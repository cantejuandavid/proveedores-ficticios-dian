#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scraper.py
==========
Mantiene actualizada la lista oficial de "Proveedores ficticios" de la DIAN.

Flujo:
  1. Abre https://www.dian.gov.co/Paginas/Inicio.aspx con Playwright (Chromium headless),
     porque la página es dinámica (JavaScript) y `requests` no ve el <li>.
     (Se bloquean imágenes/fuentes/CSS/medios para cargar solo el DOM: más rápido.)
  2. Localiza el elemento <li data-content="Proveedores ficticios"> y obtiene el href
     actual del enlace que contiene (la DIAN cambia este enlace en cada actualización).
  3. Descarga el PDF apuntado por ese enlace.
  4. Extrae la(s) tabla(s) del PDF con pdfplumber y normaliza las columnas a:
        NIT ; Razon_Social ; Resolucion ; Fecha ; Estado
  5. Si los datos cambiaron respecto al CSV actual, escribe en la raíz del repo:
        - proveedores_ficticios.csv   (UTF-8 con BOM, separador ';', con encabezados)
        - proveedores_ficticios.json  (UTF-8)
        - meta.json                   (fecha de actualización, URL del PDF, # de registros)
     Si NO cambiaron, no reescribe nada (así el commit del workflow es realmente condicional
     y `meta.json` refleja la última vez que los datos cambiaron de verdad).

Robustez:
  - La carga de la página y la descarga del PDF se reintentan ante fallos transitorios.
  - Si NO encuentra el enlace, no puede descargar, o la extracción no supera las
    validaciones mínimas, el script NO sobrescribe el CSV bueno anterior, registra el
    error y sale con código != 0.
  - Escritura atómica: primero genera archivos *.tmp, valida, y solo entonces reemplaza.

Mapeo de columnas (documentado):
  El PDF de la DIAN suele traer una tabla con columnas en este orden aproximado:
        NIT | Nombre o razón social | Resolución | Fecha | Estado/Observaciones
  El parser intenta detectar la fila de encabezado por palabras clave y mapear cada
  columna a los nombres canónicos. Si no detecta encabezado, asume el orden posicional
  anterior. Ajusta COLUMN_KEYWORDS / CANONICAL_COLUMNS si la DIAN cambia la estructura.
"""

from __future__ import annotations

import csv
import io
import json
import logging
import re
import sys
import time
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin

import requests

# ----------------------------------------------------------------------------
# Configuración
# ----------------------------------------------------------------------------

DIAN_URL = "https://www.dian.gov.co/Paginas/Inicio.aspx"
LI_SELECTOR = 'li[data-content="Proveedores ficticios"]'

# Columnas canónicas de salida (orden estable para Excel/VBA).
CANONICAL_COLUMNS = ["NIT", "Razon_Social", "Resolucion", "Fecha", "Estado"]

# Palabras clave para detectar a qué columna canónica corresponde cada encabezado del PDF.
COLUMN_KEYWORDS = {
    "NIT": ["nit", "identificacion", "documento", "cedula"],
    "Razon_Social": ["razon", "nombre", "social", "contribuyente", "proveedor"],
    "Resolucion": ["resolucion", "resolución", "acto", "numero", "número"],
    "Fecha": ["fecha", "ano", "año", "vigencia"],
    "Estado": ["estado", "observacion", "observación", "situacion", "situación"],
}

# Recursos que NO necesitamos para hallar el <li> (acelera la carga de la página).
RECURSOS_BLOQUEADOS = {"image", "font", "stylesheet", "media"}

# Validaciones mínimas para considerar la extracción exitosa.
MIN_ROWS = 5            # al menos esta cantidad de filas de datos
MIN_NIT_LIKE_RATIO = 0.5  # al menos este % de filas con un NIT plausible (dígitos)

# Reintentos ante fallos transitorios (red / render).
REINTENTOS = 3
ESPERA_REINTENTO = 4  # segundos entre intentos

# Salidas (en la raíz del repo, junto a este script).
ROOT = Path(__file__).resolve().parent
CSV_PATH = ROOT / "proveedores_ficticios.csv"
JSON_PATH = ROOT / "proveedores_ficticios.json"
META_PATH = ROOT / "meta.json"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("dian-scraper")


class ScraperError(Exception):
    """Error controlado: aborta sin sobrescribir los datos buenos."""


def _reintentar(fn, *, descripcion: str, intentos: int = REINTENTOS, espera: int = ESPERA_REINTENTO):
    """Ejecuta `fn` reintentando ante cualquier excepción transitoria."""
    ultimo_error: Exception | None = None
    for n in range(1, intentos + 1):
        try:
            return fn()
        except Exception as e:  # noqa: BLE001
            ultimo_error = e
            if n < intentos:
                log.warning("Intento %d/%d de %s falló: %s. Reintentando en %ds...",
                            n, intentos, descripcion, e, espera)
                time.sleep(espera)
            else:
                log.error("Agotados los %d intentos de %s.", intentos, descripcion)
    raise ultimo_error  # type: ignore[misc]


# ----------------------------------------------------------------------------
# Paso 1-2: obtener el enlace actual del PDF con Playwright
# ----------------------------------------------------------------------------

def _extraer_href(page) -> str:
    """Navega a la página de la DIAN y devuelve la URL absoluta del PDF."""
    page.goto(DIAN_URL, wait_until="domcontentloaded", timeout=60_000)

    # El <li> lo inyecta JS; esperamos a que aparezca (no a 'networkidle').
    page.wait_for_selector(LI_SELECTOR, timeout=30_000)

    li = page.locator(LI_SELECTOR).first
    if li.count() == 0:
        raise ScraperError(f"No se encontró el elemento {LI_SELECTOR} en la página.")

    # El enlace puede estar en un <a> hijo, en atributos data-* o embebido en el HTML.
    href = None
    anchor = li.locator("a[href]").first
    if anchor.count() > 0:
        href = anchor.get_attribute("href")

    if not href:
        for attr in ("data-href", "data-url", "data-link"):
            val = li.get_attribute(attr)
            if val:
                href = val
                break

    if not href:
        inner = li.inner_html() or ""
        m = re.search(r'href=["\']([^"\']+)["\']', inner)
        if m:
            href = m.group(1)

    if not href:
        raise ScraperError("Se encontró el <li> pero no se pudo extraer ningún enlace (href).")

    return urljoin(page.url, href.strip())


def obtener_url_pdf() -> str:
    """Renderiza la página de la DIAN (con reintentos) y devuelve la URL del PDF."""
    from playwright.sync_api import sync_playwright

    log.info("Abriendo %s con Playwright...", DIAN_URL)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            def intento() -> str:
                context = browser.new_context(user_agent=USER_AGENT, locale="es-CO")
                # Bloquear recursos pesados e innecesarios para acelerar la carga.
                context.route(
                    "**/*",
                    lambda route: (
                        route.abort()
                        if route.request.resource_type in RECURSOS_BLOQUEADOS
                        else route.continue_()
                    ),
                )
                page = context.new_page()
                try:
                    return _extraer_href(page)
                finally:
                    context.close()

            url_abs = _reintentar(intento, descripcion="cargar la página de la DIAN")
            log.info("Enlace de proveedores ficticios: %s", url_abs)
            return url_abs
        finally:
            browser.close()


# ----------------------------------------------------------------------------
# Paso 3: descargar el PDF
# ----------------------------------------------------------------------------

def descargar_pdf(url: str) -> bytes:
    log.info("Descargando PDF...")
    headers = {"User-Agent": USER_AGENT, "Accept": "application/pdf,*/*"}

    def intento() -> bytes:
        with requests.Session() as s:
            resp = s.get(url, headers=headers, timeout=120, allow_redirects=True)
        resp.raise_for_status()
        content = resp.content
        ctype = resp.headers.get("Content-Type", "")
        if not (content[:5] == b"%PDF-" or "pdf" in ctype.lower()):
            raise ScraperError(
                f"El contenido descargado no parece un PDF (Content-Type={ctype!r}, "
                f"primeros bytes={content[:8]!r})."
            )
        return content

    content = _reintentar(intento, descripcion="descargar el PDF")
    log.info("PDF descargado: %d bytes", len(content))
    return content


# ----------------------------------------------------------------------------
# Paso 4: extraer y normalizar la tabla
# ----------------------------------------------------------------------------

def _norm(s: str) -> str:
    """Minúsculas, sin acentos, sin espacios extra: para comparar encabezados."""
    s = (s or "").strip().lower()
    s = "".join(
        c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"\s+", " ", s)


def _mapear_encabezado(headers: list[str]) -> dict[int, str] | None:
    """Dado el contenido de una fila candidata a encabezado, devuelve
    {indice_columna -> nombre_canonico} si reconoce al menos NIT y Razon_Social."""
    mapeo: dict[int, str] = {}
    for idx, raw in enumerate(headers):
        h = _norm(raw)
        if not h:
            continue
        for canon, kws in COLUMN_KEYWORDS.items():
            if canon in mapeo.values():
                continue
            if any(kw in h for kw in kws):
                mapeo[idx] = canon
                break
    # Consideramos válido el encabezado si reconocemos al menos NIT y Razon_Social.
    if "NIT" in mapeo.values() and "Razon_Social" in mapeo.values():
        return mapeo
    return None


def _limpiar(celda) -> str:
    if celda is None:
        return ""
    return re.sub(r"\s+", " ", str(celda)).strip()


def _parece_nit(valor: str) -> bool:
    digitos = re.sub(r"\D", "", valor or "")
    return len(digitos) >= 5  # NIT colombiano: típicamente 8-10 dígitos


def extraer_tabla(pdf_bytes: bytes) -> list[dict]:
    """Extrae filas de todas las tablas del PDF y las normaliza a CANONICAL_COLUMNS."""
    import pdfplumber

    registros: list[dict] = []
    mapeo_global: dict[int, str] | None = None

    log.info("Extrayendo tablas con pdfplumber...")
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pnum, page in enumerate(pdf.pages, start=1):
            tablas = page.extract_tables() or []
            for tabla in tablas:
                if not tabla:
                    continue
                inicio = 0
                # Intentar detectar encabezado en la primera fila de la tabla.
                posible = _mapear_encabezado([_limpiar(c) for c in tabla[0]])
                if posible:
                    mapeo_global = posible
                    inicio = 1

                # Mapeo a usar: el detectado o, en su defecto, posicional.
                mapeo = mapeo_global
                for fila in tabla[inicio:]:
                    celdas = [_limpiar(c) for c in fila]
                    if not any(celdas):
                        continue

                    reg = {col: "" for col in CANONICAL_COLUMNS}
                    if mapeo:
                        for idx, canon in mapeo.items():
                            if idx < len(celdas):
                                reg[canon] = celdas[idx]
                    else:
                        # Fallback posicional: NIT, Razon_Social, Resolucion, Fecha, Estado
                        for i, canon in enumerate(CANONICAL_COLUMNS):
                            if i < len(celdas):
                                reg[canon] = celdas[i]

                    # Descartar filas claramente no-datos (sin NIT plausible y sin nombre).
                    if not reg["NIT"] and not reg["Razon_Social"]:
                        continue
                    registros.append(reg)

            log.info("Página %d procesada (%d registros acumulados).", pnum, len(registros))

    return registros


def validar(registros: list[dict]) -> None:
    if len(registros) < MIN_ROWS:
        raise ScraperError(
            f"Se extrajeron muy pocas filas ({len(registros)} < {MIN_ROWS}). "
            "Posible cambio de estructura del PDF; se conserva la versión anterior."
        )
    con_nit = sum(1 for r in registros if _parece_nit(r.get("NIT", "")))
    ratio = con_nit / len(registros)
    if ratio < MIN_NIT_LIKE_RATIO:
        raise ScraperError(
            f"Solo {ratio:.0%} de las filas tienen un NIT plausible "
            f"(mínimo {MIN_NIT_LIKE_RATIO:.0%}). Se conserva la versión anterior."
        )
    log.info("Validación OK: %d filas, %.0f%% con NIT plausible.", len(registros), ratio * 100)


# ----------------------------------------------------------------------------
# Paso 5: escribir salidas (atómico y solo si cambiaron los datos)
# ----------------------------------------------------------------------------

def _csv_str(registros: list[dict]) -> str:
    """Serializa los registros a CSV (separador ';', salto '\\n' determinista)."""
    buf = io.StringIO()
    writer = csv.DictWriter(
        buf, fieldnames=CANONICAL_COLUMNS, delimiter=";",
        extrasaction="ignore", lineterminator="\n",
    )
    writer.writeheader()
    writer.writerows(registros)
    return buf.getvalue()


def escribir_salidas(registros: list[dict], url_pdf: str) -> bool:
    """Escribe CSV/JSON/meta solo si el CSV cambió. Devuelve True si hubo cambios."""
    nuevo_csv = _csv_str(registros)

    # Comparar con el CSV actual (utf-8-sig descarta el BOM al leer).
    if CSV_PATH.exists():
        try:
            actual = CSV_PATH.read_text(encoding="utf-8-sig")
        except OSError:
            actual = None
        if actual == nuevo_csv:
            log.info("Datos idénticos al CSV actual (%d filas): no se reescribe nada.",
                     len(registros))
            return False

    # CSV con BOM (UTF-8) para que Excel reconozca acentos; salto '\n' (ver .gitattributes).
    tmp_csv = CSV_PATH.with_suffix(".csv.tmp")
    tmp_csv.write_text(nuevo_csv, encoding="utf-8-sig", newline="")

    tmp_json = JSON_PATH.with_suffix(".json.tmp")
    tmp_json.write_text(
        json.dumps(registros, ensure_ascii=False, indent=2), encoding="utf-8", newline="\n"
    )

    meta = {
        "fecha_actualizacion": datetime.now(timezone.utc).isoformat(),
        "url_pdf": url_pdf,
        "num_registros": len(registros),
        "columnas": CANONICAL_COLUMNS,
        "fuente": DIAN_URL,
    }
    tmp_meta = META_PATH.with_suffix(".json.tmp")
    tmp_meta.write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8", newline="\n"
    )

    # Reemplazo atómico solo cuando todo se generó bien.
    tmp_csv.replace(CSV_PATH)
    tmp_json.replace(JSON_PATH)
    tmp_meta.replace(META_PATH)
    log.info("Archivos actualizados: %s, %s, %s", CSV_PATH.name, JSON_PATH.name, META_PATH.name)
    return True


# ----------------------------------------------------------------------------
# main
# ----------------------------------------------------------------------------

def main() -> int:
    try:
        url_pdf = obtener_url_pdf()
        pdf_bytes = descargar_pdf(url_pdf)
        registros = extraer_tabla(pdf_bytes)
        validar(registros)
        cambiaron = escribir_salidas(registros, url_pdf)
        if cambiaron:
            log.info("Proceso completado: datos actualizados (%d registros).", len(registros))
        else:
            log.info("Proceso completado: sin cambios (%d registros).", len(registros))
        return 0
    except ScraperError as e:
        log.error("Fallo controlado: %s", e)
        log.error("Se conserva la última versión válida del CSV (no se sobrescribió).")
        return 2
    except Exception as e:  # noqa: BLE001
        log.exception("Fallo inesperado: %s", e)
        log.error("Se conserva la última versión válida del CSV (no se sobrescribió).")
        return 1


if __name__ == "__main__":
    sys.exit(main())
