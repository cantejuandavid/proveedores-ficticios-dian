#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scraper.py
==========
Mantiene actualizadas, desde la página de la DIAN, dos listas oficiales:
  - Proveedores ficticios
  - Contadores sancionados por la DIAN

Flujo:
  1. Abre https://www.dian.gov.co/Paginas/Inicio.aspx con Playwright (Chromium headless)
     UNA sola vez (la página es dinámica; `requests` no ve los <li>). Se bloquean
     imágenes/fuentes/CSS/medios para cargar solo el DOM (más rápido).
  2. De cada fuente localiza su <li data-content="..."> y obtiene el href ACTUAL del
     enlace que contiene (la DIAN cambia estos enlaces en cada actualización).
  3. Descarga cada PDF.
  4. Extrae la(s) tabla(s) con pdfplumber y normaliza a las columnas canónicas de la fuente.
  5. Si los datos cambiaron respecto al CSV actual, escribe los archivos de la fuente
     (csv UTF-8 BOM con ';', json y meta). Si NO cambiaron, no reescribe nada (commit
     del workflow realmente condicional; el meta refleja el último cambio real).

Robustez:
  - La carga de la página y la descarga de cada PDF se reintentan ante fallos transitorios.
  - Cada fuente se procesa de forma INDEPENDIENTE: si una falla, la otra se actualiza igual;
    la que falla conserva su última versión válida. El proceso sale con código != 0 si
    alguna fuente falló.
  - Escritura atómica: genera *.tmp, valida y solo entonces reemplaza.

Para agregar/ajustar fuentes o columnas, edita la lista FUENTES.
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
# Configuración general
# ----------------------------------------------------------------------------

DIAN_URL = "https://www.dian.gov.co/Paginas/Inicio.aspx"

# Recursos que NO necesitamos para hallar los <li> (acelera la carga de la página).
RECURSOS_BLOQUEADOS = {"image", "font", "stylesheet", "media"}

# Reintentos ante fallos transitorios (red / render).
REINTENTOS = 3
ESPERA_REINTENTO = 4  # segundos entre intentos

ROOT = Path(__file__).resolve().parent

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

# ----------------------------------------------------------------------------
# Definición de fuentes (data-driven)
# ----------------------------------------------------------------------------
#   nombre        : etiqueta legible (para logs y meta)
#   li_selector   : selector CSS del <li data-content="...">
#   csv/json/meta : nombres de archivo de salida (en la raíz del repo)
#   columnas      : columnas canónicas del CSV (orden estable para Excel/VBA)
#   keywords      : por columna canónica, palabras clave para mapear el encabezado del PDF.
#                   El match es por igualdad normalizada o por subcadena (solo si la
#                   palabra tiene >= 4 letras, para evitar falsos positivos como "no").
#   requeridas    : columnas que deben reconocerse para aceptar una fila de encabezado
#                   (y para no descartar una fila de datos).
#   col_id        : columna identificadora que debe verse "numérica" en la validación.
#   min_filas     : mínimo de filas de datos para aceptar la extracción.
#   min_id_ratio  : mínimo de filas cuyo col_id parece un número (dígitos).
FUENTES = [
    {
        "nombre": "Proveedores ficticios",
        "li_selector": 'li[data-content="Proveedores ficticios"]',
        "csv": "proveedores_ficticios.csv",
        "json": "proveedores_ficticios.json",
        "meta": "meta.json",  # se conserva el nombre histórico (compatibilidad con la macro)
        "columnas": ["NIT", "Razon_Social", "Resolucion", "Fecha", "Estado"],
        "keywords": {
            "NIT": ["nit", "identificacion", "documento", "cedula"],
            "Razon_Social": ["razon", "nombre", "social", "contribuyente", "proveedor"],
            "Resolucion": ["resolucion", "acto"],
            "Fecha": ["fecha", "ano", "vigencia"],
            "Estado": ["estado", "observacion", "situacion"],
        },
        "requeridas": ["NIT", "Razon_Social"],
        "col_id": "NIT",
        "min_filas": 5,
        "min_id_ratio": 0.5,
    },
    {
        "nombre": "Contadores sancionados por la DIAN",
        "li_selector": 'li[data-content="Contadores sancionados por la DIAN"]',
        "csv": "contadores_sancionados.csv",
        "json": "contadores_sancionados.json",
        "meta": "contadores_sancionados.meta.json",
        "columnas": [
            "No", "Nombre", "Cedula", "Inscripcion_Profesional", "Resolucion",
            "Sancion", "Fecha_Ejecutoria", "Vencimiento", "Autoridad",
        ],
        "keywords": {
            "No": ["no", "num", "numero", "item"],
            "Nombre": ["nombre"],
            "Cedula": ["cedula", "identificacion", "documento"],
            "Inscripcion_Profesional": ["inscripcion", "profesional", "tarjeta"],
            "Resolucion": ["resolucion", "acto"],
            "Sancion": ["sancion"],
            "Fecha_Ejecutoria": ["ejecutoria", "fecha"],
            "Vencimiento": ["vencimiento", "vence"],
            "Autoridad": ["autoridad", "sanciona", "entidad"],
        },
        "requeridas": ["Cedula", "Nombre"],
        "col_id": "Cedula",
        "min_filas": 3,
        "min_id_ratio": 0.5,
    },
    {
        "nombre": "Autorretenedores de renta",
        "pagina": "https://www.dian.gov.co/impuestos/Autorretenedores/Paginas/Autorretenedor-del-Impuesto-sobre-la-Renta.aspx",
        "li_selector": 'a[href*="/impuestos/Autorretenedores/Documents/"]',
        "csv": "autorretenedores_renta.csv",
        "json": "autorretenedores_renta.json",
        "meta": "autorretenedores_renta.meta.json",
        "columnas": ["NIT", "Razon_Social", "Resolucion", "Fecha"],
        "keywords": {
            "NIT": ["nit", "identificacion", "documento"],
            "Razon_Social": ["razon", "nombre", "social"],
            "Resolucion": ["resolucion", "numero", "acto"],
            "Fecha": ["fecha"],
        },
        "requeridas": ["NIT", "Razon_Social"],
        "col_id": "NIT",
        "min_filas": 100,       # son miles de autorretenedores
        "min_id_ratio": 0.9,
        # El PDF repite una fila de título y otra de encabezado en cada página:
        # quedarse solo con filas cuyo NIT sea numérico las descarta automáticamente.
        "solo_filas_con_id": True,
    },
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("dian-scraper")


class ScraperError(Exception):
    """Error controlado: aborta la fuente sin sobrescribir sus datos buenos."""


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
# Paso 1-2: obtener los enlaces actuales (una sola carga de página)
# ----------------------------------------------------------------------------

def _href_de(page, selector: str) -> str | None:
    """Devuelve la URL absoluta del enlace ubicado por `selector`, o None.
    Funciona tanto si el selector apunta al propio <a> como a un contenedor (p.ej. un <li>)."""
    el = page.locator(selector).first
    if el.count() == 0:
        return None

    # 1) el propio elemento es un <a> con href
    href = el.get_attribute("href")

    # 2) o contiene un <a href> (caso <li ...><a>...</a></li>)
    if not href:
        anchor = el.locator("a[href]").first
        if anchor.count() > 0:
            href = anchor.get_attribute("href")

    # 3) o trae el enlace en un atributo data-*
    if not href:
        for attr in ("data-href", "data-url", "data-link"):
            val = el.get_attribute(attr)
            if val:
                href = val
                break

    # 4) último recurso: buscar un href embebido en el HTML interno
    if not href:
        inner = el.inner_html() or ""
        m = re.search(r'href=["\']([^"\']+)["\']', inner)
        if m:
            href = m.group(1)

    if not href:
        return None
    return urljoin(page.url, href.strip())


def obtener_urls(fuentes: list[dict]) -> dict[str, str | None]:
    """Carga cada página necesaria (una sola vez por URL) y devuelve {nombre_fuente: url_pdf|None}.

    Las fuentes se agrupan por su página de origen: si una página falla, las fuentes de las
    otras páginas igual se resuelven."""
    from playwright.sync_api import sync_playwright

    # Agrupar fuentes por su página de origen (por defecto, la home de la DIAN).
    paginas: dict[str, list[dict]] = {}
    for f in fuentes:
        paginas.setdefault(f.get("pagina", DIAN_URL), []).append(f)

    resultado: dict[str, str | None] = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            for pagina, fs in paginas.items():
                log.info("Abriendo %s con Playwright...", pagina)

                def intento(pagina=pagina, fs=fs) -> dict[str, str | None]:
                    context = browser.new_context(user_agent=USER_AGENT, locale="es-CO")
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
                        page.goto(pagina, wait_until="domcontentloaded", timeout=60_000)
                        urls: dict[str, str | None] = {}
                        for f in fs:
                            sel = f["li_selector"]
                            try:
                                page.wait_for_selector(sel, timeout=30_000)
                            except Exception:
                                log.warning("No apareció el selector de '%s' (%s).", f["nombre"], sel)
                            urls[f["nombre"]] = _href_de(page, sel)
                        return urls
                    finally:
                        context.close()

                try:
                    urls = _reintentar(intento, descripcion=f"cargar {pagina}")
                except Exception as e:  # noqa: BLE001
                    log.error("No se pudo cargar %s: %s", pagina, e)
                    urls = {f["nombre"]: None for f in fs}
                resultado.update(urls)

            for nombre, url in resultado.items():
                log.info("Enlace [%s]: %s", nombre, url or "NO ENCONTRADO")
            return resultado
        finally:
            browser.close()


# ----------------------------------------------------------------------------
# Paso 3: descargar un PDF
# ----------------------------------------------------------------------------

def descargar_pdf(url: str) -> bytes:
    log.info("Descargando PDF: %s", url)
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
    """Minúsculas, sin acentos ni puntuación, sin espacios extra (para comparar encabezados)."""
    s = (s or "").strip().lower()
    s = "".join(
        c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn"
    )
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _coincide(h: str, kw: str) -> bool:
    """True si el encabezado `h` corresponde a la palabra clave `kw`.
    Igualdad normalizada, o subcadena solo si kw tiene >= 4 letras (evita 'no' en 'nombre')."""
    return h == kw or (len(kw) >= 4 and kw in h)


def _mapear_encabezado(headers: list[str], keywords: dict[str, list[str]],
                       requeridas: list[str]) -> dict[int, str] | None:
    """Devuelve {indice_columna -> nombre_canonico} si reconoce todas las `requeridas`."""
    mapeo: dict[int, str] = {}
    for idx, raw in enumerate(headers):
        h = _norm(raw)
        if not h:
            continue
        for canon, kws in keywords.items():
            if canon in mapeo.values():
                continue
            if any(_coincide(h, kw) for kw in kws):
                mapeo[idx] = canon
                break
    if all(col in mapeo.values() for col in requeridas):
        return mapeo
    return None


def _limpiar(celda) -> str:
    if celda is None:
        return ""
    return re.sub(r"\s+", " ", str(celda)).strip()


def _parece_id(valor: str) -> bool:
    return len(re.sub(r"\D", "", valor or "")) >= 5


def extraer_tabla(pdf_bytes: bytes, fuente: dict) -> list[dict]:
    """Extrae filas de todas las tablas del PDF y las normaliza a las columnas de la fuente."""
    import pdfplumber

    columnas = fuente["columnas"]
    keywords = fuente["keywords"]
    requeridas = fuente["requeridas"]
    col_id = fuente["col_id"]
    solo_con_id = fuente.get("solo_filas_con_id", False)

    registros: list[dict] = []
    mapeo_global: dict[int, str] | None = None

    log.info("[%s] Extrayendo tablas con pdfplumber...", fuente["nombre"])
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pnum, page in enumerate(pdf.pages, start=1):
            tablas = page.extract_tables() or []
            for tabla in tablas:
                if not tabla:
                    continue
                # Buscar la fila de encabezado en las primeras filas (hay PDF con una fila
                # de título antes del encabezado, y el encabezado se repite por página).
                inicio = 0
                for hi in range(min(5, len(tabla))):
                    posible = _mapear_encabezado(
                        [_limpiar(c) for c in tabla[hi]], keywords, requeridas
                    )
                    if posible:
                        mapeo_global = posible
                        inicio = hi + 1
                        break

                mapeo = mapeo_global
                for fila in tabla[inicio:]:
                    celdas = [_limpiar(c) for c in fila]
                    if not any(celdas):
                        continue

                    reg = {col: "" for col in columnas}
                    if mapeo:
                        for idx, canon in mapeo.items():
                            if idx < len(celdas):
                                reg[canon] = celdas[idx]
                    else:
                        # Fallback posicional: usa el orden de `columnas`.
                        for i, canon in enumerate(columnas):
                            if i < len(celdas):
                                reg[canon] = celdas[i]

                    if solo_con_id:
                        # Mantener solo filas cuyo identificador parezca un número
                        # (descarta filas de título/encabezado repetidas por página).
                        if not _parece_id(reg.get(col_id, "")):
                            continue
                    else:
                        # Descartar filas no-datos: todas las columnas requeridas vacías.
                        if all(not reg.get(c) for c in requeridas):
                            continue
                    registros.append(reg)

            log.info("[%s] Página %d procesada (%d registros).",
                     fuente["nombre"], pnum, len(registros))

    return registros


def validar(registros: list[dict], fuente: dict) -> None:
    nombre = fuente["nombre"]
    min_filas = fuente["min_filas"]
    col_id = fuente["col_id"]
    min_ratio = fuente["min_id_ratio"]

    if len(registros) < min_filas:
        raise ScraperError(
            f"Se extrajeron muy pocas filas ({len(registros)} < {min_filas}). "
            "Posible cambio de estructura del PDF; se conserva la versión anterior."
        )
    con_id = sum(1 for r in registros if _parece_id(r.get(col_id, "")))
    ratio = con_id / len(registros)
    if ratio < min_ratio:
        raise ScraperError(
            f"Solo {ratio:.0%} de las filas tienen '{col_id}' plausible "
            f"(mínimo {min_ratio:.0%}). Se conserva la versión anterior."
        )
    log.info("[%s] Validación OK: %d filas, %.0f%% con '%s' plausible.",
             nombre, len(registros), ratio * 100, col_id)


# ----------------------------------------------------------------------------
# Paso 5: escribir salidas (atómico y solo si cambiaron los datos)
# ----------------------------------------------------------------------------

def _csv_str(registros: list[dict], columnas: list[str]) -> str:
    """Serializa los registros a CSV (separador ';', salto '\\n' determinista)."""
    buf = io.StringIO()
    writer = csv.DictWriter(
        buf, fieldnames=columnas, delimiter=";",
        extrasaction="ignore", lineterminator="\n",
    )
    writer.writeheader()
    writer.writerows(registros)
    return buf.getvalue()


def escribir_salidas(registros: list[dict], url_pdf: str, fuente: dict) -> bool:
    """Escribe CSV/JSON/meta de la fuente solo si el CSV cambió. Devuelve True si hubo cambios."""
    columnas = fuente["columnas"]
    csv_path = ROOT / fuente["csv"]
    json_path = ROOT / fuente["json"]
    meta_path = ROOT / fuente["meta"]

    nuevo_csv = _csv_str(registros, columnas)

    if csv_path.exists():
        try:
            actual = csv_path.read_text(encoding="utf-8-sig")
        except OSError:
            actual = None
        if actual == nuevo_csv:
            log.info("[%s] Datos idénticos al CSV actual (%d filas): no se reescribe nada.",
                     fuente["nombre"], len(registros))
            return False

    # CSV con BOM (UTF-8) para Excel; salto '\n' (ver .gitattributes).
    tmp_csv = csv_path.with_suffix(csv_path.suffix + ".tmp")
    tmp_csv.write_text(nuevo_csv, encoding="utf-8-sig", newline="")

    tmp_json = json_path.with_suffix(json_path.suffix + ".tmp")
    tmp_json.write_text(
        json.dumps(registros, ensure_ascii=False, indent=2), encoding="utf-8", newline="\n"
    )

    meta = {
        "fuente": fuente["nombre"],
        "fecha_actualizacion": datetime.now(timezone.utc).isoformat(),
        "url_pdf": url_pdf,
        "num_registros": len(registros),
        "columnas": columnas,
        "pagina": DIAN_URL,
    }
    tmp_meta = meta_path.with_suffix(meta_path.suffix + ".tmp")
    tmp_meta.write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8", newline="\n"
    )

    tmp_csv.replace(csv_path)
    tmp_json.replace(json_path)
    tmp_meta.replace(meta_path)
    log.info("[%s] Archivos actualizados: %s, %s, %s",
             fuente["nombre"], csv_path.name, json_path.name, meta_path.name)
    return True


# ----------------------------------------------------------------------------
# main
# ----------------------------------------------------------------------------

def procesar_fuente(fuente: dict, url: str | None) -> bool:
    """Procesa una fuente. Devuelve True si todo fue bien; False si falló (controladamente)."""
    nombre = fuente["nombre"]
    try:
        if not url:
            raise ScraperError(f"No se encontró el enlace de '{nombre}' en la página.")
        pdf = descargar_pdf(url)
        registros = extraer_tabla(pdf, fuente)
        validar(registros, fuente)
        cambio = escribir_salidas(registros, url, fuente)
        log.info("[%s] %s (%d registros).",
                 nombre, "ACTUALIZADO" if cambio else "sin cambios", len(registros))
        return True
    except ScraperError as e:
        log.error("[%s] Fallo controlado: %s", nombre, e)
    except Exception as e:  # noqa: BLE001
        log.exception("[%s] Fallo inesperado: %s", nombre, e)
    log.error("[%s] Se conserva la última versión válida (no se sobrescribió).", nombre)
    return False


def main() -> int:
    try:
        urls = obtener_urls(FUENTES)
    except Exception as e:  # noqa: BLE001
        log.exception("No se pudo cargar la página de la DIAN: %s", e)
        log.error("Se conservan todas las versiones válidas anteriores.")
        return 1

    fallidas = [f["nombre"] for f in FUENTES if not procesar_fuente(f, urls.get(f["nombre"]))]

    if fallidas:
        log.error("Fuentes con error: %s.", ", ".join(fallidas))
        return 2
    log.info("Proceso completado: todas las fuentes OK.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
