# Listas DIAN — actualización automática

Mantiene actualizadas, de forma automática, tres listas oficiales de la DIAN y las publica
como CSV estables que pueden leerse desde Excel/VBA por una URL fija:

1. **Proveedores ficticios**
2. **Contadores sancionados por la DIAN**
3. **Autorretenedores del impuesto sobre la renta**

- Fuente oficial: https://www.dian.gov.co/Paginas/Inicio.aspx
- El script ([`scraper.py`](scraper.py)) renderiza con Playwright la página de cada lista,
  localiza su enlace (un `<li data-content="...">` o un `<a href>`), resuelve la URL **actual**
  de su PDF (cambia cada vez que la DIAN actualiza), lo descarga y extrae la tabla con `pdfplumber`.
- Cada lista se procesa de forma **independiente**: si una falla, las demás se actualizan igual.
- Un GitHub Action programado ([`.github/workflows/actualizar.yml`](.github/workflows/actualizar.yml))
  lo ejecuta dos veces por semana (lunes y jueves) y manualmente cuando quieras, y commitea
  los archivos solo si cambian.

## Propósito y aviso

Este proyecto tiene **fines exclusivamente académicos y de apoyo a la comunidad contable y
tributaria**. Su único objetivo es facilitar el acceso a información **pública** que la DIAN ya
publica en su página oficial —las listas de **Proveedores ficticios**, **Contadores
sancionados por la DIAN** y **Autorretenedores del impuesto sobre la renta**—, de modo que
sirva como **ayuda para la toma de decisiones**.

- No persigue ningún fin comercial, ni distinto al de consultar y consolidar esa información pública.
- No modifica, interpreta ni certifica los datos: solo reproduce lo que la DIAN publica en sus PDF.
- La **fuente oficial y única válida** sigue siendo la DIAN
  (https://www.dian.gov.co/Paginas/Inicio.aspx). Ante cualquier diferencia, prevalece la
  publicación oficial.
- Los datos pueden contener errores de extracción o estar desactualizados respecto a la fuente;
  verifica siempre contra el PDF oficial antes de tomar decisiones con efectos legales.

## URLs raw de los CSV (las que consume Excel/VBA)

```
https://raw.githubusercontent.com/cantejuandavid/proveedores-ficticios-dian/main/proveedores_ficticios.csv
https://raw.githubusercontent.com/cantejuandavid/proveedores-ficticios-dian/main/contadores_sancionados.csv
https://raw.githubusercontent.com/cantejuandavid/proveedores-ficticios-dian/main/autorretenedores_renta.csv
```

> Son las URLs fijas que consumen las macros de Excel/VBA.

## Archivos publicados

| Archivo | Descripción |
|---|---|
| `proveedores_ficticios.csv` | UTF-8 (con BOM), separador `;`, con encabezados. **Estructura estable.** |
| `proveedores_ficticios.json` | Mismos datos en JSON. |
| `meta.json` | Meta de proveedores: fecha de actualización, URL del PDF y número de registros. |
| `contadores_sancionados.csv` | UTF-8 (con BOM), separador `;`, con encabezados. **Estructura estable.** |
| `contadores_sancionados.json` | Mismos datos en JSON. |
| `contadores_sancionados.meta.json` | Meta de contadores: fecha de actualización, URL del PDF y número de registros. |
| `autorretenedores_renta.csv` | UTF-8 (con BOM), separador `;`, con encabezados. **Estructura estable.** |
| `autorretenedores_renta.json` | Mismos datos en JSON. |
| `autorretenedores_renta.meta.json` | Meta de autorretenedores: fecha de actualización, URL del PDF y número de registros. |

### Estructura fija de los CSV

```
# proveedores_ficticios.csv
NIT;Razon_Social;Resolucion;Fecha;Estado

# contadores_sancionados.csv
No;Nombre;Cedula;Inscripcion_Profesional;Resolucion;Sancion;Fecha_Ejecutoria;Vencimiento;Autoridad

# autorretenedores_renta.csv
NIT;Razon_Social;Resolucion;Fecha
```

El mapeo de columnas del PDF a estos nombres canónicos está definido por fuente en la lista
`FUENTES` de [`scraper.py`](scraper.py). El parser detecta la fila de encabezado del PDF por
palabras clave; si no lo logra, asume el orden posicional. Si la DIAN cambia la estructura de
algún PDF, ajusta esa fuente en `FUENTES` (columnas / keywords).

## Robustez

- Si **no encuentra el enlace**, no puede descargar, o la tabla no supera las validaciones
  mínimas (filas/columnas/NIT plausible), el script **no sobrescribe** el CSV bueno anterior,
  registra el error y sale con código distinto de 0 (el Action falla y conserva la última
  versión válida).
- Escritura **atómica**: genera `*.tmp`, valida y solo entonces reemplaza los archivos finales.

## Uso local

```bash
python -m venv .venv
# Windows PowerShell:  .venv\Scripts\Activate.ps1
# Linux/Mac:           source .venv/bin/activate
pip install -r requirements.txt
python -m playwright install chromium
python scraper.py
```

Códigos de salida: `0` éxito · `2` fallo controlado (estructura/validación) · `1` fallo inesperado.

## Publicar en GitHub

```bash
git init
git add .
git commit -m "Proyecto inicial: scraper DIAN + GitHub Action"
git branch -M main

# Opción A) con GitHub CLI (requiere: gh auth login)
gh repo create cantejuandavid/proveedores-ficticios-dian --public --source=. --remote=origin --push

# Opción B) manual: crea el repo en github.com y luego
git remote add origin https://github.com/cantejuandavid/proveedores-ficticios-dian.git
git push -u origin main
```

> **Nota sobre el `schedule`:** GitHub deshabilita los cron de Actions tras ~60 días sin
> actividad en el repo. El push semanal del propio workflow cuenta como actividad, así que
> normalmente se mantiene activo; si llegara a pausarse, ejecútalo una vez manualmente
> (pestaña **Actions → Run workflow**).

## Consumo desde Excel/VBA

El CSV mantiene encabezados fijos, separador `;` y UTF-8 para que una macro pueda hacer un
`GET` a la URL raw, partir por líneas y por `;`, y volcar a una hoja oculta.

> 📎 **Versión avanzada con fechas:** el módulo [`ejemplo_macro.bas`](ejemplo_macro.bas)
> carga el CSV y además muestra tres fechas en la hoja oculta (`H1:I4`):
> *Lista DIAN actualizada al* (de `meta.json`), *Última verificación del robot*
> (consultando la API de GitHub Actions, **sin generar commits**) y *Consultado por mí el*
> (`Now()` local). Requiere que el workflow se haya ejecutado al menos una vez en Actions
> para que exista un run del cual leer la fecha de verificación.

Ejemplo mínimo de referencia (no forma parte del flujo automático):

```vba
Sub CargarProveedoresFicticios()
    Const URL As String = _
        "https://raw.githubusercontent.com/cantejuandavid/proveedores-ficticios-dian/main/proveedores_ficticios.csv"
    Dim http As Object, texto As String, lineas() As String, campos() As String
    Dim i As Long, ws As Worksheet

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", URL & "?t=" & Format(Now, "yyyymmddhhnnss"), False ' evita caché
    http.Send
    If http.Status <> 200 Then
        MsgBox "Error al descargar: " & http.Status
        Exit Sub
    End If

    texto = http.responseText
    ' Quitar BOM UTF-8 si viniera
    If Len(texto) > 0 Then If AscW(Left(texto, 1)) = 65279 Then texto = Mid(texto, 2)
    texto = Replace(texto, vbCrLf, vbLf)
    texto = Replace(texto, vbCr, vbLf)
    lineas = Split(texto, vbLf)

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ProveedoresFicticios")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "ProveedoresFicticios"
        ws.Visible = xlSheetVeryHidden
    End If
    ws.Cells.Clear

    For i = LBound(lineas) To UBound(lineas)
        If Len(Trim(lineas(i))) > 0 Then
            campos = Split(lineas(i), ";")
            Dim j As Long
            For j = LBound(campos) To UBound(campos)
                ws.Cells(i + 1, j + 1).Value = campos(j)
            Next j
        End If
    Next i
End Sub
```
