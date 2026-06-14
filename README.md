# Proveedores ficticios DIAN — actualización automática

Mantiene actualizada, de forma automática, la lista oficial de **Proveedores ficticios**
de la DIAN y la publica como un CSV estable que puede leerse desde Excel/VBA por una URL fija.

- Fuente oficial: https://www.dian.gov.co/Paginas/Inicio.aspx
- El script ([`scraper.py`](scraper.py)) renderiza la página con Playwright, localiza el
  `<li data-content="Proveedores ficticios">`, resuelve el enlace **actual** del PDF
  (cambia cada vez que la DIAN actualiza), descarga el PDF y extrae la tabla con `pdfplumber`.
- Un GitHub Action programado ([`.github/workflows/actualizar.yml`](.github/workflows/actualizar.yml))
  lo ejecuta cada semana (y manualmente cuando quieras) y commitea los archivos solo si cambian.

## URL raw del CSV (la que consume Excel/VBA)

```
https://raw.githubusercontent.com/TU_USUARIO/proveedores-ficticios-dian/main/proveedores_ficticios.csv
```

> Reemplaza `TU_USUARIO` por tu usuario de GitHub. Si nombras el repo distinto, ajusta también
> el nombre del repo en la URL.

## Archivos publicados

| Archivo | Descripción |
|---|---|
| `proveedores_ficticios.csv` | UTF-8 (con BOM), separador `;`, con encabezados. **Estructura estable.** |
| `proveedores_ficticios.json` | Mismos datos en JSON. |
| `meta.json` | Fecha de actualización, URL del PDF usada y número de registros. |

### Estructura fija del CSV

```
NIT;Razon_Social;Resolucion;Fecha;Estado
```

El mapeo de columnas del PDF a estos nombres canónicos está documentado en el encabezado de
[`scraper.py`](scraper.py). El parser intenta detectar la fila de encabezado del PDF por
palabras clave; si no lo logra, asume el orden posicional anterior. Si la DIAN cambia la
estructura del PDF, ajusta `COLUMN_KEYWORDS` / `CANONICAL_COLUMNS` en el script.

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
gh repo create TU_USUARIO/proveedores-ficticios-dian --public --source=. --remote=origin --push

# Opción B) manual: crea el repo en github.com y luego
git remote add origin https://github.com/TU_USUARIO/proveedores-ficticios-dian.git
git push -u origin main
```

> **Nota sobre el `schedule`:** GitHub deshabilita los cron de Actions tras ~60 días sin
> actividad en el repo. El push semanal del propio workflow cuenta como actividad, así que
> normalmente se mantiene activo; si llegara a pausarse, ejecútalo una vez manualmente
> (pestaña **Actions → Run workflow**).

## Consumo desde Excel/VBA

El CSV mantiene encabezados fijos, separador `;` y UTF-8 para que una macro pueda hacer un
`GET` a la URL raw, partir por líneas y por `;`, y volcar a una hoja oculta. Ejemplo de
referencia (no forma parte de este repo):

```vba
Sub CargarProveedoresFicticios()
    Const URL As String = _
        "https://raw.githubusercontent.com/TU_USUARIO/proveedores-ficticios-dian/main/proveedores_ficticios.csv"
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
