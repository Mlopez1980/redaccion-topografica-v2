# Redacción topográfica v3.2 — Multitramos con auto-encadenar

- **Auto-encadenar**: si dejas vacío el inicio de un tramo, toma el fin del tramo anterior; al agregar fila nueva, copia automáticamente ese valor en la UI.
- Múltiples tramos: Est. inicio/fin, Rumbo, Distancia, Colindancia.
- Exportación a **Word (.docx)** con encabezado (logo + texto institucional).
- Parser flexible para rumbos: `N, 25, 35, 20, O`, `S 10°0'30'' E`, `N 10 5 0 O`, y palabras (Norte/Sur/Este/Oeste).

## Local
```
pip install -r requirements.txt
python app.py
```
Abrir http://127.0.0.1:5000

## Render
Build: `pip install -r requirements.txt`  
Start: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --threads 2 --timeout 120`

## Estructura
```
app.py
templates/formulario.html
static/logo_hc.png
requirements.txt
runtime.txt
render.yaml
README.md
```
