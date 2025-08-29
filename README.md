# Redacción topográfica v3.4 — Colindancia y rumbo compacto

- Redacción de cada tramo ahora incluye: **palabras + compacto** entre paréntesis, p. ej.  
  `... con rumbo Norte veinticinco grados, treinta y cinco minutos, veinte segundos Oeste (N 25° 35´20´´O).`
- La **colindancia** se normaliza a: `Colinda con ...` (evita duplicados si ya lo escribiste).
- Exportación a **Word (.docx)** con logo + encabezado.
- Endpoint `/_version` devuelve `{ "version": "v3.4-colind-rumbo", "docx": true }` para verificar el deploy.

## Deploy rápido
Build: `pip install -r requirements.txt`  
Start: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --threads 2 --timeout 120`

## Estructura
app.py
templates/formulario.html
static/logo_hc.png
requirements.txt
runtime.txt
render.yaml
README.md
