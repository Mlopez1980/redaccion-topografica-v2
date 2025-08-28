# Redacción topográfica v3 — Multitramos

- Múltiples tramos con **estación inicio/fin**, **rumbo**, **distancia** y **colindancia**.
- Exportación a **Word (.docx)** con encabezado (logo + texto).
- Rumbo aceptado como texto: `N, 25, 35, 20, O` o `N 25°35'20'' O`.

## Local
```
pip install -r requirements.txt
python app.py
```
Abrir http://127.0.0.1:5000

## Render
Build: `pip install -r requirements.txt`  
Start: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --threads 2 --timeout 120`
