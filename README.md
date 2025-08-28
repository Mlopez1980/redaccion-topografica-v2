# Redacción topográfica v3.3 — Version check

- Igual que v3.2 (multitramos, auto-encadenar, DOCX con encabezado) + endpoint **/_version** para verificar el deploy.
- Título de la página muestra la versión: **v3.3-versioncheck**.

## Comprobar versión en Render
Abre `https://TU-SERVICIO.onrender.com/_version` y debe devolver:
```
{"version":"v3.3-versioncheck","docx":true}
```

## Despliegue
Build: `pip install -r requirements.txt`  
Start: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --threads 2 --timeout 120`
