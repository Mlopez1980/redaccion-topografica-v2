# Redacción topográfica v2

- Colindancia por tramo.
- Ingreso de rumbo como texto (`N, 25, 35, 20, O`) o por campos.
- Salida en texto y compacta. Exportación a Word con encabezado de Honduras Constructores.

## Encabezado
El `.docx` incluye encabezado con el logo (`static/logo_hc.png`) y el texto:
"Este programa fue creado por Honduras Constructores S de R L".

## Estructura
```
redaccion-topografica-v2/
├─ app.py
├─ requirements.txt
├─ runtime.txt
├─ render.yaml
├─ static/
│  └─ logo_hc.png
└─ templates/
   └─ formulario.html
```

## Local
```
pip install -r requirements.txt
python app.py
```

## Render
Build: `pip install -r requirements.txt`  
Start: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --threads 2 --timeout 120`
