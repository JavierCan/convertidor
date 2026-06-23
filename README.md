# CFDI a Excel — Streamlit App

Proyecto listo para desplegar en Streamlit Community Cloud.

## Funcionalidades

- Carga uno o múltiples XML.
- Carga uno o múltiples ZIP con XML.
- Detecta la estructura XML/CFDI automáticamente.
- Preview por concepto o por comprobante.
- Con un solo XML no muestra la opción de consolidar.
- Con varios XML permite consolidar o generar archivos individuales dentro de un ZIP.
- Columnas dinámicas, sin hardcodear proveedor, UUID ni campos concretos.
- Exclusión opcional de sellos y certificados largos.
- Reporte de errores sin detener el resto del proceso.
- Frontend HTML/CSS/JS integrado dentro de Streamlit.

## Estructura

```text
cfdi_streamlit_app/
├── app.py
├── requirements.txt
├── README.md
├── assets/
│   ├── index.html
│   ├── styles.css
│   └── script.js
└── examples/
    └── ejemplo.xml
```

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy

1. Crea un repositorio en GitHub.
2. Sube todo el contenido de esta carpeta.
3. En Streamlit Community Cloud selecciona el repositorio.
4. Usa `app.py` como Main file path.
5. Pulsa **Deploy**.

Para miles de CFDI, es preferible cargarlos dentro de uno o varios ZIP.
