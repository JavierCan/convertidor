# CFDI a Excel — UI para contabilidad

Proyecto listo para desplegar en Streamlit Community Cloud.

## Arquitectura

Toda la interfaz y la lógica de negocio viven en HTML, CSS y JavaScript:

```text
cfdi_contador_frontend/
├── app.py
├── requirements.txt
├── README.md
├── build.py
├── src/
│   ├── index.html
│   ├── styles.css
│   └── app.js
└── dist/
    └── app.min.html
```

`app.py` únicamente lee y carga `dist/app.min.html`.

## Experiencia diseñada para un contador

- Lenguaje contable, no técnico.
- Carga de XML y ZIP.
- Resumen de comprobantes, conceptos y observaciones.
- Elección entre una fila por concepto o por comprobante.
- Un Excel consolidado o un Excel por XML.
- Selección de campos contables principales.
- Reordenamiento de columnas con arrastrar y soltar o flechas.
- Vista previa antes de exportar.
- Conversión de importes y cantidades a valores numéricos.
- Layouts guardados en el navegador mediante `localStorage`.
- Importación y exportación de layouts JSON.
- Procesamiento local en el navegador: los XML no se envían al servidor.

## Dependencias del frontend

El HTML usa estas bibliotecas desde CDN:

- SheetJS para generar Excel.
- JSZip para leer y crear ZIP.

El navegador del usuario necesita acceso a internet para cargar esas bibliotecas.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy en Streamlit Community Cloud

1. Sube esta carpeta a GitHub.
2. Selecciona el repositorio en Streamlit Community Cloud.
3. Usa como archivo principal:
   - `app.py`, si esta carpeta es la raíz.
   - `cfdi_contador_frontend/app.py`, si está dentro de otro repositorio.
4. Pulsa **Deploy**.

## Editar el frontend

Modifica:

```text
src/index.html
src/styles.css
src/app.js
```

Después ejecuta:

```bash
python build.py
```

Esto regenera:

```text
dist/app.min.html
```
