# Convertidor SAT a Excel

Miniapp de Streamlit que permite:

- Cargar uno o varios archivos ZIP, TXT, CSV o XML.
- Detectar automáticamente separadores como `~`, `|`, `;`, tabulador o coma.
- Convertir XML CFDI a hojas de Excel estructuradas.
- Crear hojas para resumen, emisor, receptor, conceptos, impuestos, CFDI relacionados, timbre, complementos, pagos y detalle XML completo.
- Conservar toda la información de un XML genérico en la hoja `XML_Detalle`.
- Consolidar varios XML incluidos dentro de un ZIP en un solo Excel.
- Crear un Excel independiente por cada archivo cargado.
- Crear una hoja por cada TXT/CSV incluido dentro de un ZIP.
- Descargar cada Excel individualmente.
- Descargar todos los Excel dentro de un solo ZIP.
- Mantener UUID, RFC, códigos, números largos y ceros iniciales como texto.
- Mostrar una vista previa antes de descargar.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicar en Streamlit Community Cloud

1. Crea un repositorio en GitHub.
2. Sube `app.py`, `requirements.txt` y `README.md`.
3. Crea una aplicación en Streamlit Community Cloud.
4. Selecciona `app.py` como archivo principal.
5. Pulsa **Deploy**.
