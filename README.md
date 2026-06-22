# Convertidor SAT a Excel

Miniapp de Streamlit que permite:

- Cargar archivos ZIP, TXT o CSV.
- Leer archivos delimitados por `~` u otros separadores.
- Detectar automáticamente la codificación.
- Mostrar una vista previa.
- Convertir el contenido a Excel `.xlsx`.
- Descargar el resultado desde el navegador.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Subir a Streamlit Community Cloud

1. Crea un repositorio en GitHub.
2. Sube `app.py` y `requirements.txt`.
3. Entra a Streamlit Community Cloud.
4. Selecciona el repositorio.
5. Indica `app.py` como archivo principal.
6. Pulsa **Deploy**.

No es necesario guardar los archivos cargados en disco. La conversión se realiza en memoria.
