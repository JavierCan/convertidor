# Convertidor SAT a Excel

Miniapp de Streamlit que permite convertir desde un solo archivo hasta múltiples archivos en una misma carga.

## Funciones

- Carga uno o varios archivos ZIP, TXT o CSV.
- Detecta automáticamente separadores como `~`, `|`, `;`, tabulación o coma.
- Permite indicar manualmente el separador.
- Detecta codificación UTF-8, Windows-1252 y Latin-1.
- Si un ZIP contiene varios TXT o CSV, cada uno se convierte en una hoja del mismo Excel.
- Genera un archivo XLSX independiente por cada archivo cargado.
- Permite descargar cada XLSX por separado.
- Cuando se cargan varios archivos, permite descargar todos los XLSX dentro de un solo ZIP.
- Conserva RFC, UUID, códigos y ceros a la izquierda como texto.
- Muestra una vista previa de hasta 100 registros por tabla.
- Si un archivo falla, continúa procesando los demás y muestra el error individual.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicar en Streamlit Community Cloud

1. Crea un repositorio en GitHub.
2. Sube `app.py`, `requirements.txt` y `README.md`.
3. Entra a Streamlit Community Cloud.
4. Crea una aplicación nueva y selecciona el repositorio.
5. Usa `app.py` como archivo principal.
6. Pulsa **Deploy**.

La conversión se realiza en memoria; la app no necesita guardar permanentemente los archivos cargados.
