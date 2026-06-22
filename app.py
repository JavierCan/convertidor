import io
import os
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIGURACIÓN DE LA APP
# ============================================================

st.set_page_config(
    page_title="Convertidor SAT a Excel",
    page_icon="📄",
    layout="wide",
)

st.title("📄 Convertidor de archivos SAT a Excel")
st.write(
    "Carga un archivo **ZIP, TXT o CSV** y conviértelo a un archivo "
    "**Excel (.xlsx)** listo para descargar."
)


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def limpiar_nombre_hoja(nombre: str, usados: set[str]) -> str:
    """
    Genera un nombre de hoja válido para Excel.
    """
    nombre = Path(nombre).stem

    caracteres_invalidos = ['\\', '/', '*', '?', ':', '[', ']']
    for caracter in caracteres_invalidos:
        nombre = nombre.replace(caracter, "_")

    nombre = nombre.strip() or "Hoja"
    nombre = nombre[:31]

    nombre_base = nombre
    contador = 1

    while nombre in usados:
        sufijo = f"_{contador}"
        nombre = f"{nombre_base[:31 - len(sufijo)]}{sufijo}"
        contador += 1

    usados.add(nombre)
    return nombre


def detectar_separador(contenido: bytes, codificacion: str) -> str:
    """
    Detecta el separador más probable.
    Se prioriza '~', porque es el formato habitual del archivo indicado.
    """
    texto = contenido.decode(codificacion, errors="ignore")
    primera_linea = texto.splitlines()[0] if texto.splitlines() else ""

    candidatos = ["~", "|", ";", "\t", ","]

    conteos = {
        separador: primera_linea.count(separador)
        for separador in candidatos
    }

    mejor_separador = max(conteos, key=conteos.get)

    if conteos[mejor_separador] == 0:
        return "~"

    return mejor_separador


def leer_tabla(
    contenido: bytes,
    nombre_archivo: str,
    separador_manual: str | None = None
) -> tuple[pd.DataFrame, str, str]:
    """
    Lee un TXT o CSV probando varias codificaciones.
    Mantiene los valores como texto para proteger RFC, UUID y claves.
    """

    codificaciones = [
        "utf-8-sig",
        "utf-8",
        "cp1252",
        "latin-1",
    ]

    errores = []

    for codificacion in codificaciones:
        try:
            separador = (
                separador_manual
                if separador_manual
                else detectar_separador(contenido, codificacion)
            )

            dataframe = pd.read_csv(
                io.BytesIO(contenido),
                sep=separador,
                encoding=codificacion,
                dtype=str,
                keep_default_na=False,
                engine="python",
                on_bad_lines="warn",
            )

            dataframe.columns = [
                str(columna).replace("\ufeff", "").strip()
                for columna in dataframe.columns
            ]

            # Eliminar columnas cuyo nombre y contenido estén completamente vacíos.
            columnas_validas = []
            for columna in dataframe.columns:
                nombre_no_vacio = str(columna).strip() != ""
                contenido_no_vacio = dataframe[columna].astype(str).str.strip().ne("").any()

                if nombre_no_vacio or contenido_no_vacio:
                    columnas_validas.append(columna)

            dataframe = dataframe[columnas_validas]

            if dataframe.empty and len(dataframe.columns) <= 1:
                raise ValueError(
                    f"No se detectaron datos tabulares en {nombre_archivo}."
                )

            return dataframe, codificacion, separador

        except Exception as error:
            errores.append(f"{codificacion}: {error}")

    detalle = "\n".join(errores)
    raise ValueError(
        f"No fue posible leer el archivo {nombre_archivo}.\n\n{detalle}"
    )


def extraer_tablas(
    archivo_subido,
    separador_manual: str | None = None
) -> list[dict]:
    """
    Procesa un ZIP, TXT o CSV y devuelve una lista de tablas.
    """
    nombre_archivo = archivo_subido.name
    contenido = archivo_subido.getvalue()
    extension = Path(nombre_archivo).suffix.lower()

    tablas = []

    if extension == ".zip":
        try:
            with zipfile.ZipFile(io.BytesIO(contenido), "r") as archivo_zip:
                archivos_internos = [
                    nombre
                    for nombre in archivo_zip.namelist()
                    if not nombre.endswith("/")
                    and Path(nombre).suffix.lower() in {".txt", ".csv"}
                ]

                if not archivos_internos:
                    raise ValueError(
                        "El ZIP no contiene archivos TXT o CSV compatibles."
                    )

                for archivo_interno in archivos_internos:
                    contenido_interno = archivo_zip.read(archivo_interno)

                    dataframe, codificacion, separador = leer_tabla(
                        contenido_interno,
                        archivo_interno,
                        separador_manual,
                    )

                    tablas.append(
                        {
                            "nombre": archivo_interno,
                            "dataframe": dataframe,
                            "codificacion": codificacion,
                            "separador": separador,
                        }
                    )

        except zipfile.BadZipFile as error:
            raise ValueError(
                "El archivo ZIP está dañado o no es un ZIP válido."
            ) from error

    elif extension in {".txt", ".csv"}:
        dataframe, codificacion, separador = leer_tabla(
            contenido,
            nombre_archivo,
            separador_manual,
        )

        tablas.append(
            {
                "nombre": nombre_archivo,
                "dataframe": dataframe,
                "codificacion": codificacion,
                "separador": separador,
            }
        )

    else:
        raise ValueError(
            "Formato no compatible. Selecciona un archivo ZIP, TXT o CSV."
        )

    return tablas


def ajustar_hoja_excel(worksheet, dataframe: pd.DataFrame) -> None:
    """
    Aplica formato básico al archivo Excel.
    """
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    relleno_encabezado = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )

    fuente_encabezado = Font(
        color="FFFFFF",
        bold=True,
    )

    for celda in worksheet[1]:
        celda.fill = relleno_encabezado
        celda.font = fuente_encabezado
        celda.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    for indice_columna, columna in enumerate(dataframe.columns, start=1):
        valores = dataframe[columna].astype(str)

        largo_maximo = max(
            len(str(columna)),
            valores.map(len).max() if not valores.empty else 0,
        )

        ancho = min(max(largo_maximo + 2, 12), 50)
        letra_columna = get_column_letter(indice_columna)
        worksheet.column_dimensions[letra_columna].width = ancho

        # Todos los valores se escriben como texto para evitar cambios en RFC,
        # UUID, códigos con ceros a la izquierda o números demasiado largos.
        for fila in range(2, len(dataframe) + 2):
            worksheet.cell(
                row=fila,
                column=indice_columna,
            ).number_format = "@"


def crear_excel(tablas: list[dict]) -> bytes:
    """
    Crea el archivo XLSX en memoria y devuelve sus bytes.
    """
    buffer_salida = io.BytesIO()
    nombres_usados = set()

    with pd.ExcelWriter(
        buffer_salida,
        engine="openpyxl",
    ) as writer:
        for tabla in tablas:
            dataframe = tabla["dataframe"]

            nombre_hoja = limpiar_nombre_hoja(
                tabla["nombre"],
                nombres_usados,
            )

            dataframe.to_excel(
                writer,
                sheet_name=nombre_hoja,
                index=False,
            )

            worksheet = writer.sheets[nombre_hoja]
            ajustar_hoja_excel(worksheet, dataframe)

    buffer_salida.seek(0)
    return buffer_salida.getvalue()


def nombre_salida(nombre_entrada: str) -> str:
    """
    Genera el nombre del archivo de salida.
    """
    base = Path(nombre_entrada).stem
    return f"{base}_convertido.xlsx"


# ============================================================
# BARRA LATERAL
# ============================================================

with st.sidebar:
    st.header("⚙️ Configuración")

    modo_separador = st.radio(
        "Separador de columnas",
        options=[
            "Detectar automáticamente",
            "Usar separador personalizado",
        ],
        index=0,
    )

    separador_manual = None

    if modo_separador == "Usar separador personalizado":
        separador_manual = st.text_input(
            "Separador",
            value="~",
            max_chars=3,
            help="Para los archivos del SAT normalmente se usa ~.",
        )

        if separador_manual == r"\t":
            separador_manual = "\t"

    st.markdown("---")
    st.caption(
        "La aplicación mantiene las columnas como texto para evitar "
        "que Excel modifique UUID, RFC, códigos o ceros a la izquierda."
    )


# ============================================================
# CARGA DEL ARCHIVO
# ============================================================

archivo_subido = st.file_uploader(
    "Selecciona el archivo",
    type=["zip", "txt", "csv"],
    help="Puedes cargar el ZIP original o el TXT/CSV extraído.",
)

if archivo_subido is None:
    st.info("Carga un archivo para iniciar la conversión.")
    st.stop()


# ============================================================
# PROCESAMIENTO
# ============================================================

try:
    with st.spinner("Leyendo y validando el archivo..."):
        tablas = extraer_tablas(
            archivo_subido,
            separador_manual=separador_manual,
        )

    total_registros = sum(
        len(tabla["dataframe"])
        for tabla in tablas
    )

    total_columnas = sum(
        len(tabla["dataframe"].columns)
        for tabla in tablas
    )

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Archivos procesados",
        len(tablas),
    )

    col2.metric(
        "Registros totales",
        f"{total_registros:,}",
    )

    col3.metric(
        "Columnas detectadas",
        total_columnas,
    )

    st.success("El archivo se leyó correctamente.")

    # --------------------------------------------------------
    # RESUMEN Y VISTA PREVIA
    # --------------------------------------------------------

    st.subheader("🔎 Vista previa")

    nombres_pestanas = [
        Path(tabla["nombre"]).name
        for tabla in tablas
    ]

    pestanas = st.tabs(nombres_pestanas)

    for pestana, tabla in zip(pestanas, tablas):
        with pestana:
            dataframe = tabla["dataframe"]

            st.write(
                f"**Registros:** {len(dataframe):,}  |  "
                f"**Columnas:** {len(dataframe.columns)}  |  "
                f"**Codificación:** `{tabla['codificacion']}`  |  "
                f"**Separador:** `{repr(tabla['separador'])[1:-1]}`"
            )

            st.dataframe(
                dataframe.head(100),
                use_container_width=True,
                hide_index=True,
            )

            if len(dataframe) > 100:
                st.caption(
                    "La vista previa muestra las primeras 100 filas. "
                    "El Excel incluirá todos los registros."
                )

    # --------------------------------------------------------
    # CREAR Y DESCARGAR EXCEL
    # --------------------------------------------------------

    with st.spinner("Generando el archivo Excel..."):
        excel_bytes = crear_excel(tablas)

    st.subheader("⬇️ Descargar resultado")

    st.download_button(
        label="Descargar archivo XLSX",
        data=excel_bytes,
        file_name=nombre_salida(archivo_subido.name),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True,
    )

except Exception as error:
    st.error("No se pudo convertir el archivo.")
    st.exception(error)
