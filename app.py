import io
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
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
    "Carga **uno o varios archivos ZIP, TXT o CSV**. La aplicación generará "
    "un archivo Excel por cada archivo cargado."
)


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================


def limpiar_nombre_hoja(nombre: str, usados: set[str]) -> str:
    """Genera un nombre de hoja válido y único para Excel."""
    nombre = Path(nombre).stem

    for caracter in ["\\", "/", "*", "?", ":", "[", "]"]:
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


def limpiar_nombre_archivo(nombre: str) -> str:
    """Limpia un nombre para usarlo dentro del ZIP de resultados."""
    nombre = Path(nombre).stem.strip() or "archivo"

    for caracter in ['<', '>', ':', '"', '/', '\\', '|', '?', '*']:
        nombre = nombre.replace(caracter, "_")

    return nombre


def nombre_unico(nombre_deseado: str, usados: set[str]) -> str:
    """Evita sobrescribir resultados cuando existen nombres repetidos."""
    ruta = Path(nombre_deseado)
    nombre_base = ruta.stem
    extension = ruta.suffix
    resultado = nombre_deseado
    contador = 1

    while resultado.lower() in usados:
        resultado = f"{nombre_base}_{contador}{extension}"
        contador += 1

    usados.add(resultado.lower())
    return resultado


def detectar_separador(contenido: bytes, codificacion: str) -> str:
    """Detecta el separador más probable en la primera línea."""
    texto = contenido.decode(codificacion, errors="ignore")
    lineas = texto.splitlines()
    primera_linea = lineas[0] if lineas else ""

    candidatos = ["~", "|", ";", "\t", ","]
    conteos = {
        separador: primera_linea.count(separador)
        for separador in candidatos
    }

    mejor_separador = max(conteos, key=conteos.get)
    return mejor_separador if conteos[mejor_separador] > 0 else "~"


def leer_tabla(
    contenido: bytes,
    nombre_archivo: str,
    separador_manual: str | None = None,
) -> tuple[pd.DataFrame, str, str]:
    """
    Lee un TXT o CSV probando varias codificaciones.
    Mantiene todos los valores como texto para proteger RFC, UUID y claves.
    """
    codificaciones = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]
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

            columnas_validas = []
            for columna in dataframe.columns:
                nombre_no_vacio = str(columna).strip() != ""
                contenido_no_vacio = (
                    dataframe[columna]
                    .astype(str)
                    .str.strip()
                    .ne("")
                    .any()
                )

                if nombre_no_vacio or contenido_no_vacio:
                    columnas_validas.append(columna)

            dataframe = dataframe[columnas_validas]

            if len(dataframe.columns) <= 1:
                raise ValueError(
                    "Solo se detectó una columna. Revisa el separador seleccionado."
                )

            return dataframe, codificacion, separador

        except Exception as error:
            errores.append(f"{codificacion}: {error}")

    detalle = "\n".join(errores)
    raise ValueError(
        f"No fue posible leer {nombre_archivo}.\n\n{detalle}"
    )


def extraer_tablas(
    archivo_subido,
    separador_manual: str | None = None,
) -> list[dict]:
    """Procesa un ZIP, TXT o CSV y devuelve todas las tablas encontradas."""
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
    """Aplica formato legible a una hoja de Excel."""
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    relleno_encabezado = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )
    fuente_encabezado = Font(color="FFFFFF", bold=True)

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

        for fila in range(2, len(dataframe) + 2):
            worksheet.cell(
                row=fila,
                column=indice_columna,
            ).number_format = "@"


def crear_excel(tablas: list[dict]) -> bytes:
    """Crea un XLSX en memoria con una hoja por tabla encontrada."""
    buffer_salida = io.BytesIO()
    nombres_usados = set()

    with pd.ExcelWriter(buffer_salida, engine="openpyxl") as writer:
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


def crear_zip_resultados(resultados: list[dict]) -> bytes:
    """Agrupa todos los XLSX generados dentro de un ZIP descargable."""
    buffer_zip = io.BytesIO()

    with zipfile.ZipFile(
        buffer_zip,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
    ) as archivo_zip:
        for resultado in resultados:
            archivo_zip.writestr(
                resultado["nombre_salida"],
                resultado["excel_bytes"],
            )

    buffer_zip.seek(0)
    return buffer_zip.getvalue()


def mostrar_separador(separador: str) -> str:
    """Muestra tabulación de forma entendible en la interfaz."""
    return "TAB" if separador == "\t" else separador


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

        if separador_manual == "":
            separador_manual = None

    st.markdown("---")
    st.caption(
        "Cada archivo cargado genera su propio Excel. Si cargas varios, "
        "también podrás descargarlos todos juntos dentro de un ZIP."
    )


# ============================================================
# CARGA DE UNO O VARIOS ARCHIVOS
# ============================================================

archivos_subidos = st.file_uploader(
    "Selecciona uno o varios archivos",
    type=["zip", "txt", "csv"],
    accept_multiple_files=True,
    help=(
        "Puedes seleccionar un solo archivo o varios al mismo tiempo. "
        "También puedes arrastrarlos y soltarlos aquí."
    ),
)

if not archivos_subidos:
    st.info("Carga al menos un archivo para iniciar la conversión.")
    st.stop()


# ============================================================
# PROCESAMIENTO DE TODOS LOS ARCHIVOS
# ============================================================

resultados = []
errores = []
nombres_salida_usados = set()

barra_progreso = st.progress(0, text="Preparando conversión...")

for indice, archivo_subido in enumerate(archivos_subidos, start=1):
    barra_progreso.progress(
        (indice - 1) / len(archivos_subidos),
        text=f"Procesando {archivo_subido.name}...",
    )

    try:
        tablas = extraer_tablas(
            archivo_subido,
            separador_manual=separador_manual,
        )
        excel_bytes = crear_excel(tablas)

        nombre_base = limpiar_nombre_archivo(archivo_subido.name)
        nombre_deseado = f"{nombre_base}_convertido.xlsx"
        nombre_salida = nombre_unico(
            nombre_deseado,
            nombres_salida_usados,
        )

        resultados.append(
            {
                "nombre_entrada": archivo_subido.name,
                "nombre_salida": nombre_salida,
                "tablas": tablas,
                "excel_bytes": excel_bytes,
                "registros": sum(
                    len(tabla["dataframe"])
                    for tabla in tablas
                ),
                "columnas": sum(
                    len(tabla["dataframe"].columns)
                    for tabla in tablas
                ),
            }
        )

    except Exception as error:
        errores.append(
            {
                "archivo": archivo_subido.name,
                "error": str(error),
            }
        )

barra_progreso.progress(1.0, text="Conversión finalizada.")


# ============================================================
# RESUMEN
# ============================================================

col1, col2, col3, col4 = st.columns(4)

col1.metric("Archivos cargados", len(archivos_subidos))
col2.metric("Convertidos", len(resultados))
col3.metric("Con error", len(errores))
col4.metric(
    "Registros convertidos",
    f"{sum(resultado['registros'] for resultado in resultados):,}",
)

if resultados:
    st.success(
        f"Se convirtieron correctamente {len(resultados)} "
        f"de {len(archivos_subidos)} archivo(s)."
    )

if errores:
    st.warning(
        f"No fue posible convertir {len(errores)} archivo(s). "
        "Revisa el detalle al final de la página."
    )


# ============================================================
# VISTA PREVIA Y DESCARGAS INDIVIDUALES
# ============================================================

if resultados:
    st.subheader("🔎 Resultados y vista previa")

    for numero, resultado in enumerate(resultados, start=1):
        with st.expander(
            f"{numero}. {resultado['nombre_entrada']}",
            expanded=len(resultados) == 1,
        ):
            resumen1, resumen2, resumen3 = st.columns(3)
            resumen1.metric("Tablas encontradas", len(resultado["tablas"]))
            resumen2.metric("Registros", f"{resultado['registros']:,}")
            resumen3.metric("Columnas", resultado["columnas"])

            nombres_pestanas = [
                Path(tabla["nombre"]).name
                for tabla in resultado["tablas"]
            ]
            pestanas = st.tabs(nombres_pestanas)

            for pestana, tabla in zip(pestanas, resultado["tablas"]):
                with pestana:
                    dataframe = tabla["dataframe"]

                    st.write(
                        f"**Registros:** {len(dataframe):,}  |  "
                        f"**Columnas:** {len(dataframe.columns)}  |  "
                        f"**Codificación:** `{tabla['codificacion']}`  |  "
                        f"**Separador:** `{mostrar_separador(tabla['separador'])}`"
                    )

                    st.dataframe(
                        dataframe.head(100),
                        use_container_width=True,
                        hide_index=True,
                    )

                    if len(dataframe) > 100:
                        st.caption(
                            "La vista previa muestra las primeras 100 filas. "
                            "El Excel contiene todos los registros."
                        )

            st.download_button(
                label=f"Descargar {resultado['nombre_salida']}",
                data=resultado["excel_bytes"],
                file_name=resultado["nombre_salida"],
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                key=f"descarga_individual_{numero}",
                use_container_width=True,
            )


# ============================================================
# DESCARGA GENERAL
# ============================================================

if len(resultados) > 1:
    st.subheader("📦 Descargar todos")
    zip_resultados = crear_zip_resultados(resultados)

    st.download_button(
        label=f"Descargar los {len(resultados)} Excel en un ZIP",
        data=zip_resultados,
        file_name="archivos_convertidos_excel.zip",
        mime="application/zip",
        type="primary",
        key="descarga_todos_zip",
        use_container_width=True,
    )

elif len(resultados) == 1:
    st.info(
        "Se cargó un solo archivo. Puedes descargar su Excel desde "
        "la sección de resultados."
    )


# ============================================================
# DETALLE DE ERRORES
# ============================================================

if errores:
    st.subheader("⚠️ Archivos que no pudieron convertirse")

    for error in errores:
        with st.expander(error["archivo"]):
            st.code(error["error"], language=None)
