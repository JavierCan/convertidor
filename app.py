import io
import re
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
import streamlit as st
from defusedxml import ElementTree as ET
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Convertidor SAT a Excel",
    page_icon="📄",
    layout="wide",
)


# ============================================================
# ESTILOS
# ============================================================

st.markdown(
    """
    <style>
    /* Botones de descarga verdes */
    [data-testid="stDownloadButton"] button,
    div.stDownloadButton > button,
    div.stDownloadButton > button[kind="primary"],
    div.stDownloadButton > button[kind="secondary"] {
        background-color: #28a745 !important;
        color: #ffffff !important;
        border: 1px solid #28a745 !important;
        border-radius: 12px !important;
        font-weight: 700 !important;
        padding: 0.75rem 1rem !important;
        width: 100% !important;
        min-height: 48px !important;
        font-size: 1rem !important;
        transition: all 0.2s ease-in-out !important;
    }

    [data-testid="stDownloadButton"] button:hover,
    [data-testid="stDownloadButton"] button:focus,
    div.stDownloadButton > button:hover,
    div.stDownloadButton > button:focus {
        background-color: #218838 !important;
        border-color: #1e7e34 !important;
        color: #ffffff !important;
        box-shadow: 0 4px 12px rgba(40, 167, 69, 0.28) !important;
        transform: translateY(-1px) !important;
    }

    [data-testid="stDownloadButton"] button:active,
    div.stDownloadButton > button:active {
        background-color: #1e7e34 !important;
        border-color: #1e7e34 !important;
        color: #ffffff !important;
        transform: translateY(0) !important;
    }

    .texto-descarga {
        text-align: center;
        font-size: 0.95rem;
        color: #555555;
        margin-top: 0.35rem;
        margin-bottom: 1rem;
    }

    .firma-app {
        text-align: center;
        font-size: 0.98rem;
        color: #777777;
        margin-top: 2.5rem;
        margin-bottom: 1.25rem;
        font-style: italic;
    }

    .archivo-titulo {
        font-size: 1.08rem;
        font-weight: 700;
        margin-top: 0.25rem;
        margin-bottom: 0.5rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# ENCABEZADO
# ============================================================

st.title("📄 Convertidor de archivos SAT a Excel")
st.write(
    "Carga **uno o varios archivos ZIP, TXT, CSV o XML**. La aplicación "
    "genera un archivo **Excel (.xlsx)** por cada archivo cargado y también "
    "puede agrupar todos los Excel en un solo ZIP. Los XML CFDI se organizan "
    "automáticamente en hojas como resumen, emisor, receptor, conceptos, "
    "impuestos y detalle completo del XML."
)


# ============================================================
# CONSTANTES
# ============================================================

FORMATOS_COMPATIBLES = {".txt", ".csv", ".xml"}
MAX_FILAS_POR_HOJA = 1_048_575  # Se reserva una fila para encabezados.

ORDEN_TABLAS_XML = [
    "XML_Resumen",
    "XML_Emisor",
    "XML_Receptor",
    "XML_Conceptos",
    "XML_Impuestos",
    "XML_Relacionados",
    "XML_Timbre",
    "XML_Complementos",
    "XML_Pagos",
    "XML_DoctosPago",
    "XML_TotalesPago",
    "XML_Detalle",
]


# ============================================================
# FUNCIONES GENERALES
# ============================================================


def limpiar_nombre_hoja(nombre: str, usados: Set[str]) -> str:
    """Genera un nombre de hoja válido y único para Excel."""
    nombre_limpio = Path(nombre).stem

    for caracter in ['\\', '/', '*', '?', ':', '[', ']']:
        nombre_limpio = nombre_limpio.replace(caracter, "_")

    nombre_limpio = nombre_limpio.strip() or "Hoja"
    nombre_limpio = nombre_limpio[:31]

    nombre_base = nombre_limpio
    contador = 1

    while nombre_limpio.lower() in {item.lower() for item in usados}:
        sufijo = f"_{contador}"
        nombre_limpio = f"{nombre_base[:31 - len(sufijo)]}{sufijo}"
        contador += 1

    usados.add(nombre_limpio)
    return nombre_limpio


def detectar_separador(contenido: bytes, codificacion: str) -> str:
    """Detecta el separador más probable a partir de las primeras líneas."""
    texto = contenido.decode(codificacion, errors="ignore")
    lineas = [linea for linea in texto.splitlines()[:10] if linea.strip()]

    if not lineas:
        return "~"

    candidatos = ["~", "|", ";", "\t", ","]
    puntuaciones: Dict[str, int] = {}

    for separador in candidatos:
        conteos = [linea.count(separador) for linea in lineas]
        no_cero = [conteo for conteo in conteos if conteo > 0]

        if not no_cero:
            puntuaciones[separador] = 0
            continue

        consistencia = len(no_cero) * 100
        frecuencia = sum(no_cero)
        variacion = max(no_cero) - min(no_cero)
        puntuaciones[separador] = consistencia + frecuencia - variacion

    mejor = max(puntuaciones, key=puntuaciones.get)
    return mejor if puntuaciones[mejor] > 0 else "~"


def leer_tabla(
    contenido: bytes,
    nombre_archivo: str,
    separador_manual: Optional[str] = None,
) -> Dict:
    """
    Lee un TXT o CSV probando varias codificaciones.
    Mantiene todas las columnas como texto para proteger UUID, RFC y claves.
    """
    codificaciones = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]
    errores: List[str] = []

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

            columnas_validas: List[str] = []
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

            if dataframe.empty:
                raise ValueError("El archivo no contiene registros de datos.")

            return {
                "nombre": nombre_archivo,
                "archivo_origen": nombre_archivo,
                "dataframe": dataframe,
                "codificacion": codificacion,
                "separador": separador,
                "formato": Path(nombre_archivo).suffix.upper().replace(".", ""),
            }

        except Exception as error:
            errores.append(f"{codificacion}: {error}")

    detalle = "\n".join(errores)
    raise ValueError(
        f"No fue posible leer {nombre_archivo}.\n\n{detalle}"
    )


# ============================================================
# FUNCIONES XML
# ============================================================


def nombre_local(valor) -> str:
    """Elimina el namespace de una etiqueta o atributo XML."""
    texto = str(valor)
    if "}" in texto:
        texto = texto.split("}", 1)[1]
    if ":" in texto:
        texto = texto.split(":", 1)[1]
    return texto


def detectar_codificacion_xml(contenido: bytes) -> str:
    """Obtiene la codificación declarada en el encabezado XML."""
    coincidencia = re.search(
        br"<\?xml[^>]*encoding=[\"']([^\"']+)[\"']",
        contenido[:500],
        flags=re.IGNORECASE,
    )
    if coincidencia:
        return coincidencia.group(1).decode("ascii", errors="replace")
    return "No declarada"


def atributos_elemento(elemento) -> Dict[str, str]:
    """Devuelve los atributos XML sin namespaces."""
    resultado: Dict[str, str] = {}
    repetidos: Dict[str, int] = defaultdict(int)

    for clave, valor in elemento.attrib.items():
        clave_limpia = nombre_local(clave)
        repetidos[clave_limpia] += 1

        if repetidos[clave_limpia] > 1:
            clave_limpia = f"{clave_limpia}_{repetidos[clave_limpia]}"

        resultado[clave_limpia] = str(valor)

    return resultado


def buscar_elementos(raiz, etiqueta: str) -> List:
    """Busca elementos por nombre local ignorando namespaces y mayúsculas."""
    objetivo = etiqueta.lower()
    return [
        elemento
        for elemento in raiz.iter()
        if nombre_local(elemento.tag).lower() == objetivo
    ]


def buscar_hijo_directo(elemento, etiqueta: str):
    """Busca el primer hijo directo por nombre local."""
    objetivo = etiqueta.lower()
    for hijo in list(elemento):
        if nombre_local(hijo.tag).lower() == objetivo:
            return hijo
    return None


def construir_detalle_xml(raiz, nombre_archivo: str) -> pd.DataFrame:
    """
    Genera una tabla vertical que conserva todos los atributos y textos
    encontrados en el XML, incluso cuando no es un CFDI estándar.
    """
    filas: List[Dict[str, str]] = []

    def recorrer(elemento, ruta_padre: str, indice_hermano: int) -> None:
        etiqueta = nombre_local(elemento.tag)
        ruta_actual = f"{ruta_padre}/{etiqueta}[{indice_hermano}]"

        for atributo, valor in atributos_elemento(elemento).items():
            filas.append(
                {
                    "Archivo_XML": nombre_archivo,
                    "Ruta": ruta_actual,
                    "Elemento": etiqueta,
                    "Tipo": "Atributo",
                    "Campo": atributo,
                    "Valor": valor,
                }
            )

        texto = (elemento.text or "").strip()
        if texto:
            filas.append(
                {
                    "Archivo_XML": nombre_archivo,
                    "Ruta": ruta_actual,
                    "Elemento": etiqueta,
                    "Tipo": "Texto",
                    "Campo": "Texto",
                    "Valor": texto,
                }
            )

        contadores_hijos: Dict[str, int] = defaultdict(int)
        for hijo in list(elemento):
            etiqueta_hijo = nombre_local(hijo.tag)
            contadores_hijos[etiqueta_hijo] += 1
            recorrer(
                hijo,
                ruta_actual,
                contadores_hijos[etiqueta_hijo],
            )

    recorrer(raiz, "", 1)

    return pd.DataFrame(
        filas,
        columns=[
            "Archivo_XML",
            "Ruta",
            "Elemento",
            "Tipo",
            "Campo",
            "Valor",
        ],
    )


def parsear_xml(contenido: bytes, nombre_archivo: str) -> List[Dict]:
    """
    Convierte un XML CFDI —o un XML genérico— en varias tablas de Excel.
    Siempre crea un resumen y un detalle completo para evitar pérdida de datos.
    """
    try:
        raiz = ET.fromstring(contenido)
    except ET.ParseError as error:
        raise ValueError(
            f"El XML {nombre_archivo} no es válido o está incompleto: {error}"
        ) from error
    except Exception as error:
        raise ValueError(
            f"No fue posible procesar el XML {nombre_archivo}: {error}"
        ) from error

    codificacion = detectar_codificacion_xml(contenido)
    tablas_datos: Dict[str, List[Dict]] = defaultdict(list)

    # --------------------------------------------------------
    # RESUMEN GENERAL DEL CFDI / XML
    # --------------------------------------------------------

    resumen: Dict[str, str] = {
        "Archivo_XML": nombre_archivo,
        "Elemento_Raiz": nombre_local(raiz.tag),
        "Codificacion_XML": codificacion,
    }

    for clave, valor in atributos_elemento(raiz).items():
        resumen[f"CFDI_{clave}"] = valor

    emisores = buscar_elementos(raiz, "Emisor")
    receptores = buscar_elementos(raiz, "Receptor")
    timbres = buscar_elementos(raiz, "TimbreFiscalDigital")

    if emisores:
        for clave, valor in atributos_elemento(emisores[0]).items():
            resumen[f"Emisor_{clave}"] = valor

    if receptores:
        for clave, valor in atributos_elemento(receptores[0]).items():
            resumen[f"Receptor_{clave}"] = valor

    if timbres:
        for clave, valor in atributos_elemento(timbres[0]).items():
            resumen[f"Timbre_{clave}"] = valor

    tablas_datos["XML_Resumen"].append(resumen)

    # --------------------------------------------------------
    # EMISOR, RECEPTOR Y TIMBRE
    # --------------------------------------------------------

    for numero, emisor in enumerate(emisores, start=1):
        fila = {
            "Archivo_XML": nombre_archivo,
            "Emisor_Numero": numero,
        }
        fila.update(atributos_elemento(emisor))
        tablas_datos["XML_Emisor"].append(fila)

    for numero, receptor in enumerate(receptores, start=1):
        fila = {
            "Archivo_XML": nombre_archivo,
            "Receptor_Numero": numero,
        }
        fila.update(atributos_elemento(receptor))
        tablas_datos["XML_Receptor"].append(fila)

    for numero, timbre in enumerate(timbres, start=1):
        fila = {
            "Archivo_XML": nombre_archivo,
            "Timbre_Numero": numero,
        }
        fila.update(atributos_elemento(timbre))
        tablas_datos["XML_Timbre"].append(fila)

    # --------------------------------------------------------
    # CONCEPTOS E IMPUESTOS POR CONCEPTO
    # --------------------------------------------------------

    conceptos = buscar_elementos(raiz, "Concepto")

    for numero_concepto, concepto in enumerate(conceptos, start=1):
        fila_concepto = {
            "Archivo_XML": nombre_archivo,
            "Concepto_Numero": numero_concepto,
        }
        fila_concepto.update(atributos_elemento(concepto))
        tablas_datos["XML_Conceptos"].append(fila_concepto)

        for impuesto in concepto.iter():
            tipo_impuesto = nombre_local(impuesto.tag)
            if tipo_impuesto.lower() not in {"traslado", "retencion"}:
                continue

            fila_impuesto = {
                "Archivo_XML": nombre_archivo,
                "Nivel": "Concepto",
                "Concepto_Numero": numero_concepto,
                "Tipo_Movimiento": tipo_impuesto,
            }
            fila_impuesto.update(atributos_elemento(impuesto))
            tablas_datos["XML_Impuestos"].append(fila_impuesto)

    # --------------------------------------------------------
    # IMPUESTOS GLOBALES DEL COMPROBANTE
    # --------------------------------------------------------

    impuestos_globales = buscar_hijo_directo(raiz, "Impuestos")
    if impuestos_globales is not None:
        for impuesto in impuestos_globales.iter():
            tipo_impuesto = nombre_local(impuesto.tag)
            if tipo_impuesto.lower() not in {"traslado", "retencion"}:
                continue

            fila_impuesto = {
                "Archivo_XML": nombre_archivo,
                "Nivel": "Comprobante",
                "Concepto_Numero": "",
                "Tipo_Movimiento": tipo_impuesto,
            }
            fila_impuesto.update(atributos_elemento(impuesto))
            tablas_datos["XML_Impuestos"].append(fila_impuesto)

    # --------------------------------------------------------
    # CFDI RELACIONADOS
    # --------------------------------------------------------

    for numero, relacionado in enumerate(
        buscar_elementos(raiz, "CfdiRelacionado"),
        start=1,
    ):
        fila = {
            "Archivo_XML": nombre_archivo,
            "Relacionado_Numero": numero,
        }
        fila.update(atributos_elemento(relacionado))
        tablas_datos["XML_Relacionados"].append(fila)

    # --------------------------------------------------------
    # COMPLEMENTOS
    # --------------------------------------------------------

    numero_complemento = 0
    for complemento in buscar_elementos(raiz, "Complemento"):
        for hijo in list(complemento):
            numero_complemento += 1
            fila = {
                "Archivo_XML": nombre_archivo,
                "Complemento_Numero": numero_complemento,
                "Tipo_Complemento": nombre_local(hijo.tag),
            }
            fila.update(atributos_elemento(hijo))
            tablas_datos["XML_Complementos"].append(fila)

    # --------------------------------------------------------
    # COMPLEMENTO DE PAGOS
    # --------------------------------------------------------

    pagos = buscar_elementos(raiz, "Pago")
    for numero_pago, pago in enumerate(pagos, start=1):
        fila_pago = {
            "Archivo_XML": nombre_archivo,
            "Pago_Numero": numero_pago,
        }
        fila_pago.update(atributos_elemento(pago))
        tablas_datos["XML_Pagos"].append(fila_pago)

        numero_documento = 0
        for documento in pago.iter():
            if nombre_local(documento.tag).lower() != "doctorelacionado":
                continue

            numero_documento += 1
            fila_documento = {
                "Archivo_XML": nombre_archivo,
                "Pago_Numero": numero_pago,
                "Documento_Numero": numero_documento,
            }
            fila_documento.update(atributos_elemento(documento))
            tablas_datos["XML_DoctosPago"].append(fila_documento)

    for numero, totales in enumerate(
        buscar_elementos(raiz, "Totales"),
        start=1,
    ):
        fila = {
            "Archivo_XML": nombre_archivo,
            "Totales_Numero": numero,
        }
        fila.update(atributos_elemento(totales))
        tablas_datos["XML_TotalesPago"].append(fila)

    # --------------------------------------------------------
    # DETALLE COMPLETO DEL XML
    # --------------------------------------------------------

    detalle = construir_detalle_xml(raiz, nombre_archivo)

    # --------------------------------------------------------
    # CONSTRUIR SALIDA
    # --------------------------------------------------------

    tablas: List[Dict] = []

    for nombre_tabla in ORDEN_TABLAS_XML:
        if nombre_tabla == "XML_Detalle":
            dataframe = detalle
        else:
            filas = tablas_datos.get(nombre_tabla, [])
            if not filas:
                continue
            dataframe = pd.DataFrame(filas).fillna("")

        if dataframe.empty:
            continue

        tablas.append(
            {
                "nombre": nombre_tabla,
                "archivo_origen": nombre_archivo,
                "dataframe": dataframe.astype(str),
                "codificacion": codificacion,
                "separador": None,
                "formato": "XML",
            }
        )

    return tablas


def consolidar_tablas_xml(tablas_xml: List[Dict]) -> List[Dict]:
    """Consolida tablas del mismo tipo cuando un ZIP contiene varios XML."""
    acumuladas: Dict[str, List[pd.DataFrame]] = defaultdict(list)
    archivos_por_tabla: Dict[str, List[str]] = defaultdict(list)
    codificaciones_por_tabla: Dict[str, Set[str]] = defaultdict(set)

    for tabla in tablas_xml:
        nombre = tabla["nombre"]
        acumuladas[nombre].append(tabla["dataframe"])
        archivos_por_tabla[nombre].append(tabla["archivo_origen"])
        codificaciones_por_tabla[nombre].add(tabla["codificacion"])

    resultado: List[Dict] = []

    nombres_ordenados = [
        nombre
        for nombre in ORDEN_TABLAS_XML
        if nombre in acumuladas
    ]

    for nombre in nombres_ordenados:
        dataframe = pd.concat(
            acumuladas[nombre],
            ignore_index=True,
            sort=False,
        ).fillna("")

        resultado.append(
            {
                "nombre": nombre,
                "archivo_origen": (
                    f"{len(set(archivos_por_tabla[nombre]))} XML dentro del ZIP"
                ),
                "dataframe": dataframe.astype(str),
                "codificacion": ", ".join(
                    sorted(codificaciones_por_tabla[nombre])
                ),
                "separador": None,
                "formato": "XML",
            }
        )

    return resultado


# ============================================================
# EXTRACCIÓN DE ARCHIVOS
# ============================================================


def extraer_tablas(
    archivo_subido,
    separador_manual: Optional[str] = None,
) -> List[Dict]:
    """Procesa un ZIP, TXT, CSV o XML y devuelve las tablas encontradas."""
    nombre_archivo = archivo_subido.name
    contenido = archivo_subido.getvalue()
    extension = Path(nombre_archivo).suffix.lower()
    tablas: List[Dict] = []

    if extension == ".zip":
        try:
            with zipfile.ZipFile(io.BytesIO(contenido), "r") as archivo_zip:
                archivos_internos = [
                    nombre
                    for nombre in archivo_zip.namelist()
                    if not nombre.endswith("/")
                    and Path(nombre).suffix.lower() in FORMATOS_COMPATIBLES
                ]

                if not archivos_internos:
                    raise ValueError(
                        "El ZIP no contiene archivos TXT, CSV o XML compatibles."
                    )

                tablas_xml: List[Dict] = []

                for archivo_interno in archivos_internos:
                    contenido_interno = archivo_zip.read(archivo_interno)
                    extension_interna = Path(archivo_interno).suffix.lower()

                    if extension_interna == ".xml":
                        tablas_xml.extend(
                            parsear_xml(contenido_interno, archivo_interno)
                        )
                    else:
                        tablas.append(
                            leer_tabla(
                                contenido_interno,
                                archivo_interno,
                                separador_manual,
                            )
                        )

                if tablas_xml:
                    tablas.extend(consolidar_tablas_xml(tablas_xml))

        except zipfile.BadZipFile as error:
            raise ValueError(
                "El archivo ZIP está dañado o no es un ZIP válido."
            ) from error

    elif extension in {".txt", ".csv"}:
        tablas.append(
            leer_tabla(
                contenido,
                nombre_archivo,
                separador_manual,
            )
        )

    elif extension == ".xml":
        tablas.extend(parsear_xml(contenido, nombre_archivo))

    else:
        raise ValueError(
            "Formato no compatible. Selecciona archivos ZIP, TXT, CSV o XML."
        )

    if not tablas:
        raise ValueError(
            "No se encontraron datos que pudieran convertirse a Excel."
        )

    return tablas


# ============================================================
# CREACIÓN DEL EXCEL
# ============================================================


def ajustar_hoja_excel(worksheet, dataframe: pd.DataFrame) -> None:
    """Aplica formato básico y conserva los valores como texto."""
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    worksheet.sheet_view.showGridLines = False

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

    worksheet.row_dimensions[1].height = 24

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


def dividir_dataframe(dataframe: pd.DataFrame) -> List[pd.DataFrame]:
    """Divide tablas muy grandes para respetar el límite de filas de Excel."""
    if len(dataframe) <= MAX_FILAS_POR_HOJA:
        return [dataframe]

    return [
        dataframe.iloc[inicio:inicio + MAX_FILAS_POR_HOJA].copy()
        for inicio in range(0, len(dataframe), MAX_FILAS_POR_HOJA)
    ]


def crear_excel(tablas: List[Dict]) -> bytes:
    """Crea un XLSX en memoria con una o varias hojas por tabla."""
    buffer_salida = io.BytesIO()
    nombres_usados: Set[str] = set()

    with pd.ExcelWriter(buffer_salida, engine="openpyxl") as writer:
        for tabla in tablas:
            dataframe = tabla["dataframe"].fillna("").astype(str)
            fragmentos = dividir_dataframe(dataframe)

            for numero_fragmento, fragmento in enumerate(fragmentos, start=1):
                nombre_base = tabla["nombre"]
                if len(fragmentos) > 1:
                    nombre_base = f"{nombre_base}_{numero_fragmento}"

                nombre_hoja = limpiar_nombre_hoja(
                    nombre_base,
                    nombres_usados,
                )

                fragmento.to_excel(
                    writer,
                    sheet_name=nombre_hoja,
                    index=False,
                )

                ajustar_hoja_excel(
                    writer.sheets[nombre_hoja],
                    fragmento,
                )

    buffer_salida.seek(0)
    return buffer_salida.getvalue()


def nombre_salida(nombre_entrada: str) -> str:
    """Genera el nombre del Excel resultante."""
    return f"{Path(nombre_entrada).stem}_convertido.xlsx"


def crear_zip_resultados(resultados: List[Dict]) -> bytes:
    """Agrupa todos los Excel generados en un ZIP."""
    buffer_zip = io.BytesIO()
    nombres_usados: Set[str] = set()

    with zipfile.ZipFile(
        buffer_zip,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
    ) as archivo_zip:
        for resultado in resultados:
            nombre = resultado["nombre_salida"]
            nombre_base = Path(nombre).stem
            extension = Path(nombre).suffix
            nombre_unico = nombre
            contador = 1

            while nombre_unico.lower() in nombres_usados:
                nombre_unico = f"{nombre_base}_{contador}{extension}"
                contador += 1

            nombres_usados.add(nombre_unico.lower())
            archivo_zip.writestr(
                nombre_unico,
                resultado["excel_bytes"],
            )

    buffer_zip.seek(0)
    return buffer_zip.getvalue()


def mostrar_separador(separador: Optional[str]) -> str:
    """Convierte el separador a una etiqueta legible."""
    if separador is None:
        return "No aplica"
    if separador == "\t":
        return "TAB"
    return separador


def texto_metadatos_tabla(tabla: Dict, dataframe: pd.DataFrame) -> str:
    """Crea la descripción mostrada arriba de cada vista previa."""
    partes = [
        f"**Origen:** {tabla['archivo_origen']}",
        f"**Formato:** `{tabla['formato']}`",
        f"**Registros:** {len(dataframe):,}",
        f"**Columnas:** {len(dataframe.columns)}",
    ]

    if tabla["formato"] == "XML":
        partes.append(f"**Codificación XML:** `{tabla['codificacion']}`")
    else:
        partes.append(f"**Codificación:** `{tabla['codificacion']}`")
        partes.append(
            f"**Separador:** `{mostrar_separador(tabla['separador'])}`"
        )

    return "  |  ".join(partes)


# ============================================================
# BARRA LATERAL
# ============================================================

with st.sidebar:
    st.header("⚙️ Configuración")

    modo_separador = st.radio(
        "Separador de columnas para TXT/CSV",
        options=[
            "Detectar automáticamente",
            "Usar separador personalizado",
        ],
        index=0,
    )

    separador_manual: Optional[str] = None

    if modo_separador == "Usar separador personalizado":
        separador_manual = st.text_input(
            "Separador",
            value="~",
            max_chars=3,
            help="Para estos archivos del SAT normalmente se usa ~.",
        )

        if separador_manual == r"\t":
            separador_manual = "\t"

        if separador_manual == "":
            separador_manual = None

    st.markdown("---")
    st.caption(
        "El separador solo se aplica a TXT y CSV. Los XML se detectan y "
        "procesan automáticamente. Los valores se conservan como texto para "
        "evitar que Excel modifique UUID, RFC, códigos, números largos o "
        "ceros iniciales."
    )


# ============================================================
# CARGA DE UNO O VARIOS ARCHIVOS
# ============================================================

archivos_subidos = st.file_uploader(
    "Selecciona uno o varios archivos",
    type=["zip", "txt", "csv", "xml"],
    accept_multiple_files=True,
    help=(
        "Puedes cargar uno o varios ZIP, TXT, CSV o XML. Cada archivo cargado "
        "se convertirá en un Excel independiente. Si un ZIP contiene varios "
        "XML, se consolidarán dentro de un solo Excel."
    ),
)

if not archivos_subidos:
    st.info("Carga al menos un archivo para iniciar la conversión.")
    st.stop()


# ============================================================
# PROCESAMIENTO INDEPENDIENTE DE CADA ARCHIVO
# ============================================================

resultados: List[Dict] = []
errores: List[Dict] = []

barra_progreso = st.progress(0)
texto_progreso = st.empty()

total_archivos = len(archivos_subidos)

for indice, archivo_subido in enumerate(archivos_subidos, start=1):
    texto_progreso.write(
        f"Procesando {indice} de {total_archivos}: **{archivo_subido.name}**"
    )

    try:
        tablas = extraer_tablas(
            archivo_subido,
            separador_manual=separador_manual,
        )
        excel_bytes = crear_excel(tablas)

        formatos_detectados = sorted(
            {tabla["formato"] for tabla in tablas}
        )

        resultados.append(
            {
                "nombre_entrada": archivo_subido.name,
                "nombre_salida": nombre_salida(archivo_subido.name),
                "tablas": tablas,
                "excel_bytes": excel_bytes,
                "formatos": ", ".join(formatos_detectados),
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
                "nombre": archivo_subido.name,
                "error": str(error),
            }
        )

    barra_progreso.progress(indice / total_archivos)

texto_progreso.empty()
barra_progreso.empty()


# ============================================================
# RESUMEN
# ============================================================

st.subheader("📊 Resumen del procesamiento")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Archivos cargados", total_archivos)
col2.metric("Convertidos", len(resultados))
col3.metric("Con error", len(errores))
col4.metric(
    "Registros totales",
    f"{sum(resultado['registros'] for resultado in resultados):,}",
)

if resultados:
    st.success(
        f"Se convirtieron correctamente {len(resultados)} "
        f"de {total_archivos} archivo(s)."
    )

if errores:
    st.warning(
        "Algunos archivos no pudieron convertirse. Los demás resultados "
        "siguen disponibles para descargar."
    )

    with st.expander("Ver archivos con error", expanded=True):
        for error in errores:
            st.error(f"**{error['nombre']}**\n\n{error['error']}")


# ============================================================
# RESULTADOS Y DESCARGAS INDIVIDUALES
# ============================================================

if resultados:
    st.subheader("⬇️ Descargar resultados")

    for numero, resultado in enumerate(resultados, start=1):
        with st.container(border=True):
            st.markdown(
                f'<div class="archivo-titulo">{numero}. '
                f'{resultado["nombre_entrada"]}</div>',
                unsafe_allow_html=True,
            )

            resumen1, resumen2, resumen3, resumen4 = st.columns(4)
            resumen1.metric("Formato", resultado["formatos"])
            resumen2.metric("Registros", f"{resultado['registros']:,}")
            resumen3.metric("Columnas", resultado["columnas"])
            resumen4.metric("Hojas de Excel", len(resultado["tablas"]))

            with st.expander("🔎 Ver vista previa"):
                if len(resultado["tablas"]) == 1:
                    tabla = resultado["tablas"][0]
                    dataframe = tabla["dataframe"]

                    st.write(texto_metadatos_tabla(tabla, dataframe))
                    st.dataframe(
                        dataframe.head(100),
                        use_container_width=True,
                        hide_index=True,
                    )

                else:
                    nombres_pestanas: List[str] = []
                    nombres_vistos: Dict[str, int] = defaultdict(int)

                    for tabla in resultado["tablas"]:
                        nombre_pestana = tabla["nombre"][:36]
                        nombres_vistos[nombre_pestana] += 1

                        if nombres_vistos[nombre_pestana] > 1:
                            nombre_pestana = (
                                f"{nombre_pestana[:32]} "
                                f"({nombres_vistos[nombre_pestana]})"
                            )

                        nombres_pestanas.append(nombre_pestana)

                    pestanas = st.tabs(nombres_pestanas)

                    for pestana, tabla in zip(
                        pestanas,
                        resultado["tablas"],
                    ):
                        with pestana:
                            dataframe = tabla["dataframe"]
                            st.write(texto_metadatos_tabla(tabla, dataframe))
                            st.dataframe(
                                dataframe.head(100),
                                use_container_width=True,
                                hide_index=True,
                            )

                if resultado["registros"] > 100:
                    st.caption(
                        "La vista previa muestra como máximo las primeras "
                        "100 filas de cada hoja. El Excel incluye todos los datos."
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

            st.markdown(
                '<div class="texto-descarga">⬆️ Aquí debe descargarlo ella</div>',
                unsafe_allow_html=True,
            )


# ============================================================
# DESCARGA CONJUNTA
# ============================================================

if len(resultados) > 1:
    st.subheader("📦 Descargar todos")
    zip_resultados = crear_zip_resultados(resultados)

    st.download_button(
        label="Descargar todos los Excel en un solo ZIP",
        data=zip_resultados,
        file_name="archivos_convertidos_excel.zip",
        mime="application/zip",
        key="descarga_todos_zip",
        use_container_width=True,
    )

    st.markdown(
        '<div class="texto-descarga">'
        '⬆️ Aquí puede descargar todos los archivos de una sola vez'
        '</div>',
        unsafe_allow_html=True,
    )


# ============================================================
# FIRMA
# ============================================================

st.markdown(
    '<div class="firma-app">Hecho por tu novio el ingeniero xd 😎</div>',
    unsafe_allow_html=True,
)
