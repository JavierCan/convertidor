from __future__ import annotations

import io
import re
import time
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Iterable
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="CFDI a Excel", page_icon="🧾", layout="wide")
BASE_DIR = Path(__file__).parent
ASSETS = BASE_DIR / "assets"
PREVIEW_XML_LIMIT = 20
PREVIEW_ROW_LIMIT = 25
EXCEL_MAX_ROWS = 1_048_576
TECHNICAL_FIELDS = {"Sello", "Certificado", "SelloCFD", "SelloSAT"}


def load_asset(name: str) -> str:
    p = ASSETS / name
    return p.read_text(encoding="utf-8") if p.exists() else ""


html = load_asset("index.html")
html = html.replace("/*__INLINE_CSS__*/", load_asset("styles.css"))
html = html.replace("/*__INLINE_JS__*/", load_asset("script.js"))
components.html(html, height=300, scrolling=False)

st.markdown("""
<style>
.block-container{padding-top:1rem;padding-bottom:3rem;max-width:1500px}
div[data-testid="stFileUploader"]{border:2px dashed #16a34a;border-radius:18px;padding:10px;background:#f0fdf4}
div.stButton>button,div.stDownloadButton>button{width:100%;border-radius:12px;font-weight:700;min-height:46px}
div.stButton>button[kind="primary"],div.stDownloadButton>button{background:linear-gradient(135deg,#16a34a,#15803d)!important;color:white!important;border:none!important}
div.stButton>button[kind="primary"]:hover,div.stDownloadButton>button:hover{background:linear-gradient(135deg,#15803d,#166534)!important;color:white!important}
.signature{text-align:center;color:#64748b;font-size:.95rem;margin-top:1rem;font-style:italic}
</style>
""", unsafe_allow_html=True)


def local_name(tag: str) -> str:
    return tag.split("}")[-1] if "}" in tag else tag.split(":")[-1]


def normalize_col(name: str) -> str:
    name = re.sub(r"\s+", "_", str(name).strip())
    name = re.sub(r"[^A-Za-z0-9_áéíóúÁÉÍÓÚñÑ.-]", "_", name)
    return re.sub(r"_+", "_", name).strip("_")


def identify_xml_type(root: ET.Element) -> str:
    root_name = local_name(root.tag)
    version = root.attrib.get("Version") or root.attrib.get("version") or ""
    names = {local_name(el.tag) for el in root.iter()}
    if root_name == "Comprobante":
        if {"Pagos", "Pago", "DoctoRelacionado"} & names:
            return f"CFDI {version or 'sin versión'} - Complemento de pago"
        return f"CFDI {version or 'sin versión'}"
    return f"XML genérico - {root_name}"


def collect_flat(
    element: ET.Element,
    prefix: str = "",
    skip_tags: set[str] | None = None,
    exclude_technical: bool = True,
) -> dict[str, str]:
    """Aplana atributos y textos del XML sin depender de campos específicos."""
    skip_tags = skip_tags or set()
    result: dict[str, str] = {}
    name = local_name(element.tag)
    if name in skip_tags:
        return result
    path = normalize_col(f"{prefix}_{name}" if prefix else name)

    for attr, value in element.attrib.items():
        attr_name = local_name(attr)
        if exclude_technical and attr_name in TECHNICAL_FIELDS:
            continue
        result[normalize_col(f"{path}_{attr_name}")] = value

    text = (element.text or "").strip()
    if text and len(element) == 0:
        result[normalize_col(f"{path}_Texto")] = text

    counts = Counter(local_name(child.tag) for child in element)
    occurrences: defaultdict[str, int] = defaultdict(int)

    for child in element:
        child_name = local_name(child.tag)
        occurrences[child_name] += 1
        child_prefix = path
        if counts[child_name] > 1:
            child_prefix = normalize_col(f"{path}_{child_name}_{occurrences[child_name]}")
            child_prefix = child_prefix.rsplit("_", 1)[0]

        child_data = collect_flat(
            child,
            prefix=child_prefix,
            skip_tags=skip_tags,
            exclude_technical=exclude_technical,
        )
        for key, value in child_data.items():
            candidate = key
            idx = 2
            while candidate in result and result[candidate] != value:
                candidate = f"{key}_{idx}"
                idx += 1
            result[candidate] = value

    return result


def find_elements(root: ET.Element, tag_name: str) -> list[ET.Element]:
    return [el for el in root.iter() if local_name(el.tag) == tag_name]


def parse_xml(
    raw: bytes,
    source_name: str,
    representation: str,
    exclude_technical: bool,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    try:
        root = ET.fromstring(raw)
    except ET.ParseError as exc:
        raise ValueError(f"XML inválido: {exc}") from exc

    xml_type = identify_xml_type(root)
    concepts = find_elements(root, "Concepto")
    base = {"Archivo_XML": source_name, "Tipo_XML": xml_type}
    base.update(collect_flat(root, skip_tags={"Conceptos"}, exclude_technical=exclude_technical))

    meta = {"tipo": xml_type, "conceptos": len(concepts)}

    if representation == "por_comprobante":
        row = dict(base)
        row["Total_Conceptos"] = len(concepts)
        for idx, concept in enumerate(concepts, start=1):
            flat = collect_flat(concept, prefix=f"Concepto_{idx}", exclude_technical=exclude_technical)
            row.update(flat)
        return [row], meta

    if not concepts:
        return [base], meta

    rows = []
    for idx, concept in enumerate(concepts, start=1):
        row = dict(base)
        row["Concepto_Indice"] = idx
        row.update(collect_flat(concept, exclude_technical=exclude_technical))
        rows.append(row)
    return rows, meta


def iter_xmls(uploaded_files) -> Iterable[tuple[str, bytes]]:
    for uploaded in uploaded_files:
        name = uploaded.name
        raw = uploaded.getvalue()
        suffix = Path(name).suffix.lower()
        if suffix == ".xml":
            yield name, raw
        elif suffix == ".zip":
            try:
                with zipfile.ZipFile(io.BytesIO(raw), "r") as zf:
                    for member in zf.infolist():
                        if not member.is_dir() and Path(member.filename).suffix.lower() == ".xml":
                            yield f"{Path(name).stem}/{member.filename}", zf.read(member)
            except zipfile.BadZipFile as exc:
                raise ValueError(f"{name}: ZIP inválido") from exc


def analyze(uploaded_files) -> dict[str, Any]:
    total = valid = concepts = 0
    types: Counter[str] = Counter()
    errors = []
    for name, raw in iter_xmls(uploaded_files):
        total += 1
        try:
            root = ET.fromstring(raw)
            valid += 1
            types[identify_xml_type(root)] += 1
            concepts += len(find_elements(root, "Concepto"))
        except Exception as exc:
            errors.append({"Archivo": name, "Error": str(exc)})
    return {"total": total, "valid": valid, "concepts": concepts, "types": dict(types), "errors": errors}


def make_preview(uploaded_files, representation, exclude_technical, sample_mode):
    xmls = list(iter_xmls(uploaded_files))
    if sample_mode == "Aleatorios" and len(xmls) > PREVIEW_XML_LIMIT:
        sample = pd.Series(xmls).sample(PREVIEW_XML_LIMIT, random_state=42).tolist()
    else:
        sample = xmls[:PREVIEW_XML_LIMIT]

    rows, errors = [], []
    for name, raw in sample:
        try:
            parsed, _ = parse_xml(raw, name, representation, exclude_technical)
            rows.extend(parsed)
        except Exception as exc:
            errors.append({"Archivo": name, "Error": str(exc)})
    return pd.DataFrame(rows), errors, len(sample)


def sanitize_sheet(name: str) -> str:
    return (re.sub(r"[\[\]\*\?/\\:]", "_", name)[:31] or "Hoja").strip()


def format_sheet(ws, headers: list[str]):
    fill = PatternFill("solid", fgColor="166534")
    font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, header in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = min(max(len(str(header)) + 2, 12), 40)


def rows_to_excel(rows: list[dict[str, Any]], sheet_name="CFDI_Consolidado", errors=None) -> bytes:
    columns, seen = [], set()
    for row in rows:
        for col in row:
            if col not in seen:
                seen.add(col)
                columns.append(col)

    wb = Workbook()
    wb.remove(wb.active)

    if not rows:
        ws = wb.create_sheet("Sin_datos")
        ws.append(["Mensaje"])
        ws.append(["No se generaron filas"])
    else:
        ws = None
        row_count = 0
        sheet_idx = 1
        for row in rows:
            if ws is None or row_count >= EXCEL_MAX_ROWS - 1:
                name = sheet_name if sheet_idx == 1 else f"{sheet_name}_{sheet_idx}"
                ws = wb.create_sheet(sanitize_sheet(name))
                ws.append(columns)
                format_sheet(ws, columns)
                row_count = 0
                sheet_idx += 1
            ws.append([row.get(col, "") for col in columns])
            row_count += 1

    if errors:
        ws = wb.create_sheet("Errores")
        headers = ["Archivo", "Error"]
        ws.append(headers)
        for err in errors:
            ws.append([err.get("Archivo", ""), err.get("Error", "")])
        format_sheet(ws, headers)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def process_consolidated(uploaded_files, representation, exclude_technical, progress, status):
    xmls = list(iter_xmls(uploaded_files))
    rows, errors = [], []
    for idx, (name, raw) in enumerate(xmls, start=1):
        try:
            parsed, _ = parse_xml(raw, name, representation, exclude_technical)
            rows.extend(parsed)
        except Exception as exc:
            errors.append({"Archivo": name, "Error": str(exc)})
        progress.progress(idx / len(xmls))
        status.caption(f"Procesando {idx:,} de {len(xmls):,}: {name}")
    return rows, errors


def process_individual(uploaded_files, representation, exclude_technical, progress, status):
    xmls = list(iter_xmls(uploaded_files))
    out = io.BytesIO()
    errors = []
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, (name, raw) in enumerate(xmls, start=1):
            try:
                rows, _ = parse_xml(raw, name, representation, exclude_technical)
                zf.writestr(f"{Path(name).stem}_convertido.xlsx", rows_to_excel(rows, "CFDI"))
            except Exception as exc:
                errors.append({"Archivo": name, "Error": str(exc)})
            progress.progress(idx / len(xmls))
            status.caption(f"Procesando {idx:,} de {len(xmls):,}: {name}")
        if errors:
            zf.writestr("reporte_errores.csv", pd.DataFrame(errors).to_csv(index=False).encode("utf-8-sig"))
    out.seek(0)
    return out.getvalue(), errors


for key, value in {
    "download_data": None,
    "download_name": None,
    "download_mime": None,
    "summary": None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = value


with st.sidebar:
    st.header("⚙️ Opciones avanzadas")

    exclude_technical = st.checkbox(
        "Excluir sellos y certificados largos",
        value=True,
        help=(
            "Excluye Sello, Certificado, SelloCFD y SelloSAT para reducir "
            "el tamaño del archivo. El resto de los datos se conserva."
        ),
    )

    sample_mode = st.radio(
        "Muestra del preview",
        ["Primeros archivos", "Aleatorios"],
        index=0,
        help="Con cargas grandes, el preview utiliza hasta 20 XML.",
    )

    st.markdown("---")
    st.caption(
        "Las columnas se generan dinámicamente desde cada XML. "
        "No se hardcodean proveedores, UUID ni campos concretos."
    )


# ============================================================
# 1. CARGA
# ============================================================

st.markdown("## 1. Carga tus archivos")
st.caption(
    "Puedes cargar un XML, varios XML, un ZIP con XML o varios ZIP."
)

uploaded_files = st.file_uploader(
    "Arrastra o selecciona XML y ZIP",
    type=["xml", "zip"],
    accept_multiple_files=True,
    help="Admite uno o múltiples archivos XML y ZIP.",
)

if not uploaded_files:
    st.info(
        "Carga al menos un XML o ZIP para mostrar las opciones de estructura, "
        "el modo de salida y los previews."
    )
    st.stop()

try:
    analysis = analyze(uploaded_files)
except Exception as exc:
    st.error(f"No se pudo analizar la carga: {exc}")
    st.stop()

if analysis["total"] == 0:
    st.warning("No se encontraron archivos XML dentro de la carga.")
    st.stop()


input_signature = tuple(
    sorted(
        (
            uploaded.name,
            int(getattr(uploaded, "size", len(uploaded.getvalue()))),
        )
        for uploaded in uploaded_files
    )
)

if st.session_state.get("input_signature") != input_signature:
    st.session_state.input_signature = input_signature
    st.session_state.download_data = None
    st.session_state.download_name = None
    st.session_state.download_mime = None
    st.session_state.summary = None


# ============================================================
# 2. ANÁLISIS
# ============================================================

st.markdown("## 2. Análisis automático")

c1, c2, c3, c4 = st.columns(4)
c1.metric("XML detectados", f"{analysis['total']:,}")
c2.metric("XML válidos", f"{analysis['valid']:,}")
c3.metric("Conceptos detectados", f"{analysis['concepts']:,}")
c4.metric("Tipos de estructura", f"{len(analysis['types']):,}")

if analysis["types"]:
    st.dataframe(
        pd.DataFrame(
            [
                {"Tipo detectado": name, "Cantidad": quantity}
                for name, quantity in analysis["types"].items()
            ]
        ),
        use_container_width=True,
        hide_index=True,
    )

if analysis["errors"]:
    with st.expander(
        f"⚠️ Ver {len(analysis['errors']):,} XML con problemas"
    ):
        st.dataframe(
            pd.DataFrame(analysis["errors"]),
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# 3. ESTRUCTURA
# ============================================================

st.markdown("## 3. Elige cómo se organizarán las filas")

representation_label = st.radio(
    "Estructura del Excel",
    [
        "Una fila por concepto",
        "Una fila por comprobante",
    ],
    index=0,
    horizontal=True,
    key="representation_main",
)

representation = (
    "por_concepto"
    if representation_label == "Una fila por concepto"
    else "por_comprobante"
)

description_col1, description_col2 = st.columns(2)

with description_col1:
    with st.container(border=True):
        st.markdown("### 🧾 Una fila por concepto")
        st.write(
            "Cada producto o servicio del CFDI genera una fila. "
            "Los datos generales se repiten para cada concepto."
        )
        st.caption(
            "Recomendado para Power BI, compras, productos, servicios "
            "e impuestos detallados."
        )

with description_col2:
    with st.container(border=True):
        st.markdown("### 📄 Una fila por comprobante")
        st.write(
            "Cada XML genera una sola fila. Los conceptos se colocan "
            "en grupos de columnas numeradas."
        )
        st.caption(
            "Recomendado para revisar facturas, folios, emisores, "
            "receptores y totales."
        )

st.success(
    f"Se exportará como **{representation_label.lower()}**."
)


# ============================================================
# 4. SALIDA: UNO O VARIOS XML
# ============================================================

st.markdown("## 4. Elige el tipo de descarga")

if analysis["total"] == 1:
    output_mode = "individual"

    st.info(
        "Se detectó **un solo XML**. La app generará directamente un Excel "
        "individual y no mostrará la opción de unir archivos."
    )

    with st.container(border=True):
        st.markdown("### Resultado esperado")
        st.code(
            "nombre_del_xml_convertido.xlsx\n└── Hoja: CFDI",
            language="text",
        )

else:
    output_mode_label = st.radio(
        "Se detectaron varios XML. ¿Cómo deseas descargarlos?",
        [
            "Unir todos en un solo Excel",
            "Dejar un Excel por cada XML",
        ],
        index=0,
        horizontal=True,
        key="output_mode_main",
    )

    if output_mode_label == "Unir todos en un solo Excel":
        output_mode = "consolidated"
        st.success(
            "Todos los XML se integrarán en una sola tabla. La columna "
            "`Archivo_XML` permitirá identificar el documento de origen."
        )
        st.code(
            "CFDI_Consolidado.xlsx\n└── Hoja: CFDI_Consolidado",
            language="text",
        )
    else:
        output_mode = "individual"
        st.info(
            "Se generará un ZIP con un Excel independiente por cada XML."
        )
        st.code(
            "CFDI_Individuales.zip\n"
            "├── factura_001_convertido.xlsx\n"
            "├── factura_002_convertido.xlsx\n"
            "└── reporte_errores.csv (si aplica)",
            language="text",
        )


# ============================================================
# 5. PREVIEW COMPARATIVO
# ============================================================

st.markdown("## 5. Compara las vistas previas")
st.caption(
    "Puedes revisar ambas estructuras antes de procesar todos los archivos. "
    "La pestaña marcada corresponde a la opción elegida para exportar."
)

with st.spinner("Generando previews..."):
    preview_concept_df, preview_concept_errors, concept_sample_count = make_preview(
        uploaded_files,
        "por_concepto",
        exclude_technical,
        sample_mode,
    )

    preview_document_df, preview_document_errors, document_sample_count = make_preview(
        uploaded_files,
        "por_comprobante",
        exclude_technical,
        sample_mode,
    )


def show_preview(
    dataframe,
    sample_count,
    errors,
    key_prefix,
):
    p1, p2, p3 = st.columns(3)
    p1.metric("XML usados en muestra", f"{sample_count:,}")
    p2.metric("Filas generadas", f"{len(dataframe):,}")
    p3.metric("Columnas detectadas", f"{len(dataframe.columns):,}")

    if dataframe.empty:
        st.warning("No se generaron filas para esta vista previa.")
    else:
        defaults = list(
            dataframe.columns[: min(14, len(dataframe.columns))]
        )

        visible = st.multiselect(
            "Columnas visibles",
            options=list(dataframe.columns),
            default=defaults,
            key=f"{key_prefix}_columns",
            help=(
                "Solo cambia la visualización. El Excel conservará "
                "todas las columnas detectadas."
            ),
        )

        st.dataframe(
            dataframe[visible or defaults].head(PREVIEW_ROW_LIMIT),
            use_container_width=True,
            hide_index=True,
            height=430,
        )

        st.caption(
            f"Se muestran hasta {PREVIEW_ROW_LIMIT} filas. "
            "La exportación incluirá todos los registros."
        )

    if errors:
        with st.expander(
            f"⚠️ Errores de esta muestra: {len(errors):,}"
        ):
            st.dataframe(
                pd.DataFrame(errors),
                use_container_width=True,
                hide_index=True,
            )


concept_tab, document_tab = st.tabs(
    [
        "🧾 Preview por concepto",
        "📄 Preview por comprobante",
    ]
)

with concept_tab:
    if representation == "por_concepto":
        st.success(
            "✅ Esta es la estructura seleccionada para la exportación."
        )
    else:
        st.caption(
            "Vista de comparación; actualmente no está seleccionada."
        )

    show_preview(
        preview_concept_df,
        concept_sample_count,
        preview_concept_errors,
        "concept",
    )

with document_tab:
    if representation == "por_comprobante":
        st.success(
            "✅ Esta es la estructura seleccionada para la exportación."
        )
    else:
        st.caption(
            "Vista de comparación; actualmente no está seleccionada."
        )

    show_preview(
        preview_document_df,
        document_sample_count,
        preview_document_errors,
        "document",
    )


# ============================================================
# 6. PROCESAR
# ============================================================

st.markdown("## 6. Procesar y descargar")

s1, s2, s3 = st.columns(3)
s1.metric("XML por procesar", f"{analysis['total']:,}")
s2.metric("Estructura", representation_label)
s3.metric(
    "Salida",
    (
        "Excel individual"
        if analysis["total"] == 1
        else (
            "Excel consolidado"
            if output_mode == "consolidated"
            else "ZIP con Excel individuales"
        )
    ),
)

button_label = (
    "Convertir este XML a Excel"
    if analysis["total"] == 1
    else f"Procesar {analysis['total']:,} XML"
)

if st.button(
    button_label,
    type="primary",
    use_container_width=True,
):
    st.session_state.download_data = None
    st.session_state.download_name = None
    st.session_state.download_mime = None
    st.session_state.summary = None

    progress = st.progress(0)
    status = st.empty()
    started = time.time()

    try:
        if analysis["total"] == 1 or output_mode == "consolidated":
            rows, errors = process_consolidated(
                uploaded_files,
                representation,
                exclude_technical,
                progress,
                status,
            )

            st.session_state.download_data = rows_to_excel(
                rows,
                "CFDI" if analysis["total"] == 1 else "CFDI_Consolidado",
                errors,
            )

            if analysis["total"] == 1:
                first_name = next(iter_xmls(uploaded_files))[0]
                st.session_state.download_name = (
                    f"{Path(first_name).stem}_convertido.xlsx"
                )
            else:
                st.session_state.download_name = "CFDI_Consolidado.xlsx"

            st.session_state.download_mime = (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )

            st.session_state.summary = {
                "procesados": analysis["total"],
                "correctos": analysis["total"] - len(errors),
                "errores": len(errors),
                "filas": len(rows),
                "tiempo": time.time() - started,
                "estructura": representation_label,
                "salida": (
                    "Excel individual"
                    if analysis["total"] == 1
                    else "Excel consolidado"
                ),
            }

        else:
            data, errors = process_individual(
                uploaded_files,
                representation,
                exclude_technical,
                progress,
                status,
            )

            st.session_state.download_data = data
            st.session_state.download_name = "CFDI_Individuales.zip"
            st.session_state.download_mime = "application/zip"

            st.session_state.summary = {
                "procesados": analysis["total"],
                "correctos": analysis["total"] - len(errors),
                "errores": len(errors),
                "filas": None,
                "tiempo": time.time() - started,
                "estructura": representation_label,
                "salida": "ZIP con Excel individuales",
            }

        status.success("Procesamiento completado.")

    except Exception as exc:
        st.error(
            f"No fue posible completar el procesamiento: {exc}"
        )


# ============================================================
# 7. RESULTADO
# ============================================================

if st.session_state.summary:
    summary = st.session_state.summary

    st.success("Conversión completada correctamente.")

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("XML procesados", f"{summary['procesados']:,}")
    r2.metric("Correctos", f"{summary['correctos']:,}")
    r3.metric("Con error", f"{summary['errores']:,}")
    r4.metric("Tiempo", f"{summary['tiempo']:.2f} s")

    st.caption(
        f"Estructura: {summary['estructura']} · "
        f"Salida: {summary['salida']}"
    )

if st.session_state.download_data:
    st.download_button(
        label=f"⬇️ Descargar {st.session_state.download_name}",
        data=st.session_state.download_data,
        file_name=st.session_state.download_name,
        mime=st.session_state.download_mime,
        use_container_width=True,
    )

    st.markdown(
        '<div class="signature">'
        'Aquí debe descargarlo ella 👆<br>'
        'Hecho por tu novio el ingeniero xd 😎'
        '</div>',
        unsafe_allow_html=True,
    )
