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
    st.header("⚙️ Configuración")
    representation_label = st.radio("Estructura de salida", ["Una fila por concepto", "Una fila por comprobante"])
    representation = "por_concepto" if representation_label == "Una fila por concepto" else "por_comprobante"
    exclude_technical = st.checkbox("Excluir sellos y certificados largos", value=True)
    sample_mode = st.radio("Muestra del preview", ["Primeros archivos", "Aleatorios"])
    st.markdown("---")
    st.caption("Las columnas se detectan dinámicamente. No se hardcodean proveedores ni campos concretos.")

st.markdown("## 1. Carga tus archivos")
uploaded_files = st.file_uploader("Selecciona XML o ZIP", type=["xml", "zip"], accept_multiple_files=True)
if not uploaded_files:
    st.info("Carga al menos un XML o ZIP para comenzar.")
    st.stop()

try:
    analysis = analyze(uploaded_files)
except Exception as exc:
    st.error(str(exc))
    st.stop()

if analysis["total"] == 0:
    st.warning("No se encontraron XML en la carga.")
    st.stop()

st.markdown("## 2. Análisis automático")
c1, c2, c3, c4 = st.columns(4)
c1.metric("XML detectados", f"{analysis['total']:,}")
c2.metric("XML válidos", f"{analysis['valid']:,}")
c3.metric("Conceptos", f"{analysis['concepts']:,}")
c4.metric("Estructuras", f"{len(analysis['types']):,}")
if analysis["types"]:
    st.dataframe(pd.DataFrame([{"Tipo detectado": k, "Cantidad": v} for k, v in analysis["types"].items()]), use_container_width=True, hide_index=True)
if analysis["errors"]:
    with st.expander(f"⚠️ Ver {len(analysis['errors'])} errores"):
        st.dataframe(pd.DataFrame(analysis["errors"]), use_container_width=True, hide_index=True)

st.markdown("## 3. Configura la salida")
if analysis["total"] == 1:
    output_mode = "Consolidado"
    st.info("Se detectó un solo XML. Se generará un Excel individual; no se muestra la opción de unir archivos.")
else:
    output_mode = st.radio("¿Cómo deseas generar el resultado?", ["Consolidar todos en un solo Excel", "Generar un Excel por cada XML"], horizontal=True)

st.markdown("## 4. Vista previa")
preview_df, preview_errors, sample_count = make_preview(uploaded_files, representation, exclude_technical, sample_mode)
p1, p2, p3 = st.columns(3)
p1.metric("XML usados en muestra", sample_count)
p2.metric("Filas de muestra", len(preview_df))
p3.metric("Columnas detectadas", len(preview_df.columns))
if not preview_df.empty:
    defaults = list(preview_df.columns[: min(12, len(preview_df.columns))])
    visible = st.multiselect("Columnas visibles en el preview", list(preview_df.columns), default=defaults)
    st.dataframe(preview_df[visible or defaults].head(PREVIEW_ROW_LIMIT), use_container_width=True, hide_index=True, height=420)
    st.caption(f"Vista previa de hasta {PREVIEW_ROW_LIMIT} filas. El archivo final incluirá todas las filas y columnas.")
if preview_errors:
    with st.expander(f"⚠️ Errores en preview: {len(preview_errors)}"):
        st.dataframe(pd.DataFrame(preview_errors), use_container_width=True, hide_index=True)

st.markdown("## 5. Procesar y descargar")
button_label = "Convertir XML a Excel" if analysis["total"] == 1 else f"Procesar {analysis['total']:,} CFDI"
if st.button(button_label, type="primary", use_container_width=True):
    progress = st.progress(0)
    status = st.empty()
    started = time.time()
    try:
        if analysis["total"] == 1 or output_mode == "Consolidar todos en un solo Excel":
            rows, errors = process_consolidated(uploaded_files, representation, exclude_technical, progress, status)
            st.session_state.download_data = rows_to_excel(rows, "CFDI" if analysis["total"] == 1 else "CFDI_Consolidado", errors)
            st.session_state.download_name = "CFDI_Consolidado.xlsx" if analysis["total"] > 1 else f"{Path(next(iter_xmls(uploaded_files))[0]).stem}_convertido.xlsx"
            st.session_state.download_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            st.session_state.summary = {"procesados": analysis["total"], "correctos": analysis["total"] - len(errors), "errores": len(errors), "filas": len(rows), "tiempo": time.time() - started}
        else:
            data, errors = process_individual(uploaded_files, representation, exclude_technical, progress, status)
            st.session_state.download_data = data
            st.session_state.download_name = "CFDI_Individuales.zip"
            st.session_state.download_mime = "application/zip"
            st.session_state.summary = {"procesados": analysis["total"], "correctos": analysis["total"] - len(errors), "errores": len(errors), "filas": None, "tiempo": time.time() - started}
        status.success("Procesamiento completado.")
    except Exception as exc:
        st.error(f"No fue posible completar el procesamiento: {exc}")

if st.session_state.summary:
    s = st.session_state.summary
    st.success("Conversión completada correctamente.")
    r1, r2, r3, r4 = st.columns(4)
    r1.metric("XML procesados", f"{s['procesados']:,}")
    r2.metric("Correctos", f"{s['correctos']:,}")
    r3.metric("Con error", f"{s['errores']:,}")
    r4.metric("Tiempo", f"{s['tiempo']:.2f} s")

if st.session_state.download_data:
    st.download_button(f"⬇️ Descargar {st.session_state.download_name}", st.session_state.download_data, st.session_state.download_name, st.session_state.download_mime, use_container_width=True)
    st.markdown('<div class="signature">Aquí debe descargarlo ella 👆<br>Hecho por tu novio el ingeniero xd 😎</div>', unsafe_allow_html=True)
