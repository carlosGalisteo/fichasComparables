"""
Comparables Inmobiliarios — Generador de Fichas de Testigos de Mercado
Tenerife · Powered by Claude Vision API
"""

import streamlit as st
import anthropic
import base64
import json
import io
import os
import shutil
import tempfile
from datetime import date
from pathlib import Path
import openpyxl

from fill_template import fill_comparable

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Fichas Comparables · Tenerife",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CSS PERSONALIZADO
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

  html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
  }

  /* Sidebar */
  [data-testid="stSidebar"] {
    background: #0f1117;
    border-right: 1px solid #1e2130;
  }
  [data-testid="stSidebar"] * {
    color: #e0e4f0 !important;
  }
  [data-testid="stSidebar"] .stTextInput input,
  [data-testid="stSidebar"] .stDateInput input,
  [data-testid="stSidebar"] .stTextArea textarea {
    background: #1a1d2e !important;
    border: 1px solid #2d3150 !important;
    color: #e0e4f0 !important;
    border-radius: 6px !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 12px !important;
  }
  [data-testid="stSidebar"] .stTextArea textarea::placeholder {
    color: #4b5563 !important;
  }

  /* Header principal */
  .app-header {
    background: linear-gradient(135deg, #0f1117 0%, #1a1d2e 100%);
    border: 1px solid #2d3150;
    border-radius: 12px;
    padding: 24px 32px;
    margin-bottom: 24px;
    display: flex;
    align-items: center;
    gap: 16px;
  }
  .app-header h1 {
    color: #e0e4f0;
    font-size: 22px;
    font-weight: 600;
    margin: 0;
  }
  .app-header p {
    color: #6b7280;
    font-size: 13px;
    margin: 4px 0 0 0;
  }

  /* Cards de comparable */
  .comparable-card {
    background: #fafafa;
    border: 1px solid #e5e7eb;
    border-radius: 10px;
    padding: 16px 20px;
    margin-bottom: 12px;
    transition: border-color 0.2s;
  }
  .comparable-card:hover { border-color: #3b5bdb; }
  .comparable-card.processing { border-color: #f59e0b; background: #fffbeb; }
  .comparable-card.done { border-color: #10b981; background: #f0fdf4; }
  .comparable-card.error { border-color: #ef4444; background: #fef2f2; }

  /* Badges de campo desconocido */
  .badge-unknown {
    display: inline-block;
    background: #fee2e2;
    color: #991b1b;
    font-size: 10px;
    font-weight: 600;
    padding: 2px 6px;
    border-radius: 4px;
    margin: 2px;
    font-family: 'DM Mono', monospace;
  }
  .badge-ok {
    display: inline-block;
    background: #d1fae5;
    color: #065f46;
    font-size: 10px;
    font-weight: 600;
    padding: 2px 6px;
    border-radius: 4px;
    margin: 2px;
    font-family: 'DM Mono', monospace;
  }

  /* Tabla de datos extraídos */
  .data-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 8px;
    margin-top: 12px;
  }
  .data-cell {
    background: white;
    border: 1px solid #e5e7eb;
    border-radius: 6px;
    padding: 8px 10px;
  }
  .data-cell .label {
    font-size: 9px;
    color: #9ca3af;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    font-weight: 600;
    margin-bottom: 2px;
  }
  .data-cell .value {
    font-size: 13px;
    color: #111827;
    font-weight: 500;
    font-family: 'DM Mono', monospace;
  }
  .data-cell .value.unknown { color: #dc2626; }

  /* Botón primario */
  .stButton > button[kind="primary"] {
    background: #3b5bdb !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    letter-spacing: 0.02em !important;
    transition: all 0.2s !important;
  }
  .stButton > button[kind="primary"]:hover {
    background: #2f4cc0 !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 12px rgba(59, 91, 219, 0.3) !important;
  }

  /* Notas del analista */
  .analyst-notes {
    background: #f8f9fa;
    border-left: 3px solid #f59e0b;
    padding: 10px 14px;
    border-radius: 0 6px 6px 0;
    font-size: 12px;
    color: #6b7280;
    margin-top: 10px;
    font-style: italic;
  }

  /* Download button */
  .stDownloadButton > button {
    background: #059669 !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
  }
  .stDownloadButton > button:hover {
    background: #047857 !important;
  }

  /* Ocultar hamburger y footer */
  #MainMenu { visibility: hidden; }
  footer { visibility: hidden; }
  header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CARGA DEL SYSTEM PROMPT
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data
def load_system_prompt():
    sp_path = Path(__file__).parent / "SYSTEM_PROMPT_COMPARABLES.txt"
    if sp_path.exists():
        return sp_path.read_text(encoding="utf-8")
    return ""

SYSTEM_PROMPT = load_system_prompt()

# ─────────────────────────────────────────────────────────────────────────────
# FUNCIONES CORE
# ─────────────────────────────────────────────────────────────────────────────

def image_to_base64(uploaded_file) -> tuple[str, str]:
    """Convierte un archivo subido a base64 + media_type."""
    data = uploaded_file.read()
    b64 = base64.standard_b64encode(data).decode("utf-8")
    mt = uploaded_file.type or "image/png"
    return b64, mt


def bytes_to_b64(data: bytes) -> str:
    return base64.standard_b64encode(data).decode("utf-8")


def crop_zones(img_bytes: bytes) -> dict[str, tuple[bytes, str]]:
    """
    Bloque 0: recorta la imagen en zonas estándar usando PIL.
    Devuelve dict {nombre: (bytes_png, media_type)}.
    """
    from PIL import Image as PILImage
    import io as _io

    img = PILImage.open(_io.BytesIO(img_bytes))
    w, h = img.size

    zonas = {
        "ZONA_FICHA":     img.crop((0, int(h*0.03), w, int(h*0.38))),
        "ZONA_PRECIO":    img.crop((0, int(h*0.55), w, int(h*0.75))),
        "ZONA_CABECERA":  img.crop((0, 0,            w, int(h*0.06))),
        "ZONA_UBICACION": img.crop((0, int(h*0.80),  w, int(h*0.92))),
        "ZONA_ANUNCIANTE":img.crop((0, int(h*0.10),  w, int(h*0.22))),
    }

    result = {}
    for nombre, zona in zonas.items():
        buf = _io.BytesIO()
        zona.save(buf, format="PNG")
        result[nombre] = (buf.getvalue(), "image/png")
    return result


def extract_comparable_from_image(
    client: anthropic.Anthropic,
    img_b64: str,
    img_media_type: str,
    num: int,
    url: str,
    fecha: str,
    ref_catastral: str = "",
) -> dict:
    """
    Llama a Claude con visión aplicando el preprocesado del Bloque 0:
    recorta la imagen en zonas y las envía junto con la imagen completa
    para maximizar la precisión de lectura.
    """
    # Decodificar bytes originales para recortar
    img_bytes = base64.b64decode(img_b64)
    zonas = crop_zones(img_bytes)

    # Construir el mensaje multi-imagen: zonas + imagen completa
    content_parts = []

    zona_labels = {
        "ZONA_CABECERA":   "ZONA_CABECERA — dirección, título, precio destacado",
        "ZONA_FICHA":      "ZONA_FICHA — características: sup. construida, dormitorios, baños, dotaciones, año construcción, comercializador",
        "ZONA_PRECIO":     "ZONA_PRECIO — precio exacto, precio/m², anejos incluidos",
        "ZONA_UBICACION":  "ZONA_UBICACION — municipio, barrio, zona, CP",
        "ZONA_ANUNCIANTE": "ZONA_ANUNCIANTE — nombre de la agencia / inmobiliaria anunciante",
    }

    for nombre, (zona_bytes, zona_mt) in zonas.items():
        content_parts.append({
            "type": "text",
            "text": f"--- {zona_labels.get(nombre, nombre)} ---"
        })
        content_parts.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": zona_mt,
                "data": bytes_to_b64(zona_bytes),
            }
        })

    # Imagen completa al final como referencia
    content_parts.append({
        "type": "text",
        "text": "--- IMAGEN COMPLETA (referencia) ---"
    })
    content_parts.append({
        "type": "image",
        "source": {
            "type": "base64",
            "media_type": img_media_type,
            "data": img_b64,
        }
    })

    user_msg = f"""Se te envían 5 recortes de zonas específicas del anuncio (ZONA_CABECERA, ZONA_FICHA,
ZONA_PRECIO, ZONA_UBICACION, ZONA_ANUNCIANTE) seguidos de la imagen completa como referencia.

Lee PRIMERO cada zona recortada con atención — el texto es más legible en los recortes.
Usa la imagen completa solo para confirmar datos dudosos.

DATOS APORTADOS POR EL USUARIO:
- Número de comparable: {num}
- URL del anuncio: {url if url else "No aportada"}
- Fecha de aportación: {fecha}
- Referencia catastral: {ref_catastral if ref_catastral else ""}

INSTRUCCIONES CRÍTICAS:
1. El COMERCIALIZADOR es el nombre de la inmobiliaria o agencia anunciante visible en
   ZONA_ANUNCIANTE o ZONA_FICHA. NUNCA es el nombre del portal (Idealista, Fotocasa, etc.).
   Si aparece "RMA", "Engel & Völkers", "Re/Max" u otro nombre de agencia → úsalo.
2. Lee el precio con máxima atención en ZONA_PRECIO.
3. Lee superficies y dormitorios en ZONA_FICHA.
4. Si algún valor es ambiguo → DESCONOCIDO.

Devuelve ÚNICAMENTE el objeto JSON válido, sin texto adicional, sin marcadores de código."""

    content_parts.append({"type": "text", "text": user_msg})

    response = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=2000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": content_parts}],
    )

    raw = response.content[0].text.strip()
    # Limpiar backticks residuales
    raw = raw.replace("```json", "").replace("```", "").strip()
    # Extraer el primer objeto JSON válido aunque haya texto alrededor
    start = raw.find("{")
    end   = raw.rfind("}") + 1
    if start == -1 or end == 0:
        raise json.JSONDecodeError("No se encontró objeto JSON en la respuesta", raw, 0)
    return json.loads(raw[start:end])


def build_xlsx_multisheet(comparables_data: list[dict]) -> bytes:
    """
    Genera un único .xlsx con una hoja por comparable.
    Usa fill_comparable() de fill_template.py sobre la plantilla base.
    """
    template_path = Path(__file__).parent / "COMPARABLE_FICHA_v4.xlsx"

    with tempfile.TemporaryDirectory() as tmpdir:
        # Generar un xlsx por comparable
        xlsx_paths = []
        for item in comparables_data:
            num    = item["comparable"]
            campos = item["campos"]
            flags  = item.get("flags", {})
            out_p  = Path(tmpdir) / f"comp_{num}.xlsx"
            fill_comparable(str(template_path), str(out_p), num, campos, flags)
            xlsx_paths.append(out_p)

        if len(xlsx_paths) == 1:
            return xlsx_paths[0].read_bytes()

        # Combinar todas las hojas en un único workbook
        wb_final = openpyxl.Workbook()
        wb_final.remove(wb_final.active)  # quitar hoja vacía inicial

        for p in xlsx_paths:
            wb_src = openpyxl.load_workbook(p)
            for shname in wb_src.sheetnames:
                ws_src = wb_src[shname]
                ws_dst = wb_final.create_sheet(title=shname)
                for row in ws_src.iter_rows():
                    for cell in row:
                        new_cell = ws_dst.cell(
                            row=cell.row, column=cell.column, value=cell.value
                        )
                        if cell.has_style:
                            new_cell.font      = cell.font.copy()
                            new_cell.fill      = cell.fill.copy()
                            new_cell.border    = cell.border.copy()
                            new_cell.alignment = cell.alignment.copy()
                            new_cell.number_format = cell.number_format
                # Copiar anchos de columna y altos de fila
                for col_dim in ws_src.column_dimensions.values():
                    ws_dst.column_dimensions[col_dim.index].width = col_dim.width
                for row_dim in ws_src.row_dimensions.values():
                    ws_dst.row_dimensions[row_dim.index].height = row_dim.height

        buf = io.BytesIO()
        wb_final.save(buf)
        return buf.getvalue()


def render_extracted_data(item: dict):
    """Muestra un resumen visual de los datos extraídos de un comparable."""
    campos = item.get("campos", {})
    flags  = item.get("flags", {})
    desc   = flags.get("campos_desconocidos", [])
    notas  = flags.get("notas", "")

    def v(key, default="—"):
        val = campos.get(key)
        if val is None or val == "":
            return default
        return str(val)

    def cls(key):
        return "unknown" if key in desc or v(key) == "DESCONOCIDO" else ""

    precio_raw = campos.get("B9")
    if isinstance(precio_raw, (int, float)):
        precio_fmt = f"{precio_raw:,.0f} €"
    else:
        precio_fmt = str(precio_raw) if precio_raw else "—"

    html = f"""
    <div class="data-grid">
      <div class="data-cell">
        <div class="label">Dirección</div>
        <div class="value {cls('B2')}">{v('B2')} {v('B3', '')}</div>
      </div>
      <div class="data-cell">
        <div class="label">Municipio</div>
        <div class="value {cls('F2')}">{v('F2')}</div>
      </div>
      <div class="data-cell">
        <div class="label">C.P.</div>
        <div class="value {cls('F3')}">{v('F3')}</div>
      </div>
      <div class="data-cell">
        <div class="label">Tipo</div>
        <div class="value {cls('B5')}">{v('B5')}</div>
      </div>
      <div class="data-cell">
        <div class="label">Planta</div>
        <div class="value {cls('D3')}">{v('D3')}</div>
      </div>
      <div class="data-cell">
        <div class="label">Dormitorios</div>
        <div class="value {cls('D5')}">{v('D5')}</div>
      </div>
      <div class="data-cell">
        <div class="label">Baños</div>
        <div class="value {cls('F5')}">{v('F5')}</div>
      </div>
      <div class="data-cell">
        <div class="label">Sup. construida</div>
        <div class="value {cls('B6')}">{v('B6')} {'m²' if v('B6') not in ('—','DESCONOCIDO') else ''}</div>
      </div>
      <div class="data-cell">
        <div class="label">Antigüedad</div>
        <div class="value {cls('D6')}">{v('D6')} {'años' if v('D6') not in ('—','DESCONOCIDO') else ''}</div>
      </div>
      <div class="data-cell">
        <div class="label">Precio oferta</div>
        <div class="value {cls('B9')}">{precio_fmt}</div>
      </div>
      <div class="data-cell">
        <div class="label">Fuente</div>
        <div class="value">{v('B11')}</div>
      </div>
      <div class="data-cell">
        <div class="label">Comercializador</div>
        <div class="value {cls('D11')}">{v('D11')}</div>
      </div>
    </div>
    """

    if desc:
        badges = "".join(f'<span class="badge-unknown">{c}</span>' for c in desc)
        html += f'<div style="margin-top:10px;"><strong style="font-size:11px;color:#991b1b;">Campos desconocidos:</strong> {badges}</div>'

    if flags.get("precio_con_anejos"):
        html += '<div style="margin-top:6px;"><span class="badge-unknown">⚠ Precio incluye anejos — revisar D9</span></div>'

    if notas:
        html += f'<div class="analyst-notes">💬 {notas}</div>'

    st.markdown(html, unsafe_allow_html=True)

    # Selector de estado de conservación (campo H6)
    estados = ["", "Muy bueno", "Bueno", "Medio", "Malo", "Muy malo"]
    estado_key = f"estado_{item.get('comparable', id(item))}"
    estado_actual = item.get("campos", {}).get("H6", "")
    idx_actual = estados.index(estado_actual) if estado_actual in estados else 0
    nuevo_estado = st.selectbox(
        "Estado de conservación",
        options=estados,
        index=idx_actual,
        key=estado_key,
        help="Campo H6 de la ficha — selecciona el estado del inmueble",
    )
    item["campos"]["H6"] = nuevo_estado


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS DE ESTADO
# ─────────────────────────────────────────────────────────────────────────────

def _renumber_inputs():
    """Reordena el campo num tras eliminar un comparable del lote."""
    for i, item in enumerate(st.session_state.inputs):
        item["num"] = i + 1


# ─────────────────────────────────────────────────────────────────────────────
# ESTADO DE SESIÓN
# ─────────────────────────────────────────────────────────────────────────────

if "inputs" not in st.session_state:
    st.session_state.inputs = []       # datos introducidos por el usuario

if "resultados" not in st.session_state:
    st.session_state.resultados = []   # resultados extraídos por Claude

if "fase" not in st.session_state:
    st.session_state.fase = "entrada"  # entrada | revision | analisis | resultados

if "form_key" not in st.session_state:
    st.session_state.form_key = 0      # incrementar al guardar para resetear el formulario

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("### ⚙️ Configuración")
    st.markdown("---")

    # Leer API key desde Streamlit Secrets (despliegue en cloud)
    # o permitir entrada manual como fallback
    try:
        _secret_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:
        _secret_key = ""
    if _secret_key:
        api_key = _secret_key
        st.success("🔑 API Key configurada", icon="✅")
    else:
        api_key = st.text_input(
            "API Key de Anthropic",
            type="password",
            placeholder="sk-ant-...",
            help="Necesaria para analizar imágenes con Claude. No se almacena.",
        )

    st.markdown("---")
    st.markdown("### 📋 Datos del lote")

    fecha_aportacion = st.date_input(
        "Fecha de aportación",
        value=date.today(),
        format="DD/MM/YYYY",
        help="Fecha en que se aportan las imágenes (campo H3 de la ficha)",
    )
    fecha_str = fecha_aportacion.strftime("%d/%m/%Y")

    # ── Formulario de entrada comparable a comparable ──────────────────────────
    st.markdown("---")
    _next_num = len(st.session_state.inputs) + 1
    st.markdown(f"### Comparable {_next_num}")

    _uploaded = st.file_uploader(
        "Imagen del anuncio",
        type=["png", "jpg", "jpeg", "webp"],
        key=f"img_upload_{st.session_state.form_key}",
    )
    _url_input = st.text_input(
        "URL del anuncio",
        placeholder="https://www.idealista.com/inmueble/...",
        key=f"url_{st.session_state.form_key}",
    )
    _ref_input = st.text_input(
        "Ref. catastral (opcional)",
        placeholder="0000001AA0000S0000AA",
        key=f"ref_{st.session_state.form_key}",
    )
    _estado_input = st.selectbox(
        "Estado de conservación",
        options=["Muy bueno", "Bueno", "Medio", "Malo", "Muy malo"],
        index=1,
        key=f"estado_{st.session_state.form_key}",
    )

    if st.button(
        "Guardar comparable",
        use_container_width=True,
        disabled=(_uploaded is None),
        help="Sube una imagen para poder guardar el comparable.",
    ):
        st.session_state.inputs.append({
            "num": _next_num,
            "image_bytes": _uploaded.read(),
            "image_name": _uploaded.name,
            "image_type": _uploaded.type or "image/png",
            "url": _url_input.strip(),
            "ref": _ref_input.strip(),
            "estado": _estado_input,
        })
        st.session_state.form_key += 1
        st.rerun()

    # ── Resumen del lote en el sidebar ─────────────────────────────────────────
    if st.session_state.inputs:
        st.markdown("---")
        st.markdown(f"### Lote ({len(st.session_state.inputs)} comparable(s))")

        for _item in st.session_state.inputs:
            _col1, _col2 = st.columns([4, 1])
            with _col1:
                st.markdown(f"**{_item['num']}.** {_item['image_name']}")
                _url_ok = "✓" if _item["url"] else "—"
                st.caption(f"Estado: {_item['estado']} · URL: {_url_ok}")
            with _col2:
                if st.button("✕", key=f"del_{_item['num']}"):
                    st.session_state.inputs = [
                        x for x in st.session_state.inputs if x["num"] != _item["num"]
                    ]
                    _renumber_inputs()
                    st.rerun()

        st.markdown("")
        if st.button("🗑 Limpiar todo", use_container_width=True):
            st.session_state.inputs = []
            st.session_state.resultados = []
            st.session_state.form_key += 1
            st.rerun()

    st.markdown("---")
    st.markdown(
        "<div style='font-size:11px;color:#4b5563;text-align:center;'>"
        "Fichas Comparables · Tenerife<br>"
        "Powered by Claude Vision<br>"
        "<span style='color:#6366f1;'>claude-opus-4-5</span></div>",
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
# CABECERA PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

# Logo en base64 para incrustar en el header
import base64 as _b64
from pathlib import Path as _Path
_logo_path = _Path(__file__).parent / "icono.png"
if _logo_path.exists():
    _logo_b64 = _b64.b64encode(_logo_path.read_bytes()).decode()
    _logo_html = f'<img src="data:image/png;base64,{_logo_b64}" style="height:52px;width:auto;">'
else:
    _logo_html = '<div style="font-size:36px;">🏠</div>'

st.markdown(f"""
<div class="app-header">
  {_logo_html}
  <div>
    <h1>Fichas de Testigos de Mercado</h1>
    <p>Extracción automática de comparables inmobiliarios · Tenerife · Claude Vision</p>
  </div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# RESUMEN DEL LOTE Y BOTÓN DE ANÁLISIS
# ─────────────────────────────────────────────────────────────────────────────

if st.session_state.inputs:
    st.markdown("---")
    _n = len(st.session_state.inputs)
    st.markdown(f"### Lote cargado — {_n} comparable(s)")

    for _item in st.session_state.inputs:
        _url_label = _item["url"] or "—"
        if len(_url_label) > 60:
            _url_label = _url_label[:60] + "…"
        _ref_label = _item["ref"] if _item["ref"] else "—"
        st.markdown(
            f"**{_item['num']}.** `{_item['image_name']}` · "
            f"Estado: **{_item['estado']}** · "
            f"URL: {_url_label} · Ref: {_ref_label}"
        )

    st.markdown("---")
    st.button(
        f"▶ Analizar {_n} comparable(s) con Claude",
        type="primary",
        disabled=True,
        help="El análisis con Claude se adaptará en la siguiente fase de implementación.",
    )
    st.caption("⚙️ Análisis pendiente de adaptar al nuevo flujo guiado.")
else:
    st.info("Usa el panel izquierdo para añadir comparables al lote.")


# ─────────────────────────────────────────────────────────────────────────────
# RESULTADOS
# ─────────────────────────────────────────────────────────────────────────────

if st.session_state.resultados:
    st.markdown("---")
    st.markdown(f"### 📊 Resultados — {len(st.session_state.resultados)} comparable(s)")

    for item in st.session_state.resultados:
        num       = item.get("comparable", "?")
        img_name  = item.get("img_name", "")
        campos    = item.get("campos", {})
        flags     = item.get("flags", {})
        desc      = flags.get("campos_desconocidos", [])
        tipo      = campos.get("B5", "—")
        municipio = campos.get("F2", "—")
        precio    = campos.get("B9")

        if isinstance(precio, (int, float)):
            precio_label = f"{precio:,.0f} €"
        else:
            precio_label = str(precio) if precio else "—"

        status_cls = "error" if desc else "done"
        icon = "⚠️" if desc else "✅"

        with st.expander(
            f"{icon} **COMPARABLE {num}** · {tipo} · {municipio} · {precio_label} — {img_name}",
            expanded=True,
        ):
            render_extracted_data(item)

    # ── Generación y descarga del XLSX ──────────────────────────────────────
    st.markdown("---")
    col_dl1, col_dl2, col_dl3 = st.columns([2, 2, 1])

    with col_dl1:
        nombre_archivo = st.text_input(
            "Nombre del archivo de salida",
            value=f"Comparables_{fecha_str.replace('/', '-')}",
            placeholder="Comparables_22-05-2026",
        )

    with col_dl2:
        try:
            xlsx_bytes = build_xlsx_multisheet(st.session_state.resultados)
            fname = f"{nombre_archivo}.xlsx"

            st.download_button(
                label=f"⬇ Descargar {fname}",
                data=xlsx_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Error al generar el XLSX: {e}")

    with col_dl3:
        total_unknown = sum(
            len(item.get("flags", {}).get("campos_desconocidos", []))
            for item in st.session_state.resultados
        )
        st.metric("Campos a revisar", total_unknown, delta=None)

    # ── JSON bruto (expander colapsado) ──────────────────────────────────────
    with st.expander("🔎 Ver JSON extraído (debug)", expanded=False):
        st.json(st.session_state.resultados)
