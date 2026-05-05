"""
app.py  -  Streamlit UI for the Automated CDR Tool (BAAI BGE)
--------------------------------------------------------------
Run with:
    streamlit run app.py
"""

import io
import json
import os
import sys
import tempfile
import zipfile

import pandas as pd
import streamlit as st

# ── Make src/ importable when running from the BAAI BGE/ directory ─────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from src.heading_extractor import extract_headings
from src.heading_matcher   import load_model, match_headings
from src.template_filler   import fill_template

# ---------------------------------------------------------------------------
# Page config (must be first Streamlit call)
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Automated CDR Tool",
    page_icon="📄",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Custom CSS — colourful theme
# ---------------------------------------------------------------------------

st.markdown("""
<style>
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
    min-height: 100vh;
}
[data-testid="stHeader"] { background: transparent; }
[data-testid="stFileUploader"] {
    background: rgba(255,255,255,0.06);
    border: 1px solid rgba(255,255,255,0.15);
    border-radius: 14px;
    padding: 18px 16px;
}
h1 {
    background: linear-gradient(90deg, #a78bfa, #60a5fa, #34d399);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-size: 2.4rem !important;
    font-weight: 800 !important;
    text-align: center;
    padding-bottom: 4px;
}
h3 { color: #a78bfa !important; }
p, label, .stCaption, [data-testid="stMarkdownContainer"] p { color: #cbd5e1 !important; }
[data-testid="stMetric"] {
    background: rgba(255,255,255,0.07);
    border-radius: 12px;
    padding: 10px 14px;
    border: 1px solid rgba(255,255,255,0.1);
}
[data-testid="stMetricValue"] { color: #60a5fa !important; font-size: 1.6rem !important; }
[data-testid="stMetricLabel"] { color: #94a3b8 !important; }
[data-testid="stButton"] > button[kind="primary"] {
    background: linear-gradient(90deg, #7c3aed, #2563eb) !important;
    border: none !important;
    border-radius: 10px !important;
    color: #fff !important;
    font-weight: 700 !important;
    font-size: 1.05rem !important;
    padding: 12px 0 !important;
    transition: opacity 0.2s;
}
[data-testid="stButton"] > button[kind="primary"]:hover { opacity: 0.88; }
[data-testid="stButton"] > button[kind="primary"]:disabled {
    background: rgba(255,255,255,0.1) !important;
    color: #64748b !important;
}
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(90deg, #059669, #0891b2) !important;
    border: none !important;
    border-radius: 10px !important;
    color: #fff !important;
    font-weight: 700 !important;
    font-size: 1.05rem !important;
    padding: 12px 0 !important;
}
[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }
hr { border-color: rgba(255,255,255,0.1) !important; }
.step-badge {
    display: inline-block;
    background: linear-gradient(90deg, #7c3aed, #2563eb);
    color: #fff;
    border-radius: 20px;
    padding: 3px 14px;
    font-size: 0.78rem;
    font-weight: 700;
    margin-bottom: 6px;
}
[data-testid="stFileUploaderDropzoneInstructions"] { color: #94a3b8 !important; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

@st.cache_resource(show_spinner="🔄 Loading BAAI/bge-base-en-v1.5 model …")
def get_model():
    return load_model()


def derive_output_name(input_filename: str) -> str:
    return f"op_{input_filename}"


def _load_prompts() -> dict:
    prompts_path = os.path.join(BASE_DIR, "prompts.json")
    with open(prompts_path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def auto_detect_product_name(docx_bytes: bytes) -> str | None:
    """Read product name from the input document's Word core properties.

    Tries title -> subject -> description in order; returns None if all empty.
    """
    try:
        from docx import Document as _Doc
        doc = _Doc(io.BytesIO(docx_bytes))
        cp = doc.core_properties
        for attr in ("title", "subject", "description"):
            val = (getattr(cp, attr, None) or "").strip()
            if val:
                return val
    except Exception:
        pass
    return None


def run_pipeline(
    template_bytes: bytes,
    template_name: str,
    input_bytes: bytes,
    input_name: str,
    threshold: float,
) -> tuple:
    """Run the full CDR pipeline for one input file.

    Returns:
        (output_bytes | None, matches: list[dict], detected_product_name: str | None, error: str)
    """
    prompts = _load_prompts()

    # Auto-detect product name from the input doc's Word properties;
    # fall back to prompts.json value only if doc has none set.
    product_name = auto_detect_product_name(input_bytes)
    if not product_name:
        product_name = prompts.get("product_name") or None

    with tempfile.TemporaryDirectory() as tmp:
        template_path = os.path.join(tmp, template_name)
        input_path    = os.path.join(tmp, input_name)
        output_path   = os.path.join(tmp, "output.docx")

        with open(template_path, "wb") as fh:
            fh.write(template_bytes)
        with open(input_path, "wb") as fh:
            fh.write(input_bytes)

        data_headings = extract_headings(input_path)
        if not data_headings:
            return None, [], None, f"No headings found in: {input_name}"

        model             = get_model()
        template_headings = extract_headings(template_path)
        if not template_headings:
            return None, [], None, f"No headings found in template: {template_name}"

        matches = match_headings(
            template_headings, data_headings, model, threshold=threshold,
        )

        fill_template(
            template_path, input_path, matches, output_path,
            product_name=product_name,
        )

        if not os.path.isfile(output_path):
            return None, [], None, "fill_template did not produce an output file."

        with open(output_path, "rb") as fh:
            result_bytes = fh.read()

    return result_bytes, matches, product_name, ""


def colour_score(val: float) -> str:
    """Cell colour for score column in the match table."""
    if val >= 0.85:
        return "background-color:#064e3b;color:#6ee7b7"
    if val >= 0.70:
        return "background-color:#1e3a5f;color:#93c5fd"
    return "background-color:#4c1d1d;color:#fca5a5"


def make_zip(results: list) -> bytes:
    """Bundle all successful output .docx files into a ZIP archive."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in results:
            if r["bytes"]:
                zf.writestr(r["out_name"], r["bytes"])
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------

if "results" not in st.session_state:
    st.session_state.results       = []   # list[dict] — one entry per input file
    st.session_state.template_name = ""

# ---------------------------------------------------------------------------
# Sidebar — live settings (defaults from prompts.json, editable by user)
# ---------------------------------------------------------------------------

_prompts_defaults  = _load_prompts()
_default_threshold = float(_prompts_defaults.get("threshold", 0.6))

with st.sidebar:
    st.markdown("## ⚙️ Settings")
    threshold = st.slider(
        "Match threshold",
        min_value=0.0, max_value=1.0,
        value=_default_threshold, step=0.05,
        help="Minimum cosine-similarity score to accept a heading match. "
             "Higher = stricter matching.",
    )
    st.caption(
        f"Default from prompts.json: **{_default_threshold}**  \n"
        "🟢 ≥ 0.85 · 🔵 ≥ 0.70 · 🔴 < 0.70"
    )
    st.divider()
    st.markdown(
        "**Product name** is auto-detected from each input document's "
        "Word core properties (title → subject → description).  \n"
        "Falls back to `product_name` in *prompts.json* if not found."
    )

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------

st.markdown("<h1>📄 Automated CDR Tool</h1>", unsafe_allow_html=True)
st.markdown(
    "<p style='text-align:center;color:#94a3b8;margin-top:-8px;'>"
    "BAAI/bge-base-en-v1.5 &nbsp;·&nbsp; Semantic heading matching &nbsp;·&nbsp; DOCX → DOCX"
    "</p>",
    unsafe_allow_html=True,
)
st.divider()

# ---------------------------------------------------------------------------
# Upload — STEP 1 (template) and STEP 2 (1–5 inputs)
# ---------------------------------------------------------------------------

col1, col2 = st.columns([1, 2])

with col1:
    st.markdown('<div class="step-badge">STEP 1 — Template</div>', unsafe_allow_html=True)
    st.markdown("**Upload CDR Template** (1 file)")
    st.caption("e.g. *Template_D001024021 CDR ISO 17664-2 ... .docx*")
    template_file = st.file_uploader(
        "CDR template (.docx)", type=["docx"], key="template",
        label_visibility="collapsed",
    )
    if template_file:
        st.success(f"✔ {template_file.name}")

with col2:
    st.markdown('<div class="step-badge">STEP 2 — Inputs</div>', unsafe_allow_html=True)
    st.markdown("**Upload Test Records** (any number of files)")
    st.caption("e.g. *Input1_... .docx*, *Input7_... .docx*, …")
    input_files = st.file_uploader(
        "Test records (.docx)", type=["docx"], key="inputs",
        accept_multiple_files=True,
        label_visibility="collapsed",
    )
    if input_files:
        for f in input_files:
            st.success(f"✔ {f.name}")

st.divider()

# ---------------------------------------------------------------------------
# Run — STEP 3
# ---------------------------------------------------------------------------

inputs_ready = template_file is not None and bool(input_files)
st.markdown('<div class="step-badge">STEP 3 — Run</div>', unsafe_allow_html=True)
run_btn = st.button(
    "▶  Run Pipeline",
    disabled=not inputs_ready,
    use_container_width=True,
    type="primary",
)

# ---------------------------------------------------------------------------
# Pipeline execution
# ---------------------------------------------------------------------------

if run_btn:
    st.session_state.results       = []
    st.session_state.template_name = template_file.name

    files_to_process = input_files or []
    total            = len(files_to_process)
    progress         = st.progress(0, text="Starting …")

    for idx, inp in enumerate(files_to_process):
        progress.progress(
            int(idx / total * 90),
            text=f"📂 Processing {idx + 1}/{total}: {inp.name} …",
        )

        out_bytes, matches, detected_name, err = run_pipeline(
            template_bytes=template_file.getvalue(),
            template_name=template_file.name,
            input_bytes=inp.getvalue(),
            input_name=inp.name,
            threshold=threshold,
        )

        st.session_state.results.append({
            "input_name":   inp.name,
            "out_name":     derive_output_name(inp.name),
            "bytes":        out_bytes,
            "matches":      matches,
            "product_name": detected_name,
            "error":        err,
        })

    progress.progress(100, text="✅ All done!")

# ---------------------------------------------------------------------------
# Results — rendered from session state (stable across reruns)
# ---------------------------------------------------------------------------

if st.session_state.results:
    results = st.session_state.results
    n_ok    = sum(1 for r in results if r["bytes"])
    n_fail  = len(results) - n_ok

    st.divider()
    st.markdown("### 📊 Results")
    st.info(f"**Template:** {st.session_state.template_name}")

    sc1, sc2, sc3 = st.columns(3)
    sc1.metric("Files processed", len(results))
    sc2.metric("✅ Succeeded",    n_ok)
    sc3.metric("❌ Failed",       n_fail)

    st.divider()

    # ── Per-file expanders ────────────────────────────────────────────────
    for r in results:
        icon = "✅" if r["bytes"] else "❌"
        with st.expander(f"{icon}  {r['input_name']}", expanded=(len(results) == 1)):
            if r["error"]:
                st.error(f"Pipeline failed: {r['error']}")
                continue

            matches       = r["matches"]
            matched_count = sum(1 for m in matches if m["matched_heading"] is not None)
            total_count   = len(matches)

            mc1, mc2, mc3 = st.columns(3)
            mc1.metric("Template headings", total_count)
            mc2.metric("Matched",           matched_count)
            mc3.metric("Unmatched",         total_count - matched_count)

            pname = r.get("product_name")
            if pname:
                st.caption(f"🏷 Product name detected: **{pname}**")
            else:
                st.caption("🏷 Product name: *not detected in doc properties — no header/footer replacement*")

            rows = []
            for m in matches:
                rows.append({
                    "Template Heading": m["template_heading"]["text"],
                    "Score":            round(m["score"], 2),
                    "Data-Doc Heading": (
                        m["matched_heading"]["text"]
                        if m["matched_heading"]
                        else "— NO MATCH —"
                    ),
                })
            df = pd.DataFrame(rows)
            st.dataframe(
                df.style.map(colour_score, subset=["Score"]),
                use_container_width=True,
                hide_index=True,
            )

            st.download_button(
                label=f"⬇  Download  {r['out_name']}",
                data=r["bytes"],
                file_name=r["out_name"],
                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".wordprocessingml.document"
                ),
                use_container_width=True,
                key=f"dl_{r['input_name']}",
            )

    # ── Download All as ZIP (when multiple outputs succeeded) ─────────────
    if n_ok > 1:
        st.divider()
        st.markdown(
            '<div class="step-badge">STEP 4 — Download All</div>',
            unsafe_allow_html=True,
        )
        st.markdown("**Bundle all outputs into one ZIP file**")
        st.download_button(
            label="⬇  Download All as ZIP",
            data=make_zip(results),
            file_name="CDR_outputs.zip",
            mime="application/zip",
            use_container_width=True,
        )
