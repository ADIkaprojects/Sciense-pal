# app.py
import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime

# Load environment variables from .env file
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    # Safe fallback if python-dotenv is not installed yet
    if os.path.exists(".env"):
        with open(".env", "r") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, val = line.split("=", 1)
                    os.environ[key.strip()] = val.strip()

from generate_items import run_pipeline
from curriculum_analyzer import analyze_curriculum_alignment, CHAPTER_CODES

import re

# Load chapter mapping from Science Class 6.xlsx
CHAPTER_IDS = {
    "a journey through states of water": "67ff3d8b2613c0dfa772579b",
    "beyond earth": "67ff3d8a2613c0dfa7725785",
    "diversity in the living world": "67ff3d8a2613c0dfa7725742",
    "exploring magnets": "67ff3d8a2613c0dfa772574f",
    "living creatures": "67ff3d8b2613c0dfa7725792",
    "living creatures: exploring their characteristics": "67ff3d8b2613c0dfa7725792",
    "materials around us": "67ff3d8a2613c0dfa7725765",
    "measurement of length and motion": "67ff3d8a2613c0dfa772576c",
    "method of separation in everyday life": "68d260c812d191aeb6c953e6",
    "methods of separation": "68d260c812d191aeb6c953e6",
    "mindful eating": "67ff3d8a2613c0dfa772575a",
    "mindful eating: a path to a healthy body": "67ff3d8a2613c0dfa772575a",
    "nature's treasures": "67ff3d8b2613c0dfa77257af",
    "nature’s treasures": "67ff3d8b2613c0dfa77257af",
    "temperature and its measurement": "67ff3d8a2613c0dfa7725771",
    "the wonderful world of science": "68d2604a12d191aeb6c9320b",
}

try:
    science_file_path = "Science Class 6.xlsx"
    if not os.path.exists(science_file_path):
        science_file_path = os.path.join(os.path.dirname(os.getcwd()), "Science Class 6.xlsx")
    if not os.path.exists(science_file_path):
        parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) if '__file__' in locals() else os.path.dirname(os.getcwd())
        science_file_path = os.path.join(parent_dir, "Science Class 6.xlsx")
        
    df_science = pd.read_excel(science_file_path)
    for idx, row in df_science.iterrows():
        chap = row.get("Chapter")
        cid = row.get("Chapter ID")
        if chap and cid:
            norm_name = re.sub(r'\s+', ' ', str(chap).lower().strip())
            CHAPTER_IDS[norm_name] = str(cid).strip()
except Exception as e:
    pass

# Load chapters from Question Bank Count (1).xlsx
try:
    parent_dir = os.path.dirname(os.getcwd())
    qbank_path = os.path.join(parent_dir, "Question Bank Count (1).xlsx")
    if not os.path.exists(qbank_path):
        qbank_path = "Question Bank Count (1).xlsx"
    df_qbank = pd.read_excel(qbank_path)
    df_qbank_g6 = df_qbank[df_qbank['Grade'] == 6]
    qbank_chapters = df_qbank_g6['Chapter'].dropna().map(lambda x: str(x).strip()).unique().tolist()
except Exception as e:
    # Fallback to predefined Grade 6 chapters if loading fails
    qbank_chapters = [
        "Materials Around Us",
        "Living Creatures",
        "Diversity in the Living World",
        "Exploring Magnets",
        "Mindful Eating",
        "Measurement of Length and Motion",
        "Temperature and its Measurement",
        "A Journey through States of Water",
        "Methods of Separation",
        "Nature’s Treasures",
        "Beyond Earth"
    ]

st.set_page_config(
    page_title="Science PAL — Item Bank Generator",
    page_icon="🔬",
    layout="wide"
)

# ─── Initialize Session State Defaults ──────────────────────────────────────────
if "api_key" not in st.session_state:
    st.session_state["api_key"] = os.environ.get("MISTRAL_API_KEY", "")
if "chapter_name" not in st.session_state:
    st.session_state["chapter_name"] = "Materials Around Us"
if "chapter_code" not in st.session_state:
    st.session_state["chapter_code"] = "MAT"
if "topic_code" not in st.session_state:
    st.session_state["topic_code"] = "PROP"
if "ncf_cg" not in st.session_state:
    st.session_state["ncf_cg"] = "CG-1"
if "competency" not in st.session_state:
    st.session_state["competency"] = "C-1.1"
if "learning_outcome" not in st.session_state:
    st.session_state["learning_outcome"] = "Identifies and explains different properties of materials and relates them to their uses."
if "lo_id" not in st.session_state:
    st.session_state["lo_id"] = "LO-1.1.a"
if "start_seq" not in st.session_state:
    st.session_state["start_seq"] = 1

# ─── File Upload Callback Handler ──────────────────────────────────────────────
def handle_file_upload():
    uploaded_file = st.session_state.get("uploader")
    if uploaded_file is not None:
        api_key_val = st.session_state.get("api_key", os.environ.get("MISTRAL_API_KEY", ""))
        alignment = analyze_curriculum_alignment(uploaded_file, api_key_val)
        st.session_state["chapter_name"] = alignment["chapter_name"]
        st.session_state["chapter_code"] = alignment["chapter_code"]
        st.session_state["topic_code"] = alignment["topic_code"]
        st.session_state["ncf_cg"] = alignment["ncf_cg"]
        st.session_state["competency"] = alignment["competency"]
        st.session_state["learning_outcome"] = alignment["learning_outcome"]
        st.session_state["lo_id"] = alignment["lo_id"]
        st.session_state["start_seq"] = alignment["start_seq"]

# ─── Header ────────────────────────────────────────────────────────────────────
st.title("🔬 Science PAL — Bulk Item Authoring Pipeline")
st.markdown("*Grade 6 · AI-Powered Item Bank Generator · Built for TechCurators*")
st.divider()

# ─── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Pipeline Configuration")

    # API Key
    api_key_input = st.text_input(
        "Mistral API Key",
        value=st.session_state.get("api_key", os.environ.get("MISTRAL_API_KEY", "")),
        type="password",
        help="Enter key or set MISTRAL_API_KEY env variable before launching. Enter 'mock' to run locally with high-fidelity mock AI answers.",
        placeholder="sk-..."
    )
    st.session_state["api_key"] = api_key_input
    
    if st.session_state["api_key"]:
        st.success("✅ API key loaded")
    else:
        st.warning("⚠️ API key required")

    st.divider()
    st.subheader("📚 Batch Metadata")
    st.caption("Applied to ALL items in this upload.")

    # Selectbox dropdown menu for Chapter Name
    default_chap = st.session_state.get("chapter_name", "Materials Around Us")
    default_idx = 0
    if default_chap in qbank_chapters:
        default_idx = qbank_chapters.index(default_chap)
    else:
        for idx, chap in enumerate(qbank_chapters):
            if chap.lower().strip() == default_chap.lower().strip():
                default_idx = idx
                break

    chapter_name_input = st.selectbox(
        "Chapter Name *",
        options=qbank_chapters,
        index=default_idx,
        help="Select the chapter name from the Question Bank Count list."
    )
    st.session_state["chapter_name"] = chapter_name_input
    
    # Automatically map and resolve chapter code to Chapter ID
    norm_chap = re.sub(r'\s+', ' ', chapter_name_input.lower().strip())
    chap_code = CHAPTER_IDS.get(norm_chap, chapter_name_input[:3].upper())
    st.session_state["chapter_code"] = chap_code


# ─── Main Area ────────────────────────────────────────────────────────────────
col_main, col_info = st.columns([3, 1])

with col_main:
    st.subheader("📤 Upload Question Bank")
    uploaded_file = st.file_uploader(
        "Upload `Qs_Creation_Template_PAL_Science.xlsx` (or compatible file)",
        type=["xlsx", "xls"],
        help="Must contain a sheet with the 9 required columns.",
        key="uploader",
        on_change=handle_file_upload
    )

with col_info:
    st.subheader("📋 Required Columns")
    st.code(
        "Mastery_Level\nStimulus_Text\nItem_Stem\n"
        "Option_A\nOption_B\nOption_C\nOption_D\n"
        "Correct_Option\nCorrect_Answer",
        language="text"
    )

# ─── File Preview ──────────────────────────────────────────────────────────────
if uploaded_file:
    try:
        # Load sheets to inspect
        wb = pd.ExcelFile(uploaded_file)
        matching_sheet = None
        for name in wb.sheet_names:
            if name.strip().lower() == "questions":
                matching_sheet = name
                break
        if matching_sheet is None and len(wb.sheet_names) > 0:
            matching_sheet = wb.sheet_names[0]
        if matching_sheet is not None:
            df_preview = pd.read_excel(uploaded_file, sheet_name=matching_sheet)
            df_preview.reset_index(drop=True, inplace=True)
            # Only rows with at least a stem
            valid_rows = df_preview[df_preview["Item_Stem"].notna()].copy()
            st.info(
                f"**{len(valid_rows)} item(s)** found with a valid Item Stem "
                f"out of {len(df_preview)} total rows."
            )
            
            # Display Auto-Detected Curriculum Alignment & Sequence
            st.markdown("### 🎯 Curriculum Alignment & Status")
            
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**📚 Selected Chapter:** `{st.session_state.get('chapter_name', 'N/A')}`")
                st.markdown(f"**🔑 Chapter Code:** `{st.session_state.get('chapter_code', 'N/A')}`")
            with col2:
                st.markdown(f"**🎯 Topic & Subject IDs:** `Auto-detected per item`")
                st.markdown(f"**🔢 Sequence Numbering:** `Auto-managed per topic`")
            st.divider()

            with st.expander("🔍 Preview Questions Sheet", expanded=False):
                st.dataframe(df_preview, width="stretch")

            col_m = st.columns(4)
            for i, level in enumerate(["M1", "M2", "M3", "M4"]):
                count = len(valid_rows[valid_rows["Mastery_Level"].astype(str).str.upper() == level])
                col_m[i].metric(f"{level} Items", count)
        else:
            st.error("Error: The uploaded file is empty and has no sheets.")
            valid_rows = pd.DataFrame()

    except Exception as e:
        st.error(f"Could not read Excel file: {e}")
        valid_rows = pd.DataFrame()
else:
    valid_rows = pd.DataFrame()

# ─── Validation Gate ─────────────────────────────────────────────────────────
ready = (
    uploaded_file is not None
    and st.session_state.get("api_key", "").strip() != ""
    and st.session_state.get("chapter_name", "").strip() != ""
    and st.session_state.get("chapter_code", "").strip() != ""
    and len(valid_rows) > 0
)

if not ready:
    missing = []
    if not uploaded_file: missing.append("Excel file")
    if not st.session_state.get("api_key", "").strip(): missing.append("Mistral API key")
    if not st.session_state.get("chapter_name", "").strip(): missing.append("Chapter Name")
    if not st.session_state.get("chapter_code", "").strip(): missing.append("Chapter Code")
    if missing:
        st.warning(f"⚠️ Please provide: {', '.join(missing)}")

st.divider()

# ─── Run Button ───────────────────────────────────────────────────────────────
if ready:
    st.subheader("🚀 Generate Item Bank")
    col_run, col_est = st.columns([2, 1])
    with col_est:
        est_api_calls = len(valid_rows) * 5
        est_secs = len(valid_rows) * 3 if st.session_state.get("api_key", "").strip().lower() in ("mock", "test") else len(valid_rows) * 15
        st.metric("Est. API Calls", 0 if st.session_state.get("api_key", "").strip().lower() in ("mock", "test") else est_api_calls)
        st.metric("Est. Time", f"~{max(1, est_secs // 60)} min" if est_secs >= 60 else f"{est_secs} secs")

    with col_run:
        run_clicked = st.button(
            "▶  Run Pipeline",
            type="primary",
            width="stretch"
        )

    if run_clicked:
        progress_bar = st.progress(0.0, text="Initialising pipeline…")
        status_placeholder = st.empty()
        log_expander = st.expander("📋 Live Processing Log", expanded=True)
        log_placeholder = log_expander.empty()
        live_logs = []

        def progress_cb(item_num: int, total: int, item_id: str, level: str, msg: str):
            pct = item_num / total
            progress_bar.progress(pct, text=f"Processing item {item_num}/{total} — {item_id}")
            status_placeholder.caption(f"Last: [{level}] {item_id} — {msg}")
            live_logs.append(f"[{level}] Row {item_num:>3} | {item_id:<28} | {msg}")
            log_placeholder.code("\n".join(live_logs[-25:]), language="text")

        batch_meta = {
            "chapter_name": st.session_state.get("chapter_name", "").strip(),
            "chapter_code": st.session_state.get("chapter_code", "").strip().upper(),
            "ncf_cg": st.session_state.get("ncf_cg", "").strip(),
            "competency": st.session_state.get("competency", "").strip(),
            "learning_outcome": st.session_state.get("learning_outcome", "").strip(),
            "lo_id": st.session_state.get("lo_id", "").strip(),
            "start_seq": int(st.session_state.get("start_seq", 1)),
        }

        try:
            docx_bytes, csv_bytes, summary = run_pipeline(
                excel_file=uploaded_file,
                api_key=st.session_state.get("api_key", "").strip(),
                batch_meta=batch_meta,
                progress_callback=progress_cb
            )
            progress_bar.progress(1.0, text="✅ Pipeline complete!")
            status_placeholder.empty()

            # Summary
            st.success(
                f"✅ **Done.** {summary['written']} item(s) written · "
                f"{summary['skipped']} skipped · "
                f"{summary['overrides']} override(s) applied."
            )

            st.subheader("📥 Download Outputs")
            today = datetime.today().strftime("%Y%m%d")
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                st.download_button(
                    "📄 Download Item Bank (DOCX)",
                    data=docx_bytes,
                    file_name=f"Science_PAL_G6_Item_Bank_{today}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    width="stretch",
                )
            with col_dl2:
                st.download_button(
                    "📊 Download Audit Log (CSV)",
                    data=csv_bytes,
                    file_name=f"Science_PAL_G6_Item_Bank_{today}_log.csv",
                    mime="text/csv",
                    width="stretch",
                )

        except Exception as exc:
            st.error(f"Pipeline failed: {exc}")
            st.exception(exc)
