"""
Streamlit app: Lighthouse Studio (.sav export) -> SPSS mapping & syntax automation.

RUN WITH:
    streamlit run app.py

(Do NOT run this with `python app.py` — Streamlit apps must be launched via
the `streamlit run` command so the Streamlit server/UI framework starts.)
"""

import io
import os
import tempfile

import streamlit as st
import openpyxl

from lighthouse_core import build_mapping_workbook, build_syntax_text

st.set_page_config(page_title="Lighthouse → SPSS Automation", layout="wide")

st.title("Lighthouse Studio → SPSS Mapping & Syntax Automation")
st.caption(
    "Stage 1 turns a raw .sav export into a review-ready mapping file. "
    "Stage 2 turns the finalized mapping into ready-to-run SPSS syntax."
)

tab1, tab2 = st.tabs(["Stage 1 — Generate Mapping", "Stage 2 — Generate Syntax"])

# ---------------------------------------------------------------------------
# STAGE 1
# ---------------------------------------------------------------------------
with tab1:
    st.subheader("1. Upload your raw .sav export")
    sav_file = st.file_uploader("Lighthouse SPSS export (.sav)", type=["sav"], key="sav_stage1")

    st.subheader("2. Optional: codebook for per-option grid text")
    st.caption(
        "Two columns: Variable, Option Text. Only needed for grid/multi-select "
        "items where Lighthouse's .sav export doesn't include the specific "
        "answer-option wording."
    )
    codebook_file = st.file_uploader("Codebook (.xlsx)", type=["xlsx"], key="codebook_stage1")

    if sav_file is not None:
        if st.button("Generate mapping file", type="primary"):
            with st.spinner("Reading .sav and applying rename/label rules..."):
                # pyreadstat needs a real file path, so write the upload to a temp file
                with tempfile.NamedTemporaryFile(suffix=".sav", delete=False) as tmp_sav:
                    tmp_sav.write(sav_file.getvalue())
                    sav_path = tmp_sav.name

                codebook_arg = None
                if codebook_file is not None:
                    codebook_arg = io.BytesIO(codebook_file.getvalue())

                try:
                    wb, stats = build_mapping_workbook(sav_path, codebook_path=codebook_arg)
                finally:
                    os.unlink(sav_path)

            st.success("Mapping file generated.")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total variables", stats["total"])
            c2.metric("Excluded (helper vars)", stats["excluded"])
            c3.metric("Auto-resolved", stats["auto_resolved"])
            c4.metric("Flagged for review", stats["flagged"])

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)

            st.download_button(
                label="Download mapping .xlsx",
                data=buf,
                file_name="mapping_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.info(
                "Review the yellow-highlighted rows in the downloaded file, fix any "
                "'Suggested Rename' / 'Cleaned Variable Label' values that need it, "
                "then use that file in Stage 2."
            )

# ---------------------------------------------------------------------------
# STAGE 2
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("1. Upload your finalized mapping file")
    mapping_file = st.file_uploader("Finalized mapping (.xlsx)", type=["xlsx"], key="mapping_stage2")

    st.subheader("2. Upload the original .sav (for value labels)")
    sav_file2 = st.file_uploader("Original .sav export", type=["sav"], key="sav_stage2")

    if mapping_file is not None and sav_file2 is not None:
        if st.button("Generate SPSS syntax", type="primary"):
            with st.spinner("Building .sps syntax..."):
                with tempfile.NamedTemporaryFile(suffix=".sav", delete=False) as tmp_sav:
                    tmp_sav.write(sav_file2.getvalue())
                    sav_path = tmp_sav.name

                wb = openpyxl.load_workbook(io.BytesIO(mapping_file.getvalue()), data_only=True)

                try:
                    syntax_text, stats = build_syntax_text(wb, sav_path)
                except ValueError as e:
                    st.error(str(e))
                    syntax_text = None
                finally:
                    os.unlink(sav_path)

            if syntax_text:
                st.success("Syntax generated.")
                c1, c2, c3 = st.columns(3)
                c1.metric("Variables renamed", stats["renamed"])
                c2.metric("Variable labels written", stats["labeled"])
                c3.metric("Variables with value labels", stats["value_labeled"])

                st.download_button(
                    label="Download .sps syntax",
                    data=syntax_text,
                    file_name="generated_syntax.sps",
                    mime="text/plain",
                )
                with st.expander("Preview syntax"):
                    st.code(syntax_text, language=None)
