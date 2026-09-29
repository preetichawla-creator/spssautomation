"""
Streamlit app for the SPSS mapping-file workflow.

Run with:
    streamlit run app.py

Requires generate_mapping.py to be in the same folder -- this app is just a UI
wrapper around its functions; all the actual parsing/renaming/labeling and
syntax-generation logic lives there and is unchanged.

Two steps, one tab each:
  1. Generate mapping     -- upload raw .sav + questionnaire .docx, download a draft
                             mapping .xlsx (Variable Label sheet).
  2. Generate SPSS syntax -- upload that mapping .xlsx (reviewed/edited by hand, or
                             not), download RENAME VARIABLES + VARIABLE LABELS as a
                             ready-to-run .txt syntax file.
"""
import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

from generate_mapping import build_workbook, generate_spss_syntax

st.set_page_config(page_title="SPSS Mapping & Syntax Generator", layout="wide")
st.title("SPSS Mapping & Syntax Generator")

tab_map, tab_syntax = st.tabs(["1. Generate mapping", "2. Generate SPSS syntax"])

# ----------------------------------------------------------------------------
# Step 1: raw .sav + questionnaire .docx -> draft mapping .xlsx
# ----------------------------------------------------------------------------
with tab_map:
    st.caption(
        "Upload a raw .sav data file and the Word questionnaire for the same study "
        "to get a draft Variable Label mapping file."
    )

    col1, col2 = st.columns(2)
    with col1:
        sav_file = st.file_uploader("Raw data file (.sav)", type=["sav"])
    with col2:
        qnr_file = st.file_uploader("Questionnaire (.docx)", type=["docx"])

    generate = st.button("Generate mapping file", type="primary",
                          disabled=not (sav_file and qnr_file))

    if generate:
        with tempfile.TemporaryDirectory() as tmpdir:
            sav_path = Path(tmpdir) / sav_file.name
            qnr_path = Path(tmpdir) / qnr_file.name
            sav_path.write_bytes(sav_file.getvalue())
            qnr_path.write_bytes(qnr_file.getvalue())

            with st.spinner("Parsing questionnaire and raw data, building mapping..."):
                try:
                    wb = build_workbook(str(sav_path), str(qnr_path))
                except Exception as e:
                    st.error(f"Failed to generate mapping: {e}")
                    st.stop()

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)

            st.session_state["mapping_bytes"] = buf.getvalue()
            st.session_state["mapping_ready"] = True

    if st.session_state.get("mapping_ready"):
        st.success("Mapping file generated.")

        st.download_button(
            "Download mapping .xlsx",
            data=st.session_state["mapping_bytes"],
            file_name="draft_mapping.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        wb_preview = load_workbook(io.BytesIO(st.session_state["mapping_bytes"]), read_only=True)

        ws = wb_preview["Variable Label"]
        rows = list(ws.iter_rows(values_only=True))
        df = pd.DataFrame(rows[2:], columns=rows[1])  # skip title row, use header row

        total = len(df)
        has_note = df["Notes"].fillna("") != ""
        flagged = int(has_note.sum())

        m1, m2, m3 = st.columns(3)
        m1.metric("Total variables", total)
        m2.metric("Flagged for review", flagged)
        m3.metric("Auto-resolved", total - flagged)

        preview_tab1, preview_tab2 = st.tabs(["Flagged rows (review these first)", "All variables"])

        with preview_tab1:
            flagged_df = df[has_note]
            st.dataframe(flagged_df, width="stretch", height=500)

        with preview_tab2:
            st.dataframe(df, width="stretch", height=500)

# ----------------------------------------------------------------------------
# Step 2: mapping .xlsx -> SPSS syntax .txt
# ----------------------------------------------------------------------------
with tab_syntax:
    st.caption(
        "Upload the finalized mapping file: 3 columns, always in this order -- "
        "Raw Variable, Renamed Variable, Variable Label (column headers can say "
        "anything, only the position matters) -- to get ready-to-run SPSS syntax: "
        "RENAME VARIABLES, then VARIABLE LABELS."
    )
    st.caption(
        "Value Labels and RECODE aren't generated yet -- the mapping file doesn't "
        "carry value-level data at this stage."
    )

    mapping_file = st.file_uploader("Mapping file (.xlsx)", type=["xlsx"], key="mapping_upload")

    build_syntax = st.button("Generate SPSS syntax", type="primary", disabled=not mapping_file)

    if build_syntax:
        with tempfile.TemporaryDirectory() as tmpdir:
            mapping_path = Path(tmpdir) / mapping_file.name
            mapping_path.write_bytes(mapping_file.getvalue())

            with st.spinner("Reading mapping file and building syntax..."):
                try:
                    text, warnings = generate_spss_syntax(str(mapping_path))
                except Exception as e:
                    st.error(f"Failed to generate syntax: {e}")
                    st.stop()

            st.session_state["syntax_text"] = text
            st.session_state["syntax_warnings"] = warnings
            st.session_state["syntax_ready"] = True

    if st.session_state.get("syntax_ready"):
        text = st.session_state["syntax_text"]
        warnings = st.session_state["syntax_warnings"]

        if not text:
            st.warning("Nothing to generate from this file -- see details below.")
        else:
            st.success("SPSS syntax generated.")
            st.download_button(
                "Download syntax .txt",
                data=text,
                file_name="spss_syntax.txt",
                mime="text/plain",
            )

        if warnings:
            with st.expander(f"{len(warnings)} warning(s) -- review before running in SPSS", expanded=True):
                for w in warnings:
                    st.warning(w)

        if text:
            st.text_area("Preview", text, height=500)
