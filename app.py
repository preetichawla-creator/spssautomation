"""
Streamlit app for the SPSS mapping-file workflow.

Run with:
    streamlit run app.py

Requires generate_mapping.py to be in the same folder -- this app is just a UI
wrapper around its functions; all the actual parsing/renaming/labeling and
syntax-generation logic lives there and is unchanged.

Two steps, one tab each:
  1. Generate mapping     -- upload raw .sav + questionnaire (.docx or .xlsx), download a draft
                             mapping .xlsx (Variable Label sheet).
  2. Generate SPSS syntax -- upload the finalized mapping .xlsx plus the original raw
                             .sav and questionnaire (.docx or .xlsx); download TWO ready-to-run
                             .txt syntax files: (1) RENAME VARIABLES + VARIABLE LABELS,
                             (2) RECODE + Val Lab (value labels).
"""
import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

from generate_mapping import build_workbook, generate_spss_syntax, generate_recode_value_syntax

st.set_page_config(page_title="SPSS Mapping & Syntax Generator", layout="wide")
st.title("SPSS Mapping & Syntax Generator")

tab_map, tab_syntax = st.tabs(["1. Generate mapping", "2. Generate SPSS syntax"])

# ----------------------------------------------------------------------------
# Step 1: raw .sav + questionnaire (.docx or .xlsx) -> draft mapping .xlsx
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
        qnr_file = st.file_uploader("Questionnaire (.docx or .xlsx)", type=["docx", "xlsx"])

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
        "anything, only the position matters)."
    )
    st.caption(
        "Also upload the original raw .sav and questionnaire (.docx or .xlsx) for this study -- "
        "needed to work out which values need recoding and what their value labels "
        "say. These are the same two files used in step 1."
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        mapping_file = st.file_uploader("Mapping file (.xlsx)", type=["xlsx"], key="mapping_upload")
    with col2:
        syntax_sav_file = st.file_uploader("Raw data file (.sav)", type=["sav"], key="syntax_sav_upload")
    with col3:
        syntax_qnr_file = st.file_uploader("Questionnaire (.docx or .xlsx)", type=["docx", "xlsx"], key="syntax_qnr_upload")

    build_syntax = st.button(
        "Generate SPSS syntax", type="primary",
        disabled=not (mapping_file and syntax_sav_file and syntax_qnr_file),
    )

    if build_syntax:
        with tempfile.TemporaryDirectory() as tmpdir:
            mapping_path = Path(tmpdir) / mapping_file.name
            sav_path = Path(tmpdir) / syntax_sav_file.name
            qnr_path = Path(tmpdir) / syntax_qnr_file.name
            mapping_path.write_bytes(mapping_file.getvalue())
            sav_path.write_bytes(syntax_sav_file.getvalue())
            qnr_path.write_bytes(syntax_qnr_file.getvalue())

            with st.spinner("Building rename and label syntax..."):
                try:
                    rename_text, rename_warnings = generate_spss_syntax(str(mapping_path))
                except Exception as e:
                    st.error(f"Failed to generate rename/label syntax: {e}")
                    st.stop()

            with st.spinner("Building recode and value label syntax..."):
                try:
                    recode_text, recode_warnings = generate_recode_value_syntax(
                        str(sav_path), str(qnr_path), str(mapping_path))
                except Exception as e:
                    st.error(f"Failed to generate recode/value label syntax: {e}")
                    st.stop()

            st.session_state["rename_syntax_text"] = rename_text
            st.session_state["rename_syntax_warnings"] = rename_warnings
            st.session_state["recode_syntax_text"] = recode_text
            st.session_state["recode_syntax_warnings"] = recode_warnings
            st.session_state["syntax_ready"] = True

    if st.session_state.get("syntax_ready"):
        rename_text = st.session_state["rename_syntax_text"]
        rename_warnings = st.session_state["rename_syntax_warnings"]
        recode_text = st.session_state["recode_syntax_text"]
        recode_warnings = st.session_state["recode_syntax_warnings"]

        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.markdown("**File 1: Rename + Variable Labels**")
            if rename_text:
                st.download_button(
                    "Download rename_and_labels.txt",
                    data=rename_text,
                    file_name="rename_and_labels.txt",
                    mime="text/plain",
                    key="dl_rename",
                )
            else:
                st.warning("Nothing to generate -- see warnings below.")
        with dl_col2:
            st.markdown("**File 2: Recode + Value Labels**")
            if recode_text:
                st.download_button(
                    "Download recode_and_value_labels.txt",
                    data=recode_text,
                    file_name="recode_and_value_labels.txt",
                    mime="text/plain",
                    key="dl_recode",
                )
            else:
                st.warning("Nothing to generate -- see warnings below.")

        all_warnings = [("Rename/Labels", w) for w in rename_warnings] + \
                       [("Recode/Value Labels", w) for w in recode_warnings]
        if all_warnings:
            with st.expander(f"{len(all_warnings)} warning(s) -- review before running in SPSS", expanded=True):
                for source, w in all_warnings:
                    st.warning(f"[{source}] {w}")

        preview_col1, preview_col2 = st.columns(2)
        with preview_col1:
            if rename_text:
                st.text_area("Preview: rename_and_labels.txt", rename_text, height=400, key="preview_rename")
        with preview_col2:
            if recode_text:
                st.text_area("Preview: recode_and_value_labels.txt", recode_text, height=400, key="preview_recode")
