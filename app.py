import streamlit as st
import tempfile
import os
import shutil
from processor import run_export

st.set_page_config(
    page_title="WA Exporter",
    page_icon="🍇",
    layout="centered"
)

st.title("🍇 WA Exporter")
st.caption("Clean Excel to WhatsApp import format")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

st.sidebar.header("Settings")

input_level = st.sidebar.selectbox(
    "Input Level",
    ["DUN", "PARLIMEN", "MIXED"]
)

structure = st.sidebar.selectbox(
    "Output Structure",
    [
        "DUN folder > individual DM Excel files",
        "DUN folder > DM folder > one Excel inside each DM folder",
        "DUN folder > DM folder > separate KAUM Excel files",
        "DUN folder > one DUN Excel file",
        "DUN folder > separate KAUM Excel files",
        "PARLIMEN folder > DUN folder > individual DM Excel files",
        "PARLIMEN folder > DUN folder > DM folder > separate KAUM Excel files",
        "ONE Excel file only",
        "AGE range Excel files only",
    ]
)

split_mode = st.sidebar.selectbox(
    "Split File Mode",
    [
        "All rows in one Excel",
        "Separate Excel files by KAUM",
        "Separate Excel files by gender code",
    ]
)

kaum_filter = st.sidebar.selectbox(
    "Kaum Filter",
    [
        "ALL",
        "MELAYU,CINA,INDIA",
        "MELAYU",
        "CINA",
        "INDIA",
        "LAIN-LAIN",
    ]
)

sikap_filter = st.sidebar.selectbox(
    "Sikap Filter",
    [
        "No filter",
        "HITAM",
        "KELABU",
        "PUTIH",
        "HITAM,KELABU",
        "KELABU,PUTIH",
    ]
)

party_filter = st.sidebar.text_input("Party filter", placeholder="Example: PKR")

use_age_filter = st.sidebar.checkbox("Apply age filter")
min_age = st.sidebar.number_input("Min age", min_value=0, max_value=120, value=18)
max_age_text = st.sidebar.text_input("Max age", placeholder="Blank = no max")

dedup_phone = st.sidebar.checkbox("Remove duplicate phone numbers", value=True)
dedup_nokp = st.sidebar.checkbox("Remove duplicate NO KP", value=True)
read_all_sheets = st.sidebar.checkbox("Read all sheets", value=True)

run_button = st.button("🍇 Run Export", type="primary")

if run_button:
    if not uploaded_files:
        st.error("Upload at least one Excel file first.")
        st.stop()

    with tempfile.TemporaryDirectory() as temp_dir:
        input_dir = os.path.join(temp_dir, "input")
        output_dir = os.path.join(temp_dir, "voter_outputs")
        os.makedirs(input_dir, exist_ok=True)

        saved_paths = []

        for uploaded_file in uploaded_files:
            file_path = os.path.join(input_dir, uploaded_file.name)

            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            saved_paths.append(file_path)

        config = {
            "input_level": input_level,
            "structure": structure,
            "split_mode": split_mode,
            "kaum_filter": kaum_filter,
            "sikap_filter": sikap_filter,
            "party_filter": party_filter,
            "use_age_filter": use_age_filter,
            "min_age": min_age,
            "max_age": max_age_text,
            "dedup_phone": dedup_phone,
            "dedup_nokp": dedup_nokp,
            "read_all_sheets": read_all_sheets,
            "output_dir": output_dir,
        }

        with st.status("Processing files...", expanded=True) as status:
            result = run_export(saved_paths, config)
            status.update(label="Done", state="complete")

        st.success("Export completed.")

        st.write(f"Final rows: **{result['final_rows']:,}**")
        st.write(f"Excel files created: **{result['files_created']:,}**")

        with open(result["zip_path"], "rb") as f:
            st.download_button(
                label="Download ZIP",
                data=f,
                file_name="voter_outputs.zip",
                mime="application/zip"
            )