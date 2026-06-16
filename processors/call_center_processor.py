import streamlit as st
import pandas as pd
from processors.call_center_processor import (
    read_input_file,
    standardize_columns,
    clean_text_columns,
    process_files
)

st.set_page_config(
    page_title="Call Center Cleaner",
    page_icon="📞",
    layout="wide"
)

st.title("📞 Call Center Number Cleaner")
st.caption("Clean call center Excel/CSV files, remove invalid phone numbers, split CSVs, and export ZIP.")

uploaded_files = st.file_uploader(
    "Upload Excel or CSV files",
    type=["xlsx", "csv"],
    accept_multiple_files=True
)

with st.sidebar:
    st.header("Cleaning Options")

    start_id = st.text_input(
        "Starting ID",
        value="CJ1000",
        help="Output will start from +1. Example CJ1000 → CJ1001"
    )

    remove_invalid = st.toggle("Remove invalid phone numbers", value=True)
    dedupe = st.toggle("Remove duplicate phone numbers", value=True)
    prefix_6 = st.toggle("Prefix 6 in CSV export", value=True)

    chunk_size = st.number_input(
        "CSV chunk size",
        min_value=1000,
        max_value=500000,
        value=50000,
        step=1000
    )

if uploaded_files:
    st.subheader("Preview")

    selected_preview_file = st.selectbox(
        "Choose file to preview",
        uploaded_files,
        format_func=lambda x: x.name
    )

    try:
        preview_df = read_input_file(selected_preview_file)
        preview_df = standardize_columns(preview_df)
        preview_df = clean_text_columns(preview_df)

        col1, col2 = st.columns(2)

        with col1:
            st.metric("Preview Rows", len(preview_df))

        with col2:
            st.metric("Detected Columns", len(preview_df.columns))

        with st.expander("Detected columns", expanded=False):
            st.write(list(preview_df.columns))

        st.dataframe(preview_df.head(30), use_container_width=True)

    except Exception as e:
        st.error(f"Preview error: {e}")

    st.divider()

    if st.button("Process Files", type="primary", use_container_width=True):
        try:
            with st.spinner("Processing files..."):
                zip_bytes, summary_df = process_files(
                    uploaded_files=uploaded_files,
                    start_id=start_id,
                    remove_invalid=remove_invalid,
                    dedupe=dedupe,
                    prefix_6=prefix_6,
                    chunk_size=int(chunk_size)
                )

            st.success("Processing completed.")

            st.subheader("Summary Statistics")
            st.dataframe(summary_df, use_container_width=True)

            zip_name = (
                uploaded_files[0].name.rsplit(".", 1)[0] + ".zip"
                if len(uploaded_files) == 1
                else "OUTPUT.zip"
            )

            st.download_button(
                label="Download ZIP",
                data=zip_bytes,
                file_name=zip_name,
                mime="application/zip",
                use_container_width=True
            )

        except Exception as e:
            st.error(str(e))

else:
    st.info("Upload one or more Excel/CSV files to begin.")
