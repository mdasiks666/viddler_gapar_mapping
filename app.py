import streamlit as st
import pandas as pd
import re
import os
from io import BytesIO
import zipfile

if 'viddler_map' not in st.session_state:
    st.session_state.viddler_map = {}

def load_excel_file(excel_file):
    try:
        xls = pd.ExcelFile(excel_file)
        for sheet_name in xls.sheet_names:
            data = pd.read_excel(excel_file, sheet_name=sheet_name)
            if "ViddlerMediaId" in data.columns and "GasparMediaId" in data.columns:
                ViddlerMediaId = data["ViddlerMediaId"].astype(str)
                GasparMediaId = data["GasparMediaId"].astype(str)
                for v, g in zip(ViddlerMediaId, GasparMediaId):
                    st.session_state.viddler_map[v] = g
        if not st.session_state.viddler_map:
            st.warning("ViddlerMediaId or GasparMediaId missing in given Excel file.")
            return False
        return True
    except Exception as e:
        st.error(f"Error loading mapping file: {e}")
        return False


def process_files(files):
    output_zip = BytesIO()
    with zipfile.ZipFile(output_zip, "w") as zipf:
        for file in files:
            try:
                if os.path.splitext(file.name)[1] == ".xlsx":
                    data = pd.read_excel(file)
                    idf = pd.DataFrame(data).astype(str)

                    vid = [
                        item
                        for item in idf.columns.tolist()
                        if "viddler" in str(item).strip().lower()
                        or str(item).strip().lower() == "id"
                    ]
                    if len(vid):
                        idf["Gaspar ID"] = idf[vid[0]].apply(update_gaspar_id)
                        idf["EMBED Updated"] = idf["Gaspar ID"].apply(
                            generate_embed_code
                        )

                    if "Media ID" in idf.columns:
                        idf["Media ID Updated"] = idf["Media ID"].apply(update_media_id)

                    columns = idf.columns.tolist()
                    if len(vid):
                        columns.remove("Gaspar ID")
                        b_index = columns.index(vid[0])
                        columns.insert(b_index + 1, "Gaspar ID")

                    if "Media ID" in columns:
                        columns.remove("Media ID Updated")
                        b_index = columns.index("Media ID")
                        columns.insert(b_index + 1, "Media ID Updated")

                    if "EMBED" in columns:
                        columns.remove("EMBED Updated")
                        b_index = columns.index("EMBED")
                        columns.insert(b_index + 1, "EMBED Updated")

                    idf = idf[columns]
                    output_buffer = BytesIO()
                    idf.to_excel(output_buffer, index=False)
                    zipf.writestr(f"{file.name}_updated.xlsx", output_buffer.getvalue())

            except Exception as e:
                st.error(f"Error processing {file.name}: {e}")

    output_zip.seek(0)
    return output_zip


def update_gaspar_id(viddler_id):
    if viddler_id in st.session_state.viddler_map:
        return st.session_state.viddler_map[viddler_id]
    return "ID Not Found"


def generate_embed_code(gaspar_id):
    if gaspar_id != "ID Not Found":
        return f"""<iframe frameborder='0' style='width:640px; height:480px;' src='https://media.gaspar.mheducation.com/GASPARPlayer/play.html?id={gaspar_id}' allowfullscreen></iframe>"""
    return "ID Not Found"


def update_media_id(media_id):
    match = re.search(r"_viddler_([a-z0-9]+)_", media_id)
    if match:
        viddler_id = match.group(1)
        if viddler_id in st.session_state.viddler_map:
            return re.sub(
                viddler_id,
                st.session_state.viddler_map.get(viddler_id),
                media_id,
            ).replace("viddler", "gaspar")
    return "ID Not Found"


hide_streamlit_style = """
    <style>
        .stAppHeader {
            display: none;
        }     
        .stMainBlockContainer {
            padding: 50px;
        }
        ._terminalButton_rix23_138{
            display: none !important;
        }
    </style>
"""

st.set_page_config(page_title="Viddler to Gaspar Mapping", layout="wide")
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

st.title("Viddler to Gaspar Mapping Tool")
st.write("Upload a mapping Excel file and input directory to begin processing files.")

if "mapping_loaded" not in st.session_state:
    st.session_state.mapping_loaded = False
if "processing" not in st.session_state:
    st.session_state.processing = False
if "output_ready" not in st.session_state:
    st.session_state.output_ready = False

with st.container():
    st.subheader("📁 1. Upload Mapping File")
    excel_file = st.file_uploader("Select Mapping XLSX", type=["xlsx"], key="mapping_uploader")

    if excel_file and not st.session_state.mapping_loaded:
        with st.spinner("⏳ Processing mapping file..."):
            if load_excel_file(excel_file):
                st.session_state.mapping_loaded = True
                st.success(f"Mapping loaded: {excel_file.name}", icon="✅")

if st.session_state.mapping_loaded:
    with st.container():
        st.subheader("📂 2. Upload Input Files")
        input_files = st.file_uploader(
            "Select Input Files (.xlsx)",
            type=["xlsx"],
            accept_multiple_files=True,
            key="input_files_uploader"
        )
        if input_files:
            st.success(f"{len(input_files)} Files Uploaded Successfully", icon="✅")

if st.session_state.mapping_loaded and input_files:
    st.subheader("🚀 3. Process Files")
    
    if not st.session_state.output_ready:
        if st.button("Replace IDs and Generate Output", disabled=st.session_state.processing):
            st.session_state.processing = True
            
            with st.spinner("⏳ Processing files... This may take a while..."):
                output_zip = process_files(input_files)
                st.session_state.output_zip = output_zip
                st.session_state.output_ready = True
                st.session_state.processing = False
                
            st.success("Files processed successfully!", icon="✅")

    if st.session_state.output_ready:
        st.download_button(
            label="📥 Download Updated Files",
            data=st.session_state.output_zip,
            file_name="updated_files.zip",
            mime="application/zip",
        )