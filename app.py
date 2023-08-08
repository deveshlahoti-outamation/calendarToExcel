import glob
import zipfile

import streamlit as st
import os
import shutil
from main import main, clean_files

SAVE_PATH = 'files/'


def save_uploaded_file(uploaded_file):
    # Check if the save path exists, if not, create it
    if not os.path.exists(SAVE_PATH):
        os.makedirs(SAVE_PATH)

    # Get the file from the upload widget
    for file in uploaded_file:
        with st.spinner(f'Saving {file.name}...'):
            # Create a temporary file, then move it to the desired directory
            with open(file.name, 'wb') as f:
                f.write(file.getvalue())
            shutil.move(file.name, os.path.join(SAVE_PATH, file.name))


def download_files():
    output_files = glob.glob(os.path.join('output/files', '*.xlsx'))

    # Create a Zip file
    with zipfile.ZipFile('output/files/files.zip', 'w') as zipf:
        for file in output_files:
            zipf.write(file, arcname=os.path.basename(file))  # arcname is to avoid storing the folder structure

    # Create a download button for the Zip file
    with open('output/files/files.zip', 'rb') as f:
        zip_bytes = f.read()
        st.download_button(
            label="Download",
            data=zip_bytes,
            file_name='files.zip',
            mime='application/zip'
        )


st.set_page_config(page_title="Calendar Converter", page_icon=":bookmark_tabs:", layout="centered")

st.markdown("<h1 style='text-align: center; color: blue;'>Calendar to Excel</h1>", unsafe_allow_html=True)
st.markdown("<h2 style='text-align: center; color: gray;'>Drop your file(s) here 👇</h2>", unsafe_allow_html=True)

files = st.file_uploader('', ["pdf"], accept_multiple_files=True)

if len(files) != 0:
    save_uploaded_file(files)
    main()
    with st.columns(6)[5]:
        download_files()
    clean_files()


