import streamlit as st
import pandas as pd
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
import tempfile
import openpyxl
import os

# Configure Tesseract path
if os.name == 'nt':  # Windows
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
else:  # Linux
    pytesseract.pytesseract.tesseract_cmd = "/usr/bin/tesseract"

# Function to load Excel data
@st.cache_data
def load_data(file_path):
    df = pd.read_excel(file_path)
    return df

# Function to save data back to Excel
def save_data(data, file_path):
    data.to_excel(file_path, index=False)

# Function to clean data
def clean_data(df):
    df = df.fillna('NA').astype(str)
    return df

# OCR function to extract text from image
def extract_text_from_image(image):
    try:
        text = pytesseract.image_to_string(image)
        return text
    except pytesseract.TesseractNotFoundError:
        st.error("Tesseract OCR not found. Please ensure it is installed and the path is correctly set.")
        return ""

# Function to extract text from PDF
def extract_text_from_pdf(pdf_path):
    text = ""
    images = convert_from_path(pdf_path)
    for image in images:
        text += extract_text_from_image(image) + "\n"
    return text

# Function to map extracted data to DataFrame columns
def map_data_to_columns(extracted_text, df_columns):
    lines = extracted_text.split('\n')
    data_dict = {col: [] for col in df_columns}
    
    for line in lines:
        if line.strip():
            values = line.split(',')
            for i, col in enumerate(df_columns):
                if i < len(values):
                    data_dict[col].append(values[i].strip())
    
    new_data_df = pd.DataFrame(data_dict)
    return new_data_df

def match_and_fill_data(extracted_df, original_df):
    # Example matching logic: Assuming the first column is a unique identifier
    for index, row in extracted_df.iterrows():
        unique_id = row[original_df.columns[0]]
        if unique_id in original_df[original_df.columns[0]].values:
            original_df.loc[original_df[original_df.columns[0]] == unique_id] = row
        else:
            original_df = original_df.append(row, ignore_index=True)
    return original_df

def main():
    st.title("Excel Data Management")

    st.sidebar.title('Data Entry')
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        if 'original_file_path' not in st.session_state:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
                temp_file.write(uploaded_file.getbuffer())
                st.session_state.original_file_path = temp_file.name

        if 'df' not in st.session_state:
            df = load_data(st.session_state.original_file_path)
            df = clean_data(df)
            st.session_state.df = df
        else:
            df = st.session_state.df

        st.write('Current Data:')
        st.write(st.session_state.df)

        st.sidebar.header('Upload Scanned Document')
        scanned_file = st.sidebar.file_uploader("Choose a scanned document (image or PDF)", type=["png", "jpg", "jpeg", "pdf"])

        if scanned_file is not None:
            extracted_text = ""
            if scanned_file.type == "application/pdf":
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                    temp_pdf.write(scanned_file.getbuffer())
                    temp_pdf_path = temp_pdf.name
                extracted_text = extract_text_from_pdf(temp_pdf_path)
            else:
                image = Image.open(scanned_file)
                extracted_text = extract_text_from_image(image)

            if extracted_text:
                st.text("Extracted Text:")
                st.write(extracted_text)

                new_data_df = map_data_to_columns(extracted_text, st.session_state.df.columns)
                st.write('New Data Extracted from Scanned Document:')
                st.write(new_data_df)

                if st.sidebar.button('Match and Fill Data'):
                    st.session_state.df = match_and_fill_data(new_data_df, st.session_state.df)
                    st.session_state.df = clean_data(st.session_state.df)
                    st.sidebar.success('Data matched and filled successfully!')
                    st.experimental_rerun()

        if st.button('Download Updated Data'):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as updated_file:
                save_data(st.session_state.df, updated_file.name)
                with open(updated_file.name, "rb") as file:
                    st.download_button(
                        label="Download Excel file",
                        data=file,
                        file_name="updated_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        st.header('Retrieve Data')
        filter_cols = st.multiselect('Select columns for filter:', options=df.columns)
        
        filter_values = {}
        for col in filter_cols:
            filter_values[col] = st.text_input(f'Enter value to filter {col}:')

        if filter_values:
            filtered_df = df.copy()
            for col, value in filter_values.items():
                if value:
                    filtered_df = filtered_df[filtered_df[col].str.lower() == value.lower()]
            st.write(filtered_df)

if __name__ == "__main__":
    main()
