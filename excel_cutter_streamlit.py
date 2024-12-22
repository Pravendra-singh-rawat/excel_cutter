import streamlit as st
import pandas as pd
import os
import tempfile
import shutil
import datetime

@st.cache_data
def process_excel(file):
    # Read Excel file
    excel_data = pd.ExcelFile(file)
    sheet_names = excel_data.sheet_names
    
    # Collect data from all sheets
    sheets_data = {sheet: excel_data.parse(sheet) for sheet in sheet_names}
    return sheets_data, sheet_names

@st.cache_data
def save_files_by_column(sheets_data, column):
    unique_values = set()
    for df in sheets_data.values():
        if column in df.columns:
            unique_values.update(df[column].dropna().unique())

    output_files = {}
    output_dir = tempfile.mkdtemp()
    summary = []

    for value in unique_values:
        file_path = os.path.join(output_dir, f"{value}.xlsx")
        with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
            value_summary = []

            for sheet_name, df in sheets_data.items():
                if column in df.columns:
                    filtered_df = df[df[column] == value]
                    if not filtered_df.empty:
                        # Convert datetime columns to date format
                        for col in filtered_df.select_dtypes(include=['datetime']).columns:
                            filtered_df[col] = filtered_df[col].dt.date

                        filtered_df.to_excel(writer, index=False, sheet_name=sheet_name)

                        # Autofit column width
                        worksheet = writer.sheets[sheet_name]
                        for idx, col in enumerate(filtered_df.columns):
                            max_len = max(
                                filtered_df[col].astype(str).map(len).max(), len(col)
                            ) + 2
                            worksheet.set_column(idx, idx, max_len)

                        # Add to value summary
                        value_summary.append({"Sheet": sheet_name, "Rows": len(filtered_df)})

            # Add to overall summary
            summary.append({"Value": value, "Details": value_summary, "File": file_path})

        output_files[value] = file_path

    # Create a ZIP file for all generated files
    zip_file = os.path.join(output_dir, "filtered_files.zip")
    with tempfile.TemporaryDirectory() as temp_zip_dir:
        for file_path in output_files.values():
            shutil.copy(file_path, temp_zip_dir)
        shutil.make_archive(zip_file.replace(".zip", ""), 'zip', temp_zip_dir)

    return output_files, zip_file, summary

# Streamlit UI
st.title("Excel File Generator by Column Values")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    st.success("File uploaded successfully!")

    # Process Excel file
    sheets_data, sheet_names = process_excel(uploaded_file)

    # Allow user to select column to process
    all_columns = set()
    for df in sheets_data.values():
        all_columns.update(df.columns)

    selected_column = st.selectbox("Select the column to split files by:", sorted(all_columns))

    if st.button("Generate Files"):
        output_files, zip_file, summary = save_files_by_column(sheets_data, selected_column)

        # Display summary as a table
        st.write("### File Generation Summary:")
        for item in summary:
            st.write(f"#### Value: {item['Value']}")
            value_summary_df = pd.DataFrame(item['Details'])
            st.write(value_summary_df)

            with open(item['File'], "rb") as file:
                st.download_button(
                    label=f"Download {item['Value']}.xlsx",
                    data=file,
                    file_name=os.path.basename(item['File'])
                )

        # Option to download all files as a ZIP
        with open(zip_file, "rb") as zip_data:
            st.download_button(
                label="Download All Files as ZIP",
                data=zip_data,
                file_name="filtered_files.zip"
            )
    else:
        st.info("Click 'Generate Files' to process and download.")
