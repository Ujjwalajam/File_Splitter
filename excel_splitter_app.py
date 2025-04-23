import streamlit as st
import pandas as pd
import io
import zipfile

st.title("📊 Smart Excel Splitter Utility")

# Step 1: Upload file
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

# Step 2: Choose splitting method
split_mode = st.radio("Choose split method:", ["Split by Rows", "Split by Column Value"])

# Step 3: Common input
base_name = st.text_input("Enter base name for ZIP and files (e.g., Himalaya_Urban_Retailer):", value="Enter File Name")

# Variables to be used later
rows_per_file = None
split_column = None
df = None

# Step 4: Read file and show appropriate input
if uploaded_file:
    df = pd.read_excel(uploaded_file)

    if split_mode == "Split by Rows":
        rows_per_file = st.number_input("Split into how many rows per file?", min_value=1, value=1, step=1)
    else:  # Split by column value
        column_options = df.columns.tolist()
        split_column = st.selectbox("Select column to split by:", column_options)

# Step 5: Button to process
if st.button("🚀 Split Excel File"):
    if not uploaded_file or not df is not None:
        st.error("❌ Please upload a valid Excel file.")
    elif not base_name.strip():
        st.error("❌ Please provide a valid base name.")
    else:
        try:
            # Create in-memory ZIP
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:

                if split_mode == "Split by Rows":
                    total_rows = len(df)
                    num_chunks = (total_rows + rows_per_file - 1) // rows_per_file

                    for i in range(num_chunks):
                        start_row = i * rows_per_file
                        end_row = min((i + 1) * rows_per_file, total_rows)
                        chunk = df.iloc[start_row:end_row]

                        buffer = io.BytesIO()
                        chunk.to_excel(buffer, index=False, engine='openpyxl')
                        buffer.seek(0)

                        zip_file.writestr(f"{base_name}_{i+1}.xlsx", buffer.read())

                else:  # Split by column value
                    unique_values = df[split_column].dropna().unique()
                    for val in unique_values:
                        chunk = df[df[split_column] == val]
                        buffer = io.BytesIO()
                        chunk.to_excel(buffer, index=False, engine='openpyxl')
                        buffer.seek(0)

                        sanitized_val = str(val).replace(" ", "_")
                        zip_file.writestr(f"{base_name}_{sanitized_val}.xlsx", buffer.read())

            zip_buffer.seek(0)
            st.success("✅ Files successfully split and zipped!")

            # Download button
            st.download_button(
                label="📦 Download ZIP File",
                data=zip_buffer,
                file_name=f"{base_name}_1.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"⚠️ Error occurred: {e}")
