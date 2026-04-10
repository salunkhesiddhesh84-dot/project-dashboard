import streamlit as st
import pandas as pd
import os
from dashboard_script import main

st.set_page_config(page_title="Project Dashboard", layout="centered")

st.title("📊 Project Dashboard Generator")

st.write("Upload your raw Excel file to generate dashboard report.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    # Save uploaded file
    input_path = "input.xlsx"
    output_path = "output.xlsx"

    with open(input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success("File uploaded successfully!")

    if st.button("Generate Report"):
        try:
            main(input_path, output_path)

            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Download Report",
                    data=f,
                    file_name="Project_Dashboard.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.success("Report generated successfully!")

        except Exception as e:
            st.error(f"Error: {str(e)}")
