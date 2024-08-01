import streamlit as st
import PyPDF2
import pandas as pd
import re
import os
import time 
from io import BytesIO

st.title("Insurance Details Fetch from PDF")

raw_data=[]
emp_code = []

pdf_path = st.text_input("Enter the folder path where all insurance PDFs are present")

if pdf_path:
    # List PDF files in the provided directory
    files = [file for file in os.listdir(pdf_path) if file.endswith('.pdf')]
    st.write(len(files), "PDFs are available at the mentioned path.")

    # Proceed button
    if st.button("Proceed"):

        with st.spinner("Fetch is_in progress.."):
            time.sleep(5)

            files_paths = [os.path.join(pdf_path,file_path) for file_path in files]
            
            for i in files_paths:
                password = re.findall(r"\\(.*?)\_",i) 
                emp_code.append("".join(password))
                
                with open(i, 'rb') as file:  
                    reader = PyPDF2.PdfReader(file)
                    
                    if reader.is_encrypted:
                        try:
                            reader.decrypt("".join(password))
                        except Exception as e:
                            print(f"Decryption failed: {e}")
                            raise  
                            
                    text = ""
                    for page_num in range(len(reader.pages)):
                        try:
                            page = reader.pages[page_num]
                            text += page.extract_text() or ""
                        except Exception as e:
                            print(f"Failed to extract text from page {page_num}: {e}")

                    pattern = re.compile(r'([^:\n]+?)\s*[:\n]+\s*(.*?)(?=\n[A-Za-z]|$)', re.DOTALL)
                    matches = pattern.findall(text)
                    result_dict = {key.strip(): value.strip() for key, value in matches}
                    raw_data.append(result_dict)
            df = pd.DataFrame(raw_data)
            df["emp_code"]=emp_code

            @st.cache_data
            def convert_df_to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='all pdf fetched report')
                processed_data = output.getvalue()
                return processed_data

            excel_data = convert_df_to_excel(df)
            st.download_button(
                label="Download report",
                data=excel_data,
                file_name="insurance_deatils.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        st.success("Done! ")    
else:
    st.write("Please enter a valid folder path.")





        

