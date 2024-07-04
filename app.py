import streamlit as st
from dotenv import load_dotenv
import pandas as pd
import os
from langchain_groq import ChatGroq
from pandasai import SmartDataframe
from pandasai.llm import LLM
from io import BytesIO
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows

load_dotenv()

llm = ChatGroq(
    model_name="mixtral-8x7b-32768", 
    api_key=os.getenv("GROQ_API_KEY")
)

def process_query(df, prompt):
    smart_df = SmartDataframe(df, config={"llm": llm})
    
    response = smart_df.chat(prompt)
    
    # Check if a plot was generated
    if plt.get_fignums():
        # Get the last generated plot
        fig = plt.gcf()
        
        # Save the plot to a BytesIO object
        buf = BytesIO()
        fig.savefig(buf, format="png")
        buf.seek(0)
        plt.close(fig)  
        return buf, "plot", df
    else:
        return response, "text", None

def create_excel_with_plot_and_data(plot_bytes, df):
    wb = Workbook()
    ws = wb.active

    img = Image(plot_bytes)


    ws.add_image(img, 'A1')
    data_start_row = img.height // 20 + 2  

    # Add the dataframe to the worksheet
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx + data_start_row, column=c_idx, value=value)


    excel_bytes = BytesIO()
    wb.save(excel_bytes)
    excel_bytes.seek(0)

    return excel_bytes

st.title("Excel Data Analysis and Visualization App")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    # Display the dataframe
    st.write("Data Preview:")
    st.dataframe(df.head())
    
    # Text input for the query
    prompt = st.text_input("Enter your query or plot prompt:", "")
    
    if st.button("Generate Response"):
        if prompt:
            # Process the query
            result, result_type, data = process_query(df, prompt)
            
            if result_type == "plot":
                # Display the plot
                st.image(result, caption="Generated Plot", use_column_width=True)
                
                # Create a download button for the Excel file
                excel_bytes = create_excel_with_plot_and_data(result, data)
                st.download_button(
                    label="Download Plot and Data as Excel",
                    data=excel_bytes,
                    file_name="plot_and_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                # Display the text response
                st.write("Response:")
                st.write(result)
        else:
            st.write("Please enter a query or plot prompt.")
            
#("Note: This app uses the Groq API. Make sure you have set the GROQ_API_KEY environment variable.")