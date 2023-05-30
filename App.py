#!/usr/bin/env python3
# -*- coding: utf-8 -*-

#set working directory when NOT working web-based 
#saves output to local working directory
#import os
#target_directory = "/app/firstapp"
#base_directory = os.getcwd()
#os.chdir(os.path.join(base_directory, target_directory))
#print(os.getcwd())

#set target directory to web-based
#saves output to GitHub Working Directory??
#import os
#target_directory = "CosetteTurvold/FirstApp"
#base_directory = os.getcwd()
#os.chdir(os.path.join(base_directory, target_directory))
#print(os.getcwd())


import streamlit as st
import pandas as pd
import base64


# Define the instruments and their corresponding column names
instruments = {
    'Densimeter': ['Sample ID', 'Density (g/mL)', 'Temperature °C'],
    'Instrument 2': ['Column X', 'Column Y', 'Column Z'],
    # Add more instruments and their columns as needed
}

def generate_excel_file(instrument):
    # Logic to generate the Excel file with pre-formatted columns and equations
    # You can use libraries like Pandas or openpyxl to work with Excel files

    # Example: Create a DataFrame with the specified columns
    columns = instruments[instrument]
    df = pd.DataFrame(columns=columns)

    # Save the DataFrame to an Excel file
    filename = f'{instrument}_data.xlsx'
    df.to_excel(filename, index=False)

    return filename

# Main app
def main():
    st.title('Instrument Data App')

    # Display a dropdown to select the instrument
    instrument = st.selectbox('Select an instrument', list(instruments.keys()))

    # Generate the Excel file when a button is clicked
    if st.button('Generate Excel'):
        filename = generate_excel_file(instrument)
        st.success(f'Excel file for {instrument} has been generated!')

        # Provide a download link to the file
        with open(filename, 'rb') as f:
            data = f.read()
        b64 = base64.b64encode(data).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">Download Excel File</a>'
        st.markdown(href, unsafe_allow_html=True)

if __name__ == '__main__':
    main()
