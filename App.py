import os
target_directory = "FirstApp"
base_directory = os.getcwd()
os.chdir(os.path.join(base_directory, target_directory))
print(os.getcwd())

#set target directory to web-based
#saves output to GitHub Working Directory??
#import os
#target_directory = "CosetteTurvold/FirstApp"
#base_directory = os.getcwd()
#os.chdir(os.path.join(base_directory, target_directory))
#print(os.getcwd())


import streamlit as st

import pandas as pd
import numpy as np


# Define the instruments and their corresponding column names
instruments = {
    'Instrument 1': ['Column A', 'Column B', 'Column C'],
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
    df.to_excel(f'{instrument}_data.xlsx', index=False)

# Main app
def main():
    st.title('Instrument Data App')

    # Display a dropdown to select the instrument
    instrument = st.selectbox('Select an instrument', list(instruments.keys()))

# Generate the Excel file when a button is clicked
    if st.button('Generate Excel'):
        generate_excel_file(instrument)
        st.success(f'Excel file for {instrument} has been generated!')

if __name__ == '__main__':
    main()
    
