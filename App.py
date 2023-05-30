import streamlit as st
import pandas as pd
import base64


# Define the instruments and their corresponding column names
instrument_columns = {
    'Density Meter': ['Sample ID', 'Density (g/mL)', 'Temperature °C'],
    'LECO CHN': ['Sample ID', 'Mass (g)', 'C%', 'O% (diff)'],
    # Add more instruments and their columns as needed here
}


def generate_excel_file(instrument):
    # Logic to generate the Excel file with pre-formatted columns and equations
    # You can use libraries like Pandas or openpyxl to work with Excel files

    # Example: Create a DataFrame with the specified columns
    columns = instrument_columns[instrument]
    df = pd.DataFrame(columns=columns)

    # Add three cells under the column header "O% (diff)" if the instrument is "LECO CHN"
    if instrument == "LECO CHN":
        extra_cells = ["=100-SUM(C6:C6)"] * 1  # Create cells with the desired formula
        row_data = [""] * (len(columns) - 1) + extra_cells
        df.loc[-1] = row_data
        df.index = df.index + 1  # Shifting the index to insert the new row at the top
        df = df.sort_index()  # Sorting the index to maintain the order

    # Save the DataFrame to an Excel file
    filename = f'{instrument}_data_template.xlsx'
    df.to_excel(filename, index=False)

    return filename


# Main app
def main():
    st.title('Instrument Data App')

    # Display a dropdown to select the instrument
    instrument = st.selectbox('Select an instrument', list(instrument_columns.keys()))

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
