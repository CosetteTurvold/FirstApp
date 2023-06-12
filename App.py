import streamlit as st
import pandas as pd
import base64


# Define the instruments and their corresponding column names
instrument_columns = {
    'Density Meter': ['Sample ID', 'Density (g/mL)', 'Temperature °C'],
    'LECO CHN': ['Sample ID', 'Mass (g)', 'C%', 'H%', 'N%','O% (diff)'],
    'Karl Fischer': ['Sample ID','~Vol(uL)', 'Mass(g)', 'Titrant(mL)', 'H2O%'],
    'KF & LECO CHN Combined': ['Sample ID', 'Mass (g)', 'C%', 'H%', 'N%','O% (diff)', 'Water', 'C% Dry Basis', 'H% Dry Basis', 'O% Dry Basis'],
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
        CHN_calc = [
            ['Sample 1','','','','', "=100-SUM(C2:E2)"],
            ['Sample 1','','','','', "=100-SUM(C3:E3)"],
            ['Sample 1','','','','', "=100-SUM(C4:E4)"], 
            ['', 'Average', '=AVERAGE(C2:C4)','=AVERAGE(D2:D4)','=AVERAGE(E2:E4)','=AVERAGE(F2:F4)'],
            ['', 'StDev', '=STDEV(C2:C4)','=STDEV(D2:D4)','=STDEV(E2:E4)','=STDEV(F2:F4)'],
            ['', 'RSD', '=C6/C5*100','=D6/D5*100','=E6/E5*100','=F6/F5*100'],
        ]
        for row in CHN_calc:
            df.loc[len(df)] = row
            
    elif instrument == "Karl Fischer":
            kf_calc = [
                ['Sample 1','', '', '', ''],
                ['Sample 1','', '', '', ''],
                ['Sample 1','', '', '', ''],
                ['', '', '','Mean%', '=AVERAGE(E2:E4)'],
                ['', '', '', 'Sabs%', '=STDEV(E2:E4)'],
                ['', '', '', 'Srel%', '=(D6/D5)*100'],
            ]
            for row in kf_calc:
                df.loc[len(df)] = row
                
    elif instrument == "KF & LECO CHN Combined":
        CHN_KF_calc = [
            ['Sample 1','','','','', "=100-SUM(C2:E2)",'','','',''],
            ['Sample 1','','','','', "=100-SUM(C3:E3)",'','','',''],
            ['Sample 1','','','','', "=100-SUM(C4:E4)",'','','',''], 
            ['', 'Average', '=AVERAGE(C2:C4)','=AVERAGE(D2:D4)','=AVERAGE(E2:E4)','=AVERAGE(F2:F4)', '=AVERAGE(G2:G4)', '=C5/(100-G5)*100','=(D5-(G5*0.111))/(100-G5)*100','=(F5-(0.889*G5))/(100-G5)*100'],
            ['', 'StDev', '=STDEV(C2:C4)','=STDEV(D2:D4)','=STDEV(E2:E4)','=STDEV(F2:F4)','=STDEV(G2:G4)','','',''],
            ['', 'RSD', '=C6/C5*100','=D6/D5*100','=E6/E5*100','=F6/F5*100','=G6/G5*100','','',''],
        ]
        for row in CHN_KF_calc:
            df.loc[len(df)] = row
    

    # Create an Excel writer and specify the sheet name
    filename = f'{instrument}_data_template.xlsx'
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    
    # Define a format for bold cells
    bold_format = writer.book.add_format({'bold': True})
    
    # Write the DataFrame to the default sheet
    df.to_excel(writer, index=False, sheet_name='Analysis', engine= 'xlsxwriter')
    
    # Create a additional sheet and write additional data
    if instrument == "LECO CHN":
        extra_sheet = {
            'Data': ['Additional Data 1', 'Additional Data 2', 'Additional Data 3']
            }
        additional_df = pd.DataFrame(extra_sheet)
        additional_df.to_excel(writer, index=False, sheet_name='Cresol Testing', startrow=1, startcol=0)

   # Get the worksheet for the Cresol Testing sheet
        worksheet = writer.sheets['Cresol Testing']
        
        # additional method apply bold formatting to the cell containing "Additional Data 1" (A1)
        #worksheet.write('A1', 'Additional Data 1', bold_format)
        #worksheet.write('A2', 'Additional Data 2', bold_format)

    # Write new column and row headers in Cresol Testing sheet
        headers = ['Cresol Limit Testing', 'Cresol Measured', 'Average', 'Low-value', 'High-value']
        for i, header in enumerate(headers):
            worksheet.write(0, i, header, bold_format)
        
        # Write new row headers in Cresol Testing sheet
        row_headers = ['Carbon', 'Hydrogen', 'Nitrogen', 'Oxygen']
        for i, header in enumerate(row_headers):
            worksheet.write(i+1, 0, header, bold_format)
            

  # Write Cresol Limits below the columns
        cresol_limits = [
            ['','', 77.540064516129, 76.8689132062322, 78.2112158260259], 
            ['','', 7.68705, 7.41272698524775, 7.96137301475225],
            ['','', 0.0579167768595041, 0.0151682547630326, 0.100665298955976],
            ['','', 14.772885483871, 13.9828215158141, 15.5629494519279]
             ]
            
        for row_num, row_data in enumerate(cresol_limits):
            for col_num, cell_data in enumerate(row_data):
                worksheet.write(row_num + 1, col_num + 0, cell_data)
                
    elif instrument == "Karl Fischer":
        extra_sheet = {
           'Water Standard Check': ['Water Standard (1)', 'Water Standard (2)', 'Water Standard (3)']
            }
        additional_df = pd.DataFrame(extra_sheet)
        additional_df.to_excel(writer, index=False, sheet_name='Water Standard', startrow=0, startcol=0)
        worksheet = writer.sheets['Water Standard']
        headers = ['Water Standard Check', 'Mass(g)', 'Titrant(mL)', 'H2O%']
        for i, header in enumerate(headers):
            worksheet.write(0, i, header, bold_format)
        
        # Write new row headers in Cresol Testing sheet
        row_headers = ['Water Standard (1)', 'Water Standard (2)', 'Water Standard (3)']
        for i, header in enumerate(row_headers):
            worksheet.write(i+1, 0, header, bold_format)
        
        KF_water_check = [
           ['', '', 'Mean%', '=AVERAGE(D2:D4)'],
           ['', '', 'Sabs%', '=STDEV(D2:D4)'],
           ['', '', 'Srel%', '=(D6/D5)*100'],
           ]
        for row_num, row_data in enumerate(KF_water_check):
            for col_num, cell_data in enumerate(row_data):
                worksheet.write(row_num + 4, col_num + 0, cell_data)
        

    # Close the writer and save the Excel file
    writer.close()

    return filename



# Main app
def main():
    st.title('Instrument Template Generator')

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
