import streamlit as st
import pandas as pd
import base64
import xlsxwriter


# Define the instruments and their corresponding column names
instrument_columns = {
    'Density Meter': ['Sample ID', 'Density (g/mL)', 'Temperature (°C)'],
    'LECO CHN': ['Sample ID', 'Mass (g)', 'C%', 'H%', 'N%','O% (diff)'],
    'Karl Fischer': ['Sample ID','~Vol(μL)', 'Mass(g)', 'Titrant(mL)', 'H2O%'],
    'KF & LECO CHN Combined': ['Sample ID', 'Mass (g)', 'C%', 'H%', 'N%','O% (diff)', 'Water', 'C% Dry Basis', 'H% Dry Basis', 'O% Dry Basis'],
    'Viscometer': ['Sample ID', 'Viscosity (cP)','Torque (%)','Speed (rpm)', 'Temperature (°C)'],
    'Acids Titration': ['Sample ID', 'CAN mol/kg', 'TAN mol/kg', 'PhAN mol/kg', 'Carboxylic Acid Number mg KOH/g','Total Acid Number mg KOH/g', 'Phenolic Acid Number mg KOH/g']
}


def generate_excel_file(instrument):
    # Logic to generate the Excel file with pre-formatted columns and equations
    # You can use libraries like Pandas or openpyxl to work with Excel files
    # Example: Create a DataFrame with the specified columns
    columns = instrument_columns[instrument]
    df = pd.DataFrame(columns=columns)
    
    filename = f'{instrument}_data_template.xlsx'
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    
    # Write the DataFrame to the default sheet
    df.to_excel(writer, index=False, sheet_name=('Analysis'), engine='xlsxwriter')
    worksheet = writer.sheets['Analysis']
     
    # Get the workbook from the writer
    workbook = writer.book
    
    # Define a format for bold, borders, and green conditional format
    bold_format = writer.book.add_format({'bold': True})
    bottomborder_format = workbook.add_format({'bottom':2, 'border_color':'black'})
    rightborder_format = workbook.add_format({'right':2, 'border_color':'black'})
    headerborder_format = workbook.add_format({'border':2, 'border_color':'black'})
    cornerborder_format = workbook.add_format({'right':2, 'bottom':2, 'border_color':'black'})                                   
    dashedborder_format = workbook.add_format({'border_color': 'black'})
    dashedborder_format.set_bottom(6)
    sideborder_format = workbook.add_format({'right':2,'border_color':'black'})
    sideborder_format.set_bottom(6)
    green_format = workbook.add_format({'bg_color':   '#C6EFCE',
                               'font_color': '#006100'})
    red_format = workbook.add_format({'bg_color':   '#FFC7CE',
                               'font_color': '#9C0006'})
#start creating different excel files for each instrument option
    if instrument == "LECO CHN":
        CHN_calc = [
            ['Sample 1','','','','', '=100-SUM(C2:E2)'],
            ['Sample 1','','','','', '=100-SUM(C3:E3)'],
            ['Sample 1','','','','', '=100-SUM(C4:E4)'], 
            ['', 'Average', '=AVERAGE(C2:C4)','=AVERAGE(D2:D4)','=AVERAGE(E2:E4)','=AVERAGE(F2:F4)'],
            ['', 'StDev', '=STDEV(C2:C4)','=STDEV(D2:D4)','=STDEV(E2:E4)','=STDEV(F2:F4)'],
            ['', 'RSD', '=C6/C5*100','=D6/D5*100','=E6/E5*100','=F6/F5*100'],
        ]
        for row in CHN_calc:
            df.loc[len(df)] = row

        
        # Apply all borders to top row given they are blank OR not blank
        worksheet.conditional_format('A1:F1', {'type': 'no_blanks', 'format': headerborder_format})
        worksheet.conditional_format('A1:F1', {'type': 'blanks', 'format': headerborder_format})
        
        # Apply bottom borders to A7:E7 given they are blank OR not blank
        worksheet.conditional_format('A7:E7', {'type': 'no_blanks', 'format': bottomborder_format})
        worksheet.conditional_format('A7:E7', {'type': 'blanks', 'format': bottomborder_format})
        
        # Apply dashed bottom borders to A4:E4 given they are blank OR not blank
        worksheet.conditional_format('A4:E4', {'type': 'no_blanks', 'format': dashedborder_format})
        worksheet.conditional_format('A4:E4', {'type': 'blanks', 'format': dashedborder_format})
        
        # Apply right borders to F2:F3 given they are blank OR not blank
        worksheet.conditional_format('F2:F3', {'type': 'no_blanks', 'format': rightborder_format})
        worksheet.conditional_format('F2:F3', {'type': 'blanks', 'format': rightborder_format})
        
        # Apply right borders to F5:F6 given they are blank OR not blank
        worksheet.conditional_format('F5:F6', {'type': 'no_blanks', 'format': rightborder_format})
        worksheet.conditional_format('F5:F6', {'type': 'blanks', 'format': rightborder_format})
        
        # Apply right and dashed bottom border to F4 given it is blank OR not blank
        worksheet.conditional_format('F4', {'type': 'no_blanks', 'format': sideborder_format})
        worksheet.conditional_format('F4', {'type': 'blanks', 'format': sideborder_format})
        
        # Apply corner border to F7 given it is blank OR not blank
        worksheet.conditional_format('F7', {'type': 'no_blanks', 'format': cornerborder_format})
        worksheet.conditional_format('F7', {'type': 'blanks', 'format': cornerborder_format})
        
        # Create an additional sheet and write additional data
        extra_sheetCHN = {
            'Data': ['Additional Data 1', 'Additional Data 2', 'Additional Data 3']
        }
        additional_df = pd.DataFrame(extra_sheetCHN)
        additional_df.to_excel(writer, index=False, sheet_name='Cresol Testing', startrow=1, startcol=0)
        # Get the worksheet for the Cresol Testing sheet
        worksheet = writer.sheets['Cresol Testing']
        
        #set the file to open the column width to the length of a string
        Cresolcolumnwidth = len("Cresol Triplicate ")
        #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(0, 1, Cresolcolumnwidth)
        
        #set the file to open the column width to the length of a string
        Cresolcolumnwidth = len("Cresol Measured")
        #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(1, 1, Cresolcolumnwidth)
                  
        
        #set conditional formatting for Cresol values in range green/red for in/out of range
        worksheet.conditional_format('B2', {'type':     'cell',
                                       'criteria': 'between',
                                       'minimum':  77.0,
                                       'maximum':  78.3,
                                       'format':   green_format})
        worksheet.conditional_format('B2', {'type':     'cell',
                                       'criteria': 'not between',
                                       'minimum':  77.0,
                                       'maximum':  78.3,
                                       'format':   red_format})
        
        worksheet.conditional_format('B3', {'type':     'cell',
                                       'criteria': 'between',
                                       'minimum':  7.4,
                                       'maximum':  7.9,
                                       'format':   green_format})
        worksheet.conditional_format('B3', {'type':     'cell',
                                       'criteria': 'not between',
                                       'minimum':  7.4,
                                       'maximum':  7.9,
                                       'format':   red_format})
        
        worksheet.conditional_format('B5', {'type':     'cell',
                                       'criteria': 'between',
                                       'minimum':  14.0,
                                       'maximum':  15.4,
                                       'format':   green_format})
        
        worksheet.conditional_format('B5', {'type':     'cell',
                                       'criteria': 'not between',
                                       'minimum':  14.0,
                                       'maximum':  15.4,
                                       'format':   red_format})

         # Write new column headers in Cresol Testing sheet
        headers = ['Cresol Limit Testing', 'Cresol Measured', 'Average', 'Low-value', 'High-value']
        for i, header in enumerate(headers):
             worksheet.write(0, i, header, bold_format)
             
        # Write new row headers in Cresol Testing sheet
        row_headers = ['Carbon', 'Hydrogen', 'Nitrogen', 'Oxygen']
        for i, header in enumerate(row_headers):
             worksheet.write(i+1, 0, header, bold_format)
                 

       # Write Cresol Limits below the columns
             cresol_limits = [
                 ['','=C11', 77.7, 77.0, 78.3], 
                 ['','=D11', 7.5, 7.4, 7.9],
                 ['','=E11', 0.03, 0.0, 0.1],
                 ['','=F11', 14.8, 14.0, 15.4]
                  ]
             for row_num, row_data in enumerate(cresol_limits):
                 for col_num, cell_data in enumerate(row_data):
                     worksheet.write(row_num + 1, col_num + 0, cell_data)      
        #write a table to put the three cresol values from the CHN data
        cresol_triplicate = [
            ['Cresol Triplicate Check', 'Mass (g)', 'C%', 'H%', 'N%', 'O% (diff)'],
            ['Cresol 1','','','','', '=100-SUM(C8:E8)'],
            ['Cresol 2','','','','', '=100-SUM(C9:E9)'],
            ['Cresol 3','','','','', '=100-SUM(C10:E10)'], 
            ['', 'Average', '=AVERAGE(C8:C10)','=AVERAGE(D8:D10)','=AVERAGE(E8:E10)','=AVERAGE(F8:F10)'],
            ['', 'StDev', '=STDEV(C8:C10)','=STDEV(D8:D10)','=STDEV(E8:E10)','=STDEV(F8:F10)'],
            ['', 'RSD', '=C12/C11*100','=D12/D11*100','=E12/E11*100','=F12/F11*100'],
        ]
        for row_num, row_data in enumerate(cresol_triplicate):
            for col_num, cell_data in enumerate(row_data):
                worksheet.write(row_num + 6, col_num + 0, cell_data)
        
        #write borders for the cresol triplicate data
        
        # Apply all borders to top row given they are blank OR not blank
        worksheet.conditional_format('A7:F7', {'type': 'no_blanks', 'format': headerborder_format})
        worksheet.conditional_format('A7:F7', {'type': 'blanks', 'format': headerborder_format})
        
        # Apply bottom borders to A13:E13 given they are blank OR not blank
        worksheet.conditional_format('A13:E13', {'type': 'no_blanks', 'format': bottomborder_format})
        worksheet.conditional_format('A13:E13', {'type': 'blanks', 'format': bottomborder_format})
        
        # Apply dashed bottom borders to A10:E10 given they are blank OR not blank
        worksheet.conditional_format('A10:E10', {'type': 'no_blanks', 'format': dashedborder_format})
        worksheet.conditional_format('A10:E10', {'type': 'blanks', 'format': dashedborder_format})
        
        # Apply right borders to F8:F9 given they are blank OR not blank
        worksheet.conditional_format('F8:F9', {'type': 'no_blanks', 'format': rightborder_format})
        worksheet.conditional_format('F8:F9', {'type': 'blanks', 'format': rightborder_format})
        
        # Apply right borders to F11:F12 given they are blank OR not blank
        worksheet.conditional_format('F11:F12', {'type': 'no_blanks', 'format': rightborder_format})
        worksheet.conditional_format('F11:F12', {'type': 'blanks', 'format': rightborder_format})
        
        # Apply right and dashed bottom border to F10 given it is not blank
        worksheet.conditional_format('F10', {'type': 'no_blanks', 'format': sideborder_format})
        
        # Apply corner border to F13 given it is not blank
        worksheet.conditional_format('F13', {'type': 'no_blanks', 'format': cornerborder_format})
           
        # Get the worksheet for the Cresol Testing sheet
        worksheet = writer.sheets['Cresol Testing']
            
    elif instrument == "Karl Fischer":
        kf_calc = [
            ['Sample 1','', '', '', ''],
            ['Sample 1','', '', '', ''],
            ['Sample 1','', '', '', ''],
            ['', '', '','Mean%', '=AVERAGE(E2:E4)'],
            ['', '', '', 'StDev%', '=STDEV(E2:E4)'],
            ['', '', '', 'RSD%', '=(E6/E5)*100'],
        ]
        for row in kf_calc:
            df.loc[len(df)] = row
            
            # Apply all borders to top row given they are blank OR not blank
            worksheet.conditional_format('A1:E1', {'type': 'no_blanks', 'format': headerborder_format})
            worksheet.conditional_format('A1:E1', {'type': 'blanks', 'format': headerborder_format})
            
            # Apply bottom borders to A7:D7 given they are blank OR not blank
            worksheet.conditional_format('A7:D7', {'type': 'no_blanks', 'format': bottomborder_format})
            worksheet.conditional_format('A7:D7', {'type': 'blanks', 'format': bottomborder_format})
            
            # Apply dashed bottom borders to A4:D4 given they are blank OR not blank
            worksheet.conditional_format('A4:D4', {'type': 'no_blanks', 'format': dashedborder_format})
            worksheet.conditional_format('A4:D4', {'type': 'blanks', 'format': dashedborder_format})
            
            # Apply right borders to E2:E3 given they are blank OR not blank
            worksheet.conditional_format('E2:E3', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format('E2:E3', {'type': 'blanks', 'format': rightborder_format})
            
            # Apply right borders to E5:E6 given they are blank OR not blank
            worksheet.conditional_format('E5:E6', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format('E5:E6', {'type': 'blanks', 'format': rightborder_format})
            
            # Apply right and dashed bottom border to E4 given it is blank OR not blank
            worksheet.conditional_format('E4', {'type': 'blanks', 'format': sideborder_format})
            worksheet.conditional_format('E4', {'type': 'no_blanks', 'format': sideborder_format})
            
            # Apply corner border to E7 given it is blank OR not blank
            worksheet.conditional_format('E7', {'type': 'no_blanks', 'format': cornerborder_format})
            worksheet.conditional_format('E7', {'type': 'blanks', 'format': cornerborder_format})
            
            extra_sheetKF = {
           'Water Standard Check': ['Water Standard (1)', 'Water Standard (2)', 'Water Standard (3)']
           }
        additional_df = pd.DataFrame(extra_sheetKF)
        additional_df.to_excel(writer, index=False, sheet_name='Water Standard', startrow=0, startcol=0)
        worksheet = writer.sheets['Water Standard']
        headers = ['Water Standard Check', 'Mass(g)', 'Titrant(mL)', 'H2O%']
        for i, header in enumerate(headers):
            worksheet.write(0, i, header, bold_format)

        # Write new row headers in Cresol Testing sheet
        row_headers = ['Water Standard (1)', 'Water Standard (2)', 'Water Standard (3)']
        for i, header in enumerate(row_headers):
            worksheet.write(i+1, 0, header, bold_format)
            
        #set the file to open the column width to the length of a string
        KFcolumnwidth = len("Water Standard Check")
        #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(0, 0, KFcolumnwidth)

        KF_water_check = [
           ['', '', 'Mean%', '=AVERAGE(D2:D4)'],
           ['', '', 'StDev%', '=STDEV(D2:D4)'],
           ['', '', 'RSD%', '=(D6/D5)*100'],
        ]
        for row_num, row_data in enumerate(KF_water_check):
            for col_num, cell_data in enumerate(row_data):
                worksheet.write(row_num + 4, col_num + 0, cell_data)
        
        #format conditional for RSD check        
        worksheet.conditional_format('D7', {'type':     'cell',
                                       'criteria': 'between',
                                       'minimum':  0,
                                       'maximum':  1.5,
                                       'format':   green_format})
        
        worksheet.conditional_format('D7', {'type':     'cell',
                                       'criteria': 'not between',
                                       'minimum':  0,
                                       'maximum':  1.5,
                                       'format':   red_format})

    elif instrument == "KF & LECO CHN Combined":
        CHN_KF_calc = [
            ['Sample 1','','','','', '=100-SUM(C2:E2)','','','',''],
            ['Sample 1','','','','', '=100-SUM(C3:E3)','','','',''],
            ['Sample 1','','','','', '=100-SUM(C4:E4)','','','',''], 
            ['', 'Average', '=AVERAGE(C2:C4)','=AVERAGE(D2:D4)','=AVERAGE(E2:E4)','=AVERAGE(F2:F4)', '=AVERAGE(G2:G4)', '=C5/(100-G5)*100','=(D5-(G5*0.111))/(100-G5)*100','=(F5-(0.889*G5))/(100-G5)*100'],
            ['', 'StDev', '=STDEV(C2:C4)','=STDEV(D2:D4)','=STDEV(E2:E4)','=STDEV(F2:F4)','=STDEV(G2:G4)','','',''],
            ['', 'RSD', '=C6/C5*100','=D6/D5*100','=E6/E5*100','=F6/F5*100','=G6/G5*100','','',''],
        ]
        for row in CHN_KF_calc:
            df.loc[len(df)] = row
            
            #set the file to open the column width to the length of a string
            CHNKFcolumnwidth = len("O% Dry Basis")
            #worksheet.set_column(first_col, last_col, width, cell_format, options)
            worksheet.set_column(7, 9, CHNKFcolumnwidth)
            
            # Apply all borders to top row given they are blank OR not blank
            worksheet.conditional_format('A1:J1', {'type': 'no_blanks', 'format': headerborder_format})
            worksheet.conditional_format('A1:J1', {'type': 'blanks', 'format': headerborder_format})
            
            
            # Apply bottom borders to A7:I7 given they are blank OR not blank
            worksheet.conditional_format('A7:I7', {'type': 'no_blanks', 'format': bottomborder_format})
            worksheet.conditional_format('A7:I7', {'type': 'blanks', 'format': bottomborder_format})
            
            # Apply dashed bottom borders to A4:I4 given they are blank OR not blank
            worksheet.conditional_format('A4:I4', {'type': 'no_blanks', 'format': dashedborder_format})
            worksheet.conditional_format('A4:I4', {'type': 'blanks', 'format': dashedborder_format})
            
            # Apply right borders to J2:J3 given they are blank OR not blank
            worksheet.conditional_format('J2:J3', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format('J2:J3', {'type': 'blanks', 'format': rightborder_format})
            
            # Apply right borders to J5:J6 given they are blank OR not blank
            worksheet.conditional_format('J5:J6', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format('J5:J6', {'type': 'blanks', 'format': rightborder_format})
            
            # Apply right and dashed bottom border to J4 given it is blank OR not blank
            worksheet.conditional_format('J4', {'type': 'blanks', 'format': sideborder_format})
            worksheet.conditional_format('J4', {'type': 'no_blanks', 'format': sideborder_format})
            
            # Apply corner border to J7 given it is blank OR not blank
            worksheet.conditional_format('J7', {'type': 'blanks', 'format': cornerborder_format})
            worksheet.conditional_format('J7', {'type': 'no_blanks', 'format': cornerborder_format})
            
            # Create an additional sheet and write additional data
            extra_sheetCHNKF = {
                'Data': ['Additional Data 1', 'Additional Data 2', 'Additional Data 3']
            }
        additional_df = pd.DataFrame(extra_sheetCHNKF)
        additional_df.to_excel(writer, index=False, sheet_name='Cresol Testing', startrow=1, startcol=0)
            # Get the worksheet for the Cresol Testing sheet
        worksheet = writer.sheets['Cresol Testing']
        
        #set the file to open the column width to the length of a string
        Cresolcolumnwidth = len("Cresol Measured")
        #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(0, 1, Cresolcolumnwidth)
            
        #set conditional formatting for Cresol values in range green/red for in/out of range
        worksheet.conditional_format('B2', {'type':     'cell',
                                       'criteria': 'between',
                                       'minimum':  77.0,
                                       'maximum':  78.3,
                                       'format':   green_format})
        worksheet.conditional_format('B2', {'type':     'cell',
                                       'criteria': 'not between',
                                       'minimum':  77.0,
                                       'maximum':  78.3,
                                       'format':   red_format})
        
        worksheet.conditional_format('B3', {'type':     'cell',
                                       'criteria': 'between',
                                       'minimum':  7.4,
                                       'maximum':  7.9,
                                       'format':   green_format})
        worksheet.conditional_format('B3', {'type':     'cell',
                                       'criteria': 'not between',
                                       'minimum':  7.4,
                                       'maximum':  7.9,
                                       'format':   red_format})
        worksheet.conditional_format('B5', {'type':     'cell',
                                       'criteria': 'between',
                                       'minimum':  14.0,
                                       'maximum':  15.4,
                                       'format':   green_format})
        worksheet.conditional_format('B5', {'type':     'cell',
                                       'criteria': 'not between',
                                       'minimum':  14.0,
                                       'maximum':  15.4,
                                       'format':   red_format})
                 
       
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
            ['','=C11', 77.7, 77.0, 78.3], 
            ['','=D11', 7.5, 7.4, 7.9],
            ['','=E11', 0.03, 0.0, 0.1],
            ['','=F11', 14.8, 14.0, 15.4]
             ]
                     
        for row_num, row_data in enumerate(cresol_limits):
            for col_num, cell_data in enumerate(row_data):
                      worksheet.write(row_num + 1, col_num + 0, cell_data)
                      
        #write a table to put the three cresol values from the CHN data
        cresol_triplicate = [
            ['Cresol Triplicate Check', 'Mass (g)', 'C%', 'H%', 'N%', 'O% (diff)'],
            ['Cresol 1','','','','', '=100-SUM(C8:E8)'],
            ['Cresol 2','','','','', '=100-SUM(C9:E9)'],
            ['Cresol 3','','','','', '=100-SUM(C10:E10)'], 
            ['', 'Average', '=AVERAGE(C8:C10)','=AVERAGE(D8:D10)','=AVERAGE(E8:E10)','=AVERAGE(F8:F10)'],
            ['', 'StDev', '=STDEV(C8:C10)','=STDEV(D8:D10)','=STDEV(E8:E10)','=STDEV(F8:F10)'],
            ['', 'RSD', '=C12/C11*100','=D12/D11*100','=E12/E11*100','=F12/F11*100'],
        ]
        for row_num, row_data in enumerate(cresol_triplicate):
            for col_num, cell_data in enumerate(row_data):
                worksheet.write(row_num + 6, col_num + 0, cell_data)
        
        #write borders for the cresol triplicate data
        
        # Apply all borders to top row given they are blank OR not blank
        worksheet.conditional_format('A7:F7', {'type': 'no_blanks', 'format': headerborder_format})
        worksheet.conditional_format('A7:F7', {'type': 'blanks', 'format': headerborder_format})
        
        # Apply bottom borders to A13:E13 given they are blank OR not blank
        worksheet.conditional_format('A13:E13', {'type': 'no_blanks', 'format': bottomborder_format})
        worksheet.conditional_format('A13:E13', {'type': 'blanks', 'format': bottomborder_format})
        
        # Apply dashed bottom borders to A10:E10 given they are blank OR not blank
        worksheet.conditional_format('A10:E10', {'type': 'no_blanks', 'format': dashedborder_format})
        worksheet.conditional_format('A10:E10', {'type': 'blanks', 'format': dashedborder_format})
        
        # Apply right borders to F8:F9 given they are blank OR not blank
        worksheet.conditional_format('F8:F9', {'type': 'no_blanks', 'format': rightborder_format})
        worksheet.conditional_format('F8:F9', {'type': 'blanks', 'format': rightborder_format})
        
        # Apply right borders to F11:F12 given they are blank OR not blank
        worksheet.conditional_format('F11:F12', {'type': 'no_blanks', 'format': rightborder_format})
        worksheet.conditional_format('F11:F12', {'type': 'blanks', 'format': rightborder_format})
        
        # Apply right and dashed bottom border to F10 given it is not blank
        worksheet.conditional_format('F10', {'type': 'no_blanks', 'format': sideborder_format})
        
        # Apply corner border to F7 given it is not blank
        worksheet.conditional_format('F13', {'type': 'no_blanks', 'format': cornerborder_format})

            # Get the worksheet for the Cresol Testing sheet
        worksheet = writer.sheets['Cresol Testing']
        
    elif instrument == "Viscometer":
        #set the file to open the column width to the length of a string
        Vcolumnwidth = len("Temperature °C")
        #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(0, 4, Vcolumnwidth)
    
    elif instrument == "Density Meter":
        density_calc = [
            ['Sample 1','', ''],
            ['Sample 1','', ''],
            ['Mean%', '=AVERAGE(B2:B3)', ''],
            ['StDev%', '=STDEV(B2:B3)', ''],
            ['RSD%', '=(B5/B4)*100', ''],
        ]
        for row in density_calc:
            df.loc[len(df)] = row
            
            # Apply all borders to top row given they are blank OR not blank
            worksheet.conditional_format('A1:C1', {'type': 'no_blanks', 'format': headerborder_format})
            worksheet.conditional_format('A1:C1', {'type': 'blanks', 'format': headerborder_format})
            
            # Apply bottom borders to A6:C6 given they are blank OR not blank
            worksheet.conditional_format('A6:B6', {'type': 'no_blanks', 'format': bottomborder_format})
            worksheet.conditional_format('A6:B6', {'type': 'blanks', 'format': bottomborder_format})
            
            # Apply dashed bottom borders to A3:B3 given they are blank OR not blank
            worksheet.conditional_format('A3:B3', {'type': 'no_blanks', 'format': dashedborder_format})
            worksheet.conditional_format('A3:B3', {'type': 'blanks', 'format': dashedborder_format})
            
            # Apply right borders to C2 given they are blank OR not blank
            worksheet.conditional_format('C2', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format('C2', {'type': 'blanks', 'format': rightborder_format})
            
            # Apply right borders to C4:C5 given they are blank OR not blank
            worksheet.conditional_format('C4:C5', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format('C4:C5', {'type': 'blanks', 'format': rightborder_format})
            
            # Apply right and dashed bottom border to C4' given it is blank OR not blank
            worksheet.conditional_format('C3', {'type': 'blanks', 'format': sideborder_format})
            worksheet.conditional_format('C3', {'type': 'no_blanks', 'format': sideborder_format})
            
            # Apply corner border to C6 given it is blank OR not blank
            worksheet.conditional_format('C6', {'type': 'no_blanks', 'format': cornerborder_format})
            worksheet.conditional_format('C6', {'type': 'blanks', 'format': cornerborder_format})
            
            #set the file to open the column width to the length of a string
            DensityAnalysiscolumnwidth = len("Temperature (°C)")
            #worksheet.set_column(first_col, last_col, width, cell_format, options)
            worksheet.set_column(0, 2, DensityAnalysiscolumnwidth)
            
            extra_sheetDensity = {
           'Water Check': ['Water (1)', 'Water (2)']
           }
        additional_df = pd.DataFrame(extra_sheetDensity)
        additional_df.to_excel(writer, index=False, sheet_name='Water Check', startrow=0, startcol=0)
        worksheet = writer.sheets['Water Check']
        headers = ['Water Check', 'Density (g/mL)', 'Temperature (°C)']
        for i, header in enumerate(headers):
            worksheet.write(0, i, header, bold_format)

        # Write new row headers in Cresol Testing sheet
        row_headers = ['Water (1)', 'Water (2)']
        for i, header in enumerate(row_headers):
            worksheet.write(i+1, 0, header, bold_format)
            
        #set the file to open the column width to the length of a string
        Densitycolumnwidth = len("Temperature (°C)")
        #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(0, 2, Densitycolumnwidth)
    
    elif instrument == "Acids Titration":
        Acids_calc = [
            ['Sample 1','','','=C2-B2','=B2*56.1', '=C2*56.1','=D2*56.1'],
            ['Sample 1','','','=C3-B3','=B3*56.1', '=C3*56.1','=D3*56.1'],
            ['Average','=AVERAGE(B2:B3)','=AVERAGE(C2:C3)','=AVERAGE(D2:D3)','=AVERAGE(E2:E3)', '=AVERAGE(F2:F3)','=AVERAGE(G2:G3)'], 
            ['Range', '=ABS(B3-B2)', '=ABS(C3-C2)','=ABS(D3-D2)','=ABS(E3-E2)','=ABS(F3-F2)', '=ABS(G3-G2)'],
            ]
        for row in Acids_calc:
            df.loc[len(df)] = row
            
        #set the file to open the column width to the length of a string
        Acidscolumnwidth = len("Carboxylic Acid Number mg KOH/g")
        #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(4, 6, Acidscolumnwidth) 
            
            # Apply all borders to top row given they are blank OR not blank
        worksheet.conditional_format('A1:G1', {'type': 'no_blanks', 'format': headerborder_format})
        worksheet.conditional_format('A1:G1', {'type': 'blanks', 'format': headerborder_format})
            
            
            # Apply bottom borders to A5:F5 given they are blank OR not blank
        worksheet.conditional_format('A5:F5', {'type': 'no_blanks', 'format': bottomborder_format})
        worksheet.conditional_format('A5:F5', {'type': 'blanks', 'format': bottomborder_format})
            
            
            # Apply right borders to G2:G46 given they are blank OR not blank
        worksheet.conditional_format('G2:G4', {'type': 'no_blanks', 'format': rightborder_format})
        worksheet.conditional_format('G2:G4', {'type': 'blanks', 'format': rightborder_format})
            
            # Apply corner border to G5 given it is blank OR not blank
        worksheet.conditional_format('G5', {'type': 'no_blanks', 'format': cornerborder_format})
        worksheet.conditional_format('G5', {'type': 'blanks', 'format': cornerborder_format})
    
            # Create an additional sheet and write additional data
        extra_sheetAcids = {'Sample ID': ["='Analysis'!A2"]
            }
        additional_df = pd.DataFrame(extra_sheetAcids)
        additional_df.to_excel(writer, index=False, sheet_name='Summary Analysis', startrow=0, startcol=0)
            # Get the worksheet for the Cresol Testing sheet and tell it to start writing below column
        worksheet = writer.sheets['Summary Analysis']
            
        # Write new column and row headers in Summary Analysis sheet
        headers = ['Sample ID', 'CAN mol/kg', 'TAN mol/kg', 'PhAN mol/kg', 'Carboxylic Acid Number mg KOH/g', 'Total Acid Number mg KOH/g', 'Phenolic Acid Number mg KOH/g']
        for i, header in enumerate(headers):
            worksheet.write(0, i, header, bold_format)
                     
        # Write Summary Analysis row headers or formulas below the columns
        #using a " ' ' " method because of formula to call Analysis page
        Acids_summary_analysis = [
            ['',"='Analysis'!B4","='Analysis'!C4", "='Analysis'!D4", "='Analysis'!E4", "='Analysis'!F4","='Analysis'!G4"], 
            ]
        for row_num, row_data in enumerate(Acids_summary_analysis):
            for col_num, cell_data in enumerate(row_data):
                      worksheet.write(row_num + 1, col_num + 0, cell_data)
                      
        #set the file to open the column width to the length of a string
        summaryAcidscolumnwidth = len("Carboxylic Acid Number mg K")
        #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(4, 6, summaryAcidscolumnwidth)
        
        #set the file to open the column width to the length of a string
        summaryAcidscolumnwidth = len("PhAN mol/kg")
        #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(1, 3, summaryAcidscolumnwidth)
                         
            # Get the worksheet for the Cresol Testing sheet
        worksheet = writer.sheets['Summary Analysis']

    # Write the DataFrame to the default sheet
    df.to_excel(writer, index=False, sheet_name='Analysis', engine='xlsxwriter')
    worksheet = writer.sheets['Analysis']
    
                
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
    
  
