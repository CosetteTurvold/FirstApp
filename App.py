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
    'Acids Titration': ['Sample ID', 'CAN mol/kg', 'TAN mol/kg', 'PhAN mol/kg', 'Carboxylic Acid Number mg KOH/g','Total Acid Number mg KOH/g', 'Phenolic Acid Number mg KOH/g'],
    'Carbonyls Titration': ['Sample ID', 'Carbonyls mol/kg']
}


def generate_excel_file(instrument, num_request):
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
        CHN_data = []
        CHN_template = [
            ['Sample 1','','','','', '=100-SUM(C2:E2)'],
            ['Sample 1','','','','', '=100-SUM(C3:E3)'],
            ['Sample 1','','','','', '=100-SUM(C4:E4)'], 
            ['', 'Average', '=AVERAGE(C2:C4)','=AVERAGE(D2:D4)','=AVERAGE(E2:E4)','=AVERAGE(F2:F4)'],
            ['', 'StDev', '=STDEV(C2:C4)','=STDEV(D2:D4)','=STDEV(E2:E4)','=STDEV(F2:F4)'],
            ['', 'RSD', '=C6/C5*100','=D6/D5*100','=E6/E5*100','=F6/F5*100'],
        ]
        average_row = 3  # Row number for Mean%
        stdev_row = 4    # Row number for StDev%
        rsd_row = 5      # Row number for RSD%
        sample_row1 = 0 # Row number for "Sample X" labels
        sample_row2 = 1
        sample_row3 = 2
        #for row in CHN_template:
         #   df.loc[len(df)] = row

        # Replicate kf_calc based on the number of samples
        for sample_num in range(1, num_request + 1):
            # Create a copy of the sample template for this sample
            CHN_calc = [row[:] for row in CHN_template]

            # Calculate the start row for this sample
            start_row = (sample_num - 1) * 8 + 1

            # Update the stat formulas (average)
            CHN_calc[average_row][2] = f'=AVERAGE(C{start_row + 1}:C{start_row + 3})'
            CHN_calc[average_row][3] = f'=AVERAGE(D{start_row + 1}:D{start_row + 3})'
            CHN_calc[average_row][4] = f'=AVERAGE(E{start_row + 1}:E{start_row + 3})'
            CHN_calc[average_row][5] = f'=AVERAGE(F{start_row + 1}:F{start_row + 3})'
            # Update the stat formulas (stdev)
            CHN_calc[stdev_row][2] = f'=STDEV(C{start_row + 1}:C{start_row + 3})'
            CHN_calc[stdev_row][3] = f'=STDEV(D{start_row + 1}:D{start_row + 3})'
            CHN_calc[stdev_row][4] = f'=STDEV(E{start_row + 1}:E{start_row + 3})'
            CHN_calc[stdev_row][5] = f'=STDEV(F{start_row + 1}:F{start_row + 3})'
            # Update the stat formulas (RSD)
            CHN_calc[rsd_row][2] = f'=(C{start_row + 5})/(C{start_row + 4}) * 100'
            CHN_calc[rsd_row][3] = f'=(D{start_row + 5})/(D{start_row + 4}) * 100'
            CHN_calc[rsd_row][4] = f'=(E{start_row + 5})/(E{start_row + 4}) * 100'
            CHN_calc[rsd_row][5] = f'=(F{start_row + 5})/(F{start_row + 4}) * 100'
            #make changes to the Oxygen calculations
            CHN_calc[sample_row1][-1] = f'=100-SUM(C{start_row + 1}:E{start_row + 1})'
            CHN_calc[sample_row2][-1] = f'=100-SUM(C{start_row + 2}:E{start_row + 2})'
            CHN_calc[sample_row3][-1] = f'=100-SUM(C{start_row + 3}:E{start_row + 3})'
            # Add the Sample label row
            CHN_calc[sample_row1][0] = f'Sample {sample_num}'
            CHN_calc[sample_row2][0] = f'Sample {sample_num}'
            CHN_calc[sample_row3][0] = f'Sample {sample_num}'
            # Append the sample's data to kf_data
            CHN_data.extend(CHN_calc)

            # Add an empty row between sample templates, but not after the last sample
            if sample_num < num_request:
                CHN_data.extend([[''] * len(CHN_template[0]), columns])

        # Apply all borders, conditional formatting, and create the Excel file as before

                
            # Apply all borders to top row
            worksheet.conditional_format(f'A{start_row}:F{start_row}', {'type': 'no_blanks', 'format': headerborder_format})
            worksheet.conditional_format(f'A{start_row}:F{start_row}', {'type': 'blanks', 'format': headerborder_format})

            # Apply bottom borders to A7:D7
            worksheet.conditional_format(f'A{start_row + 6}:E{start_row + 6}', {'type': 'no_blanks', 'format': bottomborder_format})
            worksheet.conditional_format(f'A{start_row + 6}:E{start_row + 6}', {'type': 'blanks', 'format': bottomborder_format})

            # Apply dashed bottom borders to A4:D4
            worksheet.conditional_format(f'A{start_row + 3}:E{start_row + 3}', {'type': 'no_blanks', 'format': dashedborder_format})
            worksheet.conditional_format(f'A{start_row + 3}:E{start_row + 3}', {'type': 'blanks', 'format': dashedborder_format})

            # Apply right borders to E2:E3
            worksheet.conditional_format(f'F{start_row + 1}:F{start_row + 2}', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format(f'F{start_row + 1}:F{start_row + 2}', {'type': 'blanks', 'format': rightborder_format})
            
            # Apply right borders to E5:E6
            worksheet.conditional_format(f'F{start_row + 4}:F{start_row + 5}', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format(f'F{start_row + 4}:F{start_row + 5}', {'type': 'blanks', 'format': rightborder_format})

            # Apply right and dashed bottom border to E4
            worksheet.conditional_format(f'F{start_row + 3}', {'type': 'blanks', 'format': sideborder_format})
            worksheet.conditional_format(f'F{start_row + 3}', {'type': 'no_blanks', 'format': sideborder_format})

            # Apply corner border to E7
            worksheet.conditional_format(f'F{start_row + 6}', {'type': 'no_blanks', 'format': cornerborder_format})
            worksheet.conditional_format(f'F{start_row + 6}', {'type': 'blanks', 'format': cornerborder_format})
            
            # Convert kf_data into a DataFrame for each sample
            df = pd.DataFrame(CHN_data, columns=columns)
            
            # Apply conditional formatting to this sample
            worksheet = writer.sheets['Analysis']
                
            # Write the DataFrame to the default sheet
            df.to_excel(writer, index=False, sheet_name='Analysis', engine='xlsxwriter')
        
        
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
            worksheet = writer.sheets['Analysis']
            
    elif instrument == "Karl Fischer":
        kf_data = []
        # Define the template for one sample
        sample_template = [
            ['Sample 1', '', '', '', ''],
            ['Sample 1', '', '', '', ''],
            ['Sample 1', '', '', '', ''],
            ['', '', '', 'Mean%', '=AVERAGE(E2:E4)'],
            ['', '', '', 'StDev%', '=STDEV(E2:E4)'],
            ['', '', '', 'RSD%', '=(E6/E5)*100'],
        ]
        average_row = 3  # Row number for Mean%
        stdev_row = 4    # Row number for StDev%
        rsd_row = 5      # Row number for RSD%
        sample_row1 = 0 # Row number for "Sample X" labels
        sample_row2 = 1
        sample_row3 = 2
        # Replicate kf_calc based on the number of samples
        for sample_num in range(1, num_request + 1):
        # Create a copy of the sample template for this sample
            kf_calc = [row[:] for row in sample_template]

        # Calculate the start row for this sample
            start_row = (sample_num - 1) * 8 + 1

        # Update the formulas for this sample
            kf_calc[average_row][-1] = f'=AVERAGE(E{start_row + 1}:E{start_row + 3})'
            kf_calc[stdev_row][-1] = f'=STDEV(E{start_row + 1}:E{start_row + 3})'
            kf_calc[rsd_row][-1] = f'=(E{start_row + 5})/(E{start_row + 4}) * 100'

        # Add the Sample label row
            kf_calc[sample_row1][0] = f'Sample {sample_num}'
            kf_calc[sample_row2][0] = f'Sample {sample_num}'
            kf_calc[sample_row3][0] = f'Sample {sample_num}'
        # Append the sample's data to kf_data
            kf_data.extend(kf_calc)

        # Add an empty row between sample templates, but not after the last sample
            if sample_num < num_request:
                kf_data.extend([[''] * len(sample_template[0]), columns])

    # Apply all borders, conditional formatting, and create the Excel file as before

            
        # Apply all borders to top row
            worksheet.conditional_format(f'A{start_row}:E{start_row}', {'type': 'no_blanks', 'format': headerborder_format})
            worksheet.conditional_format(f'A{start_row}:E{start_row}', {'type': 'blanks', 'format': headerborder_format})

        # Apply bottom borders to A7:D7
            worksheet.conditional_format(f'A{start_row + 6}:D{start_row + 6}', {'type': 'no_blanks', 'format': bottomborder_format})
            worksheet.conditional_format(f'A{start_row + 6}:D{start_row + 6}', {'type': 'blanks', 'format': bottomborder_format})

        # Apply dashed bottom borders to A4:D4
            worksheet.conditional_format(f'A{start_row + 3}:D{start_row + 3}', {'type': 'no_blanks', 'format': dashedborder_format})
            worksheet.conditional_format(f'A{start_row + 3}:D{start_row + 3}', {'type': 'blanks', 'format': dashedborder_format})

        # Apply right borders to E2:E3
            worksheet.conditional_format(f'E{start_row + 1}:E{start_row + 2}', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format(f'E{start_row + 1}:E{start_row + 2}', {'type': 'blanks', 'format': rightborder_format})
        
        # Apply right borders to E5:E6
            worksheet.conditional_format(f'E{start_row + 4}:E{start_row + 5}', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format(f'E{start_row + 4}:E{start_row + 5}', {'type': 'blanks', 'format': rightborder_format})

        # Apply right and dashed bottom border to E4
            worksheet.conditional_format(f'E{start_row + 3}', {'type': 'blanks', 'format': sideborder_format})
            worksheet.conditional_format(f'E{start_row + 3}', {'type': 'no_blanks', 'format': sideborder_format})

        # Apply corner border to E7
            worksheet.conditional_format(f'E{start_row + 6}', {'type': 'no_blanks', 'format': cornerborder_format})
            worksheet.conditional_format(f'E{start_row + 6}', {'type': 'blanks', 'format': cornerborder_format})
        
        # Convert kf_data into a DataFrame for each sample
        df = pd.DataFrame(kf_data, columns=columns)
        
        # Apply conditional formatting to this sample
        worksheet = writer.sheets['Analysis']
            
        # Write the DataFrame to the default sheet
        df.to_excel(writer, index=False, sheet_name='Analysis', engine='xlsxwriter')
    
            
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
            
            
            # Write the DataFrame to the default sheet
        df.to_excel(writer, index=False, sheet_name='Analysis', engine='xlsxwriter')
        worksheet = writer.sheets['Analysis']


    elif instrument == "KF & LECO CHN Combined":
        CHN_KF_data = []
        CHN_KF_template = [
            ['Sample 1','','','','', '=100-SUM(C2:E2)','','','',''],
            ['Sample 1','','','','', '=100-SUM(C3:E3)','','','',''],
            ['Sample 1','','','','', '=100-SUM(C4:E4)','','','',''], 
            ['', 'Average', '=AVERAGE(C2:C4)','=AVERAGE(D2:D4)','=AVERAGE(E2:E4)','=AVERAGE(F2:F4)', '=AVERAGE(G2:G4)', '=C5/(100-G5)*100','=(D5-(G5*0.111))/(100-G5)*100','=(F5-(0.889*G5))/(100-G5)*100'],
            ['', 'StDev', '=STDEV(C2:C4)','=STDEV(D2:D4)','=STDEV(E2:E4)','=STDEV(F2:F4)','=STDEV(G2:G4)','','',''],
            ['', 'RSD', '=C6/C5*100','=D6/D5*100','=E6/E5*100','=F6/F5*100','=G6/G5*100','','',''],
        ]
        average_row = 3  # Row number for Mean%
        stdev_row = 4    # Row number for StDev%
        rsd_row = 5      # Row number for RSD%
        sample_row1 = 0  # Row number for "Sample X" labels
        sample_row2 = 1
        sample_row3 = 2
        #for row in CHN_KF_calc:
         #   df.loc[len(df)] = row
            
            #set the file to open the column width to the length of a string
        CHNKFcolumnwidth = len("O% Dry Basis")
            #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(7, 9, CHNKFcolumnwidth)
            
            # Replicate kf_calc based on the number of samples
        for sample_num in range(1, num_request + 1):
                # Create a copy of the sample template for this sample
            CHN_KF_calc = [row[:] for row in CHN_KF_template]

                # Calculate the start row for this sample
            start_row = (sample_num - 1) * 8 + 1
            
            # Update the dry basis calculations
            CHN_KF_calc[average_row][7] = f'=(C{start_row + 4})/(100-G{start_row + 4})*100'
            CHN_KF_calc[average_row][8] = f'=(D{start_row + 4}-(G{start_row + 4}*0.111))/(100-G{start_row + 4})*100'
            CHN_KF_calc[average_row][9] = f'=(F{start_row + 4}-(0.889*G{start_row + 4}))/(100-G{start_row + 4})*100'
            # Update the stat formulas (average)
            CHN_KF_calc[average_row][2] = f'=AVERAGE(C{start_row + 1}:C{start_row + 3})'
            CHN_KF_calc[average_row][3] = f'=AVERAGE(D{start_row + 1}:D{start_row + 3})'
            CHN_KF_calc[average_row][4] = f'=AVERAGE(E{start_row + 1}:E{start_row + 3})'
            CHN_KF_calc[average_row][5] = f'=AVERAGE(F{start_row + 1}:F{start_row + 3})'
            CHN_KF_calc[average_row][6] = f'=AVERAGE(G{start_row + 1}:G{start_row + 3})'
            # Update the stat formulas (stdev)
            CHN_KF_calc[stdev_row][2] = f'=STDEV(C{start_row + 1}:C{start_row + 3})'
            CHN_KF_calc[stdev_row][3] = f'=STDEV(D{start_row + 1}:D{start_row + 3})'
            CHN_KF_calc[stdev_row][4] = f'=STDEV(E{start_row + 1}:E{start_row + 3})'
            CHN_KF_calc[stdev_row][5] = f'=STDEV(F{start_row + 1}:F{start_row + 3})'
            CHN_KF_calc[stdev_row][6] = f'=STDEV(G{start_row + 1}:G{start_row + 3})'
            # Update the stat formulas (RSD)
            CHN_KF_calc[rsd_row][2] = f'=(C{start_row + 5})/(C{start_row + 4}) * 100'
            CHN_KF_calc[rsd_row][3] = f'=(D{start_row + 5})/(D{start_row + 4}) * 100'
            CHN_KF_calc[rsd_row][4] = f'=(E{start_row + 5})/(E{start_row + 4}) * 100'
            CHN_KF_calc[rsd_row][5] = f'=(F{start_row + 5})/(F{start_row + 4}) * 100'
            CHN_KF_calc[rsd_row][6] = f'=(G{start_row + 5})/(G{start_row + 4}) * 100'
            #make changes to the Oxygen calculations
            CHN_KF_calc[sample_row1][5] = f'=100-SUM(C{start_row + 1}:E{start_row + 1})'
            CHN_KF_calc[sample_row2][5] = f'=100-SUM(C{start_row + 2}:E{start_row + 2})'
            CHN_KF_calc[sample_row3][5] = f'=100-SUM(C{start_row + 3}:E{start_row + 3})'
            # Add the Sample label row
            CHN_KF_calc[sample_row1][0] = f'Sample {sample_num}'
            CHN_KF_calc[sample_row2][0] = f'Sample {sample_num}'
            CHN_KF_calc[sample_row3][0] = f'Sample {sample_num}'
            # Append the sample's data to kf_data
            CHN_KF_data.extend(CHN_KF_calc)
            
            # Add an empty row between sample templates, but not after the last sample
            if sample_num < num_request:
                CHN_KF_data.extend([[''] * len(CHN_KF_template[0]), columns])
                
            # Apply all borders to top row given they are blank OR not blank
            worksheet.conditional_format(f'A{start_row}:J{start_row}', {'type': 'no_blanks', 'format': headerborder_format})
            worksheet.conditional_format(f'A{start_row}:J{start_row}', {'type': 'blanks', 'format': headerborder_format})
            
            # Apply bottom borders to A7:I7 given they are blank OR not blank
            worksheet.conditional_format(f'A{start_row + 6}:I{start_row + 6}', {'type': 'no_blanks', 'format': bottomborder_format})
            worksheet.conditional_format(f'A{start_row + 6}:I{start_row + 6}', {'type': 'blanks', 'format': bottomborder_format})
            
            # Apply dashed bottom borders to A4:I4 given they are blank OR not blank
            worksheet.conditional_format(f'A{start_row + 3}:I{start_row + 3}', {'type': 'no_blanks', 'format': dashedborder_format})
            worksheet.conditional_format(f'A{start_row + 3}:I{start_row + 3}', {'type': 'blanks', 'format': dashedborder_format})
            
            # Apply right borders to J2:J3 given they are blank OR not blank
            worksheet.conditional_format(f'J{start_row + 1}:J{start_row + 2}', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format(f'J{start_row + 1}:J{start_row + 2}', {'type': 'blanks', 'format': rightborder_format})
            
            # Apply right borders to J5:J6 given they are blank OR not blank
            worksheet.conditional_format(f'J{start_row + 4}:J{start_row + 5}', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format(f'J{start_row + 4}:J{start_row + 5}', {'type': 'blanks', 'format': rightborder_format})
            
            # Apply right and dashed bottom border to J4 given it is blank OR not blank
            worksheet.conditional_format(f'J{start_row + 3}', {'type': 'blanks', 'format': sideborder_format})
            worksheet.conditional_format(f'J{start_row + 3}', {'type': 'no_blanks', 'format': sideborder_format})
            
            # Apply corner border to J7 given it is blank OR not blank
            worksheet.conditional_format(f'J{start_row + 6}', {'type': 'blanks', 'format': cornerborder_format})
            worksheet.conditional_format(f'J{start_row + 6}', {'type': 'no_blanks', 'format': cornerborder_format})
            
            # Convert kf_data into a DataFrame for each sample
            df = pd.DataFrame(CHN_KF_data, columns=columns)
            
            # Apply conditional formatting to this sample
            worksheet = writer.sheets['Analysis']
                
            # Write the DataFrame to the default sheet
            df.to_excel(writer, index=False, sheet_name='Analysis', engine='xlsxwriter')
            
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
            worksheet = writer.sheets['Analysis']
        
    elif instrument == "Viscometer":
        viscometer_data = []
        viscometer_template = [
            ['Sample 1', '', '', '', '']
            ]
        sample_row1 = 0
        #set the file to open the column width to the length of a string
        Vcolumnwidth = len("Temperature °C")
        #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(0, 4, Vcolumnwidth)
        # Replicate kf_calc based on the number of samples
        for sample_num in range(1, num_request + 1):
        # Create a copy of the sample template for this sample
            viscometer_calc = [row[:] for row in viscometer_template]
            # Calculate the start row for this sample
            #start_row = (sample_num - 1) * 8 + 1
            # Add the Sample label row
            viscometer_calc[sample_row1][0] = f'Sample {sample_num}'
            # Append the sample's data to kf_data
            viscometer_data.extend(viscometer_calc)
            # Convert kf_data into a DataFrame for each sample
        df = pd.DataFrame(viscometer_data, columns=columns)
            
            # Apply conditional formatting to this sample
        worksheet = writer.sheets['Analysis']
                
            # Write the DataFrame to the default sheet
        df.to_excel(writer, index=False, sheet_name='Analysis', engine='xlsxwriter')
        
        extra_sheetViscosity = {
           'Standard Check': ['Standard']
           }
        additional_df = pd.DataFrame(extra_sheetViscosity)
        additional_df.to_excel(writer, index=False, sheet_name='Standard Check', startrow=0, startcol=0)
        worksheet = writer.sheets['Standard Check']
        #set the file to open the column width to the length of a string
        viscositystandardcolumnwidth = len("Temperature (°C)")
        #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(0, 4, viscositystandardcolumnwidth) 
        headers = ['Standard Check', 'Viscosity (cP)', 'Torque (%)', 'Speed (RPM)', 'Temperature (°C)', 'Spindle']
        for i, header in enumerate(headers):
            worksheet.write(0, i, header, bold_format)

        # Write new row headers in Cresol Testing sheet
        row_headers = ['Standard']
        for i, header in enumerate(row_headers):
            worksheet.write(i+1, 0, header, bold_format)
            
        # Write the DataFrame to the default sheet
        df.to_excel(writer, index=False, sheet_name='Analysis', engine='xlsxwriter')
        worksheet = writer.sheets['Analysis']

    
    elif instrument == "Density Meter":
        density_data = []
        density_template = [
            ['Sample 1','', ''],
            ['Sample 1','', ''],
            ['Mean', '=AVERAGE(B2:B3)', ''],
            ['StDev', '=STDEV(B2:B3)', ''],
            ['RSD', '=(B5/B4)*100', ''],
        ]
        average_row = 2  # Row number for Mean%
        stdev_row = 3    # Row number for StDev%
        rsd_row = 4      # Row number for RSD%
        sample_row1 = 0 # Row number for "Sample X" labels
        sample_row2 = 1
        #for row in density_calc:
        #    df.loc[len(df)] = row
        
        #set the file to open the column width to the length of a string
        DensityAnalysiscolumnwidth = len("Temperature (°C)")
        #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(0, 2, DensityAnalysiscolumnwidth)
        
        # Replicate kf_calc based on the number of samples
        for sample_num in range(1, num_request + 1):
        # Create a copy of the sample template for this sample
            density_calc = [row[:] for row in density_template]

        # Calculate the start row for this sample
            start_row = (sample_num - 1) * 7 + 1
            
            # Update the formulas for this sample
            density_calc[average_row][1] = f'=AVERAGE(B{start_row + 1}:B{start_row + 2})'
            density_calc[stdev_row][1] = f'=STDEV(B{start_row + 1}:B{start_row + 2})'
            density_calc[rsd_row][1] = f'=(B{start_row + 4})/(B{start_row + 3}) * 100'
            # Add the Sample label row
            density_calc[sample_row1][0] = f'Sample {sample_num}'
            density_calc[sample_row2][0] = f'Sample {sample_num}'
            # Append the sample's data to kf_data
            density_data.extend(density_calc)
            # Add an empty row between sample templates, but not after the last sample
            if sample_num < num_request:
                density_data.extend([[''] * len(density_template[0]), columns])
    
            # Apply all borders to top row given they are blank OR not blank
            worksheet.conditional_format(f'A{start_row}:C{start_row}', {'type': 'no_blanks', 'format': headerborder_format})
            worksheet.conditional_format(f'A{start_row}:C{start_row}', {'type': 'blanks', 'format': headerborder_format})
            
            # Apply bottom borders to A6:C6 given they are blank OR not blank
            worksheet.conditional_format(f'A{start_row + 5}:B{start_row + 5}', {'type': 'no_blanks', 'format': bottomborder_format})
            worksheet.conditional_format(f'A{start_row + 5}:B{start_row + 5}', {'type': 'blanks', 'format': bottomborder_format})
            
            # Apply dashed bottom borders to A3:B3 given they are blank OR not blank
            worksheet.conditional_format(f'A{start_row + 2}:B{start_row + 2}', {'type': 'no_blanks', 'format': dashedborder_format})
            worksheet.conditional_format(f'A{start_row + 2}:B{start_row + 2}', {'type': 'blanks', 'format': dashedborder_format})
            
            # Apply right borders to C2 given they are blank OR not blank
            worksheet.conditional_format(f'C{start_row + 1}', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format(f'C{start_row + 1}', {'type': 'blanks', 'format': rightborder_format})
            
            # Apply right borders to C4:C5 given they are blank OR not blank
            worksheet.conditional_format(f'C{start_row + 3}:C{start_row + 4}', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format(f'C{start_row + 3}:C{start_row + 4}', {'type': 'blanks', 'format': rightborder_format})
            
            # Apply right and dashed bottom border to C3' given it is blank OR not blank
            worksheet.conditional_format(f'C{start_row + 2}', {'type': 'blanks', 'format': sideborder_format})
            worksheet.conditional_format(f'C{start_row + 2}', {'type': 'no_blanks', 'format': sideborder_format})
            
            # Apply corner border to C6 given it is blank OR not blank
            worksheet.conditional_format(f'C{start_row + 5}', {'type': 'no_blanks', 'format': cornerborder_format})
            worksheet.conditional_format(f'C{start_row + 5}', {'type': 'blanks', 'format': cornerborder_format})
            
            
            # Convert kf_data into a DataFrame for each sample
        df = pd.DataFrame(density_data, columns=columns)
            
            # Apply conditional formatting to this sample
        worksheet = writer.sheets['Analysis']
                
            # Write the DataFrame to the default sheet
        df.to_excel(writer, index=False, sheet_name='Analysis', engine='xlsxwriter')
            
        extra_sheetDensity = {
           'Water Check': ['Water (1)', 'Water (2)']
           }
        additional_df = pd.DataFrame(extra_sheetDensity)
        additional_df.to_excel(writer, index=False, sheet_name='Water Check', startrow=0, startcol=0)
        worksheet = writer.sheets['Water Check']
        #set the file to open the column width to the length of a string
        densitywatercheckcolumnwidth = len("Temperature (°C)")
        #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(0, 2, densitywatercheckcolumnwidth) 
        headers = ['Water Check', 'Density (g/mL)', 'Temperature (°C)']
        for i, header in enumerate(headers):
            worksheet.write(0, i, header, bold_format)

        # Write new row headers in Cresol Testing sheet
        row_headers = ['Water (1)', 'Water (2)']
        for i, header in enumerate(row_headers):
            worksheet.write(i+1, 0, header, bold_format)
            
        # Write the DataFrame to the default sheet
        df.to_excel(writer, index=False, sheet_name='Analysis', engine='xlsxwriter')
        worksheet = writer.sheets['Analysis']
    
    elif instrument == "Acids Titration":
        acids_data = []
        acids_template = [
            ['Sample 1','','','=C2-B2','=B2*56.1', '=C2*56.1','=D2*56.1'],
            ['Sample 1','','','=C3-B3','=B3*56.1', '=C3*56.1','=D3*56.1'],
            ['Average','=AVERAGE(B2:B3)','=AVERAGE(C2:C3)','=AVERAGE(D2:D3)','=AVERAGE(E2:E3)', '=AVERAGE(F2:F3)','=AVERAGE(G2:G3)'], 
            ['Range', '=ABS(B3-B2)', '=ABS(C3-C2)','=ABS(D3-D2)','=ABS(E3-E2)','=ABS(F3-F2)', '=ABS(G3-G2)'],
            ]
        average_row = 2  # Row number for Mean%
        range_row = 3    # Row number for StDev%
        sample_row1 = 0 # Row number for "Sample X" labels
        sample_row2 = 1
        
        #set the file to open the column width to the length of a string
        Acidscolumnwidth = len("Carboxylic Acid Number mg KOH/g")
        #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(4, 6, Acidscolumnwidth) 
        
        # Replicate kf_calc based on the number of samples
        for sample_num in range(1, num_request + 1):
        # Create a copy of the sample template for this sample
            acids_calc = [row[:] for row in acids_template]

        # Calculate the start row for this sample
            start_row = (sample_num - 1) * 6 + 1 
            
       # Update the average forumlas 
            acids_calc[average_row][1] = f'=AVERAGE(B{start_row + 1}:B{start_row + 2})'
            acids_calc[average_row][2] = f'=AVERAGE(C{start_row + 1}:C{start_row + 2})'
            acids_calc[average_row][3] = f'=AVERAGE(D{start_row + 1}:D{start_row + 2})'
            acids_calc[average_row][4] = f'=AVERAGE(E{start_row + 1}:E{start_row + 2})'
            acids_calc[average_row][5] = f'=AVERAGE(F{start_row + 1}:F{start_row + 2})'
            acids_calc[average_row][6] = f'=AVERAGE(G{start_row + 1}:G{start_row + 2})'
       # Update the range forumlas      
            acids_calc[range_row][1] = f'=ABS(B{start_row + 2}-B{start_row + 1})'
            acids_calc[range_row][2] = f'=ABS(C{start_row + 2}-C{start_row + 1})'
            acids_calc[range_row][3] = f'=ABS(D{start_row + 2}-D{start_row + 1})'
            acids_calc[range_row][4] = f'=ABS(E{start_row + 2}-E{start_row + 1})'
            acids_calc[range_row][5] = f'=ABS(F{start_row + 2}-F{start_row + 1})'
            acids_calc[range_row][6] = f'=ABS(G{start_row + 2}-G{start_row + 1})'
            
        # Update the CAN calculations
            acids_calc[sample_row1][3] = f'=(C{start_row + 1}-B{start_row + 1})'
            acids_calc[sample_row1][4] = f'=(B{start_row + 1} * 56.1)'
            acids_calc[sample_row1][5] = f'=(C{start_row + 1} * 56.1)'
            acids_calc[sample_row1][6] = f'=(D{start_row + 1} * 56.1)'
            acids_calc[sample_row2][3] = f'=(C{start_row + 2}-B{start_row + 2})'
            acids_calc[sample_row2][4] = f'=(B{start_row + 2} * 56.1)'
            acids_calc[sample_row2][5] = f'=(C{start_row + 2} * 56.1)'
            acids_calc[sample_row2][6] = f'=(D{start_row + 2} * 56.1)'
            acids_calc[sample_row1][1] = 0
            acids_calc[sample_row1][2] = 0

       # Update the Sample label row
            acids_calc[sample_row1][0] = f'Sample {sample_num}'
            acids_calc[sample_row2][0] = f'Sample {sample_num}'
       # Append the sample's data to kf_data
            acids_data.extend(acids_calc)

       # Add an empty row between sample templates, but not after the last sample
            if sample_num < num_request:
               acids_data.extend([[''] * len(acids_template[0]), columns])
            
                # Apply all borders to top row given they are blank OR not blank
            worksheet.conditional_format(f'A{start_row}:G{start_row}', {'type': 'no_blanks', 'format': headerborder_format})
            worksheet.conditional_format(f'A{start_row}:G{start_row}', {'type': 'blanks', 'format': headerborder_format})
            
            # Apply bottom borders to A5:F5 given they are blank OR not blank
            worksheet.conditional_format(f'A{start_row + 4}:F{start_row + 4}', {'type': 'no_blanks', 'format': bottomborder_format})
            worksheet.conditional_format(f'A{start_row + 4}:F{start_row + 4}', {'type': 'blanks', 'format': bottomborder_format})

            # Apply right borders to G2:G46 given they are blank OR not blank
            worksheet.conditional_format(f'G{start_row + 1}:G{start_row + 3}', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format(f'G{start_row + 1}:G{start_row + 3}', {'type': 'blanks', 'format': rightborder_format})
            
            # Apply corner border to G5 given it is blank OR not blank
            worksheet.conditional_format(f'G{start_row + 4}', {'type': 'no_blanks', 'format': cornerborder_format})
            worksheet.conditional_format(f'G{start_row + 4}', {'type': 'blanks', 'format': cornerborder_format})
         

     
             
            acids_summary_data = []
            # Create an additional sheet and write additional data
            extra_sheetacids = {'Sample ID': ["='Analysis'!A2"]
                                }
            additional_df = pd.DataFrame(extra_sheetacids)
            additional_df.to_excel(writer, index=False, sheet_name='Summary Analysis', startrow=0, startcol=0)
            # Get the worksheet for the Cresol Testing sheet and tell it to start writing below column
            worksheet_summary = writer.sheets['Summary Analysis']
        
            # Write new column and row headers in Summary Analysis sheet
            headers = ['Sample ID', 'CAN mol/kg', 'TAN mol/kg', 'PhAN mol/kg', 'Carboxylic Acid Number mg KOH/g', 'Total Acid Number mg KOH/g', 'Phenolic Acid Number mg KOH/g']
            for i, header in enumerate(headers):
                worksheet_summary.write(0, i, header, bold_format)
        
            #acids_summary_data = []             
            # Write Summary Analysis row headers or formulas below the columns
            #using a " ' ' " method because of formula to call Analysis page
            acids_summary_template = [
                ['',"='Analysis'!B4","='Analysis'!C4", "='Analysis'!D4", "='Analysis'!E4", "='Analysis'!F4","='Analysis'!G4"], 
            ]
            for row_num, row_data in enumerate(acids_summary_template):
                for col_num, cell_data in enumerate(row_data):
                          worksheet_summary.write(row_num + 1, col_num + 0, cell_data)
            summary_row = 0                
            #set the file to open the column width to the length of a string
            summaryacidscolumnwidth = len("Carboxylic Acid Number mg K")
            #worksheet.set_column(first_col, last_col, width, cell_format, options)
            worksheet_summary.set_column(4, 6, summaryacidscolumnwidth)
        
            #set the file to open the column width to the length of a string
            summaryacidscolumnwidth = len("PhAN mol/kg654")
            #worksheet.set_column(first_col, last_col, width, cell_format, options)
            worksheet_summary.set_column(1, 3, summaryacidscolumnwidth)
        
            # Replicate kf_calc based on the number of samples
            for sample_num in range(1, num_request + 1):
        # Create a copy of the sample template for this sample
                acids_summary_calc = [row[:] for row in acids_summary_template]

        # Calculate the start row for this sample
                start_row = (sample_num - 1) * 6 + 1

        # Update the formulas for this sample
                acids_summary_calc[summary_row][0] = f"='Analysis'!A{start_row + 1}"
                acids_summary_calc[summary_row][1] = f"='Analysis'!B{start_row + 3}"
                acids_summary_calc[summary_row][2] = f"='Analysis'!C{start_row + 3}"
                acids_summary_calc[summary_row][3] = f"='Analysis'!D{start_row + 3}"
                acids_summary_calc[summary_row][4] = f"='Analysis'!E{start_row + 3}"
                acids_summary_calc[summary_row][5] = f"='Analysis'!F{start_row + 3}"
                acids_summary_calc[summary_row][6] = f"='Analysis'!G{start_row + 3}"

                acids_summary_data.extend(acids_summary_calc)
        
        # Convert kf_data into a DataFrame for each sample
            df = pd.DataFrame(acids_summary_data, columns=columns)
            # Get the worksheet for the Cresol Testing sheet
            worksheet_summary = writer.sheets['Summary Analysis']
            # Write the DataFrame to the default sheet
            df.to_excel(writer, index=False, sheet_name='Summary Analysis', engine='xlsxwriter')
            
            # Create an additional sheet and write additional data
            acids_vanillic = {'Expected': [5.77, 11.54]}
                                
            additional_df = pd.DataFrame(acids_vanillic)
            additional_df.to_excel(writer, index=False, sheet_name='Vanillic Validation', startrow=0, startcol=0)
            worksheet_vanillic = writer.sheets['Vanillic Validation']
        
            # Write new column and row headers in Summary Analysis sheet
            vanillic_headers = ['Expected', 'Measured', '% Diff']
            for col_num, header in enumerate(vanillic_headers):
                worksheet_vanillic.write(0, col_num, header, bold_format)
        
            acids_vanillic_calc = [
                        ['', '', '=(A2-B2)/A2 * 100'],
                        ['', '', '=(A3-B3)/A3 * 100']]
                            
            for row_num, row_data in enumerate(acids_vanillic_calc):
                for col_num, cell_data in enumerate(row_data):
                          worksheet_vanillic.write(row_num + 1, col_num + 0, cell_data)
                          
            #format conditional for vanillic check        
            worksheet_vanillic.conditional_format('C2', {'type':     'cell',
                                           'criteria': 'between',
                                           'minimum':  5.4815,
                                           'maximum':  6.0585,
                                           'format':   green_format})
            
            worksheet_vanillic.conditional_format('C2', {'type':     'cell',
                                           'criteria': 'not between',
                                           'minimum':  5.4815,
                                           'maximum':  6.0585,
                                           'format':   red_format})
            
            worksheet_vanillic.conditional_format('C3', {'type':     'cell',
                                           'criteria': 'between',
                                           'minimum':  10.963,
                                           'maximum':  12.117,
                                           'format':   green_format})
            
            worksheet_vanillic.conditional_format('C3', {'type':     'cell',
                                           'criteria': 'not between',
                                           'minimum':  10.963,
                                           'maximum':  12.117,
                                           'format':   red_format})
                        
            # Create the DataFrame for "Vanillic Validation" with the correct headers
            df_vanillic = pd.DataFrame(acids_vanillic, columns=vanillic_headers)
            # Write the DataFrame to the default sheet
           
            df_vanillic.to_excel(writer, index=False, sheet_name='Vanillic Validation', engine='xlsxwriter')

  
        df = pd.DataFrame(acids_data, columns=columns)
        worksheet = writer.sheets['Analysis']
        df.to_excel(writer, index=False, sheet_name='Analysis', engine='xlsxwriter')
    
    elif instrument == "Carbonyls Titration":
        carbonyl_data = []
        carbonyl_template = [
            ['Sample 1',''],
            ['Sample 1',''],
            ['Average','=AVERAGE(B2:B3)'], 
            ['Range', '=ABS(B2-B3)'],
            ]
        average_row = 2  # Row number for Mean%
        range_row = 3    # Row number for StDev%
        sample_row1 = 0 # Row number for "Sample X" labels
        sample_row2 = 1
        
        #set the file to open the column width to the length of a string
        carbonylcolumnwidth = len("Carbonyls mol/kg")
        #worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(0, 1, carbonylcolumnwidth) 
        
        # Replicate kf_calc based on the number of samples
        for sample_num in range(1, num_request + 1):
        # Create a copy of the sample template for this sample
            carbonyl_calc = [row[:] for row in carbonyl_template]

        # Calculate the start row for this sample
            start_row = (sample_num - 1) * 6 + 1 
            
       # Update the average forumlas 
            carbonyl_calc[average_row][1] = f'=AVERAGE(B{start_row + 1}:B{start_row + 2})'
            
       # Update the range forumlas      
            carbonyl_calc[range_row][1] = f'=ABS(B{start_row + 1}-B{start_row + 2})'
            
            carbonyl_calc[sample_row1][1] = 0

       # Update the Sample label row
            carbonyl_calc[sample_row1][0] = f'Sample {sample_num}'
            carbonyl_calc[sample_row2][0] = f'Sample {sample_num}'
       # Append the sample's data to kf_data
            carbonyl_data.extend(carbonyl_calc)

       # Add an empty row between sample templates, but not after the last sample
            if sample_num < num_request:
               carbonyl_data.extend([[''] * len(carbonyl_template[0]), columns])
            
                # Apply all borders to top row given they are blank OR not blank
            worksheet.conditional_format(f'A{start_row}:B{start_row}', {'type': 'no_blanks', 'format': headerborder_format})
            worksheet.conditional_format(f'A{start_row}:B{start_row}', {'type': 'blanks', 'format': headerborder_format})
            
            # Apply bottom borders to A5 given they are blank OR not blank
            worksheet.conditional_format(f'A{start_row + 4}', {'type': 'no_blanks', 'format': bottomborder_format})
            worksheet.conditional_format(f'A{start_row + 4}', {'type': 'blanks', 'format': bottomborder_format})

            # Apply right borders to B2:B4 given they are blank OR not blank
            worksheet.conditional_format(f'B{start_row + 1}:B{start_row + 3}', {'type': 'no_blanks', 'format': rightborder_format})
            worksheet.conditional_format(f'B{start_row + 1}:B{start_row + 3}', {'type': 'blanks', 'format': rightborder_format})
            
            # Apply corner border to B5 given it is blank OR not blank
            worksheet.conditional_format(f'B{start_row + 4}', {'type': 'no_blanks', 'format': cornerborder_format})
            worksheet.conditional_format(f'B{start_row + 4}', {'type': 'blanks', 'format': cornerborder_format})
     
             
            carbonyl_summary_data = []
            # Create an additional sheet and write additional data
            extra_sheetcarbonyl = {'Sample ID': ["='Analysis'!B4"]
                                }
            additional_df = pd.DataFrame(extra_sheetcarbonyl)
            additional_df.to_excel(writer, index=False, sheet_name='Summary Analysis', startrow=0, startcol=0)
            # Get the worksheet for the Cresol Testing sheet and tell it to start writing below column
            worksheet_summary = writer.sheets['Summary Analysis']
            
            #set the file to open the column width to the length of a string
            summarycarbonylcolumnwidth = len("Carbonyls mol/kg")
            #worksheet.set_column(first_col, last_col, width, cell_format, options)
            worksheet_summary.set_column(0, 1, summarycarbonylcolumnwidth)
        
            # Write new column and row headers in Summary Analysis sheet
            headers = ['Sample ID', 'Carbonyls mol/kg']
            for i, header in enumerate(headers):
                worksheet_summary.write(0, i, header, bold_format)
        
            #acids_summary_data = []             
            # Write Summary Analysis row headers or formulas below the columns
            #using a " ' ' " method because of formula to call Analysis page
            carbonyl_summary_template = [
                ['',"='Analysis'!B4"], 
            ]
            for row_num, row_data in enumerate(carbonyl_summary_template):
                for col_num, cell_data in enumerate(row_data):
                          worksheet_summary.write(row_num + 1, col_num + 0, cell_data)
            summary_row = 0                
        
            # Replicate kf_calc based on the number of samples
            for sample_num in range(1, num_request + 1):
        # Create a copy of the sample template for this sample
                carbonyl_summary_calc = [row[:] for row in carbonyl_summary_template]

        # Calculate the start row for this sample
                start_row = (sample_num - 1) * 6 + 1

        # Update the formulas for this sample
                carbonyl_summary_calc[summary_row][0] = f"='Analysis'!A{start_row + 1}"
                carbonyl_summary_calc[summary_row][1] = f"='Analysis'!B{start_row + 3}"


                carbonyl_summary_data.extend(carbonyl_summary_calc)
        
        # Convert kf_data into a DataFrame for each sample
            df = pd.DataFrame(carbonyl_summary_data, columns=columns)
            # Get the worksheet for the Cresol Testing sheet
            worksheet_summary = writer.sheets['Summary Analysis']
            # Write the DataFrame to the default sheet
            df.to_excel(writer, index=False, sheet_name='Summary Analysis', engine='xlsxwriter')
            
            # Create an additional sheet and write additional data
            carbonyl_4BBA = {'Expected 4-BBA Carbonyls mol/kg': [4.7, 4.7]}
                                
            additional_df = pd.DataFrame(carbonyl_4BBA)
            additional_df.to_excel(writer, index=False, sheet_name='4-BBA Validation', startrow=0, startcol=0)
            worksheet_4BBA = writer.sheets['4-BBA Validation']
            
            carbonyl4BBAcolumnwidth = len("Measured 4-BBA Carbonyls mol/kg")
            #worksheet.set_column(first_col, last_col, width, cell_format, options)
            worksheet_4BBA.set_column(0, 1, carbonyl4BBAcolumnwidth) 
            
            # Write new column and row headers in Summary Analysis sheet
            carbonyl_headers = ['Expected 4-BBA Carbonyls mol/kg', 'Measured 4-BBA Carbonyls mol/kg', '% Diff', 'Range']
            for col_num, header in enumerate(carbonyl_headers):
                worksheet_4BBA.write(0, col_num, header, bold_format)
        
            carbonyl_4BBA_calc = [
                        ['', '', '=ABS(A2-B2)/A2 * 100', '=ABS(B2-B3)'],
                        ['', '', '=ABS(A3-B3)/A3 * 100', '']]
                            
            for row_num, row_data in enumerate(carbonyl_4BBA_calc):
                for col_num, cell_data in enumerate(row_data):
                          worksheet_4BBA.write(row_num + 1, col_num + 0, cell_data)
                          
            #format conditional for vanillic check        
            worksheet_4BBA.conditional_format('C2', {'type':     'cell',
                                           'criteria': 'between',
                                           'minimum':  0,
                                           'maximum':  12.766,
                                           'format':   green_format})
            
            worksheet_4BBA.conditional_format('C2', {'type':     'cell',
                                           'criteria': 'not between',
                                           'minimum':  0,
                                           'maximum':  12.766,
                                           'format':   red_format})
            
            worksheet_4BBA.conditional_format('C3', {'type':     'cell',
                                           'criteria': 'between',
                                           'minimum':  0,
                                           'maximum':  12.766,
                                           'format':   green_format})
            
            worksheet_4BBA.conditional_format('C3', {'type':     'cell',
                                           'criteria': 'not between',
                                           'minimum':  0,
                                           'maximum':  12.766,
                                           'format':   red_format})
            worksheet_4BBA.conditional_format('D2', {'type':     'cell',
                                           'criteria': 'between',
                                           'minimum':  0,
                                           'maximum':  0.4,
                                           'format':   green_format})
            
            worksheet_4BBA.conditional_format('D2', {'type':     'cell',
                                           'criteria': 'not between',
                                           'minimum':  0,
                                           'maximum':  0.4,
                                           'format':   red_format})
                        
            # Create the DataFrame for "Vanillic Validation" with the correct headers
            df_4BBA = pd.DataFrame(carbonyl_4BBA, columns=carbonyl_headers)
            # Write the DataFrame to the default sheet
           
            df_4BBA.to_excel(writer, index=False, sheet_name='4-BBA Validation', engine='xlsxwriter')

  
        df = pd.DataFrame(carbonyl_data, columns=columns)
        worksheet = writer.sheets['Analysis']
        df.to_excel(writer, index=False, sheet_name='Analysis', engine='xlsxwriter')
                
    # Close the writer and save the Excel file
    writer.close()
    return filename



# Main app
def main():
    st.title('Instrument Template Generator')

    # Display a dropdown to select the instrument
    instrument = st.selectbox('Select an instrument', list(instrument_columns.keys()))
    
    # Allow the user to input the number of samples, step = increasing intervals of 1
    num_request = st.number_input('Number of Samples', min_value=1, value=1, step=1)
    # Generate the Excel file when a button is clicked
    if st.button('Generate Excel'):
        filename = generate_excel_file(instrument, num_request)
        st.success(f'Excel file for {instrument} with {num_request} samples has been generated!')

        # Provide a download link to the file
        with open(filename, 'rb') as f:
            data = f.read()
        b64 = base64.b64encode(data).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">Download Excel File</a>'
        st.markdown(href, unsafe_allow_html=True)

if __name__ == '__main__':
    main()
    
    
