import re
from openpyxl.styles import *
from openpyxl import *
from openpyxl.utils import *

def cal_lab_sheets() :

    workbook = load_workbook('DSA_Lab_Template.xlsx')
    # List all sheet names
    sheet_names = workbook.sheetnames

    # Print the sheet names
    print(sheet_names)
    optSheet = workbook["Optional"]

    total_roll = int(optSheet['B3'].value)
    LOcount = int(optSheet['B4'].value)
    LabTarget = float(optSheet['B5'].value)
    OralTarget = float(optSheet['B6'].value)
    AssignmentTarget = float(optSheet['B7'].value)
    ProjectTarget = float(optSheet['B8'].value)
    lab_type = optSheet['B9'].value
    assignmentCount = int(optSheet['B10'].value)
    assignment_col = [get_column_letter(i) for i in range(3, 3 + assignmentCount)]
    # expCount = int(optSheet['B11'].value)

    def cal_orals(sheet):
        targetvalue = OralTarget
            
        startRow = 4
        endRow = 0
        startCal = 0
        for i in range(1, 200):  # Corrected to iterate over a range of rows
            if sheet[f'B{i}'].value == "Count(appeared)":
                endRow =i-2
                startCal = i
                break  

        print(total_roll, startRow, endRow, startCal, targetvalue)

        # Calc. started
        sheet[f'C{startCal}'] = total_roll
        target = targetvalue*25/100
        sheet[f'C{startCal+1}'] = f'=COUNTIF(C{startRow}:C{endRow},">={target}")'
        sheet[f'C{startCal+2}'] = f'=ROUND(C{startCal+1}/C{startCal},1)*100'
        sheet[f'C{startCal+3}'] = f'=IF(C{startCal+2}<60,1,IF(AND(C{startCal+2}>59,C{startCal+2}<70),2,IF(AND(C{startCal+2}>69,C{startCal+2}<80),3,4)))'

        # Define the border style
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Make the headings bold and give them borders
        sheet[f'B{startCal+5}'] = 'LOs'
        sheet[f'B{startCal+5}'].font = Font(bold=True)
        sheet[f'B{startCal+5}'].border = thin_border

        sheet[f'C{startCal+5}'] = 'AL'
        sheet[f'C{startCal+5}'].font = Font(bold=True)
        sheet[f'C{startCal+5}'].border = thin_border
        orals_lo_arr = []

        for i in range(LOcount):
            cell_b = sheet[f'B{startCal+6+i}']
            cell_b.value = f"LO{i+1}"
            cell_b.border = thin_border

            cell_c = sheet[f'C{startCal+6+i}']
            cell_c.value = sheet[f'C{startCal+3}'].value
            cell_c.border = thin_border
            orals_lo_arr.append(f'={sheet.title}!C{startCal+6+i}')
        
        return orals_lo_arr

    def cal_ungroup_labs(sheet):
        # Extract target value from cell A2
        targetvalue = LabTarget

        # Extract expCount value from cell B2
        expCount = sheet['B2'].value
        match = re.search(r'\d+', expCount)
        if match:
            total_exp = int(match.group())

        # Find the last non-empty row
        i = 10
        while sheet[f'A{i}'].value is not None:
            i += 1

        # Last roll number row found
        endRow = i - 1
        startRow = 6
        startCol = 'C'
        offset = total_exp

        # Calculate the new column letter
        endCol = chr(ord(startCol) + offset - 1)

        #alphabet array
        alphabets = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']

        print(startRow)
        print(endRow)
        print(startCol)
        print(endCol)
        print(total_exp)
        print(total_roll)
        # Calculate the values


        # Starting column for processing
        j = 3
        i = endRow + 2
        for col in range(j, j + total_exp):
            # Assign formula for COUNTIF
            sheet.cell(row=i, column=col, value=f"=COUNTIF({alphabets[col-1]}4:{alphabets[col-1]}{endRow},\">={10 * targetvalue / 100}\")")
            print(f"=COUNTIF({alphabets[col-1]}4:{alphabets[col-1]}{endRow},\">={10 * targetvalue / 100}")
        #     # Assign formula for ROUND percentage
            sheet.cell(row=i + 1, column=col, value=f"=ROUND(({sheet.cell(row=i, column=col).coordinate}/{total_roll})*100, 1)")

            # Assign formula for grading
            sheet.cell(row=i + 2, column=col, value=(
                f"=IF({sheet.cell(row=i + 1, column=col).coordinate}<60,1,"
                f"IF(AND({sheet.cell(row=i + 1, column=col).coordinate}>=60,"
                f"{sheet.cell(row=i + 1, column=col).coordinate}<70),2,"
                f"IF(AND({sheet.cell(row=i + 1, column=col).coordinate}>=70,"
                f"{sheet.cell(row=i + 1, column=col).coordinate}<80),3,4)))"
            ))

        avgCol = chr(ord(startCol) + offset)
        print(avgCol)
        for i in range(total_roll):
            sheet[f'{avgCol}{i+6}'] = f'=ROUND(AVERAGE({startCol}{i+6}:{endCol}{i+6}),1)'

        los_dict = {}
            
        for col in range(ord(startCol), ord(endCol) + 1):
                current_col_letter = chr(col)
                cell_value = str(sheet[f'{current_col_letter}5'].value)
                
                if cell_value:  # If the cell has a value
                    values = cell_value.split(',')
                    for value in values:
                        value = value.strip()  # Remove any leading/trailing whitespace
                        if value.isnumeric():  # Ensure it's a numeric value
                            lo_key = f'LO{value}'
                            if lo_key not in los_dict:
                                los_dict[lo_key] = []
                            los_dict[lo_key].append(current_col_letter)
            
        print(los_dict)
        # Define the border style
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Define the font style for bold text
        bold_font = Font(bold=True)
        new_row = endRow+6
        # Set the headers
        sheet[f'B{new_row}'] = "LOs"
        sheet[f'C{new_row}'] = "AL"

        # Apply the border and bold font to the headers
        sheet[f'B{new_row}'].border = thin_border
        sheet[f'B{new_row}'].font = bold_font

        sheet[f'C{new_row}'].border = thin_border
        sheet[f'C{new_row}'].font = bold_font

        i = 1
        ungroup_labs_lo_arr = []
        # Iterate over each key in los_dict to create and write the formulas
        for key in los_dict.keys():
            columns = los_dict[key]  # Get the list of columns for the current LO
            column_ranges = ','.join([f'{col}{endRow+4}' for col in columns])  # Create the range for AVERAGE formula
            sheet[f'B{new_row+1}'] = key
            sheet[f'C{new_row+1}'] = f'=ROUND(AVERAGE({column_ranges}),1)'  # AVERAGE formula

            # Add border to the cells
            sheet[f'B{new_row+1}'].border = thin_border
            sheet[f'C{new_row+1}'].border = thin_border
            ungroup_labs_lo_arr.append(f'={sheet.title}!C{new_row+1}')
            
            new_row += 1  # Increment new_row for the next set of entries
        
        return ungroup_labs_lo_arr
    
    def cal_group_labs(sheet):
                
        # Find the roll count by identifying the row where "AL" is located
        rollCount = total_roll+8
        
        #alphabet array
        alphabets = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']

        startRow = 5 #start of marks row
        endRow = rollCount - 4 #end of marks row
        calRow = rollCount - 2 #calculation row started

        startCol = 'E'
        endCol = 0

        for col in range(1, sheet.max_column + 1):
            if sheet.cell(row=3, column=col).value == None:
                endCol = col - 1

        # Get the target value from cell A2
        targetvalue = LabTarget
        
        # Find the last column for processing (where data ends)
        endCol = None
        for j in range(1, 26):  # Iterate through columns in the sheet
            if sheet.cell(row=3, column=j).value is None:  # Find the first empty column in row 3
                endCol = j - 2  # The last filled column
                break

        # Getting the LOs from the LOs row

        los_dict = {}
            
        for col in range(ord(startCol), ord(alphabets[endCol]) + 1):
                current_col_letter = chr(col)
                cell_value = str(sheet[f'{current_col_letter}4'].value)
                
                if cell_value:  # If the cell has a value
                    values = cell_value.split(',')
                    for value in values:
                        value = value.strip()  # Remove any leading/trailing whitespace
                        if value.isnumeric():  # Ensure it's a numeric value
                            lo_key = f'LO{value}'
                            if lo_key not in los_dict:
                                los_dict[lo_key] = []
                            los_dict[lo_key].append(current_col_letter)
            
        print(los_dict)

        # Calc. values at End Col
        for col in range(ord(startCol), ord(alphabets[endCol]) + 1):
            current_col_letter = chr(col)
            sheet[f'{current_col_letter}{calRow}'] = f'=COUNTIF({current_col_letter}{startRow}:{current_col_letter}{endRow},">={5*targetvalue/100}")' 
            sheet[f'{current_col_letter}{calRow+1}'] = f'=ROUND((({current_col_letter}{calRow}/{total_roll})*100),1)' 
            sheet[f'{current_col_letter}{calRow+2}'] = f'=IF({current_col_letter}{calRow+1}<60,1,IF(AND({current_col_letter}{calRow+1}>59,{current_col_letter}{calRow+1}<70),2,IF(AND({current_col_letter}{calRow+1}>69,{current_col_letter}{calRow+1}<80),3,4)))'

        new_row = calRow + 4

        # Define the border style
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Define the font style for bold text
        bold_font = Font(bold=True)

        # Set the headers
        sheet[f'B{new_row}'] = "LOs"
        sheet[f'C{new_row}'] = "AL"

        # Apply the border and bold font to the headers
        sheet[f'B{new_row}'].border = thin_border
        sheet[f'B{new_row}'].font = bold_font

        sheet[f'C{new_row}'].border = thin_border
        sheet[f'C{new_row}'].font = bold_font

        i = 1
        group_labs_lo_arr = []
        # Iterate over each key in los_dict to create and write the formulas
        for key in los_dict.keys():
            columns = los_dict[key]  # Get the list of columns for the current LO
            column_ranges = ','.join([f'{col}{calRow+2}' for col in columns])  # Create the range for AVERAGE formula
            sheet[f'B{new_row+1}'] = key
            sheet[f'C{new_row+1}'] = f'=ROUND(AVERAGE({column_ranges}),1)'  # AVERAGE formula

            # Add border to the cells
            sheet[f'B{new_row+1}'].border = thin_border
            sheet[f'C{new_row+1}'].border = thin_border
            group_labs_lo_arr.append(f'={sheet.title}!C{new_row+1}')
            
            new_row += 1  # Increment new_row for the next set of entries
        
        return group_labs_lo_arr

    def cal_mini_project(sheet):
                        
        # Find the roll count by identifying the row where "AL" is located
        rollCount = total_roll+8
        
        #alphabet array
        alphabets = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']

        startRow = 5 #start of marks row
        endRow = rollCount - 4 #end of marks row
        calRow = rollCount - 2 #calculation row started

        startCol = 'E'
        endCol = 0

        for col in range(1, sheet.max_column + 1):
            if sheet.cell(row=3, column=col).value == None:
                endCol = col - 1

        # Get the target value from cell A2
        targetvalue = ProjectTarget
        
        # Find the last column for processing (where data ends)
        endCol = None
        for j in range(1, 26):  # Iterate through columns in the sheet
            if sheet.cell(row=3, column=j).value is None:  # Find the first empty column in row 3
                endCol = j - 2  # The last filled column
                break

        # Getting the LOs from the LOs row

        los_dict = {}
            
        for col in range(ord(startCol), ord(alphabets[endCol]) + 1):
                current_col_letter = chr(col)
                cell_value = str(sheet[f'{current_col_letter}4'].value)
                
                if cell_value:  # If the cell has a value
                    values = cell_value.split(',')
                    for value in values:
                        value = value.strip()  # Remove any leading/trailing whitespace
                        if value.isnumeric():  # Ensure it's a numeric value
                            lo_key = f'LO{value}'
                            if lo_key not in los_dict:
                                los_dict[lo_key] = []
                            los_dict[lo_key].append(current_col_letter)
            
        print(los_dict)

        # Calc. values at End Col
        for col in range(ord(startCol), ord(alphabets[endCol]) + 1):
            current_col_letter = chr(col)
            sheet[f'{current_col_letter}{calRow}'] = f'=COUNTIF({current_col_letter}{startRow}:{current_col_letter}{endRow},">={5*targetvalue/100}")' 
            sheet[f'{current_col_letter}{calRow+1}'] = f'=ROUND((({current_col_letter}{calRow}/{total_roll})*100),1)' 
            sheet[f'{current_col_letter}{calRow+2}'] = f'=IF({current_col_letter}{calRow+1}<60,1,IF(AND({current_col_letter}{calRow+1}>59,{current_col_letter}{calRow+1}<70),2,IF(AND({current_col_letter}{calRow+1}>69,{current_col_letter}{calRow+1}<80),3,4)))'

        new_row = calRow + 4

        # Define the border style
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Define the font style for bold text
        bold_font = Font(bold=True)

        # Set the headers
        sheet[f'B{new_row}'] = "LOs"
        sheet[f'C{new_row}'] = "AL"

        # Apply the border and bold font to the headers
        sheet[f'B{new_row}'].border = thin_border
        sheet[f'B{new_row}'].font = bold_font

        sheet[f'C{new_row}'].border = thin_border
        sheet[f'C{new_row}'].font = bold_font

        i = 1
        mini_project_lo_arr = []
        # Iterate over each key in los_dict to create and write the formulas
        for key in los_dict.keys():
            columns = los_dict[key]  # Get the list of columns for the current LO
            column_ranges = ','.join([f'{col}{calRow+2}' for col in columns])  # Create the range for AVERAGE formula
            sheet[f'B{new_row+1}'] = key
            sheet[f'C{new_row+1}'] = f'=ROUND(AVERAGE({column_ranges}),1)'  # AVERAGE formula

            # Add border to the cells
            sheet[f'B{new_row+1}'].border = thin_border
            sheet[f'C{new_row+1}'].border = thin_border
            mini_project_lo_arr.append(f'={sheet.title}!C{new_row+1}')
            
            new_row += 1  # Increment new_row for the next set of entries
        
        return mini_project_lo_arr

    def cal_assignment(sheet):
        for col in assignment_col:
        
            sheet[f'{col}{total_roll+4}'] = f'=COUNT({col}4:{col}{total_roll+3})'
            sheet[f'{col}{total_roll+5}'] = f'=ROUND(AVERAGE({col}4:{col}{total_roll+3}), 0)'
            target_cell = sheet[f'{col}{total_roll+6}']
            if target_cell.value is None:  # Check if the cell is empty
                target_cell.value = f'=COUNTIF({col}4:{col}{total_roll+3}, ">={float(AssignmentTarget) / 100 * 10}")'
            sheet[f'{col}{total_roll+7}'] = f'=ROUND({sheet[f"{col}{total_roll+6}"].coordinate} / {sheet[f"{col}{total_roll+4}"].coordinate} * 100, 1)'
            sheet[f'{col}{total_roll+8}'] = f'=COUNTIF({col}3:{col}{total_roll+3}, ">="&{col}{total_roll+5})'
            sheet[f'{col}{total_roll+9}'] = f'=ROUND({sheet[f"{col}{total_roll+8}"].coordinate} / {sheet[f"{col}{total_roll+4}"].coordinate} * 100, 1)'
            sheet[f'{col}{total_roll+10}'] = f'=IF({sheet[f"{col}{total_roll+7}"].coordinate}<60, 1, IF(AND({sheet[f"{col}{total_roll+7}"].coordinate}>59, {sheet[f"{col}{total_roll+7}"].coordinate}<70), 2, IF(AND({sheet[f"{col}{total_roll+7}"].coordinate}>69, {sheet[f"{col}{total_roll+7}"].coordinate}<80), 3, 4)))'
    
        print("Assignmet col: ", assignment_col)
        los_dict = {}
            
        for col in assignment_col:
                current_col_letter = col
                cell_value = sheet[f'{current_col_letter}3'].value
                
                if cell_value:  # If the cell has a value
                    values = cell_value.split(',')
                    for value in values:
                        value = value.strip()  # Remove any leading/trailing whitespace
                        if value.isnumeric():  # Ensure it's a numeric value
                            lo_key = f'LO{value}'
                            if lo_key not in los_dict:
                                los_dict[lo_key] = []
                            los_dict[lo_key].append(current_col_letter)
            
        print(los_dict)

        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        i = 0
        assignment_lo_arr = []
        for key in los_dict.keys():
            columns = los_dict[key]  # Get the list of columns for the current LO
            column_ranges = ','.join([f'{col}{total_roll+10}' for col in columns])  # Create the range for AVERAGE formula
            sheet[f'C{total_roll+14+i}'] = key
            sheet[f'D{total_roll+14+i}'] = f'=ROUND(AVERAGE({column_ranges}),1)'  # AVERAGE formula

            # Add border to the cells
            sheet[f'C{total_roll+14+i}'].border = thin_border
            sheet[f'D{total_roll+14+i}'].border = thin_border
            assignment_lo_arr.append(f'={sheet.title}!D{total_roll+14+i}')
            
            i += 1  # Increment new_row for the next set of entries

        return assignment_lo_arr
    
    def cal_survey(sheet):
        sheet4=workbook['Survey']

        col_list = ['G','H','I','J','K']

        if LOcount == 6:
            col_list.append('L')
        
        for col in col_list:
            
            sheet4[f'{col}{total_roll+4}'] = f'=COUNT({col}2:{col}{total_roll+1})'
            sheet4[f'{col}{total_roll+5}'] = f'=COUNTIF({col}2:{col}{total_roll+1}, ">=4")'
            sheet4[f'{col}{total_roll+6}'] = f'=ROUND(({col}{total_roll+5}/{col}{total_roll+4}*100), 1)'
            sheet4[f'{col}{total_roll+8}'] = f'=IF({col}{total_roll+6}<60,1,IF(AND({col}{total_roll+6}>59,{col}{total_roll+6}<70),2,IF(AND({col}{total_roll+6}>69,{col}{total_roll+6}<80),3,4)))'
            
        map_survey_lo_arr=[f'={sheet4.title}!G{total_roll+8}',f'={sheet4.title}!H{total_roll+8}',f'={sheet4.title}!I{total_roll+8}',f'={sheet4.title}!J{total_roll+8}',f'={sheet4.title}!K{total_roll+8}',f'={sheet4.title}!L{total_roll+8}']
        return map_survey_lo_arr


    orals_lo_value = []
    group_labs_lo_value = []
    nongroup_labs_lo_value = []
    mini_project_lo_value = []
    assignment_lo_value = []
    for i in range(len(sheet_names)):
        if(sheet_names[i] == "Orals"): orals_lo_value = cal_orals(workbook[sheet_names[i]]);
        if(sheet_names[i] == "Lab"): 
            if(lab_type=="group"):
                group_labs_lo_value = cal_group_labs(workbook[sheet_names[i]])
            else:
                nongroup_labs_lo_value = cal_ungroup_labs(workbook[sheet_names[i]])
        if(sheet_names[i] == "Mini Project"): mini_project_lo_value = cal_mini_project(workbook[sheet_names[i]])
        if(sheet_names[i] == "Assignment"): assignment_lo_value = cal_assignment(workbook[sheet_names[i]])
        if(sheet_names[i] == "Course Exit Survey"): survey_lo_value = cal_survey(workbook[sheet_names[i]])

    print("Final Vales -- ")
    print(orals_lo_value)
    print(group_labs_lo_value)
    print(nongroup_labs_lo_value)
    print(mini_project_lo_value)
    print(assignment_lo_value)
    
    # Save the modified workbook after calculations
    workbook.save('DSA_Lab_Calculated.xlsx')  # Save to a new file or overwrite the original
    print("Workbook saved successfully!")

cal_lab_sheets()