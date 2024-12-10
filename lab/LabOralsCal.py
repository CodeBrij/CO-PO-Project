import re
from openpyxl.styles import *
from openpyxl import *

# Load the workbook and select the sheet
workbook = load_workbook('./Oral_Records.xlsx')
sheet = workbook['Orals']

# Find the roll count by identifying the row where "AL" is located
# Extract target value from cell A2
target = sheet['A2'].value
match = re.search(r'\d+(\.\d+)?', target)
if match:
    targetvalue = float(match.group())
totalCOs = sheet['B2'].value
total_roll = 0
startRow = 4
endRow = 0
startCal = 0
for i in range(1, 200):  # Corrected to iterate over a range of rows
    if sheet[f'B{i}'].value == "Count(appeared)":
        total_roll = i-5
        endRow =i-2
        startCal = i
        break  

print(total_roll, startRow, endRow, startCal, targetvalue, totalCOs)

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
sheet[f'B{startCal+5}'] = 'COs'
sheet[f'B{startCal+5}'].font = Font(bold=True)
sheet[f'B{startCal+5}'].border = thin_border

sheet[f'C{startCal+5}'] = 'AL'
sheet[f'C{startCal+5}'].font = Font(bold=True)
sheet[f'C{startCal+5}'].border = thin_border

for i in range(totalCOs):
    cell_b = sheet[f'B{startCal+6+i}']
    cell_b.value = f"CO{i+1}"
    cell_b.border = thin_border

    cell_c = sheet[f'C{startCal+6+i}']
    cell_c.value = sheet[f'C{startCal+3}'].value
    cell_c.border = thin_border

workbook.save('Calculated_Orals.xlsx')