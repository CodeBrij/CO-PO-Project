import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

# Function to create a thin border
def create_border():
    thin = Side(border_style="thin", color="000000")
    return Border(left=thin, right=thin, top=thin, bottom=thin)

subject = input("Enter the subject name: ")
total_roll = int(input("Enter the total number of roll numbers: "))
total_COs = int(input("Enter the total number of COs: "))
LabTarget = float(input("Enter the Lab Target: "))

# Create workbook and sheet
workbook = Workbook()
sheet1 = workbook.active
sheet1.title = "Orals"

# Setting the Subject Lab Work as the heading in A1
sheet1['A1'] = f"{subject} Orals"
sheet1['A1'].font = Font(size=14, bold=True)  # Make the heading bold and larger
sheet1['A1'].alignment = Alignment(horizontal='center')  # Center align the heading
sheet1['A1'].border = create_border()  # Add border to heading

# Merging the heading across columns for better formatting
sheet1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_COs+3)

# Setting the Lab Target value in A2
sheet1['A2'] = f"Target = {LabTarget}%"
sheet1['A2'].border = create_border()  # Add border to Lab Target

sheet1['B2'] = total_COs

# Header row setup
sheet1['A3'] = "Roll No."
sheet1['A3'].border = create_border()  # Add border to header
sheet1['B3'] = "Name"
sheet1['B3'].border = create_border()  # Add border to header
sheet1['C3'] = "Marks(25)"
sheet1['C3'].border = create_border()  # Add border to header

for i in range(total_roll):
    cell = sheet1[f'A{i+4}']
    cell.value = i + 1
    cell.border = create_border()  # Add border to roll number cells

endCol = i + 4

# Adding footer information and applying borders
footer_info = [
    ("Count(appeared)", f'B{endCol+2}'),
    (f"Count(>={LabTarget}%)", f'B{endCol+3}'),
    (f"% count(>={LabTarget}%) w.r.t appeared", f'B{endCol+4}'),
    ("AL (All Los)", f'B{endCol+5}')
]

for text, position in footer_info:
    cell = sheet1[position]
    cell.value = text
    cell.border = create_border()  # Add border to footer cells

# Save the workbook
workbook.save("Oral_Records.xlsx")
