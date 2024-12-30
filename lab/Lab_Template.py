import openpyxl
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl import Workbook

# Function to create a thin border
def create_border():
    thin = Side(border_style="thin", color="000000")
    return Border(left=thin, right=thin, top=thin, bottom=thin)

# User inputs
subject = input("Enter the subject name: ")
division= input("Enter the division: ")
total_roll = int(input("Enter the total number of roll numbers: "))
LOcount = int(input("Enter the total number of LOs: "))
LabTarget = float(input("Enter the Lab Target: "))
OralTarget = float(input("Enter the oral Target: "))
AssignmentTarget = float(input("Enter the Assignment Target: "))
ProjectTarget = float(input("Enter the Mini Project Target: "))
lab_type = input("Enter the lab type ('non-group' or 'group'): ").strip().lower()
miniProject = bool(input("Mini Project? 0 or 1??"))
assignment = bool(input("Assignment hai? 0 or 1??"))

# Non-Group Wise Template
if lab_type == 'non-group':
    total_exp = int(input("Enter the total number of experiments: "))
    LOs = [input(f'Enter the LO for Exp{i+1}: ') for i in range(total_exp)]
else:
    total_exp = 0
    LOs = [0]

# Group Wise Template
if lab_type == 'group':
    groupSize = int(input("Enter the size of group: "))
    criteria = int(input("Enter the number of criteria for marks: "))
    critList = [input(f"Enter criteria {i + 1}: ") for i in range(criteria)]
    loList = [input(f"Enter the LO for criteria {i + 1}: ") for i in range(criteria)]
else:
    groupSize = 0
    criteria = 0
    critList = [0]
    loList = [0]

if(miniProject):
    projGroupSize = int(input("Enter the size of group: "))
    projCriteria = int(input("Enter the number of criteria for marks: "))
    projCritList = [input(f"Enter criteria {i + 1}: ") for i in range(projCriteria)]
    projLoList = [input(f"Enter the LO for criteria {i + 1}: ") for i in range(projCriteria)]

if(assignment):
    assignmentCount = int(input("Enter the number of assignment: "))
    assignmentLOs = [input(f"Enter LOs for assignment{i + 1}: ") for i in range(assignmentCount)]
     
# subject, total_roll, LOcount, LabTarget, lab_type, miniProject, assignment, total_exp, LOs, groupSize, criteria, critList, loList, projGroupSize, projCriteria, projCritList, projLoList, assignmentCount, assignmentLOs

def lab_template_generator():
    # Create workbook and sheets
    workbook = Workbook()
    orals_sheet = workbook.active
    orals_sheet.title = "Orals"

    # Common setup for Orals
    orals_sheet['A1'] = f"{subject} Orals"
    orals_sheet['A1'].font = Font(size=14, bold=True)  # Make the heading bold and larger
    orals_sheet['A1'].alignment = Alignment(horizontal='center')  # Center align the heading
    orals_sheet['A1'].border = create_border()  # Add border to heading
    orals_sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    orals_sheet['A2'] = f"Target = {OralTarget}%"
    orals_sheet['A2'].border = create_border()  # Add border to Lab Target
    orals_sheet['A3'] = "Roll No."
    orals_sheet['A3'].border = create_border()  # Add border to header
    orals_sheet['B3'] = "Name"
    orals_sheet['B3'].border = create_border()  # Add border to header
    orals_sheet['C3'] = "Marks(25)"
    orals_sheet['C3'].border = create_border()  # Add border to header

    for i in range(total_roll):
        cell = orals_sheet[f'A{i+4}']
        cell2 = orals_sheet[f'B{i+4}']
        cell3 = orals_sheet[f'C{i+4}']
        cell.value = i + 1
        
        cell.border = create_border()  # Add border to roll number cells
        cell2.border = create_border()  # Add border to roll number cells
        cell3.border = create_border()  # Add border to roll number cells

    endCol = i + 4

    footer_info = [
    ("Count(appeared)", f'B{endCol+2}', f'C{endCol+2}'),
    (f"Count(>={OralTarget}%)", f'B{endCol+3}', f'C{endCol+3}'),
    (f"% count(>={OralTarget}%) w.r.t appeared", f'B{endCol+4}', f'C{endCol+4}'),
    ("AL (All Los)", f'B{endCol+5}', f'C{endCol+5}')
    ]

    for text, position, position2 in footer_info:
        cell = orals_sheet[position]
        cell2 = orals_sheet[position]
        cell.value = text
        cell.border = create_border()  # Add border to footer cells
        cell2.border = create_border()  # Add border to footer cells


    if (lab_type=="non-group"):
        lab_sheet = workbook.create_sheet(title="Lab")
        lab_sheet['A1'] = f"{subject} Lab Work - Ungrouped"
        lab_sheet['A1'].font = Font(size=14, bold=True)
        lab_sheet['A1'].alignment = Alignment(horizontal='center')
        lab_sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_exp+6)
        lab_sheet['A2'] = f"Target = {LabTarget}"

        lab_sheet['A3'] = "Roll No."
        lab_sheet['B3'] = "Name"
        lab_sheet['B2'] = f"Total Experiment = {total_exp}"
        lab_sheet.insert_rows(4)
        lab_sheet['A5'] = "LO"
        lab_sheet.merge_cells('A5:B5')
        lab_sheet['A5'].font = Font(size=12, bold=True)
        lab_sheet['A5'].alignment = Alignment(horizontal='center')

        for i in range(total_exp):
            lab_sheet.cell(row=4, column=i+3, value=f"Exp {i+1}")
            lab_sheet.cell(row=5, column=i+3, value=LOs[i])

        lab_sheet.cell(row=4, column=total_exp+3, value="Average(15)")

        for i in range(6, total_roll + 6):
            lab_sheet[f'A{i}'] = i - 5

        lab_sheet.column_dimensions['A'].width = 10
        lab_sheet.column_dimensions['B'].width = 25

        roll_end = total_roll + 5
        lab_sheet[f'A{roll_end+2}'] = f'Count>={LabTarget}%'
        lab_sheet[f'A{roll_end+3}'] = f'%Count'
        lab_sheet[f'A{roll_end+4}'] = f'AL'

        for row in lab_sheet.iter_rows(min_row=3, max_row=roll_end+4, min_col=1, max_col=total_exp+3):
            for cell in row:
                cell.border = create_border()
                cell.alignment = Alignment(horizontal='center', vertical='center')

##############################################################################
    if (lab_type=="group"):
        lab_sheet = workbook.create_sheet(title="Lab")
        lab_sheet.merge_cells('A1:E1')
        lab_sheet['A1'] = f"{subject} Lab Work - Grouped"
        lab_sheet['A1'].font = Font(size=14, bold=True)
        lab_sheet['A1'].alignment = Alignment(horizontal='center')
        lab_sheet['A2'] = f"Target={LabTarget}"
        lab_sheet['A3'] = "Group No."
        lab_sheet['B3'] = "Roll No."
        lab_sheet['C3'] = "Name of Student"
        lab_sheet['D3'] = "Project Name"
        lab_sheet['A4'] = "LOs Mapped"
        lab_sheet.column_dimensions['C'].width = 30

        startCell = 5
        groupCount = 1

        for roll_no in range(1, total_roll + 1):
            current_row = startCell + roll_no - 1
            if (roll_no - 1) % groupSize == 0:
                lab_sheet[f"A{current_row}"] = groupCount
                lab_sheet.merge_cells(f'A{current_row}:A{min(current_row + groupSize - 1, startCell + total_roll - 1)}')
                lab_sheet.merge_cells(f'D{current_row}:D{min(current_row + groupSize - 1, startCell + total_roll - 1)}')
                groupCount += 1
            lab_sheet[f"B{current_row}"] = roll_no
            lab_sheet[f"C{current_row}"] = f"Student {roll_no}"

        for i in range(criteria):
            lab_sheet.cell(row=3, column=5 + i, value=critList[i])
            lab_sheet.cell(row=4, column=5 + i, value=loList[i])

        current_row = startCell + total_roll - 1
        lab_sheet[f'A{current_row+2}'] = f"Count>={LabTarget}%"
        lab_sheet[f'A{current_row+3}'] = f"%Count"
        lab_sheet[f'A{current_row+4}'] = "AL"

        for row in lab_sheet.iter_rows(min_row=3, max_row=current_row + 4, min_col=1, max_col=4 + criteria):
            for cell in row:
                cell.border = create_border()

    ########################################################
    
    if (miniProject):
        project_sheet = workbook.create_sheet(title="Mini Project")
        project_sheet.merge_cells('A1:E1')
        project_sheet['A1'] = f"{subject} Mini Project"
        project_sheet['A1'].font = Font(size=14, bold=True)
        project_sheet['A1'].alignment = Alignment(horizontal='center')
        project_sheet['A2'] = f"Target={LabTarget}"
        project_sheet['A3'] = "Group No."
        project_sheet['B3'] = "Roll No."
        project_sheet['C3'] = "Name of Student"
        project_sheet['D3'] = "Project Name"
        project_sheet['A4'] = "LOs Mapped"
        project_sheet.column_dimensions['C'].width = 30

        startCell = 5
        groupCount = 1

        for roll_no in range(1, total_roll + 1):
            current_row = startCell + roll_no - 1
            if (roll_no - 1) % projGroupSize == 0:
                project_sheet[f"A{current_row}"] = groupCount
                project_sheet.merge_cells(f'A{current_row}:A{min(current_row + projGroupSize - 1, startCell + total_roll - 1)}')
                project_sheet.merge_cells(f'D{current_row}:D{min(current_row + projGroupSize - 1, startCell + total_roll - 1)}')
                groupCount += 1
            project_sheet[f"B{current_row}"] = roll_no
            project_sheet[f"C{current_row}"] = f"Student {roll_no}"

        for i in range(projCriteria):
            project_sheet.cell(row=3, column=5 + i, value=projCritList[i])
            project_sheet.cell(row=4, column=5 + i, value=projLoList[i])

        current_row = startCell + total_roll - 1
        project_sheet[f'A{current_row+2}'] = f"Count>={ProjectTarget}%"
        project_sheet[f'A{current_row+3}'] = f"%Count"
        project_sheet[f'A{current_row+4}'] = "AL"

        for row in project_sheet.iter_rows(min_row=3, max_row=current_row + 4, min_col=1, max_col=4+projCriteria):
            for cell in row:
                cell.border = create_border()

##############################
    if (assignment):
        temp=len(assignmentLOs)
        # assignment_sheet=sheet4
        assignment_sheet = workbook.create_sheet(title="Assignment")
        assignment_sheet.column_dimensions['B'].width =42
        
        assignment_sheet['A2']="Roll No."
        assignment_sheet['B2']="Name"
        
        if temp==1 :
            assignment_sheet['C2']="Assignment1"
            assignment_sheet['C3']= assignmentLOs[0]
            assignment_sheet.merge_cells("A1:C1")
        
            myArr=['A','B','C']
            
        
        elif temp==2:
            assignment_sheet['C2']="Assignment1"
            assignment_sheet['D2']="Assignment2"
            
            assignment_sheet['C3']=assignmentLOs[0]
            assignment_sheet['D3']=assignmentLOs[1]
            assignment_sheet.merge_cells("A1:D1")

            myArr=['A','B', 'C', 'D']
            
        elif temp==3:
            assignment_sheet['C2']="Assignment1"
            assignment_sheet['D2']="Assignment2"
            assignment_sheet['E2']="Assignment3"
            
            assignment_sheet['C3']=assignmentLOs[0]
            assignment_sheet['D3']=assignmentLOs[1]
            assignment_sheet['E3']=assignmentLOs[2]
            assignment_sheet.merge_cells("A1:E1")
            myArr=['A','B', 'C', 'D', 'E']
        
        assignment_sheet['A1']="Type : Assignment                 Maximum Marks for each question = 10"
        assignment_sheet['A1'].font=Font(bold=True)
        
        for i in range(1,4):
            for col in  ['A','B', 'C', 'D', 'E', 'F', 'G','H','I', 'J','K','L'] :
                assignment_sheet[f'{col}{i}'].font=Font(bold=True)
                
        assignment_sheet.merge_cells('A3:B3')
        assignment_sheet['A3'] = "LO -->"
        assignment_sheet['A3'].font=Font(bold=True)
        
        
        for i in range(1 ,total_roll+1):
            assignment_sheet[f'A{i+3}']=i
        
        for i in range(total_roll+4,total_roll+12) :
            assignment_sheet.merge_cells(f'A{i}:B{i}')
            
        for i in range(1,total_roll+11):
            for j,col in enumerate(myArr[:temp+2]) :
                assignment_sheet[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
                assignment_sheet[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
                if i>total_roll+3 :
                    assignment_sheet[f'{col}{i}'].alignment= Alignment(horizontal='left', vertical='center')     
                
        for i in range(total_roll+4,total_roll+11):
            start_index=1
            for j, col in enumerate(myArr[start_index:temp+2],start=start_index+1) :
                assignment_sheet[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
                assignment_sheet[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
            if i==total_roll+10:
                assignment_sheet[f'{col}{i}'].font=Font(bold=True)        
                
    
        assignment_sheet[f'A{total_roll+4}']="Count(Attempted)"       
        assignment_sheet[f'A{total_roll+5}']="Average Marks"
        
    
        assignment_sheet[f'A{total_roll+6}']=f"Count(>={AssignmentTarget}%)"
        
        
        assignment_sheet[f'A{total_roll+7}']=f"% Count(>={AssignmentTarget}% w.r.t appeared)"

        assignment_sheet[f'A{total_roll+8}']="Count(>=Average Marks of class)"
        assignment_sheet[f'A{total_roll+9}']="% Count(>=Average Marks of class w.r.t appeared)"
        
        assignment_sheet[f'A{total_roll+10}']=f"AL(Based on >={AssignmentTarget}% Count) (All LOs)"
        assignment_sheet[f'A{total_roll+10}'].font=Font(bold=True)
        
        loTableEnd = total_roll+20
        if(LOcount == 5):
            loTableEnd = total_roll+19
        for i in range(total_roll+13,loTableEnd):
            assignment_sheet[f'C{i}'].alignment= Alignment(horizontal='center', vertical='center')     
            assignment_sheet[f'C{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
            assignment_sheet[f'D{i}'].alignment= Alignment(horizontal='center', vertical='center')     
            assignment_sheet[f'D{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
            
        assignment_sheet[f'C{total_roll+13}'] = "LOs"
        assignment_sheet[f'C{total_roll+13}'].font=Font(bold=True)
        assignment_sheet[f'D{total_roll+13}'] = "AL"
        assignment_sheet[f'D{total_roll+13}'].font=Font(bold=True)
        assignment_sheet[f'C{total_roll+14}'] = 'LO1'
        assignment_sheet[f'C{total_roll+15}'] = 'LO2'
        assignment_sheet[f'C{total_roll+16}'] = 'LO3'
        assignment_sheet[f'C{total_roll+17}'] = 'LO4'
        assignment_sheet[f'C{total_roll+18}'] = 'LO5'
        if(LOcount==6):
            assignment_sheet[f'C{total_roll+19}'] = 'LO6' 

    #########################################
    optSheet = workbook.create_sheet(title="Optional")

    optSheet['A1'] = subject
    optSheet['A3'] = "Total Students"
    optSheet['A4'] = "LOsCount"
    optSheet['A5'] = "Labs"
    optSheet['A6'] = "Orals"
    optSheet['A7'] = "Assignment"
    optSheet['A8'] = "MiniProject"
    optSheet['A9'] = "lab_type"
    optSheet['A10'] = "assignmentCount"
    
    optSheet['B3'] = total_roll
    optSheet['B4'] = LOcount
    optSheet['B5'] = LabTarget
    optSheet['B6'] = OralTarget
    optSheet['B7'] = AssignmentTarget
    optSheet['B8'] = ProjectTarget
    optSheet['B9'] = lab_type
    optSheet['B10'] = assignmentCount

    #######################################
    sheet5 = workbook.create_sheet(title="Course Exit Survey")
    sheet5.column_dimensions['B'].width =40
    sheet5.column_dimensions['C'].width =40
    sheet5.column_dimensions['F'].width =40
    
    sheet5['A1']="Sr. No."
    sheet5['B1']="Email Address"
    sheet5['C1']="Full name of Student"
    sheet5['D1']="Roll No."
    sheet5['E1']="Class"
    sheet5['F1']="Branch"
    sheet5['G1']="Q1"
    sheet5['H1']="Q2"
    sheet5['I1']="Q3"
    sheet5['J1']="Q4"
    sheet5['K1']="Q5"
    if(LOcount == 6):
        sheet5['L1']="Q6"

    col_list = ['A','B', 'C', 'D', 'E', 'F', 'G','H','I', 'J','K']

    if LOcount == 6:
        col_list.append('L')
    
    for col in col_list :
            sheet5[f'{col}1'].font=Font(bold=True)
            
    for i in range(1 ,total_roll+1):
        sheet5[f'A{i+1}']=i
        sheet5[f'E{i+1}']= division
        sheet5[f'F{i+1}']=subject
        
    for i in range(1,total_roll+2):
        for col in col_list :
            sheet5[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
            sheet5[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
    
    sheet5[f'F{total_roll+4}']= 'Total' 
    sheet5[f'F{total_roll+4}'].font=Font(bold=True) 
    sheet5[f'F{total_roll+5}']= 'SA + A Count'
    sheet5[f'F{total_roll+5}'].font=Font(bold=True)
    sheet5[f'F{total_roll+6}']= 'SA + A Percentage' 
    sheet5[f'F{total_roll+6}'].font=Font(bold=True)
    sheet5[f'F{total_roll+7}']= 'LO Mapped' 
    sheet5[f'F{total_roll+7}'].font=Font(bold=True)
    sheet5[f'F{total_roll+8}']= 'AL'
    
    col_list_2 = ['F','G','H','I','J','K']
    if LOcount == 6:
        col_list_2.append('L')

    for i in range(total_roll+4,total_roll+9):
        for col in col_list_2 :
            sheet5[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
            sheet5[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
            if col=="F":
                sheet5[f'F{i}'].alignment= Alignment(horizontal='left', vertical='center')     
    
            if i==total_roll+8:
                sheet5[f'{col}{i}'].font=Font(bold=True)
        
    sheet5[f'G{total_roll+7}']= 'LO1' 
    sheet5[f'H{total_roll+7}']= 'LO2' 
    sheet5[f'I{total_roll+7}']= 'LO3' 
    sheet5[f'J{total_roll+7}']= 'LO4' 
    sheet5[f'K{total_roll+7}']= 'LO5' 
    if(LOcount==6):
        sheet5[f'L{total_roll+7}']= 'LO6'
    

    #########################################
    lo_attainment = workbook.create_sheet(title="LO Attainment")
    
    # Save the workbook
    workbook.save(f"{subject}_Lab_Template.xlsx")
    print(f"Workbook saved successfully as {subject}_Lab_Template.xlsx")


lab_template_generator()
