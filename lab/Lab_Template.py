import openpyxl
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl import Workbook
from tkinter import filedialog
from tkinter import messagebox
import tkinter as tk
import os
from pathlib import Path
import customtkinter as ctk
from CTkMessagebox import CTkMessagebox

# Function to create a thin border


def lab_template_generator(basic_values_lo):

    def create_border():
        thin = Side(border_style="thin", color="000000")
        return Border(left=thin, right=thin, top=thin, bottom=thin)

# User inputs
    subject = basic_values_lo[0]
    division= basic_values_lo[1]
    branch = basic_values_lo[2]
    academic_year = basic_values_lo[3]
    semester = basic_values_lo[4]
    teacher_name = basic_values_lo[5]
    total_roll = int(basic_values_lo[6])
    LOcount = int(basic_values_lo[7])
    LabTarget = float(basic_values_lo[8])
    OralTarget = float(basic_values_lo[9])
    AssignmentTarget = float(basic_values_lo[10])
    ProjectTarget = float(basic_values_lo[11])
    lab_type = basic_values_lo[12]
    # miniProject = int(input("Mini Project? 0 or 1??"))
    assignmentCount = int(basic_values_lo[13])
    lo_text_array = basic_values_lo[14]

    print(assignmentCount)

    # AssignmentTarget = None
    # ProjectTarget = None

    # for i in range(LOcount):
    #     lo_text_array[i] = input(f"Enter the LO{i+1} description: ")
    # Non-Group Wise Template
    if lab_type == 'Individual Students':
        total_exp = int(basic_values_lo[15])
        LOs = basic_values_lo[16]
    else:
        total_exp = 0
        LOs = [0]

    # Group Wise Template
    if lab_type == 'Group Students':
        groupSize = int(basic_values_lo[17])
        # criteria = int(input("Enter the number of criteria for marks: "))
        critList = basic_values_lo[18]
        loList = basic_values_lo[19]
    else:
        groupSize = 0
        criteria = 0
        critList = [0]
        loList = [0]


    projGroupSize = int(basic_values_lo[20])
    # projCriteria = int(input("Enter the number of criteria for marks: "))
    # projCritList = [input(f"Enter criteria {i + 1}: ") for i in range(projCriteria)]
    projLoList = basic_values_lo[21]


    # assignmentCount = int(input("Enter the number of assignment: "))
    assignmentLOs = basic_values_lo[22]
        
# subject, total_roll, LOcount, LabTarget, lab_type, miniProject, assignment, total_exp, LOs, groupSize, criteria, critList, loList, projGroupSize, projCriteria, projCritList, projLoList, assignmentCount, assignmentLOs
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


    if (lab_type=="Individual Students"):
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
    if (lab_type=="Group Students"):
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

        for i in range(0,5):
            lab_sheet.cell(row=3, column=5 + i, value=critList[i])
            lab_sheet.cell(row=4, column=5 + i, value=loList[i])

        current_row = startCell + total_roll - 1
        lab_sheet[f'A{current_row+2}'] = f"Count>={LabTarget}%"
        lab_sheet[f'A{current_row+3}'] = f"%Count"
        lab_sheet[f'A{current_row+4}'] = "AL"

        for row in lab_sheet.iter_rows(min_row=3, max_row=current_row + 4, min_col=1, max_col=4 + 5):
            for cell in row:
                cell.border = create_border()

    ########################################################
    

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

    for i in range(0,4):
        project_sheet.cell(row=3, column=5 + i, value=f'Factor {i+1}')
        project_sheet.cell(row=4, column=5 + i, value=projLoList[i])

    current_row = startCell + total_roll - 1
    project_sheet[f'A{current_row+2}'] = f"Count>={ProjectTarget}%"
    project_sheet[f'A{current_row+3}'] = f"%Count"
    project_sheet[f'A{current_row+4}'] = "AL"

    # for row in project_sheet.iter_rows(min_row=3, max_row=current_row + 4, min_col=1, max_col=4+projCriteria):
    #     for cell in row:
    #         cell.border = create_border()

##############################
   
    
    # assignment_sheet=sheet4
    assignment_sheet = workbook.create_sheet(title="Assignment")
    assignment_sheet.column_dimensions['B'].width =42
    
    assignment_sheet['A2']="Roll No."
    assignment_sheet['B2']="Name"
    
    if assignmentCount==1 :
        assignment_sheet['C2']="Assignment1"
        assignment_sheet['C3']= assignmentLOs[0]
        assignment_sheet.merge_cells("A1:C1")
    
        myArr=['A','B','C']
        
    
    elif assignmentCount==2:
        assignment_sheet['C2']="Assignment1"
        assignment_sheet['D2']="Assignment2"
        
        assignment_sheet['C3']=assignmentLOs[0]
        assignment_sheet['D3']=assignmentLOs[1]
        assignment_sheet.merge_cells("A1:D1")

        myArr=['A','B', 'C', 'D']
        
    elif assignmentCount==3:
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
        for j,col in enumerate(myArr[:assignmentCount+2]) :
            assignment_sheet[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
            assignment_sheet[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
            if i>total_roll+3 :
                assignment_sheet[f'{col}{i}'].alignment= Alignment(horizontal='left', vertical='center')     
            
    for i in range(total_roll+4,total_roll+11):
        start_index=1
        for j, col in enumerate(myArr[start_index:assignmentCount+2],start=start_index+1) :
            assignment_sheet[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
            assignment_sheet[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))
        if i==total_roll+10:
            assignment_sheet[f'{col}{i}'].font=Font(bold=True)        

    print(assignmentCount)  

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
    lo_attainment_sheet = workbook.create_sheet(title="LO Attainment")


    lo_attainment_sheet.column_dimensions['A'].width =16
    lo_attainment_sheet.column_dimensions['B'].width =25
    lo_attainment_sheet.column_dimensions['C'].width =25
    lo_attainment_sheet.column_dimensions['D'].width =25
    lo_attainment_sheet.column_dimensions['E'].width =25
    lo_attainment_sheet.column_dimensions['F'].width =34
    lo_attainment_sheet.column_dimensions['G'].width =25
    
    for i in range (1,9):
        lo_attainment_sheet.merge_cells(f"A{i}:H{i}")
        lo_attainment_sheet[f'A{i}'].font=Font(bold=True)
        
    
    for i in range (9,15):
        lo_attainment_sheet.merge_cells(f"B{i}:H{i}") 
        
    

    lo_attainment_sheet["A1"].value="Vivekanand Education Society's Institute of Technology"
    lo_attainment_sheet["A1"].alignment= Alignment(horizontal='center', vertical='center')     
    
    lo_attainment_sheet["A2"].value="Department of "+branch+""
    lo_attainment_sheet["A2"].alignment= Alignment(horizontal='center', vertical='center')     
    
    lo_attainment_sheet["A3"].value="Academic Year :"+academic_year+""
    lo_attainment_sheet["A3"].alignment= Alignment(horizontal='center', vertical='center')     
    
    lo_attainment_sheet["A5"].value="  Subject : "+subject+"                                                                                                                                                                       Class : "+division+""
    lo_attainment_sheet["A5"].alignment= Alignment(horizontal='left', vertical='center')     
    
    lo_attainment_sheet["A6"].value="  Subject Teacher :"+teacher_name+"                                                                                                                                                                Semester : "+semester+""
    lo_attainment_sheet["A6"].alignment= Alignment(horizontal='left', vertical='center')     
    
    
    lo_attainment_sheet['A8']='Course Outcomes(COs): Upon successful completion of this course, students will be able to:'
    lo_attainment_sheet['A8'].font=Font(bold=True)
    lo_attainment_sheet["A8"].alignment= Alignment(horizontal='left', vertical='center')     
    
    lo_attainment_sheet['A9'] ='LO1'
    lo_attainment_sheet['A10']='LO2'
    lo_attainment_sheet['A11']='LO3'
    lo_attainment_sheet['A12']='LO4'
    lo_attainment_sheet['A13']='LO5'
    if(LOcount==6):
        lo_attainment_sheet['A14']='LO6'
    

    lo_attainment_sheet['B9'].value =""+lo_text_array[0]+""
    lo_attainment_sheet['B10'].value=""+lo_text_array[1]+""
    lo_attainment_sheet['B11'].value=""+lo_text_array[2]+""
    lo_attainment_sheet['B12'].value=""+lo_text_array[3]+""
    lo_attainment_sheet['B13'].value=""+lo_text_array[4]+""
    if(LOcount==6):
        lo_attainment_sheet['B14'].value=""+lo_text_array[5]+""
    rangeMax = 15
    if(LOcount==5):
        rangeMax = 14
    
    for i in range (9,rangeMax):
        lo_attainment_sheet[f'A{i}'].alignment= Alignment(horizontal='center', vertical='center')
        lo_attainment_sheet[f'B{i}'].alignment= Alignment(horizontal='left', vertical='center')         
    
    for i in range(9,rangeMax):
        for col in ['A','B', 'C', 'D', 'E', 'F', 'G','H']:
            lo_attainment_sheet[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))  
                    
    lo_attainment_sheet.merge_cells("A15:H15")
    lo_attainment_sheet.merge_cells("A16:H16")
    
    lo_attainment_sheet['A16']='CO Rubrics Mapping'
    lo_attainment_sheet['A16'].alignment= Alignment(horizontal='center', vertical='center')     
    
    lo_attainment_sheet.merge_cells("A17:H17")
    
    lo_attainment_sheet.merge_cells("A18:A19")  
    lo_attainment_sheet.merge_cells("F18:F19")
    
    lo_attainment_sheet.merge_cells("B18:E18")
    lo_attainment_sheet.merge_cells("B19:D19") 
    
    for i in range(16,21):
        for col in ['A','B', 'C', 'D', 'E', 'F', 'G','H']:
            lo_attainment_sheet[f'{col}{i}'].font=Font(bold=True)
    
    rangeMax2 = 27
    if(LOcount==5): rangeMax = 26
    for i in range(18,rangeMax2):
        for col in ['A','B', 'C', 'D', 'E', 'F']:
            lo_attainment_sheet[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000')) 
            lo_attainment_sheet[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
                
    lo_attainment_sheet['A18']='Assessment'
    lo_attainment_sheet['B18']='Direct Assessment' 
    lo_attainment_sheet['F18']='Indirect Assessment' 
    
    lo_attainment_sheet['B19']='Internal Assessment' 
    lo_attainment_sheet['E19']='External Assessment' 
    
    lo_attainment_sheet['A20']="LOs"
    lo_attainment_sheet['B20']="Lab Work"
    lo_attainment_sheet['C20']="Assignments"
    lo_attainment_sheet['D20']="Mini Project"
    lo_attainment_sheet['E20']="ESE(PR/OR)"
    lo_attainment_sheet['F20']="Course Exit Survey"
    
    lo_attainment_sheet['A21']='LO1'
    lo_attainment_sheet['A22']='LO2'
    lo_attainment_sheet['A23']='LO3'
    lo_attainment_sheet['A24']='LO4'
    lo_attainment_sheet['A25']='LO5'
    if(LOcount==6):
        lo_attainment_sheet['A26']='LO6'
    
    lo_attainment_sheet.merge_cells("A27:H27")
    lo_attainment_sheet.merge_cells("A28:H28")
    lo_attainment_sheet.merge_cells("A29:H29")
    
    lo_attainment_sheet['A28']='LO Attainment (Level)'
    lo_attainment_sheet['A28'].alignment= Alignment(horizontal='center', vertical='center')     
    
    lo_attainment_sheet.merge_cells("A30:A31")
    lo_attainment_sheet.merge_cells("G30:G31")
    
    lo_attainment_sheet.merge_cells("B30:F30")
    lo_attainment_sheet.merge_cells("B31:D31")
    
    for i in range(28,33):
        for col in ['A','B', 'C', 'D', 'E', 'F', 'G','H']:
            lo_attainment_sheet[f'{col}{i}'].font=Font(bold=True)
    rangeMax3 = 39
    if(LOcount==5): rangeMax3=38
    for i in range(30,rangeMax3):
        for col in ['A','B', 'C', 'D', 'E', 'F', 'G']:
            lo_attainment_sheet[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000')) 
            lo_attainment_sheet[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
    
    lo_attainment_sheet['A30']='Assessment'
    lo_attainment_sheet['B30']='Direct Assessment'
    lo_attainment_sheet['G30']='Indirect Assessment'
    
    lo_attainment_sheet['B31']='Internal Assessment'
    lo_attainment_sheet['E31']='External Assessment'
    lo_attainment_sheet['F31']='Attainment Level'
    
    lo_attainment_sheet['A32']="LOs"
    lo_attainment_sheet['B32']="Lab Work"
    lo_attainment_sheet['C32']="Assignments"
    lo_attainment_sheet['D32']="Mini Project"
    lo_attainment_sheet['E32']="ESE(PR/OR)"
    lo_attainment_sheet['F32']="70% (External) + 30% (Internal)"
    lo_attainment_sheet['G32']="Course Exit Survey"
    
    lo_attainment_sheet['A33']='LO1'
    lo_attainment_sheet['A34']='LO2'
    lo_attainment_sheet['A35']='LO3'
    lo_attainment_sheet['A36']='LO4'
    lo_attainment_sheet['A37']='LO5'
    if(LOcount==6):
        lo_attainment_sheet['A38']='LO6'
    
    lo_attainment_sheet.merge_cells("A39:H39")
    lo_attainment_sheet.merge_cells("A40:H40")
    lo_attainment_sheet.merge_cells("A41:H41")
    rangeMax4 = 49
    if(LOcount==5):
        rangeMax4 = 48
    for i in range(42,rangeMax4):
        for col in ['C', 'D']:
            lo_attainment_sheet[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000')) 
            lo_attainment_sheet[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')     
    
    lo_attainment_sheet['A40']='Final LO Attainment'
    lo_attainment_sheet['A40'].font=Font(bold=True)
    lo_attainment_sheet['A40'].alignment= Alignment(horizontal='center', vertical='center')     
    
    lo_attainment_sheet['C42']='Course Outcomes'
    lo_attainment_sheet['C42'].font=Font(bold=True)
    
    lo_attainment_sheet['D42']='Final LO Attainment Level'
    lo_attainment_sheet['D42'].font=Font(bold=True)
    
    lo_attainment_sheet['C43']='LO1'
    lo_attainment_sheet['C43'].font=Font(bold=True)
    
    lo_attainment_sheet['C44']='LO2'
    lo_attainment_sheet['C44'].font=Font(bold=True)
    
    lo_attainment_sheet['C45']='LO3'
    lo_attainment_sheet['C45'].font=Font(bold=True)
    
    lo_attainment_sheet['C46']='LO4'
    lo_attainment_sheet['C46'].font=Font(bold=True)
    
    lo_attainment_sheet['C47']='LO5'
    lo_attainment_sheet['C47'].font=Font(bold=True)
    
    if(LOcount == 6):
        lo_attainment_sheet['C48']='LO6'
        lo_attainment_sheet['C48'].font=Font(bold=True)

    

        
    #########################################
    #PO ATtainment#
    if True:
        po_attainment_sheet = workbook.create_sheet(title="PO Attainment")
        for i in range (1,9):
            po_attainment_sheet.merge_cells(f"A{i}:O{i}")
            po_attainment_sheet[f'A{i}'].font=Font(bold=True)
                    
        po_attainment_sheet["A1"].value="Vivekanand Education Society's Institute of Technology"
        po_attainment_sheet["A1"].alignment= Alignment(horizontal='center', vertical='center')     

        po_attainment_sheet["A2"].value="Department of "+branch+""
        po_attainment_sheet["A2"].alignment= Alignment(horizontal='center', vertical='center')     

        po_attainment_sheet["A3"].value="Academic Year :"+academic_year+""
        po_attainment_sheet["A3"].alignment= Alignment(horizontal='center', vertical='center')     

        po_attainment_sheet["A5"].value="  Subject : "+subject+"                                                                                                                                                                       Class : "+division+""
        po_attainment_sheet["A5"].alignment= Alignment(horizontal='left', vertical='center')     

        po_attainment_sheet["A6"].value="  Subject Teacher :"+teacher_name+"                                                                                                                                                                Semester : "+semester+""
        po_attainment_sheet["A6"].alignment= Alignment(horizontal='left', vertical='center')     

        po_attainment_sheet["A9"].value="Programme Outcomes(POs):"                                                                                                                       
        po_attainment_sheet["A9"].alignment= Alignment(horizontal='left', vertical='center')     
        po_attainment_sheet["A9"].font=Font(bold=True)
        po_attainment_sheet.merge_cells("A9:O9")

        po_attainment_sheet['A7'] = "Please fill up the CO - PO/PSO Mapping - Leave cell empty for no mapping"
        red_font = Font(color="FF0000") 
        po_attainment_sheet['A7'].font = red_font

        po_attainment_sheet.merge_cells("A10:O10")
        po_attainment_sheet["A10"].value="""PO1) Basic Engineering knowledge: An ability to apply the fundamental knowledge in mathematics, science and engineering to solve problems in Computer engineering.
        PO2) Problem Analysis: Identify, formulate, research literature and analyze computer engineering problems reaching substantiated conclusions using first principles of mathematics, natural sciences and computer engineering and sciences.
        PO3) Design/ Development of Solutions: Design solutions for complex computer engineering problems and design system components or processes that meet specified needs with appropriate consideration for public health and safety, cultural, societal and environmental considerations.
        PO4) Conduct investigations of complex engineering problems using research-based knowledge and research methods including design of experiments, analysis and interpretation of data and synthesis of information to provide valid conclusions
        PO5) Modern Tool Usage: Create, select and apply appropriate techniques, resources and modern computer engineering and IT tools including prediction and modeling to complex engineering activities with an understanding of the limitations. 
        PO6) The Engineer and Society: Apply reasoning informed by contextual knowledge to assess societal, health, safety, legal and cultural issues and the consequent responsibilities relevant to computer engineering practice.
        PO7) Environment and Sustainability: Understand the impact of professional computer engineering solutions in societal and environmental contexts and demonstrate knowledge of and need for sustainable development. 
        PO8) Ethics: Apply ethical principles and commit to professional ethics and responsibilities and norms of computer engineering practice.
        PO9) Individual and Team Work: Function effectively as an individual, and as a member or leader in diverse teams and in multidisciplinary settings. 
        PO10) Communication: Communicate effectively on complex engineering activities with the engineering community and with society at large, such as being able to comprehend and write effective reports and design documentation, make effective presentations and give and receive clear instructions 
        PO11) Project Management and Finance: Demonstrate knowledge and understanding of computer engineering and management principles and apply these to one's own work, as a member and leader in a team, to manage projects and in multidisciplinary environments.
        PO12) Life-long Learning: Recognize the need for and have the preparation and ability to engage in independent and lifelong learning in the broadest context of technological change.
        PSO1) Professional Skills - The ability to develop programs for computer based systems of varying complexity and domains using standard practices.
        PSO2) Successful Career - The ability to adopt skills, languages, environment and platforms for creating innovative carrier paths, being successful entrepreneurs or for pursuing higher studies."""    
        
        po_attainment_sheet["A12"].value="CO - PO/PSO Mapping"                                                                                                                       
        po_attainment_sheet["A12"].alignment= Alignment(horizontal='center', vertical='center')     
        po_attainment_sheet["A12"].font=Font(bold=True)
        po_attainment_sheet.merge_cells("A12:O12")
        

        
        po_attainment_sheet.merge_cells("B14:M14")
        po_attainment_sheet.merge_cells("N14:O14")
        po_attainment_sheet.merge_cells("A14:A15")
        
        COrange = 22
        if(LOcount == 5):
            COrange = 21
        for i in range(14,COrange):
            for col in ['A','B', 'C', 'D', 'E', 'F', 'G','H','I','J','K','L','M','N','O']:
                po_attainment_sheet[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))  
                po_attainment_sheet[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')  
            

        for col in ['B', 'C', 'D', 'E', 'F', 'G','H','I','J','K','L','M','N','O']:
            po_attainment_sheet[f'{col}15'].font=Font(bold=True)

        for i in range(14,COrange):
            po_attainment_sheet[f'A{i}'].font=Font(bold=True)

        po_attainment_sheet['A16']='LO1'
        po_attainment_sheet['A17']='LO2'
        po_attainment_sheet['A18']='LO3'
        po_attainment_sheet['A19']='LO4'
        po_attainment_sheet['A20']='LO5'
        if(LOcount==6):
            po_attainment_sheet['A21']='LO6' 

        for col,i in zip(['B', 'C', 'D', 'E', 'F', 'G','H','I','J','K','L','M'],range(1,13)):
            po_attainment_sheet[f'{col}15']=f'PO{i}'
        
        po_attainment_sheet['N15']='PSO1' 
        po_attainment_sheet['O15']='PSO2'  

        po_attainment_sheet['A14']='Course Outcomes'
        po_attainment_sheet['A14'].alignment= Alignment(horizontal='center', vertical='center',wrap_text=True)  ######  With Wrap Text
        po_attainment_sheet['B14']='Programme Outcomes' 
        po_attainment_sheet['B14'].font=Font(bold=True)
        po_attainment_sheet['N14']="PSOs" 
        po_attainment_sheet['N14'].font=Font(bold=True)
            
            
            
            
        po_attainment_sheet["A23"].value="Direct PO Attainment"                                                                                                                       
        po_attainment_sheet["A23"].alignment= Alignment(horizontal='center', vertical='center')     
        po_attainment_sheet["A23"].font=Font(bold=True)
        po_attainment_sheet.merge_cells("A23:O23")
        

        
        po_attainment_sheet.merge_cells("B25:M25")
        po_attainment_sheet.merge_cells("N25:O25")
        po_attainment_sheet.merge_cells("A25:A26")
        
        COrange = 33
        if(LOcount == 5):
            COrange = 32
        for i in range(25,COrange):
            for col in ['A','B', 'C', 'D', 'E', 'F', 'G','H','I','J','K','L','M','N','O']:
                po_attainment_sheet[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))  
                po_attainment_sheet[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')  
            

        for col in ['B', 'C', 'D', 'E', 'F', 'G','H','I','J','K','L','M','N','O']:
            po_attainment_sheet[f'{col}26'].font=Font(bold=True)

        for i in range(25,33):
            po_attainment_sheet[f'A{i}'].font=Font(bold=True)

        po_attainment_sheet['A27']='LO1'
        po_attainment_sheet['A28']='LO2'
        po_attainment_sheet['A29']='LO3'
        po_attainment_sheet['A30']='LO4'
        po_attainment_sheet['A31']='LO5'
        if(LOcount==6):
            po_attainment_sheet['A32']='CO6' 

        for col,i in zip(['B', 'C', 'D', 'E', 'F', 'G','H','I','J','K','L','M'],range(1,13)):
            po_attainment_sheet[f'{col}26']=f'PO{i}'
        
        po_attainment_sheet['N26']='PSO1' 
        po_attainment_sheet['O26']='PSO2'  

        po_attainment_sheet['A25']='Course Outcomes(COs)'
        po_attainment_sheet['A25'].alignment= Alignment(horizontal='center', vertical='center',wrap_text=True)  ######  With Wrap Text
        po_attainment_sheet['B25']='Programme Outcomes(POs)' 
        po_attainment_sheet['B25'].font=Font(bold=True)
        po_attainment_sheet['N25']="PSOs"
        po_attainment_sheet['N25'].font=Font(bold=True)



            
        po_attainment_sheet["A34"].value="Direct PO Attainment (After Applying CO-PO Mapping)"                                                                                                                       
        po_attainment_sheet["A34"].alignment= Alignment(horizontal='center', vertical='center')     
        po_attainment_sheet["A34"].font=Font(bold=True)
        po_attainment_sheet.merge_cells("A34:O34")
        

        
        po_attainment_sheet.merge_cells("B36:M36")
        po_attainment_sheet.merge_cells("N36:O36")
        po_attainment_sheet.merge_cells("A36:A37")
        
        for i in range(36,45):
            for col in ['A','B', 'C', 'D', 'E', 'F', 'G','H','I','J','K','L','M','N','O']:
                po_attainment_sheet[f'{col}{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))  
                po_attainment_sheet[f'{col}{i}'].alignment= Alignment(horizontal='center', vertical='center')  
            

        for col in ['B', 'C', 'D', 'E', 'F', 'G','H','I','J','K','L','M','N','O']:
            po_attainment_sheet[f'{col}37'].font=Font(bold=True)
            po_attainment_sheet[f'{col}44'].font=Font(bold=True)

        for i in range(36,45):
            po_attainment_sheet[f'A{i}'].font=Font(bold=True)

        po_attainment_sheet['A38']='LO1'
        po_attainment_sheet['A39']='LO2'
        po_attainment_sheet['A40']='LO3'
        po_attainment_sheet['A41']='LO4'
        po_attainment_sheet['A42']='LO5'
        if(LOcount == 6):
            po_attainment_sheet['A43']='LO6' 
        po_attainment_sheet['A44']='Avg PO'

        for col,i in zip(['B', 'C', 'D', 'E', 'F', 'G','H','I','J','K','L','M'],range(1,13)):
            po_attainment_sheet[f'{col}37']=f'PO{i}'
        
        po_attainment_sheet['N37']='PSO1' 
        po_attainment_sheet['O37']='PSO2'  

        po_attainment_sheet['A36']='Course Outcomes(COs)'
        po_attainment_sheet['A36'].alignment= Alignment(horizontal='center', vertical='center',wrap_text=True)  ######  With Wrap Text
        po_attainment_sheet['B36']='Programme Outcomes(POs)' 
        po_attainment_sheet['B36'].font=Font(bold=True)
        po_attainment_sheet['N36']="PSOs"
        po_attainment_sheet['N36'].font=Font(bold=True)

        po_attainment_sheet['A47']='AL' 
        po_attainment_sheet['A47'].font=Font(bold=True)
        po_attainment_sheet['B47']='%'
        po_attainment_sheet['B47'].font=Font(bold=True)
        po_attainment_sheet['A48']='1' 
        po_attainment_sheet['A49']='2' 
        po_attainment_sheet['A50']='3' 
        po_attainment_sheet['B48']='40' 
        po_attainment_sheet['B49']='60' 
        po_attainment_sheet['B50']='100'

        for i in range(47,51):
                po_attainment_sheet[f'A{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))  
                po_attainment_sheet[f'A{i}'].alignment= Alignment(horizontal='center', vertical='center')  
                po_attainment_sheet[f'B{i}'].border=Border(top=Side(style='thin',color='000000'),right=Side(style='thin',color='000000'),left=Side(style='thin',color='000000'),bottom=Side(style='thin',color='000000'))  
                po_attainment_sheet[f'B{i}'].alignment= Alignment(horizontal='center', vertical='center')  
            
    #######################################3
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
    optSheet['A11'] = "Exp Count"
    
    optSheet['B3'] = total_roll
    optSheet['B4'] = LOcount
    optSheet['B5'] = LabTarget
    optSheet['B6'] = OralTarget
    optSheet['B7'] = AssignmentTarget
    optSheet['B8'] = ProjectTarget
    optSheet['B9'] = lab_type
    optSheet['B10'] = assignmentCount
    optSheet['B11'] = total_exp

    selectedPath = filedialog.askdirectory()
    filepath = f'{selectedPath}/Lab_Template_{subject}_{division}_{teacher_name}_{academic_year}.xlsx'

    # Save the workbook
    workbook.save(filepath)
    print(f"Workbook saved successfully as {subject}_Lab_Template.xlsx")
    CTkMessagebox(message=f"Excel template downloaded successfully at {filepath}.",icon="check", option_1="OK")
