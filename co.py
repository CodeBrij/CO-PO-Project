import customtkinter as ctk 
from CTkMessagebox import CTkMessagebox
from tkinter import filedialog
import re
import openpyxl
from openpyxl.styles import Alignment 
from openpyxl import Workbook
import threading
from PIL import Image, ImageTk  # Required for image handling
import os
import sys

# @Sairam Konar
# PPT ke COS na normally he le like Abhi kaise lete he CA1_Co_arr=[1,2,3,4,5,6] but PPT ke liye aise le  CA1_Co_arr=[[1,2],[3,4],[5],[6]] array of array where inside array is of group COs isse kiya hoga ki strcture maintain rahega.
# Yeah Errors Dikh rahe he woh error nahi he places leave kiye he udhar code daal

class User_mode:
    def __init__(self):
        self.app = None
        self.cursor = None
        self.open_main_page()

    def go_back(self, current_window):
            current_window.destroy()  # Close the current window
            self.__init__() 

    def open_co_window(self):
            self.app.destroy() 
            
            co_window = ctk.CTk()  # Close the current window

            screen_width=co_window.winfo_screenwidth()
            screen_height=co_window.winfo_screenheight()
       
        # Calculate the coordinates for centering the window
            x_position = 0
            y_position = 0

             # Create a new CO Calculations window
            co_window.title("CO Calculations")
            co_window.geometry(f"{screen_width}x{screen_height}+{x_position}+{y_position}")

            main_frame = ctk.CTkFrame(master=co_window)
            main_frame.pack(expand=True, fill="both", padx=10, pady=10)

    # Tabview inside the frame
            tabview = ctk.CTkTabview(main_frame, corner_radius=20)
            tabview.pack(expand=True, fill="both", padx=10, pady=5)

            tabview.add(" Instructions ") 
            tabview.add(" CO Information ") 
            tabview.add(" Basic Information ") 
            # add tab at the end
            tabview.add(" CO Mapping ")
            tabview.add(" Target level of tests ")
            # tabview.add(" Lab CO ")
            tabview.add(" Upload Excel File ") 

            def switch():
                if (entry1.get() == "" or yearDropDown.get() == "Select Year" or
                    entry8.get() == "Select Department" or entry2.get() == "Select Sem" or
                    entry3.get() == "Select Subject" or entry4.get() == "" or
                    entry5.get() == "" or entry7.get() == "Select Class" or
                    entry11.get() == ""):
                    CTkMessagebox(title="Error", message="Please fill all the required fields.", icon="cancel")
                elif not entry1.get().isdigit() or int(entry1.get()) < 0:
                    CTkMessagebox(title="Invalid Input", message="Please enter valid No Of Students", icon="warning")
                elif not validate_co_string(entry11.get()):
                    CTkMessagebox(title="Invalid Input", message="Please enter the CO in valid format", icon="warning")
                elif entry10.get() == "2":
                    if entry13.get() == "Select Type" or entry14.get() == "Select Type":
                        CTkMessagebox(title="Error", message="Select type of CA", icon="cancel")
                    elif entry13.get() == "Presentation" and presentationCA1Text.get() == "":
                        CTkMessagebox(title="Error", message="Please fill the maximum students in a group.", icon="cancel")
                    elif entry13.get() == "NPTEL Course" and nptelCA1Text.get() == "":
                        CTkMessagebox(title="Error", message="Please fill the CO number for NPTEL course.", icon="cancel")
                    elif entry13.get() == "Quiz" and noCA1Entry.get() == "Select No":
                        CTkMessagebox(title="Error", message="Please select number of questions in CA1", icon="cancel")
                    elif entry14.get() == "Presentation" and presentationCA2Text.get() == "":
                        CTkMessagebox(title="Error", message="Please fill the maximum students in a group.", icon="cancel")
                    elif entry14.get() == "NPTEL Course" and nptelCA2Text.get() == "":
                        CTkMessagebox(title="Error", message="Please fill the CO number for NPTEL course.", icon="cancel")
                    elif entry14.get() == "Quiz" and noCA2Entry.get() == "Select No":
                        CTkMessagebox(title="Error", message="Please select number of questions in CA2", icon="cancel")
                    elif entry13.get() == "NPTEL Course" and not validate_co_string(nptelCA1Text.get()):
                        CTkMessagebox(title="Invalid Input", message="Please enter the CO in valid format (CA1, NPTEL Course)", icon="warning")
                    elif entry14.get() == "NPTEL Course" and not validate_co_string(nptelCA2Text.get()):
                        CTkMessagebox(title="Invalid Input", message="Please enter the CO in valid format (CA2, NPTEL Course)", icon="warning")
                    else:
                        tabview.set(" CO Mapping ")
               
                elif entry10.get() == "3":
                    if entry13.get() == "Select Type" or entry14 == "Select Type":
                        CTkMessagebox(title="Error", message="Select type of CA", icon="cancel")
                    elif entry13.get() == "Presentation" and presentationCA1Text.get() == "":
                        CTkMessagebox(title="Error", message="Please fill the maximum students in a group.", icon="cancel")
                    elif entry13.get() == "NPTEL Course" and nptelCA1Text.get() == "":
                        CTkMessagebox(title="Error", message="Please fill the CO number for NPTEL course", icon="cancel")
                    elif entry13.get() == "Quiz" and noCA1Entry.get() == "Select No":
                        CTkMessagebox(title="Error", message="Please select number of questions in CA1", icon="cancel")
                    elif entry14.get() == "Presentation" and presentationCA2Text.get() == "":
                        CTkMessagebox(title="Error", message="Please fill the maximum students in a group.", icon="cancel")
                    elif entry14.get() == "NPTEL Course" and nptelCA2Text.get() == "":
                        CTkMessagebox(title="Error", message="Please fill the CO number for NPTEL course", icon="cancel")
                    elif entry14.get() == "Quiz" and noCA2Entry.get() == "Select No":
                        CTkMessagebox(title="Error", message="Please select number of questions in CA2", icon="cancel")
                    elif entry15.get() == "Presentation" and presentationCA3Text.get() == "":
                        CTkMessagebox(title="Error", message="Please fill the maximum students in a group.", icon="cancel")
                    elif entry15.get() == "NPTEL Course" and nptelCA3Text.get() == "":
                        CTkMessagebox(title="Error", message="Please fill the CO number for NPTEL course", icon="cancel")
                    elif entry15.get() == "Quiz" and noCA3Entry.get() == "Select No":
                        CTkMessagebox(title="Error", message="Please select number of questions in CA3", icon="cancel")
                    elif entry13.get() == "NPTEL Course" and not validate_co_string(nptelCA1Text.get()):
                        CTkMessagebox(title="Error", message="Please enter the CO in valid format (CA1, NPTEL Course)", icon="cancel")
                    elif entry14.get() == "NPTEL Course" and not validate_co_string(nptelCA2Text.get()):
                        CTkMessagebox(title="Error", message="Please enter the CO in valid format (CA2, NPTEL Course)", icon="cancel")
                    elif entry15.get() == "NPTEL Course" and not validate_co_string(nptelCA3Text.get()):
                        CTkMessagebox(title="Error", message="Please enter the CO in valid format (CA3, NPTEL Course)", icon="cancel")
                    else:
                        tabview.set(" CO Mapping ")
                # self.pathName = f"{yearDropDown.get()}_{entry2.get()}_{entry3.get()}_{entry7.get()}_{entry5.get()}_{entry4.get()}.xlsx"
                # self.pathName = self.pathName.replace(" ","_")
    
            def switch1():
                if noOfCOOption.get() == "5":
                    if CO1T.get() != "" and CO2T.get() != "" and CO3T.get() != "" and CO4T.get() != "" and CO5T.get() != "":
                        tabview.set(" Basic Information ")
                        valid_CO = [1,2,3,4,5]
                    else:
                        CTkMessagebox(title="Error", message="Please enter all the 5 CO's", icon="cancel")
                elif noOfCOOption.get() == "6":
                    if CO1T.get() != "" and CO2T.get() != "" and CO3T.get() != "" and CO4T.get() != "" and CO5T.get() != "" and CO6T.get() != "":
                        tabview.set(" Basic Information ")
                        valid_CO = [1,2,3,4,5,6]
                    else:
                        CTkMessagebox(title="Error", message="Please enter all the 6 CO's", icon="cancel")
                else:
                    CTkMessagebox(title="Error", message="Please Select No Of CO's", icon="cancel")
    
            def switch2():
                if a1T.get()=="" or a2T.get()=="" or a3T.get()=="" or a4T.get()=="" or a5T.get()=="" or a6T.get()=="" or a2aT.get()=="" or a2bT.get()=="" or a3aT == "" or a3bT.get()=="":
                    CTkMessagebox(title="Error", message="Please enter CO\'s for all questions", icon="cancel")
                elif not (validate_co_string(a1T.get()) and validate_co_string(a2T.get()) and validate_co_string(a3T.get()) and validate_co_string(a4T.get()) and validate_co_string(a5T.get()) and validate_co_string(a6T.get()) and validate_co_string(a2aT.get()) and validate_co_string(a2bT.get()) and validate_co_string(a3aT.get()) and validate_co_string(a3bT.get())):
                    CTkMessagebox(title="Error", message="Please enter valid format of CO string", icon="cancel")
                elif entry10.get() == "2":
                    if entry13.get() == "Quiz":
                        check_text_CA1 = [q1TCA1.get(), q2TCA1.get(), q3TCA1.get(), q4TCA1.get(), q5TCA1.get(), q6TCA1.get(), q7TCA1.get(), q8TCA1.get(), q9TCA1.get(), q10TCA1.get()]
                        no_of_text_fields = int(noCA1Entry.get())
                        for i in range (0, no_of_text_fields):
                            if(check_text_CA1[i] == ""):
                                CTkMessagebox(title="Error", message="Please enter the CO\'s for all questions", icon="cancel")
                                return
                        for i in range (0, no_of_text_fields):
                            if not validate_co_string(check_text_CA1[i]):
                                CTkMessagebox(title="Invalid Input", message="Please enter the valid format of CO", icon="warning")
                                return
                    if entry14.get() == "Quiz":
                        check_text_CA2 = [q1TCA2.get(), q2TCA2.get(), q3TCA2.get(), q4TCA2.get(), q5TCA2.get(), q6TCA2.get(), q7TCA2.get(), q8TCA2.get(), q9TCA2.get(), q10TCA2.get()]
                        no_of_text_fields = int(noCA2Entry.get())
                        for i in range(0,no_of_text_fields):
                            if(check_text_CA2[i] == ""):
                                CTkMessagebox(title="Error", message="Please enter the CO\'s for all questions", icon="cancel")
                                return
                        for i in range(0,no_of_text_fields):
                            if not validate_co_string(check_text_CA2[i]):
                                CTkMessagebox(title="Invalid Input", message="Please enter the valid format of CO", icon="warning")
                                return
                    tabview.set(" Target level of tests ")
    
                elif entry10.get()=="3":
                    if entry13.get() == "Quiz":
                        check_text_CA1 = [q1TCA1.get(), q2TCA1.get(), q3TCA1.get(), q4TCA1.get(), q5TCA1.get(), q6TCA1.get(), q7TCA1.get(), q8TCA1.get(), q9TCA1.get(), q10TCA1.get()]
                        no_of_text_fields = int(noCA1Entry.get())
                        for i in range (0, no_of_text_fields):
                            if(check_text_CA1[i] == ""):
                                CTkMessagebox(title="Error", message="Please enter the CO\'s for all questions", icon="cancel")
                                return
                        for i in range (0, no_of_text_fields):
                            if not validate_co_string(check_text_CA1[i]):
                                CTkMessagebox(title="Invalid Input", message="Please enter the valid format of CO", icon="warning")
                                return
                    if entry14.get() == "Quiz":
                        check_text_CA2 = [q1TCA2.get(), q2TCA2.get(), q3TCA2.get(), q4TCA2.get(), q5TCA2.get(), q6TCA2.get(), q7TCA2.get(), q8TCA2.get(), q9TCA2.get(), q10TCA2.get()]
                        no_of_text_fields = int(noCA2Entry.get())
                        for i in range(0,no_of_text_fields):
                            if(check_text_CA2[i] == ""):
                                CTkMessagebox(title="Error", message="Please enter the CO\'s for all questions", icon="cancel")
                                return
                        for i in range(0,no_of_text_fields):
                            if not validate_co_string(check_text_CA2[i]):
                                CTkMessagebox(title="Invalid Input", message="Please enter the valid format of CO", icon="warning")
                                return
                    if entry15.get() == "Quiz":
                        check_text_CA3 = [q1TCA3.get(), q2TCA3.get(), q3TCA3.get(), q4TCA3.get(), q5TCA3.get(), q6TCA3.get(), q7TCA3.get(), q8TCA3.get(), q9TCA3.get(), q10TCA3.get()]
                        no_of_text_fields = int(noCA3Entry.get())
                        for i in range(0, no_of_text_fields):
                            if(check_text_CA3[i] == ""):
                                CTkMessagebox(title="Error", message="Please enter the CO\'s for all questions", icon="cancel")
                                return
                        for i in range(0,no_of_text_fields):
                            if not (validate_co_string(check_text_CA3[i])):
                                CTkMessagebox(title="Invalid Input", message="Please enter the valid format of CO", icon="warning")
                                return
                    tabview.set(" Target level of tests ")
    
            def create_button(tab, name, font_name, font_size, w, h, com, x, y):
                button = ctk.CTkButton(master=tabview.tab(tab), text=name, width=w, height=h, font=(font_name, font_size), command=com)
                button.place(x=x, y=y)
                return button
            
            def create_label(tab, name, font_type, font_size, x, y):
                label = ctk.CTkLabel(master=tabview.tab(tab), text=name, font=(font_type, font_size))
                label.place(x=x, y=y)
                return label
    
            def create_entry_box(tab, text, font_name, font_size, w, x, y):
                entry_box = ctk.CTkEntry(master=tabview.tab(tab), placeholder_text=text, font=(font_name,font_size), width=w)
                entry_box.place(x=x,y=y)
                return entry_box
            
            def create_dropdown(tab, val, font_name, font_size, w, com, x, y):
                dropdown = ctk.CTkOptionMenu(master=tabview.tab(tab), values=val, font=(font_name, font_size), width=w, command=com)
                dropdown.place(x=x, y=y)
                return dropdown
            
            def download():
                # Get the values from the Entry widgets
                # values = [entry1.get(), entry2.get(), entry3.get(), entry4.get(), entry5.get(),
                #         a1T.get(), a2T.get(), a3T.get(), a4T.get(), a5T.get(), a6T.get(),
                #         a2aT.get(), a2bT.get(), a3aT.get(), a3bT.get(),entry7.get(),entry8.get(),entry10.get(),entry11.get(),entry12.get()]
                
                # coValues =[entry13.get(),entry14.get(),entry15.get(),q1T.get(),q6T.get(),q2T.get(),q2aT.get(),q3T.get(),q2bT.get(),q4T.get(),q3aT.get(),q5T.get(),q3bT.get()]
                # coCAs=[entry13.get(),entry14.get(),entry15.get()]
                # coQuizs=[q1T.get(),q6T.get(),q2T.get(),q2aT.get(),q3T.get(),q2bT.get(),q4T.get(),q3aT.get(),q5T.get(),q3bT.get()]
    
                if noOfCOOption.get()=='6':
                    coTextArray = [CO1T.get(),CO2T.get(),CO3T.get(),CO4T.get(),CO5T.get(),CO6T.get()]
                elif noOfCOOption.get()=='5':
                    coTextArray = [CO1T.get(),CO2T.get(),CO3T.get(),CO4T.get(),CO5T.get(),'-']
                else:
                    CTkMessagebox(title="Error", message="Please Select No Of CO's", icon="cancel")
    
                print(coTextArray)
                
                values=[entry1.get(),entry8.get(),yearDropDown.get(),entry2.get(),entry3.get(),entry4.get(),entry5.get(),entry7.get(),entry11.get(),
                
    
                        entry10.get(),entry13.get(),entry14.get(),entry15.get(),
    
                        noCA1Entry.get(),noCA2Entry.get(),noCA3Entry.get(),
                        nptelCA1Text.get(),nptelCA2Text.get(),nptelCA3Text.get(),
    
                        q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get(),q9TCA1.get(),q10TCA1.get(),
                        q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get(),q8TCA2.get(),q9TCA2.get(),q10TCA2.get(),
                        q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get(),q5TCA3.get(),q6TCA3.get(),q7TCA3.get(),q8TCA3.get(),q9TCA3.get(),q10TCA3.get(),
    
                        a1T.get(), a2T.get(), a3T.get(), a4T.get(), a5T.get(), a6T.get(),a2aT.get(),a2bT.get(), a3aT.get(), a3bT.get()]
                
                basic_values=[entry1.get(),entry8.get(),yearDropDown.get(),entry2.get(),entry3.get(),entry4.get(),entry5.get(),entry7.get(),entry11.get(),None,entry10.get(),entry13.get(),entry14.get(),entry15.get(),noOfCOOption.get()]
                midSem_Co_values=[a1T.get(), a2T.get(), a3T.get(), a4T.get(), a5T.get(), a6T.get(),a2aT.get(),a2bT.get(), a3aT.get(), a3bT.get()]
                receiversEmail = emailText.get()
                
                if entry10.get()=="2":
                    al_values=[ALCA1Text.get(), ALCA2Text.get(), '-', ALMidTermText.get(), ALEndSemText.get()]
                    if entry13.get()=="Select Type":
                        CTkMessagebox(title="Error", message="Please Select Type of CA 1.", icon="cancel")
                    elif entry13.get()=="Quiz":
                        if noCA1Entry.get()=="Select No" :
                            CTkMessagebox(title="Error", message="Please Select No Of Question in CA 1.", icon="cancel")
                        elif noCA1Entry.get()=="1" :
                            CA1_Co_arr=[q1TCA1.get()]
                        elif noCA1Entry.get()=="2" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get()]
                        elif noCA1Entry.get()=="3" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get()]
                        elif noCA1Entry.get()=="4" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get()]
                        elif noCA1Entry.get()=="5" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get()]
                        elif noCA1Entry.get()=="6" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get()]
                        elif noCA1Entry.get()=="7" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get()]
                        elif noCA1Entry.get()=="8" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get()]
                        elif noCA1Entry.get()=="9" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get(),q9TCA1.get()]
                        else:
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get(),q9TCA1.get(),q10TCA1.get()]       
                    
                    elif entry13.get()=="Presentation":
                        CA1_Co_arr = [presentationCA1Text.get()]
                    else:
                        CA1_Co_arr=[nptelCA1Text.get()]
                     
                            
                    if entry14.get()=="Select Type":
                        CTkMessagebox(title="Error", message="Please Select Type of CA 2.", icon="cancel")
                    elif entry14.get()=="Quiz":
                        if noCA2Entry.get()=="Select No" :
                            CTkMessagebox(title="Error", message="Please Select No Of Question in CA 2.", icon="cancel")
                        elif noCA2Entry.get()=="1" :
                            CA2_Co_arr=[q1TCA2.get()]
                        elif noCA2Entry.get()=="2" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get()]
                        elif noCA2Entry.get()=="3" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get()]
                        elif noCA2Entry.get()=="4" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get()]
                        elif noCA2Entry.get()=="5" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get()]
                        elif noCA2Entry.get()=="6" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get()]
                        elif noCA2Entry.get()=="7" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get()]
                        elif noCA2Entry.get()=="8" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get(),q8TCA2.get()]
                        elif noCA2Entry.get()=="9" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get(),q8TCA2.get(),q9TCA2.get()]
                        else :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get(),q8TCA2.get(),q9TCA2.get(),q10TCA2.get()]
                    
                    elif entry14.get()=="Presentation":
                        CA2_Co_arr = [presentationCA2Text.get()]
                    else:
                        CA2_Co_arr=[nptelCA2Text.get()]
                        
                else:
                    al_values=[ALCA1Text.get(), ALCA2Text.get(), ALCA3Text.get(), ALMidTermText.get(), ALEndSemText.get()]
                    if entry13.get()=="Select Type":
                        CTkMessagebox(title="Error", message="Please Select Type of CA 1.", icon="cancel")
                    elif entry13.get()=="Quiz":
                        if noCA1Entry.get()=="Select No" :
                            CTkMessagebox(title="Error", message="Please Select No Of Question in CA 1.", icon="cancel")
                        elif noCA1Entry.get()=="1" :
                            CA1_Co_arr=[q1TCA1.get()]
                        elif noCA1Entry.get()=="2" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get()]
                        elif noCA1Entry.get()=="3" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get()]
                        elif noCA1Entry.get()=="4" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get()]
                        elif noCA1Entry.get()=="5" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get()]
                        elif noCA1Entry.get()=="6" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get()]
                        elif noCA1Entry.get()=="7" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get()]
                        elif noCA1Entry.get()=="8" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get()]
                        elif noCA1Entry.get()=="9" :
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get(),q9TCA1.get()]
                        else:
                            CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get(),q9TCA1.get(),q10TCA1.get()]
                    
                    elif entry13.get()=="Presentation":
                        CA1_Co_arr = [presentationCA1Text.get()]
                    else:
                        CA1_Co_arr=[nptelCA1Text.get()]
                    
                        
                    if entry14.get()=="Select Type":
                        CTkMessagebox(title="Error", message="Please Select Type of CA 2.", icon="cancel")
                    elif entry14.get()=="Quiz":
                        if noCA2Entry.get()=="Select No" :
                            CTkMessagebox(title="Error", message="Please Select No Of Question in CA 2.", icon="cancel")
                        elif noCA2Entry.get()=="1" :
                            CA2_Co_arr=[q1TCA2.get()]
                        elif noCA2Entry.get()=="2" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get()]
                        elif noCA2Entry.get()=="3" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get()]
                        elif noCA2Entry.get()=="4" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get()]
                        elif noCA2Entry.get()=="5" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get()]
                        elif noCA2Entry.get()=="6" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get()]
                        elif noCA2Entry.get()=="7" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get()]
                        elif noCA2Entry.get()=="8" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get(),q8TCA2.get()]
                        elif noCA2Entry.get()=="9" :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get(),q8TCA2.get(),q9TCA2.get()]
                        else :
                            CA2_Co_arr=[q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get(),q8TCA2.get(),q9TCA2.get(),q10TCA2.get()]
                    
                    elif entry14.get()=="Presentation":
                        CA2_Co_arr = [presentationCA2Text.get()]
                    else:
                        CA2_Co_arr=[nptelCA2Text.get()]
                        
                    if entry15.get()=="Select Type":
                        CTkMessagebox(title="Error", message="Please Select Type of CA 3.", icon="cancel")
                    elif entry15.get()=="Quiz":
                        if noCA3Entry.get()=="Select No" :
                            CTkMessagebox(title="Error", message="Please Select No Of Question in CA 3.", icon="cancel")
                        elif noCA3Entry.get()=="1" :
                            CA3_Co_arr=[q1TCA3.get()]
                        elif noCA3Entry.get()=="2" :
                            CA3_Co_arr=[q1TCA3.get(),q2TCA3.get()]
                        elif noCA3Entry.get()=="3" :
                            CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get()]
                        elif noCA3Entry.get()=="4" :
                            CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get()]
                        elif noCA3Entry.get()=="5" :
                            CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get(),q5TCA3.get()]
                        elif noCA3Entry.get()=="6" :
                            CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get(),q5TCA3.get(),q6TCA3.get()]
                        elif noCA3Entry.get()=="7" :
                            CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get(),q5TCA3.get(),q6TCA3.get(),q7TCA3.get()]
                        elif noCA3Entry.get()=="8" :
                            CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get(),q5TCA3.get(),q6TCA3.get(),q7TCA3.get(),q8TCA3.get()]
                        elif noCA3Entry.get()=="9" :
                            CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get(),q5TCA3.get(),q6TCA3.get(),q7TCA3.get(),q8TCA3.get(),q9TCA3.get()]
                        else :
                            CA3_Co_arr=[q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get(),q5TCA3.get(),q6TCA3.get(),q7TCA3.get(),q8TCA3.get(),q9TCA3.get(),q10TCA3.get()]
                    
                    elif entry15.get()=="Presentation":
                        CA3_Co_arr = [presentationCA3Text.get()]
                    else:
                        CA3_Co_arr=[nptelCA3Text.get()]
                        
                if basic_values[1]=="Select Department":
                    CTkMessagebox(title="Error", message="Please Select Department.", icon="cancel")
                elif basic_values[2]=="Select Year":
                    CTkMessagebox(title="Error", message="Please Select Year.", icon="cancel")
                elif basic_values[3]=="Select Sem":
                    CTkMessagebox(title="Error", message="Please Select Semester.", icon="cancel")
                elif basic_values[4]=="Select Subject":
                    CTkMessagebox(title="Error", message="Please Select Subject.", icon="cancel")
                elif any(value == "" for value in basic_values):
                    CTkMessagebox(title="Error", message="Please fill in all required fields.", icon="cancel")
                elif any(midCo == "" for midCo in midSem_Co_values):
                    CTkMessagebox(title="Error", message="Please fill in all required fields.", icon="cancel")
                elif any(ca1 == "" for ca1 in CA1_Co_arr):
                    CTkMessagebox(title="Error", message="Please fill in all required fields.", icon="cancel")
                elif any(ca2 == "" for ca2 in CA2_Co_arr):
                    CTkMessagebox(title="Error", message="Please fill in all required fields.", icon="cancel")
                elif entry10.get()=="3":
                    if any(ca3 == "" for ca3 in CA3_Co_arr):
                        CTkMessagebox(title="Error", message="Please fill in all required fields.", icon="cancel")
                    else :
                        import template_generator
                        template_generator.template_gen(coTextArray,basic_values,midSem_Co_values,CA1_Co_arr,CA2_Co_arr,CA3_Co_arr,al_values, receiversEmail)
                        # template_generator.template_gen(basic_values,midSem_Co_values,CA1_Co_arr,CA2_Co_arr,CA3_Co_arr)
                        # CTkMessagebox(message="Excel template downloaded successfully.",icon="check", option_1="OK")
    
                        
                else:
                    # import template_generator
                    # print(entry10.get())
                    if entry10.get()=="2": 
                        import template_generator
                        # print("Hi v1",basic_values[10]) 
                        template_generator.template_gen(coTextArray,basic_values,midSem_Co_values,CA1_Co_arr,CA2_Co_arr,[],al_values,receiversEmail)
                        # CTkMessagebox(message="Excel template downloaded successfully.",icon="check", option_1="OK")
                    elif entry10.get()=="3":
                        import template_generator
                        # print("Hi v2",basic_values[10]) 
                        # print("Hi v2",CA3_Co_arr) 
                        template_generator.template_gen(coTextArray,basic_values,midSem_Co_values,CA1_Co_arr,CA2_Co_arr,CA3_Co_arr,al_values,receiversEmail)
                        # CTkMessagebox(message="Excel template downloaded successfully .",icon="check", option_1="OK")
                
            def validate_co_string(coString):
                validate_co_array = []
                # print(coString)
                coString = coString.replace(" ", "")
                # print(coString)
                if noOfCOOption.get() == "Select No of CO\'s":
                    CTkMessagebox(title = "Error", message="Select No of CO\'s", icon="cancel")
                elif noOfCOOption.get() == "5":
                    validate_co_array = [1,2,3,4,5]
                elif noOfCOOption.get() == "6":
                    validate_co_array = [1,2,3,4,5,6]
                
                pattern = r'^(\d,)*\d$'
                if not re.match(pattern, coString):
                    return False
    
                # Extract digits from the input string
                digits = list(map(int, coString.split(',')))
                # print(digits)
    
                # Check each digit is within the valid_digits array
                if not all(digit in validate_co_array for digit in digits):
                    return False
    
                # Ensure there are no consecutive identical digits
                if len(digits) != len(set(digits)):
                    return False
    
                return True
                                            
                    
                
            def ca1(option):
                if  option == "Select Type":
                    for disca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1,noCA1Entry,nptelCA1Text,presentationCA1Text]:
                        disca.configure(state="disabled", fg_color="gray")
              
                elif option == "NPTEL Course":
                    for disca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1,noCA1Entry,presentationCA1Text]:
                        disca.configure(state="disabled", fg_color="gray")
                    nptelCA1Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"]) 
                
                elif option == "Presentation":
                    for preca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1,noCA1Entry,nptelCA1Text]:
                        preca.configure(state="disabled", fg_color="gray")
                    presentationCA1Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                else:
                    for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    noCA1Entry.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
                    nptelCA1Text.configure(state="disabled", fg_color="gray") 
                    presentationCA1Text.configure(state="disabled", fg_color="gray") 
                    
            def ca2(option):
                if  option == "Select Type":
                    for disca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2,noCA2Entry,nptelCA2Text,presentationCA2Text]:
                        disca.configure(state="disabled", fg_color="gray")
                elif option == "NPTEL Course":
                    for disca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2,noCA2Entry,presentationCA2Text]:
                        disca.configure(state="disabled", fg_color="gray")
                    nptelCA2Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                elif option == "Presentation":
                    for preca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2,noCA2Entry,nptelCA2Text]:
                        preca.configure(state="disabled", fg_color="gray")
                    presentationCA2Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                else:
                    for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    noCA2Entry.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
                    nptelCA2Text.configure(state="disabled", fg_color="gray") 
                    presentationCA2Text.configure(state="disabled", fg_color="gray")
                    
    
            def ca3(option):
                if  option == "Select Type":
                    for disca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3,noCA3Entry,nptelCA3Text,presentationCA3Text]:
                        disca.configure(state="disabled", fg_color="gray")
                    
                elif option == "NPTEL Course":
                    for disca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3,noCA3Entry,presentationCA3Text]:
                        disca.configure(state="disabled", fg_color="gray")
                    nptelCA3Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])    
                
                elif option == "Presentation":
                    for preca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3,noCA3Entry,nptelCA3Text]:
                        preca.configure(state="disabled", fg_color="gray")
                    presentationCA3Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                else:
                    for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    noCA3Entry.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
                    nptelCA3Text.configure(state="disabled", fg_color="gray") 
                    presentationCA3Text.configure(state="disabled", fg_color="gray")
                   
            def semesterAndClass(option):
                if option == "Select Year":
                    entry2.configure(values=["Select Sem"])
                    entry7.configure(values=["Select Class"])
                    # entry2_lab.configure(values=["Select Sem"])
                    # entry7_lab.configure(values=["Select Class"])
                elif option == "F.E":
                    entry2.configure(values=["Select Sem","I","II"])
                    entry7.configure(values=["Select Class", "D5A", "D5B", "D5C"])
                    # entry2_lab.configure(values=["Select Sem","I","II"])
                    # entry7_lab.configure(values=["Select Class", "D5A", "D5B", "D5C"])
                elif option == "S.E":
                    entry2.configure(values=["Select Sem","III","IV"])
                    entry7.configure(values=["Select Class", "D10A", "D10B", "D10C"])
                    # entry2_lab.configure(values=["Select Sem","III","IV"])
                    # entry7_lab.configure(values=["Select Class", "D10A", "D10B", "D10C"])
                elif option == "T.E":
                    entry2.configure(values=["Select Sem","V","VI"])
                    entry7.configure(values=["Select Class", "D15A", "D15B", "D15C"])
                    # entry2_lab.configure(values=["Select Se/m","V","VI"])
                    # entry7_lab.configure(values=["Select Class", "D15A", "D15B", "D15C"])
                elif option == "B.E":
                    entry2.configure(values=["Select Sem","VII","VIII"])
                    entry7.configure(values=["Select Class", "D20A", "D20B", "D20C"])
                    # entry2_lab.configure(values=["Select Sem","VII","VIII"])
                    # entry7_lab.configure(values=["Select Class", "D20A", "D20B", "D20C"])
    
    
            def noQuestion1(option):
                if option== "1" :
                    for ca in [q1TCA1]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option== "2" :
                    for ca in [q1TCA1,q2TCA1]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option== "3" :
                    for ca in [q1TCA1,q2TCA1,q3TCA1]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option== "4" :
                    for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                        disca.configure(state="disabled", fg_color="gray")
                  
                elif option== "5" :
                    for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                        disca.configure(state="disabled", fg_color="gray")
                  
                elif option== "6" :
                    for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                        disca.configure(state="disabled", fg_color="gray")
    
                elif option== "7" :
                    for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q8TCA1,q9TCA1,q10TCA1]:
                        disca.configure(state="disabled", fg_color="gray")
                    
                elif option== "8" :
                    for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q9TCA1,q10TCA1]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option== "9" :
                    for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q10TCA1]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option=="10" :
                    for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
                 
                else :
                    for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]:
                        ca.configure(state="disabled", fg_color="gray")  
                          
            def noQuestion2(option):
                if option== "1" :
                    for ca in [q1TCA2]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option== "2" :
                    for ca in [q1TCA2,q2TCA2]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option== "3" :
                    for ca in [q1TCA2,q2TCA2,q3TCA2]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option== "4" :
                    for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option== "5" :
                    for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option== "6" :
                    for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option== "7" :
                    for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q8TCA2,q9TCA2,q10TCA2]:
                        disca.configure(state="disabled", fg_color="gray")
                    
                elif option== "8" :
                    for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q9TCA2,q10TCA2]:
                        disca.configure(state="disabled", fg_color="gray")
                    
                elif option== "9" :
                    for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q10TCA2]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option=="10" :
                    for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                else :
                    for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]:
                        ca.configure(state="disabled", fg_color="gray")
                    
            def noQuestion3(option):
                if option== "1" :
                    for ca in [q1TCA3]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option== "2" :
                    for ca in [q1TCA3,q2TCA3]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option== "3" :
                    for ca in [q1TCA3,q2TCA3,q3TCA3]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option== "4" :
                    for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option== "5" :
                    for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                        disca.configure(state="disabled", fg_color="gray")
                    
                elif option== "6" :
                    for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                        disca.configure(state="disabled", fg_color="gray")
                    
                elif option== "7" :
                    for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q8TCA3,q9TCA3,q10TCA3]:
                        disca.configure(state="disabled", fg_color="gray")
                    
                elif option== "8" :
                    for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q9TCA3,q10TCA3]:
                        disca.configure(state="disabled", fg_color="gray")
                   
                elif option== "9" :
                    for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                        
                    for disca in [q10TCA3]:
                        disca.configure(state="disabled", fg_color="gray")
                    
                elif option=="10" :
                    for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                        ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                   
                else:
                    for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]:
                        ca.configure(state="disabled", fg_color="gray")
                    
                    
            def subject(option):
                if option == "Select Sem":
                    entry3.configure(values=["Select Subject"])
                    # entry3_lab.configure(values=["Select Subject"])
                elif option == "I":
                    entry3.configure(values=["Select Subject","Universal Human Values - 1","Fundamentals of Vedic Mathematics (Indian Knowledge System)", "Basic Electrical Engineering", "Engineering Drawing", "Engineering Mechanics", "Engineering Physics", "Matrices and Differential Calculus", "Python Programming"])
                    # entry3_lab.configure(values=["Select Subject","Universal Human Values - 1","Fundamentals of Vedic Mathematics (Indian Knowledge System)", "Basic Electrical Engineering", "Engineering Drawing", "Engineering Mechanics", "Engineering Physics", "Matrices and Differential Calculus", "Python Programming"])
                elif option == "II":
                    entry3.configure(values=["Select Subject","Universal Human Values - 2","Basic Workshop Practice", "Computer Programming", "Integral Calculus and Complex Numbers", "Biology for Engineers", "Engineering Chemistry", "Professional Communication and Ethics - 1"])
                    # entry3_lab.configure(values=["Select Subject","Universal Human Values - 2","Basic Workshop Practice", "Computer Programming", "Integral Calculus and Complex Numbers", "Biology for Engineers", "Engineering Chemistry", "Professional Communication and Ethics - 1"])
                elif option == "III":
                    entry3.configure(values=["Select Subject","Engineering Mathematics III", "Data Structures and Analysis", "Database Management System", "Principle of Communications", "Paradigm and computer programming fundamentals"])
                    # entry3_lab.configure(values=["Select Subject","Engineering Mathematics III", "Data Structures and Analysis", "Database Management System", "Principle of Communications", "Paradigm and computer programming fundamentals"])
                elif option == "IV":
                    entry3.configure(values=["Select Subject","Engineering Mathematics IV", "Computer Network and Network Design", "Operating System", "Automata Theory", "Computer Organization and Architecture"])
                    # entry3_lab.configure(values=["Select Subject","Engineering Mathematics IV", "Computer Network and Network Design", "Operating System", "Automata Theory", "Computer Organization and Architecture"])
                elif option == "V":
                    entry3.configure(values=["Select Subject","Internet Programming", "Computer Network Security", "Entrepreneurship and E- business", "Software Engineering", "Advance Data Management Technologies", "Advanced Data structure and Analysis"])
                    # entry3_lab.configure(values=["Select Subject","Internet Programming", "Computer Network Security", "Entrepreneurship and E- business", "Software Engineering", "Advance Data Management Technologies", "Advanced Data structure and Analysis"])
                elif option == "VI":
                    entry3.configure(values=["Select Subject","Data Mining & Business Intelligence", "Web X.0", "Wireless Technology", "AI and DS 1", "Optional Course 2"])
                    # entry3_lab.configure(values=["Select Subject","Data Mining & Business Intelligence", "Web X.0", "Wireless Technology", "AI and DS 1", "Optional Course 2"])
                elif option == "VII":
                    entry3.configure(values=["Select Subject","AI and DS II", "Internet of Everything", "Department Optional Course 3", "Department Optional Course 4", "Institute Optional Course 1"])
                    # entry3_lab.configure(values=["Select Subject","AI and DS II", "Internet of Everything", "Department Optional Course 3", "Department Optional Course 4", "Institute Optional Course 1"])
                elif option == "VIII":
                    entry3.configure(values=["Select Subject","Blockchain and DLT", "Department Optional Course 5", "Department Optional Course 6", "Institute Optional Course 2"])
                    # entry3_lab.configure(values=["Select Subject","Blockchain and DLT", "Department Optional Course 5", "Department Optional Course 6", "Institute Optional Course 2"])
    
            def disable(option):
                if option == "3":
                    for entry in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3,ALCA3Text]:
                        entry.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    entry15.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
                   
                else:
                    for entry in [entry15,q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3,nptelCA3Text,noCA3Entry,ALCA3Text]:
                        entry.configure(state="disabled", fg_color='gray')
    
            def validate_academic_year(event):
                new_value = event.widget.get()
                if new_value:
                    if len(new_value) != 9 or new_value[4] != '-' or not (new_value[:4].isdigit() and new_value[5:].isdigit()):
                        CTkMessagebox(title="Invalid Input", message="Academic Year format is incorrect. Please enter in the format YYYY-YYYY.", icon="warning")                
                        return False
                    else: 
                        return True
    
            def noOfCO(option):
                if option=='6':
                    CO6T.configure(state="normal",fg_color=["#F9F9FA", "#343638"])
                else :
                    CO6T.configure(state="disabled",fg_color="gray")
                   
                    
    
            def resource_path(relative_path):
                """Get the absolute path to a resource, handling PyInstaller paths."""
                if hasattr(sys, '_MEIPASS'):  # PyInstaller extracts files to _MEIPASS
                    return os.path.join(sys._MEIPASS, relative_path)
                return os.path.join(os.path.abspath("."), relative_path)

            # Use resource_path to access the image
            image_path = resource_path("CO Calculator.png")
            # Load the image and create a CTkImage
            background_image = Image.open(image_path)
            bg_image = ctk.CTkImage(background_image, size=(screen_width - 100, screen_height-130))

            # Create a frame for the "Instructions" tab content
            instructions_tab = tabview.tab(" Instructions ")
            instructions_tab.columnconfigure(0, weight=1)
            instructions_tab.rowconfigure(0, weight=1)

            # Add a label to hold the background image
            bg_label = ctk.CTkLabel(master=instructions_tab, image=bg_image, text="")
            bg_label.place(relx=0.5, rely=0.5, anchor="center")

            # If you want to overlay widgets on top of the image:
            # Example of overlaying text on the background
            overlay_label = ctk.CTkLabel(master=instructions_tab, text="")
            overlay_label.place(relx=0.5, rely=0.1, anchor="center")

            button = create_button(" Basic Information ", "Next", "Arial", 20, 150, 40, switch, 725, 690)

            label0 = create_label(" Basic Information ", "Basic Details", "Arial", 20, 725, 5)

            label1 = create_label(" Basic Information ", "No. of Students :", "Arial", 15, 200, 55)

            entry1 = create_entry_box(" Basic Information ", "Enter no of students", "Arial", 15, 300, 400, 55)

            newLabel = create_label(" Basic Information ", "Year :", "Arial", 15, 200, 155)

            yearDropDown = create_dropdown(" Basic Information ", ["Select Year", "F.E", "S.E", "T.E", "B.E"], "Arial", 15, 300, semesterAndClass, 400, 155)

            label8 = create_label(" Basic Information ", "Department :", "Arial", 15, 200, 105)

            entry8 = create_dropdown(" Basic Information ", ["Select Department", "Humanities and Applied Science(FE)", "Information Technology", "Computer", "AI and Data Science", "Electronics and Telecommunication", "Electronics", "Instrumentation"], "Arial", 15, 300, None, 400, 105)

            label2 = create_label(" Basic Information ", "Semester :", "Arial", 15, 200, 205)

            entry2 = create_dropdown(" Basic Information ", ["Select Sem"], "Arial", 15, 300, subject, 400, 205)

            label3 = create_label(" Basic Information ", "Subject :", "Arial", 15, 200, 255)

            entry3 = create_dropdown(" Basic Information ", ["Select Subject"], "Arial", 15, 300, None, 400, 255)

            label4 = create_label(" Basic Information ", "Academic Year: ", "Arial", 15, 200, 305)

            entry4 = create_entry_box(" Basic Information ", "YYYY-YYYY", "Arial", 15, 300, 400, 305)
            entry4.bind("<FocusOut>", validate_academic_year)

            label5 = create_label(" Basic Information ", "Subject Teacher :", "Arial", 15, 200, 355)

            entry5 = create_entry_box(" Basic Information ", "Subject Teacher", "Arial", 15, 300, 400, 355)

            label7 = create_label(" Basic Information ", "Class :", "Arial", 15, 200, 405)

            # entry7 = create_entry_box(" Basic Information ", "Eg.D10 C", "Arial", 15, 300, 400, 405)

            entry7 = create_dropdown(" Basic Information ", ["Select Class"], "Arial", 15, 300, None, 400, 405)

            label11 = create_label(" Basic Information ", "Endsems CO's", "Arial", 15, 875, 55)

            entry11 = create_entry_box(" Basic Information ", "1,2,3,4,5,6", "Arial", 15, 300, 1075, 55)

            # label12 = create_label(" Basic Information ", "Attainment Target :", "Arial", 15, 875, 55)

            # entry12 = create_entry_box(" Basic Information ", "52.5", "Arial", 15, 300, 1075, 55)

            label10 = create_label(" Basic Information ", "No. of CA's :", "Arial", 15, 875, 105)

            entry10 = create_dropdown(" Basic Information ", ["2", "3"], "Arial", 15, 300, disable, 1075, 105)

            label13 = create_label(" Basic Information ", "CA1 type :", "Arial", 15, 875, 155)

            entry13 = create_dropdown(" Basic Information ", ["Select Type", "Quiz", "NPTEL Course", "Presentation"], "Arial", 15, 300, ca1, 1075, 155)

            noCA1Label = create_label(" Basic Information ", "No of Question CA1 :", "Arial", 15, 875, 205)

            noCA1Entry = create_dropdown(" Basic Information ", ["Select No", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], "Arial", 15, 300, noQuestion1, 1075, 205)
            noCA1Entry.configure(state="disabled", fg_color="gray")

            label14 = create_label(" Basic Information ", "CA2 type :", "Arial", 15, 875, 255)

            entry14 = create_dropdown(" Basic Information ", ["Select Type", "Quiz", "NPTEL Course", "Presentation"], "Arial", 15, 300, ca2, 1075, 255)

            noCA2Label = create_label(" Basic Information ", "No of Question CA2 :", "Arial", 15, 875, 305)

            noCA2Entry = create_dropdown(" Basic Information ", ["Select No", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], "Arial", 15, 300, noQuestion2, 1075, 305)
            noCA2Entry.configure(state="disabled", fg_color="gray")

            label15 = create_label(" Basic Information ", "CA3 type :", "Arial", 15, 875, 355)

            entry15 = create_dropdown(" Basic Information ", ["Select Type", "Quiz", "NPTEL Course", "Presentation"], "Arial", 15, 300, ca3, 1075, 355)
            entry15.configure(state="disabled", fg_color="gray")

            noCA3Label = create_label(" Basic Information ", "No of Question CA3 :", "Arial", 15, 875, 405)

            noCA3Entry = create_dropdown(" Basic Information ", ["Select No", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], "Arial", 15, 300, noQuestion3, 1075, 405)
            noCA3Entry.configure(state="disabled", fg_color="gray")

            nptel = create_label(" Basic Information ", "CO's for NPTEL (CA)", "Arial", 20, 700, 505)

            nptelCA1Label = create_label(" Basic Information ", "CA1: ", "Arial", 15, 100, 545)

            nptelCA1Text = create_entry_box(" Basic Information ", "1,2,3,4,5,6", "Arial", 15, 300, 200, 545)
            nptelCA1Text.configure(state="disabled", fg_color="gray")

            nptelCA2Label = create_label(" Basic Information ", "CA2: ", "Arial", 15, 550, 545)

            nptelCA2Text = create_entry_box(" Basic Information ", "1,2,3,4,5,6", "Arial", 15, 300, 650, 545)
            nptelCA2Text.configure(state="disabled", fg_color="gray")

            nptelCA3Label = create_label(" Basic Information ", "CA3: ", "Arial", 15, 1000, 545)

            nptelCA3Text = create_entry_box(" Basic Information ", "1,2,3,4,5,6", "Arial", 15, 300, 1100, 545)
            nptelCA3Text.configure(state="disabled", fg_color="gray")

            presentation = create_label(" Basic Information ", "Maximum group size of Presentations (CA)", "Arial", 20, 590, 585)

            presentationCA1Label = create_label(" Basic Information ", "CA1: ", "Arial", 15, 100, 625)

            presentationCA1Text = create_entry_box(" Basic Information ", "Enter maximum number of students in a group", "Arial", 15, 300, 200, 625)
            presentationCA1Text.configure(state="disabled", fg_color="gray")

            presentationCA2Label = create_label(" Basic Information ", "CA2: ", "Arial", 15, 550, 625)

            presentationCA2Text = create_entry_box(" Basic Information ", "Enter maximum number of students in a group", "Arial", 15, 300, 650, 625)
            presentationCA2Text.configure(state="disabled", fg_color="gray")

            presentationCA3Label = create_label(" Basic Information ", "CA3: ", "Arial", 15, 1000, 625)

            presentationCA3Text = create_entry_box(" Basic Information ", "Enter maximum number of students in a group", "Arial", 15, 300, 1100, 625)
            presentationCA3Text.configure(state="disabled", fg_color="gray")


            label6 = create_label(" CO Mapping ", "COs for Midterm", "Arial", 20, 375, 20)

            a1L = create_label(" CO Mapping ", "1a :", "Arial", 15, 200, 60)
            a2L = create_label(" CO Mapping ", "1b :", "Arial", 15, 200, 110)
            a3L = create_label(" CO Mapping ", "1c :", "Arial", 15, 200, 160)
            a4L = create_label(" CO Mapping ", "1d :", "Arial", 15, 200, 210)
            a5L = create_label(" CO Mapping ", "1e :", "Arial", 15, 200, 260)
            a6L = create_label(" CO Mapping ", "1f :", "Arial", 15, 500, 60)
            a2aL = create_label(" CO Mapping ", "2a :", "Arial", 15, 500, 110)
            a2bL = create_label(" CO Mapping ", "2b :", "Arial", 15, 500, 160)
            a3aL = create_label(" CO Mapping ", "3a :", "Arial", 15, 500, 210)
            a3bL = create_label(" CO Mapping ", "3b :", "Arial", 15, 500, 260)


            a1T=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            a1T.place(x=250,y=60)

            a2T=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            a2T.place(x=250,y=110)

            a3T=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            a3T.place(x=250,y=160)

            a4T=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            a4T.place(x=250,y=210)

            a5T=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            a5T.place(x=250,y=260)

            a6T=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            a6T.place(x=550,y=60)


            a2aT=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            a2aT.place(x=550,y=110)


            a2bT=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            a2bT.place(x=550,y=160)


            a3aT=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            a3aT.place(x=550,y=210)


            a3bT=ctk.CTkEntry(master=tabview.tab(" CO Mapping "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            a3bT.place(x=550,y=260)

            #For Quiz

            # COs for CA1 Quiz
            label9 = create_label(" CO Mapping ", "COs for CA1 Quiz", "Arial", 20, 375, 340)

            q1LCA1 = create_label(" CO Mapping ", "Q1 :", "Arial", 15, 200, 380)
            q1TCA1 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 380)
            q1TCA1.configure(state="disabled", fg_color="gray")

            q2LCA1 = create_label(" CO Mapping ", "Q2 :", "Arial", 15, 200, 430)
            q2TCA1 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 430)
            q2TCA1.configure(state="disabled", fg_color="gray")

            q3LCA1 = create_label(" CO Mapping ", "Q3 :", "Arial", 15, 200, 480)
            q3TCA1 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 480)
            q3TCA1.configure(state="disabled", fg_color="gray")

            q4LCA1 = create_label(" CO Mapping ", "Q4 :", "Arial", 15, 200, 530)
            q4TCA1 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 530)
            q4TCA1.configure(state="disabled", fg_color="gray")

            q5LCA1 = create_label(" CO Mapping ", "Q5 :", "Arial", 15, 200, 580)
            q5TCA1 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 580)
            q5TCA1.configure(state="disabled", fg_color="gray")

            q6LCA1 = create_label(" CO Mapping ", "Q6 :", "Arial", 15, 500, 380)
            q6TCA1 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 380)
            q6TCA1.configure(state="disabled", fg_color="gray")

            q7LCA1 = create_label(" CO Mapping ", "Q7 :", "Arial", 15, 500, 430)
            q7TCA1 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 430)
            q7TCA1.configure(state="disabled", fg_color="gray")

            q8LCA1 = create_label(" CO Mapping ", "Q8 :", "Arial", 15, 500, 480)
            q8TCA1 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 480)
            q8TCA1.configure(state="disabled", fg_color="gray")

            q9LCA1 = create_label(" CO Mapping ", "Q9 :", "Arial", 15, 500, 530)
            q9TCA1 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 530)
            q9TCA1.configure(state="disabled", fg_color="gray")

            q10LCA1 = create_label(" CO Mapping ", "Q10 :", "Arial", 15, 500, 580)
            q10TCA1 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 580)
            q10TCA1.configure(state="disabled", fg_color="gray")

            # COs for CA2 Quiz
            label18 = create_label(" CO Mapping ", "COs for CA2 Quiz", "Arial", 20, 1025, 20)

            q1LCA2 = create_label(" CO Mapping ", "Q1 :", "Arial", 15, 850, 60)
            q1TCA2 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 900, 60)
            q1TCA2.configure(state="disabled", fg_color="gray")

            q2LCA2 = create_label(" CO Mapping ", "Q2 :", "Arial", 15, 850, 110)
            q2TCA2 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 900, 110)
            q2TCA2.configure(state="disabled", fg_color="gray")

            q3LCA2 = create_label(" CO Mapping ", "Q3 :", "Arial", 15, 850, 160)
            q3TCA2 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 900, 160)
            q3TCA2.configure(state="disabled", fg_color="gray")

            q4LCA2 = create_label(" CO Mapping ", "Q4 :", "Arial", 15, 850, 210)
            q4TCA2 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 900, 210)
            q4TCA2.configure(state="disabled", fg_color="gray")

            q5LCA2 = create_label(" CO Mapping ", "Q5 :", "Arial", 15, 850, 260)
            q5TCA2 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 900, 260)
            q5TCA2.configure(state="disabled", fg_color="gray")

            q6LCA2 = create_label(" CO Mapping ", "Q6 :", "Arial", 15, 1150, 60)
            q6TCA2 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 1200, 60)
            q6TCA2.configure(state="disabled", fg_color="gray")

            q7LCA2 = create_label(" CO Mapping ", "Q7 :", "Arial", 15, 1150, 110)
            q7TCA2 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 1200, 110)
            q7TCA2.configure(state="disabled", fg_color="gray")

            q8LCA2 = create_label(" CO Mapping ", "Q8 :", "Arial", 15, 1150, 160)
            q8TCA2 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 1200, 160)
            q8TCA2.configure(state="disabled", fg_color="gray")

            q9LCA2 = create_label(" CO Mapping ", "Q9 :", "Arial", 15, 1150, 210)
            q9TCA2 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 1200, 210)
            q9TCA2.configure(state="disabled", fg_color="gray")

            q10LCA2 = create_label(" CO Mapping ", "Q10 :", "Arial", 15, 1150, 260)
            q10TCA2 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 1200, 260)
            q10TCA2.configure(state="disabled", fg_color="gray")

            # COs for CA3 Quiz
            label21 = create_label(" CO Mapping ", "COs for CA3 Quiz", "Arial", 20, 1025, 340)

            q1LCA3 = create_label(" CO Mapping ", "Q1 :", "Arial", 15, 850, 380)
            q1TCA3 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 900, 380)
            q1TCA3.configure(state="disabled", fg_color="gray")

            q2LCA3 = create_label(" CO Mapping ", "Q2 :", "Arial", 15, 850, 430)
            q2TCA3 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 900, 430)
            q2TCA3.configure(state="disabled", fg_color="gray")

            q3LCA3 = create_label(" CO Mapping ", "Q3 :", "Arial", 15, 850, 480)
            q3TCA3 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 900, 480)
            q3TCA3.configure(state="disabled", fg_color="gray")

            q4LCA3 = create_label(" CO Mapping ", "Q4 :", "Arial", 15, 850, 530)
            q4TCA3 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 900, 530)
            q4TCA3.configure(state="disabled", fg_color="gray")

            q5LCA3 = create_label(" CO Mapping ", "Q5 :", "Arial", 15, 850, 580)
            q5TCA3 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 900, 580)
            q5TCA3.configure(state="disabled", fg_color="gray")

            q6LCA3 = create_label(" CO Mapping ", "Q6 :", "Arial", 15, 1150, 380)
            q6TCA3 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 1200, 380)
            q6TCA3.configure(state="disabled", fg_color="gray")

            q7LCA3 = create_label(" CO Mapping ", "Q7 :", "Arial", 15, 1150, 430)
            q7TCA3 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 1200, 430)
            q7TCA3.configure(state="disabled", fg_color="gray")

            q8LCA3 = create_label(" CO Mapping ", "Q8 :", "Arial", 15, 1150, 480)
            q8TCA3 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 1200, 480)
            q8TCA3.configure(state="disabled", fg_color="gray")

            q9LCA3 = create_label(" CO Mapping ", "Q9 :", "Arial", 15, 1150, 530)
            q9TCA3 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 1200, 530)
            q9TCA3.configure(state="disabled", fg_color="gray")

            q10LCA3 = create_label(" CO Mapping ", "Q10 :", "Arial", 15, 1150, 580)
            q10TCA3 = create_entry_box(" CO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 1200, 580)
            q10TCA3.configure(state="disabled", fg_color="gray")


            button = create_button(" Target level of tests ", "Download", "Arial", 20, 200, 40, download, 650, 500)

            def upload_file():
                global file_path
                file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
                print(f"Upload function: {file_path}")
                if file_path:
                    print(file_path)
                    file_name = os.path.basename(file_path)
                    print(file_name)
                    path_label.configure(text=file_name)


            def process_file():
                global file_path
                if entry10.get() == "2":
                    al_values=[ALCA1Text.get(), ALCA2Text.get(), '-', ALMidTermText.get(), ALEndSemText.get()]
                    print(al_values)
                else:
                    al_values=[ALCA1Text.get(), ALCA2Text.get(), ALCA3Text.get(), ALMidTermText.get(), ALEndSemText.get()]
                file_path = file_path
                print("File Path : : : ", file_path)
                import Cal
                Cal.cal_sheet(file_path, al_values)

            # Using create_label, create_entry_box, and create_dropdown to recreate the UI

            # CO Information
            enterCO = create_label(" CO Information ", "Enter the CO's", "Arial", 20, 700, 50)
            noOfCOLabel = create_label(" CO Information ", "Select No. of CO's: ", "Arial", 15, 550, 100)
            noOfCOOption = create_dropdown(" CO Information ", ['Select No of CO\'s', '5', '6'], "Arial", 15, 300, noOfCO, 750, 100)

            CO1L = create_label(" CO Information ", "CO1: ", "Arial", 15, 550, 150)
            CO1T = create_entry_box(" CO Information ", "", "Arial", 15, 500, 650, 150)

            CO2L = create_label(" CO Information ", "CO2: ", "Arial", 15, 550, 200)
            CO2T = create_entry_box(" CO Information ", "", "Arial", 15, 500, 650, 200)

            CO3L = create_label(" CO Information ", "CO3: ", "Arial", 15, 550, 250)
            CO3T = create_entry_box(" CO Information ", "", "Arial", 15, 500, 650, 250)

            CO4L = create_label(" CO Information ", "CO4: ", "Arial", 15, 550, 300)
            CO4T = create_entry_box(" CO Information ", "", "Arial", 15, 500, 650, 300)

            CO5L = create_label(" CO Information ", "CO5: ", "Arial", 15, 550, 350)
            CO5T = create_entry_box(" CO Information ", "", "Arial", 15, 500, 650, 350)

            CO6L = create_label(" CO Information ", "CO6: ", "Arial", 15, 550, 400)
            CO6T = create_entry_box(" CO Information ", "", "Arial", 15, 500, 650, 400)
            CO6T.configure(state="disabled", fg_color="gray")

            # Target level of tests
            ALlabel = create_label(" Target level of tests ", "Enter the AL targets for each exam", "Arial", 20, 600, 50)

            # CA1, CA2, CA3, MidTerm, EndSem, Labs
            ALCA1Label = create_label(" Target level of tests ", "CA1: ", "Arial", 15, 450, 100)
            ALCA1Text = create_entry_box(" Target level of tests ", "", "Arial", 15, 500, 575, 100)

            ALCA2Label = create_label(" Target level of tests ", "CA2: ", "Arial", 15, 450, 150)
            ALCA2Text = create_entry_box(" Target level of tests ", "", "Arial", 15, 500, 575, 150)

            ALCA3Label = create_label(" Target level of tests ", "CA3: ", "Arial", 15, 450, 200)
            ALCA3Text = create_entry_box(" Target level of tests ", "", "Arial", 15, 500, 575, 200)
            ALCA3Text.configure(state="disabled", fg_color="gray")

            ALMidTermLabel = create_label(" Target level of tests ", "Mid Term: ", "Arial", 15, 450, 250)
            ALMidTermText = create_entry_box(" Target level of tests ", "", "Arial", 15, 500, 575, 250)

            ALEndSemLabel = create_label(" Target level of tests ", "End Semester: ", "Arial", 15, 450, 300)
            ALEndSemText = create_entry_box(" Target level of tests ", "", "Arial", 15, 500, 575, 300)

            setEmailLabel = create_label(" Target level of tests ", "Enter the Email ID to send the template.", "Arial", 20, 600, 400)

            emailText = create_entry_box(" Target level of tests ", "", "Arial", 15, 500, 525, 450)

            # ALSurveyLabel = create_label(" Target level of tests ", "Survey: ", "Arial", 15, 450, 350)
            # ALSurveyText = create_entry_box(" Target level of tests ", "", "Arial", 15, 500, 575, 350)

            # Buttons
            button1 = create_button(" CO Information ", "Next", "Arial", 20, 200, 40, switch1, 725, 500)
            # button2 = create_button(" CO Mapping ", "Next", "Arial", 20, 200, 40, switch2, 725, 500)
            button2 = create_button(" CO Mapping ", "Next", "Arial", 20, 200, 40, switch2, 1000, 640)



            path_entry=ctk.CTkEntry(tabview.tab(" Upload Excel File "))

            # button_process=ctk.CTkButton(tabview.tab(" Upload Excel File "),text="Process",width=100,height=30,command=process_file)
            # button_process.place(x=500,y=500)

            upload_Label = create_label(" Upload Excel File ", "Upload you excel file with the marks entered:", "Arial", 25, 550, 50)
            path_label = create_label(" Upload Excel File ", "Path of file", "Arial", 15, 650, 110)
            button_upload = create_button(" Upload Excel File ", "Upload", "Arial", 20, 200, 40, upload_file, 400, 100)


            line = ctk.CTkFrame(master=tabview.tab(" Upload Excel File "), height=2, width=1200, fg_color="white")
            line.place(x=150,y=200)

            process_Label = create_label(" Upload Excel File ", "Process the excel file you uploaded:", "Arial", 25, 600, 250)

            setEmailProcessedLabel = create_label(" Upload Excel File ", "Enter the Email ID to send the calculated sheet.", "Arial", 20, 200, 400)

            important_label = create_label(" Upload Excel File ", "Important: Please fill the no. of CO\'s field and the CO\'s in the CO Information page and AL values in Target level of tests page before processing the file", "Arial", 20, 100, 325)
            important_label.configure(text_color="black", fg_color="yellow")

            emailTextProcessed = create_entry_box(" Upload Excel File ", "", "Arial", 15, 500, 700, 400)

            button_process = create_button(" Upload Excel File ", "Process", "Arial", 20, 200, 40, process_file, 650, 500)

            back_button = ctk.CTkButton(co_window, text="Back", command=lambda: self.go_back(co_window))
            back_button.pack(pady=20)

            co_window.mainloop()

    def open_lo_window(self):
            
            def subject(option):
                if option == "Select Sem":
                    # entry3.configure(values=["Select Subject"])
                    entry3_lab.configure(values=["Select Subject"])
                elif option == "I":
                    # entry3.configure(values=["Select Subject","Universal Human Values - 1","Fundamentals of Vedic Mathematics (Indian Knowledge System)", "Basic Electrical Engineering", "Engineering Drawing", "Engineering Mechanics", "Engineering Physics", "Matrices and Differential Calculus", "Python Programming"])
                    entry3_lab.configure(values=["Select Subject","Universal Human Values - 1","Fundamentals of Vedic Mathematics (Indian Knowledge System)", "Basic Electrical Engineering", "Engineering Drawing", "Engineering Mechanics", "Engineering Physics", "Matrices and Differential Calculus", "Python Programming"])
                elif option == "II":
                    # entry3.configure(values=["Select Subject","Universal Human Values - 2","Basic Workshop Practice", "Computer Programming", "Integral Calculus and Complex Numbers", "Biology for Engineers", "Engineering Chemistry", "Professional Communication and Ethics - 1"])
                    entry3_lab.configure(values=["Select Subject","Universal Human Values - 2","Basic Workshop Practice", "Computer Programming", "Integral Calculus and Complex Numbers", "Biology for Engineers", "Engineering Chemistry", "Professional Communication and Ethics - 1"])
                elif option == "III":
                    # entry3.configure(values=["Select Subject","Engineering Mathematics III", "Data Structures and Analysis", "Database Management System", "Principle of Communications", "Paradigm and computer programming fundamentals"])
                    entry3_lab.configure(values=["Select Subject","Engineering Mathematics III", "Data Structures and Analysis", "Database Management System", "Principle of Communications", "Paradigm and computer programming fundamentals"])
                elif option == "IV":
                    # entry3.configure(values=["Select Subject","Engineering Mathematics IV", "Computer Network and Network Design", "Operating System", "Automata Theory", "Computer Organization and Architecture"])
                    entry3_lab.configure(values=["Select Subject","Engineering Mathematics IV", "Computer Network and Network Design", "Operating System", "Automata Theory", "Computer Organization and Architecture"])
                elif option == "V":
                    # entry3.configure(values=["Select Subject","Internet Programming", "Computer Network Security", "Entrepreneurship and E- business", "Software Engineering", "Advance Data Management Technologies", "Advanced Data structure and Analysis"])
                    entry3_lab.configure(values=["Select Subject","Internet Programming", "Computer Network Security", "Entrepreneurship and E- business", "Software Engineering", "Advance Data Management Technologies", "Advanced Data structure and Analysis"])
                elif option == "VI":
                    # entry3.configure(values=["Select Subject","Data Mining & Business Intelligence", "Web X.0", "Wireless Technology", "AI and DS 1", "Optional Course 2"])
                    entry3_lab.configure(values=["Select Subject","Data Mining & Business Intelligence", "Web X.0", "Wireless Technology", "AI and DS 1", "Optional Course 2"])
                elif option == "VII":
                    # entry3.configure(values=["Select Subject","AI and DS II", "Internet of Everything", "Department Optional Course 3", "Department Optional Course 4", "Institute Optional Course 1"])
                    entry3_lab.configure(values=["Select Subject","AI and DS II", "Internet of Everything", "Department Optional Course 3", "Department Optional Course 4", "Institute Optional Course 1"])
                elif option == "VIII":
                    # entry3.configure(values=["Select Subject","Blockchain and DLT", "Department Optional Course 5", "Department Optional Course 6", "Institute Optional Course 2"])
                    entry3_lab.configure(values=["Select Subject","Blockchain and DLT", "Department Optional Course 5", "Department Optional Course 6", "Institute Optional Course 2"])
            
            def validate_academic_year(event):
                new_value = event.widget.get()
                if new_value:
                    if len(new_value) != 9 or new_value[4] != '-' or not (new_value[:4].isdigit() and new_value[5:].isdigit()):
                        CTkMessagebox(title="Invalid Input", message="Academic Year format is incorrect. Please enter in the format YYYY-YYYY.", icon="warning")                
                        return False
                    else: 
                        return True
            
            def lo_check(option):
                if option == "Select No of LO\'s":
                    for entry in LO_entry.values():
                        entry.configure(state="disabled", fg_color="gray")
                else:
                    option = int(option)
                    for i in range (1, option+1):
                        LO_entry[f"LO{i}_entry"].configure(state="normal", fg_color=["F9F9FA", "#343638"])
                    for i in range (option+1, 7):
                        LO_entry[f"LO{i}_entry"].configure(state="disabled", fg_color="gray")

            def exp_group(option):
                factor_title_entry = [factor_1_title_entry, factor_2_title_entry, factor_3_title_entry, factor_4_title_entry, factor_5_title_entry]
                factor_lo_entry = [factor_1_lo_entry, factor_2_lo_entry, factor_3_lo_entry, factor_4_lo_entry, factor_5_lo_entry]
                if option == "Group Students":
                    group_size_exp_entry.configure(state="normal", fg_color=["F9F9FA", "#343638"])
                    for entry in factor_title_entry:
                        entry.configure(state="normal", fg_color=["F9F9FA", "#343638"])
                    for entry in factor_lo_entry:
                        entry.configure(state="normal", fg_color=["F9F9FA", "#343638"])
                    for entry in exp_lo_entry.values():
                        entry.configure(state="disabled", fg_color="gray")

                else:
                    group_size_exp_entry.configure(state="disabled",fg_color="gray")
                    for entry in factor_title_entry:
                        entry.configure(state="disabled",fg_color="gray")
                    for entry in factor_lo_entry:
                        entry.configure(state="disabled",fg_color="gray")
                if option == "Individual Students":
                    no_of_exp_dropdown.configure(state="normal", fg_color=["F9F9FA", "#1F6AA5"])
                else:
                    no_of_exp_dropdown.configure(state="disabled", fg_color="gray")
                    for entry in exp_lo_entry.values():
                        entry.configure(state="disabled", fg_color="gray")

            def exp_fields(option):
                if option == "Select no of experiments":
                    for entry in exp_lo_entry.values():
                        entry.configure(state="disabled", fg_color="gray")
                else:
                    option = int(option)
                    for i in range (1, option+1):
                        exp_lo_entry[f"exp{i}_lo_entry"].configure(state="normal", fg_color=["F9F9FA", "#343638"])
                    for i in range (option+1, 16):
                        exp_lo_entry[f"exp{i}_lo_entry"].configure(state="disabled", fg_color="gray")

            def assignment_fields(option):
                if option == "Select no of Assignments":
                    for entry in assignment_lo_entry.values():
                        entry.configure(state="disabled", fg_color="gray")
                else:
                    option = int(option)
                    for i in range (1, option+1):
                        assignment_lo_entry[f"assignment_{i}_lo_entry"].configure(state="normal", fg_color=["F9F9FA", "#343638"])
                    for i in range (option+1, 4):
                        assignment_lo_entry[f"assignment_{i}_lo_entry"].configure(state="disabled", fg_color="gray")

            def validate_lo_string(loString):
                validate_lo_array = []
                # print(coString)
                loString = loString.replace(" ", "")
                # print(coString)
                if noOfLOOption.get() == "Select No of CO\'s":
                    CTkMessagebox(title = "Error", message="Select No of CO\'s", icon="cancel")
                elif noOfLOOption.get() == "5":
                    validate_lo_array = [1,2,3,4,5]
                elif noOfLOOption.get() == "6":
                    validate_lo_array = [1,2,3,4,5,6]

                pattern = r'^(\d,)*\d$'
                if not re.match(pattern, loString):
                    return False

                # Extract digits from the input string
                digits = list(map(int, loString.split(',')))
                # print(digits)

                # Check each digit is within the valid_digits array
                if not all(digit in validate_lo_array for digit in digits):
                    return False

                # Ensure there are no consecutive identical digits
                if len(digits) != len(set(digits)):
                    return False

                return True

            def switch_lab():
                if noOfLOOption.get() == "Select No of LO\'s":
                    CTkMessagebox(title = "Error", message="Select No of LO\'s", icon="cancel")
                else:
                    option = noOfLOOption.get()
                    option = int(option)
                    for i in range (1, option+1):
                        if LO_entry[f"LO{i}_entry"].get() == "":
                            CTkMessagebox(title = "Error", message="Please fill all the LO\'s", icon="cancel")
                            return
                    tabview.set(" LO Information Template generation ")

            def switch_1_lab():
                if noOfLOOption.get() == "Select No of LO\'s":
                    CTkMessagebox(title = "Error", message="Select No of LO\'s on the previous page", icon="cancel")
                if (entry1_lab.get() == "" or yearDropDown_lab.get() == "Select Year" or entry8_lab.get() == "Select Department" or entry2_lab.get() == "Select Sem" or entry3_lab.get() == "Select Subject" or entry4_lab.get() == "" or entry5_lab.get() == "" or entry7_lab.get() == "Select Class"):
                    CTkMessagebox(title = "Error", message="Please fill all the basic details", icon="cancel")
                    return
                elif no_of_assignments_dropdown.get() == "Select no of Assignments":
                    CTkMessagebox(title = "Error", message="Please select the number of assignments", icon="cancel")
                    return
                elif (oral_marks_target_entry.get() == "" or mini_project_marks_target_entry.get() == "" or term_work_marks_target_entry.get() == "" or assignment_target_entry.get() == ""):
                    CTkMessagebox(title = "Error", message="Please fill all the marks targets", icon="cancel")
                    return
                elif termWork_dropdown.get() == "Select Type":
                    CTkMessagebox(title = "Error", message="Please select the type of term work", icon="cancel")
                    return
                elif termWork_dropdown.get() == "Group Students":
                    size = group_size_exp_entry.get()
                    if size == "":
                        CTkMessagebox(title = "Error", message="Please fill the group size", icon="cancel")
                        return
                    elif (not (size.isdigit()) or int(size) < 0):
                        CTkMessagebox(title = "Error", message="Please enter a valid group size", icon="cancel")
                        return
                    elif (factor_1_title_entry.get() == "" or factor_2_title_entry.get() == "" or factor_3_title_entry.get() == "" or factor_4_title_entry.get() == "" or factor_5_title_entry.get() == ""):
                        CTkMessagebox(title = "Error", message="Please fill all the factors", icon="cancel")
                        return
                    check_text_group_LO = [factor_1_lo_entry.get(), factor_2_lo_entry.get(), factor_3_lo_entry.get(), factor_4_lo_entry.get(), factor_5_lo_entry.get()]
                    if "" in check_text_group_LO:
                        CTkMessagebox(title = "Error", message="Please fill all LO\'s for the factors", icon="cancel")
                        return
                    for text in check_text_group_LO:
                        if not validate_lo_string(text):
                            CTkMessagebox(title = "Error", message="Please fill LO\'s in valid format", icon="cancel")
                            return
                elif termWork_dropdown.get() == "Individual Students":
                    if no_of_exp_dropdown.get() == "Select no of experiments":
                        CTkMessagebox(title = "Error", message="Please select the number of experiments", icon="cancel")
                        return
                tabview.set(" LO Mapping ")



            def switch_2_lab():
                tabview.set(" Upload Excel File (Lab) ")

            def download_template_lab():
                if(noOfLOOption.get() == "Select No of LO\'s"):
                    CTkMessagebox(title = "Error", message="Enter No of LOs in page 1", icon="cancel")

                LOcount = int(noOfLOOption.get())

                if(termWork_dropdown.get() == "Individual Students"):
                    exp_no = []
                    for i in range (0,int(no_of_exp_dropdown)):
                        exp_no.append(exp_lo_entry[f"exp{i+1}_lo_entry"].get())
                    for text in exp_no:
                        if not (validate_lo_string(text)):
                            CTkMessagebox(title = "Error", message="Enter all LO\'s of Experiments", icon="cancel")
                            return
                factor_no = []
                for i in range (0,4):
                    factor_no.append(mini_project_lo_entry[f"mini_project_factor{i+1}_lo_entry"].get())
                for text in factor_no:
                    if not (validate_lo_string(text)):
                        CTkMessagebox(title = "Error", message="Enter all LO\'s of Factors of Mini Projects", icon="cancel")
                        return
                assignment_lo = []
                for i in range (0, int(no_of_assignments_dropdown)):
                    assignment_lo.appned(assignment_lo_entry[f'assignment_{i+1}_lo_entry'].get())
                for text in assignment_lo:
                    if not (validate_lo_string(text)):
                        CTkMessagebox(title = "Error", message="Enter all LO\'s of Assignments", icon="cancel")
                        return
  
                           
                
                loTextArray = [""]
                if noOfLOOption.get() == "5":
                    loTextArray = [LO_entry["LO1_entry"].get(), LO_entry["LO2_entry"].get(), LO_entry["LO3_entry"].get(), LO_entry["LO4_entry"].get(), LO_entry["LO5_entry"].get(), "-"]
                elif noOfLOOption.get() == "6":
                    loTextArray = [LO_entry["LO1_entry"].get(), LO_entry["LO2_entry"].get(), LO_entry["LO3_entry"].get(), LO_entry["LO4_entry"].get(), LO_entry["LO5_entry"].get(), LO_entry["LO6_entry"].get()]
                else:
                    CTkMessagebox(title = "Error", message="Select No of LO\'s", icon="cancel")

                basic_values_lo = [entry3_lab.get(), entry7_lab.get(), entry8_lab.get(), entry4_lab.get(), entry2_lab.get(), entry5_lab.get(), entry1_lab.get(), noOfLOOption.get(), term_work_marks_target_entry.get(), oral_marks_target_entry.get(), assignment_target_entry.get(), mini_project_marks_target_entry.get(), termWork_dropdown.get(), no_of_assignments_dropdown.get(), loTextArray, no_of_exp_dropdown.get()]
                exp_lo = []
                if no_of_exp_dropdown.get() != "Select no of experiments":
                    for i in range (0, int(no_of_exp_dropdown.get())):
                        exp_lo.append(exp_lo_entry[f"exp{i+1}_lo_entry"].get())
                basic_values_lo.append(exp_lo)
                basic_values_lo.append(group_size_exp_entry.get())
                critList = [factor_1_title_entry.get(), factor_2_title_entry.get(), factor_3_title_entry.get(), factor_4_title_entry.get(), factor_5_title_entry.get()]
                loList = [factor_1_lo_entry.get(), factor_2_lo_entry.get(), factor_3_lo_entry.get(), factor_4_lo_entry.get(), factor_5_lo_entry.get()]
                basic_values_lo.append(critList)
                basic_values_lo.append(loList)
                basic_values_lo.append(group_size_mini_project_entry.get())
                projLoList = [mini_project_lo_entry[f"mini_project_factor{i}_lo_entry"].get() for i in range (1, 5)]
                basic_values_lo.append(projLoList)
                assignmentLOs = [assignment_lo_entry[f"assignment_{i}_lo_entry"].get() for i in range (1, (int(no_of_assignments_dropdown.get()))+1)]
                basic_values_lo.append(assignmentLOs)
                print(basic_values_lo)
                from lab.Lab_Template import lab_template_generator
                lab_template_generator(basic_values_lo)


            def upload_lab_file():
                global file_path_lab
                file_path_lab = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
                print(f"Upload function: {file_path_lab}")
                if file_path_lab:
                    print(file_path_lab)
                    file_name_lab = os.path.basename(file_path_lab)
                    print(file_name_lab)
                    path_label_lab.configure(text=file_name_lab)



            def process_file_lab():
                global file_path_lab
                file_path_lab = file_path_lab
                from lab.Lab_Cal import cal_lab_sheets
                cal_lab_sheets(file_path_lab)
            
            def semesterAndClass(option):
                if option == "Select Year":
                    # entry2.configure(values=["Select Sem"])
                    # entry7.configure(values=["Select Class"])
                    entry2_lab.configure(values=["Select Sem"])
                    entry7_lab.configure(values=["Select Class"])
                elif option == "F.E":
                    # entry2.configure(values=["Select Sem","I","II"])
                    # entry7.configure(values=["Select Class", "D5A", "D5B", "D5C"])
                    entry2_lab.configure(values=["Select Sem","I","II"])
                    entry7_lab.configure(values=["Select Class", "D5A", "D5B", "D5C"])
                elif option == "S.E":
                    # entry2.configure(values=["Select Sem","III","IV"])
                    # entry7.configure(values=["Select Class", "D10A", "D10B", "D10C"])
                    entry2_lab.configure(values=["Select Sem","III","IV"])
                    entry7_lab.configure(values=["Select Class", "D10A", "D10B", "D10C"])
                elif option == "T.E":
                    # entry2.configure(values=["Select Sem","V","VI"])
                    # entry7.configure(values=["Select Class", "D15A", "D15B", "D15C"])
                    entry2_lab.configure(values=["Select Sem","V","VI"])
                    entry7_lab.configure(values=["Select Class", "D15A", "D15B", "D15C"])
                elif option == "B.E":
                    # entry2.configure(values=["Select Sem","VII","VIII"])
                    # entry7.configure(values=["Select Class", "D20A", "D20B", "D20C"])
                    entry2_lab.configure(values=["Select Sem","VII","VIII"])
                    entry7_lab.configure(values=["Select Class", "D20A", "D20B", "D20C"])
            
            def create_button(tab, name, font_name, font_size, w, h, com, x, y):
                button = ctk.CTkButton(master=tabview.tab(tab), text=name, width=w, height=h, font=(font_name, font_size), command=com)
                button.place(x=x, y=y)
                return button

            def create_label(tab, name, font_type, font_size, x, y):
                label = ctk.CTkLabel(master=tabview.tab(tab), text=name, font=(font_type, font_size))
                label.place(x=x, y=y)
                return label

            def create_entry_box(tab, text, font_name, font_size, w, x, y):
                entry_box = ctk.CTkEntry(master=tabview.tab(tab), placeholder_text=text, font=(font_name,font_size), width=w)
                entry_box.place(x=x,y=y)
                return entry_box

            def create_dropdown(tab, val, font_name, font_size, w, com, x, y):
                dropdown = ctk.CTkOptionMenu(master=tabview.tab(tab), values=val, font=(font_name, font_size), width=w, command=com)
                dropdown.place(x=x, y=y)
                return dropdown
            
            self.app.destroy() 
            
            lo_window = ctk.CTk()  # Close the current window

            screen_width=lo_window.winfo_screenwidth()
            screen_height=lo_window.winfo_screenheight()
       
        # Calculate the coordinates for centering the window
            x_losition = 0
            y_losition = 0

             # Create a new lo Calculations window
            lo_window.title("lo Calculations")
            lo_window.geometry(f"{screen_width}x{screen_height}+{x_losition}+{y_losition}")

            main_frame = ctk.CTkFrame(master=lo_window)
            main_frame.pack(expand=True, fill="both", padx=10, pady=10)

            # Tabview inside the frame
            tabview = ctk.CTkTabview(main_frame, corner_radius=20)
            tabview.pack(expand=True, fill="both", padx=10, pady=5)

            tabview.add(" LO Information ")
            tabview.add(" LO Information Template generation ")
            tabview.add(" LO Mapping ")
            tabview.add(" Upload Excel File (Lab) ")      


            label0_lab = create_label(" LO Information Template generation ", "Basic Details", "Arial", 20, 325, 5)

            label1_lab = create_label(" LO Information Template generation ", "No. of Students :", "Arial", 15, 100, 55)

            entry1_lab = create_entry_box(" LO Information Template generation ", "Enter no of students", "Arial", 15, 300, 400, 55)

            newLabel_lab = create_label(" LO Information Template generation ", "Year :", "Arial", 15, 100, 155)

            yearDropDown_lab = create_dropdown(" LO Information Template generation ", ["Select Year", "F.E", "S.E", "T.E", "B.E"], "Arial", 15, 300, semesterAndClass, 400, 155)

            label8_lab = create_label(" LO Information Template generation ", "Department :", "Arial", 15, 100, 105)

            entry8_lab = create_dropdown(" LO Information Template generation ", ["Select Department", "Humanities and Applied Science(FE)", "Information Technology", "Computer", "AI and Data Science", "Electronics and Telecommunication", "Electronics", "Instrumentation"], "Arial", 15, 300, None, 400, 105)

            label2_lab = create_label(" LO Information Template generation ", "Semester :", "Arial", 15, 100, 205)

            entry2_lab = create_dropdown(" LO Information Template generation ", ["Select Sem"], "Arial", 15, 300, subject, 400, 205)

            label3_lab = create_label(" LO Information Template generation ", "Subject :", "Arial", 15, 100, 255)

            entry3_lab = create_dropdown(" LO Information Template generation ", ["Select Subject"], "Arial", 15, 300, None, 400, 255)

            label4_lab = create_label(" LO Information Template generation ", "Academic Year: ", "Arial", 15, 100, 305)

            entry4_lab = create_entry_box(" LO Information Template generation ", "YYYY-YYYY", "Arial", 15, 300, 400, 305)
            entry4_lab.bind("<FocusOut>", validate_academic_year)

            label5_lab = create_label(" LO Information Template generation ", "Subject Teacher :", "Arial", 15, 100, 355)

            entry5_lab = create_entry_box(" LO Information Template generation ", "Subject Teacher", "Arial", 15, 300, 400, 355)

            label7_lab = create_label(" LO Information Template generation ", "Class :", "Arial", 15, 100, 405)

            # entry7 = create_entry_box(" LO Information Template generation ", "Eg.D10 C", "Arial", 15, 300, 400, 405)

            entry7_lab = create_dropdown(" LO Information Template generation ", ["Select Class"], "Arial", 15, 300, None, 400, 405)

            label9_lab = create_label(" LO Information Template generation ", "Type of Term Work : ", "Arial", 15, 100, 455)

            termWork_dropdown = create_dropdown(" LO Information Template generation ", ["Select Type", "Group Students", "Individual Students"], "Arial", 15, 300, exp_group, 400, 455)

            group_size_exp_label = create_label(" LO Information Template generation ", "Enter max no of students in group: ", "Arial", 15, 100, 505)

            group_size_exp_entry = create_entry_box(" LO Information Template generation ", "", "Arial", 15, 300, 400, 505)
            group_size_exp_entry.configure(state="disabled",fg_color="gray")

            no_of_exp_label = create_label(" LO Information Template generation ", "Enter no. of experiments: ", "Arial", 15, 100, 555)

            no_of_exp_dropdown = create_dropdown(" LO Information Template generation ", ["Select no of experiments","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15"], "Arial", 15, 300, exp_fields, 400, 555)
            no_of_exp_dropdown.configure(state="disabled",fg_color="gray")

            no_of_assignments_label = create_label(" LO Information Template generation ", "Enter no. of Assignments: ", "Arial", 15, 100, 605)

            no_of_assignments_dropdown = create_dropdown(" LO Information Template generation ", ["Select no of Assignments", "2", "3"], "Arial", 15, 300, assignment_fields, 400, 605)

            group_size_mini_project_label = create_label(" LO Information Template generation ", "Enter max no of students in group \n(Mini Project): ", "Arial", 15, 100, 655)

            group_size_mini_project_entry = create_entry_box(" LO Information Template generation ", "", "Arial", 15, 300, 400, 655)

            label_10_lab = create_label(" LO Information Template generation ", "Enter factors for group term work and their LO\'s ", "Arial", 20, 875, 5)

            label_11_lab = create_label(" LO Information Template generation ", "Factor No. ", "Arial", 15, 825, 55)

            label_12_lab = create_label(" LO Information Template generation ", "Factor Title ", "Arial", 15, 1025, 55)

            label_13_lab = create_label(" LO Information Template generation ", "Corresponding LO ", "Arial", 15, 1225, 55)

            factor_1_label = create_label(" LO Information Template generation ", "1", "Arial", 15, 825, 105)

            factor_2_label = create_label(" LO Information Template generation ", "2", "Arial", 15, 825, 155)

            factor_3_label = create_label(" LO Information Template generation ", "3", "Arial", 15, 825, 205)

            factor_4_label = create_label(" LO Information Template generation ", "4", "Arial", 15, 825, 255)

            factor_5_label = create_label(" LO Information Template generation ", "5", "Arial", 15, 825, 305)

            factor_1_title_entry = create_entry_box(" LO Information Template generation ", "Enter title", "Arial", 15, 150, 1025, 105)
            factor_1_title_entry.configure(state="disabled",fg_color="gray")

            factor_2_title_entry = create_entry_box(" LO Information Template generation ", "Enter title", "Arial", 15, 150, 1025, 155)
            factor_2_title_entry.configure(state="disabled",fg_color="gray")

            factor_3_title_entry = create_entry_box(" LO Information Template generation ", "Enter title", "Arial", 15, 150, 1025, 205)
            factor_3_title_entry.configure(state="disabled",fg_color="gray")

            factor_4_title_entry = create_entry_box(" LO Information Template generation ", "Enter title", "Arial", 15, 150, 1025, 255)
            factor_4_title_entry.configure(state="disabled",fg_color="gray")    

            factor_5_title_entry = create_entry_box(" LO Information Template generation ", "Enter title", "Arial", 15, 150, 1025, 305)
            factor_5_title_entry.configure(state="disabled",fg_color="gray")

            factor_1_lo_entry = create_entry_box(" LO Information Template generation ", "1,2,3,4,5,6", "Arial", 15, 150, 1225, 105)
            factor_1_lo_entry.configure(state="disabled",fg_color="gray")

            factor_2_lo_entry = create_entry_box(" LO Information Template generation ", "1,2,3,4,5,6", "Arial", 15, 150, 1225, 155)
            factor_2_lo_entry.configure(state="disabled",fg_color="gray")

            factor_3_lo_entry = create_entry_box(" LO Information Template generation ", "1,2,3,4,5,6", "Arial", 15, 150, 1225, 205)
            factor_3_lo_entry.configure(state="disabled",fg_color="gray")

            factor_4_lo_entry = create_entry_box(" LO Information Template generation ", "1,2,3,4,5,6", "Arial", 15, 150, 1225, 255)
            factor_4_lo_entry.configure(state="disabled",fg_color="gray")

            factor_5_lo_entry = create_entry_box(" LO Information Template generation ", "1,2,3,4,5,6", "Arial", 15, 150, 1225, 305)
            factor_5_lo_entry.configure(state="disabled",fg_color="gray")

            label_14_lab = create_label(" LO Information Template generation ", "Enter target levels of the tests: ", "Arial", 20, 950, 355)

            oral_marks_target_label = create_label(" LO Information Template generation ", "Oral: ", "Arial", 15, 825, 405)

            oral_marks_target_entry = create_entry_box(" LO Information Template generation ", "", "Arial", 15, 300, 1025, 405)

            mini_project_marks_target_label = create_label(" LO Information Template generation ", "Mini Project: ", "Arial", 15, 825, 505)

            mini_project_marks_target_entry = create_entry_box(" LO Information Template generation ", "", "Arial", 15, 300, 1025, 505)

            term_work_marks_target_label = create_label(" LO Information Template generation ", "Term Work: ", "Arial", 15, 825, 455)

            term_work_marks_target_entry = create_entry_box(" LO Information Template generation ", "", "Arial", 15, 300, 1025, 455)

            assignment_target_label = create_label(" LO Information Template generation ", "Assignment: ", "Arial", 15, 825, 555)

            assignment_target_entry = create_entry_box(" LO Information Template generation ", "", "Arial", 15, 300, 1025, 555)

            next_1_lab_button = create_button(" LO Information Template generation ", "Next", "Arial", 20, 250, 40, switch_1_lab, 1020, 600)

            label_15_lab = create_label(" LO Mapping ", "Enter the LO mapped for the following situations", "Arial", 20, 500, 5)

            exp_lo_label = create_label(" LO Mapping ", "Experiments ", "Arial", 15, 700, 55)

            exp_lo_labels = {}
            exp_lo_entry = {}

            for i in range(1,6):
                exp_lo_labels[f"exp{i}_lo_label"] = create_label(" LO Mapping ", f"{i}:", "Arial", 15, 200, 105+(50*(i-1)))
                exp_lo_entry[f"exp{i}_lo_entry"] = create_entry_box(" LO Mapping ", "1,2,3,4,5,6", "Arial", 15, 200, 300, 105+(50*(i-1)))

            c = 1
            for i in range(6,11):
                exp_lo_labels[f"exp{i}_lo_label"] = create_label(" LO Mapping ", f"{i}:", "Arial", 15, 600, 105+(50*(c-1)))
                exp_lo_entry[f"exp{i}_lo_entry"] = create_entry_box(" LO Mapping ", "1,2,3,4,5,6", "Arial", 15, 200, 700, 105+(50*(c-1)))
                c = c+ 1

            c = 1
            for i in range(11,16):
                exp_lo_labels[f"exp{i}_lo_label"] = create_label(" LO Mapping ", f"{i}:", "Arial", 15, 1000, 105+(50*(c-1)))
                exp_lo_entry[f"exp{i}_lo_entry"] = create_entry_box(" LO Mapping ", "1,2,3,4,5,6", "Arial", 15, 200, 1100, 105+(50*(c-1)))
                c = c + 1

            for entry in exp_lo_entry.values():
                entry.configure(state="disabled", fg_color="gray")

            mini_project_lo_lables = create_label(" LO Mapping ", "Mini project", "Arial", 15, 700, 375)

            mini_project_lo_label = {}
            mini_project_lo_entry = {}

            for i in range(1,5):
                mini_project_lo_label[f"mini_project_factor{i}_lo_label"] = create_label(" LO Mapping ", f"Factor {i}: ", "Arial", 15, 50 + (350*(i-1)), 425)
                mini_project_lo_entry[f"mini_project_factor{i}_lo_entry"] = create_entry_box(" LO Mapping ", "1,2,3,4,5,6", "Arial", 15, 175, 150 + (350*(i-1)), 425)

            assignment_lo_lables = create_label(" LO Mapping ", "Assignment", "Arial", 15, 700, 495)

            assignment_lo_label = {}
            assignment_lo_entry = {}

            for i in range(1,4):
                assignment_lo_label[f"assignment_{i}_lo_label"] = create_label(" LO Mapping ", f"Assignment {i}: ", "Arial", 15, 200 + (400*(i-1)), 545)
                assignment_lo_entry[f"assignment_{i}_lo_entry"] = create_entry_box(" LO Mapping ", "1,2,3,4,5,6", "Arial", 15, 175, 350 + (400*(i-1)), 545)

            for entry in assignment_lo_entry.values():
                entry.configure(state="disabled", fg_color="gray")

            next_2_lab_button = create_button(" LO Mapping ", "Download", "Arial", 20, 200, 40, download_template_lab, 650, 600)

            enterLO = create_label(" LO Information ", "Enter the LO's", "Arial", 20, 700, 50)
            noOfLOLabel = create_label(" LO Information ", "Select No. of LO's: ", "Arial", 15, 550, 100)
            noOfLOOption = create_dropdown(" LO Information ", ['Select No of LO\'s', '5', '6'], "Arial", 15, 300, lo_check, 750, 100)

            LO_label = {}
            LO_entry = {}

            for i in range(1,7):
                LO_label[f"LO{i}_label"] = create_label(" LO Information ", f"LO{i}: ", "Arial", 15, 550, 150 + (50*(i-1)))
                LO_entry[f"LO{i}_entry"] = create_entry_box(" LO Information ", "", "Arial", 15, 500, 650, 150 + (50*(i-1)))

            for entry in LO_entry.values():
                entry.configure(state="disabled", fg_color="gray")

            next_lab_button = create_button(" LO Information ", "Next", "Arial", 20, 200, 40, switch_lab, 725, 500)

            path_entry_lab=ctk.CTkEntry(tabview.tab(" Upload Excel File (Lab) "))

            # button_process=ctk.CTkButton(tabview.tab(" Upload Excel File "),text="Process",width=100,height=30,command=process_file)
            # button_process.place(x=500,y=500)

            upload_Label_lab = create_label(" Upload Excel File (Lab) ", "Upload you excel file with the marks entered:", "Arial", 25, 550, 50)
            path_label_lab = create_label(" Upload Excel File (Lab) ", "Path of file", "Arial", 15, 650, 110)
            button_upload_lab = create_button(" Upload Excel File (Lab) ", "Upload", "Arial", 20, 200, 40, upload_lab_file, 400, 100)


            line_lab = ctk.CTkFrame(master=tabview.tab(" Upload Excel File (Lab) "), height=2, width=1200, fg_color="white")
            line_lab.place(x=150,y=200)

            process_Label_lab = create_label(" Upload Excel File (Lab) ", "Process the excel file you uploaded:", "Arial", 25, 600, 250)

            setEmailProcessedLabel_lab = create_label(" Upload Excel File (Lab) ", "Enter the Email ID to send the calculated sheet.", "Arial", 20, 200, 325)

            # important_label = create_label(" Upload Excel File ", "Important: Please fill the no. of CO\'s field and the CO\'s in the CO Information page and AL values in Target level of tests page before processing the file", "Arial", 20, 100, 325)
            # important_label.configure(text_color="black", fg_color="yellow")

            emailTextProcessed_lab = create_entry_box(" Upload Excel File (Lab) ", "", "Arial", 15, 500, 700, 325)

            button_process_lab = create_button(" Upload Excel File (Lab) ", "Process", "Arial", 20, 200, 40, process_file_lab, 650, 425)

        

            back_button = ctk.CTkButton(lo_window, text="Back", command=lambda: self.go_back(lo_window))
            back_button.pack(pady=20)

            lo_window.mainloop()
    
    def open_main_page(self):
        
        ctk.set_appearance_mode("system")  # Modes: system (default), light, dark
        ctk.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green
        
        
        self.app = ctk.CTk()  # creating custom tkinter window
        self.app.title('CO-PO')
       
        screen_width=self.app.winfo_screenwidth()
        screen_height=self.app.winfo_screenheight()
       
        # Calculate the coordinates for centering the window
        x_position = 0
        y_position = 0
        
        # Set the window position and size
        self.app.geometry(f"{screen_width}x{screen_height}+{x_position}+{y_position}")

        
        
        self.main_frame = ctk.CTkFrame(master=self.app)
        self.main_frame.pack(expand=True, fill="both", padx=10, pady=10)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1)

        # # Half-width and full-height of main_frame for coframe
        # self.coframe = ctk.CTkFrame(master=self.main_frame)
        # self.coframe.place(relx=0, rely=0, relwidth=0.5, relheight=1)

        # # Half-width and full-height of main_frame for loframe
        # self.loframe = ctk.CTkFrame(master=self.main_frame)
        # self.loframe.place(relx=0.5, rely=0, relwidth=0.5, relheight=1)
        
        # Padding around coframe and loframe
        frame_padding = 0.01  # Adjust this to increase/decrease padding (e.g., 1% of the width/height)

        # Half-width and full-height of main_frame for coframe with padding
        self.coframe = ctk.CTkFrame(master=self.main_frame)
        self.coframe.place(
            relx=frame_padding,  # Start slightly inward (left padding)
            rely=frame_padding,  # Start slightly downward (top padding)
            relwidth=0.5 - (frame_padding),  # Reduce width for padding on both sides
            relheight=1 - (2*frame_padding)  # Reduce height for padding on top and bottom
        )

        # Half-width and full-height of main_frame for loframe with padding
        self.loframe = ctk.CTkFrame(master=self.main_frame)
        self.loframe.place(
            relx=0.5 + frame_padding,  # Start slightly after the midpoint (left padding)
            rely=frame_padding,  # Start slightly downward (top padding)
            relwidth=0.5 - (frame_padding),  # Reduce width for padding on both sides
            relheight=1 - (2*frame_padding)  # Reduce height for padding on top and bottom
        )
        
        def resource_path(relative_path):
            """Get the absolute path to a resource, handling PyInstaller paths."""
            if hasattr(sys, '_MEIPASS'):  # PyInstaller extracts files to _MEIPASS
                return os.path.join(sys._MEIPASS, relative_path)
            return os.path.join(os.path.abspath("."), relative_path)

        image_path = resource_path("Course-outcome-final-1.png")

        # Load the image and create a CTkImage
        background_image = Image.open(image_path)
        bg_image = ctk.CTkImage(background_image, size=(512,313))

        

        # Add a label to hold the background image
        bg_label = ctk.CTkLabel(master=self.coframe, image=bg_image, text="")
        bg_label.place(relx=0.5, rely=0.45, anchor="center")
        
        image_path1 = resource_path("Lab-outcome-final-1.png")

        # Load the image and create a CTkImage
        background_image1 = Image.open(image_path1)
        bg_image1 = ctk.CTkImage(background_image1, size=(512,313))

        

        # Add a label to hold the background image
        bg_label1 = ctk.CTkLabel(master=self.loframe, image=bg_image1, text="")
        bg_label1.place(relx=0.5, rely=0.45, anchor="center")

        cobutton = ctk.CTkButton(self.coframe,width=180,height=40, text="CO - PO",font=("Helvetica",25), command=self.open_co_window)
        cobutton.place(x=280,y=600) 
        
    
        lobutton = ctk.CTkButton(self.loframe,width=180,height=40, text="LO",font=("Helvetica", 25), command=self.open_lo_window)
        lobutton.place(x=280,y=600)

        
        
        # def start_button_event():
        #     if check_var.get()=='off':
        #         CTkMessagebox(title="Error", message="Please Check the box.",icon="cancel")
        #     else:
        #         import co
        #         self.app.destroy()
        #         co.User_mode()
                


        # check_var = ctk.StringVar(value="off")
        # checkbox = ctk.CTkCheckBox(self.main_frame, text="I have Read the instructions",font=("Helvetica", 20),
        #                                     variable=check_var, onvalue="on", offvalue="off")
        
        # checkbox.place(x=500,y=600)     
        
        # button1 = ctk.CTkButton(self.main_frame, text="Start",font=("Helvetica", 20), command=start_button_event)
        # button1.place(x=800,y=600)

         # add tab at the end # set currently visible tab

        


        # label12 = create_label(" LO Information Template generation ", "Attainment Target :", "Arial", 15, 875, 55)

        # entry12 = create_entry_box(" LO Information Template generation ", "52.5", "Arial", 15, 300, 1075, 55)

        self.app.mainloop()
        
def main():
    user_mode = User_mode()


if __name__ == "__main__":
    main()
