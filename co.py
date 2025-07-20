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

class User_mode:
    def __init__(self):
        self.app = None
        self.cursor = None
        self.open_main_page()

    def go_back(self, current_window):
            current_window.destroy()  # Close the current window
            self.__init__() 

    def open_co_window(self):
            
            def switch_to_co_information():
                if (entry1.get() == "" or yearDropDown.get() == "Select Year" or
                    entry8.get() == "Select Department" or entry2.get() == "Select Sem" or
                    entry3.get() == "Select Subject" or entry4.get() == "" or
                    entry5.get() == "" or entry7.get() == "Select Class" ):
                        return CTkMessagebox(title="Error", message="Please fill all the required fields.", icon="cancel")
                elif not validateNumberString(entry1.get()):
                    return CTkMessagebox(title="Invalid Input", message="Please enter valid No Of Students", icon="warning")
                else:
                    tabview.set(" CO Information ")
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
                    
            # def switch_to_MidTerm_EndSem():
            #     if noOfCOOption.get() == "5":
            #         if CO1T.get() != "" and CO2T.get() != "" and CO3T.get() != "" and CO4T.get() != "" and CO5T.get() != "":
            #             tabview.set(" Mid Terms & End Semesters ")
            #             valid_CO = [1,2,3,4,5]
            #         else:
            #             CTkMessagebox(title="Error", message="Please enter all the 5 CO's", icon="cancel")
            #     elif noOfCOOption.get() == "6":
            #         if CO1T.get() != "" and CO2T.get() != "" and CO3T.get() != "" and CO4T.get() != "" and CO5T.get() != "" and CO6T.get() != "":
            #             tabview.set(" Mid Terms & End Semesters ")
            #             valid_CO = [1,2,3,4,5,6]
            #         else:
            #             CTkMessagebox(title="Error", message="Please enter all the 6 CO's", icon="cancel")
            #     else:
            #         CTkMessagebox(title="Error", message="Please Select No Of CO's", icon="cancel")

#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
           
            def switch_to_MidTerm_EndSem():
                selected_CO = noOfCOOption.get()

                if selected_CO in ["5", "6"]:
                    total_cos = int(selected_CO)
                    all_entered = True
                    for i in range(1, total_cos + 1):
                        entry_value = co_desc_entry.get(f"CO{i}T").get()
                        if entry_value.strip() == "":
                            all_entered = False
                            break
                    
                    if all_entered:
                        tabview.set(" Mid Terms & End Semesters ")
                        valid_CO = list(range(1, total_cos + 1))
                    else:
                        CTkMessagebox(title="Error", message=f"Please enter all the {total_cos} CO's Description", icon="cancel")
                
                else:
                    CTkMessagebox(title="Error", message="Please Select No Of CO's", icon="cancel")
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   

            # def switch_to_CA1():
            #     if a1T.get()=="" or a2T.get()=="" or a3T.get()=="" or a4T.get()=="" or a5T.get()=="" or a6T.get()=="" or a2aT.get()=="" or a2bT.get()=="" or a3aT == "" or a3bT.get()=="" or entry11.get()=="":
            #         CTkMessagebox(title="Error", message="Please enter CO\'s for all questions", icon="cancel")
            #     elif not (validate_co_string(a1T.get()) and validate_co_string(a2T.get()) and validate_co_string(a3T.get()) and validate_co_string(a4T.get()) and validate_co_string(a5T.get()) and validate_co_string(a6T.get()) and validate_co_string(a2aT.get()) and validate_co_string(a2bT.get()) and validate_co_string(a3aT.get()) and validate_co_string(a3bT.get()) and validate_co_string(entry11.get())):
            #         CTkMessagebox(title="Error", message="Please enter valid format of CO string", icon="cancel")
            #     elif ALEndSemText.get() == "" or ALMidTermText.get()=="":
            #         CTkMessagebox(title="Error", message="Please enter target level of MidSem and End Semester", icon="cancel")
            #     elif (not validateNumberString(ALEndSemText.get()) or not validateNumberString(ALMidTermText.get())):
            #         CTkMessagebox(title="Error", message="Please enter valid target level of MidSem and End Semester", icon="cancel")
            #     else:
            #         tabview.set(" CA 1 ")

#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   

            def switch_to_CA1():
                # Check for empty COs
                for key, entry in midterm_co_entry.items():
                    if entry.get().strip() == "":
                        CTkMessagebox(title="Error", message=f"Please enter CO for {key}", icon="cancel")
                        return

                # Validate CO format
                for key, entry in midterm_co_entry.items():
                    if not validate_co_string(entry.get()):
                        CTkMessagebox(title="Error", message=f"Invalid CO format in {key}", icon="cancel")
                        return

                # Check target levels
                if ALEndSemText.get().strip() == "" or ALMidTermText.get().strip() == "":
                    CTkMessagebox(title="Error", message="Please enter target level of MidSem and End Semester", icon="cancel")
                    return

                if not (validateNumberString(ALEndSemText.get()) and validateNumberString(ALMidTermText.get())):
                    CTkMessagebox(title="Error", message="Please enter valid target level of MidSem and End Semester", icon="cancel")
                    return

                # All checks passed → switch tab
                tabview.set(" CA 1 ")

#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   

            # def switch_to_CA2():
            #     caQT = [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]
            #     caQM = [q1TCA1marks,q2TCA1marks,q3TCA1marks,q4TCA1marks,q5TCA1marks,q6TCA1marks,q7TCA1marks,q8TCA1marks,q9TCA1marks,q10TCA1marks]
            #     if noCA1Entry.get()!="Select No":
            #         option = int(noCA1Entry.get())
            #         for i in range(option,10):
            #             print(f"Ye option hai {option} {i}")
            #             caQT.pop(option)
            #             caQM.pop(option)
            #         for i in caQT:
            #             print(f"Ye CA co hai {i.get()}")
            #         for i in caQM:
            #             print(f"Ye CA marks hai {i.get()}")
                    
            #     if entry13.get()=="Select Type":
            #         return CTkMessagebox(title="Error", message="Please select Type of CA", icon="cancel")
            #     if (entry13.get()=="Quiz" or entry13.get()=="Test") :
            #         if noCA1Entry.get()=="Select No":
            #             return CTkMessagebox(title="Error", message="Please select No of questions", icon="cancel")
            #         else:
            #             for ca1 in caQT:
            #                 print(ca1.get())
            #                 if ca1.get()=="":
            #                     return CTkMessagebox(title="Error", message="Please enter CO of CA", icon="cancel")
            #                 elif not validate_co_string(ca1.get()):
            #                     return CTkMessagebox(title="Error", message="Please enter valid CO String", icon="cancel")
            #             for ca1M in caQM:
            #                 if ca1M.get()=="":
            #                     return CTkMessagebox(title="Error", message="Please enter marks of the questions", icon="cancel")
            #                 if( not validateNumberString(ca1M.get())):
            #                     return CTkMessagebox(title="Error", message="Please enter valid marks for the questions", icon="cancel")
            #     if entry13.get()=="NPTEL Course" and (nptelCA1Text.get()==""):
            #         return CTkMessagebox(title="Error", message="Please enter CO of NPTEL", icon="cancel")
            #     elif entry13.get()=="NPTEL Course" and not validate_co_string(nptelCA1Text.get()):
            #         return CTkMessagebox(title="Error", message="Please enter valid CO format", icon="cancel")
            #     if entry13.get()=="Presentation" and presentationCA1Text.get()=="":
            #         return CTkMessagebox(title="Error", message="Please enter maximum students in Presentation", icon="cancel")
            #     if (not validateNumberString(presentationCA1Text.get())):
            #         return CTkMessagebox(title="Error", message="Please enter valid number of students in Presentation", icon="cancel")
            #     if ALCA1Text.get() == "":
            #         return CTkMessagebox(title="Error", message="Please enter target level of CA1", icon="cancel")
            #     if (not validateNumberString(ALCA1Text.get())):
            #         return CTkMessagebox(title="Error", message="Please enter valid target level of CA1", icon="cancel")
            #     tabview.set(" CA 2 ")
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
            def switch_to_CA2():
                caQT = [Question_CA1_CO[f"Q{i}_questions"] for i in range(1, 11)]
                caQM = [Marks_CA1_CO[f"Q{i}_marks"] for i in range(1, 11)]

                if noCA1Entry.get() != "Select No":
                    option = int(noCA1Entry.get())
                    for i in range(option, 10):
                        print(f"Ye option hai {option} {i}")
                        caQT[i].configure(state="disabled", fg_color="gray")
                        caQM[i].configure(state="disabled", fg_color="gray")
                    caQT = caQT[:option]
                    caQM = caQM[:option]

                for i in caQT:
                    print(f"Ye CA co hai {i.get()}")
                for i in caQM:
                    print(f"Ye CA marks hai {i.get()}")

                if entry13.get() == "Select Type":
                    return CTkMessagebox(title="Error", message="Please select Type of CA", icon="cancel")

                if entry13.get() in ["Quiz", "Test"]:
                    if noCA1Entry.get() == "Select No":
                        return CTkMessagebox(title="Error", message="Please select No of questions", icon="cancel")
                    else:
                        for ca1 in caQT:
                            if ca1.get() == "":
                                return CTkMessagebox(title="Error", message="Please enter CO of CA", icon="cancel")
                            elif not validate_co_string(ca1.get()):
                                return CTkMessagebox(title="Error", message="Please enter valid CO String", icon="cancel")
                        for ca1M in caQM:
                            if ca1M.get() == "":
                                return CTkMessagebox(title="Error", message="Please enter marks of the questions", icon="cancel")
                            elif not validateNumberString(ca1M.get()):
                                return CTkMessagebox(title="Error", message="Please enter valid marks for the questions", icon="cancel")

                if entry13.get() == "NPTEL Course":
                    if nptelCA1Text.get() == "":
                        return CTkMessagebox(title="Error", message="Please enter CO of NPTEL", icon="cancel")
                    elif not validate_co_string(nptelCA1Text.get()):
                        return CTkMessagebox(title="Error", message="Please enter valid CO format", icon="cancel")

                if entry13.get() == "Presentation":
                    if presentationCA1Text.get() == "":
                        return CTkMessagebox(title="Error", message="Please enter maximum students in Presentation", icon="cancel")
                    elif not validateNumberString(presentationCA1Text.get()):
                        return CTkMessagebox(title="Error", message="Please enter valid number of students in Presentation", icon="cancel")

                if ALCA1Text.get() == "":
                    return CTkMessagebox(title="Error", message="Please enter target level of CA1", icon="cancel")
                elif not validateNumberString(ALCA1Text.get()):
                    return CTkMessagebox(title="Error", message="Please enter valid target level of CA1", icon="cancel")

                tabview.set(" CA 2 ")
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
             
            # def switch_to_CA3():
            #     caQT = [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]
            #     caQM = [q1TCA2marks,q2TCA2marks,q3TCA2marks,q4TCA2marks,q5TCA2marks,q6TCA2marks,q7TCA2marks,q8TCA2marks,q9TCA2marks,q10TCA2marks]
            #     if noCA2Entry.get()!="Select No":
            #         option = int(noCA2Entry.get())
            #         for i in range(option,10):
            #             caQT.pop(option)
            #             caQM.pop(option)
                
                    
            #     if entry14.get()=="Select Type":
            #         return CTkMessagebox(title="Error", message="Please select Type of CA", icon="cancel")
            #     if (entry14.get()=="Quiz" or entry14.get()=="Test"):
            #         if noCA2Entry.get()=="Select No":
            #             return CTkMessagebox(title="Error", message="Please select No of questions", icon="cancel")
            #         else:
            #             for ca2 in caQT:
            #                 if ca2.get()=="":
            #                     return CTkMessagebox(title="Error", message="Please enter CO of CA", icon="cancel")
            #                 elif not validate_co_string(ca2.get()):
            #                     return CTkMessagebox(title="Error", message="Please enter valid CO String", icon="cancel")
            #             for ca2M in caQM:
            #                 if ca2M.get()=="":
            #                     return CTkMessagebox(title="Error", message="Please enter marks of the questions", icon="cancel")
            #                 if not (validateNumberString(ca2M.get())):
            #                     return CTkMessagebox(title="Error", message="Please enter valid marks for the questions", icon="cancel")
            #     if entry14.get()=="NPTEL Course" and (nptelCA2Text.get()==""):
            #         return CTkMessagebox(title="Error", message="Please enter CO of NPTEL", icon="cancel")
            #     elif entry14.get()=="NPTEL Course" and not validate_co_string(nptelCA2Text.get()):
            #         return CTkMessagebox(title="Error", message="Please enter valid CO format", icon="cancel")
            #     if entry14.get()=="Presentation" and presentationCA2Text.get()=="":
            #         return CTkMessagebox(title="Error", message="Please enter maximum students in Presentation", icon="cancel")
            #     if not (validateNumberString(presentationCA2Text.get())):
            #         return CTkMessagebox(title="Error", message="Please enter valid number of students in Presentation", icon="cancel")
            #     if ALCA2Text.get() == "":
            #         return CTkMessagebox(title="Error", message="Please enter target level of CA1", icon="cancel")
            #     if not (validateNumberString(ALCA2Text.get())):
            #         return CTkMessagebox(title="Error", message="Please enter valid target level of CA1", icon="cancel")
            #     tabview.set(" CA 3 ")
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
            def switch_to_CA3():
                caQT = [Question_CA2_CO[f"Q{i}_questions"] for i in range(1, 11)]
                caQM = [Marks_CA2_CO[f"Q{i}_marks"] for i in range(1, 11)]

                if noCA2Entry.get() != "Select No":
                    option = int(noCA2Entry.get())
                    for i in range(option, 10):
                        caQT[i].configure(state="disabled", fg_color="gray")
                        caQM[i].configure(state="disabled", fg_color="gray")
                    caQT = caQT[:option]
                    caQM = caQM[:option]

                if entry14.get() == "Select Type":
                    return CTkMessagebox(title="Error", message="Please select Type of CA", icon="cancel")

                if entry14.get() in ["Quiz", "Test"]:
                    if noCA2Entry.get() == "Select No":
                        return CTkMessagebox(title="Error", message="Please select No of questions", icon="cancel")
                    else:
                        for ca2 in caQT:
                            if ca2.get() == "":
                                return CTkMessagebox(title="Error", message="Please enter CO of CA", icon="cancel")
                            elif not validate_co_string(ca2.get()):
                                return CTkMessagebox(title="Error", message="Please enter valid CO String", icon="cancel")
                        for ca2M in caQM:
                            if ca2M.get() == "":
                                return CTkMessagebox(title="Error", message="Please enter marks of the questions", icon="cancel")
                            elif not validateNumberString(ca2M.get()):
                                return CTkMessagebox(title="Error", message="Please enter valid marks for the questions", icon="cancel")

                if entry14.get() == "NPTEL Course":
                    if nptelCA2Text.get() == "":
                        return CTkMessagebox(title="Error", message="Please enter CO of NPTEL", icon="cancel")
                    elif not validate_co_string(nptelCA2Text.get()):
                        return CTkMessagebox(title="Error", message="Please enter valid CO format", icon="cancel")

                if entry14.get() == "Presentation":
                    if presentationCA2Text.get() == "":
                        return CTkMessagebox(title="Error", message="Please enter maximum students in Presentation", icon="cancel")
                    elif not validateNumberString(presentationCA2Text.get()):
                        return CTkMessagebox(title="Error", message="Please enter valid number of students in Presentation", icon="cancel")

                if ALCA2Text.get() == "":
                    return CTkMessagebox(title="Error", message="Please enter target level of CA2", icon="cancel")
                elif not validateNumberString(ALCA2Text.get()):
                    return CTkMessagebox(title="Error", message="Please enter valid target level of CA2", icon="cancel")

                tabview.set(" CA 3 ")

#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   

            # def switch_to_template():
            #     if entry10.get() == "Yes":
            #         caQT = [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]
            #         caQM = [q1TCA3marks,q2TCA3marks,q3TCA3marks,q4TCA3marks,q5TCA3marks,q6TCA3marks,q7TCA3marks,q8TCA3marks,q9TCA3marks,q10TCA3marks]
            #         if noCA3Entry.get()!="Select No":
            #             option = int(noCA3Entry.get())
            #             for i in range(option,10):
            #                 caQT.pop(option)
            #                 caQM.pop(option)
            #         if entry15.get()=="Select Type":
            #             return CTkMessagebox(title="Error", message="Please select Type of CA", icon="cancel")
            #         if (entry15.get()=="Quiz" or entry15.get()=="Test"):
            #             if noCA3Entry.get()=="Select No":
            #                 return CTkMessagebox(title="Error", message="Please select No of questions", icon="cancel")
            #             else:
            #                 for ca3 in caQT:
            #                     if ca3.get()=="":
            #                         return CTkMessagebox(title="Error", message="Please enter CO of CA", icon="cancel")
            #                     elif not validate_co_string(ca3.get()):
            #                         return CTkMessagebox(title="Error", message="Please enter valid CO String", icon="cancel")
            #                 for ca3M in caQM:
            #                     if ca3M.get()=="":
            #                         return CTkMessagebox(title="Error", message="Please enter marks of the questions", icon="cancel")
            #                     if not (validateNumberString(ca3M.get())):
            #                         return CTkMessagebox(title="Error", message="Please enter valid marks for the questions", icon="cancel")
            #         if entry15.get()=="NPTEL Course" and (nptelCA3Text.get()==""):
            #             return CTkMessagebox(title="Error", message="Please enter CO of NPTEL", icon="cancel")
            #         elif entry15.get()=="NPTEL Course" and not validate_co_string(nptelCA3Text.get()):
            #             return CTkMessagebox(title="Error", message="Please enter valid CO format", icon="cancel")
            #         if entry15.get()=="Presentation" and presentationCA3Text.get()=="":
            #             return CTkMessagebox(title="Error", message="Please enter maximum students in Presentation", icon="cancel")
            #         if not (validateNumberString(presentationCA3Text.get())):
            #             return CTkMessagebox(title="Error", message="Please enter valid number of students in Presentation", icon="cancel")
            #         if ALCA3Text.get() == "":
            #             return CTkMessagebox(title="Error", message="Please enter target level of CA1", icon="cancel")
            #         if not (validateNumberString(ALCA3Text.get())):
            #             return CTkMessagebox(title="Error", message="Please enter valid target level of CA1", icon="cancel")
            #     tabview.set(" Process Template/Calculated ")
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
            def switch_to_template():
                if entry10.get() == "Yes":
                    caQT = [Question_CA3_CO[f"Q{i}_questions"] for i in range(1, 11)]
                    caQM = [Marks_CA3_CO[f"Q{i}_marks"] for i in range(1, 11)]

                    if noCA3Entry.get() != "Select No":
                        option = int(noCA3Entry.get())
                        for i in range(option, 10):
                            caQT[i].configure(state="disabled", fg_color="gray")
                            caQM[i].configure(state="disabled", fg_color="gray")
                        caQT = caQT[:option]
                        caQM = caQM[:option]

                    if entry15.get() == "Select Type":
                        return CTkMessagebox(title="Error", message="Please select Type of CA", icon="cancel")

                    if entry15.get() in ["Quiz", "Test"]:
                        if noCA3Entry.get() == "Select No":
                            return CTkMessagebox(title="Error", message="Please select No of questions", icon="cancel")
                        else:
                            for ca3 in caQT:
                                if ca3.get() == "":
                                    return CTkMessagebox(title="Error", message="Please enter CO of CA", icon="cancel")
                                elif not validate_co_string(ca3.get()):
                                    return CTkMessagebox(title="Error", message="Please enter valid CO String", icon="cancel")
                            for ca3M in caQM:
                                if ca3M.get() == "":
                                    return CTkMessagebox(title="Error", message="Please enter marks of the questions", icon="cancel")
                                elif not validateNumberString(ca3M.get()):
                                    return CTkMessagebox(title="Error", message="Please enter valid marks for the questions", icon="cancel")

                    if entry15.get() == "NPTEL Course":
                        if nptelCA3Text.get() == "":
                            return CTkMessagebox(title="Error", message="Please enter CO of NPTEL", icon="cancel")
                        elif not validate_co_string(nptelCA3Text.get()):
                            return CTkMessagebox(title="Error", message="Please enter valid CO format", icon="cancel")

                    if entry15.get() == "Presentation":
                        if presentationCA3Text.get() == "":
                            return CTkMessagebox(title="Error", message="Please enter maximum students in Presentation", icon="cancel")
                        elif not validateNumberString(presentationCA3Text.get()):
                            return CTkMessagebox(title="Error", message="Please enter valid number of students in Presentation", icon="cancel")

                    if ALCA3Text.get() == "":
                        return CTkMessagebox(title="Error", message="Please enter target level of CA3", icon="cancel")
                    elif not validateNumberString(ALCA3Text.get()):
                        return CTkMessagebox(title="Error", message="Please enter valid target level of CA3", icon="cancel")

                tabview.set(" Process Template/Calculated ")


#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
            # def create_button(tab, name, font_name, font_size, w, h, com, x, y):
            #     button = ctk.CTkButton(master=tabview.tab(tab), text=name, width=w, height=h, font=(font_name, font_size), command=com)
            #     button.place(x=x, y=y)
            #     return button
            
            # def create_label(tab, name, font_type, font_size, x, y):
            #     label = ctk.CTkLabel(master=tabview.tab(tab), text=name, font=(font_type, font_size))
            #     label.place(x=x, y=y)
            #     return label
    
            # def create_entry_box(tab, text, font_name, font_size, w, x, y):
            #     entry_box = ctk.CTkEntry(master=tabview.tab(tab), placeholder_text=text, font=(font_name,font_size), width=w)
            #     entry_box.place(x=x,y=y)
            #     return entry_box
            
            # def create_dropdown(tab, val, font_name, font_size, w, com, x, y):
            #     dropdown = ctk.CTkOptionMenu(master=tabview.tab(tab), values=val, font=(font_name, font_size), width=w, command=com)
            #     dropdown.place(x=x, y=y)
            #     return dropdown
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
          

            # ------------------ Helper UI Functions (Responsive & Scrollable) ------------------

            def create_label(tab, name, font_type, font_size, row=None, column=None, colspan=1, sticky="nsw", padx=5, pady=15):
                label = ctk.CTkLabel(master=scroll_frames[tab], text=name, font=(font_type, font_size))
                if row is not None and column is not None:
                    label.grid(row=row, column=column, columnspan=colspan, sticky=sticky, padx=padx, pady=pady)
                else:
                    label.pack(pady=15)
                return label

            def create_entry_box(tab, text, font_name, font_size, w, row=None, column=None, colspan=1, sticky="nsw", padx=5, pady=15):
                entry_box = ctk.CTkEntry(master=scroll_frames[tab], placeholder_text=text, font=(font_name, font_size), width=w)
                if row is not None and column is not None:
                    entry_box.grid(row=row, column=column, columnspan=colspan, sticky=sticky, padx=padx, pady=pady)
                else:
                    entry_box.pack(pady=15)
                return entry_box

            def create_button(tab, name, font_name, font_size, w, h=40, com=None, row=None, column=None, colspan=1, sticky="nsw", padx=5, pady=20):
                button = ctk.CTkButton(master=scroll_frames[tab], text=name, width=w, height=h, font=(font_name, font_size), command=com)
                if row is not None and column is not None:
                    button.grid(row=row, column=column, columnspan=colspan, sticky=sticky, padx=padx, pady=pady)
                else:
                    button.pack(pady=20)
                return button

            def create_dropdown(tab, val, font_name, font_size, w, com=None, row=None, column=None, colspan=1, sticky="nsw", padx=5, pady=15):
                dropdown = ctk.CTkOptionMenu(master=scroll_frames[tab], values=val, font=(font_name, font_size), width=w, command=com)
                if row is not None and column is not None:
                    dropdown.grid(row=row, column=column, columnspan=colspan, sticky=sticky, padx=padx, pady=pady)
                else:
                    dropdown.pack(pady=15)
                return dropdown
            
            def download():
                CA1M_Co_arr = []
                CA2M_Co_arr = []
                CA3M_Co_arr = []
                # Get the values from the Entry widgets
                # values = [entry1.get(), entry2.get(), entry3.get(), entry4.get(), entry5.get(),
                #         a1T.get(), a2T.get(), a3T.get(), a4T.get(), a5T.get(), a6T.get(),
                #         a2aT.get(), a2bT.get(), a3aT.get(), a3bT.get(),entry7.get(),entry8.get(),entry10.get(),entry11.get(),entry12.get()]
                
                # coValues =[entry13.get(),entry14.get(),entry15.get(),q1T.get(),q6T.get(),q2T.get(),q2aT.get(),q3T.get(),q2bT.get(),q4T.get(),q3aT.get(),q5T.get(),q3bT.get()]
                # coCAs=[entry13.get(),entry14.get(),entry15.get()]
                # coQuizs=[q1T.get(),q6T.get(),q2T.get(),q2aT.get(),q3T.get(),q2bT.get(),q4T.get(),q3aT.get(),q5T.get(),q3bT.get()]
    #<---------------------------------------------------------------------------------------------------------------->
                # if noOfCOOption.get()=='6':
                #     coTextArray = [CO1T.get(),CO2T.get(),CO3T.get(),CO4T.get(),CO5T.get(),CO6T.get()]
                # elif noOfCOOption.get()=='5':
                #     coTextArray = [CO1T.get(),CO2T.get(),CO3T.get(),CO4T.get(),CO5T.get(),'-']
                # else:
                #     CTkMessagebox(title="Error", message="Please Select No Of CO's", icon="cancel")
    #<---------------------------------------------------------------------------------------------------------------->

                if noOfCOOption.get() == '6':
                    coTextArray = [co_desc_entry[f"CO{i}T"].get() for i in range(1, 7)]
                elif noOfCOOption.get() == '5':
                    coTextArray = [co_desc_entry[f"CO{i}T"].get() for i in range(1, 6)] + ['-']
                else:
                    CTkMessagebox(title="Error", message="Please Select No Of CO's", icon="cancel")

                print(coTextArray)
    #<---------------------------------------------------------------------------------------------------------------->
            
                # values=[entry1.get(),entry8.get(),yearDropDown.get(),entry2.get(),entry3.get(),entry4.get(),entry5.get(),entry7.get(),entry11.get(),
                
    
                #         entry10.get(),entry13.get(),entry14.get(),entry15.get(),
    
                #         noCA1Entry.get(),noCA2Entry.get(),noCA3Entry.get(),
                #         nptelCA1Text.get(),nptelCA2Text.get(),nptelCA3Text.get(),
    
                #         q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get(),q9TCA1.get(),q10TCA1.get(),
                #         q1TCA2.get(),q2TCA2.get(),q3TCA2.get(),q4TCA2.get(),q5TCA2.get(),q6TCA2.get(),q7TCA2.get(),q8TCA2.get(),q9TCA2.get(),q10TCA2.get(),
                #         q1TCA3.get(),q2TCA3.get(),q3TCA3.get(),q4TCA3.get(),q5TCA3.get(),q6TCA3.get(),q7TCA3.get(),q8TCA3.get(),q9TCA3.get(),q10TCA3.get(),
    
                #         a1T.get(), a2T.get(), a3T.get(), a4T.get(), a5T.get(), a6T.get(),a2aT.get(),a2bT.get(), a3aT.get(), a3bT.get()]
    #<---------------------------------------------------------------------------------------------------------------->
           
                values = [
                    entry1.get(), entry8.get(), yearDropDown.get(), entry2.get(), entry3.get(),
                    entry4.get(), entry5.get(), entry7.get(), entry11.get(),

                    entry10.get(), entry13.get(), entry14.get(), entry15.get(),

                    noCA1Entry.get(), noCA2Entry.get(), noCA3Entry.get(),
                    nptelCA1Text.get(), nptelCA2Text.get(), nptelCA3Text.get(),

                    *[Question_CA1_CO[f"Q{i}_questions"].get() for i in range(1, 11)],
                    *[Question_CA2_CO[f"Q{i}_questions"].get() for i in range(1, 11)],
                    *[Question_CA3_CO[f"Q{i}_questions"].get() for i in range(1, 11)],

                    *[midterm_co_entry[f"Q.{i}"].get() for i in range(1, 11)]

                ]

                basic_values=[entry1.get(),entry8.get(),yearDropDown.get(),entry2.get(),entry3.get(),entry4.get(),entry5.get(),entry7.get(),entry11.get(),None,entry10.get(),entry13.get(),entry14.get(),entry15.get(),noOfCOOption.get()]
                midSem_Co_values=[midterm_co_entry[f"Q.{i}"].get() for i in range(1, 11)]
                receiversEmail = emailText.get()
                
                if entry10.get()=="No":
                    al_values=[ALCA1Text.get(), ALCA2Text.get(), '-', ALMidTermText.get(), ALEndSemText.get()]
                    if entry13.get()=="Select Type":
                        CTkMessagebox(title="Error", message="Please Select Type of CA 1.", icon="cancel")
                    elif entry13.get()=="Quiz" or entry13.get() == "Test":
                        if noCA1Entry.get()=="Select No" :
                            CTkMessagebox(title="Error", message="Please Select No Of Question in CA 1.", icon="cancel")
                        else:
                            # caQT = [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]
                            # caQM = [q1TCA1marks,q2TCA1marks,q3TCA1marks,q4TCA1marks,q5TCA1marks,q6TCA1marks,q7TCA1marks,q8TCA1marks,q9TCA1marks,q10TCA1marks]
                            caQT = [Question_CA1_CO[f"Q{i}_questions"] for i in range(1, 11)]
                            caQM = [Marks_CA1_CO[f"Q{i}_marks"] for i in range(1, 11)]

                            option = int(noCA1Entry.get())
                            CA1_Co_arr = []
                            CA1M_Co_arr = []
                            for i in range (0, option):
                                CA1_Co_arr.append(caQT[i].get())
                                CA1M_Co_arr.append(caQM[i].get())
                            
                        # elif noCA1Entry.get()=="1" :
                        #     CA1_Co_arr=[q1TCA1.get()]
                        # elif noCA1Entry.get()=="2" :
                        #     CA1_Co_arr=[q1TCA1.get(),q2TCA1.get()]
                        # elif noCA1Entry.get()=="3" :
                        #     CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get()]
                        # elif noCA1Entry.get()=="4" :
                        #     CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get()]
                        # elif noCA1Entry.get()=="5" :
                        #     CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get()]
                        # elif noCA1Entry.get()=="6" :
                        #     CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get()]
                        # elif noCA1Entry.get()=="7" :
                        #     CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get()]
                        # elif noCA1Entry.get()=="8" :
                        #     CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get()]
                        # elif noCA1Entry.get()=="9" :
                        #     CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get(),q9TCA1.get()]
                        # else:
                        #     CA1_Co_arr=[q1TCA1.get(),q2TCA1.get(),q3TCA1.get(),q4TCA1.get(),q5TCA1.get(),q6TCA1.get(),q7TCA1.get(),q8TCA1.get(),q9TCA1.get(),q10TCA1.get()]       
                    
                    elif entry13.get()=="Presentation":
                        CA1_Co_arr = [presentationCA1Text.get()]
                    elif entry13.get()=="NPTEL Course":
                        CA1_Co_arr=[nptelCA1Text.get()]
                    else:
                        CA1_Co_arr=[1,2,3,4,5,6]

                            
                    if entry14.get()=="Select Type":
                        CTkMessagebox(title="Error", message="Please Select Type of CA 2.", icon="cancel")
                    elif entry14.get()=="Quiz" or entry14.get() == "Test":
                        if noCA2Entry.get()=="Select No" :
                            CTkMessagebox(title="Error", message="Please Select No Of Question in CA 1.", icon="cancel")
                        else:
                            # caQT = [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]
                            # caQM = [q1TCA2marks,q2TCA2marks,q3TCA2marks,q4TCA2marks,q5TCA2marks,q6TCA2marks,q7TCA2marks,q8TCA2marks,q9TCA2marks,q10TCA2marks]
                            caQT = [Question_CA2_CO[f"Q{i}_questions"] for i in range(1, 11)]
                            caQM = [Marks_CA2_CO[f"Q{i}_marks"] for i in range(1, 11)]

                            option = int(noCA2Entry.get())
                            CA2_Co_arr = []
                            CA2M_Co_arr = []
                            for i in range (0, option):
                                CA2_Co_arr.append(caQT[i].get())
                                CA2M_Co_arr.append(caQM[i].get())
                    
                    elif entry14.get()=="Presentation":
                        CA2_Co_arr = [presentationCA2Text.get()]
                    elif entry14.get()=="NPTEL Course":
                        CA2_Co_arr=[nptelCA2Text.get()]
                    else:
                        CA2_Co_arr=[1,2,3,4,5,6]
                        
                else:
                    al_values=[ALCA1Text.get(), ALCA2Text.get(), ALCA3Text.get(), ALMidTermText.get(), ALEndSemText.get()]
                    if entry13.get()=="Select Type":
                        CTkMessagebox(title="Error", message="Please Select Type of CA 1.", icon="cancel")
                    elif entry13.get()=="Quiz" or entry13.get()=="Test":
                        if noCA1Entry.get()=="Select No" :
                            CTkMessagebox(title="Error", message="Please Select No Of Question in CA 1.", icon="cancel")
                        else:
                            # caQT = [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]
                            # caQM = [q1TCA1marks,q2TCA1marks,q3TCA1marks,q4TCA1marks,q5TCA1marks,q6TCA1marks,q7TCA1marks,q8TCA1marks,q9TCA1marks,q10TCA1marks]
                            caQT = [Question_CA1_CO[f"Q{i}_questions"] for i in range(1, 11)]
                            caQM = [Marks_CA1_CO[f"Q{i}_marks"] for i in range(1, 11)]

                            option = int(noCA1Entry.get())
                            CA1_Co_arr = []
                            CA1M_Co_arr = []
                            for i in range (0, option):
                                CA1_Co_arr.append(caQT[i].get())
                                CA1M_Co_arr.append(caQM[i].get())
                    
                    elif entry13.get()=="Presentation":
                        CA1_Co_arr = [presentationCA1Text.get()]
                    elif entry13.get()=="NPTEL Course":
                        CA1_Co_arr=[nptelCA1Text.get()]
                    else:
                        CA1_Co_arr=[1,2,3,4,5,6]
                    
                        
                    if entry14.get()=="Select Type":
                        CTkMessagebox(title="Error", message="Please Select Type of CA 2.", icon="cancel")
                    elif entry14.get()=="Quiz" or entry14.get() == "Test":
                        if noCA2Entry.get()=="Select No" :
                            CTkMessagebox(title="Error", message="Please Select No Of Question in CA 1.", icon="cancel")
                        else:
                            # caQT = [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]
                            # caQM = [q1TCA2marks,q2TCA2marks,q3TCA2marks,q4TCA2marks,q5TCA2marks,q6TCA2marks,q7TCA2marks,q8TCA2marks,q9TCA2marks,q10TCA2marks]
                            caQT = [Question_CA2_CO[f"Q{i}_questions"] for i in range(1, 11)]
                            caQM = [Marks_CA2_CO[f"Q{i}_marks"] for i in range(1, 11)]

                            option = int(noCA2Entry.get())
                            CA2_Co_arr = []
                            CA2M_Co_arr = []
                            for i in range (0, option):
                                CA2_Co_arr.append(caQT[i].get())
                                CA2M_Co_arr.append(caQM[i].get())
                    
                    elif entry14.get()=="Presentation":
                        CA2_Co_arr = [presentationCA2Text.get()]
                    elif entry14.get()=="NPTEL Course":
                        CA2_Co_arr=[nptelCA2Text.get()]
                    else:
                        CA2_Co_arr=[1,2,3,4,5,6]
                        
                    if entry15.get()=="Select Type":
                        CTkMessagebox(title="Error", message="Please Select Type of CA 3.", icon="cancel")
                    elif entry15.get()=="Quiz" or entry15.get() == "Test":
                        if noCA3Entry.get()=="Select No" :
                            CTkMessagebox(title="Error", message="Please Select No Of Question in CA 3.", icon="cancel")
                        else:
                            # caQT = [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]
                            # caQM = [q1TCA3marks,q2TCA3marks,q3TCA3marks,q4TCA3marks,q5TCA3marks,q6TCA3marks,q7TCA3marks,q8TCA3marks,q9TCA3marks,q10TCA3marks]
                            caQT = [Question_CA3_CO[f"Q{i}_questions"] for i in range(1, 11)]
                            caQM = [Marks_CA3_CO[f"Q{i}_marks"] for i in range(1, 11)]

                            option = int(noCA3Entry.get())
                            CA3_Co_arr = []
                            CA3M_Co_arr = []
                            for i in range (0, option):
                                print(caQT[i].get())
                                CA3_Co_arr.append(caQT[i].get())
                                CA3M_Co_arr.append(caQM[i].get())
                    
                    elif entry15.get()=="Presentation":
                        CA3_Co_arr = [presentationCA3Text.get()]
                    elif entry15.get()=="NPTEL Course":
                        CA3_Co_arr=[nptelCA3Text.get()]
                    else: 
                        CA3_Co_arr=[1,2,3,4,5,6]
                        
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
                elif entry10.get()=="Yes":
                    print(f"Ye rha CA3 {CA3_Co_arr}")
                    if any(ca3 == "" for ca3 in CA3_Co_arr):
                        CTkMessagebox(title="Error", message="Please fill in all required fields.", icon="cancel")
                    else :
                        import template_generator
                        template_generator.template_gen(coTextArray,basic_values,midSem_Co_values,CA1_Co_arr,CA2_Co_arr,CA3_Co_arr,CA1M_Co_arr,CA2M_Co_arr,CA3M_Co_arr,al_values, receiversEmail)
                        # template_generator.template_gen(basic_values,midSem_Co_values,CA1_Co_arr,CA2_Co_arr,CA3_Co_arr)
                        # CTkMessagebox(message="Excel template downloaded successfully.",icon="check", option_1="OK")
    
                        
                else:
                    # import template_generator
                    # print(entry10.get())
                    if entry10.get()=="No": 
                        import template_generator
                        # print("Hi v1",basic_values[10]) 
                        template_generator.template_gen(coTextArray,basic_values,midSem_Co_values,CA1_Co_arr,CA2_Co_arr,[],CA1M_Co_arr,CA2M_Co_arr,[],al_values,receiversEmail)
                        # CTkMessagebox(message="Excel template downloaded successfully.",icon="check", option_1="OK")
                    elif entry10.get()=="Yes":
                        import template_generator
                        # print("Hi v2",basic_values[10]) 
                        # print("Hi v2",CA3_Co_arr) 
                        template_generator.template_gen(coTextArray,basic_values,midSem_Co_values,CA1_Co_arr,CA2_Co_arr,CA3_Co_arr,CA1M_Co_arr,CA2M_Co_arr,CA3M_Co_arr,al_values,receiversEmail)
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
                                            
                    
        #<----------------------------------------------------------------------------------------------------------------------------------------------------------------------->        
            # def ca1(option):
            #     if  option == "Select Type":
            #         noCA1Entry.configure(state="disabled", fg_color="gray",button_color="gray") 
            #         for disca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1,nptelCA1Text,presentationCA1Text, q1TCA1marks,q2TCA1marks,q3TCA1marks,q4TCA1marks,q5TCA1marks,q6TCA1marks,q7TCA1marks,q8TCA1marks,q9TCA1marks,q10TCA1marks]:
            #             disca.configure(state="disabled", fg_color="gray")
              
            #     elif option == "NPTEL Course":
            #         noCA1Entry.configure(state="disabled", fg_color="gray",button_color="gray") 
            #         for disca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1,presentationCA1Text, q1TCA1marks,q2TCA1marks,q3TCA1marks,q4TCA1marks,q5TCA1marks,q6TCA1marks,q7TCA1marks,q8TCA1marks,q9TCA1marks,q10TCA1marks]:
            #             disca.configure(state="disabled", fg_color="gray")
            #         nptelCA1Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"]) 
                
            #     elif option == "Presentation":
            #         noCA1Entry.configure(state="disabled", fg_color="gray",button_color="gray") 
            #         for preca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1,nptelCA1Text,q1TCA1marks,q2TCA1marks,q3TCA1marks,q4TCA1marks,q5TCA1marks,q6TCA1marks,q7TCA1marks,q8TCA1marks,q9TCA1marks,q10TCA1marks]:
            #             preca.configure(state="disabled", fg_color="gray")
            #         presentationCA1Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    
            #     elif option == "Quiz" or option == "Test":
            #         # for ca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1, q1TCA1marks,q2TCA1marks,q3TCA1marks,q4TCA1marks,q5TCA1marks,q6TCA1marks,q7TCA1marks,q8TCA1marks,q9TCA1marks,q10TCA1marks]:
            #         #     ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
            #         noCA1Entry.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"],button_color=["F9F9FA", "#144870"])
            #         nptelCA1Text.configure(state="disabled", fg_color="gray") 
            #         presentationCA1Text.configure(state="disabled", fg_color="gray") 
                    
            #     elif option == "Other":
            #         noCA1Entry.configure(state="disabled", fg_color="gray",button_color="gray") 
            #         for disca in [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1,presentationCA1Text, nptelCA1Text, q1TCA1marks,q2TCA1marks,q3TCA1marks,q4TCA1marks,q5TCA1marks,q6TCA1marks,q7TCA1marks,q8TCA1marks,q9TCA1marks,q10TCA1marks]:
            #             disca.configure(state="disabled", fg_color="gray")
            
        #<----------------------------------------------------------------------------------------------------------------------------------------------------------------------->        
          
            def ca1(option):
                # Prepare lists from the dictionaries
                question_entries = list(Question_CA1_CO.values())
                marks_entries = list(Marks_CA1_CO.values())

                if option == "Select Type":
                    noCA1Entry.configure(state="disabled", fg_color="gray", button_color="gray")
                    for widget in question_entries + marks_entries + [nptelCA1Text, presentationCA1Text]:
                        widget.configure(state="disabled", fg_color="gray")

                elif option == "NPTEL Course":
                    noCA1Entry.configure(state="disabled", fg_color="gray", button_color="gray")
                    for widget in question_entries + marks_entries + [presentationCA1Text]:
                        widget.configure(state="disabled", fg_color="gray")
                    nptelCA1Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])

                elif option == "Presentation":
                    noCA1Entry.configure(state="disabled", fg_color="gray", button_color="gray")
                    for widget in question_entries + marks_entries + [nptelCA1Text]:
                        widget.configure(state="disabled", fg_color="gray")
                    presentationCA1Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])

                elif option in ["Quiz", "Test"]:
                    noCA1Entry.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"], button_color=["F9F9FA", "#144870"])
                    nptelCA1Text.configure(state="disabled", fg_color="gray")
                    presentationCA1Text.configure(state="disabled", fg_color="gray")
                    # for widget in question_entries + marks_entries:
                    #     widget.configure(state="normal", fg_color=["#F9F9FA", "#343638"])

                elif option == "Other":
                    noCA1Entry.configure(state="disabled", fg_color="gray", button_color="gray")
                    for widget in question_entries + marks_entries + [nptelCA1Text, presentationCA1Text]:
                        widget.configure(state="disabled", fg_color="gray")
                            
                        
        #<----------------------------------------------------------------------------------------------------------------------------------------------------------------------->        
                    
        #     def ca2(option):
        #         if  option == "Select Type":
        #             noCA2Entry.configure(state="disabled", fg_color="gray",button_color="gray") 
        #             for disca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2,nptelCA2Text,presentationCA2Text, q1TCA2marks,q2TCA2marks,q3TCA2marks,q4TCA2marks,q5TCA2marks,q6TCA2marks,q7TCA2marks,q8TCA2marks,q9TCA2marks,q10TCA2marks]:
        #                 disca.configure(state="disabled", fg_color="gray")
        #         elif option == "NPTEL Course":
        #             noCA1Entry.configure(state="disabled", fg_color="gray",button_color="gray") 
        #             for disca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2,presentationCA2Text, q1TCA2marks,q2TCA2marks,q3TCA2marks,q4TCA2marks,q5TCA2marks,q6TCA2marks,q7TCA2marks,q8TCA2marks,q9TCA2marks,q10TCA2marks]:
        #                 disca.configure(state="disabled", fg_color="gray")
        #             nptelCA2Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         elif option == "Presentation":
        #             noCA1Entry.configure(state="disabled", fg_color="gray",button_color="gray") 
        #             for preca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2,nptelCA2Text, q1TCA2marks,q2TCA2marks,q3TCA2marks,q4TCA2marks,q5TCA2marks,q6TCA2marks,q7TCA2marks,q8TCA2marks,q9TCA2marks,q10TCA2marks]:
        #                 preca.configure(state="disabled", fg_color="gray")
        #             presentationCA2Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #         elif option == "Test" or option == "Quiz":
        #             # for ca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2, q1TCA2marks,q2TCA2marks,q3TCA2marks,q4TCA2marks,q5TCA2marks,q6TCA2marks,q7TCA2marks,q8TCA2marks,q9TCA2marks,q10TCA2marks]:
        #             #     ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        #             noCA2Entry.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"],button_color=["F9F9FA", "#144870"])
        #             nptelCA2Text.configure(state="disabled", fg_color="gray") 
        #             presentationCA2Text.configure(state="disabled", fg_color="gray")
        #         elif option == "Other":
        #             noCA1Entry.configure(state="disabled", fg_color="gray",button_color="gray") 
        #             for disca in [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2,presentationCA2Text, nptelCA2Text, q1TCA2marks,q2TCA2marks,q3TCA2marks,q4TCA2marks,q5TCA2marks,q6TCA2marks,q7TCA2marks,q8TCA2marks,q9TCA2marks,q10TCA2marks]:
        #                 disca.configure(state="disabled", fg_color="gray")
        #<----------------------------------------------------------------------------------------------------------------------------------------------------------------------->        
            def ca2(option):
                question_entries = list(Question_CA2_CO.values())
                marks_entries = list(Marks_CA2_CO.values())

                if option == "Select Type":
                    noCA2Entry.configure(state="disabled", fg_color="gray", button_color="gray")
                    for widget in question_entries + marks_entries + [nptelCA2Text, presentationCA2Text]:
                        widget.configure(state="disabled", fg_color="gray")

                elif option == "NPTEL Course":
                    noCA2Entry.configure(state="disabled", fg_color="gray", button_color="gray")
                    for widget in question_entries + marks_entries + [presentationCA2Text]:
                        widget.configure(state="disabled", fg_color="gray")
                    nptelCA2Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])

                elif option == "Presentation":
                    noCA2Entry.configure(state="disabled", fg_color="gray", button_color="gray")
                    for widget in question_entries + marks_entries + [nptelCA2Text]:
                        widget.configure(state="disabled", fg_color="gray")
                    presentationCA2Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])

                elif option in ["Test", "Quiz"]:
                    noCA2Entry.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"], button_color=["F9F9FA", "#144870"])
                    nptelCA2Text.configure(state="disabled", fg_color="gray")
                    presentationCA2Text.configure(state="disabled", fg_color="gray")
                    # for widget in question_entries + marks_entries:
                    #     widget.configure(state="normal", fg_color=["#F9F9FA", "#343638"])

                elif option == "Other":
                    noCA2Entry.configure(state="disabled", fg_color="gray", button_color="gray")
                    for widget in question_entries + marks_entries + [nptelCA2Text, presentationCA2Text]:
                        widget.configure(state="disabled", fg_color="gray")
                    

        #<----------------------------------------------------------------------------------------------------------------------------------------------------------------------->        

            # def ca3(option):
            #     if  option == "Select Type":
            #         noCA3Entry.configure(state="disabled", fg_color="gray",button_color="gray") 
            #         for disca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3,nptelCA3Text,presentationCA3Text,q1TCA3marks,q2TCA3marks,q3TCA3marks,q4TCA3marks,q5TCA3marks,q6TCA3marks,q7TCA3marks,q8TCA3marks,q9TCA3marks,q10TCA3marks]:
            #             disca.configure(state="disabled", fg_color="gray")
                    
            #     elif option == "NPTEL Course":
            #         noCA3Entry.configure(state="disabled", fg_color="gray",button_color="gray") 
            #         for disca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3,presentationCA3Text,q1TCA3marks,q2TCA3marks,q3TCA3marks,q4TCA3marks,q5TCA3marks,q6TCA3marks,q7TCA3marks,q8TCA3marks,q9TCA3marks,q10TCA3marks]:
            #             disca.configure(state="disabled", fg_color="gray")
            #         nptelCA3Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])    
                
            #     elif option == "Presentation":
            #         noCA3Entry.configure(state="disabled", fg_color="gray",button_color="gray") 
            #         for preca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3,nptelCA3Text,q1TCA3marks,q2TCA3marks,q3TCA3marks,q4TCA3marks,q5TCA3marks,q6TCA3marks,q7TCA3marks,q8TCA3marks,q9TCA3marks,q10TCA3marks]:
            #             preca.configure(state="disabled", fg_color="gray")
            #         presentationCA3Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
            #     elif option == "Quiz" or option == "Test":
            #         # for ca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3,q1TCA3marks,q2TCA3marks,q3TCA3marks,q4TCA3marks,q5TCA3marks,q6TCA3marks,q7TCA3marks,q8TCA3marks,q9TCA3marks,q10TCA3marks]:
            #         #     ca.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
            #         noCA3Entry.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"],button_color=["F9F9FA", "#144870"])
            #         nptelCA3Text.configure(state="disabled", fg_color="gray") 
            #         presentationCA3Text.configure(state="disabled", fg_color="gray")
            #     elif option == "Other":
            #         noCA1Entry.configure(state="disabled", fg_color="gray",button_color="gray") 
            #         for disca in [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3,presentationCA3Text,nptelCA3Text, q1TCA3marks,q2TCA3marks,q3TCA3marks,q4TCA3marks,q5TCA3marks,q6TCA3marks,q7TCA3marks,q8TCA3marks,q9TCA3marks,q10TCA3marks]:
            #             disca.configure(state="disabled", fg_color="gray")
        #<----------------------------------------------------------------------------------------------------------------------------------------------------------------------->        
            def ca3(option):
                question_entries = list(Question_CA3_CO.values())
                marks_entries = list(Marks_CA3_CO.values())

                if option == "Select Type":
                    noCA3Entry.configure(state="disabled", fg_color="gray", button_color="gray")
                    for widget in question_entries + marks_entries + [nptelCA3Text, presentationCA3Text]:
                        widget.configure(state="disabled", fg_color="gray")

                elif option == "NPTEL Course":
                    noCA3Entry.configure(state="disabled", fg_color="gray", button_color="gray")
                    for widget in question_entries + marks_entries + [presentationCA3Text]:
                        widget.configure(state="disabled", fg_color="gray")
                    nptelCA3Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])

                elif option == "Presentation":
                    noCA3Entry.configure(state="disabled", fg_color="gray", button_color="gray")
                    for widget in question_entries + marks_entries + [nptelCA3Text]:
                        widget.configure(state="disabled", fg_color="gray")
                    presentationCA3Text.configure(state="normal", fg_color=["#F9F9FA", "#343638"])

                elif option in ["Quiz", "Test"]:
                    noCA3Entry.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"], button_color=["F9F9FA", "#144870"])
                    nptelCA3Text.configure(state="disabled", fg_color="gray")
                    presentationCA3Text.configure(state="disabled", fg_color="gray")
                    # for widget in question_entries + marks_entries:
                    #     widget.configure(state="normal", fg_color=["#F9F9FA", "#343638"])

                elif option == "Other":
                    noCA3Entry.configure(state="disabled", fg_color="gray", button_color="gray")
                    for widget in question_entries + marks_entries + [nptelCA3Text, presentationCA3Text]:
                        widget.configure(state="disabled", fg_color="gray")

            def semesterAndClass(option):
                if option == "Select Year":
                    entry2.configure(values=["Select Sem"])
                    entry7.configure(values=["Select Class"])
                elif option == "F.E":
                    entry2.configure(values=["Select Sem","I","II"])
                    entry7.configure(values=["Select Class", "D5A", "D5B", "D5C"])
                elif option == "S.E":
                    entry2.configure(values=["Select Sem","III","IV"])
                    entry7.configure(values=["Select Class", "D10A", "D10B", "D10C"])
                elif option == "T.E":
                    entry2.configure(values=["Select Sem","V","VI"])
                    entry7.configure(values=["Select Class", "D15A", "D15B", "D15C"])
                elif option == "B.E":
                    entry2.configure(values=["Select Sem","VII","VIII"])
                    entry7.configure(values=["Select Class", "D20A", "D20B", "D20C"])
    
        #<----------------------------------------------------------------------------------------------------------------------------------------------------------------------->        

            # def noQuestion1(option):        
            #     caQT = [q1TCA1,q2TCA1,q3TCA1,q4TCA1,q5TCA1,q6TCA1,q7TCA1,q8TCA1,q9TCA1,q10TCA1]
            #     caQM = [q1TCA1marks,q2TCA1marks,q3TCA1marks,q4TCA1marks,q5TCA1marks,q6TCA1marks,q7TCA1marks,q8TCA1marks,q9TCA1marks,q10TCA1marks]
            #     if(option  == "Select No"):
            #         for ca in caQT:
            #             ca.configure(state="disabled", fg_color="gray") 
            #         for ca in caQM:
            #             ca.configure(state="disabled", fg_color="gray") 
            #     option = int(option)
            #     for i in range (0, option):
            #         caQT[i].configure(state="normal", fg_color=["#F9F9FA", "#343638"])
            #         caQM[i].configure(state="normal", fg_color=["#F9F9FA", "#343638"])
            #     for i in range (option, 10):
            #         caQT[i].configure(state="disabled", fg_color="gray")
            #         caQM[i].configure(state="disabled", fg_color="gray")

            # def noQuestion2(option):
            #     caQT = [q1TCA2,q2TCA2,q3TCA2,q4TCA2,q5TCA2,q6TCA2,q7TCA2,q8TCA2,q9TCA2,q10TCA2]
            #     caQM = [q1TCA2marks,q2TCA2marks,q3TCA2marks,q4TCA2marks,q5TCA2marks,q6TCA2marks,q7TCA2marks,q8TCA2marks,q9TCA2marks,q10TCA2marks]
            #     if(option  == "Select No"):
            #         for ca in caQT:
            #             ca.configure(state="disabled", fg_color="gray") 
            #         for ca in caQM:
            #             ca.configure(state="disabled", fg_color="gray") 
            #     option = int(option)
            #     for i in range (0, option):
            #         caQT[i].configure(state="normal", fg_color=["#F9F9FA", "#343638"])
            #         caQM[i].configure(state="normal", fg_color=["#F9F9FA", "#343638"])
            #     for i in range (option, 10):
            #         caQT[i].configure(state="disabled", fg_color="gray")
            #         caQM[i].configure(state="disabled", fg_color="gray")
           
            # def noQuestion3(option):
            #     caQT = [q1TCA3,q2TCA3,q3TCA3,q4TCA3,q5TCA3,q6TCA3,q7TCA3,q8TCA3,q9TCA3,q10TCA3]
            #     caQM = [q1TCA3marks,q2TCA3marks,q3TCA3marks,q4TCA3marks,q5TCA3marks,q6TCA3marks,q7TCA3marks,q8TCA3marks,q9TCA3marks,q10TCA3marks]
                
            #     if(option  == "Select No"):
            #         for ca in caQT:
            #             ca.configure(state="disabled", fg_color="gray") 
            #         for ca in caQM:
            #             ca.configure(state="disabled", fg_color="gray") 
            #     option = int(option)
            #     for i in range (0, option):
            #         caQT[i].configure(state="normal", fg_color=["#F9F9FA", "#343638"])
            #         caQM[i].configure(state="normal", fg_color=["#F9F9FA", "#343638"])
            #     for i in range (option, 10):
            #         caQT[i].configure(state="disabled", fg_color="gray")
            #         caQM[i].configure(state="disabled", fg_color="gray")
         #<----------------------------------------------------------------------------------------------------------------------------------------------------------------------->        
            def noQuestion1(option):
                question_entries = list(Question_CA1_CO.values())
                marks_entries = list(Marks_CA1_CO.values())

                if option == "Select No":
                    for entry in question_entries + marks_entries:
                        entry.configure(state="disabled", fg_color="gray")
                    return

                option = int(option)
                for i in range(option):
                    question_entries[i].configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    marks_entries[i].configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                for i in range(option, 10):
                    question_entries[i].configure(state="disabled", fg_color="gray")
                    marks_entries[i].configure(state="disabled", fg_color="gray")
            
            def noQuestion2(option):
                question_entries = list(Question_CA2_CO.values())
                marks_entries = list(Marks_CA2_CO.values())

                if option == "Select No":
                    for entry in question_entries + marks_entries:
                        entry.configure(state="disabled", fg_color="gray")
                    return

                option = int(option)
                for i in range(option):
                    question_entries[i].configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    marks_entries[i].configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                for i in range(option, 10):
                    question_entries[i].configure(state="disabled", fg_color="gray")
                    marks_entries[i].configure(state="disabled", fg_color="gray")

            def noQuestion3(option):
                question_entries = list(Question_CA3_CO.values())
                marks_entries = list(Marks_CA3_CO.values())

                if option == "Select No":
                    for entry in question_entries + marks_entries:
                        entry.configure(state="disabled", fg_color="gray")
                    return

                option = int(option)

                for i in range(0, option):
                    question_entries[i].configure(state="normal", fg_color=["#F9F9FA", "#343638"])
                    marks_entries[i].configure(state="normal", fg_color=["#F9F9FA", "#343638"])

                for i in range(option, 10):
                    question_entries[i].configure(state="disabled", fg_color="gray")
                    marks_entries[i].configure(state="disabled", fg_color="gray")

                    
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
                question_entries = list(Question_CA3_CO.values())
                marks_entries = list(Marks_CA3_CO.values())

                if option == "Yes":
                    entry15.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"],button_color=["F9F9FA", "#144870"])
                    ALCA3Text.configure(state='normal', fg_color=["#3B8ED0", "#343638"])
                    # Enable question/mark fields if needed:
                    # for entry in question_entries + marks_entries:
                    #     entry.configure(state="normal", fg_color=["#F9F9FA", "#343638"])

                else:
                    entry15.configure(state="disabled", fg_color="gray",button_color="gray")
                    for entry in [noCA3Entry, nptelCA3Text, ALCA3Text] + question_entries + marks_entries:
                        entry.configure(state="disabled", fg_color="gray")

            def validate_academic_year(event):
                new_value = event.widget.get()
    
                if new_value:
                    # Check basic format: Length should be 9, with a '-' in the middle, and both parts should be digits
                    if len(new_value) != 9 or new_value[4] != '-' or not (new_value[:4].isdigit() and new_value[5:].isdigit()):
                        CTkMessagebox(title="Invalid Input", message="Academic Year format is incorrect. Please enter in the format YYYY-YYYY.", icon="warning")
                        return False

                    # Extract years and validate they are consecutive
                    start_year, end_year = int(new_value[:4]), int(new_value[5:])
                    if end_year - start_year != 1:
                        CTkMessagebox(title="Invalid Input", message="Academic Year should be consecutive (e.g., 2025-2026).", icon="warning")
                        return False

                    return True

                return False
    
            def noOfCO(option):
                if option=='6':
                    co_desc_entry["CO6T"].configure(state="normal",fg_color=["#F9F9FA", "#343638"])
                else :
                    co_desc_entry["CO6T"].configure(state="disabled",fg_color="gray")
                    
    
            def resource_path(relative_path):
                """Get the absolute path to a resource, handling PyInstaller paths."""
                if hasattr(sys, '_MEIPASS'):  # PyInstaller extracts files to _MEIPASS
                    return os.path.join(sys._MEIPASS, relative_path)
                return os.path.join(os.path.abspath("."), relative_path)
            
            def validateNumberString(string):
                return string.isdigit()
             
            self.app.destroy() 
            
            co_window = ctk.CTk()  # Close the current window

            screen_width=co_window.winfo_screenwidth()
            screen_height=co_window.winfo_screenheight()
       
            # Set window size (like 80% of screen)
            window_width = int(screen_width * 0.8)
            window_height = int(screen_height * 0.8)
            # Center the window
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2

            # Create a new CO Calculations window
            co_window.title("CO Calculations")
            co_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

            main_frame = ctk.CTkFrame(master=co_window)
            main_frame.pack(expand=True, fill="both", padx=10, pady=10)

            # ---------- Back Button in topbar ----------
            topbar = ctk.CTkFrame(master=main_frame, fg_color="transparent")
            topbar.pack(side="top", fill="x", padx=0, pady=(0, 0))
            back_button = ctk.CTkButton(
                master=topbar,
                text="← Back",
                width=200,
                command=lambda: self.go_back(co_window)
            )
            back_button.pack(side="top", anchor="ne", padx=5)

            # Tabview inside the frame
            tabview = ctk.CTkTabview(main_frame, corner_radius=20)
            tabview.pack(expand=True, fill="both", padx=10, pady=5)
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
            # tabview.add(" Instructions ") 
            # tabview.add(" Basic Information ") 
            # tabview.add(" CO Information ") 
            # tabview.add(" Mid Terms & End Semesters ") 
            # tabview.add(" CA 1 ") 
            # tabview.add(" CA 2 ") 
            # tabview.add(" CA 3 ") 
            # tabview.add(" Process Template/Calculated ")
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
            
            
            # add tab at the end
            # tabview.add(" CO Mapping ")
            # tabview.add(" Target level of tests ")
            # tabview.add(" Lab CO ")

            # Add Tabs
            tab_names = [
                " Instructions ",
                " Basic Information ",
                " CO Information ",
                " Mid Terms & End Semesters ",
                " CA 1 ",
                " CA 2 ",
                " CA 3 ",
                " Process Template/Calculated "
            ]

            scroll_frames = {}  # Dictionary to store scrollable frames by tab name

            for tab_name in tab_names:
                tabview.add(tab_name)

                # Create a scrollable frame inside each tab
                scroll_frame = ctk.CTkScrollableFrame(master=tabview.tab(tab_name), label_text="")
                scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
                scroll_frames[tab_name] = scroll_frame


    
              
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->                           

            # Use resource_path to access the image
            # image_path = resource_path(f"./images/coCal.png")
            # # Load the image and create a CTkImage
            # background_image = Image.open(image_path)
            # bg_image = ctk.CTkImage(background_image, size=(screen_width - 100, screen_height-130))

            # # Create a frame for the "Instructions" tab content
            # instructions_tab = tabview.tab(" Instructions ")
            # instructions_tab.columnconfigure(0, weight=1)
            # instructions_tab.rowconfigure(0, weight=1)

            # # Add a label to hold the background image
            # bg_label = ctk.CTkLabel(master=instructions_tab, image=bg_image, text="")
            # bg_label.place(relx=0.5, rely=0.5, anchor="center")

            # # If you want to overlay widgets on top of the image:
            # # Example of overlaying text on the background
            # overlay_label = ctk.CTkLabel(master=instructions_tab, text="")
            # overlay_label.place(relx=0.5, rely=0.1, anchor="center")
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   

            

            # Reference the scrollable frame for the " Instructions " tab
            instructions_frame = scroll_frames[" Instructions "]

            width12 = instructions_frame.winfo_width()
            height12 = instructions_frame.winfo_height()
            # Load the image using PIL and create CTkImage with fixed size
            image_path = resource_path(f"./images/coCal.png")
            background_image = Image.open(image_path)
            bg_image = ctk.CTkImage(background_image,size=(1200,600))

            # Configure grid in scrollable frame
            instructions_frame.grid_columnconfigure(0, weight=1)
            instructions_frame.grid_rowconfigure(0, weight=1)

            # Add the label with image using .grid()
            bg_label = ctk.CTkLabel(master=instructions_frame, image=bg_image, text="")
            bg_label.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

            # Optional: Add overlay text or widgets
            # overlay_label = ctk.CTkLabel(master=instructions_frame, text="Welcome to Instructions", font=("Arial", 20))
            # overlay_label.grid(row=1, column=0, pady=(10, 20))

#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   

            # button = create_button(" Basic Information ", "Next", "Arial", 20, 150, 40, switch_to_co_information, 725, 490)

            # label0 = create_label(" Basic Information ", "Basic Details", "Arial", 20, 725, 5)

            # label1 = create_label(" Basic Information ", "No. of Students :", "Arial", 15, 550, 55)

            # entry1 = create_entry_box(" Basic Information ", "Enter no of students", "Arial", 15, 300, 750, 55)

            # newLabel = create_label(" Basic Information ", "Year :", "Arial", 15, 550, 155)

            # yearDropDown = create_dropdown(" Basic Information ", ["Select Year", "F.E", "S.E", "T.E", "B.E"], "Arial", 15, 300, semesterAndClass, 750, 155)

            # label8 = create_label(" Basic Information ", "Department :", "Arial", 15, 550, 105)

            # entry8 = create_dropdown(" Basic Information ", ["Select Department", "Humanities and Applied Science(FE)", "Information Technology", "Computer", "AI and Data Science", "Electronics and Telecommunication", "Electronics", "Instrumentation"], "Arial", 15, 300, None, 750, 105)

            # label2 = create_label(" Basic Information ", "Semester :", "Arial", 15,550, 205)

            # entry2 = create_dropdown(" Basic Information ", ["Select Sem"], "Arial", 15, 300, subject,750, 205)

            # label3 = create_label(" Basic Information ", "Subject :", "Arial", 15, 550, 255)

            # entry3 = create_dropdown(" Basic Information ", ["Select Subject"], "Arial", 15, 300, None, 750, 255)

            # label4 = create_label(" Basic Information ", "Academic Year: ", "Arial", 15, 550, 305)

            # entry4 = create_entry_box(" Basic Information ", "YYYY-YYYY", "Arial", 15, 300, 750, 305)
            # entry4.bind("<FocusOut>", validate_academic_year)

            # label5 = create_label(" Basic Information ", "Subject Teacher :", "Arial", 15, 550, 355)

            # entry5 = create_entry_box(" Basic Information ", "Subject Teacher", "Arial", 15, 300, 750, 355)

            # label7 = create_label(" Basic Information ", "Class :", "Arial", 15, 550, 405)

            # # entry7 = create_entry_box(" Basic Information ", "Eg.D10 C", "Arial", 15, 300, 400, 405)

            # entry7 = create_dropdown(" Basic Information ", ["Select Class"], "Arial", 15, 300, None, 750, 405)

            # # label12 = create_label(" Basic Information ", "Attainment Target :", "Arial", 15, 875, 55)

            # # entry12 = create_entry_box(" Basic Information ", "52.5", "Arial", 15, 300, 1075, 55)
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
            tab_name = " Basic Information "
            
            for i in range(5):
                scroll_frames[tab_name].grid_columnconfigure(i, weight=1)

            row_base = 0
            label0 = create_label(tab_name, "Basic Details", "Arial", 20, row=row_base, column=0, colspan=5)
            label0.configure(anchor="center", justify="center")
            label0.grid_configure(sticky="nsew")
            row_base +=2

            label_text_basic=["No. of Students :","Department :","Year :","Semester :","Subject :","Academic Year: ","Subject Teacher :","Class :"]
            for text in label_text_basic :
                create_label(tab_name, text, "Arial", 15, row=row_base, column=2, sticky="nsw")
                row_base+=1
            
            row_base=2
            entry1= create_entry_box(tab_name, "Enter no of students", "Arial", 15, 300, row=row_base, column=3,sticky="nsw")
            row_base += 1
            
            entry8 = create_dropdown(tab_name, ["Select Department", "Humanities and Applied Science(FE)", "Information Technology", "Computer", "AI and Data Science", "Electronics and Telecommunication", "Electronics", "Instrumentation"], "Arial", 15, 300, None, row=row_base, column=3,sticky="nsw")
            row_base +=1
            
            yearDropDown = create_dropdown(tab_name, ["Select Year", "F.E", "S.E", "T.E", "B.E"], "Arial", 15, 300, semesterAndClass, row=row_base, column=3,sticky="nsw")
            row_base +=1

            entry2 = create_dropdown(tab_name, ["Select Sem"], "Arial", 15, 300, subject, row=row_base, column=3,sticky="nsw")
            row_base +=1

            entry3 = create_dropdown(tab_name, ["Select Subject"], "Arial", 15, 300, subject, row=row_base, column=3,sticky="nsw")
            row_base +=1
             
            entry4= create_entry_box(tab_name, "YYYY-YYYY", "Arial", 15, 300, row=row_base, column=3,sticky="nsw")
            entry4.bind("<FocusOut>", validate_academic_year)
            row_base += 1

            entry5= create_entry_box(tab_name, "Subject Teacher", "Arial", 15, 300, row=row_base, column=3,sticky="nsw")
            row_base += 1

            entry7 = create_dropdown(tab_name, ["Select Class"], "Arial", 15, 300, None, row=row_base, column=3,sticky="nsw")
            row_base +=2

            button = create_button(tab_name, "Next", "Arial", 20, 300, 40, switch_to_co_information, row=row_base, column=2,colspan=2,sticky="")


#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   

            # label6 = create_label(" Mid Terms & End Semesters ", "COs for Midterm", "Arial", 20, 725, 20)
            # label11 = create_label(" Mid Terms & End Semesters ", "Endsems CO's", "Arial", 20, 725, 390)

            # entry11 = create_entry_box(" Mid Terms & End Semesters ", "1,2,3,4,5,6", "Arial", 15, 300, 650, 425)

            # a1L = create_label(" Mid Terms & End Semesters ", "1a :", "Arial", 15, 550, 60)
            # a2L = create_label(" Mid Terms & End Semesters ", "1b :", "Arial", 15, 550, 110)
            # a3L = create_label(" Mid Terms & End Semesters ", "1c :", "Arial", 15, 550, 160)
            # a4L = create_label(" Mid Terms & End Semesters ", "1d :", "Arial", 15, 550, 210)
            # a5L = create_label(" Mid Terms & End Semesters ", "1e :", "Arial", 15, 550, 260)
            # a6L = create_label(" Mid Terms & End Semesters ", "1f :", "Arial", 15, 550, 310)
            # a2aL = create_label(" Mid Terms & End Semesters ", "2a :", "Arial", 15, 850, 60)
            # a2bL = create_label(" Mid Terms & End Semesters ", "2b :", "Arial", 15, 850, 110)
            # a3aL = create_label(" Mid Terms & End Semesters ", "3a :", "Arial", 15, 850, 160)
            # a3bL = create_label(" Mid Terms & End Semesters ", "3b :", "Arial", 15, 850, 210)


            # a1T=ctk.CTkEntry(master=tabview.tab(" Mid Terms & End Semesters "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            # a1T.place(x=600,y=60)

            # a2T=ctk.CTkEntry(master=tabview.tab(" Mid Terms & End Semesters "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            # a2T.place(x=600,y=110)

            # a3T=ctk.CTkEntry(master=tabview.tab(" Mid Terms & End Semesters "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            # a3T.place(x=600,y=160)

            # a4T=ctk.CTkEntry(master=tabview.tab(" Mid Terms & End Semesters "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            # a4T.place(x=600,y=210)

            # a5T=ctk.CTkEntry(master=tabview.tab(" Mid Terms & End Semesters "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            # a5T.place(x=600,y=260)

            # a6T=ctk.CTkEntry(master=tabview.tab(" Mid Terms & End Semesters "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            # a6T.place(x=600,y=310)


            # a2aT=ctk.CTkEntry(master=tabview.tab(" Mid Terms & End Semesters "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            # a2aT.place(x=900,y=60)


            # a2bT=ctk.CTkEntry(master=tabview.tab(" Mid Terms & End Semesters "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            # a2bT.place(x=900,y=110)


            # a3aT=ctk.CTkEntry(master=tabview.tab(" Mid Terms & End Semesters "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            # a3aT.place(x=900,y=160)

            
            # a3bT=ctk.CTkEntry(master=tabview.tab(" Mid Terms & End Semesters "),placeholder_text="1,2,3,4,5,6",font=("Arial",15),width=150)
            # a3bT.place(x=900,y=210)

            
            # ALlabelMidTerm = create_label(" Mid Terms & End Semesters ", "Enter the Target level for Midterms and End Semsesters", "Arial", 20, 600, 505)
            # ALMidTermLabel = create_label(" Mid Terms & End Semesters ", "Mid Term: ", "Arial", 15, 450, 555)
            # ALMidTermText = create_entry_box(" Mid Terms & End Semesters ", "", "Arial", 15, 500, 550, 555)
            # ALEndSemLabel = create_label(" Mid Terms & End Semesters ", "End Semester: ", "Arial", 15, 450, 605)
            # ALEndSemText = create_entry_box(" Mid Terms & End Semesters ", "", "Arial", 15, 500, 550, 605)

            # button2 = create_button(" Mid Terms & End Semesters ", "Next", "Arial", 20, 200, 40, switch_to_CA1, 1050, 640)
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
            tab_name = " Mid Terms & End Semesters "
            
            for i in range(6):
                scroll_frames[tab_name].grid_columnconfigure(i, weight=1)

            row_base = 0
            label6= create_label(tab_name, "COs for Midterm", "Arial", 20, row=row_base, column=1, colspan=2)
            label6.configure(anchor="center", justify="center")
            label6.grid_configure(sticky="nsew")
            row_base +=1 

            Question_no=['Q.1a : ','Q.1b : ','Q.1c : ','Q.1d : ','Q.1e : ','Q.1f : ','Q.2a : ','Q.2b : ','Q.3a : ','Q.3b : ']
            midterm_co_entry={}

            index=1
            for lab_CO_mid in Question_no:
                create_label(tab_name, lab_CO_mid, "Arial", 15, row=row_base, column=1)
                midterm_co_entry[f"Q.{index}"]=create_entry_box(tab_name, "1,2,3,4,5,6", "Arial", 15, 200, row=row_base, column=2)
                row_base+=1
                index+=1

            vertical_line_co = ctk.CTkFrame(scroll_frames[tab_name], width=2, fg_color="white")
            vertical_line_co.grid(row=1, column=3, rowspan=19, sticky="ns", padx=5)

            row_base = 0
            label11= create_label(tab_name, "Endsems CO's", "Arial", 20, row=row_base, column=4, colspan=2)
            label11.configure(anchor="center", justify="center")
            label11.grid_configure(sticky="nsew")
            row_base +=1
            
            label11 = create_label(tab_name, "Endsems CO's : ", "Arial", 15, row=row_base, column=4)
            entry11 = create_entry_box(tab_name, "1,2,3,4,5,6", "Arial", 15, 200, row=row_base, column=5)

            row_base+=2
            ALlabelMidTerm= create_label(tab_name, "Target level for Midterms", "Arial", 20, row=row_base, column=4, colspan=2)
            ALlabelMidTerm.configure(anchor="center", justify="center")
            ALlabelMidTerm.grid_configure(sticky="nsew")

            row_base+=1
            ALMidTermLabel = create_label(tab_name, "Mid Term : ", "Arial", 15, row=row_base, column=4)
            ALMidTermText = create_entry_box(tab_name, "Eg.50", "Arial", 15, 200, row=row_base, column=5)
            
            row_base+=2
            ALlabelEndTerm= create_label(tab_name, "Target level for End Semesters", "Arial", 20, row=row_base, column=4, colspan=2)
            ALlabelEndTerm.configure(anchor="center", justify="center")
            ALlabelEndTerm.grid_configure(sticky="nsew")
            
            row_base+=1
            ALEndSemLabel =  create_label(tab_name, "End Semester : ", "Arial", 15, row=row_base, column=4)
            ALEndSemText = create_entry_box(tab_name, "Eg.50", "Arial", 15, 200, row=row_base, column=5)

            row_base+=1
            
            button2 = create_button(tab_name, "Next", "Arial", 20, 250, 40, switch_to_CA1, row=row_base, column=5)
            button2.grid(row=row_base, column=4,columnspan=2, rowspan=2, padx=5, pady=5, sticky="")

#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   

            # # COs for CA1 Quiz
            # label13 = create_label(" CA 1 ", "CA1 type :", "Arial", 15, 200, 55)
            # entry13 = create_dropdown(" CA 1 ", ["Select Type", "Quiz", "NPTEL Course", "Presentation", "Test",  "Other"], "Arial", 15, 300, ca1, 500, 55)
            # noCA1Label = create_label(" CA 1 ", "No of Question CA1 (Quiz/Test) :", "Arial", 15, 200, 105)
            # noCA1Entry = create_dropdown(" CA 1 ", ["Select No", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], "Arial", 15, 300, noQuestion1, 500, 105)
            # noCA1Entry.configure(state="disabled", fg_color="gray")
            # nptelCA1 = create_label(" CA 1 ", "CO's for NPTEL (CA)", "Arial", 20, 470, 505)
            # nptelCA1Label = create_label(" CA 1 ", "NPTEL: ", "Arial", 15, 350, 555)
            # nptelCA1Text = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 300, 450, 555)
            # nptelCA1Text.configure(state="disabled", fg_color="gray")
            # presentationcA1 = create_label(" CA 1 ", "Maximum group size of Presentations (CA)", "Arial", 20, 870, 505)
            # presentationCA1Label = create_label(" CA 1 ", "Group Size: ", "Arial", 15, 850, 555)
            # presentationCA1Text = create_entry_box(" CA 1 ", "Enter maximum number of students in a group", "Arial", 15, 350, 950, 555)
            # presentationCA1Text.configure(state="disabled", fg_color="gray")


            # label9 = create_label(" CA 1 ", "COs for CA1 Quiz/Test", "Arial", 20, 350, 205)
            # label9_marks = create_label(" CA 1 ", "Marks of Questions for CA1 Quiz/Test", "Arial", 20, 1000, 205)

            # q1LCA1 = create_label(" CA 1 ", "Q1 :", "Arial", 15, 200, 255)
            # q1TCA1 = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 255)
            # q1TCA1.configure(state="disabled", fg_color="gray")
            
            # q1LCA1 = create_label(" CA 1 ", "Q1 :", "Arial", 15, 200, 255)
            # q1TCA1 = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 255)
            # q1TCA1.configure(state="disabled", fg_color="gray")

            # q2LCA1 = create_label(" CA 1 ", "Q2 :", "Arial", 15, 200, 305)
            # q2TCA1 = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 305)
            # q2TCA1.configure(state="disabled", fg_color="gray")

            # q3LCA1 = create_label(" CA 1 ", "Q3 :", "Arial", 15, 200, 355)
            # q3TCA1 = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 355)
            # q3TCA1.configure(state="disabled", fg_color="gray")

            # q4LCA1 = create_label(" CA 1 ", "Q4 :", "Arial", 15, 200, 405)
            # q4TCA1 = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 405)
            # q4TCA1.configure(state="disabled", fg_color="gray")

            # q5LCA1 = create_label(" CA 1 ", "Q5 :", "Arial", 15, 200, 455)
            # q5TCA1 = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 455)
            # q5TCA1.configure(state="disabled", fg_color="gray")

            # q6LCA1 = create_label(" CA 1 ", "Q6 :", "Arial", 15, 500, 255)
            # q6TCA1 = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 255)
            # q6TCA1.configure(state="disabled", fg_color="gray")

            # q7LCA1 = create_label(" CA 1 ", "Q7 :", "Arial", 15, 500, 305)
            # q7TCA1 = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 305)
            # q7TCA1.configure(state="disabled", fg_color="gray")

            # q8LCA1 = create_label(" CA 1 ", "Q8 :", "Arial", 15, 500, 355)
            # q8TCA1 = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 355)
            # q8TCA1.configure(state="disabled", fg_color="gray")

            # q9LCA1 = create_label(" CA 1 ", "Q9 :", "Arial", 15, 500, 405)
            # q9TCA1 = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 550,405)
            # q9TCA1.configure(state="disabled", fg_color="gray")

            # q10LCA1 = create_label(" CA 1 ", "Q10 :", "Arial", 15, 500, 455)
            # q10TCA1 = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 455)
            # q10TCA1.configure(state="disabled", fg_color="gray")

            # q1LCA1marks = create_label(" CA 1 ", "Q1 :", "Arial", 15, 900, 255)
            # q1TCA1marks = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 950, 255)
            # q1TCA1marks.configure(state="disabled", fg_color="gray")

            # q2LCA1marks = create_label(" CA 1 ", "Q2 :", "Arial", 15, 900, 305)
            # q2TCA1marks = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 950, 305)
            # q2TCA1marks.configure(state="disabled", fg_color="gray")

            # q3LCA1marks = create_label(" CA 1 ", "Q3 :", "Arial", 15, 900, 355)
            # q3TCA1marks = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 950, 355)
            # q3TCA1marks.configure(state="disabled", fg_color="gray")

            # q4LCA1marks = create_label(" CA 1 ", "Q4 :", "Arial", 15, 900, 405)
            # q4TCA1marks = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 950, 405)
            # q4TCA1marks.configure(state="disabled", fg_color="gray")

            # q5LCA1marks = create_label(" CA 1 ", "Q5 :", "Arial", 15, 900, 455)
            # q5TCA1marks = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 950, 455)
            # q5TCA1marks.configure(state="disabled", fg_color="gray")

            # q6LCA1marks = create_label(" CA 1 ", "Q6 :", "Arial", 15, 1200, 255)
            # q6TCA1marks = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 1250, 255)
            # q6TCA1marks.configure(state="disabled", fg_color="gray")

            # q7LCA1marks = create_label(" CA 1 ", "Q7 :", "Arial", 15, 1200, 305)
            # q7TCA1marks = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 1250, 305)
            # q7TCA1marks.configure(state="disabled", fg_color="gray")

            # q8LCA1marks = create_label(" CA 1 ", "Q8 :", "Arial", 15, 1200, 355)
            # q8TCA1marks = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 1250, 355)
            # q8TCA1marks.configure(state="disabled", fg_color="gray")

            # q9LCA1marks = create_label(" CA 1 ", "Q9 :", "Arial", 15, 1200, 405)
            # q9TCA1marks = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 1250,405)
            # q9TCA1marks.configure(state="disabled", fg_color="gray")

            # q10LCA1marks = create_label(" CA 1 ", "Q10 :", "Arial", 15, 1200, 455)
            # q10TCA1marks = create_entry_box(" CA 1 ", "1,2,3,4,5,6", "Arial", 15, 150, 1250, 455)
            # q10TCA1marks.configure(state="disabled", fg_color="gray")

            
            # # CA1, CA2, CA3, MidTerm, EndSem, Labs
            # ALlabelCA1 = create_label(" CA 1 ", "Enter the Target level for CA 1", "Arial", 20, 600, 605)
            # ALCA1Label = create_label(" CA 1 ", "CA1: ", "Arial", 15, 450, 655)
            # ALCA1Text = create_entry_box(" CA 1 ", "", "Arial", 15, 500, 550, 655)

            # button3 = create_button(" CA 1 ", "Next", "Arial", 20, 200, 40, switch_to_CA2, 1100, 640)
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
            tab_name = " CA 1 "
            
            for i in range(7):
                scroll_frames[tab_name].grid_columnconfigure(i, weight=1)

            row_base=0

            labelCA1= create_label(tab_name, "CA 1 Details", "Arial", 20, row=row_base, column=1, colspan=2)
            labelCA1.configure(anchor="center", justify="center")
            labelCA1.grid_configure(sticky="nsew")
            row_base +=1

            label13 = create_label(tab_name, "CA1 type :", "Arial", 15, row=row_base, column=1, sticky="nsw")
            entry13 = create_dropdown(tab_name, ["Select Type", "Quiz", "NPTEL Course", "Presentation", "Test",  "Other"], "Arial", 15, 300, ca1, row=row_base, column=2,sticky="nsw")
            row_base+=1
            
            noCA1Label = create_label(tab_name, "No of Question CA1 (Quiz/Test) :", "Arial", 15, row=row_base, column=1, sticky="nsw")
            noCA1Entry = create_dropdown(tab_name, ["Select No", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], "Arial", 15, 300, noQuestion1, row=row_base, column=2,sticky="nsw")
            noCA1Entry.configure(state="disabled", fg_color="gray",button_color="gray")
            row_base+=1

            merged_frame1 = ctk.CTkFrame(scroll_frames[tab_name])
            merged_frame1.grid(row=row_base, column=1, columnspan=2, padx=5, pady=5, sticky="nsew")

            for sub_col in range(3):
                merged_frame1.grid_columnconfigure(sub_col, weight=1)

            label_title_1 = ctk.CTkLabel(merged_frame1, text="No.")
            label_title_1.grid(row=0, column=0, padx=5, pady=5, sticky="nsw")

            label_title_2 = ctk.CTkLabel(merged_frame1, text="COs")
            label_title_2.grid(row=0, column=1, padx=5, pady=5, sticky="nsw")

            label_title_3 = ctk.CTkLabel(merged_frame1, text="Marks")
            label_title_3.grid(row=0, column=2, padx=5, pady=5, sticky="nsw")

            vertical_line_ca1 = ctk.CTkFrame(scroll_frames[tab_name], width=2, fg_color="white")
            vertical_line_ca1.grid(row=1, column=3, rowspan=19, sticky="ns", padx=5)

            Question_CA1_CO={}
            Marks_CA1_CO={}

            for i in range(1,11):
                row_base+=1

                # Create a frame to be placed inside merged col 2+3
                merged_frame = ctk.CTkFrame(scroll_frames[tab_name])
                merged_frame.grid(row=row_base, column=1, columnspan=2, padx=5, pady=5, sticky="nsew")

                # Configure 3 sub-columns inside merged frame
                for sub_col in range(3):
                    merged_frame.grid_columnconfigure(sub_col, weight=1)

                # Label inside merged frame
                if i==10:
                    label = ctk.CTkLabel(merged_frame, text=f"Q.{i}:")
                else :
                    label = ctk.CTkLabel(merged_frame, text=f"Q.{i} : ")
                label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

                # Entry 1
                Question_CA1_CO[f"Q{i}_questions"] = ctk.CTkEntry(merged_frame, placeholder_text="1,2,3,4,5,6",font=("Arial", 15), width=200)
                Question_CA1_CO[f"Q{i}_questions"].grid(row=0, column=1, padx=5, pady=5, sticky="w")
                Question_CA1_CO[f"Q{i}_questions"].configure(state="disabled", fg_color="gray")

                # Entry 2
                Marks_CA1_CO[f"Q{i}_marks"] = ctk.CTkEntry(merged_frame, placeholder_text="Eg.10",font=("Arial", 15), width=200)
                Marks_CA1_CO[f"Q{i}_marks"].grid(row=0, column=2, padx=5, pady=5, sticky="w")
                Marks_CA1_CO[f"Q{i}_marks"].configure(state="disabled", fg_color="gray")

            row_base=0
            nptelCA1 = create_label(tab_name, "CO's for NPTEL (CA)", "Arial", 20, row=row_base, column=4,colspan=2, sticky="nsew")
            row_base+=1
            nptelCA1Label = create_label(tab_name, "NPTEL:", "Arial", 15, row=row_base, column=4, sticky="nsw")
            nptelCA1Text = create_entry_box(tab_name, "1,2,3,4,5,6", "Arial", 15, 300, row=row_base, column=5,sticky="nsw")
            nptelCA1Text.configure(state="disabled", fg_color="gray")
            row_base+=2
           
            presentationcA1 = create_label(tab_name, "Maximum group size of Presentations (CA)", "Arial", 20, row=row_base, column=4,colspan=2, sticky="nsew")
            row_base+=1
            presentationCA1Label = create_label(tab_name, "Group Size: ", "Arial", 15, row=row_base, column=4, sticky="nsw")
            presentationCA1Text = create_entry_box(tab_name, "Enter maximum number of students in a group", "Arial", 15, 300, row=row_base, column=5,sticky="nsw")
            presentationCA1Text.configure(state="disabled", fg_color="gray")
            row_base+=2
            
            ALlabelCA1 = create_label(tab_name, "Enter the Target level for CA 1", "Arial", 20, row=row_base, column=4,colspan=2, sticky="nsew")
            row_base+=1
            ALCA1Label = create_label(tab_name, "CA1: ", "Arial", 15, row=row_base, column=4, sticky="nsw")
            ALCA1Text = create_entry_box(tab_name, "Eg.50", "Arial", 15, 300, row=row_base, column=5,sticky="nsw")
            row_base+=1

            button3 = create_button(tab_name, "Next", "Arial", 20, 250, 40, switch_to_CA2, row=row_base, column=5)
            button3.grid(row=row_base, column=4,columnspan=2, rowspan=2, padx=5, pady=5, sticky="")

#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   

            # # COs for CA2 Quiz
            # label14 = create_label(" CA 2 ", "CA2 type :", "Arial", 15, 200, 55)
            # entry14 = create_dropdown(" CA 2 ", ["Select Type", "Quiz", "NPTEL Course", "Presentation", "Test", "Other"], "Arial", 15, 300, ca2, 500, 55)
            # noCA2Label = create_label(" CA 2 ", "No of Question CA2 (Quiz/Test) :", "Arial", 15, 200, 105)
            # noCA2Entry = create_dropdown(" CA 2 ", ["Select No", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], "Arial", 15, 300, noQuestion2, 500, 105)
            # noCA2Entry.configure(state="disabled", fg_color="gray")
            # nptelCA2 = create_label(" CA 2 ", "CO's for NPTEL (CA)", "Arial", 20, 470, 505)
            # nptelCA2Label = create_label(" CA 2 ", "NPTEL: ", "Arial", 15, 350, 555)
            # nptelCA2Text = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 300, 450, 555)
            # nptelCA2Text.configure(state="disabled", fg_color="gray")
            # presentationCA2 = create_label(" CA 2 ", "Maximum group size of Presentations (CA)", "Arial", 20, 870, 505)
            # presentationCA2Label = create_label(" CA 2 ", "Group Size: ", "Arial", 15, 850, 555)
            # presentationCA2Text = create_entry_box(" CA 2 ", "Enter maximum number of students in a group", "Arial", 15, 350, 950, 555)
            # presentationCA2Text.configure(state="disabled", fg_color="gray")


            # label18 = create_label(" CA 2 ", "COs for CA2 Quiz/Test", "Arial", 20, 350, 205)
            # label18marks = create_label(" CA 2 ", "Marks of Questions for CA2 Quiz/Test", "Arial", 20, 1000, 205)

            # q1LCA2 = create_label(" CA 2 ", "Q1 :", "Arial", 15, 200, 255)
            # q1TCA2 = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 255)
            # q1TCA2.configure(state="disabled", fg_color="gray")

            # q2LCA2 = create_label(" CA 2 ", "Q2 :", "Arial", 15, 200, 305)
            # q2TCA2 = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 305)
            # q2TCA2.configure(state="disabled", fg_color="gray")

            # q3LCA2 = create_label(" CA 2 ", "Q3 :", "Arial", 15, 200, 355)
            # q3TCA2 = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 355)
            # q3TCA2.configure(state="disabled", fg_color="gray")

            # q4LCA2 = create_label(" CA 2 ", "Q4 :", "Arial", 15, 200, 405)
            # q4TCA2 = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 405)
            # q4TCA2.configure(state="disabled", fg_color="gray")

            # q5LCA2 = create_label(" CA 2 ", "Q5 :", "Arial", 15, 200, 455)
            # q5TCA2 = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 455)
            # q5TCA2.configure(state="disabled", fg_color="gray")

            # q6LCA2 = create_label(" CA 2 ", "Q6 :", "Arial", 15, 500, 255)
            # q6TCA2 = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 255)
            # q6TCA2.configure(state="disabled", fg_color="gray")

            # q7LCA2 = create_label(" CA 2 ", "Q7 :", "Arial", 15,500, 305)
            # q7TCA2 = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 305)
            # q7TCA2.configure(state="disabled", fg_color="gray")

            # q8LCA2 = create_label(" CA 2 ", "Q8 :", "Arial", 15, 500, 355)
            # q8TCA2 = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 355)
            # q8TCA2.configure(state="disabled", fg_color="gray")

            # q9LCA2 = create_label(" CA 2 ", "Q9 :", "Arial", 15, 500, 405)
            # q9TCA2 = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 405)
            # q9TCA2.configure(state="disabled", fg_color="gray")

            # q10LCA2 = create_label(" CA 2 ", "Q10 :", "Arial", 15, 500, 455)
            # q10TCA2 = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 455)
            # q10TCA2.configure(state="disabled", fg_color="gray")
            
            # q1LCA2marks = create_label(" CA 2 ", "Q1 :", "Arial", 15, 900, 255)
            # q1TCA2marks = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 950, 255)
            # q1TCA2marks.configure(state="disabled", fg_color="gray")

            # q2LCA2marks = create_label(" CA 2 ", "Q2 :", "Arial", 15, 900, 305)
            # q2TCA2marks = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 950, 305)
            # q2TCA2marks.configure(state="disabled", fg_color="gray")

            # q3LCA2marks = create_label(" CA 2 ", "Q3 :", "Arial", 15, 900, 355)
            # q3TCA2marks = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 950, 355)
            # q3TCA2marks.configure(state="disabled", fg_color="gray")

            # q4LCA2marks = create_label(" CA 2 ", "Q4 :", "Arial", 15, 900, 405)
            # q4TCA2marks = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 950, 405)
            # q4TCA2marks.configure(state="disabled", fg_color="gray")

            # q5LCA2marks = create_label(" CA 2 ", "Q5 :", "Arial", 15, 900, 455)
            # q5TCA2marks = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 950, 455)
            # q5TCA2marks.configure(state="disabled", fg_color="gray")

            # q6LCA2marks = create_label(" CA 2 ", "Q6 :", "Arial", 15, 1200, 255)
            # q6TCA2marks = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 1250, 255)
            # q6TCA2marks.configure(state="disabled", fg_color="gray")

            # q7LCA2marks = create_label(" CA 2 ", "Q7 :", "Arial", 15, 1200, 305)
            # q7TCA2marks = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 1250, 305)
            # q7TCA2marks.configure(state="disabled", fg_color="gray")

            # q8LCA2marks = create_label(" CA 2 ", "Q8 :", "Arial", 15, 1200, 355)
            # q8TCA2marks = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 1250, 355)
            # q8TCA2marks.configure(state="disabled", fg_color="gray")

            # q9LCA2marks = create_label(" CA 2 ", "Q9 :", "Arial", 15, 1200, 405)
            # q9TCA2marks = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 1250,405)
            # q9TCA2marks.configure(state="disabled", fg_color="gray")

            # q10LCA2marks = create_label(" CA 2 ", "Q10 :", "Arial", 15, 1200, 455)
            # q10TCA2marks = create_entry_box(" CA 2 ", "1,2,3,4,5,6", "Arial", 15, 150, 1250, 455)
            # q10TCA2marks.configure(state="disabled", fg_color="gray")
            
            # ALlabelCA2 = create_label(" CA 2 ", "Enter the Target level for CA 2", "Arial", 20, 600, 605)
            # ALCA2Label = create_label(" CA 2 ", "CA2: ", "Arial", 15, 450, 655)
            # ALCA2Text = create_entry_box(" CA 2 ", "", "Arial", 15, 500, 550, 655)

            # button2 = create_button(" CA 2 ", "Next", "Arial", 20, 200, 40, switch_to_CA3, 1100, 640)
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
            tab_name = " CA 2 "
            
            for i in range(7):
                scroll_frames[tab_name].grid_columnconfigure(i, weight=1)

            row_base=0

            labelCA2= create_label(tab_name, "CA 2 Details", "Arial", 20, row=row_base, column=1, colspan=2)
            labelCA2.configure(anchor="center", justify="center")
            labelCA2.grid_configure(sticky="nsew")
            row_base +=1

            label14 = create_label(tab_name, "CA2 type :", "Arial", 15, row=row_base, column=1, sticky="nsw")
            entry14 = create_dropdown(tab_name, ["Select Type", "Quiz", "NPTEL Course", "Presentation", "Test",  "Other"], "Arial", 15, 300, ca2, row=row_base, column=2,sticky="nsw")
            row_base+=1
            
            noCA2Label = create_label(tab_name, "No of Question CA2 (Quiz/Test) :", "Arial", 15, row=row_base, column=1, sticky="nsw")
            noCA2Entry = create_dropdown(tab_name, ["Select No", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], "Arial", 15, 300, noQuestion2, row=row_base, column=2,sticky="nsw")
            noCA2Entry.configure(state="disabled", fg_color="gray",button_color="gray")
            row_base+=1

            merged_frame2 = ctk.CTkFrame(scroll_frames[tab_name])
            merged_frame2.grid(row=row_base, column=1, columnspan=2, padx=5, pady=5, sticky="nsew")

            for sub_col in range(3):
                merged_frame2.grid_columnconfigure(sub_col, weight=1)

            label_title_4 = ctk.CTkLabel(merged_frame2, text="No.")
            label_title_4.grid(row=0, column=0, padx=5, pady=5, sticky="nsw")

            label_title_5 = ctk.CTkLabel(merged_frame2, text="COs")
            label_title_5.grid(row=0, column=1, padx=5, pady=5, sticky="nsw")

            label_title_6 = ctk.CTkLabel(merged_frame2, text="Marks")
            label_title_6.grid(row=0, column=2, padx=5, pady=5, sticky="nsw")

            vertical_line_ca2 = ctk.CTkFrame(scroll_frames[tab_name], width=2, fg_color="white")
            vertical_line_ca2.grid(row=1, column=3, rowspan=19, sticky="ns", padx=5)

            Question_CA2_CO={}
            Marks_CA2_CO={}

            for i in range(1,11):
                row_base+=1

                # Create a frame to be placed inside merged col 2+3
                merged_frame = ctk.CTkFrame(scroll_frames[tab_name])
                merged_frame.grid(row=row_base, column=1, columnspan=2, padx=5, pady=5, sticky="nsew")

                # Configure 3 sub-columns inside merged frame
                for sub_col in range(3):
                    merged_frame.grid_columnconfigure(sub_col, weight=1)

                # Label inside merged frame
                if i==10:
                    label = ctk.CTkLabel(merged_frame, text=f"Q.{i}:")
                else :
                    label = ctk.CTkLabel(merged_frame, text=f"Q.{i} : ")
                label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

                # Entry 1
                Question_CA2_CO[f"Q{i}_questions"] = ctk.CTkEntry(merged_frame, placeholder_text="1,2,3,4,5,6",font=("Arial", 15), width=200)
                Question_CA2_CO[f"Q{i}_questions"].grid(row=0, column=1, padx=5, pady=5, sticky="w")
                Question_CA2_CO[f"Q{i}_questions"].configure(state="disabled", fg_color="gray")

                # Entry 2
                Marks_CA2_CO[f"Q{i}_marks"] = ctk.CTkEntry(merged_frame, placeholder_text="Eg.10",font=("Arial", 15), width=200)
                Marks_CA2_CO[f"Q{i}_marks"].grid(row=0, column=2, padx=5, pady=5, sticky="w")
                Marks_CA2_CO[f"Q{i}_marks"].configure(state="disabled", fg_color="gray")

            row_base=0
            nptelCA2 = create_label(tab_name, "CO's for NPTEL (CA)", "Arial", 20, row=row_base, column=4,colspan=2, sticky="nsew")
            row_base+=1
            nptelCA2Label = create_label(tab_name, "NPTEL:", "Arial", 15, row=row_base, column=4, sticky="nsw")
            nptelCA2Text = create_entry_box(tab_name, "1,2,3,4,5,6", "Arial", 15, 300, row=row_base, column=5,sticky="nsw")
            nptelCA2Text.configure(state="disabled", fg_color="gray")
            row_base+=2
           
            presentationcA2 = create_label(tab_name, "Maximum group size of Presentations (CA)", "Arial", 20, row=row_base, column=4,colspan=2, sticky="nsew")
            row_base+=1
            presentationCA2Label = create_label(tab_name, "Group Size: ", "Arial", 15, row=row_base, column=4, sticky="nsw")
            presentationCA2Text = create_entry_box(tab_name, "Enter maximum number of students in a group", "Arial", 15, 300, row=row_base, column=5,sticky="nsw")
            presentationCA2Text.configure(state="disabled", fg_color="gray")
            row_base+=2
            
            ALlabelCA2 = create_label(tab_name, "Enter the Target level for CA 2", "Arial", 20, row=row_base, column=4,colspan=2, sticky="nsew")
            row_base+=1
            ALCA2Label = create_label(tab_name, "CA2: ", "Arial", 15, row=row_base, column=4, sticky="nsw")
            ALCA2Text = create_entry_box(tab_name, "Eg.50", "Arial", 15, 300, row=row_base, column=5,sticky="nsw")
            row_base+=1

            button4 = create_button(tab_name, "Next", "Arial", 20, 250, 40, switch_to_CA3, row=row_base, column=5)
            button4.grid(row=row_base, column=4,columnspan=2, rowspan=2, padx=5, pady=5, sticky="")


#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   

            # # COs for CA3 Quiz
            # label15 = create_label(" CA 3 ", "CA3 type :", "Arial", 15, 200, 105)
            # nptelCA3 = create_label(" CA 3 ", "CO's for NPTEL (CA)", "Arial", 20, 470, 555)
            # entry15 = create_dropdown(" CA 3 ", ["Select Type", "Quiz", "NPTEL Course", "Presentation", "Test", "Other"], "Arial", 15, 300, ca3, 500, 105)
            # entry15.configure(state="disabled", fg_color="gray")
            # noCA3Label = create_label(" CA 3 ", "No of Question CA3 (Quiz/Test) :", "Arial", 15, 200, 155)
            # noCA3Entry = create_dropdown(" CA 3 ", ["Select No", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], "Arial", 15, 300, noQuestion3, 500, 155)
            # noCA3Entry.configure(state="disabled", fg_color="gray")
            # nptelCA3Label = create_label(" CA 3 ", "NPTEL: ", "Arial", 15, 350, 605)
            # nptelCA3Text = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 300, 450, 605)
            # nptelCA3Text.configure(state="disabled", fg_color="gray")
            # presentationCA3 = create_label(" CA 3 ", "Maximum group size of Presentations (CA)", "Arial", 20, 870, 555)
            # presentationCA3Label = create_label(" CA 3 ", "Group Size: ", "Arial", 15, 850, 605)
            # presentationCA3Text = create_entry_box(" CA 3 ", "Enter maximum number of students in a group", "Arial", 15, 350, 950, 605)
            # presentationCA3Text.configure(state="disabled", fg_color="gray")

            # label10 = create_label(" CA 3 ", "Is CA3 Applicable: ", "Arial", 15, 200, 55)

            # entry10 = create_dropdown(" CA 3 ", ["Select Yes/No", "Yes", "No"], "Arial", 15, 300, disable, 500, 55)

            # label21 = create_label(" CA 3 ", "COs for CA3 Quiz/Test", "Arial", 20, 350, 255)
            # label21marks = create_label(" CA 3 ", "Marks of Questions for CA3 Quiz/Test", "Arial", 20, 1000, 255)

            # q1LCA3 = create_label(" CA 3 ", "Q1 :", "Arial", 15, 200, 305)
            # q1TCA3 = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 305)
            # q1TCA3.configure(state="disabled", fg_color="gray")

            # q2LCA3 = create_label(" CA 3 ", "Q2 :", "Arial", 15, 200, 355)
            # q2TCA3 = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 355)
            # q2TCA3.configure(state="disabled", fg_color="gray")

            # q3LCA3 = create_label(" CA 3 ", "Q3 :", "Arial", 15, 200, 405)
            # q3TCA3 = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 405)
            # q3TCA3.configure(state="disabled", fg_color="gray")

            # q4LCA3 = create_label(" CA 3 ", "Q4 :", "Arial", 15, 200, 455)
            # q4TCA3 = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 455)
            # q4TCA3.configure(state="disabled", fg_color="gray")

            # q5LCA3 = create_label(" CA 3 ", "Q5 :", "Arial", 15, 200, 505)
            # q5TCA3 = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 250, 505)
            # q5TCA3.configure(state="disabled", fg_color="gray")

            # q6LCA3 = create_label(" CA 3 ", "Q6 :", "Arial", 15, 500, 305)
            # q6TCA3 = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 305)
            # q6TCA3.configure(state="disabled", fg_color="gray")

            # q7LCA3 = create_label(" CA 3 ", "Q7 :", "Arial", 15, 500, 355)
            # q7TCA3 = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 355)
            # q7TCA3.configure(state="disabled", fg_color="gray")

            # q8LCA3 = create_label(" CA 3 ", "Q8 :", "Arial", 15, 500, 405)
            # q8TCA3 = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 405)
            # q8TCA3.configure(state="disabled", fg_color="gray")

            # q9LCA3 = create_label(" CA 3 ", "Q9 :", "Arial", 15, 500, 455)
            # q9TCA3 = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 455)
            # q9TCA3.configure(state="disabled", fg_color="gray")

            # q10LCA3 = create_label(" CA 3 ", "Q10 :", "Arial", 15, 500, 505)
            # q10TCA3 = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 550, 505)
            # q10TCA3.configure(state="disabled", fg_color="gray")
            
            # q1LCA3marks = create_label(" CA 3 ", "Q1 :", "Arial", 15, 900, 305)
            # q1TCA3marks = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 950, 305)
            # q1TCA3marks.configure(state="disabled", fg_color="gray")

            # q2LCA3marks = create_label(" CA 3 ", "Q2 :", "Arial", 15, 900, 355)
            # q2TCA3marks = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 950, 355)
            # q2TCA3marks.configure(state="disabled", fg_color="gray")

            # q3LCA3marks = create_label(" CA 3 ", "Q3 :", "Arial", 15, 900, 405)
            # q3TCA3marks = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 950, 405)
            # q3TCA3marks.configure(state="disabled", fg_color="gray")

            # q4LCA3marks = create_label(" CA 3 ", "Q4 :", "Arial", 15, 900, 455)
            # q4TCA3marks = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 950, 455)
            # q4TCA3marks.configure(state="disabled", fg_color="gray")

            # q5LCA3marks = create_label(" CA 3 ", "Q5 :", "Arial", 15, 900, 505)
            # q5TCA3marks = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 950, 505)
            # q5TCA3marks.configure(state="disabled", fg_color="gray")

            # q6LCA3marks = create_label(" CA 3 ", "Q6 :", "Arial", 15, 1200, 305)
            # q6TCA3marks = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 1250, 305)
            # q6TCA3marks.configure(state="disabled", fg_color="gray")

            # q7LCA3marks = create_label(" CA 3 ", "Q7 :", "Arial", 15, 1200, 355)
            # q7TCA3marks = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 1250, 355)
            # q7TCA3marks.configure(state="disabled", fg_color="gray")

            # q8LCA3marks = create_label(" CA 3 ", "Q8 :", "Arial", 15, 1200, 405)
            # q8TCA3marks = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 1250, 405)
            # q8TCA3marks.configure(state="disabled", fg_color="gray")

            # q9LCA3marks = create_label(" CA 3 ", "Q9 :", "Arial", 15, 1200, 455)
            # q9TCA3marks = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 1250,455)
            # q9TCA3marks.configure(state="disabled", fg_color="gray")

            # q10LCA3marks = create_label(" CA 3 ", "Q10 :", "Arial", 15, 1200, 505)
            # q10TCA3marks = create_entry_box(" CA 3 ", "1,2,3,4,5,6", "Arial", 15, 150, 1250, 505)
            # q10TCA3marks.configure(state="disabled", fg_color="gray")

            # ALlabelCA3 = create_label(" CA 3 ", "Enter the Target level for CA 3", "Arial", 20, 600, 640)
            # ALCA3Label = create_label(" CA 3 ", "CA3: ", "Arial", 15, 450, 675)
            # ALCA3Text = create_entry_box(" CA 3 ", "", "Arial", 15, 500, 550, 675)
            # ALCA3Text.configure(state="disabled", fg_color="gray")

            # button2 = create_button(" CA 3 ", "Next", "Arial", 20, 200, 40, switch_to_template, 1100, 665)
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
            tab_name = " CA 3 "
            
            for i in range(7):
                scroll_frames[tab_name].grid_columnconfigure(i, weight=1)

            row_base=0

            labelCA3= create_label(tab_name, "CA 3 Details", "Arial", 20, row=row_base, column=1, colspan=2)
            labelCA3.configure(anchor="center", justify="center")
            labelCA3.grid_configure(sticky="nsew")
            row_base +=1

            label10 = create_label(tab_name, "Is CA3 Applicable: ", "Arial", 15, row=row_base, column=1, sticky="nsw")
            entry10 = create_dropdown(tab_name, ["Select Yes/No", "Yes", "No"], "Arial", 15, 300, disable, row=row_base, column=2,sticky="nsw")
            row_base+=1

            label15 = create_label(tab_name, "CA3 type :", "Arial", 15, row=row_base, column=1, sticky="nsw")
            entry15 = create_dropdown(tab_name, ["Select Type", "Quiz", "NPTEL Course", "Presentation", "Test",  "Other"], "Arial", 15, 300, ca3, row=row_base, column=2,sticky="nsw")
            entry15.configure(state="disabled", fg_color="gray",button_color="gray")
            row_base+=1
            
            noCA3Label = create_label(tab_name, "No of Question CA3 (Quiz/Test) :", "Arial", 15, row=row_base, column=1, sticky="nsw")
            noCA3Entry = create_dropdown(tab_name, ["Select No", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], "Arial", 15, 300, noQuestion3, row=row_base, column=2,sticky="nsw")
            noCA3Entry.configure(state="disabled", fg_color="gray",button_color="gray")
            row_base+=1

            merged_frame3 = ctk.CTkFrame(scroll_frames[tab_name])
            merged_frame3.grid(row=row_base, column=1, columnspan=2, padx=5, pady=5, sticky="nsew")

            for sub_col in range(3):
                merged_frame3.grid_columnconfigure(sub_col, weight=1)

            label_title_7 = ctk.CTkLabel(merged_frame3, text="No.")
            label_title_7.grid(row=0, column=0, padx=5, pady=5, sticky="nsw")

            label_title_8 = ctk.CTkLabel(merged_frame3, text="COs")
            label_title_8.grid(row=0, column=1, padx=5, pady=5, sticky="nsw")

            label_title_9 = ctk.CTkLabel(merged_frame3, text="Marks")
            label_title_9.grid(row=0, column=2, padx=5, pady=5, sticky="nsw")

            vertical_line_ca3 = ctk.CTkFrame(scroll_frames[tab_name], width=2, fg_color="white")
            vertical_line_ca3.grid(row=1, column=3, rowspan=19, sticky="ns", padx=5)

            Question_CA3_CO={}
            Marks_CA3_CO={}

            for i in range(1,11):
                row_base+=1

                # Create a frame to be placed inside merged col 2+3
                merged_frame = ctk.CTkFrame(scroll_frames[tab_name])
                merged_frame.grid(row=row_base, column=1, columnspan=2, padx=5, pady=5, sticky="nsew")

                # Configure 3 sub-columns inside merged frame
                for sub_col in range(3):
                    merged_frame.grid_columnconfigure(sub_col, weight=1)

                # Label inside merged frame
                if i==10:
                    label = ctk.CTkLabel(merged_frame, text=f"Q.{i}:")
                else :
                    label = ctk.CTkLabel(merged_frame, text=f"Q.{i} : ")
                label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

                # Entry 1
                Question_CA3_CO[f"Q{i}_questions"] = ctk.CTkEntry(merged_frame, placeholder_text="1,2,3,4,5,6",font=("Arial", 15), width=200)
                Question_CA3_CO[f"Q{i}_questions"].grid(row=0, column=1, padx=5, pady=5, sticky="w")
                Question_CA3_CO[f"Q{i}_questions"].configure(state="disabled", fg_color="gray")

                # Entry 2
                Marks_CA3_CO[f"Q{i}_marks"] = ctk.CTkEntry(merged_frame, placeholder_text="Eg.10",font=("Arial", 15), width=200)
                Marks_CA3_CO[f"Q{i}_marks"].grid(row=0, column=2, padx=5, pady=5, sticky="w")
                Marks_CA3_CO[f"Q{i}_marks"].configure(state="disabled", fg_color="gray")

            row_base=0
            nptelCA3 = create_label(tab_name, "CO's for NPTEL (CA)", "Arial", 20, row=row_base, column=4,colspan=2, sticky="nsew")
            row_base+=1
            nptelCA3Label = create_label(tab_name, "NPTEL:", "Arial", 15, row=row_base, column=4, sticky="nsw")
            nptelCA3Text = create_entry_box(tab_name, "1,2,3,4,5,6", "Arial", 15, 300, row=row_base, column=5,sticky="nsw")
            nptelCA3Text.configure(state="disabled", fg_color="gray")
            row_base+=2
           
            presentationcA3 = create_label(tab_name, "Maximum group size of Presentations (CA)", "Arial", 20, row=row_base, column=4,colspan=2, sticky="nsew")
            row_base+=1
            presentationCA3Label = create_label(tab_name, "Group Size: ", "Arial", 15, row=row_base, column=4, sticky="nsw")
            presentationCA3Text = create_entry_box(tab_name, "Enter maximum number of students in a group", "Arial", 15, 300, row=row_base, column=5,sticky="nsw")
            presentationCA3Text.configure(state="disabled", fg_color="gray")
            row_base+=2
            
            ALlabelCA3 = create_label(tab_name, "Enter the Target level for CA 3", "Arial", 20, row=row_base, column=4,colspan=2, sticky="nsew")
            row_base+=1
            ALCA3Label = create_label(tab_name, "CA3: ", "Arial", 15, row=row_base, column=4, sticky="nsw")
            ALCA3Text = create_entry_box(tab_name, "Eg.50", "Arial", 15, 300, row=row_base, column=5,sticky="nsw")
            ALCA3Text.configure(state="disabled", fg_color="gray")
            row_base+=1

            button5 = create_button(tab_name, "Next", "Arial", 20, 250, 40, switch_to_template, row=row_base, column=5)
            button5.grid(row=row_base, column=4,columnspan=2, rowspan=2, padx=5, pady=5, sticky="")
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   

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
                if entry10.get() == "No":
                    al_values=[ALCA1Text.get(), ALCA2Text.get(), '-', ALMidTermText.get(), ALEndSemText.get()]
                    print(al_values)
                else:
                    al_values=[ALCA1Text.get(), ALCA2Text.get(), ALCA3Text.get(), ALMidTermText.get(), ALEndSemText.get()]
                file_path = file_path
                print("File Path : : : ", file_path)
                import Cal
                # Cal.cal_sheet(file_path, al_values)
                Cal.cal_sheet(file_path, emailTextProcessed.get())

            # # Using create_label, create_entry_box, and create_dropdown to recreate the UI


#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   

            # # CO Information
            # enterCO = create_label(" CO Information ", "Enter the CO's Description", "Arial", 20, 700, 50)
            # noOfCOLabel = create_label(" CO Information ", "Select No. of CO's: ", "Arial", 15, 550, 100)
            # noOfCOOption = create_dropdown(" CO Information ", ['Select No of CO\'s', '5', '6'], "Arial", 15, 300, noOfCO, 750, 100)

            # CO1L = create_label(" CO Information ", "CO1: ", "Arial", 15, 550, 150)
            # CO1T = create_entry_box(" CO Information ", "", "Arial", 15, 500, 650, 150)

            # CO2L = create_label(" CO Information ", "CO2: ", "Arial", 15, 550, 200)
            # CO2T = create_entry_box(" CO Information ", "", "Arial", 15, 500, 650, 200)

            # CO3L = create_label(" CO Information ", "CO3: ", "Arial", 15, 550, 250)
            # CO3T = create_entry_box(" CO Information ", "", "Arial", 15, 500, 650, 250)

            # CO4L = create_label(" CO Information ", "CO4: ", "Arial", 15, 550, 300)
            # CO4T = create_entry_box(" CO Information ", "", "Arial", 15, 500, 650, 300)

            # CO5L = create_label(" CO Information ", "CO5: ", "Arial", 15, 550, 350)
            # CO5T = create_entry_box(" CO Information ", "", "Arial", 15, 500, 650, 350)

            # CO6L = create_label(" CO Information ", "CO6: ", "Arial", 15, 550, 400)
            # CO6T = create_entry_box(" CO Information ", "", "Arial", 15, 500, 650, 400)
            # CO6T.configure(state="disabled", fg_color="gray")

            # button1 = create_button(" CO Information ", "Next", "Arial", 20, 200, 40, switch_to_MidTerm_EndSem, 725, 500)
#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   
            tab_name = " CO Information "
            
            for i in range(5):
                scroll_frames[tab_name].grid_columnconfigure(i, weight=1)

            row_base = 0
            enterCO = create_label(tab_name, "Enter the CO's Description", "Arial", 20, row=row_base, column=0, colspan=5)
            enterCO.configure(anchor="center", justify="center")
            enterCO.grid_configure(sticky="nsew")

            row_base +=2

            noOfCOLabel = create_label(tab_name, "Select No. of CO's: ", "Arial", 15, row=row_base, column=2,sticky="nsw")
            noOfCOOption = create_dropdown(tab_name, ['Select No of CO\'s', '5', '6'],  "Arial", 15,300,com=noOfCO, row=row_base, column=3,sticky="nsw")

            row_base +=1

            co_desc_entry={}

            for i in range(1,7) :
                create_label(tab_name, f"CO{i}:", "Arial", 15, row=row_base, column=2, sticky="nsw")
                co_desc_entry[f"CO{i}T"]= create_entry_box(tab_name, f"Enter CO{i} Description", "Arial", 15, 500, row=row_base, column=3,sticky="nsw")
                row_base+=1
            
            co_desc_entry["CO6T"].configure(state="disabled",fg_color="gray")
            row_base+=1

            button1 = create_button(tab_name, "Next", "Arial", 20, 300, 40, switch_to_MidTerm_EndSem, row=row_base, column=2,colspan=2,sticky="")

#<--------------------------------------------------------------------------------------------------------------------------------------------------------------->   

            # # Target level of tests
            # # ALlabel = create_label(" Target level of tests ", "Enter the Target levels for each exam", "Arial", 20, 600, 50)
            # # ALSurveyLabel = create_label(" Target level of tests ", "Survey: ", "Arial", 15, 450, 350)
            # # ALSurveyText = create_entry_box(" Target level of tests ", "", "Arial", 15, 500, 575, 350)



            # setEmailLabel = create_label(" Process Template/Calculated ", "Generate Template", "Arial", 20, 650, 10)
            # setEmailIDLabel = create_label(" Process Template/Calculated ", "Enter the Email ID to send the template sheet.", "Arial", 20, 200, 50)
            # emailText = create_entry_box(" Process Template/Calculated ", "", "Arial", 15, 500, 700, 50)
            # button = create_button(" Process Template/Calculated ", "Download", "Arial", 20, 200, 40, download, 650, 100)


            # # Buttons
            # # button2 = create_button(" CO Mapping ", "Next", "Arial", 20, 200, 40, switch2, 725, 500)
            
            
            # line1 = ctk.CTkFrame(master=tabview.tab(" Process Template/Calculated "), height=2, width=1200, fg_color="white")
            # line1.place(x=150,y=200)

            # path_entry=ctk.CTkEntry(tabview.tab(" Process Template/Calculated "))

            # # button_process=ctk.CTkButton(tabview.tab(" Process Template/Calculated "),text="Process",width=100,height=30,command=process_file)
            # # button_process.place(x=500,y=500)

            # upload_Label = create_label(" Process Template/Calculated ", "Upload you excel file with the marks entered:", "Arial", 25, 550, 250)
            # path_label = create_label(" Process Template/Calculated ", "Path of file", "Arial", 15, 650, 310)
            # button_upload = create_button(" Process Template/Calculated ", "Upload", "Arial", 20, 200, 40, upload_file, 400, 300)


            # line = ctk.CTkFrame(master=tabview.tab(" Process Template/Calculated "), height=2, width=1200, fg_color="white")
            # line.place(x=150,y=400)

            # process_Label = create_label(" Process Template/Calculated ", "Process the excel file you uploaded:", "Arial", 25, 600, 450)

            # setEmailProcessedLabel = create_label(" Process Template/Calculated ", "Enter the Email ID to send the calculated sheet.", "Arial", 20, 200, 500)

            # # important_label = create_label(" Process Template/Calculated ", "Important: Please fill the no. of CO\'s field and the CO\'s in the CO Information page and AL values in Target level of tests page before processing the file", "Arial", 20, 100, 325)
            # # important_label.configure(text_color="black", fg_color="yellow")

            # emailTextProcessed = create_entry_box(" Process Template/Calculated ", "", "Arial", 15, 500, 700, 500)

            # button_process = create_button(" Process Template/Calculated ", "Process", "Arial", 20, 200, 40, process_file, 650, 550)
#<---------------------------------------------------------------------------------------------------------------------------------------------->
            tab_name = " Process Template/Calculated "
            
            for i in range(5):
                scroll_frames[tab_name].grid_columnconfigure(i, weight=1)

            row_base=0
            setEmailLabel = create_label(tab_name, "Generate Template", "Arial", 20, row=row_base, column=0, colspan=5)
            setEmailLabel.configure(anchor="center", justify="center")
            setEmailLabel.grid_configure(sticky="nsew")

            row_base+=1
            setEmailIDLabel = create_label(tab_name, "Enter the Email ID to send the template sheet :", "Arial", 15, row=row_base, column=2,sticky="nsw")
            emailText = create_entry_box(tab_name, "Eg. abc2050@gmail.com", "Arial", 15, 500, row=row_base, column=3,sticky="nsw")
            row_base+=1
            button = create_button(tab_name, "Download", "Arial", 20, 200, 40, download, row=row_base, column=2,colspan=2,sticky="")
            row_base+=1

            # Line separator
            line_separator = ctk.CTkFrame(scroll_frames[tab_name], height=2, fg_color="white")  # Use your desired color
            line_separator.grid(row=row_base, column=0, columnspan=5, sticky="ew", padx=10, pady=10)

            row_base+=1
            upload_Label = create_label(tab_name, "Upload you excel file with the marks entered", "Arial", 20, row=row_base, column=0,colspan=5, sticky="nsew", padx=10, pady=10)
            row_base+=1
            path_label = create_label(tab_name, "Path of file :", "Arial", 15, row=row_base, column=2, sticky="nsw", padx=10, pady=10)
            path_entry= create_entry_box(tab_name, "", "Arial", 15, 500, row=row_base, column=3, sticky="nsw", padx=10)
            row_base+=1
            button_upload = create_button(tab_name, "Upload", "Arial", 20, 200, 40, upload_file, row=row_base, column=2,colspan=2,sticky="")
            row_base+=1

            line_separator1 = ctk.CTkFrame(scroll_frames[tab_name], height=2, fg_color="white")  # Use your desired color
            line_separator1.grid(row=row_base, column=0, columnspan=5, sticky="ew", padx=10, pady=10)
            row_base+=1

            process_Label = create_label(tab_name, "Process the excel file you uploaded", "Arial", 20, row=row_base, column=0,colspan=5, sticky="nsew", padx=10, pady=10)
            row_base+=1
            setEmailProcessedLabel = create_label(tab_name, "Enter the Email ID to send the calculated sheet :", "Arial", 15, row=row_base, column=2,sticky="nsw")
            emailTextProcessed = create_entry_box(tab_name, "", "Arial", 15, 500, row=row_base, column=3, sticky="nsw", padx=10)
            row_base+=1
            button_process = create_button(tab_name, "Process", "Arial", 20, 200, 40, process_file, row=row_base, column=2,colspan=2,sticky="")
#<---------------------------------------------------------------------------------------------------------------------------------------------->
          
#<---------------------------------------------------------------------------------------------------------------------------------------------->
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
                # new_value = event.widget.get()
                if hasattr(event, "widget"):
                    new_value = event.widget.get()
                else:
                    # It's directly a CTkEntry widget
                    new_value = event.get()

                if new_value:
                    # Check basic format: Length should be 9, with a '-' in the middle, and both parts should be digits
                    if len(new_value) != 9 or new_value[4] != '-' or not (new_value[:4].isdigit() and new_value[5:].isdigit()):
                        CTkMessagebox(title="Invalid Input", message="Academic Year format is incorrect. Please enter in the format YYYY-YYYY.", icon="warning")
                        return False

                    # Extract years and validate they are consecutive
                    start_year, end_year = int(new_value[:4]), int(new_value[5:])
                    if end_year - start_year != 1:
                        CTkMessagebox(title="Invalid Input", message="Academic Year should be consecutive (e.g., 2025-2026).", icon="warning")
                        return False

                    return True

                return False
            
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
                # factor_title_entry = [factor_1_title_entry, factor_2_title_entry, factor_3_title_entry, factor_4_title_entry, factor_5_title_entry]
                # factor_lo_entry = [factor_1_lo_entry, factor_2_lo_entry, factor_3_lo_entry, factor_4_lo_entry, factor_5_lo_entry]
                factor_title_entry= [factor_title_entries[f"factor_{i}_title_entry"] for i in range(1, 6)]
                factor_lo_entry= [factor_lo_entries[f"factor_{i}_lo_entry"] for i in range(1, 6)]

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
                    no_of_exp_dropdown.configure(state="normal", fg_color=["F9F9FA", "#1F6AA5"],button_color=["F9F9FA", "#144870"])
                else:
                    no_of_exp_dropdown.configure(state="disabled", fg_color="gray", button_color="gray")
                    for entry in exp_lo_entry.values():
                        entry.configure(state="disabled", fg_color="gray")

            def exp_fields(option):
                if option == "Select no of experiments":
                    for entry in exp_lo_entry.values():
                        entry.configure(state="disabled", fg_color="gray")
                else:
                    option = int(option)
                    # print("Available keys:", exp_lo_entry.keys())

                    for i in range (1, option+1):
                        (exp_lo_entry[f"exp{i}_lo_entry"]).configure(state="normal", fg_color=["F9F9FA", "#343638"])
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
                        
            def validateNumberString(string):
                # print("XXXXXXXXXXXXXXXXXXXXXXXXXXXXX:",string)
                # for char in string:
                #     if char not in "0123456789":
                #         return False
                return string.isdigit()
                    
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
                    return CTkMessagebox(title = "Error", message="Select No of LO\'s on the previous page", icon="cancel")
                if (entry1_lab.get() == "" or yearDropDown_lab.get() == "Select Year" or entry8_lab.get() == "Select Department" or entry2_lab.get() == "Select Sem" or entry3_lab.get() == "Select Subject" or entry4_lab.get() == "" or entry5_lab.get() == "" or entry7_lab.get() == "Select Class"):
                    return CTkMessagebox(title = "Error", message="Please fill all the basic details", icon="cancel")
                
                elif not validate_academic_year(entry4_lab):
                    return CTkMessagebox(title = "Error", message="Please enter a valid academic year in the format YYYY-YYYY", icon="cancel")
                
                elif not validateNumberString(entry1_lab.get()):
                    return CTkMessagebox(title = "Error", message="Please enter a valid number of students", icon="cancel")
                    
                elif no_of_assignments_dropdown.get() == "Select no of Assignments":
                    return CTkMessagebox(title = "Error", message="Please select the number of assignments", icon="cancel")
                    
                elif (oral_marks_target_entry.get() == "" or mini_project_marks_target_entry.get() == "" or assignment_target_entry.get() == "" or group_size_mini_project_entry.get() == ""):
                    return CTkMessagebox(title = "Error", message="Please fill all the marks targets", icon="cancel")
                
                elif not validateNumberString(group_size_mini_project_entry.get()):
                    print(group_size_mini_project_entry.get().isdigit())
                    return CTkMessagebox(title = "Error", message="Enter Valid Group Size for Mini Project", icon="cancel")
                
                elif not validateNumberString(oral_marks_target_entry.get()):
                    print(oral_marks_target_entry.get().isdigit())
                    return CTkMessagebox(title = "Error", message="Enter Valid Target level for orals", icon="cancel")
                
                elif not validateNumberString(mini_project_marks_target_entry.get()):
                    print(mini_project_marks_target_entry.get().isdigit())
                    return CTkMessagebox(title = "Error", message="Enter Valid Target level for Mini Project", icon="cancel")
                
                elif not validateNumberString(assignment_target_entry.get()):
                    print(mini_project_marks_target_entry.get().isdigit())
                    return CTkMessagebox(title = "Error", message="Enter Valid Target level for Assignments", icon="cancel")
                
                option = int(no_of_assignments_dropdown.get())
                for i in range (1, option+1):
                    if assignment_lo_entry[f"assignment_{i}_lo_entry"].get() == "":
                        return CTkMessagebox(title = "Error", message="Please fill all the LO\'s of assignments", icon="cancel")
                    elif not validate_lo_string(assignment_lo_entry[f"assignment_{i}_lo_entry"].get()):
                        return CTkMessagebox(title = "Error", message="Please fill all the LO\'s of assignments in valid format", icon="cancel")
                # elif termWork_dropdown.get() == "Select Type":
                #     CTkMessagebox(title = "Error", message="Please select the type of term work", icon="cancel")
                #     return
                # elif termWork_dropdown.get() == "Group Students":
                #     size = group_size_exp_entry.get()
                #     if size == "":
                #         CTkMessagebox(title = "Error", message="Please fill the group size", icon="cancel")
                #         return
                #     elif (not (size.isdigit()) or int(size) < 0):
                #         CTkMessagebox(title = "Error", message="Please enter a valid group size", icon="cancel")
                #         return
                #     elif (factor_1_title_entry.get() == "" or factor_2_title_entry.get() == "" or factor_3_title_entry.get() == "" or factor_4_title_entry.get() == "" or factor_5_title_entry.get() == ""):
                #         CTkMessagebox(title = "Error", message="Please fill all the factors", icon="cancel")
                #         return
                #     check_text_group_LO = [factor_1_lo_entry.get(), factor_2_lo_entry.get(), factor_3_lo_entry.get(), factor_4_lo_entry.get(), factor_5_lo_entry.get()]
                #     if "" in check_text_group_LO:
                #         CTkMessagebox(title = "Error", message="Please fill all LO\'s for the factors", icon="cancel")
                #         return
                #     for text in check_text_group_LO:
                #         if not validate_lo_string(text):
                #             CTkMessagebox(title = "Error", message="Please fill LO\'s in valid format", icon="cancel")
                #             return
                # elif termWork_dropdown.get() == "Individual Students":
                #     if no_of_exp_dropdown.get() == "Select no of experiments":
                #         CTkMessagebox(title = "Error", message="Please select the number of experiments", icon="cancel")
                #         return
                tabview.set(" LO Mapping ")



            def switch_2_lab():
                tabview.set(" Upload Excel File (Lab) ")

            def download_template_lab():
                if(noOfLOOption.get() == "Select No of LO\'s"):
                    return CTkMessagebox(title = "Error", message="Enter No of LOs in page 1", icon="cancel")

                if (termWork_dropdown.get() == "Select Type"):
                    return CTkMessagebox(title = "Error", message="Select Type of Term Work", icon="cancel")
                
                if (term_work_marks_target_entry.get() == ""):
                    return CTkMessagebox(title = "Error", message="Enter Term Work Marks Target", icon="cancel")
                
                if not validateNumberString(term_work_marks_target_entry.get()):
                    return CTkMessagebox(title = "Error", message="Enter Valid target for Term Work Marks ", icon="cancel")
                
                if(termWork_dropdown.get() == "Group Students"):
                    if (group_size_exp_entry.get() == ""):
                        return CTkMessagebox(title = "Error", message="Enter Group Size for Experiments", icon="cancel")
                    elif not validateNumberString(group_size_exp_entry.get()):
                        return CTkMessagebox(title = "Error", message="Enter Valid Group Size for Experiments", icon="cancel")
                    # if (factor_1_title_entry.get() == "" or factor_2_title_entry.get() == "" or factor_3_title_entry.get() == "" or factor_4_title_entry.get() == "" or factor_5_title_entry.get() == ""):
                    #     return CTkMessagebox(title = "Error", message="Enter all Factors of Experiments", icon="cancel")
                    for i in range(1, 6):
                        if factor_title_entries[f"factor_{i}_title_entry"].get().strip() == "":
                            return CTkMessagebox(
                                title="Error",
                                message="Enter all Factors of Experiments",
                                icon="cancel"
                            )
                    check_text_group_LO = [factor_lo_entries[f"factor_{i}_lo_entry"].get() for i in range(1, 6)]

                    for text in check_text_group_LO:
                        if not validate_lo_string(text):
                            CTkMessagebox(title = "Error", message="Enter all LO\'s of Factors of Experiments", icon="cancel")
                            return
                elif (termWork_dropdown.get() == "Individual Students"):
                    if (no_of_exp_dropdown.get() == "Select no of experiments"):
                        return CTkMessagebox(title = "Error", message="Select No of Experiments", icon="cancel")
                    option = int(no_of_exp_dropdown.get())
                    for i in range (1, option+1):
                        if exp_lo_entry[f"exp{i}_lo_entry"].get() == "":
                            return CTkMessagebox(title = "Error", message="Enter all LO\'s of Experiments", icon="cancel")
                        elif not validate_lo_string(exp_lo_entry[f"exp{i}_lo_entry"].get()):
                            return CTkMessagebox(title = "Error", message="Enter all LO\'s of Experiments", icon="cancel")

                LOcount = int(noOfLOOption.get())

                if(termWork_dropdown.get() == "Individual Students"):
                    exp_no = []
                    for i in range (0,int(no_of_exp_dropdown.get())):
                        exp_no.append(exp_lo_entry[f"exp{i+1}_lo_entry"].get())
                    for text in exp_no:
                        if not (validate_lo_string(text)):
                            CTkMessagebox(title = "Error", message="Enter all LO\'s of Experiments", icon="cancel")
                            return
                # factor_no = []
                # for i in range (0,4):
                #     factor_no.append(mini_project_lo_entry[f"mini_project_factor{i+1}_lo_entry"].get())
                # for text in factor_no:
                #     if not (validate_lo_string(text)):
                #         CTkMessagebox(title = "Error", message="Enter all LO\'s of Factors of Mini Projects", icon="cancel")
                #         return
                assignment_lo = []
                for i in range (0, int(no_of_assignments_dropdown.get())):
                    assignment_lo.append(assignment_lo_entry[f'assignment_{i+1}_lo_entry'].get())
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
                # critList = [factor_1_title_entry.get(), factor_2_title_entry.get(), factor_3_title_entry.get(), factor_4_title_entry.get(), factor_5_title_entry.get()]
                # loList = [factor_1_lo_entry.get(), factor_2_lo_entry.get(), factor_3_lo_entry.get(), factor_4_lo_entry.get(), factor_5_lo_entry.get()]
                critList = [factor_title_entries[f"factor_{i}_title_entry"].get() for i in range(1, 6)]
                loList = [factor_lo_entries[f"factor_{i}_lo_entry"].get() for i in range(1, 6)]

                basic_values_lo.append(critList)
                basic_values_lo.append(loList)
                basic_values_lo.append(group_size_mini_project_entry.get())
                if noOfLOOption.get() == "5":
                    projLoList = "1,2,3,4,5"
                else:
                    projLoList = "1,2,3,4,5,6"
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
                    entry2_lab.configure(values=["Select Sem"])
                    entry7_lab.configure(values=["Select Class"])
                elif option == "F.E":
                    entry2_lab.configure(values=["Select Sem","I","II"])
                    entry7_lab.configure(values=["Select Class", "D5A", "D5B", "D5C"])
                elif option == "S.E":
                    entry2_lab.configure(values=["Select Sem","III","IV"])
                    entry7_lab.configure(values=["Select Class", "D10A", "D10B", "D10C"])
                elif option == "T.E":
                    entry2_lab.configure(values=["Select Sem","V","VI"])
                    entry7_lab.configure(values=["Select Class", "D15A", "D15B", "D15C"])
                elif option == "B.E":
                    entry2_lab.configure(values=["Select Sem","VII","VIII"])
                    entry7_lab.configure(values=["Select Class", "D20A", "D20B", "D20C"])
            
            # def create_button(tab, name, font_name, font_size, w, h, com, x, y):
            #     button = ctk.CTkButton(master=tabview.tab(tab), text=name, width=w, height=h, font=(font_name, font_size), command=com)
            #     button.place(x=x, y=y)
            #     return button

            # def create_label(tab, name, font_type, font_size, x, y):
            #     label = ctk.CTkLabel(master=tabview.tab(tab), text=name, font=(font_type, font_size))
            #     label.place(x=x, y=y)
            #     return label

            # def create_entry_box(tab, text, font_name, font_size, w, x, y):
            #     entry_box = ctk.CTkEntry(master=tabview.tab(tab), placeholder_text=text, font=(font_name,font_size), width=w)
            #     entry_box.place(x=x,y=y)
            #     return entry_box

            # def create_dropdown(tab, val, font_name, font_size, w, com, x, y):
            #     dropdown = ctk.CTkOptionMenu(master=tabview.tab(tab), values=val, font=(font_name, font_size), width=w, command=com)
            #     dropdown.place(x=x, y=y)
            #     return dropdown
            
            # ------------------ Helper UI Functions (Responsive & Scrollable) ------------------

            def create_label(tab, name, font_type, font_size, row=None, column=None, colspan=1, sticky="nsw", padx=5, pady=15):
                label = ctk.CTkLabel(master=scroll_frames[tab], text=name, font=(font_type, font_size))
                if row is not None and column is not None:
                    label.grid(row=row, column=column, columnspan=colspan, sticky=sticky, padx=padx, pady=pady)
                else:
                    label.pack(pady=15)
                return label

            def create_entry_box(tab, text, font_name, font_size, w, row=None, column=None, colspan=1, sticky="nsw", padx=5, pady=15):
                entry_box = ctk.CTkEntry(master=scroll_frames[tab], placeholder_text=text, font=(font_name, font_size), width=w)
                if row is not None and column is not None:
                    entry_box.grid(row=row, column=column, columnspan=colspan, sticky=sticky, padx=padx, pady=pady)
                else:
                    entry_box.pack(pady=15)
                return entry_box

            def create_button(tab, name, font_name, font_size, w, h=40, com=None, row=None, column=None, colspan=1, sticky="nsw", padx=5, pady=20):
                button = ctk.CTkButton(master=scroll_frames[tab], text=name, width=w, height=h, font=(font_name, font_size), command=com)
                if row is not None and column is not None:
                    button.grid(row=row, column=column, columnspan=colspan, sticky=sticky, padx=padx, pady=pady)
                else:
                    button.pack(pady=20)
                return button

            def create_dropdown(tab, val, font_name, font_size, w, com=None, row=None, column=None, colspan=1, sticky="nsw", padx=5, pady=15):
                dropdown = ctk.CTkOptionMenu(master=scroll_frames[tab], values=val, font=(font_name, font_size), width=w, command=com)
                if row is not None and column is not None:
                    dropdown.grid(row=row, column=column, columnspan=colspan, sticky=sticky, padx=padx, pady=pady)
                else:
                    dropdown.pack(pady=15)
                return dropdown

            
            self.app.destroy() 
            lo_window = ctk.CTk()  # Close the current window 

            screen_width=lo_window.winfo_screenwidth()
            screen_height=lo_window.winfo_screenheight()
       
            # Set window size (like 80% of screen)
            window_width = int(screen_width * 0.8)
            window_height = int(screen_height * 0.8)
            # Center the window
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2

            # Create a new CO Calculations window
            lo_window.title("LO Calculations")
            lo_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

            main_frame = ctk.CTkFrame(master=lo_window)
            main_frame.pack(expand=True, fill="both", padx=10, pady=10)
            
            # ---------- Back Button in topbar ----------
            topbar = ctk.CTkFrame(master=main_frame, fg_color="transparent")
            topbar.pack(side="top", fill="x", padx=0, pady=(0, 0))
            back_button = ctk.CTkButton(
                master=topbar,
                text="← Back",
                width=200,
                command=lambda: self.go_back(lo_window)
            )
            back_button.pack(side="top", anchor="ne", padx=5)
            # Tabview inside the frame
            tabview = ctk.CTkTabview(main_frame, corner_radius=20)
            tabview.pack(expand=True, fill="both", padx=10, pady=5)

            # tabview.add(" LO Information ")
            # tabview.add(" LO Information Template generation ")
            # tabview.add(" LO Mapping ")
            # tabview.add(" Upload Excel File (Lab) ")   
            

            # Add Tabs
            tab_names = [
                " LO Information ",
                " LO Information Template generation ",
                " LO Mapping ",
                " Upload Excel File (Lab) "
            ]

            scroll_frames = {}  # Dictionary to store scrollable frames by tab name

            for tab_name in tab_names:
                tabview.add(tab_name)

                # Create a scrollable frame inside each tab
                scroll_frame = ctk.CTkScrollableFrame(master=tabview.tab(tab_name), label_text="")
                scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
                scroll_frames[tab_name] = scroll_frame


            # label0_lab = create_label(" LO Information Template generation ", "Basic Details", "Arial", 20, 325, 5)

            # label1_lab = create_label(" LO Information Template generation ", "No. of Students :", "Arial", 15, 100, 55)
            # entry1_lab = create_entry_box(" LO Information Template generation ", "Enter no of students", "Arial", 15, 300, 400, 55)

            # newLabel_lab = create_label(" LO Information Template generation ", "Year :", "Arial", 15, 100, 155)
            # yearDropDown_lab = create_dropdown(" LO Information Template generation ", ["Select Year", "F.E", "S.E", "T.E", "B.E"], "Arial", 15, 300, semesterAndClass, 400, 155)

            # label8_lab = create_label(" LO Information Template generation ", "Department :", "Arial", 15, 100, 105)
            # entry8_lab = create_dropdown(" LO Information Template generation ", ["Select Department", "Humanities and Applied Science(FE)", "Information Technology", "Computer", "AI and Data Science", "Electronics and Telecommunication", "Electronics", "Instrumentation"], "Arial", 15, 300, None, 400, 105)

            # label2_lab = create_label(" LO Information Template generation ", "Semester :", "Arial", 15, 100, 205)
            # entry2_lab = create_dropdown(" LO Information Template generation ", ["Select Sem"], "Arial", 15, 300, subject, 400, 205)

            # label3_lab = create_label(" LO Information Template generation ", "Subject :", "Arial", 15, 100, 255)
            # entry3_lab = create_dropdown(" LO Information Template generation ", ["Select Subject"], "Arial", 15, 300, None, 400, 255)

            # label4_lab = create_label(" LO Information Template generation ", "Academic Year: ", "Arial", 15, 100, 305)
            # entry4_lab = create_entry_box(" LO Information Template generation ", "YYYY-YYYY", "Arial", 15, 300, 400, 305)
            # entry4_lab.bind("<FocusOut>", validate_academic_year)

            # label5_lab = create_label(" LO Information Template generation ", "Subject Teacher :", "Arial", 15, 100, 355)
            # entry5_lab = create_entry_box(" LO Information Template generation ", "Subject Teacher", "Arial", 15, 300, 400, 355)

            # label7_lab = create_label(" LO Information Template generation ", "Class :", "Arial", 15, 100, 405)
            # # entry7_lab = create_dropdown(" LO Information Template generation ", "Eg.D10 C", "Arial", 15, 300, 400, 405)

            # entry7_lab = create_dropdown(" LO Information Template generation ", ["Select Class"], "Arial", 15, 300, semesterAndClass, 400, 405)
            
            # assignment_head_lab = create_label(" LO Information Template generation ", "Assignment Details", "Arial", 20, 950, 225)
            # no_of_assignments_label = create_label(" LO Information Template generation ", "Enter no. of Assignments: ", "Arial", 15, 825, 275)
            # no_of_assignments_dropdown = create_dropdown(" LO Information Template generation ", ["Select no of Assignments", "2", "3"], "Arial", 15, 300, assignment_fields, 1025, 275)
            # assignment_target_label = create_label(" LO Information Template generation ", "Target for Assignment: ", "Arial", 15, 825, 325)
            # assignment_target_entry = create_entry_box(" LO Information Template generation ", "", "Arial", 15, 300, 1025, 325)
            # assignment_lo_label = {}
            # assignment_lo_entry = {}
            # for i in range(1,4):
            #     assignment_lo_label[f"assignment_{i}_lo_label"] = create_label(" LO Information Template generation ", f"LO for Assignment {i}: ", "Arial", 15, 825, 325+i*50)
            #     assignment_lo_entry[f"assignment_{i}_lo_entry"] = create_entry_box(" LO Information Template generation ", "1,2,3,4,5,6", "Arial", 15, 175, 1025, 325+i*50)

            # label_10_lab = create_label(" LO Information Template generation ", "Enter Oral Target and MiniProject Details ", "Arial", 20, 890, 5)
            # oral_marks_target_label = create_label(" LO Information Template generation ", "Oral: ", "Arial", 15, 825, 55)
            # oral_marks_target_entry = create_entry_box(" LO Information Template generation ", "", "Arial", 15, 300, 1025, 55)
            # mini_project_marks_target_label = create_label(" LO Information Template generation ", "Mini Project: ", "Arial", 15, 825, 105)
            # mini_project_marks_target_entry = create_entry_box(" LO Information Template generation ", "", "Arial", 15, 300, 1025, 105)
            # group_size_mini_project_label = create_label(" LO Information Template generation ", "No. of students in a group:\n(MiniProject) ", "Arial", 15, 825, 155)
            # group_size_mini_project_entry = create_entry_box(" LO Information Template generation ", "", "Arial", 15, 300, 1025, 155)

            # next_1_lab_button = create_button(" LO Information Template generation ", "Next", "Arial", 20, 250, 40, switch_1_lab, 1020, 600)
            
            tab_name = " LO Information Template generation "
            
            for i in range(6):
                scroll_frames[tab_name].grid_columnconfigure(i, weight=1)

            row_base = 0
            label0_lab = create_label(tab_name, "Basic Details", "Arial", 20, row=row_base, column=1, colspan=2)
            label0_lab.configure(anchor="center", justify="center")
            label0_lab.grid_configure(sticky="nsew")

            # empty_lab = create_label(tab_name, " ", "Arial", 20, row=row_base+1, column=1, colspan=2)
            row_base +=2

            label1_lab = create_label(tab_name, "No. of Students :", "Arial", 15, row=row_base, column=1)
            entry1_lab = create_entry_box(tab_name, "Enter no of students", "Arial", 15, 300, row=row_base, column=2)
            row_base +=1
            
            label8_lab = create_label(tab_name, "Department :", "Arial", 15, row=row_base, column=1)
            entry8_lab = create_dropdown(tab_name, ["Select Department", "Humanities and Applied Science(FE)", "Information Technology", "Computer", "AI and Data Science", "Electronics and Telecommunication", "Electronics", "Instrumentation"], "Arial", 15, 300, None, row=row_base, column=2)
            row_base +=1
            
            newLabel_lab = create_label(tab_name, "Year :", "Arial", 15, row=row_base, column=1)
            yearDropDown_lab = create_dropdown(tab_name, ["Select Year", "F.E", "S.E", "T.E", "B.E"], "Arial", 15, 300, semesterAndClass, row=row_base, column=2)
            row_base +=1
            
            label2_lab = create_label(tab_name, "Semester :", "Arial", 15, row=row_base, column=1)
            entry2_lab = create_dropdown(tab_name, ["Select Sem"], "Arial", 15, 300, subject, row=row_base, column=2)
            row_base +=1
            
            label3_lab = create_label(tab_name, "Subject :", "Arial", 15, row=row_base, column=1)
            entry3_lab = create_dropdown(tab_name, ["Select Subject"], "Arial", 15, 300, None, row=row_base, column=2)
            row_base +=1
            
            label4_lab = create_label(tab_name, "Academic Year: ", "Arial", 15, row=row_base, column=1)
            entry4_lab = create_entry_box(tab_name, "YYYY-YYYY", "Arial", 15, 300, row=row_base, column=2)
            entry4_lab.bind("<FocusOut>", validate_academic_year)
            row_base +=1
            
            label5_lab = create_label(tab_name, "Subject Teacher :", "Arial", 15, row=row_base, column=1)
            entry5_lab = create_entry_box(tab_name, "Subject Teacher", "Arial", 15, 300, row=row_base, column=2)
            row_base +=1
            
            label7_lab = create_label(tab_name, "Class :", "Arial", 15, row=row_base, column=1)
            entry7_lab = create_dropdown(tab_name, ["Select Class"], "Arial", 15, 300, semesterAndClass, row=row_base, column=2)

            
            # Start placing right section from row 0, column 4
            row_base = 0
            tab_name = " LO Information Template generation "

            assignment_head_lab = create_label(tab_name, "Assignment Details", "Arial", 20, row=row_base, column=4, colspan=2)
            assignment_head_lab.configure(anchor="center", justify="center")
            assignment_head_lab.grid_configure(sticky="nsew")   #center in grid

            row_base += 2
            no_of_assignments_label = create_label(tab_name, "Enter no. of Assignments: ", "Arial", 15, row=row_base, column=4)
            no_of_assignments_dropdown = create_dropdown(tab_name, ["Select no of Assignments", "2", "3"], "Arial", 15, 300, assignment_fields, row=row_base, column=5)

            row_base += 1
            assignment_target_label = create_label(tab_name, "Target for Assignment: ", "Arial", 15, row=row_base, column=4)
            assignment_target_entry = create_entry_box(tab_name, "", "Arial", 15, 300, row=row_base, column=5)

            # Assignment LO Entries
            assignment_lo_label = {}
            assignment_lo_entry = {}
            for i in range(1, 4):
                row_base += 1
                assignment_lo_label[f"assignment_{i}_lo_label"] = create_label(
                    tab_name, f"LO for Assignment {i}: ", "Arial", 15, row=row_base, column=4
                )
                assignment_lo_entry[f"assignment_{i}_lo_entry"] = create_entry_box(
                    tab_name, "1,2,3,4,5,6", "Arial", 15, 300, row=row_base, column=5
                )

            # Oral & Mini Project Section
            row_base +=1
            label_10_lab = create_label(tab_name, "Enter Oral Target and MiniProject Details ", "Arial", 20, row=row_base, column=4, colspan=2)
            label_10_lab.configure(anchor="center", justify="center")
            label_10_lab.grid_configure(sticky="nsew")

            row_base += 1
            oral_marks_target_label = create_label(tab_name, "Oral: ", "Arial", 15, row=row_base, column=4)
            oral_marks_target_entry = create_entry_box(tab_name, "", "Arial", 15, 300, row=row_base, column=5)

            row_base += 1
            mini_project_marks_target_label = create_label(tab_name, "Mini Project: ", "Arial", 15, row=row_base, column=4)
            mini_project_marks_target_entry = create_entry_box(tab_name, "", "Arial", 15, 300, row=row_base, column=5)

            row_base += 1
            group_size_mini_project_label = create_label(tab_name, "No. of students in a group:\n(MiniProject) ", "Arial", 15, row=row_base, column=4)
            group_size_mini_project_entry = create_entry_box(tab_name, "", "Arial", 15, 300, row=row_base, column=5)

            # Next Button
            next_1_lab_button = create_button(tab_name, "Next", "Arial", 20, 250, 40, switch_1_lab, row=20, column=5)

            vertical_line1 = ctk.CTkFrame(scroll_frames[tab_name], width=2, fg_color="white")
            vertical_line1.grid(row=2, column=3, rowspan=19, sticky="ns", padx=5)











            # # LO Mapping tab view
            # label9_lab = create_label(" LO Mapping ", "Type of Term Work : ", "Arial", 15, 100, 5)
            # termWork_dropdown = create_dropdown(" LO Mapping ", ["Select Type", "Group Students", "Individual Students"], "Arial", 15, 300, exp_group, 400, 5)

            # label11_lab= create_label(" LO Mapping ", " Grouped Experiments Detail ", "Arial", 20, 195, 125)

            # group_size_exp_label = create_label(" LO Mapping ", "Enter max no of students in group: ", "Arial", 15, 100, 175)
            # group_size_exp_entry = create_entry_box(" LO Mapping ", "", "Arial", 15, 300, 350, 175)
            # group_size_exp_entry.configure(state="disabled",fg_color="gray")
            
            # label_11_lab = create_label(" LO Mapping ", "Factor No. ", "Arial", 15, 100, 225)
            # label_12_lab = create_label(" LO Mapping ", "Factor Title ", "Arial", 15, 300, 225)
            # label_13_lab = create_label(" LO Mapping ", "Corresponding LO ", "Arial", 15, 500, 225)

            # factor_1_label = create_label(" LO Mapping ", "1", "Arial", 15, 100, 275)
            # factor_2_label = create_label(" LO Mapping ", "2", "Arial", 15, 100, 325)
            # factor_3_label = create_label(" LO Mapping ", "3", "Arial", 15, 100, 375)
            # factor_4_label = create_label(" LO Mapping ", "4", "Arial", 15, 100, 425)
            # factor_5_label = create_label(" LO Mapping ", "5", "Arial", 15, 100, 475)

            # factor_1_title_entry = create_entry_box(" LO Mapping ", "Enter title", "Arial", 15, 150, 300, 275)
            # factor_1_title_entry.configure(state="disabled",fg_color="gray")

            # factor_2_title_entry = create_entry_box(" LO Mapping ", "Enter title", "Arial", 15, 150, 300, 325)
            # factor_2_title_entry.configure(state="disabled",fg_color="gray")

            # factor_3_title_entry = create_entry_box(" LO Mapping ", "Enter title", "Arial", 15, 150, 300, 375)
            # factor_3_title_entry.configure(state="disabled",fg_color="gray")

            # factor_4_title_entry = create_entry_box(" LO Mapping ", "Enter title", "Arial", 15, 150, 300, 425)
            # factor_4_title_entry.configure(state="disabled",fg_color="gray")    

            # factor_5_title_entry = create_entry_box(" LO Mapping ", "Enter title", "Arial", 15, 150, 300, 475)
            # factor_5_title_entry.configure(state="disabled",fg_color="gray")

            # factor_1_lo_entry = create_entry_box(" LO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 500, 275)
            # factor_1_lo_entry.configure(state="disabled",fg_color="gray")

            # factor_2_lo_entry = create_entry_box(" LO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 500, 325)
            # factor_2_lo_entry.configure(state="disabled",fg_color="gray")

            # factor_3_lo_entry = create_entry_box(" LO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 500, 375)
            # factor_3_lo_entry.configure(state="disabled",fg_color="gray")

            # factor_4_lo_entry = create_entry_box(" LO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 500, 425)
            # factor_4_lo_entry.configure(state="disabled",fg_color="gray")

            # factor_5_lo_entry = create_entry_box(" LO Mapping ", "1,2,3,4,5,6", "Arial", 15, 150, 500, 475)
            # factor_5_lo_entry.configure(state="disabled",fg_color="gray")

            # #######
            

            # label_14_lab = create_label(" LO Mapping ", " Experiments Detail ", "Arial", 20, 1000, 55)
            
            # label_15_lab = create_label(" LO Mapping ", "Enter the LO's mapped for each experiment ", "Arial", 15, 955, 155)
            
            # no_of_exp_label = create_label(" LO Mapping ", "Enter no. of experiments: ", "Arial", 15, 825, 105)
            # no_of_exp_dropdown = create_dropdown(" LO Mapping ", ["Select no of experiments","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15"], "Arial", 15, 300, exp_fields, 1025, 105)
            # no_of_exp_dropdown.configure(state="disabled",fg_color="gray")

            # term_work_marks_target_label = create_label(" LO Mapping ", "Target for Term Work: ", "Arial", 15, 100, 55)
            # term_work_marks_target_entry = create_entry_box(" LO Mapping ", "", "Arial", 15, 300, 400, 55)

            # # exp_lo_label = create_label(" LO Mapping ", "LO for Experiments ", "Arial", 15, 1000, 155)

            # exp_lo_labels = {}
            # exp_lo_entry = {}

            # for i in range(1,9):
            #     exp_lo_labels[f"exp{i}_lo_label"] = create_label(" LO Mapping ", f"{i}:", "Arial", 15, 825, 205+(50*(i-1)))
            #     exp_lo_entry[f"exp{i}_lo_entry"] = create_entry_box(" LO Mapping ", "1,2,3,4,5,6", "Arial", 15, 100, 925, 205+(50*(i-1)))

            # c = 1
            # for i in range(9,16):
            #     exp_lo_labels[f"exp{i}_lo_label"] = create_label(" LO Mapping ", f"{i}:", "Arial", 15, 1050, 205+(50*(c-1)))
            #     exp_lo_entry[f"exp{i}_lo_entry"] = create_entry_box(" LO Mapping ", "1,2,3,4,5,6", "Arial", 15, 100, 1150, 205+(50*(c-1)))
            #     c = c+ 1

            # for entry in exp_lo_entry.values():
            #     entry.configure(state="disabled", fg_color="gray")

            # tab_name = " LO Mapping "
            # row_base = 0

            # # Column spacing for responsive layout
            # for i in range(7):  # Columns 0 to 6
            #     scroll_frames[tab_name].grid_columnconfigure(i, weight=1)

            # # Row 0: Type of Term Work
            # label9_lab = create_label(tab_name, "Type of Term Work:", "Arial", 15, row=row_base, column=1)
            # termWork_dropdown = create_dropdown(tab_name, ["Select Type", "Group Students", "Individual Students"], "Arial", 15, 300, exp_group, row=row_base, column=2)
            # row_base += 1

            # # Row 1: Term Work Target
            # term_work_marks_target_label = create_label(tab_name, "Target for Term Work:", "Arial", 15, row=row_base, column=1)
            # term_work_marks_target_entry = create_entry_box(tab_name, "", "Arial", 15, 300, row=row_base, column=2)
            # row_base += 1

            # # Row 2: Section Heading
            # label11_lab = create_label(tab_name, "Grouped Experiments Detail", "Arial", 20, row=row_base, column=1, colspan=3, sticky="nsew")
            # row_base += 1

            # # Row 3: Group Size Input
            # group_size_exp_label = create_label(tab_name, "Enter max no of students in group:", "Arial", 15, row=row_base, column=1)
            # group_size_exp_entry = create_entry_box(tab_name, "", "Arial", 15, 300, row=row_base, column=2)
            # group_size_exp_entry.configure(state="disabled", fg_color="gray")
            # row_base += 1

            # # Row 4: Factor Table Headers
            # label_11_lab = create_label(tab_name, "Factor No.", "Arial", 15, row=row_base, column=1)
            # label_12_lab = create_label(tab_name, "Factor Title", "Arial", 15, row=row_base, column=2)
            # label_13_lab = create_label(tab_name, "Corresponding LO", "Arial", 15, row=row_base, column=3)
            # row_base += 1

            # # Rows 5–9: Factor Rows
            # for i in range(1, 6):
            #     create_label(tab_name, f"{i}", "Arial", 15, row=row_base, column=1)

            #     title_entry = create_entry_box(tab_name, "Enter title", "Arial", 15, 150, row=row_base, column=2)
            #     title_entry.configure(state="disabled", fg_color="gray")

            #     lo_entry = create_entry_box(tab_name, "1,2,3,4,5,6", "Arial", 15, 150, row=row_base, column=3)
            #     lo_entry.configure(state="disabled", fg_color="gray")

            #     row_base += 1

            # # Header for Experiment Details
            # row_base=0
            # label_14_lab = create_label(tab_name, "Experiments Detail", "Arial", 20, row=row_base, column=5, colspan=3, sticky="nsew")
            # row_base += 1

            # # Row: No. of Experiments
            # no_of_exp_label = create_label(tab_name, "Enter no. of experiments:", "Arial", 15, row=row_base, column=5)
            # no_of_exp_dropdown = create_dropdown(tab_name, ["Select no of experiments"] + [str(i) for i in range(1, 16)], "Arial", 15, 300, exp_fields, row=row_base, column=6)
            # no_of_exp_dropdown.configure(state="disabled", fg_color="gray")
            # row_base += 1

            # # Row: LO Mapping instruction
            # label_15_lab = create_label(tab_name, "Enter the LO's mapped for each experiment", "Arial", 15, row=row_base, column=5, colspan=2, sticky="nsew")
            # row_base += 1

            # # Rows: Experiment LO Mapping 
            # exp_lo_labels = {}
            # exp_lo_entry = {}

            # for i in range(1, 16):
            #     exp_lo_labels[f"exp{i}_lo_label"] = create_label(tab_name, f"{i}:", "Arial", 15, row=row_base, column=5)
            #     exp_lo_entry[f"exp{i}_lo_entry"] = create_entry_box(tab_name, "1,2,3,4,5,6", "Arial", 15, 100, row=row_base, column=6)
            #     exp_lo_entry[f"exp{i}_lo_entry"].configure(state="disabled", fg_color="gray")
            #     row_base += 1






            tab_name = " LO Mapping "
            row_base = 0

            # Column spacing for responsive layout (0–6)
            for i in range(7):
                scroll_frames[tab_name].grid_columnconfigure(i, weight=1)

            # Ensure equal width for columns 4 and 5
            # scroll_frames[tab_name].grid_columnconfigure(4, weight=1, uniform="col")
            # scroll_frames[tab_name].grid_columnconfigure(5, weight=1, uniform="col")

            ### === FACTOR SECTION (col 1,2) === ###
            row_base=0
            create_label(tab_name, "Term Work Detail", "Arial", 20, row=row_base, column=1, colspan=2, sticky="nsew")
            row_base += 1

            # Row 0: Type of Term Work
            create_label(tab_name, "Type of Term Work:", "Arial", 15, row=row_base, column=1, sticky="nsw")
            termWork_dropdown = create_dropdown(tab_name, ["Select Type", "Group Students", "Individual Students"], "Arial", 15, 300, exp_group, row=row_base, column=2)
            row_base += 1

            # Row 1: Term Work Target
            create_label(tab_name, "Target for Term Work:", "Arial", 15, row=row_base, column=1, sticky="nsw")
            term_work_marks_target_entry = create_entry_box(tab_name, "", "Arial", 15, 300, row=row_base, column=2)
            row_base += 1

            # Row 2: Section Heading
            create_label(tab_name, "Grouped Experiments Detail", "Arial", 20, row=row_base, column=1, colspan=2, sticky="nsew")
            row_base += 1

            # Row 3: Group Size
            create_label(tab_name, "Enter max no of students\nin group:", "Arial", 15, row=row_base, column=1, sticky="nsw")
            group_size_exp_entry = create_entry_box(tab_name, "", "Arial", 15, 300, row=row_base, column=2)
            group_size_exp_entry.configure(state="disabled", fg_color="gray")
            row_base += 1

           # Row 4: Title Label for the merged column area
            # Header Row for Factor Section (row 4)
            header_inner_frame = ctk.CTkFrame(scroll_frames[tab_name], fg_color="transparent")
            header_inner_frame.grid(row=row_base, column=1, columnspan=2, padx=5, pady=(10, 5), sticky="nsew")

            # Configure 3 columns inside header frame
            header_inner_frame.grid_columnconfigure(0, weight=1, minsize=30)   # For "Factor No."
            header_inner_frame.grid_columnconfigure(1, weight=2)               # For Title
            header_inner_frame.grid_columnconfigure(2, weight=2)               # For LO

            # Add labels inside the header inner frame
            ctk.CTkLabel(header_inner_frame, text="Factor No.", font=("Arial", 15)).grid(row=0, column=0, padx=5,pady=10, sticky="nsw")
            ctk.CTkLabel(header_inner_frame, text="Title", font=("Arial", 15)).grid(row=0, column=1, padx=5,pady=10, sticky="nsw")
            ctk.CTkLabel(header_inner_frame, text="LO", font=("Arial", 15)).grid(row=0, column=2, padx=5,pady=10, sticky="nsw")

            row_base += 1  # Move to next row for factors

            factor_title_entries = {}
            factor_lo_entries = {}


            # Rows 5–9: Factor Details
            for i in range(1, 6):
                # Inner frame to combine Factor No., Title, and LO in one merged cell
                factor_inner_frame = ctk.CTkFrame(scroll_frames[tab_name], fg_color="transparent")
                factor_inner_frame.grid(row=row_base, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

                # Configure 3 columns inside the inner frame
                factor_inner_frame.grid_columnconfigure(0, weight=1, minsize=30)   # Factor No.
                factor_inner_frame.grid_columnconfigure(1, weight=2)               # Title
                factor_inner_frame.grid_columnconfigure(2, weight=2)               # LO

                # Factor No. Label
                ctk.CTkLabel(factor_inner_frame, text=f"{i}:", font=("Arial", 15)).grid(row=0, column=0, padx=5,pady=10, sticky="w")

                # Title Entry
                title_entry = ctk.CTkEntry(factor_inner_frame, placeholder_text="Enter title", font=("Arial", 15), width=150)
                title_entry.grid(row=0, column=1, padx=(5, 20),pady=10, sticky="ew")  # Added right-padding for space between title and LO
                title_entry.configure(state="disabled", fg_color="gray")

                # LO Entry
                lo_entry = ctk.CTkEntry(factor_inner_frame, placeholder_text="1,2,3,4,5,6", font=("Arial", 15), width=150)
                lo_entry.grid(row=0, column=2, padx=5,pady=10, sticky="ew")
                lo_entry.configure(state="disabled", fg_color="gray")

                factor_title_entries[f"factor_{i}_title_entry"] = title_entry
                factor_lo_entries[f"factor_{i}_lo_entry"] = lo_entry

                row_base += 1



            ### === EXPERIMENT SECTION (col 4,5) === ###
            row_base = 0
             
            # Row 0: Section Heading
            create_label(tab_name, "Experiments Detail", "Arial", 20, row=row_base, column=4, colspan=2, sticky="nsew")
            row_base += 1

            # create_label(tab_name, " ", "Arial", 20, row=row_base, column=4, colspan=2, sticky="nsew")
            # row_base += 1

            # Row 1: No. of Experiments
            create_label(tab_name, "Enter no. of experiments:", "Arial", 15, row=row_base, column=4, sticky="nsw")
            no_of_exp_dropdown = create_dropdown(tab_name, ["Select no of experiments"] + [str(i) for i in range(1, 16)], "Arial", 15, 300, exp_fields, row=row_base, column=5)
            no_of_exp_dropdown.configure(state="disabled", fg_color="gray", button_color="gray")
            row_base += 1

            # Row 2: Instructions
            create_label(tab_name, "Enter the LO's mapped for each experiment", "Arial", 20, row=row_base, column=4, colspan=2, sticky="nsew")
            row_base += 1

            # Headings for Exp No + LO mapping (col 4 and col 5)
            # exp_col_titles = [("Exp No.", "LO")] * 2
            # exp_col_indices = [4, 5]

            # for idx, col in enumerate(exp_col_indices):
            #     exp_title_frame = ctk.CTkFrame(scroll_frames[tab_name], fg_color="transparent")
            #     exp_title_frame.grid(row=row_base, column=col, padx=50, pady=5, sticky="nsew")
            #     exp_title_frame.grid_columnconfigure(0, weight=0, minsize=50)  # Exp No.
            #     exp_title_frame.grid_columnconfigure(1, weight=1)              # LO Entry

            #     ctk.CTkLabel(exp_title_frame, text="Exp No.", font=("Arial", 15)).grid(row=0, column=0, sticky="nsw", padx=5)
            #     ctk.CTkLabel(exp_title_frame, text="LO", font=("Arial", 15)).grid(row=0, column=1, sticky="nsew", padx=5)

            # row_base += 1

            # Entries 1-8 go in column 4, 9-15 in column 5
            exp_lo_labels = {}
            exp_lo_entry = {}

            row_base_copy = row_base  # Preserve original

            left_row = row_base_copy
            right_row = row_base_copy
            k=1
            for i in range(1, 18):
                if i <= 9:
                    col = 4
                    row = left_row
                    left_row += 1
                    padX=10
                else:
                    col = 5
                    row = right_row
                    right_row += 1
                    padX=10

                inner_frame = ctk.CTkFrame(scroll_frames[tab_name], fg_color="transparent")
                inner_frame.grid(row=row, column=col, padx=padX, pady=10, sticky="ew")

                inner_frame.grid_columnconfigure(0, weight=0, minsize=50)  # Exp number column
                inner_frame.grid_columnconfigure(1, weight=0, uniform="col")              # Entry box column

                if i==1 or i==10 :
                    ctk.CTkLabel(inner_frame, text="Exp No.", font=("Arial", 15)).grid(row=0, column=0, sticky="w", padx=5)
                    ctk.CTkLabel(inner_frame, text="LO", font=("Arial", 15)).grid(row=0, column=1, sticky="nsw", padx=(50,5))
                else :
                    exp_lo_labels[f"exp{i-1}_lo_label"] = ctk.CTkLabel(inner_frame, text=f"{k}:", font=("Arial", 15))
                    exp_lo_labels[f"exp{i-1}_lo_label"].grid(row=0, column=0, padx=5, sticky="nsw")
                    if i<10:
                        exp_lo_entry[f"exp{i-1}_lo_entry"] = ctk.CTkEntry(inner_frame, placeholder_text="1,2,3,4,5,6", font=("Arial", 15), width=150)
                        exp_lo_entry[f"exp{i-1}_lo_entry"].grid(row=0, column=1, padx=(5, 5), sticky="nsew")
                        exp_lo_entry[f"exp{i-1}_lo_entry"].configure(state="disabled", fg_color="gray")
                    else :
                        exp_lo_entry[f"exp{i-2}_lo_entry"] = ctk.CTkEntry(inner_frame, placeholder_text="1,2,3,4,5,6", font=("Arial", 15), width=150)
                        exp_lo_entry[f"exp{i-2}_lo_entry"].grid(row=0, column=1, padx=(5, 5), sticky="nsew")
                        exp_lo_entry[f"exp{i-2}_lo_entry"].configure(state="disabled", fg_color="gray")
                    k+=1

            vertical_line = ctk.CTkFrame(scroll_frames[tab_name], width=2, fg_color="white")
            vertical_line.grid(row=1, column=3, rowspan=right_row+2, sticky="ns", padx=5)














            # mini_project_lo_lables = create_label(" LO Mapping ", "Mini project", "Arial", 15, 700, 375)

            # mini_project_lo_label = {}
            # mini_project_lo_entry = {}

            # for i in range(1,5):
            #     mini_project_lo_label[f"mini_project_factor{i}_lo_label"] = create_label(" LO Mapping ", f"Factor {i}: ", "Arial", 15, 50 + (350*(i-1)), 425)
            #     mini_project_lo_entry[f"mini_project_factor{i}_lo_entry"] = create_entry_box(" LO Mapping ", "1,2,3,4,5,6", "Arial", 15, 175, 150 + (350*(i-1)), 425)
            
            # for entry in assignment_lo_entry.values():
            #     entry.configure(state="disabled", fg_color="gray")

            # next_2_lab_button = create_button(" LO Mapping ", "Download", "Arial", 20, 200, 40, download_template_lab, 650, 600)

            # enterLO = create_label(" LO Information ", "Enter the LO's Description", "Arial", 20, 700, 50)
            # noOfLOLabel = create_label(" LO Information ", "Select No. of LO's: ", "Arial", 15, 550, 100)
            # noOfLOOption = create_dropdown(" LO Information ", ['Select No of LO\'s', '5', '6'], "Arial", 15, 300, lo_check, 750, 100)

            # LO_label = {}
            # LO_entry = {}

            # for i in range(1,7):
            #     LO_label[f"LO{i}_label"] = create_label(" LO Information ", f"LO{i}: ", "Arial", 15, 550, 150 + (50*(i-1)))
            #     LO_entry[f"LO{i}_entry"] = create_entry_box(" LO Information ", "", "Arial", 15, 500, 650, 150 + (50*(i-1)))

            # for entry in LO_entry.values():
            #     entry.configure(state="disabled", fg_color="gray")

            # next_lab_button = create_button(" LO Information ", "Next", "Arial", 20, 200, 40, switch_lab, 725, 500)

            # path_entry_lab=ctk.CTkEntry(tabview.tab(" Upload Excel File (Lab) "))

            # # button_process=ctk.CTkButton(tabview.tab(" Upload Excel File "),text="Process",width=100,height=30,command=process_file)
            # # button_process.place(x=500,y=500)

            # upload_Label_lab = create_label(" Upload Excel File (Lab) ", "Upload you excel file with the marks entered:", "Arial", 25, 550, 50)
            # path_label_lab = create_label(" Upload Excel File (Lab) ", "Path of file", "Arial", 15, 650, 110)
            # button_upload_lab = create_button(" Upload Excel File (Lab) ", "Upload", "Arial", 20, 200, 40, upload_lab_file, 400, 100)


            # line_lab = ctk.CTkFrame(master=tabview.tab(" Upload Excel File (Lab) "), height=2, width=1200, fg_color="white")
            # line_lab.place(x=150,y=200)

            # process_Label_lab = create_label(" Upload Excel File (Lab) ", "Process the excel file you uploaded:", "Arial", 25, 600, 250)

            # setEmailProcessedLabel_lab = create_label(" Upload Excel File (Lab) ", "Enter the Email ID to send the calculated sheet.", "Arial", 20, 200, 325)

            # # important_label = create_label(" Upload Excel File ", "Important: Please fill the no. of CO\'s field and the CO\'s in the CO Information page and AL values in Target level of tests page before processing the file", "Arial", 20, 100, 325)
            # # important_label.configure(text_color="black", fg_color="yellow")

            # emailTextProcessed_lab = create_entry_box(" Upload Excel File (Lab) ", "", "Arial", 15, 500, 700, 325)

            # button_process_lab = create_button(" Upload Excel File (Lab) ", "Process", "Arial", 20, 200, 40, process_file_lab, 650, 425)

        

            # back_button = ctk.CTkButton(lo_window, text="Back", command=lambda: self.go_back(lo_window))
            # back_button.place(x=1300,y=40)








            # ---------- LO Information Tab ----------
            lo_frame = scroll_frames[" LO Information "]
            tab_name=" LO Information "

            for i in range(4):
                scroll_frames[tab_name].grid_columnconfigure(i, weight=1)

            enterLO = create_label(" LO Information ", "Enter the LO's Description", "Arial", 20, row=0, column=1, colspan=2, sticky="nsew", padx=10, pady=(10, 20))

            noOfLOLabel = create_label(" LO Information ", "Select No. of LO's: ", "Arial", 15, row=1, column=1, sticky="w", padx=10)
            noOfLOOption = create_dropdown(" LO Information ", ['Select No of LO\'s', '5', '6'], "Arial", 15, 500, lo_check, row=1, column=2, sticky="w", padx=10)

            LO_label = {}
            LO_entry = {}

            for i in range(1, 7):
                LO_label[f"LO{i}_label"] = create_label(" LO Information ", f"LO{i}: ", "Arial", 15, row=i + 1, column=1, sticky="w", padx=10)
                LO_entry[f"LO{i}_entry"] = create_entry_box(" LO Information ", "", "Arial", 15, 500, row=i + 1, column=2, sticky="w", padx=10)
                LO_entry[f"LO{i}_entry"].configure(state="disabled", fg_color="gray")

            next_lab_button = create_button(" LO Information ", "Next", "Arial", 20, 300, 40, switch_lab, row=8, column=1,colspan=2, sticky="", padx=20, pady=20)

            # Disable assignment LO entry fields
            for entry in assignment_lo_entry.values():
                entry.configure(state="disabled", fg_color="gray")

            # Download button
            next_2_lab_button = create_button(" LO Mapping ", "Download", "Arial", 20, 200, 20, download_template_lab, row=right_row, column=5, sticky="", padx=0, pady=0)
            next_2_lab_button.grid_configure(rowspan=2)


            # ---------- Upload Excel File (Lab) Tab ----------
            lab_frame = scroll_frames[" Upload Excel File (Lab) "]
            tab_name=" Upload Excel File (Lab) "

            for i in range(4):
                scroll_frames[tab_name].grid_columnconfigure(i, weight=1)

            upload_Label_lab = create_label(" Upload Excel File (Lab) ", "Upload your excel file with the marks entered:", "Arial", 25, row=0, column=1,colspan=2, sticky="nsew", padx=10, pady=20)

            path_label_lab = create_label(" Upload Excel File (Lab) ", "Upload Or Enter the Path of File", "Arial", 20, row=2, column=1, sticky="w", padx=10)
            path_entry_lab = ctk.CTkEntry(lab_frame, font=("Arial", 15), width=500)
            path_entry_lab.grid(row=2, column=2, padx=10, pady=5, sticky="w")

            button_upload_lab = create_button(" Upload Excel File (Lab) ", "Upload", "Arial", 20, 200, 40, upload_lab_file, row=3, column=1,colspan=2, sticky="", padx=10)

            # Line separator
            line_separator = ctk.CTkFrame(lab_frame, height=2, fg_color="white")  # Use your desired color
            line_separator.grid(row=4, column=1, columnspan=2, sticky="ew", padx=10, pady=10)


            process_Label_lab = create_label(" Upload Excel File (Lab) ", "Process the excel file you uploaded:", "Arial", 25, row=5, column=1,colspan=2, sticky="nsew", padx=10, pady=10)

            setEmailProcessedLabel_lab = create_label(" Upload Excel File (Lab) ", "Enter the Email ID to send the calculated sheet.", "Arial", 20, row=6, column=1, sticky="w", padx=10)
            emailTextProcessed_lab = create_entry_box(" Upload Excel File (Lab) ", "", "Arial", 15, 500, row=6, column=2, sticky="w", padx=10)

            button_process_lab = create_button(" Upload Excel File (Lab) ", "Process", "Arial", 20, 200, 40, process_file_lab, row=7, column=1,colspan=2, sticky="", padx=20, pady=20)


            # ---------- Back Button ----------
            # back_button = ctk.CTkButton(lo_window, text="Back", command=lambda: self.go_back(lo_window))
            # back_button.place(x=1300, y=40)  # Still using place because it's global in window
            
            lo_window.mainloop()
    
    def open_main_page(self):
        
        ctk.set_appearance_mode("dark")  # Modes: system (default), light, dark
        ctk.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green
        
        
        self.app = ctk.CTk()  # creating custom tkinter window
        self.app.title('CO-PO')

        screen_width=self.app.winfo_screenwidth()
        screen_height=self.app.winfo_screenheight()
       
        # Set window size (like 80% of screen)
        window_width = int(screen_width * 0.8)
        window_height = int(screen_height * 0.8)
        # Center the window
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        # Set the window position and size
        self.app.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Let main_frame expand when window resizes
        self.app.rowconfigure(0, weight=1)
        self.app.columnconfigure(0, weight=1)

        self.main_frame = ctk.CTkFrame(master=self.app)
        # self.main_frame.pack(expand=True, fill="both", padx=10, pady=10)
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.main_frame.columnconfigure((0,1), weight=1)
        self.main_frame.rowconfigure(0, weight=1)

        # # Half-width and full-height of main_frame for coframe
        # self.coframe = ctk.CTkFrame(master=self.main_frame)
        # self.coframe.place(relx=0, rely=0, relwidth=0.5, relheight=1)

        # # Half-width and full-height of main_frame for loframe
        # self.loframe = ctk.CTkFrame(master=self.main_frame)
        # self.loframe.place(relx=0.5, rely=0, relwidth=0.5, relheight=1)
        

        # Left frame (CO)
        self.coframe = ctk.CTkFrame(master=self.main_frame)
        self.coframe.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Right frame (LO)
        self.loframe = ctk.CTkFrame(master=self.main_frame)
        self.loframe.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")



        # # Padding around coframe and loframe
        # frame_padding = 0.01  # Adjust this to increase/decrease padding (e.g., 1% of the width/height)

        # # Half-width and full-height of main_frame for coframe with padding
        # self.coframe = ctk.CTkFrame(master=self.main_frame)
        # self.coframe.place(
        #     relx=frame_padding,  # Start slightly inward (left padding)
        #     rely=frame_padding,  # Start slightly downward (top padding)
        #     relwidth=0.5 - (frame_padding),  # Reduce width for padding on both sides
        #     relheight=1 - (2*frame_padding)  # Reduce height for padding on top and bottom
        # )

        # # Half-width and full-height of main_frame for loframe with padding
        # self.loframe = ctk.CTkFrame(master=self.main_frame)
        # self.loframe.place(
        #     relx=0.5 + frame_padding,  # Start slightly after the midpoint (left padding)
        #     rely=frame_padding,  # Start slightly downward (top padding)
        #     relwidth=0.5 - (frame_padding),  # Reduce width for padding on both sides
        #     relheight=1 - (2*frame_padding)  # Reduce height for padding on top and bottom
        # )
        
        def resource_path(relative_path):
            """Get the absolute path to a resource, handling PyInstaller paths."""
            if hasattr(sys, '_MEIPASS'):  # PyInstaller extracts files to _MEIPASS
                return os.path.join(sys._MEIPASS, relative_path)
            return os.path.join(os.path.abspath("."), relative_path)

        image_path = resource_path(f"./images/coFinal.png") 

        # Load the image and create a CTkImage
        background_image = Image.open(image_path)
        bg_image = ctk.CTkImage(background_image, size=(512,313))

        

        # # Add a label to hold the background image
        # bg_label = ctk.CTkLabel(master=self.coframe, image=bg_image, text="")
        # bg_label.place(relx=0.5, rely=0.45, anchor="center")
        
        image_path1 = resource_path(f"./images/loFinal.png")

        # Load the image and create a CTkImage
        background_image1 = Image.open(image_path1)
        bg_image1 = ctk.CTkImage(background_image1, size=(512,313))

        

        # # Add a label to hold the background image
        # bg_label1 = ctk.CTkLabel(master=self.loframe, image=bg_image1, text="")
        # bg_label1.place(relx=0.5, rely=0.45, anchor="center")

        # cobutton = ctk.CTkButton(self.coframe,width=180,height=40, text="CO - PO",font=("Helvetica",25), command=self.open_co_window)
        # cobutton.place(x=280,y=600) 
        
    
        # lobutton = ctk.CTkButton(self.loframe,width=180,height=40, text="LO",font=("Helvetica", 25), command=self.open_lo_window)
        # lobutton.place(x=280,y=600)
        


        # Configure grid for coframe and loframe
        self.coframe.grid_rowconfigure(0, weight=1)
        self.coframe.grid_rowconfigure(1, weight=0)
        self.coframe.grid_columnconfigure(0, weight=1)

        self.loframe.grid_rowconfigure(0, weight=1)
        self.loframe.grid_rowconfigure(1, weight=0)
        self.loframe.grid_columnconfigure(0, weight=1)

        # ========== CO FRAME ==========
        bg_label = ctk.CTkLabel(master=self.coframe, image=bg_image, text="")
        bg_label.grid(row=0, column=0, pady=(40, 20), sticky="")  # Center top with padding

        cobutton = ctk.CTkButton(
            self.coframe,
            width=180,
            height=40,
            text="CO - PO",
            font=("Helvetica", 25),
            command=self.open_co_window
        )
        cobutton.grid(row=1, column=0, pady=(10, 40), sticky="")

        # ========== LO FRAME ==========
        bg_label1 = ctk.CTkLabel(master=self.loframe, image=bg_image1, text="")
        bg_label1.grid(row=0, column=0, pady=(40, 20), sticky="")

        lobutton = ctk.CTkButton(
            self.loframe,
            width=180,
            height=40,
            text="LO",
            font=("Helvetica", 25),
            command=self.open_lo_window
        )
        lobutton.grid(row=1, column=0, pady=(10, 40), sticky="")

        











        
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
