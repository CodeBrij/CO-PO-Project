import customtkinter as ctk 
from CTkMessagebox import CTkMessagebox

class User:
    def __init__(self):
        self.app = None
        self.cursor = None
        self.open_instruction_page()
    
    def open_instruction_page(self):
        ctk.set_appearance_mode("dark")  # Modes: system (default), light, dark
        ctk.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green
        
        
        self.app = ctk.CTk()  # creating custom tkinter window
        #self.app.geometry("700x540")
        self.app.title('Instructions')
         # Get screen width and height
        # screen_width = self.app.winfo_screenwidth()
        # screen_height = self.app.winfo_screenheight()
       
        screen_width=1500
        screen_height=800
       
        # Calculate the coordinates for centering the window
        x_position = 0
        y_position = 0
        
        # Set the window position and size
        self.app.geometry(f"{screen_width}x{screen_height}+{x_position}+{y_position}")
        
        self.main_frame = ctk.CTkFrame(master=self.app)
        self.main_frame.pack(expand=True, fill="both", padx=10, pady=10)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(2, weight=1)

        
        def start_button_event():
            if check_var.get()=='off':
                CTkMessagebox(title="Error", message="Please Check the box.",icon="cancel")
            else:
                import mainPage
                self.app.destroy()
                mainPage.User_mode()
                


        check_var = ctk.StringVar(value="off")
        checkbox = ctk.CTkCheckBox(self.main_frame, text="I have Read the instructions",font=("Helvetica", 20),
                                            variable=check_var, onvalue="on", offvalue="off")
        
        checkbox.place(x=500,y=600)     
        
        button1 = ctk.CTkButton(self.main_frame, text="Start",font=("Helvetica", 20), command=start_button_event)
        button1.place(x=800,y=600) 
          
        self.app.mainloop()
        
if __name__ == "__main__":
    User()