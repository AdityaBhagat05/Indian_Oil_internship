


import os
import threading
import customtkinter as ctk
from tkinter import filedialog
import lookup

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class FormulaInjectorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Formula Injector Dashboard")
        self.geometry("550x600")
        
        self.target_file = None

        self.grid_columnconfigure(0, weight=1)

        title_label = ctk.CTkLabel(
            self, 
            text="Excel Formula Injector", 
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.grid(row=0, column=0, pady=(20, 10))

        self.file_btn = ctk.CTkButton(
            self, 
            text="Select Target Excel File", 
            command=self.pick_file
        )
        self.file_btn.grid(row=1, column=0, pady=10)

        config_frame = ctk.CTkFrame(self, fg_color="transparent")
        config_frame.grid(row=2, column=0, pady=15)

        self.month_entry = ctk.CTkEntry(
            config_frame, 
            placeholder_text="Month (e.g., Aug)", 
            width=120
        )
        self.month_entry.insert(0, "Aug")
        self.month_entry.grid(row=0, column=0, padx=10)

        self.year_entry = ctk.CTkEntry(
            config_frame, 
            placeholder_text="Year (e.g., 25)", 
            width=120
        )
        self.year_entry.insert(0, "25")
        self.year_entry.grid(row=0, column=1, padx=10)

        self.run_btn = ctk.CTkButton(
            self, 
            text="Inject Formulas", 
            command=self.start_process, 
            width=200
        )
        self.run_btn.grid(row=3, column=0, pady=20)

        self.log_output = ctk.CTkTextbox(
            self, 
            width=450, 
            height=250, 
            state="disabled"
        )
        self.log_output.grid(row=4, column=0, pady=(0, 20))

    def log_message(self, message):
        self.log_output.configure(state="normal")
        self.log_output.insert("end", message + "\n")
        self.log_output.see("end")
        self.log_output.configure(state="disabled")

    def pick_file(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")]
        )
        if filepath:
            self.target_file = filepath
            filename = os.path.basename(filepath)
            self.file_btn.configure(
                text=filename, 
                fg_color="#2FA572", 
                hover_color="#1D7A50"
            )
            self.log_message(f"Selected file: {filename}")

    def start_process(self):
        if not self.target_file:
            self.log_message("Error: Please select a target Excel file.")
            return

        month = self.month_entry.get().strip()
        year = self.year_entry.get().strip()

        if not month or not year:
            self.log_message("Error: Month and Year are required.")
            return

        self.run_btn.configure(state="disabled", text="Injecting...")
        threading.Thread(target=self.worker, args=(month, year), daemon=True).start()

    def worker(self, month, year):
        def gui_logger(msg):
            self.after(0, self.log_message, msg)
            
        try:
            lookup.process_excel(self.target_file, month, year, logger=gui_logger)
        except Exception as e:
            gui_logger(f"CRASHED! Here is the error: {e}")
        finally:
            self.after(0, lambda: self.run_btn.configure(state="normal", text="Inject Formulas"))

if __name__ == "__main__":
    app = FormulaInjectorApp()
    app.mainloop()