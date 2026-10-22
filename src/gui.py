"""
GUI module for VTU Result Automation
Handles the Tkinter interface for user input and configuration
"""

import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox


class AutomationGUI:
    """Manages the Tkinter GUI for the automation application"""
    
    def __init__(self, on_submit_callback):
        """
        Initialize the GUI
        
        Args:
            on_submit_callback: Function to call when user clicks Submit
        """
        self.on_submit = on_submit_callback
        
        # Create main window
        self.app = tk.Tk()
        self.app.title("Student Result Automation")
        
        # Initialize input variables
        self.driver_path = tk.StringVar()
        self.usn_file = tk.StringVar()
        self.website = tk.StringVar()
        self.save_path = tk.StringVar()
        self.start_row = tk.StringVar()
        self.end_row = tk.StringVar()
        self.auto_solve_captcha = tk.BooleanVar(value=True)
        
        self._create_widgets()
    
    def _pick_file(self, target, file_types):
        """
        Open a file picker dialog and store the selected path
        
        Args:
            target: StringVar to store the selected file path
            file_types: List of tuples defining allowed file types
        """
        path = filedialog.askopenfilename(title="Select a File", filetypes=file_types)
        if path:
            target.set(path)
    
    def _choose_driver(self):
        """Open file picker for ChromeDriver executable"""
        self._pick_file(
            self.driver_path,
            [("Executables", "*.exe"), ("All Files", "*.*")]
        )
    
    def _choose_usn_file(self):
        """Open file picker for input file containing USNs (Excel, CSV, Word DOCX, PDF, TXT)"""
        self._pick_file(
            self.usn_file,
            [
                ("All Supported Files", "*.xlsx;*.xls;*.csv;*.docx;*.pdf;*.txt"),
                ("Excel Files (*.xlsx, *.xls)", "*.xlsx;*.xls"),
                ("CSV Files (*.csv)", "*.csv"),
                ("Word Documents (*.docx)", "*.docx"),
                ("PDF Documents (*.pdf)", "*.pdf"),
                ("Text Files (*.txt)", "*.txt"),
                ("All Files", "*.*")
            ]
        )
    
    def _choose_save_path(self):
        """Open file dialog for output file path (Excel, CSV, Word DOCX, PDF)"""
        path = filedialog.asksaveasfilename(
            title="Save Results As",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel Workbook (*.xlsx)", "*.xlsx"),
                ("CSV Document (*.csv)", "*.csv"),
                ("Word Document (*.docx)", "*.docx"),
                ("PDF Document (*.pdf)", "*.pdf"),
                ("Legacy Excel 97-2003 (*.xls)", "*.xls"),
                ("All Files", "*.*")
            ]
        )
        if path:
            self.save_path.set(path)
    
    def _create_widgets(self):
        """Create and layout all GUI widgets"""
        
        # ChromeDriver path input
        tk.Label(self.app, text="CHROMEDRIVER PATH:").grid(
            row=0, column=0, sticky=tk.W, padx=10, pady=10
        )
        tk.Entry(self.app, textvariable=self.driver_path, width=40).grid(
            row=0, column=1, padx=10
        )
        tk.Button(self.app, text="CHOOSE", command=self._choose_driver).grid(
            row=0, column=2, padx=10
        )
        
        # USN file path input
        tk.Label(self.app, text="USN FILE PATH:").grid(
            row=1, column=0, sticky=tk.W, padx=10, pady=10
        )
        tk.Entry(self.app, textvariable=self.usn_file, width=40).grid(
            row=1, column=1, padx=10
        )
        tk.Button(self.app, text="CHOOSE", command=self._choose_usn_file).grid(
            row=1, column=2, padx=10
        )
        
        # Website URL input
        tk.Label(self.app, text="WEBSITE ADDRESS:").grid(
            row=3, column=0, sticky=tk.W, padx=10, pady=10
        )
        tk.Entry(self.app, textvariable=self.website, width=40).grid(
            row=3, column=1, padx=10
        )
        
        # Save path input
        tk.Label(self.app, text="SAVE PATH:").grid(
            row=4, column=0, sticky=tk.W, padx=10, pady=10
        )
        tk.Entry(self.app, textvariable=self.save_path, width=40).grid(
            row=4, column=1, padx=10
        )
        tk.Button(self.app, text="CHOOSE", command=self._choose_save_path).grid(
            row=4, column=2, padx=10
        )
        
        # Start row input
        tk.Label(self.app, text="USN START (Row):").grid(
            row=5, column=0, sticky=tk.W, padx=10, pady=10
        )
        tk.Entry(self.app, textvariable=self.start_row, width=20).grid(
            row=5, column=1, padx=10, sticky=tk.W
        )
        
        # End row input
        tk.Label(self.app, text="USN END (Row):").grid(
            row=6, column=0, sticky=tk.W, padx=10, pady=10
        )
        tk.Entry(self.app, textvariable=self.end_row, width=20).grid(
            row=6, column=1, padx=10, sticky=tk.W
        )

        # Auto-solve captcha checkbox
        tk.Checkbutton(
            self.app,
            text="Automatically solve CAPTCHA using OCR",
            variable=self.auto_solve_captcha
        ).grid(row=7, column=1, sticky=tk.W, padx=10, pady=5)
        
        # Submit and Quit buttons
        tk.Button(
            self.app, text="SUBMIT", command=self._handle_submit,
            width=15, bg="#dddddd"
        ).grid(row=8, column=0, pady=20, padx=10)
        
        tk.Button(
            self.app, text="QUIT", command=self.app.quit,
            width=15, bg="#ffcccc"
        ).grid(row=8, column=1, pady=20, padx=10)
        
        # Credit label
        tk.Label(
            self.app, text="Original by: Samarth Kashyap\nDepartment of CSE"
        ).grid(row=9, column=2, pady=10)
    
    def _handle_submit(self):
        """Handle submit button click - validate and trigger callback"""
        self.on_submit()
    
    def get_inputs(self):
        """
        Get all user input values from the GUI
        
        Returns:
            Dictionary containing all input values
        """
        return {
            "driver_path": self.driver_path.get().strip(),
            "usn_file": self.usn_file.get().strip(),
            "website": self.website.get().strip(),
            "save_path": self.save_path.get().strip(),
            "start_row": self.start_row.get(),
            "end_row": self.end_row.get(),
            "auto_solve_captcha": self.auto_solve_captcha.get()
        }
        
    def get_captcha_input(self):
        """
        Prompt user for captcha input via a dialog
        
        Returns:
            String with the captcha code entered by user
        """
        return simpledialog.askstring("Captcha Required", "Enter results page captcha:")
    
    def show_error(self, message):
        """Show an error message dialog"""
        messagebox.showerror("Error", message)
        
    def show_info(self, message):
        """Show an info message dialog"""
        messagebox.showinfo("Information", message)
        
    def show_warning(self, message):
        """Show a warning message dialog"""
        messagebox.showwarning("Warning", message)
    
    def run(self):
        """Start the GUI main loop"""
        self.app.mainloop()
