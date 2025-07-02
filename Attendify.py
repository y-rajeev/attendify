import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module='xlrd')

class ExcelCleanerApp:
    def __init__(self, master):
        self.master = master
        self.master.title("WorkLog Extractor Pro")
        self.master.geometry("500x400")
        self.master.resizable(False, False)

        self.input_file = None
        self.output_file = None

        self.create_widgets()

    def create_widgets(self):
        # Header Frame
        header = tk.Frame(self.master, bg="#4CAF50", height=60)
        header.pack(fill='x')
        tk.Label(header, text="Attendance Extractor", font=("Segoe UI", 18, "bold"), fg="white", bg="#4CAF50").pack(pady=10)

        # Body Frame
        body = tk.Frame(self.master, padx=20, pady=20)
        body.pack(fill='both', expand=True)

        # File input
        tk.Label(body, text="Select the Excel .xls File", font=("Segoe UI", 11)).grid(row=0, column=0, sticky="w")
        tk.Button(body, text="Browse File", command=self.browse_input, bg="#2196F3", fg="white", width=15).grid(row=0, column=1, padx=10)
        self.input_label = tk.Label(body, text="No file selected", fg="gray", font=("Segoe UI", 9))
        self.input_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))

        # Output selection
        tk.Label(body, text="Choose Save Location", font=("Segoe UI", 11)).grid(row=2, column=0, sticky="w")
        tk.Button(body, text="Set Output", command=self.browse_output, bg="#FF9800", fg="white", width=15).grid(row=2, column=1, padx=10)
        self.output_label = tk.Label(body, text="No destination selected", fg="gray", font=("Segoe UI", 9))
        self.output_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=(0, 10))

        # Progress bar
        self.progress = ttk.Progressbar(body, length=400, mode='determinate', maximum=100)
        self.progress.grid(row=4, column=0, columnspan=2, pady=10)

        self.status_label = tk.Label(body, text="", font=("Segoe UI", 10), fg="blue")
        self.status_label.grid(row=5, column=0, columnspan=2)

        # Run Button
        self.run_button = tk.Button(body, text="Start Extraction", command=self.run_cleaning,
                                    bg="#4CAF50", fg="white", font=("Segoe UI", 11, "bold"), width=30)
        self.run_button.grid(row=6, column=0, columnspan=2, pady=15)

        # Footer
        tk.Label(self.master, text="© 2025 Attendance Extractor | Python Tools", font=("Segoe UI", 8), fg="gray").pack(side="bottom", pady=5)

    def browse_input(self):
        self.input_file = filedialog.askopenfilename(title="Select .xls File", filetypes=[("Excel 97-2003 Files", "*.xls")])
        self.input_label.config(text=self.input_file if self.input_file else "No file selected")

    def browse_output(self):
        self.output_file = filedialog.asksaveasfilename(
            title="Save As", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
        )
        self.output_label.config(text=self.output_file if self.output_file else "No destination selected")

    def run_cleaning(self):
        if not self.input_file or not self.output_file:
            messagebox.showerror("Missing Info", "Please select both input and output files.")
            return

        self.run_button.config(state=tk.DISABLED)
        self.status_label.config(text="⏳ Processing...")
        self.progress['value'] = 0

        threading.Thread(target=self.clean_excel).start()

    def clean_excel(self):
        try:
            xls = pd.ExcelFile(self.input_file)
            sheet_count = len(xls.sheet_names)
            cleaned_data = {}

            for i, sheet in enumerate(xls.sheet_names, start=1):
                df = pd.read_excel(self.input_file, sheet_name=sheet, header=6)
                df = df.loc[:, df.columns.notna() & (df.columns != '') & ~df.columns.str.startswith('Unnamed')]
                cleaned_data[sheet] = df

                progress_value = int((i / sheet_count) * 100)
                self.progress['value'] = progress_value
                self.status_label.config(text=f"✅ Processed sheet {i} of {sheet_count}")
                self.master.update_idletasks()

            with pd.ExcelWriter(self.output_file, engine='xlsxwriter') as writer:
                for sheet, data in cleaned_data.items():
                    safe_name = sheet[:31].replace('/', '_').replace('\\', '_').replace('*', '_').replace('?', '_')
                    data.to_excel(writer, sheet_name=safe_name, index=False)

            self.status_label.config(text="🎉 Extraction Complete!")
            messagebox.showinfo("Success", f"File cleaned and saved to:\n{self.output_file}")

        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.run_button.config(state=tk.NORMAL)
            self.progress['value'] = 100

# Run the App
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCleanerApp(root)
    root.mainloop()
