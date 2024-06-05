import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from tkinter.ttk import Progressbar, Combobox
from tkinterdnd2 import TkinterDnD, DND_FILES

class CSVGradeMerger:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Grade Merger - from CodeValidator to Canvas - Minh Nguyen")

        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.column1_name = tk.StringVar()
        self.column2_name = tk.StringVar()
        self.multiplier = tk.StringVar(value="3")

        tk.Label(root, text="CSV File 1:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        self.file1_entry = tk.Entry(root, textvariable=self.file1_path, width=50)
        self.file1_entry.grid(row=0, column=1, padx=10, pady=5)
        tk.Button(root, text="Browse", command=self.browse_file1).grid(row=0, column=2, padx=10, pady=5)

        tk.Label(root, text="CSV File 2:").grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
        self.file2_entry = tk.Entry(root, textvariable=self.file2_path, width=50)
        self.file2_entry.grid(row=1, column=1, padx=10, pady=5)
        tk.Button(root, text="Browse", command=self.browse_file2).grid(row=1, column=2, padx=10, pady=5)

        tk.Label(root, text="Column in File 1:").grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
        self.column1_combobox = Combobox(root, textvariable=self.column1_name, width=47)
        self.column1_combobox.grid(row=2, column=1, padx=10, pady=5)

        tk.Label(root, text="Column in File 2:").grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)
        self.column2_combobox = Combobox(root, textvariable=self.column2_name, width=47)
        self.column2_combobox.grid(row=3, column=1, padx=10, pady=5)

        tk.Label(root, text="Multiplier:").grid(row=4, column=0, padx=10, pady=5, sticky=tk.W)
        self.multiplier_entry = tk.Entry(root, textvariable=self.multiplier, width=50)
        self.multiplier_entry.grid(row=4, column=1, padx=10, pady=5)

        self.progress = Progressbar(root, orient=tk.HORIZONTAL, length=300, mode='determinate')
        self.progress.grid(row=5, columnspan=3, pady=20)

        tk.Button(root, text="Process", command=self.process_files).grid(row=6, column=1, pady=10)

        # Enable drag and drop
        self.file1_entry.drop_target_register(DND_FILES)
        self.file1_entry.dnd_bind('<<Drop>>', self.on_file1_drop)

        self.file2_entry.drop_target_register(DND_FILES)
        self.file2_entry.dnd_bind('<<Drop>>', self.on_file2_drop)

    def browse_file1(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        self.file1_path.set(file_path)
        if file_path:
            df = pd.read_csv(file_path)
            self.column1_combobox['values'] = list(df.columns)

    def browse_file2(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        self.file2_path.set(file_path)
        if file_path:
            df = pd.read_csv(file_path)
            self.column2_combobox['values'] = list(df.columns)

    def on_file1_drop(self, event):
        file_path = event.data.strip('{}')
        self.file1_path.set(file_path)
        if file_path:
            df = pd.read_csv(file_path)
            self.column1_combobox['values'] = list(df.columns)

    def on_file2_drop(self, event):
        file_path = event.data.strip('{}')
        self.file2_path.set(file_path)
        if file_path:
            df = pd.read_csv(file_path)
            self.column2_combobox['values'] = list(df.columns)

    def process_files(self):
        file1 = self.file1_path.get()
        file2 = self.file2_path.get()
        column1 = self.column1_name.get()
        column2 = self.column2_name.get()
        try:
            multiplier = float(self.multiplier.get())
        except ValueError:
            messagebox.showerror("Error", "Multiplier must be a number")
            return

        if not file1 or not file2 or not column1 or not column2:
            messagebox.showerror("Error", "Please select both CSV files and columns")
            return

        try:
            self.progress['value'] = 20
            self.root.update_idletasks()

            df1 = pd.read_csv(file1)
            df2 = pd.read_csv(file2)

            self.progress['value'] = 40
            self.root.update_idletasks()

            # Process to match rows and copy column1 multiplied by the multiplier to column2
            df1['Short Email'] = df1['Email address'].apply(lambda x: x.split('@')[0])

            self.progress['value'] = 60
            self.root.update_idletasks()

            for index, row in df2.iterrows():
                matching_row = df1[df1['Short Email'] == row['SIS Login ID']]
                if not matching_row.empty:
                    try:
                        value = int(matching_row[column1].values[0])
                        df2.at[index, column2] = value * multiplier
                    except ValueError:
                        pass  # If the value is not an integer, do nothing

            self.progress['value'] = 80
            self.root.update_idletasks()

            # Ask user where to save the output file
            output_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")], initialfile="new.csv")
            if output_path:
                df2.to_csv(output_path, index=False)

                self.progress['value'] = 100
                self.root.update_idletasks()

                messagebox.showinfo("Success", f"Grades have been processed and saved to {output_path}")
            else:
                messagebox.showinfo("Cancelled", "Save operation cancelled")

        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = CSVGradeMerger(root)
    root.mainloop()
