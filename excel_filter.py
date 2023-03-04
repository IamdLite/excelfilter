from tkinter import filedialog, messagebox
import pandas as pd
import unicodedata
import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename, asksaveasfilename

class ExcelFilter:
    def __init__(self, master):
        self.master = master
        master.title("Excel Filter")
        master.geometry("400x250")
        
        # Create input file selection button and label
        self.input_label = ttk.Label(master, text="No file selected.")
        self.input_label.pack(pady=10)
        input_button = ttk.Button(master, text="Select input file", command=self.select_input_file)
        input_button.pack()
        
        # Create column selection drop-down menu
        self.column_label = ttk.Label(master, text="Select a column:")
        self.column_label.pack(pady=10)
        self.column_combo = ttk.Combobox(master)
        self.column_combo.pack()
        
        # Create value selection drop-down menu
        self.value_label = ttk.Label(master, text="Select a value:")
        self.value_label.pack(pady=10)
        self.value_combo = ttk.Combobox(master)
        self.value_combo.pack()
        
        # Create apply filter button
        apply_button = ttk.Button(master, text="Apply filter", command=self.apply_filter)
        apply_button.pack(pady=20)
        
        # Create output file selection button and label
        self.output_label = ttk.Label(master, text="")
        self.output_label.pack()
        output_button = ttk.Button(master, text="Select output file", command=self.select_output_file)
        output_button.pack(pady=10)
        
        # Create generate output button
        generate_button = ttk.Button(master, text="Generate output", command=self.generate_output)
        generate_button.pack(pady=20)
        
        # Initialize class variables
        self.input_file = None
        self.df = None
        self.normalized_column_names = None
        self.selected_column = None
        self.selected_value = None
        self.filters = []
        self.output_file = None
    
    def select_input_file(self):
        filetypes = [("Excel files", "*.xlsx")]
        input_file = askopenfilename(filetypes=filetypes)
        if input_file:
            self.input_file = input_file
            self.input_label.config(text=input_file)
            self.load_input_file()
    
    def load_input_file(self):
        try:
            self.df = pd.read_excel(self.input_file)
        except FileNotFoundError:
            print("The specified file does not exist.")
            exit()
        self.normalize_column_names()
        self.populate_column_combo()
        
    def normalize_column_names(self):
        self.normalized_column_names = [unicodedata.normalize("NFKD", str(c)).encode("ascii", "ignore").decode("utf-8") for c in self.df.columns]
    
    def populate_column_combo(self):
        self.column_combo["values"] = self.normalized_column_names
        self.column_combo.bind("<<ComboboxSelected>>", self.on_column_select)
    
    def on_column_select(self, event):
        self.selected_column = event.widget.get()
        self.populate_value_combo()
    
    def populate_value_combo(self):
        unique_values = self.df[self.selected_column].unique().tolist()
        self.value_combo["values"] = unique_values
    
    def apply_filter(self):
        # self.selected_value = self.value_combo.get()
        # if self.selected_value:
        #     self.filters.append((self.selected_column, self.selected_value))
        #     self.populate_value_combo()
        # # # filtered_df = df
        # # for self.selected_column_name, self.selected__value in self.filters:
        # #     filtered_df = filtered_df[filtered_df[self.selected_column_name] == self.selected_value]
        # #     if filtered_df.empty:
        # #         print("No rows found matching the specified filter.")
        # #         exit()
        
        selected_value = self.value_combo.get()
        self.filtered_df = self.df[self.df[self.selected_column] == selected_value]
        if self.filtered_df.empty:
            messagebox.showerror("Error", "No rows found matching the specified filter.")
            return
        
        self.select_output_file()
        

    def select_output_file(self):
        self.output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not self.output_file:
            self.status_label.config(text="No output file selected.")
            return
        self.generate_output()
    
    def generate_output(self):
        # self.value_combo.to_excel(self.output_file, index=False)
        # self.status_label.config(text=f"{len(self.value_combo.index)} rows were written to {self.output_file}.")
        self.filtered_df.to_excel(self.output_file, index=False)
        messagebox.showinfo("Success", f"{len(self.filtered_df.index)} rows were written to {self.output_file}.")

    
# Create a Tkinter root window
root = tk.Tk()
run = ExcelFilter(master=root)
root.mainloop()
