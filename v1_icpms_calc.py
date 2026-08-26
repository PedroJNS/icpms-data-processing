"""
==============================================================================
ICP-MS Data Processing & GUI Module (Agilent 7900)
Author: Pedro J. (PedroJNS)
License: GNU General Public License v3.0 (GPL-3.0)
Description: Integrated desktop GUI to upload Agilent 7900 MassHunter files, 
             input sample digestion data (mass & volume), and automatically 
             calculate real concentrations in solid samples (ppm and %).
==============================================================================
"""

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np


class ICPMSAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ICP-MS Data Processor - Agilent 7900")
        self.root.geometry("700x600")

        self.df_raw = None
        self.df_data = None
        self.header_row = None
        self.col_sample_idx = None
        
        # Digestion values storage: {sample_name: (mass_mg, volume_ml)}
        self.digestion_values = {}

        # --- Graphical Interface Layout ---

        # 1. File Upload Section
        self.btn_load = tk.Button(
            root, 
            text="1. Load File (Excel / CSV)", 
            command=self.load_file,
            font=("Arial", 11, "bold"),
            bg="#4CAF50",
            fg="white",
            padx=10, pady=5
        )
        self.btn_load.pack(pady=12)

        self.lbl_file = tk.Label(root, text="No file selected", font=("Arial", 9, "italic"))
        self.lbl_file.pack()

        # 2. Sample Selection Frame
        self.frame_list = tk.LabelFrame(root, text=" 2. Select samples to process ", padx=10, pady=10)
        self.frame_list.pack(fill="both", expand=True, padx=20, pady=10)

        self.scrollbar = tk.Scrollbar(self.frame_list, orient="vertical")
        self.listbox = tk.Listbox(
            self.frame_list, 
            selectmode=tk.MULTIPLE, 
            yscrollcommand=self.scrollbar.set,
            font=("Arial", 10)
        )
        self.scrollbar.config(command=self.listbox.yview)
        
        self.scrollbar.pack(side="right", fill="y")
        self.listbox.pack(side="left", fill="both", expand=True)

        # 3. Action Buttons
        self.frame_actions = tk.Frame(root)
        self.frame_actions.pack(pady=15)

        self.btn_digestion = tk.Button(
            self.frame_actions, 
            text="2. Input Digestion Data (Mass & Vol)", 
            command=self.open_digestion_dialog,
            font=("Arial", 10, "bold"),
            bg="#FF9800",
            fg="white",
            padx=8, pady=4,
            state=tk.DISABLED
        )
        self.btn_digestion.pack(side="left", padx=5)

        self.btn_process = tk.Button(
            self.frame_actions, 
            text="3. Calculate Concentrations", 
            command=self.calculate_and_display,
            font=("Arial", 10, "bold"),
            bg="#2196F3",
            fg="white",
            padx=8, pady=4,
            state=tk.DISABLED
        )
        self.btn_process.pack(side="left", padx=5)

    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Agilent 7900 Data File",
            filetypes=[("Excel and CSV Files", "*.xlsx *.xls *.csv"), ("All Files", "*.*")]
        )

        if not file_path:
            return

        try:
            if file_path.endswith('.csv'):
                self.df_raw = pd.read_csv(file_path, header=None)
            else:
                self.df_raw = pd.read_excel(file_path, header=None)

            # Locate the row containing 'Sample Name'
            self.header_row = None
            self.col_sample_idx = None

            for r_idx, row in self.df_raw.iterrows():
                for c_idx, val in enumerate(row):
                    val_str = str(val).strip().lower()
                    if "sample name" in val_str or val_str == "sample" or "nombre muestra" in val_str:
                        self.header_row = r_idx
                        self.col_sample_idx = c_idx
                        break
                if self.header_row is not None:
                    break

            if self.header_row is not None:
                headers = self.df_raw.iloc[self.header_row].fillna("").astype(str).tolist()
                self.df_data = self.df_raw.iloc[self.header_row + 1:].copy()
                self.df_data.columns = headers
            else:
                self.df_data = self.df_raw.copy()
                self.col_sample_idx = 0

            self.lbl_file.config(text=f"Loaded: {file_path.split('/')[-1]}")
            self.populate_sample_list()
            self.btn_digestion.config(state=tk.NORMAL)
            self.btn_process.config(state=tk.NORMAL)

        except Exception as e:
            messagebox.showerror("Loading Error", f"Could not parse file properly:\n{e}")

    def populate_sample_list(self):
        self.listbox.delete(0, tk.END)
        samples_series = self.df_data.iloc[:, self.col_sample_idx].dropna().astype(str)
        unique_samples = samples_series.unique()

        for sample in unique_samples:
            clean_name = sample.strip()
            if clean_name and clean_name.lower() != "nan":
                self.listbox.insert(tk.END, clean_name)

    def open_digestion_dialog(self):
        selected_indices = self.listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Warning", "Please select at least one sample from the list first.")
            return

        selected_samples = [self.listbox.get(i) for i in selected_indices]

        # Modal Dialog Window for Data Input
        dialog = tk.Toplevel(self.root)
        dialog.title("Digestion Parameters Entry")
        dialog.geometry("450x400")
        dialog.grab_set()

        lbl_info = tk.Label(dialog, text="Enter Mass (mg) and Dilution Volume (mL) for each sample:", font=("Arial", 9, "bold"))
        lbl_info.pack(pady=10)

        frame_inputs = tk.Frame(dialog)
        frame_inputs.pack(fill="both", expand=True, padx=15, pady=5)

        canvas = tk.Canvas(frame_inputs)
        scrollbar_y = tk.Scrollbar(frame_inputs, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_y.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")

        entries = {}
        for sample in selected_samples:
            row_frame = tk.Frame(scrollable_frame)
            row_frame.pack(fill="x", pady=4)

            lbl_s = tk.Label(row_frame, text=sample, width=20, anchor="w")
            lbl_s.pack(side="left")

            # Pre-fill values if previously entered, else default (100 mg, 50 mL)
            default_m, default_v = self.digestion_values.get(sample, ("100", "50"))

            ent_m = tk.Entry(row_frame, width=8)
            ent_m.insert(0, str(default_m))
            ent_m.pack(side="left", padx=2)

            lbl_m = tk.Label(row_frame, text="mg")
            lbl_m.pack(side="left", padx=(0, 10))

            ent_v = tk.Entry(row_frame, width=8)
            ent_v.insert(0, str(default_v))
            ent_v.pack(side="left", padx=2)

            lbl_v = tk.Label(row_frame, text="mL")
            lbl_v.pack(side="left")

            entries[sample] = (ent_m, ent_v)

        def save_and_close():
            try:
                for sample, (e_m, e_v) in entries.items():
                    m_val = float(e_m.get().replace(',', '.'))
                    v_val = float(e_v.get().replace(',', '.'))
                    if m_val <= 0 or v_val <= 0:
                        raise ValueError(f"Values must be > 0 for sample '{sample}'")
                    self.digestion_values[sample] = (m_val, v_val)
                dialog.destroy()
                messagebox.showinfo("Success", "Digestion parameters updated successfully.")
            except ValueError as ve:
                messagebox.showerror("Input Error", f"Invalid numeric input:\n{ve}")

        btn_save = tk.Button(dialog, text="Save Parameters", command=save_and_close, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        btn_save.pack(pady=10)

    def calculate_and_display(self):
        selected_indices = self.listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Warning", "Please select at least one sample.")
            return

        selected_samples = [self.listbox.get(i) for i in selected_indices]
        
        # Filter raw data
        col_series = self.df_data.iloc[:, self.col_sample_idx].astype(str).str.strip()
        df_filtered = self.df_data[col_series.isin(selected_samples)].copy()

        # Extract masses and volumes (or use defaults: 100 mg, 50 mL)
        masses = []
        volumes = []
        for s in df_filtered.iloc[:, self.col_sample_idx].astype(str).str.strip():
            m, v = self.digestion_values.get(s, (100.0, 50.0))
            masses.append(float(m))
            volumes.append(float(v))

        s_masses = pd.Series(masses, index=df_filtered.index)
        s_volumes = pd.Series(volumes, index=df_filtered.index)

        # Identify analytical element columns (skipping sample name column)
        element_cols = df_filtered.columns[1:]
        df_ppm = pd.DataFrame(index=df_filtered.index)
        
        FACTOR_PPM_TO_PCT = 10000.0

        for col in element_cols:
            raw_reading = pd.to_numeric(
                df_filtered[col].astype(str).str.replace(',', '.'), 
                errors='coerce'
            ).fillna(0.0)

            # Calculation: ppm (mg/kg) = [ppb (µg/L) * Volume (mL)] / Mass (mg)
            conc_ppm = (raw_reading * s_volumes) / s_masses
            df_ppm[col] = conc_ppm

        # Insert sample names back
        df_ppm.insert(0, df_filtered.columns[self.col_sample_idx], df_filtered.iloc[:, self.col_sample_idx])

        self.display_results_table(df_ppm)

    def display_results_table(self, df_result):
        results_window = tk.Toplevel(self.root)
        results_window.title(f"Processed Results - Concentrations in Solid (ppm / mg/kg)")
        results_window.geometry("1200x600")

        table_frame = tk.Frame(results_window)
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        scroll_y = tk.Scrollbar(table_frame, orient="vertical")
        scroll_x = tk.Scrollbar(table_frame, orient="horizontal")

        # Build clean column header labels
        column_titles = []
        if self.header_row is not None and self.header_row > 0:
            element_row = self.df_raw.iloc[self.header_row - 1].fillna("").astype(str)
            param_row = self.df_raw.iloc[self.header_row].fillna("").astype(str)
            
            current_elem = ""
            for e, p in zip(element_row, param_row):
                if e.strip():
                    current_elem = e.strip()
                p_str = p.strip()
                if current_elem and p_str and p_str not in ["Acq. Date-Time", "Type", "Level", "Sample Name"]:
                    column_titles.append(f"{current_elem}\n(ppm)")
                else:
                    column_titles.append(p_str)
        else:
            column_titles = [str(c) for c in df_result.columns]

        col_ids = [f"col_{i}" for i in range(len(column_titles))]

        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 9, "bold"))

        table = ttk.Treeview(
            table_frame, 
            columns=col_ids, 
            show="headings",
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set
        )

        scroll_y.config(command=table.yview)
        scroll_x.config(command=table.xview)

        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        table.pack(fill="both", expand=True)

        for col_id, title in zip(col_ids, column_titles):
            table.heading(col_id, text=title)
            max_len = max([len(line) for line in title.split('\n')])
            calculated_width = max(max_len * 10, 110)
            table.column(col_id, width=calculated_width, minwidth=80, anchor="center")

        for _, row in df_result.iterrows():
            formatted_vals = []
            for val in row:
                if isinstance(val, (int, float)):
                    formatted_vals.append(f"{val:.4f}")
                elif pd.notna(val) and str(val).lower() != "nan":
                    formatted_vals.append(str(val))
                else:
                    formatted_vals.append("")
            table.insert("", tk.END, values=formatted_vals)


if __name__ == "__main__":
    root = tk.Tk()
    app = ICPMSAnalyzerApp(root)
    root.mainloop()
