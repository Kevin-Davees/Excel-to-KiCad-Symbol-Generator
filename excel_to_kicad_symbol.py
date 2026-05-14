import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import subprocess
import os
import csv

class KiPartGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to KiCad Symbol Generator (KiPart)")
        self.root.geometry("620x580")
        
        # --- File Selection ---
        frame_files = ttk.LabelFrame(root, text="File Selection", padding=(10, 10))
        frame_files.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(frame_files, text="Input (Excel/CSV):").grid(row=0, column=0, sticky="w", pady=2)
        self.var_input = tk.StringVar()
        ttk.Entry(frame_files, textvariable=self.var_input, width=50).grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(frame_files, text="Browse...", command=self.browse_input).grid(row=0, column=2, pady=2)
        
        ttk.Label(frame_files, text="Output (.kicad_sym):").grid(row=1, column=0, sticky="w", pady=2)
        self.var_output = tk.StringVar()
        ttk.Entry(frame_files, textvariable=self.var_output, width=50).grid(row=1, column=1, padx=5, pady=2)
        ttk.Button(frame_files, text="Browse...", command=self.browse_output).grid(row=1, column=2, pady=2)

        # --- Symbol Properties ---
        frame_props = ttk.LabelFrame(root, text="Symbol Properties (Injected into CSV automatically)", padding=(10, 10))
        frame_props.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_props, text="Symbol Name:").grid(row=0, column=0, sticky="w", pady=2)
        self.var_name = tk.StringVar()
        ttk.Entry(frame_props, textvariable=self.var_name, width=30).grid(row=0, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(frame_props, text="Designation (e.g. U, J):").grid(row=1, column=0, sticky="w", pady=2)
        self.var_desig = tk.StringVar(value="U")
        ttk.Entry(frame_props, textvariable=self.var_desig, width=30).grid(row=1, column=1, sticky="w", padx=5, pady=2)

        # --- KiPart Options ---
        frame_opts = ttk.LabelFrame(root, text="KiPart Options", padding=(10, 10))
        frame_opts.pack(fill="x", padx=10, pady=5)

        self.var_overwrite = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame_opts, text="Overwrite existing library (-w)", variable=self.var_overwrite).grid(row=0, column=0, sticky="w", pady=2)

        self.var_merge = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame_opts, text="Merge into existing library (-m)", variable=self.var_merge).grid(row=1, column=0, sticky="w", pady=2)

        self.var_bundle = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame_opts, text="Bundle identical power/ground pins (-b)", variable=self.var_bundle).grid(row=2, column=0, sticky="w", pady=2)

        ttk.Label(frame_opts, text="Sort Pins by:").grid(row=0, column=1, sticky="w", padx=(20, 5), pady=2)
        self.var_sort = tk.StringVar(value="row")
        cb_sort = ttk.Combobox(frame_opts, textvariable=self.var_sort, values=["row", "num", "name"], width=10, state="readonly")
        cb_sort.grid(row=0, column=2, sticky="w", pady=2)

        ttk.Label(frame_opts, text="Default Side:").grid(row=1, column=1, sticky="w", padx=(20, 5), pady=2)
        self.var_side = tk.StringVar(value="")
        cb_side = ttk.Combobox(frame_opts, textvariable=self.var_side, values=["", "left", "right", "top", "bottom"], width=10, state="readonly")
        cb_side.grid(row=1, column=2, sticky="w", pady=2)

        ttk.Label(frame_opts, text="Extra Arguments:").grid(row=3, column=0, sticky="w", pady=10)
        self.var_extra = tk.StringVar()
        ttk.Entry(frame_opts, textvariable=self.var_extra, width=50).grid(row=3, column=1, columnspan=2, sticky="w", padx=5, pady=10)

        # --- Actions ---
        frame_actions = tk.Frame(root)
        frame_actions.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(frame_actions, text="Generate Symbol", command=self.generate, width=20).pack(side="right")
        
        # --- Log Text ---
        self.log_text = tk.Text(root, height=8, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.root.update()

    def browse_input(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.var_input.set(file_path)
            if not self.var_name.get():
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                self.var_name.set(base_name)

    def browse_output(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".kicad_sym",
            filetypes=[("KiCad Symbol Library", "*.kicad_sym"), ("Legacy KiCad Library", "*.lib"), ("All Files", "*.*")]
        )
        if file_path:
            self.var_output.set(file_path)

    def convert_to_kipart_csv(self, input_path, output_csv_path, part_name, ref_des):
        """Reads Excel/CSV and forcefully formats it to the exact spec required by KiPart"""
        try:
            if input_path.lower().endswith('.csv'):
                df = pd.read_csv(input_path)
            else:
                df = pd.read_excel(input_path)
        except Exception as e:
            raise Exception(f"Failed to read input file: {e}")

        # --- Header Detection ---
        if 'Pin' not in df.columns and 'Name' not in df.columns:
            if input_path.lower().endswith('.csv'):
                df = pd.read_csv(input_path, header=None)
            else:
                df = pd.read_excel(input_path, header=None)
                
            header_idx = -1
            for i, row in df.iterrows():
                row_vals = [str(x).strip().lower() for x in row.values if pd.notna(x)]
                if 'pin' in row_vals and 'name' in row_vals:
                    header_idx = i
                    break
            
            if header_idx != -1:
                df.columns = df.iloc[header_idx]
                df = df.iloc[header_idx+1:]
                df.reset_index(drop=True, inplace=True)
            else:
                raise Exception("Could not find 'Pin' and 'Name' headers in the file.")

        # ==========================================
        # AGGRESSIVE CLEANUP OF EXCEL "GHOST" CELLS
        # ==========================================
        
        # 1. Strip whitespace from column names and remove unnamed/ghost columns
        df.columns = [str(c).strip() for c in df.columns]
        valid_cols = [c for c in df.columns if c and 'unnamed' not in c.lower() and c.lower() != 'nan']
        df = df[valid_cols]

        # 2. Drop completely empty rows
        df.dropna(how='all', inplace=True)

        # 3. Drop rows where 'Pin' is completely missing (KiPart will crash otherwise)
        if 'Pin' in df.columns:
            df = df.dropna(subset=['Pin'])
            df = df[df['Pin'].astype(str).str.strip() != '']
            df = df[df['Pin'].astype(str).str.lower() != 'nan']

        # Replace remaining NaN with empty string
        df = df.fillna("")

        # Helper to stop integers like '1' from becoming '1.0' in pandas
        def clean_val(val):
            if pd.isna(val) or val == "" or str(val).lower() == 'nan':
                return ""
            if isinstance(val, float) and val.is_integer():
                return str(int(val))
            return str(val).strip()

        # Write to intermediate CSV file with proper Part Name and Reference injected
        with open(output_csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            
            # --- FIX APPLIED HERE ---
            # 1. Write Symbol Name and Properties on the FIRST ROW
            first_row = []
            if part_name:
                first_row.append(part_name)
            if ref_des:
                first_row.extend(['Reference', ref_des])
            writer.writerow(first_row)
                
            # 2. Write Headers on the SECOND ROW
            writer.writerow(df.columns.tolist())
            
            # 3. Write Pins Data
            for _, row in df.iterrows():
                row_data = [clean_val(x) for x in row.tolist()]
                writer.writerow(row_data)

    def generate(self):
        input_file = self.var_input.get()
        output_file = self.var_output.get()
        part_name = self.var_name.get().strip()
        ref_des = self.var_desig.get().strip()

        if not input_file or not os.path.exists(input_file):
            messagebox.showerror("Error", "Please select a valid input file.")
            return
        if not output_file:
            messagebox.showerror("Error", "Please select an output file location.")
            return
        if not part_name:
            messagebox.showerror("Error", "Please enter a Symbol Name.")
            return

        self.log_text.config(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state="disabled")

        input_dir = os.path.dirname(input_file)
        temp_csv = os.path.join(input_dir, f"{part_name}_kipart_ready.csv")
        
        self.log(f"Cleaning Excel file and converting to CSV...")
        try:
            self.convert_to_kipart_csv(input_file, temp_csv, part_name, ref_des)
            self.log(f"Created clean CSV: {os.path.basename(temp_csv)}")
        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Data Formatting Error", str(e))
            return

        cmd = ["kipart"]
        
        if self.var_overwrite.get() and not self.var_merge.get():
            cmd.append("-w")
        if self.var_merge.get():
            cmd.append("-m")
        if self.var_bundle.get():
            cmd.append("-b")
        
        if self.var_sort.get():
            cmd.extend(["-s", self.var_sort.get()])
        if self.var_side.get():
            cmd.extend(["--side", self.var_side.get()])
            
        extra_args = self.var_extra.get().strip()
        if extra_args:
            cmd.extend(extra_args.split())

        cmd.extend(["-o", output_file])
        cmd.append(temp_csv)

        self.log(f"Executing: {' '.join(cmd)}")
        
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, shell=(os.name == 'nt'))
            
            if result.stdout:
                self.log(result.stdout)
            if result.stderr:
                self.log(result.stderr)
                
            if result.returncode == 0:
                self.log("Successfully generated KiCad symbol library!")
                messagebox.showinfo("Success", f"Symbol successfully generated at:\n{output_file}")
            else:
                self.log(f"KiPart exited with error code {result.returncode}")
                messagebox.showerror("KiPart Error", f"KiPart failed with error code {result.returncode}.\nCheck the log for details.")
                
        except FileNotFoundError:
            self.log("ERROR: 'kipart' command not found. Is it installed?")
            messagebox.showerror("Error", "kipart command not found.\nPlease install it using 'pip install kipart'.")
        except Exception as e:
            self.log(f"Execution ERROR: {e}")
            messagebox.showerror("Execution Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = KiPartGUI(root)
    root.mainloop()
