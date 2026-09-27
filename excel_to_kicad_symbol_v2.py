import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import subprocess
import os
import csv
import re
import json
import math

SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".kipart_gui_settings.json")

PIN_TYPES = ["input", "output", "bidirectional", "tri-state", "passive", "free",
             "unspecified", "power_in", "power_out", "open_collector",
             "open_emitter", "no_connect"]
PIN_STYLES = ["line", "inverted", "clock", "inverted_clock", "input_low",
              "clock_low", "output_low", "clock_fall", "non_logic"]
SIDES = ["", "left", "right", "top", "bottom"]

# Aliases people commonly use in their Excel pin tables -> canonical KiPart names
COLUMN_ALIASES = {
    "pin": "Pin", "pin number": "Pin", "pin no": "Pin", "pin#": "Pin",
    "num": "Pin", "number": "Pin", "pad": "Pin", "pin num": "Pin", "no": "Pin",
    "name": "Name", "pin name": "Name", "signal": "Name", "label": "Name",
    "function": "Name",
    "type": "Type", "pin type": "Type",
    "style": "Style", "pin style": "Style", "graphic style": "Style",
    "side": "Side", "position": "Side",
    "unit": "Unit", "unit number": "Unit", "bank": "Unit",
}


def clean_val(val):
    if pd.isna(val) or val == "" or str(val).lower() == "nan":
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


def read_and_clean(input_path):
    # Reads Excel/CSV, finds the header row, normalizes column names, drops ghost rows.
    try:
        if input_path.lower().endswith(".csv"):
            df = pd.read_csv(input_path)
        else:
            df = pd.read_excel(input_path)
    except Exception as e:
        raise Exception(f"Failed to read input file: {e}")

    if "Pin" not in df.columns and "Name" not in df.columns:
        if input_path.lower().endswith(".csv"):
            df = pd.read_csv(input_path, header=None)
        else:
            df = pd.read_excel(input_path, header=None)
        header_idx = -1
        for i, row in df.iterrows():
            row_vals = [str(x).strip().lower() for x in row.values if pd.notna(x)]
            if "pin" in row_vals and "name" in row_vals:
                header_idx = i
                break
        if header_idx != -1:
            df.columns = df.iloc[header_idx]
            df = df.iloc[header_idx + 1:]
            df.reset_index(drop=True, inplace=True)
        else:
            raise Exception("Could not find 'Pin' and 'Name' headers in the file.")

    df.columns = [str(c).strip() for c in df.columns]
    valid_cols = [c for c in df.columns if c and "unnamed" not in c.lower() and c.lower() != "nan"]
    df = df[valid_cols]

    # Normalize aliased column names to what KiPart expects
    rename = {}
    for c in df.columns:
        key = c.strip().lower()
        if key in COLUMN_ALIASES and COLUMN_ALIASES[key] not in df.columns:
            rename[c] = COLUMN_ALIASES[key]
    df = df.rename(columns=rename)

    df.dropna(how="all", inplace=True)

    if "Pin" not in df.columns:
        raise Exception("No pin-number column found (expected 'Pin').")
    df = df.dropna(subset=["Pin"])
    df = df[df["Pin"].astype(str).str.strip() != ""]
    df = df[df["Pin"].astype(str).str.lower() != "nan"]

    df = df.fillna("")

    # Fill empty pin names with the pin number so KiPart never chokes
    if "Name" in df.columns:
        mask = df["Name"].astype(str).str.strip() == ""
        df.loc[mask, "Name"] = df.loc[mask, "Pin"].astype(str)

    # Validate: duplicate pin numbers
    pins = df["Pin"].astype(str)
    dupes = pins[pins.duplicated()].unique().tolist()
    return df, dupes


def apply_grouping(df, enabled, case_insensitive, spacer_rows):
    # Sorts so identical pin names are contiguous and optionally inserts blank
    # spacer rows between groups (KiPart renders them as visible gaps).
    if not enabled or "Name" not in df.columns or len(df) == 0:
        return df
    df = df.copy()
    keys = df["Name"].astype(str).str.strip()
    if case_insensitive:
        keys = keys.str.lower()
    df["_grp"] = keys
    df = df.sort_values("_grp", kind="stable").drop(columns="_grp").reset_index(drop=True)
    if spacer_rows > 0:
        out = []
        prev = None
        for _, r in df.iterrows():
            k = str(r["Name"]).strip()
            if case_insensitive:
                k = k.lower()
            if prev is not None and k != prev:
                for _ in range(spacer_rows):
                    out.append(pd.Series({c: "" for c in df.columns}))
            out.append(r)
            prev = k
        df = pd.DataFrame(out, columns=df.columns).reset_index(drop=True)
    return df


def apply_symmetric(df, mode, start_side="left"):
    # Assigns a per-pin 'Side' column for balanced left/right placement.
    if mode == "None" or len(df) == 0:
        return df
    df = df.copy()
    n = len(df)
    other = "right" if start_side == "left" else "left"
    sides = []
    if mode == "Even split (top/bottom halves)":
        half = math.ceil(n / 2)
        sides = [start_side] * half + [other] * (n - half)
    elif mode == "Interleaved (L,R,L,R...)":
        for i in range(n):
            sides.append(start_side if i % 2 == 0 else other)
    elif mode == "Alternate per name-group":
        prev, cur = None, other  # first group starts on start_side
        names = df["Name"].astype(str).str.strip().str.lower() if "Name" in df.columns else [""] * n
        for nm in names:
            if nm != prev:
                cur = other if cur == start_side else start_side
                prev = nm
            sides.append(cur)
    df["Side"] = sides
    return df


def write_kipart_csv(df, output_csv_path, part_name, ref_des):
    with open(output_csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        first_row = []
        if part_name:
            first_row.append(part_name)
        if ref_des:
            first_row.extend(["Reference", ref_des])
        writer.writerow(first_row)
        writer.writerow(df.columns.tolist())
        for _, row in df.iterrows():
            writer.writerow([clean_val(x) for x in row.tolist()])


def make_square(path):
    # Post-processes the generated .kicad_sym so the body rectangle becomes a
    # square (width = max(width, height)), keeping the same center point.
    with open(path, "r", encoding="utf-8") as f:
        text = f.read()

    def repl(m):
        x1, y1, x2, y2 = map(float, m.groups())
        cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
        half = max(abs(x2 - x1), abs(y2 - y1)) / 2
        def fmt(v):
            return str(int(v)) if float(v).is_integer() else f"{v:.2f}".rstrip("0").rstrip(".")
        return (f"(rectangle (start {fmt(cx - half)} {fmt(cy - half)}) "
                f"(end {fmt(cx + half)} {fmt(cy + half)}))")

    new = re.sub(r"\(rectangle \(start ([-\d.]+) ([-\d.]+)\) \(end ([-\d.]+) ([-\d.]+)\)\)",
                 repl, text)
    if new != text:
        with open(path, "w", encoding="utf-8") as f:
            f.write(new)
        return True
    return False


class KiPartGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to KiCad Symbol Generator (KiPart) - v2")
        self.root.geometry("780x760")

        self._build_symbol_props()
        self._build_notebook()
        self._build_options()
        self._build_actions_and_log()
        self._load_settings()

    # ------------------------------------------------------------------ UI
    def _build_symbol_props(self):
        frame = ttk.LabelFrame(self.root, text="Symbol Properties", padding=(10, 8))
        frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(frame, text="Symbol Name:").grid(row=0, column=0, sticky="w")
        self.var_name = tk.StringVar()
        ttk.Entry(frame, textvariable=self.var_name, width=32).grid(row=0, column=1, sticky="w", padx=5)
        ttk.Label(frame, text="Designation (U, J, ...):  ").grid(row=0, column=2, sticky="w", padx=(15, 0))
        self.var_desig = tk.StringVar(value="U")
        ttk.Entry(frame, textvariable=self.var_desig, width=10).grid(row=0, column=3, sticky="w", padx=5)

    def _build_notebook(self):
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=10, pady=5)

        # ---- Tab 1: file mode ----
        tab_file = ttk.Frame(self.nb, padding=8)
        self.nb.add(tab_file, text="  File Mode  ")
        ttk.Label(tab_file, text="Input (Excel/CSV):").grid(row=0, column=0, sticky="w", pady=2)
        self.var_input = tk.StringVar()
        ttk.Entry(tab_file, textvariable=self.var_input, width=62).grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(tab_file, text="Browse...", command=self.browse_input).grid(row=0, column=2, pady=2)
        ttk.Label(tab_file, text="Output (.kicad_sym):").grid(row=1, column=0, sticky="w", pady=2)
        self.var_output = tk.StringVar()
        ttk.Entry(tab_file, textvariable=self.var_output, width=62).grid(row=1, column=1, padx=5, pady=2)
        ttk.Button(tab_file, text="Browse...", command=self.browse_output).grid(row=1, column=2, pady=2)
        ttk.Button(tab_file, text="Preview / Validate File", command=self.preview_file).grid(row=2, column=1, sticky="w", pady=6)

        # ---- Tab 2: manual pin editor ----
        tab_edit = ttk.Frame(self.nb, padding=8)
        self.nb.add(tab_edit, text="  Pin Editor (Manual Entry)  ")

        cols = ("pin", "name", "type", "style", "side")
        heads = ["Pin", "Name", "Type", "Style", "Side"]
        widths = [70, 220, 130, 130, 80]
        self.tree = ttk.Treeview(tab_edit, columns=cols, show="headings", height=12, selectmode="extended")
        for c, h, w in zip(cols, heads, widths):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="w")
        vsb = ttk.Scrollbar(tab_edit, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, columnspan=8, sticky="nsew")
        vsb.grid(row=0, column=8, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        tab_edit.rowconfigure(0, weight=1)
        tab_edit.columnconfigure(0, weight=1)

        # entry row
        ttk.Label(tab_edit, text="Pin:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.ent_pin = ttk.Entry(tab_edit, width=8)
        self.ent_pin.grid(row=2, column=0, sticky="w")
        ttk.Label(tab_edit, text="Name:").grid(row=1, column=1, sticky="w", pady=(8, 0))
        self.ent_name = ttk.Entry(tab_edit, width=30)
        self.ent_name.grid(row=2, column=1, sticky="w")
        ttk.Label(tab_edit, text="Type:").grid(row=1, column=2, sticky="w", pady=(8, 0))
        self.cb_type = ttk.Combobox(tab_edit, values=PIN_TYPES, width=16, state="readonly")
        self.cb_type.set("passive")
        self.cb_type.grid(row=2, column=2, sticky="w", padx=2)
        ttk.Label(tab_edit, text="Style:").grid(row=1, column=3, sticky="w", pady=(8, 0))
        self.cb_style = ttk.Combobox(tab_edit, values=PIN_STYLES, width=16, state="readonly")
        self.cb_style.set("line")
        self.cb_style.grid(row=2, column=3, sticky="w", padx=2)
        ttk.Label(tab_edit, text="Side:").grid(row=1, column=4, sticky="w", pady=(8, 0))
        self.cb_side_pin = ttk.Combobox(tab_edit, values=SIDES, width=8, state="readonly")
        self.cb_side_pin.set("")
        self.cb_side_pin.grid(row=2, column=4, sticky="w", padx=2)

        ttk.Button(tab_edit, text="Add", command=self.add_pin).grid(row=2, column=5, padx=3)
        ttk.Button(tab_edit, text="Update Selected", command=self.update_pin).grid(row=2, column=6, padx=3)
        ttk.Button(tab_edit, text="Delete", command=self.delete_pins).grid(row=2, column=7, padx=3)

        btns = ttk.Frame(tab_edit)
        btns.grid(row=3, column=0, columnspan=8, sticky="w", pady=6)
        ttk.Button(btns, text="Move Up", command=lambda: self.move_pin(-1)).pack(side="left", padx=2)
        ttk.Button(btns, text="Move Down", command=lambda: self.move_pin(1)).pack(side="left", padx=2)
        ttk.Button(btns, text="Duplicate", command=self.duplicate_pin).pack(side="left", padx=2)
        ttk.Button(btns, text="Sort by Pin #", command=self.sort_tree).pack(side="left", padx=2)
        ttk.Button(btns, text="Clear Table", command=self.clear_tree).pack(side="left", padx=2)
        ttk.Button(btns, text="Import from File...", command=self.import_to_tree).pack(side="left", padx=(20, 2))
        ttk.Button(btns, text="Export Table to CSV...", command=self.export_tree_csv).pack(side="left", padx=2)

        self.lbl_count = ttk.Label(tab_edit, text="0 pins")
        self.lbl_count.grid(row=3, column=8, sticky="e")

    def _build_options(self):
        frame = ttk.LabelFrame(self.root, text="Layout & KiPart Options", padding=(10, 8))
        frame.pack(fill="x", padx=10, pady=5)

        # --- grouping ---
        grp = ttk.LabelFrame(frame, text="Grouping", padding=(8, 5))
        grp.grid(row=0, column=0, rowspan=3, sticky="ns", padx=(0, 10))
        self.var_group = tk.BooleanVar(value=False)
        ttk.Checkbutton(grp, text="Group pins by name", variable=self.var_group,
                        command=self._toggle_group).grid(row=0, column=0, columnspan=2, sticky="w")
        self.var_grp_case = tk.BooleanVar(value=True)
        self.cb_grp_case = ttk.Checkbutton(grp, text="Case-insensitive", variable=self.var_grp_case, state="disabled")
        self.cb_grp_case.grid(row=1, column=0, columnspan=2, sticky="w")
        ttk.Label(grp, text="Spacer rows between groups:").grid(row=2, column=0, sticky="w", pady=2)
        self.var_spacer = tk.IntVar(value=1)
        self.sp_spacer = ttk.Spinbox(grp, from_=0, to=5, textvariable=self.var_spacer, width=5, state="disabled")
        self.sp_spacer.grid(row=2, column=1, sticky="w")

        # --- symmetric distribution ---
        sym = ttk.LabelFrame(frame, text="Pin Distribution", padding=(8, 5))
        sym.grid(row=0, column=1, rowspan=3, sticky="ns", padx=(0, 10))
        ttk.Label(sym, text="Mode:").grid(row=0, column=0, sticky="w")
        self.var_dist = tk.StringVar(value="None")
        ttk.Combobox(sym, textvariable=self.var_dist, state="readonly", width=28,
                     values=["None", "Even split (top/bottom halves)",
                             "Interleaved (L,R,L,R...)", "Alternate per name-group"]
                     ).grid(row=0, column=1, sticky="w")
        ttk.Label(sym, text="Start side:").grid(row=1, column=0, sticky="w")
        self.var_dist_side = tk.StringVar(value="left")
        ttk.Combobox(sym, textvariable=self.var_dist_side, state="readonly", width=10,
                     values=["left", "right"]).grid(row=1, column=1, sticky="w")
        ttk.Label(sym, text="(Overrides any 'Side' column)", foreground="#666").grid(row=2, column=0, columnspan=2, sticky="w")

        # --- body shape ---
        shp = ttk.LabelFrame(frame, text="Body Shape", padding=(8, 5))
        shp.grid(row=0, column=2, rowspan=3, sticky="ns", padx=(0, 10))
        self.var_shape = tk.StringVar(value="Outline (default)")
        for i, s in enumerate(["Outline (default)", "Rectangle box (-r)", "Centered (-c)",
                               "Rectangle + centered (-r -c)", "Force square (post-process)"]):
            ttk.Radiobutton(shp, text=s, variable=self.var_shape, value=s).grid(row=i, column=0, sticky="w")

        # --- misc kipart flags ---
        misc = ttk.LabelFrame(frame, text="KiPart Flags", padding=(8, 5))
        misc.grid(row=0, column=3, rowspan=3, sticky="ns")
        self.var_overwrite = tk.BooleanVar(value=True)
        ttk.Checkbutton(misc, text="Overwrite library (-w)", variable=self.var_overwrite).grid(row=0, column=0, columnspan=2, sticky="w")
        self.var_merge = tk.BooleanVar(value=False)
        ttk.Checkbutton(misc, text="Merge into existing (-m)", variable=self.var_merge).grid(row=1, column=0, columnspan=2, sticky="w")
        self.var_bundle = tk.BooleanVar(value=False)
        ttk.Checkbutton(misc, text="Bundle power pins (-b)", variable=self.var_bundle).grid(row=2, column=0, columnspan=2, sticky="w")
        self.var_icase = tk.BooleanVar(value=False)
        ttk.Checkbutton(misc, text="Ignore case (-i)", variable=self.var_icase).grid(row=3, column=0, columnspan=2, sticky="w")
        self.var_fuzzy = tk.BooleanVar(value=False)
        ttk.Checkbutton(misc, text="Fuzzy name match (-f)", variable=self.var_fuzzy).grid(row=4, column=0, columnspan=2, sticky="w")
        ttk.Label(misc, text="Sort pins by:").grid(row=5, column=0, sticky="w")
        self.var_sort = tk.StringVar(value="row")
        ttk.Combobox(misc, textvariable=self.var_sort, values=["row", "num", "name"], width=8, state="readonly").grid(row=5, column=1, sticky="w")
        ttk.Label(misc, text="Default side:").grid(row=6, column=0, sticky="w")
        self.var_side = tk.StringVar(value="")
        ttk.Combobox(misc, textvariable=self.var_side, values=SIDES, width=8, state="readonly").grid(row=6, column=1, sticky="w")
        ttk.Label(misc, text="Output type:").grid(row=7, column=0, sticky="w")
        self.var_otype = tk.StringVar(value="kicad")
        ttk.Combobox(misc, textvariable=self.var_otype, values=["kicad", "lib"], width=8, state="readonly").grid(row=7, column=1, sticky="w")
        ttk.Label(misc, text="Extra arguments:").grid(row=8, column=0, sticky="w")
        self.var_extra = tk.StringVar()
        ttk.Entry(misc, textvariable=self.var_extra, width=18).grid(row=8, column=1, sticky="w")

    def _toggle_group(self):
        st = "normal" if self.var_group.get() else "disabled"
        self.cb_grp_case.config(state=st)
        self.sp_spacer.config(state=st)

    def _build_actions_and_log(self):
        frame = ttk.Frame(self.root)
        frame.pack(fill="x", padx=10, pady=2)
        ttk.Button(frame, text="Generate from FILE", command=self.generate_from_file, width=24).pack(side="left", padx=4)
        ttk.Button(frame, text="Generate from PIN TABLE", command=self.generate_from_table, width=24).pack(side="left", padx=4)
        ttk.Button(frame, text="Save Settings", command=self._save_settings).pack(side="right", padx=4)

        self.log_text = tk.Text(self.root, height=9, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(2, 10))

    # ------------------------------------------------------------ settings
    def _save_settings(self):
        data = {k: v.get() for k, v in self._setting_vars().items()}
        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump(data, f, indent=1)
            self.log("Settings saved.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save settings: {e}")

    def _load_settings(self):
        try:
            with open(SETTINGS_FILE) as f:
                data = json.load(f)
            for k, var in self._setting_vars().items():
                if k in data:
                    try:
                        var.set(data[k])
                    except Exception:
                        pass
        except Exception:
            pass

    def _setting_vars(self):
        return {
            "name": self.var_name, "desig": self.var_desig,
            "group": self.var_group, "grp_case": self.var_grp_case,
            "spacer": self.var_spacer, "dist": self.var_dist,
            "dist_side": self.var_dist_side, "shape": self.var_shape,
            "overwrite": self.var_overwrite, "merge": self.var_merge,
            "bundle": self.var_bundle, "icase": self.var_icase,
            "fuzzy": self.var_fuzzy, "sort": self.var_sort,
            "side": self.var_side, "otype": self.var_otype,
            "extra": self.var_extra,
        }

    # ------------------------------------------------------------ logging
    def log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.root.update()

    # ---------------------------------------------------------- file mode
    def browse_input(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("All Files", "*.*")])
        if file_path:
            self.var_input.set(file_path)
            if not self.var_name.get():
                self.var_name.set(os.path.splitext(os.path.basename(file_path))[0])
            if not self.var_output.get():
                self.var_output.set(os.path.splitext(file_path)[0] + ".kicad_sym")

    def browse_output(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".kicad_sym",
            filetypes=[("KiCad Symbol Library", "*.kicad_sym"), ("Legacy KiCad Library", "*.lib"), ("All Files", "*.*")])
        if file_path:
            self.var_output.set(file_path)

    def preview_file(self):
        path = self.var_input.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Please select a valid input file first.")
            return
        try:
            df, dupes = read_and_clean(path)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        msg = f"Parsed {len(df)} pins, columns: {', '.join(df.columns)}\n\nFirst rows:\n"
        msg += df.head(8).to_string(index=False)
        if dupes:
            msg += f"\n\nWARNING - duplicate pin numbers: {', '.join(map(str, dupes))}"
        messagebox.showinfo("Preview / Validation", msg)

    # -------------------------------------------------------- pin editor
    def tree_values(self):
        return [list(self.tree.item(i, "values")) for i in self.tree.get_children()]

    def _refresh_count(self):
        self.lbl_count.config(text=f"{len(self.tree.get_children())} pins")

    def on_tree_select(self, _evt=None):
        sel = self.tree.selection()
        if not sel:
            return
        v = self.tree.item(sel[0], "values")
        self.ent_pin.delete(0, tk.END); self.ent_pin.insert(0, v[0])
        self.ent_name.delete(0, tk.END); self.ent_name.insert(0, v[1])
        self.cb_type.set(v[2] or "passive")
        self.cb_style.set(v[3] or "line")
        self.cb_side_pin.set(v[4] if len(v) > 4 else "")

    def add_pin(self):
        pin = self.ent_pin.get().strip()
        if not pin:
            messagebox.showerror("Error", "Pin number is required.")
            return
        self.tree.insert("", "end", values=(pin, self.ent_name.get().strip(),
                                            self.cb_type.get(), self.cb_style.get(),
                                            self.cb_side_pin.get()))
        self._refresh_count()

    def update_pin(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showerror("Error", "Select a row to update.")
            return
        self.tree.item(sel[0], values=(self.ent_pin.get().strip(), self.ent_name.get().strip(),
                                       self.cb_type.get(), self.cb_style.get(),
                                       self.cb_side_pin.get()))

    def delete_pins(self):
        for i in self.tree.selection():
            self.tree.delete(i)
        self._refresh_count()

    def move_pin(self, delta):
        sel = self.tree.selection()
        if len(sel) != 1:
            return
        idx = self.tree.index(sel[0])
        new_idx = idx + delta
        ids = list(self.tree.get_children())
        if 0 <= new_idx < len(ids):
            self.tree.move(sel[0], "", new_idx)

    def duplicate_pin(self):
        for i in self.tree.selection():
            self.tree.insert("", "end", values=self.tree.item(i, "values"))
        self._refresh_count()

    def sort_tree(self):
        rows = self.tree_values()

        def key(r):
            s = str(r[0])
            return (0, int(s)) if s.isdigit() else (1, s)

        for i in self.tree.get_children():
            self.tree.delete(i)
        for r in sorted(rows, key=key):
            self.tree.insert("", "end", values=r)

    def clear_tree(self, silent=False):
        if not silent and not messagebox.askyesno("Confirm", "Clear all pins from the table?"):
            return
        for i in self.tree.get_children():
            self.tree.delete(i)
        self._refresh_count()

    def import_to_tree(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All Files", "*.*")])
        if not path:
            return
        try:
            df, dupes = read_and_clean(path)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self.clear_tree(silent=True)
        for _, r in df.iterrows():
            self.tree.insert("", "end", values=(
                clean_val(r.get("Pin", "")), clean_val(r.get("Name", "")),
                clean_val(r.get("Type", "")) or "passive",
                clean_val(r.get("Style", "")) or "line",
                clean_val(r.get("Side", ""))))
        self._refresh_count()
        self.log(f"Imported {len(df)} pins from {os.path.basename(path)}")
        if dupes:
            self.log(f"WARNING: duplicate pin numbers: {', '.join(map(str, dupes))}")

    def export_tree_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Pin", "Name", "Type", "Style", "Side"])
            for r in self.tree_values():
                w.writerow(r)
        self.log(f"Table exported to {path}")

    # ---------------------------------------------------------- generate
    def _check_common(self, output_file):
        part_name = self.var_name.get().strip()
        ref_des = self.var_desig.get().strip()
        if not output_file:
            messagebox.showerror("Error", "Please select an output file location (File Mode tab).")
            return None
        if not part_name:
            messagebox.showerror("Error", "Please enter a Symbol Name.")
            return None
        return part_name, ref_des

    def generate_from_file(self):
        input_file = self.var_input.get()
        output_file = self.var_output.get()
        if not input_file or not os.path.exists(input_file):
            messagebox.showerror("Error", "Please select a valid input file.")
            return
        common = self._check_common(output_file)
        if not common:
            return
        try:
            df, dupes = read_and_clean(input_file)
        except Exception as e:
            messagebox.showerror("Data Error", str(e))
            return
        if dupes and not messagebox.askyesno(
                "Duplicate Pin Numbers",
                f"Duplicate pin numbers found: {', '.join(map(str, dupes))}\nContinue anyway?"):
            return
        self._run_pipeline(df, common, os.path.dirname(input_file))

    def generate_from_table(self):
        output_file = self.var_output.get()
        common = self._check_common(output_file)
        if not common:
            return
        rows = self.tree_values()
        if not rows:
            messagebox.showerror("Error", "The pin table is empty. Add pins or import a file.")
            return
        df = pd.DataFrame(rows, columns=["Pin", "Name", "Type", "Style", "Side"])
        pins = df["Pin"].astype(str)
        dupes = pins[pins.duplicated()].unique().tolist()
        if dupes and not messagebox.askyesno(
                "Duplicate Pin Numbers",
                f"Duplicate pin numbers found: {', '.join(map(str, dupes))}\nContinue anyway?"):
            return
        self._run_pipeline(df, common, os.getcwd())

    def _run_pipeline(self, df, common, work_dir):
        part_name, ref_des = common
        output_file = self.var_output.get()

        self.log_text.config(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state="disabled")

        self.log(f"Loaded {len(df)} pins. Applying layout options...")

        df = apply_grouping(df, self.var_group.get(), self.var_grp_case.get(),
                            max(0, self.var_spacer.get()))
        if self.var_group.get():
            self.log(f"Grouped by name (spacers: {self.var_spacer.get()}).")
        df = apply_symmetric(df, self.var_dist.get(), self.var_dist_side.get())
        if self.var_dist.get() != "None":
            self.log(f"Applied distribution: {self.var_dist.get()} starting '{self.var_dist_side.get()}'.")

        temp_csv = os.path.join(work_dir, f"{part_name}_kipart_ready.csv")
        try:
            write_kipart_csv(df, temp_csv, part_name, ref_des)
            self.log(f"Created clean CSV: {os.path.basename(temp_csv)}")
        except Exception as e:
            self.log(f"ERROR writing CSV: {e}")
            messagebox.showerror("Error", str(e))
            return

        cmd = ["kipart"]
        shape = self.var_shape.get()
        if "(-r)" in shape or "Rectangle" in shape:
            cmd.append("-r")
        if "(-c)" in shape or "Centered" in shape:
            cmd.append("-c")
        if self.var_overwrite.get() and not self.var_merge.get():
            cmd.append("-w")
        if self.var_merge.get():
            cmd.append("-m")
        if self.var_bundle.get():
            cmd.append("-b")
        if self.var_icase.get():
            cmd.append("-i")
        if self.var_fuzzy.get():
            cmd.append("-f")
        if self.var_group.get() and self.var_spacer.get() > 0:
            cmd.append("-z")  # ignore blank spacer rows silently
        if self.var_sort.get():
            cmd.extend(["-s", self.var_sort.get()])
        if self.var_side.get():
            cmd.extend(["--side", self.var_side.get()])
        if self.var_otype.get() and self.var_otype.get() != "kicad":
            cmd.extend(["-t", self.var_otype.get()])
        extra = self.var_extra.get().strip()
        if extra:
            cmd.extend(extra.split())
        cmd.extend(["-o", output_file, temp_csv])

        self.log(f"Executing: {' '.join(cmd)}")
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, shell=(os.name == "nt"))
            if result.stdout:
                self.log(result.stdout)
            if result.stderr:
                self.log(result.stderr)
            if result.returncode != 0:
                self.log(f"KiPart exited with error code {result.returncode}")
                messagebox.showerror("KiPart Error",
                                     f"KiPart failed (code {result.returncode}). Check the log.")
                return
        except FileNotFoundError:
            self.log("ERROR: 'kipart' command not found. Install with: pip install kipart")
            messagebox.showerror("Error", "kipart command not found.\nInstall it with 'pip install kipart'.")
            return
        except Exception as e:
            self.log(f"Execution ERROR: {e}")
            messagebox.showerror("Execution Error", str(e))
            return

        if shape == "Force square (post-process)":
            try:
                if make_square(output_file):
                    self.log("Post-processed body rectangle into a square.")
                else:
                    self.log("WARNING: no rectangle found to square-ify (did you enable a box shape?).")
            except Exception as e:
                self.log(f"Square post-process failed: {e}")

        self.log("Successfully generated KiCad symbol library!")
        messagebox.showinfo("Success", f"Symbol generated at:\n{output_file}")


if __name__ == "__main__":
    root = tk.Tk()
    app = KiPartGUI(root)
    root.mainloop()
