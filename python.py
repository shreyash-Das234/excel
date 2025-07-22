import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
from PIL import Image, ImageTk
from io import BytesIO
import base64

DEFAULT_LOGO_B64 = """R0lGODlhMAAwAIAAAP///wAAACH5BAEAAAAALAAAAAAwADAAAAIOhI+py+0Po5y02ouz3rwFADs="""


class ExcelToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ABG-Excel")
        self.root.geometry("1400x900")
        self.root.configure(bg="#232946")

        # ---------- State ----------
        self.files = []                   # [full_path, ...]
        self.all_sheets = {}              # {full_path: {sheet_name: DataFrame}}
        self.result_df = None             # last operation result
        self.current_preview_df = None    # what's shown in preview
        self.current_preview_file = None  # file path of previewed sheet (if sheet-based)
        self.current_preview_sheet = None # sheet name

        # ---------- UI setup ----------
        self._load_logo()
        self._set_style()
        self._create_menu()
        self._create_widgets()

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(
            self.root, textvariable=self.status_var, anchor="w", style="Status.TLabel"
        )
        self.status_bar.pack(side=tk.BOTTOM, fill="x")

    def _load_logo(self):
        try:
            img = Image.open("logo.png").resize((48, 48))
        except Exception:
            try:
                raw = base64.b64decode(DEFAULT_LOGO_B64)
                img = Image.open(BytesIO(raw)).resize((48, 48))
            except Exception:
                self.logo = None
                return
        self.logo = ImageTk.PhotoImage(img)
        self.root.iconphoto(False, self.logo)

    def _set_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background="#232946")
        style.configure("TLabel", background="#232946", foreground="#fffffe", font=("Segoe UI", 12))
        style.configure("TButton",
                        background="#eebbc3",
                        foreground="#232946",
                        font=("Segoe UI", 12, "bold"),
                        padding=6)
        style.map("TButton", background=[("active", "#b8c1ec")])
        style.configure("TLabelframe", background="#232946", foreground="#eebbc3", font=("Segoe UI", 13, "bold"))
        style.configure("TLabelframe.Label", background="#232946", foreground="#eebbc3", font=("Segoe UI", 13, "bold"))
        style.configure("Status.TLabel", background="#121629", foreground="#eebbc3", font=("Segoe UI", 10, "italic"))
        style.configure("Treeview", background="#ffffff", foreground="#232946", fieldbackground="#ffffff", rowheight=24)
        style.configure("Treeview.Heading", background="#eebbc3", foreground="#232946", font=("Segoe UI", 11, "bold"))


    def _create_menu(self):
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open Excel Files...", command=self.load_files)
        file_menu.add_command(label="Export Result...", command=self.export_result)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Full Preview", command=self.full_preview)
        menubar.add_cascade(label="View", menu=view_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)

    def _show_about(self):
        messagebox.showinfo(
            "About",
            "ABG-Excel v4.0\n\n"
            "Load multiple Excel files; preview sheets; VLOOKUP & compare with unique differences; "
            "concatenate columns; find unique values; export results."
        )

    def _create_widgets(self):

        file_frame = ttk.Labelframe(self.root, text="Step 1: Load Excel Files", padding=10)
        file_frame.pack(fill="x", padx=15, pady=(15, 5))

        ttk.Button(file_frame, text="Add Files", command=self.load_files).pack(side="left", padx=5)
        ttk.Button(file_frame, text="Clear All", command=self.clear_files).pack(side="left", padx=5)

        self.file_listbox = tk.Listbox(
            file_frame, height=4, bg="#232946", fg="#eebbc3",
            selectbackground="#b8c1ec", selectforeground="#232946",
            font=("Segoe UI", 11), highlightthickness=0
        )
        self.file_listbox.pack(side="left", fill="both", expand=True, padx=(10, 0))
        f_scroll = ttk.Scrollbar(file_frame, orient="vertical", command=self.file_listbox.yview)
        f_scroll.pack(side="right", fill="y")
        self.file_listbox.config(yscrollcommand=f_scroll.set)


        sel_frame = ttk.Labelframe(self.root, text="Step 2: Select Sheet to Preview", padding=10)
        sel_frame.pack(fill="x", padx=15, pady=5)

        ttk.Label(sel_frame, text="File:").pack(side="left")
        self.preview_file_combo = ttk.Combobox(sel_frame, state="readonly", width=40)
        self.preview_file_combo.pack(side="left", padx=5)

        ttk.Label(sel_frame, text="Sheet:").pack(side="left")
        self.preview_sheet_combo = ttk.Combobox(sel_frame, state="readonly", width=30)
        self.preview_sheet_combo.pack(side="left", padx=5)

        ttk.Button(sel_frame, text="Load Sheet", command=self._preview_selected_sheet).pack(side="left", padx=10)

        self.preview_file_combo.bind("<<ComboboxSelected>>", self._update_preview_sheet_combo)


        op_frame = ttk.Labelframe(self.root, text="Step 3: Operations", padding=10)
        op_frame.pack(fill="x", padx=15, pady=5)

        ttk.Button(op_frame, text="VLOOKUP & Compare", command=self.vlookup).pack(side="left", padx=5)
        ttk.Button(op_frame, text="Compare Columns", command=self.compare_columns).pack(side="left", padx=5)
        ttk.Button(op_frame, text="Concatenate Columns", command=self.concat_columns).pack(side="left", padx=5)
        ttk.Button(op_frame, text="Find Unique Values", command=self.find_unique_values).pack(side="left", padx=5)
        ttk.Button(op_frame, text="Export Result", command=self.export_result).pack(side="right", padx=5)


        preview_frame = ttk.Labelframe(self.root, text="Step 4: Preview (first 100 rows)", padding=10)
        preview_frame.pack(fill="both", expand=True, padx=15, pady=(5, 15))

        self.preview_container = ttk.Frame(preview_frame)
        self.preview_container.pack(fill="both", expand=True)

        self._build_preview_tree()


    def load_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
        if not paths:
            return

        new_files = 0
        for p in paths:
            if p in self.files:
                continue
            try:
                xl = pd.ExcelFile(p)
                sheets_dict = {}
                for s in xl.sheet_names:
                    sheets_dict[s] = xl.parse(s)
                self.all_sheets[p] = sheets_dict
                self.files.append(p)
                new_files += 1
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read {os.path.basename(p)}:\n{e}")

      
        self._refresh_file_list()
        self._refresh_preview_file_combo()

        if self.files and not self.current_preview_df:
       
            first_fp = self.files[0]
            first_sh = next(iter(self.all_sheets[first_fp]))
            self._set_current_preview(first_fp, first_sh)

        self.set_status(f"Loaded {new_files} new file(s). Total: {len(self.files)}.")

    def clear_files(self):
        self.files.clear()
        self.all_sheets.clear()
        self.result_df = None
        self.current_preview_df = None
        self.current_preview_file = None
        self.current_preview_sheet = None
        self._refresh_file_list()
        self._refresh_preview_file_combo()
        self._update_preview_tree(pd.DataFrame())
        self.set_status("Cleared all files.")

    def _refresh_file_list(self):
        self.file_listbox.delete(0, tk.END)
        for p in self.files:
            self.file_listbox.insert(tk.END, os.path.basename(p))

    def _refresh_preview_file_combo(self):
        self.preview_file_combo['values'] = [os.path.basename(p) for p in self.files]
        if self.files:
            self.preview_file_combo.current(0)
            self._update_preview_sheet_combo()

    def _update_preview_sheet_combo(self, event=None):
        if not self.files:
            self.preview_sheet_combo['values'] = []
            return
        idx = self.preview_file_combo.current()
        fp = self.files[idx]
        sheets = list(self.all_sheets[fp].keys())
        self.preview_sheet_combo['values'] = sheets
        if sheets:
            self.preview_sheet_combo.current(0)

    def _preview_selected_sheet(self):
        if not self.files:
            return
        fp = self.files[self.preview_file_combo.current()]
        sh = self.preview_sheet_combo.get()
        self._set_current_preview(fp, sh)

    def _set_current_preview(self, file_path, sheet_name):
        df = self.all_sheets[file_path][sheet_name]
        self.current_preview_df = df
        self.current_preview_file = file_path
        self.current_preview_sheet = sheet_name
        self.result_df = df.copy()  # treat current as baseline result
        self._update_preview_tree(df)
        self.set_status(f"Previewing: {os.path.basename(file_path)} / {sheet_name}")


    def _build_preview_tree(self):
        # remove existing if any
        for w in self.preview_container.winfo_children():
            w.destroy()

        self.preview_tree = ttk.Treeview(self.preview_container, show="headings")
        self.preview_tree.pack(fill="both", expand=True)

        vsb = ttk.Scrollbar(self.preview_container, orient="vertical", command=self.preview_tree.yview)
        vsb.pack(side="right", fill="y")
        self.preview_tree.configure(yscrollcommand=vsb.set)

        hsb = ttk.Scrollbar(self.preview_container, orient="horizontal", command=self.preview_tree.xview)
        hsb.pack(side="bottom", fill="x")
        self.preview_tree.configure(xscrollcommand=hsb.set)

        self._update_preview_tree(pd.DataFrame())

    def _update_preview_tree(self, df):
        self.preview_tree.delete(*self.preview_tree.get_children())

        for c in self.preview_tree["columns"]:
            self.preview_tree.heading(c, text="")
        if df is None or df.empty:
            self.preview_tree["columns"] = []
            return
        cols = list(df.columns)
        self.preview_tree["columns"] = cols
        for c in cols:
            self.preview_tree.heading(c, text=c)
            self.preview_tree.column(c, width=120, anchor="w")
        for _, row in df.head(100).iterrows():  # limit preview
            self.preview_tree.insert("", "end", values=list(row))

    def vlookup(self):
        if len(self.files) < 2:
            messagebox.showwarning("Warning", "Load at least two files.")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("VLOOKUP & Compare")
        dlg.geometry("600x520")
        dlg.configure(bg="#232946")

        file_names = [os.path.basename(p) for p in self.files]


        ttk.Label(dlg, text="Main File:", style="TLabel").pack(pady=(10, 0))
        main_file_combo = ttk.Combobox(dlg, values=file_names, state="readonly")
        main_file_combo.pack(fill="x", padx=15, pady=5)
        if self.current_preview_file:
            main_file_combo.current(self.files.index(self.current_preview_file))
        else:
            main_file_combo.current(0)

        ttk.Label(dlg, text="Main Sheet:", style="TLabel").pack(pady=(5, 0))
        main_sheet_combo = ttk.Combobox(dlg, state="readonly")
        main_sheet_combo.pack(fill="x", padx=15, pady=5)

        ttk.Label(dlg, text="Key Column (Main):", style="TLabel").pack(pady=(5, 0))
        main_key_combo = ttk.Combobox(dlg, state="readonly")
        main_key_combo.pack(fill="x", padx=15, pady=5)


        ttk.Label(dlg, text="Lookup File:", style="TLabel").pack(pady=(15, 0))
        lookup_file_combo = ttk.Combobox(dlg, values=file_names, state="readonly")
        lookup_file_combo.pack(fill="x", padx=15, pady=5)

        if len(self.files) > 1:
            l_default = 1 if main_file_combo.current() == 0 else 0
            lookup_file_combo.current(l_default)
        else:
            lookup_file_combo.current(0)

        ttk.Label(dlg, text="Lookup Sheet:", style="TLabel").pack(pady=(5, 0))
        lookup_sheet_combo = ttk.Combobox(dlg, state="readonly")
        lookup_sheet_combo.pack(fill="x", padx=15, pady=5)

        ttk.Label(dlg, text="Key Column (Lookup):", style="TLabel").pack(pady=(5, 0))
        lookup_key_combo = ttk.Combobox(dlg, state="readonly")
        lookup_key_combo.pack(fill="x", padx=15, pady=5)

        ttk.Label(dlg, text="Columns to Fetch (Lookup):", style="TLabel").pack(pady=(5, 0))
        lookup_cols_list = tk.Listbox(
            dlg, selectmode="multiple", height=6,
            bg="#232946", fg="#eebbc3", selectbackground="#b8c1ec", selectforeground="#232946"
        )
        lookup_cols_list.pack(fill="x", padx=15, pady=5)


        unique_var = tk.BooleanVar()
        tk.Checkbutton(
            dlg, text="Show Only Unique Differences (A vs B)", variable=unique_var,
            bg="#232946", fg="#eebbc3", selectcolor="#232946", activebackground="#232946"
        ).pack(pady=8)

        def update_main_sheets(_e=None):
            fp = self.files[main_file_combo.current()]
            sheets = list(self.all_sheets[fp].keys())
            main_sheet_combo['values'] = sheets

            if self.current_preview_file == fp and self.current_preview_sheet in sheets:
                main_sheet_combo.set(self.current_preview_sheet)
            else:
                main_sheet_combo.current(0)
            update_main_cols()

        def update_main_cols(_e=None):
            fp = self.files[main_file_combo.current()]
            sh = main_sheet_combo.get()
            df = self.all_sheets[fp][sh]
            main_key_combo['values'] = list(df.columns)
            if df.columns.size:
                main_key_combo.current(0)

        def update_lookup_sheets(_e=None):
            fp = self.files[lookup_file_combo.current()]
            sheets = list(self.all_sheets[fp].keys())
            lookup_sheet_combo['values'] = sheets
            lookup_sheet_combo.current(0)
            update_lookup_cols()

        def update_lookup_cols(_e=None):
            fp = self.files[lookup_file_combo.current()]
            sh = lookup_sheet_combo.get()
            df = self.all_sheets[fp][sh]
            cols = list(df.columns)
            lookup_key_combo['values'] = cols
            if cols:
                lookup_key_combo.current(0)
            lookup_cols_list.delete(0, tk.END)
            for c in cols:
                lookup_cols_list.insert(tk.END, c)

        main_file_combo.bind("<<ComboboxSelected>>", update_main_sheets)
        main_sheet_combo.bind("<<ComboboxSelected>>", update_main_cols)
        lookup_file_combo.bind("<<ComboboxSelected>>", update_lookup_sheets)
        lookup_sheet_combo.bind("<<ComboboxSelected>>", update_lookup_cols)

        update_main_sheets()
        update_lookup_sheets()

        def perform_vlookup():
            try:

                fp_main = self.files[main_file_combo.current()]
                fp_lookup = self.files[lookup_file_combo.current()]
                sh_main = main_sheet_combo.get()
                sh_lookup = lookup_sheet_combo.get()
                key_main = main_key_combo.get()
                key_lookup = lookup_key_combo.get()
                sel_idx = lookup_cols_list.curselection()
                fetch_cols = [lookup_cols_list.get(i) for i in sel_idx]

                if not key_main or not key_lookup:
                    messagebox.showwarning("Warning", "Select key columns.")
                    return

                df_main = self.all_sheets[fp_main][sh_main]
                df_lookup = self.all_sheets[fp_lookup][sh_lookup]

                if unique_var.get():

                    set_main = set(df_main[key_main].dropna().astype(str))
                    set_lookup = set(df_lookup[key_lookup].dropna().astype(str))
                    only_in_main = sorted(set_main - set_lookup)
                    only_in_lookup = sorted(set_lookup - set_main)
                    out_df = pd.DataFrame({
                        f"Only in {os.path.basename(fp_main)}::{sh_main}": pd.Series(only_in_main),
                        f"Only in {os.path.basename(fp_lookup)}::{sh_lookup}": pd.Series(only_in_lookup)
                    })
                else:

                    if not fetch_cols:

                        fetch_cols = [c for c in df_lookup.columns if c != key_lookup]

                    cols = [key_lookup] + [c for c in fetch_cols if c != key_lookup]
                    out_df = pd.merge(
                        df_main,
                        df_lookup[cols],
                        left_on=key_main,
                        right_on=key_lookup,
                        how="left",
                        suffixes=('', '_lk')
                    )

                    if key_main != key_lookup and key_lookup in out_df.columns:
                        out_df.drop(columns=[key_lookup], inplace=True)

                self.result_df = out_df
                self.current_preview_df = out_df
                self._update_preview_tree(out_df)
                self.set_status("VLOOKUP complete.")
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"VLOOKUP failed:\n{e}")

        ttk.Button(dlg, text="Run VLOOKUP", command=perform_vlookup).pack(pady=15)


    def compare_columns(self):
        if len(self.files) < 2:
            messagebox.showwarning("Warning", "Load at least two files.")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Compare Columns")
        dlg.geometry("520x420")
        dlg.configure(bg="#232946")

        fnames = [os.path.basename(p) for p in self.files]

        # A
        ttk.Label(dlg, text="Sheet A - File:", style="TLabel").pack(pady=(10, 0))
        file_a_combo = ttk.Combobox(dlg, values=fnames, state="readonly")
        file_a_combo.pack(fill="x", padx=15, pady=5)
        file_a_combo.current(0)

        ttk.Label(dlg, text="Sheet A - Sheet:", style="TLabel").pack()
        sheet_a_combo = ttk.Combobox(dlg, state="readonly")
        sheet_a_combo.pack(fill="x", padx=15, pady=5)

        ttk.Label(dlg, text="Sheet A - Column:", style="TLabel").pack()
        col_a_combo = ttk.Combobox(dlg, state="readonly")
        col_a_combo.pack(fill="x", padx=15, pady=5)

        # B
        ttk.Label(dlg, text="Sheet B - File:", style="TLabel").pack(pady=(15, 0))
        file_b_combo = ttk.Combobox(dlg, values=fnames, state="readonly")
        file_b_combo.pack(fill="x", padx=15, pady=5)
        file_b_combo.current(1 if len(self.files) > 1 else 0)

        ttk.Label(dlg, text="Sheet B - Sheet:", style="TLabel").pack()
        sheet_b_combo = ttk.Combobox(dlg, state="readonly")
        sheet_b_combo.pack(fill="x", padx=15, pady=5)

        ttk.Label(dlg, text="Sheet B - Column:", style="TLabel").pack()
        col_b_combo = ttk.Combobox(dlg, state="readonly")
        col_b_combo.pack(fill="x", padx=15, pady=5)

        def upd_a(_e=None):
            fp = self.files[file_a_combo.current()]
            sheets = list(self.all_sheets[fp].keys())
            sheet_a_combo['values'] = sheets
            sheet_a_combo.current(0)
            upd_a_cols()

        def upd_a_cols(_e=None):
            fp = self.files[file_a_combo.current()]
            sh = sheet_a_combo.get()
            df = self.all_sheets[fp][sh]
            col_a_combo['values'] = list(df.columns)
            if df.columns.size:
                col_a_combo.current(0)

        def upd_b(_e=None):
            fp = self.files[file_b_combo.current()]
            sheets = list(self.all_sheets[fp].keys())
            sheet_b_combo['values'] = sheets
            sheet_b_combo.current(0)
            upd_b_cols()

        def upd_b_cols(_e=None):
            fp = self.files[file_b_combo.current()]
            sh = sheet_b_combo.get()
            df = self.all_sheets[fp][sh]
            col_b_combo['values'] = list(df.columns)
            if df.columns.size:
                col_b_combo.current(0)

        file_a_combo.bind("<<ComboboxSelected>>", upd_a)
        sheet_a_combo.bind("<<ComboboxSelected>>", upd_a_cols)
        file_b_combo.bind("<<ComboboxSelected>>", upd_b)
        sheet_b_combo.bind("<<ComboboxSelected>>", upd_b_cols)

        upd_a(); upd_b()

        def run_compare():
            try:
                fp_a = self.files[file_a_combo.current()]
                sh_a = sheet_a_combo.get()
                col_a = col_a_combo.get()
                fp_b = self.files[file_b_combo.current()]
                sh_b = sheet_b_combo.get()
                col_b = col_b_combo.get()

                df_a = self.all_sheets[fp_a][sh_a]
                df_b = self.all_sheets[fp_b][sh_b]

                set_a = set(df_a[col_a].dropna().astype(str))
                set_b = set(df_b[col_b].dropna().astype(str))

                only_a = sorted(set_a - set_b)
                only_b = sorted(set_b - set_a)

                out_df = pd.DataFrame({
                    f"Only in {os.path.basename(fp_a)}::{sh_a}": pd.Series(only_a),
                    f"Only in {os.path.basename(fp_b)}::{sh_b}": pd.Series(only_b)
                })
                self.result_df = out_df
                self.current_preview_df = out_df
                self._update_preview_tree(out_df)
                self.set_status("Column comparison complete.")
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Comparison failed:\n{e}")

        ttk.Button(dlg, text="Compare", command=run_compare).pack(pady=15)

    def find_unique_values(self):
        if not self.files:
            messagebox.showwarning("Warning", "Load files first.")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Find Unique Values")
        dlg.geometry("420x300")
        dlg.configure(bg="#232946")

        fnames = [os.path.basename(p) for p in self.files]

        ttk.Label(dlg, text="File:", style="TLabel").pack(pady=(10, 0))
        file_combo = ttk.Combobox(dlg, values=fnames, state="readonly")
        file_combo.pack(fill="x", padx=15, pady=5)
        file_combo.current(0)

        ttk.Label(dlg, text="Sheet:", style="TLabel").pack()
        sheet_combo = ttk.Combobox(dlg, state="readonly")
        sheet_combo.pack(fill="x", padx=15, pady=5)

        ttk.Label(dlg, text="Column:", style="TLabel").pack()
        col_combo = ttk.Combobox(dlg, state="readonly")
        col_combo.pack(fill="x", padx=15, pady=5)

        def upd_s(_e=None):
            fp = self.files[file_combo.current()]
            sheets = list(self.all_sheets[fp].keys())
            sheet_combo['values'] = sheets
            sheet_combo.current(0)
            upd_c()

        def upd_c(_e=None):
            fp = self.files[file_combo.current()]
            sh = sheet_combo.get()
            df = self.all_sheets[fp][sh]
            cols = list(df.columns)
            col_combo['values'] = cols
            if cols:
                col_combo.current(0)

        file_combo.bind("<<ComboboxSelected>>", upd_s)
        sheet_combo.bind("<<ComboboxSelected>>", upd_c)

        upd_s()

        def run_unique():
            try:
                fp = self.files[file_combo.current()]
                sh = sheet_combo.get()
                col = col_combo.get()
                df = self.all_sheets[fp][sh]
                uniq = pd.DataFrame(df[col].dropna().astype(str).unique(), columns=[f"Unique_{col}"])
                self.result_df = uniq
                self.current_preview_df = uniq
                self._update_preview_tree(uniq)
                self.set_status(f"Found {len(uniq)} unique value(s).")
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Unique extraction failed:\n{e}")

        ttk.Button(dlg, text="Find Unique", command=run_unique).pack(pady=15)

    def concat_columns(self):
        if not self.files:
            messagebox.showwarning("Warning", "Load files first.")
            return
        self._concat_dialog()

    def _concat_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Concatenate Columns")
        dlg.geometry("500x500")
        dlg.configure(bg="#232946")

        fnames = [os.path.basename(p) for p in self.files]

        ttk.Label(dlg, text="File:", style="TLabel").pack(pady=(10, 0))
        file_combo = ttk.Combobox(dlg, values=fnames, state="readonly")
        file_combo.pack(fill="x", padx=15, pady=5)
        if self.current_preview_file:
            file_combo.current(self.files.index(self.current_preview_file))
        else:
            file_combo.current(0)

        ttk.Label(dlg, text="Sheet:", style="TLabel").pack()
        sheet_combo = ttk.Combobox(dlg, state="readonly")
        sheet_combo.pack(fill="x", padx=15, pady=5)

        cols_list = tk.Listbox(
            dlg, selectmode="multiple", height=8,
            bg="#232946", fg="#eebbc3", selectbackground="#b8c1ec", selectforeground="#232946"
        )
        cols_list.pack(fill="x", padx=15, pady=5)

        ttk.Label(dlg, text="Separator:", style="TLabel").pack(pady=(10, 0))
        sep_entry = ttk.Entry(dlg)
        sep_entry.insert(0, " ")
        sep_entry.pack(fill="x", padx=15, pady=5)

        ttk.Label(dlg, text="Prefix:", style="TLabel").pack(pady=(5, 0))
        pre_entry = ttk.Entry(dlg)
        pre_entry.pack(fill="x", padx=15, pady=5)

        ttk.Label(dlg, text="Suffix:", style="TLabel").pack(pady=(5, 0))
        suf_entry = ttk.Entry(dlg)
        suf_entry.pack(fill="x", padx=15, pady=5)

        ttk.Label(dlg, text="Result Column Name:", style="TLabel").pack(pady=(5, 0))
        res_entry = ttk.Entry(dlg)
        res_entry.insert(0, "Concatenated")
        res_entry.pack(fill="x", padx=15, pady=5)

        def upd_sh(_e=None):
            fp = self.files[file_combo.current()]
            sheets = list(self.all_sheets[fp].keys())
            sheet_combo['values'] = sheets
            if self.current_preview_file == fp and self.current_preview_sheet in sheets:
                sheet_combo.set(self.current_preview_sheet)
            else:
                sheet_combo.current(0)
            upd_cols()

        def upd_cols(_e=None):
            fp = self.files[file_combo.current()]
            sh = sheet_combo.get()
            df = self.all_sheets[fp][sh]
            cols_list.delete(0, tk.END)
            for c in df.columns:
                cols_list.insert(tk.END, c)

        file_combo.bind("<<ComboboxSelected>>", upd_sh)
        sheet_combo.bind("<<ComboboxSelected>>", upd_cols)

        upd_sh()

        def preview_concat():
            sel_idx = cols_list.curselection()
            if len(sel_idx) < 2:
                messagebox.showwarning("Warning", "Select at least two columns.")
                return
            cols = [cols_list.get(i) for i in sel_idx]
            fp = self.files[file_combo.current()]
            sh = sheet_combo.get()
            df = self.all_sheets[fp][sh]
            sep = sep_entry.get()
            pre = pre_entry.get()
            suf = suf_entry.get()
            res = res_entry.get().strip() or "Concatenated"
            try:
                preview_df = df.copy()
                preview_df[res] = pre + df[cols].fillna("").astype(str).agg(sep.join, axis=1) + suf
            except Exception as e:
                messagebox.showerror("Error", f"Preview failed:\n{e}")
                return
            self._show_concat_preview_window(preview_df, fp, sh, cols, sep, pre, suf, res)

        ttk.Button(dlg, text="Preview Result (Full Screen)", command=preview_concat).pack(pady=15)

    def _show_concat_preview_window(self, preview_df, fp, sh, cols, sep, pre, suf, res_name):
        win = tk.Toplevel(self.root)
        win.title("Concatenation Preview")
        try:
            win.state('zoomed')
        except Exception:
            win.geometry("1200x700")

        lbl = ttk.Label(win, text=f"Preview: {os.path.basename(fp)} / {sh} → {res_name}", style="TLabel")
        lbl.pack(pady=5)

        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        tree = ttk.Treeview(frame, show="headings")
        tree.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(win, orient="horizontal", command=tree.xview)
        hsb.pack(fill="x")
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree["columns"] = list(preview_df.columns)
        for c in preview_df.columns:
            tree.heading(c, text=c)
            tree.column(c, width=150, anchor="w")

        for _, row in preview_df.head(1000).iterrows():
            tree.insert("", "end", values=list(row))

        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill="x", pady=10, padx=10)

        def apply_concat():
            df = self.all_sheets[fp][sh].copy()
            try:
                df[res_name] = pre + df[cols].fillna("").astype(str).agg(sep.join, axis=1) + suf
            except Exception as e:
                messagebox.showerror("Error", f"Apply failed:\n{e}")
                return
            # Persist
            self.all_sheets[fp][sh] = df
            if self.current_preview_file == fp and self.current_preview_sheet == sh:
                self.current_preview_df = df
                self.result_df = df
                self._update_preview_tree(df)
            self.set_status(f"Concatenated → {res_name}")
            messagebox.showinfo("Success", f"Column '{res_name}' added.")
            win.destroy()

        def download_concat():
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
                title="Save Concatenated Result As"
            )
            if not path:
                return
            try:
                if path.lower().endswith(".csv"):
                    preview_df.to_csv(path, index=False)
                else:
                    preview_df.to_excel(path, index=False)
                messagebox.showinfo("Saved", f"File saved:\n{path}")
            except Exception as e:
                messagebox.showerror("Error", f"Save failed:\n{e}")

        ttk.Button(btn_frame, text="Apply to Sheet", command=apply_concat).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="Download Excel", command=download_concat).pack(side="right", padx=5)

    def full_preview(self):
        df = self.result_df if self.result_df is not None else self.current_preview_df
        if df is None:
            messagebox.showwarning("Warning", "Nothing to preview.")
            return

        win = tk.Toplevel(self.root)
        win.title("Full Preview")
        try:
            win.state('zoomed')
        except Exception:
            win.geometry("1200x700")

        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        tree = ttk.Treeview(frame, show="headings")
        tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(win, orient="horizontal", command=tree.xview)
        hsb.pack(fill="x")
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree["columns"] = list(df.columns)
        for c in df.columns:
            tree.heading(c, text=c)
            tree.column(c, width=150, anchor="w")

        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row))

        ttk.Button(win, text="Download", command=self.export_result).pack(pady=8)


    def export_result(self):
        df = self.result_df if self.result_df is not None else self.current_preview_df
        if df is None:
            messagebox.showwarning("Warning", "No result data to export.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
            title="Save Result As"
        )
        if not path:
            return

        try:
            if path.lower().endswith(".csv"):
                df.to_csv(path, index=False)
            else:
                df.to_excel(path, index=False)
            messagebox.showinfo("Export Complete", f"File saved:\n{path}")
            self.set_status(f"Exported to {path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed:\n{e}")


    def set_status(self, msg):
        self.status_var.set(msg)
        self.root.update_idletasks()

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToolApp(root)
    root.mainloop()
