import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from PIL import Image, ImageTk
import threading
import openpyxl
import webbrowser
from io import BytesIO
import base64

class ExcelToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ABG-Excel")
        self.root.geometry("1400x900")
        self.root.configure(bg="#232946")
        self.files = []
        self.dfs = []
        self.sheet_names = []
        self.selected_columns = []
        self.result_df = None
        self.current_preview_df = None

        # App icon/logo
        try:
            logo_img = Image.open("logo.png").resize((48, 48))
            self.logo = ImageTk.PhotoImage(logo_img)
            self.root.iconphoto(False, self.logo)
        except Exception:
            # Use a default logo if custom logo not found
            try:
                logo_data = base64.b64decode("""R0lGODlhMAAwAIAAAP///wAAACH5BAEAAAAALAAAAAAwADAAAAIOhI+py+0Po5y02ouz3rwFADs=""")
                logo_img = Image.open(BytesIO(logo_data)).resize((48, 48))
                self.logo = ImageTk.PhotoImage(logo_img)
                self.root.iconphoto(False, self.logo)
            except Exception:
                self.logo = None

        # Style
        self.set_style()

        # Menu bar
        self.create_menu()

        # UI Elements
        self.create_widgets()

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, anchor="w", style="Status.TLabel")
        self.status_bar.pack(side=tk.BOTTOM, fill="x")

    def set_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TLabel", background="#232946", foreground="#fffffe", font=("Segoe UI", 12))
        style.configure("TButton",
                      background="#eebbc3",
                      foreground="#232946",
                      font=("Segoe UI", 12, "bold"),
                      padding=10,
                      borderwidth=0,
                      relief="flat")
        style.map("TButton",
                background=[("active", "#b8c1ec"), ("pressed", "#b8c1ec")],
                relief=[("pressed", "flat"), ("!pressed", "flat")])
        style.configure("TFrame", background="#232946")
        style.configure("TLabelframe", background="#232946", foreground="#eebbc3", font=("Segoe UI", 13, "bold"))
        style.configure("TLabelframe.Label", background="#232946", foreground="#eebbc3", font=("Segoe UI", 13, "bold"))
        style.configure("Treeview", background="#ffffff", foreground="#232946", fieldbackground="#ffffff", 
                       font=("Arial", 10), borderwidth=0, rowheight=25)
        style.configure("Treeview.Heading", background="#eebbc3", foreground="#232946", 
                       font=("Segoe UI", 11, "bold"), padding=(10,5))
        style.map("Treeview", background=[("selected", "#b8c1ec")], foreground=[("selected", "#232946")])
        style.configure("Status.TLabel", background="#121629", foreground="#eebbc3", font=("Segoe UI", 10, "italic"))
        style.configure("Preview.TFrame", background="#ffffff", borderwidth=1, relief="solid")
        style.configure("Preview.TLabel", background="#ffffff", foreground="#232946", font=("Arial", 10))

    def create_menu(self):
        menubar = tk.Menu(self.root)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open Excel Files...", command=self.load_files)
        file_menu.add_command(label="Export Result...", command=self.export_result)
        file_menu.add_command(label="Download as Excel", command=self.download_result)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Full Preview", command=lambda: self.show_full_preview())
        menubar.add_cascade(label="View", menu=view_menu)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Documentation", command=self.show_docs)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)

    def show_about(self):
        about_text = """Excel Multi-Sheet Comparator & VLOOKUP Tool
Version 2.0
© 2025 ABG Analytics

Features:
- Multiple Excel file comparison
- Sheet concatenation
- VLOOKUP functionality
- Google Sheets-like preview
- Direct download options"""
        messagebox.showinfo("About", about_text)

    def show_docs(self):
        docs_url = "https://docs.example.com/excel-tool"
        try:
            webbrowser.open_new(docs_url)
        except:
            messagebox.showinfo("Documentation", "Online documentation: " + docs_url)

    def create_widgets(self):
        # Main container with scroll
        main_container = ttk.Frame(self.root)
        main_container.pack(fill="both", expand=True)
        
        # Canvas and scrollbars
        canvas = tk.Canvas(main_container, bg="#232946", highlightthickness=0)
        scroll_y = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        scroll_x = ttk.Scrollbar(main_container, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        canvas.pack(side="left", fill="both", expand=True)
        
        # Frame inside canvas
        self.main_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        
        # Bind scroll events
        self.main_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units")))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # Header
        header_frame = ttk.Frame(self.main_frame, style="TFrame")
        header_frame.pack(fill="x", padx=20, pady=(15, 5))
        
        if self.logo:
            logo_label = ttk.Label(header_frame, image=self.logo, style="TLabel")
            logo_label.pack(side="left", padx=(0, 15))
        
        title_frame = ttk.Frame(header_frame, style="TFrame")
        title_frame.pack(side="left", fill="y")
        
        title_label = ttk.Label(title_frame, text="Excel Multi-Sheet Tool", 
                              font=("Segoe UI", 20, "bold"), style="TLabel")
        title_label.pack(anchor="w")
        
        subtitle_label = ttk.Label(title_frame, text="Compare, merge and analyze Excel sheets", 
                                 font=("Segoe UI", 12), style="TLabel")
        subtitle_label.pack(anchor="w")
        
        # Quick action buttons
        quick_btn_frame = ttk.Frame(header_frame, style="TFrame")
        quick_btn_frame.pack(side="right", padx=10)
        
        ttk.Button(quick_btn_frame, text="New", command=self.clear_files, 
                 style="TButton", width=8).pack(side="left", padx=5)
        ttk.Button(quick_btn_frame, text="Open", command=self.load_files, 
                 style="TButton", width=8).pack(side="left", padx=5)
        ttk.Button(quick_btn_frame, text="Save As", command=self.export_result, 
                 style="TButton", width=8).pack(side="left", padx=5)

        # File selection
        file_frame = ttk.Labelframe(self.main_frame, text="Step 1: Select Excel Files", padding=15)
        file_frame.pack(fill="x", padx=20, pady=10)
        
        btn_frame = ttk.Frame(file_frame, style="TFrame")
        btn_frame.pack(side="left", fill="y", padx=(0, 15))
        
        ttk.Button(btn_frame, text="Add Files", command=self.load_files, 
                 style="TButton").pack(fill="x", pady=3)
        ttk.Button(btn_frame, text="Clear All", command=self.clear_files, 
                 style="TButton").pack(fill="x", pady=3)
        
        self.file_listbox = tk.Listbox(file_frame, height=4, bg="#232946", fg="#eebbc3", 
                                     font=("Segoe UI", 11), selectbackground="#b8c1ec", 
                                     selectforeground="#232946", highlightthickness=0)
        self.file_listbox.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        scroll = ttk.Scrollbar(file_frame, orient="vertical", command=self.file_listbox.yview)
        scroll.pack(side="right", fill="y")
        self.file_listbox.config(yscrollcommand=scroll.set)

        # Sheet and column selection
        select_frame = ttk.Frame(self.main_frame, style="TFrame")
        select_frame.pack(fill="x", padx=20, pady=10)
        
        # Sheet selection
        sheet_frame = ttk.Labelframe(select_frame, text="Step 2: Select Sheet", padding=15)
        sheet_frame.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ttk.Label(sheet_frame, text="Sheet Name:", style="TLabel").pack(anchor="w")
        self.sheet_combo = ttk.Combobox(sheet_frame, state="readonly", font=("Segoe UI", 11))
        self.sheet_combo.pack(fill="x", pady=5)
        ttk.Button(sheet_frame, text="Load Sheet Data", command=self.load_sheet, 
                 style="TButton").pack(fill="x", pady=5)

        # Column selection
        col_frame = ttk.Labelframe(select_frame, text="Step 3: Select Columns", padding=15)
        col_frame.pack(side="left", fill="x", expand=True)
        
        ttk.Label(col_frame, text="Available Columns:", style="TLabel").pack(anchor="w")
        
        col_select_frame = ttk.Frame(col_frame, style="TFrame")
        col_select_frame.pack(fill="both", expand=True)
        
        self.col_listbox = tk.Listbox(col_select_frame, selectmode=tk.MULTIPLE, height=6, 
                                    bg="#232946", fg="#eebbc3", font=("Segoe UI", 11), 
                                    selectbackground="#b8c1ec", selectforeground="#232946")
        self.col_listbox.pack(side="left", fill="both", expand=True)
        
        col_scroll = ttk.Scrollbar(col_select_frame, orient="vertical", command=self.col_listbox.yview)
        col_scroll.pack(side="right", fill="y")
        self.col_listbox.config(yscrollcommand=col_scroll.set)

        # Operations
        op_frame = ttk.Labelframe(self.main_frame, text="Step 4: Select Operation", padding=15)
        op_frame.pack(fill="x", padx=20, pady=10)
        
        btn_grid = ttk.Frame(op_frame, style="TFrame")
        btn_grid.pack(fill="x")
        
        ttk.Button(btn_grid, text="Concatenate Columns", command=self.concat_columns, 
                 style="TButton").grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ttk.Button(btn_grid, text="Find Unique Rows", command=self.find_unique, 
                 style="TButton").grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(btn_grid, text="VLOOKUP", command=self.vlookup, 
                 style="TButton").grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        ttk.Button(btn_grid, text="Merge Sheets", command=self.merge_sheets, 
                 style="TButton").grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        
        btn_grid.columnconfigure(0, weight=1)
        btn_grid.columnconfigure(1, weight=1)
        btn_grid.columnconfigure(2, weight=1)
        btn_grid.columnconfigure(3, weight=1)

        # Preview area with Google Sheets-like UI
        preview_frame = ttk.Labelframe(self.main_frame, text="Step 5: Preview Results", padding=5)
        preview_frame.pack(fill="both", expand=True, padx=20, pady=(10, 15))
        
        # Toolbar for preview
        preview_toolbar = ttk.Frame(preview_frame, style="TFrame")
        preview_toolbar.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(preview_toolbar, text="Refresh", command=self.refresh_preview, 
                 style="TButton", width=10).pack(side="left", padx=2)
        ttk.Button(preview_toolbar, text="Full Screen", command=self.show_full_preview, 
                 style="TButton", width=12).pack(side="left", padx=2)
        ttk.Button(preview_toolbar, text="Download", command=self.download_result, 
                 style="TButton", width=10).pack(side="right", padx=2)
        
        # Preview grid
        self.preview_container = ttk.Frame(preview_frame, style="Preview.TFrame")
        self.preview_container.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        
        # Create treeview with Google Sheets-like appearance
        self.create_preview_grid()

        # Progress bar (hidden by default)
        self.progress = ttk.Progressbar(self.main_frame, orient="horizontal", 
                                      mode="indeterminate", length=300)
        self.progress.pack(fill="x", padx=20, pady=5)
        self.progress.pack_forget()

    def create_preview_grid(self):
        # Clear existing widgets
        for widget in self.preview_container.winfo_children():
            widget.destroy()
        
        # Create scrollable canvas
        canvas = tk.Canvas(self.preview_container, bg="white", highlightthickness=0)
        hscroll = ttk.Scrollbar(self.preview_container, orient="horizontal", command=canvas.xview)
        vscroll = ttk.Scrollbar(self.preview_container, orient="vertical", command=canvas.yview)
        canvas.configure(xscrollcommand=hscroll.set, yscrollcommand=vscroll.set)
        
        # Grid layout for scrollbars
        canvas.grid(row=0, column=0, sticky="nsew")
        vscroll.grid(row=0, column=1, sticky="ns")
        hscroll.grid(row=1, column=0, sticky="ew")
        
        self.preview_container.grid_rowconfigure(0, weight=1)
        self.preview_container.grid_columnconfigure(0, weight=1)
        
        # Frame inside canvas for content
        self.grid_frame = ttk.Frame(canvas, style="Preview.TFrame")
        canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")
        
        # Bind configuration
        self.grid_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        # Mouse wheel scrolling
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", 
            lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units")))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        
        # Initial empty grid
        self.update_preview_grid(pd.DataFrame())

    def update_preview_grid(self, df):
        # Clear existing grid
        for widget in self.grid_frame.winfo_children():
            widget.destroy()
        
        if df.empty:
            # Show empty state
            empty_label = ttk.Label(self.grid_frame, text="No data to preview. Load or process data first.", 
                                  style="Preview.TLabel", font=("Arial", 12))
            empty_label.pack(pady=50)
            return
        
        # Limit preview to first 100 rows and 20 columns for performance
        preview_df = df.head(100).iloc[:, :20]
        rows, cols = preview_df.shape
        
        # Create header row
        for j, col in enumerate(preview_df.columns):
            header = ttk.Label(self.grid_frame, text=str(col), style="Preview.TLabel", 
                             relief="solid", borderwidth=1, padding=(8,4), 
                             font=("Arial", 10, "bold"), background="#f0f0f0")
            header.grid(row=0, column=j, sticky="nsew")
        
        # Create data cells
        for i, row in preview_df.iterrows():
            for j, val in enumerate(row):
                cell = ttk.Label(self.grid_frame, text=str(val), style="Preview.TLabel", 
                               relief="solid", borderwidth=1, padding=(8,4))
                cell.grid(row=i+1, column=j, sticky="nsew")
        
        # Configure grid weights
        for i in range(rows+1):
            self.grid_frame.grid_rowconfigure(i, weight=1)
        for j in range(cols):
            self.grid_frame.grid_columnconfigure(j, weight=1)

    def refresh_preview(self):
        if self.current_preview_df is not None:
            self.update_preview_grid(self.current_preview_df)
            self.set_status("Preview refreshed")

    def show_full_preview(self):
        if self.result_df is None:
            messagebox.showwarning("Warning", "No data to preview")
            return
        
        top = tk.Toplevel(self.root)
        top.title("Full Preview - ABG Excel Tool")
        top.geometry("1200x800")
        top.configure(bg="#232946")
        
        # Create a frame for the preview
        preview_frame = ttk.Frame(top, style="TFrame")
        preview_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create a treeview widget for full preview
        tree = ttk.Treeview(preview_frame, show="headings", selectmode="extended")
        tree.pack(fill="both", expand=True, side="left")
        
        # Add scrollbars
        vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=tree.yview)
        vsb.pack(fill="y", side="right")
        hsb = ttk.Scrollbar(top, orient="horizontal", command=tree.xview)
        hsb.pack(fill="x", side="bottom")
        
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Set up columns
        tree["columns"] = list(self.result_df.columns)
        for col in self.result_df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="w")
        
        # Add data (limit to 1000 rows for performance)
        for i, row in self.result_df.head(1000).iterrows():
            tree.insert("", "end", values=list(row))
        
        # Add a download button
        btn_frame = ttk.Frame(top, style="TFrame")
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        ttk.Button(btn_frame, text="Download Full Data", command=lambda: self.download_result(self.result_df), 
                 style="TButton").pack(side="right", padx=5)

    def set_status(self, msg):
        self.status_var.set(msg)
        self.root.update_idletasks()

    def run_with_progress(self, func, *args, **kwargs):
        def wrapper():
            self.progress.pack()
            self.progress.start()
            try:
                func(*args, **kwargs)
            finally:
                self.progress.stop()
                self.progress.pack_forget()
        threading.Thread(target=wrapper).start()

    def load_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
        if files:
            self.files = list(files)  # Overwrite existing files
            self.file_listbox.delete(0, tk.END)
            for f in files:
                self.file_listbox.insert(tk.END, f)
            self.set_status(f"{len(files)} file(s) loaded")
            self.load_sheet_names()

    def clear_files(self):
        self.files = []
        self.dfs = []
        self.sheet_names = []
        self.file_listbox.delete(0, tk.END)
        self.sheet_combo['values'] = []
        self.col_listbox.delete(0, tk.END)
        self.result_df = None
        self.current_preview_df = None
        self.update_preview_grid(pd.DataFrame())
        self.set_status("Cleared all files and data")

    def load_sheet_names(self):
        if not self.files:
            return
        
        try:
            xl = pd.ExcelFile(self.files[0])
            self.sheet_names = xl.sheet_names
            self.sheet_combo['values'] = self.sheet_names
            if self.sheet_names:
                self.sheet_combo.current(0)
            self.set_status(f"Loaded {len(self.sheet_names)} sheets from first file")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read sheets: {str(e)}")
            self.set_status("Error loading sheets")

    def load_sheet(self):
        if not self.files or not self.sheet_combo.get():
            messagebox.showwarning("Warning", "Please select files and a sheet first")
            return
        
        self.set_status("Loading sheet data...")
        self.run_with_progress(self._load_sheet)

    def _load_sheet(self):
        self.dfs = []
        for f in self.files:
            try:
                df = pd.read_excel(f, sheet_name=self.sheet_combo.get())
                self.dfs.append(df)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load sheet from {f}:\n{str(e)}")
                self.set_status("Error loading sheet")
                return
        
        # Update column listbox
        self.col_listbox.delete(0, tk.END)
        for col in self.dfs[0].columns:
            self.col_listbox.insert(tk.END, col)
        
        # Show preview of first file's data
        self.current_preview_df = self.dfs[0]
        self.update_preview_grid(self.current_preview_df)
        self.set_status(f"Loaded sheet '{self.sheet_combo.get()}' from {len(self.files)} files")

    def concat_columns(self):
        indices = self.col_listbox.curselection()
        if not indices:
            messagebox.showwarning("Warning", "Please select columns to concatenate")
            return
        
        col_names = [self.col_listbox.get(i) for i in indices]
        for i, df in enumerate(self.dfs):
            self.dfs[i]['Concatenated'] = df[col_names].astype(str).agg(' '.join, axis=1)
        
        self.result_df = pd.concat(self.dfs, ignore_index=True)
        self.current_preview_df = self.result_df
        self.update_preview_grid(self.result_df)
        
        messagebox.showinfo("Success", f"Concatenated {len(col_names)} columns and merged {len(self.dfs)} sheets")
        self.set_status(f"Concatenated columns: {', '.join(col_names)}")

    def find_unique(self):
        if len(self.dfs) < 2:
            messagebox.showwarning("Warning", "You need at least 2 files to find unique rows")
            return
        
        # Create dialog to select comparison columns
        select_win = tk.Toplevel(self.root)
        select_win.title("Select Columns for Comparison")
        select_win.geometry("400x300")
        select_win.configure(bg="#232946")
        
        ttk.Label(select_win, text="Select column from first file:", style="TLabel").pack(pady=(20,5))
        col1_combo = ttk.Combobox(select_win, values=list(self.dfs[0].columns), state="readonly")
        col1_combo.pack(pady=5, padx=20, fill="x")
        col1_combo.current(0)
        
        ttk.Label(select_win, text="Select column from second file:", style="TLabel").pack(pady=(20,5))
        col2_combo = ttk.Combobox(select_win, values=list(self.dfs[1].columns), state="readonly")
        col2_combo.pack(pady=5, padx=20, fill="x")
        col2_combo.current(0)
        
        def perform_comparison():
            col1 = col1_combo.get()
            col2 = col2_combo.get()
            
            if not col1 or not col2:
                messagebox.showwarning("Warning", "Please select columns from both files")
                return
            
            try:
                # Get unique values from each file
                set1 = set(self.dfs[0][col1].dropna().astype(str))
                set2 = set(self.dfs[1][col2].dropna().astype(str))
                
                # Find unique to each set
                unique_to_1 = self.dfs[0][~self.dfs[0][col1].astype(str).isin(set2)]
                unique_to_2 = self.dfs[1][~self.dfs[1][col2].astype(str).isin(set1)]
                
                # Combine results with source markers
                unique_to_1['_Source'] = f"Only in {self.files[0]}"
                unique_to_2['_Source'] = f"Only in {self.files[1]}"
                
                self.result_df = pd.concat([unique_to_1, unique_to_2], ignore_index=True)
                self.current_preview_df = self.result_df
                self.update_preview_grid(self.result_df)
                
                select_win.destroy()
                messagebox.showinfo("Success", f"Found {len(unique_to_1)} unique rows in first file and {len(unique_to_2)} in second file")
                self.set_status("Unique rows comparison completed")
            except Exception as e:
                messagebox.showerror("Error", f"Comparison failed: {str(e)}")
                self.set_status("Error in comparison")
        
        ttk.Button(select_win, text="Find Unique Rows", command=perform_comparison, 
                 style="TButton").pack(pady=20)

    def vlookup(self):
        indices = self.col_listbox.curselection()
        if len(self.dfs) < 2 or not indices:
            messagebox.showwarning("Warning", "You need at least 2 files and select a key column for VLOOKUP")
            return
        
        key_col = self.col_listbox.get(indices[0])
        
        try:
            # Perform merge (VLOOKUP)
            result = pd.merge(self.dfs[0], self.dfs[1], on=key_col, how='left', 
                            suffixes=('_file1', '_file2'))
            
            self.result_df = result
            self.current_preview_df = self.result_df
            self.update_preview_grid(self.result_df)
            
            messagebox.showinfo("Success", f"VLOOKUP completed on column '{key_col}'")
            self.set_status(f"VLOOKUP completed on {key_col}")
        except Exception as e:
            messagebox.showerror("Error", f"VLOOKUP failed: {str(e)}")
            self.set_status("VLOOKUP failed")

    def merge_sheets(self):
        if len(self.dfs) < 2:
            messagebox.showwarning("Warning", "You need at least 2 files to merge")
            return
        
        try:
            # Simple concatenation - could be enhanced with more options
            self.result_df = pd.concat(self.dfs, ignore_index=True)
            self.current_preview_df = self.result_df
            self.update_preview_grid(self.result_df)
            
            messagebox.showinfo("Success", f"Merged {len(self.dfs)} sheets with {len(self.result_df)} total rows")
            self.set_status(f"Merged {len(self.dfs)} sheets")
        except Exception as e:
            messagebox.showerror("Error", f"Merge failed: {str(e)}")
            self.set_status("Merge failed")

    def export_result(self):
        if self.result_df is None:
            messagebox.showwarning("Warning", "No result data to export")
            return
        
        file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003", "*.xls"), ("CSV", "*.csv")],
            title="Save Result As"
        )
        
        if not file:
            return
        
        try:
            if file.endswith('.csv'):
                self.result_df.to_csv(file, index=False)
            else:
                self.result_df.to_excel(file, index=False)
            
            messagebox.showinfo("Success", f"Data successfully exported to:\n{file}")
            self.set_status(f"Exported to {file}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {str(e)}")
            self.set_status("Export failed")

    def download_result(self, df=None):
        if df is None:
            if self.result_df is None:
                messagebox.showwarning("Warning", "No result data to download")
                return
            df = self.result_df
        
        file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003", "*.xls")],
            title="Download Excel File"
        )
        
        if not file:
            return
        
        try:
            df.to_excel(file, index=False)
            messagebox.showinfo("Download Complete", f"Excel file successfully downloaded to:\n{file}")
            self.set_status(f"Downloaded to {file}")
        except Exception as e:
            messagebox.showerror("Download Failed", f"Error saving file: {str(e)}")
            self.set_status("Download failed")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToolApp(root)
    root.mainloop()