import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pythoncom
import win32com.client

# ---------------------------------------------------------------------------
# SolidWorks COM Constants
# ---------------------------------------------------------------------------
VARIANT_NONE = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
swDocPART = 1
swDocASSEMBLY = 2
swOpenDocOptions_Silent = 1

INSTANCE_SPACING_X = 0.50  # 500 mm spacing between instances of the same part
COMPONENT_SPACING_Z = 0.50  # 500 mm spacing between different components

# ---------------------------------------------------------------------------
# SolidWorks Automation Core
# ---------------------------------------------------------------------------
def connect_sw():
    try:
        app = win32com.client.GetActiveObject("SldWorks.Application")
    except Exception:
        app = win32com.client.Dispatch("SldWorks.Application")
    app.Visible = True
    app.UserControl = True
    return app

def new_assembly(app, title="New_Assembly"):
    try:
        template = app.GetUserPreferenceStringValue(1)
    except Exception:
        template = ""

    assy = None
    if template and os.path.isfile(template):
        assy = app.NewDocument(template, 0, 0, 0)
    else:
        candidates = [
            r"C:\ProgramData\SolidWorks\SOLIDWORKS 2026\templates\Assembly.asmdot",
            r"C:\ProgramData\SolidWorks\SOLIDWORKS 2025\templates\Assembly.asmdot",
            r"C:\ProgramData\SolidWorks\SOLIDWORKS 2024\templates\Assembly.asmdot",
            r"C:\ProgramData\SolidWorks\SOLIDWORKS 2023\templates\Assembly.asmdot",
            r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\english\Tutorial\assembly.asmdot",
            "",
        ]
        for path in candidates:
            try:
                assy = app.NewDocument(path, 0, 0, 0)
                if assy:
                    break
            except Exception:
                continue

    if assy is None:
        raise RuntimeError("Failed to create a new SolidWorks Assembly document.")

    model = app.ActiveDoc
    try:
        model.SetTitle2(str(title))
    except Exception:
        pass
    return model

def resolve_file_path(base_dir, comp_name):
    clean_name = str(comp_name).strip()
    direct_path = os.path.join(base_dir, clean_name)
    if os.path.isfile(direct_path):
        return direct_path

    for ext in [".sldprt", ".SLDPRT", ".sldasm", ".SLDASM"]:
        candidate = os.path.join(base_dir, f"{clean_name}{ext}")
        if os.path.isfile(candidate):
            return candidate

    return None

def open_model_in_background(app, file_path):
    ext = os.path.splitext(file_path)[1].lower()
    doc_type = swDocASSEMBLY if "asm" in ext else swDocPART

    errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

    return app.OpenDoc6(file_path, doc_type, swOpenDocOptions_Silent, "", errors, warnings)

def activate_document(app, doc_title):
    errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    app.ActivateDoc2(doc_title, False, errors)

def safe_rebuild(model):
    try:
        model.ForceRebuild3(False)
    except Exception:
        try:
            rebuild_fn = getattr(model, "EditRebuild3", None)
            if callable(rebuild_fn):
                rebuild_fn()
        except Exception:
            pass

def process_assembly(excel_path, log_callback, progress_callback):
    # Required for win32com inside background threads
    pythoncom.CoInitialize()
    try:
        if not os.path.isfile(excel_path):
            raise FileNotFoundError(f"Excel file not found: {excel_path}")

        log_callback(f"Reading BOM: {os.path.basename(excel_path)}")
        df = pd.read_excel(excel_path)
        df.columns = [str(col).strip() for col in df.columns]

        required_cols = {"Material", "Component", "Quantity", "Location"}
        if not required_cols.issubset(set(df.columns)):
            raise ValueError(f"Missing required columns. Found: {list(df.columns)}")

        log_callback("Connecting to SolidWorks...")
        app = connect_sw()

        grouped = df.groupby("Material")
        total_rows = len(df)
        processed_rows = 0

        for material_id, group in grouped:
            assembly_title = f"Assembly_{material_id}"
            log_callback(f"\nBuilding Assembly for Material: {material_id}")
            assy_model = new_assembly(app, title=assembly_title)

            title_attr = getattr(assy_model, "GetTitle", None)
            actual_title = title_attr() if callable(title_attr) else str(title_attr)

            is_first_comp = True
            comp_row_index = 0

            for _, row in group.iterrows():
                comp_name = str(row["Component"]).strip()
                qty = int(row["Quantity"])
                folder = str(row["Location"]).strip()

                model_path = resolve_file_path(folder, comp_name)
                if not model_path:
                    log_callback(f"  [ERROR] File missing: '{comp_name}' in {folder}")
                    processed_rows += 1
                    progress_callback(processed_rows, total_rows)
                    continue

                open_model_in_background(app, model_path)
                activate_document(app, actual_title)

                pos_z = comp_row_index * COMPONENT_SPACING_Z

                for instance in range(qty):
                    pos_x = instance * INSTANCE_SPACING_X
                    pos_y = 0.0

                    comp = assy_model.AddComponent5(
                        model_path, 0, "", False, "", pos_x, pos_y, pos_z
                    )

                    if comp:
                        if is_first_comp:
                            assy_model.Extension.SelectByID2(
                                comp.Name2, "COMPONENT", 0, 0, 0, False, 0, VARIANT_NONE, 0
                            )
                            assy_model.FixComponent()
                            is_first_comp = False
                            log_callback(f"  Inserted (Fixed): {comp_name} [Instance {instance + 1}/{qty}]")
                        else:
                            log_callback(f"  Inserted: {comp_name} [Instance {instance + 1}/{qty}]")
                    else:
                        log_callback(f"  [WARNING] Insert failed: {model_path}")

                comp_row_index += 1
                processed_rows += 1
                progress_callback(processed_rows, total_rows)

            safe_rebuild(assy_model)
            save_folder = str(group["Location"].iloc[0]).strip()
            output_asm_path = os.path.join(save_folder, f"{material_id}.SLDASM")
            assy_model.SaveAs3(output_asm_path, 0, 2)
            log_callback(f"Saved: {output_asm_path}")

        log_callback("\nAll assemblies built successfully.")
    finally:
        pythoncom.CoUninitialize()

# ---------------------------------------------------------------------------
# Modern GUI Interface
# ---------------------------------------------------------------------------
class AssemblyImporterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SolidWorks BOM Assembly Automator")
        self.geometry("780x620")
        self.minsize(700, 520)
        self.configure(bg="#1E1E24")

        self._init_styles()
        self._build_layout()

    def _init_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        # Configure dark theme palette
        style.configure(".", background="#1E1E24", foreground="#ECEFF4")
        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            background="#007ACC",
            foreground="#FFFFFF",
            borderwidth=0,
            padding=8,
        )
        style.map("Primary.TButton", background=[("active", "#0098FF")])

        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 9),
            background="#3E4452",
            foreground="#FFFFFF",
            borderwidth=0,
            padding=6,
        )
        style.map("Secondary.TButton", background=[("active", "#4C566A")])

        style.configure(
            "Accent.Horizontal.TProgressbar",
            troughcolor="#282C34",
            background="#007ACC",
            thickness=6,
            borderwidth=0,
        )

    def _build_layout(self):
        # Header banner
        header = tk.Frame(self, bg="#181A1F", padx=24, pady=18)
        header.pack(fill=tk.X)

        title_lbl = tk.Label(
            header,
            text="SolidWorks Automated Assembly Builder",
            font=("Segoe UI", 15, "bold"),
            bg="#181A1F",
            fg="#61AFEF",
        )
        title_lbl.pack(anchor="w")

        sub_lbl = tk.Label(
            header,
            text="Batch import components from Excel BOM with automated spacing and mates",
            font=("Segoe UI", 9),
            bg="#181A1F",
            fg="#ABB2BF",
        )
        sub_lbl.pack(anchor="w", pady=(3, 0))

        # Main content card
        content = tk.Frame(self, bg="#1E1E24", padx=24, pady=18)
        content.pack(fill=tk.BOTH, expand=True)

        # File selection section
        file_frame = tk.Frame(content, bg="#282C34", padx=16, pady=16, highlightthickness=1, highlightbackground="#3E4452")
        file_frame.pack(fill=tk.X)

        file_title = tk.Label(
            file_frame, text="BOM EXCEL SHEET", font=("Segoe UI", 9, "bold"), bg="#282C34", fg="#98C379"
        )
        file_title.pack(anchor="w", pady=(0, 8))

        entry_row = tk.Frame(file_frame, bg="#282C34")
        entry_row.pack(fill=tk.X)

        self.file_path_var = tk.StringVar()
        self.entry_path = tk.Entry(
            entry_row,
            textvariable=self.file_path_var,
            font=("Segoe UI", 10),
            bg="#1E1E24",
            fg="#ECEFF4",
            insertbackground="#ECEFF4",
            relief=tk.FLAT,
            bd=6,
        )
        self.entry_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        browse_btn = ttk.Button(entry_row, text="Browse...", style="Secondary.TButton", command=self._browse_file)
        browse_btn.pack(side=tk.RIGHT)

        # Execution Controls
        btn_frame = tk.Frame(content, bg="#1E1E24", pady=12)
        btn_frame.pack(fill=tk.X)

        self.start_btn = ttk.Button(
            btn_frame,
            text="Run Assembly Build",
            style="Primary.TButton",
            command=self._start_import_thread,
        )
        self.start_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Progress bar
        self.progress = ttk.Progressbar(content, style="Accent.Horizontal.TProgressbar", mode="determinate")
        self.progress.pack(fill=tk.X, pady=(0, 12))

        # Terminal / Log Monitor
        log_frame = tk.Frame(content, bg="#181A1F", highlightthickness=1, highlightbackground="#3E4452")
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(
            log_frame,
            bg="#181A1F",
            fg="#98C379",
            insertbackground="#98C379",
            font=("Consolas", 9),
            relief=tk.FLAT,
            padx=10,
            pady=10,
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview, bg="#181A1F")
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

    def _browse_file(self):
        chosen = filedialog.askopenfilename(
            title="Select BOM Excel Spreadsheet",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
        )
        if chosen:
            self.file_path_var.set(chosen)

    def log(self, text):
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)

    def update_progress(self, current, total):
        pct = (current / total) * 100 if total > 0 else 0
        self.progress["value"] = pct

    def _start_import_thread(self):
        excel_path = self.file_path_var.get().strip()
        if not excel_path:
            messagebox.showwarning("File Required", "Please select an Excel sheet first.")
            return

        if not os.path.isfile(excel_path):
            messagebox.showerror("File Error", "The specified Excel file does not exist.")
            return

        self.start_btn.config(state=tk.DISABLED)
        self.progress["value"] = 0
        self.log_text.delete("1.0", tk.END)

        # Execute SolidWorks background automation on a worker thread
        def worker():
            try:
                process_assembly(excel_path, self.log, self.update_progress)
                messagebox.showinfo("Success", "SolidWorks assembly build completed successfully.")
            except Exception as e:
                self.log(f"\n[CRITICAL ERROR] {e}")
                messagebox.showerror("Execution Failed", str(e))
            finally:
                self.start_btn.config(state=tk.NORMAL)

        threading.Thread(target=worker, daemon=True).start()

# ---------------------------------------------------------------------------
# Entry Point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app = AssemblyImporterApp()
    app.mainloop()
