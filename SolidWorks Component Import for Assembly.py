import os
import win32com.client
import pythoncom
import pandas as pd

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
EXCEL_FILE = r"C:\Sidhu\OneDrive\Documents\SolidWorks\Book1.xlsx"
VARIANT_NONE = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)

# SolidWorks Document Type & Option Constants
swDocPART = 1
swDocASSEMBLY = 2
swOpenDocOptions_Silent = 1

# Layout Spacing (in meters: 0.5m = 500mm spacing)
INSTANCE_SPACING_X = 0.50  # Spacing between instances of the SAME part
COMPONENT_SPACING_Z = 0.50 # Spacing between DIFFERENT component types

# ---------------------------------------------------------------------------
# COM Connection Helpers
# ---------------------------------------------------------------------------
def connect_sw():
    try:
        app = win32com.client.GetActiveObject("SldWorks.Application")
        print("Connected to active SolidWorks session.")
    except Exception:
        print("Launching new SolidWorks instance...")
        app = win32com.client.Dispatch("SldWorks.Application")
    app.Visible = True
    app.UserControl = True
    return app

def new_assembly(app, title="New_Assembly"):
    """Creates a new assembly document using default or fallback templates."""
    try:
        template = app.GetUserPreferenceStringValue(1)  # swDefaultTemplateAssembly = 1
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
            ""
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
    """Finds the component model (.sldprt or .sldasm) matching the component name."""
    clean_name = str(comp_name).strip()

    # Check if exact file name is already passed with an extension
    direct_path = os.path.join(base_dir, clean_name)
    if os.path.isfile(direct_path):
        return direct_path

    # Try common SolidWorks extensions
    for ext in [".sldprt", ".SLDPRT", ".sldasm", ".SLDASM"]:
        candidate = os.path.join(base_dir, f"{clean_name}{ext}")
        if os.path.isfile(candidate):
            return candidate

    return None

def open_model_in_background(app, file_path):
    """Silently loads the part/subassembly into memory with correct ByRef arguments."""
    ext = os.path.splitext(file_path)[1].lower()
    doc_type = swDocASSEMBLY if "asm" in ext else swDocPART

    errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

    opened_doc = app.OpenDoc6(
        file_path,
        doc_type,
        swOpenDocOptions_Silent,
        "",
        errors,
        warnings
    )
    return opened_doc

def activate_document(app, doc_title):
    """Activates a document window using proper COM ByRef variant for Errors."""
    errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    app.ActivateDoc2(doc_title, False, errors)

def safe_rebuild(model):
    """Safely triggers a rebuild without crashing across dynamic COM bindings."""
    try:
        # ForceRebuild3 is a standard callable method across all SW COM versions
        model.ForceRebuild3(False)
    except Exception:
        try:
            rebuild_fn = getattr(model, "EditRebuild3", None)
            if callable(rebuild_fn):
                rebuild_fn()
        except Exception:
            pass

# ---------------------------------------------------------------------------
# Assembly Builder
# ---------------------------------------------------------------------------
def build_assembly_from_excel(excel_path):
    if not os.path.isfile(excel_path):
        raise FileNotFoundError(f"Excel file not found at: {excel_path}")

    # Read the BOM sheet
    df = pd.read_excel(excel_path)
    df.columns = [str(col).strip() for col in df.columns]

    app = connect_sw()

    # Group by 'Material' to generate an assembly per Material entry
    grouped = df.groupby("Material")

    for material_id, group in grouped:
        assembly_title = f"Assembly_{material_id}"
        print(f"\nProcessing Assembly for Material: {material_id}")
        assy_model = new_assembly(app, title=assembly_title)
        
        # Get title as property or method safely
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
                print(f"  [ERROR] File for component '{comp_name}' not found in: {folder}")
                continue

            # Load document into memory
            open_model_in_background(app, model_path)

            # Bring focus back to the assembly document
            activate_document(app, actual_title)

            # Place each component type in its own row along Z
            pos_z = comp_row_index * COMPONENT_SPACING_Z

            for instance in range(qty):
                # Place repeated instances spaced out along X
                pos_x = instance * INSTANCE_SPACING_X
                pos_y = 0.0

                # AddComponent5: (Path, ConfigOpt, ConfigName, IsVirtual, VirtualPath, X, Y, Z)
                comp = assy_model.AddComponent5(
                    model_path,
                    0,      # swAddComponentConfigOptions_CurrentSelected = 0
                    "",     # default configuration
                    False,  # IsVirtual = False
                    "",     # virtual path
                    pos_x,
                    pos_y,
                    pos_z
                )

                if comp:
                    if is_first_comp:
                        assy_model.Extension.SelectByID2(
                            comp.Name2, "COMPONENT", 0, 0, 0, False, 0, VARIANT_NONE, 0
                        )
                        assy_model.FixComponent()
                        is_first_comp = False
                        print(f"  Inserted (Fixed): {comp_name} [Instance {instance + 1}/{qty}] at X={pos_x:.2f}m, Z={pos_z:.2f}m")
                    else:
                        print(f"  Inserted: {comp_name} [Instance {instance + 1}/{qty}] at X={pos_x:.2f}m, Z={pos_z:.2f}m")
                else:
                    print(f"  [WARNING] Failed to insert component: {model_path}")

            comp_row_index += 1

        # Safe rebuild and save
        safe_rebuild(assy_model)
        save_folder = group["Location"].iloc[0]
        output_asm_path = os.path.join(save_folder, f"{material_id}.SLDASM")
        assy_model.SaveAs3(output_asm_path, 0, 2)
        print(f"Assembly saved to: {output_asm_path}")

# ---------------------------------------------------------------------------
# Main Execution
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    build_assembly_from_excel(EXCEL_FILE)
