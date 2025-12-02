import pandas as pd
from docxtpl import DocxTemplate
from datetime import date
import tkinter as tk
from tkinter import filedialog, messagebox

# -------------------------------
# Configuration / Global Variables
# -------------------------------
spreadsheet_path = ""
template_path = ""
output_path = ""

# Starting window size parameters
WINDOW_WIDTH = 600
WINDOW_HEIGHT = 300

# -------------------------------
# Functions
# -------------------------------
def update_status(message):
    """Update the status bar text."""
    status_label.config(text=message)

def select_spreadsheet():
    global spreadsheet_path
    spreadsheet_path = filedialog.askopenfilename(
        title="Select Spreadsheet",
        filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
    )
    if spreadsheet_path:
        lbl_data.config(text=f"Data file: {spreadsheet_path}")
        update_status("Spreadsheet loaded")

def select_template():
    global template_path
    template_path = filedialog.askopenfilename(
        title="Select Word Template",
        filetypes=[("Word files", "*.docx")]
    )
    if template_path:
        lbl_template.config(text=f"Template file: {template_path}")
        update_status("Template loaded")

def generate_report():
    global spreadsheet_path, template_path, output_path
    if not spreadsheet_path or not template_path:
        messagebox.showwarning("Missing Files", "Please select both a spreadsheet and a template.")
        update_status("Missing files")
        return

    try:
        # Load spreadsheet data
        if spreadsheet_path.endswith(".csv"):
            df = pd.read_csv(spreadsheet_path)
        else:
            df = pd.read_excel(spreadsheet_path)

        records = df.to_dict(orient="records")

        # Load Word template
        doc = DocxTemplate(template_path)

        # Context for placeholders
        context = {
            "name": "John",
            "project": "AI Workflow Optimization",
            "date": date.today().strftime("%B %d, %Y"),
            "table": records
        }

        # Render and save
        doc.render(context)

        # Ask user where to save the output
        output_path = filedialog.asksaveasfilename(
            title="Save Report As",
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx")]
        )
        if output_path:
            doc.save(output_path)
            lbl_output.config(text=f"Output file: {output_path}")
            update_status("Report generated successfully")

    except Exception as e:
        messagebox.showerror("Error", str(e))
        update_status("Error during generation")

# -------------------------------
# GUI Setup
# -------------------------------
root = tk.Tk()
root.title("PyDocuGen")

# Apply starting size from parameters
root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")

# --- Spreadsheet group ---
frame_data = tk.Frame(root)
frame_data.pack(pady=10)

btn_data = tk.Button(frame_data, text="Select Spreadsheet", command=select_spreadsheet)
btn_data.pack(side=tk.TOP, pady=5)

lbl_data = tk.Label(frame_data, text="Data file: (none selected)")
lbl_data.pack(side=tk.TOP, pady=5)

# --- Template group ---
frame_template = tk.Frame(root)
frame_template.pack(pady=10)

btn_template = tk.Button(frame_template, text="Select Word Template", command=select_template)
btn_template.pack(side=tk.TOP, pady=5)

lbl_template = tk.Label(frame_template, text="Template file: (none selected)")
lbl_template.pack(side=tk.TOP, pady=5)

# --- Output group ---
frame_output = tk.Frame(root)
frame_output.pack(pady=10)

btn_generate = tk.Button(frame_output, text="Generate Report", command=generate_report)
btn_generate.pack(side=tk.TOP, pady=5)

lbl_output = tk.Label(frame_output, text="Output file: (none generated)")
lbl_output.pack(side=tk.TOP, pady=5)

# --- Status bar ---
status_label = tk.Label(root, text="Ready", bd=1, relief=tk.SUNKEN, anchor="w")
status_label.pack(side=tk.BOTTOM, fill=tk.X)

root.mainloop()