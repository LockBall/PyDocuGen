import pandas as pd
from docxtpl import DocxTemplate
from datetime import date

# Load spreadsheet data
df = pd.read_excel("data_sample.xlsx")   # or 
# df = pd.read_csv("data.csv")

# Convert DataFrame to list of dicts for docxtpl
records = df.to_dict(orient="records")

# Load the Word template
doc = DocxTemplate("template_sample.docx")

# Define context for placeholders
context = {
    "name": "your_name",
    "project": "auto doc gen",
    "date": date.today().strftime("%d %B %Y"),
    "table": records
}

# Render the document
doc.render(context)

# Save the generated file
doc.save("generated_report_with_table.docx")