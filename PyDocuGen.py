from docxtpl import DocxTemplate
from datetime import date

# Load the template
doc = DocxTemplate("template.docx")

# Define context for body placeholders only
context = {
    "name": "John",
    "project": "AI Workflow Optimization",
    "date": date.today().strftime("%B %d, %Y")
}

# Render the document
doc.render(context)

# Save the generated file
doc.save("generated_report.docx")