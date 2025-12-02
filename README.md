generate a docx document from a docx template and insert a table into the doc that contains data from an xlsx or csv spreadsheet / file  

libre office supports docx and xlsx

virtual environment if desired  
```python -m venv /path/to/new/virtual/environment```

install support packages  
```pip install docxtpl```  
```pip install pandas openpyxl python-docx```  

adjust your doc template's static components  
headers, footers, page numbers etc.  

have a ```*.csv``` or ```*.xlsx``` of the data  
Replace ```Column1, Column2, Column3``` with the exact column names from your spreadsheet.
