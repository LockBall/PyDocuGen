## Purpose
generate a docx document from a docx template and insert a table into the doc that contains data from an xlsx or csv spreadsheet / file  

## Setup
software that supports docx and xlsx files  
libre office is free and opensource and does  
https://www.libreoffice.org/

have python installed on your computer  
https://www.python.org/downloads/  
in the installer select to install tkinter

#### Recomended  
If desired, use a virtual environment to isolate support packages and changes from primary python installation  
```python -m venv /path/to/new/virtual/environment```

install support packages (in the venv)  
```pip install docxtpl```  
```pip install pandas openpyxl python-docx```  

clone this git repo in the venv

adjust your doc template's static components  
headers, footers, page numbers etc.  

have a ```*.csv``` or ```*.xlsx``` of the data  

the column names in the *.docx template must match the exact column names in the spreadsheet
e.g. Replace ```Column1, Column2, Column3``` with ```Employee,	Department,	Hours Worked``` respectively  
this has already been done in the sample files provided

