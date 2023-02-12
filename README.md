# Parse-PDF-GUI


An attempt to make a GUI app that solves a specific problem - extracting data from a PDF (structured, standardized layout on a smart card of Serbian Registration Certificate) that was created using smart card reader and special software for
eading the smart cards issued by Ministry of internal affairs of Republic of Serbia. Then saving that data as an Excel file (.xlsx) and importing it in another Excel template (for creating new documents) or importing it into Access DB (using VBA script). App was exported into .exe file using pyinstaller.

#
> By: Dušan Miletić 
>
> Date created: 24. october 2020.
#

# Dependencies

Project uses the following dependencies (which can be installed via pip):

> **PySimpleGUIQt** - for GUI creation

> **pdfplumber** - for PDF parsing

> **openpyxl** - for exporting data in a excel file (.xlsx)

# Running the app

You can run the app in two ways:

> Using Python (if you have Python and dependencies installed)
> 
> 1. ``git clone https://github.com/MDule/parse-pdf-gui.git``
> 2. ``pip install -r requirements.txt``
> 3. ``python pdf.py OR py pdf.py [for Windows] OR python3 pdf.py [for Linux]`` 

OR

> Directly by running the .exe file [Windows]
> 
> https://raw.githubusercontent.com/MDule/parse-pdf-gui/main/exe/pdf.exe
>
> SHA-256: 4826135F1FD2AE24335EA6AE1EBAF9335537BB4872F5E421014B3EA747507BAA



# How does app work

### App works in a following manner:

 1. When GUI opens, drag & drop PDF file into the corresponding field (drag & drop functionality was needed because it was easier and faster for everyone than using (file open dialog)) - found out that PySimpleGUI was based on Tkinter, PySide and WxPython which didn't support drag & drop functionality hence PySimpleGUIQt was  used.

2. Attempt to read the data from PDF file into python script as a PDF class (using pdfplumber) and it was needed to check if PDF file was created properly using pdf printers (you could extract data) because sometimes, people created a PDF file that was actually an image (i.e. JPG) inside PDF so data extraction was not possible.

3. If data could be extracted from PDF file, the file was read and one large string was created

4. String was then separated by rows (PDF had structured data, every single one
had the same structure, just different data)

5. Data needed was then selected from it's corresponding row and placed inside
a variable

6. Because the variables were in latin letters (and some documents or template were needed to be in cyrillic letters) they had to be transliterated into cyrillic letters using a custom created module

7. After transliteration, data was exported into a new Excel file - saob_data(.xlsx). 

### **Note 1**
It was easier to create a new excel file that was then imported into a new template excel document or DB then it was to directly export it into excel template (because openpyxl was deleting pictures from excel files at that time) and most of the templates had logos or other images embeded in them. Second reason, because then the python script would be too coupled with the files and applications used (like 5 separate people, with little to none IT knowledge, worked in DB and/or creating new Excel documents, doing office work at the same time and that was a bad idea.)

### **Note 2**
Mainloop of GUI has weird if statment for running because the app was placed on 
the server-pc first and all users would connect to it (file server) and use it. 
It was easier for bug fixes and improvments to have it in one place then on every
PC. But later, the possibility to run it on every pc was added, per request.

### **Disclaimer**
Provided sample PDF for testing purposes doesn't represent a real vehicles registration ID. The data has been changed to protect sensitive info, swaped with dummy data. Any similarity is entirely coincidental.
