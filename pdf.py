"""An attempt to make a GUI app that solves a specific problem - extracting data
from a PDF (structured, standardized layout on a smart card of Serbian Registration 
Certificate) that was created using smart card reader and special software for
reading the smart cards issued by Ministry of internal affairs of Republic of Serbia.
Then saving that data as an Excel file (.xlsx) and importing it in another Excel 
template (for creating new documents) or importing it into Access DB (using VBA script).
App was exported into .exe file using pyinstaller.

Date created: 24. october 2020.
By: Dušan Miletić

App works in a following manner:
1 - when GUI opens, drag & drop PDF file into the corresponding field (drag & drop 
functionality was needed because it was easier and faster for everyone than using 
(file open dialog)) - found out that PySimpleGUI was based on Tkinter, PySide and
WxPython which didn't support drag & drop functionality hence PySimpleGUIQt was used.
2 - attempt to read the data from PDF file into python script as a PDF class (using
pdfplumber) and it was needed to check if PDF file was created properly using pdf 
printers (you could extract data) because sometimes, people created a PDF file that
was actually an image (i.e. JPG) inside PDF so data extraction was not possible.
3 -  if data could be extracted from PDF file, the file was read and one large string
was created
4 - string was then separated by rows (PDF had structured data, every single one
had the same structure, just different data)
5 - data needed was then selected from it's corresponding row and placed inside
a variable
6 - because the variables were in latin letters (and some documents or template were
needed to be in cyrillic letters) they had to be transliterated into cyrillic letters
using a custom created module
7 - after transliteration, data was exported into a new Excel file - saob_data(.xlsx). 

It was easier to create a new excel file that was then imported into a new template 
excel document or DB then it was to directly export it into excel template (because openpyxl
was deleting pictures from excel files at that time) and most of the templates had 
logos or other images embeded in them. Second reason, because then the python script would be
too coupled with the files and programs used (like 5 separate people, with little to none IT 
knowledge, worked in DB and/or creating new Excel documents, doing office work at the same 
time and that was a bad idea.)

Mainloop of GUI has weird if statment for running because the app was placed on 
the server-pc first and all users would connect to it (file server) and use it. 
It was easier for bug fixes and improvments to have it in one place then on every
PC. But later, the possibility to run it on every pc was added, per request.
"""


import pdfplumber
from openpyxl import Workbook
import PySimpleGUIQt as sg
import translit

############
# FUNCTIONS
############


def data_extract_saob(path_to_file_saob):
    # opening of pdf
    pdf = pdfplumber.open(path_to_file_saob)

    # reading into class: pdf
    first_page = pdf.pages[0]
    # extraction of data from class: pdf
    # variable with extracted data as string
    str_first_page = first_page.extract_text()

    # error checking, is extracting data possible
    if str_first_page is not None:
        # string split into rows and preparing for parsing
        str_lines = str_first_page.splitlines()
        # important indexes for str_lines:
        # index 1  - vehicle number plates (20:)
        reg_oznake = str_lines[1][20:]

        # index 9  - surname of owner (9:)
        prezime = str_lines[9][9:].lower().upper().capitalize()

        # index 10 - name of owner (14:)
        ime = str_lines[10][14:].lower().capitalize()

        # index 11 - place and address of owner (17:)
        # data: city, municipality, street, number
        # extracting city
        grad_opstina_adresa = str_lines[11][17:]
        grad_end = grad_opstina_adresa.find(",")
        grad = grad_opstina_adresa[0:grad_end].lower().capitalize()
        # extracting municipality
        grad_len = int(len(grad)) + 1
        opstina_full = str_lines[11][(17 + grad_len) :]
        opstina_end = opstina_full.find(",")
        # extracting address
        adresa_end = grad_opstina_adresa.find(",,")
        adresa = (
            grad_opstina_adresa[(grad_end + 1 + opstina_end + 1) : adresa_end]
            .replace(",", " ")
            .lower()
            .title()
        )

        # index 12 - unique ID number of owner 'JMBG' (15:)
        # jmbg = str_lines[12][15:]

        # index 18 - date of first vehicle registration and year of production
        reg_datum = str_lines[18][25:35] + (".")
        proiz_datum = str_lines[18][56:]

        # index 19 - vehicle brand and model
        marka_full = str_lines[19][7:]
        marka_end = marka_full.find(" Model:")
        marka = marka_full[0:marka_end].strip()
        model = marka_full[(marka_end + 7) :].strip()

        # index 21 - color of vehicle
        boja_full = str_lines[21][6:]
        boja_end = boja_full.find(" Broj osovina")
        boja = boja_full[0:boja_end].lower().capitalize()

        # index 22 - VIN and ccm
        vin_full = str_lines[22][13:]
        vin_end = vin_full.find(" Zapremina")
        vin = vin_full[0:vin_end]
        ccm_end = vin_full.find(":") + 2
        ccm = vin_full[ccm_end:]

        # index 23 - engine number
        broj_motora_full = str_lines[23][13:]
        broj_motora_end = broj_motora_full.find(" Masa")
        broj_motora = broj_motora_full[0:broj_motora_end]

        # index 24 - engine power in kW
        kw_full = str_lines[24][14:]
        kw_end = kw_full.find(" Nosivost")
        kw = kw_full[0:kw_end]

        # index 26 - vehicle category
        kategorija_full = str_lines[26][12:]
        kategorija_end = kategorija_full.find(" masa")
        kategorija = kategorija_full[0:kategorija_end].lower().capitalize()

        # index 27 - fuel
        gorivo = str_lines[27][17:].lower()

        # index 28 - number of seats
        broj_mesta = str_lines[28][23:24]

        ###########
        # TRANSLIT
        ###########

        prezime = translit.to_cyrillic(prezime)
        ime = translit.to_cyrillic(ime)
        grad = translit.to_cyrillic(grad)
        adresa = translit.to_cyrillic(adresa)
        boja = translit.to_cyrillic(boja)
        kategorija = translit.to_cyrillic(kategorija)
        gorivo = translit.to_cyrillic(gorivo)

        ###############
        # EXPORT TO TXT
        ###############

        # make a file and write data to saob-data file
        wb = Workbook()
        sheet = wb.active

        sheet["A1"] = "PODACI SAOB:"
        sheet["A2"] = reg_oznake
        sheet["A3"] = prezime + " " + ime
        sheet["A4"] = grad
        sheet["A5"] = adresa
        sheet["A6"] = reg_datum
        sheet["A7"] = proiz_datum
        sheet["A8"] = marka
        sheet["A9"] = model
        sheet["A10"] = boja
        sheet["A11"] = vin
        sheet["A12"] = ccm
        sheet["A13"] = broj_motora
        sheet["A14"] = kw

        if kategorija == "Путницко возило" or kategorija == "Путницко возил":
            kategorija = "Путничко"
        if kategorija == "Моторцикл":
            kategorija = "Мотоцикл"
        sheet["A15"] = kategorija

        if "бензин" in gorivo:
            gorivo_if = "Бензин"
        elif "дизел" in gorivo:
            gorivo_if = "Дизел"
        sheet["A16"] = "ORIG: " + gorivo
        sheet["A17"] = gorivo_if

        sheet["A18"] = broj_mesta

        save_loc_full = path_to_file_saob[::-1]
        save_loc_end = save_loc_full.find("/")
        save_loc_end_index = len(path_to_file_saob) - save_loc_end
        save_loc_saob = path_to_file_saob[0:save_loc_end_index]
        wb.save(save_loc_saob + "saob_data.xlsx")

        return 1
    else:
        return 0


###########
#   GUI
###########

sg.theme("DarkGrey1")
sg.theme_button_color(("black", "#DCDCDC"))

layout = [
    [sg.Text(" Izvlacenje podataka iz saobracajne u PDF-u.")],
    [sg.Text(" Samo drag & drop PDF u polje ispod.")],
    [sg.InputText()],
    [sg.Button("EXTRACT")],
]

window = sg.Window("DC Mini v0.1", layout)

while True:  # Event Loop
    event, values = window.read()
    path_to_file_saob = values[0]

    if "server-pc" in path_to_file_saob:
        path_to_file_saob = values[0][5:]
        # print(6)
    else:
        path_to_file_saob = values[0][8:]
        # print(8)
    if event in (None, "EXTRACT"):
        if data_extract_saob(path_to_file_saob) == 1:
            sg.popup("Uspesno izvuceni podaci!")
        else:
            sg.popup("Iz ovog PDF-a se ne mogu izvuci podaci!")
        break


window.close()
