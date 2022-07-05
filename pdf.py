#
# pokusaj programcica koji prvo ucita podatke iz saobracajne u pdf klasu
# zatim iz pdf klase izvlaci podatke i ubacuje ih u jedan string
# zatim taj string podeli po redovima
# zatim iz svakog reda (koji sadrzi potrebne podatke) izlaci podatke
# i smesta ih u varijablu
# nakon ekstrakcije podataka vrsi transliteraciju
# i na kraju pravi excel fajl sa podacima kako bi mogao da se importuje
# u template jer trenutno openpyxl brise slike iz template-a
#

import pdfplumber
import translit
from openpyxl import Workbook
import PySimpleGUIQt as sg

############
# FUNCTIONS
############


def data_extract_saob(path_to_file_saob):

    # otvaranje pdf-a
    pdf = pdfplumber.open(path_to_file_saob)

    # ucitavanje pdf-a u class: pdf
    first_page = pdf.pages[0]
    # ekstrakcija celog teksta iz class: pdf
    # promenljiva koja ekstraktovani tekst drzi u string
    str_first_page = (first_page.extract_text())

    # error checking, da li moze da se extract data iz pdf-a
    if str_first_page is not None:

        # string delimo na redove kako bi mogla da se radi ekstrakcija podataka
        str_lines = str_first_page.splitlines()
        # vazni indexi str_lines:
        # index 1  - broj reg tablica (20:)
        reg_oznake = str_lines[1][20:]

        # index 9  - prezime vlasnika (9:)
        prezime = str_lines[9][9:].lower().upper().capitalize()

        # index 10 - ime vlasnika (14:)
        ime = str_lines[10][14:].lower().capitalize()

        # index 11 - mesto i adresa vlasnika (17:)
        # izvlacenje grada
        grad_opstina_adresa = str_lines[11][17:]
        grad_end = grad_opstina_adresa.find(",")
        grad = grad_opstina_adresa[0:grad_end].lower().capitalize()
        # izvlacenje opstine
        grad_len = int(len(grad)) + 1
        opstina_full = str_lines[11][(17 + grad_len):]
        opstina_end = opstina_full.find(",")
        # izvlacenje adrese
        adresa_end = grad_opstina_adresa.find(",,")
        adresa = grad_opstina_adresa[(grad_end + 1 + opstina_end + 1):adresa_end].replace(",", " ").lower().title()
        
        # index 12 - JMBG vlasnika (15:)
        #jmbg = str_lines[12][15:]

        # index 18 - datum prve reg i godina proizvodnje
        reg_datum = str_lines[18][25:35] + (".")
        proiz_datum = str_lines[18][56:]

        # index 19 - marka i model
        marka_full = str_lines[19][7:]
        marka_end = marka_full.find(" Model:")
        marka = marka_full[0:marka_end].strip()
        model = marka_full[(marka_end + 7):].strip()

        # index 21 - boja automobila
        boja_full = str_lines[21][6:]
        boja_end = boja_full.find(" Broj osovina")
        boja = boja_full[0:boja_end].lower().capitalize()

        # index 22 - VIN i ccm
        vin_full = str_lines[22][13:]
        vin_end = vin_full.find(" Zapremina")
        vin = vin_full[0:vin_end]
        ccm_end = vin_full.find(":") + 2
        ccm = vin_full[ccm_end:]

        # index 23 - broj motora
        broj_motora_full = str_lines[23][13:]
        broj_motora_end = broj_motora_full.find(" Masa")
        broj_motora = broj_motora_full[0:broj_motora_end]

        # index 24 - snaga motora
        kw_full = str_lines[24][14:]
        kw_end = kw_full.find(" Nosivost")
        kw = kw_full[0:kw_end]

        # index 26 - kategorija vozila
        kategorija_full = str_lines[26][12:]
        kategorija_end = kategorija_full.find(" masa")
        kategorija = kategorija_full[0:kategorija_end].lower().capitalize()
        
        # index 27 - gorivo
        gorivo = str_lines[27][17:].lower()

        #index 28 - broj mesta
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

        # napraviti i upisati podatke u saob-data file
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
        #print(6)
    else:
        path_to_file_saob = values[0][8:]
        #print(8)
    if event in (None, 'EXTRACT'):
        if data_extract_saob(path_to_file_saob) == 1:
            sg.popup("Uspesno izvuceni podaci!")
        else:
            sg.popup("Iz ovog PDF-a se ne mogu izvuci podaci!")
        break
    

window.close()

