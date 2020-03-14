import openpyxl, os, shutil
from openpyxl.styles.alignment import Alignment

cellsToValidate = ['A2', 'A8', 'G9', 'G10', 'A24', 'D24', 'H24', 'J24', 'H29', 'J29', 'J30']

excelFolder = "c:/Users/mark/PycharmProjects/excel/"

# create new folder for updated files
updateFolder = excelFolder + "backup/"
if not(os.path.exists(updateFolder)):
    os.mkdir(updateFolder)

def update_title():  # A2
    # is title in A2?
    return


def billing_date(cell_letter):  # G9
    billingDate = ws[cell_letter]
    billingDate.value = "Billing Date: 09/09/2021"
    billingDate.alignment = Alignment(horizontal="right")
    return

def due_date():  # G10
    return


def date_A24():  # A24
    return


def president_info():  # A8
    return


def bill_period():  # D24
    return


def payments():  # J24
    return


def totals():  # H29
    return


def payment():  # J30
    return


with os.scandir(excelFolder) as it:  # it = iterator returned by os.scandir(path)
    for dirEntry in it:
        if dirEntry.name.endswith(".xlsx"):
            # print("Entry name: " + entry.name)
            print("Entry path: " + dirEntry.path)

            # process spreadsheets, one for each dir entry
            wb = openpyxl.load_workbook(dirEntry.path)
            ws = wb.active
            billing_date("G9")
            saveAsName = updateFolder + dirEntry.name
            print("dirEntry.Name: " + dirEntry.name)
            print("saveAsName: " + saveAsName)
            wb.save(saveAsName)
            wb.close()
