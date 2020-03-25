import logging
import os
import sys
import openpyxl
from openpyxl.styles import Font
from openpyxl.styles.alignment import Alignment

cellsToValidate = ['A2', 'A8', 'G9', 'G10', 'A24', 'D24', 'H24', 'J24', 'H29', 'J29', 'J30']

excelFolder = "/Users/mark/PycharmProjects/excel/"

# excelFolder = "/Users/mark/Documents/Sterling Arbor Homeowners Documents/"

def get_logger():
    formatstr = "%(asctime)s %(levelname)-8s [%(funcName)s:%(lineno)d] %(message)s"
    logging.basicConfig(
        format=formatstr,
        level=logging.DEBUG,
        filename="HOA_Invoices.log")
    logger = logging.getLogger("App.log")
    return logger


def main():
    try:
        logger = get_logger()
        logger.info("Beginning main()....")

        # create new folder for updated files
        updateFolder = excelFolder + "updates/"
        if not (os.path.exists(updateFolder)):
            os.mkdir(updateFolder)
            logger.info(f"{updateFolder} created for updates.")

        def create_new_filename(fileName):
            parts = fileName.split()  # split on the space delimiter (default)
            parts[1] = "2020"  # 2nd element of list contains the year, update it to the current year
            newFileName = " ".join(parts)  # put it all back together again containg new year, delimited by spaces
            return newFileName

        def update_cell(cellToUpdate, newValue):
            cell = ws[cellToUpdate]  # A2, C5, A8, etc....
            cell.value = newValue  # string, but maybe date or numeric if necessary
            return cell

        with os.scandir(excelFolder) as it:  # it = iterator returned
            for dirEntry in it:
                if dirEntry.name.endswith(".xlsx"):
                    # print("Entry name: " + dirEntry.name)
                    # print("Entry path: " + dirEntry.path)

                    # process spreadsheets, one for each dir entry
                    wb = openpyxl.load_workbook(dirEntry.path)
                    with wb as wrkbk:
                        logger.info(f"Loaded {dirEntry.path}....")
                        # get reference to worksheet
                        ws = wb.active
                        # ws.protection.disable()
                        # wb.save()

                        # set picture in A5 - Remember to install Pillow; pip install Pillow
                        # img = Image("image004.png")
                        # ws.add_image(img, "A5")

                        # invoice date
                        cell = update_cell("A2", "2020 Invoice")
                        cell.alignment = Alignment(horizontal="center", vertical="center")

                        # officer contact info
                        cell = update_cell("A8", "Mark Greenberg - President\n205-505-9216")
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        cell.font = Font(name='Arial', size=12, b=True)

                        # title
                        cell = update_cell("C5", "Sterling Arbor\nNeighborhood Association, Inc.")
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        cell.font = Font(name='Arial', size=12, b=True)

                        # billing date
                        cell = update_cell("G9", "Billing Date: 01/01/2020")
                        cell.alignment = Alignment(horizontal="left", vertical="center")

                        # due date
                        cell = update_cell("G10", "Due Date....: 04/15/2020")
                        cell.alignment = Alignment(horizontal="left", vertical="center")

                        # activity date
                        cell = update_cell("A24", "1/01/2020")
                        cell.alignment = Alignment(horizontal="center", vertical="center")

                        # detail description
                        desc = "Homeowner Dues for Year 2020"
                        cell = update_cell("D24", desc)
                        cell.alignment = Alignment(horizontal="center", vertical="center")

                        # amount due
                        update_cell("H24", "$100")

                        # payment
                        update_cell("J24", "$0.00")

                        # total amount due
                        update_cell("H29", "$100.00")

                        # total payments
                        update_cell("J29", "$0.00")

                        # amount due
                        update_cell("J30", "$100.00")

                        # mail to
                        mailTo = "Mark Greenberg, 6983 Sterling Ln, Trussville, AL  35173"
                        cell = update_cell("C34", mailTo)
                        cell.alignment = Alignment(horizontal="left", vertical="center")

                        # save updated files to /updates folder, leaving original file as is.
                        newFileName = create_new_filename(dirEntry.name)

                        saveFileName = updateFolder + newFileName  # concat new file name with new folder name
                        # print("dirEntry.Name: " + dirEntry.name)
                        # print("saveAsName: " + saveFileName)

                        # set print area
                        ws.print_area = "A1:K49"

                        wb.save(saveFileName)
                        wb.close()

                    logger.info(f"Updated {saveFileName} closed!")

    except (RuntimeError, TypeError, NameError) as err:
        logger.critical("Unexpected Error Occurred", sys.exc.info()[0])
        logger.critical(f"Unexpected Error Occurred, error info {err}")

    except:
        logger.critical("Unexpected Error Occurred", sys.exc.info()[0])

    finally:
        wb.close()
        logger.info(f"Updated {saveFileName} closed in finally!")

    if __name__ == '__main__':
        logger = get_logger()
        main()
        logger.info("End main()....")
