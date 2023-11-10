import openpyxl.utils.exceptions
from openpyxl import Workbook, load_workbook
from datetime import datetime
from style import *


SHEET_NAME = "Journey"

def make_headlines(work_sheet):
    work_sheet["B2"] = "Date"
    work_sheet["C2"] = "Title"
    work_sheet["D2"] = "Journal"
    work_sheet.row_dimensions[2].height = 20
    work_sheet.column_dimensions['B'].width = 12
    work_sheet.column_dimensions['C'].width = 25
    work_sheet.column_dimensions['D'].width = 40


if __name__ == '__main__':
    print("Program currently only work when notes have format:\n"
          "'Day' Date Title\n"
          "(body)\n"
          "'Day' Date Title\n"
          "(body)\n")

    excel_name = input("Type name or path of the Excel you want to open or create (maybe '../Journal.xlsx'): ")
    while True:
        try:
            wb = load_workbook(filename=excel_name)
            try:
                ws = wb[SHEET_NAME]
            except KeyError:
                wb.create_sheet(SHEET_NAME)
                ws = wb[SHEET_NAME]
                make_headlines(ws)
            break
        except openpyxl.utils.exceptions.InvalidFileException:
            print("Adding .xlsx at the end")
            excel_name = excel_name + ".xlsx"
        except FileNotFoundError:
            print("Creating new .xlsx file")
            wb = Workbook()
            ws = wb.active
            ws.title = SHEET_NAME
            make_headlines(ws)
            ws = wb[SHEET_NAME]
            break

    # Opening text file
    new_rows_amount = 0
    journal_date, journal_title, journal_body = "", "", ""
    notes_file_path = input("Type path to your notes (like '../Journey.txt'):")

    while True:
        try:
            file = open(notes_file_path, "r", encoding="utf-8")
            for line in file:
                if line.lower().startswith("day"):
                    if journal_body != "":
                        journal_body = journal_body.strip("\n")
                        data = [None, journal_date, journal_title, journal_body]
                        ws.append(data)
                        new_rows_amount += 1
                        journal_body = ""
                    line = line.removeprefix("Day").strip(" ")
                    journal_date = datetime.strptime(line[0:10], '%d.%m.%Y').date()
                    journal_title = line[11:len(line)]
                else:
                    journal_body += line

            style(new_rows_amount, ws, wb)
            file.close()
            wb.save(excel_name)
            break
        except PermissionError:
            print("Can't modify Excel file while is open!")
            break
        except FileNotFoundError:
            print(f"No such file or directory: '{notes_file_path}'")
            notes_file_path = input("Type path to your notes (like '../Journey.txt'):")
