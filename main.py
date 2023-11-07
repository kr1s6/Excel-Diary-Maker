import openpyxl.utils.exceptions
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, NamedStyle, Border, Side
from datetime import datetime

title_style = NamedStyle(name="title_style")
title_style.font = Font(name="Arial")
title_style.border = Border(left=Side(style="thick"),
                            right=Side(style="thick"),
                            bottom=Side(style="thick"))
title_style.alignment = Alignment(horizontal='center',
                                  vertical='center',
                                  wrap_text=True)

normal_style = NamedStyle(name="normal_style")
normal_style.font = Font(name="Arial")
normal_style.border = Border(left=Side(style="thin"),
                             right=Side(style="thin"),
                             bottom=Side(style="thin"))
normal_style.alignment = Alignment(horizontal='center',
                                   vertical='center',
                                   wrap_text=True)

if __name__ == '__main__':
    # excel_name = input("Type name of the Excel you want to open or create: ")
    excel_name = "test"
    while True:
        try:
            new_wb = load_workbook(filename=excel_name)
        except openpyxl.utils.exceptions.InvalidFileException:
            print("Adding .xlsx at the end")
            excel_name = excel_name + ".xlsx"
        except FileNotFoundError:
            print("Creating new .xlsx file")
            new_wb = Workbook()
            new_ws = new_wb.active
            new_ws.title = "Journey"
            new_ws["A1"] = "Date"
            new_ws["B1"] = "Title"
            new_ws["C1"] = "Journal"
            new_wb.add_named_style(title_style)
            new_wb.add_named_style(normal_style)
            new_ws.row_dimensions[1].height = 20
            new_ws.column_dimensions['A'].width = 12
            new_ws.column_dimensions['B'].width = 25
            new_ws.column_dimensions['C'].width = 40
            for row in new_ws["A1:C1"]:
                for cell in row:
                    cell.style = title_style
            new_wb.save(excel_name)
            print("Excel file saved")
            break
        else:
            break

    wb = load_workbook(filename=excel_name)
    ws = wb["Journey"]

    # Opening your text file
    journal_date, journal_title, journal_body = "", "", ""

    file = open("C:/Users/okr65/Desktop/Journey10.txt", "r", encoding="utf-8")
    for line in file:
        if line.lower().startswith("day"):
            if journal_body != "":
                journal_body = journal_body.strip("\n")
                ws.append([journal_date, journal_title, journal_body])
                journal_body = ""
            line = line.removeprefix("Day").strip(" ")
            journal_date = datetime.strptime(line[0:10], '%d.%m.%Y').date()
            journal_title = line[11:len(line)]
        else:
            journal_body += line

    file.close()
    wb.save(excel_name)
