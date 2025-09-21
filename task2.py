from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import datetime

wb = Workbook()
ws = wb.active
headers = ["Тренер", "Вид спорта", "Дата и время"]

def list_clear(title):
    ws.delete_rows(1,ws.max_row)
    ws.title = title
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20


def ws_save(gym, title):
    for val in gym:
        ws.append([val[headers[0]], val[headers[1]]])
        ws.cell(ws.max_row, 3).value = datetime.datetime.strptime(val[headers[2]], "%Y-%m-%d %H:%M")
        ws.cell(ws.max_row, 3).number_format = "yyyy-MM-dd HH:mm"
    filename = prevGym + ".xlsx"
    wb.save(filename)

values = []

with open("trainings.txt", "r", encoding="utf-8") as file:
    rows = file.readlines()
    for row in rows:
        cells = row.split(" | ")
        value = {"Зал": cells[3].split("\n")[0], headers[0]: cells[2].split(": ")[1], headers[1]: cells[1], headers[2]: cells[0]}
        values.append(value)

values.sort(key=lambda x: x["Зал"])

prevGym = values[0]["Зал"]
list_clear(prevGym)
gymValues = []
for value in values:
    if value["Зал"] != prevGym:
        gymValues.sort(key=lambda x: x["Дата и время"])
        ws_save(gymValues, prevGym)
        prevGym = value["Зал"]
        list_clear(prevGym)
        gymValues = []
    else:
        gv = {headers[0]: value[headers[0]], headers[1]: value[headers[1]], headers[2]: value[headers[2]]}
        gymValues.append(gv)

ws_save(gymValues, prevGym)
