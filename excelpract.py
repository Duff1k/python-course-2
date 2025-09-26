from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

students = [
    {"Имя": "Иван", "Фамилия": "Иванов", "Пол":"М", "Возраст":"20"},
    {"Имя": "Мария", "Фамилия": "Маринова", "Пол":"Ж", "Возраст":"19"},
    {"Имя": "Алексей", "Фамилия": "Алексеев", "Пол":"М", "Возраст":"21"},
    {"Имя": "Елена", "Фамилия": "Еленова", "Пол":"Ж", "Возраст":"22"},
    {"Имя": "Елизавета", "Фамилия": "Кроликова", "Пол":"Ж", "Возраст":"23"},
]

wb = Workbook()

ws = wb.active

ws.title = "Студенты"

headers = ["Имя", "Фамилия", "Пол", "Возраст"]
ws.append(headers)


for cell in ws[1]:
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center")

ws.column_dimensions["A"].width=15
ws.column_dimensions["B"].width=20
ws.column_dimensions["C"].width=10
ws.column_dimensions["D"].width=10

for student in students:
    ws.append([student["Имя"], student["Фамилия"], student["Пол"], student["Возраст"]])

filename = "students.xlsx"
wb.save(filename)

print(f"Файл {filename}' успешно создан!")
