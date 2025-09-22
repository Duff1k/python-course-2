import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

keyword = ['Дата Время', 'Вид спорта', 'Тренер', 'Зал']
TRAININGS = []

for i in range(1, 5):

    with open("trainings.txt", "r", encoding="utf-8") as file:
        for line in file:
            clean_line = line.strip()
            if clean_line.endswith("%d" % i):
               clean_line = clean_line.split(sep=' | ')
               clean_line = dict(zip(keyword, clean_line[0:3]))
               TRAININGS.append(clean_line)

    for line in TRAININGS:
        # Убираем слово Тренер
        tr = line['Тренер']
        res_tr = tr.replace('Тренер: ', '')
        line['Тренер'] = res_tr

    sorted_TRAININGS = sorted(TRAININGS, key=lambda x: x['Дата Время'], reverse=False)
    TRAININGS.clear()

    wb = Workbook()
    ws = wb.active

    ws.title = "Зал %d" % i

    headers = ["Тренер", "Вид спорта", "Дата и время"]
    ws.append(headers)

    #Форматируем первую строку жирным
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 20

    for tr in sorted_TRAININGS:
        ws.append([tr["Тренер"], tr["Вид спорта"], tr["Дата_Время"]])

    sorted_TRAININGS.clear()

    filename = "Зал_%d.xlsx" % i
    wb.save(filename)

    print(f"Файл {filename} успешно создан!")




