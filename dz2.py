from openpyxl import Workbook
from openpyxl.styles import Font
from datetime import datetime

trainings = []
with open('trainings.txt', 'r', encoding='utf-8') as file:
    for line in file:
        if line.strip():
            date, sport, trainer, hall = line.strip().split(' | ')
            trainer = trainer.replace('Тренер: ', '')
            trainings.append((hall.strip(), trainer, sport, date))

trainings.sort(key=lambda x: (x[0], datetime.strptime(x[3], '%Y-%m-%d %H:%M')))

for hall in ['Зал 1', 'Зал 2', 'Зал 3', 'Зал 4']:
    wb = Workbook()
    ws = wb.active
    ws.title = "Расписание"
    ws.append(['Тренер', 'Вид спорта', 'Дата и время'])
    for i in ws[1]:
        i.font = Font(bold=True)
    for i, trainer, sport, date in trainings:
        if i == hall:
            ws.append([trainer, sport, date])
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 16
    wb.save(f'{hall}.xlsx')