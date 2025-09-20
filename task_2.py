from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from datetime import datetime


def create_workbook(title):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Расписание'
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 30
    ws['A1'].font = Font(bold=True)
    ws['B1'].font = Font(bold=True)
    ws['C1'].font = Font(bold=True)
    ws['A1'] = 'Тренер'
    ws['B1'] = 'Вид спорта'
    ws['C1'] = 'Дата и время'
    wb.save(f'{title}.xlsx')


data_about_workout = []
halls = []


with open("trainings.txt", "r", encoding="utf-8") as data:
    lines = data.readlines()
    for line in lines:
        line = line.strip().split(' | ')
        time = datetime.strptime(line[0],'%Y-%m-%d %H:%M').strftime(f'%d.%m.%Y %H:%M')
        dict_with_data = {
            'Дата и время' : time,
            'Спорт' : line[1],
            'Тренер' : line[2][8:],
            'Зал' : line[3]
        }
        data_about_workout.append(dict_with_data)

        if line[3] in halls:
                continue
        else:
            halls.append(line[3])

for hall in halls:
    create_workbook(hall)

data_about_workout = sorted(data_about_workout, key=lambda x: x['Зал'])

for hall in halls:
    n = 1
    for inf in data_about_workout:
        if inf['Зал'] != hall:
            continue
        else:
            writer = load_workbook(f'{inf['Зал']}.xlsx')
            writer_sheet = writer.active
            writer_sheet[f'A{n+1}'] = inf['Тренер']
            writer_sheet[f'B{n+1}'] = inf['Спорт']
            writer_sheet[f'C{n+1}'] = inf['Дата и время']
            writer.save(f'{inf['Зал']}.xlsx')
            n += 1
            print(inf)


