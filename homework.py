from openpyxl import Workbook
from openpyxl.styles import Font
from datetime import datetime

data =[]
with open('trainings.txt', 'r', encoding='utf-8') as file:
    for line in file:
        line = line.strip()
        parts = line.split(' | ')

        date_time_str, sport, trainer, hall = parts
        date_time_sort = datetime.strptime(date_time_str, '%Y-%m-%d %H:%M')
        data.append({
            'hall': hall,
            'trainer': trainer,
            'sport': sport,
            'date_time_str': date_time_str,
            'date_time_sort': date_time_sort
        })

halls = {}
for item in data:
    hall_name = item['hall']
    if hall_name not in halls:
        halls[hall_name] = []
    halls[hall_name].append(item)

# создаем Excel-файл
for hall_name, trainings in halls.items():
    # Сортируем тренировки по дате и времени
    trainings_sorted = sorted(trainings, key=lambda x: x['date_time_sort'])

    # Создаем новую книгу и лист, добавляем заголовки
    wb = Workbook()
    ws = wb.active
    ws.title = 'Расписание'
    headers = ['Тренер', 'Вид спорта', 'Дата и время']
    ws.append(headers)

    # Форматируем заголовки (жирный шрифт)
    for col in range(1, 4):
        ws.cell(row=1, column=col).font = Font(bold=True, size=14)

    # Добавляем данные
    for training in trainings_sorted:
        ws.append([training['trainer'], training['sport'], training['date_time_str']])

    # Ширина колонок
    ws.column_dimensions['A'].width = 20  # Тренер
    ws.column_dimensions['B'].width = 19  # Вид спорта
    ws.column_dimensions['C'].width = 18  # Дата и время

    # Сохраняем файл
    file_name = f'{hall_name}.xlsx'
    wb.save(file_name)
    print(f'Создан файл: {file_name}')

