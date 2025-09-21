from openpyxl import Workbook
from openpyxl.styles import Font
from datetime import datetime

def create_hall_schedule(hall, trainings):
    
    # Сортировка по дате и времени
    trainings.sort(key=lambda x: datetime.strptime(x['время'], '%Y-%m-%d %H:%M'))
    
    # Создание Excel файлов
    wb = Workbook()
    ws = wb.active
    ws.title = 'Расписание'
    
    # Заголовки жирным
    ws['A1'] = 'Тренер'
    ws['B1'] = 'Вид спорта'
    ws['C1'] = 'Дата и время'
    ws['A1'].font = Font(bold=True)
    ws['B1'].font = Font(bold=True)
    ws['C1'].font = Font(bold=True)
    
    # Заполнение данных
    row = 2
    for training in trainings:
        ws[f'A{row}'] = training['тренер']
        ws[f'B{row}'] = training['спорт']
        ws[f'C{row}'] = training['время']
        row += 1
    
    # Настройка ширины колонок
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 18
    
    # Сохраниение
    wb.save(f'{hall}.xlsx')
    print(f'Создан файл: {hall}.xlsx')

with open("trainings.txt", "r", encoding="utf-8") as file:
    lines = file.readlines()

# Сборка данных по залам
halls_data = {}
for line in lines:
    parts = line.strip().split(' | ')
    if len(parts) == 4:
        hall = parts[3]
        if hall not in halls_data:
            halls_data[hall] = []
        halls_data[hall].append({
            'тренер': parts[2][8:],
            'спорт': parts[1],
            'время': parts[0]
        })

# Создание файлов для каждого зала
for hall, trainings in halls_data.items():
    create_hall_schedule(hall, trainings)

