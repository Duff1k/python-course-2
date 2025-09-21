import openpyxl
from openpyxl.styles import Font
from datetime import datetime

# Читаем исходный файл
with open("trainings.txt", "r", encoding="utf-8") as f:
    lines = f.readlines()

# Разбор строк
data_by_hall = {"Зал 1": [], "Зал 2": [], "Зал 3": [], "Зал 4": []}

for line in lines:
    line = line.strip()
    if not line:
        continue

    # Пример строки:
    # 2025-09-20 18:00 | Бокс | Тренер: Иван Петров | Зал 1
    parts = [p.strip() for p in line.split("|")]
    datetime_str = parts[0]
    sport = parts[1]
    trainer = parts[2].replace("Тренер: ", "").strip()
    hall = parts[3]

    # Преобразуем дату в datetime для сортировки
    dt = datetime.strptime(datetime_str, "%Y-%m-%d %H:%M")

    # Сохраняем
    data_by_hall[hall].append((trainer, sport, dt))

# Функция записи в Excel
def save_schedule(hall, records):
    # Сортировка по дате/времени
    records.sort(key=lambda x: x[2])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Расписание"

    # Заголовки
    headers = ["Тренер", "Вид спорта", "Дата и время"]
    ws.append(headers)

    # Жирный шрифт для заголовков
    for col in range(1, len(headers) + 1):
        ws.cell(row=1, column=col).font = Font(bold=True)

    # Заполняем строки
    for trainer, sport, dt in records:
        ws.append([trainer, sport, dt.strftime("%Y-%m-%d %H:%M")])

    # Устанавливаем ширину колонок
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 20

    # Сохраняем файл
    filename = f"{hall}.xlsx"
    wb.save(filename)

# Создаём файлы для всех залов
for hall, records in data_by_hall.items():
    if records:
        save_schedule(hall, records)
