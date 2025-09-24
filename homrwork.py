
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from datetime import datetime



trainings = []
with open('trainings.txt', 'r', encoding='utf-8') as f:
    for line in f:
        line = line.strip()
        if not line:
            continue
        parts = [part.strip() for part in line.split(' | ')]
        if len(parts) != 4:
            continue
        dt_str, sport, trainer_raw, hall_raw = parts
        trainer = trainer_raw.replace('Тренер: ', '').strip()
        hall = hall_raw.replace('Зал ', '').strip()

        dt = datetime.strptime(dt_str, '%Y-%m-%d %H:%M')
        trainings.append({
            'datetime': dt,
            'datetime_str': dt_str,
            'sport': sport,
            'trainer': trainer,
            'hall': hall
        })


from collections import defaultdict
by_hall = defaultdict(list)
for t in trainings:
    by_hall[t['hall']].append(t)


for hall_num in sorted(by_hall.keys(), key=int):  # сортируем как числа: 1,2,3,4
    entries = by_hall[hall_num]

    entries.sort(key=lambda x: x['datetime'])


    wb = Workbook()
    ws = wb.active
    ws.title = "Расписание"


    headers = ["Тренер", "Вид спорта", "Дата и время"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)


    for entry in entries:
        ws.append([entry['trainer'], entry['sport'], entry['datetime_str']])


    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column].width = adjusted_width

   
    filename = f"Зал {hall_num}.xlsx"
    wb.save(filename)
    print(f"Создан файл: {filename}")

print("Готово!")