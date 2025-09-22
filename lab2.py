from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font

INPUT_FILE = "Trainings.txt"

halls = {"Зал 1": [], "Зал 2": [], "Зал 3": [], "Зал 4": []}

with open(INPUT_FILE, "r", encoding="utf-8") as f:
    for raw_line in f:
        line = raw_line.strip()

        if not line or "|" not in line:
            continue

        parts = [p.strip() for p in line.split("|")]
        if len(parts) != 4:
            continue

        dt_text = parts[0]
        sport = parts[1]
        trainer_part = parts[2]
        hall = parts[3]

        if "Тренер:" in trainer_part:
            trainer = trainer_part.split("Тренер:")[1].strip()
        else:
            trainer = trainer_part.strip()

        try:
            dt = datetime.strptime(dt_text, "%Y-%m-%d %H:%M")
        except ValueError:

            continue

        if hall in halls:
            halls[hall].append((dt, trainer, sport, dt_text))


for hall_name, items in halls.items():

    items.sort(key=lambda x: x[0])

    wb = Workbook()
    ws = wb.active
    ws.title = "Расписание"
    headers = ["Тренер", "Вид спорта", "Дата и время"]
    ws.append(headers)

    bold_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = bold_font

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 22

    for dt, trainer, sport, dt_text in items:

        ws.append([trainer, sport, dt_text])

    out_filename = f"{hall_name}.xlsx"
    wb.save(out_filename)
