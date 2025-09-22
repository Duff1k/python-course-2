from datetime import datetime
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# Имя входного файла, содержащего графики обучения
INPUT_FILE = "trainings.txt"
# Имя листа в сгенерированных файлах Excel
SHEET_NAME = "Schedule"
# Заголовки столбцов для файлов Excel.
HEADERS = ["Тренер", "Вид спорта", "Дата и время"]
# Список названий залов (Зал 1 – Зал 4)
HALLS = [f"Зал {i}" for i in range(1, 5)]

# Функция берёт одну строку из входного файла, разбивает её на компоненты (дата и время, вид спорта, тренер и зал) и возвращает словарь с проанализированной информацией.
def parse_line(line):
    columns = [p.strip() for p in line.split("|", 3)]
    if len(columns) != 4:
        return None
    dt_str, sport, coach, hall = columns
    coach = coach.replace("Тренер: ", "").strip()
    hall = hall.strip()

    try:
        dt = datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
    except ValueError:
        return None

    return {
        "hall": hall,
        "coach": coach,
        "sport": sport,
        "dt": dt,
        "dt_str": dt_str
    }


# Функция берет список строк из входного файла, вызывает parse_line() для каждой строки и организует проанализированные данные в defaultdict, где ключами являются названия залов, а значениями — списки проанализированных данных для этого зала.
def parse_from_list(lines):

    by_hall = defaultdict(list)

    for i, raw in enumerate(lines, start=1):
        line = raw.strip()
        if not line:
            continue

        rec = parse_line(line)
        if not rec:
            continue
        if rec["hall"] in HALLS:
            by_hall[rec["hall"]].append(rec)
    return by_hall


# Функция принимает название зала и список проанализированных данных для этого зала и создаёт файл Excel с информацией о расписании. Она задаёт заголовки столбцов, форматирует ячейки, сортирует данные по дате и времени и записывает их в файл Excel.
def write_excel_for_hall(hall_name: str, rows: list):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(HEADERS)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    rows_sorted = sorted(rows, key=lambda r: r["dt"])

    for r in rows_sorted:
        ws.append([r["coach"], r["sport"], r["dt_str"]])

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 30

    filename = f"{hall_name}.xlsx"
    wb.save(filename)
    print(f"Создан файл: {filename}")


# Эта функция вызывает "write_excel_for_hall()" для каждого зала, передавая соответствующие проанализированные данные.
def write_all_excel(by_hall: dict):
    for hall in HALLS:
        write_excel_for_hall(hall, by_hall.get(hall, []))

# Точкой входа скрипта. Функция открывает входной файл, считывает строки, вызывает parse_from_list() для организации данных по залу, а затем вызывает write_all_excel() для генерации файлов Excel.
def main():
    with open(INPUT_FILE, "r", encoding="utf-8") as file:
        lines = file.readlines()
    by_hall = parse_from_list(lines)
    write_all_excel(by_hall)


if __name__ == "__main__":
    main()