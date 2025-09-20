from datetime import datetime
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# final const
INPUT_FILE = "trainings.txt"
SHEET_NAME = "Расписание"
HEADERS = ["Тренер", "Вид спорта", "Дата и время"]
HALLS = [f"Зал {i}" for i in range(1, 5)]  # Зал 1..4


# Парсим строки
def parse_line(line):
    # Сплитим строку на составляющие
    columns = [p.strip() for p in line.split("|", 3)]
    # Проверка на адекватность
    if len(columns) != 4:
        return None
    # Присваивание
    dt_str, sport, coach, hall = columns
    coach = coach.replace("Тренер: ", "").strip()
    hall = hall.strip()
    # Парсим время
    try:
        dt = datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
    except ValueError:
        return None
    # Возвращаем объект
    return {
        "hall": hall,
        "coach": coach,
        "sport": sport,
        "dt": dt,
        "dt_str": dt_str
    }


# Парсим строки из листа
def parse_from_list(lines):
    # Создаем дефолтный
    by_hall = defaultdict(list)

    for i, raw in enumerate(lines, start=1):
        line = raw.strip()
        if not line:
            continue
        # Словарь
        rec = parse_line(line)
        if not rec:
            continue
        if rec["hall"] in HALLS:
            by_hall[rec["hall"]].append(rec)
    return by_hall


# Создание Excel
def write_excel_for_hall(hall_name: str, rows: list):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(HEADERS)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Сортировка по дате/времени
    rows_sorted = sorted(rows, key=lambda r: r["dt"])

    # Данные
    for r in rows_sorted:
        ws.append([r["coach"], r["sport"], r["dt_str"]])

    # Параметры колонок
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 28

    filename = f"{hall_name}.xlsx"
    wb.save(filename)
    print(f"Создан файл: {filename}")


# Создание всех Excel
def write_all_excel(by_hall: dict):
    for hall in HALLS:
        write_excel_for_hall(hall, by_hall.get(hall, []))


def main():
    with open(INPUT_FILE, "r", encoding="utf-8") as file:
        lines = file.readlines()
    by_hall = parse_from_list(lines)
    write_all_excel(by_hall)


if __name__ == "__main__":
    main()
