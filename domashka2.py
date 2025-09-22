import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
import os
from datetime import datetime


def parse_trainings_file(filename):
    """Чтение и разбор файла с тренировками"""
    trainings = []

    with open(filename, 'r', encoding='utf-8') as file:
        for line in file:
            line = line.strip()
            if not line:
                continue

            parts = [part.strip() for part in line.split('|')]
            if len(parts) >= 4:
                date_time = parts[0].strip()
                sport = parts[1].strip()
                coach = parts[2].replace('Тренер:', '').strip()  # Убираем "Тренер:"
                hall = parts[3].strip()

                trainings.append({
                    'Тренер': coach,
                    'Вид спорта': sport,
                    'Дата и время': date_time,
                    'Зал': hall
                })

    return trainings


def create_excel_files(trainings):
    """Создание Excel-файлов для каждого зала"""
    # Группируем тренировки по залам
    halls_data = {}

    for training in trainings:
        hall = training['Зал']
        if hall not in halls_data:
            halls_data[hall] = []

        training_data = {
            'Тренер': training['Тренер'],
            'Вид спорта': training['Вид спорта'],
            'Дата и время': training['Дата и время']
        }
        halls_data[hall].append(training_data)

    for hall, hall_trainings in halls_data.items():
        df = pd.DataFrame(hall_trainings)

        df['datetime_obj'] = pd.to_datetime(df['Дата и время'], format='%Y-%m-%d %H:%M')

        df = df.sort_values('datetime_obj')

        df = df.drop('datetime_obj', axis=1)

        filename = f'{hall}.xlsx'

        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Расписание', index=False)

            workbook = writer.book
            worksheet = writer.sheets['Расписание']

            for cell in worksheet[1]:
                cell.font = Font(bold=True)

            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter

                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass

                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f'Создан файл: {filename} с {len(df)} занятиями')


def main():
    if not os.path.exists('trainings.txt'):
        print("Файл trainings.txt не найден!")
        return

    try:
        print("Чтение файла trainings.txt...")
        trainings = parse_trainings_file('trainings.txt')

        if not trainings:
            print("Не найдено данных о тренировках!")
            return

        print(f"Найдено {len(trainings)} тренировок")

        print("Создание Excel-файлов...")
        create_excel_files(trainings)

        print("Готово! Все файлы созданы успешно.")

    except Exception as e:
        print(f"Произошла ошибка: {e}")


if __name__ == "__main__":
    main()