import pandas as pd

with open('trainings.txt', 'r', encoding='utf-8') as file:
    lines = file.readlines()

data = []
for line in lines:
    date_time, sport, coach, hall = line.strip().split('|')
    data.append({'Дата и время': date_time, 'Спорт':sport, 'Тренер': coach, 'Зал': hall})

df = pd.DataFrame(data)

for hall in df['Зал'].unique():
    df_hall = df[df['Зал'] == hall].copy()
    df_hall = df_hall.drop('Зал', axis=1)
    df_hall = df_hall.sort_values('Дата и время')
    filename = f'{hall}.xlsx'
    with pd.ExcelWriter(filename) as writer:
        df_hall.to_excel(writer, index=False, sheet_name = 'Расписание')
        workbook = writer.book
        worksheet = writer.sheets['Расписание']
        for col_idx, value in enumerate (df_hall.columns.values):
            worksheet.write(0, col_idx, value, workbook.add_format({'bold': True}))
        worksheet.set_column(0, len(df_hall.columns)-1, 25)
