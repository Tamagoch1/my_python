import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from pathlib import Path

folder_path = Path('C:/Users/KorobanNA/Desktop/All') # Путь к папке с файлами

files_list = [file for file in folder_path.glob('**/*.xls')] # Получаем список файлов для объединения

dataframes = [] # Создание пустого списка для хранения DataFrame’ов

for file_path in files_list: # Объединение всех файлов .xlsx в один большой файл
    df = pd.read_excel(file_path, header = None)
    df_final = df[5:(len(df)-2)]
    dataframes.append(df_final)
    
merged_frame = pd.concat(dataframes, ignore_index = True) # Объединение всех DataFrame’ов из списка

output_file = 'final.xlsx' # Сохранение объединенного файла в новом файле output.xlsx

writer = pd.ExcelWriter(output_file, engine = 'xlsxwriter', datetime_format = 'dd.mm.yyyy', date_format = 'dd.mm.yyyy') # Преобразование даты в нужный формат
merged_frame = merged_frame.drop_duplicates() # Убираем дубликаты
merged_frame.to_excel(writer, sheet_name='AllData', index = False, header = None)

for column in df_final: # Автоподбор ширины столбца
    column_length = max(merged_frame[column].astype(str).map(len).max(), len(str(column)))
    col_idx = merged_frame.columns.get_loc(column)
    writer.sheets['AllData'].set_column(col_idx, col_idx, column_length)
    workbook = writer.book
    worksheet = writer.sheets['AllData']
    header_format = workbook.add_format({'bg_color': '#FFFFFF'}) # Изменяем цвет ячеек на белый

writer.close()

print("Процесс завершен")
