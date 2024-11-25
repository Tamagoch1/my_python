import pandas as pd
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
merged_frame.to_excel(output_file, index = False, header = None)

print("Процесс завершен")