import tkinter as tk
from tkinter import filedialog
import pandas as pd
from openpyxl import load_workbook


root = tk.Tk()
root.withdraw()

file_paths = filedialog.askopenfilenames(filetypes=[('Excel Files', '*.xlsx')])


combined_df = pd.DataFrame()

is_first_file = True

for file_path in file_paths:
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names
    if len(sheet_names) >= 2 and is_first_file:  # Проверка на наличие второго листа
        df = pd.read_excel(xls, sheet_name=1, header=None, skiprows=7)  # Чтение второго листа
        combined_df = pd.concat([combined_df, df], ignore_index=True)
        is_first_file = False  # Устанавливаем флаг в False после первого файла
    else:
        df = pd.read_excel(xls, sheet_name=1, header=None, skiprows=8)  # Чтение второго листа
        combined_df = pd.concat([combined_df, df], ignore_index=True)

save_folder = filedialog.askdirectory()

save_path = f"{save_folder}/combined_file.xlsx"
combined_df.to_excel(save_path, index=False)

print("Файл успешно сохранен:", save_path)
