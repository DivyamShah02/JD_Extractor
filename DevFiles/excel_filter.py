import os
from library.Config import Config
import pandas as pd

# Path to the Excel file
file_path = input('Excel file path : ')
file_path = file_path.replace('"','')

sorted_excel_name = input('''Sorted excel name (Business_Place) & don't include ".xlsx" : ''')
sorted_excel_name = sorted_excel_name + '.xlsx'


cwd_path = os.getcwd()
config_path = cwd_path.replace('DevFiles', 'BotConfig\\config.ini')
config = Config(filename=config_path)

# Read the Excel file
df = pd.read_excel(file_path)

# Select columns with indices
alpha_col = {
    'a':0,
    'b':1,
    'c':2,
    'd':3,
    'e':4,
    'f':5,
    'g':6,
    'h':7,
    'i':8,
    'j':9,
    'k':10,
    'l':11,
    'm':12,
    'n':13,
    'o':14,
    'p':15,
    'q':16,
    'r':17,
    's':18,
    't':19,
    'u':20,
    'v':21,
    'w':22,
    'x':23,
    'y':24,
    'z':25,
}
selected_columns = df.iloc[:, [alpha_col['a'], alpha_col['n'], alpha_col['o']]]
# Rename the selected columns

new_column_names = ['Link', 'Name', 'Address']
selected_columns.columns = new_column_names

# Remove rows with any missing values in the selected columns
selected_columns = selected_columns.dropna(subset=['Link'])

sorted_file_path = os.path.join(config.paths.unprocessed_path, sorted_excel_name)
selected_columns.to_excel(sorted_file_path, index=False, engine='openpyxl')

print(f"Excel file '{sorted_file_path}' has been generated with the selected columns, removed empty rows, and new names.")
