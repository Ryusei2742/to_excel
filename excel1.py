

import openpyxl
import pandas as pd
import glob

export_file_path = '' #出力先
import_folder_path = '' #読み込み元
path = import_folder_path + '/' + '*.xlsx'
file_path = glob.glob(path)
# print(file_path)

df_concat = pd.DataFrame()

for i in file_path:
    df_read_excel = pd.read_excel(i)
    # print(df_read_excel.head(3))
df_concat = pd.concat([df_read_excel, df_concat])
    # print(df_concat)
df_drop = df_concat.drop('Unnamed: 0', axis = 1)
# print(df_drop.head(3))
df_sort = df_drop.sort_values(by = '数量', ascending = False)
# print(df_sort)

# df_sort.to_excel(export_file_path + '/' + 'sample_te.xlsx')

workbook = openpyxl.load_workbook(export_file_path + '/' + 'sample_tes.xlsx')
worksheet = workbook.worksheets[0]
worksheet.delete_cols(1)
workbook.save(export_file_path + '/' + 'sample.xlsx')
