# Excel自動化

import openpyxl
import pandas as pd
import glob


import_file_path = '/Users/asd2f/Python/sample-1.xlsx'
excel_sheet_name = '発注管理表'
export_file_path = '/Users/asd2f/Python/to_excel/output_kino3'

df_order = pd.read_excel(import_file_path, sheet_name = excel_sheet_name)
# print(df_order)

company_name = df_order['会社名'].unique()
# print(company_name)
# print(type(company_name))	'numpy.ndarray'
# print(type(df_order))		'pandas.core.frame.DataFrame'

# print(df_order[df_order['会社名'] == '株式会社A'])

for i in company_name:
	df_order_company = df_order[df_order['会社名'] == i]
	# print(df_order_company)
	df_order_company.to_excel('/Users/asd2f/Python/to_excel' + '/' + i + '.xlsx')


















