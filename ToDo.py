#ToDO: trying to extract the location (row_and_column_number) of detected shapes#

import pandas as pd
import xlwings as xw

excel_file_path = input("Please enter the path to your Excel file (e.g., C:\\Downloads\\Sample_list.xlsx): ")

# Connect to Excel application
app = xw.App(visible=False)

# Open the workbook
wb = app.books.open(excel_file_path)

for sheet in wb.sheets:
    print(f"Sheet: {sheet.name}")

    for shape in sheet.shapes:
        print(f"Shape found in sheet {sheet.name}: {shape.name}")

wb.close()
app.quit()







