# from openpyxl import load_workbook

# wb = load_workbook(filename='translate.xlsx')
# ws = wb.active

# ws.cell(column=3, row=3, value="kel")
# print(ws.cell(column=3, row=3).value)
# ws.cell(column=3, row=3, value="kel")

# print(ws.min_column)

# wb.save('translate.xlsx')
import mtranslate
import time

print(mtranslate.translate("Hello word", "tr", "auto"))
time.sleep(5)
