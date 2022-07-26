import openpyxl as xl

wb = xl.load_workbook('transactions.xlsx')
sheet = wb['Sheet1']
cell = sheet['a1']
cell = sheet.cell(1,1)


print(sheet.max_row)
print(sheet.max_column)

for row in range(2, sheet.max_row + 1):
    cell = sheet.cell(row, 3)
    print(cell.value)
    corrected_price = cell.value * .9
    corrected_cell = sheet.cell(row, 4)
    corrected_cell.value = corrected_price

wb.save('transactions2.xlsx')
