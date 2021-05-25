from openpyxl import load_workbook

filename = './2333.xlsx'

wb = load_workbook(filename)
ws = wb.active

for i in range(3, ws.max_row + 1):
    name = ws.cell(row=i, column=1).value
    if name == '姓名' or name == '':
        continue
    if len(name) == 2:
        ws.cell(row=i, column=1).value = name[0] + '*'
    elif len(name) == 3:
        ws.cell(row=i, column=1).value = name[0] + '**'
    elif len(name) == 4:
        ws.cell(row=i, column=1).value = name[:2] + '**'

wb.save('2222.xlsx')
