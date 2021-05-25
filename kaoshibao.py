from openpyxl import workbook, load_workbook, Workbook
import random

dicSelect = {
    1: 'A',
    2: 'B',
    3: 'C',
    4: 'D',
    12: 'AB',
    13: 'AC',
    14: 'AD',
    23: 'BC',
    24: 'BD',
    34: 'CD',
    123: 'ABC',
    124: 'ABD',
    134: 'ACD',
    234: 'BCD',
    1234: 'ABCD',
}

dicJudge = {
    '正确': 'A',
    '错误': 'B',
    'A': 'A',
    'B': 'B'
}

listtitle = ['题干（必填）', '题型（必填）', '选项A', '选项B', '选项C', '选项D', '选项E', '选项F', '选项G', '选项H', '正确答案（必填）', '解析', '章节', '难度']

listType = ['单选题', '多选题', '判断题']


def handeleExcel(ws):
    lenlist = len(listtitle)
    for i in range(1, 1 + lenlist):
        ws.cell(row=1, column=i).value = listtitle[i - 1]
        answer = ws.cell(row=i, column=8)


# 安规选择题
def fun1(path):
    wb = load_workbook(path)
    ws = wb.active
    newwb = Workbook()
    newws = newwb.active
    handeleExcel(newws)
    for row in range(2, ws.max_row + 1):
        newws.cell(row=row, column=1).value = ws.cell(row=row, column=5).value
        answer = ws.cell(row=row, column=11).value
        newws.cell(row=row, column=11).value = dicSelect[answer]
        if answer <= 4:
            newws.cell(row=row, column=2).value = listType[0]
        else:
            newws.cell(row=row, column=2).value = listType[1]
        for i in range(3, 7):
            newws.cell(row=row, column=i).value = ws.cell(row=row, column=i + 3).value
    newwb.save('@' + path)


# 安规判断题
def fun2(path):
    wb = load_workbook(path)
    ws = wb.active
    newwb = Workbook()
    newws = newwb.active
    handeleExcel(newws)
    for row in range(2, ws.max_row + 1):
        newws.cell(row=row, column=1).value = ws.cell(row=row, column=4).value
        answer = ws.cell(row=row, column=5).value
        newws.cell(row=row, column=11).value = dicJudge[answer]
        newws.cell(row=row, column=2).value = listType[2]
        newws.cell(row=row, column=3).value = '正确'
        newws.cell(row=row, column=4).value = '错误'
    newwb.save('@' + path)


# 专业
def fun3(path):
    wb = load_workbook(path)
    ws = wb.active
    newwb = Workbook()
    newws = newwb.active
    handeleExcel(newws)
    for row in range(3, ws.max_row):
        if ws.cell(row=row, column=1).value == None:
            break
        # 题干
        newws.cell(row=row, column=1).value = ws.cell(row=row, column=7).value
        # 答案
        newws.cell(row=row, column=11).value = ws.cell(row=row, column=9).value
        # 题型
        newws.cell(row=row, column=2).value = ws.cell(row=row, column=6).value
        # 解析
        newws.cell(row=row, column=12).value = ws.cell(row=row, column=13).value
        values = ws.cell(row=row, column=8).value
        arrs = values.split('$;$')
        # 选项
        if len(arrs) > 2:
            for index in range(3, len(arrs) + 3):
                newws.cell(row=row, column=index).value = arrs[index - 3]
        else:
            newws.cell(row=row, column=3).value = '正确'
            newws.cell(row=row, column=4).value = '错误'
    newwb.save('@' + path)


if __name__ == '__main__':
    # fun2('安规普考_变电判断题库.xlsx')
    # fun1('安规普考_变电选择题库.xlsx')
    fun3('抄表核算收费员网大版题库.xlsx')
