import json, re, os
from collections import defaultdict
from openpyxl import load_workbook


#233.txt 放题目
#把excel放在tiku下，记得用kaoshibao.py转化
#initTiKu  录入题库
abcToIndex = {
    'A': 3,
    'B': 4,
    'C': 5,
    'D': 6,
    'E': 7,
    'F': 8,
    'G': 9,
    'H': 10,
}
patternTitle = re.compile('[（）。，！,.() /《》<>、：:;；]')

qq = defaultdict()
with open('233.txt', 'r', encoding='utf-8') as f:
    questions = json.load(f)['data']['question']
    for q in questions:
        qq[int(q['SERIAL_NUMBER'])] = patternTitle.sub('',q['QUESTION_CONTENT']).strip().replace('　　', '')


# type 1 判断  type 2 单选  type 3 多选
def initTiKu(excelpath='tiku'):
    tiku = defaultdict(dict)
    for excel in os.listdir(excelpath):
        wb = load_workbook(os.path.join(excelpath, excel))
        ws = wb.active
        for i in range(2, ws.max_row + 1):

            if ws.cell(row=i, column=1) == None or ws.cell(row=i, column=1).value == '':
                break
            title = patternTitle.sub('', ws.cell(row=i, column=1).value).replace('　　', '')
            type = ws.cell(row=i, column=2).value.strip()
            answer = ws.cell(row=i, column=11).value.strip()
            tiku[title]['answer'] = answer
            if type == '判断题':
                tiku[title]['type'] = 1
            elif type == '单选题':
                tiku[title]['type'] = 2
                if ws.cell(row=i, column=abcToIndex[answer]).value == None:
                    tiku[title]['content'] = '竟然没答案，愚蠢'
                else:
                    tiku[title]['content'] = ws.cell(row=i, column=abcToIndex[answer]).value.strip()
            elif type == '多选题':
                tiku[title]['type'] = 3
                contents = []
                for j in answer:
                    if ws.cell(row=i, column=abcToIndex[j]).value == None:
                        contents.append('竟然没答案，愚蠢')
                    else:
                        contents.append(ws.cell(row=i, column=abcToIndex[j]).value.strip())
                tiku[title]['content'] = ' | '.join(contents)
    with open('tiku.json', 'w+', encoding='utf-8') as f:
        json.dump(tiku, f, ensure_ascii=False)

def getTiku():
    with open('tiku.json','r',encoding='utf-8') as f:
        tiku=json.load(f)
    return tiku


if __name__ == '__main__':
    # initTiKu()
    tiku=getTiku()
    for k,v in qq.items():
        print(k,tiku[v]['answer'],v)
