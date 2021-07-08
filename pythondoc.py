from docx import Document

dic = {
    '(1)': '(一)',
    '(2)': '(二)',
    '(3)': '(三)',
    '(4)': '(四)',
    '(5)': '(五)',
    '(6)': '(六)',
    '1)': '一)',
    '2)': '二)',
    '3)': '三)',
    '4)': '四)',
    '5)': '五)',
    '6)': '六)',
    '1、': '一、',
    '2、': '二、',
    '3、': '三、',
    '4、': '四、',
    '5、': '五、',
    '1）': '一)',
    '2）': '二)',
    '3）': '三)',
    '4）': '四)',
    '5）': '五)',
    '6）': '六)',
    '1.': '一、',
    '2.': '二、',
    '3.': '三、',
    '4.': '四、',
    '5.': '五、',
    '（1）': '(一)',
    '（2）': '(二)',
    '（3）': '(三)',
    '（4）': '(四)',
    '（5）': '(五)',
    '（6）': '(六)',
}

doc = Document('3.docx')
cnt = 1
for con in doc.paragraphs:
    content = con.text
    if len(con.text) != 0:
        if content[0] == 'J':
            content = str(cnt) + '、' + content[10:]
            cnt += 1
            con.text = content
        elif content[0:2] == '答：':
            content = content[2:]
            con.text = content
        elif content[0:3] in dic.keys():
            con.text = dic[content[0:3]] + content[3:]
        elif content[0:2] in dic.keys():
            con.text = dic[content[0:2]] + content[2:]
doc.save('@3.docx')
