from docx import Document
import os

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
dic1 = {'一、', '二、', '三、', '四、', '五、', '六、', '七、'}
path = '/Users/wangjinjian/Desktop/23/技能等级评价--上海公司题库--笔试/供电服务员(抄表核算收费员)-抄表核算收费员/抄表核算收费员技能笔试题库.docx'
doc = Document(path)
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
        elif content[:2] in dic1:
            con.text = ''
paths = path.split('/')
lenpath = len(paths)
fpath = '/'
for i in range(lenpath - 1):
    fpath = os.path.join(fpath, paths[i])
fpath = os.path.join(fpath, '@' + paths[lenpath - 1])
doc.save(fpath)
