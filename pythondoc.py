from docx import Document
import os


dic1 = {'一、', '二、', '三、', '四、', '五、', '六、', '七、'}
path = '/Users/wangjinjian/Desktop/233/技能等级评价--上海公司题库--笔试/变配电运行值班员(变电站运行值班员)-电力调度员（主网）/地级调度--电力调度员（主网）技能笔试题库1.docx'
doc = Document(path)
cnt = 1
for con in doc.paragraphs:
    content = con.text
    if len(con.text) != 0:
        if content[0] == 'J':
            content = str(cnt) + '、' + content[10:]
            cnt += 1
            con.text = content
        elif content[:2] in dic1:
            con.text = ''
paths = path.split('/')
lenpath = len(paths)
fpath = '/'
for i in range(lenpath - 1):
    fpath = os.path.join(fpath, paths[i])
fpath = os.path.join(fpath, '@' + paths[lenpath - 1])
doc.save(fpath)
