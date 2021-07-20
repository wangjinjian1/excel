from docx import Document
import os

dic = {'一、', '二、', '三、', '四、', '五、', '六、', '七、'}
char1 = 'J'
char2 = '答案'


class docxx:

    def __init__(self, path, ignore=('绘图题')):
        self.path = path
        paths = path.split('/')
        lenpath = len(paths)
        fpath = '/'
        for i in range(lenpath - 1):
            fpath = os.path.join(fpath, paths[i])
        self.savepath = os.path.join(fpath, '@' + paths[lenpath - 1])
        self.type = type
        self.ignore = ignore
        self.doc = Document(path)

    def fun(self):
        paras = self.doc.paragraphs
        cnt = 1
        palen = len(paras)
        skiptimes = 0
        skip = False
        for i in range(palen):
            if skiptimes != 0:
                skiptimes -= 1
                continue
            if len(paras[i].text) != 0:
                if paras[i].text[:2] in dic:
                    if paras[i].text[2:] in self.ignore:
                        skip = True
                    else:
                        skip = False
                    paras[i].text = ''
                    continue
                if skip:
                    if paras[i].text[0] == char1:
                        content = str(cnt) + '、' + paras[i].text[10:]
                        paras[i].text = content
                        cnt += 1
                    continue
                if paras[i].text[0] == char1:
                    content = str(cnt) + '、' + paras[i].text[10:]
                    paras[i].text = content
                    cnt += 1
                elif paras[i].text[:2] == char2:
                    temp = paras[i].text
                    anindex = i
                    while (anindex + 2 < palen and len(paras[anindex + 1].text) != 0):
                        temp += ' ' * 5 + paras[anindex + 1].text
                        paras[anindex + 1].text = ''
                        anindex += 1
                        skiptimes += 1
                    if anindex > i:
                        paras[i].text = temp
        self.save()

    def save(self):
        self.doc.save(self.savepath)


if __name__ == '__main__':
    path1 = '/Users/wangjinjian/Desktop/23/技能等级评价--上海公司题库--笔试/供电服务员(抄表核算收费员)-抄表核算收费员/抄表核算收费员技能笔试题库.docx'
    path = '/Users/wangjinjian/Desktop/23/技能等级评价--上海公司题库--笔试/供电服务员(装表接电工)-装表接电工/装表接电工技能笔试题库.docx'
    docxx(path).fun()
