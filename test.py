import requests, json, re
from collections import defaultdict
from openpyxl import Workbook


def charToIndex(ans):
    return ord(ans) - 65


def getQues(id='f599a5'):
    queanswer = f'https://i.kaoshiyun.com.cn/a/{id}/p/{id}.json?time=635'
    queanswer = f'https://i.kaoshiyun.com.cn/a/29c8ff/a/1904be.json?time=322'
    res = requests.get(url=queanswer)
    res.encoding = 'utf8'
    try:
        queRes = res.json()
        for re1 in queRes:
            for re in re1['questions']:
                print(re[0]['options'][0].encode('utf8').decode('unicode_escape'))
        print(queRes)
    except Exception:
        print(res.text)


def test1(id='19ada9'):
    answerurl = f'https://i.kaoshiyun.com.cn/a/{id}/a/{id}.json?time=635'
    queanswer = f'https://i.kaoshiyun.com.cn/a/{id}/p/{id}.json?time=348'
    ansRes = requests.get(url=answerurl).json()
    queRes = requests.get(url=queanswer).json()
    quesLi = []
    for q in queRes:
        quesLi.extend(q['questions'])
    with open('ikaoshiyun.txt', 'w+', encoding='utf-8') as f:
        for index, re in enumerate(zip(quesLi, ansRes)):
            f.write(f"{index + 1}  {re[1]['s']}   {'_' * 4} {re[0]['content']} \n")


def test(id='19ada9'):
    answerurl = f'https://i.kaoshiyun.com.cn/a/{id}/a/{id}.json?time=635'
    queanswer = f'https://i.kaoshiyun.com.cn/a/{id}/p/{id}.json?time=348'
    ansRes = requests.get(url=answerurl).json()
    queRes = requests.get(url=queanswer).json()
    quesLi = []
    for q in queRes:
        quesLi.extend(q['questions'])
    with open('ikaoshiyun.txt', 'w+', encoding='utf-8') as f:
        for index, re in enumerate(zip(quesLi, ansRes)):
            if 'options' in re[0]:
                f.write(f"{index + 1} {'_' * 4} {re[0]['content']} \n")
                f.write(f" {re[1]['s']}  {re[0]['options'][0].encode('utf8').decode('unicode_escape')}   \n")
            else:
                f.write(f"{index + 1} {'_' * 4} {re[0]['content']} \n")
                f.write(
                    f"{re[1]['s']}   {'_' * 4} {re[0]['content']} \n")


def ansToIndex(str):
    ans = []
    for s in str.split(','):
        ans.append((s, ord(s) - 65))
    return ans


# Multi   Single    Fill   Judge

def getQ(ids, idde='0c8e87'):
    typeDic = {
        'Single': '?????????',
        'Multi': '?????????',
        'Judge': '?????????',
        'Fill': '?????????'
    }
    wb = Workbook()
    ws = wb.active
    ws.append(['??????', '????????????', '??????A', '??????B', '??????C', '??????D', '??????E', '??????', '??????', '??????', '??????', '??????'])
    with open('ans.json', 'r', encoding='utf8') as f:
        ansDic = json.load(f)
    with open('pp.json', 'r', encoding='utf8') as f:
        ppDic = json.load(f)
    for id in ids:
        res = requests.get(f'https://i.kaoshiyun.com.cn/a/{idde}/p/{id}.json?time=200')
        ans = res.json()[0]['questions']
        for q in ans:
            wsc = []
            wsc.append(typeDic[q['type']])
            wsc.append(q['content'])
            ansOpt = []
            if q['type'] != 'Fill':
                opts = json.loads(q['options'][0])
                for opt in opts:
                    for v in opt.values():
                        wsc.append(v)
                wsc.extend([''] * (5 - len(opts)))
                wsc.append(ansDic[q['qid']].replace(',', ''))
                wsc.append(1)
                wsc.append('ttt')
                wsc.append('easy')
                wsc.append('233')
                try:
                    for index in ansToIndex(ansDic[q['qid']]):
                        ansOpt.append(opts[index[1]][index[0]])
                except Exception:
                    print(q['qid'], q['options'][0])
                if q['qid'] not in ppDic:
                    ppDic[q['qid']] = {}
                    ppDic[q['qid']]['q'] = q['content']
                    ppDic[q['qid']]['a'] = '|'.join(ansOpt)
            else:
                wsc.extend([''] * 5)
                wsc.append(ansDic[q['qid']])
                wsc.append(1)
                wsc.append('ttt')
                wsc.append('easy')
                wsc.append('233')
                if q['qid'] not in ppDic:
                    ppDic[q['qid']] = {}
                    ppDic[q['qid']]['q'] = q['content']
                    ppDic[q['qid']]['a'] = ansDic[q['qid']]
            wsc.append(q['qid'])
            try:
                ws.append(wsc)
            except Exception:
                print(wsc)
    with open('pp.json', 'w+', encoding='utf8') as f:
        json.dump(ppDic, f, ensure_ascii=False)
    wb.save('1.xlsx')


def initTiKu(id='', idde='0c8e87'):
    ansDic = {}
    res = requests.get(f'https://i.kaoshiyun.com.cn/a/{idde}/a/{id}.json?time=923')
    for ans in res.json():
        ansDic[ans['q']] = ans['s']
    with open('ans.json', 'w+', encoding='utf8') as f:
        json.dump(ansDic, f, ensure_ascii=False)


def initId(save=False):
    char = re.compile('??? (.*) ???')
    ansDic = defaultdict(list)
    res = requests.get('https://i.kaoshiyun.com.cn/a/0c8e87/0c8e87_list.json?time=999').json()
    # res = requests.get('https://i.kaoshiyun.com.cn/a/29c8ff/29c8ff_list.json?time=999').json()
    for a in res:
        key = char.search(a['parentNodeName']).group(1)
        for id in a['item']:
            ansDic[key].append(id['chapterNodeID'])
    if save:
        with open('id.json', 'w+', encoding='utf8') as f:
            json.dump(ansDic, f, ensure_ascii=False)
    else:
        print(ansDic)


def printAns(id='2cc774', preid='29c8ff', ansDic=None):
    f = open('ans.txt', 'w+', encoding='utf8')
    qres = requests.get(url=f'https://i.kaoshiyun.com.cn/a/{preid}/p/{id}.json?time=855').json()[0]['questions']
    ares = requests.get(url=f'https://i.kaoshiyun.com.cn/a/{preid}/a/{id}.json?time=698').json()
    ansDic = {}
    if not ansDic:
        for ans in ares:
            ansDic[ans['q']] = ans['s']
        with open('jsdic.json', 'w+', encoding='utf8') as f1:
            json.dump(ansDic, f1, ensure_ascii=False)
    for q in qres:
        ansOpt = []
        if q['type'] != 'Fill':
            opts = json.loads(q['options'][0])
            try:
                for index in ansToIndex(ansDic[q['qid']]):
                    ansOpt.append(opts[index[1]][index[0]])
            except Exception:
                print(q['content'], ansDic[q['qid']])
            f.write(q['qid'] + '\t' + q['content'] + '\n')
            f.write('\t' + '|'.join(ansOpt) + '\n')
        else:
            f.write(q['qid'] + '\t' + q['content'] + '\n')
            f.write('\t' + ansDic[q['qid']] + '\n')
    f.close()


if __name__ == '__main__':
    with open('id.json', 'r', encoding='utf8') as f:
        idDic = json.load(f)
    with open('ans.json', 'r', encoding='utf8') as f:
        ansDic = json.load(f)
    # initTiKu()
    # initId()
    # getQ(idDic['1'])
    # getQues('1904be')
    # initId()
    printAns(id='4225c2', preid='f7016b')
