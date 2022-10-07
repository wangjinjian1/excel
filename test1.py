import json,requests

def ansToIndex(str):
    ans = []
    for s in str.split(','):
        ans.append((s, ord(s) - 65))
    return ans

def printAns(id='', preid='', ansDic=None):
    f = open('ans.txt', 'w+', encoding='utf8')
    qress = requests.get(url=f'https://i.kaoshiyun.com.cn/a/{preid}/p/{id}.json?time=855').json()
    qres = []
    for q in qress:
        qres.extend(q['questions'])
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

if __name__=='__main__':
    printAns(id='514b780', preid='a2be50')