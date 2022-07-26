const ansDic = {"bf900a": "B", "8fe31fd": "A,B,C", "18db7b": "B", "6bc53e": "D", "691d7ea": "C", "e4aa82": "A", "ebd31a6": "A", "6eefe9": "A", "781138": "C", "6df425": "C", "e507a3": "A", "18dcaa5": "C", "ecaab6": "A", "5d55d5f": "C", "901b05": "A", "b28925f": "D", "1d07134": "D", "d04b932": "C", "975c061": "D", "7a5f23c": "A,C,D", "4136d6": "A", "13cbbaa": "A", "640e28": "D", "d0f138": "A,B,C,D", "6cb9a5b": "A,B,C", "d76dd13": "D", "3af7ea": "A,B,C,D", "870fdd2": "A", "870ee9": "A", "7ea1c46": "A,B,C", "462265c": "B", "b2556122": "B,C,D", "421ab12": "B,C,D", "897512": "B", "101a6cf": "B", "fbdcd9": "B", "21be2c6": "非结构化数据", "014bdab": "A", "235b585": "A", "9e7eda": "C", "6ebb5c": "B", "d18a80": "A,B,C", "0e2bb2": "A", "f1b6037": "C", "36814f3": "HBase&分布式", "a2d7b1": "D", "d1bc2a": "C", "374a01": "A,B,C,D", "638c061": "B", "388beb0": "A", "658ffc8": "B", "7c03fad": "B", "90f6b6": "A", "3d7b3e": "A,B,C,D", "a8c60bf": "D", "4721dd": "C", "fe72bc": "聚类", "464674": "B", "f6da63": "B,C,D", "c1a8bc": "C", "77c621": "C,D", "b3c968f": "B", "2052cba": "B,C", "9a52fd7": "A,B", "7a66a5": "B", "002f408": "B", "4f7cae3": "B", "04dad17": "C", "0adf44": "D", "d3c601": "B", "e70f81": "C", "8c7155": "A,B,C,D,E", "74dacb3": "A,B,C,D", "07bddee": "B,C,D", "9c0a04c": "D", "69c4c3": "A,B,C,D", "c9a31a": "D", "8e6037": "B", "107e9d": "B", "fa5d5b": "A", "148218": "A,B,C", "186f04": "A", "07e4c1c": "A", "42d51a": "A", "9e0d5a": "C", "cf24e8": "C", "e9ac09": "B", "6e150b3": "数据可视化", "0af471": "应用价值高", "aeee1ba": "A", "5e0751": "A", "1ec757a": "A,B,D", "bbb783b": "A,C", "b094f4": "A", "dc66d35": "C", "be098c": "A", "ed5218": "B", "4fa44a": "A,B,C,D", "59fe7f8": "B", "761a22b": "A,B,C,D"}
const arrs = document.querySelectorAll('[class="subject-box-content border-bottom swiper-slide"]')
for (let arr of arrs) {
    if (arr.getAttribute('type') == "Single" || arr.getAttribute('type') == "Judge") {
        if (arr.getAttribute('qid') in ansDic) {
            let opts = arr.querySelectorAll('[class="ksy-flex"]>span')
            for (let opt of opts) {
                if (opt.getAttribute('value') == ansDic[arr.getAttribute('qid')]) {
                    opt.setAttribute('class', 'option-strong choice-question')
                }
            }
        } 
    }else if (arr.getAttribute('type') == "Multi") {
        if (arr.getAttribute('qid') in ansDic) {
            let selects = new Set(ansDic[arr.getAttribute('qid')].split(','))
            let opts = arr.querySelectorAll('[class="ksy-flex"]>span')
            for (let opt of opts) {
                if (selects.has(opt.getAttribute('value'))) {
                    opt.setAttribute('class', 'option-strong-square choice-question')
                }
            }
        }
    }
}