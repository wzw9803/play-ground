# -*- coding: utf-8 -*-
"""回填全部定稿答复到交底书 docx(问题文本含 w:ins 修订标记,公式为 OMML)
   用法: python fill_docx.py            → 普通文本插入(已补充解答版)
         python fill_docx.py --revision → 修订形式插入(w:ins,作者=王鑫)
"""
import json
import sys
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

REVISION = '--revision' in sys.argv
# 用法: python fill_docx.py <源docx> <答案json> <输出docx> [--revision]
#       答案json格式: [{"title": "问题标题", "answer": "答复文本"}, ...](可由答复 markdown 解析生成)
if len(sys.argv) < 4:
    print('用法: python fill_docx.py <源docx> <answers.json> <输出docx> [--revision]')
    sys.exit(1)
SRC = sys.argv[1]
ANSWERS_PATH = sys.argv[2]
DST = sys.argv[3]

_ins_id = [1000]
def _next_ins_id():
    _ins_id[0] += 1
    return str(_ins_id[0])

with open(ANSWERS_PATH, encoding='utf-8') as f:
    ANSWERS = json.load(f)

def full_text(p):
    """提取段落完整文本(含 w:ins 内 w:t 及 OMML 数学公式 m:t)"""
    parts = []
    for node in p.iter():
        tag = node.tag
        if tag == qn('w:t') or tag == qn('m:t'):
            if node.text:
                parts.append(node.text)
    return ''.join(parts)

def get_answer(title_kw):
    """按标题关键词查找答复"""
    for it in ANSWERS:
        if title_kw in it['title']:
            return it['answer']
    raise KeyError('未找到答复: ' + title_kw)

# 关键词 → answers.json 索引
MAPPING = [
    ('本方案的执行主体是手机等移动端设备', 0),
    ('本方案能否适用于所有3D虚拟场景的渲染', 1),
    ('待加载的3D场景是由用户指定的', 2),
    ('本申请题目是', 3),
    ('各个预测层的输入都只是当前位置', 4),
    ('三个状态预测层的输入是否不只有当前位置', 5),
    ('三种不同类型的数据如何融合', 6),
    ('T时段内的P_final的最大值', 8),
    ('currentLODError分别表示什么含义', 9),
    ('实质就是待加载的瓦片大小', 10),
    ('是针对待加载3D场景中的所有瓦片进行排序', 11),
    ('7.1.1中哪些步骤存在核心改进', 12),
    ('distance是指什么距离', 14),
    ('tile_count是指总瓦片数量还是', 15),
    ('SSE_threshold_adjusted与SSE_adjusted的关系是', 17),
    ('在3D场景加载流程中的具体应用', 18),
    ('瓦片中心怎么理解', 19),
    ('7.1.2中哪些步骤存在核心改进', 20),
    ('每个瓦片在L0~L3中都对应空间数据', 21),
    ('至少要跳3级么', 22),
    ('两种策略只要满足一种', 23),
    ('保留是指客户端也会下载', 24),
    ('二者缺一不可', 25),
    ('瓦片元数据中携带有', 26),
    ('各个层级对应的所有空间数据', 27),
    ('最近的已加载祖先节点', 28),
    ('子节点数组对应的是', 29),
    ('遍历L1~L3中的所有瓦片', 30),
    ('动态维护该双链关系', 31),
    ('7.1.3中哪些步骤存在核心改进', 32),
    ('释放时机是指首次渲染完成', 33),
    ('钩子的作用是执行如下步骤', 34),
    ('各个参数的含义', 35),
    ('检测目的是', 36),
    ('什么情况下会出现尝试访问CPU数据', 37),
    ('数据交互的方式是', 39),
    ('LRU首次出现', 40),
    ('每个层级都有瓦片', 41),
    ('这里跳级是指', 42),
    ('卸载是指从CPU和GPU中同时删除', 43),
    ('上面哪些步骤存在核心改进', 44),
    # 特殊:同一答复的第二个问题段落
    ('P_kalman(tile, t)、P_markov(tile, t)、P_nn(tile, t)分别表示什么', 6, 'short'),
]

def make_answer_paragraph(text, bold_prefix='【答复】'):
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), '60')
    spacing.set(qn('w:after'), '60')
    pPr.append(spacing)
    p.append(pPr)
    def make_run(text_, bold=False):
        r = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        rF = OxmlElement('w:rFonts')
        rF.set(qn('w:hint'), 'eastAsia')
        rF.set(qn('w:ascii'), '微软雅黑 Light')
        rF.set(qn('w:hAnsi'), '微软雅黑 Light')
        rF.set(qn('w:eastAsia'), '微软雅黑 Light')
        rPr.append(rF)
        if bold:
            rPr.append(OxmlElement('w:b'))
        c = OxmlElement('w:color'); c.set(qn('w:val'), 'C00000'); rPr.append(c)
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '21'); rPr.append(sz)
        szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'), '21'); rPr.append(szCs)
        r.append(rPr)
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = text_
        r.append(t)
        if REVISION:
            ins = OxmlElement('w:ins')
            ins.set(qn('w:id'), _next_ins_id())
            ins.set(qn('w:author'), '王鑫')
            ins.set(qn('w:date'), '2026-08-10T00:00:00Z')
            ins.append(r)
            return ins
        return r
    p.append(make_run(bold_prefix, bold=True))
    p.append(make_run(text))
    return p

doc = Document(SRC)
paras = list(doc.element.body.iter(qn('w:p')))
texts = [full_text(p) for p in paras]

matched = {}
for i, p in enumerate(paras):
    t = texts[i].strip()
    prev = texts[i-1].strip() if i > 0 else ''
    if not t:
        continue
    # 收集该段落应插入的答复(允许同段多问题)
    to_insert = []
    if '以上公式是我们自定义的' in t:
        if 'Q1,以上公式是我们自定义的' in t or t.startswith('Q1'):
            to_insert.append(('idx', 16))
        else:
            # 向上回溯找公式段落(跳过 ``` 等代码块标记与空行)
            for j in range(i - 1, max(0, i - 4), -1):
                pt = texts[j].strip()
                if not pt or pt.startswith('```'):
                    continue
                if 'Priority' in pt or 'P_final' in pt:
                    to_insert.append(('idx', 7))
                elif 'SSE_threshold_adjusted' in pt:
                    pass
                elif 'SSE_adjusted' in pt:
                    to_insert.append(('idx', 13))
                break
    if '7.1.4中哪些步骤存在核心改进' in t:
        to_insert.append(('idx', 38))
    for kw, idx, *rest in MAPPING:
        if kw in t:
            to_insert.append(('short', idx) if rest else ('idx', idx))
    if not to_insert:
        continue
    # 构建实际插入文本(去重 + 检查是否已插入)
    texts_to_add, record = [], []
    seen = set()
    for kind, idx in to_insert:
        key = kind if kind != 'idx' else idx
        if key in seen:
            continue
        seen.add(key)
        if kind == 'idx':
            if idx in matched:
                continue
            texts_to_add.append(ANSWERS[idx]['answer'])
        elif kind == 'short':
            if ('short', idx) in matched:
                continue
            texts_to_add.append('同上一问:三种异构预测输出(P_kalman / P_markov / P_nn)统一为瓦片访问概率的具体计算方式,见上方【答复】。')
        record.append(key)
    if not texts_to_add:
        continue
    anchor = p
    for txt in texts_to_add:
        np = make_answer_paragraph(txt)
        anchor.addnext(np)
        anchor = np
    for k in record:
        matched[k] = matched.get(k, 0) + 1

print('插入统计(%d 处):' % sum(matched.values()))
for k in sorted(matched, key=str):
    print(' ', k, '->', matched[k])
missing = [i for i in range(45) if i not in matched]
print('缺失答复索引:', missing if missing else '无')
print('short(同上一问):', matched.get(('short', 6), 0))

doc.save(DST)
print('saved:', DST)
