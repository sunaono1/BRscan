# -*- coding: utf-8 -*-
"""统一设置页眉：无边框两列表格（左单位名 / 右标题顶右），横线用表格下边框。
   标题统一为「一行不折行」的写法；无页眉的文档一并补上。"""
import os, sys
import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

LEFT = '受检单位'
RIGHTS = {
 '信息安全检查服务·检查报告（模板）.docx': '信息安全检查报告  报告编号：＿＿＿＿＿＿＿＿',
 '信息安全检查报告-内网自检-2026-10-22.docx': '信息安全检查报告  报告编号：SAMPLE-2026-001',
 '信息安全检查服务·现场调研表（进场前填写）.docx': '信息安全检查服务·现场调研表（进场前填写）',
 '信息安全检查服务·现场核查表（访谈部分）.docx': '信息安全检查服务·现场核查表（访谈部分）',
 '信息安全检查服务·现场核查表（技术部分）.docx': '信息安全检查服务·现场核查表（技术部分）',
 '信息安全检查服务·限时免费现场核查手册.docx': '信息安全检查服务·限时免费现场核查手册',
 '信息安全检查服务·增值服务核查手册.docx': '信息安全检查服务·增值服务核查手册',
 '信息安全检查服务·文件清单与用法说明.docx': '信息安全检查服务·文件清单与用法说明',
 '信息安全检查服务现场检测清单10.15.docx': '信息安全检查服务·现场检测清单（客户版）',
}

def el(tag, **a):
    e = OxmlElement(tag)
    for k, v in a.items(): e.set(qn('w:' + k), str(v))
    return e

def run(text):
    r = el('w:r'); rPr = el('w:rPr')
    rPr.append(el('w:rFonts', ascii='宋体', hAnsi='宋体', eastAsia='宋体'))
    rPr.append(el('w:sz', val=18)); rPr.append(el('w:szCs', val=18))
    r.append(rPr); t = el('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text; r.append(t)
    return r

def cell(text, w, align):
    tc = el('w:tc'); tcPr = el('w:tcPr')
    tcPr.append(el('w:tcW', w=w, type='dxa')); tcPr.append(el('w:vAlign', val='bottom'))
    tc.append(tcPr)
    p = el('w:p'); pPr = el('w:pPr')
    pPr.append(el('w:spacing', lineRule='exact', line=240, before=0, after=0))
    if align: pPr.append(el('w:jc', val=align))
    p.append(pPr); p.append(run(text)); tc.append(p)
    return tc

def hdr_table(right, twips):
    lw = 1700                      # 左列固定 3.0cm：只放「受检单位」
    rw = twips - lw
    tbl = el('w:tbl'); pr = el('w:tblPr')
    pr.append(el('w:tblW', w=twips, type='dxa')); pr.append(el('w:tblLayout', type='fixed'))
    b = el('w:tblBorders')
    for s in ('top', 'left', 'right', 'insideH', 'insideV'):
        b.append(el('w:' + s, val='none', sz=0, space=0, color='auto'))
    b.append(el('w:bottom', val='single', sz=6, space=1, color='000000'))
    pr.append(b)
    mar = el('w:tblCellMar')
    for s in ('top', 'left', 'bottom', 'right'): mar.append(el('w:' + s, w=0, type='dxa'))
    pr.append(mar); tbl.append(pr)
    g = el('w:tblGrid'); g.append(el('w:gridCol', w=lw)); g.append(el('w:gridCol', w=rw)); tbl.append(g)
    tr = el('w:tr'); tr.append(cell(LEFT, lw, None)); tr.append(cell(right, rw, 'right')); tbl.append(tr)
    return tbl

def tiny_para():
    p = el('w:p'); pPr = el('w:pPr')
    pPr.append(el('w:spacing', lineRule='exact', line=20, before=0, after=0))
    p.append(pPr); return p

def set_part(hdr, right, twips):
    hdr.is_linked_to_previous = False
    root = hdr.part.element
    for ch in list(root): root.remove(ch)
    root.append(hdr_table(right, twips))
    root.append(tiny_para())

def process(path):
    name = os.path.basename(path)
    right = RIGHTS.get(name)
    if not right:
        return -1
    d = docx.Document(path)
    seen = set(); n = 0
    for sec in d.sections:
        try: part = sec.header.part
        except Exception: continue
        if id(part) in seen: continue
        twips = int((sec.page_width - sec.left_margin - sec.right_margin) / 635)
        set_part(sec.header, right, twips); seen.add(id(part)); n += 1
    d.save(path)
    return n

for p in sys.argv[1:]:
    try:
        n = process(p)
        print('%-52s %s' % (os.path.basename(p), ('页眉已设为表格 %d 节' % n) if n > 0 else '跳过（未登记）'))
    except Exception as e:
        print('%-52s 失败：%s' % (os.path.basename(p), e))
