# -*- coding: utf-8 -*-
"""把页眉从「制表位」改成「无边框两列表格」：左单位名 / 右标题顶右，横线用表格下边框。
   制表位在 Word 里有效，但 Pages / 快速预览不认，表格方案各端一致。"""
import os, sys, glob
import docx
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn('w:' + k), str(v))
    return e

def mk_run(text, size=18):
    r = el('w:r')
    rPr = el('w:rPr')
    f = el('w:rFonts', ascii='宋体', hAnsi='宋体', eastAsia='宋体')
    rPr.append(f)
    rPr.append(el('w:sz', val=size)); rPr.append(el('w:szCs', val=size))
    r.append(rPr)
    t = el('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text
    r.append(t)
    return r

def mk_cell(text, width, align):
    tc = el('w:tc')
    tcPr = el('w:tcPr')
    tcPr.append(el('w:tcW', w=width, type='dxa'))
    tcPr.append(el('w:vAlign', val='bottom'))
    tc.append(tcPr)
    p = el('w:p')
    pPr = el('w:pPr')
    pPr.append(el('w:spacing', lineRule='exact', line=240, before=0, after=0))
    if align: pPr.append(el('w:jc', val=align))
    p.append(pPr)
    p.append(mk_run(text))
    tc.append(p)
    return tc

def build_header_table(left, right, total_twips):
    lw = int(total_twips * 0.23)          # 左列约 23%（够放「受检单位」）
    rw = total_twips - lw
    tbl = el('w:tbl')
    tblPr = el('w:tblPr')
    tblPr.append(el('w:tblW', w=total_twips, type='dxa'))
    tblPr.append(el('w:tblLayout', type='fixed'))
    bdr = el('w:tblBorders')
    for side in ('top', 'left', 'right', 'insideH', 'insideV'):
        bdr.append(el('w:' + side, val='none', sz=0, space=0, color='auto'))
    bdr.append(el('w:bottom', val='single', sz=6, space=1, color='000000'))
    tblPr.append(bdr)
    mar = el('w:tblCellMar')
    for side in ('top', 'left', 'bottom', 'right'):
        mar.append(el('w:' + side, w=0, type='dxa'))
    tblPr.append(mar)
    tbl.append(tblPr)
    grid = el('w:tblGrid')
    grid.append(el('w:gridCol', w=lw)); grid.append(el('w:gridCol', w=rw))
    tbl.append(grid)
    tr = el('w:tr')
    tr.append(mk_cell(left, lw, None))
    tr.append(mk_cell(right, rw, 'right'))
    tbl.append(tr)
    return tbl

def fix_part(hdr, total_twips, tag=''):
    paras = hdr.paragraphs
    if not paras: return 0
    p0 = paras[0]
    text = p0.text
    if '\t' not in text:
        return 0
    left, right = text.split('\t', 1)
    left, right = left.strip(), right.strip()
    tbl = build_header_table(left, right, total_twips)
    p0._p.addprevious(tbl)
    # 原段落留作表后收尾段落：清空内容、去掉边框与制表位、压到 1 磅
    for r in list(p0.runs):
        r._element.getparent().remove(r._element)
    pPr = p0._p.find(qn('w:pPr'))
    if pPr is not None:
        for child in list(pPr):
            if child.tag in (qn('w:pBdr'), qn('w:tabs')):
                pPr.remove(child)
        sp = pPr.find(qn('w:spacing'))
        if sp is not None:
            sp.set(qn('w:lineRule'), 'exact'); sp.set(qn('w:line'), '20')
            sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '0')
    return 1

def process(path):
    d = docx.Document(path)
    n = 0; done = set()
    for sec in d.sections:
        for hdr in (sec.header, sec.first_page_header):
            try:
                part = hdr.part
            except Exception:
                continue
            if id(part) in done: continue
            if hdr.is_linked_to_previous: continue
            twips = int((sec.page_width - sec.left_margin - sec.right_margin) / 635)  # EMU→twips
            if fix_part(hdr, twips):
                n += 1; done.add(id(part))
    if n: d.save(path)
    return n

if __name__ == '__main__':
    paths = sys.argv[1:]
    for p in paths:
        try:
            n = process(p)
            print('%-58s 页眉改为表格 %d 处' % (os.path.basename(p), n))
        except Exception as e:
            print('%-58s 失败：%s' % (os.path.basename(p), e))
