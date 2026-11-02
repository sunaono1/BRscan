# -*- coding: utf-8 -*-
"""生成《信息安全检查服务 · 授权书》——受检单位签署，授权我方在约定范围与时间窗内开展信息安全检查。"""
import io, os
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

P   = os.environ.get('BRSCAN_ROOT', '.')   # 工作根目录：含交付物/与数据源，默认当前目录
OUT = P + '/2.0-精简核查体系/交付物/01-进场前（发客户）/信息安全检查服务·授权书.docx'
MD  = P + '/2.0-精简核查体系/交付物/文档介绍/信息安全检查服务·授权书.md'

HEI, SONG, XBS = '黑体', '宋体', '方正小标宋简体'
FS_H, FS_B, FS_TB = 14, 14, 12

doc = docx.Document()
st = doc.styles['Normal']; st.font.name = SONG; st.font.size = Pt(FS_B)
st.element.rPr.rFonts.set(qn('w:eastAsia'), SONG)

def setfont(r, name, size, bold=False):
    r.font.name = name; r.font.size = Pt(size); r.font.bold = bold
    r._element.rPr.rFonts.set(qn('w:eastAsia'), name)

def setpara(p, line=Pt(23), indent=0, before=0, after=0, align=None):
    pf = p.paragraph_format
    if line: pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY; pf.line_spacing = line
    if indent: pf.first_line_indent = Pt(indent)
    pf.space_before, pf.space_after = Pt(before), Pt(after)
    if align is not None: p.alignment = align
    return p

def para(t, name=SONG, size=FS_B, bold=False, indent=0, align=None, line=Pt(23), before=0, after=0):
    p = doc.add_paragraph(); setpara(p, line, indent, before, after, align)
    setfont(p.add_run(t), name, size, bold); return p

def h(t):
    return para(t, HEI, FS_H, False, 0, WD_ALIGN_PARAGRAPH.LEFT, Pt(23), 8, 4)

def table(cols, rows, widths, size=FS_TB, hl=18):
    t = doc.add_table(rows=1, cols=len(cols)); t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
    lay = OxmlElement('w:tblLayout'); lay.set(qn('w:type'), 'fixed'); t._tbl.tblPr.append(lay)
    for i, c in enumerate(cols):
        cell = t.rows[0].cells[i]; cell.text = ''
        p = cell.paragraphs[0]; setpara(p, Pt(16), 0, 2, 2, WD_ALIGN_PARAGRAPH.CENTER)
        setfont(p.add_run(c), HEI, size)
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r):
            p = cells[i].paragraphs[0]; setpara(p, Pt(hl), 0, 2, 2, WD_ALIGN_PARAGRAPH.CENTER if i == 0 else None)
            setfont(p.add_run(str(v)), SONG, size)
    for row in t.rows:
        for i, w in enumerate(widths): row.cells[i].width = Cm(w)
    for i, w in enumerate(widths): t.columns[i].width = Cm(w)
    return t

# 页面
s = doc.sections[0]
s.page_width, s.page_height = Cm(21), Cm(29.7)
s.top_margin = s.bottom_margin = Cm(2.4); s.left_margin = s.right_margin = Cm(3.17)
s.header_distance, s.footer_distance = Cm(1.5), Cm(1.75)

# 页眉：表格写法（左单位名 / 右标题顶右）
def _el(tag, **a):
    e = OxmlElement(tag)
    for k, v in a.items(): e.set(qn('w:' + k), str(v))
    return e
def _r9(t):
    r = _el('w:r'); p = _el('w:rPr')
    p.append(_el('w:rFonts', ascii='宋体', hAnsi='宋体', eastAsia='宋体'))
    p.append(_el('w:sz', val=18)); p.append(_el('w:szCs', val=18))
    r.append(p); x = _el('w:t'); x.set(qn('xml:space'), 'preserve'); x.text = t; r.append(x); return r
def _cell(t, w, jc):
    tc = _el('w:tc'); pr = _el('w:tcPr')
    pr.append(_el('w:tcW', w=w, type='dxa')); pr.append(_el('w:vAlign', val='bottom')); tc.append(pr)
    p = _el('w:p'); pp = _el('w:pPr')
    pp.append(_el('w:spacing', lineRule='exact', line=240, before=0, after=0))
    if jc: pp.append(_el('w:jc', val=jc))
    p.append(pp); p.append(_r9(t)); tc.append(p); return tc
_tw = int((s.page_width - s.left_margin - s.right_margin) / 635); _lw = 1700
root = s.header.part.element
for c in list(root): root.remove(c)
tbl = _el('w:tbl'); pr = _el('w:tblPr')
pr.append(_el('w:tblW', w=_tw, type='dxa')); pr.append(_el('w:tblLayout', type='fixed'))
b = _el('w:tblBorders')
for side in ('top', 'left', 'right', 'insideH', 'insideV'):
    b.append(_el('w:' + side, val='none', sz=0, space=0, color='auto'))
b.append(_el('w:bottom', val='single', sz=6, space=1, color='000000')); pr.append(b)
mar = _el('w:tblCellMar')
for side in ('top', 'left', 'bottom', 'right'): mar.append(_el('w:' + side, w=0, type='dxa'))
pr.append(mar); tbl.append(pr)
g = _el('w:tblGrid'); g.append(_el('w:gridCol', w=_lw)); g.append(_el('w:gridCol', w=_tw - _lw)); tbl.append(g)
tr = _el('w:tr'); tr.append(_cell('受检单位', _lw, None)); tr.append(_cell('信息安全检查服务·授权书', _tw - _lw, 'right'))
tbl.append(tr); root.append(tbl)
tail = _el('w:p'); tp = _el('w:pPr'); tp.append(_el('w:spacing', lineRule='exact', line=20, before=0, after=0)); tail.append(tp); root.append(tail)
# 页脚页码
fp = s.footer.paragraphs[0]; setpara(fp, Pt(12), 0, 0, 0, WD_ALIGN_PARAGRAPH.CENTER)
fl = _el('w:fldSimple', instr=' PAGE ')
rr = _el('w:r'); rp = _el('w:rPr'); rp.append(_el('w:rFonts', eastAsia='宋体')); rp.append(_el('w:sz', val=18)); rr.append(rp)
fl.append(rr); fp._p.append(fl)

MDL = []
def md(x=''): MDL.append(x)

para('编号：＿＿＿＿＿＿＿＿', SONG, 10.5, False, 0, WD_ALIGN_PARAGRAPH.LEFT, Pt(16))
para('授　权　书', XBS, 16, True, 0, WD_ALIGN_PARAGRAPH.CENTER, Pt(30), 0, 12)
md('# 授　权　书'); md()

body = ('为提升本单位信息安全防护水平、发现并整改安全隐患，本单位（以下称「授权方」）同意并授权'
        '受托方单位技术部门（以下称「受托方」）对下列受检对象开展信息安全检查，'
        '并对授权范围内的检查动作、账号使用、数据处理与报告使用作出如下约定。')
para(body, SONG, FS_B, False, 28)
md(body); md()

h('一、授权方与受托方')
table(['项目', '内容'],
      [['授权方（受检单位）', ''],
       ['受托方', '受托方单位技术部门'],
       ['授权方对接人（姓名·职务·电话）', ''],
       ['受托方项目负责人（姓名·电话）', '']],
      [6.0, 8.6], hl=20)
md('**一、授权方与受托方**'); md(); md('| 项目 | 内容 |'); md('|---|---|')
md('| 授权方（受检单位） |  |'); md('| 受托方 | 受托方单位技术部门 |')
md('| 授权方对接人 |  |'); md('| 受托方项目负责人 |  |'); md()

h('二、授权检查范围')
para('本次授权范围以双方确认的受检资产清单为准（清单见《信息安全检查服务·现场调研表（进场前填写）》第二部分），'
     '核心受检对象如下；清单如有增减，以双方签字确认的版本为准。', SONG, FS_B, False, 28)
table(['序号', '受检对象 / 系统', 'IP / 域名 / 网段', '授权动作（勾选）'],
      [['1', '', '', '□访谈　□上机只读　□工具扫描'],
       ['2', '', '', '□访谈　□上机只读　□工具扫描'],
       ['3', '', '', '□访谈　□上机只读　□工具扫描'],
       ['4', '', '', '□访谈　□上机只读　□工具扫描']],
      [1.0, 4.2, 4.2, 5.2], hl=20)
md('**二、授权检查范围**'); md(); md('| 序号 | 受检对象 / 系统 | IP / 域名 / 网段 | 授权动作 |'); md('|---|---|---|---|')
for i in range(1, 5): md('| %d |  |  | □访谈 □上机只读 □工具扫描 |' % i)
md()

h('三、授权动作与现场纪律')
for i, x in enumerate([
    '**访谈与资料查阅**：访谈负责人、系统与网络管理员及员工代表，查看资产台账、管理制度、运行记录与报表；',
    '**上机只读核查**：登录主机、网络与安全设备、云平台控制台、容器与虚拟化平台、数据库，只读读取配置与状态；'
    '全程不改配置、不装软件、不重启服务；',
    '**工具测试**：对内网与外部暴露面开展端口与服务探测、组件指纹、目录与后台探测；'
    '口令核查仅在授权方书面确认账号锁定策略后执行；',
    '**证据留存**：对检查过程与结论截图、留存命令输出与扫描报告，仅用于本次检查与报告编制；',
    '**不实施**：SQL 注入、跨站脚本、文件上传、命令执行、越权利用等专项验证；扫描发现的疑似利用点仅记录现象并建议转入专项测试。'], 1):
    p = doc.add_paragraph(); setpara(p, Pt(23), 28)
    for seg, bold in ([(x, False)] if '**' not in x else
                      [(s0, False) for s0 in [x.replace('**', '')]]):
        setfont(p.add_run('%d.%s' % (i, seg)), SONG, FS_B, False)
    md('%d. %s' % (i, x))
md()

h('四、账号与口令')
for x in ['受托方按检查需要提出账号需求，授权方开通**临时只读 / 巡检级账号**，并在检查结束后立即停用或改密；',
          '个别不便开账号的系统，可由授权方管理员**现场输入口令**，受托方不留存、不记录明文；',
          '弱口令与默认口令的核查结果只记录是否命中与数量，不记录、不留存明文口令；',
          '受托方对接触到的账号、口令与业务数据承担保密义务，检查结束后按授权方要求销毁临时材料。']:
    p = doc.add_paragraph(); setpara(p, Pt(23), 28)
    setfont(p.add_run('· ' + x.replace('**', '')), SONG, FS_B)
    md('- ' + x)
md()

h('五、授权期限与时间窗口')
table(['项目', '内容'],
      [['授权有效期', '　　年　月　日　至　　年　月　日'],
       ['现场检查时间窗', '　　年　月　日　至　　年　月　日（避开高峰：＿＿＿）'],
       ['复测时间窗', '　　年　月　日　至　　年　月　日'],
       ['扫描授权范围', '□内网　□外网 / 公网资产　□不开展扫描']],
      [3.8, 10.8], hl=20)
md('**五、授权期限与时间窗口**'); md(); md('| 项目 | 内容 |'); md('|---|---|')
md('| 授权有效期 | 　　　　年　　月　　日 至 　　　　年　　月　　日 |')
md('| 现场检查时间窗 | 　　　　年　　月　　日 至 　　　　年　　月　　日 |')
md('| 复测时间窗 | 　　　　年　　月　　日 至 　　　　年　　月　　日 |')
md('| 扫描授权范围 | □内网 □外网 / 公网资产 □不开展扫描 |'); md()

h('六、成果与保密')
for x in ['检查成果为《信息安全检查报告》，仅向授权方交付，未经授权方同意不向第三方提供、不对外引用；',
          '报告中涉及授权方的系统信息、问题与证据按双方约定的密级管理；',
          '受托方参检人员签署保密承诺，检查结束后不保留授权方业务数据副本（授权方另行要求的除外）。']:
    p = doc.add_paragraph(); setpara(p, Pt(23), 28)
    setfont(p.add_run('· ' + x), SONG, FS_B)
    md('- ' + x)
md()

h('七、签署')
para('本授权书自双方签字盖章之日起生效。授权方保证其对授权范围内的受检对象具有合法的管理权或已取得相应授权；'
     '受托方在授权范围内开展检查所产生的结论，仅对检查时点的状态有效。', SONG, FS_B, False, 28)
md('本授权书自双方签字盖章之日起生效。授权方保证其对授权范围内的受检对象具有合法的管理权或已取得相应授权。'); md()
para('', SONG, FS_B, False, 0, None, Pt(18))
t = doc.add_table(rows=0, cols=2); t.autofit = False
for r in [['授权方（盖章）', '受托方（盖章）'],
          ['授权代表（签字）：', '项目负责人（签字）：'],
          ['日期：　　年　月　日', '日期：　　年　月　日']]:
    cells = t.add_row().cells
    for i, v in enumerate(r):
        p = cells[i].paragraphs[0]; setpara(p, Pt(30), 0, 4, 4)
        setfont(p.add_run(v), SONG, FS_TB)
for row in t.rows:
    for i, w in enumerate([7.3, 7.3]): row.cells[i].width = Cm(w)
md('| 授权方（盖章） |  | 受托方（盖章） |  |'); md('|---|---|---|---|')
md('| 授权代表（签字） |  | 项目负责人（签字） |  |')
md('| 日期 | 　　　　年　　月　　日 | 日期 | 　　　　年　　月　　日 |'); md()
para('附：受检资产清单（见《信息安全检查服务·现场调研表（进场前填写）》第二部分）、帐号开通清单（见该表第三部分）。',
     SONG, 12, False, 0, None, Pt(18), 6, 0)
md('> 附：受检资产清单（见调研表第二部分）、账号开通清单（见调研表第三部分）。')

os.makedirs(os.path.dirname(OUT), exist_ok=True); os.makedirs(os.path.dirname(MD), exist_ok=True)
doc.save(OUT); io.open(MD, 'w', encoding='utf-8').write('\n'.join(MDL) + '\n')
print('已生成：', OUT); print('已生成：', MD)
