# -*- coding: utf-8 -*-
"""生成《信息安全检查服务 · 现场调研表（进场前填写）》——客户提前填，用于压缩现场访谈。
   排版按单位《制度格式说明及示例》：页眉小五宋体带横线、标题方正小标宋简体三号、四级文本层级、表内小四/小五。"""
import io, os, re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

P   = os.environ.get('BRSCAN_ROOT', '.')   # 工作根目录：含交付物/与数据源，默认当前目录
OUT = P + '/2.0-精简核查体系/交付物/信息安全检查服务·现场调研表（进场前填写）.docx'
MD  = P + '/2.0-精简核查体系/交付物/文档介绍/信息安全检查服务·现场调研表（进场前填写）.md'

HEI, KAI, SONG, XBS = '黑体', '楷体', '宋体', '方正小标宋简体'
FS_H, FS_B, FS_TB, FS_S = 14, 14, 12, 9

# ───────── 内容定义（编号回指《现场检测清单》166 条，用于核对覆盖）─────────
INFO = [('受检单位全称', ''),
        ('受检系统 / 业务范围（本次检查覆盖哪些系统）', ''),
        ('系统部署方式（单选，决定检查范围）',
         '□公有云　□私有云 / 专有云　□自建机房（物理机 / 虚拟化）　□托管 IDC　□混合部署'),
        ('对接人（姓名 · 职务 · 电话 · 邮箱）', ''),
        ('拟检查时间窗口（可进现场的时间段）', '')]

ASSETS = [  # (类别, 对应细项)
 ('服务器（物理机 / 虚拟机）', '1.1-1、1.2-1、1.3-1'),
 ('终端 PC / 笔记本', '1.1-3、5.1-1'),
 ('网络设备（交换机 / 路由器 / 防火墙）', '3.1-1、3.7-1'),
 ('安全设备（IPS / WAF / 上网行为 / 堡垒机）', '3.10-1'),
 ('云主机与云资源（含对象存储、CDN、负载均衡）', '2.1-1'),
 ('数据库与中间件', '6.6-1、6.7-1'),
 ('应用系统（对外 / 对内）', '6.1-1'),
 ('存储与备份设备', '8.6-1、8.9-1'),
 ('无线（AP / 无线控制器）', '3.8-1'),
 ('机房与供配电（含 UPS、电池、配电）', '10.1-1、11.1-1'),
 ('对外域名与公网 IP（数量、用途）', '1.4-1、6.1-1'),
 ('公网开放服务 / 端口映射（数量）', '1.4-1、3.2-1')]

ACCTS = [  # (类别, 需要的权限, 说明)
 ('Linux 服务器', '只读（sudo 只读或普通账号 + 读权限）', '按资产组提供 1～2 个；口令由管理员现场输入或提供临时账号'),
 ('Windows 服务器 / 终端', '本地管理员组只读（或指定普通账号）', '如需域环境，另提供 1 个域只读账号'),
 ('网络与安全设备', '巡检 / 只读级别', '交换机、路由器、防火墙、IPS/WAF 各 1 个；只读级别即可'),
 ('云平台（阿里云 / 华为云 / 腾讯云 / AWS）', '只读子账号（RAM / IAM）', '需能查看资产、安全组、存储桶策略、操作审计；不开写权限'),
 ('Kubernetes / 虚拟化平台', '只读（只读角色 / 只读集群角色）', 'K8s 需 api 只读；虚拟化平台需只读或审计员角色'),
 ('数据库', '只读账号（每类引擎 1 个）', 'MySQL / Redis / PostgreSQL / Oracle / SQL Server 按实际提供'),
 ('堡垒机 / 日志审计平台', '只读', '用于核对登录审计与日志留存，无需操作权限'),
 ('终端管理 / 杀毒 EDR 控制台', '只读', '用于核对终端覆盖率与病毒库更新时间'),
 ('备份系统 / 存储', '只读', '用于核对备份策略、最近成功备份与介质状态')]

DOCS = [  # (资料名称, 对应细项)
 ('资产台账（IP、责任人、用途、重要级别、上下线记录）', '1.1-1、1.1-2、1.2-1'),
 ('网络拓扑图与网段划分说明', '1.4-1'),
 ('公网开放服务清单 / 端口映射表（含审批记录）', '3.2-1、1.4-1'),
 ('应用系统清单（域名、端口、责任人、数据级别）', '6.1-1、6.1-2'),
 ('云资产清单与云服务合同 / 责任边界约定', '2.1-1、2.1-2、2.1-3、9.8-1'),
 ('账号台账 / 权限矩阵与权限复核记录', '8.2-1、8.2-2、9.2-1'),
 ('安全管理制度六类（账号、变更、应急、备份、人员、第三方）与应急预案', '9.1-1、9.3-1、9.2-2'),
 ('日志留存与集中审计说明（留存天数、外发目标）', '7.1-1、7.4-1'),
 ('备份策略、恢复演练记录、备份介质台账', '8.6-1、8.7-2、8.9-1'),
 ('重要数据与个人信息清单（含分类分级结果、外发审批记录）', '8.3-1、8.4-1、8.5-1'),
 ('机房与供配电资料（门禁记录、UPS 维保与检测报告、线缆图）——纯云上可不提供', '10.1-1、11.1-1、11.2-1'),
 ('近一年安全事件台账、上次检查 / 等保测评报告与整改记录', '9.4-1、9.4-2、9.5-1')]

FACTS = [  # (需确认事项, 对应细项)
 ('资产台账最近一次更新时间，现场是否与台账一致（不致的台数）', '1.1-2、1.1-4'),
 ('是否存在第三方托管或远程运维；托管范围、接口人与维护通道', '1.2-3、9.7-1'),
 ('云账号是否分权、关键操作是否开 MFA；AK/SK 几套、谁持有', '2.2-1、2.2-2、2.2-3'),
 ('云安全组 / 网络 ACL 是否存在 0.0.0.0/0 放通高危端口', '2.3-2、2.3-3'),
 ('无线加密方式（WPA2 / WPA3），访客网与办公网是否隔离', '3.8-1'),
 ('安全产品是否在线且为阻断模式；近半年告警是否闭环', '3.10-2'),
 ('终端是否统一部署杀毒 / EDR；移动介质是否管控（有无白名单）', '5.3-2、5.5-1'),
 ('权限是否定期复核、离职转岗是否回收（最近一次复核时间）', '8.2-2'),
 ('是否收集个人信息、是否完成重要数据分类分级', '8.3-1、8.5-1、8.5-2'),
 ('是否存在离线 / 不可变备份；最近一次恢复演练时间与结论', '8.7-1、8.8-1'),
 ('六类制度是否已发布；近一年是否自查 / 内审、是否做过演练', '9.1-1、9.5-1、9.6-2'),
 ('近一年是否发生安全事件（件数）、是否完成整改闭环', '9.4-1、9.4-2')]

SITE = [  # (事项, 说明)
 ('可进机房查看设备、线缆、UPS 与配电（纯云上填「不适用」）', '含门禁登记与陪同人员安排'),
 ('可提供跳板机 / 临时账号开通窗口', '建议约定统一开通与回收时间'),
 ('内网扫描授权与时间窗', '含扫描网段、避开业务高峰时段'),
 ('外网 / 公网资产扫描授权', '含域名与公网 IP 清单确认'),
 ('口令核查授权（是否允许、账号锁定策略）', '需书面同意；命中只记数量不留明文'),
 ('陪同人员（系统 / 网络 / 机房管理员）', '各 1 名，能回答配置与流程问题')]

def nums(rows, idx):
    out = []
    for r in rows:
        out += re.findall(r'\d+\.\d+-\d+', r[idx])
    return out

# 从清单源数据取「限时免费」条目编号集合（F = 访谈/资料/目视）。
# 数据源为清单 content.py（DATA 结构），可用环境变量 BRSCAN_CONTENT 指定；
# 缺失时跳过覆盖率统计（COVER=0），不影响调研表生成。
import importlib.util
F_SET = set()
_content_path = os.environ.get('BRSCAN_CONTENT', P + '/checklist-slim/content.py')
if os.path.exists(_content_path):
    _spec = importlib.util.spec_from_file_location('content', _content_path)
    _C = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(_C)
    for _d, _dt, _cis in _C.DATA:
        for _cno, _ct, _items in _cis:
            for _k, (_lv, _tx) in enumerate(_items, 1):
                if _lv == 'F':
                    F_SET.add('%s-%d' % (_cno.split(' ')[0], _k))
else:
    print('提示: 未找到清单数据源 content.py（BRSCAN_CONTENT 可指定），跳过覆盖率统计')

cited = set(nums(ASSETS, 1) + nums(DOCS, 1) + nums(FACTS, 1))
COVER = len(cited & F_SET)          # 命中限时免费条数
COVER_T = len(cited - F_SET)        # 引用到的技术类条目（作账号/前置依据）
F_TOTAL = len(F_SET)

# ───────── 排版基建 ─────────
doc = Document()
st = doc.styles['Normal']; st.font.name = SONG; st.font.size = Pt(FS_B)
st.element.rPr.rFonts.set(qn('w:eastAsia'), SONG)

def setfont(run, name, size, bold=False):
    run.font.name = name; run.font.size = Pt(size); run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)

def setpara(p, line=Pt(23), indent=0, before=0, after=0, align=None):
    pf = p.paragraph_format
    if line: pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY; pf.line_spacing = line
    if indent: pf.first_line_indent = Pt(indent)
    pf.space_before, pf.space_after = Pt(before), Pt(after)
    if align is not None: p.alignment = align
    return p

def para(text, name=SONG, size=FS_B, bold=False, indent=0, align=None, line=Pt(23), before=0, after=0):
    p = doc.add_paragraph(); setpara(p, line, indent, before, after, align)
    setfont(p.add_run(text), name, size, bold); return p

def h2(text):
    return para(text, HEI, FS_H, False, 0, WD_ALIGN_PARAGRAPH.LEFT, Pt(23), 8, 4)

def table(cols, rows, widths, size=FS_S, head_size=FS_S, h=15, first_col_center=True):
    t = doc.add_table(rows=1, cols=len(cols)); t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
    lay = OxmlElement('w:tblLayout'); lay.set(qn('w:type'), 'fixed'); t._tbl.tblPr.append(lay)
    for i, c in enumerate(cols):
        cell = t.rows[0].cells[i]; cell.text = ''
        p = cell.paragraphs[0]; setpara(p, Pt(h), 0, 2, 2, WD_ALIGN_PARAGRAPH.CENTER)
        setfont(p.add_run(c), HEI, head_size)
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r):
            p = cells[i].paragraphs[0]
            setpara(p, Pt(h), 0, 2, 2, WD_ALIGN_PARAGRAPH.CENTER if (i == 0 and first_col_center) else None)
            setfont(p.add_run(str(v)), SONG, size)
    for row in t.rows:
        for i, w in enumerate(widths): row.cells[i].width = Cm(w)
    for i, w in enumerate(widths): t.columns[i].width = Cm(w)
    return t

# ───────── 页面与页眉页脚 ─────────
s = doc.sections[0]
s.page_width, s.page_height = Cm(21), Cm(29.7)
s.top_margin = s.bottom_margin = Cm(2.2); s.left_margin = s.right_margin = Cm(3.17)
s.header_distance, s.footer_distance = Cm(1.5), Cm(1.75)
# 页眉：无边框两列表格（左单位名 / 右标题顶右），横线用表格下边框（WPS/Word 一致）
def _el(tag, **a):
    e = OxmlElement(tag)
    for k, v in a.items(): e.set(qn('w:' + k), str(v))
    return e
def _run9(text):
    r = _el('w:r'); rPr = _el('w:rPr')
    rPr.append(_el('w:rFonts', ascii='宋体', hAnsi='宋体', eastAsia='宋体'))
    rPr.append(_el('w:sz', val=18)); rPr.append(_el('w:szCs', val=18))
    r.append(rPr); t = _el('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text; r.append(t)
    return r
def _cell(text, w, align):
    tc = _el('w:tc'); tcPr = _el('w:tcPr')
    tcPr.append(_el('w:tcW', w=w, type='dxa')); tcPr.append(_el('w:vAlign', val='bottom')); tc.append(tcPr)
    pp = _el('w:p'); pPr = _el('w:pPr')
    pPr.append(_el('w:spacing', lineRule='exact', line=240, before=0, after=0))
    if align: pPr.append(_el('w:jc', val=align))
    pp.append(pPr); pp.append(_run9(text)); tc.append(pp)
    return tc

_tw = int((s.page_width - s.left_margin - s.right_margin) / 635)
_lw = 1700
_root = s.header.part.element
for _c in list(_root): _root.remove(_c)
_tbl = _el('w:tbl'); _pr = _el('w:tblPr')
_pr.append(_el('w:tblW', w=_tw, type='dxa')); _pr.append(_el('w:tblLayout', type='fixed'))
_b = _el('w:tblBorders')
for _side in ('top', 'left', 'right', 'insideH', 'insideV'):
    _b.append(_el('w:' + _side, val='none', sz=0, space=0, color='auto'))
_b.append(_el('w:bottom', val='single', sz=6, space=1, color='000000')); _pr.append(_b)
_mar = _el('w:tblCellMar')
for _side in ('top', 'left', 'bottom', 'right'): _mar.append(_el('w:' + _side, w=0, type='dxa'))
_pr.append(_mar); _tbl.append(_pr)
_g = _el('w:tblGrid'); _g.append(_el('w:gridCol', w=_lw)); _g.append(_el('w:gridCol', w=_tw - _lw)); _tbl.append(_g)
_tr = _el('w:tr')
_tr.append(_cell('受检单位', _lw, None))
_tr.append(_cell('信息安全检查服务·现场调研表（进场前填写）', _tw - _lw, 'right'))
_tbl.append(_tr); _root.append(_tbl)
_tail = _el('w:p'); _tp = _el('w:pPr'); _tp.append(_el('w:spacing', lineRule='exact', line=20, before=0, after=0))
_tail.append(_tp); _root.append(_tail)

fp = s.footer.paragraphs[0]; setpara(fp, Pt(12), 0, 0, 0, WD_ALIGN_PARAGRAPH.CENTER)
fl = OxmlElement('w:fldSimple'); fl.set(qn('w:instr'), ' PAGE ')
rr = OxmlElement('w:r'); rpr = OxmlElement('w:rPr')
rf = OxmlElement('w:rFonts'); rf.set(qn('w:eastAsia'), SONG); rpr.append(rf)
sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '18'); rpr.append(sz)
rr.append(rpr); fl.append(rr); fp._p.append(fl)

MDL = []
def md(x=''): MDL.append(x)

# ───────── 标题与填写说明 ─────────
para('信息安全检查 · 现场调研表（进场前填写）', XBS, 16, True, 0, WD_ALIGN_PARAGRAPH.CENTER, Pt(28), 0, 10)
md('# 信息安全检查 · 现场调研表（进场前填写）'); md()

h2('填写说明')
for i, x in enumerate([
    '本表由受检单位在检查前 **3 个工作日**填写并回复，我方据此预填核查表、准备账号与工具，现场访谈可压缩到 60 分钟以内。',
    '共 6 部分、%d 个填写位，其中大部分为勾选或填数量；**不清楚的写「待确认」**，不必勉强作答。' % (len(INFO) + len(ASSETS) + len(ACCTS) + len(DOCS) + len(FACTS) + len(SITE)),
    '数量填大致量级即可（如「约 20 台」）；IP / 域名只需填主要网段与对外地址，无需逐台列全；与责任人以资产台账为准，本表不必重复填。',
    '部署方式直接决定检查范围：**纯云上 / 托管**时，物理环境（第十域）与供电（第十一域）类检查项改由云服务商侧资料替代并标注「不适用」，云与虚拟化（第二域）为重点；**自建机房**则相反。混合部署请在第 2 部分逐类注明云上 / 自建。',
    '账号类以**临时只读账号**为主（推荐，检查结束即停用或改密）；个别不便开账号的系统，可由管理员现场代输口令，仅限少量。',
    '覆盖口径：本表采集的信息对应《信息安全检查服务现场检测清单》限时免费类 %d 条中的 **%d 条**，'
    '其余 %d 条仍需现场访谈或目视确认；另有 %d 处引用为技术类条目，用于说明账号与前置条件（表 2、表 4、表 5 末列标注编号，便于核对）。'
    % (F_TOTAL, COVER, F_TOTAL - COVER, COVER_T),
    '本表信息仅用于本次检查，我方按保密要求管理，不外传、不用于其他用途。'], 1):
    p = doc.add_paragraph(); setpara(p, Pt(23), 28)
    setfont(p.add_run('%d.%s' % (i, x.replace('**', ''))), SONG, FS_B)
    md('%d. %s' % (i, x))
md()

# ───────── 表 1 基本信息 ─────────
h2('一、基本信息（%d 项）' % len(INFO))
md('## 一、基本信息（%d 项）' % len(INFO)); md()
t1 = table(['序号', '项目', '填写内容'], [[str(i), k, v] for i, (k, v) in enumerate(INFO, 1)],
           [1.0, 5.6, 8.0], size=FS_TB, head_size=FS_TB, h=20)
md('| 序号 | 项目 | 填写内容 |'); md('|---|---|---|')
for i, (k, v) in enumerate(INFO, 1): md('| %d | %s |  |' % (i, k))
md()

# ───────── 表 2 资产摸底 ─────────
h2('二、资产摸底（%d 类）' % len(ASSETS))
md('## 二、资产摸底（%d 类）' % len(ASSETS)); md()
rows = [[str(i), k, '', '', '□云上　□自建', c] for i, (k, c) in enumerate(ASSETS, 1)]
table(['序号', '资产类别', '数量', '主要 IP / 网段 / 域名', '部署', '对应细项'],
      rows, [0.9, 4.0, 1.3, 4.8, 1.9, 2.1])
md('| 序号 | 资产类别 | 数量 | 主要 IP / 网段 / 域名 | 部署 | 对应细项 |')
md('|---|---|---|---|---|---|')
for i, (k, c) in enumerate(ASSETS, 1): md('| %d | %s |  |  | □云上 □自建 | %s |' % (i, k, c))
md()

# ───────── 表 3 账号与访问方式 ─────────
h2('三、账号与访问方式（%d 类）' % len(ACCTS))
md('## 三、账号与访问方式（%d 类）' % len(ACCTS)); md()
para('用于上机核查与工具测试，仅需只读 / 巡检级权限。推荐开通临时只读账号（检查结束即停用或改密）；'
     '个别不便开账号的系统，可由管理员现场代输口令。', SONG, FS_S, False, 0, None, Pt(18), 2, 2)
md('> 仅需只读 / 巡检级权限；检查结束后由受检单位收回。'); md()
rows = [[str(i), k, q, '□临时只读账号（推荐）　□跳板机接入　□管理员现场代输（仅少量系统）', '□已约定', d]
        for i, (k, q, d) in enumerate(ACCTS, 1)]
table(['序号', '系统 / 设备类别', '需要的权限级别', '开通方式', '权限收回', '备注'],
      rows, [0.9, 2.6, 2.4, 3.4, 1.2, 3.7])
md('| 序号 | 系统 / 设备类别 | 需要的权限级别 | 开通方式 | 权限收回 | 备注 |')
md('|---|---|---|---|---|---|')
for i, (k, q, d) in enumerate(ACCTS, 1):
    md('| %d | %s | %s | □临时只读账号（推荐）/ □跳板机 / □管理员现场代输 | □已约定 | %s |' % (i, k, q, d))
md()

# ───────── 表 4 需提供资料 ─────────
h2('四、需提前提供的资料（%d 项）' % len(DOCS))
md('## 四、需提前提供的资料（%d 项）' % len(DOCS)); md()
rows = [[str(i), k, '□可提供　□部分提供　□不能提供', '', c] for i, (k, c) in enumerate(DOCS, 1)]
table(['序号', '资料名称', '能否提供', '提供方式 / 备注', '对应细项'],
      rows, [1.0, 5.7, 2.7, 3.2, 2.0])
md('| 序号 | 资料名称 | 能否提供 | 提供方式 / 备注 | 对应细项 |')
md('|---|---|---|---|---|')
for i, (k, c) in enumerate(DOCS, 1): md('| %d | %s | □可提供 □部分提供 □不能提供 |  | %s |' % (i, k, c))
md()

# ───────── 表 5 关键情况确认 ─────────
h2('五、关键情况确认（%d 项）' % len(FACTS))
md('## 五、关键情况确认（%d 项）' % len(FACTS)); md()
para('勾选现状即可，「说明」列填数量、时间或简要情况；这些事项现场不再重复询问。', SONG, FS_S, False, 0, None, Pt(18), 2, 2)
md('> 勾选现状即可，「说明」列填数量、时间或简要情况；这些事项现场不再重复询问。'); md()
rows = [[str(i), k, '□是　□否　□部分', '', c] for i, (k, c) in enumerate(FACTS, 1)]
table(['序号', '需确认事项', '现状', '说明（数量 / 时间 / 简要情况）', '对应细项'],
      rows, [1.0, 5.2, 2.4, 4.0, 2.0])
md('| 序号 | 需确认事项 | 现状 | 说明 | 对应细项 |'); md('|---|---|---|---|---|')
for i, (k, c) in enumerate(FACTS, 1): md('| %d | %s | □是 □否 □部分 |  | %s |' % (i, k, c))
md()

# ───────── 表 6 现场条件 ─────────
h2('六、现场条件确认（%d 项）' % len(SITE))
md('## 六、现场条件确认（%d 项）' % len(SITE)); md()
rows = [[str(i), k, '□可　□不可', d] for i, (k, d) in enumerate(SITE, 1)]
table(['序号', '事项', '确认', '备注'], rows, [1.0, 6.3, 2.4, 4.9])
md('| 序号 | 事项 | 确认 | 备注 |'); md('|---|---|---|---|')
for i, (k, d) in enumerate(SITE, 1): md('| %d | %s | □可 □不可 | %s |' % (i, k, d))
md()

# ───────── 落款 ─────────
para('', SONG, FS_B, False, 0, None, Pt(18))
rows = [['受检单位（盖章）', '', '我方接收人', ''],
        ['填写人 / 联系方式', '', '接收日期', '　　　　年　　月　　日']]
t = doc.add_table(rows=0, cols=4); t.autofit = False
for r in rows:
    cells = t.add_row().cells
    for i, v in enumerate(r):
        p = cells[i].paragraphs[0]; setpara(p, Pt(22), 0, 2, 2)
        setfont(p.add_run(v), SONG, FS_TB)
for row in t.rows:
    for i, w in enumerate([3.4, 3.6, 2.6, 4.8]): row.cells[i].width = Cm(w)
md('| 受检单位（盖章） |  | 我方接收人 |  |'); md('|---|---|---|---|')
md('| 填写人 / 联系方式 |  | 接收日期 | 　　　　年　　月　　日 |'); md()
para('注：本表随《现场核查表（访谈部分 / 技术部分）》配合使用；如部分信息涉密不便书面提供，可现场口头说明并记录。',
     SONG, FS_S, False, 0, None, Pt(16), 8, 0)
md('> 注：本表随《现场核查表（访谈部分 / 技术部分）》配合使用；如部分信息涉密不便书面提供，可现场口头说明并记录。')

os.makedirs(os.path.dirname(MD), exist_ok=True)
doc.save(OUT)
io.open(MD, 'w', encoding='utf-8').write('\n'.join(MDL) + '\n')
slots = len(INFO) + len(ASSETS) + len(ACCTS) + len(DOCS) + len(FACTS) + len(SITE)
print('已生成：', OUT)
print('已生成：', MD)
print('填写位合计 %d；引用编号 %d 个（限时免费 %d / 技术类 %d）' % (slots, len(cited), COVER, COVER_T))
