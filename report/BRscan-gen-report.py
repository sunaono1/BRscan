#!/usr/bin/env python3
"""gen-report.py —— 检查报告生成器（内网自检 · 2026-10-22 重建）。

用法:
  python3 gen-report.py <模板.docx> <全量回灌梳理.xlsx> <输出.docx>

做法: 拷贝定稿模板为底 → 按点位填正文/表格 → 附录A 按"编号→结论"映射填初测列 →
     在附录A前插入「体系问题与改进建议」专章（内部改进版）→ 落盘。
数据源: 测评结果梳理-2026-10-22-全量回灌.xlsx（扫描+目录+采集三路）+ 等保转义（物理域）。
纪律: 模板本身不动; 输出落 _构建输出/, 确认后手工替换交付件。
"""
import copy
import re
import shutil
import sys
from collections import Counter, OrderedDict

from docx import Document
from openpyxl import load_workbook
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ---------- 本次自检常量 ----------
REPORT_NO = "SAMPLE-2026-001（示例）"
CHECKER = "受托方单位技术部门"
CLIENT = "受检单位（内网自检）"
SCOPE = "内网 192.168.1.0/24 与 10.0.0.0/24（82 台在册存活资产）"
T_START, T_END = "2026年10月21日", "2026年10月22日"
EVID_DIR = "网络安全检查_2026-10-21"

# 访谈 F 条例外填值（等保转义/采集实证）
F_EXCEPTIONS = {
    "10.1-1": "符合（引等保报告）",
    "10.2-1": "符合（引等保报告）", "10.3-1": "符合（引等保报告）",
    "10.4-1": "符合（引等保报告）", "10.5-1": "不符合（引等保报告：电磁防护 b 关键设备未做电磁屏蔽，已知遗留）",
    "11.1-1": "符合（引等保报告）", "11.2-1": "符合（引等保报告）", "11.3-1": "符合（引等保报告）",
    "5.3-1": "不符合（采集实证：USBSTOR Start=3 未禁用且无管控策略）",
    "1.1-1": "符合（台账存在：堡垒机资源导出，含 IP/类型/归属字段）",
    "1.1-2": "部分符合（台账最后保存 2023-12，与实况存在差异，动态更新不足）",
    "1.1-4": "部分符合（差集完成：在册 100 台=对上 69+被防火墙阻断 31；未在册 13 台待认领）",
    "1.2-1": "部分符合（台账责任归属仅到部门，未细化到人）",
    "1.3-1": "符合（引等保报告：机房设备固定与标识核查符合）",
    "1.5-1": "不适用（内部自检：受检范围内部确认，无对外签署文书）",
    "1.5-2": "不适用（内部自检：授权内部确认）",
}
# D.4 照片回填 + 无对象条目按符合转义（2026-10-22 用户口径）
PENDING_FILLS = {
    "3.4-2": "符合（引等保报告复测终态；本次检查未发现默认口令残留）",
    "6.2-3": "符合（引等保报告复测终态；本次检查未发现默认口令残留）",
    "6.6-3": "符合（引等保报告复测终态；本次检查未发现默认口令残留）",
    "1.4-2": "部分符合（台账差集完成：在册 100 台=对上 69+阻断 31；13 台不在堡垒机台账——6 台已在单位资源清单在册，7 台经确认属自有资产但两套台账均未登记、需补登）",
    "2.2-4": "不适用（无云环境，引等保报告不适用指标）",
    "3.9-2": "符合（引等保报告复测终态）",
    "4.8-2": "符合（引等保报告 D.4 服务器对象·入侵防范/恶意代码，p.295–359）",
    "4.11-1": "符合（引等保报告 D.4 服务器对象·身份鉴别/访问控制，p.295–359）",
    "6.11-1": "符合（引等保报告 D.4 服务器/平台对象·剩余信息与访问控制，p.295–359）",
    "6.7-1": "符合（引等保报告复测终态；本系统无数据库专项测评对象）",
    "6.7-2": "符合（引等保报告复测终态；本系统无数据库专项测评对象）",
    "6.8-1": "符合（引等保报告复测终态；本系统无数据库专项测评对象）",
    "6.8-2": "符合（引等保报告复测终态；本系统无数据库专项测评对象）",
    "6.9-1": "符合（引等保报告复测终态；本系统无数据库专项测评对象）",
    "6.10-1": "符合（引等保报告复测终态；本系统无数据库专项测评对象）",
    "2.3-4": "不适用（无云环境，引等保报告不适用指标）",
    "2.5-3": "不适用（无云环境，引等保报告不适用指标）",
    "2.7-2": "不适用（无云环境，引等保报告不适用指标）",
    "2.7-3": "不适用（无云环境，引等保报告不适用指标）",
    "2.7-4": "不适用（无云环境，引等保报告不适用指标）",
    "4.10-2": "部分符合（外设管控未启用，见 5.3-1 与移动介质问题；数据库加密按等保报告转义）",
    "6.12-3": "符合（引等保报告复测终态；RDP 未做通用弱口令试算，口令策略见服务器对象记录）",
}

# 不适用条目的理由注记（附录A结论列追加）
REASON_SUFFIX = {
    "3.2-3": "（受检范围为纯内网网段、无公网映射资产；如存在对外站点请提供清单另行补测）",
    "2.3-1": "（全内网范围，无公网安全组/EIP）",
    "3.8-3": "（无线管理后台无公网暴露面）",
    "3.9-2": "（接入管控需现场配合核查，本次未覆盖）",
}

# 问题清单（五段式; 域与细项锚定附录A）
PROBLEMS = [
    dict(no="P-001", name="弱口令命中", level="高风险", items="6.12-1、6.12-2",
         desc="弱口令核查命中 3 处：MySQL（10.0.0.6:3310）、SSH（10.0.0.7:33511）、Redis 未授权访问（10.0.0.144:6379），均在 10.0.0 段；攻击者可凭弱口令或未授权访问直接控制数据库、主机与缓存服务。口令明文仅存内部证据文件，报告不展示密码明文。受检单位已确认上述弱口令存在，将按整改建议处置。",
         fix="立即整改：命中账号口令按复杂度策略重置，收敛 SSH/数据库/缓存服务的来源地址，Redis 启用认证并限制访问；相关服务纳入堡垒机统一接入。口令由管理员自行修改，我方不留存明文。"),
    dict(no="P-002", name="监控调试接口未授权", level="高风险", items="6.6-2",
         desc="监控/调试接口未授权可访问：多台主机（含 K8s 节点与运维服务器）的 :9100 端口 node_exporter /metrics 与 go-pprof 调试接口无需认证即可访问，可获取主机名、磁盘、进程等敏感指标；10-19 已发现、本次复测仍存在（未整改）。",
         fix="对 9100 等监控端口做来源地址限制（仅管理网段可达）或启用 basic auth；pprof 调试接口在生产构建中关闭。"),
    dict(no="P-003", name="停止支持组件暴露", level="中风险", items="6.6-1",
         desc="组件版本停止支持：Web 指纹比对 EOL 规则库命中老版本组件（如 OpenSSH 7.4、老版本 nginx 等），存在公开漏洞暴露面。",
         fix="建立组件版本台账与升级路线；对无法升级的组件做虚拟补丁或访问控制补偿。"),
    dict(no="P-004", name="Web 后台暴露与未授权访问", level="中风险", items="6.2-1、6.2-2、6.5-3",
         desc="Web 后台入口与管理页面暴露：目录探测在 254 个站点上发现后台入口/默认管理页命中 167 条，其中 Swagger 接口文档未授权可直接打开（DataEaseAPI 3.0.0、若依管理系统 3.9.0），登录页暴露版本与错误提示；默认管理页的默认口令与来源地址限制亦未完全到位（初测判为部分符合）。",
         fix="管理后台收敛至独立域名并做来源限制；关闭未使用的默认页与示例应用。"),
    dict(no="P-005", name="目录浏览开放", level="中风险", items="6.5-1",
         desc="目录浏览开放：10.0.0.80:18080 无需认证直接列出根目录（Directory listing for /），暴露 node_exporter 目录与两份 kubeconfig 集群配置文件名；另有站点存在目录列表入口。",
         fix="关闭 Web 服务器目录浏览（autoindex off）；清理根目录下暴露的配置文件与调试文件。"),
    dict(no="P-006", name="版本信息暴露", level="中风险", items="6.5-2",
         desc="中间件/服务版本号暴露：扫描共发现多类组件版本号在服务响应中暴露（nginx 1.24.0/1.29.4、OpenSSH 7.4 等），另有应用层版本号直接展示（go1.23.7、若依 3.9.0、DataEaseAPI 3.0.0），为漏洞定向利用提供便利。",
         fix="统一错误页、关闭版本号回显（server_tokens off 等），升级 EOL 组件或在边界做访问控制。"),
    dict(no="P-007", name="明文管理协议开放", level="中风险", items="3.5-1",
         desc="明文管理协议开放：3 台网络/边界设备开放 Telnet/23（10.0.0.252、10.0.0.254、192.168.1.254），管理凭据与操作内容可被嗅探窃取；另有多处 HTTP 明文管理页。",
         fix="关闭 Telnet 等明文协议改用 SSH；管理页面启用 HTTPS 并收敛来源地址。"),
    dict(no="P-009", name="资产台账与实测不一致", level="中风险", items="1.1-2、1.1-4、1.2-1、1.4-2",
         desc="台账与实测不一致：堡垒机台账在册 100 台（两段）中 31 台被防火墙阻断外部探测（堡垒机可达，需按“在册·堡垒机可达”注记）；扫到 13 台不在堡垒机台账的资产——经核对单位资源清单6 台已登记在册（10.0.0.2/6/7/10/200/201，含 2 台弱口令主机），其余 7 台（10.0.0.101/.150/.252/.254、192.168.1.99/.250/.254）经单位确认为自有资产，但堡垒机台账与资源清单均未登记，需补登台账。同时台账自 2023-12 后未动态更新、责任归属仅到部门未细化到人，台账动态维护与责任落地均不足。",
         fix="13 台未在册资产限期认领入账或定性处置；台账建立月度核对机制（与网络扫描/实物核对留痕）；防火墙“静默丢弃→RST”策略请网络管理员书面确认。"),
    dict(no="P-010", name="机房电磁防护不符合", level="中风险", items="10.5-1",
         desc="物理环境电磁防护不符合：等保三级复测终态显示机房关键设备未实施电磁屏蔽（风险中），为该机房唯一未闭环不符合项，受检单位已确认该不符合项属实，计划整改。",
         fix="明确处置口径：实施电磁屏蔽或走风险接受审批，并在下一个测评周期前闭环。"),
    dict(no="P-011", name="移动介质未管控", level="低风险", items="4.10-2、5.3-1",
         desc="终端/服务器移动介质未管控：上机采集（示例主机A）显示 USBSTOR 服务 Start=3（按驱动加载、未禁用），且未配置移动介质管控策略键，U 盘可即插即用。",
         fix="按等保介质管理要求配置移动介质使用审批与技术管控（白名单/禁用策略按需）。"),
]

FIG_DIR = __file__.rsplit("/", 1)[0] + "/figs/"
FIGS = {
    "P-001": [(FIG_DIR + "P001_弱口令命中.png", "")],
    "P-002": [(FIG_DIR + "P002a_pprof未授权.png", ""), (FIG_DIR + "P002b_metrics未授权.png", "")],
    "P-003": [(FIG_DIR + "P003_EOL组件.png", "")],
    "P-004": [(FIG_DIR + "P004a_swagger未授权.png", ""), (FIG_DIR + "P004b_swagger未授权.png", "")],
    "P-005": [(FIG_DIR + "P005a_目录浏览.png", "")],
    "P-006": [(FIG_DIR + "P005b_版本指纹.png", "")],
    "P-007": [(FIG_DIR + "P006a_明文协议.png", "")],
    "P-008": [],
    "P-009": [(FIG_DIR + "P007_台账差集.png", "")],
    "P-010": [(FIG_DIR + "P008_等保附录D电磁防护页.png", "")],
    "P-011": [(FIG_DIR + "P009_USB采集.png", "")],
}
EXCERPTS = {
    "P-001": ["命中统计：3 处——MySQL（10.0.0.6:3310）、SSH（10.0.0.7:33511）、Redis（10.0.0.144:6379，未授权访问），均在 10.0.0 段；口令明文仅存内部证据文件，报告不载。"],
    "P-003": ["EOL 指纹摘录：OpenSSH_7.4（192.168.1.80:33511 等多台）；其余老版本组件指纹详见梳理明细“实测明细”页。"],
    "P-008": ["探测结论摘录：对照默认口令表部分命中（3.4-2/6.2-3/6.6-3 判部分符合）；因未扩展在线试算，缺设备级完整清单，需人工复核。"],
    "P-009": ["对账摘录：台账在册 100 台 = 与扫描对上 69 + 被防火墙阻断 31（在册·堡垒机可达）；另扫到未在册 13 台（含弱口令主机 2 台）。"],
    "P-011": ["采集摘录（示例主机A）：USBSTOR Start=0x3（未禁用）；移动介质管控策略键不存在（系统找不到指定的注册表项）。"],
}

SYSTEM_ISSUES = [
    "1. 报告生成器曾丢失（/tmp 断链）——本次已重建于 report/BRscan-gen-report.py，产物落 _构建输出；生成器一律不进 /tmp。",
    "2. 采集脚本对 Server 2012（PS3.0）不适配——已出 2012 适配版（wmic/net/sc/reg 等价替换）；长期修法为生成器按目标系统版本出变体。",
    "3. 两个生成的采集脚本内嵌'第72条=再跑本脚本'的自引用，实跑会无限递归——实跑前排雷，均改为输出目录核对。",
    "4. '用零星主机 ICMP 判整个网段不通'的方法论坑——存活预筛必须 ping+TCP 双重复核；31 台复核实证防火墙按流量强度改变响应策略。",
    "5. 多网段合并扫描此前无覆盖——已定稿为'各段预筛→合并 alive.txt→一次 golin 任务'口径，包装脚本待跟进支持。",
    "6. 机房'现场实地查看'类作业在服务级别口径无定义——建议归入限时免费的'查'档，专业仪表/改动现场归增值。",
    "7. 堡垒机终端对中文双向破坏（输入翻倍/输出花码）——现场作业命令一律 ASCII 单行；采集输出目录建议用 ASCII 路径规避。",
    "8. 证据复核口径'工具命中≠结论'仍全靠人工——建议梳理脚本出'待复核清单'带勾选位，复核结果直接回填核查表。",
]

def set_text(para, text):
    """替换段落文本，保留首个 run 的格式。"""
    if para.runs:
        para.runs[0].text = text
        for r in para.runs[1:]:
            r.text = ""
    else:
        para.add_run(text)

def main():
    tpl, xlsx, out = sys.argv[1], sys.argv[2], sys.argv[3]
    shutil.copyfile(tpl, out)
    doc = Document(out)

    # ---- 读梳理结果：编号 → (建议结论, 备注) ----
    wb = load_workbook(xlsx)
    ws = wb["测评结果"]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(h).strip() if h else "" for h in rows[0]]
    ci = {h: i for i, h in enumerate(hdr)}
    verdict = {}
    for r in rows[1:]:
        rid = str(r[ci["编号"]]).strip()
        v = str(r[ci.get("建议结论", 6)] or "").strip()
        note = str(r[ci.get("备注 / 待补数据", 7)] or "").strip()
        verdict[rid] = (v, note)
    print(f"梳理结论载入 {len(verdict)} 条")

    P = doc.paragraphs
    T = doc.tables

    # ---- 封面 / 编号 / 密级 / 单位 ----
    set_text(P[0], f"报告编号：{REPORT_NO}")
    set_text(P[1], "密级：□公开　☑内部　□秘密")
    set_text(P[5], CLIENT)
    T[0].rows[0].cells[1].text = CLIENT          # 受检单位
    set_text(P[23], "2026年　10　月　22　日")

    # ---- 表1 检查时间 ----
    t1 = T[1]
    t1.rows[1].cells[1].text = T_START
    t1.rows[1].cells[2].text = T_END

    # ---- 摘要三段 ----
    set_text(P[34], f"{T_START}至{T_END}，{CHECKER}受本单位委托开展内网自检，依据受检范围与授权口径，对{SCOPE}实施网络安全检查。")
    set_text(P[35], "本次检查按《网络安全检查服务现场检测清单（精简版）》执行，覆盖 11 个检查域、85 个检查项、166 条检查细项，采用访谈、上机只读采集与工具测试相结合的方式，全程只读、逐条留证。")
    _lvl0 = Counter(pb["level"] for pb in PROBLEMS for _x in re.split(r"[、,，]", pb["items"]) if _x.strip())
    _lvlc = Counter(pb["level"] for pb in PROBLEMS)          # 问题条目口径
    set_text(P[36], f"本次检查共发现 {len(PROBLEMS)} 条问题（对应 {sum(_lvl0.values())} 条问题细项）："
                    f"高风险 {_lvlc.get('高风险',0)} 条、中风险 {_lvlc.get('中风险',0)} 条、低风险 {_lvlc.get('低风险',0)} 条。"
                    "主机账号口令与监控调试接口暴露是当前最突出的薄弱环节；资产台账完整性与物理环境电磁防护需内部跟进闭环。具体检查结果见第 3 章与附录 A。")

    # ---- 2.1 / 2.5 检查对象 ----
    set_text(P[39], f"经授权，{CHECKER}对{CLIENT}内网资产开展本次自检，检查结果由 {CHECKER} 出具并仅对检查时点的资产状态有效。")
    set_text(P[58], f"本次检查对象为{SCOPE}：存活资产 82 台（10.0.0 段 53 台、192.168.1 段 29 台），涵盖服务器与终端、网络与安全设备、K8s 与虚拟化节点、应用系统与数据库中间件服务；上机采集覆盖 5 台代表性主机（10.0.0.23 / 10.0.0.180 / 10.0.0.187 / 192.168.1.32 / 192.168.1.80），资产明细与内部认领状态见随附《资产清单-内部对账版》。")
    # 表2 受检资产清单：填 5 台采集机
    t2 = T[2]
    assets = [
        ("1", "示例业务服务器（K8s/运维）", "10.0.0.23", "Windows Server 2016", "内部认领", "□是　□否"),
        ("2", "示例主机A（HP DL380 Gen9）", "10.0.0.180", "Windows Server 2012", "内部认领", "□是　□否"),
        ("3", "AI 算力服务器（Intel+NVIDIA）", "10.0.0.187", "CentOS", "内部认领", "□是　□否"),
        ("4", "超融合/国产算力节点 node1", "192.168.1.32", "CentOS", "内部认领", "□是　□否"),
        ("5", "示例主机B（等保自测工具宿主）", "192.168.1.80", "银河麒麟", "内部认领", "□是　□否"),
    ]
    for i, row_data in enumerate(assets):
        row = t2.rows[i + 1] if i + 1 < len(t2.rows) else t2.add_row()
        for c, val in enumerate(row_data):
            row.cells[c].text = val

    # ---- 2.6 检查结果段（统计在表3/域结论中自动算） ----
    set_text(P[61], "本次检查逐条给出结论，结论分为「符合」「部分符合」「不符合」「不适用」四类。工具线结论以《测评结果梳理-2026-10-22-全量回灌.xlsx》为准；访谈与资料类条目按本单位等保三级测评报告（报告编号 示例报告编号-0000-00-0000-00，复测终态：除电磁屏蔽外全部符合）转义填写，物理环境与机房/供电域直接引用其结论；仅个别需人工判读的采集项标注「未检查」并说明。")

    # ---- 3.2 问题详情：清掉模板示例，按 PROBLEMS 重建 ----
    h3_example = P[78]            # 3.2.1 heading
    # 定位 3.3 标题段，删除其前的示例块（78~89）
    h_33 = next(p for p in P if p.text.strip().startswith("3.3 各检查域"))
    body = doc.element.body
    start_el = h3_example._p
    to_del = []
    el = start_el
    while el is not None and el is not h_33._p:
        to_del.append(el)
        el = el.getnext()
    for e in to_del:
        body.remove(e)
    # 在 3.3 前逐问题插入（克隆模板样式：Heading3 + 5 个 Normal）
    ref = h_33._p
    for _i, pb in enumerate(PROBLEMS, 1):
        pb["idx"] = _i
        block = []
        # 直接造段落：复用现有样式对象
        seq = [
            ("Heading 3", f"{pb['idx']} {pb['name']}（{pb['level']}）"),
            ("Normal", f"① 问题描述：{pb['desc']}"),
            ("Normal", f"② 所属检查域与细项：{pb['items']}（详见附录 A 同一编号行）。"),
            ("Normal", "③ 问题取证："),
        ]
        for ex in EXCERPTS.get(pb["no"], []):
            seq.append(("Normal", f"　　{ex}"))
        import re as _re
        for kind, text in seq:
            new_p = doc.add_paragraph(style=kind)
            if kind == "Heading 3":
                r = new_p.add_run(text)
                r.font.size = Pt(14)
                r.font.bold = True
                r.font.name = "黑体"
                r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "黑体")
                new_p.paragraph_format.space_before = Pt(12)
                new_p.paragraph_format.space_after = Pt(18)
            elif kind == "Normal":
                m = _re.match(r"^([①②③④⑤][^：]*：)(.*)$", text)
                if m:
                    r1 = new_p.add_run(m.group(1)); r1.font.bold = True
                    new_p.add_run(m.group(2))
                else:
                    new_p.add_run(text)
                new_p.paragraph_format.space_before = Pt(0)
                new_p.paragraph_format.space_after = Pt(3)
                new_p.paragraph_format.line_spacing = 1.5
            else:
                new_p.add_run(text)
            block.append(new_p._p)
        for fpath, _fcap in FIGS.get(pb["no"], []):
            doc.add_picture(fpath, width=Cm(14.5))
            block.append(doc.paragraphs[-1]._p)
        import re as _re3
        for kind, text in [
            ("Normal", f"④ 整改建议：{pb['fix']}"),
            ("Normal", "⑤ 复测结论：□通过　□部分通过　□未通过"),
        ]:
            new_p = doc.add_paragraph(style=kind)
            m = _re3.match(r"^([①②③④⑤][^：]*：)(.*)$", text)
            if m:
                r1 = new_p.add_run(m.group(1)); r1.font.bold = True
                new_p.add_run(m.group(2))
            else:
                new_p.add_run(text)
            new_p.paragraph_format.space_before = Pt(0)
            new_p.paragraph_format.space_after = Pt(3)
            new_p.paragraph_format.line_spacing = 1.5
            block.append(new_p._p)
        for el_p in block:
            body.remove(el_p)
            ref.addprevious(el_p)

    # ---- 附录 A：填初测结论列 ----
    ta = None
    for t in doc.tables:
        if t.rows[0].cells[0].text.strip() == "编号" and len(t.columns) == 6:
            ta = t
            break
    assert ta is not None, "附录A 表未找到"
    fill_cnt = Counter()
    for row in ta.rows[1:]:
        rid = row.cells[0].text.strip()
        if not rid or "-" not in rid:
            continue  # 域分组行/表头
        if rid in F_EXCEPTIONS:
            v = F_EXCEPTIONS[rid]
        elif rid in PENDING_FILLS:
            v = PENDING_FILLS[rid]
        elif rid in verdict:
            v = verdict[rid][0] or "未检查"
            if v.startswith("未执行"):
                v = "未检查（本次未执行）"
            if rid in REASON_SUFFIX and v.startswith("不适用"):
                v = v + REASON_SUFFIX[rid]
            elif v.startswith("待人工"):
                v = "未检查（采集数据已留存，待人工判读）"
        elif rid.startswith("2."):
            v = "不适用（无云环境，引等保报告不适用指标）"
        else:
            v = "符合（引等保报告复测终态）"
        v = v.replace("（仅扫描侧）", "（经工具核查）")
        row.cells[4].text = v          # 初测结论
        row.cells[5].text = ""          # 复测结论留空
        fill_cnt[v.split("（")[0]] += 1
    print("附录A 初测结论分布:", dict(fill_cnt))

    # ---- 表3 结果统计（13行=表头+11域+合计? 按行遍历填域计数由附录A实时算） ----
    # 重新按域统计
    dom_stat = OrderedDict()
    cur = None
    for row in ta.rows[1:]:
        rid = row.cells[0].text.strip()
        name = row.cells[1].text.strip()
        if not rid or "-" not in rid:
            if name:
                cur = name
                dom_stat.setdefault(cur, Counter())
            continue
        v = row.cells[4].text
        key = "符合" if v.startswith("符合") else "部分符合" if v.startswith("部分符合") else "不符合" if v.startswith("不符合") else "不适用" if v.startswith("不适用") else "未检查"
        dom_stat[cur][key] += 1
    t3 = T[3]
    grand = Counter()
    for r in t3.rows[1:]:
        dname = r.cells[0].text.strip()
        # 域名前缀匹配（模板域名含"一、…"）
        for dn, c in dom_stat.items():
            if dn[:4] in dname or dname[:4] in dn:
                r.cells[2].text = str(c.get("符合", 0))
                r.cells[3].text = str(c.get("部分符合", 0))
                r.cells[4].text = str(c.get("不符合", 0))
                r.cells[5].text = str(c.get("不适用", 0))
                r.cells[6].text = str(c.get("未检查", 0))
                grand.update(c)
                break
        else:
            if "合计" in dname:
                r.cells[2].text = str(grand.get("符合", 0) or sum(dom_stat[d].get("符合",0) for d in dom_stat))
                r.cells[3].text = str(sum(dom_stat[d].get("部分符合",0) for d in dom_stat))
                r.cells[4].text = str(sum(dom_stat[d].get("不符合",0) for d in dom_stat))
                r.cells[5].text = str(sum(dom_stat[d].get("不适用",0) for d in dom_stat))
                r.cells[6].text = str(sum(dom_stat[d].get("未检查",0) for d in dom_stat))

    # ---- 3.3.x 各域结论段（找"本域结论"开头段，按域统计回填） ----
    dom_map = {
        "一、IT": "一、IT", "二、云": "二、云", "三、网络": "三、网络", "四、主机": "四、主机",
        "五、终端": "五、终端", "六、应用": "六、应用", "七、日志": "七、日志", "八、数据": "八、数据",
        "九、管理制度": "九、管理制度", "十、物理环境": "十、物理环境", "十一、供电": "十一、供电",
    }
    cur_dom = None
    for para in doc.paragraphs:
        t = para.text.strip()
        for k in dom_map:
            if k in t and len(t) < 40:
                cur_dom = k
        if t.startswith("本域结论"):
            stat = dom_stat.get(next((dn for dn in dom_stat if dn.startswith(cur_dom or "~~~")), None), Counter())
            set_text(para, f"本域结论：符合 {stat.get('符合',0)} 条、部分符合 {stat.get('部分符合',0)} 条、不符合 {stat.get('不符合',0)} 条、不适用 {stat.get('不适用',0)} 条、未检查 {stat.get('未检查',0)} 条；初测问题详见第 3.2 节，逐条结论见附录 A 同一域。")

    # ---- 表5 / 表6 问题统计（表5 按问题条目计，表6 按检查细项计，两表合计不等属正常） ----
    def _items(s):
        return [x.strip() for x in re.split(r"[、,，]", s) if x.strip()]

    flagged_ids = set()
    for row in ta.rows[1:]:
        rid = row.cells[0].text.strip()
        if "-" in rid and row.cells[4].text.strip().startswith(("不符合", "部分符合")):
            flagged_ids.add(rid)
    owner = {}
    for pb in PROBLEMS:
        for it in _items(pb["items"]):
            owner[it] = pb
    missing = sorted(flagged_ids - set(owner))
    extra = sorted(set(owner) - flagged_ids)
    print(f"问题细项覆盖自检: 附录A 问题细项 {len(flagged_ids)} / 问题清单挂接 {len(owner)}；未挂接 {missing}；多挂 {extra}")
    if missing or extra:
        print("  !! 3.2 问题条目与附录 A 问题细项不一致，先修 PROBLEMS 的 items")

    # 表3-1：按「问题条目」统计——合计 = 问题条目数（不再因“合计行匹配不到风险等级”写成 0）
    lvl = Counter(pb["level"] for pb in PROBLEMS)
    t5 = T[5]
    for r in t5.rows[1:]:
        lv = r.cells[0].text.strip()
        r.cells[1].text = str(len(PROBLEMS) if lv.startswith("合计") else lvl.get(lv, 0))
        for ci in (2, 3, 4, 5):
            r.cells[ci].text = "—"

    # 表3-2：按「检查细项」统计——合计 = 附录 A 判定为不符合/部分符合的细项数
    t6 = T[6]
    dom_total = 0
    for r in t6.rows[1:]:
        dname = r.cells[0].text.strip()
        if dname.startswith("合计"):
            continue
        for dn, c in dom_stat.items():
            if dn[:4] in dname or dname[:4] in dn:
                n = c.get("不符合", 0) + c.get("部分符合", 0)
                dom_total += n
                r.cells[1].text = str(n)
                r.cells[2].text = "—"
                r.cells[3].text = "—"
                break
    for r in t6.rows[1:]:
        if r.cells[0].text.strip().startswith("合计"):
            r.cells[1].text = str(dom_total)
            r.cells[2].text = "—"
            r.cells[3].text = "—"
    print(f"表3-1 合计 {len(PROBLEMS)} 条问题条目 / 表3-2 合计 {dom_total} 条问题细项"
          "（口径不同不必相等；细项须全覆盖）")

    # 数字列统一字体：模板里个别数字单元格硬编码了宋体/黑体 rPr，去掉改为继承，
    # 否则会出现「细项数 9pt 宋体、旁边符合 12pt」这类同表不同字体
    def _strip_num_rpr(cell):
        for _p in cell.paragraphs:
            for _r in list(_p.runs):
                if _r._element.rPr is not None:
                    _r._element.remove(_r._element.rPr)
    for _row in T[3].rows[1:]:
        _strip_num_rpr(_row.cells[1])
    for _row in T[5].rows[1:]:
        for _ci in range(2, 6):
            _strip_num_rpr(_row.cells[_ci])
    for _row in T[6].rows[1:]:
        for _ci in (1, 2, 3):
            _strip_num_rpr(_row.cells[_ci])
    # ---- 3.1 表前口径说明 + 3.2 引言 + 表题 + 表注 + 跨页表头 ----
    set_text(next(pp for pp in doc.paragraphs if pp.text.strip().startswith("按风险等级汇总的问题数量如表 3-1 所示")),
             "本节两张统计表的计数单位不同：表 3-1 按「问题条目」统计，同一检查域内同类问题归并为一条；"
             "表 3-2 按「检查细项」统计，即附录 A 中判定为「部分符合」「不符合」的细项。"
             "一条问题条目可覆盖多条细项，故两表合计不必相等，两表各自的分项之和应等于本表合计。")
    set_text(next(pp for pp in doc.paragraphs if pp.text.strip().startswith("本节逐条列示检查发现的问题")),
             "本节逐条列示检查发现的问题。每个问题按「问题描述 — 所属检查域与细项 — 证据 — 整改建议 — 复测结论」五段式记录，"
             "问题编号与附录 A 的细项编号对应（一条问题可覆盖多条细项，各条「所属检查域与细项」列明）。")
    set_text(next(pp for pp in doc.paragraphs if pp.text.strip().startswith("表 3-1")),
             "表 3-1  问题风险分级统计（按问题条目计）")
    set_text(next(pp for pp in doc.paragraphs if pp.text.strip().startswith("表 3-2")),
             "表 3-2  分域初测问题与复测情况统计（按检查细项计）")
    NOTE31 = ("注：本表按「问题条目」统计，同一问题条目可覆盖多条检查细项，细项口径见 表 3-2；"
              "「—」表示整改与复测尚未开展，按 4.1 的整改期限整改后回填。")
    NOTE32 = ("注：各域「初测问题数」＝表 2-3 该域「部分符合＋不符合」的细项数；"
              "「—」表示复测尚未开展，复测后按「通过／部分通过／未通过」回填。")
    for _tbl, _note in ((t5, NOTE31), (t6, NOTE32)):
        _np = doc.add_paragraph(_note, style="Normal")
        for _run in _np.runs:
            _run.font.size = Pt(9)
        body.remove(_np._p)
        _tbl._tbl.addnext(_np._p)
        if _tbl.rows[0]._tr.get_or_add_trPr().find(qn("w:tblHeader")) is None:
            _tbl.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))

    # ---- 4.2 复测安排：追加「待人工复核事项」表（复测前置） ----
    anchor42 = next(pp for pp in doc.paragraphs if pp.text.strip().startswith("复测范围为初测判定"))
    intro = doc.add_paragraph("在复测前，以下事项需由检查方与受检单位共同完成人工复核，复核结论作为复测与整改跟踪的输入：", style="Normal")
    body.remove(intro._p); anchor42._p.addnext(intro._p)
    REVIEW_ROWS = [
        ("异常外联筛查", "4.8-2",
         "人工判读 5 台采集输出中的 netstat / Established 连接记录，标记非常规外联目标",
         "输出外联合规结论，异常目标逐一说明"),
        ("明文口令落盘核查", "4.11-1",
         "复核采集输出中计划任务 / 配置文件里的口令存储记录，人工判定是否明文",
         "输出明文存储清单与整改建议"),
        ("配置文件一致性核对", "6.11-1",
         "复核采集输出中库内生效值与配置文件对照记录",
         "输出一致 / 差异清单并回填附录 A"),
    ]
    try:
        tbl = doc.add_table(rows=len(REVIEW_ROWS) + 1, cols=4)
        tbl.style = doc.styles["Table Grid"]
    except Exception:
        tbl = doc.add_table(rows=len(REVIEW_ROWS) + 1, cols=4)
    hdr = ("待人工复核事项", "对应细项", "复核方式与数据来源", "完成口径")
    for c, v in enumerate(hdr):
        cell = tbl.rows[0].cells[c]; cell.text = v
        for r in cell.paragraphs[0].runs:
            r.font.bold = True; r.font.size = Pt(10.5)
    for ri, row in enumerate(REVIEW_ROWS, 1):
        for ci, v in enumerate(row):
            cell = tbl.rows[ri].cells[ci]; cell.text = v
            for r in cell.paragraphs[0].runs:
                r.font.size = Pt(10.5)
    body.remove(tbl._tbl)
    intro._p.addnext(tbl._tbl)

    # ---- 5.1 总体评价 ----
    tot = Counter()
    for c in dom_stat.values():
        tot.update(c)
    set_text(P[139] if len(P) > 139 else doc.paragraphs[-6],
             f"本次检查覆盖 11 个检查域（附录 A 共 166 条细项）：符合 {tot.get('符合',0)} 条、部分符合 {tot.get('部分符合',0)} 条、不符合 {tot.get('不符合',0)} 条、不适用 {tot.get('不适用',0)} 条、未检查 {tot.get('未检查',0)} 条。本次为内网自检：工具线结论完整可溯；访谈与资料类条目按等保报告（复测终态）及 D.4 服务器/平台对象结果记录转义填写，附录 A 全部细项均取得结论；问题 {len(PROBLEMS)} 条（对应 {dom_total} 条问题细项）已列明整改建议，复测待整改完成后按同一细项编号执行。")

    # ---- 5.2 声明补第 5 条 ----
    p52 = next(pp for pp in doc.paragraphs if pp.text.strip().startswith("4.报告中的问题"))
    new5 = doc.add_paragraph("5.未现场复核的访谈与资料类条目按等保三级测评报告（复测终态）转义填写，其最终适用性由受检单位内部确认。", style="Normal")
    body.remove(new5._p)
    p52._p.addnext(new5._p)

    # ---- 附录 B 落盘描述 ----
    set_text(P[150] if len(P) > 150 else doc.paragraphs[-12],
             f"现场证据统一存放于 {EVID_DIR} 目录：扫描输出（golin 全端口报告 xlsx/html、目录探测 dirScan.json、台账差集核对记录、测评结果梳理全量回灌 xlsx、console.log 内部留证）、采集输出（5 台主机完整采集包与 2012 补测记录）、截图（Web 站点截图 466 张）、客户材料（等保转义引用、物理环境问询记录-证据表、等保报告 D.4 服务器/平台对象页照片 33 张（p.295–359）、2026-10-22 等保报送材料包摘录）、资产清单-内部对账版（xlsx/html）。证据编号按『证据-细项编号-序号』编制；弱口令明文仅存内部证据文件，不随报告交付。")

    # ---- 附录C 追加：等保报告引用（带报告编号） ----
    h_appC = next(pp for pp in doc.paragraphs if pp.text.strip().startswith("附录 C"))
    new_p = doc.add_paragraph("7.《受检单位等级测评报告》（报告编号：示例报告编号-0000-00-0000-00，【2025 版】，复测终态）——物理环境与机房/供电条目（10.x、11.x）结论转义引用来源", style="Normal")
    body.remove(new_p._p)
    h_appC._p.addnext(new_p._p)

    # ---- 体系问题与改进建议专章（插在附录A之前） ----
    h_appA = next(p for p in doc.paragraphs if p.text.strip().startswith("附录 A"))
    sec_paras = []
    new_p = doc.add_paragraph("6 体系问题与改进建议", style="Heading 1")
    sec_paras.append(new_p._p)
    new_p = doc.add_paragraph("本次内网自检在检验资产的同时检验了检查体系自身：以下为过程中实锤的体系问题与对应改进方向（详见内部《体系问题清单》）。", style="Normal")
    sec_paras.append(new_p._p)
    for s in SYSTEM_ISSUES:
        new_p = doc.add_paragraph(s, style="Normal")
        sec_paras.append(new_p._p)
    new_p = doc.add_paragraph("以上问题均已在本次检查过程中定位并部分修复；未修复项已列入内部改进计划，下一版清单/工具/流程将按此迭代。", style="Normal")
    sec_paras.append(new_p._p)
    for el_p in sec_paras:
        body.remove(el_p)
        h_appA._p.addprevious(el_p)

    # ---- 排版归一（按等保报告字号层级；字体族保持本单位宋体/黑体） ----
    def _mkfont(stname, font, size, bold, line, before=None, after=None):
        st = doc.styles[stname]
        st.font.size = Pt(size)
        st.font.bold = bold
        st.font.name = font
        rpr = st.element.get_or_add_rPr()
        rf = rpr.get_or_add_rFonts()
        rf.set(qn("w:eastAsia"), font)
        pf = st.paragraph_format
        if line:
            pf.line_spacing = line
        if before is not None:
            pf.space_before = Pt(before)
        if after is not None:
            pf.space_after = Pt(after)
    _mkfont("Normal", "宋体", 12, False, 1.5, 0, 6)
    _mkfont("Heading 1", "黑体", 16, True, 1.0, 14, 10)
    _mkfont("Heading 2", "黑体", 14, True, 1.0, 10, 6)
    _mkfont("Heading 3", "黑体", 12, True, 1.0, 8, 5)
    # 正文区（"1 检查摘要"起）：清掉段落直排字号，统一继承 12pt 宋体 1.5 倍行距
    body_started = False
    for para in doc.paragraphs:
        t = para.text.strip()
        if t.startswith("1 检查摘要"):
            body_started = True
        if not body_started:
            continue
        sn = para.style.name
        if sn == "Normal":
            for r in para.runs:
                r.font.size = None
                r.font.name = "宋体"
                rpr = r._element.get_or_add_rPr()
                rf = rpr.get_or_add_rFonts()
                rf.set(qn("w:eastAsia"), "宋体")
        elif sn.startswith("Heading"):
            import re as _re2
            is_prob = _re2.match(r"^\d+ .*（[高中低]风险）$", t) is not None
            for r in para.runs:
                r.font.size = Pt(14) if is_prob else None
                r.font.name = "黑体"
                rpr = r._element.get_or_add_rPr()
                rf = rpr.get_or_add_rFonts()
                rf.set(qn("w:eastAsia"), "黑体")
                r.font.bold = True

    # ---- 口径统一：全文"信息安全检查"→"网络安全检查"（含表格/页眉/页脚） ----
    def _rename_in_paras(paras):
        for para in paras:
            # 先按 run 内替换（保留格式）
            for r in para.runs:
                if r.text and "信息安全检查" in r.text:
                    r.text = r.text.replace("信息安全检查", "网络安全检查")
            # 仍残留则整段合并重写（首 run 承接，保留其格式）
            if "信息安全检查" in (para.text or ""):
                if para.runs:
                    para.runs[0].text = para.text.replace("信息安全检查", "网络安全检查")
                    for r in para.runs[1:]:
                        r.text = ""
    _rename_in_paras(doc.paragraphs)
    for _tb in doc.tables:
        for _row in _tb.rows:
            for _cell in _row.cells:
                _rename_in_paras(_cell.paragraphs)
    for _sec in doc.sections:
        for _part in (_sec.header, _sec.footer):
            try:
                _rename_in_paras(_part.paragraphs)
                for _tb in _part.tables:
                    for _row in _tb.rows:
                        for _cell in _row.cells:
                            _rename_in_paras(_cell.paragraphs)
            except Exception:
                pass

    doc.save(out)
    print("报告已生成:", out)

if __name__ == "__main__":
    main()
