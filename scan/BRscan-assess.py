#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""BRscan-assess.py —— golin 扫描报告 → 166 条清单对应测评项的建议测评结果

用法:
    python3 BRscan-assess.py <golin报告.xlsx> [选项]

选项:
    --html PATH      指定同名 html（默认自动找同目录同前缀 *.html，读取弱口令/PoC 开关）
    --map PATH       指定《测评项与命令对应表.xlsx》（默认自动定位 ../../内部支撑/）
    --ledger PATH    台账 IP 清单（一行一个，可带 # 注释）；不给则自动在 客户材料/ 里找
    --collect DIR    采集输出根目录（默认自动搜 采集输出/**/测评项_<编号>.txt）
    --dirsearch PATH dirScan.json（golin dirsearch 产物；默认自动找 扫描输出/ 下同名文件）
    --public         受检范围含公网资产（默认按全部内网处理，公网类项直接 不适用）
    --no-policy      关闭"确定值策略包"，回到纯 待人工 口径
    --out PATH       输出 xlsx（默认 BRscan-assess_<时间戳>.xlsx，落在 golin 报告同目录）

结论口径（与现场核查表一致，只出确定值）: 符合 / 部分符合 / 不符合 / 不适用 / 未执行
    扫描侧不能定论的加（仅扫描侧）后缀并在"备注/待补数据"写明还差什么、放哪个目录会自动转值。

确定值策略包（--no-policy 可关）:
    P-SCOPE  受检目标全部为 RFC1918 私网 → 公网依赖项 2.3-1 / 3.2-3 / 3.8-3 判 不适用
    P-LEDGER 台账（--ledger 或 客户材料/ 自动发现）→ 1.4-2 / 3.9-2 自动出差集:
             台账外资产→不符合、无差集→符合、台账未随案→未执行（补台账重跑即转值）
    P-EOL    内置 EOL 规则库（仅收官方有停止支持日期的版本族）比对指纹:
             命中→不符合（列命中）、未命中→符合、无指纹→未执行   [2.7-2 / 6.6-1]
    P-PATH   目录类项按 PoC/dirsearch 路径关键词判定 [6.2-1/6.4-2/6.5-1/6.5-2]:
             命中→不符合、dirsearch 已跑且无命中→符合、未跑→未执行（跑表C10后自动转值）
    P-COLLECT 采集类项查 采集输出/**/测评项_<编号>.txt: 缺→未执行、有→引用文件待判读
    P-ALIGN   证据-主题对齐：PoC 分五类（默认口令/报错页/监控调试/管理页/其余未授权），
              每类测评项只吃自己主题的证据 —— 未授权访问 ≠ 默认口令（6.2-3/6.6-3/6.5-3 只认
              默认口令类 PoC + 口令试算；6.4-1 只认报错页类；6.2-2 只认管理页类）

纪律: 弱口令只记命中数量与服务类型，不抄用户/密码明文（表C 执行纪律④）。
"""
import argparse, os, re, sys, json, datetime, zipfile, ipaddress
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ---------------------------------------------------------------- EOL 规则库（保守：仅官方有停止支持日期的版本族）
EOL_RULES = [
    ("MySQL 5.6/5.7",            r"MySQL[: ]5\.[67]",                    "5.6→2021-02-28、5.7→2023-10-31 停止支持"),
    ("MariaDB 10.0/10.1/10.2",   r"MariaDB[: ]10\.[012]\b",              "2020-03 / 2020-10 / 2022-05 停止支持"),
    ("PostgreSQL 9.x/10",        r"PostgreSQL[: ](9\.[0-9]+|10\.[0-9]+)", "9.x→2020-11、10→2022-11 停止支持"),
    ("PHP 5.x/7.0-7.4",          r"PHP[/ ]([57]\.[0-9]+)",              "5.x→2018-12、7.4→2022-11-28 停止支持"),
    ("Tomcat 6/7/8.0/8.5",       r"[Tt]omcat[/ ]?(6\.|7\.|8\.0\.|8\.5\.)", "6→2016、7→2021、8.0→2018、8.5→2024-03-31 停止支持"),
    ("Apache HTTP Server 2.2",   r"Apache[/ ]2\.2\.",                    "2017-07-11 停止支持"),
    ("IIS 6.0/7.0/7.5/8.0",      r"[Mm]icrosoft-IIS[/ ](6\.0|7\.0|7\.5|8\.0)", "随宿主 Windows 版本停止支持"),
    ("OpenSSL 1.0.x",            r"OpenSSL[/ ]1\.0\.",                   "2019-12 停止支持"),
    ("OpenSSL 1.1.0",            r"OpenSSL[/ ]1\.1\.0",                  "2019-09 停止支持"),
    ("OpenSSL 1.1.1",            r"OpenSSL[/ ]1\.1\.1",                  "2023-09 上游停止支持（发行版延长支持需复核）"),
    ("log4j-core 2.0-2.14.1",    r"log4j[- ]?core?[/ ]2\.(0|1[0-4])",    "CVE-2021-44228 影响版本域"),
]

# ---------------------------------------------------------------- 目录类项关键词（PoC / dirsearch 双源匹配）
DIR_KW = {
    "6.2-1": ["console", "manager", "admin", "login", "phpmyadmin", "wp-login", "wp-admin",
              "actuator", "dashboard", "后台", "管理"],
    "6.4-2": [".git", ".env", ".bak", ".sql", ".zip", ".tar", "backup", "dump",
              "wp-config", "config.php", "config.json", "config.yml", "config.yaml",
              "composer.json", "web.config", ".ds_store", ".htpasswd", "id_rsa", ".svn", ".hg"],
    "6.5-1": ["directory", "diretory", "listing", "dir-list", "目录浏览", "index of"],
    "6.5-2": ["version", "changelog", "readme", "package.json", "sample", "example", "历史版本"],
}

# ---- PoC 证据分类（证据-主题对齐：默认口令类项只吃默认口令类证据，未授权类不串味）
POC_CAT_RULES = [
    ("defaultpass", r"default[-_ ]?(pass|password|passwd|cred)|默认口令|默认密码|默认账号|弱口令"),
    ("errorpage",   r"报错|错误页|error[-_ ]?page|stack[-_ ]?trace|绝对路径泄露|异常回显"),
    ("mondebug",    r"pprof|metrics|exporter|actuator|prometheus|zabbix|监控|调试|debug"),
    ("mgmtpage",    r"swagger|console|管理后台|后台入口|登录页|admin[-_ ]?page|login[-_ ]?page"),
]
POC_CATS_ALL = {"defaultpass", "errorpage", "mondebug", "mgmtpage", "unauth_other"}
POC_UNAUTH_FAMILY = {"mondebug", "mgmtpage", "unauth_other"}

def poc_cat(v):
    """(url, poc名, 描述) → defaultpass / errorpage / mondebug / mgmtpage / unauth_other"""
    s = f"{v[1]} {v[2]}".lower()
    for cat, pat in POC_CAT_RULES:
        if re.search(pat, s, re.I):
            return cat
    return "unauth_other"

# ---------------------------------------------------------------- 读取 golin 报告
def load_golin(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    data = {"ports": [], "alives": [], "vulns": [], "cracks": []}
    if "扫描端口结果" in wb.sheetnames:
        for r in list(wb["扫描端口结果"].iter_rows(values_only=True))[1:]:
            if r and r[1]:
                data["ports"].append((str(r[1]).strip(), str(r[2]).strip(), (r[3] or "").strip()))
    if "存活主机" in wb.sheetnames:
        for r in list(wb["存活主机"].iter_rows(values_only=True))[1:]:
            if r and r[1]:
                data["alives"].append(str(r[1]).strip())
    if "漏洞资产" in wb.sheetnames:
        for r in list(wb["漏洞资产"].iter_rows(values_only=True))[1:]:
            if r and r[1]:
                data["vulns"].append((str(r[1]).strip(), str(r[2]).strip(), (r[3] or "").strip()))
    if "弱口令资产" in wb.sheetnames:      # 列: 序号 主机 端口 用户 密码 模式 —— 不读 用户/密码
        for r in list(wb["弱口令资产"].iter_rows(values_only=True))[1:]:
            if r and r[1]:
                data["cracks"].append((str(r[1]).strip(), str(r[2]).strip(), (r[5] or "").strip()))
    wb.close()
    return data

def load_flags(html_path):
    if not html_path or not os.path.exists(html_path):
        return None, None
    raw = open(html_path, encoding="utf-8", errors="replace").read()
    def flag(pat):
        m = re.search(pat + r'[^>]*>\s*(是|否)', raw)
        return (m.group(1) == "是") if m else None
    return flag(r"启用口令扫描"), flag(r"启用漏洞扫描")

IP_RE = re.compile(r"\b(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\b")

def ips_from_text(text):
    out = set()
    for ip in IP_RE.findall(text):
        try:
            ipaddress.ip_address(ip); out.add(ip)
        except ValueError:
            pass
    return out

def ips_from_file(path):
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in (".txt", ".csv", ".md", ".tsv", ".log", ""):
            return ips_from_text(open(path, encoding="utf-8", errors="replace").read())
        if ext in (".xlsx", ".xlsm"):
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            s = set()
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    for v in row:
                        if v is not None:
                            s |= ips_from_text(str(v))
            wb.close(); return s
        if ext == ".docx":
            with zipfile.ZipFile(path) as z:
                xml = z.read("word/document.xml").decode("utf-8", errors="replace")
            return ips_from_text(re.sub(r"<[^>]+>", " ", xml))
    except Exception:
        return set()
    return set()

def find_ledger(explicit, roots):
    """台账来源: --ledger 显式 > 各根下 客户材料/ 目录内文件（≥3 个 IP 才算台账）"""
    if explicit:
        return ips_from_file(explicit) if os.path.isfile(explicit) else set(), explicit
    for root in roots:
        if not root or not os.path.isdir(root):
            continue
        for base, dirs, files in os.walk(root):
            if os.path.basename(base) != "客户材料":
                continue
            for f in sorted(files):
                if f.startswith("~$") or f.startswith("."):
                    continue
                ips = ips_from_file(os.path.join(base, f))
                if len(ips) >= 3:
                    return ips, os.path.join(base, f)
    return None, None

def load_entries_flexible(path):
    """宽容读取 dirScan.json：兼容 golin 并发写叠出的多文档/局部损坏；去重；剔除探针路径"""
    if not os.path.exists(path):
        return []
    raw = open(path, encoding="utf-8", errors="replace").read()
    dec, out, i, n = json.JSONDecoder(), [], 0, len(raw)
    while i < n:
        if raw[i] in "[{":
            try:
                obj, end = dec.raw_decode(raw, i)
            except Exception:
                i += 1; continue
            if isinstance(obj, list):
                out.extend(x for x in obj if isinstance(x, dict))
            elif isinstance(obj, dict):
                out.append(obj)
            i = end
        else:
            i += 1
    seen, res = set(), []
    for e in out:
        u = str(e.get("Url", ""))
        if "/dsh-canary-" in u:
            continue
        k = (u, e.get("Code"), (e.get("Title") or "").strip())
        if k in seen:
            continue
        seen.add(k); res.append(e)
    return res

def load_dirscan(explicit, roots):
    if explicit and os.path.isfile(explicit):
        path = explicit
    else:
        path = None
        for root in roots:
            if root and os.path.isdir(root):
                cand = os.path.join(root, "dirScan.json")
                if os.path.isfile(cand):
                    path = cand; break
        if path is None:
            for root in roots:
                if not root or not os.path.isdir(root):
                    continue
                for base, dirs, files in os.walk(root):
                    if "dirScan.json" in files:
                        path = os.path.join(base, "dirScan.json"); break
                if path: break
    if not path:
        return None, None
    entries = load_entries_flexible(path)
    if not os.path.exists(path):
        return None, path
    return entries, path

def find_collect_files(explicit, roots, ids):
    """采集输出发现: **/测评项_<编号>.txt"""
    want = {i: [] for i in ids}
    search_roots = [explicit] if explicit else [r for r in roots if r and os.path.isdir(r)]
    for root in search_roots:
        for base, dirs, files in os.walk(root):
            for f in files:
                m = re.fullmatch(r"测评项_([\d.]+-\d+)\.txt", f)
                if m and m.group(1) in want:
                    want[m.group(1)].append(os.path.join(base, f))
    return want

def find_dirmanifest(roots):
    """接续脚本写的覆盖率清单 dirscan_manifest.json（与 dirScan.json 同目录）"""
    for root in roots:
        if not root or not os.path.isdir(root):
            continue
        cand = os.path.join(root, "dirscan_manifest.json")
        if os.path.isfile(cand):
            try:
                return json.load(open(cand, encoding="utf-8")), cand
            except Exception:
                return None, cand
        for base, dirs, files in os.walk(root):
            if "dirscan_manifest.json" in files:
                p = os.path.join(base, "dirscan_manifest.json")
                try:
                    return json.load(open(p, encoding="utf-8")), p
                except Exception:
                    return None, p
    return None, None

def dir_coverage(e):
    """(scanned, total) 或 None（无清单，旧数据）"""
    m = getattr(e, "dirmanifest", None)
    if not m:
        return None
    total = m.get("targets_total") or 0
    scanned = len(m.get("targets", []))
    return scanned, total

def load_map(map_path):
    wb = openpyxl.load_workbook(map_path, data_only=True)
    items, tabc = [], {}
    ws = wb["测评项与命令对应表"]
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if r and r[0]:
            items.append({"id": str(r[0]).strip(), "domain": (r[1] or "").strip(),
                          "text": (r[2] or "").strip(), "level": (r[3] or "").strip(),
                          "src": (r[5] or "").strip()})
    if "表C-golin条目归属" in wb.sheetnames:
        for r in list(wb["表C-golin条目归属"].iter_rows(values_only=True))[1:]:
            if r and r[0]:
                tabc[str(r[0]).strip()] = ((r[1] or "").strip(), (r[2] or "").strip())
    wb.close()
    return items, tabc

def parse_src(src):
    c = [int(n) for n in re.findall(r"表C 第 (\d+) 条", src)]
    b = [int(n) for n in re.findall(r"表B 第 (\d+) 条", src)]
    return sorted(set(c)), sorted(set(b))

# ---------------------------------------------------------------- 分类与常量
SCAN_ROWS    = {1, 2, 3, 4, 5, 6, 7, 8, 9, 12, 13, 14, 15, 27}
DIR_ROWS     = {10, 11}
COLLECT_ROWS = set(range(16, 26))
SCOPE_ITEMS  = {"2.3-1", "3.2-3", "3.8-3"}          # P-SCOPE 公网依赖项
RISK_PORTS = {"21","22","23","25","53","69","110","135","139","143","161","389","445",
              "512","513","514","873","1433","1521","2049","3306","3389","5900","6379",
              "61616","9200","11211","27017","5236","8080","8000","9090"}
MGMT_PORTS = {"623", "5900", "5988", "5989"}
CLEARTEXT_MGMT = {"21", "23", "161", "514"}

# ---------------------------------------------------------------- 证据模型
class Evidence:
    def __init__(self, d, crack_on, poc_on, ledger, ledger_src, dirscan, dirscan_src,
                 collect_files, internal, public_scope, policy, dirmanifest=None,
                 wildcard_n=0, dirscan_excluded=0):
        self.alives = d["alives"]; self.ports = d["ports"]
        self.vulns = d["vulns"]; self.cracks = d["cracks"]
        self.crack_on, self.poc_on = crack_on, poc_on
        self.ledger, self.ledger_src = ledger, ledger_src
        self.dirscan, self.dirscan_src = dirscan, dirscan_src    # dirscan=None 未跑/未找到
        self.collect_files = collect_files                        # {编号: [路径...]}
        self.internal, self.public_scope, self.policy = internal, public_scope, policy
        self.dirmanifest = dirmanifest                # 接续脚本覆盖率清单（可空）
        self.wildcard_n = wildcard_n                  # 通配响应(soft-404)站数：无法枚举 → 只能（仅扫描侧）
        self.dirscan_excluded = dirscan_excluded      # 因通配被剔除的误报条数
        self.pending_items = set()      # 处理器登记：本项仍有旁证缺口 → 结论加（仅扫描侧）

    def open_ports(self): return {(h, p) for h, p, _ in self.ports}
    def risk_hits(self):
        return sorted([(h, p) for h, p, _ in self.ports if p in RISK_PORTS], key=lambda x: (x[0], int(x[1])))
    def mgmt_hits(self):
        return sorted([(h, p) for h, p, _ in self.ports if p in MGMT_PORTS])
    def web_hits(self):  return [(h, p, pr) for h, p, pr in self.ports if "WEB" in pr.upper()]
    def db_hits(self):   return [(h, p, pr) for h, p, pr in self.ports if "数据库" in pr]
    def telnet_hits(self):
        return [(h, p, pr) for h, p, pr in self.ports if "telnet" in pr.lower() or p == "23"]
    def clear_mgmt_hits(self):
        return sorted([(h, p) for h, p, _ in self.ports if p in CLEARTEXT_MGMT], key=lambda x: (x[0], int(x[1])))
    def fingerprints(self):
        seen = []
        for _, _, pr in self.ports:
            if pr and pr not in seen: seen.append(pr)
        return seen
    def rdp_open(self): return [h for h, p, _ in self.ports if p == "3389"]
    def ledger_diff(self):
        if self.ledger is None: return None
        alive = set(self.alives)
        return sorted(alive - self.ledger), sorted(self.ledger & alive)

def head(items, n=6, fmt=lambda x: str(x)):
    s = "、".join(fmt(i) for i in items[:n])
    return s + (f" 等共 {len(items)} 项" if len(items) > n else "")

# ---------------------------------------------------------------- 表C 条目规则
def rule_1(e: Evidence):
    实测 = f"存活 {len(e.alives)} 台：{head(e.alives, 10)}；开放端口合计 {len(e.ports)} 个"
    diff = e.ledger_diff()
    if diff is None:
        if e.policy:
            return 实测, "未执行", "台账未随案：把台账放进 客户材料/（或 --ledger 指定）重跑本脚本即自动出差集（台账外资产→不符合 / 无差集→符合）"
        return 实测, "待人工", "需与台账差集；--ledger 指定台账可自动判定"
    extra, both = diff
    if extra:
        return 实测 + f"；台账外资产 {len(extra)} 台：{head(extra, 10)}", "不符合", f"台账内 {len(both)} 台命中；台账外 {len(extra)} 台须登记或核实"
    return 实测 + "；与台账无差集", "符合", f"台账（{e.ledger_src}）{len(both)} 台全部存活命中"

def rule_2(e: Evidence):
    hits = e.risk_hits()
    if not hits:
        return "高危端口集合（含 22/23/445/3306/3389/6379 等）无命中", "符合", "本网段未见高危端口暴露"
    实测 = f"高危端口命中 {len(hits)} 处：" + head([f"{h}:{p}" for h, p in hits], 8)
    return 实测, "待人工", "须与设备侧映射表/公网可达性核对（清单要求「与映射表一致」）；命中明细见实测明细"

def rule_3(e: Evidence):
    w = e.web_hits()
    return f"Web 口 {len(w)} 个", None, None

def rule_4(e: Evidence):
    return f"数据库口 {len(e.db_hits())} 个", None, None

def rule_5(e: Evidence):
    hits = e.mgmt_hits()
    if not hits:
        return "带外/管理口（623/5900/5988/5989）无命中", "符合", "本网段未见管理口暴露（隔离性仍可由表B 第 4 条佐证）"
    实测 = f"管理口命中：{head([f'{h}:{p}' for h, p in hits], 8)}"
    return 实测, "部分符合", "发现管理口暴露；隔离性与出厂默认账号须上机核实（表B 第 4 条）"

def rule_6(e: Evidence):
    tel = e.telnet_hits(); clr = e.clear_mgmt_hits()
    if tel:
        return f"Telnet 开放：{head([f'{h}:{p}' for h, p, _ in tel], 6)}", "不符合", "Telnet 明文管理协议，建议以 SSH 替代"
    if clr:
        return f"明文管理类端口开放：{head([f'{h}:{p}' for h, p in clr], 8)}", "不符合", "FTP/SNMPv1-v2c/syslog 等明文管理面，建议收敛或加密"
    plain80 = [x for x in e.web_hits() if x[1] in ("80", "81", "82", "8080", "8000", "8888")]
    if plain80:
        return f"无 Telnet；HTTP 明文站点 {len(plain80)} 处", "部分符合", "存在 HTTP 明文站点，须甄别是否管理面并建议 HTTPS 化"
    return "无 Telnet/FTP/SNMP 明文管理口，站点均为 HTTPS", "符合", ""

def rule_7(e: Evidence):
    实测 = f"扫描存活 {len(e.alives)} 台：{head(e.alives, 10)}"
    diff = e.ledger_diff()
    if diff is None:
        if e.policy:
            return 实测, "未执行", "台账未随案：台账放入 客户材料/（或 --ledger）重跑即自动出差集；接入管控侧另需表B 第 27 条"
        return 实测, "待人工", "私接判定=存活清单与台账差集；--ledger 提供台账可自动判定；另需表B 第 27 条接入管控数据"
    extra, both = diff
    if extra:
        return 实测 + f"；台账外 {len(extra)} 台：{head(extra, 10)}", "不符合", "疑似未登记接入设备，须与客户核实"
    return 实测 + "；无台账外设备", "符合", "台账（%s）无差集；接入管控（802.1X/绑定）可由表B 第 27 条补充佐证" % (e.ledger_src or "")

def rule_9(e: Evidence):
    """P-EOL: 指纹 × 内置 EOL 规则库 → 确定值（--no-policy 时退回人工比对）"""
    fps = e.fingerprints()
    if not fps:
        return "（本轮未采集到组件/版本指纹）", "未执行", "无指纹可比对"
    if not e.policy:
        return f"指纹 {len(fps)} 条", "待人工", "需人工对照 EOL/CVE 表单（--no-policy 已关规则库）"
    hits = []
    for fp in fps:
        for name, pat, note in EOL_RULES:
            if re.search(pat, fp, re.I):
                hits.append((fp, name, note))
    if hits:
        实测 = f"指纹 {len(fps)} 条，命中停止支持规则 {len(hits)} 处：" + head(
            [f"{fp[:60]}〔{n}〕" for fp, n, _ in hits], 4)
        return 实测, "不符合", "内置 EOL 规则库 %d 条（清单见判定口径）：%s" % (
            len(EOL_RULES), "；".join(sorted({f"{n}→{d}" for _, n, d in hits})))
    return (f"指纹 {len(fps)} 条，按内置 EOL 规则库（{len(EOL_RULES)} 条）未命中", "符合",
            "规则库仅收录有官方停止支持日期的版本族（MySQL/MariaDB/PostgreSQL/PHP/Tomcat/Apache/IIS/OpenSSL/log4j）；无版本号指纹不参与判定")

def rule_9_item_231(e: Evidence):
    if e.internal and not e.public_scope and e.policy:
        return (f"受检目标 {len(e.alives)} 台全部为内网地址（RFC1918），无公网 IP/EIP",
                "不适用", "公网安全组 0.0.0.0/0 放通核查针对公网资产；范围含公网时加 --public 重跑（届时仍需表B 第 10 条查云控制台）")
    return f"公网可达面参考：开放端口 {len(e.ports)} 个（{len(e.alives)} 台）", "待人工", "安全组/ACL 源 0.0.0.0/0 规则须查云控制台（表B 第 10 条），扫描报告不能替代"

def rule_10_11(e: Evidence):
    if e.dirscan:
        n = len(e.dirscan)
        return (f"dirsearch 已执行：命中路径 {n} 条（{e.dirscan_src}）", "符合" if n == 0 else "待人工",
                "本项按路径关键词判定请看备注项；命中明细见 dirScan.json")
    return ("（dirsearch 未执行，结果不进扫描报告）", "未执行",
            "跑表C 第 10/11 条（golin dirsearch）后把 dirScan.json 放 扫描输出/ 重跑本脚本即自动转值")

def rule_12(e: Evidence):
    if e.poc_on is False:
        return "（本轮未启用 PoC 扫描）", "未执行", "复测时去掉 --nopoc 重跑"
    if not e.vulns:
        return f"PoC 命中 0 条（启用状态：{e.poc_on}）", "符合", "内置 PoC 库未发现命中"
    实测 = f"PoC 命中 {len(e.vulns)} 条：" + head([f"{u}｜{c}｜{d}" for u, c, d in e.vulns], 5)
    return 实测, "不符合", "存在未授权访问/信息泄露等暴露痕迹，仅记录 URL 与现象（不尝试利用）"

def rule_13(e: Evidence):
    if e.crack_on is False:
        return "（本轮未启用弱口令扫描）", "未执行", "经书面同意并确认锁定策略后按表C 第 13 条开启"
    if not e.cracks:
        return f"弱口令命中 0（已按字典对 {len(e.ports)} 个开放服务口逐服务试算）", "符合", "命中数 0，按纪律不记明细"
    modes = {}
    for h, p, m in e.cracks:
        modes[m or "服务"] = modes.get(m or "服务", 0) + 1
    实测 = f"弱口令命中 {len(e.cracks)} 处：" + "、".join(f"{k}×{v}" for k, v in modes.items())
    return 实测, "不符合", f"命中主机端口：{head([f'{h}:{p}' for h, p, _ in e.cracks], 6)}（纪律：明文不落表，仅记命中数量）"

def rule_14(e: Evidence):
    rdp = e.rdp_open()
    rdp_crack = [c for c in e.cracks if "rdp" in (c[2] or "").lower()]
    if not rdp and not rdp_crack:
        return "本网段无 3389 开放", "不适用", "未开放 RDP，无试算对象"
    if rdp_crack:
        return f"RDP 弱口令命中 {len(rdp_crack)} 处", "不符合", "立即加固（明文不落表）"
    if e.crack_on is False:
        return f"3389 开放于 {head(sorted(set(rdp)), 5)}", "未执行", "RDP 试算需 --rdp 单开（防锁定，须书面同意）"
    return f"3389 开放于 {head(sorted(set(rdp)), 5)}", "未执行", "本轮 RDP 试算默认关闭（--rdp 未开），须单独执行表C 第 14 条"

def rule_15(e: Evidence):
    dev = [c for c in e.cracks if (c[2] or "").lower() in ("telnet", "ftp", "smb", "tomcat", "rdp")]
    if e.crack_on is False:
        return "（本轮未启用弱口令扫描）", "未执行", "设备/后台默认口令试算须开弱口令（表C 第 15 条）"
    if dev:
        return f"设备类服务口令命中 {len(dev)} 处：{head([f'{h}:{p}' for h, p, _ in dev], 6)}", "不符合", "按厂商默认口令表整改（明文不落表）"
    if e.cracks:
        return f"全服务弱口令命中 {len(e.cracks)} 处（含设备类需甄别）", "部分符合", f"命中：{head([f'{h}:{p}' for h, p, _ in e.cracks], 6)}"
    return "设备/后台类服务口（Telnet/FTP/SMB/Tomcat 等）口令试算未命中", "符合", "AAA 集中认证/出厂口令整改可由表B 第 22 条补充佐证"

RULES = {1: rule_1, 27: rule_1, 2: rule_2, 3: rule_3, 4: rule_4, 5: rule_5, 6: rule_6,
         7: rule_7, 8: None, 9: rule_9, 10: rule_10_11, 11: rule_10_11,
         12: rule_12, 13: rule_13, 14: rule_14, 15: rule_15}

# ---- 测评项级策略处理器 ----
def item_scope_323(e: Evidence):
    if e.internal and not e.public_scope and e.policy:
        return (f"受检目标 {len(e.alives)} 台全部为内网地址，无公网可达端口与 WAF 场景", "不适用",
                "公网映射/绕 WAF 核查针对公网资产；范围含公网加 --public 重跑")
    return ("公网可达面参考：开放端口 %d 个" % len(e.ports)), "待人工", "须与设备侧映射表核对（表B 第 20/21 条）"

def item_scope_383(e: Evidence):
    if e.internal and not e.public_scope and e.policy:
        return ("受检目标全部为内网地址，AP/无线控制器「可从公网访问」前提不存在", "不适用",
                "无线段专项（可达性+默认口令）若在受检范围，按表C 第 8/15 条对 WLAN 段单独跑")
    return rule_8_fallback(e)

def rule_8_fallback(e: Evidence):
    wireless_kw = ("无线", "WLAN", "UniFi", "AirOS", "TP-LINK", "无线控制器")
    sus = [f"{h}:{p}「{pr}」" for h, p, pr in e.ports if any(k.lower() in pr.lower() for k in wireless_kw)]
    实测 = f"本段服务口 {len(e.ports)} 个；疑似无线/网络设备指纹 {len(sus)} 处" + (f"：{head(sus, 5)}" if sus else "")
    return 实测, "待人工", "须先确认受检范围含无线段（表C 第 8 条）"

def make_dir_item(iid):
    """P-PATH: 目录类项 —— PoC/dirsearch 双源关键词 → 确定值"""
    def handler(e: Evidence):
        kw = DIR_KW[iid]
        blobs = [(f"{u}｜{c}｜{d}", "PoC") for u, c, d in e.vulns]
        if e.dirscan:
            blobs += [(f"{d.get('Url','')} {d.get('Title','')} {d.get('App','')}", "dirsearch") for d in e.dirscan]
        hits = []
        for text, tag in blobs:
            low = text.lower()
            if any(k.lower() in low for k in kw):
                hits.append((text, tag))
        if hits:
            备注 = "证据来源：" + "、".join(sorted({t for _, t in hits})) + f"；关键词集 {kw[:4]}…"
            cov = dir_coverage(e)
            if cov and cov[0] < cov[1]:
                备注 += f"；dirsearch 覆盖 {cov[0]}/{cov[1]} 站（其余站点命中另行补跑，不影响本条命中成立）"
            if e.wildcard_n:
                备注 += f"；另剔除通配站 {e.wildcard_n} 个的 {e.dirscan_excluded} 条 soft-404 误报"
            return (f"路径命中 {len(hits)} 处：" + head([t[:80] for t, _ in hits], 4), "不符合", 备注)
        if e.dirscan is not None:
            cov = dir_coverage(e)
            dictinfo = (e.dirmanifest or {}).get("dict", "?")
            wild_note = ""
            if e.wildcard_n:
                e.pending_items.add(iid)
                wild_note = f"；通配响应站 {e.wildcard_n} 个无法枚举（soft-404，其 {e.dirscan_excluded} 条误报已剔除），该部分只能人工核"
            if cov is None:
                return (f"PoC {len(e.vulns)} 条 + dirsearch {len(e.dirscan)} 条命中均不含本项关键词",
                        "符合", f"dirsearch 已执行（{e.dirscan_src}）；无覆盖率清单，按已见数据判定；关键词集见判定口径" + wild_note)
            scanned, total = cov
            if total and scanned < total:
                e.pending_items.add(iid)
                return (f"PoC {len(e.vulns)} 条 + dirsearch 已跑 {scanned}/{total} 站，本项关键词命中 0",
                        "符合", f"dirsearch 覆盖 {scanned}/{total}（字典 {dictinfo} 条，{e.dirscan_src}）；补齐剩余站点（重跑 BRscan-dirsearch.py）后转全量定级" + wild_note)
            return (f"PoC {len(e.vulns)} 条 + dirsearch 全量 {scanned} 站命中均不含本项关键词",
                    "符合", f"dirsearch 全覆盖 {scanned}/{total}（字典 {dictinfo} 条，{e.dirscan_src}）；关键词集见判定口径" + wild_note)
        return ("（PoC 无本项路径命中；dirsearch 未执行）", "未执行",
                "跑 BRscan-dirsearch.py（或表C 第 10/11 条手跑）后重跑本脚本即自动转值")
    return handler

def make_poc_handler(iid, cats, none_verdict, none_note):
    """P-ALIGN: 只按指定 PoC 类别取证据；类别 0 命中时给确定的 none_verdict"""
    def handler(e: Evidence):
        if e.poc_on is False:
            return "（本轮未启用 PoC 扫描）", "未执行", "复测时去掉 --nopoc 重跑"
        hits = [v for v in e.vulns if poc_cat(v) in cats]
        if hits:
            实测 = f"本项判据类 PoC 命中 {len(hits)} 条：" + head([f"{u}｜{c}｜{d}" for u, c, d in hits], 4)
            return 实测, "不符合", f"证据-主题对齐：只计类别 {sorted(cats)}（PoC 总命中 {len(e.vulns)} 条，其余类别不计入本项）"
        return (f"本项判据类 PoC 命中 0（PoC 总扫命中 {len(e.vulns)} 条，均属其他类别，不计入本项）",
                none_verdict, none_note)
    return handler

def make_defaultpass_handler(iid, pending_dir=False):
    """默认口令类项（6.2-3/6.6-3/6.5-3）：默认口令类 PoC + 口令试算 双证据；未授权类证据不计入"""
    def handler(e: Evidence):
        dp = [v for v in e.vulns if poc_cat(v) == "defaultpass"]
        r15 = rule_15(e)
        if dp:
            实测 = f"默认口令类 PoC 命中 {len(dp)} 条：" + head([f"{u}｜{c}" for u, c, _ in dp], 4)
            return 实测, "不符合", "PoC 默认口令规则命中；证据-主题对齐：未授权访问类不计入本项"
        if r15[1] == "未执行":
            return (f"默认口令类 PoC 命中 0；{r15[0]}", "未执行", r15[2])
        实测 = f"默认口令类 PoC 命中 0（PoC 总命中 {len(e.vulns)} 条均为未授权/泄露类，不计入本项）；{r15[0]}"
        备注 = r15[2] + "；证据-主题对齐：未授权访问 ≠ 默认口令"
        if pending_dir:      # 仅 6.5-3 这类还含「页面/枚举面」的项才挂覆盖率缺口
            cov = dir_coverage(e)
            dir_missing = (e.dirscan is None) or (cov is not None and cov[0] < cov[1]) or (e.wildcard_n > 0)
            if dir_missing:
                e.pending_items.add(iid)
                if e.dirscan is None:
                    备注 += "；页面/目录枚举（表C 第 10 条 / BRscan-dirsearch.py）未执行，本结论仅覆盖口令侧"
                elif e.wildcard_n > 0:
                    备注 += f"；页面/目录枚举覆盖 {cov[0]}/{cov[1]} 站，其中通配站 {e.wildcard_n} 个无法枚举，页面面按其余站点计"
                else:
                    备注 += f"；页面/目录枚举覆盖 {cov[0]}/{cov[1]} 站，本结论仅覆盖口令侧（补跑接续脚本后转全量）"
        return 实测, r15[1], 备注
    return handler

ITEM_RULE = {
    # P-SCOPE 公网依赖项（内部再按 e.policy/e.internal 判 不适用 或回退）
    "2.3-1": rule_9_item_231, "3.2-3": item_scope_323, "3.8-3": item_scope_383,
    # P-PATH 目录类项
    "6.2-1": make_dir_item("6.2-1"), "6.4-2": make_dir_item("6.4-2"),
    "6.5-1": make_dir_item("6.5-1"), "6.5-2": make_dir_item("6.5-2"),
    # P-ALIGN 痕迹/无认证/接口类（只吃未授权家族证据）
    "2.3-4": make_poc_handler("2.3-4", POC_CATS_ALL, "符合", "PoC 全类别命中 0"),
    "2.7-3": make_poc_handler("2.7-3", POC_CATS_ALL, "符合", "PoC 全类别命中 0"),
    "6.2-2": make_poc_handler("6.2-2", {"mgmtpage", "unauth_other"}, "符合",
                              "PoC 管理页未授权类命中 0；页面枚举侧见表C 第 10 条（6.2-1/6.5-1）"),
    "6.6-2": make_poc_handler("6.6-2", POC_UNAUTH_FAMILY, "符合",
                              "PoC 未授权/监控/调试类命中 0"),
    "6.4-1": make_poc_handler("6.4-1", {"errorpage"}, "符合",
                              "报错页泄漏类 PoC 命中 0（其余类别不属本项）；如需系统性核对可人工抽检 Web 错误页"),
    # P-ALIGN 默认口令类（只吃默认口令类 PoC + 口令试算）
    "6.2-3": make_defaultpass_handler("6.2-3"),
    "6.6-3": make_defaultpass_handler("6.6-3"),
    "6.5-3": make_defaultpass_handler("6.5-3", pending_dir=True),
}

ORDER = {"不符合": 3, "部分符合": 2, "待人工": 1, "符合": 0, "未执行": 0}

def evaluate_item(item, e, tabc):
    cid, bid = parse_src(item["src"])
    unknown = [c for c in cid if f"表C 第 {c} 条" not in tabc]
    scan_rows = [c for c in cid if c in SCAN_ROWS]
    dir_rows = [c for c in cid if c in DIR_ROWS]
    coll_rows = [c for c in cid if c in COLLECT_ROWS]

    # P-SCOPE 顶层覆盖（纯范围判断, 先于一切条目规则）
    if e.policy and e.internal and not e.public_scope and item["id"] in SCOPE_ITEMS:
        实测, 结论, 备注 = ITEM_RULE[item["id"]](e)
        return 实测, 结论, 备注, cid, bid

    parts, notes, verdicts = [], [], []
    seen = set()
    if e.policy and item["id"] in ITEM_RULE:   # 策略级处理器（证据-主题对齐），--no-policy 时退回条目规则
        实测, 结论, 备注 = ITEM_RULE[item["id"]](e)
        parts.append(实测); verdicts.append(结论); notes.append(备注)
        scan_rows, dir_rows = [], []     # 已由测评项级处理器给出确定值

    for c in scan_rows + (dir_rows if e.policy else []):
        fn = RULES.get(c)
        if fn is None:                  # 无条目级规则的表C 行（第 8 条由 3.8-3 项级策略接管）
            notes.append(f"表C 第 {c} 条：由对应测评项级策略处理")
            continue
        实测, 结论, 备注 = fn(e)
        if 结论 is None:                 # 支撑类条目(3/4): 只作实测补充
            notes.append(f"表C 第 {c} 条侧证：{实测}")
            continue
        key = (实测, 结论, 备注)
        if key in seen:
            continue
        seen.add(key)
        parts.append(f"[表C{c}] {实测}"); verdicts.append(结论)
        if 备注: notes.append(f"[表C{c}] {备注}")

    if dir_rows and not e.policy:
        notes.append("另需表C 第 %s 条 dirsearch 输出（不在扫描报告内；--no-policy 已关自动判定）" % "/".join(map(str, dir_rows)))

    # P-COLLECT: 采集类证据文件发现
    coll_files = e.collect_files.get(item["id"], []) if e.collect_files else []
    if coll_rows:
        if coll_files:
            notes.append("另需表C 第 %s 条采集数据：已发现 %d 份（%s），按手册判读" %
                         ("/".join(map(str, coll_rows)), len(coll_files), coll_files[0]))
        else:
            notes.append("表C 第 %s 条采集输出缺失：未见 items/测评项_%s.txt（放入 采集输出/ 后重跑即引用）" %
                         ("/".join(map(str, coll_rows)), item["id"]))
    if bid:
        notes.append("另需表B 第 %s 条上机数据" % "/".join(map(str, bid)))
    if unknown:
        notes.append(f"警告: 对应表引用的表C 第 {unknown} 条不存在（表C 仅 29 条），请核对对应表")

    if not parts:                        # 纯采集类/异常类 → 确定状态
        if coll_files:
            return (f"（扫描报告无此数据；采集输出已就位: {coll_files[0]}）", "待人工", "；".join(notes), cid, bid)
        if e.policy:
            reason = ("采集输出缺失" if coll_rows else "本项数据不产自 golin 扫描报告")
            return ("（扫描报告无对应数据）", "未执行",
                    f"{reason}；" + "；".join(notes), cid, bid)
        return ("（扫描报告无对应数据）", "待人工", "；".join(notes), cid, bid)

    verdict = max(verdicts, key=lambda v: ORDER[v])
    pending = (bool((dir_rows and not e.policy) or coll_rows or bid)
               or item["id"] in e.pending_items
               or any(v == "未执行" for v in verdicts))
    verdict_cell = verdict + ("（仅扫描侧）" if pending and verdict != "未执行" else "")
    return "；".join(parts), verdict_cell, "；".join(notes), cid, bid

# ---------------------------------------------------------------- 输出
def write_out(path, rows, e, tabc, n_items, extra_meta):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "测评结果"
    headers = ["编号", "检查域", "级别", "检测内容", "依据（表C 条目）", "实测值（扫描证据）", "建议结论（确定值）", "备注 / 待补数据"]
    ws.append(headers)
    fill = PatternFill("solid", fgColor="DDEBF7")
    for i, c in enumerate(headers, 1):
        cell = ws.cell(1, i); cell.font = Font(bold=True); cell.fill = fill
        cell.alignment = Alignment(vertical="center")
    for r in rows:
        ws.append(r)
    for i, w in enumerate([9, 22, 12, 46, 20, 60, 16, 52], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    # 结论底色
    color = {"不符合": "FFC7CE", "部分符合": "FFEB9C", "符合": "C6EFCE", "不适用": "E2E2E2", "未执行": "FCE4D6", "待人工": "FFF2CC"}
    for row in ws.iter_rows(min_row=2, min_col=7, max_col=7):
        v = (row[0].value or "").split("（")[0]
        if v in color:
            row[0].fill = PatternFill("solid", fgColor=color[v])

    ws2 = wb.create_sheet("实测明细")
    ws2.append(["类别", "主机", "端口", "协议/组件或说明"])
    for i, c in enumerate(["类别", "主机", "端口", "协议/组件或说明"], 1):
        ws2.cell(1, i).font = Font(bold=True); ws2.cell(1, i).fill = fill
    for h in e.alives: ws2.append(["存活主机", h, "", ""])
    for h, p, pr in e.ports:
        tag = "高危端口" if p in RISK_PORTS else ("管理口" if p in MGMT_PORTS else "开放端口")
        ws2.append([tag, h, p, pr])
    for u, c, d in e.vulns: ws2.append(["漏洞(PoC)", u, c, d])
    if e.dirscan:
        for d in e.dirscan:
            ws2.append(["dirsearch命中", d.get("Url", ""), str(d.get("Code", "")), f"{d.get('Title','')} {d.get('App','')}".strip()])
    for h, p, m in e.cracks: ws2.append(["弱口令命中（不记明文）", h, p, f"模式:{m}"])
    for i, w in enumerate([18, 46, 12, 70], 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws2.freeze_panes = "A2"

    ws3 = wb.create_sheet("判定口径")
    ws3.append(["项目", "口径"])
    ws3.cell(1, 1).font = Font(bold=True); ws3.cell(1, 2).font = Font(bold=True)
    gen = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    meta = [
        ("生成时间", gen),
        ("数据来源", "golin xlsx（四表）+ html 任务开关 + 测评项与命令对应表 + 确定值策略包自动发现的台账/dirScan.json/采集输出"),
        ("结论口径", "符合 / 部分符合 / 不符合 / 不适用 / 未执行 —— 只出确定值；扫描侧不能定论的加（仅扫描侧）后缀，缺失证据写明还差什么、放哪个目录会自动转值"),
        ("策略 P-SCOPE", "受检目标全部 RFC1918 私网 → 公网依赖项判 不适用（范围含公网加 --public）"),
        ("策略 P-LEDGER", "台账来源：%s；台账外资产→不符合、无差集→符合、台账未随案→未执行" % (extra_meta.get("ledger_src") or "未发现（放 客户材料/ 或 --ledger 指定即自动生效）")),
        ("策略 P-EOL", "命中→不符合、未命中→符合、无指纹→未执行；规则库仅收官方停止支持日期的版本族，无版本号指纹不判定"),
        ("策略 P-PATH", "目录类项 6.2-1/6.4-2/6.5-1/6.5-2：PoC+dirsearch 双源关键词命中→不符合；dirsearch 已跑无命中→符合；未跑→未执行（dirScan.json 放 扫描输出/ 即自动转值）。关键词：%s" %
            "；".join(f"{k}={','.join(v[:4])}" for k, v in DIR_KW.items())),
        ("策略 P-ALIGN 证据-主题对齐", "PoC 按五类归集：defaultpass 默认口令 / errorpage 报错页 / mondebug 监控调试(pprof·metrics) / mgmtpage 管理页(swagger·console) / unauth_other 其余未授权。"
            "每类测评项只吃自己主题的类别：默认口令项 6.2-3·6.6-3·6.5-3 = defaultpass PoC + 口令试算（未授权类不计入，未授权≠默认口令）；报错页 6.4-1 = errorpage；"
            "管理页 6.2-2 = mgmtpage；接口 6.6-2 = 未授权家族；痕迹 2.3-4·2.7-3 = 全类别。类别 0 命中即按该主题判 符合/未执行。"),
        ("策略 P-COLLECT", "采集类项查 采集输出/**/测评项_<编号>.txt：缺失→未执行、已就位→引用文件待判读"),
        ("dirsearch 来源", (extra_meta.get("dirscan_src") or "未发现 dirScan.json")
            + (f"；覆盖率清单 {extra_meta.get('dirmanifest_src')}（{dir_coverage(e)[0]}/{dir_coverage(e)[1]} 站）"
               if dir_coverage(e) else "；覆盖率清单未发现（跑 BRscan-dirsearch.py 会生成）")
            + (f"；通配响应站 {e.wildcard_n} 个（soft-404，已剔除 {e.dirscan_excluded} 条误报，该部分记（仅扫描侧））"
               if e.wildcard_n else "")),
        ("EOL 规则库", "\n".join(f"· {n}：{pat} → {d}" for n, pat, d in EOL_RULES)),
        ("高危端口集合", ", ".join(sorted(RISK_PORTS, key=lambda x: int(x)))),
        ("管理口集合", ", ".join(sorted(MGMT_PORTS, key=lambda x: int(x)))),
        ("条目规则", "；".join(f"表C 第 {k} 条→{fn.__name__}" for k, fn in sorted(RULES.items(), key=lambda x: x[0]) if fn)),
        ("测评项级策略", "；".join(f"{k}→{fn.__name__}" for k, fn in ITEM_RULE.items())),
        ("覆盖", f"清单 166 条中含表C(golin)引用的 {n_items} 条已逐一给出确定值建议结果；其余为访谈/表B 上机项，不归扫描报告管"),
        ("弱口令纪律", "只记命中数量与服务类型，不抄用户/密码明文（表C 执行纪律④）"),
    ]
    for k, v in meta:
        ws3.append([k, v])
    ws3.column_dimensions["A"].width = 16
    ws3.column_dimensions["B"].width = 110
    for row in ws3.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(path)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", help="golin 扫描报告 xlsx")
    ap.add_argument("--html", default=None)
    ap.add_argument("--map", dest="mapx", default=None)
    ap.add_argument("--ledger", default=None)
    ap.add_argument("--collect", default=None)
    ap.add_argument("--dirsearch", default=None)
    ap.add_argument("--public", action="store_true")
    ap.add_argument("--no-policy", action="store_true")
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    xdir = os.path.dirname(os.path.abspath(args.xlsx))
    eroot = os.path.dirname(xdir) if os.path.basename(xdir) in ("扫描输出", "采集输出") else xdir
    roots = [xdir, eroot]
    policy = not args.no_policy

    html = args.html
    if html is None:
        stem = os.path.basename(args.xlsx).replace("-GoLin.xlsx", "")
        cand = [f for f in os.listdir(xdir) if f.endswith(".html") and f.startswith(stem)]
        html = os.path.join(xdir, cand[0]) if cand else None
    mpath = args.mapx
    if mpath is None:
        base = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        for c in (os.path.join(base, "内部支撑", "信息安全检查服务·测评项与命令对应表.xlsx"),
                  os.path.join(eroot, "内部支撑", "信息安全检查服务·测评项与命令对应表.xlsx")):
            if os.path.exists(c):
                mpath = c; break
    if not mpath or not os.path.exists(mpath):
        sys.exit("找不到《测评项与命令对应表.xlsx》，用 --map 指定")

    d = load_golin(args.xlsx)
    crack_on, poc_on = load_flags(html)
    items, tabc = load_map(mpath)

    # 台账 / dirsearch / 采集 自动发现
    ledger, ledger_src = find_ledger(args.ledger, roots) if policy else (None, None)
    if ledger is not None and len(ledger) == 0:
        ledger, ledger_src = None, None
    dirscan, dirscan_src = load_dirscan(args.dirsearch, roots) if policy else (None, None)
    dirmanifest, dirmanifest_src = find_dirmanifest(roots) if policy else (None, None)
    wildcard_n, dirscan_excluded = 0, 0
    if dirscan and dirmanifest:
        wild = dirmanifest.get("wildcard_targets") or [t["url"] for t in dirmanifest.get("targets", []) if t.get("wildcard")]
        wildcard_n = len(wild)
        if wild:
            kept = [e for e in dirscan if not any(str(e.get("Url", "")).startswith(w) for w in wild)]
            dirscan_excluded = len(dirscan) - len(kept)
            dirscan = kept
        if not dirscan_excluded and dirmanifest:
            dirscan_excluded = int(dirmanifest.get("pruned_total") or 0)   # 驱动侧已剔除的累计数
    golin_ids = [it["id"] for it in items if parse_src(it["src"])[0]]
    collect_files = find_collect_files(args.collect, roots, golin_ids) if policy else {}

    hosts = set(d["alives"]) | {h for h, _, _ in d["ports"]}
    def is_private(ip):
        try: return ipaddress.ip_address(ip).is_private
        except ValueError: return False
    internal = bool(hosts) and all(is_private(h) and not ipaddress.ip_address(h).is_global for h in hosts)

    e = Evidence(d, crack_on, poc_on, ledger, ledger_src, dirscan, dirscan_src,
                 collect_files, internal, args.public, policy, dirmanifest=dirmanifest,
                 wildcard_n=wildcard_n, dirscan_excluded=dirscan_excluded)

    print(f"golin报告: {os.path.basename(args.xlsx)}")
    print(f"html开关 : 口令扫描={crack_on} 漏洞扫描={poc_on} ({os.path.basename(html) if html else '未找到'})")
    print(f"数据     : 存活 {len(e.alives)} 台 | 开放 {len(e.ports)} 口 | PoC {len(e.vulns)} | 弱口令 {len(e.cracks)} | 指纹 {len(e.fingerprints())}")
    print(f"策略包   : {'开' if policy else '关'} | 范围={'全内网→公网项不适用' if internal and not args.public else '含公网/--public'} | "
          f"台账={ledger_src or '未发现'} | dirsearch={dirscan_src or '未发现'}"
          + (f"[覆盖 {dirmanifest['targets'] and len(dirmanifest.get('targets', []))}/{dirmanifest.get('targets_total', '?')} 站]"
             if dirmanifest else "") + f" | 采集文件={sum(1 for v in collect_files.values() if v)} 项")
    print("-" * 88)

    rows, n_golin = [], 0
    for it in items:
        cid, bid = parse_src(it["src"])
        if not cid:
            continue
        n_golin += 1
        实测, 结论, 备注, cid, bid = evaluate_item(it, e, tabc)
        tabctxt = "、".join(f"第 {c} 条" for c in cid) or "—"
        rows.append([it["id"], it["domain"], it["level"], it["text"], f"表C {tabctxt}", 实测, 结论, 备注])
        print(f"  {it['id']:<8} {结论:<14} {it['text'][:32]}")

    from collections import Counter
    tally = Counter((r[6] or "").split("（")[0] for r in rows)
    print("-" * 88)
    print("结论分布: " + " | ".join(f"{k} {v}" for k, v in sorted(tally.items(), key=lambda x: -x[1])))
    if tally.get("待人工"):
        print(f"⚠ 仍有 {tally['待人工']} 条待人工（见备注里的「还差什么」）")
    out = args.out or os.path.join(xdir, f"BRscan-assess_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    write_out(out, rows, e, tabc, n_golin,
              {"ledger_src": ledger_src, "dirscan_src": dirscan_src,
               "dirmanifest_src": dirmanifest_src})
    print(f"覆盖 {n_golin} 条 golin 归口测评项 → {out}")

if __name__ == "__main__":
    main()
