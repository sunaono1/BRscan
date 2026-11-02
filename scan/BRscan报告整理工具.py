#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""BRscan报告整理工具 —— golin 扫描报告读取与整理（报告读取模块独立版）

干什么:
    读 golin 扫描报告 xlsx（+同名 html 的扫描开关）→ 一键整理成结构化摘要（markdown）:
    概览统计 / 扫描开关 / 存活主机 / 端口按主机 / 高危与管理端口提示 / Web 站点 /
    PoC 命中 / 弱口令统计 / 组件指纹清单。
    报告读取模块（load_golin / load_flags）与 BRscan-assess.py 同源，可单独 import 复用。

用法:
    python3 BRscan报告整理工具.py <golin报告.xlsx> [选项]

选项:
    --html PATH   指定同名 html（默认自动找同目录同前缀 *.html，读取口令/漏洞扫描开关）
    --out PATH    摘要 markdown 输出路径（默认 <报告同目录>/报告整理_<时间戳>.md）
    --top N       每类明细最多列 N 行（默认 20，超出截断并标注总数）

输出示例:
    python3 BRscan报告整理工具.py 扫描输出/信息安全检查_全端口-GoLin.xlsx
    python3 BRscan报告整理工具.py 报告.xlsx --out 摘要.md --top 50

纪律: 弱口令只记命中数量与服务类型/模式，不抄用户/密码明文。
"""
import argparse
import datetime
import os
import re
import sys

import openpyxl

# ================================ 报告读取模块 ================================
# （自 BRscan-assess.py 抽出的独立读取层：golin 报告 xlsx + html 开关）

def load_golin(xlsx_path):
    """读 golin 扫描报告 xlsx → {"ports", "alives", "vulns", "cracks"}。
    弱口令资产表列: 序号 主机 端口 用户 密码 模式 —— 只读 主机/端口/模式，不读 用户/密码。"""
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
    if "弱口令资产" in wb.sheetnames:
        for r in list(wb["弱口令资产"].iter_rows(values_only=True))[1:]:
            if r and r[1]:
                data["cracks"].append((str(r[1]).strip(), str(r[2]).strip(), (r[5] or "").strip()))
    wb.close()
    return data


def load_flags(html_path):
    """读 golin 同名 html 的扫描开关 → (口令扫描开, 漏洞扫描开)，缺失/未标注为 None。"""
    if not html_path or not os.path.exists(html_path):
        return None, None
    raw = open(html_path, encoding="utf-8", errors="replace").read()

    def flag(pat):
        m = re.search(pat + r"[^>]*>\s*(是|否)", raw)
        return (m.group(1) == "是") if m else None

    return flag(r"启用口令扫描"), flag(r"启用漏洞扫描")


# ================================ 整理口径 ================================

# 高危/敏感服务端口 → 提示语（整理提示用，非测评结论）
RISK_PORTS = {
    "21": "FTP", "22": "SSH", "23": "Telnet（明文）", "25": "SMTP", "69": "TFTP",
    "110": "POP3", "135": "RPC", "139": "NetBIOS", "143": "IMAP", "161": "SNMP",
    "389": "LDAP", "445": "SMB", "512/513/514": "rexec/rlogin/rsh", "873": "rsync",
    "1433": "MSSQL", "1521": "Oracle", "2049": "NFS", "3306": "MySQL", "3389": "RDP",
    "5900": "VNC", "6379": "Redis", "9200": "Elasticsearch", "11211": "Memcached",
    "27017": "MongoDB", "623": "IPMI", "61616": "ActiveMQ",
}
MGMT_PORTS = {"623", "5900", "5988", "5989"}
CLEARTEXT_MGMT = {"21", "23", "161", "514"}


def _fmt(v):
    return "是" if v else ("否" if v is False else "未标注")


def _list(items, top, fmt=lambda x: str(x)):
    shown = "、".join(fmt(i) for i in items[:top])
    if len(items) > top:
        shown += f" ……（等共 {len(items)} 项）"
    return shown or "无"


def organize(d, crack_on, poc_on, html_name, top):
    """把读取结果整理成 markdown 段落列表。"""
    alives, ports, vulns, cracks = d["alives"], d["ports"], d["vulns"], d["cracks"]
    hosts = sorted(set(alives) | {h for h, _, _ in ports})
    by_host = {}
    for h, p, pr in ports:
        by_host.setdefault(h, []).append((p, pr))
    web = [(h, p, pr) for h, p, pr in ports if "WEB" in (pr or "").upper()]
    fingerprints = []
    for _, _, pr in ports:
        if pr and pr not in fingerprints:
            fingerprints.append(pr)
    risk = sorted({(h, p) for h, p, _ in ports if p in RISK_PORTS},
                  key=lambda x: (x[0], int(x[1])))
    clear = sorted({(h, p) for h, p, _ in ports if p in CLEARTEXT_MGMT},
                   key=lambda x: (x[0], int(x[1])))
    mgmt = sorted({(h, p) for h, p, _ in ports if p in MGMT_PORTS},
                  key=lambda x: (x[0], int(x[1])))

    md = []
    md.append("# golin 扫描报告整理摘要\n")
    md.append(f"整理时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}　"
              f"整理工具：BRscan报告整理工具\n")

    # ---- 1 概览 ----
    md.append("## 1. 概览统计\n")
    md.append("| 指标 | 数值 |")
    md.append("|---|---|")
    md.append(f"| 存活主机 | {len(set(alives))} 台 |")
    md.append(f"| 开放端口 | {len(ports)} 个（涉及 {len(by_host)} 台主机） |")
    md.append(f"| PoC 命中 | {len(vulns)} 条 |")
    md.append(f"| 弱口令命中 | {len(cracks)} 处（明细不载明文） |")
    md.append(f"| 组件指纹 | {len(fingerprints)} 种 |")
    md.append("")

    # ---- 2 扫描开关 ----
    md.append("## 2. 扫描开关\n")
    md.append(f"- 启用口令扫描：{_fmt(crack_on)}"
              f"（来源：{html_name or '未找到同名 html'}）")
    md.append(f"- 启用漏洞扫描：{_fmt(poc_on)}")
    if crack_on is False and cracks:
        md.append("- ⚠ 开关显示未启用口令扫描但报告含弱口令条目，请人工核对数据来源")
    if crack_on is False and not cracks:
        md.append("- 注：未启用口令扫描，弱口令侧无覆盖不代表无弱口令")
    md.append("")

    # ---- 3 存活主机 ----
    md.append("## 3. 存活主机\n")
    md.append(_list(sorted(set(alives)), top))
    md.append("")

    # ---- 4 端口按主机 ----
    md.append("## 4. 开放端口（按主机）\n")
    md.append("| 主机 | 端口数 | 端口 / 指纹 |")
    md.append("|---|---|---|")
    for h in sorted(by_host, key=lambda x: (len(by_host[x]), x), reverse=True)[:top]:
        plist = "、".join(f"{p}{('（' + pr + '）') if pr else ''}" for p, pr in sorted(by_host[h], key=lambda t: int(t[0])))
        md.append(f"| {h} | {len(by_host[h])} | {plist} |")
    if len(by_host) > top:
        md.append(f"\n（仅列端口数前 {top} 台，共 {len(by_host)} 台）")
    md.append("")

    # ---- 5 高危与管理端口提示 ----
    md.append("## 5. 高危与管理端口提示（整理提示，非测评结论）\n")
    md.append(f"- 高危/敏感服务命中 {len(risk)} 处：" + _list(
        [f"{h}:{p}（{RISK_PORTS[p]}）" for h, p in risk], top))
    md.append(f"- 明文管理协议 {len(clear)} 处：" + _list(
        [f"{h}:{p}" for h, p in clear], top))
    md.append(f"- 带外管理口（IPMI/VNC 等）{len(mgmt)} 处：" + _list(
        [f"{h}:{p}" for h, p in mgmt], top))
    md.append("")

    # ---- 6 Web 站点 ----
    md.append("## 6. Web 站点清单\n")
    md.append(_list([f"{h}:{p}{('（' + pr + '）') if pr else ''}" for h, p, pr in web], top))
    md.append("")

    # ---- 7 PoC 命中 ----
    md.append("## 7. PoC 命中\n")
    if vulns:
        md.append("| # | URL | PoC | 描述 |")
        md.append("|---|---|---|---|")
        for i, (url, name, desc) in enumerate(vulns[:top], 1):
            md.append(f"| {i} | {url} | {name} | {(desc or '').replace('|', '／')} |")
        if len(vulns) > top:
            md.append(f"\n（仅列前 {top} 条，共 {len(vulns)} 条）")
    else:
        md.append("无（或本次未启用漏洞扫描）")
    md.append("")

    # ---- 8 弱口令统计（不抄明文） ----
    md.append("## 8. 弱口令统计（只记数量与服务类型，不载用户/密码明文）\n")
    if cracks:
        stat = {}
        for _, p, mode in cracks:
            key = f"{p}{'（' + mode + '）' if mode else ''}"
            stat[key] = stat.get(key, 0) + 1
        md.append("| 服务（端口/模式） | 命中数 |")
        md.append("|---|---|")
        for k in sorted(stat, key=lambda x: -stat[x]):
            md.append(f"| {k} | {stat[k]} |")
        md.append(f"\n合计 {len(cracks)} 处；逐条主机分布：" + _list(
            sorted({h for h, _, _ in cracks}), top))
    else:
        md.append("无命中（或本次未启用口令扫描）")
    md.append("")

    # ---- 9 组件指纹 ----
    md.append("## 9. 组件指纹清单\n")
    md.append(_list(fingerprints, top))
    md.append("")

    md.append("---")
    md.append("*注：本摘要为工具读取整理结果，命中 ≠ 最终判定，需人工复核后方可入核查表与报告。*")
    return md


def main():
    ap = argparse.ArgumentParser(description="BRscan报告整理工具 —— golin 扫描报告读取与整理")
    ap.add_argument("xlsx", help="golin 扫描报告 xlsx")
    ap.add_argument("--html", default=None, help="同名 html（默认自动查找）")
    ap.add_argument("--out", default=None, help="摘要 markdown 输出路径")
    ap.add_argument("--top", type=int, default=20, help="每类明细最多列 N 行（默认 20）")
    args = ap.parse_args()

    if not os.path.isfile(args.xlsx):
        sys.exit(f"找不到报告文件: {args.xlsx}")
    xdir = os.path.dirname(os.path.abspath(args.xlsx))

    html = args.html
    if html is None:
        stem = os.path.basename(args.xlsx).replace("-GoLin.xlsx", "")
        cand = [f for f in os.listdir(xdir) if f.endswith(".html") and f.startswith(stem)]
        html = os.path.join(xdir, cand[0]) if cand else None

    d = load_golin(args.xlsx)
    crack_on, poc_on = load_flags(html)

    print(f"报告: {os.path.basename(args.xlsx)}")
    print(f"开关: 口令扫描={_fmt(crack_on)} 漏洞扫描={_fmt(poc_on)}"
          f"（{os.path.basename(html) if html else 'html 未找到'}）")
    print(f"读取: 存活 {len(set(d['alives']))} 台 | 开放 {len(d['ports'])} 口 | "
          f"PoC {len(d['vulns'])} | 弱口令 {len(d['cracks'])} 处")
    print("-" * 60)

    md = organize(d, crack_on, poc_on,
                  os.path.basename(html) if html else None, args.top)
    out = args.out or os.path.join(
        xdir, f"报告整理_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.md")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(md) + "\n")
    print(f"摘要已生成: {out}")


if __name__ == "__main__":
    main()
