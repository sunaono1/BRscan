#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""BRscan-dirsearch.py —— 把 golin dirsearch 接进「golin 扫描 → 测评结果」链

干什么:
    读 golin 扫描报告 xlsx 的 WEB应用 站点清单 → 逐站执行 golin 自带 dirsearch
    （表C 第 10/11 条口径: --code 200,301,302,403,500 --chan 50 --timeout 3）→
    dirScan.json（golin 原生格式）+ dirscan_manifest.json（覆盖率/通配站清单）。
    跑完重跑 BRscan-assess.py，6.2-1 / 6.4-2 / 6.5-1 / 6.5-2 自动转确定值。

三道可靠性措施（治 golin dirsearch 三个坑）:
    ① 随机探针校验: 每站先用本脚本自己发 3 条随机路径；若随机路径也命中状态码集
       → 该站为通配响应(soft-404，如 MinIO/SPA 任意路径都回 403/200) → 标记跳过，
       其响应不进证据（否则目录类条目会被灌成假"不符合"）。
    ② 原子写回去重: golin 并发命中时读改写 dirScan.json 会竞态叠写/丢条（上游行为），
       每站跑完立即宽容解析 → 去重 → 原子回写（探针路径剔除）。
    ③ 双轮补漏: --passes 2（默认）整表扫两遍，第 2 轮补回第 1 轮竞态丢掉的条目。

用法:
    python3 BRscan-dirsearch.py <golin报告.xlsx> [选项]

选项:
    --dict PATH      字典（默认 dict/dir-dict.txt：400 条，按四个目录类条目主题精编）
    --builtin        用 golin 内置字典（31094 条/站，数分钟一站，全量 7 小时+，慎用）
    --code LIST      状态码过滤（默认 200,301,302,403,500，同表C 第 10 条）
    --chan N         每站并发（默认 50，同表C 第 10 条）
    --timeout N      单请求超时秒（默认 3，同 golin 默认）
    --passes N       整表轮数（默认 2，补竞态丢失；探针判通配站跳过不重跑）
    --max N          每轮最多跑 N 个站点（调试/分批）
    --include REG    只跑 URL 匹配正则的站点
    --force          不跳过清单已完成站（默认断点续跑；通配站始终跳过）
    --out-dir DIR    dirScan.json / 清单 / 日志落盘目录（默认 = 报告 xlsx 所在目录）
    --golin PATH     golin 可执行文件（默认脚本同目录 golin-1.2.9-darwin-arm64）
    --delay SEC      站点间隔（默认 0.3s）

纪律: 仅对本次授权受检范围内的 Web 站点做路径探测（表C 第 10/11 条），与扫描报告同范围。
"""
import argparse, json, os, re, ssl, subprocess, sys, time, datetime, urllib.request, random, string
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DICT = os.path.join(HERE, "字典", "目录探测字典.txt")
DEFAULT_GOLIN = os.path.join(HERE, "golin-1.2.9-darwin-arm64")
URL_RE = re.compile(r"(https?://\S+)")
CANARY_MARK = "/dsh-canary-"

# ---------------------------------------------------------------- 宽容解析 + 修复（共享逻辑）
def load_entries_flexible(path):
    """宽容读取 dirScan.json：兼容叠写/多文档/局部损坏；去重；剔除探针路径"""
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
        if CANARY_MARK in u:
            continue
        k = (u, e.get("Code"), (e.get("Title") or "").strip())
        if k in seen:
            continue
        seen.add(k); res.append(e)
    return res

def repair_dirscan(path):
    entries = load_entries_flexible(path)
    if entries or os.path.exists(path):
        tmp = path + ".tmp"
        json.dump(entries, open(tmp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        os.replace(tmp, path)
    return entries

def load_targets(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["扫描端口结果"]
    targets, seen = [], set()
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if not r or not r[1]:
            continue
        host, port, proto = str(r[1]).strip(), str(r[2]).strip(), (r[3] or "")
        if "WEB" not in proto.upper():
            continue
        m = URL_RE.search(proto)
        url = (m.group(1) if m else f"http://{host}:{port}").rstrip("/")
        if url not in seen:
            seen.add(url); targets.append(url)
    wb.close()
    return targets

def load_json(path, default):
    if os.path.exists(path):
        try:
            return json.load(open(path, encoding="utf-8"))
        except Exception:
            return default
    return default

class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *a, **k):
        return None

def _make_opener():
    ctx = ssl.create_default_context(); ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE
    return urllib.request.build_opener(_NoRedirect, urllib.request.HTTPSHandler(context=ctx))

def canary_check(url, codes, timeout):
    """本脚本自发 3 条随机路径 → (是否通配, 各响应, 是否无法判定)。
    注意: TLS 上下文必须在构造 opener 时传入（open() 不收 context 参数）。"""
    opener = _make_opener()
    got = []
    for _ in range(3):
        path = CANARY_MARK + "".join(random.choices(string.hexdigits.lower(), k=16))
        try:
            resp = opener.open(urllib.request.Request(url + path, headers={"User-Agent": "dirscan-canary"}),
                               timeout=timeout)
            got.append(getattr(resp, "status", 200))
        except urllib.error.HTTPError as he:
            got.append(he.code)
        except Exception as ex:
            got.append(f"ERR:{type(ex).__name__}")
    errs = [g for g in got if isinstance(g, str)]
    inconclusive = len(errs) == len(got)          # 三条全失败（TLS/连接问题）→ 无法判定
    wildcard = (not inconclusive) and any(str(g) in codes for g in got if not isinstance(g, str))
    return wildcard, got, inconclusive

def homo_weak(hits):
    """命中 ≥30 且 (Code, ContentType) 同一组合占 ≥95% → 高度同质（疑似通配，配合探针无法判定时兜底）"""
    if len(hits) < 30:
        return False
    from collections import Counter
    c = Counter((h.get("Code"), h.get("ContentType")) for h in hits)
    _, n = c.most_common(1)[0]
    return n / len(hits) >= 0.95

def prune_url(path, url):
    """从 dirScan.json 里剔除某站点的全部条目（通配站响应不进证据），返回剔除数"""
    entries = load_entries_flexible(path)
    kept = [e for e in entries if not str(e.get("Url", "")).startswith(url)]
    removed = len(entries) - len(kept)
    json.dump(kept, open(path + ".tmp", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    os.replace(path + ".tmp", path)
    return removed

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", help="golin 扫描报告 xlsx")
    ap.add_argument("--dict", dest="dict_path", default=DEFAULT_DICT)
    ap.add_argument("--builtin", action="store_true")
    ap.add_argument("--code", default="200,301,302,403,500")
    ap.add_argument("--chan", type=int, default=50)
    ap.add_argument("--timeout", type=int, default=3)
    ap.add_argument("--passes", type=int, default=2)
    ap.add_argument("--max", type=int, default=0)
    ap.add_argument("--include", default=None)
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--recanary", action="store_true",
                    help="只对清单里已判非通配的站重跑探针（修复旧 bug 漏判），命中通配则原地剔除其条目后退出")
    ap.add_argument("--out-dir", default=None)
    ap.add_argument("--golin", default=DEFAULT_GOLIN)
    ap.add_argument("--delay", type=float, default=0.3)
    args = ap.parse_args()

    xlsx = os.path.abspath(args.xlsx)
    out_dir = os.path.abspath(args.out_dir) if args.out_dir else os.path.dirname(xlsx)
    os.makedirs(out_dir, exist_ok=True)
    if not os.path.isfile(args.golin):
        sys.exit(f"找不到 golin: {args.golin}")
    if not args.builtin and not os.path.isfile(args.dict_path):
        sys.exit(f"找不到字典: {args.dict_path}")

    codes = [c.strip() for c in args.code.split(",") if c.strip()]
    targets = load_targets(xlsx)
    if args.include:
        targets = [t for t in targets if re.search(args.include, t)]
    manifest_path = os.path.join(out_dir, "dirscan_manifest.json")
    dirscan_path = os.path.join(out_dir, "dirScan.json")
    dict_paths = 0 if args.builtin else sum(1 for l in open(args.dict_path, encoding="utf-8") if l.strip())

    # ---- --recanary: 原地修复旧版探针漏判（HTTPS context bug），不跑字典
    if args.recanary:
        manifest = load_json(manifest_path, {"targets": []})
        recs = {t["url"]: t for t in manifest.get("targets", [])}
        flagged = pruned_total = inc_n = 0
        print(f"[--recanary] 复检清单 {len(recs)} 站探针（修复 HTTPS 探针静默失败）…")
        for url in [u for u, r in recs.items() if not r.get("wildcard")]:
            wild, got, inc = canary_check(url, codes, args.timeout)
            r = recs[url]
            if inc:
                inc_n += 1
                r["canary_inconclusive"] = True
                entries = load_entries_flexible(dirscan_path)
                hits = [e for e in entries if str(e.get("Url", "")).startswith(url)]
                if homo_weak(hits):
                    wild, got = True, list(got) + ["homo-backstop"]
            if wild:
                r.update({"wildcard": True, "canary": got})
                n = prune_url(dirscan_path, url)
                pruned_total += n; flagged += 1
                print(f"  ⚠ {url} → 通配 {got}，剔除 {n} 条")
            else:
                r["canary_last"] = got
        manifest["targets"] = list(recs.values())
        manifest["wildcard_targets"] = sorted(u for u, r in recs.items() if r.get("wildcard"))
        manifest["pruned_total"] = int(manifest.get("pruned_total") or 0) + pruned_total
        manifest["updated"] = datetime.datetime.now().isoformat(timespec="seconds")
        json.dump(manifest, open(manifest_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print(f"[--recanary 完成] 新判通配 {flagged} 站，剔除 {pruned_total} 条；探针无法判定 {inc_n} 站（后续跑字典时启用同质兜底）")
        print("下一步: 重跑 BRscan-assess.py")
        return


    print(f"Web 站点: 共 {len(targets)} 个 | 轮数 {args.passes} | 探针通配站跳过 | 断点续跑 {'关(--force)' if args.force else '开'}")
    print(f"字典    : {'golin 内置 31094 条（--builtin）' if args.builtin else f'{args.dict_path}（{dict_paths} 条）'}"
          f" | 状态码 {args.code} | 并发 {args.chan} | 超时 {args.timeout}s")
    print(f"落盘    : {out_dir}/dirScan.json + dirscan_manifest.json")
    print("-" * 88)

    log_path = os.path.join(out_dir, f"目录探测_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    log = open(log_path, "a", encoding="utf-8")
    t_all = time.time()
    processed = wildcard_n = 0

    for p in range(1, args.passes + 1):
        manifest = load_json(manifest_path, {"targets": []})
        recs = {t["url"]: t for t in manifest.get("targets", [])}
        todo = []
        for t in targets:
            r = recs.get(t)
            if r and r.get("wildcard"):
                continue                      # 通配站: 已定性, 轮轮跳过
            if r and r.get("pass") and not args.force and p > 1 and r.get("done_pass2"):
                continue                      # 第2轮只为补漏: 第1轮已完成的照跑(补竞态丢条)
            todo.append(t)
        # 第2轮语义: 对所有非通配站再跑一遍(补丢条); --max 仍生效
        if args.max:
            todo = todo[:args.max]
        if not todo:
            print(f"[第 {p}/{args.passes} 轮] 无需处理的站点")
            break
        print(f"[第 {p}/{args.passes} 轮] 待跑 {len(todo)} 站 …")
        for i, url in enumerate(todo, 1):
            # ---- ① 随机探针校验（本脚本自发, 不经过 golin, 不写任何文件）
            wild, can_codes, inc = canary_check(url, codes, args.timeout)
            manifest = load_json(manifest_path, {"targets": []})
            recs = {t["url"]: t for t in manifest.get("targets", [])}
            rec = recs.get(url, {"url": url})

            if wild:
                wildcard_n += 1
                rec.update({"url": url, "hits": 0, "wildcard": True, "canary": can_codes,
                            "canary_inconclusive": False,
                            "pass": p, "secs": 0.0, "at": datetime.datetime.now().isoformat(timespec="seconds")})
                recs[url] = rec
                print(f"[{i}/{len(todo)}] {url}  ⚠ 通配响应(soft-404) 探针={can_codes} → 标记跳过，响应不进证据")
            else:
                cmd = [args.golin, "dirsearch", "--url", url, "--code", args.code,
                       "--chan", str(args.chan), "--timeout", str(args.timeout)]
                if not args.builtin:
                    cmd += ["--file", os.path.abspath(args.dict_path)]
                t0 = time.time()
                log.write(f"\n===== [pass {p}] [{i}/{len(todo)}] {datetime.datetime.now().isoformat()} {url}\n")
                log.flush()
                with open(log_path, "a", encoding="utf-8") as sink:
                    proc = subprocess.run(cmd, cwd=out_dir, stdout=sink, stderr=subprocess.STDOUT)
                secs = time.time() - t0
                # ---- ② 原子写回去重: 宽容解析 → 去重 → 原子回写
                entries = repair_dirscan(dirscan_path)
                hits_here = [e for e in entries if str(e.get("Url", "")).startswith(url)]
                # ---- ②b 同质兜底: 探针无法判定(TLS/连不通) 且 命中≥30条且≥95%同质 → 判通配并剔除
                if (not wild) and inc and homo_weak(hits_here):
                    wild = True
                    can_codes = list(can_codes) + ["homo-backstop"]
                    n = prune_url(dirscan_path, url)
                    manifest["pruned_total"] = int(manifest.get("pruned_total") or 0) + n
                    hits_here = [e for e in load_entries_flexible(dirscan_path)
                                 if str(e.get("Url", "")).startswith(url)]
                    print(f"    ↳ 同质兜底: 探针无法判定且命中高度同质 → 剔除 {n} 条")
                prev_hits = rec.get("hits") or 0
                rec.update({"url": url, "hits": len(hits_here), "wildcard": wild,
                            "canary": can_codes, "canary_inconclusive": inc,
                            "pass": p, "secs": round(secs, 1), "rc": proc.returncode,
                            "delta": max(0, len(hits_here) - prev_hits),
                            "at": datetime.datetime.now().isoformat(timespec="seconds")})
                recs[url] = rec
                processed += 1
                if wild:
                    wildcard_n += 1
                    print(f"[{i}/{len(todo)}] {url}  ⚠ 事后判通配（{can_codes}），其响应不进证据")
                else:
                    print(f"[{i}/{len(todo)}] {url}  命中 {len(hits_here)}（本轮新增 {rec['delta']}）  {secs:.1f}s  rc={proc.returncode}")

            manifest["targets"] = list(recs.values())
            manifest.update({"source_xlsx": os.path.basename(xlsx), "code": args.code,
                             "chan": args.chan, "dict": "builtin-31094" if args.builtin else dict_paths,
                             "targets_total": len(targets),
                             "wildcard_targets": sorted(t for t, r in recs.items() if r.get("wildcard")),
                             "passes": args.passes,
                             "updated": datetime.datetime.now().isoformat(timespec="seconds")})
            json.dump(manifest, open(manifest_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
            time.sleep(args.delay)
        # ---- 轮末汇总
        entries = repair_dirscan(dirscan_path)
        print(f"[第 {p}/{args.passes} 轮结束] dirScan.json 累计有效命中 {len(entries)} 条 | 通配站 {sum(1 for r in recs.values() if r.get('wildcard'))} 个")

    log.close()
    entries = load_entries_flexible(dirscan_path)
    manifest = load_json(manifest_path, {"targets": []})
    wild = manifest.get("wildcard_targets", [])
    scanned = sum(1 for r in manifest.get("targets", []) if not r.get("wildcard"))
    print("-" * 88)
    print(f"完成: 用时 {(time.time()-t_all)/60:.1f} 分钟 | 正常站 {scanned}/{len(targets)} | 通配跳过 {len(wild)} | 有效命中 {len(entries)} 条")
    if wild:
        print(f"通配响应站（soft-404，无法枚举，测评侧只作（仅扫描侧））: {', '.join(wild)}")
    print("下一步: 重跑 BRscan-assess.py，目录类条目自动转确定值")

if __name__ == "__main__":
    main()
