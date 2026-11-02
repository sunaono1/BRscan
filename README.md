# BRscan

内网安全检查工具集：**扫描工具**（全端口扫描提速、存活预筛、目录探测接续、测评结果自动梳理）+ **报告工具**（检查报告 / 授权书 / 调研表 docx 一键生成）+ **主机只读采集脚本**。

从「扫描 → 目录探测 → 测评结论 → 检查报告」整条链路都有自动化脚本支撑，核心目标：**现场少返工、结论可溯源、报告一键出稿**。

## 目录结构

```
BRscan/
├── scan/                        # 扫描工具
│   ├── BRscan-fastscan.sh       # golin 全端口扫描提速封装（--time 2 --chan 3.0 --done 2）
│   ├── BRscan-alive-probe.py    # 扫描前存活预筛（ping + TCP 双重复核）
│   ├── BRscan-dirsearch.py      # 目录探测接续驱动（随机探针拦 soft-404 / 原子写回去重 / 双轮补漏）
│   ├── BRscan-assess.py         # 扫描报告 → 测评结论表 xlsx（确定值策略包）
│   ├── BRscan报告整理工具.py     # 报告读取模块独立版：golin 报告 → 结构化摘要 md
│   └── dict/
│       ├── dir-dict.txt         # 目录探测精编字典（400 条，四主题）
│       └── weakpass-dict.txt    # 弱口令合并字典（1203 条）
├── report/                      # 报告工具（python-docx 生成器）
│   ├── BRscan-gen-report.py     # 检查报告生成器（模板 + 梳理结果 → docx）
│   ├── BRscan-gen-auth.py       # 授权书生成器
│   ├── BRscan-gen-survey.py     # 现场调研表生成器
│   ├── BRscan-set-headers.py    # 批量统一页眉
│   └── BRscan-fix-header.py     # 页眉两栏宽度修正
├── collect/                     # 主机只读采集（按测评项编号双索引归档）
│   ├── BRscan-collect-linux.sh      # Linux：38 条目 / 76 命令（WITH_DB=1 可选数据库）
│   └── BRscan-collect-windows.ps1   # Windows：24 条目 / 40 命令
└── docs/
    └── 测速报告.md               # 全端口扫描提速的实测与调参定稿过程
```

## 扫描工具（scan/）

### BRscan-fastscan.sh —— 全端口扫描提速

只调 golin 命令行参数（工具本体不改）：

- `--time 2` 丢包端口等待 5s→2s（LAN/VPN 场景 RTT 的 ~90 倍，不丢精度）
- `--chan 3.0` 并发槽 166→624（log2(IP×端口)×CPU核数×chan）
- `--done 2` 单探测放弃上限 10min→2min

实测：整段全端口 4.4 小时 → **33.7 分钟（≈8×）**；对拍精度零漏检、241 个开放端口 0 误报。

```bash
# 需自行放置 golin 二进制（默认 ./golin，GOLIN_BIN 可覆盖）
./scan/BRscan-fastscan.sh 192.168.1.0/24              # 整段全端口
./scan/BRscan-fastscan.sh 192.168.1.51,192.168.1.52   # 指定主机（台账场景）
./scan/BRscan-fastscan.sh 192.168.1.0/24 --nocrack    # 未获口令核查书面同意时关闭弱口令
```

### BRscan-alive-probe.py —— 存活预筛

规避 `--noping` 全端口扫整段死网段的小时级陷阱（226 台死主机理论 12.3 小时）。ICMP ping 全段，不通的再 TCP 建连复核常用端口（RST/SYN-ACK 都算存活）。

```bash
python3 scan/BRscan-alive-probe.py 192.168.1.0/24 > alive.txt
./scan/BRscan-fastscan.sh $(paste -sd, alive.txt) --noping
```

注意：全静默主机（不回 ICMP 且 TCP 全程 DROP）无法与死主机区分；客户台账已给主机清单时优先直接用台账。

### BRscan-dirsearch.py —— 目录探测接续

读 golin 扫描报告的站点清单 → 逐站跑 dirsearch → `dirScan.json` 回灌测评脚本。三道可靠性措施：

1. **随机探针校验**：每站先发 3 条随机路径做探针，随机路径也命中 = 通配响应（soft-404），标记跳过不进证据；
2. **原子写回去重**：每站跑完立即解析 → 去重 → 原子写回，治 golin 并发写 `dirScan.json` 的竞态叠写/丢条；
3. **双轮补漏**：第 2 轮补回第 1 轮竞态丢掉的条目。

实测：136 站 **11.9 分钟**，自动剔除误报 **594 条**、有效命中 108。

```bash
python3 scan/BRscan-dirsearch.py <golin报告.xlsx>            # 默认字典 scan/dict/dir-dict.txt
python3 scan/BRscan-dirsearch.py <golin报告.xlsx> --passes 2  # 默认即双轮
```

### BRscan-assess.py —— 测评结果自动梳理

扫描报告 → 逐条测评结论表（xlsx），结论口径：符合 / 部分符合 / 不符合 / 不适用 / 未执行。内置**确定值策略包**：

| 策略 | 行为 |
|---|---|
| P-SCOPE | 受检目标全为 RFC1918 私网 → 公网依赖项自动判「不适用」 |
| P-LEDGER | 台账清单自动出差集：台账外资产→不符合、无差集→符合 |
| P-EOL | 内置 EOL 规则库比对组件指纹：命中停止支持版本→不符合 |
| P-PATH | 目录类项按 PoC/dirsearch 路径关键词定级 |
| P-COLLECT | 采集类项查采集输出，缺→未执行、有→引用待判读 |
| P-ALIGN | 证据-主题对齐：默认口令/报错页/管理页各归各的测评项 |

```bash
python3 scan/BRscan-assess.py <golin报告.xlsx> [--ledger 台账.txt] [--collect 采集输出/] [--no-policy]
```

纪律：弱口令只记命中数量与服务类型，不抄用户/密码明文。

### BRscan报告整理工具 —— 报告读取与整理

报告读取模块独立版（`load_golin` / `load_flags` 与 BRscan-assess.py 同源，可单独 import 复用）：读 golin 扫描报告 xlsx + 同名 html 的扫描开关，一键整理成结构化摘要 md——概览统计 / 扫描开关 / 存活主机 / 端口按主机 / 高危与管理端口提示 / Web 站点 / PoC 命中 / 弱口令统计 / 组件指纹，九段一稿出。适合先跑它看全局，再决定往哪深挖。

```bash
python3 scan/BRscan报告整理工具.py <golin报告.xlsx>            # 默认输出同目录 报告整理_<时间戳>.md
python3 scan/BRscan报告整理工具.py <golin报告.xlsx> --out 摘要.md --top 50   # 指定输出与每类明细行数
```

弱口令段只出「端口/模式 → 命中数」统计表，不抄用户/密码明文；摘要结尾自带「命中 ≠ 最终判定，需人工复核」注记。

## 报告工具（report/）

python-docx 生成器，模板 + 梳理结果 → 成品 docx：

```bash
python3 report/BRscan-gen-report.py <模板.docx> <测评结果梳理.xlsx> <输出.docx>
BRSCAN_ROOT=. python3 report/BRscan-gen-auth.py        # 授权书
BRSCAN_ROOT=. python3 report/BRscan-gen-survey.py      # 现场调研表
python3 report/BRscan-set-headers.py <目录>            # 统一页眉
```

- `BRscan-gen-report.py`：按点位填正文/表格，附录 A 按「编号→结论」映射回填，可插图（figs/）、可插「体系问题与改进建议」专章；封面/摘要/问题清单/整改建议全自动生成，人工只做校对。
- `BRscan-gen-auth.py` / `BRscan-gen-survey.py`：输出路径由 `BRSCAN_ROOT` 控制（默认当前目录）；调研表覆盖率统计依赖清单数据源 `content.py`（`BRSCAN_CONTENT` 可指定，缺失时跳过统计）。

## 主机只读采集（collect/）

```bash
sudo bash collect/BRscan-collect-linux.sh -o ./采集输出_$(date +%F)
WITH_DB=1 sudo -E bash collect/BRscan-collect-linux.sh   # 需要数据库命令时
```
```powershell
powershell -ExecutionPolicy Bypass -File .\collect\BRscan-collect-windows.ps1 -OutDir D:\采集输出
```

每个条目在输出头部和 `清单.txt` 标注对应测评项编号，并生成 `采集输出/测评项_<编号>.txt` 索引文件。

纪律：① 全程只读，不改配置、不装软件、不重启服务；② 命令由检查方提供，账号口令由管理员输入；③ 输出按条目号与测评项编号双索引归档。

## 依赖

- Python 3.9+：`openpyxl`、`python-docx`
- golin（扫描引擎，二进制不随仓库分发，`GOLIN_BIN` 指定路径）
- Linux / macOS（扫描与采集）；Windows（采集）

## 实测成效

| 环节 | 成效 |
|---|---|
| 全端口扫描 | 4.5 小时 → 35 分钟（约 8×），零漏报零误报 |
| 目录探测 | 136 站 12 分钟，594 条误报自动剔除 |
| 测评梳理 | 扫描结果自动出逐项结论，37 项工具类检查直接给确定值 |
| 主机采集 | 一键只读、按检查项自动归档取证 |
| 报告出稿 | 模板+数据一键成稿，人工只做校对 |

## License

MIT
