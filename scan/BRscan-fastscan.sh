#!/bin/bash
# BRscan-fastscan.sh —— golin 全端口扫描快速执行脚本（golin 二进制原样使用，不改）
#
# 用法:
#   ./BRscan-fastscan.sh 192.168.1.0/24                      # 整段全端口
#   ./BRscan-fastscan.sh 192.168.1.51,192.168.1.52,192.168.1.110   # 指定主机（推荐台账场景）
#   ./BRscan-fastscan.sh SEGMENT/24 --noimg                    # 追加参数透传给 golin
#
# 提速三参数（2026-10-19 在 192.168.1.0/24 实测定稿，详见 docs/测速报告.md）:
#   --time 2   丢包端口等待 5s→2s（本段 RTT 22ms，2s=RTT 的 ~90 倍，LAN/VPN 不丢精度）
#   --chan 3.0 并发槽 166→624（log2(IP×端口)×CPU核数×chan；临时端口余量充足）
#   --done 2   单探测放弃上限 10min→2min（正常探测最坏 ~15s，2min 是异常兜底）
#
# 默认全链路（弱口令+PoC+截图）；按纪律域环境或未获书面同意时追加 --nocrack 关闭弱口令。
# 弱口令字典: 自动使用本目录 dict/weakpass-dict.txt（1203 条 = weak_top1000 + 厂商默认 50
# + webadmin + golin 内置，由公开弱口令 Top1000 + 常见厂商默认口令合并而成）；无该文件则回退 golin 内置。
set -o pipefail
cd "$(dirname "$0")"
TOOL="${GOLIN_BIN:-./golin}"   # golin 二进制路径，可用环境变量 GOLIN_BIN 覆盖

TARGET="$1"; shift || true
if [ -z "$TARGET" ]; then echo "用法: $0 <IP|网段|IP列表> [golin scan 追加参数...]"; exit 1; fi

EXTRA=()
[ -f "dict/weakpass-dict.txt" ] && EXTRA+=(--passwdfile "dict/weakpass-dict.txt")

outname="信息安全检查_$(date +%Y-%m-%d_%H%M)_全端口"
start=$(date +%s)
"$TOOL" scan -i "$TARGET" -p 1-65535 \
  --time 2 --chan 3.0 --done 2 "${EXTRA[@]}" \
  --outname "$outname" "$@" 2>&1 \
  | perl -ne 'printf "%s %s", scalar(localtime()), $_' > "${outname}.console.log"
rc=${PIPESTATUS[0]}
end=$(date +%s)
echo "RESULT outname=$outname elapsed=$((end-start))s rc=$rc" | tee -a "${outname}.console.log"
echo "报告: ScanLog/${outname}-GoLin.xlsx / ${outname}-GoLin.html"
