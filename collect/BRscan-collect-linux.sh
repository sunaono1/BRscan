#!/usr/bin/env bash
# ============================================================================
#  信息安全检查服务 · Linux 主机只读采集脚本
#  由 check/gen-collect-scripts.py 从《表B-上机核查表》自动生成 —— 请勿手改
#  覆盖：表B 中 38 个条目下的 L 类命令，共 76 条（另有 DB 类命令，用 WITH_DB=1 可选执行）
#
#  用法：  sudo bash collect_linux.sh [-o 输出目录]
#          WITH_DB=1 sudo -E bash collect_linux.sh      # 附带数据库类命令（需在装有客户端的跳板机执行）
#
#  纪律：  ① 全程只读，不改配置、不装软件、不重启服务；
#          ② 命令由我方提供，账号口令由客户管理员输入；
#          ③ 输出按表B 条目号分文件归档，逐项截图须含主机名/IP、命令行、系统时间、结论行；
#          ④ 弱口令只记是否命中与数量，不写明文。
#
#  注意：  缺失的命令行工具会打印 "command not found"，属正常，记录后人工补核；
#          单条命令超过 60 秒会被 timeout 打断（若系统无 timeout 则不限时）。
# ============================================================================
set -u

OUT=""
while [ $# -gt 0 ]; do
  case "$1" in
    -o|--out) OUT="$2"; shift 2 ;;
    -h|--help) sed -n "2,20p" "$0"; exit 0 ;;
    *) OUT="$1"; shift ;;
  esac
done
STAMP="$(date +%Y%m%d)"
OUT="${OUT:-./信息安全检查_${STAMP}/linux_$(hostname)}"
mkdir -p "$OUT" || exit 1
ALL="$OUT/全部输出.txt"; MAN="$OUT/清单.txt"; : > "$ALL"; : > "$MAN"
ITEMS="$OUT/采集输出"; mkdir -p "$ITEMS"

TIMEOUT_BIN=""
command -v timeout >/dev/null 2>&1 && TIMEOUT_BIN="$OUT/.timeout" && ln -sf "$(command -v timeout)" "$TIMEOUT_BIN"

hr()  { printf "%s\n" "============================================================"; }
line(){ printf "%s\n" "------------------------------------------------------------"; }
now() { date "+%F %T"; }

# 本机标识：优先默认路由出口 IP，回退 hostname -I / ip addr / ifconfig
LOCAL_IP=''
if command -v ip >/dev/null 2>&1; then
  LOCAL_IP="$(ip route get 1.1.1.1 2>/dev/null | awk '{for(i=1;i<=NF;i++) if($i=="src"){print $(i+1); exit}}')"
fi
[ -z "$LOCAL_IP" ] && LOCAL_IP="$(hostname -I 2>/dev/null | awk '{print $1}')"
[ -z "$LOCAL_IP" ] && LOCAL_IP="$(ip -4 addr show scope global 2>/dev/null | awk '/inet /{print $2; exit}' | cut -d/ -f1)"
[ -z "$LOCAL_IP" ] && LOCAL_IP="$(ifconfig 2>/dev/null | awk '/inet /{print $2; exit}' | cut -d: -f2)"
[ -z "$LOCAL_IP" ] && LOCAL_IP='(未知)'
HOST="$(hostname 2>/dev/null || echo unknown)"

echo "采集输出目录：$OUT"
if [ "$(id -u)" -ne 0 ]; then
  echo "[!] 当前非 root：脚本内含 sudo 命令，无权限的条目会记录为失败，请人工复核。"
  if [ -t 0 ] && ! sudo -n true 2>/dev/null; then
    echo "[i] 缓存 sudo 凭据（需要口令请现场输入）："
    sudo -v || { echo "[x] 未取得 sudo 凭据，建议改用：sudo bash $0"; }
  fi
fi

{ hr; echo "# Linux 主机只读采集 · 环境快照";
  echo "# 主机：$HOST    IP：$LOCAL_IP";
  echo "# 开始时间：$(now)";
  OSNAME="$(grep -E '^PRETTY_NAME=' /etc/os-release 2>/dev/null | cut -d= -f2- | tr -d '"')"
  [ -z "$OSNAME" ] && OSNAME="$(uname -sr)"
  echo "# 操作系统：$OSNAME";
  echo "# 内核：$(uname -r)   运行：$(uptime -p 2>/dev/null || uptime)";
  echo "# 执行人（sudo 身份）：$(id)";
  echo "# 脚本来源：由《表B-上机核查表》自动生成（38 个条目 / 76 条命令）";
  hr; } | tee "$OUT/环境信息.txt" | tee -a "$ALL"

begin() {  # $1 序号  $2 检查要点  $3 检查项  $4 结果形态  $5 对应测评项
  ITEM_FILE="$ITEMS/$(printf "%02d" "$1")_$(echo "$2" | tr "/ " "__").txt"
  { hr; echo "[表B 第 $1 条] $2"; echo "对应测评项：$5"; echo "检查项：$3";
    echo "主机：$HOST    IP：$LOCAL_IP    时间：$(now)"; line; } | tee "$ITEM_FILE" | tee -a "$ALL"
  printf "第 %s 条  %-24s  测评项 %-28s %s\n" "$1" "$2" "$5" "$ITEM_FILE" >> "$MAN"
  if [ -n "$5" ]; then
    IFS="、" read -ra __ITS <<< "$5"
    for __it in "${__ITS[@]}"; do
      ITEM_INDEX="$ITEMS/测评项_${__it}.txt"
      { echo "[表B 第 $1 条] $2"; echo "检查项：$3"; echo "输出文件：$ITEM_FILE";
        echo "主机：$HOST    IP：$LOCAL_IP    时间：$(now)"; } >> "$ITEM_INDEX"
    done
  fi
}

run() {  # $1 命令（逐字取自表B）
  { printf "$ %s\n" "$1"
    if [ -n "$TIMEOUT_BIN" ]; then timeout 60 bash -c "$1" 2>&1
    else bash -c "$1" 2>&1; fi
  } | tee -a "$ITEM_FILE" | tee -a "$ALL"
}

finish() {  # $1 序号  $2 结果形态（V=实测值 / L=执行结果粘贴）
  { line
    if [ "$2" = "V" ]; then echo "实测值：　　　　　　　　　　结论：□符合 □不符合 □不适用"
    else echo "执行结果粘贴："; fi
    echo "主机：$HOST    IP：$LOCAL_IP    时间：$(now)"
    hr; } | tee -a "$ITEM_FILE" | tee -a "$ALL"
  echo ""
}


# ─── 表B 板块：硬件与资产 ──────────────────────────────
echo ""; echo "──── 硬件与资产 ────"

# 表B 第 1 条 · 资产标识核对　→　测评项 1.3-1(佐证)、1.3-2(佐证)
begin 1 "资产标识核对" "核对设备序列号 / 资产标签与台账是否一致" L "1.3-1(佐证)、1.3-2(佐证)"
run 'sudo dmidecode -t system | grep -E '"'"'Product|Serial|Manufacturer'"'"''
finish 1 L

# 表B 第 2 条 · 主机规格　→　测评项 1.3-2(佐证)
begin 2 "主机规格" "记录 CPU / 内存 / 磁盘 / 网卡，作为资产基线" L "1.3-2(佐证)"
run 'lscpu | head -20; free -h; lsblk; ip -br a'
run 'smartctl -a /dev/sda | grep -Ei '"'"'Model|Serial|Health|Power_On_Hours'"'"'   # 磁盘健康'
run 'sudo dmidecode -t memory | grep -E '"'"'Size|Locator|Speed'"'"'   # 内存插槽与容量'
finish 2 L

# 表B 第 3 条 · 运行时长　→　测评项 1.3-2(佐证)
begin 3 "运行时长" "记录设备运行时长与最近重启时间（网络设备的版本回显含运行时长与重启时间，见第 8 条）" L "1.3-2(佐证)"
run 'uptime; who -b'
finish 3 L

# 表B 第 4 条 · 管理口暴露　→　测评项 3.3-2、4.10-1
begin 4 "管理口暴露" "服务器 BMC / IPMI / iDRAC / iLO 管理口是否与业务网隔离、是否改过出厂口令" V "3.3-2、4.10-1"
run 'sudo ss -tulnp | grep -E '"'"':623|:5900|:5988|:5989'"'"''
finish 4 V

# 表B 第 5 条 · 硬件告警　→　测评项 1.3-2(佐证)
begin 5 "硬件告警" "硬件是否有告警（电源 / 风扇 / RAID / 磁盘）" L "1.3-2(佐证)"
run 'sudo ipmitool sdr type '"'"'Power Supply'"'"'; sudo ipmitool sdr type Fan'
run 'sudo cat /proc/mdstat; sudo megacli -LDInfo -Lall -aALL 2>/dev/null | head'
finish 5 L

# 表B 第 6 条 · 外设与接口　→　测评项 4.10-2、5.3-1
begin 6 "外设与接口" "服务器 USB / 光驱是否按需禁用；是否存在未授权的无线网卡或外接设备（终端 U 盘管控见第 54 条）" V "4.10-2、5.3-1"
run 'lsusb; sudo lsblk -o NAME,SIZE,TYPE,TRAN'
finish 6 V

# 表B 第 7 条 · 机房与供电核验　→　测评项 11.1-1(佐证)、11.3-1(佐证)
begin 7 "机房与供电核验" "现场核对 UPS 覆盖范围、关键设备双电源接入、机柜标识" L "11.1-1(佐证)、11.3-1(佐证)"
run 'sudo dmidecode -t 39   # 系统电源与冗余信息'
run 'sudo ipmitool sdr type '"'"'Power Unit'"'"''
finish 7 L

# ─── 表B 板块：云与虚拟化资产核查 ──────────────────────────────
echo ""; echo "──── 云与虚拟化资产核查 ────"

# 表B 第 14 条 · 容器与 K8s 控制面　→　测评项 2.6-1
begin 14 "容器与 K8s 控制面" "API Server 暴露、RBAC 配置、Secret 明文与镜像仓库权限" L "2.6-1"
run 'kubectl get svc -A | grep -E "LoadBalancer|NodePort"'
run 'kubectl get clusterrolebinding -o wide | head -20'
run 'kubectl get secret -A --no-headers | head -20'
finish 14 L

# 表B 第 16 条 · 云主机元数据访问　→　测评项 2.7-1
begin 16 "云主机元数据访问" "实例元数据服务（169.254.169.254）是否可被未授权读取" V "2.7-1"
run 'timeout 3 bash -c "</dev/tcp/169.254.169.254/80" && echo "元数据服务可达" || echo "不可达"'
finish 16 V

# 表B 第 17 条 · 虚拟化平台管理面　→　测评项 2.6-3
begin 17 "虚拟化平台管理面" "vCenter / KVM / 超融合控制台端口暴露与版本" V "2.6-3"
run 'sudo ss -tulnp | grep -E '"'"':443|:5480|:9443|:902'"'"''
finish 17 V

# ─── 表B 板块：主机核查 · Linux ──────────────────────────────
echo ""; echo "──── 主机核查 · Linux ────"

# 表B 第 28 条 · 账号安全　→　测评项 4.1-1、4.1-2
begin 28 "账号安全" "空口令、UID=0、可登录的系统账号" V "4.1-1、4.1-2"
run 'sudo awk -F: '"'"'($2==""){print "空口令: "$1}'"'"' /etc/shadow'
run 'awk -F: '"'"'($3==0){print "UID0: "$1}'"'"' /etc/passwd'
run 'awk -F: '"'"'($3>=1000 && $7!~/(nologin|false)/){print $1" "$7}'"'"' /etc/passwd'
finish 28 V

# 表B 第 29 条 · 闲置账号　→　测评项 4.2-1
begin 29 "闲置账号" "90 天未登录账号与离职人员账号" L "4.2-1"
run 'sudo lastlog -b 90'
run 'sudo awk -F: '"'"'($3>=1000){print $1}'"'"' /etc/passwd | while read u; do echo "$u $(chage -l $u | grep '"'"'最近一次'"'"')   "; done'
finish 29 L

# 表B 第 30 条 · 口令策略　→　测评项 4.1-3
begin 30 "口令策略" "最小长度 / 复杂度 / 有效期 / 锁定策略" L "4.1-3"
run 'grep -E '"'"'PASS_MAX_DAYS|PASS_MIN_DAYS|PASS_MIN_LEN|PASS_WARN_AGE|ENCRYPT_METHOD'"'"' /etc/login.defs'
run 'grep -vE '"'"'^\s*#|^$'"'"' /etc/security/pwquality.conf'
run 'grep -E '"'"'pam_pwquality|pam_unix|pam_tally|pam_faillock'"'"' /etc/pam.d/system-auth /etc/pam.d/password-auth 2>/dev/null'
finish 30 L

# 表B 第 31 条 · 特权与提权　→　测评项 4.3-1、4.11-1
begin 31 "特权与提权" "sudo 免密、可写 sudoers、SUID 异常文件" L "4.3-1、4.11-1"
run 'getent group wheel sudo; sudo cat /etc/sudoers | grep -vE '"'"'^\s*#|^$'"'"'; sudo ls -l /etc/sudoers.d/'
run 'sudo find / -xdev -perm -4000 -type f 2>/dev/null'
run 'sudo find /etc /home /opt -xdev -writable -type f 2>/dev/null | head'
finish 31 L

# 表B 第 32 条 · 远程管理　→　测评项 4.4-1、4.4-2
begin 32 "远程管理" "SSH 配置：root 登录、口令认证、端口、版本" V "4.4-1、4.4-2"
run 'sudo grep -Ei '"'"'^\s*(Port|PermitRootLogin|PasswordAuthentication|PermitEmptyPasswords|Protocol|MaxAuthTries|AllowUsers|PubkeyAuthentication)'"'"' /etc/ssh/sshd_config'
run 'ssh -V; sudo ss -tlnp | grep -E '"'"':22|:3389|:23'"'"''
finish 32 V

# 表B 第 33 条 · 访问控制主机防火墙　→　测评项 4.5-1、4.5-2
begin 33 "访问控制主机防火墙" "主机防火墙策略与高危入站规则" L "4.5-1、4.5-2"
run 'sudo iptables -L INPUT -n --line-numbers | head -30'
run 'sudo nft list ruleset 2>/dev/null | head -40'
run 'sudo firewall-cmd --list-all 2>/dev/null'
finish 33 L

# 表B 第 34 条 · 会话与超时　→　测评项 4.4-2
begin 34 "会话与超时" "空闲会话超时、并发登录限制" V "4.4-2"
run 'grep -E '"'"'TMOUT'"'"' /etc/profile /etc/bashrc 2>/dev/null'
run 'grep -E '"'"'^\s*(ClientAliveInterval|ClientAliveCountMax|MaxSessions)'"'"' /etc/ssh/sshd_config'
finish 34 V

# 表B 第 35 条 · 日志与审计　→　测评项 7.1-1、7.1-2、7.2-1、7.2-2、7.3-2、7.4-1、7.6-1
begin 35 "日志与审计" "rsyslog / auditd 运行状态、日志留存时长、是否集中外发（≥ 6 个月）；Windows 侧日志容量与模式" L "7.1-1、7.1-2、7.2-1、7.2-2、7.3-2、7.4-1、7.6-1"
run 'systemctl is-active rsyslog auditd; journalctl --disk-usage'
run 'ls -lh --time-style=long-iso /var/log/ | head -20'
run 'grep -E '"'"'^\s*dest|^\*\*'"'"' /etc/rsyslog.conf /etc/rsyslog.d/*.conf 2>/dev/null | head'
finish 35 L

# 表B 第 36 条 · 补丁与版本　→　测评项 4.6-1、4.6-2
begin 36 "补丁与版本" "OS 版本、内核、EOL 判断、安全更新遗漏" L "4.6-1、4.6-2"
run 'cat /etc/os-release; uname -r; uptime'
run 'sudo yum check-update --security 2>/dev/null | head; sudo apt-get -s upgrade 2>/dev/null | grep -i security | head'
run 'sudo --version | head -1; openssl version; ssh -V'
finish 36 L

# 表B 第 37 条 · 服务最小化　→　测评项 4.7-1
begin 37 "服务最小化" "不必要服务与监听端口" L "4.7-1"
run 'systemctl list-units --type=service --state=running'
run 'systemctl list-unit-files --state=enabled'
run 'sudo ss -tulnp'
run 'for s in telnet.socket rsh.socket vsftpd tftp snmpd rpcbind avahi-daemon; do printf '"'"'%-16s %s\n'"'"' $s "$(systemctl is-active $s 2>/dev/null)"; done'
finish 37 L

# 表B 第 38 条 · 共享服务　→　测评项 4.7-2、8.1-1
begin 38 "共享服务" "Samba / NFS 共享权限与匿名访问" L "4.7-2、8.1-1"
run 'sudo testparm -s 2>/dev/null | head -40'
run 'sudo cat /etc/exports | grep -vE '"'"'^\s*#'"'"''
run 'showmount -e 127.0.0.1; smbclient -L //127.0.0.1 -N'
finish 38 L

# 表B 第 39 条 · 恶意程序初排　→　测评项 4.8-1
begin 39 "恶意程序初排" "异常进程、预加载库、近期可疑文件落点" L "4.8-1"
run 'ps aux --sort=-%cpu | head -15'
run 'sudo find /tmp /dev/shm /var/tmp /var/www -type f -mtime -30 -ls 2>/dev/null | head'
run 'cat /etc/ld.so.preload'
finish 39 L

# 表B 第 40 条 · Web 后门　→　测评项 4.8-1
begin 40 "Web 后门" "Web 目录下的 webshell 特征文件" L "4.8-1"
run 'sudo find /var/www /usr/share/nginx /opt -type f \\( -name '"'"'*.php'"'"' -o -name '"'"'*.jsp'"'"' -o -name '"'"'*.aspx'"'"' \\) -mtime -90 -ls 2>/dev/null | head -20'
run 'sudo grep -rlE '"'"'eval\\(|base64_decode|system\\(|shell_exec|assert\\(|Request\\.Item'"'"' /var/www 2>/dev/null | head'
finish 40 L

# 表B 第 41 条 · 内核加固　→　测评项 4.9-2
begin 41 "内核加固" "sysctl 安全参数与 SELinux / AppArmor" V "4.9-2"
run 'sudo sysctl -a 2>/dev/null | grep -E '"'"'ip_forward|accept_redirects|accept_source_route|syncookies|rp_filter'"'"''
run 'getenforce; sestatus 2>/dev/null | head -5; aa-status 2>/dev/null | head -5'
finish 41 V

# 表B 第 42 条 · 计划任务与自启　→　测评项 4.9-1
begin 42 "计划任务与自启" "异常自启项与定时任务" L "4.9-1"
run 'systemctl list-timers --all --no-pager'
run 'ls -l /etc/rc.local /etc/init.d/ 2>/dev/null | head; sudo crontab -l'
run 'crontab -l; sudo cat /etc/crontab; ls -l /etc/cron.d/ /etc/cron.hourly/ 2>/dev/null'
finish 42 L

# 表B 第 43 条 · Grub 与单用户　→　测评项 4.9-2
begin 43 "Grub 与单用户" "引导口令与单用户模式保护" V "4.9-2"
run 'sudo grep -E '"'"'GRUB_TIMEOUT|password'"'"' /etc/grub.conf /boot/grub/grub.cfg /etc/default/grub 2>/dev/null | head'
finish 43 V

# 表B 第 44 条 · 数据加密　→　测评项 4.10-2、6.9-2
begin 44 "数据加密" "是否启用磁盘 / 分区加密" V "4.10-2、6.9-2"
run 'lsblk -f | grep -i crypto; sudo cryptsetup status root 2>/dev/null; sudo dmsetup status 2>/dev/null | head'
finish 44 V

# ─── 表B 板块：数据库核查 ──────────────────────────────
echo ""; echo "──── 数据库核查 ────"

# 表B 第 63 条 · 中间件 Tomcat / Nginx / IIS　→　测评项 6.5-1、6.5-2
begin 63 "中间件 Tomcat / Nginx / IIS" "默认管理页、示例应用、目录浏览、版本泄露" L "6.5-1、6.5-2"
run 'nginx -v; ls -la /etc/nginx/conf.d/ /usr/share/nginx/html/ 2>/dev/null | head'
run 'grep -E '"'"'autoindex|server_tokens|listings'"'"' /etc/nginx/nginx.conf /etc/nginx/conf.d/*.conf 2>/dev/null'
run 'ls /opt/tomcat/webapps/ /var/lib/tomcat*/webapps/ 2>/dev/null'
finish 63 L

# 表B 第 64 条 · 数据库采集脚本　→　测评项 6.11-1
begin 64 "数据库采集脚本" "整体采集数据库配置做交叉核对" L "6.11-1"
run 'sudo cat /etc/mysql/my.cnf /etc/my.cnf /etc/redis/redis.conf 2>/dev/null | grep -vE '"'"'^\s*#|^$'"'"' | head -40'
finish 64 L

# ─── 表B 板块：日志 · 审计 · 权限 · 备份 ──────────────────────────────
echo ""; echo "──── 日志 · 审计 · 权限 · 备份 ────"

# 表B 第 65 条 · 时间同步　→　测评项 7.5-1
begin 65 "时间同步" "全网时钟同步状态，日志时间戳可比对" L "7.5-1"
run 'timedatectl; chronyc sources 2>/dev/null | head'
finish 65 L

# 表B 第 66 条 · 日志完整性　→　测评项 7.4-1
begin 66 "日志完整性" "日志篡改 / 清空迹象（Windows 事件 1102 / 104）" L "7.4-1"
run 'sudo stat -c '"'"'%a %U:%G %y %n'"'"' /var/log/secure /var/log/messages'
finish 66 L

# 表B 第 67 条 · 权限矩阵　→　测评项 4.1-2、5.1-1
begin 67 "权限矩阵" "账号与权限台账、最小权限落地情况" L "4.1-2、5.1-1"
run 'sudo awk -F: '"'"'($3>=1000){print $1,$3,$7}'"'"' /etc/passwd; getent group | awk -F: '"'"'NF>3'"'"' | head -20'
finish 67 L

# 表B 第 68 条 · 敏感目录越权　→　测评项 8.1-1
begin 68 "敏感目录越权" "重要数据目录是否全员可写 / 匿名可读（Windows 侧见第 50 条）" L "8.1-1"
run 'sudo find /data /opt /srv -xdev -type d -perm -0002 2>/dev/null | head'
run 'sudo find /data -xdev -type f -perm -0004 2>/dev/null | head'
finish 68 L

# 表B 第 69 条 · 不可信外联　→　测评项 4.8-2
begin 69 "不可信外联" "主机异常外联与非常规端口连接" L "4.8-2"
run 'sudo ss -antp state established | grep -vE '"'"'127.0.0.1|::1'"'"''
finish 69 L

# 表B 第 70 条 · 备份任务　→　测评项 8.6-2(佐证)
begin 70 "备份任务" "备份任务是否成功、保留策略、是否同盘同机" L "8.6-2(佐证)"
run 'crontab -l | grep -i backup; ls -lh /backup 2>/dev/null | head'
run 'df -h | grep -Ei '"'"'backup|data'"'"''
finish 70 L

# 表B 第 71 条 · 备份可用性　→　测评项 8.8-1(佐证)
begin 71 "备份可用性" "备份文件是否可读 / 可恢复（抽样验证）" L "8.8-1(佐证)"
run 'sudo tar -tzf /backup/LAST_BACKUP.tar.gz 2>/dev/null | head   # 只读校验，不实际恢复'
run 'sudo mysqlcheck --all-databases --check-only 2>/dev/null | head   # 抽查数据库一致性'
finish 71 L

# 表B 第 72 条 · 采集脚本执行　→　测评项 内部
begin 72 "采集脚本执行" "运行采集脚本，核对输出文件是否齐全" L "内部"
run 'bash collect_linux.sh'
finish 72 L

# 表B 第 73 条 · 证据归档　→　测评项 内部
begin 73 "证据归档" "全部输出文件归档，截图须含主机名或 IP、时间与结论" L "内部"
run 'mkdir -p 信息安全检查_$(date +%Y-%m-%d) && command | tee 信息安全检查_$(date +%Y-%m-%d)/条目号_资产名.txt'
run 'sed -n '"'"'1,80p'"'"' 信息安全检查_$(date +%Y-%m-%d)/条目号_资产名.txt   # 复核输出内容完整性'
finish 73 L

if [ "${WITH_DB:-0}" = "1" ]; then
  echo ""; echo "──── 附加：数据库类命令（表B DB 类，可选）────"

  # 表B 第 55 条 · MySQL 账号　→　测评项 6.7-1、6.7-2
  begin 55 "MySQL 账号（DB）" "由数据库客户端执行，详见表B 对应条目" V "6.7-1、6.7-2"
  run 'SELECT user,host,authentication_string FROM mysql.user;'
  run 'SHOW GRANTS FOR '"'"'USER'"'"'@'"'"'HOST'"'"';'
  finish 55 V

  # 表B 第 56 条 · MySQL 口令插件　→　测评项 6.8-1
  begin 56 "MySQL 口令插件（DB）" "由数据库客户端执行，详见表B 对应条目" V "6.8-1"
  run 'show variables like '"'"'validate_password%'"'"';'
  run 'show variables like '"'"'default_password_lifetime'"'"';'
  run 'SELECT USER,DEFAULT_ROLE_USER FROM mysql.default_roles;'
  finish 56 V

  # 表B 第 57 条 · MySQL 日志与审计　→　测评项 6.8-2
  begin 57 "MySQL 日志与审计（DB）" "由数据库客户端执行，详见表B 对应条目" V "6.8-2"
  run 'show variables like '"'"'general_log%'"'"';'
  run 'show variables like '"'"'slow_query_log%'"'"';'
  run 'show variables like '"'"'log_bin'"'"';'
  run 'SHOW VARIABLES LIKE '"'"'datadir'"'"';'
  finish 57 V

  # 表B 第 58 条 · MySQL 连接与超时　→　测评项 6.9-1、6.9-2
  begin 58 "MySQL 连接与超时（DB）" "由数据库客户端执行，详见表B 对应条目" V "6.9-1、6.9-2"
  run 'show variables like '"'"'%connection_control%'"'"';'
  run 'show global variables like '"'"'connect_timeout'"'"';'
  run 'show global variables like '"'"'wait_timeout'"'"';'
  finish 58 V

  # 表B 第 59 条 · PostgreSQL 连接与恢复　→　测评项 6.10-1
  begin 59 "PostgreSQL 连接与恢复（DB）" "由数据库客户端执行，详见表B 对应条目" V "6.10-1"
  run 'SHOW listen_addresses;'
  run 'SHOW log_connections; SHOW log_disconnections; SHOW log_statement;'
  run 'SHOW log_destination; SHOW logging_collector; SHOW log_filename;'
  run 'SHOW ssl_min_protocol_version; SHOW shared_preload_libraries;'
  run 'select pg_is_in_recovery(); show transaction_read_only;'
  finish 59 V

  # 表B 第 60 条 · Redis 安全　→　测评项 6.10-1
  begin 60 "Redis 安全（DB）" "由数据库客户端执行，详见表B 对应条目" V "6.10-1"
  run 'redis-cli -h HOST -p 6379 PING'
  run 'redis-cli -h HOST CONFIG GET requirepass; redis-cli CONFIG GET bind'
  finish 60 V

  # 表B 第 61 条 · Oracle 账号与口令　→　测评项 6.10-1
  begin 61 "Oracle 账号与口令（DB）" "由数据库客户端执行，详见表B 对应条目" V "6.10-1"
  run 'SELECT * FROM DBA_PROFILES;'
  run 'SELECT USER#, NAME, TYPE#, PASSWORD, CTIME, PTIME, EXPTIME, LTIME FROM sys.user$;'
  finish 61 V

  # 表B 第 62 条 · SQL Server 安全　→　测评项 6.10-1
  begin 62 "SQL Server 安全（DB）" "由数据库客户端执行，详见表B 对应条目" V "6.10-1"
  run 'SELECT encrypt_option FROM sys.dm_exec_connections WHERE session_id = @@SPID;'
  run 'SELECT name,is_disabled FROM sys.server_principals WHERE type IN ('"'"'S'"'"','"'"'U'"'"');'
  finish 62 V
else
  echo ""; echo "[i] 已跳过数据库类命令（共 23 条）。需要时用 WITH_DB=1 重跑。"
fi

{ hr; echo "# 采集结束：$(now)"; echo "# 明细索引见 清单.txt，全量输出见 全部输出.txt"; hr; } | tee -a "$ALL"
echo ""; echo "✔ 采集完成：$OUT"
echo "  条目文件：$ITEMS/"; echo "  全量输出：$ALL"; echo "  条目索引：$MAN"
