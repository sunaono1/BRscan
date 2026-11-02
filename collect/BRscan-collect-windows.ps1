<#
.SYNOPSIS
    信息安全检查服务 · Windows 主机 / 终端只读采集脚本
.DESCRIPTION
    由 check/gen-collect-scripts.py 从《表B-上机核查表》自动生成 —— 请勿手改。
    覆盖：表B 中 24 个条目下的 W 类命令，共 40 条。

    用法（必须以管理员身份运行）：
        powershell -ExecutionPolicy Bypass -File .\collect_windows.ps1
        powershell -ExecutionPolicy Bypass -File .\collect_windows.ps1 -OutDir D:\信息安全检查_20260101

    纪律：① 全程只读，不改配置、不装软件、不重启服务；
          ② 命令由我方提供，账号口令由客户管理员输入；
          ③ 输出按表B 条目号分文件归档，逐项截图须含主机名/IP、命令行、系统时间、结论行；
          ④ 弱口令只记是否命中与数量，不写明文。

    注意：wmic 在新版 Windows 上可能已移除，对应条目会记录执行异常；
          同条目内通常另有 Get-LocalUser / Get-CimInstance 等价命令，以可执行的那条为准。
#>
param([string]$OutDir)

$ErrorActionPreference = "Continue"
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$OutputEncoding = [System.Text.Encoding]::UTF8

$Stamp  = Get-Date -Format "yyyyMMdd"
$Host_  = $env:COMPUTERNAME
if (-not $OutDir) { $OutDir = Join-Path (Get-Location) ("信息安全检查_{0}\windows_{1}" -f $Stamp, $Host_) }
$ItemDir = Join-Path $OutDir "采集输出"
New-Item -ItemType Directory -Path $ItemDir -Force | Out-Null
$All = Join-Path $OutDir "全部输出.txt"
$Man = Join-Path $OutDir "清单.txt"
"" | Set-Content -Path $All -Encoding UTF8
"" | Set-Content -Path $Man -Encoding UTF8

$script:ItemFile = $null
$HR   = "============================================================"
$LINE = "------------------------------------------------------------"

$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) { Write-Host "[!] 未以管理员身份运行：审计策略、日志、Defender 等条目会失败。建议重开管理员 PowerShell。" -ForegroundColor Yellow }

# 本机 IP：优先默认路由出口
$LocalIp = "（未知）"
try {
    $LocalIp = (Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -ne $null } | Select-Object -First 1).IPv4Address.IPAddress
} catch {}
if (-not $LocalIp) { try { $LocalIp = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike "169.254.*" -and $_.IPAddress -ne "127.0.0.1" } | Select-Object -First 1).IPAddress } catch {} }

function Write-Both([string]$Text, [string]$Path) {
    if ($Path) { Add-Content -Path $Path -Value $Text -Encoding UTF8 }
    Write-Host $Text
}

function Begin-Item([int]$No, [string]$Tag, [string]$Item, [string]$Shape, [string]$Items) {
    $safe = ($Tag -replace '[/\\:*?"<>|\s]+', '_')
    $script:ItemFile = Join-Path $ItemDir ("{0:00}_{1}.txt" -f $No, $safe)
    $head = @($HR, ("[表B 第 " + $No + " 条] " + $Tag), ("对应测评项：" + $Items), ("检查项：" + $Item),
              ("主机：" + $env:COMPUTERNAME + "    IP：" + $LocalIp + "    时间：" + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")), $LINE)
    "" | Set-Content -Path $script:ItemFile -Encoding UTF8
    foreach ($h in $head) { Write-Both $h $script:ItemFile; Add-Content -Path $All -Value $h -Encoding UTF8 }
    Add-Content -Path $Man -Value ("第 {0} 条  {1,-24} 测评项 {2,-28} {3}" -f $No, $Tag, $Items, $script:ItemFile) -Encoding UTF8
    if ($Items) {
        foreach ($it in ($Items -split "、")) {
            $idx = Join-Path $ItemDir ("测评项_" + $it + ".txt")
            Add-Content -Path $idx -Value ("[表B 第 {0} 条] {1}" -f $No, $Tag) -Encoding UTF8
            Add-Content -Path $idx -Value ("检查项：" + $Item) -Encoding UTF8
            Add-Content -Path $idx -Value ("输出文件：" + $script:ItemFile) -Encoding UTF8
        }
    }
    Write-Host ""
}

function Invoke-Item-Cmd([string]$Cmd) {
    $echo = "PS> " + $Cmd
    Write-Both $echo $script:ItemFile; Add-Content -Path $All -Value $echo -Encoding UTF8
    $res = ""
    try {
        $res = (& ([ScriptBlock]::Create($Cmd)) 2>&1 | Out-String)
    } catch {
        $res = "执行异常：" + $_.Exception.Message
    }
    if (-not $res) { $res = "（无输出）" }
    foreach ($ln in ($res -split "`r?`n")) {
        if ($ln -ne "") { Add-Content -Path $script:ItemFile -Value $ln -Encoding UTF8; Add-Content -Path $All -Value $ln -Encoding UTF8 }
    }
    Write-Host ($res.TrimEnd())
}

function End-Item([int]$No, [string]$Shape) {
    $tail = @($LINE)
    if ($Shape -eq "V") { $tail += "实测值：　　　　　　　　　　结论：□符合 □不符合 □不适用" }
    else { $tail += "执行结果粘贴：" }
    $tail += ("主机：" + $env:COMPUTERNAME + "    IP：" + $LocalIp + "    时间：" + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
    $tail += $HR
    foreach ($t in $tail) { Write-Both $t $script:ItemFile; Add-Content -Path $All -Value $t -Encoding UTF8 }
}

Write-Host ("采集输出目录：" + $OutDir)
$envTxt = Join-Path $OutDir "环境信息.txt"
$envLines = @($HR, "# Windows 主机只读采集 · 环境快照", ("# 主机：" + $env:COMPUTERNAME + "    IP：" + $LocalIp),
              ("# 开始时间：" + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")), ("# 当前用户：" + [Security.Principal.WindowsIdentity]::GetCurrent().Name),
              ("# 管理员权限：" + $IsAdmin), ("# PowerShell：" + $PSVersionTable.PSVersion.ToString()), $HR)
$envLines | Set-Content -Path $envTxt -Encoding UTF8
$envLines | Add-Content -Path $All -Encoding UTF8
$envLines | ForEach-Object { Write-Host $_ }


# ─── 表B 板块：硬件与资产 ──────────────────────────────
Write-Host ""; Write-Host "──── 硬件与资产 ────" -ForegroundColor Cyan

# 表B 第 1 条 · 资产标识核对　→　测评项 1.3-1(佐证)、1.3-2(佐证)
Begin-Item 1 "资产标识核对" "核对设备序列号 / 资产标签与台账是否一致" L "1.3-1(佐证)、1.3-2(佐证)"
Invoke-Item-Cmd 'wmic bios get serialnumber,manufacturer; wmic csproduct get name,identifyingnumber'
End-Item 1 L

# 表B 第 2 条 · 主机规格　→　测评项 1.3-2(佐证)
Begin-Item 2 "主机规格" "记录 CPU / 内存 / 磁盘 / 网卡，作为资产基线" L "1.3-2(佐证)"
Invoke-Item-Cmd 'systeminfo | findstr /C:''处理器'' /C:''内存'' /C:''系统型号'''
End-Item 2 L

# 表B 第 3 条 · 运行时长　→　测评项 1.3-2(佐证)
Begin-Item 3 "运行时长" "记录设备运行时长与最近重启时间（网络设备的版本回显含运行时长与重启时间，见第 8 条）" L "1.3-2(佐证)"
Invoke-Item-Cmd '(Get-CimInstance Win32_OperatingSystem).LastBootUpTime'
End-Item 3 L

# 表B 第 5 条 · 硬件告警　→　测评项 1.3-2(佐证)
Begin-Item 5 "硬件告警" "硬件是否有告警（电源 / 风扇 / RAID / 磁盘）" L "1.3-2(佐证)"
Invoke-Item-Cmd 'Get-CimInstance Win32_DiskDrive | select Model,Status,Size'
End-Item 5 L

# 表B 第 6 条 · 外设与接口　→　测评项 4.10-2、5.3-1
Begin-Item 6 "外设与接口" "服务器 USB / 光驱是否按需禁用；是否存在未授权的无线网卡或外接设备（终端 U 盘管控见第 54 条）" V "4.10-2、5.3-1"
Invoke-Item-Cmd 'Get-PnpDevice -Class USB | where Status -eq ''OK'' | select FriendlyName,Status'
End-Item 6 V

# ─── 表B 板块：云与虚拟化资产核查 ──────────────────────────────
Write-Host ""; Write-Host "──── 云与虚拟化资产核查 ────" -ForegroundColor Cyan

# 表B 第 16 条 · 云主机元数据访问　→　测评项 2.7-1
Begin-Item 16 "云主机元数据访问" "实例元数据服务（169.254.169.254）是否可被未授权读取" V "2.7-1"
Invoke-Item-Cmd 'Test-NetConnection 169.254.169.254 -Port 80 -InformationLevel Quiet'
End-Item 16 V

# ─── 表B 板块：主机核查 · Linux ──────────────────────────────
Write-Host ""; Write-Host "──── 主机核查 · Linux ────" -ForegroundColor Cyan

# 表B 第 35 条 · 日志与审计　→　测评项 7.1-1、7.1-2、7.2-1、7.2-2、7.3-2、7.4-1、7.6-1
Begin-Item 35 "日志与审计" "rsyslog / auditd 运行状态、日志留存时长、是否集中外发（≥ 6 个月）；Windows 侧日志容量与模式" L "7.1-1、7.1-2、7.2-1、7.2-2、7.3-2、7.4-1、7.6-1"
Invoke-Item-Cmd 'Get-WinEvent -ListLog Security | select LogName,IsEnabled,RecordCount,MaximumSizeInBytes,LogMode'
End-Item 35 L

# ─── 表B 板块：主机核查 · Windows ──────────────────────────────
Write-Host ""; Write-Host "──── 主机核查 · Windows ────" -ForegroundColor Cyan

# 表B 第 45 条 · 账号安全　→　测评项 4.1-1、4.1-2、4.1-3、5.4-1、5.4-2
Begin-Item 45 "账号安全" "本地账号、禁用状态、口令策略、空口令" L "4.1-1、4.1-2、4.1-3、5.4-1、5.4-2"
Invoke-Item-Cmd 'net user; net localgroup administrators; net accounts'
Invoke-Item-Cmd 'wmic useraccount get name,disabled,passwordrequired,passwordexpires'
Invoke-Item-Cmd 'Get-LocalUser | ft Name,Enabled,PasswordRequired,PasswordLastSet,LastLogon'
End-Item 45 L

# 表B 第 46 条 · 闲置账号　→　测评项 4.2-1、5.4-2
Begin-Item 46 "闲置账号" "长期未登录账号与离职账号" L "4.2-1、5.4-2"
Invoke-Item-Cmd 'net user administrator | findstr /C:"密码上次设置" /C:"上次登录" /C:"密码要求"'
Invoke-Item-Cmd 'Get-LocalUser | ft Name,LastLogon,PasswordLastSet'
End-Item 46 L

# 表B 第 47 条 · 远程管理　→　测评项 4.4-1、4.4-2
Begin-Item 47 "远程管理" "RDP 是否开启、NLA、端口、是否允许任意源" V "4.4-1、4.4-2"
Invoke-Item-Cmd 'reg query "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections'
Invoke-Item-Cmd 'reg query "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" /v UserAuthentication'
Invoke-Item-Cmd 'reg query "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" /v PortNumber'
End-Item 47 V

# 表B 第 48 条 · 补丁与版本　→　测评项 4.6-1、4.6-2、5.2-2
Begin-Item 48 "补丁与版本" "系统版本、EOL 判断、已装补丁" L "4.6-1、4.6-2、5.2-2"
Invoke-Item-Cmd 'systeminfo | findstr /C:"OS 名称" /C:"OS 版本" /C:"系统启动时间"'
Invoke-Item-Cmd 'wmic qfe get HotFixID,InstalledOn | sort'
Invoke-Item-Cmd 'Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 10'
End-Item 48 L

# 表B 第 49 条 · 审计策略　→　测评项 7.3-1
Begin-Item 49 "审计策略" "Windows 审计策略是否覆盖登录、账号变更、特权操作" L "7.3-1"
Invoke-Item-Cmd 'auditpol /get /category:*'
Invoke-Item-Cmd 'wevtutil gl Security'
End-Item 49 L

# 表B 第 50 条 · 共享与权限　→　测评项 4.7-2、8.1-1
Begin-Item 50 "共享与权限" "共享目录权限、匿名可读、Everyone 写权限" L "4.7-2、8.1-1"
Invoke-Item-Cmd 'net share; Get-SmbShare | % { Get-SmbShareAccess -Name $_.Name } | ft'
Invoke-Item-Cmd 'icacls "D:\Data" /T /C | findstr /I "Everyone Users 所有人"'
Invoke-Item-Cmd 'Get-SmbServerConfiguration | select EnableSMB1Protocol,RequireSecuritySignature'
End-Item 50 L

# 表B 第 51 条 · 本地提权与 UAC　→　测评项 4.3-2、5.1-1、5.1-2
Begin-Item 51 "本地提权与 UAC" "终端用户是否本地管理员、UAC 状态" V "4.3-2、5.1-1、5.1-2"
Invoke-Item-Cmd 'net localgroup administrators'
Invoke-Item-Cmd 'reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v EnableLUA'
Invoke-Item-Cmd 'reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v ConsentPromptBehaviorAdmin'
End-Item 51 V

# 表B 第 52 条 · 不必要服务　→　测评项 4.7-1
Begin-Item 52 "不必要服务" "Telnet / RemoteRegistry / FTPSVC / SNMP 等" L "4.7-1"
Invoke-Item-Cmd 'sc query RemoteRegistry; sc query TermService; sc query SNMP; sc query FTPSVC'
Invoke-Item-Cmd 'Get-Service | where {$_.Status -eq ''Running''} | select Name,DisplayName | sort Name'
End-Item 52 L

# 表B 第 53 条 · 杀毒与防护　→　测评项 5.2-1
Begin-Item 53 "杀毒与防护" "杀毒 / Defender 状态与病毒库时间" V "5.2-1"
Invoke-Item-Cmd 'Get-MpComputerStatus | select AntivirusEnabled,RealTimeProtectionEnabled,AntivirusSignatureLastUpdated'
Invoke-Item-Cmd 'Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct | select displayName,productState'
End-Item 53 V

# 表B 第 54 条 · 终端移动介质　→　测评项 5.3-1
Begin-Item 54 "终端移动介质" "U 盘是否可随意使用" V "5.3-1"
Invoke-Item-Cmd 'reg query "HKLM\SYSTEM\CurrentControlSet\Services\USBSTOR" /v Start'
Invoke-Item-Cmd 'reg query "HKLM\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices" /s 2>nul'
End-Item 54 V

# ─── 表B 板块：数据库核查 ──────────────────────────────
Write-Host ""; Write-Host "──── 数据库核查 ────" -ForegroundColor Cyan

# 表B 第 63 条 · 中间件 Tomcat / Nginx / IIS　→　测评项 6.5-1、6.5-2
Begin-Item 63 "中间件 Tomcat / Nginx / IIS" "默认管理页、示例应用、目录浏览、版本泄露" L "6.5-1、6.5-2"
Invoke-Item-Cmd 'Get-WebConfiguration //system.webServer/directoryBrowse -PSPath ''IIS:\'''
End-Item 63 L

# ─── 表B 板块：日志 · 审计 · 权限 · 备份 ──────────────────────────────
Write-Host ""; Write-Host "──── 日志 · 审计 · 权限 · 备份 ────" -ForegroundColor Cyan

# 表B 第 65 条 · 时间同步　→　测评项 7.5-1
Begin-Item 65 "时间同步" "全网时钟同步状态，日志时间戳可比对" L "7.5-1"
Invoke-Item-Cmd 'w32tm /query /status; w32tm /query /source'
End-Item 65 L

# 表B 第 66 条 · 日志完整性　→　测评项 7.4-1
Begin-Item 66 "日志完整性" "日志篡改 / 清空迹象（Windows 事件 1102 / 104）" L "7.4-1"
Invoke-Item-Cmd 'Get-WinEvent -FilterHashtable @{LogName=''Security'';ID=1102} -MaxEvents 5'
Invoke-Item-Cmd 'Get-WinEvent -FilterHashtable @{LogName=''System'';ID=104} -MaxEvents 5'
End-Item 66 L

# 表B 第 67 条 · 权限矩阵　→　测评项 4.1-2、5.1-1
Begin-Item 67 "权限矩阵" "账号与权限台账、最小权限落地情况" L "4.1-2、5.1-1"
Invoke-Item-Cmd 'Get-LocalGroupMember Administrators; Get-LocalUser | ft Name,Enabled'
End-Item 67 L

# 表B 第 69 条 · 不可信外联　→　测评项 4.8-2
Begin-Item 69 "不可信外联" "主机异常外联与非常规端口连接" L "4.8-2"
Invoke-Item-Cmd 'Get-NetTCPConnection -State Established | select RemoteAddress,RemotePort,OwningProcess | sort RemoteAddress'
End-Item 69 L

# 表B 第 70 条 · 备份任务　→　测评项 8.6-2(佐证)
Begin-Item 70 "备份任务" "备份任务是否成功、保留策略、是否同盘同机" L "8.6-2(佐证)"
Invoke-Item-Cmd 'wbadmin get status; Get-WBSummary 2>$null'
End-Item 70 L

# 表B 第 72 条 · 采集脚本执行　→　测评项 内部
Begin-Item 72 "采集脚本执行" "运行采集脚本，核对输出文件是否齐全" L "内部"
Invoke-Item-Cmd 'powershell -ExecutionPolicy Bypass -File collect_windows.ps1'
End-Item 72 L

$closing = @($HR, ("# 采集结束：" + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")), "# 明细索引见 清单.txt，全量输出见 全部输出.txt", $HR)
$closing | Add-Content -Path $All -Encoding UTF8
$closing | ForEach-Object { Write-Host $_ }
Write-Host ""
Write-Host ("✔ 采集完成：" + $OutDir) -ForegroundColor Green
Write-Host ("  条目文件：" + $ItemDir)
Write-Host ("  全量输出：" + $All)
Write-Host ("  条目索引：" + $Man)
