#!/usr/bin/env python3
"""BRscan-alive-probe.py —— 全端口扫描前的存活预筛（golin --noping 12 小时陷阱的解法）。
原理: ICMP ping 扫全段；ping 不通的再用 TCP 建连探常用端口（收到 RST/SYN-ACK 都算存活）。
输出: alive.txt（一行一个 IP，直接喂 `golin scan --ipfile alive.txt --noping`）
用法: python3 BRscan-alive-probe.py 192.168.1.0/24 [> alive.txt]
精度说明: 全静默主机（不回 ICMP 且所有 TCP 也静默 DROP）无法与死主机区分，会漏出预筛；
若客户台账已给出主机清单，优先直接用 --ipfile 台账 --noping，不需要本脚本。
"""
import sys, re, socket, subprocess, concurrent.futures as cf

COMMON = [22, 80, 443, 445, 3389, 3306, 6379, 8080, 8443]  # ping 不通时的 TCP 复核端口

def parse_cidr(s):
    m = re.match(r"(\d+\.\d+\.\d+)\.0/24$", s)
    if m:
        return [f"{m.group(1)}.{i}" for i in range(1, 255)]
    raise SystemExit(f"暂只支持 /24: {s}")

def ping_ok(ip):
    r = subprocess.run(["ping", "-c", "1", "-t", "1", ip],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return r.returncode == 0

def tcp_alive(ip):
    """收到任何 TCP 响应（开放=成功 / 关闭=拒绝）都证明主机存活；超时=静默"""
    for p in COMMON:
        s = socket.socket(); s.settimeout(0.6)
        try:
            s.connect((ip, p)); return True          # 有开放端口
        except ConnectionRefusedError:
            return True                              # 有 RST = 主机活着
        except Exception:
            pass                                     # 超时：换下一个端口
        finally:
            s.close()
    return False

def main():
    net = sys.argv[1] if len(sys.argv) > 1 else "192.168.1.0/24"
    ips = parse_cidr(net)
    alive = []
    with cf.ThreadPoolExecutor(64) as ex:
        ping_hits = dict(zip(ips, ex.map(ping_ok, ips)))
    alive = [ip for ip, ok in ping_hits.items() if ok]
    silent = [ip for ip, ok in ping_hits.items() if not ok]
    with cf.ThreadPoolExecutor(32) as ex:
        tcp_hits = dict(zip(silent, ex.map(tcp_alive, silent)))
    alive += [ip for ip, ok in tcp_hits.items() if ok]
    alive.sort(key=lambda x: list(map(int, x.split("."))))
    for ip in alive:
        print(ip)
    print(f"# ping存活={sum(ping_hits.values())} TCP复核新增={sum(tcp_hits.values())} "
          f"合计={len(alive)} 静默/死主机={len(ips)-len(alive)}", file=sys.stderr)

if __name__ == "__main__":
    main()
