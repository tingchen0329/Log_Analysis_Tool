# =============================================================================
# Design By Ting Chen
# Cisco_Log_Analysis.py
#
# 功能：讀取 Cisco / Aruba / FortiGate 設備 Log 檔案，解析設備資訊並彙整輸出成 Excel 報表；
#       並可比對上季 / 本季兩份維護報表，列出新增、移除、換機與欄位異動（維護設備比對）。
#
# 支援設備：
#   - Cisco C9800 WLC / EWC（IOS-XE，含 HA SSO 雙機判定）
#   - Cisco AireOS WLC（2504 / 3504 / 5520）
#   - Catalyst IOS / IOS-XE Switch（C2960 / C1000 / C3850 / C9K，含 Stack 堆疊）
#   - Cisco Nexus NX-OS（N5K / N9K）
#   - Cisco Catalyst 1300（show tech-support 彙整格式）
#   - Aruba / HPE AOS-CX Switch（6300 / 6400，含 VSF 堆疊）
#   - Fortinet FortiGate（FortiOS：.conf 備份檔或 CLI 輸出，含 HA 叢集）
#   - WLC 底下的 AP（C9800：show ap summary + show ap config general；
#                    AireOS：show ap join stats summary all + show ap inventory all）
#
# 使用方式：
#   GUI：python Cisco_Log_Analysis.py             （或 pythonw 無黑視窗）
#   CLI：python Cisco_Log_Analysis.py <Log資料夾> [-o 報表.xlsx] [--compare-with 上季.xlsx]
#   比對：python Cisco_Log_Analysis.py --compare 上季.xlsx 本季.xlsx [-o 比對.xlsx]
#
# V23 變更摘要（新功能：維護設備比對）：
#   1. 比對上季 / 本季兩份維護報表（本工具產出的 .xlsx），以 Serial Number 配對，輸出差異報表：
#      第一頁「設備總表」= 本季完整設備 + AP（移除的設備列在最下方），以顏色標註
#      （新增淡綠、移除淡紅、換機只標 SN 深紅、異動格子淺黃 + 註解上季值、需確認橘黃）；
#      摘要（1～6 條列，以台數計）、新增設備、移除設備、換機與成員變更、欄位異動、
#      新增 AP、移除 AP、需人工確認（無序號或序號重複）
#   2. Stack / HA 非首列的「○」先補成首列的值再比對；共用欄位（Hostname、OOB IP）只在首列比較
#   3. GUI 新增「分析 Log / 維護設備比對」切換；CLI 新增 --compare 與 --compare-with
#   4. Excel 寫入共用 write_workbook()（維護報表與比對報表格式一致）
#
# V22 變更摘要（依 Cisco / HPE Aruba 官方文件比對修正）：
#   1. 環境異常關鍵字改為字首比對：Failure / Failed / Faulty 也算異常
#      （V21 的 \bFAIL\b 只認完整單字，NX-OS 官方範例的「failure or missing」會被判為正常）
#   2. NX-OS：電源 Shutdown（拔線或斷電）判為異常；「Fan_in_PSn」列歸類為電源；
#      N3K 的風扇列「Fan-1 …」會被誤認成區段標頭而略過，改為只認「Fan:」與欄位標題列
#   3. AireOS：電源改看 show sysinfo 的「Power Supply N ... Present, OK」（V21 只看溫度）；
#      show sysinfo 沒有 CPU 欄位時，備援取 show cpu 的「Current CPU load」
#   4. Aruba：PSU「Warning」判為異常；「Absent」（插槽未安裝電源）不再判為異常；
#      風扇同時檢查 Fan tray（ready / fault / empty）
#
# V21 變更摘要：
#   1. 移除 pandas 依賴：Excel 改由 openpyxl 直接寫入，輸出內容與格式不變，exe 體積大幅縮小
#   2. 程式註解中的範例資料全部改為假資料（序號 / MAC / IP / 主機名稱）
#
# V20 變更摘要：
#   1. 新增 FortiGate 解析器：.conf 與 CLI 輸出皆可，HA 叢集兩台各一列（依序號跨檔去重）
#   2. Aruba 加強：VSF 堆疊成員、映像版本、Uptime、重開機原因、最高溫、Port 數、NTP、Core Dump
#   3. Excel：偵測到 Aruba / FortiGate 時，自動多建「Aruba Info」「Fortigate Info」分頁
#   4. 副檔名新增 .conf（FortiGate 備份檔）
#
# V19 變更摘要：
#   1. C9800 HA SSO 判定
#      - 依 show chassis（* = 本機）、show redundancy states、redun-management 判定 HA
#      - HA 兩台機箱各佔一列，Model / Serial / MAC 各自對應正確機箱，新增「HA Role」欄
#      - Active 與 Standby(-stby) 兩份 Log 同時存在時，依機箱 MAC 自動去重合併
#      - OOB IP 改取 Wireless Management Interface（WMI）主要 IP，
#        不再誤抓 RMI 使用的 secondary IP 或其他 VLAN
#   2. 架構重構與效能
#      - CommandIndex：每份 Log 只掃描一次建立「指令 → 輸出」索引，取代每次全文 Regex
#      - 指令比對支援 Cisco CLI 縮寫（sh ver / show env all / show process cpu sorted）
#      - 提示符錨定行首，消除超長單行（Aruba 41 萬字元）造成的 O(n²) Regex 回溯
#      - 解析引擎與 GUI 分離：可 CLI 批次執行、可單元測試；tkinter 改為選用
#   3. Excel：AP Info / WLC Info / Switch Info 最後一欄新增「Source File」（來源檔名）
#   4. 其他修正
#      - AireOS 判定改看行首提示符（V18 會因 CDP 內容含 "RTOS Version" 把 C2960X 誤判而漏列）
#      - IOS / NX-OS 環境指令為縮寫（show env all / show env）時 V18 讀不到，一律顯示正常
#      - NX-OS mgmt0 MAC 支援 dot 格式（V18 一律 N/A）
#      - 介面 IP 逐介面解析，不再跨介面誤抓；新增「檔名 IP」與「預設路由網段」判斷管理 IP
#      - Tk 對話框改回主執行緒呼叫（V18 在背景執行緒呼叫 filedialog，有機率當掉）
#   5. Catalyst（IOS-XE Switch / C9800）Memory 欄位只取一個來源，優先序：
#      show processes memory platform sorted → show processes memory sorted
#      → 兩者皆無時才用 show processes memory | i Pool Total（舊款 IOS 備援）
#   6. System Log 保留時間戳記（IOS / NX-OS / AireOS / Aruba / C1300）
#      - 同一訊息重複發生時合併為一筆，顯示「最後一次發生」的時間
#      - 單格上限 32,000 字元（Excel 儲存格上限 32,767），超過時捨棄最舊的紀錄
# =============================================================================


# ── 標準函式庫 ────────────────────────────────────────────────────────────────
import argparse
import ipaddress
import os
import queue
import re
import sys
import threading
import time
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from typing import Callable, Dict, Iterable, List, Optional, Tuple

# ── 第三方函式庫 ──────────────────────────────────────────────────────────────
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ── GUI（選用）：沒有 tkinter 的環境仍可使用 CLI 模式 ─────────────────────────
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:                                   # pragma: no cover
    tk = filedialog = messagebox = ttk = None


# ==============================================================================
# 1. 全域常數
# ==============================================================================

APP_VERSION = "V23"
MARK = "○"                       # Stack / HA 非首列「與首列共用」標記

# ── Syslog：擷取等級 0~4（保留時間戳記），並排除雜訊 ─────────────────────────
# 時間戳記在「%FACILITY-SEV-MNEMONIC:」之前，格式依平台而異：
#   IOS / IOS-XE : 012025: Dec 20 10:36:00.054 GMT: %PARSER-3-URLOPENFAIL: ...
#   NX-OS        : 2023 Aug 21 09:53:52 N5K-CORE-01 %AUTHPRIV-3-SYSTEM_MSG: ...
#   AireOS       : *Dot1x_NW_MsgTask_6: Nov 26 11:23:22.118: %DOT1X-3-...: ...
SYSLOG_REGEX = re.compile(r"%[A-Z0-9_]+-[0-4]-[A-Z0-9_]+:[^\n]*")
TIMESTAMP_RE = re.compile(r"\d{1,2}:\d{2}:\d{2}")   # 行首前綴含 hh:mm:ss 才視為時間戳記
NXOS_VDC_TAG_RE = re.compile(r"%\$\s*VDC-\d+\s*%\$\s*")   # NX-OS nvram 紀錄的「%$ VDC-1 %$」雜訊
SYSLOG_MAX_CHARS = 32000         # Excel 單一儲存格上限 32,767 字元，超過會開檔報錯
EXCLUDE_LOG_PREFIXES = (
    "%LINK-3-UPDOWN",
    "%DAEMON-3-SYSTEM_MSG",
    "%L2FM-4-L2FM_MAC_MOVE2",
    "%URIB-4-SYSLOG_SL_MSG_WARNING",
    "%SW_MATM-4-MACFLAP_NOTIF",
    "%SMART_LIC-3-COMM_FAILED",
)
SYSLOG_KEEP = 200                # 每台設備保留最後 N 筆不重複 Syslog
ARUBA_SYSLOG_KEEP = 50

# C1300 Syslog 格式：DD-Mon-YYYY HH:MM:SS :%FAC-(E|W)-MNE: msg（只收 Error / Warning）
# group(1) = 時間戳記、group(2) = 訊息
C1300_SYSLOG_REGEX = re.compile(
    r"(\d{2}-\w{3}-\d{4}\s+\d{2}:\d{2}:\d{2})\s+:(%[A-Z0-9_]+-[EW]-[A-Z0-9_]+:.*)"
)
C1300_EXCLUDE_PREFIXES = (
    "%LINK-W-Down", "%LINK-I-Up", "%STP-W-PORTSTATUS", "%TIMEMANAGE-",
    "%MLDP-I-", "%Entity-I-", "%INIT-I-",
)

# ── 環境異常關鍵字 ────────────────────────────────────────────────────────────
# FAIL / FAULT 採字首比對：Failure、Failed、Faulty 也算（Cisco 各平台用字不一）
ENV_BAD_REGEX = re.compile(
    r"\b(FAIL\w*|FAULT\w*|ALARM|CRITICAL|BAD|OFF|NOT\s+OK|DOWN|ERROR|ABSENT)\b", re.I
)
# C1300 不含 OFF：PoE 埠未接設備時「Oper: Off」屬正常
ENV_BAD_REGEX_C1300 = re.compile(
    r"\b(FAIL\w*|FAULT\w*|ALARM|CRITICAL|BAD|NOT\s+OK|ERROR|ABSENT)\b", re.I
)
# NX-OS 另含 Shutdown：電源拔線 / 斷電時 Status 顯示 Shutdown（N9K），風扇 / 電源缺件為 failure or missing
ENV_BAD_REGEX_NXOS = re.compile(
    r"\b(FAIL\w*|FAULT\w*|ALARM|CRITICAL|BAD|OFF|NOT\s+OK|DOWN|ERROR|ABSENT|SHUTDOWN)\b", re.I
)
FAN_WORD_RE = re.compile(r"\bFANS?\b|Fan", re.I)
PWR_WORD_RE = re.compile(r"\bPS-?\d|\bPWR|POWER|PSU|SUPPLY", re.I)
WLC_FAN_SPEED_RE = re.compile(r"Fan Speed\s+\d+%", re.I)   # C9800 風扇轉速狀態，非異常

# ── Log 區塊標頭（CommandIndex 用）────────────────────────────────────────────
# IOS / NX-OS / C9800：提示符「hostname#」或「hostname>」錨定行首（動態產生，見 build_prompt_index）
PROMPT_CANDIDATE_RE = re.compile(r"^([A-Za-z0-9][\w.\-/:]{0,62})[#>]", re.M)
# show tech 區段：「------------------ show version ------------------」（IOS 與 C1300 共用）
TECH_SECTION_RE = re.compile(r"^-{10,}[ \t]+(?P<cmd>\S[^\n]*?)[ \t]+-{10,}[ \t]*$", re.M)
# AireOS：「(Cisco Controller) >show sysinfo」
AIREOS_PROMPT_RE = re.compile(r"^\([^)\n]{1,40}\)[ \t]*>[ \t]*(?P<cmd>[^\n]*)$", re.M)
# Aruba AOS-CX show tech：「*****\nCommand : show system\n*****」
ARUBA_SECTION_RE = re.compile(
    r"^\*{5,}[ \t]*\nCommand[ \t]*:[ \t]*(?P<cmd>[^\n]*?)[ \t]*\n\*{5,}[ \t]*$", re.M
)
# 指令執行錯誤的輸出（視為空輸出，繼續找下一個同名指令）
CLI_ERROR_RE = re.compile(
    r"\A\s*(?:\^\s*\n\s*)?(?:%\s*(?:Invalid input|Ambiguous command|Incomplete command|"
    r"Unrecognized command)|Incorrect usage|This command cannot be executed)", re.I
)

NXOS_BANNER = "Cisco Nexus Operating System (NX-OS) Software"
# FortiGate 提示符：「FGT-HQ # get system status」或 VDOM 內「FGT-HQ (global) # …」
FGT_PROMPT_RE = re.compile(r"^([A-Za-z0-9][\w.\-]{0,62})(?: \([\w\-]+\))? [#$] ", re.M)
NO_MATCH_RE = re.compile(r"\A(?!)(?P<cmd>)")   # 沒有任何提示符時用的空索引
DETECT_WINDOW = 2_000_000        # 類型 / 提示符偵測只看前 2 MB（提示符一定在開頭附近出現）
IPV4 = r"\d{1,3}(?:\.\d{1,3}){3}"

# ── 介面設定解析 ──────────────────────────────────────────────────────────────
IFACE_BLOCK_RE = re.compile(r"^interface[ \t]+(\S+)[^\n]*\n((?:[ \t][^\n]*(?:\n|\Z))*)", re.M)
IP_LINE_RE = re.compile(
    rf"^[ \t]+ip address[ \t]+({IPV4})(?:/(\d{{1,2}})|[ \t]+({IPV4}))?([^\n]*)", re.M
)
MGMT_IFACE_RE = re.compile(r"^(?:mgmt\d*|Management\d*|GigabitEthernet0(?:/0)?|gi0)$", re.I)

# ── Excel 報表欄位 ────────────────────────────────────────────────────────────
WLC_TYPES = {"WLC", "AireOS-WLC"}
SWITCH_TYPES = {"Switch", "NX-OS", "C1300", "Aruba"}

REPORT_COLS = ["Hostname", "Model", "OOB IP", "Serial Number", "Firmware Version",
               "CPU Usage", "Memory Usage", "Fan Status", "Power Status", "System Log"]
SWITCH_COLS = ["Hostname", "Model", "OOB IP", "Ethernet MAC", "Serial Number",
               "Firmware Version", "CPU Usage", "Memory Usage", "Fan Status",
               "Power Status", "System Log", "Source File"]
WLC_COLS = SWITCH_COLS[:1] + ["HA Role"] + SWITCH_COLS[1:]
AP_COLS = ["AP Name", "AP Model", "Ethernet MAC", "IP Address", "AP Serial Number",
           "Software Version", "Source File"]
# 偵測到設備才會建立的廠牌專屬分頁
ARUBA_COLS = ["Hostname", "VSF Role", "Part Number", "Description", "Serial Number", "OOB IP",
              "Ethernet MAC", "Firmware Version", "Active Image", "Primary Image", "Secondary Image",
              "Uptime", "Last Reboot", "CPU Usage", "Memory Usage", "Max Temperature", "Fan Status",
              "Power Status", "Ports Up/Total", "NTP Status", "Core Dumps", "System Log", "Source File"]
FORTIGATE_COLS = ["Hostname", "HA Role", "Model", "Serial Number", "OOB IP", "Firmware Version",
                  "Operation Mode", "Uptime", "Last Reboot", "CPU Usage", "Memory Usage", "Sessions",
                  "Fan Status", "Power Status", "HA Status", "System Log", "Source File"]
VENDOR_SHEETS = (("Aruba Info", "Aruba", ARUBA_COLS), ("Fortigate Info", "FortiGate", FORTIGATE_COLS))
ACTIVE_ROLES = ("active", "primary", "master", "conductor")   # HA / 堆疊中排第一列的角色


# ==============================================================================
# 2. 公用輔助函式
# ==============================================================================

def rsearch(pattern: str, text: str, group: int = 1, flags: int = re.I | re.M,
            default: Optional[str] = None) -> Optional[str]:
    """re.search 簡寫：找到回傳指定 group（去首尾空白），找不到回傳 default。"""
    m = re.search(pattern, text, flags)
    if not m or m.group(group) is None:
        return default
    return m.group(group).strip()


def read_log(path: str) -> str:
    """讀取 Log：utf-8-sig 自動去除 BOM，errors='ignore' 略過無法解碼字元，換行統一為 \\n。"""
    with open(path, "r", encoding="utf-8-sig", errors="ignore") as f:
        return f.read()


def merge_syslog(entries: Iterable[Tuple[str, str]], keep: int,
                 exclude: Tuple[str, ...] = ()) -> str:
    """
    合併 Syslog 並保留時間戳記。entries 為 (時間戳記, 訊息)，須依「舊 → 新」順序傳入。
      1. 排除雜訊前綴（exclude）
      2. 相同訊息只留一筆，顯示最後一次發生的時間並移到最後
         （同一事件持續發生時不會洗版，例：NX-OS Leaf 865 行 → 98 筆）
      3. 保留最新 keep 筆，且總長度不超過 SYSLOG_MAX_CHARS（超過時捨棄最舊的）
    """
    latest: Dict[str, str] = {}
    for ts, msg in entries:
        msg, ts = msg.strip(), ts.strip()
        if not msg or msg.startswith(exclude):
            continue
        latest.pop(msg, None)
        latest[msg] = f"{ts} {msg}" if ts else msg

    lines, size = [], 0
    for line in reversed(list(latest.values())[-keep:]):     # 由新到舊累加，直到字數上限
        size += len(line) + 1
        if size > SYSLOG_MAX_CHARS:
            break
        lines.append(line)
    return "\n".join(reversed(lines))


def truncate_ap_firmware(version: str) -> str:
    """AP 版本只取前三段：17.9.8.5 → 17.9.8；N/A 等非版本字串原樣回傳。"""
    m = re.match(r"(\d+\.\d+\.\d+)", version)
    return m.group(1) if m else version


def mac_key(mac: str) -> str:
    """MAC 正規化為 12 位小寫十六進位（去重 / 比對用），無效回傳空字串。"""
    h = re.sub(r"[^0-9a-fA-F]", "", mac or "").lower()
    return h if len(h) == 12 else ""


def mac_format(mac: str, upper: bool = True) -> str:
    """任意格式 MAC（aabb.cc00.0002 / aabbcc-000001）轉為冒號格式 AA:BB:CC:00:00:02。"""
    h = mac_key(mac)
    if not h:
        return mac or "N/A"
    s = ":".join(h[i:i + 2] for i in range(0, 12, 2))
    return s.upper() if upper else s


def ip_sort_key(ip: str) -> tuple:
    """IP 排序鍵：有效 IP 依數值排序，其餘（N/A、○）排最後。"""
    try:
        return (0, ipaddress.ip_address(ip))
    except ValueError:
        return (1, ipaddress.ip_address("0.0.0.0"))


def ips_in_text(text: str) -> List[str]:
    """擷取字串中所有 IPv4（用於檔名 IP 提示）。"""
    return re.findall(rf"(?<![\d.])({IPV4})(?![\d.])", text or "")


def title_state(state: str) -> str:
    """Redundancy 狀態正規化：'STANDBY HOT' → 'Standby Hot'。"""
    return " ".join(w.capitalize() for w in re.split(r"[\s_]+", state.strip()) if w)


def classify_env_lines(text: str, bad_re: re.Pattern = ENV_BAD_REGEX) -> Tuple[bool, bool]:
    """
    逐行判斷環境輸出：含異常關鍵字的行依內容歸類為風扇或電源異常。
    無法歸類的異常行（例如溫度）同時標記兩者，確保異常不會被漏報。
    回傳 (fan_bad, pwr_bad)。
    """
    fan_bad = pwr_bad = False
    for line in text.splitlines():
        if not bad_re.search(line):
            continue
        is_fan, is_pwr = bool(FAN_WORD_RE.search(line)), bool(PWR_WORD_RE.search(line))
        fan_bad |= is_fan or not is_pwr
        pwr_bad |= is_pwr or not is_fan
    return fan_bad, pwr_bad


# ==============================================================================
# 3. 指令索引（CommandIndex）
# ==============================================================================

class CommandIndex:
    """
    把整份 Log 一次切成「指令 → 輸出」索引。

    V18：每取一個指令就用 Regex 從頭掃描整份檔案（13 MB 的 WLC Log × 十多次）。
    V19：建立索引時只掃描一次；查詢只比對指令清單。

    指令比對採 Cisco CLI 縮寫規則（Log 中每個字須為目標指令對應字的前綴）：
        "sh ver"                  ⇒ "show version"
        "show env all"            ⇒ "show environment all"
        "show process cpu sorted" ⇒ "show processes cpu sorted"
    查詢優先序：
        1. 沒有管線（|）過濾的完整輸出優先，其次才是「show xxx | i yyy」
        2. 同條件取第一個「非空、非錯誤訊息」的輸出
    """

    def __init__(self, text: str, header_re: re.Pattern, prompt_host: str = ""):
        self.text = text
        self.prompt_host = prompt_host
        self._cache: Dict[Tuple, str] = {}
        self._entries: List[Tuple[Tuple[str, ...], bool, int, int]] = []
        heads = list(header_re.finditer(text))
        for i, m in enumerate(heads):
            end = heads[i + 1].start() if i + 1 < len(heads) else len(text)
            base, pipe, _ = m.group("cmd").lower().partition("|")
            self._entries.append((tuple(base.split()), bool(pipe), m.end(), end))

    def __len__(self) -> int:
        return len(self._entries)

    @staticmethod
    def _abbrev_match(typed: Tuple[str, ...], full: Tuple[str, ...]) -> bool:
        """typed 每個字都是 full 對應字的前綴（至少 2 字元，避免 's' 之類的歧義）。"""
        return len(typed) == len(full) and all(
            f.startswith(t) and (len(t) >= 2 or t == f) for t, f in zip(typed, full)
        )

    def get(self, *commands: str) -> str:
        """依序嘗試多個指令名稱，回傳第一個找到的輸出；皆無則回傳空字串。"""
        key = commands
        if key in self._cache:
            return self._cache[key]
        wanted = [tuple(c.lower().split()) for c in commands]
        result = ""
        for want_pipe in (False, True):
            for full in wanted:
                for toks, has_pipe, start, end in self._entries:
                    if has_pipe != want_pipe or not self._abbrev_match(toks, full):
                        continue
                    out = self.text[start:end].strip()
                    if out and not CLI_ERROR_RE.match(out):
                        result = out
                        break
                if result:
                    break
            if result:
                break
        self._cache[key] = result
        return result


def build_prompt_index(text: str) -> Optional[CommandIndex]:
    """
    自動找出 CLI 提示符（出現最多次的「行首 hostname#」）並建立索引。
    只認行首提示符，所以 PuTTY 貼上的「明文指令清單」（前面沒有提示符）不會被誤認。
    """
    # 提示符在 Log 開頭就會大量出現，只統計前 2 MB（10 MB 檔案約省 30 ms）；找不到才掃全文
    counts = (Counter(PROMPT_CANDIDATE_RE.findall(text, 0, DETECT_WINDOW))
              or Counter(PROMPT_CANDIDATE_RE.findall(text)))
    if not counts:
        return None
    host = counts.most_common(1)[0][0]
    header = re.compile(rf"^{re.escape(host)}[#>][ \t]*(?P<cmd>[^\n]*)", re.M)
    return CommandIndex(text, header, prompt_host=host)


# ==============================================================================
# 4. 資料容器
# ==============================================================================

@dataclass
class Unit:
    """一台實體設備：單機 / Stack 成員 / HA 機箱。"""
    member_id: int
    model: str = "Unknown"
    serial: str = "Unknown"
    eth_mac: str = "N/A"
    ha_state: str = ""            # HA 狀態（Active / Standby Hot…），非 HA 為空字串
    rmi_ip: str = ""              # C9800 RMI IP（redun-management）
    is_local: bool = True         # False = HA 對端機箱（本 Log 看不到它的 CPU / Memory）
    extra: Dict[str, str] = field(default_factory=dict)   # 該台專屬欄位（例：VSF Role），覆蓋設備層級值


@dataclass
class DeviceInfo:
    """單份 Log 解析出的設備層級資訊（Stack / HA 成員共用）。"""
    hostname: str = "Unknown"
    oob_ip: str = "N/A"
    firmware: str = "N/A"
    cpu: str = "N/A"
    memory: str = "N/A"
    device_type: str = "Unknown"
    eth_mac: str = "N/A"
    ha_role: str = ""             # 非 C9800 HA 的 WLC 顯示用（例：Standalone）
    is_stack: bool = False
    is_ha: bool = False
    unit_count: int = 1


@dataclass
class IfaceIP:
    """running-config 中一筆介面 IP 設定。"""
    iface: str
    ip: str
    prefix: int                   # 0 = 未知
    secondary: bool


def parse_interface_ips(config: str) -> List[IfaceIP]:
    """
    逐「interface 區塊」解析 IP（V18 用 re.S 的 .*? 會跨介面誤抓下一個介面的 IP）。
    同名介面只取第一次出現（Log 內可能同時有 running / startup / show tech 多份設定）。
    """
    result, seen = [], set()
    for m in IFACE_BLOCK_RE.finditer(config):
        name = m.group(1)
        if name.lower() in seen:
            continue
        seen.add(name.lower())
        for ip, cidr, mask, rest in IP_LINE_RE.findall(m.group(2)):
            prefix = int(cidr) if cidr else 0
            if mask:
                try:
                    prefix = ipaddress.IPv4Network(f"0.0.0.0/{mask}").prefixlen
                except ValueError:
                    prefix = 0
            result.append(IfaceIP(name, ip, prefix, "secondary" in rest))
    return result


# ==============================================================================
# 5. 解析器基礎類別（BaseParser）
# ==============================================================================

class BaseParser:
    """
    所有解析器的共用父類別（Template Method）。

    run_all() 固定解析順序；各平台子類別只覆寫格式不同的步驟。
    指令輸出一律透過 self.cmd() 取得（CommandIndex，支援縮寫與 show tech 備援）。
    """

    DEVICE_TYPE = "Unknown"
    TECH_FALLBACK = True          # 找不到提示符指令時，改找 show tech 區段
    SYSLOG_NEWEST_FIRST = False   # show logging 排列順序：IOS / NX-OS 為舊 → 新；AireOS 為新 → 舊

    def __init__(self, content: str, source: str = "", index: Optional[CommandIndex] = None):
        self.content = content
        self.source = source                      # 來源檔名（Excel Source File）
        self.info = DeviceInfo(device_type=self.DEVICE_TYPE)
        self.units: List[Unit] = []
        self.fan_status = "正常"
        self.pwr_status = "正常"
        self.syslog = ""
        self.extra: Dict[str, str] = {}           # 廠牌專屬欄位（Aruba Info / Fortigate Info）
        self.idx = index or self._build_index()
        self._tech_idx: Optional[CommandIndex] = None
        self.ver_sec = self.cmd("show version")
        self.inv_sec = self.cmd("show inventory")

    # ── 指令區塊存取 ─────────────────────────────────────────────────────────
    def _build_index(self) -> CommandIndex:
        return build_prompt_index(self.content) or CommandIndex(self.content, TECH_SECTION_RE)

    def cmd(self, *commands: str) -> str:
        """取指令輸出：先查提示符索引，再查 show tech 區段（延遲建立）。"""
        out = self.idx.get(*commands)
        if out or not self.TECH_FALLBACK:
            return out
        if self._tech_idx is None:
            self._tech_idx = CommandIndex(self.content, TECH_SECTION_RE)
        return self._tech_idx.get(*commands)

    # ── 解析流程 ─────────────────────────────────────────────────────────────
    def run_all(self) -> "BaseParser":
        for step in (self._parse_hostname, self._parse_oob, self._parse_firmware,
                     self._parse_cpu, self._parse_memory, self._parse_inventory,
                     self._parse_env, self._parse_syslog, self._parse_eth_mac, self._parse_extra):
            step()
        self.info.unit_count = max(len(self.units), 1)
        return self

    # ── Hostname：running-config → 提示符 ────────────────────────────────────
    def _parse_hostname(self):
        self.info.hostname = (rsearch(r"^hostname\s+(\S+)", self.content)
                              or self.idx.prompt_host or "Unknown")

    # ── OOB IP：依序嘗試各策略，第一個成功者為準 ────────────────────────────
    def _parse_oob(self):
        cfg = self.cmd("show running-config") or self.content
        ips = parse_interface_ips(cfg)
        for strategy in self._oob_strategies():
            ip = strategy(ips, cfg)
            if ip:
                self.info.oob_ip = ip
                return

    def _oob_strategies(self) -> tuple:
        return (self._oob_mgmt_port, self._oob_filename_hint, self._oob_default_gateway,
                self._oob_ip_brief, self._oob_first_svi)

    @staticmethod
    def _oob_mgmt_port(ips: List[IfaceIP], cfg: str) -> Optional[str]:
        """A. 專用管理埠（Gi0/0、mgmt0…），排除 /30、/31 P2P 位址。"""
        for x in ips:
            if MGMT_IFACE_RE.match(x.iface) and not x.secondary and x.prefix < 30:
                return x.ip
        return None

    def _oob_filename_hint(self, ips: List[IfaceIP], cfg: str) -> Optional[str]:
        """B. 檔名中的 IP（工程師實際連線的位址）若有設定在設備上，即為管理 IP。"""
        configured = {x.ip for x in ips}
        return next((ip for ip in ips_in_text(self.source) if ip in configured), None)

    @staticmethod
    def _oob_default_gateway(ips: List[IfaceIP], cfg: str) -> Optional[str]:
        """C. 預設閘道 / 預設路由的下一跳所在網段的介面 IP（以子網路計算，不再只比前三段）。"""
        gw = (rsearch(rf"^ip default-gateway[ \t]+({IPV4})", cfg)
              or rsearch(rf"^ip route[ \t]+(?:vrf[ \t]+\S+[ \t]+)?0\.0\.0\.0(?:/0|[ \t]+0\.0\.0\.0)[ \t]+({IPV4})", cfg))
        if not gw:
            return None
        try:
            gw_ip = ipaddress.ip_address(gw)
            for x in ips:
                if x.prefix and gw_ip in ipaddress.ip_interface(f"{x.ip}/{x.prefix}").network:
                    return x.ip
        except ValueError:                                  # 設定中有非法位址時略過此策略
            pass
        return None

    def _oob_ip_brief(self, ips: List[IfaceIP], cfg: str) -> Optional[str]:
        """D. show ip interface brief 第一個有效 Vlan IP。"""
        for line in self.cmd("show ip interface brief").splitlines():
            if "vlan" in line.lower() and "unassigned" not in line:
                ip = rsearch(rf"({IPV4})", line)
                if ip and ip != "0.0.0.0":
                    return ip
        return None

    @staticmethod
    def _oob_first_svi(ips: List[IfaceIP], cfg: str) -> Optional[str]:
        """E. running-config 第一個 Vlan SVI 的 IP（最後備案）。"""
        return next((x.ip for x in ips if x.iface.lower().startswith("vlan")), None)

    # ── 共用：CPU 五分鐘平均 ─────────────────────────────────────────────────
    def _cpu_five_min(self) -> Optional[str]:
        sec = self.cmd("show processes cpu sorted", "show processes cpu")
        val = rsearch(r"five minutes:\s*(\d+)%", sec) or rsearch(r"five minutes:\s*(\d+)%", self.content)
        return f"{val}%" if val else None

    # ── 共用：Catalyst（IOS-XE Switch / C9800）記憶體 ───────────────────────
    PLATFORM_MEM_RE = re.compile(r"System memory:\s*(\d+)K total,\s*(\d+)K used,\s*(\d+)K free", re.I)
    POOL_MEM_RE = re.compile(r"Processor Pool Total:\s*(\d+)\s+Used:\s*(\d+)(?:\s+Free:\s*(\d+))?", re.I)

    def _catalyst_memory(self) -> Optional[str]:
        """
        依優先序只取「一個」來源，取到即回傳：
          1. show processes memory platform sorted（縮寫 show process memory platform sorted 亦可）
             System memory: 15838712K total, 4801984K used, 11036728K free
          2. show processes memory sorted
             Processor Pool Total: 1673524924 Used: 486160144 Free: 1187364780
          3. 備援：以上兩個指令都沒收集時（例：C2960 / C1000 / C9200L 的
             「show processes memory | i Pool Total」），沿用 Processor Pool 格式，避免整欄變 N/A
        """
        m = self.PLATFORM_MEM_RE.search(self.cmd("show processes memory platform sorted"))
        if m:
            t, u, f = (int(x) for x in m.groups())
            return self.format_mem(f"{t}K", f"{u}K", f"{f}K", round(u / t * 100, 1) if t else 0)

        m = (self.POOL_MEM_RE.search(self.cmd("show processes memory sorted"))
             or self.POOL_MEM_RE.search(self.cmd("show processes memory") or self.content))
        if m:
            t, u = int(m.group(1)), int(m.group(2))
            f = int(m.group(3)) if m.group(3) else t - u
            return self.format_mem(t, u, f, round(u / t * 100, 1) if t else 0)
        return None

    # ── 預設實作（子類別覆寫）───────────────────────────────────────────────
    def _parse_firmware(self): pass
    def _parse_cpu(self): pass
    def _parse_memory(self): pass
    def _parse_inventory(self): pass
    def _parse_env(self): pass
    def _parse_extra(self): pass

    def _parse_syslog(self):
        """
        等級 0~4 Syslog（保留時間戳記）：
          先用「%」開頭的 Regex 快速定位訊息，再用 rfind 往回取同一行的行首前綴；
          前綴含 hh:mm:ss 才當作時間戳記（排除設定檔等非 Log 行的雜訊前綴）。
        """
        text, entries = self.content, []
        for m in SYSLOG_REGEX.finditer(text):
            prefix = NXOS_VDC_TAG_RE.sub("", text[text.rfind("\n", 0, m.start()) + 1:m.start()])
            entries.append((prefix if TIMESTAMP_RE.search(prefix) else "", m.group(0)))
        if self.SYSLOG_NEWEST_FIRST:
            entries.reverse()                         # merge_syslog 需要「舊 → 新」
        self.syslog = merge_syslog(entries, SYSLOG_KEEP, EXCLUDE_LOG_PREFIXES)

    def _parse_eth_mac(self):
        """IOS / IOS-XE：show version 的 Base Ethernet MAC Address。"""
        mac = rsearch(r"Base\s+Ethernet\s+MAC\s+Address\s*[:\s]+([0-9a-fA-F:.]{11,17})", self.content)
        if mac:
            self.info.eth_mac = mac

    @staticmethod
    def format_mem(total, used, free, pct) -> str:
        return f"Total：{total} | Used：{used} | Free：{free} ({pct}%)"

    # ── Excel 資料列 ─────────────────────────────────────────────────────────
    def _ha_role_text(self, u: Unit) -> str:
        if not u.ha_state:
            return self.info.ha_role or "Standalone"
        text = f"{u.ha_state} (Chassis {u.member_id})"
        return f"{text}｜RMI {u.rmi_ip}" if u.rmi_ip else text

    def generate_rows(self) -> List[Dict]:
        """
        每台實體設備（Unit）一列。私有欄位（底線開頭）供合併 / 排序使用，不輸出到 Excel：
          _group    同一邏輯設備（Stack 或 HA Pair）的群組鍵
          _kind     stack / ha：非首列的「○」呈現方式不同
          _order    群組內排序（HA：Active 優先）
          _ha_key   HA 機箱 MAC，用於跨檔案去重（Active log + Standby log）
          _ha_local 此列資料是否來自該機箱本身的 Log（去重時優先保留）
        """
        units = self.units
        if not units:                       # 例：AireOS 的 AP 指令片段檔，不列出設備
            return []
        if self.info.is_ha:
            kind = "ha"
            group = "HA:" + "|".join(sorted(self._unit_key(u) or self.source for u in units))
            order_of = lambda u: (0 if u.ha_state.lower().startswith(ACTIVE_ROLES) else 1, u.member_id)
        else:
            kind, group = "stack", f"FILE:{self.source}:{id(self)}"
            order_of = lambda u: (0, u.member_id)

        units = sorted(units, key=order_of)
        rows = []
        for i, u in enumerate(units):
            first = (i == 0)
            if self.info.is_ha:
                mac = u.eth_mac
            else:
                mac = u.eth_mac if u.eth_mac != "N/A" else (self.info.eth_mac if first else MARK)
            row = {
                "Hostname":         self.info.hostname,
                "HA Role":          self._ha_role_text(u),
                "Model":            u.model,
                "OOB IP":           self.info.oob_ip,
                "Ethernet MAC":     mac,
                "Serial Number":    u.serial,
                "Firmware Version": self.info.firmware,
                "CPU Usage":        self.info.cpu if u.is_local else "N/A",
                "Memory Usage":     self.info.memory if u.is_local else "N/A",
                "Fan Status":       self.fan_status,
                "Power Status":     self.pwr_status,
                "System Log":       self.syslog if u.is_local else "",
                "Source File":      self.source,
                "_device_type":     self.info.device_type,
                "_group":           group,
                "_kind":            kind,
                "_order":           order_of(u),
                "_ha_key":          self._unit_key(u) if self.info.is_ha else "",
                "_ha_local":        u.is_local,
            }
            row.update(self.extra)
            row.update(u.extra)
            rows.append(row)
        return rows

    @staticmethod
    def _unit_key(u: Unit) -> str:
        """HA 成員的跨檔去重鍵：機箱 MAC（C9800），沒有 MAC 時用序號（FortiGate）。"""
        return mac_key(u.eth_mac) or (u.serial if u.serial not in ("Unknown", "N/A", "") else "")


# ==============================================================================
# 6. IOS / IOS-XE Switch 解析器（含 Stack）
# ==============================================================================

class IOSSwitchParser(BaseParser):
    """
    Catalyst C2960 / C1000 / C3850 / C9200 / C9300 / C9500。

    Stack 判定：show version 末端的成員表格（所有 IOS 平台皆有，C2960 不支援 show switch）
        Switch Ports Model              SW Version        SW Image
        *    1 52    C9200L-48T-4X      17.09.08          CAT9K_LITE_IOSXE
             2 52    C9200L-48T-4X      17.09.08          CAT9K_LITE_IOSXE
    成員數 > 1 即為 Stack；各成員序號從 show inventory 的 NAME:"Switch N" / NAME:"N" 取得。
    """

    DEVICE_TYPE = "Switch"
    STACK_TABLE_RE = re.compile(r"^\*?\s+(\d+)\s+\d+\s+(\S+)\s+(\S+)", re.M)

    def _parse_firmware(self):
        # 版本號由 _parse_inventory 從成員表格取得；無表格時備援取 "Version X"
        self.info.firmware = rsearch(r"\bVersion\s+([^\s,]+)", self.ver_sec, default="N/A")

    def _parse_cpu(self):
        self.info.cpu = self._cpu_five_min() or "N/A"

    def _parse_memory(self):
        self.info.memory = self._catalyst_memory() or "N/A"

    def _parse_inventory(self):
        members = self.STACK_TABLE_RE.findall(self.ver_sec)     # [(sid, model, version)]
        if not members:
            self.units.append(Unit(1))
            return
        self.info.is_stack = len(members) > 1
        for sid, model, version in members:
            self.info.firmware = version                        # 所有成員版本一致
            sn = rsearch(rf'NAME:\s*"(?:Switch\s*)?{sid}".*?PID:\s*\S+.*?SN:\s*(\S+)',
                         self.inv_sec, flags=re.S | re.I)
            self.units.append(Unit(int(sid), model, sn.strip(",") if sn else "Unknown"))

    def _parse_env(self):
        """
        show environment all（常見縮寫 show env all）逐行判斷。
        C1000 只有「Power Supply Status: Good」聲明行，出現即視為電源正常。
        """
        sec = self.cmd("show environment all", "show environment")
        fan_bad, pwr_bad = classify_env_lines(sec)
        if re.search(r"Power Supply Status:\s*Good", sec, re.I):
            pwr_bad = False
        self.fan_status = "異常" if fan_bad else "正常"
        self.pwr_status = "異常" if pwr_bad else "正常"

    def _parse_eth_mac(self):
        """
        各 Stack 成員 MAC：show version 以「Switch 0N\\n-----」分段，
        每段有 Base Ethernet MAC + System Serial Number，以序號對應到成員。
        """
        mac_re = re.compile(r"Base\s+Ethernet\s+MAC\s+Address\s*[:\s]+"
                            r"((?:[0-9a-fA-F]{2}[:\-]){5}[0-9a-fA-F]{2}|"
                            r"[0-9a-fA-F]{4}\.[0-9a-fA-F]{4}\.[0-9a-fA-F]{4})", re.I)
        sn_map: Dict[str, str] = {}
        for block in re.split(r"\n(?=Switch\s+0*\d+\s*\n[-─]+)", self.ver_sec, flags=re.I):
            mac = mac_re.search(block)
            sn = rsearch(r"System\s+Serial\s+Number\s*[:\s]+(\S+)", block)
            if mac and sn:
                sn_map[sn.rstrip(",")] = mac.group(1)

        if not sn_map:                                           # 單機：直接取第一個 Base MAC
            m = mac_re.search(self.ver_sec)
            if m:
                self.info.eth_mac = m.group(1)
                if self.units:
                    self.units[0].eth_mac = self.info.eth_mac
            return
        for u in self.units:
            u.eth_mac = sn_map.get(u.serial.rstrip(","), u.eth_mac)
        first = min(self.units, key=lambda u: u.member_id)
        self.info.eth_mac = first.eth_mac if first.eth_mac != "N/A" else next(iter(sn_map.values()))


# ==============================================================================
# 7. NX-OS 解析器（Nexus N5K / N9K）
# ==============================================================================

class NexusParser(BaseParser):
    """Nexus：固定單機；N9K 版本為「NXOS: version」，N5K 為「system: version」。"""

    DEVICE_TYPE = "NX-OS"
    TECH_FALLBACK = False         # NX-OS show tech 格式不同，不建立 IOS show tech 備援索引

    def _oob_strategies(self) -> tuple:
        # N5K 常把管理 IP 設成 secondary（mgmt0 為 /30 P2P），插在檔名提示之後
        return (self._oob_mgmt_port, self._oob_filename_hint, self._oob_nxos_secondary,
                self._oob_default_gateway, self._oob_ip_brief, self._oob_first_svi)

    @staticmethod
    def _oob_nxos_secondary(ips: List[IfaceIP], cfg: str) -> Optional[str]:
        return next((x.ip for x in ips if x.secondary), None)

    def _parse_firmware(self):
        self.info.firmware = (rsearch(r"NXOS:\s+version\s+([^\s\[]+)", self.ver_sec)
                              or rsearch(r"system:\s+version\s+(\S+)", self.ver_sec)
                              or "Unknown")

    def _parse_cpu(self):
        """N9K：five minutes: X%；N5K 無此行 → CPU states 的 user% + kernel%。"""
        cpu = self._cpu_five_min()
        if not cpu:
            m = re.search(r"CPU states\s*:\s*([\d.]+)%\s*user,\s*([\d.]+)%\s*kernel", self.content)
            cpu = f"{round(float(m.group(1)) + float(m.group(2)), 1)}%" if m else None
        self.info.cpu = cpu or "N/A"

    def _parse_memory(self):
        m = re.search(r"Memory usage:\s+(\d+)K total,\s+(\d+)K used", self.content)
        if m:
            t, u = int(m.group(1)), int(m.group(2))
            self.info.memory = self.format_mem(f"{t}K", f"{u}K", f"{t - u}K",
                                               round(u / t * 100, 1) if t else 0)

    def _parse_inventory(self):
        m = re.search(r'NAME:\s*"Chassis".*?PID:\s*(\S+).*?SN:\s*(\S+)', self.inv_sec, re.S | re.I)
        self.units.append(Unit(1, m.group(1).strip(","), m.group(2).strip(","))
                          if m else Unit(1, "N9K-Chassis", "Unknown"))

    def _parse_env(self):
        """
        show environment 分成 Fan / Power Supply 區段，各自判斷 Status 欄（溫度區段不列入）。
        官方範例狀態：ok、failure or missing、fail/shutdown、Shutdown（N9K 電源未供電）。
        """
        section, fan_bad, pwr_bad = "", False, False
        for line in self.cmd("show environment").splitlines():
            # 區段標頭為「Fan:」或欄位標題「Fan    Model …」；資料列「Fan-1 …」「Fan1(sys_fan1)」不是標頭
            head = re.match(r"^(Fan|Power Supply|Temperature|Power Usage Summary)(?=\s*:|\s{2,}|\s*$)", line)
            if head:
                section = head.group(1)
                continue
            if ENV_BAD_REGEX_NXOS.search(line):
                psu_fan = section == "Fan" and re.match(r"\s*Fan_in_PS", line, re.I)
                fan_bad |= section == "Fan" and not psu_fan      # 電源內建風扇隨電源 Shutdown，歸類為電源
                pwr_bad |= section == "Power Supply" or bool(psu_fan)
        self.fan_status = "異常" if fan_bad else "正常"
        self.pwr_status = "異常" if pwr_bad else "正常"

    def _parse_eth_mac(self):
        """vPC local system-mac → mgmt0 的 bia（NX-OS 為 dot 格式，V18 只認冒號格式）。"""
        mac = (rsearch(r"vPC\s+local\s+system-mac\s*[:\s]+([0-9a-fA-F:.]{14,17})", self.cmd("show vpc role"))
               or rsearch(r"^mgmt0 is[^\n]*\n(?:[^\n]*\n){0,3}?[^\n]*address:\s*([0-9a-fA-F:.]{14,17})",
                          self.content))
        if mac:
            self.info.eth_mac = mac_format(mac, upper=False)


# ==============================================================================
# 8. C9800 WLC 解析器（IOS-XE，含 HA SSO）
# ==============================================================================

class WLCParser(BaseParser):
    """
    Cisco Catalyst 9800 / EWC。

    HA SSO 判定依據（Cisco 官方 HA SSO 文件）：
      show chassis             列出 HA Pair 兩台機箱的 Role / MAC，「*」為目前登入的本機
      show redundancy states   my state / peer state（ACTIVE、STANDBY HOT…）
      show redundancy          Operating Redundancy Mode = sso、Hardware Mode = Duplex
      show inventory           NAME:"Chassis 1" / "Chassis 2" 各自的 PID / SN
      redun-management …       各機箱的 RMI IP
    WMI（wireless management interface）為 HA Pair 共用的管理 IP；
    RMI 會以 secondary IP 形式出現在同一個 SVI，因此管理 IP 必須取 primary。
    """

    DEVICE_TYPE = "WLC"
    CHASSIS_ROW_RE = re.compile(
        r"^[ \t]*(\*)?[ \t]*(\d+)[ \t]+([A-Za-z]+)[ \t]+"
        r"([0-9a-fA-F]{4}\.[0-9a-fA-F]{4}\.[0-9a-fA-F]{4})\b", re.M
    )

    # ── OOB IP：WMI → Service Port → 通用策略 ────────────────────────────────
    def _oob_strategies(self) -> tuple:
        return (self._oob_wmi, self._oob_mgmt_port) + super()._oob_strategies()[1:]

    @staticmethod
    def _oob_wmi(ips: List[IfaceIP], cfg: str) -> Optional[str]:
        wmi = rsearch(r"^wireless management interface\s+(\S+)", cfg)
        if wmi:
            return next((x.ip for x in ips if x.iface.lower() == wmi.lower() and not x.secondary), None)
        return None

    def _parse_firmware(self):
        self.info.firmware = rsearch(r"Cisco IOS XE Software, Version\s+([^\s,]+)",
                                     self.ver_sec, default="Unknown")

    def _parse_cpu(self):
        self.info.cpu = self._cpu_five_min() or "N/A"

    def _parse_memory(self):
        self.info.memory = self._catalyst_memory() or "N/A"

    # ── 硬體 / HA ────────────────────────────────────────────────────────────
    def _chassis_inventory(self, chassis: List[Tuple[str, str, str, str]]) -> Dict[int, Tuple[str, str]]:
        """
        show inventory → {機箱編號: (PID, SN)}。
        C9800 為 NAME:"Chassis 1" / "Chassis 2"；EWC（AP 內建 WLC）HA 時兩筆都叫 NAME:"Chassis"，
        沒有編號 → 以 show version 的 Processor board ID 認出本機那一筆，其餘依序分給對端。
        """
        inv: Dict[int, Tuple[str, str]] = {}
        unnumbered: List[Tuple[str, str]] = []
        for n, pid, sn in re.findall(
                r'NAME:\s*"Chassis(?:\s+(\d+))?"[^\n]*\n\s*PID:\s*([^\s,"]+)[^\n]*?SN:\s*([^\s,"]*)',
                self.inv_sec):
            if n:
                inv.setdefault(int(n), (pid, sn or "Unknown"))
            else:
                unnumbered.append((pid, sn or "Unknown"))
        if not unnumbered:
            return inv
        board_id = rsearch(r"Processor board ID\s+(\S+)", self.ver_sec)
        local_num = next((int(num) for star, num, _r, _m in chassis if star), 1)
        local = next((e for e in unnumbered if e[1] == board_id), unnumbered[0])
        inv.setdefault(local_num, local)
        unnumbered.remove(local)
        for _star, num, _r, _m in chassis:
            if int(num) not in inv and unnumbered:
                inv[int(num)] = unnumbered.pop(0)
        return inv

    def _parse_inventory(self):
        chassis = self.CHASSIS_ROW_RE.findall(self.cmd("show chassis"))   # [(*, num, role, mac)]
        inv = self._chassis_inventory(chassis)

        red_states = self.cmd("show redundancy states")
        red = self.cmd("show redundancy")
        my_state = rsearch(r"^\s*my state\s*=\s*\d+\s*-\s*(.+?)\s*$", red_states)
        peer_state = rsearch(r"^\s*peer state\s*=\s*\d+\s*-\s*(.+?)\s*$", red_states)
        sso_duplex = (rsearch(r"Operating Redundancy Mode\s*=\s*(\S+)", red, default="").lower() == "sso"
                      and rsearch(r"Hardware Mode\s*=\s*(\S+)", red, default="").lower() == "duplex")
        self.info.is_ha = len(chassis) >= 2 or sso_duplex

        if not self.info.is_ha:
            pid, sn = inv[min(inv)] if inv else ("Unknown-C9800", "Unknown")
            self.units.append(Unit(1, pid, sn))
            return

        # RMI：redun-management interface Vlan17 chassis 1 address A chassis 2 address B
        rmi_line = rsearch(r"^redun-management interface\s+\S+(.*)$", self.content, default="")
        rmi = {int(n): ip for n, ip in re.findall(r"chassis\s+(\d+)\s+address\s+(\S+)", rmi_line)}
        local_model = next((pid for pid, _ in inv.values()), "Unknown-C9800")

        if not chassis:                         # 只有 redundancy 資訊：僅能列出本機
            self.units.append(Unit(1, *inv.get(1, (local_model, "Unknown")),
                                   ha_state=title_state(my_state or "HA"), rmi_ip=rmi.get(1, "")))
            return

        for star, num, role, mac in chassis:
            n, local = int(num), bool(star)
            state = my_state if local else peer_state
            pid, sn = inv.get(n, (local_model, "N/A"))    # HA SSO 兩台必須同型號
            self.units.append(Unit(
                member_id=n, model=pid, serial=sn, eth_mac=mac_format(mac),
                ha_state=title_state(state or role), rmi_ip=rmi.get(n, ""), is_local=local,
            ))

        # Standby 的管理 SVI 只顯示自己的 RMI（WMI 由 Active 持有），此時 OOB IP 改用 Active 的 RMI，
        # 避免 HA 首列（Active）掛上 Standby 的位址
        active = next((u for u in self.units if u.ha_state.lower().startswith("active")), None)
        if active and not active.is_local and active.rmi_ip and self.info.oob_ip in rmi.values():
            self.info.oob_ip = active.rmi_ip

    def _parse_env(self):
        """非 Normal 且非「Fan Speed N%」（C9800 正常轉速狀態）的行才判定為異常。"""
        f_bad = p_bad = False
        for line in self.cmd("show environment all").splitlines():
            if not line.strip() or "Normal" in line or WLC_FAN_SPEED_RE.search(line):
                continue
            f_bad |= "Fan" in line
            p_bad |= "Power" in line or "PSU" in line
        self.fan_status = "異常" if f_bad else "正常"
        self.pwr_status = "異常" if p_bad else "正常"

    # ── AP 解析 ──────────────────────────────────────────────────────────────
    def parse_ap_summary(self) -> List[Dict]:
        """
        show ap summary：AP Name 可能含空白，無法 split()；
        改用表頭各欄的字元起始位置切割每一列（Column Offset）。
        """
        lines = self.cmd("show ap summary").splitlines()
        hdr_i = next((i for i, l in enumerate(lines)
                      if re.search(r"AP\s+Name\s+.*AP\s+Model\s+.*Ethernet\s+MAC", l, re.I)), -1)
        mac_re = re.compile(r"(?:[\da-fA-F]{2}[:-]){5}[\da-fA-F]{2}|[\da-fA-F]{4}\.[\da-fA-F]{4}\.[\da-fA-F]{4}")
        results = []

        if hdr_i < 0:                               # 備援：逐行 Regex
            for line in lines:
                m = re.match(r"^(\S[\S ]*?)\s{2,}(\d+)\s+(\S+)\s+([\da-fA-F:.-]+)\s+"
                             rf"([\da-fA-F:.-]+)\s+({IPV4})", line)
                if m:
                    results.append({"ap_name": m.group(1).strip(), "ap_model": m.group(3),
                                    "eth_mac": m.group(4), "ip_address": m.group(6)})
            return results

        hdr = lines[hdr_i]
        col = lambda pat: (re.search(pat, hdr, re.I).start() if re.search(pat, hdr, re.I) else -1)
        c_slots, c_model, c_eth, c_ip = col(r"Slots"), col(r"AP\s+Model"), col(r"Ethernet\s+MAC"), col(r"IP\s+Address")

        for line in lines[hdr_i + 1:]:
            if not line.strip() or line.strip().startswith("-"):
                continue
            name = line[:c_slots if c_slots > 0 else c_model].strip()
            if not name:
                continue
            seg_model = line[c_model:c_eth] if c_eth > c_model else line[c_model:]
            seg_eth = line[c_eth:c_ip] if c_ip > c_eth else line[c_eth:]
            mac = mac_re.search(seg_eth)
            ip = re.search(IPV4, line[c_ip:]) if c_ip > 0 else None
            results.append({
                "ap_name": name,
                "ap_model": (seg_model.split() or ["Unknown"])[0],
                "eth_mac": mac.group(0) if mac else "N/A",
                "ip_address": ip.group(0) if ip else "N/A",
            })
        return results

    def parse_ap_config(self) -> Dict[str, Dict]:
        """show ap config general：以「Cisco AP Name :」切段，取 Serial 與 17.x 版本（截前三段）。"""
        results = {}
        for block in re.split(r"(?=Cisco AP Name\s*:)", self.cmd("show ap config general")):
            name = rsearch(r"\ACisco AP Name\s*:\s*(.+)", block)
            if not name:
                continue
            fw = rsearch(r"Software Version\s*:\s*(17\.\S+)", block, default="N/A")
            results[name] = {
                "serial":   rsearch(r"AP Serial Number\s*:\s*(\S+)", block, default="N/A"),
                "firmware": truncate_ap_firmware(fw),
            }
        return results

    def get_ap_rows(self) -> List[Dict]:
        """以 AP Name 合併 summary（Model / MAC / IP）與 config general（Serial / 版本）。"""
        config = self.parse_ap_config()
        return [dict(ap, serial=config.get(ap["ap_name"], {}).get("serial", "N/A"),
                     firmware=config.get(ap["ap_name"], {}).get("firmware", "N/A"))
                for ap in self.parse_ap_summary()]


# ==============================================================================
# 9. AireOS WLC 解析器（2504 / 3504 / 5520）
# ==============================================================================

class AireOSParser(BaseParser):
    """
    AireOS 提示符固定為「(名稱) >」，資料主要來自 show sysinfo：
      Hostname / OOB IP / Firmware / CPU / Fan / 溫度 / MAC / Redundancy Mode
    AP：show ap join stats summary all（只取 Joined）+ show ap inventory all（PID / SN）
    """

    DEVICE_TYPE = "AireOS-WLC"
    TECH_FALLBACK = False
    SYSLOG_NEWEST_FIRST = True    # AireOS show logging 最新的在最上面

    def __init__(self, content: str, source: str = "", index: Optional[CommandIndex] = None):
        super().__init__(content, source, index)
        self.sysinfo_sec = self.cmd("show sysinfo")

    def _build_index(self) -> CommandIndex:
        host = rsearch(r"^\(([^)\n]+)\)[ \t]*>", self.content, default="")
        return CommandIndex(self.content, AIREOS_PROMPT_RE, prompt_host=host)

    def _sys(self, field_name: str, value_re: str = r"(\S+)") -> Optional[str]:
        """show sysinfo 的「欄位名稱.......... 值」格式。"""
        return rsearch(rf"{field_name}[.\s]+{value_re}", self.sysinfo_sec)

    def _parse_hostname(self):
        self.info.hostname = self._sys("System Name") or self.idx.prompt_host or "Unknown"

    def _parse_oob(self):
        self.info.oob_ip = self._sys("IP Address", r"([\d.]+)") or "N/A"

    def _parse_firmware(self):
        self.info.firmware = self._sys("Product Version", r"([\d.]+)") or "Unknown"
        mode = self._sys("Redundancy Mode")
        self.info.ha_role = "Standalone" if not mode or mode.lower() == "disabled" else f"HA ({mode})"

    def _parse_cpu(self):
        """show sysinfo 的 CPU Average Usage；部分機型 / 版本沒有此欄，改取 show cpu 的 Current CPU load。"""
        val = self._sys("CPU Average Usage", r"(\d+)") or rsearch(
            r"Current CPU(?:\(s\))? load\s*:\s*([\d.]+)\s*%", self.cmd("show cpu"))
        self.info.cpu = f"{val}%" if val else "N/A"

    def _parse_memory(self):
        sec = self.cmd("show memory statistics")
        total = rsearch(r"Total System Memory[.:\s]+(\d+)\s*bytes", sec)
        used = rsearch(r"Used System Memory[.:\s]+(\d+)\s*bytes", sec)
        free = rsearch(r"Free System Memory[.:\s]+(\d+)\s*bytes", sec)
        if total and used:
            t, u = int(total), int(used)
            f = int(free) if free else t - u
            self.info.memory = self.format_mem(t, u, f, round(u / t * 100, 1) if t else 0)

    def _parse_inventory(self):
        m = re.search(r'NAME:\s*"Chassis".*?PID:\s*([\w-]+).*?SN:\s*(\w+)', self.inv_sec, re.S | re.I)
        if m:
            self.units.append(Unit(1, m.group(1), m.group(2)))
        elif self.info.firmware != "Unknown":
            self.units.append(Unit(1, "WLC (AireOS)", "Unknown"))
        # firmware 也是 Unknown：代表只是 AP 指令片段檔，不列出任何設備

    def _parse_env(self):
        """
        Fan Status == OK 為正常。
        電源：show sysinfo 的「Power Supply N ... Present, OK」（5520 / 8540 等雙電源機型），
              Present 但不是 OK 即異常；Absent（未安裝）不算異常。
              另沿用 V21：Internal / Mgig / External 任一溫度超過 Alarm 上限 → 電源異常。
        """
        fan = self._sys("Fan Status", r"(\w+)")
        if fan:
            self.fan_status = "正常" if fan.upper() == "OK" else "異常"
        pwr_bad = any(not re.match(r"OK\b", state.strip(), re.I) for state in
                      re.findall(r"^Power Supply \d+[.\s]+Present\s*,\s*([^\n]*)$", self.sysinfo_sec, re.M | re.I))
        for limit_key, temp_key in (("Internal Temp Alarm Limits", "Internal Temperature"),
                                    ("Mgig Temp Alarm Limits", "Mgig Temperature"),
                                    ("External Temp Alarm Limits", "External Temperature")):
            limit = self._sys(re.escape(limit_key), r"[-\d]+\s+to\s+(\d+)\s*C")
            temp = self._sys(re.escape(temp_key), r"\+?(\d+)\s*C")
            if limit and temp and int(temp) > int(limit):
                pwr_bad = True
                break
        self.pwr_status = "異常" if pwr_bad else "正常"

    def _parse_eth_mac(self):
        mac = self._sys("Burned-in MAC Address", r"([0-9a-fA-F:]{17})")
        if mac:
            self.info.eth_mac = mac

    def get_ap_rows(self) -> List[Dict]:
        """join stats（只取 Joined，Not Joined 不覆蓋已 Joined 的紀錄）+ inventory 補 PID / SN。"""
        wlc_fw = truncate_ap_firmware(self.info.firmware)     # AireOS AP 版本 = WLC 版本
        aps: Dict[str, Dict] = {}
        for line in self.cmd("show ap join stats summary all").splitlines():
            m = re.match(r"\s*([\da-f:]{17})\s+([\da-f:]{17})\s+(\S+)\s+([\d.]+)\s+(\S+)", line, re.I)
            if not m:
                continue
            name = m.group(3)
            joined = bool(re.search(r"\bJoined\b", line, re.I)) and not re.search(r"\bNot\s+Joined\b", line, re.I)
            if name in aps and aps[name]["joined"] and not joined:
                continue
            aps[name] = {"ap_name": name, "ap_model": "Unknown", "eth_mac": m.group(2),
                         "ip_address": m.group(4), "serial": "N/A", "firmware": wlc_fw, "joined": joined}

        for block in re.split(r"Inventory for ", self.cmd("show ap inventory all")):
            name = rsearch(r"\A\s*(\S+)", block)
            if name in aps:
                aps[name]["ap_model"] = rsearch(r"PID:\s*([\w-]+)", block, default=aps[name]["ap_model"])
                aps[name]["serial"] = rsearch(r"SN:\s*(\w+)", block, default=aps[name]["serial"])

        return [{k: v for k, v in ap.items() if k != "joined"} for ap in aps.values() if ap["joined"]]


# ==============================================================================
# 10. Aruba / HPE AOS-CX 解析器（6300 / 6400，含 VSF 堆疊）
# ==============================================================================

class ArubaParser(BaseParser):
    """
    AOS-CX show tech：區塊邊界為「*****\\nCommand : show X\\n*****」。

    Switch Info（與 Cisco 共用欄位）：
      Hostname / 型號 / 序號 / CPU / Memory / MAC → show system
      韌體 → show version（去除 FL. 前綴）；管理 IP → interface mgmt 的 ip static
    Aruba Info（Aruba 專屬欄位）：
      VSF Role             → show vsf（Conductor / Standby / Member / Not Present），無 VSF = Standalone
      Part Number / Description → show system 的 Product Name、show module
      Active / Primary / Secondary Image → show version、show images
      Uptime / Last Reboot → show system 的 Up Time、show boot-history 上一次開機原因
      Max Temperature      → show environment 溫度表最高值
      Ports Up/Total       → show interface brief；NTP → show ntp status；Core Dumps → show core-dump all
    """

    DEVICE_TYPE = "Aruba"
    TECH_FALLBACK = False
    # show vsf 成員表：「1 aa:bb:cc:00:00:01 JL666A Conductor」；不在位的成員沒有 MAC：「4 JL662A Not Present」
    VSF_ROW_RE = re.compile(
        r"^[ \t]*(\d+)[ \t]+(?:((?:[0-9a-f]{2}:){5}[0-9a-f]{2})[ \t]+)?([A-Z0-9]{5,7})[ \t]*([^\n]*?)[ \t]*$",
        re.M | re.I,
    )

    def __init__(self, content: str, source: str = "", index: Optional[CommandIndex] = None):
        super().__init__(content, source, index)
        self.sys_sec = self.cmd("show system")
        self.env_sec = self.cmd("show environment")
        self.cfg_sec = self.cmd("show running-config")

    def _build_index(self) -> CommandIndex:
        return CommandIndex(self.content, ARUBA_SECTION_RE)

    def _parse_hostname(self):
        self.info.hostname = rsearch(r"^Hostname\s*:\s*(\S+)", self.sys_sec, default="Unknown")

    def _parse_oob(self):
        ip = rsearch(rf"interface\s+mgmt\b.*?ip\s+static\s+({IPV4})/\d+", self.cfg_sec, flags=re.S | re.I)
        if ip:
            self.info.oob_ip = ip

    def _parse_firmware(self):
        self.info.firmware = rsearch(r"Version\s*:\s*(?:\w+\.)?([\d.]+)", self.ver_sec, flags=0, default="Unknown")
        images = self.cmd("show images")
        self.extra.update({
            "Active Image":    rsearch(r"^Active Image\s*:\s*(\S+)", self.ver_sec, default="N/A"),
            "Primary Image":   rsearch(r"Primary Image\s*\n-+\s*\nVersion\s*:\s*(\S+)", images, default="N/A"),
            "Secondary Image": rsearch(r"Secondary Image\s*\n-+\s*\nVersion\s*:\s*(\S+)", images, default="N/A"),
        })

    def _parse_cpu(self):
        val = rsearch(r"CPU Util\s*\(%\s*avg\s*5\s*min\)\s*:\s*(\d+)", self.sys_sec, flags=0)
        self.info.cpu = f"{val}%" if val else "N/A"

    def _parse_memory(self):
        val = rsearch(r"Memory Usage\s*\(%\)\s*:\s*(\d+)", self.sys_sec, flags=0)
        self.info.memory = f"Usage: {val}%" if val else "N/A"      # AOS-CX 只提供百分比

    def _parse_inventory(self):
        """
        型號：Product Name 前三段（JL667A 6300F 48G），與 V19 Switch Info 相同；
        VSF：show vsf 有兩台以上成員即視為堆疊，每台一列（序號由 show module 對應）。
        """
        name = rsearch(r"Product Name\s*:\s*([^\r\n]+)", self.sys_sec, flags=0) or ""
        parts = name.split()
        pid = " ".join(parts[:3]) if len(parts) >= 3 else (name or "Unknown-Aruba")
        sn = rsearch(r"Chassis Serial Nbr\s*:\s*(\S+)", self.sys_sec, flags=0, default="Unknown")
        self.extra.update({"Part Number": parts[0] if parts else "N/A",
                           "Description": " ".join(parts[1:]) or "N/A",
                           "VSF Role": "Standalone"})

        # show module：「1/1  JL667A  6300F 48G 4SFP56 Sw   SG00TEST01 Ready」→ 成員編號: (料號, 描述, 序號)
        modules: Dict[int, Tuple[str, str, str]] = {}
        for n, part, desc, serial, _status in re.findall(
                r"^(\d+)/\d+\s+(\S+)\s+(.+?)\s{2,}(\S+)\s+(\S+)\s*$", self.cmd("show module"), re.M):
            modules.setdefault(int(n), (part, desc.strip(), serial))

        vsf = self.cmd("show vsf")
        table = re.split(r"^-{3}[- ]*$", vsf, maxsplit=1, flags=re.M)
        members = self.VSF_ROW_RE.findall(table[1]) if len(table) > 1 else []
        if len(members) < 2:
            self.units.append(Unit(1, pid, sn))
            return

        self.info.is_stack = True
        for num, mac, sku, status in members:
            n = int(num)
            part, desc, m_sn = modules.get(n, (sku, "", "Unknown"))
            model = " ".join([part] + desc.split()[:2]) if desc else sku
            self.units.append(Unit(
                member_id=n, model=model, serial=m_sn if status.lower() != "not present" else "N/A",
                eth_mac=mac_format(mac, upper=False) if mac else "N/A",
                extra={"VSF Role": status or "Unknown", "Part Number": part,
                       "Description": desc or "N/A"},
            ))

    def _parse_env(self):
        """
        依 AOS-CX 官方說明：
          風扇 / 風扇盤：ok、ready 正常；fault 異常；empty（未安裝）、uninitialized 不算異常
          電源：OK、OK* 正常；Input Fault / Output Fault / Warning 異常；Absent（插槽未安裝）不算異常
        """
        bad = re.compile(r"\b(FAIL\w*|FAULT\w*|ALARM|ERROR|CRITICAL)\b", re.I)
        psu_bad = re.compile(r"\b(FAIL\w*|FAULT\w*|ALARM|ERROR|CRITICAL|WARNING)\b", re.I)
        fan = rsearch(r"Fan (?:tray )?information(.*?)(?=\n\nshow environment|\Z)", self.env_sec,
                      flags=re.S | re.I, default="")
        psu = rsearch(r"show environment power-supply\s*\n.*?\n.*?\n.*?\n(.*?)(?=\n\n|\Z)",
                      self.env_sec, flags=re.S | re.I, default="")
        self.fan_status = "異常" if bad.search(fan) else "正常"
        self.pwr_status = "異常" if psu_bad.search(psu) else "正常"

        # 溫度表：「1/1-CPU   management-module   45.75 C   normal」→ 取最高溫，非 normal 另外標示
        temps = re.findall(r"^(\S+)\s+\S+-module\s+([\d.]+)\s*C\s+(\S+)", self.env_sec, re.M)
        if temps:
            sensor, value, status = max(temps, key=lambda t: float(t[1]))
            abnormal = [t[0] for t in temps if t[2].lower() != "normal"]
            text = f"{float(value):.1f} C（{sensor}）"
            self.extra["Max Temperature"] = text + (f"｜異常：{', '.join(abnormal)}" if abnormal else "")

    def _parse_eth_mac(self):
        """HPE 格式「aabbcc-000001」與標準格式皆轉為 aa:bb:cc:00:00:01。"""
        raw = rsearch(r"Base\s+MAC\s+Address\s*[:\s]+([0-9a-fA-F:.]{11,17}|[0-9a-fA-F]{6}-[0-9a-fA-F]{6})",
                      self.sys_sec)
        if raw:
            self.info.eth_mac = mac_format(raw, upper=False) if mac_key(raw) else raw

    def _parse_extra(self):
        """Aruba Info 的維運狀態欄位（Uptime / 重開機原因 / Port / NTP / Core Dump）。"""
        uptime = rsearch(r"^Up Time\s*:\s*(.+?)\s*$", self.sys_sec)
        # show boot-history：第一筆非「Switch boot count」的時間 + 原因，即上一次開機原因
        reboot = re.search(r"^(\d{2} \w{3} \d{2} [\d:]+)\s*:\s*(?!Switch boot count)([^,\n]+)",
                           self.cmd("show boot-history"), re.M)
        ports = re.findall(r"^\d+/\d+/\d+\s+\S+\s+\S+\s+\S+\s+(?:yes|no)\s+(up|down)\b",
                           self.cmd("show interface brief"), re.M)
        ntp = self.cmd("show ntp status")
        ntp_state = rsearch(r"^NTP\s*:\s*(\S+)", ntp)
        synced = "Not synchronized" if re.search(r"Not synchronized", ntp, re.I) else (
            rsearch(r"(Synchronized to[^\n]*)", ntp) or "")
        cores = rsearch(r"Total number of core dumps\s*:\s*(\d+)", self.cmd("show core-dump all"))
        self.extra.update({
            "Uptime":         uptime or "N/A",
            "Last Reboot":    f"{reboot.group(1)}｜{reboot.group(2).strip()}" if reboot else "N/A",
            "Ports Up/Total": f"{ports.count('up')} / {len(ports)}" if ports else "N/A",
            "NTP Status":     "｜".join(x for x in (ntp_state, synced) if x) or "N/A",
            "Core Dumps":     cores if cores is not None else "N/A",
        })

    def _parse_syslog(self):
        """
        show events -s warn（舊 → 新）：保留行首 ISO 時間戳記，
        同事件持續發生會重複很多行，以訊息內容（Event|CODE|LEVEL|MOD|-|MSG）合併。
        """
        entries = []
        for line in self.cmd("show events -s warn").splitlines():
            line = line.strip()
            if not line or line.startswith(("---", "Event logs")):
                continue
            m = re.search(r"(Event\|\d+\|[A-Z_]+\|[^|]+\|-\|.*)", line) or re.search(r":\s+(.+)", line)
            if m:
                ts = rsearch(r"^(\d{4}-\d{2}-\d{2}T\S+)", line, default="")
                entries.append((ts, m.group(1)))
        self.syslog = merge_syslog(entries, ARUBA_SYSLOG_KEEP)


# ==============================================================================
# 10-1. FortiGate 解析器（FortiOS，含 HA 叢集）
# ==============================================================================

class FortigateParser(BaseParser):
    """
    FortiGate 支援兩種來源，可同時存在於同一份檔案：
      A. 備份設定檔 .conf：首行「#config-version=FG100F-7.2.8-FW-build1639-…」，
         config system global（hostname）、config system interface（管理 IP）
      B. CLI 輸出（提示符「FGT # 」或「FGT (global) # 」）：
         get system status            → Hostname / 型號 / 版本 / 序號 / 運作模式 / HA 模式 / 重開機原因
         get system performance status → CPU（100 − idle）/ Memory / Sessions / Uptime
         get system ha status         → HA 成員（Primary / Secondary + 序號）、健康與同步狀態
         execute sensor list          → 風扇 / 電源（alarm≠0 即異常）
         execute log display          → level 為 emergency ~ warning 的事件（對應 Syslog 0~4）
    .conf 本身沒有 CPU / 風扇 / 電源資訊，這些欄位會顯示 N/A（而不是假設「正常」）。
    """

    DEVICE_TYPE = "FortiGate"
    TECH_FALLBACK = False
    LOG_LEVELS = ("emergency", "alert", "critical", "error", "warning")
    ROLE_NAMES = {"master": "Primary", "slave": "Secondary", "primary": "Primary", "secondary": "Secondary"}

    def __init__(self, content: str, source: str = "", index: Optional[CommandIndex] = None):
        super().__init__(content, source, index)
        self.status = self.cmd("get system status")
        self.perf = self.cmd("get system performance status")
        self.ha = self.cmd("get system ha status")
        self.sensor = self.cmd("execute sensor list")
        self.model = "Unknown"

    def _build_index(self) -> CommandIndex:
        hosts = Counter(FGT_PROMPT_RE.findall(self.content, 0, DETECT_WINDOW))
        if not hosts:
            return CommandIndex(self.content, NO_MATCH_RE)
        host = hosts.most_common(1)[0][0]
        header = re.compile(rf"^{re.escape(host)}(?: \([\w\-]+\))? [#$][ \t]*(?P<cmd>[^\n]*)", re.M)
        return CommandIndex(self.content, header, prompt_host=host)

    def _cfg_block(self, name: str) -> str:
        """取頂層「config <name> … end」區塊（巢狀的 end 有縮排，不會誤判結尾）。"""
        return rsearch(rf"^config {re.escape(name)}[ \t]*\n(.*?)^end\b", self.content,
                       flags=re.M | re.S, default="")

    def _parse_hostname(self):
        self.info.hostname = (rsearch(r"^Hostname:\s*(\S+)", self.status)
                              or rsearch(r'^\s*set hostname "?([^"\n]+)"?', self._cfg_block("system global"))
                              or self.idx.prompt_host or "Unknown")

    def _parse_oob(self):
        """檔名 IP（有設定在介面上）→ mgmt 介面 → 開放 https/ssh 管理的介面 → 第一個有 IP 的介面。"""
        ifaces = []
        for name, body in re.findall(r'^\s*edit "([^"]+)"\s*\n(.*?)^\s*next\b',
                                     self._cfg_block("system interface"), re.M | re.S):
            ip = rsearch(rf"set ip ({IPV4})\s+{IPV4}", body)
            if ip and ip != "0.0.0.0":
                ifaces.append((name, ip, rsearch(r"set allowaccess (.+)", body, default="")))
        configured = {ip for _, ip, _ in ifaces}
        pick = (next((ip for ip in ips_in_text(self.source) if ip in configured), None)
                or next((ip for name, ip, _ in ifaces if name.lower().startswith("mgmt")), None)
                or next((ip for _, ip, acc in ifaces if re.search(r"\b(https|ssh)\b", acc)), None)
                or next((ip for _, ip, _ in ifaces), None))
        if pick:
            self.info.oob_ip = pick

    def _parse_firmware(self):
        m = re.search(r"^Version:\s*(Forti\S+)\s+(v[\d.]+),build(\d+)", self.status, re.M)
        if m:
            self.model, self.info.firmware = m.group(1), f"{m.group(2)} build{m.group(3)}"
        else:
            m = re.search(r"^#config-version=([A-Z0-9]+)-([\d.]+)-FW-build(\d+)", self.content, re.M)
            if m:
                self.model, self.info.firmware = m.group(1), f"v{m.group(2)} build{m.group(3)}"
        mode = rsearch(r"^Current HA mode:\s*(.+?)\s*$", self.status)
        self.info.ha_role = "Standalone" if not mode or mode.lower() == "standalone" else f"HA ({mode})"
        self.extra.update({
            "Operation Mode": rsearch(r"^Operation Mode:\s*(\S+)", self.status, default="N/A"),
            "Last Reboot":    rsearch(r"^Last reboot reason:\s*(.+?)\s*$", self.status, default="N/A"),
            "HA Status":      self.info.ha_role,
        })

    def _parse_cpu(self):
        idle = rsearch(r"^CPU states:.*?(\d+)% idle", self.perf)
        self.info.cpu = f"{100 - int(idle)}%" if idle else "N/A"

    def _parse_memory(self):
        m = re.search(r"Memory:\s*(\d+)k total,\s*(\d+)k used \(([\d.]+)%\),\s*(\d+)k free", self.perf)
        if m:
            t, u, pct, f = m.groups()
            self.info.memory = self.format_mem(f"{t}K", f"{u}K", f"{f}K", pct)
        self.extra.update({
            "Uptime":   rsearch(r"^Uptime:\s*(.+?)\s*$", self.perf, default="N/A"),
            "Sessions": (lambda s: f"{s}（1 分鐘平均）" if s else "N/A")(
                rsearch(r"^Average sessions:\s*(\d+) sessions in 1 minute", self.perf)),
        })

    def _parse_inventory(self):
        """單機一列；get system ha status 列出兩台以上成員時，每台一列（依序號跨檔去重）。"""
        serial = rsearch(r"^Serial-Number:\s*(\S+)", self.status, default="Unknown")
        if self.model == "Unknown":
            self.model = rsearch(r"^Model name:\s*(\S+)", self.cmd("get hardware status"), default="Unknown")

        members = re.findall(r"^(Primary|Secondary|Master|Slave)\s*:\s*([^,\n]+?)\s*,\s*([A-Z0-9]+)\s*,"
                             r"\s*HA cluster index\s*=\s*(\d+)", self.ha, re.M | re.I)
        if len(members) < 2:
            self.units.append(Unit(1, self.model, serial))
            return

        self.info.is_ha = True
        health = rsearch(r"^HA Health Status:\s*(.+?)\s*$", self.ha, default="N/A")
        # Configuration Status：「FG100FTK19000001(updated 3 seconds ago): in-sync」（限同一行，
        # 避免誤抓 System Usage stats 區段同樣格式、但冒號後換行的那幾行）
        sync = dict(re.findall(r"^[ \t]*([A-Z0-9]+)[ \t]*\(updated[^)\n]*\)[ \t]*:[ \t]*(in-sync|out-of-sync)\b",
                               self.ha, re.M | re.I))
        for role, _name, sn, index in members:
            local = sn == serial
            extra = {"HA Status": f"{health}｜{sync.get(sn, 'N/A')}"}
            if not local:                       # 對端的硬體感測只能從它自己的 Log 取得
                extra.update({"Fan Status": "N/A", "Power Status": "N/A"})
            self.units.append(Unit(member_id=int(index) + 1, model=self.model, serial=sn,
                                   ha_state=self.ROLE_NAMES.get(role.lower(), role.title()),
                                   is_local=local, extra=extra))

    def _parse_env(self):
        """execute sensor list：alarm≠0 或 LOST / FAIL 視為異常；沒有感測資料時顯示 N/A。"""
        if not self.sensor:
            self.fan_status = self.pwr_status = "N/A"
            return
        fan_bad = pwr_bad = False
        for line in self.sensor.splitlines():
            alarm = rsearch(r"alarm=(\d+)", line)
            if not ((alarm and alarm != "0") or re.search(r"\b(LOST|FAIL\w*|FAULT\w*)\b", line, re.I)):
                continue
            is_fan = bool(re.search(r"fan", line, re.I))
            is_pwr = bool(re.search(r"\bPS\d|PSU|VIN|VOUT|POWER|PWR", line, re.I))
            fan_bad |= is_fan or not is_pwr
            pwr_bad |= is_pwr or not is_fan
        self.fan_status = "異常" if fan_bad else "正常"
        self.pwr_status = "異常" if pwr_bad else "正常"

    def _parse_syslog(self):
        """
        execute log display 的 key=value 事件，只收 emergency ~ warning：
          2024-05-01 10:00:00 [warning] logdesc: msg
        FortiOS 預設由新到舊列出，因此先依日期時間排序（ISO 格式可直接字串排序）再合併。
        """
        entries = []
        for m in re.finditer(r"^[^\n]*?\bdate=(\d{4}-\d{2}-\d{2})\s+time=(\d{2}:\d{2}:\d{2})[^\n]*",
                             self.content, re.M):
            line = m.group(0)
            level = rsearch(r'\blevel="?(\w+)"?', line, default="").lower()
            if level not in self.LOG_LEVELS:
                continue
            desc = rsearch(r'\blogdesc="([^"]*)"', line, default="")
            msg = rsearch(r'\bmsg="([^"]*)"', line, default="")
            text = f"[{level}] " + (": ".join(x for x in (desc, msg) if x) or line.strip())
            entries.append((f"{m.group(1)} {m.group(2)}", text))
        entries.sort(key=lambda e: e[0])
        self.syslog = merge_syslog(entries, SYSLOG_KEEP)

    def _parse_eth_mac(self):
        pass                                     # FortiOS 預設指令無機身 MAC，維持 N/A

    def _ha_role_text(self, u: Unit) -> str:
        return u.ha_state or self.info.ha_role or "Standalone"


# ==============================================================================
# 11. Catalyst 1300 解析器（Linux-based，show tech-support 格式）
# ==============================================================================

class C1300Parser(BaseParser):
    """
    C1300 以一條 show tech-support 彙整所有指令，區塊標頭為「-------- show X --------」。
      Hostname → show system（System Name）       OOB IP → show ip interface（vlan N UP/UP）
      Firmware → show version（Active-image）     Model / SN → show inventory（NAME:"N"）
      CPU → show cpu utilization                  Memory → memory info（Dynamic RAM）
      Fan / Power → show system 的 Fans / Temperature Status
      Stack → show stack 的 Unit Id 行數
    """

    DEVICE_TYPE = "C1300"
    TECH_FALLBACK = False

    def __init__(self, content: str, source: str = "", index: Optional[CommandIndex] = None):
        super().__init__(content, source, index)
        self.sys_sec = self.cmd("show system")
        self.cpu_sec = self.cmd("show cpu utilization")
        self.mem_sec = self.cmd("memory info")
        self.log_sec = self.cmd("show logging")
        self.ip_sec = self.cmd("show ip interface")
        self.stk_sec = self.cmd("show stack")

    def _build_index(self) -> CommandIndex:
        return CommandIndex(self.content, TECH_SECTION_RE)

    def _parse_hostname(self):
        self.info.hostname = (rsearch(r"System Name:\s+(\S+)", self.sys_sec)
                              or rsearch(r"^hostname\s+(\S+)", self.content) or "Unknown")

    def _parse_oob(self):
        ip = rsearch(rf"({IPV4})/\d+\s+vlan\s+\d+\s+UP/UP", self.ip_sec)
        if ip:
            self.info.oob_ip = ip

    def _parse_firmware(self):
        self.info.firmware = rsearch(r"Active-image:.*?\n\s*Version:\s+(\S+)", self.ver_sec,
                                     flags=re.S, default="Unknown")

    def _parse_cpu(self):
        val = rsearch(r"five minutes:\s*(\d+)%", self.cpu_sec)
        self.info.cpu = f"{val}%" if val else "N/A"

    def _parse_memory(self):
        m = re.search(r"Dynamic\s*\(OS managed\)\s*RAM\s*usage:.*?Requested\s*=\s*(\d+),\s*Free\s*=\s*(\d+),"
                      r"\s*Used\s*=\s*(\d+),\s*Usage\s*=\s*(\d+)%", self.mem_sec, re.S | re.I)
        if m:
            total, free, used, pct = m.groups()
            self.info.memory = self.format_mem(total, used, free, pct)

    def _parse_inventory(self):
        """主機箱 NAME:"1"；show stack 的 Unit 行數 > 1 即為 Stack。"""
        def inv_of(uid: int) -> Optional[Tuple[str, str]]:
            m = re.search(rf'NAME:\s*"{uid}".*?PID:\s*(\S+).*?SN:\s*(\S+)', self.inv_sec, re.S | re.I)
            return (m.group(1).strip(), m.group(2).strip()) if m else None

        pid, sn = inv_of(1) or ("Unknown-C1300", "Unknown")
        rows = re.findall(r"^\s+(\d+)\s+([\da-f:]{17})\s+(\S+)", self.stk_sec, re.M | re.I)
        if not rows:
            self.units.append(Unit(1, pid, sn))
            return
        self.info.is_stack = len(rows) > 1
        for uid_str, _mac, _role in rows:
            uid = int(uid_str)
            m_pid, m_sn = inv_of(uid) or (pid, sn if uid == 1 else "Unknown")
            self.units.append(Unit(uid, m_pid, m_sn))

    def _parse_env(self):
        fans = rsearch(r"Fans\s+Status(.*?)(?=Unit\s+Temperature|\Z)", self.sys_sec, flags=re.S | re.I, default="")
        temp = rsearch(r"Temperature.*?Status(.*?)(?=\Z|\n\n)", self.sys_sec, flags=re.S | re.I, default="")
        self.fan_status = "異常" if ENV_BAD_REGEX_C1300.search(fans) else "正常"
        self.pwr_status = "異常" if ENV_BAD_REGEX_C1300.search(temp) else "正常"

    def _parse_eth_mac(self):
        mac = rsearch(r"MAC\s+Address\s*[:\s]+([0-9a-fA-F:]{17})", self.sys_sec)
        if mac:
            self.info.eth_mac = mac

    def _parse_syslog(self):
        """C1300 show logging 為「新 → 舊」排列：反轉成舊 → 新再合併，才會保留最後一次發生的時間。"""
        entries = C1300_SYSLOG_REGEX.findall(self.log_sec)[::-1]       # [(時間戳記, 訊息)]
        self.syslog = merge_syslog(entries, SYSLOG_KEEP, C1300_EXCLUDE_PREFIXES)


# ==============================================================================
# 12. 設備類型判定
# ==============================================================================

PARSERS: Dict[str, type] = {
    "Switch": IOSSwitchParser,
    "NX-OS":  NexusParser,
    "WLC":    WLCParser,
    "AireOS": AireOSParser,
    "Aruba":  ArubaParser,
    "FortiGate": FortigateParser,
    "C1300":  C1300Parser,
}
PROMPT_BASED = (IOSSwitchParser, NexusParser, WLCParser)   # 可共用 detect 階段建好的提示符索引


def detect_type(content: str) -> Tuple[str, Optional[CommandIndex]]:
    """
    由特定到一般判斷設備類型；IOS 家族會順便回傳已建立的提示符索引供解析器重用。

      1. AireOS : 行首「(xxx) >」提示符出現 ≥ 2 次
                  （V18 只要全文含 RTOS Version 就判定，CDP 鄰居是 AireOS WLC 的交換器會被誤判）
      2. Aruba  : 含「AOS-CX」且有「Command : show」區塊
      2-1. FortiGate : 首行「#config-version=FG…」（.conf）或「Version: FortiGate-…」（get system status）
      3. NX-OS  : 開頭 3000 字元內有 NX-OS 登入 Banner
      4. C1300  : 含 Catalyst 1300 / PID: C1300-，且有 show tech 的「show system」區段
                  （先做便宜的字串比對，Regex 只在必要時才掃描）
      5. 其餘    : 取實際執行的 show version 內容判斷 NX-OS / WLC / Switch
    """
    aireos_hits = sum(1 for _ in zip(range(2), AIREOS_PROMPT_RE.finditer(content, 0, DETECT_WINDOW)))
    if aireos_hits == 1 and AIREOS_PROMPT_RE.search(content, DETECT_WINDOW):
        aireos_hits = 2
    if aireos_hits >= 2:
        return "AireOS", None
    if "AOS-CX" in content and ARUBA_SECTION_RE.search(content):
        return "Aruba", None
    if (re.match(r"\s*#config-version=F[GW]", content[:300])
            or re.search(r"^Version:\s*Forti(?:Gate|WiFi)", content[:DETECT_WINDOW], re.M)):
        return "FortiGate", None
    if NXOS_BANNER in content[:3000]:
        return "NX-OS", None
    if (("Catalyst 1300" in content or re.search(r"PID:\s*C1300-", content))
            and re.search(r"^-{10,}[ \t]+show system[ \t]+-{10,}", content, re.M)):
        return "C1300", None

    idx = build_prompt_index(content)
    ver = ((idx.get("show version") if idx else "")
           or CommandIndex(content, TECH_SECTION_RE).get("show version")
           or content[:5000])
    if "NXOS" in ver or "Nexus" in ver:
        return "NX-OS", idx
    if "Wireless LAN Controller" in ver or "C9800" in ver:
        return "WLC", idx
    if "Cisco IOS" in ver or "Software" in ver:
        return "Switch", idx
    return "Unknown", idx


# ==============================================================================
# 13. 分析引擎（與 GUI 無關，可供 CLI / 測試直接呼叫）
# ==============================================================================

LOG_EXTS = (".log", ".txt", ".conf")       # .conf = FortiGate 備份設定檔


@dataclass
class FileResult:
    """單一 Log 檔的分析結果。"""
    source: str
    dev_type: str = "Unknown"
    rows: Optional[List[Dict]] = None
    aps: Optional[List[Dict]] = None
    summary: str = ""
    error: str = ""
    seconds: float = 0.0


def analyze_file(path: str) -> FileResult:
    """讀檔 → 判定類型 → 解析 → 產生設備列與 AP 列。例外不會中斷整批處理。"""
    fname = os.path.basename(path)
    res = FileResult(fname, rows=[], aps=[])
    t0 = time.perf_counter()
    try:
        content = read_log(path)
        res.dev_type, idx = detect_type(content)
        cls = PARSERS.get(res.dev_type)
        if cls:
            parser = cls(content, source=fname,
                         index=idx if issubclass(cls, PROMPT_BASED) else None).run_all()
            res.rows = parser.generate_rows()
            if hasattr(parser, "get_ap_rows"):
                res.aps = [dict(ap, source=fname) for ap in parser.get_ap_rows()]
            layout = "HA" if parser.info.is_ha else ("Stack" if parser.info.is_stack else "Single")
            res.summary = f"{res.dev_type} | {layout} | {len(res.rows)} 台"
    except Exception as e:                                   # 單檔失敗不影響其他檔案
        res.error = f"{type(e).__name__}: {e}"
    res.seconds = time.perf_counter() - t0
    return res


def mark_member_row(row: Dict):
    """群組非首列：Stack 共用 Hostname / IP / CPU / Memory / Log；HA 只共用 Hostname / 管理 IP。"""
    if row["_kind"] == "stack":
        row.update({"Hostname": f"{MARK} (Stack)", "OOB IP": MARK, "CPU Usage": MARK,
                    "Memory Usage": MARK, "System Log": ""})
    else:
        row.update({"Hostname": f"{MARK} (HA)", "OOB IP": MARK})


def finalize_device_rows(rows: List[Dict]) -> List[Dict]:
    """
    ① HA 去重：同一機箱（MAC）只留一列，優先保留「來自該機箱自己 Log」的列
       （Active log 看得到 Standby 的型號 / 序號，但 CPU / Memory 要從 Standby log 取）
    ② 依群組（Stack / HA Pair）聚合，群組內依 _order 排序，群組間依首列 OOB IP 排序
    ③ 非首列套用「○」標記
    """
    kept: List[Dict] = []
    pos: Dict[str, int] = {}
    for r in rows:
        key = r["_ha_key"]
        if not key:
            kept.append(r)
        elif key not in pos:
            pos[key] = len(kept)
            kept.append(r)
        elif r["_ha_local"] and not kept[pos[key]]["_ha_local"]:
            kept[pos[key]] = r

    groups: Dict[str, List[Dict]] = {}
    for r in kept:
        groups.setdefault(r["_group"], []).append(r)
    for g in groups.values():
        g.sort(key=lambda r: r["_order"])

    result = []
    for g in sorted(groups.values(), key=lambda g: ip_sort_key(g[0]["OOB IP"])):
        for i, r in enumerate(g):
            if i:
                mark_member_row(r)
            result.append(r)
    return result


def finalize_ap_rows(aps: List[Dict]) -> List[Dict]:
    """AP 以（名稱, MAC）去重（HA 兩份 Log 可能重複列出），再依 IP 排序。"""
    uniq: Dict[Tuple[str, str], Dict] = {}
    for ap in aps:
        uniq.setdefault((ap["ap_name"], mac_key(ap["eth_mac"]) or ap["eth_mac"]), ap)
    return sorted(uniq.values(), key=lambda a: ip_sort_key(a.get("ip_address", "")))


def list_log_files(folder: str) -> List[str]:
    return sorted(os.path.join(folder, f) for f in os.listdir(folder)
                  if f.lower().endswith(LOG_EXTS) and not f.startswith("~$"))


def analyze_folder(folder: str, log: Callable[[str], None] = print,
                   progress: Optional[Callable[[int, int], None]] = None
                   ) -> Tuple[List[Dict], List[Dict], List[FileResult]]:
    """分析資料夾內所有 .log / .txt，回傳（設備列, AP 列, 各檔結果）。"""
    files = list_log_files(folder)
    total, results = len(files), []
    for i, path in enumerate(files, 1):
        r = analyze_file(path)
        results.append(r)
        if r.error:
            log(f"  [ERR] {r.source}: {r.error}")
        elif r.dev_type not in PARSERS:
            log(f"[{i}/{total}] {r.source} -> 型態: Unknown [WARN]")
        else:
            if r.dev_type in ("WLC", "AireOS"):
                log(f"  -> 解析到 {len(r.aps)} 台 AP（{'C9800' if r.dev_type == 'WLC' else 'AireOS'}）")
            log(f"[{i}/{total}] {r.source} -> {r.summary}")
        if progress:
            progress(i, total)

    devices = finalize_device_rows([row for r in results for row in r.rows])
    aps = finalize_ap_rows([ap for r in results for ap in r.aps])
    return devices, aps, results


# ==============================================================================
# 14. Excel 輸出
# ==============================================================================

Sheet = Tuple[List[str], List[List[str]]]       # (欄位名稱, 資料列)


def _table(rows: List[Dict], cols: List[str], missing: str = "") -> Sheet:
    """dict 資料列 → 依欄位順序排好的二維表；缺值填 missing。"""
    return cols, [[r.get(c, missing) for c in cols] for r in rows]


def build_sheets(devices: List[Dict], aps: List[Dict]) -> Dict[str, Sheet]:
    """
    Report      ：所有設備 + AP（不含 MAC / Source File）
    AP Info     ：AP 詳細資訊 + Source File
    WLC Info    ：C9800 / AireOS（含 HA Role）+ Source File
    Switch Info ：IOS / NX-OS / C1300 / Aruba + Source File
    Aruba Info / Fortigate Info：偵測到該廠牌設備才建立，含廠牌專屬欄位
    （V21 起不再使用 pandas：直接以 openpyxl 寫入，打包後的 exe 由約 60 MB 降到約 7~10 MB）
    """
    ap_report = [{"Hostname": a["ap_name"], "Model": a["ap_model"], "OOB IP": a["ip_address"],
                  "Serial Number": a["serial"], "Firmware Version": a["firmware"],
                  "CPU Usage": "N/A", "Memory Usage": "N/A", "Fan Status": "正常",
                  "Power Status": "正常", "System Log": ""} for a in aps]
    ap_info = [{"AP Name": a["ap_name"], "AP Model": a["ap_model"], "Ethernet MAC": a["eth_mac"],
                "IP Address": a["ip_address"], "AP Serial Number": a["serial"],
                "Software Version": a["firmware"], "Source File": a["source"]} for a in aps]
    sheets = {
        "Report":      _table(devices + ap_report, REPORT_COLS),
        "AP Info":     _table(ap_info, AP_COLS),
        "WLC Info":    _table([r for r in devices if r["_device_type"] in WLC_TYPES], WLC_COLS),
        "Switch Info": _table([r for r in devices if r["_device_type"] in SWITCH_TYPES], SWITCH_COLS),
    }
    for sheet, dev_type, cols in VENDOR_SHEETS:
        rows = [r for r in devices if r["_device_type"] == dev_type]
        if rows:
            sheets[sheet] = _table(rows, cols, missing="N/A")
    return sheets


def _cell_width(value) -> int:
    """以 Big5 位元組數估算顯示寬度（中文約 2 字元寬），多行取最長行。"""
    if value is None or value == "":
        return 0
    return max(len(line.encode("big5", errors="ignore")) for line in str(value).split("\n"))


def write_workbook(path: str, sheets: Dict[str, Sheet],
                   decorate: Optional[Callable[[str, object], None]] = None):
    """寫入 Excel 並套用格式：藍底白字表頭、微軟正黑體、細框線、自動欄寬（上限 100）。
    decorate(分頁名稱, worksheet)：格式套用完後的額外處理（比對報表的顏色標註）。"""
    header_fill = PatternFill("solid", fgColor="2F75B5")
    header_font = Font(name="微軟正黑體", size=12, bold=True, color="FFFFFF")
    data_font = Font(name="微軟正黑體", size=12)
    header_align = Alignment(horizontal="center")
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = Workbook()
    wb.remove(wb.active)                                 # 移除預設的空白工作表
    for name, (cols, rows) in sheets.items():
        ws = wb.create_sheet(name)
        ws.append(cols)
        for row in rows:
            ws.append(row)
            for cell in ws[ws.max_row]:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.data_type = "s"                 # 以 "=" 開頭的 Log 文字不可被當成公式
        widths = [_cell_width(c) for c in cols]
        for r_idx, row_cells in enumerate(ws.iter_rows(), 1):
            for c_idx, cell in enumerate(row_cells):
                if r_idx == 1:
                    cell.fill, cell.font, cell.alignment = header_fill, header_font, header_align
                else:
                    cell.font = data_font
                    widths[c_idx] = max(widths[c_idx], _cell_width(cell.value))
                cell.border = border
        for c_idx, w in enumerate(widths):
            ws.column_dimensions[get_column_letter(c_idx + 1)].width = min(int(w * 1.2) + 4, 100)
        if decorate:
            decorate(name, ws)
    wb.save(path)


def write_excel(path: str, devices: List[Dict], aps: List[Dict]):
    """維護報表：Report / AP Info / WLC Info / Switch Info（+ 廠牌分頁）。"""
    write_workbook(path, build_sheets(devices, aps))


def default_report_name() -> str:
    return f"Network_Analysis_{datetime.now():%m%d_%H%M}.xlsx"


# ==============================================================================
# 15. 維護設備比對（V23）：上季 / 本季兩份維護報表 → 差異報表
# ==============================================================================
#
# 輸入是本工具產出的維護報表（.xlsx），不需要原始 Log。
#   ① 讀取 Switch Info / WLC Info / Fortigate Info / AP Info
#   ② Stack / HA 非首列的「○」補成首列的值（○ (Stack) / ○ (HA) → 首列 Hostname）
#   ③ 以 Serial Number 配對（序號在換機、改名、改 IP 後都不會變）；
#      沒有序號時改用 Hostname + Model，並列入「需人工確認」
#   ④ 分類：新增 / 移除 / 換機 / 欄位異動 / 需人工確認
# 輸出第一頁「設備總表」：本季全部設備 + AP，最下方列出本季已移除的設備，依顏色標註：
#   淡綠整列 = 新增；淡紅整列（Hostname、SN 深紅）= 移除；深紅 SN = 換機；
#   淺黃格子 = 欄位異動（註解為上季值）；橘黃 = 需人工確認
# CPU / Memory 是收 Log 當下的數值，每季本來就不同，不列入異動。

COMPARE_DEVICE_SHEETS = (("Switch Info", "Switch"), ("WLC Info", "WLC"), ("Fortigate Info", "FortiGate"))
COMPARE_FIELDS = ("Model", "OOB IP", "Firmware Version", "HA Role", "Fan Status", "Power Status")
COMPARE_SHEETS = ("設備總表", "摘要", "新增設備", "移除設備", "換機與成員變更", "欄位異動",
                  "新增 AP", "移除 AP", "需人工確認")
DETAIL_SHEETS = COMPARE_SHEETS[2:]
MASTER_COLS = ["狀態", "類型", "Hostname", "Model", "OOB IP", "Ethernet MAC", "Serial Number",
               "Firmware Version", "HA Role", "CPU Usage", "Memory Usage", "Fan Status", "Power Status",
               "System Log", "Source File", "變更說明"]
AP_TO_MASTER = {"Hostname": "AP Name", "Model": "AP Model", "OOB IP": "IP Address", "Ethernet MAC": "Ethernet MAC",
                "Serial Number": "AP Serial Number", "Firmware Version": "Software Version",
                "Source File": "Source File"}
DEVICE_LIST_COLS = ["類型", "Hostname", "位置", "Model", "OOB IP", "Serial Number",
                    "Firmware Version", "Source File"]
REPLACE_COLS = ["類型", "Hostname", "位置", "判定", "上季 Model", "上季 Serial", "本季 Model",
                "本季 Serial", "本季 Source File"]
CHANGE_COLS = ["類型", "Hostname", "位置", "Serial Number", "欄位", "上季", "本季"]
REVIEW_COLS = ["類型", "Hostname", "Model", "Serial Number", "季別", "原因"]
NO_SERIAL = {"", "N/A", "NA", "-", "UNKNOWN", MARK}
SYSLOG_ALERT_RE = re.compile(r"FAIL\w*|FAULT\w*|CRITICAL|SHUTDOWN|OVER[ _-]?TEMP\w*", re.I)
SYSLOG_STAMP_RE = re.compile(
    r"^[\s*.]*(?:\w{3}\s+\d{1,2}\s+(?:\d{4}\s+)?|\d{4}[-/]\d{1,2}[-/]\d{1,2}[ T]?)?"
    r"(?:\d{1,2}:\d{2}:\d{2}(?:\.\d+)?)?\s*(?:[A-Z]{2,5}\s*)?:?\s*")
ALERT_SHOW = 5                                   # 欄位異動中每台最多列出幾筆新增告警

# 顏色沿用 KINSUS 維護報告 VBA 對帳巨集：淺黃 = 異動、淡紅 = 序號異動 / 遺失
FILL_ADD = "C6EFCE"                              # 新增：整列淡綠
FILL_REMOVED = "FFC7CE"                          # 移除：整列淡紅
FILL_RED = "FF9999"                              # 換機只標 SN；移除的 Hostname / SN
FILL_CHANGED = "FFFF99"                          # 欄位異動的格子（RGB 255,255,153）
FILL_REVIEW = "FFE699"                           # 需人工確認
STATUS_ORDER = ("新增", "移除", "換機", "異動", "需確認", "")
STATUS_PRIORITY = {"": 0, "異動": 1, "需確認": 2, "換機": 3, "新增": 4, "移除": 5}


class CompareError(ValueError):
    """輸入檔不是本工具產出的維護報表。"""


@dataclass
class ReportDevice:
    kind: str                        # Switch / WLC / FortiGate
    host: str                        # 已補齊的 Hostname（群組成員 = 首列 Hostname）
    pos: str                         # 群組內位置：成員 N（Stack）/ 機箱 N（HA）；單機為空
    values: Dict[str, str]           # 「○」已補成首列的值
    serial: str                      # 正規化序號；無效時為空字串
    raw: Dict[str, str]              # 報表原始內容（設備總表照原樣顯示「○」）
    group: int = 0
    idx: int = 0                     # 群組內第幾台（0 = 首列）

    def key(self) -> Tuple:
        return ("SN", self.serial) if self.serial else ("HM", self.kind, self.host.upper(),
                                                        self.values.get("Model", "").upper())

    def label(self) -> str:
        return f"{self.host} {self.pos}".strip()


@dataclass
class ReportData:
    name: str
    devices: List[ReportDevice]
    aps: List[Dict[str, str]]


@dataclass
class Mark:
    """設備總表中一列的標註：狀態、整列底色、單格底色、註解、變更說明。"""
    status: str = ""
    row_fill: str = ""
    cells: Dict[str, str] = field(default_factory=dict)
    comments: Dict[str, str] = field(default_factory=dict)
    notes: List[str] = field(default_factory=list)

    def set_status(self, status: str):
        if STATUS_PRIORITY[status] > STATUS_PRIORITY[self.status]:
            self.status = status

    def change(self, col: str, old: str, new: str, note: str):
        self.set_status("異動")
        self.cells.setdefault(col, FILL_CHANGED)
        self.comments[col] = f"上季：{old or '（空白）'}"
        self.notes.append(note)


@dataclass
class MasterRow:
    values: Dict[str, str]
    mark: Mark
    label: str                       # GUI / 網頁版顯示用名稱（Stack 成員顯示真實 Hostname + 位置）
    is_ap: bool = False
    from_old: bool = False           # 上季資料（本季已移除，或上季無序號無法判斷）


@dataclass
class CompareResult:
    old: ReportData
    new: ReportData
    sheets: Dict[str, List[Dict[str, str]]]      # 明細分頁名稱 → 資料列
    master: List[MasterRow] = field(default_factory=list)

    def changed_units(self) -> int:
        return sum(1 for r in self.master if FILL_CHANGED in r.mark.cells.values())

    def counts(self) -> Dict[str, int]:
        """各明細分頁的數量；欄位異動以「台數」計（同一台多個欄位只算一台）。"""
        c = {name: len(rows) for name, rows in self.sheets.items()}
        c["欄位異動"] = self.changed_units()
        return c

    def device_stats(self) -> Dict[str, int]:
        devs = [r for r in self.master if not r.is_ap]
        live = [r for r in devs if not r.from_old]
        return {
            "相符": sum(1 for r in live if r.mark.status == ""),
            "換機": sum(1 for r in live if r.mark.status == "換機"),
            "移除": sum(1 for r in devs if r.mark.status == "移除"),
            "新增": sum(1 for r in live if r.mark.status == "新增"),
            "異動": sum(1 for r in devs if FILL_CHANGED in r.mark.cells.values()),
        }

    def ap_stats(self) -> Dict[str, int]:
        aps = [r for r in self.master if r.is_ap]
        return {
            "新增": sum(1 for r in aps if r.mark.status == "新增"),
            "移除": sum(1 for r in aps if r.mark.status == "移除"),
            "換機": sum(1 for r in aps if r.mark.status == "換機"),
            "改名": sum(1 for r in aps if "Hostname" in r.mark.comments),
        }


def _norm_serial(value) -> str:
    s = str(value or "").strip().upper()
    return "" if s in NO_SERIAL else s


def _cell_text(value) -> str:
    return "" if value is None else str(value).strip()


def load_report(path: str) -> ReportData:
    """讀取維護報表：補齊「○」、標上群組位置。不是本工具產出的報表時丟出 CompareError。"""
    from openpyxl import load_workbook
    name = os.path.basename(path)
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        raise CompareError(f"{name} 無法開啟（{e}）") from e
    try:
        if "Switch Info" not in wb.sheetnames or "AP Info" not in wb.sheetnames:
            raise CompareError(f"{name} 不是本工具產出的維護報表（找不到 Switch Info / AP Info 分頁）")

        def rows_of(sheet: str) -> List[Dict[str, str]]:
            if sheet not in wb.sheetnames:
                return []
            it = wb[sheet].iter_rows(values_only=True)
            header = [_cell_text(h) for h in next(it, ())]
            return [dict(zip(header, map(_cell_text, row))) for row in it if any(c is not None for c in row)]

        devices: List[ReportDevice] = []
        group_no = 0
        for sheet, kind in COMPARE_DEVICE_SHEETS:
            rows = rows_of(sheet)
            if rows and ("Hostname" not in rows[0] or "Serial Number" not in rows[0]):
                raise CompareError(f"{name} 的 {sheet} 分頁缺少 Hostname / Serial Number 欄位")
            groups: List[List[Dict[str, str]]] = []
            for r in rows:
                if r.get("Hostname", "").startswith(MARK) and groups:
                    groups[-1].append(r)            # 「○ (Stack)」「○ (HA)」屬於上一個群組
                else:
                    groups.append([r])
            for members in groups:
                group_no += 1
                head = members[0]
                is_ha = len(members) > 1 and members[1].get("Hostname", "").endswith("(HA)")
                for i, r in enumerate(members):
                    values = {k: (head.get(k, "") if v == MARK else v) for k, v in r.items()}
                    values["Hostname"] = head.get("Hostname", "")
                    pos = (f"機箱 {i + 1}" if is_ha else f"成員 {i + 1}") if len(members) > 1 else ""
                    devices.append(ReportDevice(kind, values["Hostname"], pos, values,
                                                _norm_serial(values.get("Serial Number")), r, group_no, i))
        aps = [r for r in rows_of("AP Info") if r.get("AP Name") or r.get("AP Serial Number")]
        return ReportData(name, devices, aps)
    finally:
        wb.close()


def _alert_lines(log_text: str) -> Dict[str, str]:
    """System Log 中含 FAIL / FAULT / CRITICAL / SHUTDOWN / OVERTEMP 的訊息（去掉時間戳記後當鍵）。"""
    out: Dict[str, str] = {}
    for line in log_text.splitlines():
        line = line.strip()
        if not line or not SYSLOG_ALERT_RE.search(line):
            continue
        key = line[line.index("%"):] if "%" in line else SYSLOG_STAMP_RE.sub("", line, count=1)
        out.setdefault(key, line)
    return out


def _device_row(d: ReportDevice) -> Dict[str, str]:
    return {"類型": d.kind, "Hostname": d.host, "位置": d.pos, **{c: d.values.get(c, "") for c in
            ("Model", "OOB IP", "Serial Number", "Firmware Version", "Source File")}}


def _index(items, key_fn) -> Tuple[Dict[Tuple, object], List[object]]:
    """以 key_fn 建索引；重複的鍵只保留第一筆，其餘回傳給呼叫端列入「需人工確認」。"""
    idx: Dict[Tuple, object] = {}
    dups: List[object] = []
    for it in items:
        k = key_fn(it)
        if k in idx:
            dups.append(it)
        else:
            idx[k] = it
    return idx, dups


def _dup_reason(key: Tuple) -> str:
    return ("序號重複" if key[0] == "SN" else "無序號且名稱 / 型號重複") + "，只比對第一筆"


def _diff_devices(old: ReportData, new: ReportData, sheets: Dict[str, List[Dict[str, str]]]
                  ) -> Tuple[Dict[int, Mark], List[Tuple[ReportDevice, Mark]]]:
    """回傳（本季設備的標註 {id(device): Mark}, 本季已移除的上季設備與標註）。"""
    review = sheets["需人工確認"]
    marks: Dict[int, Mark] = {}
    removed: List[Tuple[ReportDevice, Mark]] = []
    mk = lambda d: marks.setdefault(id(d), Mark())
    describe = lambda d: {"類型": d.kind, "Hostname": d.label(), "Model": d.values.get("Model", ""),
                          "Serial Number": d.values.get("Serial Number", "")}

    def gone(d: ReportDevice, note: str, status: str = "移除"):
        m = Mark(status=status, row_fill=FILL_REMOVED if status == "移除" else "",
                 cells={"Hostname": FILL_RED, "Serial Number": FILL_RED} if status == "移除"
                 else {"狀態": FILL_REVIEW, "Serial Number": FILL_REVIEW}, notes=[note])
        removed.append((d, m))

    def needs_review(d: ReportDevice, note: str):
        m = mk(d)
        m.set_status("需確認")
        m.cells["Serial Number"] = FILL_REVIEW
        m.notes.append(note)

    o_idx, o_dups = _index(old.devices, ReportDevice.key)
    n_idx, n_dups = _index(new.devices, ReportDevice.key)
    for d in o_dups:
        review.append({**describe(d), "季別": "上季", "原因": _dup_reason(d.key())})
    for d in n_dups:
        review.append({**describe(d), "季別": "本季", "原因": _dup_reason(d.key())})
        needs_review(d, _dup_reason(d.key()))

    # ── 配對成功：逐欄比較（群組成員顯示為「○」的共用欄位只在首列比較）───────────
    changes: List[Dict[str, str]] = []
    group_sizes = Counter((d.kind, d.host) for d in new.devices)
    for k, nd in n_idx.items():
        od = o_idx.get(k)
        if od is None:
            continue
        m = mk(nd)
        if k[0] == "HM":
            review.append({**describe(nd), "季別": "兩季", "原因": "無序號，以 Hostname + Model 配對，請確認是否為同一台"})
            needs_review(nd, "無序號，以 Hostname + Model 配對")
        base = {"類型": nd.kind, "Hostname": nd.host, "位置": nd.pos,
                "Serial Number": nd.values.get("Serial Number", "")}
        shared = lambda col: nd.raw.get(col, "").startswith(MARK)
        if od.host != nd.host and not shared("Hostname"):
            changes.append({**base, "欄位": "Hostname（改名）", "上季": od.host, "本季": nd.host})
            m.change("Hostname", od.host, nd.host, f"改名：{od.host} → {nd.host}")
        for col in COMPARE_FIELDS:
            a, b = od.values.get(col, ""), nd.values.get(col, "")
            if col in nd.values and col in od.values and a != b and not shared(col):
                changes.append({**base, "欄位": col, "上季": a, "本季": b})
                m.change(col, a, b, f"{col}：{a or '—'} → {b or '—'}")
        old_alerts = _alert_lines(od.values.get("System Log", ""))
        new_alerts = [v for key, v in _alert_lines(nd.values.get("System Log", "")).items()
                      if key not in old_alerts]
        if new_alerts:
            shown = new_alerts[:ALERT_SHOW]
            more = f"\n…另 {len(new_alerts) - ALERT_SHOW} 筆" if len(new_alerts) > ALERT_SHOW else ""
            changes.append({**base, "欄位": "System Log（新增告警）", "上季": "",
                            "本季": "\n".join(shown) + more})
            m.change("System Log", "", "", f"System Log 新增告警 {len(new_alerts)} 筆")
            m.comments["System Log"] = "新增告警：\n" + "\n".join(shown) + more

    # 同一群組多台成員有相同異動（例如 Stack 一起升級韌體）→ 欄位異動分頁合併成一列
    merged: Dict[Tuple, List[Dict[str, str]]] = {}
    for c in changes:
        merged.setdefault((c["類型"], c["Hostname"], c["欄位"], c["上季"], c["本季"]), []).append(c)
    for (kind, host, *_), rows in merged.items():
        if len(rows) > 1:
            pos = (f"全部 {len(rows)} 台" if len(rows) == group_sizes[(kind, host)]
                   else "、".join(r["位置"] for r in rows))
            rows = [{**rows[0], "位置": pos, "Serial Number": ""}]
        sheets["欄位異動"].extend(rows)

    # ── 未配對：同 Hostname → 換機 / 成員增減；其餘 → 新增 / 移除 ──────────────
    o_left = [d for k, d in o_idx.items() if k not in n_idx]
    n_left = [d for k, d in n_idx.items() if k not in o_idx]
    for d in o_left:
        if not d.serial:
            review.append({**describe(d), "季別": "上季",
                           "原因": "無序號且找不到本季的同名同型號設備，無法判斷是否為移除"})
            gone(d, "上季資料：無序號，無法判斷是否移除", status="需確認")
    for d in n_left:
        if not d.serial:
            review.append({**describe(d), "季別": "本季",
                           "原因": "無序號且找不到上季的同名同型號設備，無法判斷是否為新增"})
            needs_review(d, "無序號，無法判斷是否為新增")
    o_left = [d for d in o_left if d.serial]
    n_left = [d for d in n_left if d.serial]
    old_hosts = {(d.kind, d.host.upper()) for d in old.devices}
    new_hosts = {(d.kind, d.host.upper()) for d in new.devices}

    def added(d: ReportDevice, note: str):
        sheets["新增設備"].append(_device_row(d))
        m = mk(d)
        m.set_status("新增")
        m.row_fill = FILL_ADD
        m.notes.append(note)

    def removed_dev(d: ReportDevice, note: str):
        sheets["移除設備"].append(_device_row(d))
        gone(d, note)

    for kind_host in sorted({(d.kind, d.host.upper()) for d in n_left} | {(d.kind, d.host.upper()) for d in o_left}):
        out = [d for d in o_left if (d.kind, d.host.upper()) == kind_host]
        came = [d for d in n_left if (d.kind, d.host.upper()) == kind_host]
        if kind_host not in old_hosts:              # 上季沒有這個 Hostname → 全新設備
            for d in came:
                added(d, f"新上線設備（{d.pos}）" if d.pos else "新上線設備")
            continue
        if kind_host not in new_hosts:              # 本季沒有這個 Hostname → 整組下線
            for d in out:
                removed_dev(d, f"本季找不到（上季資料，{d.pos}）" if d.pos else "本季找不到（上季資料）")
            continue
        group = len([d for d in new.devices if (d.kind, d.host.upper()) == kind_host]) > 1 or \
            len([d for d in old.devices if (d.kind, d.host.upper()) == kind_host]) > 1
        out_by_pos = {d.pos: d for d in out}
        pairs: List[Tuple[Optional[ReportDevice], Optional[ReportDevice]]] = []
        for d in came:                              # 同位置優先配對，其次依序配對
            pairs.append((out_by_pos.pop(d.pos, None), d))
        rest = list(out_by_pos.values())
        pairs = [(o or (rest.pop(0) if rest else None), n) for o, n in pairs] + [(o, None) for o in rest]
        for o, n in pairs:
            unit = "機箱" if (n or o).pos.startswith("機箱") else "成員"
            if o and n:
                verdict = f"{unit}換新" if group else "換機"
                o_sn, n_sn = o.values.get("Serial Number", ""), n.values.get("Serial Number", "")
                sheets["換機與成員變更"].append({
                    "類型": n.kind, "Hostname": n.host, "位置": n.pos, "判定": verdict,
                    "上季 Model": o.values.get("Model", ""), "上季 Serial": o_sn,
                    "本季 Model": n.values.get("Model", ""), "本季 Serial": n_sn,
                    "本季 Source File": n.values.get("Source File", "")})
                m = mk(n)
                m.set_status("換機")
                m.cells["Serial Number"] = FILL_RED
                m.comments["Serial Number"] = f"上季：{o_sn}"
                m.notes.append(f"{n.pos} {verdict[-2:]}：{o_sn} → {n_sn}" if group else f"換機：{o_sn} → {n_sn}")
                if o.values.get("Model", "") != n.values.get("Model", ""):
                    m.change("Model", o.values.get("Model", ""), n.values.get("Model", ""),
                             f"Model：{o.values.get('Model', '')} → {n.values.get('Model', '')}")
                    m.status = "換機"
            elif n:
                added(n, f"{n.pos} 新增" if n.pos else "新上線設備")
            else:
                removed_dev(o, f"{o.pos} 移除（上季資料）" if o.pos else "本季找不到（上季資料）")
    return marks, removed


def _ap_key(ap: Dict[str, str]) -> Tuple:
    sn = _norm_serial(ap.get("AP Serial Number"))
    if sn:
        return ("SN", sn)
    mac = mac_key(ap.get("Ethernet MAC", ""))
    return ("MAC", mac) if mac else ("NAME", ap.get("AP Name", "").upper())


def _diff_aps(old: ReportData, new: ReportData, sheets: Dict[str, List[Dict[str, str]]]
              ) -> Tuple[Dict[int, Mark], List[Tuple[Dict[str, str], Mark]]]:
    review = sheets["需人工確認"]
    marks: Dict[int, Mark] = {}
    mk = lambda a: marks.setdefault(id(a), Mark())
    describe = lambda a: {"類型": "AP", "Hostname": a.get("AP Name", ""), "Model": a.get("AP Model", ""),
                          "Serial Number": a.get("AP Serial Number", "")}
    o_idx, o_dups = _index(old.aps, _ap_key)
    n_idx, n_dups = _index(new.aps, _ap_key)
    for a in o_dups:
        review.append({**describe(a), "季別": "上季", "原因": _dup_reason(_ap_key(a))})
    for a in n_dups:
        review.append({**describe(a), "季別": "本季", "原因": _dup_reason(_ap_key(a))})
        m = mk(a)
        m.set_status("需確認")
        m.cells["Serial Number"] = FILL_REVIEW
        m.notes.append(_dup_reason(_ap_key(a)))
    pick = lambda a: {c: a.get(c, "") for c in AP_COLS}
    for k, a in n_idx.items():
        o = o_idx.get(k)
        if o and o.get("AP Name") != a.get("AP Name"):
            sheets["欄位異動"].append({"類型": "AP", "Hostname": a.get("AP Name", ""), "位置": "",
                                    "Serial Number": a.get("AP Serial Number", ""),
                                    "欄位": "AP Name（改名）", "上季": o.get("AP Name", ""),
                                    "本季": a.get("AP Name", "")})
            mk(a).change("Hostname", o.get("AP Name", ""), a.get("AP Name", ""),
                         f"改名：{o.get('AP Name', '')} → {a.get('AP Name', '')}")
    gone = {k: a for k, a in o_idx.items() if k not in n_idx}
    came = {k: a for k, a in n_idx.items() if k not in o_idx}
    gone_by_name = {a.get("AP Name", "").upper(): k for k, a in gone.items() if a.get("AP Name")}
    for k, a in list(came.items()):                 # 同名 AP 換了序號 → AP 換機
        ok = gone_by_name.pop(a.get("AP Name", "").upper(), None)
        if ok is None:
            continue
        o = gone.pop(ok)
        del came[k]
        o_sn, n_sn = o.get("AP Serial Number", ""), a.get("AP Serial Number", "")
        sheets["換機與成員變更"].append({
            "類型": "AP", "Hostname": a.get("AP Name", ""), "位置": "", "判定": "AP 換機",
            "上季 Model": o.get("AP Model", ""), "上季 Serial": o_sn,
            "本季 Model": a.get("AP Model", ""), "本季 Serial": n_sn,
            "本季 Source File": a.get("Source File", "")})
        m = mk(a)
        m.set_status("換機")
        m.cells["Serial Number"] = FILL_RED
        m.comments["Serial Number"] = f"上季：{o_sn}"
        m.notes.append(f"AP 換機：{o_sn} → {n_sn}")
    for a in came.values():
        sheets["新增 AP"].append(pick(a))
        m = mk(a)
        m.set_status("新增")
        m.row_fill = FILL_ADD
        m.notes.append("新上線 AP")
    removed = []
    for a in gone.values():
        sheets["移除 AP"].append(pick(a))
        removed.append((a, Mark(status="移除", row_fill=FILL_REMOVED,
                                cells={"Hostname": FILL_RED, "Serial Number": FILL_RED},
                                notes=["本季找不到（上季資料）"])))
    return marks, removed


def _master_device(d: ReportDevice, m: Mark, from_old: bool = False) -> MasterRow:
    """本季設備照報表原樣（含「○」）；上季資料列已脫離群組，顯示補齊後的 Hostname / IP。"""
    src = {**d.values, "Hostname": d.host} if from_old else d.raw
    values = {c: src.get(c, "") for c in MASTER_COLS}
    values.update({"狀態": m.status, "類型": d.kind, "變更說明": "；".join(m.notes)})
    return MasterRow(values, m, d.label() + ("（上季）" if from_old and m.status != "移除" else ""),
                     from_old=from_old)


def _master_ap(a: Dict[str, str], m: Mark, from_old: bool = False) -> MasterRow:
    values = {c: a.get(AP_TO_MASTER[c], "") if c in AP_TO_MASTER else "" for c in MASTER_COLS}
    values.update({"狀態": m.status, "類型": "AP", "變更說明": "；".join(m.notes)})
    return MasterRow(values, m, a.get("AP Name", ""), is_ap=True, from_old=from_old)


def compare_reports(old_path: str, new_path: str) -> CompareResult:
    """上季 / 本季維護報表比對；不是本工具產出的報表或無法讀取時丟出 CompareError。"""
    old, new = load_report(old_path), load_report(new_path)
    sheets: Dict[str, List[Dict[str, str]]] = {name: [] for name in DETAIL_SHEETS}
    dev_marks, dev_removed = _diff_devices(old, new, sheets)
    ap_marks, ap_removed = _diff_aps(old, new, sheets)

    # 設備總表：本季設備依群組首列 OOB IP 排序（同維護報表），接著 AP，最下方為本季已移除
    heads = {d.group: d for d in new.devices if d.idx == 0}
    devices = sorted(new.devices, key=lambda d: (ip_sort_key(heads[d.group].raw.get("OOB IP", "")),
                                                 d.group, d.idx))
    master = [_master_device(d, dev_marks.get(id(d), Mark())) for d in devices]
    master += [_master_ap(a, ap_marks.get(id(a), Mark())) for a in new.aps]
    master += [_master_device(d, m, from_old=True) for d, m in dev_removed]
    master += [_master_ap(a, m, from_old=True) for a, m in ap_removed]
    return CompareResult(old, new, sheets, master)


def _summary_rows(result: CompareResult) -> List[List[str]]:
    d, a = result.device_stats(), result.ap_stats()
    items = len(result.sheets["欄位異動"])
    side = lambda r: f"{r.name}（{len(r.devices)} 台設備、{len(r.aps)} 支 AP）"
    return [
        ["比對來源", "", ""],
        ["上季報表", side(result.old), ""],
        ["本季報表", side(result.new), ""],
        ["【設備盤點對帳結果】", "", ""],
        ["1. 序號完全相符（無變更）", f"{d['相符']} 台", "無底色"],
        ["2. 換機 / 成員換新（序號異動）", f"{d['換機']} 台", "只標 Serial Number（深紅）"],
        ["3. 移除設備（本季找不到）", f"{d['移除']} 台", "整列淡紅，Hostname 與 Serial Number 深紅"],
        ["4. 新增設備 / Stack 新成員", f"{d['新增']} 台", "整列淡綠"],
        ["5. 欄位異動（以台數計）", f"{d['異動']} 台", "只標有變的格子（淺黃），註解為上季值"],
        ["6. 需人工確認（無序號 / 序號重複）", f"{len(result.sheets['需人工確認'])} 筆", "狀態與 Serial Number 橘黃"],
        ["【AP】", "", ""],
        ["新增 AP", f"{a['新增']} 支", "整列淡綠"],
        ["移除 AP", f"{a['移除']} 支", "整列淡紅"],
        ["AP 換機", f"{a['換機']} 支", "只標 Serial Number（深紅）"],
        ["AP 改名", f"{a['改名']} 支", "只標 Hostname（淺黃）"],
        ["說明", "", ""],
        ["配對方式", "以 Serial Number 配對；CPU / Memory 為收 Log 當下的數值，不列入異動", ""],
        ["欄位異動明細", f"{items} 項（見「欄位異動」分頁）", ""],
        ["比對時間", f"{datetime.now():%Y-%m-%d %H:%M}", ""],
        ["版本", f"Network Analysis Tool {APP_VERSION}", ""],
    ]


SUMMARY_FILLS = {"2.": FILL_RED, "3.": FILL_REMOVED, "4.": FILL_ADD, "5.": FILL_CHANGED, "6.": FILL_REVIEW,
                 "新增 AP": FILL_ADD, "移除 AP": FILL_REMOVED, "AP 換機": FILL_RED, "AP 改名": FILL_CHANGED}
SUMMARY_SECTIONS = ("比對來源", "【設備盤點對帳結果】", "【AP】", "說明")
DETAIL_FILLS = {"新增設備": (FILL_ADD, None), "新增 AP": (FILL_ADD, None),
                "移除設備": (FILL_REMOVED, None), "移除 AP": (FILL_REMOVED, None),
                "換機與成員變更": (None, ("上季 Serial", "本季 Serial")),
                "欄位異動": (None, ("本季",)), "需人工確認": (None, ("原因",))}
WRAP_COLS = ("變更說明", "本季", "原因", "內容")


def build_compare_sheets(result: CompareResult) -> Dict[str, Sheet]:
    """比對報表分頁；沒有差異的分頁仍建立，內容寫「無」，代表已檢查過。"""
    col_map = {"新增設備": DEVICE_LIST_COLS, "移除設備": DEVICE_LIST_COLS, "換機與成員變更": REPLACE_COLS,
               "欄位異動": CHANGE_COLS, "新增 AP": AP_COLS, "移除 AP": AP_COLS, "需人工確認": REVIEW_COLS}
    sheets: Dict[str, Sheet] = {
        "設備總表": _table([r.values for r in result.master], MASTER_COLS),
        "摘要": (["項目", "數量", "設備總表標註方式"], _summary_rows(result)),
    }
    for name, cols in col_map.items():
        rows = result.sheets[name]
        sheets[name] = _table(rows, cols) if rows else (cols, [["無"] + [""] * (len(cols) - 1)])
    return sheets


def _decorate_compare(result: CompareResult):
    """write_workbook 的後製：設備總表上色 / 註解 / 圖例 / 篩選，摘要與明細分頁上色。"""
    from openpyxl.comments import Comment
    fill = lambda hex_: PatternFill("solid", fgColor=hex_)
    bold = Font(name="微軟正黑體", size=12, bold=True)
    wrap = Alignment(wrap_text=True, vertical="top")

    def wrap_cells(ws, header_row: int):
        cols = [c.column for c in ws[header_row] if c.value in WRAP_COLS]
        for row in ws.iter_rows(min_row=header_row + 1):
            for c in row:
                if c.column in cols and isinstance(c.value, str) and "\n" in c.value:
                    c.alignment = wrap

    def decorate(name: str, ws):
        if name == "設備總表":
            ws.insert_rows(1, 2)                         # 第 1 列標題、第 2 列圖例，表頭在第 3 列
            ws["A1"] = f"設備總表（本季：{result.new.name} ／ 上季：{result.old.name}）"
            ws["A1"].font = Font(name="微軟正黑體", size=14, bold=True)
            legend = [("圖例", None), ("新增", FILL_ADD), ("移除", FILL_REMOVED), ("換機（SN）", FILL_RED),
                      ("異動欄位", FILL_CHANGED), ("需人工確認", FILL_REVIEW), ("紅色三角 = 註解（上季值）", None)]
            for i, (text, color) in enumerate(legend, 1):
                c = ws.cell(row=2, column=i, value=text)
                c.font = Font(name="微軟正黑體", size=11, bold=color is None and i == 1)
                c.alignment = Alignment(horizontal="center" if color else "left")
                if color:
                    c.fill = fill(color)
            col_no = {name: i + 1 for i, name in enumerate(MASTER_COLS)}
            for r_idx, row in enumerate(result.master, 4):
                m = row.mark
                if m.row_fill:
                    for c_idx in range(1, len(MASTER_COLS) + 1):
                        ws.cell(row=r_idx, column=c_idx).fill = fill(m.row_fill)
                if m.status:
                    ws.cell(row=r_idx, column=1).font = Font(name="微軟正黑體", size=12, bold=True)
                if m.status == "需確認":
                    ws.cell(row=r_idx, column=1).fill = fill(FILL_REVIEW)
                for col, color in m.cells.items():
                    ws.cell(row=r_idx, column=col_no[col]).fill = fill(color)
                for col, text in m.comments.items():
                    cm = Comment(text, "Network Analysis Tool")
                    cm.width, cm.height = 260, 60 + 15 * text.count("\n")
                    ws.cell(row=r_idx, column=col_no[col]).comment = cm
            last = get_column_letter(len(MASTER_COLS))
            ws.auto_filter.ref = f"A3:{last}{max(3, len(result.master) + 3)}"
            ws.freeze_panes = "D4"
            ws.column_dimensions[get_column_letter(col_no["System Log"])].width = 60
            ws.column_dimensions[get_column_letter(col_no["變更說明"])].width = 50
            wrap_cells(ws, 3)
        elif name == "摘要":
            for row in ws.iter_rows(min_row=2):
                label = str(row[0].value or "")
                if label in SUMMARY_SECTIONS:
                    for c in row:
                        c.font, c.fill = bold, fill("F1EFE8")
                for key, color in SUMMARY_FILLS.items():
                    if label.startswith(key):
                        row[2].fill = fill(color)
        elif name in DETAIL_FILLS and result.sheets.get(name):
            row_color, cols = DETAIL_FILLS[name]
            header = [c.value for c in ws[1]]
            targets = [header.index(c) + 1 for c in (cols or ()) if c in header]
            cell_color = {"換機與成員變更": FILL_RED, "欄位異動": FILL_CHANGED, "需人工確認": FILL_REVIEW}.get(name)
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    if row_color:
                        c.fill = fill(row_color)
                    elif c.column in targets and c.value:
                        c.fill = fill(cell_color)
            ws.freeze_panes = "A2"
            wrap_cells(ws, 1)
    return decorate


def write_compare_excel(path: str, result: CompareResult):
    write_workbook(path, build_compare_sheets(result), decorate=_decorate_compare(result))


def compare_items(result: CompareResult) -> List[Tuple[str, str, str]]:
    """每台有變化的設備 / AP 一行（狀態, 名稱, 變更說明），依新增 → 移除 → 換機 → 異動 → 需確認排序；
    GUI / CLI 記錄與網頁版預覽共用。"""
    rows = [r for r in result.master if r.mark.status]
    rows.sort(key=lambda r: STATUS_ORDER.index(r.mark.status))
    return [(("AP " if r.is_ap else "") + r.mark.status, r.label, r.values["變更說明"]) for r in rows]


def compare_log_lines(result: CompareResult) -> List[str]:
    """GUI / CLI 顯示用的比對摘要。"""
    lines = [f"[比對] 上季：{result.old.name}（{len(result.old.devices)} 台設備、{len(result.old.aps)} 支 AP）",
             f"[比對] 本季：{result.new.name}（{len(result.new.devices)} 台設備、{len(result.new.aps)} 支 AP）"]
    lines += [f"[{cat}] {name}｜{desc}" if desc else f"[{cat}] {name}" for cat, name, desc in compare_items(result)]
    c = result.counts()
    unit = lambda name: "支" if "AP" in name else "筆" if name == "需人工確認" else "台"
    lines.append("｜".join(f"{name} {c[name]} {unit(name)}" for name in DETAIL_SHEETS))
    return lines


def default_compare_name(old_path: str, new_path: str) -> str:
    stem = lambda p: os.path.splitext(os.path.basename(p))[0]
    return f"維護設備比對_{stem(old_path)}_vs_{stem(new_path)}.xlsx"


# ==============================================================================
# 16. GUI（Tkinter）
# ==============================================================================

class CiscoParserApp:
    """
    操作流程：
      分析 Log    ：選擇資料夾 → 產生分析報表（背景執行緒解析）→ 選擇儲存路徑 → 輸出 Excel
      維護設備比對：選擇上季 / 本季報表 → 產生比對報表（背景執行緒比對）→ 選擇儲存路徑 → 輸出 Excel

    執行緒模型（V19 修正）：
      背景執行緒只做「解析」與「寫檔」，所有 Tk 呼叫（文字框、進度條、對話框）
      一律經由 _ui_queue 交給主執行緒的輪詢器處理；V18 在背景執行緒直接呼叫
      filedialog / messagebox，Tkinter 非 thread-safe，有機率卡死或閃退。
    """

    THEMES: Dict[str, Dict[str, str]] = {
        "light": {
            "root_bg": "#F0F2F5", "toolbar_bg": "#FFFFFF", "sep_color": "#DDE0E6",
            "body_bg": "#F0F2F5", "entry_bg": "#FFFFFF", "entry_fg": "#1A1A1A",
            "entry_border": "#C8CDD5", "text_bg": "#FFFFFF", "text_fg": "#1A1A1A",
            "text_border": "#C8CDD5", "label_fg": "#333333", "btn_folder_bg": "#EAEDF1",
            "btn_folder_fg": "#333333", "btn_run_bg": "#2F75B5", "btn_run_fg": "#FFFFFF",
            "btn_run_active": "#1A5FA0", "progress_trough": "#DDE0E6", "progress_fill": "#2F75B5",
            "toggle_track_off": "#C8CDD5", "toggle_track_on": "#2F75B5", "toggle_thumb": "#FFFFFF",
        },
        "dark": {
            "root_bg": "#1A1B2E", "toolbar_bg": "#252639", "sep_color": "#35364D",
            "body_bg": "#1A1B2E", "entry_bg": "#2E2F45", "entry_fg": "#D8D9E8",
            "entry_border": "#3E3F58", "text_bg": "#21223A", "text_fg": "#C4C5D6",
            "text_border": "#3E3F58", "label_fg": "#C4C5D6", "btn_folder_bg": "#35364D",
            "btn_folder_fg": "#C4C5D6", "btn_run_bg": "#2F75B5", "btn_run_fg": "#FFFFFF",
            "btn_run_active": "#1A5FA0", "progress_trough": "#2E2F45", "progress_fill": "#2F75B5",
            "toggle_track_off": "#555570", "toggle_track_on": "#2F75B5", "toggle_thumb": "#FFFFFF",
        },
    }
    FONT_ZH, FONT_MONO, FONT_SYM = "Microsoft JhengHei", "Consolas", "Arial"

    def __init__(self, root: "tk.Tk"):
        self.root = root
        self.root.title("Network Analysis Tool")
        self.root.resizable(True, True)

        # ── 依螢幕解析度等比例縮放視窗（基準 1920 × 1200）────────────────────
        scale = min(root.winfo_screenwidth() / 1920, root.winfo_screenheight() / 1200)
        self.root.geometry(f"{max(int(905 * scale), 480)}x{max(int(600 * scale), 380)}")
        self.root.minsize(max(int(480 * scale), 380), max(int(400 * scale), 260))

        self.selected_path = tk.StringVar()
        self.old_report = tk.StringVar()
        self.new_report = tk.StringVar()
        self.mode = "analyze"                             # analyze / compare
        self.compare_result: Optional[CompareResult] = None
        self.devices: List[Dict] = []
        self.aps: List[Dict] = []
        self._is_dark = False
        self._btn_enabled = True
        self._btn_hover = False
        self._resize_job = None
        self._ui_queue: queue.Queue = queue.Queue()      # 背景執行緒 → 主執行緒

        try:                                              # 應用程式圖示（找不到就略過）
            self.root.iconbitmap(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                              "icon_ironman.ico"))
        except Exception:
            pass

        self._setup_ui()
        self._apply_theme()
        self.root.after(100, self._poll_ui_queue)

    # ── UI 建立 ───────────────────────────────────────────────────────────────
    def _setup_ui(self):
        # 頂部工具列：選單圖示 + 標題 + Dark 切換
        self.toolbar = tk.Frame(self.root, height=48)
        self.toolbar.pack(fill="x", side="top")
        self.toolbar.pack_propagate(False)
        self.lbl_menu = tk.Label(self.toolbar, text="≡", font=(self.FONT_SYM, 16),
                                 padx=12, pady=10, cursor="hand2")
        self.lbl_menu.pack(side="left")
        self.lbl_title = tk.Label(self.toolbar, text=f"Network Analysis Tool by TING  {APP_VERSION}",
                                  font=(self.FONT_SYM, 12, "bold"), padx=4)
        self.lbl_title.pack(side="left")

        self.toggle_outer = tk.Frame(self.toolbar)
        self.toggle_outer.pack(side="right", padx=14)
        self.lbl_dark = tk.Label(self.toggle_outer, text="Dark", font=(self.FONT_SYM, 10), padx=6)
        self.lbl_dark.pack(side="left")
        self.toggle_canvas = tk.Canvas(self.toggle_outer, width=46, height=24, bd=0,
                                       highlightthickness=0, cursor="hand2")
        self.toggle_canvas.pack(side="left", pady=12)
        self.toggle_canvas.bind("<Button-1>", lambda _e: self._toggle_theme())

        self.sep_top = tk.Frame(self.root, height=1)
        self.sep_top.pack(fill="x")

        # 功能切換：分析 Log / 維護設備比對（V23）
        self.mode_outer = tk.Frame(self.root)
        self.mode_outer.pack(fill="x", padx=14, pady=(10, 0))
        self.mode_frame = tk.Frame(self.mode_outer, padx=1, pady=1)
        self.mode_frame.pack(side="left")
        self.mode_tabs: Dict[str, "tk.Label"] = {}
        for i, (mode, text) in enumerate((("analyze", "分析 Log"), ("compare", "維護設備比對"))):
            lbl = tk.Label(self.mode_frame, text=text, font=(self.FONT_ZH, 10), padx=16, pady=4,
                           cursor="hand2")
            lbl.grid(row=0, column=i, padx=(1 if i else 0, 0))
            lbl.bind("<Button-1>", lambda _e, m=mode: self.set_mode(m))
            self.mode_tabs[mode] = lbl

        # 資料夾選擇列
        self.folder_frame = tk.Frame(self.root, pady=10)
        self.folder_frame.pack(fill="x", padx=14)
        self.lbl_folder = tk.Label(self.folder_frame, text="選擇資料夾", font=(self.FONT_ZH, 10),
                                   width=8, anchor="w")
        self.lbl_folder.pack(side="left", padx=(0, 8))
        self.entry_border_frame = tk.Frame(self.folder_frame, bd=1, relief="flat")
        self.entry_border_frame.pack(side="left", fill="x", expand=True)
        self.entry_path = tk.Entry(self.entry_border_frame, textvariable=self.selected_path,
                                   state="readonly", font=(self.FONT_MONO, 10), relief="flat", bd=4)
        self.entry_path.pack(fill="x", expand=True)
        self.btn_folder_border = tk.Frame(self.folder_frame, padx=1, pady=1)
        self.btn_folder_border.pack(side="left", padx=(6, 0))
        self.btn_folder = tk.Button(self.btn_folder_border, text="🗁", font=(self.FONT_SYM, 13),
                                    relief="flat", bd=0, cursor="hand2", padx=8, pady=2,
                                    command=self.select_folder)
        self.btn_folder.pack()

        # 維護設備比對：上季 / 本季報表（切換到比對模式時取代資料夾列）
        self.compare_frame = tk.Frame(self.root, pady=4)
        self.compare_rows = [
            self._make_path_row(self.compare_frame, "上季報表", self.old_report,
                                lambda: self.select_report(self.old_report, "選擇上季維護報表")),
            self._make_path_row(self.compare_frame, "本季報表", self.new_report,
                                lambda: self.select_report(self.new_report, "選擇本季維護報表")),
        ]

        # 佈局順序：先把固定高度元件鎖在底部，Log 文字框最後 pack 取剩餘空間
        self.btn_frame = tk.Frame(self.root, height=70)
        self.btn_frame.pack(fill="x", side="bottom")
        self.btn_frame.pack_propagate(False)
        self.btn_canvas = tk.Canvas(self.btn_frame, width=220, height=46, bd=0,
                                    highlightthickness=0, cursor="hand2")
        self.btn_canvas.place(relx=0.5, rely=0.5, anchor="center")
        self.btn_canvas.bind("<Button-1>", lambda _e: self._on_run_btn_click())
        self.btn_canvas.bind("<Enter>", lambda _e: self._on_run_btn_hover(True))
        self.btn_canvas.bind("<Leave>", lambda _e: self._on_run_btn_hover(False))
        self.btn_frame.bind("<Configure>", self._on_btn_frame_resize)

        self.sep_bot = tk.Frame(self.root, height=1)
        self.sep_bot.pack(fill="x", side="bottom")

        self.prog_frame = tk.Frame(self.root)
        self.prog_frame.pack(fill="x", padx=14, pady=(0, 4), side="bottom")
        self.pb_style = ttk.Style()
        self.pb_style.layout("Flat.Horizontal.TProgressbar", [
            ("Horizontal.Progressbar.trough", {
                "children": [("Horizontal.Progressbar.pbar", {"side": "left", "sticky": "ns"})],
                "sticky": "nswe"})
        ])
        self.progress = ttk.Progressbar(self.prog_frame, orient="horizontal", mode="determinate",
                                        style="Flat.Horizontal.TProgressbar")
        self.progress.pack(fill="x")

        self.log_border_frame = tk.Frame(self.root, bd=1, relief="flat")
        self.log_border_frame.pack(fill="both", expand=True, padx=14, pady=(0, 4))
        self.txt_log = tk.Text(self.log_border_frame, font=(self.FONT_MONO, 9), relief="flat",
                               bd=6, wrap="word")
        self.txt_log.pack(fill="both", expand=True)

    def _make_path_row(self, parent, text: str, var: "tk.StringVar", command) -> Dict[str, object]:
        """標籤 + 唯讀路徑框 + 🗁 按鈕（比對模式的上季 / 本季報表列）。"""
        row = tk.Frame(parent, pady=6)
        row.pack(fill="x", padx=14)
        lbl = tk.Label(row, text=text, font=(self.FONT_ZH, 10), width=8, anchor="w")
        lbl.pack(side="left", padx=(0, 8))
        border = tk.Frame(row, bd=1, relief="flat")
        border.pack(side="left", fill="x", expand=True)
        entry = tk.Entry(border, textvariable=var, state="readonly", font=(self.FONT_MONO, 10),
                         relief="flat", bd=4)
        entry.pack(fill="x", expand=True)
        btn_border = tk.Frame(row, padx=1, pady=1)
        btn_border.pack(side="left", padx=(6, 0))
        btn = tk.Button(btn_border, text="🗁", font=(self.FONT_SYM, 13), relief="flat", bd=0,
                        cursor="hand2", padx=8, pady=2, command=command)
        btn.pack()
        return {"row": row, "label": lbl, "border": border, "entry": entry,
                "btn_border": btn_border, "btn": btn}

    def set_mode(self, mode: str):
        """切換「分析 Log」/「維護設備比對」；執行中不切換。"""
        if mode == self.mode or not self._btn_enabled:
            return
        self.mode = mode
        if mode == "compare":
            self.folder_frame.pack_forget()
            self.compare_frame.pack(fill="x", before=self.log_border_frame)
        else:
            self.compare_frame.pack_forget()
            self.folder_frame.pack(fill="x", padx=14, before=self.log_border_frame)
        self.progress.configure(value=0)
        self._apply_theme()

    # ── 主題 / 按鈕繪製 ───────────────────────────────────────────────────────
    def _theme(self) -> Dict[str, str]:
        return self.THEMES["dark" if self._is_dark else "light"]

    def _draw_toggle(self):
        """圓角 Toggle：左右半圓 + 中間矩形 + 白色拇指。"""
        t, c = self._theme(), self.toggle_canvas
        c.delete("all")
        w, h = 46, 24
        r = h // 2
        col = t["toggle_track_on"] if self._is_dark else t["toggle_track_off"]
        tx = w - r - 2 if self._is_dark else r + 2
        c.create_oval(0, 0, h, h, fill=col, outline="")
        c.create_oval(w - h, 0, w, h, fill=col, outline="")
        c.create_rectangle(r, 0, w - r, h, fill=col, outline="")
        c.create_oval(tx - r + 3, 3, tx + r - 3, h - 3, fill=t["toggle_thumb"], outline="")

    def _on_btn_frame_resize(self, event):
        """視窗寬度改變時節流（30ms）重繪按鈕。"""
        if self._resize_job:
            self.root.after_cancel(self._resize_job)
        self._resize_job = self.root.after(30, lambda: self._do_btn_resize(event.width))

    def _do_btn_resize(self, frame_width: int):
        self.btn_canvas.config(width=min(max(frame_width - 60, 180), 360))
        self._draw_run_btn(hover=self._btn_hover)

    def _draw_run_btn(self, hover: bool = False):
        """圓角矩形「產生分析報表 / 產生比對報表」按鈕（四角弧 + 兩個填充矩形）。"""
        c, t = self.btn_canvas, self._theme()
        c.delete("all")
        w, h, r = int(c.cget("width")), 46, 12
        if not self._btn_enabled:
            bg = "#888899" if self._is_dark else "#AABBD0"
        else:
            bg = t["btn_run_active"] if hover else t["btn_run_bg"]
        for x0, y0, start in ((0, 0, 90), (w - 2 * r, 0, 0), (0, h - 2 * r, 180), (w - 2 * r, h - 2 * r, 270)):
            c.create_arc(x0, y0, x0 + 2 * r, y0 + 2 * r, start=start, extent=90, fill=bg, outline="")
        c.create_rectangle(r, 0, w - r, h, fill=bg, outline="")
        c.create_rectangle(0, r, w, h - r, fill=bg, outline="")
        c.create_text(w // 2, h // 2, text="產生比對報表" if self.mode == "compare" else "產生分析報表",
                      fill=t["btn_run_fg"],
                      font=(self.FONT_ZH, 11, "bold"))

    def _on_run_btn_click(self):
        if self._btn_enabled:
            self.start_process()

    def _on_run_btn_hover(self, entering: bool):
        self._btn_hover = entering
        self._draw_run_btn(hover=entering)

    def _set_run_btn_state(self, enabled: bool):
        self._btn_enabled = enabled
        self.btn_canvas.config(cursor="hand2" if enabled else "watch")
        self._draw_run_btn(hover=False)

    def _toggle_theme(self):
        self._is_dark = not self._is_dark
        self._apply_theme()

    def _apply_theme(self):
        t = self._theme()
        for w in (self.root, self.folder_frame, self.prog_frame, self.btn_frame, self.btn_canvas,
                  self.mode_outer, self.compare_frame):
            w.config(bg=t["body_bg"] if w is not self.root else t["root_bg"])
        self.mode_frame.config(bg=t["entry_border"])
        for mode, lbl in self.mode_tabs.items():
            on = mode == self.mode
            lbl.config(bg=t["btn_run_bg"] if on else t["entry_bg"],
                       fg=t["btn_run_fg"] if on else t["label_fg"])
        for r in self.compare_rows:
            r["row"].config(bg=t["body_bg"])
            r["label"].config(bg=t["body_bg"], fg=t["label_fg"])
            r["border"].config(bg=t["entry_border"])
            r["entry"].config(bg=t["entry_bg"], fg=t["entry_fg"], readonlybackground=t["entry_bg"],
                              disabledbackground=t["entry_bg"], insertbackground=t["entry_fg"])
            r["btn_border"].config(bg=t["btn_folder_bg"] if self._is_dark else t["entry_border"])
            r["btn"].config(bg=t["btn_folder_bg"], fg=t["btn_folder_fg"],
                            activebackground=t["btn_run_bg"], activeforeground=t["btn_run_fg"])
        for w in (self.toolbar, self.toggle_outer, self.toggle_canvas):
            w.config(bg=t["toolbar_bg"])
        for w in (self.sep_top, self.sep_bot):
            w.config(bg=t["sep_color"])
        for w in (self.lbl_menu, self.lbl_title, self.lbl_dark):
            w.config(bg=t["toolbar_bg"], fg=t["label_fg"])
        self.lbl_folder.config(bg=t["body_bg"], fg=t["label_fg"])
        self.log_border_frame.config(bg=t["text_border"])
        self.entry_border_frame.config(bg=t["entry_border"])
        self.entry_path.config(bg=t["entry_bg"], fg=t["entry_fg"], readonlybackground=t["entry_bg"],
                               disabledbackground=t["entry_bg"], insertbackground=t["entry_fg"])
        # 淺色：資料夾按鈕外框顯示邊框色；深色：外框與按鈕同色
        self.btn_folder_border.config(bg=t["btn_folder_bg"] if self._is_dark else t["entry_border"])
        self.btn_folder.config(bg=t["btn_folder_bg"], fg=t["btn_folder_fg"],
                               activebackground=t["btn_run_bg"], activeforeground=t["btn_run_fg"])
        self.txt_log.config(bg=t["text_bg"], fg=t["text_fg"], insertbackground=t["text_fg"])
        self.pb_style.configure("Flat.Horizontal.TProgressbar", troughcolor=t["progress_trough"],
                                background=t["progress_fill"], borderwidth=0, thickness=6)
        self._draw_toggle()
        self._draw_run_btn(hover=self._btn_hover)

    # ── 背景執行緒 → 主執行緒 的訊息處理 ─────────────────────────────────────
    def log(self, msg: str):
        """任何執行緒皆可呼叫：訊息放入佇列，由主執行緒每 100ms 批次寫入。"""
        self._ui_queue.put(("log", msg))

    def _poll_ui_queue(self):
        lines = []
        try:
            while True:
                kind, value = self._ui_queue.get_nowait()
                if kind == "log":
                    lines.append(value)
                elif kind == "progress":
                    self.progress.configure(value=value)
                elif kind == "analyzed":
                    self._on_analyzed()
                elif kind == "compared":
                    self._on_compared(value)
                elif kind == "saved":
                    self._on_saved(value)
        except queue.Empty:
            pass
        if lines:
            self.txt_log.insert(tk.END, "\n".join(lines) + "\n")
            self.txt_log.see(tk.END)
        self.root.after(100, self._poll_ui_queue)

    # ── 主流程 ───────────────────────────────────────────────────────────────
    def select_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.selected_path.set(path)

    def select_report(self, var: "tk.StringVar", title: str):
        path = filedialog.askopenfilename(title=title, filetypes=[("Excel", "*.xlsx")])
        if path:
            var.set(path)

    def start_process(self):
        if self.mode == "compare":
            self.start_compare()
            return
        if not self.selected_path.get():
            messagebox.showwarning("警告", "請選擇資料夾")
            return
        self._set_run_btn_state(False)
        self.progress.configure(value=0)
        threading.Thread(target=self._analyze_worker, daemon=True).start()

    def _analyze_worker(self):
        """背景執行緒：解析整個資料夾（不碰任何 Tk 物件）。"""
        t0 = time.perf_counter()
        try:
            self.devices, self.aps, _ = analyze_folder(
                self.selected_path.get(), log=self.log,
                progress=lambda i, n: self._ui_queue.put(("progress", i / n * 100)))
            self.log(f"解析完成：{len(self.devices)} 台設備、{len(self.aps)} 台 AP，"
                     f"耗時 {time.perf_counter() - t0:.2f} 秒")
        except Exception as e:
            self.devices, self.aps = [], []
            self.log(f"[ERR] {e}")
        self._ui_queue.put(("analyzed", None))

    def _on_analyzed(self):
        """主執行緒：詢問儲存路徑，再交由背景執行緒寫檔。"""
        if not self.devices and not self.aps:
            messagebox.showwarning("警告", "無可分析資料")
            self._set_run_btn_state(True)
            return
        path = filedialog.asksaveasfilename(initialfile=default_report_name(), defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            self._set_run_btn_state(True)
            return

        def _save():
            try:
                write_excel(path, self.devices, self.aps)
                self._ui_queue.put(("saved", None))
            except Exception as e:
                self._ui_queue.put(("saved", e))
        threading.Thread(target=_save, daemon=True).start()

    # ── 維護設備比對（V23）────────────────────────────────────────────────────
    def start_compare(self):
        old, new = self.old_report.get(), self.new_report.get()
        if not old or not new:
            messagebox.showwarning("警告", "請選擇上季與本季的維護報表")
            return
        if os.path.abspath(old) == os.path.abspath(new):
            messagebox.showwarning("警告", "上季與本季選到同一份報表")
            return
        self._set_run_btn_state(False)
        self.progress.configure(value=0)
        threading.Thread(target=self._compare_worker, args=(old, new), daemon=True).start()

    def _compare_worker(self, old: str, new: str):
        """背景執行緒：讀取兩份報表並比對（不碰任何 Tk 物件）。"""
        t0 = time.perf_counter()
        try:
            self.compare_result = compare_reports(old, new)
            self._ui_queue.put(("progress", 100))
            for line in compare_log_lines(self.compare_result):
                self.log(line)
            self.log(f"比對完成，耗時 {time.perf_counter() - t0:.2f} 秒")
            self._ui_queue.put(("compared", None))
        except Exception as e:
            self.compare_result = None
            self.log(f"[ERR] {e}")
            self._ui_queue.put(("compared", e))

    def _on_compared(self, error: Optional[Exception]):
        """主執行緒：詢問比對報表儲存路徑，再交由背景執行緒寫檔。"""
        if error or self.compare_result is None:
            messagebox.showerror("錯誤", f"比對失敗：{error}")
            self._set_run_btn_state(True)
            return
        old, new = self.old_report.get(), self.new_report.get()
        path = filedialog.asksaveasfilename(initialfile=default_compare_name(old, new),
                                            initialdir=os.path.dirname(new), defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            self._set_run_btn_state(True)
            return
        result = self.compare_result

        def _save():
            try:
                write_compare_excel(path, result)
                self.log(f"已儲存：{path}")
                self._ui_queue.put(("saved", None))
            except Exception as e:
                self._ui_queue.put(("saved", e))
        threading.Thread(target=_save, daemon=True).start()

    def _on_saved(self, error: Optional[Exception]):
        if error:
            messagebox.showerror("錯誤", f"儲存失敗: {error}")
        else:
            messagebox.showinfo("成功", "比對報表已儲存" if self.mode == "compare" else "報表已儲存")
        self._set_run_btn_state(True)


# ==============================================================================
# 17. 程式進入點：無參數 → GUI；帶資料夾參數 → CLI 批次模式
# ==============================================================================

def run_compare_cli(old_path: str, new_path: str, out: Optional[str] = None) -> str:
    """CLI 維護設備比對：印出摘要並輸出比對報表，回傳輸出路徑。"""
    try:
        result = compare_reports(old_path, new_path)
    except CompareError as e:
        sys.exit(f"[ERR] {e}")
    out = out or os.path.join(os.path.dirname(os.path.abspath(new_path)),
                              default_compare_name(old_path, new_path))
    write_compare_excel(out, result)
    print("\n".join(compare_log_lines(result)))
    print(f"已輸出：{out}")
    return out


def main(argv: Optional[List[str]] = None):
    ap = argparse.ArgumentParser(
        description=f"Cisco / Aruba / FortiGate Log 分析工具 {APP_VERSION}",
        epilog="維護設備比對：--compare 上季.xlsx 本季.xlsx [-o 比對.xlsx]；"
               "分析後順便比對：<Log資料夾> -o 本季.xlsx --compare-with 上季.xlsx")
    ap.add_argument("folder", nargs="?", help="Log 資料夾（省略則開啟 GUI）")
    ap.add_argument("-o", "--output", help="輸出 Excel 路徑（預設存在 Log 資料夾內）")
    ap.add_argument("--compare", nargs=2, metavar=("上季.xlsx", "本季.xlsx"),
                    help="維護設備比對：比較兩份本工具產出的維護報表")
    ap.add_argument("--compare-with", metavar="上季.xlsx",
                    help="分析 Log 資料夾後，再與這份上季報表比對")
    args = ap.parse_args(argv)

    if args.compare:
        if args.folder or args.compare_with:
            ap.error("--compare 不能與 Log 資料夾或 --compare-with 同時使用")
        run_compare_cli(*args.compare, out=args.output)
        return
    if args.compare_with and not args.folder:
        ap.error("--compare-with 需要搭配 Log 資料夾")

    if args.folder:
        t0 = time.perf_counter()
        devices, aps, _ = analyze_folder(args.folder)
        out = args.output or os.path.join(args.folder, default_report_name())
        write_excel(out, devices, aps)
        print(f"完成：{len(devices)} 台設備、{len(aps)} 台 AP → {out}（{time.perf_counter() - t0:.2f} 秒）")
        if args.compare_with:
            run_compare_cli(args.compare_with, out)
        return

    if tk is None:
        sys.exit(f"此環境沒有 tkinter，請改用 CLI：python {os.path.basename(__file__)} <Log資料夾>")
    root = tk.Tk()
    CiscoParserApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
