# =============================================================================
# Design By Ting Chen
# Cisco_Log_Analysis.py
#
# 功能：讀取 Cisco / Aruba 設備 Log 檔案，解析設備資訊並彙整輸出成 Excel 報表。
#
# 支援設備：
#   - Cisco C9800 WLC / EWC（IOS-XE，含 HA SSO 雙機判定）
#   - Cisco AireOS WLC（2504 / 3504 / 5520）
#   - Catalyst IOS / IOS-XE Switch（C2960 / C1000 / C3850 / C9K，含 Stack 堆疊）
#   - Cisco Nexus NX-OS（N5K / N9K）
#   - Cisco Catalyst 1300（show tech-support 彙整格式）
#   - Aruba / HPE AOS-CX Switch（6300 / 6400）
#   - WLC 底下的 AP（C9800：show ap summary + show ap config general；
#                    AireOS：show ap join stats summary all + show ap inventory all）
#
# 使用方式：
#   GUI：python Cisco_Log_Analysis.py             （或 pythonw 無黑視窗）
#   CLI：python Cisco_Log_Analysis.py <Log資料夾> [-o 報表.xlsx]
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
from dataclasses import dataclass
from datetime import datetime
from typing import Callable, Dict, Iterable, List, Optional, Tuple

# ── 第三方函式庫 ──────────────────────────────────────────────────────────────
import pandas as pd
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

APP_VERSION = "V19"
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
ENV_BAD_REGEX = re.compile(
    r"\b(FAIL|FAULT|ALARM|CRITICAL|BAD|OFF|NOT\s+OK|DOWN|ERROR|ABSENT)\b", re.I
)
# C1300 不含 OFF：PoE 埠未接設備時「Oper: Off」屬正常
ENV_BAD_REGEX_C1300 = re.compile(
    r"\b(FAIL|FAULT|ALARM|CRITICAL|BAD|NOT\s+OK|ERROR|ABSENT)\b", re.I
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
                     self._parse_env, self._parse_syslog, self._parse_eth_mac):
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
            group = "HA:" + "|".join(sorted(mac_key(u.eth_mac) or self.source for u in units))
            order_of = lambda u: (0 if u.ha_state.lower().startswith("active") else 1, u.member_id)
        else:
            kind, group = "stack", f"FILE:{self.source}:{id(self)}"
            order_of = lambda u: (0, u.member_id)

        units = sorted(units, key=order_of)
        rows = []
        for i, u in enumerate(units):
            first = (i == 0)
            mac = u.eth_mac if u.eth_mac != "N/A" else (self.info.eth_mac if first else MARK)
            rows.append({
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
                "_ha_key":          mac_key(u.eth_mac) if self.info.is_ha else "",
                "_ha_local":        u.is_local,
            })
        return rows


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
        """show environment 分成 Fan / Power Supply 區段，各自判斷 Status 欄（溫度區段不列入）。"""
        section, fan_bad, pwr_bad = "", False, False
        for line in self.cmd("show environment").splitlines():
            head = re.match(r"^(Fan|Power Supply|Temperature|Power Usage Summary)\b", line)
            if head:
                section = head.group(1)
                continue
            if ENV_BAD_REGEX.search(line):
                fan_bad |= section == "Fan"
                pwr_bad |= section == "Power Supply"
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
        val = self._sys("CPU Average Usage", r"(\d+)")
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
        """Fan Status == OK 為正常；Internal / Mgig / External 任一溫度超過 Alarm 上限 → 電源異常。"""
        fan = self._sys("Fan Status", r"(\w+)")
        if fan:
            self.fan_status = "正常" if fan.upper() == "OK" else "異常"
        pwr_bad = False
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
# 10. Aruba / HPE AOS-CX 解析器（6300 / 6400）
# ==============================================================================

class ArubaParser(BaseParser):
    """
    AOS-CX show tech：區塊邊界為「*****\\nCommand : show X\\n*****」。
      Hostname / 型號 / 序號 / CPU / Memory / MAC → show system
      韌體 → show version（去除 FL. 前綴）；管理 IP → interface mgmt 的 ip static
    """

    DEVICE_TYPE = "Aruba"
    TECH_FALLBACK = False

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

    def _parse_cpu(self):
        val = rsearch(r"CPU Util\s*\(%\s*avg\s*5\s*min\)\s*:\s*(\d+)", self.sys_sec, flags=0)
        self.info.cpu = f"{val}%" if val else "N/A"

    def _parse_memory(self):
        val = rsearch(r"Memory Usage\s*\(%\)\s*:\s*(\d+)", self.sys_sec, flags=0)
        self.info.memory = f"Usage: {val}%" if val else "N/A"      # AOS-CX 只提供百分比

    def _parse_inventory(self):
        """型號取 Product Name 前三段（JL667A 6300F 48G），其餘為 uplink 規格描述。"""
        name = rsearch(r"Product Name\s*:\s*([^\r\n]+)", self.sys_sec, flags=0)
        parts = name.split() if name else []
        pid = " ".join(parts[:3]) if len(parts) >= 3 else (name or "Unknown-Aruba")
        sn = rsearch(r"Chassis Serial Nbr\s*:\s*(\S+)", self.sys_sec, flags=0, default="Unknown")
        self.units.append(Unit(1, pid, sn))

    def _parse_env(self):
        bad = re.compile(r"\b(FAIL|FAULT|ALARM|ERROR|ABSENT|CRITICAL)\b", re.I)
        fan = rsearch(r"Fan information(.*?)(?=\n\nshow environment|\Z)", self.env_sec, flags=re.S | re.I, default="")
        psu = rsearch(r"show environment power-supply\s*\n.*?\n.*?\n.*?\n(.*?)(?=\n\n|\Z)",
                      self.env_sec, flags=re.S | re.I, default="")
        self.fan_status = "異常" if bad.search(fan) else "正常"
        self.pwr_status = "異常" if bad.search(psu) else "正常"

    def _parse_eth_mac(self):
        """HPE 格式「aabbcc-000001」與標準格式皆轉為 aa:bb:cc:00:00:01。"""
        raw = rsearch(r"Base\s+MAC\s+Address\s*[:\s]+([0-9a-fA-F:.]{11,17}|[0-9a-fA-F]{6}-[0-9a-fA-F]{6})",
                      self.sys_sec)
        if raw:
            self.info.eth_mac = mac_format(raw, upper=False) if mac_key(raw) else raw

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
    "C1300":  C1300Parser,
}
PROMPT_BASED = (IOSSwitchParser, NexusParser, WLCParser)   # 可共用 detect 階段建好的提示符索引


def detect_type(content: str) -> Tuple[str, Optional[CommandIndex]]:
    """
    由特定到一般判斷設備類型；IOS 家族會順便回傳已建立的提示符索引供解析器重用。

      1. AireOS : 行首「(xxx) >」提示符出現 ≥ 2 次
                  （V18 只要全文含 RTOS Version 就判定，CDP 鄰居是 AireOS WLC 的交換器會被誤判）
      2. Aruba  : 含「AOS-CX」且有「Command : show」區塊
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

LOG_EXTS = (".log", ".txt")


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

def build_sheets(devices: List[Dict], aps: List[Dict]) -> Dict[str, pd.DataFrame]:
    """
    Report      ：所有設備 + AP（不含 MAC / Source File）
    AP Info     ：AP 詳細資訊 + Source File
    WLC Info    ：C9800 / AireOS（含 HA Role）+ Source File
    Switch Info ：IOS / NX-OS / C1300 / Aruba + Source File
    """
    ap_report = [{"Hostname": a["ap_name"], "Model": a["ap_model"], "OOB IP": a["ip_address"],
                  "Serial Number": a["serial"], "Firmware Version": a["firmware"],
                  "CPU Usage": "N/A", "Memory Usage": "N/A", "Fan Status": "正常",
                  "Power Status": "正常", "System Log": ""} for a in aps]
    ap_info = [{"AP Name": a["ap_name"], "AP Model": a["ap_model"], "Ethernet MAC": a["eth_mac"],
                "IP Address": a["ip_address"], "AP Serial Number": a["serial"],
                "Software Version": a["firmware"], "Source File": a["source"]} for a in aps]
    return {
        "Report":      pd.DataFrame(devices + ap_report, columns=REPORT_COLS),
        "AP Info":     pd.DataFrame(ap_info, columns=AP_COLS),
        "WLC Info":    pd.DataFrame([r for r in devices if r["_device_type"] in WLC_TYPES], columns=WLC_COLS),
        "Switch Info": pd.DataFrame([r for r in devices if r["_device_type"] in SWITCH_TYPES], columns=SWITCH_COLS),
    }


def _cell_width(value) -> int:
    """以 Big5 位元組數估算顯示寬度（中文約 2 字元寬），多行取最長行。"""
    if value is None or value == "":
        return 0
    return max(len(line.encode("big5", errors="ignore")) for line in str(value).split("\n"))


def write_excel(path: str, devices: List[Dict], aps: List[Dict]):
    """寫入 Excel 並套用格式：藍底白字表頭、微軟正黑體、細框線、自動欄寬（上限 100）。"""
    header_fill = PatternFill("solid", fgColor="2F75B5")
    header_font = Font(name="微軟正黑體", size=12, bold=True, color="FFFFFF")
    data_font = Font(name="微軟正黑體", size=12)
    header_align = Alignment(horizontal="center")
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    sheets = build_sheets(devices, aps)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=name)
            ws = writer.sheets[name]
            for col in ws.iter_cols():
                width = 0
                for cell in col:
                    if cell.row == 1:
                        cell.fill, cell.font, cell.alignment = header_fill, header_font, header_align
                    else:
                        cell.font = data_font
                    cell.border = border
                    width = max(width, _cell_width(cell.value))
                ws.column_dimensions[col[0].column_letter].width = min(int(width * 1.2) + 4, 100)


def default_report_name() -> str:
    return f"Network_Analysis_{datetime.now():%m%d_%H%M}.xlsx"


# ==============================================================================
# 15. GUI（Tkinter）
# ==============================================================================

class CiscoParserApp:
    """
    操作流程：選擇資料夾 → 產生分析報表（背景執行緒解析）→ 選擇儲存路徑 → 輸出 Excel。

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
        """圓角矩形「產生分析報表」按鈕（四角弧 + 兩個填充矩形）。"""
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
        c.create_text(w // 2, h // 2, text="產生分析報表", fill=t["btn_run_fg"],
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
        for w in (self.root, self.folder_frame, self.prog_frame, self.btn_frame, self.btn_canvas):
            w.config(bg=t["body_bg"] if w is not self.root else t["root_bg"])
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

    def start_process(self):
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

    def _on_saved(self, error: Optional[Exception]):
        if error:
            messagebox.showerror("錯誤", f"儲存失敗: {error}")
        else:
            messagebox.showinfo("成功", "報表已儲存")
        self._set_run_btn_state(True)


# ==============================================================================
# 16. 程式進入點：無參數 → GUI；帶資料夾參數 → CLI 批次模式
# ==============================================================================

def main(argv: Optional[List[str]] = None):
    ap = argparse.ArgumentParser(description=f"Cisco / Aruba Log 分析工具 {APP_VERSION}")
    ap.add_argument("folder", nargs="?", help="Log 資料夾（省略則開啟 GUI）")
    ap.add_argument("-o", "--output", help="輸出 Excel 路徑（預設存在 Log 資料夾內）")
    args = ap.parse_args(argv)

    if args.folder:
        t0 = time.perf_counter()
        devices, aps, _ = analyze_folder(args.folder)
        out = args.output or os.path.join(args.folder, default_report_name())
        write_excel(out, devices, aps)
        print(f"完成：{len(devices)} 台設備、{len(aps)} 台 AP → {out}（{time.perf_counter() - t0:.2f} 秒）")
        return

    if tk is None:
        sys.exit("此環境沒有 tkinter，請改用 CLI：python Cisco_Log_Analysis.py <Log資料夾>")
    root = tk.Tk()
    CiscoParserApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
