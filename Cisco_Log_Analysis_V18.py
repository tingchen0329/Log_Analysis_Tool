# =============================================================================
# Design By Ting Chen
# Cisco_Log_Analysis_V18.py
#
# 功能：讀取 Cisco / Aruba 設備 Log 檔案，解析設備資訊並彙整輸出成 Excel 報表。
#
# 支援設備：
#   - Cisco C9800 WLC（IOS-XE based Wireless LAN Controller）
#   - Cisco WLC 2504 / 3504 / 5520（AireOS based，傳統 WLC）
#   - Catalyst IOS / IOS-XE Switch（C2960 / C1000 / C3850 / C9K，含 Stack 堆疊）
#   - Cisco Nexus NX-OS（N5K / N9K 等，含 system: version 格式修正）
#   - Cisco Catalyst 1300（Linux-based ROS，使用 show tech-support 彙整格式）
#   - Aruba / HPE AOS-CX Switch（6300 / 6400 等系列）
#   - WLC 底下掛載的 AP（C9800 AP 從 show ap summary；AireOS AP 從 show ap join stats）
#
# V16 修正與新增：
#   1. NexusParser：
#      - _parse_firmware()：補充 N5K/N6K 格式 "system: version X"（N9K 才有 "NXOS: version"）
#      - _parse_cpu()：補充 N5K 格式 "CPU states: X% user, X% kernel"（N9K 才有 "five minutes:"）
#      - _parse_oob()：BaseParser 新增策略 E，NX-OS 專用：
#        取 "ip address X/mask secondary" 的第一個 secondary IP（N5K 管理 IP 通常設為 secondary）
#        解決 mgmt0 /30 P2P link 被誤認為管理 IP 的問題
#   2. AireOSParser（新增）：
#      支援 Cisco AireOS WLC（2504/3504/5520），Log 邊界為 "(Cisco Controller) >" 提示符
#      AP 資料來自 "show ap join stats summary all" + "show ap inventory all"
#   3. ArubaParser（新增）：
#      支援 Aruba / HPE AOS-CX Switch（show tech 格式，邊界為 "Command : show X"）
#      detect_type 識別 "AOS-CX" 關鍵字，回傳 "Aruba"
# =============================================================================


# ── 標準函式庫匯入 ──────────────────────────────────────────────────────────
import os           # 作業系統路徑操作：資料夾走訪、路徑合併
import re           # 正規表示式（Regex）：從 Log 文字擷取關鍵欄位
import threading    # 多執行緒：讓耗時解析在背景執行，避免 GUI 視窗凍結
import queue        # queue.Queue：Thread-Safe 的訊息佇列，用於批次更新 Log 介面

from dataclasses import dataclass          # dataclass 裝飾器：自動產生 __init__ 等方法
from datetime import datetime              # datetime：產生含日期時間的報表檔名
from typing import Dict, List, Any        # 型別提示：提升可讀性與 IDE 自動完成

# ── 第三方函式庫匯入 ─────────────────────────────────────────────────────────
import pandas as pd                        # pandas：建立 DataFrame、寫入 Excel

import tkinter as tk                       # tkinter：Python 內建 GUI 函式庫
from tkinter import filedialog, ttk, messagebox

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ==============================================================================
# 1. 全域設定常數
# ==============================================================================

# ── EXCLUDE_LOG_PREFIXES：Syslog 雜訊過濾清單 ────────────────────────────────
EXCLUDE_LOG_PREFIXES = [
    "%LINK-3-UPDOWN",
    "%DAEMON-3-SYSTEM_MSG",
    "%L2FM-4-L2FM_MAC_MOVE2",
    "%URIB-4-SYSLOG_SL_MSG_WARNING",
    "%SW_MATM-4-MACFLAP_NOTIF",
    "%SMART_LIC-3-COMM_FAILED",
]

# ── SYSLOG_REGEX：擷取嚴重等級 0~4 的 Syslog 訊息 ────────────────────────────
SYSLOG_REGEX = re.compile(r"%[A-Z0-9_]+-[0-4]-[A-Z0-9_]+:.*")

# ── ENV_BAD_REGEX：環境異常關鍵字偵測 ────────────────────────────────────────
ENV_BAD_REGEX = re.compile(
    r"\b(FAIL|FAULT|ALARM|CRITICAL|BAD|OFF|NOT\s+OK|DOWN|ERROR|ABSENT)\b", re.I
)

# ── ENV_BAD_REGEX_C1300：C1300 環境異常偵測（排除 OFF，避免誤判 PoE 介面狀態）─
# C1300 的 show power inline 輸出中，PoE 埠若未接設備會顯示「Oper: Off」，
# 屬於正常狀態，若使用含 OFF 的 ENV_BAD_REGEX 會造成誤判。
# C1300 的環境健康狀態只從 show system 的 Fans Status / Temperature Status 判斷，
# 這兩個區段的正常值為「OK」，異常才會出現 FAIL / ERROR 等關鍵字。
ENV_BAD_REGEX_C1300 = re.compile(
    r"\b(FAIL|FAULT|ALARM|CRITICAL|BAD|NOT\s+OK|ERROR|ABSENT)\b", re.I
)

# ── C1300_SYSLOG_REGEX：擷取 C1300 ROS 格式的 Syslog 訊息 ────────────────────
# C1300 的 Syslog 格式與 IOS 完全不同：
#   IOS 格式：%FACILITY-SEVERITY_NUM-MNEMONIC: message
#             例：%OSPF-3-ADJCHG: ...
#   C1300 格式：DD-Mon-YYYY HH:MM:SS :%FACILITY-SEVERITY_LETTER-MNEMONIC: message
#               例：07-Apr-2026 01:41:50 :%LINK-W-Down: gi1/0/14
# 嚴重等級字母對照：
#   E = Emergency / Alert / Critical / Error（對應 IOS 等級 0~3）
#   W = Warning（對應 IOS 等級 4）
#   N = Notice、I = Informational、D = Debug（一般不需要收集）
# 此 Regex 只擷取 E（Error）和 W（Warning）等級的訊息，
# 與現有 IOS 的 SYSLOG_REGEX 收集等級 0~4 的概念一致。
C1300_SYSLOG_REGEX = re.compile(
    r"\d{2}-\w{3}-\d{4}\s+\d{2}:\d{2}:\d{2}\s+:(%[A-Z0-9_]+-[EW]-[A-Z0-9_]+:.*)"
)

# ── C1300_EXCLUDE_PREFIXES：C1300 Syslog 雜訊過濾清單 ───────────────────────
# C1300 的 LINK 上下線、STP 拓撲變更、時間同步等屬於正常事件，排除以降低雜訊。
C1300_EXCLUDE_PREFIXES = [
    "%LINK-W-Down",        # 介面 Link 斷線（插拔線纜即觸發，屬正常）
    "%LINK-I-Up",          # 介面 Link 連線（同上）
    "%STP-W-PORTSTATUS",   # STP 埠狀態變化（正常拓撲收斂過程）
    "%TIMEMANAGE-",        # 時間同步相關（NTP/RTC 更新）
    "%MLDP-I-",            # MLD Proxy 狀態訊息
    "%Entity-I-",          # 硬體實體配置變更通知
    "%INIT-I-",            # 系統初始化完成通知
]


# ==============================================================================
# 2. 資料容器類別（@dataclass 自動生成 __init__ 等方法）
# ==============================================================================

@dataclass
class StackMember:
    """代表 Switch Stack 中的單一成員。"""
    member_id: int
    model: str = "Unknown"
    serial: str = "Unknown"
    eth_mac: str = "N/A"          # 各 Stack 成員的實體 MAC（從 show version Switch N 區塊比對序號取得）


@dataclass
class DeviceInfo:
    """儲存單台網路設備解析完成後的所有資訊。"""
    hostname: str = "Unknown"
    oob_ip: str = "N/A"
    firmware: str = "N/A"
    cpu: str = "N/A"
    memory: str = "N/A"
    device_type: str = "Unknown"
    eth_mac: str = "N/A"          # 設備基礎 MAC 位址（用於 WLC Info / Switch Info 分頁）
    is_stack: bool = False
    unit_count: int = 1


@dataclass
class APInfo:
    """儲存 WLC 底下單台 AP 的解析結果。"""
    ap_name: str = "Unknown"
    ap_model: str = "Unknown"
    eth_mac: str = "N/A"
    ip_address: str = "N/A"
    serial: str = "N/A"
    firmware: str = "N/A"


# ==============================================================================
# 3. 公用輔助函式
# ==============================================================================

def truncate_ap_firmware(version: str) -> str:
    """
    將 AP 韌體版本號截取至第三段數字，捨棄第四段（build 號）。

    Cisco AP 版本號格式為「主版本.次版本.修補版本.組建號」，
    例如 17.9.8.5，其中第四段 .5 為 build 號，對日常管理意義不大。
    截短後可讓報表更整潔、易讀。

    規則示例：
      17.9.8.5  → 17.9.8   （截掉第四段 .5）
      17.9.4.27 → 17.9.4   （截掉第四段 .27）
      17.9.8    → 17.9.8   （只有三段，保持不變）
      N/A       → N/A      （無版本資訊，原樣回傳）

    Regex 說明：
      r"(\d+\.\d+\.\d+)"
        \d+   → 一或多位數字（第一段主版本）
        \.    → 字面上的點（版本號分隔符）
        \d+   → 第二段次版本
        \.    → 點
        \d+   → 第三段修補版本
      match.group(1) 擷取這三段，後面的 .27 等自動被忽略。

    Args:
        version: 原始版本字串，例如 "17.9.8.5"

    Returns:
        截短後的版本字串，例如 "17.9.8"
    """
    m = re.match(r"(\d+\.\d+\.\d+)", version)
    if m:
        return m.group(1)   # 成功匹配，回傳前三段
    return version           # 無法匹配（如 "N/A"），原樣回傳


# ==============================================================================
# 4. 解析器基礎類別（BaseParser）
# ==============================================================================

class BaseParser:
    """
    所有設備解析器的共用父類別（Template Method Pattern）。

    BaseParser 定義解析流程的「骨架」（run_all 方法），
    各子類別只需覆寫各自格式不同的方法（如 _parse_firmware），
    通用邏輯（如 _parse_oob、_parse_syslog）由父類別統一實作。
    """

    def __init__(self, content: str):
        """
        初始化解析器：儲存 Log 內容、建立資料容器、預先擷取常用區塊。

        Args:
            content: 整份 Log 檔案的完整文字（utf-8 解碼後的字串）
        """
        self.content = content
        self.info = DeviceInfo()
        self.members: List[StackMember] = []
        self.fan_status = "Normal"
        self.pwr_status = "Normal"
        self.syslog = ""

        # 預先擷取最常用的兩個指令區塊，避免重複解析（效能優化）
        self.ver_sec = self._get_command_block("show version")
        self.inv_sec = self._get_command_block("show inventory")

    def _get_command_block(self, cmd: str) -> str:
        """
        從 Log 中精準擷取單一 CLI 指令的完整輸出區塊。

        Log 結構特徵：
          每個指令輸出的邊界是 CLI 提示符（hostname#），例如：
            SW1#show version      ← 指令行
            Cisco IOS Software... ← 輸出開始
            ...
            SW1#show inventory    ← 下一個提示符（邊界）

        Regex 解析策略：
          1. 找到「cmd\b」（\b 確保全詞匹配，避免 "version" 匹配到 "versioning"）
          2. 跨行抓取後續內容（re.S 讓「.」可匹配換行）
          3. 遇到下一個提示符（\n\S+#）或字串結尾（\Z）停止
          4. 使用正向前視 (?=...) 確保邊界提示符本身不被包含在結果中

        Args:
            cmd: CLI 指令名稱，例如 "show version"

        Returns:
            指令輸出純文字（去除首尾空白），找不到則回傳空字串 ""
        """
        pattern = re.compile(rf"{cmd}\b.*?\n(.*?)(?=\n\S+#|\Z)", re.S | re.I)
        match = pattern.search(self.content)
        return match.group(1).strip() if match else ""

    def format_mem(self, total: Any, used: Any, free: Any, usage_pct: Any) -> str:
        """
        將記憶體數值格式化為統一的顯示字串。

        各平台記憶體單位不同（IOS = Bytes，NX-OS = KB），
        此方法不做換算，由各子類別在呼叫前自行加上單位後綴。

        Returns:
            格式：「Total：X | Used：Y | Free：Z (P%)」
        """
        return f"Total：{total} | Used：{used} | Free：{free} ({usage_pct}%)"

    def run_all(self):
        """
        依序執行完整解析流程（Template Method Pattern 的核心）。

        呼叫順序固定，確保各方法的依賴關係正確：
          hostname → oob_ip → firmware → cpu → memory → inventory → env → syslog → eth_mac
        """
        self._parse_hostname()
        self._parse_oob()
        self._parse_firmware()
        self._parse_cpu()
        self._parse_memory()
        self._parse_inventory()
        self._parse_env()
        self._parse_syslog()
        self._parse_eth_mac()          # 最後解析 MAC，確保 device_type 已設定

    def _parse_eth_mac(self):
        """
        解析設備基礎 Ethernet MAC 位址。

        BaseParser 實作（適用於 IOS / IOS-XE Switch 及 C9800 WLC）：
          show version 固定輸出格式：
            「Base ethernet MAC Address       : AA:BB:CC:00:00:01」
            「Base Ethernet MAC Address       : aabb.cc00.0002」  ← IOS-XE Cisco dot 格式

        子類別（AireOS / Aruba / C1300）各自覆寫此方法以對應不同格式。
        NX-OS 在 show version 無標準 MAC 行，嘗試從 show interface mgmt0 抓取。
        """
        # IOS / IOS-XE：Base ethernet/Ethernet MAC Address
        m = re.search(
            r"Base\s+[Ee]thernet\s+MAC\s+[Aa]ddress\s*[:\s]+([0-9a-fA-F:.]{11,17})",
            self.content, re.I
        )
        if m:
            self.info.eth_mac = m.group(1).strip()
            return

        # NX-OS fallback：mgmt0 介面的 MAC
        m = re.search(
            r"interface\s+mgmt0.*?address:\s*([0-9a-fA-F:]{17})",
            self.content, re.S | re.I
        )
        if m:
            self.info.eth_mac = m.group(1).strip()

    def _parse_hostname(self):
        """
        解析設備主機名稱。

        方法一（優先）：hostname 指令
          running-config 格式：「hostname C2960-LAB-01」
          → 擷取 "C2960-LAB-01"（re.M 使 ^ 匹配每行行首）

        方法二（備案）：CLI 提示符反推
          Log 中每條指令前都有「<名稱>#」或「<名稱>>」的提示符
          → 擷取第一個出現的提示符名稱
        """
        m = re.search(r"^hostname\s+(\S+)", self.content, re.M | re.I)
        if m:
            self.info.hostname = m.group(1)
        else:
            m_prompt = re.search(r"\n(\S+)(?:#|>)", self.content)
            if m_prompt:
                self.info.hostname = m_prompt.group(1)

    def _parse_oob(self):
        """
        解析 OOB（Out-of-Band，帶外管理）介面的 IP 位址，採四段優先策略。

        ──────────────────────────────────────────────────────────
        「為什麼需要四段策略？」

          不同 Cisco 平台的管理 IP 設定方式差異很大：
          ┌──────────────────┬──────────────────────────────────────┐
          │ 設備類型          │ 管理 IP 設定位置                      │
          ├──────────────────┼──────────────────────────────────────┤
          │ C9K / C3850 Core │ interface GigabitEthernet0/0          │
          │                  │ （獨立的帶外管理介面，通常在 Mgmt-vrf）│
          ├──────────────────┼──────────────────────────────────────┤
          │ C2960 Access SW  │ interface Vlan<N>                    │
          │                  │ （無 L3 Mgmt 介面，Vlan SVI 作管理）   │
          │                  │ 透過 ip default-gateway 識別管理 Vlan  │
          ├──────────────────┼──────────────────────────────────────┤
          │ C1000 Access SW  │ interface Vlan<N>                    │
          │                  │ （支援 ip routing，第一個 Vlan 通常是  │
          │                  │  管理 Vlan，從 ip brief 第一行找）      │
          └──────────────────┴──────────────────────────────────────┘
        ──────────────────────────────────────────────────────────

        ── 策略 A：專用 Mgmt 介面（最直接，優先採用）──────────────────────
          適用：C3850 / C9K / WLC
          常見介面名稱：
            GigabitEthernet0/0（running-config 中 vrf forwarding Mgmt-vrf）
            Management0（部分 C9K）、Mgmt0（NX-OS）
          Log 範例：
            interface GigabitEthernet0/0
             vrf forwarding Mgmt-vrf
             ip address 198.51.100.246 255.255.0.0   ← 擷取此 IP

        ── 策略 B：ip default-gateway 前綴對應 Vlan SVI ──────────────────
          適用：C2960 等老式 Access Switch
          原理：
            1. 抓 ip default-gateway（如 192.0.2.248）
            2. 取前三段作為網段前綴（192.0.2）
            3. 找 Vlan SVI 中 ip address 符合此前綴的那個
            4. 該 SVI 的 IP（如 192.0.2.31）= 管理 IP
          說明：
            Access Switch 的管理 Vlan IP 和 default-gateway 通常在同一 /24 網段，
            比對前三段是最簡單可靠的方式。
          Log 範例：
            ip default-gateway 192.0.2.248   → prefix = "192.0.2"
            interface Vlan3
             ip address 192.0.2.31 255.255.252.0  ← 前三段符合，取此 IP

        ── 策略 C：show ip interface brief 第一個有效 Vlan IP ─────────────
          適用：C1000 等（策略 B 失敗時的備援）
          原理：
            show ip brief 列出所有介面的即時 IP 與狀態，
            取第一個「含 Vlan 字樣 + 有效 IP + 非 unassigned」的行。

        ── 策略 D：running-config 第一個 Vlan SVI IP（最後備案）──────────
          適用：以上三種皆失敗的特殊情形
          缺點：若第一個 Vlan 不是管理 Vlan，可能取到錯誤 IP。
        """
        # ── 策略 A：專用 Mgmt 介面（排除 /30、/31 P2P link）────────────
        # 問題背景（NX-OS N5K）：
        #   N5K 的 mgmt0 可能設為 /30 的 P2P 連結（如 100.64.255.1/30），
        #   並非真正的管理 IP。需排除 prefix-length >= 30 的情形，
        #   避免抓到 P2P link 而非管理 IP。
        # NX-OS 格式：ip address X.X.X.X/prefix（帶斜線，無 mask）
        # IOS 格式  ：ip address X.X.X.X mask（空格分隔）
        # 解法：只接受 CIDR prefix 1~29，或傳統 mask 格式（兩種皆相容）
        m = re.search(
            r"interface\s+(?:Mgmt\w*|GigabitEthernet0(?:/\d+)?|Management\d+|gi0)\s"
            r".*?ip address\s+(\d+\.\d+\.\d+\.\d+)"
            r"(?:/([0-9]+)|\s)",   # 擷取 prefix-length（CIDR 格式）或空格（mask 格式）
            self.content, re.S | re.I
        )
        if m:
            prefix_len = int(m.group(2)) if m.group(2) else 0   # mask 格式時 group(2) 為 None
            # prefix == 0 代表是 mask 格式（IOS），不是 /30 P2P，直接接受
            # prefix >= 30 是 P2P link（/30 或 /31），跳過繼續嘗試其他策略
            if prefix_len < 30:
                self.info.oob_ip = m.group(1)
                return
            # 若 mgmt 介面是 /30 P2P，不 return，繼續往下走策略 E

        # ── 策略 E：NX-OS secondary IP（N5K / N6K 管理 IP 通常設為 secondary）─
        # NX-OS 多重 IP 格式：
        #   ip address 100.64.10.248/16        ← primary（路由用途）
        #   ip address 198.51.100.248/16 secondary ← secondary（管理 IP）
        # 管理人員常將「帶外管理段的 IP」設為 secondary，
        # 因此取第一個 secondary IP 作為管理 IP 是最準確的策略。
        sec_m = re.search(
            r"ip address\s+(\d+\.\d+\.\d+\.\d+)/\d+\s+secondary",
            self.content, re.I
        )
        if sec_m:
            self.info.oob_ip = sec_m.group(1)
            return

        # ── 策略 B ──────────────────────────────────────────────────────
        gw_m = re.search(r"ip default-gateway\s+(\d+\.\d+\.\d+)\.\d+", self.content)
        if gw_m:
            prefix = gw_m.group(1)   # 例："192.0.2"
            vlan_m = re.search(
                rf"interface\s+Vlan\d+\s+(?:[^\n]*\n)*?\s*ip address\s+({re.escape(prefix)}\.\d+)",
                self.content, re.S | re.I
            )
            if vlan_m:
                self.info.oob_ip = vlan_m.group(1)
                return

        # ── 策略 C ──────────────────────────────────────────────────────
        brief_m = re.search(
            r"show ip interface brief\b.*?\n(.*?)(?=\n\S+#|\Z)",
            self.content, re.S | re.I
        )
        if brief_m:
            for line in brief_m.group(1).splitlines():
                if not re.search(r"Vlan", line, re.I):
                    continue
                ip_m = re.search(r"(\d+\.\d+\.\d+\.\d+)", line)
                if ip_m and "unassigned" not in line and ip_m.group(1) != "0.0.0.0":
                    self.info.oob_ip = ip_m.group(1)
                    return

        # ── 策略 D ──────────────────────────────────────────────────────
        vlan_m = re.search(
            r"interface\s+Vlan\d+.*?ip address\s+(\d+\.\d+\.\d+\.\d+)",
            self.content, re.S | re.I
        )
        if vlan_m:
            self.info.oob_ip = vlan_m.group(1)

    def _parse_syslog(self):
        """
        收集 Log 中所有嚴重等級 0~4 的 Syslog，過濾雜訊後保留最後 50 筆。

        處理流程：
          1. findall()：找出所有符合 %XXX-Y-ZZZ: 格式的行
          2. 過濾：排除 EXCLUDE_LOG_PREFIXES 中的雜訊前綴
          3. dict.fromkeys()：有序去重（Python 3.7+ dict 保證插入順序）
          4. [-50:]：只保留最後 50 筆（避免 Excel 儲存格內容過長）
          5. join：合併為多行字串
        """
        logs = SYSLOG_REGEX.findall(self.content)
        if not logs:
            return

        final = [
            l.strip() for l in logs
            if not any(l.strip().startswith(p) for p in EXCLUDE_LOG_PREFIXES)
        ]
        self.syslog = "\n".join(list(dict.fromkeys(final))[-200:])

    def generate_rows(self) -> List[Dict]:
        """
        將解析結果轉換為 Excel 報表的資料列清單。

        Stack 設備的輸出格式：
          - 第一列（編號最小的成員）：填入完整設備資訊
          - 後續列（其他 Stack 成員）：
            Hostname / OOB IP / CPU / Memory / Syslog 填「○ (Stack)」或「○」，
            表示這些欄位與第一列共用，不重複顯示。

        為什麼用「○」？
          清楚標示「此欄位由 Stack 主機提供」，
          且 IP 排序邏輯可識別「○」為無效 IP，使 Stack 列緊跟主列不被拆散。

        Returns:
            List[Dict]：每個 Dict 的 key 對應 Excel 欄位名稱
        """
        rows = []
        self.members.sort(key=lambda x: x.member_id)   # 依 Stack 編號排序

        for idx, m in enumerate(self.members):
            is_first = (idx == 0)
            # 優先用各成員自身的 MAC（Stack 情境下每台有獨立 MAC）；
            # 若成員 MAC 未解析到，is_first 時 fallback 到設備整體 MAC，其餘填 ○
            member_mac = m.eth_mac if m.eth_mac != "N/A" else (
                self.info.eth_mac if is_first else "○"
            )
            rows.append({
                "Hostname":         self.info.hostname if is_first else "○ (Stack)",
                "Model":            m.model,
                "OOB IP":           self.info.oob_ip   if is_first else "○",
                "Ethernet MAC":     member_mac,
                "Serial Number":    m.serial,
                "Firmware Version": self.info.firmware,
                "CPU Usage":        self.info.cpu       if is_first else "○",
                "Memory Usage":     self.info.memory    if is_first else "○",
                "Fan Status":       self.fan_status,
                "Power Status":     self.pwr_status,
                "System Log":       self.syslog         if is_first else "",
                "_device_type":     self.info.device_type,   # 私有欄位，用於分頁過濾，不輸出到 Report
            })
        return rows


# ==============================================================================
# 5. WLC 解析器（Cisco C9800 系列）
# ==============================================================================

class WLCParser(BaseParser):
    """
    針對 Cisco C9800 WLC 的專屬解析器。

    C9800 執行 IOS-XE，與一般 Switch 的主要差異：
      - 韌體版本格式：「Cisco IOS XE Software, Version 17.x.x」
      - 記憶體指令：show process memory sorted
      - 環境輸出：有「Fan Speed 50%」特殊狀態（節電模式，非異常）
      - 有 show ap summary / show ap config general 指令
    """

    def _parse_firmware(self):
        """
        從 show version 擷取 IOS-XE 版本號並設定設備類型為 WLC。

        目標行：「Cisco IOS XE Software, Version 17.09.08」
        [^\s,]+ → 非空白且非逗號的字元序列，確保只取純版本號。
        """
        m = re.search(r"Cisco IOS XE Software, Version\s+([^\s,]+)", self.ver_sec, re.I)
        self.info.firmware = m.group(1).strip() if m else "Unknown"
        self.info.device_type = "WLC"

    def _parse_cpu(self):
        """
        從 show processes cpu sorted 取 CPU 五分鐘平均使用率。
        格式：「CPU utilization for five seconds: 1%/0%; one minute: 2%; five minutes: 2%」
        """
        sec = self._get_command_block("show processes cpu sorted")
        m = re.search(r"five minutes:\s*(\d+)%", sec)
        self.info.cpu = f"{m.group(1)}%" if m else "N/A"

    def _parse_memory(self):
        """
        從 show process memory sorted 擷取 Processor Pool 記憶體資訊。
        格式：「Processor Pool Total: 3770676924 Used:  499321360 Free: 3271355564」
        usage_pct = Used / Total × 100，round 保留一位小數。
        """
        sec = self._get_command_block("show process memory sorted")
        m = re.search(
            r"Processor Pool Total:\s*(\d+).*?Used:\s*(\d+).*?Free:\s*(\d+)",
            sec, re.S | re.I
        )
        if m:
            total   = int(m.group(1))
            used    = int(m.group(2))
            free    = int(m.group(3))
            percent = round((used / total) * 100, 1) if total > 0 else 0
            self.info.memory = self.format_mem(m.group(1), m.group(2), m.group(3), percent)

    def _parse_inventory(self):
        """
        從 show inventory 擷取 WLC 機箱型號（PID）與序號（SN）。

        C9800 HA 雙機架構下，inventory 會有 "Chassis 1" 和 "Chassis 2"，
        Regex 用 Chassis\s*\d* 同時匹配以下格式：
          NAME: "Chassis"   → 單機
          NAME: "Chassis 1" → HA 主機（取第一個）
          NAME: "Chassis 2" → HA 備機（被忽略，因為只取第一個匹配）

        strip(',') 去除 inventory 行尾可能的殘留逗號。
        WLC 不使用 Stack，unit_count 固定為 1。
        """
        block = re.search(
            r'NAME:\s*"Chassis\s*\d*".*?PID:\s*([^\s,"]+).*?SN:\s*([^\s,"]+)',
            self.inv_sec, re.S | re.I
        )
        if block:
            self.members.append(StackMember(1, block.group(1).strip(','), block.group(2).strip(',')))
        else:
            self.members.append(StackMember(1, "Unknown-C9800", "Unknown"))

        self.info.unit_count = 1

    def _parse_env(self):
        """
        解析 show environment all，判斷風扇與電源狀態。

        WLC（C9800）特殊情況：
          「R0  Temp: FC FANS  Fan Speed 50%  23  Celsius」
          此行的 Current State 欄是「Fan Speed 50%」而非「Normal」，
          但這是 WLC 在正常溫度下的節電模式（降低風扇轉速），並非異常。

        逐行分析邏輯：
          條件一：行不含 "Normal"（正常狀態標記）
          條件二：行不含 "Fan Speed 50%"（WLC 節電特例）
          同時滿足 → 再判斷 Fan / Power 字樣 → 設定異常旗標
        """
        sec   = self._get_command_block("show environment all")
        lines = sec.splitlines()
        f_bad, p_bad = False, False

        for line in lines:
            if not line.strip():
                continue
            if not ("Normal" in line or "Fan Speed 50%" in line):
                if "Fan"   in line: f_bad = True
                if "Power" in line or "PSU" in line: p_bad = True

        self.fan_status = "異常" if f_bad else "正常"
        self.pwr_status = "異常" if p_bad else "正常"

    # ──────────────────────────────────────────────────────────────────────────
    # AP 解析方法（WLC 專屬）
    # ──────────────────────────────────────────────────────────────────────────

    def parse_ap_summary(self) -> List[Dict]:
        """
        從 show ap summary 解析所有 AP 的基本資訊（Name、Model、MAC、IP）。

        表格格式（欄位以固定字元位置對齊）：
          AP Name          Slots  AP Model  Ethernet MAC       Radio MAC     IP Address ...
          -------          -----  --------  -----------------  ----------    ----------
          AP-7F-01         2      CW9166I   aa:bb:cc:00:00:03  aa:bb:...     203.0.113.101

        AP Name 可能含有空白（如 "AP Room 101"），無法用 split() 直接切割，
        因此採用「字元偏移（Column Offset）」策略：
          Step 1：找到表頭行，用 re.search().start() 記錄各欄位在該行的字元起始位置
          Step 2：對每個資料行，按照相同的字元位置範圍切出各欄內容
          → AP Name 以「行首到 Slots 欄起始位置」為範圍，可正確處理含空白的名稱

        Returns:
            List[Dict]：每個 Dict 含 ap_name / ap_model / eth_mac / ip_address
        """
        sec = self._get_command_block("show ap summary")
        if not sec:
            return []

        results = []
        lines   = sec.splitlines()

        # ── Step 1：找表頭行，記錄各欄的字元起始位置 ──────────────────────
        header_idx = -1
        col_slots = col_model = col_eth = col_ip = -1

        for i, line in enumerate(lines):
            if re.search(r"AP\s+Name\s+.*AP\s+Model\s+.*Ethernet\s+MAC", line, re.I):
                header_idx = i
                slots_m   = re.search(r"Slots",         line, re.I)
                model_m   = re.search(r"AP\s+Model",    line, re.I)
                eth_m     = re.search(r"Ethernet\s+MAC", line, re.I)
                ip_m      = re.search(r"IP\s+Address",  line, re.I)
                col_slots = slots_m.start() if slots_m else -1
                col_model = model_m.start() if model_m else -1
                col_eth   = eth_m.start()   if eth_m   else -1
                col_ip    = ip_m.start()    if ip_m    else -1
                break

        if header_idx == -1 or col_model == -1:
            # 表頭解析失敗，改用 Regex 逐行備案擷取
            for line in lines:
                # 【修改】將 [\da-f:]{17} 替換為 [\da-fA-F:.-]+ 支援小數點格式
                m = re.match(
                    r"^(\S[\S ]*?)\s{2,}(\d+)\s+(\S+)\s+([\da-fA-F:.-]+)\s+([\da-fA-F:.-]+)\s+(\d+\.\d+\.\d+\.\d+)",
                    line, re.I
                )
                if m:
                    results.append({
                        "ap_name":    m.group(1).strip(),
                        "ap_model":   m.group(3).strip(),
                        "eth_mac":    m.group(4).strip(),
                        "ip_address": m.group(6).strip(),
                    })
            return results

        # ── Step 2：逐行依字元位置切割 ─────────────────────────────────────
        for line in lines[header_idx + 1:]:
            if not line.strip() or line.strip().startswith("-"):
                continue

            # AP Name：行首到 Slots 欄（或 Model 欄）前
            ap_name = line[:col_slots].strip() if col_slots > 0 else line[:col_model].strip()
            if not ap_name:
                continue

            # AP Model：col_model 到 col_eth 之間的第一個 token
            seg_model = line[col_model:col_eth].strip() if col_eth > col_model else line[col_model:].strip()
            ap_model  = seg_model.split()[0] if seg_model.split() else "Unknown"

            # 【修改】Ethernet MAC：支援傳統 xx:xx:xx 與 IOS-XE xxxx.xxxx.xxxx 兩種格式
            seg_eth       = line[col_eth:col_ip].strip() if col_ip > col_eth else line[col_eth:].strip()
            eth_mac_match = re.search(r"(?:[\da-fA-F]{2}[:-]){5}[\da-fA-F]{2}|[\da-fA-F]{4}\.[\da-fA-F]{4}\.[\da-fA-F]{4}", seg_eth)
            eth_mac       = eth_mac_match.group(0) if eth_mac_match else "N/A"

            # IP Address：col_ip 之後，取第一個符合 IP 格式的字串
            seg_ip    = line[col_ip:].strip() if col_ip > 0 else ""
            ip_match  = re.search(r"\d+\.\d+\.\d+\.\d+", seg_ip)
            ip_address = ip_match.group(0) if ip_match else "N/A"

            results.append({
                "ap_name":    ap_name,
                "ap_model":   ap_model if ap_model else "Unknown",
                "eth_mac":    eth_mac,
                "ip_address": ip_address,
            })

        return results

    def parse_ap_config(self) -> Dict[str, Dict]:
        """
        從 show ap config general 解析每台 AP 的 Serial Number 和 Software Version。

        格式特徵：
          每台 AP 以「Cisco AP Name   : <名稱>」開頭，後續行包含各種參數。
          Cisco AP Name   : AP-7F-01
          ...
          Software Version                  : 17.9.8.5
          AP Serial Number                  : SN000000001

        解析策略：
          以「Cisco AP Name   :」為分割點，用 re.split 切出各 AP 的子區塊，
          再在每個子區塊中分別尋找 Software Version 和 AP Serial Number。

        Software Version 過濾規則（r"17\.\S+"）：
          只接受以 "17." 開頭的版本號，其他（如 8.10.x 舊版、空白）填 "N/A"。

        V14 新增：
          呼叫 truncate_ap_firmware() 將版本號截為前三段（17.9.8.5 → 17.9.8）。

        Returns:
            Dict[str, Dict]：key = AP Name，value = {"serial": ..., "firmware": ...}
        """
        sec = self._get_command_block("show ap config general")
        if not sec:
            return {}

        results = {}

        # (?=...) 正向前視：保留分割點本身，使每個子區塊的第一行仍是 AP Name 行
        blocks = re.split(r"(?=Cisco AP Name\s*:)", sec)

        for block in blocks:
            if not block.strip():
                continue

            name_m = re.match(r"Cisco AP Name\s*:\s*(.+)", block)
            if not name_m:
                continue
            ap_name = name_m.group(1).strip()

            # 擷取 Software Version（只接受 17.x 開頭）
            fw_m   = re.search(r"Software Version\s*:\s*(17\.\S+)", block)
            raw_fw = fw_m.group(1).strip() if fw_m else "N/A"

            # ── V14 新增：截取前三段版本號（17.9.8.5 → 17.9.8）─────────────
            firmware = truncate_ap_firmware(raw_fw)

            sn_m   = re.search(r"AP Serial Number\s*:\s*(\S+)", block)
            serial = sn_m.group(1).strip() if sn_m else "N/A"

            results[ap_name] = {"serial": serial, "firmware": firmware}

        return results

    def get_ap_rows(self) -> List[Dict]:
        """
        整合 parse_ap_summary() 和 parse_ap_config() 的結果，
        以 AP Name 為 key 進行 JOIN，產生每台 AP 的完整資料列。

        為什麼分兩個指令？
          show ap summary        → 提供概覽表格（Name / Model / MAC / IP）
          show ap config general → 提供詳細設定（Serial / Software Version）
          兩者用 AP Name 作為識別鍵，透過 dict.get() 對應合併。

        Returns:
            List[Dict]：含 ap_name / ap_model / eth_mac / ip_address / serial / firmware
        """
        summary_list = self.parse_ap_summary()
        config_dict  = self.parse_ap_config()

        rows = []
        for ap in summary_list:
            name   = ap.get("ap_name", "Unknown")
            config = config_dict.get(name, {})
            rows.append({
                "ap_name":    name,
                "ap_model":   ap.get("ap_model",   "Unknown"),
                "eth_mac":    ap.get("eth_mac",     "N/A"),
                "ip_address": ap.get("ip_address",  "N/A"),
                "serial":     config.get("serial",   "N/A"),
                "firmware":   config.get("firmware", "N/A"),
            })
        return rows


# ==============================================================================
# 6. IOS Switch 解析器（Catalyst 系列，支援單機與 Stack 堆疊）
# ==============================================================================

class IOSSwitchParser(BaseParser):
    """
    針對 Cisco Catalyst IOS / IOS-XE Switch 的專屬解析器。

    支援型號：C2960 / C1000 / C3850 / C9200 / C9300 / C9500（含 Stack 堆疊）

    ╔══════════════════════════════════════════════════════════════════╗
    ║              Stack 判斷機制完整說明                               ║
    ╠══════════════════════════════════════════════════════════════════╣
    ║                                                                  ║
    ║  【什麼是 Stack？】                                               ║
    ║                                                                  ║
    ║    Stack 是 Cisco Catalyst Switch 的硬體架構，                   ║
    ║    多台實體 Switch 以 StackWise 纜線串接成「環形」，              ║
    ║    對管理者和網路設備來說，它們表現得像一台單一設備：             ║
    ║      ✅ 共用同一個 IP 位址和 hostname                            ║
    ║      ✅ 共用同一個 CLI（一個 SSH 連線即可管理所有成員）           ║
    ║      ✅ 共用同一份 running-config                                ║
    ║      ✅ 共用同一個韌體版本（強制所有成員升級到相同版本）          ║
    ║      ❌ 各自有獨立的型號（可以混用不同型號，如 C3850-48P + 12S） ║
    ║      ❌ 各自有獨立的序號（資產管理上要分開列出）                  ║
    ║                                                                  ║
    ║  【判斷依據：show version 的 Switch 成員表格】                   ║
    ║                                                                  ║
    ║    執行 show version 時，IOS 會在輸出末尾顯示一個表格，           ║
    ║    列出目前 Stack 中的所有成員：                                  ║
    ║                                                                  ║
    ║    ▶ 單機（1 台）範例：                                          ║
    ║      Switch Ports Model              SW Version        SW Image  ║
    ║      ------ ----- -----              ----------        --------- ║
    ║      *    1 28    WS-C2960X-24PS-L   15.0(2)EX4   C2960X-UNIV   ║
    ║      ↑     ↑  ↑   ↑                  ↑                           ║
    ║      *=    sid Port Model             Version                    ║
    ║      master                                                       ║
    ║                                                                  ║
    ║    ▶ Stack（4 台，C3850 混搭）範例：                             ║
    ║      Switch Ports Model              SW Version      Mode        ║
    ║      ------ ----- -----              ----------      ----        ║
    ║      *    1 20    WS-C3850-12S       16.12.09        INSTALL     ║
    ║           2 20    WS-C3850-12S       16.12.09        INSTALL     ║
    ║           3 56    WS-C3850-48T       16.12.09        INSTALL     ║
    ║           4 56    WS-C3850-48T       16.12.09        INSTALL     ║
    ║                                                                  ║
    ║    表格欄位說明：                                                 ║
    ║      *   → Stack Master 標記（每個 Stack 只有一台 Master）        ║
    ║      sid → Switch 編號（1~8，由纜線拓撲位置決定）                ║
    ║      28  → Port 數量（不使用，用 \d+ 略過）                      ║
    ║      Model → 設備型號（需擷取）                                  ║
    ║      SW Version → 韌體版本（需擷取）                             ║
    ║      Mode → BUNDLE/INSTALL（C3850 特有欄位，直接忽略）           ║
    ║                                                                  ║
    ║  【解析 Regex 逐字元解析】                                       ║
    ║                                                                  ║
    ║  pattern = re.compile(r"^\*?\s+(\d+)\s+\d+\s+(\S+)\s+(\S+)", re.M)
    ║                                                                  ║
    ║    ^       → 行首（re.M 使 ^ 匹配每一行的開頭，而非字串開頭）   ║
    ║    \*?     → 可選的 * 號（有 * 是 Master，沒有是普通成員）       ║
    ║    \s+     → 一或多個空格（Switch 欄和編號之間的對齊空格）       ║
    ║    (\d+)   → 捕捉 Switch 編號（sid），如 "1"、"2"、"3"           ║
    ║    \s+     → 空格                                                ║
    ║    \d+     → Port 數量（只匹配，不捕捉，直接略過）               ║
    ║    \s+     → 空格                                                ║
    ║    (\S+)   → 捕捉型號（非空白字元序列），如 "WS-C3850-12S"       ║
    ║    \s+     → 空格                                                ║
    ║    (\S+)   → 捕捉 SW Version，如 "16.12.09"、"15.0(2)EX4"       ║
    ║                                                                  ║
    ║  ⚠️  C3850 的 Mode 欄（BUNDLE / INSTALL）位於第三個 \S+ 之後，  ║
    ║     我們只捕捉前兩個 \S+（型號和版本），Mode 欄自動被忽略。      ║
    ║                                                                  ║
    ║  【Stack 判斷邏輯】                                              ║
    ║                                                                  ║
    ║    table_pattern.findall(ver_sec) 回傳所有匹配的 tuple list：   ║
    ║                                                                  ║
    ║    單機結果：                                                     ║
    ║      [("1", "WS-C2960X-24PS-L", "15.0(2)EX4")]                 ║
    ║      → len = 1 → is_stack = False                               ║
    ║                                                                  ║
    ║    Stack 結果：                                                  ║
    ║      [("1", "WS-C3850-12S", "16.12.09"),                       ║
    ║       ("2", "WS-C3850-12S", "16.12.09"),                       ║
    ║       ("3", "WS-C3850-48T", "16.12.09"),                       ║
    ║       ("4", "WS-C3850-48T", "16.12.09")]                       ║
    ║      → len = 4 → is_stack = True，unit_count = 4               ║
    ║                                                                  ║
    ║  【為什麼不用 "show switch" 指令？】                             ║
    ║                                                                  ║
    ║    "show switch" 只有 C9K（IOS-XE 17.x 系列）才支援：           ║
    ║      ✅ C9200 / C9300 / C9500 → 有 show switch 指令             ║
    ║      ❌ C2960 / C1000 / C3850 → 執行會出現 ^ Invalid input 錯誤 ║
    ║                                                                  ║
    ║    "show version 的 Switch 表格"是所有 IOS / IOS-XE 平台        ║
    ║    都有的輸出，可跨平台統一使用，不需要額外的指令。               ║
    ║                                                                  ║
    ╚══════════════════════════════════════════════════════════════════╝
    """

    def _parse_firmware(self):
        """
        IOS Switch 的韌體版本由 _parse_inventory() 從 show version 表格一起取得，
        此方法只負責設定設備類型標籤為 "Switch"。
        """
        self.info.device_type = "Switch"

    def _parse_cpu(self):
        """
        從 Log 全文搜尋 CPU 五分鐘平均使用率。

        格式：「CPU utilization for five seconds: 34%/2%; one minute: 12%; five minutes: 9%」
        取 five minutes 值（比 five seconds 更能反映穩態負載水位）。
        """
        m = re.search(r"five minutes:\s*(\d+)%", self.content)
        self.info.cpu = f"{m.group(1)}%" if m else "N/A"

    def _parse_memory(self):
        """
        從 Log 全文搜尋 Processor Pool 記憶體行。

        格式：「Processor Pool Total:  530761472 Used:  118563156 Free:  412198316」
        只抓 Total 和 Used，Free 以 Total - Used 計算，
        兼容部分 IOS 版本 Free 欄位未出現在過濾輸出中的情形。
        """
        m = re.search(r"Processor Pool Total:\s*(\d+).*?Used:\s*(\d+)", self.content, re.S)
        if m:
            t_val = int(m.group(1))
            u_val = int(m.group(2))
            f_val = t_val - u_val
            pct   = round((u_val / t_val) * 100, 1) if t_val > 0 else 0
            self.info.memory = self.format_mem(t_val, u_val, f_val, pct)

    def _parse_inventory(self):
        """
        解析硬體庫存，同時判斷 Stack 架構並取得每台成員的型號、序號、韌體版本。

        ── Step 1：解析 show version Switch 表格 ──────────────────────────

          從 self.ver_sec（預先擷取的 show version 輸出）中，
          用 findall() 取出所有成員的 (sid, model, version) tuple。

          is_stack 判斷：len(matches) > 1
            - True  → Stack（多台實體 Switch 組成邏輯設備）
            - False → 單機

        ── Step 2：從 show inventory 取各成員的精確序號（SN）──────────────

          show inventory 有兩種常見格式：

          格式 A（C3850 / C9K，NAME 含 "Switch" 前綴）：
            NAME: "Switch 1", DESCR: "WS-C3850-48P-S"
            PID: WS-C3850-48P-S    , VID: V07  , SN: SN000000002

          格式 B（C2960 / C1000，NAME 只有數字）：
            NAME: "1", DESCR: "WS-C2960X-24PS-L"
            PID: WS-C2960X-24PS-L  , VID: V01  , SN: SN000000003

          Regex 的 (?:Switch\\s*)? 為可選的非捕捉群組，同時匹配兩種格式：
            - "Switch 1" → (?:Switch\s*)? 匹配 "Switch "，\d+ 匹配 "1"
            - "1"        → (?:Switch\s*)? 匹配零次，\d+ 匹配 "1"

          {sid} 為 Python f-string（rf'...'）動態插入 Switch 編號，
          rf 是 raw f-string：r 防止反斜線被解釋為跳脫字元，f 允許 {} 插值。
        """
        # ── Step 1：解析 show version Switch 表格 ──────────────────────────
        # re.M 使 ^ 能匹配每一行的行首（而非只匹配整個字串的開頭）
        table_pattern = re.compile(r"^\*?\s+(\d+)\s+\d+\s+(\S+)\s+(\S+)", re.M)
        matches = table_pattern.findall(self.ver_sec)   # [(sid, model, version), ...]

        if matches:
            # ── Stack 判斷：成員數 > 1 即為 Stack ──────────────────────────
            self.info.is_stack   = len(matches) > 1    # True=Stack，False=單機
            self.info.unit_count = len(matches)         # 成員台數

            for sid, model, version in matches:
                # 所有成員共用同一 SW Version，每次迴圈都寫入，最後以表格最後一筆為準
                self.info.firmware = version

                # ── Step 2：從 show inventory 取此成員的序號 ───────────────
                # group(1) = PID 型號，group(2) = SN 序號
                inv_match = re.search(
                    rf'NAME:\s*"(?:Switch\s*)?{sid}".*?PID:\s*(\S+).*?SN:\s*(\S+)',
                    self.inv_sec, re.S | re.I
                )
                sn = inv_match.group(2).strip(',') if inv_match else "Unknown"
                self.members.append(StackMember(int(sid), model, sn))
        else:
            # 表格解析失敗（備案）
            self.members.append(StackMember(1, "Unknown", "Unknown"))
            self.info.unit_count = 1

    def _parse_env(self):
        """
        解析 show environment all，判斷風扇與電源狀態。

        各 IOS 平台輸出格式差異：
          ┌───────────────┬────────────────────────────────────────────────┐
          │ 平台          │ 環境輸出格式                                    │
          ├───────────────┼────────────────────────────────────────────────┤
          │ C2960 / C2960X│ 「FAN is OK」                                  │
          │               │ 「SYSTEM TEMPERATURE is OK」                   │
          │               │ 「1  Built-in  Good」（電源 Good 欄）          │
          ├───────────────┼────────────────────────────────────────────────┤
          │ C1000         │ 「SYSTEM TEMPERATURE is OK」（無 FAN 行！）    │
          │               │ 「Power Supply Status: Good」                  │
          ├───────────────┼────────────────────────────────────────────────┤
          │ C3850 Stack   │ 「Switch 1 FAN 1 is OK」（每台成員各自列出）  │
          │               │ 「FAN PS-1 is OK」（電源風扇）                 │
          │               │ 「1A  PWR-C1-350WAC  SN000000004  OK  Good」  │
          └───────────────┴────────────────────────────────────────────────┘

        判斷邏輯：
          ① has_fan_info = \bFAN\b 是否出現
             → C1000 無 FAN 行，has_fan_info=False，fan_bad 永遠 False（視為正常）
             → C3850 有「Switch X FAN Y is OK」，has_fan_info=True，才進行異常判斷
          ② has_pwr_info = \b(Power|PSU|PWR)\b 是否出現
          ③ fan_bad / pwr_bad = ENV_BAD_REGEX 命中 AND has_xxx_info
          ④ 若偵測到「Power Supply Status: Good」→ 強制 pwr_bad=False
             （C1000 電源格式不含 Pwr 欄，只用此聲明行判斷）
        """
        sec = self._get_command_block("show environment all")
        if not sec:
            sec = self._get_command_block("show environment")

        has_fan_info = bool(re.search(r"\bFAN\b",             sec, re.I))
        has_pwr_info = bool(re.search(r"\b(?:Power|PSU|PWR)\b", sec, re.I))

        fan_bad = bool(ENV_BAD_REGEX.search(sec)) and has_fan_info
        pwr_bad = bool(ENV_BAD_REGEX.search(sec)) and has_pwr_info

        if re.search(r"Power Supply Status:\s*Good", sec, re.I):
            pwr_bad = False   # C1000 特例：明確聲明 Good，覆蓋 pwr_bad

        self.fan_status = "異常" if fan_bad else "正常"
        self.pwr_status = "異常" if pwr_bad else "正常"

    def _parse_eth_mac(self):
        """
        IOS Switch（含 Stack）每台成員的 Ethernet MAC 解析。

        解析策略：
          ── 單機 ────────────────────────────────────────────────────────────
          直接從 show version 取「Base Ethernet MAC Address : xx:xx:xx:xx:xx:xx」，
          寫入 self.info.eth_mac 和 members[0].eth_mac。

          ── Stack ───────────────────────────────────────────────────────────
          show version 對 Stack 的格式為：
            主機（無標題，或「Switch 01」前的部分）：
              Base Ethernet MAC Address          : aa:bb:cc:00:00:04
              System Serial Number               : SN000000005

            後續成員各自有獨立區塊，以「Switch 0N」或「Switch N」為分隔：
              Switch 02
              ---------
              Base Ethernet MAC Address          : aa:bb:cc:00:00:05
              System Serial Number               : SN000000006

          比對邏輯：
            1. 先解析全文，建立 serial → MAC 的映射字典（sn_mac_map）
               同時把「Switch 0N」區塊也納入，以確保取得各成員的 MAC。
            2. 以 StackMember.serial 為 key 查 sn_mac_map，找到就寫入 member.eth_mac。
            3. members[0].eth_mac 同步寫回 self.info.eth_mac（整機代表 MAC）。

          為什麼需要用序號比對而非直接用順序？
            show version 的 Switch 區塊順序不一定與 members 編號一致，
            且主機區塊（Stack Master）的 MAC 可能出現在清單頂部，
            用序號比對可確保各成員 MAC 正確對應。
        """
        MAC_RE = re.compile(
            r"Base\s+[Ee]thernet\s+MAC\s+[Aa]ddress\s*[:\s]+"
            r"([0-9a-fA-F]{2}[:\-][0-9a-fA-F]{2}[:\-][0-9a-fA-F]{2}"
            r"[:\-][0-9a-fA-F]{2}[:\-][0-9a-fA-F]{2}[:\-][0-9a-fA-F]{2}"
            r"|[0-9a-fA-F]{4}\.[0-9a-fA-F]{4}\.[0-9a-fA-F]{4})"
        )
        SN_RE  = re.compile(r"System\s+Serial\s+Number\s*[:\s]+(\S+)")

        # ── Step 1：切割 show version 為各成員區塊 ───────────────────────
        # 以「Switch 01 / Switch 1 / Switch  2」等行作為分割點
        blocks = re.split(r"\n(?=Switch\s+0*\d+\s*\n[-─]+)", self.ver_sec, flags=re.I)

        sn_mac_map: Dict[str, str] = {}   # serial_number → eth_mac

        for block in blocks:
            mac_m = MAC_RE.search(block)
            sn_m  = SN_RE.search(block)
            if mac_m and sn_m:
                sn  = sn_m.group(1).strip().rstrip(',')
                mac = mac_m.group(1).strip()
                sn_mac_map[sn] = mac

        # ── Step 2：如果 sn_mac_map 為空（單機或 show version 格式無區塊）─
        # 直接搜整個 ver_sec 取第一個 Base MAC
        if not sn_mac_map:
            m = MAC_RE.search(self.ver_sec)
            if m:
                self.info.eth_mac = m.group(1).strip()
                if self.members:
                    self.members[0].eth_mac = self.info.eth_mac
            return

        # ── Step 3：以序號比對，寫入各 StackMember.eth_mac ──────────────
        for member in self.members:
            sn_key = member.serial.strip().rstrip(',')
            if sn_key in sn_mac_map:
                member.eth_mac = sn_mac_map[sn_key]

        # ── Step 4：整機代表 MAC = 第一個成員（Stack Master）的 MAC ──────
        if self.members:
            self.members.sort(key=lambda x: x.member_id)
            first_mac = self.members[0].eth_mac
            if first_mac != "N/A":
                self.info.eth_mac = first_mac
            elif sn_mac_map:
                # fallback：取 sn_mac_map 第一個值
                self.info.eth_mac = next(iter(sn_mac_map.values()))


# ==============================================================================
# 7. NX-OS 解析器（Cisco Nexus 系列）
# ==============================================================================

class NexusParser(BaseParser):
    """
    針對 Cisco Nexus NX-OS 設備的專屬解析器。

    NX-OS 與 IOS/IOS-XE 主要差異：
      - 版本格式  ：「NXOS: version 10.3(6)」
      - 記憶體指令：show system resources（而非 show processes memory）
      - 環境指令  ：show environment（無 "all" 後綴）
      - 架構      ：固定單機（無 Stack 功能）
    """

    def _parse_firmware(self):
        """
        從 show version 擷取 NX-OS 版本號並設定設備類型。

        NX-OS 版本格式因平台而異：
          N9K / N3K（新型）：「NXOS: version 10.3(6)」
          N5K / N6K（舊型）：「system: version 7.3(2)N1(1)」
                              （無 "NXOS:" 前綴，使用 "system:" 格式）

        解析策略：
          優先嘗試 "NXOS: version"（N9K 格式），
          若找不到再嘗試 "system: version"（N5K/N6K 格式）。
          兩者都找不到才回傳 "Unknown"。
        """
        # N9K / N3K 格式（有 NXOS: 前綴）
        m = re.search(r"NXOS:\s+version\s+([^\s\[]+)", self.ver_sec, re.I)
        if not m:
            # N5K / N6K 格式（system: version X，無 NXOS: 前綴）
            m = re.search(r"system:\s+version\s+(\S+)", self.ver_sec, re.I)
        self.info.firmware = m.group(1).strip() if m else "Unknown"
        self.info.device_type = "NX-OS"

    def _parse_cpu(self):
        """
        擷取 NX-OS CPU 使用率，相容 N9K 和 N5K 兩種輸出格式。

        N9K 格式（show system resources）：
          「CPU utilization for five seconds: 38%/0%; one minute: 17%; five minutes: 14%」
          → 直接取 "five minutes: X%"

        N5K 格式（show system resources）：
          「CPU states  :   1.00% user,   1.25% kernel,   97.75% idle」
          → N5K 無 "five minutes:" 行，改用 user + kernel 加總
          → 代表 CPU 目前的「非 idle」使用率，與 five minutes 概念相近

        解析策略：
          優先嘗試 N9K 的 "five minutes: X%" 格式，
          若找不到再嘗試 N5K 的 "CPU states: X% user, X% kernel" 格式。
        """
        # N9K 格式
        m = re.search(r"five minutes:\s*(\d+)%", self.content)
        if m:
            self.info.cpu = f"{m.group(1)}%"
            return
        # N5K 格式：user% + kernel% = 總 CPU 使用率（idle% = 100 - 結果）
        m2 = re.search(
            r"CPU states\s*:\s*([\d.]+)%\s*user,\s*([\d.]+)%\s*kernel",
            self.content
        )
        if m2:
            total = round(float(m2.group(1)) + float(m2.group(2)), 1)
            self.info.cpu = f"{total}%"

    def _parse_memory(self):
        """
        從 show system resources 擷取記憶體資訊（NX-OS 使用 KB 為單位）。

        格式：「Memory usage:   32802764K total,   13071592K used,   19731172K free」
        保留 K 後綴在輸出字串中，與 IOS 的 Bytes 格式有所區分。
        """
        m = re.search(r"Memory usage:\s+(\d+)K total,\s+(\d+)K used", self.content)
        if m:
            t = int(m.group(1))
            u = int(m.group(2))
            f = t - u
            p = round((u / t) * 100, 1) if t > 0 else 0
            self.info.memory = self.format_mem(f"{t}K", f"{u}K", f"{f}K", p)

    def _parse_inventory(self):
        """
        從 show inventory 擷取 Nexus 機箱型號與序號。

        NX-OS inventory 格式：
          NAME: "Chassis",  DESCR: "Nexus9000 C93180YC-FX3 Chassis"
          PID: N9K-C93180YC-FX3    ,  VID: V01 ,  SN: SN000000007

        Nexus 不支援 Stack，unit_count 固定為 1。
        """
        m = re.search(
            r'NAME:\s*"Chassis".*?PID:\s*(\S+).*?SN:\s*(\S+)',
            self.inv_sec, re.S | re.I
        )
        if m:
            self.members.append(StackMember(1, m.group(1).strip(','), m.group(2).strip(',')))
        else:
            self.members.append(StackMember(1, "N9K-Chassis", "Unknown"))

        self.info.unit_count = 1

    def _parse_env(self):
        """
        解析 NX-OS 的 show environment 輸出（不加 "all" 後綴）。
        使用全域 ENV_BAD_REGEX 偵測風扇（Fan）或電源（Power）異常關鍵字。
        """
        sec = self._get_command_block("show environment")
        self.fan_status = "異常" if ENV_BAD_REGEX.search(sec) and "Fan"   in sec else "正常"
        self.pwr_status = "異常" if ENV_BAD_REGEX.search(sec) and "Power" in sec else "正常"

    def _parse_eth_mac(self):
        """
        NX-OS MAC 解析（兩段策略）：

        優先 ── vPC local system-mac（show vpc role）：
          vPC 環境下，NX-OS 使用虛擬 MAC 作為 vPC peer 的共同 system-mac，
          此 MAC 才是流量轉發的實際 MAC，比 show version 的 Base MAC 更有代表性。

          show vpc role 格式：
            vPC Role status
            -------------------------------------------------
            Operational role                    : primary
            vPC local system-mac                : aa:bb:cc:00:00:06
            vPC system-priority                 : 32667
            vPC local operational system-mac    : aa:bb:cc:00:00:06

        備案 ── show interface mgmt0（父類別已實作）：
          若 show vpc role 不存在（非 vPC 環境），
          呼叫 super()._parse_eth_mac() 嘗試從 mgmt0 介面取 MAC。
        """
        vpc_sec = self._get_command_block("show vpc role")
        if vpc_sec:
            m = re.search(
                r"vPC\s+local\s+system-mac\s*[:\s]+([0-9a-fA-F:]{17})",
                vpc_sec, re.I
            )
            if m:
                self.info.eth_mac = m.group(1).strip()
                return

        # 備案：從 mgmt0 介面取 MAC（父類別邏輯）
        super()._parse_eth_mac()


# ==============================================================================
# 8. AireOS WLC 解析器（Cisco WLC 2504 / 3504 / 5520，傳統 AireOS 系列）
# ==============================================================================

class AireOSParser(BaseParser):
    """
    針對 Cisco AireOS WLC 的專屬解析器（2504 / 3504 / 5520 等型號）。

    資料來源對照：
      Hostname      → show sysinfo → System Name
      OOB IP        → show sysinfo → IP Address
      Firmware      → show sysinfo → Product Version
      CPU           → show sysinfo → CPU Average Usage
      Memory        → show memory statistics → Total/Used/Free bytes
      Model / SN    → show inventory → NAME:"Chassis" → PID / SN
      Fan Status    → show sysinfo → Fan Status (OK → 正常)
      Power Status  → show sysinfo 溫度欄位：若任一溫度超過 Alarm Limit → 異常
      AP 資料       → show ap join stats summary all（Joined 的 AP）
                      + show ap inventory all（比對 AP Name 取 PID/SN）

    Log 邊界格式（AireOS 固定提示符）：
      (Cisco Controller) >show sysinfo
      Manufacturer's Name...
      ...
      (Cisco Controller) >show inventory   ← 邊界
    """

    def _get_command_block(self, cmd: str) -> str:
        """
        AireOS 版本的指令區塊擷取器。

        AireOS CLI 提示符固定為「(Cisco Controller) >」，
        與 IOS 的「hostname#」格式不同，因此覆寫父類別方法。

        Regex 說明：
          \\(Cisco Controller\\)\\s*[>)] → 提示符（[>)] 兼容 > 和 ) 格式）
          {cmd}\\b                       → 指令名稱（全詞匹配）
          .*?\\n                         → 指令行剩餘 + 換行
          (.*?)                          → 非貪婪擷取輸出
          (?=\\n\\(Cisco Controller\\)|\\Z) → 前視到下一個提示符或結尾
        """
        pattern = re.compile(
            rf"\(Cisco Controller\)\s*[>)]\s*{re.escape(cmd)}\b.*?\n(.*?)"
            rf"(?=\n\(Cisco Controller\)|\Z)",
            re.S | re.I
        )
        m = pattern.search(self.content)
        return m.group(1).strip() if m else ""

    def __init__(self, content: str):
        """初始化：預先擷取 AireOS 常用指令區塊。"""
        super().__init__(content)
        self.sysinfo_sec = self._get_command_block("show sysinfo")
        self.inv_sec     = self._get_command_block("show inventory")   # 覆蓋父類別的空字串

    def _parse_hostname(self):
        """
        從 show sysinfo 擷取 WLC 主機名稱。
        若無 show sysinfo (例如 Log 中只有 AP 指令)，則從 CLI 提示符反推。
        """
        m = re.search(r"System Name[.\s]+([\S]+)", self.sysinfo_sec)
        if m:
            self.info.hostname = m.group(1).strip()
        else:
            # 備案：從提示符 (Cisco Controller) > 擷取
            m_prompt = re.search(r"^\((.+?)\)\s*[>)]", self.content, re.M)
            if m_prompt:
                self.info.hostname = m_prompt.group(1).strip()

    def _parse_oob(self):
        """
        從 show sysinfo 的 IP Address 欄位擷取 WLC 管理 IP。

        格式：「IP Address....................................... 203.0.113.120」
        """
        m = re.search(r"IP Address[.\s]+([\d.]+)", self.sysinfo_sec)
        if m:
            self.info.oob_ip = m.group(1)

    def _parse_firmware(self):
        """
        從 show sysinfo 的 Product Version 欄位擷取 AireOS 版本號。

        格式：「Product Version.................................. 8.10.183.0」
        """
        m = re.search(r"Product Version[.\s]+([\d.]+)", self.sysinfo_sec)
        self.info.firmware    = m.group(1).strip() if m else "Unknown"
        self.info.device_type = "AireOS-WLC"

    def _parse_cpu(self):
        """
        從 show sysinfo 的 CPU Average Usage 欄位擷取 CPU 使用率。

        格式：「CPU Average Usage................................ 0」（無 % 符號）
        """
        m = re.search(r"CPU Average Usage[.\s]+(\d+)", self.sysinfo_sec)
        self.info.cpu = f"{m.group(1)}%" if m else "N/A"

    def _parse_memory(self):
        """
        從 show memory statistics 擷取記憶體資訊（bytes 格式）。

        格式：
          Total System Memory............: 3734757376 bytes (3.47 GB)
          Used System Memory.............: 1835905024 bytes (1.70 GB)
          Free System Memory.............: 1898852352 bytes (1.76 GB)
        """
        mem_sec = self._get_command_block("show memory statistics")
        total_m = re.search(r"Total System Memory[.:\s]+([\d]+)\s*bytes", mem_sec)
        used_m  = re.search(r"Used System Memory[.:\s]+([\d]+)\s*bytes",  mem_sec)
        free_m  = re.search(r"Free System Memory[.:\s]+([\d]+)\s*bytes",  mem_sec)
        if total_m and used_m:
            t   = int(total_m.group(1))
            u   = int(used_m.group(1))
            f   = int(free_m.group(1)) if free_m else t - u
            pct = round((u / t) * 100, 1) if t > 0 else 0
            self.info.memory = self.format_mem(t, u, f, pct)

    def _parse_inventory(self):
        """
        從 show inventory 的 NAME:"Chassis" 區塊擷取 WLC 型號（PID）與序號（SN）。

        show inventory 格式：
          NAME: "Chassis"    , DESCR: "Cisco 3500 Series Wireless LAN Controller"
          PID: AIR-CT3504-K9,  VID: V02,  SN: SN000000008

        說明：
          AireOS 的 show inventory 與 IOS 格式幾乎相同，
          直接用 NAME:"Chassis" 抓取 PID/SN 即可。
        """
        m = re.search(
            r'NAME:\s*"Chassis".*?PID:\s*([\w-]+).*?SN:\s*(\w+)',
            self.inv_sec, re.S | re.I
        )
        if m:
            self.members.append(StackMember(1, m.group(1).strip(','), m.group(2).strip()))
        else:
            # 【關鍵修正】若 firmware 也是 Unknown，代表這份檔案沒有 show sysinfo
            # 說明這只是一份單純的 AP 指令片段檔，因此「不要」append 任何設備！
            if self.info.firmware != "Unknown":
                self.members.append(StackMember(1, "WLC (AireOS)", "Unknown"))
                
        self.info.unit_count = len(self.members)

    def _parse_env(self):
        """
        從 show sysinfo 判斷風扇與電源（溫度）環境狀態。

        show sysinfo 環境相關欄位：
          Operating Environment............................ Commercial (10 to 35 C)
          Internal Temp Alarm Limits....................... -10 to 80 C   ← 上限閾值
          Internal Temperature............................. +41 C         ← 實際溫度
          Mgig Temp Alarm Limits........................... -10 to 78 C
          Mgig Temperature................................. +30 C
          External Temp Alarm Limits....................... -10 to 71 C
          External Temperature............................. +24 C
          Fan Status....................................... OK

        風扇判斷：Fan Status == "OK" → 正常，否則 → 異常

        電源（溫度）判斷：
          逐一比對三組溫度（Internal / Mgig / External）：
            1. 從「X Temp Alarm Limits: -N to M C」抓取上限 M
            2. 從「X Temperature: +T C」抓取實際溫度 T
            3. 若 T > M → 超出警戒閾值 → 異常
          任一組超出 → Power Status 標記為異常
          （AireOS 無獨立 PSU 狀態行，以溫度健康代表整體電源健康）
        """
        # ── 風扇狀態 ─────────────────────────────────────────────────────
        fan_m = re.search(r"Fan Status[.\s]+(\w+)", self.sysinfo_sec)
        if fan_m:
            self.fan_status = "正常" if fan_m.group(1).strip().upper() == "OK" else "異常"

        # ── 電源（溫度）狀態：比對各溫度感測點是否超出 Alarm 上限 ────────
        temp_groups = [
            ("Internal Temp Alarm Limits",  "Internal Temperature"),
            ("Mgig Temp Alarm Limits",       "Mgig Temperature"),
            ("External Temp Alarm Limits",   "External Temperature"),
        ]
        pwr_bad = False
        for limit_key, temp_key in temp_groups:
            # 擷取上限：「-10 to 80 C」→ 取 80
            limit_m = re.search(
                rf"{re.escape(limit_key)}[.\s]+[-\d]+\s+to\s+([\d]+)\s*C",
                self.sysinfo_sec, re.I
            )
            # 擷取實際溫度：「+41 C」→ 取 41
            temp_m = re.search(
                rf"{re.escape(temp_key)}[.\s]+\+?([\d]+)\s*C",
                self.sysinfo_sec, re.I
            )
            if limit_m and temp_m:
                if int(temp_m.group(1)) > int(limit_m.group(1)):
                    pwr_bad = True
                    break

        self.pwr_status = "異常" if pwr_bad else "正常"

    def _parse_eth_mac(self):
        """
        AireOS：從 show sysinfo 的「Burned-in MAC Address」欄位擷取。
        格式：「Burned-in MAC Address............................ aa:bb:cc:00:00:07」
        """
        m = re.search(
            r"Burned-in MAC Address[.\s]+([0-9a-fA-F:]{17})",
            self.sysinfo_sec, re.I
        )
        if m:
            self.info.eth_mac = m.group(1).strip()

    def get_ap_rows(self) -> List[Dict]:
        """
        整合 AireOS 的 AP 資訊，產生可寫入報表的資料列。

        資料來源：
          show ap join stats summary all：
            欄位順序：Base Mac | AP EthernetMac | AP Name | IP Address | Status
            → 只取 Status = Joined 的 AP
            → AP EthernetMac 對應 AP Info 分頁的 Ethernet MAC 欄位

          show ap inventory all：
            格式：「Inventory for <AP-Name>」開頭，後接 PID / SN
            → 以 AP Name 為 key 對比，補入 Model（PID）和 Serial Number（SN）

        show ap join stats summary all 格式：
          Base Mac             AP EthernetMac       AP Name    IP Address   Status
          aa:bb:cc:00:00:08    aa:bb:cc:00:00:09    AP-LAB01  203.0.113.121 Joined

        show ap inventory all 格式：
          Inventory for AP-LAB01            ← AP Name
          NAME: "AP2800" ...
          PID: AIR-AP2802I-T-K9, VID: V02, SN: SN000000009

        AP Firmware Version：
          AireOS WLC 集中管理所有 AP 韌體，AP 版本與 WLC 版本相同（均為 self.info.firmware）。
          使用 truncate_ap_firmware() 截取前三段（如 8.10.183.0 → 8.10.183）。
        """
        # AireOS AP 版本 = WLC 版本（WLC 統一下發韌體給所有 AP）
        wlc_fw = truncate_ap_firmware(self.info.firmware)

        # ── Step 1：解析 show ap join stats summary all ─────────────────
        join_sec = self._get_command_block("show ap join stats summary all")
        ap_dict: Dict[str, Dict] = {}

        for line in join_sec.splitlines():
            line = line.strip()
            # 格式：BaseMac(17)  EtherMAC(17)  AP-Name  IP  Status
            m = re.match(
                r"([\da-f:]{17})\s+([\da-f:]{17})\s+(\S+)\s+([\d.]+)\s+(\S+)",
                line, re.I
            )
            if not m:
                continue
            ap_name = m.group(3)
            # 「Not Joined」在 Regex 中被切為 "Not"，需要從原始行補判斷
            is_joined = (re.search(r"\bJoined\b", line, re.I) is not None and
                         re.search(r"\bNot\s+Joined\b", line, re.I) is None)

            # 避免 Not Joined 覆蓋掉先前已記錄為 Joined 的正確狀態
            if ap_name in ap_dict and ap_dict[ap_name]["joined"] and not is_joined:
                continue

            ap_dict[ap_name] = {
                "eth_mac":    m.group(2),     # AP EthernetMac → Ethernet MAC 欄
                "ip_address": m.group(4),
                "joined":     is_joined,
                "ap_model":   "Unknown",
                "serial":     "N/A",
                "firmware":   wlc_fw,         # AP 版本繼承 WLC 版本
            }

        # ── Step 2：解析 show ap inventory all，以 AP Name 配對補入 PID/SN ─
        inv_all_sec = self._get_command_block("show ap inventory all")
        # 以「Inventory for 」為分割點，將整塊切成各 AP 的子區塊
        inv_blocks = re.split(r"Inventory for ", inv_all_sec)
        for block in inv_blocks:
            block = block.strip()
            if not block:
                continue
            # 每個子區塊的第一個 token 即為 AP Name
            ap_name_m = re.match(r"(\S+)", block)
            if not ap_name_m:
                continue
            ap_name = ap_name_m.group(1)
            pid_m   = re.search(r"PID:\s*([\w-]+)", block)
            sn_m    = re.search(r"SN:\s*(\w+)", block)
            if ap_name in ap_dict:
                if pid_m: ap_dict[ap_name]["ap_model"] = pid_m.group(1)
                if sn_m:  ap_dict[ap_name]["serial"]   = sn_m.group(1)

        # ── Step 3：只輸出 Joined 狀態的 AP ────────────────────────────
        rows = []
        for ap_name, info in ap_dict.items():
            if not info.get("joined", False):
                continue
            rows.append({
                "ap_name":    ap_name,
                "ap_model":   info["ap_model"],
                "eth_mac":    info["eth_mac"],
                "ip_address": info["ip_address"],
                "serial":     info["serial"],
                "firmware":   info["firmware"],
            })
        return rows


# ==============================================================================
# 9. Aruba / HPE AOS-CX 解析器（6300 / 6400 等系列）
# ==============================================================================

class ArubaParser(BaseParser):
    """
    針對 Aruba / HPE AOS-CX Switch 的專屬解析器（6300 / 6400 等）。

    與 Cisco IOS 的關鍵差異：
      - 使用 AOS-CX 作業系統（非 Cisco IOS），CLI 提示符為 "hostname#"
      - Log 來自 show tech，區塊邊界為 "Command : show X"
      - Hostname 在 show system → "Hostname: X"（欄位名稱含大寫 H）
      - 管理 IP 在 "interface mgmt → ip static X.X.X.X/prefix"（非 ip address）
      - 韌體版本在 show version → "Version : FL.10.16.1020"，只取數字段 "10.16.1020"
      - 型號在 show system → "Product Name : JL667A 6300F 48G 4SFP56 Sw"，只取前三段
      - 序號在 show system → "Chassis Serial Nbr: X"
      - CPU 在 show system → "CPU Util (% avg 5 min): X"
      - 記憶體在 show system → "Memory Usage (%): X"（只有百分比）
      - 環境：show environment → fan 區段 + power-supply 區段
      - Syslog 來自 show events -s warn，去重後保留不重複的事件訊息
    """

    def __init__(self, content: str):
        """初始化 ArubaParser，預先擷取 AOS-CX 常用區塊。"""
        super().__init__(content)
        self.sys_sec = self._get_aruba_block("show system")
        self.env_sec = self._get_aruba_block("show environment")
        self.cfg_sec = self._get_aruba_block("show running-config")
        self.ver_sec = self._get_aruba_block("show version")   # 覆蓋父類別空字串

    def _get_aruba_block(self, cmd: str) -> str:
        """
        從 AOS-CX show tech 格式的 Log 中擷取特定指令的輸出區塊。

        AOS-CX show tech 的結構特徵：
          *********************************
          Command : show system           ← 指令標頭
          *********************************
          Hostname : 6300                 ← 輸出開始
          ...
          *********************************
          Command : show environment      ← 下一個標頭（邊界）

        Regex：
          Command\\s*:\\s* + {cmd}\\b → 指令標頭行
          .*?\\n\\*+\\s*\\n           → 標頭後的 "*****" 分隔線
          (.*?)                       → 非貪婪擷取輸出
          (?=\\n\\*+\\s*\\nCommand\\s*:|\\Z) → 前視到下一個 Command 標頭或結尾
        """
        pattern = re.compile(
            rf"Command\s*:\s*{re.escape(cmd)}\b.*?\n\*+\s*\n(.*?)(?=\n\*+\s*\nCommand\s*:|\Z)",
            re.S | re.I
        )
        m = pattern.search(self.content)
        return m.group(1).strip() if m else ""

    def _parse_hostname(self):
        """
        從 show system 的 Hostname 欄位擷取設備名稱。

        AOS-CX show system 格式：
          Hostname               : 6300
        """
        m = re.search(r"^Hostname\s*:\s*(\S+)", self.sys_sec, re.M)
        if m:
            self.info.hostname = m.group(1).strip()

    def _parse_oob(self):
        """
        從 running-config 的 interface mgmt 區塊擷取管理 IP。

        AOS-CX interface mgmt 格式：
          interface mgmt
              ip static 198.18.0.245/24    ← 使用 "ip static"，非 "ip address"
        """
        m = re.search(
            r"interface\s+mgmt\b.*?ip\s+static\s+(\d+\.\d+\.\d+\.\d+)/\d+",
            self.cfg_sec, re.S | re.I
        )
        if m:
            self.info.oob_ip = m.group(1)

    def _parse_firmware(self):
        """
        從 show version 的 Version 欄位擷取韌體版本，去除廠商前綴字母。

        AOS-CX show version 格式：
          Version      : FL.10.16.1020

        擷取規則：
          完整值為 "FL.10.16.1020"，前綴 "FL." 為廠商代碼，
          只取純數字版本段 "10.16.1020"，方便與其他設備版本號比較。

        Regex 說明：
          Version\\s*:\\s*   → 欄位標頭
          (?:\\w+\\.)?       → 可選的前綴（FL.、SL. 等），非捕捉群組
          ([\\d.]+)          → 純數字版本（如 10.16.1020），捕捉群組 1
        """
        m = re.search(r"Version\s*:\s*(?:\w+\.)?([\d.]+)", self.ver_sec)
        self.info.firmware    = m.group(1).strip() if m else "Unknown"
        self.info.device_type = "Aruba"

    def _parse_cpu(self):
        """
        從 show system 的 CPU Util (% avg 5 min) 欄位擷取 CPU 使用率。

        AOS-CX show system 格式：
          CPU Util (% avg 5 min) : 11
        """
        m = re.search(r"CPU Util\s*\(%\s*avg\s*5\s*min\)\s*:\s*(\d+)", self.sys_sec)
        self.info.cpu = f"{m.group(1)}%" if m else "N/A"

    def _parse_memory(self):
        """
        從 show system 的 Memory Usage (%) 欄位取得記憶體使用率。

        AOS-CX show system 格式：
          Memory Usage (%)       : 21

        AOS-CX 只提供百分比，無 Total/Used/Free 絕對數值，
        因此格式為 "Usage: 21%"（與 Cisco 的 Total|Used|Free 格式有所不同）。
        """
        m = re.search(r"Memory Usage\s*\(%\)\s*:\s*(\d+)", self.sys_sec)
        self.info.memory = f"Usage: {m.group(1)}%" if m else "N/A"

    def _parse_inventory(self):
        """
        從 show system 擷取型號（Product Name 前三段）與序號（Chassis Serial Nbr）。

        AOS-CX show system 格式：
          Product Name           : JL667A 6300F 48G 4SFP56 Sw
          Chassis Serial Nbr     : SN000000010

        型號截取規則：
          完整 Product Name 為「JL667A 6300F 48G 4SFP56 Sw」，
          只取前三個空白分隔的 token：「JL667A 6300F 48G」，
          包含廠商料號（JL667A）、型號（6300F）、Port 規格（48G），足以識別設備。
          「4SFP56 Sw」為 Uplink 規格和平台描述，不納入型號。
        """
        pid_m = re.search(r"Product Name\s*:\s*([^\r\n]+)", self.sys_sec)
        sn_m  = re.search(r"Chassis Serial Nbr\s*:\s*(\S+)", self.sys_sec)

        # 取前三段 token 作為型號（JL667A 6300F 48G）
        if pid_m:
            parts = pid_m.group(1).strip().split()
            pid = " ".join(parts[:3]) if len(parts) >= 3 else pid_m.group(1).strip()
        else:
            pid = "Unknown-Aruba"

        sn = sn_m.group(1).strip() if sn_m else "Unknown"
        self.members.append(StackMember(1, pid, sn))
        self.info.unit_count = 1

    def _parse_env(self):
        """
        從 show environment 判斷 Aruba Switch 的風扇與電源狀態。

        AOS-CX show environment 格式：
          ─ 風扇 ─
          Fan information
          Location-      ...  Status  RPM
          System-1/1/1   ...  ok      4051
          ─ 電源 ─
          show environment power-supply
          Mbr/PSU  ...  PSU Status  ...
          1/1      ...  OK          ...

        判斷：若 Status 欄出現 FAIL/FAULT/ALARM/ERROR/ABSENT → 異常
        """
        env_bad = re.compile(r"\b(FAIL|FAULT|ALARM|ERROR|ABSENT|CRITICAL)\b", re.I)

        fan_sec_m = re.search(r"Fan information(.*?)(?=\n\nshow environment|\Z)", self.env_sec, re.S | re.I)
        fan_sec   = fan_sec_m.group(1) if fan_sec_m else ""
        fan_bad   = bool(env_bad.search(fan_sec)) if fan_sec else False

        psu_sec_m = re.search(r"show environment power-supply\s*\n.*?\n.*?\n.*?\n(.*?)(?=\n\n|\Z)", self.env_sec, re.S | re.I)
        psu_sec   = psu_sec_m.group(1) if psu_sec_m else ""
        pwr_bad   = bool(env_bad.search(psu_sec)) if psu_sec else False

        self.fan_status = "異常" if fan_bad else "正常"
        self.pwr_status = "異常" if pwr_bad else "正常"

    def _parse_eth_mac(self):
        """
        Aruba AOS-CX：從 show system 的「Base MAC Address」欄位擷取。

        HPE 設備的 MAC 格式與標準格式不同，有兩種可能：
          標準格式：「aa:bb:cc:00:00:0a」（含冒號，17 字元）
          HPE 格式 ：「aabbcc-00000b」（6 hex 連字號 6 hex，13 字元）

        HPE 格式轉換為標準格式（插入冒號）：
          「aabbcc-00000b」→ 去除連字號 → 「aabbcc00000b」
          → 每兩個字元插入冒號 → 「aa:bb:cc:00:00:0c」
        """
        m = re.search(
            r"Base\s+MAC\s+Address\s*[:\s]+"
            r"([0-9a-fA-F:.]{11,17}|[0-9a-fA-F]{6}-[0-9a-fA-F]{6})",
            self.sys_sec, re.I
        )
        if not m:
            return
        raw = m.group(1).strip()
        # HPE 連字號格式：xxxxxx-xxxxxx → xx:xx:xx:xx:xx:xx
        if re.match(r"^[0-9a-fA-F]{6}-[0-9a-fA-F]{6}$", raw):
            hex_str = raw.replace("-", "")
            raw = ":".join(hex_str[i:i+2] for i in range(0, 12, 2))
        self.info.eth_mac = raw

    def _parse_syslog(self):
        """
        從 show events -s warn 擷取 AOS-CX 的 Warning 以上等級事件，
        去除時間戳記後以訊息內容去重，合併重複事件為單一筆記錄。

        AOS-CX show events -s warn 格式：
          2026-04-16T06:26:50.526404+08:00 6300 ndmd[1424]: Event|6134|LOG_ERR|UMM|-|訊息內容
          2026-04-16T06:26:51.512451+08:00 6300 ndmd[1424]: Event|6134|LOG_ERR|UMM|-|訊息內容
          （相同訊息重複出現，因為事件持續發生）

        去重策略：
          1. 從每行擷取「Event|CODE|LEVEL|MOD|-|訊息」部分（去掉時間戳記）
          2. 以訊息內容（EVENT|...|訊息）為 key 做 dict.fromkeys 有序去重
          3. 保留最後 50 筆不重複事件

        格式解析：
          TIMESTAMP HOSTNAME PROCESS: Event|CODE|LEVEL|MOD|-|MESSAGE
          → 捕捉「Event|...」之後的完整訊息段
        """
        events_sec = self._get_aruba_block("show events -s warn")
        if not events_sec:
            return

        raw_msgs = []
        for line in events_sec.splitlines():
            line = line.strip()
            if not line or line.startswith("---") or line.startswith("Event logs"):
                continue
            # 擷取 Event|CODE|LEVEL|MOD|-|MESSAGE 段（去掉前面的時間戳記和主機名）
            m = re.search(r"(Event\|[\d]+\|[A-Z_]+\|[^|]+\|-\|.*)", line)
            if m:
                raw_msgs.append(m.group(1).strip())
            else:
                # 若格式不符，嘗試直接取「:」後的訊息內容
                colon_m = re.search(r":\s+(.+)", line)
                if colon_m:
                    raw_msgs.append(colon_m.group(1).strip())

        # 有序去重（dict.fromkeys 保留第一次出現順序），取後 50 筆
        unique_msgs = list(dict.fromkeys(raw_msgs))[-50:]
        self.syslog = "\n".join(unique_msgs)


# ==============================================================================
# 10. C1300 解析器（Cisco Catalyst 1300 系列，Linux-based ROS）
# ==============================================================================


class C1300Parser(BaseParser):
    """
    針對 Cisco Catalyst 1300 系列 Switch 的專屬解析器。

    ╔══════════════════════════════════════════════════════════════════════╗
    ║              C1300 與 IOS Switch 的根本差異                          ║
    ╠══════════════════════════════════════════════════════════════════════╣
    ║                                                                      ║
    ║  【作業系統差異】                                                     ║
    ║    IOS / IOS-XE：Cisco 自研 OS（MONOS / IOS-XE kernel）             ║
    ║    C1300 ROS   ：基於 Linux 的輕量 OS（Cisco SB 小型交換器系列）      ║
    ║                                                                      ║
    ║  【Log 擷取指令差異】                                                 ║
    ║    IOS     ：逐條執行 show version、show inventory … 各自輸出        ║
    ║    C1300   ：執行一條 show tech-support，彙整所有指令於同一個檔案     ║
    ║                                                                      ║
    ║  【區塊邊界格式差異】                                                 ║
    ║    IOS     ：SW1#show version      ← 提示符 + 指令名稱               ║
    ║              Cisco IOS Software…   ← 輸出開始                       ║
    ║              SW1#show inventory    ← 下一個提示符（邊界）            ║
    ║                                                                      ║
    ║    C1300   ：-- show version --    ← 固定格式標頭                   ║
    ║              Active-image: flash:… ← 輸出開始                       ║
    ║              -- show hw version -- ← 下一個標頭（邊界）              ║
    ║                                                                      ║
    ║  【因此需要】                                                         ║
    ║    1. 獨立的 _get_c1300_block() 方法（不使用父類別的版本）            ║
    ║    2. 獨立的 __init__ 預先擷取 C1300 特有的常用區塊                   ║
    ║    3. 全套覆寫各解析方法，使用 C1300 專屬格式的 Regex                 ║
    ║                                                                      ║
    ║  【C1300 各欄位來源對照表】                                           ║
    ║  ┌─────────────────┬──────────────────────────────────────────────┐  ║
    ║  │ 欄位            │ 來源指令 + 格式                                │  ║
    ║  ├─────────────────┼──────────────────────────────────────────────┤  ║
    ║  │ Hostname        │ show system → System Name: <name>            │  ║
    ║  │ OOB IP          │ show ip interface → IP/mask + vlan N + UP/UP │  ║
    ║  │ Firmware        │ show version → Active-image → Version: X.X.X │  ║
    ║  │ Model           │ show inventory → NAME:"1" → PID: C1300-...   │  ║
    ║  │ Serial          │ show inventory → NAME:"1" → SN: PVN...       │  ║
    ║  │ CPU Usage       │ show cpu utilization → five minutes: X%       │  ║
    ║  │ Memory Usage    │ memory info → Dynamic RAM Requested/Free/Used │  ║
    ║  │ Fan Status      │ show system → Unit Fans Status 區段           │  ║
    ║  │ Power Status    │ show system → Unit Temperature Status 區段    │  ║
    ║  │ Syslog          │ show logging → DD-Mon-YYYY %FAC-E/W-MNE: msg │  ║
    ║  │ Stack           │ show stack → 計算 Unit Id 行數 > 1 即為 Stack │  ║
    ║  └─────────────────┴──────────────────────────────────────────────┘  ║
    ║                                                                      ║
    ╚══════════════════════════════════════════════════════════════════════╝
    """

    def __init__(self, content: str):
        """
        初始化 C1300 解析器。

        覆寫父類別的 __init__，因為父類別的 _get_command_block 使用
        IOS 格式（hostname# + 指令名），無法識別 C1300 的 '-- show X --' 格式。
        這裡改用 _get_c1300_block() 預先擷取 C1300 常用的指令區塊。

        Args:
            content: 整份 show tech-support 檔案的文字內容
        """
        # 呼叫 BaseParser.__init__ 完成基本初始化
        # （self.content、self.info、self.members、self.fan/pwr_status、self.syslog）
        # 但父類別的 ver_sec / inv_sec 會是空字串，後續由 C1300 方法使用自己的區塊
        super().__init__(content)

        # 用 C1300 專屬方法預先擷取常用區塊（效能優化，避免重複解析）
        self.sys_sec  = self._get_c1300_block("show system")          # 系統基本資訊
        self.ver_sec  = self._get_c1300_block("show version")         # 韌體版本（覆蓋父類別）
        self.inv_sec  = self._get_c1300_block("show inventory")       # 硬體庫存（覆蓋父類別）
        self.cpu_sec  = self._get_c1300_block("show cpu utilization") # CPU 使用率
        self.mem_sec  = self._get_c1300_block("memory info")          # 記憶體資訊
        self.log_sec  = self._get_c1300_block("show logging")         # Syslog
        self.ip_sec   = self._get_c1300_block("show ip interface")    # IP 介面狀態
        self.stk_sec  = self._get_c1300_block("show stack")           # Stack 資訊

    # ──────────────────────────────────────────────────────────────────────────
    def _get_c1300_block(self, cmd: str) -> str:
        """
        從 C1300 show tech-support 格式的 Log 中擷取特定指令的輸出區塊。

        C1300 show tech-support 的結構特徵：
          每個指令輸出以固定格式標頭「-- show X --」開始，
          下一個「-- show Y --」或檔案結尾為邊界，例如：
            -- show version --           ← 本指令標頭（含任意數量的 '-'）
            Active-image: flash://...    ← 輸出開始
            ...
            -- show hw version --        ← 下一個標頭（邊界）

        Regex 說明：
          --+          → 一或多個 '-'（標頭左側的連字號）
          \s*cmd\s*   → 指令名稱（前後可有空格）
          --+          → 右側連字號
          \r?\n        → 換行（兼容 Windows \r\n 和 Unix \n）
          (.*?)        → 非貪婪擷取指令輸出（re.S 使 . 匹配換行）
          (?=\r?\n--+\s*show\s|\Z) → 正向前視：下一個標頭或字串結尾

        re.escape(cmd)：
          若指令名稱含有特殊字元（如 '.' 或 '+'），re.escape 確保被視為字面值，
          例如 "memory info" 中的空格和 "show ip interface" 均可正確匹配。

        Args:
            cmd: 指令名稱，例如 "show version"、"memory info"

        Returns:
            指令輸出的純文字（去除首尾空白），找不到則回傳空字串 ""
        """
        pattern = re.compile(
            rf"--+\s*{re.escape(cmd)}\s*--+\s*\r?\n(.*?)(?=\r?\n--+\s*show\s|\Z)",
            re.S | re.I
        )
        m = pattern.search(self.content)
        return m.group(1).strip() if m else ""

    # ──────────────────────────────────────────────────────────────────────────
    def _parse_hostname(self):
        """
        從 show system 的 System Name 欄位擷取設備名稱。

        C1300 Log 格式（show system 區塊）：
          System Name:    C1300-LAB-SW01
                          ↑ System Name 後面跟著大量空白再接名稱

        與 IOS 的差異：
          IOS 從 running-config 的「hostname <名稱>」取得，
          C1300 的 running-config 中有 hostname 行，但 show system 更直接可靠。

        備案：
          若 show system 區塊解析失敗，嘗試從 Log 第一行的 CLI 提示符推斷。
          C1300 的提示符格式為「C1300-LAB-S...#」（名稱被截斷加 ...），
          因此提示符備案準確性較低，以 show system 為主。
        """
        # 主要方法：從 show system 的 System Name 行取得
        m = re.search(r"System Name:\s+(\S+)", self.sys_sec)
        if m:
            self.info.hostname = m.group(1).strip()
            return

        # 備案：從 running-config 的 hostname 行
        m2 = re.search(r"^hostname\s+(\S+)", self.content, re.M | re.I)
        if m2:
            self.info.hostname = m2.group(1)

    # ──────────────────────────────────────────────────────────────────────────
    def _parse_oob(self):
        """
        從 show ip interface 輸出找到狀態為 UP/UP 的管理 IP。

        C1300 show ip interface 的輸出格式：
          IP Address         I/F       I/F Status   Type     ...  Status
                                       admin/oper
          -----------------  --------- ----------   -------  ...  ------
          0.0.0.0/32         vlan 1    UP/DOWN      DHCP     ...  Not received
          100.64.12.70/23    vlan 101  UP/UP        Static   ...  Valid
          ↑                            ↑
          IP/prefix-length              admin/oper 狀態

        解析策略：
          1. 在 show ip interface 區塊中搜尋格式為「IP/mask vlan N UP/UP」的行
          2. IP/mask 的寫法是 CIDR 格式（帶 /prefix），需要只擷取 IP 部分
          3. UP/UP 代表介面管理狀態（admin）和操作狀態（oper）都是啟用的，
             這是確認管理 IP 有效運作的最可靠判斷標準
          4. 若有多個 UP/UP Vlan（L3 Switch 可能有多個路由介面），
             取第一個出現的（通常是最小 Vlan ID，即管理 Vlan）
        """
        # 在 show ip interface 區塊中，找 IP/mask + vlan N + UP/UP 的行
        m = re.search(
            r"(\d+\.\d+\.\d+\.\d+)/\d+\s+vlan\s+\d+\s+UP/UP",
            self.ip_sec, re.I
        )
        if m:
            self.info.oob_ip = m.group(1)

    # ──────────────────────────────────────────────────────────────────────────
    def _parse_firmware(self):
        """
        從 show version 的 Active-image 區塊擷取目前執行的韌體版本號。

        C1300 show version 格式（含 Active 與 Inactive 兩個映像）：
          Active-image: flash://system/images/image_c1300_4.1.7.24_official_key.bin
            Version: 4.1.7.24      ← 擷取此行的版本號
            MD5 Digest: fac373...
            Date: 27-Aug-2025
          Inactive-image: flash://system/images/image_4.1.3.36.bin
            Version: 4.1.3.36      ← 舊版本，不取

        解析策略：
          1. 先找「Active-image:」行作為起點
          2. 在其後緊接的幾行中找「Version: X.X.X.X」
          3. re.S 讓 .* 跨行，確保可跨越 Active-image 行和 Version 行之間的任意內容

        設備類型標記：
          設定 device_type = "C1300"，供後續流程識別（如 GUI 日誌顯示）。
        """
        m = re.search(r"Active-image:.*?\n\s*Version:\s+(\S+)", self.ver_sec, re.S)
        self.info.firmware    = m.group(1).strip() if m else "Unknown"
        self.info.device_type = "C1300"

    # ──────────────────────────────────────────────────────────────────────────
    def _parse_cpu(self):
        """
        從 show cpu utilization 擷取 CPU 五分鐘平均使用率。

        C1300 show cpu utilization 格式：
          CPU utilization service is on.
          CPU utilization
          ---------------
          five seconds: 1%; one minute: 2%; five minutes: 2%
                                                            ↑ 擷取此值

        格式與 IOS 的差異：
          IOS：「CPU utilization for five seconds: 34%/2%; one minute: 12%; five minutes: 9%」
          C1300：「five seconds: 1%; one minute: 2%; five minutes: 2%」
          兩者都有「five minutes: X%」這個子字串，因此可使用相同的 Regex。

        使用 five minutes 值（而非 five seconds）的原因：
          五秒值受瞬間負載影響大（如 MAC 學習爆發時短暫飆高），
          五分鐘平均值更能反映設備的實際穩態負載水位。
        """
        # C1300 格式恰好包含「five minutes: X%」，與 IOS 的 Regex 相容
        m = re.search(r"five minutes:\s*(\d+)%", self.cpu_sec)
        self.info.cpu = f"{m.group(1)}%" if m else "N/A"

    # ──────────────────────────────────────────────────────────────────────────
    def _parse_memory(self):
        """
        從 memory info 的 Dynamic (OS managed) RAM 區段擷取記憶體資訊。

        C1300 memory info 包含兩種記憶體統計：
          Dynamic (OS managed) RAM usage:      ← 主記憶體（Linux OS 管理）
            Requested = 293601280, Free = 73128752, Used = 220472528, Usage = 75%
          Local (ROS managed) RAM usage:       ← ROS 內部記憶體池
            Total = 84934512, Free = 16662648, Used = 68271864, Usage = 80%

        選擇 Dynamic RAM 的原因：
          Dynamic RAM 是 Linux OS 管理的主記憶體（DRAM），代表整體系統的記憶體使用情況，
          與 IOS 的 Processor Pool 概念最接近。
          Local（ROS managed）是 ROS 應用層的內部記憶體池，較為局部，不代表整體狀態。

        欄位對應：
          Requested = 系統請求的總記憶體量（≈ Total）
          Free      = 目前可用記憶體
          Used      = 目前已使用記憶體
          Usage     = 使用率百分比（已由 ROS 計算，直接取用）
        """
        m = re.search(
            r"Dynamic\s*\(OS managed\)\s*RAM\s*usage:.*?"
            r"Requested\s*=\s*(\d+),\s*Free\s*=\s*(\d+),\s*Used\s*=\s*(\d+),\s*Usage\s*=\s*(\d+)%",
            self.mem_sec, re.S | re.I
        )
        if m:
            total = m.group(1)   # Requested（總請求量）
            free  = m.group(2)   # Free（可用量）
            used  = m.group(3)   # Used（已使用量）
            pct   = m.group(4)   # Usage（百分比，ROS 已算好）
            self.info.memory = self.format_mem(total, used, free, pct)

    # ──────────────────────────────────────────────────────────────────────────
    def _parse_inventory(self):
        """
        從 show inventory 擷取機箱型號（PID）與序號（SN），並設定 Stack 資訊。

        C1300 show inventory 格式：
          NAME: "1"  DESCR: "Catalyst 1300 Series ... (C1300-48FP-4X)"
          PID: C1300-48FP-4X  VID:  V02  SN: SN000000011

          NAME: "TenGigabitEthernet1/0/1"  DESCR: "Cisco SFP-10GBase-LR"
          PID: SFP-10G-LR-S    VID: V01     SN: SN000000012

        與 IOS Switch 的差異：
          IOS：NAME: "1", DESCR: "..."（逗號分隔）
          C1300：NAME: "1"  DESCR: "..."（兩個空格分隔，無逗號）
          Regex 使用 .*? 跨行，能同時處理兩種格式。

        Stack 判斷（透過 _parse_stack 方法）：
          C1300 支援 Stack（最多 4 台），由 show stack 的 Unit Id 行數判斷，
          在 _parse_inventory 中呼叫，確保 members 和 is_stack 同步設定。
        """
        # 從 show inventory 取主機箱（NAME: "1"）的 PID 和 SN
        m = re.search(
            r'NAME:\s*"1".*?PID:\s*(\S+).*?SN:\s*(\S+)',
            self.inv_sec, re.S | re.I
        )
        if m:
            pid = m.group(1).strip()
            sn  = m.group(2).strip()
        else:
            pid, sn = "Unknown-C1300", "Unknown"

        # 判斷 Stack 架構並建立成員清單
        self._parse_stack(pid, sn)

    # ──────────────────────────────────────────────────────────────────────────
    def _parse_stack(self, primary_pid: str, primary_sn: str):
        """
        透過 show stack 判斷 C1300 是否為 Stack 架構，並建立成員清單。

        C1300 show stack 格式：
          Topology is Chain

          Unit Id MAC Address         Role       Network Uplink
                                       Port       Port
                                       Type       Type
          ------- ------------------- ---------- ------- ------
             1     aa:bb:cc:00:00:0d    active     gi      te
             2     aa:bb:cc:00:00:0e    standby    gi      te    ← Stack 有此行

        解析策略：
          1. 用 Regex 找所有「^\s+\d+\s+MAC地址\s+角色」格式的行（Unit Id 行）
          2. 行數 > 1 → Stack；行數 = 1 → 單機
          3. 對每台成員，嘗試從 show inventory 取對應的型號和序號
             （Stack 成員在 show inventory 中以 NAME: "N" 標識，N 為 Unit Id）
          4. 若 show inventory 沒有成員 N 的記錄，使用主機箱的 PID 作為備案

        Args:
            primary_pid: 從 show inventory NAME:"1" 取到的主機箱型號
            primary_sn:  從 show inventory NAME:"1" 取到的主機箱序號
        """
        # 在 show stack 中找所有成員行（Unit Id + MAC + Role）
        unit_rows = re.findall(
            r"^\s+(\d+)\s+([\da-f:]{17})\s+(\S+)",
            self.stk_sec, re.M | re.I
        )

        if not unit_rows:
            # show stack 解析失敗：以單機備案，使用主機箱資訊
            self.members.append(StackMember(1, primary_pid, primary_sn))
            self.info.is_stack   = False
            self.info.unit_count = 1
            return

        # Stack 判斷：成員數 > 1 即為 Stack
        self.info.is_stack   = len(unit_rows) > 1
        self.info.unit_count = len(unit_rows)

        for unit_id_str, mac, role in unit_rows:
            uid = int(unit_id_str)

            # 嘗試從 show inventory 取此成員的型號和序號
            # C1300 Stack 成員的 inventory 格式：NAME: "N"（N = Unit Id）
            inv_m = re.search(
                rf'NAME:\s*"{uid}".*?PID:\s*(\S+).*?SN:\s*(\S+)',
                self.inv_sec, re.S | re.I
            )
            if inv_m:
                pid = inv_m.group(1).strip()
                sn  = inv_m.group(2).strip()
            else:
                # inventory 沒有此成員記錄，以主機箱 PID 備案，序號填 Unknown
                pid = primary_pid
                sn  = "Unknown" if uid != 1 else primary_sn

            self.members.append(StackMember(uid, pid, sn))

    # ──────────────────────────────────────────────────────────────────────────
    def _parse_env(self):
        """
        從 show system 的 Fans Status 和 Temperature Status 判斷環境健康狀態。

        C1300 show system 的環境輸出格式：
          Unit                Fans Status
          ----   -------------------------------------
           1                   OK                      ← 正常，若異常會顯示 FAIL/ERROR

          Unit Temperature (Celsius) Status
          ---- --------------------- ----------
           1            34               OK            ← 溫度狀態，正常為 OK

        設計決策：
          1. 風扇狀態（Fan Status）→ 只看 Fans Status 區段
          2. 電源狀態（Power Status）→ 從 Temperature Status 區段判斷
             C1300 沒有獨立的「電源狀態」行，以整體 show system 溫度健康代表電源整體健康
          3. 使用 ENV_BAD_REGEX_C1300（不含 OFF），避免 show power inline
             中 PoE 埠關閉（Oper: Off）誤觸異常判斷

        PoE Fault 說明（不影響電源狀態）：
          show power inline 中若有埠的 Oper 顯示 Fault，
          代表該 PoE 埠供電異常（設備相容性問題），並非主電源故障。
          主電源若有問題，會在 show system 的 Fans/Temperature 區段反映。
        """
        # ── 風扇狀態：從 Fans Status 區段擷取 ─────────────────────────────
        fans_m = re.search(
            r"Fans\s+Status(.*?)(?=Unit\s+Temperature|\Z)",
            self.sys_sec, re.S | re.I
        )
        fans_sec = fans_m.group(1) if fans_m else ""
        fan_bad  = bool(ENV_BAD_REGEX_C1300.search(fans_sec))

        # ── 電源狀態：從 Temperature Status 區段擷取（代表整體健康）──────
        temp_m = re.search(
            r"Temperature.*?Status(.*?)(?=\Z|\n\n)",
            self.sys_sec, re.S | re.I
        )
        temp_sec = temp_m.group(1) if temp_m else ""
        pwr_bad  = bool(ENV_BAD_REGEX_C1300.search(temp_sec))

        self.fan_status = "異常" if fan_bad else "正常"
        self.pwr_status = "異常" if pwr_bad else "正常"

    def _parse_eth_mac(self):
        """
        C1300（Linux-based ROS）：從 show system 的「MAC Address」欄位擷取。
        格式：「MAC Address  : aa:bb:cc:00:00:0f」
        """
        m = re.search(
            r"MAC\s+Address\s*[:\s]+([0-9a-fA-F:]{17})",
            self.sys_sec, re.I
        )
        if m:
            self.info.eth_mac = m.group(1).strip()

    # ──────────────────────────────────────────────────────────────────────────
    def _parse_syslog(self):
        """
        從 show logging 擷取 C1300 格式的 Syslog 訊息。

        C1300 Syslog 格式：
          DD-Mon-YYYY HH:MM:SS :%FACILITY-SEVERITY-MNEMONIC: message
          例：07-Apr-2026 01:41:50 :%STP-W-PORTSTATUS: gi1/0/14: STP status Forwarding

        與 IOS Syslog 的差異：
          IOS 格式：%FACILITY-SEVERITY_NUM-MNEMONIC: message
                    嚴重等級為數字（0~4）
          C1300 格式：帶日期時間前綴 + 嚴重等級為字母（E/W/N/I/D）
          ── 兩者使用不同的 Regex 完全無法互用 ──

        等級字母對照（只收 E 和 W，對應 IOS 的 0~4）：
          E = Emergency/Alert/Critical/Error（IOS 0~3）→ 嚴重，需關注
          W = Warning（IOS 4）→ 警告，需關注
          N = Notice（IOS 5）、I = Info（6）、D = Debug（7）→ 不收集

        過濾邏輯：
          1. 使用 C1300_SYSLOG_REGEX 找出所有 E/W 等級訊息
          2. 以 C1300_EXCLUDE_PREFIXES 排除雜訊（LINK 上下線等正常事件）
          3. dict.fromkeys() 有序去重，保留最後 50 筆
        """
        # C1300_SYSLOG_REGEX 的捕捉群組只取 %FAC-SEV-MNE: msg 部分（去掉日期時間前綴）
        logs_raw = C1300_SYSLOG_REGEX.findall(self.log_sec)
        if not logs_raw:
            return

        # 過濾雜訊前綴
        final = [
            l.strip() for l in logs_raw
            if not any(l.strip().startswith(p) for p in C1300_EXCLUDE_PREFIXES)
        ]

        # 有序去重 → 取後 50 筆 → 合併為多行字串
        self.syslog = "\n".join(list(dict.fromkeys(final))[-200:])


# ==============================================================================
# 9. 主應用程式（GUI）
# ==============================================================================

class CiscoParserApp:
    """
    Tkinter 圖形化介面應用程式。

    操作流程：
      1. 「選擇資料夾」→ 選取含 .log / .txt 的目錄
      2. 「產生分析報表」→ 後台自動解析所有檔案
      3. 選擇儲存路徑 → 輸出 Excel 報表（Report + AP Info 兩個分頁）

    主題切換設計：
      預設為淺色（Light）模式，點擊右上角 Dark 滑桿可切換為深色（Dark）模式。
      所有元件色彩由 THEMES 字典集中管理，_apply_theme() 統一套用。
    """

    # ── 主題色彩配置（淺色 / 深色）──────────────────────────────────────────
    THEMES: Dict[str, Dict[str, str]] = {
        "light": {
            "root_bg":          "#F0F2F5",
            "toolbar_bg":       "#FFFFFF",
            "sep_color":        "#DDE0E6",
            "body_bg":          "#F0F2F5",
            "entry_bg":         "#FFFFFF",
            "entry_fg":         "#1A1A1A",
            "entry_border":     "#C8CDD5",
            "text_bg":          "#FFFFFF",
            "text_fg":          "#1A1A1A",
            "text_border":      "#C8CDD5",
            "label_fg":         "#333333",
            "btn_folder_bg":    "#EAEDF1",
            "btn_folder_fg":    "#333333",
            "btn_run_bg":       "#2F75B5",
            "btn_run_fg":       "#FFFFFF",
            "btn_run_active":   "#1A5FA0",
            "progress_trough":  "#DDE0E6",
            "progress_fill":    "#2F75B5",
            "toggle_track_off": "#C8CDD5",
            "toggle_track_on":  "#2F75B5",
            "toggle_thumb":     "#FFFFFF",
        },
        "dark": {
            "root_bg":          "#1A1B2E",
            "toolbar_bg":       "#252639",
            "sep_color":        "#35364D",
            "body_bg":          "#1A1B2E",
            "entry_bg":         "#2E2F45",
            "entry_fg":         "#D8D9E8",
            "entry_border":     "#3E3F58",
            "text_bg":          "#21223A",
            "text_fg":          "#C4C5D6",
            "text_border":      "#3E3F58",
            "label_fg":         "#C4C5D6",
            "btn_folder_bg":    "#35364D",
            "btn_folder_fg":    "#C4C5D6",
            "btn_run_bg":       "#2F75B5",
            "btn_run_fg":       "#FFFFFF",
            "btn_run_active":   "#1A5FA0",
            "progress_trough":  "#2E2F45",
            "progress_fill":    "#2F75B5",
            "toggle_track_off": "#555570",
            "toggle_track_on":  "#2F75B5",
            "toggle_thumb":     "#FFFFFF",
        },
    }

    def __init__(self, root: tk.Tk):
        """
        初始化應用程式：設定視窗屬性並建立所有 UI 元件。

        Args:
            root: Tkinter 主視窗物件（由呼叫端的 tk.Tk() 建立）
        """
        self.root = root
        self.root.title("Network Analysis Tool")
        self.root.resizable(True, True)

        # ── 螢幕比例自適應（基準：1920 × 1200）───────────────────────────
        # 取得實際螢幕解析度，等比例縮放視窗初始大小與最小尺寸，
        # 確保在各種顯示器上按鈕區域都不會被擠出畫面。
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        scale    = min(sw / 1920, sh / 1200)          # 以短邊比例為準
        win_w    = max(int(905  * scale), 480)
        win_h    = max(int(600  * scale), 380)
        min_w    = max(int(480  * scale), 380)
        # 最小高度 = toolbar(48) + folder(46) + progress(18) + sep(2) + btn_frame(70) + padding
        min_h    = max(int(400  * scale), 260)
        self.root.geometry(f"{win_w}x{win_h}")
        self.root.minsize(min_w, min_h)

        self.selected_path = tk.StringVar()    # 綁定路徑輸入框的 tkinter 字串變數
        self.data_rows: List[Dict] = []        # 所有 Switch/WLC 設備的 Excel 資料列
        self.ap_rows:   List[Dict] = []        # 所有 AP 的原始解析資料
        self._is_dark   = False                # 當前主題狀態（False = 淺色）
        self._btn_enabled = True               # 產生報表按鈕是否可按

        # ── Log 佇列（Performance 優化）──────────────────────────────────
        # 背景 Thread 將訊息放入 queue；主 Thread 每 100ms 批次取出寫入文字框，
        # 避免高頻率的 root.after 呼叫造成 GUI 卡頓。
        self._log_queue: queue.Queue = queue.Queue()
        self._start_log_poller()               # 啟動批次輪詢

        # ── 應用程式圖示（iron man）──────────────────────────────────────
        try:
            icon_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "icon_ironman.ico"
            )
            self.root.iconbitmap(icon_path)
        except Exception:
            pass                               # 找不到圖示時靜默略過

        self._setup_ui()
        self._apply_theme()                    # 初始套用淺色主題

    # ── UI 建立 ───────────────────────────────────────────────────────────────

    def _setup_ui(self):
        """建立並排版所有 GUI 元件。"""

        # ── 字型常數（集中管理）────────────────────────────────────────────
        FONT_ZH      = "Microsoft JhengHei"   # 中文字型
        FONT_MONO    = "Consolas"             # 等寬字型（log 輸出）
        FONT_SYM     = "Arial"               # 符號 / 英文字型

        # ── 頂部工具列（白底橫欄）──────────────────────────────────────────
        self.toolbar = tk.Frame(self.root, height=48)
        self.toolbar.pack(fill="x", side="top")
        self.toolbar.pack_propagate(False)

        # 漢堡選單圖示（裝飾用）
        self.lbl_menu = tk.Label(
            self.toolbar, text="≡",
            font=(FONT_SYM, 16), padx=12, pady=10, cursor="hand2"
        )
        self.lbl_menu.pack(side="left")

        # 標題文字
        self.lbl_title = tk.Label(
            self.toolbar, text="Network Analysis Tool by TING",
            font=(FONT_SYM, 12, "bold"), padx=4
        )
        self.lbl_title.pack(side="left")

        # ── Dark 切換 Toggle（右側）──────────────────────────────────────
        self.toggle_outer = tk.Frame(self.toolbar)
        self.toggle_outer.pack(side="right", padx=14)

        self.lbl_dark = tk.Label(
            self.toggle_outer, text="Dark",
            font=(FONT_SYM, 10), padx=6
        )
        self.lbl_dark.pack(side="left")

        # Canvas 繪製圓角 Toggle Slider（寬 46 × 高 24 px）
        self.toggle_canvas = tk.Canvas(
            self.toggle_outer, width=46, height=24,
            bd=0, highlightthickness=0, cursor="hand2"
        )
        self.toggle_canvas.pack(side="left", pady=12)
        self.toggle_canvas.bind("<Button-1>", lambda _e: self._toggle_theme())

        # ── 頂部分隔線 ───────────────────────────────────────────────────
        self.sep_top = tk.Frame(self.root, height=1)
        self.sep_top.pack(fill="x")

        # ── 資料夾選擇列 ────────────────────────────────────────────────
        self.folder_frame = tk.Frame(self.root, pady=10)
        self.folder_frame.pack(fill="x", padx=14)

        self.lbl_folder = tk.Label(
            self.folder_frame, text="選擇資料夾",
            font=(FONT_ZH, 10), width=8, anchor="w"   # ← 微軟正黑體
        )
        self.lbl_folder.pack(side="left", padx=(0, 8))

        # 路徑輸入框（外框容器模擬 border）
        self.entry_border_frame = tk.Frame(self.folder_frame, bd=1, relief="flat")
        self.entry_border_frame.pack(side="left", fill="x", expand=True)

        self.entry_path = tk.Entry(
            self.entry_border_frame,
            textvariable=self.selected_path,
            state="readonly",
            font=(FONT_MONO, 10),
            relief="flat",
            bd=4,
        )
        self.entry_path.pack(fill="x", expand=True)

        # 資料夾按鈕（外框容器 → light 模式顯示邊框）
        self.btn_folder_border = tk.Frame(self.folder_frame, padx=1, pady=1)
        self.btn_folder_border.pack(side="left", padx=(6, 0))

        self.btn_folder = tk.Button(
            self.btn_folder_border, text="🗁",
            font=(FONT_SYM, 13),
            relief="flat", bd=0,
            cursor="hand2",
            padx=8, pady=2,
            command=self.select_folder
        )
        self.btn_folder.pack()

        # ── 注意佈局順序（關鍵）────────────────────────────────────────
        # pack 使用 side="bottom" 先把固定高度元件「鎖」在底部，
        # txt_log 最後才 pack(expand=True)，它只能取「剩餘」空間。
        # 視窗縮小時 txt_log 優先縮小，btn_frame / progress 始終可見。

        # ── 產生報表按鈕（固定 70px 高，鎖底）──────────────────────────
        self.btn_frame = tk.Frame(self.root, height=70)
        self.btn_frame.pack(fill="x", side="bottom")
        self.btn_frame.pack_propagate(False)

        self.btn_canvas = tk.Canvas(
            self.btn_frame, width=220, height=46,
            bd=0, highlightthickness=0, cursor="hand2"
        )
        self.btn_canvas.place(relx=0.5, rely=0.5, anchor="center")
        self.btn_canvas.bind("<Button-1>",  lambda _e: self._on_run_btn_click())
        self.btn_canvas.bind("<Enter>",     lambda _e: self._on_run_btn_hover(True))
        self.btn_canvas.bind("<Leave>",     lambda _e: self._on_run_btn_hover(False))
        self._btn_hover = False
        self.btn_frame.bind("<Configure>", self._on_btn_frame_resize)

        # ── 底部分隔線（鎖底）───────────────────────────────────────────
        self.sep_bot = tk.Frame(self.root, height=1)
        self.sep_bot.pack(fill="x", side="bottom")

        # ── 進度條（鎖底）───────────────────────────────────────────────
        self.prog_frame = tk.Frame(self.root)
        self.prog_frame.pack(fill="x", padx=14, pady=(0, 4), side="bottom")

        self.pb_style = ttk.Style()
        self.pb_style.layout("Flat.Horizontal.TProgressbar", [
            ("Horizontal.Progressbar.trough", {
                "children": [("Horizontal.Progressbar.pbar", {"side": "left", "sticky": "ns"})],
                "sticky": "nswe"
            })
        ])

        self.progress = ttk.Progressbar(
            self.prog_frame,
            orient="horizontal",
            mode="determinate",
            style="Flat.Horizontal.TProgressbar"
        )
        self.progress.pack(fill="x")

        # ── 日誌輸出文字框（最後 pack，取剩餘空間，可以縮）─────────────
        self.log_border_frame = tk.Frame(self.root, bd=1, relief="flat")
        self.log_border_frame.pack(fill="both", expand=True, padx=14, pady=(0, 4))

        self.txt_log = tk.Text(
            self.log_border_frame,
            font=(FONT_MONO, 9),
            relief="flat",
            bd=6,
            wrap="word",
        )
        self.txt_log.pack(fill="both", expand=True)

        # 字型常數儲存供 _draw_run_btn 使用
        self._font_zh = FONT_ZH

    # ── 主題切換 ──────────────────────────────────────────────────────────────

    def _draw_toggle(self):
        """在 Canvas 上繪製圓角 Toggle Slider。"""
        t   = self.THEMES["dark" if self._is_dark else "light"]
        c   = self.toggle_canvas
        c.delete("all")

        W, H = 46, 24
        r    = H // 2                               # 圓角半徑 = 12
        col  = t["toggle_track_on"] if self._is_dark else t["toggle_track_off"]
        tx   = W - r - 2 if self._is_dark else r + 2  # 拇指圓心 X

        # 軌道：左半圓 + 中間矩形 + 右半圓
        c.create_oval(0,    0, H,   H,   fill=col, outline="")
        c.create_oval(W-H, 0, W,    H,   fill=col, outline="")
        c.create_rectangle(r, 0, W-r, H, fill=col, outline="")

        # 拇指（白色圓形）
        pad = 3
        c.create_oval(tx - r + pad, pad,
                      tx + r - pad, H - pad,
                      fill=t["toggle_thumb"], outline="")

    def _on_btn_frame_resize(self, event):
        """
        視窗寬度變化時，動態調整 Canvas 按鈕的寬度，保持置中且不超出邊界。
        使用 after() 節流，避免 Configure 事件頻繁觸發重繪。
        """
        if hasattr(self, "_resize_job"):
            self.root.after_cancel(self._resize_job)
        self._resize_job = self.root.after(30, lambda: self._do_btn_resize(event.width))

    def _do_btn_resize(self, frame_width: int):
        """實際執行 Canvas 按鈕的重新定尺寸與重繪。"""
        new_w = min(max(frame_width - 60, 180), 360)
        self.btn_canvas.config(width=new_w)
        self._draw_run_btn(hover=self._btn_hover)

    def _draw_run_btn(self, hover: bool = False):
        """在 Canvas 上繪製圓角「產生分析報表」按鈕。"""
        c  = self.btn_canvas
        c.delete("all")
        t  = self.THEMES["dark" if self._is_dark else "light"]
        W  = int(self.btn_canvas.cget("width"))   # 動態取得當前寬度
        H, r = 46, 12                              # 高度與圓角半徑固定

        if not self._btn_enabled:
            bg = "#888899" if self._is_dark else "#AABBD0"
        elif hover:
            bg = t["btn_run_active"]
        else:
            bg = t["btn_run_bg"]

        fg = t["btn_run_fg"]

        # 繪製圓角矩形（四個角弧 + 三個填充矩形）
        c.create_arc(0,       0,       2*r,   2*r,   start=90,  extent=90,  fill=bg, outline="")
        c.create_arc(W-2*r,   0,       W,     2*r,   start=0,   extent=90,  fill=bg, outline="")
        c.create_arc(0,       H-2*r,   2*r,   H,     start=180, extent=90,  fill=bg, outline="")
        c.create_arc(W-2*r,   H-2*r,   W,     H,     start=270, extent=90,  fill=bg, outline="")
        c.create_rectangle(r,  0,    W-r, H,   fill=bg, outline="")
        c.create_rectangle(0,  r,    W,   H-r, fill=bg, outline="")

        # 按鈕文字（垂直置中）→ 微軟正黑體
        font_zh = getattr(self, "_font_zh", "Microsoft JhengHei")
        c.create_text(W // 2, H // 2,
                      text="產生分析報表",
                      fill=fg,
                      font=(font_zh, 11, "bold"))

    def _on_run_btn_click(self):
        """Canvas 按鈕點擊事件：僅在啟用狀態下觸發 start_process。"""
        if self._btn_enabled:
            self.start_process()

    def _on_run_btn_hover(self, entering: bool):
        """Canvas 按鈕 hover 事件：Hover 時加深背景色。"""
        self._btn_hover = entering
        self._draw_run_btn(hover=entering)

    def _set_run_btn_state(self, enabled: bool):
        """統一管理按鈕啟用 / 停用狀態並重繪。"""
        self._btn_enabled = enabled
        cursor = "hand2" if enabled else "watch"
        self.btn_canvas.config(cursor=cursor)
        self._draw_run_btn(hover=False)

    def _toggle_theme(self):
        """切換深淺色主題並重繪 UI。"""
        self._is_dark = not self._is_dark
        self._draw_toggle()
        self._apply_theme()

    def _apply_theme(self):
        """將目前主題的色彩套用到所有 UI 元件。"""
        t = self.THEMES["dark" if self._is_dark else "light"]

        # ── 視窗 & 框架背景 ─────────────────────────────────────────────
        self.root.config(bg=t["root_bg"])
        self.toolbar.config(bg=t["toolbar_bg"])
        self.toggle_outer.config(bg=t["toolbar_bg"])
        self.sep_top.config(bg=t["sep_color"])
        self.sep_bot.config(bg=t["sep_color"])
        self.folder_frame.config(bg=t["body_bg"])
        self.log_border_frame.config(bg=t["text_border"])
        self.prog_frame.config(bg=t["body_bg"])
        self.btn_frame.config(bg=t["body_bg"])

        # ── 標籤 ────────────────────────────────────────────────────────
        self.lbl_menu.config(bg=t["toolbar_bg"],  fg=t["label_fg"])
        self.lbl_title.config(bg=t["toolbar_bg"], fg=t["label_fg"])
        self.lbl_dark.config(bg=t["toolbar_bg"],  fg=t["label_fg"])
        self.lbl_folder.config(bg=t["body_bg"],   fg=t["label_fg"])

        # ── Toggle Canvas ────────────────────────────────────────────────
        self.toggle_canvas.config(bg=t["toolbar_bg"])
        self._draw_toggle()

        # ── 路徑輸入框 ───────────────────────────────────────────────────
        self.entry_border_frame.config(bg=t["entry_border"])
        self.entry_path.config(
            bg=t["entry_bg"],
            fg=t["entry_fg"],
            readonlybackground=t["entry_bg"],
            disabledbackground=t["entry_bg"],
            insertbackground=t["entry_fg"],
        )

        # ── 資料夾按鈕 ───────────────────────────────────────────────────
        # light 模式：外框容器顯示邊框色；dark 模式：外框與按鈕同色（無框感）
        folder_border_bg = t["entry_border"] if not self._is_dark else t["btn_folder_bg"]
        self.btn_folder_border.config(bg=folder_border_bg)
        self.btn_folder.config(
            bg=t["btn_folder_bg"],
            fg=t["btn_folder_fg"],
            activebackground=t["btn_run_bg"],
            activeforeground=t["btn_run_fg"],
        )

        # ── 日誌文字框 ───────────────────────────────────────────────────
        self.txt_log.config(
            bg=t["text_bg"],
            fg=t["text_fg"],
            insertbackground=t["text_fg"],
        )

        # ── 進度條 ───────────────────────────────────────────────────────
        self.pb_style.configure(
            "Flat.Horizontal.TProgressbar",
            troughcolor=t["progress_trough"],
            background=t["progress_fill"],
            borderwidth=0,
            thickness=6,
        )

        # ── 產生報表按鈕（Canvas）────────────────────────────────────────
        self.btn_frame.config(bg=t["body_bg"])
        self.btn_canvas.config(bg=t["body_bg"])
        self._draw_run_btn(hover=self._btn_hover)

    def select_folder(self):
        """開啟系統資料夾選擇對話框，將路徑寫入 selected_path。"""
        p = filedialog.askdirectory()
        if p:
            self.selected_path.set(p)

    def log(self, msg: str):
        """
        將訊息放入 Thread-Safe 佇列，由主 Thread 的輪詢器批次寫入文字框。

        Performance 優化：原本每條 log 都呼叫一次 root.after()，
        在解析大量檔案時（數十~數百條訊息/秒）會造成 GUI 事件佇列壅塞。
        改為 queue + 100ms 批次處理，大幅降低 UI 更新開銷。
        """
        self._log_queue.put(msg)

    def _start_log_poller(self):
        """
        每 100ms 批次取出 log 佇列中的所有訊息，一次寫入文字框。
        使用 root.after 遞迴排程，確保在主 Thread 中執行（Thread-Safe）。
        """
        def _poll():
            lines = []
            try:
                while True:                        # 一次取出所有待顯示訊息
                    lines.append(self._log_queue.get_nowait())
            except queue.Empty:
                pass

            if lines:
                self.txt_log.insert(tk.END, "\n".join(lines) + "\n")
                self.txt_log.see(tk.END)

            self.root.after(100, _poll)            # 100ms 後再次輪詢

        self.root.after(100, _poll)

    def detect_type(self, content: str) -> str:
        """
        根據 Log 內容的關鍵字自動判斷設備類型。

        判斷順序（由特定到一般，防止誤判）：
          1. C1300    → 全文含「Catalyst 1300」或「PID: C1300-」
                        C1300 使用 show tech-support 格式，無「Cisco IOS」字樣，
                        須在其他平台之前優先判斷，避免落入 Unknown
          2. AireOS   → 全文含「(Cisco Controller)」（AireOS WLC 固定提示符）
                        或含「RTOS Version」（AireOS 特有欄位）
                        ⚠️ 必須在 WLC 之前判斷，避免被誤判為 C9800 WLC
          3. Aruba    → 全文含「AOS-CX」（HPE Aruba AOS-CX 作業系統識別字串）
                        ⚠️ 必須在 Switch 之前判斷，避免被誤套 IOS 解析邏輯
          4. NX-OS    → show version 區塊含「NXOS」或「Nexus」
          5. WLC      → show version 區塊含「Wireless LAN Controller」或「C9800」
          6. Switch   → show version 區塊含「Cisco IOS」或「Software」
          7. Unknown  → 以上皆不符合，記錄警告後跳過此檔案

        Args:
            content: 完整 Log 文字

        Returns:
            "C1300" / "AireOS" / "Aruba" / "NX-OS" / "WLC" / "Switch" / "Unknown"
        """
        # C1300 優先判斷
        if "Catalyst 1300" in content or re.search(r"PID:\s*C1300-", content):
            return "C1300"

        # AireOS WLC（2504/3504/5520）：提示符 "(Cisco Controller)" 是最可靠的識別
        if "(Cisco Controller)" in content or "RTOS Version" in content:
            return "AireOS"

        # Aruba AOS-CX Switch：AOS-CX 是唯一識別字串
        if "AOS-CX" in content:
            return "Aruba"

        # 其餘設備：取 show version 區塊再判斷（防止 CDP neighbor 輸出干擾）
        m_ver = re.search(r"show version\b.*?\n(.*?)(?=\n\S+#|\Z)", content, re.S | re.I)
        ver_text = m_ver.group(1) if m_ver else content[:5000]

        if "NXOS" in ver_text or "Nexus" in ver_text:                    return "NX-OS"
        if "Wireless LAN Controller" in ver_text or "C9800" in ver_text: return "WLC"
        if "Cisco IOS" in ver_text or "Software" in ver_text:            return "Switch"
        return "Unknown"

    def run(self):
        """
        主要處理流程，在背景 Thread 執行（由 start_process() 啟動）。

        步驟：
          1. 取得目標資料夾中所有 .log / .txt 檔案
          2. 逐一讀取，自動判斷設備類型
          3. 建立對應解析器，呼叫 run_all() 完成解析
          4. 收集設備資料列（data_rows）
          5. WLC 額外解析 AP 資訊（ap_rows）
          6. 更新進度條（透過 root.after 在主 Thread 執行）
          7. 呼叫 save() 輸出 Excel
        """
        folder = self.selected_path.get()
        files = [
            os.path.join(folder, f) for f in os.listdir(folder)
            if f.lower().endswith(('.log', '.txt'))
        ]
        total = len(files)
        self.data_rows = []   # 每次執行前清空
        self.ap_rows   = []

        for i, fp in enumerate(files):
            fname = os.path.basename(fp)
            try:
                # errors='ignore'：遇到無法解碼的字元（如 BOM）跳過
                with open(fp, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()

                dev_type = self.detect_type(content)

                # 依設備類型選擇解析器
                if   dev_type == "NX-OS":   parser = NexusParser(content)
                elif dev_type == "WLC":     parser = WLCParser(content)
                elif dev_type == "Switch":  parser = IOSSwitchParser(content)
                elif dev_type == "C1300":   parser = C1300Parser(content)
                elif dev_type == "AireOS":  parser = AireOSParser(content)
                elif dev_type == "Aruba":   parser = ArubaParser(content)
                else:
                    self.log(f"[{i+1}/{total}] {fname} -> 型態: Unknown [WARN]")
                    continue

                parser.run_all()
                self.data_rows.extend(parser.generate_rows())

                # WLC（C9800 IOS-XE）額外解析 AP 資訊
                if dev_type == "WLC":
                    ap_data = parser.get_ap_rows()
                    self.ap_rows.extend(ap_data)
                    self.log(f"  -> 解析到 {len(ap_data)} 台 AP（C9800）")

                # AireOS WLC 額外解析 AP 資訊
                if dev_type == "AireOS":
                    ap_data = parser.get_ap_rows()
                    self.ap_rows.extend(ap_data)
                    self.log(f"  -> 解析到 {len(ap_data)} 台 AP（AireOS）")

                stack_str = "Stack" if parser.info.is_stack else "Single"
                self.log(f"[{i+1}/{total}] {fname} -> {dev_type} | {stack_str} | {parser.info.unit_count} 台")

            except Exception as e:
                self.log(f"  [ERR] {fname}: {e}")

            # 更新進度條（必須在主 Thread 中執行）
            self.root.after(0, lambda v=((i + 1) / total) * 100: self.progress.configure(value=v))

        self.save()

    def start_process(self):
        """
        「產生分析報表」按鈕的事件處理器。
        驗證路徑後，停用按鈕並在背景 Thread 中啟動 run()。

        daemon=True：主視窗關閉時，背景 Thread 自動終止，不阻塞程式退出。
        """
        if not self.selected_path.get():
            messagebox.showwarning("警告", "請選擇資料夾")
            return
        self._set_run_btn_state(False)
        threading.Thread(target=self.run, daemon=True).start()

    def save(self):
        """
        將所有解析結果輸出為 Excel 報表（.xlsx），包含四個分頁：

        【Report 分頁】
          - 所有設備 + AP（不含 Ethernet MAC）
          - Stack 設備的 ○ 列緊跟主列

        【AP Info 分頁】
          - AP 詳細資訊：Name / Model / Ethernet MAC / IP / Serial / Software Version

        【WLC Info 分頁】（新增）
          - device_type 為 "WLC"（C9800）或 "AireOS-WLC" 的設備
          - 包含 Ethernet MAC 欄位，其餘欄位同 Report

        【Switch Info 分頁】（新增）
          - device_type 為 "Switch" / "NX-OS" / "C1300" / "Aruba" 的設備
          - 包含 Ethernet MAC 欄位，其餘欄位同 Report
        """
        if not self.data_rows:
            self.root.after(0, lambda: [
                messagebox.showwarning("警告", "無可分析資料"),
                self._set_run_btn_state(True)
            ])
            return

        default_name = f"Network_Analysis_{datetime.now().strftime('%m%d_%H%M')}.xlsx"
        path = filedialog.asksaveasfilename(
            initialfile=default_name,
            defaultextension=".xlsx"
        )
        if not path:
            self.root.after(0, lambda: self._set_run_btn_state(True))
            return

        try:
            import ipaddress

            def _ip_sort_key(ip_str: str) -> tuple:
                try:
                    return (0, ipaddress.ip_address(ip_str))
                except ValueError:
                    return (1, ipaddress.ip_address("0.0.0.0"))

            def _sort_rows_by_ip(data_rows: List[Dict]) -> List[Dict]:
                """依 OOB IP 升冪排序，保持 Stack ○ 列緊跟主列。"""
                groups: List[List[Dict]] = []
                current_group: List[Dict] = []
                for row in data_rows:
                    if row.get("OOB IP") == "○":
                        current_group.append(row)
                    else:
                        if current_group:
                            groups.append(current_group)
                        current_group = [row]
                if current_group:
                    groups.append(current_group)
                groups.sort(key=lambda g: _ip_sort_key(g[0].get("OOB IP", "")))
                return [row for group in groups for row in group]

            # ── 設備列依 IP 排序 ──────────────────────────────────────────
            self.data_rows = _sort_rows_by_ip(self.data_rows)

            # ── AP 依 IP 排序 ─────────────────────────────────────────────
            self.ap_rows.sort(key=lambda r: _ip_sort_key(r.get("ip_address", "")))

            # ── WLC / Switch 設備類型常數 ─────────────────────────────────
            WLC_TYPES    = {"WLC", "AireOS-WLC"}
            SWITCH_TYPES = {"Switch", "NX-OS", "C1300", "Aruba"}

            # ── 欄位定義（Report 不含 Ethernet MAC）──────────────────────
            REPORT_COLS = [
                "Hostname", "Model", "OOB IP", "Serial Number",
                "Firmware Version", "CPU Usage", "Memory Usage",
                "Fan Status", "Power Status", "System Log",
            ]
            DETAIL_COLS = [
                "Hostname", "Model", "OOB IP", "Ethernet MAC",
                "Serial Number", "Firmware Version", "CPU Usage",
                "Memory Usage", "Fan Status", "Power Status", "System Log",
            ]

            # ── Step 1：Report 分頁（設備 + AP，不含 Ethernet MAC）────────
            ap_report_rows = [
                {
                    "Hostname":         ap["ap_name"],
                    "Model":            ap["ap_model"],
                    "OOB IP":           ap["ip_address"],
                    "Serial Number":    ap["serial"],
                    "Firmware Version": ap["firmware"],
                    "CPU Usage":        "N/A",
                    "Memory Usage":     "N/A",
                    "Fan Status":       "正常",
                    "Power Status":     "正常",
                    "System Log":       "",
                }
                for ap in self.ap_rows
            ]
            # 從 data_rows 去除私有欄位，只保留 REPORT_COLS
            device_report_rows = [
                {k: r[k] for k in REPORT_COLS if k in r}
                for r in self.data_rows
            ]
            report_rows = device_report_rows + ap_report_rows

            # ── Step 2：AP Info 分頁 ──────────────────────────────────────
            ap_info_rows = [
                {
                    "AP Name":          ap["ap_name"],
                    "AP Model":         ap["ap_model"],
                    "Ethernet MAC":     ap["eth_mac"],
                    "IP Address":       ap["ip_address"],
                    "AP Serial Number": ap["serial"],
                    "Software Version": ap["firmware"],
                }
                for ap in self.ap_rows
            ]

            # ── Step 3：WLC Info 分頁（含 Ethernet MAC）──────────────────
            wlc_rows = [
                {k: r.get(k, "N/A") for k in DETAIL_COLS}
                for r in self.data_rows
                if r.get("_device_type") in WLC_TYPES
            ]

            # ── Step 4：Switch Info 分頁（含 Ethernet MAC）───────────────
            sw_rows = [
                {k: r.get(k, "N/A") for k in DETAIL_COLS}
                for r in self.data_rows
                if r.get("_device_type") in SWITCH_TYPES
            ]

            # ── Step 5：建立 DataFrame ────────────────────────────────────
            df_report  = pd.DataFrame(report_rows,  columns=REPORT_COLS)
            df_ap_info = pd.DataFrame(ap_info_rows) if ap_info_rows else pd.DataFrame(
                columns=["AP Name", "AP Model", "Ethernet MAC", "IP Address",
                         "AP Serial Number", "Software Version"]
            )
            df_wlc = pd.DataFrame(wlc_rows, columns=DETAIL_COLS) if wlc_rows else pd.DataFrame(
                columns=DETAIL_COLS
            )
            df_sw  = pd.DataFrame(sw_rows,  columns=DETAIL_COLS) if sw_rows  else pd.DataFrame(
                columns=DETAIL_COLS
            )

            # ── Step 6：寫入 Excel 並套用格式 ────────────────────────────
            HEADER_FILL  = PatternFill("solid", fgColor="2F75B5")
            HEADER_FONT  = Font(name="微軟正黑體", size=12, bold=True, color="FFFFFF")
            DATA_FONT    = Font(name="微軟正黑體", size=12)
            HEADER_ALIGN = Alignment(horizontal="center")
            THIN_BORDER  = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )

            sheet_map = {
                "Report":      df_report,
                "AP Info":     df_ap_info,
                "WLC Info":    df_wlc,
                "Switch Info": df_sw,
            }

            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                for sheet_name, df in sheet_map.items():
                    df.to_excel(writer, index=False, sheet_name=sheet_name)

                for sheet_name in sheet_map:
                    ws = writer.sheets[sheet_name]
                    for col in ws.columns:
                        max_length = 0
                        col_letter = col[0].column_letter
                        for cell in col:
                            if cell.row == 1:
                                cell.fill      = HEADER_FILL
                                cell.font      = HEADER_FONT
                                cell.alignment = HEADER_ALIGN
                            else:
                                cell.font = DATA_FONT
                            cell.border = THIN_BORDER
                            try:
                                if cell.value:
                                    for line in str(cell.value).split('\n'):
                                        length = len(line.encode('big5', errors='ignore'))
                                        if length > max_length:
                                            max_length = length
                            except Exception:
                                pass
                        final_width = min(int(max_length * 1.2) + 4, 100)
                        ws.column_dimensions[col_letter].width = final_width

            self.root.after(0, lambda: [
                messagebox.showinfo("成功", f"報表已儲存"),
                self._set_run_btn_state(True)
            ])

        except Exception as e:
            self.root.after(0, lambda: [
                messagebox.showerror("錯誤", f"儲存失敗: {e}"),
                self._set_run_btn_state(True)
            ])


# ==============================================================================
# 程式進入點
# ==============================================================================

if __name__ == "__main__":
    root = tk.Tk()               # 建立 Tkinter 主視窗
    app  = CiscoParserApp(root)  # 建立應用程式實例（同時呼叫 _setup_ui 建立 UI）
    root.mainloop()              # 進入 Tkinter 事件迴圈（等待使用者操作直到關閉視窗）