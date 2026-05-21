# 🌐 Network Log Analysis Tool

> 自動解析 Cisco / Aruba / Fortigate 設備 Log，一鍵產生 Excel 維護報告

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

---

## 📋 目錄

- [功能概述](#-功能概述)
- [支援設備清單](#-支援設備清單)
- [環境需求](#-環境需求)
- [安裝步驟](#-安裝步驟)
- [使用說明](#-使用說明)
- [輸出報表說明](#-輸出報表說明)
- [Log 檔案準備方式](#-log-檔案準備方式)
- [常見問題](#-常見問題)

---

## ✨ 功能概述

此工具為 Python Tkinter 圖形介面應用程式，專為網路維護工程師設計。

**核心功能：**
- 📂 批次掃描資料夾中所有 Log / Conf 檔案
- 🔍 自動識別設備廠牌與型號，無需手動分類
- 📊 一鍵輸出 Excel 維護報告（4 個分頁）
- 🌓 介面支援 Light / Dark 主題切換

---

## 🖥 支援設備清單

| 廠牌 | 系列 | 作業系統 | Log 副檔名 |
|------|------|----------|------------|
| **Cisco** | Catalyst 9200 / 9300 / 9500 | IOS-XE | `.txt` `.log` |
| **Cisco** | Catalyst 3850 / 2960 / 2960X | IOS | `.txt` `.log` |
| **Cisco** | Catalyst 9800 WLC | IOS-XE | `.txt` `.log` |
| **Cisco** | WLC 2504 / 3504 / 5520 | AireOS | `.txt` `.log` |
| **Cisco** | Nexus 9000 / 7000 | NX-OS | `.txt` `.log` |
| **Cisco** | Catalyst 1300 | ROS | `.txt` `.log` |
| **HPE Aruba** | 6300 / 6400 / 8400 | AOS-CX | `.txt` `.log` |
| **Fortinet** | FortiGate 70G / 100F / 200F… | FortiOS | `.conf` `.txt` |

---

## 🔧 環境需求

| 項目 | 版本需求 |
|------|----------|
| Python | **3.8 或以上** |
| 作業系統 | Windows 10 / 11（建議）<br>macOS / Linux 可執行，但 `Microsoft JhengHei` 字型可能須額外安裝 |
| 螢幕解析度 | 建議 1920 × 1200（程式會依此比例自動縮放） |

### Python 套件

```
openpyxl
pandas
```

> ⚠️ `tkinter` 為 Python 標準函式庫，Windows 安裝 Python 時預設已包含，無需另行安裝。

---

## 📦 安裝步驟

### 1. 安裝 Python

前往 [python.org](https://www.python.org/downloads/) 下載並安裝 Python 3.8+。

> 安裝時請勾選 **「Add Python to PATH」**

### 2. 安裝相依套件

開啟命令提示字元（CMD）或 PowerShell，執行：

```bash
pip install pandas openpyxl
```

### 3. 下載程式

```bash
# 方法一：Git clone
git clone https://github.com/你的帳號/network-log-analysis.git
cd network-log-analysis

# 方法二：直接下載 ZIP
# 點擊 GitHub 頁面右上角 Code → Download ZIP，解壓縮後使用
```

### 4. 準備圖示檔（選用）

將 `icon_ironman.ico` 放在與 `Cisco_Log_Analysis_V24.py` 相同的資料夾中，即可顯示自訂視窗圖示。

---

## 🚀 使用說明

### 啟動程式

```bash
python Cisco_Log_Analysis_V24.py
```

或在 Windows 直接雙擊 `.py` 檔（需已關聯 Python 直譯器）。

---

### 操作流程

```
步驟 1：準備 Log 資料夾
         └─ 將所有設備 Log（.txt / .log / .conf）放入同一資料夾

步驟 2：選擇資料夾
         └─ 點擊右側 🗁 按鈕，選取 Log 所在資料夾

步驟 3：產生報表
         └─ 點擊「產生分析報表」按鈕
         └─ 程式自動解析所有檔案（Log 視窗即時顯示進度）

步驟 4：儲存 Excel
         └─ 解析完成後，選擇 Excel 儲存路徑
         └─ 報表命名格式：Network_Analysis_MMDD_HHMM.xlsx
```

### 介面說明

```
┌─────────────────────────────────────────┐ Dark ◯
│ ≡  Network Analysis Tool by TING        │─────┤
├─────────────────────────────────────────┤     │
│ 選擇資料夾 [___路徑___________________] │ 🗁  │
├─────────────────────────────────────────┴─────┤
│                                               │
│  [  Log 輸出視窗（即時顯示解析進度）      ]  │
│                                               │
├───────────────────────────────────────────────┤
│  ████████████░░░░  進度條                     │
├───────────────────────────────────────────────┤
│                 ╔══════════════╗              │
│                 ║  產生分析報表 ║              │
│                 ╚══════════════╝              │
└───────────────────────────────────────────────┘
```

| 元件 | 功能 |
|------|------|
| `Dark` 滑桿 | 切換深色 / 淺色主題 |
| 🗁 按鈕 | 選擇含 Log 的資料夾 |
| Log 視窗 | 即時顯示解析進度與設備識別結果 |
| 進度條 | 顯示整體解析進度（% 完成） |
| 產生分析報表 | 開始解析並輸出 Excel |

---

## 📊 輸出報表說明

Excel 報表包含 **4 個分頁**：

### 📄 Report（總覽）
所有設備 + AP 的彙整清單，**不含 Ethernet MAC**。

| 欄位 | 說明 |
|------|------|
| Hostname | 設備主機名稱 |
| Model | 硬體型號 |
| OOB IP | 管理 IP |
| Serial Number | 序號 |
| Firmware Version | 軟體版本 |
| CPU Usage | CPU 使用率 |
| Memory Usage | 記憶體使用狀況 |
| Fan Status | 風扇狀態（正常 / 異常） |
| Power Status | 電源狀態（正常 / 異常） |
| System Log | 重要 Syslog 訊息摘要 |

### 📡 AP Info
WLC 旗下所有 AP 的詳細資訊（C9800 IOS-XE / AireOS）。

| 欄位 | 說明 |
|------|------|
| AP Name | AP 名稱 |
| AP Model | AP 型號 |
| Ethernet MAC | AP 實體 MAC |
| IP Address | AP IP 位址 |
| AP Serial Number | AP 序號 |
| Software Version | AP 韌體版本 |

### 🔀 WLC Info
Cisco C9800 及 AireOS WLC 的詳細清單，**含 Ethernet MAC**。

### 🔌 Switch Info
Cisco Switch（含 Stack）、NX-OS、C1300、Aruba、FortiGate 的詳細清單，**含 Ethernet MAC**。

> Stack 設備會顯示多列，每列對應一台實體成員，各自列出其 MAC 位址（以序號比對確保正確）。

---

## 📁 Log 檔案準備方式

### Cisco IOS / IOS-XE Switch

在設備上執行以下指令，將輸出複製貼上儲存為 `.txt`：

```
show version
show inventory
show environment all
show processes cpu sorted
show processes memory sorted
show ip interface brief
show logging
```

### Cisco C9800 WLC

```
show version
show inventory
show environment all
show ap summary
show ap config general
show wireless client summary
```

### Cisco AireOS WLC

```
show sysinfo
show inventory
show memory statistics
show ap join stats summary all
show ap config general <ap-name>
```

### Cisco Nexus (NX-OS)

```
show version
show inventory
show environment
show system resources
show vpc role
show interface mgmt0
show logging last 100
```

### HPE Aruba AOS-CX

```
show system
show version
show environment
show interface brief
show logging
```

### Fortinet FortiGate

直接匯出設備備份設定檔（`.conf`）：

```
GUI：Dashboard → System → Backup
CLI：execute backup config tftp <filename> <tftp-server>
```

若需 Fan / PSU / CPU / Memory 資訊，額外附加以下 CLI 輸出：

```
execute sensor list
execute tac report
get system performance status
```

---

## ❓ 常見問題

**Q：執行後出現 `ModuleNotFoundError: No module named 'pandas'`**

```bash
pip install pandas openpyxl
```

---

**Q：選擇資料夾後按「產生分析報表」沒有反應**

確認資料夾中有 `.txt`、`.log` 或 `.conf` 檔案，且至少一個檔案可被識別為支援的設備類型。

---

**Q：某台設備的欄位顯示 N/A**

Log 檔案中缺少對應指令的輸出。請參考 [Log 檔案準備方式](#-log-檔案準備方式) 補充指令輸出。

---

**Q：FortiGate 的 Fan / PSU / CPU / Memory 全部 N/A**

`.conf` 純備份檔不包含 `execute sensor list` 和 `get system performance status` 的輸出。請將這些指令的 CLI 輸出貼附在 `.conf` 檔案末端，或另存為獨立 `.txt` 放入同一資料夾。

---

**Q：Stack Switch 的 Ethernet MAC 不正確**

程式以各成員的 **System Serial Number** 進行比對，確保每台 Stack 成員的 MAC 正確對應。若 Serial Number 在 Log 中格式異常，可能導致 MAC 無法比對，請確認 `show version` 輸出完整。

---

## 📝 版本紀錄

| 版本 | 主要更新 |
|------|----------|
| V24 | 新增 FortigateParser（.conf 支援）；NexusParser vPC MAC；Aruba HPE MAC 格式正規化 |
| V23 | Stack 各成員 MAC 以序號比對；Aruba MAC HPE 格式支援；視窗螢幕比例自適應 |
| V22 | 新增 WLC Info / Switch Info 分頁；DeviceInfo 新增 eth_mac |
| V21 | AireOS AP 版本繼承 WLC 版本；Layout 重構防止按鈕消失 |
| V20 | 中文字型統一 Microsoft JhengHei；Log 批次寫入效能優化 |
| V19 | 圓角置中按鈕；Light/Dark 資料夾按鈕邊框；Iron Man 圖示 |
| V18 | Dark Mode 切換 Slider；UI 全面重構（工具列 / 分隔線 / 扁平化）|

---

## 👤 作者

**TING CHEN** — Network Engineer  
若有問題或建議，歡迎開 [Issue](https://github.com/你的帳號/network-log-analysis/issues) 回報。
