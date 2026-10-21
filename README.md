# 🌐 Network Log Analysis Tool

> 自動解析 Cisco / Aruba / FortiGate 設備 Log，一鍵產生 Excel 維護報告

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)
![Version](https://img.shields.io/badge/Version-V21-green)

---

## 目錄

- [支援設備](#支援設備)
- [快速開始](#快速開始)
- [輸出報表](#輸出報表)
- [Log 收集指令](#log-收集指令)
- [開發與版本控制](#開發與版本控制)

---

## 支援設備

| 廠牌 | 系列 | OS | 檔案 |
|------|------|----|------|
| Cisco | Catalyst 9200 / 9300 / 9500 / 3850 / 2960 / 1000（含 Stack） | IOS / IOS-XE | `.log` `.txt` |
| Cisco | Catalyst 9800 WLC / EWC（含 **HA SSO**） | IOS-XE | `.log` `.txt` |
| Cisco | WLC 2504 / 3504 / 5520 | AireOS | `.log` `.txt` |
| Cisco | Nexus 5000 / 9000 | NX-OS | `.log` `.txt` |
| Cisco | Catalyst 1300 | show tech-support | `.log` `.txt` |
| HPE Aruba | 6300 / 6400（含 **VSF**） | AOS-CX show tech | `.log` `.txt` |
| Fortinet | FortiGate（含 **HA 叢集**） | FortiOS | `.conf` `.log` `.txt` |

設備類型依 Log 內容自動判斷，不依檔名。

---

## 快速開始

### 方式一：直接執行 exe（一般使用者，免安裝 Python）

到 [Releases](https://github.com/tingchen0329/Log_Analysis_Tool/releases) 下載 `Log_Analysis_Tool.exe`（約 7~10 MB），雙擊執行即可。

### 方式二：Python 執行（開發者）

```bash
pip install -r requirements.txt          # 只需要 openpyxl
python Cisco_Log_Analysis.py                       # GUI
python Cisco_Log_Analysis.py <Log資料夾> -o 報表.xlsx  # CLI 批次模式
```

### 自行打包 exe

在 Windows 執行 `build_exe.bat`，產出 `dist\Log_Analysis_Tool.exe`。

---

## 輸出報表

| 分頁 | 內容 |
|------|------|
| Report | 所有設備 + AP 總覽（不含 MAC） |
| AP Info | WLC 底下的 AP + Source File |
| WLC Info | C9800 / AireOS，含 **HA Role**（Active / Standby Hot、Chassis、RMI）+ Source File |
| Switch Info | IOS / NX-OS / C1300 / Aruba + Source File |
| Aruba Info | 偵測到 Aruba 才建立：VSF Role、映像版本、Uptime、重開機原因、最高溫、Port 數、NTP、Core Dump |
| Fortigate Info | 偵測到 FortiGate 才建立：HA Role / HA Status、運作模式、Uptime、Sessions、重開機原因 |

- Stack / VSF / HA 每台實體設備一列；非首列以「○」表示與首列共用。
- HA 兩台的 Log 同時放入時（C9800 Active + `-stby`、FortiGate Primary + Secondary），會自動去重合併。
- System Log 保留時間戳記，同一訊息只顯示最後一次發生。

---

## Log 收集指令

### Cisco IOS / IOS-XE / C9800

```
show version
show inventory
show environment all
show processes cpu sorted
show processes memory platform sorted
show processes memory sorted
show running-config
show logging
! C9800 另需
show chassis
show redundancy
show redundancy states
show ap summary
show ap config general
```

### Cisco NX-OS

```
show version
show inventory
show environment
show system resources
show processes cpu
show interface mgmt0
show running-config
show logging
```

### Aruba AOS-CX

`show tech`（已包含所需的 show system / version / images / vsf / module / environment / interface brief / boot-history / ntp status / core-dump / events）。

### FortiGate

`.conf` 備份檔只有設定（Hostname、型號、版本、管理 IP），健康資訊需另外收集 CLI：

```
get system status
get system performance status
get system ha status
execute sensor list
execute log display
show system interface
```

---

## 開發與版本控制

```
Cisco_Log_Analysis.py        主程式（單一檔案）
tests/test_parsers.py        單元測試
tests/fixtures/              合成測試 Log（不含客戶資料）
build_exe.bat                Windows 打包腳本
push_to_github.bat           一鍵 commit / tag / 推送
.github/workflows/build.yml  CI：push 跑測試；推送 v* tag 自動打包 exe 並建立 Release
```

```bash
python -m unittest discover -s tests -v     # 執行測試

git tag -a v22 -m v22 && git push origin main --follow-tags   # 發佈新版本（GitHub Actions 自動產出 exe）
```

> ⚠️ 本 Repo 為公開專案，`.gitignore` 已排除 `*.log` / `*.conf` / `*.txt` / `*.xlsx`，請勿提交客戶 Log 或報表。

版本紀錄見 [CHANGELOG.md](CHANGELOG.md)。

---

**TING CHEN** — Network Engineer
