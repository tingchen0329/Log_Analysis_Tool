# 🌐 Network Log Analysis Tool

> 自動解析 Cisco / Aruba / FortiGate 設備 Log，一鍵產生 Excel 維護報告

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Web-lightgrey)
![Version](https://img.shields.io/badge/Version-V22-green)

---

## 目錄

- [三種使用方式](#三種使用方式)
  - [方式一：下載 exe，離線使用](#方式一下載-exe離線使用)
  - [方式二：下載 Python 原始碼，離線使用](#方式二下載-python-原始碼離線使用)
  - [方式三：網頁版](#方式三網頁版)
- [支援設備](#支援設備)
- [輸出報表](#輸出報表)
- [Log 收集指令](#log-收集指令)
- [開發與版本控制](#開發與版本控制)

---

## 三種使用方式

三種方式共用**同一份**解析引擎 `Cisco_Log_Analysis.py`，同樣的 Log 產生的 Excel 內容完全相同。

| | 方式一：exe | 方式二：Python 原始碼 | 方式三：網頁版 |
|---|---|---|---|
| 適合 | 一般使用者 | 開發者、需要 CLI 批次處理 | 不能執行 exe 的電腦、臨時使用 |
| 需要安裝 | 不需要 | Python 3.8+ 與 openpyxl | 不需要（Chrome / Edge 瀏覽器） |
| 需要網路 | 只有下載時 | 只有安裝套件時 | 開啟網頁時（之後使用瀏覽器快取） |
| Log 是否離開電腦 | 否 | 否 | 否（在瀏覽器內解析，不上傳） |
| 操作 | GUI | GUI 或 CLI | 網頁拖放 |

### 方式一：下載 exe，離線使用

1. 到 [Releases](https://github.com/tingchen0329/Log_Analysis_Tool/releases) 下載最新版的 `Log_Analysis_Tool.exe`（約 7~10 MB）。
2. 把 exe 放在任何資料夾，雙擊執行（不需要安裝 Python，也不需要網路）。
3. 選擇 Log 資料夾 → 開始分析 → 選擇 Excel 存檔位置。

> 💡 第一次執行時 Windows SmartScreen 可能顯示「已保護您的電腦」：點「其他資訊 → 仍要執行」。
> 若公司電腦的 AppLocker / 應用程式控制擋住 exe，請改用[方式三：網頁版](#方式三網頁版)。

### 方式二：下載 Python 原始碼，離線使用

**1. 安裝 Python**

到 [python.org](https://www.python.org/downloads/) 下載 Python 3.8 以上版本。Windows 安裝時請勾選 **Add python.exe to PATH**。

**2. 下載原始碼**

在 Repo 頁面按 **Code → Download ZIP** 後解壓縮，或使用 Git：

```bash
git clone https://github.com/tingchen0329/Log_Analysis_Tool.git
cd Log_Analysis_Tool
```

**3. 安裝套件**（只需要 openpyxl）

```bash
pip install -r requirements.txt
```

<details>
<summary>完全沒有網路的電腦（離線安裝）</summary>

在可以上網的電腦下載套件，再把 `packages` 資料夾複製過去：

```bash
# 可上網的電腦
pip download -r requirements.txt -d packages

# 離線電腦
pip install --no-index --find-links packages -r requirements.txt
```

</details>

**4. 執行**

```bash
python Cisco_Log_Analysis.py                          # GUI（Windows 可用 pythonw 不顯示黑視窗）
python Cisco_Log_Analysis.py <Log資料夾> -o 報表.xlsx    # CLI 批次模式（不需要 tkinter）
```

**自行打包 exe**：在 Windows 執行 `build_exe.bat`，產出 `dist\Log_Analysis_Tool.exe`。

### 方式三：網頁版

👉 **https://tingchen0329.github.io/Log_Analysis_Tool/**

1. 開啟網頁，等右上角顯示「就緒（V22）」。
2. 把 Log 資料夾或檔案拖曳到網頁（或按「選擇資料夾」「選擇檔案」，會一併讀取子資料夾）。
3. 按「開始分析」→「下載 Excel 報表」。

- **Log 不會上傳**：網頁以 [Pyodide](https://pyodide.org/)（WebAssembly 版 Python）在瀏覽器內執行與 exe 相同的程式，Log 只存在瀏覽器記憶體中，分析完即清除。
- **全部自行託管**：Pyodide 與 openpyxl 都放在 GitHub Pages 上，不連任何外部 CDN（公司網路擋 CDN 也能用）。
- 第一次開啟需下載約 14 MB 的執行環境，之後由瀏覽器快取。
- 支援瀏覽器：Chrome / Edge 80 以上、Firefox 114 以上、Safari 15 以上（需支援 Module Web Worker）。

<details>
<summary>在自己的電腦預覽網頁版</summary>

```bash
python web/build_site.py _site                 # 下載 Pyodide 與 wheel，組出網站
python -m http.server -d _site 8000            # 開啟 http://localhost:8000/
```

不能直接雙擊 `index.html`：瀏覽器不允許 `file://` 頁面啟動 Web Worker。

</details>

---

## 支援設備

| 廠牌 | 系列 | OS | 檔案 |
|------|------|----|------|
| Cisco | Catalyst 9200 / 9300 / 9500 / 3850 / 2960 / 1000（含 Stack） | IOS / IOS-XE | `.log` `.txt` |
| Cisco | Catalyst 9800 WLC / EWC（含 **HA SSO**） | IOS-XE | `.log` `.txt` |
| Cisco | WLC 2504 / 3504 / 5520 / 8540 | AireOS | `.log` `.txt` |
| Cisco | Nexus 3000 / 5000 / 9000 | NX-OS | `.log` `.txt` |
| Cisco | Catalyst 1300（含 Stack） | show tech-support | `.log` `.txt` |
| HPE Aruba | 6300 / 6400（含 **VSF**） | AOS-CX show tech | `.log` `.txt` |
| Fortinet | FortiGate（含 **HA 叢集**） | FortiOS | `.conf` `.log` `.txt` |

設備類型依 Log 內容自動判斷，不依檔名。

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
- Fan / Power Status 依各廠牌官方狀態值判斷；未安裝的插槽（Cisco `Not Present`、Aruba `Absent`、AireOS `Absent`）不算異常。

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

### Cisco AireOS

```
config paging disable
show sysinfo
show cpu
show memory statistics
show inventory
show logging
show ap join stats summary all
show ap inventory all
```

### Cisco NX-OS

```
show version
show inventory
show environment
show system resources
show processes cpu
show vpc role
show interface mgmt0
show running-config
show logging
```

### Cisco Catalyst 1300

`show tech-support`（已包含 show system / version / inventory / stack / cpu utilization / ip interface / logging）。

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
Cisco_Log_Analysis.py          主程式（單一檔案，GUI / CLI / 網頁版共用）
tests/test_parsers.py          單元測試
tests/fixtures/                合成測試 Log（不含客戶資料）
tests/compare_versions.py      新舊版本回歸比對（用真實 Log 在本機執行）
web/index.html                 網頁版介面
web/worker.js                  網頁版 Module Web Worker（Pyodide 執行 Python）
web/build_site.py              組出網頁版網站（Pyodide + wheel 自行託管）
web/smoke_test.cjs             網頁版煙霧測試（Node 模擬瀏覽器 Worker）
web/requirements-web.txt       網頁版使用的 wheel
build_exe.bat                  Windows 打包腳本
push_to_github.bat             一鍵 commit / tag / 推送
.github/workflows/build.yml    CI：push 跑測試；推送 v* tag 自動打包 exe 並建立 Release
.github/workflows/pages.yml    CI：push 到 main 自動部署網頁版到 GitHub Pages
```

```bash
python -m unittest discover -s tests -v                       # 單元測試
python tests/compare_versions.py 舊版.py Cisco_Log_Analysis.py <真實Log資料夾>   # 改版回歸比對

python web/build_site.py _site                                # 組出網頁版
node --no-warnings web/smoke_test.cjs _site tests/fixtures    # 網頁版煙霧測試

git tag -a v22 -m v22 && git push origin main --follow-tags   # 發佈新版本（GitHub Actions 自動產出 exe）
```

- 網頁版第一次部署前，請到 Repo **Settings → Pages → Build and deployment → Source** 選 **GitHub Actions**。
- 網頁版不需要另外改程式：只要 `analyze_folder` / `write_excel` / `build_sheets` 的介面不變，push 到 main 就會自動更新。

> ⚠️ 本 Repo 為公開專案，`.gitignore` 已排除 `*.log` / `*.conf` / `*.txt` / `*.xlsx`，請勿提交客戶 Log 或報表。

版本紀錄見 [CHANGELOG.md](CHANGELOG.md)。

---

**TING CHEN** — Network Engineer
