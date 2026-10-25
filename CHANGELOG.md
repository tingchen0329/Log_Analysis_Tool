# 版本紀錄

> 註：V19 起重新整理版號。2026/05 舊版 README 內的「V20～V24」是舊編號，與下列版本內容不同。

## V23

新功能：維護設備比對（上季 / 本季兩份維護報表 → 差異報表）

- 以 Serial Number 配對；Stack / HA 的「○」先補成首列的值再比對
- 判定：新增設備、移除設備、換機、Stack 成員換新 / 新增 / 移除、HA 機箱換新、AP 換機、改名
- 欄位異動：Model、OOB IP、Firmware、HA Role、Fan / Power Status、System Log 新增告警；同群組相同異動合併成一列
- 沒有序號或序號重複時，改用 Hostname + Model 配對並列入「需人工確認」
- 比對報表分頁：摘要、新增設備、移除設備、換機與成員變更、欄位異動、新增 AP、移除 AP、需人工確認（無差異寫「無」）
- GUI 新增「分析 Log / 維護設備比對」切換；網頁版新增標籤頁（網址 `#compare`）；CLI 新增 `--compare`、`--compare-with`
- 新增合成兩季範例報表 `tests/fixtures/compare_2026Q2/Q3.xlsx` 與 18 個比對測試（29 → 47 個測試）；網頁版煙霧測試涵蓋比對

## V22

依 Cisco / HPE Aruba 官方文件的指令輸出與狀態值比對，修正環境狀態漏判：

- 環境異常關鍵字改為字首比對：Failure / Failed / Faulty 也算異常（V21 只認完整單字 FAIL / FAULT）
- NX-OS：電源 `Shutdown` 判為異常、`Fan_in_PSn` 歸類為電源；修正 N3K 風扇列 `Fan-1 …` 被誤認為區段標頭而略過
- AireOS：電源改看 show sysinfo 的 `Power Supply N ... Present, OK`（V21 只看溫度）；sysinfo 沒有 CPU 欄位時改取 `show cpu`
- Aruba：PSU `Warning` 判為異常、`Absent`（插槽未安裝）不再判為異常；風扇同時檢查 Fan tray
- 新增 IOS Stack、NX-OS、AireOS、C1300 合成測試資料與單元測試（7 → 29 個測試）
- 新增網頁版（GitHub Pages + Pyodide，瀏覽器內解析、Log 不上傳），push 到 main 自動部署
- 新增 `tests/compare_versions.py`：新舊版本以真實 Log 逐欄比對
- README 整理為三種使用方式：exe、Python 原始碼、網頁版

## V21

- 移除 pandas 依賴，Excel 改由 openpyxl 直接寫入（輸出逐格與 V20 相同），exe 由約 60 MB 降到約 7~10 MB、啟動更快
- 程式註解中的範例資料全部改為假資料；Git 歷史中的 V18 同步去識別化
- 新增 push_to_github.bat 一鍵推送

## V20

- 新增 FortiGate 解析器：支援 `.conf` 備份檔與 CLI 輸出；HA 叢集兩台各一列，依序號跨檔去重
- Aruba 加強：VSF 堆疊成員各一列、Active / Primary / Secondary 映像、Uptime、上次重開機原因、最高溫、Port Up/Total、NTP、Core Dump
- Excel：偵測到 Aruba / FortiGate 時自動新增「Aruba Info」「Fortigate Info」分頁
- 新增單元測試（合成 fixtures）、PyInstaller 打包腳本、GitHub Actions 自動建置 exe

## V19

- C9800 HA SSO 判定：兩台機箱各一列、HA Role 欄、Active / Standby Log 自動合併、OOB IP 改取 WMI
- 架構重構：CommandIndex 一次掃描建索引、支援 Cisco 指令縮寫、解析引擎與 GUI 分離（可 CLI 執行）
- Excel：AP Info / WLC Info / Switch Info 最後一欄新增 Source File
- Catalyst Memory 優先序：`show processes memory platform sorted` → `sorted` → `| i Pool Total`
- System Log 保留時間戳記；單格限制 32,000 字元以符合 Excel 上限
- 修正：AireOS 誤判、縮寫環境指令讀不到（PSU 故障被顯示為正常）、NX-OS MAC、跨介面誤抓 IP、Tk 執行緒安全

## V18

- NX-OS 判定修正；WLC CPU 指令單複數相容
