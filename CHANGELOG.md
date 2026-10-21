# 版本紀錄

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
