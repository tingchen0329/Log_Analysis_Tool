"""
維護設備比對的合成情境（假資料：文件保留 IP、虛構序號），上季 → 本季的變化：

  CORE-STK（Stack 3 台 → 4 台）：成員 2 換新、新增成員 4、成員 1 / 3 韌體 17.9.4 → 17.12.3
  ACC-SW02：下線（移除設備）
  ACC-SW05、FGT-BRANCH2：新上線（新增設備）
  ACC-SW03：改名為 ACC-SW03-NEW（同序號）
  N9K-LEAF：Power Status 正常 → 異常、System Log 新增 PS_FAIL 告警
  FGT-HQ：OOB IP 變更
  EDGE-SW：兩季都沒有序號（需人工確認）
  WLC-HA（HA 兩台）：無變化
  AP：AP-1F-02 改名、AP-2F-01 換機（同名新序號）、AP-OLD 下線、AP-3F-07 / AP-3F-08 新上線

產生範例檔（tests/fixtures/compare_2026Q2.xlsx、compare_2026Q3.xlsx）：
    python tests/compare_scenario.py
"""
import os


def _dev(dev_type, host, sn, model, ip="○", fw="17.9.4", fan="正常", pwr="正常", log="",
         src="", mac="", ha=""):
    row = {"_device_type": dev_type, "Hostname": host, "Model": model, "OOB IP": ip,
           "Ethernet MAC": mac, "Serial Number": sn, "Firmware Version": fw,
           "CPU Usage": "5%", "Memory Usage": "40%", "Fan Status": fan, "Power Status": pwr,
           "System Log": log, "Source File": src}
    if ha:
        row["HA Role"] = ha
    return row


def _ap(name, sn, model="C9120AXI", ip="", mac="", src="192.0.2.50@WLC5520.log"):
    return {"ap_name": name, "ap_model": model, "eth_mac": mac, "ip_address": ip,
            "serial": sn, "firmware": "8.10.185.0", "source": src}


def quarter(q: str):
    """回傳（devices, aps），可直接交給 write_excel()。q = "Q2" 或 "Q3"。"""
    q3 = q == "Q3"
    fw = "17.12.3" if q3 else "17.9.4"
    stk = "192.0.2.10@CORE-STK.log"
    devices = [
        _dev("Switch", "CORE-STK", "FOC0000SK01", "C9300-48P", "192.0.2.10", fw,
             log="Jun 10 08:00:01: %LINK-3-UPDOWN: Interface Gi1/0/1, changed state to down", src=stk),
        _dev("Switch", "○ (Stack)", "FOC0000SK05" if q3 else "FOC0000SK02", "C9300-48P",
             fw="17.12.3" if q3 else "17.9.4", src=stk),
        _dev("Switch", "○ (Stack)", "FOC0000SK03", "C9300-24P", fw=fw, src=stk),
    ]
    if q3:
        devices.append(_dev("Switch", "○ (Stack)", "FOC0000SK06", "C9300-24P", fw=fw, src=stk))
    else:
        devices.append(_dev("Switch", "ACC-SW02", "FOC0000AS02", "C2960X-48FPD-L", "192.0.2.12",
                            "15.2(7)E8", src="192.0.2.12@ACC-SW02.log"))
    devices.append(_dev("Switch", "ACC-SW03-NEW" if q3 else "ACC-SW03", "FOC0000AS03", "C9200L-24P-4G",
                        "192.0.2.13", "17.9.4", src="192.0.2.13@ACC-SW03.log"))
    if q3:
        devices.append(_dev("Switch", "ACC-SW05", "FOC0000AS05", "C9200L-48P-4G", "192.0.2.61",
                            "17.12.3", src="192.0.2.61@ACC-SW05.log"))
    devices.append(_dev("Switch", "EDGE-SW", "N/A", "C1000-8T-2G-L", "192.0.2.70", "15.2(7)E9",
                        src="192.0.2.70@EDGE-SW.log"))
    n9k_log = "2026 Jun 10 08:00:01 N9K-LEAF %PLATFORM-2-PS_FAIL: Power supply 2 failed or shut down"
    devices.append(_dev("NX-OS", "N9K-LEAF", "FDO0000NX01", "N9K-C93180YC-FX", "198.51.100.20", "10.3(4a)",
                        pwr="異常" if q3 else "正常", log=n9k_log if q3 else "",
                        src="198.51.100.20@N9K-LEAF.log"))
    devices += [
        _dev("WLC", "WLC-HA", "FCL0000HA01", "C9800-L-C-K9", "192.0.2.40", "17.9.5",
             src="192.0.2.40@WLC-HA.log", ha="Active"),
        _dev("WLC", "○ (HA)", "FCL0000HA02", "C9800-L-C-K9", "○", "17.9.5",
             src="192.0.2.40@WLC-HA.log", ha="Standby Hot"),
        _dev("FortiGate", "FGT-HQ", "FG100FTK0000001", "FortiGate-100F",
             "198.51.100.2" if q3 else "198.51.100.1", "v7.2.8", src="FGT-HQ.conf"),
    ]
    if q3:
        devices.append(_dev("FortiGate", "FGT-BRANCH2", "FG100FTK0000030", "FortiGate-100F",
                            "198.51.100.30", "v7.2.8", src="FGT-BRANCH2.conf"))

    aps = [
        _ap("AP-1F-01", "FJC0000AP01", ip="192.0.2.101"),
        _ap("AP-1F-02B" if q3 else "AP-1F-02", "FJC0000AP02", ip="192.0.2.102"),
        _ap("AP-2F-01", "FJC0000AP09" if q3 else "FJC0000AP03", ip="192.0.2.103"),
    ]
    aps += [_ap("AP-3F-07", "FJC0000AP07", ip="192.0.2.107"),
            _ap("AP-3F-08", "FJC0000AP08", ip="192.0.2.108")] if q3 else \
           [_ap("AP-OLD", "FJC0000AP04", model="AIR-AP2802I-E-K9", ip="192.0.2.104")]
    return devices, aps


if __name__ == "__main__":                        # 產生範例檔
    import importlib.util
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    spec = importlib.util.spec_from_file_location("cla", os.path.join(root, "Cisco_Log_Analysis.py"))
    cla = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(cla)
    for q in ("Q2", "Q3"):
        out = os.path.join(root, "tests", "fixtures", f"compare_2026{q}.xlsx")
        cla.write_excel(out, *quarter(q))
        print(f"已產生 {out}")
