"""
解析器回歸測試（只使用 tests/fixtures 的合成 Log，不含任何客戶資料）。

執行：python -m unittest discover -s tests -v
"""
import importlib.util
import os
import tempfile
import unittest

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIXTURES = os.path.join(ROOT, "tests", "fixtures")

_spec = importlib.util.spec_from_file_location("cla", os.path.join(ROOT, "Cisco_Log_Analysis.py"))
cla = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(cla)


def analyze(*names):
    """只分析指定的 fixture 檔（放到暫存資料夾，模擬使用者的 Log 資料夾）。"""
    with tempfile.TemporaryDirectory() as d:
        for n in names:
            with open(os.path.join(FIXTURES, n), encoding="utf-8") as src, \
                 open(os.path.join(d, n), "w", encoding="utf-8") as dst:
                dst.write(src.read())
        devices, aps, results = cla.analyze_folder(d, log=lambda _m: None)
    return devices, aps, results


class TestCommandIndex(unittest.TestCase):
    def test_cisco_abbreviation_and_errors(self):
        log = ("SW#sh ver\nCisco IOS Software, Version 17.9\n"
               "SW#show in\n% Ambiguous command: \"show in\"\n"
               "SW#show inventory\nNAME: \"1\"\n"
               "SW#show processes cpu | i CPU\nCPU utilization for five seconds: 1%/0%; five minutes: 3%\n"
               "SW#end\n")
        idx = cla.build_prompt_index(log)
        self.assertTrue(idx.get("show version").startswith("Cisco IOS"))
        self.assertEqual(idx.get("show inventory"), 'NAME: "1"')          # 跳過錯誤輸出
        self.assertIn("five minutes: 3%", idx.get("show processes cpu"))  # 管線過濾版本備援
        self.assertEqual(idx.get("show interfaces"), "")

    def test_pasted_command_list_is_ignored(self):
        log = "N9K# terminal length 0\n!\nshow version\nshow inventory\nN9K# show version\nNXOS: version 10.3(6)\nN9K# exit\n"
        self.assertIn("NXOS", cla.build_prompt_index(log).get("show version"))


class TestC9800HA(unittest.TestCase):
    def test_ha_pair_rows(self):
        devices, aps, _ = analyze("10.0.0.10@WLC-HA.log")
        self.assertEqual(len(devices), 2)
        active, standby = devices
        self.assertTrue(active["HA Role"].startswith("Active (Chassis 2)"))
        self.assertEqual(active["Serial Number"], "FCLTEST0002")
        self.assertEqual(active["OOB IP"], "10.0.0.10")                    # WMI primary，不是 RMI secondary
        self.assertIn("15838712K", active["Memory Usage"])                 # platform sorted 優先
        self.assertEqual(standby["Hostname"], "○ (HA)")
        self.assertEqual(standby["Serial Number"], "FCLTEST0001")
        self.assertEqual(len(aps), 1)
        self.assertEqual(aps[0]["source"], "10.0.0.10@WLC-HA.log")


class TestFortiGate(unittest.TestCase):
    def test_ha_cluster_merged_by_serial(self):
        devices, _, results = analyze("10.1.1.1@FGT-HQ-primary.log", "10.1.1.2@FGT-HQ-secondary.log")
        self.assertTrue(all(r.dev_type == "FortiGate" for r in results))
        self.assertEqual(len(devices), 2)                                  # 兩份 Log、兩台、不重複
        primary, secondary = devices
        self.assertEqual((primary["HA Role"], primary["Serial Number"]), ("Primary", "FG100FTK19000001"))
        self.assertEqual((secondary["HA Role"], secondary["Serial Number"]), ("Secondary", "FG100FTK19000002"))
        self.assertEqual(primary["Model"], "FortiGate-100F")
        self.assertEqual(primary["Firmware Version"], "v7.2.8 build1639")
        self.assertEqual(primary["OOB IP"], "10.1.1.1")
        self.assertEqual(primary["CPU Usage"], "4%")                       # 100 − 96% idle
        self.assertEqual(primary["Power Status"], "異常")                  # PS2 VIN alarm=1
        self.assertEqual(secondary["Power Status"], "正常")                # 取自 Secondary 自己的 Log
        self.assertEqual(primary["HA Status"], "OK｜in-sync")
        self.assertEqual(secondary["CPU Usage"], "1%")
        log = primary["System Log"].splitlines()
        self.assertEqual(log, ["2024-05-01 14:59:00 [warning] Power supply failure: Power supply 2 failure"])

    def test_backup_conf_only(self):
        devices, _, _ = analyze("FGT-BRANCH.conf")
        r = devices[0]
        self.assertEqual((r["Hostname"], r["Model"], r["Firmware Version"]), ("FGT-BRANCH", "FG60F", "v7.0.15 build0632"))
        self.assertEqual(r["OOB IP"], "192.168.10.1")
        self.assertEqual((r["CPU Usage"], r["Fan Status"], r["Power Status"]), ("N/A", "N/A", "N/A"))


class TestAruba(unittest.TestCase):
    def test_vsf_members_and_health(self):
        devices, _, _ = analyze("aruba_vsf_show_tech.txt")
        self.assertEqual([r["VSF Role"] for r in devices], ["Conductor", "Standby", "Not Present"])
        conductor, standby, _missing = devices
        self.assertEqual(conductor["Serial Number"], "SG00TEST01")
        self.assertEqual(standby["Serial Number"], "SG00TEST02")
        self.assertEqual(conductor["OOB IP"], "10.9.9.9")
        self.assertEqual(conductor["Fan Status"], "異常")
        self.assertEqual(conductor["Ports Up/Total"], "2 / 3")
        self.assertEqual(conductor["Core Dumps"], "1")
        self.assertEqual((conductor["Primary Image"], conductor["Secondary Image"]), ("FL.10.16.1020", "FL.10.13.1040"))
        self.assertIn("2/1-CPU", conductor["Max Temperature"])
        self.assertTrue(conductor["System Log"].startswith("2026-04-17T06:26:50"))   # 同訊息保留最後一次


class TestExcel(unittest.TestCase):
    def _sheets(self, *names):
        devices, aps, _ = analyze(*names)
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "out.xlsx")
            cla.write_excel(path, devices, aps)
            wb = load_workbook(path)
            return {ws.title: [c.value for c in ws[1]] for ws in wb.worksheets}

    def test_vendor_sheets_created_only_when_detected(self):
        sheets = self._sheets("10.0.0.10@WLC-HA.log")
        self.assertEqual(list(sheets), ["Report", "AP Info", "WLC Info", "Switch Info"])
        sheets = self._sheets("aruba_vsf_show_tech.txt", "FGT-BRANCH.conf")
        self.assertIn("Aruba Info", sheets)
        self.assertIn("Fortigate Info", sheets)
        for name in ("AP Info", "WLC Info", "Switch Info", "Aruba Info", "Fortigate Info"):
            self.assertEqual(sheets[name][-1], "Source File")


if __name__ == "__main__":
    unittest.main()
