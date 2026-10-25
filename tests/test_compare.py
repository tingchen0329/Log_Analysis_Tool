"""
維護設備比對測試（V23）：只使用 tests/compare_scenario.py 的合成資料與 tests/fixtures 的合成 Log。

執行：python -m unittest discover -s tests -v
"""
import contextlib
import importlib.util
import io
import os
import sys
import tempfile
import unittest

from openpyxl import Workbook, load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIXTURES = os.path.join(ROOT, "tests", "fixtures")
sys.path.insert(0, os.path.join(ROOT, "tests"))
import compare_scenario  # noqa: E402

_spec = importlib.util.spec_from_file_location("cla", os.path.join(ROOT, "Cisco_Log_Analysis.py"))
cla = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(cla)


def by(rows, **match):
    return [r for r in rows if all(r.get(k) == v for k, v in match.items())]


class TestCompareScenario(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.q2 = os.path.join(cls.tmp.name, "2026Q2.xlsx")
        cls.q3 = os.path.join(cls.tmp.name, "2026Q3.xlsx")
        cla.write_excel(cls.q2, *compare_scenario.quarter("Q2"))
        cla.write_excel(cls.q3, *compare_scenario.quarter("Q3"))
        cls.result = cla.compare_reports(cls.q2, cls.q3)
        cls.s = cls.result.sheets

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_counts(self):
        self.assertEqual(self.result.counts(), {
            "新增設備": 2, "移除設備": 1, "換機與成員變更": 3, "欄位異動": 6,
            "新增 AP": 2, "移除 AP": 1, "需人工確認": 1})

    def test_mark_rows_filled(self):
        stack = [d for d in self.result.new.devices if d.host == "CORE-STK"]
        self.assertEqual([d.pos for d in stack], ["成員 1", "成員 2", "成員 3", "成員 4"])
        self.assertTrue(all(d.values["OOB IP"] == "192.0.2.10" for d in stack))
        ha = [d for d in self.result.new.devices if d.host == "WLC-HA"]
        self.assertEqual([d.pos for d in ha], ["機箱 1", "機箱 2"])

    def test_added_and_removed(self):
        self.assertEqual({r["Hostname"] for r in self.s["新增設備"]}, {"ACC-SW05", "FGT-BRANCH2"})
        self.assertEqual(by(self.s["新增設備"], Hostname="FGT-BRANCH2")[0]["類型"], "FortiGate")
        self.assertEqual([r["Serial Number"] for r in self.s["移除設備"]], ["FOC0000AS02"])

    def test_stack_member_replaced_and_added(self):
        rows = by(self.s["換機與成員變更"], Hostname="CORE-STK")
        replaced = by(rows, 判定="成員換新")[0]
        self.assertEqual((replaced["位置"], replaced["上季 Serial"], replaced["本季 Serial"]),
                         ("成員 2", "FOC0000SK02", "FOC0000SK05"))
        added = by(rows, 判定="成員新增")[0]
        self.assertEqual((added["位置"], added["上季 Serial"], added["本季 Serial"]),
                         ("成員 4", "", "FOC0000SK06"))

    def test_rename_is_not_add(self):
        self.assertFalse(by(self.s["新增設備"], Hostname="ACC-SW03-NEW"))
        row = by(self.s["欄位異動"], 欄位="Hostname（改名）")[0]
        self.assertEqual((row["上季"], row["本季"]), ("ACC-SW03", "ACC-SW03-NEW"))

    def test_field_changes(self):
        fw = by(self.s["欄位異動"], Hostname="CORE-STK", 欄位="Firmware Version")
        self.assertEqual(len(fw), 1)                               # 成員 1、3 相同異動合併為一列
        self.assertEqual(fw[0]["位置"], "成員 1、成員 3")
        pwr = by(self.s["欄位異動"], Hostname="N9K-LEAF", 欄位="Power Status")[0]
        self.assertEqual((pwr["上季"], pwr["本季"]), ("正常", "異常"))
        alert = by(self.s["欄位異動"], Hostname="N9K-LEAF", 欄位="System Log（新增告警）")[0]
        self.assertIn("PS_FAIL", alert["本季"])
        ip = by(self.s["欄位異動"], Hostname="FGT-HQ", 欄位="OOB IP")[0]
        self.assertEqual((ip["上季"], ip["本季"]), ("198.51.100.1", "198.51.100.2"))
        self.assertFalse(by(self.s["欄位異動"], Hostname="WLC-HA"))  # 沒變化的 HA 不列出

    def test_aps(self):
        self.assertEqual({a["AP Name"] for a in self.s["新增 AP"]}, {"AP-3F-07", "AP-3F-08"})
        self.assertEqual([a["AP Name"] for a in self.s["移除 AP"]], ["AP-OLD"])
        swap = by(self.s["換機與成員變更"], 判定="AP 換機")[0]
        self.assertEqual((swap["Hostname"], swap["上季 Serial"], swap["本季 Serial"]),
                         ("AP-2F-01", "FJC0000AP03", "FJC0000AP09"))
        self.assertTrue(by(self.s["欄位異動"], 欄位="AP Name（改名）", 本季="AP-1F-02B"))

    def test_no_serial_needs_review(self):
        row = self.s["需人工確認"][0]
        self.assertEqual((row["Hostname"], row["季別"]), ("EDGE-SW", "兩季"))

    def test_compare_excel(self):
        out = os.path.join(self.tmp.name, "cmp.xlsx")
        cla.write_compare_excel(out, self.result)
        wb = load_workbook(out)
        self.assertEqual(wb.sheetnames, list(cla.COMPARE_SHEETS))
        summary = {r[0]: r[1] for r in wb["摘要"].iter_rows(min_row=2, values_only=True)}
        self.assertEqual((summary["上季設備數"], summary["本季設備數"], summary["新增設備"]), (10, 12, 2))
        self.assertEqual(wb["新增設備"].cell(1, 1).value, "類型")

    def test_reverse_direction(self):
        rev = cla.compare_reports(self.q3, self.q2).counts()
        self.assertEqual((rev["新增設備"], rev["移除設備"], rev["新增 AP"], rev["移除 AP"]), (1, 2, 1, 2))

    def test_committed_example_files_match_scenario(self):
        q2, q3 = (os.path.join(FIXTURES, f"compare_2026{q}.xlsx") for q in ("Q2", "Q3"))
        self.assertEqual(cla.compare_reports(q2, q3).counts(), self.result.counts())


class TestCompareEdgeCases(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmp.cleanup()

    def path(self, name):
        return os.path.join(self.tmp.name, name)

    def test_identical_reports_from_real_parser(self):
        devices, aps, _ = cla.analyze_folder(FIXTURES, log=lambda _m: None)
        cla.write_excel(self.path("a.xlsx"), devices, aps)
        counts = cla.compare_reports(self.path("a.xlsx"), self.path("a.xlsx")).counts()
        self.assertEqual({k: v for k, v in counts.items() if k != "需人工確認"},
                         {k: 0 for k in counts if k != "需人工確認"})

    def test_empty_sheets_written_as_none(self):
        devices, aps = compare_scenario.quarter("Q2")
        cla.write_excel(self.path("a.xlsx"), devices, aps)
        out = self.path("cmp.xlsx")
        cla.write_compare_excel(out, cla.compare_reports(self.path("a.xlsx"), self.path("a.xlsx")))
        self.assertEqual(load_workbook(out)["新增設備"].cell(2, 1).value, "無")

    def test_not_a_tool_report(self):
        wb = Workbook()
        wb.active.append(["Name", "Value"])
        wb.save(self.path("other.xlsx"))
        with self.assertRaises(cla.CompareError):
            cla.compare_reports(self.path("other.xlsx"), self.path("other.xlsx"))

    def test_not_an_excel_file(self):
        with open(self.path("bad.xlsx"), "w") as f:
            f.write("not a zip")
        with self.assertRaises(cla.CompareError):
            cla.load_report(self.path("bad.xlsx"))

    def test_duplicate_serial_flagged(self):
        devices, aps = compare_scenario.quarter("Q2")
        devices.append(dict(devices[-1], Hostname="FGT-DUP"))       # 同序號出現兩次
        cla.write_excel(self.path("a.xlsx"), devices, aps)
        review = cla.compare_reports(self.path("a.xlsx"), self.path("a.xlsx")).sheets["需人工確認"]
        self.assertTrue(any("序號重複" in r["原因"] for r in review))

    def test_alert_lines_ignore_timestamp(self):
        a = cla._alert_lines("Jun 10 08:00:01: %PLATFORM-2-PS_FAIL: Power supply 2 failed")
        b = cla._alert_lines("Sep 01 21:13:44: %PLATFORM-2-PS_FAIL: Power supply 2 failed")
        self.assertEqual(a.keys(), b.keys())
        self.assertEqual(cla._alert_lines("%LINK-3-UPDOWN: Interface Gi1/0/1, changed state to up"), {})


class TestCompareCli(unittest.TestCase):
    def test_compare_and_compare_with(self):
        with tempfile.TemporaryDirectory() as d:
            q2, q3 = os.path.join(d, "Q2.xlsx"), os.path.join(d, "Q3.xlsx")
            cla.write_excel(q2, *compare_scenario.quarter("Q2"))
            cla.write_excel(q3, *compare_scenario.quarter("Q3"))
            out = os.path.join(d, "cmp.xlsx")
            with contextlib.redirect_stdout(io.StringIO()) as buf:
                cla.main(["--compare", q2, q3, "-o", out])
                cla.main([FIXTURES, "-o", os.path.join(d, "now.xlsx"), "--compare-with", q2])
            self.assertIn("新增設備 2｜移除設備 1", buf.getvalue())
            self.assertIn("欄位異動", load_workbook(out).sheetnames)

            report = os.path.join(d, "now.xlsx")
            self.assertTrue(os.path.exists(report))
            self.assertTrue(os.path.exists(os.path.join(d, cla.default_compare_name(q2, report))))


if __name__ == "__main__":
    unittest.main()
