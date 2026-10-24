"""
新舊版本回歸比對：兩個版本各跑一次 analyze_folder，逐檔逐欄列出差異（忽略 _group）。
改版後用真實 Log 跑一次，每個差異都要回原始 Log 確認是「修正」而不是「退化」。

    python tests/compare_versions.py <舊版.py> <新版.py> <Log 資料夾> [<Log 資料夾> ...]

例：python tests/compare_versions.py Cisco_Log_Analysis_V21.py Cisco_Log_Analysis.py "..\\1150625"
只在本機執行；真實 Log 為客戶資料，不要提交到 Repo。
"""
import importlib.util
import sys


def load(path: str, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def run(mod, folders):
    out = {}
    for folder in folders:
        _, _, results = mod.analyze_folder(folder, log=lambda _m: None)
        for r in results:
            out[r.source] = {
                "type": r.dev_type, "error": r.error,
                "rows": [{k: v for k, v in row.items() if k != "_group"} for row in r.rows],
                "aps": r.aps,
            }
    return out


def main(old_path, new_path, *folders):
    if not folders:
        sys.exit(__doc__)
    old_mod, new_mod = load(old_path, "old_version"), load(new_path, "new_version")
    print(f"{getattr(old_mod, 'APP_VERSION', old_path)} → {getattr(new_mod, 'APP_VERSION', new_path)}")
    old, new = run(old_mod, folders), run(new_mod, folders)
    diffs = 0
    for source in sorted(set(old) | set(new)):
        a, b = old.get(source), new.get(source)
        if a is None or b is None:
            print(f"[只在一版出現] {source}")
            diffs += 1
            continue
        for key in ("type", "error"):
            if a[key] != b[key]:
                print(f"{source} | {key}: {a[key]!r} → {b[key]!r}")
                diffs += 1
        if len(a["rows"]) != len(b["rows"]):
            print(f"{source} | 列數: {len(a['rows'])} → {len(b['rows'])}")
            diffs += 1
        for i, (ra, rb) in enumerate(zip(a["rows"], b["rows"])):
            for col in sorted(set(ra) | set(rb)):
                if ra.get(col) != rb.get(col):
                    print(f"{source} | 第 {i + 1} 列 {col}: {ra.get(col)!r} → {rb.get(col)!r}")
                    diffs += 1
        if a["aps"] != b["aps"]:
            print(f"{source} | AP 清單不同（{len(a['aps'])} → {len(b['aps'])} 台）")
            diffs += 1
    print(f"共 {len(set(old) | set(new))} 個檔案，{diffs} 處差異")


if __name__ == "__main__":
    main(*sys.argv[1:])
