"""
組出網頁版的靜態網站（GitHub Pages 與本機測試共用）。

    python web/build_site.py [輸出資料夾，預設 _site]

產出內容（全部自行託管，不連任何 CDN；公司網路常擋 jsDelivr 等外部來源）：
    _site/index.html              介面
    _site/worker.js               Module Web Worker：在瀏覽器內執行 Python
    _site/Cisco_Log_Analysis.py   與桌面版 / exe 完全相同的解析引擎
    _site/pyodide/                Pyodide 執行環境（從 npm registry 下載，比對 sha512）
    _site/wheels/                 openpyxl / et_xmlfile wheel + manifest.json

本機預覽：python -m http.server -d _site 8000，再開 http://localhost:8000/
（不能直接雙擊 index.html：瀏覽器不允許 file:// 頁面啟動 Web Worker）
"""
import base64
import hashlib
import io
import json
import os
import shutil
import subprocess
import sys
import tarfile
import urllib.request

PYODIDE_VERSION = "314.0.7"
PYODIDE_SHA512 = "0YvXxEhfEdpLfb/XkM2BFAeMROq0iMUX2bzzH9pOttyMcWkwq+HbE5uyuGD82LN7y2q+SNvi/6V5JEsOlD2R1A=="
PYODIDE_FILES = ("pyodide.mjs", "pyodide.asm.mjs", "pyodide.asm.wasm", "python_stdlib.zip", "pyodide-lock.json")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEB = os.path.join(ROOT, "web")


def fetch_pyodide(dest: str):
    url = f"https://registry.npmjs.org/pyodide/-/pyodide-{PYODIDE_VERSION}.tgz"
    print(f"下載 {url}")
    with urllib.request.urlopen(url, timeout=120) as r:
        data = r.read()
    digest = base64.b64encode(hashlib.sha512(data).digest()).decode()
    if digest != PYODIDE_SHA512:
        sys.exit(f"Pyodide 套件 sha512 不符：{digest}")
    os.makedirs(dest, exist_ok=True)
    with tarfile.open(fileobj=io.BytesIO(data), mode="r:gz") as tar:
        for name in PYODIDE_FILES:
            member = tar.extractfile(f"package/{name}")
            if member is None:
                sys.exit(f"Pyodide 套件缺少 {name}")
            with open(os.path.join(dest, name), "wb") as f:
                f.write(member.read())


def fetch_wheels(dest: str):
    os.makedirs(dest, exist_ok=True)
    subprocess.run([sys.executable, "-m", "pip", "download", "--quiet", "--no-deps", "--only-binary=:all:",
                    "-r", os.path.join(WEB, "requirements-web.txt"), "-d", dest], check=True)
    wheels = sorted(f for f in os.listdir(dest) if f.endswith(".whl"))
    if not any(w.startswith("openpyxl-") for w in wheels):
        sys.exit("找不到 openpyxl wheel")
    with open(os.path.join(dest, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump({"wheels": wheels}, f)


def main(out: str = "_site"):
    out = os.path.abspath(out)
    if os.path.isdir(out):
        shutil.rmtree(out)
    os.makedirs(out)
    fetch_pyodide(os.path.join(out, "pyodide"))
    fetch_wheels(os.path.join(out, "wheels"))
    for name in ("index.html", "worker.js"):
        shutil.copy2(os.path.join(WEB, name), out)
    shutil.copy2(os.path.join(ROOT, "Cisco_Log_Analysis.py"), out)
    open(os.path.join(out, ".nojekyll"), "w").close()
    size = sum(os.path.getsize(os.path.join(d, f)) for d, _, fs in os.walk(out) for f in fs)
    print(f"完成：{out}（{size / 1024 / 1024:.1f} MB）")


if __name__ == "__main__":
    main(*sys.argv[1:2])
