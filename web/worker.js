// =============================================================================
// 網頁版 Python 執行端（Module Web Worker）
//
// 在瀏覽器內以 Pyodide 執行與桌面版完全相同的 Cisco_Log_Analysis.py：
//   1. 載入 ./pyodide/（自行託管，不連 CDN）
//   2. 把 wheels/manifest.json 列出的 openpyxl / et_xmlfile wheel 解壓到 site-packages（不用 micropip、不連外）
//   3. 載入 ./Cisco_Log_Analysis.py，呼叫 analyze_folder() 與 write_excel()
//      維護設備比對（V23）：呼叫 compare_reports() 與 write_compare_excel()
//   4. 回傳 Excel bytes 與各分頁列數
// Log 只寫進瀏覽器記憶體內的虛擬檔案系統，不會上傳到任何伺服器。
//
// Pyodide 314 起不支援 classic worker，index.html 必須以 new Worker(url, { type: "module" }) 啟動。
//
// 訊息格式
//   主執行緒 → worker：{ type: "analyze", files: [{ name, data: ArrayBuffer }] }
//                      { type: "compare", old: { name, data }, new: { name, data } }
//   worker → 主執行緒：status / ready / log / progress / done / compared / error（含 job）/ fatal
// =============================================================================

import { loadPyodide } from "./pyodide/pyodide.mjs";

const BASE = new URL(".", import.meta.url);
const send = (type, data = {}, transfer = []) => self.postMessage({ type, ...data }, transfer);

// analyze_folder / write_excel / build_sheets / compare_reports / write_compare_excel / compare_items
// 為桌面版既有介面；若修改其參數，需同步修改這裡
const PY_HELPER = `
import json
import Cisco_Log_Analysis as cla

def run_web(folder, out_path, log, progress):
    devices, aps, results = cla.analyze_folder(folder, log=log, progress=progress)
    cla.write_excel(out_path, devices, aps)
    sheets = cla.build_sheets(devices, aps)
    return json.dumps({
        "sheets": [[name, len(rows)] for name, (_cols, rows) in sheets.items()],
        "devices": len(devices),
        "aps": len(aps),
        "files": [{"source": r.source, "type": r.dev_type, "summary": r.summary,
                   "error": r.error, "seconds": round(r.seconds, 3)} for r in results],
    }, ensure_ascii=False)

def run_compare(old_path, new_path, out_path):
    try:
        r = cla.compare_reports(old_path, new_path)
    except cla.CompareError as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)
    cla.write_compare_excel(out_path, r)
    side = lambda d: {"name": d.name, "devices": len(d.devices), "aps": len(d.aps)}
    return json.dumps({
        "old": side(r.old), "new": side(r.new),
        "counts": [[name, count] for name, count in r.counts().items()],
        "items": cla.compare_items(r),
        "lines": cla.compare_log_lines(r),
        "fileName": cla.default_compare_name(old_path, new_path),
    }, ensure_ascii=False)
`;

async function fetchBytes(path) {
  const res = await fetch(new URL(path, BASE));
  if (!res.ok) throw new Error(`無法下載 ${path}（HTTP ${res.status}）`);
  return new Uint8Array(await res.arrayBuffer());
}

async function boot() {
  send("status", { text: "載入 Python 執行環境…" });
  const pyodide = await loadPyodide({
    indexURL: new URL("./pyodide/", BASE).href,
    stdout: (text) => send("log", { text }),
    stderr: (text) => send("log", { text }),
  });

  send("status", { text: "安裝 openpyxl…" });
  const manifest = JSON.parse(new TextDecoder().decode(await fetchBytes("./wheels/manifest.json")));
  pyodide.FS.mkdirTree("/wheels");
  for (const whl of manifest.wheels) {
    pyodide.FS.writeFile(`/wheels/${whl}`, await fetchBytes(`./wheels/${whl}`));
  }
  pyodide.runPython(`
import os, site, zipfile
for _name in sorted(os.listdir("/wheels")):
    with zipfile.ZipFile(os.path.join("/wheels", _name)) as _z:
        _z.extractall(site.getsitepackages()[0])
`);

  send("status", { text: "載入解析引擎…" });
  pyodide.FS.mkdirTree("/app");
  pyodide.FS.writeFile("/app/Cisco_Log_Analysis.py", await fetchBytes("./Cisco_Log_Analysis.py"));
  pyodide.runPython('import sys\nsys.path.insert(0, "/app")\n' + PY_HELPER);

  const version = pyodide.runPython("cla.APP_VERSION");
  send("ready", { version });
  return { pyodide, runWeb: pyodide.globals.get("run_web"), runCompare: pyodide.globals.get("run_compare") };
}

function resetDir(FS, dir) {
  FS.mkdirTree(dir);
  for (const name of FS.readdir(dir)) {
    if (name !== "." && name !== "..") FS.unlink(`${dir}/${name}`);
  }
}

async function analyze(engine, files) {
  const { pyodide, runWeb } = engine;
  const FS = pyodide.FS;
  resetDir(FS, "/logs");
  for (const f of files) {
    FS.writeFile(`/logs/${f.name.replace(/[\\/]/g, "_")}`, new Uint8Array(f.data));
  }
  FS.mkdirTree("/out");
  const out = "/out/report.xlsx";
  const result = JSON.parse(runWeb(
    "/logs", out,
    (text) => send("log", { text: String(text) }),
    (done, total) => send("progress", { done, total }),
  ));
  const bytes = FS.readFile(out);
  FS.unlink(out);
  resetDir(FS, "/logs");                       // 分析完即清除記憶體中的 Log
  send("done", { ...result, xlsx: bytes.buffer }, [bytes.buffer]);
}

async function compare(engine, oldFile, newFile) {
  const { pyodide, runCompare } = engine;
  const FS = pyodide.FS;
  const paths = [];
  for (const [dir, f] of [["/cmp/old", oldFile], ["/cmp/new", newFile]]) {
    resetDir(FS, dir);                         // 分開資料夾：兩份報表同名也不會互相覆蓋
    const p = `${dir}/${f.name.replace(/[\\/]/g, "_")}`;
    FS.writeFile(p, new Uint8Array(f.data));
    paths.push(p);
  }
  FS.mkdirTree("/out");
  const out = "/out/compare.xlsx";
  const result = JSON.parse(runCompare(paths[0], paths[1], out));
  resetDir(FS, "/cmp/old");
  resetDir(FS, "/cmp/new");
  if (result.error) throw new Error(result.error);
  for (const text of result.lines) send("log", { text });
  const bytes = FS.readFile(out);
  FS.unlink(out);
  send("compared", { ...result, xlsx: bytes.buffer }, [bytes.buffer]);
}

const ready = boot().catch((err) => {
  send("fatal", { text: String(err && err.message ? err.message : err) });
  throw err;
});

let busy = Promise.resolve();
self.onmessage = (event) => {
  const msg = event.data || {};
  if (msg.type !== "analyze" && msg.type !== "compare") return;
  busy = busy.then(async () => {
    try {
      if (msg.type === "compare") await compare(await ready, msg.old, msg.new);
      else await analyze(await ready, msg.files || []);
    } catch (err) {
      send("error", { job: msg.type, text: String(err && err.message ? err.message : err) });
    }
  });
};
