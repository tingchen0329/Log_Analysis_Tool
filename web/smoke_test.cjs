// =============================================================================
// 網頁版煙霧測試：在 Node 裡模擬「瀏覽器的 Module Web Worker」執行 _site/worker.js
//
//   node --no-warnings web/smoke_test.cjs <_site> <Log 資料夾> [--out 報表.xlsx]
//
// 以 vm.SourceTextModule 載入 worker.js 與 pyodide.mjs（不定義 process，Pyodide 會走瀏覽器 Worker 分支），
// fetch 與動態 import 都對應到 _site 內的檔案，因此測到的是實際部署到 GitHub Pages 的同一份程式碼。
// 檢查：版本號與 Cisco_Log_Analysis.py 一致、每個檔案都判定成功、分頁清單正確、輸出為有效的 xlsx。
// 維護設備比對（V23）：Log 資料夾內有 compare_2026Q2.xlsx / compare_2026Q3.xlsx 時一併測試比對。
// --out 可另存網頁版產生的 Excel，用來與 CPython 版逐格比對；--compare-out 另存比對報表。
// =============================================================================
"use strict";

const vm = require("node:vm");

if (typeof vm.SourceTextModule !== "function") {
  // vm 模組 API 需要 --experimental-vm-modules：自動以該旗標重新執行本腳本
  const { spawnSync } = require("node:child_process");
  const r = spawnSync(process.execPath, ["--experimental-vm-modules", "--no-warnings", __filename, ...process.argv.slice(2)],
                      { stdio: "inherit" });
  process.exit(r.status === null ? 1 : r.status);
}

const fs = require("node:fs");
const path = require("node:path");

const args = process.argv.slice(2);
const outIdx = args.indexOf("--out");
const outFile = outIdx >= 0 ? args.splice(outIdx, 2)[1] : null;
const cmpIdx = args.indexOf("--compare-out");
const compareOutFile = cmpIdx >= 0 ? args.splice(cmpIdx, 2)[1] : null;
const [siteArg, logArg] = args;
if (!siteArg || !logArg) {
  console.error("用法：node --no-warnings web/smoke_test.cjs <_site> <Log 資料夾> [--out 報表.xlsx] [--compare-out 比對.xlsx]");
  process.exit(2);
}
const SITE = path.resolve(siteArg);
const LOGS = path.resolve(logArg);
const ORIGIN = "https://pages.test/Log_Analysis_Tool/";
const LOG_EXTS = [".log", ".txt", ".conf"];
const TIMEOUT_MS = 300000;
const COMPARE_FILES = ["compare_2026Q2.xlsx", "compare_2026Q3.xlsx"];
// 與 tests/compare_scenario.py 的情境一致（tests/test_compare.py 以 CPython 驗證同一組數字）
const COMPARE_EXPECTED = { "新增設備": 2, "移除設備": 1, "換機與成員變更": 3, "欄位異動": 6,
                           "新增 AP": 2, "移除 AP": 1, "需人工確認": 1 };

function fail(message) {
  console.error(`✗ ${message}`);
  process.exit(1);
}

function toFile(url) {
  const u = new URL(url);
  if (!u.href.startsWith(ORIGIN)) throw new Error(`不允許的網址（網頁版必須完全自行託管）：${u.href}`);
  return path.join(SITE, decodeURIComponent(u.pathname.slice(new URL(ORIGIN).pathname.length)));
}

const MIME = { ".wasm": "application/wasm", ".json": "application/json", ".js": "text/javascript",
               ".mjs": "text/javascript", ".py": "text/x-python", ".zip": "application/zip" };

async function fetchShim(input) {
  const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
  let data;
  try {
    data = fs.readFileSync(toFile(url));
  } catch (err) {
    return new Response("not found", { status: 404 });
  }
  return new Response(data, { status: 200, headers: { "Content-Type": MIME[path.extname(url)] || "application/octet-stream" } });
}

// ── 模擬 Module Web Worker 的全域環境 ─────────────────────────────────────────
const messages = [];
const waiters = [];
const context = vm.createContext({
  // WebAssembly 用 vm 自己的（不從主環境傳入）：Emscripten 以 instanceof TypeError 判斷錯誤，跨 realm 會判斷失敗
  console, URL, URLSearchParams, Response, Headers, Request, TextDecoder, TextEncoder,
  setTimeout, clearTimeout, setInterval, clearInterval, queueMicrotask, performance, structuredClone,
  atob, btoa, AbortController, crypto: globalThis.crypto, fetch: fetchShim,
  navigator: { userAgent: "Mozilla/5.0 (smoke-test)", hardwareConcurrency: 2, language: "zh-TW" },
  location: new URL("worker.js", ORIGIN),
  WorkerGlobalScope: class WorkerGlobalScope { static [Symbol.hasInstance]() { return true; } },
  postMessage: (msg) => {
    messages.push(msg);
    if (msg.type === "log" || msg.type === "status") console.log(`  [${msg.type}] ${msg.text}`);
    for (const w of waiters.splice(0)) w();
  },
});
vm.runInContext("globalThis.self = globalThis;", context);

const modules = new Map();
function loadModule(url) {
  if (!modules.has(url)) {
    modules.set(url, (async () => {
      const mod = new vm.SourceTextModule(fs.readFileSync(toFile(url), "utf8"), {
        context,
        identifier: url,
        initializeImportMeta(meta) { meta.url = url; },
        importModuleDynamically: async (specifier) => {
          const dep = await loadModule(new URL(specifier, url).href);
          if (dep.status === "linked") await dep.evaluate();
          return dep;
        },
      });
      await mod.link((specifier) => loadModule(new URL(specifier, url).href));
      return mod;
    })());
  }
  return modules.get(url);
}

function waitFor(predicate, what) {
  return new Promise((resolve, reject) => {
    const check = () => {
      const hit = messages.find(predicate);
      const bad = messages.find((m) => m.type === "fatal" || m.type === "error");
      if (bad) return reject(new Error(`${what}失敗：${bad.text}`));
      if (hit) return resolve(hit);
      waiters.push(check);
    };
    check();
  });
}

async function main() {
  const started = Date.now();
  const timer = setTimeout(() => fail(`逾時（${TIMEOUT_MS / 1000} 秒）`), TIMEOUT_MS);

  const worker = await loadModule(new URL("worker.js", ORIGIN).href);
  await worker.evaluate();
  const ready = await waitFor((m) => m.type === "ready", "載入");
  const pySource = fs.readFileSync(path.join(SITE, "Cisco_Log_Analysis.py"), "utf8");
  const expectedVersion = (pySource.match(/^APP_VERSION = "([^"]+)"/m) || [])[1];
  if (ready.version !== expectedVersion) fail(`版本不符：網頁版 ${ready.version}，程式 ${expectedVersion}`);
  console.log(`✓ 就緒（${ready.version}），載入 ${((Date.now() - started) / 1000).toFixed(1)} 秒`);

  const names = fs.readdirSync(LOGS).filter((n) => LOG_EXTS.includes(path.extname(n).toLowerCase()) && !n.startsWith("~$"));
  if (!names.length) fail(`${LOGS} 沒有 Log 檔`);
  const files = names.map((name) => {
    const buf = fs.readFileSync(path.join(LOGS, name));
    return { name, data: buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) };
  });
  const t0 = Date.now();
  context.onmessage({ data: { type: "analyze", files } });
  const done = await waitFor((m) => m.type === "done", "分析");
  console.log(`✓ 分析 ${names.length} 個檔案，${((Date.now() - t0) / 1000).toFixed(2)} 秒`);

  const progress = messages.filter((m) => m.type === "progress");
  if (!progress.length || progress[progress.length - 1].done !== names.length) fail("進度回報不完整");

  const bad = done.files.filter((f) => f.error || f.type === "Unknown");
  if (bad.length) fail(`解析失敗：${bad.map((f) => `${f.source}（${f.error || "Unknown"}）`).join("、")}`);
  for (const f of done.files) console.log(`  ${f.source} → ${f.summary}`);

  const sheets = done.sheets.map(([name]) => name);
  const expected = ["Report", "AP Info", "WLC Info", "Switch Info"];
  if (done.files.some((f) => f.type === "Aruba")) expected.push("Aruba Info");
  if (done.files.some((f) => f.type === "FortiGate")) expected.push("Fortigate Info");
  if (JSON.stringify(sheets) !== JSON.stringify(expected)) fail(`分頁不符：${sheets.join(", ")}`);
  console.log(`✓ 分頁：${done.sheets.map(([n, c]) => `${n}(${c})`).join("、")}`);

  const xlsx = Buffer.from(done.xlsx);
  if (xlsx.length < 1000 || xlsx.subarray(0, 2).toString() !== "PK") fail("輸出不是有效的 xlsx");
  if (outFile) {
    fs.writeFileSync(outFile, xlsx);
    console.log(`✓ Excel 已存到 ${outFile}（${xlsx.length} bytes）`);
  }
  if (COMPARE_FILES.every((n) => fs.existsSync(path.join(LOGS, n)))) {
    const [oldFile, newFile] = COMPARE_FILES.map((name) => {
      const buf = fs.readFileSync(path.join(LOGS, name));
      return { name, data: buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) };
    });
    const t1 = Date.now();
    context.onmessage({ data: { type: "compare", old: oldFile, new: newFile } });
    const cmp = await waitFor((m) => m.type === "compared", "比對");
    const counts = Object.fromEntries(cmp.counts);
    if (JSON.stringify(counts) !== JSON.stringify(COMPARE_EXPECTED)) fail(`比對結果不符：${JSON.stringify(counts)}`);
    if (cmp.items.length !== Object.values(COMPARE_EXPECTED).reduce((a, b) => a + b, 0)) fail(`比對明細筆數不符：${cmp.items.length}`);
    const cmpXlsx = Buffer.from(cmp.xlsx);
    if (cmpXlsx.length < 1000 || cmpXlsx.subarray(0, 2).toString() !== "PK") fail("比對輸出不是有效的 xlsx");
    console.log(`✓ 維護設備比對：${cmp.counts.map(([n, c]) => `${n} ${c}`).join("、")}（${((Date.now() - t1) / 1000).toFixed(2)} 秒）`);
    if (compareOutFile) {
      fs.writeFileSync(compareOutFile, cmpXlsx);
      console.log(`✓ 比對報表已存到 ${compareOutFile}（${cmpXlsx.length} bytes）`);
    }

    // 不是本工具的報表 → 回傳 error（job = compare），不可讓 worker 卡住
    const before = messages.length;
    context.onmessage({ data: { type: "compare", old: { name: "bad.xlsx", data: new ArrayBuffer(8) },
                                new: { name: "bad2.xlsx", data: new ArrayBuffer(8) } } });
    const err = await new Promise((resolve) => {
      const check = () => {
        const hit = messages.slice(before).find((m) => m.type === "error" && m.job === "compare");
        if (hit) resolve(hit); else waiters.push(check);
      };
      check();
    });
    messages.splice(before);                       // 預期中的錯誤，移除以免影響後續檢查
    console.log(`✓ 錯誤檔案會回報：${err.text}`);
  }
  clearTimeout(timer);
  console.log("✓ 網頁版煙霧測試通過");
}

main().catch((err) => fail(err && err.stack ? err.stack : String(err)));
