const DEFAULTS = {
  urlTemplate: "https://app.mokahr.com/candidates/application/{id}/more-info",
  folder: "mokahr-words",
  delaySec: 6
};

function setStatus(msg) {
  const el = document.getElementById("status");
  el.textContent = msg;
}

function normalizeFolder(folder) {
  const f = String(folder || "").trim().replace(/^\/+|\/+$/g, "");
  return f || DEFAULTS.folder;
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function parseCsv(text) {
  // Very small CSV parser: split by lines, then by comma (handles simple cases).
  // If you need complex quoting, use XLSX instead.
  const lines = text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  return lines.map((line) => line.split(",").map((c) => c.trim()));
}

function pickRowsToJobs(rows) {
  // rows: array of arrays (cells)
  // Heuristics: allow header row. Try to find column named application_id / id.
  if (!rows.length) return [];
  const header = rows[0].map((x) => String(x || "").trim().toLowerCase());
  const hasHeader = header.some((h) => ["application_id", "appid", "id", "申请id", "申请ID"].includes(h));
  let startIdx = 0;
  let idCol = 0;
  // Default (no header): application_id, job_title, name
  let titleCol = 1;
  let nameCol = 2;
  if (hasHeader) {
    startIdx = 1;
    idCol = Math.max(
      0,
      header.findIndex((h) => ["application_id", "appid", "id", "申请id", "申请id", "申请ID".toLowerCase()].includes(h))
    );
    nameCol = header.findIndex((h) => ["name", "姓名", "candidate", "候选人"].includes(h));
    titleCol = header.findIndex((h) =>
      ["job_title", "title", "position", "职位", "岗位", "应聘职位", "应聘职位名称", "投递职位", "投递岗位"].includes(h)
    );
  }

  const jobs = [];
  for (let i = startIdx; i < rows.length; i++) {
    const row = rows[i] || [];
    const rawId = row[idCol];
    const id = String(rawId ?? "").trim();
    if (!id || !/^\d+$/.test(id)) continue;
    const name = nameCol >= 0 ? String(row[nameCol] ?? "").trim() : "";
    const title = titleCol >= 0 ? String(row[titleCol] ?? "").trim() : "";
    jobs.push({ id, name, title });
  }
  return jobs;
}

async function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(r.error);
    r.readAsArrayBuffer(file);
  });
}

async function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(r.error);
    r.readAsText(file, "utf-8");
  });
}

async function loadSettingsIntoUI() {
  const { urlTemplate, folder, delaySec, lastImportCount, running, progress, lastError, errors, queue } =
    await chrome.storage.local.get([
      "urlTemplate",
      "folder",
      "delaySec",
      "lastImportCount",
      "running",
      "progress",
      "lastError",
      "errors",
      "queue"
    ]);

  document.getElementById("urlTemplate").value = urlTemplate || DEFAULTS.urlTemplate;
  document.getElementById("folder").value = folder || DEFAULTS.folder;
  document.getElementById("delaySec").value = String(delaySec || DEFAULTS.delaySec);

  const p = progress || {};
  const errList = Array.isArray(errors) ? errors : [];
  const queueCount = Array.isArray(queue) ? queue.length : 0;
  const latestErr = lastError ? `\n错误: ${lastError}` : "";
  const startBtn = document.getElementById("start");
  startBtn.disabled = !running && queueCount === 0;
  startBtn.style.opacity = !running && queueCount === 0 ? "0.6" : "1";
  startBtn.style.cursor = !running && queueCount === 0 ? "not-allowed" : "pointer";
  const msg =
    running
      ? `运行中…\n已导入: ${lastImportCount || 0}\n当前: ${p.currentIndex ?? 0}/${p.total ?? 0}\n当前ID: ${p.currentId ?? ""}\n失败数: ${errList.length}${latestErr}`
      : `等待导入…\n队列: ${queueCount}\n上次导入: ${lastImportCount || 0}\n失败数: ${errList.length}${latestErr}`;
  setStatus(msg);
}

async function saveSettingsFromUI() {
  const urlTemplate = document.getElementById("urlTemplate").value.trim() || DEFAULTS.urlTemplate;
  const folder = normalizeFolder(document.getElementById("folder").value);
  const delaySec = Math.max(1, Math.min(60, Number(document.getElementById("delaySec").value || DEFAULTS.delaySec)));
  await chrome.storage.local.set({ urlTemplate, folder, delaySec });
  return { urlTemplate, folder, delaySec };
}

async function importFile(file) {
  if (!file) throw new Error("请选择文件");

  const name = file.name.toLowerCase();
  let rows = [];

  if (name.endsWith(".csv")) {
    const text = await readFileAsText(file);
    rows = parseCsv(text);
  } else {
    // XLSX/XLS
    const buf = await readFileAsArrayBuffer(file);
    const wb = XLSX.read(buf, { type: "array" });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
  }

  const jobs = pickRowsToJobs(rows);
  if (!jobs.length) throw new Error("没有解析到申请ID（请确保第一列是纯数字ID，或表头包含 application_id/id）");

  const dedup = [];
  const seen = new Set();
  for (const j of jobs) {
    if (seen.has(j.id)) continue;
    seen.add(j.id);
    dedup.push(j);
  }

  await chrome.storage.local.set({
    queue: dedup,
    lastImportCount: dedup.length,
    progress: { currentIndex: 0, total: dedup.length, currentId: "" },
    lastError: "",
    errors: []
  });

  return dedup.length;
}

document.addEventListener("DOMContentLoaded", async () => {
  await loadSettingsIntoUI();
  setInterval(() => {
    loadSettingsIntoUI().catch(() => {});
  }, 1000);

  document.getElementById("file").addEventListener("change", async (e) => {
    try {
      const file = e.target.files?.[0];
      const count = await importFile(file);
      await loadSettingsIntoUI();
      setStatus(`导入成功: ${count}\n可点击 Start 开始下载。`);
    } catch (err) {
      setStatus(`导入失败: ${err?.message || String(err)}`);
    }
  });

  document.getElementById("start").addEventListener("click", async () => {
    try {
      const { urlTemplate, folder, delaySec } = await saveSettingsFromUI();
      const { queue } = await chrome.storage.local.get(["queue"]);
      if (!queue?.length) {
        setStatus("请先导入表格（必须包含申请ID）");
        return;
      }
      await chrome.runtime.sendMessage({ type: "START", urlTemplate, folder, delaySec });
      // Give background a moment to update status.
      await sleep(250);
      await loadSettingsIntoUI();
    } catch (err) {
      setStatus(`启动失败: ${err?.message || String(err)}`);
    }
  });

  document.getElementById("stop").addEventListener("click", async () => {
    await chrome.runtime.sendMessage({ type: "STOP" });
    await sleep(200);
    await loadSettingsIntoUI();
    setStatus("已停止。");
  });
});

