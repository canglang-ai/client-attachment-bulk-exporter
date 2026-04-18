const STATE = {
  running: false,
  tabId: null,
  currentJob: null,
  currentIndex: 0,
  total: 0,
  urlTemplate: "",
  folder: "",
  delaySec: 6,
  // downloadId -> { jobId, startedAt }
  trackedDownloads: new Map()
};

function isMokaDownload(item) {
  const url = item?.finalUrl || item?.url || "";
  // Keep function for compatibility; do not hard-limit domains because file host can vary.
  return /^https?:\/\//i.test(url);
}

function isWordDownload(item) {
  const url = item?.finalUrl || item?.url || "";
  const filename = item?.filename || "";
  const mime = item?.mime || "";
  return (
    /\.docx?(\?|#|$)/i.test(url) ||
    /\.docx?$/i.test(filename) ||
    mime === "application/msword" ||
    mime === "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  );
}

function buildUrl(urlTemplate, id) {
  return urlTemplate.replaceAll("{id}", encodeURIComponent(String(id)));
}

function safeNamePart(s) {
  return String(s || "")
    .trim()
    .replace(/[\\/:*?"<>|]+/g, "_")
    .replace(/\s+/g, " ")
    .slice(0, 80);
}

function desiredFilename(folder, job, ext = ".docx") {
  const sub = String(folder || "mokahr-pdfs").trim().replace(/^\/+|\/+$/g, "");
  const title = safeNamePart(job?.title || "");
  const who = safeNamePart(job?.name || job?.id || "");
  // If title is missing, keep old stable naming to avoid accidentally using page <title>.
  const base = title
    ? `${title}-${who || safeNamePart(job?.id || "")}`
    : job?.name
      ? `${job.id}-${safeNamePart(job.name)}`
      : `${job?.id || ""}`;
  return `${sub}/${base}${ext}`;
}

function wordExtFromItem(itemLike) {
  const url = itemLike?.finalUrl || itemLike?.url || "";
  const filename = itemLike?.filename || "";
  const blob = `${url} ${filename}`.toLowerCase();
  return /\.doc(\?|#|$)/.test(url.toLowerCase()) || /\.doc$/.test(filename.toLowerCase()) || /\.doc\b/.test(blob)
    ? ".doc"
    : ".docx";
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

async function ensureTab(url) {
  if (STATE.tabId) {
    try {
      await chrome.tabs.get(STATE.tabId);
      await chrome.tabs.update(STATE.tabId, { url, active: true });
      return STATE.tabId;
    } catch {
      STATE.tabId = null;
    }
  }
  const tab = await chrome.tabs.create({ url, active: true });
  STATE.tabId = tab.id;
  return STATE.tabId;
}

async function waitForTabComplete(tabId, timeoutMs = 45000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const tab = await chrome.tabs.get(tabId);
    if (tab.status === "complete") return;
    await sleep(300);
  }
  throw new Error("页面加载超时（可能需要登录，或网络较慢）");
}

async function triggerDownloadInTab(tabId) {
  const res = await chrome.tabs.sendMessage(tabId, { type: "TRIGGER_WORD_DOWNLOAD" }).catch(() => null);
  if (!res?.ok) throw new Error(res?.error || "未能触发下载（页面结构变化/未登录/无Word）");
}

async function extractWordUrlInTab(tabId) {
  const res = await chrome.tabs.sendMessage(tabId, { type: "EXTRACT_WORD_URL" }).catch(() => null);
  if (!res?.ok || !res?.url) return null;
  return String(res.url);
}

async function extractJobTitleInTab(tabId) {
  const res = await chrome.tabs.sendMessage(tabId, { type: "EXTRACT_JOB_TITLE" }).catch(() => null);
  if (!res?.ok) return "";
  return String(res.title || "").trim();
}

async function downloadWordDirect(url, job, folder) {
  const filename = desiredFilename(folder, job, wordExtFromItem({ url }));
  const downloadId = await chrome.downloads.download({
    url,
    filename,
    conflictAction: "uniquify",
    saveAs: false
  });
  if (typeof downloadId === "number") {
    STATE.trackedDownloads.set(downloadId, { jobId: job.id, startedAt: Date.now() });
    return downloadId;
  }
  return null;
}

async function waitForDownloadId(downloadId, timeoutMs = 60000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const [item] = await chrome.downloads.search({ id: downloadId });
    if (!item) {
      await sleep(300);
      continue;
    }
    if (item.state === "complete") return true;
    if (item.state === "interrupted") throw new Error(`下载中断: ${item.error || "unknown"}`);
    // Consider download started successfully; do not block queue forever.
    if (item.state === "in_progress" && Date.now() - start > 6000) return true;
    await sleep(300);
  }
  throw new Error("等待下载完成超时（直链下载未完成）");
}

async function waitForDownloadForJob(jobId, timeoutMs = 60000) {
  const start = Date.now();

  while (Date.now() - start < timeoutMs) {
    let hasTracked = false;
    // Find a tracked download for this job
    for (const [downloadId, meta] of STATE.trackedDownloads.entries()) {
      if (meta.jobId !== jobId) continue;
      hasTracked = true;
      const [item] = await chrome.downloads.search({ id: downloadId });
      if (!item) continue;
      if (item.state === "complete") return true;
      if (item.state === "interrupted") throw new Error(`下载中断: ${item.error || "unknown"}`);
      if (item.state === "in_progress" && Date.now() - meta.startedAt > 6000) return true;
    }
    if (hasTracked) {
      await sleep(500);
      continue;
    }
    await sleep(500);
  }
  throw new Error("等待下载完成超时（可能下载未触发，或被浏览器拦截）");
}

async function waitForAnyTrackingForJob(jobId, timeoutMs = 8000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    for (const [, meta] of STATE.trackedDownloads.entries()) {
      if (meta.jobId === jobId) return true;
    }
    await sleep(250);
  }
  return false;
}

async function updateProgress(extra = {}) {
  await chrome.storage.local.set({
    running: STATE.running,
    progress: {
      currentIndex: STATE.currentIndex,
      total: STATE.total,
      currentId: STATE.currentJob?.id || "",
      ...extra
    }
  });
}

async function runQueue() {
  const { queue } = await chrome.storage.local.get(["queue"]);
  if (!queue?.length) throw new Error("队列为空，请先导入表格");

  STATE.total = queue.length;
  const errors = [];
  await updateProgress({ total: STATE.total, errors: 0 });

  for (let i = STATE.currentIndex; i < queue.length; i++) {
    if (!STATE.running) break;
    STATE.currentIndex = i;
    STATE.currentJob = queue[i];
    await updateProgress({ errors: errors.length });

    let done = false;
    let lastErr = null;
    for (let attempt = 1; attempt <= 3 && !done && STATE.running; attempt++) {
      try {
        const url = buildUrl(STATE.urlTemplate, STATE.currentJob.id);
        const tabId = await ensureTab(url);
        await waitForTabComplete(tabId);

        // Give SPA a bit of time to render attachments section
        await sleep(1200);

        const existingTitle = String(STATE.currentJob?.title || "").trim();
        const title = existingTitle || (await extractJobTitleInTab(tabId).catch(() => ""));
        const jobWithTitle = { ...STATE.currentJob, title };
        await chrome.storage.local.set({ activeJob: { ...jobWithTitle, folder: STATE.folder, tabId } });

        let usedDirectDownload = false;
        const directUrl = await extractWordUrlInTab(tabId);
        if (directUrl && isWordDownload({ url: directUrl })) {
          const directDownloadId = await downloadWordDirect(directUrl, jobWithTitle, STATE.folder).catch(() => null);
          if (typeof directDownloadId === "number") {
            usedDirectDownload = true;
            await waitForDownloadId(directDownloadId, 25000);
          }
        }

        if (!usedDirectDownload) {
          // Fallback: click web page icon if direct URL cannot be extracted.
          await triggerDownloadInTab(tabId);
          // In some environments, browser prompts or policy blocks download events.
          // Do not block the entire queue forever: if no event appears, continue.
          const hasTracking = await waitForAnyTrackingForJob(STATE.currentJob.id, 8000);
          if (hasTracking) {
            await waitForDownloadForJob(STATE.currentJob.id, 25000);
          }
        }
        done = true;
      } catch (err) {
        lastErr = err;
        // Small backoff before next try.
        await sleep(1200 * attempt);
      }
    }

    if (!done) {
      const reason = lastErr?.message || String(lastErr || "unknown");
      errors.push({ id: STATE.currentJob.id, reason });
      await chrome.storage.local.set({ lastError: `ID ${STATE.currentJob.id}: ${reason}`, errors });
    }

    // Delay between jobs
    await sleep(Math.max(1, STATE.delaySec) * 1000);
  }

  STATE.running = false;
  await updateProgress({ done: true, errors: errors.length });
  await chrome.storage.local.remove(["activeJob"]);
}

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  (async () => {
    if (msg?.type === "START") {
      const { urlTemplate, folder, delaySec } = msg;
      STATE.urlTemplate = urlTemplate;
      STATE.folder = folder;
      STATE.delaySec = delaySec ?? 6;
      STATE.running = true;
      STATE.currentIndex = 0;
      STATE.currentJob = null;
      STATE.trackedDownloads.clear();
      await chrome.storage.local.set({ running: true, lastError: "", errors: [] });
      await updateProgress({ currentIndex: 0 });
      runQueue().catch(async (err) => {
        STATE.running = false;
        await chrome.storage.local.set({ running: false, lastError: err?.message || String(err) });
      });
      sendResponse({ ok: true });
      return;
    }

    if (msg?.type === "STOP") {
      STATE.running = false;
      await chrome.storage.local.set({ running: false });
      sendResponse({ ok: true });
      return;
    }

    sendResponse({ ok: false, error: "unknown message" });
  })();
  return true;
});

chrome.downloads.onDeterminingFilename.addListener((item, suggest) => {
  (async () => {
    try {
      if (!STATE.running) return suggest();
      const { activeJob } = await chrome.storage.local.get(["activeJob"]);
      if (!activeJob?.id) return suggest();
      if (!isMokaDownload(item) || !isWordDownload(item)) return suggest();

      // Only rename downloads that likely belong to the currently active tab/job.
      const desired = desiredFilename(activeJob.folder, activeJob, wordExtFromItem(item));

      STATE.trackedDownloads.set(item.id, { jobId: activeJob.id, startedAt: Date.now() });
      suggest({ filename: desired, conflictAction: "uniquify" });
    } catch {
      suggest();
    }
  })();
  return true;
});

chrome.downloads.onCreated.addListener((item) => {
  // Enforce "Word only" and backup tracking in case onDeterminingFilename didn't fire.
  (async () => {
    try {
      if (!STATE.running) return;
      const { activeJob } = await chrome.storage.local.get(["activeJob"]);
      if (!activeJob?.id) return;
      if (!isMokaDownload(item)) return;

      if (!isWordDownload(item)) {
        // Cancel non-word files immediately.
        await chrome.downloads.cancel(item.id).catch(() => {});
        await chrome.downloads.erase({ id: item.id }).catch(() => {});
        return;
      }

      STATE.trackedDownloads.set(item.id, { jobId: activeJob.id, startedAt: Date.now() });
    } catch {
      // ignore
    }
  })();
});

