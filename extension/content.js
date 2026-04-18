function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function isWordText(s) {
  return /\.docx?\b/i.test(String(s || ""));
}

function getDownloadIconsWithin(root) {
  if (!root) return [];
  return Array.from(
    root.querySelectorAll(
      "span.sd-icon.icondownload, span.sd-icon-container span.sd-icon.icondownload, [class*='icondownload'], [title*='下载'], [aria-label*='下载']"
    )
  ).filter((el) => el instanceof HTMLElement);
}

function nearestDownloadIcon(root, anchorEl) {
  const icons = getDownloadIconsWithin(root);
  if (!icons.length || !anchorEl) return null;
  const a = anchorEl.getBoundingClientRect();
  const ax = a.left + a.width / 2;
  const ay = a.top + a.height / 2;
  let best = null;
  let bestDist = Number.POSITIVE_INFINITY;
  for (const icon of icons) {
    const r = icon.getBoundingClientRect();
    const ix = r.left + r.width / 2;
    const iy = r.top + r.height / 2;
    const dist = (ix - ax) ** 2 + (iy - ay) ** 2;
    if (dist < bestDist) {
      bestDist = dist;
      best = icon;
    }
  }
  return best;
}

function findWordDownloadButton() {
  // Find exact filename-like text nodes containing ".doc/.docx", then choose nearest icon.
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  const candidates = [];
  while (walker.nextNode()) {
    const node = walker.currentNode;
    const text = String(node?.nodeValue || "").trim();
    if (!text || !isWordText(text)) continue;
    const anchor = node.parentElement;
    if (!anchor) continue;

    let cur = anchor;
    for (let i = 0; i < 8 && cur && cur !== document.body; i++) {
      const blockText = String(cur.innerText || "");
      if (isWordText(blockText)) {
        const btn = nearestDownloadIcon(cur, anchor);
        if (btn) {
          candidates.push({ btn, anchor, textLen: blockText.length });
          break;
        }
      }
      cur = cur.parentElement;
    }
  }

  if (!candidates.length) return null;
  // Prefer the tightest container around the word filename text.
  candidates.sort((a, b) => a.textLen - b.textLen);
  return candidates[0].btn;
}

function toAbsUrl(url) {
  try {
    return new URL(url, location.href).toString();
  } catch {
    return null;
  }
}

function parseWordUrlFromOnclick(onclickText) {
  if (!onclickText) return null;
  const m = String(onclickText).match(/https?:\/\/[^"'\s)]+\.docx?(?:\?[^"'\s)]*)?/i);
  if (m?.[0]) return m[0];
  return null;
}

function firstNonEmptyLine(s) {
  const lines = String(s || "")
    .split(/\r?\n/)
    .map((x) => x.trim())
    .filter(Boolean);
  return lines[0] || "";
}

function normalizeJobTitle(raw) {
  let t = firstNonEmptyLine(raw);
  if (!t) return "";
  // Remove common suffix parts like "· Boss直聘 ..." etc.
  t = t.split(/[·•|｜]/)[0]?.trim() || "";
  // Some pages may append source/channel with hyphen.
  t = t.replace(/\s*[-–—]\s*(Boss直聘|BOSS直聘|拉勾|猎聘|智联|前程无忧|51job|内推|内部推荐|外部推荐).*$/i, "").trim();
  // Also drop trailing bracketed/source-ish tails.
  t = t.replace(/\s*\((主动投递|主动搜索|人才推荐|入才推荐|内推|外推)[^)]+\)\s*$/i, "").trim();
  return t;
}

function looksLikeAppliedBadgeText(s) {
  const t = String(s || "").trim();
  return /^已申请\s*\d+\s*次$/.test(t) || /^已申请\d+次$/.test(t) || t.includes("已申请");
}

function findJobTitleNearAppliedBadge() {
  // Strategy: find the "已申请X次" badge, then grab text from its right-side neighbor.
  const nodes = Array.from(document.querySelectorAll("span, div, a, button")).filter(
    (el) => el instanceof HTMLElement
  );

  for (const el of nodes) {
    const badgeText = firstNonEmptyLine(el.textContent || el.innerText || "");
    if (!looksLikeAppliedBadgeText(badgeText)) continue;

    // Try direct next siblings first (common layout).
    let cur = el.nextElementSibling;
    for (let i = 0; i < 6 && cur; i++) {
      const t = normalizeJobTitle(cur.textContent || cur.innerText || "");
      if (t && !looksLikeAppliedBadgeText(t)) return t;
      cur = cur.nextElementSibling;
    }

    // Try within same parent container: find the first text containing '·' that isn't the badge.
    const parent = el.parentElement;
    if (parent) {
      const sibs = Array.from(parent.children);
      const idx = sibs.indexOf(el);
      for (let j = idx + 1; j < Math.min(sibs.length, idx + 8); j++) {
        const s = sibs[j];
        const raw = s?.textContent || s?.innerText || "";
        if (String(raw).includes("·")) {
          const t = normalizeJobTitle(raw);
          if (t) return t;
        }
      }
    }
  }
  return "";
}

function extractJobTitleFromPage() {
  // Preferred: the dropdown text right after "已申请X次"
  const nearApplied = findJobTitleNearAppliedBadge();
  if (nearApplied) return nearApplied;

  // Fallback: old known header class (no document.title fallback)
  const selectors = [".candidate-header-info__item-pandect-current", "[class*='candidate-header-info__item-pandect-current']"];

  for (const sel of selectors) {
    const el = document.querySelector(sel);
    if (!el) continue;
    const t = normalizeJobTitle(el.textContent || el.innerText || "");
    if (t) return t;
  }

  return "";
}

function findWordUrlFromRow(btn) {
  if (!btn) return null;
  const row = btn.closest("div, li, tr");
  if (!row) return null;

  const link = row.querySelector(
    "a[href*='.doc'], a[href*='.DOC'], a[href*='.docx'], a[href*='.DOCX']"
  );
  if (link instanceof HTMLAnchorElement && link.href) return link.href;

  const anyWithData = row.querySelector("[data-url], [data-href], [data-download]");
  if (anyWithData instanceof HTMLElement) {
    const u = anyWithData.dataset?.url || anyWithData.dataset?.href || anyWithData.dataset?.download;
    const abs = toAbsUrl(u);
    if (abs && /\.docx?(\?|#|$)/i.test(abs)) return abs;
  }

  for (const el of row.querySelectorAll("[onclick]")) {
    if (!(el instanceof HTMLElement)) continue;
    const u = parseWordUrlFromOnclick(el.getAttribute("onclick"));
    if (u) return u;
  }
  return null;
}

async function triggerOnce() {
  // Give SPA some time to mount.
  for (let i = 0; i < 8; i++) {
    const btn = findWordDownloadButton();
    if (btn) {
      btn.click();
      return { ok: true };
    }
    await sleep(600);
  }
  return { ok: false, error: "找不到Word下载按钮（请确认页面上确实有 doc/docx 附件）" };
}

async function extractWordUrl() {
  for (let i = 0; i < 8; i++) {
    const btn = findWordDownloadButton();
    if (btn) {
      const url = findWordUrlFromRow(btn);
      if (url) return { ok: true, url };
      return { ok: false, error: "找到了Word行，但未提取到直链" };
    }
    await sleep(600);
  }
  return { ok: false, error: "找不到Word下载按钮" };
}

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  (async () => {
    if (msg?.type === "EXTRACT_WORD_URL") {
      const res = await extractWordUrl();
      sendResponse(res);
      return;
    }
    if (msg?.type === "EXTRACT_JOB_TITLE") {
      const title = extractJobTitleFromPage();
      sendResponse({ ok: true, title });
      return;
    }
    if (msg?.type === "TRIGGER_WORD_DOWNLOAD") {
      const res = await triggerOnce();
      sendResponse(res);
      return;
    }
    sendResponse({ ok: false, error: "unknown message" });
  })();
  return true;
});

