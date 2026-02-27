// ==UserScript==
// @name         Zendesk Weekly Report Popup
// @namespace    http://tampermonkey.net/
// @version      1.0.0
// @description  Run weekly Zendesk ticket reports (assigned, solved, and takeovers) from a popup.
// @match        https://retail-support.zendesk.com/agent/*
// @grant        none
// ==/UserScript==

(function () {
  "use strict";

  const ASSIGNEE_FALLBACK_KEYWORD = "me";
  const FIXED_WEEK_STARTS_ON = 0; // Sunday
  const TAKEOVER_BATCH_SIZE = 100;
  const TAKEOVER_BATCH_SLEEP_MS = 30_000;
  const LAUNCHER_POSITION_STORAGE_KEY = "zd-weekly-report-launcher-position";

  const BASE = "https://retail-support.zendesk.com";
  const FLATPICKR_JS_URL = "https://cdn.jsdelivr.net/npm/flatpickr@4.6.13/dist/flatpickr.min.js";
  const FLATPICKR_CSS_URL = "https://cdn.jsdelivr.net/npm/flatpickr@4.6.13/dist/flatpickr.min.css";

  const css = `
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600&display=swap');

:root {
  --zd-accent: #0ea5e9;
  --zd-accent-2: #f59e0b;
  --zd-surface: #0b132b;
  --zd-card: #ffffff;
  --zd-muted: #6b7280;
  --zd-border: #e5e7eb;
  --zd-shadow: 0 18px 46px rgba(0, 0, 0, 0.25);
  --zd-radius: 12px;
  --zd-font: 'Space Grotesk', 'Helvetica Neue', sans-serif;
}

#zd-weekly-report-launcher {
  position: fixed;
  right: 16px;
  top: 16px;
  z-index: 999999;
  background: linear-gradient(135deg, var(--zd-accent), var(--zd-accent-2));
  color: #fff;
  border: none;
  border-radius: 999px;
  padding: 12px 18px;
  cursor: pointer;
  font-size: 14px;
  font-weight: 600;
  letter-spacing: 0.2px;
  box-shadow: var(--zd-shadow);
  cursor: grab;
  user-select: none;
}
#zd-weekly-report-launcher:hover {
  transform: translateY(-1px);
  transition: transform 120ms ease, box-shadow 120ms ease;
}
#zd-weekly-report-launcher.dragging {
  cursor: grabbing;
  transform: none;
}

#zd-weekly-report-overlay {
  position: fixed;
  inset: 0;
  background: radial-gradient(circle at 20% 20%, rgba(14,165,233,0.16), transparent 50%),
              radial-gradient(circle at 80% 0%, rgba(245,158,11,0.16), transparent 45%),
              rgba(6, 8, 20, 0.55);
  z-index: 999998;
  display: none;
}

#zd-weekly-report-modal {
  position: fixed;
  top: 6%;
  left: 50%;
  transform: translateX(-50%);
  width: min(1060px, 96vw);
  max-height: 88vh;
  background: var(--zd-card);
  border-radius: var(--zd-radius);
  box-shadow: var(--zd-shadow);
  z-index: 999999;
  display: none;
  overflow: hidden;
  font-family: var(--zd-font);
}

#zd-weekly-report-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 14px 16px;
  background: linear-gradient(135deg, rgba(14,165,233,0.12), rgba(245,158,11,0.12));
  border-bottom: 1px solid var(--zd-border);
}
#zd-weekly-report-header h3 {
  margin: 0;
  font-size: 17px;
  letter-spacing: 0.2px;
}
#zd-weekly-report-header .subtext {
  margin: 2px 0 0;
  color: var(--zd-muted);
  font-size: 12px;
}
#zd-weekly-report-close {
  border: 1px solid var(--zd-border);
  background: #fff;
  border-radius: 8px;
  padding: 8px 10px;
  cursor: pointer;
}

#zd-weekly-report-controls {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 10px;
  padding: 14px 16px;
  border-bottom: 1px solid var(--zd-border);
}
#zd-weekly-report-controls label {
  display: flex;
  flex-direction: column;
  gap: 6px;
  font-size: 12px;
  color: #111827;
}
#zd-weekly-report-controls input,
#zd-weekly-report-controls select,
#zd-weekly-report-controls button {
  font-size: 13px;
  padding: 8px 10px;
  border-radius: 8px;
  border: 1px solid var(--zd-border);
  font-family: var(--zd-font);
}
#zd-weekly-report-controls input:focus,
#zd-weekly-report-controls select:focus {
  outline: 2px solid rgba(14,165,233,0.25);
  border-color: var(--zd-accent);
}
#zd-weekly-report-controls .actions {
  display: flex;
  gap: 8px;
  align-items: end;
}
#zd-run-reports {
  background: var(--zd-accent);
  color: #fff;
  border: none;
  font-weight: 600;
}
#zd-copy-output {
  background: #f3f4f6;
  border: 1px solid var(--zd-border);
}
#zd-copy-report {
  background: #ecfeff;
  border: 1px solid #bae6fd;
}

#zd-weekly-report-status {
  padding: 10px 16px;
  font-size: 12px;
  border-bottom: 1px solid var(--zd-border);
  background: #f8fafc;
  color: #0f172a;
}

#zd-weekly-report-output {
  padding: 14px 16px 18px;
  overflow: auto;
  max-height: 64vh;
  background: #f8fafc;
}

.zd-placeholder {
  color: var(--zd-muted);
  font-size: 13px;
}

.zd-card-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
  gap: 12px;
  margin-top: 10px;
}
.zd-card {
  background: #fff;
  border: 1px solid var(--zd-border);
  border-radius: 12px;
  padding: 12px;
  box-shadow: 0 4px 18px rgba(0,0,0,0.04);
}
.zd-card h4 {
  margin: 0 0 6px;
  font-size: 14px;
  color: #0f172a;
  font-weight: 600;
}
.zd-metric {
  font-size: 26px;
  font-weight: 600;
  line-height: 1.1;
  color: #0f172a;
}
.zd-subline {
  color: #334155;
  font-size: 12px;
  margin-top: 2px;
}

.zd-chip-row {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  margin-top: 8px;
}
.zd-chip {
  background: #e0f2fe;
  border: 1px solid #7dd3fc;
  border-radius: 999px;
  padding: 4px 8px;
  font-size: 12px;
  color: #0c4a6e;
  font-weight: 600;
  text-decoration: none;
  display: inline-flex;
  align-items: center;
}
.zd-chip:hover {
  background: #bae6fd;
  border-color: #38bdf8;
}
.zd-chip:focus-visible {
  outline: 2px solid #0ea5e9;
  outline-offset: 1px;
}

.zd-meta-card {
  background: linear-gradient(135deg, rgba(14,165,233,0.09), rgba(245,158,11,0.09));
  border: 1px solid var(--zd-border);
  border-radius: 12px;
  padding: 12px;
  box-shadow: 0 4px 18px rgba(0,0,0,0.04);
}
.zd-meta-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
  gap: 8px 12px;
  margin-top: 8px;
  font-size: 12px;
}
.zd-meta-grid div span {
  display: block;
  color: #475569;
}
.zd-meta-grid strong {
  color: #0f172a;
}

.zd-query {
  margin-top: 6px;
  font-size: 11px;
  color: #1e293b;
  background: #e2e8f0;
  border-radius: 8px;
  padding: 8px;
  word-break: break-word;
}

.zd-section-title {
  margin: 14px 0 4px;
  font-size: 13px;
  font-weight: 600;
}

.zd-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12px;
  margin-top: 8px;
}
.zd-table th,
.zd-table td {
  border-bottom: 1px solid var(--zd-border);
  padding: 6px 4px;
  text-align: left;
}
.zd-table th {
  color: #0f172a;
  font-weight: 600;
}
.zd-table td {
  color: #1e293b;
}
.zd-table tr:last-child td {
  border-bottom: none;
}

.zd-report-draft {
  margin-top: 14px;
  background: #fff;
  border: 1px solid var(--zd-border);
  border-radius: 12px;
  padding: 12px;
  box-shadow: 0 4px 18px rgba(0,0,0,0.04);
}
.zd-report-draft h4 {
  margin: 0 0 8px;
  font-size: 14px;
}
.zd-report-draft textarea {
  width: 100%;
  min-height: 220px;
  resize: vertical;
  border: 1px solid var(--zd-border);
  border-radius: 8px;
  padding: 10px;
  font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  font-size: 12px;
  line-height: 1.45;
  background: #fcfcfd;
}

.flatpickr-calendar {
  z-index: 1000001 !important;
  pointer-events: auto !important;
}

@media (max-width: 980px) {
  #zd-weekly-report-controls {
    grid-template-columns: 1fr;
  }
}
@media (max-width: 620px) {
  #zd-weekly-report-controls {
    grid-template-columns: 1fr;
  }
}
`;

  const style = document.createElement("style");
  style.textContent = css;
  document.head.appendChild(style);

  const ensureStylesheet = (href) => {
    const existing = document.querySelector(`link[href="${href}"]`);
    if (existing) return;
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = href;
    document.head.appendChild(link);
  };

  const loadScript = (src) =>
    new Promise((resolve, reject) => {
      const existing = document.querySelector(`script[src="${src}"]`);
      if (existing) {
        if (window.flatpickr) {
          resolve();
          return;
        }
        existing.addEventListener("load", () => resolve(), { once: true });
        existing.addEventListener("error", () => reject(new Error(`Failed to load script: ${src}`)), { once: true });
        return;
      }

      const script = document.createElement("script");
      script.src = src;
      script.async = true;
      script.addEventListener("load", () => resolve(), { once: true });
      script.addEventListener("error", () => reject(new Error(`Failed to load script: ${src}`)), { once: true });
      document.head.appendChild(script);
    });

  const ensureFlatpickr = async () => {
    if (window.flatpickr) {
      return window.flatpickr;
    }
    ensureStylesheet(FLATPICKR_CSS_URL);
    await loadScript(FLATPICKR_JS_URL);
    if (!window.flatpickr) {
      throw new Error("Flatpickr loaded but unavailable on window.");
    }
    return window.flatpickr;
  };

  const launcher = document.createElement("button");
  launcher.id = "zd-weekly-report-launcher";
  launcher.textContent = "ZD Reports";

  const overlay = document.createElement("div");
  overlay.id = "zd-weekly-report-overlay";

  const getWeekStartDate = (date, weekStartsOn) => {
    const source = date instanceof Date && !Number.isNaN(date.getTime()) ? new Date(date) : new Date();
    source.setHours(0, 0, 0, 0);
    const localDay = source.getDay();
    const diffFromWeekStart = (localDay - weekStartsOn + 7) % 7;
    source.setDate(source.getDate() - diffFromWeekStart);
    return source;
  };

  const modal = document.createElement("div");
  modal.id = "zd-weekly-report-modal";
  const initialWeekDate = getWeekStartDate(new Date(), FIXED_WEEK_STARTS_ON);
  const todayIso = initialWeekDate.toISOString().slice(0, 10);

  modal.innerHTML = `
    <div id="zd-weekly-report-header">
      <div>
        <h3>Zendesk Weekly Reports</h3>
        <div class="subtext">Run multi-report summaries with one click</div>
      </div>
      <button id="zd-weekly-report-close" type="button">Close</button>
    </div>
    <div id="zd-weekly-report-controls">
      <label>
        Week containing date (Sunday start)
        <input id="zd-week-date" type="text" value="${todayIso}" placeholder="YYYY-MM-DD" autocomplete="off">
      </label>
      <div class="actions">
        <button id="zd-run-reports" type="button">Run reports</button>
        <button id="zd-copy-report" type="button">Copy Report</button>
        <button id="zd-copy-output" type="button">Copy JSON</button>
      </div>
    </div>
    <div id="zd-weekly-report-status">Ready.</div>
    <div id="zd-weekly-report-output"><div class="zd-placeholder">No report run yet.</div></div>
  `;

  document.body.appendChild(overlay);
  document.body.appendChild(modal);
  document.body.appendChild(launcher);

  const closeButton = modal.querySelector("#zd-weekly-report-close");
  const runButton = modal.querySelector("#zd-run-reports");
  const copyReportButton = modal.querySelector("#zd-copy-report");
  const copyButton = modal.querySelector("#zd-copy-output");

  const weekDateInput = modal.querySelector("#zd-week-date");
  const statusEl = modal.querySelector("#zd-weekly-report-status");
  const outputEl = modal.querySelector("#zd-weekly-report-output");

  let latestRawOutput = "";
  let latestReportOutput = "";
  let flatpickrInstance = null;
  let isSnappingDate = false;
  let suppressLauncherClick = false;
  let assigneeOverrideKeyword = null;

  const clamp = (value, min, max) => Math.min(Math.max(value, min), max);

  const applyLauncherPosition = ({ left, top } = {}) => {
    if (typeof left === "number") {
      launcher.style.left = `${left}px`;
      launcher.style.right = "auto";
    }
    if (typeof top === "number") {
      launcher.style.top = `${top}px`;
      launcher.style.bottom = "auto";
    }
  };

  const loadLauncherPosition = () => {
    try {
      const raw = localStorage.getItem(LAUNCHER_POSITION_STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      const left = Number(parsed?.left);
      const top = Number(parsed?.top);
      if (Number.isFinite(left) && Number.isFinite(top)) {
        applyLauncherPosition({ left, top });
      }
    } catch {
      // ignore malformed persisted launcher state
    }
  };

  const saveLauncherPosition = ({ left, top }) => {
    try {
      localStorage.setItem(LAUNCHER_POSITION_STORAGE_KEY, JSON.stringify({ left, top }));
    } catch {
      // ignore storage failures
    }
  };

  const clampLauncherToViewport = ({ persist = false } = {}) => {
    const rect = launcher.getBoundingClientRect();
    const maxLeft = Math.max(8, window.innerWidth - rect.width - 8);
    const maxTop = Math.max(8, window.innerHeight - rect.height - 8);
    const left = clamp(rect.left, 8, maxLeft);
    const top = clamp(rect.top, 8, maxTop);
    const moved = Math.abs(left - rect.left) > 0.5 || Math.abs(top - rect.top) > 0.5;

    if (!moved) {
      return;
    }

    applyLauncherPosition({ left, top });
    if (persist) {
      saveLauncherPosition({ left, top });
    }
  };

  const showModal = () => {
    overlay.style.display = "block";
    modal.style.display = "block";
  };

  const hideModal = () => {
    overlay.style.display = "none";
    modal.style.display = "none";
  };

  launcher.addEventListener("click", () => {
    if (suppressLauncherClick) {
      suppressLauncherClick = false;
      return;
    }
    showModal();
  });

  launcher.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) {
      return;
    }

    const rect = launcher.getBoundingClientRect();
    const offsetX = event.clientX - rect.left;
    const offsetY = event.clientY - rect.top;
    const startX = event.clientX;
    const startY = event.clientY;
    let moved = false;

    launcher.classList.add("dragging");
    launcher.setPointerCapture(event.pointerId);

    const onPointerMove = (moveEvent) => {
      const dx = Math.abs(moveEvent.clientX - startX);
      const dy = Math.abs(moveEvent.clientY - startY);
      if (!moved && (dx > 4 || dy > 4)) {
        moved = true;
      }
      if (!moved) {
        return;
      }

      const maxLeft = Math.max(0, window.innerWidth - rect.width - 8);
      const maxTop = Math.max(0, window.innerHeight - rect.height - 8);
      const left = clamp(moveEvent.clientX - offsetX, 8, maxLeft);
      const top = clamp(moveEvent.clientY - offsetY, 8, maxTop);
      applyLauncherPosition({ left, top });
    };

    const finishDrag = (upEvent) => {
      launcher.classList.remove("dragging");
      launcher.releasePointerCapture(upEvent.pointerId);
      launcher.removeEventListener("pointermove", onPointerMove);
      launcher.removeEventListener("pointerup", finishDrag);
      launcher.removeEventListener("pointercancel", finishDrag);

      if (moved) {
        const finalRect = launcher.getBoundingClientRect();
        saveLauncherPosition({ left: finalRect.left, top: finalRect.top });
        suppressLauncherClick = true;
      }
    };

    launcher.addEventListener("pointermove", onPointerMove);
    launcher.addEventListener("pointerup", finishDrag);
    launcher.addEventListener("pointercancel", finishDrag);
  });

  window.addEventListener("resize", () => {
    clampLauncherToViewport({ persist: true });
  });

  window.addEventListener("orientationchange", () => {
    clampLauncherToViewport({ persist: true });
  });

  closeButton.addEventListener("click", hideModal);
  overlay.addEventListener("click", hideModal);

  copyButton.addEventListener("click", async () => {
    try {
      const payload = latestRawOutput || outputEl.textContent || "";
      await navigator.clipboard.writeText(payload);
      statusEl.textContent = "Copied JSON payload to clipboard.";
    } catch (error) {
      statusEl.textContent = `Copy failed: ${error?.message || error}`;
    }
  });

  copyReportButton.addEventListener("click", async () => {
    try {
      if (!latestReportOutput) {
        statusEl.textContent = "Run reports first to generate report copy text.";
        return;
      }
      await navigator.clipboard.writeText(latestReportOutput);
      statusEl.textContent = "Copied report-ready summary to clipboard.";
    } catch (error) {
      statusEl.textContent = `Copy failed: ${error?.message || error}`;
    }
  });

  const computeStartOfWeek = (anchorDate) => {
    const base = anchorDate instanceof Date && !Number.isNaN(anchorDate.getTime()) ? anchorDate : new Date();
    const startOfWeek = getWeekStartDate(base, FIXED_WEEK_STARTS_ON);
    return {
      startOfWeek,
      startOfWeekDate: startOfWeek.toISOString().slice(0, 10),
      startOfWeekIso: startOfWeek.toISOString()
    };
  };

  const sleep = (milliseconds) =>
    new Promise((resolve) => {
      window.setTimeout(resolve, milliseconds);
    });

  const buildHeaders = () => {
    const token = document.querySelector('meta[name="csrf-token"]')?.getAttribute("content") || "";
    return {
      accept: "application/json, text/javascript, */*; q=0.01",
      "x-csrf-token": token,
      "x-requested-with": "XMLHttpRequest"
    };
  };

  const fetchJson = async (url) => {
    const response = await fetch(url, {
      method: "GET",
      headers: buildHeaders(),
      mode: "cors",
      credentials: "include"
    });
    const text = await response.text();
    let data = null;
    try {
      data = JSON.parse(text);
    } catch {
      data = null;
    }
    return { response, text, data, url };
  };

  const getCurrentUserId = async () => {
    const { response, data, url } = await fetchJson(`${BASE}/api/v2/users/me.json`);
    if (!response.ok) {
      throw new Error(`Unable to fetch current user (${response.status}) at ${url}`);
    }
    const userId = data?.user?.id;
    if (!userId) {
      throw new Error("Current user ID not found in /users/me response.");
    }
    return userId;
  };

  const resolveUserIdFromAssigneeInput = async (assigneeKeyword) => {
    if (!assigneeKeyword || assigneeKeyword === "me") {
      return getCurrentUserId();
    }

    if (/^\d+$/.test(assigneeKeyword)) {
      return Number(assigneeKeyword);
    }

    const query = `type:user ${assigneeKeyword}`;
    const params = new URLSearchParams({ query });
    const { response, data, url } = await fetchJson(`${BASE}/api/v2/search.json?${params.toString()}`);
    if (!response.ok) {
      throw new Error(`Unable to resolve assignee user (${response.status}) at ${url}`);
    }

    const users = Array.isArray(data?.results)
      ? data.results.filter((result) => result?.result_type === "user")
      : [];

    if (!users.length || !users[0]?.id) {
      throw new Error(`No user found for assignee input: ${assigneeKeyword}`);
    }

    return users[0].id;
  };

  const searchAllTickets = async ({ query, sortBy = "updated_at", sortOrder = "desc", onProgress }) => {
    const params = new URLSearchParams({
      query,
      sort_by: sortBy,
      sort_order: sortOrder
    });

    let nextUrl = `${BASE}/api/v2/search.json?${params.toString()}`;
    const tickets = [];
    let page = 0;

    while (nextUrl) {
      page += 1;
      onProgress?.(`Search page ${page}...`);
      const { response, data, url } = await fetchJson(nextUrl);

      if (!response.ok) {
        throw new Error(`Search failed on page ${page} (${response.status}) at ${url}`);
      }

      const pageResults = Array.isArray(data?.results) ? data.results : [];
      tickets.push(...pageResults.filter((result) => result?.result_type === "ticket"));
      nextUrl = data?.next_page || null;
    }

    return tickets;
  };

  const chunkArray = (values = [], chunkSize = 100) => {
    const chunks = [];
    for (let i = 0; i < values.length; i += chunkSize) {
      chunks.push(values.slice(i, i + chunkSize));
    }
    return chunks;
  };

  const uniqueNumericIds = (values = []) => {
    const seen = new Set();
    const result = [];

    for (const value of values) {
      const asNumber = Number(value);
      if (!Number.isFinite(asNumber) || asNumber <= 0) {
        continue;
      }
      if (seen.has(asNumber)) {
        continue;
      }
      seen.add(asNumber);
      result.push(asNumber);
    }

    return result;
  };

  const fetchUsersByIds = async (userIds = [], onProgress) => {
    const ids = uniqueNumericIds(userIds);
    const userMap = new Map();

    if (!ids.length) {
      return userMap;
    }

    const chunks = chunkArray(ids, 100);
    for (let i = 0; i < chunks.length; i += 1) {
      const chunk = chunks[i];
      onProgress?.(`Resolving requester names ${i + 1}/${chunks.length}...`);
      const params = new URLSearchParams({ ids: chunk.join(",") });
      const { response, data, url } = await fetchJson(`${BASE}/api/v2/users/show_many.json?${params.toString()}`);

      if (!response.ok) {
        throw new Error(`Requester name lookup failed (${response.status}) at ${url}`);
      }

      const users = Array.isArray(data?.users) ? data.users : [];
      for (const user of users) {
        const id = user?.id;
        if (!id) {
          continue;
        }
        const label = String(user?.name || user?.email || id);
        userMap.set(Number(id), label);
      }
    }

    return userMap;
  };

  const fetchOrganizationsByIds = async (organizationIds = [], onProgress) => {
    const ids = uniqueNumericIds(organizationIds);
    const organizationMap = new Map();

    if (!ids.length) {
      return organizationMap;
    }

    const chunks = chunkArray(ids, 100);
    for (let i = 0; i < chunks.length; i += 1) {
      const chunk = chunks[i];
      onProgress?.(`Resolving org names ${i + 1}/${chunks.length}...`);
      const params = new URLSearchParams({ ids: chunk.join(",") });
      const { response, data, url } = await fetchJson(`${BASE}/api/v2/organizations/show_many.json?${params.toString()}`);

      if (!response.ok) {
        throw new Error(`Organization name lookup failed (${response.status}) at ${url}`);
      }

      const organizations = Array.isArray(data?.organizations) ? data.organizations : [];
      for (const organization of organizations) {
        const id = organization?.id;
        if (!id) {
          continue;
        }
        const label = String(organization?.name || id);
        organizationMap.set(Number(id), label);
      }
    }

    return organizationMap;
  };

  const enrichRowsWithNames = (rows = [], organizationMap = new Map(), userMap = new Map()) => {
    if (!Array.isArray(rows)) {
      return rows;
    }

    for (const row of rows) {
      const orgId = Number(row?.organization_id);
      const requesterId = Number(row?.requester_id);

      row.organization_name = Number.isFinite(orgId) ? organizationMap.get(orgId) || null : null;
      row.requester_name = Number.isFinite(requesterId) ? userMap.get(requesterId) || null : null;
    }

    return rows;
  };

  const enrichReportsWithNames = async (reports = [], onProgress) => {
    const allRows = reports.flatMap((report) => (Array.isArray(report?.ticket_rows) ? report.ticket_rows : []));
    const organizationIds = allRows.map((row) => row?.organization_id).filter((value) => value !== null && value !== undefined);
    const requesterIds = allRows.map((row) => row?.requester_id).filter((value) => value !== null && value !== undefined);

    const [organizationMap, userMap] = await Promise.all([
      fetchOrganizationsByIds(organizationIds, onProgress),
      fetchUsersByIds(requesterIds, onProgress)
    ]);

    for (const report of reports) {
      enrichRowsWithNames(report.ticket_rows, organizationMap, userMap);
    }
  };

  const pickTicketIdentity = (ticket) => ({
    ticket_id: ticket?.id ?? null,
    organization_id: ticket?.organization_id ?? null,
    requester_id: ticket?.requester_id ?? null
  });

  const runAssignedCreatedThisWeek = async ({ assigneeKeyword, startOfWeekDate, onProgress }) => {
    const query = `type:ticket assignee:${assigneeKeyword} created>=${startOfWeekDate}`;
    const tickets = await searchAllTickets({ query, onProgress });
    return {
      name: "Tickets Taken This Week",
      query,
      total: tickets.length,
      ticket_rows: tickets.map(pickTicketIdentity)
    };
  };

  const runAssignedSolvedUpdatedThisWeek = async ({ assigneeKeyword, startOfWeekDate, onProgress }) => {
    const query = `type:ticket assignee:${assigneeKeyword} status:solved updated>=${startOfWeekDate}`;
    const tickets = await searchAllTickets({ query, onProgress });
    return {
      name: "Assigned + Solved + Updated This Week",
      query,
      total: tickets.length,
      ticket_rows: tickets.map(pickTicketIdentity)
    };
  };

  const runOpenTicketsRemaining = async ({ assigneeKeyword, onProgress }) => {
    const query = `type:ticket assignee:${assigneeKeyword} status:open`;
    const tickets = await searchAllTickets({ query, onProgress });
    return {
      name: "Open Tickets Remaining",
      query,
      total: tickets.length,
      ticket_rows: tickets.map(pickTicketIdentity)
    };
  };

  const runCarriedOverTickets = async ({ assigneeKeyword, startOfWeekDate, onProgress }) => {
    const query = `type:ticket assignee:${assigneeKeyword} created<${startOfWeekDate} updated>=${startOfWeekDate} -status:solved -status:closed`;
    const tickets = await searchAllTickets({ query, onProgress });
    return {
      name: "Carried Over",
      query,
      total: tickets.length,
      ticket_rows: tickets.map(pickTicketIdentity)
    };
  };

  const runTakeoverReport = async ({ assigneeKeyword, myUserId, startOfWeekDate, startOfWeekIso, onProgress }) => {
    const query = `type:ticket assignee:${assigneeKeyword} updated>=${startOfWeekDate}`;
    const candidateTickets = await searchAllTickets({ query, onProgress });
    const limited = candidateTickets;

    const takeoverRows = [];

    for (let i = 0; i < limited.length; i += 1) {
      const ticket = limited[i];
      const ticketId = ticket?.id;
      if (!ticketId) {
        continue;
      }

      onProgress?.(`Audits ${i + 1}/${limited.length} (ticket ${ticketId})...`);
      let nextAuditUrl = `${BASE}/api/v2/tickets/${ticketId}/audits.json`;

      while (nextAuditUrl) {
        const { response, data, url } = await fetchJson(nextAuditUrl);

        if (!response.ok) {
          throw new Error(`Audit fetch failed (${response.status}) at ${url}`);
        }

        const audits = Array.isArray(data?.audits) ? data.audits : [];

        for (const audit of audits) {
          const auditCreatedAt = audit?.created_at;
          if (!auditCreatedAt || auditCreatedAt < startOfWeekIso) {
            continue;
          }

          const events = Array.isArray(audit?.events) ? audit.events : [];
          for (const event of events) {
            const hasPreviousAssignee = event?.previous_value !== null && event?.previous_value !== undefined && String(event?.previous_value).trim() !== "";
            const isTakeover =
              event?.type === "Change" &&
              event?.field_name === "assignee_id" &&
              hasPreviousAssignee &&
              String(event?.value) === String(myUserId) &&
              String(event?.previous_value) !== String(myUserId);

            if (isTakeover) {
              takeoverRows.push({
                ticket_id: ticketId,
                organization_id: ticket?.organization_id ?? null,
                requester_id: ticket?.requester_id ?? null,
                previous_assignee_id: event?.previous_value ?? null,
                new_assignee_id: event?.value ?? null,
                takeover_at: auditCreatedAt
              });
            }
          }
        }

        nextAuditUrl = data?.next_page || null;
      }

      const processed = i + 1;
      if (processed % TAKEOVER_BATCH_SIZE === 0 && processed < limited.length) {
        onProgress?.(`Pausing ${Math.round(TAKEOVER_BATCH_SLEEP_MS / 1000)}s after ${processed} audits...`);
        await sleep(TAKEOVER_BATCH_SLEEP_MS);
      }
    }

    const uniqueTicketIds = [...new Set(takeoverRows.map((row) => row.ticket_id))];

    return {
      name: "Takeovers This Week",
      query,
      candidate_total: candidateTickets.length,
      candidate_audited: limited.length,
      total_unique_tickets: uniqueTicketIds.length,
      total_takeover_events: takeoverRows.length,
      ticket_rows: takeoverRows
    };
  };

  const toTextBlock = (title, obj) => {
    return [`=== ${title} ===`, JSON.stringify(obj, null, 2), ""].join("\n");
  };

  const getAnchorDate = () => {
    const raw = String(weekDateInput?.value || "").trim();
    if (!raw) return new Date();
    const date = new Date(`${raw}T00:00:00`);
    return Number.isNaN(date.getTime()) ? new Date() : date;
  };

  const snapWeekDateInput = () => {
    const snapped = getWeekStartDate(getAnchorDate(), FIXED_WEEK_STARTS_ON);
    const snappedIso = snapped.toISOString().slice(0, 10);

    if (flatpickrInstance) {
      isSnappingDate = true;
      flatpickrInstance.setDate(snapped, true, "Y-m-d");
      isSnappingDate = false;
    } else {
      weekDateInput.value = snappedIso;
    }
  };

  const initDatePicker = async () => {
    try {
      const flatpickr = await ensureFlatpickr();
      flatpickrInstance = flatpickr(weekDateInput, {
        dateFormat: "Y-m-d",
        defaultDate: todayIso,
        allowInput: true,
        disableMobile: true,
        locale: {
          firstDayOfWeek: FIXED_WEEK_STARTS_ON
        },
        onChange: (selectedDates) => {
          if (isSnappingDate || !selectedDates?.length) {
            return;
          }
          const snapped = getWeekStartDate(selectedDates[0], FIXED_WEEK_STARTS_ON);
          isSnappingDate = true;
          flatpickrInstance.setDate(snapped, true, "Y-m-d");
          isSnappingDate = false;
        },
        onClose: () => {
          snapWeekDateInput();
        }
      });

      snapWeekDateInput();
      statusEl.textContent = "Ready.";
    } catch (error) {
      flatpickrInstance = null;
      weekDateInput.type = "date";
      weekDateInput.value = todayIso;
      statusEl.textContent = "Ready. Date picker fallback active.";
      console.warn("Flatpickr failed to load; using native date input.", error);
    }
  };

  window.zdWeeklyReportPopup = {
    ...(window.zdWeeklyReportPopup || {}),
    setAssignee: (keyword) => {
      const normalized = String(keyword ?? "").trim();
      assigneeOverrideKeyword = normalized || null;
      const effective = assigneeOverrideKeyword || ASSIGNEE_FALLBACK_KEYWORD;
      statusEl.textContent = `Assignee override set to: ${effective}`;
      return effective;
    },
    clearAssignee: () => {
      assigneeOverrideKeyword = null;
      statusEl.textContent = `Assignee override cleared. Using ${ASSIGNEE_FALLBACK_KEYWORD}.`;
      return ASSIGNEE_FALLBACK_KEYWORD;
    },
    getAssignee: () => assigneeOverrideKeyword || ASSIGNEE_FALLBACK_KEYWORD
  };

  const renderTicketChips = (rows = []) => {
    if (!Array.isArray(rows) || rows.length === 0) {
      return '<div class="zd-subline">No ticket IDs in range.</div>';
    }
    const sample = rows.slice(0, 24);
    const chips = sample
      .map((row) => {
        const ticketId = row.ticket_id;
        if (!ticketId) {
          return `<span class="zd-chip">#?</span>`;
        }
        const href = `${BASE}/agent/tickets/${ticketId}`;
        return `<a class="zd-chip" href="${href}" target="_blank" rel="noopener noreferrer">#${ticketId}</a>`;
      })
      .join("");
    const more = rows.length > sample.length ? `<span class="zd-subline">+${rows.length - sample.length} more</span>` : "";
    return `<div class="zd-chip-row">${chips}${more}</div>`;
  };

  const renderReportCard = (report, options = {}) => {
    const subtitle = options.subtitle || "";
    const extra = options.extra || [];
    const chips = renderTicketChips(report.ticket_rows);
    const extras = extra
      .map((item) => `<div class="zd-subline">${item}</div>`)
      .join("");
    return `
      <div class="zd-card">
        <h4>${report.name}</h4>
        <div class="zd-metric">${report.total ?? 0}</div>
        <div class="zd-subline">${subtitle}</div>
        ${extras}
        <div class="zd-query">${report.query}</div>
        ${chips}
      </div>
    `;
  };

  const renderTakeoverCard = (report) => {
    const chips = renderTicketChips(report.ticket_rows);
    return `
      <div class="zd-card">
        <h4>${report.name}</h4>
        <div class="zd-metric">${report.total_unique_tickets ?? 0}</div>
        <div class="zd-subline">Unique tickets with takeovers</div>
        <div class="zd-subline">Events: ${report.total_takeover_events ?? 0}</div>
        <div class="zd-subline">Audited ${report.candidate_audited ?? 0} of ${report.candidate_total ?? 0} candidates</div>
        <div class="zd-query">${report.query}</div>
        ${chips}
      </div>
    `;
  };

  const renderTable = (rows = [], title) => {
    if (!rows.length) return "";

    const toSafeText = (value) => {
      const text = value === null || value === undefined ? "" : String(value);
      return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\"/g, "&quot;")
        .replace(/'/g, "&#39;");
    };

    const formatEntity = (name, id) => {
      const idText = id === null || id === undefined ? "" : String(id);
      const nameText = name === null || name === undefined || String(name).trim() === "" ? "" : String(name);

      if (nameText && idText) {
        return `${toSafeText(nameText)} <span class="zd-subline">(${toSafeText(idText)})</span>`;
      }
      if (nameText) {
        return toSafeText(nameText);
      }
      return toSafeText(idText);
    };

    const limited = rows.slice(0, 25);
    const header = `
      <tr>
        <th>Ticket</th>
        <th>Organization</th>
        <th>Requester</th>
      </tr>`;
    const body = limited
      .map(
        (row) => `
        <tr>
          <td>${toSafeText(row.ticket_id ?? "")}</td>
          <td>${formatEntity(row.organization_name, row.organization_id)}</td>
          <td>${formatEntity(row.requester_name, row.requester_id)}</td>
        </tr>`
      )
      .join("");
    const more = rows.length > limited.length ? `<div class="zd-subline">Showing first ${limited.length} of ${rows.length}</div>` : "";
    return `
      <div class="zd-section-title">${title}</div>
      <table class="zd-table">${header}${body}</table>
      ${more}
    `;
  };

  const renderEntityCountTable = ({
    rows = [],
    title = "",
    nameKey = "organization_name",
    idKey = "organization_id",
    label = "Entity"
  } = {}) => {
    if (!rows.length) return "";

    const toSafeText = (value) => {
      const text = value === null || value === undefined ? "" : String(value);
      return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\"/g, "&quot;")
        .replace(/'/g, "&#39;");
    };

    const counter = new Map();

    for (const row of rows) {
      const rawName = row?.[nameKey];
      const rawId = row?.[idKey];
      const idText = rawId === null || rawId === undefined || String(rawId).trim() === "" ? "Unknown" : String(rawId);
      const nameText = rawName === null || rawName === undefined || String(rawName).trim() === "" ? "Unknown" : String(rawName);
      const key = `${idText}::${nameText}`;
      const existing = counter.get(key) || { idText, nameText, count: 0 };
      existing.count += 1;
      counter.set(key, existing);
    }

    const sorted = [...counter.values()].sort((a, b) => b.count - a.count || a.nameText.localeCompare(b.nameText));
    const limited = sorted.slice(0, 50);

    const header = `
      <tr>
        <th>${label}</th>
        <th>Count</th>
      </tr>`;
    const body = limited
      .map(
        (entry) => `
        <tr>
          <td>${toSafeText(entry.nameText)} <span class="zd-subline">(${toSafeText(entry.idText)})</span></td>
          <td>${entry.count}</td>
        </tr>`
      )
      .join("");
    const more = sorted.length > limited.length ? `<div class="zd-subline">Showing top ${limited.length} of ${sorted.length}</div>` : "";

    return `
      <div class="zd-section-title">${title}</div>
      <table class="zd-table">${header}${body}</table>
      ${more}
    `;
  };

  const asIdList = (rows = [], limit = 40) => {
    if (!Array.isArray(rows) || rows.length === 0) return "None";
    const ids = rows.slice(0, limit).map((row) => `#${row.ticket_id ?? "?"}`);
    if (rows.length > limit) {
      ids.push(`... (+${rows.length - limit} more)`);
    }
    return ids.join(", ");
  };

  const uniqueRowsByTicketId = (rows = []) => {
    const seen = new Set();
    const unique = [];

    for (const row of rows) {
      const ticketId = row?.ticket_id;
      if (ticketId === null || ticketId === undefined || String(ticketId).trim() === "") {
        continue;
      }
      const key = String(ticketId);
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      unique.push(row);
    }

    return unique;
  };

  const buildReportDraft = ({ meta, report1, report2, reportOpen, reportCarry, report3 }) => {
    const takeoverRows = Array.isArray(report3.ticket_rows) ? report3.ticket_rows : [];
    const takeoverPreview = takeoverRows
      .slice(0, 20)
      .map((row) => `- #${row.ticket_id ?? "?"}: ${row.previous_assignee_id ?? "?"} -> ${row.new_assignee_id ?? "?"} at ${row.takeover_at ?? "?"}`)
      .join("\n");
    const takeoverMore = takeoverRows.length > 20 ? `\n- ... (+${takeoverRows.length - 20} more takeover events)` : "";

    return [
      "Zendesk Weekly Report",
      `Generated: ${meta.generated_at}`,
      "",
      "Summary",
      `- Tickets Taken This Week: ${report1.total}`,
      `- Assigned + Solved + Updated This Week: ${report2.total}`,
      `- Open Tickets Remaining: ${reportOpen.total}`,
      `- Carried Over: ${reportCarry.total}`,
      `- Takeovers This Week: ${report3.total_unique_tickets} unique tickets (${report3.total_takeover_events} events)` ,
      `- Takeover audits coverage: ${report3.candidate_audited}/${report3.candidate_total}`,
      "",
      "Ticket IDs",
      `- Taken: ${asIdList(report1.ticket_rows)}`,
      `- Solved/Updated: ${asIdList(report2.ticket_rows)}`,
      `- Open Remaining: ${asIdList(reportOpen.ticket_rows)}`,
      `- Carried Over: ${asIdList(reportCarry.ticket_rows)}`,
      `- Takeover Tickets: ${asIdList(report3.ticket_rows)}`,
      "",
      "Takeover Events (sample)",
      takeoverPreview || "- None",
      takeoverMore
    ].join("\n");
  };

  const renderOutput = ({ meta, report1, report2, reportOpen, reportCarry, report3 }) => {
    const reportDraft = buildReportDraft({ meta, report1, report2, reportOpen, reportCarry, report3 });
    const aggregateRows = uniqueRowsByTicketId([
      ...(Array.isArray(report1.ticket_rows) ? report1.ticket_rows : []),
      ...(Array.isArray(report2.ticket_rows) ? report2.ticket_rows : []),
      ...(Array.isArray(reportOpen.ticket_rows) ? reportOpen.ticket_rows : []),
      ...(Array.isArray(reportCarry.ticket_rows) ? reportCarry.ticket_rows : []),
      ...(Array.isArray(report3.ticket_rows) ? report3.ticket_rows : [])
    ]);

    const summaryCards = `
      <div class="zd-card-grid">
        ${renderReportCard(report1, { subtitle: `Taken since ${meta.start_of_week_date}` })}
        ${renderReportCard(report2, { subtitle: `Solved & updated since ${meta.start_of_week_date}` })}
        ${renderReportCard(reportOpen, { subtitle: "Currently open" })}
        ${renderReportCard(reportCarry, { subtitle: `Created before, updated since ${meta.start_of_week_date}, not solved` })}
        ${renderTakeoverCard(report3)}
      </div>
    `;

    const tables = `
      ${renderTable(report1.ticket_rows, report1.name)}
      ${renderTable(report2.ticket_rows, report2.name)}
      ${renderTable(reportOpen.ticket_rows, reportOpen.name)}
      ${renderTable(reportCarry.ticket_rows, reportCarry.name)}
      ${renderTable(report3.ticket_rows, `${report3.name} (events)`) }
      ${renderEntityCountTable({ rows: aggregateRows, title: "Report Count by Organization", nameKey: "organization_name", idKey: "organization_id", label: "Organization" })}
      ${renderEntityCountTable({ rows: aggregateRows, title: "Report Count by Requester", nameKey: "requester_name", idKey: "requester_id", label: "Requester" })}
      <div class="zd-report-draft">
        <h4>Copy-ready report draft</h4>
        <textarea id="zd-report-draft-text" readonly></textarea>
      </div>
    `;

    outputEl.innerHTML = `
      <div class="zd-meta-card">
        <div class="zd-section-title">Run Meta</div>
        <div class="zd-meta-grid">
          <div><span>Generated</span><strong>${meta.generated_at}</strong></div>
        </div>
      </div>
      ${summaryCards}
      ${tables}
    `;

    const draftTextarea = outputEl.querySelector("#zd-report-draft-text");
    if (draftTextarea) {
      draftTextarea.value = reportDraft;
    }

    latestReportOutput = reportDraft;
  };

  runButton.addEventListener("click", async () => {
    runButton.disabled = true;
    outputEl.innerHTML = '<div class="zd-placeholder">Running...</div>';

    const assigneeKeyword = assigneeOverrideKeyword || ASSIGNEE_FALLBACK_KEYWORD;
    const anchorDate = getAnchorDate();
    const { startOfWeekDate, startOfWeekIso } = computeStartOfWeek(anchorDate);

    const progress = (msg) => {
      statusEl.textContent = msg;
    };

    try {
      progress("Resolving target assignee user...");
      const myUserId = await resolveUserIdFromAssigneeInput(assigneeKeyword);

      progress("Running report 1/5...");
      const report1 = await runAssignedCreatedThisWeek({ assigneeKeyword, startOfWeekDate, onProgress: progress });

      progress("Running report 2/5...");
      const report2 = await runAssignedSolvedUpdatedThisWeek({ assigneeKeyword, startOfWeekDate, onProgress: progress });

      progress("Running report 3/5...");
      const reportOpen = await runOpenTicketsRemaining({ assigneeKeyword, onProgress: progress });

      progress("Running report 4/5...");
      const reportCarry = await runCarriedOverTickets({ assigneeKeyword, startOfWeekDate, onProgress: progress });

      progress("Running report 5/5 (audits)...");
      const report3 = await runTakeoverReport({
        assigneeKeyword,
        myUserId,
        startOfWeekDate,
        startOfWeekIso,
        onProgress: progress
      });

      try {
        progress("Resolving organization/requester names...");
        await enrichReportsWithNames([report1, report2, reportOpen, reportCarry, report3], progress);
      } catch (nameError) {
        console.warn("Name resolution failed; continuing with IDs only.", nameError);
        progress("Name resolution partially failed; continuing with IDs.");
      }

      const meta = {
        assignee: assigneeKeyword,
        assignee_source: assigneeOverrideKeyword ? "override" : "default (me)",
        current_user_id: myUserId,
        week_starts_on: "Sunday",
        start_of_week_date: startOfWeekDate,
        start_of_week_iso: startOfWeekIso,
        anchor_date: anchorDate.toISOString().slice(0, 10),
        takeover_batch_size: TAKEOVER_BATCH_SIZE,
        takeover_pause_seconds: Math.round(TAKEOVER_BATCH_SLEEP_MS / 1000),
        generated_at: new Date().toISOString()
      };

      renderOutput({ meta, report1, report2, reportOpen, reportCarry, report3 });

      latestRawOutput = [
        toTextBlock("Run Meta", meta),
        toTextBlock(report1.name, report1),
        toTextBlock(report2.name, report2),
        toTextBlock(reportOpen.name, reportOpen),
        toTextBlock(reportCarry.name, reportCarry),
        toTextBlock(report3.name, report3)
      ].join("\n");

      statusEl.textContent = "Done. UI shows summaries; Copy JSON for full raw payload.";
      console.log("Zendesk weekly reports:", { meta, report1, report2, reportOpen, reportCarry, report3 });
    } catch (error) {
      outputEl.innerHTML = `<div class="zd-placeholder">Error: ${error?.message || error}</div>`;
      statusEl.textContent = "Failed.";
      console.error("Zendesk weekly report popup failed:", error);
    } finally {
      runButton.disabled = false;
    }
  });

  initDatePicker();
  loadLauncherPosition();
  clampLauncherToViewport({ persist: true });
})();
