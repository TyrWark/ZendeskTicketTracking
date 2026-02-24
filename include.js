// ==UserScript==
// @name         Zendesk Weekly Report
// @namespace    http://tampermonkey.net/
// @version      0.1.0
// @description  Run weekly Zendesk ticket reports (assigned, solved, and takeovers) from a popup.
// @author       Ty Wark
// @match        https://retail-support.zendesk.com/agent/*
// @grant        none
// ==/UserScript==

(function () {
  "use strict";

  const DEFAULT_ASSIGNEE_INPUT = "";
  const ASSIGNEE_FALLBACK_KEYWORD = "me";
  const DEFAULT_WEEK_STARTS_ON = 1; // 1 = Monday, 0 = Sunday
  const DEFAULT_MAX_TICKETS_TO_AUDIT = 200;

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
  bottom: 16px;
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
}
#zd-weekly-report-launcher:hover {
  transform: translateY(-1px);
  transition: transform 120ms ease, box-shadow 120ms ease;
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
  grid-template-columns: repeat(5, minmax(0, 1fr));
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
    grid-template-columns: repeat(2, minmax(0, 1fr));
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
  const initialWeekDate = getWeekStartDate(new Date(), DEFAULT_WEEK_STARTS_ON);
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
        Assignee (optional)
        <input id="zd-assignee" type="text" value="${DEFAULT_ASSIGNEE_INPUT}" placeholder="me (blank), user id, email, or name">
      </label>
      <label>
        Week starts on
        <select id="zd-week-start">
          <option value="1" selected>Monday</option>
          <option value="0">Sunday</option>
        </select>
      </label>
      <label>
        Week containing date
        <input id="zd-week-date" type="text" value="${todayIso}" placeholder="YYYY-MM-DD" autocomplete="off">
      </label>
      <label>
        Max audits tickets
        <input id="zd-max-audits" type="number" min="1" value="${DEFAULT_MAX_TICKETS_TO_AUDIT}">
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

  const assigneeInput = modal.querySelector("#zd-assignee");
  const weekStartSelect = modal.querySelector("#zd-week-start");
  const weekDateInput = modal.querySelector("#zd-week-date");
  const maxAuditsInput = modal.querySelector("#zd-max-audits");
  const statusEl = modal.querySelector("#zd-weekly-report-status");
  const outputEl = modal.querySelector("#zd-weekly-report-output");

  let latestRawOutput = "";
  let latestReportOutput = "";
  let flatpickrInstance = null;
  let isSnappingDate = false;

  const showModal = () => {
    overlay.style.display = "block";
    modal.style.display = "block";
  };

  const hideModal = () => {
    overlay.style.display = "none";
    modal.style.display = "none";
  };

  launcher.addEventListener("click", showModal);
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

  const computeStartOfWeek = (weekStartsOn, anchorDate) => {
    const base = anchorDate instanceof Date && !Number.isNaN(anchorDate.getTime()) ? anchorDate : new Date();
    const startOfWeek = getWeekStartDate(base, weekStartsOn);
    return {
      startOfWeek,
      startOfWeekDate: startOfWeek.toISOString().slice(0, 10),
      startOfWeekIso: startOfWeek.toISOString()
    };
  };

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

  const pickTicketIdentity = (ticket) => ({
    ticket_id: ticket?.id ?? null,
    organization_id: ticket?.organization_id ?? null,
    requester_id: ticket?.requester_id ?? null
  });

  const runAssignedCreatedThisWeek = async ({ assigneeKeyword, startOfWeekDate, onProgress }) => {
    const query = `type:ticket assignee:${assigneeKeyword} created>=${startOfWeekDate}`;
    const tickets = await searchAllTickets({ query, onProgress });
    return {
      name: "Assigned + Created This Week",
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

  const runTakeoverReport = async ({ assigneeKeyword, myUserId, startOfWeekDate, startOfWeekIso, maxTicketsToAudit, onProgress }) => {
    const query = `type:ticket assignee:${assigneeKeyword} updated>=${startOfWeekDate}`;
    const candidateTickets = await searchAllTickets({ query, onProgress });
    const limited = candidateTickets.slice(0, maxTicketsToAudit);

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
    const weekStartsOn = Number(weekStartSelect.value || DEFAULT_WEEK_STARTS_ON);
    const snapped = getWeekStartDate(getAnchorDate(), weekStartsOn);
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
          firstDayOfWeek: Number(weekStartSelect.value || DEFAULT_WEEK_STARTS_ON)
        },
        onChange: (selectedDates) => {
          if (isSnappingDate || !selectedDates?.length) {
            return;
          }
          const weekStartsOn = Number(weekStartSelect.value || DEFAULT_WEEK_STARTS_ON);
          const snapped = getWeekStartDate(selectedDates[0], weekStartsOn);
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

  weekStartSelect.addEventListener("change", () => {
    if (flatpickrInstance) {
      flatpickrInstance.set("locale", {
        firstDayOfWeek: Number(weekStartSelect.value || DEFAULT_WEEK_STARTS_ON)
      });
    }
    snapWeekDateInput();
  });

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
    const limited = rows.slice(0, 25);
    const header = `
      <tr>
        <th>Ticket</th>
        <th>Org</th>
        <th>Requester</th>
      </tr>`;
    const body = limited
      .map(
        (row) => `
        <tr>
          <td>${row.ticket_id ?? ""}</td>
          <td>${row.organization_id ?? ""}</td>
          <td>${row.requester_id ?? ""}</td>
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

  const asIdList = (rows = [], limit = 40) => {
    if (!Array.isArray(rows) || rows.length === 0) return "None";
    const ids = rows.slice(0, limit).map((row) => `#${row.ticket_id ?? "?"}`);
    if (rows.length > limit) {
      ids.push(`... (+${rows.length - limit} more)`);
    }
    return ids.join(", ");
  };

  const buildReportDraft = ({ meta, report1, report2, reportOpen, report3 }) => {
    const takeoverRows = Array.isArray(report3.ticket_rows) ? report3.ticket_rows : [];
    const takeoverPreview = takeoverRows
      .slice(0, 20)
      .map((row) => `- #${row.ticket_id ?? "?"}: ${row.previous_assignee_id ?? "?"} -> ${row.new_assignee_id ?? "?"} at ${row.takeover_at ?? "?"}`)
      .join("\n");
    const takeoverMore = takeoverRows.length > 20 ? `\n- ... (+${takeoverRows.length - 20} more takeover events)` : "";

    return [
      "Zendesk Weekly Report",
      `Generated: ${meta.generated_at}`,
      `Assignee: ${meta.assignee_input || "(blank -> me)"} [keyword: ${meta.assignee}]`,
      `Week anchor date: ${meta.anchor_date}`,
      `Week starts on: ${meta.week_starts_on}`,
      `Range start: ${meta.start_of_week_date}`,
      "",
      "Summary",
      `- Assigned + Created This Week: ${report1.total}`,
      `- Assigned + Solved + Updated This Week: ${report2.total}`,
      `- Open Tickets Remaining: ${reportOpen.total}`,
      `- Takeovers This Week: ${report3.total_unique_tickets} unique tickets (${report3.total_takeover_events} events)` ,
      `- Takeover audits coverage: ${report3.candidate_audited}/${report3.candidate_total}`,
      "",
      "Ticket IDs",
      `- Created: ${asIdList(report1.ticket_rows)}`,
      `- Solved/Updated: ${asIdList(report2.ticket_rows)}`,
      `- Open Remaining: ${asIdList(reportOpen.ticket_rows)}`,
      `- Takeover Tickets: ${asIdList(report3.ticket_rows)}`,
      "",
      "Takeover Events (sample)",
      takeoverPreview || "- None",
      takeoverMore
    ].join("\n");
  };

  const renderOutput = ({ meta, report1, report2, reportOpen, report3 }) => {
    const reportDraft = buildReportDraft({ meta, report1, report2, reportOpen, report3 });

    const summaryCards = `
      <div class="zd-card-grid">
        ${renderReportCard(report1, { subtitle: `Created since ${meta.start_of_week_date}` })}
        ${renderReportCard(report2, { subtitle: `Solved & updated since ${meta.start_of_week_date}` })}
        ${renderReportCard(reportOpen, { subtitle: "Currently open" })}
        ${renderTakeoverCard(report3)}
      </div>
    `;

    const tables = `
      ${renderTable(report1.ticket_rows, report1.name)}
      ${renderTable(report2.ticket_rows, report2.name)}
      ${renderTable(reportOpen.ticket_rows, reportOpen.name)}
      ${renderTable(report3.ticket_rows, `${report3.name} (events)`) }
      <div class="zd-report-draft">
        <h4>Copy-ready report draft</h4>
        <textarea id="zd-report-draft-text" readonly></textarea>
      </div>
    `;

    outputEl.innerHTML = `
      <div class="zd-meta-card">
        <div class="zd-section-title">Run Meta</div>
        <div class="zd-meta-grid">
          <div><span>Assignee keyword</span><strong>${meta.assignee}</strong></div>
          <div><span>Assignee input</span><strong>${meta.assignee_input || "(blank → me)"}</strong></div>
          <div><span>Week starts on</span><strong>${meta.week_starts_on}</strong></div>
          <div><span>Week anchor date</span><strong>${meta.anchor_date}</strong></div>
          <div><span>Start of week</span><strong>${meta.start_of_week_date}</strong></div>
          <div><span>Max audits</span><strong>${meta.max_tickets_to_audit}</strong></div>
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

    const weekStartsOn = Number(weekStartSelect.value || DEFAULT_WEEK_STARTS_ON);
    const maxTicketsToAudit = Number(maxAuditsInput.value || DEFAULT_MAX_TICKETS_TO_AUDIT);
    const assigneeRaw = String(assigneeInput.value || "").trim();
    const assigneeKeyword = assigneeRaw || ASSIGNEE_FALLBACK_KEYWORD;
    const anchorDate = getAnchorDate();
    const { startOfWeekDate, startOfWeekIso } = computeStartOfWeek(weekStartsOn, anchorDate);

    const progress = (msg) => {
      statusEl.textContent = msg;
    };

    try {
      progress("Resolving target assignee user...");
      const myUserId = await resolveUserIdFromAssigneeInput(assigneeKeyword);

      progress("Running report 1/4...");
      const report1 = await runAssignedCreatedThisWeek({ assigneeKeyword, startOfWeekDate, onProgress: progress });

      progress("Running report 2/4...");
      const report2 = await runAssignedSolvedUpdatedThisWeek({ assigneeKeyword, startOfWeekDate, onProgress: progress });

      progress("Running report 3/4...");
      const reportOpen = await runOpenTicketsRemaining({ assigneeKeyword, onProgress: progress });

      progress("Running report 4/4 (audits)...");
      const report3 = await runTakeoverReport({
        assigneeKeyword,
        myUserId,
        startOfWeekDate,
        startOfWeekIso,
        maxTicketsToAudit,
        onProgress: progress
      });

      const meta = {
        assignee: assigneeKeyword,
        assignee_input: assigneeRaw || null,
        current_user_id: myUserId,
        week_starts_on: weekStartsOn === 1 ? "Monday" : "Sunday",
        start_of_week_date: startOfWeekDate,
        start_of_week_iso: startOfWeekIso,
        anchor_date: anchorDate.toISOString().slice(0, 10),
        max_tickets_to_audit: maxTicketsToAudit,
        generated_at: new Date().toISOString()
      };

      renderOutput({ meta, report1, report2, reportOpen, report3 });

      latestRawOutput = [
        toTextBlock("Run Meta", meta),
        toTextBlock(report1.name, report1),
        toTextBlock(report2.name, report2),
        toTextBlock(reportOpen.name, reportOpen),
        toTextBlock(report3.name, report3)
      ].join("\n");

      statusEl.textContent = "Done. UI shows summaries; Copy JSON for full raw payload.";
      console.log("Zendesk weekly reports:", { meta, report1, report2, reportOpen, report3 });
    } catch (error) {
      outputEl.innerHTML = `<div class="zd-placeholder">Error: ${error?.message || error}</div>`;
      statusEl.textContent = "Failed.";
      console.error("Zendesk weekly report popup failed:", error);
    } finally {
      runButton.disabled = false;
    }
  });

  initDatePicker();
})();
