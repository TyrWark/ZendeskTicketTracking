// ==UserScript==
// @name         Zendesk Report Pro Serv Pulse
// @namespace    http://tampermonkey.net/
// @version      1.0.0
// @description  Silently run Zendesk weekly reports and push the payload to Pro Serv Pulse via postMessage.
// @match        https://retail-support.zendesk.com/agent/*
// @grant        none
// ==/UserScript==

/**
 * PROTOCOL — pro-serv-pulse side must:
 *
 *   Listen for the payload:
 *        window.addEventListener("message", (event) => {
 *          if (event.origin !== "https://retail-support.zendesk.com") return;
 *          if (event.data?.type === "zd-report-payload") {
 *            const { payload } = event.data; // full report data
 *          }
 *        });
 *
 *   The full payload is also written to localStorage under "zdWeeklyReportData" on the Zendesk origin
 *   as a secondary reference (readable by any TM script running on retail-support.zendesk.com).
 *
 * CONSOLE API — run in browser console to configure without a UI:
 *   window.zdPulseReport.setAssignee("username or ID")  — target a specific agent
 *   window.zdPulseReport.clearAssignee()                — revert to "me"
 *   window.zdPulseReport.setWeekDate("2026-05-18")      — target a specific week
 *   window.zdPulseReport.clearWeekDate()                — revert to current week
 *   window.zdPulseReport.run()                          — trigger programmatically
 */

(function () {
  "use strict";

  const ASSIGNEE_FALLBACK_KEYWORD = "me";
  const FIXED_WEEK_STARTS_ON = 0; // Sunday
  const TAKEOVER_BATCH_SIZE = 100;
  const TAKEOVER_BATCH_SLEEP_MS = 30_000;
  const PULSE_URL = "https://fleuriengraveson.github.io/pro-serv-pulse/";
  const PULSE_ORIGIN = "https://fleuriengraveson.github.io";
  const STORAGE_KEY = "zdWeeklyReportData";
  const LAUNCHER_POSITION_STORAGE_KEY = "zd-pulse-launcher-position";
  const BASE = "https://retail-support.zendesk.com";

  // ── Button styles ──────────────────────────────────────────────────────

  const style = document.createElement("style");
  style.textContent = `
#zd-pulse-launcher {
  position: fixed;
  right: 16px;
  top: 16px;
  z-index: 999999;
  background: linear-gradient(135deg, #0ea5e9, #f59e0b);
  color: #fff;
  border: none;
  border-radius: 999px;
  padding: 10px 16px;
  cursor: grab;
  font-size: 13px;
  font-weight: 600;
  letter-spacing: 0.2px;
  box-shadow: 0 8px 24px rgba(0,0,0,0.2);
  user-select: none;
  font-family: 'Helvetica Neue', sans-serif;
  transition: opacity 120ms ease;
  white-space: nowrap;
  min-width: 110px;
  text-align: center;
}
#zd-pulse-launcher:hover { opacity: 0.9; }
#zd-pulse-launcher.dragging { cursor: grabbing; }
#zd-pulse-launcher:disabled { cursor: not-allowed; opacity: 0.65; }
`;
  document.head.appendChild(style);

  // ── Button ─────────────────────────────────────────────────────────────

  const launcher = document.createElement("button");
  launcher.id = "zd-pulse-launcher";
  launcher.textContent = "ZD → Pulse";
  document.body.appendChild(launcher);

  // ── Button state ───────────────────────────────────────────────────────

  let statusResetTimer = null;

  const setButtonStatus = (text, { sticky = false } = {}) => {
    launcher.textContent = text;
    if (statusResetTimer) {
      clearTimeout(statusResetTimer);
      statusResetTimer = null;
    }
    if (!sticky) {
      statusResetTimer = setTimeout(() => {
        launcher.textContent = "ZD → Pulse";
        statusResetTimer = null;
      }, 4000);
    }
  };

  // ── Launcher drag ──────────────────────────────────────────────────────

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

  const saveLauncherPosition = ({ left, top }) => {
    try {
      localStorage.setItem(LAUNCHER_POSITION_STORAGE_KEY, JSON.stringify({ left, top }));
    } catch { /* ignore */ }
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
    } catch { /* ignore */ }
  };

  const clampLauncherToViewport = ({ persist = false } = {}) => {
    const rect = launcher.getBoundingClientRect();
    const maxLeft = Math.max(8, window.innerWidth - rect.width - 8);
    const maxTop = Math.max(8, window.innerHeight - rect.height - 8);
    const left = clamp(rect.left, 8, maxLeft);
    const top = clamp(rect.top, 8, maxTop);
    if (Math.abs(left - rect.left) > 0.5 || Math.abs(top - rect.top) > 0.5) {
      applyLauncherPosition({ left, top });
      if (persist) saveLauncherPosition({ left, top });
    }
  };

  let suppressLauncherClick = false;

  launcher.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) return;

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
      if (!moved && (dx > 4 || dy > 4)) moved = true;
      if (!moved) return;
      const maxLeft = Math.max(0, window.innerWidth - rect.width - 8);
      const maxTop = Math.max(0, window.innerHeight - rect.height - 8);
      applyLauncherPosition({
        left: clamp(moveEvent.clientX - offsetX, 8, maxLeft),
        top: clamp(moveEvent.clientY - offsetY, 8, maxTop)
      });
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

  window.addEventListener("resize", () => clampLauncherToViewport({ persist: true }));
  window.addEventListener("orientationchange", () => clampLauncherToViewport({ persist: true }));

  // ── Week helpers ───────────────────────────────────────────────────────

  const getWeekStartDate = (date, weekStartsOn) => {
    const source = date instanceof Date && !Number.isNaN(date.getTime()) ? new Date(date) : new Date();
    source.setHours(0, 0, 0, 0);
    const diffFromWeekStart = (source.getDay() - weekStartsOn + 7) % 7;
    source.setDate(source.getDate() - diffFromWeekStart);
    return source;
  };

  const computeWeekBounds = (anchorDate) => {
    const base = anchorDate instanceof Date && !Number.isNaN(anchorDate.getTime()) ? anchorDate : new Date();
    const startOfWeek = getWeekStartDate(base, FIXED_WEEK_STARTS_ON);
    const endOfWeek = new Date(startOfWeek);
    endOfWeek.setDate(endOfWeek.getDate() + 7);
    return {
      startOfWeek,
      startOfWeekDate: startOfWeek.toISOString().slice(0, 10),
      startOfWeekIso: startOfWeek.toISOString(),
      endOfWeek,
      endOfWeekDate: endOfWeek.toISOString().slice(0, 10),
      endOfWeekIso: endOfWeek.toISOString()
    };
  };

  // ── API utilities ──────────────────────────────────────────────────────

  const sleep = (ms) => new Promise((resolve) => { window.setTimeout(resolve, ms); });

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
    try { data = JSON.parse(text); } catch { data = null; }
    return { response, text, data, url };
  };

  const searchAllTickets = async ({ query, sortBy = "updated_at", sortOrder = "desc", onProgress }) => {
    const params = new URLSearchParams({ query, sort_by: sortBy, sort_order: sortOrder });
    let nextUrl = `${BASE}/api/v2/search.json?${params.toString()}`;
    const tickets = [];
    let page = 0;

    while (nextUrl) {
      page += 1;
      onProgress?.(`Search page ${page}...`);
      const { response, data, url } = await fetchJson(nextUrl);
      if (!response.ok) throw new Error(`Search failed on page ${page} (${response.status}) at ${url}`);
      const pageResults = Array.isArray(data?.results) ? data.results : [];
      tickets.push(...pageResults.filter((r) => r?.result_type === "ticket"));
      nextUrl = data?.next_page || null;
    }

    return tickets;
  };

  const getCurrentUserId = async () => {
    const { response, data, url } = await fetchJson(`${BASE}/api/v2/users/me.json`);
    if (!response.ok) throw new Error(`Unable to fetch current user (${response.status}) at ${url}`);
    const userId = data?.user?.id;
    if (!userId) throw new Error("Current user ID not found in /users/me response.");
    return userId;
  };

  const resolveUserIdFromAssigneeInput = async (assigneeKeyword) => {
    if (!assigneeKeyword || assigneeKeyword === "me") return getCurrentUserId();
    if (/^\d+$/.test(assigneeKeyword)) return Number(assigneeKeyword);

    const params = new URLSearchParams({ query: `type:user ${assigneeKeyword}` });
    const { response, data, url } = await fetchJson(`${BASE}/api/v2/search.json?${params.toString()}`);
    if (!response.ok) throw new Error(`Unable to resolve assignee user (${response.status}) at ${url}`);

    const users = Array.isArray(data?.results)
      ? data.results.filter((r) => r?.result_type === "user")
      : [];
    if (!users.length || !users[0]?.id) throw new Error(`No user found for assignee: ${assigneeKeyword}`);
    return users[0].id;
  };

  const chunkArray = (values = [], size = 100) => {
    const chunks = [];
    for (let i = 0; i < values.length; i += size) chunks.push(values.slice(i, i + size));
    return chunks;
  };

  const uniqueNumericIds = (values = []) => {
    const seen = new Set();
    const result = [];
    for (const value of values) {
      const n = Number(value);
      if (!Number.isFinite(n) || n <= 0 || seen.has(n)) continue;
      seen.add(n);
      result.push(n);
    }
    return result;
  };

  const fetchUsersByIds = async (userIds = [], onProgress) => {
    const ids = uniqueNumericIds(userIds);
    const userMap = new Map();
    if (!ids.length) return userMap;

    const chunks = chunkArray(ids, 100);
    for (let i = 0; i < chunks.length; i++) {
      onProgress?.(`Resolving requester names ${i + 1}/${chunks.length}...`);
      const params = new URLSearchParams({ ids: chunks[i].join(",") });
      const { response, data, url } = await fetchJson(`${BASE}/api/v2/users/show_many.json?${params.toString()}`);
      if (!response.ok) throw new Error(`User lookup failed (${response.status}) at ${url}`);
      for (const user of (Array.isArray(data?.users) ? data.users : [])) {
        if (user?.id) userMap.set(Number(user.id), String(user?.name || user?.email || user.id));
      }
    }
    return userMap;
  };

  const fetchOrganizationsByIds = async (orgIds = [], onProgress) => {
    const ids = uniqueNumericIds(orgIds);
    const orgMap = new Map();
    if (!ids.length) return orgMap;

    const chunks = chunkArray(ids, 100);
    for (let i = 0; i < chunks.length; i++) {
      onProgress?.(`Resolving org names ${i + 1}/${chunks.length}...`);
      const params = new URLSearchParams({ ids: chunks[i].join(",") });
      const { response, data, url } = await fetchJson(`${BASE}/api/v2/organizations/show_many.json?${params.toString()}`);
      if (!response.ok) throw new Error(`Org lookup failed (${response.status}) at ${url}`);
      for (const org of (Array.isArray(data?.organizations) ? data.organizations : [])) {
        if (org?.id) orgMap.set(Number(org.id), String(org?.name || org.id));
      }
    }
    return orgMap;
  };

  const enrichReportsWithNames = async (reports = [], onProgress) => {
    const allRows = reports.flatMap((r) => Array.isArray(r?.ticket_rows) ? r.ticket_rows : []);
    const orgIds = allRows.map((r) => r?.organization_id).filter((v) => v != null);
    const userIds = allRows.map((r) => r?.requester_id).filter((v) => v != null);

    const [orgMap, userMap] = await Promise.all([
      fetchOrganizationsByIds(orgIds, onProgress),
      fetchUsersByIds(userIds, onProgress)
    ]);

    for (const report of reports) {
      for (const row of (Array.isArray(report?.ticket_rows) ? report.ticket_rows : [])) {
        const orgId = Number(row?.organization_id);
        const reqId = Number(row?.requester_id);
        row.organization_name = Number.isFinite(orgId) ? orgMap.get(orgId) || null : null;
        row.requester_name = Number.isFinite(reqId) ? userMap.get(reqId) || null : null;
      }
    }
  };

  const pickTicketIdentity = (ticket) => ({
    ticket_id: ticket?.id ?? null,
    organization_id: ticket?.organization_id ?? null,
    requester_id: ticket?.requester_id ?? null
  });

  // ── Reports ────────────────────────────────────────────────────────────

  const runAssignedCreatedThisWeek = async ({ assigneeKeyword, startOfWeekDate, endOfWeekDate, onProgress }) => {
    const query = `type:ticket assignee:${assigneeKeyword} created>=${startOfWeekDate} created<${endOfWeekDate}`;
    const tickets = await searchAllTickets({ query, onProgress });
    return { name: "Tickets Taken This Week", query, total: tickets.length, ticket_rows: tickets.map(pickTicketIdentity) };
  };

  const runAssignedSolvedUpdatedThisWeek = async ({ assigneeKeyword, startOfWeekDate, endOfWeekDate, onProgress }) => {
    const query = `type:ticket assignee:${assigneeKeyword} status:solved updated>=${startOfWeekDate} updated<${endOfWeekDate}`;
    const tickets = await searchAllTickets({ query, onProgress });
    return { name: "Assigned + Solved + Updated This Week", query, total: tickets.length, ticket_rows: tickets.map(pickTicketIdentity) };
  };

  const runOpenTicketsRemaining = async ({ assigneeKeyword, endOfWeekDate, onProgress }) => {
    const query = `type:ticket assignee:${assigneeKeyword} status:open created<${endOfWeekDate}`;
    const tickets = await searchAllTickets({ query, onProgress });
    return { name: "Open Tickets Remaining", query, total: tickets.length, ticket_rows: tickets.map(pickTicketIdentity) };
  };

  const runCarriedOverTickets = async ({ assigneeKeyword, startOfWeekDate, endOfWeekDate, onProgress }) => {
    const query = `type:ticket assignee:${assigneeKeyword} created<${startOfWeekDate} updated>=${startOfWeekDate} updated<${endOfWeekDate} -status:solved -status:closed`;
    const tickets = await searchAllTickets({ query, onProgress });
    return { name: "Carried Over", query, total: tickets.length, ticket_rows: tickets.map(pickTicketIdentity) };
  };

  const runTakeoverReport = async ({ assigneeKeyword, myUserId, startOfWeekDate, endOfWeekDate, startOfWeekIso, endOfWeekIso, onProgress }) => {
    const query = `type:ticket assignee:${assigneeKeyword} updated>=${startOfWeekDate} updated<${endOfWeekDate}`;
    const candidateTickets = await searchAllTickets({ query, onProgress });
    const takeoverRows = [];

    for (let i = 0; i < candidateTickets.length; i++) {
      const ticket = candidateTickets[i];
      const ticketId = ticket?.id;
      if (!ticketId) continue;

      onProgress?.(`Audits ${i + 1}/${candidateTickets.length} (ticket ${ticketId})...`);
      let nextAuditUrl = `${BASE}/api/v2/tickets/${ticketId}/audits.json`;

      while (nextAuditUrl) {
        const { response, data, url } = await fetchJson(nextAuditUrl);
        if (!response.ok) throw new Error(`Audit fetch failed (${response.status}) at ${url}`);

        for (const audit of (Array.isArray(data?.audits) ? data.audits : [])) {
          const auditCreatedAt = audit?.created_at;
          if (!auditCreatedAt || auditCreatedAt < startOfWeekIso || auditCreatedAt >= endOfWeekIso) continue;

          for (const event of (Array.isArray(audit?.events) ? audit.events : [])) {
            const hasPreviousAssignee = event?.previous_value !== null && event?.previous_value !== undefined && String(event?.previous_value).trim() !== "";
            if (
              event?.type === "Change" &&
              event?.field_name === "assignee_id" &&
              hasPreviousAssignee &&
              String(event?.value) === String(myUserId) &&
              String(event?.previous_value) !== String(myUserId)
            ) {
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
      if (processed % TAKEOVER_BATCH_SIZE === 0 && processed < candidateTickets.length) {
        onProgress?.(`Pausing ${Math.round(TAKEOVER_BATCH_SLEEP_MS / 1000)}s after ${processed} audits...`);
        await sleep(TAKEOVER_BATCH_SLEEP_MS);
      }
    }

    const uniqueTicketIds = [...new Set(takeoverRows.map((r) => r.ticket_id))];
    return {
      name: "Takeovers This Week",
      query,
      candidate_total: candidateTickets.length,
      candidate_audited: candidateTickets.length,
      total_unique_tickets: uniqueTicketIds.length,
      total_takeover_events: takeoverRows.length,
      ticket_rows: takeoverRows
    };
  };

  const dedupeMessageRows = (rows = []) => {
    const seen = new Set();
    const unique = [];
    for (const row of rows) {
      const key = row?.id ? `comment:${row.id}` : `${row?.ticket_id}:${row?.created_at}:${row?.public}:${row?.via_channel}`;
      if (seen.has(key)) continue;
      seen.add(key);
      unique.push(row);
    }
    return unique;
  };

  const runMessagesSentThisWeek = async ({ assigneeKeyword, authorId, startOfWeekDate, endOfWeekDate, startOfWeekIso, endOfWeekIso, onProgress }) => {
    const query = `type:ticket assignee:${assigneeKeyword} updated>=${startOfWeekDate} updated<${endOfWeekDate}`;
    const candidateTickets = await searchAllTickets({ query, onProgress });
    const rows = [];

    for (let i = 0; i < candidateTickets.length; i++) {
      const ticket = candidateTickets[i];
      const ticketId = ticket?.id;
      if (!ticketId) continue;

      onProgress?.(`Message audits ${i + 1}/${candidateTickets.length} (ticket ${ticketId})...`);
      let nextAuditUrl = `${BASE}/api/v2/tickets/${ticketId}/audits.json`;

      while (nextAuditUrl) {
        const { response, data, url } = await fetchJson(nextAuditUrl);
        if (!response.ok) throw new Error(`Message audit fetch failed (${response.status}) at ${url}`);

        for (const audit of (Array.isArray(data?.audits) ? data.audits : [])) {
          const auditCreatedAt = audit?.created_at;
          if (!auditCreatedAt || auditCreatedAt < startOfWeekIso || auditCreatedAt >= endOfWeekIso) continue;
          if (String(audit?.author_id) !== String(authorId)) continue;

          for (const event of (Array.isArray(audit?.events) ? audit.events : [])) {
            if (event?.type !== "Comment") continue;
            rows.push({
              id: event?.id ?? null,
              ticket_id: ticketId,
              organization_id: ticket?.organization_id ?? null,
              requester_id: ticket?.requester_id ?? null,
              author_id: audit?.author_id ?? null,
              via_channel: event?.via?.channel ?? audit?.via?.channel ?? null,
              created_at: event?.created_at || auditCreatedAt,
              public: Boolean(event?.public)
            });
          }
        }

        nextAuditUrl = data?.next_page || null;
      }

      const processed = i + 1;
      if (processed % TAKEOVER_BATCH_SIZE === 0 && processed < candidateTickets.length) {
        onProgress?.(`Pausing ${Math.round(TAKEOVER_BATCH_SLEEP_MS / 1000)}s after ${processed} message audits...`);
        await sleep(TAKEOVER_BATCH_SLEEP_MS);
      }
    }

    const uniqueRows = dedupeMessageRows(rows);
    const publicRows = uniqueRows.filter((r) => r.public);
    const privateRows = uniqueRows.filter((r) => !r.public);

    return {
      name: "Messages Sent This Week",
      query,
      candidate_total: candidateTickets.length,
      total: uniqueRows.length,
      total_public_replies: publicRows.length,
      total_internal_notes: privateRows.length,
      ticket_rows: uniqueRows
    };
  };

  // ── Pulse iframe ───────────────────────────────────────────────────────
  //
  // A small panel iframe is injected into the Zendesk page pointing at
  // PULSE_URL. After reports finish the frame becomes visible and the button
  // changes to "▶ Send to Pulse" — giving you time to switch the DevTools
  // console context to the iframe and paste the receiver script. A second
  // click actually fires the postMessage.

  let pulseFrame      = null;
  let pulseFrameReady = false;
  let pendingPulsePayload = null;

  const FRAME_CSS_HIDDEN  = "position:fixed;bottom:60px;right:16px;width:320px;height:200px;border:2px solid #6366f1;border-radius:8px;z-index:999998;opacity:0;pointer-events:none;transition:opacity 0.3s;";
  const FRAME_CSS_VISIBLE = "position:fixed;bottom:60px;right:16px;width:320px;height:200px;border:2px solid #6366f1;border-radius:8px;z-index:999998;opacity:1;pointer-events:auto;transition:opacity 0.3s;";

  const ensurePulseFrame = () => {
    if (pulseFrame) return;
    pulseFrame = document.createElement("iframe");
    pulseFrame.src = PULSE_URL;
    pulseFrame.setAttribute("aria-hidden", "true");
    pulseFrame.style.cssText = FRAME_CSS_HIDDEN;
    pulseFrame.addEventListener("load", () => {
      pulseFrameReady = true;
      console.info(
        "[zd→pulse] Pulse iframe ready.\n" +
        "Switch the DevTools console context from \"top\" to the pro-serv-pulse iframe,\n" +
        "paste the receiver script there, then click ▶ Send to Pulse."
      );
    });
    document.body.appendChild(pulseFrame);
  };

  // Called after reports finish. Stores payload, shows the frame, and waits
  // for a second button click before actually sending.
  const sendToPulse = async (payload) => {
    ensurePulseFrame();
    if (!pulseFrameReady) {
      setButtonStatus("Waiting for Pulse...", { sticky: true });
      await new Promise((resolve) => {
        pulseFrame.addEventListener("load", resolve, { once: true });
      });
    }
    pendingPulsePayload = payload;
    pulseFrame.style.cssText = FRAME_CSS_VISIBLE;
    setButtonStatus("▶ Send to Pulse", { sticky: true });
  };

  // Second click fires this to actually post the message and hide the frame.
  const flushPulsePayload = () => {
    const payload = pendingPulsePayload;
    pendingPulsePayload = null;
    pulseFrame.contentWindow.postMessage({ type: "zd-report-payload", payload }, PULSE_ORIGIN);
    pulseFrame.style.cssText = FRAME_CSS_HIDDEN;
    setButtonStatus("Sent to Pulse ✓");
  };

  // ── Run ────────────────────────────────────────────────────────────────

  const runReports = async () => {
    launcher.disabled = true;

    const assigneeKeyword = window.zdPulseReport.getAssignee();
    const anchorDate = window.zdPulseReport.getWeekDate();
    const { startOfWeekDate, startOfWeekIso, endOfWeekDate, endOfWeekIso } = computeWeekBounds(anchorDate);

    const progress = (msg) => {
      launcher.textContent = msg.length > 22 ? `${msg.slice(0, 22)}…` : msg;
    };

    try {
      progress("Resolving user...");
      const myUserId = await resolveUserIdFromAssigneeInput(assigneeKeyword);

      progress("Report 1/6...");
      const report1 = await runAssignedCreatedThisWeek({ assigneeKeyword, startOfWeekDate, endOfWeekDate, onProgress: progress });

      progress("Report 2/6...");
      const report2 = await runAssignedSolvedUpdatedThisWeek({ assigneeKeyword, startOfWeekDate, endOfWeekDate, onProgress: progress });

      progress("Report 3/6...");
      const reportOpen = await runOpenTicketsRemaining({ assigneeKeyword, endOfWeekDate, onProgress: progress });

      progress("Report 4/6...");
      const reportCarry = await runCarriedOverTickets({ assigneeKeyword, startOfWeekDate, endOfWeekDate, onProgress: progress });

      progress("Report 5/6 (audits)...");
      const report3 = await runTakeoverReport({
        assigneeKeyword, myUserId, startOfWeekDate, endOfWeekDate, startOfWeekIso, endOfWeekIso, onProgress: progress
      });

      progress("Report 6/6 (messages)...");
      const reportMessages = await runMessagesSentThisWeek({
        assigneeKeyword, authorId: myUserId, startOfWeekDate, endOfWeekDate, startOfWeekIso, endOfWeekIso, onProgress: progress
      });

      try {
        progress("Resolving names...");
        await enrichReportsWithNames([report1, report2, reportOpen, reportCarry, report3, reportMessages], progress);
      } catch (nameError) {
        console.warn("[zd→pulse] Name resolution partially failed; continuing.", nameError);
      }

      const payload = {
        meta: {
          assignee: assigneeKeyword,
          assignee_source: window.zdPulseReport._assigneeOverride ? "override" : "default (me)",
          current_user_id: myUserId,
          week_starts_on: "Sunday",
          start_of_week_date: startOfWeekDate,
          start_of_week_iso: startOfWeekIso,
          end_of_week_date: endOfWeekDate,
          end_of_week_iso: endOfWeekIso,
          anchor_date: anchorDate.toISOString().slice(0, 10),
          generated_at: new Date().toISOString()
        },
        reports: { report1, report2, reportOpen, reportCarry, report3, reportMessages }
      };

      // Store full payload locally (same-origin reference)
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
      } catch { /* ignore storage failures */ }

      // Send totals-only to Pulse — strip ticket_rows to keep the message small
      const stripRows = ({ ticket_rows, ...rest }) => rest;
      const pulsePayload = {
        meta: payload.meta,
        reports: Object.fromEntries(
          Object.entries(payload.reports).map(([k, v]) => [k, stripRows(v)])
        )
      };

      progress("Sending to Pulse...");
      await sendToPulse(pulsePayload);

    } catch (error) {
      setButtonStatus("Error!", { sticky: false });
      console.error("[zd→pulse] Report run failed:", error);
    } finally {
      launcher.disabled = false;
    }
  };

  launcher.addEventListener("click", () => {
    if (suppressLauncherClick) {
      suppressLauncherClick = false;
      return;
    }
    if (pendingPulsePayload) {
      flushPulsePayload();
    } else {
      runReports();
    }
  });

  // ── Console API ────────────────────────────────────────────────────────

  let _assigneeOverride = null;
  let _weekDateOverride = null;

  window.zdPulseReport = {
    _assigneeOverride: null,
    _weekDateOverride: null,

    setAssignee(keyword) {
      const normalized = String(keyword ?? "").trim();
      _assigneeOverride = normalized || null;
      window.zdPulseReport._assigneeOverride = _assigneeOverride;
      console.log(`[zd→pulse] Assignee set to: ${_assigneeOverride ?? ASSIGNEE_FALLBACK_KEYWORD}`);
      return _assigneeOverride ?? ASSIGNEE_FALLBACK_KEYWORD;
    },
    clearAssignee() {
      _assigneeOverride = null;
      window.zdPulseReport._assigneeOverride = null;
      console.log(`[zd→pulse] Assignee cleared, using: ${ASSIGNEE_FALLBACK_KEYWORD}`);
      return ASSIGNEE_FALLBACK_KEYWORD;
    },
    getAssignee() {
      return _assigneeOverride || ASSIGNEE_FALLBACK_KEYWORD;
    },

    setWeekDate(isoDate) {
      const date = new Date(`${isoDate}T00:00:00`);
      if (Number.isNaN(date.getTime())) {
        console.warn(`[zd→pulse] Invalid date: ${isoDate}`);
        return null;
      }
      _weekDateOverride = date;
      window.zdPulseReport._weekDateOverride = isoDate;
      console.log(`[zd→pulse] Week date set to: ${isoDate}`);
      return isoDate;
    },
    clearWeekDate() {
      _weekDateOverride = null;
      window.zdPulseReport._weekDateOverride = null;
      console.log("[zd→pulse] Week date cleared, using current week.");
    },
    getWeekDate() {
      return _weekDateOverride instanceof Date ? _weekDateOverride : new Date();
    },

    run() {
      return runReports();
    }
  };

  // ── Init ───────────────────────────────────────────────────────────────

  loadLauncherPosition();
  clampLauncherToViewport({ persist: true });
  ensurePulseFrame(); // start loading the iframe in the background
})();
