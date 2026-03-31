/* =========================================================
   Hemmelig Outlook Add-in — taskpane.js
   Requires: Office.js loaded in taskpane.html
   ========================================================= */

"use strict";

// ── Config keys ─────────────────────────────────────────────
const KEY_SERVER   = "hemmelig_server";
const KEY_TOKEN    = "hemmelig_token";
const KEY_TEMPLATE = "hemmelig_template";

const DEFAULT_SERVER   = "https://secret.joeymeijer.nl";
const DEFAULT_TEMPLATE = `Hi,\n\nPlease find your secure information via the link below. This link will expire after {expiry} and can only be read once.\n\n🔗 {link}\n\nThis message was sent securely using Hemmelig.`;

// ── State ────────────────────────────────────────────────────
let selectedTtl   = 604800;   // 7 days default
let generatedLink = null;
let emailContent  = { to: "", subject: "", body: "" };

// ── Office initialisation ────────────────────────────────────
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loadSettings();
    loadEmailPreview();
  }
});

// ── Tab switching ────────────────────────────────────────────
function switchTab(name) {
  document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
  document.querySelectorAll(".panel").forEach(p => p.classList.remove("active"));
  document.getElementById("tab-" + name).classList.add("active");
  document.getElementById("panel-" + name).classList.add("active");
}

// ── TTL pill selection ───────────────────────────────────────
function selectTtl(el) {
  document.querySelectorAll(".pill").forEach(p => p.classList.remove("selected"));
  el.classList.add("selected");
  selectedTtl = parseInt(el.dataset.ttl, 10);
}

// ── Password toggle ──────────────────────────────────────────
function togglePassword(checkbox) {
  const field = document.getElementById("pw-field");
  field.classList.toggle("show", checkbox.checked);
  if (!checkbox.checked) {
    document.getElementById("secret-password").value = "";
  }
}

// ── Settings: load / save ────────────────────────────────────
function loadSettings() {
  const server   = localStorage.getItem(KEY_SERVER)   || DEFAULT_SERVER;
  const token    = localStorage.getItem(KEY_TOKEN)    || "";
  const template = localStorage.getItem(KEY_TEMPLATE) || DEFAULT_TEMPLATE;

  document.getElementById("cfg-server").value   = server;
  document.getElementById("cfg-token").value    = token;
  document.getElementById("cfg-template").value = template;
}

function saveSettings() {
  const server   = document.getElementById("cfg-server").value.trim().replace(/\/$/, "");
  const token    = document.getElementById("cfg-token").value.trim();
  const template = document.getElementById("cfg-template").value;

  if (!server) return showStatus("error", "⚠", "Server URL is required.", "Please enter your Hemmelig server URL.");
  if (!token)  return showStatus("error", "⚠", "API token is required.", "Please enter your API token.");

  localStorage.setItem(KEY_SERVER,   server);
  localStorage.setItem(KEY_TOKEN,    token);
  localStorage.setItem(KEY_TEMPLATE, template);

  showStatus("success", "✔", "Settings saved.", "Your configuration has been stored locally.");
  switchTab("compose");
}

// ── Connection test ──────────────────────────────────────────
async function testConnection() {
  const server = document.getElementById("cfg-server").value.trim().replace(/\/$/, "");
  const token  = document.getElementById("cfg-token").value.trim();
  const chip   = document.getElementById("conn-status");

  chip.textContent = "● Testing…";
  chip.style.color = "#e8a020";

  try {
    const res = await fetch(`${server}/api/v1/healthz`, {
      method: "GET",
      headers: { "Authorization": `Bearer ${token}` }
    });

    if (res.ok || res.status === 404) {
      // 404 just means healthz endpoint doesn't exist but server is reachable
      chip.textContent = "● Connected";
      chip.style.color = "#3eb87a";
      chip.style.borderColor = "rgba(62,184,122,0.3)";
      chip.style.background  = "rgba(62,184,122,0.1)";
    } else {
      chip.textContent = `● Error ${res.status}`;
      chip.style.color = "#e05050";
      chip.style.borderColor = "rgba(224,80,80,0.3)";
      chip.style.background  = "rgba(224,80,80,0.1)";
    }
  } catch (e) {
    chip.textContent = "● Unreachable";
    chip.style.color = "#e05050";
    chip.style.borderColor = "rgba(224,80,80,0.3)";
    chip.style.background  = "rgba(224,80,80,0.1)";
  }
}

// ── Load email preview from Outlook ─────────────────────────
function loadEmailPreview() {
  const item = Office.context.mailbox.item;

  // Subject
  item.subject.getAsync((res) => {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      emailContent.subject = res.value || "(No subject)";
      document.getElementById("prev-subject").textContent = emailContent.subject;
    }
  });

  // To recipients
  item.to.getAsync((res) => {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      const addresses = (res.value || []).map(r => r.emailAddress).join(", ");
      emailContent.to = addresses || "(No recipients)";
      document.getElementById("prev-to").textContent = emailContent.to;
    }
  });

  // Body (plain text)
  item.body.getAsync(Office.CoercionType.Text, (res) => {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      emailContent.body = res.value || "";
      const preview = emailContent.body.slice(0, 400) + (emailContent.body.length > 400 ? "…" : "");
      document.getElementById("prev-body").textContent = preview || "(Empty body)";
    }
  });
}

// ── Create Hemmelig secret ───────────────────────────────────
async function createHemmeligSecret(text, ttl, password, preventBurn) {
  const server = (localStorage.getItem(KEY_SERVER) || DEFAULT_SERVER).replace(/\/$/, "");
  const token  = localStorage.getItem(KEY_TOKEN) || "";

  if (!token) throw new Error("No API token configured. Please go to Settings.");

  const body = {
    text,
    ttl,
    preventBurn: !preventBurn,   // hemmelig: preventBurn=false means it WILL burn
    allowedIp: ""
  };

  if (password) body.password = password;

  const res = await fetch(`${server}/api/v1/secret`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${token}`
    },
    body: JSON.stringify(body)
  });

  if (!res.ok) {
    let msg = `API returned ${res.status}`;
    try {
      const err = await res.json();
      if (err.message) msg = err.message;
    } catch (_) {}
    throw new Error(msg);
  }

  const data = await res.json();

  // Hemmelig returns id and optionally a key (encryption key for the fragment)
  if (!data.id) throw new Error("Unexpected API response — no secret ID returned.");

  const link = data.key
    ? `${server}/secret/${data.id}#${data.key}`
    : `${server}/secret/${data.id}`;

  return link;
}

// ── Build secret payload ─────────────────────────────────────
function buildSecretPayload() {
  const parts = [];

  if (emailContent.subject) {
    parts.push(`Subject: ${emailContent.subject}`);
  }
  if (emailContent.to) {
    parts.push(`To: ${emailContent.to}`);
  }
  if (emailContent.body.trim()) {
    parts.push(`\n--- Message ---\n${emailContent.body.trim()}`);
  }

  return parts.join("\n");
}

// ── TTL label ────────────────────────────────────────────────
function ttlLabel(seconds) {
  if (seconds <= 86400)   return "1 day";
  if (seconds <= 604800)  return "7 days";
  if (seconds <= 2592000) return "30 days";
  return `${Math.round(seconds / 86400)} days`;
}

// ── Main action: Send Securely ───────────────────────────────
async function handleSendSecurely() {
  const btn = document.getElementById("send-btn");

  // Refresh email content first
  await refreshEmailContent();

  const payload  = buildSecretPayload();
  const password = document.getElementById("secret-password").value.trim();
  const burnAfter = document.getElementById("burn-toggle").checked;

  if (!payload.trim()) {
    return showStatus("error", "⚠", "Empty email.", "There is no content to send securely.");
  }

  // Loading state
  btn.disabled = true;
  btn.classList.add("loading");
  hideStatus();
  hideLink();

  try {
    const link = await createHemmeligSecret(payload, selectedTtl, password, burnAfter);
    generatedLink = link;

    showStatus("success", "✔", "Secret created!", `Your secret link is ready. It expires in ${ttlLabel(selectedTtl)}.`);
    showLink(link);

  } catch (err) {
    showStatus("error", "✗", "Failed to create secret.", err.message || "Unknown error. Check your settings and try again.");
  } finally {
    btn.disabled = false;
    btn.classList.remove("loading");
  }
}

// ── Inject secret link into email body ───────────────────────
function injectIntoEmail() {
  if (!generatedLink) return;

  const template = localStorage.getItem(KEY_TEMPLATE) || DEFAULT_TEMPLATE;
  const expiry   = ttlLabel(selectedTtl);

  const newBody = template
    .replace(/\{link\}/g,   generatedLink)
    .replace(/\{expiry\}/g, expiry);

  Office.context.mailbox.item.body.setAsync(
    newBody,
    { coercionType: Office.CoercionType.Text },
    (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        showStatus("success", "✔", "Email body replaced.", "The original content has been replaced with the secure link.");
      } else {
        showStatus("error", "✗", "Could not update email.", res.error?.message || "Unknown error.");
      }
    }
  );
}

// ── Copy link to clipboard ───────────────────────────────────
async function copyLink() {
  if (!generatedLink) return;

  try {
    await navigator.clipboard.writeText(generatedLink);
    const btn = document.querySelector(".link-actions .btn-sm.gold");
    const orig = btn.textContent;
    btn.textContent = "✔ Copied!";
    setTimeout(() => { btn.textContent = orig; }, 2000);
  } catch (_) {
    // Fallback for older Office hosts
    const ta = document.createElement("textarea");
    ta.value = generatedLink;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
  }
}

// ── Refresh email data ───────────────────────────────────────
function refreshEmailContent() {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;
    let done = 0;
    const check = () => { if (++done === 3) resolve(); };

    item.subject.getAsync((r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        emailContent.subject = r.value || "";
        document.getElementById("prev-subject").textContent = emailContent.subject || "—";
      }
      check();
    });

    item.to.getAsync((r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        emailContent.to = (r.value || []).map(x => x.emailAddress).join(", ");
        document.getElementById("prev-to").textContent = emailContent.to || "—";
      }
      check();
    });

    item.body.getAsync(Office.CoercionType.Text, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        emailContent.body = r.value || "";
        const preview = emailContent.body.slice(0, 400) + (emailContent.body.length > 400 ? "…" : "");
        document.getElementById("prev-body").textContent = preview || "(Empty body)";
      }
      check();
    });
  });
}

// ── UI helpers ───────────────────────────────────────────────
function showStatus(type, icon, title, detail) {
  const banner = document.getElementById("status-banner");
  document.getElementById("status-icon").textContent = icon;
  document.getElementById("status-text").innerHTML =
    `<strong>${escHtml(title)}</strong>${escHtml(detail)}`;
  banner.className = `status-banner show ${type}`;
}

function hideStatus() {
  document.getElementById("status-banner").className = "status-banner";
}

function showLink(link) {
  document.getElementById("link-value").textContent = link;
  document.getElementById("link-box").classList.add("show");
}

function hideLink() {
  document.getElementById("link-box").classList.remove("show");
  generatedLink = null;
}

function escHtml(str) {
  return (str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
