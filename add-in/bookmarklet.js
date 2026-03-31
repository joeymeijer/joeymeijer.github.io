/* ============================================================
   Hemmelig OWA Bookmarklet — bookmarklet.js
   Hosted at: https://joeymeijer.nl/add-in/bookmarklet.js
   ============================================================ */
(function () {
  "use strict";

  const PANEL_ID  = "hemmelig-panel";
  const KEY_SERVER   = "hemmelig_server";
  const KEY_TOKEN    = "hemmelig_token";
  const KEY_TEMPLATE = "hemmelig_template";
  const DEFAULT_SERVER   = "https://secret.joeymeijer.nl";
  const DEFAULT_TEMPLATE = "Hi,\n\nPlease find your secure information via the link below.\nThis link will expire after {expiry} and can only be read once.\n\n🔗 {link}\n\nThis message was sent securely using Hemmelig.";

  // ── Toggle: remove if already open ─────────────────────────
  const existing = document.getElementById(PANEL_ID);
  if (existing) { existing.remove(); return; }

  // ── Inject styles ───────────────────────────────────────────
  const style = document.createElement("style");
  style.textContent = `
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=Syne:wght@700;800&family=Inter:wght@300;400;500&display=swap');

    #hemmelig-panel {
      all: initial;
      position: fixed;
      top: 72px;
      right: 24px;
      width: 320px;
      max-height: calc(100vh - 96px);
      background: #0c0e14;
      border: 1px solid #e8a020;
      border-radius: 12px;
      box-shadow: 0 8px 48px rgba(0,0,0,.7), 0 0 0 1px rgba(232,160,32,.15);
      z-index: 2147483647;
      font-family: 'Inter', sans-serif;
      font-size: 13px;
      color: #d4d8e8;
      display: flex;
      flex-direction: column;
      overflow: hidden;
      animation: hgSlideIn .22s cubic-bezier(.22,1,.36,1);
    }

    @keyframes hgSlideIn {
      from { opacity:0; transform: translateX(24px) scale(.97); }
      to   { opacity:1; transform: translateX(0) scale(1); }
    }

    #hemmelig-panel * { box-sizing: border-box; margin: 0; padding: 0; font-family: inherit; }

    .hg-header {
      padding: 12px 14px 10px;
      border-bottom: 1px solid #1e2230;
      display: flex;
      align-items: center;
      gap: 9px;
      flex-shrink: 0;
    }

    .hg-logo {
      width: 26px; height: 26px;
      background: #e8a020;
      border-radius: 6px;
      display: flex; align-items: center; justify-content: center;
      font-size: 13px; flex-shrink: 0;
    }

    .hg-title {
      font-family: 'Syne', sans-serif;
      font-size: 13px; font-weight: 700;
      color: #fff; letter-spacing: .02em;
    }

    .hg-sub {
      font-family: 'IBM Plex Mono', monospace;
      font-size: 9px; color: #5a6080;
    }

    .hg-close {
      margin-left: auto;
      background: none; border: none;
      color: #5a6080; cursor: pointer;
      font-size: 18px; line-height: 1;
      padding: 2px 5px; border-radius: 4px;
      transition: color .15s, background .15s;
    }
    .hg-close:hover { color: #d4d8e8; background: #1e2230; }

    .hg-tabs {
      display: flex; border-bottom: 1px solid #1e2230; flex-shrink: 0;
    }

    .hg-tab {
      flex: 1; padding: 8px 0; text-align: center;
      font-family: 'Syne', sans-serif; font-size: 10px;
      font-weight: 700; letter-spacing: .07em; text-transform: uppercase;
      color: #5a6080; background: none; border: none;
      border-bottom: 2px solid transparent; cursor: pointer;
      transition: color .15s, border-color .15s;
    }
    .hg-tab.hg-active { color: #e8a020; border-bottom-color: #e8a020; }

    .hg-body {
      overflow-y: auto; padding: 14px;
      flex: 1;
    }
    .hg-body::-webkit-scrollbar { width: 3px; }
    .hg-body::-webkit-scrollbar-thumb { background: #252839; border-radius: 2px; }

    .hg-panel { display: none; }
    .hg-panel.hg-active { display: block; }

    .hg-label {
      display: block; font-family: 'IBM Plex Mono', monospace;
      font-size: 9px; color: #5a6080;
      letter-spacing: .08em; text-transform: uppercase;
      margin-bottom: 5px;
    }

    .hg-group { margin-bottom: 13px; }

    .hg-preview {
      background: #13151f; border: 1px solid #1e2230;
      border-radius: 7px; overflow: hidden;
    }

    .hg-preview-head {
      padding: 7px 9px; border-bottom: 1px solid #1e2230;
    }

    .hg-meta {
      font-family: 'IBM Plex Mono', monospace;
      font-size: 9px; color: #5a6080; margin-bottom: 2px;
    }
    .hg-meta span { color: #d4d8e8; }

    .hg-preview-body {
      padding: 8px 9px;
      font-size: 11px; color: #5a6080;
      max-height: 80px; overflow-y: auto;
      white-space: pre-wrap; word-break: break-word; line-height: 1.4;
    }

    .hg-pills { display: flex; gap: 5px; }

    .hg-pill {
      flex: 1; text-align: center;
      padding: 7px 4px; border-radius: 6px;
      border: 1px solid #1e2230; background: #13151f;
      color: #5a6080;
      font-family: 'IBM Plex Mono', monospace; font-size: 10px;
      cursor: pointer; transition: all .12s;
    }
    .hg-pill.hg-sel { border-color: #e8a020; color: #e8a020; background: rgba(232,160,32,.09); }

    .hg-toggle-row {
      display: flex; align-items: center; justify-content: space-between;
      margin-bottom: 6px;
    }

    .hg-toggle-label { font-size: 11px; color: #d4d8e8; }

    .hg-switch { position: relative; width: 30px; height: 17px; }
    .hg-switch input { display: none; }

    .hg-track {
      position: absolute; inset: 0;
      background: #1e2230; border-radius: 9px; cursor: pointer;
      transition: background .2s;
    }
    .hg-switch input:checked + .hg-track { background: #9a6a14; }

    .hg-thumb {
      position: absolute; top: 3px; left: 3px;
      width: 11px; height: 11px;
      background: #fff; border-radius: 50%;
      transition: transform .2s; pointer-events: none;
    }
    .hg-switch input:checked ~ .hg-thumb { transform: translateX(13px); }

    .hg-opt { display: none; margin-top: 6px; }
    .hg-opt.hg-show { display: block; }

    input.hg-input, textarea.hg-input {
      width: 100%; background: #13151f;
      border: 1px solid #1e2230; border-radius: 7px;
      color: #d4d8e8; font-size: 12px;
      padding: 7px 9px; outline: none;
      transition: border-color .15s;
    }
    input.hg-input:focus, textarea.hg-input:focus { border-color: #9a6a14; }
    textarea.hg-input { resize: none; height: 72px; }

    .hg-note { font-size: 10px; color: #5a6080; margin-top: 4px; line-height: 1.4; }

    .hg-status {
      margin: 0 14px 10px;
      padding: 8px 10px; border-radius: 7px;
      font-size: 11px; display: none;
      align-items: flex-start; gap: 7px; flex-shrink: 0;
    }
    .hg-status.hg-show { display: flex; }
    .hg-status.hg-err  { background: rgba(224,80,80,.12); border: 1px solid rgba(224,80,80,.3); color: #f08080; }
    .hg-status.hg-ok   { background: rgba(62,184,122,.12); border: 1px solid rgba(62,184,122,.3); color: #6de4a8; }

    .hg-status-icon { font-size: 13px; flex-shrink: 0; }
    .hg-status-body strong { display: block; margin-bottom: 1px; }

    .hg-link-box {
      margin: 0 14px 10px;
      background: #13151f; border: 1px solid #252839;
      border-radius: 7px; overflow: hidden; display: none; flex-shrink: 0;
    }
    .hg-link-box.hg-show { display: block; }

    .hg-link-label {
      padding: 5px 9px;
      background: rgba(232,160,32,.09);
      border-bottom: 1px solid #1e2230;
      font-family: 'IBM Plex Mono', monospace;
      font-size: 9px; letter-spacing: .1em;
      color: #e8a020; text-transform: uppercase;
    }

    .hg-link-val {
      padding: 7px 9px;
      font-family: 'IBM Plex Mono', monospace;
      font-size: 9px; color: #d4d8e8;
      word-break: break-all; line-height: 1.5;
    }

    .hg-link-acts {
      padding: 6px 9px; border-top: 1px solid #1e2230;
      display: flex; gap: 5px;
    }

    .hg-btn-sm {
      padding: 5px 9px; border-radius: 5px;
      font-size: 10px; font-family: 'IBM Plex Mono', monospace;
      cursor: pointer; border: 1px solid #1e2230;
      background: #0c0e14; color: #5a6080;
      transition: all .12s;
    }
    .hg-btn-sm:hover { border-color: #252839; color: #d4d8e8; }
    .hg-btn-sm.hg-gold { border-color: #9a6a14; color: #e8a020; background: rgba(232,160,32,.07); }
    .hg-btn-sm.hg-gold:hover { background: rgba(232,160,32,.14); }

    .hg-footer {
      padding: 11px 14px;
      border-top: 1px solid #1e2230; flex-shrink: 0;
    }

    .hg-btn-main {
      width: 100%; padding: 10px;
      background: #e8a020; color: #0c0e14;
      border: none; border-radius: 7px;
      font-family: 'Syne', sans-serif;
      font-size: 12px; font-weight: 700;
      letter-spacing: .04em; cursor: pointer;
      display: flex; align-items: center;
      justify-content: center; gap: 6px;
      transition: background .15s, opacity .15s;
    }
    .hg-btn-main:hover { background: #f0b030; }
    .hg-btn-main:disabled { opacity: .45; cursor: not-allowed; }

    .hg-spinner {
      width: 12px; height: 12px;
      border: 2px solid rgba(0,0,0,.25);
      border-top-color: #0c0e14;
      border-radius: 50%;
      animation: hgSpin .6s linear infinite;
      display: none;
    }
    .hg-btn-main.hg-loading .hg-spinner { display: block; }
    .hg-btn-main.hg-loading .hg-lbl     { display: none; }

    @keyframes hgSpin { to { transform: rotate(360deg); } }

    .hg-divider { border: none; border-top: 1px solid #1e2230; margin: 12px 0; }

    .hg-conn-chip {
      display: inline-flex; align-items: center; gap: 3px;
      padding: 2px 7px; border-radius: 20px;
      font-size: 9px; font-family: 'IBM Plex Mono', monospace;
      background: rgba(62,184,122,.1);
      border: 1px solid rgba(62,184,122,.25);
      color: #3eb87a; letter-spacing: .05em;
      text-transform: uppercase;
    }
  `;
  document.head.appendChild(style);

  // ── Build panel HTML ────────────────────────────────────────
  const panel = document.createElement("div");
  panel.id = PANEL_ID;
  panel.innerHTML = `
    <div class="hg-header">
      <div class="hg-logo">🔒</div>
      <div>
        <div class="hg-title">Send Securely</div>
        <div class="hg-sub">secret.joeymeijer.nl</div>
      </div>
      <button class="hg-close" id="hg-close">×</button>
    </div>

    <div class="hg-tabs">
      <button class="hg-tab hg-active" data-tab="compose">Compose</button>
      <button class="hg-tab" data-tab="settings">Settings</button>
    </div>

    <!-- COMPOSE -->
    <div class="hg-body" id="hg-body-compose" style="display:flex;flex-direction:column;gap:0;">
      <div class="hg-panel hg-active" id="hg-tab-compose">

        <div class="hg-group">
          <label class="hg-label">Email Preview</label>
          <div class="hg-preview">
            <div class="hg-preview-head">
              <div class="hg-meta">To: <span id="hg-prev-to">—</span></div>
              <div class="hg-meta">Subject: <span id="hg-prev-subj">—</span></div>
            </div>
            <div class="hg-preview-body" id="hg-prev-body">Reading email…</div>
          </div>
        </div>

        <div class="hg-group">
          <label class="hg-label">Expires after</label>
          <div class="hg-pills">
            <div class="hg-pill" data-ttl="86400">1 Day</div>
            <div class="hg-pill hg-sel" data-ttl="604800">7 Days</div>
            <div class="hg-pill" data-ttl="2592000">30 Days</div>
          </div>
        </div>

        <div class="hg-group">
          <div class="hg-toggle-row">
            <span class="hg-toggle-label">Password protect</span>
            <label class="hg-switch">
              <input type="checkbox" id="hg-pw-chk">
              <div class="hg-track"></div>
              <div class="hg-thumb"></div>
            </label>
          </div>
          <div class="hg-opt" id="hg-pw-opt">
            <input type="password" class="hg-input" id="hg-pw" placeholder="Passphrase…">
          </div>
        </div>

        <div class="hg-group">
          <div class="hg-toggle-row">
            <span class="hg-toggle-label">Burn after reading</span>
            <label class="hg-switch">
              <input type="checkbox" id="hg-burn" checked>
              <div class="hg-track"></div>
              <div class="hg-thumb"></div>
            </label>
          </div>
          <div class="hg-note">Secret is destroyed after first view.</div>
        </div>

      </div>

      <!-- SETTINGS -->
      <div class="hg-panel" id="hg-tab-settings">
        <div class="hg-group">
          <label class="hg-label">Server URL</label>
          <input type="text" class="hg-input" id="hg-server" value="${localStorage.getItem(KEY_SERVER) || DEFAULT_SERVER}">
        </div>
        <div class="hg-group">
          <label class="hg-label">API Token</label>
          <input type="password" class="hg-input" id="hg-token" value="${localStorage.getItem(KEY_TOKEN) || ''}" placeholder="Paste API token…">
        </div>
        <div class="hg-group">
          <label class="hg-label">Email Template</label>
          <textarea class="hg-input" id="hg-tmpl">${localStorage.getItem(KEY_TEMPLATE) || DEFAULT_TEMPLATE}</textarea>
          <div class="hg-note">Use {link} and {expiry} as placeholders.</div>
        </div>
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">
          <label class="hg-label" style="margin:0;">Status</label>
          <div class="hg-conn-chip" id="hg-chip">● Not tested</div>
        </div>
        <button class="hg-btn-sm hg-gold" style="width:100%;text-align:center;padding:7px;" id="hg-test">Test connection</button>
      </div>
    </div>

    <!-- Status / link (outside scroll area) -->
    <div class="hg-status" id="hg-status">
      <div class="hg-status-icon" id="hg-st-icon"></div>
      <div class="hg-status-body" id="hg-st-body"></div>
    </div>

    <div class="hg-link-box" id="hg-link-box">
      <div class="hg-link-label">🔗 Secret Link</div>
      <div class="hg-link-val" id="hg-link-val"></div>
      <div class="hg-link-acts">
        <button class="hg-btn-sm hg-gold" id="hg-copy">Copy Link</button>
        <button class="hg-btn-sm" id="hg-inject">Replace Email Body</button>
      </div>
    </div>

    <div class="hg-footer">
      <button class="hg-btn-main" id="hg-send">
        <div class="hg-spinner"></div>
        <span class="hg-lbl">🔒 &nbsp;Send Securely</span>
      </button>
      <div id="hg-save-wrap" style="display:none;margin-top:8px;">
        <button class="hg-btn-main" id="hg-save" style="background:#3eb87a;">
          <span class="hg-lbl">Save Settings</span>
        </button>
      </div>
    </div>
  `;
  document.body.appendChild(panel);

  // ── State ───────────────────────────────────────────────────
  let selectedTtl   = 604800;
  let generatedLink = null;
  let emailData     = { to: "", subject: "", body: "" };
  let activeTab     = "compose";

  // ── Helpers ─────────────────────────────────────────────────
  function esc(s) {
    return (s || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
  }

  function showStatus(type, icon, title, detail) {
    const el = document.getElementById("hg-status");
    document.getElementById("hg-st-icon").textContent = icon;
    document.getElementById("hg-st-body").innerHTML = `<strong>${esc(title)}</strong>${esc(detail)}`;
    el.className = `hg-status hg-show hg-${type}`;
  }

  function hideStatus() {
    document.getElementById("hg-status").className = "hg-status";
  }

  function showLink(url) {
    document.getElementById("hg-link-val").textContent = url;
    document.getElementById("hg-link-box").classList.add("hg-show");
  }

  function hideLink() {
    document.getElementById("hg-link-box").classList.remove("hg-show");
    generatedLink = null;
  }

  function ttlLabel(s) {
    if (s <= 86400)   return "1 day";
    if (s <= 604800)  return "7 days";
    return "30 days";
  }

  // ── OWA DOM readers ─────────────────────────────────────────
  function readOwaEmail() {
    // Subject
    const subjEl =
      document.querySelector('[aria-label="Add a subject"]') ||
      document.querySelector('[placeholder*="subject" i]') ||
      document.querySelector('input[aria-label*="Subject" i]');
    emailData.subject = subjEl ? subjEl.value.trim() : "(no subject)";

    // To
    const toTokens = document.querySelectorAll(
      '[id*="ToLine"] [data-shared-element-id], ' +
      '[aria-label*="To" i] .ms-Persona-primaryText, ' +
      '.customScrollBar [class*="wellItem"] [class*="rootPersona"]'
    );
    if (toTokens.length) {
      emailData.to = Array.from(toTokens).map(e => e.textContent.trim()).filter(Boolean).join(", ");
    } else {
      // Fallback: look for email-like text in the To line
      const toLine = document.querySelector('[id*="ToLine"]') ||
                     document.querySelector('[aria-label*="To recipients" i]');
      emailData.to = toLine ? toLine.textContent.replace(/^To\s*/i,"").trim() : "(unknown)";
    }

    // Body — OWA's contenteditable compose area
    const bodyEl =
      document.querySelector('[aria-label="Message body, press Alt+F10 to exit"]') ||
      document.querySelector('div[contenteditable="true"][role="textbox"]') ||
      document.querySelector('[aria-multiline="true"][contenteditable="true"]');
    emailData.body = bodyEl ? (bodyEl.innerText || bodyEl.textContent || "").trim() : "";

    // Update preview
    document.getElementById("hg-prev-to").textContent   = emailData.to   || "—";
    document.getElementById("hg-prev-subj").textContent = emailData.subject || "—";
    const bodyPreview = emailData.body.slice(0, 300) + (emailData.body.length > 300 ? "…" : "");
    document.getElementById("hg-prev-body").textContent = bodyPreview || "(empty)";
  }

  // ── API call ─────────────────────────────────────────────────
  async function createSecret(text, ttl, password) {
    const server = (localStorage.getItem(KEY_SERVER) || DEFAULT_SERVER).replace(/\/$/, "");
    const token  = localStorage.getItem(KEY_TOKEN) || "";
    if (!token) throw new Error("No API token. Go to Settings first.");

    const payload = { text, ttl, preventBurn: !document.getElementById("hg-burn").checked, allowedIp: "" };
    if (password) payload.password = password;

    const res = await fetch(`${server}/api/v1/secret`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
      body: JSON.stringify(payload)
    });

    if (!res.ok) {
      let msg = `API ${res.status}`;
      try { const e = await res.json(); if (e.message) msg = e.message; } catch(_) {}
      throw new Error(msg);
    }

    const data = await res.json();
    if (!data.id) throw new Error("No secret ID returned.");
    return data.key
      ? `${server}/secret/${data.id}#${data.key}`
      : `${server}/secret/${data.id}`;
  }

  // ── Inject link into OWA compose body ────────────────────────
  function injectIntoOwa(text) {
    const bodyEl =
      document.querySelector('[aria-label="Message body, press Alt+F10 to exit"]') ||
      document.querySelector('div[contenteditable="true"][role="textbox"]') ||
      document.querySelector('[aria-multiline="true"][contenteditable="true"]');

    if (!bodyEl) {
      showStatus("err","✗","Could not find email body.","Please click inside the compose window and try again.");
      return;
    }

    bodyEl.focus();
    // Select all and replace via execCommand (works in OWA's sandboxed editor)
    document.execCommand("selectAll", false, null);
    document.execCommand("insertText", false, text);
    showStatus("ok","✔","Email body replaced.","Original content replaced with the secure link.");
  }

  // ── Tab switching ────────────────────────────────────────────
  function switchTab(name) {
    activeTab = name;
    document.querySelectorAll("#hemmelig-panel .hg-tab").forEach(t => {
      t.classList.toggle("hg-active", t.dataset.tab === name);
    });
    document.querySelectorAll("#hemmelig-panel .hg-panel").forEach(p => {
      p.classList.toggle("hg-active", p.id === `hg-tab-${name}`);
    });
    document.getElementById("hg-send").closest(".hg-footer").querySelector("#hg-save-wrap").style.display =
      name === "settings" ? "block" : "none";
    document.getElementById("hg-send").style.display =
      name === "compose" ? "flex" : "none";
    hideStatus();
    hideLink();
  }

  // ── Wire events ──────────────────────────────────────────────
  document.getElementById("hg-close").addEventListener("click", () => panel.remove());

  document.querySelectorAll("#hemmelig-panel .hg-tab").forEach(t => {
    t.addEventListener("click", () => switchTab(t.dataset.tab));
  });

  document.querySelectorAll("#hemmelig-panel .hg-pill").forEach(p => {
    p.addEventListener("click", () => {
      document.querySelectorAll("#hemmelig-panel .hg-pill").forEach(x => x.classList.remove("hg-sel"));
      p.classList.add("hg-sel");
      selectedTtl = parseInt(p.dataset.ttl, 10);
    });
  });

  document.getElementById("hg-pw-chk").addEventListener("change", function() {
    document.getElementById("hg-pw-opt").classList.toggle("hg-show", this.checked);
    if (!this.checked) document.getElementById("hg-pw").value = "";
  });

  // Test connection
  document.getElementById("hg-test").addEventListener("click", async () => {
    const server = document.getElementById("hg-server").value.trim().replace(/\/$/,"");
    const token  = document.getElementById("hg-token").value.trim();
    const chip   = document.getElementById("hg-chip");
    chip.textContent = "● Testing…"; chip.style.color = "#e8a020";
    try {
      const r = await fetch(`${server}/api/v1/healthz`, { headers: { Authorization: `Bearer ${token}` } });
      if (r.ok || r.status === 404) {
        chip.textContent = "● Connected"; chip.style.color = "#3eb87a";
      } else {
        chip.textContent = `● Error ${r.status}`; chip.style.color = "#e05050";
      }
    } catch(_) {
      chip.textContent = "● Unreachable"; chip.style.color = "#e05050";
    }
  });

  // Save settings
  document.getElementById("hg-save").addEventListener("click", () => {
    const server = document.getElementById("hg-server").value.trim().replace(/\/$/,"");
    const token  = document.getElementById("hg-token").value.trim();
    const tmpl   = document.getElementById("hg-tmpl").value;
    if (!server || !token) return showStatus("err","⚠","Fill in server URL and token.","");
    localStorage.setItem(KEY_SERVER, server);
    localStorage.setItem(KEY_TOKEN, token);
    localStorage.setItem(KEY_TEMPLATE, tmpl);
    showStatus("ok","✔","Settings saved.","");
    setTimeout(() => switchTab("compose"), 800);
  });

  // Main send action
  document.getElementById("hg-send").addEventListener("click", async () => {
    readOwaEmail();
    const payload  = `Subject: ${emailData.subject}\nTo: ${emailData.to}\n\n--- Message ---\n${emailData.body}`;
    const password = document.getElementById("hg-pw").value.trim();

    if (!payload.trim()) return showStatus("err","⚠","Empty email.","Nothing to send securely.");

    const btn = document.getElementById("hg-send");
    btn.disabled = true; btn.classList.add("hg-loading");
    hideStatus(); hideLink();

    try {
      const link = await createSecret(payload, selectedTtl, password);
      generatedLink = link;
      showStatus("ok","✔","Secret created!",`Expires in ${ttlLabel(selectedTtl)}.`);
      showLink(link);
    } catch(e) {
      showStatus("err","✗","Failed.",e.message || "Unknown error.");
    } finally {
      btn.disabled = false; btn.classList.remove("hg-loading");
    }
  });

  // Copy link
  document.getElementById("hg-copy").addEventListener("click", async () => {
    if (!generatedLink) return;
    try {
      await navigator.clipboard.writeText(generatedLink);
    } catch(_) {
      const ta = document.createElement("textarea");
      ta.value = generatedLink; document.body.appendChild(ta); ta.select();
      document.execCommand("copy"); ta.remove();
    }
    const btn = document.getElementById("hg-copy");
    const orig = btn.textContent; btn.textContent = "✔ Copied!";
    setTimeout(() => btn.textContent = orig, 2000);
  });

  // Replace email body
  document.getElementById("hg-inject").addEventListener("click", () => {
    if (!generatedLink) return;
    const tmpl = localStorage.getItem(KEY_TEMPLATE) || DEFAULT_TEMPLATE;
    const text = tmpl.replace(/\{link\}/g, generatedLink).replace(/\{expiry\}/g, ttlLabel(selectedTtl));
    injectIntoOwa(text);
  });

  // ── Init ─────────────────────────────────────────────────────
  switchTab("compose");
  setTimeout(readOwaEmail, 300); // let the DOM settle

})();
