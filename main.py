"""
main.py
Identifyer for Zotero — local web app (no tkinter dependency).
Opens in your default browser at http://localhost:PORT
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import List, Optional
from urllib.parse import urlparse

# ---------------------------------------------------------------------------
# Lazy imports — GUI still starts if packages missing
# ---------------------------------------------------------------------------
IMPORT_ERROR: Optional[str] = None
try:
    from citation_parser import CitationParser
    from field_writer import write_zotero_document
    from matcher import CitationMatcher, MatchResult
    from zotero_client import ZoteroClient
    _DEFAULT_DB = str(ZoteroClient._find_default_db())
except ImportError as e:
    IMPORT_ERROR = str(e)
    _DEFAULT_DB = str(Path.home() / "Zotero" / "zotero.sqlite")

from config import load_settings, update_setting

# Prefer the user's most recently used DB path (if still valid) over the
# platform-autodetected one, so users who keep Zotero in a non-standard
# location don't have to re-pick it every time.
_saved_db = (load_settings().get("db_path") or "").strip()
if _saved_db and Path(_saved_db).exists():
    _DEFAULT_DB = _saved_db


# ---------------------------------------------------------------------------
# Shared analysis state (single-user local app)
# ---------------------------------------------------------------------------
_state: dict = {
    "status": "idle",        # idle | running | done | error
    "progress": 0,
    "message": "Ready.",
    "results": None,         # List[dict] after analysis
    "docx_path": "",
    "db_path": _DEFAULT_DB,
    "error": None,
    "accepted_suggestions": [],  # list of result indices accepted by user
    "uncited_refs": [],          # reference list entries with no in-text citation
}
_state_lock = threading.Lock()


def _set_state(**kwargs):
    with _state_lock:
        _state.update(kwargs)


def _get_state() -> dict:
    with _state_lock:
        return dict(_state)


# ---------------------------------------------------------------------------
# Analysis worker
# ---------------------------------------------------------------------------

def _run_analysis(
    docx_path: str,
    db_path: str,
    author_threshold: int = 82,
    candidate_threshold: int = 60,
    year_tolerance: int = 1,
    collection_id: Optional[int] = None,
):
    try:
        _set_state(status="running", progress=5, message="Loading Zotero library…")

        client = ZoteroClient(db_path)

        def lib_progress(done, total):
            if total:
                _set_state(progress=5 + int(40 * done / total))

        library = client.load_library(
            progress_cb=lib_progress, collection_id=collection_id
        )
        _set_state(
            progress=45,
            message=f"Loaded {len(library):,} items. Parsing document…",
        )

        parser = CitationParser(docx_path)
        citations = parser.parse()
        uncited_refs = parser.get_uncited_references(citations)
        _set_state(
            progress=55,
            message=f"Found {len(citations)} citations. Matching…",
        )

        matcher = CitationMatcher(
            library,
            author_threshold=author_threshold,
            candidate_threshold=candidate_threshold,
            year_tolerance=year_tolerance,
            bibliography=parser.author_year_ref_map,
        )

        def match_progress(done, total):
            if total:
                _set_state(progress=55 + int(40 * done / total))

        results = matcher.match_all(citations, progress_cb=match_progress)

        # Serialise to JSON-safe dicts
        rows = []
        for i, r in enumerate(results):
            cit = r.citation
            cands = [
                {
                    "key": it.key,
                    "title": it.title[:80] if it.title else "",
                    "year": it.year,
                    "authors": it.authors[:3],
                    "confidence": round(c * 100),
                }
                for it, c in (r.candidates[:5] if r.candidates else [])
            ]
            rows.append({
                "index": i,
                "display": cit.display(),
                "raw": cit.raw_text,
                "style": cit.style,
                "authors": cit.authors,
                "year": cit.year,
                "para": cit.paragraph_idx,
                "in_footnote": cit.in_footnote,
                "matched": r.matched,
                "is_suggestion": r.is_suggestion,
                "suggestion_year": r.suggestion_year,
                "is_ambiguous": r.is_ambiguous,
                "candidates": cands,
                "confidence": round(r.confidence * 100),
                "zotero_key": r.zotero_item.key if r.zotero_item else "",
                "zotero_title": r.zotero_item.title[:80] if r.zotero_item else "",
                "zotero_year": r.zotero_item.year if r.zotero_item else "",
                "zotero_authors": r.zotero_item.authors[:3] if r.zotero_item else [],
                "ref_text": cit.ref_text,
                "prefix": cit.prefix,
                "suffix": cit.suffix,
                "is_last_resort": r.is_last_resort,
                "bibliography_hint": r.bibliography_hint,
            })

        matched   = sum(1 for r in results if r.matched)
        ambiguous = sum(1 for r in results if r.is_ambiguous)
        suggestions = sum(1 for r in results if r.is_suggestion)
        unmatched = len(results) - matched - ambiguous - suggestions
        _set_state(
            status="done",
            progress=100,
            message=(
                f"Done — {matched} matched, {ambiguous} ambiguous, "
                f"{suggestions} suggestions, {unmatched} unmatched "
                f"out of {len(results)} citations."
            ),
            results=rows,
            error=None,
            accepted_suggestions=[],
            uncited_refs=uncited_refs,
            _results_internal=results,   # keep real objects for saving
        )
    except Exception as e:
        _set_state(status="error", message=str(e), error=str(e), progress=0)


# ---------------------------------------------------------------------------
# Cross-platform file picker
# macOS uses AppleScript; Windows/Linux use tkinter filedialog.
# ---------------------------------------------------------------------------

def _native_pick_file(prompt: str, file_types: str = "docx files") -> Optional[str]:
    """Open a native file-open dialog. Returns the chosen path or None."""
    if sys.platform == "darwin":
        if file_types == "docx":
            type_clause = (
                'of type {"docx", "com.microsoft.word.doc", '
                '"org.oasis-open.opendocument.text", "odt"}'
            )
        elif file_types == "sqlite":
            type_clause = 'of type {"sqlite", "db", "public.database"}'
        else:
            type_clause = ""
        script = f'POSIX path of (choose file with prompt "{prompt}" {type_clause})'
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True, timeout=60
            )
            path = result.stdout.strip()
            return path if path else None
        except Exception:
            return None
    else:
        # Windows / Linux
        import tkinter as tk
        from tkinter import filedialog
        if file_types == "docx":
            ftypes = [
                ("Documents", "*.docx *.odt"),
                ("Word Documents", "*.docx"),
                ("OpenDocument Text", "*.odt"),
                ("All files", "*.*"),
            ]
        elif file_types == "sqlite":
            ftypes = [("SQLite Databases", "*.sqlite *.db"), ("All files", "*.*")]
        else:
            ftypes = [("All files", "*.*")]
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes("-topmost", 1)
        path = filedialog.askopenfilename(title=prompt, filetypes=ftypes)
        root.destroy()
        return path if path else None


def _native_save_file(prompt: str, default_name: str, folder: str) -> Optional[str]:
    """Open a native file-save dialog. Returns the chosen path or None."""
    if sys.platform == "darwin":
        # Strip characters that could break AppleScript string literals
        safe_name   = re.sub(r'["\\\r\n]', '', default_name)
        safe_folder = re.sub(r'["\\\r\n]', '', folder)
        script = (
            f'POSIX path of (choose file name with prompt "{prompt}" '
            f'default name "{safe_name}" default location POSIX file "{safe_folder}")'
        )
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True, timeout=60
            )
            path = result.stdout.strip()
            return path if path else None
        except Exception:
            return None
    else:
        # Windows / Linux
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes("-topmost", 1)
        path = filedialog.asksaveasfilename(
            title=prompt,
            initialfile=default_name,
            initialdir=folder,
            defaultextension=Path(default_name).suffix,
        )
        root.destroy()
        return path if path else None


# ---------------------------------------------------------------------------
# Embedded HTML/CSS/JS
# ---------------------------------------------------------------------------

_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Identifyer for Zotero</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
         background: #f0f2f5; color: #1a1a2e; min-height: 100vh; }
  header { background: #2c3e50; color: #fff; padding: 16px 28px;
           display: flex; align-items: center; gap: 16px; }
  header h1 { font-size: 1.35rem; font-weight: 700; }
  header p  { font-size: 0.85rem; color: #bdc3c7; }
  .card { background: #fff; border-radius: 10px; box-shadow: 0 1px 4px rgba(0,0,0,.1);
          padding: 20px 24px; margin: 18px 24px; }
  .field-row { display: flex; align-items: center; gap: 10px; margin-bottom: 12px; }
  .field-row label { min-width: 130px; font-size: 0.9rem; color: #555; text-align: right; }
  .field-row input { flex: 1; padding: 7px 10px; border: 1px solid #ccc;
                     border-radius: 6px; font-size: 0.9rem; }
  .field-row button { padding: 7px 14px; border: none; border-radius: 6px;
                      background: #eef0f2; cursor: pointer; font-size: 0.85rem; }
  .field-row button:hover { background: #dde1e6; }
  .status-badge { font-size: 0.85rem; margin-top: 4px; }
  .status-badge.ok  { color: #27ae60; }
  .status-badge.off { color: #95a5a6; }
  .action-row { display: flex; align-items: center; gap: 16px; margin-top: 4px; }
  #analyzeBtn { padding: 9px 28px; background: #2980b9; color: #fff;
                border: none; border-radius: 7px; font-size: 0.95rem;
                cursor: pointer; font-weight: 600; }
  #analyzeBtn:hover { background: #1f6f9e; }
  #analyzeBtn:disabled { background: #95a5a6; cursor: default; }
  progress { flex: 1; height: 8px; max-width: 320px; }
  #statusMsg { font-size: 0.85rem; color: #666; }
  .tabs { display: flex; gap: 0; border-bottom: 2px solid #e0e0e0; }
  .tab { padding: 9px 20px; cursor: pointer; font-size: 0.9rem; color: #666;
          border: none; background: none; border-bottom: 3px solid transparent;
          margin-bottom: -2px; }
  .tab.active { color: #2980b9; border-bottom-color: #2980b9; font-weight: 600; }
  .tab-pane { display: none; }
  .tab-pane.active { display: block; }
  table { width: 100%; border-collapse: collapse; font-size: 0.88rem; }
  th { background: #f7f8fa; padding: 8px 10px; text-align: left;
       font-weight: 600; color: #444; border-bottom: 2px solid #e0e0e0; }
  td { padding: 7px 10px; border-bottom: 1px solid #f0f0f0; vertical-align: middle; }
  tr:hover td { background: #f9fbfd; }
  tr.sug-row td    { background: #fffbf0; }
  tr.sug-row:hover td { background: #fff5d9; }
  tr.amb-row td    { background: #f5f0ff; }
  tr.amb-row:hover td { background: #ede5ff; }
  tr.accepted-row td  { background: #f0faf4; }
  .badge-match { color: #27ae60; font-weight: 700; }
  .badge-no    { color: #c0392b; font-weight: 700; }
  .badge-sug   { color: #e67e22; font-weight: 700; }
  .badge-amb   { color: #8e44ad; font-weight: 700; }
  .badge-low   { color: #e67e22; font-weight: 700; }
  .conf        { font-size: 0.8rem; color: #888; white-space: nowrap; }
  .sug-info    { font-size: 0.8rem; color: #7f6000; margin-top: 3px; }
  .cand-list   { margin-top: 6px; border-left: 3px solid #c39bd3; padding-left: 8px; }
  .cand-item   { font-size: 0.8rem; color: #444; padding: 3px 0;
                 display: flex; align-items: center; gap: 6px; }
  .cand-item span { flex: 1; }
  /* small action buttons */
  .btn-sm { padding: 2px 9px; border: none; border-radius: 4px; cursor: pointer;
            font-size: 0.78rem; white-space: nowrap; }
  .btn-accept  { background: #27ae60; color: #fff; }
  .btn-accept:hover { background: #1e8449; }
  .btn-dismiss { background: #95a5a6; color: #fff; }
  .btn-dismiss:hover { background: #717d7e; }
  .btn-use     { background: #8e44ad; color: #fff; }
  .btn-use:hover { background: #6c3483; }
  .btn-unmatch { background: #e74c3c; color: #fff; }
  .btn-unmatch:hover { background: #c0392b; }
  .btn-retract { background: #e67e22; color: #fff; }
  .btn-retract:hover { background: #ca6f1e; }
  .btn-restore { background: #2980b9; color: #fff; }
  .btn-restore:hover { background: #1f6f9e; }
  .footer-row  { display: flex; gap: 12px; margin: 18px 24px 24px; align-items: center; }
  .btn-action  { padding: 9px 22px; border: none; border-radius: 7px;
                 cursor: pointer; font-size: 0.9rem; font-weight: 600; }
  .btn-save    { background: #27ae60; color: #fff; }
  .btn-save:hover { background: #1e8449; }
  .btn-export  { background: #7f8c8d; color: #fff; }
  .btn-export:hover { background: #626567; }
  .btn-action:disabled { background: #bdc3c7; cursor: default; }
  .note { font-size: 0.8rem; color: #95a5a6; }
  #importError { background: #fdf2f2; border: 1px solid #e74c3c; border-radius: 8px;
                 padding: 14px 18px; margin: 18px 24px; color: #c0392b; font-size: 0.9rem; }
  /* Snap sliders */
  .slider-row { margin-bottom: 10px; align-items: flex-start; }
  .snap-slider-wrap { flex: 1; max-width: 220px; }
  .snap-slider { width: 100%; cursor: pointer; accent-color: #2980b9;
                 margin: 2px 0 1px; }
  .snap-labels { display: flex; justify-content: space-between;
                 font-size: 0.7rem; color: #999; padding: 0 1px; }
  .slider-section-title { font-size: 0.78rem; font-weight: 600; color: #888;
                          text-transform: uppercase; letter-spacing: 0.05em;
                          margin: 14px 0 8px 0; padding-top: 10px;
                          border-top: 1px solid #f0f0f0; }
</style>
</head>
<body>

<header>
  <div>
    <h1>Identifyer for Zotero</h1>
    <p>Match plain-text citations to your Zotero library &amp; convert them to field codes</p>
  </div>
</header>

<div id="importError" style="display:none"></div>

<div class="card">
  <div class="field-row">
    <label>Document</label>
    <input id="docxPath" type="text" placeholder="/path/to/document.docx or .odt" />
    <button onclick="browse('docx')">Browse…</button>
  </div>
  <div class="field-row">
    <label>Zotero Database</label>
    <input id="dbPath" type="text" />
    <button onclick="browse('db')">Browse…</button>
  </div>
  <div class="field-row">
    <label>Collection</label>
    <select id="collectionSelect" style="flex:1; padding:7px 10px;
            border:1px solid #ccc; border-radius:6px; font-size:0.9rem;
            background:#fff;">
      <option value="">Whole library</option>
    </select>
  </div>
  <div class="field-row" style="margin-bottom:0">
    <label>Zotero app</label>
    <span id="zoteroStatus" class="status-badge off">○ Checking…</span>
  </div>
  <div style="margin-left:140px">
    <div class="slider-section-title">Matching sensitivity</div>
  </div>
  <div class="field-row slider-row">
    <label>Author match</label>
    <div class="snap-slider-wrap">
      <input type="range" class="snap-slider" id="sliderAuthor"
             min="0" max="2" step="1" value="1">
      <div class="snap-labels"><span>Wide</span><span>Normal</span><span>Narrow</span></div>
    </div>
  </div>
  <div class="field-row slider-row">
    <label>Candidates</label>
    <div class="snap-slider-wrap">
      <input type="range" class="snap-slider" id="sliderCandidate"
             min="0" max="2" step="1" value="1">
      <div class="snap-labels"><span>Wide</span><span>Normal</span><span>Narrow</span></div>
    </div>
  </div>
  <div class="field-row slider-row" style="margin-bottom:0">
    <label>Year tolerance</label>
    <div class="snap-slider-wrap">
      <input type="range" class="snap-slider" id="sliderTolerance"
             min="0" max="2" step="1" value="1">
      <div class="snap-labels"><span>Wide (±2)</span><span>Normal (±1)</span><span>Narrow (±0)</span></div>
    </div>
  </div>
</div>

<div class="card">
  <div class="action-row">
    <button id="analyzeBtn" onclick="analyze()">Analyze Document</button>
    <progress id="progressBar" value="0" max="100"></progress>
    <span id="statusMsg">Ready.</span>
  </div>
</div>

<div class="card" id="resultsCard" style="padding: 0; overflow:hidden;">
  <div class="tabs">
    <button class="tab active" onclick="switchTab('unmatched')">Unmatched (0)</button>
    <button class="tab" onclick="switchTab('matched')">Matched (0)</button>
    <button class="tab" onclick="switchTab('all')">All Citations (0)</button>
  </div>
  <div id="tab-unmatched" class="tab-pane active" style="padding:0">
    <table><thead><tr><th>Citation</th><th>Location</th><th></th></tr></thead>
    <tbody id="tbody-unmatched"><tr><td colspan="3" style="color:#aaa;padding:18px">
      Run an analysis to see results.</td></tr></tbody></table>
  </div>
  <div id="tab-matched" class="tab-pane" style="padding:0">
    <table><thead><tr><th>Citation</th><th>Zotero Key</th><th>Title</th><th>Conf.</th><th></th></tr></thead>
    <tbody id="tbody-matched"></tbody></table>
  </div>
  <div id="tab-all" class="tab-pane" style="padding:0">
    <table><thead><tr><th>Citation</th><th>Status</th><th>Zotero Key</th><th>Title</th><th>Conf.</th></tr></thead>
    <tbody id="tbody-all"></tbody></table>
  </div>
</div>

<div class="footer-row">
  <button class="btn-action btn-save" id="saveBtn" disabled onclick="saveDoc()">
    Save Zotero Document
  </button>
  <button class="btn-action btn-export" id="exportBtn" disabled onclick="exportUnmatched()">
    Export Unmatched to .txt
  </button>
  <span class="note">Save creates a new file with Zotero field codes — your original is unchanged.</span>
</div>

<script>
let pollTimer = null;
let allRows = [];                        // full results array from server
let acceptedDisplays  = new Set();       // display texts of accepted year-suggestions
let dismissedDisplays = new Set();       // display texts of dismissed year-suggestions
let manuallyUnmatched = new Set();       // display texts of manually unmatched confirmed matches
let selectedCandidates = new Map();      // display text → {candidateIndex, key, title, year, authors, confidence}
let savedCollectionId = null;            // restored from /state, applied once dropdown is filled

// ---- Initialise ----
window.onload = () => {
  fetch('/state').then(r => r.json()).then(s => {
    document.getElementById('docxPath').value = s.docx_path || '';
    const dbInput = document.getElementById('dbPath');
    dbInput.value = s.db_path || '';
    savedCollectionId = s.collection_id || null;
    if (dbInput.value) loadCollections(dbInput.value);
    // Reload the collection dropdown whenever the DB path is edited manually.
    dbInput.addEventListener('change', () => {
      savedCollectionId = null;  // don't re-apply after user switches DBs
      loadCollections(dbInput.value);
    });
  });
  checkImportError();
  checkZoteroStatus();
  setInterval(checkZoteroStatus, 30000);
};

function loadCollections(db) {
  const sel = document.getElementById('collectionSelect');
  sel.innerHTML = '<option value="">Whole library</option>';
  if (!db) return;
  fetch('/collections?db=' + encodeURIComponent(db))
    .then(r => r.json())
    .then(d => {
      if (!d.ok || !d.collections) return;
      for (const c of d.collections) {
        const opt = document.createElement('option');
        opt.value = c.id;
        opt.textContent = c.path + '  (' + c.item_count + ')';
        sel.appendChild(opt);
      }
      if (savedCollectionId != null) {
        sel.value = String(savedCollectionId);
        // Clear so that a later dbPath edit doesn't re-apply the stale ID.
        savedCollectionId = null;
      }
    })
    .catch(() => {});
}

function checkImportError() {
  fetch('/import-error').then(r => r.json()).then(d => {
    if (d.error) {
      const el = document.getElementById('importError');
      el.style.display = 'block';
      el.innerHTML = '<b>Missing dependency:</b> ' + d.error +
        '<br>Run in terminal: <code>python3 -m pip install -r requirements.txt</code>';
    }
  });
}

function checkZoteroStatus() {
  fetch('/zotero-ping').then(r => r.json()).then(d => {
    const el = document.getElementById('zoteroStatus');
    if (d.running) { el.textContent = '● Running';      el.className = 'status-badge ok'; }
    else           { el.textContent = '○ Not detected'; el.className = 'status-badge off'; }
  }).catch(() => {});
}

// ---- Browse ----
function browse(type) {
  fetch('/browse', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({type})
  }).then(r => r.json()).then(d => {
    if (d.path) {
      if (type === 'docx') {
        document.getElementById('docxPath').value = d.path;
      } else {
        document.getElementById('dbPath').value = d.path;
        savedCollectionId = null;
        loadCollections(d.path);
      }
    }
  });
}

// ---- Matching sensitivity sliders ----
function getThresholds() {
  const authorMap    = [75, 82, 90];   // Wide / Normal / Narrow
  const candidateMap = [50, 60, 70];
  const toleranceMap = [2,  1,  0];   // Wide=±2, Normal=±1, Narrow=±0
  return {
    author_threshold:    authorMap[+document.getElementById('sliderAuthor').value],
    candidate_threshold: candidateMap[+document.getElementById('sliderCandidate').value],
    year_tolerance:      toleranceMap[+document.getElementById('sliderTolerance').value],
  };
}

// ---- Analyze ----
function analyze() {
  const docx = document.getElementById('docxPath').value.trim();
  const db   = document.getElementById('dbPath').value.trim();
  if (!docx) { alert('Please select a document (.docx or .odt).'); return; }
  if (!db)   { alert('Please select the Zotero database.'); return; }

  document.getElementById('analyzeBtn').disabled = true;
  document.getElementById('saveBtn').disabled    = true;
  document.getElementById('exportBtn').disabled  = true;
  allRows = [];
  acceptedDisplays.clear();
  dismissedDisplays.clear();
  manuallyUnmatched.clear();
  selectedCandidates.clear();
  clearResults();

  const cidRaw = document.getElementById('collectionSelect').value;
  const collection_id = cidRaw ? parseInt(cidRaw, 10) : null;

  fetch('/analyze', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({docx_path: docx, db_path: db, collection_id,
                          ...getThresholds()})
  }).then(() => {
    pollTimer = setInterval(pollStatus, 400);
  });
}

function pollStatus() {
  fetch('/status').then(r => r.json()).then(s => {
    document.getElementById('progressBar').value = s.progress;
    document.getElementById('statusMsg').textContent = s.message;

    if (s.status === 'done') {
      clearInterval(pollTimer);
      document.getElementById('analyzeBtn').disabled = false;
      document.getElementById('saveBtn').disabled    = false;
      document.getElementById('exportBtn').disabled  = false;
      allRows = s.results || [];
      _render();
    } else if (s.status === 'error') {
      clearInterval(pollTimer);
      document.getElementById('analyzeBtn').disabled = false;
      alert('Error: ' + s.message);
    }
  });
}

// ---- User actions (all operate by display text so every occurrence moves together) ----

function acceptSuggestion(idx) {
  const row = allRows.find(r => r.index === idx);
  if (row) { acceptedDisplays.add(row.display); dismissedDisplays.delete(row.display); }
  _render();
}

function dismissSuggestion(idx) {
  const row = allRows.find(r => r.index === idx);
  if (row) { dismissedDisplays.add(row.display); acceptedDisplays.delete(row.display); }
  _render();
}

function retractAcceptance(idx) {
  const row = allRows.find(r => r.index === idx);
  if (row) acceptedDisplays.delete(row.display);
  _render();
}

function selectCandidate(idx, candIdx) {
  const row = allRows.find(r => r.index === idx);
  if (!row || !row.candidates || candIdx >= row.candidates.length) return;
  const c = row.candidates[candIdx];
  selectedCandidates.set(row.display, {
    candidateIndex: candIdx, key: c.key, title: c.title,
    year: c.year, authors: c.authors, confidence: c.confidence
  });
  manuallyUnmatched.delete(row.display);
  _render();
}

function deselectCandidate(idx) {
  const row = allRows.find(r => r.index === idx);
  if (row) selectedCandidates.delete(row.display);
  _render();
}

function unmatchItem(idx) {
  const row = allRows.find(r => r.index === idx);
  if (row) manuallyUnmatched.add(row.display);
  _render();
}

function restoreMatch(idx) {
  const row = allRows.find(r => r.index === idx);
  if (row) manuallyUnmatched.delete(row.display);
  _render();
}

// ---- Render results ----
function clearResults() {
  ['unmatched','matched','all'].forEach(t => {
    document.getElementById('tbody-' + t).innerHTML = '';
  });
  updateTabTitles(0, 0, 0);
}

function _dedup(rows) {
  const seen = new Set();
  return rows.filter(r => {
    if (seen.has(r.display)) return false;
    seen.add(r.display);
    return true;
  });
}

function _render() {
  const rows = allRows;
  if (!rows.length) return;

  const unmatchedRows  = [];
  const suggestionRows = [];
  const ambiguousRows  = [];
  const matchedRows    = [];

  rows.forEach(r => {
    if (r.matched) {
      if (manuallyUnmatched.has(r.display)) {
        unmatchedRows.push({...r, _manually_unmatched: true});
      } else {
        matchedRows.push({...r, _accepted: false, _selected: false});
      }
    } else if (r.is_suggestion) {
      if (acceptedDisplays.has(r.display)) {
        matchedRows.push({...r, _accepted: true, _selected: false});
      } else if (dismissedDisplays.has(r.display)) {
        unmatchedRows.push({...r, _dismissed: true});
      } else {
        suggestionRows.push(r);
      }
    } else if (r.is_ambiguous) {
      if (selectedCandidates.has(r.display) && !manuallyUnmatched.has(r.display)) {
        const sel = selectedCandidates.get(r.display);
        matchedRows.push({...r, _accepted: false, _selected: true, _sel: sel});
      } else if (manuallyUnmatched.has(r.display)) {
        unmatchedRows.push({...r, _manually_unmatched: true});
      } else {
        ambiguousRows.push(r);
      }
    } else {
      unmatchedRows.push(r);
    }
  });

  const dedupUnmatched   = _dedup(unmatchedRows);
  const dedupSuggestions = _dedup(suggestionRows);
  const dedupAmbiguous   = _dedup(ambiguousRows);
  const dedupMatched     = _dedup(matchedRows);

  // --- Unmatched tab ---
  const tbU = document.getElementById('tbody-unmatched');
  const umRows = [...dedupAmbiguous, ...dedupSuggestions, ...dedupUnmatched];
  if (!umRows.length) {
    tbU.innerHTML = '<tr><td colspan="3" style="color:#27ae60;padding:14px">All citations matched!</td></tr>';
  } else {
    tbU.innerHTML = umRows.map(r => {
      if (r.is_ambiguous && !r._manually_unmatched) {
        const candHtml = (r.candidates || []).map((c, ci) =>
          `<div class="cand-item">
            <span>[${esc(c.key)}] ${esc(c.authors.join(', '))} (${esc(c.year)}) — ${esc(c.title)} <span class="conf">${c.confidence}%</span></span>
            <button class="btn-sm btn-use" onclick="selectCandidate(${r.index},${ci})">Accept</button>
          </div>`
        ).join('');
        const refHint = r.ref_text ? `<div style="font-size:0.78rem;color:#666;margin-top:4px">Ref: ${esc(r.ref_text.substring(0,120))}${r.ref_text.length>120?'…':''}</div>` : '';
        return `<tr class="amb-row">
          <td>
            <span class="badge-amb">?</span> ${esc(r.display)}${refHint}
            <div class="cand-list">${candHtml}</div>
          </td>
          <td>${r.in_footnote ? 'footnote' : 'para ' + (r.para+1)}</td>
          <td><button class="btn-sm btn-unmatch" onclick="unmatchItem(${r.index})">Unmatch</button></td>
        </tr>`;
      } else if (r.is_suggestion && !r._dismissed) {
        const sugInfo = `Similar: [${esc(r.zotero_key)}] ${esc(r.zotero_authors.join(', '))} (${esc(r.zotero_year)}) — ${esc(r.zotero_title)}`;
        return `<tr class="sug-row">
          <td>
            <span class="badge-sug">⚠</span> ${esc(r.display)}<br>
            <span class="sug-info">→ ${sugInfo}</span>
          </td>
          <td>${r.in_footnote ? 'footnote' : 'para ' + (r.para+1)}</td>
          <td>
            <button class="btn-sm btn-accept"  onclick="acceptSuggestion(${r.index})">Accept</button>
            <button class="btn-sm btn-dismiss" onclick="dismissSuggestion(${r.index})">Dismiss</button>
          </td>
        </tr>`;
      } else if (r._manually_unmatched) {
        return `<tr>
          <td><span class="badge-no">✗</span> ${esc(r.display)}</td>
          <td>${r.in_footnote ? 'footnote' : 'para ' + (r.para+1)}</td>
          <td><button class="btn-sm btn-restore" onclick="restoreMatch(${r.index})">Restore</button></td>
        </tr>`;
      } else {
        const refHint = r.ref_text ? `<div style="font-size:0.78rem;color:#666;margin-top:4px">Ref: ${esc(r.ref_text.substring(0,120))}${r.ref_text.length>120?'…':''}</div>` : '';
        // Last-resort fallback: closest bibliography entry (preferred) or
        // closest Zotero item ignoring year. Signals to the user "this is
        // the only candidate we could find — if it looks unrelated, the
        // reference is probably missing from both."
        let lastResortHint = '';
        if (r.is_last_resort) {
          const bib = r.bibliography_hint
            ? `<div style="font-size:0.78rem;color:#b8860b;margin-top:4px">Bibliography (closest): ${esc(r.bibliography_hint.substring(0,140))}${r.bibliography_hint.length>140?'…':''}</div>`
            : '';
          const zot = r.zotero_key
            ? `<div style="font-size:0.78rem;color:#b8860b;margin-top:4px">Zotero (year ignored): [${esc(r.zotero_key)}] ${esc(r.zotero_authors.join(', '))} (${esc(r.zotero_year)}) — ${esc(r.zotero_title)}</div>`
            : '';
          lastResortHint = bib + zot;
        }
        return `<tr>
          <td><span class="badge-no">✗</span> ${esc(r.display)}${refHint}${lastResortHint}</td>
          <td>${r.in_footnote ? 'footnote' : 'para ' + (r.para+1)}</td>
          <td></td>
        </tr>`;
      }
    }).join('');
  }

  // --- Matched tab ---
  const tbM = document.getElementById('tbody-matched');
  if (!dedupMatched.length) {
    tbM.innerHTML = '<tr><td colspan="5" style="color:#aaa;padding:14px">No matches found.</td></tr>';
  } else {
    const matchedHtml = dedupMatched.map(r => {
      let rowClass = '', badge, note, key, title, conf, actionBtn;
      if (r._accepted) {
        rowClass = ' class="accepted-row"';
        badge = `<span class="badge-sug">⚠→✓</span>`;
        note  = ' <span style="font-size:0.75rem;color:#27ae60">(accepted)</span>';
        key   = r.zotero_key; title = r.zotero_title; conf = r.confidence;
        actionBtn = `<button class="btn-sm btn-retract" onclick="retractAcceptance(${r.index})">Retract</button>`;
      } else if (r._selected) {
        rowClass = ' class="accepted-row"';
        badge = `<span class="badge-amb">?→✓</span>`;
        note  = ' <span style="font-size:0.75rem;color:#8e44ad">(selected)</span>';
        key   = r._sel.key; title = r._sel.title; conf = r._sel.confidence;
        actionBtn = `<button class="btn-sm btn-retract" onclick="deselectCandidate(${r.index})">Retract</button>`;
      } else {
        badge = `<span class="${r.confidence >= 90 ? 'badge-match' : 'badge-low'}">✓</span>`;
        note  = '';
        key   = r.zotero_key; title = r.zotero_title; conf = r.confidence;
        actionBtn = `<button class="btn-sm btn-unmatch" onclick="unmatchItem(${r.index})">Unmatch</button>`;
      }
      return `<tr${rowClass}>
        <td>${badge} ${esc(r.display)}${note}</td>
        <td><code>${esc(key)}</code></td>
        <td>${esc(title)}</td>
        <td class="conf">${conf}%</td>
        <td>${actionBtn}</td>
      </tr>`;
    }).join('');
    tbM.innerHTML = matchedHtml;
  }

  // --- All tab ---
  const tbA = document.getElementById('tbody-all');
  tbA.innerHTML = rows.map(r => {
    let badge, statusText;
    const isManualUnmatched = r.matched && manuallyUnmatched.has(r.display);
    if (r.matched && !isManualUnmatched) {
      badge = `<span class="${r.confidence >= 90 ? 'badge-match' : 'badge-low'}">✓</span>`;
      statusText = r.confidence + '%';
    } else if (r.is_suggestion && acceptedDisplays.has(r.display)) {
      badge = '<span class="badge-sug">⚠→✓</span>';
      statusText = r.confidence + '% (accepted)';
    } else if (r.is_ambiguous && selectedCandidates.has(r.display) && !manuallyUnmatched.has(r.display)) {
      badge = '<span class="badge-amb">?→✓</span>';
      statusText = selectedCandidates.get(r.display).confidence + '% (selected)';
    } else if (r.is_ambiguous) {
      badge = '<span class="badge-amb">?</span>';
      statusText = 'ambiguous';
    } else if (r.is_suggestion) {
      badge = dismissedDisplays.has(r.display) ? '<span class="badge-no">✗</span>' : '<span class="badge-sug">⚠</span>';
      statusText = dismissedDisplays.has(r.display) ? '—' : 'suggestion';
    } else {
      badge = '<span class="badge-no">✗</span>';
      statusText = isManualUnmatched ? '(unmatched)' : '—';
    }
    return `<tr>
      <td>${esc(r.display)}</td>
      <td>${badge}</td>
      <td><code>${esc(r.zotero_key)}</code></td>
      <td>${esc(r.zotero_title)}</td>
      <td class="conf">${statusText}</td>
    </tr>`;
  }).join('');

  const umCount = dedupUnmatched.length + dedupSuggestions.length + dedupAmbiguous.length;
  updateTabTitles(umCount, dedupMatched.length, rows.length);

  if (!allRows._tabSet) {
    allRows._tabSet = true;
    switchTab(umCount > 0 ? 'unmatched' : 'matched');
  }
}

function updateTabTitles(u, m, total) {
  const tabs = document.querySelectorAll('.tab');
  tabs[0].textContent = `Unmatched (${u})`;
  tabs[1].textContent = `Matched (${m})`;
  tabs[2].textContent = `All Citations (${total})`;
}

function switchTab(name) {
  document.querySelectorAll('.tab').forEach((t, i) => {
    const names = ['unmatched','matched','all'];
    t.classList.toggle('active', names[i] === name);
  });
  document.querySelectorAll('.tab-pane').forEach(p => {
    p.classList.toggle('active', p.id === 'tab-' + name);
  });
}

// ---- Save / Export ----
function _serverState() {
  const accepted_suggestions = [];
  const selected_candidates  = {};
  const manually_unmatched   = [];
  for (const r of allRows) {
    if (r.is_suggestion && acceptedDisplays.has(r.display))
      accepted_suggestions.push(r.index);
    if (r.is_ambiguous && selectedCandidates.has(r.display))
      selected_candidates[r.index] = selectedCandidates.get(r.display).candidateIndex;
    if (r.matched && manuallyUnmatched.has(r.display))
      manually_unmatched.push(r.index);
  }
  return {accepted_suggestions, selected_candidates, manually_unmatched};
}

function saveDoc() {
  const docx = document.getElementById('docxPath').value.trim();
  fetch('/save', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({docx_path: docx, ..._serverState()})
  }).then(r => r.json()).then(d => {
    if (d.ok) alert('Saved to:\\n' + d.path + '\\n\\n' + d.count + ' citation(s) replaced with Zotero field codes.\\n\\nOpen in Word and use the Zotero plugin to refresh citations.');
    else if (d.error !== 'Cancelled.') alert('Save failed: ' + d.error);
  });
}

function exportUnmatched() {
  const docx = document.getElementById('docxPath').value.trim();
  fetch('/export', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({docx_path: docx, ..._serverState()})
  }).then(r => r.json()).then(d => {
    if (d.ok) alert('Exported ' + d.count + ' unmatched citations to:\\n' + d.path);
    else if (d.error === 'all_matched') alert('All citations are matched — nothing to export.');
    else if (d.error !== 'Cancelled.') alert('Export failed: ' + d.error);
  });
}

function esc(s) {
  if (!s) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# HTTP request handler
# ---------------------------------------------------------------------------

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # suppress access log noise

    def _send_json(self, data: dict, code: int = 200):
        body = json.dumps(data).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_html(self, html: str):
        body = html.encode()
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length", 0))
        if length:
            return json.loads(self.rfile.read(length))
        return {}

    # ------------------------------------------------------------------
    # GET
    # ------------------------------------------------------------------
    def do_GET(self):
        path = urlparse(self.path).path

        if path == "/" or path == "/index.html":
            self._send_html(_HTML)

        elif path == "/state":
            s = _get_state()
            self._send_json({
                "docx_path": s["docx_path"],
                "db_path": s["db_path"],
                "collection_id": load_settings().get("collection_id"),
            })

        elif path == "/status":
            s = _get_state()
            self._send_json({
                "status": s["status"],
                "progress": s["progress"],
                "message": s["message"],
                "results": s.get("results"),
                "error": s.get("error"),
            })

        elif path == "/zotero-ping":
            if IMPORT_ERROR:
                self._send_json({"running": False})
            else:
                running = ZoteroClient(_get_state()["db_path"]).ping_zotero()
                self._send_json({"running": running})

        elif path == "/import-error":
            self._send_json({"error": IMPORT_ERROR})

        elif path == "/collections":
            if IMPORT_ERROR:
                self._send_json({"ok": False, "error": IMPORT_ERROR, "collections": []})
                return
            # Optional ?db=... override; falls back to current state.
            from urllib.parse import parse_qs
            qs = parse_qs(urlparse(self.path).query)
            db = (qs.get("db", [""])[0] or _get_state()["db_path"]).strip()
            if not db:
                self._send_json({"ok": False, "error": "No DB path.", "collections": []})
                return
            try:
                cols = ZoteroClient(db).list_collections()
                self._send_json({"ok": True, "collections": cols})
            except Exception as e:
                self._send_json({"ok": False, "error": str(e), "collections": []})

        else:
            self.send_error(404)

    # ------------------------------------------------------------------
    # POST
    # ------------------------------------------------------------------
    def do_POST(self):
        path = urlparse(self.path).path
        data = self._read_json()

        if path == "/browse":
            ftype = data.get("type", "docx")
            if ftype == "docx":
                picked = _native_pick_file("Select Word Document", "docx")
            else:
                picked = _native_pick_file("Select Zotero Database", "sqlite")
            self._send_json({"path": picked or ""})

        elif path == "/analyze":
            if IMPORT_ERROR:
                self._send_json({"ok": False, "error": IMPORT_ERROR})
                return
            if _get_state()["status"] == "running":
                self._send_json({"ok": False, "error": "Already running"})
                return
            docx             = data.get("docx_path", "").strip()
            db               = data.get("db_path", "").strip()
            author_thresh    = int(data.get("author_threshold",    82))
            candidate_thresh = int(data.get("candidate_threshold", 60))
            year_tol         = int(data.get("year_tolerance",       1))
            # collection_id may be an integer, missing, or the sentinel 0 / ""
            # meaning "whole library".
            raw_cid = data.get("collection_id")
            try:
                cid: Optional[int] = int(raw_cid) if raw_cid else None
            except (TypeError, ValueError):
                cid = None
            # Persist DB path and collection choice so they're prefilled next time.
            if db:
                update_setting("db_path", db)
            update_setting("collection_id", cid if cid else None)
            _set_state(
                docx_path=docx,
                db_path=db,
                status="running",
                progress=0,
                message="Starting…",
                results=None,
                error=None,
                accepted_suggestions=[],
                _results_internal=None,
            )
            threading.Thread(
                target=_run_analysis,
                args=(docx, db, author_thresh, candidate_thresh, year_tol, cid),
                daemon=True,
            ).start()
            self._send_json({"ok": True})

        elif path == "/save":
            s = _get_state()
            internal = s.get("_results_internal")
            if not internal:
                self._send_json({"ok": False, "error": "No analysis results. Run Analyze first."})
                return
            docx = data.get("docx_path", s["docx_path"]).strip()
            p = Path(docx)
            out_path = _native_save_file(
                "Save Zotero Document As…",
                p.stem + "_zotero" + p.suffix,
                str(p.parent),
            )
            if not out_path:
                self._send_json({"ok": False, "error": "Cancelled."})
                return
            try:
                accepted          = set(data.get("accepted_suggestions", []))
                sel_cands         = {int(k): int(v) for k, v in
                                     data.get("selected_candidates", {}).items()}
                manually_unmatched = set(data.get("manually_unmatched", []))

                # Build an effective results list applying user overrides
                effective = []
                for i, r in enumerate(internal):
                    if i in manually_unmatched:
                        effective.append(MatchResult(
                            citation=r.citation, zotero_item=None,
                            confidence=0.0, matched=False))
                    elif i in sel_cands and r.candidates:
                        ci = sel_cands[i]
                        if 0 <= ci < len(r.candidates):
                            chosen_item, chosen_conf = r.candidates[ci]
                            effective.append(MatchResult(
                                citation=r.citation, zotero_item=chosen_item,
                                confidence=chosen_conf, matched=True))
                        else:
                            effective.append(r)
                    else:
                        effective.append(r)

                count = write_zotero_document(docx, out_path, effective,
                                              accepted_suggestions=accepted)
                self._send_json({"ok": True, "path": out_path, "count": count})
            except Exception as e:
                self._send_json({"ok": False, "error": str(e)})

        elif path == "/export":
            s = _get_state()
            results = s.get("results")
            if not results:
                self._send_json({"ok": False, "error": "No results yet."})
                return
            accepted           = set(data.get("accepted_suggestions", []))
            sel_cands          = {int(k) for k in data.get("selected_candidates", {}).keys()}
            manually_unmatched = set(data.get("manually_unmatched", []))

            unmatched = [
                r for r in results
                if not (r["matched"] and r["index"] not in manually_unmatched)
                and not (r["is_suggestion"] and r["index"] in accepted)
                and not (r["is_ambiguous"] and r["index"] in sel_cands)
            ]
            if not unmatched:
                self._send_json({"ok": False, "error": "all_matched"})
                return
            docx = data.get("docx_path", s["docx_path"]).strip()
            p = Path(docx)
            out_path = _native_save_file(
                "Export Unmatched Citations As…",
                p.stem + "_unmatched.txt",
                str(p.parent),
            )
            if not out_path:
                self._send_json({"ok": False, "error": "Cancelled."})
                return
            try:
                # Deduplicate by display text, keeping first occurrence
                seen_displays: set = set()
                deduped: list = []
                for r in unmatched:
                    if r["display"] not in seen_displays:
                        seen_displays.add(r["display"])
                        deduped.append(r)

                lines = [
                    f"Unmatched citations from: {docx}",
                    f"Total unmatched: {len(deduped)}",
                    "",
                ]
                for r in deduped:
                    loc = "footnote" if r["in_footnote"] else f"para {r['para'] + 1}"
                    sug = f"  [suggestion: {r['zotero_key']} ({r['zotero_year']})]" \
                          if r["is_suggestion"] else ""
                    line = f"{r['display']}  [{loc}]{sug}"
                    # Append the full reference entry if available (numbered or author-year)
                    if r.get("ref_text"):
                        line += f"\n    → {r['ref_text']}"
                    lines.append(line)
                    lines.append("")

                uncited_refs = s.get("uncited_refs", [])
                if uncited_refs:
                    lines += [
                        "",
                        "=" * 60,
                        f"References listed but not cited in text: {len(uncited_refs)}",
                        "",
                    ]
                    lines.extend(ref for entry in uncited_refs for ref in (entry, ""))

                Path(out_path).write_text("\n".join(lines), encoding="utf-8")
                self._send_json({"ok": True, "path": out_path, "count": len(deduped)})
            except Exception as e:
                self._send_json({"ok": False, "error": str(e)})

        else:
            self.send_error(404)


# ---------------------------------------------------------------------------
# Server startup
# ---------------------------------------------------------------------------

def find_free_port(start: int = 7474) -> int:
    import socket
    for port in range(start, start + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return start


def main():
    port = find_free_port()
    url = f"http://localhost:{port}"

    server = HTTPServer(("127.0.0.1", port), Handler)

    print(f"\n  Identifyer for Zotero running at {url}")
    print("  Press Ctrl-C to quit.\n")

    # Open browser after a short delay
    def _open():
        time.sleep(0.6)
        webbrowser.open(url)
    threading.Thread(target=_open, daemon=True).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  Stopped.")


if __name__ == "__main__":
    main()
