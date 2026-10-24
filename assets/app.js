/* CheatSheetCompanion – Seitenlogik (läuft in WebView2).
   Kommuniziert mit dem PowerShell-Host über window.chrome.webview.postMessage. */
(function () {
  'use strict';
  var cfg = window.CSC_CONFIG || {};
  var txt = cfg.text || {};
  var minHold = (typeof cfg.autoCopyMinHoldMs === 'number') ? cfg.autoCopyMinHoldMs : 200;
  window.CSC_AUTO_COPY = !!cfg.autoCopy;

  function post(msg) {
    try { if (window.chrome && window.chrome.webview) { window.chrome.webview.postMessage(msg); } } catch (e) { }
  }
  function $(sel, root) { return (root || document).querySelector(sel); }
  function $$(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }
  function slugify(t) { return String(t || '').toLowerCase().trim().replace(/[^\p{L}\p{N}\s-]/gu, '').replace(/\s+/g, '-').replace(/-+/g, '-'); }

  /* ---------- Kandidaten-Tabs ---------- */
  var tabs = $('#csc-tabs');
  var buttons = tabs ? $$('button[data-i]', tabs) : [];
  var panes = $$('.csc-pane');

  function activePane() { return $('.csc-pane.active') || panes[0] || null; }

  function activate(i, silent) {
    i = String(i);
    panes.forEach(function (p) { p.classList.toggle('active', p.getAttribute('data-i') === i); });
    buttons.forEach(function (b) { b.classList.toggle('active', b.getAttribute('data-i') === i); });
    buildTOC();
    buildLinks();
    var btn = buttons.filter(function (b) { return b.getAttribute('data-i') === i; })[0];
    if (btn && !silent) {
      post({ type: 'activeNote', index: parseInt(i, 10) || 0, rel: btn.getAttribute('data-rel') || '', name: btn.getAttribute('data-name') || '', score: btn.getAttribute('data-score') || '' });
    }
  }
  if (tabs) {
    tabs.addEventListener('click', function (ev) {
      var b = ev.target && ev.target.closest ? ev.target.closest('button[data-i]') : null;
      if (b) { activate(b.getAttribute('data-i')); }
    });
  }
  function activeIndex() { var b = buttons.filter(function (x) { return x.classList.contains('active'); })[0]; return b ? parseInt(b.getAttribute('data-i'), 10) || 0 : 0; }
  /* Tab-Kürzel: Strg+1..9, Strg+Tab / Strg+Umschalt+Tab (auch vom WPF-Fenster aus aufgerufen) */
  window.CSC_activateTab = function (i) { if (i >= 0 && i < buttons.length) { activate(i); } };
  window.CSC_cycleTab = function (d) { if (buttons.length > 1) { activate(((activeIndex() + d) % buttons.length + buttons.length) % buttons.length); } };

  /* ---------- Inhaltsverzeichnis ---------- */
  var tocPanel = $('#csc-toc-panel'), tocHandle = $('#csc-toc-handle'), tocNav = $('#csc-toc');
  function buildTOC() {
    if (!tocNav || !tocPanel || !tocHandle) { return; }
    tocNav.innerHTML = '';
    var scope = activePane();
    var heads = scope ? $$('h1, h2, h3, h4, h5, h6', scope) : [];
    heads.forEach(function (h, idx) {
      var level = parseInt(h.tagName.substring(1), 10);
      var text = (h.textContent || '').trim();
      if (!h.id) {
        var base = slugify(text) || ('h' + level + '-' + idx), id = base, n = 2;
        while (document.getElementById(id)) { id = base + '-' + (n++); }
        h.id = id;
      }
      var a = document.createElement('a');
      a.href = '#' + h.id; a.textContent = text; a.className = 'csc-lv' + level;
      a.addEventListener('click', function (e) { e.preventDefault(); h.scrollIntoView({ behavior: 'smooth', block: 'start' }); });
      tocNav.appendChild(a);
    });
    var show = heads.length > 1;
    tocPanel.style.display = show ? '' : 'none';
    tocHandle.style.display = show ? '' : 'none';
  }
  if (tocHandle && tocPanel) { tocHandle.addEventListener('click', function () { closePanels(tocPanel); tocPanel.classList.toggle('open'); }); }

  /* ---------- Interne Links, externe Links, Callouts ---------- */
  document.addEventListener('click', function (ev) {
    var t = ev.target;
    if (!t || !t.closest) { return; }
    var link = t.closest('a.internal-link');
    if (link) {
      ev.preventDefault();
      if (link.classList.contains('is-unresolved')) { return; }
      post({ type: 'openNote', rel: link.getAttribute('data-rel') || '', note: link.getAttribute('data-note') || '', anchor: link.getAttribute('data-anchor') || '' });
      return;
    }
    var title = t.closest('.callout.is-collapsible > .callout-title');
    if (title) { title.parentElement.classList.toggle('is-collapsed'); return; }
    var tag = t.closest('.tag');
    if (tag) { post({ type: 'search', query: (tag.textContent || '').trim() }); }
  });
  $$('a[href]').forEach(function (a) {
    var href = a.getAttribute('href') || '';
    if (/^https?:\/\//i.test(href) && !a.classList.contains('internal-link') && !a.querySelector('img')) { a.classList.add('external-link'); }
  });

  /* ---------- Codeblöcke: Kopieren-Knopf, Aktionen auf der "keine Notiz"-Seite ---------- */
  $$('pre').forEach(function (pre) {
    var code = pre.querySelector('code');
    if (!code) { return; }
    var btn = document.createElement('button');
    btn.type = 'button'; btn.className = 'csc-copy-code'; btn.textContent = txt.copy || 'Kopieren'; btn.title = txt.copyCode || '';
    btn.addEventListener('mousedown', function (e) { e.preventDefault(); e.stopPropagation(); });
    btn.addEventListener('click', function (e) {
      e.preventDefault(); e.stopPropagation();
      post({ type: 'copyCode', text: code.innerText.replace(/\n$/, '') });
      btn.textContent = txt.copied || 'Kopiert'; setTimeout(function () { btn.textContent = txt.copy || 'Kopieren'; }, 1200);
    });
    pre.appendChild(btn);
  });
  $$('button[data-action]').forEach(function (b) {
    b.addEventListener('click', function (e) { e.preventDefault(); post({ type: b.getAttribute('data-action') }); });
  });
  /* ---------- Zuletzt kopiert: ausfahrbare Leiste, ein Klick kopiert erneut ---------- */
  var copyPanel = $('#csc-copy-panel'), copyHandle = $('#csc-copy-handle'), copyList = $('#csc-copy-list');
  function renderCopyHistory(list) {
    if (!copyPanel || !copyHandle || !copyList) { return; }
    list = Array.isArray(list) ? list : [];
    copyList.innerHTML = '';
    list.forEach(function (item, idx) {
      var b = document.createElement('button');
      b.type = 'button'; b.className = 'csc-recopy'; b.setAttribute('data-copy', String(idx));
      b.title = (item.text || '') + (txt.copyAgain ? '\n\n' + txt.copyAgain : '');
      var code = document.createElement('code'); code.textContent = item.label || ''; b.appendChild(code);
      if (item.note) { var m = document.createElement('span'); m.className = 'csc-meta'; m.textContent = item.note; b.appendChild(m); }
      b.addEventListener('mousedown', function (e) { e.preventDefault(); });
      b.addEventListener('click', function (e) {
        e.preventDefault();
        post({ type: 'recopy', index: idx });
        copyPanel.classList.remove('open');   /* ein Klick: kopiert und schließt die Leiste */
      });
      copyList.appendChild(b);
    });
    var show = list.length > 0;
    copyHandle.style.display = show ? '' : 'none';
    if (!show) { copyPanel.classList.remove('open'); }
  }
  window.CSC_setCopyHistory = renderCopyHistory;
  var copyClear = $('#csc-copy-clear');
  if (copyClear) {
    copyClear.addEventListener('mousedown', function (e) { e.preventDefault(); e.stopPropagation(); });
    copyClear.addEventListener('click', function (e) { e.preventDefault(); post({ type: 'clearCopyHistory' }); });
  }
  if (copyHandle && copyPanel) {
    copyHandle.addEventListener('click', function () { closePanels(copyPanel); copyPanel.classList.toggle('open'); });
  }
  renderCopyHistory(cfg.copyHistory || []);
  /* ---------- Verlinkt von (Backlinks der aktiven Notiz, Leiste wie das TOC) ---------- */
  var linksPanel = $('#csc-links-panel'), linksHandle = $('#csc-links-handle'), linksList = $('#csc-links-list'), linksCount = $('#csc-links-count');
  function buildLinks() {
    if (!linksPanel || !linksHandle || !linksList) { return; }
    var scope = activePane();
    var data = scope ? $('.csc-backlinks-data', scope) : null;
    linksList.innerHTML = data ? data.innerHTML : '';
    if (linksCount) { linksCount.textContent = data ? (data.getAttribute('data-count') || '') : ''; }
    var show = !!data;
    linksHandle.style.display = show ? '' : 'none';
    if (!show) { linksPanel.classList.remove('open'); }
  }
  if (linksHandle && linksPanel) {
    linksHandle.addEventListener('click', function () { closePanels(linksPanel); linksPanel.classList.toggle('open'); });
    linksList.addEventListener('click', function (ev) { var a = ev.target && ev.target.closest ? ev.target.closest('a') : null; if (a) { linksPanel.classList.remove('open'); } });
  }
  buildLinks();
  /* Leisten schließen, sobald man woanders hinklickt oder das Fenster verlässt (except = die Leiste, die gerade aufgeht) */
  function closePanels(except) {
    [copyPanel, tocPanel, linksPanel].forEach(function (p) { if (p && p !== except) { p.classList.remove('open'); } });
  }
  window.addEventListener('blur', function () { closePanels(); });

  /* ---------- Zum Anker scrollen ---------- */
  function scrollToAnchor(anchor) {
    if (!anchor) { return; }
    var scope = activePane() || document;
    var el = document.getElementById(anchor) || document.getElementById(slugify(anchor));
    if (!el) {
      var want = String(anchor).trim().toLowerCase();
      el = $$('h1, h2, h3, h4, h5, h6', scope).filter(function (h) { return (h.textContent || '').trim().toLowerCase() === want; })[0] || null;
    }
    if (el) { el.scrollIntoView({ block: 'start' }); }
  }

  /* ---------- Auswahl kopieren (AutoCopy) ---------- */
  var downAt = 0;
  function getSel() { try { return String(window.getSelection ? window.getSelection() : '').trim(); } catch (e) { return ''; } }
  document.addEventListener('mousedown', function (ev) {
    var t = ev.target;
    downAt = performance.now();
    var inside = t && t.closest ? t.closest('#csc-copy-panel, #csc-copy-handle, #csc-toc-panel, #csc-toc-handle, #csc-links-panel, #csc-links-handle') : null;
    if (!inside) { closePanels(); }
  }, true);
  document.addEventListener('mouseup', function (ev) {
    if (ev.button === 3) { post({ type: 'back' }); return; }
    if (ev.button === 4) { post({ type: 'forward' }); return; }
    if (ev.button !== 0) { return; }
    var held = downAt > 0 ? (performance.now() - downAt) : 0;
    var s = getSel();
    if (!s) { return; }
    post({ type: 'sel', text: s.substring(0, 4000) });
    if (window.CSC_AUTO_COPY && held >= minHold) { post({ type: 'selection', text: s, held: Math.round(held) }); }
  }, true);
  document.addEventListener('keydown', function (ev) {
    if (ev.key === 'Escape') { closePanels(); }
    if (ev.altKey && ev.key === 'ArrowLeft') { post({ type: 'back' }); }
    if (ev.altKey && ev.key === 'ArrowRight') { post({ type: 'forward' }); }
    if (ev.ctrlKey && !ev.altKey && (ev.key === 'k' || ev.key === 'K')) { ev.preventDefault(); post({ type: 'focusSearch' }); }
    /* Strg+Umschalt+C: Markierung (oder ganze Notiz) formatiert kopieren, Strg+C bleibt reiner Text */
    if (ev.ctrlKey && ev.shiftKey && !ev.altKey && (ev.key === 'c' || ev.key === 'C')) { ev.preventDefault(); post({ type: 'copyFormatted' }); }
    if (ev.ctrlKey && !ev.altKey && /^[1-9]$/.test(ev.key)) { ev.preventDefault(); window.CSC_activateTab(parseInt(ev.key, 10) - 1); }
    if (ev.ctrlKey && !ev.altKey && ev.key === 'Tab') { ev.preventDefault(); window.CSC_cycleTab(ev.shiftKey ? -1 : 1); }
  });
  /* Doppelklick markiert ein Wort; den Editor öffnet er nur mit Strg (Standard), immer oder nie (EditOnDoubleClick). */
  document.addEventListener('dblclick', function (ev) {
    var mode = cfg.dblClickEdit || 'ctrl';
    if (mode === 'off') { return; }
    if (mode === 'ctrl' && !ev.ctrlKey) { return; }
    var t = ev.target;
    if (t && t.closest && (t.closest('a, button, #csc-header, #csc-toc-panel, input'))) { return; }
    if (mode === 'ctrl') { ev.preventDefault(); try { window.getSelection().removeAllRanges(); } catch (e) { } }
    post({ type: 'dblClick' });
  }, true);

  /* ---------- Start ---------- */
  activate(cfg.activeIndex || 0, true);
  if (cfg.anchor) { setTimeout(function () { scrollToAnchor(cfg.anchor); }, 30); }
  post({ type: 'ready' });
})();
