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

  /* ---------- Codeblöcke: Kopieren-Knopf (auch für später erzeugte Blöcke aus Dataview), Aktionen auf der "keine Notiz"-Seite ---------- */
  function addCopyButtons(scope) {
    $$('pre', scope || document).forEach(function (pre) {
      var code = pre.querySelector('code');
      if (!code || pre.querySelector('.csc-copy-code')) { return; }
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
  }
  addCopyButtons(document);
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
  /* ---------- Dataview: dataviewjs-Blöcke laufen mit einem dv- und app-Ersatz (nur die aktuelle Notiz, keine Vault-Abfragen) ---------- */
  function escHtml(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  /* Markdown-Teilmenge für dv.paragraph & Co.: Codeblöcke, Überschriften, Absätze, fett/kursiv, Inline-Code, Links */
  function mdInline(s) {
    return escHtml(s)
      .replace(/`([^`]+)`/g, '<code>$1</code>')
      .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
      .replace(/(^|[^*])\*([^*\n]+)\*/g, '$1<em>$2</em>')
      .replace(/\[\[([^\]|]+)\|([^\]]+)\]\]/g, '$2').replace(/\[\[([^\]]+)\]\]/g, '$1')
      .replace(/\[([^\]]+)\]\((https?:[^)\s]+)\)/g, '<a href="$2" class="external-link">$1</a>');
  }
  function mdBlocks(text) {
    var html = '', parts = String(text).split(/```(\w*)[^\n]*\n([\s\S]*?)```/g);
    for (var i = 0; i < parts.length; i += 3) {
      var plain = parts[i];
      plain.split(/\n{2,}/).forEach(function (para) {
        var lines = para.split('\n'), buf = [];
        var flush = function () { if (buf.length) { html += '<p>' + buf.map(mdInline).join('<br>') + '</p>'; buf = []; } };
        lines.forEach(function (line) {
          var h = line.match(/^(#{1,6})\s+(.*)$/);
          if (h) { flush(); html += '<h' + h[1].length + '>' + mdInline(h[2]) + '</h' + h[1].length + '>'; }
          else if (line.trim() !== '') { buf.push(line); }
        });
        flush();
      });
      if (i + 2 < parts.length) {
        html += '<pre><code' + (parts[i + 1] ? ' class="language-' + escHtml(parts[i + 1]) + '"' : '') + '>' + escHtml(parts[i + 2].replace(/\n$/, '')) + '</code></pre>';
      }
    }
    return html;
  }
  function enhance(el) {
    addCopyButtons(el);
    if (window.hljs) { $$('pre code', el).forEach(function (c) { try { window.hljs.highlightElement(c); } catch (e) { } }); }
  }
  var AsyncFunction = Object.getPrototypeOf(async function () { }).constructor;
  function makeApp(meta) {
    var f = meta.file || {};
    var file = { path: f.path || '', name: (f.name || '') + '.md', basename: f.name || '', extension: 'md', parent: { path: f.folder || '', name: (f.folder || '').replace(/^.*\//, '') } };
    var same = function (p) { p = String(p && p.path !== undefined ? p.path : (p || '')).replace(/\\/g, '/').replace(/^\//, ''); return p === file.path || p === file.basename || p === file.name || p.replace(/\.md$/i, '') === file.path.replace(/\.md$/i, ''); };
    var read = function (p) {
      return new Promise(function (resolve, reject) {
        if (same(p)) { resolve(meta.content || ''); }
        else { reject(new Error('app.vault.read("' + (p && p.path !== undefined ? p.path : p) + '") needs Obsidian (only the current note is available here)')); }
      });
    };
    return {
      vault: {
        adapter: { read: read, exists: function (p) { return Promise.resolve(same(p)); } },
        read: read, cachedRead: read,
        getAbstractFileByPath: function (p) { return same(p) ? file : null; },
        getMarkdownFiles: function () { return [file]; },
        getFiles: function () { return [file]; }
      },
      metadataCache: {
        getFileCache: function () { return { frontmatter: meta.frontmatter || {} }; },
        getCache: function () { return { frontmatter: meta.frontmatter || {} }; },
        getFirstLinkpathDest: function (p) { return same(p) ? file : null; }
      },
      workspace: { getActiveFile: function () { return file; } }
    };
  }
  function makeDv(root, page, appShim) {
    function render(el, v, block) {
      if (v === null || v === undefined) { return; }
      if (v instanceof Node) { el.appendChild(v); return; }
      if (Array.isArray(v)) { el.textContent = v.join(', '); return; }
      var s = String(v);
      if (/<\/?[a-z][^>]*>/i.test(s)) { el.innerHTML = s; enhance(el); }   /* values may carry HTML (Obsidian renders it as well) */
      else if (/```|\*\*|`|\[\[|\]\(|^#{1,6}\s|\n/.test(s)) { el.innerHTML = block ? mdBlocks(s) : mdInline(s); enhance(el); }
      else { el.textContent = s; }
    }
    function add(tag, content, cls, block) { var el = document.createElement(tag); if (cls) { el.className = cls; } render(el, content, block); root.appendChild(el); return el; }
    function unsupported(name) { return function () { throw new Error(name + ' needs Obsidian (only the current note is available here)'); }; }
    var samePath = function (p) { p = String(p || '').replace(/\\/g, '/').replace(/\.md$/i, ''); var cur = String(page.file.path || '').replace(/\.md$/i, ''); return !p || p === cur || p === page.file.name; };
    var dv = {
      container: root,
      current: function () { return page; },
      page: function (p) { if (samePath(p)) { return page; } throw new Error('dv.page("' + p + '") needs Obsidian (only the current note is available here)'); },
      pages: unsupported('dv.pages'), pagePaths: unsupported('dv.pagePaths'), query: unsupported('dv.query'), execute: unsupported('dv.execute'),
      header: function (level, text) { return add('h' + Math.min(6, Math.max(1, parseInt(level, 10) || 1)), text, null, false); },
      paragraph: function (text) { return add('div', text, 'csc-dv-p', true); },
      span: function (text) { return add('span', text, null, false); },
      el: function (tag, text, opts) { return add(tag, text, opts && opts.cls, true); },
      list: function (items) { var ul = document.createElement('ul'); (items || []).forEach(function (i) { var li = document.createElement('li'); render(li, i, false); ul.appendChild(li); }); root.appendChild(ul); return ul; },
      table: function (headers, rows) {
        var t = document.createElement('table'), tr = document.createElement('tr');
        (headers || []).forEach(function (h) { var th = document.createElement('th'); render(th, h, false); tr.appendChild(th); });
        t.appendChild(tr);
        (rows || []).forEach(function (r) { var row = document.createElement('tr'); (r || []).forEach(function (c) { var td = document.createElement('td'); render(td, c, false); row.appendChild(td); }); t.appendChild(row); });
        root.appendChild(t); return t;
      },
      date: function (v) { return v; },
      fileLink: function (path, embed, display) { return display || String(path || '').replace(/^.*\//, '').replace(/\.md$/i, ''); },
      array: function (v) { return Array.isArray(v) ? v : (v === undefined || v === null ? [] : [v]); },
      /* dv.view("folder/name", input): loads name.js or name/view.js from the vault (via the vault.csc host) and runs it with the same dv */
      view: function (path, input) {
        var rel = String(path || '').replace(/\\/g, '/').replace(/^\//, '');
        var base = 'https://vault.csc/' + rel.split('/').map(encodeURIComponent).join('/');
        var urls = /\.js$/i.test(rel) ? [base] : [base + '.js', base + '/view.js'];
        var tryNext = function (i) {
          if (i >= urls.length) { return Promise.reject(new Error('dv.view("' + path + '"): script not found in the vault')); }
          return fetch(urls[i]).then(function (r) {
            if (!r.ok) { return tryNext(i + 1); }
            return r.text().then(function (code) { return new AsyncFunction('dv', 'input', 'app', code)(dv, input, appShim); });
          });
        };
        return tryNext(0);
      }
    };
    return dv;
  }
  function processDataview(content) {
    (function () {
      var meta = {};
      try { var metaEl = $('script.csc-note-meta', content); meta = metaEl ? JSON.parse(metaEl.textContent) : {}; } catch (e) { meta = {}; }
      $$('pre > code.language-dataviewjs, pre > code.language-dataview', content).forEach(function (code) {
        var pre = code.parentNode, src = code.textContent, box = document.createElement('div');
        box.className = 'csc-dataview';
        var fail = function (e) {
          box.className = 'csc-dataview failed';
          box.innerHTML = '<div class="csc-dv-note">' + escHtml((txt.dataviewFailed || 'Dataview') + ': ' + (e && e.message ? e.message : e)) + '</div><pre><code>' + escHtml(src) + '</code></pre>';
        };
        pre.parentNode.replaceChild(box, pre);
        if (code.classList.contains('language-dataview')) {
          box.className = 'csc-dataview failed';
          box.innerHTML = '<div class="csc-dv-note">' + escHtml(txt.dataviewOnly || 'Dataview') + '</div><pre><code>' + escHtml(src) + '</code></pre>';
          return;
        }
        var page = Object.assign({}, meta.frontmatter || {});
        page.file = Object.assign({ name: '', path: '', folder: '', link: '', ext: 'md' }, meta.file || {});
        page.file.frontmatter = meta.frontmatter || {};
        var appShim = makeApp(meta);
        try {
          Promise.resolve(new AsyncFunction('dv', 'input', 'app', src)(makeDv(box, page, appShim), undefined, appShim)).catch(fail);
        } catch (e) { fail(e); }
      });
    })();
  }
  function runDataview(scope) { $$('.csc-content', scope || document).forEach(processDataview); }
  runDataview();

  /* ---------- Vorschau beim Überfahren interner Links (Inhalt kommt vom Skript per 'preview'-Nachricht) ---------- */
  /* Popups form a stack: level 0 belongs to a link in the note, level 1 to a link inside that popup, and so on.
     Each popup arms (bar fills in PREVIEW_LOCK_MS) and is then locked: it stays until Escape (topmost), its ×,
     a click outside it or a hover on another link of the same level. Closing a level closes everything above it. */
  var previews = [], previewTimer = null, previewSeq = 0;
  var PREVIEW_LOCK_MS = 1500, PREVIEW_MAX = 6;
  function previewLevelOf(node) {
    if (!node || node.nodeType !== 1) { return 0; }
    for (var i = previews.length - 1; i >= 0; i--) { if (previews[i].el.contains(node)) { return i + 1; } }
    return 0;
  }
  function closePreviews(fromLevel) {
    while (previews.length > fromLevel) {
      var p = previews.pop();
      clearTimeout(p.hideTimer); clearTimeout(p.lockTimer);
      if (p.el.parentNode) { p.el.parentNode.removeChild(p.el); }
    }
    previewSeq++;
  }
  function hidePreview() { closePreviews(0); }
  function newPreview(level, link) {
    var el = document.createElement('div');
    el.className = 'csc-preview'; el.style.zIndex = String(80 + level);
    el.innerHTML = '<div class="csc-preview-bar"></div><div class="csc-preview-title"><span class="csc-preview-name"></span><button type="button" class="csc-preview-close" title="Escape">&times;</button></div><div class="csc-preview-body csc-content"></div>';
    var p = { el: el, link: link, level: level, locked: false, hideTimer: null, lockTimer: null, seq: ++previewSeq };
    el.addEventListener('mouseenter', function () { clearTimeout(p.hideTimer); });
    el.addEventListener('mouseleave', function () { scheduleHide(p); });
    el.addEventListener('mousedown', function () { lockPreview(p); });   /* a click into a popup locks it at once */
    $('.csc-preview-close', el).addEventListener('click', function (e) { e.preventDefault(); e.stopPropagation(); closePreviews(p.level); });
    /* drag by the title bar */
    $('.csc-preview-title', el).addEventListener('mousedown', function (e) {
      if (e.button !== 0 || (e.target && e.target.closest && e.target.closest('.csc-preview-close'))) { return; }
      e.preventDefault();
      var startX = e.clientX, startY = e.clientY, left = el.offsetLeft, top = el.offsetTop;
      var move = function (ev) {
        el.style.left = Math.max(0, Math.min(window.innerWidth - 40, left + ev.clientX - startX)) + 'px';
        el.style.top = Math.max(0, Math.min(window.innerHeight - 30, top + ev.clientY - startY)) + 'px';
      };
      var up = function () { document.removeEventListener('mousemove', move, true); document.removeEventListener('mouseup', up, true); el.classList.remove('dragging'); };
      el.classList.add('dragging');
      document.addEventListener('mousemove', move, true);
      document.addEventListener('mouseup', up, true);
    });
    document.body.appendChild(el);
    previews.push(p);
    return p;
  }
  function lockPreview(p) {
    if (p.locked || previews.indexOf(p) < 0) { return; }
    p.locked = true; p.el.classList.add('locked', 'arming'); clearTimeout(p.hideTimer); clearTimeout(p.lockTimer);
  }
  function scheduleHide(p) {
    if (p.locked) { return; }
    clearTimeout(p.hideTimer);
    p.hideTimer = setTimeout(function () { if (!p.locked && previews.indexOf(p) >= 0) { closePreviews(p.level); } }, 250);
  }
  function placePreview(p) {
    var el = p.el, r = p.link.getBoundingClientRect();
    var w = Math.min(480, window.innerWidth - 24), maxH = Math.min(360, window.innerHeight - 24);
    el.style.width = w + 'px'; el.style.maxHeight = maxH + 'px';
    el.style.top = '0px';
    if (p.level > 0) {
      /* cascade: a nested popup sits a little right of and below its parent, so the parent's title bar stays reachable */
      var parent = previews[p.level - 1].el;
      var left = Math.min(parent.offsetLeft + 28, window.innerWidth - w - 8);
      var top = parent.offsetTop + 40;
      el.style.left = Math.max(8, left) + 'px';
      el.style.maxHeight = Math.max(120, Math.min(maxH, window.innerHeight - top - 8)) + 'px';
      var hh = el.offsetHeight;
      if (top + hh > window.innerHeight - 8) { top = Math.max(8, window.innerHeight - 8 - hh); }
      el.style.top = top + 'px';
      return;
    }
    el.style.left = Math.max(8, Math.min(r.left, window.innerWidth - w - 8)) + 'px';
    var h = el.offsetHeight;
    /* below the link when there is room (the popup shrinks to the space), otherwise above; never on top of the link */
    var spaceBelow = window.innerHeight - r.bottom - 16, spaceAbove = r.top - 16;
    var below = spaceBelow >= Math.min(h, 180) || spaceBelow >= spaceAbove;
    el.style.maxHeight = Math.max(120, Math.min(maxH, below ? spaceBelow : spaceAbove)) + 'px';
    h = el.offsetHeight;
    el.style.top = (below ? r.bottom + 8 : Math.max(8, r.top - 8 - h)) + 'px';
  }
  window.CSC_showPreview = function (data) {
    var p = previews[previews.length - 1];
    if (!data || !p || String(data.id) !== String(p.seq)) { return; }
    var body = $('.csc-preview-body', p.el);
    $('.csc-preview-name', p.el).textContent = data.title || '';
    body.innerHTML = data.html || '';
    enhance(body);
    runDataview(p.el);
    p.el.classList.add('open');
    placePreview(p);
    p.el.scrollTop = 0;
    requestAnimationFrame(function () { if (previews.indexOf(p) >= 0) { p.el.classList.add('arming'); } });
    p.lockTimer = setTimeout(function () { if (previews.indexOf(p) >= 0) { p.locked = true; p.el.classList.add('locked'); clearTimeout(p.hideTimer); } }, PREVIEW_LOCK_MS);
    if (data.anchor) {
      var want = String(data.anchor).trim().toLowerCase();
      var h = $$('h1, h2, h3, h4, h5, h6', body).filter(function (x) { return (x.textContent || '').trim().toLowerCase() === want; })[0];
      if (h) { p.el.scrollTop = Math.max(0, h.offsetTop - 36); }
    }
  };
  document.addEventListener('mouseover', function (ev) {
    if (!cfg.linkPreview) { return; }
    var t = ev.target, a = t && t.closest ? t.closest('a.internal-link[data-rel]') : null;
    if (!a || !a.getAttribute('data-rel') || a.classList.contains('is-unresolved')) { return; }
    var level = previewLevelOf(a);
    if (level >= PREVIEW_MAX) { return; }
    var existing = previews[level];
    if (existing && existing.link === a) { clearTimeout(existing.hideTimer); return; }
    closePreviews(level);   /* a new link on this level replaces the popup of this level and everything above */
    var p = newPreview(level, a);
    clearTimeout(previewTimer);
    previewTimer = setTimeout(function () {
      if (previews[level] !== p) { return; }
      post({ type: 'preview', rel: a.getAttribute('data-rel') || '', anchor: a.getAttribute('data-anchor') || '', id: String(p.seq) });
    }, 450);
  });
  document.addEventListener('mouseout', function (ev) {
    var t = ev.target, a = t && t.closest ? t.closest('a.internal-link') : null;
    if (!a) { return; }
    for (var i = 0; i < previews.length; i++) { if (previews[i].link === a) { clearTimeout(previewTimer); scheduleHide(previews[i]); break; } }
  });
  window.addEventListener('scroll', function (ev) {
    /* scrolling in the note or in a popup closes the unlocked popups above that level */
    var from = previewLevelOf(ev.target);
    while (from < previews.length && previews[from].locked) { from++; }
    closePreviews(from);
  }, true);

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
    closePreviews(previewLevelOf(t));   /* click outside closes all popups, a click inside one closes those above it */
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
    if (ev.key === 'Escape') { if (previews.length) { closePreviews(previews.length - 1); } else { closePanels(); } }
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
