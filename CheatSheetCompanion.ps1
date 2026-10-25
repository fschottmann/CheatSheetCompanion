#Requires -Version 7

<#PSScriptInfo

    .VERSION 3.4.0

    .GUID 1a0524f5-edba-4780-b562-6097e428b673

    .AUTHOR Frank Schottmann

    .COPYRIGHT @2024-2026 Frank Schottmann

    .RELEASENOTES
    Version 3.4.0
    - Search: aliases are searched like names, substring hits always rank above fuzzy hits, content hits in between;
      the whole pass runs in C# and typing is debounced (large vaults).
    - dataviewjs blocks run with a small dv and app replacement (current note only: dv.current, dv.page, header,
      paragraph with Markdown and HTML, list, table, el, dv.view loading scripts from the vault, app.vault.read of
      the current note, await); inline fields ("key:: value", also inside %% comments) are available like front
      matter; vault-wide queries and DQL blocks show a hint instead of raw code; works in preview popups.
    - Link preview: hovering an internal link shows the target note in a popup (LinkPreview = false turns it off);
      after a short progress bar it is locked (clickable, scrollable) until Escape, its close button or a click
      elsewhere; links inside a popup open further popups (cascaded stack of up to six levels, draggable).
    - Ranking: ProcessWeight (default 0.9) keeps the note named in the window title ahead of notes that merely
      contain the process name; equal scores prefer the longer matched name or alias.
    - IgnoreFolders accepts relative paths ("archive\old notes") and drops the notes of those folders only; images
      there stay resolvable for embeds. ini\config.local.ini overrides config.ini for personal settings (not in git).
    - Start-up with large vaults: vault walk, parallel note reading, link extraction and backlinks run in C#, and an
      unchanged vault converts nothing (Output\<vault>\index.cache remembers the file set; before, every start
      re-converted all notes). Phase timings in the debug log.
    - Fenced code detection tolerates CR before LF; the page can write diagnostics to the log.

    Version 3.3.0
    - UI texts come from lang\<code>.ini (Language = auto|de|en|...), German built in as fallback.
    - Pin reworked: pinned notes (several) stay as marked tabs while matching goes on, the best match is shown
      (MaxCandidates is the total, at least one slot for matches).
    - Copy history: texts copied through the companion are kept (CopyHistory) in a slide-out panel ("KOPIERT"
      handle like the TOC), one click copies again.
    - Tab shortcuts Ctrl+1..9 and Ctrl+Tab; forward navigation (button, Alt+Right, mouse button 5).
    - Backlinks: "Linked from" slide-out panel (handle like the TOC) with the linking notes and the mention
      lines (ShowBacklinks), like Obsidian's linked mentions.
    - Window opacity: WindowOpacity, and InactiveOpacity dims the window while another program is active
      (0 = off); full opacity again under the mouse, when collapsed and while the search is open.
    - Memory: the C# helpers ship prebuilt as dll\CheatSheetCompanionHelper.dll (tools\Build-Helper.ps1; nothing is
      compiled or written at run time, in-memory compile only as fallback), the WebView2 renderer is suspended while
      the window is collapsed, note texts are kept once (SearchContent = false drops them), bounded history.
    - Copy as…: selection or whole note formatted (HTML + text for Teams, Outlook; Ctrl+Shift+C) or as text.
    - Append clipboard to note: context menu adds the copied text as is to the end of the shown note.
    - Copy history panel has a Clear button; the mini "Copy" overlay after selecting text was removed.
    - Fixed: collapsing a maximized window had no effect (Height is ignored while maximized).
    - Bundled demo vault (vault\) with the help note, English ini comments, relative vault paths.
    - WebView2 SDK updated to 1.0.4314.35 (dll\, .NET Core build), license file added.

    Version 3.2.0
    - Double click no longer opens the editor by accident (EditOnDoubleClick = ctrl|always|off).
    - Clipboard mode: deferred retry when the clipboard is locked, same text can be looked up again,
      clipboard note keeps priority inside the window it was copied from; optional one shot hotkey (ClipboardLookup).
    - No match keeps the current note (NoMatchBehavior = keep|suggest).
    - Recently shown notes on the start page; aliases can be added from the context menu; [Rules] section
      (process name / title pattern -> note); folder selection in the new note dialog; faster vault refresh.
    - Context menu items were never shown (wrong delegate type for CustomItemSelected) - fixed.
    - Pin (persisted), notes excluded from matching (MatchIgnoreFolders / front matter "csc: ignore"),
      match reason in status line and tab tooltip (ShowMatchReason), collapsed strip shows the current note.

    Version 3.1.0
    - Fuzzy search (quick switcher) replaces the drop down, searches names and note contents.
    - Themes (themes\*.css) with automatic dark mode, selectable from the context menu; dark WPF chrome.
    - Multiple vaults ([Vaults] section), switchable from the context menu.
    - New note from window title or clipboard text (dialog).
    - Tray icon with the three toggles, zoom level persisted, optional front matter box, tag click = search.

    Version 3.0.0
    - Rewrite: no background runspaces any more (event driven: DispatcherTimer, RegisterHotKey, clipboard listener, FileSystemWatcher).
    - Matching moved to C# (Dice n-gram similarity), culture independent, Top-N with deterministic tie-break.
    - Obsidian rendering: wikilinks, embeds (![[...]]), callouts, tags, highlights, soft line breaks, YAML front matter.
    - Images and relative links resolved against the vault through a WebView2 virtual host.
    - Clipboard mode: react to copied text (e.g. inside RDP/SSH sessions).
#>

<#
    .SYNOPSIS
    Cheat Sheet Companion - context sensitive notes (Markdown / Obsidian vault) next to the active window.

    .DESCRIPTION
    Shows the notes that best match the active window title, the process name or the clipboard content.
    Notes are rendered Obsidian-like inside a WebView2 control. See README.md for configuration.

    .PARAMETER OpenNote
    Shows this note (name or search text) right after start.

    .PARAMETER AutoCloseSeconds
    Test helper: closes the window automatically after the given number of seconds.

    .PARAMETER ScreenshotPath
    Test helper: saves a PNG two seconds before auto close (page content, or the whole window when the name ends with .window.png).

    .PARAMETER SearchQuery
    Test helper: types this query into the search box once the WebView is ready (opens the result popup).
#>
[CmdletBinding()]
param(
    [string]$OpenNote = '',
    [int]$AutoCloseSeconds = 0,
    [string]$ScreenshotPath = '',
    [string]$SearchQuery = ''
)

Set-StrictMode -Off
$ErrorActionPreference = 'Continue'

#region ---------- Configuration helpers ----------

function Get-IniContent {
    <#
        .SYNOPSIS
        Reads an INI file into an ordered dictionary of ordered hashtables. Supports ';' and '#' comments.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $ini = [ordered]@{}
    if (-not (Test-Path -LiteralPath $Path)) { return $ini }
    $section = $null
    foreach ($rawLine in [System.IO.File]::ReadAllLines($Path)) {
        $line = $rawLine.Trim()
        if ($line.Length -eq 0 -or $line.StartsWith(';') -or $line.StartsWith('#')) { continue }
        if ($line -match '^\[(.+)\]$') {
            $section = $matches[1].Trim()
            if (-not $ini.Contains($section)) { $ini[$section] = [ordered]@{} }
            continue
        }
        if ($null -ne $section -and $line -match '^([^=]+?)\s*=\s*(.*)$') {
            $value = $matches[2].Trim()
            $value = ($value -replace '\s+[;#].*$', '').Trim()   # trailing comment
            $ini[$section][$matches[1].Trim()] = $value
        }
    }
    return $ini
}

function Get-Setting {
    <# Returns a typed configuration value with a default (type is derived from the default). #>
    param(
        [Parameter(Mandatory)][string]$Section,
        [Parameter(Mandatory)][string]$Key,
        $Default = ''
    )
    try {
        $sec = $script:Config[$Section]
        if ($null -eq $sec -or -not $sec.Contains($Key)) { return $Default }
        $raw = [string]$sec[$Key]
        if ([string]::IsNullOrWhiteSpace($raw)) { return $Default }
        switch ($Default) {
            { $_ -is [bool] }   { return ($raw -match '^(1|true|yes|on|ja)$') }
            { $_ -is [int] }    { return [int]::Parse($raw, [cultureinfo]::InvariantCulture) }
            { $_ -is [double] } { return [double]::Parse($raw.Replace(',', '.'), [cultureinfo]::InvariantCulture) }
            default             { return $raw }
        }
    }
    catch { return $Default }
}

#endregion

#region ---------- Texts (UI language) ----------

# German is built in; lang\<code>.ini ([Text] section, same keys, "\n" = line break) overrides it. {0}, {1} = placeholders.
$script:TextDefault = @{
    'app.newNoteDefault'     = 'Neue Notiz'
    'app.ready'              = 'Bereit'
    'status.vault'           = 'Vault: {0} ({1} Notizen)'
    'status.vaultUpdated'    = 'Vault aktualisiert: {0} Notizen.'
    'status.start'           = '{0}: {1} Notizen. Fenster: {2}, Zwischenablage: {3}.'
    'status.on'              = 'EIN'
    'status.off'             = 'AUS'
    'status.theme'           = 'Theme: {0}'
    'status.note'            = 'Notiz: {0}'
    'status.match'           = '{0}{1}  ->  {2} ({3}{4})'
    'status.noMatchSuggest'  = '{0}{1}  ->  kein Treffer (+ legt eine Notiz an)'
    'status.noMatchKeep'     = '{0}{1}  ->  kein Treffer, Ansicht bleibt (+ legt eine Notiz an)'
    'status.prefixWindow'    = 'Fenster: '
    'status.prefixClipboard' = 'Zwischenablage: '
    'status.clipboardEmpty'  = 'Zwischenablage enthält keinen Text.'
    'status.aliasExists'     = 'Alias existiert bereits: {0}'
    'status.aliasAdded'      = "Alias '{0}' -> {1}"
    'status.noNote'          = 'Keine Notiz ausgewählt.'
    'status.edit'            = 'Bearbeiten: {0}'
    'status.openFailed'      = 'Öffnen fehlgeschlagen (siehe Log).'
    'status.folderOutside'   = 'Ordner liegt außerhalb des Vaults.'
    'status.notFound'        = 'Keine Notiz gefunden für "{0}".'
    'status.followWindow'    = 'Fenster folgen: {0}'
    'status.followClipboard' = 'Zwischenablage folgen: {0}'
    'status.autoCopy'        = 'Auto Copy: {0}'
    'status.pinned'          = 'Angeheftet: {0} (bleibt als Tab stehen, Treffer werden weiter angezeigt).'
    'status.pinnedStart'     = 'Angeheftet: {0}'
    'status.unpinned'        = 'Anheften aufgehoben: {0}'
    'status.selectionCopied' = 'Auswahl kopiert.'
    'status.copied'          = 'Kopiert ({0} Zeichen).'
    'status.codeCopied'      = 'Codeblock kopiert.'
    'status.recopied'        = 'Erneut kopiert: {0}'
    'status.copiedFormatted' = 'Formatiert kopiert ({0} Zeichen, Text-Variante liegt mit in der Zwischenablage)'
    'status.noteCopied'      = 'Notiz als Text kopiert ({0} Zeichen)'
    'status.appended'        = 'An "{0}" angehängt ({1} Zeichen)'
    'status.copyCleared'     = 'Kopierverlauf geleert.'
    'status.webviewFailed'   = 'WebView2 konnte nicht initialisiert werden (Runtime installiert?).'
    'origin.title'           = 'Titel'
    'origin.process'         = 'Prozess'
    'origin.clipboard'       = 'Zwischenablage'
    'origin.rule'            = 'Regel'
    'origin.manual'          = 'Suche'
    'origin.link'            = 'Link'
    'origin.pinned'          = 'Angeheftet'
    'reason.contains'        = 'Name im Text enthalten'
    'reason.contained'       = 'Text im Namen enthalten'
    'reason.whole'           = 'ähnlich zum ganzen Text'
    'reason.token'           = "ähnlich zu Wort '{0}'"
    'reason.alias'           = "Alias '{0}': {1}"
    'reason.rule'            = "Regel '{0}'"
    'reason.pinned'          = 'Angeheftete Notiz'
    'theme.auto'             = 'Automatisch (Windows)'
    'theme.light'            = 'Hell'
    'theme.dark'             = 'Dunkel'
    'page.missingTarget'     = 'Ziel nicht gefunden'
    'page.missingNote'       = 'Notiz existiert nicht'
    'page.properties'        = 'Eigenschaften'
    'page.toc'               = 'Inhaltsverzeichnis'
    'page.tocTitle'          = 'Inhalt'
    'page.copy'              = 'Kopieren'
    'page.copied'            = 'Kopiert'
    'page.copyCode'          = 'Codeblock in die Zwischenablage kopieren'
    'page.recent'            = 'Zuletzt angezeigt:'
    'page.copyHistory'       = 'Zuletzt kopiert'
    'page.copyHandle'        = 'KOPIERT'
    'page.copyAgain'         = 'Klick kopiert den Text erneut in die Zwischenablage'
    'page.copyClear'         = 'Leeren'
    'page.copyClearTip'      = 'Kopierverlauf löschen (liegt in %LocalAppData%\CheatSheetCompanion\state.json)'
    'page.welcome'           = 'Bereit. Notizen werden passend zum aktiven Fenster{0} angezeigt.<br>Suche mit Strg+K.<br>Vault: {1} ({2} Notizen)'
    'page.welcomeClipboard'  = ' oder zur Zwischenablage'
    'page.nearest'           = 'Ähnlichste Notizen:'
    'page.noMatch'           = 'Noch keine Notiz dazu'
    'page.createNote'        = 'Notiz anlegen'
    'page.pinnedTab'         = 'Angeheftet: bleibt als Tab stehen, bis der Pin gelöst wird'
    'page.backlinks'         = 'Verlinkt von'
    'page.backlinksTip'      = 'Notizen, die auf diese Notiz verlinken (Klick öffnet die Notiz)'
    'page.linksHandle'       = 'VERLINKT'
    'page.dataviewOnly'      = 'Dataview-Abfrage: nur in Obsidian verfügbar'
    'page.dataviewFailed'    = 'Dataview-Block außerhalb von Obsidian nicht ausführbar'
    'strip.pinned'           = '[angeheftet]'
    'menu.copy'              = 'In Zwischenablage kopieren'
    'menu.copyAs'            = 'Kopieren als…'
    'menu.copySelFormatted'  = 'Markierung formatiert (Teams, Outlook)'
    'menu.copyNoteText'      = 'Ganze Notiz als Text'
    'menu.copyNoteFormatted' = 'Ganze Notiz formatiert (Teams, Outlook)'
    'menu.append'            = 'Zwischenablage an Notiz anhängen'
    'menu.edit'              = 'Notiz bearbeiten'
    'menu.alias'             = 'Alias hinzufügen (Notiz zuordnen)'
    'menu.aliasWindow'       = 'Fenstertitel'
    'menu.aliasProcess'      = 'Prozessname'
    'menu.aliasClipboard'    = 'Zwischenablage'
    'menu.aliasSelection'    = 'Markierung'
    'menu.aliasCustom'       = 'Eigener Text…'
    'menu.clipLookup'        = 'Zur Zwischenablage nachschlagen'
    'menu.pin'               = 'Notiz anheften'
    'menu.theme'             = 'Theme'
    'menu.vault'             = 'Vault'
    'tray.show'              = 'Anzeigen'
    'tray.toggle'            = 'Ein-/Ausklappen'
    'tray.search'            = 'Suchen (Strg+K)'
    'tray.pin'               = 'Notiz anheften'
    'tray.followWindow'      = 'Fenster folgen'
    'tray.followClipboard'   = 'Zwischenablage folgen'
    'tray.autoCopy'          = 'Auto Copy'
    'tray.exit'              = 'Beenden'
    'ui.stripTip'            = 'Klick: ein-/ausklappen. Rechtsklick: beenden.'
    'ui.back'                = 'Zurück (Alt+Links)'
    'ui.forward'             = 'Vorwärts (Alt+Rechts)'
    'ui.edit'                = 'Angezeigte Notiz bearbeiten (Obsidian / Standard-Editor).'
    'ui.editDblCtrl'         = ' Auch per Strg+Doppelklick im Text.'
    'ui.editDblAlways'       = ' Auch per Doppelklick im Text.'
    'ui.add'                 = 'Neue Notiz für das zuletzt aktive Fenster bzw. den kopierten Text anlegen.'
    'ui.pin'                 = 'Pin'
    'ui.pinTip'              = 'Angezeigte Notiz anheften: bleibt als Tab stehen, auch wenn Fenster oder Zwischenablage wechseln. Mehrere Notizen möglich, bleibt über einen Neustart erhalten. Tabs: Strg+1 bis 9, Strg+Tab.'
    'ui.window'              = 'Fenster'
    'ui.windowTip'           = 'Notizen passend zum aktiven Fenster anzeigen (Titel + Prozess).'
    'ui.clip'                = 'Clip'
    'ui.clipTip'             = 'Notizen passend zum kopierten Text anzeigen (z.B. aus RDP/SSH-Sitzungen).'
    'ui.autoCopy'            = 'AC'
    'ui.autoCopyTip'         = 'Auto Copy: markierten Text beim Loslassen der Maus kopieren (Mindest-Haltezeit konfigurierbar).'
    'ui.searchTip'           = 'Notiz suchen: Name oder Inhalt (Strg+K). Pfeiltasten + Enter öffnen.'
    'ui.searchHint'          = 'Suchen…  (Strg+K)'
    'ui.contentPrefix'       = 'Inhalt: '
    'ui.aliasPrefix'         = 'Alias: '
    'dlg.cancel'             = 'Abbrechen'
    'dlg.aliasTitle'         = 'Alias hinzufügen'
    'dlg.aliasFor'           = 'Alias für "{0}":'
    'dlg.aliasInfo'          = 'Der Alias wird als aliases-Eintrag ins Frontmatter geschrieben. Enthält ein Fenstertitel, Prozessname oder kopierter Text den Alias, zählt die Notiz als Treffer (1.00).'
    'dlg.aliasOk'            = 'Hinzufügen'
    'dlg.newTitle'           = 'Neue Notiz'
    'dlg.newName'            = 'Name der Notiz'
    'dlg.newFrom'            = 'Vorschlag aus'
    'dlg.newInsert'          = 'Diesen Text als Inhalt einfügen'
    'dlg.newFolder'          = 'Ordner im Vault (leer = Wurzel)'
    'dlg.newOk'              = 'Anlegen und öffnen'
    'dlg.srcWindow'          = 'Fenstertitel'
    'dlg.srcClipboard'       = 'Zwischenablage'
    'dlg.srcSelection'       = 'Markierung im Companion'
    'dlg.empty'              = '(leer)'
    'dlg.vault'              = 'Vault: {0}'
    'dlg.template'           = 'Vorlage: {0}'
}
$script:Text = @{}

function Get-Text {
    <# UI text by key (language file, then built-in German, then the key itself); -Format applies the placeholders. #>
    param([Parameter(Mandatory)][string]$Key, [object[]]$Format)
    $s = $script:Text[$Key]
    if ($null -eq $s) { $s = $script:TextDefault[$Key] }
    if ($null -eq $s) { return $Key }
    if ($null -ne $Format -and $Format.Count -gt 0) { try { return ([string]$s -f $Format) } catch { return [string]$s } }
    return [string]$s
}

function Initialize-Language {
    <# Language = auto (Windows display language) | de | en | any code with a lang\<code>.ini. Unknown -> en.ini, missing -> built-in German. #>
    param([AllowEmptyString()][string]$Requested)
    $code = ([string]$Requested).Trim().ToLowerInvariant()
    if (-not $code -or $code -eq 'auto') { $code = [cultureinfo]::CurrentUICulture.TwoLetterISOLanguageName.ToLowerInvariant() }
    $dir = Join-Path $script:App.Root 'lang'
    $file = Join-Path $dir ($code + '.ini')
    if ($code -ne 'de' -and -not (Test-Path -LiteralPath $file)) { $code = 'en'; $file = Join-Path $dir 'en.ini' }
    $script:Text = @{}
    if (Test-Path -LiteralPath $file) {
        $ini = Get-IniContent -Path $file
        if ($ini.Contains('Text')) {
            foreach ($k in $ini['Text'].Keys) { $script:Text[[string]$k] = ([string]$ini['Text'][$k]).Replace('\n', "`n") }
        }
    }
    elseif ($code -ne 'de') { $code = 'de' }
    $script:App.Language = $code
}

#endregion

#region ---------- Logging ----------

function Write-CSCLog {
    <# Lean logger: errors always go to the log file, Info only with Debug=true or -Verbose. #>
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info', 'Warn', 'Error')][string]$Level = 'Info'
    )
    try {
        $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Level, $Message
        Write-Verbose $line
        if ($Level -eq 'Error' -or $script:App.Debug) {
            Add-Content -LiteralPath $script:App.LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
        }
    }
    catch { }
}

function Set-Status {
    param([Parameter(Mandatory)][string]$Text)
    try { $script:UI.Status.Text = $Text } catch { }
}

#endregion

#region ---------- Native helpers (C#) ----------

Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, System.Windows.Forms, System.Drawing

$helperSource = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace CheatSheetCompanionHelper {

    public static class Native {
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")] static extern int GetWindowTextLength(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll", SetLastError = true)] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll", SetLastError = true)] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll", SetLastError = true)] public static extern bool AddClipboardFormatListener(IntPtr hwnd);
        [DllImport("user32.dll", SetLastError = true)] public static extern bool RemoveClipboardFormatListener(IntPtr hwnd);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] public static extern bool SetDllDirectory(string lpPathName);
        [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

        [DllImport("kernel32.dll")] static extern IntPtr GetCurrentProcess();
        [DllImport("kernel32.dll", SetLastError = true)] static extern bool SetProcessWorkingSetSize(IntPtr hProcess, IntPtr dwMin, IntPtr dwMax);
        // Gives unused pages back to the OS (collapsed window); private memory is unaffected, only the working set shrinks.
        public static bool TrimWorkingSet() { return SetProcessWorkingSetSize(GetCurrentProcess(), (IntPtr)(-1), (IntPtr)(-1)); }
        [DllImport("user32.dll", SetLastError = true)] static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll", SetLastError = true)] static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        [DllImport("user32.dll", SetLastError = true)] static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

        // Whole window alpha (WS_EX_LAYERED + LWA_ALPHA): unlike Window.Opacity this includes the WebView2 child window.
        public static bool SetWindowAlpha(IntPtr hwnd, double opacity) {
            const int GWL_EXSTYLE = -20, WS_EX_LAYERED = 0x80000;
            int ex = GetWindowLong(hwnd, GWL_EXSTYLE);
            if ((ex & WS_EX_LAYERED) == 0) SetWindowLong(hwnd, GWL_EXSTYLE, ex | WS_EX_LAYERED);
            byte a = (byte)Math.Max(0, Math.Min(255, (int)Math.Round(opacity * 255.0)));
            return SetLayeredWindowAttributes(hwnd, 0, a, 2 /* LWA_ALPHA */);
        }

        // Dark title bar (DWMWA_USE_IMMERSIVE_DARK_MODE = 20, older builds 19)
        public static void SetDarkTitleBar(IntPtr hwnd, bool dark) {
            int v = dark ? 1 : 0;
            try { if (DwmSetWindowAttribute(hwnd, 20, ref v, 4) != 0) DwmSetWindowAttribute(hwnd, 19, ref v, 4); } catch { }
        }

        public static string GetWindowTitle(IntPtr hWnd) {
            int len = GetWindowTextLength(hWnd);
            if (len <= 0) return string.Empty;
            var sb = new StringBuilder(len + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }
        public static int GetWindowProcessId(IntPtr hWnd) {
            uint pid; GetWindowThreadProcessId(hWnd, out pid); return (int)pid;
        }
    }

    // Receives WM_HOTKEY / WM_CLIPBOARDUPDATE on the WPF UI thread (HwndSource hook).
    public class MessageSink {
        public const int WM_HOTKEY = 0x0312;
        public const int WM_CLIPBOARDUPDATE = 0x031D;
        public const int WM_STYLECHANGING = 0x007C;
        public event Action<int> HotKeyPressed;
        public event Action ClipboardChanged;
        // WPF's HwndTarget strips WS_EX_LAYERED in WM_STYLECHANGING (it only allows it with AllowsTransparency).
        // With KeepLayered the message is marked handled before WPF sees it, so SetWindowAlpha works.
        public bool KeepLayered;
        public IntPtr Hook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) {
            try {
                if (msg == WM_HOTKEY) { var h = HotKeyPressed; if (h != null) h(wParam.ToInt32()); handled = true; }
                else if (msg == WM_CLIPBOARDUPDATE) { var c = ClipboardChanged; if (c != null) c(); handled = true; }
                else if (msg == WM_STYLECHANGING && KeepLayered && unchecked((int)wParam.ToInt64()) == -20 /* GWL_EXSTYLE */) { handled = true; }
            } catch { }
            return IntPtr.Zero;
        }
    }

    // Reason codes: contains (name inside text), contained (text inside name), whole (n-gram similarity), token:<word>
    public struct Match { public int Index; public double Score; public string Reason; public Match(int i, double s, string r) { Index = i; Score = s; Reason = r; } }

    // Dice coefficient over character n-grams, culture independent.
    public static class Similarity {
        public static string Normalize(string s) {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            var sb = new StringBuilder(s.Length);
            foreach (char c in s.ToLowerInvariant()) if (char.IsLetterOrDigit(c)) sb.Append(c);
            return sb.ToString();
        }
        static Dictionary<string, int> Grams(string norm, int n) {
            var d = new Dictionary<string, int>(StringComparer.Ordinal);
            if (norm.Length == 0) return d;
            int k = Math.Min(n, norm.Length);
            for (int i = 0; i + k <= norm.Length; i++) { string g = norm.Substring(i, k); int v; d.TryGetValue(g, out v); d[g] = v + 1; }
            return d;
        }
        public static double Dice(string a, string b, int n) {
            var ga = Grams(a, n); var gb = Grams(b, n);
            if (ga.Count == 0 || gb.Count == 0) return 0.0;
            int inter = 0, ca = 0, cb = 0;
            foreach (var kv in ga) { ca += kv.Value; int v; if (gb.TryGetValue(kv.Key, out v)) inter += Math.Min(kv.Value, v); }
            foreach (var kv in gb) cb += kv.Value;
            return 2.0 * inter / (ca + cb);
        }
        public static double Score(string name, string target, int n) { string r; return Score(name, target, n, out r); }
        public static double Score(string name, string target, int n, out string reason) {
            reason = string.Empty;
            string nn = Normalize(name), nt = Normalize(target);
            if (nn.Length == 0 || nt.Length == 0) return 0.0;
            if (nn.Length >= 3 && nt.Contains(nn)) { reason = "contains"; return 1.0; }
            if (nt.Length >= 3 && nn.Contains(nt)) { reason = "contained"; return 1.0; }
            double best = Dice(nn, nt, n); reason = "whole";
            if (target.Length > name.Length) {
                foreach (string tok in target.Split(new[] { ' ', '\t', '\r', '\n', '-', '_', '|', ',', ';', ':', '.', '/', '\\', '(', ')', '[', ']', '"', '\'' }, StringSplitOptions.RemoveEmptyEntries)) {
                    string ntok = Normalize(tok);
                    if (ntok.Length < 3) continue;
                    double d = Dice(nn, ntok, n) * 0.9;
                    if (d > best) { best = d; reason = "token:" + tok; }
                }
            }
            return best;
        }
        public static Match[] TopMatches(string[] names, string target, int n, double threshold, int top) {
            var list = new List<Match>();
            if (names == null || string.IsNullOrWhiteSpace(target)) return list.ToArray();
            for (int i = 0; i < names.Length; i++) {
                string why; double s = Score(names[i], target, n, out why);
                if (s >= threshold) list.Add(new Match(i, Math.Round(s, 4), why));
            }
            // equal scores: the longer matched name wins (a title containing "Programming - Active Directory Administration"
            // must prefer that note over "Active Directory"), then alphabetical for determinism
            list.Sort((x, y) => { int c = y.Score.CompareTo(x.Score); if (c != 0) return c; c = names[y.Index].Length.CompareTo(names[x.Index].Length); if (c != 0) return c; return string.Compare(names[x.Index], names[y.Index], StringComparison.OrdinalIgnoreCase); });
            if (list.Count > top) list.RemoveRange(top, list.Count - top);
            return list.ToArray();
        }
    }

    // Quick-switcher style fuzzy matching (substring > word starts > subsequence), multi token.
    public class FuzzyResult { public double Score; public int[] Positions; }
    public static class Fuzzy {
        static bool IsWordStart(string t, int i) { if (i <= 0) return true; return !char.IsLetterOrDigit(t[i - 1]); }
        public static FuzzyResult Match(string query, string text) {
            if (string.IsNullOrWhiteSpace(query) || string.IsNullOrEmpty(text)) return null;
            string t = text.ToLowerInvariant();
            double total = 0; var set = new HashSet<int>();
            foreach (string tok in query.ToLowerInvariant().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                double s; var p = MatchToken(tok, t, out s);
                if (p == null) return null;
                total += s; foreach (int i in p) set.Add(i);
            }
            total -= Math.Max(0, t.Length - query.Length) * 0.05;
            var pos = new List<int>(set); pos.Sort();
            return new FuzzyResult { Score = total, Positions = pos.ToArray() };
        }
        static List<int> MatchToken(string q, string t, out double score) {
            score = 0;
            var res = new List<int>();
            // substring hits always outrank subsequence hits (which sum per character and could exceed them for long tokens)
            int idx = t.IndexOf(q, StringComparison.Ordinal);
            if (idx >= 0) {
                score = 1000 + (idx == 0 ? 40 : (IsWordStart(t, idx) ? 25 : 0)) - idx * 0.3;
                for (int i = 0; i < q.Length; i++) res.Add(idx + i);
                return res;
            }
            int from = 0, last = -2;
            foreach (char c in q) {
                int i = t.IndexOf(c, from);
                if (i < 0) return null;
                if (i != last + 1) {
                    for (int k = i; k < t.Length && k < i + 40; k++) { if (t[k] == c && IsWordStart(t, k)) { i = k; break; } }
                }
                score += 10 + (i == last + 1 ? 6 : 0) + (IsWordStart(t, i) ? 8 : 0) - Math.Max(0, i - from) * 0.4;
                res.Add(i); last = i; from = i + 1;
            }
            return res;
        }
    }

    // Quick switcher search in one pass: names, aliases and contents, one hit per note (its best), top 'max'.
    // Kind: 0 = name, 1 = alias, 2 = content. Scores: substring in name/alias >= 1000, content 500, fuzzy below.
    public class SearchHit { public int Index; public int Kind; public double Score; public int[] Positions; public string Detail; }
    public static class SearchEngine {
        public static SearchHit[] Find(string[] display, string[][] aliases, string[] texts, string query, int max) {
            var hits = new List<SearchHit>();
            if (display == null || string.IsNullOrWhiteSpace(query)) return hits.ToArray();
            string q = query.Trim();
            string[] tokens = q.ToLowerInvariant().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < display.Length; i++) {
                SearchHit best = null;
                var m = Fuzzy.Match(q, display[i]);
                if (m != null) best = new SearchHit { Index = i, Kind = 0, Score = m.Score, Positions = m.Positions, Detail = null };
                if (aliases != null && i < aliases.Length && aliases[i] != null) {
                    foreach (string a in aliases[i]) {
                        if (string.IsNullOrEmpty(a)) continue;
                        var am = Fuzzy.Match(q, a);
                        if (am != null && (best == null || am.Score > best.Score)) best = new SearchHit { Index = i, Kind = 1, Score = am.Score, Positions = new int[0], Detail = a };
                    }
                }
                if (texts != null && i < texts.Length && !string.IsNullOrEmpty(texts[i]) && (best == null || best.Score < 500)) {
                    string t = texts[i]; int first = -1; bool ok = true;
                    foreach (string tok in tokens) {
                        int p = t.IndexOf(tok, StringComparison.OrdinalIgnoreCase);
                        if (p < 0) { ok = false; break; }
                        if (first < 0 || p < first) first = p;
                    }
                    if (ok) {
                        int start = Math.Max(0, first - 30); int len = Math.Min(90, t.Length - start);
                        string snippet = System.Text.RegularExpressions.Regex.Replace(t.Substring(start, len), "\\s+", " ").Trim();
                        best = new SearchHit { Index = i, Kind = 2, Score = 500.0 - first * 0.001, Positions = new int[0], Detail = (start > 0 ? "…" : "") + snippet + "…" };
                    }
                }
                if (best != null) hits.Add(best);
            }
            hits.Sort((x, y) => { int c = y.Score.CompareTo(x.Score); if (c != 0) return c; return string.Compare(display[x.Index], display[y.Index], StringComparison.OrdinalIgnoreCase); });
            if (hits.Count > max) hits.RemoveRange(max, hits.Count - max);
            return hits.ToArray();
        }
    }

    // One vault file. Created by VaultScanner; PowerShell reads and writes the public fields like a PSCustomObject.
    public class VaultEntry {
        public string Name, Rel, Full, Ext, PathKey, Key, Html, Display, Folder, Error;
        public string Text = string.Empty;
        public bool IsNote, IsImage, Matchable = true;   // Matchable false: MatchIgnoreFolders or front matter "csc: ignore" (still searchable)
        public DateTime LastWrite;
        public string[] Aliases = new string[0], LinkTargets = new string[0], LinkLines = new string[0];
    }
    public class VaultScanResult {
        public VaultEntry[] Entries;
        public Dictionary<string, VaultEntry> ByKey = new Dictionary<string, VaultEntry>(StringComparer.Ordinal);
        public Dictionary<string, VaultEntry> ByPath = new Dictionary<string, VaultEntry>(StringComparer.Ordinal);
        public Dictionary<string, List<VaultEntry>> NotesByName = new Dictionary<string, List<VaultEntry>>(StringComparer.Ordinal);
        public Dictionary<string, List<VaultEntry>> FilesByName = new Dictionary<string, List<VaultEntry>>(StringComparer.Ordinal);
        public string Signature;   // SHA-1 over all path keys: changes when files appear or disappear
    }
    public class BacklinkSource { public VaultEntry Source; public List<string> Contexts = new List<string>(); }

    // Walks the vault (skipping .obsidian, .trash, .git entirely and the notes of IgnoreFolders), reads every note in
    // parallel (text, front matter aliases, "csc: ignore", link targets with their lines) and builds the lookup maps.
    // Pure C#: nothing of PowerShell runs on the worker threads.
    public static class VaultScanner {
        static readonly Regex Front = new Regex(@"\A---\r?\n(.*?)\r?\n---[ \t]*(\r?\n|\z)", RegexOptions.Singleline | RegexOptions.Compiled);
        static readonly Regex IgnoreRx = new Regex(@"^csc(?:-match)?\s*:\s*(ignore|false|no|off)\s*$", RegexOptions.Multiline | RegexOptions.Compiled);
        static readonly Regex Fence = new Regex(@"^[ \t]*(`{3,}|~{3,}).*?^[ \t]*\1[ \t]*\r?$", RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.Compiled);
        static readonly Regex Inline = new Regex(@"`[^`\r\n]*`", RegexOptions.Compiled);
        static readonly Regex Wiki = new Regex(@"!?\[\[([^\[\]\|#]*)(?:#[^\[\]\|]*)?(?:\\?\|[^\[\]]*)?\]\]", RegexOptions.Compiled);
        static readonly Regex MdLink = new Regex(@"(?<!!)\[[^\]]*\]\(([^)\s]+?\.md)(?:#[^)]*)?\)", RegexOptions.Compiled);
        static readonly string[] Builtin = { ".obsidian", ".trash", ".git" };

        public static VaultScanResult Scan(string root, string[] ignoreNames, string[] ignorePaths, string[] matchIgnoreFolders, string outputDir, string[] imageExts) {
            root = root.TrimEnd('\\');
            var names = new HashSet<string>(ignoreNames ?? new string[0], StringComparer.OrdinalIgnoreCase);
            var paths = new List<string>(); foreach (string p in ignorePaths ?? new string[0]) paths.Add(p.Replace('/', '\\').Trim('\\'));
            var mif = new List<string>(); foreach (string p in matchIgnoreFolders ?? new string[0]) mif.Add(p.Replace('/', '\\').Trim('\\'));
            var images = new HashSet<string>(imageExts ?? new string[0], StringComparer.OrdinalIgnoreCase);
            var list = new List<VaultEntry>();
            Walk(new DirectoryInfo(root), root, "", names, paths, list);
            var r = new VaultScanResult();
            foreach (var e in list) {
                e.Ext = Path.GetExtension(e.Rel).ToLowerInvariant();
                e.IsNote = e.Ext == ".md";
                e.IsImage = images.Contains(e.Ext);
                e.PathKey = e.Rel.Replace('\\', '/').ToLowerInvariant();
                e.Key = e.IsNote ? e.PathKey.Substring(0, e.PathKey.Length - 3) : e.PathKey;
                e.Name = Path.GetFileNameWithoutExtension(e.Rel);
                e.Display = e.Name;
                e.Folder = Path.GetDirectoryName(e.Rel) ?? string.Empty;
                e.Html = e.IsNote ? Path.Combine(outputDir, e.Rel.Substring(0, e.Rel.Length - 3) + ".html") : null;
                if (e.IsNote) {
                    foreach (string mi in mif) { if (e.Folder.Equals(mi, StringComparison.OrdinalIgnoreCase) || e.Folder.StartsWith(mi + "\\", StringComparison.OrdinalIgnoreCase)) { e.Matchable = false; break; } }
                }
            }
            var notes = new List<VaultEntry>(); foreach (var e in list) if (e.IsNote) notes.Add(e);
            System.Threading.Tasks.Parallel.ForEach(notes, Fill);
            var sig = new StringBuilder();
            foreach (var e in list) {
                r.ByPath[e.PathKey] = e;
                if (sig.Length > 0) sig.Append('|');
                sig.Append(e.PathKey);
                string fn = Path.GetFileName(e.Rel).ToLowerInvariant();
                List<VaultEntry> l;
                if (!r.FilesByName.TryGetValue(fn, out l)) { l = new List<VaultEntry>(); r.FilesByName[fn] = l; }
                l.Add(e);
                if (e.IsNote) {
                    r.ByKey[e.Key] = e;
                    string n = e.Name.ToLowerInvariant();
                    if (!r.NotesByName.TryGetValue(n, out l)) { l = new List<VaultEntry>(); r.NotesByName[n] = l; }
                    l.Add(e);
                }
            }
            foreach (var kv in r.NotesByName) { if (kv.Value.Count > 1) foreach (var e in kv.Value) if (e.Folder.Length > 0) e.Display = e.Name + " (" + e.Folder + ")"; }
            using (var sha = System.Security.Cryptography.SHA1.Create()) r.Signature = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(sig.ToString()))).Replace("-", "");
            r.Entries = list.ToArray();
            return r;
        }
        static void Walk(DirectoryInfo dir, string root, string relDir, HashSet<string> names, List<string> paths, List<VaultEntry> list) {
            IEnumerable<FileSystemInfo> items;
            try { items = dir.EnumerateFileSystemInfos(); } catch { return; }
            bool notesIgnored = false;
            if (relDir.Length > 0) {
                foreach (string part in relDir.Split('\\')) if (names.Contains(part)) { notesIgnored = true; break; }
                if (!notesIgnored) foreach (string ip in paths) if (relDir.Equals(ip, StringComparison.OrdinalIgnoreCase) || relDir.StartsWith(ip + "\\", StringComparison.OrdinalIgnoreCase)) { notesIgnored = true; break; }
            }
            foreach (var item in items) {
                var sub = item as DirectoryInfo;
                if (sub != null) {
                    if (Array.IndexOf(Builtin, sub.Name.ToLowerInvariant()) >= 0) continue;
                    Walk(sub, root, relDir.Length == 0 ? sub.Name : relDir + "\\" + sub.Name, names, paths, list);
                    continue;
                }
                var f = item as FileInfo;
                if (f == null) continue;
                if (notesIgnored && f.Extension.Equals(".md", StringComparison.OrdinalIgnoreCase)) continue;
                list.Add(new VaultEntry { Full = f.FullName, Rel = relDir.Length == 0 ? f.Name : relDir + "\\" + f.Name, LastWrite = f.LastWriteTimeUtc });
            }
        }
        // text, aliases, ignore flag and link targets (code blocks and inline code skipped) of one note
        public static void Fill(VaultEntry n) {
            try { n.Text = File.ReadAllText(n.Full); } catch (Exception ex) { n.Error = ex.Message; return; }
            if (n.Text.StartsWith("---")) {
                var fm = Front.Match(n.Text);
                if (fm.Success) { string front = fm.Groups[1].Value; if (IgnoreRx.IsMatch(front)) n.Matchable = false; n.Aliases = ListProperty(front, "aliases"); }
            }
            if (n.Text.IndexOf('[') >= 0) {
                string t = Inline.Replace(Fence.Replace(n.Text, string.Empty), string.Empty);
                var targets = new List<string>(); var lines = new List<string>();
                foreach (string raw in t.Split('\n')) {
                    if (raw.IndexOf('[') < 0) continue;
                    string line = raw.TrimEnd('\r');
                    foreach (System.Text.RegularExpressions.Match w in Wiki.Matches(line)) { string tg = w.Groups[1].Value.Trim(); if (tg.Length > 0) { targets.Add(tg); lines.Add(line); } }
                    foreach (System.Text.RegularExpressions.Match w in MdLink.Matches(line)) { string tg = w.Groups[1].Value.Trim(); if (tg.Length > 0) { targets.Add(tg); lines.Add(line); } }
                }
                n.LinkTargets = targets.ToArray(); n.LinkLines = lines.ToArray();
            }
        }
        // Obsidian link resolution: exact key, exact path, relative to the linking note, then shortest path with that name
        public static VaultEntry Resolve(VaultScanResult r, string target, string fromRel) {
            if (target == null) return null;
            string t = target.Trim().Replace('\\', '/').TrimStart('/');
            if (t.StartsWith("./")) t = t.Substring(2);
            if (t.Length == 0) return null;
            try { t = Uri.UnescapeDataString(t); } catch { }
            string tl = t.ToLowerInvariant();
            string keyNoExt = tl.EndsWith(".md") ? tl.Substring(0, tl.Length - 3) : tl;
            VaultEntry e;
            if (r.ByKey.TryGetValue(keyNoExt, out e)) return e;
            if (r.ByPath.TryGetValue(tl, out e)) return e;
            if (!string.IsNullOrEmpty(fromRel)) {
                string dir = Path.GetDirectoryName(fromRel);
                if (!string.IsNullOrEmpty(dir)) {
                    string dl = dir.Replace('\\', '/').ToLowerInvariant();
                    if (r.ByKey.TryGetValue(dl + "/" + keyNoExt, out e)) return e;
                    if (r.ByPath.TryGetValue(dl + "/" + tl, out e)) return e;
                }
            }
            string baseName = tl.Substring(tl.LastIndexOf('/') + 1), baseNoExt = keyNoExt.Substring(keyNoExt.LastIndexOf('/') + 1);
            List<VaultEntry> l;
            if (r.NotesByName.TryGetValue(baseNoExt, out l)) return Shortest(l);
            if (r.FilesByName.TryGetValue(baseName, out l)) return Shortest(l);
            return null;
        }
        static VaultEntry Shortest(List<VaultEntry> l) { VaultEntry best = null; foreach (var e in l) if (best == null || e.Rel.Length < best.Rel.Length) best = e; return best; }
        // target note key -> linking notes with up to 5 mention lines each (self links ignored)
        public static Dictionary<string, List<BacklinkSource>> Backlinks(VaultScanResult r) {
            var map = new Dictionary<string, Dictionary<string, BacklinkSource>>(StringComparer.Ordinal);
            var memo = new Dictionary<string, VaultEntry>(StringComparer.Ordinal);
            foreach (var n in r.Entries) {
                if (!n.IsNote || n.LinkTargets.Length == 0) continue;
                var seen = new HashSet<string>(StringComparer.Ordinal);
                for (int k = 0; k < n.LinkTargets.Length; k++) {
                    string t = n.LinkTargets[k], line = n.LinkLines[k];
                    string mk = n.Folder + "|" + t;
                    VaultEntry e;
                    if (!memo.TryGetValue(mk, out e)) { e = Resolve(r, t, n.Rel); memo[mk] = e; }
                    if (e == null || !e.IsNote || e.Key == n.Key) continue;
                    if (!seen.Add(e.Key + "|" + line)) continue;
                    Dictionary<string, BacklinkSource> bySource;
                    if (!map.TryGetValue(e.Key, out bySource)) { bySource = new Dictionary<string, BacklinkSource>(StringComparer.Ordinal); map[e.Key] = bySource; }
                    BacklinkSource bs;
                    if (!bySource.TryGetValue(n.Key, out bs)) { bs = new BacklinkSource { Source = n }; bySource[n.Key] = bs; }
                    if (bs.Contexts.Count < 5) bs.Contexts.Add(line.Trim());
                }
            }
            var result = new Dictionary<string, List<BacklinkSource>>(StringComparer.Ordinal);
            foreach (var kv in map) result[kv.Key] = new List<BacklinkSource>(kv.Value.Values);
            return result;
        }
        // "key: [a, b]" or "key:" followed by "- item" lines (same rules as Get-FrontmatterList in the script)
        public static string[] ListProperty(string front, string key) {
            var items = new List<string>();
            string[] lines = front.Split('\n');
            for (int i = 0; i < lines.Length; i++) {
                string line = lines[i].TrimEnd('\r');
                var m = Regex.Match(line, "^" + Regex.Escape(key) + @"\s*:\s*(.*)$");
                if (!m.Success) continue;
                string inline = m.Groups[1].Value.Trim();
                if (inline.StartsWith("[") && inline.EndsWith("]")) {
                    foreach (string part in inline.Substring(1, inline.Length - 2).Split(',')) { string v = part.Trim().Trim('"', '\''); if (v.Length > 0) items.Add(v); }
                }
                else if (inline.Length > 0) { items.Add(inline.Trim('"', '\'')); }
                else {
                    for (int j = i + 1; j < lines.Length; j++) {
                        var d = Regex.Match(lines[j].TrimEnd('\r'), @"^\s*-\s*(.+)$");
                        if (!d.Success) break;
                        string v = d.Groups[1].Value.Trim().Trim('"', '\''); if (v.Length > 0) items.Add(v);
                    }
                }
                break;
            }
            return items.ToArray();
        }
    }

    // Watches the vault; the UI polls TakeDirtyIfQuiet() so no PowerShell code runs on foreign threads.
    public class VaultWatcher : IDisposable {
        FileSystemWatcher _w; volatile bool _dirty; long _last;
        public void Start(string path) {
            Stop();
            _w = new FileSystemWatcher(path);
            _w.IncludeSubdirectories = true;
            _w.InternalBufferSize = 64 * 1024;   // bursts (sync tools, Obsidian rename) must not overflow the default 8 KB
            _w.NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.LastWrite | NotifyFilters.Size;
            _w.Changed += OnEvent; _w.Created += OnEvent; _w.Deleted += OnEvent; _w.Renamed += OnEvent;
            _w.Error += (s, e) => { _last = DateTime.UtcNow.Ticks; _dirty = true; };
            _w.EnableRaisingEvents = true;
        }
        void OnEvent(object s, FileSystemEventArgs e) {
            string p = e.FullPath ?? string.Empty;
            if (p.IndexOf("\\.obsidian\\", StringComparison.OrdinalIgnoreCase) >= 0 || p.IndexOf("\\.trash\\", StringComparison.OrdinalIgnoreCase) >= 0 || p.IndexOf("\\.git\\", StringComparison.OrdinalIgnoreCase) >= 0) return;
            _last = DateTime.UtcNow.Ticks; _dirty = true;
        }
        public bool TakeDirtyIfQuiet(int quietMs) {
            if (!_dirty) return false;
            if ((DateTime.UtcNow.Ticks - _last) < quietMs * TimeSpan.TicksPerMillisecond) return false;
            _dirty = false; return true;
        }
        public void Stop() { if (_w != null) { try { _w.EnableRaisingEvents = false; _w.Dispose(); } catch { } _w = null; } }
        public void Dispose() { Stop(); }
    }
}
'@

function Get-HelperVersion {
    <#
        Version stamp of the helper source: the first 8 bytes of its SHA-256 as the four parts of an assembly version.
        tools\Build-Helper.ps1 compiles dll\CheatSheetCompanionHelper.dll with exactly this AssemblyVersion, so a stale
        DLL is recognised without loading it. Must stay identical to the copy in Build-Helper.ps1.
    #>
    param([Parameter(Mandatory)][string]$Source)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $b = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Source.Replace("`r`n", "`n")))
    $sha.Dispose()
    return [version]::new([BitConverter]::ToUInt16($b, 0), [BitConverter]::ToUInt16($b, 2), [BitConverter]::ToUInt16($b, 4), [BitConverter]::ToUInt16($b, 6))
}

# The prebuilt dll\CheatSheetCompanionHelper.dll (part of the delivery like the WebView2 DLLs) is loaded when its version
# matches the source above; nothing is compiled or written to disk at run time. Without it, or when it is stale, the
# source is compiled in memory as in earlier versions (costs the in-process C# compiler, about 50 MB).
$script:HelperMode = 'memory'
try {
    $helperDll = Join-Path $(if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }) 'dll\CheatSheetCompanionHelper.dll'
    if (Test-Path -LiteralPath $helperDll) {
        $want = Get-HelperVersion -Source $helperSource
        $have = [System.Reflection.AssemblyName]::GetAssemblyName($helperDll).Version
        if ($have -eq $want) { Add-Type -Path $helperDll; $script:HelperMode = 'dll' }
        else { $script:HelperMode = 'memory (dll\CheatSheetCompanionHelper.dll is stale, run tools\Build-Helper.ps1)' }
    }
    else { $script:HelperMode = 'memory (dll\CheatSheetCompanionHelper.dll missing, run tools\Build-Helper.ps1)' }
}
catch { $script:HelperMode = "memory ($_)" }
if ($script:HelperMode -ne 'dll') { Add-Type -TypeDefinition $helperSource }
Remove-Variable helperSource

#endregion

#region ---------- Paths, configuration, state ----------

$script:App = @{
    Root = $(if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path })
}
$script:App.DataDir    = Join-Path $env:LocalAppData 'CheatSheetCompanion'
$script:App.LogFile    = Join-Path $script:App.DataDir 'csc.log'
$script:App.StateFile  = Join-Path $script:App.DataDir 'state.json'
$script:App.OutputRoot = Join-Path $script:App.Root 'Output'
$script:App.ViewFile   = Join-Path $script:App.OutputRoot '_view.html'
$script:App.ThemeDir   = Join-Path $script:App.Root 'themes'
New-Item -ItemType Directory -Path $script:App.DataDir, $script:App.OutputRoot -Force | Out-Null

$script:Config = Get-IniContent -Path (Join-Path $script:App.Root 'ini\config.ini')
# ini\config.local.ini (not in git) overrides single keys: personal vault path, language, ignore folders ... stay out of the repo
$localIni = Join-Path $script:App.Root 'ini\config.local.ini'
if (Test-Path -LiteralPath $localIni) {
    $local = Get-IniContent -Path $localIni
    foreach ($sec in $local.Keys) {
        if (-not $script:Config.Contains($sec)) { $script:Config[$sec] = [ordered]@{} }
        foreach ($k in $local[$sec].Keys) { $script:Config[$sec][$k] = $local[$sec][$k] }
    }
}
Initialize-Language -Requested ([string](Get-Setting -Section Configuration -Key Language -Default 'auto'))

$script:App.Debug = (Get-Setting -Section Configuration -Key Debug -Default $false) -or ($VerbosePreference -eq 'Continue')
$script:App.SourceLinkTemplate = Get-Setting -Section Configuration -Key SourceLink -Default ''
$script:App.NewNoteFolder = Get-Setting -Section Configuration -Key NewNoteFolder -Default ''
$script:App.NewNoteTemplate = Get-Setting -Section Configuration -Key NewNoteTemplate -Default ''
$script:App.NewNoteNameMaxChars = [Math]::Max(10, (Get-Setting -Section Configuration -Key NewNoteNameMaxChars -Default 60))
$script:App.IgnoreFolders = @('.obsidian', '.trash', '.git') + @((Get-Setting -Section Configuration -Key IgnoreFolders -Default '') -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
# HTML cache: a new script version re-converts every note once (changed markup, e.g. the note meta data for Dataview)
$cacheStamp = Join-Path $script:App.OutputRoot 'cache.version'
$script:App.Version = '3.4.0'
$cacheFormat = '3.4.0'   # bump when the generated HTML changes (note meta data, markup); forces one full conversion
$script:App.ForceConvert = -not ((Test-Path -LiteralPath $cacheStamp) -and ((Get-Content -LiteralPath $cacheStamp -Raw -ErrorAction SilentlyContinue) -eq $cacheFormat))
if ($script:App.ForceConvert) { try { Set-Content -LiteralPath $cacheStamp -Value $cacheFormat -NoNewline } catch { } }
$script:App.NGram = [Math]::Max(1, (Get-Setting -Section Configuration -Key NGram -Default 3))
$script:App.MaxCandidates = [Math]::Max(1, (Get-Setting -Section Configuration -Key MaxCandidates -Default 3))
$script:App.ThresholdDefault = Get-Setting -Section Configuration -Key SimilarityThreshold -Default 0.25
$script:App.ThresholdTitle = Get-Setting -Section Configuration -Key SimilarityThresholdTitle -Default $script:App.ThresholdDefault
$script:App.ThresholdProcess = Get-Setting -Section Configuration -Key SimilarityThresholdProcess -Default $script:App.ThresholdDefault
$script:App.ThresholdClipboard = Get-Setting -Section Configuration -Key SimilarityThresholdClipboard -Default 0.4
# process name hits count less than window title hits: "Obsidian" as process must not beat the note named in the title
$script:App.ProcessWeight = [Math]::Min(1.0, [Math]::Max(0.1, [double](Get-Setting -Section Configuration -Key ProcessWeight -Default 0.9)))
$script:App.PollIntervalMs = [Math]::Max(200, (Get-Setting -Section Configuration -Key PollIntervalMs -Default 500))
$script:App.ClipboardMaxChars = [Math]::Max(20, (Get-Setting -Section Configuration -Key ClipboardMaxChars -Default 400))
$script:App.AutoCopyMinHoldMs = Get-Setting -Section Configuration -Key AutoCopyMinHoldMs -Default 200
$script:App.StrictLineBreaks = Get-Setting -Section Configuration -Key StrictLineBreaks -Default $false
$script:App.ShowFrontmatter = Get-Setting -Section Configuration -Key ShowFrontmatter -Default $false
$script:App.SearchMaxResults = [Math]::Max(3, (Get-Setting -Section Configuration -Key SearchMaxResults -Default 12))
$script:App.SearchContent = Get-Setting -Section Configuration -Key SearchContent -Default $true   # false: note texts are dropped after indexing, search by name/alias only
$script:App.EditOnDoubleClick = ([string](Get-Setting -Section Configuration -Key EditOnDoubleClick -Default 'ctrl')).ToLowerInvariant()
if ($script:App.EditOnDoubleClick -notin 'ctrl', 'always', 'off') { $script:App.EditOnDoubleClick = 'ctrl' }
$script:App.NoMatchBehavior = ([string](Get-Setting -Section Configuration -Key NoMatchBehavior -Default 'keep')).ToLowerInvariant()
$script:App.RecentMax = [Math]::Max(0, (Get-Setting -Section Configuration -Key RecentNotes -Default 8))
$script:App.ShowMatchReason = Get-Setting -Section Configuration -Key ShowMatchReason -Default $true
$script:App.MatchIgnoreFolders = @((Get-Setting -Section Configuration -Key MatchIgnoreFolders -Default '') -split ',' | ForEach-Object { $_.Trim().Trim('\', '/') -replace '/', '\' } | Where-Object { $_ })
$script:App.CopyHistoryMax = [Math]::Max(0, (Get-Setting -Section Configuration -Key CopyHistory -Default 10))
$script:App.ShowBacklinks = Get-Setting -Section Configuration -Key ShowBacklinks -Default $true
$script:App.LinkPreview = Get-Setting -Section Configuration -Key LinkPreview -Default $true   # hover popup for internal links
$script:App.WindowOpacity = [Math]::Min(1.0, [Math]::Max(0.3, [double](Get-Setting -Section Configuration -Key WindowOpacity -Default 0.92)))
$script:App.InactiveOpacity = [Math]::Min(1.0, [Math]::Max(0.0, [double](Get-Setting -Section Configuration -Key InactiveOpacity -Default 0.0)))
if ($script:App.InactiveOpacity -gt 0 -and $script:App.InactiveOpacity -lt 0.2) { $script:App.InactiveOpacity = 0.2 }   # never invisible
# Opacity is applied to the whole window (layered window alpha, includes the WebView2 area) whenever any opacity < 1 is wanted.
$script:App.UseWindowAlpha = ($script:App.WindowOpacity -lt 0.999 -or $script:App.InactiveOpacity -gt 0)
$script:App.ImageExtensions = @('.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp', '.bmp', '.ico', '.avif')

# [Rules]: process name (exact) or "title:<regex>" -> note name / vault path (score 1.00, shown first)
$script:App.Rules = @()
if ($script:Config.Contains('Rules')) {
    foreach ($key in $script:Config['Rules'].Keys) {
        $note = [string]$script:Config['Rules'][$key]
        if (-not $note) { continue }
        if ($key -match '^title\s*:\s*(.+)$') {
            try { $script:App.Rules += @{ Kind = 'title'; Regex = [regex]::new($matches[1].Trim(), 'IgnoreCase'); Note = $note; Spec = $key } }
            catch { Write-CSCLog "Rule '$key': invalid regex ($_)" 'Warn' }
        }
        else { $script:App.Rules += @{ Kind = 'process'; Process = $key.Trim(); Note = $note; Spec = $key } }
    }
}

# Runtime state (persisted toggles, theme, vault, zoom, window placement)
$script:State = @{
    FollowWindow    = Get-Setting -Section Configuration -Key FollowWindow -Default $true
    FollowClipboard = Get-Setting -Section Configuration -Key FollowClipboard -Default $false
    AutoCopy        = Get-Setting -Section Configuration -Key AutoCopy -Default $false
    Theme           = Get-Setting -Section Configuration -Key Theme -Default 'auto'
    Vault           = ''
    Zoom            = 1.0
    RecentNotes     = @()          # "vault|relative path", newest first
    PinnedNotes     = @()          # "vault|relative path" of pinned notes: stay as tabs while matching goes on
    CopyHistory     = @()          # texts copied through the companion (Text, Note, Time), newest first
    Left = $null; Top = $null; Width = $null; Height = $null
}
try {
    if (Test-Path -LiteralPath $script:App.StateFile) {
        $saved = Get-Content -LiteralPath $script:App.StateFile -Raw | ConvertFrom-Json
        foreach ($p in $saved.PSObject.Properties) { if ($script:State.ContainsKey($p.Name)) { $script:State[$p.Name] = $p.Value } }
        if ($saved.PSObject.Properties['PinnedNote'] -and $saved.Pinned -and $saved.PinnedNote) { $script:State.PinnedNotes = @([string]$saved.PinnedNote) }   # single pin of v3.2
        $script:State.PinnedNotes = @(@($script:State.PinnedNotes) | Where-Object { $_ } | ForEach-Object { [string]$_ })
    }
}
catch { Write-CSCLog "State file unreadable: $_" 'Warn' }

function Save-State {
    try {
        $w = $script:UI.Window
        if ($w -and $w.WindowState -eq 'Normal' -and $w.Height -gt 20) {
            $script:State.Left = $w.Left; $script:State.Top = $w.Top; $script:State.Width = $w.Width; $script:State.Height = $w.Height
        }
        $script:State | ConvertTo-Json -Compress -Depth 4 | Set-Content -LiteralPath $script:App.StateFile -Encoding UTF8
    }
    catch { Write-CSCLog "Save-State failed: $_" 'Warn' }
}

# ---- vault list: [Vaults] section + MarkDownFolder; relative paths count from the script folder, .\vault (demo vault) is the last resort ----
function Resolve-VaultPath {
    param([AllowEmptyString()][string]$Path)
    $p = $Path.Trim()
    if (-not $p) { return '' }
    if (-not [System.IO.Path]::IsPathRooted($p)) { $p = Join-Path $script:App.Root $p }
    return $p
}
$script:App.Vaults = [ordered]@{}
$mdFolder = Resolve-VaultPath -Path ([string](Get-Setting -Section Configuration -Key MarkDownFolder -Default ''))
if ($mdFolder -and (Test-Path -LiteralPath $mdFolder)) {
    $script:App.Vaults[(Split-Path $mdFolder -Leaf)] = (Get-Item -LiteralPath $mdFolder).FullName.TrimEnd('\')
}
elseif ($mdFolder) { Write-CSCLog "MarkDownFolder '$mdFolder' not found" 'Warn' }
if ($script:Config.Contains('Vaults')) {
    foreach ($key in $script:Config['Vaults'].Keys) {
        $p = Resolve-VaultPath -Path ([string]$script:Config['Vaults'][$key])
        if ($p -and (Test-Path -LiteralPath $p)) {
            $full = (Get-Item -LiteralPath $p).FullName.TrimEnd('\')
            if (-not ($script:App.Vaults.Values -contains $full)) { $script:App.Vaults[$key] = $full }
        }
        else { Write-CSCLog "Vault '$key' path '$p' not found" 'Warn' }
    }
}
if ($script:App.Vaults.Count -eq 0) { $script:App.Vaults['vault'] = Join-Path $script:App.Root 'vault' }

function ConvertTo-WindowsFilename {
    param([Parameter(Mandatory)][string]$Name)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() + @('[', ']', '#', '^', '|')
    $sb = [System.Text.StringBuilder]::new()
    foreach ($ch in $Name.ToCharArray()) { [void]$sb.Append($(if ($invalid -contains $ch) { '_' } else { $ch })) }
    $clean = $sb.ToString().Trim().TrimEnd('.', ' ')
    if (-not $clean) { $clean = Get-Text 'app.newNoteDefault' }
    return $clean
}

function Initialize-Vault {
    <# Activates a vault: paths, output cache folder, index; remaps the WebView host when already running. #>
    param([Parameter(Mandatory)][string]$Name)
    if (-not $script:App.Vaults.Contains($Name)) { $Name = [string]$script:App.Vaults.Keys[0] }
    $script:App.VaultName = $Name
    $script:App.Vault = $script:App.Vaults[$Name]
    $script:App.Output = Join-Path $script:App.OutputRoot (ConvertTo-WindowsFilename $Name)
    $script:App.SourceLink = $script:App.SourceLinkTemplate
    New-Item -ItemType Directory -Path $script:App.Output -Force | Out-Null
    $script:State.Vault = $Name
    $script:Index = @{ Entries = @(); Notes = @(); NoteNames = [string[]]@(); ByKey = @{}; ByPath = @{}; NotesByName = @{}; FilesByName = @{}; Signature = '' }
    [void](Update-VaultIndex)
    Write-CSCLog ("Vault '{0}' active: {1} ({2} notes)" -f $Name, $script:App.Vault, $script:Index.Notes.Count)
    if ($script:UI -and $script:UI.WebReady) {
        try {
            $core = $script:UI.WebView.CoreWebView2
            $core.ClearVirtualHostNameToFolderMapping('vault.csc')
            $core.SetVirtualHostNameToFolderMapping('vault.csc', $script:App.Vault, [Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind]::Allow)
        }
        catch { Write-CSCLog "Remap vault host: $_" 'Warn' }
        try { $script:UI.Watcher.Start($script:App.Vault) } catch { }
        $script:UI.History.Clear(); $script:UI.Forward.Clear(); $script:UI.CurrentView = $null; $script:UI.CurrentNote = $null
        Show-Welcome
        Set-Status (Get-Text 'status.vault' -Format $Name, $script:Index.Notes.Count)
    }
}

Write-CSCLog ("Start v{5} | Vaults={0} | Rules={1} | Language={2} | Debug={3} | Helper={4}" -f ($script:App.Vaults.Keys -join ','), $script:App.Rules.Count, $script:App.Language, $script:App.Debug, $script:HelperMode, $script:App.Version)

#endregion

#region ---------- Markdown (Markdig + Obsidian syntax) ----------

try { [void][Markdig.Markdown] } catch { Add-Type -Path (Join-Path $PSHOME 'Markdig.Signed.dll') }
$mdBuilder = [Markdig.MarkdownPipelineBuilder]::new()
[void][Markdig.MarkdownExtensions]::UseAdvancedExtensions($mdBuilder)
[void][Markdig.MarkdownExtensions]::UseYamlFrontMatter($mdBuilder)
if (-not $script:App.StrictLineBreaks) { [void][Markdig.MarkdownExtensions]::UseSoftlineBreakAsHardlineBreak($mdBuilder) }
$script:MdPipeline = $mdBuilder.Build()

function ConvertTo-HtmlAttribute {
    param([AllowNull()][string]$Text)
    return [System.Net.WebUtility]::HtmlEncode([string]$Text)
}

function ConvertTo-VaultUrl {
    <# Vault relative path -> https://vault.csc/... (WebView2 virtual host) #>
    param([Parameter(Mandatory)][string]$RelativePath)
    $segments = ($RelativePath -replace '\\', '/') -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }
    return 'https://vault.csc/' + ($segments -join '/')
}

function Resolve-VaultLink {
    <#
        .SYNOPSIS
        Resolves an Obsidian link target ("Note", "folder/Note", "image.png", "Note.md") like Obsidian does:
        exact path first, then relative to the linking note, then by (shortest) basename.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Target,
        [string]$FromRelativePath = ''
    )
    $t = ($Target.Trim() -replace '\\', '/').TrimStart('/')
    if ($t.StartsWith('./')) { $t = $t.Substring(2) }
    if ($t.Length -eq 0) { return $null }
    try { $t = [uri]::UnescapeDataString($t) } catch { }
    $tl = $t.ToLowerInvariant()
    $keyNoExt = if ($tl.EndsWith('.md')) { $tl.Substring(0, $tl.Length - 3) } else { $tl }
    $idx = $script:Index

    if ($idx.ByKey.ContainsKey($keyNoExt)) { return $idx.ByKey[$keyNoExt] }
    if ($idx.ByPath.ContainsKey($tl)) { return $idx.ByPath[$tl] }
    if ($FromRelativePath) {
        $dir = (Split-Path $FromRelativePath -Parent) -replace '\\', '/'
        if ($dir) {
            $dl = $dir.ToLowerInvariant()
            if ($idx.ByKey.ContainsKey("$dl/$keyNoExt")) { return $idx.ByKey["$dl/$keyNoExt"] }
            if ($idx.ByPath.ContainsKey("$dl/$tl")) { return $idx.ByPath["$dl/$tl"] }
        }
    }
    $base = ($tl -split '/')[-1]
    $baseNoExt = ($keyNoExt -split '/')[-1]
    if ($idx.NotesByName.ContainsKey($baseNoExt)) { return ($idx.NotesByName[$baseNoExt] | Sort-Object { $_.Rel.Length } | Select-Object -First 1) }
    if ($idx.FilesByName.ContainsKey($base)) { return ($idx.FilesByName[$base] | Sort-Object { $_.Rel.Length } | Select-Object -First 1) }
    return $null
}

function Protect-CodeSegment {
    <# Replaces fenced code blocks, inline code and YAML front matter by placeholders so they are not touched. #>
    param([AllowEmptyString()][string]$Markdown, [AllowEmptyCollection()][System.Collections.Generic.List[string]]$Store)
    $patterns = @(
        '(?s)\A---\r?\n.*?\r?\n---[ \t]*(\r?\n|\z)',                          # front matter
        '(?ms)^[ \t]{0,3}(`{3,}|~{3,})[^\n]*\n.*?^[ \t]{0,3}\1[ \t]*\r?$',   # fenced code (closing fence also with CR before LF)
        '``[^\n]*?``',                                                       # inline code with double backticks
        '`[^`\n]*`'                                                          # inline code
    )
    $text = $Markdown
    foreach ($p in $patterns) {
        $text = [regex]::Replace($text, $p, {
                param($m)
                $Store.Add($m.Value)
                return ([char]1) + [string]($Store.Count - 1) + ([char]1)
            })
    }
    return $text
}

function Restore-CodeSegment {
    param([AllowEmptyString()][string]$Text, [AllowEmptyCollection()][System.Collections.Generic.List[string]]$Store)
    return [regex]::Replace($Text, "`u{1}(\d+)`u{1}", { param($m) $Store[[int]$m.Groups[1].Value] })
}

function Convert-ObsidianSyntax {
    <# Pre-processes Obsidian specific markdown: %%comments%%, ![[embeds]], [[wikilinks]], #tags. #>
    param([AllowEmptyString()][string]$Markdown, [Parameter(Mandatory)][string]$FromRelativePath)

    if ([string]::IsNullOrWhiteSpace($Markdown)) { return '' }
    $store = [System.Collections.Generic.List[string]]::new()
    $text = Protect-CodeSegment -Markdown $Markdown -Store $store
    $text = [regex]::Replace($text, '(?s)%%.*?%%', '')

    $linkPattern = '\[\[([^\[\]\|#]*)(?:#([^\[\]\|]*))?(?:\\?\|([^\[\]]*))?\]\]'

    # Embeds: images, notes, other files
    $text = [regex]::Replace($text, '!' + $linkPattern, {
            param($m)
            $target = $m.Groups[1].Value.Trim(); $anchor = $m.Groups[2].Value.Trim(); $param = $m.Groups[3].Value.Trim()
            $entry = Resolve-VaultLink -Target $target -FromRelativePath $FromRelativePath
            if ($null -eq $entry) { return '<span class="csc-missing" title="' + (ConvertTo-HtmlAttribute (Get-Text 'page.missingTarget')) + '">' + (ConvertTo-HtmlAttribute $m.Value) + '</span>' }
            if ($entry.IsImage) {
                $size = ''
                if ($param -match '^(\d+)(?:x(\d+))?$') {
                    $size = ' width="' + $matches[1] + '"'
                    if ($matches[2]) { $size += ' height="' + $matches[2] + '"' }
                }
                return '<img src="' + (ConvertTo-VaultUrl $entry.Rel) + '" alt="' + (ConvertTo-HtmlAttribute $entry.Name) + '"' + $size + '>'
            }
            if ($entry.IsNote) {
                return '<div class="csc-embed" data-rel="' + (ConvertTo-HtmlAttribute $entry.Rel) + '" data-anchor="' + (ConvertTo-HtmlAttribute $anchor) + '"></div>'
            }
            return '<a class="internal-link" href="' + (ConvertTo-VaultUrl $entry.Rel) + '">' + (ConvertTo-HtmlAttribute $entry.Name) + '</a>'
        })

    # Wikilinks
    $text = [regex]::Replace($text, '(?<!!)' + $linkPattern, {
            param($m)
            $target = $m.Groups[1].Value.Trim(); $anchor = $m.Groups[2].Value.Trim(); $alias = $m.Groups[3].Value.Trim()
            $label = if ($alias) { $alias } elseif ($target) { $target + $(if ($anchor) { ' > ' + $anchor } else { '' }) } else { $anchor }
            if (-not $target) {
                return '<a class="internal-link" href="#" data-rel="" data-anchor="' + (ConvertTo-HtmlAttribute $anchor) + '">' + (ConvertTo-HtmlAttribute $label) + '</a>'
            }
            $entry = Resolve-VaultLink -Target $target -FromRelativePath $FromRelativePath
            if ($null -eq $entry) {
                return '<a class="internal-link is-unresolved" href="#" data-note="' + (ConvertTo-HtmlAttribute $target) + '" title="' + (ConvertTo-HtmlAttribute (Get-Text 'page.missingNote')) + '">' + (ConvertTo-HtmlAttribute $label) + '</a>'
            }
            if ($entry.IsNote) {
                return '<a class="internal-link" href="#" data-rel="' + (ConvertTo-HtmlAttribute $entry.Rel) + '" data-anchor="' + (ConvertTo-HtmlAttribute $anchor) + '">' + (ConvertTo-HtmlAttribute $label) + '</a>'
            }
            return '<a class="internal-link" href="' + (ConvertTo-VaultUrl $entry.Rel) + '">' + (ConvertTo-HtmlAttribute $label) + '</a>'
        })

    # Tags (#tag, not headings, not inside words/URLs/entities)
    $text = [regex]::Replace($text, '(?<=^|[\s(>])#([\p{L}_][\p{L}\p{N}_/-]*)(?![\p{L}\p{N}_/-])', { param($m) '<span class="tag">#' + $m.Groups[1].Value + '</span>' })

    return Restore-CodeSegment -Text $text -Store $store
}

function Get-FrontmatterList {
    <# Reads a list property from YAML front matter: "key: [a, b]" or "key:" followed by "- item" lines. #>
    param([AllowEmptyString()][string]$Markdown, [Parameter(Mandatory)][string]$Key)
    $m = [regex]::Match($Markdown, '(?s)\A---\r?\n(.*?)\r?\n---[ \t]*(\r?\n|\z)')
    if (-not $m.Success) { return @() }
    $lines = $m.Groups[1].Value -split '\r?\n'
    $items = [System.Collections.Generic.List[string]]::new()
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -notmatch ('^' + [regex]::Escape($Key) + '\s*:\s*(.*)$')) { continue }
        $inline = $matches[1].Trim()
        if ($inline -match '^\[(.*)\]$') {
            foreach ($part in ($matches[1] -split ',')) { $v = $part.Trim().Trim('"', "'"); if ($v) { $items.Add($v) } }
        }
        elseif ($inline) { $items.Add($inline.Trim('"', "'")) }
        else {
            for ($j = $i + 1; $j -lt $lines.Count; $j++) {
                if ($lines[$j] -match '^\s*-\s*(.+)$') { $v = $matches[1].Trim().Trim('"', "'"); if ($v) { $items.Add($v) } } else { break }
            }
        }
        break
    }
    return $items.ToArray()
}

function ConvertTo-FrontmatterValue {
    <# YAML scalar -> string, bool or number (quotes removed). #>
    param([AllowEmptyString()][string]$Raw)
    $v = $Raw.Trim()
    if ($v -match '^"(.*)"$' -or $v -match "^'(.*)'$") { return $matches[1] }
    if ($v -match '^(true|false)$') { return ($v -eq 'true') }
    $d = 0.0
    if ($v -match '^-?\d+(\.\d+)?$' -and [double]::TryParse($v, [System.Globalization.NumberStyles]::Float, [cultureinfo]::InvariantCulture, [ref]$d)) { return $d }
    return $v
}

function ConvertFrom-FrontmatterSimple {
    <# Flat YAML front matter -> ordered hashtable: scalars, inline lists [a, b] and dash lists. Nested maps become $null. #>
    param([AllowEmptyString()][string]$Markdown)
    $result = [ordered]@{}
    $m = [regex]::Match($Markdown, '(?s)\A---\r?\n(.*?)\r?\n---[ \t]*(\r?\n|\z)')
    if (-not $m.Success) { return $result }
    $lines = @($m.Groups[1].Value -split '\r?\n')
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -notmatch '^([A-Za-z0-9_][^:#]*?)\s*:\s*(.*)$') { continue }
        $key = $matches[1].Trim(); $val = $matches[2].Trim()
        if ($val -match '^\[(.*)\]$') { $result[$key] = @($matches[1] -split ',' | ForEach-Object { $_.Trim().Trim('"', "'") } | Where-Object { $_ }) }
        elseif ($val) { $result[$key] = ConvertTo-FrontmatterValue -Raw $val }
        else {
            $items = @(); $j = $i + 1
            while ($j -lt $lines.Count -and $lines[$j] -match '^\s*-\s*(.*)$') { $items += $matches[1].Trim().Trim('"', "'"); $j++ }
            if ($items.Count -gt 0) { $result[$key] = $items; $i = $j - 1 } else { $result[$key] = $null }
        }
    }
    return $result
}

function Get-InlineFields {
    <#
        Dataview inline fields in the note body: "key:: value" at the start of a line (also inside %% comments and list
        items) and "[key:: value]" / "(key:: value)" inside a line. Code blocks are skipped. Keys are stored as written
        and normalised like Dataview does (lower case, spaces as dashes).
    #>
    param([AllowEmptyString()][string]$Markdown)
    $fields = [ordered]@{}
    if (-not $Markdown -or $Markdown.IndexOf('::') -lt 0) { return $fields }
    $text = [regex]::Replace($Markdown, '(?ms)^[ \t]*(`{3,}|~{3,}).*?^[ \t]*\1[ \t]*\r?$', '')
    $text = [regex]::Replace($text, '`[^`\r\n]*`', '')
    $add = {
        param($k, $v)
        $k = $k.Trim(); if (-not $k) { return }
        $val = ConvertTo-FrontmatterValue -Raw $v
        $fields[$k] = $val
        $norm = ($k.ToLowerInvariant() -replace '\s+', '-')
        if ($norm -ne $k) { $fields[$norm] = $val }
    }
    foreach ($m in [regex]::Matches($text, '(?m)^[ \t]*(?:[-*+]\s+)?([A-Za-z][A-Za-z0-9_ -]*?)::[ \t]*(.*?)[ \t]*\r?$')) { & $add $m.Groups[1].Value $m.Groups[2].Value }
    foreach ($m in [regex]::Matches($text, '[\[(]([A-Za-z][A-Za-z0-9_ -]*?)::[ \t]*([^\])\r\n]*)[\])]')) { & $add $m.Groups[1].Value $m.Groups[2].Value }
    return $fields
}

function Get-NoteMetaScript {
    <# Note meta data (file info, front matter, inline fields) as JSON inside the note HTML; app.js feeds it to dataviewjs blocks as dv.current(). #>
    param([Parameter(Mandatory)]$Entry, [AllowEmptyString()][string]$Markdown)
    $props = Get-InlineFields -Markdown $Markdown
    $front = ConvertFrom-FrontmatterSimple -Markdown $Markdown
    foreach ($k in $front.Keys) { $props[$k] = $front[$k] }   # front matter wins over inline fields with the same key
    $meta = [ordered]@{
        file        = [ordered]@{ name = $Entry.Name; path = ($Entry.Rel -replace '\\', '/'); folder = (($Entry.Folder -replace '\\', '/')); link = $Entry.Name; ext = 'md' }
        frontmatter = $props
        content     = $Markdown   # app.vault.read() of the current note inside dataviewjs
    }
    $json = (ConvertTo-Json -InputObject $meta -Compress -Depth 5) -replace '</', '<\/'
    return '<script type="application/json" class="csc-note-meta">' + $json + '</script>' + "`n"
}

function Convert-Frontmatter {
    <# Optional: renders YAML front matter (key: value lines) as a collapsed properties box. #>
    param([AllowEmptyString()][string]$Markdown)
    if (-not $script:App.ShowFrontmatter) { return '' }
    $m = [regex]::Match($Markdown, '(?s)\A---\r?\n(.*?)\r?\n---[ \t]*(\r?\n|\z)')
    if (-not $m.Success) { return '' }
    $rows = [System.Text.StringBuilder]::new()
    foreach ($line in ($m.Groups[1].Value -split '\r?\n')) {
        if ($line -match '^\s*([^:#]+?)\s*:\s*(.*)$') {
            [void]$rows.Append('<tr><td>' + (ConvertTo-HtmlAttribute $matches[1]) + '</td><td>' + (ConvertTo-HtmlAttribute $matches[2]) + '</td></tr>')
        }
    }
    if ($rows.Length -eq 0) { return '' }
    return '<details class="csc-frontmatter"><summary>' + (ConvertTo-HtmlAttribute (Get-Text 'page.properties')) + '</summary><table>' + $rows.ToString() + '</table></details>' + "`n"
}

function Convert-Callout {
    <# Turns Obsidian callouts (> [!type]± Title) into styled blocks. #>
    param([AllowEmptyString()][string]$Html)
    if (-not $Html) { return '' }
    $aliases = @{
        'hint' = 'tip'; 'important' = 'tip'; 'caution' = 'warning'; 'attention' = 'warning'; 'error' = 'danger'
        'check' = 'success'; 'done' = 'success'; 'help' = 'question'; 'faq' = 'question'; 'fail' = 'failure'; 'missing' = 'failure'
        'cite' = 'quote'; 'summary' = 'abstract'; 'tldr' = 'abstract'
    }
    $build = {
        param([string]$type, [string]$fold, [string]$title, [string]$content)
        $type = $type.ToLowerInvariant()
        if ($aliases.ContainsKey($type)) { $type = $aliases[$type] }
        $title = $title.Trim()
        if (-not $title) { $title = $type.Substring(0, 1).ToUpperInvariant() + $type.Substring(1) }
        $cls = 'callout'
        if ($fold) { $cls += ' is-collapsible'; if ($fold -eq '-') { $cls += ' is-collapsed' } }
        $foldHtml = if ($fold) { '<span class="callout-fold">&#9660;</span>' } else { '' }
        return '<div class="' + $cls + '" data-callout="' + $type + '"><div class="callout-title"><span class="callout-icon"></span><span class="callout-title-inner">' + $title + '</span>' + $foldHtml + '</div><div class="callout-content">' + $content + '</div></div>'
    }
    $Html = [regex]::Replace($Html, '(?s)<blockquote>\s*<p>\[!([A-Za-z0-9_-]+)\]([+-]?)[ \t]*([^\r\n<]*)(?:<br />)?\r?\n?(.*?)</blockquote>', {
            param($m)
            $rest = $m.Groups[4].Value
            $content = if ($rest.TrimStart().StartsWith('</p>')) { $rest.TrimStart().Substring(4) } else { '<p>' + $rest }
            & $build $m.Groups[1].Value $m.Groups[2].Value $m.Groups[3].Value $content
        })
    $Html = [regex]::Replace($Html, '(?s)<div class="markdown-alert markdown-alert-([a-z]+)">\s*<p class="markdown-alert-title">(?:<svg.*?</svg>)?([^<]*)</p>(.*?)</div>', {
            param($m)
            & $build $m.Groups[1].Value '' '' $m.Groups[3].Value
        })
    return $Html
}

function Resolve-HtmlResource {
    <# Rewrites relative <img src> and <a href> targets (standard markdown links) to vault URLs / internal links. #>
    param([AllowEmptyString()][string]$Html, [Parameter(Mandatory)][string]$FromRelativePath)

    if (-not $Html) { return '' }
    $Html = [regex]::Replace($Html, '<img\b([^>]*?)\ssrc="([^"]*)"', {
            param($m)
            $src = $m.Groups[2].Value
            if ($src -match '^(https?:|data:|file:|blob:)') { return $m.Value }
            $entry = Resolve-VaultLink -Target $src -FromRelativePath $FromRelativePath
            if ($null -eq $entry) { return '<img class="csc-missing"' + $m.Groups[1].Value + ' src="' + $src + '"' }
            return '<img' + $m.Groups[1].Value + ' src="' + (ConvertTo-VaultUrl $entry.Rel) + '"'
        })

    $Html = [regex]::Replace($Html, '<a\b([^>]*?)\shref="([^"]*)"', {
            param($m)
            $attrs = $m.Groups[1].Value; $href = $m.Groups[2].Value
            if ($attrs -match 'internal-link' -or $href.StartsWith('#') -or $href -match '^[a-zA-Z][a-zA-Z0-9+.-]*:') { return $m.Value }
            $anchor = ''
            if ($href.Contains('#')) { $anchor = $href.Substring($href.IndexOf('#') + 1); $href = $href.Substring(0, $href.IndexOf('#')) }
            $entry = Resolve-VaultLink -Target $href -FromRelativePath $FromRelativePath
            if ($null -eq $entry) { return $m.Value }
            if ($entry.IsNote) { return '<a class="internal-link"' + $attrs + ' href="#" data-rel="' + (ConvertTo-HtmlAttribute $entry.Rel) + '" data-anchor="' + (ConvertTo-HtmlAttribute ([uri]::UnescapeDataString($anchor))) + '"' }
            return '<a' + $attrs + ' href="' + (ConvertTo-VaultUrl $entry.Rel) + '"'
        })
    return $Html
}

function ConvertTo-NoteHtml {
    <# Converts one note to HTML (cached in Output\<vault>\<relative path>.html). #>
    param([Parameter(Mandatory)]$Entry)
    try {
        $markdown = [System.IO.File]::ReadAllText($Entry.Full)
        $front = (Get-NoteMetaScript -Entry $Entry -Markdown $markdown) + (Convert-Frontmatter -Markdown $markdown)
        $markdown = Convert-ObsidianSyntax -Markdown $markdown -FromRelativePath $Entry.Rel
        $html = [Markdig.Markdown]::ToHtml($markdown, $script:MdPipeline)
        $html = Convert-Callout -Html $html
        $html = Resolve-HtmlResource -Html ($front + $html) -FromRelativePath $Entry.Rel
        # notes with unresolved links are re-converted when files appear (index.cache)
        if ($html.IndexOf('is-unresolved') -ge 0 -or $html.IndexOf('csc-missing') -ge 0) { $script:Index.Unresolved[$Entry.PathKey] = $true } else { $script:Index.Unresolved.Remove($Entry.PathKey) }
        $dir = Split-Path $Entry.Html -Parent
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        [System.IO.File]::WriteAllText($Entry.Html, $html, [System.Text.UTF8Encoding]::new($false))
        return $true
    }
    catch {
        Write-CSCLog ("Convert failed for '{0}': {1}" -f $Entry.Rel, $_) 'Error'
        return $false
    }
}

#endregion

#region ---------- Vault index ----------

$script:Index = @{ Entries = @(); Notes = @(); NoteNames = [string[]]@(); ByKey = @{}; ByPath = @{}; NotesByName = @{}; FilesByName = @{}; Backlinks = @{}; Signature = ''; Unresolved = @{} }

function Get-BacklinkIndex {
    <#
        .SYNOPSIS
        Builds target note key -> (source note key -> @{ Source; Contexts }) from [[wikilinks]], ![[embeds]] and
        [text](note.md) links of all notes. Code blocks and inline code are skipped, self links ignored.
        Needs $script:Index (Resolve-VaultLink), so it runs after the index was assigned.
    #>
    param([AllowEmptyCollection()][object[]]$Notes)
    # kept for tests and tools: the app itself uses VaultScanner.Backlinks (C#) with the same rules
    if ($script:Index.Scan) { return [CheatSheetCompanionHelper.VaultScanner]::Backlinks($script:Index.Scan) }
    return @{}
}

function Update-VaultIndex {
    <#
        .SYNOPSIS
        Scans the vault, converts new/changed notes and removes orphaned HTML files.
        Returns $true when the set of files changed.
    #>
    param([switch]$Force)

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $root = $script:App.Vault
    # IgnoreFolders: plain names match any folder of that name at any depth, entries with a separator ("archive\old notes")
    # match that relative path (and everything below it). They drop the NOTES there; images and other attachments in those
    # folders stay known so embeds keep working (an "_attachments" entry is common). .obsidian, .trash and .git are skipped completely.
    # The walk, the parallel reading of all notes (text, aliases, ignore flag, link targets) and the lookup maps are
    # built in C# (VaultScanner); the entries are VaultEntry objects used like PSCustomObjects below.
    $builtin = @('.obsidian', '.trash', '.git')
    $ignoreNames = [string[]]@($script:App.IgnoreFolders | Where-Object { $_ -notmatch '[\\/]' -and $builtin -notcontains $_ })
    $ignorePaths = [string[]]@($script:App.IgnoreFolders | Where-Object { $_ -match '[\\/]' })
    $scan = [CheatSheetCompanionHelper.VaultScanner]::Scan($root, $ignoreNames, $ignorePaths, [string[]]@($script:App.MatchIgnoreFolders), $script:App.Output, [string[]]@($script:App.ImageExtensions))
    $tScan = $sw.ElapsedMilliseconds
    $entries = $scan.Entries
    $byKey = $scan.ByKey; $byPath = $scan.ByPath; $notesByName = $scan.NotesByName; $filesByName = $scan.FilesByName
    foreach ($e in $entries) { if ($e.Error) { Write-CSCLog ("Read failed for '{0}': {1}" -f $e.Rel, $e.Error) 'Warn' } }
    $tRead = $sw.ElapsedMilliseconds
    $notes = @($entries | Where-Object IsNote | Sort-Object Display)
    $signature = $scan.Signature
    # Output\<vault>\index.cache remembers the file set and the notes with unresolved links from the last run, so a
    # start without changes converts nothing; a changed file set re-converts only notes whose links may resolve now
    $cacheFile = Join-Path $script:App.Output 'index.cache'
    $prev = $null
    try { if (Test-Path -LiteralPath $cacheFile) { $prev = Get-Content -LiteralPath $cacheFile -Raw | ConvertFrom-Json } } catch { $prev = $null }
    $prevSignature = if ($script:Index.Signature) { [string]$script:Index.Signature } elseif ($prev) { [string]$prev.Signature } else { '' }
    $fileSetChanged = ($signature -ne $prevSignature)
    $unresolved = @{}
    if ($script:Index.Unresolved) { foreach ($u in $script:Index.Unresolved.Keys) { $unresolved[$u] = $true } }
    elseif ($prev -and $prev.Unresolved) { foreach ($u in @($prev.Unresolved)) { $unresolved[[string]$u] = $true } }

    # match names: note name + front matter aliases, parallel array of entries
    $matchNames = [System.Collections.Generic.List[string]]::new()
    $matchEntries = [System.Collections.Generic.List[object]]::new()
    foreach ($n in $notes) {
        if (-not $n.Matchable) { continue }
        $matchNames.Add($n.Name); $matchEntries.Add($n)
        foreach ($a in $n.Aliases) { if ($a) { $matchNames.Add([string]$a); $matchEntries.Add($n) } }
    }
    $script:Index = @{
        Entries = $entries; Notes = $notes; NoteNames = [string[]]($notes | ForEach-Object { $_.Name }); Scan = $scan
        MatchNames = $matchNames.ToArray(); MatchEntries = $matchEntries.ToArray()
        ByKey = $byKey; ByPath = $byPath; NotesByName = $notesByName; FilesByName = $filesByName; Backlinks = @{}; Signature = $signature; Unresolved = $unresolved
    }
    $tPrep = $sw.ElapsedMilliseconds
    if ($script:App.ShowBacklinks) { $script:Index.Backlinks = [CheatSheetCompanionHelper.VaultScanner]::Backlinks($scan) }
    $tLinks = $sw.ElapsedMilliseconds
    if (-not $script:App.SearchContent) { foreach ($n in $notes) { $n.Text = '' } }   # aliases, ignore flag and backlinks are extracted above; the text itself is only needed for the content search
    # parallel arrays for the C# search (display name, aliases, text), same order as Notes
    $searchAliases = [string[][]]::new($notes.Count); $searchTexts = [string[]]::new($notes.Count); $searchDisplay = [string[]]::new($notes.Count)
    for ($i = 0; $i -lt $notes.Count; $i++) {
        $searchDisplay[$i] = [string]$notes[$i].Display
        $searchAliases[$i] = [string[]]@(@($notes[$i].Aliases) | ForEach-Object { [string]$_ } | Where-Object { $_ })
        $searchTexts[$i] = [string]$notes[$i].Text
    }
    $script:Index.SearchDisplay = $searchDisplay; $script:Index.SearchAliases = $searchAliases; $script:Index.SearchTexts = $searchTexts
    $tSearch = $sw.ElapsedMilliseconds

    $converted = 0
    foreach ($n in $notes) {
        # one cheap .NET call per note (a missing file yields year 1601 and therefore converts)
        $needs = $Force -or $script:App.ForceConvert -or ($fileSetChanged -and $unresolved.ContainsKey($n.PathKey)) -or ([System.IO.File]::GetLastWriteTimeUtc($n.Html) -lt $n.LastWrite)
        if ($needs) { if (ConvertTo-NoteHtml -Entry $n) { $converted++ } }
    }
    $tConv = $sw.ElapsedMilliseconds
    try {
        $keep = @($notes | Where-Object { $script:Index.Unresolved.ContainsKey($_.PathKey) } | ForEach-Object { $_.PathKey })
        ConvertTo-Json -InputObject @{ Signature = $signature; Unresolved = $keep } -Compress | Set-Content -LiteralPath $cacheFile -Encoding UTF8
    }
    catch { Write-CSCLog "index.cache: $_" 'Warn' }
    $valid = @{}; foreach ($n in $notes) { $valid[$n.Html.ToLowerInvariant()] = $true }
    Get-ChildItem -LiteralPath $script:App.Output -Recurse -File -Filter '*.html' -ErrorAction SilentlyContinue | Where-Object { -not $valid.ContainsKey($_.FullName.ToLowerInvariant()) } | Remove-Item -Force -ErrorAction SilentlyContinue

    Write-CSCLog ("Index: {0} files, {1} notes ({2} excluded from matching), {3} converted, {4} ms" -f $entries.Count, $notes.Count, @($notes | Where-Object { -not $_.Matchable }).Count, $converted, $sw.ElapsedMilliseconds)
    Write-CSCLog ("Index phases (ms): scan {0}, read {1}, prepare {2}, backlinks {3}, search arrays {4}, convert {5}, cleanup {6}" -f $tScan, ($tRead - $tScan), ($tPrep - $tRead), ($tLinks - $tPrep), ($tSearch - $tLinks), ($tConv - $tSearch), ($sw.ElapsedMilliseconds - $tConv))
    if ($converted -gt 0) { [GC]::Collect() }   # conversion garbage (Markdig, regex) would otherwise linger in an idle process
    $script:App.ForceConvert = $false
    return $fileSetChanged
}

function Get-OriginLabel {
    param([string]$Origin)
    if ($script:Text.ContainsKey('origin.' + $Origin) -or $script:TextDefault.ContainsKey('origin.' + $Origin)) { return (Get-Text ('origin.' + $Origin)) }
    return $Origin
}

function ConvertTo-MatchReason {
    <# Reason code of the matcher -> text (contains, contained, whole, token:<word>) #>
    param([string]$Code, [string]$MatchedName, [string]$NoteName)
    $why = switch -regex ($Code) {
        '^contains$'  { Get-Text 'reason.contains' }
        '^contained$' { Get-Text 'reason.contained' }
        '^whole$'     { Get-Text 'reason.whole' }
        '^token:(.+)$' { Get-Text 'reason.token' -Format $matches[1] }
        default       { '' }
    }
    if ($MatchedName -and $MatchedName -ne $NoteName) { $why = Get-Text 'reason.alias' -Format $MatchedName, $why }
    return $why
}

#endregion

#region ---------- Matching & search ----------

function Find-NoteCandidates {
    <#
        .SYNOPSIS
        Computes Top-N candidates for one or more targets and merges them (dedup by note, max score, merged origins).
        Targets: @(@{ Text='...'; Origin='title|process|clipboard'; Threshold=0.25 })
    #>
    param([Parameter(Mandatory)][object[]]$Targets, [int]$Top = 0)
    if ($Top -le 0) { $Top = $script:App.MaxCandidates }
    $map = [ordered]@{}
    $names = $script:Index.MatchNames
    if ($names.Count -eq 0) { return @() }
    foreach ($t in $Targets) {
        if ([string]::IsNullOrWhiteSpace($t.Text)) { continue }
        $matches = [CheatSheetCompanionHelper.Similarity]::TopMatches($names, [string]$t.Text, $script:App.NGram, [double]$t.Threshold, $Top * 2)
        $weight = if ($t.Origin -eq 'process') { $script:App.ProcessWeight } else { 1.0 }
        foreach ($m in $matches) {
            $note = $script:Index.MatchEntries[$m.Index]
            $score = [Math]::Round([double]$m.Score * $weight, 4)
            $reason = '{0}: {1} ({2})' -f (Get-OriginLabel $t.Origin), (ConvertTo-MatchReason -Code $m.Reason -MatchedName $names[$m.Index] -NoteName $note.Name), $score.ToString('0.00', [cultureinfo]::InvariantCulture)
            if ($map.Contains($note.Key)) {
                $c = $map[$note.Key]
                if ($score -gt $c.Score) { $c.Score = $score; $c.MatchLength = $names[$m.Index].Length }
                elseif ($score -eq $c.Score -and $names[$m.Index].Length -gt $c.MatchLength) { $c.MatchLength = $names[$m.Index].Length }
                if ($c.Origins -notcontains $t.Origin) { $c.Origins += $t.Origin; $c.Reasons += $reason }
            }
            else {
                $map[$note.Key] = [pscustomobject]@{ Entry = $note; Score = $score; Origins = @($t.Origin); Reasons = @($reason); MatchLength = $names[$m.Index].Length }
            }
        }
    }
    # equal scores: the longer matched name (most specific note) first, then alphabetical
    return @($map.Values | Sort-Object -Property @{ Expression = 'Score'; Descending = $true }, @{ Expression = 'MatchLength'; Descending = $true }, @{ Expression = { $_.Entry.Name }; Descending = $false } | Select-Object -First $Top)
}

function Get-RuleCandidates {
    <# [Rules] hits for the active window: process name or title pattern -> note (score 1.00, origin 'rule', shown first). #>
    param([AllowEmptyString()][string]$Title, [AllowEmptyString()][string]$Process)
    $out = @()
    foreach ($r in $script:App.Rules) {
        $hit = if ($r.Kind -eq 'process') { $Process -and ($Process -ieq $r.Process) } else { $Title -and $r.Regex.IsMatch($Title) }
        if (-not $hit) { continue }
        $entry = Resolve-VaultLink -Target $r.Note
        if ($null -eq $entry -or -not $entry.IsNote) { Write-CSCLog ("Rule '{0}': note '{1}' not found in vault" -f $r.Spec, $r.Note) 'Warn'; continue }
        if (@($out | Where-Object { $_.Entry.Key -eq $entry.Key }).Count -gt 0) { continue }
        $out += [pscustomobject]@{ Entry = $entry; Score = 1.0; Origins = @('rule'); Reasons = @((Get-Text 'reason.rule' -Format $r.Spec)) }
    }
    return $out
}

function Find-SearchResults {
    <#
        .SYNOPSIS
        Quick switcher search over note names, aliases and contents, one result per note (its best hit).
        Scores: substring hit in name/alias >= 1000, content hit (all tokens contained) 500, fuzzy subsequence hit below.
        Returns objects: Entry, Kind ('name'|'alias'|'content'), Score, Positions, Detail
    #>
    param([Parameter(Mandatory)][string]$Query)
    $q = $Query.Trim()
    if (-not $q) { return @() }
    $max = $script:App.SearchMaxResults
    if ($null -eq $script:Index.SearchDisplay -or $script:Index.SearchDisplay.Count -eq 0) { return @() }
    # the whole pass runs in C# (thousands of notes per keystroke), PowerShell only wraps the few hits
    $hits = [CheatSheetCompanionHelper.SearchEngine]::Find($script:Index.SearchDisplay, $script:Index.SearchAliases, $script:Index.SearchTexts, $q, $max)
    $kinds = @('name', 'alias', 'content')
    return @($hits | ForEach-Object {
            $n = $script:Index.Notes[$_.Index]
            [pscustomobject]@{ Entry = $n; Kind = $kinds[$_.Kind]; Score = $_.Score; Positions = $_.Positions; Detail = $(if ($_.Kind -eq 0) { $n.Folder } else { $_.Detail }) }
        })
}

#endregion

#region ---------- Themes ----------

function Get-ThemeNames {
    <# auto, light, dark + every other themes\*.css #>
    $names = [System.Collections.Generic.List[string]]@('auto', 'light', 'dark')
    if (Test-Path -LiteralPath $script:App.ThemeDir) {
        foreach ($f in (Get-ChildItem -LiteralPath $script:App.ThemeDir -Filter '*.css' | Sort-Object Name)) {
            $n = $f.BaseName.ToLowerInvariant()
            if ($names -notcontains $n) { $names.Add($n) }
        }
    }
    return $names.ToArray()
}

function Get-ThemeLabel {
    param([string]$Name)
    switch ($Name) { 'auto' { Get-Text 'theme.auto' } 'light' { Get-Text 'theme.light' } 'dark' { Get-Text 'theme.dark' } default { $Name } }
}

function Get-EffectiveTheme {
    <# Resolves 'auto' against the Windows app theme; returns @{ Name; IsDark; CssPath } #>
    $name = [string]$script:State.Theme
    if (-not $name) { $name = 'auto' }
    if ($name -eq 'auto') {
        $name = 'light'
        try {
            $v = Get-ItemPropertyValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name 'AppsUseLightTheme' -ErrorAction Stop
            if ([int]$v -eq 0) { $name = 'dark' }
        }
        catch { }
    }
    $css = $null
    if ($name -ne 'light') {
        $candidate = Join-Path $script:App.ThemeDir ($name + '.css')
        if (Test-Path -LiteralPath $candidate) { $css = $candidate } elseif ($name -ne 'dark') { $name = 'light' }
    }
    return @{ Name = $name; IsDark = ($name -match 'dark'); CssPath = $css }
}

function Set-Theme {
    param([Parameter(Mandatory)][string]$Name)
    $script:State.Theme = $Name
    Update-WindowTheme
    if ($script:UI.CurrentView) {
        $v = $script:UI.CurrentView
        Invoke-ViewNavigation -Html (New-ViewHtml -Candidates $v.Candidates -SourceText $v.SourceText -Anchor $v.Anchor -ActiveIndex ([int]$v.ActiveIndex))
    }
    else { Show-Welcome }
    Save-State
    Set-Status (Get-Text 'status.theme' -Format (Get-ThemeLabel $Name))
    Write-CSCLog "Theme = $Name"
}

function Update-WindowTheme {
    <# Colors of the WPF chrome (toolbar, status bar, popup) follow the effective theme. #>
    $dark = (Get-EffectiveTheme).IsDark
    $palette = if ($dark) {
        @{ CscWindowBg = '#262626'; CscPanelBg = '#333333'; CscFg = '#e6e6e6'; CscMutedFg = '#9b9b9b'; CscBorder = '#555555'; CscAccent = '#a88bfa' }
    }
    else {
        @{ CscWindowBg = '#f0f0f0'; CscPanelBg = '#fafafa'; CscFg = '#1e1e1e'; CscMutedFg = '#6c7076'; CscBorder = '#a0a0a0'; CscAccent = '#7852ee' }
    }
    foreach ($k in $palette.Keys) {
        $brush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($palette[$k]))
        $brush.Freeze()
        $script:UI.Window.Resources[$k] = $brush
    }
    try { if ($script:UI.WebReady) { $script:UI.WebView.DefaultBackgroundColor = [System.Drawing.ColorTranslator]::FromHtml($(if ($dark) { '#1e1e1e' } else { '#ffffff' })) } } catch { }
    try { if ($script:UI.Hwnd -ne [IntPtr]::Zero) { [CheatSheetCompanionHelper.Native]::SetDarkTitleBar($script:UI.Hwnd, $dark) } } catch { }
}

#endregion

#region ---------- View (HTML page) ----------

function Get-NoteViewHtml {
    <# Reads the cached HTML of a note and inlines embedded notes (one level deep). #>
    param([Parameter(Mandatory)]$Entry, [int]$Depth = 0)
    if (-not (Test-Path -LiteralPath $Entry.Html)) { [void](ConvertTo-NoteHtml -Entry $Entry) }
    $html = try { [System.IO.File]::ReadAllText($Entry.Html) } catch { '' }
    return [regex]::Replace($html, '<div class="csc-embed" data-rel="([^"]*)" data-anchor="([^"]*)"></div>', {
            param($m)
            $rel = [System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value)
            $target = $script:Index.ByPath[($rel -replace '\\', '/').ToLowerInvariant()]
            if ($null -eq $target) { return '<span class="csc-missing">![[' + $rel + ']]</span>' }
            $title = '<div class="markdown-embed-title"><a class="internal-link" href="#" data-rel="' + $m.Groups[1].Value + '" data-anchor="' + $m.Groups[2].Value + '">' + (ConvertTo-HtmlAttribute $target.Name) + '</a></div>'
            if ($Depth -ge 1) { return '<div class="markdown-embed">' + $title + '</div>' }
            return '<div class="markdown-embed">' + $title + '<div class="markdown-embed-content">' + (Get-NoteViewHtml -Entry $target -Depth ($Depth + 1)) + '</div></div>'
        })
}

function ConvertTo-BacklinkContext {
    <# One mention line as HTML: list/heading/quote markers stripped, links to notes shown as <mark>label</mark>. #>
    param([AllowEmptyString()][string]$Line)
    $s = $Line -replace '^(#{1,6}\s+|>\s*|[-*+]\s+(\[[ xX]\]\s+)?|\d+[.)]\s+)+', ''
    $s = Limit-Text $s 160
    $rx = [regex]'(!?\[\[(?<t>[^\[\]\|#]*)(?:#(?<a>[^\[\]\|]*))?(?:\\?\|(?<l>[^\[\]]*))?\]\])|(\[(?<ml>[^\]]*)\]\((?<mt>[^)\s]+?\.md)(?:#[^)]*)?\))'
    $out = [System.Text.StringBuilder]::new()
    $pos = 0
    foreach ($m in $rx.Matches($s)) {
        [void]$out.Append((ConvertTo-HtmlAttribute $s.Substring($pos, $m.Index - $pos)))
        $label = if ($m.Groups['l'].Success -and $m.Groups['l'].Value) { $m.Groups['l'].Value }
        elseif ($m.Groups['ml'].Success) { $m.Groups['ml'].Value }
        else { $m.Groups['t'].Value.Trim() + $(if ($m.Groups['a'].Value) { ' > ' + $m.Groups['a'].Value } else { '' }) }
        [void]$out.Append('<mark>' + (ConvertTo-HtmlAttribute $label) + '</mark>')
        $pos = $m.Index + $m.Length
    }
    [void]$out.Append((ConvertTo-HtmlAttribute $s.Substring($pos)))
    return $out.ToString()
}

function Get-BacklinksHtml {
    <# "Linked from" list of a note (Obsidian's linked mentions) as hidden data inside the pane; app.js shows it in the side panel. Empty when nothing links here. #>
    param([Parameter(Mandatory)]$Entry)
    if (-not $script:App.ShowBacklinks) { return '' }
    $map = $script:Index.Backlinks
    if ($null -eq $map -or -not $map.ContainsKey($Entry.Key)) { return '' }
    $sources = @($map[$Entry.Key] | Sort-Object { $_.Source.Display })
    if ($sources.Count -eq 0) { return '' }
    $sb = [System.Text.StringBuilder]::new()
    foreach ($s in $sources) {
        [void]$sb.Append('<li><a class="internal-link" href="#" data-rel="' + (ConvertTo-HtmlAttribute $s.Source.Rel) + '" data-anchor="" title="' + (ConvertTo-HtmlAttribute $s.Source.Rel) + '">' + (ConvertTo-HtmlAttribute $s.Source.Display) + '</a>')
        foreach ($ctx in $s.Contexts) { [void]$sb.Append('<p class="csc-bl-ctx">' + (ConvertTo-BacklinkContext -Line $ctx) + '</p>') }
        [void]$sb.Append('</li>')
    }
    return '<div class="csc-backlinks-data" hidden data-count="' + $sources.Count + '"><ul>' + $sb.ToString() + '</ul></div>'
}

function New-ViewHtml {
    <# Builds the complete page: header with candidate tabs, one pane per candidate, TOC, overlay, backlinks, config + app.js. #>
    param(
        [AllowEmptyCollection()][object[]]$Candidates = @(),   # objects with Entry, Score, Origins
        [string]$SourceText = '',
        [string]$Anchor = '',
        [int]$ActiveIndex = 0,
        [string]$BodyHtml = ''                                  # shown instead of panes when there are no candidates
    )
    $tabs = [System.Text.StringBuilder]::new()
    $panes = [System.Text.StringBuilder]::new()
    $i = 0
    $hasPinned = $false
    foreach ($c in $Candidates) {
        $origins = @($c.Origins)
        $isPinned = ($origins -contains 'pinned')
        if ($isPinned) { $hasPinned = $true }
        $originClass = if ($isPinned) { 'origin-pinned' } elseif ($origins.Count -gt 1) { 'origin-mix' } else { 'origin-' + $origins[0] }
        $scoreClass = if ($c.Score -ge 0.75) { 'score-high' } elseif ($c.Score -ge 0.5) { 'score-mid' } else { 'score-low' }
        $scoreText = $c.Score.ToString('0.00', [cultureinfo]::InvariantCulture)
        $name = ConvertTo-HtmlAttribute $c.Entry.Display
        $rel = ConvertTo-HtmlAttribute $c.Entry.Rel
        $showScore = (-not $isPinned -and $origins -notcontains 'manual' -and $origins -notcontains 'link')
        $tip = $c.Entry.Rel
        if ($isPinned) { $tip += "`n" + (Get-Text 'page.pinnedTab') }
        if ($showScore) { $tip += "`n" + (($origins | ForEach-Object { Get-OriginLabel $_ }) -join ' + ') }
        if ($script:App.ShowMatchReason -and -not $isPinned -and $c.PSObject.Properties['Reasons'] -and $c.Reasons) { $tip += "`n" + (@($c.Reasons) -join "`n") }
        $label = $(if ($isPinned) { '<span class="pin">&#128204;</span>' } else { '' }) + $name
        [void]$tabs.Append('<button class="' + $originClass + ' ' + $scoreClass + '" data-i="' + $i + '" data-rel="' + $rel + '" data-name="' + $name + '" data-score="' + $(if ($showScore) { $scoreText } else { '' }) + '" title="' + (ConvertTo-HtmlAttribute $tip) + '">' + $label + $(if ($showScore) { '<span class="score">' + $scoreText + '</span>' }) + '</button>')
        [void]$panes.Append('<section class="csc-pane" data-i="' + $i + '" data-rel="' + $rel + '"><article class="csc-content">' + (Get-NoteViewHtml -Entry $c.Entry) + '</article>' + (Get-BacklinksHtml -Entry $c.Entry) + '</section>')
        $i++
    }
    $theme = Get-EffectiveTheme
    $themeCss = if ($theme.CssPath) { '<link rel="stylesheet" href="https://app.csc/themes/' + [uri]::EscapeDataString((Split-Path $theme.CssPath -Leaf)) + '">' } else { '' }
    $userCss = if (Test-Path -LiteralPath (Join-Path $script:App.Root 'styles\userstyle.css')) { '<link rel="stylesheet" href="https://app.csc/styles/userstyle.css">' } else { '' }
    $hljs = ''
    if (Test-Path -LiteralPath (Join-Path $script:App.Root 'assets\highlight.min.js')) {
        $hlCss = if ($theme.IsDark -and (Test-Path -LiteralPath (Join-Path $script:App.Root 'assets\highlight-dark.css'))) { 'highlight-dark.css' } else { 'highlight.css' }
        if (Test-Path -LiteralPath (Join-Path $script:App.Root "assets\$hlCss")) {
            $hljs = '<link rel="stylesheet" href="https://app.csc/assets/' + $hlCss + '"><script src="https://app.csc/assets/highlight.min.js"></script><script>try{hljs.highlightAll();}catch(e){}</script>'
        }
    }
    $cfg = @{
        autoCopy = [bool]$script:State.AutoCopy
        autoCopyMinHoldMs = [int]$script:App.AutoCopyMinHoldMs
        dblClickEdit = [string]$script:App.EditOnDoubleClick
        anchor = $Anchor
        activeIndex = $ActiveIndex
        text = @{ copy = (Get-Text 'page.copy'); copied = (Get-Text 'page.copied'); copyCode = (Get-Text 'page.copyCode'); copyAgain = (Get-Text 'page.copyAgain'); dataviewOnly = (Get-Text 'page.dataviewOnly'); dataviewFailed = (Get-Text 'page.dataviewFailed') }
        copyHistory = @(Get-CopyHistoryItems)
        linkPreview = [bool]$script:App.LinkPreview
    } | ConvertTo-Json -Compress -Depth 4
    $legend = '<div class="csc-legend">' + $(if ($hasPinned) { '<span><span class="dot pinned"></span>' + (ConvertTo-HtmlAttribute (Get-Text 'origin.pinned')) + '</span>' }) +
    '<span><span class="dot title"></span>' + (ConvertTo-HtmlAttribute (Get-Text 'origin.title')) + '</span><span><span class="dot process"></span>' + (ConvertTo-HtmlAttribute (Get-Text 'origin.process')) + '</span><span><span class="dot clipboard"></span>' + (ConvertTo-HtmlAttribute (Get-Text 'origin.clipboard')) + '</span></div>'

    return @"
<!DOCTYPE html>
<html lang="$($script:App.Language)" data-theme="$($theme.Name)">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>CheatSheetCompanion</title>
<link rel="stylesheet" href="https://app.csc/assets/app.css">
$themeCss
$userCss
</head>
<body class="markdown-body">
<header id="csc-header">
  <div class="csc-row"><div id="csc-tabs" class="csc-tabs">$($tabs.ToString())</div><div class="csc-source" title="$(ConvertTo-HtmlAttribute $SourceText)">$(ConvertTo-HtmlAttribute $SourceText)</div></div>
  $(if ($Candidates.Count -gt 1) { $legend })
</header>
<main id="csc-main">$(if ($Candidates.Count -gt 0) { $panes.ToString() } else { $BodyHtml })</main>
<button id="csc-toc-handle" title="$(ConvertTo-HtmlAttribute (Get-Text 'page.toc'))"><span>TOC</span></button>
<aside id="csc-toc-panel"><h3>$(ConvertTo-HtmlAttribute (Get-Text 'page.tocTitle'))</h3><nav id="csc-toc"></nav></aside>
<button id="csc-copy-handle" title="$(ConvertTo-HtmlAttribute (Get-Text 'page.copyHistory'))"><span>$(ConvertTo-HtmlAttribute (Get-Text 'page.copyHandle'))</span></button>
<aside id="csc-copy-panel"><h3>$(ConvertTo-HtmlAttribute (Get-Text 'page.copyHistory'))<button type="button" id="csc-copy-clear" class="csc-panel-btn" title="$(ConvertTo-HtmlAttribute (Get-Text 'page.copyClearTip'))">$(ConvertTo-HtmlAttribute (Get-Text 'page.copyClear'))</button></h3><div id="csc-copy-list"></div></aside>
<button id="csc-links-handle" title="$(ConvertTo-HtmlAttribute (Get-Text 'page.backlinksTip'))"><span>$(ConvertTo-HtmlAttribute (Get-Text 'page.linksHandle'))</span></button>
<aside id="csc-links-panel"><h3>$(ConvertTo-HtmlAttribute (Get-Text 'page.backlinks')) <span class="count" id="csc-links-count"></span></h3><div id="csc-links-list"></div></aside>
<script>window.CSC_CONFIG = $cfg;</script>
<script src="https://app.csc/assets/app.js"></script>
$hljs
</body>
</html>
"@
}

function Invoke-ViewNavigation {
    <# Writes the page and navigates the WebView2 to it (or queues it until the core is ready). #>
    param([Parameter(Mandatory)][string]$Html)
    try {
        $script:UI.LastSelection = @{ Text = ''; Time = [DateTime]::MinValue }
        [System.IO.File]::WriteAllText($script:App.ViewFile, $Html, [System.Text.UTF8Encoding]::new($false))
        if ($script:UI.WebReady) {
            $script:UI.WebView.CoreWebView2.Navigate('https://app.csc/Output/_view.html?v=' + [DateTime]::UtcNow.Ticks)
        }
        else { $script:UI.PendingNavigation = $true }
    }
    catch { Write-CSCLog "Navigation failed: $_" 'Error' }
}

function Show-Candidates {
    <#
        Shows candidate tabs. Pinned notes always come first; the visible tab is the best match (first non-pinned
        candidate) or the one given with -ActiveEntry (search, links, tab clicks).
    #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Candidates, [string]$SourceText = '', [string]$Anchor = '', [switch]$NoHistory, $ActiveEntry = $null)
    $Candidates = @(Join-PinnedCandidates -Candidates $Candidates)
    if ($Candidates.Count -eq 0) { return }
    $active = -1
    if ($null -ne $ActiveEntry) {
        for ($i = 0; $i -lt $Candidates.Count; $i++) { if ($Candidates[$i].Entry.Key -eq $ActiveEntry.Key) { $active = $i; break } }
    }
    if ($active -lt 0) {
        for ($i = 0; $i -lt $Candidates.Count; $i++) { if (@($Candidates[$i].Origins) -notcontains 'pinned') { $active = $i; break } }
        if ($active -lt 0) { $active = 0 }
    }
    if (-not $NoHistory -and $script:UI.CurrentView) {
        [void]$script:UI.History.Push($script:UI.CurrentView); $script:UI.Forward.Clear()
        if ($script:UI.History.Count -gt 60) {   # bounded: a day of window switches must not pile up views
            $keep = @($script:UI.History.ToArray() | Select-Object -First 40)
            $script:UI.History.Clear(); for ($k = $keep.Count - 1; $k -ge 0; $k--) { [void]$script:UI.History.Push($keep[$k]) }
        }
    }
    $script:UI.CurrentView = @{ Candidates = $Candidates; SourceText = $SourceText; Anchor = $Anchor; ActiveIndex = $active }
    $script:UI.CurrentNote = $Candidates[$active].Entry
    Add-RecentNote -Entry $Candidates[$active].Entry
    Invoke-ViewNavigation -Html (New-ViewHtml -Candidates $Candidates -SourceText $SourceText -Anchor $Anchor -ActiveIndex $active)
    Sync-SearchBox -Entry $Candidates[$active].Entry
    Update-StripText
    Update-PinButton
}

function Get-PinnedEntries {
    <# Index entries of the pinned notes of the active vault (missing notes are skipped). #>
    $out = @()
    $prefix = $script:App.VaultName + '|'
    foreach ($key in @($script:State.PinnedNotes)) {
        $k = [string]$key
        if (-not $k.StartsWith($prefix)) { continue }
        $e = $script:Index.ByPath[($k.Substring($prefix.Length) -replace '\\', '/').ToLowerInvariant()]
        if ($e) { $out += $e }
    }
    return $out
}

function Test-NotePinned {
    param($Entry)
    if ($null -eq $Entry) { return $false }
    return (@($script:State.PinnedNotes) -contains ('{0}|{1}' -f $script:App.VaultName, $Entry.Rel))
}

function Update-PinButton {
    <# Pin button, tray item: checked when the visible note is pinned. #>
    $pinned = Test-NotePinned -Entry $script:UI.CurrentNote
    $script:UI.SuppressToggleEvents = $true
    try {
        $script:UI.TogglePinned.IsChecked = $pinned
        if ($script:UI.TrayItems -and $script:UI.TrayItems['Pinned']) { $script:UI.TrayItems['Pinned'].Checked = $pinned }
    }
    catch { } finally { $script:UI.SuppressToggleEvents = $false }
}

function Join-PinnedCandidates {
    <# Pinned notes first (origin 'pinned', no score), then the other candidates; MaxCandidates in total, at least one slot for matches. #>
    param([AllowEmptyCollection()][object[]]$Candidates)
    $rest = @($Candidates | Where-Object { @($_.Origins) -notcontains 'pinned' })
    $pins = @(Get-PinnedEntries)
    if ($pins.Count -eq 0) { return $rest }
    $keys = @{}
    $pinned = @(foreach ($p in $pins) { $keys[$p.Key] = $true; [pscustomobject]@{ Entry = $p; Score = 1.0; Origins = @('pinned'); Reasons = @((Get-Text 'reason.pinned')) } })
    $limit = [Math]::Max($pins.Count + 1, $script:App.MaxCandidates)
    return @(@($pinned + @($rest | Where-Object { -not $keys.ContainsKey($_.Entry.Key) })) | Select-Object -First $limit)
}

function Update-StripText {
    <# The collapsed strip shows the current note (name + score) so the strip is useful without expanding. #>
    try {
        $collapsed = $script:UI.Window.Height -lt 20
        $glyph = [string][char]::ConvertFromUtf32($(if ($collapsed) { 0x1F53B } else { 0x1F53A })) * 3
        $note = ''
        if ($collapsed -and $script:UI.CurrentNote) {
            $c = if ($script:UI.CurrentView) { $script:UI.CurrentView.Candidates[[int]$script:UI.CurrentView.ActiveIndex] } else { $null }
            $score = if ($c -and $c.Origins -notcontains 'manual' -and $c.Origins -notcontains 'link' -and $c.Origins -notcontains 'pinned') { ' (' + $c.Score.ToString('0.00', [cultureinfo]::InvariantCulture) + ')' } else { '' }
            $note = '   ' + (Limit-Text $script:UI.CurrentNote.Display 70) + $score + $(if (Test-NotePinned -Entry $script:UI.CurrentNote) { '  ' + (Get-Text 'strip.pinned') } else { '' })
        }
        $script:UI.MinimizeText.Text = $glyph + $note
    }
    catch { }
}

function Add-CopyHistory {
    <# Remembers a text copied through the companion (code block, selection, context menu) for re-copying. #>
    param([Parameter(Mandatory)][string]$Text)
    if ($script:App.CopyHistoryMax -le 0) { return }
    $t = $Text.Trim()
    if (-not $t) { return }
    if ($t.Length -gt 2000) { $t = $t.Substring(0, 2000) }
    $note = if ($script:UI.CurrentNote) { [string]$script:UI.CurrentNote.Display } else { '' }
    $item = @{ Text = $t; Note = $note; Time = [DateTime]::Now.ToString('s') }
    $list = @($item) + @(@($script:State.CopyHistory) | Where-Object { $null -ne $_ -and ([string]$_.Text) -ne $t })
    $script:State.CopyHistory = @($list | Select-Object -First $script:App.CopyHistoryMax)
}

function Get-CopyHistoryLabel {
    <# First line of a copied text, shortened, for menus and lists. #>
    param([AllowNull()][string]$Text, [int]$Max = 60)
    $first = [string](([string]$Text -split '\r?\n' | Where-Object { $_.Trim() } | Select-Object -First 1))
    return Limit-Text $first $Max
}

function Get-CopyHistoryItems {
    <# Items for the slide-out panel in the page: label (first line), full text (shortened), source note. #>
    if ($script:App.CopyHistoryMax -le 0) { return @() }
    $items = @()
    foreach ($h in @($script:State.CopyHistory)) {
        if ($null -eq $h -or -not $h.Text) { continue }
        $items += @{ label = (Get-CopyHistoryLabel -Text ([string]$h.Text) -Max 60); text = (Limit-Text ([string]$h.Text) 300); note = [string]$h.Note }
    }
    return $items
}

function Update-CopyHistoryPanel {
    <# Pushes the current copy history into the shown page (no reload). #>
    try {
        if (-not $script:UI.WebReady) { return }
        $json = ConvertTo-Json -InputObject @(Get-CopyHistoryItems) -Compress -Depth 3
        if ($json -eq 'null') { $json = '[]' }
        $null = $script:UI.WebView.CoreWebView2.ExecuteScriptAsync('if(window.CSC_setCopyHistory){window.CSC_setCopyHistory(' + $json + ');}')
    }
    catch { }
}

function Add-RecentNote {
    <# Remembers shown notes ("vault|relative path", newest first) for the start page. #>
    param([Parameter(Mandatory)]$Entry)
    if ($script:App.RecentMax -le 0) { return }
    $key = '{0}|{1}' -f $script:App.VaultName, $Entry.Rel
    $list = @($key) + @(@($script:State.RecentNotes) | Where-Object { $_ -and $_ -ne $key })
    $script:State.RecentNotes = @($list | Select-Object -First ([Math]::Max(20, $script:App.RecentMax * 2)))
}

function Get-RecentNotesHtml {
    <# List of the recently shown notes of the active vault (internal links), '' when there are none. #>
    if ($script:App.RecentMax -le 0) { return '' }
    $prefix = $script:App.VaultName + '|'
    $items = @()
    foreach ($k in @($script:State.RecentNotes)) {
        $k = [string]$k
        if (-not $k.StartsWith($prefix)) { continue }
        $e = $script:Index.ByPath[($k.Substring($prefix.Length) -replace '\\', '/').ToLowerInvariant()]
        if ($null -eq $e) { continue }
        $items += '<li><a class="internal-link" href="#" data-rel="' + (ConvertTo-HtmlAttribute $e.Rel) + '" data-anchor="">' + (ConvertTo-HtmlAttribute $e.Display) + '</a></li>'
        if ($items.Count -ge $script:App.RecentMax) { break }
    }
    if ($items.Count -eq 0) { return '' }
    return '<p class="csc-near-title">' + (ConvertTo-HtmlAttribute (Get-Text 'page.recent')) + '</p><ul class="csc-near">' + ($items -join '') + '</ul>'
}

function Show-Note {
    param([Parameter(Mandatory)]$Entry, [string]$Origin = 'manual', [string]$Anchor = '', [switch]$NoHistory)
    $cand = @([pscustomobject]@{ Entry = $Entry; Score = 1.0; Origins = @($Origin) })
    Show-Candidates -Candidates $cand -SourceText $Entry.Rel -Anchor $Anchor -NoHistory:$NoHistory -ActiveEntry $Entry
    Set-Status (Get-Text 'status.note' -Format $Entry.Display)
}

function Show-Welcome {
    $clip = if ($script:State.FollowClipboard) { Get-Text 'page.welcomeClipboard' } else { '' }
    $body = '<div class="csc-empty"><h2>Cheat Sheet Companion</h2><p>' + (Get-Text 'page.welcome' -Format $clip, (ConvertTo-HtmlAttribute $script:App.VaultName), $script:Index.Notes.Count) + '</p>' + (Get-RecentNotesHtml) + '</div>'
    Invoke-ViewNavigation -Html (New-ViewHtml -Candidates @() -BodyHtml $body)
}

function Show-NoMatch {
    <# "No note yet" page: offers to create a note and lists the closest notes even below the threshold. #>
    param([Parameter(Mandatory)][object[]]$Targets, [Parameter(Mandatory)][string]$SourceText)
    $loose = @($Targets | ForEach-Object { @{ Text = $_.Text; Origin = $_.Origin; Threshold = 0.05 } })
    $near = Find-NoteCandidates -Targets $loose -Top 5
    $list = ''
    if ($near.Count -gt 0) {
        $list = '<p class="csc-near-title">' + (ConvertTo-HtmlAttribute (Get-Text 'page.nearest')) + '</p><ul class="csc-near">' + (($near | ForEach-Object {
                    '<li><a class="internal-link" href="#" data-rel="' + (ConvertTo-HtmlAttribute $_.Entry.Rel) + '" data-anchor="">' + (ConvertTo-HtmlAttribute $_.Entry.Display) + '</a> <span class="score">' + $_.Score.ToString('0.00', [cultureinfo]::InvariantCulture) + '</span></li>'
                }) -join '') + '</ul>'
    }
    $body = '<div class="csc-empty csc-nomatch"><p class="csc-nomatch-source">' + (ConvertTo-HtmlAttribute (Limit-Text $SourceText 120)) + '</p><h2>' + (ConvertTo-HtmlAttribute (Get-Text 'page.noMatch')) + '</h2>' +
    '<p><button type="button" class="csc-primary" data-action="newNote">' + (ConvertTo-HtmlAttribute (Get-Text 'page.createNote')) + '</button></p>' + $list + '</div>'
    if ($script:UI.CurrentView) { [void]$script:UI.History.Push($script:UI.CurrentView) }
    $script:UI.CurrentView = $null
    $script:UI.CurrentNote = $null
    if (-not $script:UI.SearchPopup.IsOpen) {
        $script:UI.SuppressSearchEvents = $true
        try { $script:UI.SearchBox.Text = ''; $script:UI.SearchHint.Visibility = 'Visible' } catch { } finally { $script:UI.SuppressSearchEvents = $false }
    }
    Invoke-ViewNavigation -Html (New-ViewHtml -Candidates @() -SourceText $SourceText -BodyHtml $body)
}

function Restore-View {
    <# Shows a view from the back / forward stack again. #>
    param([Parameter(Mandatory)]$View)
    $active = [int]$View.ActiveIndex
    if ($active -lt 0 -or $active -ge $View.Candidates.Count) { $active = 0 }
    $script:UI.CurrentView = $View
    $script:UI.CurrentNote = $View.Candidates[$active].Entry
    Invoke-ViewNavigation -Html (New-ViewHtml -Candidates $View.Candidates -SourceText $View.SourceText -Anchor $View.Anchor -ActiveIndex $active)
    Sync-SearchBox -Entry $View.Candidates[$active].Entry
    Update-StripText
    Update-PinButton
}

function Invoke-ViewBack {
    if ($script:UI.History.Count -eq 0) { return }
    if ($script:UI.CurrentView) { [void]$script:UI.Forward.Push($script:UI.CurrentView) }
    Restore-View -View $script:UI.History.Pop()
}

function Invoke-ViewForward {
    if ($script:UI.Forward.Count -eq 0) { return }
    if ($script:UI.CurrentView) { [void]$script:UI.History.Push($script:UI.CurrentView) }
    Restore-View -View $script:UI.Forward.Pop()
}

function Show-MatchesFor {
    <#
        Runs the matcher for the given targets and shows the result. -Pinned (rule hits) come first.
        No match: NoMatchBehavior = keep (current note stays, status line only) or suggest ("no note yet" page).
    #>
    param([Parameter(Mandatory)][object[]]$Targets, [Parameter(Mandatory)][string]$SourceText, [string]$StatusPrefix = '', [AllowEmptyCollection()][object[]]$Pinned = @())
    $cands = @(Find-NoteCandidates -Targets $Targets)
    if ($Pinned.Count -gt 0) {
        $keys = @{}; foreach ($p in $Pinned) { $keys[$p.Entry.Key] = $true }
        $cands = @((@($Pinned) + @($cands | Where-Object { -not $keys.ContainsKey($_.Entry.Key) })) | Select-Object -First $script:App.MaxCandidates)
    }
    if ($cands.Count -gt 0) {
        # pinned notes stay as tabs (Join-PinnedCandidates), the best match becomes the visible tab
        Show-Candidates -Candidates $cands -SourceText $SourceText -ActiveEntry $cands[0].Entry
        $why = if ($script:App.ShowMatchReason -and $cands[0].Reasons) { ', ' + ((@($cands[0].Reasons) | ForEach-Object { $_ -replace ' \(\d\.\d\d\)$', '' }) -join '; ') } else { '' }
        Set-Status (Get-Text 'status.match' -Format $StatusPrefix, (Limit-Text $SourceText 60), $cands[0].Entry.Display, $cands[0].Score.ToString('0.00', [cultureinfo]::InvariantCulture), $why)
        Write-CSCLog ("Match '{0}': {1}" -f (Limit-Text $SourceText 80), (($cands | ForEach-Object { '{0}({1:0.00},{2})' -f $_.Entry.Name, $_.Score, ($_.Origins -join '+') }) -join ', '))
    }
    elseif ($script:App.NoMatchBehavior -eq 'suggest' -and @(Get-PinnedEntries).Count -eq 0) {
        Show-NoMatch -Targets $Targets -SourceText $SourceText
        Set-Status (Get-Text 'status.noMatchSuggest' -Format $StatusPrefix, (Limit-Text $SourceText 60))
    }
    else {
        Write-CSCLog ("No match '{0}' (view kept)" -f (Limit-Text $SourceText 80))
        Set-Status (Get-Text 'status.noMatchKeep' -Format $StatusPrefix, (Limit-Text $SourceText 60))
    }
}

function Limit-Text {
    param([AllowNull()][string]$Text, [int]$Max = 60)
    $t = ([string]$Text) -replace '\s+', ' '
    if ($t.Length -le $Max) { return $t }
    return $t.Substring(0, $Max - 1) + '…'
}

#endregion

#region ---------- Actions ----------

function Set-ClipboardTextSafe {
    param([Parameter(Mandatory)][string]$Text)
    $script:UI.LastOwnClipboardText = $Text
    $script:UI.LastOwnClipboardTime = [DateTime]::UtcNow
    for ($i = 0; $i -lt 5; $i++) {
        try { [System.Windows.Clipboard]::SetDataObject($Text, $true); Add-CopyHistory -Text $Text; Update-CopyHistoryPanel; return $true } catch { Start-Sleep -Milliseconds 40 }
    }
    Write-CSCLog 'Clipboard is locked by another application.' 'Warn'
    return $false
}

function ConvertTo-ClipboardHtml {
    <# Wraps an HTML fragment in the CF_HTML envelope (byte offsets in UTF-8) that Teams, Outlook and Word expect. #>
    param([AllowEmptyString()][string]$Fragment)
    $head = "Version:0.9`r`nStartHTML:{0:D10}`r`nEndHTML:{1:D10}`r`nStartFragment:{2:D10}`r`nEndFragment:{3:D10}`r`n"
    $pre = '<html><body><!--StartFragment-->'
    $post = '<!--EndFragment--></body></html>'
    $enc = [System.Text.Encoding]::UTF8
    $startHtml = $enc.GetByteCount(($head -f 0, 0, 0, 0))
    $startFrag = $startHtml + $enc.GetByteCount($pre)
    $endFrag = $startFrag + $enc.GetByteCount($Fragment)
    $endHtml = $endFrag + $enc.GetByteCount($post)
    return ($head -f $startHtml, $endHtml, $startFrag, $endFrag) + $pre + $Fragment + $post
}

function ConvertTo-PortableHtml {
    <# Inline styles instead of the page CSS so lists, code blocks and headings survive pasting into Teams or Outlook. #>
    param([AllowEmptyString()][string]$Html)
    $h = $Html
    $h = [regex]::Replace($h, '<a class="internal-link"[^>]*>(.*?)</a>', '$1')                       # note links mean nothing outside
    $h = [regex]::Replace($h, '<button[^>]*>.*?</button>', '')                                         # copy buttons
    $h = [regex]::Replace($h, '<pre(?=[\s>])', '<pre style="font-family:Consolas,''Courier New'',monospace;font-size:10pt;background:#f3f3f3;border:1px solid #dddddd;padding:6px 8px;white-space:pre-wrap"')
    $h = [regex]::Replace($h, '<code(?=[\s>])', '<code style="font-family:Consolas,''Courier New'',monospace;font-size:10pt"')
    $h = [regex]::Replace($h, '<(h[1-6])(?=[\s>])', '<$1 style="font-family:Segoe UI,Arial,sans-serif;margin:10pt 0 4pt"')
    $h = [regex]::Replace($h, '<(table)(?=[\s>])', '<$1 style="border-collapse:collapse"')
    $h = [regex]::Replace($h, '<(td|th)(?=[\s>])', '<$1 style="border:1px solid #bbbbbb;padding:3px 8px"')
    return '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:11pt;color:#000000">' + $h + '</div>'
}

function Set-ClipboardHtmlSafe {
    <# Puts HTML (CF_HTML) and the plain text version on the clipboard together; the receiver picks what it understands. #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Html, [Parameter(Mandatory)][AllowEmptyString()][string]$Text)
    $script:UI.LastOwnClipboardText = $Text
    $script:UI.LastOwnClipboardTime = [DateTime]::UtcNow
    $cf = ConvertTo-ClipboardHtml -Fragment (ConvertTo-PortableHtml -Html $Html)
    for ($i = 0; $i -lt 5; $i++) {
        try {
            $d = [System.Windows.DataObject]::new()
            $d.SetText($cf, [System.Windows.TextDataFormat]::Html)
            $d.SetText($Text, [System.Windows.TextDataFormat]::UnicodeText)
            [System.Windows.Clipboard]::SetDataObject($d, $true)
            Add-CopyHistory -Text $Text; Update-CopyHistoryPanel
            return $true
        }
        catch { Start-Sleep -Milliseconds 40 }
    }
    Write-CSCLog 'Clipboard is locked by another application.' 'Warn'
    return $false
}

function Invoke-WebScriptThen {
    <# Runs JavaScript in the page and hands the JSON result to -Then once the task completes (polled, UI thread never blocks). #>
    param([Parameter(Mandatory)][string]$Script, [Parameter(Mandatory)][scriptblock]$Then, $Context = $null)
    if (-not $script:UI.WebReady) { return }
    try { $task = $script:UI.WebView.CoreWebView2.ExecuteScriptAsync($Script) } catch { Write-CSCLog "Script: $_" 'Warn'; return }
    $poll = [System.Windows.Threading.DispatcherTimer]::new()
    $poll.Interval = [TimeSpan]::FromMilliseconds(30)
    $poll.Tag = @{ Task = $task; Then = $Then; Context = $Context; Started = [DateTime]::UtcNow }
    $poll.Add_Tick({
            param($s, $e)
            if (-not $s.Tag.Task.IsCompleted) { if (([DateTime]::UtcNow - $s.Tag.Started).TotalSeconds -gt 5) { $s.Stop() }; return }
            $s.Stop()
            try { & $s.Tag.Then $s.Tag.Task.Result $s.Tag.Context } catch { Write-CSCLog "Script result: $_" 'Warn' }   # no closures: script functions must stay visible
        })
    $poll.Start()
}

$script:JsPageContent = @'
(function () {
  try {
    var sel = window.getSelection();
    var whole = window.CSC_COPY_WHOLE === true;
    document.body.classList.add('csc-copying');   /* hides the copy buttons so innerText / selection text skip them */
    try {
      if (!whole && sel && sel.rangeCount > 0 && !sel.isCollapsed) {
        var div = document.createElement('div');
        div.appendChild(sel.getRangeAt(0).cloneContents());
        div.querySelectorAll('button').forEach(function (n) { n.remove(); });
        return { html: div.innerHTML, text: sel.toString(), whole: false };
      }
      var art = document.querySelector('.csc-pane.active .csc-content') || document.querySelector('.csc-content');
      if (!art) { return { html: '', text: '', whole: true }; }
      var c = art.cloneNode(true);
      c.querySelectorAll('button, .csc-backlinks-data').forEach(function (n) { n.remove(); });
      return { html: c.innerHTML, text: art.innerText, whole: true };
    } finally { document.body.classList.remove('csc-copying'); }
  } catch (e) { return { html: '', text: '', whole: false }; }
})();
'@

function Copy-NoteContent {
    <#
        Copies the selection, or the whole shown note without one (-WholeNote forces that), from the page.
        Formatted = HTML plus plain text so Teams and Outlook keep lists and code blocks; -Plain = text only.
    #>
    param([switch]$Plain, [switch]$WholeNote)
    if (-not $script:UI.WebReady) { return }
    $js = $(if ($WholeNote) { 'window.CSC_COPY_WHOLE=true;' } else { 'window.CSC_COPY_WHOLE=false;' }) + $script:JsPageContent
    Invoke-WebScriptThen -Script $js -Context @{ Plain = [bool]$Plain } -Then {
        param($json, $ctx)
        $data = try { $json | ConvertFrom-Json } catch { $null }
        if ($null -eq $data -or [string]::IsNullOrWhiteSpace([string]$data.text)) { Set-Status (Get-Text 'status.noNote'); return }
        $text = ([string]$data.text).Replace("`r`n", "`n").Replace("`n", "`r`n").Trim()
        if ($ctx.Plain) {
            if (Set-ClipboardTextSafe -Text $text) { Set-Status (Get-Text $(if ($data.whole) { 'status.noteCopied' } else { 'status.selectionCopied' }) -Format $text.Length) }
        }
        elseif (Set-ClipboardHtmlSafe -Html ([string]$data.html) -Text $text) { Set-Status (Get-Text 'status.copiedFormatted' -Format $text.Length) }
    }
}

function Add-TextToNote {
    <# Appends text as it is to the end of a note (separated by a blank line); the vault watcher re-renders the note. #>
    param([Parameter(Mandatory)]$Entry, [Parameter(Mandatory)][string]$Text)
    $t = $Text.Replace("`r`n", "`n").Trim("`r", "`n")
    if ([string]::IsNullOrWhiteSpace($t)) { return $false }
    try {
        $existing = [System.IO.File]::ReadAllText($Entry.Full)
        $nl = if ($existing.Contains("`r`n") -or -not $existing.Contains("`n")) { "`r`n" } else { "`n" }
        $body = $t.Replace("`n", $nl)
        $head = $existing.TrimEnd("`r", "`n")
        $new = $(if ($head.Length -gt 0) { $head + $nl + $nl } else { '' }) + $body + $nl
        [System.IO.File]::WriteAllText($Entry.Full, $new, [System.Text.UTF8Encoding]::new($false))
        Write-CSCLog ("Appended {0} chars to '{1}'" -f $t.Length, $Entry.Rel)
        return $true
    }
    catch { Write-CSCLog "Append failed: $_" 'Error'; Set-Status (Get-Text 'status.openFailed' -Format $_); return $false }
}

function Get-ClipboardTextSafe {
    <# '' = no text on the clipboard, $null = clipboard locked by another process (caller retries later). #>
    for ($i = 0; $i -lt 3; $i++) {
        try {
            if (-not [System.Windows.Clipboard]::ContainsText()) { return '' }
            return [System.Windows.Clipboard]::GetText()
        }
        catch { Start-Sleep -Milliseconds 30 }
    }
    return $null
}

function Invoke-ClipboardLookup {
    <#
        Reads the clipboard and shows matching notes. Called on WM_CLIPBOARDUPDATE (matching only with FollowClipboard),
        from the ClipboardLookup hotkey / context menu (-Force) and from the retry timer while the clipboard is locked
        (clipboard history and other listeners open it right after a copy, so the first read often fails).
    #>
    param([switch]$Force, [int]$Attempt = 0)
    $text = Get-ClipboardTextSafe
    if ($null -eq $text) {
        if ($Attempt -ge 5) { Write-CSCLog 'Clipboard stayed locked, giving up.' 'Warn'; return }
        $retry = [System.Windows.Threading.DispatcherTimer]::new()
        $retry.Interval = [TimeSpan]::FromMilliseconds(100 * ($Attempt + 1))
        $retry.Tag = @{ Force = [bool]$Force; Attempt = $Attempt + 1 }
        $retry.Add_Tick({ param($s, $e) $s.Stop(); try { Invoke-ClipboardLookup -Force:$s.Tag.Force -Attempt $s.Tag.Attempt } catch { Write-CSCLog "Clipboard retry: $_" 'Warn' } })
        $retry.Start()
        return
    }
    if ([string]::IsNullOrWhiteSpace($text)) { if ($Force) { Set-Status (Get-Text 'status.clipboardEmpty') }; return }
    $now = [DateTime]::UtcNow
    if (-not $Force) {
        # own copies (code block, selection) and duplicate notifications for a single copy
        if ($text -eq $script:UI.LastOwnClipboardText -and ($now - $script:UI.LastOwnClipboardTime).TotalSeconds -lt 3) { return }
        if ($text -eq $script:UI.LastClipboard.Text -and ($now - $script:UI.LastClipboard.Time).TotalSeconds -lt 1.5) { return }
    }
    $script:UI.LastClipboard = @{ Text = $text; Time = $now }
    if (-not $Force -and -not $script:State.FollowClipboard) { return }
    $sample = ($text -replace '\s+', ' ').Trim()
    if ($sample.Length -gt $script:App.ClipboardMaxChars) { $sample = $sample.Substring(0, $script:App.ClipboardMaxChars) }
    if (-not $Force -and $script:UI.CurrentView -and $script:UI.CurrentView.SourceText -eq $sample) { return }   # already showing this text
    try { $script:UI.ClipboardProcessId = [CheatSheetCompanionHelper.Native]::GetWindowProcessId([CheatSheetCompanionHelper.Native]::GetForegroundWindow()) } catch { $script:UI.ClipboardProcessId = 0 }
    Write-CSCLog ("Clipboard: '{0}'" -f (Limit-Text $sample 80))
    Show-MatchesFor -Targets @(@{ Text = $sample; Origin = 'clipboard'; Threshold = $script:App.ThresholdClipboard }) -SourceText $sample -StatusPrefix (Get-Text 'status.prefixClipboard')
}

function Add-NoteAlias {
    <# Adds an alias to the YAML front matter of a note (creates the front matter / the aliases list when needed). #>
    param([Parameter(Mandatory)]$Entry, [Parameter(Mandatory)][string]$Alias)
    $alias = ($Alias -replace '\s+', ' ').Trim()
    if (-not $alias) { return $false }
    $text = [System.IO.File]::ReadAllText($Entry.Full)
    if (@(Get-FrontmatterList -Markdown $text -Key 'aliases') -contains $alias) { Set-Status (Get-Text 'status.aliasExists' -Format $alias); return $false }
    $needsQuotes = ($alias -match '[:#\[\]{},&*?|<>=!%@`"''\\]') -or ($alias -match '^(true|false|null|yes|no|~)$') -or ($alias -match '^[\d.+-]')
    $yaml = if ($needsQuotes) { '"' + (($alias -replace '\\', '\\') -replace '"', '\"') + '"' } else { $alias }
    $nl = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }
    $m = [regex]::Match($text, '(?s)\A(---\r?\n)(.*?)(\r?\n---[ \t]*(?:\r?\n|\z))')
    if (-not $m.Success) {
        $new = '---' + $nl + 'aliases:' + $nl + '  - ' + $yaml + $nl + '---' + $nl + $text
    }
    else {
        $lines = [System.Collections.Generic.List[string]]@($m.Groups[2].Value -split '\r?\n')
        $idx = -1; $inline = ''
        for ($i = 0; $i -lt $lines.Count; $i++) { if ($lines[$i] -match '^aliases\s*:\s*(.*)$') { $idx = $i; $inline = $matches[1].Trim(); break } }
        if ($idx -lt 0) { $lines.Add('aliases:'); $lines.Add('  - ' + $yaml) }
        elseif ($inline -match '^\[(.*)\]$') {
            $inner = $matches[1].Trim()
            $lines[$idx] = 'aliases: [' + $(if ($inner) { $inner + ', ' } else { '' }) + $yaml + ']'
        }
        elseif ($inline) { $lines[$idx] = 'aliases:'; $lines.Insert($idx + 1, '  - ' + $inline); $lines.Insert($idx + 2, '  - ' + $yaml) }
        else {
            $indent = '  '; $j = $idx + 1
            if ($j -lt $lines.Count -and $lines[$j] -match '^(\s*)-\s') { $indent = $matches[1] }
            while ($j -lt $lines.Count -and $lines[$j] -match '^\s*-\s') { $j++ }
            $lines.Insert($j, $indent + '- ' + $yaml)
        }
        $new = $m.Groups[1].Value + ($lines -join $nl) + $m.Groups[3].Value + $text.Substring($m.Length)
    }
    [System.IO.File]::WriteAllText($Entry.Full, $new, [System.Text.UTF8Encoding]::new($false))
    Write-CSCLog ("Alias '{0}' added to '{1}'" -f $alias, $Entry.Rel)
    return $true
}

function Open-NoteInEditor {
    <# Opens a note in Obsidian (SourceLink, {0} = path without .md, {1} = vault name) or with the default .md handler. #>
    param($Entry)
    if ($null -eq $Entry) { $Entry = $script:UI.CurrentNote }
    if ($null -eq $Entry) { Set-Status (Get-Text 'status.noNote'); return }
    try {
        if ($script:App.SourceLink) {
            $relNoExt = ($Entry.Rel -replace '\\', '/') -replace '\.md$', ''
            $target = $script:App.SourceLink -f [uri]::EscapeDataString($relNoExt), [uri]::EscapeDataString($script:App.VaultName)
            Write-CSCLog "Open: $target"
            Start-Process $target | Out-Null
        }
        else { Start-Process -FilePath $Entry.Full | Out-Null }
        Set-Status (Get-Text 'status.edit' -Format $Entry.Display)
    }
    catch { Write-CSCLog "Open-NoteInEditor failed: $_" 'Error'; Set-Status (Get-Text 'status.openFailed') }
}

function ConvertTo-NoteNameProposal {
    <# First non-empty line of a text, whitespace collapsed, hard cut at NewNoteNameMaxChars (on a word boundary when possible). #>
    param([AllowNull()][string]$Text)
    $line = (([string]$Text -split '\r?\n' | Where-Object { $_.Trim() } | Select-Object -First 1) -replace '\s+', ' ').Trim()
    $max = $script:App.NewNoteNameMaxChars
    if ($line.Length -gt $max) {
        $cut = $line.Substring(0, $max)
        $space = $cut.LastIndexOf(' ')
        if ($space -gt ($max * 0.5)) { $cut = $cut.Substring(0, $space) }
        $line = $cut.TrimEnd(' ', '-', ',', ';', ':', '.')
    }
    return ConvertTo-WindowsFilename $(if ($line) { $line } else { Get-Text 'app.newNoteDefault' })
}

function New-ThemedDialog {
    <# Loads a dialog from XAML and applies owner, window brushes and the dark title bar. #>
    param([Parameter(Mandatory)][xml]$Xaml)
    $dlg = [Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($Xaml))
    $dlg.Owner = $script:UI.Window
    foreach ($k in 'CscWindowBg', 'CscPanelBg', 'CscFg', 'CscMutedFg', 'CscBorder', 'CscAccent') { $dlg.Resources[$k] = $script:UI.Window.Resources[$k] }
    $dlg.Add_SourceInitialized({ param($s, $e) try { [CheatSheetCompanionHelper.Native]::SetDarkTitleBar(([System.Windows.Interop.WindowInteropHelper]::new($s)).Handle, (Get-EffectiveTheme).IsDark) } catch { } })
    return $dlg
}

function Show-AliasDialog {
    <# Asks for an alias text (prefilled) that is written to the front matter of the note. Returns the text or $null. #>
    param([Parameter(Mandatory)]$Entry, [AllowEmptyString()][string]$Proposal)
    [xml]$dlgXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Alias hinzufügen" Width="460" SizeToContent="Height" WindowStartupLocation="CenterOwner" ResizeMode="NoResize" Topmost="True" ShowInTaskbar="False"
        Background="{DynamicResource CscWindowBg}" Foreground="{DynamicResource CscFg}">
    <Window.Resources>
        <Style TargetType="TextBlock"><Setter Property="Foreground" Value="{DynamicResource CscFg}"/></Style>
    </Window.Resources>
    <StackPanel Margin="14">
        <TextBlock Name="Info" TextWrapping="Wrap" />
        <TextBox Name="AliasBox" Margin="0,6,0,8" Background="{DynamicResource CscPanelBg}" Foreground="{DynamicResource CscFg}" BorderBrush="{DynamicResource CscBorder}" />
        <TextBlock Name="Hint" FontSize="11" Foreground="{DynamicResource CscMutedFg}" TextWrapping="Wrap" />
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
            <Button Name="OkButton" MinWidth="110" Height="24" IsDefault="True" />
            <Button Name="CancelButton" MinWidth="90" Height="24" Margin="8,0,0,0" IsCancel="True" />
        </StackPanel>
    </StackPanel>
</Window>
'@
    $dlg = New-ThemedDialog -Xaml $dlgXaml
    $dlg.Title = Get-Text 'dlg.aliasTitle'
    $dlg.FindName('Info').Text = Get-Text 'dlg.aliasFor' -Format $Entry.Display
    $dlg.FindName('Hint').Text = Get-Text 'dlg.aliasInfo'
    $dlg.FindName('OkButton').Content = Get-Text 'dlg.aliasOk'
    $dlg.FindName('CancelButton').Content = Get-Text 'dlg.cancel'
    $box = $dlg.FindName('AliasBox')
    $box.Text = $Proposal
    $dlg.FindName('OkButton').Add_Click({ param($s, $e) [System.Windows.Window]::GetWindow($s).DialogResult = $true })
    $dlg.Add_Loaded({ param($s, $e) $tb = $s.FindName('AliasBox'); $tb.Focus(); $tb.SelectAll() })
    if ($dlg.ShowDialog() -ne $true -or [string]::IsNullOrWhiteSpace($box.Text)) { return $null }
    return $box.Text.Trim()
}

function Get-VaultFolderList {
    <# All folders of the active vault (without ignored ones), including intermediate levels, sorted. #>
    $set = @{}
    foreach ($e in $script:Index.Entries) {
        $f = [string]$e.Folder
        while ($f) { $set[$f] = $true; $f = [string](Split-Path $f -Parent) }
    }
    return @($set.Keys | Sort-Object)
}

function Show-NewNoteDialog {
    <#
        Dialog for a new note: name (prefilled from the chosen source), source selection
        (window title / clipboard / selection in the companion), optional body text, target folder.
        Returns @{ Name; Source; InsertText; Folder } or $null.
    #>
    param([Parameter(Mandatory)][hashtable]$Sources, [Parameter(Mandatory)][string]$DefaultSource)
    [xml]$dlgXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Neue Notiz" Width="480" SizeToContent="Height" WindowStartupLocation="CenterOwner" ResizeMode="NoResize" Topmost="True" ShowInTaskbar="False"
        Background="{DynamicResource CscWindowBg}" Foreground="{DynamicResource CscFg}">
    <Window.Resources>
        <Style TargetType="RadioButton"><Setter Property="Foreground" Value="{DynamicResource CscFg}"/><Setter Property="Margin" Value="0,3,0,0"/></Style>
        <Style TargetType="TextBlock"><Setter Property="Foreground" Value="{DynamicResource CscFg}"/></Style>
    </Window.Resources>
    <StackPanel Margin="14">
        <TextBlock Name="LblName" />
        <TextBox Name="NoteName" Margin="0,4,0,10" Background="{DynamicResource CscPanelBg}" Foreground="{DynamicResource CscFg}" BorderBrush="{DynamicResource CscBorder}" />
        <TextBlock Name="LblFrom" />
        <RadioButton Name="SrcWindow" GroupName="src" />
        <RadioButton Name="SrcClipboard" GroupName="src" />
        <RadioButton Name="SrcSelection" GroupName="src" />
        <CheckBox Name="InsertText" Margin="0,10,0,0" Foreground="{DynamicResource CscFg}" />
        <TextBlock Name="LblFolder" Margin="0,10,0,0" />
        <ComboBox Name="FolderBox" IsEditable="True" Margin="0,4,0,0" Background="{DynamicResource CscPanelBg}" Foreground="{DynamicResource CscFg}" BorderBrush="{DynamicResource CscBorder}" />
        <TextBlock Name="FolderInfo" FontSize="11" Foreground="{DynamicResource CscMutedFg}" TextWrapping="Wrap" Margin="0,8,0,0" />
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
            <Button Name="OkButton" MinWidth="130" Height="24" IsDefault="True" />
            <Button Name="CancelButton" MinWidth="90" Height="24" Margin="8,0,0,0" IsCancel="True" />
        </StackPanel>
    </StackPanel>
</Window>
'@
    $dlg = New-ThemedDialog -Xaml $dlgXaml
    $dlg.Title = Get-Text 'dlg.newTitle'
    $dlg.FindName('LblName').Text = Get-Text 'dlg.newName'
    $dlg.FindName('LblFrom').Text = Get-Text 'dlg.newFrom'
    $dlg.FindName('LblFolder').Text = Get-Text 'dlg.newFolder'
    $dlg.FindName('OkButton').Content = Get-Text 'dlg.newOk'
    $dlg.FindName('CancelButton').Content = Get-Text 'dlg.cancel'
    $nameBox = $dlg.FindName('NoteName')
    $chk = $dlg.FindName('InsertText')
    $chk.Content = Get-Text 'dlg.newInsert'
    $folderBox = $dlg.FindName('FolderBox')
    foreach ($f in (Get-VaultFolderList)) { [void]$folderBox.Items.Add($f) }
    $folderBox.Text = [string]$script:App.NewNoteFolder
    $radios = @{ window = $dlg.FindName('SrcWindow'); clipboard = $dlg.FindName('SrcClipboard'); selection = $dlg.FindName('SrcSelection') }
    $labels = @{ window = (Get-Text 'dlg.srcWindow'); clipboard = (Get-Text 'dlg.srcClipboard'); selection = (Get-Text 'dlg.srcSelection') }
    $script:UI.NewNoteDialog = @{ Window = $dlg; Sources = $Sources; NameBox = $nameBox; Check = $chk; Radios = $radios; Source = $DefaultSource }
    foreach ($key in $radios.Keys) {
        $rb = $radios[$key]
        $text = [string]$Sources[$key]
        $rb.Tag = $key
        $rb.Content = $labels[$key] + ':  ' + $(if ($text) { Limit-Text $text 48 } else { Get-Text 'dlg.empty' })
        $rb.IsEnabled = [bool]$text
        $rb.Add_Checked({
                param($s, $e)
                $d = $script:UI.NewNoteDialog
                if ($null -eq $d) { return }
                $d.Source = [string]$s.Tag
                $d.NameBox.Text = ConvertTo-NoteNameProposal -Text ([string]$d.Sources[$d.Source])
                $multi = ($d.Source -ne 'window') -and -not [string]::IsNullOrWhiteSpace([string]$d.Sources[$d.Source])
                $d.Check.IsEnabled = $multi
                $d.Check.IsChecked = $multi
            })
    }
    $radios[$DefaultSource].IsChecked = $true
    $dlg.FindName('FolderInfo').Text = (Get-Text 'dlg.vault' -Format $script:App.Vault) + $(if ($script:App.NewNoteTemplate) { "`n" + (Get-Text 'dlg.template' -Format $script:App.NewNoteTemplate) } else { '' })
    $dlg.FindName('OkButton').Add_Click({ param($s, $e) [System.Windows.Window]::GetWindow($s).DialogResult = $true })
    $dlg.Add_Loaded({ param($s, $e) $tb = $s.FindName('NoteName'); $tb.Focus(); $tb.SelectAll() })
    $ok = $dlg.ShowDialog()
    $result = $script:UI.NewNoteDialog
    $script:UI.NewNoteDialog = $null
    if ($ok -ne $true -or [string]::IsNullOrWhiteSpace($nameBox.Text)) { return $null }
    $folder = ([string]$folderBox.Text).Trim().Trim('\', '/') -replace '/', '\'
    return @{ Name = $nameBox.Text.Trim(); Source = $result.Source; InsertText = ([bool]$chk.IsChecked -and $chk.IsEnabled); Folder = $folder }
}

function New-NoteForCurrentContext {
    <#
        Creates a note. Name proposal comes from the most recent of: window title, clipboard text, selection in the
        companion (changeable in the dialog). Uses NewNoteTemplate when configured.
    #>
    $sources = @{ window = [string]$script:UI.LastWindow.Text; clipboard = [string]$script:UI.LastClipboard.Text; selection = [string]$script:UI.LastSelection.Text }
    $newest = @('window', 'clipboard', 'selection') | Where-Object { $sources[$_] } | Sort-Object { $script:UI[('Last' + (Get-Culture).TextInfo.ToTitleCase($_))].Time } -Descending | Select-Object -First 1
    if (-not $newest) { $sources.window = Get-Text 'app.newNoteDefault'; $newest = 'window' }
    $answer = Show-NewNoteDialog -Sources $sources -DefaultSource $newest
    if ($null -eq $answer) { return }
    $name = ConvertTo-WindowsFilename $answer.Name
    $dir = $script:App.Vault
    if ($answer.Folder) {
        $dir = Join-Path $dir $answer.Folder
        $full = [System.IO.Path]::GetFullPath($dir)
        if (-not $full.StartsWith($script:App.Vault, [StringComparison]::OrdinalIgnoreCase)) { Set-Status (Get-Text 'status.folderOutside'); return }
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $path = Join-Path $dir ($name + '.md')
    if (-not (Test-Path -LiteralPath $path)) {
        $content = if ($answer.InsertText) { ([string]$sources[$answer.Source]).Trim() } else { '' }
        $body = $null
        if ($script:App.NewNoteTemplate) {
            $tpl = Join-Path $script:App.Vault $script:App.NewNoteTemplate
            if (Test-Path -LiteralPath $tpl) {
                $body = [System.IO.File]::ReadAllText($tpl)
                $body = $body -replace '<%\s*tp\.date\.now\([^)]*\)\s*%>', (Get-Date -Format 'yyyy-MM-dd')
                $body = $body -replace '<%\s*tp\.file\.title\s*%>', $name
                $body = $body -replace '(?s)[ \t]*<%.*?%>[ \t]*(\r?\n)?', ''   # remaining Templater commands (move, cursor, ...)
                $body = $body.Replace('{{title}}', $name).Replace('{{date}}', (Get-Date -Format 'yyyy-MM-dd')).Replace('{{time}}', (Get-Date -Format 'HH:mm'))
                if ($body.Contains('{{content}}')) { $body = $body.Replace('{{content}}', $content) }
                elseif ($content) { $body = $body.TrimEnd() + "`r`n`r`n" + $content + "`r`n" }
            }
            else { Write-CSCLog "NewNoteTemplate '$tpl' not found" 'Warn' }
        }
        if ($null -eq $body) {
            $body = "# $name`r`n`r`n"
            if ($content) { $body += $content + "`r`n" }
        }
        [System.IO.File]::WriteAllText($path, $body, [System.Text.UTF8Encoding]::new($false))
        Write-CSCLog "Created note '$path' (source=$($answer.Source), content=$($answer.InsertText))"
    }
    [void](Update-VaultIndex)
    $rel = $path.Substring($script:App.Vault.Length + 1)
    $entry = $script:Index.ByPath[($rel -replace '\\', '/').ToLowerInvariant()]
    if ($entry) { Show-Note -Entry $entry; Open-NoteInEditor -Entry $entry }
}

function Invoke-NoteSearch {
    <# Opens the best search hit for a query (used by -OpenNote and Enter without a selection). #>
    param([Parameter(Mandatory)][string]$Query)
    $hits = Find-SearchResults -Query $Query
    if ($hits.Count -eq 0) {
        $c = Find-NoteCandidates -Targets @(@{ Text = $Query; Origin = 'manual'; Threshold = $script:App.ThresholdTitle })
        if ($c.Count -gt 0) { Show-Note -Entry $c[0].Entry; return }
        Set-Status (Get-Text 'status.notFound' -Format $Query); return
    }
    Show-Note -Entry $hits[0].Entry
}

function Set-Toggle {
    <# Central toggle handling (buttons, hotkeys, tray): FollowWindow | FollowClipboard | AutoCopy #>
    param([Parameter(Mandatory)][ValidateSet('FollowWindow', 'FollowClipboard', 'AutoCopy', 'Pinned')][string]$Name, [bool]$Value)
    if ($Name -eq 'Pinned') {
        # pin / unpin the visible note (PinnedNotes list); the button mirrors the visible note
        $entry = $script:UI.CurrentNote
        if ($null -eq $entry) { Set-Status (Get-Text 'status.noNote'); Update-PinButton; return }
        $key = '{0}|{1}' -f $script:App.VaultName, $entry.Rel
        if ($Value) {
            if (-not (Test-NotePinned -Entry $entry)) { $script:State.PinnedNotes = @(@($script:State.PinnedNotes) + @($key)) }
            Set-Status (Get-Text 'status.pinned' -Format $entry.Display)
        }
        else {
            $script:State.PinnedNotes = @(@($script:State.PinnedNotes) | Where-Object { $_ -ne $key })
            Set-Status (Get-Text 'status.unpinned' -Format $entry.Display)
        }
        if ($script:UI.CurrentView) {
            # re-render: the note appears as pinned tab / turns back into a normal tab
            $cands = @($script:UI.CurrentView.Candidates | ForEach-Object {
                    if (@($_.Origins) -contains 'pinned' -and $_.Entry.Key -eq $entry.Key) { [pscustomobject]@{ Entry = $_.Entry; Score = 1.0; Origins = @('manual'); Reasons = @() } } else { $_ }
                })
            Show-Candidates -Candidates $cands -SourceText $script:UI.CurrentView.SourceText -NoHistory -ActiveEntry $entry
        }
        Update-PinButton
        Update-StripText
        Save-State
        Write-CSCLog "Pin $($entry.Rel) = $Value"
        return
    }
    $script:State[$Name] = $Value
    $script:UI.SuppressToggleEvents = $true
    try {
        $script:UI["Toggle$Name"].IsChecked = $Value
        if ($script:UI.TrayItems -and $script:UI.TrayItems[$Name]) { $script:UI.TrayItems[$Name].Checked = $Value }
    }
    catch { } finally { $script:UI.SuppressToggleEvents = $false }
    $onOff = Get-Text $(if ($Value) { 'status.on' } else { 'status.off' })
    switch ($Name) {
        'FollowWindow'    { Set-Status (Get-Text 'status.followWindow' -Format $onOff); if ($Value) { $script:UI.ForegroundMatchedKey = ''; $script:UI.ClipboardProcessId = 0 } }
        'FollowClipboard' { Set-Status (Get-Text 'status.followClipboard' -Format $onOff) }
        'AutoCopy'        {
            Set-Status (Get-Text 'status.autoCopy' -Format $onOff)
            try { if ($script:UI.WebReady) { $null = $script:UI.WebView.CoreWebView2.ExecuteScriptAsync('window.CSC_AUTO_COPY=' + $(if ($Value) { 'true' } else { 'false' }) + ';') } } catch { }
        }
    }
    Write-CSCLog "Toggle $Name = $Value"
}

function Switch-Minimized {
    <# Collapses the window to a thin strip (and back). -Restore only expands. #>
    param([switch]$Restore)
    $w = $script:UI.Window
    if ($w.Height -lt 20) {
        $w.Height = $(if ($script:UI.RestoreHeight -gt 50) { $script:UI.RestoreHeight } else { 450 })
        $script:UI.InnerGrid.Visibility = 'Visible'
        $script:UI.StatusBar.Visibility = 'Visible'
        $w.WindowStyle = [System.Windows.WindowStyle]::ToolWindow
        $w.ResizeMode = [System.Windows.ResizeMode]::CanResize
        if ($script:UI.RestoreMaximized) { $w.WindowState = [System.Windows.WindowState]::Maximized }
        Resume-WebView
        Update-StripText
        try { $w.Activate() | Out-Null } catch { }
    }
    elseif (-not $Restore) {
        $script:UI.SearchPopup.IsOpen = $false
        # a maximized window ignores Height: leave the maximized state first and bring it back when expanding
        $script:UI.RestoreMaximized = ($w.WindowState -eq [System.Windows.WindowState]::Maximized)
        if ($script:UI.RestoreMaximized) { $w.WindowState = [System.Windows.WindowState]::Normal }
        $script:UI.RestoreHeight = $w.Height
        $script:UI.InnerGrid.Visibility = 'Hidden'
        $script:UI.StatusBar.Visibility = 'Hidden'
        $w.WindowStyle = [System.Windows.WindowStyle]::None
        $w.ResizeMode = [System.Windows.ResizeMode]::NoResize
        $w.Height = 13
        Update-StripText
        Suspend-WebView
    }
    Update-WindowOpacity
}

function Suspend-WebView {
    <#
        Collapsed strip: the WebView2 renderer is suspended (releases most of its memory, like a background tab) and told
        to use little memory; the host process gives unused pages back. Navigation while collapsed resumes it automatically.
    #>
    if (-not $script:UI.WebReady) { return }
    $script:UI.Window.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [Action] {
            try {
                $core = $script:UI.WebView.CoreWebView2
                if ($null -eq $core -or $script:UI.WebView.IsVisible) { return }
                $core.MemoryUsageTargetLevel = [Microsoft.Web.WebView2.Core.CoreWebView2MemoryUsageTargetLevel]::Low
                $task = $core.TrySuspendAsync()
                $check = [System.Windows.Threading.DispatcherTimer]::new()
                $check.Interval = [TimeSpan]::FromMilliseconds(800)
                $check.Tag = $task
                $check.Add_Tick({
                        param($s, $e)
                        $s.Stop()
                        try { Write-CSCLog ("WebView2 suspended: {0} (task {1})" -f $script:UI.WebView.CoreWebView2.IsSuspended, $(if ($s.Tag.IsCompleted) { $s.Tag.Result } else { 'pending' })) } catch { }
                    })
                $check.Start()
            }
            catch { Write-CSCLog "Suspend: $_" 'Warn' }
            try { [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [void][CheatSheetCompanionHelper.Native]::TrimWorkingSet() } catch { }
        }) | Out-Null
}

function Resume-WebView {
    if (-not $script:UI.WebReady) { return }
    try {
        $core = $script:UI.WebView.CoreWebView2
        if ($null -eq $core) { return }
        if ($core.IsSuspended) { $core.Resume() }   # normally already resumed by becoming visible
        $core.MemoryUsageTargetLevel = [Microsoft.Web.WebView2.Core.CoreWebView2MemoryUsageTargetLevel]::Normal
        Write-CSCLog 'WebView2 active again (expanded)'
    }
    catch { Write-CSCLog "Resume: $_" 'Warn' }
}

function Test-CursorOverWindow {
    <# True when the mouse cursor is inside the window rectangle (WPF's IsMouseOver is unreliable over the WebView2 child window). #>
    $w = $script:UI.Window
    try {
        $src = [System.Windows.PresentationSource]::FromVisual($w)
        $scale = if ($src) { $src.CompositionTarget.TransformToDevice.M11 } else { 1.0 }
        $p = [System.Windows.Forms.Cursor]::Position
        $x = $p.X / $scale; $y = $p.Y / $scale
        return ($x -ge $w.Left -and $x -le ($w.Left + $w.ActualWidth) -and $y -ge $w.Top -and $y -le ($w.Top + $w.ActualHeight))
    }
    catch { return $false }
}

function Update-WindowOpacity {
    <#
        InactiveOpacity: the window is dimmed while another program is active. Full opacity (WindowOpacity) when the
        window or one of its dialogs is active, under the mouse cursor, collapsed to the strip or while the search is open.
        Applied as layered window alpha (Window.Opacity would leave the WebView2 child window opaque).
        Called on Activated/Deactivated and from the foreground poll timer (cursor check).
    #>
    $w = $script:UI.Window
    if ($null -eq $w -or -not $script:App.UseWindowAlpha -or $script:UI.Hwnd -eq [IntPtr]::Zero) { return }
    $target = $script:App.WindowOpacity
    $dim = $script:App.InactiveOpacity
    if ($dim -gt 0 -and $dim -lt $target -and $w.Height -ge 20 -and -not $w.IsActive -and -not $script:UI.SearchPopup.IsOpen) {
        $ownedActive = $false
        foreach ($o in $w.OwnedWindows) { if ($o.IsActive) { $ownedActive = $true; break } }
        if (-not $ownedActive -and -not (Test-CursorOverWindow)) { $target = $dim }
    }
    if ([Math]::Abs($script:UI.CurrentOpacity - $target) -gt 0.001) {
        if ([CheatSheetCompanionHelper.Native]::SetWindowAlpha($script:UI.Hwnd, $target)) {
            $script:UI.CurrentOpacity = $target
            Write-CSCLog ("Opacity -> {0}" -f $target.ToString('0.00', [cultureinfo]::InvariantCulture))
        }
    }
}

function Show-Search {
    param([string]$Query = $null)
    Switch-Minimized -Restore
    try { $script:UI.Window.Activate() | Out-Null } catch { }
    $box = $script:UI.SearchBox
    if ($null -ne $Query) { $box.Text = $Query }
    $box.Focus() | Out-Null
    $box.SelectAll()
    if ($box.Text) { Update-SearchPopup -Query $box.Text }
}

#endregion

#region ---------- Search popup ----------

function Sync-SearchBox {
    <# Shows the current note name in the search box without triggering a search. #>
    param($Entry)
    if ($script:UI.SearchPopup.IsOpen) { return }
    $script:UI.SuppressSearchEvents = $true
    try { $script:UI.SearchBox.Text = $Entry.Display } catch { } finally { $script:UI.SuppressSearchEvents = $false }
}

function New-SearchItem {
    param([Parameter(Mandatory)]$Result)
    $item = [System.Windows.Controls.ListBoxItem]::new()
    $item.Tag = $Result
    $item.Padding = '6,3'
    $panel = [System.Windows.Controls.StackPanel]::new()
    $title = [System.Windows.Controls.TextBlock]::new()
    $title.FontSize = 13
    $title.TextTrimming = 'CharacterEllipsis'
    $text = [string]$Result.Entry.Display
    $pos = @{}; foreach ($p in $Result.Positions) { $pos[[int]$p] = $true }
    $run = [System.Text.StringBuilder]::new(); $bold = $false
    $flush = {
        if ($run.Length -eq 0) { return }
        $r = [System.Windows.Documents.Run]::new($run.ToString())
        if ($bold) { $r.FontWeight = 'Bold'; $r.SetResourceReference([System.Windows.Documents.Run]::ForegroundProperty, 'CscAccent') }
        [void]$title.Inlines.Add($r)
        [void]$run.Clear()
    }
    for ($i = 0; $i -lt $text.Length; $i++) {
        $isHit = $pos.ContainsKey($i)
        if ($isHit -ne $bold) { & $flush; $bold = $isHit }
        [void]$run.Append($text[$i])
    }
    & $flush
    [void]$panel.Children.Add($title)
    if ($Result.Detail) {
        $detail = [System.Windows.Controls.TextBlock]::new()
        $detail.FontSize = 11
        $detail.TextTrimming = 'CharacterEllipsis'
        $detail.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'CscMutedFg')
        $detail.Text = $(switch ($Result.Kind) { 'content' { Get-Text 'ui.contentPrefix' } 'alias' { Get-Text 'ui.aliasPrefix' } default { '' } }) + [string]$Result.Detail
        [void]$panel.Children.Add($detail)
    }
    $item.Content = $panel
    $item.Add_PreviewMouseLeftButtonUp({ param($s, $e) Open-SearchResult -Result $s.Tag })
    return $item
}

function Update-SearchPopup {
    param([AllowEmptyString()][string]$Query)
    $popup = $script:UI.SearchPopup
    $list = $script:UI.SearchList
    if ([string]::IsNullOrWhiteSpace($Query)) { $popup.IsOpen = $false; return }
    $results = Find-SearchResults -Query $Query
    $list.Items.Clear()
    foreach ($r in $results) { [void]$list.Items.Add((New-SearchItem -Result $r)) }
    if ($results.Count -eq 0) { $popup.IsOpen = $false; return }
    $popup.Width = [Math]::Max(380, $script:UI.Window.ActualWidth - 40)
    $popup.IsOpen = $true
    $list.SelectedIndex = 0
}

function Open-SearchResult {
    param($Result)
    $script:UI.SearchPopup.IsOpen = $false
    if ($null -eq $Result) { return }
    Show-Note -Entry $Result.Entry
    try { $script:UI.WebView.Focus() | Out-Null } catch { }
}

function Move-SearchSelection {
    param([int]$Delta)
    $list = $script:UI.SearchList
    if (-not $script:UI.SearchPopup.IsOpen -or $list.Items.Count -eq 0) { return }
    $n = $list.Items.Count
    $list.SelectedIndex = (($list.SelectedIndex + $Delta) % $n + $n) % $n
    $list.ScrollIntoView($list.SelectedItem)
}

#endregion

#region ---------- WebView2 ----------

$dllDir = Join-Path $script:App.Root 'dll'
foreach ($dll in 'Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.Wpf.dll', 'WebView2Loader.dll') {
    if (-not (Test-Path -LiteralPath (Join-Path $dllDir $dll))) { throw "WebView2 library missing: $dllDir\$dll (see README)" }
}
[void][CheatSheetCompanionHelper.Native]::SetDllDirectory($dllDir)
Add-Type -Path (Join-Path $dllDir 'Microsoft.Web.WebView2.Core.dll')
Add-Type -Path (Join-Path $dllDir 'Microsoft.Web.WebView2.Wpf.dll')

function Save-WebViewScreenshot {
    <# Test helper: captures the page as PNG without blocking the UI thread (task is polled by a timer). #>
    param([Parameter(Mandatory)][string]$Path)
    if ($Path -like '*.window.png') {
        # whole window including WPF chrome and popups (screen capture of the window area plus the popup zone below)
        $w = $script:UI.Window
        $src = [System.Windows.PresentationSource]::FromVisual($w)
        $scale = if ($src) { $src.CompositionTarget.TransformToDevice.M11 } else { 1.0 }
        $rect = [System.Drawing.Rectangle]::new([int]($w.Left * $scale), [int]($w.Top * $scale), [int]($w.ActualWidth * $scale), [int](($w.ActualHeight + 360) * $scale))
        $bmp = [System.Drawing.Bitmap]::new($rect.Width, $rect.Height)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($rect.Location, [System.Drawing.Point]::Empty, $rect.Size)
        $g.Dispose()
        $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
        Write-CSCLog "Window screenshot -> $Path"
        return
    }
    $stream = [System.IO.File]::Create($Path)
    $task = $script:UI.WebView.CoreWebView2.CapturePreviewAsync([Microsoft.Web.WebView2.Core.CoreWebView2CapturePreviewImageFormat]::Png, $stream)
    $poll = [System.Windows.Threading.DispatcherTimer]::new()
    $poll.Interval = [TimeSpan]::FromMilliseconds(100)
    $poll.Tag = @{ Task = $task; Stream = $stream; Path = $Path }
    $poll.Add_Tick({
            param($s, $e)
            if (-not $s.Tag.Task.IsCompleted) { return }
            $s.Stop()
            try { $s.Tag.Stream.Dispose() } catch { }
            Write-CSCLog ("Screenshot -> {0} ({1})" -f $s.Tag.Path, $s.Tag.Task.Status)
        })
    $poll.Start()
    $script:UI.Timers += $poll
}

function Invoke-MenuAction {
    <# Dispatches context menu commands (actions are stored by CommandId, no closures needed). #>
    param([Parameter(Mandatory)][int]$CommandId)
    $a = $script:UI.MenuActions[$CommandId]
    if ($null -eq $a) { return }
    switch ($a.Kind) {
        'copy'       { if (Set-ClipboardTextSafe -Text $a.Value) { Set-Status (Get-Text 'status.selectionCopied') } }
        'recopy'     { if (Set-ClipboardTextSafe -Text $a.Value) { Set-Status (Get-Text 'status.recopied' -Format (Get-CopyHistoryLabel -Text $a.Value -Max 50)) } }
        'edit'       { Open-NoteInEditor }
        'theme'      { Set-Theme -Name $a.Value }
        'vault'      { Initialize-Vault -Name $a.Value; Save-State }
        'cliplookup' { Invoke-ClipboardLookup -Force }
        'copyfmt'    { Copy-NoteContent -Plain:([bool]$a.Plain) -WholeNote:([bool]$a.WholeNote) }
        'append'     {
            $entry = $script:UI.CurrentNote
            if ($null -eq $entry) { Set-Status (Get-Text 'status.noNote'); return }
            $text = Get-ClipboardTextSafe
            if ([string]::IsNullOrWhiteSpace($text)) { Set-Status (Get-Text 'status.clipboardEmpty'); return }
            if (Add-TextToNote -Entry $entry -Text $text) { Set-Status (Get-Text 'status.appended' -Format $entry.Display, $text.Length) }
        }
        'pin'        { Set-Toggle -Name Pinned -Value (-not (Test-NotePinned -Entry $script:UI.CurrentNote)) }
        'alias'      {
            $entry = $script:UI.CurrentNote
            if ($null -eq $entry) { Set-Status (Get-Text 'status.noNote'); return }
            $alias = Show-AliasDialog -Entry $entry -Proposal ([string]$a.Value)
            if ($alias -and (Add-NoteAlias -Entry $entry -Alias $alias)) {
                [void](Update-VaultIndex)
                Set-Status (Get-Text 'status.aliasAdded' -Format $alias, $entry.Display)
            }
        }
    }
}

function Initialize-WebViewCore {
    <# Called once when CoreWebView2 is ready: settings, virtual hosts, events. #>
    param([Parameter(Mandatory)]$Core)
    $Core.Settings.IsWebMessageEnabled = $true
    $Core.Settings.AreDefaultContextMenusEnabled = $true
    $Core.Settings.AreDevToolsEnabled = [bool]$script:App.Debug
    $Core.Settings.IsStatusBarEnabled = $false
    $Core.Settings.IsZoomControlEnabled = $true
    $allow = [Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind]::Allow
    $Core.SetVirtualHostNameToFolderMapping('app.csc', $script:App.Root, $allow)
    $Core.SetVirtualHostNameToFolderMapping('vault.csc', $script:App.Vault, $allow)

    $Core.add_NavigationStarting([System.EventHandler[Microsoft.Web.WebView2.Core.CoreWebView2NavigationStartingEventArgs]] {
            param($s, $e)
            try {
                $uri = [string]$e.Uri
                if ($uri.StartsWith('https://app.csc/') -or $uri -eq 'about:blank') { return }
                $e.Cancel = $true
                if ($uri.StartsWith('https://vault.csc/')) {
                    $rel = [uri]::UnescapeDataString($uri.Substring('https://vault.csc/'.Length)) -replace '/', '\'
                    $full = Join-Path $script:App.Vault $rel
                    if (Test-Path -LiteralPath $full) { Start-Process -FilePath $full | Out-Null }
                    return
                }
                Start-Process $uri | Out-Null
            }
            catch { Write-CSCLog "NavigationStarting: $_" 'Warn' }
        })

    $Core.add_NewWindowRequested([System.EventHandler[Microsoft.Web.WebView2.Core.CoreWebView2NewWindowRequestedEventArgs]] {
            param($s, $e)
            try { $e.Handled = $true; Start-Process ([string]$e.Uri) | Out-Null } catch { }
        })

    $Core.add_WebMessageReceived([System.EventHandler[Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs]] {
            param($s, $e)
            try {
                $msg = $e.WebMessageAsJson | ConvertFrom-Json
                switch ([string]$msg.type) {
                    'selection' {
                        if ($script:State.AutoCopy -and -not [string]::IsNullOrWhiteSpace([string]$msg.text) -and [double]$msg.held -ge [double]$script:App.AutoCopyMinHoldMs) {
                            if (Set-ClipboardTextSafe -Text ([string]$msg.text)) { Set-Status (Get-Text 'status.copied' -Format ([string]$msg.text).Length) }
                        }
                    }
                    'activeNote' {
                        $entry = $script:Index.ByPath[(([string]$msg.rel) -replace '\\', '/').ToLowerInvariant()]
                        if ($entry) {
                            $script:UI.CurrentNote = $entry
                            if ($script:UI.CurrentView -and $null -ne $msg.index) { $script:UI.CurrentView.ActiveIndex = [int]$msg.index }
                            Add-RecentNote -Entry $entry; Sync-SearchBox -Entry $entry; Update-StripText; Update-PinButton
                            Set-Status ((Get-Text 'status.note' -Format $entry.Display) + $(if ($msg.score) { ' (' + $msg.score + ')' }))
                        }
                    }
                    'clearCopyHistory' {
                        $script:State.CopyHistory = @()
                        Save-State; Update-CopyHistoryPanel; Set-Status (Get-Text 'status.copyCleared')
                    }
                    'recopy' {
                        $hist = @($script:State.CopyHistory)
                        $k = [int]$msg.index
                        if ($k -ge 0 -and $k -lt $hist.Count -and $hist[$k].Text) {
                            if (Set-ClipboardTextSafe -Text ([string]$hist[$k].Text)) { Set-Status (Get-Text 'status.recopied' -Format (Get-CopyHistoryLabel -Text ([string]$hist[$k].Text) -Max 50)) }
                        }
                    }
                    'preview' {
                        # hover popup: the rendered target note (embeds resolved) goes back to the page, which shows it next to the link
                        if ($script:App.LinkPreview -and $msg.rel) {
                            Write-CSCLog ("Preview request {0}: {1}" -f [string]$msg.id, [string]$msg.rel)
                            $entry = $script:Index.ByPath[(([string]$msg.rel) -replace '\\', '/').ToLowerInvariant()]
                            if ($entry -and $entry.IsNote) {
                                $html = Get-NoteViewHtml -Entry $entry   # keeps the note meta data, so dataviewjs blocks work in the popup too
                                if ($html.Length -gt 400000) { $html = $html.Substring(0, 400000) }
                                $json = ConvertTo-Json -InputObject @{ id = [string]$msg.id; title = $entry.Display; rel = $entry.Rel; anchor = [string]$msg.anchor; html = $html } -Compress -Depth 3
                                $null = $script:UI.WebView.CoreWebView2.ExecuteScriptAsync('if(window.CSC_showPreview){window.CSC_showPreview(' + $json + ');}')
                            }
                        }
                    }
                    'openNote' {
                        $entry = $null
                        if ($msg.rel) { $entry = $script:Index.ByPath[(([string]$msg.rel) -replace '\\', '/').ToLowerInvariant()] }
                        elseif ($msg.note) { $entry = Resolve-VaultLink -Target ([string]$msg.note) }
                        if ($entry -and $entry.IsNote) { Show-Note -Entry $entry -Origin 'link' -Anchor ([string]$msg.anchor) }
                        elseif ($msg.anchor) { $null = $script:UI.WebView.CoreWebView2.ExecuteScriptAsync('location.hash=' + (ConvertTo-Json ([string]$msg.anchor))) }
                    }
                    'search' { Show-Search -Query ([string]$msg.query) }
                    'focusSearch' { Show-Search }
                    'sel' { $script:UI.LastSelection = @{ Text = [string]$msg.text; Time = [DateTime]::UtcNow } }
                    'copyCode' { if (-not [string]::IsNullOrWhiteSpace([string]$msg.text)) { if (Set-ClipboardTextSafe -Text ([string]$msg.text)) { Set-Status (Get-Text 'status.codeCopied') } } }
                    'newNote' { New-NoteForCurrentContext }
                    'dblClick' { Open-NoteInEditor }
                    'back' { Invoke-ViewBack }
                    'forward' { Invoke-ViewForward }
                    'copyFormatted' { Copy-NoteContent }
                    'log' { Write-CSCLog ('Page: ' + [string]$msg.text) }
                    'ready' { }
                }
            }
            catch { Write-CSCLog "WebMessage: $_" 'Warn' }
        })

    $Core.add_ContextMenuRequested([System.EventHandler[Microsoft.Web.WebView2.Core.CoreWebView2ContextMenuRequestedEventArgs]] {
            param($s, $e)
            try {
                $env = $s.Environment
                $items = $e.MenuItems
                $cmdKind = [Microsoft.Web.WebView2.Core.CoreWebView2ContextMenuItemKind]::Command
                # CustomItemSelected is EventHandler<object>; a plain EventHandler is rejected and the whole menu stayed default
                $handler = [System.EventHandler[object]] { param($s2, $e2) try { Invoke-MenuAction -CommandId $s2.CommandId } catch { Write-CSCLog "MenuAction: $_" 'Warn' } }
                $script:UI.MenuActions = @{}
                $insertAt = 0
                $target = $e.ContextMenuTarget
                if ($target -and $target.HasSelection -and -not [string]::IsNullOrWhiteSpace($target.SelectionText)) {
                    $copyItem = $env.CreateContextMenuItem((Get-Text 'menu.copy'), $null, $cmdKind)
                    $script:UI.MenuActions[$copyItem.CommandId] = @{ Kind = 'copy'; Value = [string]$target.SelectionText }
                    $copyItem.add_CustomItemSelected($handler)
                    $items.Insert($insertAt++, $copyItem)
                }
                if ($script:UI.CurrentNote) {
                    # formatted (HTML + text) for Teams / Outlook, or the whole note; the plain selection copy above stays one click
                    $copyMenu = $env.CreateContextMenuItem((Get-Text 'menu.copyAs'), $null, [Microsoft.Web.WebView2.Core.CoreWebView2ContextMenuItemKind]::Submenu)
                    $variants = @()
                    if ($target -and $target.HasSelection) { $variants += @{ Label = (Get-Text 'menu.copySelFormatted'); Plain = $false; WholeNote = $false } }
                    $variants += @{ Label = (Get-Text 'menu.copyNoteText'); Plain = $true; WholeNote = $true }
                    $variants += @{ Label = (Get-Text 'menu.copyNoteFormatted'); Plain = $false; WholeNote = $true }
                    foreach ($v in $variants) {
                        $ci = $env.CreateContextMenuItem($v.Label, $null, $cmdKind)
                        $script:UI.MenuActions[$ci.CommandId] = @{ Kind = 'copyfmt'; Plain = $v.Plain; WholeNote = $v.WholeNote }
                        $ci.add_CustomItemSelected($handler)
                        $copyMenu.Children.Insert($copyMenu.Children.Count, $ci)
                    }
                    $items.Insert($insertAt++, $copyMenu)
                }
                $editItem = $env.CreateContextMenuItem((Get-Text 'menu.edit'), $null, $cmdKind)
                $script:UI.MenuActions[$editItem.CommandId] = @{ Kind = 'edit' }
                $editItem.add_CustomItemSelected($handler)
                $items.Insert($insertAt++, $editItem)
                if ($script:UI.CurrentNote) {
                    $clipText = ([string]$script:UI.LastClipboard.Text -replace '\s+', ' ').Trim()
                    $appendItem = $env.CreateContextMenuItem((Get-Text 'menu.append') + $(if ($clipText) { ':  ' + (Limit-Text $clipText 40) } else { '' }), $null, $cmdKind)
                    $script:UI.MenuActions[$appendItem.CommandId] = @{ Kind = 'append' }
                    $appendItem.add_CustomItemSelected($handler)
                    $items.Insert($insertAt++, $appendItem)
                }

                if ($script:UI.CurrentNote) {
                    # assign the current window / clipboard text to this note via a front matter alias
                    $aliasMenu = $env.CreateContextMenuItem((Get-Text 'menu.alias'), $null, [Microsoft.Web.WebView2.Core.CoreWebView2ContextMenuItemKind]::Submenu)
                    $proposals = @(
                        @{ Label = (Get-Text 'menu.aliasWindow'); Value = [string]$script:UI.LastForeground.Title }
                        @{ Label = (Get-Text 'menu.aliasProcess'); Value = [string]$script:UI.LastForeground.Process }
                        @{ Label = (Get-Text 'menu.aliasClipboard'); Value = [string]$script:UI.LastClipboard.Text }
                        @{ Label = (Get-Text 'menu.aliasSelection'); Value = $(if ($target -and $target.HasSelection) { [string]$target.SelectionText } else { '' }) }
                        @{ Label = (Get-Text 'menu.aliasCustom'); Value = ''; Always = $true }
                    )
                    foreach ($p in $proposals) {
                        $v = ($p.Value -replace '\s+', ' ').Trim()
                        if ($v.Length -gt 120) { $v = $v.Substring(0, 120) }
                        if (-not $v -and -not $p.Always) { continue }
                        $ai = $env.CreateContextMenuItem($p.Label + $(if ($v) { ':  ' + (Limit-Text $v 40) } else { '' }), $null, $cmdKind)
                        $script:UI.MenuActions[$ai.CommandId] = @{ Kind = 'alias'; Value = $v }
                        $ai.add_CustomItemSelected($handler)
                        $aliasMenu.Children.Insert($aliasMenu.Children.Count, $ai)
                    }
                    $items.Insert($insertAt++, $aliasMenu)
                }
                $clipItem = $env.CreateContextMenuItem((Get-Text 'menu.clipLookup'), $null, $cmdKind)
                $script:UI.MenuActions[$clipItem.CommandId] = @{ Kind = 'cliplookup' }
                $clipItem.add_CustomItemSelected($handler)
                $items.Insert($insertAt++, $clipItem)
                $pinItem = $env.CreateContextMenuItem((Get-Text 'menu.pin'), $null, [Microsoft.Web.WebView2.Core.CoreWebView2ContextMenuItemKind]::CheckBox)
                $pinItem.IsChecked = (Test-NotePinned -Entry $script:UI.CurrentNote)
                $script:UI.MenuActions[$pinItem.CommandId] = @{ Kind = 'pin' }
                $pinItem.add_CustomItemSelected($handler)
                $items.Insert($insertAt++, $pinItem)

                $themeMenu = $env.CreateContextMenuItem((Get-Text 'menu.theme'), $null, [Microsoft.Web.WebView2.Core.CoreWebView2ContextMenuItemKind]::Submenu)
                foreach ($name in (Get-ThemeNames)) {
                    $ti = $env.CreateContextMenuItem((Get-ThemeLabel $name), $null, [Microsoft.Web.WebView2.Core.CoreWebView2ContextMenuItemKind]::Radio)
                    $ti.IsChecked = ($name -eq [string]$script:State.Theme)
                    $script:UI.MenuActions[$ti.CommandId] = @{ Kind = 'theme'; Value = $name }
                    $ti.add_CustomItemSelected($handler)
                    $themeMenu.Children.Insert($themeMenu.Children.Count, $ti)
                }
                $items.Insert($insertAt++, $themeMenu)

                if ($script:App.Vaults.Count -gt 1) {
                    $vaultMenu = $env.CreateContextMenuItem((Get-Text 'menu.vault'), $null, [Microsoft.Web.WebView2.Core.CoreWebView2ContextMenuItemKind]::Submenu)
                    foreach ($vn in $script:App.Vaults.Keys) {
                        $vi = $env.CreateContextMenuItem([string]$vn, $null, [Microsoft.Web.WebView2.Core.CoreWebView2ContextMenuItemKind]::Radio)
                        $vi.IsChecked = ($vn -eq $script:App.VaultName)
                        $script:UI.MenuActions[$vi.CommandId] = @{ Kind = 'vault'; Value = [string]$vn }
                        $vi.add_CustomItemSelected($handler)
                        $vaultMenu.Children.Insert($vaultMenu.Children.Count, $vi)
                    }
                    $items.Insert($insertAt++, $vaultMenu)
                }
                $items.Insert($insertAt, $env.CreateContextMenuItem('', $null, [Microsoft.Web.WebView2.Core.CoreWebView2ContextMenuItemKind]::Separator))
            }
            catch { Write-CSCLog "ContextMenu: $_" 'Warn' }
        })

    $script:UI.WebReady = $true
    try { if ($script:State.Zoom -gt 0.3 -and $script:State.Zoom -lt 5) { $script:UI.WebView.ZoomFactor = [double]$script:State.Zoom } } catch { }
    Update-WindowTheme
    Write-CSCLog 'WebView2 ready'
    $pins = @(Get-PinnedEntries)
    if ($OpenNote) { Invoke-NoteSearch -Query $OpenNote }
    elseif ($pins.Count -gt 0) {
        Show-Candidates -Candidates @() -SourceText ''
        Set-Status (Get-Text 'status.pinnedStart' -Format (($pins | ForEach-Object { $_.Display }) -join ', '))
    }
    elseif ($script:UI.PendingNavigation) {
        $script:UI.PendingNavigation = $false
        $Core.Navigate('https://app.csc/Output/_view.html?v=' + [DateTime]::UtcNow.Ticks)
    }
    if ($SearchQuery) { Show-Search -Query $SearchQuery }
}

#endregion

#region ---------- UI (XAML) ----------

[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:wv="clr-namespace:Microsoft.Web.WebView2.Wpf;assembly=Microsoft.Web.WebView2.Wpf"
    Title="CheatSheetCompanion" Width="620" Height="480" MinWidth="320" MinHeight="13"
    WindowStyle="ToolWindow" Topmost="True" ShowInTaskbar="True"
    Background="{DynamicResource CscWindowBg}">
    <Window.Resources>
        <SolidColorBrush x:Key="CscWindowBg" Color="#F0F0F0"/>
        <SolidColorBrush x:Key="CscPanelBg" Color="#FAFAFA"/>
        <SolidColorBrush x:Key="CscFg" Color="#1E1E1E"/>
        <SolidColorBrush x:Key="CscMutedFg" Color="#6C7076"/>
        <SolidColorBrush x:Key="CscBorder" Color="#A0A0A0"/>
        <SolidColorBrush x:Key="CscAccent" Color="#7852EE"/>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource CscPanelBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource CscFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource CscBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Height" Value="22"/>
            <Setter Property="MinWidth" Value="26"/>
            <Setter Property="Margin" Value="3,0,0,0"/>
            <Setter Property="Padding" Value="3,0"/>
        </Style>
        <Style TargetType="ToggleButton">
            <Setter Property="Background" Value="{DynamicResource CscPanelBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource CscFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource CscBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Height" Value="22"/>
            <Setter Property="MinWidth" Value="34"/>
            <Setter Property="Margin" Value="3,0,0,0"/>
            <Setter Property="Padding" Value="6,0"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ToggleButton">
                        <Border Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="2" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="BorderBrush" Value="{DynamicResource CscAccent}"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource CscAccent}"/>
                                <Setter TargetName="Bd" Property="BorderBrush" Value="{DynamicResource CscAccent}"/>
                                <Setter Property="Foreground" Value="White"/>
                                <Setter Property="FontWeight" Value="SemiBold"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{DynamicResource CscFg}"/>
        </Style>
    </Window.Resources>
    <Grid Name="MainGrid" Margin="5">
        <Border Name="MinimizeBorder" BorderThickness="1" BorderBrush="{DynamicResource CscBorder}" Background="{DynamicResource CscPanelBg}" Margin="-10,-10,-20,-10" VerticalAlignment="Top" ToolTip="Klick: ein-/ausklappen. Rechtsklick: beenden.">
            <Button Name="MinimizeButton" Height="12" Margin="0" Padding="0" BorderThickness="0">
                <TextBlock Name="MinimizeText" FontFamily="Segoe UI Symbol" Text="&#x1f53a;&#x1f53a;&#x1f53a;" FontSize="9"/>
            </Button>
        </Border>
        <Grid Name="InnerGrid">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <DockPanel Grid.Row="0" Margin="0,12,0,0" LastChildFill="True">
                <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Top">
                    <Button Name="btnBack" ToolTip="Zurück (Alt+Links)"><TextBlock FontFamily="Segoe UI Symbol" Text="&#xE0A6;"/></Button>
                    <Button Name="btnForward" ToolTip="Vorwärts (Alt+Rechts)"><TextBlock FontFamily="Segoe UI Symbol" Text="&#xE0AD;"/></Button>
                    <Button Name="btnEditMarkdown" ToolTip="Angezeigte Notiz bearbeiten (Obsidian / Standard-Editor)."><TextBlock FontFamily="Segoe UI Symbol" Text="&#xE104;"/></Button>
                    <Button Name="btnAddMarkdown" ToolTip="Neue Notiz für das zuletzt aktive Fenster bzw. den kopierten Text anlegen."><TextBlock FontFamily="Segoe UI Symbol" Text="&#xE109;"/></Button>
                    <ToggleButton Name="TogglePinned" Content="Pin" ToolTip="Notiz anheften: Fenster und Zwischenablage ändern die Ansicht nicht mehr. Bleibt über einen Neustart erhalten."/>
                    <ToggleButton Name="ToggleFollowWindow" Content="Fenster" ToolTip="Notizen passend zum aktiven Fenster anzeigen (Titel + Prozess)."/>
                    <ToggleButton Name="ToggleFollowClipboard" Content="Clip" ToolTip="Notizen passend zum kopierten Text anzeigen (z.B. aus RDP/SSH-Sitzungen)."/>
                    <ToggleButton Name="ToggleAutoCopy" Content="AC" ToolTip="Auto Copy: markierten Text beim Loslassen der Maus kopieren (Mindest-Haltezeit konfigurierbar)."/>
                </StackPanel>
                <Grid Margin="0,0,6,0" VerticalAlignment="Top">
                    <TextBox Name="SearchBox" Height="22" VerticalContentAlignment="Center" Padding="3,0" Background="{DynamicResource CscPanelBg}" Foreground="{DynamicResource CscFg}" BorderBrush="{DynamicResource CscBorder}" ToolTip="Notiz suchen: Name oder Inhalt (Strg+K). Pfeiltasten + Enter öffnen."/>
                    <TextBlock Name="SearchHint" Text="Suchen…  (Strg+K)" Margin="6,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False" Foreground="{DynamicResource CscMutedFg}"/>
                    <Popup Name="SearchPopup" Placement="Bottom" StaysOpen="False" AllowsTransparency="True" PopupAnimation="Fade">
                        <Border Background="{DynamicResource CscPanelBg}" BorderBrush="{DynamicResource CscBorder}" BorderThickness="1" CornerRadius="4" Padding="2">
                            <ListBox Name="SearchList" MaxHeight="340" BorderThickness="0" Background="Transparent" Foreground="{DynamicResource CscFg}" HorizontalContentAlignment="Stretch"/>
                        </Border>
                    </Popup>
                </Grid>
            </DockPanel>
            <wv:WebView2 Name="WebBrowser" Grid.Row="1" Margin="0,6,0,26"/>
        </Grid>
        <StatusBar Name="StatusBar" VerticalAlignment="Bottom" Margin="-5,0,-5,-5" Height="22" Background="{DynamicResource CscPanelBg}" Foreground="{DynamicResource CscFg}">
            <StatusBarItem><TextBlock Name="txtStatus" Text="Bereit" TextTrimming="CharacterEllipsis"/></StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
'@

$window = [Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($xaml))
$script:UI = @{
    Window                = $window
    WebView               = $window.FindName('WebBrowser')
    Status                = $window.FindName('txtStatus')
    StatusBar             = $window.FindName('StatusBar')
    InnerGrid             = $window.FindName('InnerGrid')
    MinimizeText          = $window.FindName('MinimizeText')
    SearchBox             = $window.FindName('SearchBox')
    SearchHint            = $window.FindName('SearchHint')
    SearchPopup           = $window.FindName('SearchPopup')
    SearchList            = $window.FindName('SearchList')
    ToggleFollowWindow    = $window.FindName('ToggleFollowWindow')
    ToggleFollowClipboard = $window.FindName('ToggleFollowClipboard')
    ToggleAutoCopy        = $window.FindName('ToggleAutoCopy')
    TogglePinned          = $window.FindName('TogglePinned')
    WebReady              = $false
    PendingNavigation     = $false
    SuppressSearchEvents  = $false
    SuppressToggleEvents  = $false
    CurrentNote           = $null
    CurrentView           = $null
    NewNoteDialog         = $null
    History               = [System.Collections.Generic.Stack[object]]::new()
    Forward               = [System.Collections.Generic.Stack[object]]::new()
    RestoreHeight         = 480
    RestoreMaximized      = $false
    LastForeground        = @{ Key = ''; Title = ''; Process = ''; Stable = 0 }
    ForegroundMatchedKey  = ''
    LastWindow            = @{ Text = ''; Time = [DateTime]::MinValue }
    LastClipboard         = @{ Text = ''; Time = [DateTime]::MinValue }
    LastSelection         = @{ Text = ''; Time = [DateTime]::MinValue }
    ProcessNames          = @{}
    LastOwnClipboardText  = ''
    LastOwnClipboardTime  = [DateTime]::MinValue
    ClipboardProcessId    = 0
    MenuActions           = @{}
    Hwnd                  = [IntPtr]::Zero
    CurrentOpacity        = 1.0
    HotKeys               = @()
    Sink                  = $null
    Watcher               = $null
    Tray                  = $null
    TrayItems             = @{}
    Timers                = @()
}
$script:UI.SearchPopup.PlacementTarget = $script:UI.SearchBox

# --- UI texts (language file) ---
$window.FindName('MinimizeBorder').ToolTip = Get-Text 'ui.stripTip'
$window.FindName('btnBack').ToolTip = Get-Text 'ui.back'
$window.FindName('btnForward').ToolTip = Get-Text 'ui.forward'
$window.FindName('btnAddMarkdown').ToolTip = Get-Text 'ui.add'
$script:UI.TogglePinned.Content = Get-Text 'ui.pin'; $script:UI.TogglePinned.ToolTip = Get-Text 'ui.pinTip'
$script:UI.ToggleFollowWindow.Content = Get-Text 'ui.window'; $script:UI.ToggleFollowWindow.ToolTip = Get-Text 'ui.windowTip'
$script:UI.ToggleFollowClipboard.Content = Get-Text 'ui.clip'; $script:UI.ToggleFollowClipboard.ToolTip = Get-Text 'ui.clipTip'
$script:UI.ToggleAutoCopy.Content = Get-Text 'ui.autoCopy'; $script:UI.ToggleAutoCopy.ToolTip = Get-Text 'ui.autoCopyTip'
$script:UI.SearchBox.ToolTip = Get-Text 'ui.searchTip'
$script:UI.SearchHint.Text = Get-Text 'ui.searchHint'
$script:UI.Status.Text = Get-Text 'app.ready'

# --- window placement ---
if ($script:State.Width -gt 200 -and $script:State.Height -gt 50) { $window.Width = $script:State.Width; $window.Height = $script:State.Height }
if ($null -ne $script:State.Left -and $null -ne $script:State.Top) {
    $vs = [System.Windows.SystemParameters]::VirtualScreenWidth
    $vh = [System.Windows.SystemParameters]::VirtualScreenHeight
    if ($script:State.Left -ge -50 -and $script:State.Left -lt ($vs - 100) -and $script:State.Top -ge -10 -and $script:State.Top -lt ($vh - 50)) {
        $window.WindowStartupLocation = 'Manual'; $window.Left = $script:State.Left; $window.Top = $script:State.Top
    }
}
$script:UI.RestoreHeight = $window.Height

# --- opacity (WindowOpacity / InactiveOpacity; applied once the window handle exists, see SourceInitialized) ---
if ($script:App.InactiveOpacity -gt 0) {
    $window.Add_Activated({ Update-WindowOpacity })
    $window.Add_Deactivated({ Update-WindowOpacity })
}

# --- toggles ---
$script:UI.SuppressToggleEvents = $true
$script:UI.ToggleFollowWindow.IsChecked = [bool]$script:State.FollowWindow
$script:UI.ToggleFollowClipboard.IsChecked = [bool]$script:State.FollowClipboard
$script:UI.ToggleAutoCopy.IsChecked = [bool]$script:State.AutoCopy
$script:UI.TogglePinned.IsChecked = $false   # mirrors the visible note (Update-PinButton)
$script:UI.SuppressToggleEvents = $false
foreach ($name in 'FollowWindow', 'FollowClipboard', 'AutoCopy', 'Pinned') {
    $btn = $script:UI["Toggle$name"]
    $btn.Tag = $name
    $btn.Add_Click({ param($s, $e) if (-not $script:UI.SuppressToggleEvents) { Set-Toggle -Name $s.Tag -Value ([bool]$s.IsChecked) } })
}

# --- buttons ---
$window.FindName('MinimizeButton').Add_Click({ Switch-Minimized })
$window.FindName('MinimizeBorder').Add_MouseDown({ param($s, $e) if ($e.ChangedButton -eq 'Right') { $script:UI.Window.Close() } })
$window.FindName('btnBack').Add_Click({ Invoke-ViewBack })
$window.FindName('btnForward').Add_Click({ Invoke-ViewForward })
$window.FindName('btnEditMarkdown').Add_Click({ Open-NoteInEditor })
$dblHint = switch ($script:App.EditOnDoubleClick) { 'ctrl' { Get-Text 'ui.editDblCtrl' } 'always' { Get-Text 'ui.editDblAlways' } default { '' } }
$window.FindName('btnEditMarkdown').ToolTip = (Get-Text 'ui.edit') + $dblHint
$window.FindName('btnAddMarkdown').Add_Click({ New-NoteForCurrentContext })

# --- search box + popup ---
# typing is debounced: the search runs 80 ms after the last keystroke instead of on every one
$searchDebounce = [System.Windows.Threading.DispatcherTimer]::new()
$searchDebounce.Interval = [TimeSpan]::FromMilliseconds(80)
$searchDebounce.Add_Tick({ param($s, $e) $s.Stop(); try { Update-SearchPopup -Query $script:UI.SearchBox.Text } catch { Write-CSCLog "Search: $_" 'Warn' } })
$script:UI.SearchDebounce = $searchDebounce
$script:UI.SearchBox.Add_TextChanged({
        param($s, $e)
        $script:UI.SearchHint.Visibility = $(if ([string]::IsNullOrEmpty($s.Text)) { 'Visible' } else { 'Collapsed' })
        if ($script:UI.SuppressSearchEvents) { return }
        $script:UI.SearchDebounce.Stop(); $script:UI.SearchDebounce.Start()
    })
$script:UI.SearchBox.Add_PreviewKeyDown({
        param($s, $e)
        switch ($e.Key) {
            'Down'   { Move-SearchSelection 1; $e.Handled = $true }
            'Up'     { Move-SearchSelection -1; $e.Handled = $true }
            'Return' {
                $e.Handled = $true
                if ($script:UI.SearchPopup.IsOpen -and $script:UI.SearchList.SelectedItem) { Open-SearchResult -Result $script:UI.SearchList.SelectedItem.Tag }
                elseif ($s.Text) { Invoke-NoteSearch -Query $s.Text; try { $script:UI.WebView.Focus() | Out-Null } catch { } }
            }
            'Escape' {
                $e.Handled = $true
                if ($script:UI.SearchPopup.IsOpen) { $script:UI.SearchPopup.IsOpen = $false }
                else { $script:UI.SuppressSearchEvents = $true; $s.Text = ''; $script:UI.SuppressSearchEvents = $false; $script:UI.SearchHint.Visibility = 'Visible'; try { $script:UI.WebView.Focus() | Out-Null } catch { } }
            }
        }
    })
$script:UI.SearchBox.Add_GotKeyboardFocus({ param($s, $e) if ($s.Text -and -not $script:UI.SearchPopup.IsOpen) { $s.SelectAll() } })
$window.Add_PreviewKeyDown({
        param($s, $e)
        $mods = [System.Windows.Input.Keyboard]::Modifiers
        $ctrl = [bool]($mods -band [System.Windows.Input.ModifierKeys]::Control)
        if ($ctrl -and $e.Key -eq 'K') { $e.Handled = $true; Show-Search; return }
        # tab shortcuts while the WPF chrome has the focus (inside the page app.js handles them): Ctrl+1..9, Ctrl+Tab
        if ($ctrl -and $script:UI.WebReady) {
            $js = ''
            if ($e.Key -match '^(D|NumPad)([1-9])$') { $js = 'if(window.CSC_activateTab){window.CSC_activateTab(' + ([int]$matches[2] - 1) + ');}' }
            elseif ($e.Key -eq 'Tab') { $js = 'if(window.CSC_cycleTab){window.CSC_cycleTab(' + $(if ($mods -band [System.Windows.Input.ModifierKeys]::Shift) { '-1' } else { '1' }) + ');}' }
            if ($js) { $e.Handled = $true; try { $null = $script:UI.WebView.CoreWebView2.ExecuteScriptAsync($js) } catch { } }
        }
    })
$window.Add_LocationChanged({
        try {
            $script:UI.SearchPopup.IsOpen = $false
            if ($script:UI.Window.Top -lt 35 -and $script:UI.Window.Top -ne 0) { $script:UI.Window.Top = 0 }
        }
        catch { }
    })

# --- WebView2 ---
$creation = [Microsoft.Web.WebView2.Wpf.CoreWebView2CreationProperties]::new()
$creation.UserDataFolder = Join-Path $script:App.DataDir 'WebView2'
$script:UI.WebView.CreationProperties = $creation
$script:UI.WebView.Add_CoreWebView2InitializationCompleted({
        param($s, $e)
        if (-not $e.IsSuccess) {
            Write-CSCLog ("WebView2 init failed: {0}" -f $e.InitializationException) 'Error'
            Set-Status (Get-Text 'status.webviewFailed')
            return
        }
        try { Initialize-WebViewCore -Core $s.CoreWebView2 } catch { Write-CSCLog "Initialize-WebViewCore: $_" 'Error' }
    })
$script:UI.WebView.Add_ZoomFactorChanged({ param($s, $e) try { $script:State.Zoom = [double]$s.ZoomFactor } catch { } })

# --- foreground window polling (UI thread, cheap Win32 calls) ---
$fgTimer = [System.Windows.Threading.DispatcherTimer]::new()
$fgTimer.Interval = [TimeSpan]::FromMilliseconds($script:App.PollIntervalMs)
$fgTimer.Add_Tick({
        try {
            if ($script:App.InactiveOpacity -gt 0) { Update-WindowOpacity }
            $hwnd = [CheatSheetCompanionHelper.Native]::GetForegroundWindow()
            if ($hwnd -eq [IntPtr]::Zero -or $hwnd -eq $script:UI.Hwnd) { return }
            $title = [CheatSheetCompanionHelper.Native]::GetWindowTitle($hwnd)
            if ([string]::IsNullOrWhiteSpace($title)) { return }
            $procId = [CheatSheetCompanionHelper.Native]::GetWindowProcessId($hwnd)
            if ($procId -eq $PID) { return }
            $proc = $script:UI.ProcessNames[$procId]
            if ($null -eq $proc) {
                $proc = try { [System.Diagnostics.Process]::GetProcessById($procId).ProcessName } catch { '' }
                $script:UI.ProcessNames[$procId] = $proc
                if ($script:UI.ProcessNames.Count -gt 500) { $script:UI.ProcessNames.Clear() }
            }
            $key = "$procId|$title"
            $fg = $script:UI.LastForeground
            if ($key -ne $fg.Key) { $fg.Key = $key; $fg.Title = $title; $fg.Process = $proc; $fg.Stable = 0; return }
            $fg.Stable++
            if ($fg.Stable -lt 1) { return }
            if (-not $script:State.FollowWindow -or $key -eq $script:UI.ForegroundMatchedKey) { return }
            $script:UI.ForegroundMatchedKey = $key
            $script:UI.LastWindow = @{ Text = $title; Time = [DateTime]::UtcNow }
            # a note found via the clipboard keeps priority while the user stays in the program it was copied from
            # (terminal and editor titles change constantly); switching to another program re-enables window matching
            if ($script:State.FollowClipboard -and $script:UI.ClipboardProcessId -ne 0 -and $procId -eq $script:UI.ClipboardProcessId) { return }
            $script:UI.ClipboardProcessId = 0
            $targets = @(
                @{ Text = $title; Origin = 'title'; Threshold = $script:App.ThresholdTitle }
                @{ Text = $proc; Origin = 'process'; Threshold = $script:App.ThresholdProcess }
            )
            Show-MatchesFor -Targets $targets -SourceText ('{0} | {1}' -f $title, $proc) -StatusPrefix (Get-Text 'status.prefixWindow') -Pinned @(Get-RuleCandidates -Title $title -Process $proc)
        }
        catch { Write-CSCLog "Foreground tick: $_" 'Warn' }
    })
$script:UI.Timers += $fgTimer

# --- vault watcher polling ---
$script:UI.Watcher = [CheatSheetCompanionHelper.VaultWatcher]::new()
$vaultTimer = [System.Windows.Threading.DispatcherTimer]::new()
$vaultTimer.Interval = [TimeSpan]::FromMilliseconds(400)
$vaultTimer.Add_Tick({
        try {
            if (-not $script:UI.Watcher.TakeDirtyIfQuiet(300)) { return }
            $before = @{}
            if ($script:UI.CurrentView) { foreach ($c in $script:UI.CurrentView.Candidates) { $before[$c.Entry.Rel] = $(if (Test-Path -LiteralPath $c.Entry.Html) { (Get-Item -LiteralPath $c.Entry.Html).LastWriteTimeUtc } else { [datetime]0 }) } }
            $setChanged = Update-VaultIndex
            if ($script:UI.CurrentView) {
                $changed = $false
                $fresh = @()
                foreach ($c in $script:UI.CurrentView.Candidates) {
                    $e = $script:Index.ByPath[$c.Entry.PathKey]
                    if ($null -eq $e) { $changed = $true; continue }
                    $fresh += [pscustomobject]@{ Entry = $e; Score = $c.Score; Origins = $c.Origins; Reasons = $(if ($c.PSObject.Properties['Reasons']) { $c.Reasons } else { @() }) }
                    if ((Get-Item -LiteralPath $e.Html).LastWriteTimeUtc -ne $before[$c.Entry.Rel]) { $changed = $true }
                }
                if ($changed -and $fresh.Count -gt 0) {
                    Write-CSCLog 'Vault changed -> re-render current view'
                    Show-Candidates -Candidates $fresh -SourceText $script:UI.CurrentView.SourceText -NoHistory -ActiveEntry $script:UI.CurrentNote
                }
            }
            if ($setChanged) { Set-Status (Get-Text 'status.vaultUpdated' -Format $script:Index.Notes.Count) }
        }
        catch { Write-CSCLog "Vault tick: $_" 'Warn' }
    })
$script:UI.Timers += $vaultTimer

# --- tray icon ---
function New-TrayIcon {
    $bmp = [System.Drawing.Bitmap]::new(32, 32)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    $g.FillEllipse([System.Drawing.SolidBrush]::new([System.Drawing.ColorTranslator]::FromHtml('#7852ee')), 1, 1, 30, 30)
    $font = [System.Drawing.Font]::new('Segoe UI', 15, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $fmt = [System.Drawing.StringFormat]::new(); $fmt.Alignment = 'Center'; $fmt.LineAlignment = 'Center'
    $g.DrawString('CS', $font, [System.Drawing.Brushes]::White, [System.Drawing.RectangleF]::new(0, 0, 32, 32), $fmt)
    $g.Dispose()
    return [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
}

function Invoke-TrayAction {
    param([string]$Action)
    switch ($Action) {
        'show'            { Switch-Minimized -Restore; try { $script:UI.Window.Activate() | Out-Null } catch { } }
        'toggle'          { Switch-Minimized }
        'FollowWindow'    { Set-Toggle -Name FollowWindow -Value (-not $script:State.FollowWindow) }
        'FollowClipboard' { Set-Toggle -Name FollowClipboard -Value (-not $script:State.FollowClipboard) }
        'AutoCopy'        { Set-Toggle -Name AutoCopy -Value (-not $script:State.AutoCopy) }
        'Pinned'          { Set-Toggle -Name Pinned -Value (-not (Test-NotePinned -Entry $script:UI.CurrentNote)) }
        'search'          { Show-Search }
        'exit'            { $script:UI.Window.Close() }
    }
}

try {
    $tray = [System.Windows.Forms.NotifyIcon]::new()
    $tray.Icon = New-TrayIcon
    $tray.Text = 'CheatSheetCompanion'
    $menu = [System.Windows.Forms.ContextMenuStrip]::new()
    $trayHandler = { param($s, $e) Invoke-TrayAction -Action ([string]$s.Tag) }
    foreach ($def in @(
            @{ Text = (Get-Text 'tray.show'); Tag = 'show' }, @{ Text = (Get-Text 'tray.toggle'); Tag = 'toggle' }, @{ Text = (Get-Text 'tray.search'); Tag = 'search' }, @{ Text = '-' },
            @{ Text = (Get-Text 'tray.pin'); Tag = 'Pinned'; Check = $true }, @{ Text = (Get-Text 'tray.followWindow'); Tag = 'FollowWindow'; Check = $true }, @{ Text = (Get-Text 'tray.followClipboard'); Tag = 'FollowClipboard'; Check = $true }, @{ Text = (Get-Text 'tray.autoCopy'); Tag = 'AutoCopy'; Check = $true },
            @{ Text = '-' }, @{ Text = (Get-Text 'tray.exit'); Tag = 'exit' })) {
        if ($def.Text -eq '-') { [void]$menu.Items.Add([System.Windows.Forms.ToolStripSeparator]::new()); continue }
        $mi = [System.Windows.Forms.ToolStripMenuItem]::new($def.Text)
        $mi.Tag = $def.Tag
        $mi.add_Click($trayHandler)
        if ($def.Check) { $mi.Checked = ($def.Tag -ne 'Pinned' -and [bool]$script:State[$def.Tag]); $script:UI.TrayItems[$def.Tag] = $mi }
        [void]$menu.Items.Add($mi)
    }
    $tray.ContextMenuStrip = $menu
    $tray.add_MouseClick({ param($s, $e) if ($e.Button -eq 'Left') { Invoke-TrayAction -Action 'toggle' } })
    $tray.Visible = $true
    $script:UI.Tray = $tray
}
catch { Write-CSCLog "Tray icon: $_" 'Warn' }

# --- hotkeys + clipboard listener (need the window handle) ---
function ConvertTo-HotKeySpec {
    <# "ControlKey,ShiftKey,F11" -> @{ Modifiers; Key } (names from System.Windows.Forms.Keys, Alt/Ctrl/Win accepted) #>
    param([AllowEmptyString()][string]$Spec)
    if ([string]::IsNullOrWhiteSpace($Spec)) { return $null }
    $mods = 0x4000; $vk = 0   # MOD_NOREPEAT
    foreach ($part in ($Spec -split ',')) {
        $p = $part.Trim()
        if (-not $p) { continue }
        switch -regex ($p) {
            '^(L|R)?(Control|ControlKey|Ctrl)$' { $mods = $mods -bor 0x0002; break }
            '^(L|R)?(Shift|ShiftKey)$'          { $mods = $mods -bor 0x0004; break }
            '^(L|R)?(Menu|Alt)$'                { $mods = $mods -bor 0x0001; break }
            '^(L|R)?Win$'                       { $mods = $mods -bor 0x0008; break }
            default { try { $vk = [int][System.Windows.Forms.Keys]$p } catch { Write-CSCLog "Unknown key '$p' in hotkey '$Spec'" 'Warn' } }
        }
    }
    if ($vk -eq 0) { return $null }
    return @{ Modifiers = [uint32]$mods; Key = [uint32]$vk; Spec = $Spec }
}

$window.Add_SourceInitialized({
        try {
            $helper = [System.Windows.Interop.WindowInteropHelper]::new($script:UI.Window)
            $script:UI.Hwnd = $helper.Handle
            $source = [System.Windows.Interop.HwndSource]::FromHwnd($helper.Handle)
            $sink = [CheatSheetCompanionHelper.MessageSink]::new()
            $sink.KeepLayered = [bool]$script:App.UseWindowAlpha
            $script:UI.Sink = $sink
            $source.AddHook([System.Windows.Interop.HwndSourceHook]$sink.Hook)
            [CheatSheetCompanionHelper.Native]::SetDarkTitleBar($script:UI.Hwnd, (Get-EffectiveTheme).IsDark)
            Update-WindowOpacity

            $sink.add_HotKeyPressed([Action[int]] {
                    param($id)
                    try {
                        switch ($id) {
                            1 { Switch-Minimized }
                            2 { Set-Toggle -Name FollowWindow -Value (-not $script:State.FollowWindow) }
                            3 { Set-Toggle -Name FollowClipboard -Value (-not $script:State.FollowClipboard) }
                            4 { Show-Search }
                            5 { Invoke-ClipboardLookup -Force }
                        }
                    }
                    catch { Write-CSCLog "HotKey $id : $_" 'Warn' }
                })
            $sink.add_ClipboardChanged([Action] {
                    try { Invoke-ClipboardLookup } catch { Write-CSCLog "Clipboard: $_" 'Warn' }
                })

            $specs = @(
                @{ Id = 1; Spec = (Get-Setting -Section GlobalShortcut -Key DisplayCompanion -Default '') }
                @{ Id = 2; Spec = (Get-Setting -Section GlobalShortcut -Key AutoSwitch -Default '') }
                @{ Id = 3; Spec = (Get-Setting -Section GlobalShortcut -Key ClipboardSwitch -Default '') }
                @{ Id = 4; Spec = (Get-Setting -Section GlobalShortcut -Key Search -Default '') }
                @{ Id = 5; Spec = (Get-Setting -Section GlobalShortcut -Key ClipboardLookup -Default '') }
            )
            foreach ($hk in $specs) {
                $parsed = ConvertTo-HotKeySpec -Spec $hk.Spec
                if ($null -eq $parsed) { continue }
                if ([CheatSheetCompanionHelper.Native]::RegisterHotKey($script:UI.Hwnd, $hk.Id, $parsed.Modifiers, $parsed.Key)) {
                    $script:UI.HotKeys += $hk.Id
                    Write-CSCLog ("HotKey {0} registered: {1}" -f $hk.Id, $hk.Spec)
                }
                else { Write-CSCLog ("HotKey '{0}' could not be registered (already in use?)" -f $hk.Spec) 'Warn' }
            }
            if (-not [CheatSheetCompanionHelper.Native]::AddClipboardFormatListener($script:UI.Hwnd)) { Write-CSCLog 'AddClipboardFormatListener failed' 'Warn' }
        }
        catch { Write-CSCLog "SourceInitialized: $_" 'Error' }
    })

$window.Add_Loaded({
        if ($AutoCloseSeconds -gt 0) {
            $closeTimer = [System.Windows.Threading.DispatcherTimer]::new()
            $closeTimer.Interval = [TimeSpan]::FromSeconds($AutoCloseSeconds)
            $closeTimer.Add_Tick({ param($s, $e) $s.Stop(); Write-CSCLog 'AutoClose'; $script:UI.Window.Close() })
            $closeTimer.Start()
            $script:UI.Timers += $closeTimer
            if ($ScreenshotPath) {
                $shotTimer = [System.Windows.Threading.DispatcherTimer]::new()
                $shotTimer.Interval = [TimeSpan]::FromSeconds([Math]::Max(1, $AutoCloseSeconds - 2))
                $shotTimer.Add_Tick({
                        param($s, $e)
                        $s.Stop()
                        try { if ($script:UI.WebReady) { Save-WebViewScreenshot -Path $ScreenshotPath } } catch { Write-CSCLog "Screenshot: $_" 'Warn' }
                    })
                $shotTimer.Start()
                $script:UI.Timers += $shotTimer
            }
        }
        try {
            Update-WindowTheme
            $null = $script:UI.WebView.EnsureCoreWebView2Async()
            Show-Welcome
            $script:UI.Watcher.Start($script:App.Vault)
            foreach ($t in $script:UI.Timers) { if (-not $t.IsEnabled) { $t.Start() } }
            Set-Status (Get-Text 'status.start' -Format $script:App.VaultName, $script:Index.Notes.Count, (Get-Text $(if ($script:State.FollowWindow) { 'status.on' } else { 'status.off' })), (Get-Text $(if ($script:State.FollowClipboard) { 'status.on' } else { 'status.off' })))
        }
        catch { Write-CSCLog "Loaded: $_" 'Error' }
    })

$window.Add_Closing({
        try {
            foreach ($t in $script:UI.Timers) { try { $t.Stop() } catch { } }
            try { $script:UI.Watcher.Stop() } catch { }
            foreach ($id in $script:UI.HotKeys) { [void][CheatSheetCompanionHelper.Native]::UnregisterHotKey($script:UI.Hwnd, $id) }
            if ($script:UI.Hwnd -ne [IntPtr]::Zero) { [void][CheatSheetCompanionHelper.Native]::RemoveClipboardFormatListener($script:UI.Hwnd) }
            if ($script:UI.Tray) { try { $script:UI.Tray.Visible = $false; $script:UI.Tray.Dispose() } catch { } }
            Save-State
        }
        catch { Write-CSCLog "Closing: $_" 'Warn' }
    })

#endregion

#region ---------- Run ----------

Initialize-Vault -Name $(if ($script:State.Vault) { [string]$script:State.Vault } else { [string]$script:App.Vaults.Keys[0] })

[void]$window.ShowDialog()

try { $script:UI.WebView.Dispose() } catch { }
Write-CSCLog 'Exit'

#endregion
