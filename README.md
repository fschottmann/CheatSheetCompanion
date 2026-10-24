# CheatSheetCompanion

*English version: [README.en.md](README.en.md)*

## Über das Programm

Cheat Sheet Companion zeigt kontextabhängig Notizen aus einem Markdown-Ordner oder Obsidian-Vault an.
Die passende Notiz wird über die Ähnlichkeit zwischen Notizname und

- dem Titel des aktiven Fensters,
- dem Prozessnamen des aktiven Fensters oder
- dem zuletzt kopierten Text (Zwischenablage, z. B. aus RDP- oder SSH-Sitzungen)

ermittelt. Die besten Kandidaten (`MaxCandidates`, Vorgabe 5) werden als Tabs mit Ähnlichkeitswert (0.00 bis 1.00) angezeigt.

## Screenshots

Windows Terminal ist aktiv: die Notiz *PowerShell* wird über ihren Alias gefunden (1.00), zwei weitere Kandidaten stehen als Tabs bereit, die Statuszeile erklärt den Treffer.

![Windows Terminal aktiv: PowerShell-Notiz mit Kandidaten-Tabs](docs/Screenshot_WindowsTerminal.png)

Ein Browser (hier Brave) ist aktiv: die Notiz *Microsoft Edge* deckt per Alias alle Browser ab.

![Browser aktiv: Notiz Microsoft Edge](docs/Screenshot_Browser.png)

Angeheftete Notiz (📌) mit Callout, Tabelle und den Leisten TOC, COPIED und LINKED am rechten Rand.

![Angeheftete Hilfe-Notiz im Dark Theme](docs/Screenshot_pinned_note.png)

## Voraussetzungen

- PowerShell 7 x64 ([Download](https://github.com/PowerShell/PowerShell))
- WebView2 Runtime (auf Windows 11 vorinstalliert, sonst [Evergreen Runtime](https://developer.microsoft.com/microsoft-edge/webview2/))
- Die WebView2-Bibliotheken liegen im Ordner `dll\` bei.

## Start

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\CheatSheetCompanion.ps1
```

Mehr ist nicht nötig: `ini\config.ini` zeigt auf den mitgelieferten Demo-Vault `vault\`, der jede Funktion vorführt und die
Hilfe als Notiz *Help* enthält (Notepad, Windows Terminal oder einen Browser öffnen, und die passende Notiz erscheint).
Für den eigenen Vault `MarkDownFolder` in `ini\config.ini` auf den Ordner setzen (absolut oder relativ zum Skript) und bei
Bedarf `Language` (Vorgabe `auto` = Windows-Anzeigesprache).

Optionen:

| Parameter | Bedeutung |
| --- | --- |
| `-OpenNote "Name"` | Zeigt diese Notiz direkt nach dem Start. |
| `-Verbose` | Schreibt ein ausführliches Log nach `%LocalAppData%\CheatSheetCompanion\csc.log`. |
| `-AutoCloseSeconds n` | Testhilfe: schließt das Fenster nach n Sekunden. |
| `-ScreenshotPath pfad.png` | Testhilfe: speichert die gerenderte Seite als PNG (`*.window.png` = ganzes Fenster). |
| `-SearchQuery "text"` | Testhilfe: füllt das Suchfeld und öffnet die Trefferliste. |

## Bedienung

| Element | Funktion |
| --- | --- |
| Leiste oben (▲▲▲) | Klick: Fenster auf einen schmalen Streifen einklappen bzw. ausklappen. Rechtsklick: beenden. |
| Suchfeld (Strg+K) | Fuzzy-Suche wie der Quick Switcher in Obsidian: Namen zuerst, dann Inhalte mit Textausschnitt. Pfeiltasten wählen, Enter öffnet, Escape schließt. |
| ◀ ▶ | Zurück und vorwärts in den Ansichten (auch Alt+Links / Alt+Rechts oder die Maus-Daumentasten). |
| Tabs | Kandidaten wechseln per Klick, Strg+1 bis Strg+9 oder Strg+Tab / Strg+Umschalt+Tab. |
| ✎ | Angezeigte Notiz in Obsidian bzw. im Standardeditor öffnen. Im Text: Strg+Doppelklick (`EditOnDoubleClick`), ein normaler Doppelklick markiert nur das Wort. |
| + | Neue Notiz anlegen. Der Dialog schlägt den Namen aus der jüngsten Quelle vor (Fenstertitel, Zwischenablage oder Markierung im Companion), umschaltbar, hart gekürzt auf `NewNoteNameMaxChars`. Zwischenablage oder Markierung können als Inhalt übernommen werden, der Zielordner ist wählbar (Vorgabe `NewNoteFolder`), `NewNoteTemplate` liefert die Vorlage. |
| Codeblock | Beim Überfahren erscheint oben rechts *Kopieren*: kopiert den ganzen Block ohne Markieren. |
| KOPIERT (rechter Rand) | Ausfahrbare Leiste wie das TOC mit allem, was über den Companion kopiert wurde (Codeblock, Markierung, Auto Copy, Kontextmenü), `CopyHistory` Einträge, Vorgabe 10. Ein Klick auf einen Eintrag kopiert ihn erneut und schließt die Leiste; ein Klick woanders hin, Escape oder ein Fensterwechsel schließen sie ebenfalls (gilt auch für das TOC). Die Leiste erscheint nur, wenn es Einträge gibt, und bleibt über einen Neustart erhalten (`state.json`, siehe unten); *Leeren* in der Leiste löscht den Verlauf. |
| Kein Treffer | Standard: die angezeigte Notiz bleibt stehen, die Statuszeile meldet den fehlenden Treffer. Mit `NoMatchBehavior = suggest` erscheint stattdessen "Noch keine Notiz dazu" mit Knopf *Notiz anlegen* und den ähnlichsten Notizen. |
| Startseite | Zeigt die zuletzt angezeigten Notizen des Vaults (`RecentNotes`). |
| **Pin** | Angezeigte Notiz anheften, mehrere möglich: Angeheftete Notizen bleiben als farblich markierte Tabs (📌) vorne stehen, das Matching läuft weiter und der beste Treffer wird wie gewohnt angezeigt. `MaxCandidates` gilt insgesamt, ein Platz bleibt immer für Treffer. Der Schalter zeigt, ob die gerade sichtbare Notiz angeheftet ist. Bleibt über einen Neustart erhalten (auch im Tray-Menü und Kontextmenü). |
| **Fenster** | Notizen passend zum aktiven Fenster anzeigen (Titel und Prozess). |
| **Clip** | Notizen passend zum kopierten Text anzeigen. Eine so gefundene Notiz bleibt stehen, solange man im Programm bleibt, aus dem kopiert wurde (Titelwechsel in Terminals und Editoren überschreiben sie nicht); ein Programmwechsel schaltet wieder auf das Fenster um. |
| **AC** | Auto Copy: markierten Text beim Loslassen der Maustaste kopieren, wenn die Taste mindestens `AutoCopyMinHoldMs` gehalten wurde (0 = sofort). |
| Rechtsklick im Notiztext | Eigene Einträge oben im Browser-Kontextmenü: *In Zwischenablage kopieren* (bei Markierung, reiner Text), *Kopieren als…* (Markierung formatiert, ganze Notiz als Text oder formatiert; formatiert = HTML plus Text, damit Teams und Outlook Listen und Codeblöcke behalten, auch Strg+Umschalt+C), *Notiz bearbeiten*, *Zwischenablage an Notiz anhängen* (kopierter Text wird unverändert ans Ende der angezeigten Notiz geschrieben, etwa Gesprächsnotizen oder ein Absatz von einer Webseite), *Alias hinzufügen* (Fenstertitel, Prozessname, Zwischenablage, Markierung oder eigener Text als Alias ins Frontmatter der angezeigten Notiz), *Zur Zwischenablage nachschlagen* (einmalig, auch bei ausgeschaltetem Clip), *Notiz anheften*, *Theme* und bei mehreren Vaults *Vault*. |
| Tab-Tooltip / Statuszeile | Trefferbegründung (`ShowMatchReason`): Name im Text enthalten, ähnliches Wort des Titels, Alias oder Regel. |
| Eingeklappter Streifen | Zeigt den Namen der aktuellen Notiz mit Wert, bei Anheftung mit Hinweis. |
| Klick auf einen Tag | Sucht den Tag in allen Notizen. |
| TOC | Inhaltsverzeichnis der aktiven Notiz ein- und ausblenden. |
| VERLINKT (rechter Rand) | Ausfahrbare Leiste wie das TOC mit den verlinkten Erwähnungen wie in Obsidian: Notizen, die per `[[Wikilink]]`, Einbettung oder Markdown-Link auf die aktive Notiz verlinken, mit der Zeile der Erwähnung (Codeblöcke zählen nicht). Klick öffnet die Notiz. Der Griff erscheint nur, wenn es Backlinks gibt; `ShowBacklinks = false` schaltet das ab. |
| Durchsichtigkeit | `InactiveOpacity` macht das Fenster durchscheinend, solange ein anderes Programm aktiv ist, damit man sieht, was darunter liegt. Unter der Maus, eingeklappt oder bei offener Suche ist es wieder voll sichtbar; `0` schaltet das ab, `WindowOpacity` ist die normale Deckkraft. |
| Tray-Symbol | Linksklick: ein-/ausklappen. Rechtsklick: Anzeigen, Suchen, Anheften, die drei Schalter, Beenden. |

Globale Tastenkürzel (siehe `ini\config.ini`): Fenster ein-/ausklappen, Fenster-Modus, Zwischenablage-Modus, Suche, einmaliges Nachschlagen zur Zwischenablage.

### Sprache

Alle Texte der Oberfläche (Schalter, Statuszeile, Menüs, Dialoge, Seiten) kommen aus `lang\<code>.ini`. Mitgeliefert sind `de.ini`
und `en.ini`; `Language = auto` nimmt die Windows-Anzeigesprache, sonst `de`, `en` oder ein eigener Code. Für eine weitere Sprache
`lang\de.ini` nach `lang\fr.ini` kopieren, übersetzen und `Language = fr` setzen. Fehlende Schlüssel fallen auf die eingebauten
deutschen Texte zurück, `{0}`, `{1}` sind Platzhalter, `\n` ein Zeilenumbruch. Log und Konfiguration bleiben unverändert.

## Darstellung (Obsidian)

Unterstützt werden Wikilinks `[[Notiz]]`, `[[Notiz#Überschrift|Alias]]`, Einbettungen `![[Bild.png|300]]` und `![[Notiz]]`,
Callouts `> [!warning]- Titel`, Tags `#tag`, Hervorhebungen `==Text==`, Kommentare `%% %%`, YAML-Frontmatter
(optional als Kasten mit `ShowFrontmatter = true`), Tabellen, Aufgabenlisten, Fußnoten, Syntax-Highlighting in Codeblöcken
(highlight.js) sowie weiche Zeilenumbrüche wie in Obsidian (abschaltbar mit `StrictLineBreaks = true`).

Interne Links werden im Companion geöffnet. Bilder und relative Pfade werden gegen den Vault aufgelöst
(zuerst exakter Pfad, dann relativ zur Notiz, dann kürzester Pfad mit gleichem Dateinamen, wie in Obsidian).
Die Leiste *VERLINKT* am rechten Rand listet die Notizen auf, die auf die aktive Notiz verlinken (`ShowBacklinks`).

### Themes

`Theme = auto` folgt dem Windows-Modus (hell/dunkel), alternativ `light`, `dark` oder der Dateiname eines eigenen Themes
aus `themes\` (z. B. `github`). Umschalten per Rechtsklick > Theme; die Wahl wird gespeichert. Ein Theme überschreibt nur die
CSS-Variablen aus `assets\app.css` (siehe `themes\dark.css` als Vorlage). Eigene Feinkorrekturen gehören in `styles\userstyle.css`,
die zuletzt geladen wird. Die Leiste des Fensters folgt dem Theme.

## Konfiguration (`ini\config.ini`)

```ini
[GlobalShortcut]
AutoSwitch = ControlKey,ShiftKey,F2        ; Fenster-Modus umschalten
DisplayCompanion = ControlKey,ShiftKey,F11 ; Fenster ein-/ausklappen
ClipboardSwitch = ControlKey,ShiftKey,F3   ; Zwischenablage-Modus umschalten
Search =                                   ; Suche fokussieren (leer = kein globaler Hotkey, Strg+K im Fenster geht immer)
ClipboardLookup =                          ; einmalig zur Zwischenablage nachschlagen, unabhängig vom Clip-Schalter

[Configuration]
MarkDownFolder = vault                             ; Standard-Vault, absolut oder relativ zum Skriptordner (Vorgabe: Demo-Vault)
SourceLink = obsidian://open?vault={1}&file={0}    ; {0} = Notizpfad ohne .md, {1} = Vault-Name; leer = Standardprogramm
NewNoteFolder = Inbox                              ; Unterordner für neue Notizen (leer = Vault-Wurzel)
NewNoteTemplate = _templates\note.md               ; Vorlage: {{title}}, {{date}}, {{time}}, {{content}}; Templater-Datum wird ersetzt, übrige <% %> entfernt
NewNoteNameMaxChars = 60                           ; harte Kürzung des Namensvorschlags
IgnoreFolders = _templates                         ; zusätzlich zu .obsidian, .trash, .git
FollowWindow = true
FollowClipboard = false
AutoCopy = false
Language = auto                                    ; auto (Windows-Anzeigesprache) | de | en | eigener Code mit lang\<code>.ini
Theme = auto
ShowFrontmatter = false
EditOnDoubleClick = ctrl                           ; ctrl | always | off
NoMatchBehavior = keep                             ; keep = Notiz bleibt stehen, suggest = Vorschlagsseite
RecentNotes = 8                                    ; zuletzt angezeigte Notizen auf der Startseite (0 = aus)
ShowMatchReason = true                             ; Trefferbegründung in Statuszeile und Tab-Tooltip
MatchIgnoreFolders = Daily, Besprechungen          ; nicht automatisch treffen, aber suchbar; einzelne Notiz: Frontmatter "csc: ignore"
CopyHistory = 10                                   ; zuletzt kopierte Texte merken (0 = aus)
ShowBacklinks = true                               ; Leiste VERLINKT (Notizen, die hierher verlinken)
WindowOpacity = 0.92                               ; Deckkraft des Fensters (0.3 bis 1)
InactiveOpacity = 0.6                              ; Deckkraft, solange ein anderes Programm aktiv ist (0 = aus)
SimilarityThreshold = 0.25
SimilarityThresholdTitle = 0.25
SimilarityThresholdProcess = 0.25
SimilarityThresholdClipboard = 0.4
MaxCandidates = 5                                  ; angezeigte Kandidaten (Tabs)
NGram = 3
PollIntervalMs = 500
ClipboardMaxChars = 400
SearchMaxResults = 12
SearchContent = true                               ; false = Notiztexte nach dem Indizieren verwerfen (nur Namen/Aliase suchbar)
AutoCopyMinHoldMs = 150
StrictLineBreaks = false
Debug = false

[Vaults]
; weitere Vaults: Name = Pfad (Name = Obsidian-Vault-Name), Umschalten per Rechtsklick > Vault
myNotes = D:\pfad\zum\anderen\vault

[Rules]
; feste Zuordnung Programm -> Notiz (immer zuerst, Wert 1.00): Prozessname oder title:<Regex auf den Fenstertitel>
mstsc = RDP
title:.*Visual Studio Code$ = VSCode
```

Schalter, Anheftung, Theme, Vault, Zoomstufe, zuletzt angezeigte Notizen, zuletzt kopierte Texte sowie Fensterposition und -größe werden in
`%LocalAppData%\CheatSheetCompanion\state.json` gespeichert und beim nächsten Start übernommen.

## Funktionsweise

- Beim Start werden alle Notizen nach `Output\<Vault>\` konvertiert; danach nur noch geänderte (Cache nach Änderungszeit).
- Ein Dateiwächter erkennt neue, geänderte und gelöschte Notizen und aktualisiert Index und Anzeige.
- Das aktive Fenster wird alle `PollIntervalMs` Millisekunden geprüft; ein Wechsel löst die Suche aus, sobald der Titel kurz stabil ist.
- Die Ähnlichkeit ist ein Dice-Koeffizient über Zeichen-N-Gramme (in C#, kulturunabhängig). Enthält der Text den Notiznamen, zählt das als 1.00.
- Aliase aus dem Frontmatter (`aliases: [Browser, Chrome]`) zählen wie Notiznamen. So lässt sich eine Notiz gezielt an Fenstertitel oder Prozessnamen binden; der Companion schreibt sie auf Wunsch selbst (Rechtsklick > Alias hinzufügen).
- `[Rules]` ordnen ein Programm fest einer Notiz zu, unabhängig von der Ähnlichkeit. Aliase gehören zur Notiz (wandern mit dem Vault), Regeln zur Installation.
- Notizen in `MatchIgnoreFolders` oder mit `csc: ignore` im Frontmatter (etwa Daily Notes, Besprechungen) liefern keine Zufallstreffer, bleiben aber über die Suche erreichbar.
- Zwischenablage: Windows meldet jede Änderung (kein Polling). Ist die Zwischenablage gerade gesperrt (Verlaufsfunktion, andere Programme), wird kurz darauf erneut gelesen. Eigene Kopien des Companions lösen keine Suche aus.
- Es laufen keine Hintergrund-Runspaces; alle Ereignisse werden auf dem UI-Thread verarbeitet (Timer, Hotkeys, Zwischenablage, WebView2-Nachrichten).
- Speicher: Die C#-Helfer liegen fertig übersetzt in `dll\CheatSheetCompanionHelper.dll`; das Skript lädt sie wie die WebView2-DLLs und kompiliert zur Laufzeit nichts und schreibt keine Dateien (nur wenn die DLL fehlt oder nicht zum Quelltext passt, wird wie früher im Speicher kompiliert, das kostet rund 50 MB). Nach Änderungen am C#-Teil `tools\Build-Helper.ps1` ausführen. Im eingeklappten Zustand wird der WebView2-Renderer angehalten (wie ein Hintergrund-Tab) und der Prozess gibt ungenutzte Seiten frei; beim Ausklappen läuft er weiter. Der Verlauf ist begrenzt.
- Index: Notizname, Aliase, der Ignore-Hinweis und die Backlinks werden beim Indizieren aus den Dateien gelesen. Der Text selbst bleibt nur für die Inhaltssuche im Speicher (Größe des Markdown mal zwei, bei den meisten Vaults wenige MB); `SearchContent = false` verwirft ihn, die Suche findet dann nur Namen und Aliase.

## Projektstruktur

| Pfad | Inhalt |
| --- | --- |
| `CheatSheetCompanion.ps1` | Hauptskript (WPF, WebView2, Index, Matching, Rendering, Suche) |
| `assets\app.css`, `assets\app.js` | Basis-Theme und Seitenlogik im WebView |
| `assets\highlight*.js/css` | highlight.js (BSD-3) für Codeblöcke, helle und dunkle Variante |
| `themes\*.css` | Themes (`dark`, `github`); eigene Dateien werden automatisch im Menü angeboten |
| `styles\userstyle.css` | Eigene CSS-Anpassungen |
| `ini\config.ini` | Konfiguration |
| `lang\*.ini` | Texte der Oberfläche (`de`, `en`); Vorlage für weitere Sprachen |
| `dll\` | WebView2-Bibliotheken (SDK 1.0.4314.35, .NET-Core-Build aus `lib_manual\netcoreapp3.0` plus x64-Loader) und `CheatSheetCompanionHelper.dll` (vorkompilierte C#-Helfer) |
| `tools\Build-Helper.ps1` | Baut `dll\CheatSheetCompanionHelper.dll` neu aus dem C#-Quelltext im Skript |
| `vault\` | Demo-Vault mit Hilfe-Notiz und Beispielen für jede Funktion (Vorgabe in der ini, Rückfall ohne Konfiguration) |
| `docs\` | Screenshots für das README |
| `Output\` | HTML-Cache (wird automatisch erzeugt) |

## Lint

```powershell
Invoke-ScriptAnalyzer -Path .\CheatSheetCompanion.ps1 -ExcludeRule PSAvoidTrailingWhitespace, PSReviewUnusedParameter, PSUseDeclaredVarsMoreThanAssignments, PSAvoidGlobalVars, PSAvoidUsingEmptyCatchBlock, PSAvoidAssignmentToAutomaticVariable, PSUseOutputTypeCorrectly, PSUseSingularNouns, PSUseShouldProcessForStateChangingFunctions
```

## Lizenz

Apache License 2.0, siehe [LICENSE](LICENSE). highlight.js steht unter BSD-3, die WebView2-Bibliotheken (SDK 1.0.4314.35) unter Microsofts BSD-artiger WebView2-SDK-Lizenz, siehe `dll\LICENSE-WebView2.txt`.
