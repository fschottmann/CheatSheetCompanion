---
aliases:
  - CheatSheetCompanion
  - Manual
  - Hilfe
tags:
  - help
  - csc
---
# Help

CheatSheetCompanion keeps the right note on top of whatever you are working in. It watches the **active window** (title and process name) and, if you switch it on, the **clipboard**, and shows the notes of this vault whose names, aliases or rules fit best. This demo vault is small on purpose: every feature has a note that shows it.

> [!tip] Try it right now
> 1. Open **Notepad** or **Windows Terminal**: the matching note from `Apps` or `Dev` appears as a tab.
> 2. Press the **Clip** button, then copy `sudo systemctl restart nginx` somewhere: [[Linux Shell]] appears, found through its alias *nginx*.
> 3. Press **Ctrl+K** and type `regex`: the search finds names, aliases and note contents.
> 4. Right-click into this text and look at the menu.

## The window

| Element | What it does |
| --- | --- |
| ▲▲▲ (top bar) | Click: collapse the window to a thin strip and back. Right-click: quit. The strip shows the current note. |
| Search box | Fuzzy search like Obsidian's quick switcher (Ctrl+K). Names first, then contents with a snippet. Arrow keys and Enter. |
| ◀ ▶ | Back and forward through the shown views (also Alt+Left / Alt+Right or the mouse thumb buttons). |
| ✎ | Open the shown note in Obsidian or the default editor (`SourceLink`). In the text: Ctrl+double click. |
| + | Create a new note. The name is proposed from the active window or the clipboard, the folder is selectable, `_templates\note.md` is the template. |
| **Pin** | Keep the shown note as a tab (📌) while matching goes on. Several pins are possible, they survive a restart. |
| **Window** | Follow the active window (title and process). |
| **Clip** | Follow the copied text, for example from RDP or SSH sessions, chats or web pages. |
| **AC** | Auto Copy: text you select in the note is copied when you release the mouse button. |
| Tabs | Up to `MaxCandidates` notes with their score (0.00 to 1.00). Click, Ctrl+1 to Ctrl+9 or Ctrl+Tab. The colour shows the source: window title, process, clipboard, rule or pin. |
| Status line | What was matched and why (`ShowMatchReason`). |

## How a note is found

- **Name**: the window title or copied text contains the note name, or a word of the title is similar to it (n-gram similarity, thresholds in `ini\config.ini`).
- **Aliases** in the front matter count like names. [[Linux Shell]] is found through *nginx*, *ssh* or *bash*; [[Microsoft Edge]] through *Chrome* or *Browser*. Add aliases from the right-click menu without opening the editor.
- **Rules** (`[Rules]` in the ini) bind a process or a title pattern to a note, always first with score 1.00.
- Notes in `MatchIgnoreFolders` (here: `Daily`) or with `csc: ignore` in the front matter (see [[Meeting Notes]]) are never matched automatically but stay searchable.
- No match: the shown note stays and the status line tells you (`NoMatchBehavior = keep`), or a page with the closest notes and a *Create note* button appears (`suggest`).

## Handles at the right edge

| Handle | Content |
| --- | --- |
| **TOC** | Table of contents of the active note. |
| **COPIED** | Everything you copied through the companion. One click copies it again. |
| **LINKED** | Notes that link to the active note, with the line of the mention (Obsidian's linked mentions). |

The panels close when you click elsewhere, press Escape or switch to another program.

## Right-click menu

- *Copy to clipboard*: the selection as plain text.
- *Copy as…*: selection or whole note **formatted** (HTML plus text, so Teams and Outlook keep lists, tables and code blocks) or the whole note as text. Ctrl+Shift+C does the same for the selection.
- *Edit note*, *Append clipboard to note* (the copied text goes to the end of the note as it is: a command, meeting notes, a paragraph from a web page).
- *Add alias*: window title, process name, clipboard, selection or your own text.
- *Look up clipboard* once, *Pin note*, *Theme*, *Vault* (with several vaults).

## Rendering

See [[Markdown Showcase]]: wikilinks, embeds, callouts, tags, highlights, tables, task lists, footnotes, front matter and syntax highlighting. Code blocks get a *Copy* button on hover.

## Looks and behaviour

- **Themes**: `auto` follows Windows light or dark, `themes\*.css` adds your own. The window chrome follows.
- **Opacity**: `InactiveOpacity` makes the window see-through while another program is active; it is fully visible under the mouse.
- **Language**: `Language = auto` takes the Windows display language, `de` and `en` are included, `lang\<code>.ini` adds more.
- **Memory**: collapsed, the renderer sleeps like a background tab.

## Keyboard

| Keys | Action |
| --- | --- |
| Ctrl+Shift+F11 | Collapse or expand (global hotkey) |
| Ctrl+Shift+F2 / F3 | Toggle Window / Clip mode (global) |
| Ctrl+K | Search |
| Ctrl+1 … 9, Ctrl+Tab | Switch tabs |
| Alt+Left / Alt+Right | Back / forward |
| Ctrl+Shift+C | Copy selection formatted |
| Escape | Close panels |

Everything else lives in `ini\config.ini` (comments explain each key) and in the README.
