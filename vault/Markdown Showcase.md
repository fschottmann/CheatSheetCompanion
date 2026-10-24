---
aliases:
  - Rendering
  - Obsidian Markdown
tags:
  - demo
  - markdown
created: 2026-10-24
status: example
---
# Markdown Showcase

This note shows what the companion renders. Set `ShowFrontmatter = true` to see the properties above as a box. %% This comment is hidden, like in Obsidian. %%

## Text

Normal, **bold**, *italic*, ~~strikethrough~~, ==highlighted== and `inline code`. Keys look like <kbd>Ctrl</kbd>+<kbd>K</kbd>.
A single line break is kept (soft break), as in Obsidian's default. Tags: #demo #markdown/nested

## Links

- Wikilink: [[Help]]
- Alias: [[Help|The manual]]
- Heading: [[Dev/Regex#Quantifiers]]
- Markdown link: [PowerShell cheat sheet](Dev/PowerShell.md)
- External: [Obsidian](https://obsidian.md)
- Missing note: [[Does not exist]] (shown muted)

Clicking a note link opens the note in the companion; ◀ brings you back.

## Embeds

Image with a width: ![[logo.png|240]]

A section of another note:

![[Dev/Regex#Quantifiers]]

## Callouts

> [!note] Note
> Plain callout with a title.

> [!tip]
> Tip without a title.

> [!warning]- Collapsed warning
> Click the title to expand.

> [!example] Steps
> 1. First
> 2. Second

## Lists and tasks

- Item
    - Nested item
- [ ] Open task
- [x] Done task

1. Ordered
2. List

## Table

| Shortcut | Action | Scope |
| --- | --- | --- |
| Ctrl+K | Search | Companion |
| Ctrl+Shift+C | Copy formatted | Note text |
| `Ctrl+Tab` | Next tab | Companion |

## Code

Hover a block for the *Copy* button. Languages are highlighted.

```powershell
Get-Process | Where-Object { $_.CPU -gt 10 } | Sort-Object CPU -Descending | Select-Object -First 5
$name = "Value $env:USERNAME"
```

```bash
sudo systemctl restart nginx && sudo journalctl -u nginx -n 50
```

```json
{ "theme": "auto", "pins": ["Help"] }
```

## Quote and footnote

> Knowledge is only useful when it is where you need it.[^1]

[^1]: Which is the whole point of this tool.
