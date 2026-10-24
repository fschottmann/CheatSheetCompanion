---
aliases:
  - regex101
  - Regular Expressions
tags:
  - dev
---
# Regex

Found through the alias *regex101* when that site is open in the browser.

## Character classes

| Pattern | Matches |
| --- | --- |
| `.` | Any character except newline |
| `\d` `\w` `\s` | Digit, word character, whitespace |
| `[a-z]` `[^0-9]` | Range, negated set |
| `\b` | Word boundary |

## Quantifiers

| Pattern | Meaning |
| --- | --- |
| `*` `+` `?` | 0 or more, 1 or more, 0 or 1 |
| `{3}` `{2,5}` | Exactly 3, between 2 and 5 |
| `*?` `+?` | Lazy versions |

## Groups and lookaround

```
(abc)        capturing group, $1 in replacements
(?<name>abc) named group
(?:abc)      non capturing
(?=abc)      lookahead, (?!abc) negative
(?<=abc)     lookbehind, (?<!abc) negative
```

## Examples

```powershell
'2026-10-24' -match '^(?<y>\d{4})-(?<m>\d{2})-(?<d>\d{2})$'; $Matches.y
'a1b22c333' -replace '\d+', '#'
[regex]::Matches('x=1, y=22', '\d+').Value
```

Flags: `(?i)` ignore case, `(?m)` multiline anchors, `(?s)` dot matches newline. Used in [[Notepad]] and [[PowerShell]].
