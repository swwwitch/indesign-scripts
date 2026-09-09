# Delete Markdown markers and apply styles

[![Direct](https://img.shields.io/badge/Direct%20Link-IdRemoveMarkerApplyStyle.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdRemoveMarkerApplyStyle.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdRemoveMarkerApplyStyle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Applies a paragraph style and a character style to paragraphs carrying a leading marker — Markdown headings, bullets, numbered lists and the like — then deletes the marker together with the spaces that follow it. Markers can be detected from the search target, and the number of matches is shown before the run.

## Features

- The search text comes either from a text field or from automatic detection
- Detection walks the paragraphs of the search target and lists every marker that repeats at least twice, most frequent first, with its count. Changing the search target runs the detection again
- Line-head symbols picked up: `#` `*` `-` `>` `+` `~` `=` `|` `_` `` ` `` plus the Japanese bullets `●` `○` `◎` `◆` `◇` `■` `□` `▲` `△` `▼` `▽` `★` `☆` `・` `※` and the full-width `＊` `＃` `－` `＋` `＞`
- Numbered lists cover `1.` `1)` `（一）` `一.` `一、` `一）`, searched with GREP because the number changes. Kanji numerals include 十, 百 and 千
- Bold `**text**` is treated as an enclosing marker, and both `**` are deleted
- The styles are applied to each paragraph that carries the marker, and the marker itself is removed together with the spaces, full-width spaces and tabs that follow it
- The Matches row in the Find What panel shows how many matches the current settings find, recounted as you type or change the search target, with text on hidden or locked layers left out
- Match the heading level picks the style named `h3` / `heading 3` for the selected heading level (on by default)
- Apply every Markdown heading level processes `#` through `######` in one run, assigning `h1` to `h6` per level
- The search target can be the story, the document or all open documents, chosen with radio buttons
- Styles inside style groups are available and applied
- Every control carries a tooltip describing what it does
- The whole run is a single undo step (Cmd+Z)

## Usage

1. Place the cursor in text that contains the markers (opening the document is enough when the whole document is the target)
2. Run the script
3. Choose the search text, the styles and the options, then click OK

## Notes and limitations

- A marker chosen by automatic detection is searched at the line head (GREP `^`), so the same symbol in the middle of a paragraph is left alone. Bold `**` is the exception, since it encloses the text rather than starting the line.
- Text entered by hand is not anchored to the line head, but it never matches part of a longer run of the same character: `##` does not match `### Heading`.
- Spaces, full-width spaces and tabs after the marker are deleted with it, however many there are. A line without any separator (`##Heading`) is handled too.
- The heading style is looked up as `h3`, then `H3`, then `heading 3`, then `Heading 3`, and the first match is used.
- While applying every heading level, the Find What and Styles to Apply panels are dimmed and their settings — the character style included — are not used.
- Heading levels without a matching paragraph style are skipped silently, so a document that defines only h1 to h3 raises no error.
- The story scope needs a text selection or the cursor placed in a story; without one there is no search target and the run stops.
- With all open documents the style lists come from the active document. Documents without a paragraph style of the same name are skipped and reported by name.
- Undo is per document. After a run over all open documents, undo in each document separately.
- Find options (width sensitive, kana sensitive, include footnotes and master pages, and so on) are set by the script, so the last Find/Change settings are never inherited. Every search is a GREP search, which is always case sensitive.
- Hidden layers, locked layers and locked stories are outside the search, and the detection counts leave out the text on hidden and locked layers as well.
- Counts come in two kinds. The detection list, and Matches while a detected marker is selected, are paragraph counts: `##Heading` and `## Heading` are the same marker, and a paragraph with two bold spans still counts as one. Matches for text entered by hand, and the completion message, count the markers to delete, so each bold span adds two.
- The character style is applied to the whole paragraph that carries the marker, not to the enclosed text alone.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/style/IdRemoveMarkerApplyStyle.jsx` |
| Version | v1.1.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-09-09 |
| Last updated | 2026-09-10 |
| Article | https://note.com/dtp_tranist/n/n3a0d4c0dacdb |

## Change log

### v1.1.0 (2026-09-10)

Behaviour

- Spaces, full-width spaces and tabs after the marker are now deleted with it
- Text entered by hand is searched with GREP too, so it no longer matches part of a longer run of the same character (`##` against `### Heading`)
- `##Heading` and `## Heading` are now counted as the same marker
- Text on hidden or locked layers is excluded from the detection counts

Dialog

- A Matches row was added to the Find What panel, showing how many matches the current settings find
- The search target moved into a panel of its own, as a column of radio buttons instead of a dropdown
- The Options panel is now named Markdown, and the Styles panel Styles to Apply
- The search scope carries a tooltip per option, and the style rows carry one on their labels too

Wording

- The dialog and the undo step are now named Delete Markers and Apply Styles, since a character style is applied too
- The heading checkbox is now Match the heading level, and the scope panel Search In
- The completion message reports the markers deleted

### v1.0 (2026-09-09)

- First release

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
