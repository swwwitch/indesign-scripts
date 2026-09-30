# Find and change Markdown syntax, then apply styles

[![Direct](https://img.shields.io/badge/Direct%20Link-IdFindChangeByListMarkdown.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdFindChangeByListMarkdown.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdFindChangeByListMarkdown.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Runs a batch of Markdown find/change operations and applies the matching paragraph and character styles. Based on the FindChangeByList.jsx bundled with InDesign.

### Features

- Handles headings, bullets, numbered lists, blockquotes, code blocks, tables and images
- Switches between HTML and MS Word style-name presets
- Scope can be the document, the story or the selection
- Cleans up leading/trailing spaces, blank lines, HTML comments and horizontal rules

### Usage

1. Open the document containing the Markdown text
2. Run the script
3. Choose the scope, format and per-item styles, then click Run

### Notes and limitations

- The operation order is fixed to respect dependencies (code blocks, then numbered lists, then headings, and so on).
- Every replacement is a single undo step.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/text/IdFindChangeByListMarkdown.jsx` |
| Version | v1.0.3 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-03-17 |
| Last updated | 2026-10-01 |
| Article | https://note.com/dtp_tranist/n/n8c0211d92c96 |

### Update History

#### v1.0.3 (2026-10-01)

- A button row with only right-side buttons is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones

#### v1.0.2 (2026-09-30)

- Unified the button row, spacing and localization code with the shared parts. A button row with only right-side buttons is now centered

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
