# Switch font variants and weights in bulk

[![Direct](https://img.shields.io/badge/Direct%20Link-IdFontConverter.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/font/IdFontConverter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdFontConverter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Switches font variants (character set, P, UD, N, NT, weight) in bulk across a selection, story, spread or the whole document.

### Features

- Target the selection, the story, the whole document, or the active spread
- Toggle character set (Std / Pro / Pr5 / Pr6), N and NT, and UD and P between keep / off / on
- Max and MaxN presets move fonts to the richest available character set
- Paragraph styles, character styles, composite fonts, and locked or hidden objects can all be included
- A preview lists every change (old to new) using Japanese font names and warns about fonts that are not installed

### Usage

1. Open the target document (select text or frames if you target the selection)
2. Run the script
3. Choose the target and conversion options, then click Run
4. Review the preview and confirm only the changes you want

### Notes and limitations

- AXIS (Type Project) fonts use a different character-set scheme and get dedicated handling: width and Joyo are preserved, and only N and Std/Pro are switched.
- When no matching weight exists, the nearest weight is substituted.
- Changes are applied per textStyleRange and the whole run is a single undo step.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/font/IdFontConverter.jsx` |
| Version | v1.1.3 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-06-17 |
| Last updated | 2026-10-01 |
| Article | https://note.com/dtp_tranist/n/n261c771b4b41 |

### Update History
- v1.1.4 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)

#### v1.1.3 (2026-10-01)

- A button row with only right-side buttons is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones

#### v1.1.2 (2026-09-30)

- Unified the button row, panel spacing and localization code with the shared parts. Radio-button and checkbox panels are tighter

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
