# Insert pages keeping the current parent page

[![Direct](https://img.shields.io/badge/Direct%20Link-IdAddPagesUsingCurrentMaster.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/page/IdAddPagesUsingCurrentMaster.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAddPagesUsingCurrentMaster.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Inserts a given number of pages right after the current page, carrying over the parent (master) page applied to it.

### Features

- Shows the current page name and the applied parent page in the dialog
- Prompts for the number of pages to insert (default 2)
- Applies the current parent page to every inserted page
- Adds pages at document level so they reflow into the correct spreads

### Usage

1. Make the page you want to insert after the active page
2. Run the script
3. Enter the number of pages and click OK

### Notes and limitations

- An active InDesign document is required.
- Unlike the built-in Insert Pages dialog, this keeps the parent page of the selected page.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/page/IdAddPagesUsingCurrentMaster.jsx` |
| Version | v1.3.2 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2025-06-26 |
| Last updated | 2026-10-01 |
| Article | https://note.com/dtp_tranist/n/n2d1a76097a39 |

### Update History

#### v1.3.2 (2026-10-01)

- A button row with only right-side buttons is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones

#### v1.3.1 (2026-09-30)

- Unified the button row, spacing and localization code with the shared parts (no change in behavior)

- v1.3.0 (2026-09-27): Added stepper buttons to the page count field. The arrow keys step it as well (to the next whole number; Shift to the next multiple of ten; never below 1)

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
