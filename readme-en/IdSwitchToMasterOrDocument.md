# Switch between the parent page and the document page

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSwitchToMasterOrDocument.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/page/IdSwitchToMasterOrDocument.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSwitchToMasterOrDocument.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Detects whether the active page is a parent (master) page or a document page and switches to the other one.

### Features

- From a document page, jumps to the applied parent page
- From a parent page, returns to the document page you came from
- On spreads, picks the parent page side that matches the page

### Usage

1. Make the page you want to switch from active
2. Run the script

### Notes and limitations

- The page to return to is stored temporarily in the document's `label` property.
- Pages with no applied parent page cannot be switched.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/page/IdSwitchToMasterOrDocument.jsx` |
| Version | v1.0.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2025-07-02 |
| Last updated | 2026-09-30 |

### Update History
- v1.0.2 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)

#### v1.0.1 (2026-09-30)

- Unified the localization code with the shared part (no change in behavior)

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
