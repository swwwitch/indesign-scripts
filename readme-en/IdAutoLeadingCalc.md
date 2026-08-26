# IdAutoLeadingCalc

[![Direct](https://img.shields.io/badge/Direct%20Link-IdAutoLeadingCalc.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdAutoLeadingCalc.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAutoLeadingCalc.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Back-calculates the leading percentage from each paragraph's absolute leading and point size, then switches the paragraph to Auto leading.

## Features

- Per paragraph, derives the percentage as leading divided by point size and stores it as the auto-leading amount
- Switches leading to Auto and sets the leading model to the top/right of the virtual body
- Works with text selections, selected text frames and groups (including nested frames)
- Re-selects afterwards so the Character panel refreshes

## Usage

1. Select the target text or text frames
2. Run the script

## Notes and limitations

- Paragraphs already set to Auto leading are skipped, since there is nothing to back-calculate.
- No dialog is shown; the change is applied in place.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/text/IdAutoLeadingCalc.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-07-09 |
| Last updated | 2026-07-09 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
