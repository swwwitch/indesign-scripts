# Apply and manage GREP styles on paragraph styles

[![Direct](https://img.shields.io/badge/Direct%20Link-IdGrepStyleApplier.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdGrepStyleApplier.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdGrepStyleApplier.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Applies and manages GREP styles on paragraph styles: pick a rule and a character style, then apply it to several paragraph styles at once.

## Features

- Choose from saved rules (bullet label, language, no-break ending, TOC number, inline graphic)
- Add custom rules through a dedicated dialog
- Select or create the character style, with rule-specific defaults applied automatically
- Select multiple target paragraph styles (Option/Alt-click toggles select all)
- An existing GREP style with the same expression in the same paragraph style is overwritten

## Usage

1. Open the target document
2. Run the script
3. Choose the rule, character style and target paragraph styles, then click OK

## Notes and limitations

- The OK button stays disabled until both a paragraph style and a character style are selected.
- IdNestedStyleSetup.jsx is identical; using just one of the two is recommended.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/style/IdGrepStyleApplier.jsx` |
| Version | v1.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-05-03 |
| Last updated | 2026-05-03 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
