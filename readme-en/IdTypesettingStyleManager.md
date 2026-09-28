# Configure composition settings for paragraph styles

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTypesettingStyleManager.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdTypesettingStyleManager.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTypesettingStyleManager.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Sets typesetting options (kinsoku, mojikumi, grid alignment, hyphenation, and more) for paragraph styles from a single dialog.

### Features

- Target the selection, all styles, or a specified set
- Reads the current typesetting, language and hyphenation settings from the selected paragraph as defaults
- Apply presets (Western typesetting, grid-first, ignore grid, source code, InDesign defaults)
- Export the current settings to the Desktop as a preset code snippet
- Hyphenation-related controls enable and disable with the hyphenation checkbox
- The number fields step with the ∧∨ buttons or the arrow keys to the next whole number (1.5 → 2); Shift snaps to the next multiple of 10, Option steps by 0.1

### Usage

1. Open the target document (place the cursor in a paragraph to seed the defaults)
2. Run the script
3. Choose the settings and click OK

### Notes and limitations

- [No Paragraph Style], [Basic Paragraph] and styles inside groups whose name starts with "_" are excluded.
- Overrides in the selection are always cleared after applying.
- Quotes, language and units are written to the application preferences.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/style/IdTypesettingStyleManager.jsx` |
| Version | v1.2.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-05-06 |
| Last updated | 2026-09-27 |
| Article | https://note.com/dtp_tranist/n/n7f67e8da571f |

### Update History

#### v1.2.0 (2026-09-27)

- Added stepper buttons to the number fields (Auto leading and hyphenation). The arrow keys step them the same way (to the next whole number; Shift to the next multiple of ten)
- Fixed an error on load that kept the script from starting

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
