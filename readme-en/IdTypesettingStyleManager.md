# Configure composition settings for paragraph styles

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTypesettingStyleManager.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdTypesettingStyleManager.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTypesettingStyleManager.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Sets typesetting options (kinsoku, mojikumi, grid alignment, hyphenation, and more) for paragraph styles from a single dialog.

## Features

- Target the selection, all styles, or a specified set
- Reads the current typesetting, language and hyphenation settings from the selected paragraph as defaults
- Apply presets (Western typesetting, grid-first, ignore grid, source code, InDesign defaults)
- Export the current settings to the Desktop as a preset code snippet
- Hyphenation-related controls enable and disable with the hyphenation checkbox

## Usage

1. Open the target document (place the cursor in a paragraph to seed the defaults)
2. Run the script
3. Choose the settings and click OK

## Notes and limitations

- [No Paragraph Style], [Basic Paragraph] and styles inside groups whose name starts with "_" are excluded.
- Overrides in the selection are always cleared after applying.
- Quotes, language and units are written to the application preferences.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/style/IdTypesettingStyleManager.jsx` |
| Version | v1.1.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-05-06 |
| Last updated | 2026-05-07 |
| Article | https://note.com/dtp_tranist/n/n7f67e8da571f |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
