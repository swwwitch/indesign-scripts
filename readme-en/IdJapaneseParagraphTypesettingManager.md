# Apply Japanese composition settings from a matrix

[![Direct](https://img.shields.io/badge/Direct%20Link-IdJapaneseParagraphTypesettingManager.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdJapaneseParagraphTypesettingManager.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdJapaneseParagraphTypesettingManager.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Reviews and bulk-applies Japanese typesetting settings (kinsoku set, kinsoku adjustment, mojikumi, composer) for paragraph styles in a matrix UI.

### Features

- One row per paragraph style showing its current typesetting settings
- The top "All" row acts as a copy source; per-column Apply buttons push its value to every style
- Existing paragraph style settings are read and used as the dialog defaults
- Handles both custom mojikumi tables and built-in preset names
- Walks paragraph style groups recursively and skips excluded groups

### Usage

1. Open the target document
2. Run the script
3. Adjust the matrix, then click OK

### Notes and limitations

- [No Paragraph Style], [Basic Paragraph] and styles inside groups whose name starts with "_" are excluded.
- Defaults can be changed via the DEFAULT_* variables at the top of the script.
- The whole run is a single undo step.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/style/IdJapaneseParagraphTypesettingManager.jsx` |
| Version | v1.3.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-05-05 |
| Last updated | 2026-10-01 |

### Update History

#### v1.3.1 (2026-10-01)

- A button row with only right-side buttons is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones

#### v1.3.0 (2026-09-30)

- Fixed paragraph styles with no mojikumi set silently getting the default mojikumi when you clicked OK (now shown and kept as none)
- Fixed built-in mojikumi presets not being recognized
- Shared the kinsoku, mojikumi and composer lists with IdTypesettingStyleManager; composer names now follow the standard naming
- Unified the button row, spacing and localization code with the shared parts

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
