# Transpose table rows and columns

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSwapTableRowColumn.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdSwapTableRowColumn.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSwapTableRowColumn.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Transposes the rows and columns of the selected table. Identical to IdTransposeTableRowsCols.jsx apart from the filename.

### Features

- Swaps the contents plus point size, font, text colour, cell fill colour and tint
- Dialog options for how merged cells are handled
- A checkbox controls whether header rows are transposed

### Usage

1. Select a table, a cell, or text inside a table
2. Run the script
3. Choose how header rows and merged cells are handled, then click OK

### Notes and limitations

- Identical to IdTransposeTableRowsCols.jsx; using just one of the two is recommended.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/table/IdSwapTableRowColumn.jsx` |
| Version | v1.1.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2025-11-25 |
| Last updated | 2026-10-01 |

### Update History
- v1.1.2 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)

#### v1.1.1 (2026-10-01)

- A button row with only right-side buttons is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones

#### v1.1.0 (2026-09-30)

- Fixed the table not being found when the cursor or selected text was inside a cell
- Fixed the result shifting for tables with footer rows
- Unified the transpose logic with IdTransposeTableRowsCols and the button row with the shared part. Buttons are now centered

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
