# Transpose table rows and columns

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTransposeTableRowsCols.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdTransposeTableRowsCols.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTransposeTableRowsCols.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Transposes the rows and columns of the selected table, with dialog options for header rows and merged cells.

## Features

- Swaps the contents plus point size, font, text colour, cell fill colour and tint
- When merged cells exist, choose between cancelling and unmerging first
- A checkbox controls whether header rows are transposed
- Header and footer row counts are restored as closely as possible

## Usage

1. Select a table, a cell, or text inside a table
2. Run the script
3. Choose how header rows and merged cells are handled, then click OK

## Notes and limitations

- When the table has neither header rows nor merged cells the dialog is skipped.
- Rows or columns are temporarily added to square the table, then removed afterwards.
- Based on Table Transpose v1.0 by Iain Anderson.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/table/IdTransposeTableRowsCols.jsx` |
| Version | v1.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2025-11-25 |
| Last updated | 2025-11-25 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
