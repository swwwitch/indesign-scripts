# Transpose table rows and columns

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTransposeTableRowsCols.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdTransposeTableRowsCols.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTransposeTableRowsCols.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Transposes the whole table the selection sits in, however much of it is selected. The dialog chooses whether to keep the header row setting and how merged cells are handled.

![Dialog with a checkbox for keeping the header row setting, and radio buttons for how merged cells are handled](../png/ss-436-440-144-20260907-183901.png)

## Features

- Works from a cursor position, a few selected characters, several cells or the whole table alike
- Swaps the contents plus point size, font, text colour, cell fill colour and tint
- When merged cells exist, choose between cancelling and unmerging first
- A checkbox controls whether the header row setting is kept after transposing (the whole table is always transposed)
- Footer row counts are restored as closely as possible

## Usage

1. Place the cursor inside the table, or select cells or text
2. Run the script
3. Choose whether to keep the header rows and how merged cells are handled, then click OK

## Notes and limitations

- Any selection inside a table transposes that whole table; a partial selection cannot be transposed on its own.
- With nested tables, the inner table holding the selection is the one transposed.
- When the table has neither header rows nor merged cells the dialog is skipped.
- Rows or columns are temporarily added to square the table, then removed afterwards.
- Empty cells receive a single space, which the transpose needs in order to run.

## Original

Table Transpose (modified for robustness)

Original: Table Transpose v1.0 by Iain Anderson

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/table/IdTransposeTableRowsCols.jsx` |
| Version | v1.0.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2025-11-25 |
| Last updated | 2026-09-07 |
| Article | https://note.com/dtp_tranist/n/nc6dbdb3af6a1 |

## Changelog

### v1.0.1 (2026-09-07)

- Fixed a partial cell selection not reaching the whole table; the script now walks up the parent chain to find it
- Fixed a table with footer rows possibly transposing incorrectly, because a padding row could land before the footer; the header and footer designation is now dropped before transposing
- Added a link to the article
- Reworded the dialog to match what it actually does ("Include header rows" became "Keep header row setting", "Do nothing (cancel)" became "Cancel the operation", and so on)
- Added tooltips to the checkbox and the radio buttons
- The merged-cell panel is now disabled as a whole when the table has no merged cells
- Tidied the header, naming and JSDoc to match the house rules (no behavior change)
- Factored out the cell-format swap and split the transpose into padding, triangle swapping and trimming steps

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
