# Apply alternating row shading to a table

[![Direct](https://img.shields.io/badge/Direct%20Link-IdZebraRowFill.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdZebraRowFill.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdZebraRowFill.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Applies alternating (zebra) fills to the selected table cells, based on the row order within the selection, with a live preview. A given number of rows from the top, and columns from the left, can be left out of the fill.

## Features

- Separate colour and tint (0-100%) for the Odd-numbered Rows and Even-numbered Rows panels
- The most frequent colour and tint combinations in the selection become the starting values
- Swap odd and even, in the Options panel, exchanges the settings of the two panels
- Rows and Columns, in the Skip panel, leave a number of rows from the top, and columns from the left, out of the fill
- Every change is previewed on the document right away
- The button at the bottom left switches the document window between the Normal and Preview screen modes
- Fields step by 1 with the arrow keys and by 10 with Shift; the tint slider steps by 1%, 10% with Shift and 5% with Option
- Selecting None or Paper dims the tint controls automatically
- Every control carries a tooltip

## Usage

1. Select the table cells you want to fill
2. Run the script
3. Set the colours, tints and skip counts, then click OK

## Options

| Item | Description |
| --- | --- |
| Odd-numbered Rows / Even-numbered Rows | Counted from the top of the selection, not by the row numbers of the table. Skipped rows are not counted |
| Swap odd and even | Exchanges the colour and tint of the two panels |
| Skip: Rows | Leaves the given number of rows at the top of the selection out of the fill |
| Skip: Columns | Leaves the given number of columns at the left of the selection out of the fill |
| Normal / Preview | Switches the screen mode of the document window. The button shows the mode you are switching to |

## Notes and limitations

- Registration is always hidden from the colour list; Paper, None and Black are shown with localized names.
- Skipped cells are restored to the colour and tint they had before the script ran.
- The cell selection is cleared while the dialog is open, so that it does not hide the preview.
- Cancelling rolls back everything the preview applied.
- The steps are recorded for undo as Fill Preview and Apply Alternating Fills.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/table/IdZebraRowFill.jsx` |
| Version | v1.2.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-04-17 |
| Last updated | 2026-09-16 |
| Article | https://note.com/dtp_tranist/n/n20ff60f6b508 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
