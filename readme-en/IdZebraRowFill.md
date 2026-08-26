# IdZebraRowFill

[![Direct](https://img.shields.io/badge/Direct%20Link-IdZebraRowFill.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdZebraRowFill.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdZebraRowFill.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Applies alternating (zebra) fills to the selected table cells, based on the row order within the selection, with a live preview.

## Features

- Separate colour and tint (0-100%) for odd and even rows
- A Swap button exchanges the odd and even settings
- Skip a number of rows from the top and columns from the left of the selection
- Fields step by 1 with the arrow keys and by 10 with Shift; the tint slider steps by 1%, 10% with Shift and 5% with Option
- Selecting None or Paper dims the tint controls automatically

## Usage

1. Select the table cells you want to fill
2. Run the script
3. Set the odd and even colours, tints and skip counts, then click OK

## Notes and limitations

- Registration is always hidden from the colour list; Paper and None are shown with localized names.
- Skipped cells are restored to the colour they had when the dialog opened.
- Cancelling rolls back everything the preview applied.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/table/IdZebraRowFill.jsx` |
| Version | v1.2.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-04-17 |
| Last updated | 2026-04-17 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
