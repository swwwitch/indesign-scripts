# Unmerge cells in a table

[![Direct](https://img.shields.io/badge/Direct%20Link-IdCellUnmerge.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdCellUnmerge.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCellUnmerge.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Unmerges merged table cells, with dialog options for distributing the original text and for the target scope.

## Features

- Switch between Default (no distribution) and Distribute text
- Target the whole table or only the selected cells
- Distribution is limited to the cells produced by that merged cell
- Cells selected more than once are processed only once

## Usage

1. Select the table cells
2. Run the script
3. Choose the distribution and scope, then click OK

## Notes and limitations

- Distribute text puts the same text into every cell produced by that merged cell. Neighbouring cells are left untouched.
- The distributed text is the text the merged cell was showing. Formatting is not carried over; the text is inserted as plain text.
- A failure on one cell does not abort the whole run.
- The whole run collapses into a single undo step.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/table/IdCellUnmerge.jsx` |
| Version | v1.0.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-04-17 |
| Last updated | 2026-08-27 |

## Changelog

### v1.0.1 (2026-08-27)

- Fixed Selected cells only, where cells were left merged or picked up text from another cell. The cause was treating `contents` of a merged cell as a string when it returns an array of the constituent cells' texts
- Limited Distribute text to the cells produced by that merged cell
- Fixed Whole table missing merged cells. Each unmerge changed the cell count and shifted the references that followed
- Wrapped the whole script in an IIFE so nothing leaks into the global scope

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
