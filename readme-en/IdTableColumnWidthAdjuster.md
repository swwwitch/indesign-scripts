# Adjust table column widths in bulk

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTableColumnWidthAdjuster.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdTableColumnWidthAdjuster.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableColumnWidthAdjuster.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Adjusts table column widths, either per column or through a single batch entry.

### Features

- Choose between per-column input and batch input
- Size either by width or by character count
- Per column: width, character-count equivalent, left/right inset and auto-fit
- Auto-fit widens a column until its second line disappears, or estimates from the character count when there is no second line
- "Apply to All Columns" mirrors the value you are editing to every column

### Usage

1. Place the cursor in a cell or select the table
2. Run the script
3. Choose the input method and values, then click OK

### Notes and limitations

- Column widths and insets follow the document unit settings.
- Character-count conversion uses the most dominant font size in the table.
- Changes apply live and are reverted on cancel.
- Batch input accepts spaces or commas (for example 30 50 70 70 or 30, 50, 70, 70).

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/table/IdTableColumnWidthAdjuster.jsx` |
| Version | v1.3.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-04-19 |
| Last updated | 2026-10-01 |

### Update History

#### v1.3.1 (2026-10-01)

- Fixed Auto Fit on wrapping columns adding the left/right inset twice, which made the column too wide
- Fixed character-count conversion and the unit label when the ruler unit is Q, Ha, agates, ciceros or pixels
- Fixed headers and column numbers showing in black on a dark UI
- Added tooltips to Set by Character Count, Apply to All Columns, L/R Inset, Auto Fit and the batch input field
- Shortened the English headers and the batch input hint
- Cleaned up the code (split long functions, merged duplicates, removed unused code)

#### v1.3.0 (2026-09-30)

- Fixed an error when applying values to all columns
- The screen-mode toggle now shows the mode it switches to, as in the other scripts
- Unified the button row, table selection and localization code with the shared parts

#### v1.2.0 (2026-09-27)

- Added stepper buttons to each column's Width, Character Count and Left/Right Inset fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten; Option by 0.1)

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
