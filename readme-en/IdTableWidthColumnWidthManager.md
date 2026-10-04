# Adjust the table width and the column widths together

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTableWidthColumnWidthManager.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdTableWidthColumnWidthManager.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableWidthColumnWidthManager.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Resolves the table from the current selection and adjusts the table width and column widths together, with a live preview.

### Features

- Table width: do not change, auto, fit to parent frame, or custom
- Column width: equal widths, fit to content, fit to content+, adjust last column, or custom
- Always-on preview, and the original selection is restored afterwards
- Custom values use the current ruler unit and support stepper buttons and arrow-key stepping (Shift: to the next multiple of 10, Option: ±0.1)

### Usage

1. Place the cursor inside a table
2. Run the script
3. Choose the table and column width options, then click OK

### Notes and limitations

- When Column Width is Custom, it takes precedence over Table Width, and the table width is updated to column width times column count.
- Fit to Content+ uses the fitted widths as a base and spreads the table-width difference evenly across the columns.
- Adjust Last Column keeps every other column at its current width.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/table/IdTableWidthColumnWidthManager.jsx` |
| Version | v1.3.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-04-18 |
| Last updated | 2026-09-30 |

### Update History
- v1.3.1 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)

#### v1.3.0 (2026-09-30)

- The table can now be found from any text selection, such as a word or paragraph
- The screen-mode toggle now shows the mode it switches to, as in the other scripts
- Unified the button row, table selection and localization code with the shared parts

#### v1.2.0 (2026-09-27)

- Added stepper buttons to the Custom fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten; Option by 0.1)

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
