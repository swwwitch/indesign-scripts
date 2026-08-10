# TableWidthColumnWidthManager

Resolves the table from the current selection and adjusts the table width and column widths together, with a live preview.

## Features

- Table width: do not change, auto, fit to parent frame, or custom
- Column width: equal widths, fit to content, fit to content+, adjust last column, or custom
- Always-on preview, and the original selection is restored afterwards
- Custom values use the current ruler unit and support arrow-key stepping (Shift: ±10, Option: ±0.1)

## Usage

1. Place the cursor inside a table
2. Run the script
3. Choose the table and column width options, then click OK

## Notes and limitations

- When Column Width is Custom, it takes precedence over Table Width, and the table width is updated to column width times column count.
- Fit to Content+ uses the fitted widths as a base and spreads the table-width difference evenly across the columns.
- Adjust Last Column keeps every other column at its current width.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/table/TableWidthColumnWidthManager.jsx` |
| Version | v1.1.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-04-18 |
| Last updated | 2026-05-05 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
