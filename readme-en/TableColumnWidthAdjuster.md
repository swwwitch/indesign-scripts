# TableColumnWidthAdjuster

Adjusts table column widths, either per column or through a single batch entry.

## Features

- Choose between per-column input and batch input
- Size either by width or by character count
- Per column: width, character-count equivalent, left/right inset and auto-fit
- Auto-fit widens a column until its second line disappears, or estimates from the character count when there is no second line
- "Apply to All Columns" mirrors the value you are editing to every column

## Usage

1. Place the cursor in a cell or select the table
2. Run the script
3. Choose the input method and values, then click OK

## Notes and limitations

- Column widths and insets follow the document unit settings.
- Character-count conversion uses the most dominant font size in the table.
- Changes apply live and are reverted on cancel.
- Batch input accepts spaces or commas (for example 30 50 70 70 or 30, 50, 70, 70).

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/table/TableColumnWidthAdjuster.jsx` |
| Version | v1.1.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-04-19 |
| Last updated | 2026-04-19 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
