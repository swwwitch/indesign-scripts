# TableColumnEqualizer

Matches the width of the table at the cursor to the width of its parent text frame.

## Features

- Walks up the parent chain to find the text frame and applies its width to the table
- Column widths are redistributed by InDesign's default behaviour

## Usage

1. Place the cursor inside the table
2. Run the script

## Notes and limitations

- If the table is not inside a text frame the script reports it and stops.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/table/TableColumnEqualizer.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-04-17 |
| Last updated | 2026-04-17 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
