# SetLastColumnToFrameWidth

Resizes the last column of the table at the cursor so the whole table matches the width of its parent text frame.

## Features

- Only the last column is stretched or shrunk; the other column widths stay as they are
- The width of the text frame holding the table is used as the reference

## Usage

1. Place the cursor inside the table
2. Run the script

## Notes and limitations

- The table must live inside a text frame.
- If the other columns already exceed the frame width, the resulting last-column width can be negative.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/table/SetLastColumnToFrameWidth.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-04-17 |
| Last updated | 2026-04-17 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
