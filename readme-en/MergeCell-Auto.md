# MergeCell-Auto

Automatically merges adjacent cells with identical contents, horizontally and vertically.

## Features

- Merge in a single direction or in both directions with a priority order
- Choose horizontal or vertical (in both-directions mode this is the one processed first)
- Target the whole table or only the selected cells
- Contents are compared after trimming surrounding whitespace

## Usage

1. Select a table or cells inside a table
2. Run the script
3. Choose the merge mode, direction and scope, then click OK

## Notes and limitations

- "Selected cells only" uses the rectangle that encloses the selected cells.
- The whole run is a single undo step.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/table/MergeCell-Auto.jsx` |
| Version | v0.2 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-04-17 |
| Last updated | 2026-04-17 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
