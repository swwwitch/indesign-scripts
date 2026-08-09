# SmartBorderBuilder

Draws and clears table-cell borders with a live preview, letting you set the mode, weight, colour and tint.

## Features

- Modes: all, outer only, inner only, horizontal only, vertical only, bottom only, right only, header row, header column, clear left/right, clear all
- Keyboard shortcuts for mode switching (A/E/I/H/V/B/U/L/R/C)
- The weight follows the document's stroke-weight unit, with presets and arrow-key stepping
- Border colour from a document swatch, plus a 0-100 tint
- A Standard Mode / Preview toggle button switches the screen mode
- Remembers the settings confirmed with OK and restores them on the next run

## Usage

1. Select the table cells you want to change
2. Run the script
3. Set the mode, weight, colour and tint, then click OK

## Notes and limitations

- Partial selections are rebuilt from the selection rectangle; merged cells are resolved by finding the cell covering each coordinate.
- Turn off Clear Existing Borders to overwrite while keeping the current borders.
- The selection state from before the run is restored when the dialog closes.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/table/SmartBorderBuilder.jsx` |
| Version | v1.6.7 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-04-11 |
| Last updated | 2026-04-17 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
