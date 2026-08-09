# ResetHorizontalVerticalScale

Scans every story in the active document, including table cells and nested tables, and resets horizontal and vertical text scaling to 100%.

## Features

- Resets any style range whose horizontal or vertical scale is not 100%
- Recurses into table cells and nested tables
- Reports the number of changed text ranges when finished

## Usage

1. Open the target document
2. Run the script

## Notes and limitations

- The whole run is a single undo step.
- Every story in the active document is processed; the scope cannot be narrowed.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/text/ResetHorizontalVerticalScale.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-07-19 |
| Last updated | 2026-07-19 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
