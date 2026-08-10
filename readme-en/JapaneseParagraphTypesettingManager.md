# JapaneseParagraphTypesettingManager

Reviews and bulk-applies Japanese typesetting settings (kinsoku set, kinsoku adjustment, mojikumi, composer) for paragraph styles in a matrix UI.

## Features

- One row per paragraph style showing its current typesetting settings
- The top "All" row acts as a copy source; per-column Apply buttons push its value to every style
- Existing paragraph style settings are read and used as the dialog defaults
- Handles both custom mojikumi tables and built-in preset names
- Walks paragraph style groups recursively and skips excluded groups

## Usage

1. Open the target document
2. Run the script
3. Adjust the matrix, then click OK

## Notes and limitations

- [No Paragraph Style], [Basic Paragraph] and styles inside groups whose name starts with "_" are excluded.
- Defaults can be changed via the DEFAULT_* variables at the top of the script.
- The whole run is a single undo step.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/style/JapaneseParagraphTypesettingManager.jsx` |
| Version | v1.2.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-05-05 |
| Last updated | 2026-05-06 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
