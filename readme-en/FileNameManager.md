# FileNameManager

Edits the active document's filename segment by segment (base / date / text / version) and renames, saves as, or saves a copy.

## Features

- Modes: rename the file, Save As (default), or Save a Copy
- Splits the filename into base / date / text / version while preserving their original order
- Timestamp: none or YYYYMMDD; version: none, -vN, or -v0N
- Title: none, parent folder name, or a custom value
- Separator: leave as is, -, or _ (applied consistently across the filename)

## Usage

1. Open the target document
2. Run the script
3. Choose the mode and segments, then click OK

## Notes and limitations

- The destination is always the folder of the current file; unsaved documents prompt for a folder.
- Selecting a version bumps the original number by one, or adds a new one.
- See Id-FileNameManager.jsx for a more feature-rich variant.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/document/FileNameManager.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-05-27 |
| Last updated | 2026-05-27 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
