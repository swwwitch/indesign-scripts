# Edit text in a dialog, then replace or insert

[![Direct](https://img.shields.io/badge/Direct%20Link-IdEditTextsByDialog.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdEditTextsByDialog.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdEditTextsByDialog.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Edits text in a multi-line dialog, then replaces the selection, inserts at the caret, or creates a new text frame.

## Features

- Replaces selected text, inserts at an insertion point, or appends to a selected text frame
- Creates a new text frame at the page centre when nothing is selected
- Treats `@#` as a visible marker for a forced line break (\n)
- A Remove Breaks button strips line breaks and markers at once

## Usage

1. Select the text you want to edit (running with no selection also works)
2. Run the script
3. Edit the field and click OK

## Notes and limitations

- Enter in the field inserts a paragraph return (\r).
- Use the Add @# button to add a forced line break (it is appended to the end of the field).

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/text/IdEditTextsByDialog.jsx` |
| Version | v0.1.4 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2025-05-28 |
| Last updated | 2026-09-25 |

## Changelog

### v0.1.4 (2026-09-25)

- Renamed the buttons to Remove Breaks / Add @#
- Added tooltips to the input field and the left-hand buttons
- Shows a message and exits when no document is open
- Fixed an error when a single special character was selected

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
