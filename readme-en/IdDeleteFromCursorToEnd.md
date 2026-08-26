# IdDeleteFromCursorToEnd

[![Direct](https://img.shields.io/badge/Direct%20Link-IdDeleteFromCursorToEnd.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdDeleteFromCursorToEnd.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdDeleteFromCursorToEnd.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Deletes from the cursor to the end of the current paragraph, or deletes just a trailing mark when the cursor sits right before one.

## Features

- Never crosses a cell boundary inside a table
- When the character after the cursor is one of 。！？,.、，． and is the paragraph's last character, only that character is deleted
- A trailing 。 is kept by default (toggle with `KEEP_TRAILING_MARU`)
- A 。 immediately before the cursor is included in the deletion
- Uses cut by default so the deleted text goes to the clipboard (toggle with `COPY_TO_CLIPBOARD`)

## Usage

1. Place the cursor inside a text frame
2. Run the script

## Notes and limitations

- A trailing line break is not treated as the last character; the character before it is used instead.
- The whole run is a single undo step.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/text/IdDeleteFromCursorToEnd.jsx` |
| Version | v1.2.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-06-27 |
| Last updated | 2026-07-05 |
| Article | https://note.com/dtp_tranist/n/nf0b1e27e1f81 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
