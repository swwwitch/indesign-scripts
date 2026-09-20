# Link two text frames

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTextFrameLinker.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdTextFrameLinker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTextFrameLinker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Links the two selected text frames so that they share a single story.

## Features

- Links the two selected text frames through `nextTextFrame`
- Leaves frame size and position untouched
- Runs without a dialog
- Reports when the selection is not exactly two text frames
- Reports and stops when the first frame has overset text

## Usage

1. Select two text frames, source first and destination second
2. Run the script

## Notes and limitations

- The direction follows the selection order: the frame selected first becomes the source.
- If the first frame has overset text, linking would push out the content of the second frame, so the script stops.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/text/IdTextFrameLinker.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-09-20 |
| Last updated | 2026-09-20 |
| Article | https://note.com/dtp_tranist/n/n04ceaf4955a0 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
