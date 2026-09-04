# Split each paragraph into its own text frame

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSplitParagraph.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdSplitParagraph.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSplitParagraph.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Splits each paragraph in the selected text frame into its own text frame, keeping the original position and width (height for vertical text). Vertical text frames are supported as well.

## Features

- Places each frame on the original paragraph baseline (vertically for horizontal text, horizontally for vertical text)
- Keeps the original width (height for vertical text)
- Inherits the text frame preferences
- Supports vertical text frames (overset is resolved by growing the frame leftward)
- When overset text exists, offers Expand to resolve or Run anyway

## Usage

1. Select exactly one text frame
2. Run the script
3. If there is overset text, choose how to proceed and click OK

## Notes and limitations

- Threaded text frames, anchored frames, and frames nested inside another object are not supported; the script reports this and stops.
- Rotated frames are out of scope (the split uses the unrotated bounding box).
- Empty paragraphs are excluded from the output. When no paragraph can be split, the source frame is left untouched.
- Choosing Run anyway means hidden overset text is not output.
- If the overset text cannot be resolved, the frame is restored to its original size and the run is cancelled.
- The whole operation reverts in a single undo.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/text/IdSplitParagraph.jsx` |
| Version | v1.1.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-03-16 |
| Last updated | 2026-06-30 |
| Article | https://note.com/dtp_tranist/n/n8793ea71526b |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
