# Split each paragraph into its own text frame

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSplitParagraph.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdSplitParagraph.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSplitParagraph.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Splits each paragraph in the selected text frame into its own text frame, keeping the original position and width.

## Features

- Corrects the Y position from the original paragraph baseline
- Keeps the original left/right coordinates so the width is preserved
- Inherits the text frame preferences
- When overset text exists, offers Expand to resolve or Run anyway

## Usage

1. Select exactly one text frame
2. Run the script
3. If there is overset text, choose how to proceed and click OK

## Notes and limitations

- Empty paragraphs are excluded from the output.
- Choosing Run anyway means hidden overset text is not output.
- The whole operation reverts in a single undo.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/text/IdSplitParagraph.jsx` |
| Version | v1.1.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-03-16 |
| Last updated | 2026-06-30 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
