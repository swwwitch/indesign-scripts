# IdAutoParagraphStyleGenerator

[![Direct](https://img.shields.io/badge/Direct%20Link-IdAutoParagraphStyleGenerator.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/IdAutoParagraphStyleGenerator.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAutoParagraphStyleGenerator.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Groups unstyled paragraphs (Basic Paragraph and friends) by font, size and leading, then generates and applies a paragraph style for each group.

## Features

- Only font family/style, point size and leading are considered
- Grouping is normalized to points so pt/Q rounding does not split groups
- Style names follow the text-size unit set in Units & Increments (pt or Q)
- Paragraphs with mixed formatting are skipped and reported

## Usage

1. Open the target document
2. Run the script
3. Check the created and applied counts in the result dialog

## Notes and limitations

- Generated names look like `AutoStyle_1_10pt`.
- A numeric suffix is added when a style of the same name already exists.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/IdAutoParagraphStyleGenerator.jsx` |
| Version | v3.4 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-02-13 |
| Last updated | 2026-03-14 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
