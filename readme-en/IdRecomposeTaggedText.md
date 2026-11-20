# Convert placed tagged text markup into formatting

[![Direct](https://img.shields.io/badge/Direct%20Link-IdRecomposeTaggedText.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdRecomposeTaggedText.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdRecomposeTaggedText.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Re-imports tagged text markup that sits in a text frame as plain text (`<ParaStyle:Heading>`, `<cSize:20>`, and so on), turning the tags into formatting. With nothing selected, every story in the document that contains tag markup is converted.

### Features

- Writes the text to a tagged text file and reads it back in place with InDesign's importer, so every tag works, not just paragraph and character styles
- The scope follows the selection
  - Text frame (including inside groups) or caret: the whole story (or the cell, inside a table)
  - Selected text: just that range
  - Table or cells: the selected cells
  - Nothing selected: every story that contains tag markup
- Paragraphs without tags keep the paragraph style they had before
- The whole run is a single undo step

### Usage

1. Select the text frames, text or cells that contain the markup (deselect everything to process the whole document)
2. Run the script

### Notes and limitations

- Text that contains tables, anchored objects, footnotes, page numbers or other special characters is skipped, because re-importing would lose them. The count is reported.
- Untagged parts are re-imported too, so manual (character-level) formatting applied before conversion is lost. Paragraph styles are kept.
- `<CharStyle:name>` is not applied if the character style does not exist. Create the styles first, or include definitions such as `<DefineParaStyle:…>` in the text.
- A leading encoding header such as `<SJIS-MAC>` is removed before re-importing.
- The tagged text import options (typographer's quotes, remove text formatting) are changed only while the script runs and restored afterwards.
- Nothing is shown when everything converts. Texts that could not be converted, such as locked stories, are counted and reported.
- Based on an idea from id_recomposeTaggedText_a.jsx (z-, 2012).

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/text/IdRecomposeTaggedText.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-21 |
| Last updated | 2026-11-21 |

### Update History

- v1.0.0 (2026-11-21) First release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
