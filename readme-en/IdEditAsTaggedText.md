# Edit the selected text as tagged text

[![Direct](https://img.shields.io/badge/Direct%20Link-IdEditAsTaggedText.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdEditAsTaggedText.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdEditAsTaggedText.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Exports the selected text as tagged text and shows it in a dialog box. Edit the tags or the text and click OK to read it back in place with the new formatting.

### Features

- The range follows the selection
  - Selected text: that range
  - Caret or text frame: the whole story (or the cell, inside a table)
- Japanese and other non-ASCII characters are shown as characters, not `<0x6587>` (special spaces and control characters stay as `<0x2002>` and so on)
- Each line is a paragraph; `<0x000A>` is a forced line break
- If you delete the `<ParaStyle:…>` at the start, the original paragraph style is kept
- Reset restores the text as it was before editing
- The whole run is a single undo step

### Usage

1. Select the text to edit (or select a text frame, or place the caret)
2. Run the script
3. Edit the tagged text in the dialog box and click OK

### Notes and limitations

- Text that contains tables, anchored objects, footnotes, page numbers or other special characters cannot be edited, because re-importing would lose them.
- Type a literal `<`, `>` or `\` as `\<`, `\>` or `\\`, as tagged text requires.
- Tags that name a paragraph or character style that does not exist are not applied.
- Tags are shown in the verbose form (`<cSize:>` etc.). For the abbreviated form (`<cs:>` etc.), set `TAG_FORM` at the top of the script to `TagTextForm.ABBREVIATED`.
- Clicking OK without changes does nothing.
- The tagged text export and import options (character set, tag form, typographer's quotes, etc.) are changed only while the script runs and restored afterwards.
- If the import fails, the original text is read back in.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/text/IdEditAsTaggedText.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-21 |
| Last updated | 2026-11-21 |

### Update History

- v1.0.0 (2026-11-21) First release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
