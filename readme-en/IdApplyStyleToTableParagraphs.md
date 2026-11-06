# Apply a paragraph style to paragraphs containing only a table

[![Direct](https://img.shields.io/badge/Direct%20Link-IdApplyStyleToTableParagraphs.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdApplyStyleToTableParagraphs.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdApplyStyleToTableParagraphs.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

Finds the paragraphs that contain nothing but a table in the selection, the story or the document, and applies the paragraph style you choose.

### Features

- Choose the scope: selection, story or document (tables inside cells are included)
- Choose the paragraph style from a menu ("p.table" is selected first)
- Applies the style when a paragraph holds only tables (the trailing return is ignored), including paragraphs with two or more tables in a row
- Clears manual formatting (overrides) on the paragraphs it changes
- A single undo reverts the whole run

### How to use

1. Select the text, text frames or tables to check (not needed for the whole document)
2. Run the script
3. Set Scope and Paragraph Style to Apply, then click OK
4. The number of paragraphs changed is shown

### Limitations and notes

- Selection is available when text, a text frame, a table or cells are selected. With only a cursor, Story is selected first.
- Story is available when text or a table is selected, or when there is a cursor.
- A table counts when its anchor character is inside the selected text. Selecting only the text inside a cell does not include that table itself.
- If the document has no "p.table" style, OK stays disabled until you choose a style. To change the initial style, edit `DEFAULT_PARAGRAPH_STYLE_NAME` at the top of the script.
- A paragraph is skipped if it also holds text, spaces, anchored objects or a break character (such as a page break).
- A line separated only by a forced line break is skipped, because the paragraph as a whole holds other characters.
- Paragraphs that cannot be changed, such as those in locked stories, are skipped and counted in the message. Text on locked layers is still changed.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/style/IdApplyStyleToTableParagraphs.jsx` |
| Version | v1.1.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-06 |
| Last updated | 2026-11-06 |
| Article | https://note.com/dtp_tranist/n/na8313158d9ce |

### Update History

- v1.1.0 (2026-11-06) Added a dialog box to choose the scope (selection, story or document) and the paragraph style
- v1.0.0 (2026-11-06) First release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
