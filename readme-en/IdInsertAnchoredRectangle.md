# Create an anchored rectangle and a text anchor at the cursor

[![Direct](https://img.shields.io/badge/Direct%20Link-IdInsertAnchoredRectangle.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdInsertAnchoredRectangle.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdInsertAnchoredRectangle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

Creates a small anchored (inline) rectangle at the text cursor, plus a text anchor right before it.

The rectangle and the text anchor share the same name, so you can use the anchor as a target for hyperlinks and cross-references while seeing its position on the page as a rectangle.

### Features

- Inserts a 12 × 12 pt rectangle with a 0.5 pt stroke and no fill as an inline anchored object at the cursor (or at the start of the selected text)
- Creates a text anchor right before the rectangle's marker character
- The text anchor is named "anchored_rectangle_<time>"; the same name goes into the rectangle's script label
- Selects the rectangle afterwards and shows the text anchor name
- A single Undo reverts everything

### How to use

1. Place the text cursor where you want the target (or select text) with the Type tool
2. Run the script
3. The rectangle is inserted and the new text anchor name is shown
4. Choose this text anchor as the destination in the Hyperlinks panel or Insert Cross-Reference

### Notes

- Does not run when a text frame is selected with the Selection tool, or when an object is selected.
- Cannot be used on text on the pasteboard or in overset text.
- The stroke color follows the document's default graphic frame settings.
- The size, stroke weight and name prefix can be changed in the settings at the top of the script.
- Not yet tested in InDesign.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/document/IdInsertAnchoredRectangle.jsx` |
| Version | v2.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-06 |
| Last updated | 2026-11-06 |

### Update History

- v2.0.0 (2026-11-06): Renamed from IdInsertPageCrossReference and replaced with creating an anchored rectangle and a text anchor at the cursor
- v1.0.0 (2026-11-06): Initial release (inserted a paragraph's page number as a "p.<number>" cross-reference)

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
