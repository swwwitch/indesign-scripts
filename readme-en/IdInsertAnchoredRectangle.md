# Create an anchored rectangle and a text anchor at the cursor

[![Direct](https://img.shields.io/badge/Direct%20Link-IdInsertAnchoredRectangle.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdInsertAnchoredRectangle.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdInsertAnchoredRectangle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

Creates a small non-printing anchored (custom) rectangle at the text cursor, plus a text anchor right before it.

The anchor name and the rectangle's fill and stroke are set in a dialog box. You can use the anchor as a target for hyperlinks and cross-references while seeing its position on the page as a rectangle.

### Features

- Inserts a 12 × 12 pt rectangle as an anchored object at the cursor (or at the start of the selected text), with Position set to Custom in Anchored Object Options
- The rectangle's vertical position follows the line with the cursor (its bottom sits on the line's baseline) and moves with the line when text reflows
- The rectangle's horizontal position is chosen in the dialog box: Left of page or Right of page (default: left). The rectangle is centered halfway between the page edge and the margin
- Fill and stroke are chosen from the document's swatches. Defaults: fill "C=0 M=100 Y=100 K=0" (created if the document does not have it), stroke [None]
- A color chip next to each pop-up menu shows the chosen swatch (an approximate color for display; [None] is white)
- When a stroke is chosen, its weight can be set (default 0.5 pt)
- Sets the rectangle to Nonprinting (Attributes panel), so it does not appear in print or PDF export
- Creates a text anchor right before the rectangle's marker character
- The text anchor name can be edited in the dialog box. Default: p + page number + _ + sequence number (e.g. `p12_01`; the sequence number is the lowest free number on that page)
- The same name goes into the rectangle's script label
- Selects the rectangle afterwards
- A single Undo reverts everything

### How to use

1. Place the text cursor where you want the target (or select text) with the Type tool
2. Run the script
3. Check the name, fill, stroke and position in the dialog box and click OK
4. The rectangle and the text anchor are created
5. Choose this text anchor as the destination in the Hyperlinks panel or Insert Cross-Reference

### Notes

- Does not run when a text frame is selected with the Selection tool, or when an object is selected.
- Cannot be used on text on the pasteboard or in overset text.
- The horizontal position comes from the margins of the page at creation. If text reflows onto a page with different left and right margins, the rectangle is no longer halfway.
- If a text anchor with the same name exists, OK shows a message and returns to the dialog box.
- The size, default fill, default stroke weight and name form can be changed in the settings at the top of the script.
- Not yet tested in InDesign.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/document/IdInsertAnchoredRectangle.jsx` |
| Version | v2.7.2 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-06 |
| Last updated | 2026-11-06 |

### Update History

- v2.7.2 (2026-11-06): Fixed the color chip being gray for every swatch
- v2.7.1 (2026-11-06): Fixed the color chip not changing when another swatch was chosen
- v2.7.0 (2026-11-06): Added color chips next to the fill and stroke pop-up menus
- v2.6.0 (2026-11-06): Moved the rectangle from the page edge to halfway between the page edge and the margin
- v2.5.0 (2026-11-06): Added Position (Left of page / Right of page) to the dialog box
- v2.4.0 (2026-11-06): Fixed the rectangle being created at the top left of the page; its vertical position now follows the line with the cursor (bottom on the line's baseline)
- v2.3.0 (2026-11-06): Added a dialog box for the text anchor name (default: page number and sequence number) and the rectangle's fill, stroke and stroke weight. Removed the message shown after creation. Fixed the rectangle size being read in the document's units instead of pt
- v2.2.0 (2026-11-06): The rectangle is now filled with M100 Y100 (swatch "C=0 M=100 Y=100 K=0", created if missing) with no stroke
- v2.1.0 (2026-11-06): Changed the anchored object position from Inline to Custom and set the rectangle to Nonprinting
- v2.0.0 (2026-11-06): Renamed from IdInsertPageCrossReference and replaced with creating an anchored rectangle and a text anchor at the cursor
- v1.0.0 (2026-11-06): Initial release (inserted a paragraph's page number as a "p.<number>" cross-reference)

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
