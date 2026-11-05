# List text anchors and jump to them

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTextAnchorListPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdTextAnchorListPalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTextAnchorListPalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

Lists the text anchors in the active document (made with New Hyperlink Destination in the Hyperlinks panel menu, with the type set to Text Anchor) in page order in a floating palette, with three columns: page, anchor name and text.

Hyperlink Destination Options only shows anchor names, so finding the page and paragraph of each one normally means going to them one by one. This script shows where every anchor is and its paragraph at once, and with the palette left open, a double-click takes you to one after another. Handy for checking the targets of cross-references and hyperlinks.

### Features

- Shows a three-column list: Page | Anchor Name | Text (the paragraph that holds the anchor, shortened with "..." when long)
- Covers anchors on document pages, parent pages and the pasteboard, in tables and inside anchored objects
- Order: document pages, then parent pages, then the pasteboard, then overset text; within a page, frames top to bottom, then left to right, and text order within a frame
- A double-click (or Show) shows the page of that anchor and selects the text of the paragraph that holds it (without the paragraph return). The palette stays open, so you can go to another anchor right away
- If another document is in front, it switches to the listed document first
- Refresh reads the anchors of the active document again, and chooses the same anchor again if one with that name is still there
- The palette title shows the count

### How to use

1. Open a document and run the script
2. Check where each anchor is and its text
3. Double-click the anchor to go to (or choose it and click Show). Its page is shown and the text of the paragraph that holds the anchor is selected
4. Double-click another anchor to go on. Click Refresh after adding or removing anchors, or to look at another document
5. Click Close when you are done

Esc works as Close.

### Notes

- The document is not changed (Show only changes the page shown and the selection).
- Hidden hyperlink destinations are not listed (set `INCLUDE_HIDDEN_ANCHORS` at the top of the script to `true` to include them).
- Paragraph destinations (used by cross-references to paragraphs), URL destinations and page destinations are not listed.
- Anchors in overset text are listed as "Overset" but cannot be shown.
- The text may not be selectable for anchors on locked or hidden layers.
- For an anchor in an empty paragraph, the text cursor is placed at the anchor instead.
- The list is not updated automatically. Click Refresh after adding, removing or moving anchors, or editing the text. If the listed document has been closed or the anchor removed, showing it asks you to click Refresh.
- Running the script again while the palette is open closes it and opens a new one.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/document/IdTextAnchorListPalette.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-06 |
| Last updated | 2026-11-06 |
| Article | https://note.com/dtp_tranist/n/n23fa2f14621f |
| Tested with | InDesign 2026 (v21.6) / macOS |

### Update History

- v1.0.0 (2026-11-06): Initial release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
