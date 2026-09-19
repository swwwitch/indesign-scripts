# Replace the fonts used in a document in bulk

[![Direct](https://img.shields.io/badge/Direct%20Link-IdReplaceDocumentFonts.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/font/IdReplaceDocumentFonts.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdReplaceDocumentFonts.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Lists the fonts used in the document by family and style, and replaces the selected ones with another font in a single pass.

## Features

- Two lists (source and target) of the fonts in use, with multiple selection on the source side
- Selecting a family row selects every style in that family
- Show PostScript names toggles between family plus style and the PostScript name
- Update paragraph and character styles replaces the font in styles as well as in the text
- Each row shows how many places use the font, and fonts that are missing are marked "not installed"
- Replace All unifies every font in use on a single font
- Master pages, hidden layers, footnotes and table cells are all covered (the script uses InDesign's own find engine)

## Usage

1. Open the target document
2. Run the script
3. Pick the source fonts on the left and the target font on the right, then click Replace Fonts
4. To replace another font, just change the selection — the list refreshes after every replacement

Replace All unifies every font in use on the target font when one is selected, or on the first source font when none is.

## Notes and limitations

- Only fonts already used in the document can be chosen as the target.
- Locked layers and locked stories are left untouched, although they are included in the counts.
- Undo takes one step per source font.
- With Update paragraph and character styles off, the replacement becomes a local override: reapplying the paragraph style restores the original font.
- Members of a composite font cannot be replaced.
- Usage counts come from one search per font, so building the list takes a while on documents with many fonts or pages. Set `SHOW_USAGE_COUNT` to `false` near the top of the script to skip the counts.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/font/IdReplaceDocumentFonts.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-09-20 |
| Last updated | 2026-09-20 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
