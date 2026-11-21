# Turn the whole document into CMYK monochrome (text in K100)

[![Direct](https://img.shields.io/badge/Direct%20Link-IdConvertDocumentToGrayscale.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdConvertDocumentToGrayscale.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdConvertDocumentToGrayscale.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

Converts every color in the document to a CMYK gray that uses only K, and sets text to K100 ([Black]). Useful for preparing a document for monochrome printing. It has the same purpose as the Illustrator version ([ConvertDocumentToGrayscale](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertDocumentToGrayscale.md)).

### Features

The script runs these steps in order.

1. Sets the transparency blend space (Edit > Transparency Blend Space) to CMYK
2. Converts swatch colors and unnamed colors to CMYK process colors that use only K (RGB, Lab and spot colors too; spot colors become process colors)
3. Sets the fill and stroke of paragraph styles, character styles and local text formatting to K100 ([Black]) (text in tables, footnotes and parent pages too)
4. Merges swatches that ended up with the same gray into one, named "C=0 M=0 Y=0 K=n" (K100 is merged into [Black])
5. Deletes unused swatches

Locked layers and objects are unlocked while processing, then restored.

### Usage

1. Open the target document
2. Run the script

### Notes and limitations

- There is no dialog. The whole document is processed regardless of the selection.
- InDesign has no Convert to Grayscale, so the gray is the K value whose lightness (L*) in Lab matches the color (converted with the document's color settings). Colors that already use only K keep their value.
- Only solid fill and stroke colors of text become K100. White ([Paper], K0) stays white (set KEEP_WHITE_TEXT at the top of the script to false to make it K100 too). Text with a gradient keeps the gradient, whose colors become gray.
- The contents of placed images and graphics (PSD, AI, PDF and so on) are not converted.
- Mixed inks are not converted.
- Colors used in gradient stops or as the base of tints are not deleted even when unused.
- The Find/Change criteria are cleared (the search scope settings are restored).
- A single Undo reverts everything.

### Script info

| Item | Details |
| --- | --- |
| File | `jsx/document/IdConvertDocumentToGrayscale.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-21 |
| Last updated | 2026-11-21 |

### Update History

- v1.0.0 (2026-11-21): Initial release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
