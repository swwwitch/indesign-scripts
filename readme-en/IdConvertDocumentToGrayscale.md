# Turn the whole document into CMYK monochrome (text in K100)

[![Direct](https://img.shields.io/badge/Direct%20Link-IdConvertDocumentToGrayscale.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdConvertDocumentToGrayscale.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdConvertDocumentToGrayscale.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

Converts every color in the document to a CMYK gray that uses only K, and sets text to K100 ([Black]). Linked AI, EPS, PSD, TIFF and JPEG files can be turned monochrome in Illustrator and Photoshop and relinked too. Useful for preparing a document for monochrome printing. It has the same purpose as the Illustrator version ([ConvertDocumentToGrayscale](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertDocumentToGrayscale.md)).

### Features

The script runs these steps in order.

1. Sets the transparency blend space (Edit > Transparency Blend Space) to CMYK
2. Converts swatch colors and unnamed colors to CMYK process colors that use only K (RGB, Lab and spot colors too; spot colors become process colors)
3. Sets the fill and stroke of paragraph styles, character styles and local text formatting to K100 ([Black]) (text in tables, footnotes and parent pages too)
4. Merges swatches that ended up with the same gray into one, named "C=0 M=0 Y=0 K=n" (K100 is merged into [Black])
5. Deletes unused swatches

Locked layers and objects are unlocked while processing, then restored.

#### Linked files

When AI, EPS, PSD, TIFF or JPEG files are linked, the script asks whether to convert them too. With Yes, it does the following before the document colors (with No, only the document colors are converted).

1. Sends AI and EPS files to Illustrator through BridgeTalk and runs the work of the Illustrator script ConvertDocumentToGrayscale (Convert to CMYK, CMYK color mode, grayscale contents of used symbols with unused symbols deleted, Convert to Grayscale, K100 text, delete swatches)
2. Sends PSD, TIFF and JPEG files to Photoshop through BridgeTalk, sets Image > Mode > Grayscale and makes the text layers black (100% gray). Layers are kept
3. Saves them in the same folder as "name_gray.ai" / "name_gray.eps" / "name_gray.psd" / "name_gray.tif" / "name_gray.jpg" and closes them (.tiff and .jpeg are saved as .tif and .jpg)
4. Relinks to the saved files

While it runs, a progress bar shows the file being processed. Illustrator and Photoshop are launched when needed if they are not running. At the end, the number of converted files and relinked links, and any files that could not be converted, are shown.

### Usage

1. Open the target document
2. Run the script
3. When AI, EPS, PSD, TIFF or JPEG files are linked, click Yes to convert them too, or No to convert only the document colors

### Notes and limitations

- The whole document is processed regardless of the selection.
- InDesign has no Convert to Grayscale, so the gray is the K value whose lightness (L*) in Lab matches the color (converted with the document's color settings). Colors that already use only K keep their value.
- Only solid fill and stroke colors of text become K100. White ([Paper], K0) stays white (set KEEP_WHITE_TEXT at the top of the script to false to make it K100 too). Text with a gradient keeps the gradient, whose colors become gray.
- Only linked AI, EPS, PSD, TIFF and JPEG files are converted. The contents of PDF, PNG and other files, and embedded links, are not.
- The original linked files are not changed. An existing "_gray" file of the same name is overwritten. Links whose names end in "_gray" (already converted), missing links and files open in Illustrator or Photoshop are skipped.
- AI files are saved with PDF compatibility, EPS files with Illustrator's default EPS settings, and PSD files with layers, alpha channels, spot colors and the color profile. TIFF files keep the same with LZW compression, and JPEG files are saved with the color profile at quality 12 (maximum); re-saving a JPEG loses a little quality.
- Each PSD text layer is made black as a whole (several colors in one layer become a single black). White text stays white.
- Symbol options in AI and EPS files, such as the registration point, the dynamic type and 9-slice scaling, are not carried over.
- When several versions of Illustrator or Photoshop are running, the newest one is used. To pin a version, set ILLUSTRATOR_TARGET / PHOTOSHOP_TARGET at the top of the script to e.g. "illustrator-31.0". Each file is given up to 300 seconds (FILE_TIMEOUT_SECONDS).
- Mixed inks are not converted.
- Colors used in gradient stops or as the base of tints are not deleted even when unused.
- The Find/Change criteria are cleared (the search scope settings are restored).
- A single Undo reverts the relinking and the document conversion. The saved files remain.

### Script info

| Item | Details |
| --- | --- |
| File | `jsx/document/IdConvertDocumentToGrayscale.jsx` |
| Version | v1.2.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-21 |
| Last updated | 2026-11-21 |

### Update History

- v1.2.0 (2026-11-21): Linked TIFF and JPEG files are also converted to grayscale in Photoshop and relinked
- v1.1.0 (2026-11-21): Added turning linked AI, EPS and PSD files monochrome in Illustrator and Photoshop and relinking them (merged IdGrayscaleLinkedFiles)
- v1.0.0 (2026-11-21): Initial release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
