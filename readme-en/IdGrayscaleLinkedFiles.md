# Turn linked AI, EPS and PSD files into monochrome and relink

[![Direct](https://img.shields.io/badge/Direct%20Link-IdGrayscaleLinkedFiles.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdGrayscaleLinkedFiles.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdGrayscaleLinkedFiles.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

Turns the AI and EPS files linked in the document into CMYK monochrome (text in K100) in Illustrator and the linked PSD files into grayscale in Photoshop, saves them under a name ending in "_gray" and relinks to them. Combined with [Turn the whole document into CMYK monochrome](IdConvertDocumentToGrayscale.md), placed illustrations and images become monochrome too.

### Features

1. Collects the AI, EPS and PSD links in the document (a file placed several times is processed once)
2. Sends AI and EPS files to Illustrator through BridgeTalk, opens them and runs these steps (the same as the Illustrator script [ConvertDocumentToGrayscale](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertDocumentToGrayscale.md))
   - Edit > Edit Colors > Convert to CMYK, and CMYK document color mode
   - Deletes unused symbols and converts the contents of the used ones to grayscale
   - Edit > Edit Colors > Convert to Grayscale
   - Sets the fill and stroke of text to K100
   - Deletes the swatches
3. Sends PSD files to Photoshop through BridgeTalk, sets Image > Mode > Grayscale and makes the text layers black (100% gray). Layers are kept
4. Saves them in the same folder as "name_gray.ai" / "name_gray.eps" / "name_gray.psd" and closes them
5. Relinks the InDesign links to the saved files

In Illustrator, locked or hidden layers and objects are unlocked while processing, then restored. In Photoshop, locked text layers are made black too.

### Usage

1. Open the target document
2. Run the script
3. Check the number of files and click Yes

Illustrator and Photoshop are launched when needed if they are not running. While it runs, a progress bar shows the file being processed. At the end, the number of converted files and relinked links, and any files that could not be converted, are shown.

### Notes and limitations

- The original files are not changed. An existing "_gray" file of the same name is overwritten.
- Links whose names end in "_gray" (already converted), embedded links and missing links are skipped.
- Files that are open in Illustrator or Photoshop are skipped to protect unsaved edits. Close them and run again.
- AI files are saved with PDF compatibility. EPS files are saved with Illustrator's default EPS settings. PSD files are saved with layers, alpha channels, spot colors and the color profile.
- White text (0% gray) stays white in AI, EPS and PSD files alike (set KEEP_WHITE_TEXT at the top of the script to false to make it K100 too).
- Symbol options such as the registration point, the dynamic type and 9-slice scaling are not carried over. Linked images, the contents of patterns and brushes, and graphic styles are not converted.
- Each PSD text layer is made black as a whole (several colors in one layer become a single black). Adjustment layers and smart objects are kept, and no confirmation is shown when converting to grayscale.
- When several versions are running, the newest one is used. To pin a version, set ILLUSTRATOR_TARGET / PHOTOSHOP_TARGET at the top of the script to e.g. "illustrator-31.0".
- Each file is given up to 300 seconds (FILE_TIMEOUT_SECONDS).
- A single Undo reverts the relinking. The saved files remain.

### Script info

| Item | Details |
| --- | --- |
| File | `jsx/document/IdGrayscaleLinkedFiles.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-21 |
| Last updated | 2026-11-21 |

### Update History

- v1.0.0 (2026-11-21): Initial release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
