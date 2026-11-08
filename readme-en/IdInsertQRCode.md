# Insert a QR code at the text cursor or into a frame

[![Direct](https://img.shields.io/badge/Direct%20Link-IdInsertQRCode.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdInsertQRCode.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdInsertQRCode.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

Inserts a QR code at the text cursor, into the selected frame, or at the center of the page, with the size and color you specify.

With Object > Generate QR Code you still have to place the code and adjust its size afterward. This script puts it inline at the text cursor, or on the page at a set size, in one step.

### Key features

- The target follows the selection

  | Selection | Inserted into |
  | --- | --- |
  | Text cursor | The cursor position (inline anchored object) |
  | Text | Right after the selected text (the text stays) |
  | Graphic frame (rectangle, ellipse, polygon) or the image in it | That frame (keeps the frame's size) |
  | Nothing | The center of the page in view |

- Two types: Web Hyperlink and Plain Text
- Selected text becomes the initial content (if it is not a URL, the dialog box starts on Plain Text)
- Size in mm (default 20 mm). Use the steppers or the up and down arrow keys; values typed in other units are converted
- Color from the document's swatches (default [Black], with a color chip)
- The QR code's frame is selected afterward
- One undo reverts it

### How to use

1. Place the text cursor where the QR code goes, or select a frame (select nothing for the center of the page)
2. Run the script
3. Choose the type and enter the URL or text
4. Set the size and color, then click OK

### Limitations and notes

- Inserting at the text cursor and into a selected frame has not been tested in InDesign yet.
- A QR code cannot be inserted into text on the pasteboard or in overset text.
- Selecting a text frame does not work. Place the text cursor or clear the selection.
- If the selected frame has contents, the script asks before replacing them (the default is No).
- If the text frame or table row is too small, the inserted QR code becomes overset. Overset QR codes may crash InDesign, so enlarge the frame.
- The content is a single line. Selected text with line breaks or anchored objects is not used as the initial content.
- [None], [Paper] and gradient swatches are left out of the color list.
- For email, text message and business card QR codes, use Object > Generate QR Code.

### Script info

| Item | Details |
| --- | --- |
| File | `jsx/document/IdInsertQRCode.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-06 |
| Last updated | 2026-11-06 |

### Update History

- v1.0.0 (2026-11-06): Initial release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
