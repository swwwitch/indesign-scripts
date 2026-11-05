# List the URLs of QR codes in a document

[![Direct](https://img.shields.io/badge/Direct%20Link-IdQRCodeURLList.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdQRCodeURLList.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdQRCodeURLList.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

Lists, in page order, the URLs of the QR codes in the active document (made with Object > Generate QR Code).

A QR code does not show its URL, and checking it normally means opening Edit QR Code for each one. This script shows the URLs of all QR codes in the document at once, which helps when checking link targets before sending files to print or making a checklist.

### Features

- Shows a three-column list: Page | Type | URL
- Covers QR codes on document pages, parent pages and the pasteboard, inside groups, as anchored objects, and on locked or hidden layers
- Lists "Web Hyperlink" QR codes, plus "Plain Text" QR codes whose text is a URL (starting with `http://`, `https://` or `ftp://`)
- Order: document pages, then parent pages, then the pasteboard; within a page, top to bottom, then left to right
- Copy All puts "page<tab>URL" lines on the clipboard, ready to paste into a spreadsheet
- Select (or a double-click) shows the page of that QR code and selects its frame
- The dialog title shows the count

### How to use

1. Open a document and run the script
2. Check the URLs in the list
3. Click Copy All to copy them to the clipboard if needed
4. Choose a QR code and click Select (or double-click it). The dialog box closes, its page is shown and the frame is selected

Enter/Return works as Select, and Esc as Close.

### Notes

- The contents of a QR code are not exposed to scripts as a property, so the script exports each QR code frame to a temporary snippet (.idms), reads it and deletes it right away. This takes about 0.01 seconds per QR code.
- The document is not changed (Select only changes the selection and the page shown).
- QR codes in overset text are not read, because exporting them as a snippet may crash InDesign; the URL column shows "(cannot be read in overset text)" instead (not verified).
- Email, text message and vCard QR codes are not listed.
- Placed (linked) EPS files are ignored; QR codes are stored as embedded EPS.
- QR codes on locked or hidden layers are listed, but Select may not be able to select them.
- The list reflects the moment you run the script. Run it again after editing QR codes.
- Copying to the clipboard goes through AppleScript on Mac and VBScript on Windows.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/document/IdQRCodeURLList.jsx` |
| Version | v1.0.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-05 |
| Last updated | 2026-11-05 |
| Article | https://note.com/dtp_tranist/n/ne2cdd5f60080 |
| Tested with | InDesign 2026 (v21.6) / macOS |

### Update History

- v1.0.1 (2026-11-05) Moved the version from the title bar to the bottom left of the dialog, dimmed (shared part added)
- v1.0.0 (2026-11-05): Initial release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
