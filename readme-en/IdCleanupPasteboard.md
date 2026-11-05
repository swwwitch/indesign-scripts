# Delete Objects Outside the Pages (on the Pasteboard)

[![Direct](https://img.shields.io/badge/Direct%20Link-IdCleanupPasteboard.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/misc/IdCleanupPasteboard.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCleanupPasteboard.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Deletes the objects left outside the pages (on the pasteboard) from the document in one go.

### Features

- **Scope**: the current spread (the parent spread, if you are editing one), the entire document (default), or all open documents
- **Delete Outside This Area**
  - Bleed: beyond the document's bleed
  - Slug: beyond the document's slug
  - Distance from pages (default): beyond a given distance (mm) from the page edges. Use the up/down buttons or the arrow keys to change it (Shift snaps to 10s, Option steps by 0.1). Other units (such as "1 in") and arithmetic are converted to mm
  - On facing pages, the inside and outside bleed/slug values are applied to the correct side of each page
- **Options**
  - Include parent pages: also process the parent spreads (not available with Current spread)
  - Keep threaded text frames (on by default): keep text frames threaded to other frames, and groups that contain them
  - Keep hidden objects (on by default): keep hidden objects and objects on hidden layers
  - Delete locked objects: unlock and delete locked objects too
  - Confirm before deleting (on by default): shows the number of objects and lets you choose Delete, Select and Show, or Cancel
- **Layers to Keep**: objects on the layers selected in the list (Cmd/Ctrl-click for several) are kept. With All open documents, layers with the same names are kept
- Remembers the dialog settings for the next run
- Shows the number of deleted objects when done

### Usage

1. Open a document and run the script
2. Choose the scope, area, options and layers to keep, then click OK
3. In the confirmation dialog, click Delete. To see the objects first, click Select and Show

### Notes

- Only objects entirely outside the area are deleted. Objects that overlap a page (or the area) even slightly, or only touch its edge, are kept.
- Objects are tested by their visible bounds (including the stroke). Groups are tested as a whole, not item by item.
- Objects on locked layers are never deleted (and not counted).
- Select and Show shows the first spread with objects to delete and selects them (locked and hidden objects are not selected). It is not available with All open documents.
- Works with right-to-left binding and with spreads that contain pages of different sizes.
- Works with any ruler units (including Q/H and pixels). The units and ruler origin are switched only while objects are checked and are restored afterwards.
- The layer list shows the layers of the current document.
- Settings are saved to `IdCleanupPasteboard-prefs.txt` in `Folder.userData` (~/Library/Application Support on Mac).
- Guides are not affected.
- The deletion can be undone with a single Undo.

### Reference

Marc Autret's CleanupPasteboard is a script with the same purpose. This script was written from a spec, without using its code.

- http://indiscripts.com

### Script info

| Item | Details |
| --- | --- |
| File | `jsx/misc/IdCleanupPasteboard.jsx` |
| Version | v1.0.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-10-29 |
| Last updated | 2026-11-05 |

### Update History

- v1.0.1 (2026-11-05) Moved the version from the title bar to the bottom left of the dialog, dimmed (shared part added)
- v1.0.0 (2026-10-29): Initial release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
