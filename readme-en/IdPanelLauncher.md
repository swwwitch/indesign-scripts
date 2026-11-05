# Search and show panels

[![Direct](https://img.shields.io/badge/Direct%20Link-IdPanelLauncher.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/misc/IdPanelLauncher.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdPanelLauncher.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Lists every panel in the Window menu in a single list, including those inside submenus (Color, Type & Tables, etc.). Filter it with the search field and show the panels you pick.

### Features

- Shows the panels of the Window menu in a three-column list: Name | Category | Open (in menu order)
  - Items inside submenus are expanded; the Category column shows the submenu name
  - Open panels are marked with "✓"
  - Extension panels (CC Libraries, Overlays, etc.) are included
- Filter by name and category with the search field (as you type)
  - Separate words with spaces to match all of them (e.g. "type char")
  - Case-insensitive. Japanese hiragana is matched as katakana, and kanji names also match their reading
  - The (×) button next to the field clears the search
- Switch the sort order with radio buttons
  - **Default**: the order of the Window menu (submenu items stay where the submenu is)
  - **Alphabetical**: by panel name (Japanese names are sorted by their reading)
  - The selected panels stay selected when you re-sort
- Filtering selects the first panel, so Enter shows it right away
- Pick several panels to show them at once
- Picking a panel that is already open closes and reopens it to bring it to the front
- The dialog title shows the number of panels listed

### Usage

1. Run the script
2. Type part of a panel name in the search field
3. Show it in one of these ways
   - Enter / Return
   - Double-click the item in the list
   - Click Show

Press Down in the search field to move to the list. Cmd (Ctrl) / Shift-click to pick more than one.

### Notes

- The list is the section of the Window menu between the first and second separators (where the panels are). Arrange, Workspace, Application Frame, Contextual Task Bar and the document window names are left out. If the separators are not found, every item of the Window menu is listed.
- Showing a panel is the same as choosing its menu item (`app.menuActions.itemByID(…).invoke()`). InDesign's Window menu items toggle the panel, so an open panel is chosen twice to reopen it.
- Readings for the kanji in the Japanese panel names are kept in a table in the script (KANJI_READINGS). Kanji not in the table sort by character code.
- The open mark (✓) reflects the menu check marks when the script starts.
- Panels cannot be opened while a modal dialog is shown, so they are shown after the dialog closes.
- The script uses `#targetengine` because it is meant to be run often. Assigning a keyboard shortcut is handy.
- Esc closes the dialog.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/misc/IdPanelLauncher.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-05 |
| Last updated | 2026-11-05 |
| Tested with | InDesign 2026 (v21.6) / macOS |

### Update History

- v1.0.0 (2026-11-05): Initial release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
