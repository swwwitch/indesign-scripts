# Search and show panels

[![Direct](https://img.shields.io/badge/Direct%20Link-IdPanelLauncher.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/misc/IdPanelLauncher.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdPanelLauncher.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Lists every panel in the Window menu in a single list, including those inside submenus (Color, Type & Tables, etc.). Filter it with the search field and show the panels you pick.

### Features

- Shows the panels of the Window menu in a four-column list: Name | Mark | Category | Open (in menu order)
  - Items inside submenus are expanded; the Category column shows the submenu name
  - Open panels are marked with "✓"
  - The Mark column shows ♥ (frequently used) or ♢ (rarely used)
  - Extension panels (CC Libraries, Overlays, etc.) are included
- Filter by name and category with the search field (as you type)
  - Separate words with spaces to match all of them (e.g. "type char")
  - Case-insensitive. Japanese hiragana is matched as katakana, and kanji names also match their reading
  - The (×) button next to the field clears the search
- Mark panels and filter by the marks
  - Select panels in the list and click ♥ at the bottom left for ♥ (frequently used) or ♢ for ♢ (rarely used). If all the selected panels already have the mark, clicking again removes it
  - A panel has either ♥ or ♢. You can mark several panels at once
  - Filter checkboxes: **♥ Frequent** (only ♥ panels) / **Hide ♢** (leaves out ♢ panels). With both off, every panel is listed
  - The marks and the filter are kept for the next run
- Set buttons show a fixed group of panels at once (regardless of the filter, search or selection; the dialog closes and the panels open)
  - **Tables**: Cell Styles, Table Styles, Table
  - **Text**: Character, Paragraph, Character Styles, Paragraph Styles, Glyphs, Text Wrap
  - **Book**: Cross-References, Hyperlinks
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
- Set panels are looked up by panel name (Japanese or English). Panels that are not found are listed under "Could not show" after the dialog closes. Edit PANEL_SETS in the script to change the groups.
- The ♥ / ♢ marks and the filter are saved to `~/Library/Application Support/IdPanelLauncher-settings.txt` (`Folder.userData`). Marks are keyed by category and panel name, so they do not carry over when you switch InDesign's UI language.
- Panels cannot be opened while a modal dialog is shown, so they are shown after the dialog closes.
- The script uses `#targetengine` because it is meant to be run often. Assigning a keyboard shortcut is handy.
- Esc closes the dialog.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/misc/IdPanelLauncher.jsx` |
| Version | v1.1.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-05 |
| Last updated | 2026-11-06 |
| Tested with | InDesign 2026 (v21.6) / macOS |

### Update History

- v1.1.1 (2026-11-06): The filter is now two checkboxes instead of radio buttons, so ♥ Frequent and Hide ♢ can be on together
- v1.1.0 (2026-11-06): Added the ♥ (frequently used) and ♢ (rarely used) marks and filtering by them, and set buttons that show a fixed group of panels at once (Tables, Text, Book)
- v1.0.0 (2026-11-05): Initial release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
