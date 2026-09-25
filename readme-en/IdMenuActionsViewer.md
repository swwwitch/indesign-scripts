# Look up menu actions and copy their invoke code

[![Direct](https://img.shields.io/badge/Direct%20Link-IdMenuActionsViewer.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/misc/IdMenuActionsViewer.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdMenuActionsViewer.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Lists InDesign menu actions (menu commands) and lets you filter them by area, menu, name, and ID. You can copy the key strings (`$ID/…`) of the selected action and the code that invokes it from a script.

Based on Peter Kahrel's `menu_actions.jsx`.

## Features

- Shows menu actions in a two-column list: Area | Name
- Filters
  - **Area**: pick one area
  - **Menu**: pick File / Edit / View / Format / Find / Tools / Window / Panel Menus to show actions whose area starts with that menu name. The area choices are narrowed as well
  - **Name / ID**: text contained in the name (case-insensitive; regular expressions allowed). Digits only also match the action with that ID
- Switch the sort order with radio buttons (Area / Name / ID)
- The selected action's details are shown below the list automatically
  - **Key strings**: e.g. `$ID/Find/Change...` (multiple key strings are separated by " | ")
  - **Invoke code**: `app.menuActions.itemByID(18694).invoke();`
- Copy either one to the clipboard with its Copy button
- When the filter leaves a single action, it is selected and its details are shown
- A progress bar is shown while loading
- The dialog title shows the number of listed actions

## Usage

1. Run the script
2. Narrow down the actions by area, menu, name, or ID (press Enter to apply Name / ID)
3. Select an action in the list
4. Click Copy next to the invoke code or key strings and paste it into your script

The copied code can be used like this:

```javascript
app.menuActions.itemByID(18694).invoke();
app.menuActions.item("$ID/Find/Change...").invoke();
```

## Notes

- Font and style names (IDs 57603–61066), recent files and scripts (`.indd`, `.jsx`, `.jsxbin`), and the English "Text Selection" and "Menu:Insert" areas are not listed.
- The Menu filter matches the beginning of the area name. Both Japanese and English names are checked, but some menus may match nothing depending on the InDesign version and language.
- Action IDs can differ between InDesign versions and environments. For scripts you distribute, consider using the key strings (`$ID/…`) instead.
- Copying to the clipboard uses AppleScript on Mac and VBScript on Windows.
- Name / ID filtering is applied when you press Enter, to avoid rebuilding thousands of rows on every keystroke.

## Original / Credits

Based on Peter Kahrel's `menu_actions.jsx`.

- https://creativepro.com/menu_actions/
- http://kasyan.ho.com.ua/open_menu_item.html

## Script info

| Item | Details |
| --- | --- |
| File | `jsx/misc/IdMenuActionsViewer.jsx` |
| Version | v1.0.14 |
| Original author | Peter Kahrel |
| Modified by | Masahiro Takano (@swwwitch) |
| First release | 2026-09-25 |
| Last updated | 2026-09-25 |

## Update history

- v1.0.14 (2026-09-25): Initial release

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
