# Look up menu actions and copy their invoke code

[![Direct](https://img.shields.io/badge/Direct%20Link-IdMenuActionsViewer.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/misc/IdMenuActionsViewer.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdMenuActionsViewer.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Lists InDesign menu actions (menu commands) and lets you filter them by category, area, name, and ID. You can copy the key strings (`$ID/…`) of the selected action and the code that invokes it from a script.

Based on Peter Kahrel's `menu_actions.jsx`.

![Dialog with Category, Area, Name / ID, and Sort by at the top, a list in the middle, and Key strings and Invoke code fields at the bottom](../png/ss-1184-1698-144-20260925-094953.png)

## Features

- Shows menu actions in a two-column list: Name | Area
- Filters
  - **Category**: the part of the area name before ":" (Edit Menu, Panel Menus, etc.). Menu-related categories come first, the rest below a divider. The area choices are narrowed as well
  - **Area**: pick one area
  - **Name / ID**: text contained in the name (case-insensitive; regular expressions allowed). Digits only also match the action with that ID
- Switch the sort order with radio buttons (Name / Area; Area by default)
- The selected action's details are shown below the list automatically
  - **Key strings**: e.g. `$ID/Find/Change...` (multiple key strings are separated by " | ")
  - **Invoke code**: `app.menuActions.itemByID(18694).invoke();`
- Copy either one to the clipboard with its Copy button
- When the filter leaves a single action, it is selected and its details are shown
- A progress bar is shown while loading
- The dialog title shows the number of listed actions

## Usage

1. Run the script
2. Narrow down the actions by category, area, name, or ID (Name / ID filters as you type)
3. Select an action in the list
4. Click Copy next to the invoke code or key strings and paste it into your script

The copied code can be used like this:

```javascript
app.menuActions.itemByID(18694).invoke();
app.menuActions.item("$ID/Find/Change...").invoke();
```

## Notes

- Font and style names (IDs 57603–61066), recent files and scripts (`.indd`, `.jsx`, `.jsxbin`), names without any letter or digit (such as "(" or ")"), document names in the Window menu, and the English "Text Selection" and "Menu:Insert" areas are not listed.
- Names longer than 20 characters are shortened with "…" in the list, because ScriptUI on Mac widens the first column to its longest text. Filtering uses the full name.
- The category choices are built from the actual area names, so they vary with the InDesign version and language.
- Action IDs can differ between InDesign versions and environments. For scripts you distribute, consider using the key strings (`$ID/…`) instead.
- Copying to the clipboard uses AppleScript on Mac and VBScript on Windows.
- Name / ID rebuilds the list on every keystroke. With many actions listed, each keystroke may take a moment.
- Enter / Return and Esc close the dialog.

## Original / Credits

Based on Peter Kahrel's `menu_actions.jsx`. Many thanks to Peter Kahrel for sharing his scripts.

- https://creativepro.com/menu_actions/
- http://kasyan.ho.com.ua/open_menu_item.html

Main changes from the original:

- Modal dialog instead of a persistent palette
- Filters by category, area, and name / ID (can be combined)
- Much faster sorting (about 4 s → 0.05 s), switchable between name and area
- Shows and copies the selected action's key strings and invoke code (`app.menuActions.itemByID(…).invoke();`)
- Japanese / English UI

## Script info

| Item | Details |
| --- | --- |
| File | `jsx/misc/IdMenuActionsViewer.jsx` |
| Version | v1.0.24 |
| Original author | Peter Kahrel |
| Modified by | Masahiro Takano (@swwwitch) |
| First release | 2026-09-25 |
| Last updated | 2026-09-25 |
| Article | https://note.com/dtp_tranist/n/n5038c9d2cc85 |

## Update history

- v1.0.24 (2026-09-25): Initial release

## License

The original `menu_actions.jsx` is copyright Peter Kahrel. This modified version is released under the MIT License with the permission of the original author.

- Modifications: Copyright (c) 2026 Masahiro Takano (@swwwitch)
- MIT License — <http://opensource.org/licenses/mit-license.php>
