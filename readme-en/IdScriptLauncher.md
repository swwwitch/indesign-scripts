# Filter scripts by keyword and run them

[![Direct](https://img.shields.io/badge/Direct%20Link-IdScriptLauncher.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/misc/IdScriptLauncher.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdScriptLauncher.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

A launcher that filters .jsx / .js / .jsxbin files in a chosen folder by keyword and runs the selected script on the spot. Folders and file names are shown in two side-by-side lists.

## Features

- Typing in the keyword field filters both the folder list and the file name list instantly (space-separated terms are AND-matched)
- Matching ignores case, full-width vs. half-width, hiragana vs. katakana, and voiced/small kana differences
- Words that appear often in the filtered results become one-click keyword buttons automatically
- "Include subdirectories" switches between the folders directly under the target folder and the whole tree
- "Full path" switches the target folder display between the full path and a `~`-abbreviated one
- Preferences let you change the target folder and the keyword button rules (occurrences and number of keywords)
- The target folder and keyword button settings are restored the next time you run the script

## Usage

1. Run the script
2. On the first run, choose the script folder to search (later runs reopen the last folder)
3. Type a keyword or click a keyword button to filter
4. Select a file name and click [Run], or double-click it

### Keyboard

| Key | Action |
| --- | --- |
| `↓` (keyword field) | Move to the file name list |
| `Enter` | Run the selected script (from anywhere in the dialog) |
| `↑` `↓` in a Preferences number field | ±1 (`shift` for ±10, `option` for ±0.1) |

### Mouse

| Action | Result |
| --- | --- |
| Double-click a file name | Run it |
| `option` + double-click a file name | Reveal it in the Finder instead of running it |
| Double-click a folder | Open that folder in the Finder |
| `option` + keyword button | Append the word to the current keyword instead of replacing it (AND search) |

## Notes and limitations

- Undo grouping is left to the launched script (it is not wrapped in `doScript`).
- Only ASCII words of three characters or more become keyword buttons. Japanese file names can still be filtered, but they never become buttons.
- If the launcher itself sits in the target folder, it is left out of the list.
- Alias folders are not followed, to avoid loops.
- "Reveal in Finder" uses `/Applications/RevealInFinder.app`. Without it, the enclosing folder is opened instead (as on platforms other than macOS).
- Settings are stored in `IdScriptLauncher-prefs.txt` in the user data folder.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/misc/IdScriptLauncher.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-08-26 |
| Last updated | 2026-08-26 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
