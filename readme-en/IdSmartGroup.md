# Group nearby objects into rows or columns

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSmartGroup.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/group/IdSmartGroup.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSmartGroup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Groups the selected objects by horizontal (row) or vertical (column) proximity, with a live red-frame preview of each group while the dialog is open.

### Features

- Switch the direction between horizontal and vertical with radio buttons
- Adjust the tolerance from 0 to 50 with a slider (in ruler units)
- Group extents are drawn live as red frames on a non-printing layer at the top (toggle with Show preview)
- The preview is removed automatically when the dialog closes

### Usage

1. Select the objects you want to group
2. Run the script
3. Adjust the direction and tolerance, check the preview, then click OK

### Notes and limitations

- The preview lives on a non-printing layer named "SmartGroup Preview" and is deleted on exit.
- A group needs at least two objects.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/group/IdSmartGroup.jsx` |
| Version | v1.1.2 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-04-11 |
| Last updated | 2026-10-01 |

### Update History
- v1.1.3 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)

#### v1.1.2 (2026-10-01)

- A button row with only right-side buttons is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones

#### v1.1.1 (2026-09-30)

- Unified the button row, spacing and localization code with the shared parts (no change in behavior)

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
