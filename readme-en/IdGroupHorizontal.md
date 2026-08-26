# Group objects into rows by vertical position

[![Direct](https://img.shields.io/badge/Direct%20Link-IdGroupHorizontal.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/group/IdGroupHorizontal.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdGroupHorizontal.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Splits the selected objects into rows by vertical proximity and groups each row.

## Features

- Objects whose centre Y values differ by less than the tolerance count as one row
- Only rows with two or more objects are grouped
- Reports the number of groups created

## Usage

1. Select the objects you want to group
2. Run the script

## Notes and limitations

- Change the tolerance via `ROW_TOLERANCE` at the top of the script (current ruler units).
- Use IdSmartGroup.jsx when you need finer control.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/group/IdGroupHorizontal.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-04-11 |
| Last updated | 2026-04-17 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
