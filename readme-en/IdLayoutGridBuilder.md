# IdLayoutGridBuilder

[![Direct](https://img.shields.io/badge/Direct%20Link-IdLayoutGridBuilder.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/IdLayoutGridBuilder.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdLayoutGridBuilder.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Builds a content area, title area, frames, a column/row grid and dividers on the active page, with a live preview.

## Features

- Configure the content area (border, corner radius, extension, cap style), the title area and the footer column area
- Set page margins, frame insets and content-region offsets independently
- Create a grid from column and row counts with gaps, and draw dividers (solid, dashed or dotted)
- Character count per line is computed from the base text size and leading
- Everything is drawn on a preview layer so you can adjust while watching

## Usage

1. Make the target page active
2. Run the script
3. Configure the panels, check the preview, then click OK

## Notes and limitations

- The outer area and the real content region are kept separate: the title area and the column area (plus its gap) are subtracted from the content region.
- Page bounds used by the auto-adjust features are expressed in spread coordinates so spreads work correctly.
- Dashed and dotted strokes are resolved via itemByName("破線 (3 & 2)" / "点線 (1 & 1)").

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/IdLayoutGridBuilder.jsx` |
| Version | v0.2.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-03-13 |
| Last updated | 2026-03-15 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
