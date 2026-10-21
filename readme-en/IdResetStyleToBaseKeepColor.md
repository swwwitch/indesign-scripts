# Remove non-color overrides from derived styles

[![Direct](https://img.shields.io/badge/Direct%20Link-IdResetStyleToBaseKeepColor.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdResetStyleToBaseKeepColor.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdResetStyleToBaseKeepColor.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Removes every override except colors from paragraph and character styles that are based on another style, so they follow their base style again. Character colors, paragraph shading, and the colors of underlines, rules and borders are kept.

### Features

- Works on paragraph styles and character styles, including those inside style groups
- Choose which colors to keep by category: character color (fill, stroke, tint, weight, overprint), paragraph shading (on/off, color, tint, overprint), underline and strikethrough colors, paragraph rule and border colors, other colors (kenten, ruby, etc.)
- Styles based directly on [Basic Paragraph] are skipped by default (optional)
- The dialog box shows how many styles and properties will be reset
- A single undo restores everything

### How to use

1. Open a document
2. Run the script, choose the targets and the colors to keep, then click OK

### Notes

- Paragraph styles are reset by assigning the base style's value, which makes the property inherit again (verified). Character styles are set back to unset
- Next Style and export tags are not formatting and are left as they are
- Nested styles, GREP styles and nested line styles are not touched
- Non-color measurements of paragraph shading, such as offsets and width, go back to the base

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/style/IdResetStyleToBaseKeepColor.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-10-21 |
| Last updated | 2026-10-21 |

### Update History

#### v1.0.0 (2026-10-21)

- First release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
