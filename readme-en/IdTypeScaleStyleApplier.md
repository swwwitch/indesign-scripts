# Apply a type scale to paragraph styles (body, headings, lists, tables)

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTypeScaleStyleApplier.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdTypeScaleStyleApplier.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTypeScaleStyleApplier.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Builds a type scale from a base size and ratio, then applies it to the body, heading, list and table paragraph styles.

## Features

- Lists default to 100% and tables to 94% of the body size
- Set leading (%) per body and heading, the kerning method, and the size rounding step
- Space before and after is derived from fixed percentages and can be overridden per row in the preview (size likewise)
- Fonts can be shared, specified separately, or left unchanged; headings can reference the body font and change only the weight
- Font information is not loaded at startup, only when "Include fonts & styles" is pressed (the list is cached on disk)

## Usage

1. Open the target document
2. Run the script
3. Set the base size and ratio, adjust rows in the preview, then click OK

## Notes and limitations

- sameParaStyleSpacing follows per-style rules: 0 for ul-li, the same value as spaceBefore for p and ol-li, and unchanged otherwise.
- Forcing the justification is off by default (`ENABLE_JUSTIFICATION`).
- The whole run is a single undo step.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/style/IdTypeScaleStyleApplier.jsx` |
| Version | v1.6.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-05-05 |
| Last updated | 2026-06-30 |
| Article | https://note.com/dtp_tranist/n/n4f9b0666db66 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
