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
- Fonts are left untouched while "Include fonts & styles" is off; turning it on offers "same for body and headings" or "specify separately"
- The weight can be changed on its own, without picking a font family; the options come from the font each paragraph style already uses
- With "Size only" on, everything but the point size (font, leading, spacing, kerning) keeps its original value
- The font list is loaded at startup and cached on disk

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
| Version | v1.6.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-05-05 |
| Last updated | 2026-09-20 |
| Article | https://note.com/dtp_tranist/n/n4f9b0666db66 |

## Changelog

### v1.6.1 (2026-09-20)

- "Size only" now defaults to off. With the previous default, leading and paragraph spacing were never applied unless the box was cleared first
- "Include fonts & styles" became a checkbox and moved into the Font Assignment panel. Clearing it means what "Do not change fonts" used to mean, so that radio button was removed
- Fixed the weight not being changeable on its own when no font family is selected; the options now come from the font each paragraph style already uses
- Fixed space before and after being read in the document's ruler units, which made them about 2.83x too large in millimeter documents
- Fixed the default base size being off when the text size unit is not points (Q, for example); the unit now comes from the document rather than the application preferences

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
