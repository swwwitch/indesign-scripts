# Register paragraph and character styles in bulk

[![Direct](https://img.shields.io/badge/Direct%20Link-IdStyleSetup.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdStyleSetup.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdStyleSetup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Registers paragraph and character styles, their groups, inheritance and GREP styles in a single pass. Style names can follow either HTML or Word.

## Features

- A dialog at launch picks the style-name scheme (HTML or Word)
- Runs in four stages: create styles and groups, apply attributes, set GREP styles, and reorder
- basedOn relationships are established before attributes are applied, and shared typesetting settings live on the base styles
- Shared GREP rules (lang-US, no-break, inline-graphic) sit on base-regex and are inherited by child styles
- ul-li carries its own li-label rule, so the three shared rules are also set directly to avoid losing inheritance
- A progress bar is shown while running

## Style names

| HTML | Word |
| --- | --- |
| `h1`–`h6` | `Heading 1`–`Heading 6` |
| `p` | `Normal` |
| `ul-li` | `List Paragraph` |
| `ol-li` | `List Number` |
| `strong-bold` | `Strong` |
| `em-italic` | `Emphasis` |

Choosing Word means a Word file placed with "Preserve Styles and Formatting from Text and Tables" needs no style mapping. Styles missing from the table above (`p.caption`, `p.code`, `link`, and grouped ones such as `base-regex`, `th` and `toc-h1`) keep the same name under either scheme.

## Usage

1. Open the target document
2. Run the script
3. Pick the style-name scheme in the dialog and click OK

## Notes and limitations

- Existing same-named styles are left untouched by default. Set `OVERWRITE_EXISTING_STYLES` to true to re-apply every attribute and rebuild the GREP rules. Running the Word scheme on a document that already holds a placed Word manuscript therefore treats `Heading 1` and friends as existing styles.
- Word puts both bullet and numbered lists into "List Paragraph", but InDesign cannot hold two styles of the same name, so numbered lists are mapped to `List Number` instead. Edit `WORD_STYLE_NAMES` at the top of the script to change the mapping.
- Attributes the script does not set (font, weight, colour, and so on) are never reset.
- The whole run is a single undo step.
- Kerning method names are localized, so the candidates are tried in order. On a locale where none of them match, the setting is left as it is.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/style/IdStyleSetup.jsx` |
| Version | v1.4.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-05-03 |
| Last updated | 2026-09-19 |
| Article | https://note.com/dtp_tranist/n/nfe87ec253780 |

## Update history

### v1.4.0 (2026-09-19)

- Added a dialog at launch to choose the style-name scheme: HTML or Word
- The Word scheme registers `Heading 1`–`Heading 6`, `Normal`, `List Paragraph`, `List Number`, `Strong` and `Emphasis`, so placing a Word manuscript needs no style mapping
- The mapping lives in `WORD_STYLE_NAMES` at the top of the script

### v1.3.3 (2026-09-01)

- Added the `p.table` paragraph style for tables, based on `p`

### v1.3.2 (2026-09-01)

- Fixed settings such as turning off the keep options on `p` sometimes having no effect. Attributes were applied before basedOn was assigned, so the parent's (body-text) settings could be inherited again and cancel them out. Inheritance is now established before attributes are applied
- Fixed `heading` and h1–h6 losing their inheritance entirely when `body-text` is missing from the `basestyle` group
- Fixed the kerning method assignment failing on non-Japanese InDesign, which rolled back the whole script. The localized names (`和文等幅`, `Japanese Mojikumi`, and so on) are now tried in order
- Removed unused UI helpers and brought the JSDoc return types and comments in line with the implementation

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
