# Register paragraph and character styles in bulk

[![Direct](https://img.shields.io/badge/Direct%20Link-IdStyleSetup.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdStyleSetup.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdStyleSetup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Registers paragraph and character styles, their groups, inheritance and GREP styles in a single pass. Fonts can be changed in one place through `base-font`, `base-heading` or `base-text`, and style names can follow either HTML or Word.

### Features

- A dialog at launch picks the style-name scheme (HTML or Word)
- Runs in four stages: create styles and groups, apply attributes, set GREP styles, and reorder
- basedOn relationships are established before attributes are applied, and shared typesetting settings live on the base styles
- Fonts change in one place per family: `base-font` reaches everything, `base-heading` the headings and TOC title, `base-text` body text, lists, table cells and TOC entries
- GREP styles are assigned by role: Latin text on headings, p and lists; line-end no-break on p and lists; anchors on p, lists and table cells; the label only on ul-li
- Table cells build on `base-table` (92% horizontal and vertical scale), split into body cells under `td-left` (td-justify / td-justify-all / td-center / td-right / td-left ul-li for bullets) and header cells under `th-left` (th-center / th-center-W with Paper-colored text)
- Each GREP expression is prefixed with a short comment such as `(?#欧文)`
- A progress bar is shown while running

### Style structure

No style sets a font. Changing `base-font`, `base-heading` or `base-text` updates every style beneath it.

```
[Basic Paragraph]
└─ base-font               parent of every paragraph style; change the font here to update everything
   ├─ base-heading         parent of headings. Metrics, align left, keep with next 2 lines, keep all lines together, no hyphenation
   │  │                    GREP: Latin
   │  ├─ h1–h6             next style: p
   │  └─ toc-title         TOC title
   ├─ base-text            parent of body text. Japanese mojikumi kerning, justify with last line left, keep all lines together, no hyphenation
   │  ├─ p                 all keep options off
   │  │  │                 GREP: Latin, line-end no-break, anchor
   │  │  └─ p.table        for tables (inherits p's GREP)
   │  ├─ p.caption         keep with previous, next style: p
   │  ├─ p.code            no language, ligatures off, align left
   │  ├─ p.img             align center
   │  ├─ ul-li             bullets (li-bullet), keep with previous, space between same-style paragraphs 0, tab stop at the font size
   │  │                    GREP: Latin, line-end no-break, anchor, label
   │  ├─ ol-li             numbering (li-num)
   │  │                    GREP: Latin, line-end no-break, anchor
   │  ├─ base-table        shared table settings. 92% horizontal and vertical scale
   │  │  │                 GREP: anchor
   │  │  ├─ td-left        parent of body cells. Align left
   │  │  │  ├─ td-justify      justify, last line left
   │  │  │  ├─ td-justify-all  justify all lines
   │  │  │  ├─ td-center       align center
   │  │  │  ├─ td-right        align right
   │  │  │  └─ td-left ul-li   bullets inside cells (li-bullet), keep with previous, space between same-style paragraphs 0, tab stop at the font size
   │  │  └─ th-left        parent of header cells. Align left
   │  │     ├─ th-center       align center
   │  │     └─ th-center-W     align center, Paper text
   │  └─ base-toc          parent of TOC entries. Align left, keep with next 2 lines, keep all lines together, no hyphenation
   │     ├─ toc-h1
   │     ├─ toc-h2
   │     └─ toc-h3         not kept with next
   ├─ page-number          folio. Align away from spine
   ├─ running-head         running head
   └─ thumb-index          thumb index
```

### Style names

| HTML | Word |
| --- | --- |
| `h1`–`h6` | `Heading 1`–`Heading 6` |
| `p` | `Normal` |
| `ul-li` | `List Paragraph` |
| `ol-li` | `List Number` |
| `strong-bold` | `Strong` |
| `em-italic` | `Emphasis` |

Choosing Word means a Word file placed with "Preserve Styles and Formatting from Text and Tables" needs no style mapping. Styles missing from the table above (`p.caption`, `p.code`, `link`, and grouped ones such as `base-font`, `td-left` and `toc-h1`) keep the same name under either scheme.

### Usage

1. Open the target document
2. Run the script
3. Pick the style-name scheme in the dialog and click OK

### Notes and limitations

- Existing same-named styles are left untouched by default. Set `OVERWRITE_EXISTING_STYLES` to true to re-apply every attribute and rebuild the GREP rules. Running the Word scheme on a document that already holds a placed Word manuscript therefore treats `Heading 1` and friends as existing styles. Running on a document set up by v1.4.3 or earlier creates the base styles again under their new names and duplicates the GREP styles, so enable replacement.
- Word puts both bullet and numbered lists into "List Paragraph", but InDesign cannot hold two styles of the same name, so numbered lists are mapped to `List Number` instead. Edit `WORD_STYLE_NAMES` at the top of the script to change the mapping.
- Attributes the script does not set (font, weight, and so on) are never reset. The only colour it sets is the Paper text on `th-center-W`.
- The whole run is a single undo step.
- Kerning method names are localized, so the candidates are tried in order. On a locale where none of them match, the setting is left as it is.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/style/IdStyleSetup.jsx` |
| Version | v1.5.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-05-03 |
| Last updated | 2026-10-01 |
| Article | https://note.com/dtp_tranist/n/nfe87ec253780 |

### Update History

#### v1.5.1 (2026-10-01)

- Added `td-left ul-li` (based on `td-left`) for bulleted lists inside table cells
- Set the tab stop of `ul-li` and `td-left ul-li` to the font size (one em)

#### v1.5.0 (2026-10-01)

- Reorganized inheritance: `base-font` is now the parent of everything, with headings going through `base-heading` and body text (including p.code, p.img, table cells and TOC entries) through `base-text`. Renamed base styles: `base-regex` → `base-font`, `body-text` → `base-text`, `heading` → `base-heading`
- Moved GREP styles off the base styles: Latin text on headings, p and lists; line-end no-break on p and lists; anchors on p, lists and table cells
- Reorganized table cell styles: `td-left` and `th-left` now sit under `base-table` (92% horizontal and vertical scale), with new `td-justify` (justify, last line left), `td-justify-all` (justify all lines), `td-center`, `td-right`, `th-center` and `th-center-W` (Paper text). Removed `th`, `td` and `th-right`
- Each GREP expression now starts with a short comment such as `(?#欧文)`
- Set the page-number paragraph style to Align Away from Spine

#### v1.4.3 (2026-10-01)

- Unified the dialog opacity to 0.98 and aligned the progress palette with the other scripts

#### v1.4.2 (2026-10-01)

- A button row with only right-side buttons is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones

#### v1.4.1 (2026-09-30)

- Unified the button row, spacing and localization code with the shared parts. Buttons are now centered

#### v1.4.0 (2026-09-19)

- Added a dialog at launch to choose the style-name scheme: HTML or Word
- The Word scheme registers `Heading 1`–`Heading 6`, `Normal`, `List Paragraph`, `List Number`, `Strong` and `Emphasis`, so placing a Word manuscript needs no style mapping
- The mapping lives in `WORD_STYLE_NAMES` at the top of the script

#### v1.3.3 (2026-09-01)

- Added the `p.table` paragraph style for tables, based on `p`

#### v1.3.2 (2026-09-01)

- Fixed settings such as turning off the keep options on `p` sometimes having no effect. Attributes were applied before basedOn was assigned, so the parent's (body-text) settings could be inherited again and cancel them out. Inheritance is now established before attributes are applied
- Fixed `heading` and h1–h6 losing their inheritance entirely when `body-text` is missing from the `basestyle` group
- Fixed the kerning method assignment failing on non-Japanese InDesign, which rolled back the whole script. The localized names (`和文等幅`, `Japanese Mojikumi`, and so on) are now tried in order
- Removed unused UI helpers and brought the JSDoc return types and comments in line with the implementation

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
