# Convert Word paragraph style names to HTML element names

[![Direct](https://img.shields.io/badge/Direct%20Link-IdConvertWordStyleNamesToHtml.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdConvertWordStyleNamesToHtml.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdConvertWordStyleNamesToHtml.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

Renames paragraph styles imported from MS Word (Heading 1 / Normal / Quote and the like) to the matching HTML element names (h1 / p / blockquote and so on) in one pass. There is no dialog: the script runs right away and shows a report.

### Features

- `Heading 1` to `Heading 6` become `h1` to `h6` (also without the space, as in `Heading1`)
- Other names are mapped as in the table below
- Case, full-width spaces and repeated whitespace are ignored when matching
- Style groups are processed recursively
- When a style with the target name already exists, the Word style is merged into it (paragraphs using it switch to the existing style, and the Word style is deleted)
- A report lists how many styles were converted, merged, skipped or failed, with details
- The whole run can be undone in a single step (Cmd+Z)

| Word style name | Converted to |
| --- | --- |
| Normal / Default Paragraph Style / Body Text / Body Text 2 / No Spacing | p |
| Heading 1-6 | h1-h6 |
| Title / TOC Heading | h1 |
| Subtitle | h2 |
| Quote / Intense Quote | blockquote |
| List Paragraph / List Bullet / List Number | li |
| Caption | figcaption |
| Header | header |
| Footer | footer |
| Hyperlink | a |
| Emphasis | em |
| Strong | strong |

### Usage

1. Open the document that contains the placed or imported Word file
2. Run the script
3. Check the report

### Notes and limitations

- Only paragraph styles are renamed. Character styles are left as they are.
- Styles whose names start with a bracket, such as `[Basic Paragraph]`, are skipped.
- Names not in the table are skipped as "no match", and names that are already converted are skipped as "no change needed".
- The merge target is searched in the same group first, then at the top level of the document, then in other groups; the first match is used.
- The mapping can be changed in `STYLE_NAME_MAP` at the top of the script.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/style/IdConvertWordStyleNamesToHtml.jsx` |
| Version | v1.0.2 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-09-20 |
| Last updated | 2026-09-30 |

### Update History

#### v1.0.2 (2026-09-30)

- Unified the localization code with the shared part (no change in behavior)

- v1.0.1 (2026-09-20) Initial version

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
