# Split Heading Styles by Part to Format the TOC per Part

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSplitHeadingStylesByPart.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdSplitHeadingStylesByPart.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSplitHeadingStylesByPart.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

InDesign builds a table of contents per paragraph style, so every "h1" gets the same TOC format whichever part it is in. This script assigns the headings listed in the TOC to per-part variant styles (such as "h1.blue") and adds a TOC format for each variant (such as "toc-h1.blue") to the TOC style. After that, a plain Update Table of Contents formats each part differently.

To reformat an existing table of contents without touching the body styles, use [IdApplyTocStylesByPart](IdApplyTocStylesByPart.md).

### Features

- **Parts**: each paragraph in the chosen style (by default the top-level style of the TOC style, such as "h0") starts a part that runs up to the next one, in page order
- **Variant mapping**: the body style "h1.blue" maps to the TOC format of "h1" (such as "toc-h1") plus the same suffix: "toc-h1.blue"
- **Variant per part**: choose a suffix (blue, pink, ...) for each part. The default is the variant used most often within the part (such as "h2.blue"). The suffixes offered are collected from the TOC format styles that have a ".suffix" variant
- **Headings to split**: choose which of the headings in the TOC style to reassign
- Missing body variants are created based on the original style, so the text looks the same. Missing TOC format variants are created based on the TOC format
- Adds an entry for each variant to the TOC style, copying the level, page number position, separator and so on from the original entry
- **Update the table of contents**: rebuilds the existing TOC with the selected TOC style
- **Remove forced line breaks in the TOC**: removes the forced line breaks copied from the headings, which remain even when the TOC style's Remove Forced Line Break is on
- The whole run can be undone with a single Undo

### Usage

1. Create per-part variants of the TOC format styles (for example "toc-h0.blue" based on "toc-h0"). One is enough; variants for the other levels are created automatically
2. Run the script and choose the TOC style and the part style
3. Choose a suffix for each part and click OK

### Notes

- Headings before the first part are left unchanged. Setting a part to None returns its headings to the original styles.
- Paragraphs in variants that are not among the suffixes offered (such as "h0.lead") are left unchanged.
- Styles inside style groups are matched by name only. If several groups contain the same name, the first one found is used.
- Paragraphs on parent pages, on the pasteboard, in overset text or in table cells are not processed.
- The "." separator can be changed with `VARIANT_SEPARATOR` at the top of the script.

### Script info

| Item | Details |
| --- | --- |
| File | `jsx/style/IdSplitHeadingStylesByPart.jsx` |
| Version | v1.0.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-10-30 |
| Last updated | 2026-11-05 |

### Update History

- v1.0.1 (2026-11-05) Moved the version from the title bar to the bottom left of the dialog, small and dimmed (shared part added)
- v1.0.0 (2026-10-30): Initial release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
