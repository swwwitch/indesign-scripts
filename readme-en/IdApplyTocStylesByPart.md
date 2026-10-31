# Reassign TOC Entries to Per-Part Variant Styles

[![Direct](https://img.shields.io/badge/Direct%20Link-IdApplyTocStylesByPart.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdApplyTocStylesByPart.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdApplyTocStylesByPart.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Run after updating the table of contents to reassign its entries to per-part variant styles (such as "toc-h1.blue"). The body styles are left unchanged, so run it each time the TOC is updated.

To split the body headings into variant styles so that a plain Update Table of Contents is enough, use [IdSplitHeadingStylesByPart](IdSplitHeadingStylesByPart.md).

### Features

- **Parts**: each paragraph in the chosen body style (by default the top-level style of the TOC style, such as "h0") starts a part that runs up to the next one
- **Matching the TOC**: each part heading in the TOC (such as a "toc-h0" paragraph) is matched with a body part by its text, and the entries up to the next part heading get that part's variant. Headings that do not match are paired by order and listed in the result
- **Variant per part**: choose a suffix (blue, pink, ...) for each part. The default is the variant used most often within the part (such as "h2.blue")
- A "toc-h1" paragraph becomes "toc-h1.blue" with the part's suffix. If that style does not exist, the base format is kept and the style is listed in the result
- **Remove forced line breaks in the TOC**: removes the forced line breaks copied from the headings, which remain even when the TOC style's Remove Forced Line Break is on
- The whole run can be undone with a single Undo

### Usage

1. Create per-part variants of the TOC format styles (for example "toc-h1.blue" based on "toc-h1")
2. Update the table of contents
3. Run the script, check the suffix for each part and click OK

### Notes

- The TOC is found as the story that uses the TOC style's format styles.
- Entries before the first part heading (such as "Introduction") get the base format.
- Styles inside style groups are matched by name only.
- The "." separator can be changed with `VARIANT_SEPARATOR` at the top of the script.

### Script info

| Item | Details |
| --- | --- |
| File | `jsx/style/IdApplyTocStylesByPart.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-10-30 |
| Last updated | 2026-10-30 |

### Update History

- v1.0.0 (2026-10-30): Initial release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
