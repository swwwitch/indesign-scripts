# Apply rotated Roman characters, Tate-Chu-Yoko and number rewriting to vertical text

[![Direct](https://img.shields.io/badge/Direct%20Link-IdVerticalTextFormatter.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdVerticalTextFormatter.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdVerticalTextFormatter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

Turns the selected text vertical and applies Rotate Roman Characters, Tate-Chu-Yoko, number rewriting (full-width digits, kanji numerals, Japanese era years), kinsoku and mojikumi in one go. This is the InDesign version of the Illustrator script [VerticalTextFormatter](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/VerticalTextFormatter.md).

### Features

- Simple, Detailed and Wizard modes (switch at the top left of the dialog box) and presets (7 built-in, save and delete your own)
- Rotation and Tate-Chu-Yoko: rotate Roman characters (skip text of N or more characters) / TCY numbers (2, 3, 4 digits; skip numbers next to separators) / TCY 2–3 capitals / TCY "!!" and "!?" / fixed-width forms by length (half, third and quarter width)
- Character Conversion: separators (、 / remove / keep) / decimal . → ・ / range - ~ → 〜 / full-width alnum to half-width / symbols to full-width / quotes to 〝〟 / ! ? to full-width / single letters to full-width / capitals of up to 3 letters to full-width (PDF → ＰＤＦ) / short words to full-width
- Numbers: single digits to full-width / numbers of 3+ digits to full-width (2026 → ２０２６) / non-TCY digits to full-width / units to squares / Western years to Japanese eras / kanji numerals to Arabic / convert to kanji numerals (digit by digit, 十百千, 万億兆, 零, 一千) / number hyphens
- Character and Paragraph: kinsoku set and mojikumi set (from the document; Soft kinsoku and line-end punctuation half-width by default) / leading as an auto-leading percentage / proportional metrics for full-width alphanumerics / do not tighten 一 / height to a multiple of the type size / auto-size frames (width only, top-right reference)
- Check Coloring: full-width alphanumerics and ！？ red, half-width blue, fixed-width glyphs green, the rest [Black] (swatches such as Check_FullWidth are added)
- Conversion Check... lists what the current settings change, in a separate dialog box
- Text that is already vertical is first returned to horizontal form, so the script can be run again with other settings (kanji numerals and eras are reverted only in stories last processed with them on)
- The whole run is a single undo step
- Settings are saved and restored next time. Japanese / English UI switches automatically

### How to use

1. Select text frames (or a range of text with the Type tool)
2. Run the script and choose the options in the dialog box
3. Check the list of changes with Conversion Check... and click OK

### Notes

- Selecting a text frame targets the whole story, threaded frames included. The orientation changes per story (also when a range of text is selected).
- Text in table cells and footnotes is not processed. Locked or hidden text frames (or layers) are skipped.
- When a rewrite changes the number of characters (3,000 → 三千 etc.), the rewritten part takes the formatting of its first character.
- At the end, leading is set to Auto, auto kerning to None (0) and justification to justify with last line top.
- Height to a multiple of the type size leaves rotated or sheared frames alone; with auto-sizing on, the height follows the auto-sizing.
- Auto-size frames applies only to text frames that are not threaded.
- The Simple mode and the wizard leave kinsoku, mojikumi and leading unchanged (except the wizard's "Set up paragraphs for vertical text").
- Processing is recorded in a story label (key VerticalTextFormatter).
- Differences from the Illustrator version: rotation, Tate-Chu-Yoko and glyph forms are set through the DOM (no temporary action). Kinsoku and mojikumi are chosen from the document's lists (Illustrator's Soft v2 and Solid do not exist in InDesign, so the defaults are Soft and line-end punctuation half-width).
- Settings are saved in `~/Library/Application Support/indesign-scripts/IdVerticalTextFormatter.json`, presets in `IdVerticalTextFormatterPresets.json` in the same folder.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/text/IdVerticalTextFormatter.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-01 |
| Last updated | 2026-11-01 |

### Update History

- v1.0.0 (2026-11-01) First release (ported from the Illustrator VerticalTextFormatter v1.28.0)
