# Set paragraph rules from a palette with a live preview

[![Direct](https://img.shields.io/badge/Direct%20Link-IdParagraphRulesPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdParagraphRulesPalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdParagraphRulesPalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Recreates InDesign's Paragraph Rules dialog box as a palette and sets the rule above and rule below on the paragraphs of the selected text, text frames or table cells with a live preview.

### Features

- Rule Above and Rule Below are shown side by side on one screen (no switching as in the built-in dialog box)
- Weight, Type, Color, Tint, Overprint Stroke, Gap Color, Gap Tint, Overprint Gap, Width (Column / Text), Offset, indents and Keep In Frame (rule above only)
- Only the items you change after opening are written, so other settings and values that differ between paragraphs stay as they are
- With Preview on, the settings are applied to the document temporarily
- OK applies everything as a single undo step
- A swatch chip is shown next to each color; (Text Color) shows the paragraph's text color
- Items are dimmed to match the settings (everything when Rule On is off, the gap items for solid rules, tint and overprint for Text Color, None and Paper)
- Indent labels read Left / Right in horizontal text and Top / Bottom in vertical text
- Numeric fields step with the ∧∨ buttons and the ↑↓ keys (Shift snaps to 10s, Option steps by 0.1), and accept arithmetic and other units (typing "1in" in a mm field gives "25.4 mm")
- Lengths follow the document's units (stroke, horizontal and vertical)
- Japanese / English UI

### Usage

1. Select text, text frames or table cells
2. Run the script
3. Set the rule above and rule below in the palette (check the result with Preview)
4. Click OK to apply. Closing with Cancel, Esc or the close box restores the original

### Notes and limitations

- The values shown in the palette are read from the first paragraph of the selection
- Running the script again while the palette is open closes the old palette (reverting its preview) and opens a new one
- If the target text is deleted while the palette is open, the preview or apply fails with an alert

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/text/IdParagraphRulesPalette.jsx` |
| Version | v1.3.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-10-19 |
| Last updated | 2026-10-19 |

### Update History

#### v1.3.0 (2026-10-19)

- First release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
