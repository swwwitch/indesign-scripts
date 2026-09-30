# Delete a marker and apply a paragraph style (simple version)

[![Direct](https://img.shields.io/badge/Direct%20Link-IdRemoveMarkerApplyStyleSimple.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdRemoveMarkerApplyStyleSimple.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdRemoveMarkerApplyStyleSimple.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

Applies a paragraph style to paragraphs that contain the given search text (a marker), then deletes the marker. This is a simple version of [IdRemoveMarkerApplyStyle](IdRemoveMarkerApplyStyle.md) that keeps only three settings: the search text, the paragraph style and the search scope.

### Features

- Applies the chosen paragraph style to paragraphs containing the search text (`###` by default) and deletes the marker
- Spaces, full-width spaces and tabs right after the marker are deleted with it
- Never matches part of a longer run of the same character (`##` does not match `### Heading`)
- The search scope is Story, Document or All Documents. The dialog opens on Story when text is selected, otherwise on Document
- "Matches" shows how many locations the current settings find, and is recounted whenever the search text or scope changes
- Paragraph styles inside style groups can be chosen
- The whole run can be undone in a single step (Cmd+Z)

### Usage

1. Place the cursor in the text that contains the markers (to search the whole document, just have it open)
2. Run the script
3. Set the search text, the style to apply and the search scope, then click OK

### Notes and limitations

- The search text is not anchored to the start of the paragraph, so markers in the middle of a paragraph are deleted too.
- The search is width sensitive and kana sensitive, and does not inherit the previous Find/Change settings.
- Footnotes, master pages, hidden layers, locked layers and locked stories are not searched (change this in the settings at the top of the script).
- With All Documents, documents that lack the chosen paragraph style are skipped and listed when the run finishes.
- To detect line-start markers automatically, apply character styles or search with GREP, use [IdRemoveMarkerApplyStyle](IdRemoveMarkerApplyStyle.md).

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/style/IdRemoveMarkerApplyStyleSimple.jsx` |
| Version | v1.2.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-09-09 |
| Last updated | 2026-10-01 |

### Update History

#### v1.2.0 (2026-10-01)

- Moved OK / Cancel from the right-hand column to the standard button row at the bottom of the dialog (Cancel, then OK)

#### v1.1.0 (2026-09-30)

- OK is now disabled when nothing matches (a search string of only spaces could delete spaces in the text)
- Trailing spaces and tabs in the search string are now ignored
- Shared the search code with IdRemoveMarkerApplyStyle and unified localization and spacing with the shared parts

- v1.0 (2026-09-10) Initial version

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
