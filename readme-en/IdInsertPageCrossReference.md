# Insert a paragraph's page number as a "p.<number>" cross-reference

[![Direct](https://img.shields.io/badge/Direct%20Link-IdInsertPageCrossReference.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdInsertPageCrossReference.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdInsertPageCrossReference.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

Inserts the page number of the third paragraph of the first story at the end of the second story, as a cross-reference such as "p.12".

A page number typed as text has to be fixed by hand when pages are added or removed. As a cross-reference, the number follows page changes and also works as a link to the target.

### Features

- Creates a paragraph destination at the target paragraph and inserts a cross-reference at the insertion point
- Creates the cross-reference format "p.ページ番号" (the text "p." plus the page number), or reuses a format with that name
- A single Undo reverts everything

### How to use

1. Open a document and run the script
2. "p.<number>" is inserted at the end of the second story and added to the Cross-References panel

### Notes

- The target (third paragraph of the first story) and the insertion point (end of the second story) are fixed. Story order does not follow page position or creation order, so a different paragraph may be used in some documents.
- If there are fewer than two stories, or the first story has fewer than three paragraphs, the script shows a message and stops.
- When the target is in overset text or on the pasteboard, InDesign's cross-reference handling applies.
- Not yet tested in InDesign.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/document/IdInsertPageCrossReference.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-06 |
| Last updated | 2026-11-06 |

### Update History

- v1.0.0 (2026-11-06): Initial release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
