# Delete unused styles and swatches

[![Direct](https://img.shields.io/badge/Direct%20Link-IdDeleteUnused.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/misc/IdDeleteUnused.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdDeleteUnused.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Deletes unused styles (paragraph, character, object, table, cell), parent pages, empty pages, swatches, and composite fonts in the active document. Before deleting, the candidates are listed so you can uncheck anything you want to keep.

## Features

- Choose what to delete with checkboxes grouped into Styles, Pages, and Other panels (only swatches are checked on first run)
  - Styles: paragraph / character / object / table / cell styles
  - Pages: parent pages / empty pages
  - Other: swatches / composite fonts
- Before deleting, candidates are listed with kind and name. Unchecked items are kept
- Styles referenced by other styles or document settings are kept as "in use", not only those applied to text or tables
- Styles, parent pages, and composite fonts that become unused once child styles, derived parent pages, or empty pages are deleted are listed as well
- Colors used in gradient stops or as the base of tints are kept, since deleting them would change the gradient or tint
- "Delete empty style groups" option also deletes empty style groups, including those emptied by this deletion
- The dialog settings are remembered and used as the defaults next time
- Shows how many items were deleted per kind, plus the total

## Usage

1. Open a document and run the script
2. Choose what to delete and options, then click [OK]
3. In the confirmation list, uncheck anything you want to keep, then click [Delete]

### Mouse

| Action | Result |
| --- | --- |
| `Option` + click a checkbox under Items to Delete | Check only that item |
| `Option` + click again while only that item is checked | Check all items |
| Double-click a row in the confirmation list | Toggle its check |
| [Check All] / [Uncheck All] in the confirmation list | Toggle all rows |

## What counts as "in use"

| Kind | Counted as in use |
| --- | --- |
| Paragraph styles | Applied to text; Based On / Next Style of another paragraph style; object styles; cell styles; TOC styles (title, entries, entry format); footnote options; text defaults |
| Character styles | Applied to text; Based On of another character style; nested / GREP / line styles; bullets and numbering; drop caps; TOC styles (page number, separator); footnote options; text defaults |
| Object styles | Applied to a page item; Based On of another object style; defaults for new objects |
| Table styles | Applied to a table; Based On of another table style |
| Cell styles | Applied to a cell; Based On of another cell style; table style regions (header, footer, body, left and right columns) |
| Swatches | Not in InDesign's unused swatches; gradient stops; base colors of tints |
| Parent pages | Applied to a page; basis of another parent page |
| Empty pages | Pages with any object on them (parent page objects do not count) |
| Composite fonts | Applied to text; font of a paragraph or character style; text default font |

## Notes and limitations

- Default styles in [ ] such as [No Paragraph Style], [Basic Paragraph], and [None] are never deleted.
- Unnamed colors are not listed. If a default swatch that cannot be deleted (such as [Paper]) is listed, it is kept.
- A checked candidate that is still referenced by an unchecked item is not deleted, and is counted as "kept" in the result.
- Swatches that become unused only after styles are deleted are not in the confirmation list. Run the script again to catch them.
- Character styles referenced only by formatting applied directly to paragraphs (e.g. local nested styles) are not detected as in use.
- Table and cell styles are checked on tables directly in stories. Tables nested inside cells are not examined.
- Text search covers hidden and locked layers, locked stories, master pages, and footnotes. The search scope settings are restored afterwards, but the Find/Change query is cleared.
- Only the active document is processed.
- An empty page is not deleted if it would be the last page of the document.
- Composite fonts are matched by comparing their names with the family part of applied font names.
- Everything can be restored with a single undo.
- Dialog settings are saved to `IdDeleteUnused-prefs.txt` in the user data folder.

## Script info

| Item | Details |
| --- | --- |
| File | `jsx/misc/IdDeleteUnused.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-09-24 |
| Last updated | 2026-09-24 |
| Article | https://note.com/dtp_tranist/n/n879f09b72808 |

## Update history

- v1.0.0 (2026-09-24): Initial release

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
