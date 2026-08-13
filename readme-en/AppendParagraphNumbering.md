# AppendParagraphNumbering

Appends sequential numbering when the same text repeats within the same paragraph style, judging duplicates per parent heading.

## Features

- Lists paragraph styles with checkboxes so you can narrow the targets
- Switches the scope between story and document
- Switches between full-width and half-width brackets (Japanese UI only)
- Includes a Remove button that strips existing numbering

## Usage

1. Open the target document
2. Run the script
3. Choose the paragraph styles and scope, then click OK

## Notes and limitations

- Targets the active document.
- Both numbering and removal are a single undo step.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/page/AppendParagraphNumbering.jsx` |
| Version | v1.2.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2025-06-30 |
| Last updated | 2026-08-13 |
| Article | https://note.com/dtp_tranist/n/nc96549bb60f9 |

## Changelog

### v1.2.0 (2026-08-13)

- Fixed paragraphs without a recognized parent heading being dropped. Only the names in `HEADING_LEVEL_MAP` (`h1`–`h6` / `Heading 1`–`Heading 6`) count as headings, so documents using any other heading style found no targets at all
- Fixed selected targets being silently skipped when the document contained empty paragraphs or single-character headings
- Fixed the Cancel button not closing the dialog in the Japanese UI
- Text frames on parent pages are now excluded correctly even when they sit inside a group
- The dialog now closes when Remove runs, instead of leaving a stale list behind
- Added messages for "no document open" and for Story scope with nothing selected
- Sped up the analysis pass

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
