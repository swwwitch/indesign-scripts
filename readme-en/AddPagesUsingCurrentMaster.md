# AddPagesUsingCurrentMaster

Inserts a given number of pages right after the current page, carrying over the parent (master) page applied to it.

## Features

- Shows the current page name and the applied parent page in the dialog
- Prompts for the number of pages to insert (default 2)
- Applies the current parent page to every inserted page
- Adds pages at document level so they reflow into the correct spreads

## Usage

1. Make the page you want to insert after the active page
2. Run the script
3. Enter the number of pages and click OK

## Notes and limitations

- An active InDesign document is required.
- Unlike the built-in Insert Pages dialog, this keeps the parent page of the selected page.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/page/AddPagesUsingCurrentMaster.jsx` |
| Version | v1.2.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2025-06-26 |
| Last updated | 2026-06-30 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
