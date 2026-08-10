# IdFontConverter

Switches font variants (character set, P, UD, N, NT, weight) in bulk across a selection, story, spread or the whole document.

## Features

- Target the selection, the story, the whole document, or the active spread
- Toggle character set (Std / Pro / Pr5 / Pr6), N and NT, and UD and P between keep / off / on
- Max and MaxN presets move fonts to the richest available character set
- Paragraph styles, character styles, composite fonts, and locked or hidden objects can all be included
- A preview lists every change (old to new) using Japanese font names and warns about fonts that are not installed

## Usage

1. Open the target document (select text or frames if you target the selection)
2. Run the script
3. Choose the target and conversion options, then click Run
4. Review the preview and confirm only the changes you want

## Notes and limitations

- AXIS (Type Project) fonts use a different character-set scheme and get dedicated handling: width and Joyo are preserved, and only N and Std/Pro are switched.
- When no matching weight exists, the nearest weight is substituted.
- Changes are applied per textStyleRange and the whole run is a single undo step.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/font/IdFontConverter.jsx` |
| Version | v1.1.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-06-17 |
| Last updated | 2026-06-30 |
| Article | https://note.com/dtp_tranist/n/n261c771b4b41 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
