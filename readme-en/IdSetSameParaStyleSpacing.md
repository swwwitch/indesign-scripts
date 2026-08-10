# IdSetSameParaStyleSpacing

Sets a paragraph style's "Spacing (Same Style)" at the style-definition level.

## Features

- Collects paragraph styles recursively, including those inside style groups
- Spacing can be Ignore, 0, or a specific value
- The value follows the document display units and supports arrow-key stepping (Shift snaps to multiples of 10, Option steps by 0.1)
- With text selected, the applied paragraph style is preselected and its current value shown

## Usage

1. Open the target document (place the cursor in the target paragraph for a convenient default)
2. Run the script
3. Choose the paragraph style and spacing, then click OK

## Notes and limitations

- The change is written to the style definition, so it affects every paragraph using that style.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/style/IdSetSameParaStyleSpacing.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-06-30 |
| Last updated | 2026-06-30 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
