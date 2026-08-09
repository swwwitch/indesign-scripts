# ParagraphShadingMatchRule

Sets an invisible paragraph rule above on each selected paragraph, sized to match the paragraph shading height, so it acts as a layout spacer.

## Features

- The rule weight is the shading top offset plus the point size
- The rule colour is set to None so nothing is visible on screen or in output
- Keep Rule Above In Frame is enabled automatically

## Usage

1. Select the target text
2. Run the script

## Notes and limitations

- This is not for drawing a visible line: a thick invisible rule is used as a spacer to approximate the shaded area's height.
- A missing shading offset is treated as 0.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/ParagraphShadingMatchRule.jsx` |
| Version | v1.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-04-12 |
| Last updated | 2026-04-12 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
