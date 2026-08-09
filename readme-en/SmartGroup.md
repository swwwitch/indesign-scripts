# SmartGroup

Groups the selected objects by horizontal (row) or vertical (column) proximity, with a live red preview of each group while the dialog is open.

## Features

- Switch the direction between horizontal and vertical with radio buttons
- Adjust the tolerance from 0 to 50 with a slider
- Group extents are drawn live as red rectangles on a non-printing layer
- The preview is removed automatically when the dialog closes

## Usage

1. Select the objects you want to group
2. Run the script
3. Adjust the direction and tolerance, check the preview, then click OK

## Notes and limitations

- The preview lives on a non-printing layer named "SmartGroup Preview" and is deleted on exit.
- A group needs at least two objects.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/group/SmartGroup.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-04-11 |
| Last updated | 2026-04-17 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
