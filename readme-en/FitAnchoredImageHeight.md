# FitAnchoredImageHeight

Scales inline anchored graphic frames so their height matches the surrounding text size, keeping the aspect ratio.

## Features

- Finds paragraphs containing anchored objects with a GREP search for `~a`
- Uses the largest non-image point size in the paragraph as the target height
- `resize()` scales the placed image along with the frame
- The scope is document, story or selection (set at the top of the script)

## Usage

1. Open the target document (select text first if you target a story or selection)
2. Run the script

## Notes and limitations

- Paragraphs that contain only the image are skipped.
- Units are switched to millimetres during the run and always restored afterwards.
- Change the scope via `SEARCH_SCOPE` (default `"document"`).

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/frame/FitAnchoredImageHeight.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-06-01 |
| Last updated | 2026-06-01 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
