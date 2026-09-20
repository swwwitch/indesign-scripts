# Link text frames into one story

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTextFrameLinker.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdTextFrameLinker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTextFrameLinker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Links the selected text frames, in selection order, so that they share a single story, then optionally fits the height of the first frame and removes the frames left empty.

## Features

- Links two or more selected text frames through `nextTextFrame`, in selection order
- "Fit the height of the first text frame": grows the frame vertically until the text fits, keeping its width
- "Delete the emptied text frame": removes only the frames left empty after linking, keeping the first one
- Leaves frame size and position untouched when both options are off
- Reports when fewer than two items, or anything other than text frames, are selected
- Reports and stops when linking would create a circular thread

## Usage

1. Select two or more text frames, in the order they should be linked
2. Run the script
3. Choose what to do after linking with the checkboxes, then click OK

## Notes and limitations

- The order follows the selection order: the frame selected first becomes the head of the thread. A marquee selection returns the frames in stacking order, so click them one by one.
- Text already in the later frames is not lost: the stories are merged in selection order.
- If a selected frame is already threaded, the existing thread is not broken; the frames are re-spliced into it.
- Only empty frames after the first are deleted; the first frame and any frame that still holds text are kept.
- The checkbox defaults live in the "Default settings" block at the top of the script (`DEFAULT_FIT_FIRST_HEIGHT` / `DEFAULT_DELETE_EMPTY_FRAME`).

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/text/IdTextFrameLinker.jsx` |
| Version | v1.0.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2023-12-26 |
| Last updated | 2026-09-20 |
| Article | https://note.com/dtp_tranist/n/n04ceaf4955a0 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
