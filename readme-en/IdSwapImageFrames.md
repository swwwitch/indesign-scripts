# Rotate images between the selected frames

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSwapImageFrames.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/frame/IdSwapImageFrames.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSwapImageFrames.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Rotates the selected image frames, moving either the linked images alone or the frames together with their positions.

## Features

- Move the images only removes each existing graphic and re-places the linked images in rotating order
- Move the frames keeps every frame's content and rotates only the positions
- Choose the fit after placing (Fill Frame Proportionally or Fit Content Proportionally) and the alignment reference (top left or centre)
- The order follows the visible layout: top to bottom, then left to right. Each item moves to the next frame, and the last one wraps around to the first
- Every option carries a tooltip describing what it does
- The whole run is a single undo step (Cmd+Z)

## Usage

1. Select two or more frames containing placed images
2. Run the script
3. Choose the swap mode and options, then click OK

## Notes and limitations

- A frame is processed only once, even when it is selected twice or selected together with the graphic inside it.
- Move the images only assumes each frame holds a single primary graphic. Frames whose graphic has no link (embedded images, for example) cannot be processed.
- Move the frames never resizes a frame. With frames of different sizes the result depends on the anchor you pick (top left or centre).
- If removal, placement or fitting fails, the run stops at that point and an error is shown. Whatever has already been swapped can be reverted with undo.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/frame/IdSwapImageFrames.jsx` |
| Version | v1.0.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-03-28 |
| Last updated | 2026-09-09 |
| Article | https://note.com/dtp_tranist/n/n6dee03ae96e2 |

## Change log

### v1.0.1 (2026-09-09)

- Reworded the dialog so it reads as a rotation rather than a two-way swap, and aligned the fit options with the standard InDesign terms
- Error messages now state that the run stopped partway and can be reverted with undo
- Added tooltips to every radio button and panel
- Simplified frame detection when a placed graphic itself is selected: any child of an image frame now resolves to that frame, regardless of its type
- Added the link to the introductory article
- Internal cleanup (consistent naming, shared rotation helper, tidier error handling)

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
