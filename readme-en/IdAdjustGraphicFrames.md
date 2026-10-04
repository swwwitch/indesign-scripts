# Batch-adjust anchored graphic frames

[![Direct](https://img.shields.io/badge/Direct%20Link-IdAdjustGraphicFrames.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/frame/IdAdjustGraphicFrames.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAdjustGraphicFrames.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Collects text-anchored graphic frames and adjusts their width, frame size and image scale in one pass.

### Features

- Target the document, the story, or the selection
- Frame width: leave unchanged, fit to the parent frame, or fit to the page margins (with an "only when wider" toggle)
- Fit frames to content, and match inline images to the surrounding text size
- Round image scale down in 1%, 5% or 10% steps (optionally limited to 72/96/144 ppi images, with an optional re-fit)
- Broken aspect ratios are always corrected against the horizontal scale

### Usage

1. Open the target document (select text or frames if you target a story or the selection)
2. Run the script
3. Set the target, width, frame size and scale options, then click OK

### Notes and limitations

- Only text-anchored graphic frames are processed; independent frames and text frames are skipped.
- True inline images with surrounding text get height matching only: scale rounding and width fitting are skipped.
- The whole run is a single undo step.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/frame/IdAdjustGraphicFrames.jsx` |
| Version | v1.0.2 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-06-02 |
| Last updated | 2026-10-01 |

### Update History
- v1.0.3 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)

#### v1.0.2 (2026-10-01)

- A button row with only right-side buttons is now centered in dialogs up to 200 px wide (inside the margins) and right-aligned in wider ones

#### v1.0.1 (2026-09-30)

- Unified the button row, panel spacing and localization code with the shared parts. Buttons are now centered and panels are tighter

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
