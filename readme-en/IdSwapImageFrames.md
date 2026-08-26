# IdSwapImageFrames

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSwapImageFrames.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/IdSwapImageFrames.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSwapImageFrames.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Rotates the contents or positions of the selected image frames, either by swapping linked images or by swapping the frames themselves.

## Features

- Swap linked images removes each existing graphic and re-places the linked images in rotating order
- Swap frames keeps every frame's content and rotates only the positions
- Choose the fit after placing (fill or fit proportionally) and the position anchor (top left or centre)
- The order follows the visible layout: top to bottom, then left to right

## Usage

1. Select two or more frames containing placed images
2. Run the script
3. Choose the swap mode and options, then click OK

## Notes and limitations

- A frame selected more than once is processed only once.
- Swap linked images assumes each frame holds a single primary graphic.
- If removal, placement or fitting fails the run is cancelled and an error is shown.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/IdSwapImageFrames.jsx` |
| Version | v1.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-03-28 |
| Last updated | 2026-03-28 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
