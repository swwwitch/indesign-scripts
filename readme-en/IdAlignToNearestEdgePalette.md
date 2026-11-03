# Palette to align objects to the nearest guide, margin, page edge, or bleed in a direction

[![Direct](https://img.shields.io/badge/Direct%20Link-IdAlignToNearestEdgePalette.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/frame/IdAlignToNearestEdgePalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAlignToNearestEdgePalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Moves the selected objects in the direction chosen on the palette and aligns their edge or center to the nearest guide, margin guide, page edge, or bleed edge in that direction.

### Features

- Pick how to align with a 3×3 grid of buttons and two more below it (each icon shows the target line and the object)
  - Top, bottom, left, right: aligns the edge facing the direction (the left edge for left) to the nearest line in that direction
  - Top left, top right, bottom left, bottom right: aligns two edges to the nearest line in each direction
  - Center (cross): aligns the center to the nearest line on each axis
  - Horizontal center, vertical center: aligns the center on one axis only
- The target is the nearest line among guides (spread, page, and parent page), margin guides, page edges, and bleed edges
- The edge buttons move on to the next line when pressed again (for example margin → page edge → bleed). The center buttons do nothing when the center is already on a line
- Each object moves to its own nearest line
- With "Move selection as a unit", everything moves by the same amount, based on the bounds of the whole selection, keeping relative positions
- While the palette is focused, the arrow keys act like the top, bottom, left, and right buttons. Esc closes the palette
- Each action is a single undo step
- "Move selection as a unit" and the palette position are remembered until InDesign quits. Japanese / English UI is chosen automatically

### Usage

1. Run the script to open the palette
2. Select the objects to move (when text is selected with the Type tool, its text frame is used)
3. Click an alignment button

### Notes and limitations

- Edges are measured on the visible bounds, including the stroke.
- Objects with no target line stay where they are (a diagonal button moves only along the axes that have a line). A warning appears only when nothing moved.
- Locked objects, objects on locked layers, and anchored objects are skipped.
- Guides on hidden layers are ignored. Only guides from the parent page applied directly to the page are used (not from the parent page it is based on).
- Column guides are not used.
- Running the script while the palette is open closes the old palette and opens a new one.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/frame/IdAlignToNearestEdgePalette.jsx` |
| Version | v1.1.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-03 |
| Last updated | 2026-11-03 |

### Update History

- v1.1.0 (2026-11-03) Buttons redrawn as a 3×3 grid plus two, with the same icons as the Illustrator script AiAlignToArtboardPalette. Added the diagonals (two edges), center (cross), horizontal center and vertical center. "Move selection as a unit" now uses the bounds of the whole selection
- v1.0.0 (2026-11-03) First release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
