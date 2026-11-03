# Palette to align objects to the nearest guide, margin, page edge, or bleed in a direction

[![Direct](https://img.shields.io/badge/Direct%20Link-IdAlignToNearestEdgePalette.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/frame/IdAlignToNearestEdgePalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAlignToNearestEdgePalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Moves the selected objects in the direction chosen on the palette and aligns their edge to the nearest guide, margin guide, page edge, or bleed edge in that direction.

### Features

- Four directions: ←, →, ↑, ↓. The edge facing the direction is aligned (the left edge for ←, the top edge for ↑, and so on)
- The target is the nearest line in that direction among guides (spread, page, and parent page), margin guides, page edges, and bleed edges
- Pressing again moves on to the next line (for example margin → page edge → bleed)
- Each object moves to its own nearest line
- With "Move selection as a unit", everything moves by the same amount, based on the object furthest in the direction, keeping relative positions
- While the palette is focused, the arrow keys work too. Esc closes the palette
- Each move is a single undo step
- "Move selection as a unit" and the palette position are remembered until InDesign quits. Japanese / English UI is chosen automatically

### Usage

1. Run the script to open the palette
2. Select the objects to move (when text is selected with the Type tool, its text frame is used)
3. Click the button for the direction

### Notes and limitations

- Edges are measured on the visible bounds, including the stroke.
- Objects with no line in the direction stay where they are. A warning appears only when nothing moved.
- Locked objects, objects on locked layers, and anchored objects are skipped.
- Guides on hidden layers are ignored. Only guides from the parent page applied directly to the page are used (not from the parent page it is based on).
- Column guides are not used.
- Running the script while the palette is open closes the old palette and opens a new one.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/frame/IdAlignToNearestEdgePalette.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-03 |
| Last updated | 2026-11-03 |

### Update History

- v1.0.0 (2026-11-03) First release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
