# Palette to align objects to the nearest guide, margin, page edge, or bleed in a direction

[![Direct](https://img.shields.io/badge/Direct%20Link-IdAlignToNearestEdgePalette.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/frame/IdAlignToNearestEdgePalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAlignToNearestEdgePalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Moves the selected objects in the direction chosen on the palette and aligns their edge to the nearest guide, margin guide, page edge, or bleed edge in that direction. The center buttons align to the center of the page or of the margins.

### Features

- Pick how to align with a 3×3 grid of buttons and two more below it (each icon shows the target line and the object)
  - Top, bottom, left, right: aligns the edge facing the direction (the left edge for left) to the nearest line in that direction
  - Top left, top right, bottom left, bottom right: aligns two edges to the nearest line in each direction
  - Center (cross): aligns the center to the center of the page (or of the margins)
  - Horizontal center, vertical center: aligns the center to the center of the page (or of the margins) on one axis only
- "Center Alignment Basis" chooses what the center buttons align to: Page, or Margins (the area inside the margin guides)
- Edges go to the nearest line among guides (spread, page, and parent page), margin guides, page edges, and bleed edges
- The edge buttons move on to the next line when pressed again (for example margin → page edge → bleed)
- Option (Alt)-click an edge button to treat the object as centered first (at the center of the "Center Alignment Basis" area) and align from there to the nearest line (Option with an arrow key does the same)
- Each object moves to its own nearest line
- With "Move selection as a unit", everything moves by the same amount, based on the bounds of the whole selection, keeping relative positions
- While the palette is focused, the arrow keys act like the top, bottom, left, and right buttons. Esc closes the palette
- Each action is a single undo step
- Run it with new initial values for the variables `SW_ACTION`, `SW_FROM_CENTER`, `SW_MOVE_AS_UNIT`, and `SW_CENTER_BASIS` at the top, and it does that one button without the palette (so each button can get a keyboard shortcut, for example with LayoutKeys)
- "Center Alignment Basis", "Move selection as a unit", and the palette position are remembered until InDesign quits. Japanese / English UI is chosen automatically

### Usage

1. Run the script to open the palette
2. Select the objects to move (when text is selected with the Type tool, its text frame is used)
3. Click an alignment button

### Keyboard shortcuts for each button

With `SW_ACTION` empty (the default), the script opens the palette. Run it with a button name instead, and it aligns like that button without the palette.

| Variable | Value |
| --- | --- |
| `SW_ACTION` | `upLeft` (top left), `up` (top), `upRight` (top right), `left`, `center`, `right`, `downLeft` (bottom left), `down` (bottom), `downRight` (bottom right), `centerH` (horizontal center), `centerV` (vertical center) |
| `SW_FROM_CENTER` | `true` centers the object first, then aligns (like Option-click) |
| `SW_MOVE_AS_UNIT` | Sets "Move selection as a unit" to `true` or `false`; empty follows the palette |
| `SW_CENTER_BASIS` | Sets "Center Alignment Basis" to `page` or `margin`; empty follows the palette |

In LayoutKeys, register this script once and click "Add a Row for Each SW_ACTION" in the Hotkey table of the detail pane to get one row per button, then assign a shortcut to each row (set SW_FROM_CENTER to true on rows that should center first). The file itself is not changed. With `SW_CENTER_BASIS` and `SW_MOVE_AS_UNIT` empty, "Center Alignment Basis" and "Move selection as a unit" follow the palette's settings (until InDesign quits).

### Notes and limitations

- Edges are measured on the visible bounds, including the stroke.
- The center buttons use the page that contains the object's center (one page of a spread), or the area inside its margins. When the center is off the pages, the nearest page is used.
- Objects with no target line stay where they are (a diagonal button moves only along the axes that have a line). A warning appears only when nothing moved.
- Locked objects, objects on locked layers, and anchored objects are skipped.
- Guides on hidden layers are ignored. Only guides from the parent page applied directly to the page are used (not from the parent page it is based on).
- Column guides are not used.
- Running the script while the palette is open closes the old palette and opens a new one.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/frame/IdAlignToNearestEdgePalette.jsx` |
| Version | v1.2.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-03 |
| Last updated | 2026-11-05 |

### Update History

- v1.2.1 (2026-11-05) Moved the version from the title bar to the bottom left of the dialog, dimmed (shared part added)
- v1.2.0 (2026-11-03) The center (cross), horizontal center, and vertical center buttons now always align to the center of the page instead of the nearest line. Added "Center Alignment Basis" (Page / Margins) and Option (Alt)-click to center first, then align. Added the variables `SW_ACTION`, `SW_FROM_CENTER`, `SW_MOVE_AS_UNIT`, and `SW_CENTER_BASIS` to run one button without the palette
- v1.1.0 (2026-11-03) Buttons redrawn as a 3×3 grid plus two, with the same icons as the Illustrator script AiAlignToArtboardPalette. Added the diagonals (two edges), center (cross), horizontal center and vertical center. "Move selection as a unit" now uses the bounds of the whole selection
- v1.0.0 (2026-11-03) First release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
