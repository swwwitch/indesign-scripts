# AGENTS.md

This repository contains Adobe InDesign ExtendScript (`.jsx`).

## SCRIPTMETA

Every script carries a [SCRIPTMETA](https://github.com/Yamonov/SCRIPTMETA) block (Local Subset) right before
the header comment. It is generated from the basic info block and the 概要 / Overview text — do not edit it by hand.
Update checking is not used: write no `Meta-URL` and no `SCRIPTMETA.txt`.

- After bumping `SCRIPT_VERSION`, editing the overview or a Japanese README title, or adding, moving or renaming a script, run
  `python3 helpers/sync-scriptmeta.py` (`--check` reports files out of sync)
- The same tool writes `SCRIPT_SOURCE_URL` (the GitHub link to the latest source) as the last line of the README /
  article group in the basic info block, from the file's path, and realigns that group's `=`. Do not edit it by hand
- It also writes `SCRIPT_DESCRIPTION_JA` from the `# ` title of the README that `SCRIPT_README_JA` points to,
  as a one-line group between the first group and the README / article group, and uses the same title as
  SCRIPTMETA `Name`. To change it, edit the README title and run the tool. A script without a README gets a
  hand-written `SCRIPT_DESCRIPTION_JA` in the same style (a short phrase saying what it does); the tool keeps it and uses it as `Name`
- Script-ID is `com.swwwitch.indesign.<SCRIPT_NAME>`. Never change a published Script-ID, even when renaming a script;
  names with characters other than `A-Z a-z 0-9 . _ -` get an entry in `ID_OVERRIDES`
- `jsx/_templates/` is excluded

## UI

- Do not put `SCRIPT_VERSION` in a window title. Show it with the VersionFooter part (`jsx/_templates/VersionFooter.jsx`):
  call `addVersionFooter(window)` after every control is added, right before `show()`. It draws the script name and version (left) and "DTP Transit" (right) small and gray
  at the bottom left (dot glyphs in `onDraw`; needs the UITheme part for `isDarkUI()`) and takes over the button row's bottom margin. Progress, confirm and other
  small sub-windows get no footer. InDesign's ScriptUI ignores font sizes, which is why the digits are drawn as dots
