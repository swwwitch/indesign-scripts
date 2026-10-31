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
  SCRIPTMETA `Name`. Without a README there is no `SCRIPT_DESCRIPTION_JA` and `Name` is `SCRIPT_NAME`.
  To change it, edit the README title and run the tool
- Script-ID is `com.swwwitch.indesign.<SCRIPT_NAME>`. Never change a published Script-ID, even when renaming a script;
  names with characters other than `A-Z a-z 0-9 . _ -` get an entry in `ID_OVERRIDES`
- `jsx/_templates/` is excluded
