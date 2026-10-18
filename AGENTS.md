# AGENTS.md

This repository contains Adobe InDesign ExtendScript (`.jsx`).

## SCRIPTMETA

Every script carries a [SCRIPTMETA](https://github.com/Yamonov/SCRIPTMETA) block (Update Profile) right before
the header comment, and `SCRIPTMETA.txt` at the repository root lists the latest version of each Script-ID.
Both are generated from the basic info block and the 概要 / Overview text — do not edit them by hand.

- After bumping `SCRIPT_VERSION`, editing the overview, or adding a script, run `python3 helpers/sync-scriptmeta.py`
  and commit the `.jsx` together with `SCRIPTMETA.txt` (`--check` reports files out of sync)
- Script-ID is `com.swwwitch.indesign.<SCRIPT_NAME>`. Never change a published Script-ID, even when renaming a script;
  names with characters other than `A-Z a-z 0-9 . _ -` get an entry in `ID_OVERRIDES`
- `jsx/_templates/` is excluded
