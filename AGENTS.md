# AGENTS.md

This repository contains Adobe InDesign ExtendScript (`.jsx`).

## SCRIPTMETA

Every script carries a [SCRIPTMETA](https://github.com/Yamonov/SCRIPTMETA) block (Local Subset) right before
the header comment. It is generated from the basic info block and the 概要 / Overview text — do not edit it by hand.
Update checking is not used: write no `Meta-URL` and no `SCRIPTMETA.txt`.

- After bumping `SCRIPT_VERSION`, editing the overview, or adding a script, run `python3 helpers/sync-scriptmeta.py`
  (`--check` reports files out of sync)
- Script-ID is `com.swwwitch.indesign.<SCRIPT_NAME>`. Never change a published Script-ID, even when renaming a script;
  names with characters other than `A-Z a-z 0-9 . _ -` get an entry in `ID_OVERRIDES`
- `jsx/_templates/` is excluded
