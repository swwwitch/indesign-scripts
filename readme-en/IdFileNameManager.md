# Edit the file name by segment, then rename or save

[![Direct](https://img.shields.io/badge/Direct%20Link-IdFileNameManager.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdFileNameManager.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdFileNameManager.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Edits the active document's filename segment by segment (base / title / status / timestamp / page / version) and renames, saves as, or saves a copy.

### Features

- Segments: base, title, status, timestamp, page and version, with a customizable order
- Insert a production status (wip, draft, review, approved, flattened, and so on) from a dropdown
- Sequence numbers (page01 / page001) auto-bump to the folder maximum plus one on save
- Version format can be v1, v01 or v001, also auto-bumping to the folder maximum plus one
- Filename cleanup: separator unification, NFC normalization, and ASCII conversion of symbols and corporate abbreviations

### Usage

1. Open the target document
2. Run the script
3. Choose the mode, segments and order, then click OK

The destination folder is fixed before the dialog opens, so the incremented number shown in the preview matches the file that is actually saved. For an unsaved document, you are asked for a destination folder before the dialog opens.

### Mode: three ways to save

| Mode | Behavior |
| --- | --- |
| Rename Original | Save under the new name, then move the original to the Trash |
| Save As | Save under the new name. The original is kept, and the active document switches to the new file |
| Save a Copy | Save over the original, then create a copy under the new name. The active document stays on the original |

The default selection is always "Save As".

Only InDesign format (.indd) is ever written, so **"Rename Original" and "Save a Copy" are unavailable when a non-.indd document is open (IDML, for example).** The former would convert the format and then delete the original; the latter would produce a file whose contents are still in the original format but whose extension says .indd. Use "Save As" in that case — the original is kept.

### Timestamp

Choose from "None", "YYYYMMDD", and "YYYY-MM-DD". "Append HHMM" adds `-HHMM` at the end.

Turning on "Append HHMM" while the timestamp is "None" would print no time at all, so **"YYYYMMDD" is selected automatically at that moment**. Explicitly switching back to "None" afterwards is still allowed.

### Auto-increment

Sequence and version numbers are incremented by actually scanning the same folder. The scan targets ".indd files with the same prefix / suffix", so unrelated files do not pull the number around.

The bump runs **after** formatting (clean, convert, separator unification). Since existing filenames in the folder are stored already-formatted, matching against a pre-formatting name could miss the maximum and fail to bump.

A v-number or sequence is recognized only when it is **delimited by a separator (`-` `_` `.`) or by the ends of the name**. In `rev1-catalog-v02.indd`, the `rev1` is "revision 1" rather than a version, so the trailing `v02` is what gets bumped.

The folder listing is read once while the dialog is open and reused for the preview; it is re-read for the actual numbering after you press OK. This avoids a full scan on every keystroke in folders holding thousands of files.

### Notes and limitations

- When the original has no v-number, the version is added in v01 form.
- Settings are saved on each run and become the defaults next time.
- A name consisting only of symbols or emoji reduces to an empty string after formatting, so it is rejected instead of being saved.

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/document/IdFileNameManager.jsx` |
| Version | v1.3.2 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-05-27 |
| Last updated | 2026-09-12 |
| Article | https://note.com/dtp_tranist/n/nc88dd887eb1c |

### Changelog

- v1.3.1 (2026-05-29) Segment editing, ordering and formatting in place
- v1.3.2 (2026-09-12) Restored `setupPanel`, which was called from five places while undefined, and fixed two leftover `L()` calls to `getLabel()` — either one prevented the dialog from opening. Fixed the three "Segment Order" radio buttons not being mutually exclusive. Fixed an argument mismatch that passed `prefs` into the sort panel's `currentOrderAvailable` (leaving "Match Current" always enabled). Fixed a case where a rename differing only in letter case or kana composition sent the freshly saved file to the Trash. Rename and Save a Copy are now unavailable for non-.indd documents. Moved the empty-name check to run **after** formatting. Moved numbering to run after formatting. Fixed "Save a Copy" failing after an overwrite was approved, because `File.copy()` does not overwrite. Version and sequence numbers are now matched as delimited segments. Turning on "Append HHMM" now selects a timestamp format automatically. The destination folder is now fixed before the dialog opens. Unified the formatting pipeline shared by the preview and the actual save, and cached the folder listing. Aligned the JSDoc with the real signatures and removed unused layout helpers.

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
