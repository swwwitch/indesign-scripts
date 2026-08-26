# IdFileNameManager

[![Direct](https://img.shields.io/badge/Direct%20Link-IdFileNameManager.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdFileNameManager.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdFileNameManager.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

Edits the active document's filename segment by segment (base / title / status / timestamp / page / version) and renames, saves as, or saves a copy.

## Features

- Segments: base, title, status, timestamp, page and version, with a customizable order
- Insert a production status (wip, draft, review, approved, flattened, and so on) from a dropdown
- Sequence numbers (page01 / page001) auto-bump to the folder maximum plus one on save
- Version format can be v1, v01 or v001, also auto-bumping to the folder maximum plus one
- Filename cleanup: separator unification, NFC normalization, and ASCII conversion of symbols and corporate abbreviations

## Usage

1. Open the target document
2. Run the script
3. Choose the mode, segments and order, then click OK

## Notes and limitations

- The destination is always the folder of the current file; unsaved documents prompt for a folder.
- When the original has no v-number, the version is added in v01 form.
- Settings are saved on each run and become the defaults next time.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/document/IdFileNameManager.jsx` |
| Version | v1.3.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-05-27 |
| Last updated | 2026-05-29 |
| Article | https://note.com/dtp_tranist/n/nc88dd887eb1c |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
