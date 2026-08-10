# InDesignStyleSetup

Registers paragraph and character styles, their groups, inheritance and GREP styles in a single pass.

## Features

- Runs in four stages: create styles and groups, apply attributes, set GREP styles, and reorder
- basedOn relationships are always set, and shared typesetting settings live on the base styles
- Shared GREP rules (lang-US, no-break, inline-graphic) sit on base-regex and are inherited by child styles
- ul-li carries its own li-label rule, so the three shared rules are also set directly to avoid losing inheritance
- A progress bar is shown while running

## Usage

1. Open the target document
2. Run the script

## Notes and limitations

- Existing same-named styles are left untouched by default. Set `OVERWRITE_EXISTING_STYLES` to true to re-apply every attribute and rebuild the GREP rules.
- Attributes the script does not set (font, weight, colour, and so on) are never reset.
- The whole run is a single undo step.

## Script info

| Item | Value |
| --- | --- |
| File | `jsx/style/InDesignStyleSetup.jsx` |
| Version | v1.3.1 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-05-03 |
| Last updated | 2026-06-30 |
| Article | https://note.com/dtp_tranist/n/nfe87ec253780 |

## License

MIT License — <http://opensource.org/licenses/mit-license.php>
