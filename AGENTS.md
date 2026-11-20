# AGENTS.md

This repository contains Adobe InDesign ExtendScript (`.jsx`).

## SCRIPTMETA (basic info)

A script's basic info lives only in the SCRIPTMETA block at the very top of the file, in the SCRIPTMETA 2.1
multilingual form ([MetadataAuthoring.md](https://github.com/Yamonov/SCRIPTMETAKit/blob/main/docs/MetadataAuthoring.md)).
There is no hand-written `var SCRIPT_* = "…"` block and no `### 概要` / `### Overview` header comment any more.
Update checking is not used: write no `Meta-URL` and no `SCRIPTMETA.txt`.

```js
/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.<ScriptName>
Meta-Default-language=ja
Name=<README の見出し>
Name[en]=<English README title>
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.6.0
Release-Date=2026-02-24
Tags=<folder tags>
X-com.swwwitch-Script-Name=<ScriptName>
X-com.swwwitch-Updated=2026-07-26
X-com.swwwitch-README-JA=https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/<ScriptName>.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/<ScriptName>.md
X-com.swwwitch-Article-URL=https://note.com/dtp_tranist/n/xxxxxxxx
X-com.swwwitch-Source-URL=https://github.com/swwwitch/indesign-scripts/blob/main/jsx/<category>/<ScriptName>.jsx
Description-BEGIN
[ja]
（1〜2文の説明）
[en]
(one or two sentences)
Description-END
SCRIPTMETA-END
*/
```

- Edit by hand: `Version` (no `v`), `Author`, `Release-Date` (first release), the Description, and the
  `X-com.swwwitch-*` fields `Script-Name`, `Modified-By`, `Updated`, `README-JA` / `README-EN` (only when the README
  files exist), `Article-URL`, `Reference-URL`, `Pro-URL`. Leave out a field that has no value
- `X-com.swwwitch-Script-Name` becomes `SCRIPT_NAME`, which is also the key of saved settings and dialog positions.
  Changing it loses what users have saved
- Description: `[ja]` and `[en]` labels on their own lines, one or two sentences each, then point nowhere else
  (the README links are in the fields). Write a literal `[`, `]`, `\` as `\[`, `\]`, `\\`. Never `*/`
- Notes that do not fit in the Description (`### 注意` / `### Notes`) go in a `/* */` comment right after the SCRIPTMETA block
- When adapting someone else's script, `Author` is the original author and `X-com.swwwitch-Modified-By=Masahiro Takano (@swwwitch)`.
  Put the original credit (`@author` / `@discussion`) in a JSDoc block after the Script info part
- `python3 helpers/sync-scriptmeta.py` (`--check` reports files out of sync) fills in the rest. Run it after editing
  SCRIPTMETA or a README title, and after adding, moving or renaming a script:
  - `Script-ID` only when it is missing (`com.swwwitch.indesign.<Script-Name>`; names with characters other than
    `A-Z a-z 0-9 . _ -` get an entry in `ID_OVERRIDES`). Never change a published Script-ID, even when renaming a script
  - `Name` / `Name[en]` from the `# ` titles of the READMEs that `README-JA` / `README-EN` point to. To change them, edit
    the README titles. A script without a README keeps a hand-written `Name` (a short phrase saying what it does)
  - `Meta-Default-language`, `Target-App`, `Tags` (from the top folder under `jsx/`, `FOLDER_TAGS` in the tool; a new folder
    needs an entry there), `X-com.swwwitch-Source-URL` (from the file's path), percent-encoding of non-ASCII in URLs,
    and the field order
  - the Script info part below, rewritten to match `jsx/_templates/ScriptMeta.jsx`
- `jsx/_templates/` is excluded

## Basic info (Script info part)

The 基本情報 / Basic info block holds the Script info part (`jsx/_templates/ScriptMeta.jsx`), at the top level outside the IIFE:

```js
// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
…
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php
```

It reads the SCRIPTMETA of `$.fileName` when the script loads and sets `SCRIPT_NAME`, `SCRIPT_VERSION` (`"v" + Version`),
`SCRIPT_AUTHOR`, `SCRIPT_RELEASED`, `SCRIPT_UPDATED`, `SCRIPT_DESCRIPTION_JA` (`Name`), `SCRIPT_README_JA` / `SCRIPT_README_EN`,
`SCRIPT_ARTICLE_URL` and `SCRIPT_SOURCE_URL` (empty string when a field is missing). Use those variables as before; do not
assign them by hand and do not edit the part (the sync tool rewrites it). In code, write the SCRIPTMETA marker strings split
(`"SCRIPTMETA" + "-BEGIN"`) so apps that scan for SCRIPTMETA do not pick them up.

## Version

Bump `Version` in SCRIPTMETA whenever a change alters behavior; do not wait to be told.
Version numbers are tied to the READMEs and published articles, so keep them in step.

- Bug fix or small behavior change: bump the patch (`1.6.0` → `1.6.1`; two-part versions grow a third part, `1.1` → `1.1.1`)
- Larger feature addition: bump the minor (`1.6.1` → `1.7.0`)
- Set `X-com.swwwitch-Updated` to that day and add the same version to the change log in `readme-ja/` and `readme-en/`
  (create the section when it is missing)
- Do not bump for links, comments, alignment or JSDoc only
- Bump once per unreleased change set: count from the version in the last commit, not from today's working copy
- New scripts start at `1.0.0`

## UI

- Do not put `SCRIPT_VERSION` in a window title. Show it with the VersionFooter part (`jsx/_templates/VersionFooter.jsx`):
  call `addVersionFooter(window)` after every control is added, right before `show()`. It draws the script name and version (left) and "DTP Transit" (right) small and gray
  at the bottom left (dot glyphs in `onDraw`; needs the UITheme part for `isDarkUI()`) and takes over the button row's bottom margin. Progress, confirm and other
  small sub-windows get no footer. InDesign's ScriptUI ignores font sizes, which is why the digits are drawn as dots
  Clicking the name and version opens the part's About This Script dialog (script name with a copy button, the Japanese README title,
  version, README / source code / article links, Show in Finder — in Path Finder when it is running, through
  `/Applications/OpenInFileViewer.app`). It reads the `SCRIPT_*` basic info with `typeof` and carries its own
  labels (`VERSION_INFO_LABELS`), so paste the whole part unchanged; do not add its labels to `LABELS`
