# Count keywords and unify spelling variants

[![Direct](https://img.shields.io/badge/Direct%20Link-IdCountKeywordPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdCountKeywordPalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCountKeywordPalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

A floating palette that counts how many times each of up to five keywords appears in the selected stories or in the whole document, lists where they are, and unifies spelling variants.

### Features

- Five keyword fields: list spelling variants side by side and count each of them at once
- Each count appears to the right of its field, and the total in the Total row (e.g. "7 (Story)")
- Skip parts of longer keywords in the Keywords panel (on by default): a keyword found as part of a longer keyword is not counted for the shorter one
- Two-column palette: the Matches list on the right shows the page and the surrounding text of each match (e.g. "1 | 前の【取引】の話")
- Set how many characters to show before and after with the Characters Around slider (default 3, 0–20). Releasing it rebuilds only the list, without counting again
- Hide matches at line start or end (on by default): matches right after or before a line break (a return, forced line break, column/frame/page break, or the start or end of a story or cell) are left out of the list and of Unify (they are still counted)
- Wrap matches in 【】 (off by default): wraps each match in 【】 in the list
- Click a row to select that text and show it at the zoom set with the Zoom slider (default 200%, 50–800%)
- **Exclude**: select a row and click Exclude, or double-click the row, to exclude it (it is unchecked). Excluded rows are left out of the counts and of Unify All, and the Total row shows how many are excluded. Click again to include it. Exclusions stay after counting again or replacing
- **Unify all**: choose the target with the radio button to the left of a keyword field. Rows of the other fields get "→ target". Unify All replaces all rows that are not excluded with the target
- **Change one at a time**: the picked text goes into Change to; edit it and click Change. The next row of the same keyword is selected afterwards
- Replacing can be undone in one step (Cmd+Z), and the palette counts again afterwards
- The page column shows "A-Parent" style names for parent pages, "Overset" for overset text and "Pasteboard" for text off the page
- Choose the scope: Story (the stories containing the selected text, text frames or table cells) or Document (Document by default)
- When several stories are selected, shows their total
- Also counts text in tables, footnotes, hidden layers, and locked layers and stories
- Options: Case Sensitive, Whole Word, Include Parent Pages (document scope only) and Exclude Tables of Contents (on by default; stories made with Table of Contents are not searched)
- Load Selection puts the selected text into the keyword field you last typed in (or the first empty one)
- Remembers the last keywords, scope, options and unify target until InDesign quits
- Japanese / English UI

### Usage

1. Select the text to count, or place the cursor in the target story (no selection is needed for the whole document)
2. Run the script (if the selected text fits in one paragraph, it opens with that text in the first empty field)
3. Set one to five keywords, the scope and options, then click Find (or press Return in a keyword field)
4. To unify, choose the target radio button, exclude the rows you want to keep, and click Unify All. To fix one match at a time, pick a row, edit Change to and click Change
5. Click Close or press Esc to close

### Notes and limitations

- For Story, the scope is taken from the selection at the moment you click Find. Without a text selection it shows "Please select text."
- After changing keywords, the scope or the find options, Unify All and Change stay disabled until you click Find again
- Matches whose text was edited after counting are skipped, and the number skipped is shown
- Replaced text takes the formatting of the first character of the original match
- A `^` in the keyword is searched literally (special characters such as `^p` cannot be used)
- After counting, the Find what text in Find/Change is cleared. The find options are restored
- The list follows the order of the keyword fields, then the order found. Up to the first 500 matches are shown (Unify All also replaces the ones not shown)
- In the list, returns appear as "¶", tabs as spaces, and special characters such as page numbers are left out
- The total is the sum of the counts. The same keyword entered in two fields is counted twice
- The version appears at the bottom left of the palette
- Running the script again while the palette is open closes it and opens a new one

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/text/IdCountKeywordPalette.jsx` |
| Version | v1.3.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-05 |
| Last updated | 2026-11-05 |
| Article | https://note.com/dtp_tranist/n/n9f6d4d2bbeed |

### Update History

#### v1.3.0 (2026-11-05)

- Reworded the UI to match InDesign's Find/Change: Unify → Unify All, Apply → Change, Replace with → Change to, Use Selection → Load Selection, and clearer option names. The Keywords panel title explains the ○ unify-target mark. Added a tooltip to Find
- Internal cleanup (panels built by separate functions, duplicate code merged)

#### v1.2.1 (2026-11-05)

- Renamed the Count button to Find

#### v1.2.0 (2026-11-05)

- Added a × button to the right of each keyword field to clear it
- Moved the Characters Around and Zoom sliders to the top of the Matches panel
- Added space above the Total row

#### v1.1.1 (2026-11-05)

- Added space below the Zoom row
- Renamed the palette title to "Spelling Consistency Helper"

#### v1.1.0 (2026-11-05)

- Added Unify, which replaces matches with the chosen target at once, and Replace with / Apply for one match at a time
- Added Exclude to leave list rows out of the counts and of Unify (double-click also toggles)
- Added the options Hide matches next to line breaks, Mark with 【】 and Skip Tables of Contents
- Characters Around and Zoom are now sliders; the picked text is shown at the set zoom
- The right panel is now named Matches in English and 該当箇所 in Japanese. Document is the default scope
- Moved the version from the title bar to the bottom left, dimmed (shared part added)

#### v1.0.0 (2026-11-05)

- First release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
