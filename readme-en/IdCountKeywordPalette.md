# Count keyword occurrences in a palette

[![Direct](https://img.shields.io/badge/Direct%20Link-IdCountKeywordPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdCountKeywordPalette.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCountKeywordPalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README-en.md)

---

A floating palette that counts how many times each of up to five keywords appears in the selected stories or in the whole document.

### Features

- Five keyword fields: list spelling variants side by side and count each of them at once
- Each count appears to the right of its field, and the total in the Total row (e.g. "7 (Story)")
- Skip matches inside other keywords (on by default): a keyword found as part of a longer keyword is not counted for the shorter one
- Two-column palette: the Matches list on the right shows the page and the surrounding text of each match (e.g. "1 | 前の【取引】の話")
- Set how many characters to show before and after with Characters Around (default 2, 0–50). Changing it rebuilds only the list, without counting again
- Click a row to select and show that text
- The page column shows "A-Parent" style names for parent pages, "Overset" for overset text and "Pasteboard" for text off the page
- Choose the scope: Story (the stories containing the selected text, text frames or table cells) or Document
- When several stories are selected, shows their total
- Also counts text in tables, footnotes, hidden layers, and locked layers and stories
- Options: Case Sensitive, Whole Word and Include Parent Pages (Include Parent Pages is for the document scope only)
- Keep the palette open and count again after changing the keyword or the selection
- Use Selection puts the selected text into the keyword field you last typed in (or the first empty one)
- Remembers the last keyword, scope and options until InDesign quits
- Japanese / English UI

### Usage

1. Select the text to count, or place the cursor in the target story (no selection is needed for the whole document)
2. Run the script (if the selected text fits in one paragraph, it opens with that text in the first empty field)
3. Set one to five keywords, the scope and options, then click Count (or press Return in the keyword field)
4. Click Close or press Esc to close

### Notes and limitations

- For Story, the scope is taken from the selection at the moment you click Count. Without a text selection it shows "Please select text."
- A `^` in the keyword is searched literally (special characters such as `^p` cannot be used)
- After counting, the Find what text in Find/Change is cleared. The find options are restored
- The list follows the order of the keyword fields, then the order found. Up to the first 500 matches are shown
- In the list, returns appear as "¶", tabs as spaces, and special characters such as page numbers are left out
- The total is the sum of the counts. The same keyword entered in two fields is counted twice
- Running the script again while the palette is open closes it and opens a new one

### Script info

| Item | Value |
| --- | --- |
| File | `jsx/text/IdCountKeywordPalette.jsx` |
| Version | v1.0.0 |
| Author | Masahiro Takano (@swwwitch) |
| First release | 2026-11-05 |
| Last updated | 2026-11-05 |

### Update History

#### v1.0.0 (2026-11-05)

- First release

### License

MIT License — <http://opensource.org/licenses/mit-license.php>
