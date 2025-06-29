# Adobe InDesign Scripts

- After publication, we continue to fix bugs and make adjustments as we use the scripts in daily work.
- If you encounter any issues or have suggestions like “It didn’t work as expected” or “It would be better if it did this,” we’d appreciate your feedback. If possible, please include the relevant document so we can verify and reproduce the issue (any artwork you send will not be shared externally).

[DTP Transitで公開しているスクリプトについて｜DTP Transit 別館](https://note.com/dtp_tranist/n/n60092f59a341)

## Pages

### Insert pages using the current parent page

By default, InDesign’s “Insert Pages” dialog does not reference the master page of the currently selected page.

This script refers to the master (parent) page of the currently selected page and inserts the specified number of pages directly after the current page.

https://github.com/swwwitch/indesign-scripts/blob/c580906e01ba767b8c08feba7b35deb693ab3a94/jsx/AddPagesUsingCurrentMaster.jsx

<img alt="" src="png/ss-420-332-72-20250626-135739.png" width="50%" />

## Paragraph styles

When the same text is repeated with the same paragraph style, this script appends numbering at the end.

![](png/ss-860-722-72-20250630-045123.png)

https://github.com/swwwitch/indesign-scripts/blob/509d5929089edb0523461ea2f49b262469fd9a84/jsx/AppendParagraphNumbering.jsx