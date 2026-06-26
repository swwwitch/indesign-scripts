#target indesign

/*

### 概要

- 選択した段落の「段落の開始位置」を「次の段（フレーム）」に設定します。
- テキストフレームの最終行などを次の段や次のフレームへ送りたいときに使います。

### 処理の流れ

- ドキュメントと選択の有無を確認
- 選択が段落開始位置を持てるテキストか確認
- `startParagraph` に `NEXT_COLUMN` を設定

*/

/*

### Overview

- Sets the "Start Paragraph" option of the selected paragraph to "Next Column (Frame)".
- Useful when you want to push the last line of a text frame to the next column or frame.

### Flow

- Check that a document and a selection exist
- Verify that the selection is text that supports the start-paragraph option
- Set `startParagraph` to `NEXT_COLUMN`

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

// =========================================
// ローカライズ / Localization
// =========================================

/* 言語判定（ja で始まれば日本語、それ以外は英語）/ Detect language (ja-prefixed locale is Japanese, otherwise English) */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

var LABELS = {
    alert: {
        noDocument: {
            ja: "ドキュメントを開いてください。",
            en: "Please open a document."
        },
        noSelection: {
            ja: "段落またはテキストを選択してください。",
            en: "Please select a paragraph or text."
        },
        unsupported: {
            ja: "選択対象には段落開始位置を設定できません。",
            en: "The selected object does not support paragraph start options."
        }
    }
};

/* ネストしたキーからラベルを取得 / Resolve a label from a nested key path */
function L(category, key) {
    return LABELS[category][key][currentLanguage];
}

// =========================================
// メイン処理 / Main
// =========================================
(function () {
    /* ドキュメントが開いているか確認 / Make sure a document is open */
    if (app.documents.length === 0) {
        alert(L("alert", "noDocument"));
        return;
    }

    var currentDocument = app.activeDocument;

    /* 選択があるか確認 / Make sure something is selected */
    if (currentDocument.selection.length === 0) {
        alert(L("alert", "noSelection"));
        return;
    }

    var selectedTextSelection = currentDocument.selection[0];

    /* 段落開始位置を設定できる対象か確認 / Verify the selection supports the start-paragraph option */
    if (!selectedTextSelection || !selectedTextSelection.hasOwnProperty("startParagraph")) {
        alert(L("alert", "unsupported"));
        return;
    }

    /* 段落の開始位置を「次の段」に設定 / Set the paragraph to start in the next column */
    selectedTextSelection.startParagraph = StartParagraph.NEXT_COLUMN;
}());
