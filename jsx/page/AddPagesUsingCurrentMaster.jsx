#target indesign

/*

### 概要

- ダイアログで挿入するページ数を指定し、現在のページの直後にページを追加します。
- 追加したページには現在のページのマスターを適用します。

### 処理の流れ

1. ダイアログでページ数を入力（初期値 2）
2. 現在のページの直後に指定数のページを追加
3. 現在のマスターを各ページに適用
4. ユーザーがキャンセルした場合は処理を中止

### 制限事項

- アクティブな InDesign ドキュメントが必要です。

---

### Overview

- Prompts for the number of pages to insert and adds them right after the current page.
- Applies the current page's master to the inserted pages.

### Process flow

1. Input the number of pages in the dialog (default 2)
2. Insert the specified number of pages after the current page
3. Apply the current master to each inserted page
4. Cancel the operation if the user closes the dialog

### Requirements

- An active InDesign document is required.

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.2.1";

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================
var DEFAULT_PAGE_COUNT = 2; /* ダイアログの初期ページ数 / Default page count in the dialog */

// =========================================
// ローカライズ / Localization
// =========================================
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "ページを挿入", en: "Insert Pages" }
    },
    field: {
        pageCount: { ja: "挿入するページ数", en: "Number of pages to insert" }
    },
    tooltip: {
        insertAfter: { ja: "指定ページの次に挿入", en: "Insert after the specified page" }
    },
    undo: {
        insertPages: { ja: "ページを挿入", en: "Insert Pages" }
    },
    page: {
        current: { ja: "現在のページ", en: "Current page" }
    },
    master: {
        current: { ja: "親（マスター）", en: "Master" }
    },
    alert: {
        invalidNumber: {
            ja: "1以上の数値を入力してください。",
            en: "Please enter a number greater than 0."
        },
        noDocument: {
            ja: "ドキュメントが開いていません。",
            en: "No document is open."
        }
    }
};

/* ドット区切りキーでラベルを取得 / Look up a label by dot-separated key */
function getLabel(key) {
    var node = LABELS;
    var parts = key.split(".");
    for (var idx = 0; idx < parts.length; idx++) {
        node = node[parts[idx]];
        if (!node) return key;
    }
    return node[currentLanguage] || node.en || key;
}

/* ラベル取得のショートハンド / Shorthand for getLabel */
function L(key) {
    return getLabel(key);
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return getLabel(key) + (currentLanguage === "ja" ? "：" : ": ");
}

// =========================================
// ダイアログ / Dialog
// =========================================
/* ラベル群の幅を最大値に揃える / Match all label widths to the widest one */
function alignLabelWidths(labelList) {
    var maxWidth = 0;
    for (var i = 0; i < labelList.length; i++) {
        var labelWidth = labelList[i].preferredSize.width;
        if (labelWidth > maxWidth) maxWidth = labelWidth;
    }
    for (var j = 0; j < labelList.length; j++) {
        labelList[j].preferredSize.width = maxWidth;
    }
}

/* ページ数を尋ねるダイアログを表示し、結果を返す（キャンセル時は null）/ Show the page-count dialog and return the result (null if cancelled) */
function showPageCountDialog(defaultPageCount) {
    var pageInsertDialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
    pageInsertDialog.orientation = "column";
    pageInsertDialog.alignChildren = ["left", "top"];
    pageInsertDialog.spacing = 10;
    pageInsertDialog.margins = 20;

    var currentPage = app.activeWindow.activePage;
    var currentAppliedMaster = currentPage.appliedMaster;

    var labelStatics = [];

    /* 現在のページ表示グループ / Display the current page */
    var currentPageGroup = pageInsertDialog.add("group");
    currentPageGroup.orientation = "row";
    currentPageGroup.alignChildren = ["left", "center"];

    var currentPageLabel = currentPageGroup.add("statictext", undefined, labelText("page.current"));
    labelStatics.push(currentPageLabel);
    currentPageGroup.add("statictext", undefined, currentPage.name);

    /* マスター名表示グループ / Display current master name */
    var masterNameGroup = pageInsertDialog.add("group");
    masterNameGroup.orientation = "row";
    masterNameGroup.alignChildren = ["left", "center"];

    var masterLabel = masterNameGroup.add("statictext", undefined, labelText("master.current"));
    labelStatics.push(masterLabel);
    var masterName = currentAppliedMaster ? currentAppliedMaster.name : "-";
    masterNameGroup.add("statictext", undefined, masterName);

    /* 入力フィールドグループ / Group for label and input field */
    var pageCountInputGroup = pageInsertDialog.add("group");
    pageCountInputGroup.orientation = "row";
    pageCountInputGroup.alignChildren = ["left", "center"];

    var pageCountLabel = pageCountInputGroup.add("statictext", undefined, labelText("field.pageCount"));
    labelStatics.push(pageCountLabel);
    var pageCountInput = pageCountInputGroup.add("edittext", undefined, defaultPageCount.toString());
    pageCountInput.characters = 5;
    pageCountInput.active = true;
    pageCountInput.helpTip = L("tooltip.insertAfter");
    pageCountLabel.helpTip = L("tooltip.insertAfter");

    alignLabelWidths(labelStatics);

    /* ボタングループ / Buttons group */
    var dialogButtonGroup = pageInsertDialog.add("group");
    dialogButtonGroup.orientation = "row";
    dialogButtonGroup.alignment = "center";
    var cancelButton = dialogButtonGroup.add("button", undefined, "Cancel", { name: "cancel" });
    var okButton = dialogButtonGroup.add("button", undefined, "OK", { name: "ok" });
    dialogButtonGroup.margins = [0, 10, 0, 0];

    var dialogResult = null;

    /* OKボタン処理 / OK button handler */
    okButton.onClick = function () {
        var enteredPageCount = parseInt(pageCountInput.text, 10);
        if (!isNaN(enteredPageCount) && enteredPageCount > 0) {
            dialogResult = {
                pageCount: enteredPageCount,
                selectedMaster: currentAppliedMaster
            };
            pageInsertDialog.close();
        } else {
            alert(L("alert.invalidNumber"));
        }
    };

    /* キャンセルボタン処理 / Cancel button handler */
    cancelButton.onClick = function () {
        pageInsertDialog.close();
    };

    pageInsertDialog.show();
    return dialogResult;
}

// =========================================
// ページ挿入 / Page insertion
// =========================================
/* ダイアログ入力に基づき、現在ページの直後へページを追加してマスターを適用 / Add pages after the current page and apply the master, based on dialog input */
function insertPages() {
    /* ドキュメントが開いていなければ中止 / Abort if no document is open */
    if (app.documents.length === 0) {
        alert(L("alert.noDocument"));
        return;
    }

    var dialogResult = showPageCountDialog(DEFAULT_PAGE_COUNT);
    if (!dialogResult) return; /* ユーザーがキャンセルした場合は終了 / Exit if cancelled */

    var pageCount = dialogResult.pageCount;
    var selectedMaster = dialogResult.selectedMaster;

    /* 一括 undo になるよう doScript でラップ / Wrap in doScript so the whole insertion is a single undo step */
    app.doScript(function () {
        var doc = app.activeDocument;
        var currentPage = app.activeWindow.activePage;

        /* ドキュメント基準で追加し、スプレッドへ正しく流し込む / Add at document level so pages reflow into proper spreads */
        for (var i = 0; i < pageCount; i++) {
            var newPage = doc.pages.add(LocationOptions.AFTER, currentPage);
            if (selectedMaster) newPage.appliedMaster = selectedMaster;
            currentPage = newPage;
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, L("undo.insertPages"));
}

// =========================================
// 実行 / Run
// =========================================
insertPages();

})();
