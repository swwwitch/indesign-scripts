#target indesign

/*
AddPagesUsingCurrentMaster.jsx

ページ挿入スクリプト（Adobe InDesign用）
Page Insertion Script for Adobe InDesign

-------------------------------
選択ダイアログを表示し、指定したページ数だけ現在のページの後に挿入します。
Displays a dialog and inserts the specified number of pages after the current page.

処理の流れ / Process flow:
1. ダイアログでページ数を入力（初期値 2） / Input the number of pages (default 2)
2. 現在のページの直後に指定数のページを追加 / Insert pages after current page
3. 現在のマスターを各ページに適用 / Apply current master to inserted pages
4. ユーザーがキャンセルした場合は処理を中止 / Cancel if user closes dialog

限定条件 / Requirements:
- InDesignドキュメントがアクティブであること / An active InDesign document is required

更新日：2025-06-26
Last updated: 2025-06-26
*/

(function () {

// -------------------------------
// 日英ラベル定義　Define label
// -------------------------------
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var lang = getCurrentLang();
var LABELS = {
    dialogTitle: {
        ja: "ページを挿入",
        en: "Insert Pages"
    },
    pageCountLabel: {
        ja: "挿入するページ数",
        en: "Number of pages to insert"
    },
    masterPrefix: {
        ja: "現在のマスター：",
        en: "Master:"
    }
};

// ScriptUI ダイアログでページ数を尋ねる関数
// Prompt user for number of pages to insert; returns null if cancelled
function showPageCountDialog(defaultPageCount) {
    var pageInsertDialog = new Window("dialog", LABELS.dialogTitle[lang]);
    pageInsertDialog.orientation = "column";
    pageInsertDialog.alignChildren = ["left", "top"];
    pageInsertDialog.spacing = 10;
    pageInsertDialog.margins = 20;

    var currentPage = app.activeWindow.activePage;
    var currentAppliedMaster = currentPage.appliedMaster;

    // マスター名表示グループ / Display current master name
    var masterNameGroup = pageInsertDialog.add("group");
    masterNameGroup.orientation = "row";
    masterNameGroup.alignChildren = ["left", "center"];

    var masterName = currentAppliedMaster ? currentAppliedMaster.name : "-";
    masterNameGroup.add("statictext", undefined, LABELS.masterPrefix[lang] + " " + masterName);

    // 入力フィールドグループ / Group for label and input field
    var pageCountInputGroup = pageInsertDialog.add("group");
    pageCountInputGroup.orientation = "row";
    pageCountInputGroup.alignChildren = ["left", "center"];

    pageCountInputGroup.add("statictext", undefined, LABELS.pageCountLabel[lang]);
    var pageCountInput = pageCountInputGroup.add("edittext", undefined, defaultPageCount.toString());
    pageCountInput.characters = 5;
    pageCountInput.active = true;

    // ボタングループ / Buttons group
    var dialogButtonGroup = pageInsertDialog.add("group");
    dialogButtonGroup.orientation = "row";
    dialogButtonGroup.alignment = "center";
    var cancelButton = dialogButtonGroup.add("button", undefined, "Cancel");
    var okButton = dialogButtonGroup.add("button", undefined, "OK");
    dialogButtonGroup.margins = [0, 10, 0, 0];

    var dialogResult = null;

    // OKボタン処理 / OK button handler
    okButton.onClick = function() {
        var enteredPageCount = parseInt(pageCountInput.text, 10);
        if (!isNaN(enteredPageCount) && enteredPageCount > 0) {
            dialogResult = {
                pageCount: enteredPageCount,
                selectedMaster: currentAppliedMaster
            };
            pageInsertDialog.close();
        } else {
            alert("1 以上の数値を入力してください / Please enter a number greater than 0.");
        }
    };

    // キャンセルボタン処理 / Cancel button handler
    cancelButton.onClick = function() {
        pageInsertDialog.close();
    };

    pageInsertDialog.show();
    return dialogResult;
}

// ページ挿入処理
// Insert specified number of pages after current page, applying current master
function insertPages() {
    var dialogResult = showPageCountDialog(2);
    if (!dialogResult) return; // ユーザーがキャンセルした場合は終了 / Exit if cancelled

    var pageCount = dialogResult.pageCount;
    var selectedMaster = dialogResult.selectedMaster;

    var currentPage = app.activeWindow.activePage;
    var spread = currentPage.parent;

    // 指定数だけページを追加し、マスターを適用 / Add pages and apply master
    for (var i = 0; i < pageCount; i++) {
        var newPage = spread.pages.add(LocationOptions.AFTER, currentPage);
        if (selectedMaster) newPage.appliedMaster = selectedMaster;
        currentPage = newPage;
    }
}

// スクリプト実行 / Run the script
insertPages();

})();