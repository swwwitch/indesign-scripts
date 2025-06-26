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
    promptMessage: {
        ja: "挿入するページ数",
        en: "Number of pages to insert"
    }
};

// ScriptUI ダイアログでページ数を尋ねる関数
// Prompt user for number of pages to insert; returns null if cancelled
function promptPageCountUI(defaultCount) {
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = ["left", "top"];
    dialog.spacing = 10;
    dialog.margins = 20;

    var currentPage = app.activeWindow.activePage;
    var currentAppliedMaster = currentPage.appliedMaster;

    // マスター名表示グループ / Display current master name
    var masterGroup = dialog.add("group");
    masterGroup.orientation = "row";
    masterGroup.alignChildren = ["left", "center"];

    var masterName = currentAppliedMaster ? currentAppliedMaster.name : "-";
    masterGroup.add("statictext", undefined, (lang === "ja" ? "現在のマスター：" : "Master:") + " " + masterName);

    // 入力フィールドグループ / Group for label and input field
    var inputGroup = dialog.add("group");
    inputGroup.orientation = "row";
    inputGroup.alignChildren = ["left", "center"];

    inputGroup.add("statictext", undefined, LABELS.promptMessage[lang]);
    var input = inputGroup.add("edittext", undefined, defaultCount.toString());
    input.characters = 5;
    input.active = true;

    // ボタングループ / Buttons group
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";
    var cancelBtn = buttonGroup.add("button", undefined, "Cancel");
    var okBtn = buttonGroup.add("button", undefined, "OK");
    buttonGroup.margins = [0, 10, 0, 0];

    var result = null;

    // OKボタン処理 / OK button handler
    okBtn.onClick = function() {
        var num = parseInt(input.text, 10);
        if (!isNaN(num) && num > 0) {
            result = {
                pageCount: num,
                selectedMaster: currentAppliedMaster
            };
            dialog.close();
        } else {
            alert("1 以上の数値を入力してください / Please enter a number greater than 0.");
        }
    };

    // キャンセルボタン処理 / Cancel button handler
    cancelBtn.onClick = function() {
        dialog.close();
    };

    dialog.show();
    return result;
}

// ページ挿入処理
// Insert specified number of pages after current page, applying current master
function insertPages() {
    var result = promptPageCountUI(2);
    if (!result) return; // ユーザーがキャンセルした場合は終了 / Exit if cancelled

    var pageCount = result.pageCount;
    var selectedMaster = result.selectedMaster;

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