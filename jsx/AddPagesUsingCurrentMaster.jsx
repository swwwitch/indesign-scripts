#target indesign

/*
 * AddPagesUsingCurrentMaster.jsx
 *
 * 現在のページに適用されている親（マスター）ページを引き継いだまま、指定した枚数のページを直後に挿入します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddPagesUsingCurrentMaster";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-06-30";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/AddPagesUsingCurrentMaster.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/AddPagesUsingCurrentMaster.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* ダイアログの初期ページ数 / Default page count in the dialog */
var DEFAULT_PAGE_COUNT = 2;

// ==============================
// UIレイアウトの共通設定 / Shared UI layout
// ==============================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */

/**
 * ウィンドウの共通設定を適用する
 * @param {Window} win 対象ウィンドウ
 * @param {number} [spacing] 要素間隔。省略時は WINDOW_SPACING
 * @returns {void}
 */
function setupWindow(win, spacing) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * 行グループの共通設定を適用する（ボタン列など）
 * @param {Group} group 対象グループ
 * @param {string} [alignment] 配置。省略時は "left"
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// =========================================
// ラベル定義 / Labels
// =========================================

/**
 * UI 言語を判定する
 * @returns {string} "ja" または "en"
 */
function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var currentLanguage = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "ページを挿入", en: "Insert Pages" }
    },
    field: {
        pageCount: { ja: "挿入するページ数", en: "Number of pages to insert" }
    },
    page: {
        current: { ja: "現在のページ", en: "Current page" }
    },
    master: {
        current: { ja: "親（マスター）", en: "Parent page" }
    },
    button: {
        ok:     { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    tooltip: {
        insertAfter: { ja: "指定ページの次に挿入", en: "Insert after the specified page" }
    },
    undo: {
        insertPages: { ja: "ページを挿入", en: "Insert Pages" }
    },
    alert: {
        invalidNumber: { ja: "1以上の数値を入力してください。", en: "Please enter a number greater than 0." },
        noDocument:    { ja: "ドキュメントが開いていません。", en: "No document is open." }
    }
};

/**
 * ドット区切りキーでラベルを取得する
 * @param {string} labelKey 例: "dialog.title"
 * @returns {string} 現在の言語のラベル文字列。見つからない場合はキーをそのまま返す
 */
function getLabel(labelKey) {
    var node = LABELS;
    var keyParts = labelKey.split(".");
    for (var i = 0; i < keyParts.length; i++) {
        node = node[keyParts[i]];
        if (!node) return labelKey;
    }
    return node[currentLanguage] || node.en || labelKey;
}

/**
 * コロン付きラベルを取得する（日本語は全角コロン、英語は半角コロン）
 * @param {string} labelKey 例: "field.pageCount"
 * @returns {string} コロンを付与したラベル文字列
 */
function getLabelWithColon(labelKey) {
    return getLabel(labelKey) + (currentLanguage === "ja" ? "：" : ": ");
}

// =========================================
// ダイアログ / Dialog
// =========================================

/**
 * 見出しラベル群の幅を最大値に揃える
 * @param {Array<StaticText>} labelControls 幅を揃える StaticText の配列
 * @returns {void}
 */
function alignLabelWidths(labelControls) {
    var maxWidth = 0;
    for (var i = 0; i < labelControls.length; i++) {
        var labelWidth = labelControls[i].preferredSize.width;
        if (labelWidth > maxWidth) maxWidth = labelWidth;
    }
    for (var j = 0; j < labelControls.length; j++) {
        labelControls[j].preferredSize.width = maxWidth;
    }
}

/**
 * 挿入ページ数を尋ねるダイアログを表示する
 * @param {number} defaultPageCount 入力欄の初期値
 * @returns {{pageCount: number, appliedMaster: MasterSpread}|null} 入力結果。キャンセル時は null
 */
function showPageCountDialog(defaultPageCount) {
    var pageInsertDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(pageInsertDialog, 8);
    pageInsertDialog.alignChildren = ["left", "top"];

    var currentPage          = app.activeWindow.activePage;
    var currentAppliedMaster = currentPage.appliedMaster;

    var headingLabels = [];

    /* 現在のページ表示行 / Row showing the current page */
    var currentPageRow = pageInsertDialog.add("group");
    setupRow(currentPageRow, "left", 4);
    currentPageRow.alignChildren = ["left", "center"];

    var currentPageLabel = currentPageRow.add("statictext", undefined, getLabelWithColon("page.current"));
    headingLabels.push(currentPageLabel);
    currentPageRow.add("statictext", undefined, currentPage.name);

    /* 親ページ名表示行 / Row showing the applied parent page */
    var masterNameRow = pageInsertDialog.add("group");
    setupRow(masterNameRow, "left", 4);
    masterNameRow.alignChildren = ["left", "center"];

    var masterNameLabel = masterNameRow.add("statictext", undefined, getLabelWithColon("master.current"));
    headingLabels.push(masterNameLabel);
    masterNameRow.add("statictext", undefined, currentAppliedMaster ? currentAppliedMaster.name : "-");

    /* ページ数入力行 / Row for the page-count field */
    var pageCountRow = pageInsertDialog.add("group");
    setupRow(pageCountRow, "left", 4);
    pageCountRow.alignChildren = ["left", "center"];

    var pageCountLabel = pageCountRow.add("statictext", undefined, getLabelWithColon("field.pageCount"));
    headingLabels.push(pageCountLabel);
    var pageCountInput = pageCountRow.add("edittext", undefined, defaultPageCount.toString());
    pageCountInput.characters = 5;
    pageCountInput.active = true;
    pageCountInput.helpTip = getLabel("tooltip.insertAfter");
    pageCountLabel.helpTip = getLabel("tooltip.insertAfter");

    alignLabelWidths(headingLabels);

    /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
    var dialogButtonRow = pageInsertDialog.add("group");
    setupRow(dialogButtonRow, "center", 8);
    dialogButtonRow.margins = [0, 10, 0, 0];
    var cancelButton = dialogButtonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var okButton     = dialogButtonRow.add("button", undefined, getLabel("button.ok"), { name: "ok" });

    var dialogResult = null;

    okButton.onClick = function () {
        var enteredPageCount = parseInt(pageCountInput.text, 10);
        if (!isNaN(enteredPageCount) && enteredPageCount > 0) {
            dialogResult = {
                pageCount: enteredPageCount,
                appliedMaster: currentAppliedMaster
            };
            pageInsertDialog.close();
        } else {
            alert(getLabel("alert.invalidNumber"));
        }
    };

    cancelButton.onClick = function () {
        pageInsertDialog.close();
    };

    pageInsertDialog.show();
    return dialogResult;
}

// =========================================
// ページ挿入 / Page insertion
// =========================================

/**
 * ダイアログの入力に従い、現在ページの直後にページを追加して親ページを適用する
 * @returns {void}
 */
function insertPagesAfterCurrentPage() {
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var dialogResult = showPageCountDialog(DEFAULT_PAGE_COUNT);
    if (!dialogResult) return; /* キャンセル時は何もしない / Do nothing when cancelled */

    var pageCount     = dialogResult.pageCount;
    var appliedMaster = dialogResult.appliedMaster;

    /* 一括 undo になるよう doScript でラップ / Wrap in doScript so the whole insertion is a single undo step */
    app.doScript(function () {
        var activeDoc  = app.activeDocument;
        var insertAfterPage = app.activeWindow.activePage;

        /* ドキュメント基準で追加し、スプレッドへ正しく流し込む / Add at document level so pages reflow into proper spreads */
        for (var i = 0; i < pageCount; i++) {
            var newPage = activeDoc.pages.add(LocationOptions.AFTER, insertAfterPage);
            if (appliedMaster) newPage.appliedMaster = appliedMaster;
            insertAfterPage = newPage;
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.insertPages"));
}

// =========================================
// 実行 / Run
// =========================================
insertPagesAfterCurrentPage();

})();
