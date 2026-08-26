#target indesign

/*
 * IdAddPagesUsingCurrentMaster.jsx
 *
 * 現在のページに適用されている親（マスター）ページを引き継いだまま、指定した枚数のページを直後に挿入します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdAddPagesUsingCurrentMaster"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-14";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAddPagesUsingCurrentMaster.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAddPagesUsingCurrentMaster.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n2d1a76097a39"; /* 紹介記事 / article URL */

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
 * @param {Window} targetWindow 対象ウィンドウ
 * @param {number} [spacing] 要素間隔。省略時は WINDOW_SPACING
 * @returns {void}
 */
function setupWindow(targetWindow, spacing) {
    targetWindow.orientation = "column";
    targetWindow.alignChildren = "fill";
    targetWindow.margins = WINDOW_MARGINS;
    targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * 行グループの共通設定を適用する（ボタン列など）
 * @param {Group} targetGroup 対象グループ
 * @param {string} [alignment] 配置。省略時は "left"
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupRow(targetGroup, alignment, spacing) {
    targetGroup.orientation = "row";
    targetGroup.alignment = alignment || "left";
    targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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
        currentPage:    { ja: "現在のページ", en: "Current page" },
        appliedMaster:  { ja: "親（マスター）", en: "Parent page" },
        pageCount:      { ja: "挿入するページ数", en: "Number of pages to insert" }
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
        noDocument:    { ja: "ドキュメントが開いていません。", en: "No document is open." },
        noLayoutWindow: {
            ja: "レイアウトウィンドウがアクティブではありません。ドキュメントのページを表示してから実行してください。",
            en: "No layout window is active. Switch to a document layout view and try again."
        },
        masterPageActive: {
            ja: "親（マスター）ページが表示されています。ドキュメントのページを表示してから実行してください。",
            en: "A parent (master) page is displayed. Switch to a document page and try again."
        }
    }
};

/**
 * ドット区切りキーでラベルを取得する
 * @param {string} labelKey 例: "dialog.title"
 * @returns {string} 現在の言語のラベル文字列。見つからない場合はキーをそのまま返す
 */
function getLabel(labelKey) {
    var labelNode = LABELS;
    var keyParts = labelKey.split(".");
    for (var i = 0; i < keyParts.length; i++) {
        labelNode = labelNode[keyParts[i]];
        if (!labelNode) return labelKey;
    }
    return labelNode[currentLanguage] || labelNode.en || labelKey;
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
// 実行前のチェック / Preconditions
// =========================================

/**
 * レイアウトウィンドウがアクティブかどうかを判定する
 * @returns {boolean} レイアウトウィンドウなら true
 */
function isLayoutWindowActive() {
    try {
        return (app.activeWindow.constructor.name === "LayoutWindow");
    } catch (e) {
        return false;
    }
}

/**
 * 指定ページが親（マスター）ページかどうかを判定する
 * @param {Page} targetPage 判定するページ
 * @returns {boolean} 親ページなら true
 */
function isMasterPage(targetPage) {
    try {
        return (targetPage.parent instanceof MasterSpread);
    } catch (e) {
        return false;
    }
}

// =========================================
// ダイアログ / Dialog
// =========================================

/**
 * 全角数字を半角に変換する
 * @param {string} inputText 変換する文字列
 * @returns {string} 全角数字を半角に置き換えた文字列
 */
function toHalfWidthDigits(inputText) {
    return String(inputText).replace(/[０-９]/g, function (fullWidthDigit) {
        return String.fromCharCode(fullWidthDigit.charCodeAt(0) - 0xFEE0);
    });
}

/**
 * 入力文字列を挿入ページ数として解釈する
 * @param {string} inputText 入力欄の文字列
 * @returns {number} 1以上の整数。解釈できない場合は 0
 */
function parsePageCount(inputText) {
    var normalizedText = toHalfWidthDigits(inputText).replace(/^\s+|\s+$/g, "");
    if (!/^[0-9]+$/.test(normalizedText)) return 0;
    var parsedCount = parseInt(normalizedText, 10);
    return (parsedCount > 0) ? parsedCount : 0;
}

/**
 * 各行の左側ラベルの幅を最大値に揃える
 * @param {Array<StaticText>} rowLabelControls 幅を揃える StaticText の配列
 * @returns {void}
 */
function alignLabelWidths(rowLabelControls) {
    var maxLabelWidth = 0;
    for (var i = 0; i < rowLabelControls.length; i++) {
        var labelWidth = rowLabelControls[i].preferredSize.width;
        if (labelWidth > maxLabelWidth) maxLabelWidth = labelWidth;
    }
    for (var j = 0; j < rowLabelControls.length; j++) {
        rowLabelControls[j].preferredSize.width = maxLabelWidth;
    }
}

/**
 * 挿入ページ数を尋ねるダイアログを表示する
 * @param {number} defaultPageCount 入力欄の初期値
 * @returns {{pageCount: number, insertAfterPage: Page, appliedMaster: MasterSpread}|null} 入力結果。キャンセル時は null
 */
function showPageCountDialog(defaultPageCount) {
    var pageInsertDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(pageInsertDialog, 8);
    pageInsertDialog.alignChildren = ["left", "top"];

    var currentPage          = app.activeWindow.activePage;
    var currentAppliedMaster = currentPage.appliedMaster;

    var rowLabelControls = [];

    /* 現在のページ表示行 / Row showing the current page */
    var currentPageRow = pageInsertDialog.add("group");
    setupRow(currentPageRow, "left", 4);
    currentPageRow.alignChildren = ["left", "center"];

    var currentPageLabel = currentPageRow.add("statictext", undefined, getLabelWithColon("field.currentPage"));
    rowLabelControls.push(currentPageLabel);
    currentPageRow.add("statictext", undefined, currentPage.name);

    /* 親ページ名表示行 / Row showing the applied parent page */
    var appliedMasterRow = pageInsertDialog.add("group");
    setupRow(appliedMasterRow, "left", 4);
    appliedMasterRow.alignChildren = ["left", "center"];

    var appliedMasterLabel = appliedMasterRow.add("statictext", undefined, getLabelWithColon("field.appliedMaster"));
    rowLabelControls.push(appliedMasterLabel);
    appliedMasterRow.add("statictext", undefined, currentAppliedMaster ? currentAppliedMaster.name : "-");

    /* ページ数入力行 / Row for the page-count field */
    var pageCountRow = pageInsertDialog.add("group");
    setupRow(pageCountRow, "left", 4);
    pageCountRow.alignChildren = ["left", "center"];

    var pageCountLabel = pageCountRow.add("statictext", undefined, getLabelWithColon("field.pageCount"));
    rowLabelControls.push(pageCountLabel);
    var pageCountInput = pageCountRow.add("edittext", undefined, defaultPageCount.toString());
    pageCountInput.characters = 5;
    pageCountInput.active = true;
    pageCountInput.helpTip = getLabel("tooltip.insertAfter");
    pageCountLabel.helpTip = getLabel("tooltip.insertAfter");

    alignLabelWidths(rowLabelControls);

    /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
    var dialogButtonRow = pageInsertDialog.add("group");
    setupRow(dialogButtonRow, "center", 8);
    dialogButtonRow.margins = [0, 10, 0, 0];
    /* キャンセルは name: "cancel" の既定動作（クリック・ESC で閉じる）に任せる / Cancel relies on the built-in behavior of name: "cancel" */
    dialogButtonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var okButton = dialogButtonRow.add("button", undefined, getLabel("button.ok"), { name: "ok" });

    var pageInsertSettings = null;

    okButton.onClick = function () {
        var enteredPageCount = parsePageCount(pageCountInput.text);
        if (enteredPageCount > 0) {
            pageInsertSettings = {
                pageCount: enteredPageCount,
                insertAfterPage: currentPage,
                appliedMaster: currentAppliedMaster
            };
            pageInsertDialog.close();
        } else {
            alert(getLabel("alert.invalidNumber"));
        }
    };

    pageInsertDialog.show();
    return pageInsertSettings;
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

    /* 挿入位置はアクティブなレイアウトウィンドウのページから決まる / The insertion point comes from the active layout window */
    if (!isLayoutWindowActive()) {
        alert(getLabel("alert.noLayoutWindow"));
        return;
    }

    /* 親ページは document.pages のメンバーではないため挿入位置に使えない / A parent page is not a member of document.pages */
    if (isMasterPage(app.activeWindow.activePage)) {
        alert(getLabel("alert.masterPageActive"));
        return;
    }

    var pageInsertSettings = showPageCountDialog(DEFAULT_PAGE_COUNT);
    if (!pageInsertSettings) return; /* キャンセル時は何もしない / Do nothing when cancelled */

    var pageCount     = pageInsertSettings.pageCount;
    var appliedMaster = pageInsertSettings.appliedMaster;

    /* 一括 undo になるよう doScript でラップ / Wrap in doScript so the whole insertion is a single undo step */
    app.doScript(function () {
        var activeDocument  = app.activeDocument;
        var insertAfterPage = pageInsertSettings.insertAfterPage;

        /* ドキュメント基準で追加し、スプレッドへ正しく流し込む / Add at document level so pages reflow into proper spreads */
        for (var i = 0; i < pageCount; i++) {
            var insertedPage = activeDocument.pages.add(LocationOptions.AFTER, insertAfterPage);
            /* 現在のページが［なし］なら挿入したページも［なし］にそろえる / Match [None] as well, not just a named parent */
            insertedPage.appliedMaster = appliedMaster ? appliedMaster : NothingEnum.nothing;
            insertAfterPage = insertedPage;
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.insertPages"));
}

// =========================================
// 実行 / Run
// =========================================
insertPagesAfterCurrentPage();

})();
