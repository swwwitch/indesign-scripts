#target indesign

/*
 * IdEditTextsByDialog.jsx
 *
 * 複数行入力ダイアログでテキストを編集し、選択範囲の置換・カーソル位置への挿入・新規テキストフレーム作成を行います。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdEditTextsByDialog";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v0.1.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-06-26";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdEditTextsByDialog.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdEditTextsByDialog.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 強制改行（\n）を入力欄で表す可視マーカー / Visible marker that stands for a forced line break (\n) in the input field */
var SOFT_BREAK_MARKER = "@#";

/* 新規テキストフレームの幅の目安（1 文字あたりの推定幅と上下限）/ Width estimate for a new text frame (per-character width and its bounds) */
var NEW_FRAME_WIDTH_PER_CHAR = 12;
var NEW_FRAME_WIDTH_MIN      = 100;
var NEW_FRAME_WIDTH_MAX      = 800;
var NEW_FRAME_HEIGHT         = 20;

// =========================================
// レイアウト設定 / Layout settings
// =========================================

/* 入力欄のサイズ [幅, 高さ]（px）/ Size of the input field [width, height] (px) */
var INPUT_BOX_SIZE = [350, 160];

/* ボタン列を左右に分けるスペーサーの幅（px）/ Width of the spacer that splits the button row (px) */
var BUTTON_ROW_SPACER_WIDTH = 30;

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

var currentLang = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "テキスト編集", en: "Edit Text" },
        note:  { ja: "@#で強制改行（\\n）、fn + returnで確定", en: "@# = forced line break (\\n), fn + return to confirm" }
    },
    button: {
        ok:              { ja: "OK", en: "OK" },
        cancel:          { ja: "キャンセル", en: "Cancel" },
        clearLineBreaks: { ja: "改行全削除", en: "Clear All" },
        insertSoftBreak: { ja: "@# 挿入", en: "Insert @#" }
    },
    alert: {
        noTextFrame:   { ja: "テキストフレームを選択してください。", en: "Please select a text frame." },
        errorOccurred: { ja: "エラーが発生しました：\n", en: "An error occurred:\n" }
    },
    undo: {
        editText: { ja: "テキスト編集", en: "Edit Text" }
    }
};

/**
 * ラベルを現在の言語で取得する
 * @param {object} labelEntry ja / en を持つラベルオブジェクト
 * @returns {string} 現在の言語のラベル文字列
 */
function localize(labelEntry) {
    return labelEntry[currentLang];
}

// =========================================
// ダイアログ / Dialog
// =========================================

/**
 * 複数行テキストの編集ダイアログを表示する
 * @param {string} initialText 入力欄の初期値
 * @returns {string|null} 入力されたテキスト。キャンセル時は null
 */
function showMultilineTextDialog(initialText) {
    var textEditDialog = new Window("dialog", localize(LABELS.dialog.title) + " " + SCRIPT_VERSION);
    setupWindow(textEditDialog, 8);

    var inputBox = textEditDialog.add("edittext", undefined, initialText || "", { multiline: true });
    inputBox.preferredSize = INPUT_BOX_SIZE;

    var noteRow = textEditDialog.add("group");
    setupRow(noteRow, "center", 0);
    noteRow.add("statictext", undefined, localize(LABELS.dialog.note));

    /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
    var dialogButtonRow = textEditDialog.add("group");
    setupRow(dialogButtonRow, "fill", 8);

    var clearLineBreaksButton = dialogButtonRow.add("button", undefined, localize(LABELS.button.clearLineBreaks));
    clearLineBreaksButton.alignment = "left";

    var insertSoftBreakButton = dialogButtonRow.add("button", undefined, localize(LABELS.button.insertSoftBreak));
    insertSoftBreakButton.alignment = "left";

    /* 左右のボタンを分けるスペーサー / Spacer that separates the left and right button clusters */
    dialogButtonRow.add("statictext", undefined, "").preferredSize.width = BUTTON_ROW_SPACER_WIDTH;

    dialogButtonRow.add("button", undefined, localize(LABELS.button.cancel), { name: "cancel" });
    dialogButtonRow.add("button", undefined, localize(LABELS.button.ok), { name: "ok" });

    clearLineBreaksButton.onClick = function () {
        inputBox.text = inputBox.text
            .replace(/\\[nr]/g, "")   /* 文字列としての \n / \r を削除 / Remove literal \n and \r */
            .replace(/[\n\r]/g, "")   /* 実際の改行を削除 / Remove actual line breaks */
            .replace(/@#/g, "");      /* 可視マーカーを削除 / Remove the visible marker */
        inputBox.active = true;
        inputBox.selection = [inputBox.text.length, inputBox.text.length];
    };

    insertSoftBreakButton.onClick = function () {
        inputBox.text += SOFT_BREAK_MARKER;
        inputBox.selection = [inputBox.text.length, inputBox.text.length];
        inputBox.active = true;
    };

    inputBox.active = true;
    if (!initialText) inputBox.selection = [0, 0];

    return (textEditDialog.show() === 1) ? inputBox.text : null;
}

// =========================================
// テキスト挿入 / Text insertion
// =========================================

/**
 * 入力文字列を InDesign の改行コードへ正規化する
 * @param {string} rawText ダイアログで入力された文字列
 * @returns {string} 段落は \r、強制改行は \n に揃えた文字列
 */
function normalizeLineBreaks(rawText) {
    return rawText
        .replace(/\\n/g, "\n")
        .replace(/\\r/g, "\r")
        .replace(/\r\n/g, "\r")
        .replace(/\n/g, "\r")
        .replace(/\r{2,}/g, "\r")
        .replace(/@#/g, "\n");
}

/**
 * 選択がテキスト編集の対象になり得るか判定する
 * @returns {boolean} テキスト／テキストフレームが 1 つ選択されていれば true
 */
function hasEditableTextSelection() {
    return !!(app.selection && app.selection.length === 1 &&
        (app.selection[0] instanceof TextFrame ||
            app.selection[0].hasOwnProperty("contents") ||
            app.selection[0].hasOwnProperty("insertionPoints")));
}

/**
 * 選択範囲・挿入ポイント・テキストフレームのいずれかにテキストを流し込む
 * @param {string} textToApply 適用するテキスト
 * @returns {void}
 */
function replaceTextInSelection(textToApply) {
    try {
        if (!app.selection || app.selection.length !== 1) {
            alert(localize(LABELS.alert.noTextFrame));
            return;
        }

        var selectedObject = app.selection[0];

        /* 選択テキストがあれば置換 / Replace when text is selected */
        if (selectedObject.hasOwnProperty("contents") && selectedObject.contents !== "") {
            selectedObject.contents = textToApply;
            return;
        }

        /* テキストフレーム選択時はストーリー末尾へ挿入 / Append to the story when a text frame is selected */
        if (selectedObject instanceof TextFrame && selectedObject.parentStory) {
            selectedObject.parentStory.insertionPoints[-1].contents = textToApply;
            return;
        }

        /* 挿入ポイントのみのときはその位置へ挿入 / Insert at the caret when only an insertion point is active */
        if (selectedObject.hasOwnProperty("insertionPoints")) {
            selectedObject.insertionPoints[0].contents = textToApply;
            return;
        }

        alert(localize(LABELS.alert.noTextFrame));
    } catch (e) {
        alert(localize(LABELS.alert.errorOccurred) + e);
    }
}

/**
 * アクティブページの中央に新規テキストフレームを作成する
 * @param {string} textToApply 流し込むテキスト
 * @returns {void}
 */
function createTextFrameAtPageCenter(textToApply) {
    var activeDoc  = app.activeDocument;
    var frameWidth = Math.min(Math.max(textToApply.length * NEW_FRAME_WIDTH_PER_CHAR, NEW_FRAME_WIDTH_MIN), NEW_FRAME_WIDTH_MAX);

    /* bounds は [上, 左, 下, 右] / bounds is [top, left, bottom, right] */
    var pageBounds = app.activeWindow.activePage.bounds;
    var centerY = (pageBounds[0] + pageBounds[2]) / 2;
    var centerX = (pageBounds[1] + pageBounds[3]) / 2;

    var frameLeft = centerX - frameWidth / 2;
    var frameTop  = centerY - NEW_FRAME_HEIGHT / 2;

    var newTextFrame = activeDoc.textFrames.add();
    newTextFrame.geometricBounds = [frameTop, frameLeft, frameTop + NEW_FRAME_HEIGHT, frameLeft + frameWidth];
    newTextFrame.contents = textToApply;
}

// =========================================
// メイン処理 / Main
// =========================================

/**
 * ダイアログでテキストを編集し、選択状態に応じて反映する
 * @returns {void}
 */
function main() {
    /* 選択テキストを初期値にする（\n は可視マーカーへ）/ Seed the field with the selected text (\n shown as the marker) */
    var initialText = "";
    if (app.selection && app.selection.length === 1 && app.selection[0].hasOwnProperty("contents")) {
        initialText = app.selection[0].contents.replace(/\n/g, SOFT_BREAK_MARKER);
    }

    var userInput = showMultilineTextDialog(initialText);
    if (!userInput) return;

    var hasSelection = hasEditableTextSelection();
    var normalizedText = normalizeLineBreaks(userInput);

    app.doScript(function () {
        if (hasSelection) {
            replaceTextInSelection(normalizedText);
        } else {
            createTextFrameAtPageCenter(normalizedText);
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.FAST_ENTIRE_SCRIPT, localize(LABELS.undo.editText));
}

main();
