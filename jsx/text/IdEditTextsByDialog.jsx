#target indesign

/*

### 概要

複数行入力ダイアログでテキストを編集し、選択範囲の置換・カーソル位置への挿入・新規テキストフレーム作成を行います。

詳細は README を参照してください。

### Overview

Edits text in a multi-line dialog, then replaces the selection, inserts at the cursor, or creates a new text frame.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdEditTextsByDialog";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v0.1.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-25";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdEditTextsByDialog.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdEditTextsByDialog.md"; /* README (English) */

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

(function () {

    // =========================================
    // ラベル定義 / Labels
    // =========================================

    /**
     * UI 言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getUiLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getUiLang();

    /* {marker} は SOFT_BREAK_MARKER に置き換わる / {marker} is replaced with SOFT_BREAK_MARKER */
    var LABELS = {
        dialog: {
            title: { ja: "テキストを編集", en: "Edit Text" },
            note:  { ja: "{marker} で強制改行、fn + return で確定", en: "{marker} = forced line break, fn + Return to confirm" }
        },
        button: {
            ok:               { ja: "OK", en: "OK" },
            cancel:           { ja: "キャンセル", en: "Cancel" },
            removeLineBreaks: { ja: "改行を削除", en: "Remove Breaks" },
            appendMarker:     { ja: "{marker} を追加", en: "Add {marker}" }
        },
        tooltip: {
            textInput: {
                ja: "return で段落改行、{marker} で強制改行になります。",
                en: "Return starts a new paragraph; {marker} becomes a forced line break."
            },
            removeLineBreaks: {
                ja: "入力欄の改行と {marker} をすべて削除して 1 行にします。",
                en: "Removes every line break and {marker} to join the text into one line."
            },
            appendMarker: {
                ja: "入力欄の末尾に強制改行のマーカー（{marker}）を追加します。",
                en: "Appends the forced line break marker ({marker}) to the end of the field."
            }
        },
        alert: {
            noDocument:    { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noTextTarget:  { ja: "テキストまたはテキストフレームを選択してください。", en: "Please select text or a text frame." },
            errorOccurred: { ja: "エラーが発生しました：\n", en: "An error occurred:\n" }
        },
        undo: {
            editText: { ja: "テキストを編集", en: "Edit Text" }
        }
    };

    /**
     * ラベルを現在の言語で取得する（{marker} は SOFT_BREAK_MARKER に置換）
     * @param {object} labelEntry ja / en を持つラベルオブジェクト
     * @returns {string} 現在の言語のラベル文字列
     */
    function getLabel(labelEntry) {
        return labelEntry[uiLang].replace(/\{marker\}/g, SOFT_BREAK_MARKER);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 入力欄にフォーカスし、キャレットを末尾へ置く
     * @param {EditText} textInput 対象の入力欄
     * @returns {void}
     */
    function placeCaretAtEnd(textInput) {
        textInput.active = true;
        textInput.selection = [textInput.text.length, textInput.text.length];
    }

    /**
     * 複数行テキストの編集ダイアログを表示する
     * @param {string} initialText 入力欄の初期値
     * @returns {string|null} 入力されたテキスト。キャンセル時は null
     */
    function showMultilineTextDialog(initialText) {
        var editTextDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(editTextDialog, 8);

        var textInput = editTextDialog.add("edittext", undefined, initialText || "", { multiline: true });
        textInput.preferredSize = INPUT_BOX_SIZE;
        textInput.helpTip = getLabel(LABELS.tooltip.textInput);

        var noteRow = editTextDialog.add("group");
        setupRow(noteRow, "center", 0);
        noteRow.add("statictext", undefined, getLabel(LABELS.dialog.note));

        /* ボタンエリア（左右分割）/ Button area (split left and right) */
        var btnRowGroup = editTextDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var btnRemoveLineBreaks = btnLeftGroup.add("button", undefined, getLabel(LABELS.button.removeLineBreaks));
        btnRemoveLineBreaks.helpTip = getLabel(LABELS.tooltip.removeLineBreaks);
        var btnAppendMarker = btnLeftGroup.add("button", undefined, getLabel(LABELS.button.appendMarker));
        btnAppendMarker.helpTip = getLabel(LABELS.tooltip.appendMarker);

        /* スペーサー（伸縮）/ Spacer (stretchable) */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        btnRemoveLineBreaks.onClick = function () {
            textInput.text = textInput.text
                .replace(/\\[nr]/g, "")   /* 文字列としての \n / \r を削除 / Remove literal \n and \r */
                .replace(/[\n\r]/g, "")   /* 実際の改行を削除 / Remove actual line breaks */
                .split(SOFT_BREAK_MARKER).join("");
            placeCaretAtEnd(textInput);
        };

        btnAppendMarker.onClick = function () {
            textInput.text += SOFT_BREAK_MARKER;
            placeCaretAtEnd(textInput);
        };

        textInput.active = true;
        if (!initialText) textInput.selection = [0, 0];

        return (editTextDialog.show() === 1) ? textInput.text : null;
    }

    // =========================================
    // テキスト反映 / Text application
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
            .split(SOFT_BREAK_MARKER).join("\n");
    }

    /**
     * 選択が 1 つだけならそれを返す
     * @returns {object|null} 選択オブジェクト。0 個・複数なら null
     */
    function getSingleSelection() {
        return (app.selection && app.selection.length === 1) ? app.selection[0] : null;
    }

    /**
     * テキスト編集の対象になり得るか判定する
     * @param {object|null} selectedObject 選択オブジェクト
     * @returns {boolean} テキスト／挿入ポイント／テキストフレームなら true
     */
    function isEditableTextTarget(selectedObject) {
        return !!selectedObject &&
            (selectedObject instanceof TextFrame ||
                selectedObject.hasOwnProperty("contents") ||
                selectedObject.hasOwnProperty("insertionPoints"));
    }

    /**
     * 選択範囲・挿入ポイント・テキストフレームのいずれかにテキストを流し込む
     * @param {object} selectedObject 選択オブジェクト
     * @param {string} textToApply 適用するテキスト
     * @returns {void}
     */
    function applyTextToSelection(selectedObject, textToApply) {
        /* 選択テキスト、または文字のあるテキストフレームは中身を置換 / Replace selected text, or the contents of a non-empty text frame */
        if (selectedObject.hasOwnProperty("contents") && selectedObject.contents !== "") {
            selectedObject.contents = textToApply;
            return;
        }

        /* 空のテキストフレームはストーリー末尾へ挿入 / Append to the story of an empty text frame */
        if (selectedObject instanceof TextFrame && selectedObject.parentStory) {
            selectedObject.parentStory.insertionPoints[-1].contents = textToApply;
            return;
        }

        /* 挿入ポイントのみのときはその位置へ挿入 / Insert at the caret when only an insertion point is active */
        if (selectedObject.hasOwnProperty("insertionPoints")) {
            selectedObject.insertionPoints[0].contents = textToApply;
            return;
        }

        alert(getLabel(LABELS.alert.noTextTarget));
    }

    /**
     * アクティブページの中央に新規テキストフレームを作成する
     * @param {string} textToApply 流し込むテキスト
     * @returns {void}
     */
    function createTextFrameAtPageCenter(textToApply) {
        var frameWidth = Math.min(Math.max(textToApply.length * NEW_FRAME_WIDTH_PER_CHAR, NEW_FRAME_WIDTH_MIN), NEW_FRAME_WIDTH_MAX);

        /* bounds は [上, 左, 下, 右] / bounds is [top, left, bottom, right] */
        var pageBounds = app.activeWindow.activePage.bounds;
        var frameLeft = (pageBounds[1] + pageBounds[3]) / 2 - frameWidth / 2;
        var frameTop  = (pageBounds[0] + pageBounds[2]) / 2 - NEW_FRAME_HEIGHT / 2;

        var newTextFrame = app.activeDocument.textFrames.add();
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
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var selectedObject = getSingleSelection();

        /* 選択テキストを初期値にする（\n は可視マーカーへ）/ Seed the field with the selected text (\n shown as the marker) */
        var initialText = "";
        if (selectedObject && selectedObject.hasOwnProperty("contents") && typeof selectedObject.contents === "string") {
            initialText = selectedObject.contents.replace(/\n/g, SOFT_BREAK_MARKER);
        }

        var userInput = showMultilineTextDialog(initialText);
        if (!userInput) return;

        var normalizedText = normalizeLineBreaks(userInput);
        var hasTextTarget = isEditableTextTarget(selectedObject);

        /* ロックされたフレームなどで DOM への代入が失敗しうる / DOM writes can fail, e.g. on a locked frame */
        try {
            app.doScript(function () {
                if (hasTextTarget) {
                    applyTextToSelection(selectedObject, normalizedText);
                } else {
                    createTextFrameAtPageCenter(normalizedText);
                }
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undo.editText));
        } catch (e) {
            alert(getLabel(LABELS.alert.errorOccurred) + e);
        }
    }

    main();

})();
