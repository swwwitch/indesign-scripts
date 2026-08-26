#target indesign

/*
 * IdSplitParagraph.jsx
 *
 * 選択したテキストフレーム内の各段落を、元の位置と幅を保ったまま独立したテキストフレームへ分割します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdSplitParagraph";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-06-30";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSplitParagraph.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSplitParagraph.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* フレームを広げる 1 回あたりの量（pt）/ Height growth per step (pt) */
var FRAME_GROW_STEP_PT = 12;

/* オーバーセット解消の最大試行回数 / Max iterations to resolve overset */
var MAX_GROW_ITERATIONS = 200;

/* 幅復元後の再改行を救済する最大試行回数 / Max iterations after the width is restored */
var MAX_REFLOW_ITERATIONS = 20;

// ==============================
// UIレイアウトの共通設定 / Shared UI layout
// ==============================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
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
 * パネルの共通設定を適用する
 * @param {Panel} panel 対象パネル
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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
        overflowTitle: { ja: "オーバーセットテキストの確認", en: "Overset Text Detected" }
    },
    message: {
        overflowPrompt: {
            ja: "オーバーセットテキストがあります。処理方法を選択してください。",
            en: "The text frame contains overset text. Choose how to proceed."
        }
    },
    panel: {
        overflowOption: { ja: "処理方法", en: "How to proceed" }
    },
    radio: {
        expandFrame:   { ja: "フレームを拡張して解消する", en: "Expand frame to resolve" },
        ignoreOverset: {
            ja: "そのまま実行する（あふれたテキストは失われます）",
            en: "Run anyway (overset text will be lost)"
        }
    },
    tooltip: {
        expandFrame: {
            ja: "各段落を分割する前にフレームの高さを広げ、隠れているテキストをすべて表示してから処理します。",
            en: "Grows the frame height to reveal all hidden text before splitting each paragraph."
        },
        ignoreOverset: {
            ja: "現在表示されている段落だけを分割します。あふれて隠れているテキストは出力されません。",
            en: "Splits only the currently visible paragraphs. Hidden overset text will not be output."
        }
    },
    button: {
        ok:     { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    error: {
        docRequired:  { ja: "ドキュメントを開いてください。", en: "Please open a document." },
        selectFrame:  {
            ja: "テキストフレームを1つだけ選択して実行してください。",
            en: "Select exactly one text frame before running."
        },
        overflowFail: {
            ja: "オーバーセットテキストを解消できなかったため中断しました。",
            en: "Could not resolve overset text. Process cancelled."
        }
    },
    undo: {
        splitParagraphs: { ja: "段落ごとにフレーム分割", en: "Split Paragraphs Into Frames" }
    }
};

/**
 * ドット区切りキーでラベルを取得する
 * @param {string} labelKey 例: "dialog.overflowTitle"
 * @returns {string} 現在の言語のラベル文字列
 */
function getLabel(labelKey) {
    var keyParts = labelKey.split(".");
    var node = LABELS;
    for (var i = 0; i < keyParts.length; i++) {
        node = node[keyParts[i]];
    }
    return node[currentLanguage];
}

// =========================================
// 前提チェック / Preconditions
// =========================================
if (app.documents.length === 0) {
    alert(getLabel("error.docRequired"));
    return;
}

if (app.selection.length !== 1) {
    alert(getLabel("error.selectFrame"));
    return;
}

var sourceFrame = app.selection[0];
if (!sourceFrame || sourceFrame.constructor.name !== "TextFrame") {
    alert(getLabel("error.selectFrame"));
    return;
}

// =========================================
// 元フレーム情報 / Source frame info
// =========================================
var parentContainer = sourceFrame.parent;

/* 元フレームの座標 [上, 左, 下, 右] / Original frame bounds [top, left, bottom, right] */
var sourceBounds = sourceFrame.geometricBounds;
var sourceLeft   = sourceBounds[1];
var sourceRight  = sourceBounds[3];

// =========================================
// 補助関数 / Helpers
// =========================================

/**
 * オーバーセットが解消するまでフレーム高さを少しずつ広げる
 * @param {TextFrame} targetFrame 対象のテキストフレーム
 * @param {number} maxIterations 最大試行回数
 * @returns {void}
 */
function growFrameUntilFits(targetFrame, maxIterations) {
    var iterationCount = 0;
    while (targetFrame.overflows && iterationCount < maxIterations) {
        var bounds = targetFrame.geometricBounds;
        targetFrame.geometricBounds = [bounds[0], bounds[1], bounds[2] + FRAME_GROW_STEP_PT, bounds[3]];
        iterationCount++;
    }
}

// =========================================
// ダイアログ / Dialog
// =========================================

/**
 * オーバーセットテキストの処理方法をダイアログで確認する
 * @returns {string} "expand" / "ignore" / "cancel"
 */
function askOverflowHandling() {
    var overflowDialog = new Window("dialog", getLabel("dialog.overflowTitle") + " " + SCRIPT_VERSION);
    setupWindow(overflowDialog, 10);

    overflowDialog.add("statictext", undefined, getLabel("message.overflowPrompt"));

    /* 処理方法の選択パネル / Panel for choosing the handling method */
    var overflowOptionPanel = overflowDialog.add("panel", undefined, getLabel("panel.overflowOption"));
    setupPanel(overflowOptionPanel, 6);
    overflowOptionPanel.alignChildren = ["left", "top"];

    var expandFrameRadio   = overflowOptionPanel.add("radiobutton", undefined, getLabel("radio.expandFrame"));
    var ignoreOversetRadio = overflowOptionPanel.add("radiobutton", undefined, getLabel("radio.ignoreOverset"));
    expandFrameRadio.helpTip   = getLabel("tooltip.expandFrame");
    ignoreOversetRadio.helpTip = getLabel("tooltip.ignoreOverset");
    expandFrameRadio.value = true;

    /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
    var dialogButtonRow = overflowDialog.add("group");
    setupRow(dialogButtonRow, "right", 8);
    dialogButtonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    dialogButtonRow.add("button", undefined, getLabel("button.ok"), { name: "ok" });

    if (overflowDialog.show() !== 1) return "cancel";
    return expandFrameRadio.value ? "expand" : "ignore";
}

// =========================================
// 分割処理 / Split
// =========================================

/**
 * 1 段落を独立したテキストフレームに分割して配置する
 * @param {Paragraph} paragraph 分割元の段落
 * @returns {void}
 */
function createFrameForParagraph(paragraph) {
    /* 空行・行なしの段落はスキップ（仕様）/ Skip blank or line-less paragraphs (by design) */
    if (paragraph.contents.replace(/\s/g, "") === "" || paragraph.lines.length === 0) return;

    /* 元テキストの正確な Y 座標（ベースライン）/ Exact original Y (baseline) */
    var sourceBaselineY = paragraph.lines[0].baseline;

    /* 設定とサイズ・位置を引き継いだ新しいフレームを作成 / Create a new frame inheriting preferences, size, and position */
    var newFrame = parentContainer.textFrames.add();
    newFrame.textFramePreferences.properties = sourceFrame.textFramePreferences.properties;
    newFrame.geometricBounds = sourceBounds;

    paragraph.duplicate(LocationOptions.AT_BEGINNING, newFrame.insertionPoints.item(0));

    /* 余分な空段落や改行を削除 / Remove the extra trailing paragraph and return character */
    if (newFrame.paragraphs.length > 1) {
        newFrame.paragraphs.item(-1).remove();
    }
    if (newFrame.characters.length > 0 && newFrame.characters.item(-1).contents === "\r") {
        newFrame.characters.item(-1).remove();
    }

    /* いったんコンテンツに合わせる（高さと幅が縮む）/ Fit to content (height and width shrink) */
    newFrame.fit(FitOptions.FRAME_TO_CONTENT);

    /* 行が消えたフレームは破棄 / Discard the frame if it ends up with no lines */
    if (newFrame.lines.length === 0) {
        newFrame.remove();
        return;
    }

    /* Y はズレを補正し、X は元フレーム幅へ戻す / Correct Y by the offset, restore X to the original width */
    var dy = sourceBaselineY - newFrame.lines[0].baseline;
    var fittedBounds = newFrame.geometricBounds;
    newFrame.geometricBounds = [fittedBounds[0] + dy, sourceLeft, fittedBounds[2] + dy, sourceRight];

    /* 幅復元で再改行され、高さ不足になるケースを救済 / Rescue height shortage caused by reflow after the width is restored */
    growFrameUntilFits(newFrame, MAX_REFLOW_ITERATIONS);
}

/**
 * 選択フレーム内の全段落を独立したフレームへ分割する
 * @returns {string} "ok" または "overflowFail"
 */
function splitParagraphsIntoFrames() {
    /* 「拡張」選択時はオーバーセットを解消してから座標を再取得 / On "expand", clear overset then re-read bounds */
    if (overflowChoice === "expand") {
        growFrameUntilFits(sourceFrame, MAX_GROW_ITERATIONS);
        if (sourceFrame.overflows) return "overflowFail";
        sourceBounds = sourceFrame.geometricBounds;
        sourceLeft   = sourceBounds[1];
        sourceRight  = sourceBounds[3];
    }

    var paragraphList = sourceFrame.paragraphs.everyItem().getElements();
    for (var i = 0; i < paragraphList.length; i++) {
        createFrameForParagraph(paragraphList[i]);
    }

    sourceFrame.remove();
    return "ok";
}

// =========================================
// 実行 / Run
// =========================================

/* オーバーセットがあれば先に処理方法を確認（ダイアログは取り消し対象外）/ Ask first if overset (the dialog stays outside undo) */
var overflowChoice = "none";
if (sourceFrame.overflows) {
    overflowChoice = askOverflowHandling();
    if (overflowChoice === "cancel") return;
}

/* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
var resultStatus = app.doScript(
    splitParagraphsIntoFrames,
    ScriptLanguage.JAVASCRIPT,
    undefined,
    UndoModes.ENTIRE_SCRIPT,
    getLabel("undo.splitParagraphs")
);

if (resultStatus === "overflowFail") {
    alert(getLabel("error.overflowFail"));
}

})();
