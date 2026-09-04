#target indesign

/*

### 概要

選択したテキストフレーム内の各段落を、元の位置と幅（縦組みは高さ）を保ったまま独立したテキストフレームへ分割します。
縦組みにも対応しています。

詳細は README を参照してください。

### Overview

Splits each paragraph in the selected text frame into its own text frame, keeping the original position and width (height for vertical text).
Vertical text frames are supported as well.

See the README for details.

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
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n8793ea71526b"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* フレームを広げる 1 回あたりの量（横組みは下、縦組みは左）/ Growth per step (downward for horizontal, leftward for vertical) */
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
        oversetTitle: { ja: "オーバーセットテキストの確認", en: "Overset Text Detected" }
    },
    message: {
        oversetPrompt: {
            ja: "オーバーセットテキストがあります。処理方法を選択してください。",
            en: "The text frame contains overset text. Choose how to proceed."
        }
    },
    panel: {
        oversetHandling: { ja: "処理方法", en: "How to proceed" }
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
        oversetFail: {
            ja: "オーバーセットテキストを解消できなかったため中断しました。",
            en: "Could not resolve overset text. Process cancelled."
        },
        threadedFrame: {
            ja: "連結されたテキストフレームには対応していません。連結を解除してから実行してください。",
            en: "Threaded text frames are not supported. Unthread the frame before running."
        },
        noParagraph: {
            ja: "分割できる段落がないため、何も変更していません。",
            en: "No paragraph to split. Nothing was changed."
        },
        unsupportedParent: {
            ja: "アンカー付きフレームや、ほかのオブジェクトの内側にあるフレームには対応していません。",
            en: "Anchored frames and frames nested inside another object are not supported."
        }
    },
    undo: {
        splitParagraphs: { ja: "段落ごとにフレーム分割", en: "Split Paragraphs Into Frames" }
    }
};

/**
 * ドット区切りキーでラベルを取得する
 * @param {string} labelKey 例: "dialog.oversetTitle"
 * @returns {string} 現在の言語のラベル文字列
 */
function getLabel(labelKey) {
    var keyParts = labelKey.split(".");
    var labelNode = LABELS;
    for (var i = 0; i < keyParts.length; i++) {
        labelNode = labelNode[keyParts[i]];
    }
    return labelNode[currentLanguage];
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

/* 連結フレームは元テキストが残るため対象外 / Threaded frames are excluded: the original text would survive in the thread */
if (sourceFrame.parentStory.textContainers.length > 1) {
    alert(getLabel("error.threadedFrame"));
    return;
}

/* 分割後のフレームを追加できる親か（アンカー付きフレームの親は Character で追加できない）
   / Parents that can hold the new frames (an anchored frame's parent is a Character and cannot) */
var FRAME_CONTAINER_TYPES = { Spread: true, MasterSpread: true, Page: true, Layer: true, Group: true };
if (!FRAME_CONTAINER_TYPES[String(sourceFrame.parent.constructor.name)]) {
    alert(getLabel("error.unsupportedParent"));
    return;
}

// =========================================
// 元フレーム情報 / Source frame info
// =========================================
var sourceParent = sourceFrame.parent;

/* 縦組みかどうか / Whether the story runs vertically */
var isVerticalStory = (sourceFrame.parentStory.storyPreferences.storyOrientation === StoryHorizontalOrVertical.VERTICAL);

/* 元フレームの座標 [上, 左, 下, 右] / Original frame bounds [top, left, bottom, right] */
var sourceBounds = sourceFrame.geometricBounds;

// =========================================
// 補助関数 / Helpers
// =========================================

/**
 * オーバーセットが解消するまでフレームを少しずつ広げる（横組みは下、縦組みは左）
 * @param {TextFrame} targetFrame 対象のテキストフレーム
 * @param {number} maxIterations 最大試行回数
 * @returns {void}
 */
function growFrameUntilFits(targetFrame, maxIterations) {
    var iterationCount = 0;
    while (targetFrame.overflows && iterationCount < maxIterations) {
        var currentBounds = targetFrame.geometricBounds;
        /* 横組みは下辺を、縦組みは左辺を伸ばす / Extend the bottom edge for horizontal text, the left edge for vertical */
        targetFrame.geometricBounds = isVerticalStory
            ? [currentBounds[0], currentBounds[1] - FRAME_GROW_STEP_PT, currentBounds[2], currentBounds[3]]
            : [currentBounds[0], currentBounds[1], currentBounds[2] + FRAME_GROW_STEP_PT, currentBounds[3]];
        iterationCount++;
    }
}

/**
 * 段落の複製で生じた末尾の空段落と改行を取り除く
 * @param {TextFrame} targetFrame 対象のテキストフレーム
 * @returns {void}
 */
function removeTrailingBreak(targetFrame) {
    if (targetFrame.paragraphs.length > 1) {
        targetFrame.paragraphs.item(-1).remove();
    }
    if (targetFrame.characters.length > 0 && targetFrame.characters.item(-1).contents === "\r") {
        targetFrame.characters.item(-1).remove();
    }
}

// =========================================
// ダイアログ / Dialog
// =========================================

/**
 * オーバーセットテキストの処理方法をダイアログで確認する
 * @returns {string} "expand" / "ignore" / "cancel"
 */
function askOversetHandling() {
    var oversetDialog = new Window("dialog", getLabel("dialog.oversetTitle") + " " + SCRIPT_VERSION);
    setupWindow(oversetDialog, 10);

    oversetDialog.add("statictext", undefined, getLabel("message.oversetPrompt"));

    /* 処理方法の選択パネル / Panel for choosing the handling method */
    var oversetHandlingPanel = oversetDialog.add("panel", undefined, getLabel("panel.oversetHandling"));
    setupPanel(oversetHandlingPanel, 6);
    oversetHandlingPanel.alignChildren = ["left", "top"];

    var expandFrameRadio   = oversetHandlingPanel.add("radiobutton", undefined, getLabel("radio.expandFrame"));
    var ignoreOversetRadio = oversetHandlingPanel.add("radiobutton", undefined, getLabel("radio.ignoreOverset"));
    expandFrameRadio.helpTip   = getLabel("tooltip.expandFrame");
    ignoreOversetRadio.helpTip = getLabel("tooltip.ignoreOverset");
    expandFrameRadio.value = true;

    /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
    var btnRowGroup = oversetDialog.add("group");
    setupRow(btnRowGroup, "right", 8);
    btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

    if (oversetDialog.show() !== 1) return "cancel";
    return expandFrameRadio.value ? "expand" : "ignore";
}

// =========================================
// 分割処理 / Split
// =========================================

/**
 * 1 段落を独立したテキストフレームに分割して配置する
 * @param {Paragraph} paragraph 分割元の段落
 * @returns {boolean} フレームを作成したら true
 */
function createFrameForParagraph(paragraph) {
    /* 空行・行なしの段落はスキップ（仕様）/ Skip blank or line-less paragraphs (by design) */
    if (paragraph.contents.replace(/\s/g, "") === "" || paragraph.lines.length === 0) return false;

    /* 元テキストの正確なベースライン位置（横組みは Y、縦組みは X）/ Exact original baseline position (Y for horizontal, X for vertical) */
    var sourceBaseline = paragraph.lines[0].baseline;

    /* 設定とサイズ・位置を引き継いだ新しいフレームを作成 / Create a new frame inheriting preferences, size, and position */
    var paragraphFrame = sourceParent.textFrames.add();
    paragraphFrame.textFramePreferences.properties = sourceFrame.textFramePreferences.properties;
    paragraphFrame.geometricBounds = sourceBounds;

    paragraph.duplicate(LocationOptions.AT_BEGINNING, paragraphFrame.insertionPoints.item(0));
    removeTrailingBreak(paragraphFrame);

    /* いったんコンテンツに合わせる（高さと幅が縮む）/ Fit to content (height and width shrink) */
    paragraphFrame.fit(FitOptions.FRAME_TO_CONTENT);

    /* 行が消えたフレームは破棄 / Discard the frame if it ends up with no lines */
    if (paragraphFrame.lines.length === 0) {
        paragraphFrame.remove();
        return false;
    }

    /* 流れ方向はベースラインのズレを補正し、直交方向は元フレームの位置へ戻す
       / Correct the flow axis by the baseline offset, restore the cross axis to the source frame */
    var baselineDelta = sourceBaseline - paragraphFrame.lines[0].baseline;
    var fittedBounds = paragraphFrame.geometricBounds;
    paragraphFrame.geometricBounds = isVerticalStory
        ? [sourceBounds[0], fittedBounds[1] + baselineDelta, sourceBounds[2], fittedBounds[3] + baselineDelta]
        : [fittedBounds[0] + baselineDelta, sourceBounds[1], fittedBounds[2] + baselineDelta, sourceBounds[3]];

    /* 幅（縦組みは高さ）を戻した際の再改行で不足するケースを救済 / Rescue the shortage caused by reflow after the cross axis is restored */
    growFrameUntilFits(paragraphFrame, MAX_REFLOW_ITERATIONS);
    return true;
}

/**
 * 選択フレーム内の全段落を独立したフレームへ分割する
 * @returns {string} "ok" / "oversetFail" / "noParagraph"
 */
function splitParagraphsIntoFrames() {
    /* 中断時に元フレームを戻すための座標 / Bounds used to restore the source frame when aborting */
    var originalBounds = sourceBounds;

    /* 「拡張」選択時はオーバーセットを解消してから座標を再取得 / On "expand", clear overset then re-read bounds */
    if (oversetChoice === "expand") {
        growFrameUntilFits(sourceFrame, MAX_GROW_ITERATIONS);
        if (sourceFrame.overflows) {
            sourceFrame.geometricBounds = originalBounds;
            return "oversetFail";
        }
        sourceBounds = sourceFrame.geometricBounds;
    }

    var paragraphList = sourceFrame.paragraphs.everyItem().getElements();
    var createdCount = 0;
    for (var i = 0; i < paragraphList.length; i++) {
        if (createFrameForParagraph(paragraphList[i])) createdCount++;
    }

    /* 1 つも作れなかったときは元フレームを残す / Keep the source frame when nothing was created */
    if (createdCount === 0) {
        sourceFrame.geometricBounds = originalBounds;
        return "noParagraph";
    }

    sourceFrame.remove();
    return "ok";
}

// =========================================
// 実行 / Run
// =========================================

/* オーバーセットがあれば先に処理方法を確認（ダイアログは取り消し対象外）/ Ask first if overset (the dialog stays outside undo) */
var oversetChoice = "none";
if (sourceFrame.overflows) {
    oversetChoice = askOversetHandling();
    if (oversetChoice === "cancel") return;
}

/* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
var splitResult = app.doScript(
    splitParagraphsIntoFrames,
    ScriptLanguage.JAVASCRIPT,
    undefined,
    UndoModes.ENTIRE_SCRIPT,
    getLabel("undo.splitParagraphs")
);

if (splitResult === "oversetFail") {
    alert(getLabel("error.oversetFail"));
} else if (splitResult === "noParagraph") {
    alert(getLabel("error.noParagraph"));
}

})();
