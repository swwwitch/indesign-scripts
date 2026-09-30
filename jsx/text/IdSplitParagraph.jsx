#target indesign

/*

### 概要

選択したテキストフレーム内の各段落を、元の位置と幅（縦組みは高さ）を保ったまま独立したテキストフレームへ分割します。
縦組みにも対応しています。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSplitParagraph.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n8793ea71526b

### Overview

Splits each paragraph in the selected text frame into its own text frame, keeping the original position and width (height for vertical text).
Vertical text frames are supported as well.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSplitParagraph.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdSplitParagraph";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSplitParagraph.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSplitParagraph.md"; /* README (English) */
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

// UIレイアウト（再利用パーツ） / UI layout (reusable)

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

/**
 * ウィンドウの共通設定
 * @param {Window} targetWindow - 対象のウィンドウ
 * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
 * @returns {void}
 */
function setupWindow(targetWindow, spacing) {
    targetWindow.orientation = "column";
    targetWindow.alignChildren = "fill";
    targetWindow.margins = WINDOW_MARGINS;
    targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
 * @param {Panel} targetPanel - 対象のパネル
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
 * @returns {void}
 */
function setupPanel(targetPanel, spacing) {
    targetPanel.orientation = "column";
    targetPanel.alignChildren = ["fill", "top"];
    targetPanel.alignment = "fill";
    targetPanel.margins = PANEL_MARGINS;
    targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * タブの共通設定
 * @param {Tab} targetTab - 対象のタブ
 * @param {number} [spacing] - 要素間隔（省略時は変えない）
 * @returns {void}
 */
function setupTab(targetTab, spacing) {
    targetTab.orientation = "column";
    targetTab.alignChildren = "fill";
    targetTab.margins = TAB_MARGINS;
    if (typeof spacing === "number") targetTab.spacing = spacing;
}

/**
 * 横並びの行グループの共通設定（ボタン列など）。
 * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
 * @param {Group} rowGroup - 対象のグループ
 * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
 * @returns {void}
 */
function setupRow(rowGroup, rowAlignment, spacing) {
    rowGroup.orientation = "row";
    rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
    rowGroup.alignChildren = ["left", "center"];
    rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
 * @param {Button} targetButton - 対象のボタン
 * @param {number} trimPixels - 詰める量（px）
 * @returns {void}
 */
function trimButtonHeight(targetButton, trimPixels) {
    /* レイアウト前は size が無い / size is not set until the layout runs */
    if (!targetButton.size) return;
    targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
}

// UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

// ボタン行（再利用パーツ） / Button row (reusable)

var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

/**
 * ダイアログ下部のボタン行を作る。
 * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
 * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
 * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
 * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
 */
function addButtonRow(parent, rowOptions) {
    var isCentered = !!(rowOptions && rowOptions.centered);
    var btnRowGroup = parent.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
    btnRowGroup.spacing = BUTTON_ROW_SPACING;

    if (isCentered) {
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
    }

    btnRowGroup.alignment = ["fill", "bottom"];

    var btnLeftGroup = btnRowGroup.add("group");
    btnLeftGroup.alignChildren = ["left", "center"];
    btnLeftGroup.spacing = BUTTON_ROW_SPACING;

    /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
    var spacer = btnRowGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 0;

    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.alignChildren = ["right", "center"];
    btnRightGroup.spacing = BUTTON_ROW_SPACING;

    return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
}

/**
 * 左のグループにボタンが無い（右のボタンだけの）とき、行を左右中央に並べ直す。
 * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
 * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
 * @returns {void}
 */
function centerButtonRowIfRightOnly(buttonRow) {
    if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
    var btnRowGroup = buttonRow.rowGroup;
    /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
    btnRowGroup.remove(buttonRow.leftGroup);
    btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
    btnRowGroup.alignment = ["center", "bottom"];
    btnRowGroup.alignChildren = ["center", "center"];
    buttonRow.leftGroup = null;
}

// ボタン行（再利用パーツ）ここまで / End of the reusable button row

// =========================================
// ラベル定義 / Labels
// =========================================

// ローカライズ（再利用パーツ） / Localization (reusable)

/**
 * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
 * @returns {string} "ja" または "en"
 */
function getCurrentLang() {
    return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
}

var uiLang = getCurrentLang();

/**
 * LABELS から今の UI 言語の文言を取り出す。
 * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
 * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
 * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
 */
function getLabel(labelRef, placeholderValues) {
    var labelEntry = labelRef;
    if (typeof labelRef === "string") {
        var labelPathKeys = labelRef.split(".");
        labelEntry = LABELS;
        for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
            labelEntry = labelEntry[labelPathKeys[i]];
        }
    }
    var labelString;
    if (typeof labelEntry === "string") labelString = labelEntry;
    else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
    else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
    else return (typeof labelRef === "string") ? labelRef : "";
    return fillLabelPlaceholders(String(labelString), placeholderValues);
}

/**
 * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
 * @param {string|Object} labelRef - getLabel と同じ
 * @param {Object|Array} [placeholderValues] - getLabel と同じ
 * @returns {string} コロン付きの文言
 */
function labelText(labelRef, placeholderValues) {
    return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
}

/**
 * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
 * @param {string|Object} labelRef - getLabel と同じ
 * @param {string|number} value - コロンのあとに続ける値
 * @returns {string} 項目名と値をつないだ文字列
 */
function labelValueText(labelRef, value) {
    return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
}

/**
 * 文言の {name} や %1 に値を差し込む
 * @param {string} labelString - 文言
 * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
 * @returns {string} 差し込んだ文言
 */
function fillLabelPlaceholders(labelString, placeholderValues) {
    if (placeholderValues == null) return labelString;
    if (placeholderValues instanceof Array) {
        /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
        for (var i = placeholderValues.length; i >= 1; i--) {
            labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
        }
        return labelString;
    }
    for (var placeholderKey in placeholderValues) {
        if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
        labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
    }
    return labelString;
}

// ローカライズ（再利用パーツ）ここまで / End of the reusable localization

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

    /* ボタン行 / Button row */
    var buttonRow = addButtonRow(oversetDialog);
    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    centerButtonRowIfRightOnly(buttonRow);

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
