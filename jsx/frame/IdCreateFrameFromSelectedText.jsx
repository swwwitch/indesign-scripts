#target indesign

/*

### 概要

選択したテキストのサイズを基準に、インライン（アンカー付き）またはページ上へグラフィックフレームを作成します。

詳細は README を参照してください。

### Overview

Creates a graphic frame sized from the selected text, either inline (anchored) or placed on the page.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdCreateFrameFromSelectedText"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.6";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-03-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCreateFrameFromSelectedText.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCreateFrameFromSelectedText.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ndd1b7c5246a3"; /* 紹介記事 / article URL */

// Original idea
// DTP Script note
// https://note.com/yosi2631/n/ned2dbc1cb79d

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 既定で選ばれる段落スタイル名の手がかり / Hints used to preselect a paragraph style */
var DEFAULT_PARA_STYLE_NAME = "p.img";
var DEFAULT_PARA_STYLE_HINT = ".img";

/* 「なし」オブジェクトスタイルの名前（日英）/ Names of the "None" object style (JA/EN) */
var NONE_OBJECT_STYLE_NAMES = ["[なし]", "[None]"];

/* 高さ入力の初期値 / Default values of the height field */
var DEFAULT_HEIGHT_IN_LINES = "1";
var DEFAULT_HEIGHT_IN_MM    = "40";

/* 行送りを取得できないときのフォールバック（pt）/ Fallback used when leading cannot be read (pt) */
var FALLBACK_POINT_SIZE = 14;
var FALLBACK_LEADING    = 14;

/* mm をポイントへ換算する係数 / Millimeter-to-point conversion factor */
var MM_TO_POINTS = 2.834645669;

// =========================================
// レイアウト設定 / Layout settings
// =========================================

/* 高さ入力欄の文字数 / Character width of the height field */
var HEIGHT_INPUT_CHARACTERS = 8;

// ==============================
// UIレイアウトの共通設定 / Shared UI layout
// ==============================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

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

var currentLang = getCurrentLang();

var LABELS = {
    dialog: {
        title:            { ja: "選択した文字列からフレーム作成", en: "Create Frame from Selected Text" },
        titleNoSelection: { ja: "グラフィックフレーム作成", en: "Create Graphic Frame" }
    },
    panel: {
        method:         { ja: "配置方法", en: "Placement" },
        frameSize:      { ja: "フレームサイズ", en: "Frame Size" },
        width:          { ja: "幅", en: "Width" },
        height:         { ja: "高さ", en: "Height" },
        textSettings:   { ja: "テキスト設定", en: "Text Settings" },
        objectSettings: { ja: "オブジェクト設定", en: "Object Settings" }
    },
    radio: {
        graphicFrame: { ja: "グラフィックフレーム", en: "Graphic Frame" },
        inlineFrame:  { ja: "インライン（アンカー付き）", en: "Inline (Anchored)" },
        widthText:    { ja: "選択したテキスト", en: "Selected Text" },
        widthColumn:  { ja: "カラム幅", en: "Column Width" },
        widthFrame:   { ja: "親フレーム", en: "Parent Frame" },
        widthMargin:  { ja: "ページのマージン", en: "Page Margins" },
        heightLines:  { ja: "行数", en: "Line Count" },
        heightSize:   { ja: "サイズ指定", en: "Specify Size" }
    },
    field: {
        paraStyle: { ja: "挿入行の段落スタイル:", en: "Paragraph Style for Inserted Line:" },
        objStyle:  { ja: "オブジェクトスタイル:", en: "Object Style:" },
        wrap:      { ja: "テキストの回り込み:", en: "Text Wrap:" }
    },
    checkbox: {
        autoLeading: { ja: "行送り：自動", en: "Leading: Auto" }
    },
    wrap: {
        none:        { ja: "なし", en: "None" },
        boundingBox: { ja: "境界線ボックスで回り込む", en: "Wrap Around Bounding Box" },
        contour:     { ja: "オブジェクトのシェイプで回り込む", en: "Wrap Around Object Shape" },
        jumpObject:  { ja: "オブジェクトを挟んで回り込む", en: "Jump Object" },
        nextColumn:  { ja: "次の段へテキストを送る", en: "Jump to Next Column" }
    },
    unit: {
        lines: { ja: "行", en: "lines" },
        mm:    { ja: "mm", en: "mm" }
    },
    button: {
        ok:     { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    tooltip: {
        heightLines: {
            ja: "行数モードの高さは概算です。1行目の文字サイズ＋残り行数×行送りで計算します。",
            en: "Line Count height is approximate. It is calculated as first-line point size plus leading for the remaining lines."
        }
    },
    alert: {
        noDocument:          { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        selectText:          { ja: "テキストを選択してください。", en: "Please select text." },
        selectTextBeforeRun: { ja: "テキスト項目を1つだけ選択してから実行してください。", en: "Select only one text item before running the script." },
        boundsError:         { ja: "選択したテキストの座標を取得できませんでした。", en: "Could not get the bounds of the selected text." },
        parentPageError:     { ja: "配置先のページを取得できませんでした。", en: "Could not determine the destination page." },
        parentFrameError:    { ja: "挿入位置の親テキストフレームを取得できませんでした。", en: "Could not get the parent text frame for the insertion point." }
    },
    undo: {
        createFrame: { ja: "フレーム作成", en: "Create Frame" }
    }
};

/**
 * ドット区切りキーでラベルを取得する
 * @param {string} labelKey 例: "dialog.title"
 * @returns {string} 現在の言語のラベル文字列
 */
function getLabel(labelKey) {
    var node = LABELS;
    var keyParts = labelKey.split(".");
    for (var i = 0; i < keyParts.length; i++) {
        node = node[keyParts[i]];
        if (!node) return labelKey;
    }
    return node[currentLang] || node.en || labelKey;
}

// =========================================
// テキスト判定と座標取得 / Text detection and bounds
// =========================================

/**
 * 選択オブジェクトがテキストかどうかを判定する
 * @param {object} selectionItem 選択オブジェクト
 * @returns {boolean} テキストなら true
 */
function isTextSelection(selectionItem) {
    if (!selectionItem) return false;

    try {
        if (selectionItem.hasOwnProperty("baseline")) return true;
    } catch (e) {}

    try {
        if (selectionItem.hasOwnProperty("characters") && selectionItem.characters && selectionItem.characters.length > 0) {
            return true;
        }
    } catch (e) {}

    try {
        if (selectionItem.constructor && selectionItem.constructor.name) {
            var typeName = String(selectionItem.constructor.name);
            if (typeName === "Text" || typeName === "Word" || typeName === "Line" ||
                typeName === "Paragraph" || typeName === "TextStyleRange" || typeName === "Character") {
                return true;
            }
        }
    } catch (e) {}

    return false;
}

/**
 * アウトライン化した一時オブジェクトから外接矩形を求め、そのオブジェクトを削除する
 * @param {Array} outlineItems アウトライン化で得られたオブジェクトの配列
 * @returns {Array<number>|null} [上, 左, 下, 右]。求められない場合は null
 */
function getAndRemoveOutlineBounds(outlineItems) {
    var top = null, left = null, bottom = null, right = null;

    try {
        for (var i = 0; i < outlineItems.length; i++) {
            var itemBounds = outlineItems[i].geometricBounds;
            if (top === null    || itemBounds[0] < top)    top    = itemBounds[0];
            if (left === null   || itemBounds[1] < left)   left   = itemBounds[1];
            if (bottom === null || itemBounds[2] > bottom) bottom = itemBounds[2];
            if (right === null  || itemBounds[3] > right)  right  = itemBounds[3];
        }
    } finally {
        for (var j = outlineItems.length - 1; j >= 0; j--) {
            try { outlineItems[j].remove(); } catch (e) {}
        }
    }

    if (top === null) return null;
    return [top, left, bottom, right];
}

/**
 * 1 文字分の外接矩形を求める
 * @param {Character} characterObject 対象の文字
 * @returns {Array<number>|null} [上, 左, 下, 右]。求められない場合は null
 */
function getCharacterOutlineBounds(characterObject) {
    try {
        var outlineItems = characterObject.createOutlines(false);
        if (!outlineItems || outlineItems.length === 0) return null;
        return getAndRemoveOutlineBounds(outlineItems);
    } catch (e) {
        return null;
    }
}

/**
 * 選択テキスト全体を一度にアウトライン化して外接矩形を求める
 * @param {object} textObject 対象のテキスト
 * @returns {Array<number>|null} [上, 左, 下, 右]。求められない場合は null
 */
function getTextSelectionBoundsSingleOutline(textObject) {
    try {
        if (!textObject || !textObject.characters || textObject.characters.length === 0) return null;
        var outlineItems = textObject.createOutlines(false);
        if (!outlineItems || outlineItems.length === 0) return null;
        return getAndRemoveOutlineBounds(outlineItems);
    } catch (e) {
        return null;
    }
}

/**
 * 1 文字ずつアウトライン化して外接矩形を合成する
 * @param {object} textObject 対象のテキスト
 * @returns {Array<number>|null} [上, 左, 下, 右]。求められない場合は null
 */
function getTextSelectionBoundsPerCharacter(textObject) {
    try {
        var characterList = textObject.characters;
        if (!characterList || characterList.length === 0) return null;

        var mergedTop = null, mergedLeft = null, mergedBottom = null, mergedRight = null;

        for (var i = 0; i < characterList.length; i++) {
            var character = characterList[i];

            try {
                /* 改行類はアウトライン化できないので除外 / Skip break characters, which cannot be outlined */
                if (character.contents === "\r" || character.contents === "\n" || character.contents === "\u0003") continue;
            } catch (e) {}

            var characterBounds = getCharacterOutlineBounds(character);
            if (!characterBounds) continue;

            if (mergedTop === null    || characterBounds[0] < mergedTop)    mergedTop    = characterBounds[0];
            if (mergedLeft === null   || characterBounds[1] < mergedLeft)   mergedLeft   = characterBounds[1];
            if (mergedBottom === null || characterBounds[2] > mergedBottom) mergedBottom = characterBounds[2];
            if (mergedRight === null  || characterBounds[3] > mergedRight)  mergedRight  = characterBounds[3];
        }

        if (mergedTop === null) {
            try {
                return textObject.parentTextFrames[0].geometricBounds;
            } catch (e) {
                return null;
            }
        }

        return [mergedTop, mergedLeft, mergedBottom, mergedRight];
    } catch (e) {
        return null;
    }
}

/**
 * 選択テキストの外接矩形を求める
 * @param {object} textObject 対象のテキスト
 * @returns {Array<number>|null} [上, 左, 下, 右]。求められない場合は null
 */
function getTextSelectionBounds(textObject) {
    return getTextSelectionBoundsSingleOutline(textObject) || getTextSelectionBoundsPerCharacter(textObject);
}

/**
 * 挿入ポイントが属するテキストフレームを取得する
 * @param {InsertionPoint} insertionPoint 対象の挿入ポイント
 * @returns {TextFrame|null} テキストフレーム。取得できない場合は null
 */
function getInsertionPointTextFrame(insertionPoint) {
    try {
        if (insertionPoint && insertionPoint.parentTextFrames && insertionPoint.parentTextFrames.length > 0) {
            return insertionPoint.parentTextFrames[0];
        }
    } catch (e) {}
    return null;
}

/**
 * テキストが配置されているページを取得する
 * @param {object} textObject 対象のテキスト
 * @returns {Page|null} ページ。取得できない場合は null
 */
function getParentPage(textObject) {
    try {
        if (textObject.parentTextFrames.length > 0) return textObject.parentTextFrames[0].parentPage;
    } catch (e) {}
    return null;
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {

    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var selectedText     = null;
    var isInsertionPoint = false;
    var hasNoSelection   = false;

    if (app.selection.length === 0) {
        hasNoSelection = true;
    } else if (app.selection.length !== 1) {
        alert(getLabel("alert.selectTextBeforeRun"));
        return;
    } else {
        selectedText = app.selection[0];
        if (selectedText.constructor && selectedText.constructor.name === "InsertionPoint") {
            isInsertionPoint = true;
        } else if (!isTextSelection(selectedText)) {
            alert(getLabel("alert.selectText"));
            return;
        }
    }

    var activeDoc = app.activeDocument;

    /* 選択テキストの外接矩形（挿入ポイント・未選択時は取得しない）/ Bounds of the selection (skipped for insertion points and empty selections) */
    var selectionBounds = null;
    var hasTextBounds   = false;
    if (!isInsertionPoint && !hasNoSelection) {
        selectionBounds = getTextSelectionBounds(selectedText);
        if (!selectionBounds) {
            alert(getLabel("alert.boundsError"));
            return;
        }
        hasTextBounds = true;
    }

    var selectionWidth  = hasTextBounds ? (selectionBounds[3] - selectionBounds[1]) : 0;
    var selectionHeight = hasTextBounds ? (selectionBounds[2] - selectionBounds[0]) : 0;

    /* オブジェクトスタイル一覧 / Object style list */
    var objectStyleNames = [];
    for (var objectStyleIndex = 0; objectStyleIndex < activeDoc.objectStyles.length; objectStyleIndex++) {
        objectStyleNames.push(activeDoc.objectStyles[objectStyleIndex].name);
    }

    /* 段落スタイル一覧と既定選択位置 / Paragraph style list and the default selection */
    var paragraphStyleNames = [];
    var defaultParaStyleIndex = 0;
    for (var paraStyleIndex = 0; paraStyleIndex < activeDoc.paragraphStyles.length; paraStyleIndex++) {
        var paragraphStyleName = activeDoc.paragraphStyles[paraStyleIndex].name;
        paragraphStyleNames.push(paragraphStyleName);
        if (paragraphStyleName === DEFAULT_PARA_STYLE_NAME || paragraphStyleName.indexOf(DEFAULT_PARA_STYLE_HINT) !== -1) {
            defaultParaStyleIndex = paraStyleIndex;
        }
    }

    /* 回り込みの選択肢 / Text wrap options */
    var wrapOptionLabels = [
        getLabel("wrap.none"),
        getLabel("wrap.boundingBox"),
        getLabel("wrap.contour"),
        getLabel("wrap.jumpObject"),
        getLabel("wrap.nextColumn")
    ];
    var wrapOptionModes = [
        TextWrapModes.NONE,
        TextWrapModes.BOUNDING_BOX_TEXT_WRAP,
        TextWrapModes.CONTOUR,
        TextWrapModes.JUMP_OBJECT_TEXT_WRAP,
        TextWrapModes.NEXT_COLUMN_TEXT_WRAP
    ];

    /* 「なし」オブジェクトスタイルの位置 / Index of the "None" object style */
    var noneObjectStyleIndex = 0;
    for (var i = 0; i < objectStyleNames.length; i++) {
        if (objectStyleNames[i] === NONE_OBJECT_STYLE_NAMES[0] || objectStyleNames[i] === NONE_OBJECT_STYLE_NAMES[1]) {
            noneObjectStyleIndex = i;
            break;
        }
    }

    // ---------------------------------------
    // ダイアログ / Dialog
    // ---------------------------------------
    var dialogTitle = (isInsertionPoint || hasNoSelection)
        ? getLabel("dialog.titleNoSelection")
        : getLabel("dialog.title");

    var createFrameDialog = new Window("dialog", dialogTitle + " " + SCRIPT_VERSION);
    setupWindow(createFrameDialog);

    /* 配置方法パネル / Placement panel */
    var placementPanel = createFrameDialog.add("panel", undefined, getLabel("panel.method"));
    setupPanel(placementPanel, 6);
    placementPanel.alignChildren = ["left", "top"];

    var graphicFrameRadio = placementPanel.add("radiobutton", undefined, getLabel("radio.graphicFrame"));
    var inlineFrameRadio  = placementPanel.add("radiobutton", undefined, getLabel("radio.inlineFrame"));
    if (isInsertionPoint || hasNoSelection) {
        graphicFrameRadio.value = true;
        /* 未選択時は挿入位置がないためインラインを無効化 / Inline needs an insertion point, so disable it when nothing is selected */
        if (hasNoSelection) inlineFrameRadio.enabled = false;
    } else {
        inlineFrameRadio.value = true;
    }

    /* 2カラムレイアウト / Two-column layout */
    var mainColumnsRow = createFrameDialog.add("group");
    setupRow(mainColumnsRow, "fill", COLUMN_SPACING);
    mainColumnsRow.alignChildren = ["fill", "top"];

    var leftColumn = mainColumnsRow.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = ["fill", "top"];
    leftColumn.spacing = PANEL_SPACING;

    var frameSizePanel = leftColumn.add("panel", undefined, getLabel("panel.frameSize"));
    setupPanel(frameSizePanel, 8);

    /* 幅パネル / Width panel */
    var widthPanel = frameSizePanel.add("panel", undefined, getLabel("panel.width"));
    setupPanel(widthPanel, 6);
    widthPanel.alignChildren = ["left", "top"];

    var widthFromTextRadio        = widthPanel.add("radiobutton", undefined, getLabel("radio.widthText"));
    var widthFromColumnRadio      = widthPanel.add("radiobutton", undefined, getLabel("radio.widthColumn"));
    var widthFromParentFrameRadio = widthPanel.add("radiobutton", undefined, getLabel("radio.widthFrame"));
    var widthFromMarginRadio      = widthPanel.add("radiobutton", undefined, getLabel("radio.widthMargin"));

    if (isInsertionPoint || hasNoSelection) {
        widthFromMarginRadio.value = true;
        widthFromTextRadio.enabled = false;
        if (hasNoSelection) {
            widthFromColumnRadio.enabled = false;
            widthFromParentFrameRadio.enabled = false;
        }
    } else {
        widthFromColumnRadio.value = true;
    }

    /* 高さパネル（テキスト外接矩形が使えないときのみ）/ Height panel (only when text bounds are unavailable) */
    var heightInput      = null;
    var heightLinesRadio = null;
    var heightSizeRadio  = null;
    if (!hasTextBounds) {
        var heightPanel = frameSizePanel.add("panel", undefined, getLabel("panel.height"));
        setupPanel(heightPanel, 6);
        heightPanel.alignChildren = ["left", "top"];

        heightLinesRadio = heightPanel.add("radiobutton", undefined, getLabel("radio.heightLines"));
        heightLinesRadio.helpTip = getLabel("tooltip.heightLines");
        heightSizeRadio = heightPanel.add("radiobutton", undefined, getLabel("radio.heightSize"));

        var heightInputRow = heightPanel.add("group");
        setupRow(heightInputRow, "left", 6);

        heightInput = heightInputRow.add("edittext", undefined, hasNoSelection ? DEFAULT_HEIGHT_IN_MM : DEFAULT_HEIGHT_IN_LINES);
        heightInput.characters = HEIGHT_INPUT_CHARACTERS;
        if (!hasNoSelection) heightInput.helpTip = getLabel("tooltip.heightLines");

        var heightUnitLabel = heightInputRow.add("statictext", undefined,
            hasNoSelection ? getLabel("unit.mm") : getLabel("unit.lines"));

        /* 未選択時は行送りを参照できないため行数モードを無効化 / Without a selection there is no leading to read, so disable line-count mode */
        if (hasNoSelection) {
            heightSizeRadio.value = true;
            heightLinesRadio.enabled = false;
        } else {
            heightLinesRadio.value = true;
        }

        heightLinesRadio.onClick = function () {
            heightInput.text = DEFAULT_HEIGHT_IN_LINES;
            heightUnitLabel.text = getLabel("unit.lines");
        };
        heightSizeRadio.onClick = function () {
            heightInput.text = DEFAULT_HEIGHT_IN_MM;
            heightUnitLabel.text = getLabel("unit.mm");
        };
    }

    /* 右カラム / Right column */
    var rightColumn = mainColumnsRow.add("group");
    rightColumn.orientation = "column";
    rightColumn.alignChildren = ["fill", "top"];
    rightColumn.spacing = PANEL_SPACING;

    var textSettingsPanel = rightColumn.add("panel", undefined, getLabel("panel.textSettings"));
    setupPanel(textSettingsPanel, 6);

    var paraStyleLabel    = textSettingsPanel.add("statictext", undefined, getLabel("field.paraStyle"));
    var paraStyleDropdown = textSettingsPanel.add("dropdownlist", undefined, paragraphStyleNames);
    paraStyleDropdown.selection = defaultParaStyleIndex;

    var autoLeadingCheckbox = textSettingsPanel.add("checkbox", undefined, getLabel("checkbox.autoLeading"));
    autoLeadingCheckbox.value = true;

    var objectSettingsPanel = rightColumn.add("panel", undefined, getLabel("panel.objectSettings"));
    setupPanel(objectSettingsPanel, 6);

    objectSettingsPanel.add("statictext", undefined, getLabel("field.objStyle"));
    var objectStyleDropdown = objectSettingsPanel.add("dropdownlist", undefined, objectStyleNames);
    objectStyleDropdown.selection = noneObjectStyleIndex;

    var wrapLabel    = objectSettingsPanel.add("statictext", undefined, getLabel("field.wrap"));
    var wrapDropdown = objectSettingsPanel.add("dropdownlist", undefined, wrapOptionLabels);
    wrapDropdown.selection = 0;

    /**
     * 現在の選択状態に合わせてコントロールの有効／無効を切り替える
     * @returns {void}
     */
    function updateDialogState() {
        var isInlinePlacement = inlineFrameRadio.value;
        var isNoneObjectStyle = (objectStyleDropdown.selection.index === noneObjectStyleIndex);

        /* グラフィックフレームでは挿入行の段落スタイルを使わない / The inserted-line paragraph style only applies to inline placement */
        paraStyleLabel.enabled      = isInlinePlacement;
        paraStyleDropdown.enabled   = isInlinePlacement;
        autoLeadingCheckbox.enabled = isInlinePlacement;

        var disableWidthFromText        = !hasTextBounds;
        var disableWidthFromColumn      = hasNoSelection;
        var disableWidthFromParentFrame = isInlinePlacement || hasNoSelection;
        var disableWidthFromMargin      = isInlinePlacement;

        widthFromTextRadio.enabled        = !disableWidthFromText;
        widthFromColumnRadio.enabled      = !disableWidthFromColumn;
        widthFromParentFrameRadio.enabled = !disableWidthFromParentFrame;
        widthFromMarginRadio.enabled      = !disableWidthFromMargin;

        /* 選択中の項目が無効になったら、有効な項目へ移す / Move the selection to an enabled option when the current one is disabled */
        if ((widthFromTextRadio.value && disableWidthFromText) ||
            (widthFromColumnRadio.value && disableWidthFromColumn) ||
            (widthFromParentFrameRadio.value && disableWidthFromParentFrame) ||
            (widthFromMarginRadio.value && disableWidthFromMargin)) {
            if (!disableWidthFromText) {
                widthFromTextRadio.value = true;
            } else if (!disableWidthFromColumn) {
                widthFromColumnRadio.value = true;
            } else if (!disableWidthFromParentFrame) {
                widthFromParentFrameRadio.value = true;
            } else if (!disableWidthFromMargin) {
                widthFromMarginRadio.value = true;
            }
        }

        /* 回り込みはページ配置かつオブジェクトスタイルが「なし」のときだけ有効 / Text wrap applies only to page placement with the "None" object style */
        wrapLabel.enabled    = !isInlinePlacement && isNoneObjectStyle;
        wrapDropdown.enabled = !isInlinePlacement && isNoneObjectStyle;
    }

    updateDialogState();
    graphicFrameRadio.onClick     = updateDialogState;
    inlineFrameRadio.onClick      = updateDialogState;
    objectStyleDropdown.onChange  = updateDialogState;

    /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
    var dialogButtonRow = createFrameDialog.add("group");
    setupRow(dialogButtonRow, "center", 8);
    dialogButtonRow.margins = [0, 10, 0, 0];
    dialogButtonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    dialogButtonRow.add("button", undefined, getLabel("button.ok"), { name: "ok" });

    if (createFrameDialog.show() !== 1) return;

    // ---------------------------------------
    // 入力値の確定 / Resolve settings
    // ---------------------------------------
    var useInlinePlacement    = inlineFrameRadio.value;
    var selectedObjStyleIndex = objectStyleDropdown.selection.index;
    var selectedWrapMode      = wrapOptionModes[wrapDropdown.selection.index];
    var selectedParaStyleIdx  = paraStyleDropdown.selection.index;
    var useAutoLeading        = autoLeadingCheckbox.value;

    /* フレームの高さ / Frame height */
    var frameHeight = selectionHeight;
    if (heightLinesRadio && heightSizeRadio) {
        if (heightLinesRadio.value) {
            /* 行数モード: 1行目の文字サイズ + 残り行数 × 行送り / Line-count mode: first-line size plus leading for the remaining lines */
            var lineCount = parseFloat(heightInput.text) || 1;
            if (lineCount < 1) lineCount = 1;
            try {
                var referenceFrame = isInsertionPoint
                    ? getInsertionPointTextFrame(selectedText)
                    : selectedText.parentTextFrames[0];
                if (!referenceFrame) throw new Error("No parent text frame");

                var referenceStory = referenceFrame.parentStory;
                var insertionIndex = isInsertionPoint ? selectedText.index : selectedText.characters[0].index;
                var referenceCharacter = (insertionIndex < referenceStory.characters.length)
                    ? referenceStory.characters[insertionIndex]
                    : referenceStory.characters[referenceStory.characters.length - 1];

                var referencePointSize = referenceCharacter.pointSize;
                var referenceLeading   = referenceCharacter.leading;
                if (referenceLeading === Leading.AUTO) {
                    referenceLeading = referencePointSize * (referenceCharacter.autoLeading / 100);
                }
                frameHeight = referencePointSize + Math.max(0, lineCount - 1) * referenceLeading;
            } catch (e) {
                frameHeight = FALLBACK_POINT_SIZE + Math.max(0, lineCount - 1) * FALLBACK_LEADING;
            }
        } else {
            /* サイズ指定モード: mm 入力をポイントへ変換 / Size mode: convert the millimeter input to points */
            var heightInMillimeters = parseFloat(heightInput.text) || parseFloat(DEFAULT_HEIGHT_IN_MM);
            frameHeight = heightInMillimeters * MM_TO_POINTS;
        }
    }

    /* フレームの幅 / Frame width */
    var frameWidth = selectionWidth;
    if (widthFromColumnRadio.value) {
        try {
            var columnSourceFrame = selectedText.parentTextFrames[0];
            var columnWidth = columnSourceFrame.textFramePreferences.textColumnFixedWidth;
            if (columnWidth <= 0) {
                var columnFrameBounds = columnSourceFrame.geometricBounds;
                var textFrameWidth = columnFrameBounds[3] - columnFrameBounds[1];
                var columnCount    = columnSourceFrame.textFramePreferences.textColumnCount;
                var columnGutter   = columnSourceFrame.textFramePreferences.textColumnGutter;
                var insetLeft      = columnSourceFrame.textFramePreferences.insetSpacing[1];
                var insetRight     = columnSourceFrame.textFramePreferences.insetSpacing[3];
                columnWidth = (textFrameWidth - insetLeft - insetRight - columnGutter * (columnCount - 1)) / columnCount;
            }
            frameWidth = columnWidth;
        } catch (e) {}
    } else if (widthFromParentFrameRadio.value) {
        try {
            var parentFrameBounds = selectedText.parentTextFrames[0].geometricBounds;
            frameWidth = parentFrameBounds[3] - parentFrameBounds[1];
        } catch (e) {}
    } else if (widthFromMarginRadio.value) {
        try {
            var marginPage;
            if (hasNoSelection) {
                marginPage = app.activeWindow.activePage;
            } else if (isInsertionPoint) {
                var insertionFrame = getInsertionPointTextFrame(selectedText);
                marginPage = insertionFrame ? insertionFrame.parentPage : null;
            } else {
                marginPage = getParentPage(selectedText);
            }
            if (marginPage) {
                var marginPageBounds = marginPage.bounds;
                var pageMargins = marginPage.marginPreferences;
                frameWidth = (marginPageBounds[3] - marginPageBounds[1]) - pageMargins.left - pageMargins.right;
            }
        } catch (e) {}
    }

    // ---------------------------------------
    // フレーム作成 / Frame creation
    // ---------------------------------------

    /**
     * インライン（アンカー付き）でフレームを挿入する
     * @returns {void}
     */
    function createInlineFrame() {
        var anchorFrame = getInsertionPointTextFrame(selectedText);
        if (!anchorFrame) {
            alert(getLabel("alert.parentFrameError"));
            return;
        }

        var anchorStory = anchorFrame.parentStory;
        var charIndex = isInsertionPoint ? selectedText.index : selectedText.characters[0].index;

        /* 元のテキストの段落スタイルを控える / Remember the original paragraph style */
        var originalParaStyle = null;
        var originalParagraph = null;
        if (!isInsertionPoint) {
            originalParaStyle = selectedText.characters[0].appliedParagraphStyle;
            originalParagraph = selectedText.characters[0].paragraphs[0];
        }

        /* 直前が改行でなければ改行を挿入 / Insert a return unless the previous character already is one */
        var needsLeadingReturn = (charIndex > 0) && (anchorStory.characters[charIndex - 1].contents !== "\r");
        if (needsLeadingReturn) {
            anchorStory.insertionPoints[charIndex].contents = "\r";
            charIndex = charIndex + 1;
        }

        var anchorInsertionPoint = anchorStory.insertionPoints[charIndex];
        var anchoredRectangle = anchorInsertionPoint.rectangles.add();
        anchoredRectangle.geometricBounds = [0, 0, frameHeight, frameWidth];
        anchoredRectangle.contentType = ContentType.GRAPHIC_TYPE;
        anchoredRectangle.anchoredObjectSettings.anchoredPosition = AnchorPosition.ANCHORED;
        anchoredRectangle.appliedObjectStyle = activeDoc.objectStyles[selectedObjStyleIndex];

        /* フレーム挿入で文字が 1 つ増えるため、その次の位置に改行を入れる / The inline frame adds one character, so the return goes after it */
        anchorStory.insertionPoints[charIndex + 1].contents = "\r";

        var anchorParagraph = anchorInsertionPoint.paragraphs[0];
        anchorParagraph.appliedParagraphStyle = activeDoc.paragraphStyles[selectedParaStyleIdx];

        if (useAutoLeading) {
            anchorParagraph.autoLeading = 100;
            anchorParagraph.leading = Leading.AUTO;
        }

        if (originalParagraph && originalParaStyle) {
            try {
                originalParagraph.appliedParagraphStyle = originalParaStyle;
            } catch (e) {}
        }
    }

    /**
     * ページ上にグラフィックフレームを作成する
     * @returns {void}
     */
    function createGraphicFrameOnPage() {
        var destinationPage;
        if (hasNoSelection) {
            destinationPage = app.activeWindow.activePage;
        } else if (isInsertionPoint) {
            var insertionFrame = getInsertionPointTextFrame(selectedText);
            destinationPage = insertionFrame ? insertionFrame.parentPage : null;
        } else {
            destinationPage = getParentPage(selectedText);
        }

        if (!destinationPage) {
            alert(getLabel("alert.parentPageError"));
            return;
        }

        var frameTop, frameLeft;
        if (hasTextBounds) {
            frameTop  = selectionBounds[0];
            frameLeft = selectionBounds[1];
        } else {
            /* 挿入ポイント・未選択時はマージン左上を基準にする / Anchor to the top-left margin for insertion points and empty selections */
            var pageMargins = destinationPage.marginPreferences;
            var destinationBounds = destinationPage.bounds;
            frameTop  = destinationBounds[0] + pageMargins.top;
            frameLeft = destinationBounds[1] + pageMargins.left;
        }

        var graphicRectangle = destinationPage.rectangles.add();
        graphicRectangle.geometricBounds = [frameTop, frameLeft, frameTop + frameHeight, frameLeft + frameWidth];
        graphicRectangle.contentType = ContentType.GRAPHIC_TYPE;
        graphicRectangle.appliedObjectStyle = activeDoc.objectStyles[selectedObjStyleIndex];

        /* 回り込みはオブジェクトスタイルが「なし」のときだけ設定 / Apply text wrap only with the "None" object style */
        if (selectedObjStyleIndex === noneObjectStyleIndex) {
            graphicRectangle.textWrapPreferences.textWrapMode = selectedWrapMode;
        }
    }

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(function () {
        if (useInlinePlacement) {
            createInlineFrame();
        } else {
            createGraphicFrameOnPage();
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.createFrame"));

})();
