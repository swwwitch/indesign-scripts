#target indesign
#targetengine "SmartBorderBuilderEngine"

/*

### 概要

選択した表セルに対して、モード・線幅・カラー・濃淡を指定しながら罫線をプレビュー付きで描画・消去します。

詳細は README を参照してください。

### Overview

Draws and clears strokes on the selected table cells with a live preview, choosing the mode, stroke weight, color and tint.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdSmartBorderBuilder";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.8";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-15";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSmartBorderBuilder.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSmartBorderBuilder.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n8c1bcb9a2844"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 前回設定を持ち回すためのグローバルキー / Global key that carries the previous settings between runs */
var SESSION_STATE_KEY = "__SmartBorderBuilderLast";

/* 線幅プリセットの候補（現在の線幅単位で解釈。先頭に「なし」が付く）/ Weight presets, read in the current stroke unit ("None" is added first) */
var WEIGHT_PRESET_VALUES = ["0.1", "0.2", "0.25", "0.35", "0.5"];

/* 濃淡（Tint）の初期値と範囲 / Initial value and range of the tint control */
var TINT_DEFAULT = 100;
var TINT_MIN     = 0;
var TINT_MAX     = 100;

/* Shift を押しながらスライダーを動かしたときの濃淡の刻み / Tint step while dragging the slider with Shift */
var TINT_SNAP_STEP = 10;

/* ↑↓キー1回の増減量。Shift 併用時はこの10倍の刻みに揃える / Arrow-key step; with Shift the value snaps to 10x this step */
var WEIGHT_ARROW_STEP = 0.1;
var TINT_ARROW_STEP   = 1;

/* ［既存の罫線を消してから引く］を切り替えるショートカットキー / Shortcut key that toggles Clear Existing Borders First */
var CLEAR_FIRST_SHORTCUT_KEY = "M";

// =========================================
// レイアウト設定 / Layout settings
// =========================================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 10;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 10;                 /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 10;                 /* 2カラムの間隔 / gap between columns */

/* 入力欄の行・ボタン同士の間隔 / Spacing inside input rows and between buttons */
var CONTROL_SPACING = 8;

/* パネルを持たないオプション群の余白 [左,上,右,下] / Margins of the option groups that have no panel */
var OPTION_GROUP_MARGINS = [16, 10, 16, 10];

/* 線幅プリセットのラジオボタンの間隔 / Spacing between the weight preset radio buttons */
var WEIGHT_PRESET_SPACING = 4;

/* 線幅入力欄・濃淡入力欄の文字数と最小幅（px）/ Character width and minimum width of the weight and tint fields (px) */
var WEIGHT_INPUT_CHARACTERS = 6;
var WEIGHT_INPUT_MIN_WIDTH  = 60;
var TINT_INPUT_CHARACTERS   = 4;
var TINT_INPUT_MIN_WIDTH    = 45;

/* カラーの色見本とドロップダウンの寸法（px）/ Sizes of the swatch chip and dropdown (px) */
var SWATCH_PREVIEW_SIZE        = 18;
var SWATCH_ROW_SPACING         = 6;
var SWATCH_DROPDOWN_WIDTH      = 90;
var SWATCH_DROPDOWN_MIN_HEIGHT = 22;

/* ボタン列の上余白と、左右を分けるスペーサーの最小幅（px）/ Top margin of the button row and minimum width of its spacer (px) */
var BUTTON_ROW_TOP_MARGIN       = 8;
var BUTTON_ROW_SPACER_MIN_WIDTH = 40;

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
 * パネルの共通設定を適用する
 * @param {Panel} targetPanel 対象パネル
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupPanel(targetPanel, spacing) {
    targetPanel.orientation = "column";
    targetPanel.alignChildren = ["fill", "top"];
    targetPanel.alignment = ["fill", "top"];  /* 横並びの中でも縦に伸ばさない / Do not stretch vertically inside a row */
    targetPanel.margins = PANEL_MARGINS;
    targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * 行グループの共通設定を適用する
 * @param {Group} targetGroup 対象グループ
 * @param {string} [alignment] 横方向の配置。省略時は "left"
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupRow(targetGroup, alignment, spacing) {
    targetGroup.orientation = "row";
    targetGroup.alignment = [alignment || "left", "center"];  /* 横と天地を対で / Pair horizontal with vertical */
    targetGroup.alignChildren = ["left", "center"];           /* 親の fill 継承を打ち消す / Cancel the inherited fill */
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

var currentLang = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "罫線の設定", en: "Border Settings" }
    },
    panel: {
        mode:   { ja: "モード", en: "Mode" },
        style:  { ja: "線の設定", en: "Stroke" },
        weight: { ja: "線幅", en: "Weight" },
        color:  { ja: "カラー", en: "Color" }
    },
    radio: {
        all:            { ja: "すべて", en: "All Borders" },
        outer:          { ja: "外枠のみ", en: "Outer Borders" },
        innerOnly:      { ja: "内側のみ", en: "Inner Borders" },
        horizontal:     { ja: "水平線のみ", en: "Horizontal Borders" },
        vertical:       { ja: "垂直線のみ", en: "Vertical Borders" },
        bottomOnly:     { ja: "下端のみ", en: "Bottom Edge Only" },
        rightOnly:      { ja: "右端のみ", en: "Right Edge Only" },
        headerRow:      { ja: "見出し行", en: "Header Row" },
        headerColumn:   { ja: "見出し列", en: "Header Column" },
        clearLeftRight: { ja: "左右の外枠を消去", en: "Clear Outer Left/Right Borders" },
        allOff:         { ja: "すべて消去", en: "Clear All Borders" },
        weightNone:     { ja: "なし", en: "None" }
    },
    checkbox: {
        clearFirst: { ja: "既存の罫線を消してから引く", en: "Clear Existing Borders First" }
    },
    fieldLabel: {
        tint: { ja: "濃淡", en: "Tint" }
    },
    swatch: {
        black: { ja: "黒", en: "Black" },
        paper: { ja: "紙色", en: "Paper" },
        none:  { ja: "なし", en: "None" }
    },
    button: {
        ok:                { ja: "OK", en: "OK" },
        cancel:            { ja: "キャンセル", en: "Cancel" },
        screenModeNormal:  { ja: "標準モード", en: "Normal" },
        screenModePreview: { ja: "プレビュー", en: "Preview" }
    },
    tooltip: {
        all: {
            ja: "ショートカット：A／Option（Alt）+クリックで［既存の罫線を消してから引く］を切り替え",
            en: "Shortcut: A / Option (Alt)-click toggles Clear Existing Borders First"
        },
        outer: {
            ja: "ショートカット：E／Option（Alt）+クリックで［既存の罫線を消してから引く］を切り替え",
            en: "Shortcut: E / Option (Alt)-click toggles Clear Existing Borders First"
        },
        innerOnly: {
            ja: "ショートカット：I／Option（Alt）+クリックで［既存の罫線を消してから引く］を切り替え",
            en: "Shortcut: I / Option (Alt)-click toggles Clear Existing Borders First"
        },
        horizontal: {
            ja: "外枠の上下を含む、すべての水平線を引きます。\nショートカット：H／Option（Alt）+クリックで［既存の罫線を消してから引く］を切り替え",
            en: "Draws every horizontal line, including the outer top and bottom.\nShortcut: H / Option (Alt)-click toggles Clear Existing Borders First"
        },
        vertical: {
            ja: "外枠の左右を含む、すべての垂直線を引きます。\nショートカット：V／Option（Alt）+クリックで［既存の罫線を消してから引く］を切り替え",
            en: "Draws every vertical line, including the outer left and right.\nShortcut: V / Option (Alt)-click toggles Clear Existing Borders First"
        },
        bottomOnly: {
            ja: "ショートカット：B／Option（Alt）+クリックで［既存の罫線を消してから引く］を切り替え",
            en: "Shortcut: B / Option (Alt)-click toggles Clear Existing Borders First"
        },
        rightOnly: {
            ja: "Option（Alt）+クリックで［既存の罫線を消してから引く］を切り替え",
            en: "Option (Alt)-click toggles Clear Existing Borders First"
        },
        headerRow: {
            ja: "表の上端・1行目の下・表の下端に線を引きます。表全体を選択しているときに使えます。\nショートカット：U／Option（Alt）+クリックで［既存の罫線を消してから引く］を切り替え",
            en: "Draws the top edge, the line below the first row, and the bottom edge. Available when the whole table is selected.\nShortcut: U / Option (Alt)-click toggles Clear Existing Borders First"
        },
        headerColumn: {
            ja: "表の左端・1列目の右・表の右端に線を引きます。表全体を選択しているときに使えます。\nショートカット：L／Option（Alt）+クリックで［既存の罫線を消してから引く］を切り替え",
            en: "Draws the left edge, the line right of the first column, and the right edge. Available when the whole table is selected.\nShortcut: L / Option (Alt)-click toggles Clear Existing Borders First"
        },
        clearLeftRight: {
            ja: "ショートカット：R／Option（Alt）+クリックで［既存の罫線を消してから引く］を切り替え",
            en: "Shortcut: R / Option (Alt)-click toggles Clear Existing Borders First"
        },
        allOff: {
            ja: "ショートカット：C／Option（Alt）+クリックで［既存の罫線を消してから引く］を切り替え",
            en: "Shortcut: C / Option (Alt)-click toggles Clear Existing Borders First"
        },
        clearFirst: {
            ja: "オンにすると、選択範囲の罫線をいったん消してから引きます。\nMキーでも切り替えられます。",
            en: "When on, borders in the selection are cleared before drawing.\nPress M to toggle."
        },
        weightInput: {
            ja: "↑↓キーで0.1ずつ、Shift+↑↓キーで1ずつ増減します。",
            en: "Up/Down arrow keys change the value by 0.1, or by 1 with Shift."
        },
        swatchDropdown: {
            ja: "「なし」と「紙色」では濃淡を指定できません。",
            en: "Tint is not available for None or Paper."
        },
        tintInput: {
            ja: "↑↓キーで1ずつ、Shift+↑↓キーで10ずつ増減します。",
            en: "Up/Down arrow keys change the value by 1, or by 10 with Shift."
        },
        tintSlider: { ja: "Shift+ドラッグで10%刻み", en: "Shift-drag to snap to 10% steps" },
        screenMode: {
            ja: "ドキュメントウィンドウの画面モードを、標準モードとプレビューで切り替えます。",
            en: "Switches the document window between the Normal and Preview screen modes."
        }
    },
    alert: {
        noCellSelection: { ja: "表のセルを選択してください。", en: "Please select table cells." },
        invalidWeight:   { ja: "線幅には0以上の数値を入力してください。", en: "Enter a value of 0 or greater for the stroke weight." },
        invalidTint:     { ja: "濃淡には0〜100の数値を入力してください。", en: "Enter a value between 0 and 100 for tint." }
    },
    undo: {
        previewBorders: { ja: "罫線プレビュー", en: "Border Preview" },
        applyBorders:   { ja: "罫線の設定", en: "Apply Border Settings" }
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
    return labelNode[currentLang] || labelNode.en || labelKey;
}

/**
 * 項目名にコロンを付けて返す（日本語は全角、英語は半角）
 * @param {string} labelKey 例: "fieldLabel.tint"
 * @returns {string} コロン付きのラベル文字列
 */
function labelText(labelKey) {
    return getLabel(labelKey) + (currentLang === "ja" ? "：" : ":");
}

// =========================================
// 罫線モード / Border modes
// =========================================

/**
 * @typedef {object} BorderMode
 * @property {string} id 識別子。LABELS.radio / LABELS.tooltip のキーと、前回設定の保存値を兼ねる
 * @property {string} shortcutKey 選択用のショートカットキー。無い場合は空文字
 * @property {boolean} needsWholeTable 表全体を選択しているときだけ使えるか
 * @property {function(BorderTarget, BorderStroke, boolean): void} apply 罫線を適用する関数
 */

/* ラジオボタンの並び順どおりに定義する / Listed in radio-button order */
var BORDER_MODES = [
    { id: "all",            shortcutKey: "A", needsWholeTable: false, apply: applyAllBorders },
    { id: "outer",          shortcutKey: "E", needsWholeTable: false, apply: applyOuterBorders },
    { id: "innerOnly",      shortcutKey: "I", needsWholeTable: false, apply: applyInnerBorders },
    { id: "horizontal",     shortcutKey: "H", needsWholeTable: false, apply: applyHorizontalBorders },
    { id: "vertical",       shortcutKey: "V", needsWholeTable: false, apply: applyVerticalBorders },
    { id: "bottomOnly",     shortcutKey: "B", needsWholeTable: false, apply: applyBottomEdgeBorder },
    { id: "rightOnly",      shortcutKey: "",  needsWholeTable: false, apply: applyRightEdgeBorder },
    { id: "headerRow",      shortcutKey: "U", needsWholeTable: true,  apply: applyHeaderRowBorders },
    { id: "headerColumn",   shortcutKey: "L", needsWholeTable: true,  apply: applyHeaderColumnBorders },
    { id: "clearLeftRight", shortcutKey: "R", needsWholeTable: false, apply: clearOuterLeftRightBorders },
    { id: "allOff",         shortcutKey: "C", needsWholeTable: false, apply: clearAllBorders }
];

/**
 * ショートカットキーに対応するモードを探す
 * @param {string} keyName 押されたキーの名前
 * @returns {BorderMode|null} 該当するモード。無い場合は null
 */
function findModeByShortcutKey(keyName) {
    for (var i = 0; i < BORDER_MODES.length; i++) {
        if (BORDER_MODES[i].shortcutKey !== "" && BORDER_MODES[i].shortcutKey === keyName) return BORDER_MODES[i];
    }
    return null;
}

// =========================================
// 前回設定の保存と復元 / Remembering the previous settings
// =========================================

/**
 * 前回 OK で確定した設定を読み込む
 * @returns {object|null} 前回の設定。保存されていない場合は null
 */
function loadLastSettings() {
    return $.global[SESSION_STATE_KEY] || null;
}

/**
 * 次回起動時に復元できるよう、確定した設定を保存する
 * @param {object} settings 保存する設定
 * @returns {void}
 */
function saveLastSettings(settings) {
    $.global[SESSION_STATE_KEY] = settings;
}

// =========================================
// セルの取得 / Cell lookup
// =========================================

/**
 * @typedef {object} CellRange
 * @property {number} startRow 開始行
 * @property {number} endRow 終了行
 * @property {number} startCol 開始列
 * @property {number} endCol 終了列
 */

/**
 * @typedef {object} CellEntry
 * @property {Cell} cell セル
 * @property {CellRange} range 結合を考慮した占有範囲
 */

/**
 * @typedef {object} CellBounds
 * @property {number} minRow 最初の行
 * @property {number} maxRow 最後の行
 * @property {number} minCol 最初の列
 * @property {number} maxCol 最後の列
 */

/**
 * @typedef {object} BorderTarget
 * @property {CellBounds} bounds 選択範囲の行・列の範囲
 * @property {CellEntry[]} selectedEntries 選択したセル
 * @property {CellEntry[]} blockEntries 選択範囲の矩形に掛かるセル
 */

/**
 * 選択オブジェクトから表セルを取り出す
 * @param {object} selectionItem 選択オブジェクト
 * @returns {Cell[]} 表セルの配列
 */
function getCellsFromSelectionItem(selectionItem) {
    var cells = [];
    try {
        if (selectionItem.constructor.name === "Cell") {
            cells.push(selectionItem);
        } else if (selectionItem.hasOwnProperty("cells") && selectionItem.cells.length > 0) {
            for (var i = 0; i < selectionItem.cells.length; i++) {
                cells.push(selectionItem.cells[i]);
            }
        } else if (selectionItem.parent && selectionItem.parent.constructor.name === "Cell") {
            cells.push(selectionItem.parent);
        }
    } catch (e) {}
    return cells;
}

/**
 * 選択から対象の表セルを重複なく集める
 * @param {Array} selectionItems 現在の選択
 * @returns {CellEntry[]} 選択されたセル
 */
function collectSelectedCellEntries(selectionItems) {
    var entries = [];
    var seenKeys = {};
    var i, j, cells, range, cellKey;

    for (i = 0; i < selectionItems.length; i++) {
        cells = getCellsFromSelectionItem(selectionItems[i]);
        for (j = 0; j < cells.length; j++) {
            range = getCellRange(cells[j]);
            cellKey = [cells[j].parent.id, range.startRow, range.endRow, range.startCol, range.endCol].join(":");
            if (seenKeys[cellKey]) continue;
            seenKeys[cellKey] = true;
            entries.push({ cell: cells[j], range: range });
        }
    }
    return entries;
}

/**
 * セルの結合数を読む。読めない場合は 1 とみなす
 * @param {Cell} cell 対象のセル
 * @param {string} propertyName "rowSpan" または "columnSpan"
 * @returns {number} 1 以上の結合数
 */
function readCellSpan(cell, propertyName) {
    try {
        return Math.max(1, Number(cell[propertyName]) || 1);
    } catch (e) {
        return 1;
    }
}

/**
 * 結合を考慮したセルの占有範囲を求める
 * @param {Cell} cell 対象のセル
 * @returns {CellRange} 行と列の占有範囲
 */
function getCellRange(cell) {
    var startRow = cell.parentRow.index;
    var startCol = cell.parentColumn.index;
    return {
        startRow: startRow,
        endRow: startRow + readCellSpan(cell, "rowSpan") - 1,
        startCol: startCol,
        endCol: startCol + readCellSpan(cell, "columnSpan") - 1
    };
}

/**
 * 2 つの区間が重なるかを判定する
 * @param {number} startA 区間 A の始まり
 * @param {number} endA 区間 A の終わり
 * @param {number} startB 区間 B の始まり
 * @param {number} endB 区間 B の終わり
 * @returns {boolean} 重なっていれば true
 */
function spansOverlap(startA, endA, startB, endB) {
    return !(endA < startB || endB < startA);
}

/**
 * セル全体の行・列の範囲を求める
 * @param {CellEntry[]} entries 対象のセル
 * @returns {CellBounds} 行と列の範囲
 */
function getEntriesBounds(entries) {
    var bounds = { minRow: 999999, maxRow: -1, minCol: 999999, maxCol: -1 };
    var range;
    for (var i = 0; i < entries.length; i++) {
        range = entries[i].range;
        if (range.startRow < bounds.minRow) bounds.minRow = range.startRow;
        if (range.endRow > bounds.maxRow) bounds.maxRow = range.endRow;
        if (range.startCol < bounds.minCol) bounds.minCol = range.startCol;
        if (range.endCol > bounds.maxCol) bounds.maxCol = range.endCol;
    }
    return bounds;
}

/**
 * 表のセルのうち、指定した範囲に掛かるものを集める
 * @param {Table} table 対象の表
 * @param {CellBounds} bounds 行と列の範囲
 * @returns {CellEntry[]} 範囲に掛かるセル
 */
function collectCellEntriesInBounds(table, bounds) {
    var tableCells = table.cells;
    var cellCount = tableCells.length;
    var entries = [];
    var i, cell, range;

    for (i = 0; i < cellCount; i++) {
        cell = tableCells[i];
        range = getCellRange(cell);
        if (spansOverlap(range.startRow, range.endRow, bounds.minRow, bounds.maxRow) &&
            spansOverlap(range.startCol, range.endCol, bounds.minCol, bounds.maxCol)) {
            entries.push({ cell: cell, range: range });
        }
    }
    return entries;
}

/**
 * 罫線を引く対象（範囲・選択セル・矩形内のセル）をまとめる
 * @param {CellEntry[]} selectedEntries 選択したセル
 * @returns {BorderTarget} 罫線の適用対象
 */
function buildBorderTarget(selectedEntries) {
    var bounds = getEntriesBounds(selectedEntries);
    return {
        bounds: bounds,
        selectedEntries: selectedEntries,
        blockEntries: collectCellEntriesInBounds(selectedEntries[0].cell.parent, bounds)
    };
}

/**
 * 表全体が選択されているかを判定する
 * @param {CellEntry[]} selectedEntries 選択したセル
 * @returns {boolean} 表全体なら true
 */
function isWholeTableSelected(selectedEntries) {
    return selectedEntries.length === selectedEntries[0].cell.parent.cells.length;
}

/**
 * 控えておいた選択状態を復元する
 * @param {Array} selectionItems 実行前の選択
 * @returns {void}
 */
function restoreSelection(selectionItems) {
    var validItems = [];
    for (var i = 0; i < selectionItems.length; i++) {
        if (selectionItems[i] && selectionItems[i].isValid !== false) validItems.push(selectionItems[i]);
    }
    if (validItems.length === 0) return;
    try {
        app.select(validItems);
    } catch (e) {}
}

// =========================================
// 線幅と濃淡の値 / Weight and tint values
// =========================================

/**
 * @typedef {object} StrokeUnitInfo
 * @property {string} label 入力欄の横に出す単位
 * @property {string} suffix 線幅の文字列に付ける単位
 * @property {string} defaultWeightText 線幅入力欄の初期値
 */

/**
 * ドキュメントの線幅単位を調べる
 * @returns {StrokeUnitInfo} 線幅単位の表示と初期値
 */
function getStrokeUnitInfo() {
    switch (app.activeDocument.viewPreferences.strokeMeasurementUnits) {
        case MeasurementUnits.POINTS:
            return { label: "pt", suffix: "pt", defaultWeightText: "0.25" };
        case MeasurementUnits.MILLIMETERS:
            return { label: "mm", suffix: "mm", defaultWeightText: "0.1" };
        case MeasurementUnits.CENTIMETERS:
            return { label: "cm", suffix: "cm", defaultWeightText: "0.1" };
        case MeasurementUnits.INCHES:
            return { label: "in", suffix: "in", defaultWeightText: "0.1" };
        case MeasurementUnits.PICAS:
            return { label: "pica", suffix: "p", defaultWeightText: "0.1" };
        case MeasurementUnits.Q:
            return { label: "Q", suffix: "q", defaultWeightText: "0.1" };
        default:
            return { label: "pt", suffix: "pt", defaultWeightText: "0.1" };
    }
}

/**
 * 線幅入力欄の文字列を取得する。空欄なら初期値を返す
 * @param {object} ui UI オブジェクト
 * @param {object} state 状態オブジェクト
 * @returns {string} 線幅を表す文字列
 */
function getWeightText(ui, state) {
    var weightText = String(ui.weightInput.text).replace(/^\s+|\s+$/g, "");
    return (weightText !== "") ? weightText : state.strokeUnit.defaultWeightText;
}

/**
 * 濃淡の文字列を数値にして有効範囲に収める
 * @param {string} text 濃淡の文字列
 * @returns {number} 範囲内に収めた値。数値でなければ NaN
 */
function parseTintText(text) {
    var tint = parseFloat(String(text));
    if (isNaN(tint)) return NaN;
    return Math.min(TINT_MAX, Math.max(TINT_MIN, tint));
}

// =========================================
// スウォッチ / Swatches
// =========================================

/* 特別なスウォッチの名前（英語版・日本語版の表記を含む）/ Names of the built-in swatches, including Japanese variants */
var SPECIAL_SWATCH_NAMES = {
    none:         ["None", "[None]", "なし", "[なし]"],
    black:        ["Black", "[Black]", "ブラック", "黒"],
    paper:        ["Paper", "[Paper]", "紙色", "[紙色]"],
    registration: ["Registration", "[Registration]", "レジストレーション", "[レジストレーション]"]
};

/**
 * スウォッチ名が特別なスウォッチのどれに当たるかを調べる
 * @param {string} swatchName スウォッチ名
 * @returns {string} "none" / "black" / "paper" / "registration"。どれでもなければ空文字
 */
function getSpecialSwatchKind(swatchName) {
    var swatchKind, names, i;
    for (swatchKind in SPECIAL_SWATCH_NAMES) {
        if (!SPECIAL_SWATCH_NAMES.hasOwnProperty(swatchKind)) continue;
        names = SPECIAL_SWATCH_NAMES[swatchKind];
        for (i = 0; i < names.length; i++) {
            if (names[i] === swatchName) return swatchKind;
        }
    }
    return "";
}

/**
 * カラー候補として表示するスウォッチ一覧を作る（レジストレーションは除く）
 * @returns {Array<{swatchName: string, displayName: string}>} スウォッチ名と表示名の配列
 */
function getSwatchEntries() {
    var swatches = app.activeDocument.swatches;
    var entries = [];
    var i, swatchName, swatchKind;

    for (i = 0; i < swatches.length; i++) {
        swatchName = String(swatches[i].name);
        swatchKind = getSpecialSwatchKind(swatchName);
        if (swatchKind === "registration") continue;
        entries.push({
            swatchName: swatchName,
            displayName: swatchKind ? getLabel("swatch." + swatchKind) : swatchName
        });
    }
    return entries;
}

/**
 * 既定で選択するカラーの位置を求める（黒があれば黒）
 * @param {Array<{swatchName: string}>} swatchEntries スウォッチ一覧
 * @returns {number} 既定で選ぶ位置
 */
function getDefaultSwatchIndex(swatchEntries) {
    for (var i = 0; i < swatchEntries.length; i++) {
        if (getSpecialSwatchKind(swatchEntries[i].swatchName) === "black") return i;
    }
    return 0;
}

/**
 * 名前からスウォッチを取得する
 * @param {string} swatchName スウォッチ名
 * @returns {Swatch|null} スウォッチ。見つからない場合は null
 */
function getSwatchByName(swatchName) {
    if (!swatchName) return null;
    var swatch = app.activeDocument.swatches.itemByName(swatchName);
    return swatch.isValid ? swatch : null;
}

/**
 * 色見本の描画に使う RGBA 値を求める
 * @param {Swatch} swatch 対象のスウォッチ
 * @param {string} swatchKind getSpecialSwatchKind() の結果
 * @returns {number[]} 0〜1 の RGBA 値
 */
function getSwatchPreviewColor(swatch, swatchKind) {
    var colorValue, cyan, magenta, yellow, black;

    if (swatchKind === "none" || swatchKind === "paper") return [1, 1, 1, 1];
    if (swatchKind === "black" || swatchKind === "registration") return [0, 0, 0, 1];

    try {
        if (swatch.hasOwnProperty("colorValue")) {
            colorValue = swatch.colorValue;
            if (swatch.space === ColorSpace.RGB) {
                return [colorValue[0] / 255, colorValue[1] / 255, colorValue[2] / 255, 1];
            }
            if (swatch.space === ColorSpace.CMYK) {
                cyan = colorValue[0] / 100;
                magenta = colorValue[1] / 100;
                yellow = colorValue[2] / 100;
                black = colorValue[3] / 100;
                return [(1 - cyan) * (1 - black), (1 - magenta) * (1 - black), (1 - yellow) * (1 - black), 1];
            }
        }
    } catch (e) {}

    return [0.5, 0.5, 0.5, 1];
}

// =========================================
// UI構築 / Build UI
// =========================================

/**
 * 罫線設定ダイアログを組み立てる
 * @param {object} state 状態オブジェクト
 * @returns {object} ダイアログとコントロールをまとめた UI オブジェクト
 */
function buildDialog(state) {
    var dlg = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    var ui = { dlg: dlg };
    setupWindow(dlg);

    /* 左：モード、右：スタイル / Left: mode, right: style */
    var columnsGroup = dlg.add("group");
    columnsGroup.orientation = "row";
    columnsGroup.alignChildren = ["fill", "top"];
    columnsGroup.alignment = ["fill", "top"];
    columnsGroup.spacing = COLUMN_SPACING;

    var modeColumnGroup = columnsGroup.add("group");
    modeColumnGroup.orientation = "column";
    modeColumnGroup.alignChildren = ["fill", "top"];
    modeColumnGroup.alignment = ["fill", "top"];
    modeColumnGroup.spacing = WINDOW_SPACING;

    addModePanel(modeColumnGroup, ui, state.isWholeTableSelected);
    addClearFirstOption(modeColumnGroup, ui);

    var stylePanel = columnsGroup.add("panel", undefined, getLabel("panel.style"));
    setupPanel(stylePanel);
    addWeightPanel(stylePanel, ui, state);
    addColorPanel(stylePanel, ui);

    addButtonRow(dlg, ui);
    return ui;
}

/**
 * モードのラジオボタンを並べたパネルを追加する
 * @param {Group} parent 追加先のグループ
 * @param {object} ui UI オブジェクト
 * @param {boolean} wholeTableSelected 表全体を選択しているか
 * @returns {void}
 */
function addModePanel(parent, ui, wholeTableSelected) {
    var modePanel = parent.add("panel", undefined, getLabel("panel.mode"));
    var i, borderMode, modeRadio;
    setupPanel(modePanel);

    ui.modeRadios = {};
    for (i = 0; i < BORDER_MODES.length; i++) {
        borderMode = BORDER_MODES[i];
        modeRadio = modePanel.add("radiobutton", undefined, getLabel("radio." + borderMode.id));
        modeRadio.helpTip = getLabel("tooltip." + borderMode.id);
        if (borderMode.needsWholeTable) modeRadio.enabled = wholeTableSelected;
        ui.modeRadios[borderMode.id] = modeRadio;
    }
    ui.modeRadios.all.value = true;
}

/**
 * ［既存の罫線を消してから引く］チェックボックスを追加する
 * @param {Group} parent 追加先のグループ
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function addClearFirstOption(parent, ui) {
    var clearFirstGroup = parent.add("group");
    clearFirstGroup.orientation = "column";
    clearFirstGroup.alignChildren = "left";
    clearFirstGroup.alignment = ["fill", "top"];
    clearFirstGroup.margins = OPTION_GROUP_MARGINS;

    ui.cbClearFirst = clearFirstGroup.add("checkbox", undefined, getLabel("checkbox.clearFirst"));
    ui.cbClearFirst.helpTip = getLabel("tooltip.clearFirst");
}

/**
 * 線幅の入力欄とプリセットを並べたパネルを追加する
 * @param {Panel} parent 追加先のパネル
 * @param {object} ui UI オブジェクト
 * @param {object} state 状態オブジェクト
 * @returns {void}
 */
function addWeightPanel(parent, ui, state) {
    var weightPanel = parent.add("panel", undefined, getLabel("panel.weight"));
    var presetTexts = ["0"].concat(WEIGHT_PRESET_VALUES);
    var i, presetLabel;
    setupPanel(weightPanel, CONTROL_SPACING);

    var weightInputRow = weightPanel.add("group");
    setupRow(weightInputRow, "left", CONTROL_SPACING);
    ui.weightInput = weightInputRow.add("edittext", undefined, state.strokeUnit.defaultWeightText);
    ui.weightInput.characters = WEIGHT_INPUT_CHARACTERS;
    ui.weightInput.minimumSize.width = WEIGHT_INPUT_MIN_WIDTH;
    ui.weightInput.helpTip = getLabel("tooltip.weightInput");
    weightInputRow.add("statictext", undefined, state.strokeUnit.label);

    var weightPresetGroup = weightPanel.add("group");
    weightPresetGroup.orientation = "column";
    weightPresetGroup.alignChildren = ["left", "center"];
    weightPresetGroup.alignment = ["fill", "top"];
    weightPresetGroup.margins = OPTION_GROUP_MARGINS;
    weightPresetGroup.spacing = WEIGHT_PRESET_SPACING;

    ui.weightPresets = [];
    for (i = 0; i < presetTexts.length; i++) {
        presetLabel = (i === 0) ? getLabel("radio.weightNone") : presetTexts[i];
        ui.weightPresets.push({
            text: presetTexts[i],
            radio: weightPresetGroup.add("radiobutton", undefined, presetLabel)
        });
    }
    syncWeightPresets(ui, state);
}

/**
 * カラーと濃淡のパネルを追加する
 * @param {Panel} parent 追加先のパネル
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function addColorPanel(parent, ui) {
    var colorPanel = parent.add("panel", undefined, getLabel("panel.color"));
    setupPanel(colorPanel);

    var swatchRow = colorPanel.add("group");
    setupRow(swatchRow, "left", SWATCH_ROW_SPACING);

    ui.swatchPreviewBox = swatchRow.add("group");
    ui.swatchPreviewBox.preferredSize = [SWATCH_PREVIEW_SIZE, SWATCH_PREVIEW_SIZE];
    ui.swatchPreviewBox.minimumSize = [SWATCH_PREVIEW_SIZE, SWATCH_PREVIEW_SIZE];
    ui.swatchPreviewBox.maximumSize = [SWATCH_PREVIEW_SIZE, SWATCH_PREVIEW_SIZE];

    ui.swatchDropdown = addSwatchDropdown(swatchRow, getSwatchEntries());
    ui.swatchDropdown.helpTip = getLabel("tooltip.swatchDropdown");

    var tintInputRow = colorPanel.add("group");
    setupRow(tintInputRow, "fill", CONTROL_SPACING);
    tintInputRow.add("statictext", undefined, labelText("fieldLabel.tint"));
    ui.tintInput = tintInputRow.add("edittext", undefined, String(TINT_DEFAULT));
    ui.tintInput.characters = TINT_INPUT_CHARACTERS;
    ui.tintInput.minimumSize.width = TINT_INPUT_MIN_WIDTH;
    ui.tintInput.helpTip = getLabel("tooltip.tintInput");

    ui.tintSlider = colorPanel.add("slider", undefined, TINT_DEFAULT, TINT_MIN, TINT_MAX);
    ui.tintSlider.helpTip = getLabel("tooltip.tintSlider");

    refreshSwatchControls(ui);
}

/**
 * スウォッチ選択のドロップダウンを追加する
 * @param {Group} parent 追加先のグループ
 * @param {Array<{swatchName: string, displayName: string}>} swatchEntries スウォッチ一覧
 * @returns {DropDownList} 追加したドロップダウン
 */
function addSwatchDropdown(parent, swatchEntries) {
    var displayNames = [];
    var i;
    for (i = 0; i < swatchEntries.length; i++) {
        displayNames.push(swatchEntries[i].displayName);
    }

    var swatchDropdown = parent.add("dropdownlist", undefined, displayNames);
    swatchDropdown.minimumSize.height = SWATCH_DROPDOWN_MIN_HEIGHT;
    swatchDropdown.minimumSize.width = SWATCH_DROPDOWN_WIDTH;
    swatchDropdown.preferredSize.width = SWATCH_DROPDOWN_WIDTH;

    /* 表示名とは別に実際のスウォッチ名を持たせる / Keep the real swatch name apart from the display name */
    for (i = 0; i < swatchDropdown.items.length; i++) {
        swatchDropdown.items[i]._swatchName = swatchEntries[i].swatchName;
    }
    if (swatchDropdown.items.length > 0) {
        swatchDropdown.selection = getDefaultSwatchIndex(swatchEntries);
    }
    return swatchDropdown;
}

/**
 * ダイアログ下部のボタン列を追加する
 * @param {Window} dlg 対象のダイアログ
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function addButtonRow(dlg, ui) {
    // メイングループ（横並び） / Main group (horizontal layout)
    var btnRowGroup = dlg.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
    btnRowGroup.alignment = ["fill", "bottom"];

    // 左側グループ / Left-side button group
    var btnLeftGroup = btnRowGroup.add("group");
    btnLeftGroup.alignChildren = ["left", "center"];
    ui.btnScreenMode = btnLeftGroup.add("button", undefined, getScreenModeButtonLabel());
    ui.btnScreenMode.helpTip = getLabel("tooltip.screenMode");

    // スペーサー（伸縮）/ Spacer (stretchable)
    var spacer = btnRowGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = BUTTON_ROW_SPACER_MIN_WIDTH;

    // 右側グループ / Right-side button group
    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.alignChildren = ["right", "center"];
    btnRightGroup.spacing = CONTROL_SPACING;
    var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
}

// =========================================
// UIの値の読み書き / Read and write UI values
// =========================================

/**
 * 選択中のモードを取得する
 * @param {object} ui UI オブジェクト
 * @returns {BorderMode|null} 選択中のモード。無い場合は null
 */
function getSelectedMode(ui) {
    for (var i = 0; i < BORDER_MODES.length; i++) {
        if (ui.modeRadios[BORDER_MODES[i].id].value) return BORDER_MODES[i];
    }
    return null;
}

/**
 * モードのラジオボタンを選択する。無効なモードは選ばない
 * @param {object} ui UI オブジェクト
 * @param {string} modeId モードの識別子
 * @returns {boolean} 選択できたら true
 */
function selectModeRadio(ui, modeId) {
    var modeRadio = ui.modeRadios[modeId];
    if (!modeRadio || !modeRadio.enabled) return false;
    modeRadio.value = true;
    return true;
}

/**
 * 線幅入力欄の値に合わせてプリセットの選択状態を揃える
 * @param {object} ui UI オブジェクト
 * @param {object} state 状態オブジェクト
 * @returns {void}
 */
function syncWeightPresets(ui, state) {
    var weightValue = parseFloat(getWeightText(ui, state));
    if (isNaN(weightValue)) return;
    for (var i = 0; i < ui.weightPresets.length; i++) {
        ui.weightPresets[i].radio.value = (parseFloat(ui.weightPresets[i].text) === weightValue);
    }
}

/**
 * ドロップダウンで選択中のスウォッチ名を取得する
 * @param {object} ui UI オブジェクト
 * @returns {string} スウォッチ名。未選択なら空文字
 */
function getSelectedSwatchName(ui) {
    var selectedItem = ui.swatchDropdown.selection;
    return selectedItem ? String(selectedItem._swatchName) : "";
}

/**
 * カラーのドロップダウンをスウォッチ名で選択する
 * @param {object} ui UI オブジェクト
 * @param {string} swatchName 選択したいスウォッチ名
 * @returns {void}
 */
function selectSwatchByName(ui, swatchName) {
    var dropdownItems = ui.swatchDropdown.items;
    for (var i = 0; i < dropdownItems.length; i++) {
        if (dropdownItems[i]._swatchName === swatchName) {
            ui.swatchDropdown.selection = i;
            return;
        }
    }
}

/**
 * 選択中のスウォッチに合わせて色見本と濃淡コントロールを更新する
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function refreshSwatchControls(ui) {
    var swatchName = getSelectedSwatchName(ui);
    var swatchKind = getSpecialSwatchKind(swatchName);
    var swatch = getSwatchByName(swatchName);
    var previewGraphics = ui.swatchPreviewBox.graphics;
    var tintAdjustable = (swatchKind !== "none" && swatchKind !== "paper");

    if (swatch) {
        previewGraphics.backgroundColor = previewGraphics.newBrush(
            previewGraphics.BrushType.SOLID_COLOR,
            getSwatchPreviewColor(swatch, swatchKind)
        );
        ui.dlg.update();
    }
    ui.tintInput.enabled = tintAdjustable;
    ui.tintSlider.enabled = tintAdjustable;
}

/**
 * 濃淡入力欄の値を整数に丸めて範囲に収め、スライダーへ反映する
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function commitTintInput(ui) {
    var tint = parseTintText(ui.tintInput.text);
    if (isNaN(tint)) return;
    tint = Math.round(tint);
    ui.tintInput.text = String(tint);
    ui.tintSlider.value = tint;
}

/**
 * スライダーの値を入力欄へ反映する
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function syncTintInputFromSlider(ui) {
    ui.tintInput.text = String(Math.round(ui.tintSlider.value));
}

/**
 * Shift を押していればスライダーを刻みに揃え、確定時は整数に丸める
 * @param {Slider} tintSlider 濃淡のスライダー
 * @param {boolean} roundToInteger Shift なしのときに整数へ丸めるか
 * @returns {void}
 */
function snapTintSlider(tintSlider, roundToInteger) {
    if (ScriptUI.environment.keyboardState.shiftKey) {
        tintSlider.value = Math.round(tintSlider.value / TINT_SNAP_STEP) * TINT_SNAP_STEP;
    } else if (roundToInteger) {
        tintSlider.value = Math.round(tintSlider.value);
    }
}

/**
 * 現在プレビュー表示になっているかを判定する
 * @returns {boolean} プレビュー表示なら true
 */
function isPreviewScreenMode() {
    try {
        return app.activeWindow.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE;
    } catch (e) {
        return false;
    }
}

/**
 * 標準モードとプレビューを切り替える
 * @returns {void}
 */
function toggleScreenMode() {
    try {
        app.activeWindow.screenMode = isPreviewScreenMode() ? ScreenModeOptions.PREVIEW_OFF : ScreenModeOptions.PREVIEW_TO_PAGE;
    } catch (e) {}
}

/**
 * 画面モードの切り替えボタンに出すラベルを返す（切り替え先のモード名）
 * @returns {string} ボタンに表示する文字列
 */
function getScreenModeButtonLabel() {
    return getLabel(isPreviewScreenMode() ? "button.screenModeNormal" : "button.screenModePreview");
}

/**
 * @typedef {object} BorderStroke
 * @property {string|number} weight 線幅（単位付きの文字列、または 0）
 * @property {Swatch|null} swatch 罫線のカラー
 * @property {number} tint 濃淡
 */

/**
 * @typedef {object} BorderSettings
 * @property {BorderMode|null} mode 選択中のモード
 * @property {boolean} clearFirst 描画前に既存の罫線を消すか
 * @property {string} weightText 線幅入力欄の文字列
 * @property {string} swatchName 選択中のスウォッチ名
 * @property {boolean} isWeightValid 線幅が 0 以上の数値か
 * @property {boolean} isTintValid 濃淡が数値か
 * @property {BorderStroke} stroke 罫線に設定する値
 */

/**
 * ダイアログの値から罫線の設定を読み取る
 * @param {object} ui UI オブジェクト
 * @param {object} state 状態オブジェクト
 * @returns {BorderSettings} 罫線の設定
 */
function readBorderSettings(ui, state) {
    var weightText = getWeightText(ui, state);
    var weightValue = parseFloat(weightText);
    var swatchName = getSelectedSwatchName(ui);
    var tint = parseTintText(ui.tintInput.text);

    return {
        mode: getSelectedMode(ui),
        clearFirst: ui.cbClearFirst.value,
        weightText: weightText,
        swatchName: swatchName,
        isWeightValid: !isNaN(weightValue) && weightValue >= 0,
        isTintValid: !isNaN(tint),
        stroke: {
            weight: (weightValue === 0) ? 0 : String(weightValue) + state.strokeUnit.suffix,
            swatch: getSwatchByName(swatchName),
            tint: tint
        }
    };
}

/**
 * 罫線を適用できる設定かどうかを判定する
 * @param {BorderSettings} settings 罫線の設定
 * @returns {boolean} 適用できれば true
 */
function canApplySettings(settings) {
    return !!(settings.mode && settings.isWeightValid && settings.isTintValid && settings.stroke.swatch);
}

/**
 * 前回保存した設定をダイアログへ復元する
 * @param {object} ui UI オブジェクト
 * @param {object} state 状態オブジェクト
 * @returns {void}
 */
function restoreLastSettings(ui, state) {
    var lastSettings = loadLastSettings();
    if (!lastSettings) return;

    /* 使えないモード（表の一部選択時の見出し行など）は既定の［すべて］のまま / Unavailable modes keep the default "All" */
    if (lastSettings.mode) selectModeRadio(ui, String(lastSettings.mode));

    if (lastSettings.weight != null && String(lastSettings.weight).length > 0) {
        ui.weightInput.text = String(lastSettings.weight);
        syncWeightPresets(ui, state);
    }

    if (lastSettings.color) {
        selectSwatchByName(ui, String(lastSettings.color));
        refreshSwatchControls(ui);
    }

    if (lastSettings.tint != null && !isNaN(Number(lastSettings.tint))) {
        ui.tintInput.text = String(Number(lastSettings.tint));
        commitTintInput(ui);
    }

    if (typeof lastSettings.clearFirst === "boolean") {
        ui.cbClearFirst.value = lastSettings.clearFirst;
    }
}

// =========================================
// イベント / Events
// =========================================

/**
 * ダイアログのコントロールにイベントを結び付ける
 * @param {object} ui UI オブジェクト
 * @param {object} state 状態オブジェクト
 * @returns {void}
 */
function bindDialogEvents(ui, state) {
    bindModeEvents(ui, state);
    bindWeightEvents(ui, state);
    bindColorEvents(ui, state);
    bindShortcutKeys(ui, state);

    ui.cbClearFirst.onClick = function () {
        updatePreview(ui, state);
    };

    ui.btnScreenMode.onClick = function () {
        toggleScreenMode();
        ui.btnScreenMode.text = getScreenModeButtonLabel();
    };

    ui.dlg.onShow = function () {
        ui.btnScreenMode.text = getScreenModeButtonLabel();
        updatePreview(ui, state);
    };
}

/**
 * モードのラジオボタンのクリックを登録する（Option+クリックで［既存の罫線を消してから引く］を切り替え）
 * @param {object} ui UI オブジェクト
 * @param {object} state 状態オブジェクト
 * @returns {void}
 */
function bindModeEvents(ui, state) {
    for (var i = 0; i < BORDER_MODES.length; i++) {
        ui.modeRadios[BORDER_MODES[i].id].onClick = onModeRadioClick;
    }

    /**
     * モードのラジオボタンが押されたときの処理
     * @returns {void}
     */
    function onModeRadioClick() {
        if (ScriptUI.environment.keyboardState.altKey) {
            ui.cbClearFirst.value = !ui.cbClearFirst.value;
        }
        updatePreview(ui, state);
    }
}

/**
 * 線幅の入力欄とプリセットのイベントを登録する
 * @param {object} ui UI オブジェクト
 * @param {object} state 状態オブジェクト
 * @returns {void}
 */
function bindWeightEvents(ui, state) {
    for (var i = 0; i < ui.weightPresets.length; i++) {
        ui.weightPresets[i].radio.onClick = createWeightPresetClickHandler(ui.weightInput, ui.weightPresets[i].text, onWeightChanged);
    }
    ui.weightInput.onChange = onWeightChanged;
    changeValueByArrowKey(ui.weightInput, WEIGHT_ARROW_STEP, onWeightChanged);

    /**
     * 線幅が変わったときの処理
     * @returns {void}
     */
    function onWeightChanged() {
        syncWeightPresets(ui, state);
        updatePreview(ui, state);
    }
}

/**
 * 線幅プリセットのクリック時の処理を作る
 * @param {EditText} weightInput 線幅の入力欄
 * @param {string} presetText プリセットの線幅文字列
 * @param {function} onWeightChanged 線幅を変えたあとに呼ぶ処理
 * @returns {function} クリック時の処理
 */
function createWeightPresetClickHandler(weightInput, presetText, onWeightChanged) {
    return function () {
        weightInput.text = presetText;
        onWeightChanged();
    };
}

/**
 * カラーと濃淡のイベントを登録する
 * @param {object} ui UI オブジェクト
 * @param {object} state 状態オブジェクト
 * @returns {void}
 */
function bindColorEvents(ui, state) {
    ui.swatchDropdown.onChange = function () {
        refreshSwatchControls(ui);
        updatePreview(ui, state);
    };

    ui.tintInput.onChange = onTintInputChanged;
    changeValueByArrowKey(ui.tintInput, TINT_ARROW_STEP, onTintInputChanged);

    ui.tintSlider.onChanging = function () {
        snapTintSlider(ui.tintSlider, false);
        syncTintInputFromSlider(ui);
    };

    ui.tintSlider.onChange = function () {
        snapTintSlider(ui.tintSlider, true);
        syncTintInputFromSlider(ui);
        updatePreview(ui, state);
    };

    /**
     * 濃淡の入力欄が変わったときの処理
     * @returns {void}
     */
    function onTintInputChanged() {
        commitTintInput(ui);
        updatePreview(ui, state);
    }
}

/**
 * モード切り替えと［既存の罫線を消してから引く］のショートカットキーを登録する
 * @param {object} ui UI オブジェクト
 * @param {object} state 状態オブジェクト
 * @returns {void}
 */
function bindShortcutKeys(ui, state) {
    ui.dlg.addEventListener("keydown", function (event) {
        var keyName = String(event.keyName);
        var borderMode;

        if (keyName === CLEAR_FIRST_SHORTCUT_KEY) {
            ui.cbClearFirst.value = !ui.cbClearFirst.value;
        } else {
            borderMode = findModeByShortcutKey(keyName);
            if (!borderMode || !selectModeRadio(ui, borderMode.id)) return;
        }
        event.preventDefault();
        updatePreview(ui, state);
    });
}

/**
 * 数値入力欄に ↑↓ キーでの増減を付ける（Shift 併用で10倍の刻みに揃える）
 * @param {EditText} editText 対象の入力欄
 * @param {number} step 1回の増減量
 * @param {function} onAfterChange 値を変えたあとに呼ぶ処理
 * @returns {void}
 */
function changeValueByArrowKey(editText, step, onAfterChange) {
    var largeStep = step * 10;
    var precision = 1 / step;

    editText.addEventListener("keydown", function (event) {
        var direction = getArrowKeyDirection(event.keyName);
        var value = Number(editText.text);
        if (direction === 0 || isNaN(value)) return;

        if (ScriptUI.environment.keyboardState.shiftKey || event.shiftKey) {
            value = (direction > 0)
                ? Math.floor(value / largeStep) * largeStep + largeStep
                : Math.ceil(value / largeStep) * largeStep - largeStep;
        } else {
            value += direction * step;
        }
        if (value < 0) value = 0;

        editText.text = String(Math.round(value * precision) / precision);
        event.preventDefault();
        onAfterChange();
    });
}

/**
 * キー名から増減の向きを求める（環境によるキー名の違いを吸収）
 * @param {string} keyName イベントから得たキー名
 * @returns {number} 増やすなら 1、減らすなら -1、対象外なら 0
 */
function getArrowKeyDirection(keyName) {
    switch (String(keyName)) {
        case "Up":
        case "UpArrow":
        case "PageUp":
            return 1;
        case "Down":
        case "DownArrow":
        case "PageDown":
            return -1;
        default:
            return 0;
    }
}

// =========================================
// プレビューと確定 / Preview & Apply
// =========================================

/**
 * 罫線を1つの取り消し単位として適用する
 * @param {BorderTarget} target 罫線の適用対象
 * @param {BorderSettings} settings 罫線の設定
 * @param {string} undoName 取り消し履歴に出す名前
 * @returns {void}
 */
function applyBordersAsUndoStep(target, settings, undoName) {
    app.doScript(function () {
        settings.mode.apply(target, settings.stroke, settings.clearFirst);
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, undoName);
}

/**
 * 現在の設定で罫線のプレビューを描き直す
 * @param {object} ui UI オブジェクト
 * @param {object} state 状態オブジェクト
 * @returns {void}
 */
function updatePreview(ui, state) {
    var settings = readBorderSettings(ui, state);

    clearPreview(state);
    if (!canApplySettings(settings)) return;

    try {
        applyBordersAsUndoStep(state.target, settings, getLabel("undo.previewBorders"));
    } catch (e) {
        return;
    }
    state.previewed = true;
    app.activeDocument.recompose();
}

/**
 * プレビューとして適用した罫線を取り消す
 * @param {object} state 状態オブジェクト
 * @returns {void}
 */
function clearPreview(state) {
    if (!state.previewed) return;
    state.previewed = false;
    try {
        app.undo();
    } catch (e) {}
    app.activeDocument.recompose();
}

/**
 * ダイアログの設定を確定して罫線を適用し、次回のために設定を保存する
 * @param {object} ui UI オブジェクト
 * @param {object} state 状態オブジェクト
 * @returns {void}
 */
function applyDialogSettings(ui, state) {
    var settings = readBorderSettings(ui, state);

    clearPreview(state);
    if (!settings.isWeightValid) {
        alert(getLabel("alert.invalidWeight"));
        return;
    }
    if (!settings.isTintValid) {
        alert(getLabel("alert.invalidTint"));
        return;
    }
    if (!canApplySettings(settings)) return;

    applyBordersAsUndoStep(state.target, settings, getLabel("undo.applyBorders"));
    saveLastSettings({
        mode: settings.mode.id,
        weight: settings.weightText,
        color: settings.swatchName,
        tint: settings.stroke.tint,
        clearFirst: settings.clearFirst
    });
}

// =========================================
// 罫線の描画 / Draw borders
// =========================================

/* セルの4辺。線のプロパティ名の接頭辞を兼ねる / The four cell edges, also the prefixes of the stroke property names */
var CELL_EDGES = ["top", "bottom", "left", "right"];

/* 4辺すべてを対象にするフラグ / Flags that select all four edges */
var ALL_EDGE_FLAGS = { top: true, bottom: true, left: true, right: true };

/**
 * セルの1辺に線幅・カラー・濃淡を設定する
 * @param {Cell} cell 対象のセル
 * @param {string} edgeName "top" / "bottom" / "left" / "right"
 * @param {BorderStroke} stroke 設定する値
 * @returns {void}
 */
function setEdgeStroke(cell, edgeName, stroke) {
    cell[edgeName + "EdgeStrokeWeight"] = stroke.weight;
    cell[edgeName + "EdgeStrokeColor"] = stroke.swatch;
    cell[edgeName + "EdgeStrokeTint"] = stroke.tint;
}

/**
 * セルの1辺の罫線を消去する（線幅 0、カラーなし、濃淡 100）
 * @param {Cell} cell 対象のセル
 * @param {string} edgeName "top" / "bottom" / "left" / "right"
 * @returns {void}
 */
function clearEdgeStroke(cell, edgeName) {
    cell[edgeName + "EdgeStrokeWeight"] = 0;
    try {
        cell[edgeName + "EdgeStrokeColor"] = NothingEnum.NOTHING;
        cell[edgeName + "EdgeStrokeTint"] = 100;
    } catch (e) {}
}

/**
 * セルの4辺の線幅を 0 にする
 * @param {CellEntry[]} entries 対象のセル
 * @returns {void}
 */
function zeroEdgeWeights(entries) {
    for (var i = 0; i < entries.length; i++) {
        for (var j = 0; j < CELL_EDGES.length; j++) {
            entries[i].cell[CELL_EDGES[j] + "EdgeStrokeWeight"] = 0;
        }
    }
}

/**
 * セルごとに選んだ辺へ罫線を引く
 * @param {CellEntry[]} entries 対象のセル
 * @param {BorderStroke} stroke 設定する値
 * @param {boolean} clearFirst 先に4辺の線幅を 0 にするか
 * @param {function(CellEntry): object} pickEdges 引く辺を { top, bottom, left, right } のフラグで返す関数
 * @returns {void}
 */
function strokeCellEdges(entries, stroke, clearFirst, pickEdges) {
    var i, j, edgeFlags;
    if (clearFirst) zeroEdgeWeights(entries);
    for (i = 0; i < entries.length; i++) {
        edgeFlags = pickEdges(entries[i]);
        for (j = 0; j < CELL_EDGES.length; j++) {
            if (edgeFlags[CELL_EDGES[j]]) setEdgeStroke(entries[i].cell, CELL_EDGES[j], stroke);
        }
    }
}

/**
 * セルごとに選んだ辺の罫線を消去する
 * @param {CellEntry[]} entries 対象のセル
 * @param {function(CellEntry): object} pickEdges 消す辺を { top, bottom, left, right } のフラグで返す関数
 * @returns {void}
 */
function clearCellEdges(entries, pickEdges) {
    var i, j, edgeFlags;
    for (i = 0; i < entries.length; i++) {
        edgeFlags = pickEdges(entries[i]);
        for (j = 0; j < CELL_EDGES.length; j++) {
            if (edgeFlags[CELL_EDGES[j]]) clearEdgeStroke(entries[i].cell, CELL_EDGES[j]);
        }
    }
}

/**
 * セルが選択範囲のどの辺に接しているかを求める
 * @param {CellRange} range セルの占有範囲
 * @param {CellBounds} bounds 選択範囲の行・列の範囲
 * @returns {object} 各辺に接しているかを示すフラグ
 */
function getBoundaryEdgeFlags(range, bounds) {
    return {
        top: range.startRow === bounds.minRow,
        bottom: range.endRow === bounds.maxRow,
        left: range.startCol === bounds.minCol,
        right: range.endCol === bounds.maxCol
    };
}

/**
 * 指定した側の隣に、同じ集合のセルがあるかを判定する
 * @param {CellEntry} entry 対象のセル
 * @param {CellEntry[]} entries 隣を探すセルの集合
 * @param {string} side "left" / "right" / "bottom"
 * @returns {boolean} 隣にセルがあれば true
 */
function hasSelectedNeighbor(entry, entries, side) {
    var baseRange = entry.range;
    var i, otherRange;

    for (i = 0; i < entries.length; i++) {
        if (entries[i] === entry) continue;
        otherRange = entries[i].range;

        if (side === "bottom") {
            if (spansOverlap(baseRange.startCol, baseRange.endCol, otherRange.startCol, otherRange.endCol) &&
                baseRange.endRow + 1 === otherRange.startRow) return true;
        } else if (spansOverlap(baseRange.startRow, baseRange.endRow, otherRange.startRow, otherRange.endRow)) {
            if (side === "left" && otherRange.endCol + 1 === baseRange.startCol) return true;
            if (side === "right" && baseRange.endCol + 1 === otherRange.startCol) return true;
        }
    }
    return false;
}

/**
 * ［すべて］選択範囲のすべての罫線を引く
 * @param {BorderTarget} target 罫線の適用対象
 * @param {BorderStroke} stroke 設定する値
 * @returns {void}
 */
function applyAllBorders(target, stroke) {
    strokeCellEdges(target.blockEntries, stroke, false, function () {
        return ALL_EDGE_FLAGS;
    });
}

/**
 * ［外枠のみ］選択範囲の外周だけに罫線を引く
 * @param {BorderTarget} target 罫線の適用対象
 * @param {BorderStroke} stroke 設定する値
 * @param {boolean} clearFirst 描画前に既存の罫線を消すか
 * @returns {void}
 */
function applyOuterBorders(target, stroke, clearFirst) {
    strokeCellEdges(target.blockEntries, stroke, clearFirst, function (entry) {
        return getBoundaryEdgeFlags(entry.range, target.bounds);
    });
}

/**
 * ［内側のみ］選択範囲の内側だけに罫線を引く
 * @param {BorderTarget} target 罫線の適用対象
 * @param {BorderStroke} stroke 設定する値
 * @param {boolean} clearFirst 描画前に既存の罫線を消すか
 * @returns {void}
 */
function applyInnerBorders(target, stroke, clearFirst) {
    var entries = target.blockEntries;
    strokeCellEdges(entries, stroke, clearFirst, function (entry) {
        return {
            bottom: hasSelectedNeighbor(entry, entries, "bottom"),
            right: hasSelectedNeighbor(entry, entries, "right")
        };
    });
}

/**
 * ［水平線のみ］外周の上下を含む水平方向の罫線を引く
 * @param {BorderTarget} target 罫線の適用対象
 * @param {BorderStroke} stroke 設定する値
 * @param {boolean} clearFirst 描画前に既存の罫線を消すか
 * @returns {void}
 */
function applyHorizontalBorders(target, stroke, clearFirst) {
    var entries = target.blockEntries;
    strokeCellEdges(entries, stroke, clearFirst, function (entry) {
        var boundaryFlags = getBoundaryEdgeFlags(entry.range, target.bounds);
        return {
            top: boundaryFlags.top,
            bottom: boundaryFlags.bottom || hasSelectedNeighbor(entry, entries, "bottom")
        };
    });
}

/**
 * ［垂直線のみ］外周の左右を含む垂直方向の罫線を引く
 * @param {BorderTarget} target 罫線の適用対象
 * @param {BorderStroke} stroke 設定する値
 * @param {boolean} clearFirst 描画前に既存の罫線を消すか
 * @returns {void}
 */
function applyVerticalBorders(target, stroke, clearFirst) {
    var entries = target.blockEntries;
    strokeCellEdges(entries, stroke, clearFirst, function (entry) {
        var boundaryFlags = getBoundaryEdgeFlags(entry.range, target.bounds);
        return {
            left: boundaryFlags.left,
            right: boundaryFlags.right || hasSelectedNeighbor(entry, entries, "right")
        };
    });
}

/**
 * ［下端のみ］選択範囲の最下辺だけに罫線を引く
 * @param {BorderTarget} target 罫線の適用対象
 * @param {BorderStroke} stroke 設定する値
 * @param {boolean} clearFirst 描画前に既存の罫線を消すか
 * @returns {void}
 */
function applyBottomEdgeBorder(target, stroke, clearFirst) {
    strokeCellEdges(target.blockEntries, stroke, clearFirst, function (entry) {
        return { bottom: getBoundaryEdgeFlags(entry.range, target.bounds).bottom };
    });
}

/**
 * ［右端のみ］選択範囲の最右辺だけに罫線を引く
 * @param {BorderTarget} target 罫線の適用対象
 * @param {BorderStroke} stroke 設定する値
 * @param {boolean} clearFirst 描画前に既存の罫線を消すか
 * @returns {void}
 */
function applyRightEdgeBorder(target, stroke, clearFirst) {
    strokeCellEdges(target.blockEntries, stroke, clearFirst, function (entry) {
        return { right: getBoundaryEdgeFlags(entry.range, target.bounds).right };
    });
}

/**
 * ［見出し行］表の上端・先頭行の下・表の下端に罫線を引く
 * @param {BorderTarget} target 罫線の適用対象
 * @param {BorderStroke} stroke 設定する値
 * @param {boolean} clearFirst 描画前に既存の罫線を消すか
 * @returns {void}
 */
function applyHeaderRowBorders(target, stroke, clearFirst) {
    strokeCellEdges(target.selectedEntries, stroke, clearFirst, function (entry) {
        var boundaryFlags = getBoundaryEdgeFlags(entry.range, target.bounds);
        return {
            top: boundaryFlags.top,
            bottom: boundaryFlags.top || boundaryFlags.bottom
        };
    });
}

/**
 * ［見出し列］表の左端・先頭列の右・表の右端に罫線を引く
 * @param {BorderTarget} target 罫線の適用対象
 * @param {BorderStroke} stroke 設定する値
 * @param {boolean} clearFirst 描画前に既存の罫線を消すか
 * @returns {void}
 */
function applyHeaderColumnBorders(target, stroke, clearFirst) {
    strokeCellEdges(target.selectedEntries, stroke, clearFirst, function (entry) {
        var boundaryFlags = getBoundaryEdgeFlags(entry.range, target.bounds);
        return {
            left: boundaryFlags.left,
            right: boundaryFlags.left || boundaryFlags.right
        };
    });
}

/**
 * ［左右の外枠を消去］選択ブロックの左端と右端の罫線だけを消去する
 * @param {BorderTarget} target 罫線の適用対象
 * @returns {void}
 */
function clearOuterLeftRightBorders(target) {
    var entries = target.selectedEntries;
    clearCellEdges(entries, function (entry) {
        return {
            left: !hasSelectedNeighbor(entry, entries, "left"),
            right: !hasSelectedNeighbor(entry, entries, "right")
        };
    });
}

/**
 * ［すべて消去］選択範囲のすべての罫線を消去する
 * @param {BorderTarget} target 罫線の適用対象
 * @returns {void}
 */
function clearAllBorders(target) {
    clearCellEdges(target.blockEntries, function () {
        return ALL_EDGE_FLAGS;
    });
}

// =========================================
// メイン処理 / Main
// =========================================

/**
 * 選択セルを調べてダイアログを開き、確定した罫線を適用する
 * @returns {void}
 */
function main() {
    var originalSelection = app.selection;
    var selectedEntries = collectSelectedCellEntries(originalSelection);
    if (selectedEntries.length === 0) {
        alert(getLabel("alert.noCellSelection"));
        return;
    }

    var state = {
        target: buildBorderTarget(selectedEntries),
        strokeUnit: getStrokeUnitInfo(),
        isWholeTableSelected: isWholeTableSelected(selectedEntries),
        previewed: false
    };

    app.selection = NothingEnum.NOTHING;

    var ui = buildDialog(state);
    bindDialogEvents(ui, state);
    restoreLastSettings(ui, state);

    if (ui.dlg.show() === 1) {
        applyDialogSettings(ui, state);
    } else {
        clearPreview(state);
    }
    restoreSelection(originalSelection);
}

main();

})();
