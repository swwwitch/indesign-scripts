#target indesign

/*

### 概要

選択した表セルに、選択範囲内の行の並びを基準に交互の塗り（縞模様）をプレビュー付きで適用します。
上から／左から指定した数の行・列は、塗りの対象外にできます。

詳細は README を参照してください。

### Overview

Applies alternating fills (zebra striping) to the selected table cells, based on the row order within the selection, with a live preview.
A given number of rows from the top, and columns from the left, can be left out of the fill.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdZebraRowFill";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-16";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdZebraRowFill.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdZebraRowFill.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n20ff60f6b508"; /* 紹介記事 / article URL */

// Original idea
// KK sawa

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 濃淡（Tint）の初期値と範囲 / Initial value and range of the tint control */
var TINT_DEFAULT = 100;
var TINT_MIN     = 0;
var TINT_MAX     = 100;

/* 濃淡スライダーの刻み幅（通常 / Shift / Option）/ Tint slider steps (normal, Shift, Option) */
var TINT_STEP_NORMAL = 1;
var TINT_STEP_SHIFT  = 10;
var TINT_STEP_OPTION = 5;

/* ↑↓キー1回の増減量。Shift 併用時はこの10倍の刻みに揃える / Arrow-key step; with Shift the value snaps to 10x this step */
var TINT_ARROW_STEP = 1;
var SKIP_ARROW_STEP = 1;

/* ［行のスキップ］［列のスキップ］の初期値 / Initial number of rows and columns to skip */
var SKIP_COUNT_DEFAULT = 1;

// =========================================
// レイアウト設定 / Layout settings
// =========================================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 10;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 20;                 /* 2カラムの間隔 / gap between columns */

/* 行パネル内の間隔と、濃淡グループの余白 [左,上,右,下] / Spacing inside a row panel and margins of its tint group */
var ROW_PANEL_SPACING  = 6;
var TINT_GROUP_MARGINS = [0, 20, 0, 10];

/* 入力欄の行・ボタン同士の間隔と、項目名と色見本の間隔 / Spacing inside control rows and next to the swatch chip */
var CONTROL_SPACING    = 8;
var SWATCH_ROW_SPACING = 6;

/* カラードロップダウン・濃淡入力欄・スライダー・スキップ入力欄の幅（px）
   / Widths of the color dropdown, tint field, slider and skip fields (px) */
var COLOR_DROPDOWN_WIDTH = 140;
var TINT_INPUT_WIDTH     = 50;
var TINT_SLIDER_WIDTH    = 150;
var SKIP_INPUT_WIDTH     = 30;

/* スキップのチェックボックスの幅。行と列で入力欄の位置をそろえる（px）
   / Width of the skip checkboxes so both fields line up (px) */
var SKIP_CHECKBOX_WIDTH = 80;

/* 色見本の一辺（px）/ Size of the swatch chip (px) */
var SWATCH_CHIP_SIZE = 18;

/* ボタン列の上余白と、左右を分けるスペーサーの最小幅（px）/ Top margin of the button row and minimum width of its spacer (px) */
var BUTTON_ROW_TOP_MARGIN       = 8;
var BUTTON_ROW_SPACER_MIN_WIDTH = 40;

/* ダイアログの不透明度 / Dialog opacity */
var DIALOG_OPACITY = 0.97;

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
        title: { ja: "行の塗りを交互に設定", en: "Alternating Row Fill" }
    },
    panel: {
        oddRows:  { ja: "奇数番目の行", en: "Odd-numbered Rows" },
        evenRows: { ja: "偶数番目の行", en: "Even-numbered Rows" },
        options:  { ja: "オプション", en: "Options" },
        skip:     { ja: "スキップ", en: "Skip" }
    },
    fieldLabel: {
        color: { ja: "カラー", en: "Color" },
        tint:  { ja: "濃淡", en: "Tint" }
    },
    checkbox: {
        swapRows:    { ja: "奇数行と偶数行を入れ替え", en: "Swap odd and even" },
        skipRows:    { ja: "行", en: "Rows" },
        skipColumns: { ja: "列", en: "Columns" }
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
        oddRows: {
            ja: "選択範囲の上から数えて1・3・5…番目の行です。除外した行は数に入りません。",
            en: "The 1st, 3rd, 5th and so on row counted from the top of the selection. Skipped rows are not counted."
        },
        evenRows: {
            ja: "選択範囲の上から数えて2・4・6…番目の行です。除外した行は数に入りません。",
            en: "The 2nd, 4th, 6th and so on row counted from the top of the selection. Skipped rows are not counted."
        },
        colorDropdown: {
            ja: "「なし」と「紙色」では濃淡を指定できません。",
            en: "Tint is not available for None or Paper."
        },
        tintInput: {
            ja: "↑↓キーで1ずつ、Shift+↑↓キーで10ずつ増減します。",
            en: "Up/Down arrow keys change the value by 1, or by 10 with Shift."
        },
        tintSlider: {
            ja: "Shift+ドラッグで10%刻み、Option（Alt）+ドラッグで5%刻みになります。",
            en: "Shift-drag snaps to 10% steps, Option (Alt)-drag to 5% steps."
        },
        swapRows: {
            ja: "奇数行と偶数行のカラー・濃淡を入れ替えます。",
            en: "Exchanges the color and tint between the odd and even rows."
        },
        skipRows: {
            ja: "選択範囲の上から指定した行数を塗りの対象から外し、実行前のカラーに戻します。",
            en: "Leaves the given number of rows at the top of the selection out of the fill, restoring the color they had before the script ran."
        },
        skipColumns: {
            ja: "選択範囲の左から指定した列数を塗りの対象から外し、実行前のカラーに戻します。",
            en: "Leaves the given number of columns at the left of the selection out of the fill, restoring the color they had before the script ran."
        },
        screenMode: {
            ja: "ドキュメントウィンドウの画面モードを、標準モードとプレビューで切り替えます。",
            en: "Switches the document window between the Normal and Preview screen modes."
        }
    },
    alert: {
        noDocument:      { ja: "ドキュメントを開いてください。", en: "Please open a document." },
        noCellSelection: { ja: "セルを選択してください。", en: "Please select table cells." },
        noUsableSwatch:  { ja: "使用可能なカラーがありません。", en: "No usable colors are available." }
    },
    undo: {
        previewFills: { ja: "塗りプレビュー", en: "Fill Preview" },
        applyFills:   { ja: "セルの塗りを交互に設定", en: "Apply Alternating Fills" }
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
// スウォッチ / Swatches
// =========================================

/* 特別扱いするスウォッチ名（角括弧・大文字小文字は無視して照合）/ Special swatch names (brackets and case are ignored) */
var SPECIAL_SWATCH_NAMES = {
    none:         ["None", "なし"],
    black:        ["Black", "ブラック", "黒"],
    paper:        ["Paper", "紙色", "紙"],
    registration: ["Registration", "レジストレーション", "トンボ用"]
};

/* 色見本に使う RGB（0〜1）/ RGB values (0-1) used for the swatch chip */
var CHIP_RGB_BLACK    = [0, 0, 0];
var CHIP_RGB_WHITE    = [1, 1, 1];
var CHIP_RGB_FALLBACK = [0.5, 0.5, 0.5];

/**
 * @typedef {object} SwatchChoice
 * @property {string} swatchName ドキュメント上のスウォッチ名
 * @property {string} displayName ドロップダウンに表示する名前
 * @property {string} swatchKind 特別なスウォッチの種別。通常のカラーは空文字
 */

/**
 * 比較しやすいようにスウォッチ名を正規化する（角括弧・前後の空白・大文字小文字を落とす）
 * @param {string} swatchName スウォッチ名
 * @returns {string} 正規化した名前
 */
function normalizeSwatchName(swatchName) {
    if (swatchName == null) return "";
    return String(swatchName).replace(/^\[|\]$/g, "").replace(/^\s+|\s+$/g, "").toLowerCase();
}

/**
 * スウォッチ名が特別なスウォッチのどれに当たるかを調べる
 * @param {string} swatchName スウォッチ名
 * @returns {string} "none" / "black" / "paper" / "registration"。どれでもなければ空文字
 */
function getSpecialSwatchKind(swatchName) {
    var normalizedName = normalizeSwatchName(swatchName);
    var swatchKind, names, i;
    for (swatchKind in SPECIAL_SWATCH_NAMES) {
        if (!SPECIAL_SWATCH_NAMES.hasOwnProperty(swatchKind)) continue;
        names = SPECIAL_SWATCH_NAMES[swatchKind];
        for (i = 0; i < names.length; i++) {
            if (normalizeSwatchName(names[i]) === normalizedName) return swatchKind;
        }
    }
    return "";
}

/**
 * 濃淡を指定できないスウォッチかどうかを判定する
 * @param {string} swatchKind getSpecialSwatchKind() が返した種別
 * @returns {boolean} 濃淡を指定できないなら true
 */
function isTintDisabledKind(swatchKind) {
    return swatchKind === "none" || swatchKind === "paper";
}

/**
 * カラー候補として表示するスウォッチ一覧を作る（レジストレーションは除く）
 * @returns {Array<SwatchChoice>} スウォッチ一覧
 */
function getSwatchChoices() {
    var swatches = app.activeDocument.swatches;
    var swatchChoices = [];
    var i, swatchName, swatchKind;

    for (i = 0; i < swatches.length; i++) {
        swatchName = String(swatches[i].name);
        swatchKind = getSpecialSwatchKind(swatchName);
        if (swatchKind === "registration") continue;
        swatchChoices.push({
            swatchName: swatchName,
            displayName: swatchKind ? getLabel("swatch." + swatchKind) : swatchName,
            swatchKind: swatchKind
        });
    }
    return swatchChoices;
}

/**
 * 名前からスウォッチを取得する
 * @param {string} swatchName スウォッチ名
 * @returns {Swatch|null} 見つかったスウォッチ。無ければ null
 */
function findSwatch(swatchName) {
    var swatch = app.activeDocument.swatches.itemByName(String(swatchName));
    return (swatch && swatch.isValid) ? swatch : null;
}

/**
 * 一覧の中からスウォッチ名に一致する位置を求める
 * @param {Array<SwatchChoice>} swatchChoices スウォッチ一覧
 * @param {string} swatchName 探すスウォッチ名
 * @returns {number} 見つかった位置。無ければ 0
 */
function findSwatchChoiceIndex(swatchChoices, swatchName) {
    for (var i = 0; i < swatchChoices.length; i++) {
        if (swatchChoices[i].swatchName === swatchName) return i;
    }
    return 0;
}

/**
 * スウォッチのカラー値を色見本用の RGB に変換する
 * @param {Swatch|null} swatch 対象のスウォッチ
 * @returns {Array<number>} 0〜1 の RGB 値
 */
function getSwatchChipRGB(swatch) {
    if (!swatch || !swatch.isValid) return CHIP_RGB_FALLBACK;

    var swatchKind = getSpecialSwatchKind(String(swatch.name));
    if (swatchKind === "paper" || swatchKind === "none") return CHIP_RGB_WHITE;
    if (swatchKind === "black" || swatchKind === "registration") return CHIP_RGB_BLACK;

    /* グラデーションや混合インキは色値を読めない / Gradients and mixed inks have no readable color value */
    try {
        var colorValues = swatch.colorValue;
        if (swatch.space === ColorSpace.RGB) {
            return [colorValues[0] / 255, colorValues[1] / 255, colorValues[2] / 255];
        }
        if (swatch.space === ColorSpace.CMYK) {
            var cyanValue    = colorValues[0] / 100;
            var magentaValue = colorValues[1] / 100;
            var yellowValue  = colorValues[2] / 100;
            var blackValue   = colorValues[3] / 100;
            return [
                (1 - cyanValue) * (1 - blackValue),
                (1 - magentaValue) * (1 - blackValue),
                (1 - yellowValue) * (1 - blackValue)
            ];
        }
    } catch (e) { }
    return CHIP_RGB_FALLBACK;
}

/**
 * 色見本をスウォッチのカラーで塗り直す
 * @param {Group} swatchChip 色見本のグループ
 * @param {Swatch|null} swatch 表示するスウォッチ
 * @returns {void}
 */
function paintSwatchChip(swatchChip, swatch) {
    var rgb = getSwatchChipRGB(swatch);
    swatchChip.graphics.backgroundColor = swatchChip.graphics.newBrush(
        swatchChip.graphics.BrushType.SOLID_COLOR,
        [rgb[0], rgb[1], rgb[2], 1]
    );
    if (swatchChip.window) swatchChip.window.update();
}

// =========================================
// 画面モードの切り替え / Screen mode
// =========================================

/**
 * ドキュメントウィンドウがプレビュー表示かどうかを判定する
 * @returns {boolean} プレビュー表示なら true
 */
function isPreviewScreenMode() {
    /* ストーリーエディターがアクティブだと screenMode を持たない / A Story window has no screenMode */
    try {
        var docWindow = app.activeWindow;
        return !!(docWindow && docWindow.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE);
    } catch (e) {
        return false;
    }
}

/**
 * 標準表示とプレビュー表示を切り替える
 * @returns {void}
 */
function toggleScreenMode() {
    try {
        var docWindow = app.activeWindow;
        if (!docWindow) return;
        docWindow.screenMode = isPreviewScreenMode()
            ? ScreenModeOptions.PREVIEW_OFF
            : ScreenModeOptions.PREVIEW_TO_PAGE;
    } catch (e) { }
}

/**
 * 画面モードの切り替えボタンに表示する文字列を返す
 * @returns {string} ボタンに表示する文字列
 */
function getScreenModeButtonLabel() {
    return getLabel(isPreviewScreenMode() ? "button.screenModeNormal" : "button.screenModePreview");
}

// =========================================
// 数値の正規化とキー操作 / Value handling
// =========================================

/**
 * 修飾キーに応じた濃淡スライダーの刻み幅を返す
 * @returns {number} 刻み幅
 */
function getTintSliderStep() {
    var keyboard = ScriptUI.environment.keyboardState;
    if (keyboard.shiftKey) return TINT_STEP_SHIFT;
    if (keyboard.altKey) return TINT_STEP_OPTION;
    return TINT_STEP_NORMAL;
}

/**
 * 濃淡を有効範囲の整数に収める
 * @param {number} value 入力された値
 * @returns {number} 整えた値
 */
function normalizeTint(value) {
    if (isNaN(value)) value = TINT_DEFAULT;
    if (value < TINT_MIN) value = TINT_MIN;
    if (value > TINT_MAX) value = TINT_MAX;
    return Math.round(value);
}

/**
 * 濃淡スライダーの値を、修飾キーに応じた刻み幅に合わせる
 * @param {number} value スライダーの値
 * @returns {number} 整えた値
 */
function snapTintToSliderStep(value) {
    var step = getTintSliderStep();
    return normalizeTint(Math.round(normalizeTint(value) / step) * step);
}

/**
 * スキップ数を 0 以上の整数に整える
 * @param {number} value 入力された値
 * @returns {number} 整えた値
 */
function normalizeSkipCount(value) {
    if (isNaN(value) || value < 0) return 0;
    return Math.round(value);
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

/**
 * 数値入力欄に↑↓キーでの増減を付ける
 * @param {EditText} editText 対象の入力欄
 * @param {number} step ↑↓キー1回の増減量。Shift 併用時はこの10倍の刻みに揃える
 * @param {function} normalizeValue 値を有効範囲に収める関数
 * @param {function} onAfterChange 値の変更後に呼ぶ処理
 * @returns {void}
 */
function changeValueByArrowKey(editText, step, normalizeValue, onAfterChange) {
    var largeStep = step * 10;

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

        editText.text = String(normalizeValue(value));
        event.preventDefault();
        onAfterChange();
    });
}

// =========================================
// 選択セルと初期値 / Selected cells & defaults
// =========================================

/**
 * @typedef {object} CellFill
 * @property {Swatch|null} color 塗りのカラー
 * @property {number|null} tint 濃淡。指定しない場合は null
 */

/**
 * 選択中の表セルを配列で取得する
 * @returns {Array<Cell>} 選択していたセル。表以外を選んでいる場合は空配列
 */
function getSelectedCells() {
    if (app.selection.length === 0) return [];

    /* 表以外を選んでいると cells を持たない / A non-table selection has no cells */
    var selectedCells;
    try {
        selectedCells = app.selection[0].cells;
    } catch (e) {
        return [];
    }

    var cells = [];
    for (var i = 0; selectedCells && i < selectedCells.length; i++) {
        cells.push(selectedCells[i]);
    }
    return cells;
}

/**
 * セルの現在の塗りを控える（スキップ時に戻すため）
 * @param {Cell} cell 対象のセル
 * @returns {CellFill} 控えた塗り
 */
function captureCellFill(cell) {
    var fillColor = cell.fillColor;
    var fillTint = cell.fillTint;
    var swatchKind = (fillColor && fillColor.name) ? getSpecialSwatchKind(String(fillColor.name)) : "";
    var hasTint = !isTintDisabledKind(swatchKind) && !isNaN(fillTint) &&
        fillTint >= TINT_MIN && fillTint <= TINT_MAX;

    return { color: fillColor, tint: hasTint ? fillTint : null };
}

/**
 * 選択セルすべての塗りを控える
 * @param {Array<Cell>} cells 対象のセル
 * @returns {Array<CellFill>} セルと同じ並びの塗り
 */
function captureCellFills(cells) {
    var fills = [];
    for (var i = 0; i < cells.length; i++) {
        fills.push(captureCellFill(cells[i]));
    }
    return fills;
}

/**
 * 選択セルの塗り（カラー＋濃淡）を、使われている数の多い順に並べる
 * @param {Array<Cell>} cells 対象のセル
 * @param {{colorName: string, tint: number}} fallbackFill 塗りを読めなかったときに使う値
 * @returns {Array<{colorName: string, tint: number, count: number}>} 多い順の組み合わせ
 */
function countFillUsages(cells, fallbackFill) {
    var usageByKey = {};
    var usages = [];
    var i, fillColor, fillTint, colorName, tint, usageKey;

    for (i = 0; i < cells.length; i++) {
        fillColor = cells[i].fillColor;
        fillTint = cells[i].fillTint;
        colorName = (fillColor && fillColor.name) ? String(fillColor.name) : fallbackFill.colorName;
        tint = (!isNaN(fillTint) && fillTint >= TINT_MIN && fillTint <= TINT_MAX) ? fillTint : fallbackFill.tint;

        usageKey = colorName + "||" + tint;
        if (!usageByKey[usageKey]) {
            usageByKey[usageKey] = { colorName: colorName, tint: tint, count: 0 };
            usages.push(usageByKey[usageKey]);
        }
        usageByKey[usageKey].count++;
    }

    usages.sort(function (a, b) { return b.count - a.count; });
    return usages;
}

/**
 * 選択セルでよく使われている塗りから、奇数行・偶数行の初期値を決める
 * @param {Array<Cell>} cells 対象のセル
 * @param {string} fallbackColorName 塗りを読めなかったときに使うスウォッチ名
 * @returns {{odd: {colorName: string, tint: number}, even: {colorName: string, tint: number}}} 初期値
 */
function detectDefaultFills(cells, fallbackColorName) {
    var fallbackFill = { colorName: fallbackColorName, tint: TINT_DEFAULT };
    var usages = countFillUsages(cells, fallbackFill);
    var mostUsed = (usages.length > 0) ? usages[0] : fallbackFill;

    return {
        odd: mostUsed,
        even: (usages.length > 1) ? usages[1] : mostUsed
    };
}

// =========================================
// ダイアログの組み立て / Dialog
// =========================================

/**
 * @typedef {object} RowFillControls
 * @property {Panel} panel パネル本体
 * @property {DropDownList} colorDropdown カラーのドロップダウン
 * @property {Group} swatchChip 選択中のカラーを示す色見本
 * @property {EditText} tintInput 濃淡の入力欄
 * @property {Slider} tintSlider 濃淡のスライダー
 * @property {Array<SwatchChoice>} swatchChoices ドロップダウンと同じ並びのスウォッチ一覧
 */

/**
 * 行のカラーと濃淡を設定するパネルを作る（奇数行・偶数行で共通）
 * @param {Group} parent 追加先のグループ
 * @param {string} titleKey パネル名のラベルキー
 * @param {string} tooltipKey パネルに付けるツールチップのキー
 * @param {Array<SwatchChoice>} swatchChoices スウォッチ一覧
 * @param {{colorName: string, tint: number}} defaultFill カラーと濃淡の初期値
 * @param {function} onChange 値が変わったときに呼ぶ処理
 * @returns {RowFillControls} 作成したコントロール一式
 */
function createRowFillPanel(parent, titleKey, tooltipKey, swatchChoices, defaultFill, onChange) {
    var rowFillPanel = parent.add("panel", undefined, getLabel(titleKey));
    setupPanel(rowFillPanel, ROW_PANEL_SPACING);
    rowFillPanel.alignChildren = ["left", "top"];
    rowFillPanel.helpTip = getLabel(tooltipKey);

    var colorGroup = rowFillPanel.add("group");
    colorGroup.orientation = "column";
    colorGroup.alignChildren = ["left", "top"];

    var colorLabelRow = colorGroup.add("group");
    setupRow(colorLabelRow, "left", SWATCH_ROW_SPACING);
    colorLabelRow.add("statictext", undefined, labelText("fieldLabel.color"));

    var swatchChip = colorLabelRow.add("group");
    swatchChip.preferredSize = [SWATCH_CHIP_SIZE, SWATCH_CHIP_SIZE];
    swatchChip.minimumSize   = [SWATCH_CHIP_SIZE, SWATCH_CHIP_SIZE];
    swatchChip.maximumSize   = [SWATCH_CHIP_SIZE, SWATCH_CHIP_SIZE];

    var displayNames = [];
    for (var i = 0; i < swatchChoices.length; i++) {
        displayNames.push(swatchChoices[i].displayName);
    }
    var colorDropdown = colorGroup.add("dropdownlist", undefined, displayNames);
    colorDropdown.preferredSize.width = COLOR_DROPDOWN_WIDTH;
    colorDropdown.helpTip = getLabel("tooltip.colorDropdown");
    colorDropdown.selection = findSwatchChoiceIndex(swatchChoices, defaultFill.colorName);

    var tintGroup = rowFillPanel.add("group");
    tintGroup.orientation = "column";
    tintGroup.alignChildren = ["left", "top"];
    tintGroup.margins = TINT_GROUP_MARGINS;

    var tintInputRow = tintGroup.add("group");
    setupRow(tintInputRow, "left", CONTROL_SPACING);
    tintInputRow.add("statictext", undefined, labelText("fieldLabel.tint"));

    var tintInput = tintInputRow.add("edittext", undefined, String(defaultFill.tint));
    tintInput.preferredSize.width = TINT_INPUT_WIDTH;
    tintInput.helpTip = getLabel("tooltip.tintInput");

    var tintSlider = tintGroup.add("slider", undefined, defaultFill.tint, TINT_MIN, TINT_MAX);
    tintSlider.preferredSize.width = TINT_SLIDER_WIDTH;
    tintSlider.helpTip = getLabel("tooltip.tintSlider");

    var rowFill = {
        panel: rowFillPanel,
        colorDropdown: colorDropdown,
        swatchChip: swatchChip,
        tintInput: tintInput,
        tintSlider: tintSlider,
        swatchChoices: swatchChoices
    };

    /* ハンドラーは初期値を入れ終えてから付ける（selection への代入で発火するため）
       / Attach the handlers after the initial values are in place */
    colorDropdown.onChange = function () {
        refreshRowFill(rowFill);
        onChange();
    };
    tintSlider.onChanging = function () {
        var tint = snapTintToSliderStep(this.value);
        this.value = tint;
        tintInput.text = String(tint);
        onChange();
    };
    tintInput.onChange = function () {
        setRowFillTint(rowFill, normalizeTint(parseFloat(this.text)));
        onChange();
    };
    changeValueByArrowKey(tintInput, TINT_ARROW_STEP, normalizeTint, function () {
        setRowFillTint(rowFill, normalizeTint(parseFloat(tintInput.text)));
        onChange();
    });

    return rowFill;
}

/**
 * パネルで選択中のスウォッチを取得する
 * @param {RowFillControls} rowFill 対象のパネル
 * @returns {SwatchChoice} 選択中のスウォッチ
 */
function getRowFillChoice(rowFill) {
    var selectedIndex = rowFill.colorDropdown.selection ? rowFill.colorDropdown.selection.index : 0;
    return rowFill.swatchChoices[selectedIndex];
}

/**
 * パネルの濃淡を入力欄とスライダーの両方に反映する
 * @param {RowFillControls} rowFill 対象のパネル
 * @param {number} tint 設定する濃淡
 * @returns {void}
 */
function setRowFillTint(rowFill, tint) {
    rowFill.tintInput.text = String(tint);
    rowFill.tintSlider.value = tint;
}

/**
 * 選択中のカラーに合わせて、色見本と濃淡コントロールの状態を更新する
 * @param {RowFillControls} rowFill 対象のパネル
 * @returns {void}
 */
function refreshRowFill(rowFill) {
    var swatchChoice = getRowFillChoice(rowFill);
    var tintEnabled = !isTintDisabledKind(swatchChoice.swatchKind);

    rowFill.tintInput.enabled = tintEnabled;
    rowFill.tintSlider.enabled = tintEnabled;
    paintSwatchChip(rowFill.swatchChip, findSwatch(swatchChoice.swatchName));
}

/**
 * オプションのパネルとスキップのパネルを追加する
 * @param {Window} dialog 対象のダイアログ
 * @param {object} ui UI オブジェクト
 * @param {function} onChange 値が変わったときに呼ぶ処理
 * @returns {void}
 */
function addOptions(dialog, ui, onChange) {
    /* 左にオプション、右にスキップの2カラム / Two columns: options on the left, skip on the right */
    var optionPanelsGroup = dialog.add("group");
    setupRow(optionPanelsGroup, "fill", COLUMN_SPACING);
    optionPanelsGroup.alignChildren = ["left", "top"];

    var optionsPanel = optionPanelsGroup.add("panel", undefined, getLabel("panel.options"));
    setupPanel(optionsPanel, ROW_PANEL_SPACING);
    optionsPanel.alignChildren = ["left", "top"];

    ui.swapCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.swapRows"));
    ui.swapCheckbox.helpTip = getLabel("tooltip.swapRows");
    ui.swapCheckbox.value = false;
    ui.swapCheckbox.onClick = onChange;

    /* 選択範囲の上から／左から、指定数を塗りの対象外にする / Leave the given rows and columns out of the fill */
    var skipPanel = optionPanelsGroup.add("panel", undefined, getLabel("panel.skip"));
    setupPanel(skipPanel, ROW_PANEL_SPACING);
    skipPanel.alignChildren = ["left", "top"];

    var skipRowControls = addSkipRow(skipPanel, "checkbox.skipRows", "tooltip.skipRows", onChange);
    ui.skipRowCheckbox = skipRowControls.checkbox;
    ui.skipRowInput = skipRowControls.input;

    var skipColumnControls = addSkipRow(skipPanel, "checkbox.skipColumns", "tooltip.skipColumns", onChange);
    ui.skipColumnCheckbox = skipColumnControls.checkbox;
    ui.skipColumnInput = skipColumnControls.input;
}

/**
 * スキップ数の1行（チェックボックス＋入力欄）を追加する
 * @param {Panel} parent 追加先のパネル
 * @param {string} labelKey チェックボックスのラベルキー
 * @param {string} tooltipKey ツールチップのキー
 * @param {function} onChange 値が変わったときに呼ぶ処理
 * @returns {{checkbox: Checkbox, input: EditText}} 作成したコントロール
 */
function addSkipRow(parent, labelKey, tooltipKey, onChange) {
    var skipGroup = parent.add("group");
    setupRow(skipGroup, "left", CONTROL_SPACING);

    var skipCheckbox = skipGroup.add("checkbox", undefined, labelText(labelKey));
    skipCheckbox.preferredSize.width = SKIP_CHECKBOX_WIDTH;
    var skipInput = skipGroup.add("edittext", undefined, String(SKIP_COUNT_DEFAULT));
    skipInput.preferredSize.width = SKIP_INPUT_WIDTH;

    skipCheckbox.helpTip = getLabel(tooltipKey);
    skipInput.helpTip = getLabel(tooltipKey);
    skipInput.enabled = skipCheckbox.value;

    skipCheckbox.onClick = function () {
        skipInput.enabled = skipCheckbox.value;
        onChange();
    };
    skipInput.onChange = function () {
        this.text = String(normalizeSkipCount(parseInt(this.text, 10)));
        onChange();
    };
    changeValueByArrowKey(skipInput, SKIP_ARROW_STEP, normalizeSkipCount, onChange);

    return { checkbox: skipCheckbox, input: skipInput };
}

/**
 * ダイアログ下部のボタン列を追加する（左：画面モード／右：キャンセル・OK）
 * @param {Window} dialog 対象のダイアログ
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function addButtonRow(dialog, ui) {
    /* メイングループ（横並び）/ Main group (horizontal layout) */
    var btnRowGroup = dialog.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
    btnRowGroup.alignment = ["fill", "bottom"];

    /* 左側グループ / Left-side button group */
    var btnLeftGroup = btnRowGroup.add("group");
    btnLeftGroup.alignChildren = ["left", "center"];
    ui.btnScreenMode = btnLeftGroup.add("button", undefined, getScreenModeButtonLabel());
    ui.btnScreenMode.helpTip = getLabel("tooltip.screenMode");
    ui.btnScreenMode.onClick = function () {
        toggleScreenMode();
        ui.btnScreenMode.text = getScreenModeButtonLabel();
    };

    /* スペーサー（伸縮）/ Spacer (stretchable) */
    var spacer = btnRowGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = BUTTON_ROW_SPACER_MIN_WIDTH;

    /* 右側グループ / Right-side button group */
    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.alignChildren = ["right", "center"];
    btnRightGroup.alignment = ["right", "center"];
    btnRightGroup.margins = 0;
    btnRightGroup.spacing = CONTROL_SPACING;
    var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
}

/**
 * ダイアログを組み立てる
 * @param {Array<SwatchChoice>} swatchChoices スウォッチ一覧
 * @param {{odd: object, even: object}} defaultFills 奇数行・偶数行の初期値
 * @param {function} onChange 値が変わったときに呼ぶ処理
 * @returns {object} ダイアログとコントロールをまとめた UI オブジェクト
 */
function buildDialog(swatchChoices, defaultFills, onChange) {
    var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    dialog.opacity = DIALOG_OPACITY;
    setupWindow(dialog);

    /* 左に奇数行、右に偶数行の2カラム / Two columns: odd rows on the left, even rows on the right */
    var rowPanelsGroup = dialog.add("group");
    setupRow(rowPanelsGroup, "fill", COLUMN_SPACING);
    rowPanelsGroup.alignChildren = ["left", "top"];

    var ui = { dialog: dialog };
    ui.oddRowFill = createRowFillPanel(rowPanelsGroup, "panel.oddRows", "tooltip.oddRows",
        swatchChoices, defaultFills.odd, onChange);
    ui.evenRowFill = createRowFillPanel(rowPanelsGroup, "panel.evenRows", "tooltip.evenRows",
        swatchChoices, defaultFills.even, onChange);

    addOptions(dialog, ui, onChange);
    addButtonRow(dialog, ui);

    refreshRowFill(ui.oddRowFill);
    refreshRowFill(ui.evenRowFill);
    return ui;
}

// =========================================
// UIの値の読み取り / Read UI values
// =========================================

/**
 * @typedef {object} ZebraSettings
 * @property {CellFill} oddFill 奇数行の塗り
 * @property {CellFill} evenFill 偶数行の塗り
 * @property {boolean} swapRows 奇数行と偶数行を入れ替えるか
 * @property {number} skipRowCount 上からスキップする行数
 * @property {number} skipColumnCount 左からスキップする列数
 */

/**
 * パネルの設定を、セルに適用する形の塗りとして読み取る
 * @param {RowFillControls} rowFill 対象のパネル
 * @returns {CellFill} 読み取った塗り
 */
function readRowFill(rowFill) {
    var swatchChoice = getRowFillChoice(rowFill);
    return {
        color: findSwatch(swatchChoice.swatchName),
        tint: isTintDisabledKind(swatchChoice.swatchKind) ? null : normalizeTint(parseFloat(rowFill.tintInput.text))
    };
}

/**
 * スキップ数を読み取る（チェックが外れていれば 0）
 * @param {Checkbox} skipCheckbox スキップのチェックボックス
 * @param {EditText} skipInput スキップ数の入力欄
 * @returns {number} スキップする数
 */
function readSkipCount(skipCheckbox, skipInput) {
    return skipCheckbox.value ? normalizeSkipCount(parseInt(skipInput.text, 10)) : 0;
}

/**
 * ダイアログの設定をまとめて読み取る
 * @param {object} ui UI オブジェクト
 * @returns {ZebraSettings} 読み取った設定
 */
function readZebraSettings(ui) {
    return {
        oddFill: readRowFill(ui.oddRowFill),
        evenFill: readRowFill(ui.evenRowFill),
        swapRows: ui.swapCheckbox.value,
        skipRowCount: readSkipCount(ui.skipRowCheckbox, ui.skipRowInput),
        skipColumnCount: readSkipCount(ui.skipColumnCheckbox, ui.skipColumnInput)
    };
}

// =========================================
// 塗りの適用 / Apply fills
// =========================================

/**
 * @typedef {object} ZebraContext
 * @property {Array<Cell>} targetCells 対象のセル
 * @property {Array<CellFill>} originalFills 実行前の塗り（セルと同じ並び）
 * @property {boolean} previewApplied プレビューを当てたままかどうか
 */

/**
 * 選択セルが使っている行または列の位置を、重複なく昇順で集める
 * @param {Array<Cell>} cells 対象のセル
 * @param {string} parentName "parentRow" または "parentColumn"
 * @returns {Array<number>} 昇順に並べた位置
 */
function collectSortedIndices(cells, parentName) {
    var sortedIndices = [];
    var seen = {};
    for (var i = 0; i < cells.length; i++) {
        var parentIndex = cells[i][parentName].index;
        if (seen[parentIndex]) continue;
        seen[parentIndex] = true;
        sortedIndices.push(parentIndex);
    }
    sortedIndices.sort(function (a, b) { return a - b; });
    return sortedIndices;
}

/**
 * 先頭から指定数ぶんを「スキップする位置」として控える
 * @param {Array<number>} sortedIndices 昇順に並べた位置
 * @param {number} skipCount スキップする数
 * @returns {object} スキップする位置をキーに持つオブジェクト
 */
function buildSkipLookup(sortedIndices, skipCount) {
    var skipped = {};
    for (var i = 0; i < skipCount && i < sortedIndices.length; i++) {
        skipped[sortedIndices[i]] = true;
    }
    return skipped;
}

/**
 * スキップしていない行に、選択範囲の上から 0, 1, 2, … の順番を振る
 * @param {Array<number>} rowIndices 昇順に並べた行の位置
 * @param {object} skippedRows スキップする行
 * @returns {object} 行の位置をキー、順番を値に持つオブジェクト
 */
function buildRowOrderMap(rowIndices, skippedRows) {
    var rowOrderByIndex = {};
    var rowOrder = 0;
    for (var i = 0; i < rowIndices.length; i++) {
        if (skippedRows[rowIndices[i]]) continue;
        rowOrderByIndex[rowIndices[i]] = rowOrder++;
    }
    return rowOrderByIndex;
}

/**
 * セルに塗りを設定する
 * @param {Cell} cell 対象のセル
 * @param {CellFill} fill 設定する塗り。tint が null のときは濃淡を触らない
 * @returns {void}
 */
function applyCellFill(cell, fill) {
    if (!fill || !fill.color) return;
    cell.fillColor = fill.color;
    if (fill.tint === null) return;

    /* グラデーションなど濃淡を持たない塗りでは無視する / Fills without a tint, such as gradients, are left alone */
    try {
        cell.fillTint = fill.tint;
    } catch (e) { }
}

/**
 * 選択範囲内の行の並びを基準に、交互の塗りを適用する
 * @param {ZebraContext} context 対象のセルと実行前の塗り
 * @param {ZebraSettings} settings ダイアログの設定
 * @returns {void}
 */
function applyAlternatingFills(context, settings) {
    var targetCells = context.targetCells;
    if (targetCells.length === 0) return;

    var oddFill = settings.swapRows ? settings.evenFill : settings.oddFill;
    var evenFill = settings.swapRows ? settings.oddFill : settings.evenFill;

    var rowIndices = collectSortedIndices(targetCells, "parentRow");
    var columnIndices = collectSortedIndices(targetCells, "parentColumn");
    var skippedRows = buildSkipLookup(rowIndices, settings.skipRowCount);
    var skippedColumns = buildSkipLookup(columnIndices, settings.skipColumnCount);
    var rowOrderByIndex = buildRowOrderMap(rowIndices, skippedRows);

    for (var i = 0; i < targetCells.length; i++) {
        var cell = targetCells[i];
        var rowIndex = cell.parentRow.index;
        var columnIndex = cell.parentColumn.index;

        /* スキップした行・列は実行前の塗りに戻す / Skipped rows and columns go back to their previous fill */
        if (skippedRows[rowIndex] || skippedColumns[columnIndex]) {
            applyCellFill(cell, context.originalFills[i]);
            continue;
        }

        var rowOrder = rowOrderByIndex[rowIndex];
        if (rowOrder === undefined) continue;
        applyCellFill(cell, (rowOrder % 2 === 0) ? oddFill : evenFill);
    }
}

// =========================================
// プレビューと確定 / Preview & apply
// =========================================

/**
 * プレビューとして適用した塗りを取り消す
 * @param {ZebraContext} context 対象のセルと実行前の塗り
 * @returns {void}
 */
function clearPreview(context) {
    if (!context.previewApplied) return;
    context.previewApplied = false;

    /* プレビュー1回分の取り消し。失敗しても続行する / Undo the single preview step; keep going if it fails */
    try {
        app.undo();
        app.activeDocument.recompose();
    } catch (e) { }
}

/**
 * 現在の設定で塗りのプレビューを描き直す
 * @param {ZebraContext} context 対象のセルと実行前の塗り
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function refreshPreview(context, ui) {
    clearPreview(context);
    var settings = readZebraSettings(ui);

    /* 取り消しを1ステップにまとめる / Keep the preview to a single undo step */
    try {
        app.doScript(function () {
            applyAlternatingFills(context, settings);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.previewFills"));
        context.previewApplied = true;
        app.activeDocument.recompose();
    } catch (e) {
        context.previewApplied = false;
    }
}

/**
 * 選択セルに交互の塗りを適用するダイアログを表示して実行する
 * @returns {void}
 */
function main() {
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var targetCells = getSelectedCells();
    if (targetCells.length === 0) {
        alert(getLabel("alert.noCellSelection"));
        return;
    }

    var swatchChoices = getSwatchChoices();
    if (swatchChoices.length === 0) {
        alert(getLabel("alert.noUsableSwatch"));
        return;
    }

    var context = {
        targetCells: targetCells,
        originalFills: captureCellFills(targetCells),
        previewApplied: false
    };
    var defaultFills = detectDefaultFills(targetCells, swatchChoices[0].swatchName);

    /* セルの青い選択ハイライトがプレビューの邪魔になるため、選択は一旦解除する
       / The blue selection highlight hides the preview, so clear the selection */
    app.selection = null;

    var ui = null;
    ui = buildDialog(swatchChoices, defaultFills, function () {
        refreshPreview(context, ui);
    });

    refreshPreview(context, ui);

    if (ui.dialog.show() === 1) {
        clearPreview(context);
        var settings = readZebraSettings(ui);
        app.doScript(function () {
            applyAlternatingFills(context, settings);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.applyFills"));
    } else {
        clearPreview(context);
    }
}

main();

})();
