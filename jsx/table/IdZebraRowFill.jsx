#target indesign

/*

### 概要

選択した表セルに対して、選択範囲内の行の並びを基準に交互の塗り（縞模様）をプレビュー付きで適用します。

詳細は README を参照してください。

### Overview

Applies alternating fills (zebra striping) to the selected table cells, based on the row order within the selection, with a live preview.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdZebraRowFill";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdZebraRowFill.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdZebraRowFill.md

// Original idea
// KK sawa

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

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

// =========================================
// レイアウト設定 / Layout settings
// =========================================

/* カラードロップダウン・濃淡入力欄・スライダー・スキップ入力欄の幅（px）
   / Widths of the colour dropdown, tint field, slider and skip fields (px) */
var COLOR_DROPDOWN_WIDTH = 140;
var TINT_INPUT_WIDTH     = 50;
var TINT_SLIDER_WIDTH    = 150;
var SKIP_INPUT_WIDTH     = 30;

/* 色見本の一辺（px）/ Size of the swatch preview box (px) */
var SWATCH_PREVIEW_SIZE = [18, 18];

// ==============================
// UIレイアウトの共通設定 / Shared UI layout
// ==============================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 20;                 /* 2カラムの間隔 / gap between columns */

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

(function () {

    // =========================================
    // ラベル定義 / Labels
    // =========================================

    /**
     * UI 言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        var isJapanese = false;
        try {
            if (app.locale && app.locale === Locale.JAPANESE) isJapanese = true;
        } catch (e) {}
        try {
            if (!isJapanese && $.locale && $.locale.toString().indexOf("ja") === 0) isJapanese = true;
        } catch (e) {}
        return isJapanese ? "ja" : "en";
    }

    var currentLang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "セルの塗りを交互に反復", en: "Repeat Cell Fills Every Other Row" }
        },
        panel: {
            oddRows:  { ja: "奇数行", en: "Odd Rows" },
            evenRows: { ja: "偶数行", en: "Even Rows" }
        },
        field: {
            color:       { ja: "カラー：", en: "Color:" },
            tint:        { ja: "濃淡：", en: "Tint:" },
            skipRows:    { ja: "行のスキップ", en: "Skip top rows" },
            skipColumns: { ja: "列のスキップ", en: "Skip left columns" }
        },
        unit: {
            row:    { ja: "行", en: "rows" },
            column: { ja: "列", en: "columns" }
        },
        swatch: {
            black: { ja: "黒", en: "Black" },
            paper: { ja: "紙色", en: "Paper Color" },
            none:  { ja: "なし", en: "None" }
        },
        button: {
            ok:           { ja: "OK", en: "OK" },
            cancel:       { ja: "キャンセル", en: "Cancel" },
            swap:         { ja: "交換", en: "Swap" },
            preview:      { ja: "プレビュー", en: "Preview" },
            standardMode: { ja: "標準モード", en: "Standard" }
        },
        alert: {
            openDocument:   { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectCells:    { ja: "セルを選択してください。", en: "Please select table cells." },
            noUsableColors: { ja: "使用可能なカラーがありません。", en: "No usable colors are available." }
        },
        undo: {
            preview: { ja: "塗りプレビュー", en: "Fill Preview" },
            apply:   { ja: "塗りの設定", en: "Apply Fill" }
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
        return node[currentLang] || node.en || labelKey;
    }

    /**
     * 選択セルに交互の塗りを適用するダイアログを表示して実行する
     * @returns {void}
     */
    function main() {
        // -------------------------------------------------------------
        // 前提チェック
        // -------------------------------------------------------------
        if (app.documents.length === 0) {
            alert(getLabel("alert.openDocument"));
            return;
        }

        var doc = app.activeDocument;

        // 選択セルを保持し、プレビュー中は選択ハイライトを一旦消して見やすくする
        var baseTargetCells = [];
        var selCells;
        try {
            selCells = app.selection[0].cells;
        } catch (e) {
            alert(getLabel("alert.selectCells"));
            return;
        }
        if (!selCells || selCells.length === 0) {
            alert(getLabel("alert.selectCells"));
            return;
        }
        for (var bi = 0; bi < selCells.length; bi++) {
            baseTargetCells.push(selCells[bi]);
        }
        // 元の塗り状態を保持（スキップ時に復元するため）
        var originalCellStyles = [];
        for (var oi = 0; oi < baseTargetCells.length; oi++) {
            var oc = baseTargetCells[oi];
            originalCellStyles.push({
                color: oc.fillColor,
                tint: oc.fillTint
            });
        }
        // プレビュー時にセルの青いハイライトが邪魔になるため、選択は一旦解除
        app.selection = null;

        // -------------------------------------------------------------
        // スウォッチ取得（Registration を除外）
        // -------------------------------------------------------------
        var REG_NAMES = ["Registration", "[Registration]", "トンボ用"];
        var PAPER_NAMES = ["Paper", "[Paper]", "紙色", "[紙色]"];
        var NONE_NAMES = ["None", "[None]", "なし", "[なし]"];

        /**
         * 名前がリストに含まれるかを判定する
         * @param {string} name 判定する名前
         * @param {Array<string>} list 照合するリスト
         * @returns {boolean} 含まれていれば true
         */
        function isInList(name, list) {
            for (var i = 0; i < list.length; i++) {
                if (name === list[i]) return true;
            }
            return false;
        }

        /**
         * レジストレーションのスウォッチ名かどうかを判定する
         * @param {string} name スウォッチ名
         * @returns {boolean} レジストレーションなら true
         */
        function isRegistrationName(name) {
            if (!name) return false;
            if (isInList(name, REG_NAMES)) return true;
            return name.toLowerCase() === "registration";
        }

        /**
         * 紙色のスウォッチ名かどうかを判定する
         * @param {string} name スウォッチ名
         * @returns {boolean} 紙色なら true
         */
        function isPaperName(name) {
            return isInList(name, PAPER_NAMES);
        }

        /**
         * 「なし」のスウォッチ名かどうかを判定する
         * @param {string} name スウォッチ名
         * @returns {boolean} 「なし」なら true
         */
        function isNoneName(name) {
            return isInList(name, NONE_NAMES);
        }

        var swatches = [];
        for (var i = 0; i < doc.swatches.length; i++) {
            var sw = doc.swatches[i];
            var name = sw.name;

            if (isRegistrationName(name)) continue;

            var label = name;
            if (isPaperName(name)) {
                label = getLabel("swatch.paper");
            } else if (isNoneName(name)) {
                label = getLabel("swatch.none");
            }
            if (name === "Black") label = getLabel("swatch.black");

            swatches.push({
                name: name,
                label: label
            });
        }

        if (swatches.length === 0) {
            alert(getLabel("alert.noUsableColors"));
            return;
        }

        // -------------------------------------------------------------
        // プレビュー管理ステート
        // -------------------------------------------------------------
        var state = { previewed: false };

        // -------------------------------------------------------------
        // スウォッチプレビュー用ヘルパー
        // -------------------------------------------------------------
        /**
         * 比較しやすいようにスウォッチ名を正規化する
         * @param {string} n スウォッチ名
         * @returns {string} 正規化した名前
         */
        function normalizeSwatchName(n) {
            if (n == null) return "";
            n = String(n);
            n = n.replace(/^\[|\]$/g, "");
            n = n.replace(/^\s+|\s+$/g, "").toLowerCase();
            return n;
        }

        /**
         * 色見本で黒として扱うスウォッチ名かどうかを判定する
         * @param {string} n スウォッチ名
         * @returns {boolean} 黒として扱うなら true
         */
        function isBlackPreviewName(n) {
            var x = normalizeSwatchName(n);
            return x === "black" || x === "ブラック" || x === "黒" || x === "registration";
        }

        /**
         * 色見本で白として扱うスウォッチ名かどうかを判定する
         * @param {string} n スウォッチ名
         * @returns {boolean} 白として扱うなら true
         */
        function isWhitePreviewName(n) {
            var x = normalizeSwatchName(n);
            return x === "paper" || x === "紙色" || x === "none" || x === "なし";
        }

        /**
         * スウォッチのカラー値を表示用の RGB に変換する
         * @param {Swatch} swatch 対象のスウォッチ
         * @returns {Array<number>|null} RGB 値の配列。変換できない場合は null
         */
        function convertSwatchToPreviewRGB(swatch) {
            if (!swatch) return [0.5, 0.5, 0.5];
            var sname = swatch.name != null ? String(swatch.name) : "";
            if (isWhitePreviewName(sname)) return [1, 1, 1];
            if (isBlackPreviewName(sname)) return [0, 0, 0];
            try {
                if (swatch.hasOwnProperty("colorValue")) {
                    var vals = swatch.colorValue;
                    if (swatch.space === ColorSpace.RGB) {
                        return [vals[0] / 255, vals[1] / 255, vals[2] / 255];
                    }
                    if (swatch.space === ColorSpace.CMYK) {
                        var c = vals[0] / 100,
                            m = vals[1] / 100,
                            y = vals[2] / 100,
                            k = vals[3] / 100;
                        return [(1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k)];
                    }
                }
            } catch (e) { }
            return [0.5, 0.5, 0.5];
        }

        /**
         * スウォッチの色見本を表示する枠を作る
         * @param {object} parent 追加先のコンテナ
         * @returns {Group} 色見本用のグループ
         */
        function createSwatchPreviewBox(parent) {
            var box = parent.add("group");
            box.preferredSize = SWATCH_PREVIEW_SIZE;
            box.minimumSize = [18, 18];
            box.maximumSize = [18, 18];
            return box;
        }

        /**
         * 選択中のスウォッチに合わせて色見本を塗り直す
         * @param {Group} previewBox 色見本のグループ
         * @param {string} swatchName スウォッチ名
         * @returns {void}
         */
        function updateSwatchPreview(previewBox, swatchName) {
            if (!previewBox) return;
            var swatch = null;
            try {
                swatch = doc.swatches.itemByName(String(swatchName));
                if (!swatch || !swatch.isValid) return;
            } catch (e) {
                return;
            }
            var rgb = convertSwatchToPreviewRGB(swatch);
            try {
                previewBox.graphics.backgroundColor = previewBox.graphics.newBrush(
                    previewBox.graphics.BrushType.SOLID_COLOR,
                    [rgb[0], rgb[1], rgb[2], 1]
                );
                if (dialog) dialog.update();
            } catch (e2) { }
        }

        // -------------------------------------------------------------
        // プレビュー表示モード切替ヘルパー
        // -------------------------------------------------------------
        /**
         * 現在プレビュー表示になっているかを判定する
         * @returns {boolean} プレビュー表示なら true
         */
        function isPreviewScreenMode() {
            try {
                return app.activeWindow && app.activeWindow.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE;
            } catch (e) {
                return false;
            }
        }

        /**
         * 標準表示とプレビュー表示を切り替える
         * @returns {void}
         */
        function togglePreviewScreenMode() {
            try {
                var w = app.activeWindow;
                if (!w) return;
                if (w.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE) {
                    w.screenMode = ScreenModeOptions.PREVIEW_OFF;
                } else {
                    w.screenMode = ScreenModeOptions.PREVIEW_TO_PAGE;
                }
            } catch (e) { }
        }

        /**
         * 画面モードに応じたトグルボタンのラベルを返す
         * @returns {string} ボタンに表示する文字列
         */
        function getPreviewToggleButtonLabel() {
            return isPreviewScreenMode()
                ? getLabel("button.preview")
                : getLabel("button.standardMode");
        }

        /**
         * トグルボタンのラベルを現在の画面モードに合わせて更新する
         * @param {Button} button 対象のボタン
         * @returns {void}
         */
        function updatePreviewToggleButtonLabel(button) {
            if (!button) return;
            button.text = getPreviewToggleButtonLabel();
        }

        // -------------------------------------------------------------
        // デフォルトのカラー／濃淡（選択セル内で使われている設定をもとに決定）
        // -------------------------------------------------------------
        var defaultOddColorName = swatches[0].name;
        var defaultEvenColorName = swatches[0].name;
        var defaultOddTint = 100;
        var defaultEvenTint = 100;

        try {
            var firstCell = baseTargetCells[0];
            var fillColor = firstCell.fillColor;
            var fillTint = firstCell.fillTint;

            if (fillColor && fillColor.name) {
                defaultOddColorName = fillColor.name;
                defaultEvenColorName = fillColor.name;
            }
            if (!isNaN(fillTint)) {
                defaultOddTint = fillTint;
                defaultEvenTint = fillTint;
            }

            var comboCounts = {};

            for (var ci = 0; ci < baseTargetCells.length; ci++) {
                var c = baseTargetCells[ci];
                var cFillColor = c.fillColor;
                var cFillTint = c.fillTint;

                var cColorName = (cFillColor && cFillColor.name) ? cFillColor.name : defaultOddColorName;
                var cTint = (!isNaN(cFillTint)) ? cFillTint : defaultOddTint;

                var key = cColorName + "||" + cTint;
                if (!comboCounts[key]) {
                    comboCounts[key] = {
                        count: 0,
                        colorName: cColorName,
                        tint: cTint
                    };
                }
                comboCounts[key].count++;
            }

            var combos = [];
            for (var kCombo in comboCounts) {
                combos.push(comboCounts[kCombo]);
            }
            combos.sort(function (a, b) {
                return b.count - a.count;
            });

            if (combos.length >= 1) {
                defaultOddColorName = combos[0].colorName;
                defaultOddTint = combos[0].tint;
            }
            if (combos.length >= 2) {
                defaultEvenColorName = combos[1].colorName;
                defaultEvenTint = combos[1].tint;
            } else {
                defaultEvenColorName = defaultOddColorName;
                defaultEvenTint = defaultOddTint;
            }
        } catch (e) { }

        // -------------------------------------------------------------
        // ダイアログボックス作成
        // -------------------------------------------------------------
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        var dialogOpacity = 0.97;


        dialog.opacity = dialogOpacity;
        setupWindow(dialog, 10);

        /* メインパネル（左右2カラム）/ Main area with two columns */
        var mainGroup = dialog.add("group");
        setupRow(mainGroup, "fill", COLUMN_SPACING);
        mainGroup.alignChildren = "top";

        // -------------------------------------------------------------
        // 左カラム：奇数行パネル
        // -------------------------------------------------------------
        var oddPanel = mainGroup.add("panel", undefined, getLabel("panel.oddRows"));

        setupPanel(oddPanel, 6);

        oddPanel.alignChildren = "left";

        var oddColorPanel = oddPanel.add("group");
        oddColorPanel.orientation = "column";
        oddColorPanel.alignChildren = "left";

        var oddColorLabelRow = oddColorPanel.add("group");
        oddColorLabelRow.orientation = "row";
        oddColorLabelRow.alignChildren = ["left", "center"];
        oddColorLabelRow.add("statictext", undefined, getLabel("field.color"));
        var oddPreviewBox = createSwatchPreviewBox(oddColorLabelRow);

        var oddSwatchLabels = [];
        for (var j = 0; j < swatches.length; j++) {
            oddSwatchLabels.push(swatches[j].label);
        }
        var oddColorDropdown = oddColorPanel.add("dropdownlist", undefined, oddSwatchLabels);
        oddColorDropdown.preferredSize.width = COLOR_DROPDOWN_WIDTH;

        var oddDefaultIdx = 0;
        for (var j2 = 0; j2 < swatches.length; j2++) {
            if (swatches[j2].name === defaultOddColorName) {
                oddDefaultIdx = j2;
                break;
            }
        }
        oddColorDropdown.selection = oddDefaultIdx;

        oddColorDropdown.onChange = function () {
            updateTintUI(true);
            updateSwatchPreview(oddPreviewBox, getOddSelectedColorName());
            updatePreview();
        };

        // 濃淡設定（奇数）
        var oddTintPanel = oddPanel.add("group");
        oddTintPanel.orientation = "column";
        oddTintPanel.alignChildren = "left";
        oddTintPanel.margins = [0, 20, 0, 10];

        var oddTintLabelGroup = oddTintPanel.add("group");
        oddTintLabelGroup.orientation = "row";
        oddTintLabelGroup.add("statictext", undefined, getLabel("field.tint"));
        var oddTintText = oddTintLabelGroup.add("edittext", undefined, String(defaultOddTint));
        oddTintText.preferredSize.width = TINT_INPUT_WIDTH;

        var oddSlider = oddTintPanel.add("slider", undefined, defaultOddTint, 0, 100);
        oddSlider.preferredSize.width = TINT_SLIDER_WIDTH;

        // -------------------------------------------------------------
        // 右カラム：偶数行パネル
        // -------------------------------------------------------------
        var evenPanel = mainGroup.add("panel", undefined, getLabel("panel.evenRows"));

        setupPanel(evenPanel, 6);

        evenPanel.alignChildren = "left";

        var evenColorPanel = evenPanel.add("group");
        evenColorPanel.orientation = "column";
        evenColorPanel.alignChildren = "left";

        var evenColorLabelRow = evenColorPanel.add("group");
        evenColorLabelRow.orientation = "row";
        evenColorLabelRow.alignChildren = ["left", "center"];
        evenColorLabelRow.add("statictext", undefined, getLabel("field.color"));
        var evenPreviewBox = createSwatchPreviewBox(evenColorLabelRow);

        var evenSwatchLabels = [];
        for (var k = 0; k < swatches.length; k++) {
            evenSwatchLabels.push(swatches[k].label);
        }
        var evenColorDropdown = evenColorPanel.add("dropdownlist", undefined, evenSwatchLabels);
        evenColorDropdown.preferredSize.width = COLOR_DROPDOWN_WIDTH;

        var evenDefaultIdx = 0;
        for (var k2 = 0; k2 < swatches.length; k2++) {
            if (swatches[k2].name === defaultEvenColorName) {
                evenDefaultIdx = k2;
                break;
            }
        }
        evenColorDropdown.selection = evenDefaultIdx;

        evenColorDropdown.onChange = function () {
            updateTintUI(false);
            updateSwatchPreview(evenPreviewBox, getEvenSelectedColorName());
            updatePreview();
        };

        // 濃淡設定（偶数）
        var evenTintPanel = evenPanel.add("group");
        evenTintPanel.orientation = "column";
        evenTintPanel.alignChildren = "left";
        evenTintPanel.margins = [0, 20, 0, 10];

        var evenTintLabelGroup = evenTintPanel.add("group");
        evenTintLabelGroup.orientation = "row";
        evenTintLabelGroup.add("statictext", undefined, getLabel("field.tint"));
        var evenTintText = evenTintLabelGroup.add("edittext", undefined, String(defaultEvenTint));
        evenTintText.preferredSize.width = TINT_INPUT_WIDTH;

        var evenSlider = evenTintPanel.add("slider", undefined, defaultEvenTint, 0, 100);
        evenSlider.preferredSize.width = TINT_SLIDER_WIDTH;

        // -------------------------------------------------------------
        // オプションエリア（交換／スキップ設定）
        // -------------------------------------------------------------
        var optionGroup = dialog.add("group");
        optionGroup.orientation = "column";
        optionGroup.alignChildren = "left";
        optionGroup.alignment = ["left", "top"];

        var swapCheckbox = optionGroup.add("checkbox", undefined, getLabel("button.swap"));
        swapCheckbox.value = false;
        swapCheckbox.onClick = function () {
            updatePreview();
        };

        // 行のスキップ：選択範囲内で、上から指定行のみカラーリングしない
        var rowSkipGroup = optionGroup.add("group");
        rowSkipGroup.orientation = "row";
        var rowSkipCheckbox = rowSkipGroup.add("checkbox", undefined, getLabel("field.skipRows"));
        var rowSkipText = rowSkipGroup.add("edittext", undefined, "1");
        rowSkipText.preferredSize.width = SKIP_INPUT_WIDTH;
        rowSkipGroup.add("statictext", undefined, getLabel("unit.row"));

        // 列のスキップ：選択範囲内で、左から指定列のみカラーリングしない
        var colSkipGroup = optionGroup.add("group");
        colSkipGroup.orientation = "row";
        var colSkipCheckbox = colSkipGroup.add("checkbox", undefined, getLabel("field.skipColumns"));
        var colSkipText = colSkipGroup.add("edittext", undefined, "1");
        colSkipText.preferredSize.width = SKIP_INPUT_WIDTH;
        colSkipGroup.add("statictext", undefined, getLabel("unit.column"));

        /**
         * スキップ数を 0 以上の整数に整える
         * @param {number} v 入力された値
         * @returns {number} 整えた値
         */
        function adjustSkipValue(v) {
            if (isNaN(v)) v = 0;
            if (v < 0) v = 0;
            return Math.round(v);
        }

        /**
         * キー操作で増減した値を表示用の文字列にする
         * @param {number} value 数値
         * @returns {string} 表示する文字列
         */
        function formatArrowValue(value) {
            return String(Math.round(value));
        }

        /**
         * 入力欄に上下キーでの増減操作を追加する
         * @param {EditText} editText 対象の入力欄
         * @param {boolean} allowNegative 負の値を許可するか
         * @param {function} onAfterChange 値の変更後に呼ぶ処理
         * @param {object} options 最小値などの追加設定
         * @returns {void}
         */
        function bindArrowKeyValueControl(editText, allowNegative, onAfterChange, options) {
            options = options || {};
            var minValue = (options.min != null) ? options.min : (allowNegative ? null : 0);
            var maxValue = (options.max != null) ? options.max : null;

            editText.addEventListener("keydown", function (event) {
                var key = event.keyName;
                if (key !== "Up" && key !== "Down") return;

                var value = Number(editText.text);
                if (isNaN(value)) return;

                var keyboard = ScriptUI.environment.keyboardState;
                var delta = keyboard.shiftKey ? 10 : 1;

                if (keyboard.shiftKey) {
                    if (key === "Up") {
                        value = Math.ceil((value + 1) / delta) * delta;
                    } else {
                        value = Math.floor((value - 1) / delta) * delta;
                    }
                } else {
                    if (key === "Up") {
                        value += delta;
                    } else {
                        value -= delta;
                    }
                }

                value = Math.round(value);

                if (!allowNegative && value < 0) value = 0;
                if (minValue != null && value < minValue) value = minValue;
                if (maxValue != null && value > maxValue) value = maxValue;

                editText.text = formatArrowValue(value);

                if (typeof onAfterChange === "function") {
                    onAfterChange(value, editText);
                }

                try { event.preventDefault(); } catch (e) { }
            });
        }

        /**
         * スキップ設定に応じて関連コントロールの状態を更新する
         * @returns {void}
         */
        function updateSkipUI() {
            rowSkipText.enabled = rowSkipCheckbox.value;
            colSkipText.enabled = colSkipCheckbox.value;
        }
        updateSkipUI();

        rowSkipCheckbox.onClick = function () {
            updateSkipUI();
            updatePreview();
        };
        colSkipCheckbox.onClick = function () {
            updateSkipUI();
            updatePreview();
        };
        rowSkipText.onChange = function () {
            var v = adjustSkipValue(parseInt(this.text, 10));
            this.text = String(v);
            updatePreview();
        };
        colSkipText.onChange = function () {
            var v = adjustSkipValue(parseInt(this.text, 10));
            this.text = String(v);
            updatePreview();
        };
        bindArrowKeyValueControl(rowSkipText, false, function () {
            var v = adjustSkipValue(parseInt(rowSkipText.text, 10));
            rowSkipText.text = String(v);
            updatePreview();
        }, { min: 0 });

        bindArrowKeyValueControl(colSkipText, false, function () {
            var v = adjustSkipValue(parseInt(colSkipText.text, 10));
            colSkipText.text = String(v);
            updatePreview();
        }, { min: 0 });

        // -------------------------------------------------------------
        // ボタングループ（左：プレビューモード切替／右：キャンセル・OK）
        // -------------------------------------------------------------
        var bottomGroup = dialog.add("group");
        setupRow(bottomGroup, "fill", 8);
        bottomGroup.alignChildren = "fill";
        bottomGroup.margins = [0, 8, 0, 0];

        var btnPreviewToggle = bottomGroup.add("button", undefined, getPreviewToggleButtonLabel());
        btnPreviewToggle.alignment = ["left", "center"];
        btnPreviewToggle.onClick = function () {
            togglePreviewScreenMode();
            updatePreviewToggleButtonLabel(btnPreviewToggle);
        };

        var spacer = bottomGroup.add("group");
        spacer.alignment = ["fill", "top"];

        /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
        var buttonGroup = bottomGroup.add("group");
        setupRow(buttonGroup, "right", 8);
        buttonGroup.alignChildren = "right";

        var cancelButton = buttonGroup.add("button", undefined, getLabel("button.cancel"), {
            name: "cancel"
        });
        var okButton = buttonGroup.add("button", undefined, getLabel("button.ok"), {
            name: "ok"
        });

        // -------------------------------------------------------------
        // 濃淡ディム制御用ヘルパー
        // -------------------------------------------------------------
        /**
         * 奇数行で選択中のカラー名を取得する
         * @returns {string} スウォッチ名
         */
        function getOddSelectedColorName() {
            if (oddColorDropdown.selection) {
                return swatches[oddColorDropdown.selection.index].name;
            }
            return swatches[0].name;
        }

        /**
         * 偶数行で選択中のカラー名を取得する
         * @returns {string} スウォッチ名
         */
        function getEvenSelectedColorName() {
            if (evenColorDropdown.selection) {
                return swatches[evenColorDropdown.selection.index].name;
            }
            return swatches[0].name;
        }

        /**
         * 選択中のスウォッチに応じて濃淡コントロールの有効／無効を切り替える
         * @param {boolean} isOdd 奇数行側なら true
         * @returns {void}
         */
        function updateTintUI(isOdd) {
            var name = isOdd ? getOddSelectedColorName() : getEvenSelectedColorName();
            var disabled = isNoneName(name) || isPaperName(name);

            if (isOdd) {
                oddSlider.enabled = !disabled;
                oddTintText.enabled = !disabled;
            } else {
                evenSlider.enabled = !disabled;
                evenTintText.enabled = !disabled;
            }
        }

        updateTintUI(true);
        updateTintUI(false);
        updateSwatchPreview(oddPreviewBox, getOddSelectedColorName());
        updateSwatchPreview(evenPreviewBox, getEvenSelectedColorName());

        // -------------------------------------------------------------
        // プレビュー制御（doScript + app.undo）
        // -------------------------------------------------------------
        /**
         * プレビューとして適用した塗りを取り消す
         * @returns {void}
         */
        function clearPreview() {
            if (!state.previewed) return;
            try { app.undo(); } catch (e) { }
            state.previewed = false;
            try { doc.recompose(); } catch (e2) { }
        }

        /**
         * 現在の設定で塗りのプレビューを描き直す
         * @returns {void}
         */
        function updatePreview() {
            clearPreview();
            var oddColorName = getOddSelectedColorName();
            var evenColorName = getEvenSelectedColorName();
            var oddTint = adjustTintTextValue(parseFloat(oddTintText.text));
            var evenTint = adjustTintTextValue(parseFloat(evenTintText.text));
            var swap = swapCheckbox.value;
            var rowSkip = rowSkipCheckbox.value ? Math.max(0, parseInt(rowSkipText.text, 10) || 0) : 0;
            var colSkip = colSkipCheckbox.value ? Math.max(0, parseInt(colSkipText.text, 10) || 0) : 0;
            try {
                app.doScript(function () {
                    applyZebra(oddColorName, evenColorName, oddTint, evenTint, swap, rowSkip, colSkip);
                }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.preview"));
                state.previewed = true;
                try { doc.recompose(); } catch (e) { }
            } catch (e2) {
                state.previewed = false;
            }
        }

        // -------------------------------------------------------------
        // 濃淡入力ヘルパー：スライダーは 1 / 10 / 5 % 刻み、テキストは ↑↓ / Shift+↑↓ に対応
        // -------------------------------------------------------------
        /**
         * 修飾キーに応じた濃淡スライダーの刻み幅を返す
         * @returns {number} 刻み幅
         */
        function getTintSliderStep() {
            var ks = ScriptUI.environment.keyboardState;
            if (ks.shiftKey) return 10;
            if (ks.altKey) return 5;
            return 1;
        }

        /**
         * スライダーの値を有効範囲と刻み幅に合わせて整える
         * @param {number} v スライダーの値
         * @returns {number} 整えた値
         */
        function adjustTintSliderValue(v) {
            if (isNaN(v)) v = 100;
            if (v < 0) v = 0;
            if (v > 100) v = 100;
            var step = getTintSliderStep();
            v = Math.round(v / step) * step;
            if (v < 0) v = 0;
            if (v > 100) v = 100;
            return v;
        }

        /**
         * 入力欄の濃淡を有効範囲に収める
         * @param {number} v 入力された値
         * @returns {number} 整えた値
         */
        function adjustTintTextValue(v) {
            if (isNaN(v)) v = 100;
            if (v < 0) v = 0;
            if (v > 100) v = 100;
            return Math.round(v);
        }

        oddSlider.onChanging = function () {
            var v = adjustTintSliderValue(this.value);
            this.value = v;
            oddTintText.text = formatArrowValue(v);
            updatePreview();
        };

        oddTintText.onChange = function () {
            var v = adjustTintTextValue(parseFloat(this.text));
            this.text = formatArrowValue(v);
            oddSlider.value = v;
            updatePreview();
        };

        bindArrowKeyValueControl(oddTintText, false, function () {
            var v = adjustTintTextValue(parseFloat(oddTintText.text));
            oddTintText.text = formatArrowValue(v);
            oddSlider.value = v;
            updatePreview();
        }, { min: 0, max: 100 });

        evenSlider.onChanging = function () {
            var v = adjustTintSliderValue(this.value);
            this.value = v;
            evenTintText.text = formatArrowValue(v);
            updatePreview();
        };

        evenTintText.onChange = function () {
            var v = adjustTintTextValue(parseFloat(this.text));
            this.text = formatArrowValue(v);
            evenSlider.value = v;
            updatePreview();
        };

        bindArrowKeyValueControl(evenTintText, false, function () {
            var v = adjustTintTextValue(parseFloat(evenTintText.text));
            evenTintText.text = formatArrowValue(v);
            evenSlider.value = v;
            updatePreview();
        }, { min: 0, max: 100 });

        // -------------------------------------------------------------
        // 実行処理
        // -------------------------------------------------------------
        /**
         * 奇数行・偶数行に交互の塗りを適用する
         * @param {string} oddColorName 奇数行のスウォッチ名
         * @param {string} evenColorName 偶数行のスウォッチ名
         * @param {number} oddTint 奇数行の濃淡
         * @param {number} evenTint 偶数行の濃淡
         * @param {boolean} swap 奇数行と偶数行を入れ替えるか
         * @param {number} rowSkip 上からスキップする行数
         * @param {number} colSkip 左からスキップする列数
         * @returns {void}
         */
        function applyZebra(oddColorName, evenColorName, oddTint, evenTint, swap, rowSkip, colSkip) {
            if (swap) {
                var tmpName = oddColorName;
                oddColorName = evenColorName;
                evenColorName = tmpName;
                var tmpTint = oddTint;
                oddTint = evenTint;
                evenTint = tmpTint;
            }

            rowSkip = rowSkip || 0;
            colSkip = colSkip || 0;

            var oddSwatch = doc.swatches.item(oddColorName);
            var evenSwatch = doc.swatches.item(evenColorName);

            if (!baseTargetCells || baseTargetCells.length === 0) return;

            var targetCells = [];
            for (var i = 0; i < baseTargetCells.length; i++) {
                targetCells.push(baseTargetCells[i]);
            }

            // 選択範囲内の行・列インデックスを昇順で収集
            var uniqueRowIndices = [];
            var uniqueColIndices = [];
            var seenRows = {};
            var seenCols = {};
            for (var uri = 0; uri < targetCells.length; uri++) {
                var rIdx = targetCells[uri].parentRow.index;
                var cIdx = targetCells[uri].parentColumn.index;
                if (!seenRows[rIdx]) { seenRows[rIdx] = true; uniqueRowIndices.push(rIdx); }
                if (!seenCols[cIdx]) { seenCols[cIdx] = true; uniqueColIndices.push(cIdx); }
            }
            uniqueRowIndices.sort(function (a, b) { return a - b; });
            uniqueColIndices.sort(function (a, b) { return a - b; });

            // 選択範囲内で、上から rowSkip 行・左から colSkip 列をカラーリング対象外にする
            var rowsToSkip = {};
            for (var rs = 0; rs < rowSkip && rs < uniqueRowIndices.length; rs++) {
                rowsToSkip[String(uniqueRowIndices[rs])] = true;
            }
            var colsToSkip = {};
            for (var cs = 0; cs < colSkip && cs < uniqueColIndices.length; cs++) {
                colsToSkip[String(uniqueColIndices[cs])] = true;
            }

            // スキップされていない行に、選択範囲内で上から 0, 1, 2, ... の順番を振る
            var rowIndexMap = {};
            var orderCounter = 0;
            for (var uri2 = 0; uri2 < uniqueRowIndices.length; uri2++) {
                var ridx = uniqueRowIndices[uri2];
                if (rowsToSkip[String(ridx)]) continue;
                rowIndexMap[ridx] = orderCounter++;
            }

            for (var n = 0; n < targetCells.length; n++) {
                var cell = targetCells[n];
                var tableRowIndex = cell.parentRow.index;
                var tableColIndex = cell.parentColumn.index;

                // スキップ対象は元のカラーに戻す
                if (rowsToSkip[String(tableRowIndex)] || colsToSkip[String(tableColIndex)]) {
                    try {
                        var original = originalCellStyles[n];
                        if (original) {
                            cell.fillColor = original.color;
                            if (!isNaN(original.tint)) {
                                cell.fillTint = original.tint;
                            }
                        }
                    } catch (e) { }
                    continue;
                }

                var localRowOrder = rowIndexMap[tableRowIndex];
                if (localRowOrder === undefined) continue;

                if (localRowOrder % 2 === 0) {
                    cell.fillColor = oddSwatch;
                    if (!isNoneName(oddColorName) && !isPaperName(oddColorName)) {
                        try { cell.fillTint = oddTint; } catch (e) { }
                    }
                } else {
                    cell.fillColor = evenSwatch;
                    if (!isNoneName(evenColorName) && !isPaperName(evenColorName)) {
                        try { cell.fillTint = evenTint; } catch (e) { }
                    }
                }
            }
        }

        // 初期プレビュー
        updatePreview();

        var result = dialog.show();

        if (result == 1) {
            clearPreview();
            var finalOddColorName = getOddSelectedColorName();
            var finalEvenColorName = getEvenSelectedColorName();
            var finalOddTint = adjustTintTextValue(parseFloat(oddTintText.text));
            var finalEvenTint = adjustTintTextValue(parseFloat(evenTintText.text));
            var finalSwap = swapCheckbox.value;
            var finalRowSkip = rowSkipCheckbox.value ? Math.max(0, parseInt(rowSkipText.text, 10) || 0) : 0;
            var finalColSkip = colSkipCheckbox.value ? Math.max(0, parseInt(colSkipText.text, 10) || 0) : 0;
            app.doScript(function () {
                applyZebra(finalOddColorName, finalEvenColorName, finalOddTint, finalEvenTint,
                    finalSwap, finalRowSkip, finalColSkip);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.apply"));
        } else {
            clearPreview();
        }
    }

    main();

})();
