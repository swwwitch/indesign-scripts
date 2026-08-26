#target indesign

/*

### 概要

選択した表について、内容が同じ隣接セルを水平方向・垂直方向に自動で結合します。

詳細は README を参照してください。

### Overview

Merges adjacent cells with identical contents in the selected table, horizontally and vertically.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdCellMergeAuto";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-27";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCellMergeAuto.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCellMergeAuto.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/na84f68305844"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

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
    targetPanel.alignment = "fill";
    targetPanel.margins = PANEL_MARGINS;
    targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * 行グループの共通設定を適用する（ボタン列など）
 * @param {Group} targetGroup 対象グループ
 * @param {string} [alignment] 横方向の配置。省略時は "left"
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupRow(targetGroup, alignment, spacing) {
    targetGroup.orientation = "row";
    targetGroup.alignment = [alignment || "left", "center"];  /* 横と天地を対で / Pair the horizontal and vertical alignment */
    targetGroup.alignChildren = ["left", "center"];           /* 親の fill 継承を打ち消す / Cancel the inherited fill */
    targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var currentLang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "自動でセル結合", en: "Auto Merge Cells" }
        },
        panel: {
            mergeMode: { ja: "結合", en: "Merge" },
            direction: { ja: "方向", en: "Direction" },
            scope:     { ja: "対象", en: "Scope" }
        },
        radio: {
            singleDirection: { ja: "単一方向", en: "Single direction" },
            bothDirections:  { ja: "両方向（優先付き）", en: "Both directions (with priority)" },
            horizontal:      { ja: "水平", en: "Horizontal" },
            vertical:        { ja: "垂直", en: "Vertical" },
            wholeTable:      { ja: "表全体", en: "Whole table" },
            selectedCells:   { ja: "選択セルのみ", en: "Selected cells only" }
        },
        tooltip: {
            mergeMode: {
                ja: "「単一方向」は［方向］で選んだ側だけ、「両方向」は水平・垂直の両方を結合します。",
                en: "Single direction merges only the side chosen under Direction; Both directions merges horizontally and vertically."
            },
            direction: {
                ja: "「両方向」のときは、ここで選んだ側から先に結合します。",
                en: "In Both directions mode, the side chosen here is merged first."
            },
            scope: {
                ja: "「選択セルのみ」は、選択したセルを囲む長方形の範囲が対象になります。",
                en: "Selected cells only targets the rectangle that encloses the selected cells."
            }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            selectTable:   { ja: "表、または表の中のセルを選択してください。", en: "Select a table or cells inside a table." },
            tableNotFound: { ja: "表が見つかりません。", en: "No table was found." },
            noCellRange: {
                ja: "「選択セルのみ」が選択されていますが、有効なセル選択が見つかりませんでした。\n表全体を対象にするか、複数セルを選択してください。",
                en: "\"Selected cells only\" is chosen, but no valid cell selection was found.\nTarget the whole table or select multiple cells."
            }
        },
        undo: {
            autoMerge: { ja: "自動でセル結合", en: "Auto Merge Cells" }
        }
    };

    /**
     * ラベルを現在の言語で取得する
     * @param {object} labelEntry ja / en を持つラベルオブジェクト
     * @returns {string} 現在の言語のラベル文字列
     */
    function getLabel(labelEntry) {
        return labelEntry[currentLang];
    }

    // =========================================
    // テキスト処理 / Text helpers
    // =========================================

    /**
     * 前後の空白を取り除く（ES3 環境向けの自前実装）
     * @param {*} rawValue 対象の値
     * @returns {string} 前後の空白を除いた文字列
     */
    function customTrim(rawValue) {
        var textValue = (typeof rawValue === "string") ? rawValue : String(rawValue);
        return textValue.replace(/^\s+|\s+$/g, "");
    }

    /**
     * 比較用にセルのテキストを取り出す
     * @param {Cell} targetCell 対象のセル
     * @returns {string} 前後の空白を除いたセルの内容
     */
    function getCellText(targetCell) {
        if (!targetCell || !targetCell.isValid) return "";

        var cellContents = targetCell.contents;

        /* contents が配列で返る場合に備えて連結する / contents can come back as an array */
        if (cellContents instanceof Array) cellContents = cellContents.join("");

        return customTrim(cellContents);
    }

    // =========================================
    // 表と選択範囲の取得 / Table and range lookup
    // =========================================

    /**
     * 選択オブジェクトから対象の表を特定する
     * @param {object} selectionItem 選択オブジェクト
     * @returns {Table|null} 対象の表。特定できない場合は null
     */
    function resolveTargetTable(selectionItem) {
        if (selectionItem.constructor.name === "Table") return selectionItem;
        if (selectionItem.parent && selectionItem.parent.constructor.name === "Table") return selectionItem.parent;
        if (selectionItem.tables && selectionItem.tables.length > 0) return selectionItem.tables[0];
        if (selectionItem.parent && selectionItem.parent.parent &&
            selectionItem.parent.parent.constructor.name === "Table") {
            return selectionItem.parent.parent;
        }
        return null;
    }

    /**
     * 選択セルを囲む矩形範囲を求める
     * @param {Table} targetTable 対象の表
     * @returns {{rowStart: number, rowEnd: number, colStart: number, colEnd: number}|null} 矩形範囲。選択がなければ null
     */
    function getSelectedCellRange(targetTable) {
        var tableCells = targetTable.cells;

        var minRow = Number.MAX_VALUE;
        var maxRow = -1;
        var minCol = Number.MAX_VALUE;
        var maxCol = -1;

        for (var i = 0; i < tableCells.length; i++) {
            var targetCell = tableCells[i];
            if (!targetCell || !targetCell.isValid) continue;

            /* selected を持たないセルは常に false になるので安全に判定できる
               / Cells without a selected property simply evaluate to false */
            if (!targetCell.selected) continue;

            minRow = Math.min(minRow, targetCell.row.index);
            maxRow = Math.max(maxRow, targetCell.row.index);
            minCol = Math.min(minCol, targetCell.column.index);
            maxCol = Math.max(maxCol, targetCell.column.index);
        }

        if (maxRow < 0) return null;

        return { rowStart: minRow, rowEnd: maxRow, colStart: minCol, colEnd: maxCol };
    }

    // =========================================
    // セル結合 / Cell merging
    // =========================================

    /**
     * 走査方向と直交する側のインデックスを取り出す
     * @param {Cell} targetCell 対象のセル
     * @param {boolean} isHorizontal true なら水平方向の走査
     * @returns {number} 水平方向なら列インデックス、垂直方向なら行インデックス
     */
    function getCrossIndex(targetCell, isHorizontal) {
        return isHorizontal ? targetCell.column.index : targetCell.row.index;
    }

    /**
     * 内容が同じ隣接セルを指定方向に結合する
     * @param {Table} targetTable 対象の表
     * @param {object|null} cellRange 対象の矩形範囲。null なら表全体
     * @param {boolean} isHorizontal true なら水平方向、false なら垂直方向
     * @returns {void}
     */
    function mergeSameContentCells(targetTable, cellRange, isHorizontal) {
        /* 水平方向なら行を、垂直方向なら列を 1 本ずつ走査する
           / Scan row by row when horizontal, column by column when vertical */
        var scanLines = isHorizontal ? targetTable.rows : targetTable.columns;

        /* 走査する行・列の範囲と、その中で対象になるセルの範囲
           / Bounds for the lines to scan, and for the cells inside a line */
        var targetRange = cellRange ||
            { rowStart: 0, rowEnd: Number.MAX_VALUE, colStart: 0, colEnd: Number.MAX_VALUE };
        var lineStart = isHorizontal ? targetRange.rowStart : targetRange.colStart;
        var lineEnd   = isHorizontal ? targetRange.rowEnd   : targetRange.colEnd;
        var cellStart = isHorizontal ? targetRange.colStart : targetRange.rowStart;
        var cellEnd   = isHorizontal ? targetRange.colEnd   : targetRange.rowEnd;

        for (var lineIndex = lineStart; lineIndex < scanLines.length && lineIndex <= lineEnd; lineIndex++) {
            var currentLine = scanLines[lineIndex];
            if (!currentLine.isValid) continue;

            var lineCells = currentLine.cells;
            var cellIndex = 0;

            while (cellIndex < lineCells.length - 1) {
                var headCell = lineCells[cellIndex];
                var nextCell = lineCells[cellIndex + 1];

                if (!headCell || !headCell.isValid || !nextCell || !nextCell.isValid ||
                    getCrossIndex(headCell, isHorizontal) < cellStart ||
                    getCrossIndex(nextCell, isHorizontal) > cellEnd) {
                    cellIndex++;
                    continue;
                }

                var headText = getCellText(headCell);
                if (headText !== getCellText(nextCell)) {
                    cellIndex++;
                    continue;
                }

                /* またぎ方が揃わないセル同士は merge() が失敗するので、その組は飛ばす
                   / merge() fails when the spans do not line up, so skip that pair */
                try {
                    headCell.merge(nextCell);
                } catch (e) {
                    cellIndex++;
                    continue;
                }

                headCell.contents = headText;
                lineCells = currentLine.cells; /* 結合でセル配列が変わるため取り直す / Refresh after the merge */
            }
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ラジオボタンを縦に並べたパネルを追加する
     * @param {Window|Group} parentContainer 追加先
     * @param {object} panelLabel パネル見出しのラベル
     * @param {Array.<object>} radioLabels ラジオボタンのラベル配列
     * @param {number} selectedIndex 最初に選択するラジオボタンのインデックス
     * @param {object} [tooltipLabel] パネルとラジオボタンに付けるツールチップのラベル
     * @returns {Array.<RadioButton>} 追加したラジオボタン
     */
    function addRadioPanel(parentContainer, panelLabel, radioLabels, selectedIndex, tooltipLabel) {
        var radioPanel = parentContainer.add("panel", undefined, getLabel(panelLabel));
        setupPanel(radioPanel, 6);
        radioPanel.alignChildren = ["left", "top"];

        /* パネルだけに付けると枠の上でしか出ないので、ラジオボタンにも同じ説明を持たせる
           / A panel-only helpTip shows on the frame alone, so give the radio buttons the same text */
        var tooltipText = tooltipLabel ? getLabel(tooltipLabel) : "";
        radioPanel.helpTip = tooltipText;

        var radioButtons = [];
        for (var i = 0; i < radioLabels.length; i++) {
            radioButtons[i] = radioPanel.add("radiobutton", undefined, getLabel(radioLabels[i]));
            radioButtons[i].helpTip = tooltipText;
        }
        radioButtons[selectedIndex].value = true;

        return radioButtons;
    }

    /**
     * 結合方法を指定するダイアログを表示する
     * @returns {{passes: Array.<boolean>, useSelectionOnly: boolean}|null} 設定内容。キャンセル時は null
     */
    function showAutoMergeDialog() {
        var autoMergeDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(autoMergeDialog, 10);

        var settingsColumn = autoMergeDialog.add("group");
        settingsColumn.orientation = "column";
        settingsColumn.alignChildren = ["fill", "top"];
        settingsColumn.spacing = PANEL_SPACING;

        var mergeModeRadios = addRadioPanel(settingsColumn, LABELS.panel.mergeMode,
            [LABELS.radio.singleDirection, LABELS.radio.bothDirections], 1, LABELS.tooltip.mergeMode);
        var directionRadios = addRadioPanel(settingsColumn, LABELS.panel.direction,
            [LABELS.radio.horizontal, LABELS.radio.vertical], 0, LABELS.tooltip.direction);
        var scopeRadios = addRadioPanel(settingsColumn, LABELS.panel.scope,
            [LABELS.radio.wholeTable, LABELS.radio.selectedCells], 0, LABELS.tooltip.scope);

        /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
        var btnRowGroup = autoMergeDialog.add("group");
        setupRow(btnRowGroup, "right", 8);
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        if (autoMergeDialog.show() !== 1) return null;

        var useBothDirections = mergeModeRadios[1].value;
        var startsHorizontal = directionRadios[0].value;

        /* 結合を試す順番。true は水平、false は垂直 / Passes to run; true is horizontal, false is vertical */
        return {
            passes: useBothDirections ? [startsHorizontal, !startsHorizontal] : [startsHorizontal],
            useSelectionOnly: scopeRadios[1].value
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログの設定に従って同一内容のセルを自動結合する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        if (app.selection.length === 0) {
            alert(getLabel(LABELS.alert.selectTable));
            return;
        }

        var targetTable = resolveTargetTable(app.selection[0]);
        if (!targetTable) {
            alert(getLabel(LABELS.alert.tableNotFound));
            return;
        }

        var dialogResult = showAutoMergeDialog();
        if (!dialogResult) return;

        /* 対象範囲。null なら表全体 / The target range; null means the whole table */
        var cellRange = null;
        if (dialogResult.useSelectionOnly) {
            cellRange = getSelectedCellRange(targetTable);
            if (!cellRange) {
                alert(getLabel(LABELS.alert.noCellRange));
                return;
            }
        }

        /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
        app.doScript(function () {
            for (var i = 0; i < dialogResult.passes.length; i++) {
                mergeSameContentCells(targetTable, cellRange, dialogResult.passes[i]);
            }
        }, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undo.autoMerge));
    }

    main();

})();
