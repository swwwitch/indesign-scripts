#target indesign

/*
 * MergeCell-Auto.jsx
 *
 * 選択した表について、内容が同じ隣接セルを水平方向・垂直方向に自動で結合します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "MergeCell-Auto";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v0.2";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/MergeCell-Auto.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/MergeCell-Auto.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 結合の進め方 / How the merge is carried out */
var MERGE_MODE_HORIZONTAL_ONLY  = "horizontalOnly";
var MERGE_MODE_VERTICAL_ONLY    = "verticalOnly";
var MERGE_MODE_HORIZONTAL_FIRST = "horizontalFirst";
var MERGE_MODE_VERTICAL_FIRST   = "verticalFirst";

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
        title: { ja: "自動でセル結合", en: "Auto Merge Cells" }
    },
    panel: {
        mergeMode: { ja: "結合", en: "Merge" },
        direction: { ja: "方向", en: "Direction" },
        scope:     { ja: "対象", en: "Scope" }
    },
    radio: {
        singleDirection:  { ja: "単一方向", en: "Single direction" },
        bothDirections:   { ja: "両方向（優先付き）", en: "Both directions (with priority)" },
        horizontal:       { ja: "水平", en: "Horizontal" },
        vertical:         { ja: "垂直", en: "Vertical" },
        wholeTable:       { ja: "表全体", en: "Whole table" },
        selectedCells:    { ja: "選択セルのみ", en: "Selected cells only" }
    },
    button: {
        ok:     { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    alert: {
        selectTable:  { ja: "表、または表の中のセルを選択してください。", en: "Select a table or cells inside a table." },
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
function localize(labelEntry) {
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
 * @param {Cell} cell 対象のセル
 * @returns {string} 前後の空白を除いたセルの内容
 */
function getCellText(cell) {
    if (!cell || !cell.isValid) return "";

    var cellContents = cell.contents;

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
    if (!tableCells || tableCells.length === 0) return null;

    var minRow = Number.MAX_VALUE;
    var maxRow = -1;
    var minCol = Number.MAX_VALUE;
    var maxCol = -1;

    for (var i = 0; i < tableCells.length; i++) {
        var cell = tableCells[i];
        if (!cell || !cell.isValid) continue;

        /* selected を持たないセルは常に false になるので安全に判定できる
           / Cells without a selected property simply evaluate to false */
        if (!cell.selected) continue;

        var rowIndex = cell.row.index;
        var colIndex = cell.column.index;

        if (rowIndex < minRow) minRow = rowIndex;
        if (rowIndex > maxRow) maxRow = rowIndex;
        if (colIndex < minCol) minCol = colIndex;
        if (colIndex > maxCol) maxCol = colIndex;
    }

    if (maxRow < 0 || maxCol < 0) return null;

    return { rowStart: minRow, rowEnd: maxRow, colStart: minCol, colEnd: maxCol };
}

// =========================================
// セル結合 / Cell merging
// =========================================

/**
 * 内容が同じ隣接セルを水平方向に結合する
 * @param {Table} targetTable 対象の表
 * @param {object|null} cellRange 対象の矩形範囲。null なら表全体
 * @returns {void}
 */
function mergeHorizontalSafe(targetTable, cellRange) {
    var tableRows = targetTable.rows;

    for (var rowIndex = 0; rowIndex < tableRows.length; rowIndex++) {
        var currentRow = tableRows[rowIndex];
        if (!currentRow.isValid) continue;
        if (cellRange && (rowIndex < cellRange.rowStart || rowIndex > cellRange.rowEnd)) continue;

        var rowCells = currentRow.cells;
        var cellIndex = 0;

        while (cellIndex < rowCells.length - 1) {
            var leftCell  = rowCells[cellIndex];
            var rightCell = rowCells[cellIndex + 1];

            if (!leftCell || !leftCell.isValid || !rightCell || !rightCell.isValid) {
                cellIndex++;
                continue;
            }

            if (cellRange &&
                (leftCell.column.index < cellRange.colStart || rightCell.column.index > cellRange.colEnd)) {
                cellIndex++;
                continue;
            }

            var leftText  = getCellText(leftCell);
            var rightText = getCellText(rightCell);

            if (leftText !== rightText) {
                cellIndex++;
                continue;
            }

            try {
                leftCell.merge(rightCell);
                leftCell.contents = leftText;
                rowCells = currentRow.cells; /* 結合でセル配列が変わるため取り直す / Refresh after the merge */
            } catch (e) {
                cellIndex++;
            }
        }
    }
}

/**
 * 内容が同じ隣接セルを垂直方向に結合する
 * @param {Table} targetTable 対象の表
 * @param {object|null} cellRange 対象の矩形範囲。null なら表全体
 * @returns {void}
 */
function mergeVerticalSafe(targetTable, cellRange) {
    for (var colIndex = 0; colIndex < targetTable.columns.length; colIndex++) {
        var currentColumn = targetTable.columns[colIndex];
        if (!currentColumn.isValid) continue;
        if (cellRange && (colIndex < cellRange.colStart || colIndex > cellRange.colEnd)) continue;

        var columnCells = currentColumn.cells;
        var cellIndex = 0;

        while (cellIndex < columnCells.length - 1) {
            var upperCell = columnCells[cellIndex];
            var lowerCell = columnCells[cellIndex + 1];

            if (!upperCell || !upperCell.isValid || !lowerCell || !lowerCell.isValid) {
                cellIndex++;
                continue;
            }

            if (cellRange &&
                (upperCell.row.index < cellRange.rowStart || lowerCell.row.index > cellRange.rowEnd)) {
                cellIndex++;
                continue;
            }

            var upperText = getCellText(upperCell);
            var lowerText = getCellText(lowerCell);

            if (upperText !== lowerText) {
                cellIndex++;
                continue;
            }

            try {
                upperCell.merge(lowerCell);
                upperCell.contents = upperText;
                columnCells = currentColumn.cells; /* 結合でセル配列が変わるため取り直す / Refresh after the merge */
            } catch (e) {
                cellIndex++;
            }
        }
    }
}

// =========================================
// ダイアログ / Dialog
// =========================================

/**
 * 結合方法を指定するダイアログを表示する
 * @returns {{mergeMode: string, useSelectionOnly: boolean}|null} 設定内容。キャンセル時は null
 */
function showAutoMergeDialog() {
    var autoMergeDialog = new Window("dialog", localize(LABELS.dialog.title) + " " + SCRIPT_VERSION);
    setupWindow(autoMergeDialog, 10);

    var settingsRow = autoMergeDialog.add("group");
    setupRow(settingsRow, "fill", COLUMN_SPACING);
    settingsRow.alignChildren = ["fill", "top"];

    var settingsColumn = settingsRow.add("group");
    settingsColumn.orientation = "column";
    settingsColumn.alignChildren = ["fill", "top"];
    settingsColumn.spacing = PANEL_SPACING;

    /* 結合パネル / Merge-mode panel */
    var mergeModePanel = settingsColumn.add("panel", undefined, localize(LABELS.panel.mergeMode));
    setupPanel(mergeModePanel, 6);
    mergeModePanel.alignChildren = ["left", "top"];

    var singleDirectionRadio = mergeModePanel.add("radiobutton", undefined, localize(LABELS.radio.singleDirection));
    var bothDirectionsRadio  = mergeModePanel.add("radiobutton", undefined, localize(LABELS.radio.bothDirections));
    bothDirectionsRadio.value = true;

    /* 方向パネル / Direction panel */
    var directionPanel = settingsColumn.add("panel", undefined, localize(LABELS.panel.direction));
    setupPanel(directionPanel, 6);
    directionPanel.alignChildren = ["left", "top"];

    var horizontalRadio = directionPanel.add("radiobutton", undefined, localize(LABELS.radio.horizontal));
    horizontalRadio.value = true;
    directionPanel.add("radiobutton", undefined, localize(LABELS.radio.vertical));

    /* 対象パネル / Scope panel */
    var scopePanel = settingsColumn.add("panel", undefined, localize(LABELS.panel.scope));
    setupPanel(scopePanel, 6);
    scopePanel.alignChildren = ["left", "top"];

    var wholeTableRadio = scopePanel.add("radiobutton", undefined, localize(LABELS.radio.wholeTable));
    wholeTableRadio.value = true;
    scopePanel.add("radiobutton", undefined, localize(LABELS.radio.selectedCells));

    /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
    var dialogButtonRow = autoMergeDialog.add("group");
    setupRow(dialogButtonRow, "right", 8);
    dialogButtonRow.add("button", undefined, localize(LABELS.button.cancel), { name: "cancel" });
    dialogButtonRow.add("button", undefined, localize(LABELS.button.ok), { name: "ok" });

    if (autoMergeDialog.show() !== 1) return null;

    var mergeMode;
    if (singleDirectionRadio.value) {
        mergeMode = horizontalRadio.value ? MERGE_MODE_HORIZONTAL_ONLY : MERGE_MODE_VERTICAL_ONLY;
    } else {
        mergeMode = horizontalRadio.value ? MERGE_MODE_HORIZONTAL_FIRST : MERGE_MODE_VERTICAL_FIRST;
    }

    return { mergeMode: mergeMode, useSelectionOnly: !wholeTableRadio.value };
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
        alert(localize(LABELS.alert.selectTable));
        return;
    }

    var targetTable = resolveTargetTable(app.selection[0]);
    if (!targetTable) {
        alert(localize(LABELS.alert.tableNotFound));
        return;
    }

    var dialogResult = showAutoMergeDialog();
    if (dialogResult === null) return;

    /* 対象範囲。null なら表全体 / The target range; null means the whole table */
    var cellRange = null;
    if (dialogResult.useSelectionOnly) {
        cellRange = getSelectedCellRange(targetTable);
        if (!cellRange) {
            alert(localize(LABELS.alert.noCellRange));
            return;
        }
    }

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(function () {
        switch (dialogResult.mergeMode) {
            case MERGE_MODE_HORIZONTAL_ONLY:
                mergeHorizontalSafe(targetTable, cellRange);
                break;
            case MERGE_MODE_VERTICAL_ONLY:
                mergeVerticalSafe(targetTable, cellRange);
                break;
            case MERGE_MODE_HORIZONTAL_FIRST:
                mergeHorizontalSafe(targetTable, cellRange);
                mergeVerticalSafe(targetTable, cellRange);
                break;
            case MERGE_MODE_VERTICAL_FIRST:
                mergeVerticalSafe(targetTable, cellRange);
                mergeHorizontalSafe(targetTable, cellRange);
                break;
        }
    }, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, localize(LABELS.undo.autoMerge));
}

main();
