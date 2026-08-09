#target indesign

/*
 * swap-table-row-column.jsx
 *
 * 選択した表の行と列を入れ替えます。ヘッダー行の扱いとセル結合の処理方法をダイアログで選べます。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "swap-table-row-column";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-11-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/swap-table-row-column.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/swap-table-row-column.md

// Original idea
// Table Transpose v1.0 by Iain Anderson

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* セル結合の扱い / How merged cells are handled */
var MERGE_MODE_STOP    = "stop";     /* 結合があれば中止 / Cancel when merged cells exist */
var MERGE_MODE_UNMERGE = "unmerge";  /* 転置前に結合を解除 / Unmerge before transposing */

/* 空セルに入れる代替文字（段落を1つ確保するため）/ Placeholder put in empty cells so each has one paragraph */
var EMPTY_CELL_PLACEHOLDER = " ";

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
    var isJapanese = false;
    try {
        if (app.locale === Locale.JAPANESE) isJapanese = true;
    } catch (e) {}
    try {
        if (!isJapanese && String($.locale).indexOf("ja") === 0) isJapanese = true;
    } catch (e) {}
    return isJapanese ? "ja" : "en";
}

var currentLang = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "行と列を入れ替え", en: "Transpose Rows and Columns" }
    },
    panel: {
        mergedCells: { ja: "セル結合", en: "Merged cells" }
    },
    checkbox: {
        includeHeader: { ja: "ヘッダー行を対象にする", en: "Include header rows" }
    },
    radio: {
        mergeStop:    { ja: "しない（終了）", en: "Do nothing (cancel)" },
        mergeUnmerge: { ja: "転置前にセル結合を解除", en: "Unmerge before transposing" }
    },
    button: {
        ok:     { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    alert: {
        noDocument:  { ja: "ドキュメントが開いていません。", en: "No document is open." },
        noSelection: {
            ja: "表、または表を含むテキストフレームを選択してから実行してください。",
            en: "Please select a table or a text frame containing a table, then run this script."
        },
        noTable: {
            ja: "選択範囲から表を特定できませんでした。\n表、セル、または表内のテキストを選択して再度お試しください。",
            en: "Could not find a table from the selection.\nPlease select a table, cell, or text inside a table and try again."
        },
        mergeStopped: {
            ja: "セル結合があるため、処理を中止しました。\nセル結合の扱いを変更して再度お試しください。",
            en: "The table contains merged cells. The operation has been cancelled.\nChange how merged cells are handled and try again."
        }
    },
    undo: {
        transposeTable: { ja: "行と列を入れ替え", en: "Transpose Rows and Columns" }
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
// 表の取得 / Table lookup
// =========================================

/**
 * 選択オブジェクトから対象の表を特定する
 * @param {object} selectionItem 選択オブジェクト
 * @returns {Table|null} 対象の表。特定できない場合は null
 */
function resolveTableFromSelection(selectionItem) {
    if (!selectionItem) return null;

    if (selectionItem.constructor.name === "Table") return selectionItem;

    /* セル選択時の parent は Table / The parent of a selected cell is the table */
    if (selectionItem.constructor.name === "Cell") return selectionItem.parent;

    try {
        if (selectionItem.tables && selectionItem.tables.length > 0) return selectionItem.tables[0];
    } catch (e) {}

    try {
        if (selectionItem.parent && selectionItem.parent.tables && selectionItem.parent.tables.length > 0) {
            return selectionItem.parent.tables[0];
        }
    } catch (e) {}

    return null;
}

/**
 * 表にセル結合があるかを判定する
 * @param {Table} targetTable 対象の表
 * @returns {boolean} 結合セルがあれば true
 */
function hasMergedCells(targetTable) {
    try {
        var tableCells = targetTable.cells;
        for (var i = 0; i < tableCells.length; i++) {
            if (tableCells[i].rowSpan > 1 || tableCells[i].columnSpan > 1) return true;
        }
    } catch (e) {}
    return false;
}

/**
 * セルの最初の段落を取得する
 * @param {Cell} targetCell 対象のセル
 * @returns {Paragraph|null} 最初の段落。取得できない場合は null
 */
function getFirstParagraph(targetCell) {
    try {
        if (targetCell.paragraphs.length > 0) return targetCell.paragraphs.item(0);
        if (targetCell.texts && targetCell.texts.length > 0 && targetCell.texts[0].paragraphs.length > 0) {
            return targetCell.texts[0].paragraphs.item(0);
        }
    } catch (e) {}
    return null;
}

// =========================================
// ダイアログ / Dialog
// =========================================

/**
 * ヘッダー行とセル結合の扱いを尋ねるダイアログを表示する
 * @param {boolean} tableHasHeader ヘッダー行があるか
 * @param {boolean} tableHasMerge セル結合があるか
 * @returns {{includeHeader: boolean, mergeMode: string}|null} 設定内容。キャンセル時は null
 */
function showTransposeDialog(tableHasHeader, tableHasMerge) {
    var transposeDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(transposeDialog);
    transposeDialog.alignChildren = ["left", "top"];

    var includeHeaderCheckbox = transposeDialog.add("checkbox", undefined, getLabel("checkbox.includeHeader"));
    includeHeaderCheckbox.value = true;
    if (!tableHasHeader) {
        includeHeaderCheckbox.enabled = false;
        includeHeaderCheckbox.value = false;
    }

    /* セル結合の扱いパネル / Panel for merged-cell handling */
    var mergedCellsPanel = transposeDialog.add("panel", undefined, getLabel("panel.mergedCells"));
    setupPanel(mergedCellsPanel, 6);
    mergedCellsPanel.alignChildren = ["left", "top"];

    var mergeStopRadio    = mergedCellsPanel.add("radiobutton", undefined, getLabel("radio.mergeStop"));
    var mergeUnmergeRadio = mergedCellsPanel.add("radiobutton", undefined, getLabel("radio.mergeUnmerge"));
    mergeUnmergeRadio.value = true;
    if (!tableHasMerge) {
        mergeStopRadio.enabled = false;
        mergeUnmergeRadio.enabled = false;
    }

    /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
    var dialogButtonRow = transposeDialog.add("group");
    setupRow(dialogButtonRow, "right", 8);
    dialogButtonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    dialogButtonRow.add("button", undefined, getLabel("button.ok"), { name: "ok" });

    if (transposeDialog.show() !== 1) return null;

    return {
        includeHeader: includeHeaderCheckbox.value,
        mergeMode: mergeStopRadio.value ? MERGE_MODE_STOP : MERGE_MODE_UNMERGE
    };
}

// =========================================
// 転置処理 / Transpose
// =========================================

/**
 * 2 つのセルの内容と書式を入れ替える
 * @param {Cell} cellA 入れ替え元のセル
 * @param {Cell} cellB 入れ替え先のセル
 * @returns {void}
 */
function swapCells(cellA, cellB) {
    /* テキスト内容 / Cell contents */
    var contentsBuffer = cellA.contents;
    cellA.contents = cellB.contents;
    cellB.contents = contentsBuffer;

    var paragraphA = getFirstParagraph(cellA);
    var paragraphB = getFirstParagraph(cellB);

    if (paragraphA && paragraphB) {
        /* 文字サイズ / Point size */
        try {
            var pointSizeBuffer = paragraphA.pointSize;
            paragraphA.pointSize = paragraphB.pointSize;
            paragraphB.pointSize = pointSizeBuffer;
        } catch (e) {}

        /* フォントとフォントスタイル / Font and font style */
        try {
            var fontBuffer      = paragraphA.appliedFont;
            var fontStyleBuffer = paragraphA.fontStyle;
            paragraphA.appliedFont = paragraphB.appliedFont;
            paragraphA.fontStyle   = paragraphB.fontStyle;
            paragraphB.appliedFont = fontBuffer;
            paragraphB.fontStyle   = fontStyleBuffer;
        } catch (e) {}

        /* 文字色 / Text fill color */
        try {
            var textFillBuffer = paragraphA.fillColor;
            paragraphA.fillColor = paragraphB.fillColor;
            paragraphB.fillColor = textFillBuffer;
        } catch (e) {}
    }

    /* セルの塗り色 / Cell fill color */
    try {
        var cellFillBuffer = cellA.fillColor;
        cellA.fillColor = cellB.fillColor;
        cellB.fillColor = cellFillBuffer;
    } catch (e) {}

    /* セルのティント / Cell fill tint */
    try {
        var fillTintBuffer = cellA.fillTint;
        cellA.fillTint = cellB.fillTint;
        cellB.fillTint = fillTintBuffer;
    } catch (e) {}
}

/**
 * 表の行と列を入れ替える
 * @param {Table} targetTable 対象の表
 * @param {boolean} includeHeader ヘッダー行も転置対象にするか
 * @param {string} mergeMode MERGE_MODE_STOP または MERGE_MODE_UNMERGE
 * @returns {string} "ok" または "mergeStopped"
 */
function transposeTable(targetTable, includeHeader, mergeMode) {
    /* 元のヘッダー／フッター行数を控える / Remember the original header and footer row counts */
    var originalHeaderRowCount = includeHeader ? targetTable.headerRowCount : 0;
    var originalFooterRowCount = targetTable.footerRowCount;

    if (hasMergedCells(targetTable)) {
        if (mergeMode === MERGE_MODE_STOP) return "mergeStopped";
        try {
            targetTable.unmerge();
        } catch (e) {
            /* 解除できなくても単純な表なら続行できる / Simple tables can continue even if unmerge fails */
        }
    }

    var rowCount    = targetTable.rows.length;
    var columnCount = targetTable.columnCount;
    var originalSize = 0;
    var paddedAxis   = "none"; /* "columns" / "rows" / "none" */

    /* 転置しやすいよう、いったん正方形に揃える / Pad the table to a square so it can be transposed in place */
    if (rowCount > columnCount) {
        for (var addedColumn = columnCount; addedColumn < rowCount; addedColumn++) {
            targetTable.columns.add(LocationOptions.atEnd);
        }
        originalSize = columnCount;
        paddedAxis   = "columns";
    } else if (rowCount < columnCount) {
        for (var addedRow = rowCount; addedRow < columnCount; addedRow++) {
            targetTable.rows.add(LocationOptions.atEnd);
        }
        originalSize = rowCount;
        paddedAxis   = "rows";
    }

    rowCount    = targetTable.rows.length;
    columnCount = targetTable.columnCount;

    /* 空セルに段落を1つ確保する / Ensure every cell has at least one paragraph */
    var totalCellCount = rowCount * columnCount;
    for (var i = 0; i < totalCellCount; i++) {
        try {
            var currentCell = targetTable.cells.item(i);
            if (currentCell.contents === "") currentCell.contents = EMPTY_CELL_PLACEHOLDER;
        } catch (e) {}
    }

    /* 上三角と下三角を入れ替える / Swap the upper and lower triangles */
    for (var row = 0; row < rowCount; row++) {
        for (var col = row + 1; col < columnCount; col++) {
            var upperIndex = col + (row * columnCount);
            var lowerIndex = row + (col * columnCount);
            swapCells(targetTable.cells.item(upperIndex), targetTable.cells.item(lowerIndex));
        }
    }

    /* 正方形にするため増やした行・列を戻す / Remove the rows or columns added for padding */
    if (paddedAxis === "columns") {
        while (targetTable.rows.length > originalSize) {
            try {
                targetTable.rows.lastItem().remove();
            } catch (e) {
                break;
            }
        }
    } else if (paddedAxis === "rows") {
        while (targetTable.columnCount > originalSize) {
            try {
                targetTable.columns.lastItem().remove();
            } catch (e) {
                break;
            }
        }
    }

    /* ヘッダー／フッター行を元の設定に近い形で復元 / Restore header and footer rows as closely as possible */
    try {
        var totalRowCount  = targetTable.rows.length;
        var newHeaderCount = Math.min(originalHeaderRowCount, totalRowCount);
        var newFooterCount = Math.min(originalFooterRowCount, Math.max(0, totalRowCount - newHeaderCount));
        targetTable.headerRowCount = newHeaderCount;
        targetTable.footerRowCount = newFooterCount;
    } catch (e) {}

    return "ok";
}

// =========================================
// メイン処理 / Main
// =========================================

if (app.documents.length === 0) {
    alert(getLabel("alert.noDocument"));
    return;
}

if (app.selection.length === 0) {
    alert(getLabel("alert.noSelection"));
    return;
}

var targetTable = resolveTableFromSelection(app.selection[0]);
if (!targetTable) {
    alert(getLabel("alert.noTable"));
    return;
}

var tableHasMerge  = hasMergedCells(targetTable);
var tableHasHeader = (targetTable.headerRowCount > 0);

var includeHeader = false;
var mergeMode     = MERGE_MODE_UNMERGE;

/* 選択の余地がない場合はダイアログを省略 / Skip the dialog when there is nothing to choose */
if (tableHasHeader || tableHasMerge) {
    var dialogResult = showTransposeDialog(tableHasHeader, tableHasMerge);
    if (dialogResult === null) return;
    includeHeader = dialogResult.includeHeader;
    mergeMode     = dialogResult.mergeMode;
}

/* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
var transposeStatus = app.doScript(function () {
    return transposeTable(targetTable, includeHeader, mergeMode);
}, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.transposeTable"));

if (transposeStatus === "mergeStopped") {
    alert(getLabel("alert.mergeStopped"));
}

})();
