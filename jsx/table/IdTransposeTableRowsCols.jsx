#target indesign

/*

### 概要

選択した表の行と列を入れ替えます。ヘッダー行の扱いとセル結合の処理方法をダイアログで選べます。

詳細は README を参照してください。

### Overview

Transposes the rows and columns of the selected table. The dialog picks how header rows and merged cells are handled.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTransposeTableRowsCols";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-11-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-07";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTransposeTableRowsCols.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTransposeTableRowsCols.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nc6dbdb3af6a1"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @discussion Table Transpose (modified for robustness)
 * Original: Table Transpose v1.0 by Iain Anderson
 */

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
    try {
        if (app.locale === Locale.JAPANESE) return "ja";
    } catch (e) {}
    return (String($.locale).indexOf("ja") === 0) ? "ja" : "en";
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
    var btnRowGroup = transposeDialog.add("group");
    setupRow(btnRowGroup, "right", 8);
    btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

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
 * 2 つのオブジェクトの同名プロパティを入れ替える
 * @param {object} objectA 入れ替え元のオブジェクト
 * @param {object} objectB 入れ替え先のオブジェクト
 * @param {string[]} propertyNames 入れ替えるプロパティ名
 * @returns {void}
 */
function swapProperties(objectA, objectB, propertyNames) {
    for (var i = 0; i < propertyNames.length; i++) {
        var propertyName = propertyNames[i];
        /* 対応していないプロパティは飛ばす / Skip properties the object does not support */
        try {
            var valueBuffer = objectA[propertyName];
            objectA[propertyName] = objectB[propertyName];
            objectB[propertyName] = valueBuffer;
        } catch (e) {}
    }
}

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
        /* 文字サイズと文字色 / Point size and text fill color */
        swapProperties(paragraphA, paragraphB, ["pointSize", "fillColor"]);

        /* フォントとフォントスタイルは対で入れ替える（先にフォントを変えるとスタイルが変わるため）
           / Font and style are swapped as a pair (changing the font first can reset the style) */
        try {
            var fontBuffer      = paragraphA.appliedFont;
            var fontStyleBuffer = paragraphA.fontStyle;
            paragraphA.appliedFont = paragraphB.appliedFont;
            paragraphA.fontStyle   = paragraphB.fontStyle;
            paragraphB.appliedFont = fontBuffer;
            paragraphB.fontStyle   = fontStyleBuffer;
        } catch (e) {}
    }

    /* セルの塗り色とティント / Cell fill color and tint */
    swapProperties(cellA, cellB, ["fillColor", "fillTint"]);
}

/**
 * 転置しやすいよう、いったん表を正方形に揃える
 * @param {Table} targetTable 対象の表
 * @returns {{paddedAxis: string, originalSize: number}} 足した軸（"columns" / "rows" / "none"）と、その軸の元のサイズ
 */
function padTableToSquare(targetTable) {
    var rowCount    = targetTable.rows.length;
    var columnCount = targetTable.columnCount;

    if (rowCount > columnCount) {
        for (var addedColumn = columnCount; addedColumn < rowCount; addedColumn++) {
            targetTable.columns.add(LocationOptions.atEnd);
        }
        return { paddedAxis: "columns", originalSize: columnCount };
    }

    if (rowCount < columnCount) {
        for (var addedRow = rowCount; addedRow < columnCount; addedRow++) {
            targetTable.rows.add(LocationOptions.atEnd);
        }
        return { paddedAxis: "rows", originalSize: rowCount };
    }

    return { paddedAxis: "none", originalSize: rowCount };
}

/**
 * 空セルに代替文字を入れて段落を1つ確保する
 * @param {Table} targetTable 対象の表
 * @returns {void}
 */
function fillEmptyCells(targetTable) {
    var tableCells = targetTable.cells;
    for (var i = 0; i < tableCells.length; i++) {
        try {
            if (tableCells[i].contents === "") tableCells[i].contents = EMPTY_CELL_PLACEHOLDER;
        } catch (e) {}
    }
}

/**
 * 正方形に揃えた表の上三角と下三角を入れ替える
 * @param {Table} targetTable 正方形に揃えた表
 * @returns {void}
 */
function swapTriangles(targetTable) {
    var rowCount    = targetTable.rows.length;
    var columnCount = targetTable.columnCount;

    for (var row = 0; row < rowCount; row++) {
        for (var col = row + 1; col < columnCount; col++) {
            var upperIndex = col + (row * columnCount);
            var lowerIndex = row + (col * columnCount);
            swapCells(targetTable.cells.item(upperIndex), targetTable.cells.item(lowerIndex));
        }
    }
}

/**
 * 正方形にするため増やした行・列を取り除く
 * @param {Table} targetTable 対象の表
 * @param {string} paddedAxis padTableToSquare が返した軸
 * @param {number} originalSize 残すサイズ
 * @returns {void}
 */
function removePadding(targetTable, paddedAxis, originalSize) {
    /* 列を足した表は転置後に行が余る（その逆も同じ）/ Padding columns leaves surplus rows after transposing, and vice versa */
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
}

/**
 * ヘッダー／フッター行を元の設定に近い形で復元する
 * @param {Table} targetTable 対象の表
 * @param {number} headerRowCount 元のヘッダー行数
 * @param {number} footerRowCount 元のフッター行数
 * @returns {void}
 */
function restoreHeaderFooterRows(targetTable, headerRowCount, footerRowCount) {
    try {
        var totalRowCount  = targetTable.rows.length;
        var newHeaderCount = Math.min(headerRowCount, totalRowCount);
        var newFooterCount = Math.min(footerRowCount, Math.max(0, totalRowCount - newHeaderCount));
        targetTable.headerRowCount = newHeaderCount;
        targetTable.footerRowCount = newFooterCount;
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

    var padding = padTableToSquare(targetTable);
    fillEmptyCells(targetTable);
    swapTriangles(targetTable);
    removePadding(targetTable, padding.paddedAxis, padding.originalSize);
    restoreHeaderFooterRows(targetTable, originalHeaderRowCount, originalFooterRowCount);

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
