#target indesign

/*

### 概要

表の中を選択して実行すると、その表全体の行と列を入れ替えます。ヘッダー行の指定を引き継ぐかと、セル結合の扱いをダイアログで選べます。

詳細は README を参照してください。

### Overview

Transposes the whole table the selection sits in, however much of it is selected. The dialog chooses whether to keep the header row setting and how merged cells are handled.

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

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTransposeTableRowsCols.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTransposeTableRowsCols.md"; /* README (English) */
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
        mergedCells: { ja: "セル結合があるとき", en: "When merged cells exist" }
    },
    checkbox: {
        keepHeaderRows: { ja: "ヘッダー行の指定を引き継ぐ", en: "Keep header row setting" }
    },
    radio: {
        mergeStop:    { ja: "処理を中止", en: "Cancel the operation" },
        mergeUnmerge: { ja: "解除してから入れ替え", en: "Unmerge, then transpose" }
    },
    tooltip: {
        keepHeaderRows: {
            ja: "オンにすると、転置後の表でも先頭から同じ行数をヘッダー行に設定します。\n転置の対象は、オン・オフにかかわらず常に表全体です。",
            en: "When on, the same number of leading rows becomes header rows after transposing.\nThe whole table is transposed either way."
        },
        mergeStop: {
            ja: "結合セルが1つでもあれば、何も変更せずに終了します。",
            en: "If the table contains any merged cell, nothing is changed and the script stops."
        },
        mergeUnmerge: {
            ja: "すべてのセル結合を解除してから入れ替えます。結合は復元されません。",
            en: "Unmerges every merged cell, then transposes. The merges are not restored."
        }
    },
    button: {
        ok:     { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    alert: {
        noDocument:  { ja: "ドキュメントが開いていません。", en: "No document is open." },
        noSelection: {
            ja: "表、セル、または表内のテキストを選択してから実行してください。",
            en: "Please select a table, cell, or text inside a table, then run this script."
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
    var labelNode = LABELS;
    var keyParts = labelKey.split(".");
    for (var i = 0; i < keyParts.length; i++) {
        labelNode = labelNode[keyParts[i]];
        if (!labelNode) return labelKey;
    }
    return labelNode[currentLang] || labelNode.en || labelKey;
}

// =========================================
// 表の取得 / Table lookup
// =========================================

/* 親をたどる上限。表 → セル → 表 の入れ子でも十分な深さ / Parent-walk limit, deep enough for nested tables */
var PARENT_LOOKUP_LIMIT = 8;

/**
 * 選択の祖先をたどって表そのものを探す
 * @param {object} selectionItem 選択オブジェクト
 * @returns {Table|null} 選択を含む表。見つからない場合は null
 */
function findAncestorTable(selectionItem) {
    var candidate = selectionItem;
    for (var i = 0; i < PARENT_LOOKUP_LIMIT; i++) {
        if (!candidate) return null;
        if (candidate.constructor.name === "Table") return candidate;
        try {
            candidate = candidate.parent;
        } catch (e) {
            return null;
        }
    }
    return null;
}

/**
 * 選択とその祖先が持っている表を探す
 * @param {object} selectionItem 選択オブジェクト
 * @returns {Table|null} 最初に見つかった表。見つからない場合は null
 */
function findContainedTable(selectionItem) {
    var candidate = selectionItem;
    for (var i = 0; i < PARENT_LOOKUP_LIMIT; i++) {
        if (!candidate) return null;
        try {
            if (candidate.tables && candidate.tables.length > 0) return candidate.tables[0];
        } catch (e) {}
        try {
            candidate = candidate.parent;
        } catch (e) {
            return null;
        }
    }
    return null;
}

/**
 * 選択オブジェクトから対象の表を特定する
 * カーソル位置、セル内の数文字、複数セル、全セルのいずれでも、それを含む表全体を返す。
 * @param {object} selectionItem 選択オブジェクト
 * @returns {Table|null} 対象の表。特定できない場合は null
 */
function resolveTableFromSelection(selectionItem) {
    if (!selectionItem) return null;

    /* 祖先の表を優先する。先に tables を見ると、入れ子の表や同じストーリー内の別の表を掴む
       / An ancestor table wins: checking tables first would grab a nested table, or another table in the same story */
    return findAncestorTable(selectionItem) || findContainedTable(selectionItem);
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
    } catch (e) {}
    return null;
}

// =========================================
// ダイアログ / Dialog
// =========================================

/**
 * ヘッダー行とセル結合の扱いを尋ねるダイアログを表示する
 * @param {boolean} tableHasHeaderRows ヘッダー行があるか
 * @param {boolean} tableHasMergedCells セル結合があるか
 * @returns {{keepHeaderRows: boolean, mergeMode: string}|null} 設定内容。キャンセル時は null
 */
function showTransposeDialog(tableHasHeaderRows, tableHasMergedCells) {
    var transposeDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(transposeDialog);
    transposeDialog.alignChildren = ["left", "top"];

    var keepHeaderRowsCheckbox = transposeDialog.add("checkbox", undefined, getLabel("checkbox.keepHeaderRows"));
    keepHeaderRowsCheckbox.helpTip = getLabel("tooltip.keepHeaderRows");
    keepHeaderRowsCheckbox.value = true;
    if (!tableHasHeaderRows) {
        keepHeaderRowsCheckbox.enabled = false;
        keepHeaderRowsCheckbox.value = false;
    }

    /* セル結合の扱いパネル / Panel for merged-cell handling */
    var mergedCellsPanel = transposeDialog.add("panel", undefined, getLabel("panel.mergedCells"));
    setupPanel(mergedCellsPanel, 6);
    mergedCellsPanel.alignChildren = ["left", "top"];

    var mergeStopRadio    = mergedCellsPanel.add("radiobutton", undefined, getLabel("radio.mergeStop"));
    var mergeUnmergeRadio = mergedCellsPanel.add("radiobutton", undefined, getLabel("radio.mergeUnmerge"));
    mergeStopRadio.helpTip    = getLabel("tooltip.mergeStop");
    mergeUnmergeRadio.helpTip = getLabel("tooltip.mergeUnmerge");
    mergeUnmergeRadio.value = true;

    /* 結合がない表では選ぶ余地がないのでパネルごと無効にする / Disable the whole panel when there is nothing to choose */
    if (!tableHasMergedCells) mergedCellsPanel.enabled = false;

    /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
    var btnRowGroup = transposeDialog.add("group");
    setupRow(btnRowGroup, "right", 8);
    btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

    if (transposeDialog.show() !== 1) return null;

    return {
        keepHeaderRows: keepHeaderRowsCheckbox.value,
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
 * 行または列を末尾に追加する
 * @param {Rows|Columns} targetCollection 対象の行または列のコレクション
 * @param {number} addCount 追加する数
 * @returns {void}
 */
function appendItems(targetCollection, addCount) {
    for (var i = 0; i < addCount; i++) {
        targetCollection.add(LocationOptions.atEnd);
    }
}

/**
 * 行または列を末尾から取り除く
 * @param {Rows|Columns} targetCollection 対象の行または列のコレクション
 * @param {number} keepCount 残す数
 * @returns {void}
 */
function trimItems(targetCollection, keepCount) {
    while (targetCollection.length > keepCount) {
        try {
            targetCollection.lastItem().remove();
        } catch (e) {
            break;
        }
    }
}

/**
 * 転置しやすいよう、いったん表を正方形に揃える
 * @param {Table} targetTable 対象の表
 * @returns {{paddedAxis: string, originalAxisSize: number}} 足した軸（"columns" / "rows" / "none"）と、その軸の元のサイズ
 */
function padTableToSquare(targetTable) {
    var rowCount    = targetTable.rows.length;
    var columnCount = targetTable.columnCount;

    /* 少ない方の軸を多い方に合わせて増やす / Grow the shorter axis to match the longer one */
    if (rowCount > columnCount) {
        appendItems(targetTable.columns, rowCount - columnCount);
        return { paddedAxis: "columns", originalAxisSize: columnCount };
    }

    if (rowCount < columnCount) {
        appendItems(targetTable.rows, columnCount - rowCount);
        return { paddedAxis: "rows", originalAxisSize: rowCount };
    }

    return { paddedAxis: "none", originalAxisSize: rowCount };
}

/**
 * 空セルに代替文字を入れて段落を1つ確保する
 * @param {Table} targetTable 対象の表
 * @returns {void}
 */
function fillEmptyCells(targetTable) {
    var tableCells = targetTable.cells;
    for (var i = 0; i < tableCells.length; i++) {
        /* 内容を持てないセルは飛ばす / Skip cells that cannot take contents */
        try {
            if (tableCells[i].contents === "") tableCells[i].contents = EMPTY_CELL_PLACEHOLDER;
        } catch (e) {}
    }
}

/**
 * 正方形に揃えた表の上三角と下三角を入れ替えて転置する
 * @param {Table} targetTable 正方形に揃えた表
 * @returns {void}
 */
function transposeSquareTable(targetTable) {
    var rowCount    = targetTable.rows.length;
    var columnCount = targetTable.columnCount;

    for (var rowIndex = 0; rowIndex < rowCount; rowIndex++) {
        for (var columnIndex = rowIndex + 1; columnIndex < columnCount; columnIndex++) {
            var upperCellIndex = columnIndex + (rowIndex * columnCount);
            var lowerCellIndex = rowIndex + (columnIndex * columnCount);
            swapCells(targetTable.cells.item(upperCellIndex), targetTable.cells.item(lowerCellIndex));
        }
    }
}

/**
 * 正方形にするため増やした行・列を取り除く
 * @param {Table} targetTable 対象の表
 * @param {string} paddedAxis padTableToSquare が返した軸
 * @param {number} originalAxisSize 残すサイズ
 * @returns {void}
 */
function removePadding(targetTable, paddedAxis, originalAxisSize) {
    /* 列を足した表は転置後に行が余る（その逆も同じ）/ Padding columns leaves surplus rows after transposing, and vice versa */
    if (paddedAxis === "columns") {
        trimItems(targetTable.rows, originalAxisSize);
    } else if (paddedAxis === "rows") {
        trimItems(targetTable.columns, originalAxisSize);
    }
}

/**
 * ヘッダー／フッター行数を、行数に収まる範囲で設定する
 * 0 を渡すと指定を解除できる。
 * @param {Table} targetTable 対象の表
 * @param {number} headerRowCount 設定したいヘッダー行数
 * @param {number} footerRowCount 設定したいフッター行数
 * @returns {void}
 */
function setHeaderFooterRows(targetTable, headerRowCount, footerRowCount) {
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
 * @param {boolean} keepHeaderRows 転置後にヘッダー行の指定を引き継ぐか
 * @param {string} mergeMode MERGE_MODE_STOP または MERGE_MODE_UNMERGE
 * @returns {string} "ok" または "mergeStopped"
 */
function transposeTable(targetTable, keepHeaderRows, mergeMode) {
    /* 引き継がない場合はヘッダー指定を外す / Drop the header designation when it is not kept */
    var originalHeaderRowCount = keepHeaderRows ? targetTable.headerRowCount : 0;
    var originalFooterRowCount = targetTable.footerRowCount;

    if (hasMergedCells(targetTable)) {
        if (mergeMode === MERGE_MODE_STOP) return "mergeStopped";
        try {
            targetTable.unmerge();
        } catch (e) {
            /* 解除できなくても単純な表なら続行できる / Simple tables can continue even if unmerge fails */
        }
    }

    /* ヘッダー／フッターの指定をいったん外す。付いたままだと、末尾に足した行が
       フッターの手前に入って cells のインデックス計算がずれる
       / Drop the header and footer designation first: otherwise an appended row can land
       before the footer rows and throw off the cell index math */
    setHeaderFooterRows(targetTable, 0, 0);

    var squarePadding = padTableToSquare(targetTable);
    fillEmptyCells(targetTable);
    transposeSquareTable(targetTable);
    removePadding(targetTable, squarePadding.paddedAxis, squarePadding.originalAxisSize);

    /* 元の設定に近い形で戻す / Restore the original designation as closely as possible */
    setHeaderFooterRows(targetTable, originalHeaderRowCount, originalFooterRowCount);

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

var tableHasMergedCells = hasMergedCells(targetTable);
var tableHasHeaderRows  = (targetTable.headerRowCount > 0);

var keepHeaderRows = false;
var mergeMode      = MERGE_MODE_UNMERGE;

/* 選択の余地がない場合はダイアログを省略 / Skip the dialog when there is nothing to choose */
if (tableHasHeaderRows || tableHasMergedCells) {
    var dialogSettings = showTransposeDialog(tableHasHeaderRows, tableHasMergedCells);
    if (dialogSettings === null) return;
    keepHeaderRows = dialogSettings.keepHeaderRows;
    mergeMode      = dialogSettings.mergeMode;
}

/* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
var transposeStatus = app.doScript(function () {
    return transposeTable(targetTable, keepHeaderRows, mergeMode);
}, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.transposeTable"));

if (transposeStatus === "mergeStopped") {
    alert(getLabel("alert.mergeStopped"));
}

})();
