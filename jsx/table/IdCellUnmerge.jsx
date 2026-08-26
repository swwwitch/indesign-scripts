#target indesign

/*
 * IdCellUnmerge.jsx
 *
 * 表の結合セルを解除します。解除後のセルへ元のテキストを複製するかどうかと、対象範囲をダイアログで選べます。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdCellUnmerge";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCellUnmerge.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCellUnmerge.md

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
        title: { ja: "セルの結合解除", en: "Unmerge Cells" }
    },
    panel: {
        merge: { ja: "結合", en: "Merge" },
        scope: { ja: "対象", en: "Scope" }
    },
    radio: {
        noDistribute:  { ja: "デフォルト（分配なし）", en: "Default (no distribution)" },
        distribute:    { ja: "テキストを分配", en: "Distribute text" },
        wholeTable:    { ja: "表全体", en: "Whole table" },
        selectedCells: { ja: "選択したセルのみ", en: "Selected cells only" }
    },
    button: {
        ok:     { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    alert: {
        noDocument:    { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        selectCells:   { ja: "表のセルを選択してください。", en: "Please select table cells." },
        noTargetCells: { ja: "対象となるセルが選択されていません。", en: "No target cells are selected." }
    },
    undo: {
        unmergeCells: { ja: "セルの結合解除", en: "Unmerge Cells" }
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
// 表とセルの取得 / Table and cell lookup
// =========================================

/**
 * 任意のオブジェクトから親方向にたどって表を探す
 * @param {object} pageObject 起点となるオブジェクト
 * @returns {Table|null} 表。見つからない場合は null
 */
function getParentTable(pageObject) {
    var currentObject = pageObject;
    while (currentObject) {
        if (currentObject.constructor && currentObject.constructor.name === "Table") return currentObject;
        if (!currentObject.parent || currentObject.parent === currentObject) break;
        currentObject = currentObject.parent;
    }
    return null;
}

/**
 * 選択オブジェクトから対象の表を取得する
 * @param {object} selectionItem 選択オブジェクト
 * @returns {Table|null} 対象の表。特定できない場合は null
 */
function getTableFromSelection(selectionItem) {
    if (!selectionItem) return null;

    var typeName = selectionItem.constructor && selectionItem.constructor.name;
    if (typeName === "Table") return selectionItem;
    if (typeName === "Cell") return getParentTable(selectionItem);
    if (typeName === "Cells") return (selectionItem.length > 0) ? getParentTable(selectionItem[0]) : null;

    return getParentTable(selectionItem);
}

/**
 * 選択範囲から対象の表に属するセルを集める
 * @param {Array} selectionItems 選択オブジェクトの配列
 * @param {Table} targetTable 対象の表
 * @returns {Array<Cell>} 対象セルの配列
 */
function getSelectedCells(selectionItems, targetTable) {
    var collectedCells = [];

    for (var i = 0; i < selectionItems.length; i++) {
        var selectionItem = selectionItems[i];
        var typeName = selectionItem.constructor && selectionItem.constructor.name;

        if (typeName === "Cell") {
            if (getParentTable(selectionItem) === targetTable) collectedCells.push(selectionItem);
            continue;
        }

        if (typeName === "Cells") {
            var cellElements = selectionItem.getElements();
            for (var j = 0; j < cellElements.length; j++) {
                if (getParentTable(cellElements[j]) === targetTable) collectedCells.push(cellElements[j]);
            }
            continue;
        }

        /* セル内のテキストが選択されているケース / The selection is text inside a cell */
        var currentObject = selectionItem;
        while (currentObject && currentObject !== currentObject.parent) {
            if (currentObject.constructor && currentObject.constructor.name === "Cell") {
                if (getParentTable(currentObject) === targetTable) collectedCells.push(currentObject);
                break;
            }
            currentObject = currentObject.parent;
        }
    }

    return collectedCells;
}

/**
 * セル配列から重複を取り除く（id ベース）
 * @param {Array<Cell>} cells セルの配列
 * @returns {Array<Cell>} 重複を除いたセルの配列
 */
function uniqueCells(cells) {
    var uniqueList = [];
    var seenCellIds = {};

    for (var i = 0; i < cells.length; i++) {
        var cell = cells[i];
        if (!cell || !cell.isValid) continue;
        if (seenCellIds[cell.id]) continue;
        seenCellIds[cell.id] = true;
        uniqueList.push(cell);
    }

    return uniqueList;
}

// =========================================
// ダイアログ / Dialog
// =========================================

/**
 * 結合解除の設定ダイアログを表示する
 * @returns {{distributeText: boolean, wholeTable: boolean}|null} 設定内容。キャンセル時は null
 */
function showUnmergeDialog() {
    var unmergeDialog = new Window("dialog", localize(LABELS.dialog.title) + " " + SCRIPT_VERSION);
    setupWindow(unmergeDialog, 10);

    /* 結合パネル / Merge panel */
    var mergePanel = unmergeDialog.add("panel", undefined, localize(LABELS.panel.merge));
    setupPanel(mergePanel, 6);
    mergePanel.alignChildren = ["left", "top"];

    mergePanel.add("radiobutton", undefined, localize(LABELS.radio.noDistribute));
    var distributeTextRadio = mergePanel.add("radiobutton", undefined, localize(LABELS.radio.distribute));
    distributeTextRadio.value = true;

    /* 対象パネル / Scope panel */
    var scopePanel = unmergeDialog.add("panel", undefined, localize(LABELS.panel.scope));
    setupPanel(scopePanel, 6);
    scopePanel.alignChildren = ["left", "top"];

    var wholeTableRadio = scopePanel.add("radiobutton", undefined, localize(LABELS.radio.wholeTable));
    wholeTableRadio.value = true;
    scopePanel.add("radiobutton", undefined, localize(LABELS.radio.selectedCells));

    /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
    var dialogButtonRow = unmergeDialog.add("group");
    setupRow(dialogButtonRow, "center", 8);
    dialogButtonRow.margins = [0, 10, 0, 0];
    dialogButtonRow.add("button", undefined, localize(LABELS.button.cancel), { name: "cancel" });
    dialogButtonRow.add("button", undefined, localize(LABELS.button.ok), { name: "ok" });

    if (unmergeDialog.show() !== 1) return null;

    return {
        distributeText: distributeTextRadio.value,
        wholeTable: wholeTableRadio.value
    };
}

// =========================================
// 結合解除 / Unmerge
// =========================================

/**
 * 結合セルを解除し、必要に応じて元のテキストを複製する
 * @param {boolean} distributeText 解除後のすべてのセルに元のテキストを複製するか
 * @param {boolean} wholeTable 表全体を対象にするか（false なら選択セルのみ）
 * @returns {void}
 */
function runUnmerge(distributeText, wholeTable) {
    var selectionItems = app.selection;
    if (!selectionItems || selectionItems.length === 0) {
        alert(localize(LABELS.alert.selectCells));
        return;
    }

    var targetTable = getTableFromSelection(selectionItems[0]);
    if (!targetTable) {
        alert(localize(LABELS.alert.selectCells));
        return;
    }

    var cellsToCheck;
    if (wholeTable) {
        cellsToCheck = targetTable.cells.everyItem().getElements();
    } else {
        cellsToCheck = getSelectedCells(selectionItems, targetTable);
        if (!cellsToCheck || cellsToCheck.length === 0) {
            alert(localize(LABELS.alert.noTargetCells));
            return;
        }
    }

    cellsToCheck = uniqueCells(cellsToCheck);

    for (var i = 0; i < cellsToCheck.length; i++) {
        var cell = cellsToCheck[i];
        if (!cell || !cell.isValid) continue;

        /* 行方向または列方向に 2 つ以上を跨いでいれば結合セル / A cell spanning more than one row or column is merged */
        if (cell.rowSpan <= 1 && cell.columnSpan <= 1) continue;

        var originalContents = cell.contents;

        /* unmerge() は解除後のセル配列を返す / unmerge() returns the resulting cells */
        var unmergedCells;
        try {
            unmergedCells = cell.unmerge();
        } catch (e) {
            continue;
        }

        if (!distributeText || originalContents === "") continue;

        for (var j = 0; j < unmergedCells.length; j++) {
            if (!unmergedCells[j] || !unmergedCells[j].isValid) continue;
            try {
                unmergedCells[j].contents = originalContents;
            } catch (e) {
                /* 1 セルの失敗で全体を止めない / One failed cell must not abort the run */
            }
        }
    }
}

// =========================================
// メイン処理 / Main
// =========================================

/**
 * ダイアログを表示し、選んだ条件で結合セルを解除する
 * @returns {void}
 */
function main() {
    if (app.documents.length === 0) {
        alert(localize(LABELS.alert.noDocument));
        return;
    }
    if (!app.selection || app.selection.length === 0) {
        alert(localize(LABELS.alert.selectCells));
        return;
    }

    var dialogResult = showUnmergeDialog();
    if (dialogResult === null) return;

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(function () {
        runUnmerge(dialogResult.distributeText, dialogResult.wholeTable);
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, localize(LABELS.undo.unmergeCells));
}

main();
