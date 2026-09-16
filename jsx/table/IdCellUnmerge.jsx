#target indesign

/*

### 概要

表の結合セルを解除します。解除後のセルへ元のテキストを複製するかどうかと、対象範囲（表全体・選択セルのみ）をダイアログで選べます。

詳細は README を参照してください。

### Overview

Unmerges merged cells in a table. The dialog picks whether to copy the original text into the resulting cells and the scope (the whole table or only the selected cells).

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdCellUnmerge";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-16";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCellUnmerge.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCellUnmerge.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n175525637a3d"; /* 紹介記事 / article URL */

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
            title: { ja: "セルの結合を解除", en: "Unmerge Cells" }
        },
        panel: {
            text:  { ja: "解除後のテキスト", en: "Text after unmerge" },
            scope: { ja: "対象", en: "Scope" }
        },
        radio: {
            keepInOriginal: { ja: "元のセルにのみ残す", en: "Keep in the original cell" },
            copyToAll:      { ja: "すべてのセルに複製", en: "Copy to all cells" },
            wholeTable:     { ja: "表全体", en: "Whole table" },
            selectedCells:  { ja: "選択セルのみ", en: "Selected cells only" }
        },
        tooltip: {
            text: {
                ja: "「すべてのセルに複製」を選ぶと、その結合セルを解除してできたセルすべてに同じテキストが入ります。隣接する別のセルには影響しません。書式は引き継がれず、プレーンテキストになります。",
                en: "Copy to all cells puts the same text into every cell the unmerge created. Neighboring cells are left alone, and the text is inserted as plain text without its formatting."
            },
            scope: {
                ja: "「選択セルのみ」は、選択範囲に含まれる結合セルだけを解除します。",
                en: "Selected cells only unmerges the merged cells inside the selection."
            }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:    { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectTable:   { ja: "表、または表の中のセルを選択してください。", en: "Select a table or cells inside a table." },
            noTargetCells: { ja: "対象となるセルが選択されていません。", en: "No target cells are selected." }
        },
        undo: {
            unmergeCells: { ja: "セルの結合を解除", en: "Unmerge Cells" }
        }
    };

    /**
     * ラベルを現在の言語で取得する
     * @param {object} labelSet ja / en を持つラベルオブジェクト
     * @returns {string} 現在の言語のラベル文字列
     */
    function getLabel(labelSet) {
        return labelSet[currentLang];
    }

    // =========================================
    // 表とセルの取得 / Table and cell lookup
    // =========================================

    /**
     * 任意のオブジェクトから親方向にたどって表を探す
     * @param {object} startObject 起点となるオブジェクト
     * @returns {Table|null} 表。見つからない場合は null
     */
    function getParentTable(startObject) {
        var ancestor = startObject;
        while (ancestor) {
            if (ancestor.constructor && ancestor.constructor.name === "Table") return ancestor;
            if (!ancestor.parent || ancestor.parent === ancestor) break;
            ancestor = ancestor.parent;
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
     * 2 つの DOM オブジェクトが同じ表かどうかを id で判定する
     * @param {Table|null} tableA 比較する表
     * @param {Table|null} tableB 比較する表
     * @returns {boolean} 同じ表なら true
     */
    function isSameTable(tableA, tableB) {
        if (!tableA || !tableB) return false;
        try {
            return tableA.id === tableB.id;
        } catch (e) {
            return false;
        }
    }

    /**
     * 選択オブジェクトを個々のセルへ解決する
     * 複数セルの選択は 1 個の Cell として返り、getElements() では展開されないため、
     * cells コレクションを使って実体のセルへ展開する
     * @param {object} selectionItem 選択オブジェクト
     * @returns {Array<Cell>} 実体のセル配列
     */
    function resolveCellElements(selectionItem) {
        try {
            var cellElements = selectionItem.cells.everyItem().getElements();
            if (cellElements && cellElements.length > 1) return cellElements;
        } catch (e) {}
        return [selectionItem];
    }

    /**
     * セルの表示テキストを文字列で取得する
     * 結合セルの contents は構成セルごとのテキストを配列で返すため、そのままでは使えない
     * @param {Cell} cell 対象のセル
     * @returns {string} セルのテキスト
     */
    function getCellText(cell) {
        try {
            /* Text の contents は常に文字列 / Text.contents is always a string */
            return String(cell.texts[0].contents);
        } catch (e) {}

        var cellContents = cell.contents;
        if (cellContents instanceof Array) {
            for (var i = 0; i < cellContents.length; i++) {
                if (cellContents[i] !== "") return String(cellContents[i]);
            }
            return "";
        }

        return String(cellContents);
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

            if (typeName === "Cell" || typeName === "Cells") {
                var cellElements = resolveCellElements(selectionItem);
                for (var j = 0; j < cellElements.length; j++) {
                    if (isSameTable(getParentTable(cellElements[j]), targetTable)) collectedCells.push(cellElements[j]);
                }
                continue;
            }

            /* セル内のテキストが選択されているケース / The selection is text inside a cell */
            var ancestor = selectionItem;
            while (ancestor && ancestor !== ancestor.parent) {
                if (ancestor.constructor && ancestor.constructor.name === "Cell") {
                    if (isSameTable(getParentTable(ancestor), targetTable)) collectedCells.push(ancestor);
                    break;
                }
                ancestor = ancestor.parent;
            }
        }

        return collectedCells;
    }

    /**
     * セル配列から重複を取り除く（id ベース）
     * @param {Array<Cell>} cells セルの配列
     * @returns {Array<Cell>} 重複を除いたセルの配列
     */
    function dedupeCells(cells) {
        var dedupedCells = [];
        var seenCellIds = {};

        for (var i = 0; i < cells.length; i++) {
            var cell = cells[i];
            if (!cell || !cell.isValid) continue;

            var cellId;
            try {
                cellId = cell.id;
            } catch (e) {
                continue;
            }

            if (seenCellIds[cellId]) continue;
            seenCellIds[cellId] = true;
            dedupedCells.push(cell);
        }

        return dedupedCells;
    }

    /**
     * 表の一部だけがセル選択されているかを判定する
     * テキストカーソルを置いただけの状態は「一部の選択」とみなさない
     * @param {Array} selectionItems 選択オブジェクトの配列
     * @param {Table} targetTable 対象の表
     * @returns {boolean} 表の一部が選択されていれば true
     */
    function isPartialCellSelection(selectionItems, targetTable) {
        if (!targetTable) return false;

        /* セルを選んだときだけ対象を絞る。テキスト選択は表全体のつもりで実行していることが多い
           / Narrow the scope only for a cell selection; a text selection usually means the whole table */
        var hasCellSelection = false;
        for (var i = 0; i < selectionItems.length; i++) {
            var typeName = selectionItems[i].constructor && selectionItems[i].constructor.name;
            if (typeName === "Cell" || typeName === "Cells") {
                hasCellSelection = true;
                break;
            }
        }
        if (!hasCellSelection) return false;

        var selectedCells = dedupeCells(getSelectedCells(selectionItems, targetTable));
        if (selectedCells.length === 0) return false;

        try {
            return selectedCells.length < targetTable.cells.length;
        } catch (e) {
            return false;
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
     * 結合解除の設定ダイアログを表示する
     * @param {boolean} selectedCellsByDefault ［対象］の初期値を「選択セルのみ」にするか
     * @returns {{copyToAllCells: boolean, wholeTable: boolean}|null} 設定内容。キャンセル時は null
     */
    function showUnmergeDialog(selectedCellsByDefault) {
        var unmergeDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(unmergeDialog, 10);

        var textRadios = addRadioPanel(unmergeDialog, LABELS.panel.text,
            [LABELS.radio.keepInOriginal, LABELS.radio.copyToAll], 1, LABELS.tooltip.text);
        var scopeRadios = addRadioPanel(unmergeDialog, LABELS.panel.scope,
            [LABELS.radio.wholeTable, LABELS.radio.selectedCells],
            selectedCellsByDefault ? 1 : 0, LABELS.tooltip.scope);

        /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
        var btnRowGroup = unmergeDialog.add("group");
        setupRow(btnRowGroup, "right", 8);
        btnRowGroup.margins = [0, 10, 0, 0];
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        if (unmergeDialog.show() !== 1) return null;

        return {
            copyToAllCells: textRadios[1].value,
            wholeTable: scopeRadios[0].value
        };
    }

    // =========================================
    // 結合解除 / Unmerge
    // =========================================

    /**
     * 結合セルを解除し、必要に応じて元のテキストを複製する
     * @param {boolean} copyToAllCells 解除後のすべてのセルに元のテキストを複製するか
     * @param {boolean} wholeTable 表全体を対象にするか（false なら選択セルのみ）
     * @returns {void}
     */
    function runUnmerge(copyToAllCells, wholeTable) {
        var selectionItems = app.selection;
        if (!selectionItems || selectionItems.length === 0) {
            alert(getLabel(LABELS.alert.selectTable));
            return;
        }

        var targetTable = getTableFromSelection(selectionItems[0]);
        if (!targetTable) {
            alert(getLabel(LABELS.alert.selectTable));
            return;
        }

        var candidateCells;
        if (wholeTable) {
            candidateCells = targetTable.cells.everyItem().getElements();
        } else {
            candidateCells = getSelectedCells(selectionItems, targetTable);
            if (!candidateCells || candidateCells.length === 0) {
                alert(getLabel(LABELS.alert.noTargetCells));
                return;
            }
        }

        candidateCells = dedupeCells(candidateCells);

        /* 解除するたびに表のセル数が増えてインデックスがずれるため、後ろから処理する
           / Each unmerge adds cells and shifts the indexes after it, so walk the list backwards */
        for (var i = candidateCells.length - 1; i >= 0; i--) {
            var cell = candidateCells[i];
            if (!cell || !cell.isValid) continue;

            /* 行方向または列方向に 2 つ以上を跨いでいれば結合セル / A cell spanning more than one row or column is merged */
            if (cell.rowSpan <= 1 && cell.columnSpan <= 1) continue;

            var mergedCellText = getCellText(cell);

            /* 解除で増えたセルを後から特定するため、解除前のセル id を控える
               / Record the cell ids before unmerging so the cells it creates can be identified afterwards */
            var cellIdsBeforeUnmerge = collectCellIds(targetTable);

            try {
                cell.unmerge();
            } catch (e) {
                continue;
            }

            if (!copyToAllCells || mergedCellText === "") continue;

            /* 解除で増えたセルだけがこの結合セルの内側 / Only the cells the unmerge created belong to this merged cell */
            copyTextToCells(getCellsAddedSince(targetTable, cellIdsBeforeUnmerge), mergedCellText);
        }
    }

    /**
     * 表に含まれるセルの id を集める
     * @param {Table} targetTable 対象の表
     * @returns {object} id をキーにしたルックアップ
     */
    function collectCellIds(targetTable) {
        var cellIds = {};
        var allCells = targetTable.cells.everyItem().getElements();

        for (var i = 0; i < allCells.length; i++) {
            try {
                cellIds[allCells[i].id] = true;
            } catch (e) {}
        }

        return cellIds;
    }

    /**
     * 控えた id に含まれない（＝あとから増えた）セルを集める
     * @param {Table} targetTable 対象の表
     * @param {object} previousCellIds 解除前のセル id のルックアップ
     * @returns {Array<Cell>} 増えたセルの配列
     */
    function getCellsAddedSince(targetTable, previousCellIds) {
        var addedCells = [];
        var allCells = targetTable.cells.everyItem().getElements();

        for (var i = 0; i < allCells.length; i++) {
            var cell = allCells[i];
            if (!cell || !cell.isValid) continue;
            try {
                if (!previousCellIds[cell.id]) addedCells.push(cell);
            } catch (e) {}
        }

        return addedCells;
    }

    /**
     * 解除で増えたセルへ元のテキストを複製する
     * @param {Array<Cell>} createdCells 解除で増えたセルの配列
     * @param {string} mergedCellText 元のテキスト
     * @returns {void}
     */
    function copyTextToCells(createdCells, mergedCellText) {
        for (var i = 0; i < createdCells.length; i++) {
            var targetCell = createdCells[i];
            if (!targetCell || !targetCell.isValid) continue;

            try {
                targetCell.contents = mergedCellText;
            } catch (e) {
                /* 1 セルの失敗で全体を止めない / One failed cell must not abort the run */
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
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var selectionItems = app.selection;
        if (!selectionItems || selectionItems.length === 0) {
            alert(getLabel(LABELS.alert.selectTable));
            return;
        }

        var targetTable = getTableFromSelection(selectionItems[0]);
        if (!targetTable) {
            alert(getLabel(LABELS.alert.selectTable));
            return;
        }

        /* 表の一部を選んで実行したときは、そのまま［選択セルのみ］で始められるようにする
           / Start on Selected cells only when the run began from a partial selection */
        var unmergeSettings = showUnmergeDialog(isPartialCellSelection(selectionItems, targetTable));
        if (unmergeSettings === null) return;

        /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
        app.doScript(function () {
            runUnmerge(unmergeSettings.copyToAllCells, unmergeSettings.wholeTable);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undo.unmergeCells));
    }

    main();

})();
