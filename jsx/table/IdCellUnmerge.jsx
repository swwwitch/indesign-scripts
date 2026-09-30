#target indesign

/*

### 概要

表の結合セルを解除します。解除後のセルへ元のテキストを複製するかどうかと、対象範囲（表全体・選択セルのみ）をダイアログで選べます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCellUnmerge.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n175525637a3d

### Overview

Unmerges merged cells in a table. The dialog picks whether to copy the original text into the resulting cells and the scope (the whole table or only the selected cells).

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCellUnmerge.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdCellUnmerge";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCellUnmerge.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCellUnmerge.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n175525637a3d"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// ==============================
// UIレイアウトの共通設定 / Shared UI layout
// ==============================

// UIレイアウト（再利用パーツ） / UI layout (reusable)

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

/**
 * ウィンドウの共通設定
 * @param {Window} targetWindow - 対象のウィンドウ
 * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
 * @returns {void}
 */
function setupWindow(targetWindow, spacing) {
    targetWindow.orientation = "column";
    targetWindow.alignChildren = "fill";
    targetWindow.margins = WINDOW_MARGINS;
    targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
 * @param {Panel} targetPanel - 対象のパネル
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
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
 * タブの共通設定
 * @param {Tab} targetTab - 対象のタブ
 * @param {number} [spacing] - 要素間隔（省略時は変えない）
 * @returns {void}
 */
function setupTab(targetTab, spacing) {
    targetTab.orientation = "column";
    targetTab.alignChildren = "fill";
    targetTab.margins = TAB_MARGINS;
    if (typeof spacing === "number") targetTab.spacing = spacing;
}

/**
 * 横並びの行グループの共通設定（ボタン列など）。
 * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
 * @param {Group} rowGroup - 対象のグループ
 * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
 * @returns {void}
 */
function setupRow(rowGroup, rowAlignment, spacing) {
    rowGroup.orientation = "row";
    rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
    rowGroup.alignChildren = ["left", "center"];
    rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
 * @param {Button} targetButton - 対象のボタン
 * @param {number} trimPixels - 詰める量（px）
 * @returns {void}
 */
function trimButtonHeight(targetButton, trimPixels) {
    /* レイアウト前は size が無い / size is not set until the layout runs */
    if (!targetButton.size) return;
    targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
}

// UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

// ボタン行（再利用パーツ） / Button row (reusable)

var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

/**
 * ダイアログ下部のボタン行を作る。
 * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
 * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
 * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
 * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
 */
function addButtonRow(parent, rowOptions) {
    var isCentered = !!(rowOptions && rowOptions.centered);
    var btnRowGroup = parent.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
    btnRowGroup.spacing = BUTTON_ROW_SPACING;

    if (isCentered) {
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
    }

    btnRowGroup.alignment = ["fill", "bottom"];

    var btnLeftGroup = btnRowGroup.add("group");
    btnLeftGroup.alignChildren = ["left", "center"];
    btnLeftGroup.spacing = BUTTON_ROW_SPACING;

    /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
    var spacer = btnRowGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 0;

    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.alignChildren = ["right", "center"];
    btnRightGroup.spacing = BUTTON_ROW_SPACING;

    return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
}

/**
 * 左のグループにボタンが無い（右のボタンだけの）とき、行を左右中央に並べ直す。
 * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
 * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
 * @returns {void}
 */
function centerButtonRowIfRightOnly(buttonRow) {
    if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
    var btnRowGroup = buttonRow.rowGroup;
    /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
    btnRowGroup.remove(buttonRow.leftGroup);
    btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
    btnRowGroup.alignment = ["center", "bottom"];
    btnRowGroup.alignChildren = ["center", "center"];
    buttonRow.leftGroup = null;
}

// ボタン行（再利用パーツ）ここまで / End of the reusable button row

(function () {

    // =========================================
    // ラベル定義 / Labels
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

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

    // =========================================
    // 表とセルの取得 / Table and cell lookup
    // =========================================

    // 表の選択（再利用パーツ） / Table selection (reusable)

    var TABLE_PARENT_LOOKUP_LIMIT = 20; /* 親をたどる上限（無限ループよけ） / max parent hops (guards against loops) */

    /**
     * 親をたどって、いちばん近い表を返す（セル・行・列・セル内のテキストや挿入点に対応）
     * @param {Object} startItem - たどり始めるオブジェクト
     * @returns {Table|null} 見つかった表。表の中でなければ null
     */
    function findParentTable(startItem) {
        var node = startItem;
        for (var i = 0; i < TABLE_PARENT_LOOKUP_LIMIT && node; i++) {
            try {
                var typeName = node.constructor.name;
                if (typeName === "Table") return node;
                if (typeName === "Document" || typeName === "Application") return null;
                node = node.parent;
            } catch (e) {
                return null;
            }
        }
        return null;
    }

    /**
     * 選択から対象の表を返す。表の中の選択ならその表、表を含むテキストフレームやテキストなら最初の表
     * @param {Object} selectionItem - app.selection[0] など
     * @returns {Table|null} 対象の表。見つからなければ null
     */
    function getTableFromSelection(selectionItem) {
        if (!selectionItem) return null;
        var parentTable = findParentTable(selectionItem);
        if (parentTable) return parentTable;
        try {
            if (selectionItem.tables && selectionItem.tables.length > 0) return selectionItem.tables[0];
        } catch (e) {
            /* tables を持たない選択（画像など） / Selections without tables (images, etc.) */
        }
        return null;
    }

    /**
     * 選択からセルを1つずつの配列にして返す。セル選択・表の選択はその全セル、セル内のテキストや挿入点はそのセル
     * @param {Object} selectionItem - app.selection[0] など
     * @returns {Cell[]} セルの配列。表の外なら空配列
     */
    function getSelectedCells(selectionItem) {
        var selectedCells = [];
        if (!selectionItem) return selectedCells;
        try {
            var typeName = selectionItem.constructor.name;
            if (typeName === "Cell" || typeName === "Table") {
                /* 複数セルの選択も1つの Cell で返るので .cells で展開する / A multi-cell selection is one Cell; expand it via .cells */
                var cellCollection = selectionItem.cells;
                for (var i = 0; i < cellCollection.length; i++) selectedCells.push(cellCollection[i]);
                return selectedCells;
            }
            var node = selectionItem;
            for (var j = 0; j < TABLE_PARENT_LOOKUP_LIMIT && node; j++) {
                var nodeType = node.constructor.name;
                if (nodeType === "Cell") {
                    selectedCells.push(node);
                    break;
                }
                if (nodeType === "Table" || nodeType === "Document" || nodeType === "Application") break;
                node = node.parent;
            }
        } catch (e) {
            /* 親をたどれない選択は空のまま / Leave empty when the parent chain cannot be followed */
        }
        return selectedCells;
    }

    /**
     * 2つの表が同じ表かどうかを返す
     * @param {Table} tableA - 表
     * @param {Table} tableB - 表
     * @returns {boolean} 同じ表なら true
     */
    function isSameTable(tableA, tableB) {
        if (!tableA || !tableB) return false;
        try {
            return tableA.id === tableB.id && tableA.parent.id === tableB.parent.id;
        } catch (e) {
            return false;
        }
    }

    // 表の選択（再利用パーツ）ここまで / End of the reusable table selection

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
    function collectTargetTableCells(selectionItems, targetTable) {
        var collectedCells = [];

        for (var i = 0; i < selectionItems.length; i++) {
            var itemCells = getSelectedCells(selectionItems[i]);
            for (var j = 0; j < itemCells.length; j++) {
                if (isSameTable(findParentTable(itemCells[j]), targetTable)) collectedCells.push(itemCells[j]);
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
            if (typeName === "Cell") {
                hasCellSelection = true;
                break;
            }
        }
        if (!hasCellSelection) return false;

        var selectedCells = dedupeCells(collectTargetTableCells(selectionItems, targetTable));
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

        var buttonRow = addButtonRow(unmergeDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        centerButtonRowIfRightOnly(buttonRow);

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
            candidateCells = collectTargetTableCells(selectionItems, targetTable);
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
