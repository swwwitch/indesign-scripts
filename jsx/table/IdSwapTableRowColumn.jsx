#target indesign

/*

### 概要

選択した表の行と列を入れ替えます。ヘッダー行の扱いとセル結合の処理方法をダイアログで選べます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSwapTableRowColumn.md

### Overview

Transposes the rows and columns of the selected table. The dialog picks how header rows and merged cells are handled.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSwapTableRowColumn.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdSwapTableRowColumn";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-11-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSwapTableRowColumn.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSwapTableRowColumn.md"; /* README (English) */

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

// =========================================
// 表の取得 / Table lookup
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
 * 選択とその祖先が持っている表を探す（表の外のカーソルやテキストフレームの選択向け）
 * @param {object} selectionItem 選択オブジェクト
 * @returns {Table|null} 最初に見つかった表。見つからない場合は null
 */
function findContainedTable(selectionItem) {
    var candidate = selectionItem;
    for (var i = 0; i < TABLE_PARENT_LOOKUP_LIMIT; i++) {
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
    return getTableFromSelection(selectionItem) || findContainedTable(selectionItem);
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

    var buttonRow = addButtonRow(transposeDialog);
    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    centerButtonRowIfRightOnly(buttonRow);

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
