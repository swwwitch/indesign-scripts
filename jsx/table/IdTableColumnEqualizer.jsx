#target indesign

/*

### 概要

カーソルのある表の幅を、親テキストフレームの幅にそろえます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableColumnEqualizer.md

### Overview

Fits the width of the table at the cursor to the width of its parent text frame.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableColumnEqualizer.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTableColumnEqualizer";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableColumnEqualizer.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableColumnEqualizer.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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
        alert: {
            placeCursorInTable: { ja: "表内にカーソルを置いてください。", en: "Place the cursor inside a table." },
            frameNotFound:      { ja: "表がテキストフレーム内に見つかりません。", en: "The table is not inside a text frame." }
        },
        undo: {
            fitTableWidth: { ja: "表の幅をフレーム幅に合わせる", en: "Fit Table Width to Frame" }
        }
    };

    // =========================================
    // 表とフレームの取得 / Table and frame lookup
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
     * カーソル位置から対象の表を特定する
     * @param {object} cursorItem 選択オブジェクト（挿入ポイントまたはテキスト）
     * @returns {Table|null} 対象の表。特定できない場合は null
     */
    function resolveTableFromCursor(cursorItem) {
        /* セル内のカーソルやテキストは、親をたどってその表を取る / Inside a cell, walk up to its table */
        var parentTable = findParentTable(cursorItem);
        if (parentTable) return parentTable;

        /* 本文で表のアンカーにかかるテキストを選んでいるときは、その位置の表を取る
           / For text in the story that spans a table anchor, take the table at that position */
        if (cursorItem.constructor.name !== "Text") return null;
        if (cursorItem.parentTextFrames.length === 0) return null;

        var parentFrame = cursorItem.parentTextFrames[0];
        for (var i = 0; i < parentFrame.tables.length; i++) {
            var candidateTable = parentFrame.tables[i];
            var tableStart = candidateTable.storyOffset.index;
            var tableEnd   = tableStart + candidateTable.characters.length;
            if (tableStart <= cursorItem.index && cursorItem.index <= tableEnd) return candidateTable;
        }
        return null;
    }

    /**
     * 表を内包しているテキストフレームを親方向にたどって探す
     * @param {Table} targetTable 対象の表
     * @returns {TextFrame|null} テキストフレーム。見つからない場合は null
     */
    function findEnclosingTextFrame(targetTable) {
        var container = targetTable.parent;
        while (container.constructor.name !== "TextFrame" && container.constructor.name !== "Story") {
            container = container.parent;
        }
        return (container.constructor.name === "TextFrame") ? container : null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 表の幅を親テキストフレームの幅にそろえる
     * @returns {void}
     */
    function main() {
        var selectionItems = app.selection;

        if (selectionItems.length === 0 || !selectionItems[0].hasOwnProperty("baseline")) {
            alert(getLabel(LABELS.alert.placeCursorInTable));
            return;
        }

        var targetTable = resolveTableFromCursor(selectionItems[0]);
        if (!targetTable) {
            alert(getLabel(LABELS.alert.placeCursorInTable));
            return;
        }

        var enclosingFrame = findEnclosingTextFrame(targetTable);
        if (!enclosingFrame) {
            alert(getLabel(LABELS.alert.frameNotFound));
            return;
        }

        var frameBounds = enclosingFrame.geometricBounds;
        targetTable.width = frameBounds[3] - frameBounds[1];
    }

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undo.fitTableWidth));

})();
