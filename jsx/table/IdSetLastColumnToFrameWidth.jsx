#target indesign

/*

### 概要

カーソルのある表の最終列幅を調整し、表全体の幅を親テキストフレームの幅に合わせます。

詳細は README を参照してください。

### Overview

Adjusts the last column of the table at the cursor so that the whole table matches the width of its parent text frame.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdSetLastColumnToFrameWidth";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSetLastColumnToFrameWidth.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSetLastColumnToFrameWidth.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

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
        alert: {
            placeCursorInTable: { ja: "表内にカーソルを置いてください。", en: "Place the cursor inside a table." }
        },
        undo: {
            fitLastColumn: { ja: "最終列をフレーム幅に合わせる", en: "Fit Last Column to Frame Width" }
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
    // 表の取得 / Table lookup
    // =========================================

    /**
     * カーソル位置から対象の表を特定する
     * @param {object} cursorItem 選択オブジェクト（挿入ポイントまたはテキスト）
     * @returns {Table|null} 対象の表。特定できない場合は null
     */
    function resolveTableFromCursor(cursorItem) {
        if (cursorItem.constructor.name === "InsertionPoint") {
            if (cursorItem.parent.constructor.name === "Cell") return cursorItem.parent.parent;
            return null;
        }

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

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 最終列の幅を調整し、表の幅を親テキストフレームに合わせる
     * @returns {void}
     */
    function main() {
        var selectionItems = app.selection;

        if (selectionItems.length === 0 ||
            !selectionItems[0].hasOwnProperty("baseline") ||
            !(selectionItems[0].constructor.name === "InsertionPoint" || selectionItems[0].constructor.name === "Text")) {
            alert(localize(LABELS.alert.placeCursorInTable));
            return;
        }

        var targetTable = resolveTableFromCursor(selectionItems[0]);
        if (!targetTable) {
            alert(localize(LABELS.alert.placeCursorInTable));
            return;
        }

        /* 表を含むテキストフレームの幅 / Width of the text frame that holds the table */
        var frameBounds = targetTable.parent.geometricBounds;
        var textFrameWidth = frameBounds[3] - frameBounds[1];

        /* 最終列を除いた列幅の合計 / Total width of every column except the last one */
        var otherColumnsWidth = 0;
        for (var i = 0; i < targetTable.columns.length - 1; i++) {
            otherColumnsWidth += targetTable.columns[i].width;
        }

        targetTable.columns[targetTable.columns.length - 1].width = textFrameWidth - otherColumnsWidth;
    }

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, localize(LABELS.undo.fitLastColumn));

})();
