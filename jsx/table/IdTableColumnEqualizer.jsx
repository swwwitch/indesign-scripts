#target indesign

/*
 * IdTableColumnEqualizer.jsx
 *
 * カーソルのある表の幅を、親テキストフレームの幅にそろえます。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTableColumnEqualizer";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableColumnEqualizer.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableColumnEqualizer.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

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
        placeCursorInTable: { ja: "表内にカーソルを置いてください。", en: "Place the cursor inside a table." },
        frameNotFound:      { ja: "表がテキストフレーム内に見つかりません。", en: "The table is not inside a text frame." }
    },
    undo: {
        fitTableWidth: { ja: "表の幅をフレーム幅に合わせる", en: "Fit Table Width to Frame" }
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
// 表とフレームの取得 / Table and frame lookup
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
        alert(localize(LABELS.alert.placeCursorInTable));
        return;
    }

    var targetTable = resolveTableFromCursor(selectionItems[0]);
    if (!targetTable) {
        alert(localize(LABELS.alert.placeCursorInTable));
        return;
    }

    var enclosingFrame = findEnclosingTextFrame(targetTable);
    if (!enclosingFrame) {
        alert(localize(LABELS.alert.frameNotFound));
        return;
    }

    var frameBounds = enclosingFrame.geometricBounds;
    targetTable.width = frameBounds[3] - frameBounds[1];
}

/* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, localize(LABELS.undo.fitTableWidth));
