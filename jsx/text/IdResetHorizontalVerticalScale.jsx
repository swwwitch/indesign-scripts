#target indesign

/*

### 概要

ドキュメント内すべてのストーリー（表セル・入れ子の表を含む）を走査し、文字の水平・垂直比率を 100% に戻します。

詳細は README を参照してください。

### Overview

Walks every story in the document, including table cells and nested tables, and resets the horizontal and vertical character scale to 100%.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdResetHorizontalVerticalScale"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-19";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdResetHorizontalVerticalScale.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdResetHorizontalVerticalScale.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 変倍なしとみなす比率（%）/ The scale value treated as "no scaling" (%) */
var NORMAL_SCALE_PERCENT = 100;

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
        noDocument:   { ja: "ドキュメントを開いてください。", en: "Please open a document." },
        doneTitle:    { ja: "完了しました。", en: "Done." },
        changedCount: { ja: "\n変更した文字範囲: ", en: "\nText ranges changed: " }
    },
    undo: {
        resetScale: { ja: "文字の水平・垂直比率を100%にする", en: "Reset Horizontal / Vertical Scale" }
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
// 前提チェック / Preconditions
// =========================================
if (app.documents.length === 0) {
    alert(localize(LABELS.alert.noDocument));
    return;
}

var targetDocument = app.activeDocument;

// =========================================
// 変倍のリセット / Scale reset
// =========================================

/**
 * 文字スタイル範囲のうち、比率が 100% でないものを 100% に戻す
 * @param {Array<TextStyleRange>} styleRanges 対象の文字スタイル範囲の配列
 * @returns {number} リセットした文字範囲の件数
 */
function resetScaleInRanges(styleRanges) {
    var resetCount = 0;

    for (var i = 0; i < styleRanges.length; i++) {
        var styleRange = styleRanges[i];
        if (styleRange.horizontalScale !== NORMAL_SCALE_PERCENT ||
            styleRange.verticalScale !== NORMAL_SCALE_PERCENT) {
            styleRange.horizontalScale = NORMAL_SCALE_PERCENT;
            styleRange.verticalScale = NORMAL_SCALE_PERCENT;
            resetCount++;
        }
    }

    return resetCount;
}

/**
 * テキストコンテナの比率を戻し、内包する表のセルへ再帰する
 * @param {Story|Text} textContainer textStyleRanges と tables を持つオブジェクト
 * @returns {number} リセットした文字範囲の件数
 */
function resetScaleInContainer(textContainer) {
    var resetCount = resetScaleInRanges(textContainer.textStyleRanges.everyItem().getElements());

    var tables = textContainer.tables.everyItem().getElements();
    for (var i = 0; i < tables.length; i++) {
        var cells = tables[i].cells.everyItem().getElements();
        for (var j = 0; j < cells.length; j++) {
            /* セル内テキストを同じ処理で再帰（入れ子の表にも対応）/ Recurse into the cell's text, covering nested tables */
            resetCount += resetScaleInContainer(cells[j].texts[0]);
        }
    }

    return resetCount;
}

// =========================================
// メイン処理 / Main
// =========================================

/**
 * 全ストーリー（表セルを含む）の文字比率を 100% に戻す
 * @returns {void}
 */
function main() {
    var stories = targetDocument.stories.everyItem().getElements();
    var resetRangeCount = 0;

    for (var i = 0; i < stories.length; i++) {
        resetRangeCount += resetScaleInContainer(stories[i]);
    }

    alert(localize(LABELS.alert.doneTitle) + localize(LABELS.alert.changedCount) + resetRangeCount);
}

/* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, localize(LABELS.undo.resetScale));

})();
