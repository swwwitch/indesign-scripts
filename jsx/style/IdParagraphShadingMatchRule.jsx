#target indesign

/*

### 概要

選択テキストの各段落に、段落背景色の高さに合わせた不可視の段落境界線（前境界線）をスペーサーとして設定します。

詳細は README を参照してください。

### Overview

Sets an invisible paragraph rule above on each selected paragraph, sized to the paragraph shading height so it acts as a spacer.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdParagraphShadingMatchRule";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-12";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-12";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdParagraphShadingMatchRule.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdParagraphShadingMatchRule.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* テキストとして扱う選択オブジェクトの種別 / Selection types treated as text */
    var TEXT_TYPE_NAMES = {
        Text: true,
        InsertionPoint: true,
        Word: true,
        Line: true,
        TextStyleRange: true,
        Paragraph: true
    };

    /* 不可視の段落境界線に使うスウォッチ名 / Swatch name used for the invisible paragraph rule */
    var NONE_SWATCH_NAME = "None";

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
        error: {
            selectText: { ja: "テキストを選択してください。", en: "Please select text." }
        },
        undo: {
            setParagraphRules: { ja: "段落ルール設定", en: "Set Paragraph Rules" }
        }
    };

    /**
     * ドット区切りキーでラベルを取得する
     * @param {string} labelKey 例: "error.selectText"
     * @returns {string} 現在の言語のラベル文字列
     */
    function getLabel(labelKey) {
        var keyParts = labelKey.split(".");
        var node = LABELS;

        for (var i = 0; i < keyParts.length; i++) {
            if (!node || !node.hasOwnProperty(keyParts[i])) {
                throw new Error("Missing label path: " + labelKey);
            }
            node = node[keyParts[i]];
        }

        if (!node || !node.hasOwnProperty(currentLang)) {
            throw new Error("Missing label language: " + labelKey + " / " + currentLang);
        }

        return node[currentLang];
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * 数値に変換する（変換できない場合は 0）
     * @param {*} rawValue 変換元の値
     * @returns {number} 数値。変換できない場合は 0
     */
    function toNumberOrZero(rawValue) {
        var numericValue = Number(rawValue);
        return isNaN(numericValue) ? 0 : numericValue;
    }

    /**
     * 選択オブジェクトから段落を持つテキストを取り出す
     * @param {object} selectionItem 選択オブジェクト
     * @returns {object|null} テキストオブジェクト。該当しない場合は null
     */
    function getTargetTextFromSelection(selectionItem) {
        if (!selectionItem) return null;
        if (selectionItem.hasOwnProperty("baseline")) return selectionItem;
        if (TEXT_TYPE_NAMES[selectionItem.constructor.name]) return selectionItem;
        if (selectionItem.hasOwnProperty("parentStory")) return selectionItem.parentStory;
        return null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択テキストの各段落に不可視の段落境界線を設定する
     * @returns {void}
     */
    function main() {
        if (app.selection.length === 0) {
            alert(getLabel("error.selectText"));
            return;
        }

        var targetText = getTargetTextFromSelection(app.selection[0]);
        if (!targetText || !targetText.paragraphs || targetText.paragraphs.length === 0) {
            alert(getLabel("error.selectText"));
            return;
        }

        var noneSwatch = app.activeDocument.swatches.itemByName(NONE_SWATCH_NAME);
        var paragraphs = targetText.paragraphs;

        for (var i = 0; i < paragraphs.length; i++) {
            var paragraph = paragraphs[i];

            /* 段落背景色の上オフセットとフォントサイズの合計が、背景領域の高さの目安になる
               / The shading top offset plus the font size approximates the height of the shaded area */
            var shadingTopOffsetPt = toNumberOrZero(paragraph.paragraphShadingTopOffset);
            var fontSizePt         = toNumberOrZero(paragraph.pointSize);

            /* 見た目の線ではなく、レイアウト調整用の不可視スペーサーとして使う
               / Used as an invisible layout spacer, not as a visible line */
            paragraph.ruleAbove            = true;
            paragraph.ruleAboveLineWeight  = shadingTopOffsetPt + fontSizePt;
            paragraph.ruleAboveColor       = noneSwatch;
            paragraph.keepRuleAboveInFrame = true;
        }
    }

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT,
        getLabel("undo.setParagraphRules") + " " + SCRIPT_VERSION);

})();
