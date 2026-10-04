#target indesign

/*

### 概要

選択テキストの各段落に、段落背景色の高さに合わせた不可視の段落境界線（前境界線）をスペーサーとして設定します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdParagraphShadingMatchRule.md

### Overview

Sets an invisible paragraph rule above on each selected paragraph, sized to the paragraph shading height so it acts as a spacer.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdParagraphShadingMatchRule.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdParagraphShadingMatchRule";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-12";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdParagraphShadingMatchRule.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdParagraphShadingMatchRule.md"; /* README (English) */

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
        error: {
            selectText: { ja: "テキストを選択してください。", en: "Please select text." }
        },
        undo: {
            setParagraphRules: { ja: "段落ルール設定", en: "Set Paragraph Rules" }
        }
    };

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
