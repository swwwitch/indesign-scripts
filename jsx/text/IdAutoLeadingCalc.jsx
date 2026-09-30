#target indesign

/*

### 概要

選択テキストの現在の行送り（絶対値）と文字サイズから行送り％を段落ごとに逆算し、自動行送りに切り替えます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAutoLeadingCalc.md

### Overview

Derives the leading percentage of each paragraph from its current absolute leading and font size, then switches it to auto leading.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAutoLeadingCalc.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdAutoLeadingCalc";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAutoLeadingCalc.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAutoLeadingCalc.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 逆算した行送り％の丸め桁数（小数第 1 位）/ Rounding of the back-calculated leading percentage (1 decimal) */
var LEADING_PERCENT_ROUND_FACTOR = 10;

/* テキストとして扱う選択の型 / Selection types treated as text */
var TEXT_SELECTION_TYPES = {
    InsertionPoint: true, Character: true, Word: true, Line: true,
    TextStyleRange: true, Paragraph: true, TextColumn: true, Text: true
};

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
    alert: {
        noDocument:  { ja: "ドキュメントを開いてください。", en: "Please open a document." },
        noSelection: { ja: "テキストが選択されていません。", en: "No text is selected." }
    },
    undo: {
        applyAutoLeading: { ja: "自動行送りに変換", en: "Convert to Auto Leading" }
    }
};

// =========================================
// 段落の収集 / Paragraph collection
// =========================================

/**
 * オブジェクトの型名を安全に取得する
 * @param {object} targetObject 対象オブジェクト
 * @returns {string} 型名。取得できない場合は空文字
 */
function getConstructorName(targetObject) {
    try {
        return (targetObject && targetObject.constructor) ? targetObject.constructor.name : "";
    } catch (e) {
        return "";
    }
}

/**
 * 選択がテキスト（キャレット・文字範囲）かどうかを判定する
 * @param {object} selectionItem 選択オブジェクト
 * @returns {boolean} テキスト上の選択なら true
 */
function isTextSelection(selectionItem) {
    if (selectionItem == null) return false;
    return TEXT_SELECTION_TYPES[selectionItem.constructor.name] === true;
}

/**
 * テキストオブジェクトの段落を対象配列へ追加する
 * @param {object} textObject paragraphs を持つテキストオブジェクト
 * @param {Array<Paragraph>} paragraphTargets 追加先の配列
 * @returns {void}
 */
function addParagraphs(textObject, paragraphTargets) {
    try {
        var paragraphs = textObject.paragraphs;
        for (var i = 0; i < paragraphs.length; i++) {
            paragraphTargets.push(paragraphs[i]);
        }
    } catch (e) {}
}

/**
 * 選択項目 1 つから対象の段落を収集する
 * @param {object} selectionItem 選択オブジェクト
 * @param {Array<Paragraph>} paragraphTargets 追加先の配列
 * @returns {void}
 */
function collectParagraphsFromItem(selectionItem, paragraphTargets) {
    if (!selectionItem) return;
    var typeName = getConstructorName(selectionItem);

    /* テキスト編集モード：選択が触れている段落全体が対象 / Text-edit mode: every touched paragraph */
    if (isTextSelection(selectionItem)) {
        addParagraphs(selectionItem, paragraphTargets);
        return;
    }

    /* 選択ツールでのフレーム選択：そのフレームの本文が対象 / A selected frame contributes its own text */
    if (typeName === "TextFrame") {
        try {
            if (selectionItem.texts.length > 0) addParagraphs(selectionItem.texts[0], paragraphTargets);
        } catch (e) {}
        return;
    }

    /* グループは内部（ネスト含む）の全テキストフレームが対象 / A group contributes every nested text frame */
    if (typeName === "Group") {
        try {
            var innerItems = selectionItem.allPageItems;
            for (var i = 0; i < innerItems.length; i++) {
                if (getConstructorName(innerItems[i]) === "TextFrame" && innerItems[i].texts.length > 0) {
                    addParagraphs(innerItems[i].texts[0], paragraphTargets);
                }
            }
        } catch (e) {}
        return;
    }

    /* 長方形などテキストを保持し得る図形フレーム / Rectangles and similar shapes that may hold text */
    try {
        if (selectionItem.texts && selectionItem.texts.length > 0 && selectionItem.texts[0].contents.length > 0) {
            addParagraphs(selectionItem.texts[0], paragraphTargets);
        }
    } catch (e) {}
}

// =========================================
// 自動行送りの適用 / Auto-leading
// =========================================

/**
 * 1 段落の行送り％を逆算し、自動行送りとして適用する
 * @param {Paragraph} paragraph 対象の段落
 * @returns {void}
 */
function applyAutoLeadingToParagraph(paragraph) {
    try {
        if (paragraph.characters.length === 0) return;

        var firstCharacter = paragraph.characters[0];
        var pointSize      = firstCharacter.pointSize;
        var leadingValue   = firstCharacter.leading;

        /* すでに自動行送りなら逆算できないためスキップ / Already Auto: nothing to back-calculate */
        if (leadingValue === Leading.AUTO) return;

        var leadingInPoints = Number(leadingValue);
        if (isNaN(pointSize) || pointSize <= 0 || isNaN(leadingInPoints) || leadingInPoints <= 0) return;

        var leadingPercent = Math.round((leadingInPoints / pointSize) * 100 * LEADING_PERCENT_ROUND_FACTOR) /
            LEADING_PERCENT_ROUND_FACTOR;

        paragraph.autoLeading = leadingPercent;
        paragraph.leading = Leading.AUTO;

        /* 行送りの基準は仮想ボディの上／右に固定 / Fix the leading basis to the top/right of the virtual body */
        paragraph.leadingModel = LeadingModel.LEADING_MODEL_AKI_BELOW;
    } catch (e) {}
}

// =========================================
// メイン処理 / Main
// =========================================

/**
 * 選択テキストの各段落を自動行送りへ変換する
 * @returns {void}
 */
function main() {
    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }

    var selectionItems = app.selection;
    if (!selectionItems || selectionItems.length === 0) {
        alert(getLabel(LABELS.alert.noSelection));
        return;
    }

    var paragraphTargets = [];
    for (var i = 0; i < selectionItems.length; i++) {
        collectParagraphsFromItem(selectionItems[i], paragraphTargets);
    }

    if (paragraphTargets.length === 0) {
        alert(getLabel(LABELS.alert.noSelection));
        return;
    }

    for (var j = 0; j < paragraphTargets.length; j++) {
        applyAutoLeadingToParagraph(paragraphTargets[j]);
    }

    /* 文字パネルの表示を更新するため、いったん選択を解除して同じ選択を選び直す
       / Deselect and re-select so the Character panel refreshes its cached values */
    try {
        app.select(NothingEnum.NOTHING);
        app.select(selectionItems);
    } catch (e) {}
}

/* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undo.applyAutoLeading));

})();
