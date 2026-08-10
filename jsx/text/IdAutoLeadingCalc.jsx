#target indesign

/*
 * IdAutoLeadingCalc.jsx
 *
 * 選択テキストの現在の行送り（絶対値）と文字サイズから行送り％を段落ごとに逆算し、自動行送りに切り替えます。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdAutoLeadingCalc";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-09";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAutoLeadingCalc.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAutoLeadingCalc.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 逆算した行送り％の丸め桁数（小数第 1 位）/ Rounding of the back-calculated leading percentage (1 decimal) */
var LEADING_PERCENT_ROUND_FACTOR = 10;

/* 段落を取り出せる選択オブジェクトの型 / Selection types that expose paragraphs */
var TEXT_SELECTION_TYPES = {
    Character: true, Word: true, TextStyleRange: true, Paragraph: true,
    Line: true, TextColumn: true, Text: true, InsertionPoint: true, Story: true
};

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
        noDocument:  { ja: "ドキュメントを開いてください。", en: "Please open a document." },
        noSelection: { ja: "テキストが選択されていません。", en: "No text is selected." }
    },
    undo: {
        applyAutoLeading: { ja: "自動行送りに変換", en: "Convert to Auto Leading" }
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
    if (TEXT_SELECTION_TYPES[typeName]) {
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
        alert(localize(LABELS.alert.noDocument));
        return;
    }

    var selectionItems = app.selection;
    if (!selectionItems || selectionItems.length === 0) {
        alert(localize(LABELS.alert.noSelection));
        return;
    }

    var paragraphTargets = [];
    for (var i = 0; i < selectionItems.length; i++) {
        collectParagraphsFromItem(selectionItems[i], paragraphTargets);
    }

    if (paragraphTargets.length === 0) {
        alert(localize(LABELS.alert.noSelection));
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
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, localize(LABELS.undo.applyAutoLeading));

})();
