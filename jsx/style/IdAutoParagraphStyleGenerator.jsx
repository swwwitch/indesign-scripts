#target indesign

/*

### 概要

［基本段落］などスタイル未設定の段落をフォント・サイズ・行送りごとにまとめ、段落スタイルを自動生成して適用します。

詳細は README を参照してください。

### Overview

Groups unstyled paragraphs such as [Basic Paragraph] by font, size and leading, then generates and applies paragraph styles for them.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdAutoParagraphStyleGenerator"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v3.4";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-03-14";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAutoParagraphStyleGenerator.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAutoParagraphStyleGenerator.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 自動生成の対象とする段落スタイル名 / Paragraph style names treated as "unstyled" */
var TARGET_STYLE_NAMES = ["[基本段落]", "基本段落", "[Basic Paragraph]", "No Paragraph Style"];

/* 生成する段落スタイル名の接頭辞 / Prefix of the generated paragraph style names */
var GENERATED_STYLE_PREFIX = "AutoStyle_";

/* グループ判定に使う丸め幅（pt）と、スタイル名に出す表示の丸め幅 / Rounding step for grouping (pt) and for the displayed size label */
var GROUPING_ROUND_STEP = 0.01;
var SIZE_LABEL_ROUND_STEP = 0.1;

// =========================================
// 単位換算 / Unit conversion
// =========================================

var POINTS_PER_MM = 72 / 25.4;
var POINTS_PER_Q  = POINTS_PER_MM * 0.25;  /* 1Q = 0.25mm */
var POINTS_PER_H  = POINTS_PER_Q;          /* 1H = 0.25mm（和文組版の H）/ 1H = 0.25mm (Japanese typesetting) */

/**
 * 単位が Q かどうかを判定する
 * @param {MeasurementUnits} unitEnum 判定する単位
 * @returns {boolean} Q なら true
 */
function isUnitQ(unitEnum) {
    try {
        return unitEnum === MeasurementUnits.Q;
    } catch (e) {
        return false;
    }
}

/**
 * 単位が H かどうかを判定する
 * @param {MeasurementUnits} unitEnum 判定する単位
 * @returns {boolean} H なら true
 */
function isUnitH(unitEnum) {
    try {
        if (typeof MeasurementUnits.HA !== "undefined" && unitEnum === MeasurementUnits.HA) return true;
    } catch (e) {}
    return false;
}

/**
 * 現在の単位の数値をポイントへ正規化する
 * @param {number} value 現在の単位での数値
 * @param {MeasurementUnits} unitEnum 現在の単位
 * @returns {number} ポイント値
 */
function convertToPoints(value, unitEnum) {
    if (typeof value !== "number") return value;
    if (isUnitQ(unitEnum)) return value * POINTS_PER_Q;
    if (isUnitH(unitEnum)) return value * POINTS_PER_H;
    return value; /* pt 前提 / Assume points */
}

/**
 * ポイント値を現在の単位の数値へ戻す
 * @param {number} pointValue ポイント値
 * @param {MeasurementUnits} unitEnum 現在の単位
 * @returns {number} 現在の単位での数値
 */
function convertFromPoints(pointValue, unitEnum) {
    if (typeof pointValue !== "number") return pointValue;
    if (isUnitQ(unitEnum)) return pointValue / POINTS_PER_Q;
    if (isUnitH(unitEnum)) return pointValue / POINTS_PER_H;
    return pointValue; /* pt 前提 / Assume points */
}

/**
 * 指定した刻みで数値を丸める
 * @param {number} value 対象の数値
 * @param {number} step 丸めの刻み
 * @returns {number} 丸めた数値
 */
function roundToStep(value, step) {
    return Math.round(value / step) * step;
}

/**
 * 環境設定［単位と増減値］＞［テキストサイズ］の単位を取得する
 * @param {Document} targetDoc 対象ドキュメント
 * @returns {MeasurementUnits|null} テキストサイズの単位。取得できない場合は null
 */
function getTextSizeUnit(targetDoc) {
    try {
        return targetDoc.viewPreferences.textSizeMeasurementUnits;
    } catch (e) {
        return null;
    }
}

/**
 * 環境設定［単位と増減値］＞［組版］の単位を取得する
 * @param {Document} targetDoc 対象ドキュメント
 * @returns {MeasurementUnits|null} 組版の単位。取得できない場合は null
 */
function getTypographicUnit(targetDoc) {
    try {
        return targetDoc.viewPreferences.typographicMeasurementUnits;
    } catch (e) {
        return null;
    }
}

/**
 * ポイント値からスタイル名に使うサイズ表記を作る
 * @param {number} pointValue ポイント値
 * @param {MeasurementUnits} unitEnum テキストサイズの単位
 * @returns {{value: string, suffix: string}} 表示用の数値と単位サフィックス
 */
function formatSizeLabelFromPoints(pointValue, unitEnum) {
    if (isUnitQ(unitEnum)) {
        var quarterMillimeters = roundToStep(pointValue / POINTS_PER_Q, SIZE_LABEL_ROUND_STEP);
        return { value: (Math.round(quarterMillimeters * 10) / 10).toString(), suffix: "Q" };
    }
    var roundedPoints = roundToStep(pointValue, SIZE_LABEL_ROUND_STEP);
    return { value: (Math.round(roundedPoints * 10) / 10).toString(), suffix: "pt" };
}

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
        noTarget:     { ja: "対象となる段落が見つかりませんでした。", en: "No target paragraph was found." },
        skippedNote:  { ja: "\n（書式混在によりスキップした段落: ", en: "\n(Paragraphs skipped due to mixed formatting: " },
        doneTitle:    { ja: "完了しました。", en: "Done." },
        createdCount: { ja: "\n作成したスタイル数: ", en: "\nStyles created: " },
        appliedCount: { ja: "\n適用した段落数: ", en: "\nParagraphs styled: " },
        skippedCount: { ja: "\n\n※書式混在でスキップ: ", en: "\n\nSkipped (mixed formatting): " }
    },
    undo: {
        generateStyles: { ja: "段落スタイル自動整理", en: "Generate Paragraph Styles" }
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
// ユーティリティ / Utilities
// =========================================

/**
 * オブジェクトが空かどうかを判定する
 * @param {object} targetObject 判定するオブジェクト
 * @returns {boolean} 自前のプロパティが 1 つもなければ true
 */
function isObjectEmpty(targetObject) {
    for (var propertyName in targetObject) {
        if (targetObject.hasOwnProperty(propertyName)) return false;
    }
    return true;
}

/**
 * 段落スタイル名が自動生成の対象かどうかを判定する
 * @param {string} styleName 段落スタイル名
 * @returns {boolean} 対象なら true
 */
function isTargetStyleName(styleName) {
    for (var i = 0; i < TARGET_STYLE_NAMES.length; i++) {
        if (styleName === TARGET_STYLE_NAMES[i]) return true;
    }
    return false;
}

/**
 * 既存スタイルと重複しないスタイル名を作る
 * @param {Document} targetDoc 対象ドキュメント
 * @param {string} baseStyleName 基準となるスタイル名
 * @returns {string} 重複しないスタイル名
 */
function makeUniqueStyleName(targetDoc, baseStyleName) {
    var uniqueStyleName = baseStyleName;
    var duplicateIndex = 2;
    while (targetDoc.paragraphStyles.itemByName(uniqueStyleName).isValid) {
        uniqueStyleName = baseStyleName + "_" + duplicateIndex;
        duplicateIndex++;
    }
    return uniqueStyleName;
}

// =========================================
// 段落の収集 / Paragraph collection
// =========================================

/**
 * 対象段落をフォント・サイズ・行送りごとにまとめる
 * @param {Document} targetDoc 対象ドキュメント
 * @param {MeasurementUnits} textSizeUnit テキストサイズの単位
 * @param {MeasurementUnits} typographicUnit 組版の単位
 * @returns {{groups: object, skippedCount: number}} グループと、書式混在でスキップした段落数
 */
function collectParagraphGroups(targetDoc, textSizeUnit, typographicUnit) {
    var styleGroups  = {};
    var skippedCount = 0;

    for (var i = 0; i < targetDoc.stories.length; i++) {
        var story = targetDoc.stories[i];

        for (var j = 0; j < story.paragraphs.length; j++) {
            var paragraph = story.paragraphs[j];
            if (!isTargetStyleName(paragraph.appliedParagraphStyle.name)) continue;

            var fontFamilyName = paragraph.appliedFont.name;
            var fontStyleName  = paragraph.fontStyle;

            /* pointSize / leading は現在の単位での数値になり得るため pt に正規化 / Normalize to points; these can be in the current unit */
            var pointSizeInPoints = convertToPoints(paragraph.pointSize, textSizeUnit);
            var leadingInPoints   = paragraph.leading;
            if (leadingInPoints !== Leading.AUTO) {
                leadingInPoints = convertToPoints(leadingInPoints, typographicUnit);
            }

            /* 段落内で書式が混在していると数値以外が返るためスキップ / Mixed formatting yields non-numeric values, so skip */
            if (typeof pointSizeInPoints !== "number" ||
                (leadingInPoints !== Leading.AUTO && typeof leadingInPoints !== "number")) {
                skippedCount++;
                continue;
            }

            /* 浮動小数の揺れで別グループにならないよう正規化 / Normalize so float noise does not split groups */
            var normalizedPointSize = roundToStep(pointSizeInPoints, GROUPING_ROUND_STEP);
            var normalizedLeading   = (leadingInPoints === Leading.AUTO)
                ? Leading.AUTO
                : roundToStep(leadingInPoints, GROUPING_ROUND_STEP);
            var leadingKeyPart = (leadingInPoints === Leading.AUTO) ? "Auto" : normalizedLeading;

            var groupKey = fontFamilyName + "_" + fontStyleName + "_" + normalizedPointSize + "_" + leadingKeyPart;

            if (!styleGroups[groupKey]) {
                styleGroups[groupKey] = {
                    appliedFont: paragraph.appliedFont,
                    fontStyle: fontStyleName,
                    pointSizeInPoints: normalizedPointSize,
                    leadingInPoints: normalizedLeading,
                    paragraphs: []
                };
            }
            styleGroups[groupKey].paragraphs.push(paragraph);
        }
    }

    return { groups: styleGroups, skippedCount: skippedCount };
}

// =========================================
// メイン処理 / Main
// =========================================

if (app.documents.length === 0) {
    alert(localize(LABELS.alert.noDocument));
    return;
}

var activeDoc       = app.activeDocument;
var textSizeUnit    = getTextSizeUnit(activeDoc);
var typographicUnit = getTypographicUnit(activeDoc);

var collectResult = collectParagraphGroups(activeDoc, textSizeUnit, typographicUnit);
var styleGroups   = collectResult.groups;
var skippedCount  = collectResult.skippedCount;

if (isObjectEmpty(styleGroups)) {
    var noTargetMessage = localize(LABELS.alert.noTarget);
    if (skippedCount > 0) {
        noTargetMessage += localize(LABELS.alert.skippedNote) + skippedCount + ")";
    }
    alert(noTargetMessage);
    return;
}

var createdStyleCount   = 0;
var appliedParagraphCount = 0;

app.doScript(function () {
    for (var groupKey in styleGroups) {
        if (!styleGroups.hasOwnProperty(groupKey)) continue;

        var styleGroup = styleGroups[groupKey];

        /* テキストサイズ単位に連動したスタイル名を作る / Build a style name that follows the text-size unit */
        var sizeLabel = formatSizeLabelFromPoints(styleGroup.pointSizeInPoints, textSizeUnit);
        var baseStyleName = GENERATED_STYLE_PREFIX + (createdStyleCount + 1) + "_" + sizeLabel.value + sizeLabel.suffix;

        var newParagraphStyle = activeDoc.paragraphStyles.add({
            name: makeUniqueStyleName(activeDoc, baseStyleName),
            appliedFont: styleGroup.appliedFont,
            fontStyle: styleGroup.fontStyle,
            pointSize: convertFromPoints(styleGroup.pointSizeInPoints, textSizeUnit),
            leading: (styleGroup.leadingInPoints === Leading.AUTO)
                ? Leading.AUTO
                : convertFromPoints(styleGroup.leadingInPoints, typographicUnit)
        });

        for (var i = 0; i < styleGroup.paragraphs.length; i++) {
            try {
                /* 第 2 引数 true でオーバーライドを消去 / Pass true to clear overrides */
                styleGroup.paragraphs[i].applyParagraphStyle(newParagraphStyle, true);
                appliedParagraphCount++;
            } catch (e) {}
        }
        createdStyleCount++;
    }
}, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, localize(LABELS.undo.generateStyles));

var resultMessage = localize(LABELS.alert.doneTitle) +
    localize(LABELS.alert.createdCount) + createdStyleCount +
    localize(LABELS.alert.appliedCount) + appliedParagraphCount;

if (skippedCount > 0) {
    resultMessage += localize(LABELS.alert.skippedCount) + skippedCount;
}

alert(resultMessage);

})();
