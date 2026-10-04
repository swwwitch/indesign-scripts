#target indesign

/*

### 概要

任意の ExtendScript ファイル（.jsx / .jsxbin / .js）をダイアログで選んで実行するランチャーです。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdScriptRunner.md

### Overview

A launcher that runs any ExtendScript file (.jsx / .jsxbin / .js) chosen from a dialog.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdScriptRunner.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdScriptRunner";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdScriptRunner.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdScriptRunner.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* ファイル選択ダイアログの絞り込み条件 / Filter used by the file-picker dialog */
var SCRIPT_FILE_FILTER = "ExtendScript:*.jsx;*.jsxbin;*.js";

/* 実行を許可する拡張子 / Extensions that may be executed */
var EXECUTABLE_EXTENSION_PATTERN = /\.(jsx|jsxbin|js)$/i;

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
    dialog: {
        title: { ja: "実行するスクリプトを選択してください", en: "Select a script to execute" }
    },
    alert: {
        fileNotFound: { ja: "選択したスクリプトファイルが存在しません:", en: "The selected script file does not exist:" },
        invalidFile:  { ja: "実行できないファイル形式です:", en: "Unsupported file type:" },
        errorTitle:   { ja: "スクリプトの実行中にエラーが発生しました:", en: "An error occurred while executing the script:" }
    },
    errorField: {
        file:        { ja: "ファイル", en: "File" },
        line:        { ja: "行番号", en: "Line" },
        errorNumber: { ja: "エラー番号", en: "Error Number" }
    }
};

// =========================================
// ファイル選択とエラー整形 / File picking and error formatting
// =========================================

/**
 * 実行するスクリプトファイルを選ばせる
 * @returns {File|null} 選択したファイル。キャンセル時は null
 */
function pickScriptFile() {
    return File.openDialog(getLabel(LABELS.dialog.title), SCRIPT_FILE_FILTER, false);
}

/**
 * 実行時エラーの内容を表示用の文字列に整形する
 * @param {File} targetFile 実行しようとしたファイル
 * @param {Error} caughtError 捕捉した例外
 * @returns {string} 表示するエラーメッセージ
 */
function buildErrorText(targetFile, caughtError) {
    var errorFileName = (caughtError && caughtError.fileName)
        ? File(caughtError.fileName).name
        : (targetFile ? targetFile.name : "unknown");
    var errorLine    = (caughtError && caughtError.line !== undefined) ? caughtError.line : "unknown";
    var errorNumber  = (caughtError && caughtError.number !== undefined) ? caughtError.number : "unknown";
    var errorMessage = (caughtError && caughtError.message) ? caughtError.message : String(caughtError);

    return getLabel(LABELS.alert.errorTitle) + "\n\n" +
        labelValueText(LABELS.errorField.file, errorFileName) + "\n" +
        labelValueText(LABELS.errorField.line, errorLine) + "\n" +
        labelValueText(LABELS.errorField.errorNumber, errorNumber) + "\n\n" +
        errorMessage;
}

// =========================================
// メイン処理 / Main
// =========================================

var selectedScriptFile = pickScriptFile();
if (!selectedScriptFile) return;

if (!selectedScriptFile.exists) {
    alert(getLabel(LABELS.alert.fileNotFound) + "\n" + selectedScriptFile.fsName);
    return;
}

if (!EXECUTABLE_EXTENSION_PATTERN.test(selectedScriptFile.name)) {
    alert(getLabel(LABELS.alert.invalidFile) + "\n" + selectedScriptFile.name);
    return;
}

try {
    /* 実行されるスクリプト側で取り消し単位を管理するため、ここでは doScript でラップしない
       / The launched script manages its own undo grouping, so no doScript wrapper here */
    $.evalFile(selectedScriptFile);
} catch (e) {
    alert(buildErrorText(selectedScriptFile, e));
}

})();
