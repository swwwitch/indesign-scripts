#target indesign

/*

### 概要

任意の ExtendScript ファイル（.jsx / .jsxbin / .js）をダイアログで選んで実行するランチャーです。

詳細は README を参照してください。

### Overview

A launcher that runs any ExtendScript file (.jsx / .jsxbin / .js) chosen from a dialog.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdScriptRunner";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdScriptRunner.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdScriptRunner.md

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

/**
 * UI 言語を判定する
 * @returns {string} "ja" または "en"
 */
function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var currentLang = getCurrentLang();

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

/**
 * ラベルを現在の言語で取得する
 * @param {object} labelEntry ja / en を持つラベルオブジェクト
 * @returns {string} 現在の言語のラベル文字列
 */
function localize(labelEntry) {
    return labelEntry[currentLang];
}

// =========================================
// ファイル選択とエラー整形 / File picking and error formatting
// =========================================

/**
 * 実行するスクリプトファイルを選ばせる
 * @param {Folder} [startFolder] ダイアログの初期フォルダ
 * @returns {File|null} 選択したファイル。キャンセル時は null
 */
function pickScriptFile(startFolder) {
    try {
        if (startFolder && startFolder.exists) Folder.current = startFolder;
    } catch (e) {}
    return File.openDialog(localize(LABELS.dialog.title), SCRIPT_FILE_FILTER, false);
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

    return localize(LABELS.alert.errorTitle) + "\n\n" +
        localize(LABELS.errorField.file) + ": " + errorFileName + "\n" +
        localize(LABELS.errorField.line) + ": " + errorLine + "\n" +
        localize(LABELS.errorField.errorNumber) + ": " + errorNumber + "\n\n" +
        errorMessage;
}

// =========================================
// メイン処理 / Main
// =========================================

var selectedScriptFile = pickScriptFile(null);
if (!selectedScriptFile) return;

if (!selectedScriptFile.exists) {
    alert(localize(LABELS.alert.fileNotFound) + "\n" + selectedScriptFile.fsName);
    return;
}

if (!EXECUTABLE_EXTENSION_PATTERN.test(selectedScriptFile.name)) {
    alert(localize(LABELS.alert.invalidFile) + "\n" + selectedScriptFile.name);
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
