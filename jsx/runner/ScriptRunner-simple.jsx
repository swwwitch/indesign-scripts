#target indesign

/*
 * ScriptRunner-simple.jsx
 *
 * ExtendScript ファイルを 1 つ選んで実行するだけの最小構成のランチャーです。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ScriptRunner-simple";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/ScriptRunner-simple.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/ScriptRunner-simple.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* ファイル選択ダイアログの絞り込み条件 / Filter used by the file-picker dialog */
var SCRIPT_FILE_FILTER = "*.jsx;*.js;*.jsxbin";

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
        title: { ja: "実行するExtendScriptを選択", en: "Select an ExtendScript to run" }
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
// メイン処理 / Main
// =========================================

var selectedScriptFile = File.openDialog(localize(LABELS.dialog.title), SCRIPT_FILE_FILTER);

/* 実行されるスクリプト側で取り消し単位を管理するため、ここでは doScript でラップしない
   / The launched script manages its own undo grouping, so no doScript wrapper here */
if (selectedScriptFile) $.evalFile(selectedScriptFile);

})();
