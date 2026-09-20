#target indesign

/*

### 概要

選択した2つのテキストフレームを連結し、1つのストーリーにします。

詳細は README を参照してください。

### Overview

Links the two selected text frames so that they share a single story.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTextFrameLinker";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2023年12月26日";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-20";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTextFrameLinker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTextFrameLinker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n04ceaf4955a0"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

    var currentLanguage = getCurrentLang();

    var LABELS = {
        alert: {
            noDocument:    { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            needTwoFrames: { ja: "2つのテキストフレームを選択してください。", en: "Please select two text frames." },
            overflows:     { ja: "最初のテキストフレームにはオーバーフローテキストがあります。", en: "The first text frame has overset text." }
        },
        undo: {
            linkFrames: { ja: "テキストフレームを連結", en: "Link Text Frames" }
        }
    };

    /**
     * ラベルを現在の言語で取得する
     * @param {object} labelEntry ja / en を持つラベルオブジェクト
     * @returns {string} 現在の言語のラベル文字列
     */
    function getLabel(labelEntry) {
        return labelEntry[currentLanguage];
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択がちょうど2つのテキストフレームかを判定する
     * @param {Array} selectedItems 選択オブジェクトの配列
     * @returns {boolean} 2つともテキストフレームなら true
     */
    function isTwoTextFrames(selectedItems) {
        if (selectedItems.length !== 2) {
            return false;
        }
        for (var i = 0; i < selectedItems.length; i++) {
            if (!(selectedItems[i] instanceof TextFrame)) {
                return false;
            }
        }
        return true;
    }

    /**
     * 選択した2つのテキストフレームを連結する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var selectedItems = app.activeDocument.selection;
        if (!isTwoTextFrames(selectedItems)) {
            alert(getLabel(LABELS.alert.needTwoFrames));
            return;
        }

        var firstFrame = selectedItems[0];
        var secondFrame = selectedItems[1];

        /* あふれたテキストが2つ目のフレームの内容を押し出すため、連結しない / Skip linking: overset text would push out the second frame's own content */
        if (firstFrame.overflows) {
            alert(getLabel(LABELS.alert.overflows));
            return;
        }

        firstFrame.nextTextFrame = secondFrame;
    }

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undo.linkFrames));

})();
