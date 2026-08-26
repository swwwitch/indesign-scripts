#target indesign

/*

### 概要

選択した段落の「段落の開始位置」を「次の段（フレーム）」に設定します。

詳細は README を参照してください。

### Overview

Sets "Start Paragraph" to "In Next Column" (next frame) on the selected paragraphs.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdKeepOptionNextColumn";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-06-27";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdKeepOptionNextColumn.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdKeepOptionNextColumn.md

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
            noDocument:  { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noSelection: { ja: "段落またはテキストを選択してください。", en: "Please select a paragraph or text." },
            unsupported: { ja: "選択対象には段落開始位置を設定できません。", en: "The selected object does not support paragraph start options." }
        },
        undo: {
            startNextColumn: { ja: "段落の開始位置を次の段に設定", en: "Start Paragraph in Next Column" }
        }
    };

    /**
     * ラベルを現在の言語で取得する
     * @param {object} labelEntry ja / en を持つラベルオブジェクト
     * @returns {string} 現在の言語のラベル文字列
     */
    function localize(labelEntry) {
        return labelEntry[currentLanguage];
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択した段落の開始位置を「次の段」に設定する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(localize(LABELS.alert.noDocument));
            return;
        }

        var currentDocument = app.activeDocument;
        if (currentDocument.selection.length === 0) {
            alert(localize(LABELS.alert.noSelection));
            return;
        }

        var selectedText = currentDocument.selection[0];

        /* 段落開始位置を設定できる対象か確認 / Verify the selection supports the start-paragraph option */
        if (!selectedText || !selectedText.hasOwnProperty("startParagraph")) {
            alert(localize(LABELS.alert.unsupported));
            return;
        }

        selectedText.startParagraph = StartParagraph.NEXT_COLUMN;
    }

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, localize(LABELS.undo.startNextColumn));

})();
