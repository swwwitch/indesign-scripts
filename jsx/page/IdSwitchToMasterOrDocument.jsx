#target indesign

/*

### 概要

アクティブページが親（マスター）ページかドキュメントページかを判定し、対応するもう一方へ切り替えます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSwitchToMasterOrDocument.md

### Overview

Detects whether the active page is a parent (master) page or a document page and switches to its counterpart.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSwitchToMasterOrDocument.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdSwitchToMasterOrDocument";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSwitchToMasterOrDocument.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSwitchToMasterOrDocument.md"; /* README (English) */

// Original idea
// https://creativepro.com/files/kahrel/indesign/go_to_master.html

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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
            noReturnPage:  { ja: "戻るページ情報がありません。", en: "No return page information found." },
            noAppliedMaster: { ja: "このページには親ページが適用されていません。", en: "No parent page is applied to this page." },
            errorOccurred: { ja: "エラーが発生しました: ", en: "An error occurred: " }
        },
        undo: {
            switchPage: { ja: "親ページ／ドキュメントページの切り替え", en: "Switch Parent / Document Page" }
        }
    };

    // =========================================
    // ページ探索 / Page lookup
    // =========================================

    /**
     * ドキュメント内からページ名が一致するページを探す
     * @param {Document} targetDoc 対象ドキュメント
     * @param {string} pageName 探すページ名
     * @returns {Page|null} 見つかったページ。存在しない場合は null
     */
    function findPageByName(targetDoc, pageName) {
        var documentPages = targetDoc.pages;
        for (var i = 0; i < documentPages.length; i++) {
            if (documentPages[i].name === pageName) return documentPages[i];
        }
        return null;
    }

    /**
     * ドキュメントページに適用されている親ページ側の対応ページを取得する
     * @param {Page} documentPage 対象のドキュメントページ
     * @returns {Page|null} 対応する親ページ。適用がない場合は null
     */
    function getAppliedMasterPage(documentPage) {
        if (!documentPage.appliedMaster) {
            alert(getLabel(LABELS.alert.noAppliedMaster));
            return null;
        }
        var masterPages = documentPage.appliedMaster.pages;
        var masterPageIndex = (masterPages.length === 1)
            ? 0
            : (documentPage.side === PageSideOptions.LEFT_HAND ? 0 : 1);
        return masterPages[masterPageIndex];
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 親ページとドキュメントページを相互に切り替える
     * @returns {void}
     */
    function main() {
        try {
            var layoutWindow = app.windows[0];
            var activeDoc    = app.activeDocument;
            var activePage   = layoutWindow.activePage;

            if (activePage.parent instanceof MasterSpread) {
                /* 親ページ表示中 → 退避したドキュメントページへ戻る / On a parent page: return to the stored document page */
                var returnPage = findPageByName(activeDoc, activeDoc.label);
                if (returnPage) {
                    layoutWindow.activePage = returnPage;
                    activeDoc.label = "";
                } else {
                    alert(getLabel(LABELS.alert.noReturnPage));
                }
            } else {
                /* ドキュメントページ表示中 → 親ページへ移動 / On a document page: jump to the parent page */
                activeDoc.label = activePage.name;
                var masterPage = getAppliedMasterPage(activePage);
                if (masterPage) layoutWindow.activePage = masterPage;
            }
        } catch (e) {
            alert(getLabel(LABELS.alert.errorOccurred) + e);
        }
    }

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undo.switchPage));

})();
