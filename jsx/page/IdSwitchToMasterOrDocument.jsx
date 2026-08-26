#target indesign

/*
 * IdSwitchToMasterOrDocument.jsx
 *
 * アクティブページが親（マスター）ページかドキュメントページかを判定し、対応するもう一方へ切り替えます。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdSwitchToMasterOrDocument";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-07-02";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSwitchToMasterOrDocument.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSwitchToMasterOrDocument.md

// Original idea
// https://creativepro.com/files/kahrel/indesign/go_to_master.html

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

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
        noReturnPage:  { ja: "戻るページ情報がありません。", en: "No return page information found." },
        noAppliedMaster: { ja: "このページには親ページが適用されていません。", en: "No parent page is applied to this page." },
        errorOccurred: { ja: "エラーが発生しました: ", en: "An error occurred: " }
    },
    undo: {
        switchPage: { ja: "親ページ／ドキュメントページの切り替え", en: "Switch Parent / Document Page" }
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
        alert(localize(LABELS.alert.noAppliedMaster));
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
                alert(localize(LABELS.alert.noReturnPage));
            }
        } else {
            /* ドキュメントページ表示中 → 親ページへ移動 / On a document page: jump to the parent page */
            activeDoc.label = activePage.name;
            var masterPage = getAppliedMasterPage(activePage);
            if (masterPage) layoutWindow.activePage = masterPage;
        }
    } catch (e) {
        alert(localize(LABELS.alert.errorOccurred) + e);
    }
}

/* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, localize(LABELS.undo.switchPage));
