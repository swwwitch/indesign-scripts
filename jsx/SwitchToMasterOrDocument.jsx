#target indesign

/*
    スクリプト名：SwitchToMasterOrDocument.jsx

    スクリプトの概要：
    アクティブページがドキュメントページの場合はマスターにジャンプ、
    マスターページの場合は元のドキュメントページに戻ります。

    Overview:
    If the active page is a document page, switch to its master page.
    If it is a master page, switch back to the previously active document page.

    処理の流れ / Flow:
    1. 現在のページを判定 / Check current page
    2. ドキュメントページならマスターへジャンプ / Jump to master from document
    3. マスターページなら元のドキュメントページに戻る / Return to document from master

    更新履歴 / Update history:
    2025-07-02 改訂：ドキュメントページとマスターページの切り替え対応
    2025-07-02 Updated: Support switching between document and master pages

    限定条件 / Notes:
    - ドキュメントが開いており、アクティブページが存在する必要があります。
    - Requires an open document and an active page.

    Original Idea: https://creativepro.com/files/kahrel/indesign/go_to_master.html
*/

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

// -------------------------------
// ラベル定義 (日英)  Define labels (JA/EN)
// -------------------------------
var lang = getCurrentLang();
var LABELS = {
    noReturnPage: { ja: "戻るページ情報がありません。", en: "No return page information found." },
    noMaster: { ja: "このページにはマスターが適用されていません。", en: "No master applied to this page." },
    errorOccurred: { ja: "エラーが発生しました: ", en: "An error occurred: " }
};

function main() {
    try {
        var win = app.windows[0];
        var doc = app.activeDocument;
        var currentPage = win.activePage;

        if (currentPage.parent instanceof MasterSpread) {
            // マスターページ表示中 → 保存されたページへ戻る / On master page: return to saved page
            var targetPage = findPageByName(doc, doc.label);
            if (targetPage) {
                win.activePage = targetPage;
                doc.label = "";
            } else {
                alert(LABELS.noReturnPage[lang]);
            }
        } else {
            // ドキュメントページ表示中 → マスターへ / On document page: jump to master
            doc.label = currentPage.name;
            var masterPage = getMasterPage(currentPage);
            if (masterPage) win.activePage = masterPage;
        }
    } catch (e) {
        alert(LABELS.errorOccurred[lang] + e);
    }
}

// ページ名でページを探す / Find page by name
function findPageByName(doc, name) {
    var pages = doc.pages;
    for (var i = 0; i < pages.length; i++) {
        if (pages[i].name === name) return pages[i];
    }
    return null;
}

// マスター側のページを取得 / Get corresponding master page
function getMasterPage(page) {
    if (!page.appliedMaster) {
        alert(LABELS.noMaster[lang]);
        return null;
    }
    var index = (page.appliedMaster.pages.length === 1) ? 0
        : (page.side === PageSideOptions.LEFT_HAND ? 0 : 1);
    return page.appliedMaster.pages[index];
}

main();