#target indesign

/*

### 概要

選択したフレーム、指定したページ、またはドキュメント内のアンカー付きオブジェクトを、アンカーやフレームの種類で絞り込んで解除します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdReleaseAnchoredObjects.md

### Overview

Releases the anchored objects in the selected frames, on a chosen page, or throughout the document, filtered by anchor type and frame type.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdReleaseAnchoredObjects.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdReleaseAnchoredObjects";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-25";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdReleaseAnchoredObjects.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdReleaseAnchoredObjects.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    /* グラフィックフレームとして扱う種別 / Types treated as graphic frames */
    var GRAPHIC_FRAME_TYPES = ["Rectangle", "Oval", "Polygon"];

    var LABELS = {
        dialog: {
            title: { ja: "アンカー付きオブジェクトを解除", en: "Release Anchored Objects" }
        },
        panel: {
            target:       { ja: "対象", en: "Target" },
            anchorType:   { ja: "アンカーの種類", en: "Anchor Type" },
            frameType:    { ja: "フレームの種類", en: "Frame Type" },
            afterRelease: { ja: "解除後", en: "After Release" }
        },
        radio: {
            selectedFrames: { ja: "選択したフレーム", en: "Selected Frames" },
            currentPage:    { ja: "現在のページ", en: "Current Page" },
            wholeDocument:  { ja: "ドキュメント", en: "Document" }
        },
        checkbox: {
            includeParentPages: { ja: "親ページを含める", en: "Include Parent Pages" },
            inline:             { ja: "インライン", en: "Inline" },
            aboveLine:          { ja: "行の上", en: "Above Line" },
            custom:             { ja: "カスタム", en: "Custom" },
            textFrame:          { ja: "テキストフレーム", en: "Text Frames" },
            graphicFrame:       { ja: "グラフィックフレーム", en: "Graphic Frames" },
            selectReleased:     { ja: "解除したオブジェクトを選択", en: "Select Released Objects" }
        },
        tooltip: {
            selectedFrames: {
                ja: "選択中のアンカー付きオブジェクトと、選択したテキストフレームやテキストに含まれるものを解除します。",
                en: "Releases the selected anchored objects and those inside the selected text frames or text."
            },
            currentPage: {
                ja: "指定したページにあるアンカー付きオブジェクトを解除します。",
                en: "Releases the anchored objects on the specified page."
            },
            pageName: {
                ja: "ページ番号。初期値は表示中のページです。↑↓キーで前後のページ、shiftキー併用で10ページずつ移動します。",
                en: "Page number. Defaults to the page currently shown. Use the Up/Down arrow keys to move between pages, or add Shift to move 10 pages."
            },
            wholeDocument: {
                ja: "ドキュメント内のすべてのアンカー付きオブジェクトを解除します。",
                en: "Releases every anchored object in the document."
            },
            includeParentPages: {
                ja: "親ページ上のアンカー付きオブジェクトも解除します。［ドキュメント］のときだけ有効です。",
                en: "Also releases anchored objects on parent pages. Available only when Document is selected."
            },
            graphicFrame: {
                ja: "長方形・楕円・多角形のフレームです。画像の入っていないフレームも含みます。",
                en: "Rectangle, ellipse and polygon frames, including empty ones."
            },
            optionClick: {
                ja: "option＋クリックで、これ以外をオフにします。もう一度 option＋クリックで、すべてオンに戻します。",
                en: "Option-click to turn off all the others. Option-click again to turn them all back on."
            },
            selectReleased: {
                ja: "終了後、解除したオブジェクトを選択します。複数のスプレッドにまたがるときは、表示中のスプレッドのものだけを選択します。",
                en: "Selects the released objects when done. If they span several spreads, only those on the spread currently shown are selected."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument:       { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            pageNotFound:     { ja: "ページ「{name}」が見つかりません。", en: "Page \"{name}\" was not found." },
            noAnchoredObject: { ja: "条件に合うアンカー付きオブジェクトがありません。", en: "No anchored objects match the settings." },
            released:         { ja: "{count} 件のアンカー付きオブジェクトを解除しました。", en: "Released {count} anchored object(s)." },
            skipped:          { ja: "{count} 件は解除できなかったため飛ばしました。", en: "{count} object(s) could not be released and were skipped." }
        }
    };

    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /**
     * LABELS から現在の言語の文字列を返す
     * @param {object} labelSet - { ja, en } の組
     * @returns {string}
     */
    function getLabel(labelSet) {
        return labelSet[uiLang];
    }

    /**
     * 配列に値が含まれるか
     * @param {Array} candidates - 調べる配列
     * @param {*} searchValue - 探す値
     * @returns {boolean}
     */
    function arrayContains(candidates, searchValue) {
        for (var i = 0; i < candidates.length; i++) {
            if (candidates[i] === searchValue) return true;
        }
        return false;
    }

    /**
     * アンカー付きオブジェクト（親が文字のフレーム）かどうか
     * @param {object} pageItem - ページアイテム
     * @returns {boolean}
     */
    function isAnchoredFrame(pageItem) {
        var typeName = pageItem.constructor.name;
        if (typeName !== "TextFrame" && !arrayContains(GRAPHIC_FRAME_TYPES, typeName)) return false;
        return pageItem.parent.constructor.name === "Character";
    }

    /**
     * ページアイテムの配列からアンカー付きオブジェクトを抜き出して追加する
     * @param {Array} pageItems - 調べるページアイテム
     * @param {Array} anchoredFrames - 追加先
     * @returns {void}
     */
    function pushAnchoredFrames(pageItems, anchoredFrames) {
        for (var i = 0; i < pageItems.length; i++) {
            if (isAnchoredFrame(pageItems[i])) anchoredFrames.push(pageItems[i]);
        }
    }

    /**
     * 選択からアンカー付きオブジェクトを集める。テキストフレームやテキストの選択時は、その中に含まれるものを対象にする
     * @returns {Array}
     */
    function collectSelectedAnchoredFrames() {
        var anchoredFrames = [];
        var selectionItems = app.selection;
        for (var i = 0; i < selectionItems.length; i++) {
            var selectedItem = selectionItems[i];
            if (isAnchoredFrame(selectedItem)) {
                anchoredFrames.push(selectedItem);
            } else if (selectedItem.constructor.name === "TextFrame") {
                pushAnchoredFrames(selectedItem.texts[0].allPageItems, anchoredFrames);
            } else if (selectedItem.hasOwnProperty("parentTextFrames")) {
                pushAnchoredFrames(selectedItem.allPageItems, anchoredFrames);
            }
        }
        return anchoredFrames;
    }

    /**
     * 対象範囲に応じてアンカー付きオブジェクトを集める
     * @param {object} releaseOptions - ダイアログの設定
     * @param {Array} selectedAnchoredFrames - 選択から集めたもの
     * @returns {Array}
     */
    function collectTargetFrames(releaseOptions, selectedAnchoredFrames) {
        if (releaseOptions.scope === "selection") return selectedAnchoredFrames;

        var activeDoc = app.activeDocument;
        var docAnchoredFrames = [];
        pushAnchoredFrames(activeDoc.allPageItems, docAnchoredFrames);
        if (releaseOptions.scope === "document") return docAnchoredFrames;

        var targetPageId = releaseOptions.targetPage.id;
        var pageAnchoredFrames = [];
        for (var i = 0; i < docAnchoredFrames.length; i++) {
            var parentPage = docAnchoredFrames[i].parentPage; /* オーバーセットやペーストボード上は null / null when overset or on the pasteboard */
            if (parentPage && parentPage.id === targetPageId) pageAnchoredFrames.push(docAnchoredFrames[i]);
        }
        return pageAnchoredFrames;
    }

    /**
     * 親ページ上にあるか
     * @param {PageItem} pageItem - ページアイテム
     * @returns {boolean}
     */
    function isOnParentPage(pageItem) {
        var parentPage = pageItem.parentPage;
        return parentPage !== null && parentPage.parent.constructor.name === "MasterSpread";
    }

    /**
     * アンカーの種類・フレームの種類・親ページの設定で絞り込む
     * @param {Array} anchoredFrames - アンカー付きオブジェクト
     * @param {object} releaseOptions - ダイアログの設定
     * @returns {Array}
     */
    function filterByOptions(anchoredFrames, releaseOptions) {
        var matchedFrames = [];
        for (var i = 0; i < anchoredFrames.length; i++) {
            var anchoredFrame = anchoredFrames[i];
            var isTextFrame = (anchoredFrame.constructor.name === "TextFrame");
            if (isTextFrame ? !releaseOptions.includeTextFrames : !releaseOptions.includeGraphicFrames) continue;
            if (!arrayContains(releaseOptions.anchorPositions, anchoredFrame.anchoredObjectSettings.anchoredPosition)) continue;
            if (releaseOptions.scope === "document" && !releaseOptions.includeParentPages && isOnParentPage(anchoredFrame)) continue;
            matchedFrames.push(anchoredFrame);
        }
        return matchedFrames;
    }

    /**
     * アンカー付きオブジェクトを1つずつ解除する（位置はそのまま、親はスプレッドになる）
     * @param {Array} anchoredFrames - 対象のアンカー付きオブジェクト
     * @returns {Array} 解除できたオブジェクト
     */
    function releaseEachFrame(anchoredFrames) {
        var releasedFrames = [];
        for (var i = 0; i < anchoredFrames.length; i++) {
            try {
                anchoredFrames[i].anchoredObjectSettings.releaseAnchoredObject();
                releasedFrames.push(anchoredFrames[i]);
            } catch (e) {
                /* ロック・ロックされたレイヤー上などで解除できないものは飛ばす / Skip items that cannot be released */
            }
        }
        return releasedFrames;
    }

    /**
     * 解除したオブジェクトを選択する。スプレッドをまたぐときは表示中のスプレッドのものだけにする
     * @param {Array} releasedFrames - 解除したオブジェクト
     * @returns {void}
     */
    function selectReleasedFrames(releasedFrames) {
        var activeSpread = app.activeDocument.layoutWindows[0].activeSpread;
        var framesOnSpread = [];
        for (var i = 0; i < releasedFrames.length; i++) {
            if (releasedFrames[i].parent.id === activeSpread.id) framesOnSpread.push(releasedFrames[i]);
        }
        try {
            app.select(framesOnSpread.length > 0 ? framesOnSpread : NothingEnum.NOTHING);
        } catch (e) {
            /* ロックされたものが混ざると選択できない / Selection fails when locked items are included */
        }
    }

    /**
     * ページ番号欄の値からページを探す。見つからなければ null
     * @param {string} pageName - ページ番号
     * @returns {Page|null}
     */
    function findPageByName(pageName) {
        var foundPage = app.activeDocument.pages.itemByName(pageName);
        return foundPage.isValid ? foundPage : null;
    }

    /**
     * ページ番号欄で↑↓キーを押したら、前後のページへ移動する（shift 併用で10ページ）
     * @param {EditText} pageNameField - ページ番号欄
     * @param {KeyboardEvent} event - キーイベント
     * @returns {void}
     */
    function stepPageByArrowKey(pageNameField, event) {
        if (event.keyName != "Up" && event.keyName != "Down") return;
        var docPages = app.activeDocument.pages;
        var typedPage = findPageByName(pageNameField.text);
        var currentIndex = typedPage ? typedPage.documentOffset : 0;
        var stepAmount = ScriptUI.environment.keyboardState.shiftKey ? 10 : 1;
        var nextIndex = currentIndex + (event.keyName == "Up" ? stepAmount : -stepAmount);
        nextIndex = Math.max(0, Math.min(docPages.length - 1, nextIndex));
        pageNameField.text = docPages[nextIndex].name;
        pageNameField.notify("onChanging");
        event.preventDefault();
    }

    /**
     * ページ番号欄の値から対象ページを決める。見つからなければメッセージを出して null
     * @param {string} pageName - ページ番号欄の値
     * @param {Page} activePage - 表示中のページ（親ページのこともある）
     * @returns {Page|null}
     */
    function resolveTargetPage(pageName, activePage) {
        if (pageName === activePage.name) return activePage; /* 親ページは pages から引けないので表示中のものを使う / Parent pages are not in pages */
        var foundPage = findPageByName(pageName);
        if (foundPage === null) alert(getLabel(LABELS.alert.pageNotFound).replace("{name}", pageName));
        return foundPage;
    }

    /**
     * 対象パネルを作る
     * @param {Window} optionsDialog - ダイアログ
     * @param {boolean} hasSelection - 選択にアンカー付きオブジェクトがあるか
     * @param {string} initialPageName - ページ番号欄の初期値
     * @returns {object} 各コントロール
     */
    function buildTargetPanel(optionsDialog, hasSelection, initialPageName) {
        var targetPanel = optionsDialog.add("panel", undefined, getLabel(LABELS.panel.target));
        setupPanel(targetPanel);
        var selectedFramesRadio = targetPanel.add("radiobutton", undefined, getLabel(LABELS.radio.selectedFrames));

        /* ページ番号欄と並べるため別グループ。排他は selectScope() で取る / Separate group for the page field; exclusivity handled in selectScope() */
        var currentPageRow = targetPanel.add("group");
        currentPageRow.orientation = "row";
        currentPageRow.alignChildren = ["left", "center"];
        currentPageRow.spacing = 6;
        var currentPageRadio = currentPageRow.add("radiobutton", undefined, getLabel(LABELS.radio.currentPage));
        var pageNameField = currentPageRow.add("edittext", undefined, initialPageName);
        pageNameField.characters = 5;

        var wholeDocumentRadio = targetPanel.add("radiobutton", undefined, getLabel(LABELS.radio.wholeDocument));

        var parentPagesGroup = targetPanel.add("group");
        parentPagesGroup.margins = [18, 0, 0, 0]; /* ［ドキュメント］の下に字下げ / Indent under Document */
        var includeParentPagesCheckbox = parentPagesGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.includeParentPages));

        selectedFramesRadio.helpTip = getLabel(LABELS.tooltip.selectedFrames);
        currentPageRadio.helpTip = getLabel(LABELS.tooltip.currentPage);
        pageNameField.helpTip = getLabel(LABELS.tooltip.pageName);
        wholeDocumentRadio.helpTip = getLabel(LABELS.tooltip.wholeDocument);
        includeParentPagesCheckbox.helpTip = getLabel(LABELS.tooltip.includeParentPages);

        selectedFramesRadio.enabled = hasSelection;

        var currentScope = hasSelection ? "selection" : "page";
        function selectScope(targetScope) {
            currentScope = targetScope;
            selectedFramesRadio.value = (targetScope === "selection");
            currentPageRadio.value = (targetScope === "page");
            wholeDocumentRadio.value = (targetScope === "document");
            includeParentPagesCheckbox.enabled = (targetScope === "document");
        }
        function scopeSelector(targetScope) {
            return function () { selectScope(targetScope); };
        }
        selectScope(currentScope);
        selectedFramesRadio.onClick = scopeSelector("selection");
        currentPageRadio.onClick = scopeSelector("page");
        wholeDocumentRadio.onClick = scopeSelector("document");
        pageNameField.onChanging = scopeSelector("page"); /* 入力したら［現在のページ］に切り替え / Typing switches to Current Page */
        pageNameField.addEventListener("keydown", function (event) {
            stepPageByArrowKey(pageNameField, event);
        });

        return {
            getTargetScope: function () { return currentScope; },
            pageNameField: pageNameField,
            includeParentPagesCheckbox: includeParentPagesCheckbox
        };
    }

    /**
     * チェックボックスを横に並べたパネルを作る。初期値はすべてオン
     * @param {Window} optionsDialog - ダイアログ
     * @param {object} panelLabel - パネル名の LABELS
     * @param {Array} checkboxLabels - チェックボックスの LABELS
     * @returns {Array} チェックボックス
     */
    function buildCheckboxRowPanel(optionsDialog, panelLabel, checkboxLabels) {
        var checkboxPanel = optionsDialog.add("panel", undefined, getLabel(panelLabel));
        setupPanel(checkboxPanel);
        checkboxPanel.orientation = "row";
        var panelCheckboxes = [];
        for (var i = 0; i < checkboxLabels.length; i++) {
            var newCheckbox = checkboxPanel.add("checkbox", undefined, getLabel(checkboxLabels[i]));
            newCheckbox.value = true;
            panelCheckboxes.push(newCheckbox);
        }
        return panelCheckboxes;
    }

    /**
     * option＋クリックで「これだけオン」、もう一度で「すべてオン」にする
     * @param {Array} panelCheckboxes - 同じパネルのチェックボックス
     * @param {Function} afterClick - クリック後に呼ぶ処理
     * @returns {void}
     */
    function enableOptionClickSolo(panelCheckboxes, afterClick) {
        var optionClickTip = getLabel(LABELS.tooltip.optionClick);
        for (var i = 0; i < panelCheckboxes.length; i++) {
            var panelCheckbox = panelCheckboxes[i];
            panelCheckbox.helpTip = panelCheckbox.helpTip ? panelCheckbox.helpTip + "\n" + optionClickTip : optionClickTip;
            panelCheckbox.onClick = soloHandler(panelCheckbox);
        }

        function soloHandler(clickedCheckbox) {
            return function () {
                if (ScriptUI.environment.keyboardState.altKey) {
                    /* クリックで反転済みなので、オフになった＝もともと単独オンだった / Value is already toggled: now off means it was the only one on */
                    var othersOff = true;
                    for (var j = 0; j < panelCheckboxes.length; j++) {
                        if (panelCheckboxes[j] !== clickedCheckbox && panelCheckboxes[j].value) othersOff = false;
                    }
                    var selectAll = othersOff && !clickedCheckbox.value;
                    for (var k = 0; k < panelCheckboxes.length; k++) {
                        panelCheckboxes[k].value = selectAll || panelCheckboxes[k] === clickedCheckbox;
                    }
                }
                afterClick();
            };
        }
    }

    /**
     * 1つ以上オンになっているか
     * @param {Array} panelCheckboxes - チェックボックス
     * @returns {boolean}
     */
    function hasAnyChecked(panelCheckboxes) {
        for (var i = 0; i < panelCheckboxes.length; i++) {
            if (panelCheckboxes[i].value) return true;
        }
        return false;
    }

    /**
     * オンになっているチェックボックスに対応する値を集める
     * @param {Array} panelCheckboxes - チェックボックス
     * @param {Array} checkboxValues - 各チェックボックスに対応する値
     * @returns {Array}
     */
    function collectCheckedValues(panelCheckboxes, checkboxValues) {
        var checkedValues = [];
        for (var i = 0; i < panelCheckboxes.length; i++) {
            if (panelCheckboxes[i].value) checkedValues.push(checkboxValues[i]);
        }
        return checkedValues;
    }

    /**
     * パネルの共通設定を適用する
     * @param {Panel} targetPanel - 対象パネル
     * @returns {void}
     */
    function setupPanel(targetPanel) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = "left";
        targetPanel.margins = [15, 20, 15, 10];
    }

    /**
     * 設定ダイアログを表示する
     * @param {boolean} hasSelection - 選択にアンカー付きオブジェクトがあるか
     * @returns {object|null} 設定。キャンセル時は null
     */
    function showOptionsDialog(hasSelection) {
        var optionsDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        optionsDialog.orientation = "column";
        optionsDialog.alignChildren = "fill";

        /* 表示中のページ。ストーリーエディターが前面でも取れるようにレイアウトウィンドウから / Use the layout window so it works with the story editor in front */
        var activePage = app.activeDocument.layoutWindows[0].activePage;
        var targetControls = buildTargetPanel(optionsDialog, hasSelection, activePage.name);
        var anchorTypeCheckboxes = buildCheckboxRowPanel(optionsDialog, LABELS.panel.anchorType,
            [LABELS.checkbox.inline, LABELS.checkbox.aboveLine, LABELS.checkbox.custom]);
        var frameTypeCheckboxes = buildCheckboxRowPanel(optionsDialog, LABELS.panel.frameType,
            [LABELS.checkbox.textFrame, LABELS.checkbox.graphicFrame]);
        frameTypeCheckboxes[1].helpTip = getLabel(LABELS.tooltip.graphicFrame);
        var afterReleaseCheckboxes = buildCheckboxRowPanel(optionsDialog, LABELS.panel.afterRelease, [LABELS.checkbox.selectReleased]);
        afterReleaseCheckboxes[0].helpTip = getLabel(LABELS.tooltip.selectReleased);

        var btnRowGroup = optionsDialog.add("group");
        btnRowGroup.alignment = "right";
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        /* 種類が1つも選ばれていなければ実行できない / Require at least one type in each group */
        function updateOKEnabled() {
            btnOK.enabled = hasAnyChecked(anchorTypeCheckboxes) && hasAnyChecked(frameTypeCheckboxes);
        }
        enableOptionClickSolo(anchorTypeCheckboxes, updateOKEnabled);
        enableOptionClickSolo(frameTypeCheckboxes, updateOKEnabled);

        /* ページが見つからなければダイアログを閉じない / Keep the dialog open when the page is not found */
        var targetPage = activePage;
        btnOK.onClick = function () {
            if (targetControls.getTargetScope() === "page") {
                targetPage = resolveTargetPage(targetControls.pageNameField.text, activePage);
                if (targetPage === null) return;
            }
            optionsDialog.close(1);
        };

        if (optionsDialog.show() !== 1) return null;

        return {
            scope: targetControls.getTargetScope(),
            targetPage: targetPage,
            includeParentPages: targetControls.includeParentPagesCheckbox.value,
            anchorPositions: collectCheckedValues(anchorTypeCheckboxes,
                [AnchorPosition.INLINE_POSITION, AnchorPosition.ABOVE_LINE, AnchorPosition.ANCHORED]),
            includeTextFrames: frameTypeCheckboxes[0].value,
            includeGraphicFrames: frameTypeCheckboxes[1].value,
            selectReleased: afterReleaseCheckboxes[0].value
        };
    }

    /**
     * 完了メッセージを作る
     * @param {number} releasedCount - 解除した件数
     * @param {number} skippedCount - 飛ばした件数
     * @returns {string}
     */
    function buildResultMessage(releasedCount, skippedCount) {
        var resultMessage = getLabel(LABELS.alert.released).replace("{count}", releasedCount);
        if (skippedCount > 0) resultMessage += "\n" + getLabel(LABELS.alert.skipped).replace("{count}", skippedCount);
        return resultMessage;
    }

    /**
     * 設定ダイアログを出し、条件に合うアンカー付きオブジェクトを解除する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var selectedAnchoredFrames = collectSelectedAnchoredFrames();

        var releaseOptions = showOptionsDialog(selectedAnchoredFrames.length > 0);
        if (releaseOptions === null) return;

        var anchoredFrames = filterByOptions(collectTargetFrames(releaseOptions, selectedAnchoredFrames), releaseOptions);
        if (anchoredFrames.length === 0) {
            alert(getLabel(LABELS.alert.noAnchoredObject));
            return;
        }

        var releasedFrames = releaseEachFrame(anchoredFrames);
        if (releaseOptions.selectReleased) selectReleasedFrames(releasedFrames);
        alert(buildResultMessage(releasedFrames.length, anchoredFrames.length - releasedFrames.length));
    }

    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.dialog.title));
})();
