#target indesign

/*

### 概要

選択したフレーム、指定したページ、またはドキュメント内のアンカー付きオブジェクトを、アンカーやフレームの種類で絞り込んで解除します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdReleaseAnchoredObjects.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ne3ee16f466bf

### Overview

Releases the anchored objects in the selected frames, on a chosen page, or throughout the document, filtered by anchor type and frame type.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdReleaseAnchoredObjects.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdReleaseAnchoredObjects";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdReleaseAnchoredObjects.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdReleaseAnchoredObjects.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ne3ee16f466bf"; /* 紹介記事 / article URL */

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

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

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

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

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
     * グラフィックフレーム（長方形・楕円・多角形）の種別かどうか
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} グラフィックフレームなら true
     */
    function isGraphicFrameType(pageItem) {
        return arrayContains(GRAPHIC_FRAME_TYPES, pageItem.constructor.name);
    }

    /**
     * フレームがテキストにアンカーされているか（親が文字なら、インライン・行の上・カスタムのいずれか）
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} アンカー付きなら true
     */
    function isAnchoredFrame(pageItem) {
        return pageItem.parent.constructor.name === "Character";
    }

    /**
     * 解除の対象になるアンカー付きオブジェクト（テキストフレームかグラフィックフレーム）かどうか
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} 対象なら true
     */
    function isReleasableFrame(pageItem) {
        if (pageItem.constructor.name !== "TextFrame" && !isGraphicFrameType(pageItem)) return false;
        return isAnchoredFrame(pageItem);
    }

    /**
     * ページアイテムの配列からアンカー付きオブジェクトを抜き出して追加する
     * @param {Array} pageItems - 調べるページアイテム
     * @param {Array} anchoredFrames - 追加先
     * @returns {void}
     */
    function pushAnchoredFrames(pageItems, anchoredFrames) {
        for (var i = 0; i < pageItems.length; i++) {
            if (isReleasableFrame(pageItems[i])) anchoredFrames.push(pageItems[i]);
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
            if (isReleasableFrame(selectedItem)) {
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
        if (foundPage === null) alert(getLabel(LABELS.alert.pageNotFound, { name: pageName }));
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
        setupPanel(targetPanel, 6);
        var selectedFramesRadio = targetPanel.add("radiobutton", undefined, getLabel(LABELS.radio.selectedFrames));

        /* ページ番号欄と並べるため別グループ。排他は selectScope() で取る / Separate group for the page field; exclusivity handled in selectScope() */
        var currentPageRow = targetPanel.add("group");
        setupRow(currentPageRow, "left", 6);
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
        checkboxPanel.alignChildren = ["left", "center"]; /* 横並びなので伸ばさない / Row layout: do not stretch the checkboxes */
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
     * 設定ダイアログを表示する
     * @param {boolean} hasSelection - 選択にアンカー付きオブジェクトがあるか
     * @returns {object|null} 設定。キャンセル時は null
     */
    function showOptionsDialog(hasSelection) {
        var optionsDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(optionsDialog);

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

        var buttonRow = addButtonRow(optionsDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

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
        var resultMessage = getLabel(LABELS.alert.released, { count: releasedCount });
        if (skippedCount > 0) resultMessage += "\n" + getLabel(LABELS.alert.skipped, { count: skippedCount });
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
