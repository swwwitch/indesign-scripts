#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdInsertPageCrossReference
Name=カーソル位置にアンカー付き長方形とテキストアンカーを作成
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=2.1.0
Release-Date=2026-11-06
Description-BEGIN
カーソルを置いた位置に、アンカー付き（カスタム）でプリントしない小さな長方形と、その直前のテキストアンカーを作成します。
長方形とテキストアンカーは同じ名前で対応づけられ、ハイパーリンクや相互参照の行き先として使えます。

Creates a small non-printing anchored (custom) rectangle at the text cursor, plus a text anchor right before it.
The rectangle and the text anchor share the same name and can serve as a target for hyperlinks and cross-references.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

カーソルを置いた位置に、アンカー付き（カスタム）でプリントしない小さな長方形と、その直前のテキストアンカーを作成します。
長方形とテキストアンカーは同じ名前で対応づけられ、ハイパーリンクや相互参照の行き先として使えます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdInsertAnchoredRectangle.md

### Overview

Creates a small non-printing anchored (custom) rectangle at the text cursor, plus a text anchor right before it.
The rectangle and the text anchor share the same name and can serve as a target for hyperlinks and cross-references.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdInsertAnchoredRectangle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdInsertAnchoredRectangle";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-06";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "カーソル位置にアンカー付き長方形とテキストアンカーを作成"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdInsertAnchoredRectangle.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdInsertAnchoredRectangle.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdInsertAnchoredRectangle.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 基本設定 / Settings
    // =========================================

    /* 長方形の幅・高さ（pt）/ Rectangle width and height (pt) */
    var RECT_WIDTH  = 12;
    var RECT_HEIGHT = 12;

    /* 長方形の線幅（pt）/ Rectangle stroke weight (pt) */
    var RECT_STROKE_WEIGHT = 0.5;

    /* テキストアンカー名・スクリプトラベルの接頭辞 / Prefix of the text anchor name and script label */
    var ANCHOR_NAME_PREFIX = "anchored_rectangle";

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
        undoName: { ja: "アンカー付き長方形とテキストアンカーを作成", en: "Create Anchored Rectangle and Text Anchor" },
        alert: {
            noDocument:     { ja: "InDesignドキュメントを開いてください。", en: "Open an InDesign document." },
            noTextSelected: {
                ja: "参照先にしたいテキスト内へカーソルを置くか、テキストを選択してから実行してください。",
                en: "Place the text cursor in the text you want to target, or select text, and run the script again."
            },
            notOnPage: {
                ja: "ペーストボード上またはオーバーセット状態のテキストには、このスクリプトを使用できません。",
                en: "This script cannot be used on text on the pasteboard or in overset text."
            },
            noAnchorCharacter: {
                ja: "アンカー付きオブジェクトのマーカー文字を取得できませんでした。",
                en: "Could not get the marker character of the anchored object."
            },
            created: {
                ja: "アンカー付き長方形とテキストアンカーを作成しました。\r\rテキストアンカー名 :\r%1",
                en: "Created an anchored rectangle and a text anchor.\r\rText anchor name:\r%1"
            }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }
    if (app.selection.length !== 1 || !isTextSelection(app.selection[0])) {
        alert(getLabel(LABELS.alert.noTextSelected));
        return;
    }

    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undoName));

    /**
     * カーソル位置にアンカー付き長方形を入れ、その直前にテキストアンカーを作る
     * @returns {void}
     */
    function main() {
        var doc = app.activeDocument;
        var insertionPoint = app.selection[0].insertionPoints[0];

        var page = getPageOf(insertionPoint);
        if (!page) {
            alert(getLabel(LABELS.alert.notOnPage));
            return;
        }

        var anchorName = makeUniqueDestinationName(doc, ANCHOR_NAME_PREFIX);
        var rectangle = createRectangle(doc, page, anchorName);

        rectangle.anchoredObjectSettings.insertAnchoredObject(insertionPoint, AnchorPosition.ANCHORED);

        /* アンカー化した長方形の親は、オブジェクト置換文字（Character）/ An anchored rectangle's parent is its object replacement Character */
        var anchorCharacter = rectangle.parent;
        if (!anchorCharacter || !anchorCharacter.isValid || anchorCharacter.constructor.name !== "Character") {
            alert(getLabel(LABELS.alert.noAnchorCharacter));
            return;
        }

        /* マーカー文字そのものではなく、直前の挿入点をテキストアンカーにする / Use the insertion point before the marker, not the marker itself */
        var destination = doc.hyperlinkTextDestinations.add(anchorCharacter.insertionPoints[0], { name: anchorName });

        app.select(rectangle);
        alert(getLabel(LABELS.alert.created, [destination.name]));
    }

    /**
     * 選択がテキスト（カーソル位置を含む）かどうか
     * @param {Object} selectedItem - app.selection の要素
     * @returns {boolean} テキストなら true（テキストフレームそのものの選択は false）
     */
    function isTextSelection(selectedItem) {
        return selectedItem != null && selectedItem.isValid && selectedItem.hasOwnProperty("baseline");
    }

    /**
     * 挿入点が載っているページを返す
     * @param {InsertionPoint} insertionPoint - 対象の挿入点
     * @returns {Page|null} ページ。オーバーセットやペーストボード上なら null
     */
    function getPageOf(insertionPoint) {
        var frames = insertionPoint.parentTextFrames;
        if (frames.length === 0) return null;
        var page = frames[0].parentPage;
        return (page && page.isValid) ? page : null;
    }

    /**
     * 塗りなしの長方形をページ上に作る（位置はアンカー化で変わるので、大きさだけ決める）
     * @param {Document} doc - 対象ドキュメント
     * @param {Page} page - 作成するページ
     * @param {string} labelName - スクリプトラベルに入れる名前
     * @returns {Rectangle} 作成した長方形
     */
    function createRectangle(doc, page, labelName) {
        var rectangle = page.rectangles.add();
        rectangle.geometricBounds = [0, 0, RECT_HEIGHT, RECT_WIDTH];
        rectangle.strokeWeight = RECT_STROKE_WEIGHT;
        /* 目印なのでプリント・書き出しには出さない / A marker only, so keep it out of print and export */
        rectangle.nonprinting = true;
        var noneSwatch = doc.swatches.itemByName("None");
        if (noneSwatch.isValid) rectangle.fillColor = noneSwatch;
        /* テキストアンカーと同じ名前で追えるようにする / Lets the rectangle be matched with its text anchor by name */
        rectangle.label = labelName;
        return rectangle;
    }

    /**
     * ドキュメント内で重複しないテキストアンカー名を作る
     * @param {Document} doc - 対象ドキュメント
     * @param {string} prefix - 名前の接頭辞
     * @returns {string} テキストアンカー名
     */
    function makeUniqueDestinationName(doc, prefix) {
        var baseName = prefix + "_" + new Date().getTime();
        var candidate = baseName;
        for (var i = 1; doc.hyperlinkTextDestinations.itemByName(candidate).isValid; i++) {
            candidate = baseName + "_" + i;
        }
        return candidate;
    }

})();
