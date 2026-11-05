#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdInsertPageCrossReference
Name=段落のページ番号を「p.番号」の相互参照で挿入
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.0.0
Release-Date=2026-11-06
Description-BEGIN
1つ目のストーリーの3段落目が載っているページの番号を、2つ目のストーリーの末尾に「p.12」の形の相互参照で挿入します。
番号は固定の文字ではないので、ページが動いても相互参照の更新で追従します。

Inserts the page number of the third paragraph of the first story at the end of the second story, as a cross-reference such as "p.12".
The number is not fixed text, so it follows page changes when cross-references are updated.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

1つ目のストーリーの3段落目が載っているページの番号を、2つ目のストーリーの末尾に「p.12」の形の相互参照で挿入します。
番号は固定の文字ではないので、ページが動いても相互参照の更新で追従します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdInsertPageCrossReference.md

### Overview

Inserts the page number of the third paragraph of the first story at the end of the second story, as a cross-reference such as "p.12".
The number is not fixed text, so it follows page changes when cross-references are updated.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdInsertPageCrossReference.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdInsertPageCrossReference";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-06";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "段落のページ番号を「p.番号」の相互参照で挿入"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdInsertPageCrossReference.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdInsertPageCrossReference.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdInsertPageCrossReference.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 基本設定 / Settings
    // =========================================

    /* 挿入する相互参照の書式名 / Name of the cross-reference format to insert */
    var XREF_FORMAT_NAME = "p.ページ番号";

    /* ページ番号の前に付ける文字列 / Text placed before the page number */
    var XREF_PREFIX = "p.";

    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "ページ参照の挿入");

    /**
     * stories[0] の3段落目を参照先にして、stories[1] の末尾へ「p.番号」の相互参照を挿入する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) { alert("ドキュメントが開いていません"); return; }
        var doc = app.activeDocument;
        if (doc.stories.length < 2) { alert("ストーリーが2つ以上必要です"); return; }

        var destPara = doc.stories[0].paragraphs[2];
        if (!destPara.isValid) { alert("参照先の段落がありません"); return; }

        var format = getPageNumberFormat(doc);
        var dest = doc.paragraphDestinations.add(destPara);
        var src = doc.crossReferenceSources.add(doc.stories[1].insertionPoints[-1], format);
        doc.hyperlinks.add(src, dest, { visible: false });
    }

    /**
     * 「p.番号」の相互参照書式を返す。無ければ作る
     * @param {Document} doc - 対象ドキュメント
     * @returns {CrossReferenceFormat} 相互参照書式
     */
    function getPageNumberFormat(doc) {
        var format = doc.crossReferenceFormats.itemByName(XREF_FORMAT_NAME);
        if (format.isValid) return format;

        format = doc.crossReferenceFormats.add({ name: XREF_FORMAT_NAME });
        format.buildingBlocks.add(BuildingBlockTypes.CUSTOM_STRING_BUILDING_BLOCK, { customText: XREF_PREFIX });
        format.buildingBlocks.add(BuildingBlockTypes.PAGE_NUMBER_BUILDING_BLOCK);
        return format;
    }

})();
