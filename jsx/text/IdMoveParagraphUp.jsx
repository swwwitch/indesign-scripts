#target indesign
#targetengine "IdMoveParagraph"

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdMoveParagraphUp
Meta-Default-language=ja
Name=段落をひとつ上の段落と入れ替える
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.0.1
Release-Date=2026-09-30
Tags=text,テキスト
X-com.swwwitch-Script-Name=IdMoveParagraphUp
X-com.swwwitch-Updated=2026-09-30
X-com.swwwitch-README-JA=https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdMoveParagraphUp.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdMoveParagraphUp.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdMoveParagraphUp.jsx
Description-BEGIN
[ja]
カーソルのある段落を、ひとつ上の段落と入れ替えます。
sky-chaser-high 氏の moveLineUp.jsx（Illustrator）と、それを段落単位にした Illustrator 版 MoveParagraphUp.jsx から着想を得て、InDesign 向けに書き起こしたものです。
[en]
Swaps the paragraph containing the cursor with the paragraph above it.
Written for InDesign, inspired by moveLineUp.jsx by sky-chaser-high and its paragraph-based Illustrator variant MoveParagraphUp.jsx.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読んで SCRIPT_* に入れる（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; it is read into SCRIPT_* at run time */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME           = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名 / script name */
var SCRIPT_VERSION        = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : ""; /* バージョン / version */
var SCRIPT_AUTHOR         = getScriptMetaValue("Author");                       /* 作者 / author */
var SCRIPT_RELEASED       = getScriptMetaValue("Release-Date");                 /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED        = getScriptMetaValue("X-com.swwwitch-Updated");       /* 更新日 / last updated */
var SCRIPT_DESCRIPTION_JA = getScriptMetaValue("Name");                         /* README の見出し / README title (Japanese) */
var SCRIPT_README_JA      = getScriptMetaValue("X-com.swwwitch-README-JA");     /* README（日本語） */
var SCRIPT_README_EN      = getScriptMetaValue("X-com.swwwitch-README-EN");     /* README (English) */
var SCRIPT_ARTICLE_URL    = getScriptMetaValue("X-com.swwwitch-Article-URL");   /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL     = getScriptMetaValue("X-com.swwwitch-Source-URL");    /* 最新版のソース / latest source */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * 着想元 / Inspired by
 * @author sky-chaser-high
 * @discussion https://github.com/sky-chaser-high/adobe-illustrator-scripts/blob/main/README_ja.md#%E8%A1%8C%E3%82%92%E4%B8%8A--%E4%B8%8B%E3%81%B8%E7%A7%BB%E5%8B%95
 */

(function () {

    /* テキストとして扱う選択の型 / Selection types treated as text */
    var TEXT_SELECTION_TYPES = {
        InsertionPoint: true, Character: true, Word: true, Line: true,
        TextStyleRange: true, Paragraph: true, TextColumn: true, Text: true
    };

    /**
     * 選択がテキスト（キャレット・文字範囲）かどうかを判定する
     * @param {object} selectionItem 選択オブジェクト
     * @returns {boolean} テキスト上の選択なら true
     */
    function isTextSelection(selectionItem) {
        if (selectionItem == null) return false;
        return TEXT_SELECTION_TYPES[selectionItem.constructor.name] === true;
    }

    /**
     * 選択中のテキストを確かめ、1回の取り消しで戻せるように段落を移動する
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) return;
        if (app.selection.length !== 1) return;

        /* 文字ツールでテキストを選択、またはテキスト内にカーソルがあるときだけ実行する / Run only when text is selected or the caret is inside text */
        var selectedText = app.selection[0];
        if (!isTextSelection(selectedText)) return;

        app.doScript(function () {
            moveCurrentParagraphUp(selectedText);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.FAST_ENTIRE_SCRIPT, SCRIPT_NAME);
    }

    /**
     * カーソルのある段落をひとつ上の段落と入れ替え、選択位置を追従させる。
     * @param {Text} selectedText - 選択中のテキスト（キャレットのみの場合を含む）
     * @returns {void}
     */
    function moveCurrentParagraphUp(selectedText) {
        var caret = selectedText.insertionPoints[0];
        var paragraph = selectedText.paragraphs[0];
        var previousParagraph = getPreviousParagraph(paragraph);

        /* 先頭段落は上へ動かせない / The first paragraph cannot move up */
        if (!previousParagraph) return;

        /* 段落内での選択位置と長さを控える / Remember the selection relative to the paragraph */
        var offsetInParagraph = caret.index - paragraph.insertionPoints[0].index;
        var selectionLength = selectedText.characters.length;

        /* 入れ替え後に移動した段落を引き直すための足場（位置で解決される） */
        /* Anchor resolved by position to re-find the moved paragraph after the swap */
        var anchor = previousParagraph.insertionPoints[0];

        swapWithPreviousParagraph(paragraph, previousParagraph);

        restoreSelection(anchor.paragraphs[0], offsetInParagraph, selectionLength);
    }

    /**
     * 同じテキストの流れ（ストーリーまたはセル）の中で、前の段落を返す。
     * @param {Paragraph} paragraph - 基準の段落
     * @returns {Paragraph|null} 前の段落。無ければ null
     */
    function getPreviousParagraph(paragraph) {
        try {
            var previousParagraph = paragraph.parent.paragraphs.previousItem(paragraph);
            if (previousParagraph && previousParagraph.isValid) return previousParagraph;
        } catch (e) {}
        return null;
    }

    /**
     * 段落が段落区切りで終わっているかどうかを判定する。
     * @param {Paragraph} paragraph - 判定する段落
     * @returns {boolean} 末尾が段落区切りなら true
     */
    function endsWithParagraphBreak(paragraph) {
        return /\r$/.test(paragraph.contents);
    }

    /**
     * 指定した段落を、ひとつ上の段落と入れ替える。
     * 段落を上の段落の前へ move() するので、文字・段落の書式ごと移る。
     * 動かす段落が末尾で改行を持たないとき（ストーリーやセルの最終段落）は、
     * 一時的に改行を足してから入れ替え、最後に余った改行を取り除く。
     * @param {Paragraph} paragraph - 上へ移動する段落
     * @param {Paragraph} previousParagraph - その前の段落
     * @returns {void}
     */
    function swapWithPreviousParagraph(paragraph, previousParagraph) {
        var isLastParagraph = !endsWithParagraphBreak(paragraph);
        if (isLastParagraph) paragraph.insertionPoints[-1].contents = "\r";

        /* 参照は位置で解決されるため、move() の後で paragraph は使わない */
        /* Text references resolve by position; do not reuse paragraph after move() */
        var container = paragraph.parent;
        paragraph.move(LocationOptions.BEFORE, previousParagraph.insertionPoints[0]);

        /* 足した改行の分だけ末尾に空段落が残るので、その直前の改行を削除する */
        /* Remove the break that now leaves an empty paragraph at the end */
        if (isLastParagraph) container.characters[-1].remove();
    }

    /**
     * 入れ替え後の段落の中に、元の選択（キャレットまたは範囲）を復帰させる。
     * 段落の外へはみ出す範囲は、キャレットだけを戻す。
     * @param {Paragraph} paragraph - 復帰先の段落
     * @param {number} offsetInParagraph - 段落先頭からの選択開始位置
     * @param {number} selectionLength - 選択していた文字数（キャレットなら 0）
     * @returns {void}
     */
    function restoreSelection(paragraph, offsetInParagraph, selectionLength) {
        /* 改行の後ろにはキャレットを置かない / Keep the caret before the paragraph break */
        var lastOffset = paragraph.characters.length - (endsWithParagraphBreak(paragraph) ? 1 : 0);
        var startOffset = Math.min(offsetInParagraph, lastOffset);

        if (selectionLength > 0 && startOffset + selectionLength <= paragraph.characters.length) {
            app.select(paragraph.characters.itemByRange(startOffset, startOffset + selectionLength - 1));
        } else {
            app.select(paragraph.insertionPoints[startOffset]);
        }
    }

    main();

})();
