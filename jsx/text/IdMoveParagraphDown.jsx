#target indesign
#targetengine "IdMoveParagraph"

/*

### 概要

カーソルのある段落を、ひとつ下の段落と入れ替えます。
sky-chaser-high 氏の moveLineDown.jsx（Illustrator）と、それを段落単位にした Illustrator 版 MoveParagraphDown.jsx から着想を得て、InDesign 向けに書き起こしたものです。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdMoveParagraphDown.md

### Overview

Swaps the paragraph containing the cursor with the paragraph below it.
Written for InDesign, inspired by moveLineDown.jsx by sky-chaser-high and its paragraph-based Illustrator variant MoveParagraphDown.jsx.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdMoveParagraphDown.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdMoveParagraphDown";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdMoveParagraphDown.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdMoveParagraphDown.md"; /* README (English) */

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
            moveCurrentParagraphDown(selectedText);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.FAST_ENTIRE_SCRIPT, SCRIPT_NAME);
    }

    /**
     * カーソルのある段落をひとつ下の段落と入れ替え、選択位置を追従させる。
     * @param {Text} selectedText - 選択中のテキスト（キャレットのみの場合を含む）
     * @returns {void}
     */
    function moveCurrentParagraphDown(selectedText) {
        var caret = selectedText.insertionPoints[0];
        var paragraph = selectedText.paragraphs[0];
        var nextParagraph = getNextParagraph(paragraph);

        /* 最終段落は下へ動かせない / The last paragraph cannot move down */
        if (!nextParagraph) return;

        /* 段落内での選択位置と長さを控える / Remember the selection relative to the paragraph */
        var offsetInParagraph = caret.index - paragraph.insertionPoints[0].index;
        var selectionLength = selectedText.characters.length;

        /* 入れ替え後に先頭の段落を引き直すための足場（位置で解決される） */
        /* Anchor resolved by position to re-find the paragraphs after the swap */
        var anchor = paragraph.insertionPoints[0];

        swapWithNextParagraph(paragraph, nextParagraph);

        var movedParagraph = getNextParagraph(anchor.paragraphs[0]);
        if (movedParagraph) restoreSelection(movedParagraph, offsetInParagraph, selectionLength);
    }

    /**
     * 同じテキストの流れ（ストーリーまたはセル）の中で、次の段落を返す。
     * @param {Paragraph} paragraph - 基準の段落
     * @returns {Paragraph|null} 次の段落。無ければ null
     */
    function getNextParagraph(paragraph) {
        try {
            var nextParagraph = paragraph.parent.paragraphs.nextItem(paragraph);
            if (nextParagraph && nextParagraph.isValid) return nextParagraph;
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
     * 指定した段落を、ひとつ下の段落と入れ替える。
     * 下の段落を上の段落の前へ move() するので、文字・段落の書式ごと移る。
     * 下の段落が末尾で改行を持たないとき（ストーリーやセルの最終段落）は、
     * 一時的に改行を足してから入れ替え、最後に余った改行を取り除く。
     * @param {Paragraph} paragraph - 下へ移動する段落
     * @param {Paragraph} nextParagraph - その次の段落
     * @returns {void}
     */
    function swapWithNextParagraph(paragraph, nextParagraph) {
        var isLastParagraph = !endsWithParagraphBreak(nextParagraph);
        if (isLastParagraph) nextParagraph.insertionPoints[-1].contents = "\r";

        /* 参照は位置で解決されるため、move() の後で paragraph は使わない */
        /* Text references resolve by position; do not reuse paragraph after move() */
        var container = paragraph.parent;
        nextParagraph.move(LocationOptions.BEFORE, paragraph.insertionPoints[0]);

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
