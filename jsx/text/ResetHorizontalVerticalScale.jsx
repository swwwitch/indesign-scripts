#target "InDesign"
/*
概要：
ドキュメント内すべてのストーリーを走査し、文字の水平比率・垂直比率が
100% でない範囲を 100% に戻します（変倍のリセット）。表（テーブル）セル
内の文字も、入れ子の表を含めて再帰的に対象とします。処理全体を1つの
取り消し（Undo）にまとめ、最後に変更した文字範囲の件数を表示します。

Overview:
Scans every story in the active document and resets any text whose
horizontal or vertical scale is not 100% back to 100% (clears scaling).
Text inside table cells is also covered, recursing into nested tables.
The whole run is a single undo step, and the number of changed text
ranges is shown when finished.
*/

(function () {
    /* ドキュメントの有無を確認 / Ensure a document is open */
    if (app.documents.length === 0) {
        alert("ドキュメントを開いてください。");
        return;
    }

    var targetDocument = app.activeDocument;

    /**
     * 文字範囲の配列を走査し、水平・垂直比率が 100% でないものを 100% に戻す。
     * Reset every style range in the given array whose scale is not 100%.
     * @param {TextStyleRange[]} styleRanges 対象の文字スタイル範囲の配列 / style ranges to process
     * @returns {number} リセットした文字範囲の件数 / number of ranges that were reset
     */
    function resetScaleInRanges(styleRanges) {
        var resetCount = 0;

        for (var k = 0; k < styleRanges.length; k++) {
            var styleRange = styleRanges[k];

            if (
                styleRange.horizontalScale !== 100 ||
                styleRange.verticalScale !== 100
            ) {
                styleRange.horizontalScale = 100;
                styleRange.verticalScale = 100;
                resetCount++;
            }
        }

        return resetCount;
    }

    /**
     * テキストコンテナ（ストーリーまたはセル内テキスト）の文字比率を戻し、
     * 内包する表のセルへ再帰する（入れ子の表にも対応）。
     * Reset a text container's scale, then recurse into its table cells.
     * @param {Story|Text} textContainer textStyleRanges と tables を持つオブジェクト / a story or cell text
     * @returns {number} リセットした文字範囲の件数 / number of ranges that were reset
     */
    function resetScaleInContainer(textContainer) {
        var resetCount = resetScaleInRanges(
            textContainer.textStyleRanges.everyItem().getElements()
        );

        var tables = textContainer.tables.everyItem().getElements();
        for (var tableIndex = 0; tableIndex < tables.length; tableIndex++) {
            var cells = tables[tableIndex].cells.everyItem().getElements();
            for (var cellIndex = 0; cellIndex < cells.length; cellIndex++) {
                /* セル内テキストを同じ処理で再帰 / Recurse into the cell's text */
                resetCount += resetScaleInContainer(cells[cellIndex].texts[0]);
            }
        }

        return resetCount;
    }

    app.doScript(
        /* 全ストーリー（表セル含む）の文字比率を100%に戻す / Reset every story's scale, including table cells */
        function resetTextScaleToHundred() {
            var stories = targetDocument.stories.everyItem().getElements();
            var resetRangeCount = 0;

            for (var i = 0; i < stories.length; i++) {
                resetRangeCount += resetScaleInContainer(stories[i]);
            }

            alert("完了しました。\n変更した文字範囲: " + resetRangeCount);
        },
        ScriptLanguage.JAVASCRIPT,
        undefined,
        UndoModes.ENTIRE_SCRIPT,
        "文字の水平・垂直比率を100%にする"
    );
})();
