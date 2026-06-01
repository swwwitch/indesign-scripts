#target indesign

/*

# インラインアンカー画像の高さ調整 / Fit Anchored Image Height

## 概要 / Overview

インラインアンカーされた画像フレームの高さを、周囲の文字サイズに合わせて自動調整するスクリプトです。
Scales inline anchored graphic frames to match the surrounding text size.

## 処理の流れ / Flow

1. 指定スコープ（ドキュメント／ストーリー／選択範囲）から GREP（`~a`）でアンカーオブジェクトを含む段落を検索する。
2. 画像のみの段落（前後に文字なし）は何もしない。
3. 文字がある段落では、画像以外の最大文字サイズ（pt）を取得する。
4. フレームを内容に合わせてから、縦横比を保ったまま文字サイズの高さへスケールする。

## メモ / Notes

- 単位は処理中だけ mm に変更し、終了時（finally）に必ず元へ戻す。
- `resize()` を使うため、フレームだけでなく中の画像も一緒に拡大縮小される。

*/

(function () {

    // =========================================
    // バージョン / Version
    // =========================================

    var SCRIPT_VERSION = "1.0.0";

    // =========================================
    // 定数 / Constants
    // =========================================

    /* インラインアンカー（Object Replacement Character / U+FFFC） / Inline anchored-object marker */
    var ANCHORED_OBJECT_MARKER = String.fromCharCode(0xFFFC);

    /* サイズ計算で無視する文字（改行 CR/LF・タブ・半角/全角スペース・アンカーマーカー） / Characters ignored when measuring text size */
    var IGNORABLE_CHAR_PATTERN = new RegExp("^[\\r\\n\\t 　" + ANCHORED_OBJECT_MARKER + "]$");

    /* pt → mm 換算係数（1pt = 0.352777778mm） / Point-to-millimeter conversion factor */
    var POINT_TO_MM = 0.352777778;

    // =========================================
    // 設定 / Settings
    // =========================================

    /* 処理範囲の切り替え / Processing scope
       "document"  … ドキュメント全体 / Whole document
       "story"     … カーソル／選択のあるストーリー / Story of the current selection
       "selection" … 選択範囲のみ / Current selection only */
    var SEARCH_SCOPE = "document";

    // =========================================
    // 事前チェック / Pre-checks
    // =========================================

    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var targetDocument = app.activeDocument;

    /* スコープに応じた検索対象を決定（無効ならメッセージを出して終了） / Resolve the findGrep target for the chosen scope */
    var searchTarget = resolveSearchTarget();
    if (searchTarget == null) {
        return;
    }

    /* finally で元に戻すため、変更前の状態を退避 / Save current state so finally can restore it */
    var originalHorizontalUnit = targetDocument.viewPreferences.horizontalMeasurementUnits;
    var originalVerticalUnit = targetDocument.viewPreferences.verticalMeasurementUnits;

    // =========================================
    // メイン処理 / Main
    // =========================================

    try {

        resetGrepPreferences();

        targetDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
        targetDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;

        app.findGrepPreferences.findWhat = "~a";

        var anchorSearchResults = searchTarget.findGrep();

        /* 同じ段落を二重処理しないための記録 / Track processed paragraphs to avoid duplicates */
        var processedParagraphKeys = {};

        for (var i = 0; i < anchorSearchResults.length; i++) {

            var targetParagraph = anchorSearchResults[i].paragraphs[0];

            /* 1段落に複数のアンカーがあると findGrep が同じ段落を複数回返すため、処理済みはスキップ。
               Paragraph は id を持たないので「親ストーリーの id ＋ 開始文字位置」で識別する /
               Skip paragraphs already handled (paragraphs have no id, so key by story id + start index) */
            var paragraphKey = targetParagraph.parentStory.id + "_" + targetParagraph.index;
            if (processedParagraphKeys[paragraphKey]) {
                continue;
            }
            processedParagraphKeys[paragraphKey] = true;

            var textWithoutSpaces = targetParagraph.contents.replace(/[\r\n\t 　]/g, "");

            /* 画像のみの段落（前後に文字なし）は何もしない / Skip paragraphs that contain only the image */
            if (textWithoutSpaces == ANCHORED_OBJECT_MARKER) {
                continue;
            }

            /* 前後に文字がある場合：文字サイズに合わせる（縦横比を保持） / Match the surrounding text size, keeping aspect ratio */
            var surroundingFontSizePt = getMaxTextPointSizeInParagraph(targetParagraph);

            if (surroundingFontSizePt <= 0) {
                continue;
            }

            var pageItemsInParagraph = targetParagraph.allPageItems;

            for (var j = 0; j < pageItemsInParagraph.length; j++) {

                var pageItem = pageItemsInParagraph[j];

                if (isGraphicFrame(pageItem)) {
                    pageItem.fit(FitOptions.FRAME_TO_CONTENT);   // 先にフレームを内容に合わせる / Fit frame to content first
                    scaleFrameToHeightInPoints(pageItem, surroundingFontSizePt);
                }
            }
        }

    } finally {
        /* fit()/resize() などで例外が出ても、GREP設定と単位設定を必ず元へ戻す / Always restore GREP and unit settings even on error */
        resetGrepPreferences();

        targetDocument.viewPreferences.horizontalMeasurementUnits = originalHorizontalUnit;
        targetDocument.viewPreferences.verticalMeasurementUnits = originalVerticalUnit;
    }

    // =========================================
    // 関数 / Functions
    // =========================================

    /* SEARCH_SCOPE に応じて findGrep を実行する対象を返す（無効なら alert して null） / Resolve the object to run findGrep on, based on SEARCH_SCOPE */
    function resolveSearchTarget() {

        if (SEARCH_SCOPE != "document" && SEARCH_SCOPE != "story" && SEARCH_SCOPE != "selection") {
            alert("SEARCH_SCOPE の値が不正です: " + SEARCH_SCOPE);
            return null;
        }

        if (SEARCH_SCOPE == "document") {
            return targetDocument;
        }

        /* story / selection は選択（またはテキストカーソル）が必要 / story and selection need a text selection */
        if (app.selection.length === 0) {
            alert("ストーリー／選択範囲を対象にするには、テキストを選択するかカーソルを置いてください。");
            return null;
        }

        var selectionTextObject = getSelectionTextObject(app.selection[0]);
        if (selectionTextObject == null) {
            alert("テキストが選択されていません。");
            return null;
        }

        if (SEARCH_SCOPE == "story") {
            return selectionTextObject.parentStory;
        }

        /* SEARCH_SCOPE == "selection" */
        return selectionTextObject;
    }

    /* 選択オブジェクトから findGrep 可能なテキストオブジェクトを取り出す / Extract a findGrep-capable text object from the selection */
    function getSelectionTextObject(selectedItem) {

        var typeName = selectedItem.constructor.name;

        /* テキストツールでの選択・カーソル（テキスト範囲） / Text-tool selection or insertion point */
        if (
            typeName == "InsertionPoint" || typeName == "Character" || typeName == "Word"
            || typeName == "Line" || typeName == "TextStyleRange" || typeName == "Paragraph"
            || typeName == "TextColumn" || typeName == "Text"
        ) {
            return selectedItem;
        }

        /* テキストフレームが選択されている場合は、その中のテキスト / A selected text frame -> its text */
        if (typeName == "TextFrame" && selectedItem.texts.length > 0) {
            return selectedItem.texts[0];
        }

        return null;
    }

    /* GREP 検索・置換条件を丸ごと初期化（findWhat 以外のオプションも含めて確実にクリア） / Fully reset find/change GREP preferences */
    function resetGrepPreferences() {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
    }

    /* 画像が入ったフレーム（長方形・楕円・多角形）かどうか。空フレームや画像なしフレームは除外 / Whether the page item is a graphic frame that actually contains an image */
    function isGraphicFrame(pageItem) {
        return (
            pageItem instanceof Rectangle
            || pageItem instanceof Oval
            || pageItem instanceof Polygon
        ) && pageItem.graphics.length > 0;
    }

    /* 配置画像を含むフレームを、縦横比を保ったまま指定の高さ（pt）へスケール / Scale a graphic frame (and its image) to a target height in points, keeping aspect ratio */
    // resize() は geometricBounds の書き換えと違い、フレームだけでなく中身の画像も一緒に変形する
    function scaleFrameToHeightInPoints(imageFrame, targetHeightPt) {

        var frameBounds = imageFrame.geometricBounds;          // [y1, x1, y2, x2]（mm）
        var currentHeightMM = frameBounds[2] - frameBounds[0];

        if (currentHeightMM <= 0) {
            return;
        }

        var scaleFactor = (targetHeightPt * POINT_TO_MM) / currentHeightMM;

        /* 左上を固定して、フレームと中身の画像をまとめて等倍スケール / Scale frame and image together from the top-left anchor */
        imageFrame.resize(
            CoordinateSpaces.INNER_COORDINATES,
            AnchorPoint.TOP_LEFT_ANCHOR,
            ResizeMethods.MULTIPLYING_CURRENT_DIMENSIONS_BY,
            [scaleFactor, scaleFactor]
        );
    }

    /* 段落内の画像以外の文字サイズの最大値を取得（改行・アンカー・各種スペースは除外） / Get the largest point size of real text in the paragraph */
    function getMaxTextPointSizeInParagraph(paragraph) {

        var paragraphCharacters = paragraph.characters;
        var maxPointSize = 0;

        for (var i = 0; i < paragraphCharacters.length; i++) {

            var character = paragraphCharacters[i];

            if (IGNORABLE_CHAR_PATTERN.test(character.contents)) {
                continue;
            }

            var pointSize = Number(character.pointSize);

            if (pointSize > maxPointSize) {
                maxPointSize = pointSize;
            }
        }

        return maxPointSize;
    }

})();