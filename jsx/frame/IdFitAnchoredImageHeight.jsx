#target indesign

/*
 * IdFitAnchoredImageHeight.jsx
 *
 * インラインアンカーされた画像フレームの高さを、同じ段落の文字サイズに合わせて縦横比を保ったまま調整します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdFitAnchoredImageHeight";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-06-01";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdFitAnchoredImageHeight.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdFitAnchoredImageHeight.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 処理範囲 / Processing scope
   "document"  … ドキュメント全体 / Whole document
   "story"     … カーソル／選択のあるストーリー / Story of the current selection
   "selection" … 選択範囲のみ / Current selection only */
var SEARCH_SCOPE = "document";

/* インラインアンカー（Object Replacement Character / U+FFFC）/ Inline anchored-object marker */
var ANCHORED_OBJECT_MARKER = String.fromCharCode(0xFFFC);

/* サイズ計算で無視する文字（改行・タブ・半角/全角スペース・アンカーマーカー）/ Characters ignored when measuring text size */
var IGNORABLE_CHAR_PATTERN = new RegExp("^[\\r\\n\\t 　" + ANCHORED_OBJECT_MARKER + "]$");

/* pt → mm 換算係数（1pt = 0.352777778mm）/ Point-to-millimeter conversion factor */
var POINT_TO_MM = 0.352777778;

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
        noDocument:  { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        invalidScope: { ja: "SEARCH_SCOPE の値が不正です: ", en: "Invalid SEARCH_SCOPE value: " },
        needSelection: {
            ja: "ストーリー／選択範囲を対象にするには、テキストを選択するかカーソルを置いてください。",
            en: "To target a story or selection, select text or place the cursor first."
        },
        noTextSelected: { ja: "テキストが選択されていません。", en: "No text is selected." }
    },
    undo: {
        fitAnchoredImages: { ja: "アンカー画像の高さ調整", en: "Fit Anchored Image Height" }
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
// 前提チェック / Preconditions
// =========================================
if (app.documents.length === 0) {
    alert(localize(LABELS.alert.noDocument));
    return;
}

var targetDocument = app.activeDocument;

// =========================================
// 検索対象の判定 / Search target
// =========================================

/**
 * 選択オブジェクトから findGrep を実行できるテキストオブジェクトを取り出す
 * @param {object} selectedItem 選択オブジェクト
 * @returns {object|null} テキストオブジェクト。取り出せない場合は null
 */
function getSelectionTextObject(selectedItem) {
    var typeName = selectedItem.constructor.name;

    /* テキストツールでの選択・カーソル / Text-tool selection or insertion point */
    if (typeName === "InsertionPoint" || typeName === "Character" || typeName === "Word" ||
        typeName === "Line" || typeName === "TextStyleRange" || typeName === "Paragraph" ||
        typeName === "TextColumn" || typeName === "Text") {
        return selectedItem;
    }

    /* テキストフレーム選択時はその中のテキスト / A selected text frame yields its text */
    if (typeName === "TextFrame" && selectedItem.texts.length > 0) return selectedItem.texts[0];

    return null;
}

/**
 * SEARCH_SCOPE に応じて findGrep の対象を決める
 * @returns {object|null} 検索対象。決められない場合は null
 */
function resolveSearchTarget() {
    if (SEARCH_SCOPE !== "document" && SEARCH_SCOPE !== "story" && SEARCH_SCOPE !== "selection") {
        alert(localize(LABELS.alert.invalidScope) + SEARCH_SCOPE);
        return null;
    }

    if (SEARCH_SCOPE === "document") return targetDocument;

    /* story / selection は選択（またはテキストカーソル）が必要 / story and selection need a text selection */
    if (app.selection.length === 0) {
        alert(localize(LABELS.alert.needSelection));
        return null;
    }

    var selectionTextObject = getSelectionTextObject(app.selection[0]);
    if (selectionTextObject == null) {
        alert(localize(LABELS.alert.noTextSelected));
        return null;
    }

    return (SEARCH_SCOPE === "story") ? selectionTextObject.parentStory : selectionTextObject;
}

// =========================================
// 補助関数 / Helpers
// =========================================

/**
 * GREP の検索・置換条件をすべて初期化する
 * @returns {void}
 */
function resetGrepPreferences() {
    app.findGrepPreferences = NothingEnum.NOTHING;
    app.changeGrepPreferences = NothingEnum.NOTHING;
}

/**
 * 画像が入ったフレーム（長方形・楕円・多角形）かどうかを判定する
 * @param {PageItem} pageItem 判定するオブジェクト
 * @returns {boolean} 画像入りのフレームなら true
 */
function isGraphicFrame(pageItem) {
    return (pageItem instanceof Rectangle || pageItem instanceof Oval || pageItem instanceof Polygon) &&
        pageItem.graphics.length > 0;
}

/**
 * 段落内で画像以外の文字サイズの最大値を求める
 * @param {Paragraph} paragraph 対象の段落
 * @returns {number} 最大の文字サイズ（pt）
 */
function getMaxTextPointSizeInParagraph(paragraph) {
    var paragraphCharacters = paragraph.characters;
    var maxPointSize = 0;

    for (var i = 0; i < paragraphCharacters.length; i++) {
        var character = paragraphCharacters[i];
        if (IGNORABLE_CHAR_PATTERN.test(character.contents)) continue;

        var pointSize = Number(character.pointSize);
        if (pointSize > maxPointSize) maxPointSize = pointSize;
    }

    return maxPointSize;
}

/**
 * 画像フレームを縦横比を保ったまま指定の高さへスケールする
 * @param {PageItem} imageFrame 対象の画像フレーム
 * @param {number} targetHeightPt 目標の高さ（pt）
 * @returns {void}
 */
function scaleFrameToHeightInPoints(imageFrame, targetHeightPt) {
    /* geometricBounds は [上, 左, 下, 右]（mm）/ geometricBounds is [top, left, bottom, right] in mm */
    var frameBounds = imageFrame.geometricBounds;
    var currentHeightMM = frameBounds[2] - frameBounds[0];
    if (currentHeightMM <= 0) return;

    var scaleFactor = (targetHeightPt * POINT_TO_MM) / currentHeightMM;

    /* resize() は geometricBounds の書き換えと違い、中の画像も一緒に変形する
       / Unlike rewriting geometricBounds, resize() scales the placed image too */
    imageFrame.resize(
        CoordinateSpaces.INNER_COORDINATES,
        AnchorPoint.TOP_LEFT_ANCHOR,
        ResizeMethods.MULTIPLYING_CURRENT_DIMENSIONS_BY,
        [scaleFactor, scaleFactor]
    );
}

// =========================================
// メイン処理 / Main
// =========================================

/**
 * アンカー画像を含む段落を走査し、画像の高さを文字サイズに合わせる
 * @param {object} searchTarget findGrep を実行する対象
 * @returns {void}
 */
function fitAnchoredImages(searchTarget) {
    /* finally で元に戻すため、変更前の単位を退避 / Save the units so finally can restore them */
    var originalHorizontalUnit = targetDocument.viewPreferences.horizontalMeasurementUnits;
    var originalVerticalUnit   = targetDocument.viewPreferences.verticalMeasurementUnits;

    try {
        resetGrepPreferences();

        targetDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
        targetDocument.viewPreferences.verticalMeasurementUnits   = MeasurementUnits.MILLIMETERS;

        app.findGrepPreferences.findWhat = "~a";
        var anchorSearchResults = searchTarget.findGrep();

        /* 同じ段落を二重処理しないための記録 / Track processed paragraphs to avoid duplicates */
        var processedParagraphKeys = {};

        for (var i = 0; i < anchorSearchResults.length; i++) {
            var targetParagraph = anchorSearchResults[i].paragraphs[0];

            /* 1 段落に複数のアンカーがあると findGrep が同じ段落を複数回返す。
               Paragraph は id を持たないため「親ストーリーの id ＋ 開始位置」で識別する
               / Paragraphs have no id, so key them by story id plus start index */
            var paragraphKey = targetParagraph.parentStory.id + "_" + targetParagraph.index;
            if (processedParagraphKeys[paragraphKey]) continue;
            processedParagraphKeys[paragraphKey] = true;

            var textWithoutSpaces = targetParagraph.contents.replace(/[\r\n\t 　]/g, "");

            /* 画像のみの段落（前後に文字なし）は対象外 / Skip paragraphs that contain only the image */
            if (textWithoutSpaces === ANCHORED_OBJECT_MARKER) continue;

            var surroundingFontSizePt = getMaxTextPointSizeInParagraph(targetParagraph);
            if (surroundingFontSizePt <= 0) continue;

            var pageItemsInParagraph = targetParagraph.allPageItems;
            for (var j = 0; j < pageItemsInParagraph.length; j++) {
                var pageItem = pageItemsInParagraph[j];
                if (!isGraphicFrame(pageItem)) continue;

                /* 先にフレームを内容に合わせる / Fit the frame to its content first */
                pageItem.fit(FitOptions.FRAME_TO_CONTENT);
                scaleFrameToHeightInPoints(pageItem, surroundingFontSizePt);
            }
        }
    } finally {
        /* 例外が出ても GREP 設定と単位設定は必ず元へ戻す / Always restore GREP and unit settings, even on error */
        resetGrepPreferences();
        targetDocument.viewPreferences.horizontalMeasurementUnits = originalHorizontalUnit;
        targetDocument.viewPreferences.verticalMeasurementUnits   = originalVerticalUnit;
    }
}

/* スコープに応じた検索対象を決定（決まらなければ何もしない）/ Resolve the search target; do nothing when it cannot be resolved */
var searchTarget = resolveSearchTarget();
if (searchTarget == null) return;

/* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
app.doScript(function () {
    fitAnchoredImages(searchTarget);
}, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, localize(LABELS.undo.fitAnchoredImages));

})();
