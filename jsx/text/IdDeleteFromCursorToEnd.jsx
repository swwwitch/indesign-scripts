#target indesign

/*
 * IdDeleteFromCursorToEnd.jsx
 *
 * カーソル位置からその段落の末尾までをまとめて削除します。段落末尾の記号の直前では、その記号 1 文字だけを削除します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdDeleteFromCursorToEnd";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-05";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdDeleteFromCursorToEnd.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdDeleteFromCursorToEnd.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf0b1e27e1f81"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 末尾が「。」のとき「。」を残すか（true: 残す / false: 削除する）
   / Keep a trailing "。" when present (true: keep, false: delete) */
var KEEP_TRAILING_MARU = true;

/* 削除した文字列をクリップボードへ入れるか（true: cut / false: remove）
   / Put the deleted text on the clipboard (true: cut, false: remove) */
var COPY_TO_CLIPBOARD = true;

/* 段落末尾でカーソルの直後にあるとき、1 文字だけ削除する記号
   / Marks deleted alone when they sit right after the cursor at the paragraph end */
var TRAILING_MARKS = ["。", "！", "？", ",", ".", "、", "，", "．"];

/* 親をたどってセルを探すときの最大階層 / Maximum depth when walking up to find a containing cell */
var MAX_CELL_LOOKUP_DEPTH = 6;

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

var currentLanguage = getCurrentLang();

var LABELS = {
    alert: {
        noDocument:   { ja: "ドキュメントを開いてください。", en: "Please open a document." },
        noTextCursor: { ja: "テキストフレーム内にカーソルを置いてください。", en: "Place the cursor inside a text frame." }
    },
    error: {
        /* 末尾コロンは日本語が全角、英語が半角 / Trailing colon: full-width JA, half-width EN */
        failed: { ja: "処理に失敗しました：", en: "Processing failed: " }
    },
    undo: {
        deleteToEnd: { ja: "カーソル以降を削除", en: "Delete from cursor to end" }
    }
};

/**
 * ドット区切りキーでラベルを取得する
 * @param {string} labelKey 例: "alert.noDocument"
 * @returns {string} 現在の言語のラベル文字列。見つからない場合はキーをそのまま返す
 */
function getLabel(labelKey) {
    var keyParts = labelKey.split(".");
    var node = LABELS;
    for (var i = 0; i < keyParts.length; i++) {
        if (node == null) return labelKey;
        node = node[keyParts[i]];
    }
    if (node == null) return labelKey;
    if (node[currentLanguage] != null) return node[currentLanguage];
    return (node.en != null) ? node.en : labelKey;
}

// =========================================
// 判定用のユーティリティ / Predicates
// =========================================

/**
 * 選択がテキスト上の挿入点または範囲かどうかを判定する
 * @param {object} selectionItem 選択オブジェクト
 * @returns {boolean} テキスト上の選択なら true
 */
function isTextSelection(selectionItem) {
    if (selectionItem == null) return false;
    var typeName = selectionItem.constructor.name;
    return typeName === "InsertionPoint" || typeName === "Text" || typeName === "Character" ||
        typeName === "Word" || typeName === "Line" || typeName === "Paragraph" ||
        typeName === "TextStyleRange";
}

/**
 * 改行文字（段落区切り）かどうかを判定する
 * @param {string} characterContents 文字の内容
 * @returns {boolean} 改行なら true
 */
function isLineBreak(characterContents) {
    return characterContents === "\r" || characterContents === "\n";
}

/**
 * TRAILING_MARKS に含まれる記号かどうかを判定する
 * @param {string} characterContents 文字の内容
 * @returns {boolean} 対象の記号なら true
 */
function isTrailingMark(characterContents) {
    for (var i = 0; i < TRAILING_MARKS.length; i++) {
        if (TRAILING_MARKS[i] === characterContents) return true;
    }
    return false;
}

// =========================================
// 位置の算出 / Offset helpers
// =========================================

/**
 * 末尾改行を除いた「最後の表示文字」のコンテナ内相対位置を求める
 * @param {object} textObject characters を持つテキストオブジェクト
 * @param {number} baseIndex コンテナ先頭の index
 * @returns {number} 相対位置。文字がない場合は -1
 */
function lastVisibleCharOffset(textObject, baseIndex) {
    var characterList = textObject.characters;
    var characterCount = characterList.length;
    if (characterCount === 0) return -1;

    var lastCharacter = characterList.item(characterCount - 1);
    if (isLineBreak(lastCharacter.contents)) {
        return (characterCount >= 2) ? (characterList.item(characterCount - 2).index - baseIndex) : -1;
    }
    return lastCharacter.index - baseIndex;
}

/**
 * 選択が表のセル内にあるとき、そのセルを返す
 * @param {object} selectionItem 選択オブジェクト
 * @returns {Cell|null} セル。セル内でなければ null
 */
function getContainingCell(selectionItem) {
    try {
        var currentNode = selectionItem;
        for (var depth = 0; currentNode != null && depth < MAX_CELL_LOOKUP_DEPTH; depth++) {
            var typeName = currentNode.constructor.name;
            if (typeName === "Cell") return currentNode;

            /* セル外のコンテナまで上がったら打ち切る（無効オブジェクト回避）
               / Stop at a non-cell container to avoid touching invalid objects */
            if (typeName === "Story" || typeName === "Document" || typeName === "Application") return null;

            currentNode = currentNode.parent;
        }
    } catch (e) {}
    return null;
}

// =========================================
// 削除処理 / Deletion
// =========================================

/**
 * 指定した範囲を削除する（設定に応じて cut または remove）
 * @param {object} textRange 削除する文字範囲
 * @returns {void}
 */
function deleteTextRange(textRange) {
    if (COPY_TO_CLIPBOARD) {
        app.select(textRange); /* cut は対象の選択が必要 / cut needs a selection */
        app.cut();
    } else {
        textRange.remove();
    }
}

/**
 * カーソルから現在の段落末尾までを削除する（セル境界は越えない）
 * @param {object} selectionItem テキスト上の選択または挿入点
 * @returns {void}
 */
function deleteFromCursorToEnd(selectionItem) {
    /* 削除の上限となるコンテナ。セル内ならそのセル、そうでなければストーリー。
       story.characters はセル内テキストを含まないため、コンテナ自身の文字コレクションを使う。
       / Bound the deletion by the cell when inside one, otherwise by the story.
         story.characters excludes cell text, so always use the container's own collection. */
    var containingCell = getContainingCell(selectionItem);
    var textContainer  = containingCell ? containingCell.texts[0] : selectionItem.parentStory;
    var characterList  = textContainer.characters;

    /* コンテナ先頭の index を基準に、各 index を 0 始まりの相対位置へ変換
       / Use the container's first index as the base for container-relative offsets */
    var baseIndex = textContainer.insertionPoints[0].index;

    /* カーソル直後が TRAILING_MARKS の記号で、それが段落の最後の表示文字なら、その 1 文字だけを削除して終了
       （KEEP_TRAILING_MARU より優先）
       / When the char right after the cursor is a trailing mark and is the paragraph's last visible char,
         delete just that char and stop. This takes priority over KEEP_TRAILING_MARU. */
    var cursorOffset       = selectionItem.insertionPoints[0].index - baseIndex;
    var paragraphEndOffset = lastVisibleCharOffset(selectionItem.paragraphs[0], baseIndex);
    if (cursorOffset === paragraphEndOffset && isTrailingMark(characterList.item(cursorOffset).contents)) {
        deleteTextRange(characterList.item(cursorOffset));
        return;
    }

    /* カーソル位置。直前が「。」ならその「。」も削除対象に含める
       / Start at the cursor, also taking in a "。" immediately before it */
    var startOffset = cursorOffset;
    if (startOffset > 0 && characterList.item(startOffset - 1).contents === "。") {
        startOffset = startOffset - 1;
    }

    var containerLastOffset = lastVisibleCharOffset(textContainer, baseIndex);
    if (containerLastOffset < 0 || startOffset > containerLastOffset) return;

    /* 終端はカーソルのある段落の末尾。コンテナ末尾で上限クランプする（保険）
       / End at the cursor's paragraph end, clamped to the container end as a safeguard */
    var endOffset = Math.min(paragraphEndOffset, containerLastOffset);

    /* 空段落では段落末尾が -1 になるため、item アクセスの前にここで弾く
       / An empty paragraph yields -1, so bail out before any item() access */
    if (endOffset < startOffset) return;

    /* 末尾が「。」かつ残す設定なら、その手前を終端にする / Keep a trailing "。" when configured */
    if (KEEP_TRAILING_MARU && characterList.item(endOffset).contents === "。") {
        endOffset = endOffset - 1;
        if (endOffset < startOffset) return;
    }

    /* 座標系の曖昧さを避けるため、文字オブジェクトで範囲を指定する
       / Specify the range with character objects to avoid index-coordinate ambiguity */
    deleteTextRange(characterList.itemByRange(characterList.item(startOffset), characterList.item(endOffset)));
}

// =========================================
// メイン処理 / Main
// =========================================

if (app.documents.length === 0) {
    alert(getLabel("alert.noDocument"));
    return;
}

var currentSelection = app.selection[0];
if (!isTextSelection(currentSelection)) {
    alert(getLabel("alert.noTextCursor"));
    return;
}

/* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
app.doScript(function () {
    try {
        deleteFromCursorToEnd(currentSelection);
    } catch (e) {
        alert(getLabel("error.failed") + e.message);
    }
}, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.deleteToEnd"));

})();
