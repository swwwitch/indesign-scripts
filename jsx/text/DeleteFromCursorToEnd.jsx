#target indesign

/*

### 概要

カーソル位置から、その段落の末尾までをまとめて削除するスクリプトです。

- カーソル位置から現在の段落末尾まで削除します。セル内ではセル境界を越えません。
- 削除範囲の末尾が「。」の場合は残します（`KEEP_TRAILING_MARU` で切替）。
- カーソルの直前が「。」の場合は、その「。」も削除対象に含めます。
- 末尾の改行（段落区切り）は「最後の文字」とみなさず、手前を末尾として扱います。
- 既定では `cut` で削除し、削除した文字列をクリップボードへ入れます（`COPY_TO_CLIPBOARD` を `false` にすると `remove` でクリップボードを汚しません）。
- Undo 1 回でまとめて戻せます。

**使い方**：テキストフレーム内にカーソルを置いた状態で実行します。

### 紹介記事（note）

https://note.com/dtp_tranist/n/nf0b1e27e1f81

### Overview

Deletes from the cursor to the end of the current paragraph in a single action.

- Deletes from the cursor to the end of the current paragraph. Inside a table cell, it never crosses the cell boundary.
- Keeps a trailing `。` when present (toggle with `KEEP_TRAILING_MARU`).
- Also removes a `。` placed immediately before the cursor.
- Treats a trailing line break (paragraph separator) as not the last character, using the character before it as the end.
- Uses `cut` by default, placing the deleted text on the clipboard (set `COPY_TO_CLIPBOARD` to `false` to use `remove` and keep the clipboard intact).
- Can be undone in a single step.

**Usage**: Run with the cursor placed inside a text frame.

作成日 / Created: 2024-08-15
更新日 / Updated: 2026-06-27

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.1.0";

(function () {
    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 末尾が「。」のとき「。」を残すか / Keep a trailing "。" when present */
    //   true  … 残す   / keep
    //   false … 削除する / delete
    var KEEP_TRAILING_MARU = true;

    /* 削除した文字列をクリップボードへ入れるか / Put the deleted text on the clipboard */
    //   true  … cut（クリップボードに入る／既存のコピー内容は上書き）/ cut (goes to clipboard)
    //   false … remove（クリップボードを汚さない）               / remove (keeps clipboard intact)
    var COPY_TO_CLIPBOARD = true;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 実行環境の言語を判定（ja / en）/ Detect the runtime language (ja / en) */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noTextCursor: { ja: "テキストフレーム内にカーソルを置いてください。", en: "Place the cursor inside a text frame." }
        },
        error: {
            /* 末尾コロンは日本語は全角・英語は半角 / Trailing colon: full-width JA, half-width EN */
            failed: { ja: "処理に失敗しました：", en: "Processing failed: " }
        },
        undo: {
            stepName: { ja: "カーソル以降を削除", en: "Delete from cursor to end" }
        }
    };

    /* ラベル取得（ドット区切りで階層参照）。未解決時は en → キー文字列にフォールバック /
       Get a localized label (dot-path lookup). Falls back to en, then to the key string. */
    function L(key) {
        var pathParts = key.split(".");
        var entry = LABELS;
        for (var i = 0; i < pathParts.length; i++) {
            if (entry == null) {
                return key;
            }
            entry = entry[pathParts[i]];
        }
        if (entry == null) {
            return key;
        }
        if (entry[currentLanguage] != null) {
            return entry[currentLanguage];
        }
        return entry.en != null ? entry.en : key;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* ドキュメントが開いているか / A document must be open */
    if (app.documents.length === 0) {
        alert(L("alert.noDocument"));
        return;
    }

    /* 選択がテキスト上の挿入点／範囲か / The selection must be a text insertion point or range */
    var selection = app.selection[0];
    if (!isTextSelection(selection)) {
        alert(L("alert.noTextCursor"));
        return;
    }

    /* Undo 1 回でまとめて戻せるようにラップ / Wrap so it can be undone in a single step */
    app.doScript(function () {
        try {
            deleteFromCursorToEnd(selection);
        } catch (e) {
            alert(L("error.failed") + e.message);
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, L("undo.stepName"));

    /* カーソルから現在の段落末尾まで削除（セル境界は越えない）/
       Delete from the cursor to the end of the current paragraph (never crossing a cell boundary) */
    function deleteFromCursorToEnd(selection) {
        /* 削除の上限となるコンテナ。セル内ならそのセル、そうでなければストーリー。
           story.characters はセル内テキストを含まないため、コンテナ自身の文字コレクションで扱う。
           Container that bounds the deletion: the cell when inside one, otherwise the story.
           story.characters does not include cell text, so always use the container's own collection. */
        var cell = getContainingCell(selection);
        var textContainer = cell ? cell.texts[0] : selection.parentStory;

        var characters = textContainer.characters;

        /* コンテナ先頭の index を基準に、各 index をコンテナ内 0 始まりの相対位置へ変換 /
           Use the container's first index as the base to get container-relative offsets */
        var baseIndex = textContainer.insertionPoints[0].index;

        /* カーソル位置（相対）。直前が「。」ならその「。」も削除対象に含める /
           Cursor offset; also include a "。" immediately before the cursor */
        var startOffset = selection.insertionPoints[0].index - baseIndex;
        if (startOffset > 0 && characters.item(startOffset - 1).contents === "。") {
            startOffset = startOffset - 1;
        }

        /* コンテナ末尾（末尾の改行は除く）の相対位置 / Last visible offset of the container */
        var containerLastOffset = lastVisibleCharOffset(textContainer, baseIndex);
        if (containerLastOffset < 0 || startOffset > containerLastOffset) {
            return;
        }

        /* 終端はカーソルのある段落の末尾。コンテナ末尾で上限クランプ（保険）/
           End at the cursor's paragraph end, clamped to the container end as a safeguard */
        var paragraphLastOffset = lastVisibleCharOffset(selection.paragraphs[0], baseIndex);
        var endOffset = Math.min(paragraphLastOffset, containerLastOffset);

        /* 削除対象が無ければ終了。空段落では段落末尾が -1 になるため、
           「。」判定（item アクセス）より前にここで弾いて endOffset を非負に保つ。
           Bail out on an empty range first (empty paragraph yields -1), keeping
           endOffset non-negative before the "。" check below. */
        if (endOffset < startOffset) {
            return;
        }

        /* 末尾が「。」かつ残す設定なら、その手前を終端にする / Keep a trailing "。" when configured */
        if (KEEP_TRAILING_MARU && characters.item(endOffset).contents === "。") {
            endOffset = endOffset - 1;
            if (endOffset < startOffset) {
                return; /* 「。」を残すと削除対象が無くなる / Nothing left after keeping "。" */
            }
        }

        /* 範囲を文字オブジェクトで指定して削除（座標系の曖昧さを回避）/
           Delete the range by character objects to avoid index-coordinate ambiguity */
        var deleteRange = characters.itemByRange(characters.item(startOffset), characters.item(endOffset));
        if (COPY_TO_CLIPBOARD) {
            app.select(deleteRange); /* cut は対象選択が必要 / cut needs a selection */
            app.cut();
        } else {
            deleteRange.remove();
        }
    }

    /* 選択がテキスト上の挿入点または範囲か / Whether the selection is a text insertion point or range */
    function isTextSelection(obj) {
        if (obj == null) {
            return false;
        }
        var typeName = obj.constructor.name;
        return typeName === "InsertionPoint" ||
            typeName === "Text" ||
            typeName === "Character" ||
            typeName === "Word" ||
            typeName === "Line" ||
            typeName === "Paragraph" ||
            typeName === "TextStyleRange";
    }

    /* 改行文字（段落区切り）か / Whether the character content is a line break */
    function isLineBreak(content) {
        return content === "\r" || content === "\n";
    }

    /* 末尾改行を除いた「最後の表示文字」の、コンテナ内相対位置を返す（空なら -1）/
       Container-relative offset of the last visible char, excluding a trailing break; -1 if empty */
    function lastVisibleCharOffset(textObject, baseIndex) {
        var characters = textObject.characters;
        var charCount = characters.length;
        if (charCount === 0) {
            return -1;
        }
        var lastChar = characters.item(charCount - 1);
        if (isLineBreak(lastChar.contents)) {
            return charCount >= 2 ? (characters.item(charCount - 2).index - baseIndex) : -1;
        }
        return lastChar.index - baseIndex;
    }

    /* 選択（挿入点）が表のセル内ならそのセルを返す。親を辿って Cell を探し、
       Story / Document などコンテナ階層に達したら打ち切る（無効オブジェクト回避）。
       Walk up parents to find the containing cell; stop at the story/document level
       to avoid touching invalid objects. Returns null when not inside a cell. */
    function getContainingCell(selection) {
        try {
            var node = selection;
            for (var depth = 0; node != null && depth < 6; depth++) {
                var typeName = node.constructor.name;
                if (typeName === "Cell") {
                    return node;
                }
                /* セル外のコンテナまで上がったら終了 / Reached a non-cell container; give up */
                if (typeName === "Story" || typeName === "Document" || typeName === "Application") {
                    return null;
                }
                node = node.parent;
            }
        } catch (e) {}
        return null;
    }
})();
