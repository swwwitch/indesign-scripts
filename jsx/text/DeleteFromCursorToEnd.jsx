#target indesign

/**
 * DeleteFromCursorToEnd.jsx
 *
 * カーソル位置から後ろをまとめて削除するスクリプト。
 *
 * 動作概要：
 * - カーソル以降を、段落内で収まる場合は段落の末尾まで、
 *   それを越える場合はストーリーの末尾まで選択して削除します。
 * - 削除範囲の末尾が「。」の場合は、その「。」を残します
 *   （KEEP_TRAILING_MARU を false にすると末尾の「。」も削除します）。
 * - カーソルの直前が「。」の場合は、その「。」も削除対象に含めます。
 * - 末尾の改行（段落区切り）は「最後の文字」とみなさず、手前を末尾として扱います。
 * - 表のセル内にカーソルがある場合は、セルをまたがず、そのセル内の末尾までを上限とします。
 * - 削除はデフォルトでクリップボードを汚さない remove を使用します
 *   （COPY_TO_CLIPBOARD を true にすると cut でクリップボードへ入れます）。
 * - Undo 1 回でまとめて戻せるよう doScript（ENTIRE_SCRIPT）でラップしています。
 *
 * 使い方：テキストフレーム内にカーソルを置いた状態で実行します。
 *
 * 作成日：2024-08-15
 * 更新日：2026-06-27
 */

(function() {
    // ===== 設定 =====
    // 削除範囲の末尾が「。」の場合に「。」を残すかどうか
    //   true  … 末尾の「。」を残す
    //   false … 末尾の「。」も削除する
    var KEEP_TRAILING_MARU = true;

    // 削除した文字列をクリップボードに入れるかどうか
    //   true  … cut（クリップボードに入る／既存のコピー内容は上書きされる）
    //   false … remove（クリップボードを汚さない）
    var COPY_TO_CLIPBOARD = false;

    // ドキュメントが開いているか確認
    if (app.documents.length === 0) {
        alert("ドキュメントを開いてください。");
        return;
    }

    // 現在の選択を取得
    var sel = app.selection[0];

    // 選択がテキストフレームまたはテキスト内（挿入点）であることを確認
    if (!sel || !sel.hasOwnProperty("baseline")) {
        alert("テキストフレーム内にカーソルを置いてください。");
        return;
    }

    // Undo 1回でまとめて戻せるようにラップ
    app.doScript(function() {
        try {
            removeAfterwards(sel);
        } catch (e) {
            alert("処理に失敗しました：" + e.message);
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "カーソル以降を削除");

    function removeAfterwards(sel) {
        // 削除の上限となるコンテナ。表のセル内ならそのセル、そうでなければストーリー。
        // story.characters はセル内テキストを含まないため、必ずコンテナ自身の
        // 文字コレクションで範囲を扱い、セルをまたいだ削除も防ぐ。
        var cell = getContainingCell(sel);
        var container = cell ? cell.texts[0] : sel.parentStory;

        var chars = container.characters;
        if (chars.length === 0) {
            return;
        }

        // コンテナ先頭の位置を基準（base）に、各 index をコンテナ内 0 始まりの
        // 相対位置へ変換する（挿入点・文字とも index はストーリー基準のため）。
        var base = container.insertionPoints[0].index;

        // カーソル位置（コンテナ基準）。直前が「。」ならその「。」も削除対象に含める
        var startPos = sel.insertionPoints[0].index - base;
        if (startPos > 0 && chars[startPos - 1].contents === "。") {
            startPos = startPos - 1;
        }

        // コンテナ末尾（末尾の改行は除く）の相対位置
        var containerLastPos = lastVisiblePos(container, base);
        if (containerLastPos < 0 || startPos > containerLastPos) {
            return;
        }

        // 段落内で収まる場合は段落末尾、越える場合はコンテナ末尾を終端とする
        var paraLastPos = lastVisiblePos(sel.paragraphs[0], base);
        var endPos = Math.min(paraLastPos, containerLastPos);

        // 末尾が「。」かつ残す設定なら、その手前を終端にする
        if (KEEP_TRAILING_MARU && chars[endPos].contents === "。") {
            endPos = endPos - 1;
        }

        // 残すべき「。」より後ろに削除対象が無い場合は何もしない
        if (endPos < startPos) {
            return;
        }

        // 範囲を削除（位置でなく文字オブジェクトで指定し、座標系の曖昧さを回避）。
        // 設定によりクリップボードへ入れるか入れないかを切り替える。
        var range = chars.itemByRange(chars[startPos], chars[endPos]);
        if (COPY_TO_CLIPBOARD) {
            app.select(range); // cut はクリップボード対象を選択しておく必要がある
            app.cut();
        } else {
            range.remove();
        }
    }

    // 改行文字（段落区切り）かどうか
    function isLineBreak(ch) {
        return ch === "\r" || ch === "\n";
    }

    // テキスト（段落／セル／ストーリー）の末尾改行を除いた「最後の表示文字」の
    // 位置を、base を基準としたコンテナ内 0 始まりの相対位置で返す。空なら -1。
    function lastVisiblePos(textObj, base) {
        var n = textObj.characters.length;
        if (n === 0) {
            return -1;
        }
        var last = textObj.characters[-1];
        if (isLineBreak(last.contents)) {
            return n >= 2 ? (textObj.characters[-2].index - base) : -1;
        }
        return last.index - base;
    }

    // 選択（挿入点）が表のセル内にあればそのセルを、無ければ null を返す。
    // セル内テキストの挿入点は parent が Cell になる。通常本文では Story 等になる。
    function getContainingCell(sel) {
        try {
            var p = sel.parent;
            if (p && p.constructor && p.constructor.name === "Cell") {
                return p;
            }
        } catch (e) {}
        return null;
    }
})();
