#target inDesign;

/*
スクリプトの概要：
このスクリプトは、InDesign で選択中のテキストまたはカーソル位置に対し、ユーザーがダイアログで入力した複数行テキストを挿入・置換します。
また、「@# 挿入」ボタンを追加し、入力欄の末尾に視覚的な強制改行マーカーを追記できます。

主な仕様：
- テキスト選択中はその範囲を置換、カーソル位置のみなら挿入、TextFrame選択時は末尾に挿入
- 入力欄では Enter により段落（\r）を挿入可能
- @# を入力すると視覚的な強制改行（\n）として扱い、確定時に \n に復元
- ［改行削除］ボタンにより \n / \r / \\n / \\r / @# を一括削除
- ボタン（改行削除 / @# 挿入 / OK / キャンセル）は同一行に横並び
- 「@# 挿入」ボタンは入力欄の末尾に @# を追記
- 処理は Undo 対応で実行され、一括で戻すことが可能

※ テキストが選択されていない／テキストフレームが未選択でも、入力内容を新規テキストフレームとして挿入可能

処理の流れ：
1. ロケールをもとに UI ラベルを定義
2. 選択テキストを初期値としてダイアログに表示（\n → @#）
3. 入力を正規化（\r 統一、\n 復元）し、テキスト挿入または置換
4. 「@# 挿入」ボタンで入力欄末尾に強制改行マーカーを追加

対象オブジェクト：
- TextFrame
- Text（挿入ポイントまたは選択範囲）

更新履歴：
- 作成日：2025-05-27
- 更新日：2025-05-28（UI整理、改行変換調整、ボタン整列、ラベル分離）
- 0.1.0：初版リリース
- 0.1.1：改行削除ボタンの挙動修正、ラベル定義のスコープ外宣言
- 0.1.2：テキストが選択されていない／テキストフレームが未選択でも、入力内容を新規テキストフレームとして挿入可能
- 0.1.3：「@# 挿入」ボタンを追加（入力欄末尾に追記）
*/

// ラベル定義（日本語／英語）をスコープ外で宣言し、クロージャで再利用可能にする
var LABELS = {
    errorNoTextFrame: {
        ja: "テキストフレームを選択してください。", // 使用箇所: 選択が不正な場合のエラーダイアログ
        en: "Please select a text frame."
    },
    errorOccurred: {
        ja: "エラーが発生しました：\n", // 使用箇所: 例外発生時のエラーダイアログ
        en: "An error occurred:\n"
    },
    dialogTitle: {
        ja: "テキスト編集", // 使用箇所: ダイアログタイトル
        en: "Edit Text"
    },
    buttonOK: {
        ja: "OK", // 使用箇所: ダイアログOKボタン
        en: "OK"
    },
    buttonCancel: {
        ja: "キャンセル", // 使用箇所: ダイアログキャンセルボタン
        en: "Cancel"
    },
    buttonClearAll: {
        ja: "改行全削除", // 使用箇所: 改行削除ボタン
        en: "Clear All"
    },
    labelNote: {
        ja: "@#で強制改行（\\n）、fn + returnで確定", // 使用箇所: ダイアログの説明ラベル
        en: "@# = forced line break (\\n), fn + return to confirm"
    },
    buttonInsertSoftBreak: {
        ja: "@# 挿入", // 使用箇所: @# 挿入ボタン
        en: "Insert @#"
    }
};

function main() {
    // ロケールを取得し、日本語か英語かを判定
    var lang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    // ダイアログを表示：初期テキスト付き、改行削除・@# 挿入・OK・キャンセルは同一行に横並び
    function showMultilineDialog(initialText) {
        var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";


        var inputBox = dialog.add("edittext", undefined, initialText || "", {
            multiline: true
        });
        inputBox.preferredSize = [350, 160];


        var labelGroup = dialog.add("group");
        labelGroup.orientation = "column";
        labelGroup.alignChildren = ["center", "top"];
        labelGroup.margins = [0, 0, 0, 10]; // 下に余白を追加

        labelGroup.add("statictext", undefined, LABELS.labelNote[lang]);

        // ボタングループ（改行削除、@# 挿入、OK/キャンセル）を右下に配置
        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = "fill";

        var clearBtn = buttonGroup.add("button", undefined, LABELS.buttonClearAll[lang]);

        // 「@# 挿入」ボタンを追加
        var insertSoftBreakBtn = buttonGroup.add("button", undefined, LABELS.buttonInsertSoftBreak[lang]);

        // スペーサーを追加してボタン間に余白を確保（幅を30に変更）
        buttonGroup.add("statictext", undefined, "").preferredSize.width = 30;

        var cancelBtn = buttonGroup.add("button", undefined, LABELS.buttonCancel[lang], {
            name: "cancel"
        });
        var okBtn = buttonGroup.add("button", undefined, LABELS.buttonOK[lang], {
            name: "ok"
        });

        clearBtn.onClick = function() {
            inputBox.text = inputBox.text
                .replace(/\\[nr]/g, "") // remove \n and \r literals
                .replace(/[\n\r]/g, "") // remove actual line breaks
                .replace(/@#/g, ""); // remove visual marker
            inputBox.active = true;
            inputBox.selection = [inputBox.text.length, inputBox.text.length];
        };

        // 「@# 挿入」ボタンのクリック時の挙動（末尾に @# を追記）
        insertSoftBreakBtn.onClick = function() {
            inputBox.text += "@#";
            inputBox.selection = [inputBox.text.length, inputBox.text.length];
            inputBox.active = true;
        };

        inputBox.active = true; // 入力欄にフォーカス
        // 入力なしのときは先頭に、初期値があれば選択状態を維持
        if (!initialText) {
            inputBox.selection = [0, 0]; // 入力なしのときは先頭に
        }
        if (dialog.show() === 1) {
            return inputBox.text;
        } else {
            return null;
        }
    }

    // 選択テキストがあれば初期値に設定（\n を @# に変換）
    var initialText = "";
    if (app.selection && app.selection.length === 1 && app.selection[0].hasOwnProperty("contents")) {
        initialText = app.selection[0].contents.replace(/\n/g, "@#");
    }

    var userInput = showMultilineDialog(initialText);

    var validSelection = app.selection && app.selection.length === 1 &&
        (app.selection[0] instanceof TextFrame ||
            app.selection[0].hasOwnProperty("contents") ||
            app.selection[0].hasOwnProperty("insertionPoints"));

    if (userInput) {
        userInput = userInput
            .replace(/\\n/g, "\n")
            .replace(/\\r/g, "\r")
            .replace(/\r\n/g, "\r")
            .replace(/\n/g, "\r")
            .replace(/\r{2,}/g, "\r")
            .replace(/@#/g, "\n");

        app.doScript(function() {
            if (validSelection) {
                replaceTextInSelection(userInput);
            } else {
                // 選択がない場合は新規テキストフレームを作成（入力長に応じてサイズ可変）
                var doc = app.activeDocument;
                var textLength = userInput.length;
                var width = Math.min(Math.max(textLength * 12, 100), 800); // 最小100、最大800
                var height = 20;

                // 現在のページの表示範囲の中心座標を取得
                var bounds = app.activeWindow.activePage.bounds; // [y1, x1, y2, x2]
                var centerY = (bounds[0] + bounds[2]) / 2;
                var centerX = (bounds[1] + bounds[3]) / 2;

                var left = centerX - width / 2;
                var top = centerY - height / 2;

                var tf = doc.textFrames.add();
                tf.geometricBounds = [top, left, top + height, left + width / 2];
                tf.contents = userInput;
            }
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.FAST_ENTIRE_SCRIPT);
    }

    // 挿入ポイントか選択テキストを処理対象とする
    function replaceTextInSelection(inputText) {
        try {
            // 選択状態を確認（1つのみ選択されているか）
            if (!app.selection || app.selection.length !== 1) {
                alert(LABELS.errorNoTextFrame[lang]);
                return;
            }

            var selectedObj = app.selection[0];

            // 選択が Text オブジェクトで、内容があれば置換
            if (selectedObj.hasOwnProperty("contents") && selectedObj.contents !== "") {
                selectedObj.contents = inputText;
                return;
            }

            // TextFrame が選択されている場合（最後尾の挿入ポイントに挿入）
            if (selectedObj instanceof TextFrame && selectedObj.parentStory) {
                selectedObj.parentStory.insertionPoints[-1].contents = inputText;
                return;
            }

            // 挿入ポイントのみが選択されている場合
            if (selectedObj.hasOwnProperty("insertionPoints")) {
                selectedObj.insertionPoints[0].contents = inputText;
                return;
            }

            alert(LABELS.errorNoTextFrame[lang]);
        } catch (e) {
            alert(LABELS.errorOccurred[lang] + e);
        }
    }
}

// スクリプトのエントリポイント
main();