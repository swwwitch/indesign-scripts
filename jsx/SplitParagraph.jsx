//=========================================================
/*

### 概要

- 選択したテキストフレーム内の各段落を、段落ごとに独立したテキストフレームへ分割します。
- 元段落のベースライン位置を基準に Y 位置を補正します。
- 元フレームの左右座標を維持し、幅を保持します。
- テキストフレーム設定（textFramePreferences）を引き継ぎます。
- オーバーセットテキストがある場合は確認ダイアログを表示し、「拡張して解消」または「そのまま実行」を選べます。
- 「拡張して解消」を選ぶと、フレーム高さを自動調整してオーバーセットを解消してから分割します。
- 一連の処理は1回の取り消し（編集 > 取り消し）でまとめて元に戻せます。

仕様上の注意：

- 空段落（空行）は出力対象から除外されます。
- 「そのまま実行」を選んだ場合、あふれて隠れているテキストは出力されません（失われます）。
- 想定用途：見出しや段落を個別オブジェクトとして配置・アニメーション・レイアウト調整するための前処理。

### Overview

- Splits each paragraph in the selected text frame into its own independent text frame.
- Corrects the Y position based on the original paragraph baseline.
- Keeps the original left/right coordinates so the width is preserved.
- Inherits the text frame preferences (textFramePreferences).
- Shows a confirmation dialog when overset text exists, offering "Expand to resolve" or "Run anyway".
- When "Expand to resolve" is chosen, grows the frame height to clear the overset before splitting.
- The whole operation can be reverted in a single undo (Edit > Undo).

Notes:

- Empty paragraphs (blank lines) are excluded from the output.
- When "Run anyway" is chosen, hidden overset text is not output (it is lost).
- Intended as a pre-process for placing, animating, or laying out headings and paragraphs as separate objects.

*/
//=========================================================

#target indesign

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.1.0";

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var FRAME_GROW_STEP_PT = 12;     /* フレームを広げる1回あたりの量(pt) / Height growth per step (pt) */
    var MAX_GROW_ITERATIONS = 200;   /* オーバーセット解消の最大試行回数 / Max iterations to resolve overset */
    var MAX_REFLOW_ITERATIONS = 20;  /* 幅復元後の再改行救済の最大試行回数 / Max iterations after width restore */

    // =========================================
    // ローカライズ / Localization
    // =========================================
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    var LABELS = {
        dialog: {
            overflowTitle: { ja: "オーバーセットテキストの確認", en: "Overset Text Detected" }
        },
        message: {
            overflowPrompt: {
                ja: "オーバーセットテキストがあります。処理方法を選択してください。",
                en: "The text frame contains overset text. Choose how to proceed."
            }
        },
        radio: {
            expandFrame: { ja: "フレームを拡張して解消する", en: "Expand frame to resolve" },
            ignoreOverset: {
                ja: "そのまま実行する（あふれたテキストは失われます）",
                en: "Run anyway (overset text will be lost)"
            }
        },
        tooltip: {
            expandFrame: {
                ja: "各段落を分割する前にフレームの高さを広げ、隠れているテキストをすべて表示してから処理します。",
                en: "Grows the frame height to reveal all hidden text before splitting each paragraph."
            },
            ignoreOverset: {
                ja: "現在表示されている段落だけを分割します。あふれて隠れているテキストは出力されません。",
                en: "Splits only the currently visible paragraphs. Hidden overset text will not be output."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        error: {
            docRequired: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectFrame: {
                ja: "テキストフレームを1つだけ選択して実行してください。",
                en: "Select exactly one text frame before running."
            },
            overflowFail: {
                ja: "オーバーセットテキストを解消できなかったため中断しました。",
                en: "Could not resolve overset text. Process cancelled."
            }
        },
        undo: {
            splitParagraphs: { ja: "段落ごとにフレーム分割", en: "Split Paragraphs Into Frames" }
        }
    };

    /* ドット区切りのキーでラベルを取得 / Look up a label by dotted key path */
    function L(keyPath) {
        var keys = keyPath.split(".");
        var node = LABELS;
        for (var k = 0; k < keys.length; k++) {
            node = node[keys[k]];
        }
        return node[currentLanguage];
    }

    // =========================================
    // 前提チェック / Preconditions
    // =========================================
    if (app.documents.length === 0) {
        alert(L("error.docRequired"));
        return;
    }

    if (app.selection.length !== 1) {
        alert(L("error.selectFrame"));
        return;
    }

    var sourceFrame = app.selection[0];
    if (!sourceFrame || sourceFrame.constructor.name !== "TextFrame") {
        alert(L("error.selectFrame"));
        return;
    }

    // =========================================
    // 元フレーム情報 / Source frame info
    // =========================================
    var parentContainer = sourceFrame.parent;

    /* 元フレームの座標 [上, 左, 下, 右] / Original frame bounds [top, left, bottom, right] */
    var sourceBounds = sourceFrame.geometricBounds;
    var sourceLeft = sourceBounds[1];  /* 元フレームの左端 / Original left edge */
    var sourceRight = sourceBounds[3]; /* 元フレームの右端 / Original right edge */

    // =========================================
    // 補助関数 / Helpers
    // =========================================
    /* オーバーセットが解消するまでフレーム高さを少しずつ広げる / Grow frame height step by step until overset clears */
    function growFrameUntilFits(targetFrame, maxIterations) {
        var iterations = 0;
        while (targetFrame.overflows && iterations < maxIterations) {
            var bounds = targetFrame.geometricBounds;
            targetFrame.geometricBounds = [bounds[0], bounds[1], bounds[2] + FRAME_GROW_STEP_PT, bounds[3]];
            iterations++;
        }
    }

    /* オーバーセット処理方法をダイアログで確認 / Ask the user how to handle overset text */
    function askOverflowHandling() {
        var overflowDialog = new Window("dialog", L("dialog.overflowTitle") + " " + SCRIPT_VERSION);
        overflowDialog.orientation = "column";
        overflowDialog.alignChildren = ["fill", "top"];
        overflowDialog.margins = [15, 15, 15, 15];

        overflowDialog.add("statictext", undefined, L("message.overflowPrompt"));

        /* 処理方法の選択パネル / Panel for choosing the handling method */
        var overflowOptionPanel = overflowDialog.add("panel");
        overflowOptionPanel.orientation = "column";
        overflowOptionPanel.alignChildren = ["left", "top"];
        overflowOptionPanel.margins = [12, 12, 12, 10];

        var expandFrameRadio = overflowOptionPanel.add("radiobutton", undefined, L("radio.expandFrame"));
        var ignoreOversetRadio = overflowOptionPanel.add("radiobutton", undefined, L("radio.ignoreOverset"));
        expandFrameRadio.helpTip = L("tooltip.expandFrame");
        ignoreOversetRadio.helpTip = L("tooltip.ignoreOverset");
        expandFrameRadio.value = true;

        /* ボタン行（Cancel → OK） / Button row (Cancel → OK) */
        var dialogButtonRow = overflowDialog.add("group");
        dialogButtonRow.alignment = "right";
        var cancelButton = dialogButtonRow.add("button", undefined, L("button.cancel"), { name: "cancel" });
        var okButton = dialogButtonRow.add("button", undefined, "OK", { name: "ok" });

        if (overflowDialog.show() !== 1) {
            return "cancel";
        }
        return expandFrameRadio.value ? "expand" : "ignore";
    }

    /* 1段落を独立したフレームに分割して配置 / Split one paragraph into its own positioned frame */
    function createFrameForParagraph(paragraph) {
        /* 空行・行なしの段落はスキップ（仕様） / Skip blank or line-less paragraphs (by design) */
        if (paragraph.contents.replace(/\s/g, '') === "" || paragraph.lines.length === 0) {
            return;
        }

        /* 元テキストの正確な Y 座標（ベースライン） / Exact original Y (baseline) */
        var sourceBaselineY = paragraph.lines[0].baseline;

        /* 新しいフレームを作成し、設定とサイズ・位置を引き継ぐ / Create a new frame inheriting preferences, size, and position */
        var newFrame = parentContainer.textFrames.add();
        newFrame.textFramePreferences.properties = sourceFrame.textFramePreferences.properties;
        newFrame.geometricBounds = sourceBounds;

        /* 段落テキストを複製 / Duplicate the paragraph text */
        paragraph.duplicate(LocationOptions.AT_BEGINNING, newFrame.insertionPoints.item(0));

        /* 余分な空段落や改行を削除 / Remove the extra trailing paragraph and return character */
        if (newFrame.paragraphs.length > 1) {
            newFrame.paragraphs.item(-1).remove();
        }
        if (newFrame.characters.length > 0 && newFrame.characters.item(-1).contents === "\r") {
            newFrame.characters.item(-1).remove();
        }

        /* いったんコンテンツに合わせる（高さと幅が縮む） / Fit to content (height and width shrink) */
        newFrame.fit(FitOptions.FRAME_TO_CONTENT);

        /* 行が消えたフレームは破棄 / Discard the frame if it ends up with no lines */
        if (newFrame.lines.length === 0) {
            newFrame.remove();
            return;
        }

        /* Y はズレを補正し、X は元フレーム幅へ戻す / Correct Y by the offset, restore X to the original width */
        var dy = sourceBaselineY - newFrame.lines[0].baseline;
        var fittedBounds = newFrame.geometricBounds;
        newFrame.geometricBounds = [fittedBounds[0] + dy, sourceLeft, fittedBounds[2] + dy, sourceRight];

        /* 幅復元で再改行され、高さ不足になるケースを救済 / Rescue height shortage caused by reflow after width restore */
        growFrameUntilFits(newFrame, MAX_REFLOW_ITERATIONS);
    }

    /* 全段落をフレーム分割（doScript から1つの取り消しとして実行） / Split all paragraphs (run via doScript as a single undo) */
    function splitParagraphsIntoFrames() {
        /* 「拡張」選択時はオーバーセットを解消してから座標を再取得 / On "expand", clear overset then re-read bounds */
        if (overflowChoice === "expand") {
            growFrameUntilFits(sourceFrame, MAX_GROW_ITERATIONS);
            if (sourceFrame.overflows) {
                return "overflowFail";
            }
            sourceBounds = sourceFrame.geometricBounds;
            sourceLeft = sourceBounds[1];
            sourceRight = sourceBounds[3];
        }

        /* 各段落を独立フレームへ / Each paragraph into its own frame */
        var paragraphList = sourceFrame.paragraphs.everyItem().getElements();
        for (var i = 0; i < paragraphList.length; i++) {
            createFrameForParagraph(paragraphList[i]);
        }

        /* 元のテキストフレームを削除 / Remove the original text frame */
        sourceFrame.remove();
        return "ok";
    }

    // =========================================
    // 実行 / Run
    // =========================================
    /* オーバーセットがあれば先に処理方法を確認（ダイアログは取り消し対象外） / Ask handling first if overset (dialog stays outside undo) */
    var overflowChoice = "none";
    if (sourceFrame.overflows) {
        overflowChoice = askOverflowHandling();
        if (overflowChoice === "cancel") {
            return;
        }
    }

    /* 変更を1つの取り消しステップとして実行 / Run all changes as a single undo step */
    var resultStatus = app.doScript(
        splitParagraphsIntoFrames,
        ScriptLanguage.JAVASCRIPT,
        undefined,
        UndoModes.ENTIRE_SCRIPT,
        L("undo.splitParagraphs")
    );

    if (resultStatus === "overflowFail") {
        alert(L("error.overflowFail"));
    }

})();
