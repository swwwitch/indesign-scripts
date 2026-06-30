//=========================================================
/*

### 概要

- 選択したテキストフレーム内の各段落を、段落ごとに独立したテキストフレームへ分割します。
- 元段落のベースライン位置を基準に Y 位置を補正します。
- 元フレームの左右座標を維持し、幅を保持します。
- テキストフレーム設定（textFramePreferences）を引き継ぎます。
- オーバーセットテキストがある場合は確認ダイアログを表示します。
- 必要に応じてフレーム高さを自動調整してオーバーセットを解消します。

仕様上の注意：

- 空段落（空行）は出力対象から除外されます。
- 想定用途：見出しや段落を個別オブジェクトとして配置・アニメーション・レイアウト調整するための前処理。

### Overview

- Splits each paragraph in the selected text frame into its own independent text frame.
- Corrects the Y position based on the original paragraph baseline.
- Keeps the original left/right coordinates so the width is preserved.
- Inherits the text frame preferences (textFramePreferences).
- Shows a confirmation dialog when overset text exists.
- Optionally grows the frame height to resolve overset text.

Notes:

- Empty paragraphs (blank lines) are excluded from the output.
- Intended as a pre-process for placing, animating, or laying out headings and paragraphs as separate objects.

*/
//=========================================================

#target indesign

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

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
                ja: "そのまま実行する（オーバーセットは消失します）",
                en: "Run anyway (overset text will be lost)"
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
    // オーバーセット処理 / Overset handling
    // =========================================
    /* オーバーセットがあれば処理方法を確認し、必要なら拡張して解消 / Confirm overset handling, expand to resolve if requested */
    if (sourceFrame.overflows) {
        var overflowChoice = askOverflowHandling();
        if (overflowChoice === "cancel") {
            return;
        }
        if (overflowChoice === "expand") {
            resolveOverflowByExpandingFrame(sourceFrame);
            if (sourceFrame.overflows) {
                alert(L("error.overflowFail"));
                return;
            }
        }
    }

    /* オーバーセット処理方法をダイアログで確認 / Ask the user how to handle overset text */
    function askOverflowHandling() {
        var overflowDialog = new Window("dialog", L("dialog.overflowTitle") + " " + SCRIPT_VERSION);
        overflowDialog.orientation = "column";
        overflowDialog.alignChildren = ["fill", "top"];
        overflowDialog.margins = [15, 15, 15, 10];

        overflowDialog.add("statictext", undefined, L("message.overflowPrompt"));

        /* 処理方法の選択パネル / Panel for choosing the handling method */
        var overflowOptionPanel = overflowDialog.add("panel");
        overflowOptionPanel.orientation = "column";
        overflowOptionPanel.alignChildren = ["left", "top"];
        overflowOptionPanel.margins = [12, 12, 12, 10];

        var expandFrameRadio = overflowOptionPanel.add("radiobutton", undefined, L("radio.expandFrame"));
        var ignoreOversetRadio = overflowOptionPanel.add("radiobutton", undefined, L("radio.ignoreOverset"));
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

    /* フレーム高さを少しずつ広げてオーバーセットを解消 / Grow frame height step by step to clear overset */
    function resolveOverflowByExpandingFrame(targetFrame) {
        var growIterations = 0;

        while (targetFrame.overflows && growIterations < MAX_GROW_ITERATIONS) {
            var currentBounds = targetFrame.geometricBounds;
            targetFrame.geometricBounds = [currentBounds[0], currentBounds[1], currentBounds[2] + FRAME_GROW_STEP_PT, currentBounds[3]];
            growIterations++;
        }
    }

    // =========================================
    // 段落分割 / Split paragraphs
    // =========================================
    var parentContainer = sourceFrame.parent;
    var paragraphList = sourceFrame.paragraphs.everyItem().getElements();

    /* 元フレームの座標 [上, 左, 下, 右] / Original frame bounds [top, left, bottom, right] */
    var sourceBounds = sourceFrame.geometricBounds;
    var sourceLeft = sourceBounds[1];  /* 元フレームの左端 / Original left edge */
    var sourceRight = sourceBounds[3]; /* 元フレームの右端 / Original right edge */

    for (var i = 0; i < paragraphList.length; i++) {
        var paragraph = paragraphList[i];

        /* 空行はスキップ（仕様） / Skip blank lines (by design) */
        if (paragraph.contents.replace(/\s/g, '') === "") {
            continue;
        }

        /* 念のため、行を持たない段落はスキップ / Skip paragraphs that have no lines */
        if (paragraph.lines.length === 0) {
            continue;
        }

        /* 元テキストの正確な Y 座標（ベースライン） / Exact original Y (baseline) */
        var sourceBaselineY = paragraph.lines[0].baseline;

        /* 新しいテキストフレームを作成し、設定を引き継ぐ / Create a new frame and inherit preferences */
        var newFrame = parentContainer.textFrames.add();
        newFrame.textFramePreferences.properties = sourceFrame.textFramePreferences.properties;

        /* 一旦、元フレームと全く同じサイズ・位置にする / Match the original size and position first */
        newFrame.geometricBounds = sourceBounds;

        /* 段落テキストを複製 / Duplicate the paragraph text */
        paragraph.duplicate(LocationOptions.AT_BEGINNING, newFrame.insertionPoints.item(0));

        /* 余分な空段落や改行を削除 / Remove the extra trailing paragraph and return */
        if (newFrame.paragraphs.length > 1) {
            newFrame.paragraphs.item(-1).remove();
        }
        if (newFrame.characters.length > 0 && newFrame.characters.item(-1).contents === "\r") {
            newFrame.characters.item(-1).remove();
        }

        /* いったんコンテンツに合わせる（高さと幅が縮む） / Fit to content (height and width shrink) */
        newFrame.fit(FitOptions.FRAME_TO_CONTENT);

        if (newFrame.lines.length > 0) {
            /* 縮んだ後の Y 座標とのズレを計算 / Compute the offset against the shrunk Y */
            var fittedBaselineY = newFrame.lines[0].baseline;
            var dy = sourceBaselineY - fittedBaselineY;

            var fittedBounds = newFrame.geometricBounds;

            /* Y はズレを補正し、X は元フレーム幅へ戻す / Correct Y by the offset, restore X to the original width */
            newFrame.geometricBounds = [
                fittedBounds[0] + dy,
                sourceLeft,
                fittedBounds[2] + dy,
                sourceRight
            ];

            /* 幅復元で再改行され、高さ不足になるケースを救済 / Rescue height shortage caused by reflow after width restore */
            var reflowIterations = 0;
            while (newFrame.overflows && reflowIterations < MAX_REFLOW_ITERATIONS) {
                var currentBounds = newFrame.geometricBounds;
                newFrame.geometricBounds = [currentBounds[0], currentBounds[1], currentBounds[2] + FRAME_GROW_STEP_PT, currentBounds[3]];
                reflowIterations++;
            }
        } else {
            newFrame.remove();
        }
    }

    /* 元のテキストフレームを削除 / Remove the original text frame */
    sourceFrame.remove();

})();
