#target indesign

/*

### 概要

選択した複数のテキストフレームを、選択順に連結して1つのストーリーにします。連結後の処理はダイアログで選べます。

詳細は README を参照してください。

### Overview

Links the selected text frames, in selection order, so that they share a single story. What happens after linking is chosen in a dialog.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTextFrameLinker";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2023-12-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-20";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTextFrameLinker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTextFrameLinker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n04ceaf4955a0"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// 初期設定 / Default settings
// =========================================
var DEFAULT_DELETE_EMPTY_FRAME = false;  /* 空になったフレームを削除 / delete the emptied frame */
var DEFAULT_FIT_FIRST_HEIGHT   = false;  /* 1つ目のフレームの高さを調整 / fit the first frame height */

// =========================================
// レイアウト / Layout
// =========================================
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */
var BUTTON_SPACING = 8;                  /* ボタン間隔 / button spacing */

(function () {

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
        dialog: {
            title: { ja: "テキストフレームを連結", en: "Link Text Frames" }
        },
        panel: {
            afterLinking: { ja: "連結後の処理", en: "After Linking" }
        },
        checkbox: {
            fitFirstHeight:   { ja: "1つ目のテキストフレームの高さを調整", en: "Fit the height of the first text frame" },
            deleteEmptyFrame: { ja: "空になったテキストフレームを削除", en: "Delete the emptied text frame" }
        },
        tooltip: {
            fitFirstHeight:   { ja: "幅はそのままに、テキストが収まる高さまで1つ目のフレームを広げます。", en: "Grows the first frame vertically until the text fits, keeping its width." },
            deleteEmptyFrame: { ja: "連結後に空になったフレームだけを削除します。1つ目のフレームは残します。", en: "Deletes only the frames left empty after linking. The first frame is always kept." }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument:    { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            needTwoFrames: { ja: "2つ以上のテキストフレームを選択してください。", en: "Please select two or more text frames." },
            cycle:         { ja: "連結すると循環するため実行できません。選択の順序を確認してください。", en: "Linking these frames would create a circular thread." },
            errorOccurred: { ja: "エラーが発生しました: ", en: "An error occurred: " }
        },
        undo: {
            linkFrames: { ja: "テキストフレームを連結", en: "Link Text Frames" }
        }
    };

    /**
     * ラベルを現在の言語で取得する
     * @param {object} labelEntry ja / en を持つラベルオブジェクト
     * @returns {string} 現在の言語のラベル文字列
     */
    function getLabel(labelEntry) {
        return labelEntry[currentLanguage];
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ウィンドウの共通設定を適用する
     * @param {Window} targetWindow 対象ウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

    /**
     * パネルの共通設定を適用する
     * @param {Panel} targetPanel 対象パネル
     * @returns {void}
     */
    function setupPanel(targetPanel) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = PANEL_SPACING;
    }

    /**
     * 行グループの共通設定を適用する（ボタン列など）
     * @param {Group} targetGroup 対象グループ
     * @param {string} [alignment] 横方向の配置。省略時は "left"
     * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
     * @returns {void}
     */
    function setupRow(targetGroup, alignment, spacing) {
        targetGroup.orientation = "row";
        targetGroup.alignment = [alignment || "left", "center"];  /* 横と天地を対で / Pair horizontal with vertical */
        targetGroup.alignChildren = ["left", "center"];           /* 親の fill 継承を打ち消す / Cancel the inherited fill */
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 連結後の処理を選ぶダイアログを表示する
     * @returns {?{fitFirstHeight: boolean, deleteEmptyFrame: boolean}} キャンセル時は null
     */
    function showLinkDialog() {
        var linkDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(linkDialog);

        var afterLinkingPanel = linkDialog.add("panel", undefined, getLabel(LABELS.panel.afterLinking));
        setupPanel(afterLinkingPanel);

        var fitFirstHeightCheckbox = afterLinkingPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.fitFirstHeight));
        var deleteEmptyFrameCheckbox = afterLinkingPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.deleteEmptyFrame));
        fitFirstHeightCheckbox.helpTip = getLabel(LABELS.tooltip.fitFirstHeight);
        deleteEmptyFrameCheckbox.helpTip = getLabel(LABELS.tooltip.deleteEmptyFrame);
        fitFirstHeightCheckbox.value = DEFAULT_FIT_FIRST_HEIGHT;
        deleteEmptyFrameCheckbox.value = DEFAULT_DELETE_EMPTY_FRAME;

        /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
        var btnRowGroup = linkDialog.add("group");
        setupRow(btnRowGroup, "right", BUTTON_SPACING);
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        if (linkDialog.show() !== 1) {
            return null;
        }

        return {
            fitFirstHeight: (fitFirstHeightCheckbox.value === true),
            deleteEmptyFrame: (deleteEmptyFrameCheckbox.value === true)
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択からテキストフレームだけを取り出す
     * @param {Array} selectedItems 選択オブジェクトの配列
     * @returns {?Array<TextFrame>} 2つ以上すべてテキストフレームなら配列、そうでなければ null
     */
    function getSelectedTextFrames(selectedItems) {
        if (selectedItems.length < 2) {
            return null;
        }
        var frames = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (!(selectedItems[i] instanceof TextFrame)) {
                return null;
            }
            frames.push(selectedItems[i]);
        }
        return frames;
    }

    /**
     * 連結先のフレームを返す（予定の連結を優先し、なければ現在の連結先）
     * @param {TextFrame} frame 対象のテキストフレーム
     * @param {object} plannedLinks 連結元のid をキーに連結先フレームを持つ対応表
     * @returns {?TextFrame} 連結先。なければ null
     */
    function nextFrameOf(frame, plannedLinks) {
        var key = String(frame.id);
        if (plannedLinks[key]) {
            return plannedLinks[key];
        }
        return (frame.nextTextFrame instanceof TextFrame) ? frame.nextTextFrame : null;
    }

    /**
     * 連結すると循環になるかを、実際に連結する前に判定する
     * @param {Array<TextFrame>} frames 選択順のテキストフレーム
     * @returns {boolean} 循環になるなら true
     */
    function wouldCycle(frames) {
        /* 予定の連結を既存の連結に重ねた状態で辿る / Walk the existing thread with the planned links applied */
        var plannedLinks = {};
        for (var i = 0; i < frames.length - 1; i++) {
            plannedLinks[String(frames[i].id)] = frames[i + 1];
        }
        for (var j = 0; j < frames.length; j++) {
            var visitedIds = {};
            var frame = frames[j];
            while (frame) {
                var key = String(frame.id);
                if (visitedIds[key]) {
                    return true;
                }
                visitedIds[key] = true;
                frame = nextFrameOf(frame, plannedLinks);
            }
        }
        return false;
    }

    /**
     * フレームに文字が入っているかを判定する
     * @param {TextFrame} frame 対象のテキストフレーム
     * @returns {boolean} 文字があれば true
     */
    function frameHasText(frame) {
        return String(frame.contents).length > 0;
    }

    /**
     * フレームの高さを、テキストが収まるまで広げる（幅・上端はそのまま）
     * @param {TextFrame} frame 対象のテキストフレーム
     * @returns {void}
     */
    function fitFrameHeight(frame) {
        var bounds = frame.geometricBounds;
        frame.fit(FitOptions.FRAME_TO_CONTENT);
        var fittedBounds = frame.geometricBounds;
        /* 幅は変えずに高さだけ反映する / Apply the fitted height while keeping the width */
        frame.geometricBounds = [bounds[0], bounds[1], fittedBounds[2], bounds[3]];
    }

    /**
     * 選択したテキストフレームを選択順に連結する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var frames = getSelectedTextFrames(app.activeDocument.selection);
        if (!frames) {
            alert(getLabel(LABELS.alert.needTwoFrames));
            return;
        }

        /* 循環はInDesignが例外を投げるため、事前に判定して案内する / InDesign throws on a circular thread, so report it up front */
        if (wouldCycle(frames)) {
            alert(getLabel(LABELS.alert.cycle));
            return;
        }

        var linkSettings = showLinkDialog();
        if (!linkSettings) {
            return;
        }

        try {
            /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
            app.doScript(function () {
                for (var i = 0; i < frames.length - 1; i++) {
                    frames[i].nextTextFrame = frames[i + 1];
                }

                if (linkSettings.fitFirstHeight) {
                    fitFrameHeight(frames[0]);
                }

                /* 1つ目は残し、文字が残っているフレームも削除しない / Keep the first frame, and never delete a frame that still holds text */
                if (linkSettings.deleteEmptyFrame) {
                    for (var j = frames.length - 1; j >= 1; j--) {
                        if (!frameHasText(frames[j])) {
                            frames[j].remove();
                        }
                    }
                }
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undo.linkFrames));
        } catch (e) {
            alert(getLabel(LABELS.alert.errorOccurred) + e);
        }
    }

    main();

})();
