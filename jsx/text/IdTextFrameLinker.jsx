#target indesign

/*

### 概要

選択した複数のテキストフレームを、選択順に連結して1つのストーリーにします。連結後の処理はダイアログで選べます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTextFrameLinker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n04ceaf4955a0

### Overview

Links the selected text frames, in selection order, so that they share a single story. What happens after linking is chosen in a dialog.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTextFrameLinker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTextFrameLinker";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2023-12-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

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
// UIレイアウト（再利用パーツ） / UI layout (reusable)

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

/**
 * ウィンドウの共通設定
 * @param {Window} targetWindow - 対象のウィンドウ
 * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
 * @returns {void}
 */
function setupWindow(targetWindow, spacing) {
    targetWindow.orientation = "column";
    targetWindow.alignChildren = "fill";
    targetWindow.margins = WINDOW_MARGINS;
    targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
 * @param {Panel} targetPanel - 対象のパネル
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
 * @returns {void}
 */
function setupPanel(targetPanel, spacing) {
    targetPanel.orientation = "column";
    targetPanel.alignChildren = ["fill", "top"];
    targetPanel.alignment = "fill";
    targetPanel.margins = PANEL_MARGINS;
    targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * タブの共通設定
 * @param {Tab} targetTab - 対象のタブ
 * @param {number} [spacing] - 要素間隔（省略時は変えない）
 * @returns {void}
 */
function setupTab(targetTab, spacing) {
    targetTab.orientation = "column";
    targetTab.alignChildren = "fill";
    targetTab.margins = TAB_MARGINS;
    if (typeof spacing === "number") targetTab.spacing = spacing;
}

/**
 * 横並びの行グループの共通設定（ボタン列など）。
 * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
 * @param {Group} rowGroup - 対象のグループ
 * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
 * @returns {void}
 */
function setupRow(rowGroup, rowAlignment, spacing) {
    rowGroup.orientation = "row";
    rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
    rowGroup.alignChildren = ["left", "center"];
    rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
 * @param {Button} targetButton - 対象のボタン
 * @param {number} trimPixels - 詰める量（px）
 * @returns {void}
 */
function trimButtonHeight(targetButton, trimPixels) {
    /* レイアウト前は size が無い / size is not set until the layout runs */
    if (!targetButton.size) return;
    targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
}

// UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

// ボタン行（再利用パーツ） / Button row (reusable)

var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

/**
 * ダイアログ下部のボタン行を作る。
 * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
 * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
 * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
 * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
 */
function addButtonRow(parent, rowOptions) {
    var isCentered = !!(rowOptions && rowOptions.centered);
    var btnRowGroup = parent.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
    btnRowGroup.spacing = BUTTON_ROW_SPACING;

    if (isCentered) {
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
    }

    btnRowGroup.alignment = ["fill", "bottom"];

    var btnLeftGroup = btnRowGroup.add("group");
    btnLeftGroup.alignChildren = ["left", "center"];
    btnLeftGroup.spacing = BUTTON_ROW_SPACING;

    /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
    var spacer = btnRowGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 0;

    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.alignChildren = ["right", "center"];
    btnRightGroup.spacing = BUTTON_ROW_SPACING;

    return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
}

/**
 * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
 * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
 * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
 * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
 * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
 * @returns {void}
 */
function alignRightOnlyButtonRow(buttonRow) {
    if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
    var dialogWindow = buttonRow.rowGroup.window;
    dialogWindow.addEventListener("show", function () {
        if (!buttonRow.leftGroup) return;
        var btnRowGroup = buttonRow.rowGroup;
        /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
        if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
        /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
        btnRowGroup.remove(buttonRow.leftGroup);
        btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        buttonRow.leftGroup = null;
        dialogWindow.layout.layout(true);
    });
}

// ボタン行（再利用パーツ）ここまで / End of the reusable button row

(function () {

    // =========================================
    // ラベル定義 / Labels
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

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

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 連結後の処理を選ぶダイアログを表示する
     * @returns {?{fitFirstHeight: boolean, deleteEmptyFrame: boolean}} キャンセル時は null
     */
    function showLinkDialog() {
        var linkDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(linkDialog);

        var afterLinkingPanel = linkDialog.add("panel", undefined, getLabel(LABELS.panel.afterLinking));
        setupPanel(afterLinkingPanel, 6);

        var fitFirstHeightCheckbox = afterLinkingPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.fitFirstHeight));
        var deleteEmptyFrameCheckbox = afterLinkingPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.deleteEmptyFrame));
        fitFirstHeightCheckbox.helpTip = getLabel(LABELS.tooltip.fitFirstHeight);
        deleteEmptyFrameCheckbox.helpTip = getLabel(LABELS.tooltip.deleteEmptyFrame);
        fitFirstHeightCheckbox.value = DEFAULT_FIT_FIRST_HEIGHT;
        deleteEmptyFrameCheckbox.value = DEFAULT_DELETE_EMPTY_FRAME;

        /* ボタン行 / Button row */
        var buttonRow = addButtonRow(linkDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

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
