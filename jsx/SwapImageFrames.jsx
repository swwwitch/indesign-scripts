#target indesign

/*
 * SwapImageFrames.jsx
 *
 * 選択した画像入りフレームを、リンク画像だけ／フレーム位置ごとのいずれかの方法で順送りに入れ替えます。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SwapImageFrames";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-03-28";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/SwapImageFrames.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/SwapImageFrames.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 同じ行とみなす垂直方向のズレの許容値 / Vertical tolerance that still counts as the same row */
var ROW_TOLERANCE = 12;

/* 画像フレームとして扱うオブジェクト種別 / Object types treated as image frames */
var IMAGE_FRAME_TYPES = ["Rectangle", "Oval", "Polygon"];

/* フレームの中身として扱う配置画像の種別 / Placed-graphic types recognized inside a frame */
var PLACED_GRAPHIC_TYPES = ["Image", "PDF", "EPS", "ImportedPage"];

// ==============================
// UIレイアウトの共通設定 / Shared UI layout
// ==============================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */

/**
 * ウィンドウの共通設定を適用する
 * @param {Window} win 対象ウィンドウ
 * @param {number} [spacing] 要素間隔。省略時は WINDOW_SPACING
 * @returns {void}
 */
function setupWindow(win, spacing) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルの共通設定を適用する
 * @param {Panel} panel 対象パネル
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * 行グループの共通設定を適用する（ボタン列など）
 * @param {Group} group 対象グループ
 * @param {string} [alignment] 配置。省略時は "left"
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

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
    dialog: {
        title: { ja: "フレーム／画像の入れ替え", en: "Swap Frames / Images" }
    },
    panel: {
        swapMode:    { ja: "入れ替えモード", en: "Swap Mode" },
        frameAnchor: { ja: "フレームの位置", en: "Frame Position" },
        fitOption:   { ja: "配置後のフィット", en: "Fit After Placing" }
    },
    radio: {
        swapByFrame:       { ja: "フレームごと入れ替え", en: "Swap frames" },
        swapGraphicOnly:   { ja: "画像リンクを入れ替え", en: "Swap linked images" },
        anchorTopLeft:     { ja: "左上", en: "Top Left" },
        anchorCenter:      { ja: "中央", en: "Center" },
        fitFillProportion: { ja: "フレームに合わせて塗り（トリミングあり）", en: "Fill Proportionally (crop)" },
        fitProportionally: { ja: "全体を表示（余白あり）", en: "Fit Proportionally" }
    },
    button: {
        ok:     { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    alert: {
        noDocument:         { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        selectFrames:       { ja: "画像入りフレームを2つ以上選択してください。", en: "Select at least two frames containing placed images." },
        selectLinkedFrames: { ja: "リンク画像が入ったフレームを2つ以上選択してください。", en: "Select at least two frames containing linked images." },
        sortError:          { ja: "フレームの並び順を決定できませんでした。", en: "Could not determine the frame order." },
        graphicRemoveError: { ja: "既存画像の削除に失敗したため、画像リンクの入れ替えを中止しました。", en: "The swap was canceled because an existing graphic could not be removed." },
        graphicPlaceError:  { ja: "画像の再配置に失敗したため、画像リンクの入れ替えを中止しました。", en: "The swap was canceled because a linked image could not be placed." },
        graphicFitError:    { ja: "画像は再配置されましたが、フィット処理に失敗しました。", en: "The linked image was replaced, but the fit operation failed." },
        genericError:       { ja: "予期しないエラーが発生したため、処理を中止しました。", en: "The operation was canceled due to an unexpected error." }
    },
    undo: {
        swapFrames: { ja: "フレーム／画像の入れ替え", en: "Swap Frames / Images" }
    }
};

/**
 * ドット区切りキーでラベルを取得する
 * @param {string} labelKey 例: "dialog.title"
 * @returns {string} 現在の言語のラベル文字列。見つからない場合はキーをそのまま返す
 */
function getLabel(labelKey) {
    var node = LABELS;
    var keyParts = labelKey.split(".");
    for (var i = 0; i < keyParts.length; i++) {
        node = node[keyParts[i]];
        if (!node) return labelKey;
    }
    return node[currentLang] || node.en || labelKey;
}

// =========================================
// フレーム判定 / Frame detection
// =========================================

/**
 * 種別名がリストに含まれるかを判定する
 * @param {Array<string>} typeNames 判定に使う種別名の配列
 * @param {string} typeName 判定する種別名
 * @returns {boolean} 含まれていれば true
 */
function isTypeIn(typeNames, typeName) {
    for (var i = 0; i < typeNames.length; i++) {
        if (typeNames[i] === typeName) return true;
    }
    return false;
}

/**
 * フレームの座標情報をまとめて取得する
 * @param {PageItem} frame 対象フレーム
 * @returns {{top: number, left: number, bottom: number, right: number, width: number, height: number, centerX: number, centerY: number}} 座標情報
 */
function getFrameBounds(frame) {
    var bounds = frame.geometricBounds;
    return {
        top: bounds[0],
        left: bounds[1],
        bottom: bounds[2],
        right: bounds[3],
        width: bounds[3] - bounds[1],
        height: bounds[2] - bounds[0],
        centerX: (bounds[1] + bounds[3]) / 2,
        centerY: (bounds[0] + bounds[2]) / 2
    };
}

/**
 * 重複選択を除くためのフレーム識別キーを作る
 * @param {PageItem} frame 対象フレーム
 * @returns {string} 識別キー。取得できない場合は空文字
 */
function getFrameKey(frame) {
    if (!frame) return "";

    try {
        if (frame.id !== undefined) return String(frame.id);
    } catch (e) {}

    try {
        var bounds = frame.geometricBounds;
        return [frame.constructor.name, bounds[0], bounds[1], bounds[2], bounds[3]].join("|");
    } catch (e) {}

    return "";
}

/**
 * 選択項目から画像フレーム本体を取り出す
 * @param {object} selectionItem 選択項目
 * @returns {PageItem|null} 画像フレーム。該当しない場合は null
 */
function getFrameFromSelectionItem(selectionItem) {
    if (!selectionItem) return null;

    var typeName = selectionItem.constructor.name;
    if (isTypeIn(IMAGE_FRAME_TYPES, typeName)) return selectionItem;

    if (isTypeIn(PLACED_GRAPHIC_TYPES, typeName) && selectionItem.parent) {
        if (isTypeIn(IMAGE_FRAME_TYPES, selectionItem.parent.constructor.name)) return selectionItem.parent;
    }

    if (selectionItem.parent && isTypeIn(IMAGE_FRAME_TYPES, selectionItem.parent.constructor.name)) {
        return selectionItem.parent;
    }

    return null;
}

/**
 * 見た目の位置（上から下、左から右）でフレームを並べ替える
 * @param {Array<PageItem>} frames 対象フレームの配列
 * @returns {Array<PageItem>} 並べ替えた配列
 */
function sortFramesByVisualOrder(frames) {
    var sortedFrames = frames.slice(0);

    sortedFrames.sort(function (a, b) {
        var boundsA = getFrameBounds(a);
        var boundsB = getFrameBounds(b);
        if (Math.abs(boundsA.centerY - boundsB.centerY) > ROW_TOLERANCE) {
            return boundsA.centerY - boundsB.centerY;
        }
        return boundsA.left - boundsB.left;
    });

    return sortedFrames;
}

/**
 * 選択項目から重複を除いた画像フレームの配列を作る
 * @param {Array} selectionItems 選択項目の配列
 * @returns {Array<PageItem>} 画像入りフレームの配列
 */
function collectImageFrames(selectionItems) {
    var imageFrames = [];
    var seenFrameKeys = {};

    for (var i = 0; i < selectionItems.length; i++) {
        var frame = getFrameFromSelectionItem(selectionItems[i]);
        if (!frame) continue;
        if (frame.allGraphics.length === 0) continue;

        var frameKey = getFrameKey(frame);
        if (frameKey && seenFrameKeys[frameKey]) continue;
        if (frameKey) seenFrameKeys[frameKey] = true;

        imageFrames.push(frame);
    }

    return imageFrames;
}

// =========================================
// 入れ替え処理 / Swap operations
// =========================================

/**
 * フレーム内の主画像を取得する
 * @param {PageItem} frame 対象フレーム
 * @returns {Graphic|null} 主画像。存在しない場合は null
 */
function getPrimaryGraphic(frame) {
    if (!frame || frame.allGraphics.length === 0) return null;
    return frame.allGraphics[0];
}

/**
 * フレーム内の既存画像をすべて削除する
 * @param {PageItem} frame 対象フレーム
 * @returns {void}
 */
function removeExistingGraphics(frame) {
    if (!frame) return;
    while (frame.allGraphics.length > 0) {
        frame.allGraphics[0].remove();
    }
}

/**
 * フレームの画像を差し替えてフィットさせる
 * @param {PageItem} frame 対象フレーム
 * @param {File} imageFile 配置する画像ファイル
 * @param {FitOptions} fitOption 配置後のフィット方法
 * @returns {void}
 */
function replaceFrameGraphic(frame, imageFile, fitOption) {
    if (!frame || !imageFile) throw new Error("Invalid frame or file.");

    try {
        removeExistingGraphics(frame);
    } catch (e) {
        throw new Error(getLabel("alert.graphicRemoveError"));
    }

    var placedItems = frame.place(imageFile);
    if (!placedItems || placedItems.length === 0) throw new Error("Place failed.");

    try {
        frame.fit(fitOption);
    } catch (e) {
        throw new Error(getLabel("alert.graphicFitError"));
    }
}

/**
 * フレームを保存しておいた位置へ移動する
 * @param {PageItem} frame 対象フレーム
 * @param {object} targetBounds 移動先の座標情報
 * @param {string} anchorMode "center" または "topLeft"
 * @returns {void}
 */
function moveFrameToStoredPosition(frame, targetBounds, anchorMode) {
    if (!frame || !targetBounds) return;

    var currentBounds = getFrameBounds(frame);
    var dx, dy;

    if (anchorMode === "center") {
        dx = targetBounds.centerX - currentBounds.centerX;
        dy = targetBounds.centerY - currentBounds.centerY;
    } else {
        dx = targetBounds.left - currentBounds.left;
        dy = targetBounds.top - currentBounds.top;
    }

    frame.move(undefined, [dx, dy]);
}

// =========================================
// ダイアログ / Dialog
// =========================================

/**
 * 入れ替え方法を指定するダイアログを表示する
 * @returns {{swapMode: string, fitOption: FitOptions, anchorMode: string}|null} 設定内容。キャンセル時は null
 */
function showSwapDialog() {
    var swapDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(swapDialog);

    /* 入れ替えモードパネル / Swap-mode panel */
    var swapModePanel = swapDialog.add("panel", undefined, getLabel("panel.swapMode"));
    setupPanel(swapModePanel, 6);
    swapModePanel.alignChildren = ["left", "top"];

    var swapByFrameRadio     = swapModePanel.add("radiobutton", undefined, getLabel("radio.swapByFrame"));
    var swapGraphicOnlyRadio = swapModePanel.add("radiobutton", undefined, getLabel("radio.swapGraphicOnly"));
    swapGraphicOnlyRadio.value = true;

    /* フレーム位置パネル / Frame-anchor panel */
    var frameAnchorPanel = swapDialog.add("panel", undefined, getLabel("panel.frameAnchor"));
    setupPanel(frameAnchorPanel, 6);
    frameAnchorPanel.alignChildren = ["left", "top"];

    var anchorTopLeftRadio = frameAnchorPanel.add("radiobutton", undefined, getLabel("radio.anchorTopLeft"));
    var anchorCenterRadio  = frameAnchorPanel.add("radiobutton", undefined, getLabel("radio.anchorCenter"));
    anchorTopLeftRadio.value = true;

    /* フィットパネル / Fit-option panel */
    var fitOptionPanel = swapDialog.add("panel", undefined, getLabel("panel.fitOption"));
    setupPanel(fitOptionPanel, 6);
    fitOptionPanel.alignChildren = ["left", "top"];

    var fitFillRadio         = fitOptionPanel.add("radiobutton", undefined, getLabel("radio.fitFillProportion"));
    var fitProportionalRadio = fitOptionPanel.add("radiobutton", undefined, getLabel("radio.fitProportionally"));
    fitFillRadio.value = true;

    /**
     * 入れ替えモードに応じて関連パネルの有効／無効を切り替える
     * @returns {void}
     */
    function reflectDialogState() {
        var isGraphicOnly = (swapGraphicOnlyRadio.value === true);

        fitOptionPanel.enabled       = isGraphicOnly;
        fitFillRadio.enabled         = isGraphicOnly;
        fitProportionalRadio.enabled = isGraphicOnly;

        frameAnchorPanel.enabled     = !isGraphicOnly;
        anchorTopLeftRadio.enabled   = !isGraphicOnly;
        anchorCenterRadio.enabled    = !isGraphicOnly;
    }

    swapByFrameRadio.onClick     = reflectDialogState;
    swapGraphicOnlyRadio.onClick = reflectDialogState;
    reflectDialogState();

    /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
    var dialogButtonRow = swapDialog.add("group");
    setupRow(dialogButtonRow, "right", 8);
    dialogButtonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    dialogButtonRow.add("button", undefined, getLabel("button.ok"), { name: "ok" });

    if (swapDialog.show() !== 1) return null;

    return {
        swapMode: swapByFrameRadio.value ? "frame" : "graphicOnly",
        fitOption: fitFillRadio.value ? FitOptions.FILL_PROPORTIONALLY : FitOptions.PROPORTIONALLY,
        anchorMode: anchorTopLeftRadio.value ? "topLeft" : "center"
    };
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {

    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var selectionItems = app.selection;
    if (!selectionItems || selectionItems.length < 2) {
        alert(getLabel("alert.selectFrames"));
        return;
    }

    var imageFrames = collectImageFrames(selectionItems);

    if (imageFrames.length > 1) {
        try {
            imageFrames = sortFramesByVisualOrder(imageFrames);
        } catch (e) {
            alert(getLabel("alert.sortError"));
            return;
        }
    }

    if (imageFrames.length < 2) {
        alert(getLabel("alert.selectFrames"));
        return;
    }

    var dialogResult = showSwapDialog();
    if (dialogResult === null) return;

    var swapMode   = dialogResult.swapMode;
    var fitOption  = dialogResult.fitOption;
    var anchorMode = dialogResult.anchorMode;

    /* 画像リンク入れ替えでは、先に全リンクファイルを控えておく / For link swapping, capture every linked file first */
    var linkedFiles = [];
    if (swapMode === "graphicOnly") {
        for (var i = 0; i < imageFrames.length; i++) {
            var primaryGraphic = getPrimaryGraphic(imageFrames[i]);
            var itemLink = primaryGraphic ? primaryGraphic.itemLink : null;
            if (!itemLink || !itemLink.filePath) {
                alert(getLabel("alert.selectLinkedFrames"));
                return;
            }
            linkedFiles.push(File(itemLink.filePath));
        }
    }

    try {
        /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
        app.doScript(function () {
            if (swapMode === "graphicOnly") {
                var lastFile = linkedFiles[linkedFiles.length - 1];
                for (var i = linkedFiles.length - 1; i > 0; i--) {
                    replaceFrameGraphic(imageFrames[i], linkedFiles[i - 1], fitOption);
                }
                replaceFrameGraphic(imageFrames[0], lastFile, fitOption);
                return;
            }

            var originalBounds = [];
            for (var j = 0; j < imageFrames.length; j++) {
                originalBounds.push(getFrameBounds(imageFrames[j]));
            }

            var lastBounds = originalBounds[originalBounds.length - 1];
            for (var k = imageFrames.length - 1; k > 0; k--) {
                moveFrameToStoredPosition(imageFrames[k], originalBounds[k - 1], anchorMode);
            }
            moveFrameToStoredPosition(imageFrames[0], lastBounds, anchorMode);

        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.swapFrames"));
    } catch (e) {
        var errorMessage = (e && e.message) ? e.message : "";

        if (errorMessage === getLabel("alert.graphicFitError")) {
            alert(getLabel("alert.graphicFitError"));
        } else if (errorMessage === "Place failed." || errorMessage === "Invalid frame or file.") {
            alert(getLabel("alert.graphicPlaceError"));
        } else if (errorMessage === getLabel("alert.graphicRemoveError")) {
            alert(getLabel("alert.graphicRemoveError"));
        } else {
            alert(getLabel("alert.genericError"));
        }
    }

})();
