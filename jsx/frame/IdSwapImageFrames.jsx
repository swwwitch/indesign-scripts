#target indesign

/*

### 概要

選択した画像入りフレームを、リンク画像だけ／フレームの位置ごと、いずれかの方法で順送りに入れ替えます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSwapImageFrames.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n6dee03ae96e2

### Overview

Rotates the selected image frames, moving either the linked images alone or the frames together with their positions.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSwapImageFrames.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdSwapImageFrames";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSwapImageFrames.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSwapImageFrames.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n6dee03ae96e2"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 同じ行とみなす垂直方向のズレの許容値 / Vertical tolerance that still counts as the same row */
var SAME_ROW_TOLERANCE = 12;

/* 画像フレームとして扱うオブジェクト種別 / Object types treated as image frames */
var IMAGE_FRAME_TYPES = ["Rectangle", "Oval", "Polygon"];

// ==============================
// UIレイアウトの共通設定 / Shared UI layout
// ==============================

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

var LABELS = {
    dialog: {
        title: { ja: "フレーム／画像を順送り", en: "Rotate Frames / Images" }
    },
    panel: {
        swapMode:    { ja: "入れ替え方法", en: "Swap Method" },
        frameAnchor: { ja: "位置合わせの基準", en: "Alignment Reference" },
        fitOption:   { ja: "配置後のフィット", en: "Fit After Placing" }
    },
    radio: {
        swapByFrame:       { ja: "フレームごと（位置を移動）", en: "Move the frames" },
        swapGraphicOnly:   { ja: "画像だけ入れ替え（フレームは固定）", en: "Move the images only" },
        anchorTopLeft:     { ja: "左上", en: "Top Left" },
        anchorCenter:      { ja: "中央", en: "Center" },
        fitFillProportion: { ja: "フレームに均等に流し込む", en: "Fill Frame Proportionally" },
        fitProportionally: { ja: "内容を縦横比率に応じて合わせる", en: "Fit Content Proportionally" }
    },
    tooltip: {
        swapByFrame: {
            ja: "画像はフレームに入れたまま、フレームの位置だけを順送りします。",
            en: "Keeps each graphic in its frame and rotates only the frame positions."
        },
        swapGraphicOnly: {
            ja: "フレームは動かさず、中のリンク画像だけを順送りします。",
            en: "Leaves every frame in place and rotates only the linked images."
        },
        frameAnchor: {
            ja: "「フレームごと（位置を移動）」のときだけ有効です。",
            en: "Available only when Move the frames is selected."
        },
        anchorTopLeft: {
            ja: "移動先フレームの左上に合わせます。サイズが違うと右下がずれます。",
            en: "Aligns with the top-left corner of the destination frame."
        },
        anchorCenter: {
            ja: "移動先フレームの中心に合わせます。サイズが違っても中心がそろいます。",
            en: "Aligns with the centre of the destination frame."
        },
        fitOption: {
            ja: "「画像だけ入れ替え（フレームは固定）」のときだけ有効です。",
            en: "Available only when Move the images only is selected."
        },
        fitFillProportion: {
            ja: "縦横比を保ったままフレーム全体を埋めます。はみ出した部分は隠れます。",
            en: "Fills the whole frame proportionally; anything outside is cropped."
        },
        fitProportionally: {
            ja: "縦横比を保ったまま画像全体を収めます。フレームに余白が出ることがあります。",
            en: "Fits the whole graphic proportionally; the frame may show empty space."
        }
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
        graphicRemoveError: {
            ja: "既存画像の削除に失敗したため、途中で中止しました。取り消しで元に戻せます。",
            en: "Stopped partway because an existing graphic could not be removed. Use undo to revert."
        },
        graphicPlaceError: {
            ja: "画像の再配置に失敗したため、途中で中止しました。取り消しで元に戻せます。",
            en: "Stopped partway because a linked image could not be placed. Use undo to revert."
        },
        graphicFitError: {
            ja: "画像を再配置しましたが、フィット処理に失敗したため中止しました。取り消しで元に戻せます。",
            en: "The graphic was replaced but could not be fitted, so the run stopped. Use undo to revert."
        },
        genericError: {
            ja: "予期しないエラーが発生したため、途中で中止しました。取り消しで元に戻せます。",
            en: "Stopped partway due to an unexpected error. Use undo to revert."
        }
    },
    undo: {
        swapFrames: { ja: "フレーム／画像の順送り", en: "Rotate Frames / Images" }
    }
};

// =========================================
// フレーム判定 / Frame detection
// =========================================

/**
 * 画像フレームとして扱う種別かを判定する
 * @param {string} typeName オブジェクトの種別名
 * @returns {boolean} 画像フレームとして扱う種別なら true
 */
function isImageFrameType(typeName) {
    for (var i = 0; i < IMAGE_FRAME_TYPES.length; i++) {
        if (IMAGE_FRAME_TYPES[i] === typeName) return true;
    }
    return false;
}

/**
 * フレームの位置情報（左上と中心）を取得する
 * @param {PageItem} frame 対象フレーム
 * @returns {{top: number, left: number, centerX: number, centerY: number}} 位置情報
 */
function getFrameBounds(frame) {
    var bounds = frame.geometricBounds;
    return {
        top: bounds[0],
        left: bounds[1],
        centerX: (bounds[1] + bounds[3]) / 2,
        centerY: (bounds[0] + bounds[2]) / 2
    };
}

/**
 * 選択項目から画像フレーム本体を取り出す
 * @param {object} selectionItem 選択項目（フレーム本体、またはフレーム内の配置画像）
 * @returns {PageItem|null} 画像フレーム。該当しない場合は null
 */
function getFrameFromSelectionItem(selectionItem) {
    if (!selectionItem) return null;
    if (isImageFrameType(selectionItem.constructor.name)) return selectionItem;

    /* 配置画像を直接選んでいる場合は親フレームへたどる / Walk up when the placed graphic itself is selected */
    var parentItem = selectionItem.parent;
    if (parentItem && isImageFrameType(parentItem.constructor.name)) return parentItem;

    return null;
}

/**
 * 見た目の位置（上から下、左から右）でフレームを並べ替える
 * @param {Array} frames 対象フレームの配列
 * @returns {Array} 並べ替えた配列
 */
function sortFramesByVisualOrder(frames) {
    var sortedFrames = frames.slice(0);

    sortedFrames.sort(function (frameA, frameB) {
        var boundsA = getFrameBounds(frameA);
        var boundsB = getFrameBounds(frameB);
        if (Math.abs(boundsA.centerY - boundsB.centerY) > SAME_ROW_TOLERANCE) {
            return boundsA.centerY - boundsB.centerY;
        }
        return boundsA.left - boundsB.left;
    });

    return sortedFrames;
}

/**
 * 選択項目から重複を除いた画像入りフレームの配列を作る
 * @param {Array} selectionItems 選択項目の配列
 * @returns {Array} 画像入りフレームの配列
 */
function collectImageFrames(selectionItems) {
    var imageFrames = [];
    var seenFrameIds = {};

    for (var i = 0; i < selectionItems.length; i++) {
        var frame = getFrameFromSelectionItem(selectionItems[i]);
        if (!frame) continue;
        if (frame.allGraphics.length === 0) continue;

        /* フレームと中の画像を同時に選んだときも1回だけ処理する / Handle a frame once even when its graphic is selected too */
        var frameId = String(frame.id);
        if (seenFrameIds[frameId]) continue;
        seenFrameIds[frameId] = true;

        imageFrames.push(frame);
    }

    return imageFrames;
}

// =========================================
// 入れ替え処理 / Swap operations
// =========================================

/**
 * 入れ替え処理の中断用エラーを作る（メッセージはローカライズ済み）
 * @param {string} labelKey アラート用ラベルのキー
 * @returns {Error} 表示用メッセージを持つエラー
 */
function createSwapError(labelKey) {
    var swapError = new Error(getLabel(labelKey));
    swapError.isSwapError = true;
    return swapError;
}

/**
 * フレーム内の主画像を取得する
 * @param {PageItem} frame 対象フレーム
 * @returns {object} 主画像。存在しない場合は null
 */
function getPrimaryGraphic(frame) {
    if (!frame || frame.allGraphics.length === 0) return null;
    return frame.allGraphics[0];
}

/**
 * 各フレームのリンク画像ファイルを並び順どおりに集める
 * @param {Array} imageFrames 対象フレームの配列
 * @returns {Array} リンクファイルの配列。1つでもリンクがない場合は null
 */
function collectLinkedFiles(imageFrames) {
    var linkedFiles = [];

    for (var i = 0; i < imageFrames.length; i++) {
        var primaryGraphic = getPrimaryGraphic(imageFrames[i]);
        var itemLink = primaryGraphic ? primaryGraphic.itemLink : null;
        if (!itemLink || !itemLink.filePath) return null;
        linkedFiles.push(File(itemLink.filePath));
    }

    return linkedFiles;
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
    if (!frame || !imageFile) throw createSwapError("alert.graphicPlaceError");

    try {
        removeExistingGraphics(frame);
    } catch (e) {
        throw createSwapError("alert.graphicRemoveError");
    }

    var placedItems = frame.place(imageFile);
    if (!placedItems || placedItems.length === 0) throw createSwapError("alert.graphicPlaceError");

    try {
        frame.fit(fitOption);
    } catch (e) {
        throw createSwapError("alert.graphicFitError");
    }
}

/**
 * フレームを指定の位置情報に合わせて移動する
 * @param {PageItem} frame 対象フレーム
 * @param {object} targetBounds 移動先の位置情報
 * @param {string} anchorMode "center" または "topLeft"
 * @returns {void}
 */
function moveFrameToBounds(frame, targetBounds, anchorMode) {
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

/**
 * 1つ後ろへずらす順送りを実行する（末尾は先頭へ回す）
 * @param {number} itemCount 対象件数
 * @param {function} applyRotation 移動先indexと取得元indexを受け取る処理
 * @returns {void}
 */
function rotateInOrder(itemCount, applyRotation) {
    for (var i = itemCount - 1; i > 0; i--) {
        applyRotation(i, i - 1);
    }
    applyRotation(0, itemCount - 1);
}

/**
 * フレームはそのままに、リンク画像だけを順送りする
 * @param {Array} imageFrames 対象フレームの配列
 * @param {Array} linkedFiles 実行前に控えたリンクファイルの配列
 * @param {FitOptions} fitOption 配置後のフィット方法
 * @returns {void}
 */
function rotateLinkedImages(imageFrames, linkedFiles, fitOption) {
    rotateInOrder(imageFrames.length, function (targetIndex, sourceIndex) {
        replaceFrameGraphic(imageFrames[targetIndex], linkedFiles[sourceIndex], fitOption);
    });
}

/**
 * 中身はそのままに、フレームの位置だけを順送りする
 * @param {Array} imageFrames 対象フレームの配列
 * @param {string} anchorMode "center" または "topLeft"
 * @returns {void}
 */
function rotateFramePositions(imageFrames, anchorMode) {
    var originalBoundsList = [];
    for (var i = 0; i < imageFrames.length; i++) {
        originalBoundsList.push(getFrameBounds(imageFrames[i]));
    }

    rotateInOrder(imageFrames.length, function (targetIndex, sourceIndex) {
        moveFrameToBounds(imageFrames[targetIndex], originalBoundsList[sourceIndex], anchorMode);
    });
}

// =========================================
// ダイアログ / Dialog
// =========================================

/**
 * 入れ替え方法を指定するダイアログを表示する
 * @returns {object} 設定内容 {swapMode, fitOption, anchorMode}。キャンセル時は null
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
    swapByFrameRadio.helpTip     = getLabel("tooltip.swapByFrame");
    swapGraphicOnlyRadio.helpTip = getLabel("tooltip.swapGraphicOnly");
    swapGraphicOnlyRadio.value = true;

    /* フレーム位置パネル / Frame-anchor panel */
    var frameAnchorPanel = swapDialog.add("panel", undefined, getLabel("panel.frameAnchor"));
    setupPanel(frameAnchorPanel, 6);
    frameAnchorPanel.alignChildren = ["left", "top"];
    frameAnchorPanel.helpTip = getLabel("tooltip.frameAnchor");

    var anchorTopLeftRadio = frameAnchorPanel.add("radiobutton", undefined, getLabel("radio.anchorTopLeft"));
    var anchorCenterRadio  = frameAnchorPanel.add("radiobutton", undefined, getLabel("radio.anchorCenter"));
    anchorTopLeftRadio.helpTip = getLabel("tooltip.anchorTopLeft");
    anchorCenterRadio.helpTip  = getLabel("tooltip.anchorCenter");
    anchorTopLeftRadio.value = true;

    /* フィットパネル / Fit-option panel */
    var fitOptionPanel = swapDialog.add("panel", undefined, getLabel("panel.fitOption"));
    setupPanel(fitOptionPanel, 6);
    fitOptionPanel.alignChildren = ["left", "top"];
    fitOptionPanel.helpTip = getLabel("tooltip.fitOption");

    var fitFillRadio         = fitOptionPanel.add("radiobutton", undefined, getLabel("radio.fitFillProportion"));
    var fitProportionalRadio = fitOptionPanel.add("radiobutton", undefined, getLabel("radio.fitProportionally"));
    fitFillRadio.helpTip         = getLabel("tooltip.fitFillProportion");
    fitProportionalRadio.helpTip = getLabel("tooltip.fitProportionally");
    fitFillRadio.value = true;

    /**
     * 入れ替えモードに応じて関連パネルの有効／無効を切り替える
     * @returns {void}
     */
    function updatePanelAvailability() {
        var isGraphicOnly = (swapGraphicOnlyRadio.value === true);

        fitOptionPanel.enabled       = isGraphicOnly;
        fitFillRadio.enabled         = isGraphicOnly;
        fitProportionalRadio.enabled = isGraphicOnly;

        frameAnchorPanel.enabled     = !isGraphicOnly;
        anchorTopLeftRadio.enabled   = !isGraphicOnly;
        anchorCenterRadio.enabled    = !isGraphicOnly;
    }

    swapByFrameRadio.onClick     = updatePanelAvailability;
    swapGraphicOnlyRadio.onClick = updatePanelAvailability;
    updatePanelAvailability();

    var buttonRow = addButtonRow(swapDialog);
    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    alignRightOnlyButtonRow(buttonRow);

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
    if (imageFrames.length < 2) {
        alert(getLabel("alert.selectFrames"));
        return;
    }

    try {
        imageFrames = sortFramesByVisualOrder(imageFrames);
    } catch (e) {
        alert(getLabel("alert.sortError"));
        return;
    }

    var swapSettings = showSwapDialog();
    if (!swapSettings) return;

    var isGraphicOnly = (swapSettings.swapMode === "graphicOnly");

    /* 画像リンク入れ替えでは、先に全リンクファイルを控えておく / For link swapping, capture every linked file first */
    var linkedFiles = null;
    if (isGraphicOnly) {
        linkedFiles = collectLinkedFiles(imageFrames);
        if (!linkedFiles) {
            alert(getLabel("alert.selectLinkedFrames"));
            return;
        }
    }

    try {
        /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
        app.doScript(function () {
            if (isGraphicOnly) {
                rotateLinkedImages(imageFrames, linkedFiles, swapSettings.fitOption);
            } else {
                rotateFramePositions(imageFrames, swapSettings.anchorMode);
            }
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.swapFrames"));
    } catch (e) {
        alert((e && e.isSwapError) ? e.message : getLabel("alert.genericError"));
    }

})();
