#target indesign

/*

### 概要

選択したオブジェクトを水平方向（行）または垂直方向（列）の近さでまとめてグループ化します。

詳細は README を参照してください。

### Overview

Groups the selected objects by proximity, either horizontally (rows) or vertically (columns).

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdSmartGroup";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSmartGroup.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSmartGroup.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 許容値スライダーの初期値・最小値・最大値 / Initial, minimum and maximum of the tolerance slider */
var TOLERANCE_DEFAULT = 5;
var TOLERANCE_MIN     = 0;
var TOLERANCE_MAX     = 50;

/* プレビュー用の一時レイヤー名とスウォッチ名 / Names of the temporary preview layer and swatch */
var PREVIEW_LAYER_NAME   = "SmartGroup Preview";
var PREVIEW_SWATCH_NAME  = "SmartGroup_Preview_Red";

/* プレビュー矩形の CMYK 値と不透明度 / CMYK value and opacity of the preview rectangle */
var PREVIEW_SWATCH_CMYK = [0, 100, 100, 0];
var PREVIEW_OPACITY     = 40;

// =========================================
// レイアウト設定 / Layout settings
// =========================================

/* 許容値スライダーと数値表示の幅（px）/ Width of the tolerance slider and its value readout (px) */
var TOLERANCE_SLIDER_WIDTH = 180;
var TOLERANCE_VALUE_WIDTH  = 30;

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

    var currentLang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "グループ化の設定", en: "Smart Group Settings" }
        },
        panel: {
            direction: { ja: "グループ化する方向", en: "Grouping direction" },
            tolerance: { ja: "許容値", en: "Tolerance" }
        },
        radio: {
            horizontal: { ja: "水平方向（横並び）", en: "Horizontal (rows)" },
            vertical:   { ja: "垂直方向（縦並び）", en: "Vertical (columns)" }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        direction: {
            horizontal: { ja: "水平方向", en: "horizontally" },
            vertical:   { ja: "垂直方向", en: "vertically" }
        },
        alert: {
            noSelection: { ja: "アイテムを選択してください。", en: "Please select one or more items." }
        },
        undo: {
            smartGroup: { ja: "スマートグループ化", en: "Smart Group" }
        }
    };

    /**
     * ラベルを現在の言語で取得する
     * @param {object} labelEntry ja / en を持つラベルオブジェクト
     * @returns {string} 現在の言語のラベル文字列
     */
    function localize(labelEntry) {
        return labelEntry[currentLang];
    }

    // =========================================
    // プレビュー用リソース / Preview resources
    // =========================================

    /* プレビュー矩形の参照を保持する / Holds references to the preview rectangles */
    var previewRectangles = [];

    /**
     * プレビュー用の赤スウォッチを取得する（なければ作成）
     * @param {Document} targetDoc 対象ドキュメント
     * @returns {Color} プレビュー用スウォッチ
     */
    function getPreviewSwatch(targetDoc) {
        for (var i = 0; i < targetDoc.colors.length; i++) {
            if (targetDoc.colors[i].name === PREVIEW_SWATCH_NAME) return targetDoc.colors[i];
        }
        return targetDoc.colors.add({
            name: PREVIEW_SWATCH_NAME,
            model: ColorModel.PROCESS,
            space: ColorSpace.CMYK,
            colorValue: PREVIEW_SWATCH_CMYK
        });
    }

    /**
     * プレビュー用の非印刷レイヤーを取得する（なければ作成）
     * @param {Document} targetDoc 対象ドキュメント
     * @returns {Layer} プレビュー用レイヤー
     */
    function getPreviewLayer(targetDoc) {
        for (var i = 0; i < targetDoc.layers.length; i++) {
            if (targetDoc.layers[i].name === PREVIEW_LAYER_NAME) return targetDoc.layers[i];
        }
        return targetDoc.layers.add({ name: PREVIEW_LAYER_NAME, printable: false });
    }

    /**
     * プレビューをレイヤーごと削除する
     * @param {Document} targetDoc 対象ドキュメント
     * @returns {void}
     */
    function clearPreview(targetDoc) {
        for (var i = 0; i < targetDoc.layers.length; i++) {
            if (targetDoc.layers[i].name === PREVIEW_LAYER_NAME) {
                try { targetDoc.layers[i].remove(); } catch (e) {}
                break;
            }
        }
        previewRectangles = [];
        try { app.redraw(); } catch (e) {}
    }

    // =========================================
    // グループ計算 / Group computation
    // =========================================

    /**
     * 選択オブジェクトの境界情報を取り出す
     * @param {Array} selectedItems 選択オブジェクトの配列
     * @returns {Array<{pageItem: PageItem, top: number, left: number, bottom: number, right: number}>} 境界情報つきの配列
     */
    function collectItemBounds(selectedItems) {
        var itemBounds = [];
        for (var i = 0; i < selectedItems.length; i++) {
            /* geometricBounds は [上, 左, 下, 右] / geometricBounds is [top, left, bottom, right] */
            var bounds = selectedItems[i].geometricBounds;
            itemBounds.push({
                pageItem: selectedItems[i],
                top: bounds[0],
                left: bounds[1],
                bottom: bounds[2],
                right: bounds[3]
            });
        }
        return itemBounds;
    }

    /**
     * 指定方向の中心座標を返す
     * @param {{top: number, left: number, bottom: number, right: number}} itemBound 境界情報
     * @param {string} direction "horizontal"（Y 中心）または "vertical"（X 中心）
     * @returns {number} 中心座標
     */
    function getCenterAlongAxis(itemBound, direction) {
        return (direction === "horizontal")
            ? (itemBound.top + itemBound.bottom) / 2
            : (itemBound.left + itemBound.right) / 2;
    }

    /**
     * 中心座標の近さでオブジェクトを行または列にまとめる
     * @param {Array} itemBounds 境界情報つきの配列
     * @param {string} direction "horizontal" または "vertical"
     * @param {number} tolerance 同じ行／列とみなす許容値
     * @returns {Array<Array>} まとめた結果
     */
    function computeGroups(itemBounds, direction, tolerance) {
        var sortedItems = itemBounds.slice();
        sortedItems.sort(function (a, b) {
            return getCenterAlongAxis(a, direction) - getCenterAlongAxis(b, direction);
        });

        var groups = [];
        var currentGroup = [];

        for (var i = 0; i < sortedItems.length; i++) {
            if (currentGroup.length === 0) {
                currentGroup.push(sortedItems[i]);
                continue;
            }
            var previousItem = currentGroup[currentGroup.length - 1];
            var centerDelta = Math.abs(
                getCenterAlongAxis(sortedItems[i], direction) - getCenterAlongAxis(previousItem, direction)
            );
            if (centerDelta <= tolerance) {
                currentGroup.push(sortedItems[i]);
            } else {
                groups.push(currentGroup);
                currentGroup = [sortedItems[i]];
            }
        }
        if (currentGroup.length > 0) groups.push(currentGroup);
        return groups;
    }

    /**
     * グループ全体を囲む境界を求める
     * @param {Array} groupItems 同じグループに属する境界情報の配列
     * @returns {Array<number>} [上, 左, 下, 右]
     */
    function getGroupBounds(groupItems) {
        var top    = groupItems[0].top;
        var left   = groupItems[0].left;
        var bottom = groupItems[0].bottom;
        var right  = groupItems[0].right;
        for (var i = 1; i < groupItems.length; i++) {
            top    = Math.min(top, groupItems[i].top);
            left   = Math.min(left, groupItems[i].left);
            bottom = Math.max(bottom, groupItems[i].bottom);
            right  = Math.max(right, groupItems[i].right);
        }
        return [top, left, bottom, right];
    }

    /**
     * 現在の設定でグループ範囲を示すプレビュー矩形を描き直す
     * @param {Document} targetDoc 対象ドキュメント
     * @param {Array} itemBounds 境界情報つきの配列
     * @param {string} direction "horizontal" または "vertical"
     * @param {number} tolerance 許容値
     * @returns {void}
     */
    function updatePreview(targetDoc, itemBounds, direction, tolerance) {
        clearPreview(targetDoc);

        var groups        = computeGroups(itemBounds, direction, tolerance);
        var previewLayer  = getPreviewLayer(targetDoc);
        var previewSwatch = getPreviewSwatch(targetDoc);
        var noneSwatch    = targetDoc.swatches.itemByName("[None]");

        for (var i = 0; i < groups.length; i++) {
            if (groups[i].length < 2) continue;
            var parentPage = groups[i][0].pageItem.parentPage;
            if (!parentPage) continue;

            var previewRectangle = parentPage.rectangles.add(previewLayer);
            previewRectangle.geometricBounds = getGroupBounds(groups[i]);
            previewRectangle.fillColor       = previewSwatch;
            previewRectangle.strokeColor     = noneSwatch;
            previewRectangle.opacity         = PREVIEW_OPACITY;
            previewRectangles.push(previewRectangle);
        }
        try { app.redraw(); } catch (e) {}
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 方向と許容値を指定するダイアログを表示する
     * @param {Document} targetDoc 対象ドキュメント
     * @param {Array} itemBounds 境界情報つきの配列
     * @returns {{direction: string, tolerance: number}|null} 設定内容。キャンセル時は null
     */
    function showSmartGroupDialog(targetDoc, itemBounds) {
        var smartGroupDialog = new Window("dialog", localize(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(smartGroupDialog);

        /* 方向パネル / Direction panel */
        var directionPanel = smartGroupDialog.add("panel", undefined, localize(LABELS.panel.direction));
        setupPanel(directionPanel, 6);
        directionPanel.alignChildren = ["left", "top"];

        var horizontalRadio = directionPanel.add("radiobutton", undefined, localize(LABELS.radio.horizontal));
        var verticalRadio   = directionPanel.add("radiobutton", undefined, localize(LABELS.radio.vertical));
        horizontalRadio.value = true;

        /* 許容値パネル / Tolerance panel */
        var tolerancePanel = smartGroupDialog.add("panel", undefined, localize(LABELS.panel.tolerance));
        setupPanel(tolerancePanel, 6);

        var toleranceRow = tolerancePanel.add("group");
        setupRow(toleranceRow, "left", 8);
        toleranceRow.alignChildren = ["left", "center"];

        var toleranceSlider = toleranceRow.add("slider", undefined, TOLERANCE_DEFAULT, TOLERANCE_MIN, TOLERANCE_MAX);
        toleranceSlider.preferredSize.width = TOLERANCE_SLIDER_WIDTH;
        var toleranceValueLabel = toleranceRow.add("statictext", undefined, String(TOLERANCE_DEFAULT));
        toleranceValueLabel.preferredSize.width = TOLERANCE_VALUE_WIDTH;

        /**
         * 現在の入力値でプレビューを描き直す
         * @returns {void}
         */
        function refreshPreview() {
            var direction = horizontalRadio.value ? "horizontal" : "vertical";
            var tolerance = Math.round(toleranceSlider.value);
            toleranceValueLabel.text = String(tolerance);
            updatePreview(targetDoc, itemBounds, direction, tolerance);
        }

        toleranceSlider.onChanging = refreshPreview;
        horizontalRadio.onClick    = refreshPreview;
        verticalRadio.onClick      = refreshPreview;
        smartGroupDialog.onShow    = refreshPreview;

        /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
        var dialogButtonRow = smartGroupDialog.add("group");
        setupRow(dialogButtonRow, "center", 8);
        dialogButtonRow.add("button", undefined, localize(LABELS.button.cancel), { name: "cancel" });
        dialogButtonRow.add("button", undefined, localize(LABELS.button.ok), { name: "ok" });

        var accepted = smartGroupDialog.show() === 1;

        /* ダイアログを閉じたらプレビューを片付ける / Clean up the preview once the dialog closes */
        clearPreview(targetDoc);

        if (!accepted) return null;
        return {
            direction: horizontalRadio.value ? "horizontal" : "vertical",
            tolerance: Math.round(toleranceSlider.value)
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトを方向と許容値に応じてグループ化する
     * @returns {void}
     */
    function main() {
        if (app.selection.length === 0) {
            alert(localize(LABELS.alert.noSelection));
            return;
        }

        var activeDoc  = app.activeDocument;
        var itemBounds = collectItemBounds(app.selection);

        var dialogResult = showSmartGroupDialog(activeDoc, itemBounds);
        if (dialogResult === null) return; /* キャンセル / Cancelled */

        var groups = computeGroups(itemBounds, dialogResult.direction, dialogResult.tolerance);
        var createdGroupCount = 0;

        for (var i = 0; i < groups.length; i++) {
            var groupItems = groups[i];
            if (groupItems.length < 2) continue;

            app.select(groupItems[0].pageItem);
            for (var j = 1; j < groupItems.length; j++) {
                app.select(groupItems[j].pageItem, SelectionOptions.ADD_TO);
            }
            activeDoc.groups.add(app.selection);
            createdGroupCount++;
        }

        var directionLabel = (dialogResult.direction === "horizontal")
            ? localize(LABELS.direction.horizontal)
            : localize(LABELS.direction.vertical);

        alert(currentLang === "ja"
            ? directionLabel + "で " + createdGroupCount + " 個のグループを作成しました。"
            : createdGroupCount + " group(s) created " + directionLabel + ".");
    }

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, localize(LABELS.undo.smartGroup));

})();
