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
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-25";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSmartGroup.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSmartGroup.md"; /* README (English) */

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

/* プレビュー枠の線色（CMYK）・線幅・不透明度 / Stroke color (CMYK), weight and opacity of the preview frame */
var PREVIEW_SWATCH_CMYK   = [0, 100, 100, 0];
var PREVIEW_STROKE_WEIGHT = "10pt";
var PREVIEW_OPACITY       = 50;

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
    function getCurrentUILang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentUILang();

    var LABELS = {
        dialog: {
            title: { ja: "スマートグループ", en: "Smart Group" }
        },
        panel: {
            direction:         { ja: "グループ化する方向", en: "Grouping direction" },
            tolerance:         { ja: "許容値", en: "Tolerance" },
            toleranceWithUnit: { ja: "許容値（%1）", en: "Tolerance (%1)" }
        },
        checkbox: {
            showPreview: { ja: "プレビューを表示", en: "Show preview" }
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
        tooltip: {
            horizontal: {
                ja: "上下の中心がそろったオブジェクトを、横一列ずつグループ化します。",
                en: "Groups objects whose vertical centers line up, one row at a time."
            },
            vertical: {
                ja: "左右の中心がそろったオブジェクトを、縦一列ずつグループ化します。",
                en: "Groups objects whose horizontal centers line up, one column at a time."
            },
            tolerance: {
                ja: "隣り合うオブジェクトの中心のずれがこの値以内なら、同じ行／列とみなします（定規の単位）。",
                en: "Objects whose centers are within this distance of their neighbor count as the same row/column (ruler units)."
            }
        },
        alert: {
            noSelection: { ja: "アイテムを選択してください。", en: "Please select one or more items." },
            result:      { ja: "%1で %2 個のグループを作成しました。", en: "%2 group(s) created %1." },
            noClusters:  { ja: "グループ化できる並びがありませんでした。", en: "No rows or columns to group were found." }
        },
        undo: {
            smartGroup: { ja: "スマートグループ化", en: "Smart Group" }
        }
    };

    /**
     * ラベル定義から現在の UI 言語の文字列を取り出す
     * @param {{ja: string, en: string}} labelSet 言語別のラベル定義
     * @returns {string} 現在の UI 言語の文字列
     */
    function getLabel(labelSet) {
        return labelSet[uiLang] || labelSet.en;
    }

    /**
     * ラベル内のプレースホルダー（%1, %2 …）を値で置き換える
     * @param {string} template プレースホルダーを含む文字列
     * @param {Array} values 差し込む値
     * @returns {string} 置き換え後の文字列
     */
    function formatLabel(template, values) {
        var text = template;
        for (var i = 0; i < values.length; i++) {
            text = text.split("%" + (i + 1)).join(String(values[i]));
        }
        return text;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /**
     * 定規の単位から表示用の単位名を返す
     * @param {MeasurementUnits} rulerUnit 対象の単位
     * @returns {string} 単位名（表示しない単位は空文字）
     */
    function getUnitName(rulerUnit) {
        switch (rulerUnit) {
            case MeasurementUnits.MILLIMETERS:    return "mm";
            case MeasurementUnits.CENTIMETERS:    return "cm";
            case MeasurementUnits.POINTS:         return "pt";
            case MeasurementUnits.INCHES:         return "in";
            case MeasurementUnits.INCHES_DECIMAL: return "in";
            case MeasurementUnits.PICAS:          return "p";
            case MeasurementUnits.CICEROS:        return "c";
            case MeasurementUnits.PIXELS:         return "px";
            case MeasurementUnits.Q:              return "Q";
            case MeasurementUnits.HA:             return "H";
            default:                              return "";
        }
    }

    /**
     * 比べる座標の向きに合った定規の単位名を返す
     * 水平方向（行）は Y 座標を比べるので縦の定規、垂直方向（列）は横の定規を使う
     * @param {Document} targetDoc 対象ドキュメント
     * @param {string} direction "horizontal" または "vertical"
     * @returns {string} 単位名
     */
    function getToleranceUnitName(targetDoc, direction) {
        var viewPrefs = targetDoc.viewPreferences;
        return getUnitName((direction === "horizontal")
            ? viewPrefs.verticalMeasurementUnits
            : viewPrefs.horizontalMeasurementUnits);
    }

    /**
     * 許容値パネルの見出しを単位付きで返す
     * @param {Document} targetDoc 対象ドキュメント
     * @param {string} direction "horizontal" または "vertical"
     * @returns {string} 見出し
     */
    function getTolerancePanelTitle(targetDoc, direction) {
        var unitName = getToleranceUnitName(targetDoc, direction);
        return unitName
            ? formatLabel(getLabel(LABELS.panel.toleranceWithUnit), [unitName])
            : getLabel(LABELS.panel.tolerance);
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビュー用の赤スウォッチを取得する（なければ作成）
     * @param {Document} targetDoc 対象ドキュメント
     * @returns {Color} プレビュー用スウォッチ
     */
    function getPreviewSwatch(targetDoc) {
        var previewSwatch = targetDoc.colors.itemByName(PREVIEW_SWATCH_NAME);
        if (previewSwatch.isValid) return previewSwatch;
        return targetDoc.colors.add({
            name: PREVIEW_SWATCH_NAME,
            model: ColorModel.PROCESS,
            space: ColorSpace.CMYK,
            colorValue: PREVIEW_SWATCH_CMYK
        });
    }

    /**
     * プレビュー用の非印刷レイヤーを最上位に取得する（なければ作成）
     * @param {Document} targetDoc 対象ドキュメント
     * @returns {Layer} プレビュー用レイヤー
     */
    function getPreviewLayer(targetDoc) {
        var previewLayer = targetDoc.layers.itemByName(PREVIEW_LAYER_NAME);
        if (previewLayer.isValid) return previewLayer;
        previewLayer = targetDoc.layers.add({ name: PREVIEW_LAYER_NAME, printable: false });
        previewLayer.move(LocationOptions.AT_BEGINNING);
        return previewLayer;
    }

    /**
     * プレビューをレイヤーごと削除する
     * @param {Document} targetDoc 対象ドキュメント
     * @returns {void}
     */
    function clearPreview(targetDoc) {
        var previewLayer = targetDoc.layers.itemByName(PREVIEW_LAYER_NAME);
        if (previewLayer.isValid) previewLayer.remove();
        app.redraw();
    }

    /**
     * プレビューのレイヤーとスウォッチをすべて片付ける
     * @param {Document} targetDoc 対象ドキュメント
     * @returns {void}
     */
    function removePreviewResources(targetDoc) {
        clearPreview(targetDoc);
        var previewSwatch = targetDoc.colors.itemByName(PREVIEW_SWATCH_NAME);
        if (previewSwatch.isValid) previewSwatch.remove();
    }

    /**
     * 現在の設定でまとまりの範囲を示すプレビュー枠を描き直す
     * @param {Document} targetDoc 対象ドキュメント
     * @param {Array} itemBounds 境界情報つきの配列
     * @param {string} direction "horizontal" または "vertical"
     * @param {number} tolerance 許容値
     * @returns {void}
     */
    function updatePreview(targetDoc, itemBounds, direction, tolerance) {
        clearPreview(targetDoc);

        var clusters      = computeClusters(itemBounds, direction, tolerance);
        var previewLayer  = getPreviewLayer(targetDoc);
        var previewSwatch = getPreviewSwatch(targetDoc);
        var noneSwatch    = targetDoc.swatches.itemByName("[None]");

        for (var i = 0; i < clusters.length; i++) {
            if (clusters[i].length < 2) continue;
            var parentPage = clusters[i][0].pageItem.parentPage;
            if (!parentPage) continue;

            var previewRectangle = parentPage.rectangles.add(previewLayer);
            previewRectangle.geometricBounds = getClusterBounds(clusters[i]);
            previewRectangle.fillColor       = noneSwatch;
            previewRectangle.strokeColor     = previewSwatch;
            previewRectangle.strokeWeight    = PREVIEW_STROKE_WEIGHT;
            previewRectangle.opacity         = PREVIEW_OPACITY;
        }
        app.redraw();
    }

    // =========================================
    // まとまりの計算 / Cluster computation
    // =========================================

    /**
     * 選択オブジェクトの境界と中心座標を取り出す
     * @param {Array} selectedItems 選択オブジェクトの配列
     * @returns {Array<{pageItem: PageItem, top: number, left: number, bottom: number, right: number, centerX: number, centerY: number}>} 境界情報つきの配列
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
                right: bounds[3],
                centerX: (bounds[1] + bounds[3]) / 2,
                centerY: (bounds[0] + bounds[2]) / 2
            });
        }
        return itemBounds;
    }

    /**
     * 指定方向で比べる中心座標を返す
     * @param {{centerX: number, centerY: number}} itemBound 境界情報
     * @param {string} direction "horizontal"（Y 中心）または "vertical"（X 中心）
     * @returns {number} 中心座標
     */
    function getCenterAlongAxis(itemBound, direction) {
        return (direction === "horizontal") ? itemBound.centerY : itemBound.centerX;
    }

    /**
     * 中心座標の近さでオブジェクトを行または列のまとまりに分ける
     * @param {Array} itemBounds 境界情報つきの配列
     * @param {string} direction "horizontal" または "vertical"
     * @param {number} tolerance 同じ行／列とみなす許容値
     * @returns {Array<Array>} まとまりごとの配列
     */
    function computeClusters(itemBounds, direction, tolerance) {
        var sortedItems = itemBounds.slice();
        sortedItems.sort(function (a, b) {
            return getCenterAlongAxis(a, direction) - getCenterAlongAxis(b, direction);
        });

        var clusters = [];
        var currentCluster = [sortedItems[0]];

        for (var i = 1; i < sortedItems.length; i++) {
            var previousItem = currentCluster[currentCluster.length - 1];
            var centerDelta = getCenterAlongAxis(sortedItems[i], direction) - getCenterAlongAxis(previousItem, direction);
            if (centerDelta <= tolerance) {
                currentCluster.push(sortedItems[i]);
            } else {
                clusters.push(currentCluster);
                currentCluster = [sortedItems[i]];
            }
        }
        clusters.push(currentCluster);
        return clusters;
    }

    /**
     * まとまり全体を囲む境界を求める
     * @param {Array} clusterItems 同じまとまりに属する境界情報の配列
     * @returns {Array<number>} [上, 左, 下, 右]
     */
    function getClusterBounds(clusterItems) {
        var top    = clusterItems[0].top;
        var left   = clusterItems[0].left;
        var bottom = clusterItems[0].bottom;
        var right  = clusterItems[0].right;
        for (var i = 1; i < clusterItems.length; i++) {
            top    = Math.min(top, clusterItems[i].top);
            left   = Math.min(left, clusterItems[i].left);
            bottom = Math.max(bottom, clusterItems[i].bottom);
            right  = Math.max(right, clusterItems[i].right);
        }
        return [top, left, bottom, right];
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 方向パネルを作る
     * @param {Window} parentDialog 親ダイアログ
     * @returns {{horizontal: RadioButton, vertical: RadioButton}} 方向のラジオボタン
     */
    function buildDirectionPanel(parentDialog) {
        var directionPanel = parentDialog.add("panel", undefined, getLabel(LABELS.panel.direction));
        setupPanel(directionPanel, 6);
        directionPanel.alignChildren = ["left", "top"];

        var horizontalRadio = directionPanel.add("radiobutton", undefined, getLabel(LABELS.radio.horizontal));
        var verticalRadio   = directionPanel.add("radiobutton", undefined, getLabel(LABELS.radio.vertical));
        horizontalRadio.helpTip = getLabel(LABELS.tooltip.horizontal);
        verticalRadio.helpTip   = getLabel(LABELS.tooltip.vertical);
        horizontalRadio.value = true;

        return { horizontal: horizontalRadio, vertical: verticalRadio };
    }

    /**
     * 許容値パネルを作る
     * @param {Window} parentDialog 親ダイアログ
     * @param {string} panelTitle パネルの見出し
     * @returns {{panel: Panel, slider: Slider, valueLabel: StaticText}} パネル・スライダー・数値表示
     */
    function buildTolerancePanel(parentDialog, panelTitle) {
        var tolerancePanel = parentDialog.add("panel", undefined, panelTitle);
        setupPanel(tolerancePanel, 6);

        var toleranceRowGroup = tolerancePanel.add("group");
        setupRow(toleranceRowGroup, "left", 8);
        toleranceRowGroup.alignChildren = ["left", "center"];

        var toleranceSlider = toleranceRowGroup.add("slider", undefined, TOLERANCE_DEFAULT, TOLERANCE_MIN, TOLERANCE_MAX);
        toleranceSlider.preferredSize.width = TOLERANCE_SLIDER_WIDTH;
        toleranceSlider.helpTip = getLabel(LABELS.tooltip.tolerance);

        var toleranceValueLabel = toleranceRowGroup.add("statictext", undefined, String(TOLERANCE_DEFAULT));
        toleranceValueLabel.preferredSize.width = TOLERANCE_VALUE_WIDTH;

        return { panel: tolerancePanel, slider: toleranceSlider, valueLabel: toleranceValueLabel };
    }

    /**
     * ［プレビューを表示］チェックボックスを作る（初期値 ON）
     * @param {Window} parentDialog 親ダイアログ
     * @returns {Checkbox} チェックボックス
     */
    function buildPreviewCheckbox(parentDialog) {
        var previewCheckbox = parentDialog.add("checkbox", undefined, getLabel(LABELS.checkbox.showPreview));
        previewCheckbox.value = true;
        return previewCheckbox;
    }

    /**
     * キャンセル／OK ボタンの行を作る（幅いっぱいには広げない）
     * @param {Window} parentDialog 親ダイアログ
     * @returns {void}
     */
    function buildButtonRow(parentDialog) {
        var btnRowGroup = parentDialog.add("group");
        setupRow(btnRowGroup, "center", 8);
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
    }

    /**
     * 方向と許容値を指定するダイアログを表示する
     * @param {Document} targetDoc 対象ドキュメント
     * @param {Array} itemBounds 境界情報つきの配列
     * @returns {{direction: string, tolerance: number}|null} 設定内容。キャンセル時は null
     */
    function showSmartGroupDialog(targetDoc, itemBounds) {
        var smartGroupDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(smartGroupDialog);

        var directionRadios   = buildDirectionPanel(smartGroupDialog);
        var toleranceControls = buildTolerancePanel(smartGroupDialog, getTolerancePanelTitle(targetDoc, "horizontal"));
        var previewCheckbox   = buildPreviewCheckbox(smartGroupDialog);
        buildButtonRow(smartGroupDialog);

        /* プレビュー用レイヤーを足すとアクティブレイヤーが変わるので控えておく / Adding the preview layer changes the active layer, so remember it */
        var originalActiveLayer = targetDoc.activeLayer;

        /**
         * ダイアログの現在の設定を読み取る
         * @returns {{direction: string, tolerance: number}} 設定内容
         */
        function readSettings() {
            return {
                direction: directionRadios.horizontal.value ? "horizontal" : "vertical",
                tolerance: Math.round(toleranceControls.slider.value)
            };
        }

        /**
         * 現在の入力値でプレビューを描き直す
         * @returns {void}
         */
        function refreshPreview() {
            var settings = readSettings();
            toleranceControls.valueLabel.text = String(settings.tolerance);
            toleranceControls.panel.text = getTolerancePanelTitle(targetDoc, settings.direction);
            if (previewCheckbox.value) {
                updatePreview(targetDoc, itemBounds, settings.direction, settings.tolerance);
            } else {
                clearPreview(targetDoc);
            }
        }

        toleranceControls.slider.onChanging = refreshPreview;
        directionRadios.horizontal.onClick  = refreshPreview;
        directionRadios.vertical.onClick    = refreshPreview;
        previewCheckbox.onClick             = refreshPreview;
        smartGroupDialog.onShow             = refreshPreview;

        var accepted = smartGroupDialog.show() === 1;

        /* ダイアログを閉じたらプレビューを片付ける / Clean up the preview once the dialog closes */
        removePreviewResources(targetDoc);
        targetDoc.activeLayer = originalActiveLayer;

        return accepted ? readSettings() : null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 2 つ以上のオブジェクトを含むまとまりをグループ化する
     * @param {Document} targetDoc 対象ドキュメント
     * @param {Array<Array>} clusters まとまりごとの配列
     * @returns {number} 作成したグループの数
     */
    function groupClusters(targetDoc, clusters) {
        var createdGroupCount = 0;
        for (var i = 0; i < clusters.length; i++) {
            /* グループ化には 2 つ以上のオブジェクトが必要 / Grouping requires at least two items */
            if (clusters[i].length < 2) continue;

            var itemsToGroup = [];
            for (var j = 0; j < clusters[i].length; j++) {
                itemsToGroup.push(clusters[i][j].pageItem);
            }
            targetDoc.groups.add(itemsToGroup);
            createdGroupCount++;
        }
        return createdGroupCount;
    }

    /**
     * 選択オブジェクトを方向と許容値に応じてグループ化する
     * @returns {void}
     */
    function main() {
        if (app.selection.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var activeDoc  = app.activeDocument;
        var itemBounds = collectItemBounds(app.selection);

        var dialogResult = showSmartGroupDialog(activeDoc, itemBounds);
        if (dialogResult === null) return; /* キャンセル / Cancelled */

        var clusters = computeClusters(itemBounds, dialogResult.direction, dialogResult.tolerance);
        var createdGroupCount = groupClusters(activeDoc, clusters);

        if (createdGroupCount === 0) {
            alert(getLabel(LABELS.alert.noClusters));
            return;
        }
        alert(formatLabel(getLabel(LABELS.alert.result), [
            getLabel(LABELS.direction[dialogResult.direction]),
            createdGroupCount
        ]));
    }

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undo.smartGroup));

})();
