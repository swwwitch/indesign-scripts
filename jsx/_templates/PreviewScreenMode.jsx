#target indesign

/*

### 概要

ドキュメントウィンドウの画面モード（標準モード／プレビュー）を切り替えるボタンの再利用テンプレートです。
プレビュー付きダイアログのボタン行（左側）に置き、ボタンの文字は「押すと切り替わる先」を表示します。

### Overview

A reusable template for a button that switches the document window between the Normal and Preview screen modes.
Place it on the left of a preview dialog's button row; the button shows the mode it switches to.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PreviewScreenMode";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // 【移植手順 / How to port】
    // 1. 「（再利用パーツ）」の行から「ここまで」の行までをまるごと、コピー先の IIFE 内に貼る（ローカライズの部品より後ならどこでもよい）。
    //    識別子は isPreviewScreenMode / togglePreviewScreenMode / getScreenModeButtonLabel / addScreenModeButton
    // 2. コピー先の LABELS に次の3つを足す（このファイルの LABELS から写す）
    //      button.screenModePreview … 標準モードのときの表示（押すとプレビューへ）
    //      button.screenModeNormal  … プレビューのときの表示（押すと標準モードへ）
    //      tooltip.screenMode
    // 3. ボタン行の左のグループに addScreenModeButton(buttonRow.leftGroup) で置く。
    //    切り替えのあとに何かしたいときは第2引数に関数を渡す（例: プレビューの描き直し）
    // 4. 同じ役割の既存の関数（isInPreviewScreenMode、toggleScreenMode、getPreviewToggleButtonLabel など）は消して、これに寄せる

    // プレビュー画面モード（再利用パーツ） / Preview screen mode (reusable)

    /**
     * 作業中のウィンドウがプレビュー画面モードかどうかを返す
     * @returns {boolean} プレビューなら true。ストーリーエディターなど screenMode の無いウィンドウでは false
     */
    function isPreviewScreenMode() {
        try {
            var docWindow = app.activeWindow;
            return !!(docWindow && docWindow.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE);
        } catch (e) {
            /* ストーリーエディターのウィンドウには screenMode が無い / Story editor windows have no screenMode */
            return false;
        }
    }

    /**
     * 作業中のウィンドウの画面モードを、標準モードとプレビューで切り替える
     * @returns {void}
     */
    function togglePreviewScreenMode() {
        try {
            var docWindow = app.activeWindow;
            if (!docWindow) return;
            docWindow.screenMode = isPreviewScreenMode() ? ScreenModeOptions.PREVIEW_OFF : ScreenModeOptions.PREVIEW_TO_PAGE;
        } catch (e) {
            /* screenMode の無いウィンドウでは何もしない / Nothing to do for windows without a screenMode */
        }
    }

    /**
     * 画面モードの切り替えボタンの文字（押すと切り替わる先）を返す
     * @returns {string} プレビュー中は「標準モード」、それ以外は「プレビュー」
     */
    function getScreenModeButtonLabel() {
        return getLabel(isPreviewScreenMode() ? LABELS.button.screenModeNormal : LABELS.button.screenModePreview);
    }

    /**
     * 画面モードの切り替えボタンを足す
     * @param {Group} parent - ボタンを足す先（ふつうはボタン行の左のグループ）
     * @param {Function} [onToggle] - 切り替えたあとに呼ぶ関数
     * @returns {Button} 足したボタン
     */
    function addScreenModeButton(parent, onToggle) {
        var btnScreenMode = parent.add("button", undefined, getScreenModeButtonLabel());
        btnScreenMode.helpTip = getLabel(LABELS.tooltip.screenMode);
        btnScreenMode.onClick = function () {
            togglePreviewScreenMode();
            btnScreenMode.text = getScreenModeButtonLabel();
            if (onToggle) onToggle();
        };
        return btnScreenMode;
    }

    // プレビュー画面モード（再利用パーツ）ここまで / End of the reusable preview screen mode

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "プレビュー画面モード", en: "Preview Screen Mode" }
        },
        button: {
            screenModePreview: { ja: "プレビュー", en: "Preview" },
            screenModeNormal: { ja: "標準モード", en: "Normal Mode" },
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            screenMode: {
                ja: "ドキュメントウィンドウの画面モードを、標準モードとプレビューで切り替えます。",
                en: "Switches the document window between the Normal and Preview screen modes."
            }
        }
    };

    /**
     * 現在の UI 言語のラベルを返す
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} ラベル
     */
    function getLabel(labelEntry) {
        return labelEntry[uiLang] || labelEntry.en;
    }

    // =========================================
    // デモ / Demo
    // =========================================
    /**
     * デモのダイアログを表示する
     * @returns {void}
     */
    function showDemoDialog() {
        var demoDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        demoDialog.orientation = "row";
        demoDialog.margins = 16;
        demoDialog.spacing = 10;
        addScreenModeButton(demoDialog);
        demoDialog.add("button", undefined, getLabel(LABELS.button.close), { name: "ok" });
        demoDialog.show();
    }

    if (app.documents.length === 0) {
        alert("Open a document first.");
        return;
    }
    showDemoDialog();

})();
