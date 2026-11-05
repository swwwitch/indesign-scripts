#target indesign

/*

### 概要

ダイアログ・パレットの左下にバージョン（SCRIPT_VERSION）を薄く表示する再利用テンプレートです。
タイトルにはバージョンを入れず、この部品で下端に表示します。

### Overview

A reusable template that shows the version (SCRIPT_VERSION) dimmed at the bottom left of a dialog or palette.
The version is no longer put in the title; this part shows it at the bottom instead.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "VersionFooter";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-05";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // 【移植手順 / How to port】
    // 1. 「（再利用パーツ）」の行から「ここまで」の行までをまるごと、コピー先の IIFE 内（ボタン行の部品のあと）に貼る。
    //    識別子は VERSION_FOOTER_* / addVersionFooter
    // 2. ウィンドウのタイトルにはバージョンを入れない
    //      var mainDialog = new Window("dialog", getLabel("dialog.title"));
    // 3. 部品をすべて足したあと、show() の直前で addVersionFooter() を呼ぶ（ウィンドウは縦並び。何度呼んでも1回だけ足す）
    //      addVersionFooter(mainDialog);
    //      mainDialog.show();
    // 4. 直前の行（ボタン行など）の下余白は 0 になり、ウィンドウの下余白は VERSION_FOOTER_BOTTOM_MARGIN になる。
    //    バージョンの上の間隔は、ウィンドウの間隔に VERSION_FOOTER_TOP_MARGIN を足したもの
    // 5. 確認・進捗などの小さいウィンドウには付けない（メインのダイアログ・パレットだけ）
    // 6. InDesign の ScriptUI は文字サイズの指定を無視する（statictext も onDraw の drawString も）ので、
    //    小さくはできない。標準の大きさのまま無効（ディム表示）にして薄くする

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */

    /**
     * ウィンドウの下端にバージョン（SCRIPT_VERSION）を薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * 文字サイズは変えられない（InDesign の ScriptUI は無視する）ので、無効にしてディム表示の薄さにする
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {StaticText|null} バージョンの文字。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];
        var versionText = footerGroup.add("statictext", undefined, SCRIPT_VERSION);
        versionText.enabled = false;
        targetWindow.versionFooter = versionText;
        return versionText;
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

    // =========================================
    // デモ / Demo
    // =========================================
    /**
     * デモのダイアログを表示する
     * @returns {void}
     */
    function showDemoDialog() {
        var demoDialog = new Window("dialog", "Version Footer");
        demoDialog.orientation = "column";
        demoDialog.alignChildren = ["fill", "top"];
        demoDialog.margins = 16;
        demoDialog.spacing = 12;

        var contentPanel = demoDialog.add("panel", undefined, "Content");
        contentPanel.preferredSize = [320, 80];

        var btnRowGroup = demoDialog.add("group");
        btnRowGroup.alignment = ["right", "bottom"];
        btnRowGroup.margins = [0, 5, 0, 14];
        btnRowGroup.add("button", undefined, "Cancel", { name: "cancel" });
        btnRowGroup.add("button", undefined, "OK", { name: "ok" });

        addVersionFooter(demoDialog);
        demoDialog.show();
    }

    showDemoDialog();

})();
