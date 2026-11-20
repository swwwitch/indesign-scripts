#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdCellMergeAuto
Name=内容が同じセルを自動で結合
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.1.5
Release-Date=2026-04-17
Description-BEGIN
選択した表について、内容が同じ隣接セルを自動で結合します。結合する方向（水平・垂直・両方向）と、対象（表全体・選択セルのみ）をダイアログで選べます。

Merges adjacent cells with identical contents in the selected table. The dialog picks the merge direction (horizontal, vertical, or both) and the scope (the whole table or only the selected cells).
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択した表について、内容が同じ隣接セルを自動で結合します。結合する方向（水平・垂直・両方向）と、対象（表全体・選択セルのみ）をダイアログで選べます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCellMergeAuto.md

note記事も参照してください。
https://note.com/dtp_tranist/n/na84f68305844

### Overview

Merges adjacent cells with identical contents in the selected table. The dialog picks the merge direction (horizontal, vertical, or both) and the scope (the whole table or only the selected cells).

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCellMergeAuto.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdCellMergeAuto";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "内容が同じセルを自動で結合"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCellMergeAuto.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCellMergeAuto.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/na84f68305844"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdCellMergeAuto.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

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

// ボタン行（再利用パーツ） / Button row (reusable)

var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
    btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
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

// UI の明暗（再利用パーツ） / UI theme (reusable)

/**
 * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
 * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
 */
function isDarkUI() {
    try {
        if (app.preferences && app.preferences.getRealPreference) {
            return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
        }
        return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
    } catch (e) {
        return false;
    }
}

// UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

// バージョン表示（再利用パーツ） / Version footer (reusable)

var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

/* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
   9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
   / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
     nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
var VERSION_FOOTER_GLYPHS = {
    "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
    "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
    "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
    "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
    "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
    "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
    "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
    "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
    "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
    "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
    "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
    "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
    "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
    "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
    "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
    "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
    "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
    "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
    "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
    "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
    "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
    "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
    "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
    "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
    "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
    "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
    "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
    "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
    "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
    "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
    "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
    "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
    "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
    "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
    "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
    "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
    "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
    "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
    "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
    "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
    "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
    "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
    "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
    "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
    "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
    "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
    "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
    "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
    "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
    "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
    "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
    "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
    "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
    "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
    "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
    "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
    "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
    "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
    "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
    "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
    "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
    "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
    ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
    "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
    "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
    "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
    " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
};

/**
 * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPT_VERSION）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
 * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
 * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
 * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
 * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
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

    /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
       左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
       / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
         or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
    var versionText = String(SCRIPT_VERSION);
    var leftCandidates = [versionText];
    if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
    var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
    versionCanvas.alignment = ["fill", "center"];
    bindVersionInfoClick(versionCanvas); /* クリックで「このスクリプトについて」を開く / click opens About This Script */
    var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
    brandCanvas.alignment = ["right", "center"];
    brandCanvas.maximumSize = brandCanvas.preferredSize;
    targetWindow.versionFooter = footerGroup;
    return versionCanvas;
}

/**
 * 文字列のすべての文字に字形があるかを返す
 * @param {string} dotText - 文字列
 * @returns {boolean} すべてあれば true
 */
function hasVersionFooterGlyphs(dotText) {
    for (var i = 0; i < dotText.length; i++) {
        if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
    }
    return dotText.length > 0;
}

/**
 * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
 * @param {Group} parentGroup - 足す先
 * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
 * @returns {Group} 描くグループ
 */
function addVersionFooterCanvas(parentGroup, dotTexts) {
    var dotCanvas = parentGroup.add("group");
    var shortestText = dotTexts[dotTexts.length - 1];
    dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
    dotCanvas.helpTip = dotTexts[0];
    dotCanvas.onDraw = function () {
        var availableWidth = this.size ? this.size.width : 0;
        var drawText = shortestText;
        for (var i = 0; i < dotTexts.length; i++) {
            if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                drawText = dotTexts[i];
                break;
            }
        }
        drawVersionFooterDots(this.graphics, drawText);
    };
    return dotCanvas;
}

/**
 * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
 * @param {string} versionText - 文字列
 * @returns {number} 幅（ドット）
 */
function measureVersionFooterDots(versionText) {
    var dotCount = 0;
    for (var i = 0; i < versionText.length; i++) {
        var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
        if (!glyphRows) continue;
        if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
        dotCount += glyphRows[0].length;
    }
    return dotCount;
}

/**
 * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
 * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
 * @param {string} versionText - 文字列
 * @returns {void}
 */
function drawVersionFooterDots(canvasGraphics, versionText) {
    var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
    var dotSize = VERSION_FOOTER_PIXEL_SIZE;
    var cursorDots = 0;
    canvasGraphics.newPath();
    for (var i = 0; i < versionText.length; i++) {
        var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
        if (!glyphRows) continue;
        if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
        for (var row = 0; row < glyphRows.length; row++) {
            var rowText = glyphRows[row];
            var runStart = -1;
            for (var col = 0; col <= rowText.length; col++) {
                var isDot = col < rowText.length && rowText.charAt(col) === "#";
                if (isDot && runStart < 0) runStart = col;
                if (!isDot && runStart >= 0) {
                    canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                    runStart = -1;
                }
            }
        }
        cursorDots += glyphRows[0].length;
    }
    canvasGraphics.fillPath(footerBrush);
}

// -----------------------------------------
// このスクリプトについて / About this script
// -----------------------------------------
/* 左下のスクリプト名・バージョンをクリックすると開く。基本情報ブロックの SCRIPT_* を読み、
   この部品だけで動くよう、文言（VERSION_INFO_LABELS）とレイアウトの値も部品の中に持つ
   Opened by clicking the name and version at the bottom left. Reads the SCRIPT_* basic info,
   and keeps its own labels and layout values so the part works on its own */

var VERSION_INFO_LANG = (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en"; /* 文言の言語 / label language */
var VERSION_INFO_IS_INDESIGN = (String(app.name).indexOf("InDesign") >= 0);         /* InDesign で動いているか / running in InDesign */
/* 実行中のスクリプトのファイル。ボタンを押すころには $.fileName が別のものを指すことがあるので、読み込んだ時点で控える
   The running script file, captured at load time; $.fileName may point elsewhere by the time the button is clicked */
var VERSION_INFO_SCRIPT_FILE = new File($.fileName);
var VERSION_INFO_REVEAL_APP_PATH = "/Applications/RevealInFinder.app";   /* Finder で選択表示する補助アプリ / helper app that reveals a file in the Finder */
var VERSION_INFO_REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";  /* 補助アプリが読むパス（アプリの旧名のまま。変えると読めない）/ path the helper reads (keeps the app's old name) */

var VERSION_INFO_WINDOW_MARGINS = 16;               /* ウィンドウの余白 / window margin */
var VERSION_INFO_WINDOW_SPACING = 12;               /* ウィンドウ内の間隔 / window spacing */
var VERSION_INFO_PANEL_MARGINS = [16, 20, 16, 12];  /* パネルの余白 [左,上,右,下] / panel margins */
var VERSION_INFO_ROW_SPACING = 6;                   /* 項目名と値の間隔 / gap between a label and its value */
var VERSION_INFO_BUTTON_SPACING = 10;               /* ボタンどうしの間隔 / gap between buttons */
var VERSION_INFO_BUTTON_ROW_MARGINS = [0, 5, 0, 14]; /* ボタン行の余白（Illustrator 標準のダイアログに合わせる）/ button row margins */
var VERSION_INFO_COPY_BUTTON_SIZE = 22;             /* コピーボタンの一辺 / copy button size */
/* コピーボタンの配色。∧∨（ステップボタン）と同じく、背景に重ねる半透明 / Copy button colors, translucent overlays like the stepper */
var VERSION_INFO_COPY_FILL_COLOR    = isDarkUI() ? [0, 0, 0, 0.10] : [1, 1, 1, 0.50]; /* 地 / background */
var VERSION_INFO_COPY_FRAME_COLOR   = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.15]; /* 枠線 / frame */
var VERSION_INFO_COPY_PRESSED_COLOR = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 押下中 / pressed */
var VERSION_INFO_COPY_ICON_COLOR    = isDarkUI() ? [1, 1, 1, 0.85] : [0, 0, 0, 0.70]; /* アイコン / icon */

var VERSION_INFO_LABELS = {
    dialog: {
        title: { ja: "このスクリプトについて", en: "About This Script" }
    },
    panel: {
        generalInfo: { ja: "基本情報", en: "General" },
        links: { ja: "リンク", en: "Links" }
    },
    fieldLabel: {
        scriptName: { ja: "スクリプト名", en: "Script Name" },
        scriptFeature: { ja: "機能", en: "Feature" },
        scriptVersion: { ja: "バージョン", en: "Version" }
    },
    button: {
        openReadme: { ja: "README", en: "README" },
        openSourceCode: { ja: "ソースコード", en: "Source Code" },
        openArticle: { ja: "記事（note）", en: "Article (note)" },
        revealInFinder: { ja: "Finder に表示", en: "Show in Finder" },
        revealInExplorer: { ja: "エクスプローラーに表示", en: "Show in Explorer" },
        close: { ja: "閉じる", en: "Close" }
    },
    tooltip: {
        versionInfo: { ja: "クリックで、このスクリプトについての情報を表示", en: "Click for info about this script" },
        copyScriptName: { ja: "スクリプト名をコピー", en: "Copy Script Name" }
    },
    alert: {
        copyFailed: { ja: "スクリプト名をコピーできませんでした。", en: "Could not copy the script name." }
    }
};

/**
 * VERSION_INFO_LABELS の文言を今の言語で返す
 * @param {Object} labelEntry - { ja, en }
 * @returns {string} 文言
 */
function getVersionInfoLabel(labelEntry) {
    return labelEntry[VERSION_INFO_LANG] || labelEntry.en;
}

/**
 * 基本情報ブロックの値を返す（変数が無い・文字列でなければ空文字）。typeof で読むので、未定義でも例外にならない
 * @param {string} variableKey - "name" / "version" / "feature" / "readme" / "source" / "article"
 * @returns {string} 値
 */
function getVersionInfoValue(variableKey) {
    var infoValue;
    if (variableKey === "name") infoValue = (typeof SCRIPT_NAME !== "undefined") ? SCRIPT_NAME : "";
    else if (variableKey === "version") infoValue = (typeof SCRIPT_VERSION !== "undefined") ? SCRIPT_VERSION : "";
    else if (variableKey === "feature") infoValue = (typeof SCRIPT_DESCRIPTION_JA !== "undefined") ? SCRIPT_DESCRIPTION_JA : "";
    else if (variableKey === "readme") {
        if (VERSION_INFO_LANG === "ja") infoValue = (typeof SCRIPT_README_JA !== "undefined") ? SCRIPT_README_JA : "";
        else infoValue = (typeof SCRIPT_README_EN !== "undefined") ? SCRIPT_README_EN : "";
    }
    else if (variableKey === "source") infoValue = (typeof SCRIPT_SOURCE_URL !== "undefined") ? SCRIPT_SOURCE_URL : "";
    else if (variableKey === "article") infoValue = (typeof SCRIPT_ARTICLE_URL !== "undefined") ? SCRIPT_ARTICLE_URL : "";
    return (infoValue === undefined || infoValue === null) ? "" : String(infoValue);
}

/**
 * 左下のスクリプト名・バージョンをクリックしたら、「このスクリプトについて」のダイアログを開くようにする
 * @param {Group} versionCanvas - addVersionFooterCanvas() で作った左の欄
 * @returns {void}
 */
function bindVersionInfoClick(versionCanvas) {
    versionCanvas.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.versionInfo);
    /* mousedown で開くと、離したときのイベントが閉じたあとの欄に届くので mouseup で開く
       Open on mouseup so the release does not land on the footer after the dialog closes */
    versionCanvas.addEventListener("mouseup", function () {
        showVersionInfoDialog(versionCanvas.window.type === "palette");
    });
}

/**
 * 「このスクリプトについて」のダイアログを開く。スクリプト名（コピーボタン付き）・機能・バージョンの［基本情報］パネル、
 * README・ソースコード・記事を開く［リンク］パネル、［Finder に表示］［閉じる］を並べる
 * @param {boolean} isFromPalette - 常駐パレットから開いたなら true（Illustrator のコピーを BridgeTalk 経由にする）
 * @returns {void}
 */
function showVersionInfoDialog(isFromPalette) {
    var infoDialog = new Window("dialog", getVersionInfoLabel(VERSION_INFO_LABELS.dialog.title));
    infoDialog.orientation = "column";
    infoDialog.alignChildren = "fill";
    infoDialog.margins = VERSION_INFO_WINDOW_MARGINS;
    infoDialog.spacing = VERSION_INFO_WINDOW_SPACING;

    var infoPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.generalInfo);
    /* コピーボタンはスクリプト名だけ / Only the script name gets a copy button */
    addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptName, getVersionInfoValue("name"), isFromPalette);
    /* 機能は README の見出しで日本語だけなので、日本語の画面だけに出す / The feature is the Japanese README title, shown only in the Japanese UI */
    if (VERSION_INFO_LANG === "ja" && getVersionInfoValue("feature") !== "") {
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptFeature, getVersionInfoValue("feature"), null);
    }
    addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptVersion, getVersionInfoValue("version"), null);
    alignVersionInfoLabels(infoPanel);

    /* URL が無いボタンはディム表示にして、並びはスクリプトごとに変えない / Dim buttons without a URL; keep the same buttons everywhere */
    var linksPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.links);
    var linkButtonRow = linksPanel.add("group");
    linkButtonRow.orientation = "row";
    linkButtonRow.alignment = ["left", "center"];
    linkButtonRow.alignChildren = ["left", "center"];
    linkButtonRow.spacing = VERSION_INFO_BUTTON_SPACING;
    addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openReadme, getVersionInfoValue("readme"));
    addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openSourceCode, getVersionInfoValue("source"));
    addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openArticle, getVersionInfoValue("article"));

    /* ボタン行：左に［Finder に表示］、右に［閉じる］ / Button row: Show in Finder at the left, Close at the right */
    var buttonRow = infoDialog.add("group");
    buttonRow.orientation = "row";
    buttonRow.alignment = ["fill", "bottom"];
    buttonRow.alignChildren = ["left", "center"];
    buttonRow.margins = VERSION_INFO_BUTTON_ROW_MARGINS;
    buttonRow.spacing = VERSION_INFO_BUTTON_SPACING;
    var revealLabel = (File.fs === "Windows") ? VERSION_INFO_LABELS.button.revealInExplorer : VERSION_INFO_LABELS.button.revealInFinder;
    var revealButton = buttonRow.add("button", undefined, getVersionInfoLabel(revealLabel));
    if (VERSION_INFO_SCRIPT_FILE.exists) {
        revealButton.helpTip = VERSION_INFO_SCRIPT_FILE.fsName;
        revealButton.onClick = function () {
            revealVersionInfoScriptFile(VERSION_INFO_SCRIPT_FILE);
            infoDialog.close();
        };
    } else {
        revealButton.enabled = false;
    }
    /* 余りの幅を吸って、［閉じる］を右端に寄せる / Absorbs the extra width so Close sits at the right edge */
    var buttonSpacer = buttonRow.add("group");
    buttonSpacer.alignment = ["fill", "fill"];
    buttonSpacer.minimumSize.width = 0;
    var closeButton = buttonRow.add("button", undefined, getVersionInfoLabel(VERSION_INFO_LABELS.button.close), { name: "ok" });
    closeButton.alignment = ["right", "center"];
    /* return でも esc でも閉じる / Close with both Return and Esc */
    infoDialog.defaultElement = closeButton;
    infoDialog.cancelElement = closeButton;
    infoDialog.onShow = function () {
        pinVersionInfoCopyButton(infoPanel);
    };
    infoDialog.show();
}

/**
 * 「このスクリプトについて」のダイアログにパネルを足す
 * @param {Window} infoDialog - 足す先のダイアログ
 * @param {Object} titleEntry - パネル名の文言 { ja, en }
 * @returns {Panel} パネル
 */
function addVersionInfoPanel(infoDialog, titleEntry) {
    var infoPanel = infoDialog.add("panel", undefined, getVersionInfoLabel(titleEntry));
    infoPanel.orientation = "column";
    infoPanel.alignChildren = ["fill", "top"];
    infoPanel.alignment = "fill";
    infoPanel.margins = VERSION_INFO_PANEL_MARGINS;
    infoPanel.spacing = VERSION_INFO_WINDOW_SPACING;
    return infoPanel;
}

/**
 * ［基本情報］パネルに「項目名 : 値」の1行を足す（値は見せるだけなので statictext）
 * @param {Panel} infoPanel - 足す先のパネル
 * @param {Object} labelEntry - 項目名の文言 { ja, en }
 * @param {string} valueText - 表示する値
 * @param {boolean|null} copyFromPalette - 値の右にコピーボタンを置くなら true / false（常駐パレットから開いたか）、置かないなら null
 * @returns {StaticText} 値の statictext
 */
function addVersionInfoRow(infoPanel, labelEntry, valueText, copyFromPalette) {
    var infoRow = infoPanel.add("group");
    infoRow.orientation = "row";
    infoRow.alignment = ["fill", "center"];
    infoRow.alignChildren = ["left", "center"];
    infoRow.spacing = VERSION_INFO_ROW_SPACING;
    /* 項目名のコロンは日本語が「 :」、英語が「:」 / Colon: " :" in Japanese, ":" in English */
    var rowLabel = infoRow.add("statictext", undefined, getVersionInfoLabel(labelEntry) + (VERSION_INFO_LANG === "ja" ? " :" : ":"));
    rowLabel.justify = "right";
    var valueLabel = infoRow.add("statictext", undefined, valueText);
    if (copyFromPalette === null) return valueLabel;
    /* statictext の fill では伸びないことがあるので、伸びるスペーサーでコピーボタンをパネルの右端に寄せる
       A stretching spacer (statictext fill is unreliable) pushes the copy button to the panel's right edge */
    var copySpacer = infoRow.add("group");
    copySpacer.alignment = ["fill", "fill"];
    copySpacer.minimumSize.width = 0;
    infoRow.copyButton = addVersionInfoCopyButton(infoRow, valueLabel, copyFromPalette);
    return valueLabel;
}

/**
 * コピーボタンを［基本情報］パネルの右端に置き直す（表示時、レイアウトが決まったあとに呼ぶ）。
 * 行の fill・スペーサーだけでは、項目の値が短い行がパネルの幅まで伸びないことがあるので、行の幅とボタンの位置を直接入れる
 * @param {Panel} infoPanel - ［基本情報］パネル
 * @returns {void}
 */
function pinVersionInfoCopyButton(infoPanel) {
    var innerWidth = infoPanel.size.width - VERSION_INFO_PANEL_MARGINS[0] - VERSION_INFO_PANEL_MARGINS[2];
    for (var i = 0; i < infoPanel.children.length; i++) {
        var infoRow = infoPanel.children[i];
        if (!infoRow.copyButton || !infoRow.size || infoRow.size.width >= innerWidth) continue;
        infoRow.size = [innerWidth, infoRow.size.height];
        infoRow.copyButton.location = [innerWidth - infoRow.copyButton.size.width, infoRow.copyButton.location[1]];
    }
}

/**
 * ［基本情報］パネルの項目名を、いちばん長いものの幅にそろえる
 * @param {Panel} infoPanel - 対象のパネル
 * @returns {void}
 */
function alignVersionInfoLabels(infoPanel) {
    var rowLabels = [];
    var labelWidth = 0;
    for (var i = 0; i < infoPanel.children.length; i++) {
        var rowLabel = infoPanel.children[i].children[0]; /* 行の先頭は項目名 / each row starts with its label */
        rowLabels.push(rowLabel);
        labelWidth = Math.max(labelWidth, rowLabel.graphics.measureString(rowLabel.text)[0]);
    }
    for (var j = 0; j < rowLabels.length; j++) {
        rowLabels[j].preferredSize.width = Math.ceil(labelWidth) + 2;
    }
}

/**
 * URL を開くボタンを足す。URL が無ければディム表示にする。開いたらダイアログは閉じる
 * @param {Window} infoDialog - ボタンを置くダイアログ
 * @param {Group} buttonGroup - 足す先のグループ
 * @param {Object} labelEntry - ボタンの文言 { ja, en }
 * @param {string} targetUrl - 開く URL（空ならディム表示）
 * @returns {void}
 */
function addVersionInfoLinkButton(infoDialog, buttonGroup, labelEntry, targetUrl) {
    var linkButton = buttonGroup.add("button", undefined, getVersionInfoLabel(labelEntry));
    if (!targetUrl) {
        linkButton.enabled = false;
        return;
    }
    linkButton.helpTip = targetUrl;
    linkButton.onClick = function () {
        openVersionInfoUrl(targetUrl);
        infoDialog.close();
    };
}

/**
 * 値をクリップボードへコピーするアイコンのボタンを足す。コピーしたらアイコンをチェックマークに変え、ポインタが外れたら戻す
 * @param {Group} infoRow - 足す先の行
 * @param {StaticText} valueLabel - コピーする値の statictext
 * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
 * @returns {Group} ボタンとして使う group
 */
function addVersionInfoCopyButton(infoRow, valueLabel, isFromPalette) {
    var buttonSize = VERSION_INFO_COPY_BUTTON_SIZE;
    var copyButton = infoRow.add("group");
    copyButton.preferredSize = [buttonSize, buttonSize];
    copyButton.minimumSize = [buttonSize, buttonSize];
    copyButton.maximumSize = [buttonSize, buttonSize];
    copyButton.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.copyScriptName);
    copyButton.isPressed = false;
    copyButton.isCopied = false;

    copyButton.onDraw = function () {
        var buttonGraphics = copyButton.graphics;
        var fillColor = copyButton.isPressed ? VERSION_INFO_COPY_PRESSED_COLOR : VERSION_INFO_COPY_FILL_COLOR;
        buttonGraphics.newPath();
        buttonGraphics.rectPath(0, 0, buttonSize, buttonSize);
        buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, fillColor));
        buttonGraphics.newPath();
        buttonGraphics.rectPath(0.5, 0.5, buttonSize - 1, buttonSize - 1);
        buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_FRAME_COLOR, 1));
        if (copyButton.isCopied) drawVersionInfoCheckIcon(buttonGraphics, buttonSize);
        else drawVersionInfoCopyIcon(buttonGraphics, buttonSize);
    };

    /**
     * 状態を変えて描き直す（group には notify() が無いので隠して再表示する）
     * @param {boolean} isPressed - 押下中なら true
     * @param {boolean} isCopied - コピー済みなら true
     * @returns {void}
     */
    function repaint(isPressed, isCopied) {
        if (copyButton.isPressed === isPressed && copyButton.isCopied === isCopied) return;
        copyButton.isPressed = isPressed;
        copyButton.isCopied = isCopied;
        copyButton.hide();
        copyButton.show();
    }
    copyButton.addEventListener("mousedown", function () { repaint(true, false); });
    copyButton.addEventListener("mouseup", function () {
        if (!copyButton.isPressed) return;
        var hasCopied = copyVersionInfoText(valueLabel.text, isFromPalette);
        repaint(false, hasCopied);
        if (!hasCopied) alert(getVersionInfoLabel(VERSION_INFO_LABELS.alert.copyFailed));
    });
    /* 押したまま外へ出たとき・コピーのあと離れたときは元の見た目に戻す / Reset when the pointer leaves */
    copyButton.addEventListener("mouseout", function () { repaint(false, false); });
    return copyButton;
}

/**
 * コピーのアイコン（後ろの四角に前の四角を重ねた形）を描く
 * @param {ScriptUIGraphics} buttonGraphics - 描画先
 * @param {number} buttonSize - ボタンの一辺
 * @returns {void}
 */
function drawVersionInfoCopyIcon(buttonGraphics, buttonSize) {
    var iconPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1);
    var sheetSize = 9;
    var offset = 3;
    var frontLeft = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
    var frontTop = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
    var backLeft = frontLeft - offset;
    var backTop = frontTop - offset;
    /* 後ろの四角は、前の四角に隠れる部分を描かない（上と左の辺だけ）/ Draw only the visible top and left of the back sheet */
    buttonGraphics.newPath();
    buttonGraphics.moveTo(frontLeft, backTop + sheetSize);
    buttonGraphics.lineTo(backLeft, backTop + sheetSize);
    buttonGraphics.lineTo(backLeft, backTop);
    buttonGraphics.lineTo(backLeft + sheetSize, backTop);
    buttonGraphics.lineTo(backLeft + sheetSize, frontTop);
    buttonGraphics.strokePath(iconPen);
    buttonGraphics.newPath();
    buttonGraphics.rectPath(frontLeft, frontTop, sheetSize, sheetSize);
    buttonGraphics.strokePath(iconPen);
}

/**
 * コピーしたあとのチェックマークを描く
 * @param {ScriptUIGraphics} buttonGraphics - 描画先
 * @param {number} buttonSize - ボタンの一辺
 * @returns {void}
 */
function drawVersionInfoCheckIcon(buttonGraphics, buttonSize) {
    var center = buttonSize / 2;
    buttonGraphics.newPath();
    buttonGraphics.moveTo(center - 4.5, center);
    buttonGraphics.lineTo(center - 1.5, center + 3);
    buttonGraphics.lineTo(center + 4.5, center - 3.5);
    buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1.5));
}

/**
 * 文字列をクリップボードにコピーする。
 * InDesign は AppleScript（Windows は VBScript）の doScript で書き込む。
 * Illustrator は文字列を直接送る API が無いので、一時レイヤーのテキストフレームを選んでコピーする（プレーンテキストも入る）。
 * Illustrator の常駐パレットからは DOM に触れないので、同じ処理を BridgeTalk でメインエンジンに渡す
 * @param {string} copyText - コピーする文字列
 * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
 * @returns {boolean} コピーの操作まで進めたら true
 */
function copyVersionInfoText(copyText, isFromPalette) {
    try {
        if (VERSION_INFO_IS_INDESIGN) {
            if (File.fs === "Windows") {
                app.doScript('CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""') + '"', ScriptLanguage.VISUAL_BASIC);
            } else {
                app.doScript('set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"', ScriptLanguage.APPLESCRIPT_LANGUAGE);
            }
            return true;
        }
        var copyScript = buildVersionInfoCopyScript(copyText);
        if (!isFromPalette) return eval(copyScript) === true;
        var copyMessage = new BridgeTalk();
        copyMessage.target = BridgeTalk.appSpecifier; /* 複数バージョンを起動していても、今のバージョンに送る / the running version, even with several installed */
        copyMessage.body = copyScript;
        copyMessage.send();
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * Illustrator でコピーする処理を文字列で組み立てる（BridgeTalk でも送れるよう、関数の toString() は使わない）。
 * 一時レイヤーにテキストフレームを作り、フレームを選んでメニューコマンドでコピーして、レイヤーごと消す。
 * 文字として選択（textRange.select()）してコピーすると「オブジェクトをコピーできません」の警告が出るので使わない。
 * 選択と作業レイヤーは元に戻す。ドキュメントが無ければ一時ドキュメントを作って閉じる
 * @param {string} copyText - コピーする文字列
 * @returns {string} 実行すると true / false を返すスクリプト
 */
function buildVersionInfoCopyScript(copyText) {
    return [
        "(function (copyText) {",
        "    var usingTempDoc = (app.documents.length === 0);",
        "    var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;",
        "    var savedSelection = (!usingTempDoc && targetDoc.selection instanceof Array) ? targetDoc.selection : [];",
        "    var savedActiveLayer = usingTempDoc ? null : targetDoc.activeLayer;",
        "    var tempLayer = null;",
        "    var hasCopied = false;",
        "    try {",
        "        tempLayer = targetDoc.layers.add();",
        "        var tempFrame = tempLayer.textFrames.add();",
        "        tempFrame.contents = copyText;",
        "        app.redraw();",
        "        app.executeMenuCommand('deselectall');",
        "        tempFrame.selected = true;",
        "        app.redraw();",
        "        app.executeMenuCommand('copy');",
        "        app.redraw();",
        "        hasCopied = true;",
        "    } catch (e) { }",
        "    if (tempLayer) { try { tempLayer.remove(); } catch (e) { } }",
        "    if (usingTempDoc) { targetDoc.close(SaveOptions.DONOTSAVECHANGES); return hasCopied; }",
        "    try { targetDoc.activeLayer = savedActiveLayer; } catch (e) { }",
        "    targetDoc.selection = null;",
        "    for (var i = 0; i < savedSelection.length; i++) { try { savedSelection[i].selected = true; } catch (e) { } }",
        "    app.redraw();",
        "    return hasCopied;",
        "})(" + quoteVersionInfoString(copyText) + ");"
    ].join(" "); /* 行コメントを含まないので1行につなぐ（BridgeTalk の転送で改行が化けない）/ joined on one line; there are no line comments */
}

/**
 * 文字列を、スクリプトに埋め込める二重引用符のリテラルにする
 * @param {string} sourceText - 文字列
 * @returns {string} "…" の形のリテラル
 */
function quoteVersionInfoString(sourceText) {
    return '"' + String(sourceText).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r/g, "\\r").replace(/\n/g, "\\n") + '"';
}

/**
 * ファイルを Finder で表示する。RevealInFinder.app があればファイルを選択した状態で表示し、無ければ入っているフォルダーを開く
 * @param {File} scriptFile - 表示するファイル
 * @returns {void}
 */
function revealVersionInfoScriptFile(scriptFile) {
    /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
    var hasRevealApp = (File.fs === "Macintosh")
        && (new Folder(VERSION_INFO_REVEAL_APP_PATH).exists || new File(VERSION_INFO_REVEAL_APP_PATH).exists);
    /* アプリは一時ファイルからパスを読むので、書き出してから起動する / The app reads the path from a temp file */
    if (hasRevealApp && writeVersionInfoRevealPath(scriptFile.fsName) && new File(VERSION_INFO_REVEAL_APP_PATH).execute()) return;
    /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
    scriptFile.parent.execute();
}

/**
 * RevealInFinder.app に渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
 * @param {string} targetPath - 絶対パス
 * @returns {boolean} 書き出せたら true
 */
function writeVersionInfoRevealPath(targetPath) {
    var pathFile = new File(VERSION_INFO_REVEAL_PATH_FILE);
    pathFile.encoding = "UTF-8";
    pathFile.lineFeed = "Unix";
    if (!pathFile.open("w")) return false;
    var hasWritten = pathFile.write(targetPath);
    pathFile.close();
    return hasWritten;
}

/**
 * URL を既定のブラウザーで開く（一時的なショートカットファイルを作って開く）
 * @param {string} targetUrl - 開く URL
 * @returns {void}
 */
function openVersionInfoUrl(targetUrl) {
    var isMac = (File.fs === "Macintosh");
    var shortcutFile = new File(Folder.temp.fsName + "/VersionInfoLink" + (isMac ? ".webloc" : ".url"));
    shortcutFile.encoding = "UTF-8";
    shortcutFile.open("w");
    if (isMac) {
        var escapedUrl = targetUrl.replace(/&/g, "&amp;").replace(/</g, "&lt;");
        shortcutFile.write('<?xml version="1.0" encoding="UTF-8"?>\n'
            + '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
            + '<plist version="1.0"><dict><key>URL</key><string>' + escapedUrl + '</string></dict></plist>\n');
    } else {
        shortcutFile.write("[InternetShortcut]\r\nURL=" + targetUrl + "\r\n");
    }
    shortcutFile.close();
    shortcutFile.execute();
}

// バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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
            title: { ja: "自動でセル結合", en: "Auto Merge Cells" }
        },
        panel: {
            mergeMode: { ja: "結合", en: "Merge" },
            direction: { ja: "方向", en: "Direction" },
            scope:     { ja: "対象", en: "Scope" }
        },
        radio: {
            singleDirection: { ja: "単一方向", en: "Single direction" },
            bothDirections:  { ja: "両方向（優先付き）", en: "Both directions (with priority)" },
            horizontal:      { ja: "水平", en: "Horizontal" },
            vertical:        { ja: "垂直", en: "Vertical" },
            wholeTable:      { ja: "表全体", en: "Whole table" },
            selectedCells:   { ja: "選択セルのみ", en: "Selected cells only" }
        },
        tooltip: {
            mergeMode: {
                ja: "「単一方向」は［方向］で選んだ側だけ、「両方向」は水平・垂直の両方を結合します。",
                en: "Single direction merges only the side chosen under Direction; Both directions merges horizontally and vertically."
            },
            direction: {
                ja: "「両方向」のときは、ここで選んだ側から先に結合します。",
                en: "In Both directions mode, the side chosen here is merged first."
            },
            scope: {
                ja: "「選択セルのみ」は、選択したセルを囲む長方形の範囲が対象になります。",
                en: "Selected cells only targets the rectangle that encloses the selected cells."
            }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            selectTable:   { ja: "表、または表の中のセルを選択してください。", en: "Select a table or cells inside a table." },
            tableNotFound: { ja: "表が見つかりません。", en: "No table was found." },
            noCellRange: {
                ja: "「選択セルのみ」が選択されていますが、有効なセル選択が見つかりませんでした。\n表全体を対象にするか、複数セルを選択してください。",
                en: "\"Selected cells only\" is chosen, but no valid cell selection was found.\nTarget the whole table or select multiple cells."
            }
        },
        undo: {
            autoMerge: { ja: "自動でセル結合", en: "Auto Merge Cells" }
        }
    };

    // =========================================
    // テキスト処理 / Text helpers
    // =========================================

    /**
     * 前後の空白を取り除く（ES3 環境向けの自前実装）
     * @param {*} rawValue 対象の値
     * @returns {string} 前後の空白を除いた文字列
     */
    function customTrim(rawValue) {
        var textValue = (typeof rawValue === "string") ? rawValue : String(rawValue);
        return textValue.replace(/^\s+|\s+$/g, "");
    }

    /**
     * 比較用にセルのテキストを取り出す
     * @param {Cell} targetCell 対象のセル
     * @returns {string} 前後の空白を除いたセルの内容
     */
    function getCellText(targetCell) {
        if (!targetCell || !targetCell.isValid) return "";

        var cellContents = targetCell.contents;

        /* contents が配列で返る場合に備えて連結する / contents can come back as an array */
        if (cellContents instanceof Array) cellContents = cellContents.join("");

        return customTrim(cellContents);
    }

    // =========================================
    // 表と選択範囲の取得 / Table and range lookup
    // =========================================

    // 表の選択（再利用パーツ） / Table selection (reusable)

    var TABLE_PARENT_LOOKUP_LIMIT = 20; /* 親をたどる上限（無限ループよけ） / max parent hops (guards against loops) */

    /**
     * 親をたどって、いちばん近い表を返す（セル・行・列・セル内のテキストや挿入点に対応）
     * @param {Object} startItem - たどり始めるオブジェクト
     * @returns {Table|null} 見つかった表。表の中でなければ null
     */
    function findParentTable(startItem) {
        var node = startItem;
        for (var i = 0; i < TABLE_PARENT_LOOKUP_LIMIT && node; i++) {
            try {
                var typeName = node.constructor.name;
                if (typeName === "Table") return node;
                if (typeName === "Document" || typeName === "Application") return null;
                node = node.parent;
            } catch (e) {
                return null;
            }
        }
        return null;
    }

    /**
     * 選択から対象の表を返す。表の中の選択ならその表、表を含むテキストフレームやテキストなら最初の表
     * @param {Object} selectionItem - app.selection[0] など
     * @returns {Table|null} 対象の表。見つからなければ null
     */
    function getTableFromSelection(selectionItem) {
        if (!selectionItem) return null;
        var parentTable = findParentTable(selectionItem);
        if (parentTable) return parentTable;
        try {
            if (selectionItem.tables && selectionItem.tables.length > 0) return selectionItem.tables[0];
        } catch (e) {
            /* tables を持たない選択（画像など） / Selections without tables (images, etc.) */
        }
        return null;
    }

    /**
     * 選択からセルを1つずつの配列にして返す。セル選択・表の選択はその全セル、セル内のテキストや挿入点はそのセル
     * @param {Object} selectionItem - app.selection[0] など
     * @returns {Cell[]} セルの配列。表の外なら空配列
     */
    function getSelectedCells(selectionItem) {
        var selectedCells = [];
        if (!selectionItem) return selectedCells;
        try {
            var typeName = selectionItem.constructor.name;
            if (typeName === "Cell" || typeName === "Table") {
                /* 複数セルの選択も1つの Cell で返るので .cells で展開する / A multi-cell selection is one Cell; expand it via .cells */
                var cellCollection = selectionItem.cells;
                for (var i = 0; i < cellCollection.length; i++) selectedCells.push(cellCollection[i]);
                return selectedCells;
            }
            var node = selectionItem;
            for (var j = 0; j < TABLE_PARENT_LOOKUP_LIMIT && node; j++) {
                var nodeType = node.constructor.name;
                if (nodeType === "Cell") {
                    selectedCells.push(node);
                    break;
                }
                if (nodeType === "Table" || nodeType === "Document" || nodeType === "Application") break;
                node = node.parent;
            }
        } catch (e) {
            /* 親をたどれない選択は空のまま / Leave empty when the parent chain cannot be followed */
        }
        return selectedCells;
    }

    /**
     * 2つの表が同じ表かどうかを返す
     * @param {Table} tableA - 表
     * @param {Table} tableB - 表
     * @returns {boolean} 同じ表なら true
     */
    function isSameTable(tableA, tableB) {
        if (!tableA || !tableB) return false;
        try {
            return tableA.id === tableB.id && tableA.parent.id === tableB.parent.id;
        } catch (e) {
            return false;
        }
    }

    // 表の選択（再利用パーツ）ここまで / End of the reusable table selection

    /**
     * 選択セルを囲む矩形範囲を求める
     * @returns {{rowStart: number, rowEnd: number, colStart: number, colEnd: number}|null} 矩形範囲。選択がなければ null
     */
    function getSelectedCellRange() {
        var minRow = Number.MAX_VALUE;
        var maxRow = -1;
        var minCol = Number.MAX_VALUE;
        var maxCol = -1;

        for (var i = 0; i < app.selection.length; i++) {
            var selectedCells = getSelectedCells(app.selection[i]);

            for (var j = 0; j < selectedCells.length; j++) {
                var targetCell = selectedCells[j];
                if (!targetCell || !targetCell.isValid) continue;

                minRow = Math.min(minRow, targetCell.parentRow.index);
                maxRow = Math.max(maxRow, targetCell.parentRow.index);
                minCol = Math.min(minCol, targetCell.parentColumn.index);
                maxCol = Math.max(maxCol, targetCell.parentColumn.index);
            }
        }

        if (maxRow < 0) return null;

        return { rowStart: minRow, rowEnd: maxRow, colStart: minCol, colEnd: maxCol };
    }

    // =========================================
    // セル結合 / Cell merging
    // =========================================

    /**
     * 走査方向と直交する側のインデックスを取り出す
     * @param {Cell} targetCell 対象のセル
     * @param {boolean} isHorizontal true なら水平方向の走査
     * @returns {number} 水平方向なら列インデックス、垂直方向なら行インデックス
     */
    function getCrossIndex(targetCell, isHorizontal) {
        return isHorizontal ? targetCell.parentColumn.index : targetCell.parentRow.index;
    }

    /**
     * 内容が同じ隣接セルを指定方向に結合する
     * @param {Table} targetTable 対象の表
     * @param {object|null} cellRange 対象の矩形範囲。null なら表全体
     * @param {boolean} isHorizontal true なら水平方向、false なら垂直方向
     * @returns {void}
     */
    function mergeSameContentCells(targetTable, cellRange, isHorizontal) {
        /* 水平方向なら行を、垂直方向なら列を 1 本ずつ走査する
           / Scan row by row when horizontal, column by column when vertical */
        var scanLines = isHorizontal ? targetTable.rows : targetTable.columns;

        /* 走査する行・列の範囲と、その中で対象になるセルの範囲
           / Bounds for the lines to scan, and for the cells inside a line */
        var targetRange = cellRange ||
            { rowStart: 0, rowEnd: Number.MAX_VALUE, colStart: 0, colEnd: Number.MAX_VALUE };
        var lineStart = isHorizontal ? targetRange.rowStart : targetRange.colStart;
        var lineEnd   = isHorizontal ? targetRange.rowEnd   : targetRange.colEnd;
        var cellStart = isHorizontal ? targetRange.colStart : targetRange.rowStart;
        var cellEnd   = isHorizontal ? targetRange.colEnd   : targetRange.rowEnd;

        for (var lineIndex = lineStart; lineIndex < scanLines.length && lineIndex <= lineEnd; lineIndex++) {
            var currentLine = scanLines[lineIndex];
            if (!currentLine.isValid) continue;

            var lineCells = currentLine.cells;
            var cellIndex = 0;

            while (cellIndex < lineCells.length - 1) {
                var headCell = lineCells[cellIndex];
                var nextCell = lineCells[cellIndex + 1];

                if (!headCell || !headCell.isValid || !nextCell || !nextCell.isValid) {
                    cellIndex++;
                    continue;
                }

                /* 表全体のときは範囲を見に行かない / Skip the bounds lookup when the whole table is the target */
                if (cellRange &&
                    (getCrossIndex(headCell, isHorizontal) < cellStart ||
                     getCrossIndex(nextCell, isHorizontal) > cellEnd)) {
                    cellIndex++;
                    continue;
                }

                var headText = getCellText(headCell);
                if (headText !== getCellText(nextCell)) {
                    cellIndex++;
                    continue;
                }

                /* またぎ方が揃わないセル同士は merge() が失敗するので、その組は飛ばす
                   / merge() fails when the spans do not line up, so skip that pair */
                try {
                    headCell.merge(nextCell);
                } catch (e) {
                    cellIndex++;
                    continue;
                }

                headCell.contents = headText;
                lineCells = currentLine.cells; /* 結合でセル配列が変わるため取り直す / Refresh after the merge */
            }
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ラジオボタンを縦に並べたパネルを追加する
     * @param {Window|Group} parentContainer 追加先
     * @param {object} panelLabel パネル見出しのラベル
     * @param {Array.<object>} radioLabels ラジオボタンのラベル配列
     * @param {number} selectedIndex 最初に選択するラジオボタンのインデックス
     * @param {object} [tooltipLabel] パネルとラジオボタンに付けるツールチップのラベル
     * @returns {Array.<RadioButton>} 追加したラジオボタン
     */
    function addRadioPanel(parentContainer, panelLabel, radioLabels, selectedIndex, tooltipLabel) {
        var radioPanel = parentContainer.add("panel", undefined, getLabel(panelLabel));
        setupPanel(radioPanel, 6);
        radioPanel.alignChildren = ["left", "top"];

        /* パネルだけに付けると枠の上でしか出ないので、ラジオボタンにも同じ説明を持たせる
           / A panel-only helpTip shows on the frame alone, so give the radio buttons the same text */
        var tooltipText = tooltipLabel ? getLabel(tooltipLabel) : "";
        radioPanel.helpTip = tooltipText;

        var radioButtons = [];
        for (var i = 0; i < radioLabels.length; i++) {
            radioButtons[i] = radioPanel.add("radiobutton", undefined, getLabel(radioLabels[i]));
            radioButtons[i].helpTip = tooltipText;
        }
        radioButtons[selectedIndex].value = true;

        return radioButtons;
    }

    /**
     * 結合方法を指定するダイアログを表示する
     * @returns {{passes: Array.<boolean>, useSelectionOnly: boolean}|null} 設定内容。キャンセル時は null
     */
    function showAutoMergeDialog() {
        var autoMergeDialog = new Window("dialog", getLabel(LABELS.dialog.title));
        setupWindow(autoMergeDialog, 10);

        var settingsColumn = autoMergeDialog.add("group");
        settingsColumn.orientation = "column";
        settingsColumn.alignChildren = ["fill", "top"];
        settingsColumn.spacing = PANEL_SPACING;

        var mergeModeRadios = addRadioPanel(settingsColumn, LABELS.panel.mergeMode,
            [LABELS.radio.singleDirection, LABELS.radio.bothDirections], 1, LABELS.tooltip.mergeMode);
        var directionRadios = addRadioPanel(settingsColumn, LABELS.panel.direction,
            [LABELS.radio.horizontal, LABELS.radio.vertical], 0, LABELS.tooltip.direction);
        var scopeRadios = addRadioPanel(settingsColumn, LABELS.panel.scope,
            [LABELS.radio.wholeTable, LABELS.radio.selectedCells], 0, LABELS.tooltip.scope);

        var buttonRow = addButtonRow(autoMergeDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        addVersionFooter(autoMergeDialog);
        if (autoMergeDialog.show() !== 1) return null;

        var useBothDirections = mergeModeRadios[1].value;
        var startsHorizontal = directionRadios[0].value;

        /* 結合を試す順番。true は水平、false は垂直 / Passes to run; true is horizontal, false is vertical */
        return {
            passes: useBothDirections ? [startsHorizontal, !startsHorizontal] : [startsHorizontal],
            useSelectionOnly: scopeRadios[1].value
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログの設定に従って同一内容のセルを自動結合する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        if (app.selection.length === 0) {
            alert(getLabel(LABELS.alert.selectTable));
            return;
        }

        var targetTable = getTableFromSelection(app.selection[0]);
        if (!targetTable) {
            alert(getLabel(LABELS.alert.tableNotFound));
            return;
        }

        var dialogResult = showAutoMergeDialog();
        if (!dialogResult) return;

        /* 対象範囲。null なら表全体 / The target range; null means the whole table */
        var cellRange = null;
        if (dialogResult.useSelectionOnly) {
            cellRange = getSelectedCellRange();
            if (!cellRange) {
                alert(getLabel(LABELS.alert.noCellRange));
                return;
            }
        }

        /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
        app.doScript(function () {
            for (var i = 0; i < dialogResult.passes.length; i++) {
                mergeSameContentCells(targetTable, cellRange, dialogResult.passes[i]);
            }
        }, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undo.autoMerge));
    }

    main();

})();
