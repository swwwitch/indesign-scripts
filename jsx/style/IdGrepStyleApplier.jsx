#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdGrepStyleApplier
Name=段落スタイルの正規表現スタイルを適用・管理
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.0.6
Release-Date=2026-05-03
Description-BEGIN
段落スタイルに正規表現スタイル（GREP スタイル）を適用・管理します。ルールと文字スタイルを選び、複数の段落スタイルへまとめて反映できます。

Applies and manages GREP styles on paragraph styles. Pick a rule and a character style and push it to several paragraph styles at once.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

段落スタイルに正規表現スタイル（GREP スタイル）を適用・管理します。ルールと文字スタイルを選び、複数の段落スタイルへまとめて反映できます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdGrepStyleApplier.md

### Overview

Applies and manages GREP styles on paragraph styles. Pick a rule and a character style and push it to several paragraph styles at once.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdGrepStyleApplier.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdGrepStyleApplier";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "段落スタイルの正規表現スタイルを適用・管理"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdGrepStyleApplier.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdGrepStyleApplier.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdGrepStyleApplier.jsx"; /* 最新版のソース / latest source */

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
    /* 余りの幅を取り、コピーボタンを右端に置く / Take the extra width so the copy button sits at the right */
    valueLabel.alignment = ["fill", "center"];
    if (copyFromPalette !== null) addVersionInfoCopyButton(infoRow, valueLabel, copyFromPalette);
    return valueLabel;
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

var LABELS = {
    dialog: {
        title:      { ja: "正規表現スタイルを適用", en: "Apply GREP Styles" },
        addRuleTitleDefault: { ja: "新規ルール", en: "New Rule" }
    },
    panel: {
        regex:                 { ja: "正規表現のルール", en: "GREP Rules" },
        characterStyle:        { ja: "適用する文字スタイル", en: "Character Style to Apply" },
        targetParagraphStyles: { ja: "適用先の段落スタイル", en: "Target Paragraph Styles" }
    },
    button: {
        addRule: { ja: "＋ ルール追加", en: "+ Add Rule" },
        create:  { ja: "作成", en: "Create" },
        cancel:  { ja: "キャンセル", en: "Cancel" },
        ok:      { ja: "OK", en: "OK" }
    },
    hint: {
        multipleSelect: { ja: "複数の段落スタイルを選択できます", en: "Multiple selection allowed" }
    },
    prompt: {
        addRuleTitle:      { ja: "追加する正規表現の管理用の名称を入力してください。", en: "Enter a management name for the GREP expression to add." },
        addRuleExpression: { ja: "正規表現を入力してください。", en: "Enter a GREP expression." }
    },
    tooltip: {
        regexRuleList:          { ja: "登録済みの正規表現ルールを選択します。右側で適用先の段落スタイルを指定します。", en: "Select a saved GREP rule. Configure target paragraph styles on the right." },
        selectedExpression:     { ja: "選択中の正規表現です。\\t などの制御文字は、文字として読める形で表示します。", en: "The selected GREP expression. Control characters such as \\t are shown literally." },
        addRuleButton:          { ja: "新しい正規表現ルールを追加します。", en: "Add a new GREP rule." },
        characterStyleDropdown: { ja: "適用する文字スタイルを選択します。", en: "Select the character style to apply." },
        newCharacterStyleName:  { ja: "新しく作成する文字スタイル名を入力します。", en: "Enter a name for the new character style." },
        createCharacterStyle:   { ja: "文字スタイルを作成します。ルールによっては言語設定・分割禁止・前後アキなどの初期設定を自動適用します。", en: "Create a character style. Depending on the selected rule, default settings such as language, no-break, or spacing are applied automatically." },
        paragraphStyleList:     { ja: "正規表現スタイルを適用する段落スタイルを選択します。Option/Altクリックで全選択／全解除できます。", en: "Select paragraph styles to apply the GREP style to. Option/Alt-click toggles select all." },
        addRuleTitleInput:      { ja: "このルールを識別するための管理用名称です。処理内容には影響しません。", en: "Management name used to identify this rule. It does not affect processing." },
        addRuleExpressionInput: { ja: "適用する正規表現を入力します。例：(?<=\\t)\\d+", en: "Enter the GREP expression to apply. Example: (?<=\\t)\\d+" },
        addRuleDialogOk:        { ja: "入力した名称と正規表現でルールを追加します。", en: "Add a rule using the entered name and GREP expression." },
        addRuleDialogCancel:    { ja: "ルール追加をキャンセルします。", en: "Cancel adding the rule." },
        mainOk:                 { ja: "選択した文字スタイルと段落スタイルに、選択中の正規表現スタイルを適用します。", en: "Apply the selected GREP style using the selected character and paragraph styles." },
        mainCancel:             { ja: "処理を実行せずに閉じます。", en: "Close without applying changes." }
    },
    rule: {
        bulletLabel:    { ja: "箇条書きのラベル", en: "Bullet Label" },
        language:       { ja: "言語設定", en: "Language" },
        sumaru:         { ja: "スマル", en: "No-break Ending" },
        tocNumber:      { ja: "目次の数字", en: "TOC Number" },
        inlineGraphic:  { ja: "インライングラフィック", en: "Inline Graphic" }
    },
    undo: {
        applyGrepStyles: { ja: "正規表現スタイルを適用", en: "Apply GREP Styles" }
    },
    error: {
        noDocument:                 { ja: "ドキュメントを開いてから実行してください。", en: "Open a document before running this script." },
        noParagraphStyles:          { ja: "段落スタイルが見つかりません。", en: "No paragraph styles were found." },
        emptyCharacterStyleName:    { ja: "文字スタイル名を入力してください。", en: "Enter a character style name." },
        duplicateCharacterStyleName:{ ja: "同名の文字スタイルが既にあります。", en: "A character style with the same name already exists." }
    }
};

(function () {
    if (app.documents.length === 0) {
        alert(getLabel("error.noDocument"));
        return;
    }
    var doc = app.activeDocument;

    var nestedGrepRules = [
        {
            key: "bulletLabel",
            title: getLabel("rule.bulletLabel"),
            defaultParagraph: "ul-li",
            defaultCharacter: "li-label",
            autoSelect: true,
            expression: "^.+?(?=：)"
        },
        {
            key: "language",
            title: getLabel("rule.language"),
            defaultParagraph: "p",
            defaultCharacter: "lang-US",
            autoSelect: true,
            expression: "[\\u\\l]",
            apply: function (style) {
                var langNames = ["English: USA", "英語：米国"];
                for (var languageNameIndex = 0; languageNameIndex < langNames.length; languageNameIndex++) {
                    var candidate = app.languagesWithVendors.itemByName(langNames[languageNameIndex]);
                    if (candidate.isValid) {
                        style.appliedLanguage = candidate;
                        break;
                    }
                }
            }
        },
        {
            key: "sumaru",
            title: getLabel("rule.sumaru"),
            defaultParagraph: "p",
            defaultCharacter: "sumaru",
            autoSelect: true,
            expression: "..[。」』？！…]?$",
            apply: function (style) {
                style.noBreak = true;
            }
        },
        {
            key: "tocNumber",
            title: getLabel("rule.tocNumber"),
            defaultParagraph: "p",
            defaultCharacter: "",
            autoSelect: false,
            expression: "(?<=\\t)\\d+"
        },
        {
            key: "inlineGraphic",
            title: getLabel("rule.inlineGraphic"),
            defaultParagraph: "p",
            defaultCharacter: "inline-graphic",
            autoSelect: true,
            expression: "~a",
            apply: function (style) {
                style.leadingAki = 0.25;
                style.trailingAki = 0.25;
            }
        }
    ];

    /**
     * 「[...]」で始まる既定スタイルを除いたスタイル名を集める
     * @param {object} styleCollection スタイルのコレクション
     * @returns {Array<string>} スタイル名の配列
     */
    function collectVisibleStyleNames(styleCollection) {
        var visibleStyleNames = [];
        for (var styleIndex = 0; styleIndex < styleCollection.length; styleIndex++) {
            var styleName = styleCollection[styleIndex].name;
            if (styleName.charAt(0) === "[") continue;
            visibleStyleNames.push(styleName);
        }
        return visibleStyleNames;
    }

    /**
     * 名前の一覧から一致する位置を探す
     * @param {Array<string>} nameList 名前の一覧
     * @param {string} targetName 探す名前
     * @returns {number} 見つかった位置。なければ -1
     */
    function findNameIndex(nameList, targetName) {
        for (var nameIndex = 0; nameIndex < nameList.length; nameIndex++) {
            if (nameList[nameIndex] === targetName) return nameIndex;
        }
        return -1;
    }

    /**
     * 名前でスタイルを探す（スタイルグループ内も対象）
     * @param {array} styleCollection allParagraphStyles / allCharacterStyles
     * @param {string} styleName 探すスタイル名
     * @returns {object} 見つかったスタイル。無い場合は null
     */
    function findStyleByName(styleCollection, styleName) {
        for (var i = 0; i < styleCollection.length; i++) {
            if (styleCollection[i].name === styleName) return styleCollection[i];
        }

        return null;
    }

    /**
     * 正規表現スタイルの設定ダイアログを表示する
     * @param {Array<object>} grepRules 正規表現ルールの一覧
     * @param {Array<string>} paragraphStyleNames 段落スタイル名の一覧
     * @param {Array<string>} characterStyleNames 文字スタイル名の一覧
     * @param {Document} targetDocument 対象ドキュメント
     * @returns {Array<object>|null} 適用するスタイルの組み合わせ。キャンセル時は null
     */
    function showDialog(grepRules, paragraphStyleNames, characterStyleNames, targetDocument) {
        var dialogWindow = new Window("dialog", getLabel("dialog.title"));
        setupWindow(dialogWindow);

        var mainContentGroup = dialogWindow.add("group");
        setupRow(mainContentGroup, "fill", COLUMN_SPACING);
        mainContentGroup.alignChildren = ["fill", "fill"];

        var leftColumn = mainContentGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];
        leftColumn.spacing = 8;

        var regexPanel = leftColumn.add("panel", undefined, getLabel("panel.regex"));
        setupPanel(regexPanel, 6);
        var grepRuleTitles = [];
        for (var ruleIndex = 0; ruleIndex < grepRules.length; ruleIndex++) grepRuleTitles.push(grepRules[ruleIndex].title);
        var grepRuleListbox = regexPanel.add("listbox", undefined, grepRuleTitles);
        grepRuleListbox.alignment = ["fill", "top"];
        grepRuleListbox.preferredSize.height = 140;
        grepRuleListbox.helpTip = getLabel("tooltip.regexRuleList");

        var selectedRegexGroup = regexPanel.add("group");
        selectedRegexGroup.orientation = "column";
        selectedRegexGroup.alignChildren = "left";
        selectedRegexGroup.spacing = 3;
        var selectedExpressionText = selectedRegexGroup.add("edittext", undefined, "");
        // selectedExpressionText.alignment = ["fill", "top"];
        selectedExpressionText.preferredSize.width = 180;
        selectedExpressionText.enabled = false;
        selectedExpressionText.helpTip = getLabel("tooltip.selectedExpression");

        var btnAddRule = regexPanel.add("button", undefined, getLabel("button.addRule"));
        btnAddRule.alignment = "right";
        btnAddRule.helpTip = getLabel("tooltip.addRuleButton");

        /* 共通の文字スタイルパネル / Shared character style panel */
        var sharedCharacterStylePanel = leftColumn.add("panel", undefined, getLabel("panel.characterStyle"));
        setupPanel(sharedCharacterStylePanel, 6);

        var sharedCharacterStyleDropdown = sharedCharacterStylePanel.add("dropdownlist", undefined, characterStyleNames);
        sharedCharacterStyleDropdown.preferredSize.width = 160;
        sharedCharacterStyleDropdown.helpTip = getLabel("tooltip.characterStyleDropdown");

        var characterStyleCreateGroup = sharedCharacterStylePanel.add("group");
        setupRow(characterStyleCreateGroup, "left", 4);
        var newCharacterStyleNameInput = characterStyleCreateGroup.add("edittext", undefined, "");
        newCharacterStyleNameInput.preferredSize.width = 120;
        newCharacterStyleNameInput.helpTip = getLabel("tooltip.newCharacterStyleName");
        var btnCreateCharacterStyle = characterStyleCreateGroup.add("button", undefined, getLabel("button.create"));
        btnCreateCharacterStyle.helpTip = getLabel("tooltip.createCharacterStyle");

        /* 右カラム（縦構造）/ Right column with vertical layout */
        var paragraphStyleColumn = mainContentGroup.add("group");
        paragraphStyleColumn.orientation = "column";
        paragraphStyleColumn.alignChildren = ["fill", "top"];

        var paragraphStylePanelStack = paragraphStyleColumn.add("group");
        paragraphStylePanelStack.orientation = "stack";
        paragraphStylePanelStack.alignChildren = ["fill", "top"];

        var ruleRows = [];

        /**
         * ルールに対応する既定の文字スタイル名を返す
         * @param {object} grepRule 正規表現ルール
         * @returns {string} 文字スタイル名
         */
        function getDefaultCharacterStyleNameForRule(grepRule) {
            if (!grepRule) return "";
            return grepRule.defaultCharacter || "";
        }

        /**
         * 選択したルールに合わせて表示と候補を更新する
         * @param {number} selectedRuleIndex 選択したルールの位置
         * @returns {void}
         */
        function updateSelectedGrepRule(selectedRuleIndex) {
            for (var rowIndex = 0; rowIndex < ruleRows.length; rowIndex++) {
                ruleRows[rowIndex].paragraphStylePanel.visible = (rowIndex === selectedRuleIndex);
            }
            var expressionForDisplay = (selectedRuleIndex >= 0 && ruleRows[selectedRuleIndex]) ? ruleRows[selectedRuleIndex].rule.expression : "";
            /* 表示用にエスケープ（制御文字のみ）/ Escape control characters for display (\t, \n, \r) */
            selectedExpressionText.text = expressionForDisplay
                .replace(/\t/g, "\\t")
                .replace(/\n/g, "\\n")
                .replace(/\r/g, "\\r");

            if (selectedRuleIndex >= 0 && ruleRows[selectedRuleIndex]) {
                var selectedGrepRule = ruleRows[selectedRuleIndex].rule;
                if (selectedGrepRule.autoSelect) {
                    var defaultCharacterStyleName = getDefaultCharacterStyleNameForRule(selectedGrepRule);
                    var defaultCharacterStyleIndex = findNameIndex(characterStyleNames, defaultCharacterStyleName);
                    sharedCharacterStyleDropdown.selection = (defaultCharacterStyleIndex >= 0) ? defaultCharacterStyleIndex : null;
                } else {
                    sharedCharacterStyleDropdown.selection = null;
                }
            } else {
                sharedCharacterStyleDropdown.selection = null;
            }
        }

        /**
         * ルール 1 件分の表示行を作る
         * @param {object} grepRule 正規表現ルール
         * @returns {string} リストに表示する文字列
         */
        function createGrepRuleRow(grepRule) {
            var paragraphStylePanel = paragraphStylePanelStack.add("panel", undefined, getLabel("panel.targetParagraphStyles"));
            setupPanel(paragraphStylePanel, 6);
            paragraphStylePanel.preferredSize = [280, 300];
            paragraphStylePanel.visible = false;

            paragraphStylePanel.add("statictext", undefined, getLabel("hint.multipleSelect"));
            var paragraphStyleListbox = paragraphStylePanel.add("listbox", undefined, paragraphStyleNames, { multiselect: true });
            paragraphStyleListbox.alignment = ["fill", "fill"];
            paragraphStyleListbox.preferredSize.height = 300;
            paragraphStyleListbox.helpTip = getLabel("tooltip.paragraphStyleList");

            paragraphStyleListbox.onClick = function () {
                /* Option/Altクリックで全選択トグル / Toggle select all with Option/Alt-click */
                var isAlt = ScriptUI.environment.keyboardState.altKey;
                if (!isAlt) return;

                var shouldSelectAll = false;
                for (var paragraphStyleIndex = 0; paragraphStyleIndex < paragraphStyleListbox.items.length; paragraphStyleIndex++) {
                    if (!paragraphStyleListbox.items[paragraphStyleIndex].selected) {
                        shouldSelectAll = true;
                        break;
                    }
                }

                for (var selectIndex = 0; selectIndex < paragraphStyleListbox.items.length; selectIndex++) {
                    paragraphStyleListbox.items[selectIndex].selected = shouldSelectAll;
                }
            };
            for (var defaultParagraphStyleIndex = 0; defaultParagraphStyleIndex < paragraphStyleListbox.items.length; defaultParagraphStyleIndex++) {
                if (paragraphStyleListbox.items[defaultParagraphStyleIndex].text === grepRule.defaultParagraph) {
                    paragraphStyleListbox.items[defaultParagraphStyleIndex].selected = true;
                }
            }

            /* 文字スタイルは共通パネルに統合 / Character style controls are unified in the shared panel */

            var ruleRow = {
                rule: grepRule,
                paragraphStylePanel: paragraphStylePanel,
                paragraphStyleListbox: paragraphStyleListbox
            };
            ruleRows.push(ruleRow);

            return ruleRow;
        }

        for (var grepRuleIndex = 0; grepRuleIndex < grepRules.length; grepRuleIndex++) {
            createGrepRuleRow(grepRules[grepRuleIndex]);
        }

        if (ruleRows.length > 0) {
            grepRuleListbox.selection = 0;
            updateSelectedGrepRule(0);
        }

        grepRuleListbox.onChange = function () {
            var selectedIndex = grepRuleListbox.selection ? grepRuleListbox.selection.index : -1;
            updateSelectedGrepRule(selectedIndex);
        };

        btnAddRule.onClick = function () {

            var ruleDialog = new Window("dialog", getLabel("button.addRule"));
            setupWindow(ruleDialog, 10);

            // --- タイトル入力 ---
            var titleGroup = ruleDialog.add("group");
            titleGroup.orientation = "column";
            titleGroup.alignChildren = "left";

            titleGroup.add("statictext", undefined, getLabel("prompt.addRuleTitle"));
            var titleInput = titleGroup.add("edittext", undefined, getLabel("dialog.addRuleTitleDefault"));
            titleInput.preferredSize.width = 240;
            titleInput.helpTip = getLabel("tooltip.addRuleTitleInput");

            // --- 正規表現入力 ---
            var exprGroup = ruleDialog.add("group");
            exprGroup.orientation = "column";
            exprGroup.alignChildren = "left";

            exprGroup.add("statictext", undefined, getLabel("prompt.addRuleExpression"));
            var expressionInput = exprGroup.add("edittext", undefined, "");
            expressionInput.preferredSize.width = 240;
            expressionInput.helpTip = getLabel("tooltip.addRuleExpressionInput");

            // --- ボタン ---
            /* ボタン行（キャンセル → OK）/ Button row (Cancel, then OK) */
            var ruleButtonRow = addButtonRow(ruleDialog);
            var btnRuleCancel = ruleButtonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
            var btnRuleOK = ruleButtonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
            btnRuleCancel.helpTip = getLabel("tooltip.addRuleDialogCancel");
            btnRuleOK.helpTip = getLabel("tooltip.addRuleDialogOk");
            alignRightOnlyButtonRow(ruleButtonRow);

            // --- OKボタン制御 ---
            /**
             * 入力状態に応じてコントロールの有効／無効を切り替える
             * @returns {void}
             */
            function updateState() {
                btnRuleOK.enabled = (titleInput.text.length > 0 && expressionInput.text.length > 0);
            }

            titleInput.onChanging = updateState;
            expressionInput.onChanging = updateState;

            updateState();

            // --- 実行 ---
            if (ruleDialog.show() !== 1) return;

            var newGrepRule = {
                key: "custom_" + (new Date().getTime()),
                title: titleInput.text,
                defaultParagraph: paragraphStyleNames.length > 0 ? paragraphStyleNames[0] : "",
                defaultCharacter: "",
                autoSelect: false,
                expression: expressionInput.text
            };

            grepRules.push(newGrepRule);
            grepRuleListbox.add("item", newGrepRule.title);

            createGrepRuleRow(newGrepRule);
            grepRuleListbox.selection = grepRuleListbox.items.length - 1;
            updateSelectedGrepRule(ruleRows.length - 1);

            dialogWindow.layout.layout(true);
        };

        btnCreateCharacterStyle.onClick = function () {
            var newCharacterStyleName = newCharacterStyleNameInput.text;
            if (!newCharacterStyleName || newCharacterStyleName.length === 0) {
                alert(getLabel("error.emptyCharacterStyleName"));
                return;
            }
            if (findNameIndex(characterStyleNames, newCharacterStyleName) >= 0) {
                alert(getLabel("error.duplicateCharacterStyleName"));
                return;
            }
            var newCharacterStyle = doc.characterStyles.add({ name: newCharacterStyleName });
            var selectedGrepRuleIndex = grepRuleListbox.selection ? grepRuleListbox.selection.index : -1;
            var selectedGrepRule = (selectedGrepRuleIndex >= 0 && ruleRows[selectedGrepRuleIndex]) ? ruleRows[selectedGrepRuleIndex].rule : null;
            if (selectedGrepRule && typeof selectedGrepRule.apply === "function") {
                selectedGrepRule.apply(newCharacterStyle);
            }
            characterStyleNames.push(newCharacterStyleName);
            sharedCharacterStyleDropdown.add("item", newCharacterStyleName);
            sharedCharacterStyleDropdown.selection = sharedCharacterStyleDropdown.items.length - 1;
            newCharacterStyleNameInput.text = "";
            dialogWindow.layout.layout(true);
        };

        /* ボタン行（キャンセル → OK）/ Button row (Cancel, then OK) */
        var buttonRow = addButtonRow(dialogWindow);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        btnCancel.helpTip = getLabel("tooltip.mainCancel");
        btnOK.helpTip = getLabel("tooltip.mainOk");
        alignRightOnlyButtonRow(buttonRow);

        /**
         * 必須項目の入力状況に応じて OK ボタンの有効／無効を切り替える
         * @returns {void}
         */
        function updateOkButtonState() {
            var hasCharacter = sharedCharacterStyleDropdown.selection !== null;
            var hasParagraph = false;

            for (var okRuleRowIndex = 0; okRuleRowIndex < ruleRows.length; okRuleRowIndex++) {
                var listbox = ruleRows[okRuleRowIndex].paragraphStyleListbox;
                for (var okParagraphStyleIndex = 0; okParagraphStyleIndex < listbox.items.length; okParagraphStyleIndex++) {
                    if (listbox.items[okParagraphStyleIndex].selected) {
                        hasParagraph = true;
                        break;
                    }
                }
                if (hasParagraph) break;
            }

            btnOK.enabled = hasCharacter && hasParagraph;
        }

        sharedCharacterStyleDropdown.onChange = updateOkButtonState;

        for (var okBindingRowIndex = 0; okBindingRowIndex < ruleRows.length; okBindingRowIndex++) {
            (function (listbox) {
                listbox.onChange = updateOkButtonState;
            })(ruleRows[okBindingRowIndex].paragraphStyleListbox);
        }

        updateOkButtonState();

        addVersionFooter(dialogWindow);
        if (dialogWindow.show() !== 1) return null;

        var styleRegistrations = [];
        for (var ruleRowIndex = 0; ruleRowIndex < ruleRows.length; ruleRowIndex++) {
            var ruleRow = ruleRows[ruleRowIndex];
            var selectedParagraphStyleNames = [];
            for (var paragraphStyleItemIndex = 0; paragraphStyleItemIndex < ruleRow.paragraphStyleListbox.items.length; paragraphStyleItemIndex++) {
                if (ruleRow.paragraphStyleListbox.items[paragraphStyleItemIndex].selected) {
                    selectedParagraphStyleNames.push(ruleRow.paragraphStyleListbox.items[paragraphStyleItemIndex].text);
                }
            }
            var selectedCharacterStyleName = (sharedCharacterStyleDropdown.selection) ? sharedCharacterStyleDropdown.selection.text : null;
            if (selectedParagraphStyleNames.length === 0 || !selectedCharacterStyleName) continue;
            for (var selectedParagraphStyleIndex = 0; selectedParagraphStyleIndex < selectedParagraphStyleNames.length; selectedParagraphStyleIndex++) {
                styleRegistrations.push({
                    paragraph: selectedParagraphStyleNames[selectedParagraphStyleIndex],
                    character: selectedCharacterStyleName,
                    expression: ruleRow.rule.expression
                });
            }
        }
        return styleRegistrations;
    }

    /**
     * 選択した段落スタイルへ正規表現スタイルを適用する
     * @param {Document} targetDocument 対象ドキュメント
     * @param {Array<object>} styleRegistrations 適用するスタイルの組み合わせ
     * @returns {void}
     */
    function applyNestedGrepStyleSettings(targetDocument, styleRegistrations) {
        var paragraphStyleCollection = targetDocument.allParagraphStyles;
        var characterStyleCollection = targetDocument.allCharacterStyles;

        for (var registrationIndex = 0; registrationIndex < styleRegistrations.length; registrationIndex++) {
            var styleRegistration = styleRegistrations[registrationIndex];
            var paragraphStyle = findStyleByName(paragraphStyleCollection, styleRegistration.paragraph);
            var characterStyle = findStyleByName(characterStyleCollection, styleRegistration.character);
            if (!paragraphStyle || !characterStyle) continue;

            for (var grepStyleIndex = paragraphStyle.nestedGrepStyles.length - 1; grepStyleIndex >= 0; grepStyleIndex--) {
                var existingNestedGrepStyle = paragraphStyle.nestedGrepStyles[grepStyleIndex];
                if (existingNestedGrepStyle.grepExpression === styleRegistration.expression) {
                    existingNestedGrepStyle.remove();
                }
            }

            var newNestedGrepStyle = paragraphStyle.nestedGrepStyles.add();
            newNestedGrepStyle.appliedCharacterStyle = characterStyle;
            newNestedGrepStyle.grepExpression = styleRegistration.expression;
        }
    }

    var paragraphStyleNames = collectVisibleStyleNames(doc.allParagraphStyles);
    var characterStyleNames = collectVisibleStyleNames(doc.allCharacterStyles);

    if (paragraphStyleNames.length === 0) {
        alert(getLabel("error.noParagraphStyles"));
        return;
    }

    var styleRegistrations = showDialog(nestedGrepRules, paragraphStyleNames, characterStyleNames, doc);
    if (!styleRegistrations) return;

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(function () {
        applyNestedGrepStyleSettings(doc, styleRegistrations);
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.applyGrepStyles"));
})();