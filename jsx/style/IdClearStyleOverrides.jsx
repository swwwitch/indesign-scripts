#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdClearStyleOverrides
Name=テキスト・表・オブジェクトのスタイルオーバーライドを一括消去
Author=Gregor Fellenz (grefel)
Target-App=InDesign
Version=1.3.9
Release-Date=2020-06-09
Description-BEGIN
テキスト・表・オブジェクトに適用されているスタイルのオーバーライドをまとめて消去します。
処理する範囲はドキュメント全体・ストーリー・選択範囲から選べ、特定のスタイルが適用された箇所だけに絞り込むこともできます。

Clears style overrides from text, tables and objects in one pass.
The scope can be the whole document, a story or the current selection, and processing can be narrowed to where a particular style is applied.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

テキスト・表・オブジェクトに適用されているスタイルのオーバーライドをまとめて消去します。
処理する範囲はドキュメント全体・ストーリー・選択範囲から選べ、特定のスタイルが適用された箇所だけに絞り込むこともできます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdClearStyleOverrides.md

### Overview

Clears style overrides from text, tables and objects in one pass.
The scope can be the whole document, a story or the current selection, and processing can be narrowed to where a particular style is applied.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdClearStyleOverrides.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdClearStyleOverrides";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.9";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Gregor Fellenz (grefel)";      /* 作者 / author */
var SCRIPT_MODIFIED = "Masahiro Takano (@swwwitch)";  /* 改変 / modified by */
var SCRIPT_RELEASED = "2020-06-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "テキスト・表・オブジェクトのスタイルオーバーライドを一括消去"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdClearStyleOverrides.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdClearStyleOverrides.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdClearStyleOverrides.jsx"; /* 最新版のソース / latest source */

// Copyright (C) 2020 Gregor Fellenz <http://www.publishingx.de>
// Copyright (C) 2026 Masahiro Takano (@swwwitch)
//
// Modification notice (GNU GPL v3 section 5a) / 改変の告知（GNU GPL v3 第5条a）:
//   2020-06-09  Gregor Fellenz — original work "Clear Overrides"
//               https://github.com/grefel/clearOverrides
//   2026-09-20  Masahiro Takano — 全面的な書き直し。日本語／英語のローカライズ、
//               UIの再構成、セルスタイルでの絞り込みを追加。
//               Rewritten throughout: Japanese/English localization,
//               reorganized UI, added filtering by cell style.
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program. The full text is distributed alongside this
// script as IdClearStyleOverrides-LICENSE.txt. If it is missing, see
// <https://www.gnu.org/licenses/>.

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* ダイアログを開いたときのチェック状態 / Checkbox states on launch */
var DEFAULT_PROCESS_TEXT    = true;
var DEFAULT_PROCESS_TABLES  = false;
var DEFAULT_PROCESS_OBJECTS = false;

/* 「消去する種類」の初期選択 / Preselected override type (0:両方 1:段落のみ 2:文字のみ) */
var DEFAULT_OVERRIDE_TYPE_INDEX = 0;

/* 「対象」の初期選択 / Preselected scope (0:ドキュメント 1:ストーリー 2:選択範囲) */
var DEFAULT_SCOPE_INDEX = 0;

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
        title: { ja: "スタイルのオーバーライドを消去", en: "Clear style overrides" }
    },
    panel: {
        scope:  { ja: "対象", en: "Scope" },
        text:   { ja: "テキスト（段落スタイルと文字スタイル）", en: "Text (paragraph & character styles)" },
        table:  { ja: "表スタイルとセルスタイル", en: "Table & cell styles" },
        object: { ja: "オブジェクトスタイル", en: "Object styles" }
    },
    checkbox: {
        text:   { ja: "テキストのオーバーライドを消去", en: "Clear text overrides" },
        table:  { ja: "表のオーバーライドを消去", en: "Clear table overrides" },
        object: { ja: "オブジェクトのオーバーライドを消去", en: "Clear object overrides" }
    },
    fieldLabel: {
        paragraphStyle: { ja: "対象の段落スタイル", en: "Target paragraph style" },
        characterStyle: { ja: "対象の文字スタイル", en: "Target character style" },
        tableStyle:     { ja: "対象の表スタイル", en: "Target table style" },
        cellStyle:      { ja: "対象のセルスタイル", en: "Target cell style" },
        objectStyle:    { ja: "対象のオブジェクトスタイル", en: "Target object style" },
        overrideType:   { ja: "消去するオーバーライド", en: "Overrides to clear" }
    },
    radio: {
        scopeDocument:  { ja: "ドキュメント", en: "Document" },
        scopeStory:     { ja: "ストーリー", en: "Story" },
        scopeSelection: { ja: "選択範囲", en: "Selection" },
        bothTypes:     { ja: "両方", en: "Both" },
        paragraphOnly: { ja: "段落のみ", en: "Paragraph only" },
        characterOnly: { ja: "文字のみ", en: "Character only" }
    },
    dropdown: {
        allStyles: { ja: "すべてのスタイル", en: "All styles" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" },
        run:    { ja: "オーバーライドを消去", en: "Clear overrides" }
    },
    tooltip: {
        clearText: {
            ja: "本文・表の中のテキスト・脚注に付いた、段落スタイルと文字スタイルのオーバーライドを消去します。",
            en: "Clears paragraph and character style overrides on story text, text inside tables, and footnotes."
        },
        clearTable: {
            ja: "表スタイルのオーバーライドと、表の中のセルスタイルのオーバーライドを消去します。",
            en: "Clears table style overrides and the cell style overrides inside the table."
        },
        clearObject: {
            ja: "オブジェクトスタイルのオーバーライドを消去します。対象が「ドキュメント」のときは、マスターページ上のフレームも含みます。",
            en: "Clears object style overrides. When the scope is \"Document\", frames on parent (master) pages are included too."
        },
        scope: {
            ja: "「ストーリー」は選択箇所を含むストーリー全体（連結したテキストフレームすべて）、「選択範囲」は選択しているテキストとオブジェクトだけが対象になります。",
            en: "\"Story\" targets the whole story containing the selection, including every threaded frame. \"Selection\" targets only the selected text and objects."
        },
        overrideType: {
            ja: "「段落のみ」は行揃えや段落前後のアキなど段落レベルの手動設定だけを、「文字のみ」はフォントやサイズなど文字レベルの手動設定だけを消去します。",
            en: "\"Paragraph only\" clears paragraph-level local formatting such as alignment and space before/after; \"Character only\" clears character-level local formatting such as font and size."
        },
        styleFilter: {
            ja: "「すべてのスタイル」のままにすると、すべてのスタイルが対象になります。特定のスタイルを選ぶと、そのスタイルが適用された箇所だけに絞り込まれます。",
            en: "Leave \"All styles\" to target every style. Pick a style to narrow the scope to where that style is applied."
        }
    },
    alert: {
        noDocument:  { ja: "ドキュメントを開いてから実行してください。", en: "Please run with an active document." },
        noSelection: { ja: "対象が選択されていません。テキストやオブジェクトを選択してから実行してください。", en: "Nothing is selected. Select text or objects before running." },
        failed:      { ja: "オーバーライドの消去中にエラーが発生しました。", en: "An error occurred while clearing overrides." }
    },
    undo: {
        clearOverrides: { ja: "オーバーライドを消去", en: "Clear style overrides" }
    }
};

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

var LABEL_WIDTH    = 200;                /* 項目名の幅 / field label width */
var DROPDOWN_WIDTH = 220;                /* ドロップダウンの幅 / dropdown width */
var RADIO_SPACING  = 12;                 /* ラジオボタン間の間隔 / gap between radio buttons */

/**
 * 右揃えの項目名を行に追加する
 * @param {Group} parentGroup 追加先の行グループ
 * @param {string} labelKey ラベルキー
 * @returns {StaticText} 追加した項目名
 */
function addRowLabel(parentGroup, labelKey) {
    var rowLabel = parentGroup.add("statictext", undefined, labelText(labelKey));
    rowLabel.preferredSize.width = LABEL_WIDTH;
    rowLabel.justify = "right";
    return rowLabel;
}

// =========================================
// 検索 / Find
// =========================================

var ALL_TEXT_GREP   = "(?s).+";   /* 改行を含むすべてのテキストにマッチ / Match every text run, newlines included */
var TABLE_MARKER    = "<0016>";   /* 表のアンカー文字 U+0016 / Table anchor character U+0016 */
var ALL_STYLES      = null;       /* スタイルで絞り込まないことを表す値 / Sentinel meaning "do not filter by style" */

/* 検索の種類ごとの app プロパティ名 / app property names per search kind */
var SEARCH_MODES = {
    grep: {
        optionsKey:   "findChangeGrepOptions",
        findPrefsKey: "findGrepPreferences",
        findMethod:   "findGrep"
    },
    text: {
        optionsKey:   "findChangeTextOptions",
        findPrefsKey: "findTextPreferences",
        findMethod:   "findText"
    },
    object: {
        optionsKey:   "findChangeObjectOptions",
        findPrefsKey: "findObjectPreferences",
        findMethod:   "findObject",
        objectType:   ObjectTypes.ALL_FRAMES_TYPE
    }
};

/* 退避・復帰の対象になる検索オプション / Search options that are saved and restored */
var SEARCH_OPTION_KEYS = [
    "includeFootnotes",
    "includeHiddenLayers",
    "includeLockedLayersForFind",
    "includeLockedStoriesForFind",
    "includeMasterPages",
    "searchBackwards",
    "objectType"
];

/* 検索時に使う値。ここに無いキーは退避・復帰だけ行う / Values applied while searching */
var SEARCH_OPTION_VALUES = {
    includeFootnotes: true,
    includeHiddenLayers: true,
    includeLockedLayersForFind: false,
    includeLockedStoriesForFind: false,
    includeMasterPages: true,
    searchBackwards: false
};

/**
 * 指定した範囲を検索し、見つかったオブジェクトを返す
 * 検索オプションと検索条件は、呼び出し前の状態に必ず戻す
 * @param {Document|Story|Text} searchTarget 検索する範囲
 * @param {string} searchModeKey "grep" / "text" / "object"
 * @param {object} findProperties 検索条件のプロパティ
 * @returns {Array<object>} 見つかったオブジェクトの配列
 */
function findInTarget(searchTarget, searchModeKey, findProperties) {
    var searchMode = SEARCH_MODES[searchModeKey];
    var searchOptions = app[searchMode.optionsKey];
    var savedOptions = {};
    var optionKey;
    var i;

    /* 現在の設定を退避（そのバージョンに無いキーは触らない）/ Save current options, skipping keys this version lacks */
    for (i = 0; i < SEARCH_OPTION_KEYS.length; i++) {
        optionKey = SEARCH_OPTION_KEYS[i];
        if (searchOptions.hasOwnProperty(optionKey)) {
            savedOptions[optionKey] = searchOptions[optionKey];
        }
    }

    try {
        for (i = 0; i < SEARCH_OPTION_KEYS.length; i++) {
            optionKey = SEARCH_OPTION_KEYS[i];
            if (savedOptions.hasOwnProperty(optionKey) && SEARCH_OPTION_VALUES.hasOwnProperty(optionKey)) {
                searchOptions[optionKey] = SEARCH_OPTION_VALUES[optionKey];
            }
        }
        if (searchMode.hasOwnProperty("objectType")) {
            searchOptions.objectType = searchMode.objectType;
        }

        app[searchMode.findPrefsKey] = NothingEnum.nothing;
        app[searchMode.findPrefsKey].properties = findProperties;
        return searchTarget[searchMode.findMethod](true);
    } finally {
        app[searchMode.findPrefsKey] = NothingEnum.nothing;
        for (i = 0; i < SEARCH_OPTION_KEYS.length; i++) {
            optionKey = SEARCH_OPTION_KEYS[i];
            if (savedOptions.hasOwnProperty(optionKey)) {
                searchOptions[optionKey] = savedOptions[optionKey];
            }
        }
    }
}

// =========================================
// 対象の解決 / Resolving the scope
// =========================================

/* 選択がテキストとみなせる型 / Selection types treated as text */
var TEXT_SELECTION_TYPES = {
    Text: true,
    Word: true,
    Character: true,
    Paragraph: true,
    Line: true,
    TextColumn: true,
    TextStyleRange: true,
    InsertionPoint: true
};

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
 * 現在の選択を、テキストとページアイテムに振り分ける
 * @returns {object} texts / items を持つオブジェクト
 */
function collectSelectionTargets() {
    var textTargets = [];
    var itemTargets = [];
    var cellTargets = [];
    var selectedObjects = app.selection;

    for (var i = 0; i < selectedObjects.length; i++) {
        var selectedObject = selectedObjects[i];
        var typeName = selectedObject.constructor.name;

        if (TEXT_SELECTION_TYPES.hasOwnProperty(typeName)) {
            textTargets.push(selectedObject);
        } else if (typeName === "Cell") {
            /* 複数セル選択は1つの Cell として返るので展開する /
               A multi-cell selection arrives as one Cell; expand it */
            var resolvedCells = getSelectedCells(selectedObject);
            for (var c = 0; c < resolvedCells.length; c++) {
                cellTargets.push(resolvedCells[c]);
                textTargets.push(resolvedCells[c].texts[0]);
            }
        } else {
            itemTargets.push(selectedObject);
            /* テキストフレームは中のテキストも対象にする / A selected text frame also contributes its text */
            if (typeName === "TextFrame") {
                textTargets.push(selectedObject.texts[0]);
            }
        }
    }
    return { texts: textTargets, items: itemTargets, cells: cellTargets };
}

/**
 * 選択箇所が属するストーリーと、そのストーリーのテキストフレームを集める
 * @param {object} selectionTargets collectSelectionTargets() の戻り値
 * @returns {object} texts（ストーリー）/ items（テキストフレーム）/ cells を持つオブジェクト
 */
function collectStoryTargets(selectionTargets) {
    var storyTargets = [];
    var frameTargets = [];
    var seenStoryIds = {};

    for (var i = 0; i < selectionTargets.texts.length; i++) {
        var parentStory = selectionTargets.texts[i].parentStory;
        if (seenStoryIds.hasOwnProperty(parentStory.id)) continue;
        seenStoryIds[parentStory.id] = true;
        storyTargets.push(parentStory);

        var textContainers = parentStory.textContainers;
        for (var j = 0; j < textContainers.length; j++) {
            frameTargets.push(textContainers[j]);
        }
    }
    return { texts: storyTargets, items: frameTargets, cells: [] };
}

/**
 * 長さのないカーソル位置を取り除く
 * @param {Array<object>} textTargets テキストの配列
 * @returns {Array<object>} InsertionPoint を除いた配列
 */
function excludeInsertionPoints(textTargets) {
    var pickedTargets = [];
    for (var i = 0; i < textTargets.length; i++) {
        if (textTargets[i].constructor.name !== "InsertionPoint") pickedTargets.push(textTargets[i]);
    }
    return pickedTargets;
}

/**
 * 「対象」の選択に応じて、処理する範囲を返す
 * 選択は処理の途中で変わりうるので、消去を始める前に確定させる
 * @param {string} scopeKey "document" / "story" / "selection"
 * @param {Document} targetDocument 対象ドキュメント
 * @returns {object} texts（検索する範囲）/ items（オブジェクト。null はドキュメント全体から検索）/ cells（選択されたセル）
 */
function resolveScopeTargets(scopeKey, targetDocument) {
    if (scopeKey === "document") {
        /* items が null なら、必要になった時点でドキュメント全体を検索する / null defers the document-wide object search */
        return { texts: [targetDocument], items: null, cells: [] };
    }
    var selectionTargets = collectSelectionTargets();
    if (scopeKey === "selection") {
        /* カーソルを置いただけの状態は「選択なし」として扱う。
           ストーリーの特定には使えるので、ここでだけ取り除く /
           A bare caret counts as no selection here, though it still identifies a story */
        return {
            texts: excludeInsertionPoints(selectionTargets.texts),
            items: selectionTargets.items,
            cells: selectionTargets.cells
        };
    }
    return collectStoryTargets(selectionTargets);
}

// =========================================
// オーバーライドの消去 / Clearing overrides
// =========================================

/**
 * 適用スタイルが処理対象かどうかを判定する
 * @param {object} appliedStyle 対象に適用されているスタイル
 * @param {object} selectedStyle ダイアログで選ばれたスタイル。ALL_STYLES なら常に true
 * @returns {boolean} 処理対象なら true
 */
function isTargetStyle(appliedStyle, selectedStyle) {
    return selectedStyle === ALL_STYLES || appliedStyle.id === selectedStyle.id;
}

/**
 * テキスト検索の条件を組み立てる（スタイルでの絞り込みを含む）
 * @param {string} findWhat 検索する文字列
 * @param {object} settings ダイアログで決めた設定
 * @returns {object} 検索条件のプロパティ
 */
function buildTextFindProperties(findWhat, settings) {
    var findProperties = { findWhat: findWhat };
    if (settings.paragraphStyle !== ALL_STYLES) {
        findProperties.appliedParagraphStyle = settings.paragraphStyle;
    }
    if (settings.characterStyle !== ALL_STYLES) {
        findProperties.appliedCharacterStyle = settings.characterStyle;
    }
    return findProperties;
}

/**
 * 見つかったテキストのオーバーライドをまとめて消去する
 * @param {Array<Text>} foundTexts 消去対象のテキスト
 * @param {OverrideType} overrideType 消去するオーバーライドの種類
 * @returns {void}
 */
function clearFoundTextOverrides(foundTexts, overrideType) {
    for (var i = 0; i < foundTexts.length; i++) {
        foundTexts[i].clearOverrides(overrideType);
    }
}

/**
 * テキストのスタイルオーバーライドを消去する
 * @param {Array<object>} textTargets 検索する範囲の配列
 * @param {object} settings ダイアログで決めた設定
 * @returns {void}
 */
function clearTextOverrides(textTargets, settings) {
    for (var i = 0; i < textTargets.length; i++) {
        clearFoundTextOverrides(findInTarget(textTargets[i], "grep", buildTextFindProperties(ALL_TEXT_GREP, settings)), settings.overrideType);
        /* フレームやセルの中身が表だけの場合、GREP検索では拾えないので別途処理 / A frame or cell holding only a table is not matched by GREP */
        clearFoundTextOverrides(findInTarget(textTargets[i], "text", buildTextFindProperties(TABLE_MARKER, settings)), settings.overrideType);
    }
}

/**
 * 表の中のセルスタイルのオーバーライドを消去する
 * @param {Table} targetTable 対象の表
 * @param {object} selectedCellStyle 対象のセルスタイル。ALL_STYLES ならすべて
 * @returns {void}
 */
function clearCellOverrides(targetTable, selectedCellStyle) {
    /* 絞り込みが無ければ一括で処理する / Clear in one call when no style filter is set */
    if (selectedCellStyle === ALL_STYLES) {
        targetTable.cells.everyItem().clearCellStyleOverrides();
        return;
    }
    var tableCells = targetTable.cells.everyItem().getElements();
    for (var i = 0; i < tableCells.length; i++) {
        if (isTargetStyle(tableCells[i].appliedCellStyle, selectedCellStyle)) {
            tableCells[i].clearCellStyleOverrides();
        }
    }
}

/**
 * 表スタイルとセルスタイルのオーバーライドを消去する
 * 選択されたセルは表のアンカー文字では見つからないので、別経路で処理する
 * @param {object} scopeTargets resolveScopeTargets() が返した処理範囲
 * @param {object} selectedTableStyle 対象の表スタイル。ALL_STYLES ならすべて
 * @param {object} selectedCellStyle 対象のセルスタイル。ALL_STYLES ならすべて
 * @returns {void}
 */
function clearTableOverrides(scopeTargets, selectedTableStyle, selectedCellStyle) {
    /* 同じ表を二度処理しないための控え / Keeps a table from being processed twice */
    var processedTableIds = {};
    var i;

    /* 選択されたセル：そのセルだけ消し、属する表の表スタイルは表ごとに1回 /
       Selected cells: clear just those cells, and their table's style once per table */
    for (i = 0; i < scopeTargets.cells.length; i++) {
        var targetCell = scopeTargets.cells[i];
        if (isTargetStyle(targetCell.appliedCellStyle, selectedCellStyle)) {
            targetCell.clearCellStyleOverrides();
        }
        var cellTable = findParentTable(targetCell);
        if (cellTable === null || processedTableIds.hasOwnProperty(cellTable.id)) continue;
        processedTableIds[cellTable.id] = true;
        if (isTargetStyle(cellTable.appliedTableStyle, selectedTableStyle)) {
            cellTable.clearTableStyleOverrides();
        }
    }

    /* 検索で見つかる表：表全体のセルを処理する / Tables found by search: every cell in them */
    for (i = 0; i < scopeTargets.texts.length; i++) {
        var foundMarkers = findInTarget(scopeTargets.texts[i], "text", { findWhat: TABLE_MARKER });
        for (var j = 0; j < foundMarkers.length; j++) {
            var foundTable = foundMarkers[j].tables[0];
            if (processedTableIds.hasOwnProperty(foundTable.id)) continue;
            processedTableIds[foundTable.id] = true;
            if (isTargetStyle(foundTable.appliedTableStyle, selectedTableStyle)) {
                foundTable.clearTableStyleOverrides();
                clearCellOverrides(foundTable, selectedCellStyle);
            }
        }
    }
}

/**
 * オブジェクトスタイルのオーバーライドを消去する
 * @param {Document} targetDocument 対象ドキュメント
 * @param {Array<PageItem>} objectTargets 対象のページアイテム。null ならドキュメント全体から検索する
 * @param {object} selectedObjectStyle 対象のオブジェクトスタイル。ALL_STYLES ならすべて
 * @returns {void}
 */
function clearObjectOverrides(targetDocument, objectTargets, selectedObjectStyle) {
    var candidateItems = (objectTargets === null) ? findInTarget(targetDocument, "object", {}) : objectTargets;
    for (var i = 0; i < candidateItems.length; i++) {
        /* ガイドなどオブジェクトスタイルを持たない選択は飛ばす /
           Skip a selection that has no object style, such as a guide */
        if (!candidateItems[i].hasOwnProperty("appliedObjectStyle")) continue;
        if (isTargetStyle(candidateItems[i].appliedObjectStyle, selectedObjectStyle)) {
            candidateItems[i].clearObjectStyleOverrides();
        }
    }
}

/**
 * 設定に従って、選ばれた種類のオーバーライドをすべて消去する
 * @param {Document} targetDocument 対象ドキュメント
 * @param {object} scopeTargets resolveScopeTargets() が返した処理範囲
 * @param {object} settings ダイアログで決めた設定
 * @returns {void}
 */
function clearStyleOverrides(targetDocument, scopeTargets, settings) {
    if (settings.processObjects) clearObjectOverrides(targetDocument, scopeTargets.items, settings.objectStyle);
    if (settings.processTables)  clearTableOverrides(scopeTargets, settings.tableStyle, settings.cellStyle);
    if (settings.processText)    clearTextOverrides(scopeTargets.texts, settings);
}

// =========================================
// UI / Dialog
// =========================================

/* 「対象」の選択肢 / Scope choices */
var SCOPE_CHOICES = [
    { labelKey: "radio.scopeDocument",  value: "document" },
    { labelKey: "radio.scopeStory",     value: "story" },
    { labelKey: "radio.scopeSelection", value: "selection" }
];

/* 「消去する種類」の選択肢 / Override type choices */
var OVERRIDE_TYPE_CHOICES = [
    { labelKey: "radio.bothTypes",     value: OverrideType.ALL },
    { labelKey: "radio.paragraphOnly", value: OverrideType.PARAGRAPH_ONLY },
    { labelKey: "radio.characterOnly", value: OverrideType.CHARACTER_ONLY }
];

/**
 * ラジオボタンの並びを追加する（排他になるよう専用グループに入れる）
 * @param {Panel|Group} parentContainer 追加先
 * @param {Array<object>} choices labelKey / value を持つ選択肢の配列
 * @param {number} defaultIndex 初期選択の位置
 * @param {string} tooltipKey tooltipのラベルキー
 * @param {string} [alignment] 横方向の配置。省略時は "left"
 * @returns {Array<RadioButton>} 追加したラジオボタン
 */
function addRadioRow(parentContainer, choices, defaultIndex, tooltipKey, alignment) {
    var radioGroup = parentContainer.add("group");
    setupRow(radioGroup, alignment || "left");
    radioGroup.spacing = RADIO_SPACING;

    var radioButtons = [];
    for (var i = 0; i < choices.length; i++) {
        var radioButton = radioGroup.add("radiobutton", undefined, getLabel(choices[i].labelKey));
        radioButton.helpTip = getLabel(tooltipKey);
        radioButtons.push(radioButton);
    }
    radioButtons[defaultIndex].value = true;
    return radioButtons;
}

/**
 * ラジオボタンで選ばれている値を取得する
 * @param {Array<RadioButton>} radioButtons 対象のラジオボタン
 * @param {Array<object>} choices 対応する選択肢の配列
 * @returns {object} 選ばれている選択肢の value。未選択なら先頭の value
 */
function getSelectedRadioValue(radioButtons, choices) {
    for (var i = 0; i < radioButtons.length; i++) {
        if (radioButtons[i].value) return choices[i].value;
    }
    return choices[0].value;
}

/**
 * 「対象」パネルを追加する
 * @param {Window} parentWindow 追加先のウィンドウ
 * @returns {Array<RadioButton>} 追加したラジオボタン
 */
function addScopePanel(parentWindow) {
    var scopePanel = parentWindow.add("panel", undefined, getLabel("panel.scope"));
    setupPanel(scopePanel);
    return addRadioRow(scopePanel, SCOPE_CHOICES, DEFAULT_SCOPE_INDEX, "tooltip.scope", "center");
}

/**
 * スタイルグループ名を含めた表示名を返す（グループは : 区切り）
 * @param {ParagraphStyle|CharacterStyle|TableStyle|CellStyle|ObjectStyle} style 対象のスタイル
 * @returns {string} グループ名を含む表示名
 */
function getStyleDisplayName(style) {
    var displayName = style.name;
    var container = style.parent;
    while (container.constructor.name.match(/Group$/)) {
        displayName = container.name + ":" + displayName;
        container = container.parent;
    }
    return displayName;
}

/**
 * ドロップダウンに [すべて] と各スタイルを並べる
 * @param {DropDownList} styleDropdown 対象のドロップダウン
 * @param {Array<object>} styleCollection 並べるスタイルの配列
 * @returns {void}
 */
function fillStyleDropdown(styleDropdown, styleCollection) {
    styleDropdown.add("item", getLabel("dropdown.allStyles"));
    for (var i = 0; i < styleCollection.length; i++) {
        var styleItem = styleDropdown.add("item", getStyleDisplayName(styleCollection[i]));
        styleItem.styleRef = styleCollection[i];
    }
    styleDropdown.selection = 0;
}

/**
 * ドロップダウンで選ばれているスタイルを取得する
 * @param {DropDownList} styleDropdown 対象のドロップダウン
 * @returns {object} 選ばれたスタイル。先頭の [すべて] の場合は ALL_STYLES
 */
function getSelectedStyle(styleDropdown) {
    var selectedItem = styleDropdown.selection;
    if (!selectedItem || selectedItem.index === 0) return ALL_STYLES;
    return selectedItem.styleRef;
}

/**
 * 種類ごとのパネル（チェックボックス＋スタイル絞り込み行）を追加する
 * @param {Window} parentWindow 追加先のウィンドウ
 * @param {string} panelLabelKey パネル見出しのラベルキー
 * @param {string} checkboxLabelKey チェックボックスのラベルキー
 * @param {string} tooltipKey チェックボックスのtooltipキー
 * @param {boolean} initialValue チェックボックスの初期状態
 * @returns {object} panel / checkbox / optionRows を持つオブジェクト
 */
function addOverrideSection(parentWindow, panelLabelKey, checkboxLabelKey, tooltipKey, initialValue) {
    var sectionPanel = parentWindow.add("panel", undefined, getLabel(panelLabelKey));
    setupPanel(sectionPanel);

    var enableCheckbox = sectionPanel.add("checkbox", undefined, getLabel(checkboxLabelKey));
    enableCheckbox.alignment = ["left", "top"];
    enableCheckbox.value = initialValue;
    enableCheckbox.helpTip = getLabel(tooltipKey);

    var overrideSection = {
        panel: sectionPanel,
        checkbox: enableCheckbox,
        optionRows: [],
        onToggle: null   /* チェック状態が変わったときに呼ぶ / Called when the checkbox changes */
    };

    enableCheckbox.onClick = function () {
        for (var i = 0; i < overrideSection.optionRows.length; i++) {
            overrideSection.optionRows[i].enabled = enableCheckbox.value;
        }
        if (overrideSection.onToggle) overrideSection.onToggle();
    };

    return overrideSection;
}

/**
 * 処理パネルにスタイル絞り込みの行を追加する
 * @param {object} overrideSection addOverrideSection() が返したオブジェクト
 * @param {string} fieldLabelKey 項目名のラベルキー
 * @param {Array<object>} styleCollection ドロップダウンに並べるスタイルの配列
 * @returns {DropDownList} 追加したドロップダウン
 */
function addStyleFilterRow(overrideSection, fieldLabelKey, styleCollection) {
    var filterRow = overrideSection.panel.add("group");
    setupRow(filterRow, "fill");
    filterRow.enabled = overrideSection.checkbox.value;
    addRowLabel(filterRow, fieldLabelKey);

    var styleDropdown = filterRow.add("dropdownlist");
    styleDropdown.preferredSize.width = DROPDOWN_WIDTH;
    styleDropdown.maximumSize.width = DROPDOWN_WIDTH;
    styleDropdown.alignment = ["fill", "center"];
    styleDropdown.helpTip = getLabel("tooltip.styleFilter");
    fillStyleDropdown(styleDropdown, styleCollection);

    overrideSection.optionRows.push(filterRow);
    return styleDropdown;
}

/**
 * 処理パネルに「消去する種類」の行を追加する
 * @param {object} overrideSection addOverrideSection() が返したオブジェクト
 * @returns {Array<RadioButton>} 追加したラジオボタン
 */
function addOverrideTypeRow(overrideSection) {
    var typeRow = overrideSection.panel.add("group");
    setupRow(typeRow, "fill");
    typeRow.enabled = overrideSection.checkbox.value;
    addRowLabel(typeRow, "fieldLabel.overrideType");

    var typeRadios = addRadioRow(typeRow, OVERRIDE_TYPE_CHOICES, DEFAULT_OVERRIDE_TYPE_INDEX, "tooltip.overrideType");
    overrideSection.optionRows.push(typeRow);
    return typeRadios;
}

/**
 * 設定ダイアログを表示する
 * @param {Document} targetDocument 対象ドキュメント
 * @returns {object} 設定オブジェクト。キャンセルされた場合は null
 */
function showSettingsDialog(targetDocument) {
    var settingsDialog = new Window("dialog", getLabel("dialog.title"));
    setupWindow(settingsDialog);

    var scopeRadios = addScopePanel(settingsDialog);

    var textSection = addOverrideSection(settingsDialog, "panel.text", "checkbox.text", "tooltip.clearText", DEFAULT_PROCESS_TEXT);
    var paragraphStyleDropdown = addStyleFilterRow(textSection, "fieldLabel.paragraphStyle", targetDocument.allParagraphStyles);
    var characterStyleDropdown = addStyleFilterRow(textSection, "fieldLabel.characterStyle", targetDocument.allCharacterStyles);
    var overrideTypeRadios = addOverrideTypeRow(textSection);

    var tableSection = addOverrideSection(settingsDialog, "panel.table", "checkbox.table", "tooltip.clearTable", DEFAULT_PROCESS_TABLES);
    var tableStyleDropdown = addStyleFilterRow(tableSection, "fieldLabel.tableStyle", targetDocument.allTableStyles);
    var cellStyleDropdown = addStyleFilterRow(tableSection, "fieldLabel.cellStyle", targetDocument.allCellStyles);

    var objectSection = addOverrideSection(settingsDialog, "panel.object", "checkbox.object", "tooltip.clearObject", DEFAULT_PROCESS_OBJECTS);
    var objectStyleDropdown = addStyleFilterRow(objectSection, "fieldLabel.objectStyle", targetDocument.allObjectStyles);

    var buttonRow = addButtonRow(settingsDialog);
    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnRun = buttonRow.rightGroup.add("button", undefined, getLabel("button.run"), { name: "ok" });
    alignRightOnlyButtonRow(buttonRow);

    /* 3つとも処理しない設定では実行できないようにする / Nothing to process means nothing to run */
    var overrideSections = [textSection, tableSection, objectSection];
    var updateRunButton = function () {
        var anyChecked = false;
        for (var i = 0; i < overrideSections.length; i++) {
            if (overrideSections[i].checkbox.value) anyChecked = true;
        }
        btnRun.enabled = anyChecked;
    };
    for (var i = 0; i < overrideSections.length; i++) {
        overrideSections[i].onToggle = updateRunButton;
    }
    updateRunButton();

    addVersionFooter(settingsDialog);
    if (settingsDialog.show() !== 1) return null;

    return {
        scope:          getSelectedRadioValue(scopeRadios, SCOPE_CHOICES),
        processText:    textSection.checkbox.value,
        paragraphStyle: getSelectedStyle(paragraphStyleDropdown),
        characterStyle: getSelectedStyle(characterStyleDropdown),
        overrideType:   getSelectedRadioValue(overrideTypeRadios, OVERRIDE_TYPE_CHOICES),
        processTables:  tableSection.checkbox.value,
        tableStyle:     getSelectedStyle(tableStyleDropdown),
        cellStyle:      getSelectedStyle(cellStyleDropdown),
        processObjects: objectSection.checkbox.value,
        objectStyle:    getSelectedStyle(objectStyleDropdown)
    };
}

// =========================================
// メイン処理 / Main
// =========================================

if (app.documents.length === 0 || app.layoutWindows.length === 0) {
    alert(getLabel("alert.noDocument"));
    return;
}

var targetDocument = app.activeDocument;
var settings = showSettingsDialog(targetDocument);
if (settings === null) return;

/* 消去を始めると選択が変わりうるので、処理範囲は先に確定させる / Resolve the scope before anything changes the selection */
var scopeTargets = resolveScopeTargets(settings.scope, targetDocument);
if (settings.scope !== "document" && scopeTargets.texts.length === 0 && scopeTargets.items.length === 0) {
    alert(getLabel("alert.noSelection"));
    return;
}

/* 検索系のバージョン差による警告を抑える / Avoid version-related warnings from find/change */
var savedScriptVersion = app.scriptPreferences.version;
app.scriptPreferences.version = parseInt(app.version, 10);

try {
    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(function () {
        clearStyleOverrides(targetDocument, scopeTargets, settings);
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.clearOverrides"));
} catch (err) {
    alert(getLabel("alert.failed") + "\n" + err + " (line " + err.line + ")");
} finally {
    app.scriptPreferences.version = savedScriptVersion;
}

})();
