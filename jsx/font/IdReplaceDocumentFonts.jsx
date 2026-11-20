#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdReplaceDocumentFonts
Name=ドキュメントで使用中のフォントを一括置換
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.0.6
Release-Date=2026-09-20
Description-BEGIN
ドキュメントで使用中のフォントをファミリー／スタイル単位で一覧し、
選んだフォントを別のフォントへまとめて置き換えます。
段落スタイル・文字スタイルのフォントも同時に更新できます。

Lists the fonts used in the document by family and style, and replaces
the selected ones with another font in a single pass. Paragraph and
character styles can be updated at the same time.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

ドキュメントで使用中のフォントをファミリー／スタイル単位で一覧し、
選んだフォントを別のフォントへまとめて置き換えます。
段落スタイル・文字スタイルのフォントも同時に更新できます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdReplaceDocumentFonts.md

### Overview

Lists the fonts used in the document by family and style, and replaces
the selected ones with another font in a single pass. Paragraph and
character styles can be updated at the same time.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdReplaceDocumentFonts.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdReplaceDocumentFonts";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "ドキュメントで使用中のフォントを一括置換"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdReplaceDocumentFonts.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdReplaceDocumentFonts.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/font/IdReplaceDocumentFonts.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* スタイル行の字下げ / Indent used for style rows */
    var STYLE_ROW_INDENT = "　　";

    /* PostScript名表示の初期状態 / Initial state of the PostScript-name display */
    var SHOW_POSTSCRIPT_NAME_DEFAULT = false;

    /* ［段落・文字スタイルも更新］の初期状態 / Initial state of the style-update option */
    var UPDATE_STYLES_DEFAULT = true;

    /* 使用箇所の件数を表示するか。フォントごとに検索するため、大きなドキュメントでは一覧の作成に時間がかかる
       Whether to show usage counts. Each font is searched separately, so building the list is slower on large documents */
    var SHOW_USAGE_COUNT = true;

    /* 検索の対象範囲 / Scope of the search */
    var INCLUDE_MASTER_PAGES  = true; /* 親ページ / master pages */
    var INCLUDE_HIDDEN_LAYERS = true; /* 非表示レイヤー / hidden layers */
    var INCLUDE_FOOTNOTES     = true; /* 脚注 / footnotes */

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

    var LIST_LABEL_SPACING    = 6;    /* 見出しとリストの間隔 / spacing between a label and its list */
    var OPTION_SPACING        = 4;    /* チェックボックスの間隔 / spacing between checkboxes */
    var LISTBOX_HEIGHT        = 300;  /* リストの高さ / list height */
    var LISTBOX_WIDTH_MIN     = 200;  /* リスト幅の下限 / minimum list width */
    var LISTBOX_WIDTH_MAX     = 600;  /* リスト幅の上限 / maximum list width */
    var LISTBOX_CHAR_WIDTH    = 9;    /* 1文字あたりの概算幅 / approximate width per character */
    var LISTBOX_WIDTH_PADDING = 60;   /* リスト幅の余裕 / extra width added to the list */

    // =========================================
    // ローカライズ / Localization
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
            title: { ja: "フォント置換", en: "Replace Fonts" }
        },
        fieldLabel: {
            sourceFonts: { ja: "置換元フォント（複数選択可）", en: "Source Fonts (Multiple Selection)" },
            targetFont: { ja: "置換先フォント", en: "Target Font" }
        },
        checkbox: {
            postScriptName: { ja: "PostScript名で表示", en: "Show PostScript names" },
            updateStyles: { ja: "段落・文字スタイルも更新", en: "Update paragraph and character styles" }
        },
        button: {
            close: { ja: "閉じる", en: "Close" },
            replaceAll: { ja: "全置換", en: "Replace All" },
            replace: { ja: "フォントを置換", en: "Replace Fonts" }
        },
        list: {
            notInstalled: { ja: "未インストール", en: "not installed" }
        },
        tooltip: {
            sourceFonts: {
                ja: "置換元のフォントを選びます。ファミリー名の行を選ぶと、そのファミリーのスタイルがすべて選ばれます。",
                en: "Pick the fonts to replace. Selecting a family row selects every style in that family."
            },
            targetFont: {
                ja: "置換先のフォントを選びます。ファミリー名の行は選べません。",
                en: "Pick the font to replace them with. Family rows cannot be selected."
            },
            postScriptName: {
                ja: "ファミリー名とスタイル名の代わりに、PostScript名で一覧します。",
                en: "List the fonts by PostScript name instead of family and style."
            },
            updateStyles: {
                ja: "テキストだけでなく、段落スタイル・文字スタイルに設定されたフォントも置き換えます。",
                en: "Replace the font in paragraph and character styles as well as in the text itself."
            },
            replaceAll: {
                ja: "使用中のすべてのフォントを置換先フォントに置き換えます。置換先を選んでいないときは、置換元の1つ目のフォントに揃えます。",
                en: "Replace every font in use with the target font. With no target selected, the first source font is used instead."
            },
            replace: {
                ja: "選んだ置換元フォントを、置換先フォントに置き換えます。",
                en: "Replace the selected source fonts with the target font."
            }
        },
        alert: {
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            noFontsFound: {
                ja: "ドキュメント内に使用中のフォントが見つかりません。",
                en: "No fonts in use were found in the document."
            },
            noSourceFont: {
                ja: "置換元フォントを1つ以上選択してください。",
                en: "Please select at least one source font."
            },
            noTargetFont: {
                ja: "置換先フォントを選択してください。",
                en: "Please select a target font."
            },
            selectFonts: {
                ja: "置換先または置換元フォントを選択してください。\n（または、置換元だけ選んで全置換することも可能です）",
                en: "Please select either a target font or source fonts.\n(Alternatively, select only the source fonts to replace all.)"
            },
            targetNotFound: {
                ja: "置換先フォントが見つかりません（%1）。",
                en: "Target font not found (%1)."
            }
        }
    };

    // =========================================
    // 状態 / State
    // =========================================

    var doc = null;
    var mainDialog = null;
    var sourceFontListBox = null;
    var targetFontListBox = null;
    var postScriptNameCheckbox = null;
    var updateStylesCheckbox = null;

    /* 収集した使用中フォント / Fonts collected from the document */
    var usedFontMap = {};

    /* リストに並べるフォント（ファミリー見出しを含む）/ Fonts listed in the boxes, family headers included */
    var flatFontList = [];

    /* 選択の再帰更新を防ぐフラグ / Guard against recursive selection updates */
    var isUpdatingSelection = false;

    // =========================================
    // 検索・置換の下ごしらえ / Find and change setup
    // =========================================

    /**
     * 検索・置換の設定を初期化する
     * @returns {void}
     */
    function resetFindChangePreferences() {
        app.findTextPreferences = NothingEnum.NOTHING;
        app.changeTextPreferences = NothingEnum.NOTHING;
    }

    /**
     * 検索範囲のオプションをそろえる
     * @param {boolean} includesLocked - ロックされたレイヤー・ストーリーを含めるか
     * @returns {void}
     */
    function setupFindChangeOptions(includesLocked) {
        var options = app.findChangeTextOptions;
        options.includeMasterPages = INCLUDE_MASTER_PAGES;
        options.includeHiddenLayers = INCLUDE_HIDDEN_LAYERS;
        options.includeFootnotes = INCLUDE_FOOTNOTES;
        options.includeLockedLayersForFind = includesLocked;
        options.includeLockedStoriesForFind = includesLocked;
    }

    /**
     * 指定したフォントを使っているテキストを検索する
     * @param {Font} font - 検索するフォント
     * @returns {Array} 見つかったテキストの配列
     */
    function findTextOfFont(font) {
        resetFindChangePreferences();
        setupFindChangeOptions(true);
        app.findTextPreferences.appliedFont = font;
        var found = doc.findText();
        resetFindChangePreferences();
        return found;
    }

    // =========================================
    // フォントの収集 / Collecting fonts
    // =========================================

    /**
     * スタイルに設定されているフォント名を取り出す
     * @param {ParagraphStyle|CharacterStyle} style - 対象スタイル
     * @returns {string} 「ファミリー名＋タブ＋スタイル名」。未設定なら空文字列
     */
    function getStyleFontName(style) {
        var appliedFont = style.appliedFont;
        var familyName = "";

        if (appliedFont && appliedFont.fontFamily) {
            familyName = appliedFont.fontFamily;
        } else if (appliedFont && appliedFont.constructor === String) {
            familyName = String(appliedFont);
        }
        if (familyName === "") return "";

        var styleName = style.fontStyle;
        if (!styleName || styleName.constructor !== String) return "";
        return familyName + "\t" + String(styleName);
    }

    /**
     * ドキュメント内のすべての段落スタイル・文字スタイルを集める
     * @returns {Array} スタイルの配列（スタイルグループ内も含む）
     */
    function collectAllStyles() {
        return doc.allParagraphStyles.concat(doc.allCharacterStyles);
    }

    /**
     * 段落・文字スタイルが使っているフォントを数える
     * @returns {object} フォント名 → 件数のマップ
     */
    function collectStyleFontCounts() {
        var counts = {};
        var styles = collectAllStyles();

        for (var i = 0; i < styles.length; i++) {
            var fontName = getStyleFontName(styles[i]);
            if (fontName === "") continue;
            counts[fontName] = (counts[fontName] || 0) + 1;
        }
        return counts;
    }

    /**
     * PostScript名を取り出す（取得できないときはフォント名で代用）
     * @param {Font} font - 対象フォント
     * @param {string} fallbackName - 取得できないときに使う名前
     * @returns {string} PostScript名
     */
    function getPostScriptName(font, fallbackName) {
        /* 環境にないフォントは PostScript 名を返さないことがある / Fonts that are not installed may not expose a PostScript name */
        var postScriptName = "";
        try {
            postScriptName = font.postscriptName;
        } catch (e) {
            postScriptName = "";
        }
        if (postScriptName && postScriptName.constructor === String) return String(postScriptName);
        return fallbackName.replace("\t", " ");
    }

    /**
     * ドキュメントで使用中のフォントをファミリー別に集める
     * @returns {object} ファミリー名 → フォント名 → フォント情報のマップ
     */
    function collectUsedFonts() {
        var styleFontCounts = collectStyleFontCounts();
        var fonts = doc.fonts.everyItem().getElements();
        var fontMap = {};

        for (var i = 0; i < fonts.length; i++) {
            var font = fonts[i];
            var family = font.fontFamily;
            var styleName = font.fontStyleName;
            var fontName = family + "\t" + styleName;

            if (!fontMap[family]) fontMap[family] = {};
            fontMap[family][fontName] = {
                font: font,
                name: fontName,
                family: family,
                style: styleName,
                postScriptName: getPostScriptName(font, fontName),
                isMissing: (font.status !== FontStatus.INSTALLED),
                textCount: SHOW_USAGE_COUNT ? findTextOfFont(font).length : 0,
                styleCount: styleFontCounts[fontName] || 0
            };
        }
        return fontMap;
    }

    // =========================================
    // リストの組み立て / Building the list
    // =========================================

    /**
     * ［段落・文字スタイルも更新］がオンかを調べる
     * @returns {boolean} オンなら true
     */
    function updatesStyles() {
        return updateStylesCheckbox ? updateStylesCheckbox.value : UPDATE_STYLES_DEFAULT;
    }

    /**
     * 件数と未インストールの印をまとめた接尾辞を作る（日本語は全角括弧、英語は半角括弧）
     * @param {object} fontInfo - フォント情報
     * @returns {string} 括弧付きの接尾辞。示すものがなければ空文字列
     */
    function labelSuffix(fontInfo) {
        var parts = [];

        if (SHOW_USAGE_COUNT) {
            var count = fontInfo.textCount + (updatesStyles() ? fontInfo.styleCount : 0);
            parts.push(String(count));
        }
        if (fontInfo.isMissing) parts.push(getLabel(LABELS.list.notInstalled));
        if (parts.length === 0) return "";

        var body = parts.join(uiLang === "ja" ? "／" : " / ");
        return (uiLang === "ja") ? "（" + body + "）" : " (" + body + ")";
    }

    /**
     * リスト1行分のデータを作る
     * @param {object} fontInfo - フォント情報
     * @param {string} labelBody - 接尾辞を除いた表示名
     * @returns {object} リスト行
     */
    function createFontRow(fontInfo, labelBody) {
        return {
            label: labelBody + labelSuffix(fontInfo),
            name: fontInfo.name,
            family: fontInfo.family,
            style: fontInfo.style,
            font: fontInfo.font,
            isMissing: fontInfo.isMissing,
            isHeader: false
        };
    }

    /**
     * 収集したフォントをリスト表示用の1次元配列にする
     * @param {object} fontMap - collectUsedFonts() が返したマップ
     * @returns {Array<object>} ファミリー見出しとスタイル行を並べた配列（PostScript名表示中は見出しなしの1フォント1行）
     */
    function buildFlatFontList(fontMap) {
        var showsPostScriptName = postScriptNameCheckbox ? postScriptNameCheckbox.value : SHOW_POSTSCRIPT_NAME_DEFAULT;
        var fontList = [];

        for (var family in fontMap) {
            var styles = fontMap[family];
            var fontNames = [];
            for (var fontName in styles) {
                fontNames.push(fontName);
            }

            /* PostScript名表示のときは見出しを立てず、1フォント1行で並べる / In PostScript-name mode, list one row per font with no headers */
            if (showsPostScriptName) {
                for (var p = 0; p < fontNames.length; p++) {
                    var psFont = styles[fontNames[p]];
                    fontList.push(createFontRow(psFont, psFont.postScriptName));
                }
                continue;
            }

            /* スタイルが1つだけのファミリーは見出しを立てず1行で見せる / Show single-style families on one row */
            if (fontNames.length === 1) {
                var onlyFont = styles[fontNames[0]];
                fontList.push(createFontRow(onlyFont, onlyFont.family + " " + onlyFont.style));
                continue;
            }

            fontList.push({ label: family, family: family, isHeader: true });
            for (var i = 0; i < fontNames.length; i++) {
                var font = styles[fontNames[i]];
                fontList.push(createFontRow(font, STYLE_ROW_INDENT + font.style));
            }
        }
        return fontList;
    }

    /**
     * 表示形式に合わせてリストを作り直す（選択はフォント名で引き継ぐ）
     * @returns {void}
     */
    function rebuildFontList() {
        var previousSourceFontNames = getSelectedSourceFontNames();
        var previousTargetFontName = getSelectedTargetFontName();

        flatFontList = buildFlatFontList(usedFontMap);
        populateFontListBoxes();
        restoreSelection(previousSourceFontNames, previousTargetFontName);
    }

    /**
     * フォント一覧を集め直してリストを作り直す
     * @returns {void}
     */
    function reloadFontList() {
        usedFontMap = collectUsedFonts();
        rebuildFontList();
    }

    // =========================================
    // 置換処理 / Replacing fonts
    // =========================================

    /**
     * フォント名から Font を取得する
     * @param {string} fontName - フォント名（ファミリー名＋タブ＋スタイル名）
     * @returns {Font|null} 見つからなければ null
     */
    function findFontByName(fontName) {
        var font = doc.fonts.itemByName(fontName);
        return font.isValid ? font : null;
    }

    /**
     * 置換元として選択されているフォント名を取り出す（見出し行は除く）
     * @returns {Array<string>} フォント名の配列
     */
    function getSelectedSourceFontNames() {
        var fontNames = [];
        if (!sourceFontListBox || !sourceFontListBox.selection) return fontNames;

        for (var i = 0; i < sourceFontListBox.selection.length; i++) {
            var listEntry = flatFontList[sourceFontListBox.selection[i].index];
            if (!listEntry.isHeader) fontNames.push(listEntry.name);
        }
        return fontNames;
    }

    /**
     * 置換先リストでフォント（見出し行以外）が選ばれているか調べる
     * @returns {boolean} 選ばれていれば true
     */
    function hasTargetFontSelection() {
        if (!targetFontListBox || !targetFontListBox.selection) return false;
        return !flatFontList[targetFontListBox.selection.index].isHeader;
    }

    /**
     * 置換先として選択されているフォント名を取り出す（見出し行は除く）
     * @returns {string} 未選択・見出し行のときは空文字列
     */
    function getSelectedTargetFontName() {
        if (!hasTargetFontSelection()) return "";
        return flatFontList[targetFontListBox.selection.index].name;
    }

    /**
     * 置換先として選択されているフォントを取り出す
     * @returns {Font|null} 未選択・見出し行のときは null
     */
    function getSelectedTargetFont() {
        var fontName = getSelectedTargetFontName();
        if (fontName === "") return null;

        var targetFont = findFontByName(fontName);
        if (!targetFont) alert(getLabel(LABELS.alert.targetNotFound, [fontName.replace("\t", " ")]));
        return targetFont;
    }

    /**
     * 使用中フォントの名前をすべて集める
     * @param {string} [excludedFontName] - 除外するフォント名
     * @returns {Array<string>} フォント名の配列
     */
    function collectAllFontNames(excludedFontName) {
        var fontNames = [];
        for (var i = 0; i < flatFontList.length; i++) {
            if (flatFontList[i].isHeader) continue;
            if (flatFontList[i].name === excludedFontName) continue;
            fontNames.push(flatFontList[i].name);
        }
        return fontNames;
    }

    /**
     * テキストに使われているフォントを置き換える（ロックされたレイヤー・ストーリーは対象外）
     * @param {string} sourceFontName - 置換元のフォント名
     * @param {Font} targetFont - 置換先フォント
     * @returns {void}
     */
    function changeTextFont(sourceFontName, targetFont) {
        var sourceFont = findFontByName(sourceFontName);
        if (!sourceFont) return;

        resetFindChangePreferences();
        setupFindChangeOptions(false);
        app.findTextPreferences.appliedFont = sourceFont;
        app.changeTextPreferences.appliedFont = targetFont;
        doc.changeText();
        resetFindChangePreferences();
    }

    /**
     * 段落・文字スタイルに設定されたフォントを置き換える
     * @param {string} sourceFontName - 置換元のフォント名
     * @param {Font} targetFont - 置換先フォント
     * @returns {void}
     */
    function changeStyleFont(sourceFontName, targetFont) {
        var styles = collectAllStyles();

        for (var i = 0; i < styles.length; i++) {
            if (getStyleFontName(styles[i]) !== sourceFontName) continue;
            /* ファミリーとスタイルは別のプロパティなので、それぞれに入れる / Family and style are separate properties, so set both */
            styles[i].appliedFont = targetFont.fontFamily;
            styles[i].fontStyle = targetFont.fontStyleName;
        }
    }

    /**
     * 指定したフォントを置換先フォントに置き換える
     * @param {Array<string>} sourceFontNames - 置換元のフォント名
     * @param {Font} targetFont - 置換先フォント
     * @returns {void}
     */
    function replaceFonts(sourceFontNames, targetFont) {
        if (!targetFont || sourceFontNames.length === 0) return;

        var targetFontName = targetFont.fontFamily + "\t" + targetFont.fontStyleName;
        var updatesStyleFont = updatesStyles();

        for (var i = 0; i < sourceFontNames.length; i++) {
            if (sourceFontNames[i] === targetFontName) continue;
            changeTextFont(sourceFontNames[i], targetFont);
            if (updatesStyleFont) changeStyleFont(sourceFontNames[i], targetFont);
        }
        reloadFontList();
    }

    // =========================================
    // イベントハンドラー / Event handlers
    // =========================================

    /**
     * 置換元リストの選択を整える（見出し行はファミリー内の全スタイルに展開する）
     * @returns {void}
     */
    function handleSourceFontSelection() {
        if (isUpdatingSelection || !sourceFontListBox.selection) return;

        var expandedSelection = [];
        for (var i = 0; i < sourceFontListBox.selection.length; i++) {
            var selectedIndex = sourceFontListBox.selection[i].index;
            var listEntry = flatFontList[selectedIndex];

            if (!listEntry.isHeader) {
                expandedSelection.push(sourceFontListBox.items[selectedIndex]);
                continue;
            }
            for (var j = 0; j < flatFontList.length; j++) {
                if (!flatFontList[j].isHeader && flatFontList[j].family === listEntry.family) {
                    expandedSelection.push(sourceFontListBox.items[j]);
                }
            }
        }

        isUpdatingSelection = true;
        sourceFontListBox.selection = expandedSelection;
        isUpdatingSelection = false;
    }

    /**
     * 置換先リストで見出し行が選ばれたら選択を解除する
     * @returns {void}
     */
    function handleTargetFontSelection() {
        if (isUpdatingSelection || !targetFontListBox.selection) return;
        if (flatFontList[targetFontListBox.selection.index].isHeader) {
            targetFontListBox.selection = null;
        }
    }

    /**
     * ［PostScript名で表示］［段落・文字スタイルも更新］：表示を切り替えてリストを作り直す
     * @returns {void}
     */
    function handleDisplayModeChange() {
        rebuildFontList();
    }

    /**
     * ［フォントを置換］：選んだ置換元フォントを置換先フォントに置き換える
     * @returns {void}
     */
    function handleReplaceClick() {
        var sourceFontNames = getSelectedSourceFontNames();
        if (sourceFontNames.length === 0) {
            alert(getLabel(LABELS.alert.noSourceFont));
            return;
        }
        if (!hasTargetFontSelection()) {
            alert(getLabel(LABELS.alert.noTargetFont));
            return;
        }
        replaceFonts(sourceFontNames, getSelectedTargetFont());
    }

    /**
     * ［全置換］：使用中のすべてのフォントを1つのフォントにそろえる
     * @returns {void}
     */
    function handleReplaceAllClick() {
        /* 置換先が選ばれていれば、それにすべてをそろえる / Unify on the target font when one is selected */
        var targetFont = getSelectedTargetFont();
        if (targetFont) {
            replaceFonts(collectAllFontNames(), targetFont);
            return;
        }

        /* 置換先がなければ、置換元の1つ目にそろえる / Otherwise unify on the first source font */
        var sourceFontNames = getSelectedSourceFontNames();
        if (sourceFontNames.length === 0) {
            alert(getLabel(LABELS.alert.selectFonts));
            return;
        }

        var fallbackFont = findFontByName(sourceFontNames[0]);
        if (!fallbackFont) {
            alert(getLabel(LABELS.alert.targetNotFound, [sourceFontNames[0].replace("\t", " ")]));
            return;
        }
        replaceFonts(collectAllFontNames(sourceFontNames[0]), fallbackFont);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 見出し付きのフォントリストを1カラム分追加する
     * @param {Group} parent - 追加先のグループ
     * @param {object} labelSet - 見出しのラベル（ja/en）
     * @param {object} tooltipSet - ツールチップのラベル（ja/en）
     * @param {boolean} allowsMultiple - 複数選択を許可するか
     * @returns {ListBox} 追加したリストボックス
     */
    function addFontListColumn(parent, labelSet, tooltipSet, allowsMultiple) {
        var columnGroup = parent.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];
        columnGroup.spacing = LIST_LABEL_SPACING;
        columnGroup.add("statictext", undefined, labelText(labelSet));

        var fontListBox = columnGroup.add("listbox", undefined, [], { multiselect: allowsMultiple });
        fontListBox.preferredSize.height = LISTBOX_HEIGHT;
        fontListBox.tabEnabled = true;
        fontListBox.helpTip = getLabel(tooltipSet);
        return fontListBox;
    }

    /**
     * ダイアログを組み立てる
     * @returns {void}
     */
    function buildDialog() {
        mainDialog = new Window("dialog", getLabel(LABELS.dialog.title));
        setupWindow(mainDialog);

        var listGroup = mainDialog.add("group");
        listGroup.orientation = "row";
        listGroup.alignChildren = ["fill", "top"];
        listGroup.spacing = COLUMN_SPACING;

        sourceFontListBox = addFontListColumn(listGroup, LABELS.fieldLabel.sourceFonts, LABELS.tooltip.sourceFonts, true);
        targetFontListBox = addFontListColumn(listGroup, LABELS.fieldLabel.targetFont, LABELS.tooltip.targetFont, false);

        /* 表示オプション / Display options */
        var optionGroup = mainDialog.add("group");
        optionGroup.orientation = "column";
        optionGroup.alignChildren = ["left", "center"];
        optionGroup.spacing = OPTION_SPACING;

        postScriptNameCheckbox = optionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.postScriptName));
        postScriptNameCheckbox.value = SHOW_POSTSCRIPT_NAME_DEFAULT;
        postScriptNameCheckbox.helpTip = getLabel(LABELS.tooltip.postScriptName);

        updateStylesCheckbox = optionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.updateStyles));
        updateStylesCheckbox.value = UPDATE_STYLES_DEFAULT;
        updateStylesCheckbox.helpTip = getLabel(LABELS.tooltip.updateStyles);

        /* ボタン行（閉じるは左、置換系は右）/ Button row: Close on the left, replace buttons on the right */
        var buttonRow = addButtonRow(mainDialog);
        var btnClose = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.close), { name: "cancel" });
        var btnReplaceAll = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.replaceAll));
        var btnReplace = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.replace), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        btnReplaceAll.helpTip = getLabel(LABELS.tooltip.replaceAll);
        btnReplace.helpTip = getLabel(LABELS.tooltip.replace);

        sourceFontListBox.onChange = handleSourceFontSelection;
        targetFontListBox.onChange = handleTargetFontSelection;
        postScriptNameCheckbox.onClick = handleDisplayModeChange;
        updateStylesCheckbox.onClick = handleDisplayModeChange;
        btnReplaceAll.onClick = handleReplaceAllClick;
        btnReplace.onClick = handleReplaceClick;
    }

    /**
     * 2つのリストにフォント一覧を流し込み、幅をそろえる
     * @returns {void}
     */
    function populateFontListBoxes() {
        var listBoxWidth = calculateListBoxWidth(flatFontList);

        isUpdatingSelection = true;
        sourceFontListBox.removeAll();
        targetFontListBox.removeAll();
        for (var i = 0; i < flatFontList.length; i++) {
            sourceFontListBox.add("item", flatFontList[i].label);
            targetFontListBox.add("item", flatFontList[i].label);
        }
        isUpdatingSelection = false;

        sourceFontListBox.preferredSize.width = listBoxWidth;
        targetFontListBox.preferredSize.width = listBoxWidth;
    }

    /**
     * フォント名を手がかりに、作り直したリストの選択を元に戻す
     * @param {Array<string>} sourceFontNames - 置換元として選択されていたフォント名
     * @param {string} targetFontName - 置換先として選択されていたフォント名
     * @returns {void}
     */
    function restoreSelection(sourceFontNames, targetFontName) {
        var restoredSelection = [];
        var restoredTargetIndex = -1;

        for (var i = 0; i < flatFontList.length; i++) {
            if (flatFontList[i].isHeader) continue;

            for (var j = 0; j < sourceFontNames.length; j++) {
                if (flatFontList[i].name === sourceFontNames[j]) {
                    restoredSelection.push(sourceFontListBox.items[i]);
                    break;
                }
            }
            if (restoredTargetIndex === -1 && flatFontList[i].name === targetFontName) {
                restoredTargetIndex = i;
            }
        }

        isUpdatingSelection = true;
        sourceFontListBox.selection = restoredSelection;
        targetFontListBox.selection = (restoredTargetIndex === -1) ? null : restoredTargetIndex;
        isUpdatingSelection = false;
    }

    /**
     * いちばん長いラベルからリストの幅を見積もる
     * @param {Array<object>} fontList - リストに並べるフォント
     * @returns {number} リストの幅（px）
     */
    function calculateListBoxWidth(fontList) {
        var maxLength = 0;
        for (var i = 0; i < fontList.length; i++) {
            if (fontList[i].isHeader) continue;
            if (fontList[i].label.length > maxLength) maxLength = fontList[i].label.length;
        }
        var estimatedWidth = maxLength * LISTBOX_CHAR_WIDTH + LISTBOX_WIDTH_PADDING;
        return Math.min(LISTBOX_WIDTH_MAX, Math.max(LISTBOX_WIDTH_MIN, estimatedWidth));
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * スクリプトのエントリーポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        doc = app.activeDocument;

        usedFontMap = collectUsedFonts();
        flatFontList = buildFlatFontList(usedFontMap);
        if (flatFontList.length === 0) {
            alert(getLabel(LABELS.alert.noFontsFound));
            return;
        }

        buildDialog();
        populateFontListBoxes();

        /* 先頭のフォントを選んでおく / Preselect the first font */
        sourceFontListBox.selection = 0;
        targetFontListBox.selection = 0;
        handleSourceFontSelection();

        addVersionFooter(mainDialog);
        mainDialog.show();
    }

    main();

})();
