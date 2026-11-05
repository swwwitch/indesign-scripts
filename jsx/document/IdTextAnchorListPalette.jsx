#target indesign
#targetengine "IdTextAnchorListPalette"

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdTextAnchorListPalette
Name=テキストアンカーを一覧して表示
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.0.0
Release-Date=2026-11-06
Description-BEGIN
アクティブなドキュメントのテキストアンカーを、ページ・アンカー名・テキスト（アンカーのある段落）の3列で、常駐パレットにページ順に一覧します。
パレットを開いたまま、ダブルクリックでそのアンカーのページを表示して段落のテキストを選択でき、［更新］で一覧を読み直せます。

Lists the text anchors in the active document in page order in a floating palette, with three columns: page, anchor name and text (the paragraph that holds the anchor).
With the palette left open, double-click an anchor to show its page and select the paragraph text; Refresh reads the list again.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

アクティブなドキュメントのテキストアンカーを、ページ・アンカー名・テキスト（アンカーのある段落）の3列で、常駐パレットにページ順に一覧します。
パレットを開いたまま、ダブルクリックでそのアンカーのページを表示して段落のテキストを選択でき、［更新］で一覧を読み直せます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTextAnchorListPalette.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n23fa2f14621f

### Overview

Lists the text anchors in the active document in page order in a floating palette, with three columns: page, anchor name and text (the paragraph that holds the anchor).
With the palette left open, double-click an anchor to show its page and select the paragraph text; Refresh reads the list again.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTextAnchorListPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTextAnchorListPalette";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-06";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "テキストアンカーを一覧して表示"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTextAnchorListPalette.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTextAnchorListPalette.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n23fa2f14621f"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdTextAnchorListPalette.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 基本設定 / Settings
    // =========================================

    /* 非表示のハイパーリンク先（hidden）も一覧に入れる / Also list hidden destinations */
    var INCLUDE_HIDDEN_ANCHORS = false;

    /* テキスト欄に出す最大文字数（超えた分は「...」）/ Max characters shown in the text column (the rest becomes "...") */
    var TEXT_MAX_LENGTH = 60;

    /* 並べ替えキーの座標に足す値（負の座標を正にする）/ Offset that makes coordinates positive in the sort key */
    var SORT_COORDINATE_OFFSET = 100000;

    /* パレットの参照を残す $.global のキー（再実行で前回のパレットを閉じるため）
       / $.global key holding the palette, so a rerun can close the previous one */
    var PALETTE_GLOBAL_KEY = "__IdTextAnchorListPalette";

    /* 冒頭の #targetengine と同じ名前 / Same name as the #targetengine directive at the top */
    var PALETTE_ENGINE_NAME = "IdTextAnchorListPalette";

    /* ランチャーの $.evalFile から呼ばれると #targetengine が効かず、呼び出し元のエンジンが終わるとパレットが消える。
       このファイルを doScript で自分のエンジンに開き直す
       / Run through a launcher's $.evalFile, #targetengine is ignored and the palette vanishes with the caller's engine,
         so rerun this file with doScript in its own engine */
    if ($.engineName !== PALETTE_ENGINE_NAME) {
        app.doScript(File($.fileName), ScriptLanguage.JAVASCRIPT);
        return;
    }

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

    var LIST_SIZE     = [600, 320];          /* リストの寸法 [幅,高さ] / list size */
    var COLUMN_WIDTHS = [90, 170, 340];      /* 列幅（ページ・アンカー名・テキスト）/ column widths (page, anchor name, text) */

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
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match */
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
        "I": ["#", "#", "#", "#", "#", "#", "#", ".", "."],
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
        "r": ["....", "....", "#.##", "##..", "#...", "#...", "#...", "....", "...."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", "#...#", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#", "##", ".#", ".#", ".#", ".#", ".#", "..", ".."],
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

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

    var LABELS = {
        palette: {
            title: { ja: "テキストアンカー", en: "Text Anchors" }
        },
        column: {
            page: { ja: "ページ", en: "Page" },
            name: { ja: "アンカー名", en: "Anchor Name" },
            text: { ja: "テキスト", en: "Text" }
        },
        location: {
            pasteboard: { ja: "ペーストボード", en: "Pasteboard" },
            overset:    { ja: "オーバーセット", en: "Overset" }
        },
        button: {
            refresh: { ja: "更新", en: "Refresh" },
            show:    { ja: "表示", en: "Show" },
            close:   { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            list: {
                ja: "ダブルクリックすると、そのアンカーのページを表示してテキストを選択します。",
                en: "Double-click to show the page of that anchor and select its text."
            },
            refresh: {
                ja: "アクティブなドキュメントのアンカーを読み直します。アンカーを追加・削除したときや、ドキュメントを切り替えたときに使います。",
                en: "Reads the anchors of the active document again. Use it after adding or removing anchors, or after switching documents."
            },
            show: {
                ja: "リストで選んだアンカーのページを表示し、アンカーのある段落のテキストを選択します。",
                en: "Shows the page of the anchor chosen in the list and selects the text of the paragraph that holds it."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Open a document first." },
            noAnchor:   { ja: "テキストアンカーは見つかりませんでした。", en: "No text anchors were found." },
            closedDocument: {
                ja: "一覧を作ったドキュメントが閉じられています。［更新］で読み直してください。",
                en: "The document of this list has been closed. Click Refresh to read it again."
            },
            removedAnchor: {
                ja: "このアンカーは削除されています。［更新］で読み直してください。",
                en: "This anchor has been removed. Click Refresh to read the list again."
            },
            showFailed: {
                ja: "アンカーを表示できませんでした（オーバーセット、ロック・非表示のレイヤーなど）。",
                en: "Could not show the anchor (overset text, locked or hidden layer, etc.)."
            }
        },
        counter: {
            itemUnit: { ja: "件", en: " items" }
        }
    };

    // =========================================
    // アンカーの読み取り / Reading anchors
    // =========================================

    /**
     * @typedef {Object} AnchorEntry
     * @property {HyperlinkTextDestination} destination - テキストアンカー
     * @property {InsertionPoint} insertionPoint - アンカーの位置
     * @property {string} name - アンカー名
     * @property {string} text - アンカーのある段落のテキスト（表示用に整えたもの）
     * @property {string} pageLabel - ページの表示名
     * @property {Page|null} page - 移動先のページ（ペーストボードは null）
     * @property {Spread|MasterSpread|null} spread - 移動先のスプレッド（オーバーセットは null）
     * @property {string} sortKey - 並べ替えキー
     */

    /**
     * 数値を指定桁の0埋め文字列にする（並べ替えキー用）
     * @param {number} value - 0以上の数値
     * @param {number} digits - 桁数
     * @returns {string} 0埋めした文字列
     */
    function padNumber(value, digits) {
        var numberText = String(Math.max(0, Math.round(value)));
        while (numberText.length < digits) numberText = "0" + numberText;
        return numberText;
    }

    /**
     * 段落のテキストを1行の表示用に整える（改行・タブを空白に、制御文字と U+FEFF・U+FFFC（アンカー付きオブジェクトなどの目印）を除き、長ければ切り詰める）
     * @param {string} rawText - 段落のテキスト
     * @returns {string} 表示用のテキスト
     */
    function formatAnchorText(rawText) {
        var displayText = String(rawText)
            .split(String.fromCharCode(0xFEFF)).join("")
            .split(String.fromCharCode(0xFFFC)).join("")
            .replace(/[\r\n\t]+/g, " ")
            .replace(/[\x00-\x1F]/g, "")
            .replace(/\s+/g, " ")
            .replace(/^\s+|\s+$/g, "");
        if (displayText.length > TEXT_MAX_LENGTH) displayText = displayText.substring(0, TEXT_MAX_LENGTH) + "...";
        return displayText;
    }

    /**
     * アンカーのある段落のテキストを返す
     * @param {InsertionPoint} insertionPoint - アンカーの位置
     * @returns {string} 表示用のテキスト。読めなければ空文字
     */
    function readParagraphText(insertionPoint) {
        try {
            var paragraphs = insertionPoint.paragraphs;
            if (paragraphs.length === 0) return "";
            return formatAnchorText(paragraphs[0].contents);
        } catch (e) {
            return "";
        }
    }

    /**
     * フレームの上にあるスプレッドを探す（ペーストボードの判定用）。
     * アンカー付きオブジェクトの中のテキストは、文字を収めたテキストフレームからたどる
     * @param {PageItem} frame - 対象のフレーム
     * @returns {Spread|MasterSpread|null} スプレッド。オーバーセットの文字に入ったものは null
     */
    function findParentSpread(frame) {
        var ancestor = frame.parent;
        while (ancestor) {
            var ancestorType = ancestor.constructor.name;
            if (ancestorType === "Spread" || ancestorType === "MasterSpread") return ancestor;
            if (ancestorType === "Document" || ancestorType === "Application") return null;
            if (ancestorType === "Character") {
                var textFrames = ancestor.parentTextFrames;
                /* 収まるテキストフレームが無い文字＝オーバーセット / A character without a text frame is overset */
                if (!textFrames || textFrames.length === 0) return null;
                ancestor = textFrames[0];
                continue;
            }
            ancestor = ancestor.parent;
        }
        return null;
    }

    /**
     * スプレッドの表示名を作る（親ページはその名前、ドキュメントのスプレッドは「1-2」のようなページ番号）
     * @param {Spread|MasterSpread} spread - 対象のスプレッド
     * @returns {string} 表示名
     */
    function buildSpreadLabel(spread) {
        if (spread.constructor.name === "MasterSpread") return spread.name;
        var spreadPages = spread.pages;
        if (spreadPages.length === 0) return "";
        var firstPageName = spreadPages[0].name;
        var lastPageName = spreadPages[spreadPages.length - 1].name;
        return (firstPageName === lastPageName) ? firstPageName : firstPageName + "-" + lastPageName;
    }

    /**
     * アンカーの場所（ページの表示名・移動先・並べ替えキーの先頭）を調べる。
     * 並びはドキュメントページ → 親ページ → ペーストボード → オーバーセット。同じ場所の中ではフレームの上 → 左
     * @param {InsertionPoint} insertionPoint - アンカーの位置
     * @returns {{pageLabel: string, page: Page|null, spread: Spread|MasterSpread|null, sortPrefix: string}} 場所
     */
    function describeLocation(insertionPoint) {
        var textFrames = insertionPoint.parentTextFrames;
        /* 収まるテキストフレームが無い＝オーバーセット / No text frame means overset */
        if (!textFrames || textFrames.length === 0) {
            return { pageLabel: getLabel(LABELS.location.overset), page: null, spread: null, sortPrefix: "3000000" + padNumber(0, 14) };
        }
        var frame = textFrames[0];
        var bounds = frame.geometricBounds;
        var positionKey = padNumber(bounds[0] + SORT_COORDINATE_OFFSET, 7) + padNumber(bounds[1] + SORT_COORDINATE_OFFSET, 7);

        var parentPage = frame.parentPage;
        if (parentPage) {
            var pageSpread = parentPage.parent;
            if (pageSpread.constructor.name === "MasterSpread") {
                return { pageLabel: pageSpread.name, page: parentPage, spread: pageSpread, sortPrefix: "1" + padNumber(pageSpread.index, 6) + positionKey };
            }
            return { pageLabel: parentPage.name, page: parentPage, spread: pageSpread, sortPrefix: "0" + padNumber(parentPage.documentOffset, 6) + positionKey };
        }
        var parentSpread = findParentSpread(frame);
        if (parentSpread) {
            return {
                pageLabel: getLabel(LABELS.location.pasteboard) + (uiLang === "ja" ? "（" + buildSpreadLabel(parentSpread) + "）" : " (" + buildSpreadLabel(parentSpread) + ")"),
                page: null,
                spread: parentSpread,
                sortPrefix: "2" + padNumber(parentSpread.index, 6) + positionKey
            };
        }
        return { pageLabel: getLabel(LABELS.location.overset), page: null, spread: null, sortPrefix: "3000000" + padNumber(0, 14) };
    }

    /**
     * ハイパーリンク先を一覧に入れるか判定する（非表示のものは INCLUDE_HIDDEN_ANCHORS で切り替え）
     * @param {HyperlinkTextDestination} destination - 対象のハイパーリンク先
     * @returns {boolean} 入れるなら true
     */
    function isListedAnchor(destination) {
        if (INCLUDE_HIDDEN_ANCHORS) return true;
        try {
            return !destination.hidden;
        } catch (e) {
            return true;
        }
    }

    /**
     * ドキュメントのテキストアンカーを集めてページ順に並べる
     * @param {Document} doc - 対象のドキュメント
     * @returns {AnchorEntry[]} アンカーの一覧
     */
    function collectAnchors(doc) {
        var destinations = doc.hyperlinkTextDestinations.everyItem().getElements();
        var entries = [];
        for (var i = 0; i < destinations.length; i++) {
            var destination = destinations[i];
            if (!isListedAnchor(destination)) continue;
            var insertionPoint;
            try {
                insertionPoint = destination.destinationText;
            } catch (e) {
                continue;
            }
            if (!insertionPoint || !insertionPoint.isValid) continue;
            var location = describeLocation(insertionPoint);
            entries.push({
                destination: destination,
                insertionPoint: insertionPoint,
                name: destination.name,
                text: readParagraphText(insertionPoint),
                pageLabel: location.pageLabel,
                page: location.page,
                spread: location.spread,
                /* 場所 → フレームの上 → 左 → ストーリー内の位置 → 見つかった順 / location, frame top, left, position in the story, then discovery order */
                sortKey: location.sortPrefix + padNumber(insertionPoint.index, 9) + padNumber(i, 6)
            });
        }
        return sortEntries(entries);
    }

    /**
     * 並べ替えキーの順に並べる（比較関数つきの sort() は遅く並びも狂うため、キーの文字列で並べる）
     * @param {AnchorEntry[]} entries - アンカーの一覧
     * @returns {AnchorEntry[]} 並べ替えた一覧
     */
    function sortEntries(entries) {
        var entryByKey = {};
        var sortKeys = [];
        for (var i = 0; i < entries.length; i++) {
            entryByKey[entries[i].sortKey] = entries[i];
            sortKeys.push(entries[i].sortKey);
        }
        sortKeys.sort();
        var sortedEntries = [];
        for (var j = 0; j < sortKeys.length; j++) {
            sortedEntries.push(entryByKey[sortKeys[j]]);
        }
        return sortedEntries;
    }

    // =========================================
    // 表示・操作 / Display & actions
    // =========================================

    /**
     * 数値を3桁区切りの文字列にする
     * @param {number} value - 数値
     * @returns {string} 3桁区切りの文字列
     */
    function formatThousands(value) {
        return String(value).replace(/(\d)(?=(\d\d\d)+$)/g, "$1,");
    }

    /**
     * 件数入りのパレットタイトルを作る
     * @param {number} itemCount - 件数
     * @returns {string} パレットタイトル
     */
    function buildPaletteTitle(itemCount) {
        var countText = formatThousands(itemCount) + getLabel(LABELS.counter.itemUnit);
        return getLabel(LABELS.palette.title)
            + (uiLang === "ja" ? "（" + countText + "）" : " (" + countText + ")");
    }

    /**
     * アンカーのある段落のテキスト（末尾の改行を除く）を返す
     * @param {InsertionPoint} insertionPoint - アンカーの位置
     * @returns {Text|InsertionPoint} 段落のテキスト。空の段落や読めないときはアンカーの位置
     */
    function getParagraphRange(insertionPoint) {
        try {
            var paragraphs = insertionPoint.paragraphs;
            if (paragraphs.length === 0) return insertionPoint;
            var paragraphCharacters = paragraphs[0].characters;
            var lastIndex = paragraphCharacters.length - 1;
            /* 段落末の改行は選択に含めない / Leave the paragraph return out of the selection */
            if (lastIndex >= 0 && paragraphCharacters[lastIndex].contents === "\r") lastIndex--;
            if (lastIndex < 0) return insertionPoint;
            return paragraphCharacters.itemByRange(paragraphCharacters[0], paragraphCharacters[lastIndex]);
        } catch (e) {
            return insertionPoint;
        }
    }

    /**
     * 2つのドキュメントの参照が同じドキュメントを指すか判定する（InDesign の参照は === で比べられないので、指定子で比べる）
     * @param {Document} firstDocument - 1つめ
     * @param {Document} secondDocument - 2つめ
     * @returns {boolean} 同じなら true
     */
    function isSameDocument(firstDocument, secondDocument) {
        return firstDocument.toSpecifier() === secondDocument.toSpecifier();
    }

    /**
     * アンカーのページ（またはスプレッド）を表示し、アンカーのある段落のテキストを選択する。
     * 一覧を作ったドキュメントがアクティブでなければ、先にアクティブにする
     * @param {Document} doc - 一覧を作ったドキュメント
     * @param {AnchorEntry} entry - 対象のアンカー
     * @returns {void}
     */
    function showAnchor(doc, entry) {
        if (!doc.isValid) {
            alert(getLabel(LABELS.alert.closedDocument));
            return;
        }
        if (!entry.destination.isValid || !entry.insertionPoint.isValid) {
            alert(getLabel(LABELS.alert.removedAnchor));
            return;
        }
        try {
            if (!isSameDocument(app.activeDocument, doc)) app.activeDocument = doc;
            if (entry.page) {
                app.activeWindow.activePage = entry.page;
            } else if (entry.spread) {
                app.activeWindow.activeSpread = entry.spread;
            } else {
                throw new Error(getLabel(LABELS.location.overset));
            }
        } catch (e) {
            alert(getLabel(LABELS.alert.showFailed) + "\n" + e.message);
            return;
        }
        /* ［ハイパーリンク］パネルの「ハイパーリンク先へ移動」と同じ操作で、アンカーの位置までスクロールする（失敗してもページは表示済み）。
           選択をアンカーの位置に変えることがあるので、段落の選択より先に呼ぶ
           / Same as "Go to Destination" in the Hyperlinks panel; scrolls to the anchor (the page is already shown if this fails).
             It may move the selection to the anchor, so call it before selecting the paragraph */
        try {
            entry.destination.showDestination();
        } catch (e) {}
        try {
            app.select(getParagraphRange(entry.insertionPoint));
        } catch (e) {
            alert(getLabel(LABELS.alert.showFailed) + "\n" + e.message);
        }
    }

    /**
     * リストの中身を一覧で入れ直す。選んでいたアンカーと同じ名前の行があれば選び直す
     * @param {ListBox} anchorList - リスト
     * @param {AnchorEntry[]} entries - アンカーの一覧
     * @param {string|null} keepName - 選び直すアンカー名（無ければ先頭）
     * @returns {void}
     */
    function fillAnchorList(anchorList, entries, keepName) {
        anchorList.removeAll();
        var selectIndex = 0;
        for (var i = 0; i < entries.length; i++) {
            var listItem = anchorList.add("item", entries[i].pageLabel);
            listItem.subItems[0].text = entries[i].name;
            listItem.subItems[1].text = entries[i].text;
            if (keepName !== null && entries[i].name === keepName) selectIndex = i;
        }
        if (entries.length > 0) anchorList.selection = selectIndex;
    }

    /**
     * 一覧のパレットを組み立てる
     * @param {Document} doc - 一覧を作ったドキュメント
     * @param {AnchorEntry[]} entries - アンカーの一覧
     * @returns {Window} パレット
     */
    function buildPalette(doc, entries) {
        var listedDocument = doc;
        var listedEntries = entries;

        var anchorPalette = new Window("palette", buildPaletteTitle(entries.length)); /* バージョンは左下に出す / the version is shown at the bottom left */
        setupWindow(anchorPalette);

        var anchorList = anchorPalette.add("listbox", [0, 0, LIST_SIZE[0], LIST_SIZE[1]], "", {
            multiselect: false,
            numberOfColumns: 3,
            showHeaders: true,
            columnTitles: [
                getLabel(LABELS.column.page),
                getLabel(LABELS.column.name),
                getLabel(LABELS.column.text)
            ],
            columnWidths: COLUMN_WIDTHS
        });
        anchorList.preferredSize = LIST_SIZE;
        anchorList.helpTip = getLabel(LABELS.tooltip.list);
        fillAnchorList(anchorList, listedEntries, null);

        var buttonRow = addButtonRow(anchorPalette);
        var btnRefresh = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.refresh));
        btnRefresh.helpTip = getLabel(LABELS.tooltip.refresh);
        var btnClose = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.close));
        var btnShow = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.show));
        btnShow.helpTip = getLabel(LABELS.tooltip.show);

        /**
         * 選んでいる行のアンカーを表示する（パレットは開いたまま）
         * @returns {void}
         */
        function showSelectedAnchor() {
            if (anchorList.selection === null) return;
            showAnchor(listedDocument, listedEntries[anchorList.selection.index]);
        }

        anchorList.onChange = function () {
            btnShow.enabled = (anchorList.selection !== null);
        };
        anchorList.onDoubleClick = showSelectedAnchor;
        btnShow.onClick = showSelectedAnchor;
        btnRefresh.onClick = function () {
            if (app.documents.length === 0) {
                alert(getLabel(LABELS.alert.noDocument));
                return;
            }
            var keepName = (anchorList.selection !== null && listedDocument.isValid && isSameDocument(listedDocument, app.activeDocument))
                ? listedEntries[anchorList.selection.index].name
                : null;
            listedDocument = app.activeDocument;
            listedEntries = collectAnchors(listedDocument);
            fillAnchorList(anchorList, listedEntries, keepName);
            anchorPalette.text = buildPaletteTitle(listedEntries.length);
            btnShow.enabled = (anchorList.selection !== null);
        };
        btnClose.onClick = function () {
            anchorPalette.close();
        };

        /* パレットは Esc で閉じないので、［閉じる］と同じにする / palettes do not close on Esc, so treat it as Close */
        anchorPalette.addEventListener("keydown", function (event) {
            if (event.keyName !== "Escape") return;
            event.preventDefault();
            anchorPalette.close();
        });

        /* 閉じるときは参照を外すだけ（DOM には触れない）/ On close, only drop the reference (no DOM access) */
        anchorPalette.onClose = function () {
            $.global[PALETTE_GLOBAL_KEY] = null;
        };

        btnShow.enabled = (anchorList.selection !== null);
        addVersionFooter(anchorPalette);
        return anchorPalette;
    }

    // =========================================
    // 実行 / Main
    // =========================================

    /**
     * パレットを開く（前回のパレットが開いていれば閉じて作り直す）
     * @returns {void}
     */
    function main() {
        if ($.global[PALETTE_GLOBAL_KEY]) $.global[PALETTE_GLOBAL_KEY].close();

        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var doc = app.activeDocument;
        var anchorEntries = collectAnchors(doc);
        if (anchorEntries.length === 0) {
            alert(getLabel(LABELS.alert.noAnchor));
            return;
        }
        var anchorPalette = buildPalette(doc, anchorEntries);
        /* パレットは show() から即座に戻るので、参照を $.global に残す / palettes return from show() at once, so keep a reference */
        $.global[PALETTE_GLOBAL_KEY] = anchorPalette;
        anchorPalette.show();
    }

    main();

}());
