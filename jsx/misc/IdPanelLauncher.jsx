#target indesign
#targetengine "IdPanelLauncher"

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdPanelLauncher
Name=パネルを検索して表示
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.1.1
Release-Date=2026-11-05
Description-BEGIN
［ウィンドウ］メニューのパネルを、サブメニューの中まで展開して1つのリストに並べます。
検索欄で絞り込み、選んだパネルを表示します。

Lists every panel in the Window menu, including those inside submenus, in a single list.
Filter it with the search field and show the panels you pick.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

［ウィンドウ］メニューのパネルを、サブメニューの中まで展開して1つのリストに並べます。
検索欄で絞り込み、選んだパネルを表示します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdPanelLauncher.md

### Overview

Lists every panel in the Window menu, including those inside submenus, in a single list.
Filter it with the search field and show the panels you pick.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdPanelLauncher.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdPanelLauncher";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-06";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "パネルを検索して表示"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdPanelLauncher.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdPanelLauncher.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/misc/IdPanelLauncher.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 基本設定 / Settings
    // =========================================

    /* ［ウィンドウ］メニューの区切り線で分けた区画のうち、パネルが並ぶ区画（0 始まり）。
       0 はアレンジ・ワークスペース、2 はアプリケーションフレームやドキュメントのウィンドウ
       / Section of the Window menu (split by separators, 0-based) that holds the panels.
       0 is Arrange / Workspace, 2 and later are Application Frame and the document windows */
    var PANEL_SECTION_INDEX = 1;

    /* サブメニューの階層の区切り（分類の列）/ Separator between submenu levels (category column) */
    var CATEGORY_LEVEL_SEPARATOR = " > ";

    /* 表示中のパネルの印 / Mark for panels that are open */
    var OPEN_MARK = "✓";

    /* 並び順（ラジオボタンの順）。menu＝デフォルト（メニューの並び）、kana＝あいうえお順
       / Sort orders (radio button order): menu = default (menu order), kana = alphabetical */
    var SORT_KEYS = ["menu", "kana"];
    var DEFAULT_SORT_KEY = "menu";

    /* パネルに付ける印（リストの印の列と、印を付けるボタンの文字）。fav＝よく使う、rare＝あまり使わない（どちらか一方）
       / Marks for panels (shown in the Mark column and on the mark buttons): fav = frequently used, rare = rarely used (one or the other) */
    var MARK_GLYPHS = {
        fav:  "\u2665",  /* ♥ */
        rare: "\u2662"   /* ♢ */
    };

    /* 絞り込みのチェックボックス（並びの順）。fav＝よく使う（♥）だけ、hideRare＝あまり使わない（♢）を除く。両方オフならすべて
       / Filter checkboxes (in order): fav = only ♥ panels, hideRare = hide ♢ panels. Both off shows everything */
    var FILTER_KEYS = ["fav", "hideRare"];

    /* 印と絞り込みを残す設定ファイル（Folder.userData の中）/ File that keeps the marks and the filter (in Folder.userData) */
    var SETTINGS_FILE_NAME = "IdPanelLauncher-settings.txt";

    /* まとめて表示するパネルのセット（ボタンの順）。panels はメニューに出るパネル名（日本語版・英語版）で、上から順に表示する
       / Panel sets shown together (button order). panels are the menu names (Japanese / English), shown top to bottom */
    var PANEL_SETS = [
        {
            label: { ja: "表", en: "Tables" },
            panels: [
                { ja: "セルスタイル", en: "Cell Styles" },
                { ja: "表スタイル", en: "Table Styles" },
                { ja: "表", en: "Table" }
            ]
        },
        {
            label: { ja: "テキスト", en: "Text" },
            panels: [
                { ja: "文字", en: "Character" },
                { ja: "段落", en: "Paragraph" },
                { ja: "文字スタイル", en: "Character Styles" },
                { ja: "段落スタイル", en: "Paragraph Styles" },
                { ja: "字形", en: "Glyphs" },
                { ja: "テキストの回り込み", en: "Text Wrap" }
            ]
        },
        {
            label: { ja: "ブック", en: "Book" },
            panels: [
                { ja: "相互参照", en: "Cross-References" },
                { ja: "ハイパーリンク", en: "Hyperlinks" }
            ]
        }
    ];

    /* あいうえお順と検索で使う、パネル名・分類の漢字の読み（長い語から置き換える）
       / Readings of the kanji in Japanese panel and category names, used for the alphabetical order */
    var KANJI_READINGS = [
        ["相互参照", "ソウゴサンショウ"], ["編集関連", "ヘンシュウカンレン"],
        ["整列", "セイレツ"], ["変形", "ヘンケイ"], ["段落", "ダンラク"], ["文字", "モジ"],
        ["効果", "コウカ"], ["透明", "トウメイ"], ["分割", "ブンカツ"], ["統合", "トウゴウ"],
        ["分版", "ブンパン"], ["索引", "サクイン"], ["字形", "ジケイ"], ["条件", "ジョウケン"],
        ["情報", "ジョウホウ"], ["数式", "スウシキ"], ["生成", "セイセイ"], ["注釈", "チュウシャク"],
        ["変更", "ヘンコウ"], ["履歴", "リレキ"], ["出力", "シュツリョク"], ["書式", "ショシキ"],
        ["回り込み", "マワリコミ"], ["表", "ヒョウ"], ["線", "セン"]
    ];

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

    var LIST_SIZE     = [440, 460];        /* リストの寸法 [幅,高さ] / list size */
    var COLUMN_WIDTHS = [200, 30, 160, 40]; /* 列幅（名前・印・分類・表示中）/ column widths (name, mark, category, open) */
    var ROW_LABEL_WIDTH = 66;              /* 行ラベル（検索・絞り込み・並び順・セット）の幅 / width of the row labels (search, filter, sort, sets) */
    var MARK_BUTTON_WIDTH = 40;            /* ♥・♢ボタンの幅 / width of the ♥ and ♢ buttons */

    /* 検索欄のクリアボタン（自前描画、IdScriptLauncher と同じ）/ Clear button next to the search field (custom drawn, same as IdScriptLauncher) */
    var CLEAR_BUTTON_SIZE  = 20;           /* クリアボタンの一辺 / clear button size */
    var CLEAR_CIRCLE_INSET = 2;            /* 円と外周の間隔 / inset of the circle */
    var CLEAR_GLYPH_INSET  = 6;            /* ×と外周の間隔 / inset of the × glyph */
    var CLEAR_STROKE_WIDTH = 1.5;          /* 円と×の線幅 / stroke width */

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
        "a": ["....", "....", "....", ".###", "#..#", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "#...", "###.", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", "....", ".###", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", "...#", ".###", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", "....", ".##.", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", ".#.", "###", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", "....", ".###", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "#...", "###.", "#..#", "#..#", "#..#", "....", "...."],
        "i": [".", "#", ".", "#", "#", "#", "#", ".", "."],
        "j": ["..", ".#", "..", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#...", "#..#", "#.#.", "###.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "....", "###.", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", "....", ".##.", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "....", "###.", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", "....", ".###", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["....", "....", "....", "#.##", "##..", "#...", "#...", "....", "...."],
        "s": ["....", "....", "....", ".###", "##..", "..##", "###.", "....", "...."],
        "t": ["...", ".#.", ".#.", "###", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "....", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", ".....", "#...#", "#...#", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", ".....", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "....", "#..#", ".##.", ".##.", "#..#", "....", "...."],
        "y": ["....", "....", "....", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "....", "####", "..#.", ".#..", "####", "....", "...."],
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

    /**
     * クリアボタンの丸と×を描く
     * @param {Button} button - 対象ボタン
     * @returns {void}
     */
    function drawClearButton(button) {
        var graphics = button.graphics;
        var size = button.size[0];
        var useLightUI = !isDarkUI();

        /* 主役ではないので、通常時も少し薄めに描く。押している間は濃くして手応えを出す / Slightly muted, darker while pressed */
        var glyphColor = useLightUI ? [0.45, 0.45, 0.45, 1] : [0.72, 0.72, 0.72, 1];
        if (button.pressed) glyphColor = useLightUI ? [0.20, 0.20, 0.20, 1] : [0.95, 0.95, 0.95, 1];

        /* 検索欄が空のときは無効にしてあるので、さらに淡くする / Dim further while disabled */
        if (button.enabled === false) glyphColor = useLightUI ? [0.78, 0.78, 0.78, 1] : [0.42, 0.42, 0.42, 1];

        /* 前回の描画を消すため、まず親と同じ色で塗りつぶす。InDesignでは背景ブラシが取れないので何もしない / No-op in InDesign */
        try {
            graphics.newPath();
            graphics.rectPath(0, 0, size, size);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) {}

        var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, CLEAR_STROKE_WIDTH);
        var circleSize = size - CLEAR_CIRCLE_INSET * 2;
        graphics.newPath();
        graphics.ellipsePath(CLEAR_CIRCLE_INSET, CLEAR_CIRCLE_INSET, circleSize, circleSize);
        graphics.strokePath(pen);

        var glyphStart = CLEAR_GLYPH_INSET;
        var glyphEnd = size - CLEAR_GLYPH_INSET;
        graphics.newPath();
        graphics.moveTo(glyphStart, glyphStart);
        graphics.lineTo(glyphEnd, glyphEnd);
        graphics.strokePath(pen);
        graphics.newPath();
        graphics.moveTo(glyphEnd, glyphStart);
        graphics.lineTo(glyphStart, glyphEnd);
        graphics.strokePath(pen);
    }

    /**
     * 検索欄を消すクリアボタンを作る
     * @param {Group} parent - 追加先のグループ
     * @returns {Button} 丸に×を自前描画したボタン
     */
    function addClearButton(parent) {
        var button = parent.add("button", undefined, "");
        button.helpTip = getLabel(LABELS.tooltip.clearSearch);
        button.alignment = ["right", "center"];
        button.preferredSize = [CLEAR_BUTTON_SIZE, CLEAR_BUTTON_SIZE];
        button.minimumSize = [CLEAR_BUTTON_SIZE, CLEAR_BUTTON_SIZE];
        button.maximumSize = [CLEAR_BUTTON_SIZE, CLEAR_BUTTON_SIZE];
        button.pressed = false;

        /* 実際の状態は呼び出し側が入力欄に合わせる / The caller syncs this with the field */
        button.enabled = false;

        button.onDraw = function () {
            drawClearButton(this);
        };
        button.addEventListener("mousedown", function () {
            this.pressed = true;
            this.notify("onDraw");
        });
        button.addEventListener("mouseup", function () {
            this.pressed = false;
            this.notify("onDraw");
        });
        /* ボタンの外でマウスを離したときも押下表示を戻す / Reset the pressed look on mouseout */
        button.addEventListener("mouseout", function () {
            if (!this.pressed) return;
            this.pressed = false;
            this.notify("onDraw");
        });

        return button;
    }

    var LABELS = {
        dialog: {
            title: { ja: "パネル", en: "Panels" }
        },
        column: {
            name:     { ja: "名前", en: "Name" },
            mark:     { ja: "印", en: "Mark" },
            category: { ja: "分類", en: "Category" },
            open:     { ja: "表示中", en: "Open" }
        },
        fieldLabel: {
            search:    { ja: "検索", en: "Search" },
            filter:    { ja: "絞り込み", en: "Show" },
            sortOrder: { ja: "並び順", en: "Sort by" },
            panelSet:  { ja: "セット", en: "Sets" }
        },
        filter: {
            fav:      { ja: "\u2665 よく使う", en: "\u2665 Frequent" },
            hideRare: { ja: "\u2662 を除く", en: "Hide \u2662" }
        },
        sort: {
            menu: { ja: "デフォルト", en: "Default" },
            kana: { ja: "あいうえお順", en: "Alphabetical" }
        },
        button: {
            show:   { ja: "表示", en: "Show" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            search: {
                ja: "名前と分類（サブメニュー名）で絞り込みます。空白で区切ると、すべての語を含むものに絞ります。ひらがなはカタカナとして探し、漢字の名前は読みでも探します。↓でリストへ移ります。",
                en: "Filters by name and category (submenu name). Separate words with spaces to match all of them. Press Down to move to the list."
            },
            filter: {
                ja: "♥ よく使う：♥ を付けたパネルだけを出します。♢ を除く：♢（あまり使わない）を付けたパネルを隠します。両方オフならすべて出します。",
                en: "\u2665 Frequent lists only the panels marked \u2665. Hide \u2662 leaves out the panels marked \u2662 (rarely used). With both off, every panel is listed."
            },
            markFav: {
                ja: "選んだパネルに ♥（よく使う）を付けます。すべてに付いていれば外します。",
                en: "Marks the selected panels with \u2665 (frequently used). Removes the mark if they all have it."
            },
            markRare: {
                ja: "選んだパネルに ♢（あまり使わない）を付けます。すべてに付いていれば外します。",
                en: "Marks the selected panels with \u2662 (rarely used). Removes the mark if they all have it."
            },
            sortOrder: {
                ja: "リストの並び順を切り替えます。デフォルトは［ウィンドウ］メニューの並び（サブメニューの中身はサブメニューの位置）、あいうえお順はパネル名の読みの順です。",
                en: "Changes the order of the list. Default follows the Window menu (submenu items stay where the submenu is); Alphabetical sorts by panel name."
            },
            clearSearch: { ja: "検索をクリア", en: "Clear search" },
            panelSet: { ja: "%1 をまとめて表示します。", en: "Shows %1 together." },
            list: {
                ja: "ダブルクリックか Enter で表示します。⌘（Ctrl）／Shift＋クリックで複数選べます。左下の ♥・♢ で印を付けます。",
                en: "Double-click or press Enter to show. Cmd (Ctrl) / Shift-click to pick more than one. Use \u2665 / \u2662 at the bottom left to mark them."
            }
        },
        counter: {
            itemUnit: { ja: "件", en: " panels" }
        },
        alert: {
            notFound:   { ja: "［ウィンドウ］メニューにパネルが見つかりませんでした。", en: "No panels were found in the Window menu." },
            showFailed: { ja: "表示できなかったパネル", en: "Could not show" }
        }
    };

    // =========================================
    // データ / Data
    // =========================================

    /**
     * @typedef {Object} PanelEntry
     * @property {string} name - メニューに出る名前
     * @property {string} category - 親のサブメニュー名（「カラー」「書式と表」など。直下の項目は空文字）
     * @property {number} actionId - メニューアクションの ID
     * @property {boolean} isOpen - 表示中（メニューにチェックが付いている）なら true。ダイアログを出す前に読む
     * @property {number} menuIndex - メニューの並び順（0 始まり。並び順の同順位もこれで決める）
     * @property {string} mark - 印（"fav"＝よく使う、"rare"＝あまり使わない、無ければ空文字）
     */

    /**
     * メニュー要素の種類を返す（menuElements の要素は MenuElement のままなので getElements() で解決してから使う）
     * @param {Object} menuElement - メニュー要素
     * @returns {string} "MenuItem" / "Submenu" / "MenuSeparator" など
     */
    function getMenuElementType(menuElement) {
        return menuElement.constructor.name;
    }

    /**
     * メニュー項目を1件の PanelEntry にする
     * @param {MenuItem} menuItem - メニュー項目
     * @param {string} categoryName - 親のサブメニュー名
     * @returns {PanelEntry} パネルの情報
     */
    function buildPanelEntry(menuItem, categoryName) {
        var menuAction = menuItem.associatedMenuAction;
        return {
            name: menuItem.name,
            category: categoryName,
            actionId: menuAction.id,
            isOpen: menuAction.checked,
            menuIndex: -1,
            mark: ""
        };
    }

    /**
     * メニュー要素を順に読み、項目は entries に足し、サブメニューは中まで展開する
     * @param {Object[]} menuElements - getElements() で解決したメニュー要素
     * @param {string} categoryName - 親のサブメニュー名（直下なら空文字）
     * @param {PanelEntry[]} entries - 足す先
     * @returns {void}
     */
    function appendMenuEntries(menuElements, categoryName, entries) {
        for (var i = 0; i < menuElements.length; i++) {
            var elementType = getMenuElementType(menuElements[i]);
            if (elementType === "MenuItem") {
                entries.push(buildPanelEntry(menuElements[i], categoryName));
            } else if (elementType === "Submenu") {
                var childCategory = (categoryName === "")
                    ? menuElements[i].name
                    : categoryName + CATEGORY_LEVEL_SEPARATOR + menuElements[i].name;
                appendMenuEntries(menuElements[i].menuElements.everyItem().getElements(), childCategory, entries);
            }
        }
    }

    /**
     * ［ウィンドウ］メニュー直下の要素のうち、パネルの区画（PANEL_SECTION_INDEX）の要素を返す。
     * 区切り線が足りず区画が見つからないときは、直下の要素をすべて返す
     * @param {Object[]} topElements - ［ウィンドウ］メニュー直下の要素
     * @returns {Object[]} パネルの区画の要素
     */
    function pickPanelSection(topElements) {
        var sectionIndex = 0;
        var sectionElements = [];
        for (var i = 0; i < topElements.length; i++) {
            if (getMenuElementType(topElements[i]) === "MenuSeparator") {
                sectionIndex++;
                continue;
            }
            if (sectionIndex === PANEL_SECTION_INDEX) sectionElements.push(topElements[i]);
        }
        return (sectionElements.length > 0) ? sectionElements : topElements;
    }

    /**
     * ［ウィンドウ］メニューのパネルを、サブメニューの中まで展開してメニューの並び順で集める
     * @returns {PanelEntry[]} パネルの一覧
     */
    function collectPanelEntries() {
        var windowMenu = app.menus.item("$ID/Main").submenus.item("$ID/&Window");
        var entries = [];
        appendMenuEntries(pickPanelSection(windowMenu.menuElements.everyItem().getElements()), "", entries);
        for (var i = 0; i < entries.length; i++) {
            entries[i].menuIndex = i;
        }
        return entries;
    }

    /**
     * @typedef {Object} LauncherSettings
     * @property {Object} marks - パネルの鍵（getPanelKey）から印（"fav"／"rare"）への対応
     * @property {Object} filters - 絞り込み（FILTER_KEYS の鍵ごとにオンなら true）
     */

    /**
     * 設定ファイルを返す
     * @returns {File} 設定ファイル
     */
    function getSettingsFile() {
        return File(Folder.userData.fsName + "/" + SETTINGS_FILE_NAME);
    }

    /**
     * 印を残すときのパネルの鍵（分類とパネル名をタブでつなぐ。メニューアクションの ID はエクステンションのパネルで変わりうるので使わない）
     * @param {PanelEntry} panelEntry - 対象のパネル
     * @returns {string} 鍵
     */
    function getPanelKey(panelEntry) {
        return panelEntry.category + "\t" + panelEntry.name;
    }

    /**
     * 設定ファイルを読む。1行に「filter=オンの絞り込み（カンマ区切り）」か「fav=鍵」「rare=鍵」
     * @returns {LauncherSettings} 読んだ設定（読めなければ印なし・絞り込みなし）
     */
    function loadSettings() {
        var settings = { marks: {}, filters: { fav: false, hideRare: false } };
        var settingsFile = getSettingsFile();
        if (!settingsFile.exists) return settings;
        var rawText = "";
        try {
            settingsFile.encoding = "UTF-8";
            if (settingsFile.open("r")) rawText = settingsFile.read();
        } catch (e) {
        } finally {
            try { settingsFile.close(); } catch (e2) {}
        }
        var lines = String(rawText).split(/\r\n|\r|\n/);
        for (var i = 0; i < lines.length; i++) {
            /* 鍵（パネル名）にも = が入りうるので最初の = だけで区切る / Split on the first = only */
            var separatorIndex = lines[i].indexOf("=");
            if (separatorIndex <= 0) continue;
            var lineKey = lines[i].substring(0, separatorIndex);
            var lineValue = lines[i].substring(separatorIndex + 1);
            if (lineKey === "filter") readFilterValue(lineValue, settings.filters);
            else if (MARK_GLYPHS.hasOwnProperty(lineKey)) settings.marks[lineValue] = lineKey;
        }
        return settings;
    }

    /**
     * 設定ファイルに書き出す（今のメニューに無いパネルの印も残す）。書き出せなくても操作は続ける
     * @param {LauncherSettings} settings - 書き出す設定
     * @returns {void}
     */
    function saveSettings(settings) {
        var onFilterKeys = [];
        for (var i = 0; i < FILTER_KEYS.length; i++) {
            if (settings.filters[FILTER_KEYS[i]]) onFilterKeys.push(FILTER_KEYS[i]);
        }
        var lines = ["filter=" + onFilterKeys.join(",")];
        for (var panelKey in settings.marks) {
            if (settings.marks.hasOwnProperty(panelKey)) lines.push(settings.marks[panelKey] + "=" + panelKey);
        }
        var settingsFile = getSettingsFile();
        try {
            settingsFile.encoding = "UTF-8";
            settingsFile.lineFeed = "Unix";
            if (settingsFile.open("w")) settingsFile.write(lines.join("\n"));
        } catch (e) {
        } finally {
            try { settingsFile.close(); } catch (e2) {}
        }
    }

    /**
     * 「filter=」の値（カンマ区切りの鍵。v1.1.0 の1つだけの値や "all" も同じ書き方で読める）を絞り込みに写す
     * @param {string} filterValue - 値
     * @param {Object} filters - 写す先（FILTER_KEYS の鍵ごとの真偽）
     * @returns {void}
     */
    function readFilterValue(filterValue, filters) {
        var onKeys = "," + filterValue + ",";
        for (var i = 0; i < FILTER_KEYS.length; i++) {
            filters[FILTER_KEYS[i]] = (onKeys.indexOf("," + FILTER_KEYS[i] + ",") >= 0);
        }
    }

    /**
     * 読んだ印をパネルに写す
     * @param {PanelEntry[]} entries - パネルの一覧
     * @param {Object} marks - パネルの鍵から印への対応
     * @returns {void}
     */
    function applySavedMarks(entries, marks) {
        for (var i = 0; i < entries.length; i++) {
            var panelKey = getPanelKey(entries[i]);
            entries[i].mark = marks.hasOwnProperty(panelKey) ? marks[panelKey] : "";
        }
    }

    /**
     * パネルの印を付け替え、設定の対応にも写す（空文字なら外す）
     * @param {PanelEntry} panelEntry - 対象のパネル
     * @param {string} markKey - "fav"／"rare"／""
     * @param {Object} marks - パネルの鍵から印への対応
     * @returns {void}
     */
    function setPanelMark(panelEntry, markKey, marks) {
        panelEntry.mark = markKey;
        if (markKey === "") delete marks[getPanelKey(panelEntry)];
        else marks[getPanelKey(panelEntry)] = markKey;
    }

    /**
     * パネルがすべて同じ印か
     * @param {PanelEntry[]} entries - パネル
     * @param {string} markKey - 印
     * @returns {boolean} 1件以上あり、すべてその印なら true
     */
    function allHaveMark(entries, markKey) {
        if (entries.length === 0) return false;
        for (var i = 0; i < entries.length; i++) {
            if (entries[i].mark !== markKey) return false;
        }
        return true;
    }

    /**
     * 絞り込み（オンのものすべて）に合うパネルだけを返す
     * @param {PanelEntry[]} entries - パネルの一覧
     * @param {Object} filters - FILTER_KEYS の鍵ごとの真偽
     * @returns {PanelEntry[]} 条件に合うパネル
     */
    function filterByMark(entries, filters) {
        var matched = [];
        for (var i = 0; i < entries.length; i++) {
            if (filters.fav && entries[i].mark !== "fav") continue;
            if (filters.hideRare && entries[i].mark === "rare") continue;
            matched.push(entries[i]);
        }
        return matched;
    }

    /**
     * セットのパネルを、セットに書いた順で探す（名前が日本語版か英語版のどちらかと同じもの）
     * @param {PanelEntry[]} entries - すべてのパネル
     * @param {Object} panelSet - PANEL_SETS の1件
     * @returns {{entries: PanelEntry[], missingNames: string[]}} 見つかったパネルと、見つからなかったパネルの名前
     */
    function findPanelSetEntries(entries, panelSet) {
        var found = { entries: [], missingNames: [] };
        for (var i = 0; i < panelSet.panels.length; i++) {
            var setPanel = panelSet.panels[i];
            var matchedEntry = null;
            for (var j = 0; j < entries.length && !matchedEntry; j++) {
                if (entries[j].name === setPanel.ja || entries[j].name === setPanel.en) matchedEntry = entries[j];
            }
            if (matchedEntry) found.entries.push(matchedEntry);
            else found.missingNames.push(getLabel(setPanel));
        }
        return found;
    }

    /**
     * セットのパネル名を、UI の言語で区切ってつなぐ（ツールチップ用）
     * @param {Object} panelSet - PANEL_SETS の1件
     * @returns {string} 「セルスタイル・表スタイル・表」のような文字列
     */
    function joinPanelSetNames(panelSet) {
        var panelNames = [];
        for (var i = 0; i < panelSet.panels.length; i++) {
            panelNames.push(getLabel(panelSet.panels[i]));
        }
        return panelNames.join(uiLang === "ja" ? "・" : ", ");
    }

    /**
     * 検索用に文字列をそろえる（小文字にし、ひらがなをカタカナにする）
     * @param {string} sourceText - 元の文字列
     * @returns {string} そろえた文字列
     */
    function normalizeForSearch(sourceText) {
        return String(sourceText).toLowerCase().replace(/[ぁ-ゖ]/g, function (hiragana) {
            return String.fromCharCode(hiragana.charCodeAt(0) + 0x60);
        });
    }

    /**
     * 検索欄の文字列を語に分ける（空白は半角・全角とも区切り）
     * @param {string} searchText - 検索欄の文字列
     * @returns {string[]} そろえた語の一覧（空欄なら空配列）
     */
    function splitSearchWords(searchText) {
        var trimmedText = normalizeForSearch(searchText).replace(/^[\s　]+|[\s　]+$/g, "");
        return (trimmedText === "") ? [] : trimmedText.split(/[\s　]+/);
    }

    /**
     * 名前と分類（または漢字の読み）に、すべての語が含まれるパネルだけを返す
     * @param {PanelEntry[]} entries - パネルの一覧
     * @param {string} searchText - 検索欄の文字列
     * @returns {PanelEntry[]} 条件に合うパネル
     */
    function filterPanelEntries(entries, searchText) {
        var searchWords = splitSearchWords(searchText);
        if (searchWords.length === 0) return entries.slice(0);
        var matched = [];
        for (var i = 0; i < entries.length; i++) {
            /* 漢字の名前は読みでも探す（「しょしき」で［書式］）/ Kanji names also match their reading */
            var sourceText = entries[i].name + "\n" + entries[i].category;
            var haystack = normalizeForSearch(sourceText) + "\n" + normalizeForSearch(convertKanjiToReading(sourceText));
            var isMatch = true;
            for (var j = 0; j < searchWords.length; j++) {
                if (haystack.indexOf(searchWords[j]) < 0) {
                    isMatch = false;
                    break;
                }
            }
            if (isMatch) matched.push(entries[i]);
        }
        return matched;
    }

    /* 並べ替えキーの区切り（どの文字より前に並ぶ）/ Sort key separator (sorts before any character) */
    var SORT_KEY_SEPARATOR = "\u0001";

    /**
     * メニューの並び順を桁そろえした文字列にする（文字列の並べ替えで数値順になるように）
     * @param {number} menuIndex - メニューの並び順
     * @returns {string} 5桁にゼロ埋めした並び順
     */
    function padMenuIndex(menuIndex) {
        return ("00000" + menuIndex).slice(-5);
    }

    /**
     * 文字列の漢字を KANJI_READINGS の読み（カタカナ）に置き換える
     * @param {string} sourceText - 元の文字列
     * @returns {string} 漢字を読みにした文字列
     */
    function convertKanjiToReading(sourceText) {
        var readingText = String(sourceText);
        for (var i = 0; i < KANJI_READINGS.length; i++) {
            readingText = readingText.split(KANJI_READINGS[i][0]).join(KANJI_READINGS[i][1]);
        }
        return readingText;
    }

    /**
     * あいうえお順の並べ替え用に、名前を読みのカタカナにそろえる。
     * 漢字は KANJI_READINGS で読みに置き換え、長音（ー）と中黒（・）は読みの順に関係しないので除く
     * @param {string} panelName - パネル名
     * @returns {string} 並べ替え用の読み（英字は小文字）
     */
    function buildReadingKey(panelName) {
        return normalizeForSearch(convertKanjiToReading(panelName)).replace(/[ー・]/g, "");
    }

    /**
     * 並べ替え用の文字列キーを作る（同順位はメニューの並び順で決める）
     * @param {PanelEntry} panelEntry - 対象のパネル
     * @param {string} sortKey - SORT_KEYS のどれか
     * @returns {string} 並べ替えキー
     */
    function buildSortKey(panelEntry, sortKey) {
        var orderKey = padMenuIndex(panelEntry.menuIndex);
        if (sortKey === "kana") return buildReadingKey(panelEntry.name) + SORT_KEY_SEPARATOR + orderKey;
        return orderKey;
    }

    /**
     * 指定した並び順に並べ替えた新しい配列を返す。
     * ExtendScript の比較関数つき sort() は遅く並びも狂うことがあるので、文字列キーを組み込みの sort() で並べる
     * @param {PanelEntry[]} entries - パネルの一覧（並べ替えない）
     * @param {string} sortKey - SORT_KEYS のどれか
     * @returns {PanelEntry[]} 並べ替えた一覧
     */
    function sortPanelEntries(entries, sortKey) {
        var sortKeys = [];
        for (var i = 0; i < entries.length; i++) {
            sortKeys.push(buildSortKey(entries[i], sortKey) + SORT_KEY_SEPARATOR + i);
        }
        sortKeys.sort();
        var sortedEntries = [];
        for (var j = 0; j < sortKeys.length; j++) {
            sortedEntries.push(entries[Number(sortKeys[j].substring(sortKeys[j].lastIndexOf(SORT_KEY_SEPARATOR) + 1))]);
        }
        return sortedEntries;
    }

    /**
     * パネルを表示する（メニューの項目を選ぶのと同じ）。
     * メニューの項目は表示／非表示の切り替えなので、表示中のパネルは2回選んで閉じてから開き直し、前面に出す
     * （同じスクリプトの中では、選んだ直後の checked が古い値のまま返るので当てにしない）
     * @param {PanelEntry[]} entries - 表示するパネル
     * @returns {string[]} 表示できなかったパネルの名前
     */
    function showPanels(entries) {
        var failedNames = [];
        for (var i = 0; i < entries.length; i++) {
            /* 無効な項目や、拡張パネルの読み込み失敗は例外になる / Disabled items or broken extension panels throw */
            try {
                var menuAction = app.menuActions.itemByID(entries[i].actionId);
                menuAction.invoke();
                if (entries[i].isOpen) menuAction.invoke();
            } catch (e) {
                failedNames.push(entries[i].name);
            }
        }
        return failedNames;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 件数入りのダイアログタイトルを作る
     * @param {number} itemCount - 表示中の件数
     * @returns {string} ダイアログタイトル
     */
    function buildDialogTitle(itemCount) {
        var countText = itemCount + getLabel(LABELS.counter.itemUnit);
        return getLabel(LABELS.dialog.title) + (uiLang === "ja" ? "（" + countText + "）" : " (" + countText + ")");
    }

    /**
     * 名前・印・分類・表示中の4列のリストを作る
     * @param {Window} parent - 追加先
     * @returns {ListBox} パネルの一覧
     */
    function buildPanelList(parent) {
        var panelList = parent.add("listbox", [0, 0, LIST_SIZE[0], LIST_SIZE[1]], "", {
            multiselect: true,
            numberOfColumns: 4,
            showHeaders: true,
            columnTitles: [
                getLabel(LABELS.column.name),
                getLabel(LABELS.column.mark),
                getLabel(LABELS.column.category),
                getLabel(LABELS.column.open)
            ],
            columnWidths: COLUMN_WIDTHS
        });
        panelList.preferredSize = LIST_SIZE;
        panelList.alignment = ["fill", "fill"];
        panelList.helpTip = getLabel(LABELS.tooltip.list);
        return panelList;
    }

    /**
     * 右揃えのラベル付きの行を作る
     * @param {Window} parent - 追加先
     * @param {Object} labelSet - 行ラベル
     * @returns {Group} 行グループ（入力部品はここへ追加する）
     */
    function addFieldRow(parent, labelSet) {
        var rowGroup = parent.add("group");
        setupRow(rowGroup, ["fill", "top"]);
        var rowLabel = rowGroup.add("statictext", undefined, labelText(labelSet));
        /* 幅を固定しないと右揃えが効かない / Right alignment needs a fixed width */
        rowLabel.minimumSize.width = ROW_LABEL_WIDTH;
        rowLabel.preferredSize.width = ROW_LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowGroup;
    }

    /**
     * ラジオボタンの行を作る（同じ行グループに入れて排他にする）
     * @param {Window} parent - 追加先
     * @param {Object} rowLabelSet - 行ラベル
     * @param {string[]} optionKeys - 選択肢の鍵（ラジオボタンの順）
     * @param {Object} optionLabels - 鍵ごとの文言
     * @param {Object} tooltipLabel - ツールチップ
     * @param {string} selectedKey - 最初に選ぶ鍵
     * @returns {RadioButton[]} optionKeys と同じ順のラジオボタン
     */
    function buildRadioRow(parent, rowLabelSet, optionKeys, optionLabels, tooltipLabel, selectedKey) {
        var radioRow = addFieldRow(parent, rowLabelSet);
        var radios = [];
        for (var i = 0; i < optionKeys.length; i++) {
            var radio = radioRow.add("radiobutton", undefined, getLabel(optionLabels[optionKeys[i]]));
            radio.helpTip = getLabel(tooltipLabel);
            radio.value = (optionKeys[i] === selectedKey);
            radios.push(radio);
        }
        return radios;
    }

    /**
     * セットのボタンの行を作る
     * @param {Window} parent - 追加先
     * @returns {Button[]} PANEL_SETS と同じ順のボタン
     */
    function buildPanelSetRow(parent) {
        var setRow = addFieldRow(parent, LABELS.fieldLabel.panelSet);
        var setButtons = [];
        for (var i = 0; i < PANEL_SETS.length; i++) {
            var setButton = setRow.add("button", undefined, getLabel(PANEL_SETS[i].label));
            setButton.helpTip = getLabel(LABELS.tooltip.panelSet, [joinPanelSetNames(PANEL_SETS[i])]);
            setButtons.push(setButton);
        }
        return setButtons;
    }

    /**
     * チェックボックスの行を作る
     * @param {Window} parent - 追加先
     * @param {Object} rowLabelSet - 行ラベル
     * @param {string[]} optionKeys - 選択肢の鍵（並びの順）
     * @param {Object} optionLabels - 鍵ごとの文言
     * @param {Object} tooltipLabel - ツールチップ
     * @param {Object} checkedKeys - 鍵ごとの最初の値
     * @returns {Checkbox[]} optionKeys と同じ順のチェックボックス
     */
    function buildCheckboxRow(parent, rowLabelSet, optionKeys, optionLabels, tooltipLabel, checkedKeys) {
        var checkboxRow = addFieldRow(parent, rowLabelSet);
        var checkboxes = [];
        for (var i = 0; i < optionKeys.length; i++) {
            var checkbox = checkboxRow.add("checkbox", undefined, getLabel(optionLabels[optionKeys[i]]));
            checkbox.helpTip = getLabel(tooltipLabel);
            checkbox.value = !!checkedKeys[optionKeys[i]];
            checkboxes.push(checkbox);
        }
        return checkboxes;
    }

    /**
     * 印を付けるボタン（♥・♢）を作る
     * @param {Group} parent - 追加先
     * @param {string} markKey - "fav"／"rare"
     * @param {Object} tooltipLabel - ツールチップ
     * @returns {Button} ボタン
     */
    function addMarkButton(parent, markKey, tooltipLabel) {
        var markButton = parent.add("button", undefined, MARK_GLYPHS[markKey]);
        markButton.preferredSize.width = MARK_BUTTON_WIDTH;
        markButton.helpTip = getLabel(tooltipLabel);
        return markButton;
    }

    /**
     * 検索欄の行（検索欄とクリアボタン）を作る
     * @param {Window} parent - 追加先
     * @returns {{input: EditText, clearButton: Button}} 検索欄とクリアボタン
     */
    function buildSearchRow(parent) {
        var searchRow = addFieldRow(parent, LABELS.fieldLabel.search);
        var searchInput = searchRow.add("edittext", undefined, "");
        searchInput.alignment = ["fill", "center"];
        searchInput.helpTip = getLabel(LABELS.tooltip.search);
        searchInput.active = true;
        var clearButton = addClearButton(searchRow);
        return { input: searchInput, clearButton: clearButton };
    }

    /**
     * Enter／Return のキーイベントか
     * @param {KeyboardEvent} keyEvent - キーイベント
     * @returns {boolean} Enter／Return なら true
     */
    function isEnterKey(keyEvent) {
        return keyEvent.keyName === "Enter" || keyEvent.keyName === "Return";
    }

    /**
     * ダイアログを表示し、表示するパネルを選ばせる。印と絞り込みは変えるたびに設定ファイルへ書き出す
     * @param {PanelEntry[]} allEntries - すべてのパネル
     * @param {LauncherSettings} settings - 読んだ設定（印と絞り込みをここに書き換える）
     * @returns {{entries: PanelEntry[], missingNames: string[]}} 表示するパネル（キャンセルなら空配列）と、セットのうち見つからなかったパネルの名前
     */
    function showPanelDialog(allEntries, settings) {
        var visibleEntries = [];
        var pickedEntries = [];
        var missingNames = [];
        var lastSearchText = null;
        var currentSortKey = DEFAULT_SORT_KEY;

        var mainDialog = new Window("dialog", getLabel(LABELS.dialog.title));
        setupWindow(mainDialog);

        var searchUI = buildSearchRow(mainDialog);
        var searchInput = searchUI.input;
        var clearButton = searchUI.clearButton;
        var filterChecks = buildCheckboxRow(mainDialog, LABELS.fieldLabel.filter, FILTER_KEYS, LABELS.filter, LABELS.tooltip.filter, settings.filters);
        var sortRadios = buildRadioRow(mainDialog, LABELS.fieldLabel.sortOrder, SORT_KEYS, LABELS.sort, LABELS.tooltip.sortOrder, currentSortKey);
        var setButtons = buildPanelSetRow(mainDialog);
        var panelList = buildPanelList(mainDialog);

        var buttonRow = addButtonRow(mainDialog);
        var markButtons = {
            fav:  addMarkButton(buttonRow.leftGroup, "fav", LABELS.tooltip.markFav),
            rare: addMarkButton(buttonRow.leftGroup, "rare", LABELS.tooltip.markRare)
        };
        buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnShow = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.show), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /**
         * リストで選んでいるパネルを返す
         * @returns {PanelEntry[]} 選んでいるパネル（無ければ空配列）
         */
        function getSelectedEntries() {
            var selectedEntries = [];
            var selectedItems = panelList.selection;
            if (!selectedItems) return selectedEntries;
            for (var i = 0; i < selectedItems.length; i++) {
                selectedEntries.push(visibleEntries[selectedItems[i].index]);
            }
            return selectedEntries;
        }

        /**
         * 絞り込み・並べ替えの結果で一覧を作り直す。
         * 前に選んでいたパネルが残っていれば選び直し、無ければ先頭を選ぶ（Enter ですぐ表示できるように）
         * @param {PanelEntry[]} keepEntries - 選び直すパネル
         * @returns {void}
         */
        function rebuildList(keepEntries) {
            visibleEntries = sortPanelEntries(filterPanelEntries(filterByMark(allEntries, settings.filters), searchInput.text), currentSortKey);
            panelList.removeAll();
            var reselectIndexes = [];
            for (var i = 0; i < visibleEntries.length; i++) {
                var listItem = panelList.add("item", visibleEntries[i].name);
                listItem.subItems[0].text = visibleEntries[i].mark ? MARK_GLYPHS[visibleEntries[i].mark] : "";
                listItem.subItems[1].text = visibleEntries[i].category;
                listItem.subItems[2].text = visibleEntries[i].isOpen ? OPEN_MARK : "";
                for (var j = 0; j < keepEntries.length; j++) {
                    if (keepEntries[j] === visibleEntries[i]) reselectIndexes.push(i);
                }
            }
            for (var k = 0; k < reselectIndexes.length; k++) {
                panelList.items[reselectIndexes[k]].selected = true;
            }
            if (reselectIndexes.length === 0 && visibleEntries.length > 0) panelList.selection = 0;
            mainDialog.text = buildDialogTitle(visibleEntries.length);
            btnShow.enabled = (visibleEntries.length > 0);
            updateMarkButtons();
        }

        /**
         * ♥・♢ボタンを選択に合わせる（選択が無ければ無効）
         * @returns {void}
         */
        function updateMarkButtons() {
            var selectedEntries = getSelectedEntries();
            for (var markKey in markButtons) {
                if (!markButtons.hasOwnProperty(markKey)) continue;
                markButtons[markKey].enabled = (selectedEntries.length > 0);
            }
        }

        /**
         * 選んだパネルの印を切り替える（すべてその印なら外し、そうでなければ付ける。♥と♢はどちらか一方）。
         * 絞り込みで外れたパネルは一覧から消える
         * @param {string} markKey - "fav"／"rare"
         * @returns {void}
         */
        function toggleMark(markKey) {
            var selectedEntries = getSelectedEntries();
            if (selectedEntries.length === 0) return;
            var newMark = allHaveMark(selectedEntries, markKey) ? "" : markKey;
            for (var i = 0; i < selectedEntries.length; i++) {
                setPanelMark(selectedEntries[i], newMark, settings.marks);
            }
            saveSettings(settings);
            rebuildList(selectedEntries);
            panelList.active = true;
        }

        /**
         * 絞り込みのチェックボックスで一覧を作り直し、選んでいたパネルを選び直す
         * @param {string} filterKey - 切り替えた絞り込み
         * @param {boolean} isOn - オンなら true
         * @returns {void}
         */
        function applyFilter(filterKey, isOn) {
            if (settings.filters[filterKey] === isOn) return;
            settings.filters[filterKey] = isOn;
            saveSettings(settings);
            rebuildList(getSelectedEntries());
        }

        /**
         * 検索欄の文字が前回から変わっていれば一覧を作り直し、先頭を選ぶ
         * （onChanging と onChange の両方から呼ばれても1回で済ませる）
         * @returns {void}
         */
        function applySearchText() {
            /* 検索欄が空ならクリアボタンを無効にする / Disable the clear button while the field is empty */
            var hasText = (searchInput.text !== "");
            if (clearButton.enabled !== hasText) {
                clearButton.enabled = hasText;
                clearButton.notify("onDraw");
            }
            if (searchInput.text === lastSearchText) return;
            lastSearchText = searchInput.text;
            rebuildList([]);
        }

        /**
         * 並び順のラジオボタンで並べ替え、選んでいたパネルを選び直す
         * @param {string} sortKey - 選んだ並び順
         * @returns {void}
         */
        function applySortKey(sortKey) {
            if (sortKey === currentSortKey) return;
            currentSortKey = sortKey;
            rebuildList(getSelectedEntries());
        }

        /**
         * リストで選んだパネルを控えてダイアログを閉じる（選択が無ければ先頭のパネル）。
         * モーダルダイアログの表示中は DOM を変更できないので、表示は閉じたあとで行う
         * @returns {void}
         */
        function pickAndClose() {
            if (visibleEntries.length === 0) return;
            pickedEntries = getSelectedEntries();
            if (pickedEntries.length === 0) pickedEntries.push(visibleEntries[0]);
            mainDialog.close(1);
        }

        /**
         * セットのパネルを控えてダイアログを閉じる（絞り込み・検索・選択に関係なく、セットに書いたパネル）
         * @param {Object} panelSet - PANEL_SETS の1件
         * @returns {void}
         */
        function pickSetAndClose(panelSet) {
            var found = findPanelSetEntries(allEntries, panelSet);
            pickedEntries = found.entries;
            missingNames = found.missingNames;
            mainDialog.close(1);
        }

        searchInput.onChanging = applySearchText;
        searchInput.onChange = applySearchText;
        for (var i = 0; i < sortRadios.length; i++) {
            /* i をその場で閉じ込める / Capture i for each radio */
            sortRadios[i].onClick = (function (sortKey) {
                return function () { applySortKey(sortKey); };
            }(SORT_KEYS[i]));
        }
        for (var f = 0; f < filterChecks.length; f++) {
            filterChecks[f].onClick = (function (filterKey) {
                return function () { applyFilter(filterKey, this.value); };
            }(FILTER_KEYS[f]));
        }
        for (var p = 0; p < setButtons.length; p++) {
            setButtons[p].onClick = (function (panelSet) {
                return function () { pickSetAndClose(panelSet); };
            }(PANEL_SETS[p]));
        }
        markButtons.fav.onClick = function () { toggleMark("fav"); };
        markButtons.rare.onClick = function () { toggleMark("rare"); };
        panelList.onChange = updateMarkButtons;
        searchInput.addEventListener("keydown", function (keyEvent) {
            if (keyEvent.keyName === "Down" && panelList.items.length > 0) {
                panelList.active = true;
                if (!panelList.selection) panelList.selection = 0;
                keyEvent.preventDefault();
            } else if (isEnterKey(keyEvent)) {
                pickAndClose();
                keyEvent.preventDefault();
                if (keyEvent.stopPropagation) keyEvent.stopPropagation();
            }
        });
        clearButton.onClick = function () {
            searchInput.text = "";
            /* 続けて打ち直せるよう、フォーカスは入力欄へ戻す / Put the focus back in the field */
            searchInput.active = true;
            applySearchText();
        };
        panelList.onDoubleClick = pickAndClose;
        panelList.addEventListener("keydown", function (keyEvent) {
            if (isEnterKey(keyEvent)) {
                pickAndClose();
                keyEvent.preventDefault();
                if (keyEvent.stopPropagation) keyEvent.stopPropagation();
            }
        });
        btnShow.onClick = pickAndClose;

        applySearchText();
        addVersionFooter(mainDialog);
        if (mainDialog.show() !== 1) return { entries: [], missingNames: [] };
        return { entries: pickedEntries, missingNames: missingNames };
    }

    // =========================================
    // 実行 / Run
    // =========================================

    var allEntries = collectPanelEntries();
    if (allEntries.length === 0) {
        alert(getLabel(LABELS.alert.notFound));
        return;
    }
    var settings = loadSettings();
    applySavedMarks(allEntries, settings.marks);
    var dialogResult = showPanelDialog(allEntries, settings);
    var failedNames = dialogResult.missingNames.concat(showPanels(dialogResult.entries));
    if (failedNames.length > 0) {
        alert(labelText(LABELS.alert.showFailed) + "\n" + failedNames.join("\n"));
    }

}());
