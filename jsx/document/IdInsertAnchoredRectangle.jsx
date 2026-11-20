#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdInsertPageCrossReference
Name=カーソル位置にアンカー付き長方形とテキストアンカーを作成
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=2.7.7
Release-Date=2026-11-06
Description-BEGIN
カーソルを置いた位置に、アンカー付き（カスタム）でプリントしない小さな長方形と、その直前のテキストアンカーを作成します。
テキストアンカーの名前と長方形の塗り・線・位置（ページの左／右）はダイアログボックスで指定でき、ハイパーリンクや相互参照の行き先として使えます。

Creates a small non-printing anchored (custom) rectangle at the text cursor, plus a text anchor right before it.
The anchor name and the rectangle's fill, stroke and position (left or right of the page) are set in a dialog box, and the anchor can serve as a target for hyperlinks and cross-references.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

カーソルを置いた位置に、アンカー付き（カスタム）でプリントしない小さな長方形と、その直前のテキストアンカーを作成します。
テキストアンカーの名前と長方形の塗り・線・位置（ページの左／右）はダイアログボックスで指定でき、ハイパーリンクや相互参照の行き先として使えます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdInsertAnchoredRectangle.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nde76a2dc06ae

### Overview

Creates a small non-printing anchored (custom) rectangle at the text cursor, plus a text anchor right before it.
The anchor name and the rectangle's fill, stroke and position (left or right of the page) are set in a dialog box, and the anchor can serve as a target for hyperlinks and cross-references.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdInsertAnchoredRectangle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdInsertAnchoredRectangle";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.7.7";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "カーソル位置にアンカー付き長方形とテキストアンカーを作成"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdInsertAnchoredRectangle.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdInsertAnchoredRectangle.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nde76a2dc06ae"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdInsertAnchoredRectangle.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 基本設定 / Settings
    // =========================================

    /* 長方形の幅・高さ（pt）/ Rectangle width and height (pt) */
    var RECT_WIDTH  = 12;
    var RECT_HEIGHT = 12;

    /* 塗りの初期値の CMYK（線の初期値はなし）/ Default fill CMYK (the stroke defaults to none) */
    var DEFAULT_FILL_CMYK = [0, 100, 100, 0];

    /* 線幅の初期値（pt）/ Default stroke weight (pt) */
    var DEFAULT_STROKE_WEIGHT = 0.5;

    /* 長方形を置くページの端の初期値（"left" / "right"）/ Default page edge for the rectangle ("left" / "right") */
    var DEFAULT_PAGE_SIDE = "left";

    /* テキストアンカー名の初期値の形（p12_01）/ Form of the default text anchor name (p12_01) */
    var ANCHOR_NAME_PAGE_PREFIX = "p";
    var ANCHOR_NAME_SEPARATOR   = "_";
    var ANCHOR_NAME_DIGITS      = 2;

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

    /* 行ラベルの幅（px）/ Row label width (px) */
    var ROW_LABEL_WIDTH = 56;

    /* 色見本の一辺（px）/ Side of the color chip (px) */
    var SWATCH_CHIP_SIZE = 16;

    /* 名前の欄・線幅の欄の文字数 / Characters of the name and stroke weight fields */
    var NAME_FIELD_CHARACTERS   = 20;
    var WEIGHT_FIELD_CHARACTERS = 5;

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
        undoName: { ja: "アンカー付き長方形とテキストアンカーを作成", en: "Create Anchored Rectangle and Text Anchor" },
        dialog: {
            title: { ja: "アンカー付き長方形とテキストアンカーを作成", en: "Create Anchored Rectangle and Text Anchor" }
        },
        panel: {
            anchor:    { ja: "テキストアンカー", en: "Text Anchor" },
            rectangle: { ja: "長方形", en: "Rectangle" }
        },
        field: {
            name:         { ja: "名前", en: "Name" },
            fill:         { ja: "塗り", en: "Fill" },
            stroke:       { ja: "線", en: "Stroke" },
            strokeWeight: { ja: "線幅", en: "Weight" },
            pageSide:     { ja: "位置", en: "Position" }
        },
        radio: {
            pageLeft:  { ja: "ページの左", en: "Left of page" },
            pageRight: { ja: "ページの右", en: "Right of page" }
        },
        swatch: {
            None:         { ja: "なし", en: "None" },
            Paper:        { ja: "紙色", en: "Paper" },
            Black:        { ja: "黒", en: "Black" },
            Registration: { ja: "レジストレーション", en: "Registration" }
        },
        tooltip: {
            name: { ja: "初期値は「p＋ページ番号＋_＋通し番号」です", en: "Defaults to p + page number + _ + sequence number" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument:     { ja: "InDesignドキュメントを開いてください。", en: "Open an InDesign document." },
            noTextSelected: {
                ja: "参照先にしたいテキスト内へカーソルを置くか、テキストを選択してから実行してください。",
                en: "Place the text cursor in the text you want to target, or select text, and run the script again."
            },
            notOnPage: {
                ja: "ペーストボード上またはオーバーセット状態のテキストには、このスクリプトを使用できません。",
                en: "This script cannot be used on text on the pasteboard or in overset text."
            },
            noAnchorCharacter: {
                ja: "アンカー付きオブジェクトのマーカー文字を取得できませんでした。",
                en: "Could not get the marker character of the anchored object."
            },
            emptyName:     { ja: "テキストアンカーの名前を入力してください。", en: "Enter a name for the text anchor." },
            duplicateName: { ja: "同じ名前のテキストアンカーがあります : %1", en: "A text anchor with this name already exists: %1" },
            invalidWeight: { ja: "線幅には 0 より大きい数値を入力してください。", en: "Enter a stroke weight greater than 0." }
        }
    };

    // =========================================
    // 共通部品 / Shared parts
    // =========================================

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

    // =========================================
    // ダイアログボックス / Dialog box
    // =========================================

    /**
     * 右揃えの行ラベルを足す
     * @param {Group} rowGroup - 行グループ
     * @param {Object} labelEntry - { ja, en }
     * @returns {StaticText} 行ラベル
     */
    function addRowLabel(rowGroup, labelEntry) {
        var rowLabel = rowGroup.add("statictext", undefined, labelText(labelEntry));
        rowLabel.preferredSize.width = ROW_LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowLabel;
    }

    /**
     * 「ラベル + スウォッチのドロップダウンリスト」の行を足す
     * @param {Panel} parentPanel - 行を足すパネル
     * @param {Object} labelEntry - { ja, en }
     * @param {Array} swatchChoices - getSwatchChoices() の items
     * @param {number} selectedIndex - 最初に選ぶ項目の番号
     * @param {Function} [onSelectionChange] - 選択が変わったときに呼ぶ関数
     * @returns {DropDownList} ドロップダウンリスト
     */
    function addSwatchRow(parentPanel, labelEntry, swatchChoices, selectedIndex, onSelectionChange) {
        var swatchRow = parentPanel.add("group");
        setupRow(swatchRow, "fill");
        addRowLabel(swatchRow, labelEntry);
        var swatchLabels = [];
        for (var i = 0; i < swatchChoices.length; i++) swatchLabels.push(swatchChoices[i].label);
        var swatchDropdown = swatchRow.add("dropdownlist", undefined, swatchLabels);
        swatchDropdown.alignment = ["fill", "center"];
        swatchDropdown.selection = selectedIndex;

        var swatchChip = addSwatchChip(swatchRow);
        paintSwatchChip(swatchChip, swatchChoices[selectedIndex].rgb);
        swatchDropdown.onChange = function () {
            paintSwatchChip(swatchChip, swatchChoices[swatchDropdown.selection.index].rgb);
            swatchChip.window.update();
            if (onSelectionChange) onSelectionChange();
        };
        return swatchDropdown;
    }

    /**
     * 色見本を足す
     * @param {Group} rowGroup - 行グループ
     * @returns {Group} 色見本
     */
    function addSwatchChip(rowGroup) {
        var swatchChip = rowGroup.add("group");
        swatchChip.minimumSize = swatchChip.preferredSize = swatchChip.maximumSize = [SWATCH_CHIP_SIZE, SWATCH_CHIP_SIZE];
        return swatchChip;
    }

    /**
     * 色見本を塗る
     * @param {Group} swatchChip - 色見本
     * @param {Array|null} chipRgb - [R, G, B]（0〜1）。［なし］の null は白で塗る
     * @returns {void}
     */
    function paintSwatchChip(swatchChip, chipRgb) {
        var fillRgb = chipRgb || [1, 1, 1];
        swatchChip.graphics.backgroundColor = swatchChip.graphics.newBrush(
            swatchChip.graphics.BrushType.SOLID_COLOR,
            [fillRgb[0], fillRgb[1], fillRgb[2], 1]
        );
    }

    /**
     * 名前と塗り・線を指定するダイアログボックスを表示する
     * @param {string} defaultName - 名前の初期値
     * @param {{items: Array, fillIndex: number, strokeIndex: number}} swatchChoices - getSwatchChoices() の戻り値
     * @param {Object} usedNames - 使用中のテキストアンカー名（名前をキーにしたオブジェクト）
     * @returns {{anchorName: string, fillChoice: Object, strokeChoice: Object, strokeWeight: number, pageSide: string}|null} 指定内容。キャンセルなら null
     */
    function showOptionsDialog(defaultName, swatchChoices, usedNames) {
        var optionsDialog = new Window("dialog", getLabel(LABELS.dialog.title));
        setupWindow(optionsDialog);

        var anchorPanel = optionsDialog.add("panel", undefined, getLabel(LABELS.panel.anchor));
        setupPanel(anchorPanel);
        var nameRow = anchorPanel.add("group");
        setupRow(nameRow, "fill");
        addRowLabel(nameRow, LABELS.field.name);
        var nameField = nameRow.add("edittext", undefined, defaultName);
        nameField.characters = NAME_FIELD_CHARACTERS;
        nameField.alignment = ["fill", "center"];
        nameField.helpTip = getLabel(LABELS.tooltip.name);

        var rectanglePanel = optionsDialog.add("panel", undefined, getLabel(LABELS.panel.rectangle));
        setupPanel(rectanglePanel);
        var fillDropdown = addSwatchRow(rectanglePanel, LABELS.field.fill, swatchChoices.items, swatchChoices.fillIndex);
        var strokeDropdown = addSwatchRow(rectanglePanel, LABELS.field.stroke, swatchChoices.items, swatchChoices.strokeIndex, updateWeightEnabled);

        var weightRow = rectanglePanel.add("group");
        setupRow(weightRow, "fill");
        addRowLabel(weightRow, LABELS.field.strokeWeight);
        var weightField = weightRow.add("edittext", undefined, String(DEFAULT_STROKE_WEIGHT));
        weightField.characters = WEIGHT_FIELD_CHARACTERS;
        var weightUnit = weightRow.add("statictext", undefined, "pt");

        /**
         * 線が「なし」のときは線幅の欄を無効にする
         * @returns {void}
         */
        function updateWeightEnabled() {
            var hasStroke = swatchChoices.items[strokeDropdown.selection.index].name !== "None";
            weightField.enabled = hasStroke;
            weightUnit.enabled = hasStroke;
        }
        updateWeightEnabled();

        var sideRow = rectanglePanel.add("group");
        setupRow(sideRow, "fill");
        addRowLabel(sideRow, LABELS.field.pageSide);
        var pageLeftRadio = sideRow.add("radiobutton", undefined, getLabel(LABELS.radio.pageLeft));
        var pageRightRadio = sideRow.add("radiobutton", undefined, getLabel(LABELS.radio.pageRight));
        pageLeftRadio.value = (DEFAULT_PAGE_SIDE !== "right");
        pageRightRadio.value = (DEFAULT_PAGE_SIDE === "right");

        var buttonRow = addButtonRow(optionsDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        var dialogResult = null;
        btnOK.onClick = function () {
            var anchorName = nameField.text.replace(/^\s+|\s+$/g, "");
            if (anchorName === "") {
                alert(getLabel(LABELS.alert.emptyName));
                return;
            }
            if (usedNames.hasOwnProperty(anchorName)) {
                alert(getLabel(LABELS.alert.duplicateName, [anchorName]));
                return;
            }
            var strokeWeight = parseFloat(weightField.text);
            if (weightField.enabled && !(strokeWeight > 0)) {
                alert(getLabel(LABELS.alert.invalidWeight));
                return;
            }
            dialogResult = {
                anchorName: anchorName,
                fillChoice: swatchChoices.items[fillDropdown.selection.index],
                strokeChoice: swatchChoices.items[strokeDropdown.selection.index],
                strokeWeight: strokeWeight,
                pageSide: pageRightRadio.value ? "right" : "left"
            };
            optionsDialog.close(1);
        };

        alignRightOnlyButtonRow(buttonRow);
        addVersionFooter(optionsDialog);
        nameField.active = true;
        return (optionsDialog.show() === 1) ? dialogResult : null;
    }

    // =========================================
    // 処理 / Processing
    // =========================================

    /**
     * 選択がテキスト（カーソル位置を含む）かどうか
     * @param {Object} selectedItem - app.selection の要素
     * @returns {boolean} テキストなら true（テキストフレームそのものの選択は false）
     */
    function isTextSelection(selectedItem) {
        return selectedItem != null && selectedItem.isValid && selectedItem.hasOwnProperty("baseline");
    }

    /**
     * 挿入点が載っているページを返す
     * @param {InsertionPoint} insertionPoint - 対象の挿入点
     * @returns {Page|null} ページ。オーバーセットやペーストボード上なら null
     */
    function getPageOf(insertionPoint) {
        var frames = insertionPoint.parentTextFrames;
        if (frames.length === 0) return null;
        var page = frames[0].parentPage;
        return (page && page.isValid) ? page : null;
    }

    /**
     * 使用中のテキストアンカー名を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object} 名前をキーにしたオブジェクト
     */
    function getUsedAnchorNames(doc) {
        var usedNames = {};
        var destinations = doc.hyperlinkTextDestinations;
        for (var i = 0; i < destinations.length; i++) usedNames[destinations[i].name] = true;
        return usedNames;
    }

    /**
     * 名前の初期値を作る（p12_01 の形。通し番号はそのページで空いている最小の番号）
     * @param {string} pageName - ページ番号
     * @param {Object} usedNames - 使用中のテキストアンカー名
     * @returns {string} テキストアンカー名
     */
    function makeDefaultAnchorName(pageName, usedNames) {
        var candidate;
        var sequence = 0;
        do {
            sequence++;
            var sequenceText = String(sequence);
            while (sequenceText.length < ANCHOR_NAME_DIGITS) sequenceText = "0" + sequenceText;
            candidate = ANCHOR_NAME_PAGE_PREFIX + pageName + ANCHOR_NAME_SEPARATOR + sequenceText;
        } while (usedNames.hasOwnProperty(candidate));
        return candidate;
    }

    /**
     * CMYK の値からスウォッチ名を作る
     * @param {Array} cmykValues - [C, M, Y, K]（0〜100）
     * @returns {string} 「C=0 M=100 Y=100 K=0」の形の名前
     */
    function makeCmykColorName(cmykValues) {
        return "C=" + cmykValues[0] + " M=" + cmykValues[1] + " Y=" + cmykValues[2] + " K=" + cmykValues[3];
    }

    /**
     * CMYK を色見本用の RGB に簡易換算する
     * @param {Array} cmykValues - [C, M, Y, K]（0〜100）
     * @returns {Array} [R, G, B]（0〜1）
     */
    function cmykToRgb(cmykValues) {
        var keyFactor = 1 - cmykValues[3] / 100;
        return [
            (1 - cmykValues[0] / 100) * keyFactor,
            (1 - cmykValues[1] / 100) * keyFactor,
            (1 - cmykValues[2] / 100) * keyFactor
        ];
    }

    /**
     * スウォッチのプロパティを読む。種類に無いプロパティは例外になるので undefined にする
     * @param {Swatch} swatchItem - スウォッチ
     * @param {string} propertyName - プロパティ名
     * @returns {*} 値。その種類に無ければ undefined
     */
    function readSwatchProperty(swatchItem, propertyName) {
        var propertyValue;
        try {
            /* return の式の中で例外が出ると catch に届かなかったので、代入してから返す / An exception in a return expression did not reach catch, so assign first */
            propertyValue = swatchItem[propertyName];
        } catch (e) {
            propertyValue = undefined;
        }
        return propertyValue;
    }

    /**
     * スウォッチの色見本の RGB を返す（表示用の簡易換算）。
     * doc.swatches の要素は種類によらず constructor.name が "Swatch" なので、持っているプロパティで見分ける。
     * 濃淡は白と混ぜ、グラデーションは先頭の色、Lab は明度の灰色、それ以外（混合インキなど）は中間の灰色にする
     * @param {Swatch} swatchItem - スウォッチ
     * @returns {Array} [R, G, B]（0〜1）
     */
    function getPreviewRgb(swatchItem) {
        var tintValue = readSwatchProperty(swatchItem, "tintValue");
        var baseColor = readSwatchProperty(swatchItem, "baseColor");
        if (tintValue !== undefined && baseColor) {
            var baseRgb = getPreviewRgb(baseColor);
            var tintRatio = (tintValue < 0) ? 1 : tintValue / 100;
            return [
                1 - (1 - baseRgb[0]) * tintRatio,
                1 - (1 - baseRgb[1]) * tintRatio,
                1 - (1 - baseRgb[2]) * tintRatio
            ];
        }

        var gradientStops = readSwatchProperty(swatchItem, "gradientStops");
        if (gradientStops && gradientStops.length > 0) return getPreviewRgb(gradientStops[0].stopColor);

        var colorValues = readSwatchProperty(swatchItem, "colorValue");
        if (colorValues) {
            var colorSpace = swatchItem.space;
            if (colorSpace === ColorSpace.CMYK) return cmykToRgb(colorValues);
            if (colorSpace === ColorSpace.RGB) return [colorValues[0] / 255, colorValues[1] / 255, colorValues[2] / 255];
            if (colorSpace === ColorSpace.LAB) return [colorValues[0] / 100, colorValues[0] / 100, colorValues[0] / 100];
        }
        return [0.5, 0.5, 0.5];
    }

    /**
     * ダイアログボックスのドロップダウンリストに並べるスウォッチの一覧を作る。
     * 塗りの初期値のスウォッチがドキュメントに無ければ、作成前の項目として末尾に足す
     * @param {Document} doc - 対象ドキュメント
     * @returns {{items: Array, fillIndex: number, strokeIndex: number}} 項目 { label, name, swatch, rgb } と、塗り・線の初期値の番号
     */
    function getSwatchChoices(doc) {
        var defaultFillName = makeCmykColorName(DEFAULT_FILL_CMYK);
        var choices = { items: [], fillIndex: -1, strokeIndex: 0 };
        var swatches = doc.swatches;
        for (var i = 0; i < swatches.length; i++) {
            var swatchName = swatches[i].name;
            var isBuiltIn = LABELS.swatch.hasOwnProperty(swatchName);
            choices.items.push({
                label: isBuiltIn ? "[" + getLabel(LABELS.swatch[swatchName]) + "]" : swatchName,
                name: swatchName,
                swatch: swatches[i],
                rgb: (swatchName === "None") ? null : getPreviewRgb(swatches[i])
            });
            if (swatchName === defaultFillName) choices.fillIndex = i;
            if (swatchName === "None") choices.strokeIndex = i;
        }
        if (choices.fillIndex < 0) {
            choices.items.push({ label: defaultFillName, name: defaultFillName, swatch: null, rgb: cmykToRgb(DEFAULT_FILL_CMYK) });
            choices.fillIndex = choices.items.length - 1;
        }
        return choices;
    }

    /**
     * ダイアログボックスの項目からスウォッチを返す。作成前の項目ならプロセスカラーを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {{label: string, name: string, swatch: Swatch|null}} swatchChoice - スウォッチの項目
     * @returns {Swatch} スウォッチ
     */
    function resolveSwatch(doc, swatchChoice) {
        if (swatchChoice.swatch) return swatchChoice.swatch;
        /* 塗りと線で同じ項目を選んだときは、2回目はできている色を使う / When fill and stroke share the item, the second call finds the new color */
        var color = doc.colors.itemByName(swatchChoice.name);
        if (color.isValid) return color;
        return doc.colors.add({ name: swatchChoice.name, model: ColorModel.PROCESS, space: ColorSpace.CMYK, colorValue: DEFAULT_FILL_CMYK });
    }

    /**
     * 指定の塗り・線でプリントしない長方形をページ上に作る（位置はアンカー化で変わるので、大きさだけ決める）
     * @param {Document} doc - 対象ドキュメント
     * @param {Page} page - 作成するページ
     * @param {{anchorName: string, fillChoice: Object, strokeChoice: Object, strokeWeight: number, pageSide: string}} rectangleOptions - ダイアログボックスの指定内容
     * @returns {Rectangle} 作成した長方形
     */
    function createRectangle(doc, page, rectangleOptions) {
        var rectangle = page.rectangles.add();
        /* 素の数値は定規の単位で解釈されるので pt を付ける / Plain numbers follow the ruler units, so add pt */
        rectangle.geometricBounds = ["0pt", "0pt", RECT_HEIGHT + "pt", RECT_WIDTH + "pt"];
        rectangle.fillColor = resolveSwatch(doc, rectangleOptions.fillChoice);
        rectangle.strokeColor = resolveSwatch(doc, rectangleOptions.strokeChoice);
        if (rectangleOptions.strokeChoice.name !== "None") rectangle.strokeWeight = rectangleOptions.strokeWeight + "pt";
        /* 目印なのでプリント・書き出しには出さない / A marker only, so keep it out of print and export */
        rectangle.nonprinting = true;
        /* テキストアンカーと同じ名前で追えるようにする / Lets the rectangle be matched with its text anchor by name */
        rectangle.label = rectangleOptions.anchorName;
        return rectangle;
    }

    /**
     * ページの左または右のマージンの幅を返す（見開きの左ページは left が内側＝画面の右なので入れ替える）
     * @param {Page} page - 対象のページ
     * @param {boolean} isRightSide - 右のマージンなら true
     * @returns {number} マージンの幅（定規の単位）
     */
    function getSideMargin(page, isRightSide) {
        var marginPrefs = page.marginPreferences;
        var isLeftHand = (page.side === PageSideOptions.LEFT_HAND);
        return (isRightSide !== isLeftHand) ? marginPrefs.right : marginPrefs.left;
    }

    /**
     * カーソル位置にアンカー付き長方形を入れ、その直前にテキストアンカーを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {Page} page - 長方形を作るページ
     * @param {InsertionPoint} insertionPoint - カーソル位置
     * @param {{anchorName: string, fillChoice: Object, strokeChoice: Object, strokeWeight: number, pageSide: string}} rectangleOptions - ダイアログボックスの指定内容
     * @returns {void}
     */
    function insertAnchoredRectangle(doc, page, insertionPoint, rectangleOptions) {
        var rectangle = createRectangle(doc, page, rectangleOptions);
        var anchorSettings = rectangle.anchoredObjectSettings;
        anchorSettings.insertAnchoredObject(insertionPoint, AnchorPosition.ANCHORED);

        /* カスタムのままだとページの左上に残るので、下端を行のベースラインに揃える（文章が流れても行に付いて動く）/ Custom leaves it at the top left of the page, so sit its bottom on the line's baseline (it follows the line when text reflows) */
        anchorSettings.verticalReferencePoint = VerticallyRelativeTo.LINE_BASELINE;
        anchorSettings.anchorYoffset = 0;

        /* 横はページの端とマージンの中間に中心を置く。見開きでも左右が入れ替わらないよう、ノドは基準にしない / Center it halfway between the page edge and the margin; not spine relative, so left and right stay put on facing pages */
        var isRightSide = (rectangleOptions.pageSide === "right");
        anchorSettings.spineRelative = false;
        anchorSettings.horizontalReferencePoint = AnchoredRelativeTo.PAGE_EDGE;
        anchorSettings.horizontalAlignment = isRightSide ? HorizontalAlignment.RIGHT_ALIGN : HorizontalAlignment.LEFT_ALIGN;
        anchorSettings.anchorPoint = AnchorPoint.BOTTOM_CENTER_ANCHOR;
        var halfMargin = getSideMargin(page, isRightSide) / 2;
        anchorSettings.anchorXoffset = isRightSide ? -halfMargin : halfMargin;

        /* アンカー化した長方形の親は、オブジェクト置換文字（Character）/ An anchored rectangle's parent is its object replacement Character */
        var anchorCharacter = rectangle.parent;
        if (!anchorCharacter || !anchorCharacter.isValid || anchorCharacter.constructor.name !== "Character") {
            alert(getLabel(LABELS.alert.noAnchorCharacter));
            return;
        }

        /* マーカー文字そのものではなく、直前の挿入点をテキストアンカーにする / Use the insertion point before the marker, not the marker itself */
        doc.hyperlinkTextDestinations.add(anchorCharacter.insertionPoints[0], { name: rectangleOptions.anchorName });
        app.select(rectangle);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめ、ダイアログボックスの指定で長方形とテキストアンカーを作る
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        if (app.selection.length !== 1 || !isTextSelection(app.selection[0])) {
            alert(getLabel(LABELS.alert.noTextSelected));
            return;
        }

        var doc = app.activeDocument;
        var insertionPoint = app.selection[0].insertionPoints[0];
        var page = getPageOf(insertionPoint);
        if (!page) {
            alert(getLabel(LABELS.alert.notOnPage));
            return;
        }

        /* モーダルのダイアログボックスの表示中は DOM を変えられないので、閉じてから作る / The DOM cannot change while the modal dialog is open, so create everything after it closes */
        var usedNames = getUsedAnchorNames(doc);
        var rectangleOptions = showOptionsDialog(makeDefaultAnchorName(page.name, usedNames), getSwatchChoices(doc), usedNames);
        if (!rectangleOptions) return;

        app.doScript(function () {
            insertAnchoredRectangle(doc, page, insertionPoint, rectangleOptions);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undoName));
    }

    main();

})();
