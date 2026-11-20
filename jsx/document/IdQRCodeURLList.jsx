#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdQRCodeURLList
Name=ドキュメント内のQRコードのURLを一覧
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.1.1
Release-Date=2026-11-05
Description-BEGIN
アクティブなドキュメントにある QR コード（［オブジェクト］→［QR コードを生成］で作ったもの）の URL を、ページ順に一覧します。
入稿前のリンク先の確認に使えます。一覧はクリップボードにコピーでき、選んだ QR コードのフレームの選択や、種類・URL の変更もできます。

Lists, in page order, the URLs of the QR codes in the active document (made with Object > Generate QR Code).
Handy for checking link targets before sending files to print. You can copy the list, select the frame of the chosen QR code, and change its type or URL.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

アクティブなドキュメントにある QR コード（［オブジェクト］→［QR コードを生成］で作ったもの）の URL を、ページ順に一覧します。
入稿前のリンク先の確認に使えます。一覧はクリップボードにコピーでき、選んだ QR コードのフレームの選択や、種類・URL の変更もできます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdQRCodeURLList.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ne2cdd5f60080

### Overview

Lists, in page order, the URLs of the QR codes in the active document (made with Object > Generate QR Code).
Handy for checking link targets before sending files to print. You can copy the list, select the frame of the chosen QR code, and change its type or URL.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdQRCodeURLList.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdQRCodeURLList";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "ドキュメント内のQRコードのURLを一覧"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdQRCodeURLList.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdQRCodeURLList.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ne2cdd5f60080"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdQRCodeURLList.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 基本設定 / Settings
    // =========================================

    /* プレーンテキストの QR コードも、中身が URL なら一覧に入れる / Also list plain-text QR codes whose text is a URL */
    var INCLUDE_PLAIN_TEXT_URLS = true;

    /* URL とみなすプレーンテキスト / Plain text treated as a URL */
    var URL_PATTERN = /^(?:https?|ftp):\/\/\S+$/i;

    /* QR コードの種類（スニペットの要素名）/ QR code types (element names in a snippet) */
    var QR_ELEMENT_PATTERN = /<(HyperlinkQRCode|PlainTextQRCode|EmailQRCode|TextMsgQRCode|VCardQRCode)\b([^>]*)>/;

    /* QR コードの色（スニペットのスウォッチ名）/ QR code color (swatch name in a snippet) */
    var QR_SWATCH_PATTERN = /<QrCodeSwatch\b[^>]*>([^<]*)<\/QrCodeSwatch>/;

    /* QR コードを埋め込んだ EPS のリンク名（オーバーセットの判定用）/ Link name of the embedded EPS of a QR code (for overset items) */
    var QR_LINK_NAME_PATTERN = /^QR Code/;

    /* 並べ替えキーの座標に足す値（負の座標を正にする）/ Offset that makes coordinates positive in the sort key */
    var SORT_COORDINATE_OFFSET = 100000;

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
    var COLUMN_WIDTHS = [110, 130, 360];     /* 列幅（ページ・種類・URL）/ column widths (page, type, URL) */
    var EDIT_LABEL_WIDTH     = 50;           /* 編集ダイアログの項目名の幅 / label width in the edit dialog */
    var EDIT_URL_FIELD_WIDTH = 420;          /* 編集ダイアログの URL 欄の幅 / URL field width in the edit dialog */

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
            title:     { ja: "QR コードの URL", en: "QR Code URLs" },
            editTitle: { ja: "QR コードを編集", en: "Edit QR Code" }
        },
        column: {
            page: { ja: "ページ", en: "Page" },
            type: { ja: "種類", en: "Type" },
            url:  { ja: "URL", en: "URL" }
        },
        fieldLabel: {
            type: { ja: "種類", en: "Type" },
            url:  { ja: "URL", en: "URL" }
        },
        qrType: {
            hyperlink: { ja: "Web ハイパーリンク", en: "Web Hyperlink" },
            plainText: { ja: "プレーンテキスト", en: "Plain Text" }
        },
        location: {
            pasteboard: { ja: "ペーストボード", en: "Pasteboard" },
            overset:    { ja: "オーバーセット", en: "Overset" }
        },
        message: {
            unreadable:   { ja: "（オーバーセットのため読み取れません）", en: "(cannot be read in overset text)" },
            multipleURLs: { ja: "（複数の URL）", en: "(multiple URLs)" }
        },
        button: {
            copyAll: { ja: "すべてコピー", en: "Copy All" },
            edit:    { ja: "編集...", en: "Edit..." },
            select:  { ja: "選択", en: "Select" },
            close:   { ja: "閉じる", en: "Close" },
            cancel:  { ja: "キャンセル", en: "Cancel" },
            ok:      { ja: "OK", en: "OK" }
        },
        tooltip: {
            list: {
                ja: "ダブルクリックすると、その QR コードのフレームを選択してダイアログボックスを閉じます。",
                en: "Double-click to select the frame of that QR code and close the dialog box."
            },
            copyAll: {
                ja: "一覧を「ページ<タブ>URL」の行にしてクリップボードにコピーします。表計算アプリにそのまま貼り付けられます。",
                en: "Copies the list to the clipboard as \"page<tab>URL\" lines, ready to paste into a spreadsheet."
            },
            edit: {
                ja: "リストで選んだ QR コードの種類と URL を変更します。複数選ぶと、まとめて変更できます。",
                en: "Changes the type and URL of the QR codes chosen in the list. Choose several to change them together."
            },
            editType: {
                ja: "QR コードの種類です。複数の種類が混ざっているときは、どちらも選ばなければ種類を変えません。",
                en: "The QR code type. When the chosen QR codes are of mixed types, leaving both off keeps each type."
            },
            editURL: {
                ja: "QR コードの URL です。URL が異なる QR コードを複数選んだときは変更できません。",
                en: "The URL of the QR code. It cannot be changed when the chosen QR codes have different URLs."
            },
            select: {
                ja: "リストで選んだ QR コードのフレームを選択し、そのページを表示します。",
                en: "Selects the frame of the QR code chosen in the list and shows its page."
            }
        },
        alert: {
            noDocument:   { ja: "ドキュメントを開いてから実行してください。", en: "Open a document first." },
            noQRCode:     { ja: "URL の QR コードは見つかりませんでした。", en: "No QR codes with a URL were found." },
            copyFailed:   { ja: "クリップボードにコピーできませんでした。", en: "Could not copy to the clipboard." },
            selectFailed: {
                ja: "フレームを選択できませんでした（ロック・非表示のレイヤー、オーバーセットなど）。",
                en: "Could not select the frame (locked or hidden layer, overset text, etc.)."
            },
            emptyURL: { ja: "URL を入力してください。", en: "Enter a URL." },
            plainTextNotURL: {
                ja: "プレーンテキストの QR コードは http:// https:// ftp:// で始まる URL にしてください。これ以外の内容は一覧に出なくなります。",
                en: "For a Plain Text QR code, use a URL starting with http://, https:// or ftp://. Any other text would drop out of the list."
            },
            editFailed: {
                ja: "%1件の QR コードを変更できませんでした（ロックされたオブジェクトなど）。\n%2",
                en: "Could not change %1 QR code(s) (locked objects, etc.).\n%2"
            }
        },
        counter: {
            itemUnit: { ja: "件", en: " items" }
        }
    };

    // =========================================
    // QR コードの読み取り / Reading QR codes
    // =========================================

    /**
     * @typedef {Object} QRCodeEntry
     * @property {PageItem} frame - QR コードを入れたフレーム
     * @property {string|null} qrType - 種類（"hyperlink" / "plainText"）。オーバーセットは読み取れないので null
     * @property {string} typeName - 種類の表示名
     * @property {string} url - URL
     * @property {string|null} swatchName - QR コードのスウォッチ名。指定なしは null
     * @property {string} pageLabel - ページの表示名
     * @property {Page|null} page - 移動先のページ（ペーストボードは null）
     * @property {Spread|MasterSpread|null} spread - 移動先のスプレッド（オーバーセットは null）
     * @property {string} sortKey - 並べ替えキー
     */

    /**
     * QR コードの候補か判定する。QR コードは埋め込みの EPS として入っている
     * @param {Graphic} graphic - 対象のグラフィック
     * @returns {boolean} 候補なら true
     */
    function isQRCodeCandidate(graphic) {
        if (graphic.constructor.name !== "EPS") return false;
        var graphicLink = graphic.itemLink;
        /* リンクした EPS ファイルは QR コードではない / A linked EPS file is not a QR code */
        return !graphicLink || graphicLink.status === LinkStatus.LINK_EMBEDDED;
    }

    /**
     * QR コードを埋め込んだ EPS のリンク名か判定する（スニペットに書き出せないオーバーセットのもの用）
     * @param {Graphic} graphic - 対象のグラフィック
     * @returns {boolean} QR コードのリンク名なら true
     */
    function isQRCodeLinkName(graphic) {
        var graphicLink = graphic.itemLink;
        return !!graphicLink && QR_LINK_NAME_PATTERN.test(graphicLink.name);
    }

    /**
     * フレームをスニペットに書き出し、QR コードの要素を取り出す。
     * QR コードの中身は DOM のプロパティに無く、スニペット（IDMS）の XML にだけ出る
     * @param {PageItem} frame - QR コードを入れたフレーム
     * @param {File} snippetFile - 一時ファイル
     * @returns {{elementName: string, attributes: string, swatchName: (string|null)}|null} 要素名・属性の文字列・スウォッチ名。QR コードでなければ null
     */
    function readQRCodeElement(frame, snippetFile) {
        frame.exportFile(ExportFormat.INDESIGN_SNIPPET, snippetFile);
        snippetFile.encoding = "UTF-8";
        snippetFile.open("r");
        var snippetText = snippetFile.read();
        snippetFile.close();
        snippetFile.remove();
        var elementMatch = snippetText.match(QR_ELEMENT_PATTERN);
        if (!elementMatch) return null;
        var swatchMatch = snippetText.match(QR_SWATCH_PATTERN);
        return {
            elementName: elementMatch[1],
            attributes: elementMatch[2],
            swatchName: swatchMatch ? decodeXMLText(swatchMatch[1]) : null
        };
    }

    /**
     * XML の属性値を取り出し、文字参照を戻す
     * @param {string} attributes - 属性の並び
     * @param {string} attributeName - 属性名
     * @returns {string} 属性値。無ければ空文字
     */
    function getXMLAttribute(attributes, attributeName) {
        var attributeMatch = attributes.match(new RegExp("\\b" + attributeName + "=\"([^\"]*)\""));
        return attributeMatch ? decodeXMLText(attributeMatch[1]) : "";
    }

    /**
     * XML の文字参照を戻す
     * @param {string} xmlText - XML の属性値や要素の文字列
     * @returns {string} 文字参照を戻した文字列
     */
    function decodeXMLText(xmlText) {
        return xmlText
            .replace(/&#x([0-9A-Fa-f]+);/g, function (whole, hex) { return String.fromCharCode(parseInt(hex, 16)); })
            .replace(/&#(\d+);/g, function (whole, digits) { return String.fromCharCode(parseInt(digits, 10)); })
            .replace(/&lt;/g, "<")
            .replace(/&gt;/g, ">")
            .replace(/&quot;/g, "\"")
            .replace(/&apos;/g, "'")
            .replace(/&amp;/g, "&");
    }

    /**
     * QR コードの要素から種類と URL を取り出す
     * @param {{elementName: string, attributes: string}} qrElement - readQRCodeElement() の戻り値
     * @returns {{qrType: string, url: string}|null} URL でなければ null
     */
    function extractURL(qrElement) {
        if (qrElement.elementName === "HyperlinkQRCode") {
            return { qrType: "hyperlink", url: getXMLAttribute(qrElement.attributes, "UrlLink") };
        }
        if (INCLUDE_PLAIN_TEXT_URLS && qrElement.elementName === "PlainTextQRCode") {
            var plainText = getXMLAttribute(qrElement.attributes, "PlainText").replace(/^\s+|\s+$/g, "");
            if (URL_PATTERN.test(plainText)) {
                return { qrType: "plainText", url: plainText };
            }
        }
        return null;
    }

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
     * フレームの上にあるスプレッドを探す（ペーストボードの判定用）。
     * アンカー付きオブジェクトは、文字を収めたテキストフレームからたどる
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
     * フレームの場所（ページの表示名・移動先・並べ替えキーの先頭）を調べる。
     * 並びはドキュメントページ → 親ページ → ペーストボード → オーバーセット
     * @param {PageItem} frame - 対象のフレーム
     * @returns {{pageLabel: string, page: Page|null, spread: Spread|MasterSpread|null, sortPrefix: string}} 場所
     */
    function describeLocation(frame) {
        var parentPage = frame.parentPage;
        if (parentPage) {
            var pageSpread = parentPage.parent;
            if (pageSpread.constructor.name === "MasterSpread") {
                return { pageLabel: pageSpread.name, page: parentPage, spread: pageSpread, sortPrefix: "1" + padNumber(pageSpread.index, 6) };
            }
            return { pageLabel: parentPage.name, page: parentPage, spread: pageSpread, sortPrefix: "0" + padNumber(parentPage.documentOffset, 6) };
        }
        var parentSpread = findParentSpread(frame);
        if (parentSpread) {
            return {
                pageLabel: getLabel(LABELS.location.pasteboard) + (uiLang === "ja" ? "（" + buildSpreadLabel(parentSpread) + "）" : " (" + buildSpreadLabel(parentSpread) + ")"),
                page: null,
                spread: parentSpread,
                sortPrefix: "2" + padNumber(parentSpread.index, 6)
            };
        }
        return { pageLabel: getLabel(LABELS.location.overset), page: null, spread: null, sortPrefix: "3000000" };
    }

    /**
     * ドキュメントの QR コードのうち、URL のものを集めてページ順に並べる
     * @param {Document} doc - 対象のドキュメント
     * @returns {QRCodeEntry[]} QR コードの一覧
     */
    function collectQRCodes(doc) {
        var graphics = doc.allGraphics;
        var snippetFile = new File(Folder.temp.fsName + "/" + SCRIPT_NAME + "_" + $.hiresTimer + ".idms");
        var entries = [];
        for (var i = 0; i < graphics.length; i++) {
            if (!isQRCodeCandidate(graphics[i])) continue;
            var frame = graphics[i].parent;
            var location = describeLocation(frame);
            var urlInfo;
            var swatchName = null;
            if (location.spread) {
                var qrElement = readQRCodeElement(frame, snippetFile);
                if (!qrElement) continue;
                urlInfo = extractURL(qrElement);
                if (!urlInfo) continue;
                swatchName = qrElement.swatchName;
            } else {
                /* オーバーセットの文字に入ったフレームをスニペットに書き出すと InDesign が落ちる / Exporting a frame in overset text as a snippet crashes InDesign */
                if (!isQRCodeLinkName(graphics[i])) continue;
                urlInfo = { qrType: null, url: getLabel(LABELS.message.unreadable) };
            }
            var bounds = frame.geometricBounds;
            entries.push({
                frame: frame,
                qrType: urlInfo.qrType,
                typeName: urlInfo.qrType ? getLabel(LABELS.qrType[urlInfo.qrType]) : "-",
                url: urlInfo.url,
                swatchName: swatchName,
                pageLabel: location.pageLabel,
                page: location.page,
                spread: location.spread,
                /* 場所 → 上 → 左 → 見つかった順 / location, top, left, then discovery order */
                sortKey: location.sortPrefix
                    + padNumber(bounds[0] + SORT_COORDINATE_OFFSET, 7)
                    + padNumber(bounds[1] + SORT_COORDINATE_OFFSET, 7)
                    + padNumber(i, 6)
            });
        }
        return sortEntries(entries);
    }

    /**
     * 並べ替えキーの順に並べる（比較関数つきの sort() は遅く並びも狂うため、キーの文字列で並べる）
     * @param {QRCodeEntry[]} entries - QR コードの一覧
     * @returns {QRCodeEntry[]} 並べ替えた一覧
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
     * 件数入りのダイアログタイトルを作る
     * @param {number} itemCount - 件数
     * @returns {string} ダイアログタイトル
     */
    function buildDialogTitle(itemCount) {
        var countText = formatThousands(itemCount) + getLabel(LABELS.counter.itemUnit);
        return getLabel(LABELS.dialog.title)
            + (uiLang === "ja" ? "（" + countText + "）" : " (" + countText + ")");
    }

    /**
     * 一覧を「ページ<タブ>URL」の行にまとめる
     * @param {QRCodeEntry[]} entries - QR コードの一覧
     * @returns {string} 改行区切りの文字列
     */
    function buildCopyText(entries) {
        var lines = [];
        for (var i = 0; i < entries.length; i++) {
            lines.push(entries[i].pageLabel + "\t" + entries[i].url);
        }
        return lines.join("\n");
    }

    /**
     * 文字列をクリップボードにコピーする（Mac は AppleScript、Windows は VBScript 経由）
     * @param {string} copyText - コピーする文字列
     * @returns {void}
     */
    function copyToClipboard(copyText) {
        var isMac = (File.fs === "Macintosh");
        var scriptSource = isMac
            ? 'set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"'
            : 'CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""').replace(/\r?\n/g, '" & vbCrLf & "').replace(/\t/g, '" & vbTab & "') + '"';
        /* スクリプト実行の許可がないと失敗する / Fails without permission to run scripts */
        try {
            app.doScript(scriptSource, isMac ? ScriptLanguage.APPLESCRIPT_LANGUAGE : ScriptLanguage.VISUAL_BASIC);
        } catch (e) {
            alert(getLabel(LABELS.alert.copyFailed) + "\n" + e.message);
        }
    }

    /**
     * QR コードのページ（またはスプレッド）を表示し、フレームを選択する
     * @param {QRCodeEntry} entry - 対象の QR コード
     * @returns {void}
     */
    function selectQRCode(entry) {
        try {
            if (entry.page) {
                app.activeWindow.activePage = entry.page;
            } else if (entry.spread) {
                app.activeWindow.activeSpread = entry.spread;
            }
            app.select(entry.frame);
        } catch (e) {
            alert(getLabel(LABELS.alert.selectFailed) + "\n" + e.message);
        }
    }

    /**
     * @typedef {Object} QRCodeChange
     * @property {QRCodeEntry} entry - 変更する QR コード
     * @property {string} qrType - 変更後の種類（"hyperlink" / "plainText"）
     * @property {string} url - 変更後の URL
     */

    /**
     * 編集の指定を、実際に変わる QR コードごとの変更にする
     * @param {QRCodeEntry[]} targets - 編集の対象
     * @param {string|null} newType - 新しい種類。null なら種類を変えない
     * @param {string|null} newURL - 新しい URL。null なら URL を変えない
     * @returns {QRCodeChange[]} 変更の一覧（変わらないものは含めない）
     */
    function buildChanges(targets, newType, newURL) {
        var changes = [];
        for (var i = 0; i < targets.length; i++) {
            var qrType = newType || targets[i].qrType;
            var url = (newURL !== null) ? newURL : targets[i].url;
            if (qrType === targets[i].qrType && url === targets[i].url) continue;
            changes.push({ entry: targets[i], qrType: qrType, url: url });
        }
        return changes;
    }

    /**
     * プレーンテキストになる変更のうち、中身が URL でないものがあるか判定する（一覧から外れてしまう）
     * @param {QRCodeChange[]} changes - 変更の一覧
     * @returns {boolean} あれば true
     */
    function hasNonURLPlainText(changes) {
        for (var i = 0; i < changes.length; i++) {
            if (changes[i].qrType === "plainText" && !URL_PATTERN.test(changes[i].url)) return true;
        }
        return false;
    }

    /**
     * 同じフレームに QR コードを作り直す。種類や中身を変えるプロパティは無いが、
     * 作り直しても大きさ・位置・回転は残る（色は元のスウォッチを渡さないと残らない）
     * @param {Document} doc - 対象のドキュメント
     * @param {QRCodeChange} change - 変更の内容
     * @returns {void}
     */
    function rebuildQRCode(doc, change) {
        var frame = change.entry.frame;
        var createMethod = (change.qrType === "plainText") ? "createPlainTextQRCode" : "createHyperlinkQRCode";
        if (change.entry.swatchName === null) {
            frame[createMethod](change.url);
        } else {
            frame[createMethod](change.url, doc.swatches.itemByName(change.entry.swatchName));
        }
    }

    /**
     * 変更をまとめて適用する（取り消しは1回で済む）
     * @param {Document} doc - 対象のドキュメント
     * @param {QRCodeChange[]} changes - 変更の一覧
     * @returns {{failedCount: number, lastMessage: string}} 失敗した件数と最後のエラー
     */
    function applyChanges(doc, changes) {
        var applyResult = { failedCount: 0, lastMessage: "" };
        app.doScript(function () {
            for (var i = 0; i < changes.length; i++) {
                /* ロックされたフレームは例外になる。残りの QR コードは続ける / A locked frame throws; carry on with the rest */
                try {
                    rebuildQRCode(doc, changes[i]);
                } catch (e) {
                    applyResult.failedCount++;
                    applyResult.lastMessage = e.message;
                }
            }
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.dialog.editTitle));
        return applyResult;
    }

    /**
     * 編集の対象に共通する値を返す
     * @param {QRCodeEntry[]} targets - 編集の対象
     * @param {string} propertyName - 調べるプロパティ名
     * @returns {string|null} すべて同じならその値、異なれば null
     */
    function getCommonValue(targets, propertyName) {
        for (var i = 1; i < targets.length; i++) {
            if (targets[i][propertyName] !== targets[0][propertyName]) return null;
        }
        return targets[0][propertyName];
    }

    /**
     * 種類と URL の編集ダイアログを表示する
     * @param {QRCodeEntry[]} targets - 編集の対象
     * @returns {QRCodeChange[]|null} 変更の一覧。キャンセルしたときや何も変わらないときは null
     */
    function showEditDialog(targets) {
        var editDialog = new Window("dialog", getLabel(LABELS.dialog.editTitle));
        setupWindow(editDialog);

        var commonURL = getCommonValue(targets, "url");
        var selectedType = getCommonValue(targets, "qrType");

        var typeRow = editDialog.add("group");
        setupRow(typeRow, "fill");
        addEditFieldLabel(typeRow, LABELS.fieldLabel.type);
        var radioHyperlink = typeRow.add("radiobutton", undefined, getLabel(LABELS.qrType.hyperlink));
        var radioPlainText = typeRow.add("radiobutton", undefined, getLabel(LABELS.qrType.plainText));
        radioHyperlink.value = (selectedType === "hyperlink");
        radioPlainText.value = (selectedType === "plainText");
        radioHyperlink.helpTip = radioPlainText.helpTip = getLabel(LABELS.tooltip.editType);

        var urlRow = editDialog.add("group");
        setupRow(urlRow, "fill");
        addEditFieldLabel(urlRow, LABELS.fieldLabel.url);
        var urlField = urlRow.add("edittext", undefined, (commonURL !== null) ? commonURL : getLabel(LABELS.message.multipleURLs));
        urlField.preferredSize.width = EDIT_URL_FIELD_WIDTH;
        urlField.enabled = (commonURL !== null);
        urlField.helpTip = getLabel(LABELS.tooltip.editURL);

        var buttonRow = addButtonRow(editDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        editDialog.defaultElement = btnOK;
        editDialog.cancelElement = btnCancel;
        alignRightOnlyButtonRow(buttonRow);

        var changes = null;

        radioHyperlink.onClick = function () {
            selectedType = "hyperlink";
        };
        radioPlainText.onClick = function () {
            selectedType = "plainText";
        };
        btnOK.onClick = function () {
            var newURL = null;
            if (commonURL !== null) {
                newURL = urlField.text.replace(/^\s+|\s+$/g, "");
                if (newURL === "") {
                    alert(getLabel(LABELS.alert.emptyURL));
                    return;
                }
            }
            var pendingChanges = buildChanges(targets, selectedType, newURL);
            if (hasNonURLPlainText(pendingChanges)) {
                alert(getLabel(LABELS.alert.plainTextNotURL));
                return;
            }
            changes = pendingChanges;
            editDialog.close(1);
        };
        btnCancel.onClick = function () {
            editDialog.close(2);
        };

        editDialog.show();
        return (changes && changes.length > 0) ? changes : null;
    }

    /**
     * 編集ダイアログの行ラベルを足す
     * @param {Group} rowGroup - 行
     * @param {Object} labelRef - 項目名の文言
     * @returns {void}
     */
    function addEditFieldLabel(rowGroup, labelRef) {
        var fieldLabel = rowGroup.add("statictext", undefined, labelText(labelRef));
        fieldLabel.preferredSize.width = EDIT_LABEL_WIDTH;
        fieldLabel.justify = "right";
    }

    /**
     * リストで選んでいる QR コードを返す
     * @param {ListBox} qrList - 一覧のリスト
     * @param {QRCodeEntry[]} entries - QR コードの一覧
     * @returns {QRCodeEntry[]} 選んでいる QR コード
     */
    function getSelectedEntries(qrList, entries) {
        var selectedEntries = [];
        for (var i = 0; i < qrList.items.length; i++) {
            if (qrList.items[i].selected) selectedEntries.push(entries[i]);
        }
        return selectedEntries;
    }

    /**
     * 編集できる QR コードだけを返す（オーバーセットは読み取れず、作り直すと InDesign が落ちるおそれがある）
     * @param {QRCodeEntry[]} entries - QR コードの一覧
     * @returns {QRCodeEntry[]} 編集できる QR コード
     */
    function getEditableEntries(entries) {
        var editableEntries = [];
        for (var i = 0; i < entries.length; i++) {
            if (entries[i].qrType) editableEntries.push(entries[i]);
        }
        return editableEntries;
    }

    /**
     * 編集した QR コードをリストで選び直す
     * @param {ListBox} qrList - 一覧のリスト
     * @param {QRCodeEntry[]} entries - QR コードの一覧
     * @param {number[]} frameIds - 選ぶフレームの id
     * @returns {boolean} 1つでも選んだら true
     */
    function restoreSelection(qrList, entries, frameIds) {
        var idText = "," + frameIds.join(",") + ",";
        var hasSelection = false;
        for (var i = 0; i < entries.length; i++) {
            if (idText.indexOf("," + entries[i].frame.id + ",") < 0) continue;
            qrList.items[i].selected = true;
            hasSelection = true;
        }
        return hasSelection;
    }

    /**
     * 一覧のダイアログを表示する
     * @param {QRCodeEntry[]} entries - QR コードの一覧
     * @param {number[]} selectedFrameIds - 最初に選んでおくフレームの id（無ければ先頭を選ぶ）
     * @returns {{chosenEntry: (QRCodeEntry|undefined), changes: (QRCodeChange[]|undefined)}|null} ［選択］なら chosenEntry、編集なら changes、閉じただけなら null
     */
    function showListDialog(entries, selectedFrameIds) {
        var listDialog = new Window("dialog", buildDialogTitle(entries.length));
        setupWindow(listDialog);

        var qrList = listDialog.add("listbox", [0, 0, LIST_SIZE[0], LIST_SIZE[1]], "", {
            multiselect: true,
            numberOfColumns: 3,
            showHeaders: true,
            columnTitles: [
                getLabel(LABELS.column.page),
                getLabel(LABELS.column.type),
                getLabel(LABELS.column.url)
            ],
            columnWidths: COLUMN_WIDTHS
        });
        qrList.preferredSize = LIST_SIZE;
        qrList.helpTip = getLabel(LABELS.tooltip.list);
        for (var i = 0; i < entries.length; i++) {
            var listItem = qrList.add("item", entries[i].pageLabel);
            listItem.subItems[0].text = entries[i].typeName;
            listItem.subItems[1].text = entries[i].url;
        }
        if (!restoreSelection(qrList, entries, selectedFrameIds)) qrList.items[0].selected = true;

        var buttonRow = addButtonRow(listDialog);
        var btnCopyAll = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.copyAll));
        btnCopyAll.helpTip = getLabel(LABELS.tooltip.copyAll);
        var btnEdit = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.edit));
        btnEdit.helpTip = getLabel(LABELS.tooltip.edit);
        var btnClose = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.close), { name: "cancel" });
        var btnSelect = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.select), { name: "ok" });
        btnSelect.helpTip = getLabel(LABELS.tooltip.select);
        listDialog.defaultElement = btnSelect;
        listDialog.cancelElement = btnClose;

        var changes = null;

        /**
         * 選んでいる件数に合わせてボタンを切り替える（［選択］は1件のときだけ）
         * @returns {void}
         */
        function updateButtons() {
            var selectedEntries = getSelectedEntries(qrList, entries);
            btnSelect.enabled = (selectedEntries.length === 1);
            btnEdit.enabled = (getEditableEntries(selectedEntries).length > 0);
        }

        qrList.onChange = updateButtons;
        qrList.onDoubleClick = function () {
            if (getSelectedEntries(qrList, entries).length === 1) listDialog.close(1);
        };
        btnCopyAll.onClick = function () {
            copyToClipboard(buildCopyText(entries));
        };
        btnEdit.onClick = function () {
            changes = showEditDialog(getEditableEntries(getSelectedEntries(qrList, entries)));
            if (changes) listDialog.close(3);
        };
        btnSelect.onClick = function () {
            listDialog.close(1);
        };
        btnClose.onClick = function () {
            listDialog.close(2);
        };

        updateButtons();

        /* モーダルのダイアログボックスを開いたままでは選択も QR コードも変えられないため、閉じてから処理する / Neither the selection nor a QR code can change while the modal dialog box is open */
        addVersionFooter(listDialog);
        var closeResult = listDialog.show();
        if (closeResult === 3) return { changes: changes };
        if (closeResult !== 1) return null;
        return { chosenEntry: getSelectedEntries(qrList, entries)[0] };
    }

    // =========================================
    // 実行 / Main
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }
    var activeDoc = app.activeDocument;
    var selectedFrameIds = [];

    /* 編集したら読み直して一覧に戻る / After an edit, read the document again and return to the list */
    while (true) {
        var qrEntries = collectQRCodes(activeDoc);
        if (qrEntries.length === 0) {
            alert(getLabel(LABELS.alert.noQRCode));
            return;
        }
        var listResult = showListDialog(qrEntries, selectedFrameIds);
        if (!listResult) return;
        if (listResult.chosenEntry) {
            selectQRCode(listResult.chosenEntry);
            return;
        }
        var applyResult = applyChanges(activeDoc, listResult.changes);
        if (applyResult.failedCount > 0) {
            alert(getLabel(LABELS.alert.editFailed, [applyResult.failedCount, applyResult.lastMessage]));
        }
        selectedFrameIds = [];
        for (var i = 0; i < listResult.changes.length; i++) {
            selectedFrameIds.push(listResult.changes[i].entry.frame.id);
        }
    }

}());
