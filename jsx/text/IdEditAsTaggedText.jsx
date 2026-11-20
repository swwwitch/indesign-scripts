#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdEditAsTaggedText
Meta-Default-language=ja
Name=選択したテキストをタグ付きテキストとして編集
Name[en]=Edit the selected text as tagged text
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.0.0
Release-Date=2026-11-21
Tags=text,テキスト
X-com.swwwitch-Script-Name=IdEditAsTaggedText
X-com.swwwitch-Updated=2026-11-21
X-com.swwwitch-README-JA=https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdEditAsTaggedText.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdEditAsTaggedText.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdEditAsTaggedText.jsx
Description-BEGIN
[ja]
選択したテキストをタグ付きテキストに書き出してダイアログボックスで編集し、編集結果を同じ位置に読み込み直します。
[en]
Exports the selected text as tagged text, lets you edit it in a dialog box, and reads the result back in place.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読んで SCRIPT_* に入れる（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; it is read into SCRIPT_* at run time */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME           = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名 / script name */
var SCRIPT_VERSION        = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : ""; /* バージョン / version */
var SCRIPT_AUTHOR         = getScriptMetaValue("Author");                       /* 作者 / author */
var SCRIPT_RELEASED       = getScriptMetaValue("Release-Date");                 /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED        = getScriptMetaValue("X-com.swwwitch-Updated");       /* 更新日 / last updated */
var SCRIPT_DESCRIPTION_JA = getScriptMetaValue("Name");                         /* README の見出し / README title (Japanese) */
var SCRIPT_README_JA      = getScriptMetaValue("X-com.swwwitch-README-JA");     /* README（日本語） */
var SCRIPT_README_EN      = getScriptMetaValue("X-com.swwwitch-README-EN");     /* README (English) */
var SCRIPT_ARTICLE_URL    = getScriptMetaValue("X-com.swwwitch-Article-URL");   /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL     = getScriptMetaValue("X-com.swwwitch-Source-URL");    /* 最新版のソース / latest source */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

/* タグの書き方。TagTextForm.VERBOSE（<cSize:> など）または TagTextForm.ABBREVIATED（<cs:> など）
   Tag form: TagTextForm.VERBOSE (<cSize:> etc.) or TagTextForm.ABBREVIATED (<cs:> etc.) */
var TAG_FORM = TagTextForm.VERBOSE;

// =========================================
// レイアウト設定 / Layout settings
// =========================================

/* 入力欄のサイズ [幅, 高さ]（px） / Size of the input field [width, height] (px) */
var INPUT_BOX_SIZE = [560, 320];

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
/* 左下のスクリプト名・バージョンをクリックすると開く。SCRIPTMETA から読んだ SCRIPT_*（スクリプト情報の部品）を読み、
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
var VERSION_INFO_FILE_VIEWER_APP_PATH = "/Applications/OpenInFileViewer.app";   /* Path Finder が起動中なら Path Finder、そうでなければ Finder で選択表示する補助アプリ / helper app that reveals a file in Path Finder when it is running, otherwise in the Finder */
var VERSION_INFO_FILE_VIEWER_PATH_FILE = "/tmp/open_in_file_viewer_path.txt"; /* OpenInFileViewer.app が読むパス / path the helper reads */

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
 * SCRIPT_* の値を返す（変数が無い・文字列でなければ空文字）。typeof で読むので、未定義でも例外にならない
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
 * ファイルを選択した状態で表示する。OpenInFileViewer.app があれば Path Finder が起動中なら Path Finder、そうでなければ Finder で表示し、
 * 無ければ RevealInFinder.app で Finder に表示する。どちらも無いときは入っているフォルダーを開く
 * @param {File} scriptFile - 表示するファイル
 * @returns {void}
 */
function revealVersionInfoScriptFile(scriptFile) {
    if (launchVersionInfoHelperApp(VERSION_INFO_FILE_VIEWER_APP_PATH, VERSION_INFO_FILE_VIEWER_PATH_FILE, scriptFile.fsName)) return;
    if (launchVersionInfoHelperApp(VERSION_INFO_REVEAL_APP_PATH, VERSION_INFO_REVEAL_PATH_FILE, scriptFile.fsName)) return;
    /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
    scriptFile.parent.execute();
}

/**
 * 補助アプリが読む一時ファイルにパスを書き出してから、アプリを起動する
 * @param {string} appPath - 補助アプリ（.app）のパス
 * @param {string} pathFilePath - アプリが読む一時ファイルのパス
 * @param {string} targetPath - 表示するファイルの絶対パス
 * @returns {boolean} 起動できたら true。Mac 以外・アプリが無い・起動できないときは false
 */
function launchVersionInfoHelperApp(appPath, pathFilePath, targetPath) {
    /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
    if (File.fs !== "Macintosh" || !(new Folder(appPath).exists || new File(appPath).exists)) return false;
    return writeVersionInfoRevealPath(pathFilePath, targetPath) && new File(appPath).execute();
}

/**
 * 補助アプリに渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
 * @param {string} pathFilePath - 一時ファイルのパス
 * @param {string} targetPath - 絶対パス
 * @returns {boolean} 書き出せたら true
 */
function writeVersionInfoRevealPath(pathFilePath, targetPath) {
    var pathFile = new File(pathFilePath);
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
            title: { ja: "タグ付きテキストとして編集", en: "Edit as Tagged Text" },
            note:  { ja: "改行で段落を区切ります。強制改行は <0x000A>、文字の < > \\ は \\< \\> \\\\ です。", en: "Each line is a paragraph. <0x000A> is a forced line break; type < > \\ as \\< \\> \\\\." }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            reset:  { ja: "リセット", en: "Reset" }
        },
        tooltip: {
            reset: { ja: "編集前のタグ付きテキストに戻します。", en: "Restores the tagged text as it was before editing." }
        },
        alert: {
            noDocument:   { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noTextTarget: { ja: "テキストまたはテキストフレームを1つ選択してください。", en: "Please select text or a single text frame." },
            hasObjects:   { ja: "表・アンカー付きオブジェクト・脚注などを含むテキストは編集できません。", en: "Text that contains tables, anchored objects, footnotes, etc. cannot be edited." },
            placeFailed:  { ja: "タグ付きテキストを読み込めなかったので、元のテキストに戻しました。\n", en: "The tagged text could not be imported, so the original text was restored.\n" }
        },
        undo: {
            edit: { ja: "タグ付きテキストとして編集", en: "Edit as Tagged Text" }
        }
    };

    // =========================================
    // 判定 / Detection
    // =========================================

    /* 読み直すと失われる文字（表・アンカー付きオブジェクト・脚注・ページ番号などの特殊文字）。タブ・改行・「ここまでインデント」・右インデントタブは除く
       Characters lost on re-import (tables, anchored objects, footnotes, page numbers, …); tabs, breaks, indent-to-here and right indent tabs excluded */
    var OBJECT_CHAR_PATTERN = /[\u0000-\u0006\u000B\u000C\u000E-\u001F\uFFFC]/;

    /* 書き出した本文の始まり（行頭の段落スタイルのタグ。略式は pstyle） / Start of the exported body: a paragraph style tag at a line start (pstyle in the abbreviated form) */
    var BODY_START_PATTERN = /\n<(?:ParaStyle|pstyle):/;

    /**
     * 文字を入力欄に読める形で出すか判定する（ASCII・制御文字・見えない文字は <0xXXXX> のまま）
     * @param {number} charCode - 文字コード
     * @returns {boolean} 文字で出すなら true
     */
    function isDisplayableChar(charCode) {
        if (charCode < 0x80) return false;
        if (charCode === 0x00A0 || charCode === 0x00AD || charCode === 0xFEFF) return false;
        if (charCode >= 0x2000 && charCode <= 0x200F) return false; /* 特殊なスペース・幅なし文字 / special spaces and zero-width characters */
        if (charCode >= 0x2028 && charCode <= 0x202F) return false; /* 行・段落区切り・方向制御 / line/paragraph separators and direction marks */
        if (charCode >= 0x205F && charCode <= 0x206F) return false;
        return true;
    }

    // =========================================
    // 対象 / Target
    // =========================================

    /**
     * @typedef {Object} EditTarget
     * @property {Object} container - 書き換える範囲を持つ入れ物（Story または Cell）
     * @property {Text} textRange - 書き換える範囲
     * @property {boolean} isWhole - 入れ物全体なら true
     */

    /**
     * 配列で返るプロパティの先頭を返す（itemByRange の範囲などは parent・index が配列になる）
     * @param {*} propertyValue - プロパティの値
     * @returns {*} 先頭の値、または値そのもの
     */
    function firstValue(propertyValue) {
        return (propertyValue instanceof Array) ? propertyValue[0] : propertyValue;
    }

    /**
     * 選択から編集する範囲を決める（文字の選択はその範囲、カーソル・テキストフレームはストーリー全体、セルの中ならそのセル）
     * @param {Array} selectedItems - app.selection
     * @returns {EditTarget|null} 対象。テキストでなければ null
     */
    function getEditTarget(selectedItems) {
        if (selectedItems.length !== 1) return null;
        var selectedItem = selectedItems[0];
        var itemType = selectedItem.constructor.name;

        if (itemType === "TextFrame") {
            var frameStory = selectedItem.parentStory;
            return { container: frameStory, textRange: frameStory.texts[0], isWhole: true };
        }
        if (!selectedItem.hasOwnProperty("insertionPoints") || !selectedItem.hasOwnProperty("parentStory")) return null;

        var container = firstValue(selectedItem.insertionPoints[0].parent);
        if (itemType === "InsertionPoint" || selectedItem.characters.length === 0) {
            return { container: container, textRange: container.texts[0], isWhole: true };
        }
        return { container: container, textRange: selectedItem.texts[0], isWhole: false };
    }

    // =========================================
    // 書き出し / Export
    // =========================================

    /**
     * 一時ファイルを作る
     * @param {string} suffix - ファイル名の末尾
     * @returns {File} 一時ファイル（まだ書いていない）
     */
    function createTempFile(suffix) {
        return new File(Folder.temp.fsName + "/" + SCRIPT_NAME + "_" + new Date().getTime() + "_" + suffix + ".txt");
    }

    /**
     * ファイルを ASCII で読む
     * @param {File} sourceFile - 読むファイル
     * @returns {string} 中身
     */
    function readAsciiFile(sourceFile) {
        sourceFile.encoding = "ASCII";
        sourceFile.open("r");
        var fileText = sourceFile.read();
        sourceFile.close();
        return fileText;
    }

    /**
     * 文字列を ASCII・LF で書く（CR だとタグ付きテキストと認識されない）
     * @param {File} targetFile - 書くファイル
     * @param {string} fileText - 中身
     * @returns {void}
     */
    function writeAsciiFile(targetFile, fileText) {
        targetFile.encoding = "ASCII";
        targetFile.lineFeed = "Unix";
        targetFile.open("w");
        targetFile.write(fileText);
        targetFile.close();
    }

    /**
     * 範囲をタグ付きテキスト（ASCII。ASCII 以外は <0xXXXX>）で書き出して読む。書き出しの設定は元に戻す
     * @param {Text} textRange - 書き出す範囲
     * @returns {string} ファイルの中身
     */
    function exportTaggedText(textRange) {
        var exportPreferences = app.taggedTextExportPreferences;
        var savedCharacterSet = exportPreferences.characterSet;
        var savedTagForm = exportPreferences.tagForm;
        var exportFile = createTempFile("export");
        try {
            exportPreferences.characterSet = TagTextExportCharacterSet.ASCII;
            exportPreferences.tagForm = TAG_FORM;
            textRange.exportFile(ExportFormat.TAGGED_TEXT, exportFile);
            return readAsciiFile(exportFile);
        } finally {
            exportPreferences.characterSet = savedCharacterSet;
            exportPreferences.tagForm = savedTagForm;
            if (exportFile.exists) exportFile.remove();
        }
    }

    /**
     * position の「<」で始まるタグの閉じ「>」の位置を返す（入れ子の <…> と「\」のエスケープを考慮）。
     * ExtendScript の正規表現は入れ子の量指定を誤評価するので、1文字ずつ数える
     * @param {string} text - 文字列
     * @param {number} position - タグの「<」の位置
     * @returns {number} 閉じ「>」の位置（閉じていなければ -1）
     */
    function findTagEnd(text, position) {
        var depth = 0;
        for (var i = position; i < text.length; i++) {
            var character = text.charAt(i);
            if (character === "\\") {
                i++;
            } else if (character === "<") {
                depth++;
            } else if (character === ">") {
                depth--;
                if (depth === 0) return i;
            }
        }
        return -1;
    }

    /**
     * 書き出したタグ付きテキストを、ヘッダー・定義と本文に分ける（本文は行頭の段落スタイルのタグから）
     * @param {string} taggedText - 書き出したファイルの中身
     * @returns {{prologue: string, body: string}} prologue はヘッダーと定義（編集しない）、body は本文
     */
    function splitTaggedText(taggedText) {
        var bodyStart = taggedText.search(BODY_START_PATTERN) + 1;
        if (bodyStart === 0) bodyStart = taggedText.indexOf("\n") + 1;
        return {
            prologue: taggedText.substr(0, bodyStart),
            body: taggedText.substr(bodyStart).replace(/[\r\n]+$/, "")
        };
    }

    // =========================================
    // 入力欄との変換 / Conversion for the input field
    // =========================================

    /**
     * 本文の <0xXXXX> を読める文字に戻す（「\」でエスケープした部分・制御文字・見えない文字はそのまま）
     * @param {string} bodyText - 書き出した本文（ASCII）
     * @returns {string} 入力欄に出す文字列
     */
    function decodeForDisplay(bodyText) {
        var displayParts = [];
        var position = 0;
        while (position < bodyText.length) {
            var character = bodyText.charAt(position);
            if (character === "\\") {
                displayParts.push(bodyText.substr(position, 2));
                position += 2;
                continue;
            }
            var escapeMatch = (character === "<") ? /^<0x([0-9A-Fa-f]{4})>/.exec(bodyText.substr(position, 8)) : null;
            if (escapeMatch && isDisplayableChar(parseInt(escapeMatch[1], 16))) {
                displayParts.push(String.fromCharCode(parseInt(escapeMatch[1], 16)));
                position += 8;
                continue;
            }
            displayParts.push(character);
            position++;
        }
        return displayParts.join("");
    }

    /**
     * 入力欄の文字列をファイルに書ける本文に変える（改行は LF、ASCII 以外は <0xXXXX>）
     * @param {string} displayText - 入力欄の文字列
     * @returns {string} 本文
     */
    function encodeForFile(displayText) {
        var normalizedText = displayText.replace(/\r\n?/g, "\n");
        var fileParts = [];
        for (var i = 0; i < normalizedText.length; i++) {
            var charCode = normalizedText.charCodeAt(i);
            if (charCode === 10 || charCode === 9 || (charCode >= 32 && charCode <= 126)) fileParts.push(normalizedText.charAt(i));
            else fileParts.push("<0x" + ("000" + charCode.toString(16).toUpperCase()).slice(-4) + ">");
        }
        return fileParts.join("");
    }

    /**
     * 文字より前に続くタグの中に ParaStyle があるか判定する
     * @param {string} bodyText - 本文
     * @returns {boolean} ParaStyle があれば true
     */
    function startsWithParaStyle(bodyText) {
        var position = 0;
        while (bodyText.charAt(position) === "<") {
            if (bodyText.substr(position, 11) === "<ParaStyle:" || bodyText.substr(position, 8) === "<pstyle:") return true;
            var tagEnd = findTagEnd(bodyText, position);
            if (tagEnd < 0) return false;
            position = tagEnd + 1;
        }
        return false;
    }

    /**
     * 段落スタイルを ParaStyle タグに書く名前に変える（グループは「:」で区切り、名前の中の「\」「:」「<」「>」はエスケープ）
     * @param {ParagraphStyle} paragraphStyle - 段落スタイル
     * @returns {string} タグに書く名前（［段落スタイルなし］は空、［基本段落］は NormalParagraphStyle）
     */
    function paragraphStyleTagName(paragraphStyle) {
        var nameParts = [];
        var styleNode = paragraphStyle;
        while (styleNode.constructor.name !== "Document") {
            nameParts.unshift(styleNode.name.replace(/\\/g, "\\\\").replace(/:/g, "\\:"));
            styleNode = styleNode.parent;
        }
        var rootStyles = styleNode.paragraphStyles;
        if (paragraphStyle.id === rootStyles[0].id) return "";
        if (paragraphStyle.id === rootStyles[1].id) return "NormalParagraphStyle";
        return nameParts.join(":").replace(/([\\:<>])/g, "\\$1");
    }

    /**
     * 編集した本文からファイルの中身を作る。
     * 本文の先頭に段落のタグが無いと、すぐ後の文字のタグも効かず段落が［段落スタイルなし］になるので、書き換える位置の段落スタイルを補う
     * @param {string} prologue - ヘッダーと定義
     * @param {string} displayText - 入力欄の文字列
     * @param {ParagraphStyle} hostParagraphStyle - 書き換える位置の段落スタイル
     * @returns {string} ファイルの中身
     */
    function buildTaggedFileText(prologue, displayText, hostParagraphStyle) {
        var bodyText = encodeForFile(displayText);
        if (!startsWithParaStyle(bodyText)) bodyText = "<ParaStyle:" + encodeForFile(paragraphStyleTagName(hostParagraphStyle)) + ">" + bodyText;
        return prologue + bodyText;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * タグ付きテキストの編集ダイアログボックスを表示する
     * @param {string} initialText - 入力欄の初期値
     * @returns {string|null} 編集した文字列。キャンセル時は null
     */
    function showEditDialog(initialText) {
        var editDialog = new Window("dialog", getLabel(LABELS.dialog.title));
        setupWindow(editDialog, 8);

        var textInput = editDialog.add("edittext", undefined, initialText, { multiline: true, scrolling: true });
        textInput.preferredSize = INPUT_BOX_SIZE;
        textInput.alignment = ["fill", "fill"];

        var noteText = editDialog.add("statictext", undefined, getLabel(LABELS.dialog.note));
        noteText.alignment = ["left", "center"];

        /* ボタンエリア（左右分割） / Button area (split left and right) */
        var buttonRow = addButtonRow(editDialog);
        var btnReset = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.reset));
        btnReset.helpTip = getLabel(LABELS.tooltip.reset);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        btnReset.onClick = function () {
            textInput.text = initialText;
            textInput.active = true;
        };

        textInput.active = true;
        addVersionFooter(editDialog);
        return (editDialog.show() === 1) ? textInput.text : null;
    }

    // =========================================
    // 読み込み / Import
    // =========================================

    /**
     * 読み込みの設定を変換向けにし、元の値を返す（引用符を勝手に変えない・書式を残す・ダイアログボックスを出さない）
     * @returns {Object} 元の設定
     */
    function prepareImportPreferences() {
        var importPreferences = app.taggedTextImportPreferences;
        var savedPreferences = {
            useTypographersQuotes: importPreferences.useTypographersQuotes,
            removeTextFormatting:  importPreferences.removeTextFormatting,
            userInteractionLevel:  app.scriptPreferences.userInteractionLevel
        };
        importPreferences.useTypographersQuotes = false;
        importPreferences.removeTextFormatting = false;
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
        return savedPreferences;
    }

    /**
     * 読み込みの設定を元に戻す
     * @param {Object} savedPreferences - prepareImportPreferences が返した値
     * @returns {void}
     */
    function restoreImportPreferences(savedPreferences) {
        var importPreferences = app.taggedTextImportPreferences;
        importPreferences.useTypographersQuotes = savedPreferences.useTypographersQuotes;
        importPreferences.removeTextFormatting = savedPreferences.removeTextFormatting;
        app.scriptPreferences.userInteractionLevel = savedPreferences.userInteractionLevel;
    }

    /**
     * 範囲を消し、同じ位置に編集したタグ付きテキストを読み込む。
     * 読み込みに失敗したときは、書き出したままのファイルを読み込んで元に戻してから例外を投げる
     * @param {EditTarget} target - 対象
     * @param {string} editedFileText - 編集したファイルの中身
     * @param {string} originalFileText - 書き出したままのファイルの中身
     * @returns {void}
     */
    function replaceWithTaggedText(target, editedFileText, originalFileText) {
        var editedFile = createTempFile("edited");
        var savedPreferences = prepareImportPreferences();
        try {
            writeAsciiFile(editedFile, editedFileText);
            var insertIndex = target.isWhole ? 0 : firstValue(target.textRange.insertionPoints[0].index);
            target.textRange.remove();
            try {
                target.container.insertionPoints[insertIndex].place(editedFile);
            } catch (placeError) {
                writeAsciiFile(editedFile, originalFileText);
                target.container.insertionPoints[insertIndex].place(editedFile);
                throw placeError;
            }
        } finally {
            restoreImportPreferences(savedPreferences);
            if (editedFile.exists) editedFile.remove();
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択したテキストをタグ付きテキストとして編集する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var target = getEditTarget(app.selection);
        if (!target) {
            alert(getLabel(LABELS.alert.noTextTarget));
            return;
        }
        if (OBJECT_CHAR_PATTERN.test(String(target.textRange.contents))) {
            alert(getLabel(LABELS.alert.hasObjects));
            return;
        }

        var hostParagraphStyle = firstValue(target.textRange.insertionPoints[0].appliedParagraphStyle);
        var originalFileText = exportTaggedText(target.textRange);
        var taggedParts = splitTaggedText(originalFileText);
        var initialText = decodeForDisplay(taggedParts.body);

        var editedText = showEditDialog(initialText);
        if (editedText === null) return;
        editedText = editedText.replace(/\r\n?/g, "\n");
        if (editedText === initialText) return;

        var editedFileText = buildTaggedFileText(taggedParts.prologue, editedText, hostParagraphStyle);
        try {
            app.doScript(function () {
                replaceWithTaggedText(target, editedFileText, originalFileText);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undo.edit));
        } catch (e) {
            alert(getLabel(LABELS.alert.placeFailed) + e);
        }
    }

    main();

})();
