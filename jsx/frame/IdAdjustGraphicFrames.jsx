#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdAdjustGraphicFrames
Meta-Default-language=ja
Name=アンカー付きグラフィックフレームの幅・サイズ・縮尺を一括調整
Name[en]=Batch-adjust anchored graphic frames
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.0.6
Release-Date=2026-06-02
Tags=frame,フレーム
X-com.swwwitch-Script-Name=IdAdjustGraphicFrames
X-com.swwwitch-Updated=2026-11-20
X-com.swwwitch-README-JA=https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAdjustGraphicFrames.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAdjustGraphicFrames.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/indesign-scripts/blob/main/jsx/frame/IdAdjustGraphicFrames.jsx
Description-BEGIN
[ja]
テキストにアンカーされたグラフィックフレームを集め、フレーム幅・フレームサイズ・画像の縮尺率をまとめて調整します。
[en]
Collects graphic frames anchored in text and adjusts their frame width, frame size and image scale together.
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

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 縮尺率の切り捨て単位の選択肢（％）と既定値 / Round-down step choices (%) and the default */
    var ROUND_PRECISION_OPTIONS = [1, 5, 10];
    var DEFAULT_ROUND_PRECISION = 5;

    /* 既定で「切り捨てる」を ON にする / Turn on "Round Down" by default */
    var DEFAULT_ROUND_SCALE = true;

    /* 切り捨て対象を元解像度 72/96/144 ppi の画像に限定する（既定 ON）
       / Limit round-down to images whose actual resolution is 72, 96 or 144 ppi (on by default) */
    var DEFAULT_ROUND_ONLY_72_96_144 = true;
    var ROUND_TARGET_PPI = [72, 96, 144];

    /* 縦横比の崩れ（横スケール≠縦スケール）を同率とみなす許容差（％）
       / Tolerance (%) for treating horizontal and vertical scales as already uniform */
    var ASPECT_RATIO_TOLERANCE = 0.1;

    /* 既定で「フレームを内容に合わせる」を ON にする / Turn on "Fit Frame to Content" by default */
    var DEFAULT_FIT_FRAME_TO_CONTENT = true;

    /* 既定で「インライン画像を文字サイズに合わせる」を ON にする / Turn on inline-image matching by default */
    var DEFAULT_ADJUST_INLINE_IMAGE_IN_TEXT = true;

    /* 既定で「合わせ先より大きい場合のみ」を ON にする / Turn on "Only when wider than target" by default */
    var DEFAULT_WIDTH_ONLY_IF_LARGER = true;

    /* 既定で切り捨て後の「再フィット」を ON にする / Re-fit after rounding down by default */
    var DEFAULT_REFIT_AFTER_ROUND = true;

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

    var LABELS = {
        dialog: {
            title: { ja: "アンカー付きオブジェクトのフレーム調整", en: "Adjust Anchored Object Frames" }
        },
        panel: {
            target: { ja: "対象", en: "Target" },
            width: { ja: "フレーム幅", en: "Frame Width" },
            frameSize: { ja: "フレームサイズ", en: "Frame Size" },
            scale: { ja: "縮尺率", en: "Scale" }
        },
        target: {
            document: { ja: "ドキュメント", en: "Document" },
            story: { ja: "ストーリー", en: "Story" },
            selection: { ja: "選択範囲", en: "Selection" }
        },
        fit: {
            frameToContent: { ja: "フレームを内容に合わせる", en: "Fit Frame to Content" },
            adjustInlineImageInText: { ja: "インライン画像を文字サイズに合わせる", en: "Match Inline Images to Text Size" }
        },
        width: {
            keep: { ja: "変更しない", en: "Don't Change" },
            fitToParent: { ja: "親フレームに合わせる", en: "Fit to Parent Frame" },
            onlyIfLarger: { ja: "親フレーム／マージンより大きい場合のみ", en: "Only When Wider Than Target" },
            fitToMargin: { ja: "マージンに合わせる", en: "Fit to Margins" }
        },
        scale: {
            round: { ja: "縮尺率を切り捨てる", en: "Round Scale Down" },
            precision: { ja: "単位", en: "Step" },
            refit: { ja: "調整後にフレームを内容へ合わせる", en: "Fit Frame to Content After Adjusting" },
            only7296144: { ja: "スクショのみ（72／96／144 ppi の画像）", en: "Screenshots only (72/96/144 ppi images)" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tip: {
            targetDocument: { ja: "ドキュメント内の、テキストにアンカーされたすべてのグラフィックフレーム", en: "All text-anchored graphic frames in the document" },
            targetStory: { ja: "選択中のテキストフレーム／テキスト範囲／挿入点が属するストーリーにアンカーされたグラフィックフレーム", en: "Graphic frames anchored in the story of the selected text frame / text range / insertion point" },
            targetSelection: { ja: "選択中の、テキストにアンカーされたグラフィックフレーム（グループは展開。テキスト選択時はそのストーリーを対象）", en: "Selected text-anchored graphic frames (groups expanded; a text selection targets its story)" },
            fitFrameToContent: { ja: "フレームの大きさを、配置された内容のバウンディングボックスに合わせます。", en: "Resize the frame to match the placed content's bounds." },
            adjustInlineImageInText: { ja: "真のインライン配置の画像の高さを、同じ段落内の文字サイズに合わせます（行揃え／カスタム配置は対象外）。", en: "Match true inline image height to the text size in the same paragraph (above-line / custom positions excluded)." },
            widthKeep: { ja: "フレームの幅は変更しません。", en: "Leave the frame width unchanged." },
            fitToParent: { ja: "アンカーされたフレームの幅を、親テキストフレームの内寸に合わせます。", en: "Match the anchored frame width to the parent text frame's inner width." },
            onlyIfLarger: { ja: "オンにすると、合わせ先（親フレーム／マージン）より広いフレームだけを調整します（縮小のみ）。オフの場合は、小さいフレームも合わせ先の幅に広げます。", en: "When on, only frames wider than the target (parent frame / margins) are adjusted (shrink only). When off, smaller frames are also expanded to the target width." },
            fitToMargin: { ja: "フレームをページのマージン（版面）幅に合わせて配置します。", en: "Fit frames to the page margin live-area width." },
            round: { ja: "画像の拡大縮小率を、指定した単位で切り捨てます。", en: "Round the image scale down to the selected step." },
            precision: { ja: "切り捨ての刻み幅です。小さいほど元の倍率に近づきます。", en: "Round-down step. Smaller values stay closer to the original scale." },
            refit: { ja: "縮尺率を切り捨てた後、フレームを内容に合わせ直します。", en: "After rounding the scale down, fit the frame to the content again." },
            only7296144: { ja: "元解像度が 72／96／144 ppi の画像だけを切り捨て対象にします。", en: "Limit round-down to images whose actual resolution is 72, 96, or 144 ppi." }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noTextFrame: { ja: "おっと!!!\nテキストフレーム、テキスト範囲、挿入点、またはアンカーオブジェクトマーカーを選択してください。", en: "Oops!!!\nPlease select a text frame, text range, insertion point, or anchored object marker." },
            noSelection: { ja: "フレームを選択してください。", en: "Please select one or more frames." },
            noFrames: { ja: "対象となるフレームが見つかりませんでした。", en: "No target frames were found." },
            done: { ja: "{count} 個のフレームを調整しました。", en: "Adjusted {count} frame(s)." },
            doneNone: { ja: "調整対象のフレームはありませんでした。", en: "No frames needed adjustment." }
        }
    };

    main();

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * 検証・ダイアログ・収集・適用の流れを制御する
     * @returns {void}
     */
    function main() {
        /* ドキュメントの有無を確認 / Require an open document */
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var targetDocument = app.activeDocument;

        /* 選択中のテキストフレームを取得（初期選択の判定用）/ Get the selected text frame (for default state) */
        var selectedTextFrame = getSelectedTextFrame();

        /* ダイアログで設定を取得 / Get settings from dialog */
        var settings = showOptionsDialog(selectedTextFrame !== null);
        if (settings === null) return; // キャンセル / Cancelled

        /* 対象に応じてフレームを収集 / Collect frames by target */
        var targetFrames = collectFrames(targetDocument, settings.target, selectedTextFrame);
        if (targetFrames === null) return; // 選択不足などで中断 / Aborted (e.g. nothing selected)
        if (targetFrames.length === 0) {
            alert(getLabel("alert.noFrames"));
            return;
        }

        /* 取り消しをひとまとめに / Apply in a single undo step */
        var processedCount = 0;
        app.doScript(
            function () { processedCount = applyToFrames(targetFrames, settings); },
            ScriptLanguage.JAVASCRIPT,
            undefined,
            UndoModes.ENTIRE_SCRIPT,
            getLabel("dialog.title")
        );

        /* 完了メッセージ（0 件は専用文言）/ Completion message (dedicated text for zero) */
        alert(processedCount > 0 ? getLabel("alert.done", { count: processedCount }) : getLabel("alert.doneNone"));
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 調整内容を指定するダイアログを表示する
     * @param {boolean} hasTextFrameSelection テキストフレームが選択されているか
     * @returns {object|null} 設定内容。キャンセル時は null
     */
    function showOptionsDialog(hasTextFrameSelection) {
        var dialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(dialog, 10);

        /* 対象はカラム貫通（全幅）/ Target spans the full width */
        var targetControls = buildTargetPanel(dialog, hasTextFrameSelection);

        /* 1 カラム：フレームの幅 → フレームの大きさ → 縮尺率 / One column: Width → Frame size → Scale */
        var widthControls = buildWidthPanel(dialog);
        var frameToContentControls = buildFrameToContentPanel(dialog);
        var scaleControls = buildScalePanel(dialog);

        /* ボタン / Buttons (Mac: Cancel → OK) */
        var buttonRow = addButtonRow(dialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        addVersionFooter(dialog);
        if (dialog.show() !== 1) return null;

        return collectSettings(targetControls, frameToContentControls, widthControls, scaleControls);
    }

    /**
     * 対象範囲のパネルを組み立てる
     * @param {Window} dialog 対象のダイアログ
     * @param {boolean} hasTextFrameSelection テキストフレームが選択されているか
     * @returns {object} パネル内のコントロール
     */
    function buildTargetPanel(parent, hasTextFrameSelection) {
        var panel = parent.add("panel", undefined, getLabel("panel.target"));
        setupPanel(panel);

        /* ラジオを横並び＋左右中央に / Radios in a row, centered horizontally */
        var radioGroup = panel.add("group");
        setupRow(radioGroup, "center");

        var documentRadio = radioGroup.add("radiobutton", undefined, getLabel("target.document"));
        var storyRadio = radioGroup.add("radiobutton", undefined, getLabel("target.story"));
        var selectionRadio = radioGroup.add("radiobutton", undefined, getLabel("target.selection"));
        documentRadio.helpTip = getLabel("tip.targetDocument");
        storyRadio.helpTip = getLabel("tip.targetStory");
        selectionRadio.helpTip = getLabel("tip.targetSelection");

        /* 初期選択は常にドキュメント / Default is always Document */
        documentRadio.value = true;
        /* ストーリーはテキストフレーム選択時のみ選択可能 / Story is selectable only when a text frame is selected */
        storyRadio.enabled = hasTextFrameSelection;
        /* 選択範囲は何か選択しているときのみ選択可能 / Selection is selectable only when something is selected */
        selectionRadio.enabled = (app.selection.length > 0);

        return { documentRadio: documentRadio, storyRadio: storyRadio, selectionRadio: selectionRadio };
    }

    /**
     * フレームサイズのパネルを組み立てる
     * @param {Window} dialog 対象のダイアログ
     * @returns {object} パネル内のコントロール
     */
    function buildFrameToContentPanel(parent) {
        var panel = parent.add("panel", undefined, getLabel("panel.frameSize"));
        setupPanel(panel, 6);

        var frameToContentCheckbox = panel.add("checkbox", undefined, getLabel("fit.frameToContent"));
        frameToContentCheckbox.value = DEFAULT_FIT_FRAME_TO_CONTENT;
        frameToContentCheckbox.helpTip = getLabel("tip.fitFrameToContent");

        var adjustInlineImageInTextCheckbox = panel.add("checkbox", undefined, getLabel("fit.adjustInlineImageInText"));
        adjustInlineImageInTextCheckbox.value = DEFAULT_ADJUST_INLINE_IMAGE_IN_TEXT;
        adjustInlineImageInTextCheckbox.helpTip = getLabel("tip.adjustInlineImageInText");

        return {
            frameToContentCheckbox: frameToContentCheckbox,
            adjustInlineImageInTextCheckbox: adjustInlineImageInTextCheckbox
        };
    }

    /**
     * フレーム幅のパネルを組み立てる
     * @param {Window} dialog 対象のダイアログ
     * @returns {object} パネル内のコントロール
     */
    function buildWidthPanel(parent) {
        var panel = parent.add("panel", undefined, getLabel("panel.width"));
        setupPanel(panel, 6);

        var keepRadio = panel.add("radiobutton", undefined, getLabel("width.keep"));
        var fitToParentRadio = panel.add("radiobutton", undefined, getLabel("width.fitToParent"));
        var fitToMarginRadio = panel.add("radiobutton", undefined, getLabel("width.fitToMargin"));
        var onlyIfLargerCheckbox = panel.add("checkbox", undefined, getLabel("width.onlyIfLarger"));
        fitToParentRadio.value = true;
        onlyIfLargerCheckbox.value = DEFAULT_WIDTH_ONLY_IF_LARGER;
        keepRadio.helpTip = getLabel("tip.widthKeep");
        fitToParentRadio.helpTip = getLabel("tip.fitToParent");
        fitToMarginRadio.helpTip = getLabel("tip.fitToMargin");
        onlyIfLargerCheckbox.helpTip = getLabel("tip.onlyIfLarger");

        /**
         * 幅の指定に応じて「合わせ先より大きい場合のみ」の有効／無効を切り替える
         * @returns {void}
         */
        function syncOnlyIfLargerEnabled() {
            onlyIfLargerCheckbox.enabled = fitToParentRadio.value || fitToMarginRadio.value;
        }
        keepRadio.onClick = syncOnlyIfLargerEnabled;
        fitToParentRadio.onClick = syncOnlyIfLargerEnabled;
        fitToMarginRadio.onClick = syncOnlyIfLargerEnabled;
        syncOnlyIfLargerEnabled();

        return {
            keepRadio: keepRadio,
            fitToParentRadio: fitToParentRadio,
            fitToMarginRadio: fitToMarginRadio,
            onlyIfLargerCheckbox: onlyIfLargerCheckbox
        };
    }

    /**
     * 縮尺率のパネルを組み立てる
     * @param {Window} dialog 対象のダイアログ
     * @returns {object} パネル内のコントロール
     */
    function buildScalePanel(parent) {
        var panel = parent.add("panel", undefined, getLabel("panel.scale"));
        setupPanel(panel, 6);

        var roundCheckbox = panel.add("checkbox", undefined, getLabel("scale.round"));
        roundCheckbox.value = DEFAULT_ROUND_SCALE;
        roundCheckbox.helpTip = getLabel("tip.round");

        var precisionGroup = panel.add("group");
        setupRow(precisionGroup);
        /* コロンは日本語は全角、英語は半角 / Colon: full-width JA, half-width EN */
        var precisionLabel = precisionGroup.add("statictext", undefined, labelText("scale.precision"));
        precisionLabel.helpTip = getLabel("tip.precision");

        var precisionRadios = [];
        for (var i = 0; i < ROUND_PRECISION_OPTIONS.length; i++) {
            var radio = precisionGroup.add("radiobutton", undefined, ROUND_PRECISION_OPTIONS[i] + "%");
            if (ROUND_PRECISION_OPTIONS[i] === DEFAULT_ROUND_PRECISION) radio.value = true;
            radio.helpTip = getLabel("tip.precision");
            precisionRadios.push(radio);
        }

        var refitCheckbox = panel.add("checkbox", undefined, getLabel("scale.refit"));
        refitCheckbox.value = DEFAULT_REFIT_AFTER_ROUND;
        refitCheckbox.helpTip = getLabel("tip.refit");

        var only7296144Checkbox = panel.add("checkbox", undefined, getLabel("scale.only7296144"));
        only7296144Checkbox.value = DEFAULT_ROUND_ONLY_72_96_144;
        only7296144Checkbox.helpTip = getLabel("tip.only7296144");

        /**
         * 切り捨ての ON/OFF に応じて関連項目を切り替える
         * @returns {void}
         */
        function syncRoundEnabled() {
            precisionGroup.enabled = roundCheckbox.value;
            refitCheckbox.enabled = roundCheckbox.value;
            only7296144Checkbox.enabled = roundCheckbox.value;
        }
        roundCheckbox.onClick = syncRoundEnabled;
        syncRoundEnabled();

        return {
            roundCheckbox: roundCheckbox,
            precisionRadios: precisionRadios,
            refitCheckbox: refitCheckbox,
            only7296144Checkbox: only7296144Checkbox
        };
    }

    /**
     * ダイアログの入力内容を設定オブジェクトにまとめる
     * @returns {object} 適用に使う設定
     */
    function collectSettings(targetControls, frameToContentControls, widthControls, scaleControls) {
        /* 対象 / Target */
        var target = "document";
        if (targetControls.storyRadio.value) target = "story";
        else if (targetControls.selectionRadio.value) target = "selection";

        /* 丸め精度 / Rounding precision */
        var roundPrecision = DEFAULT_ROUND_PRECISION;
        for (var j = 0; j < scaleControls.precisionRadios.length; j++) {
            if (scaleControls.precisionRadios[j].value) {
                roundPrecision = ROUND_PRECISION_OPTIONS[j];
                break;
            }
        }

        /* フレーム幅の調整モード / Frame width mode */
        var widthMode = "keep";
        if (widthControls.fitToParentRadio.value) widthMode = "parent";
        else if (widthControls.fitToMarginRadio.value) widthMode = "margin";

        return {
            target: target,
            fitFrameToContent: frameToContentControls.frameToContentCheckbox.value,
            adjustInlineImageInText: frameToContentControls.adjustInlineImageInTextCheckbox.value,
            widthMode: widthMode,
            widthOnlyIfLarger: widthControls.onlyIfLargerCheckbox.value,
            roundScale: scaleControls.roundCheckbox.value,
            roundPrecision: roundPrecision,
            refitAfterRound: scaleControls.refitCheckbox.value,
            roundOnly7296144: scaleControls.only7296144Checkbox.value
        };
    }

    /**
     * 選択からテキストフレームを取り出す
     * @returns {TextFrame|null} テキストフレーム。取得できない場合は null
     */
    function getSelectedTextFrame() {
        if (app.selection.length === 0) return null;
        var selectedItem = app.selection[0];
        try {
            if (selectedItem.constructor.name === "TextFrame") return selectedItem;
            if (selectedItem.parentTextFrames && selectedItem.parentTextFrames.length > 0) return selectedItem.parentTextFrames[0];
            if (selectedItem.parent && selectedItem.parent.parentTextFrames && selectedItem.parent.parentTextFrames.length > 0) return selectedItem.parent.parentTextFrames[0];
            /* ^~a$（アンカーオブジェクトマーカーのみ選択）にも対応
               Support selections consisting only of an anchored-object marker */
            if (selectedItem.constructor.name === "Character" && selectedItem.contents === "\uFFFC") {
                if (selectedItem.parentTextFrames && selectedItem.parentTextFrames.length > 0) {
                    return selectedItem.parentTextFrames[0];
                }
            }
        } catch (e) { }
        return null;
    }

    // =========================================
    // フレーム収集 / Frame collection
    // =========================================

    /**
     * 対象範囲に応じて調整するフレームを集める
     * @param {Document} targetDocument 対象ドキュメント
     * @param {string} target 対象範囲を表す識別子
     * @param {TextFrame} selectedTextFrame 選択中のテキストフレーム
     * @returns {Array<PageItem>|null} 対象フレーム。中断時は null
     */
    function collectFrames(targetDocument, target, selectedTextFrame) {
        if (target === "story") {
            if (selectedTextFrame === null) {
                alert(getLabel("alert.noTextFrame"));
                return null;
            }
            return collectFramesFromStory(targetDocument, selectedTextFrame.parentStory);
        }
        if (target === "selection") {
            if (app.selection.length === 0) {
                alert(getLabel("alert.noSelection"));
                return null;
            }
            return collectFramesFromSelection(app.selection);
        }
        return collectFramesFromDocument(targetDocument);
    }

    /**
     * 編集できるフレームかどうかを判定する
     * @param {PageItem} item 対象のオブジェクト
     * @returns {boolean} 編集できるなら true
     */
    function isEditableFrame(item) {
        if (!isGraphicFrameType(item)) return false;
        /* テキストにアンカーされたフレームのみ対象（独立配置は対象外）/ Only frames anchored into text (free-floating frames are excluded) */
        if (getOwningStory(item) === null) return false;
        try {
            /* マスターページ上の項目は対象外 / Skip items on master spreads */
            var page = item.parentPage;
            if (page !== null && page.parent.constructor.name === "MasterSpread") return false;
            /* 非表示・ロックは対象外 / Skip hidden or locked items */
            if (item.visible === false) return false;
            if (item.locked === true) return false;
            /* ロックレイヤー・非表示レイヤーは対象外 / Skip items on locked or hidden layers */
            var itemLayer = item.itemLayer;
            if (itemLayer !== undefined && (itemLayer.locked === true || itemLayer.visible === false)) return false;
        } catch (e) {
            return false;
        }
        return true;
    }

    /**
     * グラフィックフレーム（長方形・楕円・多角形）の種別かどうか
     * @param {PageItem} pageItem 判定するオブジェクト
     * @returns {boolean} グラフィックフレームなら true
     */
    function isGraphicFrameType(pageItem) {
        var typeName = pageItem.constructor.name;
        return typeName === "Rectangle" || typeName === "Oval" || typeName === "Polygon";
    }

    /**
     * オブジェクトが属するストーリーを取得する
     * @param {PageItem} pageItem 対象のオブジェクト
     * @returns {Story|null} ストーリー。取得できない場合は null
     */
    function getOwningStory(item) {
        try {
            /* storyOffset から親ストーリーを取得 / Read parent story from storyOffset */
            var anchorPoint = item.storyOffset;
            if (anchorPoint && anchorPoint.isValid) {
                if (anchorPoint.parentStory && anchorPoint.parentStory.isValid) return anchorPoint.parentStory;
                if (anchorPoint.parent && anchorPoint.parent.parentStory && anchorPoint.parent.parentStory.isValid) return anchorPoint.parent.parentStory;
            }
        } catch (e) { }

        try {
            /* インラインアンカーでは parent が Character になることがある / Inline anchored items may have a Character parent */
            var parent = item.parent;
            if (parent && parent.parentStory && parent.parentStory.isValid) return parent.parentStory;
            if (parent && parent.parent && parent.parent.parentStory && parent.parent.parentStory.isValid) return parent.parent.parentStory;
        } catch (err) { }

        return null;
    }

    /**
     * ストーリーにアンカーされたフレームを集める
     * @param {Story} story 対象のストーリー
     * @returns {Array<PageItem>} 対象フレーム
     */
    function collectFramesFromStory(targetDocument, story) {
        var allItems = targetDocument.allPageItems;
        var collectedFrames = [];
        for (var i = 0; i < allItems.length; i++) {
            if (!isEditableFrame(allItems[i])) continue;
            var owningStory = getOwningStory(allItems[i]);
            if (owningStory !== null && owningStory.id === story.id) collectedFrames.push(allItems[i]);
        }
        return collectedFrames;
    }

    /**
     * ドキュメント全体からアンカーされたフレームを集める
     * @param {Document} targetDocument 対象ドキュメント
     * @returns {Array<PageItem>} 対象フレーム
     */
    function collectFramesFromDocument(targetDocument) {
        var allItems = targetDocument.allPageItems;
        var collectedFrames = [];
        for (var i = 0; i < allItems.length; i++) {
            if (isEditableFrame(allItems[i])) collectedFrames.push(allItems[i]);
        }
        return collectedFrames;
    }

    /**
     * 選択からアンカーされたフレームを集める
     * @returns {Array<PageItem>|null} 対象フレーム。中断時は null
     */
    function collectFramesFromSelection(selection) {
        var collectedFrames = [];
        var seenIds = {};
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            var typeName = item.constructor.name;

            if (typeName === "TextFrame") {
                appendFrames(collectedFrames, seenIds, collectFramesFromStory(app.activeDocument, item.parentStory));
            } else if (typeName === "Character" && item.contents === "\uFFFC") {
                appendFrames(collectedFrames, seenIds, collectFramesFromStory(app.activeDocument, item.parentStory));
            } else if (typeName === "Group") {
                appendFrames(collectedFrames, seenIds, collectFramesFromGroup(item));
            } else if (isEditableFrame(item)) {
                appendFrame(collectedFrames, seenIds, item);
            } else {
                /* テキスト範囲・挿入点などのテキスト選択は、そのストーリーを対象に
                   Text range / insertion point → collect from its story */
                var selectionStory = getSelectionStory(item);
                if (selectionStory !== null) {
                    appendFrames(collectedFrames, seenIds, collectFramesFromStory(app.activeDocument, selectionStory));
                }
            }
        }
        return collectedFrames;
    }

    /**
     * 選択からストーリーを取り出す
     * @returns {Story|null} ストーリー。取得できない場合は null
     */
    function getSelectionStory(item) {
        try {
            if (item.parentStory && item.parentStory.isValid) return item.parentStory;
        } catch (e) { }
        return null;
    }

    /**
     * グループ内のフレームを再帰的に集める
     * @param {Group} group 対象のグループ
     * @param {Array<PageItem>} frames 収集先の配列
     * @returns {void}
     */
    function collectFramesFromGroup(group) {
        var collectedFrames = [];
        var groupItems = group.allPageItems;
        for (var i = 0; i < groupItems.length; i++) {
            if (isEditableFrame(groupItems[i])) collectedFrames.push(groupItems[i]);
        }
        return collectedFrames;
    }

    /**
     * 重複を避けながらフレームの配列を追加する
     * @param {Array<PageItem>} frames 収集先の配列
     * @param {Array<PageItem>} newFrames 追加するフレーム
     * @returns {void}
     */
    function appendFrames(targetFrames, seenIds, sourceFrames) {
        for (var i = 0; i < sourceFrames.length; i++) {
            appendFrame(targetFrames, seenIds, sourceFrames[i]);
        }
    }

    /**
     * 重複を避けながらフレームを 1 つ追加する
     * @param {Array<PageItem>} frames 収集先の配列
     * @param {PageItem} frame 追加するフレーム
     * @returns {void}
     */
    function appendFrame(targetFrames, seenIds, frame) {
        var key = "" + frame.id;
        if (seenIds[key]) return;
        seenIds[key] = true;
        targetFrames.push(frame);
    }

    // =========================================
    // 実行 / Apply
    // =========================================

    /**
     * 収集したフレームへ設定を適用する
     * @param {Array<PageItem>} frames 対象フレーム
     * @param {object} settings 適用する設定
     * @returns {number} 実際に調整した件数
     */
    function applyToFrames(targetFrames, settings) {
        var processedCount = 0;
        for (var i = 0; i < targetFrames.length; i++) {
            var frame = targetFrames[i];
            /* 変化検出のため処理前の状態を記録 / Snapshot the state before processing, to detect real changes */
            var beforeSignature = frameSignature(frame);
            applyToFrame(frame, settings);
            /* インライン画像を前後の文字サイズに合わせる / Match inline graphics to surrounding text size */
            if (settings.adjustInlineImageInText) matchInlineHeightToText(frame);
            /* フレームか内容が実際に変化したものだけを数える / Count only frames whose geometry or content actually changed */
            if (frameSignature(frame) !== beforeSignature) processedCount++;
        }
        return processedCount;
    }

    /**
     * 変更の有無を比べるためのフレームの状態を文字列化する
     * @param {PageItem} frame 対象のフレーム
     * @returns {string} 状態を表す文字列
     */
    function frameSignature(frame) {
        var parts = [];
        try {
            var bounds = frame.geometricBounds; // [y1, x1, y2, x2]
            parts.push(roundForCompare(bounds[0]), roundForCompare(bounds[1]), roundForCompare(bounds[2]), roundForCompare(bounds[3]));
        } catch (e) {
            parts.push("nb");
        }
        var graphic = getSingleImage(frame);
        if (graphic !== null) {
            try {
                parts.push(roundForCompare(graphic.horizontalScale), roundForCompare(graphic.verticalScale));
            } catch (eScale) { parts.push("ns"); }
            try {
                /* 画像自身の位置（フィットで動くため変化検出に含める）/ The image's own position (moves on fit, so include it) */
                var graphicBounds = graphic.geometricBounds; // [y1, x1, y2, x2]
                parts.push(roundForCompare(graphicBounds[0]), roundForCompare(graphicBounds[1]), roundForCompare(graphicBounds[2]), roundForCompare(graphicBounds[3]));
            } catch (eBounds) { parts.push("ng"); }
        }
        return parts.join(",");
    }

    /**
     * 比較用に数値を丸める
     * @param {number} value 対象の数値
     * @returns {number} 丸めた数値
     */
    function roundForCompare(value) {
        if (typeof value !== "number" || !isFinite(value)) return "x";
        return Math.round(value * 10000) / 10000;
    }

    /**
     * 1 つのフレームへ設定を適用する
     * @param {PageItem} frame 対象のフレーム
     * @param {object} settings 適用する設定
     * @returns {void}
     */
    function applyToFrame(frame, settings) {
        /* 縦横比の崩れた画像を縦横同率に正す（設定に依らず常に実行）/ Always correct distorted (non-uniform) image scaling */
        correctImageAspectRatio(frame);

        /* 前後に文字がある真のインライン画像は、ここでは何もしない（縮尺率切り捨て・内容フィット・フレーム幅調整は行わない）。
           高さ合わせは applyToFrames 側が「インライン画像を文字サイズに合わせる」ON のときだけ実行する
           True inline images with surrounding text are left untouched here (no round-down, fit-to-content, or width change).
           Height matching is handled by the caller only when "Match Inline Images to Text Size" is on */
        if (getInlineSurroundingPointSize(frame) > 0) return;

        /* 縮尺率を切り捨てる / Round the image scale down */
        if (settings.roundScale) {
            roundImageScale(frame, settings.roundPrecision, settings.refitAfterRound, settings.roundOnly7296144);
        }

        /* フレームを内容に合わせる / Fit frame to content */
        if (settings.fitFrameToContent) {
            safeFit(frame, FitOptions.frameToContent);
        }

        /* フレームの幅の調整 / Adjust frame width */
        if (settings.widthMode === "parent") {
            fitWidthToParentFrame(frame, settings.widthOnlyIfLarger);
        } else if (settings.widthMode === "margin") {
            fitWidthToMargin(frame, settings.widthOnlyIfLarger);
        }
    }

    /**
     * フィット処理を例外を握りつぶして実行する
     * @param {PageItem} frame 対象のフレーム
     * @param {FitOptions} fitOption フィット方法
     * @returns {boolean} フィットできたら true
     */
    function safeFit(frame, fitOption) {
        try {
            frame.fit(fitOption);
            return true;
        } catch (e) {
            return false; // ロック・空フレームなど / Locked, empty, etc.
        }
    }

    /**
     * フレーム内の画像が 1 点だけならそれを返す
     * @param {PageItem} frame 対象のフレーム
     * @returns {Graphic|null} 画像。該当しない場合は null
     */
    function getSingleImage(frame) {
        try {
            if (frame.allGraphics && frame.allGraphics.length === 1) {
                var graphic = frame.allGraphics[0];
                var typeName = graphic.constructor.name;
                if (typeName === "Image" || typeName === "PDF" || typeName === "EPS") return graphic;
            }
        } catch (e) { }
        return null;
    }

    /**
     * 指定した刻み幅で切り捨てる
     * @param {number} value 対象の数値
     * @param {number} step 刻み幅
     * @returns {number} 切り捨てた数値
     */
    function floorToStep(value, step) {
        return Math.floor(value / step) * step;
    }

    /**
     * 切り捨て対象の解像度かどうかを判定する
     * @param {Graphic} image 対象の画像
     * @returns {boolean} 対象なら true
     */
    function isTargetResolution(graphic) {
        try {
            var actualPpi = graphic.actualPpi; // [horizontal, vertical]
            if (!actualPpi || actualPpi.length < 2) return false;
            for (var i = 0; i < ROUND_TARGET_PPI.length; i++) {
                if (Math.round(actualPpi[0]) === ROUND_TARGET_PPI[i] ||
                    Math.round(actualPpi[1]) === ROUND_TARGET_PPI[i]) {
                    return true;
                }
            }
        } catch (e) { }
        return false;
    }

    /**
     * 画像の縮尺率を指定単位で切り捨てる
     * @param {PageItem} frame 対象のフレーム
     * @param {object} settings 適用する設定
     * @returns {boolean} 変更したら true
     */
    function roundImageScale(frame, precision, refit, onlyTargetPpi) {
        var graphic = getSingleImage(frame);
        if (graphic === null) return false;
        /* 解像度フィルタ：72/96/144 ppi 以外はスキップ / Resolution filter: skip non-72/96/144 ppi images */
        if (onlyTargetPpi && !isTargetResolution(graphic)) return false;
        try {
            /* 横スケールを基準に切り捨て、縦にも同じ値を適用（縦横同率に揃える）/ Floor horizontal scale, apply the same value to vertical */
            var flooredScale = floorToStep(graphic.horizontalScale, precision);
            /* 0 など不正値になる場合は最小ステップに留める（エラー 11268 回避）/ Clamp to one step if flooring hits 0, etc. (avoids error 11268) */
            if (flooredScale < precision) flooredScale = precision;
            if (!isValidScale(flooredScale)) return false;

            graphic.horizontalScale = flooredScale;
            graphic.verticalScale = flooredScale;
            if (refit) safeFit(frame, FitOptions.frameToContent); // 再フィット / Re-fit
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 崩れた縦横比を横スケール基準に揃える
     * @param {PageItem} frame 対象のフレーム
     * @returns {boolean} 変更したら true
     */
    function correctImageAspectRatio(frame) {
        var graphic = getSingleImage(frame);
        if (graphic === null) return false;
        try {
            var hScale = graphic.horizontalScale;
            var vScale = graphic.verticalScale;
            if (!isValidScale(hScale) || !isValidScale(vScale)) return false;
            /* 既に同率（許容差内）なら触らない / Already uniform within tolerance */
            if (Math.abs(hScale - vScale) <= ASPECT_RATIO_TOLERANCE) return false;
            /* 横スケールを基準に縦を合わせる / Match vertical scale to the horizontal scale */
            graphic.verticalScale = hScale;
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * アンカー元の親テキストフレームを取得する
     * @param {PageItem} frame 対象のフレーム
     * @returns {TextFrame|null} 親テキストフレーム。取得できない場合は null
     */
    function getParentTextFrame(frame) {
        try {
            var anchorCharacter = getAnchorCharacter(frame);
            if (anchorCharacter !== null && anchorCharacter.parentTextFrames && anchorCharacter.parentTextFrames.length > 0) {
                return anchorCharacter.parentTextFrames[0];
            }
        } catch (e) { }

        try {
            var parent = frame.parent;
            if (parent && parent.parentTextFrames && parent.parentTextFrames.length > 0) {
                return parent.parentTextFrames[0];
            }
            if (parent && parent.parent && parent.parent.parentTextFrames && parent.parent.parentTextFrames.length > 0) {
                return parent.parent.parentTextFrames[0];
            }
        } catch (err) { }

        return null;
    }

    /**
     * フレーム幅を親テキストフレームの内寸に合わせる
     * @param {PageItem} frame 対象のフレーム
     * @param {object} settings 適用する設定
     * @returns {boolean} 変更したら true
     */
    function fitWidthToParentFrame(frame, onlyIfLarger) {
        try {
            var parentTextFrame = getParentTextFrame(frame);
            if (parentTextFrame === null) return false;

            var parentBounds = parentTextFrame.geometricBounds; // [y1, x1, y2, x2]
            var parentWidth = parentBounds[3] - parentBounds[1];

            /* テキストフレームの左右インセットを控除（取得不可なら 0）/ Subtract left/right text insets (0 if unavailable) */
            var insets = getTextFrameInsets(parentTextFrame);
            var contentWidth = parentWidth - insets.left - insets.right;
            if (contentWidth <= 0) return false; // 不正な幅は触らない / Skip invalid widths

            var bounds = frame.geometricBounds; // [y1, x1, y2, x2]
            var currentWidth = bounds[3] - bounds[1];
            /* 親の内寸以下のフレームはそのまま / Leave frames that already fit within the parent */
            if (onlyIfLarger && currentWidth <= contentWidth) return false;

            /* 左端を保持して親の内寸幅へ。広い→収める / 狭い→いっぱいに流し込む（マージンと同じ挙動）
               Keep the left edge, snap to the parent width. Wider→fit inside / narrower→fill (same as margin) */
            applyWidthFit(frame, bounds[1], bounds[1] + contentWidth, currentWidth);
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 指定した幅にフレームを合わせ、内容を収め直す
     * @param {PageItem} frame 対象のフレーム
     * @param {number} targetWidth 目標の幅
     * @param {object} settings 適用する設定
     * @returns {boolean} 変更したら true
     */
    function applyWidthFit(frame, left, right, originalWidth) {
        var bounds = frame.geometricBounds; // [y1, x1, y2, x2]
        frame.geometricBounds = [bounds[0], left, bounds[2], right];

        var targetWidth = right - left;
        if (originalWidth > targetWidth) {
            safeFit(frame, FitOptions.proportionally);   // 広い：内容を内側に収める / Wider: fit inside
        } else {
            safeFit(frame, FitOptions.fillProportionally); // 狭い：いっぱいに流し込む / Narrower: fill
        }
        safeFit(frame, FitOptions.frameToContent);
    }

    /**
     * テキストフレームの内側余白を取得する
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @returns {object} 各辺の余白
     */
    function getTextFrameInsets(textFrame) {
        var result = { left: 0, right: 0 };
        try {
            var insetPrefs = textFrame.textFramePreferences;
            var spacing = insetPrefs.insetSpacing; // 数値 or [top, left, bottom, right]
            if (spacing instanceof Array) {
                result.left = spacing[1];
                result.right = spacing[3];
            } else if (typeof spacing === "number") {
                result.left = spacing;
                result.right = spacing;
            }
        } catch (e) { }
        return result;
    }

    /**
     * フレーム幅をページのマージン幅に合わせる
     * @param {PageItem} frame 対象のフレーム
     * @param {object} settings 適用する設定
     * @returns {boolean} 変更したら true
     */
    function fitWidthToMargin(frame, onlyIfLarger) {
        try {
            var page = frame.parentPage;
            if (page === null) return false;

            /* マージン量はページ単位で取得（見開き左右でも正しい）/ Read margins per page (correct on facing pages) */
            var margins = page.marginPreferences;
            var pageBounds = page.bounds; // [y1, x1, y2, x2]
            var left = pageBounds[1] + margins.left;
            var right = pageBounds[3] - margins.right;
            var marginWidth = right - left;
            if (marginWidth <= 0) return false; // 不正な版面幅は触らない / Skip invalid live-area widths

            /* 元のフレーム幅を記録してから版面幅へ / Remember the original width, then snap to the live area */
            var bounds = frame.geometricBounds; // [y1, x1, y2, x2]
            var originalWidth = bounds[3] - bounds[1];
            /* 版面幅以下のフレームはそのまま（縮小のみ）/ Leave frames within the live area (shrink only) */
            if (onlyIfLarger && originalWidth <= marginWidth) return false;

            /* 版面幅へ。広い→収める / 狭い→いっぱいに流し込む / Snap to the live area. Wider→fit / narrower→fill */
            applyWidthFit(frame, left, right, originalWidth);
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * インライン画像の高さを同じ段落の文字サイズに合わせる
     * @param {PageItem} frame 対象のフレーム
     * @returns {boolean} 変更したら true
     */
    function matchInlineHeightToText(frame) {
        try {
            /* インラインでない・前後に文字が無い（画像のみの段落）ときは 0 / 0 when not inline or when there is no surrounding text */
            var targetPoint = getInlineSurroundingPointSize(frame);
            if (targetPoint <= 0) return false;

            var graphic = getSingleImage(frame);
            if (graphic === null) return false; // 単一画像でなければ対象外 / Only single-image frames

            /* まずフレームを画像実寸に合わせ、現在の高さ（pt）を取得 / Fit to content first, then read the height in points */
            safeFit(frame, FitOptions.frameToContent);
            var currentPoint = getFrameHeightPoint(frame);
            if (!isFinite(currentPoint) || currentPoint <= 0) return false;

            /* 目標高さになるよう画像を比例拡大縮小 / Scale the image proportionally to reach the target height */
            var factor = targetPoint / currentPoint;
            var currentHScale = graphic.horizontalScale;
            var currentVScale = graphic.verticalScale;
            if (!isValidScale(currentHScale) || !isValidScale(currentVScale)) return false;

            var newHScale = currentHScale * factor;
            var newVScale = currentVScale * factor;
            /* 不正・範囲外のスケール値は適用しない（エラー 11268 回避）/ Skip invalid/out-of-range scales (avoids error 11268) */
            if (!isValidScale(newHScale) || !isValidScale(newVScale)) return false;

            graphic.horizontalScale = newHScale;
            graphic.verticalScale = newVScale;

            /* 縮小後の画像にフレームを合わせ直す / Re-fit the frame to the scaled image */
            safeFit(frame, FitOptions.frameToContent);
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * スケール値として有効かどうかを判定する
     * @param {number} value スケール値
     * @returns {boolean} 有効なら true
     */
    function isValidScale(scale) {
        return typeof scale === "number" && isFinite(scale) && scale > 0 && scale <= 1000000;
    }

    /**
     * インライン画像の前後にある文字サイズを取得する
     * @param {PageItem} frame 対象のフレーム
     * @returns {number} 文字サイズ（pt）
     */
    function getInlineSurroundingPointSize(frame) {
        var anchorCharacter = getAnchorCharacter(frame);
        if (anchorCharacter === null) return 0; // インラインでなければ 0 / Not inline
        try {
            return getSurroundingTextPointSize(anchorCharacter.paragraphs[0]);
        } catch (e) {
            return 0;
        }
    }

    /**
     * 真のインライン配置かどうかを判定する
     * @param {PageItem} frame 対象のフレーム
     * @returns {boolean} インライン配置なら true
     */
    function isInlineAnchored(frame) {
        try {
            var anchoredSettings = frame.anchoredObjectSettings;
            return anchoredSettings && anchoredSettings.anchoredPosition === AnchorPosition.INLINE_POSITION;
        } catch (e) {
            return false;
        }
    }

    /**
     * フレームがテキストにアンカーされているか（親が文字なら、インライン・行の上・カスタムのいずれか）
     * @param {PageItem} pageItem 判定するオブジェクト
     * @returns {boolean} アンカー付きなら true
     */
    function isAnchoredFrame(pageItem) {
        return pageItem.parent.constructor.name === "Character";
    }

    /**
     * インライン配置のアンカー付きオブジェクトについて、アンカーを表す文字を取得する
     * @param {PageItem} frame 対象のフレーム
     * @returns {Character|null} アンカー文字。インラインでない・取得できない場合は null
     */
    function getAnchorCharacter(frame) {
        try {
            if (!isInlineAnchored(frame)) return null; // 行揃え・カスタム配置は対象外 / Skip above-line / custom positions
            if (isAnchoredFrame(frame)) return frame.parent;
        } catch (e) { }
        return null;
    }

    /**
     * アンカー文字のある段落で、本文の文字サイズ（最大値）を求める
     * @param {Paragraph} paragraph アンカー文字のある段落
     * @returns {number} 文字サイズ（pt）
     */
    function getSurroundingTextPointSize(paragraph) {
        var chars = paragraph.characters.everyItem().getElements();
        var maxSize = 0;
        for (var i = 0; i < chars.length; i++) {
            var content = chars[i].contents;
            if (content === "\r" || content === "\n") continue;              // 改行 / line breaks
            if (content === "\uFFFC") continue;                              // アンカーオブジェクトマーカー / anchored object marker
            if (content === " " || content === "\u3000" || content === "\t") continue; // 空白 / spaces
            try {
                var pointSize = chars[i].pointSize;
                if (pointSize !== NothingEnum.NOTHING) {
                    pointSize = Number(pointSize);
                    if (pointSize > maxSize) maxSize = pointSize;
                }
            } catch (e) { }
        }
        return maxSize;
    }

    /**
     * フレームの高さをポイントで取得する
     * @param {PageItem} frame 対象のフレーム
     * @returns {number} 高さ（pt）
     */
    function getFrameHeightPoint(frame) {
        var view = app.activeDocument.viewPreferences;
        var oldH = view.horizontalMeasurementUnits;
        var oldV = view.verticalMeasurementUnits;
        try {
            view.horizontalMeasurementUnits = MeasurementUnits.POINTS;
            view.verticalMeasurementUnits = MeasurementUnits.POINTS;
            var bounds = frame.geometricBounds; // [y1, x1, y2, x2]（ポイント / points）
            return bounds[2] - bounds[0];
        } catch (e) {
            return 0;
        } finally {
            view.horizontalMeasurementUnits = oldH;
            view.verticalMeasurementUnits = oldV;
        }
    }

})();
