#target indesign
#targetengine "IdAlignToNearestEdgePalette"

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdAlignToNearestEdgePalette
Name=移動方向でいちばん近いガイド・マージン・ページ・裁ち落としに揃えるパレット
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.2.1
Release-Date=2026-11-03
Description-BEGIN
選択したオブジェクトを、パレットのボタンで選んだ方向へ動かし、その方向でいちばん近いガイド・マージンガイド・ページのエッジ・裁ち落としに辺を揃えます。中央のボタンは、ページかマージンの内側の中央に揃えます。

Moves the selected objects in the direction chosen on the palette and aligns their edge to the nearest guide, margin guide, page edge, or bleed edge in that direction. The center buttons align to the center of the page or of the margins.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択したオブジェクトを、パレットのボタンで選んだ方向へ動かし、その方向でいちばん近いガイド・マージンガイド・ページのエッジ・裁ち落としに辺を揃えます。中央のボタンは、ページかマージンの内側の中央に揃えます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAlignToNearestEdgePalette.md

### Overview

Moves the selected objects in the direction chosen on the palette and aligns their edge to the nearest guide, margin guide, page edge, or bleed edge in that direction. The center buttons align to the center of the page or of the margins.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAlignToNearestEdgePalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdAlignToNearestEdgePalette";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-05";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "移動方向でいちばん近いガイド・マージン・ページ・裁ち落としに揃えるパレット"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAlignToNearestEdgePalette.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAlignToNearestEdgePalette.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/frame/IdAlignToNearestEdgePalette.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 辺と線が一致しているとみなす誤差（pt）/ Tolerance for an edge to count as on a line (pt) */
    var EDGE_TOLERANCE_POINTS = 0.001;

    // =========================================
    // 外部からの実行 / Running from outside
    // =========================================

    /* LayoutKeys などで初期値を差し替えると、パレットを出さずにそのボタンの揃え方で実行する。1行に1つずつ書く（差し替えの対象になる形） */
    /* When a launcher such as LayoutKeys changes these initial values, the button runs without the palette. Keep one per line (the form launchers can change) */
    /* 行末の [a|b|…] は LayoutKeys が候補として読む / LayoutKeys reads the [a|b|…] list as the choices */
    var SW_ACTION       = "";     /* 空ならパレット [upLeft|up|upRight|left|center|right|downLeft|down|downRight|centerH|centerV] / empty for the palette */
    var SW_FROM_CENTER  = false;  /* true なら、いったん中央に移してから揃える（Option＋クリックと同じ）/ true: center first, like Option-click */
    var SW_MOVE_AS_UNIT = "";     /* ［選択をまとめて移動］。空ならパレットの設定 [true|false] / "Move selection as a unit"; empty follows the palette */
    var SW_CENTER_BASIS = "";     /* ［中央揃えの基準］。空ならパレットの設定 [page|margin] / "Center Alignment Basis"; empty follows the palette */

    // =========================================
    // 揃え方 / Alignment modes
    // =========================================

    /* 軸ごとの揃え方。edgeIndex は visibleBounds [天, 左, 地, 右] の添字（中心は null。中心はページの中央に揃える）、sign は座標が増える向きなら 1 */
    /* One mode per axis. edgeIndex indexes visibleBounds [top, left, bottom, right] (null for the center, which goes to the page center); sign is 1 when the coordinate grows */
    var AXIS_MODES = {
        left:    { axis: "x", edgeIndex: 1, sign: -1 },
        right:   { axis: "x", edgeIndex: 3, sign: 1 },
        top:     { axis: "y", edgeIndex: 0, sign: -1 },
        bottom:  { axis: "y", edgeIndex: 2, sign: 1 },
        centerX: { axis: "x", edgeIndex: null, sign: 0 },
        centerY: { axis: "y", edgeIndex: null, sign: 0 }
    };

    /* 軸の両端の visibleBounds の添字 / visibleBounds indices of both ends of each axis */
    var AXIS_BOUNDS_INDICES = { x: [1, 3], y: [0, 2] };

    /* ボタンごとの揃え方（斜めは2軸）/ The axis modes of each button; a diagonal uses two */
    var ACTION_AXIS_MODES = {
        upLeft:    ["top", "left"],
        up:        ["top"],
        upRight:   ["top", "right"],
        left:      ["left"],
        center:    ["centerX", "centerY"],
        right:     ["right"],
        downLeft:  ["bottom", "left"],
        down:      ["bottom"],
        downRight: ["bottom", "right"],
        centerH:   ["centerX"],
        centerV:   ["centerY"]
    };

    /* 十字に並べるボタン / The buttons laid out as a 3×3 cross */
    var MOVE_BUTTON_ROWS = [
        ["upLeft",   "up",     "upRight"],
        ["left",     "center", "right"],
        ["downLeft", "down",   "downRight"]
    ];

    /* 十字の下に並べる、片方の軸だけ中央に揃えるボタン / The single-axis center buttons below the cross */
    var CENTER_BUTTON_KEYS = ["centerH", "centerV"];

    /* パレットで押した矢印キー（ScriptUI の keyName）とボタン / Arrow keys pressed on the palette (ScriptUI keyName) */
    var MOVE_ACTION_BY_KEY_NAME = { Up: "up", Left: "left", Right: "right", Down: "down" };

    // =========================================
    // セッション記憶 / Session memory
    // =========================================

    /* パレットの参照と設定を InDesign の終了まで $.global に残す / Keeps the palette and settings on $.global until InDesign quits */
    var PALETTE_GLOBAL_KEY = "idAlignToNearestEdgePalette";
    var SETTINGS_GLOBAL_KEY = "idAlignToNearestEdgeSettings";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var ICON_SIZE      = 30;  /* ボタン1個の大きさ（px）/ size of each button (px) */
    var CROSS_GAP      = 2;   /* 十字のボタンどうしの間隔 / gap between the cross buttons */
    var CENTER_ROW_TOP = 8;   /* 十字と中央揃えボタンの間隔 / gap between the cross and the center buttons */

    /* アイコンの形。すべて ICON_SIZE に対する比率 / Icon shapes, all as ratios of ICON_SIZE */
    var ICON_RULE_INSET     = 0.12;  /* ケイ線の端の余白 / inset at both ends of a rule */
    var ICON_RULE_OFFSET    = 0.17;  /* 端に置くケイ線の位置 / position of a rule at an edge */
    var ICON_RULE_CLEARANCE = 0.05;  /* 端のケイ線とオブジェクトのすき間 / gap between an edge rule and the block */
    var ICON_BLOCK_WIDTH    = 0.38;  /* オブジェクトの幅 / width of the block */
    var ICON_BLOCK_HEIGHT   = 0.30;  /* オブジェクトの高さ / height of the block */

    /* 方向ボタンのアイコンで、寄せ先を表すケイ線を引く辺（斜めは2辺）/ The edges that get a rule in each icon; a diagonal gets two */
    var MOVE_ICON_RULES = {
        up:        ["top"],
        down:      ["bottom"],
        left:      ["left"],
        right:     ["right"],
        upLeft:    ["top", "left"],
        upRight:   ["top", "right"],
        downLeft:  ["bottom", "left"],
        downRight: ["bottom", "right"]
    };

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

    // =========================================
    // 配色 / Colors
    // =========================================

    /* パレットの地に重ねる半透明の白・黒（地の明るさが分からないため、重ねる色で明暗を付ける）*/
    /* Translucent overlays on the palette background (its exact color is not readable, so shade with overlays) */
    var ICON_UI_DARK      = isDarkUI();
    var ICON_COLOR        = ICON_UI_DARK ? [0.85, 0.85, 0.85, 1] : [0.25, 0.25, 0.25, 1]; /* ケイ線とオブジェクト / rules and block */
    var ICON_BASE_BG      = ICON_UI_DARK ? [1, 1, 1, 0.06]       : [1, 1, 1, 0.60];       /* 通常時の地 / background */
    var ICON_HOVER_BG     = ICON_UI_DARK ? [1, 1, 1, 0.16]       : [0, 0, 0, 0.08];       /* マウスオーバー時の地 / background while hovered */
    var ICON_BORDER_COLOR = ICON_UI_DARK ? [1, 1, 1, 0.25]       : [0, 0, 0, 0.30];       /* マウスオーバー時の枠 / border while hovered */

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

    var LABELS = {
        palette: {
            title: { ja: "近くの線に揃える", en: "Align to Nearest Edge" }
        },
        action: {
            upLeft:    { ja: "左上", en: "Top Left" },
            up:        { ja: "上", en: "Top" },
            upRight:   { ja: "右上", en: "Top Right" },
            left:      { ja: "左", en: "Left" },
            center:    { ja: "中央", en: "Center" },
            right:     { ja: "右", en: "Right" },
            downLeft:  { ja: "左下", en: "Bottom Left" },
            down:      { ja: "下", en: "Bottom" },
            downRight: { ja: "右下", en: "Bottom Right" },
            centerH:   { ja: "水平方向中央", en: "Horizontal Center" },
            centerV:   { ja: "垂直方向中央", en: "Vertical Center" }
        },
        tooltip: {
            upLeft:    { ja: "上辺と左辺を、それぞれの方向でいちばん近い線に揃える", en: "Aligns the top and left edges to the nearest line in each direction" },
            up:        { ja: "上辺を、上方向でいちばん近い線に揃える（↑キー）", en: "Aligns the top edge to the nearest line above (Up Arrow)" },
            upRight:   { ja: "上辺と右辺を、それぞれの方向でいちばん近い線に揃える", en: "Aligns the top and right edges to the nearest line in each direction" },
            left:      { ja: "左辺を、左方向でいちばん近い線に揃える（←キー）", en: "Aligns the left edge to the nearest line on the left (Left Arrow)" },
            center:    { ja: "中心を、ページ（またはマージンの内側）の中央に揃える", en: "Aligns the center to the center of the page (or of the margins)" },
            right:     { ja: "右辺を、右方向でいちばん近い線に揃える（→キー）", en: "Aligns the right edge to the nearest line on the right (Right Arrow)" },
            downLeft:  { ja: "下辺と左辺を、それぞれの方向でいちばん近い線に揃える", en: "Aligns the bottom and left edges to the nearest line in each direction" },
            down:      { ja: "下辺を、下方向でいちばん近い線に揃える（↓キー）", en: "Aligns the bottom edge to the nearest line below (Down Arrow)" },
            downRight: { ja: "下辺と右辺を、それぞれの方向でいちばん近い線に揃える", en: "Aligns the bottom and right edges to the nearest line in each direction" },
            centerH:   { ja: "水平方向の中心を、ページ（またはマージンの内側）の左右の中央に揃える", en: "Aligns the horizontal center to the horizontal center of the page (or of the margins)" },
            centerV:   { ja: "垂直方向の中心を、ページ（またはマージンの内側）の天地の中央に揃える", en: "Aligns the vertical center to the vertical center of the page (or of the margins)" },
            centerBasis: { ja: "中央揃えのボタン（十字・水平方向中央・垂直方向中央）で、どの範囲の中央に揃えるか", en: "Which area the center buttons (cross, horizontal center, vertical center) align to" },
            edgeSuffix: { ja: "続けて押すと次の線へ進む", en: "Press again to move on to the next line" },
            optionFromCenter: { ja: "Option（Alt）＋クリックで、いったん中央（中央揃えの基準の範囲）に移してから揃える", en: "Option (Alt)-click: center first (within the center alignment basis), then align" },
            targetSuffix: { ja: "揃える先：ガイド・マージンガイド・ページのエッジ・裁ち落とし", en: "Targets: guides, margin guides, page edges, bleed" },
            moveAsUnit: { ja: "選択全体の境界を基準に、相対位置を保ったまま動かす", en: "Moves everything by the same amount, based on the bounds of the whole selection" }
        },
        panel: {
            centerBasis: { ja: "中央揃えの基準", en: "Center Alignment Basis" }
        },
        radio: {
            page: { ja: "ページ", en: "Page" },
            margin: { ja: "マージン", en: "Margins" }
        },
        checkbox: {
            moveAsUnit: { ja: "選択をまとめて移動", en: "Move selection as a unit" }
        },
        button: {
            close: { ja: "閉じる", en: "Close" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noMovableItem: { ja: "移動できるオブジェクトを選択してください（ロックしたもの・アンカー付きオブジェクトは対象外）。", en: "Please select objects that can be moved (locked and anchored objects are skipped)." },
            noLineAhead: { ja: "揃える先が見つかりませんでした。", en: "No target line was found." },
            unknownAction: { ja: "SW_ACTION の値「{name}」はボタンの名前ではありません。", en: "SW_ACTION \"{name}\" is not a button name." }
        },
        undo: {
            align: { ja: "近くの線に揃える（{name}）", en: "Align to Nearest Edge ({name})" }
        }
    };

    // =========================================
    // 設定 / Settings
    // =========================================

    /**
     * @typedef {Object} PaletteSettings
     * @property {boolean} moveAsUnit - 選択をまとめて移動するなら true
     * @property {string} centerBasis - 中央揃えの基準（"page"＝ページ、"margin"＝マージンの内側）
     * @property {number[]|null} paletteLocation - 前回閉じたときのパレットの位置 [x, y]
     */

    /**
     * 前回の設定を返す（無ければ既定の設定）。返したオブジェクトへの変更はそのまま次回に残る
     * @returns {PaletteSettings} 設定
     */
    function getSessionSettings() {
        if (!$.global[SETTINGS_GLOBAL_KEY]) $.global[SETTINGS_GLOBAL_KEY] = { moveAsUnit: false, paletteLocation: null };
        /* 前の版で残した設定には centerBasis が無い / Settings kept by an earlier version lack centerBasis */
        if (!$.global[SETTINGS_GLOBAL_KEY].centerBasis) $.global[SETTINGS_GLOBAL_KEY].centerBasis = "page";
        return $.global[SETTINGS_GLOBAL_KEY];
    }

    // =========================================
    // 対象のオブジェクト / Target objects
    // =========================================

    /**
     * 選択から動かせるページアイテムを集める（文字の選択はそのテキストフレームにする。
     * ロックしたもの・ロックしたレイヤーのもの・アンカー付きオブジェクトは除く）
     * @param {Array<Object>} selectedObjects - app.selection
     * @returns {PageItem[]} 動かせるページアイテム（重複なし）
     */
    function collectMovableItems(selectedObjects) {
        var movableItems = [];
        var seenIds = {};
        for (var i = 0; i < selectedObjects.length; i++) {
            var pageItem = selectedObjects[i];
            if (!pageItem.hasOwnProperty("visibleBounds") && pageItem.hasOwnProperty("parentTextFrames")) {
                pageItem = pageItem.parentTextFrames[0];
            }
            if (!pageItem || !pageItem.isValid || !pageItem.hasOwnProperty("visibleBounds")) continue;
            if (pageItem.locked || pageItem.itemLayer.locked) continue;
            if (pageItem.parent.constructor.name === "Character") continue;
            if (seenIds["id:" + pageItem.id]) continue;
            seenIds["id:" + pageItem.id] = true;
            movableItems.push(pageItem);
        }
        return movableItems;
    }

    /**
     * ページアイテムが載っているスプレッド（親ページのスプレッドを含む）を返す
     * @param {PageItem} pageItem - 対象
     * @returns {Spread|MasterSpread|null} スプレッド。見つからなければ null
     */
    function findParentSpread(pageItem) {
        var container = pageItem.parent;
        while (container && container.constructor.name !== "Document") {
            if (container.constructor.name === "Spread" || container.constructor.name === "MasterSpread") return container;
            container = container.parent;
        }
        return null;
    }

    // =========================================
    // 揃える先の線 / Target lines
    // =========================================

    /**
     * スプレッドにある、ガイド・マージンガイド・ページのエッジ・裁ち落としの線の座標をまとめて集める
     * @param {Spread|MasterSpread} spread - 対象のスプレッド
     * @param {string} axis - "x"（縦の線の X 座標）または "y"（横の線の Y 座標）
     * @param {Document} doc - 対象ドキュメント
     * @returns {number[]} 座標（pt、スプレッド基準）
     */
    function collectTargetLines(spread, axis, doc) {
        return collectGuideLines(spread, axis)
            .concat(collectMarginLines(spread, axis))
            .concat(collectPageEdgeLines(spread, axis))
            .concat(collectBleedLines(spread, axis, doc));
    }

    /**
     * 指定した向きのガイドの座標を足す（非表示のレイヤーのガイドは除く）
     * @param {number[]} lines - 足す先
     * @param {Guides} guides - ガイドのコレクション
     * @param {HorizontalOrVertical} orientation - ガイドの向き
     * @param {number} offset - 座標に足す量（親ページのガイドを本文のページへ写すとき）
     * @param {number[]|null} range - [最小, 最大]。範囲外のガイドは除く（null なら全部）
     * @returns {void}
     */
    function addGuideLines(lines, guides, orientation, offset, range) {
        for (var i = 0; i < guides.length; i++) {
            var guide = guides[i];
            if (guide.orientation !== orientation) continue;
            if (!guide.itemLayer.visible) continue;
            var guidePosition = guide.location;
            if (range && (guidePosition < range[0] - EDGE_TOLERANCE_POINTS || guidePosition > range[1] + EDGE_TOLERANCE_POINTS)) continue;
            lines.push(guidePosition + offset);
        }
    }

    /**
     * ページに当たっている親ページ（左ページは親の最初のページ、それ以外は最後のページ）を返す
     * @param {Page} page - 本文のページ
     * @returns {Page|null} 親ページ。［なし］なら null
     */
    function getAppliedMasterPage(page) {
        var masterSpread = page.appliedMaster;
        if (!masterSpread || masterSpread.constructor.name !== "MasterSpread") return null;
        var masterPages = masterSpread.pages;
        if (masterPages.length === 0) return null;
        return (page.side === PageSideOptions.LEFT_HAND) ? masterPages[0] : masterPages[masterPages.length - 1];
    }

    /**
     * スプレッドのガイド（スプレッド・各ページ・親ページのガイド）の座標を集める
     * @param {Spread|MasterSpread} spread - 対象のスプレッド
     * @param {string} axis - "x" または "y"
     * @returns {number[]} 座標
     */
    function collectGuideLines(spread, axis) {
        var isVertical = (axis === "x");
        var orientation = isVertical ? HorizontalOrVertical.VERTICAL : HorizontalOrVertical.HORIZONTAL;
        var boundsIndex = isVertical ? 1 : 0;
        var lines = [];
        addGuideLines(lines, spread.guides, orientation, 0, null);
        var pages = spread.pages;
        for (var i = 0; i < pages.length; i++) {
            addGuideLines(lines, pages[i].guides, orientation, 0, null);
            var masterPage = getAppliedMasterPage(pages[i]);
            if (!masterPage) continue;
            var masterBounds = masterPage.bounds;
            var offset = pages[i].bounds[boundsIndex] - masterBounds[boundsIndex];
            addGuideLines(lines, masterPage.guides, orientation, offset, null);
            /* 親のスプレッドの垂直ガイドは、この親ページの幅にあるものだけを写す（水平ガイドは見開き全体にかかる） */
            /* Map only the parent spread's vertical guides within this parent page; horizontal ones span the spread */
            addGuideLines(lines, masterPage.parent.guides, orientation, offset, isVertical ? [masterBounds[1], masterBounds[3]] : null);
        }
        return lines;
    }

    /**
     * ページのマージンの内側（マージンガイドで囲んだ範囲）の境界を返す（見開きの左ページは左右の内側・外側を入れ替える）
     * @param {Page} page - 対象のページ
     * @returns {number[]} [天, 左, 地, 右]
     */
    function getMarginAreaBounds(page) {
        var pageBounds = page.bounds;
        var marginPrefs = page.marginPreferences;
        var isLeftHand = (page.side === PageSideOptions.LEFT_HAND);
        return [
            pageBounds[0] + marginPrefs.top,
            pageBounds[1] + (isLeftHand ? marginPrefs.right : marginPrefs.left),
            pageBounds[2] - marginPrefs.bottom,
            pageBounds[3] - (isLeftHand ? marginPrefs.left : marginPrefs.right)
        ];
    }

    /**
     * 各ページのマージンガイドの座標を集める
     * @param {Spread|MasterSpread} spread - 対象のスプレッド
     * @param {string} axis - "x" または "y"
     * @returns {number[]} 座標
     */
    function collectMarginLines(spread, axis) {
        var lines = [];
        var endIndices = AXIS_BOUNDS_INDICES[axis];
        var pages = spread.pages;
        for (var i = 0; i < pages.length; i++) {
            var marginBounds = getMarginAreaBounds(pages[i]);
            lines.push(marginBounds[endIndices[0]], marginBounds[endIndices[1]]);
        }
        return lines;
    }

    /**
     * 各ページの端の座標を集める
     * @param {Spread|MasterSpread} spread - 対象のスプレッド
     * @param {string} axis - "x"（左右の端）または "y"（天地の端）
     * @returns {number[]} 座標
     */
    function collectPageEdgeLines(spread, axis) {
        var lines = [];
        var pages = spread.pages;
        for (var i = 0; i < pages.length; i++) {
            var pageBounds = pages[i].bounds;
            if (axis === "x") lines.push(pageBounds[1], pageBounds[3]);
            else lines.push(pageBounds[0], pageBounds[2]);
        }
        return lines;
    }

    /**
     * 見開きの裁ち落とし線の座標を返す。
     * 左端は左ページなら小口・それ以外はノド、右端は左ページならノド・それ以外は小口の裁ち落としを使う
     * （単ページのドキュメントは「ノド＝左」「小口＝右」になる）
     * @param {Spread|MasterSpread} spread - 対象のスプレッド
     * @param {string} axis - "x"（左右）または "y"（天地）
     * @param {Document} doc - 対象ドキュメント
     * @returns {number[]} 座標
     */
    function collectBleedLines(spread, axis, doc) {
        var docPrefs = doc.documentPreferences;
        var pages = spread.pages;
        if (axis === "y") {
            var topY = pages[0].bounds[0];
            var bottomY = pages[0].bounds[2];
            for (var j = 1; j < pages.length; j++) {
                topY = Math.min(topY, pages[j].bounds[0]);
                bottomY = Math.max(bottomY, pages[j].bounds[2]);
            }
            return [topY - docPrefs.documentBleedTopOffset, bottomY + docPrefs.documentBleedBottomOffset];
        }
        var insideBleed = docPrefs.documentBleedInsideOrLeftOffset;
        var outsideBleed = docPrefs.documentBleedOutsideOrRightOffset;
        var leftmostPage = pages[0];
        var rightmostPage = pages[0];
        for (var i = 1; i < pages.length; i++) {
            if (pages[i].bounds[1] < leftmostPage.bounds[1]) leftmostPage = pages[i];
            if (pages[i].bounds[3] > rightmostPage.bounds[3]) rightmostPage = pages[i];
        }
        var isLeftmostLeftHand = (leftmostPage.side === PageSideOptions.LEFT_HAND);
        var isRightmostLeftHand = (rightmostPage.side === PageSideOptions.LEFT_HAND);
        return [
            leftmostPage.bounds[1] - (isLeftmostLeftHand ? outsideBleed : insideBleed),
            rightmostPage.bounds[3] + (isRightmostLeftHand ? insideBleed : outsideBleed)
        ];
    }

    /**
     * 辺から移動方向に見て、いちばん近い線を返す（辺が載っている線は除くので、続けて押すと次の線へ進む）
     * @param {number[]} lines - 線の座標
     * @param {number} edgePosition - 辺の座標
     * @param {number} sign - 移動方向（座標が増える向きなら 1、減る向きなら -1）
     * @returns {number|null} 線の座標。無ければ null
     */
    function findNearestLineAhead(lines, edgePosition, sign) {
        var nearestPosition = null;
        var nearestDistance = 0;
        for (var i = 0; i < lines.length; i++) {
            var distance = (lines[i] - edgePosition) * sign;
            if (distance <= EDGE_TOLERANCE_POINTS) continue;
            if (nearestPosition === null || distance < nearestDistance) {
                nearestPosition = lines[i];
                nearestDistance = distance;
            }
        }
        return nearestPosition;
    }

    /**
     * オブジェクトが載っているページを返す（中心を含むページ。どのページにも無ければ、中心にいちばん近いページ）
     * @param {Spread|MasterSpread} spread - 対象のスプレッド
     * @param {number[]} bounds - オブジェクトの境界 [天, 左, 地, 右]
     * @returns {Page|null} ページ。スプレッドにページが無ければ null
     */
    function findCenterPage(spread, bounds) {
        var centerX = (bounds[1] + bounds[3]) / 2;
        var centerY = (bounds[0] + bounds[2]) / 2;
        var pages = spread.pages;
        var nearestPage = null;
        var nearestDistance = 0;
        for (var i = 0; i < pages.length; i++) {
            var pageBounds = pages[i].bounds;
            /* ページの外なら、いちばん近い辺までの距離（中なら 0）/ Distance to the page rectangle (0 when inside) */
            var distanceX = Math.max(pageBounds[1] - centerX, 0, centerX - pageBounds[3]);
            var distanceY = Math.max(pageBounds[0] - centerY, 0, centerY - pageBounds[2]);
            var distance = distanceX * distanceX + distanceY * distanceY;
            if (nearestPage === null || distance < nearestDistance) {
                nearestPage = pages[i];
                nearestDistance = distance;
            }
        }
        return nearestPage;
    }

    /**
     * 1つの軸で、揃える先までの移動量を求める（辺はいちばん近い線へ、中心はページかマージンの内側の中央へ）
     * @param {number[]} lines - その軸の線の座標
     * @param {number[]} bounds - オブジェクトの境界 [天, 左, 地, 右]
     * @param {Object} axisMode - AXIS_MODES の値
     * @param {number[]|null} centerBasisBounds - 中心を揃える範囲 [天, 左, 地, 右]
     * @returns {number|null} 移動量（pt）。揃える先が無ければ null
     */
    function findAxisOffset(lines, bounds, axisMode, centerBasisBounds) {
        if (axisMode.edgeIndex === null) {
            if (!centerBasisBounds) return null;
            var endIndices = AXIS_BOUNDS_INDICES[axisMode.axis];
            var itemCenter = (bounds[endIndices[0]] + bounds[endIndices[1]]) / 2;
            var basisCenter = (centerBasisBounds[endIndices[0]] + centerBasisBounds[endIndices[1]]) / 2;
            return basisCenter - itemCenter;
        }
        var edgePosition = bounds[axisMode.edgeIndex];
        var edgeTarget = findNearestLineAhead(lines, edgePosition, axisMode.sign);
        return (edgeTarget === null) ? null : edgeTarget - edgePosition;
    }

    // =========================================
    // 移動 / Moving
    // =========================================

    /**
     * 定規の単位を縦横とも pt、原点をスプレッドにして処理を実行し、終わったら（失敗しても）元に戻す
     * @param {Document} doc - 対象ドキュメント
     * @param {Function} task - 実行する処理（引数なし）
     * @returns {*} task の戻り値
     */
    function runInPointsWithSpreadOrigin(doc, task) {
        var viewPrefs = doc.viewPreferences;
        var savedHorizontalUnits = viewPrefs.horizontalMeasurementUnits;
        var savedVerticalUnits = viewPrefs.verticalMeasurementUnits;
        var savedRulerOrigin = viewPrefs.rulerOrigin;
        viewPrefs.horizontalMeasurementUnits = MeasurementUnits.POINTS;
        viewPrefs.verticalMeasurementUnits = MeasurementUnits.POINTS;
        viewPrefs.rulerOrigin = RulerOrigin.SPREAD_ORIGIN;
        try {
            return task();
        } finally {
            viewPrefs.horizontalMeasurementUnits = savedHorizontalUnits;
            viewPrefs.verticalMeasurementUnits = savedVerticalUnits;
            viewPrefs.rulerOrigin = savedRulerOrigin;
        }
    }

    /**
     * 選択全体の境界を返す
     * @param {PageItem[]} pageItems - 対象
     * @returns {number[]} [天, 左, 地, 右]
     */
    function getUnionBounds(pageItems) {
        var unionBounds = pageItems[0].visibleBounds.slice(0);
        for (var i = 1; i < pageItems.length; i++) {
            var itemBounds = pageItems[i].visibleBounds;
            unionBounds[0] = Math.min(unionBounds[0], itemBounds[0]);
            unionBounds[1] = Math.min(unionBounds[1], itemBounds[1]);
            unionBounds[2] = Math.max(unionBounds[2], itemBounds[2]);
            unionBounds[3] = Math.max(unionBounds[3], itemBounds[3]);
        }
        return unionBounds;
    }

    /**
     * オブジェクトを、ボタンの揃え方でいちばん近い線へ動かす
     * @param {PageItem[]} movableItems - 動かすオブジェクト
     * @param {string} actionKey - ACTION_AXIS_MODES のキー
     * @param {PaletteSettings} alignSettings - moveAsUnit と centerBasis を使う
     * @param {boolean} fromCenter - いったん中央（centerBasis の範囲）へ移したものとして揃えるなら true
     * @param {Document} doc - 対象ドキュメント
     * @returns {number} 揃える先が見つかったオブジェクトの数
     */
    function alignItemsToNearestLine(movableItems, actionKey, alignSettings, fromCenter, doc) {
        var axisModeKeys = ACTION_AXIS_MODES[actionKey];
        var linesBySpread = {};

        /**
         * スプレッドの線を、スプレッドと軸ごとに一度だけ集める
         * @param {Spread|MasterSpread} spread - 対象のスプレッド
         * @param {string} axis - "x" または "y"
         * @returns {number[]} 座標
         */
        function getTargetLines(spread, axis) {
            var cacheKey = spread.constructor.name + ":" + spread.id + ":" + axis;
            if (!linesBySpread[cacheKey]) linesBySpread[cacheKey] = collectTargetLines(spread, axis, doc);
            return linesBySpread[cacheKey];
        }

        /**
         * 境界から、揃える先までの移動量 [x, y] を求める（見つからない軸は 0）
         * @param {number[]} bounds - 境界 [天, 左, 地, 右]
         * @param {PageItem} spreadItem - 線やページを調べるスプレッドを決めるオブジェクト
         * @returns {number[]|null} [x, y]。どの軸にも揃える先が無ければ null
         */
        function findOffsets(bounds, spreadItem) {
            var spread = findParentSpread(spreadItem);
            if (!spread) return null;
            var centerPage = findCenterPage(spread, bounds);
            var centerBasisBounds = null;
            if (centerPage) centerBasisBounds = (alignSettings.centerBasis === "margin") ? getMarginAreaBounds(centerPage) : centerPage.bounds;
            var offsets = { x: 0, y: 0 };
            var isFound = false;
            var startBounds = bounds;
            /* Option＋クリックは、いったん中央へ移したものとして寄せ先を探す。中央への移動量を足し、境界も中央へずらしてから辺の移動量を出す（動かすのは最後の1回だけ）*/
            /* Option-click: treat the object as centered first. Add the centering offset and shift the bounds before the edge offsets; it moves only once at the end */
            if (fromCenter && centerBasisBounds) {
                offsets.x = findAxisOffset([], bounds, AXIS_MODES.centerX, centerBasisBounds);
                offsets.y = findAxisOffset([], bounds, AXIS_MODES.centerY, centerBasisBounds);
                startBounds = [bounds[0] + offsets.y, bounds[1] + offsets.x, bounds[2] + offsets.y, bounds[3] + offsets.x];
                isFound = true;
            }
            for (var i = 0; i < axisModeKeys.length; i++) {
                var axisMode = AXIS_MODES[axisModeKeys[i]];
                var axisLines = (axisMode.edgeIndex === null) ? [] : getTargetLines(spread, axisMode.axis);
                var axisOffset = findAxisOffset(axisLines, startBounds, axisMode, centerBasisBounds);
                if (axisOffset === null) continue;
                offsets[axisMode.axis] += axisOffset;
                isFound = true;
            }
            return isFound ? [offsets.x, offsets.y] : null;
        }

        /**
         * 移動量が誤差より大きいときだけ動かす
         * @param {PageItem} pageItem - 動かすオブジェクト
         * @param {number[]} offsets - [x, y]
         * @returns {void}
         */
        function moveByOffsets(pageItem, offsets) {
            if (Math.abs(offsets[0]) <= EDGE_TOLERANCE_POINTS && Math.abs(offsets[1]) <= EDGE_TOLERANCE_POINTS) return;
            pageItem.move(undefined, offsets);
        }

        if (alignSettings.moveAsUnit) {
            /* 選択全体の境界を基準に、全体を同じ量だけ動かす / Move everything by one amount, based on the bounds of the whole selection */
            var unitOffsets = findOffsets(getUnionBounds(movableItems), movableItems[0]);
            if (!unitOffsets) return 0;
            for (var j = 0; j < movableItems.length; j++) moveByOffsets(movableItems[j], unitOffsets);
            return movableItems.length;
        }

        var foundCount = 0;
        for (var k = 0; k < movableItems.length; k++) {
            var itemOffsets = findOffsets(movableItems[k].visibleBounds, movableItems[k]);
            if (!itemOffsets) continue;
            moveByOffsets(movableItems[k], itemOffsets);
            foundCount++;
        }
        return foundCount;
    }

    /**
     * 今の選択を、ボタンの揃え方でいちばん近い線へ揃える（パレットのボタン・矢印キーから呼ぶ）
     * @param {string} actionKey - ACTION_AXIS_MODES のキー
     * @param {boolean} fromCenter - いったん中央へ移してから揃えるなら true（Option／Alt キー）
     * @param {boolean} [moveAsUnit] - ［選択をまとめて移動］をこの値にする（省略時はパレットの設定）
     * @param {string} [centerBasis] - ［中央揃えの基準］を "page" か "margin" にする（省略時はパレットの設定）
     * @returns {void}
     */
    function runAlign(actionKey, fromCenter, moveAsUnit, centerBasis) {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var targetDocument = app.activeDocument;
        var movableItems = collectMovableItems(app.selection);
        if (movableItems.length === 0) {
            alert(getLabel("alert.noMovableItem"));
            return;
        }
        var sessionSettings = getSessionSettings();
        /* パレットの設定は書き換えずに、この1回だけ差し替える / Override for this run only, leaving the palette setting as it is */
        var alignSettings = {
            moveAsUnit: (typeof moveAsUnit === "boolean") ? moveAsUnit : sessionSettings.moveAsUnit,
            centerBasis: (centerBasis === "page" || centerBasis === "margin") ? centerBasis : sessionSettings.centerBasis
        };
        var foundCount = 0;
        /* 処理全体（定規の切り替えと復元を含む）を1回の取り消しにまとめる / One undo step, including the ruler switch and restore */
        app.doScript(function () {
            foundCount = runInPointsWithSpreadOrigin(targetDocument, function () {
                return alignItemsToNearestLine(movableItems, actionKey, alignSettings, fromCenter, targetDocument);
            });
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.align", { name: getLabel("action." + actionKey) }));

        if (foundCount === 0) alert(getLabel("alert.noLineAhead"));
    }

    // =========================================
    // アイコンの描画 / Icon drawing
    // =========================================

    /**
     * 塗りつぶした長方形を描く（newPath() を先に呼ばないとパスが重なり、塗りが線画になる）
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {number} x - 左端
     * @param {number} y - 上端
     * @param {number} width - 幅
     * @param {number} height - 高さ
     * @param {number[]} color - [r, g, b, a]
     * @returns {void}
     */
    function fillRect(graphics, x, y, width, height, color) {
        graphics.newPath();
        graphics.rectPath(x, y, width, height);
        graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
    }

    /**
     * 端に置くケイ線の座標を返す（太さ1の線がピクセル境界に乗るよう .5 の位置に取る）
     * @param {number} size - アイコンの一辺
     * @param {boolean} isStart - 左・上なら true、右・下なら false
     * @returns {number} ケイ線の中心の座標
     */
    function getEdgeRulePosition(size, isStart) {
        return isStart ? Math.round(size * ICON_RULE_OFFSET) + 0.5 : Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5;
    }

    /**
     * アイコンの中心の座標を返す（太さ1のケイ線がピクセル境界に乗るよう .5 の位置に取る）
     * @param {number} size - アイコンの一辺
     * @returns {number} 中心の座標
     */
    function getIconCenter(size) {
        return Math.round(size / 2) + 0.5;
    }

    /**
     * 長さを奇数に丸める（中心が .5 の位置にあるので、奇数なら両側が同じ幅になる）
     * @param {number} rawLength - 丸める前の長さ
     * @returns {number} 奇数の長さ
     */
    function roundToOddLength(rawLength) {
        var rounded = Math.round(rawLength);
        if (rounded % 2 !== 0) return rounded;
        return (rounded > 1) ? rounded - 1 : 1;
    }

    /**
     * アイコンに置くオブジェクトの大きさを返す（どのアイコンも同じ長方形）
     * @param {number} size - アイコンの一辺
     * @returns {number[]} [幅, 高さ]
     */
    function getBlockSize(size) {
        return [roundToOddLength(size * ICON_BLOCK_WIDTH), roundToOddLength(size * ICON_BLOCK_HEIGHT)];
    }

    /**
     * 端のケイ線（寄せ先を表す線）を1本描く
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {number} size - アイコンの一辺
     * @param {string} side - "top" / "bottom" / "left" / "right"
     * @param {number[]} color - [r, g, b, a]
     * @returns {number} 引いたケイ線の座標
     */
    function drawEdgeRule(graphics, size, side, color) {
        var ruleInset = Math.round(size * ICON_RULE_INSET);
        var ruleLength = size - ruleInset * 2;
        var rulePosition = getEdgeRulePosition(size, side === "top" || side === "left");
        if (side === "top" || side === "bottom") fillRect(graphics, ruleInset, rulePosition - 0.5, ruleLength, 1, color);
        else fillRect(graphics, rulePosition - 0.5, ruleInset, 1, ruleLength, color);
        return rulePosition;
    }

    /**
     * アイコンの中央を貫くケイ線を1本描く
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {number} size - アイコンの一辺
     * @param {boolean} isVerticalRule - 縦のケイ線なら true
     * @param {number[]} color - [r, g, b, a]
     * @returns {void}
     */
    function drawCenterRule(graphics, size, isVerticalRule, color) {
        var ruleInset = Math.round(size * ICON_RULE_INSET);
        var ruleLength = size - ruleInset * 2;
        var iconCenter = getIconCenter(size);
        if (isVerticalRule) fillRect(graphics, iconCenter - 0.5, ruleInset, 1, ruleLength, color);
        else fillRect(graphics, ruleInset, iconCenter - 0.5, ruleLength, 1, color);
    }

    /**
     * ボタンの絵（ケイ線と、そこへ寄せたオブジェクト）を描く。
     * 方向のボタンは端のケイ線、中央のボタンは中央を貫くケイ線（十字は2本）にオブジェクトを重ねる
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {string} actionKey - ACTION_AXIS_MODES のキー
     * @param {number} size - アイコンの一辺
     * @param {number[]} color - [r, g, b, a]
     * @returns {void}
     */
    function drawActionIcon(graphics, actionKey, size, color) {
        var blockSize = getBlockSize(size);
        var iconCenter = getIconCenter(size);
        /* ケイ線のない軸では中央に置く / Centered on any axis without an edge rule */
        var blockX = iconCenter - blockSize[0] / 2;
        var blockY = iconCenter - blockSize[1] / 2;

        if (actionKey === "center" || actionKey === "centerH") drawCenterRule(graphics, size, true, color);
        if (actionKey === "center" || actionKey === "centerV") drawCenterRule(graphics, size, false, color);

        var ruleSides = MOVE_ICON_RULES[actionKey] || [];
        var clearance = Math.round(size * ICON_RULE_CLEARANCE);
        for (var i = 0; i < ruleSides.length; i++) {
            var side = ruleSides[i];
            var rulePosition = drawEdgeRule(graphics, size, side, color);
            /* ケイ線は太さ1なので、その端から数えてすき間を取る / The rule is 1 unit thick, so measure from its edge */
            if (side === "top") blockY = rulePosition + 0.5 + clearance;
            if (side === "bottom") blockY = rulePosition - 0.5 - clearance - blockSize[1];
            if (side === "left") blockX = rulePosition + 0.5 + clearance;
            if (side === "right") blockX = rulePosition - 0.5 - clearance - blockSize[0];
        }

        /* ケイ線を先に描き、オブジェクトを上に重ねる（中央のケイ線がオブジェクトの下を通って見える）*/
        /* Rules first, block on top, so a center rule runs behind it */
        fillRect(graphics, blockX, blockY, blockSize[0], blockSize[1], color);
    }

    /**
     * ボタンの下地（地と、マウスオーバー中だけの枠）を描く
     * @param {Button} iconButton - 対象のボタン
     * @returns {void}
     */
    function drawButtonBase(iconButton) {
        var graphics = iconButton.graphics;
        var width = iconButton.size[0];
        var height = iconButton.size[1];
        fillRect(graphics, 0, 0, width, height, iconButton.isHover ? ICON_HOVER_BG : ICON_BASE_BG);
        if (!iconButton.isHover) return;
        /* 0.5 ずらすと1pxの線がピクセル境界に乗ってくっきり出る / The 0.5 offset keeps the 1px border crisp */
        graphics.newPath();
        graphics.rectPath(0.5, 0.5, width - 1, height - 1);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, ICON_BORDER_COLOR, 1));
    }

    // =========================================
    // パレット / Palette
    // =========================================

    /**
     * ボタンを描き直す（notify が使えない環境では隠して再表示する）
     * @param {Button} iconButton - 対象のボタン
     * @returns {void}
     */
    function redrawIconButton(iconButton) {
        try {
            iconButton.notify("onDraw");
        } catch (e) {
            iconButton.hide();
            iconButton.show();
        }
    }

    /**
     * Option（Alt）キーが押されているかを返す
     * @returns {boolean} 押されていれば true（取得できない環境では false）
     */
    function isAltPressed() {
        try {
            return ScriptUI.environment.keyboardState.altKey === true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 揃え方のボタンを1つ足す（画像なしの iconbutton はクリックが届かないことがあるので button に描く）
     * @param {Group} rowGroup - 足す先の行
     * @param {string} actionKey - ACTION_AXIS_MODES のキー
     * @returns {Button} 足したボタン
     */
    function addActionButton(rowGroup, actionKey) {
        var actionButton = rowGroup.add("button", undefined, "");
        actionButton.minimumSize = actionButton.preferredSize = actionButton.maximumSize = [ICON_SIZE, ICON_SIZE];
        var isCenterAction = (actionKey === "center" || actionKey === "centerH" || actionKey === "centerV");
        /* 中心はページの中央に揃えるので、揃える先の説明は辺のボタンだけに付ける / Only the edge buttons list the target lines; centers go to the page center */
        actionButton.helpTip = getLabel("tooltip." + actionKey) +
            (isCenterAction ? "" : "\n" + getLabel("tooltip.edgeSuffix") + "\n" + getLabel("tooltip.optionFromCenter") + "\n" + getLabel("tooltip.targetSuffix"));
        actionButton.isHover = false;
        actionButton.onDraw = function () {
            drawButtonBase(actionButton);
            drawActionIcon(actionButton.graphics, actionKey, ICON_SIZE, ICON_COLOR);
        };
        actionButton.addEventListener("mouseover", function () { actionButton.isHover = true; redrawIconButton(actionButton); });
        actionButton.addEventListener("mouseout", function () { actionButton.isHover = false; redrawIconButton(actionButton); });
        actionButton.onClick = function () { runAlign(actionKey, isAltPressed()); };
        return actionButton;
    }

    /**
     * 3×3 の十字のボタンと、その下に中央揃えのボタン2つを並べる
     * @param {Window} targetWindow - 足す先のパレット
     * @returns {void}
     */
    function addActionButtonGrid(targetWindow) {
        var gridGroup = targetWindow.add("group");
        gridGroup.orientation = "column";
        gridGroup.alignment = ["center", "top"];
        gridGroup.alignChildren = ["center", "top"];
        gridGroup.spacing = CROSS_GAP;

        for (var i = 0; i < MOVE_BUTTON_ROWS.length; i++) {
            var crossRowGroup = gridGroup.add("group");
            setupRow(crossRowGroup, "center", CROSS_GAP);
            for (var j = 0; j < MOVE_BUTTON_ROWS[i].length; j++) addActionButton(crossRowGroup, MOVE_BUTTON_ROWS[i][j]);
        }

        var centerRowGroup = gridGroup.add("group");
        setupRow(centerRowGroup, "center", CROSS_GAP);
        /* 十字の4段目に見えないよう、上に少し余白を取る / A little room so it does not read as a fourth row of the cross */
        centerRowGroup.margins = [0, CENTER_ROW_TOP, 0, 0];
        for (var k = 0; k < CENTER_BUTTON_KEYS.length; k++) addActionButton(centerRowGroup, CENTER_BUTTON_KEYS[k]);
    }

    /**
     * 揃え方のボタンを並べたパレットを表示する
     * @returns {void}
     */
    function showAlignPalette() {
        var sessionSettings = getSessionSettings();
        var alignPalette = new Window("palette", getLabel("palette.title"));
        setupWindow(alignPalette);

        addActionButtonGrid(alignPalette);

        var centerBasisPanel = alignPalette.add("panel", undefined, getLabel("panel.centerBasis"));
        setupPanel(centerBasisPanel, 6);
        centerBasisPanel.helpTip = getLabel("tooltip.centerBasis");
        /* ページとマージンは縦に並べる（パネルに直接置くので同じ親の中で排他になる）/ Page and Margins stacked; both in the panel, so they stay exclusive */
        var pageBasisRadio = centerBasisPanel.add("radiobutton", undefined, getLabel("radio.page"));
        var marginBasisRadio = centerBasisPanel.add("radiobutton", undefined, getLabel("radio.margin"));
        pageBasisRadio.value = (sessionSettings.centerBasis !== "margin");
        marginBasisRadio.value = (sessionSettings.centerBasis === "margin");
        pageBasisRadio.onClick = function () { sessionSettings.centerBasis = "page"; };
        marginBasisRadio.onClick = function () { sessionSettings.centerBasis = "margin"; };

        var moveAsUnitCheckbox = alignPalette.add("checkbox", undefined, getLabel("checkbox.moveAsUnit"));
        moveAsUnitCheckbox.helpTip = getLabel("tooltip.moveAsUnit");
        moveAsUnitCheckbox.value = sessionSettings.moveAsUnit;
        moveAsUnitCheckbox.onClick = function () { sessionSettings.moveAsUnit = moveAsUnitCheckbox.value; };

        var buttonRow = addButtonRow(alignPalette);
        var btnClose = buttonRow.rightGroup.add("button", undefined, getLabel("button.close"));
        btnClose.onClick = function () { alignPalette.close(); };
        alignRightOnlyButtonRow(buttonRow);

        /* パレットは Esc で閉じないので自分で受ける。矢印キーは上下左右のボタンと同じ / Palettes ignore Esc, so handle it; arrow keys act like the edge buttons */
        alignPalette.addEventListener("keydown", function (event) {
            if (event.keyName === "Escape") {
                event.preventDefault();
                alignPalette.close();
                return;
            }
            var actionKey = MOVE_ACTION_BY_KEY_NAME[event.keyName];
            if (!actionKey) return;
            event.preventDefault();
            runAlign(actionKey, event.altKey === true);
        });

        alignPalette.onClose = function () {
            sessionSettings.paletteLocation = [alignPalette.location[0], alignPalette.location[1]];
            $.global[PALETTE_GLOBAL_KEY] = null;
        };

        if (sessionSettings.paletteLocation) alignPalette.location = sessionSettings.paletteLocation;

        /* パレットは show() から即座に戻るので、参照を $.global に残す / Keep a reference: palettes return from show() at once */
        $.global[PALETTE_GLOBAL_KEY] = alignPalette;
        addVersionFooter(alignPalette);
        alignPalette.show();
    }

    /* SW_ACTION があれば、パレットを出さずにそのボタンの揃え方で実行する（パレットの設定は同じエンジンなので共通）/ With SW_ACTION, run that button without the palette (the palette settings are shared through the same engine) */
    /* 「, SW_…」と書くと差し替え側が複数宣言とみなして止まるので、いったん別の変数に受ける / Writing ", SW_…" looks like a multiple declaration to launchers, so copy it first */
    if (SW_ACTION !== "") {
        var isFromCenter = (SW_FROM_CENTER === true);
        /* SW_MOVE_AS_UNIT は文字列 "true"／"false"。空ならパレットの設定 / A string "true" or "false"; empty follows the palette */
        var moveAsUnitOverride = (SW_MOVE_AS_UNIT === "") ? undefined : (String(SW_MOVE_AS_UNIT) === "true");
        var centerBasisOverride = String(SW_CENTER_BASIS);
        if (ACTION_AXIS_MODES.hasOwnProperty(SW_ACTION)) runAlign(SW_ACTION, isFromCenter, moveAsUnitOverride, centerBasisOverride);
        else alert(getLabel("alert.unknownAction", { name: SW_ACTION }));
        return;
    }

    /* 前回のパレットが開いていれば閉じる（開いたままだと修正前のコードが動き続ける） / Close a palette left open, which would keep running the old code */
    if ($.global[PALETTE_GLOBAL_KEY]) $.global[PALETTE_GLOBAL_KEY].close();

    showAlignPalette();

})();
