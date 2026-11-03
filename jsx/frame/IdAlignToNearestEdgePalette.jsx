#target indesign
#targetengine "IdAlignToNearestEdgePalette"

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdAlignToNearestEdgePalette
Name=移動方向でいちばん近いガイド・マージン・ページ・裁ち落としに揃えるパレット
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.0.0
Release-Date=2026-11-03
Description-BEGIN
選択したオブジェクトを、パレットのボタンで選んだ方向へ動かし、その方向でいちばん近いガイド・マージンガイド・ページのエッジ・裁ち落としに辺を揃えます。

Moves the selected objects in the direction chosen on the palette and aligns their edge to the nearest guide, margin guide, page edge, or bleed edge in that direction.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択したオブジェクトを、パレットのボタンで選んだ方向へ動かし、その方向でいちばん近いガイド・マージンガイド・ページのエッジ・裁ち落としに辺を揃えます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAlignToNearestEdgePalette.md

### Overview

Moves the selected objects in the direction chosen on the palette and aligns their edge to the nearest guide, margin guide, page edge, or bleed edge in that direction.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAlignToNearestEdgePalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdAlignToNearestEdgePalette";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-03";                   /* 更新日 / last updated */

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
    // 移動方向 / Directions
    // =========================================

    /* axis は線の向き、edgeIndex は visibleBounds [天, 左, 地, 右] の添字、sign は座標が増える向きなら 1 */
    /* axis: which lines to use; edgeIndex: index into visibleBounds [top, left, bottom, right]; sign: 1 when the coordinate grows */
    var DIRECTIONS = {
        left:  { axis: "x", edgeIndex: 1, sign: -1 },
        right: { axis: "x", edgeIndex: 3, sign: 1 },
        up:    { axis: "y", edgeIndex: 0, sign: -1 },
        down:  { axis: "y", edgeIndex: 2, sign: 1 }
    };

    /* パレットで押した矢印キー（ScriptUI の keyName）と方向 / Arrow keys pressed on the palette (ScriptUI keyName) */
    var ARROW_KEY_DIRECTIONS = { Left: "left", Right: "right", Up: "up", Down: "down" };

    // =========================================
    // セッション記憶 / Session memory
    // =========================================

    /* パレットの参照と設定を InDesign の終了まで $.global に残す / Keeps the palette and settings on $.global until InDesign quits */
    var PALETTE_GLOBAL_KEY = "idAlignToNearestEdgePalette";
    var SETTINGS_GLOBAL_KEY = "idAlignToNearestEdgeSettings";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIRECTION_BUTTON_SIZE = [40, 28];  /* 方向ボタンの大きさ / size of a direction button */
    var DIRECTION_SPACING = 4;             /* 方向ボタンどうしの間隔 / gap between direction buttons */

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
        panel: {
            direction: { ja: "移動方向", en: "Direction" }
        },
        button: {
            left: { ja: "←", en: "←" },
            right: { ja: "→", en: "→" },
            up: { ja: "↑", en: "↑" },
            down: { ja: "↓", en: "↓" },
            close: { ja: "閉じる", en: "Close" }
        },
        checkbox: {
            moveAsUnit: { ja: "選択をまとめて移動", en: "Move selection as a unit" }
        },
        tooltip: {
            left: { ja: "左方向でいちばん近いガイド・マージンガイド・ページのエッジ・裁ち落としに左辺を揃える（←キー）", en: "Aligns the left edge to the nearest guide, margin, page edge, or bleed on the left (Left Arrow)" },
            right: { ja: "右方向でいちばん近いガイド・マージンガイド・ページのエッジ・裁ち落としに右辺を揃える（→キー）", en: "Aligns the right edge to the nearest guide, margin, page edge, or bleed on the right (Right Arrow)" },
            up: { ja: "上方向でいちばん近いガイド・マージンガイド・ページのエッジ・裁ち落としに上辺を揃える（↑キー）", en: "Aligns the top edge to the nearest guide, margin, page edge, or bleed above (Up Arrow)" },
            down: { ja: "下方向でいちばん近いガイド・マージンガイド・ページのエッジ・裁ち落としに下辺を揃える（↓キー）", en: "Aligns the bottom edge to the nearest guide, margin, page edge, or bleed below (Down Arrow)" },
            moveAsUnit: { ja: "移動方向の先頭にあるオブジェクトを基準に、相対位置を保ったまま動かす", en: "Moves everything by the same amount, based on the object furthest in the direction" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noMovableItem: { ja: "移動できるオブジェクトを選択してください（ロックしたもの・アンカー付きオブジェクトは対象外）。", en: "Please select objects that can be moved (locked and anchored objects are skipped)." },
            noLineAhead: { ja: "移動方向に揃える先が見つかりませんでした。", en: "No target was found in that direction." }
        },
        undo: {
            left: { ja: "左方向の線に揃える", en: "Align Left to Nearest Edge" },
            right: { ja: "右方向の線に揃える", en: "Align Right to Nearest Edge" },
            up: { ja: "上方向の線に揃える", en: "Align Top to Nearest Edge" },
            down: { ja: "下方向の線に揃える", en: "Align Bottom to Nearest Edge" }
        }
    };

    // =========================================
    // 設定 / Settings
    // =========================================

    /**
     * @typedef {Object} PaletteSettings
     * @property {boolean} moveAsUnit - 選択をまとめて移動するなら true
     * @property {number[]|null} paletteLocation - 前回閉じたときのパレットの位置 [x, y]
     */

    /**
     * 前回の設定を返す（無ければ既定の設定）。返したオブジェクトへの変更はそのまま次回に残る
     * @returns {PaletteSettings} 設定
     */
    function getSessionSettings() {
        if (!$.global[SETTINGS_GLOBAL_KEY]) $.global[SETTINGS_GLOBAL_KEY] = { moveAsUnit: false, paletteLocation: null };
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
     * 各ページのマージンガイドの座標を集める（見開きの左ページは左右の内側・外側を入れ替える）
     * @param {Spread|MasterSpread} spread - 対象のスプレッド
     * @param {string} axis - "x" または "y"
     * @returns {number[]} 座標
     */
    function collectMarginLines(spread, axis) {
        var lines = [];
        var pages = spread.pages;
        for (var i = 0; i < pages.length; i++) {
            var pageBounds = pages[i].bounds;
            var marginPrefs = pages[i].marginPreferences;
            if (axis === "y") {
                lines.push(pageBounds[0] + marginPrefs.top, pageBounds[2] - marginPrefs.bottom);
                continue;
            }
            var isLeftHand = (pages[i].side === PageSideOptions.LEFT_HAND);
            lines.push(pageBounds[1] + (isLeftHand ? marginPrefs.right : marginPrefs.left));
            lines.push(pageBounds[3] - (isLeftHand ? marginPrefs.left : marginPrefs.right));
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
     * オブジェクトを移動方向に動かし、その方向の辺をいちばん近い線に揃える
     * @param {PageItem[]} movableItems - 動かすオブジェクト
     * @param {string} directionKey - "left" / "right" / "up" / "down"
     * @param {boolean} moveAsUnit - まとめて動かすなら true
     * @param {Document} doc - 対象ドキュメント
     * @returns {number} 動かしたオブジェクトの数
     */
    function alignItemsToNearestLine(movableItems, directionKey, moveAsUnit, doc) {
        var direction = DIRECTIONS[directionKey];
        var linesBySpread = {};

        /**
         * スプレッドの線を、スプレッドごとに一度だけ集める
         * @param {PageItem} pageItem - スプレッドを調べるオブジェクト
         * @returns {number[]} 座標（スプレッドが無ければ空）
         */
        function getTargetLines(pageItem) {
            var spread = findParentSpread(pageItem);
            if (!spread) return [];
            var spreadKey = spread.constructor.name + ":" + spread.id;
            if (!linesBySpread[spreadKey]) linesBySpread[spreadKey] = collectTargetLines(spread, direction.axis, doc);
            return linesBySpread[spreadKey];
        }

        /**
         * 移動方向の軸に沿って動かす
         * @param {PageItem} pageItem - 動かすオブジェクト
         * @param {number} distance - 動かす量（pt）
         * @returns {void}
         */
        function moveAlongAxis(pageItem, distance) {
            pageItem.move(undefined, (direction.axis === "x") ? [distance, 0] : [0, distance]);
        }

        if (moveAsUnit) {
            /* 移動方向の先頭にあるオブジェクトの辺を基準に、全体を同じ量だけ動かす / Move everything by one amount, based on the object furthest ahead */
            var leadingItem = movableItems[0];
            var leadingEdge = movableItems[0].visibleBounds[direction.edgeIndex];
            for (var i = 1; i < movableItems.length; i++) {
                var itemEdge = movableItems[i].visibleBounds[direction.edgeIndex];
                if ((itemEdge - leadingEdge) * direction.sign > 0) {
                    leadingEdge = itemEdge;
                    leadingItem = movableItems[i];
                }
            }
            var unitTarget = findNearestLineAhead(getTargetLines(leadingItem), leadingEdge, direction.sign);
            if (unitTarget === null) return 0;
            for (var j = 0; j < movableItems.length; j++) moveAlongAxis(movableItems[j], unitTarget - leadingEdge);
            return movableItems.length;
        }

        var movedCount = 0;
        for (var k = 0; k < movableItems.length; k++) {
            var edgePosition = movableItems[k].visibleBounds[direction.edgeIndex];
            var targetPosition = findNearestLineAhead(getTargetLines(movableItems[k]), edgePosition, direction.sign);
            if (targetPosition === null) continue;
            moveAlongAxis(movableItems[k], targetPosition - edgePosition);
            movedCount++;
        }
        return movedCount;
    }

    /**
     * 今の選択を、指定した方向のいちばん近い線へ揃える（パレットのボタン・矢印キーから呼ぶ）
     * @param {string} directionKey - "left" / "right" / "up" / "down"
     * @returns {void}
     */
    function runAlign(directionKey) {
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
        var moveAsUnit = getSessionSettings().moveAsUnit;
        var movedCount = 0;
        /* 処理全体（定規の切り替えと復元を含む）を1回の取り消しにまとめる / One undo step, including the ruler switch and restore */
        app.doScript(function () {
            movedCount = runInPointsWithSpreadOrigin(targetDocument, function () {
                return alignItemsToNearestLine(movableItems, directionKey, moveAsUnit, targetDocument);
            });
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo." + directionKey));

        if (movedCount === 0) alert(getLabel("alert.noLineAhead"));
    }

    // =========================================
    // パレット / Palette
    // =========================================

    /**
     * 方向ボタンを足す
     * @param {Group} rowGroup - 足す先の行
     * @param {string} directionKey - "left" / "right" / "up" / "down"
     * @returns {Button} 足したボタン
     */
    function addDirectionButton(rowGroup, directionKey) {
        var btnDirection = rowGroup.add("button", undefined, getLabel("button." + directionKey));
        btnDirection.preferredSize = DIRECTION_BUTTON_SIZE;
        btnDirection.helpTip = getLabel("tooltip." + directionKey);
        btnDirection.onClick = function () { runAlign(directionKey); };
        return btnDirection;
    }

    /**
     * 方向ボタンを十字に並べたパレットを表示する
     * @returns {void}
     */
    function showAlignPalette() {
        var sessionSettings = getSessionSettings();
        var alignPalette = new Window("palette", getLabel("palette.title") + " " + SCRIPT_VERSION);
        setupWindow(alignPalette);

        var directionPanel = alignPalette.add("panel", undefined, getLabel("panel.direction"));
        setupPanel(directionPanel, DIRECTION_SPACING);
        directionPanel.alignChildren = ["center", "top"];

        var upRowGroup = directionPanel.add("group");
        setupRow(upRowGroup, "center", DIRECTION_SPACING);
        addDirectionButton(upRowGroup, "up");

        var middleRowGroup = directionPanel.add("group");
        setupRow(middleRowGroup, "center", DIRECTION_SPACING);
        addDirectionButton(middleRowGroup, "left");
        /* 十字の中央はあける / Leave the center of the cross empty */
        middleRowGroup.add("group").preferredSize = DIRECTION_BUTTON_SIZE;
        addDirectionButton(middleRowGroup, "right");

        var downRowGroup = directionPanel.add("group");
        setupRow(downRowGroup, "center", DIRECTION_SPACING);
        addDirectionButton(downRowGroup, "down");

        var moveAsUnitCheckbox = alignPalette.add("checkbox", undefined, getLabel("checkbox.moveAsUnit"));
        moveAsUnitCheckbox.helpTip = getLabel("tooltip.moveAsUnit");
        moveAsUnitCheckbox.value = sessionSettings.moveAsUnit;
        moveAsUnitCheckbox.onClick = function () { sessionSettings.moveAsUnit = moveAsUnitCheckbox.value; };

        var buttonRow = addButtonRow(alignPalette);
        var btnClose = buttonRow.rightGroup.add("button", undefined, getLabel("button.close"));
        btnClose.onClick = function () { alignPalette.close(); };
        alignRightOnlyButtonRow(buttonRow);

        /* パレットは Esc で閉じないので自分で受ける。矢印キーは方向ボタンと同じ / Palettes ignore Esc, so handle it; arrow keys act like the buttons */
        alignPalette.addEventListener("keydown", function (event) {
            if (event.keyName === "Escape") {
                event.preventDefault();
                alignPalette.close();
                return;
            }
            var directionKey = ARROW_KEY_DIRECTIONS[event.keyName];
            if (!directionKey) return;
            event.preventDefault();
            runAlign(directionKey);
        });

        alignPalette.onClose = function () {
            sessionSettings.paletteLocation = [alignPalette.location[0], alignPalette.location[1]];
            $.global[PALETTE_GLOBAL_KEY] = null;
        };

        if (sessionSettings.paletteLocation) alignPalette.location = sessionSettings.paletteLocation;

        /* パレットは show() から即座に戻るので、参照を $.global に残す / Keep a reference: palettes return from show() at once */
        $.global[PALETTE_GLOBAL_KEY] = alignPalette;
        alignPalette.show();
    }

    /* 前回のパレットが開いていれば閉じる（開いたままだと修正前のコードが動き続ける） / Close a palette left open, which would keep running the old code */
    if ($.global[PALETTE_GLOBAL_KEY]) $.global[PALETTE_GLOBAL_KEY].close();

    showAlignPalette();

})();
