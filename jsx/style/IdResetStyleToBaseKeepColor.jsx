#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdResetStyleToBaseKeepColor
Name=派生スタイルのオーバーライドをカラー以外削除
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.0.1
Release-Date=2026-10-21
Description-BEGIN
ほかのスタイルを基準にした段落スタイル・文字スタイルから、カラー以外のオーバーライドを削除して基準スタイルの設定に戻します。
文字カラー・段落背景色・下線や境界線のカラーなどは残します。

Removes every override except colors from paragraph and character styles that are based on another style, so they follow their base style again.
Character colors, paragraph shading, and the colors of underlines, rules and borders are kept.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

ほかのスタイルを基準にした段落スタイル・文字スタイルから、カラー以外のオーバーライドを削除して基準スタイルの設定に戻します。
文字カラー・段落背景色・下線や境界線のカラーなどは残します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdResetStyleToBaseKeepColor.md

### Overview

Removes every override except colors from paragraph and character styles that are based on another style, so they follow their base style again.
Character colors, paragraph shading, and the colors of underlines, rules and borders are kept.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdResetStyleToBaseKeepColor.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdResetStyleToBaseKeepColor";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-21";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-05";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "派生スタイルのオーバーライドをカラー以外削除"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdResetStyleToBaseKeepColor.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdResetStyleToBaseKeepColor.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdResetStyleToBaseKeepColor.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 残すカラーの分類。上から順に判定し、どれかに当てはまる項目はオーバーライドを残す
       / Color categories to keep; a property matching any checked category keeps its override */
    var COLOR_CATEGORIES = [
        { key: "characterColor", pattern: /^(fillColor|fillTint|overprintFill|strokeColor|strokeTint|strokeWeight|overprintStroke|gradient(Fill|Stroke)(Angle|Length|Start))$/ },
        { key: "paragraphShading", pattern: /^paragraphShading(On|Color|Tint|Overprint)$/ },
        { key: "underlineStrike", pattern: /^(underline|strikeThrough)(Gap)?(Color|Tint|Overprint)$/ },
        { key: "ruleBorder", pattern: /^(ruleAbove|ruleBelow|paragraphBorder)(Gap)?(Color|Tint|Overprint)$/ },
        { key: "otherColor", pattern: /(Color|Tint)$|Overprint/ }
    ];

    /* 書式ではないのでオーバーライドとして扱わない項目 / Properties that are not formatting and are never reset */
    var NON_FORMAT_KEYS = {
        id: true, index: true, label: true, name: true, parent: true, basedOn: true, nextStyle: true,
        imported: true, isValid: true, properties: true, styleUniqueId: true,
        styleExportTagMaps: true, emitCss: true, includeClass: true
    };

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
            title: { ja: "スタイルを基準に戻す（カラーは残す）", en: "Reset Styles to Base (Keep Colors)" }
        },
        panel: {
            target: { ja: "対象", en: "Target" },
            keep: { ja: "残すカラー", en: "Colors to Keep" }
        },
        checkbox: {
            paragraphStyles: { ja: "段落スタイル", en: "Paragraph styles" },
            characterStyles: { ja: "文字スタイル", en: "Character styles" },
            includeBasicParagraph: { ja: "［基本段落］を基準にしたスタイルも含める", en: "Include styles based on [Basic Paragraph]" },
            characterColor: { ja: "文字カラー（塗り・線）", en: "Character color (fill and stroke)" },
            paragraphShading: { ja: "段落背景色", en: "Paragraph shading" },
            underlineStrike: { ja: "下線・打ち消し線のカラー", en: "Underline and strikethrough colors" },
            ruleBorder: { ja: "段落境界線・段落囲みのカラー", en: "Paragraph rule and border colors" },
            otherColor: { ja: "その他のカラー（圏点・ルビなど）", en: "Other colors (kenten, ruby, etc.)" }
        },
        summary: {
            count: { ja: "%1 個のスタイルで %2 項目", en: "%2 properties in %1 styles" }
        },
        tooltip: {
            includeBasicParagraph: { ja: "オフのときは、［基本段落］を直接の基準にしたスタイルを処理しません", en: "When off, styles based directly on [Basic Paragraph] are left as they are" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        message: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            nothingToDo: { ja: "削除するオーバーライドはありません。", en: "There are no overrides to remove." },
            done: { ja: "%1 個のスタイルで %2 項目のオーバーライドを削除しました。", en: "Removed %2 overrides from %1 styles." },
            failed: { ja: "削除できなかった項目：%1", en: "Properties that could not be reset: %1" }
        },
        undo: {
            reset: { ja: "スタイルを基準に戻す", en: "Reset Styles to Base" }
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
        "v": [".....", ".....", "#...#", "#...#", "#...#", ".#.#.", "..#..", ".....", "....."],
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

    // =========================================
    // スタイルの判定 / Style inspection
    // =========================================

    /**
     * 値を比べられる文字列にする（スウォッチやフォントは種類と id、配列は要素ごと）
     * @param {*} value - プロパティの値
     * @returns {string} 比較用の文字列
     */
    function toComparableKey(value) {
        if (value === null || value === undefined) return String(value);
        if (value instanceof Array) {
            var itemKeys = [];
            for (var i = 0; i < value.length; i++) itemKeys.push(toComparableKey(value[i]));
            return "[" + itemKeys.join(",") + "]";
        }
        if (typeof value === "object") {
            if (value.constructor.name === "Object") return value.toSource();
            /* スウォッチなどは id、フォントは name で見分ける / Swatches have an id, fonts are told apart by name */
            var identity;
            try { identity = value.id; } catch (e) {}
            if (identity === undefined) {
                try { identity = value.name; } catch (e2) {}
            }
            return value.constructor.name + "#" + String(identity);
        }
        return typeof value + ":" + String(value);
    }

    /**
     * プロパティ名が、チェックされた残すカラーの分類に当てはまるか
     * @param {string} propertyKey - プロパティ名
     * @param {Object} keepFlags - 分類の key → true/false
     * @returns {boolean} 残すなら true
     */
    function isKeptColorKey(propertyKey, keepFlags) {
        for (var i = 0; i < COLOR_CATEGORIES.length; i++) {
            if (COLOR_CATEGORIES[i].pattern.test(propertyKey)) return !!keepFlags[COLOR_CATEGORIES[i].key];
        }
        return false;
    }

    /**
     * スタイルの基準を返す。基準なし（［段落スタイルなし］［なし］）なら null
     * @param {ParagraphStyle|CharacterStyle} style - スタイル
     * @param {ParagraphStyle|CharacterStyle} noneStyle - ［段落スタイルなし］または［なし］
     * @returns {ParagraphStyle|CharacterStyle|null} 基準スタイル
     */
    function getBaseStyle(style, noneStyle) {
        /* ［段落スタイルなし］［なし］などのルートスタイルは basedOn を読むと例外（Error 516） / Reading basedOn on a root style throws Error 516 */
        var baseStyle;
        try { baseStyle = style.basedOn; } catch (e) { return null; }
        if (!baseStyle || typeof baseStyle === "string") return null;
        if (baseStyle.id === noneStyle.id) return null;
        return baseStyle;
    }

    /**
     * 基準からの深さ（基準スタイルを先に処理するため）
     * @param {ParagraphStyle|CharacterStyle} style - スタイル
     * @param {ParagraphStyle|CharacterStyle} noneStyle - ［段落スタイルなし］または［なし］
     * @returns {number} 深さ
     */
    function getStyleDepth(style, noneStyle) {
        var depth = 0;
        var baseStyle = getBaseStyle(style, noneStyle);
        while (baseStyle && depth < 100) {
            depth++;
            baseStyle = getBaseStyle(baseStyle, noneStyle);
        }
        return depth;
    }

    /**
     * 段落スタイルのうち、基準と値が異なる（＝オーバーライドしている）カラー以外の項目を集める
     * @param {ParagraphStyle} style - 段落スタイル
     * @param {ParagraphStyle} baseStyle - 基準スタイル
     * @param {Object} keepFlags - 残すカラーの分類
     * @returns {string[]} プロパティ名
     */
    function collectParagraphOverrideKeys(style, baseStyle, keepFlags) {
        var styleProps = style.properties;
        var baseProps = baseStyle.properties;
        var overrideKeys = [];
        for (var propertyKey in styleProps) {
            if (NON_FORMAT_KEYS[propertyKey] || isKeptColorKey(propertyKey, keepFlags)) continue;
            if (!(propertyKey in baseProps)) continue;
            if (toComparableKey(styleProps[propertyKey]) !== toComparableKey(baseProps[propertyKey])) overrideKeys.push(propertyKey);
        }
        return overrideKeys;
    }

    /**
     * 文字スタイルのうち、値が設定されている（NOTHING ではない）カラー以外の項目を集める
     * @param {CharacterStyle} style - 文字スタイル
     * @param {Object} keepFlags - 残すカラーの分類
     * @returns {string[]} プロパティ名
     */
    function collectCharacterOverrideKeys(style, keepFlags) {
        var styleProps = style.properties;
        var overrideKeys = [];
        for (var propertyKey in styleProps) {
            if (NON_FORMAT_KEYS[propertyKey] || isKeptColorKey(propertyKey, keepFlags)) continue;
            if (styleProps[propertyKey] == NothingEnum.NOTHING) continue;
            overrideKeys.push(propertyKey);
        }
        return overrideKeys;
    }

    /**
     * 処理するスタイルと、外すオーバーライドの一覧を作る（基準スタイルが先）
     * @param {Document} doc - ドキュメント
     * @param {Object} options - { paragraph, character, includeBasicParagraph, keep }
     * @returns {Array} { style, baseStyle, isParagraph, keys } の配列
     */
    function buildResetPlan(doc, options) {
        var plan = [];
        if (options.paragraph) appendStylesToPlan(plan, doc.allParagraphStyles, doc.paragraphStyles[0], true, options, doc.paragraphStyles[1]);
        if (options.character) appendStylesToPlan(plan, doc.allCharacterStyles, doc.characterStyles[0], false, options, null);
        return plan;
    }

    /**
     * スタイルの配列から、基準を持つものを深さ順に plan へ足す
     * @param {Array} plan - 足し先
     * @param {Array} styles - allParagraphStyles / allCharacterStyles
     * @param {ParagraphStyle|CharacterStyle} noneStyle - ［段落スタイルなし］または［なし］
     * @param {boolean} isParagraph - 段落スタイルなら true
     * @param {Object} options - buildResetPlan と同じ
     * @param {ParagraphStyle|null} basicParagraphStyle - ［基本段落］（文字スタイルでは null）
     * @returns {void}
     */
    function appendStylesToPlan(plan, styles, noneStyle, isParagraph, options, basicParagraphStyle) {
        var entries = [];
        for (var i = 0; i < styles.length; i++) {
            var style = styles[i];
            var baseStyle = getBaseStyle(style, noneStyle);
            if (!baseStyle) continue;
            if (basicParagraphStyle && !options.includeBasicParagraph && baseStyle.id === basicParagraphStyle.id) continue;
            var keys = isParagraph ? collectParagraphOverrideKeys(style, baseStyle, options.keep) : collectCharacterOverrideKeys(style, options.keep);
            if (keys.length === 0) continue;
            entries.push({ style: style, baseStyle: baseStyle, isParagraph: isParagraph, keys: keys, depth: getStyleDepth(style, noneStyle) });
        }
        /* 比較関数つきの sort は遅いので、深さごとに振り分ける / Bucket by depth instead of a comparator sort */
        var buckets = [];
        for (var j = 0; j < entries.length; j++) {
            if (!buckets[entries[j].depth]) buckets[entries[j].depth] = [];
            buckets[entries[j].depth].push(entries[j]);
        }
        for (var depth = 0; depth < buckets.length; depth++) {
            if (!buckets[depth]) continue;
            for (var k = 0; k < buckets[depth].length; k++) plan.push(buckets[depth][k]);
        }
    }

    /**
     * plan の件数を数える
     * @param {Array} plan - buildResetPlan の戻り値
     * @returns {{styles: number, keys: number}} スタイル数と項目数
     */
    function countPlan(plan) {
        var keyCount = 0;
        for (var i = 0; i < plan.length; i++) keyCount += plan[i].keys.length;
        return { styles: plan.length, keys: keyCount };
    }

    // =========================================
    // オーバーライドの削除 / Removing overrides
    // =========================================

    /**
     * plan にしたがってオーバーライドを外す。
     * 段落スタイルは基準と同じ値を入れると継承に戻り、文字スタイルは NOTHING を入れると未設定に戻る
     * @param {Array} plan - buildResetPlan の戻り値
     * @returns {{styles: number, keys: number, failedKeys: string[]}} 結果
     */
    function applyResetPlan(plan) {
        var result = { styles: 0, keys: 0, failedKeys: [] };
        var failedKeySet = {};
        for (var i = 0; i < plan.length; i++) {
            var entry = plan[i];
            /* 基準スタイルを先に戻したので、値は今の基準から読み直す / Read the base again since it may have been reset just before */
            var baseProps = entry.isParagraph ? entry.baseStyle.properties : null;
            var resetCount = 0;
            for (var j = 0; j < entry.keys.length; j++) {
                var propertyKey = entry.keys[j];
                try {
                    entry.style[propertyKey] = entry.isParagraph ? baseProps[propertyKey] : NothingEnum.NOTHING;
                    resetCount++;
                } catch (e) {
                    if (!failedKeySet[propertyKey]) {
                        failedKeySet[propertyKey] = true;
                        result.failedKeys.push(propertyKey);
                    }
                }
            }
            if (resetCount > 0) result.styles++;
            result.keys += resetCount;
        }
        return result;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 設定ダイアログボックスを表示する
     * @param {Document} doc - ドキュメント
     * @returns {Object|null} options。キャンセルなら null
     */
    function showOptionsDialog(doc) {
        var dialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(dialog);

        var targetPanel = dialog.add("panel", undefined, getLabel("panel.target"));
        setupPanel(targetPanel, 6);
        var chkParagraph = targetPanel.add("checkbox", undefined, getLabel("checkbox.paragraphStyles"));
        var chkCharacter = targetPanel.add("checkbox", undefined, getLabel("checkbox.characterStyles"));
        var chkBasicParagraph = targetPanel.add("checkbox", undefined, getLabel("checkbox.includeBasicParagraph"));
        chkBasicParagraph.helpTip = getLabel("tooltip.includeBasicParagraph");
        chkParagraph.value = true;
        chkCharacter.value = true;
        chkBasicParagraph.value = false;

        var keepPanel = dialog.add("panel", undefined, getLabel("panel.keep"));
        setupPanel(keepPanel, 6);
        var keepCheckboxes = {};
        for (var i = 0; i < COLOR_CATEGORIES.length; i++) {
            var categoryKey = COLOR_CATEGORIES[i].key;
            keepCheckboxes[categoryKey] = keepPanel.add("checkbox", undefined, getLabel("checkbox." + categoryKey));
            keepCheckboxes[categoryKey].value = true;
        }

        var summaryText = dialog.add("statictext", undefined, "");
        summaryText.alignment = ["fill", "center"];

        var buttonRow = addButtonRow(dialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /**
         * ダイアログボックスの状態を options にまとめる
         * @returns {Object} options
         */
        function readOptions() {
            var keepFlags = {};
            for (var categoryKey in keepCheckboxes) keepFlags[categoryKey] = keepCheckboxes[categoryKey].value;
            return {
                paragraph: chkParagraph.value,
                character: chkCharacter.value,
                includeBasicParagraph: chkBasicParagraph.value,
                keep: keepFlags
            };
        }

        /**
         * 件数の表示と［OK］の可否を更新する
         * @returns {void}
         */
        function updateSummary() {
            chkBasicParagraph.enabled = chkParagraph.value;
            var counts = countPlan(buildResetPlan(doc, readOptions()));
            summaryText.text = getLabel("summary.count", [counts.styles, counts.keys]);
            btnOK.enabled = counts.keys > 0;
        }

        chkParagraph.onClick = updateSummary;
        chkCharacter.onClick = updateSummary;
        chkBasicParagraph.onClick = updateSummary;
        for (var keepKey in keepCheckboxes) keepCheckboxes[keepKey].onClick = updateSummary;
        updateSummary();

        addVersionFooter(dialog);
        if (dialog.show() !== 1) return null;
        return readOptions();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * エントリポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("message.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var options = showOptionsDialog(doc);
        if (!options) return;

        var plan = buildResetPlan(doc, options);
        if (plan.length === 0) {
            alert(getLabel("message.nothingToDo"));
            return;
        }

        var result;
        /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
        app.doScript(function () { result = applyResetPlan(plan); }, ScriptLanguage.JAVASCRIPT, undefined,
            UndoModes.ENTIRE_SCRIPT, getLabel("undo.reset") + " " + SCRIPT_VERSION);

        var message = getLabel("message.done", [result.styles, result.keys]);
        if (result.failedKeys.length > 0) message += "\n\n" + getLabel("message.failed", [result.failedKeys.join(", ")]);
        alert(message);
    }

    main();

})();
