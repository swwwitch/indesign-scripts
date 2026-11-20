#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdApplyTocStylesByPart
Name=目次の項目を部ごとの派生スタイルに当て直す
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.0.2
Release-Date=2026-10-30
Description-BEGIN
目次を更新したあとに実行し、目次の項目を部（部の扉の段落スタイル）ごとに派生の段落スタイル（toc-h1.blue など）へ当て直します。
本文の段落スタイルは変更しません。

Run after updating the table of contents to reassign its entries to per-part variant paragraph styles (such as toc-h1.blue).
The paragraph styles in the body text are left unchanged.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

目次を更新したあとに実行し、目次の項目を部（部の扉の段落スタイル）ごとに派生の段落スタイル（toc-h1.blue など）へ当て直します。
本文の段落スタイルは変更しません。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdApplyTocStylesByPart.md

### Overview

Run after updating the table of contents to reassign its entries to per-part variant paragraph styles (such as toc-h1.blue).
The paragraph styles in the body text are left unchanged.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdApplyTocStylesByPart.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdApplyTocStylesByPart";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "目次の項目を部ごとの派生スタイルに当て直す"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdApplyTocStylesByPart.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdApplyTocStylesByPart.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdApplyTocStylesByPart.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 派生スタイル名の区切り（「h1.blue」の「.」） / Separator between a style name and its variant suffix ("." in "h1.blue") */
    var VARIANT_SEPARATOR = ".";

    /* 部の見出しとして表示する文字数 / Characters of each part heading shown in the dialog */
    var PART_TITLE_MAX_LENGTH = 24;

    /* InDesign がマーカー位置に入れる文字。ExtendScript では "﻿" と書くと空文字列になる
       / The character InDesign puts at marker positions; "﻿" becomes an empty string in ExtendScript */
    var MARKER_CHAR = String.fromCharCode(0xFEFF);

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
            title: { ja: "目次の書式を部ごとに当て直す", en: "Apply TOC Formats by Part" }
        },
        panel: {
            toc: { ja: "目次", en: "Table of Contents" },
            parts: { ja: "部ごとの派生スタイル", en: "Variant per Part" },
            options: { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            tocStyle: { ja: "目次スタイル", en: "TOC style" },
            partStyle: { ja: "部の段落スタイル", en: "Part style" }
        },
        dropdown: {
            none: { ja: "なし", en: "None" }
        },
        checkbox: {
            removeBreaks: { ja: "目次の強制改行を削除", en: "Remove forced line breaks in the TOC" }
        },
        format: {
            partRow: { ja: "p.%1 %2", en: "p. %1 %2" }
        },
        tooltip: {
            partStyle: { ja: "この段落から、次の同じスタイルの段落の手前までを1つの部とみなします", en: "Each paragraph in this style starts a part that runs up to the next one" },
            parts: { ja: "初期値は、部の中ですでに使われている派生スタイルの多数決です", en: "The default is the variant used most often within the part" },
            removeBreaks: { ja: "目次スタイルの［強制改行を削除］がオンでも残る強制改行を削除します", en: "Removes the forced line breaks that remain even when the TOC style's Remove Forced Line Break is on" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        message: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noTocStyle: { ja: "目次スタイルがありません。", en: "The document has no TOC styles." },
            noEntries: { ja: "目次スタイル「%1」に段落スタイルが登録されていません。", en: "The TOC style \"%1\" includes no paragraph styles." },
            noVariants: {
                ja: "目次の書式の派生スタイルが見つかりません。\n「%1」のように、目次の書式の段落スタイル名に「%2接尾辞」を付けた段落スタイルを作ってから実行してください。",
                en: "No variant TOC format styles were found.\nCreate paragraph styles named like \"%1\" (the TOC format style name plus \"%2suffix\") and run the script again."
            },
            noParts: { ja: "本文に段落スタイル「%1」の段落がありません。", en: "No paragraphs in the body use the paragraph style \"%1\"." },
            noToc: { ja: "目次スタイル「%1」の書式を使った目次が見つかりません。", en: "No table of contents using the formats of the TOC style \"%1\" was found." },
            nothingToDo: { ja: "変更する段落・強制改行はありません。", en: "There is nothing to change." },
            done: { ja: "目次の %1 段落のスタイルを変更し、強制改行を %2 個削除しました。", en: "Changed the style of %1 TOC paragraphs and removed %2 forced line breaks." },
            missingStyles: { ja: "見つからないため元の書式のままにした段落スタイル：%1", en: "Missing paragraph styles (the base format was kept): %1" },
            unmatchedParts: { ja: "本文と一致しなかった部の見出し（順番で対応させました）：%1", en: "Part headings that did not match the body (matched by order): %1" }
        },
        undo: {
            apply: { ja: "目次の書式を部ごとに当て直す", en: "Apply TOC Formats by Part" }
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

    var FIELD_LABEL_WIDTH = 110;   /* 項目名の幅 / width of field labels */
    var FIELD_INPUT_WIDTH = 200;   /* ドロップダウンの幅 / width of field dropdowns */
    var PART_LABEL_WIDTH  = 250;   /* 部の見出しの幅 / width of part headings */
    var VARIANT_WIDTH     = 120;   /* 派生スタイルのドロップダウンの幅 / width of variant dropdowns */

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
    // 文字列 / Text
    // =========================================

    /**
     * 段落の文字列から、マーカー位置の文字と末尾の改行を除き、改行を空白にする（表示用）
     * @param {string} text - 段落の contents
     * @returns {string} 表示用の文字列
     */
    function toDisplayText(text) {
        return text.split(MARKER_CHAR).join("").replace(/\r$/, "").replace(/[\r\n]/g, " ");
    }

    /**
     * 見出しを比べるための文字列にする（マーカー位置の文字・空白・改行を除く）
     * @param {string} text - 段落の contents
     * @returns {string} 比較用の文字列
     */
    function toMatchKey(text) {
        return text.split(MARKER_CHAR).join("").replace(/[\s　]/g, "");
    }

    /**
     * 長い文字列を切り詰める
     * @param {string} text - 文字列
     * @param {number} maxLength - 最大の文字数
     * @returns {string} 切り詰めた文字列
     */
    function truncateText(text, maxLength) {
        return (text.length > maxLength) ? text.substr(0, maxLength) + "..." : text;
    }

    // =========================================
    // ドキュメントの解析 / Document analysis
    // =========================================

    /**
     * everyItem() のプロパティを必ず配列で受け取る（要素が1つだと配列にならない）
     * @param {*} value - everyItem().xxx の値
     * @returns {Array} 配列
     */
    function toArray(value) {
        return (value instanceof Array) ? value : [value];
    }

    /**
     * 段落スタイルを名前で引ける表にする（グループ内も含む。同名は先のものを使う）
     * @param {Document} doc - ドキュメント
     * @returns {Object} 名前 → ParagraphStyle
     */
    function buildStyleMap(doc) {
        var styleMap = {};
        var allStyles = doc.allParagraphStyles;
        for (var i = 0; i < allStyles.length; i++) {
            if (!styleMap.hasOwnProperty(allStyles[i].name)) styleMap[allStyles[i].name] = allStyles[i];
        }
        return styleMap;
    }

    /**
     * 目次エントリーの書式の段落スタイル名。［同一スタイル］などで段落スタイルでなければ空文字
     * @param {TOCStyleEntry} entry - 目次スタイルの項目
     * @returns {string} 段落スタイル名
     */
    function getEntryFormatName(entry) {
        try {
            var formatStyle = entry.formatStyle;
            if (formatStyle && typeof formatStyle === "object" && formatStyle.constructor.name === "ParagraphStyle") return formatStyle.name;
        } catch (e) {}
        return "";
    }

    /**
     * 段落の載っているページを返す。ページ上にない（親ページ・ペーストボード・オーバーセット）なら null
     * @param {Paragraph} para - 段落
     * @returns {Page|null} ページ
     */
    function getParagraphPage(para) {
        try {
            var frames = para.parentTextFrames;
            if (frames.length === 0 || !frames[0].parentPage) return null;
            var page = frames[0].parentPage;
            return (page.parent.constructor.name === "Spread") ? page : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * 目次スタイルと部の段落スタイルから、見出し・部・派生スタイルの候補を調べる
     * @param {Document} doc - ドキュメント
     * @param {TOCStyle} tocStyle - 目次スタイル
     * @param {string|null} partBaseName - 部の段落スタイル名（null なら最上位レベルの項目）
     * @returns {Object} { styleMap, baseEntries, partBaseName, candidates, tocStories, records, parts }
     */
    function analyzeDocument(doc, tocStyle, partBaseName) {
        var analysis = { styleMap: buildStyleMap(doc), baseEntries: [], partBaseName: "", candidates: [], tocStories: [], records: [], parts: [] };
        var styleMap = analysis.styleMap;

        /* 項目のうち、ほかの項目の派生（「h1」に対する「h1.blue」）ではないものを元の見出しとする
           / Entries that are not variants of another entry are the base headings */
        var entries = tocStyle.tocStyleEntries.everyItem().getElements();
        var entryNameSet = {};
        for (var i = 0; i < entries.length; i++) entryNameSet[entries[i].name] = true;
        for (i = 0; i < entries.length; i++) {
            var separatorIndex = entries[i].name.lastIndexOf(VARIANT_SEPARATOR);
            if (separatorIndex > 0 && entryNameSet[entries[i].name.substr(0, separatorIndex)]) continue;
            analysis.baseEntries.push({ name: entries[i].name, entry: entries[i], level: entries[i].level, formatName: getEntryFormatName(entries[i]) });
        }
        if (analysis.baseEntries.length === 0) return analysis;

        var baseSet = {};
        var topEntry = analysis.baseEntries[0];
        for (i = 0; i < analysis.baseEntries.length; i++) {
            baseSet[analysis.baseEntries[i].name] = analysis.baseEntries[i];
            if (analysis.baseEntries[i].level < topEntry.level) topEntry = analysis.baseEntries[i];
        }
        analysis.partBaseName = (partBaseName && baseSet[partBaseName]) ? partBaseName : topEntry.name;

        /* 派生の候補は、目次の書式の段落スタイルに「.接尾辞」を付けたものから集める
           / Variant suffixes come from the TOC format styles ("toc-h1" → "toc-h1.blue") */
        var candidateSet = {};
        var tocFormatSet = {};
        for (i = 0; i < analysis.baseEntries.length; i++) {
            var formatName = analysis.baseEntries[i].formatName;
            if (!formatName) continue;
            tocFormatSet[formatName] = true;
            var prefix = formatName + VARIANT_SEPARATOR;
            for (var styleName in styleMap) {
                if (!styleMap.hasOwnProperty(styleName) || styleName.indexOf(prefix) !== 0 || styleName.length === prefix.length) continue;
                tocFormatSet[styleName] = true;
                var suffix = styleName.substr(prefix.length);
                if (!candidateSet[suffix]) {
                    candidateSet[suffix] = true;
                    analysis.candidates.push(suffix);
                }
            }
        }

        var stories = doc.stories.everyItem().getElements();
        for (var s = 0; s < stories.length; s++) {
            if (stories[s].paragraphs.length === 0) continue;
            var paraStyles = toArray(stories[s].paragraphs.everyItem().appliedParagraphStyle);

            /* 目次の書式を使った段落があるストーリーは目次とみなし、本文の見出しから外す
               / A story using a TOC format style is a table of contents, not body text */
            var isTocStory = false;
            for (var t = 0; t < paraStyles.length; t++) {
                if (tocFormatSet[paraStyles[t].name]) { isTocStory = true; break; }
            }
            if (isTocStory) {
                analysis.tocStories.push(stories[s]);
                continue;
            }

            for (var p = 0; p < paraStyles.length; p++) {
                var parsed = parseVariantName(paraStyles[p].name, baseSet, candidateSet);
                if (!parsed) continue;
                var para = stories[s].paragraphs[p];
                var page = getParagraphPage(para);
                if (!page) continue;
                analysis.records.push({ para: para, styleName: paraStyles[p].name, base: parsed.base, suffix: parsed.suffix, offset: page.documentOffset, pageName: page.name, part: null });
            }
        }

        buildParts(analysis);
        return analysis;
    }

    /**
     * 段落スタイル名を、元の見出しと派生の接尾辞に分ける
     * @param {string} styleName - 段落スタイル名
     * @param {Object} nameSet - 元の名前の集合
     * @param {Object} candidateSet - 接尾辞の候補の集合
     * @returns {{base: string, suffix: string}|null} 当てはまらなければ null
     */
    function parseVariantName(styleName, nameSet, candidateSet) {
        if (nameSet[styleName]) return { base: styleName, suffix: "" };
        var separatorIndex = styleName.lastIndexOf(VARIANT_SEPARATOR);
        if (separatorIndex <= 0) return null;
        var baseName = styleName.substr(0, separatorIndex);
        var suffix = styleName.substr(separatorIndex + VARIANT_SEPARATOR.length);
        return (nameSet[baseName] && candidateSet[suffix]) ? { base: baseName, suffix: suffix } : null;
    }

    /**
     * 部の段落をページ順に並べて部を作り、見出しを部に振り分け、部ごとの派生の初期値を多数決で決める
     * @param {Object} analysis - analyzeDocument の途中結果
     * @returns {void}
     */
    function buildParts(analysis) {
        var parts = [];
        for (var i = 0; i < analysis.records.length; i++) {
            var record = analysis.records[i];
            if (record.base !== analysis.partBaseName) continue;
            parts.push({ start: record.offset, pageName: record.pageName, title: toDisplayText(record.para.contents), matchKey: toMatchKey(record.para.contents), votes: {}, variant: "" });
        }
        /* 部は多くても数十なので比較関数つきの sort で足りる / Only a few dozen parts at most, so a comparator sort is fine */
        parts.sort(function (a, b) { return a.start - b.start; });
        for (i = 0; i < parts.length; i++) parts[i].index = i;

        for (i = 0; i < analysis.records.length; i++) {
            var rec = analysis.records[i];
            for (var k = parts.length - 1; k >= 0; k--) {
                if (parts[k].start <= rec.offset) { rec.part = parts[k]; break; }
            }
            if (rec.part && rec.suffix) rec.part.votes[rec.suffix] = (rec.part.votes[rec.suffix] || 0) + 1;
        }

        for (i = 0; i < parts.length; i++) {
            var bestCount = 0;
            for (var suffix in parts[i].votes) {
                if (parts[i].votes.hasOwnProperty(suffix) && parts[i].votes[suffix] > bestCount) {
                    parts[i].variant = suffix;
                    bestCount = parts[i].votes[suffix];
                }
            }
        }
        analysis.parts = parts;
    }

    // =========================================
    // 目次の変更 / Changing the TOC
    // =========================================

    /**
     * ストーリー内の強制改行を削除する（検索設定は元に戻す）
     * @param {Story} story - ストーリー
     * @returns {number} 削除した数
     */
    function removeForcedLineBreaks(story) {
        var savedFind = app.findGrepPreferences.properties;
        var savedChange = app.changeGrepPreferences.properties;
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "\\n";
        app.changeGrepPreferences.changeTo = "";
        var removedCount = story.changeGrep().length;
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.properties = savedFind;
        app.changeGrepPreferences.properties = savedChange;
        return removedCount;
    }

    /**
     * 目次の部の見出しに当たる本文の部を返す。文字列が一致する部を優先し、無ければまだ使っていない部を順に当てる
     * @param {Array} parts - analyzeDocument の parts
     * @param {string} contents - 目次の段落の contents
     * @param {Object} usedParts - 使った部の index → true
     * @param {string[]} unmatchedTitles - 一致しなかった見出しを足す先
     * @returns {Object|null} 部。余りが無ければ null
     */
    function matchPart(parts, contents, usedParts, unmatchedTitles) {
        /* ノンブルの前のタブより後ろは比べない / Ignore the page number after the tab */
        var matchKey = toMatchKey(contents.split("\t")[0]);
        var i;
        for (i = 0; i < parts.length; i++) {
            if (!usedParts[i] && parts[i].matchKey === matchKey) {
                usedParts[i] = true;
                return parts[i];
            }
        }
        unmatchedTitles.push(toDisplayText(contents.split("\t")[0]));
        for (i = 0; i < parts.length; i++) {
            if (!usedParts[i]) {
                usedParts[i] = true;
                return parts[i];
            }
        }
        return null;
    }

    /**
     * 目次の段落を、部ごとの派生の書式に当て直す
     * @param {Object} analysis - analyzeDocument の戻り値
     * @param {Object} options - { variants: string[], removeBreaks: boolean }
     * @returns {Object} 結果
     */
    function applyTocStyles(analysis, options) {
        var result = { paragraphs: 0, removedBreaks: 0, missingNames: [], unmatchedTitles: [] };
        var styleMap = analysis.styleMap;
        var formatSet = {};
        var partFormatName = "";
        for (var b = 0; b < analysis.baseEntries.length; b++) {
            var formatName = analysis.baseEntries[b].formatName;
            if (!formatName) continue;
            formatSet[formatName] = true;
            if (analysis.baseEntries[b].name === analysis.partBaseName) partFormatName = formatName;
        }
        var candidateSet = {};
        for (var c = 0; c < analysis.candidates.length; c++) candidateSet[analysis.candidates[c]] = true;
        var missingSet = {};

        for (var s = 0; s < analysis.tocStories.length; s++) {
            var story = analysis.tocStories[s];
            var paras = story.paragraphs.everyItem().getElements();
            var paraStyles = toArray(story.paragraphs.everyItem().appliedParagraphStyle);
            var usedParts = {};
            var currentPart = null;
            for (var p = 0; p < paras.length; p++) {
                var parsed = parseVariantName(paraStyles[p].name, formatSet, candidateSet);
                if (!parsed) continue;
                if (parsed.base === partFormatName) currentPart = matchPart(analysis.parts, paras[p].contents, usedParts, result.unmatchedTitles);

                var variant = currentPart ? options.variants[currentPart.index] : "";
                var targetName = parsed.base;
                if (variant) {
                    var variantName = parsed.base + VARIANT_SEPARATOR + variant;
                    if (styleMap.hasOwnProperty(variantName)) targetName = variantName;
                    else if (!missingSet[variantName]) {
                        missingSet[variantName] = true;
                        result.missingNames.push(variantName);
                    }
                }
                if (paraStyles[p].name === targetName) continue;
                /* 手で付けた書式は残す / Keep local overrides */
                paras[p].applyParagraphStyle(styleMap[targetName], false);
                result.paragraphs++;
            }
            if (options.removeBreaks) result.removedBreaks += removeForcedLineBreaks(story);
        }
        return result;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名つきの行を作る
     * @param {Panel|Group} parent - 行を足す先
     * @param {string} labelKey - 項目名のラベルのパス
     * @returns {Group} 行
     */
    function addFieldRow(parent, labelKey) {
        var rowGroup = parent.add("group");
        setupRow(rowGroup);
        var fieldLabel = rowGroup.add("statictext", undefined, labelText(labelKey));
        fieldLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        fieldLabel.justify = "right";
        return rowGroup;
    }

    /**
     * 設定ダイアログボックスを表示する
     * @param {Document} doc - ドキュメント
     * @param {Object} analysis - analyzeDocument の戻り値
     * @param {Object} dialogState - { tocStyleIndex, partBaseName }。選び直したときはここを書き換える
     * @param {boolean} canRun - false なら［OK］を押せなくする（目次スタイルの選び直しだけできる）
     * @returns {Object|null} options。キャンセルなら null、目次スタイルか部の段落スタイルを選び直したら { rebuild: true }
     */
    function showOptionsDialog(doc, analysis, dialogState, canRun) {
        var dialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(dialog);

        var tocPanel = dialog.add("panel", undefined, getLabel("panel.toc"));
        setupPanel(tocPanel, 6);
        var ddTocStyle = addFieldRow(tocPanel, "fieldLabel.tocStyle").add("dropdownlist", undefined, toArray(doc.tocStyles.everyItem().name));
        ddTocStyle.preferredSize.width = FIELD_INPUT_WIDTH;
        ddTocStyle.selection = dialogState.tocStyleIndex;

        var partStyleRow = addFieldRow(tocPanel, "fieldLabel.partStyle");
        var baseNames = [];
        var partStyleIndex = 0;
        for (var i = 0; i < analysis.baseEntries.length; i++) {
            baseNames.push(analysis.baseEntries[i].name);
            if (analysis.baseEntries[i].name === analysis.partBaseName) partStyleIndex = i;
        }
        var ddPartStyle = partStyleRow.add("dropdownlist", undefined, baseNames);
        ddPartStyle.preferredSize.width = FIELD_INPUT_WIDTH;
        if (baseNames.length > 0) ddPartStyle.selection = partStyleIndex;
        ddPartStyle.helpTip = getLabel("tooltip.partStyle");

        var partsPanel = dialog.add("panel", undefined, getLabel("panel.parts"));
        setupPanel(partsPanel, 6);
        partsPanel.helpTip = getLabel("tooltip.parts");
        var variantItems = [getLabel("dropdown.none")].concat(analysis.candidates);
        var variantDropdowns = [];
        for (i = 0; i < analysis.parts.length; i++) {
            var part = analysis.parts[i];
            var partRow = partsPanel.add("group");
            setupRow(partRow);
            var partLabel = partRow.add("statictext", undefined, getLabel("format.partRow", [part.pageName, truncateText(part.title, PART_TITLE_MAX_LENGTH)]));
            partLabel.preferredSize.width = PART_LABEL_WIDTH;
            var ddVariant = partRow.add("dropdownlist", undefined, variantItems);
            ddVariant.preferredSize.width = VARIANT_WIDTH;
            var variantIndex = 0;
            for (var c = 0; c < analysis.candidates.length; c++) {
                if (analysis.candidates[c] === part.variant) variantIndex = c + 1;
            }
            ddVariant.selection = variantIndex;
            variantDropdowns.push(ddVariant);
        }

        var optionsPanel = dialog.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel, 6);
        var chkRemoveBreaks = optionsPanel.add("checkbox", undefined, getLabel("checkbox.removeBreaks"));
        chkRemoveBreaks.helpTip = getLabel("tooltip.removeBreaks");
        chkRemoveBreaks.value = true;

        var buttonRow = addButtonRow(dialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        btnOK.enabled = canRun;
        alignRightOnlyButtonRow(buttonRow);

        /* 目次スタイル・部の段落スタイルを選び直したら、調べ直して開き直す / Re-analyze and reopen after another TOC style or part style is chosen */
        var REBUILD_CODE = 3;
        ddTocStyle.onChange = function () {
            dialogState.tocStyleIndex = ddTocStyle.selection.index;
            dialogState.partBaseName = null;
            dialog.close(REBUILD_CODE);
        };
        ddPartStyle.onChange = function () {
            dialogState.partBaseName = ddPartStyle.selection.text;
            dialog.close(REBUILD_CODE);
        };

        addVersionFooter(dialog);
        var dialogResult = dialog.show();
        if (dialogResult === REBUILD_CODE) return { rebuild: true };
        if (dialogResult !== 1) return null;

        var options = { variants: [], removeBreaks: chkRemoveBreaks.value };
        for (i = 0; i < variantDropdowns.length; i++) {
            options.variants.push(variantDropdowns[i].selection.index === 0 ? "" : analysis.candidates[variantDropdowns[i].selection.index - 1]);
        }
        return options;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 解析結果に問題があればメッセージを返す
     * @param {TOCStyle} tocStyle - 目次スタイル
     * @param {Object} analysis - analyzeDocument の戻り値
     * @returns {string} メッセージ。問題なければ空文字
     */
    function getAnalysisProblem(tocStyle, analysis) {
        if (analysis.baseEntries.length === 0) return getLabel("message.noEntries", [tocStyle.name]);
        if (analysis.candidates.length === 0) {
            var sampleFormat = analysis.baseEntries[0].formatName || "toc-h1";
            return getLabel("message.noVariants", [sampleFormat + VARIANT_SEPARATOR + "blue", VARIANT_SEPARATOR]);
        }
        if (analysis.parts.length === 0) return getLabel("message.noParts", [analysis.partBaseName]);
        if (analysis.tocStories.length === 0) return getLabel("message.noToc", [tocStyle.name]);
        return "";
    }

    /**
     * 結果のメッセージを組み立てる
     * @param {Object} result - applyTocStyles の戻り値
     * @returns {string} メッセージ
     */
    function buildResultMessage(result) {
        var lines = [getLabel("message.done", [result.paragraphs, result.removedBreaks])];
        if (result.missingNames.length > 0) lines.push(getLabel("message.missingStyles", [result.missingNames.join(", ")]));
        if (result.unmatchedTitles.length > 0) lines.push(getLabel("message.unmatchedParts", [result.unmatchedTitles.join(", ")]));
        return lines.join("\n\n");
    }

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
        if (doc.tocStyles.length === 0) {
            alert(getLabel("message.noTocStyle"));
            return;
        }

        var dialogState = { tocStyleIndex: 0, partBaseName: null };
        var tocStyle, analysis, options;
        while (true) {
            tocStyle = doc.tocStyles[dialogState.tocStyleIndex];
            analysis = analyzeDocument(doc, tocStyle, dialogState.partBaseName);
            var problem = getAnalysisProblem(tocStyle, analysis);
            if (problem) {
                alert(problem);
                /* 目次スタイルがほかにあれば、選び直せるようにダイアログボックスは開く / With other TOC styles available, still open the dialog so another can be chosen */
                if (doc.tocStyles.length === 1) return;
            }
            options = showOptionsDialog(doc, analysis, dialogState, !problem);
            if (!options) return;
            if (!options.rebuild) break;
        }

        var result;
        /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
        app.doScript(function () { result = applyTocStyles(analysis, options); }, ScriptLanguage.JAVASCRIPT, undefined,
            UndoModes.ENTIRE_SCRIPT, getLabel("undo.apply") + " " + SCRIPT_VERSION);

        if (result.paragraphs === 0 && result.removedBreaks === 0 && result.unmatchedTitles.length === 0) {
            alert(getLabel("message.nothingToDo"));
            return;
        }
        alert(buildResultMessage(result));
    }

    main();

})();
