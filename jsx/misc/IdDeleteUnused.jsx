#target indesign

/*

### 概要

現在のドキュメントで使われていないスタイル（段落・文字・オブジェクト・表・セル）、親ページ、空ページ、スウォッチ、合成フォントを削除します。
削除する前に一覧で確認でき、残したい項目はチェックを外せます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdDeleteUnused.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n879f09b72808

### Overview

Deletes unused styles (paragraph, character, object, table, cell), parent pages, empty pages, swatches, and composite fonts in the active document.
The candidates are listed before deletion so you can uncheck anything you want to keep.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdDeleteUnused.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdDeleteUnused";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdDeleteUnused.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdDeleteUnused.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n879f09b72808"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

    /* 確認リスト / Confirmation list */
    var CONFIRM_LIST_SIZE     = [560, 320];            /* リストの寸法 [幅,高さ] / list size */
    var CONFIRM_COLUMN_WIDTHS = [170, 380];            /* 列幅（種類・名前） / column widths (kind, name) */

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
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
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
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
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

    var LABELS = {
        dialog: {
            title:   { ja: "未使用項目の削除", en: "Delete Unused Items" },
            confirm: { ja: "削除する項目の確認", en: "Confirm Items to Delete" }
        },
        panel: {
            styleTargets: { ja: "スタイル", en: "Styles" },
            pageTargets:  { ja: "ページ", en: "Pages" },
            otherTargets: { ja: "その他", en: "Other" },
            options:      { ja: "オプション", en: "Options" }
        },
        checkbox: {
            paragraphStyle:    { ja: "段落スタイル", en: "Paragraph styles" },
            characterStyle:    { ja: "文字スタイル", en: "Character styles" },
            objectStyle:       { ja: "オブジェクトスタイル", en: "Object styles" },
            tableStyle:        { ja: "表スタイル", en: "Table styles" },
            cellStyle:         { ja: "セルスタイル", en: "Cell styles" },
            swatch:            { ja: "スウォッチ", en: "Swatches" },
            masterSpread:      { ja: "親ページ", en: "Parent pages" },
            emptyPage:         { ja: "空ページ", en: "Empty pages" },
            compositeFont:     { ja: "合成フォント", en: "Composite fonts" },
            removeEmptyGroups: { ja: "空のスタイルグループを削除", en: "Delete empty style groups" }
        },
        kindName: {
            paragraphStyle: { ja: "段落スタイル", en: "Paragraph style" },
            characterStyle: { ja: "文字スタイル", en: "Character style" },
            objectStyle:    { ja: "オブジェクトスタイル", en: "Object style" },
            tableStyle:     { ja: "表スタイル", en: "Table style" },
            cellStyle:      { ja: "セルスタイル", en: "Cell style" },
            swatch:         { ja: "スウォッチ", en: "Swatch" },
            masterSpread:   { ja: "親ページ", en: "Parent page" },
            emptyPage:      { ja: "空ページ", en: "Empty page" },
            compositeFont:  { ja: "合成フォント", en: "Composite font" },
            styleGroup:     { ja: "スタイルグループ", en: "Style group" }
        },
        columnTitle: {
            kind: { ja: "種類", en: "Kind" },
            name: { ja: "名前", en: "Name" }
        },
        tooltip: {
            paragraphStyle: {
                ja: "テキストにも、他のスタイルの「基準」「次のスタイル」、オブジェクトスタイル、セルスタイル、目次スタイル、脚注の設定にも使われていない段落スタイルを削除します。［ ］で囲まれた既定のスタイルは残します。",
                en: "Deletes paragraph styles not used in text, as Based On / Next Style of other styles, or in object, cell, TOC styles and footnote options. Default styles in [ ] are kept."
            },
            characterStyle: {
                ja: "テキストにも、他のスタイルの「基準」、先頭文字スタイル・正規表現スタイル・行スタイル、箇条書き、ドロップキャップ、目次スタイル、脚注の設定にも使われていない文字スタイルを削除します。［なし］は残します。",
                en: "Deletes character styles not used in text, as Based On, in nested / GREP / line styles, bullets and numbering, drop caps, TOC styles, or footnote options. [None] is kept."
            },
            objectStyle: {
                ja: "どのオブジェクトにも、他のスタイルの「基準」にも、新規オブジェクトの既定にも使われていないオブジェクトスタイルを削除します。［ ］で囲まれた既定のスタイルは残します。",
                en: "Deletes object styles not applied to any object, not used as Based On, and not set as the default for new objects. Default styles in [ ] are kept."
            },
            tableStyle: {
                ja: "どの表にも、他の表スタイルの「基準」にも使われていない表スタイルを削除します。［ ］で囲まれた既定のスタイルは残します。",
                en: "Deletes table styles not applied to any table and not used as Based On. Default styles in [ ] are kept."
            },
            cellStyle: {
                ja: "どのセルにも、他のセルスタイルの「基準」、表スタイルの領域（ヘッダー行・本文行など）にも使われていないセルスタイルを削除します。［なし］は残します。",
                en: "Deletes cell styles not applied to any cell, not used as Based On, and not used in table style regions (header rows, body rows, etc.). [None] is kept."
            },
            swatch: {
                ja: "使われていないスウォッチを削除します。グラデーションの分岐点や濃淡の元になっているカラー、名前のないカラー、削除できない既定のスウォッチは残します。",
                en: "Deletes unused swatches. Colors used in gradient stops or as the base of tints, unnamed colors, and default swatches that cannot be deleted are kept."
            },
            masterSpread: {
                ja: "どのページにも、他の親ページの「基準」にも使われていない親ページを削除します。",
                en: "Deletes parent pages not applied to any page and not used as the basis of another parent page."
            },
            emptyPage: {
                ja: "オブジェクトが1つも無いページを削除します。親ページのオブジェクトしか無いページも空とみなします。ドキュメントの最後の1ページは残します。",
                en: "Deletes pages with no objects. Pages showing only parent page objects count as empty. The last remaining page is kept."
            },
            compositeFont: {
                ja: "テキストにも、段落スタイル・文字スタイル、テキストの既定にも使われていない合成フォントを削除します。",
                en: "Deletes composite fonts not used in text, in paragraph or character styles, or in text defaults."
            },
            removeEmptyGroups: {
                ja: "段落・文字・オブジェクト・表・セルのスタイルグループのうち、中身が空のもの（削除で空になるものを含む）を削除します。",
                en: "Deletes paragraph, character, object, table, and cell style groups that are empty, including those emptied by this deletion."
            },
            optionClick: {
                ja: "option（Alt）+クリック：この項目だけをON。もう一度押すとすべてをON",
                en: "Option (Alt)-click: check only this item. Do it again to check all"
            },
            confirmList: {
                ja: "行をダブルクリックするとチェックを切り替えます。チェックを外した項目は削除しません。",
                en: "Double-click a row to toggle its check. Unchecked items are kept."
            }
        },
        button: {
            cancel:     { ja: "キャンセル", en: "Cancel" },
            ok:         { ja: "OK", en: "OK" },
            remove:     { ja: "削除", en: "Delete" },
            checkAll:   { ja: "すべてON", en: "Check All" },
            uncheckAll: { ja: "すべてOFF", en: "Uncheck All" }
        },
        message: {
            confirmCount: {
                ja: "%1個の未使用項目が見つかりました。チェックを外した項目は残します。",
                en: "Found %1 unused items. Unchecked items will be kept."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noneFound:  { ja: "未使用の項目は見つかりませんでした。", en: "No unused items were found." },
            result:     { ja: "削除しました。", en: "Deleted." },
            countLine:  { ja: "%1：%2個", en: "%1: %2" },
            totalLine:  { ja: "合計：%1個", en: "Total: %1" },
            keptLine: {
                ja: "チェックを外した項目から参照されているなどの理由で残した項目：%1個",
                en: "Kept because still referenced by unchecked items, etc.: %1"
            }
        },
        undoName: { ja: "未使用項目の削除", en: "Delete Unused Items" }
    };

    // =========================================
    // 削除対象の定義 / Target definitions
    // =========================================

    /* 削除対象の種類（ダイアログ・確認リスト・結果の並び順） / Target kinds, in display order */
    var TARGET_KEYS = ["paragraphStyle", "characterStyle", "objectStyle", "tableStyle", "cellStyle", "masterSpread", "emptyPage", "swatch", "compositeFont"];

    /* 削除する項目のパネル分け / Panels for the target checkboxes */
    var TARGET_PANELS = [
        { labelKey: "styleTargets", targetKeys: ["paragraphStyle", "characterStyle", "objectStyle", "tableStyle", "cellStyle"] },
        { labelKey: "pageTargets",  targetKeys: ["masterSpread", "emptyPage"] },
        { labelKey: "otherTargets", targetKeys: ["swatch", "compositeFont"] }
    ];

    /* 初回の初期値 / Defaults on first run */
    var DEFAULT_CHECKED_TARGETS = { swatch: true };
    var DEFAULT_REMOVE_EMPTY_GROUPS = false;

    /* スタイルの種類ごとのコレクション名とテキスト検索のプロパティ名
       / Collection names and find property per style kind */
    var STYLE_KINDS = {
        paragraphStyle: { allStyles: "allParagraphStyles", styles: "paragraphStyles", groups: "paragraphStyleGroups", findProperty: "appliedParagraphStyle" },
        characterStyle: { allStyles: "allCharacterStyles", styles: "characterStyles", groups: "characterStyleGroups", findProperty: "appliedCharacterStyle" },
        objectStyle:    { allStyles: "allObjectStyles",    styles: "objectStyles",    groups: "objectStyleGroups",    findProperty: null },
        tableStyle:     { allStyles: "allTableStyles",     styles: "tableStyles",     groups: "tableStyleGroups",     findProperty: null },
        cellStyle:      { allStyles: "allCellStyles",      styles: "cellStyles",      groups: "cellStyleGroups",      findProperty: null }
    };
    var STYLE_KIND_KEYS = ["paragraphStyle", "characterStyle", "objectStyle", "tableStyle", "cellStyle"];

    /* 確認リストでのスタイルグループの種類名 / Kind key for style groups in the list */
    var GROUP_KIND_KEY = "styleGroup";

    /* 合成フォントは、フォント名のファミリー部分との照合に使うので名前から作ったキーで記録する
       / Composite fonts are keyed by name, to match the family part of applied font names */
    var COMPOSITE_FONT_KEY_PREFIX = "compositeFont:";

    // =========================================
    // 設定の記憶 / Saved settings
    // =========================================

    /* InDesignには任意の値を残す環境設定APIが無いので、設定ファイルに key=value で書き出す
       / InDesign has no scriptable preference store, so settings go to a key=value file */
    var PREFS_FILE_NAME              = "IdDeleteUnused-prefs.txt";
    var PREF_KEY_TARGETS             = "targets";
    var PREF_KEY_REMOVE_EMPTY_GROUPS = "removeEmptyGroups";

    /**
     * 設定ファイルを返す
     * @returns {File} 設定ファイル
     */
    function getPrefsFile() {
        return File(Folder.userData.fsName + "/" + PREFS_FILE_NAME);
    }

    /**
     * 設定ファイルを読み出す
     * @returns {object} キーと文字列値の対応。読めなければ空のオブジェクト
     */
    function loadPrefs() {
        var raw = readTextFile(getPrefsFile());
        if (!raw) return {};

        var prefs = {};
        var lines = String(raw).split("\n");
        for (var i = 0; i < lines.length; i++) {
            /* 値にも = が入りうるので最初の = だけで区切る / Split on the first = only */
            var separatorIndex = lines[i].indexOf("=");
            if (separatorIndex > 0) prefs[lines[i].substring(0, separatorIndex)] = trimWhitespace(lines[i].substring(separatorIndex + 1));
        }
        return prefs;
    }

    /**
     * 設定ファイルを書き出す
     * @param {object} prefs - キーと値の対応
     * @returns {void}
     */
    function savePrefs(prefs) {
        var lines = [];
        for (var key in prefs) {
            if (prefs.hasOwnProperty(key)) lines.push(key + "=" + prefs[key]);
        }

        /* 保存できなくても操作は続けられるので、書き出せたかどうかは見ない / A failed save must not break the run */
        writeTextFile(getPrefsFile(), lines.join("\n"));
    }

    /**
     * テキストファイルを読み込む
     * @param {File} textFile - 読み込むファイル
     * @returns {string} 中身。読めなければ空文字
     */
    function readTextFile(textFile) {
        if (!textFile.exists) return "";
        try {
            textFile.encoding = "UTF-8";
            if (!textFile.open("r")) return "";
            return textFile.read();
        } catch (e) {
            return "";
        } finally {
            try { textFile.close(); } catch (e) {}
        }
    }

    /**
     * テキストファイルへ書き出す
     * @param {File} textFile - 書き出すファイル
     * @param {string} text - 書き出す内容
     * @returns {boolean} 書き出せたら true
     */
    function writeTextFile(textFile, text) {
        try {
            textFile.encoding = "UTF-8";
            textFile.lineFeed = "Unix";
            if (!textFile.open("w")) return false;

            textFile.write(text);
            return true;
        } catch (e) {
            return false;
        } finally {
            try { textFile.close(); } catch (e) {}
        }
    }

    /**
     * 前後の空白を取り除く
     * @param {string} value - 対象の文字列
     * @returns {string} 前後の空白を除いた文字列
     */
    function trimWhitespace(value) {
        return String(value).replace(/^\s+|\s+$/g, "");
    }

    /**
     * 前回のダイアログの状態を読み出す。記録が無ければ初期値を返す
     * @returns {{targets: Object<string, boolean>, removeEmptyGroups: boolean}} ダイアログの状態
     */
    function loadDialogState() {
        var dialogState = { targets: DEFAULT_CHECKED_TARGETS, removeEmptyGroups: DEFAULT_REMOVE_EMPTY_GROUPS };
        var prefs = loadPrefs();

        if (prefs.hasOwnProperty(PREF_KEY_REMOVE_EMPTY_GROUPS)) dialogState.removeEmptyGroups = prefs[PREF_KEY_REMOVE_EMPTY_GROUPS] === "true";
        if (prefs.hasOwnProperty(PREF_KEY_TARGETS)) {
            dialogState.targets = {};
            var savedKeys = prefs[PREF_KEY_TARGETS].split(",");
            for (var j = 0; j < savedKeys.length; j++) dialogState.targets[savedKeys[j]] = true;
        }
        return dialogState;
    }

    /**
     * ダイアログの状態を設定ファイルに書き出す
     * @param {{targets: Object<string, boolean>, removeEmptyGroups: boolean}} dialogState - ダイアログの状態
     * @returns {void}
     */
    function saveDialogState(dialogState) {
        var checkedKeys = [];
        for (var i = 0; i < TARGET_KEYS.length; i++) {
            if (dialogState.targets[TARGET_KEYS[i]]) checkedKeys.push(TARGET_KEYS[i]);
        }
        var prefs = {};
        prefs[PREF_KEY_TARGETS] = checkedKeys.join(",");
        prefs[PREF_KEY_REMOVE_EMPTY_GROUPS] = dialogState.removeEmptyGroups;
        savePrefs(prefs);
    }

    // =========================================
    // 設定ダイアログ / Settings dialog
    // =========================================

    /**
     * 削除する項目のチェックボックスを、スタイル・ページ・その他のパネルに分けて並べる
     * @param {Window} parent - 追加先のウィンドウ
     * @param {Object<string, boolean>} checkedTargets - ONにしておく種類
     * @returns {Object<string, Checkbox>} 種類ごとのチェックボックス
     */
    function addTargetCheckboxes(parent, checkedTargets) {
        var targetCheckboxes = {};
        for (var i = 0; i < TARGET_PANELS.length; i++) {
            var targetPanel = parent.add("panel", undefined, getLabel(LABELS.panel[TARGET_PANELS[i].labelKey]));
            setupPanel(targetPanel, 6);

            var targetKeys = TARGET_PANELS[i].targetKeys;
            for (var j = 0; j < targetKeys.length; j++) {
                var targetKey = targetKeys[j];
                var targetCheckbox = targetPanel.add("checkbox", undefined, getLabel(LABELS.checkbox[targetKey]));
                targetCheckbox.value = checkedTargets[targetKey] === true;
                targetCheckbox.helpTip = getLabel(LABELS.tooltip[targetKey]) + "\n\n" + getLabel(LABELS.tooltip.optionClick);
                targetCheckboxes[targetKey] = targetCheckbox;
            }
        }
        return targetCheckboxes;
    }

    /**
     * オプションのチェックボックスを並べる
     * @param {Window} parent - 追加先のウィンドウ
     * @param {boolean} removeEmptyGroups - 「空のスタイルグループを削除」の初期値
     * @returns {Checkbox} 「空のスタイルグループを削除」のチェックボックス
     */
    function addOptionCheckboxes(parent, removeEmptyGroups) {
        var optionPanel = parent.add("panel", undefined, getLabel(LABELS.panel.options));
        setupPanel(optionPanel, 6);

        var emptyGroupsCheckbox = optionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.removeEmptyGroups));
        emptyGroupsCheckbox.value = removeEmptyGroups;
        emptyGroupsCheckbox.helpTip = getLabel(LABELS.tooltip.removeEmptyGroups);
        return emptyGroupsCheckbox;
    }

    /**
     * option（Alt）+クリックで、押した項目だけをONにする
     * すでにそれだけがONの状態で押したときは、すべてをONに戻す
     * onClick の時点で押した項目の値は反転済みなので、反転前の状態から判定する
     * @param {Object<string, Checkbox>} targetCheckboxes - 種類ごとのチェックボックス
     * @param {string} clickedKey - 押された項目の種類
     * @returns {void}
     */
    function applyOptionClick(targetCheckboxes, clickedKey) {
        var wasChecked = !targetCheckboxes[clickedKey].value;
        var othersUnchecked = true;
        for (var i = 0; i < TARGET_KEYS.length; i++) {
            if (TARGET_KEYS[i] !== clickedKey && targetCheckboxes[TARGET_KEYS[i]].value) othersUnchecked = false;
        }
        var checkAll = wasChecked && othersUnchecked;
        for (var j = 0; j < TARGET_KEYS.length; j++) {
            targetCheckboxes[TARGET_KEYS[j]].value = checkAll || TARGET_KEYS[j] === clickedKey;
        }
    }

    /**
     * 設定ダイアログを表示して、削除する項目とオプションを選ばせる
     * 開くときは前回の状態を復元し、OK で閉じたら保存する
     * @returns {{targets: Object<string, boolean>, removeEmptyGroups: boolean}|null} 選んだ内容。キャンセル時は null
     */
    function showSettingsDialog() {
        var savedState = loadDialogState();

        var settingsDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(settingsDialog);

        var targetCheckboxes = addTargetCheckboxes(settingsDialog, savedState.targets);
        var emptyGroupsCheckbox = addOptionCheckboxes(settingsDialog, savedState.removeEmptyGroups);
        var buttonRow = addButtonRow(settingsDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /**
         * 削除する項目が1つもONでなければ OK を押せなくする / Disable OK when no kind is checked
         * @returns {void}
         */
        function updateOKButton() {
            var hasTarget = false;
            for (var i = 0; i < TARGET_KEYS.length; i++) {
                if (targetCheckboxes[TARGET_KEYS[i]].value) hasTarget = true;
            }
            btnOK.enabled = hasTarget;
        }
        /**
         * チェックボックスのクリック処理を作る（option 併用の判定と OK ボタンの更新）
         * @param {string} targetKey - 対象の種類
         * @returns {Function} onClick に渡す関数
         */
        function createCheckboxClickHandler(targetKey) {
            return function () {
                if (ScriptUI.environment.keyboardState.altKey) applyOptionClick(targetCheckboxes, targetKey);
                updateOKButton();
            };
        }
        for (var i = 0; i < TARGET_KEYS.length; i++) {
            targetCheckboxes[TARGET_KEYS[i]].onClick = createCheckboxClickHandler(TARGET_KEYS[i]);
        }
        updateOKButton();

        if (settingsDialog.show() !== 1) return null;

        var selectedTargets = {};
        for (var j = 0; j < TARGET_KEYS.length; j++) {
            selectedTargets[TARGET_KEYS[j]] = targetCheckboxes[TARGET_KEYS[j]].value;
        }
        var dialogState = {
            targets: selectedTargets,
            removeEmptyGroups: emptyGroupsCheckbox.value
        };
        saveDialogState(dialogState);
        return dialogState;
    }

    // =========================================
    // 確認ダイアログ / Confirmation dialog
    // =========================================

    /**
     * 確認リストの種類の列に出す文字列を返す
     * @param {{kind: string, styleKind: string}} candidate - 削除候補
     * @returns {string} 種類名。スタイルグループは元のスタイルの種類を添える
     */
    function getKindText(candidate) {
        if (candidate.kind !== GROUP_KIND_KEY) return getLabel(LABELS.kindName[candidate.kind]);
        return getLabel(LABELS.kindName[GROUP_KIND_KEY]) + " (" + getLabel(LABELS.kindName[candidate.styleKind]) + ")";
    }

    /**
     * 削除候補を一覧で見せ、チェックを付けたものだけを返す
     * @param {Array<Object>} candidates - 削除候補
     * @returns {Array<Object>|null} チェックを付けた候補。キャンセル時は null
     */
    function showConfirmDialog(candidates) {
        var confirmDialog = new Window("dialog", getLabel(LABELS.dialog.confirm));
        setupWindow(confirmDialog);

        confirmDialog.add("statictext", undefined, getLabel(LABELS.message.confirmCount, [candidates.length]));

        var candidateList = confirmDialog.add("listbox", undefined, [], {
            numberOfColumns: 2,
            showHeaders: true,
            columnTitles: [getLabel(LABELS.columnTitle.kind), getLabel(LABELS.columnTitle.name)],
            columnWidths: CONFIRM_COLUMN_WIDTHS
        });
        candidateList.preferredSize = CONFIRM_LIST_SIZE;
        candidateList.helpTip = getLabel(LABELS.tooltip.confirmList);

        for (var i = 0; i < candidates.length; i++) {
            var listItem = candidateList.add("item", getKindText(candidates[i]));
            listItem.subItems[0].text = candidates[i].name;
            listItem.checked = true;
        }

        /* Mac の listbox はチェック欄を直接クリックしても切り替わらないので、ダブルクリックで切り替える
           / Checkboxes in a Mac listbox do not respond to clicks, so double-click toggles them */
        candidateList.onDoubleClick = function () {
            if (candidateList.selection) candidateList.selection.checked = !candidateList.selection.checked;
        };

        /* 左にチェックの一括切り替え、右に［キャンセル］［削除］ / Check-all toggles on the left, Cancel and Delete on the right */
        var buttonRow = addButtonRow(confirmDialog);
        var btnCheckAll = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.checkAll));
        var btnUncheckAll = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.uncheckAll));
        buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.remove), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        /**
         * リストのチェックをまとめて切り替える
         * @param {boolean} checked - ONにするか
         * @returns {Function} onClick に渡す関数
         */
        function createCheckAllHandler(checked) {
            return function () {
                for (var j = 0; j < candidateList.items.length; j++) candidateList.items[j].checked = checked;
            };
        }
        btnCheckAll.onClick = createCheckAllHandler(true);
        btnUncheckAll.onClick = createCheckAllHandler(false);

        if (confirmDialog.show() !== 1) return null;

        var confirmedCandidates = [];
        for (var k = 0; k < candidateList.items.length; k++) {
            if (candidateList.items[k].checked) confirmedCandidates.push(candidates[k]);
        }
        return confirmedCandidates;
    }

    // =========================================
    // 参照の収集 / Reference collection
    // =========================================

    /**
     * 種類ごとの空の記録を作る
     * @returns {Object<string, Object<string, boolean>>} 種類をキーにした、IDの記録
     */
    function createKindMap() {
        var kindMap = {};
        for (var i = 0; i < TARGET_KEYS.length; i++) kindMap[TARGET_KEYS[i]] = {};
        return kindMap;
    }

    /**
     * 値が指定クラスのオブジェクトなら、使用中として記録する
     * 未設定のときは文字列や NothingEnum が返るので、クラスで見分ける
     * @param {Object<string, boolean>} referenceMap - IDをキーにした使用中の記録
     * @param {*} value - プロパティ値
     * @param {Function} domClass - ParagraphStyle などのクラス
     * @returns {void}
     */
    function markReferenced(referenceMap, value, domClass) {
        if (value instanceof domClass) referenceMap[value.id] = true;
    }

    /**
     * ルートスタイル（［段落スタイルなし］や［なし］）を除いたスタイルの一覧を返す
     * ルートスタイルは basedOn などを読むだけで「ルートスタイルに対する無効な要求」の例外になる
     * @param {Document} doc - 対象ドキュメント
     * @param {string} styleKind - STYLE_KINDS のキー
     * @returns {Array<Object>} ルートスタイルを除いたスタイル
     */
    function getNonRootStyles(doc, styleKind) {
        var allStyles = doc[STYLE_KINDS[styleKind].allStyles];
        var rootStyleId = doc[STYLE_KINDS[styleKind].styles][0].id;
        var styles = [];
        for (var i = 0; i < allStyles.length; i++) {
            if (allStyles[i].id !== rootStyleId) styles.push(allStyles[i]);
        }
        return styles;
    }

    /**
     * 表とセルに適用されている表スタイル・セルスタイルを記録する
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, Object<string, boolean>>} directUse - 種類ごとの使用中の記録
     * @returns {void}
     */
    function markTableUse(doc, directUse) {
        for (var i = 0; i < doc.stories.length; i++) {
            var tables = doc.stories[i].tables;
            for (var j = 0; j < tables.length; j++) {
                markReferenced(directUse.tableStyle, tables[j].appliedTableStyle, TableStyle);
                if (tables[j].cells.length === 0) continue;
                var cellStyleList = tables[j].cells.everyItem().appliedCellStyle;
                if (!(cellStyleList instanceof Array)) cellStyleList = [cellStyleList];
                for (var k = 0; k < cellStyleList.length; k++) {
                    markReferenced(directUse.cellStyle, cellStyleList[k], CellStyle);
                }
            }
        }
    }

    /**
     * 目次スタイルから参照されている段落・文字スタイルを記録する
     * 目次の項目は、対象の段落スタイルを名前で持つ
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, Object<string, boolean>>} directUse - 種類ごとの使用中の記録
     * @returns {void}
     */
    function markTocUse(doc, directUse) {
        var tocEntryNames = {};
        for (var i = 0; i < doc.tocStyles.length; i++) {
            var tocStyle = doc.tocStyles[i];
            markReferenced(directUse.paragraphStyle, tocStyle.titleStyle, ParagraphStyle);
            for (var j = 0; j < tocStyle.tocStyleEntries.length; j++) {
                var tocEntry = tocStyle.tocStyleEntries[j];
                tocEntryNames[tocEntry.name] = true;
                markReferenced(directUse.paragraphStyle, tocEntry.formatStyle, ParagraphStyle);
                markReferenced(directUse.characterStyle, tocEntry.pageNumberStyle, CharacterStyle);
                markReferenced(directUse.characterStyle, tocEntry.separatorStyle, CharacterStyle);
            }
        }
        var paragraphStyles = doc.allParagraphStyles;
        for (var k = 0; k < paragraphStyles.length; k++) {
            if (tocEntryNames[paragraphStyles[k].name]) directUse.paragraphStyle[paragraphStyles[k].id] = true;
        }
    }

    /**
     * appliedFont の値から、合成フォントの記録に使うキーを作る
     * フォント名は「ファミリー名\tスタイル名」の形なので、ファミリー名の部分で合成フォント名と照合する
     * 未設定のスタイルでは空文字が返る
     * @param {*} fontValue - appliedFont の値（Font か文字列）
     * @returns {string} 合成フォントの記録キー。フォント名が取れなければ空文字
     */
    function getFontFamilyKey(fontValue) {
        var fontName = (typeof fontValue === "string") ? fontValue : fontValue.name;
        if (!fontName) return "";
        return COMPOSITE_FONT_KEY_PREFIX + String(fontName).split("\t")[0];
    }

    /**
     * 書式範囲のフォントを使用中として記録する
     * @param {TextStyleRanges} textStyleRanges - 調べる書式範囲
     * @param {Object<string, boolean>} fontUse - 合成フォントの使用中の記録
     * @returns {void}
     */
    function markRangeFonts(textStyleRanges, fontUse) {
        if (textStyleRanges.length === 0) return;
        var fontValues = textStyleRanges.everyItem().appliedFont;
        if (!(fontValues instanceof Array)) fontValues = [fontValues];
        for (var i = 0; i < fontValues.length; i++) {
            var fontKey = getFontFamilyKey(fontValues[i]);
            if (fontKey) fontUse[fontKey] = true;
        }
    }

    /**
     * テキスト（ストーリー・表のセル・脚注）に使われているフォントを記録する
     * 合成フォントのオブジェクトは検索条件に入れられない（Font か文字列しか受け付けない）ので、書式範囲を直接読む
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, boolean>} fontUse - 合成フォントの使用中の記録
     * @returns {void}
     */
    function markTextFonts(doc, fontUse) {
        for (var i = 0; i < doc.stories.length; i++) {
            var story = doc.stories[i];
            markRangeFonts(story.textStyleRanges, fontUse);
            for (var j = 0; j < story.tables.length; j++) {
                var cells = story.tables[j].cells;
                for (var k = 0; k < cells.length; k++) markRangeFonts(cells[k].texts[0].textStyleRanges, fontUse);
            }
            for (var m = 0; m < story.footnotes.length; m++) {
                markRangeFonts(story.footnotes[m].texts[0].textStyleRanges, fontUse);
            }
        }
    }

    /**
     * 削除候補に左右されない使用状況を集める（ページアイテム・表への適用、ドキュメント設定）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, boolean>} selectedTargets - 種類ごとのON/OFF（重い調べものを省くのに使う）
     * @returns {Object<string, Object<string, boolean>>} 種類ごとの使用中のID
     */
    function collectDirectUse(doc, selectedTargets) {
        var directUse = createKindMap();

        var pageItems = doc.allPageItems;
        for (var i = 0; i < pageItems.length; i++) {
            markReferenced(directUse.objectStyle, pageItems[i].appliedObjectStyle, ObjectStyle);
        }
        var itemDefaults = doc.pageItemDefaults;
        markReferenced(directUse.objectStyle, itemDefaults.appliedTextObjectStyle, ObjectStyle);
        markReferenced(directUse.objectStyle, itemDefaults.appliedGraphicObjectStyle, ObjectStyle);
        markReferenced(directUse.objectStyle, itemDefaults.appliedGridObjectStyle, ObjectStyle);

        markTableUse(doc, directUse);
        markTocUse(doc, directUse);

        markReferenced(directUse.paragraphStyle, doc.footnoteOptions.footnoteTextStyle, ParagraphStyle);
        markReferenced(directUse.characterStyle, doc.footnoteOptions.footnoteMarkerStyle, CharacterStyle);
        markReferenced(directUse.paragraphStyle, doc.textDefaults.appliedParagraphStyle, ParagraphStyle);
        markReferenced(directUse.characterStyle, doc.textDefaults.appliedCharacterStyle, CharacterStyle);

        if (selectedTargets.compositeFont) {
            var defaultFontKey = getFontFamilyKey(doc.textDefaults.appliedFont);
            if (defaultFontKey) directUse.compositeFont[defaultFontKey] = true;
            markTextFonts(doc, directUse.compositeFont);
        }
        return directUse;
    }

    /**
     * 段落スタイルから、他の段落スタイルと文字スタイルへの参照を記録する
     * @param {ParagraphStyle} paragraphStyle - 参照元の段落スタイル
     * @param {Object<string, Object<string, boolean>>} crossRefs - 種類ごとの参照の記録
     * @returns {void}
     */
    function markParagraphStyleRefs(paragraphStyle, crossRefs) {
        markReferenced(crossRefs.paragraphStyle, paragraphStyle.basedOn, ParagraphStyle);
        /* 自分自身を「次のスタイル」にしているのは参照に数えない / Self as Next Style does not count */
        if (paragraphStyle.nextStyle instanceof ParagraphStyle && paragraphStyle.nextStyle.id !== paragraphStyle.id) {
            crossRefs.paragraphStyle[paragraphStyle.nextStyle.id] = true;
        }
        markReferenced(crossRefs.characterStyle, paragraphStyle.bulletsCharacterStyle, CharacterStyle);
        markReferenced(crossRefs.characterStyle, paragraphStyle.numberingCharacterStyle, CharacterStyle);
        markReferenced(crossRefs.characterStyle, paragraphStyle.dropCapStyle, CharacterStyle);

        var nestedCollections = [paragraphStyle.nestedStyles, paragraphStyle.nestedGrepStyles, paragraphStyle.nestedLineStyles];
        for (var i = 0; i < nestedCollections.length; i++) {
            for (var j = 0; j < nestedCollections[i].length; j++) {
                markReferenced(crossRefs.characterStyle, nestedCollections[i][j].appliedCharacterStyle, CharacterStyle);
            }
        }
    }

    /**
     * 表スタイルから、他の表スタイルと各領域のセルスタイルへの参照を記録する
     * @param {TableStyle} tableStyle - 参照元の表スタイル
     * @param {Object<string, Object<string, boolean>>} crossRefs - 種類ごとの参照の記録
     * @returns {void}
     */
    function markTableStyleRefs(tableStyle, crossRefs) {
        markReferenced(crossRefs.tableStyle, tableStyle.basedOn, TableStyle);
        var regionProperties = ["headerRegionCellStyle", "footerRegionCellStyle", "bodyRegionCellStyle", "leftColumnRegionCellStyle", "rightColumnRegionCellStyle"];
        for (var i = 0; i < regionProperties.length; i++) {
            markReferenced(crossRefs.cellStyle, tableStyle[regionProperties[i]], CellStyle);
        }
    }

    /**
     * 段落・文字スタイルのフォントを、合成フォントの参照として記録する
     * @param {ParagraphStyle|CharacterStyle} style - 参照元のスタイル
     * @param {Object<string, Object<string, boolean>>} crossRefs - 種類ごとの参照の記録
     * @returns {void}
     */
    function markStyleFont(style, crossRefs) {
        var fontKey = getFontFamilyKey(style.appliedFont);
        if (fontKey) crossRefs.compositeFont[fontKey] = true;
    }

    /**
     * スタイル同士の参照を集める。ignoredIds にある参照元（削除予定のもの）は数えない
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, boolean>} ignoredIds - 削除予定のID
     * @param {Object<string, Object<string, boolean>>} crossRefs - 種類ごとの参照の記録
     * @returns {void}
     */
    function markStyleCrossRefs(doc, ignoredIds, crossRefs) {
        var paragraphStyles = getNonRootStyles(doc, "paragraphStyle");
        for (var i = 0; i < paragraphStyles.length; i++) {
            if (ignoredIds[paragraphStyles[i].id]) continue;
            markParagraphStyleRefs(paragraphStyles[i], crossRefs);
            markStyleFont(paragraphStyles[i], crossRefs);
        }
        var characterStyles = getNonRootStyles(doc, "characterStyle");
        for (var j = 0; j < characterStyles.length; j++) {
            if (ignoredIds[characterStyles[j].id]) continue;
            markReferenced(crossRefs.characterStyle, characterStyles[j].basedOn, CharacterStyle);
            markStyleFont(characterStyles[j], crossRefs);
        }
        var objectStyles = getNonRootStyles(doc, "objectStyle");
        for (var k = 0; k < objectStyles.length; k++) {
            if (ignoredIds[objectStyles[k].id]) continue;
            markReferenced(crossRefs.objectStyle, objectStyles[k].basedOn, ObjectStyle);
            markReferenced(crossRefs.paragraphStyle, objectStyles[k].appliedParagraphStyle, ParagraphStyle);
        }
        var tableStyles = getNonRootStyles(doc, "tableStyle");
        for (var m = 0; m < tableStyles.length; m++) {
            if (!ignoredIds[tableStyles[m].id]) markTableStyleRefs(tableStyles[m], crossRefs);
        }
        var cellStyles = getNonRootStyles(doc, "cellStyle");
        for (var n = 0; n < cellStyles.length; n++) {
            if (ignoredIds[cellStyles[n].id]) continue;
            markReferenced(crossRefs.cellStyle, cellStyles[n].basedOn, CellStyle);
            markReferenced(crossRefs.paragraphStyle, cellStyles[n].appliedParagraphStyle, ParagraphStyle);
        }
    }

    /**
     * スタイル・合成フォント・スウォッチ・ページ・親ページの間の参照を集める。ignoredIds にある参照元（削除予定のもの）は数えない
     * スウォッチはグラデーションの分岐点と濃淡の元のカラー。unusedSwatches にはこれらが含まれ、
     * 消すとグラデーションや濃淡が変わってしまう
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, boolean>} ignoredIds - 削除予定のID
     * @returns {Object<string, Object<string, boolean>>} 種類ごとの参照されているID
     */
    function collectCrossRefs(doc, ignoredIds) {
        var crossRefs = createKindMap();
        markStyleCrossRefs(doc, ignoredIds, crossRefs);

        for (var i = 0; i < doc.gradients.length; i++) {
            if (ignoredIds[doc.gradients[i].id]) continue;
            var gradientStops = doc.gradients[i].gradientStops;
            for (var j = 0; j < gradientStops.length; j++) {
                markReferenced(crossRefs.swatch, gradientStops[j].stopColor, Swatch);
            }
        }
        for (var k = 0; k < doc.tints.length; k++) {
            if (!ignoredIds[doc.tints[k].id]) markReferenced(crossRefs.swatch, doc.tints[k].baseColor, Swatch);
        }

        /* 空ページを消すと、そこに適用していた親ページが未使用になりうる / Deleting empty pages may free their parent pages */
        for (var p = 0; p < doc.pages.length; p++) {
            if (!ignoredIds[doc.pages[p].id]) markReferenced(crossRefs.masterSpread, doc.pages[p].appliedMaster, MasterSpread);
        }
        for (var m = 0; m < doc.masterSpreads.length; m++) {
            if (ignoredIds[doc.masterSpreads[m].id]) continue;
            var masterPages = doc.masterSpreads[m].pages;
            for (var n = 0; n < masterPages.length; n++) {
                markReferenced(crossRefs.masterSpread, masterPages[n].appliedMaster, MasterSpread);
            }
        }
        return crossRefs;
    }

    // =========================================
    // 削除候補の洗い出し / Finding candidates
    // =========================================

    /**
     * テキスト検索の範囲を、非表示・ロック・マスター・脚注まで広げる
     * @returns {Object<string, boolean>} 変更前の設定
     */
    function widenFindScope() {
        var findOptions = app.findChangeTextOptions;
        var scopeKeys = ["includeHiddenLayers", "includeLockedLayersForFind", "includeLockedStoriesForFind", "includeMasterPages", "includeFootnotes"];
        var savedScope = {};
        for (var i = 0; i < scopeKeys.length; i++) {
            savedScope[scopeKeys[i]] = findOptions[scopeKeys[i]];
            findOptions[scopeKeys[i]] = true;
        }
        return savedScope;
    }

    /**
     * スタイルがテキストに適用されているかを検索で調べる
     * @param {Document} doc - 対象ドキュメント
     * @param {string} findProperty - "appliedParagraphStyle" または "appliedCharacterStyle"
     * @param {ParagraphStyle|CharacterStyle} style - 調べるスタイル
     * @returns {boolean} 適用されていれば true
     */
    function isAppliedToText(doc, findProperty, style) {
        app.findTextPreferences = NothingEnum.NOTHING;
        app.findTextPreferences[findProperty] = style;
        var isApplied = doc.findText().length > 0;
        app.findTextPreferences = NothingEnum.NOTHING;
        return isApplied;
    }

    /**
     * 選ばれた種類の、削除候補になりうる項目を並べる
     * ［ ］で囲まれた既定のスタイルと、名前のないカラーは最初から外す
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, boolean>} selectedTargets - 種類ごとのON/OFF
     * @returns {Array<{kind: string, item: Object, id: (number|string), name: string}>} 候補になりうる項目
     */
    function listRemovableItems(doc, selectedTargets) {
        var removableItems = [];
        /**
         * 項目を候補の形にして追加する
         * @param {string} kind - 種類
         * @param {Object} item - DOMオブジェクト
         * @returns {void}
         */
        function pushItem(kind, item) {
            removableItems.push({ kind: kind, item: item, id: item.id, name: item.name });
        }

        for (var i = 0; i < STYLE_KIND_KEYS.length; i++) {
            var styleKind = STYLE_KIND_KEYS[i];
            if (!selectedTargets[styleKind]) continue;
            var styles = getNonRootStyles(doc, styleKind);
            for (var j = 0; j < styles.length; j++) {
                if (styles[j].name.charAt(0) !== "[") pushItem(styleKind, styles[j]);
            }
        }
        if (selectedTargets.swatch) {
            var unusedSwatchList = doc.unusedSwatches;
            for (var k = 0; k < unusedSwatchList.length; k++) {
                if (unusedSwatchList[k].name !== "") pushItem("swatch", unusedSwatchList[k]);
            }
        }
        if (selectedTargets.masterSpread) {
            for (var m = 0; m < doc.masterSpreads.length; m++) pushItem("masterSpread", doc.masterSpreads[m]);
        }
        if (selectedTargets.emptyPage) {
            for (var n = 0; n < doc.pages.length; n++) {
                if (doc.pages[n].pageItems.length === 0) pushItem("emptyPage", doc.pages[n]);
            }
        }
        if (selectedTargets.compositeFont) {
            /* ［No composite font］など［ ］で始まる既定のものは外す / Skip defaults in [ ] */
            for (var q = 0; q < doc.compositeFonts.length; q++) {
                var compositeFont = doc.compositeFonts[q];
                if (compositeFont.name.charAt(0) === "[") continue;
                removableItems.push({ kind: "compositeFont", item: compositeFont, id: COMPOSITE_FONT_KEY_PREFIX + compositeFont.name, name: compositeFont.name });
            }
        }
        return removableItems;
    }

    /**
     * 項目が使用中かを判定する。テキストへの適用は重いので最後に調べ、結果を控える
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} removableItem - listRemovableItems() の要素
     * @param {Object<string, Object<string, boolean>>} directUse - 削除候補に左右されない使用状況
     * @param {Object<string, Object<string, boolean>>} crossRefs - 項目同士の参照
     * @param {Object<string, boolean>} textUseCache - テキストへの適用の判定結果
     * @returns {boolean} 使用中なら true
     */
    function isInUse(doc, removableItem, directUse, crossRefs, textUseCache) {
        var kind = removableItem.kind;
        if (directUse[kind][removableItem.id] || crossRefs[kind][removableItem.id]) return true;

        var findProperty = STYLE_KINDS[kind] ? STYLE_KINDS[kind].findProperty : null;
        if (!findProperty) return false;
        if (!textUseCache.hasOwnProperty(removableItem.id)) {
            textUseCache[removableItem.id] = isAppliedToText(doc, findProperty, removableItem.item);
        }
        return textUseCache[removableItem.id];
    }

    /**
     * 空のスタイルグループを、子グループが先に来る順で集める
     * deletedIds にあるスタイルは削除済みとみなす
     * @param {Object} groupCollection - paragraphStyleGroups などのコレクション
     * @param {string} styleKind - STYLE_KINDS のキー
     * @param {Object<string, boolean>} deletedIds - 削除予定のID
     * @param {Array<Object>} emptyGroups - 見つかったグループの追加先
     * @returns {boolean} コレクション内のグループがすべて空なら true
     */
    function collectEmptyGroups(groupCollection, styleKind, deletedIds, emptyGroups) {
        var allEmpty = true;
        for (var i = 0; i < groupCollection.length; i++) {
            var styleGroup = groupCollection[i];
            var subgroupsEmpty = collectEmptyGroups(styleGroup[STYLE_KINDS[styleKind].groups], styleKind, deletedIds, emptyGroups);
            var styles = styleGroup[STYLE_KINDS[styleKind].styles];
            var stylesDeleted = true;
            for (var j = 0; j < styles.length; j++) {
                if (!deletedIds[styles[j].id]) stylesDeleted = false;
            }
            if (subgroupsEmpty && stylesDeleted) {
                emptyGroups.push(styleGroup);
            } else {
                allEmpty = false;
            }
        }
        return allEmpty;
    }

    /**
     * 空になるスタイルグループを削除候補の形で返す
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, boolean>} deletedIds - 削除予定のID
     * @returns {Array<Object>} スタイルグループの削除候補
     */
    function findEmptyGroupCandidates(doc, deletedIds) {
        var groupCandidates = [];
        for (var i = 0; i < STYLE_KIND_KEYS.length; i++) {
            var styleKind = STYLE_KIND_KEYS[i];
            var emptyGroups = [];
            collectEmptyGroups(doc[STYLE_KINDS[styleKind].groups], styleKind, deletedIds, emptyGroups);
            for (var j = 0; j < emptyGroups.length; j++) {
                groupCandidates.push({ kind: GROUP_KIND_KEY, styleKind: styleKind, item: emptyGroups[j], id: emptyGroups[j].id, name: emptyGroups[j].name });
            }
        }
        return groupCandidates;
    }

    /**
     * 削除候補を洗い出す
     * 子スタイル・派生した親ページ・空ページなどを消すと元が未使用になることがあるので、候補が増えなくなるまで繰り返す
     * @param {Document} doc - 対象ドキュメント
     * @param {{targets: Object<string, boolean>, removeEmptyGroups: boolean}} dialogState - 設定ダイアログの内容
     * @returns {Array<Object>} 削除候補（種類の並び順）
     */
    function findCandidates(doc, dialogState) {
        var removableItems = listRemovableItems(doc, dialogState.targets);
        var directUse = collectDirectUse(doc, dialogState.targets);
        var textUseCache = {};
        var candidateIds = {};

        var addedCount;
        do {
            addedCount = 0;
            var crossRefs = collectCrossRefs(doc, candidateIds);
            for (var i = 0; i < removableItems.length; i++) {
                if (candidateIds[removableItems[i].id]) continue;
                if (isInUse(doc, removableItems[i], directUse, crossRefs, textUseCache)) continue;
                candidateIds[removableItems[i].id] = true;
                addedCount++;
            }
        } while (addedCount > 0);

        var candidates = [];
        for (var j = 0; j < removableItems.length; j++) {
            if (candidateIds[removableItems[j].id]) candidates.push(removableItems[j]);
        }
        if (dialogState.removeEmptyGroups) candidates = candidates.concat(findEmptyGroupCandidates(doc, candidateIds));
        return candidates;
    }

    /**
     * 検索範囲を広げて削除候補を洗い出す。テキスト検索の範囲は必ず元に戻す
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} dialogState - 設定ダイアログの内容
     * @returns {Array<Object>} 削除候補
     */
    function findCandidatesWithWideScope(doc, dialogState) {
        var savedScope = widenFindScope();
        try {
            return findCandidates(doc, dialogState);
        } finally {
            /* 検索範囲はユーザーの設定なので、失敗しても戻す / Always restore the user's find scope */
            app.findChangeTextOptions.properties = savedScope;
        }
    }

    // =========================================
    // 削除 / Deletion
    // =========================================

    /**
     * チェックを外した項目から参照されている候補を外す
     * 外した候補がさらに別の候補を参照していることがあるので、外すものが無くなるまで繰り返す
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<Object>} confirmedItems - チェックを付けた候補（スタイルグループを除く）
     * @returns {Array<Object>} 実際に削除する候補
     */
    function excludeStillReferenced(doc, confirmedItems) {
        var confirmedKinds = {};
        for (var n = 0; n < confirmedItems.length; n++) confirmedKinds[confirmedItems[n].kind] = true;
        var directUse = collectDirectUse(doc, confirmedKinds);
        var deletingIds = {};
        for (var i = 0; i < confirmedItems.length; i++) deletingIds[confirmedItems[i].id] = true;

        var droppedCount;
        do {
            droppedCount = 0;
            var crossRefs = collectCrossRefs(doc, deletingIds);
            for (var j = 0; j < confirmedItems.length; j++) {
                var confirmedItem = confirmedItems[j];
                if (!deletingIds[confirmedItem.id]) continue;
                if (directUse[confirmedItem.kind][confirmedItem.id] || crossRefs[confirmedItem.kind][confirmedItem.id]) {
                    delete deletingIds[confirmedItem.id];
                    droppedCount++;
                }
            }
        } while (droppedCount > 0);

        var deletingItems = [];
        for (var k = 0; k < confirmedItems.length; k++) {
            if (deletingIds[confirmedItems[k].id]) deletingItems.push(confirmedItems[k]);
        }
        return deletingItems;
    }

    /**
     * 項目を削除する
     * スウォッチと合成フォントには削除可否のプロパティが無いので、remove() の例外で見分ける
     * 空ページは、ドキュメントの最後の1ページになったら残す
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} candidate - 削除候補
     * @returns {boolean} 削除できたら true
     */
    function removeCandidate(doc, candidate) {
        if (candidate.kind === "emptyPage" && doc.pages.length <= 1) return false;
        if (candidate.kind !== "swatch" && candidate.kind !== "compositeFont") {
            candidate.item.remove();
            return true;
        }
        try {
            candidate.item.remove();
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * スタイルグループが空かを調べる
     * @param {Object} candidate - スタイルグループの削除候補
     * @returns {boolean} スタイルもサブグループも無ければ true
     */
    function isGroupEmpty(candidate) {
        var styleGroup = candidate.item;
        return styleGroup[STYLE_KINDS[candidate.styleKind].styles].length === 0 &&
            styleGroup[STYLE_KINDS[candidate.styleKind].groups].length === 0;
    }

    /**
     * チェックを付けた候補を削除する
     * スタイルグループは最後に、実際に空になったものだけを消す（子グループが先の順）
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<Object>} confirmedCandidates - チェックを付けた候補
     * @param {Object<string, number>} removedCounts - 種類ごとの削除数（加算していく）
     * @returns {void}
     */
    function removeConfirmedCandidates(doc, confirmedCandidates, removedCounts) {
        var confirmedItems = [];
        var groupCandidates = [];
        for (var i = 0; i < confirmedCandidates.length; i++) {
            if (confirmedCandidates[i].kind === GROUP_KIND_KEY) {
                groupCandidates.push(confirmedCandidates[i]);
            } else {
                confirmedItems.push(confirmedCandidates[i]);
            }
        }

        var deletingItems = excludeStillReferenced(doc, confirmedItems);
        for (var j = 0; j < deletingItems.length; j++) {
            if (removeCandidate(doc, deletingItems[j])) removedCounts[deletingItems[j].kind]++;
        }
        for (var k = 0; k < groupCandidates.length; k++) {
            if (!isGroupEmpty(groupCandidates[k])) continue;
            groupCandidates[k].item.remove();
            removedCounts[GROUP_KIND_KEY]++;
        }
    }

    /**
     * 種類ごとの削除数を 0 で用意する
     * @returns {Object<string, number>} 種類（スタイルグループを含む）ごとの削除数
     */
    function createCountMap() {
        var countMap = {};
        countMap[GROUP_KIND_KEY] = 0;
        for (var i = 0; i < TARGET_KEYS.length; i++) countMap[TARGET_KEYS[i]] = 0;
        return countMap;
    }

    /**
     * チェックを付けた候補を、1回の取り消し単位で削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<Object>} confirmedCandidates - チェックを付けた候補
     * @returns {Object<string, number>} 種類ごとの削除数
     */
    function removeConfirmedWithUndo(doc, confirmedCandidates) {
        var removedCounts = createCountMap();
        app.doScript(function () {
            removeConfirmedCandidates(doc, confirmedCandidates, removedCounts);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undoName));
        return removedCounts;
    }

    // =========================================
    // 結果 / Result
    // =========================================

    /**
     * 種類ごとの削除数・合計・残した数を表示する
     * @param {{targets: Object<string, boolean>, removeEmptyGroups: boolean}} dialogState - 設定ダイアログの内容
     * @param {Object<string, number>} removedCounts - 種類ごとの削除数
     * @param {number} confirmedCount - チェックを付けた候補の数
     * @returns {void}
     */
    function showResult(dialogState, removedCounts, confirmedCount) {
        var resultLines = [getLabel(LABELS.alert.result), ""];

        var shownKeys = [];
        for (var i = 0; i < TARGET_KEYS.length; i++) {
            if (dialogState.targets[TARGET_KEYS[i]]) shownKeys.push(TARGET_KEYS[i]);
        }
        if (dialogState.removeEmptyGroups) shownKeys.push(GROUP_KIND_KEY);

        var totalCount = 0;
        for (var j = 0; j < shownKeys.length; j++) {
            var removedCount = removedCounts[shownKeys[j]];
            totalCount += removedCount;
            resultLines.push(getLabel(LABELS.alert.countLine, [getLabel(LABELS.kindName[shownKeys[j]]), removedCount]));
        }
        resultLines.push("");
        resultLines.push(getLabel(LABELS.alert.totalLine, [totalCount]));
        if (confirmedCount > totalCount) resultLines.push(getLabel(LABELS.alert.keptLine, [confirmedCount - totalCount]));

        alert(resultLines.join("\n"), getLabel(LABELS.dialog.title));
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * 設定ダイアログ → 候補の洗い出し → 確認リスト → 削除 → 結果表示
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument), getLabel(LABELS.dialog.title));
            return;
        }

        var dialogState = showSettingsDialog();
        if (!dialogState) return;

        var doc = app.activeDocument;
        var candidates = findCandidatesWithWideScope(doc, dialogState);
        if (candidates.length === 0) {
            alert(getLabel(LABELS.alert.noneFound), getLabel(LABELS.dialog.title));
            return;
        }

        var confirmedCandidates = showConfirmDialog(candidates);
        if (!confirmedCandidates || confirmedCandidates.length === 0) return;

        var removedCounts = removeConfirmedWithUndo(doc, confirmedCandidates);
        showResult(dialogState, removedCounts, confirmedCandidates.length);
    }

    main();

})();
