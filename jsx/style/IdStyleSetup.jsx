#target indesign

/*

### 概要

段落スタイル・文字スタイルとそのグループ、継承関係、正規表現スタイル（説明付き）までを一括で登録します。フォントは base-font（全体）・base-heading（見出し系）・base-text（本文系）で一括変更できます。
スタイル名は HTML（h1 / p）と Word 対応（Heading 1 / Normal）から選べ、既定では既存の同名スタイルには手を触れません。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdStyleSetup.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nfe87ec253780

### Overview

Registers paragraph and character styles together with their groups, inheritance and GREP styles (with comments) in one pass. Fonts can be changed in one place through base-font (everything), base-heading (headings) or base-text (body text).
Style names can follow either HTML (h1 / p) or Word (Heading 1 / Normal), and existing same-named styles are left untouched by default.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdStyleSetup.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdStyleSetup";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdStyleSetup.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdStyleSetup.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nfe87ec253780"; /* 紹介記事 / article URL */

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

    /* 同名スタイルが既にある場合の挙動 / Behavior when a same-named style already exists.
       true:  既存スタイルを置き換える。本スクリプトが設定する全属性を再適用し、GREP は消してから付け直す
              / Replace the existing style: re-apply every attribute this script sets and rebuild its GREP rules
       false: 既存スタイルには触れず、新規作成したスタイルにだけ属性を適用する
              / Leave existing styles untouched and apply attributes only to newly created ones
       本スクリプトが扱わない属性（フォント・サイズなど）はリセットしません。スタイル実体も削除しないため、
       適用済みテキストとの関連は保たれます。
       / Attributes this script does not set are left alone, and no style object is deleted,
         so text keeps its style association. */
    var OVERWRITE_EXISTING_STYLES = false;

    /* ［Word 対応］を選んだときのスタイル名 / Style names used when the "Word" scheme is chosen.
       キーは HTML 側の名前。ここに無いスタイル（p.caption、link、グループ内のスタイルなど）は
       どちらの体系でも同じ名前のままです /
       Keys are the HTML names; styles not listed here (p.caption, link, grouped styles, …)
       keep the same name in both schemes.
       Word の「見出し1〜5」「標準」「リスト段落」「強調太字」「強調斜体」に対応する名前なので、
       Word 原稿を配置するときにスタイルのマッピングが不要になります /
       These match Word's Heading 1–5, Normal, List Paragraph, Strong and Emphasis, so placing a
       Word manuscript needs no style mapping. */
    var WORD_STYLE_NAMES = {
        "h1":          "Heading 1",
        "h2":          "Heading 2",
        "h3":          "Heading 3",
        "h4":          "Heading 4",
        "h5":          "Heading 5",
        "h6":          "Heading 6",
        "p":           "Normal",
        "ul-li":       "List Paragraph",
        "ol-li":       "List Number",
        "strong-bold": "Strong",
        "em-italic":   "Emphasis"
    };

    // =========================================
    // レイアウト設定 / Layout settings
    // =========================================

    /* 進捗バーの幅と高さ（px）/ Width and height of the progress bar (px) */
    var PROGRESS_BAR_WIDTH  = 320;
    var PROGRESS_BAR_HEIGHT = 12;

    /* ラジオボタンの下に添えるスタイル名の見本の字下げ（px）/ Indent of the sample style names under each radio button (px) */
    var SCHEME_SAMPLE_INDENT = 18;

    /* ダイアログの不透明度 / Dialog opacity */
    var DIALOG_OPACITY = 0.98;

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
            title: { ja: "スタイル一括登録", en: "Register Styles" }
        },
        panel: {
            styleNames: { ja: "スタイル名", en: "Style names" }
        },
        scheme: {
            html: { ja: "HTML", en: "HTML" },
            word: { ja: "Word 対応", en: "Word" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "OK", en: "OK" }
        },
        tooltip: {
            schemeHtml: {
                ja: "h1〜h6、p、ul-li など、HTML の要素名に合わせたスタイル名で登録します。",
                en: "Registers styles named after HTML elements (h1–h6, p, ul-li, …)."
            },
            schemeWord: {
                ja: "Word の「見出し1〜5」「標準」「リスト段落」に対応するスタイル名で登録します。Word 原稿を配置するときにスタイルのマッピングが不要になります。",
                en: "Registers styles matching Word's Heading 1–5, Normal and List Paragraph, so placing a Word manuscript needs no style mapping."
            }
        },
        progress: {
            title:   { ja: "スタイル一括登録", en: "Register Styles" },
            styles:  { ja: "スタイルとグループを作成中…", en: "Creating styles and groups…" },
            attrs:   { ja: "属性を適用中…", en: "Applying attributes…" },
            grep:    { ja: "正規表現スタイルを設定中…", en: "Setting GREP styles…" },
            reorder: { ja: "並び替え中…", en: "Reordering…" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document before running." }
        },
        undo: {
            registerStyles: { ja: "スタイル一括登録", en: "Register Styles" }
        }
    };

    // =========================================
    // スタイル名の体系 / Style name schemes
    // =========================================

    /* 選べるスタイル名の体系 / Selectable style-name schemes */
    var STYLE_NAME_SCHEMES = {
        html: {},                /* 読み替えなし。キーがそのままスタイル名 / No mapping: the keys are the style names */
        word: WORD_STYLE_NAMES
    };

    /* 選択中の体系のスタイル名マップ。ダイアログの結果で差し替える /
       Style-name map of the selected scheme; replaced with the dialog's result */
    var activeStyleNameMap = STYLE_NAME_SCHEMES.html;

    /**
     * HTML 側の名前を、選択中の体系でのスタイル名に読み替える
     * ※ 読み替えの対象はグループに入れていないスタイルだけ。グループ内のスタイル
     *   （base-font、td-left、toc-h1、lang-US など）はどちらの体系でも名前を変えません /
     *   Only root-level styles are renamed; grouped styles (base-font, td-left, toc-h1, lang-US, …)
     *   keep the same name in both schemes.
     * @param {string} htmlStyleName HTML 側のスタイル名（例: "h1"）
     * @returns {string} 選択中の体系でのスタイル名。対応が無ければ引数をそのまま返す
     */
    function styleName(htmlStyleName) {
        var mappedName = activeStyleNameMap[htmlStyleName];
        return (typeof mappedName === "string") ? mappedName : htmlStyleName;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* ラジオボタンの下に出す見本に使う代表スタイル / Representative styles shown under each radio button */
    var SCHEME_SAMPLE_KEYS = ["h1", "p", "ul-li", "strong-bold"];

    /**
     * ラジオボタンの下に添える、代表スタイル名の見本を作る
     * @param {object} styleNameMap 体系のスタイル名マップ
     * @returns {string} 「 / 」区切りのスタイル名
     */
    function buildSchemeSample(styleNameMap) {
        var sampleNames = [];
        for (var sampleIndex = 0; sampleIndex < SCHEME_SAMPLE_KEYS.length; sampleIndex++) {
            var sampleKey = SCHEME_SAMPLE_KEYS[sampleIndex];
            var mappedName = styleNameMap[sampleKey];
            sampleNames.push((typeof mappedName === "string") ? mappedName : sampleKey);
        }
        return sampleNames.join(" / ");
    }

    /**
     * スタイル名の体系を選ぶダイアログを表示する
     * @returns {object|null} 選んだ体系のスタイル名マップ。キャンセルした場合は null
     */
    function showStyleSchemeDialog() {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dialog.opacity = DIALOG_OPACITY;
        setupWindow(dialog);

        var schemePanel = dialog.add("panel", undefined, getLabel("panel.styleNames"));
        setupPanel(schemePanel);

        /* ラジオは同じ親の直下どうしでしか排他にならないため、2つとも panel の直下に置く /
           Radio buttons are exclusive only among siblings, so both go directly under the panel */
        var rdoHtml = schemePanel.add("radiobutton", undefined, getLabel("scheme.html"));
        rdoHtml.helpTip = getLabel("tooltip.schemeHtml");
        var htmlSample = schemePanel.add("statictext", undefined, buildSchemeSample(STYLE_NAME_SCHEMES.html));
        htmlSample.indent = SCHEME_SAMPLE_INDENT;

        var rdoWord = schemePanel.add("radiobutton", undefined, getLabel("scheme.word"));
        rdoWord.helpTip = getLabel("tooltip.schemeWord");
        var wordSample = schemePanel.add("statictext", undefined, buildSchemeSample(STYLE_NAME_SCHEMES.word));
        wordSample.indent = SCHEME_SAMPLE_INDENT;

        rdoHtml.value = true;

        /* ボタン行（キャンセル → OK）/ Button row (Cancel, then OK) */
        var buttonRow = addButtonRow(dialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        if (dialog.show() !== 1) return null;
        return rdoWord.value ? STYLE_NAME_SCHEMES.word : STYLE_NAME_SCHEMES.html;
    }

    // =========================================
    // プログレスバー / Progress bar
    // =========================================

    /**
     * 進捗表示用のパレットを作る
     * @param {number} totalSteps 全体のステップ数
     * @returns {object} 更新と終了を行うオブジェクト
     */
    function createProgressWindow(totalSteps) {
        var progressWindow = new Window("palette", getLabel("progress.title") + "  " + SCRIPT_VERSION);
        setupWindow(progressWindow, 10);

        var progressMessage = progressWindow.add("statictext", undefined, "");
        progressMessage.preferredSize.width = PROGRESS_BAR_WIDTH;

        var progressBar = progressWindow.add("progressbar", undefined, 0, totalSteps);
        progressBar.preferredSize.width = PROGRESS_BAR_WIDTH;
        progressBar.preferredSize.height = PROGRESS_BAR_HEIGHT;

        progressWindow.show();

        return {
            step: function (message) {
                progressBar.value += 1;
                progressMessage.text = message;
                progressWindow.update();
            },
            close: function () {
                progressWindow.close();
            }
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * スタイルとグループを作成し、属性・GREP・並び順をまとめて適用する
     * @returns {void}
     */
    function main() {
        var doc = app.activeDocument;

        // =========================================
        // スタイル名定義 / Style name definitions
        // =========================================

        /* グループに入れないスタイルは styleName() を通す。選んだ体系に応じて
           HTML 名（h1 / p / …）と Word 名（Heading 1 / Normal / …）を切り替える /
           Root-level style names go through styleName(), which switches between the HTML names
           (h1 / p / …) and the Word names (Heading 1 / Normal / …) depending on the chosen scheme. */
        var paragraphStyleNames = [
            styleName("h1"), styleName("h2"), styleName("h3"), styleName("h4"), styleName("h5"), styleName("h6"),
            styleName("ul-li"), styleName("ol-li"),
            styleName("p"), styleName("p.caption"), styleName("p.code"), styleName("p.img"), styleName("p.table")
        ];

        var characterStyleNames = [
            styleName("strong-bold"), styleName("em-italic"),
            styleName("link"), styleName("code-normal"), styleName("code-strong"), styleName("highlighter")
        ];

        var paragraphStyleGroupNames = [
            "basestyle", "table", "toc", "book"
        ];

        var paragraphStylesInGroups = [
            { group: "basestyle", styles: ["base-font", "base-heading", "base-text", "base-table", "base-toc"] },
            { group: "table", styles: ["td-left", "td-justify", "td-justify-all", "td-center", "td-right", "td-left ul-li", "th-left", "th-center", "th-center-W"] },
            { group: "toc", styles: ["toc-title", "toc-h1", "toc-h2", "toc-h3"] },
            { group: "book", styles: ["page-number", "running-head", "thumb-index"] }
        ];

        var characterStyleGroupNames = [
            "table", "auto-apply"
        ];

        var characterStylesInGroups = [
            { group: "table", styles: ["td-bold"] },
            { group: "auto-apply", styles: ["no-break", "lang-US", "inline-graphic", "li-label", "li-bullet", "li-num"] }
        ];

        // =========================================
        // スタイル／グループ作成（既存ならスキップ） / Ensure styles & groups (skip if exists)
        // =========================================

        /* 今回のスクリプト実行で新規作成したスタイルを記録（属性適用の対象判定に使う）/
           コンテナ（doc またはグループ）ごとに区別するため container.id + styleName を複合キーにする /
           Record styles created during this run (used by attribute guards).
           Keyed by container.id + styleName so same-named styles in different groups don't collide. */
        var paragraphStyleKeysCreatedThisRun = {};
        var characterStyleKeysCreatedThisRun = {};

        /**
         * スタイルの所在を表す一意なキーを作る
         * @param {object} styleContainer スタイルのコンテナ
         * @param {string} styleName スタイル名
         * @returns {string} 識別キー
         */
        function styleContainerKey(styleContainer, styleName) {
            return styleContainer.id + "\t" + styleName;
        }

        /**
         * 段落スタイルグループを取得する（なければ作成）
         * @param {Document} doc 対象ドキュメント
         * @param {string} groupName グループ名
         * @returns {ParagraphStyleGroup} スタイルグループ
         */
        function ensureParagraphStyleGroup(doc, groupName) {
            var styleGroup = doc.paragraphStyleGroups.itemByName(groupName);
            if (!styleGroup.isValid) {
                styleGroup = doc.paragraphStyleGroups.add({ name: groupName });
            }
            return styleGroup;
        }

        /**
         * 文字スタイルグループを取得する（なければ作成）
         * @param {Document} doc 対象ドキュメント
         * @param {string} groupName グループ名
         * @returns {CharacterStyleGroup} スタイルグループ
         */
        function ensureCharacterStyleGroup(doc, groupName) {
            var styleGroup = doc.characterStyleGroups.itemByName(groupName);
            if (!styleGroup.isValid) {
                styleGroup = doc.characterStyleGroups.add({ name: groupName });
            }
            return styleGroup;
        }

        /**
         * 段落スタイルを取得する（なければ作成）
         * @param {object} styleContainer スタイルのコンテナ
         * @param {string} styleName スタイル名
         * @returns {ParagraphStyle} 段落スタイル
         */
        function ensureParagraphStyle(styleContainer, styleName) {
            var paragraphStyle = styleContainer.paragraphStyles.itemByName(styleName);
            if (!paragraphStyle.isValid) {
                paragraphStyle = styleContainer.paragraphStyles.add({ name: styleName });
                paragraphStyleKeysCreatedThisRun[styleContainerKey(styleContainer, styleName)] = true;
            }
            return paragraphStyle;
        }

        /**
         * 文字スタイルを取得する（なければ作成）
         * @param {object} styleContainer スタイルのコンテナ
         * @param {string} styleName スタイル名
         * @returns {CharacterStyle} 文字スタイル
         */
        function ensureCharacterStyle(styleContainer, styleName) {
            var characterStyle = styleContainer.characterStyles.itemByName(styleName);
            if (!characterStyle.isValid) {
                characterStyle = styleContainer.characterStyles.add({ name: styleName });
                characterStyleKeysCreatedThisRun[styleContainerKey(styleContainer, styleName)] = true;
            }
            return characterStyle;
        }

        /**
         * その段落スタイルへ属性を適用してよいかを判定する
         * @param {object} styleContainer スタイルのコンテナ
         * @param {string} styleName スタイル名
         * @returns {boolean} 適用してよければ true
         */
        function shouldApplyAttributesToParagraphStyle(styleContainer, styleName) {
            return OVERWRITE_EXISTING_STYLES ||
                paragraphStyleKeysCreatedThisRun[styleContainerKey(styleContainer, styleName)] === true;
        }

        /**
         * その文字スタイルへ属性を適用してよいかを判定する
         * @param {object} styleContainer スタイルのコンテナ
         * @param {string} styleName スタイル名
         * @returns {boolean} 適用してよければ true
         */
        function shouldApplyAttributesToCharacterStyle(styleContainer, styleName) {
            return OVERWRITE_EXISTING_STYLES ||
                characterStyleKeysCreatedThisRun[styleContainerKey(styleContainer, styleName)] === true;
        }

        // =========================================
        // 個別スタイルの属性適用 / Style-specific property settings
        // =========================================
        // ※ 属性・basedOn とも既存スタイルへの適用は OVERWRITE_EXISTING_STYLES 次第（shouldApplyAttributesTo* がガード）/
        //   Both attributes and basedOn are applied to existing styles only when OVERWRITE_EXISTING_STYLES is on
        //   (guarded by shouldApplyAttributesTo*).

        /**
         * 基準スタイルへ共通の組版設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyBaseGroupStyleSettings(doc) {
            var baseGroup = doc.paragraphStyleGroups.itemByName("basestyle");
            if (!baseGroup.isValid) return;

            if (shouldApplyAttributesToParagraphStyle(baseGroup, "base-text")) {
                var baseTextStyle = baseGroup.paragraphStyles.itemByName("base-text");
                if (baseTextStyle.isValid) {
                    setKerningMethodByNames(baseTextStyle, KERNING_METHOD_MOJIKUMI_NAMES);
                    baseTextStyle.justification = Justification.LEFT_JUSTIFIED;
                    baseTextStyle.keepLinesTogether = true;
                    baseTextStyle.keepAllLinesTogether = true;
                    baseTextStyle.hyphenation = false;
                }
            }

            if (shouldApplyAttributesToParagraphStyle(baseGroup, "base-heading")) {
                var headingStyle = baseGroup.paragraphStyles.itemByName("base-heading");
                if (headingStyle.isValid) {
                    setKerningMethodByNames(headingStyle, KERNING_METHOD_METRICS_NAMES);
                    headingStyle.justification = Justification.LEFT_ALIGN;
                    headingStyle.keepWithNext = 2;
                    headingStyle.keepLinesTogether = true;
                    headingStyle.keepAllLinesTogether = true;
                    headingStyle.hyphenation = false;
                }
            }

            if (shouldApplyAttributesToParagraphStyle(baseGroup, "base-toc")) {
                var tocStyle = baseGroup.paragraphStyles.itemByName("base-toc");
                if (tocStyle.isValid) {
                    tocStyle.justification = Justification.LEFT_ALIGN;
                    tocStyle.keepWithNext = 2;
                    tocStyle.keepLinesTogether = true;
                    tocStyle.keepAllLinesTogether = true;
                    tocStyle.hyphenation = false;
                }
            }
        }

        /**
         * 基準スタイルの継承関係（basedOn）を設定する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyBaseStyleBasedOn(doc) {
            var baseGroup = doc.paragraphStyleGroups.itemByName("basestyle");
            if (!baseGroup.isValid) return;
            var baseStyle = baseGroup.paragraphStyles.itemByName("base-font");
            if (!baseStyle.isValid) return;

            // base-font はすべての段落スタイルの親。見出し系は base-heading、本文系は base-text を経由する /
            // base-font is the root of every paragraph style; headings go through base-heading, body text through base-text
            // base-text → base-font（本文・リスト・表セル・目次項目の親）/ (parent of body, lists, table cells, TOC entries)
            var baseTextStyle = baseGroup.paragraphStyles.itemByName("base-text");
            if (baseTextStyle.isValid &&
                shouldApplyAttributesToParagraphStyle(baseGroup, "base-text")) {
                baseTextStyle.basedOn = baseStyle;
            }

            // base-heading → base-font（h1〜h6 と目次タイトルの親）/ (parent of h1–h6 and the TOC title)
            var headingStyle = baseGroup.paragraphStyles.itemByName("base-heading");
            if (headingStyle.isValid &&
                shouldApplyAttributesToParagraphStyle(baseGroup, "base-heading")) {
                headingStyle.basedOn = baseStyle;
            }

            // base-toc → base-text（目次項目 toc-h1〜toc-h3 の親）/ (parent of TOC entries toc-h1–toc-h3)
            var tocBaseStyle = baseGroup.paragraphStyles.itemByName("base-toc");
            if (tocBaseStyle.isValid && baseTextStyle.isValid &&
                shouldApplyAttributesToParagraphStyle(baseGroup, "base-toc")) {
                tocBaseStyle.basedOn = baseTextStyle;
            }

            // book グループ（ノンブル・柱・ツメ）→ base-font / book group (folio, running head, thumb index) → base-font
            var bookGroup = doc.paragraphStyleGroups.itemByName("book");
            if (bookGroup.isValid) {
                var bookStyleNames = ["page-number", "running-head", "thumb-index"];
                for (var bookIndex = 0; bookIndex < bookStyleNames.length; bookIndex++) {
                    if (!shouldApplyAttributesToParagraphStyle(bookGroup, bookStyleNames[bookIndex])) continue;
                    var bookStyle = bookGroup.paragraphStyles.itemByName(bookStyleNames[bookIndex]);
                    if (bookStyle.isValid) bookStyle.basedOn = baseStyle;
                }
            }

            // p / ul-li / ol-li / p.caption / p.code / p.img → base-text
            if (baseTextStyle.isValid) {
                var basedOnBaseStyleNames = [styleName("p"), styleName("ul-li"), styleName("ol-li"), styleName("p.caption"),
                                             styleName("p.code"), styleName("p.img")];
                for (var basedOnIndex = 0; basedOnIndex < basedOnBaseStyleNames.length; basedOnIndex++) {
                    var basedOnStyleName = basedOnBaseStyleNames[basedOnIndex];
                    if (!shouldApplyAttributesToParagraphStyle(doc, basedOnStyleName)) continue;
                    var basedOnTargetStyle = doc.paragraphStyles.itemByName(basedOnStyleName);
                    if (basedOnTargetStyle.isValid) basedOnTargetStyle.basedOn = baseTextStyle;
                }
            }

            // p.table → p
            var bodyParagraphStyle = doc.paragraphStyles.itemByName(styleName("p"));
            if (bodyParagraphStyle.isValid &&
                shouldApplyAttributesToParagraphStyle(doc, styleName("p.table"))) {
                var tableParagraphStyle = doc.paragraphStyles.itemByName(styleName("p.table"));
                if (tableParagraphStyle.isValid) tableParagraphStyle.basedOn = bodyParagraphStyle;
            }

            // h1〜h6 → base-heading
            if (headingStyle.isValid) {
                var headingBasedOnNames = [styleName("h1"), styleName("h2"), styleName("h3"),
                                           styleName("h4"), styleName("h5"), styleName("h6")];
                for (var headingBasedOnIndex = 0; headingBasedOnIndex < headingBasedOnNames.length; headingBasedOnIndex++) {
                    var headingBasedOnName = headingBasedOnNames[headingBasedOnIndex];
                    if (!shouldApplyAttributesToParagraphStyle(doc, headingBasedOnName)) continue;
                    var headingTargetStyle = doc.paragraphStyles.itemByName(headingBasedOnName);
                    if (headingTargetStyle.isValid) headingTargetStyle.basedOn = headingStyle;
                }
            }
        }

        /**
         * 段落スタイルの「次のスタイル」を設定する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyNextStyleSettings(doc) {
            var bodyParagraphStyle = doc.paragraphStyles.itemByName(styleName("p"));
            if (!bodyParagraphStyle.isValid) return;
            var nextStyleTargetNames = [styleName("h1"), styleName("h2"), styleName("h3"), styleName("h4"),
                                        styleName("h5"), styleName("h6"), styleName("p.caption")];
            for (var nextStyleIndex = 0; nextStyleIndex < nextStyleTargetNames.length; nextStyleIndex++) {
                var nextStyleTargetName = nextStyleTargetNames[nextStyleIndex];
                if (!shouldApplyAttributesToParagraphStyle(doc, nextStyleTargetName)) continue;
                var nextStyleTargetStyle = doc.paragraphStyles.itemByName(nextStyleTargetName);
                if (nextStyleTargetStyle.isValid) nextStyleTargetStyle.nextStyle = bodyParagraphStyle;
            }
        }

        /**
         * 段落の分離禁止に関する設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyKeepTogetherSettings(doc) {
            var keepWithPreviousStyleNames = [styleName("ul-li"), styleName("p.caption")];
            for (var keepWithPreviousIndex = 0; keepWithPreviousIndex < keepWithPreviousStyleNames.length; keepWithPreviousIndex++) {
                var keepWithPreviousName = keepWithPreviousStyleNames[keepWithPreviousIndex];
                if (!shouldApplyAttributesToParagraphStyle(doc, keepWithPreviousName)) continue;
                var keepWithPreviousStyle = doc.paragraphStyles.itemByName(keepWithPreviousName);
                if (keepWithPreviousStyle.isValid) {
                    keepWithPreviousStyle.keepWithPrevious = true;
                }
            }

            // p は分離禁止オプションをすべて OFF（base-text からの継承も含めて打ち消す）/
            // p turns off all keep options (also overriding what is inherited from base-text)
            if (shouldApplyAttributesToParagraphStyle(doc, styleName("p"))) {
                var bodyKeepStyle = doc.paragraphStyles.itemByName(styleName("p"));
                if (bodyKeepStyle.isValid) {
                    bodyKeepStyle.keepLinesTogether = false;
                    bodyKeepStyle.keepAllLinesTogether = false;
                    bodyKeepStyle.keepWithNext = 0;
                    bodyKeepStyle.keepWithPrevious = false;
                }
            }
        }

        /**
         * 画像用段落スタイルの設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyImageParagraphSettings(doc) {
            if (!shouldApplyAttributesToParagraphStyle(doc, styleName("p.img"))) return;
            var imageParagraphStyle = doc.paragraphStyles.itemByName(styleName("p.img"));
            if (imageParagraphStyle.isValid) {
                imageParagraphStyle.justification = Justification.CENTER_ALIGN;
            }
        }

        /**
         * ノンブル用段落スタイルの設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyPageNumberSettings(doc) {
            var bookGroup = doc.paragraphStyleGroups.itemByName("book");
            if (!bookGroup.isValid) return;
            if (!shouldApplyAttributesToParagraphStyle(bookGroup, "page-number")) return;
            var pageNumberStyle = bookGroup.paragraphStyles.itemByName("page-number");
            if (pageNumberStyle.isValid) {
                // 小口揃え / Align away from spine
                pageNumberStyle.justification = Justification.AWAY_FROM_BINDING_SIDE;
            }
        }

        /**
         * 箇条書き・番号リストの設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyListSettings(doc) {
            if (shouldApplyAttributesToParagraphStyle(doc, styleName("ul-li"))) {
                var bulletListStyle = doc.paragraphStyles.itemByName(styleName("ul-li"));
                if (bulletListStyle.isValid) applyBulletListSettings(doc, bulletListStyle);
            }
            if (shouldApplyAttributesToParagraphStyle(doc, styleName("ol-li"))) {
                var numberedListStyle = doc.paragraphStyles.itemByName(styleName("ol-li"));
                if (numberedListStyle.isValid) {
                    numberedListStyle.bulletsAndNumberingListType = ListType.NUMBERED_LIST;
                    var numberingCharacterStyle = resolveCharacterStyle(doc, "li-num");
                    if (numberingCharacterStyle) {
                        numberedListStyle.numberingCharacterStyle = numberingCharacterStyle.style;
                    }
                }
            }
        }

        /**
         * 箇条書きの設定（記号・同じスタイル間のスペース・タブ位置）を段落スタイルへ適用する
         * @param {Document} doc 対象ドキュメント
         * @param {ParagraphStyle} bulletListStyle 対象の段落スタイル
         * @returns {void}
         */
        function applyBulletListSettings(doc, bulletListStyle) {
            bulletListStyle.bulletsAndNumberingListType = ListType.BULLET_LIST;
            var bulletCharacterStyle = resolveCharacterStyle(doc, "li-bullet");
            if (bulletCharacterStyle) {
                bulletListStyle.bulletsCharacterStyle = bulletCharacterStyle.style;
            }
            // 同じスタイルが連続する段落間のスペースを 0 に（対応バージョンのみ。
            //   プロパティ名はバージョン差があるため候補から存在するものを設定）/
            // Space between paragraphs using the same style = 0 (only on supporting versions)
            setOptionalProperty(bulletListStyle,
                ["sameParaStyleSpacing", "spaceBetweenParagraphsUsingSameStyle", "spaceBetweenParagraphs", "spaceBetweenSameParagraphStyles", "spaceBetweenSameStyleParagraphs"], 0);
            setTabStopAtFontSize(bulletListStyle);
        }

        /**
         * 段落スタイルのタブ位置を、そのスタイルの文字サイズ（1字分）に置き直す
         * ※ 文字サイズの単位（pt / Q）に左右されないよう、読み書きの間だけスクリプトの単位をポイントにする /
         *   Script units are switched to points while reading and writing, so the text-size unit (pt / Q) doesn't matter.
         * @param {ParagraphStyle} paragraphStyle 対象の段落スタイル
         * @returns {void}
         */
        function setTabStopAtFontSize(paragraphStyle) {
            var savedMeasurementUnit = app.scriptPreferences.measurementUnit;
            app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
            try {
                for (var tabStopIndex = paragraphStyle.tabStops.length - 1; tabStopIndex >= 0; tabStopIndex--) {
                    paragraphStyle.tabStops[tabStopIndex].remove();
                }
                paragraphStyle.tabStops.add({ alignment: TabStopAlignment.LEFT_ALIGN, position: paragraphStyle.pointSize });
            } finally {
                app.scriptPreferences.measurementUnit = savedMeasurementUnit;
            }
        }

        /**
         * 表のスタイル（base-table と table グループの td-* / th-*）の継承関係と属性を適用する
         * ※ 関数内で basedOn → 属性の順に処理する。td-left ul-li のタブ位置は継承後の文字サイズで決まる /
         *   Handles basedOn before attributes internally, so td-left ul-li's tab stop uses the inherited font size.
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyTableStyleSettings(doc) {
            var tableGroup = doc.paragraphStyleGroups.itemByName("table");
            var baseGroup = doc.paragraphStyleGroups.itemByName("basestyle");
            if (!tableGroup.isValid || !baseGroup.isValid) return;

            // base-table: 表共通の基本テキスト（base-text → base-table、水平・垂直比率 92%）/
            // base-table: shared text for tables (base-text → base-table, horizontal/vertical scale 92%)
            var baseTableStyle = baseGroup.paragraphStyles.itemByName("base-table");
            if (!baseTableStyle.isValid) return;
            if (shouldApplyAttributesToParagraphStyle(baseGroup, "base-table")) {
                var baseTextStyle = baseGroup.paragraphStyles.itemByName("base-text");
                if (baseTextStyle.isValid) baseTableStyle.basedOn = baseTextStyle;
                baseTableStyle.horizontalScale = 92;
                baseTableStyle.verticalScale = 92;
            }

            // 親 → 子の順に並べる（子の basedOn を張る時点で親の basedOn が確定しているように）/
            // Listed parent-first so a parent's basedOn is settled before its children point at it
            // td-left: 本文セルの基準 / base of body cells, th-left: 表内の見出しの基準 / base of header cells
            // bulletList: 箇条書き（ul-li と同じ設定＋前の段落と連動）/ bullets (same as ul-li, plus keep with previous)
            var paperSwatch = doc.swatches.itemByName("Paper");
            var tableCellDefinitions = [
                { name: "td-left", parent: "base-table", justification: Justification.LEFT_ALIGN },
                { name: "td-justify", parent: "td-left", justification: Justification.LEFT_JUSTIFIED },
                { name: "td-justify-all", parent: "td-left", justification: Justification.FULLY_JUSTIFIED },
                { name: "td-center", parent: "td-left", justification: Justification.CENTER_ALIGN },
                { name: "td-right", parent: "td-left", justification: Justification.RIGHT_ALIGN },
                { name: "td-left ul-li", parent: "td-left", bulletList: true },
                { name: "th-left", parent: "base-table", justification: Justification.LEFT_ALIGN },
                { name: "th-center", parent: "th-left", justification: Justification.CENTER_ALIGN },
                { name: "th-center-W", parent: "th-left", justification: Justification.CENTER_ALIGN, fillColor: paperSwatch }
            ];
            for (var cellIndex = 0; cellIndex < tableCellDefinitions.length; cellIndex++) {
                var cellDefinition = tableCellDefinitions[cellIndex];
                if (!shouldApplyAttributesToParagraphStyle(tableGroup, cellDefinition.name)) continue;
                var cellStyle = tableGroup.paragraphStyles.itemByName(cellDefinition.name);
                if (!cellStyle.isValid) continue;
                var cellParentStyle = (cellDefinition.parent === "base-table")
                    ? baseTableStyle
                    : tableGroup.paragraphStyles.itemByName(cellDefinition.parent);
                if (cellParentStyle.isValid) cellStyle.basedOn = cellParentStyle;
                if (cellDefinition.justification) cellStyle.justification = cellDefinition.justification;
                if (cellDefinition.fillColor && cellDefinition.fillColor.isValid) {
                    cellStyle.fillColor = cellDefinition.fillColor;
                }
                if (cellDefinition.bulletList) {
                    applyBulletListSettings(doc, cellStyle);
                    cellStyle.keepWithPrevious = true;
                }
            }
        }

        /**
         * 目次見出しスタイルの継承関係を設定する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyTocSubheadingBasedOn(doc) {
            var tocGroup = doc.paragraphStyleGroups.itemByName("toc");
            if (!tocGroup.isValid) return;
            var baseGroup = doc.paragraphStyleGroups.itemByName("basestyle");
            if (!baseGroup.isValid) return;

            // toc-title は見出し系（base-heading）、toc-h1〜toc-h3 は本文系（base-toc → base-text）/
            // toc-title follows the headings (base-heading); toc-h1–toc-h3 follow body text (base-toc → base-text)
            var tocParentByName = {
                "toc-title": baseGroup.paragraphStyles.itemByName("base-heading"),
                "toc-h1": baseGroup.paragraphStyles.itemByName("base-toc"),
                "toc-h2": baseGroup.paragraphStyles.itemByName("base-toc"),
                "toc-h3": baseGroup.paragraphStyles.itemByName("base-toc")
            };
            for (var tocSubheadingName in tocParentByName) {
                var tocParentStyle = tocParentByName[tocSubheadingName];
                if (!tocParentStyle.isValid) continue;
                if (!shouldApplyAttributesToParagraphStyle(tocGroup, tocSubheadingName)) continue;
                var tocSubheadingStyle = tocGroup.paragraphStyles.itemByName(tocSubheadingName);
                if (tocSubheadingStyle.isValid) tocSubheadingStyle.basedOn = tocParentStyle;
            }
        }

        /**
         * 目次末端スタイルの個別設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyTocLeafOverrides(doc) {
            var tocGroup = doc.paragraphStyleGroups.itemByName("toc");
            if (!tocGroup.isValid) return;
            if (!shouldApplyAttributesToParagraphStyle(tocGroup, "toc-h3")) return;
            var tocH3Style = tocGroup.paragraphStyles.itemByName("toc-h3");
            if (tocH3Style.isValid) {
                tocH3Style.keepWithNext = 0;
            }
        }

        /**
         * 環境によって有無が変わるプロパティを安全に設定する
         * @param {object} targetObject 設定先のオブジェクト
         * @param {Array<string>} candidateNames 試すプロパティ名
         * @param {*} value 設定する値
         * @returns {string|null} 設定できたプロパティ名。どれも無ければ null
         */
        function setOptionalProperty(targetObject, candidateNames, value) {
            var availableProperties = targetObject.reflect.properties;
            for (var candidateIndex = 0; candidateIndex < candidateNames.length; candidateIndex++) {
                var candidateName = candidateNames[candidateIndex];
                for (var propertyIndex = 0; propertyIndex < availableProperties.length; propertyIndex++) {
                    if (String(availableProperties[propertyIndex].name) === candidateName) {
                        targetObject[candidateName] = value;
                        return candidateName;
                    }
                }
            }
            return null;
        }

        /**
         * 名前から文字スタイルを取得する（ルート → characterStyleGroupNames の各グループの順に探す）
         * @param {Document} doc 対象ドキュメント
         * @param {string} styleName 文字スタイル名
         * @returns {object|null} { style: 文字スタイル, container: 所属コンテナ }。見つからない場合は null
         */
        function resolveCharacterStyle(doc, styleName) {
            var rootStyle = doc.characterStyles.itemByName(styleName);
            if (rootStyle.isValid) return { style: rootStyle, container: doc };
            for (var groupNameIndex = 0; groupNameIndex < characterStyleGroupNames.length; groupNameIndex++) {
                var characterGroup = doc.characterStyleGroups.itemByName(characterStyleGroupNames[groupNameIndex]);
                if (!characterGroup.isValid) continue;
                var groupedStyle = characterGroup.characterStyles.itemByName(styleName);
                if (groupedStyle.isValid) return { style: groupedStyle, container: characterGroup };
            }
            return null;
        }

        /**
         * インライングラフィック用の前後アキを設定する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyInlineGraphicSpacing(doc) {
            var resolved = resolveCharacterStyle(doc, "inline-graphic");
            if (!resolved) return;
            if (!shouldApplyAttributesToCharacterStyle(resolved.container, "inline-graphic")) return;
            resolved.style.leadingAki = 0.25;
            resolved.style.trailingAki = 0.25;
        }

        /**
         * リンク用文字スタイルの設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyLinkSettings(doc) {
            var resolved = resolveCharacterStyle(doc, styleName("link"));
            if (!resolved) return;
            if (!shouldApplyAttributesToCharacterStyle(resolved.container, styleName("link"))) return;
            resolved.style.underline = false;
        }

        /**
         * 分割禁止の文字スタイル設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyNoBreakSettings(doc) {
            var resolved = resolveCharacterStyle(doc, "no-break");
            if (!resolved) return;
            if (!shouldApplyAttributesToCharacterStyle(resolved.container, "no-break")) return;
            resolved.style.noBreak = true;
        }

        /**
         * 文字スタイルの継承関係（basedOn）を設定する
         * @param {Document} doc 対象ドキュメント
         * @param {string} targetStyleName 対象のスタイル名
         * @param {string} parentStyleName 継承元のスタイル名
         * @returns {void}
         */
        function applyCharacterStyleBasedOn(doc, targetStyleName, parentStyleName) {
            var target = resolveCharacterStyle(doc, targetStyleName);
            var parent = resolveCharacterStyle(doc, parentStyleName);
            if (!target || !parent) return;
            if (!shouldApplyAttributesToCharacterStyle(target.container, targetStyleName)) return;
            target.style.basedOn = parent.style;
        }

        /**
         * 候補名から言語設定を解決する
         * @param {Array<string>} languageNames 言語名の候補
         * @returns {Language|null} 言語。見つからない場合は null
         */
        function resolveLanguageByNames(languageNames) {
            for (var languageNameIndex = 0; languageNameIndex < languageNames.length; languageNameIndex++) {
                var languageEntry = app.languagesWithVendors.itemByName(languageNames[languageNameIndex]);
                if (languageEntry.isValid) return languageEntry;
            }
            return null;
        }

        var ENGLISH_USA_LANGUAGE_NAMES = ["English: USA", "英語：米国"];
        var NO_LANGUAGE_NAMES = ["[No Language]", "[言語なし]", "[なし]"];

        /* カーニング方式は UI 言語で表示名が変わるため候補を順に試す /
           Kerning method names are localized, so try the candidates in order */
        var KERNING_METHOD_MOJIKUMI_NAMES = ["和文等幅", "Japanese Mojikumi"];
        var KERNING_METHOD_METRICS_NAMES  = ["メトリクス", "Metrics"];

        /**
         * 候補名からカーニング方式を設定する
         * @param {ParagraphStyle} paragraphStyle 対象の段落スタイル
         * @param {Array<string>} kerningMethodNames カーニング方式名の候補
         * @returns {boolean} 設定できたら true
         */
        function setKerningMethodByNames(paragraphStyle, kerningMethodNames) {
            for (var kerningMethodIndex = 0; kerningMethodIndex < kerningMethodNames.length; kerningMethodIndex++) {
                try {
                    paragraphStyle.kerningMethod = kerningMethodNames[kerningMethodIndex];
                    return true;
                } catch (e) {
                    // この環境には無い表示名。次の候補を試す / Not available in this locale; try the next candidate
                }
            }
            return false;
        }

        /**
         * lang-US スタイルに英語（米国）を設定する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyLangUSLanguageSetting(doc) {
            var resolved = resolveCharacterStyle(doc, "lang-US");
            if (!resolved) return;
            if (!shouldApplyAttributesToCharacterStyle(resolved.container, "lang-US")) return;
            var englishLanguage = resolveLanguageByNames(ENGLISH_USA_LANGUAGE_NAMES);
            if (englishLanguage) resolved.style.appliedLanguage = englishLanguage;
        }

        /**
         * コード用文字スタイルの言語設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyCodeNormalLanguageSetting(doc) {
            var resolved = resolveCharacterStyle(doc, styleName("code-normal"));
            if (!resolved) return;
            if (!shouldApplyAttributesToCharacterStyle(resolved.container, styleName("code-normal"))) return;
            var noLanguage = resolveLanguageByNames(NO_LANGUAGE_NAMES);
            if (noLanguage) resolved.style.appliedLanguage = noLanguage;
            resolved.style.ligatures = false;
        }

        /**
         * コード用段落スタイルの設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyCodeParagraphSettings(doc) {
            if (!shouldApplyAttributesToParagraphStyle(doc, styleName("p.code"))) return;
            var codeParagraphStyle = doc.paragraphStyles.itemByName(styleName("p.code"));
            if (!codeParagraphStyle.isValid) return;
            var noLanguage = resolveLanguageByNames(NO_LANGUAGE_NAMES);
            if (noLanguage) codeParagraphStyle.appliedLanguage = noLanguage;
            codeParagraphStyle.ligatures = false;
            codeParagraphStyle.justification = Justification.LEFT_ALIGN;
            codeParagraphStyle.hyphenation = false;
        }

        /**
         * すべてのスタイル属性の適用処理をまとめて実行する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyAllStyleAttributes(doc) {
            // 先に継承関係（basedOn）を確定させてから属性を適用する。逆順だと、親と同値の代入が
            //   override として残らず、後から張った basedOn の継承値で打ち消されることがある
            //   （例: p の分離禁止 OFF が base-text の ON に戻る）/
            // Set inheritance (basedOn) first, then attributes: assigning a value equal to the parent's
            //   may not register as an override, so a basedOn applied afterwards can undo it
            //   (e.g. p's keep options going back to base-text's ON).
            // ※ applyTableStyleSettings は関数内で basedOn → 属性の順になっているため、この並びのままでよい /
            //   applyTableStyleSettings already does basedOn → attributes internally, so it stays put.
            applyBaseStyleBasedOn(doc);
            applyTocSubheadingBasedOn(doc);
            applyCharacterStyleBasedOn(doc, styleName("highlighter"), styleName("strong-bold"));
            applyCharacterStyleBasedOn(doc, styleName("code-strong"), styleName("code-normal"));
            applyCharacterStyleBasedOn(doc, "li-label", styleName("strong-bold"));
            applyCharacterStyleBasedOn(doc, "td-bold", styleName("strong-bold"));

            applyBaseGroupStyleSettings(doc);
            applyNextStyleSettings(doc);
            applyKeepTogetherSettings(doc);
            applyListSettings(doc);
            applyImageParagraphSettings(doc);
            applyPageNumberSettings(doc);
            applyTableStyleSettings(doc);
            applyTocLeafOverrides(doc);
            applyInlineGraphicSpacing(doc);
            applyLinkSettings(doc);
            applyNoBreakSettings(doc);
            applyLangUSLanguageSetting(doc);
            applyCodeNormalLanguageSetting(doc);
            applyCodeParagraphSettings(doc);
        }

        // =========================================
        // 正規表現スタイル（ネスト GREP） / Nested GREP styles
        // =========================================

        /**
         * 見出し・本文・リスト・表セルに正規表現スタイルを設定する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyNestedGrepStyleSettings(doc) {
            // group: 段落スタイルの所属グループ名（null はルート）/ owning group name (null = root)
            // 割り当て / Assignment:
            //   strong-bold    … base-heading（h1〜h6・toc-title へ継承）、th-left（th-center・th-center-W へ継承）
            //   lang-US        … base-heading、p、ul-li、ol-li
            //   no-break       … p、ul-li、ol-li
            //   inline-graphic … p、ul-li、ol-li、base-table、th-left
            //   li-label       … ul-li のみ
            //   p.table は p を、td-* は base-table を継承する（own GREP を持たないので継承される）。
            //   th-left は太字を持つので base-table の継承が切れ、アンカーも直接持たせる /
            //   p.table inherits from p, td-* from base-table (they have no own GREP, so the rules carry over).
            //   th-left carries its own bold rule, which cuts off base-table's, so it also gets the anchor directly.
            // 独自の GREP を1つでも持つと InDesign は GREP の継承を切る（own リストが継承分を置き換える）ため、
            //   GREP を持つスタイルには必要なものをすべて直接設定する。継承は切れているので二重にはならない /
            //   Once a style has any own GREP, InDesign stops inheriting (the own list replaces the inherited one),
            //   so every style that carries GREP gets all of its rules directly. No duplication, since inheritance is off.
            //   ※ 手動で GREP を足すと UI が継承分を own へコピーしてから追加するので継承が残って見えるが、
            //     スクリプトの nestedGrepStyles.add() はコピーしないため継承が切れる /
            //   NOTE: manual add copies inherited rules into the own list first (so they appear to persist), but
            //     scripted add() does not copy them, so inheritance is severed.
            //   li-label は最後に置き、重なる範囲で優先させる / li-label is last so it wins on overlapping ranges.
            var GREP_BOLD = "(?#太字)(.+)";
            var GREP_LANG_US = "(?#欧文)[\\u\\l]";
            var GREP_NO_BREAK = "(?#行末分離禁止)..[。」』？！…]?$";
            var GREP_INLINE_GRAPHIC = "(?#アンカー)~a";
            var GREP_LI_LABEL = "(?#ラベル)^.+?(?=：)";
            var nestedGrepRules = [
                { group: "basestyle", paragraph: "base-heading", character: styleName("strong-bold"), expression: GREP_BOLD },
                { group: "basestyle", paragraph: "base-heading", character: "lang-US", expression: GREP_LANG_US },
                { group: null, paragraph: styleName("p"), character: "lang-US", expression: GREP_LANG_US },
                { group: null, paragraph: styleName("p"), character: "no-break", expression: GREP_NO_BREAK },
                { group: null, paragraph: styleName("p"), character: "inline-graphic", expression: GREP_INLINE_GRAPHIC },
                { group: null, paragraph: styleName("ul-li"), character: "lang-US", expression: GREP_LANG_US },
                { group: null, paragraph: styleName("ul-li"), character: "no-break", expression: GREP_NO_BREAK },
                { group: null, paragraph: styleName("ul-li"), character: "inline-graphic", expression: GREP_INLINE_GRAPHIC },
                { group: null, paragraph: styleName("ul-li"), character: "li-label", expression: GREP_LI_LABEL },
                { group: null, paragraph: styleName("ol-li"), character: "lang-US", expression: GREP_LANG_US },
                { group: null, paragraph: styleName("ol-li"), character: "no-break", expression: GREP_NO_BREAK },
                { group: null, paragraph: styleName("ol-li"), character: "inline-graphic", expression: GREP_INLINE_GRAPHIC },
                { group: "basestyle", paragraph: "base-table", character: "inline-graphic", expression: GREP_INLINE_GRAPHIC },
                { group: "table", paragraph: "th-left", character: styleName("strong-bold"), expression: GREP_BOLD },
                { group: "table", paragraph: "th-left", character: "inline-graphic", expression: GREP_INLINE_GRAPHIC }
            ];

            // 置き換えモード（OVERWRITE_EXISTING_STYLES）では、各対象スタイルの既存 GREP を
            //   一度だけ全削除してから付け直す（同じ実行内で重複削除しないよう id で記録）/
            //   In replace mode, clear each target style's existing GREP once before re-adding.
            var grepClearedStyleIds = {};

            for (var grepRuleIndex = 0; grepRuleIndex < nestedGrepRules.length; grepRuleIndex++) {
                var grepRuleDefinition = nestedGrepRules[grepRuleIndex];
                // 段落スタイルのコンテナを解決（グループ指定があればそのグループ、無ければ doc）/
                // Resolve the paragraph style container (group if specified, otherwise doc)
                var paragraphContainer = grepRuleDefinition.group
                    ? doc.paragraphStyleGroups.itemByName(grepRuleDefinition.group)
                    : doc;
                if (!paragraphContainer.isValid) continue;
                if (!shouldApplyAttributesToParagraphStyle(paragraphContainer, grepRuleDefinition.paragraph)) continue;
                var targetParagraphStyle = paragraphContainer.paragraphStyles.itemByName(grepRuleDefinition.paragraph);
                // 文字スタイルはルート → 各グループの順で解決 / Resolve character style across root and the style groups
                var resolvedCharacter = resolveCharacterStyle(doc, grepRuleDefinition.character);
                if (!targetParagraphStyle.isValid || !resolvedCharacter) continue;

                // 置き換えモードでは、このスタイルの既存 GREP を一度だけ全削除（末尾から削除して添字ずれ回避）/
                // In replace mode, clear this style's existing GREP once (remove from the end to avoid index shift)
                if (OVERWRITE_EXISTING_STYLES && !grepClearedStyleIds[targetParagraphStyle.id]) {
                    for (var grepClearIndex = targetParagraphStyle.nestedGrepStyles.length - 1; grepClearIndex >= 0; grepClearIndex--) {
                        targetParagraphStyle.nestedGrepStyles[grepClearIndex].remove();
                    }
                    grepClearedStyleIds[targetParagraphStyle.id] = true;
                }

                var targetCharacterStyle = resolvedCharacter.style;
                var hasSameGrepStyle = false;
                for (var nestedGrepStyleIndex = 0; nestedGrepStyleIndex < targetParagraphStyle.nestedGrepStyles.length; nestedGrepStyleIndex++) {
                    var existingGrepStyle = targetParagraphStyle.nestedGrepStyles[nestedGrepStyleIndex];
                    // 文字スタイルは名前ではなく一意な id で比較（別グループの同名スタイルと誤判定しない）/
                    // Compare applied character style by unique id, not name (avoids same-name collisions across groups)
                    var existingCharacterStyle = existingGrepStyle.appliedCharacterStyle;
                    if (existingGrepStyle.grepExpression === grepRuleDefinition.expression &&
                        existingCharacterStyle.isValid &&
                        existingCharacterStyle.id === targetCharacterStyle.id) {
                        hasSameGrepStyle = true;
                        break;
                    }
                }
                if (!hasSameGrepStyle) {
                    var newGrepStyle = targetParagraphStyle.nestedGrepStyles.add();
                    newGrepStyle.appliedCharacterStyle = targetCharacterStyle;
                    newGrepStyle.grepExpression = grepRuleDefinition.expression;
                }
            }
        }

        // =========================================
        // パネル上の並び替え / Reorder styles in the panel
        // =========================================

        /**
         * 段落スタイルの並び順を整える
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function reorderParagraphStyles(doc) {
            // ルート段落スタイルを配列順に末尾へ移動 / Move root styles to the end in array order
            for (var rootIndex = 0; rootIndex < paragraphStyleNames.length; rootIndex++) {
                var rootStyle = doc.paragraphStyles.itemByName(paragraphStyleNames[rootIndex]);
                if (rootStyle.isValid) rootStyle.move(LocationOptions.AT_END, doc);
            }
            // 段落スタイルグループを配列順に末尾へ移動 / Move groups to the end in array order
            for (var groupIndex = 0; groupIndex < paragraphStyleGroupNames.length; groupIndex++) {
                var styleGroup = doc.paragraphStyleGroups.itemByName(paragraphStyleGroupNames[groupIndex]);
                if (styleGroup.isValid) styleGroup.move(LocationOptions.AT_END, doc);
            }
            // 各グループ内の段落スタイルを配列順に末尾へ移動 / Move grouped styles to the end of each group
            for (var entryIndex = 0; entryIndex < paragraphStylesInGroups.length; entryIndex++) {
                var groupEntry = paragraphStylesInGroups[entryIndex];
                var targetGroup = doc.paragraphStyleGroups.itemByName(groupEntry.group);
                if (!targetGroup.isValid) continue;
                for (var styleNameIndex = 0; styleNameIndex < groupEntry.styles.length; styleNameIndex++) {
                    var groupedStyle = targetGroup.paragraphStyles.itemByName(groupEntry.styles[styleNameIndex]);
                    if (groupedStyle.isValid) groupedStyle.move(LocationOptions.AT_END, targetGroup);
                }
            }
        }

        /**
         * 文字スタイルの並び順を整える
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function reorderCharacterStyles(doc) {
            // ルート文字スタイルを配列順に末尾へ移動 / Move root styles to the end in array order
            for (var rootCharacterIndex = 0; rootCharacterIndex < characterStyleNames.length; rootCharacterIndex++) {
                var rootCharacter = doc.characterStyles.itemByName(characterStyleNames[rootCharacterIndex]);
                if (rootCharacter.isValid) rootCharacter.move(LocationOptions.AT_END, doc);
            }
            // 文字スタイルグループを配列順に末尾へ移動 / Move groups to the end in array order
            for (var characterGroupOrderIndex = 0; characterGroupOrderIndex < characterStyleGroupNames.length; characterGroupOrderIndex++) {
                var characterGroup = doc.characterStyleGroups.itemByName(characterStyleGroupNames[characterGroupOrderIndex]);
                if (characterGroup.isValid) characterGroup.move(LocationOptions.AT_END, doc);
            }
            // 各グループ内の文字スタイルを配列順に末尾へ移動 / Move grouped styles to the end of each group
            for (var characterEntryIndex = 0; characterEntryIndex < characterStylesInGroups.length; characterEntryIndex++) {
                var characterGroupEntry = characterStylesInGroups[characterEntryIndex];
                var targetCharacterGroup = doc.characterStyleGroups.itemByName(characterGroupEntry.group);
                if (!targetCharacterGroup.isValid) continue;
                for (var characterNameIndex = 0; characterNameIndex < characterGroupEntry.styles.length; characterNameIndex++) {
                    var groupedCharacter = targetCharacterGroup.characterStyles.itemByName(characterGroupEntry.styles[characterNameIndex]);
                    if (groupedCharacter.isValid) groupedCharacter.move(LocationOptions.AT_END, targetCharacterGroup);
                }
            }
        }

        // =========================================
        // メイン処理 / Main execution
        // =========================================

        // 処理中はプログレスバーを表示（作成→属性→GREP→並び替えの 4 段階）/
        // Show a progress bar while processing (4 phases: create → attributes → GREP → reorder)
        var progress = createProgressWindow(4);
        try {

        progress.step(getLabel("progress.styles"));

        for (var paragraphGroupIndex = 0; paragraphGroupIndex < paragraphStyleGroupNames.length; paragraphGroupIndex++) {
            ensureParagraphStyleGroup(doc, paragraphStyleGroupNames[paragraphGroupIndex]);
        }

        for (var characterGroupIndex = 0; characterGroupIndex < characterStyleGroupNames.length; characterGroupIndex++) {
            ensureCharacterStyleGroup(doc, characterStyleGroupNames[characterGroupIndex]);
        }

        for (var paragraphStyleIndex = 0; paragraphStyleIndex < paragraphStyleNames.length; paragraphStyleIndex++) {
            ensureParagraphStyle(doc, paragraphStyleNames[paragraphStyleIndex]);
        }

        for (var paragraphGroupStyleIndex = 0; paragraphGroupStyleIndex < paragraphStylesInGroups.length; paragraphGroupStyleIndex++) {
            var paragraphGroupEntry = paragraphStylesInGroups[paragraphGroupStyleIndex];
            var paragraphStyleGroup = ensureParagraphStyleGroup(doc, paragraphGroupEntry.group);
            for (var paragraphStyleNameIndex = 0; paragraphStyleNameIndex < paragraphGroupEntry.styles.length; paragraphStyleNameIndex++) {
                ensureParagraphStyle(paragraphStyleGroup, paragraphGroupEntry.styles[paragraphStyleNameIndex]);
            }
        }

        for (var characterStyleIndex = 0; characterStyleIndex < characterStyleNames.length; characterStyleIndex++) {
            ensureCharacterStyle(doc, characterStyleNames[characterStyleIndex]);
        }

        for (var characterGroupStyleIndex = 0; characterGroupStyleIndex < characterStylesInGroups.length; characterGroupStyleIndex++) {
            var characterGroupEntry = characterStylesInGroups[characterGroupStyleIndex];
            var characterStyleGroup = ensureCharacterStyleGroup(doc, characterGroupEntry.group);
            for (var characterStyleNameIndex = 0; characterStyleNameIndex < characterGroupEntry.styles.length; characterStyleNameIndex++) {
                ensureCharacterStyle(characterStyleGroup, characterGroupEntry.styles[characterStyleNameIndex]);
            }
        }

        progress.step(getLabel("progress.attrs"));
        applyAllStyleAttributes(doc);

        progress.step(getLabel("progress.grep"));
        applyNestedGrepStyleSettings(doc);

        // パネル上の並び順を配列順に揃える（既存スタイルも含む） /
        // Reorder styles in the panel (including existing ones)
        progress.step(getLabel("progress.reorder"));
        reorderParagraphStyles(doc);
        reorderCharacterStyles(doc);

        } finally {
            // 例外時もプログレスパレットを確実に閉じる / Always close the palette, even on error
            progress.close();
        }
    }

    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    // スタイル名の体系をダイアログで決めてから本処理へ。キャンセルならアンドゥ単位を作らずに終える /
    // Pick the style-name scheme first; cancelling exits without creating an undo step
    var selectedStyleNameMap = showStyleSchemeDialog();
    if (!selectedStyleNameMap) return;
    activeStyleNameMap = selectedStyleNameMap;

    // 全処理を 1 つのアンドゥ単位にまとめて実行 /
    // Run everything as a single undo step
    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined,
        UndoModes.ENTIRE_SCRIPT, getLabel("undo.registerStyles"));

})();
