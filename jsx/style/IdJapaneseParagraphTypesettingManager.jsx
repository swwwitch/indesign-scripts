#target indesign

/*

### 概要

段落スタイルの日本語組版設定（禁則処理セット・禁則調整方式・文字組みアキ量・コンポーザー）をマトリックス UI で確認・一括適用します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdJapaneseParagraphTypesettingManager.md

### Overview

Reviews and batch-applies the Japanese composition settings of paragraph styles (kinsoku set, kinsoku adjustment, mojikumi and composer) through a matrix UI.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdJapaneseParagraphTypesettingManager.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdJapaneseParagraphTypesettingManager"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";          /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-05";                           /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                           /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdJapaneseParagraphTypesettingManager.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdJapaneseParagraphTypesettingManager.md"; /* README (English) */

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

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 既定で選択する設定名 / Values preselected on launch */
var DEFAULT_KINSOKU_SET_NAME  = "弱い禁則";
var DEFAULT_KINSOKU_TYPE_NAME = "調整量を優先";
var DEFAULT_MOJIKUMI_NAME     = "行末約物半角";
var DEFAULT_COMPOSER_NAME     = "日本語単数行コンポーザー";

/* 対象から除外する段落スタイルグループの接頭辞 / Prefix marking style groups to skip */
var EXCLUDED_STYLE_GROUP_PREFIX = "_";

// =========================================
// レイアウト設定 / Layout settings
// =========================================

/* マトリックス UI の列幅（px）/ Column widths of the matrix UI (px) */
var COLUMN_WIDTH_NAME      = 160;  /* 段落スタイル名 / paragraph style name */
var COLUMN_WIDTH_DROPDOWN  = 120;  /* 禁則処理セット・禁則調整方式 / kinsoku set and adjustment */
var COLUMN_WIDTH_DROPDOWN_WIDE = 200; /* 文字組み・コンポーザー / mojikumi and composer */

/* マトリックス各行の要素間隔（px）/ Spacing between controls in a matrix row (px) */
var MATRIX_ROW_SPACING = 8;

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
        title: { ja: "日本語文字組版設定", en: "Japanese Typesetting Settings" }
    },
    panel: {
        typesetting: { ja: "組版設定", en: "Typesetting" }
    },
    button: {
        ok:      { ja: "OK", en: "OK" },
        cancel:  { ja: "キャンセル", en: "Cancel" },
        reflect: { ja: "↓ 反映", en: "↓ Apply" }
    },
    column: {
        styleName:   { ja: "段落スタイル", en: "Paragraph Style" },
        kinsokuSet:  { ja: "禁則処理セット", en: "Kinsoku Set" },
        kinsokuType: { ja: "禁則調整方式", en: "Kinsoku Adjustment" },
        mojikumi:    { ja: "文字組み", en: "Mojikumi" },
        composer:    { ja: "コンポーザー", en: "Composer" },
        bulkSource:  { ja: "すべて", en: "All" }
    },
    alert: {
        noDocument:      { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document before running." },
        noKinsokuTables: { ja: "このドキュメントには禁則処理セットがありません。", en: "This document has no kinsoku tables." },
        noParagraphStyles: { ja: "適用可能な段落スタイルがありません。", en: "There are no applicable paragraph styles." },
        partialFailurePrefix: { ja: "適用しましたが、", en: "Applied, but " },
        partialFailureSuffix: { ja: " 件の段落スタイルでエラーが発生しました。", en: " paragraph style(s) reported an error." }
    },
    undo: {
        applyTypesetting: { ja: "日本語文字組版設定の適用", en: "Apply Japanese Typesetting Settings" }
    }
};

// =========================================
// 文字組みアキ量プリセット定義 / Mojikumi preset definitions
// =========================================

/* 文字組み「なし」の表示名（プリセットや読み取り結果と名前で照合する） / Display name of "no mojikumi" (matched by name against presets and read values) */
var MOJIKUMI_NONE_NAME = "なし";

/*
組み込みの文字組みアキ量 preset enum と表示名の対応表。
InDesign の enum 名から UI 表示用ラベルへ変換するために使用。

Mapping table between built-in mojikumi preset enums and display labels.
Used to convert internal InDesign enum names into user-facing labels.
*/
var MOJIKUMI_LABELS = {
    "LINE_END_ALL_ONE_HALF_EM_ENUM": "行末約物半角",
    "ONE_EM_INDENT_LINE_END_UKE_ONE_HALF_EM_ENUM": "行末受け約物半角・段落1字下げ（起こし全角）",
    "ONE_OR_ONE_HALF_EM_INDENT_LINE_END_UKE_ONE_HALF_EM_ENUM": "行末受け約物半角・段落1字下げ（起こし食い込み）",
    "ONE_OR_ONE_HALF_EM_INDENT_LINE_END_ALL_ONE_EM_ENUM": "約物全角・段落1字下げ",
    "ONE_EM_INDENT_LINE_END_ALL_ONE_EM_ENUM": "約物全角・段落1字下げ（起こし全角）",
    "ONE_EM_INDENT_LINE_END_ALL_NO_FLOAT_ENUM": "行末約物全角/半角・段落1字下げ",
    "ONE_EM_INDENT_LINE_END_UKE_NO_FLOAT_ENUM": "行末受け約物全角／半角・段落1字下げ（起こし全角）",
    "ONE_OR_ONE_HALF_EM_INDENT_LINE_END_UKE_NO_FLOAT_ENUM": "行末受け約物全角／半角・段落1字下げ（起こし食い込み）",
    "ONE_EM_INDENT_LINE_END_ALL_ONE_HALF_EM_ENUM": "行末約物半角・段落1字下げ",
    "LINE_END_ALL_ONE_EM_ENUM": "約物全角",
    "LINE_END_UKE_NO_FLOAT_ENUM": "行末受け約物全角／半角",
    "ONE_OR_ONE_HALF_EM_INDENT_LINE_END_PERIOD_ONE_EM_ENUM": "行末句点全角・段落1字下げ",
    "ONE_EM_INDENT_LINE_END_PERIOD_ONE_EM_ENUM": "行末句点全角・段落1字下げ（起こし全角）",
    "LINE_END_PERIOD_ONE_EM_ENUM": "行末句点全角",
    "TRAD_CHINESE_DEFAULT": "繁体字中国語デフォルト",
    "SIMP_CHINESE_DEFAULT": "簡体字中国語デフォルト"
};

// =========================================
// 配列の検索 / Array lookup
// =========================================

/**
 * 配列の中で値が一致する（===）最初の位置を探す
 * @param {Array} items 探す配列（表示名・列挙値・オブジェクト参照など）
 * @param {*} target 探す値
 * @returns {number} 見つかった位置。なければ -1
 */
function findIndexInArray(items, target) {
    for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
        if (items[itemIndex] === target) return itemIndex;
    }
    return -1;
}

/**
 * 既定値の名前から選択位置を求める（見つからなければ先頭）
 * @param {Array<string>} names 表示名の一覧
 * @param {string} defaultName 既定値の名前
 * @returns {number} 選択する位置
 */
function getDefaultIndexByName(names, defaultName) {
    var foundIndex = findIndexInArray(names, defaultName);
    return foundIndex >= 0 ? foundIndex : 0;
}

// =========================================
// ドキュメント情報の取得 / Document data collection
// =========================================

/**
 * ドキュメント内の禁則処理セットを集める
 * @param {Document} documentObject 対象ドキュメント
 * @returns {{tables: Array, names: Array<string>}} 禁則処理セットと表示名
 */
function collectKinsokuTables(documentObject) {
    var tables = [];
    var names = [];
    for (var kinsokuTableIndex = 0; kinsokuTableIndex < documentObject.kinsokuTables.length; kinsokuTableIndex++) {
        var kinsokuTable = documentObject.kinsokuTables.item(kinsokuTableIndex);
        tables.push(kinsokuTable);
        names.push(kinsokuTable.name);
    }
    return { tables: tables, names: names };
}

/**
 * 禁則調整方式の選択肢を作る
 * @returns {{values: Array, names: Array<string>}} 調整方式と表示名
 */
function createKinsokuTypeOptions() {
    return {
        names: ["追い込み優先", "追い出し優先", "追い出しのみ", "調整量を優先"],
        values: [
            KinsokuType.KINSOKU_PUSH_IN_FIRST,
            KinsokuType.KINSOKU_PUSH_OUT_FIRST,
            KinsokuType.KINSOKU_PUSH_OUT_ONLY,
            KinsokuType.KINSOKU_PRIORITIZE_ADJUSTMENT_AMOUNT
        ]
    };
}

/**
 * ドキュメント内の文字組みアキ量設定を集める（先頭は「なし」）
 * @param {Document} documentObject 対象ドキュメント
 * @returns {{tables: Array, names: Array<string>}} 文字組み設定と表示名（「なし」の設定は null）
 */
function collectMojikumiTables(documentObject) {
    var tables = [null];
    var names = [MOJIKUMI_NONE_NAME];
    for (var mojikumiTableIndex = 0; mojikumiTableIndex < documentObject.mojikumiTables.length; mojikumiTableIndex++) {
        var mojikumiTable = documentObject.mojikumiTables.item(mojikumiTableIndex);
        tables.push(mojikumiTable);
        names.push(mojikumiTable.name);
    }
    return { tables: tables, names: names };
}

/**
 * 対象にする段落スタイルをグループ込みで集める
 * @param {Document} documentObject 対象ドキュメント
 * @returns {{styles: Array<ParagraphStyle>, names: Array<string>}} 段落スタイルと表示名
 */
function collectTargetParagraphStyles(documentObject) {
    var styles = [];
    var names = [];

    /**
     * スタイルグループを再帰的にたどって段落スタイルを集める
     * @param {object} container 段落スタイルのコンテナ
     * @param {string} prefix グループ名の接頭辞
     * @returns {void}
     */
    function walk(container, prefix) {
        for (var paragraphStyleIndex = 0; paragraphStyleIndex < container.paragraphStyles.length; paragraphStyleIndex++) {
            var paragraphStyle = container.paragraphStyles.item(paragraphStyleIndex);
            var paragraphStyleName = paragraphStyle.name;
            if (paragraphStyleName === "[段落スタイルなし]" || paragraphStyleName === "[No Paragraph Style]") continue;
            if (paragraphStyleName === "[基本段落]" || paragraphStyleName === "[Basic Paragraph]") continue;
            styles.push(paragraphStyle);
            names.push(prefix + paragraphStyleName);
        }
        for (var styleGroupIndex = 0; styleGroupIndex < container.paragraphStyleGroups.length; styleGroupIndex++) {
            var styleGroup = container.paragraphStyleGroups.item(styleGroupIndex);
            if (styleGroup.name.charAt(0) === EXCLUDED_STYLE_GROUP_PREFIX) continue;
            walk(styleGroup, prefix + styleGroup.name + " / ");
        }
    }

    walk(documentObject, "");
    return { styles: styles, names: names };
}

/**
 * コンポーザーの選択肢と適用用エイリアスを作る
 * @returns {{names: Array<string>, aliases: Array<Array<string>>}} 表示名と、ロケール・バージョン違いの名前の候補
 */
function createComposerOptions() {
    var entries = [
        { name: "日本語段落コンポーザー", aliases: ["Adobe 日本語段落コンポーザー", "Adobe Japanese Paragraph Composer"] },
        { name: "日本語単数行コンポーザー", aliases: ["Adobe 日本語単数行コンポーザー", "Adobe Japanese Single-line Composer"] },
        { name: "多言語対応段落コンポーザー", aliases: ["$ID/HL Composer Optyca", "Adobe World-Ready Paragraph Composer", "Adobe 多言語対応段落コンポーザー", "Adobe World-Ready 段落コンポーザー"] },
        { name: "多言語対応単数行コンポーザー", aliases: ["$ID/HL Single Optyca", "Adobe World-Ready Single-line Composer", "Adobe 多言語対応単数行コンポーザー", "Adobe World-Ready 単数行コンポーザー"] },
        { name: "欧文段落コンポーザー", aliases: ["$ID/HL Composer", "Adobe Paragraph Composer", "Adobe 欧文段落コンポーザー", "Adobe 段落コンポーザー"] },
        { name: "欧文単数行コンポーザー", aliases: ["$ID/HL Single", "Adobe Single-line Composer", "Adobe 欧文単数行コンポーザー", "Adobe 単数行コンポーザー"] }
    ];
    var names = [];
    var aliases = [];
    for (var entryIndex = 0; entryIndex < entries.length; entryIndex++) {
        names.push(entries[entryIndex].name);
        aliases.push(entries[entryIndex].aliases);
    }
    return { names: names, aliases: aliases };
}

// =========================================
// 設定値と選択肢の照合 / Matching setting values to options
// =========================================

/**
 * コンポーザー名（エイリアスのどれか）から選択位置を探す
 * @param {Array<Array<string>>} aliasesList createComposerOptions() の aliases
 * @param {string} target 探すコンポーザー名
 * @returns {number} 見つかった位置。なければ -1
 */
function findIndexByComposerAliases(aliasesList, target) {
    if (!target) return -1;
    for (var listIndex = 0; listIndex < aliasesList.length; listIndex++) {
        if (findIndexInArray(aliasesList[listIndex], target) >= 0) return listIndex;
    }
    return -1;
}

/**
 * エイリアスを順に試して段落スタイルにコンポーザーを設定する
 * @param {ParagraphStyle} targetParagraphStyle 対象の段落スタイル
 * @param {Array<string>} aliases コンポーザー名の候補
 * @returns {boolean} 設定できたら true
 */
function applyComposerAliases(targetParagraphStyle, aliases) {
    if (!aliases) return false;
    for (var aliasIndex = 0; aliasIndex < aliases.length; aliasIndex++) {
        try {
            targetParagraphStyle.composer = aliases[aliasIndex];
            return true;
        } catch (composerAliasError) { }
    }
    return false;
}

/**
 * 禁則処理セットの設定値から選択位置を探す（参照・ID・名前の順に照合）
 * @param {*} kinsokuValue 禁則処理セットの設定値
 * @param {Array} kinsokuTables 禁則処理セットの一覧
 * @param {Array<string>} kinsokuNames 禁則処理セットの表示名
 * @returns {number} 見つかった位置。なければ -1
 */
function findKinsokuIndexFromValue(kinsokuValue, kinsokuTables, kinsokuNames) {
    if (kinsokuValue === null || kinsokuValue === undefined) return -1;
    for (var refIndex = 0; refIndex < kinsokuTables.length; refIndex++) {
        if (kinsokuTables[refIndex] === kinsokuValue) return refIndex;
        try {
            if (kinsokuTables[refIndex].id !== undefined && kinsokuValue.id !== undefined && kinsokuTables[refIndex].id === kinsokuValue.id) return refIndex;
        } catch (eKinsokuIdCompare) { }
    }
    var kinsokuValueName = null;
    try { kinsokuValueName = kinsokuValue.name; } catch (eKinsokuValueName) { }
    if (typeof kinsokuValueName === "string" && kinsokuValueName.length > 0) {
        return findIndexInArray(kinsokuNames, kinsokuValueName);
    }
    return -1;
}

/**
 * 文字組みアキ量設定の表示名を求める
 * @param {*} mojikumiValue 文字組みの設定値（MojikumiTable・文字列・NothingEnum・組み込みプリセットの列挙値）
 * @returns {string} 表示名。求められなければ空文字
 */
function resolveMojikumiName(mojikumiValue) {
    if (mojikumiValue === null || mojikumiValue === undefined || mojikumiValue === NothingEnum.NOTHING) {
        return MOJIKUMI_NONE_NAME;
    }
    if (typeof mojikumiValue === "string") {
        return mojikumiValue;
    }

    /* MojikumiTable は .name を持つ。プリセットの列挙値は持たないため toString() で照合 / MojikumiTable has .name; preset enums are matched via toString() */
    try {
        if (mojikumiValue.isValid && typeof mojikumiValue.name === "string" && mojikumiValue.name.length > 0) {
            return mojikumiValue.name;
        }
    } catch (mojikumiNameError) { }

    var mojikumiKey = "";
    try { mojikumiKey = mojikumiValue.toString(); } catch (mojikumiStringError) { }
    for (var enumKey in MOJIKUMI_LABELS) {
        if (mojikumiKey.indexOf(enumKey) !== -1) {
            return MOJIKUMI_LABELS[enumKey];
        }
    }

    return "";
}

// =========================================
// 段落スタイル設定の読み取り / Paragraph style setting readers
// =========================================

/**
 * 段落スタイルの現在の組版設定を読み取る
 * @param {ParagraphStyle} paragraphStyle 対象の段落スタイル
 * @param {object} kinsokuTableData 禁則処理セットの一覧
 * @param {object} kinsokuTypeOptions 禁則調整方式の一覧
 * @param {object} mojikumiTableData 文字組み設定の一覧
 * @param {object} composerOptions コンポーザーの一覧
 * @param {object} defaultIndexes 既定の選択位置
 * @returns {object} 各設定の選択位置
 */
function readParagraphStyleTypesettingSettings(paragraphStyle, kinsokuTableData, kinsokuTypeOptions, mojikumiTableData, composerOptions, defaultIndexes) {
    var styleSettings = {
        kinsokuIndex: defaultIndexes.kinsokuIndex,
        kinsokuTypeIndex: defaultIndexes.kinsokuTypeIndex,
        mojikumiIndex: defaultIndexes.mojikumiIndex,
        composerIndex: defaultIndexes.composerIndex
    };

    try {
        var kinsokuIndex = findKinsokuIndexFromValue(paragraphStyle.kinsokuSet, kinsokuTableData.tables, kinsokuTableData.names);
        if (kinsokuIndex >= 0) styleSettings.kinsokuIndex = kinsokuIndex;
    } catch (kinsokuReadError) { }

    try {
        var kinsokuTypeIndex = findIndexInArray(kinsokuTypeOptions.values, paragraphStyle.kinsokuType);
        if (kinsokuTypeIndex >= 0) styleSettings.kinsokuTypeIndex = kinsokuTypeIndex;
    } catch (kinsokuTypeReadError) { }

    try {
        var mojikumiIndex = findIndexInArray(mojikumiTableData.names, resolveMojikumiName(paragraphStyle.mojikumi));
        if (mojikumiIndex >= 0) styleSettings.mojikumiIndex = mojikumiIndex;
    } catch (mojikumiReadError) { }

    try {
        var composerIndex = findIndexByComposerAliases(composerOptions.aliases, paragraphStyle.composer);
        if (composerIndex >= 0) styleSettings.composerIndex = composerIndex;
    } catch (composerReadError) { }

    return styleSettings;
}

// =========================================
// ダイアログ UI 生成 / Dialog UI builders
// =========================================

/**
 * マトリックス UI に 1 セル分のコントロールを追加する
 * @param {object} parent 追加先のコンテナ
 * @param {string} controlType コントロールの種類
 * @param {object} properties コントロールのプロパティ
 * @param {number} width 列幅（px）
 * @returns {object} 追加したコントロール
 */
function addMatrixCell(parent, controlType, properties, width) {
    var control;
    if (controlType === "statictext") {
        control = parent.add("statictext", undefined, properties.text);
    } else if (controlType === "dropdownlist") {
        control = parent.add("dropdownlist", undefined, properties.items);
        control.selection = properties.selection;
    }

    if (!control) {
        throw new Error("Unsupported control type: " + controlType);
    }

    control.preferredSize.width = width;
    return control;
}

/**
 * 列単位で値を反映するボタンを追加する
 * @param {object} parent 追加先のコンテナ
 * @param {number} width 列幅（px）
 * @returns {Button} 追加したボタン
 */
function addReflectButton(parent, width) {
    var cell = parent.add("group");
    cell.preferredSize.width = width;
    cell.alignChildren = "center";
    return cell.add("button", undefined, getLabel("button.reflect"));
}

/**
 * 縦方向の余白を追加する
 * @param {object} parent 追加先のコンテナ
 * @param {number} height 余白の高さ（px）
 * @returns {Group} 追加した余白のグループ
 */
function addVerticalSpacer(parent, height) {
    var spacer = parent.add("group");
    spacer.preferredSize.height = height;
    return spacer;
}

/**
 * マトリックス UI の見出し行を追加する
 * @param {Panel} matrixPanel 追加先のパネル
 * @returns {void}
 */
function addHeaderRow(matrixPanel) {
    var headerRowGroup = matrixPanel.add("group");
    setupRow(headerRowGroup, "left", MATRIX_ROW_SPACING);
    addMatrixCell(headerRowGroup, "statictext", { text: getLabel("column.styleName") }, COLUMN_WIDTH_NAME);
    addMatrixCell(headerRowGroup, "statictext", { text: getLabel("column.kinsokuSet") }, COLUMN_WIDTH_DROPDOWN);
    addMatrixCell(headerRowGroup, "statictext", { text: getLabel("column.kinsokuType") }, COLUMN_WIDTH_DROPDOWN);
    addMatrixCell(headerRowGroup, "statictext", { text: getLabel("column.mojikumi") }, COLUMN_WIDTH_DROPDOWN_WIDE);
    addMatrixCell(headerRowGroup, "statictext", { text: getLabel("column.composer") }, COLUMN_WIDTH_DROPDOWN_WIDE);
}

/**
 * 一括反映のコピー元となる「すべて」行を追加する
 * @param {Panel} matrixPanel 追加先のパネル
 * @param {Array<string>} kinsokuNames 禁則処理セットの表示名
 * @param {Array<string>} kinsokuTypeNames 禁則調整方式の表示名
 * @param {Array<string>} mojikumiNames 文字組み設定の表示名
 * @param {Array<string>} composerNames コンポーザーの表示名
 * @param {object} defaultIndexes 既定の選択位置
 * @returns {object} コピー元のコントロール
 */
function addBulkSourceRow(matrixPanel, kinsokuNames, kinsokuTypeNames, mojikumiNames, composerNames, defaultIndexes) {
    var bulkSourceRowGroup = matrixPanel.add("group");
    setupRow(bulkSourceRowGroup, "left", MATRIX_ROW_SPACING);
    addMatrixCell(bulkSourceRowGroup, "statictext", { text: getLabel("column.bulkSource") }, COLUMN_WIDTH_NAME);

    return {
        kinsoku: addMatrixCell(bulkSourceRowGroup, "dropdownlist", { items: kinsokuNames, selection: defaultIndexes.kinsokuIndex }, COLUMN_WIDTH_DROPDOWN),
        kinsokuType: addMatrixCell(bulkSourceRowGroup, "dropdownlist", { items: kinsokuTypeNames, selection: defaultIndexes.kinsokuTypeIndex }, COLUMN_WIDTH_DROPDOWN),
        mojikumi: addMatrixCell(bulkSourceRowGroup, "dropdownlist", { items: mojikumiNames, selection: defaultIndexes.mojikumiIndex }, COLUMN_WIDTH_DROPDOWN_WIDE),
        composer: addMatrixCell(bulkSourceRowGroup, "dropdownlist", { items: composerNames, selection: defaultIndexes.composerIndex }, COLUMN_WIDTH_DROPDOWN_WIDE)
    };
}

/**
 * 列ごとの反映ボタンを並べた行を追加する
 * @param {Panel} matrixPanel 追加先のパネル
 * @returns {object} 反映ボタン
 */
function addReflectButtonRow(matrixPanel) {
    var reflectRowGroup = matrixPanel.add("group");
    setupRow(reflectRowGroup, "left", MATRIX_ROW_SPACING);
    addMatrixCell(reflectRowGroup, "statictext", { text: "" }, COLUMN_WIDTH_NAME);

    return {
        kinsoku: addReflectButton(reflectRowGroup, COLUMN_WIDTH_DROPDOWN),
        kinsokuType: addReflectButton(reflectRowGroup, COLUMN_WIDTH_DROPDOWN),
        mojikumi: addReflectButton(reflectRowGroup, COLUMN_WIDTH_DROPDOWN_WIDE),
        composer: addReflectButton(reflectRowGroup, COLUMN_WIDTH_DROPDOWN_WIDE)
    };
}

/**
 * 段落スタイルごとの設定行を追加する
 * @param {Panel} matrixPanel 追加先のパネル
 * @param {Array<string>} kinsokuNames 禁則処理セットの表示名
 * @param {Array<string>} kinsokuTypeNames 禁則調整方式の表示名
 * @param {Array<string>} mojikumiNames 文字組み設定の表示名
 * @param {Array<string>} composerNames コンポーザーの表示名
 * @param {Array<string>} paragraphStyleDisplayNames 段落スタイルの表示名
 * @param {Array<object>} initialSettingsByStyle 各スタイルの初期設定
 * @returns {Array<object>} 行ごとのコントロール
 */
function addParagraphStyleSettingRows(matrixPanel, kinsokuNames, kinsokuTypeNames, mojikumiNames, composerNames, paragraphStyleDisplayNames, initialSettingsByStyle) {
    var styleSettingRows = [];

    for (var paragraphStyleIndex = 0; paragraphStyleIndex < paragraphStyleDisplayNames.length; paragraphStyleIndex++) {
        var initialStyleSettings = initialSettingsByStyle[paragraphStyleIndex];
        var rowGroup = matrixPanel.add("group");
        setupRow(rowGroup, "left", MATRIX_ROW_SPACING);

        addMatrixCell(rowGroup, "statictext", { text: paragraphStyleDisplayNames[paragraphStyleIndex] }, COLUMN_WIDTH_NAME);
        styleSettingRows.push({
            kinsoku: addMatrixCell(rowGroup, "dropdownlist", { items: kinsokuNames, selection: initialStyleSettings.kinsokuIndex }, COLUMN_WIDTH_DROPDOWN),
            kinsokuType: addMatrixCell(rowGroup, "dropdownlist", { items: kinsokuTypeNames, selection: initialStyleSettings.kinsokuTypeIndex }, COLUMN_WIDTH_DROPDOWN),
            mojikumi: addMatrixCell(rowGroup, "dropdownlist", { items: mojikumiNames, selection: initialStyleSettings.mojikumiIndex }, COLUMN_WIDTH_DROPDOWN_WIDE),
            composer: addMatrixCell(rowGroup, "dropdownlist", { items: composerNames, selection: initialStyleSettings.composerIndex }, COLUMN_WIDTH_DROPDOWN_WIDE)
        });
    }

    return styleSettingRows;
}

/**
 * 反映ボタンにコピー処理を結び付ける
 * @param {object} reflectButtons 反映ボタン
 * @param {object} bulkSourceControls コピー元のコントロール
 * @param {Array<object>} styleSettingRows 行ごとのコントロール
 * @returns {void}
 */
function bindBulkCopyButtons(reflectButtons, bulkSourceControls, styleSettingRows) {
    reflectButtons.kinsoku.onClick = createDropdownBulkCopyHandler(bulkSourceControls.kinsoku, styleSettingRows, "kinsoku");
    reflectButtons.kinsokuType.onClick = createDropdownBulkCopyHandler(bulkSourceControls.kinsokuType, styleSettingRows, "kinsokuType");
    reflectButtons.mojikumi.onClick = createDropdownBulkCopyHandler(bulkSourceControls.mojikumi, styleSettingRows, "mojikumi");
    reflectButtons.composer.onClick = createDropdownBulkCopyHandler(bulkSourceControls.composer, styleSettingRows, "composer");
}

/**
 * 各行の選択内容を読み取る
 * @param {Array<object>} styleSettingRows 行ごとのコントロール
 * @returns {Array<object>} 段落スタイルごとの設定
 */
function readStyleSettingRows(styleSettingRows) {
    var settingsByStyle = [];

    for (var rowIndex = 0; rowIndex < styleSettingRows.length; rowIndex++) {
        settingsByStyle.push({
            kinsokuIndex: styleSettingRows[rowIndex].kinsoku.selection.index,
            kinsokuTypeIndex: styleSettingRows[rowIndex].kinsokuType.selection.index,
            mojikumiIndex: styleSettingRows[rowIndex].mojikumi.selection.index,
            composerIndex: styleSettingRows[rowIndex].composer.selection.index
        });
    }

    return settingsByStyle;
}

// =========================================
// UI 操作の反映 / UI value propagation
// =========================================

/**
 * 1 列分の値を全行へコピーするハンドラを作る
 * @param {DropDownList} sourceControl コピー元のドロップダウン
 * @param {Array<object>} styleSettingRows 行ごとのコントロール
 * @param {string} controlKey 対象の列を表すキー
 * @returns {function} クリックハンドラ
 */
function createDropdownBulkCopyHandler(sourceControl, styleSettingRows, controlKey) {
    return function () {
        if (!sourceControl.selection) return;
        var selectedIndex = sourceControl.selection.index;
        for (var rowIndex = 0; rowIndex < styleSettingRows.length; rowIndex++) {
            styleSettingRows[rowIndex][controlKey].selection = selectedIndex;
        }
    };
}

// =========================================
// ダイアログ制御 / Dialog controller
// =========================================

/**
 * 組版設定のマトリックスダイアログを表示する
 * @param {Array<string>} kinsokuNames 禁則処理セットの表示名
 * @param {Array<string>} kinsokuTypeNames 禁則調整方式の表示名
 * @param {Array<string>} mojikumiNames 文字組み設定の表示名
 * @param {Array<string>} composerNames コンポーザーの表示名
 * @param {Array<string>} paragraphStyleDisplayNames 段落スタイルの表示名
 * @param {Array<object>} initialSettingsByStyle 各スタイルの初期設定
 * @param {object} defaultIndexes 既定の選択位置
 * @returns {object|null} 設定内容。キャンセル時は null
 */
function showTypesettingSettingsDialog(kinsokuNames, kinsokuTypeNames, mojikumiNames, composerNames, paragraphStyleDisplayNames, initialSettingsByStyle, defaultIndexes) {
    var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(dialog, 10);

    var matrixPanel = dialog.add("panel", undefined, getLabel("panel.typesetting"));
    setupPanel(matrixPanel, 2);
    addHeaderRow(matrixPanel);
    addVerticalSpacer(matrixPanel, 6);

    var bulkSourceControls = addBulkSourceRow(matrixPanel, kinsokuNames, kinsokuTypeNames, mojikumiNames, composerNames, defaultIndexes);
    addVerticalSpacer(matrixPanel, 3);

    var reflectButtons = addReflectButtonRow(matrixPanel);
    addVerticalSpacer(matrixPanel, 8);

    var styleSettingRows = addParagraphStyleSettingRows(
        matrixPanel,
        kinsokuNames,
        kinsokuTypeNames,
        mojikumiNames,
        composerNames,
        paragraphStyleDisplayNames,
        initialSettingsByStyle
    );
    bindBulkCopyButtons(reflectButtons, bulkSourceControls, styleSettingRows);
    addVerticalSpacer(dialog, 10);

    var buttonRow = addButtonRow(dialog);
    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    alignRightOnlyButtonRow(buttonRow);

    if (dialog.show() !== 1) return null;

    return { settingsByStyle: readStyleSettingRows(styleSettingRows) };
}

// =========================================
// 設定の適用 / Apply settings
// =========================================

/**
 * 読み取った設定を各段落スタイルへ適用する
 * @param {Array<ParagraphStyle>} targetParagraphStyles 対象の段落スタイル
 * @param {Array<object>} settingsByStyle 段落スタイルごとの設定
 * @param {object} lookupTables 禁則・文字組み・コンポーザーの参照表
 * @returns {void}
 */
function applyTypesettingSettingsToParagraphStyles(targetParagraphStyles, settingsByStyle, lookupTables) {
    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(
        function () {
            var skipped = 0;

            for (var styleIndex = 0; styleIndex < targetParagraphStyles.length; styleIndex++) {
                var paragraphStyle = targetParagraphStyles[styleIndex];
                var styleSettings = settingsByStyle[styleIndex];
                try {
                    paragraphStyle.kinsokuSet = lookupTables.kinsokuTables[styleSettings.kinsokuIndex];
                    paragraphStyle.kinsokuType = lookupTables.kinsokuTypeValues[styleSettings.kinsokuTypeIndex];
                    var mojikumiTable = lookupTables.mojikumiTables[styleSettings.mojikumiIndex];
                    paragraphStyle.mojikumi = mojikumiTable === null ? NothingEnum.NOTHING : mojikumiTable;
                    /* コンポーザーはロケールやバージョンで受理名が違うため候補を順に試し、どれも通らなければ失敗として数える / Try each composer alias; count the style as failed if none is accepted */
                    if (!applyComposerAliases(paragraphStyle, lookupTables.composerAliases[styleSettings.composerIndex])) throw new Error("composer");
                } catch (e) {
                    skipped++;
                    continue;
                }
            }

            if (skipped > 0) {
                alert(getLabel("alert.partialFailurePrefix") + skipped + getLabel("alert.partialFailureSuffix"));
            }
        },
        ScriptLanguage.JAVASCRIPT,
        undefined,
        UndoModes.ENTIRE_SCRIPT,
        getLabel("undo.applyTypesetting")
    );
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var activeDocument = app.activeDocument;

    var kinsokuTableData = collectKinsokuTables(activeDocument);
    if (kinsokuTableData.tables.length === 0) {
        alert(getLabel("alert.noKinsokuTables"));
        return;
    }

    var kinsokuTypeOptions = createKinsokuTypeOptions();
    var mojikumiTableData = collectMojikumiTables(activeDocument);
    var targetParagraphStyleData = collectTargetParagraphStyles(activeDocument);

    if (targetParagraphStyleData.styles.length === 0) {
        alert(getLabel("alert.noParagraphStyles"));
        return;
    }

    var composerOptions = createComposerOptions();
    var defaultIndexes = {
        kinsokuIndex: getDefaultIndexByName(kinsokuTableData.names, DEFAULT_KINSOKU_SET_NAME),
        kinsokuTypeIndex: getDefaultIndexByName(kinsokuTypeOptions.names, DEFAULT_KINSOKU_TYPE_NAME),
        mojikumiIndex: getDefaultIndexByName(mojikumiTableData.names, DEFAULT_MOJIKUMI_NAME),
        composerIndex: getDefaultIndexByName(composerOptions.names, DEFAULT_COMPOSER_NAME)
    };

    // 各段落スタイルの現在値を読み取り / Read current settings from each paragraph style
    var initialSettingsByStyle = [];
    for (var styleIndex = 0; styleIndex < targetParagraphStyleData.styles.length; styleIndex++) {
        initialSettingsByStyle.push(readParagraphStyleTypesettingSettings(
            targetParagraphStyleData.styles[styleIndex], kinsokuTableData, kinsokuTypeOptions, mojikumiTableData, composerOptions, defaultIndexes
        ));
    }

    var result = showTypesettingSettingsDialog(
        kinsokuTableData.names,
        kinsokuTypeOptions.names,
        mojikumiTableData.names,
        composerOptions.names,
        targetParagraphStyleData.names,
        initialSettingsByStyle,
        defaultIndexes
    );
    if (result === null) return;

    applyTypesettingSettingsToParagraphStyles(targetParagraphStyleData.styles, result.settingsByStyle, {
        kinsokuTables: kinsokuTableData.tables,
        kinsokuTypeValues: kinsokuTypeOptions.values,
        mojikumiTables: mojikumiTableData.tables,
        composerAliases: composerOptions.aliases
    });

})();