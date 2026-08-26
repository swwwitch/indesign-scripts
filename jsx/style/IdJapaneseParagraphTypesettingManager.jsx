#target indesign

/*

### 概要

段落スタイルの日本語組版設定（禁則処理セット・禁則調整方式・文字組みアキ量・コンポーザー）をマトリックス UI で確認・一括適用します。

詳細は README を参照してください。

### Overview

Reviews and batch-applies the Japanese composition settings of paragraph styles (kinsoku set, kinsoku adjustment, mojikumi and composer) through a matrix UI.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdJapaneseParagraphTypesettingManager"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                               /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";          /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-05";                           /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-05-06";                           /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdJapaneseParagraphTypesettingManager.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdJapaneseParagraphTypesettingManager.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// ==============================
// UIレイアウトの共通設定 / Shared UI layout
// ==============================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

/**
 * ウィンドウの共通設定を適用する
 * @param {Window} win 対象ウィンドウ
 * @param {number} [spacing] 要素間隔。省略時は WINDOW_SPACING
 * @returns {void}
 */
function setupWindow(win, spacing) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルの共通設定を適用する
 * @param {Panel} panel 対象パネル
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * 行グループの共通設定を適用する（ボタン列など）
 * @param {Group} group 対象グループ
 * @param {string} [alignment] 配置。省略時は "left"
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

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

/**
 * UI 言語を判定する
 * @returns {string} "ja" または "en"
 */
function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var currentLang = getCurrentLang();

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
        mojikumi:    { ja: "文字組みアキ量設定", en: "Mojikumi" },
        composer:    { ja: "コンポーザー", en: "Composer" },
        bulkSource:  { ja: "すべて", en: "All" }
    },
    value: {
        none: { ja: "なし", en: "None" }
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

/**
 * ドット区切りキーでラベルを取得する
 * @param {string} labelKey 例: "dialog.title"
 * @returns {string} 現在の言語のラベル文字列。見つからない場合はキーをそのまま返す
 */
function getLabel(labelKey) {
    var node = LABELS;
    var keyParts = labelKey.split(".");
    for (var i = 0; i < keyParts.length; i++) {
        node = node[keyParts[i]];
        if (!node) return labelKey;
    }
    return node[currentLang] || node.en || labelKey;
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
 * ドキュメント内の文字組みアキ量設定を集める
 * @param {Document} documentObject 対象ドキュメント
 * @returns {{tables: Array, names: Array<string>}} 文字組み設定と表示名
 */
function collectMojikumiTables(documentObject) {
    var tables = [null];
    var names = [getLabel("value.none")];
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
 * @returns {{names: Array<string>, table: Array<object>}} 表示名と定義
 */
function createComposerOptions() {
    var table = [
        {
            uiLabel: "Adobe 日本語段落コンポーザー",
            aliases: ["Adobe 日本語段落コンポーザー", "Adobe Japanese Paragraph Composer"]
        },
        {
            uiLabel: "Adobe 日本語単数行コンポーザー",
            aliases: ["Adobe 日本語単数行コンポーザー", "Adobe Japanese Single-line Composer"]
        },
        {
            uiLabel: "Adobe World-Ready 段落コンポーザー",
            aliases: ["$ID/HL Composer Optyca", "Adobe World-Ready Paragraph Composer", "Adobe 多言語対応段落コンポーザー", "Adobe World-Ready 段落コンポーザー"]
        },
        {
            uiLabel: "Adobe World-Ready 単数行コンポーザー",
            aliases: ["$ID/HL Single Optyca", "Adobe World-Ready Single-line Composer", "Adobe 多言語対応単数行コンポーザー", "Adobe World-Ready 単数行コンポーザー"]
        },
        {
            uiLabel: "Adobe 欧文段落コンポーザー",
            aliases: ["$ID/HL Composer", "Adobe Paragraph Composer", "Adobe 欧文段落コンポーザー", "Adobe 段落コンポーザー"]
        },
        {
            uiLabel: "Adobe 欧文単数行コンポーザー",
            aliases: ["$ID/HL Single", "Adobe Single-line Composer", "Adobe 欧文単数行コンポーザー", "Adobe 単数行コンポーザー"]
        }
    ];
    var names = [];
    for (var composerIndex = 0; composerIndex < table.length; composerIndex++) {
        names.push(table[composerIndex].uiLabel.replace(/^Adobe\s+/, ""));
    }
    return { names: names, table: table };
}

// =========================================
// デフォルト値の解決 / Default value resolution
// =========================================

/**
 * 名前が一致する項目の位置を探す
 * @param {Array} namedItems 名前を持つ項目の配列
 * @param {string} targetName 探す名前
 * @returns {number} 見つかった位置。なければ -1
 */
function findIndexByName(namedItems, targetName) {
    for (var itemIndex = 0; itemIndex < namedItems.length; itemIndex++) {
        var item = namedItems[itemIndex];

        // null（例：文字組みの「なし」）はスキップ / Skip null entries such as Mojikumi "None"
        if (item === null) continue;

        // name を持つオブジェクト（通常ケース） / Object with a name property, the normal case
        if (item && typeof item.name === "string") {
            if (item.name === targetName) return itemIndex;
        }

        // 表示名だけの文字列配列にも対応 / Also support plain string name arrays
        if (typeof item === "string") {
            if (item === targetName) return itemIndex;
        }
    }
    return -1;
}

/**
 * 既定値の名前から選択位置を求める
 * @param {Array<string>} names 表示名の一覧
 * @param {string} defaultName 既定値の名前
 * @returns {number} 選択する位置
 */
function getDefaultIndexByName(names, defaultName) {
    for (var nameIndex = 0; nameIndex < names.length; nameIndex++) {
        if (names[nameIndex] === defaultName) return nameIndex;
    }
    return 0;
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
        var currentKinsokuSet = paragraphStyle.kinsokuSet;
        var kinsokuIndex = findIndexByName(kinsokuTableData.tables, currentKinsokuSet.name);
        if (kinsokuIndex >= 0) styleSettings.kinsokuIndex = kinsokuIndex;
    } catch (e) { }

    try {
        var currentKinsokuType = paragraphStyle.kinsokuType;
        for (var kinsokuTypeIndex = 0; kinsokuTypeIndex < kinsokuTypeOptions.values.length; kinsokuTypeIndex++) {
            if (kinsokuTypeOptions.values[kinsokuTypeIndex] === currentKinsokuType) {
                styleSettings.kinsokuTypeIndex = kinsokuTypeIndex;
                break;
            }
        }
    } catch (e) { }

    try {
        var currentMojikumi = paragraphStyle.mojikumi;
        var mojikumiName = "";

        // mojikumi は MojikumiTable / String / NothingEnum / 組み込みプリセット enum のいずれか / Mojikumi can be a table, string, NothingEnum, or built-in preset enum
        if (typeof currentMojikumi === "string") {
            mojikumiName = currentMojikumi;
        } else if (currentMojikumi && currentMojikumi !== NothingEnum.NOTHING) {
            // MojikumiTable は .name を持つ。プリセット enum は持たないため toString() で照合 / MojikumiTable has .name; preset enums are matched via toString()
            try {
                if (currentMojikumi.isValid && typeof currentMojikumi.name === "string") {
                    mojikumiName = currentMojikumi.name;
                }
            } catch (eName) { }

            if (!mojikumiName) {
                var mojikumiKey = "";
                try { mojikumiKey = currentMojikumi.toString(); } catch (eStr) { }
                for (var enumKey in MOJIKUMI_LABELS) {
                    if (mojikumiKey.indexOf(enumKey) !== -1) {
                        mojikumiName = MOJIKUMI_LABELS[enumKey];
                        break;
                    }
                }
            }
        }

        if (mojikumiName) {
            var mojikumiIndex = findIndexByName(mojikumiTableData.tables, mojikumiName);
            if (mojikumiIndex >= 0) styleSettings.mojikumiIndex = mojikumiIndex;
        }
    } catch (e) { }

    try {
        var currentComposer = paragraphStyle.composer;
        var matchedComposer = false;
        for (var composerIndex = 0; composerIndex < composerOptions.table.length; composerIndex++) {
            var composerAliases = composerOptions.table[composerIndex].aliases;
            for (var aliasIndex = 0; aliasIndex < composerAliases.length; aliasIndex++) {
                if (composerAliases[aliasIndex] === currentComposer) {
                    styleSettings.composerIndex = composerIndex;
                    matchedComposer = true;
                    break;
                }
            }
            if (matchedComposer) break;
        }
    } catch (e) { }

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
    return cell.add("button", undefined, "↓ 反映");
}

/**
 * 縦方向の余白を追加する
 * @param {object} parent 追加先のコンテナ
 * @param {number} height 余白の高さ（px）
 * @returns {void}
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
    headerRowGroup.orientation = "row";
    headerRowGroup.spacing = MATRIX_ROW_SPACING;
    addMatrixCell(headerRowGroup, "statictext", { text: "段落スタイル" }, COLUMN_WIDTH_NAME);
    addMatrixCell(headerRowGroup, "statictext", { text: "禁則処理セット" }, COLUMN_WIDTH_DROPDOWN);
    addMatrixCell(headerRowGroup, "statictext", { text: "禁則調整方式" }, COLUMN_WIDTH_DROPDOWN);
    addMatrixCell(headerRowGroup, "statictext", { text: "文字組み" }, COLUMN_WIDTH_DROPDOWN_WIDE);
    addMatrixCell(headerRowGroup, "statictext", { text: "コンポーザー" }, COLUMN_WIDTH_DROPDOWN_WIDE);
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
    bulkSourceRowGroup.orientation = "row";
    bulkSourceRowGroup.spacing = MATRIX_ROW_SPACING;
    addMatrixCell(bulkSourceRowGroup, "statictext", { text: "すべて" }, COLUMN_WIDTH_NAME);

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
    reflectRowGroup.orientation = "row";
    reflectRowGroup.spacing = MATRIX_ROW_SPACING;
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
        rowGroup.orientation = "row";
        rowGroup.spacing = MATRIX_ROW_SPACING;

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

    /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
    var buttonGroup = dialog.add("group");
    setupRow(buttonGroup, "right", 8);
    buttonGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    buttonGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

    if (dialog.show() !== 1) return null;

    return { settingsByStyle: readStyleSettingRows(styleSettingRows) };
}

// =========================================
// 設定の適用 / Apply settings
// =========================================

/**
 * エイリアスを順に試して段落スタイルにコンポーザーを設定する
 * @param {ParagraphStyle} paragraphStyle 対象の段落スタイル
 * @param {Array<string>} aliases コンポーザー名の候補
 * @returns {void}
 */
function applyComposerToParagraphStyle(paragraphStyle, aliases) {
    var lastError = null;
    for (var aliasIndex = 0; aliasIndex < aliases.length; aliasIndex++) {
        try {
            paragraphStyle.composer = aliases[aliasIndex];
            return;
        } catch (e) {
            lastError = e;
        }
    }
    if (lastError) throw lastError;
}

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
                    applyComposerToParagraphStyle(paragraphStyle, lookupTables.composerTable[styleSettings.composerIndex].aliases);
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
        composerTable: composerOptions.table
    });

})();