#target indesign

/*
概要

JapaneseParagraphTypesettingManager.jsx
--------------------------------------------------------
アクティブドキュメント内の段落スタイルを一覧化し、
日本語組版に関する主要設定をマトリックス UI で確認・編集・一括適用する。

対象設定:
・禁則処理セット
・禁則調整方式
・文字組みアキ量設定
・コンポーザー

主な機能:
・段落スタイルごとに現在の組版設定を 1 行で表示
・先頭の「すべて」行を一括反映用のコピー元として使用
・列ごとの「↓ 反映」ボタンで、コピー元の値を該当列だけ全段落スタイルへ反映
・既存の段落スタイル設定を読み取り、ダイアログ初期値に反映
・カスタム文字組みアキ量設定と組み込みプリセット名の取得に対応
・段落スタイルグループを再帰的に走査し、対象外グループを除外

デフォルト値:
・スクリプト冒頭の DEFAULT_* 変数で変更可能
・指定した初期値が見つからない場合や範囲外の場合は先頭項目に戻す

除外条件:
・[段落スタイルなし] / [No Paragraph Style]
・[基本段落] / [Basic Paragraph]
・名前が "_" で始まるスタイルグループ配下の段落スタイル
*/

var SCRIPT_VERSION = "v1.1.0";

// =========================================
// 設定 / Settings
// =========================================

var DIALOG_TITLE = "日本語文字組版設定";

var DEFAULT_KINSOKU_SET_NAME = "弱い禁則";
var DEFAULT_KINSOKU_TYPE_NAME = "調整量を優先";
var DEFAULT_MOJIKUMI_NAME = "行末約物半角";
var DEFAULT_COMPOSER_NAME = "日本語単数行コンポーザー";

// マトリックス UI の列幅 / Matrix UI column widths
var W_NAME = 160;
var W_DROP = 120;       // 禁則処理セット・禁則調整方式
var W_DROP_WIDE = 200;  // 文字組み・コンポーザー

/* 組み込みプリセット enum → 表示名 / Built-in mojikumi preset enum to label */
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
// UI 共通設定 / Shared UI setup
// =========================================

/* パネルの基本レイアウトを設定 / Set shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = "left";
    panel.alignment = "fill";
    panel.margins = [15, 12, 15, 6];
    if (typeof spacing === "number") {
        panel.spacing = spacing;
    }
}

// =========================================
// ドキュメント情報の取得 / Document data collection
// =========================================

/* ドキュメント内の禁則処理セットを取得 / Collect kinsoku tables from the document */
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

/* 禁則調整方式の候補を定義 / Define kinsoku type options */
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

/* ドキュメント内の文字組み設定を取得 / Collect mojikumi tables from the document */
function collectMojikumiTables(documentObject) {
    var tables = [null];
    var names = ["なし"];
    for (var mojikumiTableIndex = 0; mojikumiTableIndex < documentObject.mojikumiTables.length; mojikumiTableIndex++) {
        var mojikumiTable = documentObject.mojikumiTables.item(mojikumiTableIndex);
        tables.push(mojikumiTable);
        names.push(mojikumiTable.name);
    }
    return { tables: tables, names: names };
}

/* 対象にする段落スタイルを収集 / Collect applicable paragraph styles */
function collectTargetParagraphStyles(documentObject) {
    var styles = [];
    var names = [];

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
            if (styleGroup.name.charAt(0) === "_") continue;
            walk(styleGroup, prefix + styleGroup.name + " / ");
        }
    }

    walk(documentObject, "");
    return { styles: styles, names: names };
}

/* コンポーザーの候補を定義 / Define composer options */
function createComposerOptions() {
    var values = [
        "Adobe 日本語段落コンポーザー",
        "Adobe 日本語単数行コンポーザー",
        "Adobe World-Ready 段落コンポーザー",
        "Adobe World-Ready 単数行コンポーザー",
        "Adobe 段落コンポーザー",
        "Adobe 単数行コンポーザー"
    ];
    var names = [];
    for (var composerIndex = 0; composerIndex < values.length; composerIndex++) {
        names.push(values[composerIndex].replace(/^Adobe\s+/, ""));
    }
    return { names: names, values: values };
}

// =========================================
// デフォルト値の解決 / Default value resolution
// =========================================

/* 名前に一致する項目のインデックスを探す（null/文字列対応） / Find an item index by name (handles null/string safely) */
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

/* 名前配列から既定値のインデックスを探す / Resolve a default index from names */
function getDefaultIndexByName(names, defaultName) {
    for (var nameIndex = 0; nameIndex < names.length; nameIndex++) {
        if (names[nameIndex] === defaultName) return nameIndex;
    }
    return 0;
}

/* 名前または値から既定値のインデックスを探す / Resolve a default index from names or values */
function getDefaultIndexByNameOrValue(names, values, defaultName) {
    for (var itemIndex = 0; itemIndex < names.length; itemIndex++) {
        if (names[itemIndex] === defaultName) return itemIndex;
        if (values && itemIndex < values.length && values[itemIndex] === defaultName) return itemIndex;
    }
    return 0;
}


// =========================================
// 段落スタイル設定の読み取り / Paragraph style setting readers
// =========================================

/* 段落スタイルから現在の組版設定を読み取る / Read current composition settings from a paragraph style */
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
        for (var composerIndex = 0; composerIndex < composerOptions.values.length; composerIndex++) {
            if (composerOptions.values[composerIndex] === currentComposer || composerOptions.names[composerIndex] === currentComposer) {
                styleSettings.composerIndex = composerIndex;
                break;
            }
        }
    } catch (e) { }

    return styleSettings;
}

// =========================================
// ダイアログ UI 生成 / Dialog UI builders
// =========================================

/* マトリックス UI のセルを追加 / Add a matrix UI cell */
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

/* 列ごとの反映ボタンを追加 / Add a column reflect button */
function addReflectButton(parent, width) {
    var cell = parent.add("group");
    cell.preferredSize.width = width;
    cell.alignChildren = "center";
    return cell.add("button", undefined, "↓ 反映");
}

/* 縦方向のスペーサーを追加 / Add a vertical spacer */
function addVerticalSpacer(parent, height) {
    var spacer = parent.add("group");
    spacer.preferredSize.height = height;
    return spacer;
}

/* ヘッダー行を追加 / Add the header row */
function addHeaderRow(matrixPanel) {
    var headerRowGroup = matrixPanel.add("group");
    headerRowGroup.orientation = "row";
    headerRowGroup.spacing = 8;
    addMatrixCell(headerRowGroup, "statictext", { text: "段落スタイル" }, W_NAME);
    addMatrixCell(headerRowGroup, "statictext", { text: "禁則処理セット" }, W_DROP);
    addMatrixCell(headerRowGroup, "statictext", { text: "禁則調整方式" }, W_DROP);
    addMatrixCell(headerRowGroup, "statictext", { text: "文字組み" }, W_DROP_WIDE);
    addMatrixCell(headerRowGroup, "statictext", { text: "コンポーザー" }, W_DROP_WIDE);
}

/* 一括反映用の「すべて」行を追加 / Add the bulk source row for bulk reflection */
function addBulkSourceRow(matrixPanel, kinsokuNames, kinsokuTypeNames, mojikumiNames, composerNames, defaultIndexes) {
    var bulkSourceRowGroup = matrixPanel.add("group");
    bulkSourceRowGroup.orientation = "row";
    bulkSourceRowGroup.spacing = 8;
    addMatrixCell(bulkSourceRowGroup, "statictext", { text: "すべて" }, W_NAME);

    return {
        kinsoku: addMatrixCell(bulkSourceRowGroup, "dropdownlist", { items: kinsokuNames, selection: defaultIndexes.kinsokuIndex }, W_DROP),
        kinsokuType: addMatrixCell(bulkSourceRowGroup, "dropdownlist", { items: kinsokuTypeNames, selection: defaultIndexes.kinsokuTypeIndex }, W_DROP),
        mojikumi: addMatrixCell(bulkSourceRowGroup, "dropdownlist", { items: mojikumiNames, selection: defaultIndexes.mojikumiIndex }, W_DROP_WIDE),
        composer: addMatrixCell(bulkSourceRowGroup, "dropdownlist", { items: composerNames, selection: defaultIndexes.composerIndex }, W_DROP_WIDE)
    };
}

/* 反映ボタン行を追加 / Add the reflect button row */
function addReflectButtonRow(matrixPanel) {
    var reflectRowGroup = matrixPanel.add("group");
    reflectRowGroup.orientation = "row";
    reflectRowGroup.spacing = 8;
    addMatrixCell(reflectRowGroup, "statictext", { text: "" }, W_NAME);

    return {
        kinsoku: addReflectButton(reflectRowGroup, W_DROP),
        kinsokuType: addReflectButton(reflectRowGroup, W_DROP),
        mojikumi: addReflectButton(reflectRowGroup, W_DROP_WIDE),
        composer: addReflectButton(reflectRowGroup, W_DROP_WIDE)
    };
}

/* 段落スタイルごとの設定行を追加 / Add setting rows for paragraph styles */
function addParagraphStyleSettingRows(matrixPanel, kinsokuNames, kinsokuTypeNames, mojikumiNames, composerNames, paragraphStyleDisplayNames, initialSettingsByStyle) {
    var styleSettingRows = [];

    for (var paragraphStyleIndex = 0; paragraphStyleIndex < paragraphStyleDisplayNames.length; paragraphStyleIndex++) {
        var initialStyleSettings = initialSettingsByStyle[paragraphStyleIndex];
        var rowGroup = matrixPanel.add("group");
        rowGroup.orientation = "row";
        rowGroup.spacing = 8;

        addMatrixCell(rowGroup, "statictext", { text: paragraphStyleDisplayNames[paragraphStyleIndex] }, W_NAME);
        styleSettingRows.push({
            kinsoku: addMatrixCell(rowGroup, "dropdownlist", { items: kinsokuNames, selection: initialStyleSettings.kinsokuIndex }, W_DROP),
            kinsokuType: addMatrixCell(rowGroup, "dropdownlist", { items: kinsokuTypeNames, selection: initialStyleSettings.kinsokuTypeIndex }, W_DROP),
            mojikumi: addMatrixCell(rowGroup, "dropdownlist", { items: mojikumiNames, selection: initialStyleSettings.mojikumiIndex }, W_DROP_WIDE),
            composer: addMatrixCell(rowGroup, "dropdownlist", { items: composerNames, selection: initialStyleSettings.composerIndex }, W_DROP_WIDE)
        });
    }

    return styleSettingRows;
}

/* 反映ボタンに列単位の一括コピー処理を割り当て / Bind per-column bulk-copy actions to reflect buttons */
function bindBulkCopyButtons(reflectButtons, bulkSourceControls, styleSettingRows) {
    reflectButtons.kinsoku.onClick = createDropdownBulkCopyHandler(bulkSourceControls.kinsoku, styleSettingRows, "kinsoku");
    reflectButtons.kinsokuType.onClick = createDropdownBulkCopyHandler(bulkSourceControls.kinsokuType, styleSettingRows, "kinsokuType");
    reflectButtons.mojikumi.onClick = createDropdownBulkCopyHandler(bulkSourceControls.mojikumi, styleSettingRows, "mojikumi");
    reflectButtons.composer.onClick = createDropdownBulkCopyHandler(bulkSourceControls.composer, styleSettingRows, "composer");
}

/* ダイアログ上の段落スタイル設定行から値を読み取る / Read setting values from paragraph style rows */
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

/* コピー元ドロップダウンの値を全スタイル行へ反映 / Copy a source dropdown value to all style setting rows */
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

/* 組版設定ダイアログを表示して結果を返す / Show the typesetting settings dialog and return the result */
function showTypesettingSettingsDialog(kinsokuNames, kinsokuTypeNames, mojikumiNames, composerNames, paragraphStyleDisplayNames, initialSettingsByStyle, defaultIndexes) {
    var dialog = new Window("dialog", DIALOG_TITLE + " " + SCRIPT_VERSION);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.margins = [15, 12, 15, 16];
    dialog.spacing = 10;

    var matrixPanel = dialog.add("panel", undefined, "組版設定");
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

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";
    buttonGroup.add("button", undefined, "キャンセル", { name: "cancel" });
    buttonGroup.add("button", undefined, "OK", { name: "ok" });

    if (dialog.show() !== 1) return null;

    return { settingsByStyle: readStyleSettingRows(styleSettingRows) };
}

// =========================================
// 設定の適用 / Apply settings
// =========================================

/* ダイアログの設定を段落スタイルに適用 / Apply dialog settings to paragraph styles */
function applyTypesettingSettingsToParagraphStyles(targetParagraphStyles, settingsByStyle, lookupTables) {
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
                    paragraphStyle.composer = lookupTables.composerValues[styleSettings.composerIndex];
                } catch (e) {
                    skipped++;
                    continue;
                }
            }

            if (skipped > 0) {
                alert("適用しましたが、" + skipped + " 件の段落スタイルでエラーが発生しました。");
            }
        },
        ScriptLanguage.JAVASCRIPT,
        undefined,
        UndoModes.ENTIRE_SCRIPT,
        "日本語文字組版設定を段落スタイルに適用"
    );
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {
    if (app.documents.length === 0) {
        alert("ドキュメントを開いてから実行してください。");
        return;
    }

    var activeDocument = app.activeDocument;

    var kinsokuTableData = collectKinsokuTables(activeDocument);
    if (kinsokuTableData.tables.length === 0) {
        alert("このドキュメントには禁則処理セットがありません。");
        return;
    }

    var kinsokuTypeOptions = createKinsokuTypeOptions();
    var mojikumiTableData = collectMojikumiTables(activeDocument);
    var targetParagraphStyleData = collectTargetParagraphStyles(activeDocument);

    if (targetParagraphStyleData.styles.length === 0) {
        alert("適用可能な段落スタイルがありません。");
        return;
    }

    var composerOptions = createComposerOptions();
    var defaultIndexes = {
        kinsokuIndex: getDefaultIndexByName(kinsokuTableData.names, DEFAULT_KINSOKU_SET_NAME),
        kinsokuTypeIndex: getDefaultIndexByName(kinsokuTypeOptions.names, DEFAULT_KINSOKU_TYPE_NAME),
        mojikumiIndex: getDefaultIndexByName(mojikumiTableData.names, DEFAULT_MOJIKUMI_NAME),
        composerIndex: getDefaultIndexByNameOrValue(composerOptions.names, composerOptions.values, DEFAULT_COMPOSER_NAME)
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
        composerValues: composerOptions.values
    });

})();