#target indesign

/*
 * TypesettingStyleManager.jsx
 *
 * 段落スタイルの文字組版設定（禁則・文字組み・グリッド揃え・ハイフネーションなど）をダイアログでまとめて設定します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TypesettingStyleManager";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-05-07";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/TypesettingStyleManager.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/TypesettingStyleManager.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n7f67e8da571f"; /* 紹介記事 / article URL */

// Original idea
// 欧文組版でのハイフネーション設定：コンさん
// https://typesetterkon.blogspot.com/2011/06/indesign5.html

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
        title:            { ja: "文字組版設定（一括）", en: "Typesetting Settings (Batch)" },
        stylePicker:      { ja: "段落スタイルの選択", en: "Select Paragraph Styles" },
        presetNameInput:  { ja: "プリセット名の入力", en: "Enter Preset Name" }
    },
    panel: {
        targetStyles:     { ja: "対象の段落スタイル", en: "Target Paragraph Styles" },
        preset:           { ja: "プリセット", en: "Preset" },
        basicSettings:    { ja: "基本設定", en: "Basic Settings" },
        quotes:           { ja: "引用符（環境設定）", en: "Quotes (Preferences)" },
        units:            { ja: "単位（環境設定）", en: "Units (Preferences)" },
        japaneseTypeset:  { ja: "日本語文字組版", en: "Japanese Typesetting" },
        hyphenation:      { ja: "ハイフネーション", en: "Hyphenation" },
        hyphenateBreak:   { ja: "ハイフンで区切る", en: "Hyphenate" }
    },
    field: {
        autoKerning:      { ja: getLabel("field.autoKerning"), en: "Kerning: " },
        autoLeading:      { ja: getLabel("field.autoLeading"), en: "Auto Leading: " },
        characterAlign:   { ja: getLabel("field.characterAlign"), en: "Character Alignment: " },
        leadingModel:     { ja: getLabel("field.leadingModel"), en: "Leading Model: " },
        gridAlignment:    { ja: getLabel("field.gridAlignment"), en: "Align to Grid: " },
        composer:         { ja: getLabel("field.composer"), en: "Composer: " },
        doubleQuote:      { ja: getLabel("field.doubleQuote"), en: "Double Quotes: " },
        singleQuote:      { ja: getLabel("field.singleQuote"), en: "Quotes: " },
        textSize:         { ja: getLabel("field.textSize"), en: "Text Size: " },
        typography:       { ja: getLabel("field.typography"), en: "Typography: " },
        language:         { ja: getLabel("field.language"), en: "Language: " },
        kinsokuSet:       { ja: getLabel("field.kinsokuSet"), en: "Kinsoku Set: " },
        kinsokuType:      { ja: getLabel("field.kinsokuType"), en: "Kinsoku Adjustment: " },
        kinsokuHangType:  { ja: getLabel("field.kinsokuHangType"), en: "Hanging Punctuation: " },
        mojikumi:         { ja: getLabel("field.mojikumi"), en: "Mojikumi: " },
        minWordLength:    { ja: getLabel("field.minWordLength"), en: "Shortest Word: " },
        afterFirst:       { ja: getLabel("field.afterFirst"), en: "After First: " },
        beforeLast:       { ja: getLabel("field.beforeLast"), en: "Before Last: " },
        maxHyphens:       { ja: getLabel("field.maxHyphens"), en: "Hyphen Limit: " },
        hyphenationZone:  { ja: getLabel("field.hyphenationZone"), en: "Hyphenation Zone: " },
        presetName:       { ja: "プリセット名（書き出しファイル名にも使用）：", en: "Preset name (also used as the export filename): " },
        targetCount:      { ja: "対象: ", en: "Targets: " }
    },
    radio: {
        targetSelection:  { ja: "選択中", en: "Selection" },
        targetAll:        { ja: "すべて", en: "All" },
        targetSpecified:  { ja: "指定", en: "Specified" },
        languageJapanese: { ja: "日本語", en: "Japanese" },
        languageEnglish:  { ja: "英語", en: "English" },
        languageNone:     { ja: "なし", en: "None" }
    },
    checkbox: {
        typographersQuotes: { ja: "英文引用符を使用", en: "Use Typographer's Quotes" },
        noBreak:            { ja: "分離禁止処理", en: "No Break" },
        digitsRotation:     { ja: "連数字処理", en: "Tatechuyoko" },
        rotateInVertical:   { ja: "縦組み中の文字回転", en: "Rotate Characters in Vertical Text" },
        absorbTrailingSpace:{ ja: "全角スペースを行末吸収", en: "Absorb Trailing Full-width Space" },
        arbitraryHyphen:    { ja: "欧文泣き別れ", en: "Allow Arbitrary Hyphenation" },
        capitalizedWords:   { ja: "大文字の単語", en: "Capitalized Words" },
        acrossColumns:      { ja: "段間、フレームにわたる単語", en: "Words Across Columns and Frames" },
        lastWord:           { ja: "段落末尾の単語", en: "Last Word" },
        ligatures:          { ja: "欧文合字", en: "Ligatures" },
        hyphenation:        { ja: "ハイフネーション", en: "Hyphenation" }
    },
    button: {
        ok:        { ja: "OK", en: "OK" },
        cancel:    { ja: "キャンセル", en: "Cancel" },
        select:    { ja: "選択", en: "Select" },
        selectAll: { ja: "全選択", en: "Select All" },
        clearAll:  { ja: "全解除", en: "Clear All" },
        export:    { ja: "書き出し", en: "Export" }
    },
    unit: {
        item:      { ja: " 件", en: " items" },
        character: { ja: "文字", en: "characters" },
        hyphen:    { ja: "ハイフン", en: "hyphens" }
    },
    hint: {
        multiSelect: { ja: "Shift / Cmd（Ctrl）+ クリックで複数選択", en: "Shift / Cmd (Ctrl) + click to select multiple" }
    },
    export: {
        codeHeader:      { ja: "// PRESETS マップに以下を追加してください（プリセットドロップダウン項目への追加もお忘れなく）", en: "// Add the following to the PRESETS map (and to the preset dropdown items as well)" },
        overwritePrefix: { ja: "「", en: "\"" },
        overwriteSuffix: { ja: ".jsx」は既にデスクトップに存在します。上書きしますか？", en: ".jsx\" already exists on the Desktop. Overwrite it?" },
        savedPrefix:     { ja: "プリセット「", en: "Preset \"" },
        savedSuffix:     { ja: "」をデスクトップに書き出しました。", en: "\" was exported to the Desktop." }
    },
    alert: {
        openFileFailed:       { ja: "ファイルを開けませんでした。", en: "The file could not be opened." },
        exportErrorPrefix:    { ja: "書き出しエラー: ", en: "Export error: " },
        partialFailurePrefix: { ja: "適用しましたが、", en: "Applied, but " },
        partialFailureSuffix: { ja: " 件の段落スタイルでエラーが発生しました。\n\n", en: " paragraph style(s) reported an error.\n\n" },
        noDocument:           { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document before running." },
        noKinsokuTables:      { ja: "このドキュメントには禁則処理セットがありません。", en: "This document has no kinsoku tables." },
        noParagraphStyles:    { ja: "適用可能な段落スタイルがありません。", en: "There are no applicable paragraph styles." },
        noTargetStyles:       { ja: "適用対象の段落スタイルが見つかりません。選択範囲、または指定した段落スタイルを確認してください。", en: "No target paragraph style was found. Check the selection or the specified styles." }
    },
    undo: {
        applyTypesetting: { ja: getLabel("undo.applyTypesetting"), en: "Apply Typesetting Settings to Paragraph Styles" }
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
// ユーザー設定 / User settings
// =========================================

// 言語の候補（先頭から順に試行） / Language candidates (tried in order)
var LANGUAGE_CANDIDATES = {
    "ja": ["日本語", "Japanese"],
    "en": ["英語：米国", "English: USA"],
    "none": ["[言語なし]", "[No Language]"]
};

// 二重引用符の候補 / Double quote options
var DOUBLE_QUOTE_OPTIONS = [
    "“”",
    "«»",
    "„“",
    "『』",
    "「」",
    "\"\""
];

// 引用符の候補 / Single quote options
var SINGLE_QUOTE_OPTIONS = [
    "‘’",
    "‹›",
    "‚‘",
    "〈〉",
    "''"
];

// ドロップダウン幅 / Dropdown widths
var W_DROP = 130;

// パネルの共通マージン / Shared panel margins

// =========================================
// 文字組みアキ量プリセット定義 / Mojikumi preset definitions
// =========================================

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
// UI 共通設定 / Shared UI setup
// =========================================

// =========================================
// ドキュメント情報の取得 / Document data collection
// =========================================

/**
 * ドキュメント内の禁則処理セットを集める
 * @param {Document} documentObject 対象ドキュメント
 * @returns {object} 禁則処理セットと表示名
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
 * @returns {object} 調整方式と表示名
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
 * ぶら下がり方法の選択肢を作る
 * @returns {object} ぶら下がり方法と表示名
 */
function createKinsokuHangTypeOptions() {
    return {
        names: ["なし", "標準", "強制"],
        values: [
            KinsokuHangTypes.NONE,
            KinsokuHangTypes.KINSOKU_HANG_REGULAR,
            KinsokuHangTypes.KINSOKU_HANG_FORCE
        ]
    };
}

/**
 * ドキュメント内の文字組みアキ量設定を集める
 * @param {Document} documentObject 対象ドキュメント
 * @returns {object} 文字組み設定と表示名
 */
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

/**
 * 対象にする段落スタイルをグループ込みで集める
 * @param {Document} documentObject 対象ドキュメント
 * @returns {object} 段落スタイルと表示名
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
            if (styleGroup.name.charAt(0) === "_") continue;
            walk(styleGroup, prefix + styleGroup.name + " / ");
        }
    }

    walk(documentObject, "");
    return { styles: styles, names: names };
}

/**
 * 行送りの基準位置の選択肢を作る
 * @returns {object} 基準位置と表示名
 */
function createLeadingModelOptions() {
    return {
        names: ["仮想ボディの上/右", "仮想ボディの中央", "欧文ベースライン", "仮想ボディの下/左"],
        values: [
            LeadingModel.LEADING_MODEL_AKI_BELOW,
            LeadingModel.LEADING_MODEL_CENTER,
            LeadingModel.LEADING_MODEL_ROMAN,
            LeadingModel.LEADING_MODEL_AKI_ABOVE
        ]
    };
}

/**
 * 自動カーニング方式の選択肢を作る
 * @returns {object} カーニング方式と表示名
 */
function createKerningMethodOptions() {
    var values = ["メトリクス", "オプティカル", "和文等幅", "0"];
    return { names: values, values: values };
}

/**
 * グリッド揃えの選択肢を作る
 * @returns {object} グリッド揃えと表示名
 */
function createGridAlignmentOptions() {
    return {
        names: [
            "なし",
            "欧文ベースライン",
            "仮想ボディの上/右",
            "仮想ボディの中央",
            "仮想ボディの下/左",
            "平均字面の上/右",
            "平均字面の下/左"
        ],
        values: [
            GridAlignment.NONE,
            GridAlignment.ALIGN_BASELINE,
            GridAlignment.ALIGN_EM_TOP,
            GridAlignment.ALIGN_EM_CENTER,
            GridAlignment.ALIGN_EM_BOTTOM,
            GridAlignment.ALIGN_ICF_TOP,
            GridAlignment.ALIGN_ICF_BOTTOM
        ]
    };
}

/**
 * 文字揃えの選択肢を作る
 * @returns {object} 文字揃えと表示名
 */
function createCharacterAlignmentOptions() {
    return {
        names: [
            "欧文ベースライン",
            "仮想ボディの上/右",
            "仮想ボディの中央",
            "仮想ボディの下/左",
            "平均字面の上/右",
            "平均字面の下/左"
        ],
        values: [
            CharacterAlignment.ALIGN_BASELINE,
            CharacterAlignment.ALIGN_EM_TOP,
            CharacterAlignment.ALIGN_EM_CENTER,
            CharacterAlignment.ALIGN_EM_BOTTOM,
            CharacterAlignment.ALIGN_ICF_TOP,
            CharacterAlignment.ALIGN_ICF_BOTTOM
        ]
    };
}

/**
 * コンポーザーの選択肢と適用用エイリアスを作る
 * @returns {object} 表示名と定義
 */
function createComposerOptions() {
    var entries = [
        { name: "多言語対応単数行コンポーザー", aliases: ["$ID/HL Single Optyca", "Adobe World-Ready Single-line Composer", "Adobe 多言語対応単数行コンポーザー", "Adobe World-Ready 単数行コンポーザー"] },
        { name: "多言語対応段落コンポーザー", aliases: ["$ID/HL Composer Optyca", "Adobe World-Ready Paragraph Composer", "Adobe 多言語対応段落コンポーザー", "Adobe World-Ready 段落コンポーザー"] },
        { name: "日本語単数行コンポーザー", aliases: ["Adobe 日本語単数行コンポーザー", "Adobe Japanese Single-line Composer"] },
        { name: "日本語段落コンポーザー", aliases: ["Adobe 日本語段落コンポーザー", "Adobe Japanese Paragraph Composer"] },
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

/**
 * コンポーザー名から選択位置を探す
 * @param {Array<object>} composerTable コンポーザーの定義
 * @param {string} composerName 探すコンポーザー名
 * @returns {number} 見つかった位置。なければ -1
 */
function findIndexByComposerAliases(aliasesList, target) {
    if (!target) return -1;
    for (var listIndex = 0; listIndex < aliasesList.length; listIndex++) {
        var aliases = aliasesList[listIndex];
        for (var aliasIndex = 0; aliasIndex < aliases.length; aliasIndex++) {
            if (aliases[aliasIndex] === target) return listIndex;
        }
    }
    return -1;
}

/**
 * エイリアスを順に試して段落スタイルにコンポーザーを設定する
 * @param {ParagraphStyle} paragraphStyle 対象の段落スタイル
 * @param {Array<string>} aliases コンポーザー名の候補
 * @returns {void}
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

// =========================================
// デフォルト値の解決 / Default value resolution
// =========================================

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

/**
 * 名前または値から既定の選択位置を求める
 * @param {object} options 選択肢の定義
 * @param {*} defaultNameOrValue 既定の名前または値
 * @returns {number} 選択する位置
 */
function getDefaultIndexByNameOrValue(names, values, defaultName) {
    for (var itemIndex = 0; itemIndex < names.length; itemIndex++) {
        if (names[itemIndex] === defaultName) return itemIndex;
        if (values && itemIndex < values.length && values[itemIndex] === defaultName) return itemIndex;
    }
    return 0;
}

// =========================================
// ダイアログ UI / Dialog UI
// =========================================

/**
 * ラベル付きドロップダウンの行を追加する
 * @param {object} parent 追加先のコンテナ
 * @param {string} labelText ラベルの文字列
 * @param {Array<string>} items 選択肢
 * @param {number} defaultIndex 既定の選択位置
 * @returns {DropDownList} 追加したドロップダウン
 */
function addDropdownRow(parent, labelText, items, selectionIndex) {
    var row = parent.add("group");
    row.orientation = "row";
    row.alignChildren = ["left", "center"];
    row.spacing = 8;

    var label = row.add("statictext", undefined, labelText);
    label.preferredSize.width = 120;

    var dropdown = row.add("dropdownlist", undefined, items);
    dropdown.selection = selectionIndex;
    dropdown.preferredSize.width = W_DROP;

    return dropdown;
}

/**
 * ラベル付き数値入力の行を追加する
 * @param {object} parent 追加先のコンテナ
 * @param {string} labelText ラベルの文字列
 * @param {*} defaultValue 初期値
 * @param {string} unitText 単位の文字列
 * @returns {EditText} 追加した入力欄
 */
function addNumberRow(parent, labelText, defaultValue, suffixText) {
    var row = parent.add("group");
    row.orientation = "row";
    row.alignChildren = ["left", "center"];
    row.spacing = 8;

    var label = row.add("statictext", undefined, labelText);
    label.preferredSize.width = 120;

    var input = row.add("edittext", undefined, String(defaultValue));
    input.preferredSize.width = 50;
    input.justify = "right";

    if (typeof suffixText === "string" && suffixText.length > 0) {
        row.add("statictext", undefined, suffixText);
    }

    return input;
}

/**
 * 段落スタイルを複数選択するダイアログを表示する
 * @param {Array<string>} paragraphStyleNames 段落スタイルの表示名
 * @param {Array<string>} selectedNames すでに選択済みの名前
 * @returns {Array<string>|null} 選択した名前。キャンセル時は null
 */
function showParagraphStylePicker(paragraphStyleNames, currentSelectedIndexes) {
    var picker = new Window("dialog", getLabel("dialog.stylePicker"));
    setupWindow(picker, 10);

    var targetCountText = picker.add("statictext", undefined, getLabel("field.targetCount") + paragraphStyleNames.length + getLabel("unit.item"));
    targetCountText.alignment = "left";

    // listbox はスタイル数が多くても自動でスクロールバーが付く /
    // listbox shows a scrollbar automatically when items overflow
    var styleListbox = picker.add("listbox", undefined, paragraphStyleNames, { multiselect: true });
    styleListbox.preferredSize = [360, 320];

    /**
     * その段落スタイルが選択済みかを判定する
     * @param {string} styleName 段落スタイル名
     * @returns {boolean} 選択済みなら true
     */
    function isCurrentlySelected(index) {
        if (!currentSelectedIndexes) return true;
        for (var selectedIndexPosition = 0; selectedIndexPosition < currentSelectedIndexes.length; selectedIndexPosition++) {
            if (currentSelectedIndexes[selectedIndexPosition] === index) return true;
        }
        return false;
    }

    var initialSelectionIndexes = [];
    for (var nameIndex = 0; nameIndex < paragraphStyleNames.length; nameIndex++) {
        if (isCurrentlySelected(nameIndex)) initialSelectionIndexes.push(nameIndex);
    }
    if (initialSelectionIndexes.length > 0) styleListbox.selection = initialSelectionIndexes;

    var toolRow = picker.add("group");
    toolRow.alignment = "left";
    toolRow.spacing = 6;
    var selectAllButton = toolRow.add("button", undefined, getLabel("button.selectAll"));
    var clearAllButton = toolRow.add("button", undefined, getLabel("button.clearAll"));

    selectAllButton.onClick = function () {
        var allIndexes = [];
        for (var allIndex = 0; allIndex < styleListbox.items.length; allIndex++) allIndexes.push(allIndex);
        styleListbox.selection = allIndexes;
    };
    clearAllButton.onClick = function () {
        styleListbox.selection = null;
    };

    var hint = picker.add("statictext", undefined, getLabel("hint.multiSelect"));
    hint.alignment = "left";

    var buttonGroup = picker.add("group");
    buttonGroup.alignment = "right";
    buttonGroup.add("button", undefined, "キャンセル", { name: "cancel" });
    buttonGroup.add("button", undefined, "OK", { name: "ok" });

    if (picker.show() !== 1) return null;

    var selectedIndexes = [];
    if (styleListbox.selection) {
        for (var selectionItemIndex = 0; selectionItemIndex < styleListbox.selection.length; selectionItemIndex++) {
            selectedIndexes.push(styleListbox.selection[selectionItemIndex].index);
        }
    }
    return selectedIndexes;
}

/**
 * 表示名から選択位置を探す
 * @param {object} options 選択肢の定義
 * @param {string} displayName 探す表示名
 * @returns {number} 見つかった位置。なければ -1
 */
function findIndexByDisplayName(names, target) {
    for (var nameIndex = 0; nameIndex < names.length; nameIndex++) {
        if (names[nameIndex] === target) return nameIndex;
    }
    return -1;
}

/**
 * 列挙値から選択位置を探す
 * @param {object} options 選択肢の定義
 * @param {*} enumValue 探す列挙値
 * @returns {number} 見つかった位置。なければ -1
 */
function findIndexByEnumValue(values, target) {
    for (var valueIndex = 0; valueIndex < values.length; valueIndex++) {
        if (values[valueIndex] === target) return valueIndex;
    }
    return -1;
}

/**
 * 列挙値からドロップダウンを安全に選択する
 * @param {DropDownList} dropdown 対象のドロップダウン
 * @param {object} options 選択肢の定義
 * @param {*} enumValue 設定する列挙値
 * @returns {void}
 */
function safeAssignDropdownFromEnum(dropdown, values, target) {
    var foundIndex = findIndexByEnumValue(values, target);
    if (foundIndex < 0) return false;
    dropdown.selection = foundIndex;
    return true;
}

/**
 * 名前からドロップダウンを安全に選択する
 * @param {DropDownList} dropdown 対象のドロップダウン
 * @param {Array<string>} names 表示名の一覧
 * @param {string} targetName 設定する名前
 * @returns {void}
 */
function safeAssignDropdownFromName(dropdown, names, targetName) {
    if (!targetName) return false;
    var foundIndex = findIndexByDisplayName(names, targetName);
    if (foundIndex < 0) return false;
    dropdown.selection = foundIndex;
    return true;
}

/**
 * チェックボックスへ値を安全に設定する
 * @param {Checkbox} checkbox 対象のチェックボックス
 * @param {*} value 設定する値
 * @returns {void}
 */
function safeAssignCheckbox(checkbox, value) {
    checkbox.value = !!value;
}

/**
 * 数値入力欄へ値を安全に設定する
 * @param {EditText} editText 対象の入力欄
 * @param {*} value 設定する値
 * @returns {void}
 */
function safeAssignNumberInput(input, value) {
    if (typeof value === "number") input.text = String(value);
}

/**
 * 存在しない場合もあるプロパティを安全に設定する
 * @param {object} targetObject 設定先のオブジェクト
 * @param {string} propertyName プロパティ名
 * @param {*} value 設定する値
 * @returns {boolean} 設定できたら true
 */
function safeSetProperty(targetObject, propertyName, value) {
    try {
        targetObject[propertyName] = value;
        return true;
    } catch (propertySetError) { }
    return false;
}

/* プリセット定義（段落スタイル設定と環境設定を分離） / Preset definitions split into paragraph style settings and app preferences */

var PRESETS = {
    "欧文組版": {
        styleSettings: {
            kinsoku: "弱い禁則",
            kinsokuType: "調整量を優先",
            kinsokuHangType: "なし",
            bunriKinshi: true,
            mojikumi: "なし",
            leadingModel: "欧文ベースライン",
            rensuuji: true,
            rotateSingleByte: false,
            absorbLineEndIdeographicSpace: true,
            latinWordBreak: false,
            kerningMethod: "メトリクス",
            autoLeading: 120,
            characterAlignment: "欧文ベースライン",
            gridAlignment: "なし",
            composer: "欧文段落コンポーザー",
            hyphenation: true,
            hyphenateWordsLongerThan: 6,
            hyphenateAfterFirst: 3,
            hyphenateBeforeLast: 3,
            hyphenateLadderLimit: 2,
            hyphenationZone: 6,
            hyphenateCapitalizedWords: false,
            hyphenateAcrossColumns: false,
            hyphenateLastWord: false,
            ligatures: true,
            language: "en"
        },
        appPreferences: {
            useSmartQuotes: true,
            doubleQuotes: "“”",
            singleQuotes: "‘’",
            textSizeUnit: "ポイント",
            compositionUnit: "ポイント"
        }
    },
    "グリッド優先": {
        styleSettings: {
            kinsoku: "強い禁則",
            kinsokuType: "追い込み優先",
            kinsokuHangType: "なし",
            bunriKinshi: true,
            mojikumi: "行末約物半角",
            leadingModel: "仮想ボディの中央",
            rensuuji: true,
            rotateSingleByte: false,
            absorbLineEndIdeographicSpace: true,
            latinWordBreak: false,
            kerningMethod: "和文等幅",
            autoLeading: 100,
            characterAlignment: "仮想ボディの中央",
            gridAlignment: "仮想ボディの中央",
            composer: "日本語単数行コンポーザー",
            hyphenation: false,
            hyphenateWordsLongerThan: 6,
            hyphenateAfterFirst: 3,
            hyphenateBeforeLast: 3,
            hyphenateLadderLimit: 2,
            hyphenationZone: 10,
            hyphenateCapitalizedWords: false,
            hyphenateAcrossColumns: false,
            hyphenateLastWord: false,
            ligatures: true,
            language: "ja"
        },
        appPreferences: {
            useSmartQuotes: false,
            doubleQuotes: "“”",
            singleQuotes: "‘’",
            textSizeUnit: "級",
            compositionUnit: "歯"
        }
    },
    "グリッド無視": {
        styleSettings: {
            kinsoku: "弱い禁則",
            kinsokuType: "調整量を優先",
            kinsokuHangType: "なし",
            bunriKinshi: true,
            mojikumi: "行末約物半角",
            leadingModel: "仮想ボディの中央",
            rensuuji: true,
            rotateSingleByte: false,
            absorbLineEndIdeographicSpace: true,
            latinWordBreak: false,
            kerningMethod: "メトリクス",
            autoLeading: 175,
            characterAlignment: "仮想ボディの中央",
            gridAlignment: "なし",
            composer: "日本語段落コンポーザー",
            hyphenation: false,
            hyphenateWordsLongerThan: 6,
            hyphenateAfterFirst: 3,
            hyphenateBeforeLast: 3,
            hyphenateLadderLimit: 2,
            hyphenationZone: 10,
            hyphenateCapitalizedWords: false,
            hyphenateAcrossColumns: false,
            hyphenateLastWord: false,
            ligatures: true,
            language: "ja"
        },
        appPreferences: {
            useSmartQuotes: true,
            doubleQuotes: "“”",
            singleQuotes: "‘’",
            textSizeUnit: "ポイント",
            compositionUnit: "ポイント"
        }
    },
    "ソースコード": {
        styleSettings: {
            kinsoku: "弱い禁則",
            kinsokuType: "追い込み優先",
            kinsokuHangType: "なし",
            bunriKinshi: false,
            mojikumi: "なし",
            leadingModel: "欧文ベースライン",
            rensuuji: false,
            rotateSingleByte: false,
            absorbLineEndIdeographicSpace: false,
            latinWordBreak: false,
            kerningMethod: "0",
            autoLeading: 120,
            characterAlignment: "欧文ベースライン",
            gridAlignment: "なし",
            composer: "欧文単数行コンポーザー",
            hyphenation: false,
            hyphenateWordsLongerThan: 6,
            hyphenateAfterFirst: 3,
            hyphenateBeforeLast: 3,
            hyphenateLadderLimit: 2,
            hyphenationZone: 1.25,
            hyphenateCapitalizedWords: false,
            hyphenateAcrossColumns: false,
            hyphenateLastWord: false,
            ligatures: false,
            language: "none"
        },
        appPreferences: {
            useSmartQuotes: false,
            doubleQuotes: "\"\"",
            singleQuotes: "''",
            textSizeUnit: "ポイント",
            compositionUnit: "ポイント"
        }
    },
    "InDesignのデフォルト": {
        styleSettings: {
            kinsoku: "強い禁則",
            kinsokuType: "追い込み優先",
            kinsokuHangType: "なし",
            bunriKinshi: true,
            mojikumi: "行末約物半角",
            leadingModel: "仮想ボディの上/右",
            rensuuji: true,
            rotateSingleByte: false,
            absorbLineEndIdeographicSpace: true,
            latinWordBreak: false,
            kerningMethod: "和文等幅",
            autoLeading: 175,
            characterAlignment: "仮想ボディの中央",
            gridAlignment: "なし",
            composer: "日本語段落コンポーザー",
            hyphenation: true,
            hyphenateWordsLongerThan: 5,
            hyphenateAfterFirst: 2,
            hyphenateBeforeLast: 2,
            hyphenateLadderLimit: 3,
            hyphenationZone: 10,
            hyphenateCapitalizedWords: true,
            hyphenateAcrossColumns: true,
            hyphenateLastWord: true,
            ligatures: true,
            language: "ja"
        },
        appPreferences: {
            useSmartQuotes: false,
            doubleQuotes: "“”",
            singleQuotes: "‘’",
            textSizeUnit: "ポイント",
            compositionUnit: "ポイント"
        }
    }
};

/**
 * 段落スタイル向けプリセット項目の定義を作る
 * @returns {object} プリセット項目の定義
 */
function createStylePresetFields(dialogUi, dialogData) {
    return [
        { key: "kinsoku", type: "dd", control: dialogUi.kinsokuDropdown, names: dialogData.kinsokuNames },
        { key: "kinsokuType", type: "dd", control: dialogUi.kinsokuTypeDropdown, names: dialogData.kinsokuTypeNames },
        { key: "kinsokuHangType", type: "dd", control: dialogUi.kinsokuHangTypeDropdown, names: dialogData.kinsokuHangTypeNames },
        { key: "bunriKinshi", type: "cb", control: dialogUi.bunriKinshiCheckbox },
        { key: "mojikumi", type: "dd", control: dialogUi.mojikumiDropdown, names: dialogData.mojikumiNames },
        { key: "leadingModel", type: "dd", control: dialogUi.leadingModelDropdown, names: dialogData.leadingModelNames },
        { key: "rensuuji", type: "cb", control: dialogUi.rensuujiCheckbox },
        { key: "rotateSingleByte", type: "cb", control: dialogUi.rotateSingleByteCheckbox },
        { key: "absorbLineEndIdeographicSpace", type: "cb", control: dialogUi.absorbLineEndIdeographicSpaceCheckbox },
        { key: "latinWordBreak", type: "cb", control: dialogUi.latinWordBreakCheckbox },
        { key: "kerningMethod", type: "dd", control: dialogUi.kerningMethodDropdown, names: dialogData.kerningMethodNames },
        { key: "autoLeading", type: "in", control: dialogUi.autoLeadingInput },
        { key: "characterAlignment", type: "dd", control: dialogUi.characterAlignmentDropdown, names: dialogData.characterAlignmentNames },
        { key: "gridAlignment", type: "dd", control: dialogUi.gridAlignmentDropdown, names: dialogData.gridAlignmentNames },
        { key: "composer", type: "dd", control: dialogUi.composerDropdown, names: dialogData.composerNames },
        { key: "hyphenation", type: "cb", control: dialogUi.hyphenationCheckbox },
        { key: "hyphenateWordsLongerThan", type: "in", control: dialogUi.hyphenateWordsLongerThanInput },
        { key: "hyphenateAfterFirst", type: "in", control: dialogUi.hyphenateAfterFirstInput },
        { key: "hyphenateBeforeLast", type: "in", control: dialogUi.hyphenateBeforeLastInput },
        { key: "hyphenateLadderLimit", type: "in", control: dialogUi.hyphenateLadderLimitInput },
        { key: "hyphenationZone", type: "in", control: dialogUi.hyphenationZoneInput },
        { key: "hyphenateCapitalizedWords", type: "cb", control: dialogUi.hyphenateCapitalizedWordsCheckbox },
        { key: "hyphenateAcrossColumns", type: "cb", control: dialogUi.hyphenateAcrossColumnsCheckbox },
        { key: "hyphenateLastWord", type: "cb", control: dialogUi.hyphenateLastWordCheckbox },
        { key: "ligatures", type: "cb", control: dialogUi.ligaturesCheckbox }
    ];
}

/**
 * 環境設定向けプリセット項目の定義を作る
 * @returns {object} プリセット項目の定義
 */
function createAppPreferencePresetFields(dialogUi) {
    return [
        { key: "textSizeUnit", type: "dd", control: dialogUi.textSizeUnitDropdown, names: ["ポイント", "級", "アメリカ式ポイント"] },
        { key: "compositionUnit", type: "dd", control: dialogUi.compositionUnitDropdown, names: ["ポイント", "歯", "U", "倍", "ミルス", "アメリカ式ポイント"] }
    ];
}

/**
 * プリセット項目の定義をまとめて作る
 * @returns {object} プリセット項目の定義
 */
function createPresetFields(dialogUi, dialogData) {
    return {
        styleFields: createStylePresetFields(dialogUi, dialogData),
        appPreferenceFields: createAppPreferencePresetFields(dialogUi)
    };
}

/**
 * ハイフネーションの ON/OFF に応じて関連項目を切り替える
 * @returns {void}
 */
function updateHyphenationControlsEnabled(dialogUi) {
    var isEnabled = dialogUi.hyphenationCheckbox.value;
    dialogUi.hyphenateWordsLongerThanInput.parent.enabled = isEnabled;
    dialogUi.hyphenateAfterFirstInput.parent.enabled = isEnabled;
    dialogUi.hyphenateBeforeLastInput.parent.enabled = isEnabled;
    dialogUi.hyphenateLadderLimitInput.parent.enabled = isEnabled;
    dialogUi.hyphenationZoneInput.parent.enabled = isEnabled;
    dialogUi.hyphenateCapitalizedWordsCheckbox.enabled = isEnabled;
    dialogUi.hyphenateAcrossColumnsCheckbox.enabled = isEnabled;
    dialogUi.hyphenateLastWordCheckbox.enabled = isEnabled;
}

/**
 * 対象範囲のラジオボタンを選択状態にする
 * @param {string} targetValue 選択する対象範囲
 * @returns {void}
 */
function activateTargetRadio(dialogUi, activeRadio) {
    var targetRadios = [dialogUi.targetAllRadio, dialogUi.targetSelectionRadio, dialogUi.targetSelectedParagraphsRadio];
    for (var radioIndex = 0; radioIndex < targetRadios.length; radioIndex++) {
        targetRadios[radioIndex].value = (targetRadios[radioIndex] === activeRadio);
    }
}

/**
 * 言語のラジオボタンを選択状態にする
 * @param {string} languageKey 選択する言語のキー
 * @returns {void}
 */
function activateLanguageRadio(dialogUi, activeRadio) {
    var languageRadios = [dialogUi.languageJapaneseRadio, dialogUi.languageEnglishRadio, dialogUi.languageNoneRadio];
    for (var radioIndex = 0; radioIndex < languageRadios.length; radioIndex++) {
        languageRadios[radioIndex].value = (languageRadios[radioIndex] === activeRadio);
    }
}

/**
 * 選択中の言語キーを取得する
 * @returns {string} 言語のキー
 */
function getLanguageSelection(dialogUi) {
    if (dialogUi.languageEnglishRadio.value) return "en";
    if (dialogUi.languageNoneRadio.value) return "none";
    return "ja";
}

/**
 * 選択から最初の段落を取り出す
 * @returns {Paragraph|null} 段落。取得できない場合は null
 */
function getFirstParagraphFromSelection() {
    try {
        var selectionItems = app.selection;
        if (!selectionItems || selectionItems.length === 0) return null;
        var firstSelectedItem = selectionItems[0];
        if (firstSelectedItem.paragraphs && firstSelectedItem.paragraphs.length > 0) {
            return firstSelectedItem.paragraphs.firstItem();
        }
        if (firstSelectedItem.parentStory && firstSelectedItem.parentStory.paragraphs && firstSelectedItem.parentStory.paragraphs.length > 0) {
            return firstSelectedItem.parentStory.paragraphs.firstItem();
        }
    } catch (selectionParagraphReadError) { }
    return null;
}

/**
 * 文字組みアキ量設定の表示名を求める
 * @param {*} mojikumiValue 文字組みの設定値
 * @returns {string} 表示名
 */
function resolveMojikumiName(mojikumiValue) {
    if (mojikumiValue === null || mojikumiValue === undefined || mojikumiValue === NothingEnum.NOTHING) {
        return "なし";
    }
    if (typeof mojikumiValue === "string") {
        return mojikumiValue;
    }

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

/**
 * 言語オブジェクトから言語キーを求める
 * @param {*} languageValue 言語の設定値
 * @returns {string} 言語のキー
 */
function resolveLanguageKey(appliedLanguage) {
    var languageName = "";
    if (appliedLanguage === null || appliedLanguage === undefined) {
        languageName = "[言語なし]";
    } else if (typeof appliedLanguage === "string") {
        languageName = appliedLanguage;
    } else {
        try { languageName = appliedLanguage.name; } catch (languageNameError) { }
    }

    for (var languageKey in LANGUAGE_CANDIDATES) {
        var languageCandidates = LANGUAGE_CANDIDATES[languageKey];
        for (var candidateIndex = 0; candidateIndex < languageCandidates.length; candidateIndex++) {
            if (languageCandidates[candidateIndex] === languageName) {
                return languageKey;
            }
        }
    }

    return null;
}

/**
 * 禁則処理セットの設定値から選択位置を探す
 * @param {object} kinsokuTableData 禁則処理セットの一覧
 * @param {*} kinsokuValue 禁則処理セットの設定値
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
        for (var nameIndex = 0; nameIndex < kinsokuNames.length; nameIndex++) {
            if (kinsokuNames[nameIndex] === kinsokuValueName) return nameIndex;
        }
    }
    return -1;
}

/**
 * 選択段落から禁則関連の設定を読み取る
 * @param {Paragraph} paragraph 対象の段落
 * @param {object} lookupTables 選択肢の参照表
 * @returns {object} 読み取った設定
 */
function loadKinsokuSettingsFromParagraph(paragraphObject, appliedParagraphStyle, dialogUi, dialogData, dialogLookupTables) {
    var kinsokuIndex = -1;
    if (appliedParagraphStyle) {
        try { kinsokuIndex = findKinsokuIndexFromValue(appliedParagraphStyle.kinsokuSet, dialogLookupTables.kinsokuTables, dialogData.kinsokuNames); } catch (eKinsokuFromStyle) { }
    }
    if (kinsokuIndex < 0) {
        try { kinsokuIndex = findKinsokuIndexFromValue(paragraphObject.kinsokuSet, dialogLookupTables.kinsokuTables, dialogData.kinsokuNames); } catch (eKinsokuFromParagraph) { }
    }
    if (kinsokuIndex >= 0) {
        dialogUi.kinsokuDropdown.selection = kinsokuIndex;
    }

    var assignedKinsokuType = false;
    try {
        assignedKinsokuType = safeAssignDropdownFromEnum(dialogUi.kinsokuTypeDropdown, dialogLookupTables.kinsokuTypeValues, paragraphObject.kinsokuType);
    } catch (eKinsokuTypeFromParagraph) { }
    if (!assignedKinsokuType && appliedParagraphStyle) {
        try { safeAssignDropdownFromEnum(dialogUi.kinsokuTypeDropdown, dialogLookupTables.kinsokuTypeValues, appliedParagraphStyle.kinsokuType); } catch (eKinsokuTypeFromStyle) { }
    }

    var assignedKinsokuHangType = false;
    try {
        assignedKinsokuHangType = safeAssignDropdownFromEnum(dialogUi.kinsokuHangTypeDropdown, dialogLookupTables.kinsokuHangTypeValues, paragraphObject.kinsokuHangType);
    } catch (eKinsokuHangTypeFromParagraph) { }
    if (!assignedKinsokuHangType && appliedParagraphStyle) {
        try { safeAssignDropdownFromEnum(dialogUi.kinsokuHangTypeDropdown, dialogLookupTables.kinsokuHangTypeValues, appliedParagraphStyle.kinsokuHangType); } catch (eKinsokuHangTypeFromStyle) { }
    }
}

/**
 * 選択段落からハイフネーション設定を読み取る
 * @param {Paragraph} paragraph 対象の段落
 * @returns {object} 読み取った設定
 */
function loadHyphenationSettingsFromParagraph(paragraphObject, dialogUi) {
    safeAssignCheckbox(dialogUi.hyphenationCheckbox, paragraphObject.hyphenation);
    safeAssignNumberInput(dialogUi.hyphenateWordsLongerThanInput, paragraphObject.hyphenateWordsLongerThan);
    safeAssignNumberInput(dialogUi.hyphenateAfterFirstInput, paragraphObject.hyphenateAfterFirst);
    safeAssignNumberInput(dialogUi.hyphenateBeforeLastInput, paragraphObject.hyphenateBeforeLast);
    try { safeAssignNumberInput(dialogUi.hyphenateLadderLimitInput, paragraphObject.hyphenateLadderLimit); } catch (hyphenateLadderLimitReadError) { }

    try {
        var savedMeasurementUnit = app.scriptPreferences.measurementUnit;
        app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
        var hyphenationZonePt;
        try {
            hyphenationZonePt = paragraphObject.hyphenationZone;
        } catch (eHyphenationZoneRead) {
            hyphenationZonePt = null;
        }
        app.scriptPreferences.measurementUnit = savedMeasurementUnit;
        if (typeof hyphenationZonePt === "number") {
            var hyphenationZoneMm = hyphenationZonePt / 2.834645669;
            dialogUi.hyphenationZoneInput.text = String(Math.round(hyphenationZoneMm * 100) / 100);
        }
    } catch (eHyphenationZone) { }

    try { safeAssignCheckbox(dialogUi.hyphenateCapitalizedWordsCheckbox, paragraphObject.hyphenateCapitalizedWords); } catch (eHyphenateCapitalizedWords) { }
    try { safeAssignCheckbox(dialogUi.hyphenateAcrossColumnsCheckbox, paragraphObject.hyphenateAcrossColumns); } catch (eHyphenateAcrossColumns) { }
    try { safeAssignCheckbox(dialogUi.hyphenateLastWordCheckbox, paragraphObject.hyphenateLastWord); } catch (eHyphenateLastWord) { }
}

/**
 * 選択段落から組版設定をまとめて読み取る
 * @param {Paragraph} paragraph 対象の段落
 * @param {object} lookupTables 選択肢の参照表
 * @returns {object} 読み取った設定
 */
function loadSettingsFromParagraph(paragraphObject, dialogUi, dialogData) {
    if (!paragraphObject || !dialogData.lookupTables) return;

    var dialogLookupTables = dialogData.lookupTables;
    var appliedParagraphStyle = null;
    try {
        if (paragraphObject.appliedParagraphStyle && paragraphObject.appliedParagraphStyle.isValid) {
            appliedParagraphStyle = paragraphObject.appliedParagraphStyle;
        }
    } catch (eAppliedStyle) { }
    loadKinsokuSettingsFromParagraph(paragraphObject, appliedParagraphStyle, dialogUi, dialogData, dialogLookupTables);
    try {
        safeAssignDropdownFromName(dialogUi.mojikumiDropdown, dialogData.mojikumiNames, resolveMojikumiName(paragraphObject.mojikumi));
    } catch (eMojikumi) { }

    safeAssignDropdownFromEnum(dialogUi.leadingModelDropdown, dialogLookupTables.leadingModelValues, paragraphObject.leadingModel);
    safeAssignDropdownFromEnum(dialogUi.characterAlignmentDropdown, dialogLookupTables.characterAlignmentValues, paragraphObject.characterAlignment);
    safeAssignDropdownFromEnum(dialogUi.gridAlignmentDropdown, dialogLookupTables.gridAlignmentValues, paragraphObject.gridAlignment);
    safeAssignDropdownFromEnum(dialogUi.kerningMethodDropdown, dialogLookupTables.kerningMethodValues, paragraphObject.kerningMethod);
    safeAssignNumberInput(dialogUi.autoLeadingInput, paragraphObject.autoLeading);
    try {
        var matchedComposerIndex = findIndexByComposerAliases(dialogLookupTables.composerAliases, paragraphObject.composer);
        if (matchedComposerIndex >= 0) dialogUi.composerDropdown.selection = matchedComposerIndex;
    } catch (eComposer) { }
    safeAssignCheckbox(dialogUi.bunriKinshiCheckbox, paragraphObject.bunriKinshi);
    safeAssignCheckbox(dialogUi.rensuujiCheckbox, paragraphObject.rensuuji);
    safeAssignCheckbox(dialogUi.rotateSingleByteCheckbox, paragraphObject.rotateSingleByteCharacters);
    safeAssignCheckbox(dialogUi.absorbLineEndIdeographicSpaceCheckbox, paragraphObject.treatIdeographicSpaceAsSpace);
    try {
        safeAssignCheckbox(dialogUi.latinWordBreakCheckbox, paragraphObject.allowArbitraryHyphenation);
    } catch (eLatinWordBreakFromParagraph) {
        if (appliedParagraphStyle) {
            try { safeAssignCheckbox(dialogUi.latinWordBreakCheckbox, appliedParagraphStyle.allowArbitraryHyphenation); } catch (eLatinWordBreakFromStyle) { }
        }
    }
    try { safeAssignCheckbox(dialogUi.ligaturesCheckbox, paragraphObject.ligatures); } catch (eLigatures) { }
    loadHyphenationSettingsFromParagraph(paragraphObject, dialogUi);

    try {
        var matchedLanguageKey = resolveLanguageKey(paragraphObject.appliedLanguage);
        if (matchedLanguageKey === "ja") activateLanguageRadio(dialogUi, dialogUi.languageJapaneseRadio);
        else if (matchedLanguageKey === "en") activateLanguageRadio(dialogUi, dialogUi.languageEnglishRadio);
        else if (matchedLanguageKey === "none") activateLanguageRadio(dialogUi, dialogUi.languageNoneRadio);
    } catch (eLang) { }

    updateHyphenationControlsEnabled(dialogUi);
}

/**
 * プリセットの値をダイアログへ反映する
 * @param {string} presetKey プリセットのキー
 * @returns {void}
 */
function applyPreset(presetName, dialogUi, presetFields) {
    var preset = PRESETS[presetName];
    if (!preset) return;
    var presetStyleSettings = preset.styleSettings || preset;
    var presetAppPreferences = preset.appPreferences || preset;

    var mergedPresetFields = presetFields.styleFields.concat(presetFields.appPreferenceFields);

    for (var fieldIndex = 0; fieldIndex < mergedPresetFields.length; fieldIndex++) {
        var presetField = mergedPresetFields[fieldIndex];
        if (presetStyleSettings[presetField.key] === undefined && presetAppPreferences[presetField.key] === undefined) continue;
        var presetValue = presetStyleSettings[presetField.key] !== undefined ? presetStyleSettings[presetField.key] : presetAppPreferences[presetField.key];
        if (presetField.type === "dd") {
            safeAssignDropdownFromName(presetField.control, presetField.names, presetValue);
        } else if (presetField.type === "cb") {
            presetField.control.value = !!presetValue;
        } else if (presetField.type === "in") {
            presetField.control.text = String(presetValue);
        }
    }

    if (presetAppPreferences.useSmartQuotes !== undefined) {
        dialogUi.useTypographersQuotesCheckbox.value = !!presetAppPreferences.useSmartQuotes;
    }
    if (presetAppPreferences.doubleQuotes !== undefined) {
        safeAssignDropdownFromName(dialogUi.smartQuoteDropdown, DOUBLE_QUOTE_OPTIONS, presetAppPreferences.doubleQuotes);
    }
    if (presetAppPreferences.singleQuotes !== undefined) {
        safeAssignDropdownFromName(dialogUi.smartSingleQuoteDropdown, SINGLE_QUOTE_OPTIONS, presetAppPreferences.singleQuotes);
    }
    if (presetStyleSettings.language !== undefined) {
        if (presetStyleSettings.language === "en") activateLanguageRadio(dialogUi, dialogUi.languageEnglishRadio);
        else if (presetStyleSettings.language === "none") activateLanguageRadio(dialogUi, dialogUi.languageNoneRadio);
        else activateLanguageRadio(dialogUi, dialogUi.languageJapaneseRadio);
    }
    updateHyphenationControlsEnabled(dialogUi);
}

/**
 * プリセット名を入力するダイアログを表示する
 * @returns {string|null} 入力された名前。キャンセル時は null
 */
function showPresetNameInputDialog() {
    var nameDialog = new Window("dialog", getLabel("dialog.presetNameInput"));
    setupWindow(nameDialog, 10);

    nameDialog.add("statictext", undefined, getLabel("field.presetName"));
    var nameInput = nameDialog.add("edittext", undefined, "");
    nameInput.preferredSize = [320, -1];
    nameInput.active = true;

    var buttonGroup = nameDialog.add("group");
    buttonGroup.alignment = "right";
    buttonGroup.add("button", undefined, "キャンセル", { name: "cancel" });
    buttonGroup.add("button", undefined, "OK", { name: "ok" });

    if (nameDialog.show() !== 1) return null;
    var presetName = nameInput.text;
    if (!presetName) return null;
    return presetName.replace(/^\s+|\s+$/g, "");
}

/**
 * 現在の設定をプリセット定義のコードとして組み立てる
 * @param {string} presetName プリセット名
 * @returns {string} 書き出すコード
 */
function buildPresetCodeSnippet(presetName, presetFields, dialogUi) {
    var appPreferenceKeys = {
        useSmartQuotes: true,
        doubleQuotes: true,
        singleQuotes: true,
        textSizeUnit: true,
        compositionUnit: true
    };

    var styleSettingLines = [];
    var appPreferenceLines = [];

    styleSettingLines.push("        language: \"" + getLanguageSelection(dialogUi) + "\"");
    appPreferenceLines.push("        useSmartQuotes: " + (dialogUi.useTypographersQuotesCheckbox.value ? "true" : "false"));
    appPreferenceLines.push("        doubleQuotes: \"" + (dialogUi.smartQuoteDropdown.selection ? dialogUi.smartQuoteDropdown.selection.text : "") + "\"");
    appPreferenceLines.push("        singleQuotes: \"" + (dialogUi.smartSingleQuoteDropdown.selection ? dialogUi.smartSingleQuoteDropdown.selection.text : "") + "\"");

    var mergedPresetFields = presetFields.styleFields.concat(presetFields.appPreferenceFields);

    for (var fieldIndex = 0; fieldIndex < mergedPresetFields.length; fieldIndex++) {
        var presetField = mergedPresetFields[fieldIndex];
        var valueText;

        if (presetField.type === "dd") {
            valueText = "\"" + presetField.control.selection.text + "\"";
        } else if (presetField.type === "cb") {
            valueText = presetField.control.value ? "true" : "false";
        } else {
            valueText = presetField.control.text;
            if (isNaN(parseFloat(valueText)) || String(parseFloat(valueText)) !== valueText) {
                valueText = "\"" + valueText + "\"";
            }
        }

        if (appPreferenceKeys[presetField.key]) {
            appPreferenceLines.push("        " + presetField.key + ": " + valueText);
        } else {
            styleSettingLines.push("        " + presetField.key + ": " + valueText);
        }
    }

    var lines = [];
    lines.push(getLabel("export.codeHeader"));
    lines.push("PRESETS[\"" + presetName + "\"] = {");
    lines.push("    styleSettings: {");
    lines.push(styleSettingLines.join(",\n"));
    lines.push("    },");
    lines.push("    appPreferences: {");
    lines.push(appPreferenceLines.join(",\n"));
    lines.push("    }");
    lines.push("};");
    return lines.join("\n");
}

/**
 * ファイル名に使えない文字を置き換える
 * @param {string} fileName 元のファイル名
 * @returns {string} 安全なファイル名
 */
function sanitizeFileName(name) {
    return name.replace(/[\/\\:*?"<>|]/g, "_");
}

/**
 * 現在の設定をプリセットコードとしてデスクトップへ書き出す
 * @returns {void}
 */
function exportPresetCode(presetFields, dialogUi) {
    var presetName = showPresetNameInputDialog();
    if (!presetName) return;

    var code = buildPresetCodeSnippet(presetName, presetFields, dialogUi);
    var safeFileName = sanitizeFileName(presetName);

    try {
        var file = File(Folder.desktop + "/" + encodeURI(safeFileName) + ".jsx");
        if (file.exists) {
            if (!confirm(getLabel("export.overwritePrefix") + safeFileName + getLabel("export.overwriteSuffix"))) return;
        }
        file.encoding = "UTF-8";
        if (file.open("w")) {
            file.write(code);
            file.close();
            var savedMessage = getLabel("export.savedPrefix") + presetName + getLabel("export.savedSuffix");
            if (safeFileName !== presetName) {
                savedMessage += "\nファイル名: " + safeFileName + ".jsx";
            }
            alert(savedMessage);
        } else {
            alert(getLabel("alert.openFileFailed"));
        }
    } catch (eExport) {
        alert(getLabel("alert.exportErrorPrefix") + eExport.message);
    }
}

/**
 * 文字組版設定ダイアログを組み立てる
 * @param {object} lookupTables 選択肢の参照表
 * @param {object} initialSettings 初期値
 * @returns {object} ダイアログとコントロール
 */
function createDialogUI(dialogData) {
    var defaultIndexes = dialogData.defaultIndexes;
    var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(dialog, 10);

    var topColumnsGroup = dialog.add("group");
    topColumnsGroup.orientation = "row";
    topColumnsGroup.alignChildren = ["fill", "top"];
    topColumnsGroup.spacing = 10;

    var targetPanel = topColumnsGroup.add("panel", undefined, getLabel("panel.targetStyles"));
    setupPanel(targetPanel, 8);
    targetPanel.orientation = "row";
    targetPanel.alignChildren = ["left", "center"];

    var targetSelectedParagraphsRadio = targetPanel.add("radiobutton", undefined, getLabel("radio.targetSelection"));
    var targetAllRadio = targetPanel.add("radiobutton", undefined, getLabel("radio.targetAll"));
    var targetSelectionRadio = targetPanel.add("radiobutton", undefined, getLabel("radio.targetSpecified"));
    var targetSelectionButton = targetPanel.add("button", undefined, getLabel("button.select"));
    targetSelectedParagraphsRadio.value = true;

    var presetPanel = topColumnsGroup.add("panel", undefined, getLabel("panel.preset"));
    setupPanel(presetPanel, 8);
    presetPanel.orientation = "row";

    var presetDropdown = presetPanel.add("dropdownlist", undefined, ["欧文組版", "グリッド優先", "グリッド無視", "ソースコード", "InDesignのデフォルト"]);
    presetDropdown.selection = null;
    presetDropdown.preferredSize.width = W_DROP;
    var presetExportButton = presetPanel.add("button", undefined, getLabel("button.export"));

    var columnsGroup = dialog.add("group");
    columnsGroup.orientation = "row";
    columnsGroup.alignChildren = ["fill", "top"];
    columnsGroup.spacing = 10;

    var leftColumn = columnsGroup.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = "fill";
    leftColumn.spacing = 10;

    var rightColumn = columnsGroup.add("group");
    rightColumn.orientation = "column";
    rightColumn.alignChildren = "fill";
    rightColumn.spacing = 10;

    var compositionExtraPanel = leftColumn.add("panel", undefined, getLabel("panel.basicSettings"));
    setupPanel(compositionExtraPanel, 8);

    var kerningMethodDropdown = addDropdownRow(compositionExtraPanel, getLabel("field.autoKerning"), dialogData.kerningMethodNames, defaultIndexes.kerningMethodIndex);
    var autoLeadingInput = addNumberRow(compositionExtraPanel, getLabel("field.autoLeading"), defaultIndexes.autoLeadingPercent, "%");
    var characterAlignmentDropdown = addDropdownRow(compositionExtraPanel, getLabel("field.characterAlign"), dialogData.characterAlignmentNames, defaultIndexes.characterAlignmentIndex);
    var leadingModelDropdown = addDropdownRow(compositionExtraPanel, getLabel("field.leadingModel"), dialogData.leadingModelNames, defaultIndexes.leadingModelIndex);
    var gridAlignmentDropdown = addDropdownRow(compositionExtraPanel, getLabel("field.gridAlignment"), dialogData.gridAlignmentNames, defaultIndexes.gridAlignmentIndex);
    var composerDropdown = addDropdownRow(compositionExtraPanel, getLabel("field.composer"), dialogData.composerNames, defaultIndexes.composerIndex);

    var compositionOptionalPanel = rightColumn.add("panel", undefined, getLabel("panel.quotes"));
    setupPanel(compositionOptionalPanel, 8);

    var useTypographersQuotesInitial;

    try {
        useTypographersQuotesInitial = !!app.textPreferences.typographersQuotes;
    } catch (typographersQuotesReadError) {
        useTypographersQuotesInitial = true;
    }

    var useTypographersQuotesCheckbox = compositionOptionalPanel.add(
        "checkbox",
        undefined,
        "英文引用符を使用"
    );

    useTypographersQuotesCheckbox.value = useTypographersQuotesInitial;

    var smartQuoteDropdown = addDropdownRow(
        compositionOptionalPanel,
        getLabel("field.doubleQuote"),
        DOUBLE_QUOTE_OPTIONS,
        0
    );

    var smartSingleQuoteDropdown = addDropdownRow(
        compositionOptionalPanel,
        getLabel("field.singleQuote"),
        SINGLE_QUOTE_OPTIONS,
        0
    );

    var unitsPanel = rightColumn.add("panel", undefined, getLabel("panel.units"));
    setupPanel(unitsPanel, 8);

    var textSizeUnitNames = ["ポイント", "級", "アメリカ式ポイント"];
    var textSizeUnitDropdown = addDropdownRow(unitsPanel, getLabel("field.textSize"), textSizeUnitNames, 0);
    textSizeUnitDropdown.preferredSize.width = W_DROP;

    var compositionUnitNames = ["ポイント", "歯", "U", "倍", "ミルス", "アメリカ式ポイント"];
    var compositionUnitDropdown = addDropdownRow(unitsPanel, getLabel("field.typography"), compositionUnitNames, 0);
    compositionUnitDropdown.preferredSize.width = W_DROP;

    var languageRow = compositionExtraPanel.add("group");
    languageRow.orientation = "row";
    languageRow.alignChildren = ["left", "center"];
    languageRow.spacing = 8;

    var languageLabel = languageRow.add("statictext", undefined, getLabel("field.language"));
    languageLabel.preferredSize.width = 120;

    var languageJapaneseRadio = languageRow.add("radiobutton", undefined, getLabel("radio.languageJapanese"));
    var languageEnglishRadio = languageRow.add("radiobutton", undefined, getLabel("radio.languageEnglish"));
    var languageNoneRadio = languageRow.add("radiobutton", undefined, getLabel("radio.languageNone"));
    languageJapaneseRadio.value = (defaultIndexes.language !== "en" && defaultIndexes.language !== "none");
    languageEnglishRadio.value = (defaultIndexes.language === "en");
    languageNoneRadio.value = (defaultIndexes.language === "none");

    var ligaturesCheckbox = compositionExtraPanel.add("checkbox", undefined, getLabel("checkbox.ligatures"));
    ligaturesCheckbox.value = defaultIndexes.ligatures;

    var compositionPanel = leftColumn.add("panel", undefined, getLabel("panel.japaneseTypeset"));
    setupPanel(compositionPanel, 8);

    var kinsokuDropdown = addDropdownRow(compositionPanel, getLabel("field.kinsokuSet"), dialogData.kinsokuNames, defaultIndexes.kinsokuIndex);
    var kinsokuTypeDropdown = addDropdownRow(compositionPanel, getLabel("field.kinsokuType"), dialogData.kinsokuTypeNames, defaultIndexes.kinsokuTypeIndex);
    var kinsokuHangTypeDropdown = addDropdownRow(compositionPanel, getLabel("field.kinsokuHangType"), dialogData.kinsokuHangTypeNames, defaultIndexes.kinsokuHangTypeIndex);

    var bunriKinshiCheckbox = compositionPanel.add("checkbox", undefined, getLabel("checkbox.noBreak"));
    bunriKinshiCheckbox.value = defaultIndexes.bunriKinshi;

    var mojikumiDropdown = addDropdownRow(compositionPanel, getLabel("field.mojikumi"), dialogData.mojikumiNames, defaultIndexes.mojikumiIndex);

    var compositionCheckboxesGroup = compositionPanel.add("group");
    compositionCheckboxesGroup.orientation = "row";
    compositionCheckboxesGroup.alignChildren = ["fill", "top"];
    compositionCheckboxesGroup.spacing = 16;
    compositionCheckboxesGroup.margins = [0, 10, 0, 0];

    var compositionCheckboxesLeft = compositionCheckboxesGroup.add("group");
    compositionCheckboxesLeft.orientation = "column";
    compositionCheckboxesLeft.alignChildren = "left";
    compositionCheckboxesLeft.spacing = 4;

    var compositionCheckboxesRight = compositionCheckboxesGroup.add("group");
    compositionCheckboxesRight.orientation = "column";
    compositionCheckboxesRight.alignChildren = "left";
    compositionCheckboxesRight.spacing = 4;

    var rensuujiCheckbox = compositionCheckboxesLeft.add("checkbox", undefined, getLabel("checkbox.digitsRotation"));
    rensuujiCheckbox.value = defaultIndexes.rensuuji;
    var rotateSingleByteCheckbox = compositionCheckboxesLeft.add("checkbox", undefined, getLabel("checkbox.rotateInVertical"));
    rotateSingleByteCheckbox.value = defaultIndexes.rotateSingleByte;
    var absorbLineEndIdeographicSpaceCheckbox = compositionCheckboxesRight.add("checkbox", undefined, getLabel("checkbox.absorbTrailingSpace"));
    absorbLineEndIdeographicSpaceCheckbox.value = defaultIndexes.absorbLineEndIdeographicSpace;
    var latinWordBreakCheckbox = compositionCheckboxesRight.add("checkbox", undefined, getLabel("checkbox.arbitraryHyphen"));
    latinWordBreakCheckbox.value = defaultIndexes.latinWordBreak;


    var hyphenationPanel = rightColumn.add("panel", undefined, getLabel("panel.hyphenation"));
    setupPanel(hyphenationPanel, 8);
    var hyphenationCheckbox = hyphenationPanel.add("checkbox", undefined, getLabel("checkbox.hyphenation"));
    hyphenationCheckbox.value = defaultIndexes.hyphenation;
    var hyphenateWordsLongerThanInput = addNumberRow(hyphenationPanel, getLabel("field.minWordLength"), defaultIndexes.hyphenateWordsLongerThan, "文字");
    var hyphenateAfterFirstInput = addNumberRow(hyphenationPanel, getLabel("field.afterFirst"), defaultIndexes.hyphenateAfterFirst, "文字");
    var hyphenateBeforeLastInput = addNumberRow(hyphenationPanel, getLabel("field.beforeLast"), defaultIndexes.hyphenateBeforeLast, "文字");
    var hyphenateLadderLimitInput = addNumberRow(hyphenationPanel, getLabel("field.maxHyphens"), defaultIndexes.hyphenateLadderLimit, "ハイフン");
    var hyphenationZoneInput = addNumberRow(hyphenationPanel, getLabel("field.hyphenationZone"), defaultIndexes.hyphenationZoneMm, "mm");

    var hyphenateBreakPanel = hyphenationPanel.add("panel", undefined, getLabel("panel.hyphenateBreak"));
    setupPanel(hyphenateBreakPanel, 8);
    hyphenateBreakPanel.alignChildren = "left";
    var hyphenateCapitalizedWordsCheckbox = hyphenateBreakPanel.add("checkbox", undefined, getLabel("checkbox.capitalizedWords"));
    hyphenateCapitalizedWordsCheckbox.value = defaultIndexes.hyphenateCapitalizedWords;
    var hyphenateAcrossColumnsCheckbox = hyphenateBreakPanel.add("checkbox", undefined, getLabel("checkbox.acrossColumns"));
    hyphenateAcrossColumnsCheckbox.value = defaultIndexes.hyphenateAcrossColumns;
    var hyphenateLastWordCheckbox = hyphenateBreakPanel.add("checkbox", undefined, getLabel("checkbox.lastWord"));
    hyphenateLastWordCheckbox.value = defaultIndexes.hyphenateLastWord;

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";
    buttonGroup.margins = [0, 10, 0, 0];
    buttonGroup.add("button", undefined, "キャンセル", { name: "cancel" });
    buttonGroup.add("button", undefined, "OK", { name: "ok" });

    return {
        dialog: dialog,
        targetSelectedParagraphsRadio: targetSelectedParagraphsRadio,
        targetAllRadio: targetAllRadio,
        targetSelectionRadio: targetSelectionRadio,
        targetSelectionButton: targetSelectionButton,
        presetDropdown: presetDropdown,
        presetExportButton: presetExportButton,
        kinsokuDropdown: kinsokuDropdown,
        kinsokuTypeDropdown: kinsokuTypeDropdown,
        kinsokuHangTypeDropdown: kinsokuHangTypeDropdown,
        bunriKinshiCheckbox: bunriKinshiCheckbox,
        mojikumiDropdown: mojikumiDropdown,
        leadingModelDropdown: leadingModelDropdown,
        rensuujiCheckbox: rensuujiCheckbox,
        rotateSingleByteCheckbox: rotateSingleByteCheckbox,
        absorbLineEndIdeographicSpaceCheckbox: absorbLineEndIdeographicSpaceCheckbox,
        latinWordBreakCheckbox: latinWordBreakCheckbox,
        kerningMethodDropdown: kerningMethodDropdown,
        autoLeadingInput: autoLeadingInput,
        characterAlignmentDropdown: characterAlignmentDropdown,
        gridAlignmentDropdown: gridAlignmentDropdown,
        languageJapaneseRadio: languageJapaneseRadio,
        languageEnglishRadio: languageEnglishRadio,
        languageNoneRadio: languageNoneRadio,
        composerDropdown: composerDropdown,
        ligaturesCheckbox: ligaturesCheckbox,
        hyphenationCheckbox: hyphenationCheckbox,
        hyphenateWordsLongerThanInput: hyphenateWordsLongerThanInput,
        hyphenateAfterFirstInput: hyphenateAfterFirstInput,
        hyphenateBeforeLastInput: hyphenateBeforeLastInput,
        hyphenateLadderLimitInput: hyphenateLadderLimitInput,
        hyphenationZoneInput: hyphenationZoneInput,
        hyphenateCapitalizedWordsCheckbox: hyphenateCapitalizedWordsCheckbox,
        hyphenateAcrossColumnsCheckbox: hyphenateAcrossColumnsCheckbox,
        hyphenateLastWordCheckbox: hyphenateLastWordCheckbox,
        useTypographersQuotesCheckbox: useTypographersQuotesCheckbox,
        smartQuoteDropdown: smartQuoteDropdown,
        smartSingleQuoteDropdown: smartSingleQuoteDropdown,
        textSizeUnitDropdown: textSizeUnitDropdown,
        compositionUnitDropdown: compositionUnitDropdown,
        selectedStyleIndexes: null
    };
}

/**
 * ダイアログのコントロールにイベントを結び付ける
 * @returns {void}
 */
function bindDialogEvents(dialogUi, dialogData, presetFields) {
    dialogUi.targetAllRadio.onClick = function () {
        activateTargetRadio(dialogUi, dialogUi.targetAllRadio);
    };
    dialogUi.targetSelectionRadio.onClick = function () {
        activateTargetRadio(dialogUi, dialogUi.targetSelectionRadio);
    };
    dialogUi.targetSelectedParagraphsRadio.onClick = function () {
        activateTargetRadio(dialogUi, dialogUi.targetSelectedParagraphsRadio);
        loadSettingsFromParagraph(getFirstParagraphFromSelection(), dialogUi, dialogData);
    };

    dialogUi.targetSelectionButton.onClick = function () {
        var pickerResult = showParagraphStylePicker(dialogData.paragraphStyleNames, dialogUi.selectedStyleIndexes);
        if (pickerResult !== null) {
            dialogUi.selectedStyleIndexes = pickerResult;
            activateTargetRadio(dialogUi, dialogUi.targetSelectionRadio);
        }
    };


    dialogUi.languageJapaneseRadio.onClick = function () { activateLanguageRadio(dialogUi, dialogUi.languageJapaneseRadio); };
    dialogUi.languageEnglishRadio.onClick = function () { activateLanguageRadio(dialogUi, dialogUi.languageEnglishRadio); };
    dialogUi.languageNoneRadio.onClick = function () { activateLanguageRadio(dialogUi, dialogUi.languageNoneRadio); };

    dialogUi.hyphenationCheckbox.onClick = function () {
        updateHyphenationControlsEnabled(dialogUi);
    };
    dialogUi.presetDropdown.onChange = function () {
        if (!dialogUi.presetDropdown.selection) return;
        applyPreset(dialogUi.presetDropdown.selection.text, dialogUi, presetFields);
    };
    dialogUi.presetExportButton.onClick = function () {
        exportPresetCode(presetFields, dialogUi);
    };
}

/**
 * 段落スタイルへ適用する設定を組み立てる
 * @returns {object} 適用する設定
 */
function buildStyleSettingsResult(dialogUi) {
    return {
        targetMode: dialogUi.targetAllRadio.value ? "all" : (dialogUi.targetSelectionRadio.value ? "specified" : "selectedParagraphs"),
        selectedStyleIndexes: dialogUi.selectedStyleIndexes,
        kinsokuIndex: dialogUi.kinsokuDropdown.selection.index,
        kinsokuTypeIndex: dialogUi.kinsokuTypeDropdown.selection.index,
        kinsokuHangTypeIndex: dialogUi.kinsokuHangTypeDropdown.selection.index,
        mojikumiIndex: dialogUi.mojikumiDropdown.selection.index,
        leadingModelIndex: dialogUi.leadingModelDropdown.selection.index,
        characterAlignmentIndex: dialogUi.characterAlignmentDropdown.selection.index,
        gridAlignmentIndex: dialogUi.gridAlignmentDropdown.selection.index,
        kerningMethodIndex: dialogUi.kerningMethodDropdown.selection.index,
        autoLeadingPercent: parseFloat(dialogUi.autoLeadingInput.text),
        composerIndex: dialogUi.composerDropdown.selection.index,
        hyphenation: dialogUi.hyphenationCheckbox.value,
        bunriKinshi: dialogUi.bunriKinshiCheckbox.value,
        rensuuji: dialogUi.rensuujiCheckbox.value,
        rotateSingleByte: dialogUi.rotateSingleByteCheckbox.value,
        absorbLineEndIdeographicSpace: dialogUi.absorbLineEndIdeographicSpaceCheckbox.value,
        latinWordBreak: dialogUi.latinWordBreakCheckbox.value,
        hyphenateWordsLongerThan: parseInt(dialogUi.hyphenateWordsLongerThanInput.text, 10),
        hyphenateAfterFirst: parseInt(dialogUi.hyphenateAfterFirstInput.text, 10),
        hyphenateBeforeLast: parseInt(dialogUi.hyphenateBeforeLastInput.text, 10),
        hyphenateLadderLimit: parseInt(dialogUi.hyphenateLadderLimitInput.text, 10),
        hyphenationZoneMm: parseFloat(dialogUi.hyphenationZoneInput.text),
        hyphenateCapitalizedWords: dialogUi.hyphenateCapitalizedWordsCheckbox.value,
        hyphenateAcrossColumns: dialogUi.hyphenateAcrossColumnsCheckbox.value,
        hyphenateLastWord: dialogUi.hyphenateLastWordCheckbox.value,
        ligatures: dialogUi.ligaturesCheckbox.value,
        language: getLanguageSelection(dialogUi)
    };
}

/**
 * 環境設定へ適用する設定を組み立てる
 * @returns {object} 適用する設定
 */
function buildAppPreferencesResult(dialogUi) {
    return {
        useSmartQuotes: dialogUi.useTypographersQuotesCheckbox.value,
        doubleQuotes: dialogUi.smartQuoteDropdown.selection
            ? dialogUi.smartQuoteDropdown.selection.text
            : null,
        singleQuotes: dialogUi.smartSingleQuoteDropdown.selection
            ? dialogUi.smartSingleQuoteDropdown.selection.text
            : null,
        textSizeUnit: dialogUi.textSizeUnitDropdown.selection
            ? dialogUi.textSizeUnitDropdown.selection.text
            : null,
        compositionUnit: dialogUi.compositionUnitDropdown.selection
            ? dialogUi.compositionUnitDropdown.selection.text
            : null
    };
}

/**
 * 段落スタイル用と環境設定用の結果をまとめる
 * @param {object} styleSettings 段落スタイル向けの設定
 * @param {object} appPreferences 環境設定向けの設定
 * @returns {object} まとめた設定
 */
function mergeDialogResults(styleSettingsResult, appPreferencesResult) {
    for (var appPreferenceKey in appPreferencesResult) {
        styleSettingsResult[appPreferenceKey] = appPreferencesResult[appPreferenceKey];
    }
    return styleSettingsResult;
}

/**
 * ダイアログの入力内容を結果オブジェクトにまとめる
 * @returns {object} 適用に使う設定
 */
function buildDialogResult(dialogUi) {
    return mergeDialogResults(
        buildStyleSettingsResult(dialogUi),
        buildAppPreferencesResult(dialogUi)
    );
}

/**
 * 文字組版設定ダイアログを表示する
 * @param {object} lookupTables 選択肢の参照表
 * @param {object} initialSettings 初期値
 * @returns {object|null} 設定内容。キャンセル時は null
 */
function showTypesettingSettingsDialog(kinsokuNames, kinsokuTypeNames, kinsokuHangTypeNames, mojikumiNames, leadingModelNames, characterAlignmentNames, gridAlignmentNames, kerningMethodNames, composerNames, paragraphStyleNames, defaultIndexes, lookupTables) {
    var dialogData = {
        kinsokuNames: kinsokuNames,
        kinsokuTypeNames: kinsokuTypeNames,
        kinsokuHangTypeNames: kinsokuHangTypeNames,
        mojikumiNames: mojikumiNames,
        leadingModelNames: leadingModelNames,
        characterAlignmentNames: characterAlignmentNames,
        gridAlignmentNames: gridAlignmentNames,
        kerningMethodNames: kerningMethodNames,
        composerNames: composerNames,
        paragraphStyleNames: paragraphStyleNames,
        defaultIndexes: defaultIndexes,
        lookupTables: lookupTables
    };

    var dialogUi = createDialogUI(dialogData);
    var presetFields = createPresetFields(dialogUi, dialogData);
    bindDialogEvents(dialogUi, dialogData, presetFields);
    updateHyphenationControlsEnabled(dialogUi);

    if (dialogUi.targetSelectedParagraphsRadio.value) {
        var selectedParagraph = getFirstParagraphFromSelection();

        if (selectedParagraph) {
            loadSettingsFromParagraph(selectedParagraph, dialogUi, dialogData);
        } else {
            activateTargetRadio(dialogUi, dialogUi.targetAllRadio);
        }
    }

    if (dialogUi.dialog.show() !== 1) return null;
    return buildDialogResult(dialogUi);
}

// =========================================
// オーバーライドの消去 / Clear overrides
// =========================================

/**
 * 選択範囲の文字オーバーライドを消去する
 * @returns {void}
 */
function clearTextOverridesInSelection() {
    var selectionItems = app.selection;
    if (!selectionItems || selectionItems.length === 0) return;

    for (var selectionIndex = 0; selectionIndex < selectionItems.length; selectionIndex++) {
        var selectedItem = selectionItems[selectionIndex];
        try { selectedItem.clearOverrides(OverrideType.ALL); } catch (clearItemOverridesError) { }
        try { selectedItem.texts[0].clearOverrides(OverrideType.ALL); } catch (clearTextOverridesError) { }
        try { selectedItem.paragraphs.everyItem().clearOverrides(OverrideType.ALL); } catch (clearParagraphOverridesError) { }
    }
}

/**
 * 選択があればオーバーライドを消去する
 * @returns {void}
 */
function clearOverridesIfActive() {
    clearTextOverridesInSelection();
    try { app.menuActions.itemByID(8489).invoke(); } catch (clearOverridesMenuActionError) { }
    try { app.redraw(); } catch (redrawError) { }
}

/**
 * 配列にその段落スタイルが含まれるかを判定する
 * @param {Array<ParagraphStyle>} styles 段落スタイルの配列
 * @param {ParagraphStyle} targetStyle 探す段落スタイル
 * @returns {boolean} 含まれていれば true
 */
function containsParagraphStyle(paragraphStyles, targetParagraphStyle) {
    for (var paragraphStyleIndex = 0; paragraphStyleIndex < paragraphStyles.length; paragraphStyleIndex++) {
        if (paragraphStyles[paragraphStyleIndex] === targetParagraphStyle) return true;
    }
    return false;
}

/**
 * 適用対象なら選択中の段落スタイルを追加する
 * @param {Array<ParagraphStyle>} styles 収集先の配列
 * @param {ParagraphStyle} paragraphStyle 追加する段落スタイル
 * @returns {void}
 */
function addSelectedParagraphStyleIfApplicable(resultStyles, allParagraphStyles, paragraphStyle) {
    if (!paragraphStyle || !paragraphStyle.isValid) return;
    if (containsParagraphStyle(resultStyles, paragraphStyle)) return;
    if (!containsParagraphStyle(allParagraphStyles, paragraphStyle)) return;
    resultStyles.push(paragraphStyle);
}

/**
 * 選択から対象の段落スタイルを集める
 * @returns {Array<ParagraphStyle>} 段落スタイルの配列
 */
function collectParagraphStylesFromSelection(allParagraphStyles) {
    var resultStyles = [];
    var selectionItems = app.selection;
    if (!selectionItems || selectionItems.length === 0) return resultStyles;

    for (var selectionIndex = 0; selectionIndex < selectionItems.length; selectionIndex++) {
        var selectedItem = selectionItems[selectionIndex];
        try {
            if (selectedItem.paragraphs && selectedItem.paragraphs.length > 0) {
                for (var paragraphIndex = 0; paragraphIndex < selectedItem.paragraphs.length; paragraphIndex++) {
                    addSelectedParagraphStyleIfApplicable(resultStyles, allParagraphStyles, selectedItem.paragraphs.item(paragraphIndex).appliedParagraphStyle);
                }
                continue;
            }
        } catch (eParagraphs) { }

        try {
            if (selectedItem.parentStory && selectedItem.parentStory.paragraphs && selectedItem.parentStory.paragraphs.length > 0) {
                for (var storyParagraphIndex = 0; storyParagraphIndex < selectedItem.parentStory.paragraphs.length; storyParagraphIndex++) {
                    addSelectedParagraphStyleIfApplicable(resultStyles, allParagraphStyles, selectedItem.parentStory.paragraphs.item(storyParagraphIndex).appliedParagraphStyle);
                }
            }
        } catch (eStory) { }
    }

    return resultStyles;
}

/**
 * 対象範囲の指定に応じて適用先の段落スタイルを求める
 * @param {string} targetMode 対象範囲を表す識別子
 * @param {Array<ParagraphStyle>} allStyles すべての段落スタイル
 * @param {Array<string>} specifiedNames 指定した段落スタイル名
 * @returns {Array<ParagraphStyle>} 適用先の段落スタイル
 */
function resolveTargetParagraphStyles(allParagraphStyles, dialogResult) {
    if (!dialogResult || dialogResult.targetMode === "all") {
        return allParagraphStyles;
    }

    if (dialogResult.targetMode === "specified") {
        var selectedStyles = [];
        var selectedIndexes = dialogResult.selectedStyleIndexes;
        if (!selectedIndexes || selectedIndexes.length === 0) return selectedStyles;

        for (var indexPosition = 0; indexPosition < selectedIndexes.length; indexPosition++) {
            var styleIndex = selectedIndexes[indexPosition];
            if (styleIndex >= 0 && styleIndex < allParagraphStyles.length) {
                selectedStyles.push(allParagraphStyles[styleIndex]);
            }
        }
        return selectedStyles;
    }

    if (dialogResult.targetMode === "selectedParagraphs") {
        return collectParagraphStylesFromSelection(allParagraphStyles);
    }

    return allParagraphStyles;
}

// =========================================
// 設定の適用 / Apply settings
// =========================================

/**
 * 辞書の引用符設定を適用する
 * @param {object} settings 適用する設定
 * @returns {void}
 */
function applyDictionaryQuoteSettings(dialogResult) {
    try {
        if (dialogResult.doubleQuotes) {
            app.languagesWithVendors.everyItem().doubleQuotes =
                dialogResult.doubleQuotes;
        }

        if (dialogResult.singleQuotes) {
            app.languagesWithVendors.everyItem().singleQuotes =
                dialogResult.singleQuotes;
        }
    } catch (dictionaryQuotesApplyError) { }
}

/**
 * 環境設定側の項目を適用する
 * @param {object} settings 適用する設定
 * @returns {void}
 */
function applyAppPreferenceSettings(dialogResult) {
    try {
        app.textPreferences.typographersQuotes =
            !!dialogResult.useSmartQuotes;
    } catch (typographersQuotesApplyError) { }

    applyDictionaryQuoteSettings(dialogResult);

    // テキストサイズの単位 / Text size measurement units
    try {
        if (dialogResult.textSizeUnit === "ポイント") app.viewPreferences.textSizeMeasurementUnits = TextSizeMeasurementUnits.POINTS;
        else if (dialogResult.textSizeUnit === "級") app.viewPreferences.textSizeMeasurementUnits = TextSizeMeasurementUnits.Q;
        else if (dialogResult.textSizeUnit === "アメリカ式ポイント") app.viewPreferences.textSizeMeasurementUnits = TextSizeMeasurementUnits.AMERICAN_POINTS;
    } catch (textSizeUnitApplyError) { }

    // 組版の単位 / Composition (vertical) measurement units
    try {
        if (dialogResult.compositionUnit === "ポイント") app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
        else if (dialogResult.compositionUnit === "歯") app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.HA;
        else if (dialogResult.compositionUnit === "U") app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.U;
        else if (dialogResult.compositionUnit === "倍") app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.BAI;
        else if (dialogResult.compositionUnit === "ミルス") app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILS;
        else if (dialogResult.compositionUnit === "アメリカ式ポイント") app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.AMERICAN_POINTS;
    } catch (compositionUnitApplyError) { }
}

/**
 * 確定した設定を対象の段落スタイルへ適用する
 * @param {Array<ParagraphStyle>} targetParagraphStyles 適用先の段落スタイル
 * @param {object} settings 適用する設定
 * @param {object} lookupTables 選択肢の参照表
 * @returns {void}
 */
function applyTypesettingSettingsToAll(targetParagraphStyles, dialogResult, lookupTables) {
    app.doScript(
        function () {
            var kinsokuTable = lookupTables.kinsokuTables[dialogResult.kinsokuIndex];
            var kinsokuTypeValue = lookupTables.kinsokuTypeValues[dialogResult.kinsokuTypeIndex];
            var kinsokuHangTypeValue = lookupTables.kinsokuHangTypeValues[dialogResult.kinsokuHangTypeIndex];
            var mojikumiTable = lookupTables.mojikumiTables[dialogResult.mojikumiIndex];
            var leadingModelValue = lookupTables.leadingModelValues[dialogResult.leadingModelIndex];
            var characterAlignmentValue = lookupTables.characterAlignmentValues[dialogResult.characterAlignmentIndex];
            var gridAlignmentValue = lookupTables.gridAlignmentValues[dialogResult.gridAlignmentIndex];
            var kerningMethodValue = lookupTables.kerningMethodValues[dialogResult.kerningMethodIndex];
            var autoLeadingPercent = dialogResult.autoLeadingPercent;
            var composerAliasList = lookupTables.composerAliases[dialogResult.composerIndex];
            var hyphenationValue = dialogResult.hyphenation;
            var bunriKinshiValue = dialogResult.bunriKinshi;

            var skipped = 0;
            var errorDetails = [];

            for (var styleIndex = 0; styleIndex < targetParagraphStyles.length; styleIndex++) {
                var paragraphStyle = targetParagraphStyles[styleIndex];
                try {
                    paragraphStyle.kinsokuSet = kinsokuTable;
                    paragraphStyle.kinsokuType = kinsokuTypeValue;
                    paragraphStyle.kinsokuHangType = kinsokuHangTypeValue;
                    paragraphStyle.mojikumi = mojikumiTable === null ? NothingEnum.NOTHING : mojikumiTable;
                    paragraphStyle.leadingModel = leadingModelValue;
                    paragraphStyle.characterAlignment = characterAlignmentValue;
                    paragraphStyle.gridAlignment = gridAlignmentValue;
                    paragraphStyle.kerningMethod = kerningMethodValue;
                    if (!isNaN(dialogResult.autoLeadingPercent)) {
                        paragraphStyle.autoLeading = dialogResult.autoLeadingPercent;
                    }
                    paragraphStyle.hyphenation = hyphenationValue;
                    paragraphStyle.bunriKinshi = bunriKinshiValue;
                } catch (applyError) {
                    skipped++;
                    var styleNameForError = "(unknown)";
                    try { styleNameForError = paragraphStyle.name; } catch (eName) { }
                    errorDetails.push("[" + styleNameForError + "] " + applyError);
                }

                // composer はロケールやバージョンで受理名が異なるため alias を順に試行 /
                // Composer name varies by locale/version; try aliases in order
                applyComposerAliases(paragraphStyle, composerAliasList);

                // プロパティ名が不確実なものは安全代入で適用 / Apply uncertain properties via safe assignment
                safeSetProperty(paragraphStyle, "rensuuji", dialogResult.rensuuji);
                safeSetProperty(paragraphStyle, "rotateSingleByteCharacters", dialogResult.rotateSingleByte);
                safeSetProperty(paragraphStyle, "treatIdeographicSpaceAsSpace", dialogResult.absorbLineEndIdeographicSpace);
                // 欧文泣き別れ / Latin word break (allowArbitraryHyphenation)
                safeSetProperty(paragraphStyle, "allowArbitraryHyphenation", dialogResult.latinWordBreak);

                // ハイフネーション詳細設定 / Hyphenation detail settings
                if (!isNaN(dialogResult.hyphenateWordsLongerThan)) {
                    safeSetProperty(paragraphStyle, "hyphenateWordsLongerThan", dialogResult.hyphenateWordsLongerThan);
                }
                if (!isNaN(dialogResult.hyphenateAfterFirst)) {
                    safeSetProperty(paragraphStyle, "hyphenateAfterFirst", dialogResult.hyphenateAfterFirst);
                }
                if (!isNaN(dialogResult.hyphenateBeforeLast)) {
                    safeSetProperty(paragraphStyle, "hyphenateBeforeLast", dialogResult.hyphenateBeforeLast);
                }
                if (!isNaN(dialogResult.hyphenateLadderLimit)) {
                    safeSetProperty(paragraphStyle, "hyphenateLadderLimit", dialogResult.hyphenateLadderLimit);
                }
                if (!isNaN(dialogResult.hyphenationZoneMm)) {
                    safeSetProperty(paragraphStyle, "hyphenationZone", dialogResult.hyphenationZoneMm + "mm");
                }

                safeSetProperty(paragraphStyle, "hyphenateCapitalizedWords", dialogResult.hyphenateCapitalizedWords);
                safeSetProperty(paragraphStyle, "hyphenateAcrossColumns", dialogResult.hyphenateAcrossColumns);
                safeSetProperty(paragraphStyle, "hyphenateLastWord", dialogResult.hyphenateLastWord);

                safeSetProperty(paragraphStyle, "ligatures", dialogResult.ligatures);

                if (dialogResult.language && LANGUAGE_CANDIDATES[dialogResult.language]) {
                    var languageCandidates = LANGUAGE_CANDIDATES[dialogResult.language];
                    for (var languageCandidateIndex = 0; languageCandidateIndex < languageCandidates.length; languageCandidateIndex++) {
                        try {
                            var candidateLanguage = app.languagesWithVendors.itemByName(languageCandidates[languageCandidateIndex]);
                            if (candidateLanguage.isValid) {
                                paragraphStyle.appliedLanguage = candidateLanguage;
                                break;
                            }
                        } catch (eLangApply) { }
                    }
                }
            }

            applyAppPreferenceSettings(dialogResult);

            if (skipped > 0) {
                alert(getLabel("alert.partialFailurePrefix") + skipped + getLabel("alert.partialFailureSuffix") + errorDetails.join("\n"));
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
    var kinsokuHangTypeOptions = createKinsokuHangTypeOptions();
    var mojikumiTableData = collectMojikumiTables(activeDocument);
    var targetParagraphStyleData = collectTargetParagraphStyles(activeDocument);
    var targetParagraphStyles = targetParagraphStyleData.styles;

    if (targetParagraphStyles.length === 0) {
        alert(getLabel("alert.noParagraphStyles"));
        return;
    }

    var leadingModelOptions = createLeadingModelOptions();
    var characterAlignmentOptions = createCharacterAlignmentOptions();
    var gridAlignmentOptions = createGridAlignmentOptions();
    var kerningMethodOptions = createKerningMethodOptions();
    var composerOptions = createComposerOptions();
    var defaultPreset = PRESETS["InDesignのデフォルト"].styleSettings;
    var defaultIndexes = {
        kinsokuIndex: getDefaultIndexByName(kinsokuTableData.names, defaultPreset.kinsoku),
        kinsokuTypeIndex: getDefaultIndexByName(kinsokuTypeOptions.names, defaultPreset.kinsokuType),
        kinsokuHangTypeIndex: getDefaultIndexByName(kinsokuHangTypeOptions.names, defaultPreset.kinsokuHangType),
        mojikumiIndex: getDefaultIndexByName(mojikumiTableData.names, defaultPreset.mojikumi),
        leadingModelIndex: getDefaultIndexByName(leadingModelOptions.names, defaultPreset.leadingModel),
        characterAlignmentIndex: getDefaultIndexByName(characterAlignmentOptions.names, defaultPreset.characterAlignment),
        gridAlignmentIndex: getDefaultIndexByName(gridAlignmentOptions.names, defaultPreset.gridAlignment),
        kerningMethodIndex: getDefaultIndexByName(kerningMethodOptions.names, defaultPreset.kerningMethod),
        autoLeadingPercent: defaultPreset.autoLeading,
        composerIndex: (function () {
            var aliasMatch = findIndexByComposerAliases(composerOptions.aliases, defaultPreset.composer);
            return aliasMatch >= 0 ? aliasMatch : getDefaultIndexByName(composerOptions.names, defaultPreset.composer);
        })(),
        hyphenation: defaultPreset.hyphenation,
        bunriKinshi: defaultPreset.bunriKinshi,
        rensuuji: defaultPreset.rensuuji,
        rotateSingleByte: defaultPreset.rotateSingleByte,
        absorbLineEndIdeographicSpace: defaultPreset.absorbLineEndIdeographicSpace,
        latinWordBreak: defaultPreset.latinWordBreak,
        hyphenateWordsLongerThan: defaultPreset.hyphenateWordsLongerThan,
        hyphenateAfterFirst: defaultPreset.hyphenateAfterFirst,
        hyphenateBeforeLast: defaultPreset.hyphenateBeforeLast,
        hyphenateLadderLimit: defaultPreset.hyphenateLadderLimit,
        hyphenationZoneMm: defaultPreset.hyphenationZone,
        hyphenateCapitalizedWords: defaultPreset.hyphenateCapitalizedWords,
        hyphenateAcrossColumns: defaultPreset.hyphenateAcrossColumns,
        hyphenateLastWord: defaultPreset.hyphenateLastWord,
        ligatures: defaultPreset.ligatures,
        language: defaultPreset.language
    };

    var lookupTables = {
        kinsokuTables: kinsokuTableData.tables,
        kinsokuTypeValues: kinsokuTypeOptions.values,
        kinsokuHangTypeValues: kinsokuHangTypeOptions.values,
        mojikumiTables: mojikumiTableData.tables,
        leadingModelValues: leadingModelOptions.values,
        characterAlignmentValues: characterAlignmentOptions.values,
        gridAlignmentValues: gridAlignmentOptions.values,
        kerningMethodValues: kerningMethodOptions.values,
        composerAliases: composerOptions.aliases
    };

    var dialogResult = showTypesettingSettingsDialog(
        kinsokuTableData.names,
        kinsokuTypeOptions.names,
        kinsokuHangTypeOptions.names,
        mojikumiTableData.names,
        leadingModelOptions.names,
        characterAlignmentOptions.names,
        gridAlignmentOptions.names,
        kerningMethodOptions.names,
        composerOptions.names,
        targetParagraphStyleData.names,
        defaultIndexes,
        lookupTables
    );
    if (dialogResult === null) return;

    var resolvedTargetParagraphStyles = resolveTargetParagraphStyles(targetParagraphStyles, dialogResult);
    if (resolvedTargetParagraphStyles.length === 0) {
        alert(getLabel("alert.noTargetStyles"));
        return;
    }

    applyTypesettingSettingsToAll(resolvedTargetParagraphStyles, dialogResult, lookupTables);

    clearOverridesIfActive();

})();