#target indesign

/*
概要

TypesettingStyleManager.jsx
--------------------------------------------------------
アクティブドキュメントの段落スタイルに対して、
文字組版・基本文字設定・ハイフネーション設定・欧文合字（ligatures）を一括で適用する管理ツール。

対象設定:
・自動カーニング
・自動行送り
・文字揃え
・行送りの基準位置
・グリッド揃え
・コンポーザー
・欧文合字
・スマート引用符
・言語
・禁則処理セット
・禁則調整方式
・ぶら下がり方法
・文字組みアキ量設定
・ハイフネーションおよび詳細設定
・分離禁止、連数字処理、縦組み中の文字回転、行末全角スペース吸収など

主な機能:
・ダイアログで各設定を選び、OK 時に対象段落スタイルへ反映
・対象を「選択中」「すべて」「指定」から選択
・選択中の段落から現在の組版設定・言語・ハイフネーション設定を読み込み、初期値として利用
・段落スタイルグループを再帰的に走査し、対象外グループを除外
・ハイフネーションの ON/OFF に応じて関連項目を有効／無効化
・プリセット（欧文組版／グリッド優先／グリッド無視／ソースコード／InDesignのデフォルト）の適用と、現在の設定（言語・引用符を含む）をプリセットコードとして書き出し
・対象段落スタイル数を選択ダイアログ内に表示
・適用後、選択範囲のオーバーライドを常に消去

デフォルト値:
・プリセット「InDesignのデフォルト」で初期値を管理
・ダイアログの初期選択にも使用

除外条件:
・[段落スタイルなし] / [No Paragraph Style]
・[基本段落] / [Basic Paragraph]
・名前が "_" で始まるスタイルグループ配下の段落スタイル

補足:
・欧文泣き別れは UI 項目として保持しているが、対応する ParagraphStyle プロパティが未確認のため未適用
・スマート引用符は app.textPreferences.useSmartQuotes を使用
・言語は app.languagesWithVendors から候補名を順に探索して適用
・文字組みアキ量設定は、組み込みプリセット enum とカスタム MojikumiTable の両方を可能な範囲で読み取る
・禁則設定は、選択段落から読めない場合に適用段落スタイル側も参照する
・コンポーザーは「基本設定」パネル内で設定
・一部のプリセット項目（グリッド優先、グリッド無視、ソースコード）は将来拡張用の空定義として保持
・設定適用時は lookupTables にまとめた enum / table 値を参照して段落スタイルへ反映

ハイフネーション設定：コンさん
https://typesetterkon.blogspot.com/2011/06/indesign5.html

*/

var SCRIPT_VERSION = "v1.0.0";

// =========================================
// 設定 / Settings
// =========================================

var DIALOG_TITLE = "文字組版設定（一括）";

// 言語の候補（先頭から順に試行） / Language candidates (tried in order)
var LANGUAGE_CANDIDATES = {
    "ja": ["日本語", "Japanese"],
    "en": ["英語：米国", "English: USA"],
    "none": ["[言語なし]", "[No Language]"]
};

// ドロップダウン幅 / Dropdown widths
var W_DROP = 180;

// パネルの共通マージン / Shared panel margins
var PANEL_MARGINS = [15, 20, 15, 10];

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

/* パネルの基本レイアウトを設定 / Set shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = "left";
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
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

/* ぶら下がり方法の候補を定義 / Define kinsoku hang type options */
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

/* 行送りの基準位置の候補を定義 / Define leading model options */
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

/* 自動カーニングの候補を定義 / Define kerning method options */
function createKerningMethodOptions() {
    var values = ["メトリクス", "オプティカル", "和文等幅", "0"];
    return { names: values, values: values };
}

/* グリッド揃えの候補を定義 / Define grid alignment options */
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

/* 文字揃えの候補を定義 / Define character alignment options */
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

/* コンポーザーの候補を定義 / Define composer options
   aliases は読み込み判定と適用試行に使う候補名（先頭から順に試す） /
   aliases are candidate names used for both detection and assignment fallback */
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

/* aliases 配列群から target に一致するインデックスを返す / Find index whose alias list contains target */
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

/* aliases から段落スタイルへ composer を順に試行 / Try assigning composer aliases in order */
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
// ダイアログ UI / Dialog UI
// =========================================

/* ラベル付きドロップダウンを 1 行追加 / Add a labeled dropdown row */
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

/* ラベル付き数値入力を 1 行追加 / Add a labeled numeric input row */
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

/* 段落スタイル選択ピッカーを表示 / Show the paragraph style picker dialog */
function showParagraphStylePicker(paragraphStyleNames, currentSelectedIndexes) {
    var picker = new Window("dialog", "段落スタイルの選択");
    picker.orientation = "column";
    picker.alignChildren = "fill";
    picker.margins = [15, 12, 15, 16];
    picker.spacing = 10;

    var targetCountText = picker.add("statictext", undefined, "対象: " + paragraphStyleNames.length + " 件");
    targetCountText.alignment = "left";

    var listPanel = picker.add("panel");
    listPanel.orientation = "column";
    listPanel.alignChildren = "left";
    listPanel.margins = [10, 10, 10, 10];
    listPanel.spacing = 4;

    var checkboxes = [];
    function isCurrentlySelected(index) {
        if (!currentSelectedIndexes) return true;
        for (var selectedIndexPosition = 0; selectedIndexPosition < currentSelectedIndexes.length; selectedIndexPosition++) {
            if (currentSelectedIndexes[selectedIndexPosition] === index) return true;
        }
        return false;
    }

    function bindAltClickToggle(checkbox) {
        checkbox.onClick = function () {
            if (ScriptUI.environment.keyboardState.altKey) {
                var newValue = checkbox.value;
                for (var checkboxIndex = 0; checkboxIndex < checkboxes.length; checkboxIndex++) {
                    checkboxes[checkboxIndex].value = newValue;
                }
            }
        };
    }

    for (var nameIndex = 0; nameIndex < paragraphStyleNames.length; nameIndex++) {
        var checkbox = listPanel.add("checkbox", undefined, paragraphStyleNames[nameIndex]);
        checkbox.value = isCurrentlySelected(nameIndex);
        bindAltClickToggle(checkbox);
        checkboxes.push(checkbox);
    }

    var hint = picker.add("statictext", undefined, "Option（Alt）+ クリックで全選択／全解除を切り換え");
    hint.alignment = "left";

    var buttonGroup = picker.add("group");
    buttonGroup.alignment = "right";
    buttonGroup.add("button", undefined, "キャンセル", { name: "cancel" });
    buttonGroup.add("button", undefined, "OK", { name: "ok" });

    if (picker.show() !== 1) return null;

    var selectedIndexes = [];
    for (var checkboxIndex = 0; checkboxIndex < checkboxes.length; checkboxIndex++) {
        if (checkboxes[checkboxIndex].value) selectedIndexes.push(checkboxIndex);
    }
    return selectedIndexes;
}

/* 表示名からインデックスを取得 / Find index by display name */
function findIndexByDisplayName(names, target) {
    for (var nameIndex = 0; nameIndex < names.length; nameIndex++) {
        if (names[nameIndex] === target) return nameIndex;
    }
    return -1;
}

/* enum 値からインデックスを取得 / Find index by enum value */
function findIndexByEnumValue(values, target) {
    for (var valueIndex = 0; valueIndex < values.length; valueIndex++) {
        if (values[valueIndex] === target) return valueIndex;
    }
    return -1;
}

/* enum 値をドロップダウンに反映 / Safely assign enum value to dropdown */
function safeAssignDropdownFromEnum(dropdown, values, target) {
    var foundIndex = findIndexByEnumValue(values, target);
    if (foundIndex < 0) return false;
    dropdown.selection = foundIndex;
    return true;
}

/* 表示名をドロップダウンに反映 / Safely assign display name to dropdown */
function safeAssignDropdownFromName(dropdown, names, targetName) {
    if (!targetName) return false;
    var foundIndex = findIndexByDisplayName(names, targetName);
    if (foundIndex < 0) return false;
    dropdown.selection = foundIndex;
    return true;
}

/* 真偽値をチェックボックスに反映 / Safely assign boolean value to checkbox */
function safeAssignCheckbox(checkbox, value) {
    checkbox.value = !!value;
}

/* 数値を入力欄に反映 / Safely assign number value to text input */
function safeAssignNumberInput(input, value) {
    if (typeof value === "number") input.text = String(value);
}

/* プロパティ代入を安全に試行 / Safely try assigning a property */
function safeSetProperty(targetObject, propertyName, value) {
    try {
        targetObject[propertyName] = value;
        return true;
    } catch (propertySetError) { }
    return false;
}

/* プリセット定義（一部は将来拡張用の空定義） / Preset definitions; some entries are placeholders for future expansion */

var PRESETS = {
    "欧文組版": {
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
        gridAlignment: "欧文ベースライン",
        composer: "欧文段落コンポーザー",
        hyphenation: true,
        hyphenateWordsLongerThan: 6,
        hyphenateAfterFirst: 3,
        hyphenateBeforeLast: 3,
        hyphenateLadderLimit: 2,
        hyphenationZone: 1.25,
        hyphenateCapitalizedWords: false,
        hyphenateAcrossColumns: false,
        hyphenateLastWord: false,
        useSmartQuotes: true,
        ligatures: true,
        language: "en"
    },
    "グリッド優先": {
        kinsoku: "弱い禁則",
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
        useSmartQuotes: true,
        ligatures: true,
        language: "ja"
    },
    "グリッド無視": {
        kinsoku: "弱い禁則",
        kinsokuType: "調整量を優先",
        kinsokuHangType: "なし",
        bunriKinshi: true,
        mojikumi: "行末約物半角",
        leadingModel: "欧文ベースライン",
        rensuuji: true,
        rotateSingleByte: false,
        absorbLineEndIdeographicSpace: true,
        latinWordBreak: false,
        kerningMethod: "メトリクス",
        autoLeading: 175,
        characterAlignment: "欧文ベースライン",
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
        useSmartQuotes: true,
        ligatures: true,
        language: "ja"
    },
    "ソースコード": {
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
        useSmartQuotes: false,
        ligatures: false,
        language: "none"
    },
    "InDesignのデフォルト": {
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
        useSmartQuotes: false,
        ligatures: true,
        language: "ja"
    }
};

/* プリセット対象フィールド定義を作成 / Create preset field definitions */
function createPresetFields(dialogUi, dialogData) {
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
        /* 欧文合字も通常フィールドとしてプリセットに書き出す / Export ligatures as a regular preset field */
        { key: "ligatures", type: "cb", control: dialogUi.ligaturesCheckbox }
    ];
}

/* ハイフネーション関連 UI の有効状態を更新 / Update hyphenation control enabled states */
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

/* 対象ラジオボタンを排他的に切り替え / Activate one target radio exclusively */
function activateTargetRadio(dialogUi, activeRadio) {
    var targetRadios = [dialogUi.targetAllRadio, dialogUi.targetSelectionRadio, dialogUi.targetSelectedParagraphsRadio];
    for (var radioIndex = 0; radioIndex < targetRadios.length; radioIndex++) {
        targetRadios[radioIndex].value = (targetRadios[radioIndex] === activeRadio);
    }
}

/* 言語ラジオボタンを排他的に切り替え / Activate one language radio exclusively */
function activateLanguageRadio(dialogUi, activeRadio) {
    var languageRadios = [dialogUi.languageJapaneseRadio, dialogUi.languageEnglishRadio, dialogUi.languageNoneRadio];
    for (var radioIndex = 0; radioIndex < languageRadios.length; radioIndex++) {
        languageRadios[radioIndex].value = (languageRadios[radioIndex] === activeRadio);
    }
}

/* 言語ラジオボタンの選択値を取得 / Get selected language key */
function getLanguageSelection(dialogUi) {
    if (dialogUi.languageEnglishRadio.value) return "en";
    if (dialogUi.languageNoneRadio.value) return "none";
    return "ja";
}

/* 選択から最初の段落を取得 / Get first paragraph from the current selection */
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

/* 文字組み値から表示名を解決 / Resolve display name from a mojikumi value */
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

/* 適用言語から言語キーを解決 / Resolve language key from an applied language value */
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

/* 禁則処理セット値からドロップダウンインデックスを解決 / Resolve dropdown index from a kinsoku set value */
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

/* 禁則設定を選択段落から読み込む / Load kinsoku settings from a paragraph */
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

/* ハイフネーション設定を選択段落から読み込む / Load hyphenation settings from a paragraph */
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

/* 選択段落から設定を読み込む / Load settings from a paragraph */
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

/* プリセットを UI に反映 / Apply preset values to UI controls */
function applyPreset(presetName, dialogUi, presetFields) {
    var preset = PRESETS[presetName];
    if (!preset) return;

    for (var fieldIndex = 0; fieldIndex < presetFields.length; fieldIndex++) {
        var presetField = presetFields[fieldIndex];
        if (preset[presetField.key] === undefined) continue;
        if (presetField.type === "dd") {
            safeAssignDropdownFromName(presetField.control, presetField.names, preset[presetField.key]);
        } else if (presetField.type === "cb") {
            presetField.control.value = !!preset[presetField.key];
        } else if (presetField.type === "in") {
            presetField.control.text = String(preset[presetField.key]);
        }
    }

    if (preset.useSmartQuotes !== undefined) {
        dialogUi.smartQuoteRadio.value = !!preset.useSmartQuotes;
        dialogUi.straightQuoteRadio.value = !preset.useSmartQuotes;
    }
    if (preset.language !== undefined) {
        if (preset.language === "en") activateLanguageRadio(dialogUi, dialogUi.languageEnglishRadio);
        else if (preset.language === "none") activateLanguageRadio(dialogUi, dialogUi.languageNoneRadio);
        else activateLanguageRadio(dialogUi, dialogUi.languageJapaneseRadio);
    }
    updateHyphenationControlsEnabled(dialogUi);
}

/* プリセット名入力ダイアログを表示 / Show preset name input dialog */
function showPresetNameInputDialog() {
    var nameDialog = new Window("dialog", "プリセット名の入力");
    nameDialog.orientation = "column";
    nameDialog.alignChildren = "fill";
    nameDialog.margins = [15, 12, 15, 16];
    nameDialog.spacing = 10;

    nameDialog.add("statictext", undefined, "プリセット名（書き出しファイル名にも使用）：");
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

/* プリセットコードを生成 / Build preset code snippet */
function buildPresetCodeSnippet(presetName, presetFields, dialogUi) {
    var lines = [];
    lines.push("// PRESETS マップに以下を追加してください（プリセットドロップダウン項目への追加もお忘れなく）");
    lines.push("PRESETS[\"" + presetName + "\"] = {");
    for (var fieldIndex = 0; fieldIndex < presetFields.length; fieldIndex++) {
        var presetField = presetFields[fieldIndex];
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
        lines.push("    " + presetField.key + ": " + valueText + ",");
    }
    lines.push("    useSmartQuotes: " + (dialogUi.smartQuoteRadio.value ? "true" : "false") + ",");
    lines.push("    language: \"" + getLanguageSelection(dialogUi) + "\"");
    lines.push("};");
    return lines.join("\n");
}

/* プリセットコードを書き出す / Export preset code */
function exportPresetCode(presetFields, dialogUi) {
    var presetName = showPresetNameInputDialog();
    if (!presetName) return;

    var code = buildPresetCodeSnippet(presetName, presetFields, dialogUi);

    try {
        var file = File(Folder.desktop + "/" + encodeURI(presetName) + ".jsx");
        if (file.exists) {
            if (!confirm("「" + presetName + ".jsx」は既にデスクトップに存在します。上書きしますか？")) return;
        }
        file.encoding = "UTF-8";
        if (file.open("w")) {
            file.write(code);
            file.close();
            alert("プリセット「" + presetName + "」をデスクトップに書き出しました。");
        } else {
            alert("ファイルを開けませんでした。");
        }
    } catch (eExport) {
        alert("書き出しエラー: " + eExport.message);
    }
}

/* ダイアログ UI を作成 / Create dialog UI */
function createDialogUI(dialogData) {
    var defaultIndexes = dialogData.defaultIndexes;
    var dialog = new Window("dialog", DIALOG_TITLE + " " + SCRIPT_VERSION);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.margins = [15, 12, 15, 16];
    dialog.spacing = 10;

    var topColumnsGroup = dialog.add("group");
    topColumnsGroup.orientation = "row";
    topColumnsGroup.alignChildren = ["fill", "top"];
    topColumnsGroup.spacing = 10;

    var targetPanel = topColumnsGroup.add("panel", undefined, "対象の段落スタイル");
    setupPanel(targetPanel, 8);
    targetPanel.orientation = "row";
    targetPanel.alignChildren = ["left", "center"];

    var targetSelectedParagraphsRadio = targetPanel.add("radiobutton", undefined, "選択中");
    var targetAllRadio = targetPanel.add("radiobutton", undefined, "すべて");
    var targetSelectionRadio = targetPanel.add("radiobutton", undefined, "指定");
    var targetSelectionButton = targetPanel.add("button", undefined, "選択");
    targetSelectedParagraphsRadio.value = true;

    var presetPanel = topColumnsGroup.add("panel", undefined, "プリセット");
    setupPanel(presetPanel, 8);
    presetPanel.orientation = "row";

    var presetDropdown = presetPanel.add("dropdownlist", undefined, ["欧文組版", "グリッド優先", "グリッド無視", "ソースコード", "InDesignのデフォルト"]);
    presetDropdown.selection = null;
    presetDropdown.preferredSize.width = W_DROP;
    var presetExportButton = presetPanel.add("button", undefined, "書き出し");

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

    var compositionExtraPanel = leftColumn.add("panel", undefined, "基本設定");
    setupPanel(compositionExtraPanel, 8);

    var kerningMethodDropdown = addDropdownRow(compositionExtraPanel, "自動カーニング：", dialogData.kerningMethodNames, defaultIndexes.kerningMethodIndex);
    var autoLeadingInput = addNumberRow(compositionExtraPanel, "自動行送り：", defaultIndexes.autoLeadingPercent, "%");
    var characterAlignmentDropdown = addDropdownRow(compositionExtraPanel, "文字揃え：", dialogData.characterAlignmentNames, defaultIndexes.characterAlignmentIndex);
    var leadingModelDropdown = addDropdownRow(compositionExtraPanel, "行送りの基準位置：", dialogData.leadingModelNames, defaultIndexes.leadingModelIndex);
    var gridAlignmentDropdown = addDropdownRow(compositionExtraPanel, "グリッド揃え：", dialogData.gridAlignmentNames, defaultIndexes.gridAlignmentIndex);
    var composerDropdown = addDropdownRow(compositionExtraPanel, "コンポーザー：", dialogData.composerNames, defaultIndexes.composerIndex);

    var ligaturesCheckbox = compositionExtraPanel.add("checkbox", undefined, "欧文合字");
    ligaturesCheckbox.value = defaultIndexes.ligatures;

    var smartQuotesInitial;
    try {
        smartQuotesInitial = !!app.textPreferences.useSmartQuotes;
    } catch (smartQuotesReadError) {
        smartQuotesInitial = true;
    }

    var smartQuoteRow = compositionExtraPanel.add("group");
    smartQuoteRow.orientation = "row";
    smartQuoteRow.alignChildren = ["left", "center"];
    smartQuoteRow.spacing = 8;

    var smartQuoteLabel = smartQuoteRow.add("statictext", undefined, "引用符：");
    smartQuoteLabel.preferredSize.width = 120;

    var smartQuoteRadio = smartQuoteRow.add("radiobutton", undefined, "“”");
    var straightQuoteRadio = smartQuoteRow.add("radiobutton", undefined, "\"\"");
    smartQuoteRadio.value = smartQuotesInitial;
    straightQuoteRadio.value = !smartQuotesInitial;

    var languageRow = compositionExtraPanel.add("group");
    languageRow.orientation = "row";
    languageRow.alignChildren = ["left", "center"];
    languageRow.spacing = 8;

    var languageLabel = languageRow.add("statictext", undefined, "言語：");
    languageLabel.preferredSize.width = 120;

    var languageJapaneseRadio = languageRow.add("radiobutton", undefined, "日本語");
    var languageEnglishRadio = languageRow.add("radiobutton", undefined, "英語");
    var languageNoneRadio = languageRow.add("radiobutton", undefined, "なし");
    languageJapaneseRadio.value = (defaultIndexes.language !== "en" && defaultIndexes.language !== "none");
    languageEnglishRadio.value = (defaultIndexes.language === "en");
    languageNoneRadio.value = (defaultIndexes.language === "none");

    var compositionPanel = leftColumn.add("panel", undefined, "日本語文字組版");
    setupPanel(compositionPanel, 8);

    var kinsokuDropdown = addDropdownRow(compositionPanel, "禁則処理セット：", dialogData.kinsokuNames, defaultIndexes.kinsokuIndex);
    var kinsokuTypeDropdown = addDropdownRow(compositionPanel, "禁則調整方式：", dialogData.kinsokuTypeNames, defaultIndexes.kinsokuTypeIndex);
    var kinsokuHangTypeDropdown = addDropdownRow(compositionPanel, "ぶら下がり方法：", dialogData.kinsokuHangTypeNames, defaultIndexes.kinsokuHangTypeIndex);

    var bunriKinshiCheckbox = compositionPanel.add("checkbox", undefined, "分離禁止処理");
    bunriKinshiCheckbox.value = defaultIndexes.bunriKinshi;

    var mojikumiDropdown = addDropdownRow(compositionPanel, "文字組みアキ量：", dialogData.mojikumiNames, defaultIndexes.mojikumiIndex);

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

    var rensuujiCheckbox = compositionCheckboxesLeft.add("checkbox", undefined, "連数字処理");
    rensuujiCheckbox.value = defaultIndexes.rensuuji;
    var rotateSingleByteCheckbox = compositionCheckboxesLeft.add("checkbox", undefined, "縦組み中の文字回転");
    rotateSingleByteCheckbox.value = defaultIndexes.rotateSingleByte;
    var absorbLineEndIdeographicSpaceCheckbox = compositionCheckboxesRight.add("checkbox", undefined, "全角スペースを行末吸収");
    absorbLineEndIdeographicSpaceCheckbox.value = defaultIndexes.absorbLineEndIdeographicSpace;
    var latinWordBreakCheckbox = compositionCheckboxesRight.add("checkbox", undefined, "欧文泣き別れ");
    latinWordBreakCheckbox.value = defaultIndexes.latinWordBreak;


    var hyphenationPanel = rightColumn.add("panel", undefined, "ハイフネーション");
    setupPanel(hyphenationPanel, 8);
    var hyphenationCheckbox = hyphenationPanel.add("checkbox", undefined, "ハイフネーション");
    hyphenationCheckbox.value = defaultIndexes.hyphenation;
    var hyphenateWordsLongerThanInput = addNumberRow(hyphenationPanel, "単語の最初文字数：", defaultIndexes.hyphenateWordsLongerThan, "文字");
    var hyphenateAfterFirstInput = addNumberRow(hyphenationPanel, "先頭の後：", defaultIndexes.hyphenateAfterFirst, "文字");
    var hyphenateBeforeLastInput = addNumberRow(hyphenationPanel, "最後の前：", defaultIndexes.hyphenateBeforeLast, "文字");
    var hyphenateLadderLimitInput = addNumberRow(hyphenationPanel, "最大のハイフン数：", defaultIndexes.hyphenateLadderLimit, "ハイフン");
    var hyphenationZoneInput = addNumberRow(hyphenationPanel, "領域：", defaultIndexes.hyphenationZoneMm, "mm");

    var hyphenationCheckboxesGroup = hyphenationPanel.add("group");
    hyphenationCheckboxesGroup.orientation = "column";
    hyphenationCheckboxesGroup.alignChildren = "left";
    hyphenationCheckboxesGroup.spacing = 4;
    hyphenationCheckboxesGroup.margins = [0, 10, 0, 0];
    var hyphenateCapitalizedWordsCheckbox = hyphenationCheckboxesGroup.add("checkbox", undefined, "大文字の単語");
    hyphenateCapitalizedWordsCheckbox.value = defaultIndexes.hyphenateCapitalizedWords;
    var hyphenateAcrossColumnsCheckbox = hyphenationCheckboxesGroup.add("checkbox", undefined, "段間、フレームにわたる単語");
    hyphenateAcrossColumnsCheckbox.value = defaultIndexes.hyphenateAcrossColumns;
    var hyphenateLastWordCheckbox = hyphenationCheckboxesGroup.add("checkbox", undefined, "段落末尾の単語");
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
        smartQuoteRadio: smartQuoteRadio,
        straightQuoteRadio: straightQuoteRadio,
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
        selectedStyleIndexes: null
    };
}

/* ダイアログイベントを接続 / Bind dialog events */
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

    dialogUi.smartQuoteRadio.onClick = function () {
        dialogUi.smartQuoteRadio.value = true;
        dialogUi.straightQuoteRadio.value = false;
    };
    dialogUi.straightQuoteRadio.onClick = function () {
        dialogUi.straightQuoteRadio.value = true;
        dialogUi.smartQuoteRadio.value = false;
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

/* ダイアログの入力値を結果オブジェクトにする / Build dialog result object */
function buildDialogResult(dialogUi) {
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
        useSmartQuotes: dialogUi.smartQuoteRadio.value,
        language: getLanguageSelection(dialogUi)
    };
}

/* ダイアログを表示して結果を返す / Show the dialog and return the result */
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
        loadSettingsFromParagraph(getFirstParagraphFromSelection(), dialogUi, dialogData);
    }

    if (dialogUi.dialog.show() !== 1) return null;
    return buildDialogResult(dialogUi);
}

// =========================================
// オーバーライドの消去 / Clear overrides
// =========================================

/* 選択範囲の段落オーバーライドを消去 / Clear paragraph overrides on the current selection */
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

/* 選択範囲＋メニューアクションでオーバーライドを消去 / Clear overrides via selection clear + menu action */
function clearOverridesIfActive() {
    clearTextOverridesInSelection();
    try { app.menuActions.itemByID(8489).invoke(); } catch (clearOverridesMenuActionError) { }
    try { app.redraw(); } catch (redrawError) { }
}

/* 配列内に同一段落スタイルがあるか判定 / Check whether the same paragraph style already exists in an array */
function containsParagraphStyle(paragraphStyles, targetParagraphStyle) {
    for (var paragraphStyleIndex = 0; paragraphStyleIndex < paragraphStyles.length; paragraphStyleIndex++) {
        if (paragraphStyles[paragraphStyleIndex] === targetParagraphStyle) return true;
    }
    return false;
}

/* 適用対象候補に含まれる段落スタイルだけを追加 / Add paragraph style only when it is included in applicable targets */
function addSelectedParagraphStyleIfApplicable(resultStyles, allParagraphStyles, paragraphStyle) {
    if (!paragraphStyle || !paragraphStyle.isValid) return;
    if (containsParagraphStyle(resultStyles, paragraphStyle)) return;
    if (!containsParagraphStyle(allParagraphStyles, paragraphStyle)) return;
    resultStyles.push(paragraphStyle);
}

/* 選択範囲で使用されている段落スタイルを取得 / Collect paragraph styles used in the current selection */
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

/* 指定インデックスから適用対象の段落スタイルを抽出 / Resolve target paragraph styles from selected indexes */
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

/* ダイアログの設定を対象段落スタイルに適用 / Apply dialog settings to target paragraph styles */
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
                // 欧文泣き別れ: 対応する ParagraphStyle プロパティが DOM 上で確認できず未適用 / Latin word break: no confirmed DOM property
                // safeSetProperty(paragraphStyle, "???", dialogResult.latinWordBreak);

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

            try { app.textPreferences.useSmartQuotes = !!dialogResult.useSmartQuotes; } catch (eSq) { }

            if (skipped > 0) {
                alert("適用しましたが、" + skipped + " 件の段落スタイルでエラーが発生しました。\n\n" + errorDetails.join("\n"));
            }
        },
        ScriptLanguage.JAVASCRIPT,
        undefined,
        UndoModes.ENTIRE_SCRIPT,
        "文字組版設定を段落スタイルに適用"
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
    var kinsokuHangTypeOptions = createKinsokuHangTypeOptions();
    var mojikumiTableData = collectMojikumiTables(activeDocument);
    var targetParagraphStyleData = collectTargetParagraphStyles(activeDocument);
    var targetParagraphStyles = targetParagraphStyleData.styles;

    if (targetParagraphStyles.length === 0) {
        alert("適用可能な段落スタイルがありません。");
        return;
    }

    var leadingModelOptions = createLeadingModelOptions();
    var characterAlignmentOptions = createCharacterAlignmentOptions();
    var gridAlignmentOptions = createGridAlignmentOptions();
    var kerningMethodOptions = createKerningMethodOptions();
    var composerOptions = createComposerOptions();
    var defaultPreset = PRESETS["InDesignのデフォルト"];
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
        alert("適用対象の段落スタイルが見つかりません。選択範囲、または指定した段落スタイルを確認してください。");
        return;
    }

    applyTypesettingSettingsToAll(resolvedTargetParagraphStyles, dialogResult, lookupTables);

    clearOverridesIfActive();

})();