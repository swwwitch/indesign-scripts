#target indesign

/*

### 概要

段落スタイルに正規表現スタイル（GREP スタイル）を適用・管理します。ルールと文字スタイルを選び、複数の段落スタイルへまとめて反映できます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdGrepStyleApplier.md

### Overview

Applies and manages GREP styles on paragraph styles. Pick a rule and a character style and push it to several paragraph styles at once.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdGrepStyleApplier.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdGrepStyleApplier";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdGrepStyleApplier.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdGrepStyleApplier.md"; /* README (English) */

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
        title:      { ja: "正規表現スタイルを適用", en: "Apply GREP Styles" },
        addRuleTitleDefault: { ja: "新規ルール", en: "New Rule" }
    },
    panel: {
        regex:                 { ja: "正規表現のルール", en: "GREP Rules" },
        characterStyle:        { ja: "適用する文字スタイル", en: "Character Style to Apply" },
        targetParagraphStyles: { ja: "適用先の段落スタイル", en: "Target Paragraph Styles" }
    },
    button: {
        addRule: { ja: "＋ ルール追加", en: "+ Add Rule" },
        create:  { ja: "作成", en: "Create" },
        cancel:  { ja: "キャンセル", en: "Cancel" },
        ok:      { ja: "OK", en: "OK" }
    },
    hint: {
        multipleSelect: { ja: "複数の段落スタイルを選択できます", en: "Multiple selection allowed" }
    },
    prompt: {
        addRuleTitle:      { ja: "追加する正規表現の管理用の名称を入力してください。", en: "Enter a management name for the GREP expression to add." },
        addRuleExpression: { ja: "正規表現を入力してください。", en: "Enter a GREP expression." }
    },
    tooltip: {
        regexRuleList:          { ja: "登録済みの正規表現ルールを選択します。右側で適用先の段落スタイルを指定します。", en: "Select a saved GREP rule. Configure target paragraph styles on the right." },
        selectedExpression:     { ja: "選択中の正規表現です。\\t などの制御文字は、文字として読める形で表示します。", en: "The selected GREP expression. Control characters such as \\t are shown literally." },
        addRuleButton:          { ja: "新しい正規表現ルールを追加します。", en: "Add a new GREP rule." },
        characterStyleDropdown: { ja: "適用する文字スタイルを選択します。", en: "Select the character style to apply." },
        newCharacterStyleName:  { ja: "新しく作成する文字スタイル名を入力します。", en: "Enter a name for the new character style." },
        createCharacterStyle:   { ja: "文字スタイルを作成します。ルールによっては言語設定・分割禁止・前後アキなどの初期設定を自動適用します。", en: "Create a character style. Depending on the selected rule, default settings such as language, no-break, or spacing are applied automatically." },
        paragraphStyleList:     { ja: "正規表現スタイルを適用する段落スタイルを選択します。Option/Altクリックで全選択／全解除できます。", en: "Select paragraph styles to apply the GREP style to. Option/Alt-click toggles select all." },
        addRuleTitleInput:      { ja: "このルールを識別するための管理用名称です。処理内容には影響しません。", en: "Management name used to identify this rule. It does not affect processing." },
        addRuleExpressionInput: { ja: "適用する正規表現を入力します。例：(?<=\\t)\\d+", en: "Enter the GREP expression to apply. Example: (?<=\\t)\\d+" },
        addRuleDialogOk:        { ja: "入力した名称と正規表現でルールを追加します。", en: "Add a rule using the entered name and GREP expression." },
        addRuleDialogCancel:    { ja: "ルール追加をキャンセルします。", en: "Cancel adding the rule." },
        mainOk:                 { ja: "選択した文字スタイルと段落スタイルに、選択中の正規表現スタイルを適用します。", en: "Apply the selected GREP style using the selected character and paragraph styles." },
        mainCancel:             { ja: "処理を実行せずに閉じます。", en: "Close without applying changes." }
    },
    rule: {
        bulletLabel:    { ja: "箇条書きのラベル", en: "Bullet Label" },
        language:       { ja: "言語設定", en: "Language" },
        sumaru:         { ja: "スマル", en: "No-break Ending" },
        tocNumber:      { ja: "目次の数字", en: "TOC Number" },
        inlineGraphic:  { ja: "インライングラフィック", en: "Inline Graphic" }
    },
    undo: {
        applyGrepStyles: { ja: "正規表現スタイルを適用", en: "Apply GREP Styles" }
    },
    error: {
        noDocument:                 { ja: "ドキュメントを開いてから実行してください。", en: "Open a document before running this script." },
        noParagraphStyles:          { ja: "段落スタイルが見つかりません。", en: "No paragraph styles were found." },
        emptyCharacterStyleName:    { ja: "文字スタイル名を入力してください。", en: "Enter a character style name." },
        duplicateCharacterStyleName:{ ja: "同名の文字スタイルが既にあります。", en: "A character style with the same name already exists." }
    }
};

(function () {
    if (app.documents.length === 0) {
        alert(getLabel("error.noDocument"));
        return;
    }
    var doc = app.activeDocument;

    var nestedGrepRules = [
        {
            key: "bulletLabel",
            title: getLabel("rule.bulletLabel"),
            defaultParagraph: "ul-li",
            defaultCharacter: "li-label",
            autoSelect: true,
            expression: "^.+?(?=：)"
        },
        {
            key: "language",
            title: getLabel("rule.language"),
            defaultParagraph: "p",
            defaultCharacter: "lang-US",
            autoSelect: true,
            expression: "[\\u\\l]",
            apply: function (style) {
                var langNames = ["English: USA", "英語：米国"];
                for (var languageNameIndex = 0; languageNameIndex < langNames.length; languageNameIndex++) {
                    var candidate = app.languagesWithVendors.itemByName(langNames[languageNameIndex]);
                    if (candidate.isValid) {
                        style.appliedLanguage = candidate;
                        break;
                    }
                }
            }
        },
        {
            key: "sumaru",
            title: getLabel("rule.sumaru"),
            defaultParagraph: "p",
            defaultCharacter: "sumaru",
            autoSelect: true,
            expression: "..[。」』？！…]?$",
            apply: function (style) {
                style.noBreak = true;
            }
        },
        {
            key: "tocNumber",
            title: getLabel("rule.tocNumber"),
            defaultParagraph: "p",
            defaultCharacter: "",
            autoSelect: false,
            expression: "(?<=\\t)\\d+"
        },
        {
            key: "inlineGraphic",
            title: getLabel("rule.inlineGraphic"),
            defaultParagraph: "p",
            defaultCharacter: "inline-graphic",
            autoSelect: true,
            expression: "~a",
            apply: function (style) {
                style.leadingAki = 0.25;
                style.trailingAki = 0.25;
            }
        }
    ];

    /**
     * 「[...]」で始まる既定スタイルを除いたスタイル名を集める
     * @param {object} styleCollection スタイルのコレクション
     * @returns {Array<string>} スタイル名の配列
     */
    function collectVisibleStyleNames(styleCollection) {
        var visibleStyleNames = [];
        for (var styleIndex = 0; styleIndex < styleCollection.length; styleIndex++) {
            var styleName = styleCollection[styleIndex].name;
            if (styleName.charAt(0) === "[") continue;
            visibleStyleNames.push(styleName);
        }
        return visibleStyleNames;
    }

    /**
     * 名前の一覧から一致する位置を探す
     * @param {Array<string>} nameList 名前の一覧
     * @param {string} targetName 探す名前
     * @returns {number} 見つかった位置。なければ -1
     */
    function findNameIndex(nameList, targetName) {
        for (var nameIndex = 0; nameIndex < nameList.length; nameIndex++) {
            if (nameList[nameIndex] === targetName) return nameIndex;
        }
        return -1;
    }

    /**
     * 名前でスタイルを探す（スタイルグループ内も対象）
     * @param {array} styleCollection allParagraphStyles / allCharacterStyles
     * @param {string} styleName 探すスタイル名
     * @returns {object} 見つかったスタイル。無い場合は null
     */
    function findStyleByName(styleCollection, styleName) {
        for (var i = 0; i < styleCollection.length; i++) {
            if (styleCollection[i].name === styleName) return styleCollection[i];
        }

        return null;
    }

    /**
     * 正規表現スタイルの設定ダイアログを表示する
     * @param {Array<object>} grepRules 正規表現ルールの一覧
     * @param {Array<string>} paragraphStyleNames 段落スタイル名の一覧
     * @param {Array<string>} characterStyleNames 文字スタイル名の一覧
     * @param {Document} targetDocument 対象ドキュメント
     * @returns {Array<object>|null} 適用するスタイルの組み合わせ。キャンセル時は null
     */
    function showDialog(grepRules, paragraphStyleNames, characterStyleNames, targetDocument) {
        var dialogWindow = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialogWindow);

        var mainContentGroup = dialogWindow.add("group");
        setupRow(mainContentGroup, "fill", COLUMN_SPACING);
        mainContentGroup.alignChildren = ["fill", "fill"];

        var leftColumn = mainContentGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];
        leftColumn.spacing = 8;

        var regexPanel = leftColumn.add("panel", undefined, getLabel("panel.regex"));
        setupPanel(regexPanel, 6);
        var grepRuleTitles = [];
        for (var ruleIndex = 0; ruleIndex < grepRules.length; ruleIndex++) grepRuleTitles.push(grepRules[ruleIndex].title);
        var grepRuleListbox = regexPanel.add("listbox", undefined, grepRuleTitles);
        grepRuleListbox.alignment = ["fill", "top"];
        grepRuleListbox.preferredSize.height = 140;
        grepRuleListbox.helpTip = getLabel("tooltip.regexRuleList");

        var selectedRegexGroup = regexPanel.add("group");
        selectedRegexGroup.orientation = "column";
        selectedRegexGroup.alignChildren = "left";
        selectedRegexGroup.spacing = 3;
        var selectedExpressionText = selectedRegexGroup.add("edittext", undefined, "");
        // selectedExpressionText.alignment = ["fill", "top"];
        selectedExpressionText.preferredSize.width = 180;
        selectedExpressionText.enabled = false;
        selectedExpressionText.helpTip = getLabel("tooltip.selectedExpression");

        var btnAddRule = regexPanel.add("button", undefined, getLabel("button.addRule"));
        btnAddRule.alignment = "right";
        btnAddRule.helpTip = getLabel("tooltip.addRuleButton");

        /* 共通の文字スタイルパネル / Shared character style panel */
        var sharedCharacterStylePanel = leftColumn.add("panel", undefined, getLabel("panel.characterStyle"));
        setupPanel(sharedCharacterStylePanel, 6);

        var sharedCharacterStyleDropdown = sharedCharacterStylePanel.add("dropdownlist", undefined, characterStyleNames);
        sharedCharacterStyleDropdown.preferredSize.width = 160;
        sharedCharacterStyleDropdown.helpTip = getLabel("tooltip.characterStyleDropdown");

        var characterStyleCreateGroup = sharedCharacterStylePanel.add("group");
        setupRow(characterStyleCreateGroup, "left", 4);
        var newCharacterStyleNameInput = characterStyleCreateGroup.add("edittext", undefined, "");
        newCharacterStyleNameInput.preferredSize.width = 120;
        newCharacterStyleNameInput.helpTip = getLabel("tooltip.newCharacterStyleName");
        var btnCreateCharacterStyle = characterStyleCreateGroup.add("button", undefined, getLabel("button.create"));
        btnCreateCharacterStyle.helpTip = getLabel("tooltip.createCharacterStyle");

        /* 右カラム（縦構造）/ Right column with vertical layout */
        var paragraphStyleColumn = mainContentGroup.add("group");
        paragraphStyleColumn.orientation = "column";
        paragraphStyleColumn.alignChildren = ["fill", "top"];

        var paragraphStylePanelStack = paragraphStyleColumn.add("group");
        paragraphStylePanelStack.orientation = "stack";
        paragraphStylePanelStack.alignChildren = ["fill", "top"];

        var ruleRows = [];

        /**
         * ルールに対応する既定の文字スタイル名を返す
         * @param {object} grepRule 正規表現ルール
         * @returns {string} 文字スタイル名
         */
        function getDefaultCharacterStyleNameForRule(grepRule) {
            if (!grepRule) return "";
            return grepRule.defaultCharacter || "";
        }

        /**
         * 選択したルールに合わせて表示と候補を更新する
         * @param {number} selectedRuleIndex 選択したルールの位置
         * @returns {void}
         */
        function updateSelectedGrepRule(selectedRuleIndex) {
            for (var rowIndex = 0; rowIndex < ruleRows.length; rowIndex++) {
                ruleRows[rowIndex].paragraphStylePanel.visible = (rowIndex === selectedRuleIndex);
            }
            var expressionForDisplay = (selectedRuleIndex >= 0 && ruleRows[selectedRuleIndex]) ? ruleRows[selectedRuleIndex].rule.expression : "";
            /* 表示用にエスケープ（制御文字のみ）/ Escape control characters for display (\t, \n, \r) */
            selectedExpressionText.text = expressionForDisplay
                .replace(/\t/g, "\\t")
                .replace(/\n/g, "\\n")
                .replace(/\r/g, "\\r");

            if (selectedRuleIndex >= 0 && ruleRows[selectedRuleIndex]) {
                var selectedGrepRule = ruleRows[selectedRuleIndex].rule;
                if (selectedGrepRule.autoSelect) {
                    var defaultCharacterStyleName = getDefaultCharacterStyleNameForRule(selectedGrepRule);
                    var defaultCharacterStyleIndex = findNameIndex(characterStyleNames, defaultCharacterStyleName);
                    sharedCharacterStyleDropdown.selection = (defaultCharacterStyleIndex >= 0) ? defaultCharacterStyleIndex : null;
                } else {
                    sharedCharacterStyleDropdown.selection = null;
                }
            } else {
                sharedCharacterStyleDropdown.selection = null;
            }
        }

        /**
         * ルール 1 件分の表示行を作る
         * @param {object} grepRule 正規表現ルール
         * @returns {string} リストに表示する文字列
         */
        function createGrepRuleRow(grepRule) {
            var paragraphStylePanel = paragraphStylePanelStack.add("panel", undefined, getLabel("panel.targetParagraphStyles"));
            setupPanel(paragraphStylePanel, 6);
            paragraphStylePanel.preferredSize = [280, 300];
            paragraphStylePanel.visible = false;

            paragraphStylePanel.add("statictext", undefined, getLabel("hint.multipleSelect"));
            var paragraphStyleListbox = paragraphStylePanel.add("listbox", undefined, paragraphStyleNames, { multiselect: true });
            paragraphStyleListbox.alignment = ["fill", "fill"];
            paragraphStyleListbox.preferredSize.height = 300;
            paragraphStyleListbox.helpTip = getLabel("tooltip.paragraphStyleList");

            paragraphStyleListbox.onClick = function () {
                /* Option/Altクリックで全選択トグル / Toggle select all with Option/Alt-click */
                var isAlt = ScriptUI.environment.keyboardState.altKey;
                if (!isAlt) return;

                var shouldSelectAll = false;
                for (var paragraphStyleIndex = 0; paragraphStyleIndex < paragraphStyleListbox.items.length; paragraphStyleIndex++) {
                    if (!paragraphStyleListbox.items[paragraphStyleIndex].selected) {
                        shouldSelectAll = true;
                        break;
                    }
                }

                for (var selectIndex = 0; selectIndex < paragraphStyleListbox.items.length; selectIndex++) {
                    paragraphStyleListbox.items[selectIndex].selected = shouldSelectAll;
                }
            };
            for (var defaultParagraphStyleIndex = 0; defaultParagraphStyleIndex < paragraphStyleListbox.items.length; defaultParagraphStyleIndex++) {
                if (paragraphStyleListbox.items[defaultParagraphStyleIndex].text === grepRule.defaultParagraph) {
                    paragraphStyleListbox.items[defaultParagraphStyleIndex].selected = true;
                }
            }

            /* 文字スタイルは共通パネルに統合 / Character style controls are unified in the shared panel */

            var ruleRow = {
                rule: grepRule,
                paragraphStylePanel: paragraphStylePanel,
                paragraphStyleListbox: paragraphStyleListbox
            };
            ruleRows.push(ruleRow);

            return ruleRow;
        }

        for (var grepRuleIndex = 0; grepRuleIndex < grepRules.length; grepRuleIndex++) {
            createGrepRuleRow(grepRules[grepRuleIndex]);
        }

        if (ruleRows.length > 0) {
            grepRuleListbox.selection = 0;
            updateSelectedGrepRule(0);
        }

        grepRuleListbox.onChange = function () {
            var selectedIndex = grepRuleListbox.selection ? grepRuleListbox.selection.index : -1;
            updateSelectedGrepRule(selectedIndex);
        };

        btnAddRule.onClick = function () {

            var ruleDialog = new Window("dialog", getLabel("button.addRule") + " " + SCRIPT_VERSION);
            setupWindow(ruleDialog, 10);

            // --- タイトル入力 ---
            var titleGroup = ruleDialog.add("group");
            titleGroup.orientation = "column";
            titleGroup.alignChildren = "left";

            titleGroup.add("statictext", undefined, getLabel("prompt.addRuleTitle"));
            var titleInput = titleGroup.add("edittext", undefined, getLabel("dialog.addRuleTitleDefault"));
            titleInput.preferredSize.width = 240;
            titleInput.helpTip = getLabel("tooltip.addRuleTitleInput");

            // --- 正規表現入力 ---
            var exprGroup = ruleDialog.add("group");
            exprGroup.orientation = "column";
            exprGroup.alignChildren = "left";

            exprGroup.add("statictext", undefined, getLabel("prompt.addRuleExpression"));
            var expressionInput = exprGroup.add("edittext", undefined, "");
            expressionInput.preferredSize.width = 240;
            expressionInput.helpTip = getLabel("tooltip.addRuleExpressionInput");

            // --- ボタン ---
            /* ボタン行（キャンセル → OK）/ Button row (Cancel, then OK) */
            var ruleButtonRow = addButtonRow(ruleDialog);
            var btnRuleCancel = ruleButtonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
            var btnRuleOK = ruleButtonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
            btnRuleCancel.helpTip = getLabel("tooltip.addRuleDialogCancel");
            btnRuleOK.helpTip = getLabel("tooltip.addRuleDialogOk");
            alignRightOnlyButtonRow(ruleButtonRow);

            // --- OKボタン制御 ---
            /**
             * 入力状態に応じてコントロールの有効／無効を切り替える
             * @returns {void}
             */
            function updateState() {
                btnRuleOK.enabled = (titleInput.text.length > 0 && expressionInput.text.length > 0);
            }

            titleInput.onChanging = updateState;
            expressionInput.onChanging = updateState;

            updateState();

            // --- 実行 ---
            if (ruleDialog.show() !== 1) return;

            var newGrepRule = {
                key: "custom_" + (new Date().getTime()),
                title: titleInput.text,
                defaultParagraph: paragraphStyleNames.length > 0 ? paragraphStyleNames[0] : "",
                defaultCharacter: "",
                autoSelect: false,
                expression: expressionInput.text
            };

            grepRules.push(newGrepRule);
            grepRuleListbox.add("item", newGrepRule.title);

            createGrepRuleRow(newGrepRule);
            grepRuleListbox.selection = grepRuleListbox.items.length - 1;
            updateSelectedGrepRule(ruleRows.length - 1);

            dialogWindow.layout.layout(true);
        };

        btnCreateCharacterStyle.onClick = function () {
            var newCharacterStyleName = newCharacterStyleNameInput.text;
            if (!newCharacterStyleName || newCharacterStyleName.length === 0) {
                alert(getLabel("error.emptyCharacterStyleName"));
                return;
            }
            if (findNameIndex(characterStyleNames, newCharacterStyleName) >= 0) {
                alert(getLabel("error.duplicateCharacterStyleName"));
                return;
            }
            var newCharacterStyle = doc.characterStyles.add({ name: newCharacterStyleName });
            var selectedGrepRuleIndex = grepRuleListbox.selection ? grepRuleListbox.selection.index : -1;
            var selectedGrepRule = (selectedGrepRuleIndex >= 0 && ruleRows[selectedGrepRuleIndex]) ? ruleRows[selectedGrepRuleIndex].rule : null;
            if (selectedGrepRule && typeof selectedGrepRule.apply === "function") {
                selectedGrepRule.apply(newCharacterStyle);
            }
            characterStyleNames.push(newCharacterStyleName);
            sharedCharacterStyleDropdown.add("item", newCharacterStyleName);
            sharedCharacterStyleDropdown.selection = sharedCharacterStyleDropdown.items.length - 1;
            newCharacterStyleNameInput.text = "";
            dialogWindow.layout.layout(true);
        };

        /* ボタン行（キャンセル → OK）/ Button row (Cancel, then OK) */
        var buttonRow = addButtonRow(dialogWindow);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        btnCancel.helpTip = getLabel("tooltip.mainCancel");
        btnOK.helpTip = getLabel("tooltip.mainOk");
        alignRightOnlyButtonRow(buttonRow);

        /**
         * 必須項目の入力状況に応じて OK ボタンの有効／無効を切り替える
         * @returns {void}
         */
        function updateOkButtonState() {
            var hasCharacter = sharedCharacterStyleDropdown.selection !== null;
            var hasParagraph = false;

            for (var okRuleRowIndex = 0; okRuleRowIndex < ruleRows.length; okRuleRowIndex++) {
                var listbox = ruleRows[okRuleRowIndex].paragraphStyleListbox;
                for (var okParagraphStyleIndex = 0; okParagraphStyleIndex < listbox.items.length; okParagraphStyleIndex++) {
                    if (listbox.items[okParagraphStyleIndex].selected) {
                        hasParagraph = true;
                        break;
                    }
                }
                if (hasParagraph) break;
            }

            btnOK.enabled = hasCharacter && hasParagraph;
        }

        sharedCharacterStyleDropdown.onChange = updateOkButtonState;

        for (var okBindingRowIndex = 0; okBindingRowIndex < ruleRows.length; okBindingRowIndex++) {
            (function (listbox) {
                listbox.onChange = updateOkButtonState;
            })(ruleRows[okBindingRowIndex].paragraphStyleListbox);
        }

        updateOkButtonState();

        if (dialogWindow.show() !== 1) return null;

        var styleRegistrations = [];
        for (var ruleRowIndex = 0; ruleRowIndex < ruleRows.length; ruleRowIndex++) {
            var ruleRow = ruleRows[ruleRowIndex];
            var selectedParagraphStyleNames = [];
            for (var paragraphStyleItemIndex = 0; paragraphStyleItemIndex < ruleRow.paragraphStyleListbox.items.length; paragraphStyleItemIndex++) {
                if (ruleRow.paragraphStyleListbox.items[paragraphStyleItemIndex].selected) {
                    selectedParagraphStyleNames.push(ruleRow.paragraphStyleListbox.items[paragraphStyleItemIndex].text);
                }
            }
            var selectedCharacterStyleName = (sharedCharacterStyleDropdown.selection) ? sharedCharacterStyleDropdown.selection.text : null;
            if (selectedParagraphStyleNames.length === 0 || !selectedCharacterStyleName) continue;
            for (var selectedParagraphStyleIndex = 0; selectedParagraphStyleIndex < selectedParagraphStyleNames.length; selectedParagraphStyleIndex++) {
                styleRegistrations.push({
                    paragraph: selectedParagraphStyleNames[selectedParagraphStyleIndex],
                    character: selectedCharacterStyleName,
                    expression: ruleRow.rule.expression
                });
            }
        }
        return styleRegistrations;
    }

    /**
     * 選択した段落スタイルへ正規表現スタイルを適用する
     * @param {Document} targetDocument 対象ドキュメント
     * @param {Array<object>} styleRegistrations 適用するスタイルの組み合わせ
     * @returns {void}
     */
    function applyNestedGrepStyleSettings(targetDocument, styleRegistrations) {
        var paragraphStyleCollection = targetDocument.allParagraphStyles;
        var characterStyleCollection = targetDocument.allCharacterStyles;

        for (var registrationIndex = 0; registrationIndex < styleRegistrations.length; registrationIndex++) {
            var styleRegistration = styleRegistrations[registrationIndex];
            var paragraphStyle = findStyleByName(paragraphStyleCollection, styleRegistration.paragraph);
            var characterStyle = findStyleByName(characterStyleCollection, styleRegistration.character);
            if (!paragraphStyle || !characterStyle) continue;

            for (var grepStyleIndex = paragraphStyle.nestedGrepStyles.length - 1; grepStyleIndex >= 0; grepStyleIndex--) {
                var existingNestedGrepStyle = paragraphStyle.nestedGrepStyles[grepStyleIndex];
                if (existingNestedGrepStyle.grepExpression === styleRegistration.expression) {
                    existingNestedGrepStyle.remove();
                }
            }

            var newNestedGrepStyle = paragraphStyle.nestedGrepStyles.add();
            newNestedGrepStyle.appliedCharacterStyle = characterStyle;
            newNestedGrepStyle.grepExpression = styleRegistration.expression;
        }
    }

    var paragraphStyleNames = collectVisibleStyleNames(doc.allParagraphStyles);
    var characterStyleNames = collectVisibleStyleNames(doc.allCharacterStyles);

    if (paragraphStyleNames.length === 0) {
        alert(getLabel("error.noParagraphStyles"));
        return;
    }

    var styleRegistrations = showDialog(nestedGrepRules, paragraphStyleNames, characterStyleNames, doc);
    if (!styleRegistrations) return;

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(function () {
        applyNestedGrepStyleSettings(doc, styleRegistrations);
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.applyGrepStyles"));
})();