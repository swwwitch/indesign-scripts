#target indesign

/*
 * NestedStyleSetup.jsx
 *
 * 段落スタイルに正規表現スタイル（GREP スタイル）を適用・管理します。ルールと文字スタイルを選び、複数の段落スタイルへまとめて反映できます。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "NestedStyleSetup";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-05-03";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/NestedStyleSetup.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/NestedStyleSetup.md

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

/**
 * コロン付きラベルを取得する（日本語は全角コロン、英語は半角コロン）
 * @param {string} labelKey 例: "panel.regex"
 * @returns {string} コロンを付与したラベル文字列
 */
function getLabelWithColon(labelKey) {
    return getLabel(labelKey) + (currentLang === "ja" ? "：" : ":");
}

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
            defaultCharacter: "currentLang-US",
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
     * 名前からスタイルを探す
     * @param {object} styleCollection スタイルのコレクション
     * @param {string} targetStyleName 探すスタイル名
     * @returns {object|null} スタイル。見つからない場合は null
     */
    function findStyleByName(styleCollection, targetStyleName) {
        for (var styleIndex = 0; styleIndex < styleCollection.length; styleIndex++) {
            if (styleCollection[styleIndex].name === targetStyleName) return styleCollection[styleIndex];
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

        var addGrepRuleButton = regexPanel.add("button", undefined, getLabel("button.addRule"));
        addGrepRuleButton.alignment = "right";
        addGrepRuleButton.helpTip = getLabel("tooltip.addRuleButton");

        /* 共通の文字スタイルパネル / Shared character style panel */
        var sharedCharacterStylePanel = leftColumn.add("panel", undefined, getLabel("panel.characterStyle"));
        setupPanel(sharedCharacterStylePanel, 6);

        var sharedCharacterStyleDropdown = sharedCharacterStylePanel.add("dropdownlist", undefined, characterStyleNames);
        sharedCharacterStyleDropdown.preferredSize.width = 160;
        sharedCharacterStyleDropdown.helpTip = getLabel("tooltip.characterStyleDropdown");

        var characterStyleCreateGroup = sharedCharacterStylePanel.add("group");
        characterStyleCreateGroup.orientation = "row";
        characterStyleCreateGroup.spacing = 4;
        var newCharacterStyleNameInput = characterStyleCreateGroup.add("edittext", undefined, "");
        newCharacterStyleNameInput.preferredSize.width = 120;
        newCharacterStyleNameInput.helpTip = getLabel("tooltip.newCharacterStyleName");
        var createCharacterStyleButton = characterStyleCreateGroup.add("button", undefined, getLabel("button.create"));
        createCharacterStyleButton.helpTip = getLabel("tooltip.createCharacterStyle");

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

        addGrepRuleButton.onClick = function () {

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
            /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */

            var btnGroup = ruleDialog.add("group");

            setupRow(btnGroup, "right", 8);

            var cancelBtn = btnGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
            var okBtn = btnGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
            cancelBtn.helpTip = getLabel("tooltip.addRuleDialogCancel");
            okBtn.helpTip = getLabel("tooltip.addRuleDialogOk");

            // --- OKボタン制御 ---
            /**
             * 入力状態に応じてコントロールの有効／無効を切り替える
             * @returns {void}
             */
            function updateState() {
                okBtn.enabled = (titleInput.text.length > 0 && expressionInput.text.length > 0);
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

        createCharacterStyleButton.onClick = function () {
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

        /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
        var buttonRow = dialogWindow.add("group");
        setupRow(buttonRow, "right", 8);
        var cancelButton = buttonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var okButton = buttonRow.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        cancelButton.helpTip = getLabel("tooltip.mainCancel");
        okButton.helpTip = getLabel("tooltip.mainOk");

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

            okButton.enabled = hasCharacter && hasParagraph;
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