#target indesign

/*
 * 概要
 *
 * アクティブな InDesign ドキュメントの段落スタイルに対し、
 * 正規表現スタイル（GREP スタイル）を適用・管理する。
 *
 * - ダイアログボックスは日本語／英語環境に応じて表示を切り替える
 * - タイトルバーにスクリプトのバージョンを表示する
 * - 左カラムの「正規表現のルール」パネルで、登録済みルールの選択とカスタムルールの追加を行う
 * - ルール追加は専用ダイアログで、管理用名称と正規表現を入力する
 * - 選択中の正規表現は `\\t` などを文字として読める形で表示する
 * - 左カラムの「適用する文字スタイル」パネルで、適用する文字スタイルを選択／新規作成する
 * - ルールに応じて、候補の文字スタイルが存在する場合は自動選択する
 * - 新規文字スタイル作成時、選択中ルールに応じた初期設定を自動適用する
 * - 右カラムの「適用先の段落スタイル」パネルで、対象の段落スタイルを複数選択する
 * - 段落スタイル一覧は Option/Alt クリックで全選択／全解除できる
 * - 必須項目（段落スタイル／文字スタイル）が未選択の場合、OKボタンは無効化される
 * - 同じ段落スタイル内に同一の正規表現がある場合は上書きする
 * - 主要UIには日本語／英語の tooltip を設定する
 *
 * 既定ルール:
 * - 箇条書きのラベル      : `^.+?(?=：)` → 文字スタイル候補「li-label」
 * - 言語設定              : `[\\u\\l]` → 文字スタイル候補「lang-US」／新規作成時に言語「英語：米国」を設定
 * - スマル                : `..[。」』？！…]?$` → 文字スタイル候補「sumaru」／新規作成時に分割禁止を設定
 * - 目次の数字            : `(?<=\\t)\\d+`
 * - インライングラフィック: `~a` → 文字スタイル候補「inline-graphic」／新規作成時に前後四分アキを設定
 */

/*
 * Summary
 *
 * Apply and manage GREP styles for paragraph styles
 * in the active InDesign document.
 *
 * - Switch UI labels between Japanese and English based on the locale
 * - Show the script version in the dialog title bar
 * - Select saved rules and add custom rules in the left-column GREP Rules panel
 * - Add custom rules with a dedicated dialog for a management name and GREP expression
 * - Display the selected GREP expression with escaped characters such as `\\t` shown literally
 * - Select or create the character style in the left-column Character Style to Apply panel
 * - Automatically select the preferred character style when it exists
 * - Apply rule-specific defaults when creating a new character style
 * - Select target paragraph styles in the right-column Target Paragraph Styles panel
 * - Option/Alt-click the paragraph style list to select or deselect all items
 * - Disable the OK button when required selections are missing
 * - Overwrite existing GREP styles with the same expression in the same paragraph style
 * - Provide localized tooltips for the main UI controls
 */

// =========================================
// バージョンとローカライズ / Version and localization
// =========================================

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "正規表現スタイルを適用", en: "Apply GREP Styles" },
    noDocumentError: { ja: "ドキュメントを開いてから実行してください。", en: "Open a document before running this script." },
    noParagraphStylesError: { ja: "段落スタイルが見つかりません。", en: "No paragraph styles were found." },
    regexPanel: { ja: "正規表現のルール", en: "GREP Rules" },
    addRuleButton: { ja: "＋ ルール追加", en: "+ Add Rule" },
    characterStylePanel: { ja: "適用する文字スタイル", en: "Character Style to Apply" },
    createButton: { ja: "作成", en: "Create" },
    targetParagraphStylesPanel: { ja: "適用先の段落スタイル", en: "Target Paragraph Styles" },
    multipleSelectHint: { ja: "複数の段落スタイルを選択できます", en: "Multiple selection allowed" },
    cancelButton: { ja: "キャンセル", en: "Cancel" },
    okButton: { ja: "OK", en: "OK" },
    addRuleTitlePrompt: { ja: "追加する正規表現の管理用の名称を入力してください。", en: "Enter a management name for the GREP expression to add." },
    addRuleTitleDefault: { ja: "新規ルール", en: "New Rule" },
    addRuleExpressionPrompt: { ja: "正規表現を入力してください。", en: "Enter a GREP expression." },
    // --- Tooltip labels for UI ---
    regexRuleListTip: { ja: "登録済みの正規表現ルールを選択します。右側で適用先の段落スタイルを指定します。", en: "Select a saved GREP rule. Configure target paragraph styles on the right." },
    selectedExpressionTip: { ja: "選択中の正規表現です。\\t などの制御文字は、文字として読める形で表示します。", en: "The selected GREP expression. Control characters such as \\t are shown literally." },
    addRuleButtonTip: { ja: "新しい正規表現ルールを追加します。", en: "Add a new GREP rule." },
    characterStyleDropdownTip: { ja: "適用する文字スタイルを選択します。", en: "Select the character style to apply." },
    newCharacterStyleNameTip: { ja: "新しく作成する文字スタイル名を入力します。", en: "Enter a name for the new character style." },
    createCharacterStyleButtonTip: { ja: "文字スタイルを作成します。ルールによっては言語設定・分割禁止・前後アキなどの初期設定を自動適用します。", en: "Create a character style. Depending on the selected rule, default settings such as language, no-break, or spacing are applied automatically." },
    paragraphStyleListTip: { ja: "正規表現スタイルを適用する段落スタイルを選択します。Option/Altクリックで全選択／全解除できます。", en: "Select paragraph styles to apply the GREP style to. Option/Alt-click toggles select all." },
    addRuleTitleInputTip: { ja: "このルールを識別するための管理用名称です。処理内容には影響しません。", en: "Management name used to identify this rule. It does not affect processing." },
    addRuleExpressionInputTip: { ja: "適用する正規表現を入力します。例：(?<=\\t)\\d+", en: "Enter the GREP expression to apply. Example: (?<=\\t)\\d+" },
    addRuleDialogOkTip: { ja: "入力した名称と正規表現でルールを追加します。", en: "Add a rule using the entered name and GREP expression." },
    addRuleDialogCancelTip: { ja: "ルール追加をキャンセルします。", en: "Cancel adding the rule." },
    mainOkButtonTip: { ja: "選択した文字スタイルと段落スタイルに、選択中の正規表現スタイルを適用します。", en: "Apply the selected GREP style using the selected character and paragraph styles." },
    mainCancelButtonTip: { ja: "処理を実行せずに閉じます。", en: "Close without applying changes." },
    // ---
    emptyCharacterStyleNameError: { ja: "文字スタイル名を入力してください。", en: "Enter a character style name." },
    duplicateCharacterStyleNameError: { ja: "同名の文字スタイルが既にあります。", en: "A character style with the same name already exists." },
    ruleBulletLabel: { ja: "箇条書きのラベル", en: "Bullet Label" },
    ruleLanguage: { ja: "言語設定", en: "Language" },
    ruleSumaru: { ja: "スマル", en: "No-break Ending" },
    ruleTocNumber: { ja: "目次の数字", en: "TOC Number" },
    ruleInlineGraphic: { ja: "インライングラフィック", en: "Inline Graphic" }
};

function getLabel(labelKey) {
    return LABELS[labelKey] ? LABELS[labelKey][lang] : labelKey;
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(labelKey) {
    return getLabel(labelKey) + (lang === "ja" ? "：" : ":");
}

(function () {
    if (app.documents.length === 0) {
        alert(getLabel("noDocumentError"));
        return;
    }
    var doc = app.activeDocument;

    var nestedGrepRules = [
        {
            key: "bulletLabel",
            title: getLabel("ruleBulletLabel"),
            defaultParagraph: "ul-li",
            defaultCharacter: "li-label",
            autoSelect: true,
            expression: "^.+?(?=：)"
        },
        {
            key: "language",
            title: getLabel("ruleLanguage"),
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
            title: getLabel("ruleSumaru"),
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
            title: getLabel("ruleTocNumber"),
            defaultParagraph: "p",
            defaultCharacter: "",
            autoSelect: false,
            expression: "(?<=\\t)\\d+"
        },
        {
            key: "inlineGraphic",
            title: getLabel("ruleInlineGraphic"),
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

    var PANEL_MARGINS = [15, 20, 15, 10];

    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            panel.spacing = spacing;
        }
    }

    function collectVisibleStyleNames(styleCollection) {
        var visibleStyleNames = [];
        for (var styleIndex = 0; styleIndex < styleCollection.length; styleIndex++) {
            var styleName = styleCollection[styleIndex].name;
            if (styleName.charAt(0) === "[") continue;
            visibleStyleNames.push(styleName);
        }
        return visibleStyleNames;
    }

    function findNameIndex(nameList, targetName) {
        for (var nameIndex = 0; nameIndex < nameList.length; nameIndex++) {
            if (nameList[nameIndex] === targetName) return nameIndex;
        }
        return -1;
    }

    function findStyleByName(styleCollection, targetStyleName) {
        for (var styleIndex = 0; styleIndex < styleCollection.length; styleIndex++) {
            if (styleCollection[styleIndex].name === targetStyleName) return styleCollection[styleIndex];
        }
        return null;
    }

    function showDialog(grepRules, paragraphStyleNames, characterStyleNames, targetDocument) {
        var dialogWindow = new Window("dialog", getLabel("dialogTitle") + " " + SCRIPT_VERSION);
        dialogWindow.alignChildren = "fill";
        dialogWindow.margins = 16;
        dialogWindow.spacing = 12;

        var mainContentGroup = dialogWindow.add("group");
        mainContentGroup.orientation = "row";
        mainContentGroup.alignChildren = ["fill", "fill"];
        mainContentGroup.spacing = 12;

        var leftColumn = mainContentGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];
        leftColumn.spacing = 8;

        var regexPanel = leftColumn.add("panel", undefined, getLabel("regexPanel"));
        setupPanel(regexPanel, 6);
        var grepRuleTitles = [];
        for (var ruleIndex = 0; ruleIndex < grepRules.length; ruleIndex++) grepRuleTitles.push(grepRules[ruleIndex].title);
        var grepRuleListbox = regexPanel.add("listbox", undefined, grepRuleTitles);
        grepRuleListbox.alignment = ["fill", "top"];
        grepRuleListbox.preferredSize.height = 140;
        grepRuleListbox.helpTip = getLabel("regexRuleListTip");

        var selectedRegexGroup = regexPanel.add("group");
        selectedRegexGroup.orientation = "column";
        selectedRegexGroup.alignChildren = "left";
        selectedRegexGroup.spacing = 3;
        var selectedExpressionText = selectedRegexGroup.add("edittext", undefined, "");
        // selectedExpressionText.alignment = ["fill", "top"];
        selectedExpressionText.preferredSize.width = 180;
        selectedExpressionText.enabled = false;
        selectedExpressionText.helpTip = getLabel("selectedExpressionTip");

        var addGrepRuleButton = regexPanel.add("button", undefined, getLabel("addRuleButton"));
        addGrepRuleButton.alignment = "right";
        addGrepRuleButton.helpTip = getLabel("addRuleButtonTip");

        /* 共通の文字スタイルパネル / Shared character style panel */
        var sharedCharacterStylePanel = leftColumn.add("panel", undefined, getLabel("characterStylePanel"));
        setupPanel(sharedCharacterStylePanel, 6);

        var sharedCharacterStyleDropdown = sharedCharacterStylePanel.add("dropdownlist", undefined, characterStyleNames);
        sharedCharacterStyleDropdown.preferredSize.width = 160;
        sharedCharacterStyleDropdown.helpTip = getLabel("characterStyleDropdownTip");

        var characterStyleCreateGroup = sharedCharacterStylePanel.add("group");
        characterStyleCreateGroup.orientation = "row";
        characterStyleCreateGroup.spacing = 4;
        var newCharacterStyleNameInput = characterStyleCreateGroup.add("edittext", undefined, "");
        newCharacterStyleNameInput.preferredSize.width = 120;
        newCharacterStyleNameInput.helpTip = getLabel("newCharacterStyleNameTip");
        var createCharacterStyleButton = characterStyleCreateGroup.add("button", undefined, getLabel("createButton"));
        createCharacterStyleButton.helpTip = getLabel("createCharacterStyleButtonTip");

        /* 右カラム（縦構造）/ Right column with vertical layout */
        var paragraphStyleColumn = mainContentGroup.add("group");
        paragraphStyleColumn.orientation = "column";
        paragraphStyleColumn.alignChildren = ["fill", "top"];

        var paragraphStylePanelStack = paragraphStyleColumn.add("group");
        paragraphStylePanelStack.orientation = "stack";
        paragraphStylePanelStack.alignChildren = ["fill", "top"];

        var ruleRows = [];

        function getDefaultCharacterStyleNameForRule(grepRule) {
            if (!grepRule) return "";
            return grepRule.defaultCharacter || "";
        }

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

        function createGrepRuleRow(grepRule) {
            var paragraphStylePanel = paragraphStylePanelStack.add("panel", undefined, getLabel("targetParagraphStylesPanel"));
            setupPanel(paragraphStylePanel, 6);
            paragraphStylePanel.preferredSize = [280, 300];
            paragraphStylePanel.visible = false;

            paragraphStylePanel.add("statictext", undefined, getLabel("multipleSelectHint"));
            var paragraphStyleListbox = paragraphStylePanel.add("listbox", undefined, paragraphStyleNames, { multiselect: true });
            paragraphStyleListbox.alignment = ["fill", "fill"];
            paragraphStyleListbox.preferredSize.height = 300;
            paragraphStyleListbox.helpTip = getLabel("paragraphStyleListTip");

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

            var ruleDialog = new Window("dialog", getLabel("addRuleButton") + " " + SCRIPT_VERSION);
            ruleDialog.orientation = "column";
            ruleDialog.alignChildren = "fill";
            ruleDialog.margins = 16;
            ruleDialog.spacing = 10;

            // --- タイトル入力 ---
            var titleGroup = ruleDialog.add("group");
            titleGroup.orientation = "column";
            titleGroup.alignChildren = "left";

            titleGroup.add("statictext", undefined, labelText("addRuleTitlePrompt"));
            var titleInput = titleGroup.add("edittext", undefined, getLabel("addRuleTitleDefault"));
            titleInput.preferredSize.width = 240;
            titleInput.helpTip = getLabel("addRuleTitleInputTip");

            // --- 正規表現入力 ---
            var exprGroup = ruleDialog.add("group");
            exprGroup.orientation = "column";
            exprGroup.alignChildren = "left";

            exprGroup.add("statictext", undefined, labelText("addRuleExpressionPrompt"));
            var expressionInput = exprGroup.add("edittext", undefined, "");
            expressionInput.preferredSize.width = 240;
            expressionInput.helpTip = getLabel("addRuleExpressionInputTip");

            // --- ボタン ---
            var btnGroup = ruleDialog.add("group");
            btnGroup.alignment = "right";

            var cancelBtn = btnGroup.add("button", undefined, getLabel("cancelButton"), { name: "cancel" });
            var okBtn = btnGroup.add("button", undefined, getLabel("okButton"), { name: "ok" });
            cancelBtn.helpTip = getLabel("addRuleDialogCancelTip");
            okBtn.helpTip = getLabel("addRuleDialogOkTip");

            // --- OKボタン制御 ---
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
                alert(getLabel("emptyCharacterStyleNameError"));
                return;
            }
            if (findNameIndex(characterStyleNames, newCharacterStyleName) >= 0) {
                alert(getLabel("duplicateCharacterStyleNameError"));
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

        var buttonRow = dialogWindow.add("group");
        buttonRow.alignment = "right";
        var cancelButton = buttonRow.add("button", undefined, getLabel("cancelButton"), { name: "cancel" });
        var okButton = buttonRow.add("button", undefined, getLabel("okButton"), { name: "ok" });
        cancelButton.helpTip = getLabel("mainCancelButtonTip");
        okButton.helpTip = getLabel("mainOkButtonTip");

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
        alert(getLabel("noParagraphStylesError"));
        return;
    }

    var styleRegistrations = showDialog(nestedGrepRules, paragraphStyleNames, characterStyleNames, doc);
    if (!styleRegistrations) return;

    applyNestedGrepStyleSettings(doc, styleRegistrations);
})();