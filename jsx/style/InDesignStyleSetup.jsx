#target indesign

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.1.0";

/*
概要

アクティブな InDesign ドキュメントに段落スタイル／文字スタイルと
それぞれのグループ（フォルダー）を一括登録する。

登録内容:
- 段落スタイル（ルート）: h1〜h6, ul-li, p, p.caption, p.code, p.img, ol-li
- 段落スタイルグループ:
    table -> th, td, td-left, td-center, td-right
    toc   -> toc-title, toc-h1, toc-h2, toc-h3
    book  -> page-number, running-head, thumb-index
- 文字スタイル（ルート）: strong-bold, em-italic, link, code-normal, code-strong,
                           inline-graphic, li-label, lang-US, highlighter, sumaru
- 文字スタイルグループ:
    table -> td-bold
    toc   -> toc-h1, toc-h2, toc-h3
    book  -> page-number, running-head, thumb-index

仕様:
- 同じ場所に同名のスタイル／グループが既に存在する場合は追加せずスキップ
  （ルート直下と各グループ内を分けて判定）
- p に行揃え「均等配置（最終行左揃え）」を毎回上書き設定
- td-left / td-center / td-right に行揃えを毎回上書き設定
- inline-graphic に前後四分のアキ（leadingAki/trailingAki = 0.25）を毎回上書き設定
- lang-US の言語を「英語：米国」に設定
- code-normal の言語を「なし」に設定
- sumaru の noBreak（分割禁止）を有効化

基準スタイル（basedOn）:
- td-left / td-center / td-right → td （table グループ内）
- toc-h2 / toc-h3                → toc-h1 （toc グループ内）
- code-strong                    → code-normal
- highlighter                    → strong-bold
- li-label                       → strong-bold

正規表現スタイル（重複追加しない）:
- ul-li : `^.+?(?=：)`        → li-label
- p     : `[\u\l]`             → lang-US
- p     : `..[。」』？！…]?$`  → sumaru
- p     : `~a`                 → inline-graphic

並び順:
- パネル上の並び順は配列順に揃える（既存スタイルも move() で並び替え）
*/

(function () {
    if (app.documents.length === 0) {
        alert("ドキュメントを開いてから実行してください。");
        return;
    }
    var doc = app.activeDocument;

    // =========================================
    // スタイル名定義 / Style name definitions
    // =========================================

    var paragraphStyleNames = [
        "h1", "h2", "h3", "h4", "h5", "h6",
        "ul-li",
        "p", "p.caption", "p.code", "p.img",
        "ol-li"
    ];

    var characterStyleNames = [
        "strong-bold", "em-italic",
        "link", "code-normal", "code-strong", "inline-graphic", "li-label",
        "lang-US", "highlighter", "sumaru"
    ];

    var paragraphStyleGroupNames = [
        "table", "toc", "book"
    ];

    var paragraphStylesInGroups = [
        { group: "table", styles: ["th", "td", "td-left", "td-center", "td-right"] },
        { group: "toc", styles: ["toc-title", "toc-h1", "toc-h2", "toc-h3"] },
        { group: "book", styles: ["page-number", "running-head", "thumb-index"] }
    ];

    var characterStyleGroupNames = [
        "table", "toc", "book"
    ];

    var characterStylesInGroups = [
        { group: "table", styles: ["td-bold"] },
        { group: "toc", styles: ["toc-h1", "toc-h2", "toc-h3"] },
        { group: "book", styles: ["page-number", "running-head", "thumb-index"] }
    ];

    // =========================================
    // スタイル／グループ作成（既存ならスキップ） / Ensure styles & groups (skip if exists)
    // =========================================

    /* 段落スタイルグループを取得、無ければ作成 / Get a paragraph style group, create if missing */
    function ensureParagraphStyleGroup(doc, groupName) {
        var styleGroup = doc.paragraphStyleGroups.itemByName(groupName);
        if (!styleGroup.isValid) {
            styleGroup = doc.paragraphStyleGroups.add({ name: groupName });
        }
        return styleGroup;
    }

    /* 文字スタイルグループを取得、無ければ作成 / Get a character style group, create if missing */
    function ensureCharacterStyleGroup(doc, groupName) {
        var styleGroup = doc.characterStyleGroups.itemByName(groupName);
        if (!styleGroup.isValid) {
            styleGroup = doc.characterStyleGroups.add({ name: groupName });
        }
        return styleGroup;
    }

    /* 指定コンテナ（ドキュメントまたはグループ）の段落スタイルを取得、無ければ作成 /
       Get a paragraph style in the given container (doc or group), create if missing */
    function ensureParagraphStyle(container, styleName) {
        var paragraphStyle = container.paragraphStyles.itemByName(styleName);
        if (!paragraphStyle.isValid) {
            paragraphStyle = container.paragraphStyles.add({ name: styleName });
        }
        return paragraphStyle;
    }

    /* 指定コンテナの文字スタイルを取得、無ければ作成 /
       Get a character style in the given container, create if missing */
    function ensureCharacterStyle(container, styleName) {
        var characterStyle = container.characterStyles.itemByName(styleName);
        if (!characterStyle.isValid) {
            characterStyle = container.characterStyles.add({ name: styleName });
        }
        return characterStyle;
    }

    // =========================================
    // 個別スタイルの属性適用 / Style-specific property settings
    // =========================================

    /* p に均等配置（最終行左揃え）を設定 / Apply justify-with-last-line-left to "p" */
    function applyBodyParagraphStyleSettings(doc) {
        var bodyParagraphStyle = doc.paragraphStyles.itemByName("p");
        if (bodyParagraphStyle.isValid) {
            bodyParagraphStyle.justification = Justification.LEFT_JUSTIFIED;
        }
    }

    /* table グループ内 td-* の行揃えと td ベースを設定 /
       Set alignment and basedOn=td for td-left/td-center/td-right */
    function applyTableParagraphStyleSettings(doc) {
        var tableGroup = doc.paragraphStyleGroups.itemByName("table");
        if (tableGroup.isValid) {
            var tdBase = tableGroup.paragraphStyles.itemByName("td");
            var tdLeft = tableGroup.paragraphStyles.itemByName("td-left");
            if (tdLeft.isValid) {
                if (tdBase.isValid) tdLeft.basedOn = tdBase;
                tdLeft.justification = Justification.LEFT_ALIGN;
            }
            var tdCenter = tableGroup.paragraphStyles.itemByName("td-center");
            if (tdCenter.isValid) {
                if (tdBase.isValid) tdCenter.basedOn = tdBase;
                tdCenter.justification = Justification.CENTER_ALIGN;
            }
            var tdRight = tableGroup.paragraphStyles.itemByName("td-right");
            if (tdRight.isValid) {
                if (tdBase.isValid) tdRight.basedOn = tdBase;
                tdRight.justification = Justification.RIGHT_ALIGN;
            }
        }
    }

    /* toc グループ内 toc-h2 / toc-h3 を toc-h1 ベースに /
       Set basedOn=toc-h1 for toc-h2 and toc-h3 */
    function applyTocParagraphStyleSettings(doc) {
        var tocGroup = doc.paragraphStyleGroups.itemByName("toc");
        if (tocGroup.isValid) {
            var tocH1 = tocGroup.paragraphStyles.itemByName("toc-h1");
            if (tocH1.isValid) {
                var tocH2 = tocGroup.paragraphStyles.itemByName("toc-h2");
                if (tocH2.isValid) tocH2.basedOn = tocH1;
                var tocH3 = tocGroup.paragraphStyles.itemByName("toc-h3");
                if (tocH3.isValid) tocH3.basedOn = tocH1;
            }
        }
    }

    /* inline-graphic に前後四分のアキを設定 /
       Apply quarter-em leading/trailing spacing to inline-graphic */
    function applyInlineGraphicCharacterStyleSettings(doc) {
        var inlineGraphic = doc.characterStyles.itemByName("inline-graphic");
        if (inlineGraphic.isValid) {
            inlineGraphic.leadingAki = 0.25;
            inlineGraphic.trailingAki = 0.25;
        }
    }

    /* sumaru の分割禁止を有効化 / Enable noBreak on sumaru */
    function applySumaruCharacterStyleSettings(doc) {
        var sumaru = doc.characterStyles.itemByName("sumaru");
        if (sumaru.isValid) {
            sumaru.noBreak = true;
        }
    }

    /* highlighter を strong-bold ベースに / Set basedOn=strong-bold on highlighter */
    function applyHighlighterCharacterStyleSettings(doc) {
        var highlighter = doc.characterStyles.itemByName("highlighter");
        var strongBold = doc.characterStyles.itemByName("strong-bold");
        if (highlighter.isValid && strongBold.isValid) {
            highlighter.basedOn = strongBold;
        }
    }

    /* code-strong を code-normal ベースに / Set basedOn=code-normal on code-strong */
    function applyCodeStrongCharacterStyleSettings(doc) {
        var codeStrong = doc.characterStyles.itemByName("code-strong");
        var codeNormal = doc.characterStyles.itemByName("code-normal");
        if (codeStrong.isValid && codeNormal.isValid) {
            codeStrong.basedOn = codeNormal;
        }
    }

    /* li-label を strong-bold ベースに / Set basedOn=strong-bold on li-label */
    function applyLiLabelCharacterStyleSettings(doc) {
        var liLabel = doc.characterStyles.itemByName("li-label");
        var strongBoldForLi = doc.characterStyles.itemByName("strong-bold");
        if (liLabel.isValid && strongBoldForLi.isValid) {
            liLabel.basedOn = strongBoldForLi;
        }
    }

    /* lang-US の言語を「英語：米国」に（環境差を考慮した候補順） /
       Set applied language to English: USA (try multiple localized names) */
    function applyEnglishCharacterStyleSettings(doc) {
        var langUS = doc.characterStyles.itemByName("lang-US");
        if (langUS.isValid) {
            var langNames = ["English: USA", "英語：米国"];
            for (var languageIndex = 0; languageIndex < langNames.length; languageIndex++) {
                var candidateLanguage = app.languagesWithVendors.itemByName(langNames[languageIndex]);
                if (candidateLanguage.isValid) {
                    langUS.appliedLanguage = candidateLanguage;
                    break;
                }
            }
        }
    }

    /* code-normal の言語を「なし」に（環境差を考慮した候補順） /
       Set applied language to [No Language] (try multiple localized names) */
    function applyCodeNormalCharacterStyleSettings(doc) {
        var codeNormal = doc.characterStyles.itemByName("code-normal");
        if (codeNormal.isValid) {
            var noLangNames = ["[No Language]", "[言語なし]", "[なし]"];
            for (var noLangIndex = 0; noLangIndex < noLangNames.length; noLangIndex++) {
                var noLangCandidate = app.languagesWithVendors.itemByName(noLangNames[noLangIndex]);
                if (noLangCandidate.isValid) {
                    codeNormal.appliedLanguage = noLangCandidate;
                    break;
                }
            }
        }
    }

    /* 個別スタイルの属性適用をまとめて実行 / Run all style-specific settings */
    function applyStyleSettings(doc) {
        applyBodyParagraphStyleSettings(doc);
        applyTableParagraphStyleSettings(doc);
        applyTocParagraphStyleSettings(doc);
        applyInlineGraphicCharacterStyleSettings(doc);
        applyHighlighterCharacterStyleSettings(doc);
        applyCodeStrongCharacterStyleSettings(doc);
        applyLiLabelCharacterStyleSettings(doc);
        applySumaruCharacterStyleSettings(doc);
        applyEnglishCharacterStyleSettings(doc);
        applyCodeNormalCharacterStyleSettings(doc);
    }

    // =========================================
    // 正規表現スタイル（ネスト GREP） / Nested GREP styles
    // =========================================

    /* 段落スタイルに正規表現スタイルを設定（同条件のものがあれば重複追加しない） /
       Add nested GREP styles to paragraph styles (skip duplicates) */
    function applyNestedGrepStyleSettings(doc) {
        var nestedGrepRules = [
            { paragraph: "ul-li", character: "li-label", expression: "^.+?(?=：)" },
            { paragraph: "p", character: "lang-US", expression: "[\\u\\l]" },
            { paragraph: "p", character: "sumaru", expression: "..[。」』？！…]?$" },
            { paragraph: "p", character: "inline-graphic", expression: "~a" }
        ];

        for (var grepRuleIndex = 0; grepRuleIndex < nestedGrepRules.length; grepRuleIndex++) {
            var grepRuleDefinition = nestedGrepRules[grepRuleIndex];
            var targetParagraphStyle = doc.paragraphStyles.itemByName(grepRuleDefinition.paragraph);
            var targetCharacterStyle = doc.characterStyles.itemByName(grepRuleDefinition.character);
            if (!targetParagraphStyle.isValid || !targetCharacterStyle.isValid) continue;
            var hasSameGrepStyle = false;
            for (var nestedGrepStyleIndex = 0; nestedGrepStyleIndex < targetParagraphStyle.nestedGrepStyles.length; nestedGrepStyleIndex++) {
                var existingGrepStyle = targetParagraphStyle.nestedGrepStyles[nestedGrepStyleIndex];
                if (existingGrepStyle.grepExpression === grepRuleDefinition.expression &&
                    existingGrepStyle.appliedCharacterStyle.name === targetCharacterStyle.name) {
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

    /* 段落スタイル・グループを配列順にパネル末尾へ並べ替える /
       Reorder paragraph styles and groups to match the declared array order */
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

    /* 文字スタイル・グループを配列順にパネル末尾へ並べ替える /
       Reorder character styles and groups to match the declared array order */
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

    applyStyleSettings(doc);
    applyNestedGrepStyleSettings(doc);

    // パネル上の並び順を配列順に揃える（既存スタイルも含む） /
    // Reorder styles in the panel (including existing ones)
    reorderParagraphStyles(doc);
    reorderCharacterStyles(doc);
})();
