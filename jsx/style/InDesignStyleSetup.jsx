#target indesign

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.1.0";

// =========================================
// 動作スイッチ / Behavior switches
// =========================================

/* true: 既存スタイルの属性（行揃え／言語／basedOn など）も上書きする /
   false: 既存スタイルには触れず、今回新規作成したスタイルにのみ属性を適用する */
var OVERWRITE_EXISTING_STYLES = false;

/* true: 各スタイルに基準スタイル（basedOn）を設定する /
   false: basedOn は設定しない（その他の属性は対象） */
var APPLY_BASED_ON_RELATIONSHIPS = true;

/*
概要

アクティブな InDesign ドキュメントに段落スタイル／文字スタイルと
それぞれのグループ（フォルダー）を一括登録する。

登録内容:
- 段落スタイル（ルート）: h1〜h6, ul-li, ol-li, p, p.caption, p.code, p.img
- 段落スタイルグループ:
    table -> th, th-left, th-center, th-right, td, td-left, td-center, td-right
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
- h1〜h6 に行揃え「左揃え」を毎回上書き設定
- p に行揃え「均等配置（最終行左揃え）」を毎回上書き設定
- th-left / th-center / th-right に行揃えを毎回上書き設定
- td-left / td-center / td-right に行揃えを毎回上書き設定
- inline-graphic に前後四分のアキ（leadingAki/trailingAki = 0.25）を毎回上書き設定
- lang-US の言語を「英語：米国」に設定
- code-normal の言語を「なし」に設定
- sumaru の noBreak（分割禁止）を有効化

基準スタイル（basedOn）:
- th-left / th-center / th-right → th （table グループ内）
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

動作スイッチ:
- OVERWRITE_EXISTING_STYLES: 既存スタイルの属性も上書きするか（既定 false）
- APPLY_BASED_ON_RELATIONSHIPS: 基準スタイル（basedOn）を設定するか（既定 true）
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
        "ul-li", "ol-li",
        "p", "p.caption", "p.code", "p.img"
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
        { group: "table", styles: ["th", "th-left", "th-center", "th-right", "td", "td-left", "td-center", "td-right"] },
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

    /* 今回のスクリプト実行で新規作成したスタイル名を記録（属性適用の対象判定に使う） /
       Names of paragraph/character styles created during this run (used by attribute guards) */
    var paragraphStyleNamesCreatedThisRun = {};
    var characterStyleNamesCreatedThisRun = {};

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
    function ensureParagraphStyle(styleContainer, styleName) {
        var paragraphStyle = styleContainer.paragraphStyles.itemByName(styleName);
        if (!paragraphStyle.isValid) {
            paragraphStyle = styleContainer.paragraphStyles.add({ name: styleName });
            paragraphStyleNamesCreatedThisRun[styleName] = true;
        }
        return paragraphStyle;
    }

    /* 指定コンテナの文字スタイルを取得、無ければ作成 /
       Get a character style in the given container, create if missing */
    function ensureCharacterStyle(styleContainer, styleName) {
        var characterStyle = styleContainer.characterStyles.itemByName(styleName);
        if (!characterStyle.isValid) {
            characterStyle = styleContainer.characterStyles.add({ name: styleName });
            characterStyleNamesCreatedThisRun[styleName] = true;
        }
        return characterStyle;
    }

    /* 段落スタイルへ属性を適用してよいか判定（新規作成 or 上書きスイッチ ON） /
       Whether to apply attributes to a paragraph style (newly created OR overwrite switch on) */
    function shouldApplyAttributesToParagraphStyle(styleName) {
        return OVERWRITE_EXISTING_STYLES || paragraphStyleNamesCreatedThisRun[styleName] === true;
    }

    /* 文字スタイルへ属性を適用してよいか判定 /
       Whether to apply attributes to a character style */
    function shouldApplyAttributesToCharacterStyle(styleName) {
        return OVERWRITE_EXISTING_STYLES || characterStyleNamesCreatedThisRun[styleName] === true;
    }

    // =========================================
    // 個別スタイルの属性適用 / Style-specific property settings
    // =========================================
    // ※ 既存スタイルへの上書きは OVERWRITE_EXISTING_STYLES、basedOn は APPLY_BASED_ON_RELATIONSHIPS で制御。
    //   shouldApplyAttributesTo* がガード。
    //   Guarded by OVERWRITE_EXISTING_STYLES (overwrite) and APPLY_BASED_ON_RELATIONSHIPS (basedOn).

    /* p に均等配置（最終行左揃え）を設定 / Apply justify-with-last-line-left to "p" */
    function applyParagraphJustificationForP(doc) {
        if (!shouldApplyAttributesToParagraphStyle("p")) return;
        var bodyParagraphStyle = doc.paragraphStyles.itemByName("p");
        if (bodyParagraphStyle.isValid) {
            bodyParagraphStyle.justification = Justification.LEFT_JUSTIFIED;
        }
    }

    /* h1〜h6 に左揃えを設定 / Apply left alignment to h1–h6 */
    function applyHeadingJustification(doc) {
        var headingStyleNames = ["h1", "h2", "h3", "h4", "h5", "h6"];
        for (var headingIndex = 0; headingIndex < headingStyleNames.length; headingIndex++) {
            var headingName = headingStyleNames[headingIndex];
            if (!shouldApplyAttributesToParagraphStyle(headingName)) continue;
            var headingStyle = doc.paragraphStyles.itemByName(headingName);
            if (headingStyle.isValid) {
                headingStyle.justification = Justification.LEFT_ALIGN;
            }
        }
    }

    /* table グループ内 th-* / td-* の行揃えと th / td ベースを設定 /
       Set alignment and basedOn=th/td for th-* and td-* */
    function applyTableCellAlignmentAndBasedOn(doc) {
        var tableGroup = doc.paragraphStyleGroups.itemByName("table");
        if (!tableGroup.isValid) return;

        var thBaseStyle = tableGroup.paragraphStyles.itemByName("th");
        var thAlignmentTargets = [
            { name: "th-left", justification: Justification.LEFT_ALIGN },
            { name: "th-center", justification: Justification.CENTER_ALIGN },
            { name: "th-right", justification: Justification.RIGHT_ALIGN }
        ];
        for (var thIndex = 0; thIndex < thAlignmentTargets.length; thIndex++) {
            var thTarget = thAlignmentTargets[thIndex];
            if (!shouldApplyAttributesToParagraphStyle(thTarget.name)) continue;
            var thTargetStyle = tableGroup.paragraphStyles.itemByName(thTarget.name);
            if (!thTargetStyle.isValid) continue;
            if (APPLY_BASED_ON_RELATIONSHIPS && thBaseStyle.isValid) {
                thTargetStyle.basedOn = thBaseStyle;
            }
            thTargetStyle.justification = thTarget.justification;
        }

        var tdBaseStyle = tableGroup.paragraphStyles.itemByName("td");
        var tdAlignmentTargets = [
            { name: "td-left", justification: Justification.LEFT_ALIGN },
            { name: "td-center", justification: Justification.CENTER_ALIGN },
            { name: "td-right", justification: Justification.RIGHT_ALIGN }
        ];
        for (var tdIndex = 0; tdIndex < tdAlignmentTargets.length; tdIndex++) {
            var tdTarget = tdAlignmentTargets[tdIndex];
            if (!shouldApplyAttributesToParagraphStyle(tdTarget.name)) continue;
            var tdTargetStyle = tableGroup.paragraphStyles.itemByName(tdTarget.name);
            if (!tdTargetStyle.isValid) continue;
            if (APPLY_BASED_ON_RELATIONSHIPS && tdBaseStyle.isValid) {
                tdTargetStyle.basedOn = tdBaseStyle;
            }
            tdTargetStyle.justification = tdTarget.justification;
        }
    }

    /* toc グループ内 toc-h2 / toc-h3 を toc-h1 ベースに /
       Set basedOn=toc-h1 for toc-h2 and toc-h3 */
    function applyTocSubheadingBasedOn(doc) {
        if (!APPLY_BASED_ON_RELATIONSHIPS) return;
        var tocGroup = doc.paragraphStyleGroups.itemByName("toc");
        if (!tocGroup.isValid) return;
        var tocH1Style = tocGroup.paragraphStyles.itemByName("toc-h1");
        if (!tocH1Style.isValid) return;

        var tocSubheadingNames = ["toc-h2", "toc-h3"];
        for (var tocIndex = 0; tocIndex < tocSubheadingNames.length; tocIndex++) {
            var tocSubheadingName = tocSubheadingNames[tocIndex];
            if (!shouldApplyAttributesToParagraphStyle(tocSubheadingName)) continue;
            var tocSubheadingStyle = tocGroup.paragraphStyles.itemByName(tocSubheadingName);
            if (tocSubheadingStyle.isValid) tocSubheadingStyle.basedOn = tocH1Style;
        }
    }

    /* inline-graphic に前後四分のアキを設定 /
       Apply quarter-em leading/trailing spacing to inline-graphic */
    function applyInlineGraphicSpacing(doc) {
        if (!shouldApplyAttributesToCharacterStyle("inline-graphic")) return;
        var inlineGraphicStyle = doc.characterStyles.itemByName("inline-graphic");
        if (inlineGraphicStyle.isValid) {
            inlineGraphicStyle.leadingAki = 0.25;
            inlineGraphicStyle.trailingAki = 0.25;
        }
    }

    /* sumaru の分割禁止を有効化 / Enable noBreak on sumaru */
    function applySumaruNoBreak(doc) {
        if (!shouldApplyAttributesToCharacterStyle("sumaru")) return;
        var sumaruStyle = doc.characterStyles.itemByName("sumaru");
        if (sumaruStyle.isValid) {
            sumaruStyle.noBreak = true;
        }
    }

    /* 1 つの文字スタイルの basedOn を別の文字スタイルに設定する共通処理 /
       Common helper: set basedOn of one character style to another */
    function applyCharacterStyleBasedOn(doc, targetStyleName, parentStyleName) {
        if (!APPLY_BASED_ON_RELATIONSHIPS) return;
        if (!shouldApplyAttributesToCharacterStyle(targetStyleName)) return;
        var targetStyle = doc.characterStyles.itemByName(targetStyleName);
        var parentStyle = doc.characterStyles.itemByName(parentStyleName);
        if (targetStyle.isValid && parentStyle.isValid) {
            targetStyle.basedOn = parentStyle;
        }
    }

    /* lang-US の言語を「英語：米国」に（環境差を考慮した候補順） /
       Set applied language to English: USA (try multiple localized names) */
    function applyLangUSLanguageSetting(doc) {
        if (!shouldApplyAttributesToCharacterStyle("lang-US")) return;
        var langUSStyle = doc.characterStyles.itemByName("lang-US");
        if (!langUSStyle.isValid) return;
        var englishLanguageNames = ["English: USA", "英語：米国"];
        for (var englishNameIndex = 0; englishNameIndex < englishLanguageNames.length; englishNameIndex++) {
            var englishLanguage = app.languagesWithVendors.itemByName(englishLanguageNames[englishNameIndex]);
            if (englishLanguage.isValid) {
                langUSStyle.appliedLanguage = englishLanguage;
                break;
            }
        }
    }

    /* code-normal の言語を「なし」に（環境差を考慮した候補順） /
       Set applied language to [No Language] (try multiple localized names) */
    function applyCodeNormalLanguageSetting(doc) {
        if (!shouldApplyAttributesToCharacterStyle("code-normal")) return;
        var codeNormalStyle = doc.characterStyles.itemByName("code-normal");
        if (!codeNormalStyle.isValid) return;
        var noLanguageNames = ["[No Language]", "[言語なし]", "[なし]"];
        for (var noLanguageNameIndex = 0; noLanguageNameIndex < noLanguageNames.length; noLanguageNameIndex++) {
            var noLanguageEntry = app.languagesWithVendors.itemByName(noLanguageNames[noLanguageNameIndex]);
            if (noLanguageEntry.isValid) {
                codeNormalStyle.appliedLanguage = noLanguageEntry;
                break;
            }
        }
    }

    /* 個別スタイルの属性適用をまとめて実行 / Run all style-specific settings */
    function applyAllStyleAttributes(doc) {
        applyHeadingJustification(doc);
        applyParagraphJustificationForP(doc);
        applyTableCellAlignmentAndBasedOn(doc);
        applyTocSubheadingBasedOn(doc);
        applyInlineGraphicSpacing(doc);
        applyCharacterStyleBasedOn(doc, "highlighter", "strong-bold");
        applyCharacterStyleBasedOn(doc, "code-strong", "code-normal");
        applyCharacterStyleBasedOn(doc, "li-label", "strong-bold");
        applySumaruNoBreak(doc);
        applyLangUSLanguageSetting(doc);
        applyCodeNormalLanguageSetting(doc);
    }

    // =========================================
    // 正規表現スタイル（ネスト GREP） / Nested GREP styles
    // =========================================

    /* 段落スタイルに正規表現スタイルを設定（同条件のものがあれば重複追加しない） /
       Add nested GREP styles to paragraph styles (skip duplicates).
       既存段落スタイルに対する追加は OVERWRITE_EXISTING_STYLES で制御 /
       Adding to pre-existing paragraph styles is gated by OVERWRITE_EXISTING_STYLES. */
    function applyNestedGrepStyleSettings(doc) {
        var nestedGrepRules = [
            { paragraph: "ul-li", character: "li-label", expression: "^.+?(?=：)" },
            { paragraph: "p", character: "lang-US", expression: "[\\u\\l]" },
            { paragraph: "p", character: "sumaru", expression: "..[。」』？！…]?$" },
            { paragraph: "p", character: "inline-graphic", expression: "~a" }
        ];

        for (var grepRuleIndex = 0; grepRuleIndex < nestedGrepRules.length; grepRuleIndex++) {
            var grepRuleDefinition = nestedGrepRules[grepRuleIndex];
            if (!shouldApplyAttributesToParagraphStyle(grepRuleDefinition.paragraph)) continue;
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

    applyAllStyleAttributes(doc);
    applyNestedGrepStyleSettings(doc);

    // パネル上の並び順を配列順に揃える（既存スタイルも含む） /
    // Reorder styles in the panel (including existing ones)
    reorderParagraphStyles(doc);
    reorderCharacterStyles(doc);
})();
