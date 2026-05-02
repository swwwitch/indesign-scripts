#target indesign

    /*
     * InDesignStyleSetup.jsx
     *
     * アクティブな InDesign ドキュメントに段落スタイル／文字スタイルと
     * それぞれのグループ（フォルダー）を一括登録する。
     *
     * 登録内容:
     *   - 段落スタイル（ルート）: h1〜h6, p, p.caption, p.code, p.img, ol-li, ul-li
     *   - 段落スタイルグループ:
     *       table  -> th, td, td-left, td-center, td-right
     *       toc    -> toc-h1, toc-h2, toc-h3
     *       book   -> page-number, running-head, thumb-index
     *   - 文字スタイル（ルート）: strong-bold, em-italic, link, code-strong,
     *                              inline-graphic, li-label, lang-US
     *   - 文字スタイルグループ:
     *       table  -> td-bold
     *       toc    -> toc-h1, toc-h2, toc-h3
     *       book   -> page-number, running-head, thumb-index
     *
     * 仕様:
     *   - 同じ場所に同名のスタイル／グループが既に存在する場合は追加せずスキップ
     *     （ルート直下と各グループ内を分けて判定）
     *   - p に行揃え「均等配置（最終行左揃え）」を毎回上書き設定
     *   - td-left / td-center / td-right に行揃えを毎回上書き設定
     *   - inline-graphic に前後四分のアキ（leadingAki/trailingAki = 0.25）を毎回上書き設定
     *   - lang-US の言語を「英語：米国」に設定
     *   - 正規表現スタイル（重複追加しない）:
     *       ul-li : `^.+?(?=：)` → li-label
     *       p     : `[\u\l]`     → lang-US
     *   - 最後に menuActions(8511, 8505) を invoke（ソート）
     */

    (function () {
        if (app.documents.length === 0) {
            alert("ドキュメントを開いてから実行してください。");
            return;
        }
        var doc = app.activeDocument;

        var paragraphStyleNames = [
            "h1", "h2", "h3", "h4", "h5", "h6",
            "p", "p.caption", "p.code", "p.img",
            "ol-li", "ul-li"
        ];

        var characterStyleNames = [
            "strong-bold", "em-italic",
            "link", "code-strong", "inline-graphic", "li-label",
            "lang-US"
        ];

        var paragraphStyleGroupNames = [
            "table", "toc", "book"
        ];

        var paragraphStylesInGroups = [
            { group: "table", styles: ["th", "td", "td-left", "td-center", "td-right"] },
            { group: "toc", styles: ["toc-h1", "toc-h2", "toc-h3"] },
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

        function ensureParagraphStyleGroup(doc, groupName) {
            var styleGroup = doc.paragraphStyleGroups.itemByName(groupName);
            if (!styleGroup.isValid) {
                styleGroup = doc.paragraphStyleGroups.add({ name: groupName });
            }
            return styleGroup;
        }

        function ensureCharacterStyleGroup(doc, groupName) {
            var styleGroup = doc.characterStyleGroups.itemByName(groupName);
            if (!styleGroup.isValid) {
                styleGroup = doc.characterStyleGroups.add({ name: groupName });
            }
            return styleGroup;
        }

        function ensureParagraphStyle(container, styleName) {
            var paragraphStyle = container.paragraphStyles.itemByName(styleName);
            if (!paragraphStyle.isValid) {
                paragraphStyle = container.paragraphStyles.add({ name: styleName });
            }
            return paragraphStyle;
        }

        function ensureCharacterStyle(container, styleName) {
            var characterStyle = container.characterStyles.itemByName(styleName);
            if (!characterStyle.isValid) {
                characterStyle = container.characterStyles.add({ name: styleName });
            }
            return characterStyle;
        }

        function applyBodyParagraphStyleSettings(doc) {
            // p: 均等配置（最終行左揃え / Justify with last line aligned left）
            var bodyParagraphStyle = doc.paragraphStyles.itemByName("p");
            if (bodyParagraphStyle.isValid) {
                bodyParagraphStyle.justification = Justification.LEFT_JUSTIFIED;
            }
        }

        function applyTableParagraphStyleSettings(doc) {
            // table グループ内 td-* の行揃え
            var tableGroup = doc.paragraphStyleGroups.itemByName("table");
            if (tableGroup.isValid) {
                var tdLeft = tableGroup.paragraphStyles.itemByName("td-left");
                if (tdLeft.isValid) tdLeft.justification = Justification.LEFT_ALIGN;
                var tdCenter = tableGroup.paragraphStyles.itemByName("td-center");
                if (tdCenter.isValid) tdCenter.justification = Justification.CENTER_ALIGN;
                var tdRight = tableGroup.paragraphStyles.itemByName("td-right");
                if (tdRight.isValid) tdRight.justification = Justification.RIGHT_ALIGN;
            }
        }

        function applyInlineGraphicCharacterStyleSettings(doc) {
            // inline-graphic: 前後四分のアキ
            var inlineGraphic = doc.characterStyles.itemByName("inline-graphic");
            if (inlineGraphic.isValid) {
                inlineGraphic.leadingAki = 0.25;
                inlineGraphic.trailingAki = 0.25;
            }
        }

        function applyEnglishCharacterStyleSettings(doc) {
            // lang-US: 言語を「英語：米国」に（環境によって名称が違うので候補を順に試す）
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

        function applyStyleSettings(doc) {
            applyBodyParagraphStyleSettings(doc);
            applyTableParagraphStyleSettings(doc);
            applyInlineGraphicCharacterStyleSettings(doc);
            applyEnglishCharacterStyleSettings(doc);
        }

        function applyNestedGrepStyleSettings(doc) {
            // 正規表現スタイル（重複追加しない）
            //   ul-li : 「：」直前までを li-label
            //   p     : 半角英字（\u\l）を lang-US
            var nestedGrepRules = [
                { paragraph: "ul-li", character: "li-label", expression: "^.+?(?=：)" },
                { paragraph: "p", character: "lang-US", expression: "[\\u\\l]" }
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

        function invokeMenuActionByID(actionID) {
            var menuAction = app.menuActions.itemByID(actionID);
            if (menuAction.isValid) {
                menuAction.invoke();
            }
        }

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

        // UI更新・再描画（InDesign の内部メニューコマンドID）
        invokeMenuActionByID(8511);
        invokeMenuActionByID(8505);
    })();
