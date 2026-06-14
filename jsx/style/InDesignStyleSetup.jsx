#target indesign

/*

### 概要

アクティブな InDesign ドキュメントに、段落スタイル／文字スタイルとそれぞれのグループ（フォルダー）を一括登録し、見出し・本文・表・目次・コードなどの定番属性をまとめて設定する。

- 同名のスタイル／グループが既にあれば作成はスキップ（重複登録しない）
- 属性の適用は動作スイッチで制御： `OVERWRITE_EXISTING_STYLES`（既存スタイルにも上書きするか。既定 true）/ `APPLY_BASED_ON_RELATIONSHIPS`（基準スタイルを設定するか。既定 true）
- 共通の正規表現スタイル（GREP）は basestyle グループの base-regex に集約し、p / h1〜h6 / ul-li / ol-li / p.caption を base-regex 基準にして継承させる
- パネル上の並び順は配列の宣言順に揃える（既存スタイルも move() で並び替え）
- 全処理は `app.doScript`（UndoModes.ENTIRE_SCRIPT）で 1 つのアンドゥ単位にまとめる

### 登録内容

- 段落スタイル（ルート）: h1〜h6, ul-li, ol-li, p, p.caption, p.code, p.img
- 段落スタイルグループ
  - basestyle: base-regex
  - table: th, th-left, th-center, th-right, td, td-left, td-center, td-right
  - toc: toc-title, toc-h1, toc-h2, toc-h3
  - book: page-number, running-head, thumb-index
- 文字スタイル（ルート）: strong-bold, em-italic, link, code-normal, code-strong, inline-graphic, li-label, lang-US, highlighter, sumaru
- 文字スタイルグループ
  - table: td-bold
  - toc: toc-h1, toc-h2, toc-h3
  - book: page-number, running-head, thumb-index

### 仕様

同じ場所に同名のスタイル／グループが既にあれば追加せずスキップ（ルート直下と各グループ内を分けて判定）。
以下の属性は `OVERWRITE_EXISTING_STYLES` に従って適用（既定 true なら既存スタイルにも毎回上書き、false なら今回新規作成分のみ）。

**行揃え**

- h1〜h6 / toc-h1〜toc-h3: 左揃え
- p: 均等配置（最終行左揃え）
- th-left / th-center / th-right: 左 / 中央 / 右
- td-left / td-center / td-right: 左 / 中央 / 右
- p.code: 左

**段落分離禁止（Keep Options）**

- h1〜h6 / toc-h1〜toc-h3: 次の段落を保持「2行」＋ すべての行を分離禁止
- ul-li / p.caption: 前の段落から分離しない
- ol-li / ul-li / p.caption / p.code: すべての行を分離禁止
- th / td: すべての行を分離禁止（th-* / td-* は basedOn で継承）

**次のスタイル**

- h1〜h6 → p（toc-h1〜toc-h3 は対象外）

**言語・欧文合字・ハイフネーション**

- lang-US: 言語「英語：米国」
- code-normal: 言語「なし」
- p.code: 言語「なし」・欧文合字オフ・ハイフネーションオフ

**その他**

- inline-graphic: 前後四分のアキ（leadingAki / trailingAki = 0.25）
- sumaru: noBreak（分割禁止）

### 基準スタイル（basedOn）

- th-left / th-center / th-right → th（table グループ内）
- td-left / td-center / td-right → td（table グループ内）
- toc-h2 / toc-h3 → toc-h1（toc グループ内）
- code-strong → code-normal
- highlighter → strong-bold
- li-label → strong-bold
- p / h1〜h6 / ul-li / ol-li / p.caption → base-regex（basestyle グループ内。base-regex の GREP を継承）

### 正規表現スタイル（重複追加しない）

- ul-li : `^.+?(?=：)` → li-label
- base-regex（basestyle）: `[\u\l]` → lang-US
- base-regex（basestyle）: `..[。」』？！…]?$` → sumaru
- base-regex（basestyle）: `~a` → inline-graphic

### 変更履歴

v1.2.0
- basestyle グループと base-regex スタイルを追加。共通 GREP（lang-US / sumaru / inline-graphic）を base-regex に集約し、p / h1〜h6 / ul-li / ol-li / p.caption を base-regex 基準にして継承
- h1〜h6 / toc-h1〜toc-h3 に段落分離禁止（次の段落を保持2行・すべての行を分離禁止）を設定。h1〜h6 の「次のスタイル」を p に設定
- ul-li / p.caption に「前の段落から分離しない」、ol-li / ul-li / p.caption / p.code に「すべての行を分離禁止」、th / td に「すべての行を分離禁止」（th-* / td-* は basedOn 継承）を設定
- p.code に 言語「なし」・欧文合字オフ・行揃え「左」・ハイフネーションオフ を設定
- OVERWRITE_EXISTING_STYLES の既定を false → true に変更（既存同名スタイルも上書き）
- 全処理を app.doScript（UndoModes.ENTIRE_SCRIPT）で 1 アンドゥにまとめ、全体を IIFE 化
- バグ修正: 「すべての行を分離禁止」が効かない問題（マスターの keepLinesTogether を併設）
- 堅牢化: 作成済み判定を container.id + styleName の複合キーに、ネスト GREP の重複判定を id 比較に
- リファクタ: 言語解決を resolveLanguageByNames() に共通化、関数名の整理

*/

(function () {

    // =========================================
    // バージョン / Version
    // =========================================

    var SCRIPT_VERSION = "v1.2.0";

    // =========================================
    // 動作スイッチ / Behavior switches
    // =========================================

    /* true: 既存スタイルの属性（行揃え／言語／basedOn など）も上書きする /
       false: 既存スタイルには触れず、今回新規作成したスタイルにのみ属性を適用する */
    var OVERWRITE_EXISTING_STYLES = true;

    /* true: 各スタイルに基準スタイル（basedOn）を設定する /
       false: basedOn は設定しない（その他の属性は対象） */
    var APPLY_BASED_ON_RELATIONSHIPS = true;

    // =========================================
    // メイン処理 / Main
    // =========================================

    function main() {
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
            "basestyle", "table", "toc", "book"
        ];

        var paragraphStylesInGroups = [
            { group: "basestyle", styles: ["base-regex"] },
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

        /* 今回のスクリプト実行で新規作成したスタイルを記録（属性適用の対象判定に使う）/
           コンテナ（doc またはグループ）ごとに区別するため container.id + styleName を複合キーにする /
           Record styles created during this run (used by attribute guards).
           Keyed by container.id + styleName so same-named styles in different groups don't collide. */
        var paragraphStyleKeysCreatedThisRun = {};
        var characterStyleKeysCreatedThisRun = {};

        /* コンテナ（doc またはグループ）とスタイル名から複合キーを作る /
           Build a composite key from a container (doc or group) and a style name */
        function styleContainerKey(styleContainer, styleName) {
            return styleContainer.id + "\t" + styleName;
        }

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
                paragraphStyleKeysCreatedThisRun[styleContainerKey(styleContainer, styleName)] = true;
            }
            return paragraphStyle;
        }

        /* 指定コンテナの文字スタイルを取得、無ければ作成 /
           Get a character style in the given container, create if missing */
        function ensureCharacterStyle(styleContainer, styleName) {
            var characterStyle = styleContainer.characterStyles.itemByName(styleName);
            if (!characterStyle.isValid) {
                characterStyle = styleContainer.characterStyles.add({ name: styleName });
                characterStyleKeysCreatedThisRun[styleContainerKey(styleContainer, styleName)] = true;
            }
            return characterStyle;
        }

        /* 段落スタイルへ属性を適用してよいか判定（新規作成 or 上書きスイッチ ON） /
           Whether to apply attributes to a paragraph style (newly created OR overwrite switch on) */
        function shouldApplyAttributesToParagraphStyle(styleContainer, styleName) {
            return OVERWRITE_EXISTING_STYLES ||
                paragraphStyleKeysCreatedThisRun[styleContainerKey(styleContainer, styleName)] === true;
        }

        /* 文字スタイルへ属性を適用してよいか判定 /
           Whether to apply attributes to a character style */
        function shouldApplyAttributesToCharacterStyle(styleContainer, styleName) {
            return OVERWRITE_EXISTING_STYLES ||
                characterStyleKeysCreatedThisRun[styleContainerKey(styleContainer, styleName)] === true;
        }

        // =========================================
        // 個別スタイルの属性適用 / Style-specific property settings
        // =========================================
        // ※ 既存スタイルへの上書きは OVERWRITE_EXISTING_STYLES、basedOn は APPLY_BASED_ON_RELATIONSHIPS で制御。
        //   shouldApplyAttributesTo* がガード。
        //   Guarded by OVERWRITE_EXISTING_STYLES (overwrite) and APPLY_BASED_ON_RELATIONSHIPS (basedOn).

        /* p（本文）に均等配置（最終行左揃え）を設定 / Apply justify-with-last-line-left to "p" (body) */
        function applyBodyParagraphJustification(doc) {
            if (!shouldApplyAttributesToParagraphStyle(doc, "p")) return;
            var bodyParagraphStyle = doc.paragraphStyles.itemByName("p");
            if (bodyParagraphStyle.isValid) {
                bodyParagraphStyle.justification = Justification.LEFT_JUSTIFIED;
            }
        }

        /* ルート段落スタイル（p / h1〜h6 / ul-li / ol-li / p.caption）の基準を
           basestyle グループの base-regex に設定（base-regex の GREP などを継承）/
           Set basedOn = "base-regex" (basestyle group) for p, h1–h6, ul-li, ol-li, p.caption */
        function applyBaseStyleBasedOn(doc) {
            if (!APPLY_BASED_ON_RELATIONSHIPS) return;
            var baseGroup = doc.paragraphStyleGroups.itemByName("basestyle");
            if (!baseGroup.isValid) return;
            var baseStyle = baseGroup.paragraphStyles.itemByName("base-regex");
            if (!baseStyle.isValid) return;

            var basedOnBaseStyleNames = ["p", "h1", "h2", "h3", "h4", "h5", "h6", "ul-li", "ol-li", "p.caption"];
            for (var basedOnIndex = 0; basedOnIndex < basedOnBaseStyleNames.length; basedOnIndex++) {
                var basedOnStyleName = basedOnBaseStyleNames[basedOnIndex];
                if (!shouldApplyAttributesToParagraphStyle(doc, basedOnStyleName)) continue;
                var basedOnTargetStyle = doc.paragraphStyles.itemByName(basedOnStyleName);
                if (basedOnTargetStyle.isValid) basedOnTargetStyle.basedOn = baseStyle;
            }
        }

        /* 見出し系スタイルに左揃え・「次の段落を保持：2行」・「すべての行を分離禁止」を設定 /
           （nextStyle が渡された場合は「次のスタイル」も設定）/
           Apply left alignment, keep-with-next (2 lines) and keep-all-lines-together
           (and the next style when provided) to the given heading-like styles */
        function applyHeadingLikeSettings(styleContainer, headingStyleNames, nextStyle) {
            for (var headingIndex = 0; headingIndex < headingStyleNames.length; headingIndex++) {
                var headingName = headingStyleNames[headingIndex];
                if (!shouldApplyAttributesToParagraphStyle(styleContainer, headingName)) continue;
                var headingStyle = styleContainer.paragraphStyles.itemByName(headingName);
                if (headingStyle.isValid) {
                    headingStyle.justification = Justification.LEFT_ALIGN;
                    headingStyle.keepWithNext = 2;
                    headingStyle.keepLinesTogether = true;
                    headingStyle.keepAllLinesTogether = true;
                    if (nextStyle && nextStyle.isValid) {
                        headingStyle.nextStyle = nextStyle;
                    }
                }
            }
        }

        /* h1〜h6 と toc グループ内 toc-h1〜toc-h3 に見出し系設定を適用 /
           （h1〜h6 は「次のスタイル」を p に設定。toc 系は対象外）/
           Apply heading-like settings to h1–h6 (next style = p) and toc-h1–toc-h3 (toc group) */
        function applyHeadingSettings(doc) {
            var bodyParagraphStyle = doc.paragraphStyles.itemByName("p");
            applyHeadingLikeSettings(doc, ["h1", "h2", "h3", "h4", "h5", "h6"], bodyParagraphStyle);
            var tocGroup = doc.paragraphStyleGroups.itemByName("toc");
            if (tocGroup.isValid) {
                applyHeadingLikeSettings(tocGroup, ["toc-h1", "toc-h2", "toc-h3"], null);
            }
        }

        /* 段落分離禁止オプションを設定 /
           Apply keep options:
           - 前の段落から分離しない（泣き別れ禁止）: ul-li / p.caption /
             keep-with-previous: ul-li, p.caption
           - すべての行を分離禁止: ol-li / ul-li / p.caption / p.code /
             keep-all-lines-together: ol-li, ul-li, p.caption, p.code */
        function applyKeepTogetherSettings(doc) {
            var keepWithPreviousStyleNames = ["ul-li", "p.caption"];
            for (var keepWithPreviousIndex = 0; keepWithPreviousIndex < keepWithPreviousStyleNames.length; keepWithPreviousIndex++) {
                var keepWithPreviousName = keepWithPreviousStyleNames[keepWithPreviousIndex];
                if (!shouldApplyAttributesToParagraphStyle(doc, keepWithPreviousName)) continue;
                var keepWithPreviousStyle = doc.paragraphStyles.itemByName(keepWithPreviousName);
                if (keepWithPreviousStyle.isValid) {
                    keepWithPreviousStyle.keepWithPrevious = true;
                }
            }

            var keepAllLinesStyleNames = ["ol-li", "ul-li", "p.caption", "p.code"];
            for (var keepAllLinesIndex = 0; keepAllLinesIndex < keepAllLinesStyleNames.length; keepAllLinesIndex++) {
                var keepAllLinesName = keepAllLinesStyleNames[keepAllLinesIndex];
                if (!shouldApplyAttributesToParagraphStyle(doc, keepAllLinesName)) continue;
                var keepAllLinesStyle = doc.paragraphStyles.itemByName(keepAllLinesName);
                if (keepAllLinesStyle.isValid) {
                    keepAllLinesStyle.keepLinesTogether = true;
                    keepAllLinesStyle.keepAllLinesTogether = true;
                }
            }
        }

        /* table グループ内 th / td に「すべての行を分離禁止」、th-* / td-* に行揃えと th / td ベースを設定 /
           Set keep-all-lines-together on th/td, plus alignment and basedOn=th/td for th-* and td-* */
        function applyTableCellSettings(doc) {
            var tableGroup = doc.paragraphStyleGroups.itemByName("table");
            if (!tableGroup.isValid) return;

            // th / td に「すべての行を分離禁止」（th-* / td-* は basedOn で継承）/
            // keep-all-lines-together on th/td (variants inherit via basedOn)
            var tableBaseStyleNames = ["th", "td"];
            for (var tableBaseIndex = 0; tableBaseIndex < tableBaseStyleNames.length; tableBaseIndex++) {
                var tableBaseName = tableBaseStyleNames[tableBaseIndex];
                if (!shouldApplyAttributesToParagraphStyle(tableGroup, tableBaseName)) continue;
                var tableBaseStyle = tableGroup.paragraphStyles.itemByName(tableBaseName);
                if (tableBaseStyle.isValid) {
                    tableBaseStyle.keepLinesTogether = true;
                    tableBaseStyle.keepAllLinesTogether = true;
                }
            }

            var thBaseStyle = tableGroup.paragraphStyles.itemByName("th");
            var thAlignmentTargets = [
                { name: "th-left", justification: Justification.LEFT_ALIGN },
                { name: "th-center", justification: Justification.CENTER_ALIGN },
                { name: "th-right", justification: Justification.RIGHT_ALIGN }
            ];
            for (var thIndex = 0; thIndex < thAlignmentTargets.length; thIndex++) {
                var thTarget = thAlignmentTargets[thIndex];
                if (!shouldApplyAttributesToParagraphStyle(tableGroup, thTarget.name)) continue;
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
                if (!shouldApplyAttributesToParagraphStyle(tableGroup, tdTarget.name)) continue;
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
                if (!shouldApplyAttributesToParagraphStyle(tocGroup, tocSubheadingName)) continue;
                var tocSubheadingStyle = tocGroup.paragraphStyles.itemByName(tocSubheadingName);
                if (tocSubheadingStyle.isValid) tocSubheadingStyle.basedOn = tocH1Style;
            }
        }

        /* inline-graphic に前後四分のアキを設定 /
           Apply quarter-em leading/trailing spacing to inline-graphic */
        function applyInlineGraphicSpacing(doc) {
            if (!shouldApplyAttributesToCharacterStyle(doc, "inline-graphic")) return;
            var inlineGraphicStyle = doc.characterStyles.itemByName("inline-graphic");
            if (inlineGraphicStyle.isValid) {
                inlineGraphicStyle.leadingAki = 0.25;
                inlineGraphicStyle.trailingAki = 0.25;
            }
        }

        /* sumaru の分割禁止を有効化 / Enable noBreak on sumaru */
        function applySumaruNoBreak(doc) {
            if (!shouldApplyAttributesToCharacterStyle(doc, "sumaru")) return;
            var sumaruStyle = doc.characterStyles.itemByName("sumaru");
            if (sumaruStyle.isValid) {
                sumaruStyle.noBreak = true;
            }
        }

        /* 1 つの文字スタイルの basedOn を別の文字スタイルに設定する共通処理 /
           Common helper: set basedOn of one character style to another */
        function applyCharacterStyleBasedOn(doc, targetStyleName, parentStyleName) {
            if (!APPLY_BASED_ON_RELATIONSHIPS) return;
            if (!shouldApplyAttributesToCharacterStyle(doc, targetStyleName)) return;
            var targetStyle = doc.characterStyles.itemByName(targetStyleName);
            var parentStyle = doc.characterStyles.itemByName(parentStyleName);
            if (targetStyle.isValid && parentStyle.isValid) {
                targetStyle.basedOn = parentStyle;
            }
        }

        /* 候補名（環境の UI 言語差を吸収）から言語を解決。見つからなければ null /
           Resolve a language by trying candidate names; returns null if none match */
        function resolveLanguageByNames(languageNames) {
            for (var languageNameIndex = 0; languageNameIndex < languageNames.length; languageNameIndex++) {
                var languageEntry = app.languagesWithVendors.itemByName(languageNames[languageNameIndex]);
                if (languageEntry.isValid) return languageEntry;
            }
            return null;
        }

        var ENGLISH_USA_LANGUAGE_NAMES = ["English: USA", "英語：米国"];
        var NO_LANGUAGE_NAMES = ["[No Language]", "[言語なし]", "[なし]"];

        /* lang-US の言語を「英語：米国」に / Set applied language to English: USA */
        function applyLangUSLanguageSetting(doc) {
            if (!shouldApplyAttributesToCharacterStyle(doc, "lang-US")) return;
            var langUSStyle = doc.characterStyles.itemByName("lang-US");
            if (!langUSStyle.isValid) return;
            var englishLanguage = resolveLanguageByNames(ENGLISH_USA_LANGUAGE_NAMES);
            if (englishLanguage) langUSStyle.appliedLanguage = englishLanguage;
        }

        /* code-normal の言語を「なし」に / Set applied language to [No Language] */
        function applyCodeNormalLanguageSetting(doc) {
            if (!shouldApplyAttributesToCharacterStyle(doc, "code-normal")) return;
            var codeNormalStyle = doc.characterStyles.itemByName("code-normal");
            if (!codeNormalStyle.isValid) return;
            var noLanguage = resolveLanguageByNames(NO_LANGUAGE_NAMES);
            if (noLanguage) codeNormalStyle.appliedLanguage = noLanguage;
        }

        /* p.code に 言語なし・欧文合字オフ・左揃え・ハイフネーションオフ を設定 /
           Apply [No Language], ligatures off, left alignment and hyphenation off to p.code */
        function applyCodeParagraphSettings(doc) {
            if (!shouldApplyAttributesToParagraphStyle(doc, "p.code")) return;
            var codeParagraphStyle = doc.paragraphStyles.itemByName("p.code");
            if (!codeParagraphStyle.isValid) return;
            var noLanguage = resolveLanguageByNames(NO_LANGUAGE_NAMES);
            if (noLanguage) codeParagraphStyle.appliedLanguage = noLanguage;
            codeParagraphStyle.ligatures = false;
            codeParagraphStyle.justification = Justification.LEFT_ALIGN;
            codeParagraphStyle.hyphenation = false;
        }

        /* 個別スタイルの属性適用をまとめて実行 / Run all style-specific settings */
        function applyAllStyleAttributes(doc) {
            applyHeadingSettings(doc);
            applyKeepTogetherSettings(doc);
            applyBodyParagraphJustification(doc);
            applyBaseStyleBasedOn(doc);
            applyTableCellSettings(doc);
            applyTocSubheadingBasedOn(doc);
            applyInlineGraphicSpacing(doc);
            applyCharacterStyleBasedOn(doc, "highlighter", "strong-bold");
            applyCharacterStyleBasedOn(doc, "code-strong", "code-normal");
            applyCharacterStyleBasedOn(doc, "li-label", "strong-bold");
            applySumaruNoBreak(doc);
            applyLangUSLanguageSetting(doc);
            applyCodeNormalLanguageSetting(doc);
            applyCodeParagraphSettings(doc);
        }

        // =========================================
        // 正規表現スタイル（ネスト GREP） / Nested GREP styles
        // =========================================

        /* 段落スタイルに正規表現スタイルを設定（同条件のものがあれば重複追加しない） /
           Add nested GREP styles to paragraph styles (skip duplicates).
           既存段落スタイルに対する追加は OVERWRITE_EXISTING_STYLES で制御 /
           Adding to pre-existing paragraph styles is gated by OVERWRITE_EXISTING_STYLES. */
        function applyNestedGrepStyleSettings(doc) {
            // group: 段落スタイルの所属グループ名（null はルート）/ owning group name (null = root)
            var nestedGrepRules = [
                { group: null, paragraph: "ul-li", character: "li-label", expression: "^.+?(?=：)" },
                { group: "basestyle", paragraph: "base-regex", character: "lang-US", expression: "[\\u\\l]" },
                { group: "basestyle", paragraph: "base-regex", character: "sumaru", expression: "..[。」』？！…]?$" },
                { group: "basestyle", paragraph: "base-regex", character: "inline-graphic", expression: "~a" }
            ];

            for (var grepRuleIndex = 0; grepRuleIndex < nestedGrepRules.length; grepRuleIndex++) {
                var grepRuleDefinition = nestedGrepRules[grepRuleIndex];
                // 段落スタイルのコンテナを解決（グループ指定があればそのグループ、無ければ doc）/
                // Resolve the paragraph style container (group if specified, otherwise doc)
                var paragraphContainer = grepRuleDefinition.group
                    ? doc.paragraphStyleGroups.itemByName(grepRuleDefinition.group)
                    : doc;
                if (!paragraphContainer.isValid) continue;
                if (!shouldApplyAttributesToParagraphStyle(paragraphContainer, grepRuleDefinition.paragraph)) continue;
                var targetParagraphStyle = paragraphContainer.paragraphStyles.itemByName(grepRuleDefinition.paragraph);
                var targetCharacterStyle = doc.characterStyles.itemByName(grepRuleDefinition.character);
                if (!targetParagraphStyle.isValid || !targetCharacterStyle.isValid) continue;
                var hasSameGrepStyle = false;
                for (var nestedGrepStyleIndex = 0; nestedGrepStyleIndex < targetParagraphStyle.nestedGrepStyles.length; nestedGrepStyleIndex++) {
                    var existingGrepStyle = targetParagraphStyle.nestedGrepStyles[nestedGrepStyleIndex];
                    // 文字スタイルは名前ではなく一意な id で比較（別グループの同名スタイルと誤判定しない）/
                    // Compare applied character style by unique id, not name (avoids same-name collisions across groups)
                    var existingCharacterStyle = existingGrepStyle.appliedCharacterStyle;
                    if (existingGrepStyle.grepExpression === grepRuleDefinition.expression &&
                        existingCharacterStyle.isValid &&
                        existingCharacterStyle.id === targetCharacterStyle.id) {
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
    }

    // 全処理を 1 つのアンドゥ単位にまとめて実行 /
    // Run everything as a single undo step
    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined,
        UndoModes.ENTIRE_SCRIPT, "スタイル一括登録");

})();
