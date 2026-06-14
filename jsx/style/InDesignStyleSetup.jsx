#target indesign

/*

    ### 概要

    アクティブな InDesign ドキュメントに、段落スタイル／文字スタイルとそれぞれのグループ（フォルダー）を一括登録し、見出し・本文・表・目次・コードなどの定番属性をまとめて設定する。

    - 同名のスタイル／グループが既にあれば作成はスキップ（重複登録しない）
    - 属性の適用は動作スイッチ `OVERWRITE_EXISTING_STYLES`（同名の既存スタイルを置き換えるか。既定 false）で制御。basedOn（基準スタイル）は常に設定する
    - 継承は basestyle グループに集約。共通の正規表現スタイル（GREP）は base-regex に持たせ、body-text / heading を base-regex 基準にすることで、本文系（p / ul-li / ol-li / p.caption・表セル）は body-text、見出し系（h1〜h6）は heading を経由して GREP を継承する（目次系 base-toc / toc-title / toc-h1〜h3 は base-regex を基準にしないため GREP は継承しない）
    - 行揃え・カーニング・段落分離禁止などの定番属性も body-text / heading / base-toc の基準スタイルに集約し、配下のスタイルへ basedOn で継承させる
    - パネル上の並び順は配列の宣言順に揃える（既存スタイルも move() で並び替え）
    - 全処理は `app.doScript`（UndoModes.ENTIRE_SCRIPT）で 1 つのアンドゥ単位にまとめる

    #### ユーザーが手動で設定する項目（本スクリプトは未設定）

    本スクリプトは行揃え・カーニング・分離禁止・言語・基準スタイル・GREP などの「構造」を整えるが、
    フォント・級数（サイズ）・行送り・文字色などの「見た目」は設定しない。以下は実行後に各自で設定する。

    - 全般: フォント／級数／行送り／文字色（特に p・h1〜h6）
    - strong-bold: 太字（フォントスタイルを Bold に）
    - em-italic: 斜体（フォントスタイルを Italic に）
    - code-normal / code-strong / p.code: 等幅フォント
    - highlighter: 強調表現（蛍光ペン／背景色など。基準は strong-bold）
    - link: リンク色（下線なしは設定済み）

    #### basestyle グループの各スタイルと設定

    継承の最上位を basestyle グループに集約し、本文系・見出し系・目次系の基準として使う。

    - base-regex: 共通 GREP（lang-US / no-break / inline-graphic）を保持する継承の最上位。直接の子は body-text / heading（base-toc は基準にしない）
    - body-text: 本文系の基準。カーニング「和文等幅」・行揃え「均等配置（最終行左揃え）」・すべての行を分離禁止・ハイフネーション OFF。base-regex を継承（GREP も継承）。p / ul-li / ol-li / p.caption・表セル（th / td）の基準（p は分離禁止のみ OFF で上書き）
    - heading: 見出し系の基準。カーニング「メトリクス」・行揃え「左揃え」・段落分離禁止（次の段落を保持2行＋すべての行を分離禁止）・ハイフネーション OFF。base-regex を継承（GREP も継承）。h1〜h6 の基準
    - base-toc: 目次系の基準。行揃え「左揃え」・段落分離禁止（次の段落を保持2行＋すべての行を分離禁止）・ハイフネーション OFF。basedOn なし（base-regex の GREP は継承しない）。toc-title / toc-h1 / toc-h2 / toc-h3 の基準

    ### 登録内容

    - 段落スタイル（ルート）: h1〜h6, ul-li, ol-li, p, p.caption, p.code, p.img
    - 段落スタイルグループ
      - basestyle: base-regex, body-text, heading, base-toc
      - table: th, th-left, th-center, th-right, td, td-left, td-center, td-right
      - toc: toc-title, toc-h1, toc-h2, toc-h3
      - book: page-number, running-head, thumb-index
    - 文字スタイル（ルート）: strong-bold, em-italic, link, code-normal, code-strong, highlighter
    - 文字スタイルグループ
      - table: td-bold
      - auto-apply: no-break, lang-US, inline-graphic, li-label, li-bullet, li-num

    ### 仕様

    同じ場所に同名のスタイル／グループが既にあれば追加せずスキップ（ルート直下と各グループ内を分けて判定）。
    以下の属性は `OVERWRITE_EXISTING_STYLES`（置き換えスイッチ）に従って適用。既定 false：同名の既存スタイルには
    一切触れず今回新規作成分のみ。true：同名の既存スタイルを置き換え（全属性を再適用、GREP は消してから付け直し）。

    **行揃え**

    - heading: 左揃え（h1〜h6 は basedOn で継承）
    - toc: 左揃え（toc-h1〜toc-h3 は basedOn で継承）
    - body-text: 均等配置（最終行左揃え）（p / ul-li / ol-li / p.caption・表セルは basedOn で継承）
    - th-left / th-center / th-right: 左 / 中央 / 右
    - td-left / td-center / td-right: 左 / 中央 / 右
    - p.code: 左
    - p.img: 中央

    **カーニング**

    - body-text: 和文等幅（本文系は basedOn で継承）
    - heading: メトリクス（h1〜h6 は basedOn で継承）

    **段落分離禁止（Keep Options）**

    - heading（→ h1〜h6）/ base-toc（→ toc-title / toc-h1〜toc-h3）: 次の段落を保持「2行」＋ すべての行を分離禁止（basedOn で継承）
    - body-text: すべての行を分離禁止（p / ul-li / ol-li / p.caption・th / td は basedOn で継承）
    - p: 分離禁止オプションをすべて OFF（すべての行を分離禁止／次の段落を保持／前の段落から分離しない。body-text の継承も打ち消す）
    - ul-li / p.caption: 前の段落から分離しない
    - p.code: すべての行を分離禁止（body-text を継承しないため単独設定）

    **箇条書き（リスト）**

    - ul-li: 箇条書き「記号」（bulletsAndNumberingListType = BULLET_LIST）。行頭記号に文字スタイル li-bullet を適用（bulletsCharacterStyle）
    - ol-li: 箇条書き「自動番号」（bulletsAndNumberingListType = NUMBERED_LIST）。番号に文字スタイル li-num を適用（numberingCharacterStyle）

    **次のスタイル**

    - h1〜h6 / p.caption → p（toc-h1〜toc-h3 は対象外）

    **言語・欧文合字・ハイフネーション**

    - body-text / heading / base-toc: ハイフネーション OFF（配下は basedOn で継承）
    - lang-US: 言語「英語：米国」
    - code-normal: 言語「なし」・欧文合字オフ
    - p.code: 言語「なし」・欧文合字オフ・ハイフネーションオフ

    **その他**

    - inline-graphic: 前後四分のアキ（leadingAki / trailingAki = 0.25）
    - link: 下線なし（underline = false）
    - no-break: noBreak（分割禁止）

    ### 基準スタイル（basedOn）

    - body-text → base-regex（basestyle グループ内。base-regex の GREP を継承）
    - heading → base-regex（basestyle グループ内。base-regex の GREP を継承）
    - p / ul-li / ol-li / p.caption → body-text（ルート → basestyle グループ）
    - h1〜h6 → heading（ルート → basestyle グループ）
    - th / td → body-text（table → basestyle グループ）
    - th-left / th-center / th-right → th（table グループ内）
    - td-left / td-center / td-right → td（table グループ内）
    - toc-title / toc-h1 / toc-h2 / toc-h3 → base-toc（toc グループ → basestyle グループ）
    - code-strong → code-normal
    - highlighter → strong-bold
    - li-label → strong-bold

    ### 正規表現スタイル（追加のみ・旧ルールは削除しない）

    同条件（GREP 式＋適用文字スタイル id が一致）のルールがあれば重複追加しない。
    既存の GREP ルールは削除しない仕様（追加のみ）。本スクリプトの定義に無い旧ルールもそのまま残す。

    共通3つ（lang-US / no-break / inline-graphic）は base-regex に持たせ、own GREP を持たない子
    （p / ol-li / p.caption / h1〜h6 / 表セルなど）へ basedOn で継承させる。一方 ul-li は li-label を
    自身に持つため、InDesign の仕様で GREP の継承が切れる（own GREP を1つでも持つと親の GREP を継承しない）。
    そのため ul-li には共通3つ（lang-US / no-break / inline-graphic）も直接設定する。

    - base-regex（basestyle）: `[\u\l]` → lang-US
    - base-regex（basestyle）: `..[。」』？！…]?$` → no-break
    - base-regex（basestyle）: `~a` → inline-graphic
    - ul-li（共通3つ＋固有1つを直接設定）: `[\u\l]` → lang-US / `..[。」』？！…]?$` → no-break / `~a` → inline-graphic / `^.+?(?=：)` → li-label

    ### 変更履歴

    v1.3.0
    - ul-li の GREP 継承切れを修正（ul-li に共通3つ＋li-label を直接設定）
    - OVERWRITE_EXISTING_STYLES を「置き換え」化（既定 false。ON で全属性再適用＋GREP は消して付け直し）
    - APPLY_BASED_ON_RELATIONSHIPS スイッチを廃止（basedOn は常に設定）
    - 概要に「ユーザーが手動で設定する項目（フォント・太字・色など）」を追記

    v1.2.5
    - auto-apply グループを追加し no-break / lang-US / inline-graphic / li-label / li-bullet / li-num を集約
    - スタイル名を整理（body-text / base-toc / auto-apply / no-break）、文字スタイルの toc / book グループを廃止
    - ハイフネーション OFF・分離禁止を基準スタイルへ集約、ul-li=箇条書き記号 / ol-li=自動番号
    - p.caption の次スタイル=p、p.img 中央、code-normal 欧文合字オフ、link 下線なし

    v1.2.0
    - basestyle に body-text / heading / base-toc を追加し継承を再編、共通 GREP を base-regex に集約
    - カーニング・行揃え・段落分離禁止を基準スタイルへ集約（配下は basedOn で継承）
    - p.code に言語なし・欧文合字オフ等、OVERWRITE_EXISTING_STYLES の既定を true に変更
    - 全処理を 1 アンドゥ（ENTIRE_SCRIPT）化＆ IIFE 化、各種バグ修正・堅牢化

    */

(function () {

    // =========================================
    // バージョン / Version
    // =========================================

    var SCRIPT_VERSION = "v1.3.0";

    // =========================================
    // 動作スイッチ / Behavior switches
    // =========================================

    /* 対象ドキュメントに同名の段落スタイル／文字スタイルが既にある場合の挙動（全スタイル・全属性が対象）/
       Behavior when a same-named paragraph/character style already exists (all styles, all attributes).
       true:  既存スタイルを「置き換え」る。本スクリプトが設定する全属性（行揃え／カーニング／言語／basedOn など）を
              再適用し、GREP は既存ルールを消してから付け直す（add-only ではなく完全置換）/
              replace the existing same-named style: re-apply every attribute this script sets, and
              fully replace its GREP rules (clear existing, then add).
       false: 既存スタイルには一切触れず、今回新規作成したスタイルにのみ属性を適用する（現在の設定）/
              leave existing same-named styles untouched; apply only to newly created ones.
       ※ 本スクリプトが扱わない属性（任意のフォント・サイズ等）はリセットしない。スタイル実体は削除しないため、
          適用済みテキストとの関連は保たれる / Attributes this script doesn't set are not reset; the style
          object itself is not deleted, so text keeps its style association. */
    var OVERWRITE_EXISTING_STYLES = false;

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
            "link", "code-normal", "code-strong", "highlighter"
        ];

        var paragraphStyleGroupNames = [
            "basestyle", "table", "toc", "book"
        ];

        var paragraphStylesInGroups = [
            { group: "basestyle", styles: ["base-regex", "body-text", "heading", "base-toc"] },
            { group: "table", styles: ["th", "th-left", "th-center", "th-right", "td", "td-left", "td-center", "td-right"] },
            { group: "toc", styles: ["toc-title", "toc-h1", "toc-h2", "toc-h3"] },
            { group: "book", styles: ["page-number", "running-head", "thumb-index"] }
        ];

        var characterStyleGroupNames = [
            "table", "auto-apply"
        ];

        var characterStylesInGroups = [
            { group: "table", styles: ["td-bold"] },
            { group: "auto-apply", styles: ["no-break", "lang-US", "inline-graphic", "li-label", "li-bullet", "li-num"] }
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
        // ※ 既存スタイルへの上書きは OVERWRITE_EXISTING_STYLES（shouldApplyAttributesTo* がガード）。basedOn は常に設定。
        //   Guarded by OVERWRITE_EXISTING_STYLES (overwrite). basedOn is always set.

        /* basestyle グループの body-text / heading / base-toc に基本属性を設定 /
           - body-text: カーニング「和文等幅」・行揃え「均等配置（最終行左揃え）」・すべての行を分離禁止（p 等は basedOn で継承）
           - heading: カーニング「メトリクス」・行揃え「左揃え」・段落分離禁止（次の段落を保持2行＋すべての行を分離禁止）（h1〜h6 は basedOn で継承）
           - toc: 行揃え「左揃え」・段落分離禁止（次の段落を保持2行＋すべての行を分離禁止）（toc-h1〜toc-h3 は basedOn で継承）/
           Apply base attributes to body-text, heading and toc; descendants inherit via basedOn */
        function applyBaseGroupStyleSettings(doc) {
            var baseGroup = doc.paragraphStyleGroups.itemByName("basestyle");
            if (!baseGroup.isValid) return;

            if (shouldApplyAttributesToParagraphStyle(baseGroup, "body-text")) {
                var bodyTextStyle = baseGroup.paragraphStyles.itemByName("body-text");
                if (bodyTextStyle.isValid) {
                    bodyTextStyle.kerningMethod = "和文等幅";
                    bodyTextStyle.justification = Justification.LEFT_JUSTIFIED;
                    bodyTextStyle.keepLinesTogether = true;
                    bodyTextStyle.keepAllLinesTogether = true;
                    bodyTextStyle.hyphenation = false;
                }
            }

            if (shouldApplyAttributesToParagraphStyle(baseGroup, "heading")) {
                var headingStyle = baseGroup.paragraphStyles.itemByName("heading");
                if (headingStyle.isValid) {
                    headingStyle.kerningMethod = "メトリクス";
                    headingStyle.justification = Justification.LEFT_ALIGN;
                    headingStyle.keepWithNext = 2;
                    headingStyle.keepLinesTogether = true;
                    headingStyle.keepAllLinesTogether = true;
                    headingStyle.hyphenation = false;
                }
            }

            if (shouldApplyAttributesToParagraphStyle(baseGroup, "base-toc")) {
                var tocStyle = baseGroup.paragraphStyles.itemByName("base-toc");
                if (tocStyle.isValid) {
                    tocStyle.justification = Justification.LEFT_ALIGN;
                    tocStyle.keepWithNext = 2;
                    tocStyle.keepLinesTogether = true;
                    tocStyle.keepAllLinesTogether = true;
                    tocStyle.hyphenation = false;
                }
            }
        }

        /* basestyle グループ内の基準連鎖を設定（base-regex の GREP などを継承）/
           - body-text → base-regex
           - heading → base-regex（basestyle グループ内）
           - p / ul-li / ol-li / p.caption → body-text（ルート → basestyle グループ）
           - h1〜h6 → heading（ルート → basestyle グループ。heading 経由で base-regex の GREP を継承）/
           Set basedOn chain: base-regex → body-text → (p/ul-li/ol-li/p.caption); base-regex → heading → h1–h6 */
        function applyBaseStyleBasedOn(doc) {
            var baseGroup = doc.paragraphStyleGroups.itemByName("basestyle");
            if (!baseGroup.isValid) return;
            var baseStyle = baseGroup.paragraphStyles.itemByName("base-regex");
            if (!baseStyle.isValid) return;

            // body-text を base-regex 基準に（heading も base-regex を基準にする。base-toc は基準にしない）/
            // body-text → base-regex (heading is also based on base-regex; base-toc is not)
            var bodyTextStyle = baseGroup.paragraphStyles.itemByName("body-text");
            if (bodyTextStyle.isValid &&
                shouldApplyAttributesToParagraphStyle(baseGroup, "body-text")) {
                bodyTextStyle.basedOn = baseStyle;
            }
            if (!bodyTextStyle.isValid) return;

            // heading を base-regex 基準に（h1〜h6 から継承させるための中間スタイル）/
            // heading → base-regex (intermediate style inherited by h1–h6)
            var headingStyle = baseGroup.paragraphStyles.itemByName("heading");
            if (headingStyle.isValid &&
                shouldApplyAttributesToParagraphStyle(baseGroup, "heading")) {
                headingStyle.basedOn = baseStyle;
            }

            // p / ul-li / ol-li / p.caption → body-text
            var basedOnBaseStyleNames = ["p", "ul-li", "ol-li", "p.caption"];
            for (var basedOnIndex = 0; basedOnIndex < basedOnBaseStyleNames.length; basedOnIndex++) {
                var basedOnStyleName = basedOnBaseStyleNames[basedOnIndex];
                if (!shouldApplyAttributesToParagraphStyle(doc, basedOnStyleName)) continue;
                var basedOnTargetStyle = doc.paragraphStyles.itemByName(basedOnStyleName);
                if (basedOnTargetStyle.isValid) basedOnTargetStyle.basedOn = bodyTextStyle;
            }

            // h1〜h6 → heading
            if (headingStyle.isValid) {
                var headingBasedOnNames = ["h1", "h2", "h3", "h4", "h5", "h6"];
                for (var headingBasedOnIndex = 0; headingBasedOnIndex < headingBasedOnNames.length; headingBasedOnIndex++) {
                    var headingBasedOnName = headingBasedOnNames[headingBasedOnIndex];
                    if (!shouldApplyAttributesToParagraphStyle(doc, headingBasedOnName)) continue;
                    var headingTargetStyle = doc.paragraphStyles.itemByName(headingBasedOnName);
                    if (headingTargetStyle.isValid) headingTargetStyle.basedOn = headingStyle;
                }
            }
        }

        /* h1〜h6 / p.caption の「次のスタイル」を p に設定 /
           （行揃え・カーニング・段落分離禁止は heading（basestyle）に設定済みで basedOn 継承。
             toc-h1〜toc-h3 も同様に toc から継承するため、ここでは toc を扱わない）/
           Set next style = p for h1–h6 and p.caption */
        function applyNextStyleSettings(doc) {
            var bodyParagraphStyle = doc.paragraphStyles.itemByName("p");
            if (!bodyParagraphStyle.isValid) return;
            var nextStyleTargetNames = ["h1", "h2", "h3", "h4", "h5", "h6", "p.caption"];
            for (var nextStyleIndex = 0; nextStyleIndex < nextStyleTargetNames.length; nextStyleIndex++) {
                var nextStyleTargetName = nextStyleTargetNames[nextStyleIndex];
                if (!shouldApplyAttributesToParagraphStyle(doc, nextStyleTargetName)) continue;
                var nextStyleTargetStyle = doc.paragraphStyles.itemByName(nextStyleTargetName);
                if (nextStyleTargetStyle.isValid) nextStyleTargetStyle.nextStyle = bodyParagraphStyle;
            }
        }

        /* 段落分離禁止オプションを設定 /
           Apply keep options:
           - 前の段落から分離しない（泣き別れ禁止）: ul-li / p.caption /
             keep-with-previous: ul-li, p.caption
           - すべての行を分離禁止: body-text に集約（ol-li / ul-li / p.caption は basedOn で継承）。
             p.code は body-text を継承しないため単独で設定。p は body-text を継承するが OFF で上書き /
             keep-all-lines-together: lives on body-text (ol-li/ul-li/p.caption inherit);
             p.code set directly; p overrides it OFF */
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

            // body-text を継承しない p.code は単独で「すべての行を分離禁止」を設定 /
            // p.code does not inherit from body-text, so set keep-all-lines-together directly
            if (shouldApplyAttributesToParagraphStyle(doc, "p.code")) {
                var codeKeepStyle = doc.paragraphStyles.itemByName("p.code");
                if (codeKeepStyle.isValid) {
                    codeKeepStyle.keepLinesTogether = true;
                    codeKeepStyle.keepAllLinesTogether = true;
                }
            }

            // p は分離禁止オプションをすべて OFF（body-text からの継承も含めて打ち消す）/
            // p turns off all keep options (also overriding what is inherited from body-text)
            if (shouldApplyAttributesToParagraphStyle(doc, "p")) {
                var bodyKeepStyle = doc.paragraphStyles.itemByName("p");
                if (bodyKeepStyle.isValid) {
                    bodyKeepStyle.keepLinesTogether = false;
                    bodyKeepStyle.keepAllLinesTogether = false;
                    bodyKeepStyle.keepWithNext = 0;
                    bodyKeepStyle.keepWithPrevious = false;
                }
            }
        }

        /* p.img（画像用段落）を中央寄せに / Center-align p.img (image paragraph) */
        function applyImageParagraphSettings(doc) {
            if (!shouldApplyAttributesToParagraphStyle(doc, "p.img")) return;
            var imageParagraphStyle = doc.paragraphStyles.itemByName("p.img");
            if (imageParagraphStyle.isValid) {
                imageParagraphStyle.justification = Justification.CENTER_ALIGN;
            }
        }

        /* ul-li に箇条書き「記号」（行頭記号に文字スタイル li-bullet を適用）、
           ol-li に箇条書き「自動番号」（番号に文字スタイル li-num を適用）を設定 /
           Set bullet list on ul-li (bullet uses li-bullet) and numbered list on ol-li (number uses li-num) */
        function applyListSettings(doc) {
            if (shouldApplyAttributesToParagraphStyle(doc, "ul-li")) {
                var bulletListStyle = doc.paragraphStyles.itemByName("ul-li");
                if (bulletListStyle.isValid) {
                    bulletListStyle.bulletsAndNumberingListType = ListType.BULLET_LIST;
                    var bulletCharacterStyle = resolveCharacterStyle(doc, "li-bullet");
                    if (bulletCharacterStyle) {
                        bulletListStyle.bulletsCharacterStyle = bulletCharacterStyle.style;
                    }
                    // 同じスタイルが連続する段落間のスペースを 0 に（対応バージョンのみ。
                    //   プロパティ名はバージョン差があるため候補から存在するものを設定）/
                    // Space between paragraphs using the same style = 0 (only on supporting versions)
                    setOptionalProperty(bulletListStyle,
                        ["spaceBetweenParagraphs", "spaceBetweenSameParagraphStyles", "spaceBetweenSameStyleParagraphs"], 0);
                }
            }
            if (shouldApplyAttributesToParagraphStyle(doc, "ol-li")) {
                var numberedListStyle = doc.paragraphStyles.itemByName("ol-li");
                if (numberedListStyle.isValid) {
                    numberedListStyle.bulletsAndNumberingListType = ListType.NUMBERED_LIST;
                    var numberingCharacterStyle = resolveCharacterStyle(doc, "li-num");
                    if (numberingCharacterStyle) {
                        numberedListStyle.numberingCharacterStyle = numberingCharacterStyle.style;
                    }
                }
            }
        }

        /* table グループ内 th / td を body-text ベースに（「すべての行を分離禁止」は body-text から継承）、
           th-* / td-* に行揃えと th / td ベースを設定 /
           Set basedOn=body-text on th/td (keep-all-lines-together inherited), plus alignment and basedOn=th/td for th-* and td-* */
        function applyTableCellSettings(doc) {
            var tableGroup = doc.paragraphStyleGroups.itemByName("table");
            if (!tableGroup.isValid) return;

            // body-text（basestyle グループ）を th / td の基準にし、th-* / td-* へ basedOn 経由で継承させる /
            // body-text (basestyle group) is the base of th/td; th-*/td-* inherit via basedOn
            var baseGroup = doc.paragraphStyleGroups.itemByName("basestyle");
            var bodyTextStyle = baseGroup.isValid ? baseGroup.paragraphStyles.itemByName("body-text") : null;

            // th / td を body-text ベースに（「すべての行を分離禁止」は body-text から継承するため個別設定しない）/
            // basedOn=body-text on th/td (keep-all-lines-together inherited from body-text)
            var tableBaseStyleNames = ["th", "td"];
            for (var tableBaseIndex = 0; tableBaseIndex < tableBaseStyleNames.length; tableBaseIndex++) {
                var tableBaseName = tableBaseStyleNames[tableBaseIndex];
                if (!shouldApplyAttributesToParagraphStyle(tableGroup, tableBaseName)) continue;
                var tableBaseStyle = tableGroup.paragraphStyles.itemByName(tableBaseName);
                if (tableBaseStyle.isValid) {
                    if (bodyTextStyle && bodyTextStyle.isValid) {
                        tableBaseStyle.basedOn = bodyTextStyle;
                    }
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
                if (thBaseStyle.isValid) {
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
                if (tdBaseStyle.isValid) {
                    tdTargetStyle.basedOn = tdBaseStyle;
                }
                tdTargetStyle.justification = tdTarget.justification;
            }
        }

        /* toc グループ内 toc-title / toc-h1 / toc-h2 / toc-h3 を basestyle グループの base-toc ベースに /
           Set basedOn=base-toc (basestyle group) for toc-title, toc-h1, toc-h2 and toc-h3 */
        function applyTocSubheadingBasedOn(doc) {
            var tocGroup = doc.paragraphStyleGroups.itemByName("toc");
            if (!tocGroup.isValid) return;
            var baseGroup = doc.paragraphStyleGroups.itemByName("basestyle");
            if (!baseGroup.isValid) return;
            var tocBaseStyle = baseGroup.paragraphStyles.itemByName("base-toc");
            if (!tocBaseStyle.isValid) return;

            var tocSubheadingNames = ["toc-title", "toc-h1", "toc-h2", "toc-h3"];
            for (var tocIndex = 0; tocIndex < tocSubheadingNames.length; tocIndex++) {
                var tocSubheadingName = tocSubheadingNames[tocIndex];
                if (!shouldApplyAttributesToParagraphStyle(tocGroup, tocSubheadingName)) continue;
                var tocSubheadingStyle = tocGroup.paragraphStyles.itemByName(tocSubheadingName);
                if (tocSubheadingStyle.isValid) tocSubheadingStyle.basedOn = tocBaseStyle;
            }
        }

        /* toc-h3 の「次の段落を保持」を 0 に設定（base-toc から継承した 2 を打ち消す）/
           Set keep-with-next = 0 on toc-h3 (cancels the 2 inherited from base-toc) */
        function applyTocLeafOverrides(doc) {
            var tocGroup = doc.paragraphStyleGroups.itemByName("toc");
            if (!tocGroup.isValid) return;
            if (!shouldApplyAttributesToParagraphStyle(tocGroup, "toc-h3")) return;
            var tocH3Style = tocGroup.paragraphStyles.itemByName("toc-h3");
            if (tocH3Style.isValid) {
                tocH3Style.keepWithNext = 0;
            }
        }

        /* バージョンによって存在しないプロパティを安全に設定する。
           候補名を順に reflect で存在確認し、最初に見つかったものへ value を設定。
           設定できた名前を返す（どれも未対応なら null）/
           Safely set a possibly-unsupported property: try candidate names, set the first one
           that exists (checked via reflect). Returns the name used, or null if none supported. */
        function setOptionalProperty(targetObject, candidateNames, value) {
            var availableProperties = targetObject.reflect.properties;
            for (var candidateIndex = 0; candidateIndex < candidateNames.length; candidateIndex++) {
                var candidateName = candidateNames[candidateIndex];
                for (var propertyIndex = 0; propertyIndex < availableProperties.length; propertyIndex++) {
                    if (String(availableProperties[propertyIndex].name) === candidateName) {
                        targetObject[candidateName] = value;
                        return candidateName;
                    }
                }
            }
            return null;
        }

        /* 文字スタイルを名前で解決（ルート → auto-apply グループの順で探す）。
           見つかれば { style: 文字スタイル, container: doc または auto-apply グループ } を返す。無ければ null /
           Resolve a character style by name (root first, then the "auto-apply" group).
           Returns { style, container } or null. container is used for the create/overwrite guard. */
        function resolveCharacterStyle(doc, styleName) {
            var rootStyle = doc.characterStyles.itemByName(styleName);
            if (rootStyle.isValid) return { style: rootStyle, container: doc };
            var autoApplyGroup = doc.characterStyleGroups.itemByName("auto-apply");
            if (autoApplyGroup.isValid) {
                var groupedStyle = autoApplyGroup.characterStyles.itemByName(styleName);
                if (groupedStyle.isValid) return { style: groupedStyle, container: autoApplyGroup };
            }
            return null;
        }

        /* inline-graphic に前後四分のアキを設定 /
           Apply quarter-em leading/trailing spacing to inline-graphic */
        function applyInlineGraphicSpacing(doc) {
            var resolved = resolveCharacterStyle(doc, "inline-graphic");
            if (!resolved) return;
            if (!shouldApplyAttributesToCharacterStyle(resolved.container, "inline-graphic")) return;
            resolved.style.leadingAki = 0.25;
            resolved.style.trailingAki = 0.25;
        }

        /* link の下線をなしに / Turn off underline on link */
        function applyLinkSettings(doc) {
            var resolved = resolveCharacterStyle(doc, "link");
            if (!resolved) return;
            if (!shouldApplyAttributesToCharacterStyle(resolved.container, "link")) return;
            resolved.style.underline = false;
        }

        /* no-break の分割禁止を有効化 / Enable noBreak on no-break */
        function applyNoBreakSettings(doc) {
            var resolved = resolveCharacterStyle(doc, "no-break");
            if (!resolved) return;
            if (!shouldApplyAttributesToCharacterStyle(resolved.container, "no-break")) return;
            resolved.style.noBreak = true;
        }

        /* 1 つの文字スタイルの basedOn を別の文字スタイルに設定する共通処理（グループ越え可）/
           Common helper: set basedOn of one character style to another (across groups) */
        function applyCharacterStyleBasedOn(doc, targetStyleName, parentStyleName) {
            var target = resolveCharacterStyle(doc, targetStyleName);
            var parent = resolveCharacterStyle(doc, parentStyleName);
            if (!target || !parent) return;
            if (!shouldApplyAttributesToCharacterStyle(target.container, targetStyleName)) return;
            target.style.basedOn = parent.style;
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
            var resolved = resolveCharacterStyle(doc, "lang-US");
            if (!resolved) return;
            if (!shouldApplyAttributesToCharacterStyle(resolved.container, "lang-US")) return;
            var englishLanguage = resolveLanguageByNames(ENGLISH_USA_LANGUAGE_NAMES);
            if (englishLanguage) resolved.style.appliedLanguage = englishLanguage;
        }

        /* code-normal の言語を「なし」に・欧文合字オフ / Set applied language to [No Language] and ligatures off */
        function applyCodeNormalLanguageSetting(doc) {
            var resolved = resolveCharacterStyle(doc, "code-normal");
            if (!resolved) return;
            if (!shouldApplyAttributesToCharacterStyle(resolved.container, "code-normal")) return;
            var noLanguage = resolveLanguageByNames(NO_LANGUAGE_NAMES);
            if (noLanguage) resolved.style.appliedLanguage = noLanguage;
            resolved.style.ligatures = false;
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
            applyBaseGroupStyleSettings(doc);
            applyNextStyleSettings(doc);
            applyKeepTogetherSettings(doc);
            applyListSettings(doc);
            applyImageParagraphSettings(doc);
            applyBaseStyleBasedOn(doc);
            applyTableCellSettings(doc);
            applyTocSubheadingBasedOn(doc);
            applyTocLeafOverrides(doc);
            applyInlineGraphicSpacing(doc);
            applyLinkSettings(doc);
            applyCharacterStyleBasedOn(doc, "highlighter", "strong-bold");
            applyCharacterStyleBasedOn(doc, "code-strong", "code-normal");
            applyCharacterStyleBasedOn(doc, "li-label", "strong-bold");
            applyNoBreakSettings(doc);
            applyLangUSLanguageSetting(doc);
            applyCodeNormalLanguageSetting(doc);
            applyCodeParagraphSettings(doc);
        }

        // =========================================
        // 正規表現スタイル（ネスト GREP） / Nested GREP styles
        // =========================================

        /* 段落スタイルに正規表現スタイルを設定 / Set nested GREP styles on paragraph styles.
           OVERWRITE_EXISTING_STYLES（置き換え）= ON: 既存スタイルの GREP を全削除してから本定義を付け直す
             （add-only ではなく完全置換。旧ルール・残存ルールも消える）/
             replace mode on: clear the style's existing GREP entirely, then add this definition.
           OFF: 同名の既存スタイルには触れない（guard でスキップ）。今回新規作成したスタイルにのみ付ける /
             off: leave existing same-named styles untouched; only newly created styles get GREP. */
        function applyNestedGrepStyleSettings(doc) {
            // group: 段落スタイルの所属グループ名（null はルート）/ owning group name (null = root)
            // base-regex: 共通3つ。own GREP を持たない子（p / ol-li / p.caption / h1〜h6 / 表セル）へ basedOn で継承 /
            //   base-regex: 3 shared rules, inherited via basedOn by children that have no own GREP.
            // ul-li: own GREP（li-label）を1つでも持つと InDesign は GREP の継承を切る（own リストが継承分を置き換える）ため、
            //   共通3つも ul-li に直接設定する。継承は切れているので二重にはならない /
            //   ul-li: once a style has any own GREP, InDesign stops inheriting (the own list replaces the
            //   inherited one), so set the 3 shared rules directly on ul-li too. No duplication, since inheritance is off.
            //   ※ 手動で GREP を足すと UI が継承分を own へコピーしてから追加するので継承が残って見えるが、
            //     スクリプトの nestedGrepStyles.add() はコピーしないため継承が切れる（h1 等は own GREP が無いので継承表示される）/
            //   NOTE: manual add copies inherited rules into the own list first (so they appear to persist), but
            //     scripted add() does not copy them, so inheritance is severed (h1 etc. have no own GREP, so they still inherit).
            //   li-label は最後に置き、重なる範囲で優先させる / li-label is last so it wins on overlapping ranges.
            var nestedGrepRules = [
                { group: "basestyle", paragraph: "base-regex", character: "lang-US", expression: "[\\u\\l]" },
                { group: "basestyle", paragraph: "base-regex", character: "no-break", expression: "..[。」』？！…]?$" },
                { group: "basestyle", paragraph: "base-regex", character: "inline-graphic", expression: "~a" },
                { group: null, paragraph: "ul-li", character: "lang-US", expression: "[\\u\\l]" },
                { group: null, paragraph: "ul-li", character: "no-break", expression: "..[。」』？！…]?$" },
                { group: null, paragraph: "ul-li", character: "inline-graphic", expression: "~a" },
                { group: null, paragraph: "ul-li", character: "li-label", expression: "^.+?(?=：)" }
            ];

            // 置き換えモード（OVERWRITE_EXISTING_STYLES）では、各対象スタイルの既存 GREP を
            //   一度だけ全削除してから付け直す（同じ実行内で重複削除しないよう id で記録）/
            //   In replace mode, clear each target style's existing GREP once before re-adding.
            var grepClearedStyleIds = {};

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
                // 文字スタイルはルート → auto-apply グループの順で解決 / Resolve character style across root and "auto-apply" group
                var resolvedCharacter = resolveCharacterStyle(doc, grepRuleDefinition.character);
                if (!targetParagraphStyle.isValid || !resolvedCharacter) continue;

                // 置き換えモードでは、このスタイルの既存 GREP を一度だけ全削除（末尾から削除して添字ずれ回避）/
                // In replace mode, clear this style's existing GREP once (remove from the end to avoid index shift)
                if (OVERWRITE_EXISTING_STYLES && !grepClearedStyleIds[targetParagraphStyle.id]) {
                    for (var grepClearIndex = targetParagraphStyle.nestedGrepStyles.length - 1; grepClearIndex >= 0; grepClearIndex--) {
                        targetParagraphStyle.nestedGrepStyles[grepClearIndex].remove();
                    }
                    grepClearedStyleIds[targetParagraphStyle.id] = true;
                }

                var targetCharacterStyle = resolvedCharacter.style;
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
