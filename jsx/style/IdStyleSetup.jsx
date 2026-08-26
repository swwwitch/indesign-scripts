#target indesign

/*

### 概要

段落スタイル・文字スタイルとそのグループ、継承関係、正規表現スタイルまでを一括で登録します。

詳細は README を参照してください。

### Overview

Registers paragraph and character styles together with their groups, inheritance and GREP styles in one pass.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdStyleSetup";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-06-30";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdStyleSetup.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdStyleSetup.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nfe87ec253780"; /* 紹介記事 / article URL */

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

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 同名スタイルが既にある場合の挙動 / Behavior when a same-named style already exists.
       true:  既存スタイルを置き換える。本スクリプトが設定する全属性を再適用し、GREP は消してから付け直す
              / Replace the existing style: re-apply every attribute this script sets and rebuild its GREP rules
       false: 既存スタイルには触れず、新規作成したスタイルにだけ属性を適用する
              / Leave existing styles untouched and apply attributes only to newly created ones
       本スクリプトが扱わない属性（フォント・サイズなど）はリセットしません。スタイル実体も削除しないため、
       適用済みテキストとの関連は保たれます。
       / Attributes this script does not set are left alone, and no style object is deleted,
         so text keeps its style association. */
    var OVERWRITE_EXISTING_STYLES = false;

    // =========================================
    // レイアウト設定 / Layout settings
    // =========================================

    /* 進捗バーの幅と高さ（px）/ Width and height of the progress bar (px) */
    var PROGRESS_BAR_WIDTH  = 320;
    var PROGRESS_BAR_HEIGHT = 12;

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
        progress: {
            title:   { ja: "スタイル一括登録", en: "Register Styles" },
            styles:  { ja: "スタイルとグループを作成中…", en: "Creating styles and groups…" },
            attrs:   { ja: "属性を適用中…", en: "Applying attributes…" },
            grep:    { ja: "正規表現スタイルを設定中…", en: "Setting GREP styles…" },
            reorder: { ja: "並び替え中…", en: "Reordering…" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document before running." }
        },
        undo: {
            registerStyles: { ja: "スタイル一括登録", en: "Register Styles" }
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
    // プログレスバー / Progress bar
    // =========================================

    /**
     * 進捗表示用のパレットを作る
     * @param {number} totalSteps 全体のステップ数
     * @returns {object} 更新と終了を行うオブジェクト
     */
    function createProgressWindow(totalSteps) {
        var progressWindow = null;
        try {
            progressWindow = new Window("palette", getLabel("progress.title") + "  " + SCRIPT_VERSION, undefined, { closeButton: false });
        } catch (e) {
            progressWindow = null;
        }
        if (!progressWindow) {
            return { step: function () {}, close: function () {} };
        }
        setupWindow(progressWindow, 10);

        var progressMessage = progressWindow.add("statictext", undefined, "");
        progressMessage.preferredSize.width = PROGRESS_BAR_WIDTH;

        var progressBar = progressWindow.add("progressbar", undefined, 0, totalSteps);
        progressBar.preferredSize.width = PROGRESS_BAR_WIDTH;
        progressBar.preferredSize.height = PROGRESS_BAR_HEIGHT;

        progressWindow.show();

        return {
            step: function (message) {
                progressBar.value += 1;
                progressMessage.text = message;
                progressWindow.update();
            },
            close: function () {
                progressWindow.close();
            }
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * スタイルとグループを作成し、属性・GREP・並び順をまとめて適用する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
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

        /**
         * スタイルの所在を表す一意なキーを作る
         * @param {object} styleContainer スタイルのコンテナ
         * @param {string} styleName スタイル名
         * @returns {string} 識別キー
         */
        function styleContainerKey(styleContainer, styleName) {
            return styleContainer.id + "\t" + styleName;
        }

        /**
         * 段落スタイルグループを取得する（なければ作成）
         * @param {Document} doc 対象ドキュメント
         * @param {string} groupName グループ名
         * @returns {ParagraphStyleGroup} スタイルグループ
         */
        function ensureParagraphStyleGroup(doc, groupName) {
            var styleGroup = doc.paragraphStyleGroups.itemByName(groupName);
            if (!styleGroup.isValid) {
                styleGroup = doc.paragraphStyleGroups.add({ name: groupName });
            }
            return styleGroup;
        }

        /**
         * 文字スタイルグループを取得する（なければ作成）
         * @param {Document} doc 対象ドキュメント
         * @param {string} groupName グループ名
         * @returns {CharacterStyleGroup} スタイルグループ
         */
        function ensureCharacterStyleGroup(doc, groupName) {
            var styleGroup = doc.characterStyleGroups.itemByName(groupName);
            if (!styleGroup.isValid) {
                styleGroup = doc.characterStyleGroups.add({ name: groupName });
            }
            return styleGroup;
        }

        /**
         * 段落スタイルを取得する（なければ作成）
         * @param {object} styleContainer スタイルのコンテナ
         * @param {string} styleName スタイル名
         * @returns {ParagraphStyle} 段落スタイル
         */
        function ensureParagraphStyle(styleContainer, styleName) {
            var paragraphStyle = styleContainer.paragraphStyles.itemByName(styleName);
            if (!paragraphStyle.isValid) {
                paragraphStyle = styleContainer.paragraphStyles.add({ name: styleName });
                paragraphStyleKeysCreatedThisRun[styleContainerKey(styleContainer, styleName)] = true;
            }
            return paragraphStyle;
        }

        /**
         * 文字スタイルを取得する（なければ作成）
         * @param {object} styleContainer スタイルのコンテナ
         * @param {string} styleName スタイル名
         * @returns {CharacterStyle} 文字スタイル
         */
        function ensureCharacterStyle(styleContainer, styleName) {
            var characterStyle = styleContainer.characterStyles.itemByName(styleName);
            if (!characterStyle.isValid) {
                characterStyle = styleContainer.characterStyles.add({ name: styleName });
                characterStyleKeysCreatedThisRun[styleContainerKey(styleContainer, styleName)] = true;
            }
            return characterStyle;
        }

        /**
         * その段落スタイルへ属性を適用してよいかを判定する
         * @param {object} styleContainer スタイルのコンテナ
         * @param {string} styleName スタイル名
         * @returns {boolean} 適用してよければ true
         */
        function shouldApplyAttributesToParagraphStyle(styleContainer, styleName) {
            return OVERWRITE_EXISTING_STYLES ||
                paragraphStyleKeysCreatedThisRun[styleContainerKey(styleContainer, styleName)] === true;
        }

        /**
         * その文字スタイルへ属性を適用してよいかを判定する
         * @param {object} styleContainer スタイルのコンテナ
         * @param {string} styleName スタイル名
         * @returns {boolean} 適用してよければ true
         */
        function shouldApplyAttributesToCharacterStyle(styleContainer, styleName) {
            return OVERWRITE_EXISTING_STYLES ||
                characterStyleKeysCreatedThisRun[styleContainerKey(styleContainer, styleName)] === true;
        }

        // =========================================
        // 個別スタイルの属性適用 / Style-specific property settings
        // =========================================
        // ※ 既存スタイルへの上書きは OVERWRITE_EXISTING_STYLES（shouldApplyAttributesTo* がガード）。basedOn は常に設定。
        //   Guarded by OVERWRITE_EXISTING_STYLES (overwrite). basedOn is always set.

        /**
         * 基準スタイルへ共通の組版設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
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

        /**
         * 基準スタイルの継承関係（basedOn）を設定する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
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

        /**
         * 段落スタイルの「次のスタイル」を設定する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
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

        /**
         * 段落の分離禁止に関する設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
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

        /**
         * 画像用段落スタイルの設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyImageParagraphSettings(doc) {
            if (!shouldApplyAttributesToParagraphStyle(doc, "p.img")) return;
            var imageParagraphStyle = doc.paragraphStyles.itemByName("p.img");
            if (imageParagraphStyle.isValid) {
                imageParagraphStyle.justification = Justification.CENTER_ALIGN;
            }
        }

        /**
         * 箇条書き・番号リストの設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
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
                        ["sameParaStyleSpacing", "spaceBetweenParagraphsUsingSameStyle", "spaceBetweenParagraphs", "spaceBetweenSameParagraphStyles", "spaceBetweenSameStyleParagraphs"], 0);
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

        /**
         * 表セル用スタイルの設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
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

        /**
         * 目次見出しスタイルの継承関係を設定する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
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

        /**
         * 目次末端スタイルの個別設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyTocLeafOverrides(doc) {
            var tocGroup = doc.paragraphStyleGroups.itemByName("toc");
            if (!tocGroup.isValid) return;
            if (!shouldApplyAttributesToParagraphStyle(tocGroup, "toc-h3")) return;
            var tocH3Style = tocGroup.paragraphStyles.itemByName("toc-h3");
            if (tocH3Style.isValid) {
                tocH3Style.keepWithNext = 0;
            }
        }

        /**
         * 環境によって有無が変わるプロパティを安全に設定する
         * @param {object} targetObject 設定先のオブジェクト
         * @param {Array<string>} candidateNames 試すプロパティ名
         * @param {*} value 設定する値
         * @returns {boolean} 設定できたら true
         */
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

        /**
         * 名前から文字スタイルを取得する
         * @param {Document} doc 対象ドキュメント
         * @param {string} styleName 文字スタイル名
         * @returns {CharacterStyle|null} 文字スタイル。見つからない場合は null
         */
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

        /**
         * インライングラフィック用の前後アキを設定する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyInlineGraphicSpacing(doc) {
            var resolved = resolveCharacterStyle(doc, "inline-graphic");
            if (!resolved) return;
            if (!shouldApplyAttributesToCharacterStyle(resolved.container, "inline-graphic")) return;
            resolved.style.leadingAki = 0.25;
            resolved.style.trailingAki = 0.25;
        }

        /**
         * リンク用文字スタイルの設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyLinkSettings(doc) {
            var resolved = resolveCharacterStyle(doc, "link");
            if (!resolved) return;
            if (!shouldApplyAttributesToCharacterStyle(resolved.container, "link")) return;
            resolved.style.underline = false;
        }

        /**
         * 分割禁止の文字スタイル設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyNoBreakSettings(doc) {
            var resolved = resolveCharacterStyle(doc, "no-break");
            if (!resolved) return;
            if (!shouldApplyAttributesToCharacterStyle(resolved.container, "no-break")) return;
            resolved.style.noBreak = true;
        }

        /**
         * 文字スタイルの継承関係（basedOn）を設定する
         * @param {Document} doc 対象ドキュメント
         * @param {string} targetStyleName 対象のスタイル名
         * @param {string} parentStyleName 継承元のスタイル名
         * @returns {void}
         */
        function applyCharacterStyleBasedOn(doc, targetStyleName, parentStyleName) {
            var target = resolveCharacterStyle(doc, targetStyleName);
            var parent = resolveCharacterStyle(doc, parentStyleName);
            if (!target || !parent) return;
            if (!shouldApplyAttributesToCharacterStyle(target.container, targetStyleName)) return;
            target.style.basedOn = parent.style;
        }

        /**
         * 候補名から言語設定を解決する
         * @param {Array<string>} languageNames 言語名の候補
         * @returns {Language|null} 言語。見つからない場合は null
         */
        function resolveLanguageByNames(languageNames) {
            for (var languageNameIndex = 0; languageNameIndex < languageNames.length; languageNameIndex++) {
                var languageEntry = app.languagesWithVendors.itemByName(languageNames[languageNameIndex]);
                if (languageEntry.isValid) return languageEntry;
            }
            return null;
        }

        var ENGLISH_USA_LANGUAGE_NAMES = ["English: USA", "英語：米国"];
        var NO_LANGUAGE_NAMES = ["[No Language]", "[言語なし]", "[なし]"];

        /**
         * lang-US スタイルに英語（米国）を設定する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyLangUSLanguageSetting(doc) {
            var resolved = resolveCharacterStyle(doc, "lang-US");
            if (!resolved) return;
            if (!shouldApplyAttributesToCharacterStyle(resolved.container, "lang-US")) return;
            var englishLanguage = resolveLanguageByNames(ENGLISH_USA_LANGUAGE_NAMES);
            if (englishLanguage) resolved.style.appliedLanguage = englishLanguage;
        }

        /**
         * コード用文字スタイルの言語設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
        function applyCodeNormalLanguageSetting(doc) {
            var resolved = resolveCharacterStyle(doc, "code-normal");
            if (!resolved) return;
            if (!shouldApplyAttributesToCharacterStyle(resolved.container, "code-normal")) return;
            var noLanguage = resolveLanguageByNames(NO_LANGUAGE_NAMES);
            if (noLanguage) resolved.style.appliedLanguage = noLanguage;
            resolved.style.ligatures = false;
        }

        /**
         * コード用段落スタイルの設定を適用する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
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

        /**
         * すべてのスタイル属性の適用処理をまとめて実行する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
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

        /**
         * 基準スタイルと ul-li に正規表現スタイルを設定する
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
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

        /**
         * 段落スタイルの並び順を整える
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
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

        /**
         * 文字スタイルの並び順を整える
         * @param {Document} doc 対象ドキュメント
         * @returns {void}
         */
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

        // 処理中はプログレスバーを表示（作成→属性→GREP→並び替えの 4 段階）/
        // Show a progress bar while processing (4 phases: create → attributes → GREP → reorder)
        var progress = createProgressWindow(4);
        try {

        progress.step(getLabel("progress.styles"));

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

        progress.step(getLabel("progress.attrs"));
        applyAllStyleAttributes(doc);

        progress.step(getLabel("progress.grep"));
        applyNestedGrepStyleSettings(doc);

        // パネル上の並び順を配列順に揃える（既存スタイルも含む） /
        // Reorder styles in the panel (including existing ones)
        progress.step(getLabel("progress.reorder"));
        reorderParagraphStyles(doc);
        reorderCharacterStyles(doc);

        } finally {
            // 例外時もプログレスパレットを確実に閉じる / Always close the palette, even on error
            progress.close();
        }
    }

    // 全処理を 1 つのアンドゥ単位にまとめて実行 /
    // Run everything as a single undo step
    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined,
        UndoModes.ENTIRE_SCRIPT, getLabel("undo.registerStyles"));

})();
