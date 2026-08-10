#target indesign

/*
 * AppendParagraphNumbering.jsx
 *
 * 同じテキストが同じ段落スタイルで繰り返すとき、親見出しの階層を見ながら末尾に連番を付けます。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AppendParagraphNumbering";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-06-30";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/AppendParagraphNumbering.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/AppendParagraphNumbering.md

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

    /* 段落スタイル名と見出しレベルの対応 / Map of paragraph style names to heading levels */
    var HEADING_LEVEL_MAP = {
        "Heading 1": 1, "h1": 1,
        "Heading 2": 2, "h2": 2,
        "Heading 3": 3, "h3": 3,
        "Heading 4": 4, "h4": 4,
        "Heading 5": 5, "h5": 5,
        "Heading 6": 6, "h6": 6
    };

    /* ナンバリング対象から除外する段落スタイル / Paragraph styles excluded from numbering */
    var IGNORE_STYLE_NAMES = ["p.img", "p.table"];

    /* 末尾ナンバリングを見つける正規表現 / Pattern that matches trailing numbering */
    var NUMBERING_PATTERN = /[（\(][0-9０-９]+[）\)]$/;

    // =========================================
    // レイアウト設定 / Layout settings
    // =========================================

    /* 対象リストのサイズ [幅, 高さ]（px）/ Size of the target list [width, height] (px) */
    var TARGET_LIST_SIZE = [400, 400];

    /* 進捗バーのサイズ [幅, 高さ]（px）/ Size of the progress bar [width, height] (px) */
    var PROGRESS_BAR_SIZE = [330, 7];

    /* リスト表示で省略を始める文字数と、省略後に残す文字数 / Length that triggers truncation, and the kept length */
    var LIST_TEXT_MAX_LENGTH  = 28;
    var LIST_TEXT_KEEP_LENGTH = 25;

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
    var currentLanguage = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "末尾にナンバリング追加", en: "Append Numbering at End" }
        },
        panel: {
            paragraphStyle: { ja: "段落スタイル", en: "Paragraph Style" },
            target: { ja: "対象", en: "Target" }
        },
        radio: {
            story: { ja: "ストーリー", en: "Story" },
            document: { ja: "ドキュメント", en: "Document" },
            fullWidth: { ja: "全角", en: "Full-width" },
            halfWidth: { ja: "半角", en: "Half-width" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            deleteItem: { ja: "削除", en: "Delete" }
        },
        progress: {
            title: { ja: "解析中", en: "Analyzing" }
        },
        alert: {
            noTargets: {
                ja: "ナンバリング対象が見つかりませんでした。",
                en: "No numbering targets were found."
            },
            notStory: {
                ja: "選択したオブジェクトはストーリーとして認識できません。ドキュメント全体を対象にします。",
                en: "The selected object is not recognized as a story. The entire document will be processed."
            },
            removed: {
                ja: "選択したテキストから番号を削除しました。",
                en: "Removed numbering from the selected text."
            }
        }
    };

    /**
     * ドット区切りキーでラベルを取得する
     * @param {string} labelKey 例: "dialog.title"
     * @returns {string} 現在の言語のラベル文字列。見つからない場合はキーをそのまま返す
     */
    function getLabel(labelKey) {
        var keyParts = labelKey.split(".");
        var node = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            if (node == null) return labelKey;
            node = node[keyParts[i]];
        }
        if (node == null) return labelKey;
        return (node[currentLanguage] != null) ? node[currentLanguage] : node.en;
    }


    // =========================================
    // ヘルパー / Helpers
    // =========================================

    /**
     * ナンバリング対象から除外する段落スタイルかどうかを判定する
     * @param {string} styleName 段落スタイル名
     * @returns {boolean} 除外対象なら true
     */
    function isIgnoredStyle(styleName) {
        for (var i = 0; i < IGNORE_STYLE_NAMES.length; i++) {
            if (IGNORE_STYLE_NAMES[i] === styleName) return true;
        }
        return false;
    }

    /**
     * 段落が親（マスター）ページ上にあるかを判定する
     * @param {Paragraph} paragraph 対象の段落
     * @returns {boolean} 親ページ上なら true
     */
    function isOnMasterSpread(paragraph) {
        if (paragraph.parentTextFrames && paragraph.parentTextFrames.length > 0) {
            var textFrame = paragraph.parentTextFrames[0];
            if (textFrame.parent.constructor.name === "MasterSpread") return true;
        }
        return false;
    }

    /**
     * 末尾の改行を取り除く
     * @param {string} text 対象の文字列
     * @returns {string} 末尾の改行を除いた文字列
     */
    function trimTrailingBreaks(text) {
        return text.replace(/[\r\n]+$/, "");
    }

    /**
     * 見出しスタックを更新しながら、段落の識別キーを作る
     * @param {Paragraph} paragraph 対象の段落
     * @param {Array<object>} headingStack 見出しの階層スタック
     * @returns {object} 識別キーと関連情報
     */
    function buildKeyForParagraph(paragraph, headingStack) {
        var content = trimTrailingBreaks(paragraph.contents);
        var styleName = paragraph.appliedParagraphStyle.name;
        var cleanedText = content.replace(NUMBERING_PATTERN, "");
        var headingLevel = HEADING_LEVEL_MAP[styleName] || 99;

        /* 現在のレベル以上の親をスタックから除去 / Pop parents at the same or deeper level */
        while (headingStack.length > 0 && headingStack[headingStack.length - 1].level >= headingLevel) {
            headingStack.pop();
        }
        /* 見出しなら親としてスタックに積む / Push headings onto the heading stack */
        if (headingLevel <= 6) {
            headingStack.push({
                key: cleanedText,
                style: styleName,
                level: headingLevel,
                counter: (headingStack.length > 0 ? headingStack[headingStack.length - 1].counter : 0) + 1
            });
        }
        var parentKey = headingStack.length > 0 ? headingStack[headingStack.length - 1].key : "";
        var parentStyleName = headingStack.length > 0 ? headingStack[headingStack.length - 1].style : "";
        var parentLevelCounter = headingStack.length > 0 ? headingStack[headingStack.length - 1].counter : 0;

        return {
            key: styleName + "___" + cleanedText + "___" + parentStyleName + "___" + parentKey + "___" + parentLevelCounter,
            style: styleName,
            cleanedText: cleanedText,
            content: content,
            parentKey: parentKey,
            parentStyle: parentStyleName,
            parentCounter: parentLevelCounter
        };
    }

    /**
     * 段落末尾の既存ナンバリングを削除する
     * @param {Paragraph} paragraph 対象の段落
     * @returns {void}
     */
    function removeExistingNumbering(paragraph) {
        var content = trimTrailingBreaks(paragraph.contents);
        var match = content.match(NUMBERING_PATTERN);
        if (!match) return;
        var lengthToRemove = match[0].length;
        var endIndex = paragraph.characters.length - 1;
        if (paragraph.characters[endIndex].contents == "\r" || paragraph.characters[endIndex].contents == "\n") {
            endIndex--;
        }
        var startIndex = endIndex - lengthToRemove + 1;
        paragraph.characters.itemByRange(startIndex, endIndex).remove();
    }


    // =========================================
    // メイン処理 / Main
    // =========================================
    /**
     * 対象を解析し、ダイアログの指定に従ってナンバリングを付与する
     * @returns {void}
     */
    function main() {
        var activeDoc = app.activeDocument;
        var allStories = activeDoc.stories;
        var occurrenceMap = {};

        /* 解析中プログレスバー / Progress bar while analyzing */
        var progressWindow = new Window("palette", getLabel("progress.title"));
        progressWindow.orientation = "column";
        progressWindow.alignChildren = ["fill", "top"];
        progressWindow.margins = WINDOW_MARGINS;
        var progressBar = progressWindow.add("progressbar", undefined, 0, allStories.length);
        progressBar.preferredSize = PROGRESS_BAR_SIZE;
        progressWindow.show();

        /* 全ストーリーを走査して同一テキストの出現を集計 / Scan all stories and count repeated text */
        for (var i = 0; i < allStories.length; i++) {
            var paragraphs = allStories[i].paragraphs;
            var headingStack = [];
            for (var j = 0; j < paragraphs.length; j++) {
                var paragraph = paragraphs[j];
                if (isOnMasterSpread(paragraph)) continue;

                var content = trimTrailingBreaks(paragraph.contents);
                var cleanedText = content.replace(NUMBERING_PATTERN, "");
                /* 空行・1文字以下・空白のみはスキップ / Skip empty, single-char, or whitespace-only */
                if (cleanedText.length <= 1) continue;
                if (cleanedText.match(/^\s+$/)) continue;

                var styleName = paragraph.appliedParagraphStyle.name;
                if (isIgnoredStyle(styleName)) continue;

                var paragraphInfo = buildKeyForParagraph(paragraph, headingStack);
                if (!occurrenceMap[paragraphInfo.key]) {
                    occurrenceMap[paragraphInfo.key] = {
                        style: paragraphInfo.style,
                        text: paragraphInfo.cleanedText,
                        count: 1,
                        parent: paragraphInfo.parentKey,
                        parentStyle: paragraphInfo.parentStyle,
                        parentCounter: paragraphInfo.parentCounter
                    };
                } else {
                    occurrenceMap[paragraphInfo.key].count++;
                }
            }
            progressBar.value = i + 1;
            progressWindow.update();
        }
        progressWindow.close();

        /* 親があり2回以上出現するものだけ対象に / Keep entries that have a parent and repeat */
        var numberingTargets = [];
        for (var mapKey in occurrenceMap) {
            var mapEntry = occurrenceMap[mapKey];
            if (mapEntry.count >= 2 && mapEntry.parent !== "") {
                numberingTargets.push(mapEntry);
            }
        }

        /* 出現回数の多い順、同数ならスタイル名順 / Sort by count desc, then style name */
        numberingTargets.sort(function (a, b) {
            if (b.count !== a.count) {
                return b.count - a.count;
            }
            return a.style.toLowerCase() < b.style.toLowerCase() ? -1 : 1;
        });

        if (numberingTargets.length === 0) {
            alert(getLabel("alert.noTargets"));
            return;
        }

        // -----------------------------------------
        // UIダイアログ / Dialog
        // -----------------------------------------
        var dialogWindow = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialogWindow, 10);

        var fullWidthBtn, halfWidthBtn;

        var contentGroup = dialogWindow.add("group");
        setupRow(contentGroup, "fill", COLUMN_SPACING);
        contentGroup.alignChildren = ["fill", "top"];

        /* 左カラム：スタイル・対象・全角半角 / Left column: styles, target, brackets */
        var optionColumn = contentGroup.add("group");
        optionColumn.orientation = "column";
        optionColumn.alignChildren = ["fill", "top"];
        optionColumn.spacing = PANEL_SPACING;

        var stylePanel = optionColumn.add("panel", undefined, getLabel("panel.paragraphStyle"));
        setupPanel(stylePanel, 6);
        stylePanel.alignChildren = ["left", "top"];

        var targetPanel = optionColumn.add("panel", undefined, getLabel("panel.target"));
        setupPanel(targetPanel, 6);
        targetPanel.orientation = "row";
        targetPanel.alignChildren = ["left", "top"];
        var storyRadio = targetPanel.add("radiobutton", undefined, getLabel("radio.story"));
        var documentRadio = targetPanel.add("radiobutton", undefined, getLabel("radio.document"));
        storyRadio.value = true;

        /* 全角／半角選択（日本語UIのみ）/ Full/half-width selection (Japanese UI only) */
        if (currentLanguage === "ja") {
            var bracketRadioGroup = optionColumn.add("group");
            setupRow(bracketRadioGroup, "center", 8);
            fullWidthBtn = bracketRadioGroup.add("radiobutton", undefined, getLabel("radio.fullWidth"));
            halfWidthBtn = bracketRadioGroup.add("radiobutton", undefined, getLabel("radio.halfWidth"));
            fullWidthBtn.value = true;
        }

        /* 右カラム：対象リスト / Right column: target list */
        var listColumn = contentGroup.add("group");
        listColumn.orientation = "column";
        listColumn.alignChildren = ["fill", "top"];
        var targetListBox = listColumn.add("listbox", undefined, "", { multiselect: true });
        targetListBox.preferredSize = TARGET_LIST_SIZE;

        for (var i = 0; i < numberingTargets.length; i++) {
            var baseText = numberingTargets[i].text;
            var displayText = baseText.length > LIST_TEXT_MAX_LENGTH ? baseText.substring(0, LIST_TEXT_KEEP_LENGTH) + "…" : baseText;
            var countText = (currentLanguage === "ja") ? "（" + numberingTargets[i].count + "）" : " (" + numberingTargets[i].count + ")";
            var listItem = targetListBox.add("item", numberingTargets[i].style + ": " + displayText + countText);
            listItem.helpTip = numberingTargets[i].text;
        }

        /* 対象スタイルのチェックボックスを生成 / Build checkboxes for the involved styles */
        var styleNameSet = {};
        for (var i = 0; i < numberingTargets.length; i++) {
            styleNameSet[numberingTargets[i].style] = true;
        }
        var sortedStyleNames = [];
        for (var styleNameKey in styleNameSet) {
            sortedStyleNames.push(styleNameKey);
        }
        sortedStyleNames.sort();
        var styleCheckboxes = {};
        for (var i = 0; i < sortedStyleNames.length; i++) {
            var checkboxStyleName = sortedStyleNames[i];
            var styleCheckbox = stylePanel.add("checkbox", undefined, checkboxStyleName);
            styleCheckbox.value = true;
            styleCheckboxes[checkboxStyleName] = styleCheckbox;
        }
        stylePanel.layout.layout(true);

        /**
         * 対象リストの有効／無効を現在の設定に合わせて切り替える
         * @returns {void}
         */
        function updateListBoxEnabled() {
            for (var i = 0; i < targetListBox.items.length; i++) {
                var listItem = targetListBox.items[i];
                var styleName = numberingTargets[i].style;
                if (styleCheckboxes[styleName].value) {
                    listItem.enabled = true;
                } else {
                    listItem.enabled = false;
                    listItem.selected = false;
                }
            }
        }
        for (var checkboxStyle in styleCheckboxes) {
            styleCheckboxes[checkboxStyle].onClick = updateListBoxEnabled;
        }
        if (targetListBox.items.length > 0) {
            targetListBox.items[0].selected = true;
        }

        /**
         * リストで選択中の識別キーを集める
         * @returns {object} 識別キーをキーに持つマップ
         */
        function getSelectedKeys() {
            var selectedKeyMap = {};
            for (var i = 0; i < targetListBox.items.length; i++) {
                if (targetListBox.items[i].selected) {
                    var target = numberingTargets[i];
                    var selectedKey = target.style + "___" + target.text + "___" + (target.parentStyle || "") + "___" + (target.parent || "") + "___" + (target.parentCounter || 0);
                    selectedKeyMap[selectedKey] = 1;
                }
            }
            return selectedKeyMap;
        }

        /**
         * 対象範囲に応じて処理するストーリーを求める
         * @returns {Array<Story>} 対象のストーリー
         */
        function getTargetStories() {
            if (storyRadio.value && app.selection.length > 0) {
                var selectionItem = app.selection[0];
                var parentStory = null;
                if (selectionItem.hasOwnProperty("parentStory")) {
                    parentStory = selectionItem.parentStory;
                } else if (selectionItem.parent && selectionItem.parent.hasOwnProperty("parentStory")) {
                    parentStory = selectionItem.parent.parentStory;
                }
                if (parentStory) {
                    return [parentStory];
                }
                alert(getLabel("alert.notStory"));
                return allStories;
            }
            return allStories;
        }

        // -----------------------------------------
        // ボタン / Buttons
        // -----------------------------------------
        /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
        var buttonGroup = dialogWindow.add("group");
        setupRow(buttonGroup, "fill", 8);
        buttonGroup.alignChildren = ["fill", "center"];
        var leftButtonGroup = buttonGroup.add("group");
        setupRow(leftButtonGroup, "left", 8);
        var cancelBtn = leftButtonGroup.add("button", undefined, getLabel("button.cancel"));
        var rightButtonGroup = buttonGroup.add("group");
        setupRow(rightButtonGroup, "right", 8);
        var deleteBtn = rightButtonGroup.add("button", undefined, getLabel("button.deleteItem"));

        /* 選択テキストから既存ナンバリングを削除（undoは1ステップ）/ Remove numbering from selected text (single undo) */
        deleteBtn.onClick = function () {
            var targetStories = getTargetStories();
            var selectedKeyMap = getSelectedKeys();
            /**
             * 選択した対象から既存のナンバリングを削除する
             * @returns {void}
             */
            function removeNumberingTask() {
                for (var i = 0; i < targetStories.length; i++) {
                    var paragraphs = targetStories[i].paragraphs;
                    var headingStack = [];
                    for (var j = 0; j < paragraphs.length; j++) {
                        var paragraph = paragraphs[j];
                        if (isOnMasterSpread(paragraph)) continue;
                        var styleName = paragraph.appliedParagraphStyle.name;
                        if (isIgnoredStyle(styleName)) continue;
                        var paragraphInfo = buildKeyForParagraph(paragraph, headingStack);
                        if (!(paragraphInfo.key in selectedKeyMap)) continue;
                        removeExistingNumbering(paragraph);
                    }
                }
            }
            app.doScript(removeNumberingTask, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Remove Paragraph Numbering");
            alert(getLabel("alert.removed"));
        };

        var okBtn = rightButtonGroup.add("button", undefined, "OK");

        var dialogResult = dialogWindow.show();
        if (dialogResult != 1) return;

        /* 括弧の全角／半角 / Full or half-width brackets */
        var leftParen = "（";
        var rightParen = "）";
        if (currentLanguage === "ja" && halfWidthBtn && halfWidthBtn.value) {
            leftParen = "(";
            rightParen = ")";
        }

        var targetStories = getTargetStories();
        var selectedKeyMap = getSelectedKeys();

        /**
         * 選択した対象の末尾にナンバリングを付与する
         * @returns {void}
         */
        function applyNumberingTask() {
            var counterByKey = {};
            for (var seedKey in selectedKeyMap) {
                counterByKey[seedKey] = 1;
            }
            for (var i = 0; i < targetStories.length; i++) {
                var paragraphs = targetStories[i].paragraphs;
                var headingStack = [];
                for (var j = 0; j < paragraphs.length; j++) {
                    var paragraph = paragraphs[j];
                    if (isOnMasterSpread(paragraph)) continue;
                    var styleName = paragraph.appliedParagraphStyle.name;
                    if (isIgnoredStyle(styleName)) continue;

                    var paragraphInfo = buildKeyForParagraph(paragraph, headingStack);
                    if (!(paragraphInfo.key in counterByKey)) continue;

                    removeExistingNumbering(paragraph);

                    var insertIndex = paragraph.characters.length;
                    if (insertIndex > 0 && (paragraph.characters[insertIndex - 1].contents == "\r" || paragraph.characters[insertIndex - 1].contents == "\n")) {
                        insertIndex--;
                    }
                    paragraph.insertionPoints[insertIndex].contents = leftParen + String(counterByKey[paragraphInfo.key]) + rightParen;
                    counterByKey[paragraphInfo.key]++;
                }
            }
        }
        app.doScript(applyNumberingTask, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Append Paragraph Numbering");
    }

    main();

})();
