#target indesign

/*

### 概要

- 同じテキストが同じ段落スタイルで繰り返すとき、末尾にナンバリングを追加します
- 親見出しの階層を考慮し、親単位で同一テキストを正確に判別します
- 段落スタイル選択、ストーリー／ドキュメント範囲、全角／半角括弧の切り替えに対応（全角／半角は日本語UIのみ）
- 既存ナンバリングの削除にも対応

### 限定条件

- InDesignで動作
- 対象は開いているアクティブドキュメント

*/

/*

### Overview

- When the same text repeats with the same paragraph style, appends numbering at the end
- Considers heading hierarchy to accurately judge repetitions per parent
- Supports paragraph style selection, story/document scope, and full-width/half-width brackets (full/half is Japanese UI only)
- Includes numbering removal

### Limitations

- Runs in InDesign
- Targets the active (open) document

*/


// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.1.2";

/*
    更新履歴 / Changelog:
    - v1.0.0 (2025-06-29): 初版
    - v1.0.7 (2025-07-01): p.img, p.table スタイルの段落を無視
    - v1.0.8 (2025-07-02): 段落スタイル選択パネル追加
    - v1.0.9 (2025-07-03): 削除ボタン追加
    - v1.1.0 (2025-07-04): 階層判定ロジック改良、UIラベル整理
    - v1.1.1 (2026-06-30): 対象ラジオが無視されるバグ修正、削除をundoにまとめる、
                           instanceof を constructor.name に変更、IIFE化・ローカライズ整理
    - v1.1.2 (2026-06-30): 変数・パネル・関数名を具体名に整理、コメント拡充、jsx/page/ へ移動
*/


(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 段落見出しレベル設定 / Heading level map */
    var headingLevelMap = {
        "Heading 1": 1, "h1": 1,
        "Heading 2": 2, "h2": 2,
        "Heading 3": 3, "h3": 3,
        "Heading 4": 4, "h4": 4,
        "Heading 5": 5, "h5": 5,
        "Heading 6": 6, "h6": 6
    };

    /* ナンバリング対象から除外する段落スタイル / Paragraph styles to ignore */
    var ignoreStyleNames = ["p.img", "p.table"];

    /* 末尾ナンバリングの正規表現 / Trailing numbering pattern */
    var numberingRegex = /[（\(][0-9０-９]+[）\)]$/;


    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在の言語を判定 / Detect current language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
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

    /* ドット区切りのキーで LABELS を引く / Look up a label by dotted key */
    function L(path) {
        var parts = path.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node == null) return path;
            node = node[parts[i]];
        }
        if (node == null) return path;
        return (node[currentLanguage] != null) ? node[currentLanguage] : node.en;
    }


    // =========================================
    // ヘルパー / Helpers
    // =========================================

    /* 除外スタイルかどうか / Whether the style is in the ignore list */
    function isIgnoredStyle(styleName) {
        for (var i = 0; i < ignoreStyleNames.length; i++) {
            if (ignoreStyleNames[i] === styleName) return true;
        }
        return false;
    }

    /* マスターページ上の段落か / Whether the paragraph sits on a master spread */
    function isOnMasterSpread(paragraph) {
        if (paragraph.parentTextFrames && paragraph.parentTextFrames.length > 0) {
            var textFrame = paragraph.parentTextFrames[0];
            if (textFrame.parent.constructor.name === "MasterSpread") return true;
        }
        return false;
    }

    /* 末尾の改行を除いた本文を取得 / Get contents without trailing line breaks */
    function trimTrailingBreaks(text) {
        return text.replace(/[\r\n]+$/, "");
    }

    /* 段落から識別キーを生成（見出しスタックを更新しながら）/ Build an identity key, updating the heading stack */
    function buildKeyForParagraph(paragraph, headingStack) {
        var content = trimTrailingBreaks(paragraph.contents);
        var styleName = paragraph.appliedParagraphStyle.name;
        var cleanedText = content.replace(numberingRegex, "");
        var headingLevel = headingLevelMap[styleName] || 99;

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

    /* 既存のナンバリングを段落末尾から削除 / Remove existing numbering at the paragraph end */
    function removeExistingNumbering(paragraph) {
        var content = trimTrailingBreaks(paragraph.contents);
        var match = content.match(numberingRegex);
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
    function main() {
        var activeDoc = app.activeDocument;
        var allStories = activeDoc.stories;
        var occurrenceMap = {};

        /* 解析中プログレスバー / Progress bar while analyzing */
        var progressWindow = new Window("palette", L("progress.title"));
        progressWindow.orientation = "column";
        progressWindow.alignChildren = ["fill", "top"];
        progressWindow.margins = [20, 20, 20, 20];
        var progressBar = progressWindow.add("progressbar", undefined, 0, allStories.length);
        progressBar.preferredSize = [330, 7];
        progressWindow.show();

        /* 全ストーリーを走査して同一テキストの出現を集計 / Scan all stories and count repeated text */
        for (var i = 0; i < allStories.length; i++) {
            var paragraphs = allStories[i].paragraphs;
            var headingStack = [];
            for (var j = 0; j < paragraphs.length; j++) {
                var paragraph = paragraphs[j];
                if (isOnMasterSpread(paragraph)) continue;

                var content = trimTrailingBreaks(paragraph.contents);
                var cleanedText = content.replace(numberingRegex, "");
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
            alert(L("alert.noTargets"));
            return;
        }

        // -----------------------------------------
        // UIダイアログ / Dialog
        // -----------------------------------------
        var dialogWindow = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
        dialogWindow.orientation = "column";
        dialogWindow.alignChildren = ["fill", "top"];

        var fullWidthBtn, halfWidthBtn;

        var contentGroup = dialogWindow.add("group");
        contentGroup.orientation = "row";
        contentGroup.alignChildren = ["fill", "top"];

        /* 左カラム：スタイル・対象・全角半角 / Left column: styles, target, brackets */
        var optionColumn = contentGroup.add("group");
        optionColumn.orientation = "column";
        optionColumn.alignChildren = ["fill", "top"];

        var stylePanel = optionColumn.add("panel", undefined, L("panel.paragraphStyle"));
        stylePanel.orientation = "column";
        stylePanel.alignChildren = ["left", "top"];
        stylePanel.margins = [15, 20, 15, 10];

        var targetPanel = optionColumn.add("panel", undefined, L("panel.target"));
        targetPanel.orientation = "row";
        targetPanel.alignChildren = ["left", "top"];
        var storyRadio = targetPanel.add("radiobutton", undefined, L("radio.story"));
        var documentRadio = targetPanel.add("radiobutton", undefined, L("radio.document"));
        targetPanel.margins = [15, 20, 15, 10];
        storyRadio.value = true;

        /* 全角／半角選択（日本語UIのみ）/ Full/half-width selection (Japanese UI only) */
        if (currentLanguage === "ja") {
            var bracketRadioGroup = optionColumn.add("group");
            bracketRadioGroup.orientation = "row";
            bracketRadioGroup.alignment = "center";
            fullWidthBtn = bracketRadioGroup.add("radiobutton", undefined, L("radio.fullWidth"));
            halfWidthBtn = bracketRadioGroup.add("radiobutton", undefined, L("radio.halfWidth"));
            fullWidthBtn.value = true;
        }

        /* 右カラム：対象リスト / Right column: target list */
        var listColumn = contentGroup.add("group");
        listColumn.orientation = "column";
        listColumn.alignChildren = ["fill", "top"];
        var targetListBox = listColumn.add("listbox", undefined, "", { multiselect: true });
        targetListBox.preferredSize = [400, 400];

        for (var i = 0; i < numberingTargets.length; i++) {
            var baseText = numberingTargets[i].text;
            var displayText = baseText.length > 28 ? baseText.substring(0, 25) + "…" : baseText;
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

        /* チェックボックス連動でリスト項目の有効/無効を切り替え / Toggle list items per checkbox */
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

        /* 選択中の項目からキー集合を取得 / Collect keys of the selected items */
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

        /* 処理対象のストーリーを取得（ダイアログの選択を都度反映）/ Resolve target stories per current selection */
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
                alert(L("alert.notStory"));
                return allStories;
            }
            return allStories;
        }

        // -----------------------------------------
        // ボタン / Buttons
        // -----------------------------------------
        var buttonGroup = dialogWindow.add("group");
        buttonGroup.alignment = "fill";
        buttonGroup.alignChildren = ["fill", "center"];
        var leftButtonGroup = buttonGroup.add("group");
        leftButtonGroup.alignment = "left";
        var cancelBtn = leftButtonGroup.add("button", undefined, L("button.cancel"));
        var rightButtonGroup = buttonGroup.add("group");
        rightButtonGroup.alignment = ["right", "center"];
        var deleteBtn = rightButtonGroup.add("button", undefined, L("button.deleteItem"));

        /* 選択テキストから既存ナンバリングを削除（undoは1ステップ）/ Remove numbering from selected text (single undo) */
        deleteBtn.onClick = function () {
            var targetStories = getTargetStories();
            var selectedKeyMap = getSelectedKeys();
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
            alert(L("alert.removed"));
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

        /* 選択テキストの末尾に連番を付与 / Append sequential numbering to selected text */
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
