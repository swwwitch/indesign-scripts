#target indesign

/*
    スクリプト名：AppendParagraphNumbering.jsx

    概要 (日本語):
    同じテキストが同じ段落スタイルで繰り返すとき、末尾にナンバリングします。
    UIで全角/半角括弧の選択が可能（日本語UIのみ）。
    検索時には全角・半角数字を認識し、挿入時には半角数字を使用します。

    Description (English):
    When the same text repeats with the same paragraph style, this script appends numbering at the end.
    Recognizes both full-width and half-width digits when searching, always inserts half-width numbers.

    限定条件:
    - InDesignで動作
    - 対象は開いているアクティブドキュメント

    作成日：2025-06-29
    更新履歴:
    -v1.0.0 (2025-06-29): 初版
    -v1.0.7 (2025-07-01): p.img, p.tableスタイルの段落を無視するよう修正
    -v1.0.8 (2025-07-02): 段落スタイルをリストアップし、リンクのない段落スタイルをチェックボックスで選択可能に
    -v1.0.9 (2025-07-03): ［削除］ボタンを追加
*/

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var lang = getCurrentLang();
var LABELS = {
    confirmTitle: {
        ja: "末尾にナンバリング追加",
        en: "Append Numbering at End"
    },
    target: {
        ja: "対象",
        en: "Target"
    },
    story: {
        ja: "ストーリー",
        en: "Story"
    },
    document: {
        ja: "ドキュメント",
        en: "Document"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    deleteBtn: {
        ja: "削除",
        en: "Delete"
    }
};

function main() {
    var doc = app.activeDocument;
    var stories = doc.stories;
    var comboMap = {};

    // プログレスバー表示
    var progressWin = new Window("palette", "解析中");
    progressWin.orientation = "column";
    progressWin.alignChildren = ["fill", "top"];
    progressWin.margins = [20, 20, 20, 20];
    var progressBar = progressWin.add("progressbar", undefined, 0, stories.length);
    progressBar.preferredSize = [330, 7];
    progressWin.show();

    for (var i = 0; i < stories.length; i++) {
        var paragraphs = stories[i].paragraphs;
        for (var j = 0; j < paragraphs.length; j++) {
            var para = paragraphs[j];
            // マスターページ上の段落はスキップ
            if (para.parentTextFrames && para.parentTextFrames.length > 0) {
                var tf = para.parentTextFrames[0];
                if (tf.parent instanceof MasterSpread) continue;
            }
            var content = para.contents.replace(/[\r\n]+$/, "");
            var cleaned = content.replace(/[（\(][0-9０-９]+[）\)]$/, "");
            if (!cleaned.match(/^(.+)$/)) continue;

            var styleName = para.appliedParagraphStyle.name;
            if (styleName == "p.img" || styleName == "p.table") continue;
            if (cleaned === "") continue;

            var key = styleName + "___" + cleaned;
            if (!comboMap[key]) {
                comboMap[key] = {
                    style: styleName,
                    text: cleaned,
                    count: 1
                };
            } else {
                comboMap[key].count++;
            }
        }
        progressBar.value = i + 1;
        progressWin.update();
    }

    progressWin.close();

    var targets = [];
    for (var key in comboMap) {
        if (comboMap[key].count >= 2) {
            targets.push(comboMap[key]);
        }
    }

    targets.sort(function(a, b) {
        if (b.count !== a.count) {
            return b.count - a.count;
        }
        return a.style.toLowerCase() < b.style.toLowerCase() ? -1 : 1;
    });

    if (targets.length === 0) {
        alert("ナンバリング対象が見つかりませんでした。");
        return;
    }

    var dialog = new Window("dialog", LABELS.confirmTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = ["fill", "top"];

    var fullWidthBtn, halfWidthBtn;

    var mainGroup = dialog.add("group");
    mainGroup.orientation = "row";
    mainGroup.alignChildren = ["fill", "top"];

    // 左カラム
    var leftGroup = mainGroup.add("group");
    leftGroup.orientation = "column";
    leftGroup.alignChildren = ["fill", "top"];

    // 段落スタイルパネル（チェックボックスを追加）
    var stylePanel = leftGroup.add("panel", undefined, "段落スタイル");
    stylePanel.orientation = "column";
    stylePanel.alignChildren = ["left", "top"];
    stylePanel.margins = [15, 20, 15, 10];

    // 対象パネル
    var targetGroup = leftGroup.add("panel", undefined, LABELS.target[lang]);
    targetGroup.orientation = "row";
    targetGroup.alignChildren = ["left", "top"];
    var storyRadio = targetGroup.add("radiobutton", undefined, LABELS.story[lang]);
    var documentRadio = targetGroup.add("radiobutton", undefined, LABELS.document[lang]);
    targetGroup.margins = [15, 20, 15, 10];
    storyRadio.value = true;

    // 全角/半角選択（日本語UIのみ）
    if (lang === "ja") {
        var radioGroup = leftGroup.add("group");
        radioGroup.orientation = "row";
        radioGroup.alignment = "center";
        fullWidthBtn = radioGroup.add("radiobutton", undefined, "全角");
        halfWidthBtn = radioGroup.add("radiobutton", undefined, "半角");
        fullWidthBtn.value = true;
    }

    // 右カラム
    var rightGroup = mainGroup.add("group");
    rightGroup.orientation = "column";
    rightGroup.alignChildren = ["fill", "top"];
    var listBox = rightGroup.add("listbox", undefined, "", {
        multiselect: true
    });
    listBox.preferredSize = [400, 400];

    for (var i = 0; i < targets.length; i++) {
        var baseText = targets[i].text;
        var displayText = baseText.length > 28 ? baseText.substring(0, 25) + "…" : baseText;
        var label = targets[i].style + ": " + displayText + "（" + targets[i].count + "）";
        var item = listBox.add("item", label);
        item.helpTip = targets[i].text;
    }

    // 対象の段落スタイルを抽出し、チェックボックスを追加
    var styleSet = {};
    for (var i = 0; i < targets.length; i++) {
        styleSet[targets[i].style] = true;
    }

    var sortedStyles = [];
    for (var styleName in styleSet) {
        sortedStyles.push(styleName);
    }
    sortedStyles.sort();

    var styleCheckboxes = {};
    for (var i = 0; i < sortedStyles.length; i++) {
        var styleName = sortedStyles[i];
        var cb = stylePanel.add("checkbox", undefined, styleName);
        cb.value = true;
        styleCheckboxes[styleName] = cb;
    }
    stylePanel.layout.layout(true);

    // listBoxのアイテムの有効/無効をチェックボックスに連動
    function updateListBoxEnabled() {
        for (var i = 0; i < listBox.items.length; i++) {
            var item = listBox.items[i];
            var styleName = targets[i].style;
            if (styleCheckboxes[styleName].value) {
                item.enabled = true;
            } else {
                item.enabled = false;
                item.selected = false;
            }
        }
    }

    for (var style in styleCheckboxes) {
        styleCheckboxes[style].onClick = updateListBoxEnabled;
    }

    if (listBox.items.length > 0) {
        listBox.items[0].selected = true;
    }

    // 選択された項目のキーを取得するヘルパー関数
    function getSelectedKeys() {
        var keys = {};
        for (var i = 0; i < listBox.items.length; i++) {
            if (listBox.items[i].selected) {
                var key = targets[i].style + "___" + targets[i].text;
                keys[key] = 1;
            }
        }
        return keys;
    }

    // 選択されたストーリーを取得するヘルパー関数
    function getProcessStories() {
        var processStories = [];
        if (storyRadio.value && app.selection.length > 0) {
            var sel = app.selection[0];
            var parentStory = null;

            if (sel.hasOwnProperty("parentStory")) {
                parentStory = sel.parentStory;
            } else if (sel.parent && sel.parent.hasOwnProperty("parentStory")) {
                parentStory = sel.parent.parentStory;
            }

            if (parentStory) {
                processStories = [parentStory];
            } else {
                alert("選択したオブジェクトはストーリーとして認識できません。ドキュメント全体を対象にします。");
                processStories = stories;
            }
        } else {
            processStories = stories;
        }
        return processStories;
    }

    // 既存のナンバリングを削除する共通関数
    function removeExistingNumbering(para) {
        var regex = /[（\(][0-9０-９]+[）\)]$/;
        var content = para.contents.replace(/[\r\n]+$/, "");
        if (regex.test(content)) {
            var match = content.match(regex);
            if (match) {
                var lengthToRemove = match[0].length;
                var endIndex = para.characters.length - 1;
                if (para.characters[endIndex].contents == "\r" || para.characters[endIndex].contents == "\n") {
                    endIndex--;
                }
                var startIndex = endIndex - lengthToRemove + 1;
                para.characters.itemByRange(startIndex, endIndex).remove();
            }
        }
    }

    var processStories = getProcessStories();

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "fill";
    buttonGroup.alignChildren = ["fill", "center"];

    var leftButtons = buttonGroup.add("group");
    leftButtons.alignment = "left";
    var cancelBtn = leftButtons.add("button", undefined, LABELS.cancel[lang]);

    var rightButtons = buttonGroup.add("group");
    rightButtons.alignment = ["right", "center"];
    var deleteBtn = rightButtons.add("button", undefined, LABELS.deleteBtn[lang]);

    deleteBtn.onClick = function() {
        var selectedKeys = getSelectedKeys();
        var delCounts = {};
        for (var key in selectedKeys) {
            delCounts[key] = true;
        }

        for (var i = 0; i < processStories.length; i++) {
            var paragraphs = processStories[i].paragraphs;
            for (var j = 0; j < paragraphs.length; j++) {
                var para = paragraphs[j];
                if (para.parentTextFrames && para.parentTextFrames.length > 0) {
                    var tf = para.parentTextFrames[0];
                    if (tf.parent instanceof MasterSpread) continue;
                }
                var content = para.contents.replace(/[\r\n]+$/, "");
                var styleName = para.appliedParagraphStyle.name;
                if (styleName == "p.img" || styleName == "p.table") continue;
                var regex = /[（\(][0-9０-９]+[）\)]$/;
                var cleaned = content.replace(regex, "");
                var key = styleName + "___" + cleaned;

                if (!(key in delCounts)) continue;

                removeExistingNumbering(para);
            }
        }

        alert("選択したテキストから番号を削除しました。");
    };

    var okBtn = rightButtons.add("button", undefined, LABELS.ok[lang]);

    var result = dialog.show();
    if (result != 1) return;

    var leftParen = "（";
    var rightParen = "）";
    if (lang === "ja" && halfWidthBtn && halfWidthBtn.value) {
        leftParen = "(";
        rightParen = ")";
    }

    var selectedKeys = getSelectedKeys();

    function applyNumbering() {
        var localCounts = {};
        for (var key in selectedKeys) {
            localCounts[key] = 1;
        }

        for (var i = 0; i < processStories.length; i++) {
            var paragraphs = processStories[i].paragraphs;
            for (var j = 0; j < paragraphs.length; j++) {
                var para = paragraphs[j];
                // マスターページ上の段落はスキップ
                if (para.parentTextFrames && para.parentTextFrames.length > 0) {
                    var tf = para.parentTextFrames[0];
                    if (tf.parent instanceof MasterSpread) continue;
                }
                var content = para.contents.replace(/[\r\n]+$/, "");
                var styleName = para.appliedParagraphStyle.name;
                if (styleName == "p.img" || styleName == "p.table") continue;
                var regex = /[（\(][0-9０-９]+[）\)]$/;
                var cleaned = content.replace(regex, "");
                var key = styleName + "___" + cleaned;

                if (!(key in localCounts)) continue;

                if (regex.test(content)) {
                    var match = content.match(regex);
                    if (match) {
                        var lengthToRemove = match[0].length;
                        var endIndex = para.characters.length - 1;
                        if (para.characters[endIndex].contents == "\r" || para.characters[endIndex].contents == "\n") {
                            endIndex--;
                        }
                        var startIndex = endIndex - lengthToRemove + 1;
                        para.characters.itemByRange(startIndex, endIndex).remove();
                    }
                }

                var insertPos = para.characters.length;
                if (insertPos > 0 && (para.characters[insertPos - 1].contents == "\r" || para.characters[insertPos - 1].contents == "\n")) {
                    insertPos--;
                }

                para.insertionPoints[insertPos].contents = leftParen + String(localCounts[key]) + rightParen;
                localCounts[key]++;
            }
        }
    }

    app.doScript(applyNumbering, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Append Paragraph Numbering");
}

main();