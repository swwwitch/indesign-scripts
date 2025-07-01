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
    -v1.0.7 (2025-07-01):p.img, p.tableスタイルの段落を無視するよう修正
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
    }
};

function main() {
    var doc = app.activeDocument;
    var stories = doc.stories;
    var comboMap = {};

    // 段落内容とスタイルを分析
    for (var i = 0; i < stories.length; i++) {
        var paragraphs = stories[i].paragraphs;
        for (var j = 0; j < paragraphs.length; j++) {
            var para = paragraphs[j];
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
    }

    var targets = [];
    for (var key in comboMap) {
        if (comboMap[key].count >= 2) {
            targets.push(comboMap[key]);
        }
    }

    targets.sort(function(a, b) {
        return b.count - a.count;
    });

    if (targets.length === 0) {
        alert("ナンバリング対象が見つかりませんでした。");
        return;
    }

    var dialog = new Window("dialog", LABELS.confirmTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = ["fill", "top"];

    // 対象パネル
    var targetGroup = dialog.add("panel", undefined, LABELS.target[lang]);
    targetGroup.orientation = "row";
    targetGroup.alignChildren = ["left", "top"];
    var storyRadio = targetGroup.add("radiobutton", undefined, LABELS.story[lang]);
    var documentRadio = targetGroup.add("radiobutton", undefined, LABELS.document[lang]);
    targetGroup.margins = [15, 20, 15, 10];
    storyRadio.value = true;

    var listBox = dialog.add("listbox", undefined, "", {
        multiselect: true
    });
    listBox.preferredSize = [400, 150];

    for (var i = 0; i < targets.length; i++) {
        var baseText = targets[i].text;
        var displayText = baseText;

        if (displayText.length > 28) {
            displayText = displayText.substring(0, 25) + "…";
        }

        var label = targets[i].style + ": " + displayText + "（" + targets[i].count + "）";
        var item = listBox.add("item", label);
        item.helpTip = targets[i].text;
    }

    if (listBox.items.length > 0) {
        listBox.items[0].selected = true;
    }

    var fullWidthBtn, halfWidthBtn;
    if (lang === "ja") {
        var radioGroup = dialog.add("group");
        radioGroup.orientation = "row";
        radioGroup.alignment = "center";
        fullWidthBtn = radioGroup.add("radiobutton", undefined, "全角");
        halfWidthBtn = radioGroup.add("radiobutton", undefined, "半角");
        fullWidthBtn.value = true;
    }

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.add("button", undefined, LABELS.cancel[lang], {
        name: "cancel"
    });
    buttonGroup.add("button", undefined, LABELS.ok[lang], {
        name: "ok"
    });

    var result = dialog.show();
    if (result != 1) return;

    var leftParen = "（";
    var rightParen = "）";
    if (lang === "ja" && halfWidthBtn && halfWidthBtn.value) {
        leftParen = "(";
        rightParen = ")";
    }

    var selectedKeys = {};
    for (var i = 0; i < listBox.items.length; i++) {
        if (listBox.items[i].selected) {
            var key = targets[i].style + "___" + targets[i].text;
            selectedKeys[key] = 1;
        }
    }

    var processStories = [];
    if (storyRadio.value && app.selection.length > 0) {
        // selection[0] がテキストフレーム、テキスト、あるいは段落オブジェクトのときのみ parentStory を取得
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

    function applyNumbering() {
        var localCounts = {};
        for (var key in selectedKeys) {
            localCounts[key] = 1;
        }

        for (var i = 0; i < processStories.length; i++) {
            var paragraphs = processStories[i].paragraphs;
            for (var j = 0; j < paragraphs.length; j++) {
                var para = paragraphs[j];
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