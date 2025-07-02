#target indesign

/*
    スクリプト名：AppendParagraphNumbering.jsx

    概要 (日本語):
    同じテキストが同じ段落スタイルで繰り返すとき、末尾にナンバリングを追加します。
    親見出しの階層を考慮し、親単位で同一テキストを正確に判別します。
    UIで段落スタイル選択、ストーリー/ドキュメント範囲選択、全角/半角括弧切り替えが可能（日本語UIのみ）。
    既存のナンバリング削除機能も搭載。

    Description (English):
    When the same text repeats with the same paragraph style, this script appends numbering at the end.
    Considers heading hierarchy to accurately judge repetitions per parent.
    Supports paragraph style selection, story/document scope, full-width/half-width brackets (Japanese UI only).
    Includes numbering removal.

    限定条件:
    - InDesignで動作
    - 対象は開いているアクティブドキュメント

    作成日：2025-06-29
    更新履歴:
    -v1.0.0 (2025-06-29): 初版
    -v1.0.7 (2025-07-01): p.img, p.tableスタイルの段落を無視するよう修正
    -v1.0.8 (2025-07-02): 段落スタイル選択パネル追加
    -v1.0.9 (2025-07-03): 削除ボタン追加
    -v1.1.0 (2025-07-04): 階層判定ロジック改良、UIラベル整理
*/


// --- コードの最適化提案 ---
// ・headingLevelMap を外部設定化または共通ファイル化
// ・progressWin や listBox 初期化処理を関数化
// ・comboMap 処理を別関数に分離しテスト可能にする
// ・ラベルや UI 構造を多言語対応ファイルとして分離

// --- ラベル定義 (UIに出る順に並べる) ---
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var lang = getCurrentLang();
var LABELS = {
    confirmTitle: { ja: "末尾にナンバリング追加", en: "Append Numbering at End" },
    target: { ja: "対象", en: "Target" },
    story: { ja: "ストーリー", en: "Story" },
    document: { ja: "ドキュメント", en: "Document" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    deleteBtn: { ja: "削除", en: "Delete" }
};

// --- 段落見出しレベル設定 ---
var headingLevelMap = {
    "Heading 1": 1, "h1": 1,
    "Heading 2": 2, "h2": 2,
    "Heading 3": 3, "h3": 3,
    "Heading 4": 4, "h4": 4,
    "Heading 5": 5, "h5": 5,
    "Heading 6": 6, "h6": 6
};

function main() {
    var doc = app.activeDocument;
    var stories = doc.stories;
    var comboMap = {};

    // --- プログレスバー表示 ---
    var progressWin = new Window("palette", "解析中");
    progressWin.orientation = "column";
    progressWin.alignChildren = ["fill", "top"];
    progressWin.margins = [20, 20, 20, 20];
    var progressBar = progressWin.add("progressbar", undefined, 0, stories.length);
    progressBar.preferredSize = [330, 7];
    progressWin.show();

    for (var i = 0; i < stories.length; i++) {
        var paragraphs = stories[i].paragraphs;
        // 親段落スタイル情報を保持
        var currentParentKey = "";
        var currentParentStyleName = "";
        var lastHeadingLevel = 0;
        var parentCounter = 0; // 親出現カウンタ
        // --- 親階層スタックで親情報を管理 ---
        var parentStack = [];
        for (var j = 0; j < paragraphs.length; j++) {
            var para = paragraphs[j];
            // マスターページ上の段落はスキップ
            if (para.parentTextFrames && para.parentTextFrames.length > 0) {
                var tf = para.parentTextFrames[0];
                if (tf.parent instanceof MasterSpread) continue;
            }
            var content = para.contents.replace(/[\r\n]+$/, "");
            // 末尾のナンバリングを除去
            var cleaned = content.replace(/[（\(][0-9０-９]+[）\)]$/, "");
            // 空行や1文字以下、空白のみはスキップ
            if (!cleaned.match(/^(.+)$/)) continue;
            if (cleaned.length <= 1) continue;
            if (cleaned.match(/^\s+$/)) continue;

            var styleName = para.appliedParagraphStyle.name;
            // イメージ・テーブル用スタイルはスキップ
            if (styleName == "p.img" || styleName == "p.table") continue;
            if (cleaned === "") continue;

            var level = headingLevelMap ? (headingLevelMap[styleName] || 99) : 99;
            // スタックから現在のレベル以上のものをすべて削除
            while (parentStack.length > 0 && parentStack[parentStack.length - 1].level >= level) {
                parentStack.pop();
            }
            // 新しい親情報をスタックに追加
            if (level <= 6) {
                parentStack.push({
                    key: cleaned,
                    style: styleName,
                    level: level,
                    counter: (parentStack.length > 0 ? parentStack[parentStack.length - 1].counter : 0) + 1
                });
            }
            // 親情報を取得
            var parentKey = parentStack.length > 0 ? parentStack[parentStack.length - 1].key : "";
            var parentStyleName = parentStack.length > 0 ? parentStack[parentStack.length - 1].style : "";
            var parentLevelCounter = parentStack.length > 0 ? parentStack[parentStack.length - 1].counter : 0;
            var key = styleName + "___" + cleaned + "___" + parentStyleName + "___" + parentKey + "___" + parentLevelCounter;

            if (!comboMap[key]) {
                comboMap[key] = {
                    style: styleName,
                    text: cleaned,
                    count: 1,
                    parent: parentKey,
                    parentStyle: parentStyleName,
                    parentCounter: parentLevelCounter
                };
            } else {
                comboMap[key].count++;
            }
        }
        progressBar.value = i + 1;
        progressWin.update();
    }

    progressWin.close();

    // --- 1回しか出現しないものは除外 ---
    for (var key in comboMap) {
        if (comboMap[key].count === 1) {
            delete comboMap[key];
        }
    }

    // --- comboMap生成後、targets抽出前に key 一覧を alert で表示 ---

    // --- 対象リスト生成 ---

    var targets = [];
    for (var key in comboMap) {
        var entry = comboMap[key];
        if (entry.count >= 2 && entry.parent !== "") {
            targets.push(entry);
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

    // --- UIダイアログ生成 ---
    var dialog = new Window("dialog", LABELS.confirmTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = ["fill", "top"];

    var fullWidthBtn, halfWidthBtn;

    var mainGroup = dialog.add("group");
    mainGroup.orientation = "row";
    mainGroup.alignChildren = ["fill", "top"];

    // 左カラム（スタイル・対象・全角半角）
    var leftGroup = mainGroup.add("group");
    leftGroup.orientation = "column";
    leftGroup.alignChildren = ["fill", "top"];

    // 段落スタイルパネル
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

    // 右カラム（リストボックス）
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

    // チェックボックス連動でリスト有効/無効
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

    // 選択された項目のキーを取得
    function getSelectedKeys() {
        var keys = {};
        for (var i = 0; i < listBox.items.length; i++) {
            if (listBox.items[i].selected) {
                var styleName = targets[i].style;
                var cleaned = targets[i].text;
                var parentStyleName = targets[i].parentStyle || "";
                var parentKey = targets[i].parent || "";
                var parentCounter = targets[i].parentCounter || 0;
                var key = styleName + "___" + cleaned + "___" + parentStyleName + "___" + parentKey + "___" + parentCounter;
                keys[key] = 1;
            }
        }
        return keys;
    }

    // 対象ストーリー取得
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

    // 既存のナンバリングを削除
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

    // --- ボタン ---
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
            // --- 親階層スタックで親情報を管理 ---
            var parentStack = [];
            for (var j = 0; j < paragraphs.length; j++) {
                var para = paragraphs[j];
                if (para.parentTextFrames && para.parentTextFrames.length > 0) {
                    var tf = para.parentTextFrames[0];
                    if (tf.parent instanceof MasterSpread) continue;
                }
                var content = para.contents.replace(/[\r\n]+$/, "");
                var styleName = para.appliedParagraphStyle.name;
                if (styleName == "p.img" || styleName == "p.table") continue;
                var level = headingLevelMap ? (headingLevelMap[styleName] || 99) : 99;
                var regex = /[（\(][0-9０-９]+[）\)]$/;
                var cleaned = content.replace(regex, "");
                // スタックから現在のレベル以上のものをすべて削除
                while (parentStack.length > 0 && parentStack[parentStack.length - 1].level >= level) {
                    parentStack.pop();
                }
                // 新しい親情報をスタックに追加
                if (level <= 6) {
                    parentStack.push({
                        key: cleaned,
                        style: styleName,
                        level: level,
                        counter: (parentStack.length > 0 ? parentStack[parentStack.length - 1].counter : 0) + 1
                    });
                }
                var parentKey = parentStack.length > 0 ? parentStack[parentStack.length - 1].key : "";
                var parentStyleName = parentStack.length > 0 ? parentStack[parentStack.length - 1].style : "";
                var parentLevelCounter = parentStack.length > 0 ? parentStack[parentStack.length - 1].counter : 0;
                var key = styleName + "___" + cleaned + "___" + parentStyleName + "___" + parentKey + "___" + parentLevelCounter;
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

    // --- ナンバリング付与 ---
    function applyNumbering() {
        var localCounts = {};
        for (var key in selectedKeys) {
            localCounts[key] = 1;
        }
        for (var i = 0; i < processStories.length; i++) {
            var paragraphs = processStories[i].paragraphs;
            // --- 親階層スタックで親情報を管理 ---
            var parentStack = [];
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
                var level = headingLevelMap ? (headingLevelMap[styleName] || 99) : 99;
                var regex = /[（\(][0-9０-９]+[）\)]$/;
                var cleaned = content.replace(regex, "");
                // スタックから現在のレベル以上のものをすべて削除
                while (parentStack.length > 0 && parentStack[parentStack.length - 1].level >= level) {
                    parentStack.pop();
                }
                // 新しい親情報をスタックに追加
                if (level <= 6) {
                    parentStack.push({
                        key: cleaned,
                        style: styleName,
                        level: level,
                        counter: (parentStack.length > 0 ? parentStack[parentStack.length - 1].counter : 0) + 1
                    });
                }
                var parentKey = parentStack.length > 0 ? parentStack[parentStack.length - 1].key : "";
                var parentStyleName = parentStack.length > 0 ? parentStack[parentStack.length - 1].style : "";
                var parentLevelCounter = parentStack.length > 0 ? parentStack[parentStack.length - 1].counter : 0;
                var key = styleName + "___" + cleaned + "___" + parentStyleName + "___" + parentKey + "___" + parentLevelCounter;
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