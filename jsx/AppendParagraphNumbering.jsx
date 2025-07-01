#target indesign

/*
    スクリプト名：AppendParagraphNumbering.jsx
    Script Name: AppendParagraphNumbering.jsx

    概要:
    同じテキストが同じ段落スタイルで繰り返される場合、その末尾にナンバリングを追加します。
    UIで全角/半角括弧の選択が可能（日本語UIのみ）。
    検索時には全角・半角数字を認識し、挿入時には半角数字を使用します。

    Overview:
    If the same text repeats with the same paragraph style, add numbering at the end.
    UI allows selection between full-width/half-width parentheses (Japanese UI only).
    Recognizes full-width and half-width numbers during search, uses half-width numbers when inserting.

    動作環境:
    - Adobe InDesign
    - アクティブなドキュメントを対象

    Environment:
    - Adobe InDesign
    - Targets active document

    作成日：2025-06-29
    更新履歴:
    - v1.0.0 (2025-06-29): 初版
    - v1.0.1 (2025-07-01): アンカー付きオブジェクトのみの行を除外
    - v1.0.2 (2025-07-10): 段落スタイル一覧を追加、リストのスタイルを増加
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
    paragraphStyles: {
        ja: "段落スタイル",
        en: "Paragraph Styles"
    },
    numberingStyles: {
        ja: "スタイル",
        en: "Numbering Style"
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

// ドキュメント内の段落のうち、アンカー付きオブジェクトのみで構成される段落を取得する
// Get paragraphs consisting only of anchored objects
function findAnchorOnlyParagraphs(doc) {
    var paragraphs = doc.stories.everyItem().paragraphs.everyItem().getElements();
    var targetParas = [];

    for (var i = 0; i < paragraphs.length; i++) {
        var para = paragraphs[i];
        // 段落内容から改行・空白を除去してアンカー文字だけか確認
        // Remove line breaks and spaces from paragraph content and check if only anchor character
        var cleaned = para.contents.replace(/[\r\n\s]/g, "");
        if (cleaned == "\uFFFC") {
            targetParas.push(para);
        }
    }

    return targetParas.slice();
}

// 配列に要素が含まれているか判定するヘルパー関数
// Helper function to check if array contains an item
function isInArray(array, item) {
    for (var i = 0; i < array.length; i++) {
        if (array[i] === item) {
            return true;
        }
    }
    return false;
}

// 数値をローマ数字（Ⅰ～Ⅹ）に変換（1～10対応）
// Convert number to Roman numeral (I to X) for 1 to 10
function romanNumeral(num) {
    var romans = ["","I","II","III","IV","V","VI","VII","VIII","IX","X"];
    return num <= 10 ? romans[num] : num;
}

// 数値をアルファベット大文字（A～Z）に変換（1～26対応）
// Convert number to uppercase alphabet letter (A to Z) for 1 to 26
function alphaLetter(num) {
    var letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    return num <= 26 ? letters[num - 1] : num;
}

function main() {
    var doc = app.activeDocument;
    var stories = doc.stories;
    var comboMap = {};

    // アンカー付きオブジェクトのみの段落を取得
    // Get paragraphs consisting only of anchored objects
    var anchorParas = findAnchorOnlyParagraphs(doc);

    // 段落数を事前に数える
    // Count total paragraphs in advance
    var paraCount = 0;
    for (var i = 0; i < stories.length; i++) {
        paraCount += stories[i].paragraphs.length;
    }

    // プログレスバー表示
    // Show progress bar
    var progressWin = new Window("palette", "解析中...");
    var progressBar = progressWin.add("progressbar", undefined, 0, paraCount);
    progressBar.preferredSize = [300, 7];
    progressWin.show();

    // 全ストーリーの段落を走査し、段落スタイルとテキスト内容の組み合わせを集計
    // Traverse all story paragraphs and aggregate combinations of paragraph style and text content
    for (var i = 0; i < stories.length; i++) {
        var paragraphs = stories[i].paragraphs;
        for (var j = 0; j < paragraphs.length; j++) {
            var para = paragraphs[j];
            if (isInArray(anchorParas, para)) {
                progressBar.value++;
                if (progressBar.value % 20 == 0) progressWin.update();
                continue; // アンカーのみ段落は除外
                          // Exclude paragraphs with only anchors
            }

            var content = para.contents.replace(/[\r\n]+$/, "");
            // 末尾の番号付き括弧やパターンを除去
            // Remove numbered parentheses or patterns at the end
            var cleaned = content.replace(/[（\(][0-9０-９]+[）\)]$/, "");
            // 空白のみの段落は除外
            // Exclude paragraphs with only spaces
            var trimmed = cleaned.replace(/\s|　/g, "");
            if (trimmed === "") {
                progressBar.value++;
                if (progressBar.value % 20 == 0) progressWin.update();
                continue;
            }
            // アンカーオブジェクト含む段落は除外
            // Exclude paragraphs containing anchored objects
            if (cleaned.indexOf("~a") !== -1) {
                progressBar.value++;
                if (progressBar.value % 20 == 0) progressWin.update();
                continue;
            }
            // 空文字列は除外
            // Exclude empty strings
            if (cleaned === "") {
                progressBar.value++;
                if (progressBar.value % 20 == 0) progressWin.update();
                continue;
            }

            var styleName = para.appliedParagraphStyle.name;
            // Group "part 1", "part 2" etc. as the same for key
            var baseText = cleaned;
            baseText = baseText.replace(/ part [0-9]+$/i, "");
            baseText = baseText.replace(/ part [IVXLCDM]+$/i, "");
            baseText = baseText.replace(/ Pt\. [0-9]+$/i, "");
            baseText = baseText.replace(/\([A-Z]\)$/i, "");
            baseText = baseText.replace(/\([IVXLCDM]+\)$/i, "");
            baseText = baseText.replace(/\([ivxlcdm]+\)$/i, "");
            baseText = baseText.replace(/- [0-9]+$/i, "");
            baseText = baseText.replace(/その[0-9]+$/i, "");
            baseText = baseText.replace(/\([0-9]+\)$/, "");

            var key = styleName + "___" + baseText;
            if (!comboMap[key]) {
                comboMap[key] = {
                    style: styleName,
                    text: cleaned,
                    count: 1
                };
            } else {
                comboMap[key].count++;
            }
            progressBar.value++;
            if (progressBar.value % 20 == 0) progressWin.update();
        }
    }
    progressWin.close();

    // ナンバリング対象となる2回以上出現する組み合わせを抽出
    // Extract combinations appearing two or more times as numbering targets
    var targets = [];
    for (var key in comboMap) {
        if (comboMap[key].count >= 2) {
            targets.push(comboMap[key]);
        }
    }

    // 出現回数の多い順にソート
    // Sort by descending occurrence count
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

    var mainGroup = dialog.add("group");
    mainGroup.orientation = "row";
    mainGroup.alignChildren = ["fill", "top"];

    var leftGroup = mainGroup.add("group");
    leftGroup.orientation = "column";
    leftGroup.alignChildren = ["fill", "top"];
    leftGroup.preferredSize = [200, 400];

    var styleCheckboxGroup = leftGroup.add("panel", undefined, LABELS.paragraphStyles[lang]);
    styleCheckboxGroup.orientation = "column";
    styleCheckboxGroup.alignChildren = ["left", "top"];
    styleCheckboxGroup.preferredSize = [200, 400];
    styleCheckboxGroup.margins = [15, 20, 15, 10];

    var styleSet = {};
    for (var i = 0; i < targets.length; i++) {
        styleSet[targets[i].style] = true;
    }
    var styleNames = [];
    for (var styleName in styleSet) {
        styleNames.push(styleName);
    }
    styleNames.sort();
    var styleCheckboxes = [];
    for (var i = 0; i < styleNames.length; i++) {
        var cb = styleCheckboxGroup.add("checkbox", undefined, styleNames[i]);
        cb.value = true;
        styleCheckboxes.push(cb);
        // チェックボックス変更時に中央リストを更新
        // Update center list when checkbox changes
        cb.onClick = function() {
            updateListBox(targets, styleCheckboxes, listBox);
        };
    }

    var centerGroup = mainGroup.add("group");
    centerGroup.orientation = "column";
    centerGroup.alignChildren = ["fill", "top"];

    var rightGroup = mainGroup.add("group");
    rightGroup.orientation = "column";
    rightGroup.alignChildren = ["fill", "top"];

    var targetGroup = rightGroup.add("panel", undefined, LABELS.target[lang]);
    targetGroup.orientation = "column";
    targetGroup.alignChildren = ["left", "top"];
    var storyRadio = targetGroup.add("radiobutton", undefined, LABELS.story[lang]);
    var documentRadio = targetGroup.add("radiobutton", undefined, LABELS.document[lang]);
    targetGroup.margins = [15, 20, 15, 10];
    storyRadio.value = true;

    var listBox = centerGroup.add("listbox", undefined, "", {
        multiselect: true
    });
    listBox.preferredSize = [400, 500];

    // 初期表示のリストボックスに対象を追加
    // Add targets to listbox on initial display
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

    var numberingStylePanel = rightGroup.add("panel", undefined, LABELS.numberingStyles[lang]);
    var numberingRadioGroup = numberingStylePanel.add("group");
    numberingRadioGroup.orientation = "column";
    numberingRadioGroup.alignChildren = ["left", "top"];
    numberingRadioGroup.margins = [15, 20, 15, 10];

    var col1 = numberingRadioGroup.add("group");
    col1.orientation = "column";
    col1.alignChildren = ["left", "top"];
    var col2 = numberingRadioGroup.add("group");
    col2.orientation = "column";
    col2.alignChildren = ["left", "top"];
    var col3 = numberingRadioGroup.add("group");
    col3.orientation = "column";
    col3.alignChildren = ["left", "top"];
    var col4 = numberingRadioGroup.add("group");
    col4.orientation = "column";
    col4.alignChildren = ["left", "top"];

    var styleParen = col1.add("radiobutton", undefined, "(1)");
    var stylePart1 = col1.add("radiobutton", undefined, "part 1");

    var stylePartI = col2.add("radiobutton", undefined, "part I");
    var stylePt1 = col2.add("radiobutton", undefined, "Pt. 1");

    var styleParenA = col3.add("radiobutton", undefined, "(A)");
    var styleParenRI = col3.add("radiobutton", undefined, "(I)");

    var styleParenRi = col4.add("radiobutton", undefined, "(i)");
    var styleDashNum = col4.add("radiobutton", undefined, "- 1");
    var styleSono = col4.add("radiobutton", undefined, "その1");
    var styleNone = col4.add("radiobutton", undefined, (lang === "ja") ? "なし" : "None");

    styleParen.value = true;

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

    // 選択されたリストアイテムのキーを収集
    // Collect keys of selected list items
    var selectedKeys = {};
    for (var i = 0; i < listBox.items.length; i++) {
        if (listBox.items[i].selected) {
            var key = targets[i].style + "___" + targets[i].text;
            selectedKeys[key] = 1;
        }
    }

    var processStories = [];
    if (storyRadio.value && app.selection.length > 0) {
        // 選択オブジェクトがストーリーとして認識できればそのストーリーのみ処理
        // If selected object can be recognized as a story, process only that story
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

    // ナンバリングを段落末尾に適用する処理
    // Process to apply numbering at the end of paragraphs
    function applyNumbering() {
        var localCounts = {};
        for (var key in selectedKeys) {
            localCounts[key] = 1;
        }

        // 定義: 選択されたナンバリングスタイルに応じて正規表現を決定（厳密なパターン）
        var regex = null;
        if (styleParen.value) {
            regex = /\([0-9]+\)$/;
        } else if (stylePart1.value) {
            regex = / part \d+$/i;
        } else if (stylePartI.value) {
            regex = / part [IVXLCDM]+$/i;
        } else if (stylePt1.value) {
            regex = / Pt\. \d+$/;
        } else if (styleParenA.value) {
            regex = /\([A-Z]\)$/;
        } else if (styleParenRI.value) {
            regex = /\([IVXLCDM]+\)$/;
        } else if (styleParenRi.value) {
            regex = /\([ivxlcdm]+\)$/;
        } else if (styleDashNum.value) {
            regex = /- \d+$/;
        } else if (styleSono.value) {
            regex = /その\d+$/;
        } else if (styleNone.value) {
            // 既存番号削除用: すべての番号パターンを削除
            regex = /(\([0-9]+\)| part \d+| part [IVXLCDM]+| Pt\. \d+|\([A-Z]\)|\([IVXLCDM]+\)|\([ivxlcdm]+\)|- \d+|その\d+)$/i;
        }

        for (var i = 0; i < processStories.length; i++) {
            var paragraphs = processStories[i].paragraphs;
            for (var j = 0; j < paragraphs.length; j++) {
                var para = paragraphs[j];
                if (isInArray(anchorParas, para)) continue; // アンカーのみ段落は除外
                                                          // Exclude paragraphs with only anchors

                var content = para.contents.replace(/[\r\n]+$/, "");
                var styleName = para.appliedParagraphStyle.name;

                // 既存の番号付きパターンを判定する正規表現（選択スタイルのみに限定）
                var cleaned = regex ? content.replace(regex, "") : content;
                var key = styleName + "___" + cleaned;

                if (!(key in localCounts)) continue;

                // 既存の番号付き部分を削除
                // Remove existing numbered part (always run, always remove if matched)
                if (regex && regex.test(content)) {
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

                // ○なしの場合は削除のみで次へ
                if (styleNone.value) {
                    // 番号削除のみ
                    continue;
                }

                // 挿入位置を取得（末尾の改行を除く）
                // Get insertion position (excluding trailing line breaks)
                var insertPos = para.characters.length;
                if (insertPos > 0 && (para.characters[insertPos - 1].contents == "\r" || para.characters[insertPos - 1].contents == "\n")) {
                    insertPos--;
                }

                var numberStr = "";
                if (styleParen.value) {
                    numberStr = "(" + localCounts[key] + ")";
                } else if (stylePart1.value) {
                    numberStr = " part " + localCounts[key];
                } else if (stylePartI.value) {
                    numberStr = " part " + romanNumeral(localCounts[key]);
                } else if (stylePt1.value) {
                    numberStr = " Pt. " + localCounts[key];
                } else if (styleParenA.value) {
                    numberStr = "(" + alphaLetter(localCounts[key]) + ")";
                } else if (styleParenRI.value) {
                    numberStr = "(" + romanNumeral(localCounts[key]) + ")";
                } else if (styleParenRi.value) {
                    numberStr = "(" + romanNumeral(localCounts[key]).toLowerCase() + ")";
                } else if (styleDashNum.value) {
                    numberStr = "- " + localCounts[key];
                } else if (styleSono.value) {
                    numberStr = "その" + localCounts[key];
                }
                para.insertionPoints[insertPos].contents = numberStr;
                localCounts[key]++;
            }
        }
    }

    app.doScript(applyNumbering, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Append Paragraph Numbering");

    // リストボックスを更新する関数
    // Function to update the list box
    function updateListBox(targets, styleCheckboxes, listBox) {
        // Instead of clearing all, hide items matching unchecked styles
        var checkedStyles = [];
        for (var j = 0; j < styleCheckboxes.length; j++) {
            if (styleCheckboxes[j].value) {
                checkedStyles.push(styleCheckboxes[j].text);
            }
        }
        // Hide items with unchecked styles, show only checked
        var visibleCount = 0;
        for (var i = 0; i < listBox.items.length; i++) {
            var item = listBox.items[i];
            // Extract style from label: "Style: ..."
            var label = item.text;
            var styleInLabel = label.split(":")[0];
            if (checkedStyles.indexOf(styleInLabel) === -1) {
                item.visible = false;
            } else {
                item.visible = true;
                visibleCount++;
            }
        }
        // Select first visible item if any
        if (visibleCount > 0) {
            for (var i = 0; i < listBox.items.length; i++) {
                if (listBox.items[i].visible) {
                    listBox.items[i].selected = true;
                    break;
                }
            }
        }
    }
}

main();