#target indesign

/*

### 概要

同じテキストが同じ段落スタイルで繰り返すとき、末尾に連番を付けます。直近の親見出しごとに重複を判定し、選択範囲・ストーリー・ドキュメントから適用範囲を選べます。

詳細は README を参照してください。

### Overview

Appends sequential numbers to text that repeats with the same paragraph style. Duplicates are grouped by the nearest parent heading, and the scope can be a selection, a story, or the whole document.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AppendParagraphNumbering";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-13";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/AppendParagraphNumbering.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/AppendParagraphNumbering.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nc96549bb60f9"; /* 紹介記事 / article URL */

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

    /* 見出しとして扱う最大レベルと、見出し以外に与えるレベル / Deepest heading level, and the level for non-headings */
    var MAX_HEADING_LEVEL = 6;
    var NON_HEADING_LEVEL = 99;

    /* ナンバリング対象から除外する段落スタイル / Paragraph styles excluded from numbering */
    var IGNORE_STYLE_NAMES = ["p.img", "p.table"];

    /* 末尾ナンバリングを見つける正規表現 / Pattern that matches trailing numbering */
    var NUMBERING_PATTERN = /[（\(][0-9０-９]+[）\)]$/;

    /* 識別キーを連結する区切り / Separator used to join the key */
    var KEY_SEPARATOR = "___";

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

    /* ボタン行の余白 [左,上,右,下] とスペーサーの最小幅 / Button row margins [l,t,r,b] and the spacer minimum width */
    var BUTTON_ROW_MARGINS = [10, 10, 10, 0];
    var BUTTON_ROW_SPACER_MIN_WIDTH = 0;

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
            selection: { ja: "選択範囲", en: "Selection" },
            story: { ja: "ストーリー", en: "Story" },
            document: { ja: "ドキュメント", en: "Document" },
            fullWidth: { ja: "全角", en: "Full-width" },
            halfWidth: { ja: "半角", en: "Half-width" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            deleteItem: { ja: "削除", en: "Delete" },
            apply: { ja: "追加", en: "Add" }
        },
        progress: {
            title: { ja: "解析中", en: "Analyzing" }
        },
        alert: {
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            noTargets: {
                ja: "ナンバリング対象が見つかりませんでした。",
                en: "No numbering targets were found."
            },
            noSelection: {
                ja: "テキストが選択されていません。ドキュメント全体を対象にします。",
                en: "Nothing is selected. The entire document will be processed."
            },
            noRange: {
                ja: "選択範囲がありません。ストーリー全体を対象にします。",
                en: "No text range is selected. The entire story will be processed."
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
     * everyItem() の戻り値を配列に正規化する（要素が1件のときスカラーで返るため）
     * @param {*} value everyItem() で取得した値
     * @returns {Array} 正規化した配列
     */
    function toArray(value) {
        return (value instanceof Array) ? value : [value];
    }

    /**
     * ストーリーが親（マスター）ページ上にあるかを判定する
     * @param {Story} story 対象のストーリー
     * @returns {boolean} 親ページ上なら true
     */
    function isMasterStory(story) {
        var containers = story.textContainers;
        if (!containers || containers.length === 0) return false;
        /* グループ内のフレームもあるのでスプレッドに達するまで親を遡る / Walk up until a spread is reached */
        var node = containers[0];
        while (node) {
            var typeName = node.constructor.name;
            if (typeName === "MasterSpread") return true;
            if (typeName === "Spread" || typeName === "Document" || typeName === "Application") return false;
            node = node.parent;
        }
        return false;
    }

    /**
     * 走査対象から外す段落かを判定する（解析・付与・削除で共通）
     * @param {string} cleanedText ナンバリングを除いた本文
     * @param {string} styleName 段落スタイル名
     * @returns {boolean} 除外するなら true
     */
    function isSkippedParagraph(cleanedText, styleName) {
        /* 空行・1文字以下・空白のみはスキップ / Skip empty, single-char, or whitespace-only */
        if (cleanedText.length <= 1) return true;
        if (/^\s+$/.test(cleanedText)) return true;
        for (var i = 0; i < IGNORE_STYLE_NAMES.length; i++) {
            if (IGNORE_STYLE_NAMES[i] === styleName) return true;
        }
        return false;
    }

    /**
     * 末尾の改行を取り除く
     * @param {string} text 対象の文字列
     * @returns {string} 末尾の改行を除いた文字列
     */
    function trimTrailingBreaks(text) {
        return String(text).replace(/[\r\n]+$/, "");
    }

    /**
     * 段落本文からナンバリングを除いた比較用テキストを作る
     * @param {string} contents 段落の内容
     * @returns {string} 比較用テキスト
     */
    function toCleanedText(contents) {
        return trimTrailingBreaks(contents).replace(NUMBERING_PATTERN, "");
    }

    /**
     * 段落末尾の改行を除いた最後の文字位置を返す
     * @param {Paragraph} paragraph 対象の段落
     * @returns {number} 文字インデックス。該当がなければ -1
     */
    function getLastVisibleIndex(paragraph) {
        var index = paragraph.characters.length - 1;
        while (index >= 0) {
            var character = paragraph.characters[index].contents;
            if (character !== "\r" && character !== "\n") break;
            index--;
        }
        return index;
    }

    /**
     * 進捗バーを表示しながら処理を実行する
     * @param {number} maxValue 進捗の最大値
     * @param {function} task 進捗更新関数を受け取る処理
     * @returns {void}
     */
    function withProgressBar(maxValue, task) {
        var progressWindow = new Window("palette", getLabel("progress.title"));
        progressWindow.orientation = "column";
        progressWindow.alignChildren = ["fill", "top"];
        progressWindow.margins = WINDOW_MARGINS;
        var progressBar = progressWindow.add("progressbar", undefined, 0, maxValue);
        progressBar.preferredSize = PROGRESS_BAR_SIZE;
        progressWindow.show();

        /* 解析が失敗してもパレットを残さない / Never leave the palette behind when the scan throws */
        try {
            task(function (value) {
                progressBar.value = value;
                progressWindow.update();
            });
        } finally {
            progressWindow.close();
        }
    }

    // =========================================
    // 段落の走査 / Paragraph scanning
    // =========================================

    /**
     * 見出しスタックを更新しながら、段落の識別キーを作る
     * @param {string} cleanedText ナンバリングを除いた本文
     * @param {string} styleName 段落スタイル名
     * @param {Array<object>} headingStack 見出しの階層スタック
     * @returns {object} 識別キーと関連情報
     */
    function buildKeyForParagraph(cleanedText, styleName, headingStack) {
        var headingLevel = HEADING_LEVEL_MAP[styleName];
        if (typeof headingLevel !== "number") headingLevel = NON_HEADING_LEVEL;

        /* 現在のレベル以上の親をスタックから除去 / Pop parents at the same or deeper level */
        while (headingStack.length > 0 && headingStack[headingStack.length - 1].level >= headingLevel) {
            headingStack.pop();
        }
        /* 見出しなら親としてスタックに積む / Push headings onto the heading stack */
        if (headingLevel <= MAX_HEADING_LEVEL) {
            headingStack.push({ text: cleanedText, style: styleName, level: headingLevel });
        }

        /* 親は直近の見出しだけを見る。祖先まで含めると章ごとに分かれ、繰り返しと見なされなくなる
           / Use only the nearest heading; including ancestors splits repeats per chapter */
        var parent = (headingStack.length > 0) ? headingStack[headingStack.length - 1] : null;
        var parentLabel = parent ? (parent.style + ":" + parent.text) : "";

        return {
            key: styleName + KEY_SEPARATOR + cleanedText + KEY_SEPARATOR + parentLabel,
            style: styleName,
            text: cleanedText,
            parentLabel: parentLabel
        };
    }

    /**
     * ストーリーを走査して、対象になる段落の情報を集める
     * @param {Story} story 対象のストーリー
     * @returns {Array<object>} 段落番号と識別キーを持つ情報の一覧
     */
    function scanStory(story) {
        var scanned = [];
        var paragraphCount = story.paragraphs.length;
        if (isMasterStory(story) || paragraphCount === 0) return scanned;

        /* 内容とスタイルを一括取得して段落ごとの DOM アクセスを減らす / Bulk-read to cut per-paragraph DOM access */
        var contentsList = toArray(story.paragraphs.everyItem().contents);
        var styleList = toArray(story.paragraphs.everyItem().appliedParagraphStyle);

        /* 段落数と合わなければ取りこぼすので個別取得に切り替える / Fall back per paragraph when the bulk read does not line up */
        if (contentsList.length !== paragraphCount || styleList.length !== paragraphCount) {
            contentsList = [];
            styleList = [];
            for (var n = 0; n < paragraphCount; n++) {
                contentsList.push(story.paragraphs[n].contents);
                styleList.push(story.paragraphs[n].appliedParagraphStyle);
            }
        }

        var headingStack = [];
        for (var i = 0; i < contentsList.length; i++) {
            var cleanedText = toCleanedText(contentsList[i]);
            var styleName = styleList[i].name;
            if (isSkippedParagraph(cleanedText, styleName)) continue;

            var info = buildKeyForParagraph(cleanedText, styleName, headingStack);
            info.index = i;
            scanned.push(info);
        }
        return scanned;
    }

    /**
     * 対象ストーリーを走査し、識別キーが一致した段落を処理する
     * @param {Array<Story>} stories 対象のストーリー
     * @param {object} keyMap 対象の識別キーを持つマップ
     * @param {function} handler 一致した段落に対する処理（引数: 段落, 識別キー）
     * @returns {void}
     */
    function eachMatchedParagraph(scopes, keyMap, handler) {
        for (var i = 0; i < scopes.length; i++) {
            var scope = scopes[i];
            /* 範囲を絞る場合も走査はストーリー全体で行う。手前の見出しを見ないと親が変わってしまう
               / Always scan the whole story: skipping earlier headings would change the parent */
            var scanned = scanStory(scope.story);
            for (var j = 0; j < scanned.length; j++) {
                if (!(scanned[j].key in keyMap)) continue;
                /* 付与も削除も段落数を変えないので、走査時の段落番号をそのまま使える
                   / Neither handler changes the paragraph count, so scanned indexes stay valid */
                var paragraph = scope.story.paragraphs[scanned[j].index];
                if (!isInScope(paragraph, scope)) continue;
                handler(paragraph, scanned[j].key);
            }
        }
    }

    // =========================================
    // 解析 / Analysis
    // =========================================

    /**
     * 全ストーリーを解析し、ナンバリング対象の候補を求める
     * @param {Stories} allStories 対象のストーリー
     * @returns {Array<object>} 出現回数の多い順に並べた候補
     */
    function findNumberingTargets(allStories) {
        var occurrenceMap = {};

        withProgressBar(allStories.length, function (setProgress) {
            for (var i = 0; i < allStories.length; i++) {
                var scanned = scanStory(allStories[i]);
                for (var j = 0; j < scanned.length; j++) {
                    var entry = occurrenceMap[scanned[j].key];
                    if (entry) {
                        entry.count++;
                    } else {
                        scanned[j].count = 1;
                        occurrenceMap[scanned[j].key] = scanned[j];
                    }
                }
                setProgress(i + 1);
            }
        });

        /* 2回以上出現するものを対象に。見出しスタイルは HEADING_LEVEL_MAP にある名前だけを親として
           扱うので、親の有無は条件にしない
           / Keep whatever repeats; only names in HEADING_LEVEL_MAP count as parents, so a parent is not required */
        var numberingTargets = [];
        for (var mapKey in occurrenceMap) {
            var mapEntry = occurrenceMap[mapKey];
            if (mapEntry.count >= 2) {
                numberingTargets.push(mapEntry);
            }
        }

        /* 出現回数の多い順、同数ならスタイル名順 / Sort by count desc, then style name */
        numberingTargets.sort(function (a, b) {
            if (b.count !== a.count) return b.count - a.count;
            return a.style.toLowerCase() < b.style.toLowerCase() ? -1 : 1;
        });

        return numberingTargets;
    }

    /* 対象範囲の指定 / Scope of the target range */
    var SCOPE_SELECTION = "selection";
    var SCOPE_STORY = "story";
    var SCOPE_DOCUMENT = "document";

    /**
     * ストーリーを丸ごと対象にする範囲へ変換する
     * @param {Array<Story>|Stories} stories 対象のストーリー
     * @returns {Array<object>} 範囲を絞らない対象範囲の一覧
     */
    function toWholeStoryScopes(stories) {
        var scopes = [];
        for (var i = 0; i < stories.length; i++) {
            scopes.push({ story: stories[i], fromOffset: null, toOffset: null });
        }
        return scopes;
    }

    /**
     * 段落が対象範囲に含まれるかを判定する
     * @param {Paragraph} paragraph 対象の段落
     * @param {object} scope 対象範囲 { fromOffset, toOffset }
     * @returns {boolean} 含まれるなら true
     */
    function isInScope(paragraph, scope) {
        if (scope.fromOffset === null) return true;
        /* 段落の一部でも選択範囲にかかっていれば対象にする / A partial overlap is enough */
        var paraStart = paragraph.characters[0].index;
        var paraEnd = paragraph.characters[-1].index;
        return paraEnd >= scope.fromOffset && paraStart <= scope.toOffset;
    }

    /**
     * 対象範囲に応じて処理するストーリーと文字位置を求める
     * @param {string} scopeMode SCOPE_SELECTION / SCOPE_STORY / SCOPE_DOCUMENT のいずれか
     * @param {Stories} allStories ドキュメント内の全ストーリー
     * @returns {Array<object>} 対象範囲 { story, fromOffset, toOffset } の一覧
     */
    function resolveTargetScopes(scopeMode, allStories) {
        if (scopeMode === SCOPE_DOCUMENT) return toWholeStoryScopes(allStories);

        if (app.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return toWholeStoryScopes(allStories);
        }
        var selectionItem = app.selection[0];
        var parentStory = null;
        if (selectionItem.hasOwnProperty("parentStory")) {
            parentStory = selectionItem.parentStory;
        } else if (selectionItem.parent && selectionItem.parent.hasOwnProperty("parentStory")) {
            parentStory = selectionItem.parent.parentStory;
        }
        if (!parentStory) {
            alert(getLabel("alert.notStory"));
            return toWholeStoryScopes(allStories);
        }
        if (scopeMode === SCOPE_STORY) return toWholeStoryScopes([parentStory]);

        /* 選択範囲：選択した文字の範囲だけに絞る / Selection: narrow down to the selected characters */
        var selectedChars = selectionItem.hasOwnProperty("characters") ? selectionItem.characters : null;
        if (!selectedChars || selectedChars.length === 0) {
            alert(getLabel("alert.noRange"));
            return toWholeStoryScopes([parentStory]);
        }
        return [{
            story: parentStory,
            fromOffset: selectedChars[0].index,
            toOffset: selectedChars[-1].index
        }];
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 対象リストを作成して候補を並べる
     * @param {Group} parentGroup 追加先のグループ
     * @param {Array<object>} numberingTargets ナンバリング対象の候補
     * @returns {ListBox} 作成したリストボックス
     */
    function buildTargetListBox(parentGroup, numberingTargets) {
        var targetListBox = parentGroup.add("listbox", undefined, "", { multiselect: true });
        targetListBox.preferredSize = TARGET_LIST_SIZE;

        for (var i = 0; i < numberingTargets.length; i++) {
            var target = numberingTargets[i];
            var displayText = (target.text.length > LIST_TEXT_MAX_LENGTH) ? target.text.substring(0, LIST_TEXT_KEEP_LENGTH) + "…" : target.text;
            var countText = (currentLanguage === "ja") ? "（" + target.count + "）" : " (" + target.count + ")";
            var listItem = targetListBox.add("item", target.style + ": " + displayText + countText);
            /* 省略された全文と、親見出しがあればその見出しを添える / Show the full text, plus the parent heading when there is one */
            listItem.helpTip = target.parentLabel ? (target.text + "\n" + target.parentLabel) : target.text;
        }
        if (targetListBox.items.length > 0) {
            targetListBox.items[0].selected = true;
        }
        return targetListBox;
    }

    /**
     * 段落スタイルの絞り込みチェックボックスを作り、対象リストと連動させる
     * @param {Panel} panel 追加先のパネル
     * @param {Array<object>} numberingTargets ナンバリング対象の候補
     * @param {ListBox} targetListBox 連動させる対象リスト
     * @returns {void}
     */
    function buildStyleFilter(panel, numberingTargets, targetListBox) {
        var styleNames = [];
        var seenStyles = {};
        for (var i = 0; i < numberingTargets.length; i++) {
            var styleName = numberingTargets[i].style;
            if (seenStyles[styleName] === true) continue;
            seenStyles[styleName] = true;
            styleNames.push(styleName);
        }
        styleNames.sort();

        var styleCheckboxes = {};

        /**
         * 対象リストの有効／無効を現在のチェック状態に合わせて切り替える
         * @returns {void}
         */
        function updateListBoxEnabled() {
            for (var i = 0; i < targetListBox.items.length; i++) {
                var listItem = targetListBox.items[i];
                listItem.enabled = styleCheckboxes[numberingTargets[i].style].value;
                if (!listItem.enabled) listItem.selected = false;
            }
        }

        for (var i = 0; i < styleNames.length; i++) {
            var styleCheckbox = panel.add("checkbox", undefined, styleNames[i]);
            styleCheckbox.value = true;
            styleCheckbox.onClick = updateListBoxEnabled;
            styleCheckboxes[styleNames[i]] = styleCheckbox;
        }
    }

    /**
     * 対象選択ダイアログを組み立てる
     * @param {Array<object>} numberingTargets ナンバリング対象の候補
     * @param {Stories} allStories ドキュメント内の全ストーリー
     * @returns {object} ウィンドウと選択内容を取り出す関数を持つオブジェクト
     */
    function buildDialog(numberingTargets, allStories) {
        var dialogWindow = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialogWindow, 10);

        var halfWidthBtn = null;

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
        targetPanel.alignChildren = ["left", "top"];
        var selectionRadio = targetPanel.add("radiobutton", undefined, getLabel("radio.selection"));
        var storyRadio = targetPanel.add("radiobutton", undefined, getLabel("radio.story"));
        targetPanel.add("radiobutton", undefined, getLabel("radio.document"));
        storyRadio.value = true;

        /* 全角／半角選択（日本語UIのみ）/ Full/half-width selection (Japanese UI only) */
        if (currentLanguage === "ja") {
            var bracketRadioGroup = optionColumn.add("group");
            setupRow(bracketRadioGroup, "center", 8);
            var fullWidthBtn = bracketRadioGroup.add("radiobutton", undefined, getLabel("radio.fullWidth"));
            halfWidthBtn = bracketRadioGroup.add("radiobutton", undefined, getLabel("radio.halfWidth"));
            fullWidthBtn.value = true;
        }

        /* 右カラム：対象リスト / Right column: target list */
        var listColumn = contentGroup.add("group");
        listColumn.orientation = "column";
        listColumn.alignChildren = ["fill", "top"];
        var targetListBox = buildTargetListBox(listColumn, numberingTargets);
        buildStyleFilter(stylePanel, numberingTargets, targetListBox);

        /**
         * リストで選択中の識別キーを集める
         * @returns {object} 識別キーをキーに持つマップ
         */
        function getSelectedKeys() {
            var selectedKeyMap = {};
            for (var i = 0; i < targetListBox.items.length; i++) {
                if (targetListBox.items[i].selected) {
                    selectedKeyMap[numberingTargets[i].key] = 1;
                }
            }
            return selectedKeyMap;
        }

        /**
         * 現在の対象範囲に応じた処理対象を求める
         * @returns {Array<object>} 対象範囲 { story, fromOffset, toOffset } の一覧
         */
        function getTargetScopes() {
            var scopeMode = SCOPE_DOCUMENT;
            if (selectionRadio.value) scopeMode = SCOPE_SELECTION;
            else if (storyRadio.value) scopeMode = SCOPE_STORY;
            return resolveTargetScopes(scopeMode, allStories);
        }

        /**
         * ナンバリングに使う括弧を求める
         * @returns {object} 左右の括弧 { left, right }
         */
        function getBrackets() {
            if (halfWidthBtn && halfWidthBtn.value) return { left: "(", right: ")" };
            return { left: "（", right: "）" };
        }

        /* ボタン行（横並び）/ Button row (horizontal layout) */
        var btnRowGroup = dialogWindow.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = BUTTON_ROW_MARGINS;
        btnRowGroup.alignment = ["fill", "bottom"];

        /* 左側グループ / Left-side button group */
        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var deleteBtn = btnLeftGroup.add("button", undefined, getLabel("button.deleteItem"));

        /* スペーサー（伸縮）/ Spacer (stretchable) */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = BUTTON_ROW_SPACER_MIN_WIDTH;

        /* 右側グループ / Right-side button group */
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        /* ラベルが "OK" / "Cancel" でないと既定の割り当てが効かないので name を明示
           / Labels other than "OK" / "Cancel" need an explicit name */
        btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel("button.apply"), { name: "ok" });

        /* 選択テキストから既存ナンバリングを削除（undoは1ステップ）/ Remove numbering from selected text (single undo) */
        deleteBtn.onClick = function () {
            var targetScopes = getTargetScopes();
            var selectedKeyMap = getSelectedKeys();
            /* リストが古くなるので、書き換える前にダイアログを閉じる / Close first: the list goes stale once text changes */
            dialogWindow.close(2);

            app.doScript(function () {
                eachMatchedParagraph(targetScopes, selectedKeyMap, removeExistingNumbering);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Remove Paragraph Numbering");
            alert(getLabel("alert.removed"));
        };

        return {
            window: dialogWindow,
            getSelectedKeys: getSelectedKeys,
            getTargetScopes: getTargetScopes,
            getBrackets: getBrackets
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 段落末尾の既存ナンバリングを削除する
     * @param {Paragraph} paragraph 対象の段落
     * @returns {void}
     */
    function removeExistingNumbering(paragraph) {
        var match = trimTrailingBreaks(paragraph.contents).match(NUMBERING_PATTERN);
        if (!match) return;
        var endIndex = getLastVisibleIndex(paragraph);
        var startIndex = endIndex - match[0].length + 1;
        if (startIndex < 0) return;
        paragraph.characters.itemByRange(startIndex, endIndex).remove();
    }

    /**
     * 選択した対象の末尾にナンバリングを付与する
     * @param {Array<Story>} stories 対象のストーリー
     * @param {object} keyMap 対象の識別キーを持つマップ
     * @param {object} brackets 使用する括弧 { left, right }
     * @returns {void}
     */
    function applyNumbering(scopes, keyMap, brackets) {
        var counterByKey = {};
        eachMatchedParagraph(scopes, keyMap, function (paragraph, key) {
            removeExistingNumbering(paragraph);
            var counter = (counterByKey[key] || 0) + 1;
            counterByKey[key] = counter;
            paragraph.insertionPoints[getLastVisibleIndex(paragraph) + 1].contents = brackets.left + counter + brackets.right;
        });
    }

    /**
     * 対象を解析し、ダイアログの指定に従ってナンバリングを付与する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var allStories = app.activeDocument.stories;
        var numberingTargets = findNumberingTargets(allStories);
        if (numberingTargets.length === 0) {
            alert(getLabel("alert.noTargets"));
            return;
        }

        var dialog = buildDialog(numberingTargets, allStories);
        if (dialog.window.show() != 1) return;

        var targetScopes = dialog.getTargetScopes();
        var selectedKeyMap = dialog.getSelectedKeys();
        var brackets = dialog.getBrackets();

        app.doScript(function () {
            applyNumbering(targetScopes, selectedKeyMap, brackets);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Append Paragraph Numbering");
    }

    main();

})();
