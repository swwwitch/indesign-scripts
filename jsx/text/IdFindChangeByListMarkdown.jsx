#target indesign

/*
 * IdFindChangeByListMarkdown.jsx
 *
 * Markdown 記法をまとめて検索・置換し、対応する段落スタイルと文字スタイルを適用します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdFindChangeByListMarkdown";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-03-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdFindChangeByListMarkdown.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdFindChangeByListMarkdown.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n8c0211d92c96"; /* 紹介記事 / article URL */

// Original idea
// Adobe InDesign 付属の FindChangeByList.jsx / Based on the bundled FindChangeByList.jsx

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* ソース形式ごとの既定スタイル名 / Default style names per source format */
var STYLE_PRESETS = {
    html: {
        paragraph: "p", h1: "h1", h2: "h2", h3: "h3", h4: "h4", h5: "h5", h6: "h6",
        bullet: "ul-li", bulletSub: "ul-ul-li", numbered: "ol-li",
        blockquote: "p.blockquote", caption: "p.caption",
        codeBlock: "p.code", codeInline: "code",
        image: "p.img", cell: "td", headerCell: "th",
        bold: "strong-bold", emphasis: "em-i-marker", link: "link"
    },
    msword: {
        paragraph: "Paragraph", h1: "Header1", h2: "Header2", h3: "Header3", h4: "Header4", h5: "Header4", h6: "Header4",
        bullet: "BulList > first", bulletSub: "BulList > BulList > first", numbered: "NumList > first",
        blockquote: "Blockquote > Paragraph", caption: "Caption",
        codeBlock: "CodeBlock", codeInline: "Code",
        image: "Figure", cell: "TablePar", headerCell: "TablePar > TableHeader",
        bold: "Bold", emphasis: "Italic", link: "Link"
    }
};

/* 画像挿入用に一時作成するグラフィックフレームの高さ / Height of the temporary graphic frame used for images */
var TEMP_IMAGE_FRAME_HEIGHT = 50;

// =========================================
// レイアウト設定 / Layout settings
// =========================================

/* 見出しラベルとドロップダウンの幅（px）/ Width of the row label and the style dropdown (px) */
var ROW_LABEL_WIDTH     = 120;
var STYLE_DROPDOWN_WIDTH = 160;

/* スコープ・形式行のラベル幅（px）/ Label width of the scope and format rows (px) */
var SCOPE_LABEL_WIDTH = 80;

/* プログレスバーの幅（px）/ Width of the progress bar (px) */
var PROGRESS_BAR_WIDTH = 300;

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
    dialog: {
        title: { ja: "Markdown → 段落スタイル変換", en: "Markdown to Paragraph Styles" }
    },
    progress: {
        title: { ja: "変換中…", en: "Converting…" },
        label: { ja: "Markdown → 段落スタイル変換", en: "Markdown to Paragraph Styles" }
    },
    scope: {
        label:     { ja: "スコープ", en: "Scope" },
        document:  { ja: "ドキュメント", en: "Document" },
        story:     { ja: "ストーリー", en: "Story" },
        selection: { ja: "選択範囲", en: "Selection" }
    },
    format: {
        label:  { ja: "形式", en: "Format" },
        html:   { ja: "HTML", en: "HTML" },
        msword: { ja: "MS Word", en: "MS Word" }
    },
    panel: {
        headings:  { ja: "基本段落と見出し", en: "Paragraph & Headings" },
        list:      { ja: "リスト", en: "List" },
        options:   { ja: "オプション", en: "Options" },
        code:      { ja: "ソースコード", en: "Source Code" },
        image:     { ja: "画像", en: "Image" },
        table:     { ja: "表組み", en: "Table" },
        charStyle: { ja: "文字スタイル", en: "Character Styles" },
        space:     { ja: "スペース", en: "Space" },
        cleanup:   { ja: "クリーンアップ", en: "Cleanup" }
    },
    row: {
        paragraph:  { ja: "段落", en: "Paragraph" },
        heading1:   { ja: "見出し1", en: "Heading 1" },
        heading2:   { ja: "見出し2", en: "Heading 2" },
        heading3:   { ja: "見出し3", en: "Heading 3" },
        heading4:   { ja: "見出し4", en: "Heading 4" },
        heading5:   { ja: "見出し5", en: "Heading 5" },
        heading6:   { ja: "見出し6", en: "Heading 6" },
        bullet:     { ja: "箇条書き", en: "Bullet List" },
        bulletSub:  { ja: "箇条書き（サブ）", en: "Bullet List (Sub)" },
        numbered:   { ja: "番号リスト", en: "Numbered List" },
        blockquote: { ja: "引用", en: "Blockquote" },
        caption:    { ja: "キャプション", en: "Caption" },
        codeBlock:  { ja: "コードブロック", en: "Code Block" },
        codeInline: { ja: "インライン", en: "Inline" },
        image:      { ja: "画像", en: "Image" },
        cell:       { ja: "セル", en: "Cell" },
        headerCell: { ja: "見出しセル", en: "Header Cell" },
        bold:       { ja: "太字", en: "Bold" },
        emphasis:   { ja: "強調", en: "Emphasis" },
        link:       { ja: "リンク", en: "Link" }
    },
    checkbox: {
        trimLeadingSpace:  { ja: "行頭のスペースを削除", en: "Remove leading spaces" },
        trimTrailingSpace: { ja: "行末のスペースを削除", en: "Remove trailing spaces" },
        convertNbsp:       { ja: "&nbsp; を半角スペースに変換", en: "Convert &nbsp; to space" },
        removeBlankLines:  { ja: "連続する空行を削除", en: "Remove consecutive blank lines" },
        removeComments:    { ja: "コメントアウトを削除", en: "Remove HTML comments" },
        removeRules:       { ja: "水平線を削除", en: "Remove horizontal rules" },
        removeGeta:        { ja: "〓を削除", en: "Remove 〓 marks" }
    },
    button: {
        run:    { ja: "実行", en: "Run" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    alert: {
        noDocument:        { ja: "ドキュメントが開かれていません。", en: "No documents are open." },
        noSelection:       { ja: "ストーリーまたは選択範囲で実行するには、テキストを選択してください。", en: "To run on Story or Selection, select some text first." },
        invalidSelection:  { ja: "選択範囲モードではテキストを選択してください。ストーリーモードではテキストまたはテキストフレームを選択してください。", en: "For Selection mode, select text. For Story mode, select text or a text frame." },
        noTarget:          { ja: "変換対象が選択されていません。", en: "No conversion targets selected." },
        noMatch:           { ja: "該当するMarkdown記法は見つかりませんでした。", en: "No matching Markdown syntax was found." }
    },
    result: {
        separator:   { ja: "：", en: ": " },
        countSuffix: { ja: " 件", en: " items" }
    },
    undo: {
        convertMarkdown: { ja: "Markdown → 段落スタイル変換", en: "Markdown to Paragraph Styles" }
    }
};

/**
 * ドット区切りキーでラベルを取得する
 * @param {string} labelKey 例: "dialog.title"
 * @returns {string} 現在の言語のラベル文字列
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
// UI 補助 / UI helpers
// =========================================

/**
 * ドロップダウンの項目を表示名で選択する
 * @param {DropDownList} dropdown 対象のドロップダウン
 * @param {string} itemName 選択したい項目名
 * @returns {void}
 */
function selectDropdownByName(dropdown, itemName) {
    for (var i = 0; i < dropdown.items.length; i++) {
        if (dropdown.items[i].text === itemName) {
            dropdown.selection = i;
            return;
        }
    }
}

/**
 * スタイル選択ドロップダウンを追加する（推奨候補を先頭にまとめる）
 * @param {Group} parentGroup 追加先のグループ
 * @param {Array<string>} styleNames ドキュメント内のスタイル名一覧
 * @param {string} preferredName 既定で選ぶスタイル名
 * @param {number} dropdownWidth ドロップダウンの幅（px）
 * @param {string} [presetKey] STYLE_PRESETS 内のキー名
 * @returns {DropDownList} 追加したドロップダウン
 */
function addStyleDropdown(parentGroup, styleNames, preferredName, dropdownWidth, presetKey) {
    /* 推奨候補（preferred + 各プリセット）を重複なく集める / Collect preferred and preset candidates without duplicates */
    var candidateNames = [preferredName];
    if (presetKey) {
        for (var presetName in STYLE_PRESETS) {
            if (!STYLE_PRESETS.hasOwnProperty(presetName)) continue;
            var candidate = STYLE_PRESETS[presetName][presetKey];
            if (!candidate) continue;

            var isDuplicate = false;
            for (var i = 0; i < candidateNames.length; i++) {
                if (candidateNames[i] === candidate) {
                    isDuplicate = true;
                    break;
                }
            }
            if (!isDuplicate) candidateNames.push(candidate);
        }
    }

    /* 推奨候補を先頭に、残りのスタイル名を後ろに並べる / Candidates first, remaining document styles after */
    var orderedNames = [];
    for (var j = 0; j < candidateNames.length; j++) {
        orderedNames.push(candidateNames[j]);
    }
    for (var k = 0; k < styleNames.length; k++) {
        var alreadyListed = false;
        for (var m = 0; m < candidateNames.length; m++) {
            if (styleNames[k] === candidateNames[m]) {
                alreadyListed = true;
                break;
            }
        }
        if (!alreadyListed) orderedNames.push(styleNames[k]);
    }

    var styleDropdown = parentGroup.add("dropdownlist", [0, 0, dropdownWidth, 24], orderedNames);
    styleDropdown.selection = 0;
    return styleDropdown;
}

/**
 * ドロップダウンで選択中のスタイル名を取得する
 * @param {DropDownList} dropdown 対象のドロップダウン
 * @returns {string|null} 選択中のスタイル名。未選択時は null
 */
function getSelectedStyleName(dropdown) {
    return dropdown.selection ? dropdown.selection.text : null;
}

/**
 * チェックボックス＋ラベル＋スタイル選択の 1 行を追加する
 * @param {Panel} parentPanel 追加先のパネル
 * @param {string} rowLabel 行のラベル
 * @param {Array<string>} styleNames スタイル名一覧
 * @param {string} preferredName 既定で選ぶスタイル名
 * @param {string} presetKey STYLE_PRESETS 内のキー名
 * @returns {{checkbox: Checkbox, dropdown: DropDownList}} 追加したコントロール
 */
function addStyleRow(parentPanel, rowLabel, styleNames, preferredName, presetKey) {
    var row = parentPanel.add("group");
    setupRow(row, "left", 6);

    var rowCheckbox = row.add("checkbox", undefined, "");
    rowCheckbox.value = true;
    row.add("statictext", [0, 0, ROW_LABEL_WIDTH, 20], rowLabel);
    var rowDropdown = addStyleDropdown(row, styleNames, preferredName, STYLE_DROPDOWN_WIDTH, presetKey);

    return { checkbox: rowCheckbox, dropdown: rowDropdown };
}

// =========================================
// 検索対象の判定 / Target resolution
// =========================================

/**
 * 検索・置換の対象として使えるかを判定する
 * @param {object} selectionItem 選択オブジェクト
 * @returns {boolean} changeGrep / changeText を持っていれば true
 */
function canUseAsTextTarget(selectionItem) {
    if (!selectionItem) return false;
    try {
        if (typeof selectionItem.changeGrep === "function" && typeof selectionItem.changeText === "function") return true;
    } catch (e) {}
    return false;
}

/**
 * 選択オブジェクトから対象ストーリーを取り出す
 * @param {object} selectionItem 選択オブジェクト
 * @returns {Story|null} 対象ストーリー。取得できない場合は null
 */
function resolveStoryTargetFromSelection(selectionItem) {
    if (!selectionItem) return null;
    try {
        if (selectionItem.hasOwnProperty("parentStory") && selectionItem.parentStory != null) {
            return selectionItem.parentStory;
        }
    } catch (e) {}
    try {
        if (selectionItem.hasOwnProperty("texts") && selectionItem.texts.length > 0 &&
            selectionItem.texts[0].parentStory != null) {
            return selectionItem.texts[0].parentStory;
        }
    } catch (e) {}
    return null;
}

// =========================================
// 検索・置換 / Find and change
// =========================================

/**
 * 正規表現で検索・置換し、スタイルを適用する
 * @param {object} target 検索対象
 * @param {string} findWhat 検索する正規表現
 * @param {string} changeTo 置換後の文字列
 * @param {string|null} paragraphStyleName 適用する段落スタイル名
 * @param {string|null} characterStyleName 適用する文字スタイル名
 * @param {string|null} findParagraphStyleName 検索対象を絞り込む段落スタイル名
 * @returns {number} 置換した件数
 */
function findChangeGrep(target, findWhat, changeTo, paragraphStyleName, characterStyleName, findParagraphStyleName) {
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    app.findGrepPreferences.findWhat = findWhat;
    app.changeGrepPreferences.changeTo = changeTo;

    if (findParagraphStyleName != null) {
        try {
            app.findGrepPreferences.appliedParagraphStyle = findParagraphStyleName;
        } catch (e) {
            /* スタイルが見つからない場合は絞り込まない / Do not narrow the search when the style is missing */
        }
    }

    if (paragraphStyleName != null) {
        try {
            app.changeGrepPreferences.appliedParagraphStyle = paragraphStyleName;
        } catch (e) {
            app.findGrepPreferences = NothingEnum.nothing;
            app.changeGrepPreferences = NothingEnum.nothing;
            return 0;
        }
    }

    if (characterStyleName != null) {
        try {
            if (characterStyleName === "[None]" || characterStyleName === "[なし]") {
                app.changeGrepPreferences.appliedCharacterStyle = app.activeDocument.characterStyles[0];
            } else {
                app.changeGrepPreferences.appliedCharacterStyle = characterStyleName;
            }
        } catch (e) {
            app.findGrepPreferences = NothingEnum.nothing;
            app.changeGrepPreferences = NothingEnum.nothing;
            return 0;
        }
    }

    var changedRanges = target.changeGrep();

    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    return changedRanges.length;
}

/**
 * 文字列検索で置換する
 * @param {object} target 検索対象
 * @param {string} findWhat 検索する文字列
 * @param {string} changeTo 置換後の文字列
 * @returns {number} 置換した件数
 */
function findChangeText(target, findWhat, changeTo) {
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    app.findTextPreferences.findWhat = findWhat;
    app.changeTextPreferences.changeTo = changeTo;

    var changedRanges = target.changeText();

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    return changedRanges.length;
}

// =========================================
// ダイアログと実行 / Dialog and run
// =========================================

/**
 * 変換設定ダイアログを表示し、選択内容に従って変換を実行する
 * @returns {void}
 */
function showDialog() {
    var activeDoc = app.documents.item(0);

    /* ドキュメントのスタイル名一覧 / Style names in the document */
    var paragraphStyleNames = [];
    var allParagraphStyles = activeDoc.allParagraphStyles;
    for (var i = 0; i < allParagraphStyles.length; i++) {
        paragraphStyleNames.push(allParagraphStyles[i].name);
    }
    var characterStyleNames = [];
    var allCharacterStyles = activeDoc.allCharacterStyles;
    for (var j = 0; j < allCharacterStyles.length; j++) {
        characterStyleNames.push(allCharacterStyles[j].name);
    }

    var markdownDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(markdownDialog);

    /* スコープ・形式（中央寄せ、内部は左揃え）/ Scope and format (centered block, left-aligned inside) */
    var scopeAndFormatColumn = markdownDialog.add("group");
    scopeAndFormatColumn.orientation = "column";
    scopeAndFormatColumn.alignment = ["center", "top"];
    scopeAndFormatColumn.alignChildren = ["left", "top"];
    scopeAndFormatColumn.spacing = 6;

    var scopeRow = scopeAndFormatColumn.add("group");
    setupRow(scopeRow, "left", 6);
    scopeRow.add("statictext", [0, 0, SCOPE_LABEL_WIDTH, 20], getLabel("scope.label") + getLabel("result.separator"));
    var scopeDocumentRadio  = scopeRow.add("radiobutton", undefined, getLabel("scope.document"));
    var scopeStoryRadio     = scopeRow.add("radiobutton", undefined, getLabel("scope.story"));
    var scopeSelectionRadio = scopeRow.add("radiobutton", undefined, getLabel("scope.selection"));
    scopeStoryRadio.value = true;

    var formatRow = scopeAndFormatColumn.add("group");
    setupRow(formatRow, "left", 6);
    formatRow.add("statictext", [0, 0, SCOPE_LABEL_WIDTH, 20], getLabel("format.label") + getLabel("result.separator"));
    var formatHtmlRadio   = formatRow.add("radiobutton", undefined, getLabel("format.html"));
    var formatMsWordRadio = formatRow.add("radiobutton", undefined, getLabel("format.msword"));
    formatHtmlRadio.value = true;

    /* 2カラム / Two columns */
    var columnsRow = markdownDialog.add("group");
    setupRow(columnsRow, "fill", COLUMN_SPACING);
    columnsRow.alignChildren = ["fill", "top"];

    var leftColumn = columnsRow.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = ["fill", "top"];
    leftColumn.spacing = PANEL_SPACING;

    var rightColumn = columnsRow.add("group");
    rightColumn.orientation = "column";
    rightColumn.alignChildren = ["fill", "top"];
    rightColumn.spacing = PANEL_SPACING;

    // ---- 左列 / Left column ----

    var headingPanel = leftColumn.add("panel", undefined, getLabel("panel.headings"));
    setupPanel(headingPanel, 4);

    var paragraphRow = addStyleRow(headingPanel, getLabel("row.paragraph"), paragraphStyleNames, "p", "paragraph");

    var headingDefinitions = [
        { label: getLabel("row.heading1"), markdown: "#",      styleName: "h1", presetKey: "h1" },
        { label: getLabel("row.heading2"), markdown: "##",     styleName: "h2", presetKey: "h2" },
        { label: getLabel("row.heading3"), markdown: "###",    styleName: "h3", presetKey: "h3" },
        { label: getLabel("row.heading4"), markdown: "####",   styleName: "h4", presetKey: "h4" },
        { label: getLabel("row.heading5"), markdown: "#####",  styleName: "h5", presetKey: "h5" },
        { label: getLabel("row.heading6"), markdown: "######", styleName: "h6", presetKey: "h6" }
    ];

    var headingRows = [];
    for (var k = 0; k < headingDefinitions.length; k++) {
        headingRows.push(addStyleRow(headingPanel, headingDefinitions[k].label, paragraphStyleNames,
            headingDefinitions[k].styleName, headingDefinitions[k].presetKey));
    }

    var listPanel = leftColumn.add("panel", undefined, getLabel("panel.list"));
    setupPanel(listPanel, 4);
    var bulletRow    = addStyleRow(listPanel, getLabel("row.bullet"), paragraphStyleNames, "ul-li", "bullet");
    var bulletSubRow = addStyleRow(listPanel, getLabel("row.bulletSub"), paragraphStyleNames, "ul-ul-li", "bulletSub");
    var numberedRow  = addStyleRow(listPanel, getLabel("row.numbered"), paragraphStyleNames, "ol-li", "numbered");

    var optionPanel = leftColumn.add("panel", undefined, getLabel("panel.options"));
    setupPanel(optionPanel, 4);
    var blockquoteRow = addStyleRow(optionPanel, getLabel("row.blockquote"), paragraphStyleNames, "p.blockquote", "blockquote");
    var captionRow    = addStyleRow(optionPanel, getLabel("row.caption"), paragraphStyleNames, "p.caption", "caption");

    var codePanel = leftColumn.add("panel", undefined, getLabel("panel.code"));
    setupPanel(codePanel, 4);
    var codeBlockRow  = addStyleRow(codePanel, getLabel("row.codeBlock"), paragraphStyleNames, "p.code", "codeBlock");
    var codeInlineRow = addStyleRow(codePanel, getLabel("row.codeInline"), characterStyleNames, "code", "codeInline");

    // ---- 右列 / Right column ----

    var imagePanel = rightColumn.add("panel", undefined, getLabel("panel.image"));
    setupPanel(imagePanel, 4);
    var imageRow = addStyleRow(imagePanel, getLabel("row.image"), paragraphStyleNames, "p.img", "image");

    var tablePanel = rightColumn.add("panel", undefined, getLabel("panel.table"));
    setupPanel(tablePanel, 4);
    var cellRow       = addStyleRow(tablePanel, getLabel("row.cell"), paragraphStyleNames, "td", "cell");
    var headerCellRow = addStyleRow(tablePanel, getLabel("row.headerCell"), paragraphStyleNames, "th", "headerCell");

    var charStylePanel = rightColumn.add("panel", undefined, getLabel("panel.charStyle"));
    setupPanel(charStylePanel, 4);

    var charStyleDefinitions = [
        { label: getLabel("row.bold"),     find: "\\*\\*([^*\\r\\n]+)\\*\\*", styleName: "strong-bold",  presetKey: "bold" },
        { label: getLabel("row.emphasis"), find: "\\*([^*\\r\\n]+)\\*",       styleName: "em-i-marker",  presetKey: "emphasis" },
        { label: getLabel("row.link"),     find: "\\[([^\\]\\r\\n]+)\\]\\(([^)\\r\\n]+)\\)", styleName: "link", changeTo: "$1 <$2>", presetKey: "link" }
    ];

    var charStyleRows = [];
    for (var n = 0; n < charStyleDefinitions.length; n++) {
        charStyleRows.push(addStyleRow(charStylePanel, charStyleDefinitions[n].label, characterStyleNames,
            charStyleDefinitions[n].styleName, charStyleDefinitions[n].presetKey));
    }

    var spacePanel = rightColumn.add("panel", undefined, getLabel("panel.space"));
    setupPanel(spacePanel, 4);
    spacePanel.alignChildren = ["left", "top"];
    var trimLeadingSpaceCheckbox = spacePanel.add("checkbox", undefined, getLabel("checkbox.trimLeadingSpace"));
    trimLeadingSpaceCheckbox.value = true;
    var trimTrailingSpaceCheckbox = spacePanel.add("checkbox", undefined, getLabel("checkbox.trimTrailingSpace"));
    trimTrailingSpaceCheckbox.value = true;
    var convertNbspCheckbox = spacePanel.add("checkbox", undefined, getLabel("checkbox.convertNbsp"));
    convertNbspCheckbox.value = true;

    var cleanupPanel = rightColumn.add("panel", undefined, getLabel("panel.cleanup"));
    setupPanel(cleanupPanel, 4);
    cleanupPanel.alignChildren = ["left", "top"];
    var removeBlankLinesCheckbox = cleanupPanel.add("checkbox", undefined, getLabel("checkbox.removeBlankLines"));
    removeBlankLinesCheckbox.value = true;
    var removeCommentsCheckbox = cleanupPanel.add("checkbox", undefined, getLabel("checkbox.removeComments"));
    removeCommentsCheckbox.value = true;
    var removeRulesCheckbox = cleanupPanel.add("checkbox", undefined, getLabel("checkbox.removeRules"));
    removeRulesCheckbox.value = true;
    var removeGetaCheckbox = cleanupPanel.add("checkbox", undefined, getLabel("checkbox.removeGeta"));
    removeGetaCheckbox.value = false;

    /* ソース形式プリセットの切り替え / Source-format preset switching */
    var styleDropdownByPresetKey = {
        paragraph: paragraphRow.dropdown,
        h1: headingRows[0].dropdown, h2: headingRows[1].dropdown, h3: headingRows[2].dropdown,
        h4: headingRows[3].dropdown, h5: headingRows[4].dropdown, h6: headingRows[5].dropdown,
        bullet: bulletRow.dropdown, bulletSub: bulletSubRow.dropdown, numbered: numberedRow.dropdown,
        blockquote: blockquoteRow.dropdown, caption: captionRow.dropdown,
        codeBlock: codeBlockRow.dropdown, codeInline: codeInlineRow.dropdown,
        image: imageRow.dropdown, cell: cellRow.dropdown, headerCell: headerCellRow.dropdown,
        bold: charStyleRows[0].dropdown, emphasis: charStyleRows[1].dropdown, link: charStyleRows[2].dropdown
    };

    /**
     * 指定したプリセットのスタイル名を各ドロップダウンへ反映する
     * @param {string} presetName STYLE_PRESETS のキー名
     * @returns {void}
     */
    function applyPreset(presetName) {
        var preset = STYLE_PRESETS[presetName];
        if (!preset) return;
        for (var presetKey in preset) {
            if (preset.hasOwnProperty(presetKey) && styleDropdownByPresetKey[presetKey]) {
                selectDropdownByName(styleDropdownByPresetKey[presetKey], preset[presetKey]);
            }
        }
    }

    formatHtmlRadio.onClick   = function () { applyPreset("html"); };
    formatMsWordRadio.onClick = function () { applyPreset("msword"); };

    /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
    var dialogButtonRow = markdownDialog.add("group");
    setupRow(dialogButtonRow, "right", 8);
    dialogButtonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    dialogButtonRow.add("button", undefined, getLabel("button.run"), { name: "ok" });

    if (markdownDialog.show() !== 1) return;

    // ---------------------------------------
    // 検索対象の決定 / Resolve the target
    // ---------------------------------------
    var target;
    if (scopeSelectionRadio.value) {
        if (app.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }
        target = app.selection[0];
        if (!canUseAsTextTarget(target)) {
            alert(getLabel("alert.invalidSelection"));
            return;
        }
    } else if (scopeStoryRadio.value) {
        if (app.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }
        target = resolveStoryTargetFromSelection(app.selection[0]);
        if (target == null) {
            alert(getLabel("alert.invalidSelection"));
            return;
        }
    } else {
        target = app.documents.item(0);
    }

    // ---------------------------------------
    // 変換リストの作成 / Build the operation list
    // ---------------------------------------
    var operations = [];
    var paragraphStyleName = paragraphRow.checkbox.value ? getSelectedStyleName(paragraphRow.dropdown) : null;

    /* 0. 全体に段落スタイル p、文字スタイルなしを適用 / Reset everything to the base paragraph style with no character style */
    operations.push({ find: ".+", changeTo: "$0", style: "p", charStyle: "[None]" });

    /* 0b. 画像挿入用のグラフィックフレームをクリップボードへ / Copy a graphic frame to the clipboard for image insertion */
    if (imageRow.checkbox.value) {
        operations.push({ type: "createFrame" });
    }

    /* 1. フェンスドコードブロックを先に処理して中身を保護 / Handle fenced code blocks first to protect their contents */
    if (codeBlockRow.checkbox.value) {
        operations.push({
            find: "^```[^\\r]*\\r([\\s\\S]+?)\\r```",
            changeTo: "$1",
            style: getSelectedStyleName(codeBlockRow.dropdown)
        });
    }

    /* 1b. 空行とスペース＋改行の前処理 / Pre-pass for blank lines and space-before-return */
    if (removeBlankLinesCheckbox.value) {
        operations.push({ find: "\\r\\r+", changeTo: "\\r", style: null });
        operations.push({ find: " \\r", changeTo: "\\r", style: null });
    }

    /* 2. 番号リスト（見出しより先に処理して「### 1.」の誤マッチを防ぐ）/ Numbered lists before headings, to avoid matching "### 1." */
    if (numberedRow.checkbox.value) {
        operations.push({ find: "^\\d+\\.\\s(.+)", changeTo: "$1", style: getSelectedStyleName(numberedRow.dropdown) });
    }

    /* 3. 見出し（###### → # の順）/ Headings, from the deepest level up */
    for (var h = headingDefinitions.length - 1; h >= 0; h--) {
        if (headingRows[h].checkbox.value) {
            operations.push({
                find: "^" + headingDefinitions[h].markdown + " (.+)",
                changeTo: "$1",
                style: getSelectedStyleName(headingRows[h].dropdown)
            });
        }
    }

    /* 4. 箇条書き（リンク付き）/ Bullet items that contain a link */
    if (bulletRow.checkbox.value) {
        operations.push({
            find: "- \\[([^\\]\\r\\n]+)\\]\\(([^)\\r\\n]+)\\)",
            changeTo: "$1\r$2",
            style: getSelectedStyleName(bulletRow.dropdown)
        });
    }

    /* 5. 画像 / Images */
    if (imageRow.checkbox.value) {
        var imageStyleName = getSelectedStyleName(imageRow.dropdown);
        operations.push({ find: "^\\s*<img [^>]*src=\"(.+?)\".+>", changeTo: "~C\\r$1", style: imageStyleName });
        operations.push({ find: "^!\\[\\]\\((.+)\\)", changeTo: "~C\\r$1", style: imageStyleName });
    }

    /* 6. 箇条書き / Bullet items */
    if (bulletRow.checkbox.value) {
        var bulletStyleName = getSelectedStyleName(bulletRow.dropdown);
        operations.push({ find: "^- (.+)", changeTo: "$1", style: bulletStyleName });
        operations.push({ find: "^\\* (.+)", changeTo: "$1", style: bulletStyleName });
        operations.push({ find: "^\\+ (.+)", changeTo: "$1", style: bulletStyleName });
        operations.push({ find: "^・(.+)", changeTo: "$1", style: bulletStyleName });
    }
    if (bulletSubRow.checkbox.value) {
        operations.push({ find: "^\\t- (.+)", changeTo: "$1", style: getSelectedStyleName(bulletSubRow.dropdown) });
    }

    /* 7. 引用 / Blockquotes */
    if (blockquoteRow.checkbox.value) {
        operations.push({ find: "^>\\s?(.+)", changeTo: "$1", style: getSelectedStyleName(blockquoteRow.dropdown) });
    }

    /* 8. 行頭スペース削除 / Remove leading spaces */
    if (trimLeadingSpaceCheckbox.value) {
        operations.push({ find: "^[\\s|　]", changeTo: "", style: null });
        operations.push({ find: "\\s*\\t", changeTo: "\\t", style: null });
    }

    /* 9. 文字スタイル（太字 → 強調 → リンクの順）/ Character styles: bold, then emphasis, then link */
    for (var c = 0; c < charStyleDefinitions.length; c++) {
        if (charStyleRows[c].checkbox.value) {
            operations.push({
                find: charStyleDefinitions[c].find,
                changeTo: charStyleDefinitions[c].changeTo || "$1",
                charStyle: getSelectedStyleName(charStyleRows[c].dropdown),
                style: null
            });
        }
    }

    /* 9b. 〓 で囲まれたリンク / Links wrapped in 〓 markers */
    if (charStyleRows.length > 2 && charStyleRows[2].checkbox.value) {
        var linkStyleName = getSelectedStyleName(charStyleRows[2].dropdown);
        if (linkStyleName) {
            operations.push({ find: "〓(.+?)〓", changeTo: "<$1>", charStyle: linkStyleName, style: null });
        }
    }

    /* 10. コードブロック / Code blocks */
    if (codeBlockRow.checkbox.value) {
        var codeBlockStyleName = getSelectedStyleName(codeBlockRow.dropdown);
        operations.push({ find: "^<pre><code>(.+)<\\/code><\\/pre>", changeTo: "$1", style: codeBlockStyleName });
        operations.push({ find: "^<pre>(.+)<\\/pre>", changeTo: "$1", style: codeBlockStyleName });
        operations.push({ find: "^<pre>\\r(.+)\\r<\\/pre>", changeTo: "$1", style: codeBlockStyleName });
        operations.push({ find: "^<pre>\\r", changeTo: "<pre>", style: codeBlockStyleName });
        operations.push({ find: "\\r<\\/pre>(.+)", changeTo: "</pre>", style: codeBlockStyleName });
        operations.push({ find: "^<code>(.+)<\\/code>", changeTo: "$1", style: codeBlockStyleName });
    }

    /* 11. インラインコード / Inline code */
    if (codeInlineRow.checkbox.value) {
        operations.push({
            find: "`([^`\\r\\n]+)`",
            changeTo: "$1",
            charStyle: getSelectedStyleName(codeInlineRow.dropdown),
            style: null
        });
    }

    /* 12. キャプション / Captions */
    if (captionRow.checkbox.value) {
        var captionStyleName = getSelectedStyleName(captionRow.dropdown);
        operations.push({ find: "^[\\[［]?キャプション[\\]］：:]?(.+)", changeTo: "$1", style: captionStyleName });
        operations.push({ find: "^<!-- キャプション -->(.+)<!-- \\/キャプション -->", changeTo: "$1", style: captionStyleName });
        operations.push({ find: "^▲(.+)", changeTo: "$1", style: captionStyleName });
        operations.push({ find: "^caption[：:]\\s?(.+)", changeTo: "$1", style: captionStyleName });
    }

    /* 13. 表組み / Tables */
    var headerCellStyleName = headerCellRow.checkbox.value ? getSelectedStyleName(headerCellRow.dropdown) : null;
    var cellStyleName       = cellRow.checkbox.value ? getSelectedStyleName(cellRow.dropdown) : null;

    /* 13a. 区切り行の直前をヘッダー行として扱う / Treat the line above the separator as the header row */
    if (headerCellRow.checkbox.value) {
        operations.push({ find: "^(.*\\|.*)$\\r?\\n^[-:| ]+$", changeTo: "$1", style: headerCellStyleName });
    }
    /* 13b. 区切り行を削除 / Remove the separator row */
    if (cellRow.checkbox.value || headerCellRow.checkbox.value) {
        operations.push({ find: "^\\|?(\\s*:?-{3,}:?\\s*\\|)+\\s*:?-{3,}:?\\s*\\|?\\r?", changeTo: "", style: null });
    }
    /* 13c. 本文セル行の整形 / Clean up body cell rows */
    if (cellRow.checkbox.value) {
        operations.push({ find: "^\\|\\s*", changeTo: "", style: cellStyleName });
        operations.push({ find: "\\s*\\|\\s*$", changeTo: "", style: cellStyleName });
        operations.push({ find: "\\s*\\|\\s*", changeTo: "\\t", style: cellStyleName });
    }
    /* 13d. ヘッダーセル行の整形（findStyle でヘッダー行のみ対象）/ Clean up header rows, narrowed by findStyle */
    if (headerCellRow.checkbox.value) {
        operations.push({ find: "^\\|\\s*", changeTo: "", style: null, findStyle: headerCellStyleName });
        operations.push({ find: "\\s*\\|\\s*$", changeTo: "", style: null, findStyle: headerCellStyleName });
        operations.push({ find: "\\s*\\|\\s*", changeTo: "\\t", style: null, findStyle: headerCellStyleName });
    }

    /* 14. クリーンアップ（コメント、水平線）/ Cleanup: comments and horizontal rules */
    if (removeCommentsCheckbox.value) {
        operations.push({ find: "<!-- (.+) -->", changeTo: "\\r", style: paragraphStyleName });
    }
    if (removeRulesCheckbox.value) {
        operations.push({ find: "^-+", changeTo: "\\r", style: paragraphStyleName });
    }

    /* 15. 空行削除（後処理）/ Remove blank lines (post-pass) */
    if (removeBlankLinesCheckbox.value) {
        operations.push({ find: "\\r\\r+", changeTo: "\\r", style: null });
    }

    /* 16. スペース / Spaces */
    if (convertNbspCheckbox.value) {
        operations.push({ find: "&nbsp;", changeTo: " ", style: null, type: "text" });
    }
    if (trimTrailingSpaceCheckbox.value) {
        operations.push({ find: "\\s+$", changeTo: " ", style: null });
    }

    /* 17. 〓 の削除 / Remove 〓 markers */
    if (removeGetaCheckbox.value) {
        operations.push({ find: "〓", changeTo: "", style: null });
    }

    /* 18. 行頭以外の画像タグも拾う / Catch image tags that are not at the start of a line */
    if (imageRow.checkbox.value) {
        operations.push({
            find: "<img [^>]*src=\"(.+?)\".+>",
            changeTo: "~C\\r$1",
            style: getSelectedStyleName(imageRow.dropdown)
        });
    }

    if (operations.length === 0) {
        alert(getLabel("alert.noTarget"));
        return;
    }

    // ---------------------------------------
    // 実行 / Run
    // ---------------------------------------
    var progressWindow = new Window("palette", getLabel("progress.title"));
    progressWindow.add("statictext", undefined, getLabel("progress.label"));
    var progressBar = progressWindow.add("progressbar", [0, 0, PROGRESS_BAR_WIDTH, 20], 0, operations.length);
    var progressCountLabel = progressWindow.add("statictext", [0, 0, PROGRESS_BAR_WIDTH, 20], "");
    progressWindow.show();

    var totalChangeCount = 0;

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(function () {
        for (var i = 0; i < operations.length; i++) {
            var operation = operations[i];
            progressBar.value = i + 1;
            progressCountLabel.text = (i + 1) + " / " + operations.length;
            progressWindow.update();

            if (operation.type === "createFrame") {
                /* 画像挿入用のグラフィックフレームを作ってクリップボードへコピー / Build a graphic frame and copy it for image insertion */
                var currentPage = activeDoc.layoutWindows[0].activePage;
                var pageMargins = currentPage.marginPreferences;
                var pageBounds  = currentPage.bounds;
                var marginWidth = (pageBounds[3] - pageBounds[1]) - pageMargins.left - pageMargins.right;

                var temporaryFrame = currentPage.rectangles.add({
                    geometricBounds: [
                        pageMargins.top,
                        pageMargins.left,
                        pageMargins.top + TEMP_IMAGE_FRAME_HEIGHT,
                        pageMargins.left + marginWidth
                    ],
                    contentType: ContentType.GRAPHIC_TYPE
                });
                temporaryFrame.fillColor   = activeDoc.swatches.itemByName("None");
                temporaryFrame.strokeColor = activeDoc.swatches.itemByName("None");
                app.select(temporaryFrame);
                app.copy();
                temporaryFrame.remove();
                continue;
            }

            var changeCount = (operation.type === "text")
                ? findChangeText(target, operation.find, operation.changeTo)
                : findChangeGrep(target, operation.find, operation.changeTo,
                    operation.style, operation.charStyle, operation.findStyle);

            totalChangeCount += changeCount;
        }
    }, ScriptLanguage.javascript, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.convertMarkdown"));

    progressWindow.close();

    if (totalChangeCount === 0) {
        alert(getLabel("alert.noMatch"));
    }
}

// =========================================
// メイン処理 / Main
// =========================================

/**
 * ドキュメントの有無を確認してダイアログを表示する
 * @returns {void}
 */
function main() {
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }
    showDialog();
}

main();

})();
