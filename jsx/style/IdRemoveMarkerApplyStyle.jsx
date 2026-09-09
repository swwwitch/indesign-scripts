#target indesign

/*

### 概要

行頭の目印（Markdown の見出し・箇条書き・番号リストなど）を手がかりに段落スタイルと文字スタイルを適用し、その目印を削除します。目印は検索対象から自動判別して候補に並べられます。

詳細は README を参照してください。

### Overview

Applies a paragraph style and a character style to paragraphs carrying a leading marker — Markdown headings, bullets, numbered lists and the like — then deletes the marker. Markers are detected from the search target and offered as a list.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdRemoveMarkerApplyStyle";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                        /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)"; /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-09";                  /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-09";                  /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdRemoveMarkerApplyStyle.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdRemoveMarkerApplyStyle.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// 基本設定 / Settings
// =========================================

/* 検索文字列の指定方法 / How the search text is specified */
var SEARCH_MODE_TEXT = 0;
var SEARCH_MODE_AUTO = 1;

/* ダイアログの初期状態 / Dialog defaults */
var DEFAULT_SEARCH_MODE      = SEARCH_MODE_TEXT;
var DEFAULT_SEARCH_TEXT      = "###";
var DEFAULT_MATCH_LEVEL      = true;  /* 見出しレベルに合わせて段落スタイルを選ぶ / match the style to the level */
var DEFAULT_APPLY_ALL_LEVELS = false; /* 見出しを連続適用する / apply every heading level */

/* 自動判別で拾う行頭記号 / Line-head markers picked up by auto detection */
/* 半角記号に加え、箇条書きに使う約物と全角記号を含める / ASCII symbols plus Japanese bullets and full-width symbols */
var LINE_HEAD_MARKER_PATTERN = /^([#*\->+~=|_`●○◎◆◇■□▲△▼▽★☆・※＊＃－＋＞]{1,10})[ \t　]?/;

/* 太字（**文字列**）。行頭記号ではなく囲みなので個別に判定する /
   Bold spans are enclosing, not line-head markers */
var BOLD_PATTERN = /\*\*[^*]+\*\*/;
var BOLD_SEARCH_TEXT = "\\*\\*";

/* 行頭の "**" は太字として拾うので、行頭記号からは除く / "**" at the line head belongs to bold */
var EXCLUDED_LINE_HEAD_PATTERN = /^\*{2,}/;

/* 漢数字（十・百・千を含む） / Kanji numerals */
var KANJI_NUMBER = "[〇一二三四五六七八九十百千]+";

/* 自動判別で拾うナンバリング（数字が変わるので GREP 検索になる） / Numbering (searched with GREP) */
/* pattern は JavaScript の正規表現と InDesign の GREP の両方で通る書き方にそろえる /
   Each pattern is written so it works both as a JavaScript RegExp and as an InDesign GREP query */
var NUMBERED_LIST_PATTERNS = [
    { pattern: "^\\d+\\. ", labelKey: "marker.numberedDot" },
    { pattern: "^\\d+\\) ", labelKey: "marker.numberedParen" },
    { pattern: "^[（(]" + KANJI_NUMBER + "[）)][ 　]?", labelKey: "marker.kanjiParen" },
    { pattern: "^" + KANJI_NUMBER + "\\. ", labelKey: "marker.kanjiDot" },
    { pattern: "^" + KANJI_NUMBER + "、", labelKey: "marker.kanjiComma" },
    { pattern: "^" + KANJI_NUMBER + "[）)][ 　]?", labelKey: "marker.kanjiCloseParen" }
];

/* 判定用の正規表現をあらかじめ用意する / Compile the detection patterns once */
for (var patternIndex = 0; patternIndex < NUMBERED_LIST_PATTERNS.length; patternIndex++) {
    NUMBERED_LIST_PATTERNS[patternIndex].regex =
        new RegExp(NUMBERED_LIST_PATTERNS[patternIndex].pattern);
}

/* 自動判別で「繰り返し」とみなす最小段落数 / Minimum paragraphs to treat a marker as repeated */
var MIN_REPEAT_COUNT = 2;

/* Markdown の見出しマーカー（# 〜 ######） / Markdown heading markers */
var HEADING_MARKER_PATTERN = /^#{1,6} ?$/;
var MAX_HEADING_LEVEL = 6;

/* 見出しマーカーの後ろの区切り（無くてもよい） / Separator after a heading marker (optional) */
var HEADING_SEPARATOR_PATTERN = "[ 　\\t]?";

/* 見出しレベルに対応する段落スタイル名（先に見つかったものを使う） / Paragraph style names per heading level */
var HEADING_STYLE_NAME_FORMATS = ["h%1", "H%1", "heading %1", "Heading %1"];

/* 検索オプション（前回の「検索/置換」の設定を引き継がない） / Find options (never inherited) */
var FIND_CASE_SENSITIVE         = true;  /* 大文字と小文字を区別 / case sensitive */
var FIND_WIDTH_SENSITIVE        = true;  /* 半角と全角を区別（＃ と # を分ける） / width sensitive */
var FIND_KANA_SENSITIVE         = true;  /* ひらがなとカタカナを区別 / kana sensitive */
var FIND_INCLUDE_FOOTNOTES      = true;  /* 脚注を含む / include footnotes */
var FIND_INCLUDE_MASTER_PAGES   = true;  /* マスターページを含む / include master pages */
var FIND_INCLUDE_HIDDEN_LAYERS  = false; /* 非表示レイヤーを含む / include hidden layers */
var FIND_INCLUDE_LOCKED_LAYERS  = false; /* ロックされたレイヤーを含む / include locked layers */
var FIND_INCLUDE_LOCKED_STORIES = false; /* ロックされたストーリーを含む / include locked stories */

/* スコープ（検索対象） / Search scope */
var SCOPE_ALL_DOCUMENTS = 0;
var SCOPE_DOCUMENT      = 1;
var SCOPE_STORY         = 2;

/* ダイアログの初期スコープ / Default scope */
var DEFAULT_SCOPE = SCOPE_STORY;

// =========================================
// UIレイアウトの共通設定 / Shared UI layout
// =========================================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;               /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;               /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 8;                /* パネル内の要素間隔 / panel spacing */

/* 行の寸法 / Row metrics */
var LABEL_WIDTH        = 110; /* ラベル・ラジオボタンの幅 / label & radio button width */
var SEARCH_TEXT_LENGTH = 16;  /* 検索文字列入力欄の文字数 / search field length */
var BUTTON_WIDTH       = 90;  /* ダイアログボタンの幅 / dialog button width */

/**
 * ウィンドウの共通設定を適用する
 * @param {Window} win 対象ウィンドウ
 * @returns {void}
 */
function setupWindow(win) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = WINDOW_SPACING;
}

/**
 * パネルの共通設定を適用する
 * @param {Panel} panel 対象パネル
 * @returns {void}
 */
function setupPanel(panel) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = PANEL_SPACING;
}

/**
 * 行グループの共通設定を適用する
 * @param {Group} group 対象グループ
 * @returns {void}
 */
function setupRow(group) {
    group.orientation = "row";
    group.alignment = ["fill", "top"];
    group.alignChildren = ["left", "center"];
    group.spacing = PANEL_SPACING;
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
        title: { ja: "検索文字列を削除して段落スタイルを適用", en: "Apply Paragraph Style and Delete Marker" }
    },
    panel: {
        searchText: { ja: "検索文字列", en: "Find What" },
        style:      { ja: "置換スタイル", en: "Styles to Apply" },
        options:    { ja: "オプション", en: "Options" }
    },
    marker: {
        detected:        { ja: "%1（%2箇所）", en: "%1 (%2)" },
        numberedDot:     { ja: "1. 2. 3.", en: "1. 2. 3." },
        numberedParen:   { ja: "1) 2) 3)", en: "1) 2) 3)" },
        kanjiParen:      { ja: "（一）（二）（三）", en: "(1) (2) (3) in kanji" },
        kanjiDot:        { ja: "一. 二. 三.", en: "1. 2. 3. in kanji" },
        kanjiComma:      { ja: "一、二、三、", en: "1, 2, 3, in kanji" },
        kanjiCloseParen: { ja: "一）二）三）", en: "1) 2) 3) in kanji" },
        bold:            { ja: "**文字列**", en: "**text**" }
    },
    mode: {
        text: { ja: "文字列を指定", en: "Enter text" },
        auto: { ja: "自動判別", en: "Detect automatically" }
    },
    field: {
        paragraphStyle: { ja: "段落スタイル", en: "Paragraph style" },
        characterStyle: { ja: "文字スタイル", en: "Character style" },
        scope:          { ja: "検索対象", en: "Search" },
        matchLevel:     { ja: "Markdown のレベルを合わせる", en: "Match the paragraph style to the level" },
        applyAllLevels: { ja: "Markdown の連続適用（見出し）", en: "Apply every Markdown heading level" }
    },
    option: {
        none: { ja: "（なし）", en: "(None)" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" },
        ok:     { ja: "OK", en: "OK" }
    },
    tooltip: {
        searchText:     { ja: "削除したい目印の文字列を入力します。例：###", en: "Enter the marker text to delete, for example ###." },
        auto:           { ja: "検索対象の行頭に繰り返し現れる記号（Markdown 記法を含む）を多い順に並べます。検索対象を変えると再判定します。", en: "Lists the line-head markers that repeat in the search target, most frequent first. Changing the search target runs the detection again." },
        paragraphStyle: { ja: "目印が見つかった段落に適用する段落スタイルです。", en: "The paragraph style applied to paragraphs that contain the marker." },
        characterStyle: { ja: "（なし）以外を選ぶと、その段落全体に文字スタイルを適用します。", en: "Anything other than (None) applies a character style to the whole paragraph." },
        scope:          { ja: "「ストーリー」は、テキストの選択またはカーソル位置が必要です。", en: "Story needs a text selection or the cursor placed in a story." },
        matchLevel:     { ja: "### を選ぶと h3／heading 3 の段落スタイルを自動で選びます。選び直しても構いません。", en: "Selecting ### picks the paragraph style named h3 / heading 3. You can still change it by hand." },
        applyAllLevels: { ja: "見出し # 〜 ###### の6レベルをまとめて処理し、各レベルに h1／heading 1 …… の段落スタイルを割り当てます。検索文字列と置換スタイルの設定は使いません。対応するスタイルが無いレベルは飛ばします。", en: "Processes all six heading levels (# through ######) in one run, applying h1 / heading 1 and so on. The Find What and Styles panels are not used. Levels without a matching style are skipped." },
        cancel:         { ja: "何も変更せずに閉じます。", en: "Close without making any changes." },
        ok:             { ja: "目印を削除してスタイルを適用します。", en: "Delete the markers and apply the styles." }
    },
    scope: {
        allDocuments: { ja: "すべてのドキュメント", en: "All Documents" },
        document:     { ja: "ドキュメント", en: "Document" },
        story:        { ja: "ストーリー", en: "Story" }
    },
    error: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noTarget:   { ja: "検索対象が見つかりません。\nテキストを選択するか、カーソルを配置してください。", en: "No search target found.\nSelect text or place the cursor in a story." }
    },
    result: {
        processed: { ja: "%1箇所を処理しました。", en: "Processed %1 location(s)." },
        skipped:   { ja: "段落スタイルが見つからないため、次のドキュメントはスキップしました。", en: "Skipped these documents because the paragraph style was not found." },
        partial:   { ja: "文字スタイルが見つからないため、次のドキュメントは段落スタイルだけ適用しました。", en: "Applied only the paragraph style in these documents because the character style was not found." }
    },
    undo: {
        apply: { ja: "検索文字列を削除して段落スタイルを適用", en: "Apply Paragraph Style and Delete Marker" }
    }
};

/**
 * ドット区切りキーでラベルを取得する
 * @param {string} labelKey 例: "field.paragraphStyle"
 * @returns {string} 現在の言語のラベル文字列。見つからない場合はキーをそのまま返す
 */
function getLabel(labelKey) {
    var labelNode = LABELS;
    var keyParts = labelKey.split(".");
    for (var i = 0; i < keyParts.length; i++) {
        labelNode = labelNode[keyParts[i]];
        if (!labelNode) return labelKey;
    }
    return labelNode[currentLang] || labelNode.en || labelKey;
}

/**
 * コロン付きラベルを取得する（日本語は全角コロン、英語は半角コロン）
 * @param {string} labelKey 例: "field.paragraphStyle"
 * @returns {string} コロンを付与したラベル文字列
 */
function getLabelWithColon(labelKey) {
    return getLabel(labelKey) + (currentLang === "ja" ? "：" : ":");
}

/**
 * ラベル内のプレースホルダー（%1, %2 …）を値で置き換える
 * @param {string} template プレースホルダーを含む文字列
 * @param {array} replacementValues 差し込む値
 * @returns {string} 置き換え後の文字列
 */
function formatLabel(template, replacementValues) {
    var formattedText = template;
    for (var i = 0; i < replacementValues.length; i++) {
        formattedText = formattedText.replace("%" + (i + 1), replacementValues[i]);
    }
    return formattedText;
}

// =========================================
// 共通処理 / Helpers
// =========================================

/**
 * 名前でスタイルを探す（スタイルグループ内も対象）
 * @param {array} styleCollection allParagraphStyles / allCharacterStyles
 * @param {string} styleName 探すスタイル名
 * @returns {object} 見つかったスタイル。無い場合は null
 */
function findStyleByName(styleCollection, styleName) {
    for (var i = 0; i < styleCollection.length; i++) {
        if (styleCollection[i].name === styleName) return styleCollection[i];
    }

    return null;
}

/**
 * スタイル名の一覧を取得する
 * @param {array} styleCollection allParagraphStyles / allCharacterStyles
 * @param {boolean} includeNoneOption 先頭に「（なし）」を追加するかどうか
 * @returns {array} スタイル名の配列
 */
function getStyleNames(styleCollection, includeNoneOption) {
    var styleNames = includeNoneOption ? [getLabel("option.none")] : [];

    for (var i = 0; i < styleCollection.length; i++) {
        styleNames.push(styleCollection[i].name);
    }

    return styleNames;
}

/**
 * GREP 検索用に正規表現の特殊文字をエスケープする
 * @param {string} plainText エスケープする文字列
 * @returns {string} エスケープした文字列
 */
function escapeGrepText(plainText) {
    return plainText.replace(/([\\^$.|?*+()\[\]{}])/g, "\\$1");
}

/**
 * 行頭の目印を GREP パターンに組み立てる
 * @param {string} markerText 目印の文字列
 * @returns {string} 行頭に固定した GREP パターン
 * @description 区切りが無い目印は、同じ記号が続く場合に一致しないよう先読みで止める。
 *              これが無いと "#" が "## 見出し" の1文字目に一致してしまう
 */
function buildLineHeadGrep(markerText) {
    var grepPattern = "^" + escapeGrepText(markerText);

    if (/[ \t　]$/.test(markerText)) return grepPattern;

    var lastCharacter = markerText.charAt(markerText.length - 1);

    return grepPattern + "(?!" + escapeGrepText(lastCharacter) + ")";
}

/**
 * 同じ文字を繰り返した文字列を作る
 * @param {string} character 繰り返す文字
 * @param {number} repeatCount 繰り返す回数
 * @returns {string} 繰り返した文字列
 */
function repeatCharacter(character, repeatCount) {
    var repeated = "";

    for (var i = 0; i < repeatCount; i++) {
        repeated += character;
    }

    return repeated;
}

/**
 * 見出しマーカーからレベルを取り出す
 * @param {string} markerText 検索文字列
 * @returns {number} 見出しレベル（1〜6）。見出しマーカーでない場合は 0
 */
function getHeadingLevel(markerText) {
    if (!HEADING_MARKER_PATTERN.test(markerText)) return 0;

    return markerText.replace(/[^#]/g, "").length;
}

/**
 * 見出しレベルに対応する段落スタイルをドロップダウンから探す
 * @param {DropDownList} styleDropdown 段落スタイルのドロップダウン
 * @param {number} headingLevel 見出しレベル（1〜6）
 * @returns {number} 見つかった項目のインデックス。無い場合は -1
 */
function findHeadingStyleIndex(styleDropdown, headingLevel) {
    for (var i = 0; i < HEADING_STYLE_NAME_FORMATS.length; i++) {
        var headingStyleName = formatLabel(HEADING_STYLE_NAME_FORMATS[i], [headingLevel]);

        for (var j = 0; j < styleDropdown.items.length; j++) {
            if (styleDropdown.items[j].text === headingStyleName) return j;
        }
    }

    return -1;
}

/**
 * 見出しレベルに対応する段落スタイルを探す
 * @param {Document} targetDoc 対象ドキュメント
 * @param {number} headingLevel 見出しレベル（1〜6）
 * @returns {ParagraphStyle} 見つかった段落スタイル。無い場合は null
 */
function findHeadingStyle(targetDoc, headingLevel) {
    var paragraphStyles = targetDoc.allParagraphStyles;

    for (var i = 0; i < HEADING_STYLE_NAME_FORMATS.length; i++) {
        var headingStyle = findStyleByName(paragraphStyles,
            formatLabel(HEADING_STYLE_NAME_FORMATS[i], [headingLevel]));

        if (headingStyle) return headingStyle;
    }

    return null;
}

/**
 * ドキュメントから適用するスタイルを取得する
 * @param {Document} targetDoc 対象ドキュメント
 * @param {string} paragraphStyleName 段落スタイル名
 * @param {string} characterStyleName 文字スタイル名。適用しない場合は null
 * @param {number} headingLevel 見出しレベル（1〜6）。段落スタイル名を使う場合は 0
 * @returns {object} { paragraphStyle: ParagraphStyle, characterStyle: CharacterStyle,
 *                     missingCharacterStyle: string }。
 *                   段落スタイルが無い場合は { skip: true, missing: string }。
 *                   missing が null のときは報告せずに飛ばす
 */
function resolveStyles(targetDoc, paragraphStyleName, characterStyleName, headingLevel) {
    var paragraphStyle;

    if (headingLevel > 0) {
        /* その書類に無いレベルは報告せずに飛ばす / A level the document does not define is skipped quietly */
        paragraphStyle = findHeadingStyle(targetDoc, headingLevel);
        if (!paragraphStyle) return { skip: true, missing: null };
    } else {
        paragraphStyle = findStyleByName(targetDoc.allParagraphStyles, paragraphStyleName);
        if (!paragraphStyle) return { skip: true, missing: paragraphStyleName };
    }

    var characterStyle = null;
    var missingCharacterStyle = null;

    if (characterStyleName) {
        characterStyle = findStyleByName(targetDoc.allCharacterStyles, characterStyleName);

        /* 文字スタイルが無くても段落スタイルは適用する / Apply the paragraph style even without it */
        if (!characterStyle) missingCharacterStyle = characterStyleName;
    }

    return {
        paragraphStyle: paragraphStyle,
        characterStyle: characterStyle,
        missingCharacterStyle: missingCharacterStyle
    };
}

/**
 * 実行する検索の一覧を作る
 * @param {object} dialogSettings ダイアログの設定
 * @returns {array} { searchText: string, useGrep: boolean, headingLevel: number } の配列
 */
function buildSearchPlans(dialogSettings) {
    if (!dialogSettings.applyAllLevels) {
        return [{
            searchText: dialogSettings.searchText,
            useGrep: dialogSettings.useGrep,
            headingLevel: 0
        }];
    }

    var searchPlans = [];

    /* 先読みでレベルを確定し、後ろの区切り（半角／全角スペース・タブ）は有無どちらも拾う /
       A lookahead pins the level, and the separator after it is optional */
    for (var headingLevel = MAX_HEADING_LEVEL; headingLevel >= 1; headingLevel--) {
        searchPlans.push({
            searchText: "^" + repeatCharacter("#", headingLevel) +
                "(?!#)" + HEADING_SEPARATOR_PATTERN,
            useGrep: true,
            headingLevel: headingLevel
        });
    }

    return searchPlans;
}

/**
 * スキップしたドキュメントを重複なく記録する
 * @param {array} skippedDocuments 記録先の配列
 * @param {string} skippedEntry 「ドキュメント名: スタイル名」の文字列
 * @returns {void}
 */
function addSkippedDocument(skippedDocuments, skippedEntry) {
    for (var i = 0; i < skippedDocuments.length; i++) {
        if (skippedDocuments[i] === skippedEntry) return;
    }

    skippedDocuments.push(skippedEntry);
}

/**
 * 判別した行頭記号の表示ラベル一覧を取得する
 * @param {array} detectedMarkers { markerText: string, labelKey: string, count: number } の配列
 * @returns {array} ドロップダウンに並べる文字列の配列
 */
function getDetectedMarkerLabels(detectedMarkers) {
    var markerLabels = [];

    for (var i = 0; i < detectedMarkers.length; i++) {
        var displayText = detectedMarkers[i].labelKey ?
            getLabel(detectedMarkers[i].labelKey) :
            detectedMarkers[i].markerText;

        markerLabels.push(formatLabel(getLabel("marker.detected"),
            [displayText, detectedMarkers[i].count]));
    }

    return markerLabels;
}

/**
 * 行ラベルを追加する（幅を固定して右揃え）
 * @param {Group} parentGroup 追加先の行グループ
 * @param {string} labelKey ラベルキー
 * @returns {StaticText} 追加したラベル
 */
function addRowLabel(parentGroup, labelKey) {
    var rowLabel = parentGroup.add("statictext", undefined, getLabelWithColon(labelKey));
    rowLabel.preferredSize.width = LABEL_WIDTH;
    rowLabel.justify = "right";
    return rowLabel;
}

/**
 * 選択範囲の親ストーリーを取得する
 * @returns {Story} 選択（またはカーソル位置）のストーリー。取得できない場合は null
 */
function getSelectedStory() {
    if (app.selection.length === 0) return null;

    try {
        return app.selection[0].parentStory;
    } catch (e) {
        /* テキスト以外が選択されている / Selection is not text */
        return null;
    }
}

/**
 * スコープに応じた検索対象を取得する
 * @param {number} searchScope SCOPE_* のいずれか
 * @param {Document} activeDoc アクティブドキュメント
 * @returns {array} { doc: Document, searchRange: Document|Story } の配列
 */
function getSearchTargets(searchScope, activeDoc) {
    var searchTargets = [];

    if (searchScope === SCOPE_ALL_DOCUMENTS) {
        for (var i = 0; i < app.documents.length; i++) {
            searchTargets.push({ doc: app.documents[i], searchRange: app.documents[i] });
        }
        return searchTargets;
    }

    if (searchScope === SCOPE_DOCUMENT) {
        searchTargets.push({ doc: activeDoc, searchRange: activeDoc });
        return searchTargets;
    }

    /* ストーリーは選択（またはカーソル位置）が必要 / The story scope needs a text selection */
    var story = getSelectedStory();
    if (story) searchTargets.push({ doc: activeDoc, searchRange: story });

    return searchTargets;
}

/**
 * 検索範囲に含まれる段落を取得する
 * @param {object} searchRange 検索対象（Document / Story）
 * @returns {array} 段落の配列
 */
function getParagraphsInRange(searchRange) {
    /* Story はそのまま段落を持つ / A Story exposes paragraphs directly */
    if (!(searchRange instanceof Document)) {
        return searchRange.paragraphs.everyItem().getElements();
    }

    /* Document はストーリーごとにたどる / A Document is walked story by story */
    var paragraphs = [];

    for (var i = 0; i < searchRange.stories.length; i++) {
        paragraphs = paragraphs.concat(
            searchRange.stories[i].paragraphs.everyItem().getElements());
    }

    return paragraphs;
}

/**
 * 検索オプションを既定値にそろえる
 * @param {object} findChangeOptions findChangeTextOptions / findChangeGrepOptions
 * @returns {void}
 */
function applyFindChangeOptions(findChangeOptions) {
    findChangeOptions.widthSensitive = FIND_WIDTH_SENSITIVE;
    findChangeOptions.kanaSensitive = FIND_KANA_SENSITIVE;
    findChangeOptions.includeFootnotes = FIND_INCLUDE_FOOTNOTES;
    findChangeOptions.includeMasterPages = FIND_INCLUDE_MASTER_PAGES;
    findChangeOptions.includeHiddenLayers = FIND_INCLUDE_HIDDEN_LAYERS;
    findChangeOptions.includeLockedLayersForFind = FIND_INCLUDE_LOCKED_LAYERS;
    findChangeOptions.includeLockedStoriesForFind = FIND_INCLUDE_LOCKED_STORIES;
}

/**
 * 検索条件を初期化する（テキスト検索・GREP 検索の両方）
 * @returns {void}
 */
function resetFindPreferences() {
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    applyFindChangeOptions(app.findChangeTextOptions);
    applyFindChangeOptions(app.findChangeGrepOptions);

    /* テキスト検索だけが持つオプション / Options that only text search has */
    app.findChangeTextOptions.caseSensitive = FIND_CASE_SENSITIVE;
    app.findChangeTextOptions.wholeWord = false;
}

/**
 * 段落の行頭にある目印を判定する
 * @param {string} paragraphText 段落の文字列
 * @returns {object} { key: string, markerText: string, searchText: string, useGrep: boolean,
 *                     labelKey: string }。目印がない場合は null
 * @description 行頭記号の searchText は行頭（^）に固定した GREP。段落の途中にある同じ記号は拾わない。
 *              太字（**文字列**）だけは囲みなので、行頭に固定せず ** そのものを削除する
 */
function matchLineHeadMarker(paragraphText) {
    for (var i = 0; i < NUMBERED_LIST_PATTERNS.length; i++) {
        if (!NUMBERED_LIST_PATTERNS[i].regex.test(paragraphText)) continue;

        return {
            key: NUMBERED_LIST_PATTERNS[i].labelKey,
            markerText: "",
            searchText: NUMBERED_LIST_PATTERNS[i].pattern,
            useGrep: true,
            labelKey: NUMBERED_LIST_PATTERNS[i].labelKey
        };
    }

    var matched = LINE_HEAD_MARKER_PATTERN.exec(paragraphText);

    if (matched && !EXCLUDED_LINE_HEAD_PATTERN.test(matched[0])) {
        return {
            key: matched[0],
            markerText: matched[0],
            searchText: buildLineHeadGrep(matched[0]),
            useGrep: true,
            labelKey: null
        };
    }

    /* 太字は行頭に限らないので、行頭記号を拾えなかったときに判定する /
       Bold is not tied to the line head, so it is checked last */
    if (BOLD_PATTERN.test(paragraphText)) {
        return {
            key: "marker.bold",
            markerText: "",
            searchText: BOLD_SEARCH_TEXT,
            useGrep: true,
            labelKey: "marker.bold"
        };
    }

    return null;
}

/**
 * 検索対象を調べ、行頭に繰り返し現れる記号を多い順に集める
 * @param {array} searchTargets 検索対象の配列
 * @returns {array} { markerText: string, searchText: string, useGrep: boolean,
 *                   labelKey: string, count: number } の配列
 */
function detectLineHeadMarkers(searchTargets) {
    var markerEntries = {};

    for (var i = 0; i < searchTargets.length; i++) {
        var paragraphs = getParagraphsInRange(searchTargets[i].searchRange);

        for (var j = 0; j < paragraphs.length; j++) {
            var entry = matchLineHeadMarker(paragraphs[j].contents);
            if (!entry) continue;

            if (markerEntries[entry.key]) {
                markerEntries[entry.key].count++;
            } else {
                entry.count = 1;
                markerEntries[entry.key] = entry;
            }
        }
    }

    var detectedMarkers = [];

    for (var key in markerEntries) {
        if (!markerEntries.hasOwnProperty(key)) continue;
        if (markerEntries[key].count < MIN_REPEAT_COUNT) continue;

        detectedMarkers.push(markerEntries[key]);
    }

    /* 多い順、同数なら長い記号を先に（### を # より先に） / Most frequent first, longer marker wins a tie */
    detectedMarkers.sort(function (a, b) {
        if (b.count !== a.count) return b.count - a.count;
        return b.markerText.length - a.markerText.length;
    });

    return detectedMarkers;
}

/**
 * 検索範囲内の検索文字列にスタイルを適用して削除する
 * @param {object} searchRange 検索対象（Document / Story）
 * @param {ParagraphStyle} paragraphStyle 適用する段落スタイル
 * @param {CharacterStyle} characterStyle 適用する文字スタイル。適用しない場合は null
 * @param {boolean} useGrep GREP 検索を使うかどうか
 * @returns {number} 処理した箇所数
 */
function applyStylesAndRemoveMarker(searchRange, paragraphStyle, characterStyle, useGrep) {
    var foundTexts = useGrep ? searchRange.findGrep() : searchRange.findText();
    var appliedCount = 0;

    /* 後ろから処理 / Process from the end */
    for (var i = foundTexts.length - 1; i >= 0; i--) {
        try {
            var foundText = foundTexts[i];

            /* 検索文字列を含む段落にスタイルを適用 / Apply the styles to the paragraph */
            var targetParagraph = foundText.paragraphs[0];
            targetParagraph.appliedParagraphStyle = paragraphStyle;

            if (characterStyle) {
                targetParagraph.appliedCharacterStyle = characterStyle;
            }

            /* 検索文字列を削除 / Delete the marker */
            foundText.remove();

            appliedCount++;
        } catch (e) {
            /* 処理できない箇所はスキップ / Skip locations that cannot be processed */
        }
    }

    return appliedCount;
}

/**
 * ダイアログを表示して設定を取得する
 * @param {Document} activeDoc アクティブドキュメント
 * @returns {object} { searchText: string, useGrep: boolean, applyAllLevels: boolean,
 *                     paragraphStyleName: string, characterStyleName: string,
 *                     searchScope: number }。キャンセル時は null
 * @description 「自動判別」の候補は検索対象を変えるたびに取り直す
 */
function showDialog(activeDoc) {
    var dialog = new Window("dialog", getLabel("dialog.title"));
    setupWindow(dialog);

    /* 検索文字列パネル / Find What panel */
    var searchPanel = dialog.add("panel", undefined, getLabel("panel.searchText"));
    setupPanel(searchPanel);

    var textModeRow = searchPanel.add("group");
    setupRow(textModeRow);
    var textModeRadio = textModeRow.add("radiobutton", undefined, getLabel("mode.text"));
    textModeRadio.preferredSize.width = LABEL_WIDTH;
    var searchTextField = textModeRow.add("edittext", undefined, DEFAULT_SEARCH_TEXT);
    searchTextField.characters = SEARCH_TEXT_LENGTH;
    searchTextField.alignment = ["fill", "center"];
    searchTextField.helpTip = getLabel("tooltip.searchText");

    var autoModeRow = searchPanel.add("group");
    setupRow(autoModeRow);
    var autoModeRadio = autoModeRow.add("radiobutton", undefined, getLabel("mode.auto"));
    autoModeRadio.preferredSize.width = LABEL_WIDTH;
    autoModeRadio.helpTip = getLabel("tooltip.auto");
    var detectedMarkerDropdown = autoModeRow.add("dropdownlist", undefined, []);
    detectedMarkerDropdown.alignment = ["fill", "center"];
    detectedMarkerDropdown.helpTip = getLabel("tooltip.auto");

    /* 置換スタイルパネル / Styles panel */
    var stylePanel = dialog.add("panel", undefined, getLabel("panel.style"));
    setupPanel(stylePanel);

    var paragraphStyleRow = stylePanel.add("group");
    setupRow(paragraphStyleRow);
    addRowLabel(paragraphStyleRow, "field.paragraphStyle");
    var paragraphStyleDropdown = paragraphStyleRow.add(
        "dropdownlist", undefined, getStyleNames(activeDoc.allParagraphStyles, false));
    paragraphStyleDropdown.selection = 0;
    paragraphStyleDropdown.alignment = ["fill", "center"];
    paragraphStyleDropdown.helpTip = getLabel("tooltip.paragraphStyle");

    var characterStyleRow = stylePanel.add("group");
    setupRow(characterStyleRow);
    addRowLabel(characterStyleRow, "field.characterStyle");
    var characterStyleDropdown = characterStyleRow.add(
        "dropdownlist", undefined, getStyleNames(activeDoc.allCharacterStyles, true));
    characterStyleDropdown.selection = 0;
    characterStyleDropdown.alignment = ["fill", "center"];
    characterStyleDropdown.helpTip = getLabel("tooltip.characterStyle");

    /* オプションパネル / Options panel */
    var optionsPanel = dialog.add("panel", undefined, getLabel("panel.options"));
    setupPanel(optionsPanel);

    var scopeRow = optionsPanel.add("group");
    setupRow(scopeRow);
    addRowLabel(scopeRow, "field.scope");
    var searchScopeDropdown = scopeRow.add("dropdownlist", undefined, [
        getLabel("scope.allDocuments"),
        getLabel("scope.document"),
        getLabel("scope.story")
    ]);
    searchScopeDropdown.selection = DEFAULT_SCOPE;
    searchScopeDropdown.alignment = ["fill", "center"];
    searchScopeDropdown.helpTip = getLabel("tooltip.scope");

    var matchLevelRow = optionsPanel.add("group");
    setupRow(matchLevelRow);
    var matchLevelCheckbox =
        matchLevelRow.add("checkbox", undefined, getLabel("field.matchLevel"));
    matchLevelCheckbox.value = DEFAULT_MATCH_LEVEL;
    matchLevelCheckbox.helpTip = getLabel("tooltip.matchLevel");

    var applyAllLevelsRow = optionsPanel.add("group");
    setupRow(applyAllLevelsRow);
    var applyAllLevelsCheckbox =
        applyAllLevelsRow.add("checkbox", undefined, getLabel("field.applyAllLevels"));
    applyAllLevelsCheckbox.value = DEFAULT_APPLY_ALL_LEVELS;
    applyAllLevelsCheckbox.helpTip = getLabel("tooltip.applyAllLevels");

    /* ボタンエリア / Button row */
    var btnRowGroup = dialog.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.alignment = ["fill", "bottom"];

    var spacer = btnRowGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 0;

    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.alignChildren = ["right", "center"];
    var btnCancel = btnRightGroup.add(
        "button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOk = btnRightGroup.add(
        "button", undefined, getLabel("button.ok"), { name: "ok" });
    btnCancel.preferredSize.width = BUTTON_WIDTH;
    btnOk.preferredSize.width = BUTTON_WIDTH;
    btnCancel.helpTip = getLabel("tooltip.cancel");
    btnOk.helpTip = getLabel("tooltip.ok");

    var searchModeRadios = [textModeRadio, autoModeRadio];
    var detectedMarkers = [];
    var detectedMarkerCache = {};

    /**
     * 選ばれた指定方法に合わせて、ラジオボタンと入力欄の状態をそろえる
     * @param {number} searchMode SEARCH_MODE_* のいずれか
     * @returns {void}
     */
    function selectSearchMode(searchMode) {
        for (var i = 0; i < searchModeRadios.length; i++) {
            searchModeRadios[i].value = (i === searchMode);
        }

        searchTextField.enabled = (searchMode === SEARCH_MODE_TEXT);
        detectedMarkerDropdown.enabled =
            (searchMode === SEARCH_MODE_AUTO) && (detectedMarkers.length > 0);
        btnOk.enabled = (searchMode === SEARCH_MODE_AUTO) ?
            (detectedMarkers.length > 0) :
            (searchTextField.text !== "");

        updateHeadingOptions();
    }

    /**
     * 現在選ばれている目印の文字列を取得する（GREP に組み立てる前の記号そのまま）
     * @returns {string} 目印の文字列。取得できない場合は空文字列
     */
    function getCurrentMarkerText() {
        if (getSelectedSearchMode() !== SEARCH_MODE_AUTO) return searchTextField.text;

        return detectedMarkerDropdown.selection ?
            detectedMarkers[detectedMarkerDropdown.selection.index].markerText :
            "";
    }

    /**
     * 見出しマーカーの選択に合わせて、Markdown 関連のオプションを更新する
     * @returns {void}
     */
    function updateHeadingOptions() {
        var headingLevel = getHeadingLevel(getCurrentMarkerText());
        var isHeadingMarker = (headingLevel > 0);

        /* 見出しマーカー以外を選んでいる間は両方とも使えない。チェック状態は残す /
           Both options need a heading marker; the checked state is kept as it is */
        var applyAllLevels = isHeadingMarker && applyAllLevelsCheckbox.value;

        applyAllLevelsCheckbox.enabled = isHeadingMarker;

        /* 連続適用中は検索文字列もスタイルもレベルごとに決まるので、両パネルを使わない /
           While applying every level both panels are unused */
        searchPanel.enabled = !applyAllLevels;
        stylePanel.enabled = !applyAllLevels;
        matchLevelCheckbox.enabled = isHeadingMarker && !applyAllLevels;

        if (!matchLevelCheckbox.enabled || !matchLevelCheckbox.value) return;

        var styleIndex = findHeadingStyleIndex(paragraphStyleDropdown, headingLevel);
        if (styleIndex >= 0) paragraphStyleDropdown.selection = styleIndex;
    }

    /**
     * 現在選ばれている指定方法を取得する
     * @returns {number} SEARCH_MODE_* のいずれか
     */
    function getSelectedSearchMode() {
        for (var i = 0; i < searchModeRadios.length; i++) {
            if (searchModeRadios[i].value) return i;
        }
        return SEARCH_MODE_TEXT;
    }

    /**
     * 検索対象を調べ直し、見つかった行頭記号をドロップダウンに展開する
     * @returns {void}
     */
    function refreshDetectedMarkers() {
        var searchScope = searchScopeDropdown.selection.index;

        /* 同じ検索対象を選び直したときは走査し直さない / Reuse the result for a scope already scanned */
        if (!detectedMarkerCache[searchScope]) {
            detectedMarkerCache[searchScope] = detectLineHeadMarkers(
                getSearchTargets(searchScope, activeDoc));
        }

        detectedMarkers = detectedMarkerCache[searchScope];

        var markerLabels = getDetectedMarkerLabels(detectedMarkers);

        detectedMarkerDropdown.removeAll();
        for (var i = 0; i < markerLabels.length; i++) {
            detectedMarkerDropdown.add("item", markerLabels[i]);
        }
        detectedMarkerDropdown.selection = (detectedMarkers.length > 0) ? 0 : null;

        autoModeRadio.enabled = (detectedMarkers.length > 0);

        /* 候補が無くなったら文字列指定に戻す / Fall back to manual entry when nothing was detected */
        var searchMode = getSelectedSearchMode();
        if (detectedMarkers.length === 0 && searchMode === SEARCH_MODE_AUTO) {
            searchMode = SEARCH_MODE_TEXT;
        }
        selectSearchMode(searchMode);
    }

    textModeRadio.onClick = function () { selectSearchMode(SEARCH_MODE_TEXT); };
    autoModeRadio.onClick = function () { selectSearchMode(SEARCH_MODE_AUTO); };
    searchTextField.onChanging = function () { selectSearchMode(getSelectedSearchMode()); };
    detectedMarkerDropdown.onChange = updateHeadingOptions;
    matchLevelCheckbox.onClick = updateHeadingOptions;
    applyAllLevelsCheckbox.onClick = updateHeadingOptions;
    searchScopeDropdown.onChange = refreshDetectedMarkers;

    selectSearchMode(DEFAULT_SEARCH_MODE);
    refreshDetectedMarkers();

    if (dialog.show() !== 1) return null;

    var selectedMarker = (getSelectedSearchMode() === SEARCH_MODE_AUTO) ?
        detectedMarkers[detectedMarkerDropdown.selection.index] :
        { searchText: searchTextField.text, useGrep: false };

    var applyAllLevels = applyAllLevelsCheckbox.enabled && applyAllLevelsCheckbox.value;

    /* ディム表示にしたパネルの設定は使わない / Settings in the dimmed panels are not used */
    var characterStyleName = (applyAllLevels || characterStyleDropdown.selection.index === 0) ?
        null :
        characterStyleDropdown.selection.text;

    return {
        searchText: selectedMarker.searchText,
        useGrep: selectedMarker.useGrep,
        applyAllLevels: applyAllLevels,
        paragraphStyleName: paragraphStyleDropdown.selection.text,
        characterStyleName: characterStyleName,
        searchScope: searchScopeDropdown.selection.index
    };
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {
    if (app.documents.length === 0) {
        alert(getLabel("error.noDocument"));
        return;
    }

    var doc = app.activeDocument;

    var dialogSettings = showDialog(doc);
    if (!dialogSettings) return;

    var searchTargets = getSearchTargets(dialogSettings.searchScope, doc);
    if (searchTargets.length === 0) {
        alert(getLabel("error.noTarget"));
        return;
    }

    var searchPlans = buildSearchPlans(dialogSettings);
    var processedCount = 0;
    var skippedDocuments = [];
    var partialDocuments = [];

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(function () {
        for (var planIndex = 0; planIndex < searchPlans.length; planIndex++) {
            var searchPlan = searchPlans[planIndex];

            /* 検索条件を初期化 / Reset the find preferences */
            resetFindPreferences();

            if (searchPlan.useGrep) {
                app.findGrepPreferences.findWhat = searchPlan.searchText;
            } else {
                app.findTextPreferences.findWhat = searchPlan.searchText;
            }

            for (var i = 0; i < searchTargets.length; i++) {
                var resolvedStyles = resolveStyles(
                    searchTargets[i].doc,
                    dialogSettings.paragraphStyleName,
                    dialogSettings.characterStyleName,
                    searchPlan.headingLevel);

                /* 段落スタイルを持たないドキュメントはスキップ / Skip documents without the paragraph style */
                if (resolvedStyles.skip) {
                    if (resolvedStyles.missing) {
                        addSkippedDocument(skippedDocuments,
                            searchTargets[i].doc.name + ": " + resolvedStyles.missing);
                    }
                    continue;
                }

                if (resolvedStyles.missingCharacterStyle) {
                    addSkippedDocument(partialDocuments,
                        searchTargets[i].doc.name + ": " + resolvedStyles.missingCharacterStyle);
                }

                processedCount += applyStylesAndRemoveMarker(
                    searchTargets[i].searchRange,
                    resolvedStyles.paragraphStyle,
                    resolvedStyles.characterStyle,
                    searchPlan.useGrep);
            }
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.apply"));

    /* 検索条件をクリア / Clear the find preferences */
    resetFindPreferences();

    var resultMessage = formatLabel(getLabel("result.processed"), [processedCount]);

    if (partialDocuments.length > 0) {
        resultMessage += "\n\n" + getLabel("result.partial") +
            "\n" + partialDocuments.join("\n");
    }

    if (skippedDocuments.length > 0) {
        resultMessage += "\n\n" + getLabel("result.skipped") +
            "\n" + skippedDocuments.join("\n");
    }

    alert(resultMessage);
})();
