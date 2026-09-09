#target indesign

/*

### 概要

行頭の目印（Markdown の見出し・箇条書き・番号リストなど）を手がかりに段落スタイルと文字スタイルを適用し、その目印を後ろに続くスペースごと削除します。目印は検索対象から自動判別でき、対象箇所の件数を確かめてから実行できます。

詳細は README を参照してください。

### Overview

Applies a paragraph style and a character style to paragraphs carrying a leading marker — Markdown headings, bullets, numbered lists and the like — then deletes the marker together with the spaces that follow it. Markers can be detected from the search target, and the number of matches is shown before the run.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdRemoveMarkerApplyStyle";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)"; /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-09";                  /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-10";                  /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdRemoveMarkerApplyStyle.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdRemoveMarkerApplyStyle.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n3a0d4c0dacdb"; /* 紹介記事 / article URL */

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
var HEADING_MARKER_PATTERN = /^#{1,6}[ \t　]*$/;
var MAX_HEADING_LEVEL = 6;

/* 目印に続く区切り（見つかれば一緒に削除する） / Separator after the marker (deleted together) */
/* 半角スペース・全角スペース・タブ。連続していてもまとめて削除し、無くてもよい /
   Spaces, full-width spaces and tabs: all of them are deleted, and none is fine too */
var TRAILING_SEPARATOR_PATTERN = "[ 　\\t]*";

/* 見出しレベルに対応する段落スタイル名（先に見つかったものを使う） / Paragraph style names per heading level */
var HEADING_STYLE_NAME_FORMATS = ["h%1", "H%1", "heading %1", "Heading %1"];

/* 検索オプション（前回の「検索/置換」の設定を引き継がない） / Find options (never inherited) */
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
var RADIO_SPACING  = 4;                /* ラジオボタンの間隔 / radio button spacing */

/* 行の寸法 / Row metrics */
var LABEL_WIDTH        = 110; /* ラベル・ラジオボタンの幅 / label & radio button width */
var SEARCH_TEXT_LENGTH = 16;  /* 検索文字列入力欄の文字数 / search field length */
var BUTTON_WIDTH       = 90;  /* ダイアログボタンの幅 / dialog button width */
var COUNT_WIDTH        = 60;  /* 対象箇所の数値の幅 / match count width */

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
        title: { ja: "目印を削除してスタイルを適用", en: "Delete Markers and Apply Styles" }
    },
    panel: {
        searchText: { ja: "検索文字列", en: "Find What" },
        style:      { ja: "適用するスタイル", en: "Styles to Apply" },
        scope:      { ja: "検索対象", en: "Search In" },
        markdown:   { ja: "Markdown", en: "Markdown" }
    },
    marker: {
        detected:        { ja: "%1（%2箇所）", en: "%1 (%2 found)" },
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
        matchCount:     { ja: "対象箇所", en: "Matches" },
        matchLevel:     { ja: "見出しレベルに合わせる", en: "Match the heading level" },
        applyAllLevels: { ja: "連続適用（見出しのみ）", en: "Apply every Markdown heading level" }
    },
    option: {
        none: { ja: "（なし）", en: "(None)" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" },
        ok:     { ja: "OK", en: "OK" }
    },
    tooltip: {
        searchText: {
            ja: "削除したい目印の文字列を入力します。例：###\n同じ文字が続く並びの一部には一致しません。「##」は「###」に一致しません。\n目印に続くスペースやタブも一緒に削除します。",
            en: "Enter the marker text to delete, for example ###.\nIt never matches part of a longer run of the same character, so ## does not match ###.\nSpaces and tabs after the marker are deleted with it."
        },
        auto: {
            ja: "検索対象の行頭に繰り返し現れる記号（Markdown 記法を含む）を多い順に並べます。検索対象を変えると再判定します。非表示・ロックされたレイヤーのテキストは数えません。",
            en: "Lists the line-head markers that repeat in the search target, most frequent first. Changing the search target runs the detection again. Text on hidden or locked layers is not counted."
        },
        matchCount: {
            ja: "現在の設定で見つかる件数です。「自動判別」は段落数、「文字列を指定」は削除する箇所の数を表示します。非表示・ロックされたレイヤーのテキストは数えません。",
            en: "The number of matches for the current settings: paragraphs when a marker is detected automatically, markers to delete when the text is entered by hand. Text on hidden or locked layers is not counted."
        },
        paragraphStyle: { ja: "目印が見つかった段落に適用する段落スタイルです。", en: "The paragraph style applied to paragraphs that contain the marker." },
        characterStyle: { ja: "（なし）以外を選ぶと、その段落全体に文字スタイルを適用します。", en: "Anything other than (None) applies a character style to the whole paragraph." },
        scope: {
            allDocuments: {
                ja: "開いているドキュメントすべてが対象です。スタイルの一覧はアクティブドキュメントのものなので、同名のスタイルが無いドキュメントは飛ばします。",
                en: "Every open document is searched. The style lists come from the active document, so a document without a style of the same name is skipped."
            },
            document: {
                ja: "アクティブドキュメント全体が対象です。",
                en: "The whole active document is searched."
            },
            story: {
                ja: "テキストの選択、またはカーソルを置いたストーリーが対象です。未選択のまま実行すると対象が見つかりません。",
                en: "The story that holds the selection or the cursor is searched. Without one there is no search target."
            }
        },
        matchLevel:     { ja: "### を選ぶと h3／heading 3 の段落スタイルを自動で選びます。選び直しても構いません。", en: "Selecting ### picks the paragraph style named h3 / heading 3. You can still change it by hand." },
        applyAllLevels: { ja: "見出し # 〜 ###### の6レベルをまとめて処理し、各レベルに h1／heading 1 …… の段落スタイルを割り当てます。検索文字列と適用するスタイルの設定は使いません。対応するスタイルが無いレベルは飛ばします。", en: "Processes all six heading levels (# through ######) in one run, applying h1 / heading 1 and so on. The Find What and Styles panels are not used. Levels without a matching style are skipped." },
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
        processed: { ja: "%1箇所の目印を削除しました。", en: "Deleted %1 marker(s)." },
        skipped:   { ja: "段落スタイルが見つからないため、次のドキュメントはスキップしました。", en: "Skipped these documents because the paragraph style was not found." },
        partial:   { ja: "文字スタイルが見つからないため、次のドキュメントは段落スタイルだけ適用しました。", en: "Applied only the paragraph style in these documents because the character style was not found." }
    },
    undo: {
        apply: { ja: "目印を削除してスタイルを適用", en: "Delete Markers and Apply Styles" }
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
 * @description エスケープの対象は GREP と JavaScript の正規表現で共通なので、件数の集計にも使う
 */
function escapeGrepText(plainText) {
    return plainText.replace(/([\\^$.|?*+()\[\]{}])/g, "\\$1");
}

/**
 * 末尾の区切り（半角／全角スペース・タブ）を落とす
 * @param {string} text 対象の文字列
 * @returns {string} 末尾の区切りを落とした文字列
 * @description 区切りは目印と一緒に一致範囲で拾うので、パターンに含めない
 */
function trimTrailingSeparators(text) {
    return text.replace(/[ \t　]+$/, "");
}

/**
 * 目印の直後に同じ記号が続くときは一致させない先読みを組み立てる
 * @param {string} markerText 目印の文字列
 * @returns {string} 先読みのパターン
 * @description これが無いと "#" が "## 見出し" の1文字目に一致してしまう
 */
function buildRunGuard(markerText) {
    return "(?!" + escapeGrepText(markerText.charAt(markerText.length - 1)) + ")";
}

/**
 * 行頭の目印を GREP パターンに組み立てる
 * @param {string} markerText 目印の文字列（区切りを含まない記号そのもの）
 * @returns {string} 行頭に固定した GREP パターン
 * @description 続く区切り（スペース・タブ）は一致範囲に含めて、目印と一緒に削除する
 */
function buildLineHeadGrep(markerText) {
    return "^" + escapeGrepText(markerText) + buildRunGuard(markerText) +
        TRAILING_SEPARATOR_PATTERN;
}

/**
 * 検索文字列にぴったり一致する GREP パターンを組み立てる
 * @param {string} searchText 検索文字列
 * @returns {string} GREP パターン
 * @description 前後を後読み・先読みで止めることで、同じ文字が続く並びの一部には一致させない。
 *              これが無いと "##" が "### 見出し" の先頭2文字に一致してしまう。
 *              続く区切り（スペース・タブ）は一致範囲に含めて、目印と一緒に削除する
 */
function buildExactGrep(searchText) {
    var markerText = trimTrailingSeparators(searchText);

    if (markerText === "") return escapeGrepText(searchText);

    return "(?<!" + escapeGrepText(markerText.charAt(0)) + ")" +
        escapeGrepText(markerText) + buildRunGuard(markerText) +
        TRAILING_SEPARATOR_PATTERN;
}

/**
 * 同じ文字を繰り返した文字列を作る
 * @param {string} character 繰り返す文字
 * @param {number} repeatCount 繰り返す回数
 * @returns {string} 繰り返した文字列
 */
function repeatCharacter(character, repeatCount) {
    var repeatedText = "";

    for (var i = 0; i < repeatCount; i++) {
        repeatedText += character;
    }

    return repeatedText;
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
 * 見出しレベルに対応する段落スタイル名の候補を作る
 * @param {number} headingLevel 見出しレベル（1〜6）
 * @returns {array} スタイル名の配列（先に見つかったものを使う順）
 */
function getHeadingStyleNames(headingLevel) {
    var headingStyleNames = [];

    for (var i = 0; i < HEADING_STYLE_NAME_FORMATS.length; i++) {
        headingStyleNames.push(formatLabel(HEADING_STYLE_NAME_FORMATS[i], [headingLevel]));
    }

    return headingStyleNames;
}

/**
 * 見出しレベルに対応する段落スタイルをドロップダウンから探す
 * @param {DropDownList} styleDropdown 段落スタイルのドロップダウン
 * @param {number} headingLevel 見出しレベル（1〜6）
 * @returns {number} 見つかった項目のインデックス。無い場合は -1
 */
function findHeadingStyleIndex(styleDropdown, headingLevel) {
    var headingStyleNames = getHeadingStyleNames(headingLevel);

    for (var i = 0; i < headingStyleNames.length; i++) {
        for (var j = 0; j < styleDropdown.items.length; j++) {
            if (styleDropdown.items[j].text === headingStyleNames[i]) return j;
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
    var headingStyleNames = getHeadingStyleNames(headingLevel);

    for (var i = 0; i < headingStyleNames.length; i++) {
        var headingStyle = findStyleByName(targetDoc.allParagraphStyles, headingStyleNames[i]);

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
 *                   段落スタイルが無い場合は { skip: true, missingParagraphStyle: string }。
 *                   missingParagraphStyle が null のときは報告せずに飛ばす
 */
function resolveStyles(targetDoc, paragraphStyleName, characterStyleName, headingLevel) {
    var paragraphStyle;

    if (headingLevel > 0) {
        /* その書類に無いレベルは報告せずに飛ばす / A level the document does not define is skipped quietly */
        paragraphStyle = findHeadingStyle(targetDoc, headingLevel);
        if (!paragraphStyle) return { skip: true, missingParagraphStyle: null };
    } else {
        paragraphStyle = findStyleByName(targetDoc.allParagraphStyles, paragraphStyleName);
        if (!paragraphStyle) return { skip: true, missingParagraphStyle: paragraphStyleName };
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
 * @returns {array} { searchText: string, headingLevel: number } の配列
 */
function buildSearchPlans(dialogSettings) {
    if (!dialogSettings.applyAllLevels) {
        return [{ searchText: dialogSettings.searchText, headingLevel: 0 }];
    }

    var searchPlans = [];

    /* 先読みでレベルを確定し、後ろの区切り（半角／全角スペース・タブ）は続くだけまとめて拾う /
       A lookahead pins the level, and every separator after it is picked up */
    for (var headingLevel = MAX_HEADING_LEVEL; headingLevel >= 1; headingLevel--) {
        searchPlans.push({
            searchText: "^" + repeatCharacter("#", headingLevel) +
                "(?!#)" + TRAILING_SEPARATOR_PATTERN,
            headingLevel: headingLevel
        });
    }

    return searchPlans;
}

/**
 * 見つからなかったスタイルを書類ごとに重複なく記録する
 * @param {array} documentNotes 記録先の配列
 * @param {string} documentNote 「ドキュメント名: スタイル名」の文字列
 * @returns {void}
 */
function addDocumentNote(documentNotes, documentNote) {
    for (var i = 0; i < documentNotes.length; i++) {
        if (documentNotes[i] === documentNote) return;
    }

    documentNotes.push(documentNote);
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
 * 選ばれているラジオボタンの位置を取得する
 * @param {array} radioButtons ラジオボタンの配列
 * @returns {number} 選ばれている位置。無い場合は 0
 */
function getSelectedRadioIndex(radioButtons) {
    for (var i = 0; i < radioButtons.length; i++) {
        if (radioButtons[i].value) return i;
    }

    return 0;
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
 * ストーリーが検索対象になるかを判定する
 * @param {Story} story 対象のストーリー
 * @returns {boolean} 検索対象なら true
 * @description 検索オプションに合わせ、非表示レイヤー・ロックされたレイヤーだけに
 *              置かれているストーリーは対象外にする。件数の判定を実際の検索とそろえるため。
 *              フレームを持たないストーリーは対象に含める
 */
function isSearchableStory(story) {
    var textContainers = story.textContainers;

    if (!textContainers || textContainers.length === 0) return true;

    for (var i = 0; i < textContainers.length; i++) {
        var itemLayer = textContainers[i].itemLayer;

        if ((FIND_INCLUDE_HIDDEN_LAYERS || itemLayer.visible) &&
            (FIND_INCLUDE_LOCKED_LAYERS || !itemLayer.locked)) return true;
    }

    return false;
}

/**
 * everyItem() の戻り値を配列に正規化する（要素が1件のときスカラーで返るため）
 * @param {*} everyItemValue everyItem() で取得した値
 * @returns {array} 正規化した配列
 */
function toArray(everyItemValue) {
    return (everyItemValue instanceof Array) ? everyItemValue : [everyItemValue];
}

/**
 * 検索範囲に含まれる段落の文字列を取得する
 * @param {object} searchRange 検索対象（Document / Story）
 * @returns {array} 段落の文字列の配列
 */
function getParagraphTexts(searchRange) {
    /* Story はそのまま段落を持つ。Document はストーリーごとにたどる /
       A Story exposes paragraphs directly, a Document is walked story by story */
    var stories = (searchRange instanceof Document) ?
        searchRange.stories.everyItem().getElements() : [searchRange];

    var paragraphTexts = [];

    for (var i = 0; i < stories.length; i++) {
        if (!isSearchableStory(stories[i])) continue;

        paragraphTexts = paragraphTexts.concat(
            toArray(stories[i].paragraphs.everyItem().contents));
    }

    return paragraphTexts;
}

/**
 * 検索対象に含まれる段落の文字列をまとめて取得する
 * @param {array} searchTargets 検索対象の配列
 * @returns {array} 段落の文字列の配列
 */
function getParagraphTextsInTargets(searchTargets) {
    var paragraphTexts = [];

    for (var i = 0; i < searchTargets.length; i++) {
        paragraphTexts = paragraphTexts.concat(
            getParagraphTexts(searchTargets[i].searchRange));
    }

    return paragraphTexts;
}

/**
 * 検索文字列に該当する箇所を数える
 * @param {array} paragraphTexts 段落の文字列の配列
 * @param {string} searchText 検索文字列
 * @returns {number} 該当箇所の数
 * @description モーダルダイアログの表示中は findGrep() を使えないので、段落の文字列を直接数える。
 *              ExtendScript の正規表現には後読みが無いため、直前の1文字だけ自分で見る
 */
function countMatches(paragraphTexts, searchText) {
    var markerText = trimTrailingSeparators(searchText);

    if (markerText === "") return 0;

    var firstCharacter = markerText.charAt(0);
    var matchPattern = new RegExp(escapeGrepText(markerText) + buildRunGuard(markerText), "g");
    var matchCount = 0;

    for (var i = 0; i < paragraphTexts.length; i++) {
        var paragraphText = paragraphTexts[i];
        var matchResult;

        matchPattern.lastIndex = 0;

        while ((matchResult = matchPattern.exec(paragraphText)) !== null) {
            /* 同じ文字が続く並びの一部は数えない / Skip a match inside a longer run */
            if (matchResult.index === 0 ||
                paragraphText.charAt(matchResult.index - 1) !== firstCharacter) {
                matchCount++;
            }
        }
    }

    return matchCount;
}

/**
 * 検索オプションを既定値にそろえる
 * @param {object} findChangeOptions findChangeGrepOptions
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
 * 検索条件と検索オプションを初期化する
 * @returns {void}
 * @description 前回の「検索/置換」の設定を引き継がないよう、実行の前後でそろえ直す
 */
function resetFindPreferences() {
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    applyFindChangeOptions(app.findChangeGrepOptions);
}

/**
 * 段落の行頭にある目印を判定する
 * @param {string} paragraphText 段落の文字列
 * @returns {object} { key: string, markerText: string, searchText: string, labelKey: string }。
 *                   目印がない場合は null
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
            labelKey: NUMBERED_LIST_PATTERNS[i].labelKey
        };
    }

    var matchResult = LINE_HEAD_MARKER_PATTERN.exec(paragraphText);

    /* 区切りを含まない記号そのものを目印にする（"## " と "##" を同じものとして数える） /
       The marker is the symbol run itself, so "## " and "##" count as one */
    if (matchResult && !EXCLUDED_LINE_HEAD_PATTERN.test(matchResult[1])) {
        var lineHeadMarker = matchResult[1];

        return {
            key: lineHeadMarker,
            markerText: lineHeadMarker,
            searchText: buildLineHeadGrep(lineHeadMarker),
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
            labelKey: "marker.bold"
        };
    }

    return null;
}

/**
 * 段落の文字列を調べ、行頭に繰り返し現れる記号を多い順に集める
 * @param {array} paragraphTexts 段落の文字列の配列
 * @returns {array} { markerText: string, searchText: string, labelKey: string, count: number }
 *                  の配列
 */
function detectLineHeadMarkers(paragraphTexts) {
    var markerEntries = {};

    for (var i = 0; i < paragraphTexts.length; i++) {
        var markerEntry = matchLineHeadMarker(paragraphTexts[i]);
        if (!markerEntry) continue;

        if (markerEntries[markerEntry.key]) {
            markerEntries[markerEntry.key].count++;
        } else {
            markerEntry.count = 1;
            markerEntries[markerEntry.key] = markerEntry;
        }
    }

    var detectedMarkers = [];

    for (var markerKey in markerEntries) {
        if (!markerEntries.hasOwnProperty(markerKey)) continue;
        if (markerEntries[markerKey].count < MIN_REPEAT_COUNT) continue;

        detectedMarkers.push(markerEntries[markerKey]);
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
 * @returns {number} 処理した箇所数
 */
function applyStylesAndRemoveMarker(searchRange, paragraphStyle, characterStyle) {
    var foundTexts = searchRange.findGrep();
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
 * 検索文字列パネルを作る
 * @param {Window} dialog 追加先のダイアログ
 * @returns {object} { panel: Panel, modeRadios: array, searchTextField: EditText,
 *                     detectedMarkerDropdown: DropDownList, matchCountValue: StaticText }
 */
function addSearchPanel(dialog) {
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
    textModeRadio.helpTip = getLabel("tooltip.searchText");

    var autoModeRow = searchPanel.add("group");
    setupRow(autoModeRow);
    var autoModeRadio = autoModeRow.add("radiobutton", undefined, getLabel("mode.auto"));
    autoModeRadio.preferredSize.width = LABEL_WIDTH;
    autoModeRadio.helpTip = getLabel("tooltip.auto");
    var detectedMarkerDropdown = autoModeRow.add("dropdownlist", undefined, []);
    detectedMarkerDropdown.alignment = ["fill", "center"];
    detectedMarkerDropdown.helpTip = getLabel("tooltip.auto");

    var matchCountRow = searchPanel.add("group");
    setupRow(matchCountRow);
    addRowLabel(matchCountRow, "field.matchCount").helpTip = getLabel("tooltip.matchCount");
    var matchCountValue = matchCountRow.add("statictext", undefined, "0");
    matchCountValue.preferredSize.width = COUNT_WIDTH;
    matchCountValue.helpTip = getLabel("tooltip.matchCount");

    return {
        panel: searchPanel,
        modeRadios: [textModeRadio, autoModeRadio],
        searchTextField: searchTextField,
        detectedMarkerDropdown: detectedMarkerDropdown,
        matchCountValue: matchCountValue
    };
}

/**
 * スタイルパネルを作る
 * @param {Window} dialog 追加先のダイアログ
 * @param {Document} activeDoc アクティブドキュメント
 * @returns {object} { panel: Panel, paragraphStyleDropdown: DropDownList,
 *                     characterStyleDropdown: DropDownList }
 */
function addStylePanel(dialog, activeDoc) {
    var stylePanel = dialog.add("panel", undefined, getLabel("panel.style"));
    setupPanel(stylePanel);

    var paragraphStyleRow = stylePanel.add("group");
    setupRow(paragraphStyleRow);
    addRowLabel(paragraphStyleRow, "field.paragraphStyle").helpTip =
        getLabel("tooltip.paragraphStyle");
    var paragraphStyleDropdown = paragraphStyleRow.add(
        "dropdownlist", undefined, getStyleNames(activeDoc.allParagraphStyles, false));
    paragraphStyleDropdown.selection = 0;
    paragraphStyleDropdown.alignment = ["fill", "center"];
    paragraphStyleDropdown.helpTip = getLabel("tooltip.paragraphStyle");

    var characterStyleRow = stylePanel.add("group");
    setupRow(characterStyleRow);
    addRowLabel(characterStyleRow, "field.characterStyle").helpTip =
        getLabel("tooltip.characterStyle");
    var characterStyleDropdown = characterStyleRow.add(
        "dropdownlist", undefined, getStyleNames(activeDoc.allCharacterStyles, true));
    characterStyleDropdown.selection = 0;
    characterStyleDropdown.alignment = ["fill", "center"];
    characterStyleDropdown.helpTip = getLabel("tooltip.characterStyle");

    return {
        panel: stylePanel,
        paragraphStyleDropdown: paragraphStyleDropdown,
        characterStyleDropdown: characterStyleDropdown
    };
}

/**
 * 検索対象パネルを作る（縦並びのラジオボタン）
 * @param {Window} dialog 追加先のダイアログ
 * @returns {array} SCOPE_* の並び順に対応するラジオボタンの配列
 */
function addScopePanel(dialog) {
    var scopePanel = dialog.add("panel", undefined, getLabel("panel.scope"));
    setupPanel(scopePanel);
    scopePanel.spacing = RADIO_SPACING;

    var scopeLabelKeys = ["scope.allDocuments", "scope.document", "scope.story"];
    var searchScopeRadios = [];

    for (var i = 0; i < scopeLabelKeys.length; i++) {
        var scopeRadio = scopePanel.add("radiobutton", undefined, getLabel(scopeLabelKeys[i]));
        scopeRadio.helpTip = getLabel("tooltip." + scopeLabelKeys[i]);
        searchScopeRadios.push(scopeRadio);
    }

    searchScopeRadios[DEFAULT_SCOPE].value = true;

    return searchScopeRadios;
}

/**
 * Markdown パネルを作る
 * @param {Window} dialog 追加先のダイアログ
 * @returns {object} { matchLevelCheckbox: Checkbox, applyAllLevelsCheckbox: Checkbox }
 */
function addMarkdownPanel(dialog) {
    var markdownPanel = dialog.add("panel", undefined, getLabel("panel.markdown"));
    setupPanel(markdownPanel);

    var matchLevelRow = markdownPanel.add("group");
    setupRow(matchLevelRow);
    var matchLevelCheckbox =
        matchLevelRow.add("checkbox", undefined, getLabel("field.matchLevel"));
    matchLevelCheckbox.value = DEFAULT_MATCH_LEVEL;
    matchLevelCheckbox.helpTip = getLabel("tooltip.matchLevel");

    var applyAllLevelsRow = markdownPanel.add("group");
    setupRow(applyAllLevelsRow);
    var applyAllLevelsCheckbox =
        applyAllLevelsRow.add("checkbox", undefined, getLabel("field.applyAllLevels"));
    applyAllLevelsCheckbox.value = DEFAULT_APPLY_ALL_LEVELS;
    applyAllLevelsCheckbox.helpTip = getLabel("tooltip.applyAllLevels");

    return {
        matchLevelCheckbox: matchLevelCheckbox,
        applyAllLevelsCheckbox: applyAllLevelsCheckbox
    };
}

/**
 * ボタンエリアを作る
 * @param {Window} dialog 追加先のダイアログ
 * @returns {object} { btnOk: Button, btnCancel: Button }
 */
function addButtonRow(dialog) {
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

    return { btnOk: btnOk, btnCancel: btnCancel };
}

/**
 * ダイアログを表示して設定を取得する
 * @param {Document} activeDoc アクティブドキュメント
 * @returns {object} { searchText: string, applyAllLevels: boolean,
 *                     paragraphStyleName: string, characterStyleName: string,
 *                     searchScope: number }。キャンセル時は null
 * @description 「自動判別」の候補と対象箇所は、検索対象を変えるたびに取り直す
 */
function showDialog(activeDoc) {
    var dialog = new Window("dialog", getLabel("dialog.title"));
    setupWindow(dialog);

    var searchControls = addSearchPanel(dialog);
    var styleControls = addStylePanel(dialog, activeDoc);
    var searchScopeRadios = addScopePanel(dialog);
    var markdownControls = addMarkdownPanel(dialog);
    var dialogButtons = addButtonRow(dialog);

    var searchModeRadios = searchControls.modeRadios;
    var searchTextField = searchControls.searchTextField;
    var detectedMarkerDropdown = searchControls.detectedMarkerDropdown;
    var matchCountValue = searchControls.matchCountValue;
    var paragraphStyleDropdown = styleControls.paragraphStyleDropdown;
    var characterStyleDropdown = styleControls.characterStyleDropdown;
    var matchLevelCheckbox = markdownControls.matchLevelCheckbox;
    var applyAllLevelsCheckbox = markdownControls.applyAllLevelsCheckbox;
    var btnOk = dialogButtons.btnOk;

    var detectedMarkers = [];
    var detectedMarkerCache = {};
    var paragraphTextCache = {};

    /**
     * 現在選ばれている指定方法を取得する
     * @returns {number} SEARCH_MODE_* のいずれか
     */
    function getSelectedSearchMode() {
        return getSelectedRadioIndex(searchModeRadios);
    }

    /**
     * 現在の検索対象に含まれる段落の文字列を取得する
     * @returns {array} 段落の文字列の配列
     * @description 同じ検索対象を選び直したときは走査し直さない。
     *              自動判別と対象箇所の集計で同じ結果を使い回す
     */
    function getCurrentParagraphTexts() {
        var searchScope = getSelectedRadioIndex(searchScopeRadios);

        if (!paragraphTextCache[searchScope]) {
            paragraphTextCache[searchScope] = getParagraphTextsInTargets(
                getSearchTargets(searchScope, activeDoc));
        }

        return paragraphTextCache[searchScope];
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
     * 対象箇所の表示を更新する
     * @returns {void}
     * @description 自動判別は判別時に数えた段落数をそのまま出し、
     *              文字列を指定したときは入力のたびに数え直す
     */
    function updateMatchCount() {
        if (getSelectedSearchMode() === SEARCH_MODE_AUTO) {
            matchCountValue.text = detectedMarkerDropdown.selection ?
                detectedMarkers[detectedMarkerDropdown.selection.index].count + "" :
                "0";
            return;
        }

        matchCountValue.text =
            countMatches(getCurrentParagraphTexts(), searchTextField.text) + "";
    }

    /**
     * 見出しマーカーの選択に合わせて、Markdown パネルの状態を更新する
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
        searchControls.panel.enabled = !applyAllLevels;
        styleControls.panel.enabled = !applyAllLevels;
        matchLevelCheckbox.enabled = isHeadingMarker && !applyAllLevels;

        if (!matchLevelCheckbox.enabled || !matchLevelCheckbox.value) return;

        var styleIndex = findHeadingStyleIndex(paragraphStyleDropdown, headingLevel);
        if (styleIndex >= 0) paragraphStyleDropdown.selection = styleIndex;
    }

    /**
     * 目印の選び直しに合わせて、Markdown パネルと対象箇所を更新する
     * @returns {void}
     */
    function refreshMarkerSelection() {
        updateHeadingOptions();
        updateMatchCount();
    }

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

        refreshMarkerSelection();
    }

    /**
     * 検索対象を調べ直し、見つかった行頭記号をドロップダウンに展開する
     * @returns {void}
     */
    function refreshDetectedMarkers() {
        var searchScope = getSelectedRadioIndex(searchScopeRadios);

        /* 同じ検索対象を選び直したときは判別し直さない / Reuse the result for a scope already scanned */
        if (!detectedMarkerCache[searchScope]) {
            detectedMarkerCache[searchScope] =
                detectLineHeadMarkers(getCurrentParagraphTexts());
        }

        detectedMarkers = detectedMarkerCache[searchScope];

        var markerLabels = getDetectedMarkerLabels(detectedMarkers);

        detectedMarkerDropdown.removeAll();
        for (var i = 0; i < markerLabels.length; i++) {
            detectedMarkerDropdown.add("item", markerLabels[i]);
        }
        detectedMarkerDropdown.selection = (detectedMarkers.length > 0) ? 0 : null;

        searchModeRadios[SEARCH_MODE_AUTO].enabled = (detectedMarkers.length > 0);

        /* 候補が無くなったら文字列指定に戻す / Fall back to manual entry when nothing was detected */
        var searchMode = getSelectedSearchMode();
        if (detectedMarkers.length === 0 && searchMode === SEARCH_MODE_AUTO) {
            searchMode = SEARCH_MODE_TEXT;
        }
        selectSearchMode(searchMode);
    }

    searchModeRadios[SEARCH_MODE_TEXT].onClick = function () {
        selectSearchMode(SEARCH_MODE_TEXT);
    };
    searchModeRadios[SEARCH_MODE_AUTO].onClick = function () {
        selectSearchMode(SEARCH_MODE_AUTO);
    };
    searchTextField.onChanging = function () { selectSearchMode(getSelectedSearchMode()); };
    detectedMarkerDropdown.onChange = refreshMarkerSelection;
    matchLevelCheckbox.onClick = updateHeadingOptions;
    applyAllLevelsCheckbox.onClick = updateHeadingOptions;

    for (var i = 0; i < searchScopeRadios.length; i++) {
        searchScopeRadios[i].onClick = refreshDetectedMarkers;
    }

    selectSearchMode(DEFAULT_SEARCH_MODE);
    refreshDetectedMarkers();

    if (dialog.show() !== 1) return null;

    var selectedMarker = (getSelectedSearchMode() === SEARCH_MODE_AUTO) ?
        detectedMarkers[detectedMarkerDropdown.selection.index] :
        { searchText: buildExactGrep(searchTextField.text) };

    var applyAllLevels = applyAllLevelsCheckbox.enabled && applyAllLevelsCheckbox.value;

    /* ディム表示にしたパネルの設定は使わない / Settings in the dimmed panels are not used */
    var characterStyleName = (applyAllLevels || characterStyleDropdown.selection.index === 0) ?
        null :
        characterStyleDropdown.selection.text;

    return {
        searchText: selectedMarker.searchText,
        applyAllLevels: applyAllLevels,
        paragraphStyleName: paragraphStyleDropdown.selection.text,
        characterStyleName: characterStyleName,
        searchScope: getSelectedRadioIndex(searchScopeRadios)
    };
}

// =========================================
// 実行 / Run
// =========================================

/**
 * ひとつの検索を検索対象ぶん実行する
 * @param {object} searchPlan { searchText: string, headingLevel: number }
 * @param {array} searchTargets 検索対象の配列
 * @param {object} dialogSettings ダイアログの設定
 * @param {object} runResult 結果の記録先
 * @returns {void}
 */
function runSearchPlan(searchPlan, searchTargets, dialogSettings, runResult) {
    /* 検索条件を初期化 / Reset the find preferences */
    resetFindPreferences();
    app.findGrepPreferences.findWhat = searchPlan.searchText;

    for (var i = 0; i < searchTargets.length; i++) {
        var targetDoc = searchTargets[i].doc;
        var resolvedStyles = resolveStyles(
            targetDoc,
            dialogSettings.paragraphStyleName,
            dialogSettings.characterStyleName,
            searchPlan.headingLevel);

        /* 段落スタイルを持たないドキュメントはスキップ / Skip documents without the paragraph style */
        if (resolvedStyles.skip) {
            if (resolvedStyles.missingParagraphStyle) {
                addDocumentNote(runResult.skippedDocumentNotes,
                    targetDoc.name + ": " + resolvedStyles.missingParagraphStyle);
            }
            continue;
        }

        if (resolvedStyles.missingCharacterStyle) {
            addDocumentNote(runResult.partialDocumentNotes,
                targetDoc.name + ": " + resolvedStyles.missingCharacterStyle);
        }

        runResult.processedCount += applyStylesAndRemoveMarker(
            searchTargets[i].searchRange,
            resolvedStyles.paragraphStyle,
            resolvedStyles.characterStyle);
    }
}

/**
 * 検索の一覧を順に実行する
 * @param {array} searchPlans 検索の一覧
 * @param {array} searchTargets 検索対象の配列
 * @param {object} dialogSettings ダイアログの設定
 * @returns {object} { processedCount: number, skippedDocumentNotes: array,
 *                     partialDocumentNotes: array }
 */
function runSearchPlans(searchPlans, searchTargets, dialogSettings) {
    var runResult = {
        processedCount: 0,
        skippedDocumentNotes: [],
        partialDocumentNotes: []
    };

    for (var i = 0; i < searchPlans.length; i++) {
        runSearchPlan(searchPlans[i], searchTargets, dialogSettings, runResult);
    }

    return runResult;
}

/**
 * 完了メッセージを組み立てる
 * @param {object} runResult 実行結果
 * @returns {string} 表示するメッセージ
 */
function buildResultMessage(runResult) {
    var resultMessage = formatLabel(getLabel("result.processed"), [runResult.processedCount]);

    if (runResult.partialDocumentNotes.length > 0) {
        resultMessage += "\n\n" + getLabel("result.partial") +
            "\n" + runResult.partialDocumentNotes.join("\n");
    }

    if (runResult.skippedDocumentNotes.length > 0) {
        resultMessage += "\n\n" + getLabel("result.skipped") +
            "\n" + runResult.skippedDocumentNotes.join("\n");
    }

    return resultMessage;
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {
    if (app.documents.length === 0) {
        alert(getLabel("error.noDocument"));
        return;
    }

    var activeDoc = app.activeDocument;

    var dialogSettings = showDialog(activeDoc);
    if (!dialogSettings) return;

    var searchTargets = getSearchTargets(dialogSettings.searchScope, activeDoc);
    if (searchTargets.length === 0) {
        alert(getLabel("error.noTarget"));
        return;
    }

    var searchPlans = buildSearchPlans(dialogSettings);
    var runResult = null;

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(function () {
        runResult = runSearchPlans(searchPlans, searchTargets, dialogSettings);
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.apply"));

    /* 検索条件をクリア / Clear the find preferences */
    resetFindPreferences();

    alert(buildResultMessage(runResult));
})();
