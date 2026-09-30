#target indesign

/*

### 概要

指定した検索文字列を手がかりに段落スタイルを適用し、その文字列を削除します。検索対象はストーリー・ドキュメント・すべてのドキュメントから選べます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdRemoveMarkerApplyStyleSimple.md

### Overview

Applies a paragraph style to paragraphs containing the given text, then deletes that text. The search target can be the story, the document or all open documents.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdRemoveMarkerApplyStyleSimple.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdRemoveMarkerApplyStyleSimple"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";    /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-09";                     /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                     /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdRemoveMarkerApplyStyleSimple.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdRemoveMarkerApplyStyleSimple.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// 基本設定 / Settings
// =========================================

/* 検索文字列の初期値 / Default search text */
var DEFAULT_SEARCH_TEXT = "###";

/* 検索オプション（前回の「検索/置換」の設定を引き継がない） / Find options (never inherited) */
var FIND_WIDTH_SENSITIVE        = true;  /* 半角と全角を区別（＃ と # を分ける） / width sensitive */
var FIND_KANA_SENSITIVE         = true;  /* ひらがなとカタカナを区別 / kana sensitive */
var FIND_INCLUDE_FOOTNOTES      = false; /* 脚注を含む / include footnotes */
var FIND_INCLUDE_MASTER_PAGES   = false; /* マスターページを含む / include master pages */
var FIND_INCLUDE_HIDDEN_LAYERS  = false; /* 非表示レイヤーを含む / include hidden layers */
var FIND_INCLUDE_LOCKED_LAYERS  = false; /* ロックされたレイヤーを含む / include locked layers */
var FIND_INCLUDE_LOCKED_STORIES = false; /* ロックされたストーリーを含む / include locked stories */

/* 検索文字列に続く区切り（見つかれば一緒に削除する） / Separator after the marker (deleted together) */
/* 半角スペース・全角スペース・タブ。連続していてもまとめて削除し、無くてもよい /
   Spaces, full-width spaces and tabs: all of them are deleted, and none is fine too */
var TRAILING_SEPARATOR_PATTERN = "[ \u3000\\t]*";

/* スコープ（検索対象） / Search scope */
var SCOPE_ALL_DOCUMENTS = 0;
var SCOPE_DOCUMENT      = 1;
var SCOPE_STORY         = 2;

/* ダイアログの初期スコープ / Default scope */
var DEFAULT_SCOPE = SCOPE_STORY;

// =========================================
// レイアウト / Layout
// =========================================

/* ウィンドウの余白と間隔 / Window margins and spacing */
// UIレイアウト（再利用パーツ） / UI layout (reusable)

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

/**
 * ウィンドウの共通設定
 * @param {Window} targetWindow - 対象のウィンドウ
 * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
 * @returns {void}
 */
function setupWindow(targetWindow, spacing) {
    targetWindow.orientation = "column";
    targetWindow.alignChildren = "fill";
    targetWindow.margins = WINDOW_MARGINS;
    targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
 * @param {Panel} targetPanel - 対象のパネル
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
 * @returns {void}
 */
function setupPanel(targetPanel, spacing) {
    targetPanel.orientation = "column";
    targetPanel.alignChildren = ["fill", "top"];
    targetPanel.alignment = "fill";
    targetPanel.margins = PANEL_MARGINS;
    targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * タブの共通設定
 * @param {Tab} targetTab - 対象のタブ
 * @param {number} [spacing] - 要素間隔（省略時は変えない）
 * @returns {void}
 */
function setupTab(targetTab, spacing) {
    targetTab.orientation = "column";
    targetTab.alignChildren = "fill";
    targetTab.margins = TAB_MARGINS;
    if (typeof spacing === "number") targetTab.spacing = spacing;
}

/**
 * 横並びの行グループの共通設定（ボタン列など）。
 * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
 * @param {Group} rowGroup - 対象のグループ
 * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
 * @returns {void}
 */
function setupRow(rowGroup, rowAlignment, spacing) {
    rowGroup.orientation = "row";
    rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
    rowGroup.alignChildren = ["left", "center"];
    rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
 * @param {Button} targetButton - 対象のボタン
 * @param {number} trimPixels - 詰める量（px）
 * @returns {void}
 */
function trimButtonHeight(targetButton, trimPixels) {
    /* レイアウト前は size が無い / size is not set until the layout runs */
    if (!targetButton.size) return;
    targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
}

// UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

var ROW_SPACING    = 8;  /* 行の間隔 / row spacing */
var RADIO_SPACING  = 4;  /* ラジオボタンの間隔 / radio button spacing */

/* 行の寸法 / Row metrics */
var LABEL_WIDTH        = 108; /* ラベルの幅（全角1.5文字分の余裕を持たせる） / label width */
var SEARCH_TEXT_LENGTH = 24;  /* 検索文字列入力欄の文字数 / search field length */
var BUTTON_WIDTH       = 90;  /* ダイアログボタンの幅 / dialog button width */
var COUNT_WIDTH        = 60;  /* 対象箇所の数値の幅 / match count width */

/**
 * 列グループの共通設定を適用する
 * @param {Group} group 対象グループ
 * @param {string} horizontalAlignment 水平方向の配置（"fill" / "right" など）
 * @returns {void}
 */
function setupColumn(group, horizontalAlignment) {
    group.orientation = "column";
    group.alignment = [horizontalAlignment, "top"];
    group.alignChildren = ["fill", "top"];
    group.spacing = ROW_SPACING;
}

// =========================================
// ラベル定義 / Labels
// =========================================

// ローカライズ（再利用パーツ） / Localization (reusable)

/**
 * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
 * @returns {string} "ja" または "en"
 */
function getCurrentLang() {
    return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
}

var uiLang = getCurrentLang();

/**
 * LABELS から今の UI 言語の文言を取り出す。
 * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
 * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
 * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
 */
function getLabel(labelRef, placeholderValues) {
    var labelEntry = labelRef;
    if (typeof labelRef === "string") {
        var labelPathKeys = labelRef.split(".");
        labelEntry = LABELS;
        for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
            labelEntry = labelEntry[labelPathKeys[i]];
        }
    }
    var labelString;
    if (typeof labelEntry === "string") labelString = labelEntry;
    else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
    else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
    else return (typeof labelRef === "string") ? labelRef : "";
    return fillLabelPlaceholders(String(labelString), placeholderValues);
}

/**
 * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
 * @param {string|Object} labelRef - getLabel と同じ
 * @param {Object|Array} [placeholderValues] - getLabel と同じ
 * @returns {string} コロン付きの文言
 */
function labelText(labelRef, placeholderValues) {
    return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
}

/**
 * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
 * @param {string|Object} labelRef - getLabel と同じ
 * @param {string|number} value - コロンのあとに続ける値
 * @returns {string} 項目名と値をつないだ文字列
 */
function labelValueText(labelRef, value) {
    return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
}

/**
 * 文言の {name} や %1 に値を差し込む
 * @param {string} labelString - 文言
 * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
 * @returns {string} 差し込んだ文言
 */
function fillLabelPlaceholders(labelString, placeholderValues) {
    if (placeholderValues == null) return labelString;
    if (placeholderValues instanceof Array) {
        /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
        for (var i = placeholderValues.length; i >= 1; i--) {
            labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
        }
        return labelString;
    }
    for (var placeholderKey in placeholderValues) {
        if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
        labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
    }
    return labelString;
}

// ローカライズ（再利用パーツ）ここまで / End of the reusable localization

var LABELS = {
    dialog: {
        title: { ja: "検索文字列を削除して段落スタイルを適用", en: "Apply Paragraph Style and Delete Marker" }
    },
    field: {
        searchText:     { ja: "検索文字列", en: "Find What" },
        paragraphStyle: { ja: "置換スタイル", en: "Style to Apply" },
        scope:          { ja: "検索対象", en: "Search" },
        matchCount:     { ja: "対象箇所", en: "Matches" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" },
        ok:     { ja: "OK", en: "OK" }
    },
    tooltip: {
        searchText: {
            ja: "削除したい目印の文字列を入力します。例：###\n同じ文字が続く並びの一部には一致しません。「##」は「###」に一致しません。",
            en: "Enter the marker text to delete, for example ###.\nIt never matches part of a longer run of the same character, so ## does not match ###."
        },
        paragraphStyle: { ja: "検索文字列が見つかった段落に適用する段落スタイルです。", en: "The paragraph style applied to paragraphs that contain the search text." },
        scope:          { ja: "「ストーリー」は、テキストの選択またはカーソル位置が必要です。", en: "Story needs a text selection or the cursor placed in a story." },
        cancel:         { ja: "何も変更せずに閉じます。", en: "Close without making any changes." },
        ok:             { ja: "検索文字列を削除して段落スタイルを適用します。", en: "Delete the search text and apply the paragraph style." }
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
        skipped:   { ja: "段落スタイルが見つからないため、次のドキュメントはスキップしました。", en: "Skipped these documents because the paragraph style was not found." }
    },
    undo: {
        apply: { ja: "検索文字列を削除して段落スタイルを適用", en: "Apply Paragraph Style and Delete Marker" }
    }
};

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
 * 行ラベルを追加する（幅を固定して右揃え）
 * @param {Group} parentGroup 追加先の行グループ
 * @param {string} labelKey ラベルキー
 * @returns {StaticText} 追加したラベル
 */
function addRowLabel(parentGroup, labelKey) {
    var rowLabel = parentGroup.add("statictext", undefined, labelText(labelKey));
    rowLabel.preferredSize.width = LABEL_WIDTH;
    rowLabel.justify = "right";
    return rowLabel;
}

/**
 * ラベル付きの行を追加する（ラベルは addRowLabel() で幅を固定して右揃え）
 * @param {Group} parentGroup 追加先のグループ
 * @param {string} labelKey ラベルキー
 * @returns {Group} 追加した行グループ
 */
function addLabeledRow(parentGroup, labelKey) {
    var labeledRow = parentGroup.add("group");
    setupRow(labeledRow, "fill", ROW_SPACING);

    addRowLabel(labeledRow, labelKey);

    return labeledRow;
}

/**
 * 行に縦並びのラジオボタンを追加する
 * @param {Group} parentRow 追加先の行グループ
 * @param {array} labelKeys 項目のラベルキーの配列
 * @param {number} selectedIndex 初期選択のインデックス
 * @param {string} tooltipKey ツールチップのラベルキー
 * @returns {array} 追加したラジオボタンの配列
 */
function addRadioColumn(parentRow, labelKeys, selectedIndex, tooltipKey) {
    /* ラベルは1行目のラジオボタンに合わせて上端にそろえる / The label sits at the top */
    parentRow.alignChildren = ["left", "top"];

    var radioGroup = parentRow.add("group");
    setupColumn(radioGroup, "fill");
    radioGroup.spacing = RADIO_SPACING;

    var radioButtons = [];

    for (var i = 0; i < labelKeys.length; i++) {
        var radioButton = radioGroup.add("radiobutton", undefined, getLabel(labelKeys[i]));
        radioButton.helpTip = getLabel(tooltipKey);
        radioButtons.push(radioButton);
    }

    radioButtons[selectedIndex].value = true;

    return radioButtons;
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
 * 検索条件と検索オプションを初期化する
 * @returns {void}
 * @description 前回の「検索/置換」の設定を引き継がないよう、実行の前後でそろえ直す
 */
function resetFindPreferences() {
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    var findChangeOptions = app.findChangeGrepOptions;
    findChangeOptions.widthSensitive = FIND_WIDTH_SENSITIVE;
    findChangeOptions.kanaSensitive = FIND_KANA_SENSITIVE;
    findChangeOptions.includeFootnotes = FIND_INCLUDE_FOOTNOTES;
    findChangeOptions.includeMasterPages = FIND_INCLUDE_MASTER_PAGES;
    findChangeOptions.includeHiddenLayers = FIND_INCLUDE_HIDDEN_LAYERS;
    findChangeOptions.includeLockedLayersForFind = FIND_INCLUDE_LOCKED_LAYERS;
    findChangeOptions.includeLockedStoriesForFind = FIND_INCLUDE_LOCKED_STORIES;
}

/**
 * 初期選択にする検索対象を決める
 * @param {Document} activeDoc アクティブドキュメント
 * @returns {number} SCOPE_* のいずれか
 * @description ストーリーはテキスト選択かカーソル位置が必要なので、無いときはドキュメントにする
 */
function getInitialScope(activeDoc) {
    if (DEFAULT_SCOPE !== SCOPE_STORY) return DEFAULT_SCOPE;

    return (getSearchTargets(SCOPE_STORY, activeDoc).length > 0) ? SCOPE_STORY : SCOPE_DOCUMENT;
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
 * マスタースプレッド上のストーリーかどうかを判定する
 * @param {Story} story 対象のストーリー
 * @returns {boolean} マスタースプレッド上なら true
 */
function isMasterSpreadStory(story) {
    var textContainers = story.textContainers;

    if (!textContainers || textContainers.length === 0) return false;

    return isOnMasterSpread(textContainers[0]);
}

/**
 * マスタースプレッド上のオブジェクトかどうかを判定する
 * @param {object} pageItem 対象のページアイテム
 * @returns {boolean} マスタースプレッド上なら true
 * @description グループの中に置かれていることもあるので、親をたどって判定する
 */
function isOnMasterSpread(pageItem) {
    var ancestor = pageItem;

    while (ancestor && ancestor.parent) {
        ancestor = ancestor.parent;

        if (ancestor instanceof MasterSpread) return true;
        if (ancestor instanceof Spread || ancestor instanceof Document) return false;
    }

    return false;
}

/**
 * 検索範囲に含まれる段落の文字列を取得する
 * @param {object} searchRange 検索対象（Document / Story）
 * @returns {array} 段落の文字列の配列
 */
function getParagraphTexts(searchRange) {
    /* Story はそのまま段落を持つ。Document はストーリーごとにたどる /
       A Story exposes paragraphs directly, a Document is walked story by story */
    var isDocumentRange = (searchRange instanceof Document);
    var stories = isDocumentRange ?
        searchRange.stories.everyItem().getElements() : [searchRange];

    var paragraphTexts = [];

    for (var i = 0; i < stories.length; i++) {
        /* ドキュメント全体の検索は検索オプションに合わせてマスターページを除く。
           ストーリー指定でマスターページを選んだときは、そのまま数える /
           A document-wide search follows the find option; a story picked by hand is always counted */
        if (isDocumentRange && !FIND_INCLUDE_MASTER_PAGES &&
            isMasterSpreadStory(stories[i])) continue;

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
 * 段落の文字列を照合できる形にそろえる
 * @param {string} paragraphText 段落の文字列
 * @returns {string} 末尾の改行を外した文字列
 * @description 表を含む段落は contents が配列で返る。末尾の改行を外して "$" を段落の終わりに合わせる
 */
function toPlainParagraphText(paragraphText) {
    return String(paragraphText).replace(/[\r\n]+$/, "");
}

/**
 * 検索文字列（目印）に該当する箇所を数える
 * @param {array} paragraphTexts 段落の文字列の配列
 * @param {string} searchText 検索文字列
 * @returns {number} 該当箇所の数
 * @description モーダルダイアログの表示中は findGrep() を使えないので、段落の文字列を直接数える。
 *              ExtendScript の正規表現には後読みが無いため、直前の1文字だけ自分で見る。
 *              末尾の区切りは buildExactGrep() と同じく落として数える
 */
function countMarkerMatches(paragraphTexts, searchText) {
    var markerText = trimTrailingSeparators(searchText);

    if (markerText === "") return 0;

    var firstCharacter = markerText.charAt(0);
    var matchPattern = new RegExp(escapeGrepText(markerText) + buildRunGuard(markerText), "g");
    var matchCount = 0;

    for (var i = 0; i < paragraphTexts.length; i++) {
        var paragraphText = toPlainParagraphText(paragraphTexts[i]);
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
 * 検索範囲内の検索文字列に段落スタイルを適用して削除する
 * @param {object} searchRange 検索対象（Document / Story）
 * @param {ParagraphStyle} paragraphStyle 適用する段落スタイル
 * @returns {number} 処理した箇所数
 */
function applyStyleAndRemoveMarker(searchRange, paragraphStyle) {
    var foundTexts = searchRange.findGrep();
    var appliedCount = 0;

    /* 後ろから処理 / Process from the end */
    for (var i = foundTexts.length - 1; i >= 0; i--) {
        try {
            var foundText = foundTexts[i];

            /* 検索文字列を含む段落にスタイルを適用 / Apply the style to the paragraph */
            foundText.paragraphs[0].appliedParagraphStyle = paragraphStyle;

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
 * @returns {object} { searchText: string, paragraphStyleName: string, searchScope: number }。
 *                   キャンセル時は null
 */
function showDialog(activeDoc) {
    var dialog = new Window("dialog", getLabel("dialog.title"));
    setupWindow(dialog);

    var mainGroup = dialog.add("group");
    setupRow(mainGroup, "fill", COLUMN_SPACING);
    mainGroup.alignChildren = ["fill", "top"];

    /* 入力欄 / Fields */
    var fieldGroup = mainGroup.add("group");
    setupColumn(fieldGroup, "fill");

    var searchTextRow = addLabeledRow(fieldGroup, "field.searchText");
    var searchTextField = searchTextRow.add("edittext", undefined, DEFAULT_SEARCH_TEXT);
    searchTextField.characters = SEARCH_TEXT_LENGTH;
    searchTextField.alignment = ["fill", "center"];
    searchTextField.helpTip = getLabel("tooltip.searchText");

    var paragraphStyleRow = addLabeledRow(fieldGroup, "field.paragraphStyle");
    var paragraphStyleDropdown = paragraphStyleRow.add(
        "dropdownlist", undefined, getStyleNames(activeDoc.allParagraphStyles));
    paragraphStyleDropdown.selection = 0;
    paragraphStyleDropdown.alignment = ["fill", "center"];
    paragraphStyleDropdown.helpTip = getLabel("tooltip.paragraphStyle");

    var searchScopeRadios = addRadioColumn(
        addLabeledRow(fieldGroup, "field.scope"),
        ["scope.allDocuments", "scope.document", "scope.story"],
        getInitialScope(activeDoc), "tooltip.scope");

    /* ボタンエリア / Button column */
    var btnColumnGroup = mainGroup.add("group");
    setupColumn(btnColumnGroup, "right");

    var btnOK = btnColumnGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    var btnCancel = btnColumnGroup.add(
        "button", undefined, getLabel("button.cancel"), { name: "cancel" });
    btnOK.preferredSize.width = BUTTON_WIDTH;
    btnCancel.preferredSize.width = BUTTON_WIDTH;
    btnOK.helpTip = getLabel("tooltip.ok");
    btnCancel.helpTip = getLabel("tooltip.cancel");

    /* 対象箇所（ほかの行と同じラベル付きの行。検索文字列・検索対象を変えるたびに数え直す） /
       Match count, shown as a labeled row and refreshed on every change */
    var matchCountRow = addLabeledRow(dialog, "field.matchCount");
    var matchCountValue = matchCountRow.add("statictext", undefined, "0");
    matchCountValue.preferredSize.width = COUNT_WIDTH;

    var paragraphTextCache = {};

    /**
     * 現在の検索対象に含まれる段落の文字列を取得する
     * @returns {array} 段落の文字列の配列
     * @description 同じ検索対象を選び直したときや、文字を打つたびには走査し直さない
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
     * 該当箇所を数え直して表示し、実行できるかどうかを切り替える
     * @returns {void}
     * @description 該当箇所が無いまま実行すると、区切りだけの検索文字列（スペースなど）が
     *              検索対象すべてのスペースを削除してしまうので、0件では実行させない
     */
    function updateMatchCount() {
        var matchCount = countMarkerMatches(getCurrentParagraphTexts(), searchTextField.text);

        matchCountValue.text = matchCount + "";
        btnOK.enabled = (matchCount > 0);
    }

    searchTextField.onChanging = updateMatchCount;
    updateMatchCount();

    for (var i = 0; i < searchScopeRadios.length; i++) {
        searchScopeRadios[i].onClick = updateMatchCount;
    }

    searchTextField.active = true;

    if (dialog.show() !== 1) return null;

    return {
        searchText: searchTextField.text,
        paragraphStyleName: paragraphStyleDropdown.selection.text,
        searchScope: getSelectedRadioIndex(searchScopeRadios)
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

    var processedCount = 0;
    var skippedDocuments = [];

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(function () {
        /* 検索条件を初期化 / Reset the find preferences */
        resetFindPreferences();
        app.findGrepPreferences.findWhat = buildExactGrep(dialogSettings.searchText);

        for (var i = 0; i < searchTargets.length; i++) {
            var paragraphStyle = findStyleByName(
                searchTargets[i].doc.allParagraphStyles, dialogSettings.paragraphStyleName);

            /* 段落スタイルを持たないドキュメントはスキップ / Skip documents without the paragraph style */
            if (!paragraphStyle) {
                skippedDocuments.push(
                    searchTargets[i].doc.name + ": " + dialogSettings.paragraphStyleName);
                continue;
            }

            processedCount += applyStyleAndRemoveMarker(
                searchTargets[i].searchRange, paragraphStyle);
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.apply"));

    /* 検索条件をクリア / Clear the find preferences */
    resetFindPreferences();

    var resultMessage = getLabel("result.processed", [processedCount]);

    if (skippedDocuments.length > 0) {
        resultMessage += "\n\n" + getLabel("result.skipped") +
            "\n" + skippedDocuments.join("\n");
    }

    alert(resultMessage);
})();
