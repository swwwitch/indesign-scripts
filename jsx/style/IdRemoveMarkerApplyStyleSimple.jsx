#target indesign

/*

### 概要

指定した検索文字列を手がかりに段落スタイルを適用し、その文字列を削除します。検索対象はストーリー・ドキュメント・すべてのドキュメントから選べます。

詳細は README を参照してください。

### Overview

Applies a paragraph style to paragraphs containing the given text, then deletes that text. The search target can be the story, the document or all open documents.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdRemoveMarkerApplyStyleSimple"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                           /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";    /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-09";                     /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-10";                     /* 更新日 / last updated */

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
var FIND_INCLUDE_FOOTNOTES      = true;  /* 脚注を含む / include footnotes */
var FIND_INCLUDE_MASTER_PAGES   = true;  /* マスターページを含む / include master pages */
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
var WINDOW_MARGINS = 16; /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12; /* ウィンドウ内の要素間隔 / window spacing */
var COLUMN_SPACING = 16; /* 入力欄とボタン列の間隔 / gap between the fields and the buttons */
var ROW_SPACING    = 8;  /* 行の間隔 / row spacing */
var RADIO_SPACING  = 4;  /* ラジオボタンの間隔 / radio button spacing */

/* 行の寸法 / Row metrics */
var LABEL_WIDTH        = 108; /* ラベルの幅（全角1.5文字分の余裕を持たせる） / label width */
var SEARCH_TEXT_LENGTH = 24;  /* 検索文字列入力欄の文字数 / search field length */
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
 * 行グループの共通設定を適用する
 * @param {Group} group 対象グループ
 * @returns {void}
 */
function setupRow(group) {
    group.orientation = "row";
    group.alignment = ["fill", "top"];
    group.alignChildren = ["left", "center"];
    group.spacing = ROW_SPACING;
}

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

// =========================================
// 共通処理 / Helpers
// =========================================

/**
 * GREP 検索用に正規表現の特殊文字をエスケープする
 * @param {string} plainText エスケープする文字列
 * @returns {string} エスケープした文字列
 */
function escapeGrepText(plainText) {
    return plainText.replace(/([\\^$.|?*+()\[\]{}])/g, "\\$1");
}

/**
 * 検索文字列にぴったり一致する GREP パターンを組み立てる
 * @param {string} searchText 検索文字列
 * @returns {string} GREP パターン
 * @description 前後を先読み・後読みで止めることで、同じ文字が続く並びの一部には一致させない。
 *              これが無いと "##" が "### 見出し" の先頭2文字に一致してしまう。
 *              続く区切り（スペース・タブ）は一致範囲に含めて、目印と一緒に削除する
 */
function buildExactGrep(searchText) {
    var firstCharacter = searchText.charAt(0);
    var lastCharacter = searchText.charAt(searchText.length - 1);

    return "(?<!" + escapeGrepText(firstCharacter) + ")" +
        escapeGrepText(searchText) +
        "(?!" + escapeGrepText(lastCharacter) + ")" +
        TRAILING_SEPARATOR_PATTERN;
}

/**
 * 名前で段落スタイルを探す（スタイルグループ内も対象）
 * @param {array} styleCollection allParagraphStyles
 * @param {string} styleName 探すスタイル名
 * @returns {ParagraphStyle} 見つかったスタイル。無い場合は null
 */
function findStyleByName(styleCollection, styleName) {
    for (var i = 0; i < styleCollection.length; i++) {
        if (styleCollection[i].name === styleName) return styleCollection[i];
    }

    return null;
}

/**
 * スタイル名の一覧を取得する
 * @param {array} styleCollection allParagraphStyles
 * @returns {array} スタイル名の配列
 */
function getStyleNames(styleCollection) {
    var styleNames = [];

    for (var i = 0; i < styleCollection.length; i++) {
        styleNames.push(styleCollection[i].name);
    }

    return styleNames;
}

/**
 * ラベル付きの行を追加する（ラベルは幅を固定して右揃え）
 * @param {Group} parentGroup 追加先のグループ
 * @param {string} labelKey ラベルキー
 * @returns {Group} 追加した行グループ
 */
function addLabeledRow(parentGroup, labelKey) {
    var labeledRow = parentGroup.add("group");
    setupRow(labeledRow);

    var rowLabel = labeledRow.add("statictext", undefined, getLabelWithColon(labelKey));
    rowLabel.preferredSize.width = LABEL_WIDTH;
    rowLabel.justify = "right";

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
    try {
        searchTargets.push({ doc: activeDoc, searchRange: app.selection[0].parentStory });
    } catch (e) {
        /* 選択が無い、またはテキスト以外が選択されている / No selection, or it is not text */
    }

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
 * @param {*} value everyItem() で取得した値
 * @returns {array} 正規化した配列
 */
function toArray(value) {
    return (value instanceof Array) ? value : [value];
}

/**
 * ストーリーが検索対象になるかを判定する
 * @param {Story} story 対象のストーリー
 * @returns {boolean} 検索対象なら true
 * @description 検索オプションに合わせ、非表示レイヤー・ロックされたレイヤーだけに
 *              置かれているストーリーは対象外にする。フレームを持たないストーリーは対象に含める
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
 * 検索範囲に含まれる段落の文字列を取得する
 * @param {object} searchRange 検索対象（Document / Story）
 * @returns {array} 段落の文字列の配列
 */
function getParagraphTexts(searchRange) {
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
 * 検索対象に含まれる該当箇所を数える
 * @param {array} searchTargets 検索対象の配列
 * @param {string} searchText 検索文字列
 * @returns {number} 該当箇所の数
 * @description モーダルダイアログの表示中は findGrep() を使えないので、段落の文字列を直接数える。
 *              ExtendScript の正規表現には後読みが無いため、直前の1文字だけ自分で見る
 */
function countMatches(searchTargets, searchText) {
    if (searchText === "") return 0;

    var firstCharacter = searchText.charAt(0);
    var lastCharacter = searchText.charAt(searchText.length - 1);
    var matchPattern = new RegExp(
        escapeGrepText(searchText) + "(?!" + escapeGrepText(lastCharacter) + ")", "g");

    var matchCount = 0;

    for (var i = 0; i < searchTargets.length; i++) {
        var paragraphTexts = getParagraphTexts(searchTargets[i].searchRange);

        for (var j = 0; j < paragraphTexts.length; j++) {
            var paragraphText = paragraphTexts[j];
            var matched;

            matchPattern.lastIndex = 0;

            while ((matched = matchPattern.exec(paragraphText)) !== null) {
                /* 同じ文字が続く並びの一部は数えない / Skip a match inside a longer run */
                if (matched.index === 0 ||
                    paragraphText.charAt(matched.index - 1) !== firstCharacter) {
                    matchCount++;
                }
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
    setupRow(mainGroup);
    mainGroup.alignChildren = ["fill", "top"];
    mainGroup.spacing = COLUMN_SPACING;

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

    var btnOk = btnColumnGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    var btnCancel = btnColumnGroup.add(
        "button", undefined, getLabel("button.cancel"), { name: "cancel" });
    btnOk.preferredSize.width = BUTTON_WIDTH;
    btnCancel.preferredSize.width = BUTTON_WIDTH;
    btnOk.helpTip = getLabel("tooltip.ok");
    btnCancel.helpTip = getLabel("tooltip.cancel");

    /* 対象箇所（検索文字列・検索対象を変えるたびに数え直す） / Match count, refreshed on every change */
    /* 対象箇所（ほかの行と同じラベル付きの行にそろえる） / Match count, shown as a labeled row */
    var matchCountRow = addLabeledRow(dialog, "field.matchCount");
    var matchCountValue = matchCountRow.add("statictext", undefined, "0");
    matchCountValue.preferredSize.width = COUNT_WIDTH;

    /**
     * 該当箇所を数え直して表示する
     * @returns {void}
     */
    function updateMatchCount() {
        var searchTargets = getSearchTargets(
            getSelectedRadioIndex(searchScopeRadios), activeDoc);

        matchCountValue.text = countMatches(searchTargets, searchTextField.text) + "";
    }

    /* 検索文字列が空のままでは実行できない / An empty search text cannot be run */
    searchTextField.onChanging = function () {
        btnOk.enabled = (searchTextField.text !== "");
        updateMatchCount();
    };
    searchTextField.onChanging();

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

    var resultMessage = getLabel("result.processed").replace("%1", processedCount);

    if (skippedDocuments.length > 0) {
        resultMessage += "\n\n" + getLabel("result.skipped") +
            "\n" + skippedDocuments.join("\n");
    }

    alert(resultMessage);
})();
