#target indesign

/*

### 概要

テキスト・表・オブジェクトに適用されているスタイルのオーバーライドをまとめて消去します。
処理する範囲はドキュメント全体・ストーリー・選択範囲から選べ、特定のスタイルが適用された箇所だけに絞り込むこともできます。

詳細は README を参照してください。

### Overview

Clears style overrides from text, tables and objects in one pass.
The scope can be the whole document, a story or the current selection, and processing can be narrowed to where a particular style is applied.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdClearStyleOverrides";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Gregor Fellenz (grefel)";      /* 作者 / author */
var SCRIPT_MODIFIED = "Masahiro Takano (@swwwitch)";  /* 改変 / modified by */
var SCRIPT_RELEASED = "2020-06-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-20";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdClearStyleOverrides.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdClearStyleOverrides.md"; /* README (English) */

// Copyright (C) 2020 Gregor Fellenz <http://www.publishingx.de>
// Copyright (C) 2026 Masahiro Takano (@swwwitch)
//
// Modification notice (GNU GPL v3 section 5a) / 改変の告知（GNU GPL v3 第5条a）:
//   2020-06-09  Gregor Fellenz — original work "Clear Overrides"
//               https://github.com/grefel/clearOverrides
//   2026-09-20  Masahiro Takano — 全面的な書き直し。日本語／英語のローカライズ、
//               UIの再構成、セルスタイルでの絞り込みを追加。
//               Rewritten throughout: Japanese/English localization,
//               reorganized UI, added filtering by cell style.
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program. The full text is distributed alongside this
// script as IdClearStyleOverrides-LICENSE.txt. If it is missing, see
// <https://www.gnu.org/licenses/>.

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* ダイアログを開いたときのチェック状態 / Checkbox states on launch */
var DEFAULT_PROCESS_TEXT    = true;
var DEFAULT_PROCESS_TABLES  = false;
var DEFAULT_PROCESS_OBJECTS = false;

/* 「消去する種類」の初期選択 / Preselected override type (0:両方 1:段落のみ 2:文字のみ) */
var DEFAULT_OVERRIDE_TYPE_INDEX = 0;

/* 「対象」の初期選択 / Preselected scope (0:ドキュメント 1:ストーリー 2:選択範囲) */
var DEFAULT_SCOPE_INDEX = 0;


// =========================================
// ラベル定義 / Labels
// =========================================

/**
 * UI 言語を判定する
 * @returns {string} "ja" または "en"
 */
function getUiLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var uiLang = getUiLang();

var LABELS = {
    dialog: {
        title: { ja: "スタイルのオーバーライドを消去", en: "Clear style overrides" }
    },
    panel: {
        scope:  { ja: "対象", en: "Scope" },
        text:   { ja: "テキスト（段落スタイルと文字スタイル）", en: "Text (paragraph & character styles)" },
        table:  { ja: "表スタイルとセルスタイル", en: "Table & cell styles" },
        object: { ja: "オブジェクトスタイル", en: "Object styles" }
    },
    checkbox: {
        text:   { ja: "テキストのオーバーライドを消去", en: "Clear text overrides" },
        table:  { ja: "表のオーバーライドを消去", en: "Clear table overrides" },
        object: { ja: "オブジェクトのオーバーライドを消去", en: "Clear object overrides" }
    },
    fieldLabel: {
        paragraphStyle: { ja: "対象の段落スタイル", en: "Target paragraph style" },
        characterStyle: { ja: "対象の文字スタイル", en: "Target character style" },
        tableStyle:     { ja: "対象の表スタイル", en: "Target table style" },
        cellStyle:      { ja: "対象のセルスタイル", en: "Target cell style" },
        objectStyle:    { ja: "対象のオブジェクトスタイル", en: "Target object style" },
        overrideType:   { ja: "消去するオーバーライド", en: "Overrides to clear" }
    },
    radio: {
        scopeDocument:  { ja: "ドキュメント", en: "Document" },
        scopeStory:     { ja: "ストーリー", en: "Story" },
        scopeSelection: { ja: "選択範囲", en: "Selection" },
        bothTypes:     { ja: "両方", en: "Both" },
        paragraphOnly: { ja: "段落のみ", en: "Paragraph only" },
        characterOnly: { ja: "文字のみ", en: "Character only" }
    },
    dropdown: {
        allStyles: { ja: "すべてのスタイル", en: "All styles" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" },
        run:    { ja: "オーバーライドを消去", en: "Clear overrides" }
    },
    tooltip: {
        clearText: {
            ja: "本文・表の中のテキスト・脚注に付いた、段落スタイルと文字スタイルのオーバーライドを消去します。",
            en: "Clears paragraph and character style overrides on story text, text inside tables, and footnotes."
        },
        clearTable: {
            ja: "表スタイルのオーバーライドと、表の中のセルスタイルのオーバーライドを消去します。",
            en: "Clears table style overrides and the cell style overrides inside the table."
        },
        clearObject: {
            ja: "オブジェクトスタイルのオーバーライドを消去します。対象が「ドキュメント」のときは、マスターページ上のフレームも含みます。",
            en: "Clears object style overrides. When the scope is \"Document\", frames on parent (master) pages are included too."
        },
        scope: {
            ja: "「ストーリー」は選択箇所を含むストーリー全体（連結したテキストフレームすべて）、「選択範囲」は選択しているテキストとオブジェクトだけが対象になります。",
            en: "\"Story\" targets the whole story containing the selection, including every threaded frame. \"Selection\" targets only the selected text and objects."
        },
        overrideType: {
            ja: "「段落のみ」は行揃えや段落前後のアキなど段落レベルの手動設定だけを、「文字のみ」はフォントやサイズなど文字レベルの手動設定だけを消去します。",
            en: "\"Paragraph only\" clears paragraph-level local formatting such as alignment and space before/after; \"Character only\" clears character-level local formatting such as font and size."
        },
        styleFilter: {
            ja: "「すべてのスタイル」のままにすると、すべてのスタイルが対象になります。特定のスタイルを選ぶと、そのスタイルが適用された箇所だけに絞り込まれます。",
            en: "Leave \"All styles\" to target every style. Pick a style to narrow the scope to where that style is applied."
        }
    },
    alert: {
        noDocument:  { ja: "ドキュメントを開いてから実行してください。", en: "Please run with an active document." },
        noSelection: { ja: "対象が選択されていません。テキストやオブジェクトを選択してから実行してください。", en: "Nothing is selected. Select text or objects before running." },
        failed:      { ja: "オーバーライドの消去中にエラーが発生しました。", en: "An error occurred while clearing overrides." }
    },
    undo: {
        clearOverrides: { ja: "オーバーライドを消去", en: "Clear style overrides" }
    }
};

/**
 * ドット区切りキーでラベルを取得する
 * @param {string} labelKey 例: "dialog.title"
 * @returns {string} 現在の言語のラベル文字列。見つからない場合はキーをそのまま返す
 */
function getLabel(labelKey) {
    var node = LABELS;
    var keyParts = labelKey.split(".");
    for (var i = 0; i < keyParts.length; i++) {
        node = node[keyParts[i]];
        if (!node) return labelKey;
    }
    return node[uiLang] || node.en || labelKey;
}

/**
 * コロン付きの項目名を返す（日本語は全角、英語は半角）
 * @param {string} labelKey ラベルキー
 * @returns {string} コロンを付けた項目名
 */
function labelText(labelKey) {
    return getLabel(labelKey) + (uiLang === "ja" ? "：" : ":");
}

// =========================================
// レイアウト / Layout
// =========================================

var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */
var LABEL_WIDTH    = 200;                /* 項目名の幅 / field label width */
var DROPDOWN_WIDTH = 220;                /* ドロップダウンの幅 / dropdown width */
var BUTTON_ROW_TOP_MARGIN = 4;           /* ボタン行の上余白 / button row top margin */
var BUTTON_SPACING = 8;                  /* ボタン間の間隔 / gap between buttons */
var RADIO_SPACING  = 12;                 /* ラジオボタン間の間隔 / gap between radio buttons */

/**
 * ウィンドウの共通設定を適用する
 * @param {Window} targetWindow 対象ウィンドウ
 * @returns {void}
 */
function setupWindow(targetWindow) {
    targetWindow.orientation = "column";
    targetWindow.alignChildren = "fill";
    targetWindow.margins = WINDOW_MARGINS;
    targetWindow.spacing = WINDOW_SPACING;
}

/**
 * パネルの共通設定を適用する
 * @param {Panel} targetPanel 対象パネル
 * @returns {void}
 */
function setupPanel(targetPanel) {
    targetPanel.orientation = "column";
    targetPanel.alignChildren = ["fill", "top"];
    targetPanel.alignment = "fill";
    targetPanel.margins = PANEL_MARGINS;
    targetPanel.spacing = PANEL_SPACING;
}

/**
 * 行グループの共通設定を適用する
 * @param {Group} targetGroup 対象グループ
 * @param {string} [alignment] 横方向の配置。省略時は "left"
 * @returns {void}
 */
function setupRow(targetGroup, alignment) {
    targetGroup.orientation = "row";
    targetGroup.alignment = [alignment || "left", "center"];
    targetGroup.alignChildren = ["left", "center"];
    targetGroup.spacing = PANEL_SPACING;
}

/**
 * 右揃えの項目名を行に追加する
 * @param {Group} parentGroup 追加先の行グループ
 * @param {string} labelKey ラベルキー
 * @returns {StaticText} 追加した項目名
 */
function addRowLabel(parentGroup, labelKey) {
    var rowLabel = parentGroup.add("statictext", undefined, labelText(labelKey));
    rowLabel.preferredSize.width = LABEL_WIDTH;
    rowLabel.justify = "right";
    return rowLabel;
}

// =========================================
// 検索 / Find
// =========================================

var ALL_TEXT_GREP   = "(?s).+";   /* 改行を含むすべてのテキストにマッチ / Match every text run, newlines included */
var TABLE_MARKER    = "<0016>";   /* 表のアンカー文字 U+0016 / Table anchor character U+0016 */
var ALL_STYLES      = null;       /* スタイルで絞り込まないことを表す値 / Sentinel meaning "do not filter by style" */

/* 検索の種類ごとの app プロパティ名 / app property names per search kind */
var SEARCH_MODES = {
    grep: {
        optionsKey:   "findChangeGrepOptions",
        findPrefsKey: "findGrepPreferences",
        findMethod:   "findGrep"
    },
    text: {
        optionsKey:   "findChangeTextOptions",
        findPrefsKey: "findTextPreferences",
        findMethod:   "findText"
    },
    object: {
        optionsKey:   "findChangeObjectOptions",
        findPrefsKey: "findObjectPreferences",
        findMethod:   "findObject",
        objectType:   ObjectTypes.ALL_FRAMES_TYPE
    }
};

/* 退避・復帰の対象になる検索オプション / Search options that are saved and restored */
var SEARCH_OPTION_KEYS = [
    "includeFootnotes",
    "includeHiddenLayers",
    "includeLockedLayersForFind",
    "includeLockedStoriesForFind",
    "includeMasterPages",
    "searchBackwards",
    "objectType"
];

/* 検索時に使う値。ここに無いキーは退避・復帰だけ行う / Values applied while searching */
var SEARCH_OPTION_VALUES = {
    includeFootnotes: true,
    includeHiddenLayers: true,
    includeLockedLayersForFind: false,
    includeLockedStoriesForFind: false,
    includeMasterPages: true,
    searchBackwards: false
};

/**
 * 指定した範囲を検索し、見つかったオブジェクトを返す
 * 検索オプションと検索条件は、呼び出し前の状態に必ず戻す
 * @param {Document|Story|Text} searchTarget 検索する範囲
 * @param {string} searchModeKey "grep" / "text" / "object"
 * @param {object} findProperties 検索条件のプロパティ
 * @returns {Array<object>} 見つかったオブジェクトの配列
 */
function findInTarget(searchTarget, searchModeKey, findProperties) {
    var searchMode = SEARCH_MODES[searchModeKey];
    var searchOptions = app[searchMode.optionsKey];
    var savedOptions = {};
    var optionKey;
    var i;

    /* 現在の設定を退避（そのバージョンに無いキーは触らない）/ Save current options, skipping keys this version lacks */
    for (i = 0; i < SEARCH_OPTION_KEYS.length; i++) {
        optionKey = SEARCH_OPTION_KEYS[i];
        if (searchOptions.hasOwnProperty(optionKey)) {
            savedOptions[optionKey] = searchOptions[optionKey];
        }
    }

    try {
        for (i = 0; i < SEARCH_OPTION_KEYS.length; i++) {
            optionKey = SEARCH_OPTION_KEYS[i];
            if (savedOptions.hasOwnProperty(optionKey) && SEARCH_OPTION_VALUES.hasOwnProperty(optionKey)) {
                searchOptions[optionKey] = SEARCH_OPTION_VALUES[optionKey];
            }
        }
        if (searchMode.hasOwnProperty("objectType")) {
            searchOptions.objectType = searchMode.objectType;
        }

        app[searchMode.findPrefsKey] = NothingEnum.nothing;
        app[searchMode.findPrefsKey].properties = findProperties;
        return searchTarget[searchMode.findMethod](true);
    } finally {
        app[searchMode.findPrefsKey] = NothingEnum.nothing;
        for (i = 0; i < SEARCH_OPTION_KEYS.length; i++) {
            optionKey = SEARCH_OPTION_KEYS[i];
            if (savedOptions.hasOwnProperty(optionKey)) {
                searchOptions[optionKey] = savedOptions[optionKey];
            }
        }
    }
}

// =========================================
// 対象の解決 / Resolving the scope
// =========================================

/* 選択がテキストとみなせる型 / Selection types treated as text */
var TEXT_SELECTION_TYPES = {
    Text: true,
    Word: true,
    Character: true,
    Paragraph: true,
    Line: true,
    TextColumn: true,
    TextStyleRange: true,
    InsertionPoint: true
};

/**
 * 選択されたセル範囲を個々のセルへ展開する
 * 複数セルを選んでも Cell 1つとして返り、cells.length は実際の要素数と合わないことがある
 * @param {Cell} selectedCell 選択されたセル（範囲のこともある）
 * @returns {Array<Cell>} 個々のセル
 */
function resolveCellElements(selectedCell) {
    try {
        /* 単一セルでは cells を辿れないことがあるので保護する / A single cell may not expose .cells */
        var cellElements = selectedCell.cells.everyItem().getElements();
        if (cellElements && cellElements.length > 1) return cellElements;
    } catch (err) {}
    return [selectedCell];
}

/**
 * 現在の選択を、テキストとページアイテムに振り分ける
 * @returns {object} texts / items を持つオブジェクト
 */
function collectSelectionTargets() {
    var textTargets = [];
    var itemTargets = [];
    var cellTargets = [];
    var selectedObjects = app.selection;

    for (var i = 0; i < selectedObjects.length; i++) {
        var selectedObject = selectedObjects[i];
        var typeName = selectedObject.constructor.name;

        if (TEXT_SELECTION_TYPES.hasOwnProperty(typeName)) {
            textTargets.push(selectedObject);
        } else if (typeName === "Cell") {
            /* 複数セル選択は1つの Cell として返るので展開する /
               A multi-cell selection arrives as one Cell; expand it */
            var resolvedCells = resolveCellElements(selectedObject);
            for (var c = 0; c < resolvedCells.length; c++) {
                cellTargets.push(resolvedCells[c]);
                textTargets.push(resolvedCells[c].texts[0]);
            }
        } else {
            itemTargets.push(selectedObject);
            /* テキストフレームは中のテキストも対象にする / A selected text frame also contributes its text */
            if (typeName === "TextFrame") {
                textTargets.push(selectedObject.texts[0]);
            }
        }
    }
    return { texts: textTargets, items: itemTargets, cells: cellTargets };
}

/**
 * 選択箇所が属するストーリーと、そのストーリーのテキストフレームを集める
 * @param {object} selectionTargets collectSelectionTargets() の戻り値
 * @returns {object} texts（ストーリー）/ items（テキストフレーム）/ cells を持つオブジェクト
 */
function collectStoryTargets(selectionTargets) {
    var storyTargets = [];
    var frameTargets = [];
    var seenStoryIds = {};

    for (var i = 0; i < selectionTargets.texts.length; i++) {
        var parentStory = selectionTargets.texts[i].parentStory;
        if (seenStoryIds.hasOwnProperty(parentStory.id)) continue;
        seenStoryIds[parentStory.id] = true;
        storyTargets.push(parentStory);

        var textContainers = parentStory.textContainers;
        for (var j = 0; j < textContainers.length; j++) {
            frameTargets.push(textContainers[j]);
        }
    }
    return { texts: storyTargets, items: frameTargets, cells: [] };
}

/**
 * 長さのないカーソル位置を取り除く
 * @param {Array<object>} textTargets テキストの配列
 * @returns {Array<object>} InsertionPoint を除いた配列
 */
function excludeInsertionPoints(textTargets) {
    var pickedTargets = [];
    for (var i = 0; i < textTargets.length; i++) {
        if (textTargets[i].constructor.name !== "InsertionPoint") pickedTargets.push(textTargets[i]);
    }
    return pickedTargets;
}

/**
 * 「対象」の選択に応じて、処理する範囲を返す
 * 選択は処理の途中で変わりうるので、消去を始める前に確定させる
 * @param {string} scopeKey "document" / "story" / "selection"
 * @param {Document} targetDocument 対象ドキュメント
 * @returns {object} texts（検索する範囲）/ items（オブジェクト。null はドキュメント全体から検索）/ cells（選択されたセル）
 */
function resolveScopeTargets(scopeKey, targetDocument) {
    if (scopeKey === "document") {
        /* items が null なら、必要になった時点でドキュメント全体を検索する / null defers the document-wide object search */
        return { texts: [targetDocument], items: null, cells: [] };
    }
    var selectionTargets = collectSelectionTargets();
    if (scopeKey === "selection") {
        /* カーソルを置いただけの状態は「選択なし」として扱う。
           ストーリーの特定には使えるので、ここでだけ取り除く /
           A bare caret counts as no selection here, though it still identifies a story */
        return {
            texts: excludeInsertionPoints(selectionTargets.texts),
            items: selectionTargets.items,
            cells: selectionTargets.cells
        };
    }
    return collectStoryTargets(selectionTargets);
}

// =========================================
// オーバーライドの消去 / Clearing overrides
// =========================================

/**
 * 適用スタイルが処理対象かどうかを判定する
 * @param {object} appliedStyle 対象に適用されているスタイル
 * @param {object} selectedStyle ダイアログで選ばれたスタイル。ALL_STYLES なら常に true
 * @returns {boolean} 処理対象なら true
 */
function isTargetStyle(appliedStyle, selectedStyle) {
    return selectedStyle === ALL_STYLES || appliedStyle.id === selectedStyle.id;
}

/**
 * テキスト検索の条件を組み立てる（スタイルでの絞り込みを含む）
 * @param {string} findWhat 検索する文字列
 * @param {object} settings ダイアログで決めた設定
 * @returns {object} 検索条件のプロパティ
 */
function buildTextFindProperties(findWhat, settings) {
    var findProperties = { findWhat: findWhat };
    if (settings.paragraphStyle !== ALL_STYLES) {
        findProperties.appliedParagraphStyle = settings.paragraphStyle;
    }
    if (settings.characterStyle !== ALL_STYLES) {
        findProperties.appliedCharacterStyle = settings.characterStyle;
    }
    return findProperties;
}

/**
 * 見つかったテキストのオーバーライドをまとめて消去する
 * @param {Array<Text>} foundTexts 消去対象のテキスト
 * @param {OverrideType} overrideType 消去するオーバーライドの種類
 * @returns {void}
 */
function clearFoundTextOverrides(foundTexts, overrideType) {
    for (var i = 0; i < foundTexts.length; i++) {
        foundTexts[i].clearOverrides(overrideType);
    }
}

/**
 * テキストのスタイルオーバーライドを消去する
 * @param {Array<object>} textTargets 検索する範囲の配列
 * @param {object} settings ダイアログで決めた設定
 * @returns {void}
 */
function clearTextOverrides(textTargets, settings) {
    for (var i = 0; i < textTargets.length; i++) {
        clearFoundTextOverrides(findInTarget(textTargets[i], "grep", buildTextFindProperties(ALL_TEXT_GREP, settings)), settings.overrideType);
        /* フレームやセルの中身が表だけの場合、GREP検索では拾えないので別途処理 / A frame or cell holding only a table is not matched by GREP */
        clearFoundTextOverrides(findInTarget(textTargets[i], "text", buildTextFindProperties(TABLE_MARKER, settings)), settings.overrideType);
    }
}

/**
 * 表の中のセルスタイルのオーバーライドを消去する
 * @param {Table} targetTable 対象の表
 * @param {object} selectedCellStyle 対象のセルスタイル。ALL_STYLES ならすべて
 * @returns {void}
 */
function clearCellOverrides(targetTable, selectedCellStyle) {
    /* 絞り込みが無ければ一括で処理する / Clear in one call when no style filter is set */
    if (selectedCellStyle === ALL_STYLES) {
        targetTable.cells.everyItem().clearCellStyleOverrides();
        return;
    }
    var tableCells = targetTable.cells.everyItem().getElements();
    for (var i = 0; i < tableCells.length; i++) {
        if (isTargetStyle(tableCells[i].appliedCellStyle, selectedCellStyle)) {
            tableCells[i].clearCellStyleOverrides();
        }
    }
}

/**
 * 表スタイルとセルスタイルのオーバーライドを消去する
 * 選択されたセルは表のアンカー文字では見つからないので、別経路で処理する
 * @param {object} scopeTargets resolveScopeTargets() が返した処理範囲
 * @param {object} selectedTableStyle 対象の表スタイル。ALL_STYLES ならすべて
 * @param {object} selectedCellStyle 対象のセルスタイル。ALL_STYLES ならすべて
 * @returns {void}
 */
function clearTableOverrides(scopeTargets, selectedTableStyle, selectedCellStyle) {
    /* 同じ表を二度処理しないための控え / Keeps a table from being processed twice */
    var processedTableIds = {};
    var i;

    /* 選択されたセル：そのセルだけ消し、属する表の表スタイルは表ごとに1回 /
       Selected cells: clear just those cells, and their table's style once per table */
    for (i = 0; i < scopeTargets.cells.length; i++) {
        var targetCell = scopeTargets.cells[i];
        if (isTargetStyle(targetCell.appliedCellStyle, selectedCellStyle)) {
            targetCell.clearCellStyleOverrides();
        }
        var cellTable = targetCell.parent;
        if (processedTableIds.hasOwnProperty(cellTable.id)) continue;
        processedTableIds[cellTable.id] = true;
        if (isTargetStyle(cellTable.appliedTableStyle, selectedTableStyle)) {
            cellTable.clearTableStyleOverrides();
        }
    }

    /* 検索で見つかる表：表全体のセルを処理する / Tables found by search: every cell in them */
    for (i = 0; i < scopeTargets.texts.length; i++) {
        var foundMarkers = findInTarget(scopeTargets.texts[i], "text", { findWhat: TABLE_MARKER });
        for (var j = 0; j < foundMarkers.length; j++) {
            var foundTable = foundMarkers[j].tables[0];
            if (processedTableIds.hasOwnProperty(foundTable.id)) continue;
            processedTableIds[foundTable.id] = true;
            if (isTargetStyle(foundTable.appliedTableStyle, selectedTableStyle)) {
                foundTable.clearTableStyleOverrides();
                clearCellOverrides(foundTable, selectedCellStyle);
            }
        }
    }
}

/**
 * オブジェクトスタイルのオーバーライドを消去する
 * @param {Document} targetDocument 対象ドキュメント
 * @param {Array<PageItem>} objectTargets 対象のページアイテム。null ならドキュメント全体から検索する
 * @param {object} selectedObjectStyle 対象のオブジェクトスタイル。ALL_STYLES ならすべて
 * @returns {void}
 */
function clearObjectOverrides(targetDocument, objectTargets, selectedObjectStyle) {
    var candidateItems = (objectTargets === null) ? findInTarget(targetDocument, "object", {}) : objectTargets;
    for (var i = 0; i < candidateItems.length; i++) {
        /* ガイドなどオブジェクトスタイルを持たない選択は飛ばす /
           Skip a selection that has no object style, such as a guide */
        if (!candidateItems[i].hasOwnProperty("appliedObjectStyle")) continue;
        if (isTargetStyle(candidateItems[i].appliedObjectStyle, selectedObjectStyle)) {
            candidateItems[i].clearObjectStyleOverrides();
        }
    }
}

/**
 * 設定に従って、選ばれた種類のオーバーライドをすべて消去する
 * @param {Document} targetDocument 対象ドキュメント
 * @param {object} scopeTargets resolveScopeTargets() が返した処理範囲
 * @param {object} settings ダイアログで決めた設定
 * @returns {void}
 */
function clearStyleOverrides(targetDocument, scopeTargets, settings) {
    if (settings.processObjects) clearObjectOverrides(targetDocument, scopeTargets.items, settings.objectStyle);
    if (settings.processTables)  clearTableOverrides(scopeTargets, settings.tableStyle, settings.cellStyle);
    if (settings.processText)    clearTextOverrides(scopeTargets.texts, settings);
}

// =========================================
// UI / Dialog
// =========================================

/* 「対象」の選択肢 / Scope choices */
var SCOPE_CHOICES = [
    { labelKey: "radio.scopeDocument",  value: "document" },
    { labelKey: "radio.scopeStory",     value: "story" },
    { labelKey: "radio.scopeSelection", value: "selection" }
];

/* 「消去する種類」の選択肢 / Override type choices */
var OVERRIDE_TYPE_CHOICES = [
    { labelKey: "radio.bothTypes",     value: OverrideType.ALL },
    { labelKey: "radio.paragraphOnly", value: OverrideType.PARAGRAPH_ONLY },
    { labelKey: "radio.characterOnly", value: OverrideType.CHARACTER_ONLY }
];

/**
 * ラジオボタンの並びを追加する（排他になるよう専用グループに入れる）
 * @param {Panel|Group} parentContainer 追加先
 * @param {Array<object>} choices labelKey / value を持つ選択肢の配列
 * @param {number} defaultIndex 初期選択の位置
 * @param {string} tooltipKey tooltipのラベルキー
 * @param {string} [alignment] 横方向の配置。省略時は "left"
 * @returns {Array<RadioButton>} 追加したラジオボタン
 */
function addRadioRow(parentContainer, choices, defaultIndex, tooltipKey, alignment) {
    var radioGroup = parentContainer.add("group");
    setupRow(radioGroup, alignment || "left");
    radioGroup.spacing = RADIO_SPACING;

    var radioButtons = [];
    for (var i = 0; i < choices.length; i++) {
        var radioButton = radioGroup.add("radiobutton", undefined, getLabel(choices[i].labelKey));
        radioButton.helpTip = getLabel(tooltipKey);
        radioButtons.push(radioButton);
    }
    radioButtons[defaultIndex].value = true;
    return radioButtons;
}

/**
 * ラジオボタンで選ばれている値を取得する
 * @param {Array<RadioButton>} radioButtons 対象のラジオボタン
 * @param {Array<object>} choices 対応する選択肢の配列
 * @returns {object} 選ばれている選択肢の value。未選択なら先頭の value
 */
function getSelectedRadioValue(radioButtons, choices) {
    for (var i = 0; i < radioButtons.length; i++) {
        if (radioButtons[i].value) return choices[i].value;
    }
    return choices[0].value;
}

/**
 * 「対象」パネルを追加する
 * @param {Window} parentWindow 追加先のウィンドウ
 * @returns {Array<RadioButton>} 追加したラジオボタン
 */
function addScopePanel(parentWindow) {
    var scopePanel = parentWindow.add("panel", undefined, getLabel("panel.scope"));
    setupPanel(scopePanel);
    return addRadioRow(scopePanel, SCOPE_CHOICES, DEFAULT_SCOPE_INDEX, "tooltip.scope", "center");
}

/**
 * スタイルグループ名を含めた表示名を返す（グループは : 区切り）
 * @param {ParagraphStyle|CharacterStyle|TableStyle|CellStyle|ObjectStyle} style 対象のスタイル
 * @returns {string} グループ名を含む表示名
 */
function getStyleDisplayName(style) {
    var displayName = style.name;
    var container = style.parent;
    while (container.constructor.name.match(/Group$/)) {
        displayName = container.name + ":" + displayName;
        container = container.parent;
    }
    return displayName;
}

/**
 * ドロップダウンに [すべて] と各スタイルを並べる
 * @param {DropDownList} styleDropdown 対象のドロップダウン
 * @param {Array<object>} styleCollection 並べるスタイルの配列
 * @returns {void}
 */
function fillStyleDropdown(styleDropdown, styleCollection) {
    styleDropdown.add("item", getLabel("dropdown.allStyles"));
    for (var i = 0; i < styleCollection.length; i++) {
        var styleItem = styleDropdown.add("item", getStyleDisplayName(styleCollection[i]));
        styleItem.styleRef = styleCollection[i];
    }
    styleDropdown.selection = 0;
}

/**
 * ドロップダウンで選ばれているスタイルを取得する
 * @param {DropDownList} styleDropdown 対象のドロップダウン
 * @returns {object} 選ばれたスタイル。先頭の [すべて] の場合は ALL_STYLES
 */
function getSelectedStyle(styleDropdown) {
    var selectedItem = styleDropdown.selection;
    if (!selectedItem || selectedItem.index === 0) return ALL_STYLES;
    return selectedItem.styleRef;
}

/**
 * 種類ごとのパネル（チェックボックス＋スタイル絞り込み行）を追加する
 * @param {Window} parentWindow 追加先のウィンドウ
 * @param {string} panelLabelKey パネル見出しのラベルキー
 * @param {string} checkboxLabelKey チェックボックスのラベルキー
 * @param {string} tooltipKey チェックボックスのtooltipキー
 * @param {boolean} initialValue チェックボックスの初期状態
 * @returns {object} panel / checkbox / optionRows を持つオブジェクト
 */
function addOverrideSection(parentWindow, panelLabelKey, checkboxLabelKey, tooltipKey, initialValue) {
    var sectionPanel = parentWindow.add("panel", undefined, getLabel(panelLabelKey));
    setupPanel(sectionPanel);

    var enableCheckbox = sectionPanel.add("checkbox", undefined, getLabel(checkboxLabelKey));
    enableCheckbox.alignment = ["left", "top"];
    enableCheckbox.value = initialValue;
    enableCheckbox.helpTip = getLabel(tooltipKey);

    var overrideSection = {
        panel: sectionPanel,
        checkbox: enableCheckbox,
        optionRows: [],
        onToggle: null   /* チェック状態が変わったときに呼ぶ / Called when the checkbox changes */
    };

    enableCheckbox.onClick = function () {
        for (var i = 0; i < overrideSection.optionRows.length; i++) {
            overrideSection.optionRows[i].enabled = enableCheckbox.value;
        }
        if (overrideSection.onToggle) overrideSection.onToggle();
    };

    return overrideSection;
}

/**
 * 処理パネルにスタイル絞り込みの行を追加する
 * @param {object} overrideSection addOverrideSection() が返したオブジェクト
 * @param {string} fieldLabelKey 項目名のラベルキー
 * @param {Array<object>} styleCollection ドロップダウンに並べるスタイルの配列
 * @returns {DropDownList} 追加したドロップダウン
 */
function addStyleFilterRow(overrideSection, fieldLabelKey, styleCollection) {
    var filterRow = overrideSection.panel.add("group");
    setupRow(filterRow, "fill");
    filterRow.enabled = overrideSection.checkbox.value;
    addRowLabel(filterRow, fieldLabelKey);

    var styleDropdown = filterRow.add("dropdownlist");
    styleDropdown.preferredSize.width = DROPDOWN_WIDTH;
    styleDropdown.maximumSize.width = DROPDOWN_WIDTH;
    styleDropdown.alignment = ["fill", "center"];
    styleDropdown.helpTip = getLabel("tooltip.styleFilter");
    fillStyleDropdown(styleDropdown, styleCollection);

    overrideSection.optionRows.push(filterRow);
    return styleDropdown;
}

/**
 * 処理パネルに「消去する種類」の行を追加する
 * @param {object} overrideSection addOverrideSection() が返したオブジェクト
 * @returns {Array<RadioButton>} 追加したラジオボタン
 */
function addOverrideTypeRow(overrideSection) {
    var typeRow = overrideSection.panel.add("group");
    setupRow(typeRow, "fill");
    typeRow.enabled = overrideSection.checkbox.value;
    addRowLabel(typeRow, "fieldLabel.overrideType");

    var typeRadios = addRadioRow(typeRow, OVERRIDE_TYPE_CHOICES, DEFAULT_OVERRIDE_TYPE_INDEX, "tooltip.overrideType");
    overrideSection.optionRows.push(typeRow);
    return typeRadios;
}

/**
 * ダイアログ下部のボタン行を追加する
 * @param {Window} parentWindow 追加先のウィンドウ
 * @returns {Button} 実行ボタン
 */
function addButtonRow(parentWindow) {
    var btnRowGroup = parentWindow.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
    btnRowGroup.alignment = ["right", "bottom"];
    btnRowGroup.alignChildren = ["right", "center"];
    btnRowGroup.spacing = BUTTON_SPACING;
    btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    return btnRowGroup.add("button", undefined, getLabel("button.run"), { name: "ok" });
}

/**
 * 設定ダイアログを表示する
 * @param {Document} targetDocument 対象ドキュメント
 * @returns {object} 設定オブジェクト。キャンセルされた場合は null
 */
function showSettingsDialog(targetDocument) {
    var settingsDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(settingsDialog);

    var scopeRadios = addScopePanel(settingsDialog);

    var textSection = addOverrideSection(settingsDialog, "panel.text", "checkbox.text", "tooltip.clearText", DEFAULT_PROCESS_TEXT);
    var paragraphStyleDropdown = addStyleFilterRow(textSection, "fieldLabel.paragraphStyle", targetDocument.allParagraphStyles);
    var characterStyleDropdown = addStyleFilterRow(textSection, "fieldLabel.characterStyle", targetDocument.allCharacterStyles);
    var overrideTypeRadios = addOverrideTypeRow(textSection);

    var tableSection = addOverrideSection(settingsDialog, "panel.table", "checkbox.table", "tooltip.clearTable", DEFAULT_PROCESS_TABLES);
    var tableStyleDropdown = addStyleFilterRow(tableSection, "fieldLabel.tableStyle", targetDocument.allTableStyles);
    var cellStyleDropdown = addStyleFilterRow(tableSection, "fieldLabel.cellStyle", targetDocument.allCellStyles);

    var objectSection = addOverrideSection(settingsDialog, "panel.object", "checkbox.object", "tooltip.clearObject", DEFAULT_PROCESS_OBJECTS);
    var objectStyleDropdown = addStyleFilterRow(objectSection, "fieldLabel.objectStyle", targetDocument.allObjectStyles);

    var btnRun = addButtonRow(settingsDialog);

    /* 3つとも処理しない設定では実行できないようにする / Nothing to process means nothing to run */
    var overrideSections = [textSection, tableSection, objectSection];
    var updateRunButton = function () {
        var anyChecked = false;
        for (var i = 0; i < overrideSections.length; i++) {
            if (overrideSections[i].checkbox.value) anyChecked = true;
        }
        btnRun.enabled = anyChecked;
    };
    for (var i = 0; i < overrideSections.length; i++) {
        overrideSections[i].onToggle = updateRunButton;
    }
    updateRunButton();

    settingsDialog.center();
    if (settingsDialog.show() !== 1) return null;

    return {
        scope:          getSelectedRadioValue(scopeRadios, SCOPE_CHOICES),
        processText:    textSection.checkbox.value,
        paragraphStyle: getSelectedStyle(paragraphStyleDropdown),
        characterStyle: getSelectedStyle(characterStyleDropdown),
        overrideType:   getSelectedRadioValue(overrideTypeRadios, OVERRIDE_TYPE_CHOICES),
        processTables:  tableSection.checkbox.value,
        tableStyle:     getSelectedStyle(tableStyleDropdown),
        cellStyle:      getSelectedStyle(cellStyleDropdown),
        processObjects: objectSection.checkbox.value,
        objectStyle:    getSelectedStyle(objectStyleDropdown)
    };
}

// =========================================
// メイン処理 / Main
// =========================================

if (app.documents.length === 0 || app.layoutWindows.length === 0) {
    alert(getLabel("alert.noDocument"));
    return;
}

var targetDocument = app.activeDocument;
var settings = showSettingsDialog(targetDocument);
if (settings === null) return;

/* 消去を始めると選択が変わりうるので、処理範囲は先に確定させる / Resolve the scope before anything changes the selection */
var scopeTargets = resolveScopeTargets(settings.scope, targetDocument);
if (settings.scope !== "document" && scopeTargets.texts.length === 0 && scopeTargets.items.length === 0) {
    alert(getLabel("alert.noSelection"));
    return;
}

/* 検索系のバージョン差による警告を抑える / Avoid version-related warnings from find/change */
var savedScriptVersion = app.scriptPreferences.version;
app.scriptPreferences.version = parseInt(app.version, 10);

try {
    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(function () {
        clearStyleOverrides(targetDocument, scopeTargets, settings);
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.clearOverrides"));
} catch (err) {
    alert(getLabel("alert.failed") + "\n" + err + " (line " + err.line + ")");
} finally {
    app.scriptPreferences.version = savedScriptVersion;
}

})();
