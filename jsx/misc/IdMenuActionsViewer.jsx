#target indesign

/*

### 概要

InDesign のメニューアクションを一覧し、エリア・メニュー・名前・ID で絞り込んで調べます。
選択したアクションのキーストリング（$ID/…）と実行コード（app.menuActions.itemByID(…).invoke();）をコピーできます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdMenuActionsViewer.md

### Overview

Lists InDesign menu actions and lets you filter them by area, menu, name, and ID.
You can copy the key strings ($ID/…) and the invoke code (app.menuActions.itemByID(…).invoke();) of the selected action.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdMenuActionsViewer.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdMenuActionsViewer";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.14";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Peter Kahrel";                 /* 作者 / author */
var SCRIPT_MODIFIED = "Masahiro Takano (@swwwitch)";  /* 改変 / modified by */
var SCRIPT_RELEASED = "2026-09-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-25";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdMenuActionsViewer.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdMenuActionsViewer.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @author Peter Kahrel（原作 menu_actions.jsx）
 * @discussion https://creativepro.com/menu_actions/
 * http://kasyan.ho.com.ua/open_menu_item.html
 */

(function () {

    // =========================================
    // 基本設定 / Settings
    // =========================================

    /* フォント名・スタイル名などのノイズを隠す / Hide fonts, style names and other noise */
    var HIDE_NOISE = true;

    /* フォント名・スタイル名が並ぶ ID の範囲 / ID range occupied by font and style names */
    var NOISE_ID_MIN = 57603;
    var NOISE_ID_MAX = 61066;

    /* 名前末尾の「 (xyz:)」を取り除く / Strip the occasional " (xyz:)" code */
    var CODE_SUFFIX_PATTERN = / \([a-z]+?:\)/;

    /* ファイル名（最近使用したファイル・スクリプト）/ File names (recent files, scripts) */
    var FILE_NAME_PATTERN = /\.(?:indd|jsx?(?:bin)?)$/i;

    /* 除外するエリア（英語版の名称）/ Areas to exclude (English names) */
    var NOISE_AREA_PATTERN = /^(?:Text Selection|Menu:Insert)/i;

    /*
     * メニューによる絞り込み。エリア名が areaPrefixes のどれかで始まれば対象にする（日本語版・英語版の両方を並べる）
     * Menu filter: an area matches when it starts with one of areaPrefixes (both Japanese and English names)
     */
    var MENU_FILTERS = [
        { labelKey: "file",      areaPrefixes: ["ファイル", "File"] },
        { labelKey: "edit",      areaPrefixes: ["編集", "Edit"] },
        { labelKey: "view",      areaPrefixes: ["表示", "View"] },
        { labelKey: "format",    areaPrefixes: ["フォーマット", "Format"] },
        { labelKey: "find",      areaPrefixes: ["検索", "Find"] },
        { labelKey: "tools",     areaPrefixes: ["ツール", "Tool"] },
        { labelKey: "window",    areaPrefixes: ["ウィンドウ", "Window"] },
        { labelKey: "panelMenu", areaPrefixes: ["パネルメニュー", "Panel Menu"] }
    ];

    /* 並び順ラジオボタンの順（リストの列順にそろえる）/ Sort radio order (matches the list columns) */
    var SORT_KEYS = ["area", "name", "id"];
    var DEFAULT_SORT_INDEX = 0;

    // =========================================
    // レイアウト / Layout
    // =========================================
    var WINDOW_MARGINS        = 16;              /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING        = 12;              /* ウィンドウ内の要素間隔 / window spacing */
    var LIST_SIZE             = [560, 520];      /* リストの寸法 [幅,高さ] / list size */
    var COLUMN_WIDTHS         = [70, 470];       /* 列幅（エリア・名前）/ column widths (area, name) */
    var FILTER_FIELD_WIDTH    = 300;             /* エリア・名前・ID の欄の幅 / width of the area and name/ID fields */
    var ROW_LABEL_WIDTH       = 130;             /* 行ラベルの幅 / row label width */
    var COPY_BUTTON_WIDTH     = 70;              /* ［コピー］ボタンの幅 / width of the Copy buttons */
    var BUTTON_ROW_TOP_MARGIN = 10;              /* ボタン列の上余白 / top margin above the button row */
    var PROGRESS_BAR_WIDTH    = 320;             /* 進行状況バーの幅 / progress bar width */
    var PROGRESS_STEP         = 250;             /* 進行状況を更新する間隔（件）/ items between progress updates */

    // =========================================
    // ラベル定義 / Labels
    // =========================================

    /**
     * UI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentUILang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentUILang();

    var LABELS = {
        dialog: {
            title: { ja: "メニューアクション", en: "Menu Actions" }
        },
        fieldLabel: {
            name:       { ja: "名前", en: "Name" },
            nameOrId:   { ja: "名前・ID", en: "Name / ID" },
            area:       { ja: "エリア", en: "Area" },
            menu:       { ja: "メニュー", en: "Menu" },
            id:         { ja: "ID", en: "ID" },
            sortOrder:  { ja: "並び順", en: "Sort by" },
            keyStrings: { ja: "キーストリング", en: "Key strings" },
            invokeCode: { ja: "実行コード", en: "Invoke code" }
        },
        dropdown: {
            allAreas: { ja: "［すべて］", en: "[All]" },
            allMenus: { ja: "［すべて］", en: "[All]" }
        },
        menuFilter: {
            file:      { ja: "ファイル", en: "File" },
            edit:      { ja: "編集", en: "Edit" },
            view:      { ja: "表示", en: "View" },
            format:    { ja: "フォーマット", en: "Format" },
            find:      { ja: "検索", en: "Find" },
            tools:     { ja: "ツール", en: "Tools" },
            window:    { ja: "ウィンドウ", en: "Window" },
            panelMenu: { ja: "パネルメニュー", en: "Panel Menus" }
        },
        button: {
            copy:  { ja: "コピー", en: "Copy" },
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            menu: {
                ja: "エリア名がこのメニュー名で始まるアクションに絞り込みます。エリアの候補も絞られます。",
                en: "Filters to actions whose area starts with this menu name. The area choices are narrowed as well."
            },
            nameOrId: {
                ja: "名前に含まれる文字で絞り込みます（大文字小文字を区別しません）。正規表現も使えます。数字だけを入れると、その ID のアクションも対象にします。Enter で確定します。",
                en: "Filters by text contained in the name (case-insensitive). Regular expressions are allowed. Digits only also match the action with that ID. Press Enter to apply."
            },
            copyKeyStrings: {
                ja: "選択したアクション名に対応するキーストリング（$ID/…）をクリップボードにコピーします。複数ある場合は「 | 」で区切ります。",
                en: "Copies the key strings ($ID/…) for the selected action name to the clipboard. Multiple key strings are separated by \" | \"."
            },
            copyInvokeCode: {
                ja: "選択したアクションを実行するコード（app.menuActions.itemByID(…).invoke();）をクリップボードにコピーします。",
                en: "Copies the code that invokes the selected action (app.menuActions.itemByID(…).invoke();) to the clipboard."
            }
        },
        message: {
            noKeyStrings: { ja: "（見つかりません）", en: "(not found)" }
        },
        alert: {
            copyFailed: { ja: "クリップボードにコピーできませんでした。", en: "Could not copy to the clipboard." }
        },
        progress: {
            loading:  { ja: "メニューアクションを読み込み中…", en: "Loading menu actions…" },
            building: { ja: "リストを作成中…", en: "Building the list…" }
        },
        counter: {
            itemUnit: { ja: "件", en: " items" }
        }
    };

    /**
     * 現在のUI言語のラベルを返す
     * @param {Object} labelSet - { ja: string, en: string } 形式のラベル
     * @returns {string} ラベル文字列
     */
    function getLabel(labelSet) {
        return labelSet[uiLang] || labelSet.en;
    }

    /**
     * 言語別のコロンを付けたラベルを返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - { ja: string, en: string } 形式のラベル
     * @returns {string} コロン付きラベル
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // データ / Data
    // =========================================

    /**
     * @typedef {Object} ActionEntry
     * @property {string} name - 表示用の名前（コード接尾辞を除いたもの）
     * @property {string} area - エリア
     * @property {number} id - アクション ID
     */

    /**
     * 一覧に出さないアクションか判定する
     * @param {string} actionName - アクション名
     * @param {string} areaName - エリア名
     * @param {number} actionId - アクション ID
     * @returns {boolean} 除外するなら true
     */
    function isNoiseAction(actionName, areaName, actionId) {
        if (!HIDE_NOISE) {
            return false;
        }
        return (actionId >= NOISE_ID_MIN && actionId <= NOISE_ID_MAX)
            || FILE_NAME_PATTERN.test(actionName)
            || NOISE_AREA_PATTERN.test(areaName);
    }

    /**
     * メニューアクションを収集する
     * 1件ずつプロパティを読むと遅いので、everyItem() で列ごとにまとめて取る
     * @param {Object} progressWindow - createProgressWindow() の戻り値
     * @returns {ActionEntry[]} アクションの一覧
     */
    function collectActions(progressWindow) {
        var allItems = app.menuActions.everyItem();
        var actionNames = allItems.name;
        progressWindow.update(1);
        var areaNames = allItems.area;
        progressWindow.update(2);
        var actionIds = allItems.id;
        progressWindow.update(3);

        var entries = [];
        for (var i = 0; i < actionIds.length; i++) {
            if (isNoiseAction(actionNames[i], areaNames[i], actionIds[i])) {
                continue;
            }
            entries.push({
                name: actionNames[i].replace(CODE_SUFFIX_PATTERN, ""),
                area: areaNames[i],
                id: actionIds[i]
            });
        }
        return entries;
    }

    /**
     * エリア名の一覧を重複なく昇順で返す
     * @param {ActionEntry[]} entries - アクションの一覧
     * @returns {string[]} エリア名の一覧
     */
    function collectAreaNames(entries) {
        var seen = {};
        var areaNames = [];
        for (var i = 0; i < entries.length; i++) {
            if (!seen[entries[i].area]) {
                seen[entries[i].area] = true;
                areaNames.push(entries[i].area);
            }
        }
        return areaNames.sort();
    }

    /* 並べ替えキーの区切り（どの文字より前に並ぶ）/ Sort key separator (sorts before any character) */
    var SORT_KEY_SEPARATOR = "\u0001";

    /**
     * ID を桁そろえした文字列にする（文字列の並べ替えで数値順になるように）
     * @param {number} actionId - アクション ID
     * @returns {string} 10桁にゼロ埋めした ID
     */
    function padActionId(actionId) {
        return ("0000000000" + actionId).slice(-10);
    }

    /**
     * 並べ替え用の文字列キーを作る（同順位は名前・ID で決める）
     * @param {ActionEntry} actionEntry - 対象のアクション
     * @param {string} sortKey - "name" / "area" / "id"
     * @returns {string} 並べ替えキー
     */
    function buildSortKey(actionEntry, sortKey) {
        var paddedId = padActionId(actionEntry.id);
        if (sortKey === "id") {
            return paddedId;
        }
        var nameKey = actionEntry.name + SORT_KEY_SEPARATOR + paddedId;
        return (sortKey === "area") ? actionEntry.area + SORT_KEY_SEPARATOR + nameKey : nameKey;
    }

    /**
     * 指定したキーでアクションを並べ替える
     * 比較関数を渡すと ExtendScript では数秒かかるため、文字列キーを組み込みの sort() で並べる
     * @param {ActionEntry[]} entries - 並べ替えるアクションの一覧（直接並べ替える）
     * @param {string} sortKey - "name" / "area" / "id"
     * @returns {void}
     */
    function sortActions(entries, sortKey) {
        var sortKeys = [];
        for (var i = 0; i < entries.length; i++) {
            sortKeys.push(buildSortKey(entries[i], sortKey) + SORT_KEY_SEPARATOR + i);
        }
        sortKeys.sort();

        var sortedEntries = [];
        for (var j = 0; j < sortKeys.length; j++) {
            var originalIndex = Number(sortKeys[j].substring(sortKeys[j].lastIndexOf(SORT_KEY_SEPARATOR) + 1));
            sortedEntries.push(entries[originalIndex]);
        }
        for (var k = 0; k < sortedEntries.length; k++) {
            entries[k] = sortedEntries[k];
        }
    }

    /**
     * 名前の検索文字列から照合用の正規表現を作る（正規表現として不正なら文字どおりに探す）
     * @param {string} searchText - 検索文字列
     * @returns {RegExp|null} 正規表現。空欄なら null
     */
    function buildNamePattern(searchText) {
        if (searchText === "") {
            return null;
        }
        /* 入力途中の不正な正規表現は例外になる / Incomplete regex input throws */
        try {
            return new RegExp(searchText, "i");
        } catch (e) {
            return new RegExp(searchText.replace(/[\\^$.*+?()[\]{}|\/]/g, "\\$&"), "i");
        }
    }

    /**
     * エリア名がメニューの絞り込み条件に合うか判定する
     * @param {string} areaName - エリア名
     * @param {Object|null} menuFilter - MENU_FILTERS の要素。null ならすべて
     * @returns {boolean} 合えば true
     */
    function matchesMenuFilter(areaName, menuFilter) {
        if (menuFilter === null) {
            return true;
        }
        for (var i = 0; i < menuFilter.areaPrefixes.length; i++) {
            if (areaName.indexOf(menuFilter.areaPrefixes[i]) === 0) {
                return true;
            }
        }
        return false;
    }

    /**
     * メニューの絞り込み条件に合うエリア名だけを返す
     * @param {string[]} areaNames - エリア名の一覧
     * @param {Object|null} menuFilter - MENU_FILTERS の要素。null ならすべて
     * @returns {string[]} 条件に合うエリア名
     */
    function filterAreaNames(areaNames, menuFilter) {
        var matched = [];
        for (var i = 0; i < areaNames.length; i++) {
            if (matchesMenuFilter(areaNames[i], menuFilter)) {
                matched.push(areaNames[i]);
            }
        }
        return matched;
    }

    /**
     * 名前・ID、エリア、メニューの条件で絞り込む
     * @param {ActionEntry[]} entries - アクションの一覧
     * @param {string} searchText - 名前の検索文字列（数字だけなら ID とも照合する）
     * @param {string|null} areaName - エリア名。null ならすべて
     * @param {Object|null} menuFilter - MENU_FILTERS の要素。null ならすべて
     * @returns {ActionEntry[]} 条件に合うアクション
     */
    function filterActions(entries, searchText, areaName, menuFilter) {
        var trimmedText = searchText.replace(/^\s+|\s+$/g, "");
        var namePattern = buildNamePattern(trimmedText);
        var idText = /^\d+$/.test(trimmedText) ? trimmedText : null;
        var matched = [];
        for (var i = 0; i < entries.length; i++) {
            var matchesId = (idText !== null && String(entries[i].id) === idText);
            if (namePattern && !matchesId && !namePattern.test(entries[i].name)) {
                continue;
            }
            if (areaName !== null && entries[i].area !== areaName) {
                continue;
            }
            if (!matchesMenuFilter(entries[i].area, menuFilter)) {
                continue;
            }
            matched.push(entries[i]);
        }
        return matched;
    }

    /**
     * 数値を3桁区切りの文字列にする
     * @param {number} value - 数値
     * @returns {string} 3桁区切りの文字列
     */
    function formatThousands(value) {
        return String(value).replace(/(\d)(?=(\d\d\d)+$)/g, "$1,");
    }

    /**
     * 件数入りのダイアログタイトルを作る
     * @param {number} itemCount - 表示中の件数
     * @returns {string} ダイアログタイトル
     */
    function buildDialogTitle(itemCount) {
        var countText = formatThousands(itemCount) + getLabel(LABELS.counter.itemUnit);
        return getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION
            + (uiLang === "ja" ? "（" + countText + "）" : " (" + countText + ")");
    }

    /**
     * アクションを実行するコードを作る
     * @param {ActionEntry} actionEntry - 対象のアクション
     * @returns {string} app.menuActions.itemByID(…).invoke(); 形式のコード
     */
    function buildInvokeCode(actionEntry) {
        return "app.menuActions.itemByID(" + actionEntry.id + ").invoke();";
    }

    // =========================================
    // 進行状況 / Progress
    // =========================================

    /**
     * 進行状況バーの小さなパレットを表示する
     * @param {Object} messageSet - 表示する文言
     * @param {number} maxValue - バーの最大値
     * @returns {{update: function(number): void, reset: function(Object, number): void, close: function(): void}} 操作用の関数
     */
    function createProgressWindow(messageSet, maxValue) {
        var progressPalette = new Window("palette", getLabel(LABELS.dialog.title));
        progressPalette.orientation = "column";
        progressPalette.alignChildren = ["fill", "top"];
        progressPalette.margins = WINDOW_MARGINS;

        var messageText = progressPalette.add("statictext", undefined, getLabel(messageSet));
        var progressBar = progressPalette.add("progressbar", undefined, 0, maxValue);
        progressBar.preferredSize.width = PROGRESS_BAR_WIDTH;

        progressPalette.show();
        progressPalette.update();

        return {
            update: function (value) {
                progressBar.value = value;
                progressPalette.update();
            },
            reset: function (nextMessageSet, nextMaxValue) {
                messageText.text = getLabel(nextMessageSet);
                progressBar.maxvalue = nextMaxValue;
                progressBar.value = 0;
                progressPalette.update();
            },
            close: function () {
                progressPalette.close();
            }
        };
    }

    // =========================================
    // ダイアログ部品 / Dialog parts
    // =========================================

    /**
     * ラベルと入力部品を1行に並べる
     * @param {Group} parent - 追加先
     * @param {Object} labelSet - 行ラベル
     * @returns {Group} 行グループ（入力部品はここへ追加する）
     */
    function addFieldRow(parent, labelSet) {
        var rowGroup = parent.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignment = ["fill", "top"];
        rowGroup.alignChildren = ["left", "center"];
        var rowLabel = rowGroup.add("statictext", undefined, labelText(labelSet));
        /* 幅を固定しないと右揃えが効かない / Right alignment needs a fixed width */
        rowLabel.minimumSize.width = ROW_LABEL_WIDTH;
        rowLabel.preferredSize.width = ROW_LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowGroup;
    }

    /**
     * エリア・名前の2列のリストを作る（ID は実行コードの欄で確認する）
     * @param {Group} parent - 追加先
     * @returns {ListBox} アクション一覧
     */
    function buildActionList(parent) {
        var actionList = parent.add("listbox", undefined, "", {
            multiselect: false,
            numberOfColumns: 2,
            showHeaders: true,
            columnTitles: [
                getLabel(LABELS.fieldLabel.area),
                getLabel(LABELS.fieldLabel.name)
            ],
            columnWidths: COLUMN_WIDTHS
        });
        actionList.preferredSize = LIST_SIZE;
        return actionList;
    }

    /**
     * エリアのドロップダウンの項目を入れ替える（先頭は［すべて］）
     * @param {DropDownList} areaDropdown - エリアのドロップダウン
     * @param {string[]} areaNames - 並べるエリア名
     * @returns {void}
     */
    function fillAreaDropdown(areaDropdown, areaNames) {
        areaDropdown.removeAll();
        areaDropdown.add("item", getLabel(LABELS.dropdown.allAreas));
        for (var i = 0; i < areaNames.length; i++) {
            areaDropdown.add("item", areaNames[i]);
        }
        areaDropdown.selection = 0;
    }

    /**
     * リスト上の絞り込み条件（エリア・メニュー・名前と ID・並び順）の行を作る
     * @param {Group} parent - 追加先
     * @param {string[]} areaNames - エリア名の一覧
     * @returns {{areaDropdown: DropDownList, menuDropdown: DropDownList, searchInput: EditText, sortRadios: RadioButton[]}} 作成した部品
     */
    function buildFilterRows(parent, areaNames) {
        var areaDropdown = addFieldRow(parent, LABELS.fieldLabel.area).add("dropdownlist", undefined, []);
        areaDropdown.preferredSize.width = FILTER_FIELD_WIDTH;
        fillAreaDropdown(areaDropdown, areaNames);

        var menuItems = [getLabel(LABELS.dropdown.allMenus)];
        for (var i = 0; i < MENU_FILTERS.length; i++) {
            menuItems.push(getLabel(LABELS.menuFilter[MENU_FILTERS[i].labelKey]));
        }
        var menuDropdown = addFieldRow(parent, LABELS.fieldLabel.menu).add("dropdownlist", undefined, menuItems);
        menuDropdown.preferredSize.width = FILTER_FIELD_WIDTH;
        menuDropdown.helpTip = getLabel(LABELS.tooltip.menu);
        menuDropdown.selection = 0;

        var searchInput = addFieldRow(parent, LABELS.fieldLabel.nameOrId).add("edittext", undefined, "");
        searchInput.preferredSize.width = FILTER_FIELD_WIDTH;
        searchInput.helpTip = getLabel(LABELS.tooltip.nameOrId);
        searchInput.active = true;

        /* 並び順の項目名は列名と同じ。同じ行グループに入れて排他にする / Reuse column names; same parent keeps them exclusive */
        var sortRowGroup = addFieldRow(parent, LABELS.fieldLabel.sortOrder);
        var sortRadios = [];
        for (var j = 0; j < SORT_KEYS.length; j++) {
            sortRadios.push(sortRowGroup.add("radiobutton", undefined, getLabel(LABELS.fieldLabel[SORT_KEYS[j]])));
        }
        sortRadios[DEFAULT_SORT_INDEX].value = true;

        return {
            areaDropdown: areaDropdown,
            menuDropdown: menuDropdown,
            searchInput: searchInput,
            sortRadios: sortRadios
        };
    }

    /**
     * リスト下の1行（ラベル・読み取り専用の欄・［コピー］ボタン）を作る
     * @param {Group} parent - 追加先
     * @param {Object} labelSet - 行ラベル
     * @param {Object} tooltipSet - ［コピー］ボタンのツールチップ
     * @returns {{valueField: EditText, btnCopy: Button}} 作成した部品
     */
    function addDetailRow(parent, labelSet, tooltipSet) {
        var rowGroup = addFieldRow(parent, labelSet);

        /* コピーできるよう読み取り専用の欄に出す / Read-only field so the text can be copied */
        var valueField = rowGroup.add("edittext", undefined, "", { readonly: true });
        valueField.alignment = ["fill", "center"];

        /* 余った幅を吸わないよう幅を固定 / Fix the width so it does not absorb spare space */
        var btnCopy = rowGroup.add("button", undefined, getLabel(LABELS.button.copy));
        btnCopy.alignment = ["right", "center"];
        btnCopy.minimumSize.width = COPY_BUTTON_WIDTH;
        btnCopy.preferredSize.width = COPY_BUTTON_WIDTH;
        btnCopy.maximumSize.width = COPY_BUTTON_WIDTH;
        btnCopy.helpTip = getLabel(tooltipSet);

        return { valueField: valueField, btnCopy: btnCopy };
    }

    /**
     * ボタンエリア（右端に［閉じる］）を作る
     * @param {Window} dialogWindow - 追加先のダイアログ
     * @returns {void}
     */
    function buildButtonRow(dialogWindow) {
        var btnRowGroup = dialogWindow.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["right", "bottom"];
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.close), { name: "cancel" });
    }

    // =========================================
    // 表示 / Display
    // =========================================

    /**
     * アクションのキーストリングを1行にまとめる
     * @param {ActionEntry} actionEntry - 対象のアクション
     * @returns {string} 「 | 」区切りのキーストリング。なければ空文字
     */
    function buildKeyStringsText(actionEntry) {
        return app.findKeyStrings(actionEntry.name).join(" | ");
    }

    /**
     * 文字列をクリップボードにコピーする（Mac は AppleScript、Windows は VBScript 経由）
     * @param {string} copyText - コピーする文字列
     * @returns {void}
     */
    function copyToClipboard(copyText) {
        var isMac = (File.fs === "Macintosh");
        var scriptSource = isMac
            ? 'set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"'
            : 'CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""') + '"';
        /* スクリプト実行の許可がないと失敗する / Fails without permission to run scripts */
        try {
            app.doScript(scriptSource, isMac ? ScriptLanguage.APPLESCRIPT_LANGUAGE : ScriptLanguage.VISUAL_BASIC);
        } catch (e) {
            alert(getLabel(LABELS.alert.copyFailed) + "\n" + e.message);
        }
    }

    /**
     * メインダイアログを作って表示する
     * @param {ActionEntry[]} allActions - すべてのアクション
     * @param {Object} progressWindow - createProgressWindow() の戻り値（最初の一覧作成後に閉じる）
     * @returns {void}
     */
    function showMainDialog(allActions, progressWindow) {
        var visibleActions = [];

        var mainDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        mainDialog.orientation = "column";
        mainDialog.alignChildren = ["fill", "top"];
        mainDialog.margins = WINDOW_MARGINS;
        mainDialog.spacing = WINDOW_SPACING;

        /* 上から絞り込み条件・リスト・詳細行 / Filters, list and detail rows from top to bottom */
        var allAreaNames = collectAreaNames(allActions);
        var filterControls = buildFilterRows(mainDialog, allAreaNames);
        var actionList = buildActionList(mainDialog);
        var keyStringsRow = addDetailRow(mainDialog, LABELS.fieldLabel.keyStrings, LABELS.tooltip.copyKeyStrings);
        var invokeCodeRow = addDetailRow(mainDialog, LABELS.fieldLabel.invokeCode, LABELS.tooltip.copyInvokeCode);
        buildButtonRow(mainDialog);

        /**
         * 一覧で選択中のアクションを返す
         * @returns {ActionEntry|null} 選択中のアクション。なければ null
         */
        function getSelectedAction() {
            return actionList.selection ? visibleActions[actionList.selection.index] : null;
        }

        /**
         * 選択が変わったら詳細欄を空にし、選択の有無でボタンの有効／無効を切り替える
         * @returns {void}
         */
        function updateDetailRows() {
            var selectedAction = getSelectedAction();
            var keyStringsText = selectedAction ? buildKeyStringsText(selectedAction) : "";
            if (selectedAction && keyStringsText === "") {
                keyStringsRow.valueField.text = getLabel(LABELS.message.noKeyStrings);
            } else {
                keyStringsRow.valueField.text = keyStringsText;
            }
            invokeCodeRow.valueField.text = selectedAction ? buildInvokeCode(selectedAction) : "";
            keyStringsRow.btnCopy.enabled = (keyStringsText !== "");
            invokeCodeRow.btnCopy.enabled = (selectedAction !== null);
        }

        /**
         * 選択中のメニューの絞り込み条件を返す
         * @returns {Object|null} MENU_FILTERS の要素。［すべて］なら null
         */
        function getSelectedMenuFilter() {
            var menuIndex = filterControls.menuDropdown.selection.index;
            return (menuIndex === 0) ? null : MENU_FILTERS[menuIndex - 1];
        }

        /**
         * メニューを変えたら、エリアの候補を絞り直して一覧を作り直す
         * @returns {void}
         */
        function handleMenuChange() {
            /* 項目の入れ替えで onChange が走らないよう外しておく / Detach so refilling does not fire onChange */
            filterControls.areaDropdown.onChange = null;
            fillAreaDropdown(filterControls.areaDropdown, filterAreaNames(allAreaNames, getSelectedMenuFilter()));
            filterControls.areaDropdown.onChange = handleFilterChange;
            refreshList(null);
        }

        /**
         * 選択中の並び順のキーを返す
         * @returns {string} "area" / "name" / "id"
         */
        function getSelectedSortKey() {
            for (var i = 0; i < filterControls.sortRadios.length; i++) {
                if (filterControls.sortRadios[i].value) {
                    return SORT_KEYS[i];
                }
            }
            return SORT_KEYS[DEFAULT_SORT_INDEX];
        }

        /**
         * 絞り込み・並べ替えの結果で一覧を作り直す
         * @param {Object} [listProgress] - 進行状況の表示先（最初の作成時だけ渡す）
         * @returns {void}
         */
        function refreshList(listProgress) {
            var areaSelection = filterControls.areaDropdown.selection;
            var areaName = (areaSelection.index === 0) ? null : areaSelection.text;
            visibleActions = filterActions(allActions, filterControls.searchInput.text, areaName, getSelectedMenuFilter());
            sortActions(visibleActions, getSelectedSortKey());

            if (listProgress) {
                listProgress.reset(LABELS.progress.building, visibleActions.length);
            }
            actionList.removeAll();
            for (var i = 0; i < visibleActions.length; i++) {
                var listItem = actionList.add("item", visibleActions[i].area);
                listItem.subItems[0].text = visibleActions[i].name;
                if (listProgress && i % PROGRESS_STEP === 0) {
                    listProgress.update(i);
                }
            }
            mainDialog.text = buildDialogTitle(visibleActions.length);
            /* 1件に絞れたら選択して詳細を出す / Select it when only one action is left */
            if (visibleActions.length === 1) {
                actionList.selection = 0;
            }
            updateDetailRows();
        }

        /**
         * 絞り込み条件の変更時に一覧を作り直す（進行状況は出さない）
         * @returns {void}
         */
        function handleFilterChange() {
            refreshList(null);
        }

        /* 全件の作り直しは重いので確定時に絞り込む / Rebuilding is heavy, so filter on commit */
        filterControls.searchInput.onChange = handleFilterChange;
        filterControls.areaDropdown.onChange = handleFilterChange;
        filterControls.menuDropdown.onChange = handleMenuChange;
        for (var i = 0; i < filterControls.sortRadios.length; i++) {
            filterControls.sortRadios[i].onClick = handleFilterChange;
        }
        actionList.onChange = updateDetailRows;
        keyStringsRow.btnCopy.onClick = function () {
            copyToClipboard(buildKeyStringsText(getSelectedAction()));
        };
        invokeCodeRow.btnCopy.onClick = function () {
            copyToClipboard(buildInvokeCode(getSelectedAction()));
        };

        refreshList(progressWindow);
        progressWindow.close();
        mainDialog.show();
    }

    /* 読み込み3段階（名前・エリア・ID）/ Three loading steps (names, areas, IDs) */
    var loadingProgress = createProgressWindow(LABELS.progress.loading, 3);
    showMainDialog(collectActions(loadingProgress), loadingProgress);

}());
