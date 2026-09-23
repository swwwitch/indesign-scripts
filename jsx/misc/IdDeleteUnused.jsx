#target indesign

/*

### 概要

使われていないスタイル（段落・文字・オブジェクト・表・セル）、スウォッチ、親ページを削除します。
削除する前に一覧で確認でき、残したい項目はチェックを外せます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdDeleteUnused.md

### Overview

Deletes unused styles (paragraph, character, object, table, cell), swatches, and parent pages.
The candidates are listed before deletion so you can uncheck anything you want to keep.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdDeleteUnused.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdDeleteUnused";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-24";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdDeleteUnused.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdDeleteUnused.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */

    /* 確認リスト / Confirmation list */
    var CONFIRM_LIST_SIZE         = [560, 320];        /* リストの寸法 [幅,高さ] / list size */
    var CONFIRM_COLUMN_WIDTHS     = [170, 380];        /* 列幅（種類・名前） / column widths (kind, name) */
    var CONFIRM_COLUMN_WIDTHS_DOC = [150, 250, 150];   /* 列幅（種類・名前・ドキュメント） / column widths with document */

    /* ボタン / Buttons */
    var BUTTON_ROW_TOP_MARGIN = 10;          /* ボタン列の上余白 / top margin above the button row */

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
            title:   { ja: "未使用項目の削除", en: "Delete Unused Items" },
            confirm: { ja: "削除する項目の確認", en: "Confirm Items to Delete" }
        },
        panel: {
            scope:   { ja: "対象", en: "Target" },
            targets: { ja: "削除する項目", en: "Items to Delete" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            activeDocument: { ja: "このドキュメント", en: "This document" },
            allDocuments:   { ja: "すべてのドキュメント", en: "All documents" }
        },
        checkbox: {
            paragraphStyle:    { ja: "段落スタイル", en: "Paragraph styles" },
            characterStyle:    { ja: "文字スタイル", en: "Character styles" },
            objectStyle:       { ja: "オブジェクトスタイル", en: "Object styles" },
            tableStyle:        { ja: "表スタイル", en: "Table styles" },
            cellStyle:         { ja: "セルスタイル", en: "Cell styles" },
            swatch:            { ja: "スウォッチ", en: "Swatches" },
            masterSpread:      { ja: "親ページ", en: "Parent pages" },
            removeEmptyGroups: { ja: "空のスタイルグループを削除", en: "Delete empty style groups" }
        },
        kindName: {
            paragraphStyle: { ja: "段落スタイル", en: "Paragraph style" },
            characterStyle: { ja: "文字スタイル", en: "Character style" },
            objectStyle:    { ja: "オブジェクトスタイル", en: "Object style" },
            tableStyle:     { ja: "表スタイル", en: "Table style" },
            cellStyle:      { ja: "セルスタイル", en: "Cell style" },
            swatch:         { ja: "スウォッチ", en: "Swatch" },
            masterSpread:   { ja: "親ページ", en: "Parent page" },
            styleGroup:     { ja: "スタイルグループ", en: "Style group" }
        },
        columnTitle: {
            kind:     { ja: "種類", en: "Kind" },
            name:     { ja: "名前", en: "Name" },
            document: { ja: "ドキュメント", en: "Document" }
        },
        tooltip: {
            paragraphStyle: {
                ja: "テキストにも、他のスタイルの「基準」「次のスタイル」、オブジェクトスタイル、セルスタイル、目次スタイル、脚注の設定にも使われていない段落スタイルを削除します。［ ］で囲まれた既定のスタイルは残します。",
                en: "Deletes paragraph styles not used in text, as Based On / Next Style of other styles, or in object, cell, TOC styles and footnote options. Default styles in [ ] are kept."
            },
            characterStyle: {
                ja: "テキストにも、他のスタイルの「基準」、先頭文字スタイル・正規表現スタイル・行スタイル、箇条書き、ドロップキャップ、目次スタイル、脚注の設定にも使われていない文字スタイルを削除します。［なし］は残します。",
                en: "Deletes character styles not used in text, as Based On, in nested / GREP / line styles, bullets and numbering, drop caps, TOC styles, or footnote options. [None] is kept."
            },
            objectStyle: {
                ja: "どのオブジェクトにも、他のスタイルの「基準」にも、新規オブジェクトの既定にも使われていないオブジェクトスタイルを削除します。［ ］で囲まれた既定のスタイルは残します。",
                en: "Deletes object styles not applied to any object, not used as Based On, and not set as the default for new objects. Default styles in [ ] are kept."
            },
            tableStyle: {
                ja: "どの表にも、他の表スタイルの「基準」にも使われていない表スタイルを削除します。［ ］で囲まれた既定のスタイルは残します。",
                en: "Deletes table styles not applied to any table and not used as Based On. Default styles in [ ] are kept."
            },
            cellStyle: {
                ja: "どのセルにも、他のセルスタイルの「基準」、表スタイルの領域（ヘッダー行・本文行など）にも使われていないセルスタイルを削除します。［なし］は残します。",
                en: "Deletes cell styles not applied to any cell, not used as Based On, and not used in table style regions (header rows, body rows, etc.). [None] is kept."
            },
            swatch: {
                ja: "使われていないスウォッチを削除します。グラデーションの分岐点や濃淡の元になっているカラー、名前のないカラー、削除できない既定のスウォッチは残します。",
                en: "Deletes unused swatches. Colors used in gradient stops or as the base of tints, unnamed colors, and default swatches that cannot be deleted are kept."
            },
            masterSpread: {
                ja: "どのページにも、他の親ページの「基準」にも使われていない親ページを削除します。",
                en: "Deletes parent pages not applied to any page and not used as the basis of another parent page."
            },
            removeEmptyGroups: {
                ja: "段落・文字・オブジェクト・表・セルのスタイルグループのうち、中身が空のもの（削除で空になるものを含む）を削除します。",
                en: "Deletes paragraph, character, object, table, and cell style groups that are empty, including those emptied by this deletion."
            },
            optionClick: {
                ja: "option（Alt）+クリック：この項目だけをON。もう一度押すとすべてをON",
                en: "Option (Alt)-click: check only this item. Do it again to check all"
            },
            confirmList: {
                ja: "チェックを外した項目は削除しません。行をダブルクリックしてもチェックを切り替えられます。",
                en: "Unchecked items are kept. Double-click a row to toggle its check."
            }
        },
        button: {
            cancel:     { ja: "キャンセル", en: "Cancel" },
            ok:         { ja: "OK", en: "OK" },
            remove:     { ja: "削除", en: "Delete" },
            checkAll:   { ja: "すべてON", en: "Check All" },
            uncheckAll: { ja: "すべてOFF", en: "Uncheck All" }
        },
        message: {
            confirmCount: {
                ja: "%1個の未使用項目が見つかりました。チェックを外した項目は残します。",
                en: "Found %1 unused items. Unchecked items will be kept."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noneFound:  { ja: "未使用の項目は見つかりませんでした。", en: "No unused items were found." },
            result:     { ja: "削除しました。", en: "Deleted." },
            docCount:   { ja: "%1個のドキュメントを処理しました。", en: "Processed %1 documents." },
            countLine:  { ja: "%1：%2個", en: "%1: %2" },
            totalLine:  { ja: "合計：%1個", en: "Total: %1" },
            keptLine: {
                ja: "チェックを外した項目から参照されているなどの理由で残した項目：%1個",
                en: "Kept because still referenced by unchecked items, etc.: %1"
            }
        },
        undoName: { ja: "未使用項目の削除", en: "Delete Unused Items" }
    };

    /**
     * ラベル定義から現在のUI言語の文字列を取り出す
     * @param {{ja: string, en: string}} labelSet - 言語別のラベル定義
     * @returns {string} 現在のUI言語の文字列
     */
    function getLabel(labelSet) {
        return labelSet[uiLang] || labelSet.en;
    }

    /**
     * ラベル内のプレースホルダー（%1, %2 …）を値で置き換える
     * @param {string} template - プレースホルダーを含む文字列
     * @param {Array<string>} values - 差し込む値
     * @returns {string} 置き換え後の文字列
     */
    function formatLabel(template, values) {
        var text = template;
        for (var i = 0; i < values.length; i++) {
            text = text.split("%" + (i + 1)).join(String(values[i]));
        }
        return text;
    }

    // =========================================
    // 削除対象の定義 / Target definitions
    // =========================================

    /* 削除対象の種類（ダイアログ・確認リスト・結果の並び順） / Target kinds, in display order */
    var TARGET_KEYS = ["paragraphStyle", "characterStyle", "objectStyle", "tableStyle", "cellStyle", "swatch", "masterSpread"];

    /* 初回の初期値 / Defaults on first run */
    var DEFAULT_CHECKED_TARGETS = { swatch: true };
    var DEFAULT_REMOVE_EMPTY_GROUPS = false;

    /* スタイルの種類ごとのコレクション名とテキスト検索のプロパティ名
       / Collection names and find property per style kind */
    var STYLE_KINDS = {
        paragraphStyle: { allStyles: "allParagraphStyles", styles: "paragraphStyles", groups: "paragraphStyleGroups", findProperty: "appliedParagraphStyle" },
        characterStyle: { allStyles: "allCharacterStyles", styles: "characterStyles", groups: "characterStyleGroups", findProperty: "appliedCharacterStyle" },
        objectStyle:    { allStyles: "allObjectStyles",    styles: "objectStyles",    groups: "objectStyleGroups",    findProperty: null },
        tableStyle:     { allStyles: "allTableStyles",     styles: "tableStyles",     groups: "tableStyleGroups",     findProperty: null },
        cellStyle:      { allStyles: "allCellStyles",      styles: "cellStyles",      groups: "cellStyleGroups",      findProperty: null }
    };
    var STYLE_KIND_KEYS = ["paragraphStyle", "characterStyle", "objectStyle", "tableStyle", "cellStyle"];

    /* 確認リストでのスタイルグループの種類名 / Kind key for style groups in the list */
    var GROUP_KIND_KEY = "styleGroup";

    // =========================================
    // 設定の記憶 / Saved settings
    // =========================================

    /* InDesignには任意の値を残す環境設定APIが無いので、設定ファイルに key=value で書き出す
       / InDesign has no scriptable preference store, so settings go to a key=value file */
    var PREFS_FILE_NAME              = "IdDeleteUnused-prefs.txt";
    var PREF_KEY_ALL_DOCUMENTS       = "allDocuments";
    var PREF_KEY_TARGETS             = "targets";
    var PREF_KEY_REMOVE_EMPTY_GROUPS = "removeEmptyGroups";

    /**
     * 設定ファイルを返す
     * @returns {File} 設定ファイル
     */
    function getPrefsFile() {
        return File(Folder.userData.fsName + "/" + PREFS_FILE_NAME);
    }

    /**
     * 前回のダイアログの状態を読み出す。記録が無ければ初期値を返す
     * @returns {{allDocuments: boolean, targets: Object<string, boolean>, removeEmptyGroups: boolean}} ダイアログの状態
     */
    function loadDialogState() {
        var dialogState = { allDocuments: false, targets: DEFAULT_CHECKED_TARGETS, removeEmptyGroups: DEFAULT_REMOVE_EMPTY_GROUPS };

        var prefsFile = getPrefsFile();
        prefsFile.encoding = "UTF-8";
        if (!prefsFile.exists || !prefsFile.open("r")) return dialogState;
        var lines = prefsFile.read().split("\n");
        prefsFile.close();

        var prefs = {};
        for (var i = 0; i < lines.length; i++) {
            var separatorIndex = lines[i].indexOf("=");
            if (separatorIndex > 0) prefs[lines[i].substring(0, separatorIndex)] = lines[i].substring(separatorIndex + 1);
        }

        if (prefs.hasOwnProperty(PREF_KEY_ALL_DOCUMENTS)) dialogState.allDocuments = prefs[PREF_KEY_ALL_DOCUMENTS] === "true";
        if (prefs.hasOwnProperty(PREF_KEY_REMOVE_EMPTY_GROUPS)) dialogState.removeEmptyGroups = prefs[PREF_KEY_REMOVE_EMPTY_GROUPS] === "true";
        if (prefs.hasOwnProperty(PREF_KEY_TARGETS)) {
            dialogState.targets = {};
            var savedKeys = prefs[PREF_KEY_TARGETS].split(",");
            for (var j = 0; j < savedKeys.length; j++) dialogState.targets[savedKeys[j]] = true;
        }
        return dialogState;
    }

    /**
     * ダイアログの状態を設定ファイルに書き出す
     * @param {{allDocuments: boolean, targets: Object<string, boolean>, removeEmptyGroups: boolean}} dialogState - ダイアログの状態
     * @returns {void}
     */
    function saveDialogState(dialogState) {
        var checkedKeys = [];
        for (var i = 0; i < TARGET_KEYS.length; i++) {
            if (dialogState.targets[TARGET_KEYS[i]]) checkedKeys.push(TARGET_KEYS[i]);
        }
        var lines = [
            PREF_KEY_ALL_DOCUMENTS + "=" + dialogState.allDocuments,
            PREF_KEY_TARGETS + "=" + checkedKeys.join(","),
            PREF_KEY_REMOVE_EMPTY_GROUPS + "=" + dialogState.removeEmptyGroups
        ];

        /* 保存できなくても削除は続けられるので、書き出せたかどうかは見ない / A failed save must not break the run */
        var prefsFile = getPrefsFile();
        prefsFile.encoding = "UTF-8";
        if (!prefsFile.open("w")) return;
        prefsFile.write(lines.join("\n"));
        prefsFile.close();
    }

    // =========================================
    // ダイアログ共通 / Dialog helpers
    // =========================================

    /**
     * ウィンドウの共通設定を適用する
     * @param {Window} win - 対象ウィンドウ
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
     * @param {Panel} panel - 対象パネル
     * @returns {void}
     */
    function setupPanel(panel) {
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.margins = PANEL_MARGINS;
        panel.spacing = PANEL_SPACING;
    }

    /**
     * ボタン列を追加する（左側は任意、右側にキャンセルと確定）
     * @param {Window} parent - 追加先のウィンドウ
     * @param {Array<{ja: string, en: string}>} leftLabels - 左側に並べるボタンのラベル
     * @param {{ja: string, en: string}} okLabel - 確定ボタンのラベル
     * @returns {{left: Array<Button>, ok: Button}} 左側のボタンと確定ボタン
     */
    function addButtonRow(parent, leftLabels, okLabel) {
        // メイングループ（横並び） / Main group (horizontal layout)
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        // 左側グループ / Left-side button group
        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var leftButtons = [];
        for (var i = 0; i < leftLabels.length; i++) {
            leftButtons.push(btnLeftGroup.add("button", undefined, getLabel(leftLabels[i])));
        }

        // スペーサー（伸縮）/ Spacer (stretchable)
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        // 右側グループ / Right-side button group
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel(okLabel), { name: "ok" });
        return { left: leftButtons, ok: btnOK };
    }

    // =========================================
    // 設定ダイアログ / Settings dialog
    // =========================================

    /**
     * 対象ドキュメントのラジオボタンを並べる
     * @param {Window} parent - 追加先のウィンドウ
     * @param {boolean} allDocuments - 「すべてのドキュメント」を選んだ状態にするか
     * @returns {RadioButton} 「すべてのドキュメント」のラジオボタン
     */
    function addScopeRadios(parent, allDocuments) {
        var scopePanel = parent.add("panel", undefined, getLabel(LABELS.panel.scope));
        setupPanel(scopePanel);

        var activeDocumentRadio = scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.activeDocument));
        var allDocumentsRadio = scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.allDocuments));
        activeDocumentRadio.value = !allDocuments;
        allDocumentsRadio.value = allDocuments;
        return allDocumentsRadio;
    }

    /**
     * 削除する項目のチェックボックスを並べる
     * @param {Window} parent - 追加先のウィンドウ
     * @param {Object<string, boolean>} checkedTargets - ONにしておく種類
     * @returns {Object<string, Checkbox>} 種類ごとのチェックボックス
     */
    function addTargetCheckboxes(parent, checkedTargets) {
        var targetPanel = parent.add("panel", undefined, getLabel(LABELS.panel.targets));
        setupPanel(targetPanel);

        var targetCheckboxes = {};
        for (var i = 0; i < TARGET_KEYS.length; i++) {
            var targetKey = TARGET_KEYS[i];
            var targetCheckbox = targetPanel.add("checkbox", undefined, getLabel(LABELS.checkbox[targetKey]));
            targetCheckbox.value = checkedTargets[targetKey] === true;
            targetCheckbox.helpTip = getLabel(LABELS.tooltip[targetKey]) + "\n\n" + getLabel(LABELS.tooltip.optionClick);
            targetCheckboxes[targetKey] = targetCheckbox;
        }
        return targetCheckboxes;
    }

    /**
     * オプションのチェックボックスを並べる
     * @param {Window} parent - 追加先のウィンドウ
     * @param {boolean} removeEmptyGroups - 「空のスタイルグループを削除」の初期値
     * @returns {Checkbox} 「空のスタイルグループを削除」のチェックボックス
     */
    function addOptionCheckboxes(parent, removeEmptyGroups) {
        var optionPanel = parent.add("panel", undefined, getLabel(LABELS.panel.options));
        setupPanel(optionPanel);

        var emptyGroupsCheckbox = optionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.removeEmptyGroups));
        emptyGroupsCheckbox.value = removeEmptyGroups;
        emptyGroupsCheckbox.helpTip = getLabel(LABELS.tooltip.removeEmptyGroups);
        return emptyGroupsCheckbox;
    }

    /**
     * option（Alt）+クリックで、押した項目だけをONにする
     * すでにそれだけがONの状態で押したときは、すべてをONに戻す
     * onClick の時点で押した項目の値は反転済みなので、反転前の状態から判定する
     * @param {Object<string, Checkbox>} targetCheckboxes - 種類ごとのチェックボックス
     * @param {string} clickedKey - 押された項目の種類
     * @returns {void}
     */
    function applyOptionClick(targetCheckboxes, clickedKey) {
        var wasChecked = !targetCheckboxes[clickedKey].value;
        var othersUnchecked = true;
        for (var i = 0; i < TARGET_KEYS.length; i++) {
            if (TARGET_KEYS[i] !== clickedKey && targetCheckboxes[TARGET_KEYS[i]].value) othersUnchecked = false;
        }
        var checkAll = wasChecked && othersUnchecked;
        for (var j = 0; j < TARGET_KEYS.length; j++) {
            targetCheckboxes[TARGET_KEYS[j]].value = checkAll || TARGET_KEYS[j] === clickedKey;
        }
    }

    /**
     * 設定ダイアログを表示して、対象ドキュメント・削除する項目・オプションを選ばせる
     * 開くときは前回の状態を復元し、OK で閉じたら保存する
     * @returns {{allDocuments: boolean, targets: Object<string, boolean>, removeEmptyGroups: boolean}|null} 選んだ内容。キャンセル時は null
     */
    function showSettingsDialog() {
        var savedState = loadDialogState();

        var settingsDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(settingsDialog);

        var allDocumentsRadio = addScopeRadios(settingsDialog, savedState.allDocuments);
        var targetCheckboxes = addTargetCheckboxes(settingsDialog, savedState.targets);
        var emptyGroupsCheckbox = addOptionCheckboxes(settingsDialog, savedState.removeEmptyGroups);
        var btnOK = addButtonRow(settingsDialog, [], LABELS.button.ok).ok;

        /* 削除する項目が1つもONでなければ OK を押せなくする / Disable OK when no kind is checked */
        function updateOKButton() {
            var hasTarget = false;
            for (var i = 0; i < TARGET_KEYS.length; i++) {
                if (targetCheckboxes[TARGET_KEYS[i]].value) hasTarget = true;
            }
            btnOK.enabled = hasTarget;
        }
        /**
         * チェックボックスのクリック処理を作る（option 併用の判定と OK ボタンの更新）
         * @param {string} targetKey - 対象の種類
         * @returns {Function} onClick に渡す関数
         */
        function createCheckboxClickHandler(targetKey) {
            return function () {
                if (ScriptUI.environment.keyboardState.altKey) applyOptionClick(targetCheckboxes, targetKey);
                updateOKButton();
            };
        }
        for (var i = 0; i < TARGET_KEYS.length; i++) {
            targetCheckboxes[TARGET_KEYS[i]].onClick = createCheckboxClickHandler(TARGET_KEYS[i]);
        }
        updateOKButton();

        if (settingsDialog.show() !== 1) return null;

        var selectedTargets = {};
        for (var j = 0; j < TARGET_KEYS.length; j++) {
            selectedTargets[TARGET_KEYS[j]] = targetCheckboxes[TARGET_KEYS[j]].value;
        }
        var dialogState = {
            allDocuments: allDocumentsRadio.value,
            targets: selectedTargets,
            removeEmptyGroups: emptyGroupsCheckbox.value
        };
        saveDialogState(dialogState);
        return dialogState;
    }

    // =========================================
    // 確認ダイアログ / Confirmation dialog
    // =========================================

    /**
     * 確認リストの種類の列に出す文字列を返す
     * @param {{kind: string, styleKind: string}} candidate - 削除候補
     * @returns {string} 種類名。スタイルグループは元のスタイルの種類を添える
     */
    function getKindText(candidate) {
        if (candidate.kind !== GROUP_KIND_KEY) return getLabel(LABELS.kindName[candidate.kind]);
        return getLabel(LABELS.kindName[GROUP_KIND_KEY]) + " (" + getLabel(LABELS.kindName[candidate.styleKind]) + ")";
    }

    /**
     * 削除候補を一覧で見せ、チェックを付けたものだけを返す
     * @param {Array<Object>} candidates - 削除候補
     * @param {boolean} showDocumentColumn - ドキュメントの列を出すか
     * @returns {Array<Object>|null} チェックを付けた候補。キャンセル時は null
     */
    function showConfirmDialog(candidates, showDocumentColumn) {
        var confirmDialog = new Window("dialog", getLabel(LABELS.dialog.confirm));
        setupWindow(confirmDialog);

        confirmDialog.add("statictext", undefined, formatLabel(getLabel(LABELS.message.confirmCount), [candidates.length]));

        var columnTitles = [getLabel(LABELS.columnTitle.kind), getLabel(LABELS.columnTitle.name)];
        if (showDocumentColumn) columnTitles.push(getLabel(LABELS.columnTitle.document));
        var candidateList = confirmDialog.add("listbox", undefined, [], {
            numberOfColumns: columnTitles.length,
            showHeaders: true,
            columnTitles: columnTitles,
            columnWidths: showDocumentColumn ? CONFIRM_COLUMN_WIDTHS_DOC : CONFIRM_COLUMN_WIDTHS
        });
        candidateList.preferredSize = CONFIRM_LIST_SIZE;
        candidateList.helpTip = getLabel(LABELS.tooltip.confirmList);

        for (var i = 0; i < candidates.length; i++) {
            var listItem = candidateList.add("item", getKindText(candidates[i]));
            listItem.subItems[0].text = candidates[i].name;
            if (showDocumentColumn) listItem.subItems[1].text = candidates[i].doc.name;
            listItem.checked = true;
        }

        candidateList.onDoubleClick = function () {
            if (candidateList.selection) candidateList.selection.checked = !candidateList.selection.checked;
        };

        var buttons = addButtonRow(confirmDialog, [LABELS.button.checkAll, LABELS.button.uncheckAll], LABELS.button.remove);
        /**
         * リストのチェックをまとめて切り替える
         * @param {boolean} checked - ONにするか
         * @returns {Function} onClick に渡す関数
         */
        function createCheckAllHandler(checked) {
            return function () {
                for (var j = 0; j < candidateList.items.length; j++) candidateList.items[j].checked = checked;
            };
        }
        buttons.left[0].onClick = createCheckAllHandler(true);
        buttons.left[1].onClick = createCheckAllHandler(false);

        if (confirmDialog.show() !== 1) return null;

        var confirmedCandidates = [];
        for (var k = 0; k < candidateList.items.length; k++) {
            if (candidateList.items[k].checked) confirmedCandidates.push(candidates[k]);
        }
        return confirmedCandidates;
    }

    // =========================================
    // 参照の収集 / Reference collection
    // =========================================

    /**
     * 種類ごとの空の記録を作る
     * @returns {Object<string, Object<string, boolean>>} 種類をキーにした、IDの記録
     */
    function createKindMap() {
        var kindMap = {};
        for (var i = 0; i < TARGET_KEYS.length; i++) kindMap[TARGET_KEYS[i]] = {};
        return kindMap;
    }

    /**
     * 値が指定クラスのオブジェクトなら、使用中として記録する
     * 未設定のときは文字列や NothingEnum が返るので、クラスで見分ける
     * @param {Object<string, boolean>} referenceMap - IDをキーにした使用中の記録
     * @param {*} value - プロパティ値
     * @param {Function} domClass - ParagraphStyle などのクラス
     * @returns {void}
     */
    function markReferenced(referenceMap, value, domClass) {
        if (value instanceof domClass) referenceMap[value.id] = true;
    }

    /**
     * ルートスタイル（［段落スタイルなし］や［なし］）を除いたスタイルの一覧を返す
     * ルートスタイルは basedOn などを読むだけで「ルートスタイルに対する無効な要求」の例外になる
     * @param {Document} doc - 対象ドキュメント
     * @param {string} styleKind - STYLE_KINDS のキー
     * @returns {Array<Object>} ルートスタイルを除いたスタイル
     */
    function getNonRootStyles(doc, styleKind) {
        var allStyles = doc[STYLE_KINDS[styleKind].allStyles];
        var rootStyleId = doc[STYLE_KINDS[styleKind].styles][0].id;
        var styles = [];
        for (var i = 0; i < allStyles.length; i++) {
            if (allStyles[i].id !== rootStyleId) styles.push(allStyles[i]);
        }
        return styles;
    }

    /**
     * 表とセルに適用されている表スタイル・セルスタイルを記録する
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, Object<string, boolean>>} directUse - 種類ごとの使用中の記録
     * @returns {void}
     */
    function markTableUse(doc, directUse) {
        for (var i = 0; i < doc.stories.length; i++) {
            var tables = doc.stories[i].tables;
            for (var j = 0; j < tables.length; j++) {
                markReferenced(directUse.tableStyle, tables[j].appliedTableStyle, TableStyle);
                if (tables[j].cells.length === 0) continue;
                var cellStyleList = tables[j].cells.everyItem().appliedCellStyle;
                if (!(cellStyleList instanceof Array)) cellStyleList = [cellStyleList];
                for (var k = 0; k < cellStyleList.length; k++) {
                    markReferenced(directUse.cellStyle, cellStyleList[k], CellStyle);
                }
            }
        }
    }

    /**
     * 目次スタイルから参照されている段落・文字スタイルを記録する
     * 目次の項目は、対象の段落スタイルを名前で持つ
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, Object<string, boolean>>} directUse - 種類ごとの使用中の記録
     * @returns {void}
     */
    function markTocUse(doc, directUse) {
        var tocEntryNames = {};
        for (var i = 0; i < doc.tocStyles.length; i++) {
            var tocStyle = doc.tocStyles[i];
            markReferenced(directUse.paragraphStyle, tocStyle.titleStyle, ParagraphStyle);
            for (var j = 0; j < tocStyle.tocStyleEntries.length; j++) {
                var tocEntry = tocStyle.tocStyleEntries[j];
                tocEntryNames[tocEntry.name] = true;
                markReferenced(directUse.paragraphStyle, tocEntry.formatStyle, ParagraphStyle);
                markReferenced(directUse.characterStyle, tocEntry.pageNumberStyle, CharacterStyle);
                markReferenced(directUse.characterStyle, tocEntry.separatorStyle, CharacterStyle);
            }
        }
        var paragraphStyles = doc.allParagraphStyles;
        for (var k = 0; k < paragraphStyles.length; k++) {
            if (tocEntryNames[paragraphStyles[k].name]) directUse.paragraphStyle[paragraphStyles[k].id] = true;
        }
    }

    /**
     * 削除候補に左右されない使用状況を集める（ページアイテム・表・ページへの適用、ドキュメント設定）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object<string, Object<string, boolean>>} 種類ごとの使用中のID
     */
    function collectDirectUse(doc) {
        var directUse = createKindMap();

        var pageItems = doc.allPageItems;
        for (var i = 0; i < pageItems.length; i++) {
            markReferenced(directUse.objectStyle, pageItems[i].appliedObjectStyle, ObjectStyle);
        }
        var itemDefaults = doc.pageItemDefaults;
        markReferenced(directUse.objectStyle, itemDefaults.appliedTextObjectStyle, ObjectStyle);
        markReferenced(directUse.objectStyle, itemDefaults.appliedGraphicObjectStyle, ObjectStyle);
        markReferenced(directUse.objectStyle, itemDefaults.appliedGridObjectStyle, ObjectStyle);

        markTableUse(doc, directUse);
        markTocUse(doc, directUse);

        markReferenced(directUse.paragraphStyle, doc.footnoteOptions.footnoteTextStyle, ParagraphStyle);
        markReferenced(directUse.characterStyle, doc.footnoteOptions.footnoteMarkerStyle, CharacterStyle);
        markReferenced(directUse.paragraphStyle, doc.textDefaults.appliedParagraphStyle, ParagraphStyle);
        markReferenced(directUse.characterStyle, doc.textDefaults.appliedCharacterStyle, CharacterStyle);

        for (var j = 0; j < doc.pages.length; j++) {
            markReferenced(directUse.masterSpread, doc.pages[j].appliedMaster, MasterSpread);
        }
        return directUse;
    }

    /**
     * 段落スタイルから、他の段落スタイルと文字スタイルへの参照を記録する
     * @param {ParagraphStyle} paragraphStyle - 参照元の段落スタイル
     * @param {Object<string, Object<string, boolean>>} crossRefs - 種類ごとの参照の記録
     * @returns {void}
     */
    function markParagraphStyleRefs(paragraphStyle, crossRefs) {
        markReferenced(crossRefs.paragraphStyle, paragraphStyle.basedOn, ParagraphStyle);
        /* 自分自身を「次のスタイル」にしているのは参照に数えない / Self as Next Style does not count */
        if (paragraphStyle.nextStyle instanceof ParagraphStyle && paragraphStyle.nextStyle.id !== paragraphStyle.id) {
            crossRefs.paragraphStyle[paragraphStyle.nextStyle.id] = true;
        }
        markReferenced(crossRefs.characterStyle, paragraphStyle.bulletsCharacterStyle, CharacterStyle);
        markReferenced(crossRefs.characterStyle, paragraphStyle.numberingCharacterStyle, CharacterStyle);
        markReferenced(crossRefs.characterStyle, paragraphStyle.dropCapStyle, CharacterStyle);

        var nestedCollections = [paragraphStyle.nestedStyles, paragraphStyle.nestedGrepStyles, paragraphStyle.nestedLineStyles];
        for (var i = 0; i < nestedCollections.length; i++) {
            for (var j = 0; j < nestedCollections[i].length; j++) {
                markReferenced(crossRefs.characterStyle, nestedCollections[i][j].appliedCharacterStyle, CharacterStyle);
            }
        }
    }

    /**
     * 表スタイルから、他の表スタイルと各領域のセルスタイルへの参照を記録する
     * @param {TableStyle} tableStyle - 参照元の表スタイル
     * @param {Object<string, Object<string, boolean>>} crossRefs - 種類ごとの参照の記録
     * @returns {void}
     */
    function markTableStyleRefs(tableStyle, crossRefs) {
        markReferenced(crossRefs.tableStyle, tableStyle.basedOn, TableStyle);
        var regionProperties = ["headerRegionCellStyle", "footerRegionCellStyle", "bodyRegionCellStyle", "leftColumnRegionCellStyle", "rightColumnRegionCellStyle"];
        for (var i = 0; i < regionProperties.length; i++) {
            markReferenced(crossRefs.cellStyle, tableStyle[regionProperties[i]], CellStyle);
        }
    }

    /**
     * スタイル同士の参照を集める。ignoredIds にある参照元（削除予定のもの）は数えない
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, boolean>} ignoredIds - 削除予定のID
     * @param {Object<string, Object<string, boolean>>} crossRefs - 種類ごとの参照の記録
     * @returns {void}
     */
    function markStyleCrossRefs(doc, ignoredIds, crossRefs) {
        var paragraphStyles = getNonRootStyles(doc, "paragraphStyle");
        for (var i = 0; i < paragraphStyles.length; i++) {
            if (!ignoredIds[paragraphStyles[i].id]) markParagraphStyleRefs(paragraphStyles[i], crossRefs);
        }
        var characterStyles = getNonRootStyles(doc, "characterStyle");
        for (var j = 0; j < characterStyles.length; j++) {
            if (!ignoredIds[characterStyles[j].id]) markReferenced(crossRefs.characterStyle, characterStyles[j].basedOn, CharacterStyle);
        }
        var objectStyles = getNonRootStyles(doc, "objectStyle");
        for (var k = 0; k < objectStyles.length; k++) {
            if (ignoredIds[objectStyles[k].id]) continue;
            markReferenced(crossRefs.objectStyle, objectStyles[k].basedOn, ObjectStyle);
            markReferenced(crossRefs.paragraphStyle, objectStyles[k].appliedParagraphStyle, ParagraphStyle);
        }
        var tableStyles = getNonRootStyles(doc, "tableStyle");
        for (var m = 0; m < tableStyles.length; m++) {
            if (!ignoredIds[tableStyles[m].id]) markTableStyleRefs(tableStyles[m], crossRefs);
        }
        var cellStyles = getNonRootStyles(doc, "cellStyle");
        for (var n = 0; n < cellStyles.length; n++) {
            if (ignoredIds[cellStyles[n].id]) continue;
            markReferenced(crossRefs.cellStyle, cellStyles[n].basedOn, CellStyle);
            markReferenced(crossRefs.paragraphStyle, cellStyles[n].appliedParagraphStyle, ParagraphStyle);
        }
    }

    /**
     * スタイル・スウォッチ・親ページ同士の参照を集める。ignoredIds にある参照元（削除予定のもの）は数えない
     * スウォッチはグラデーションの分岐点と濃淡の元のカラー。unusedSwatches にはこれらが含まれ、
     * 消すとグラデーションや濃淡が変わってしまう
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, boolean>} ignoredIds - 削除予定のID
     * @returns {Object<string, Object<string, boolean>>} 種類ごとの参照されているID
     */
    function collectCrossRefs(doc, ignoredIds) {
        var crossRefs = createKindMap();
        markStyleCrossRefs(doc, ignoredIds, crossRefs);

        for (var i = 0; i < doc.gradients.length; i++) {
            if (ignoredIds[doc.gradients[i].id]) continue;
            var gradientStops = doc.gradients[i].gradientStops;
            for (var j = 0; j < gradientStops.length; j++) {
                markReferenced(crossRefs.swatch, gradientStops[j].stopColor, Swatch);
            }
        }
        for (var k = 0; k < doc.tints.length; k++) {
            if (!ignoredIds[doc.tints[k].id]) markReferenced(crossRefs.swatch, doc.tints[k].baseColor, Swatch);
        }

        for (var m = 0; m < doc.masterSpreads.length; m++) {
            if (ignoredIds[doc.masterSpreads[m].id]) continue;
            var masterPages = doc.masterSpreads[m].pages;
            for (var n = 0; n < masterPages.length; n++) {
                markReferenced(crossRefs.masterSpread, masterPages[n].appliedMaster, MasterSpread);
            }
        }
        return crossRefs;
    }

    // =========================================
    // 削除候補の洗い出し / Finding candidates
    // =========================================

    /**
     * テキスト検索の範囲を、非表示・ロック・マスター・脚注まで広げる
     * @returns {Object<string, boolean>} 変更前の設定
     */
    function widenFindScope() {
        var findOptions = app.findChangeTextOptions;
        var scopeKeys = ["includeHiddenLayers", "includeLockedLayersForFind", "includeLockedStoriesForFind", "includeMasterPages", "includeFootnotes"];
        var savedScope = {};
        for (var i = 0; i < scopeKeys.length; i++) {
            savedScope[scopeKeys[i]] = findOptions[scopeKeys[i]];
            findOptions[scopeKeys[i]] = true;
        }
        return savedScope;
    }

    /**
     * スタイルがテキストに適用されているかを検索で調べる
     * @param {Document} doc - 対象ドキュメント
     * @param {string} findProperty - "appliedParagraphStyle" または "appliedCharacterStyle"
     * @param {ParagraphStyle|CharacterStyle} style - 調べるスタイル
     * @returns {boolean} 適用されていれば true
     */
    function isAppliedToText(doc, findProperty, style) {
        app.findTextPreferences = NothingEnum.NOTHING;
        app.findTextPreferences[findProperty] = style;
        var isApplied = doc.findText().length > 0;
        app.findTextPreferences = NothingEnum.NOTHING;
        return isApplied;
    }

    /**
     * 選ばれた種類の、削除候補になりうる項目を並べる
     * ［ ］で囲まれた既定のスタイルと、名前のないカラーは最初から外す
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, boolean>} selectedTargets - 種類ごとのON/OFF
     * @returns {Array<{kind: string, item: Object, id: number, name: string, doc: Document}>} 候補になりうる項目
     */
    function listRemovableItems(doc, selectedTargets) {
        var removableItems = [];
        /**
         * 項目を候補の形にして追加する
         * @param {string} kind - 種類
         * @param {Object} item - DOMオブジェクト
         * @returns {void}
         */
        function pushItem(kind, item) {
            removableItems.push({ kind: kind, item: item, id: item.id, name: item.name, doc: doc });
        }

        for (var i = 0; i < STYLE_KIND_KEYS.length; i++) {
            var styleKind = STYLE_KIND_KEYS[i];
            if (!selectedTargets[styleKind]) continue;
            var styles = getNonRootStyles(doc, styleKind);
            for (var j = 0; j < styles.length; j++) {
                if (styles[j].name.charAt(0) !== "[") pushItem(styleKind, styles[j]);
            }
        }
        if (selectedTargets.swatch) {
            var unusedSwatchList = doc.unusedSwatches;
            for (var k = 0; k < unusedSwatchList.length; k++) {
                if (unusedSwatchList[k].name !== "") pushItem("swatch", unusedSwatchList[k]);
            }
        }
        if (selectedTargets.masterSpread) {
            for (var m = 0; m < doc.masterSpreads.length; m++) pushItem("masterSpread", doc.masterSpreads[m]);
        }
        return removableItems;
    }

    /**
     * 項目が使用中かを判定する。テキストへの適用は重いので最後に調べ、結果を控える
     * @param {Object} removableItem - listRemovableItems() の要素
     * @param {Object<string, Object<string, boolean>>} directUse - 削除候補に左右されない使用状況
     * @param {Object<string, Object<string, boolean>>} crossRefs - 項目同士の参照
     * @param {Object<string, boolean>} textUseCache - テキストへの適用の判定結果
     * @returns {boolean} 使用中なら true
     */
    function isInUse(removableItem, directUse, crossRefs, textUseCache) {
        var kind = removableItem.kind;
        if (directUse[kind][removableItem.id] || crossRefs[kind][removableItem.id]) return true;

        var findProperty = STYLE_KINDS[kind] ? STYLE_KINDS[kind].findProperty : null;
        if (!findProperty) return false;
        if (!textUseCache.hasOwnProperty(removableItem.id)) {
            textUseCache[removableItem.id] = isAppliedToText(removableItem.doc, findProperty, removableItem.item);
        }
        return textUseCache[removableItem.id];
    }

    /**
     * 空のスタイルグループを、子グループが先に来る順で集める
     * deletedIds にあるスタイルは削除済みとみなす
     * @param {Object} groupCollection - paragraphStyleGroups などのコレクション
     * @param {string} styleKind - STYLE_KINDS のキー
     * @param {Object<string, boolean>} deletedIds - 削除予定のID
     * @param {Array<Object>} emptyGroups - 見つかったグループの追加先
     * @returns {boolean} コレクション内のグループがすべて空なら true
     */
    function collectEmptyGroups(groupCollection, styleKind, deletedIds, emptyGroups) {
        var allEmpty = true;
        for (var i = 0; i < groupCollection.length; i++) {
            var styleGroup = groupCollection[i];
            var subgroupsEmpty = collectEmptyGroups(styleGroup[STYLE_KINDS[styleKind].groups], styleKind, deletedIds, emptyGroups);
            var styles = styleGroup[STYLE_KINDS[styleKind].styles];
            var stylesDeleted = true;
            for (var j = 0; j < styles.length; j++) {
                if (!deletedIds[styles[j].id]) stylesDeleted = false;
            }
            if (subgroupsEmpty && stylesDeleted) {
                emptyGroups.push(styleGroup);
            } else {
                allEmpty = false;
            }
        }
        return allEmpty;
    }

    /**
     * 空になるスタイルグループを削除候補の形で返す
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, boolean>} deletedIds - 削除予定のID
     * @returns {Array<Object>} スタイルグループの削除候補
     */
    function findEmptyGroupCandidates(doc, deletedIds) {
        var groupCandidates = [];
        for (var i = 0; i < STYLE_KIND_KEYS.length; i++) {
            var styleKind = STYLE_KIND_KEYS[i];
            var emptyGroups = [];
            collectEmptyGroups(doc[STYLE_KINDS[styleKind].groups], styleKind, deletedIds, emptyGroups);
            for (var j = 0; j < emptyGroups.length; j++) {
                groupCandidates.push({ kind: GROUP_KIND_KEY, styleKind: styleKind, item: emptyGroups[j], id: emptyGroups[j].id, name: emptyGroups[j].name, doc: doc });
            }
        }
        return groupCandidates;
    }

    /**
     * 1つのドキュメントの削除候補を洗い出す
     * 子スタイルや派生した親ページを消すと元が未使用になることがあるので、候補が増えなくなるまで繰り返す
     * @param {Document} doc - 対象ドキュメント
     * @param {{targets: Object<string, boolean>, removeEmptyGroups: boolean}} dialogState - 設定ダイアログの内容
     * @returns {Array<Object>} 削除候補（種類の並び順）
     */
    function findCandidates(doc, dialogState) {
        var removableItems = listRemovableItems(doc, dialogState.targets);
        var directUse = collectDirectUse(doc);
        var textUseCache = {};
        var candidateIds = {};

        var addedCount;
        do {
            addedCount = 0;
            var crossRefs = collectCrossRefs(doc, candidateIds);
            for (var i = 0; i < removableItems.length; i++) {
                if (candidateIds[removableItems[i].id]) continue;
                if (isInUse(removableItems[i], directUse, crossRefs, textUseCache)) continue;
                candidateIds[removableItems[i].id] = true;
                addedCount++;
            }
        } while (addedCount > 0);

        var candidates = [];
        for (var j = 0; j < removableItems.length; j++) {
            if (candidateIds[removableItems[j].id]) candidates.push(removableItems[j]);
        }
        if (dialogState.removeEmptyGroups) candidates = candidates.concat(findEmptyGroupCandidates(doc, candidateIds));
        return candidates;
    }

    /**
     * 対象ドキュメントすべての削除候補を洗い出す。テキスト検索の範囲は必ず元に戻す
     * @param {Array<Document>} targetDocs - 対象ドキュメント
     * @param {Object} dialogState - 設定ダイアログの内容
     * @returns {Array<Object>} 削除候補
     */
    function findAllCandidates(targetDocs, dialogState) {
        var candidates = [];
        var savedScope = widenFindScope();
        try {
            for (var i = 0; i < targetDocs.length; i++) {
                candidates = candidates.concat(findCandidates(targetDocs[i], dialogState));
            }
        } finally {
            /* 検索範囲はユーザーの設定なので、失敗しても戻す / Always restore the user's find scope */
            app.findChangeTextOptions.properties = savedScope;
        }
        return candidates;
    }

    // =========================================
    // 削除 / Deletion
    // =========================================

    /**
     * チェックを外した項目から参照されている候補を外す
     * 外した候補がさらに別の候補を参照していることがあるので、外すものが無くなるまで繰り返す
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<Object>} confirmedItems - チェックを付けた候補（スタイルグループを除く）
     * @returns {Array<Object>} 実際に削除する候補
     */
    function excludeStillReferenced(doc, confirmedItems) {
        var directUse = collectDirectUse(doc);
        var deletingIds = {};
        for (var i = 0; i < confirmedItems.length; i++) deletingIds[confirmedItems[i].id] = true;

        var droppedCount;
        do {
            droppedCount = 0;
            var crossRefs = collectCrossRefs(doc, deletingIds);
            for (var j = 0; j < confirmedItems.length; j++) {
                var confirmedItem = confirmedItems[j];
                if (!deletingIds[confirmedItem.id]) continue;
                if (directUse[confirmedItem.kind][confirmedItem.id] || crossRefs[confirmedItem.kind][confirmedItem.id]) {
                    delete deletingIds[confirmedItem.id];
                    droppedCount++;
                }
            }
        } while (droppedCount > 0);

        var deletingItems = [];
        for (var k = 0; k < confirmedItems.length; k++) {
            if (deletingIds[confirmedItems[k].id]) deletingItems.push(confirmedItems[k]);
        }
        return deletingItems;
    }

    /**
     * 項目を削除する
     * Swatch には削除可否のプロパティが無いので、既定のスウォッチは remove() の例外で見分ける
     * @param {Object} candidate - 削除候補
     * @returns {boolean} 削除できたら true
     */
    function removeCandidate(candidate) {
        if (candidate.kind !== "swatch") {
            candidate.item.remove();
            return true;
        }
        try {
            candidate.item.remove();
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * スタイルグループが空かを調べる
     * @param {Object} candidate - スタイルグループの削除候補
     * @returns {boolean} スタイルもサブグループも無ければ true
     */
    function isGroupEmpty(candidate) {
        var styleGroup = candidate.item;
        return styleGroup[STYLE_KINDS[candidate.styleKind].styles].length === 0 &&
            styleGroup[STYLE_KINDS[candidate.styleKind].groups].length === 0;
    }

    /**
     * 1つのドキュメントで、チェックを付けた候補を削除する
     * スタイルグループは最後に、実際に空になったものだけを消す（子グループが先の順）
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<Object>} confirmedCandidates - このドキュメントのチェックを付けた候補
     * @param {Object<string, number>} removedCounts - 種類ごとの削除数（加算していく）
     * @returns {void}
     */
    function removeConfirmedCandidates(doc, confirmedCandidates, removedCounts) {
        var confirmedItems = [];
        var groupCandidates = [];
        for (var i = 0; i < confirmedCandidates.length; i++) {
            if (confirmedCandidates[i].kind === GROUP_KIND_KEY) {
                groupCandidates.push(confirmedCandidates[i]);
            } else {
                confirmedItems.push(confirmedCandidates[i]);
            }
        }

        var deletingItems = excludeStillReferenced(doc, confirmedItems);
        for (var j = 0; j < deletingItems.length; j++) {
            if (removeCandidate(deletingItems[j])) removedCounts[deletingItems[j].kind]++;
        }
        for (var k = 0; k < groupCandidates.length; k++) {
            if (!isGroupEmpty(groupCandidates[k])) continue;
            groupCandidates[k].item.remove();
            removedCounts[GROUP_KIND_KEY]++;
        }
    }

    /**
     * 候補をドキュメントごとに分け、ドキュメントごとに1回の取り消し単位で削除する
     * @param {Array<Document>} targetDocs - 対象ドキュメント
     * @param {Array<Object>} confirmedCandidates - チェックを付けた候補
     * @returns {Object<string, number>} 種類ごとの削除数
     */
    function removeAllConfirmed(targetDocs, confirmedCandidates) {
        var removedCounts = { styleGroup: 0 };
        for (var i = 0; i < TARGET_KEYS.length; i++) removedCounts[TARGET_KEYS[i]] = 0;

        for (var j = 0; j < targetDocs.length; j++) {
            var doc = targetDocs[j];
            var docCandidates = [];
            for (var k = 0; k < confirmedCandidates.length; k++) {
                if (confirmedCandidates[k].doc === doc) docCandidates.push(confirmedCandidates[k]);
            }
            if (docCandidates.length === 0) continue;
            app.doScript(function () {
                removeConfirmedCandidates(doc, docCandidates, removedCounts);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undoName));
        }
        return removedCounts;
    }

    // =========================================
    // 結果 / Result
    // =========================================

    /**
     * 種類ごとの削除数・合計・残した数を表示する
     * @param {{targets: Object<string, boolean>, removeEmptyGroups: boolean}} dialogState - 設定ダイアログの内容
     * @param {Object<string, number>} removedCounts - 種類ごとの削除数
     * @param {number} confirmedCount - チェックを付けた候補の数
     * @param {number} documentCount - 処理したドキュメント数
     * @returns {void}
     */
    function showResult(dialogState, removedCounts, confirmedCount, documentCount) {
        var resultLines = [getLabel(LABELS.alert.result)];
        if (documentCount > 1) resultLines.push(formatLabel(getLabel(LABELS.alert.docCount), [documentCount]));
        resultLines.push("");

        var shownKeys = [];
        for (var i = 0; i < TARGET_KEYS.length; i++) {
            if (dialogState.targets[TARGET_KEYS[i]]) shownKeys.push(TARGET_KEYS[i]);
        }
        if (dialogState.removeEmptyGroups) shownKeys.push(GROUP_KIND_KEY);

        var totalCount = 0;
        for (var j = 0; j < shownKeys.length; j++) {
            var removedCount = removedCounts[shownKeys[j]];
            totalCount += removedCount;
            resultLines.push(formatLabel(getLabel(LABELS.alert.countLine), [getLabel(LABELS.kindName[shownKeys[j]]), removedCount]));
        }
        resultLines.push("");
        resultLines.push(formatLabel(getLabel(LABELS.alert.totalLine), [totalCount]));
        if (confirmedCount > totalCount) resultLines.push(formatLabel(getLabel(LABELS.alert.keptLine), [confirmedCount - totalCount]));

        alert(resultLines.join("\n"), getLabel(LABELS.dialog.title));
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * 設定ダイアログ → 候補の洗い出し → 確認リスト → 削除 → 結果表示
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument), getLabel(LABELS.dialog.title));
            return;
        }

        var dialogState = showSettingsDialog();
        if (!dialogState) return;

        var targetDocs = dialogState.allDocuments ? app.documents.everyItem().getElements() : [app.activeDocument];
        var candidates = findAllCandidates(targetDocs, dialogState);
        if (candidates.length === 0) {
            alert(getLabel(LABELS.alert.noneFound), getLabel(LABELS.dialog.title));
            return;
        }

        var confirmedCandidates = showConfirmDialog(candidates, targetDocs.length > 1);
        if (!confirmedCandidates || confirmedCandidates.length === 0) return;

        var removedCounts = removeAllConfirmed(targetDocs, confirmedCandidates);
        showResult(dialogState, removedCounts, confirmedCandidates.length, targetDocs.length);
    }

    main();

})();
