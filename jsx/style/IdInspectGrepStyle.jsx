#target indesign

/*

### 概要

ドキュメント内の段落スタイルに設定された正規表現スタイル（GREPスタイル）を一覧表示し、テキストファイルへ書き出します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdInspectGrepStyle.md

### Overview

Lists the GREP styles defined in the paragraph styles of a document and exports them to a text file.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdInspectGrepStyle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdInspectGrepStyle";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdInspectGrepStyle.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdInspectGrepStyle.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// ==============================
// UIレイアウトの共通設定 / Shared UI layout
// ==============================

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

(function () {

    // =========================================
    // レイアウト設定 / Layout settings
    // =========================================

    /* 一覧の列幅（px）/ Column widths of the result list (px) */
    var RESULT_COLUMN_WIDTHS = [160, 160, 300];

    /* 一覧の推奨サイズと最小サイズ [幅, 高さ]（px）/ Preferred and minimum size of the result list (px) */
    var RESULT_LIST_SIZE     = [660, 480];
    var RESULT_LIST_MIN_SIZE = [360, 200];

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    var LABELS = {
        dialog: {
            title: { ja: "正規表現スタイル一覧", en: "GREP Style Inspector" }
        },
        panel: {
            sort: { ja: "ソート基準", en: "Sort By" }
        },
        column: {
            paragraphStyle: { ja: "段落スタイル", en: "Paragraph Style" },
            characterStyle: { ja: "文字スタイル", en: "Character Style" },
            grepExpression: { ja: "正規表現", en: "GREP Expression" }
        },
        button: {
            exportText: { ja: "テキストに書き出し…", en: "Export to Text..." },
            close:      { ja: "閉じる", en: "Close" }
        },
        exportText: {
            sectionAll:               { ja: "■ 一覧", en: "■ List" },
            sectionUniqueExpressions: { ja: "■ 正規表現一覧（重複なし・{count}件）", en: "■ GREP Expressions (unique: {count})" },
            filePrefix:               { ja: "正規表現スタイル一覧", en: "GREPStyleInspector" },
            complete:                 { ja: "書き出しました。", en: "Export complete." }
        },
        value: {
            unavailable: { ja: "取得不可", en: "Unavailable" }
        },
        error: {
            noDocument:            { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noGrepStyles:          { ja: "正規表現スタイルが見つかりませんでした。", en: "No GREP styles were found." },
            exportFailed:          { ja: "書き出しに失敗しました。", en: "Export failed." },
            openExportedFileFailed:{ ja: "書き出したファイルを開けませんでした。", en: "The exported file could not be opened." }
        }
    };

    // =========================================
    // メイン処理 / Main process
    // =========================================
    // 一覧表示とテキスト書き出しだけでドキュメントを変更しないため、doScript でのラップは不要
    // / This script only lists and exports; it never edits the document, so no doScript wrapper is needed

    if (app.documents.length === 0) {
        alert(getLabel("error.noDocument"));
        return;
    }

    var activeDocument = app.activeDocument;
    var grepStyleRows = collectGrepStyleRows(activeDocument);

    if (grepStyleRows.length === 0) {
        alert(getLabel("error.noGrepStyles"));
        return;
    }

    showResultDialog(grepStyleRows);

    // =========================================
    // 正規表現スタイルの収集 / GREP style collection
    // =========================================

    /**
     * 段落スタイルから正規表現スタイルを集めて一覧用の行にする
     * @param {Document} activeDocument 対象ドキュメント
     * @returns {Array<Array<string>>} 段落スタイル・文字スタイル・正規表現の行
     */
    function collectGrepStyleRows(activeDocument) {
        var grepStyleRows = [];

        /* ドキュメント内の全段落スタイルを取得 / Get all paragraph styles in the document */
        var paragraphStyles = activeDocument.allParagraphStyles;

        for (var paragraphStyleIndex = 0; paragraphStyleIndex < paragraphStyles.length; paragraphStyleIndex++) {

            var paragraphStyle = paragraphStyles[paragraphStyleIndex];

            try {
                var nestedGrepStyles = paragraphStyle.nestedGrepStyles;

                if (!nestedGrepStyles || nestedGrepStyles.length === 0) {
                    continue;
                }

                for (var grepStyleIndex = 0; grepStyleIndex < nestedGrepStyles.length; grepStyleIndex++) {

                    var nestedGrepStyle = nestedGrepStyles[grepStyleIndex];

                    var paragraphStyleName = getStylePath(paragraphStyle);
                    var characterStyleName = getNestedGrepCharacterStyleName(nestedGrepStyle);
                    var grepExpression = getNestedGrepExpression(nestedGrepStyle);

                    grepStyleRows.push([paragraphStyleName, characterStyleName, grepExpression]);
                }

            } catch (nestedGrepStyleError) {
                /* 正規表現スタイルを取得できないスタイルは無視 / Ignore styles whose GREP styles cannot be read */
            }
        }

        return grepStyleRows;
    }

    /**
     * 正規表現スタイルに適用されている文字スタイル名を取得する
     * @param {NestedGrepStyle} nestedGrepStyle 対象の正規表現スタイル
     * @returns {string} 文字スタイル名
     */
    function getNestedGrepCharacterStyleName(nestedGrepStyle) {
        try {
            return getStylePath(nestedGrepStyle.appliedCharacterStyle);
        } catch (characterStyleError) {
            return getLabel("value.unavailable");
        }
    }

    /**
     * 正規表現スタイルの検索式を取得する
     * @param {NestedGrepStyle} nestedGrepStyle 対象の正規表現スタイル
     * @returns {string} 正規表現
     */
    function getNestedGrepExpression(nestedGrepStyle) {
        try {
            return nestedGrepStyle.grepExpression;
        } catch (grepExpressionError) {
            return getLabel("value.unavailable");
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 収集した正規表現スタイルの一覧ダイアログを表示する
     * @param {Array<Array<string>>} grepStyleRows 一覧の行
     * @returns {void}
     */
    function showResultDialog(grepStyleRows) {

        var countText = (uiLang === "ja")
            ? grepStyleRows.length + "件"
            : grepStyleRows.length + " items";

        var dialog = new Window(
            "dialog",
            getLabel("dialog.title") + " " + SCRIPT_VERSION + "（" + countText + "）"
        );

        setupWindow(dialog, 10);
        dialog.alignChildren = ["fill", "fill"];
        dialog.resizeable = true;

        var sortPanel = dialog.add("panel", undefined, getLabel("panel.sort"));
        setupPanel(sortPanel, COLUMN_SPACING);
        sortPanel.orientation = "row";
        sortPanel.alignChildren = ["left", "center"];

        var sortRadioButtons = [
            sortPanel.add("radiobutton", undefined, getLabel("column.paragraphStyle")),
            sortPanel.add("radiobutton", undefined, getLabel("column.characterStyle")),
            sortPanel.add("radiobutton", undefined, getLabel("column.grepExpression"))
        ];
        sortRadioButtons[0].value = true;

        var resultListBox = dialog.add("listbox", undefined, "", {
            numberOfColumns: 3,
            showHeaders: true,
            columnTitles: [getLabel("column.paragraphStyle"), getLabel("column.characterStyle"), getLabel("column.grepExpression")],
            columnWidths: RESULT_COLUMN_WIDTHS
        });
        resultListBox.preferredSize = RESULT_LIST_SIZE;
        resultListBox.minimumSize = RESULT_LIST_MIN_SIZE;
        resultListBox.alignment = ["fill", "fill"];

        /**
         * 指定した列で並べ替えて一覧を描き直す
         * @param {number} sortColumnIndex 並べ替えに使う列の位置
         * @returns {void}
         */
        function populateResultList(sortColumnIndex) {
            var sortedRows = sortRowsByColumn(grepStyleRows, sortColumnIndex);
            resultListBox.removeAll();
            for (var rowIndex = 0; rowIndex < sortedRows.length; rowIndex++) {
                var listItem = resultListBox.add("item", sortedRows[rowIndex][0]);
                listItem.subItems[0].text = sortedRows[rowIndex][1];
                listItem.subItems[1].text = sortedRows[rowIndex][2];
            }
        }

        /**
         * 選択中のソート基準の列位置を取得する
         * @returns {number} 列の位置
         */
        function getSelectedSortColumnIndex() {
            for (var sortRadioIndex = 0; sortRadioIndex < sortRadioButtons.length; sortRadioIndex++) {
                if (sortRadioButtons[sortRadioIndex].value) return sortRadioIndex;
            }
            return 0;
        }

        populateResultList(getSelectedSortColumnIndex());

        for (var sortButtonIndex = 0; sortButtonIndex < sortRadioButtons.length; sortButtonIndex++) {
            sortRadioButtons[sortButtonIndex].onClick = function () {
                populateResultList(getSelectedSortColumnIndex());
            };
        }

        /* ボタン行（左に書き出し、右に閉じる）/ Button row (Export on the left, Close on the right) */
        var buttonRow = addButtonRow(dialog);
        var btnExportText = buttonRow.leftGroup.add("button", undefined, getLabel("button.exportText"));
        var btnClose = buttonRow.rightGroup.add("button", undefined, getLabel("button.close"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        btnExportText.onClick = function () {
            exportGrepStyleRows(activeDocument, grepStyleRows);
        };

        dialog.onResizing = dialog.onResize = function () {
            this.layout.resize();
        };

        dialog.show();
    }

    /**
     * 指定した列で行を並べ替える（同じ値どうしは元の順を保つ）
     * @param {Array<Array<string>>} grepStyleRows 一覧の行
     * @param {number} sortColumnIndex 並べ替えに使う列の位置
     * @returns {Array<Array<string>>} 並べ替えた行
     */
    function sortRowsByColumn(grepStyleRows, sortColumnIndex) {
        /* ExtendScript の比較関数つき sort() は遅いうえ並びが狂うことがあるので、
           「列の値＋区切り＋元の位置（ゼロ埋め）」の文字列キーを引数なしの sort() で並べる /
           A comparator sort is slow and unreliable in ExtendScript, so sort string keys
           ("value + separator + zero-padded original index") with a plain sort() */
        var sortKeys = [];
        for (var rowIndex = 0; rowIndex < grepStyleRows.length; rowIndex++) {
            var paddedIndex = String(1000000 + rowIndex).substring(1);
            sortKeys.push(String(grepStyleRows[rowIndex][sortColumnIndex]) + "\u0001" + paddedIndex);
        }
        sortKeys.sort();

        var sortedRows = [];
        for (var keyIndex = 0; keyIndex < sortKeys.length; keyIndex++) {
            var originalIndex = parseInt(sortKeys[keyIndex].substring(sortKeys[keyIndex].length - 6), 10);
            sortedRows.push(grepStyleRows[originalIndex]);
        }
        return sortedRows;
    }

    // =========================================
    // 書き出し処理 / Export utilities
    // =========================================

    /**
     * 一覧と重複を除いた正規表現をテキストへ書き出す
     * @param {Document} activeDocument 対象ドキュメント
     * @param {Array<Array<string>>} grepStyleRows 一覧の行
     * @returns {void}
     */
    function exportGrepStyleRows(activeDocument, grepStyleRows) {
        var exportFile = null;
        try {
            var exportLines = buildExportLines(grepStyleRows);
            exportFile = createExportFile(activeDocument);
            exportFile.encoding = "UTF-8";
            if (!exportFile.open("w")) {
                throw new Error(exportFile.error || getLabel("error.exportFailed"));
            }
            exportFile.write(exportLines.join("\r"));
            exportFile.close();
        } catch (exportError) {
            try { if (exportFile) exportFile.close(); } catch (closeError) { }
            alert(getLabel("error.exportFailed") + "\n\n" + (exportError && exportError.message ? exportError.message : exportError));
            return;
        }

        alert(getLabel("exportText.complete") + "\n\n" + exportFile.fsName);

        try {
            exportFile.execute();
        } catch (openError) {
            alert(getLabel("error.openExportedFileFailed") + "\n\n" + openError.message);
        }
    }

    /**
     * 書き出すテキストの行を組み立てる
     * @param {Array<Array<string>>} grepStyleRows 一覧の行
     * @returns {Array<string>} 書き出す行
     */
    function buildExportLines(grepStyleRows) {
        var exportLines = [getLabel("exportText.sectionAll"), getLabel("column.paragraphStyle") + "\t" + getLabel("column.characterStyle") + "\t" + getLabel("column.grepExpression")];
        for (var exportRowIndex = 0; exportRowIndex < grepStyleRows.length; exportRowIndex++) {
            exportLines.push(grepStyleRows[exportRowIndex].join("\t"));
        }

        var uniqueExpressions = getUniqueExpressions(grepStyleRows);
        exportLines.push("");
        exportLines.push(getLabel("exportText.sectionUniqueExpressions", { count: uniqueExpressions.length }));
        for (var uniqueExpressionIndex = 0; uniqueExpressionIndex < uniqueExpressions.length; uniqueExpressionIndex++) {
            exportLines.push(uniqueExpressions[uniqueExpressionIndex]);
        }

        return exportLines;
    }

    /**
     * 重複を除いた正規表現の一覧を作る
     * @param {Array<Array<string>>} grepStyleRows 一覧の行
     * @returns {Array<string>} 正規表現の配列
     */
    function getUniqueExpressions(grepStyleRows) {
        var seenExpressions = {};
        var uniqueExpressions = [];
        for (var expressionIndex = 0; expressionIndex < grepStyleRows.length; expressionIndex++) {
            var currentExpression = grepStyleRows[expressionIndex][2];
            if (!seenExpressions[currentExpression]) {
                seenExpressions[currentExpression] = true;
                uniqueExpressions.push(currentExpression);
            }
        }
        return uniqueExpressions;
    }

    /**
     * ファイル名に使うタイムスタンプを作る
     * @returns {string} タイムスタンプ文字列
     */
    function createTimestamp() {
        var exportDate = new Date();
        /**
         * 数値を 2 桁のゼロ埋め文字列にする
         * @param {number} value 対象の数値
         * @returns {string} 2 桁の文字列
         */
        function pad2(value) {
            return (value < 10 ? "0" : "") + value;
        }
        return exportDate.getFullYear() +
            pad2(exportDate.getMonth() + 1) +
            pad2(exportDate.getDate()) + "-" +
            pad2(exportDate.getHours()) +
            pad2(exportDate.getMinutes()) +
            pad2(exportDate.getSeconds());
    }

    /**
     * 書き出し先のファイルを作る
     * @param {Document} activeDocument 対象ドキュメント
     * @returns {File} 書き出し先のファイル
     */
    function createExportFile(activeDocument) {
        var documentName = sanitizeFileName(activeDocument.name.replace(/\.indd$/i, ""));
        var fileName = getLabel("exportText.filePrefix") + "-" + documentName + "-" + createTimestamp() + ".txt";
        return File(Folder.desktop + "/" + encodeURI(fileName));
    }

    /**
     * ファイル名に使えない文字を置き換える
     * @param {string} fileName 元のファイル名
     * @returns {string} 安全なファイル名
     */
    function sanitizeFileName(fileName) {
        return fileName.replace(/[\\\/:\*\?"<>\|]/g, "_");
    }

    // =========================================
    // スタイル名処理 / Style name utilities
    // =========================================

    /**
     * スタイルグループを含めたスタイルのパスを取得する
     * @param {object} styleObject 対象のスタイル
     * @returns {string} スタイルのパス
     */
    function getStylePath(styleObject) {

        if (!styleObject || !styleObject.isValid) {
            return "";
        }

        var stylePathNames = [];
        var currentObject = styleObject;

        while (currentObject && currentObject.isValid) {

            if (currentObject.constructor.name === "Document") {
                break;
            }

            if (currentObject.name !== undefined) {
                stylePathNames.unshift(currentObject.name);
            }

            currentObject = currentObject.parent;
        }

        return stylePathNames.join("/");
    }

})();