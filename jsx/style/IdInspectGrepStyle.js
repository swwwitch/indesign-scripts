#target indesign

/*
 * IdInspectGrepStyle.jsx
 *
 * ドキュメント内の段落スタイルに設定された正規表現スタイル（GREP スタイル）を一覧表示し、テキストへ書き出します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdInspectGrepStyle";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-05-04";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdInspectGrepStyle.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdInspectGrepStyle.md

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
        export: {
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
        return node[currentLang] || node.en || labelKey;
    }

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

        var countText = (currentLang === "ja")
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

        /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
        var buttonGroup = dialog.add("group");
        setupRow(buttonGroup, "right", 8);
        buttonGroup.alignment = ["right", "bottom"];

        var exportTextButton = buttonGroup.add("button", undefined, getLabel("button.exportText"));
        buttonGroup.add("button", undefined, getLabel("button.close"), { name: "ok" });

        exportTextButton.onClick = function () {
            exportGrepStyleRows(activeDocument, grepStyleRows);
        };

        dialog.onResizing = dialog.onResize = function () {
            this.layout.resize();
        };

        dialog.show();
    }

    /**
     * 指定した列で行を並べ替える
     * @param {Array<Array<string>>} grepStyleRows 一覧の行
     * @param {number} sortColumnIndex 並べ替えに使う列の位置
     * @returns {Array<Array<string>>} 並べ替えた行
     */
    function sortRowsByColumn(grepStyleRows, sortColumnIndex) {
        var sortedRows = grepStyleRows.slice();
        sortedRows.sort(function (firstRow, secondRow) {
            if (firstRow[sortColumnIndex] < secondRow[sortColumnIndex]) return -1;
            if (firstRow[sortColumnIndex] > secondRow[sortColumnIndex]) return 1;
            return 0;
        });
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

        alert(getLabel("export.complete") + "\n\n" + exportFile.fsName);

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
        var exportLines = [getLabel("export.sectionAll"), getLabel("column.paragraphStyle") + "\t" + getLabel("column.characterStyle") + "\t" + getLabel("column.grepExpression")];
        for (var exportRowIndex = 0; exportRowIndex < grepStyleRows.length; exportRowIndex++) {
            exportLines.push(grepStyleRows[exportRowIndex].join("\t"));
        }

        var uniqueExpressions = getUniqueExpressions(grepStyleRows);
        exportLines.push("");
        exportLines.push(getLabel("export.sectionUniqueExpressions").replace("{count}", uniqueExpressions.length));
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
        var fileName = getLabel("export.filePrefix") + "-" + documentName + "-" + createTimestamp() + ".txt";
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