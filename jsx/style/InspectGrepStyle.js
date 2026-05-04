#target indesign

/*
概要

アクティブドキュメントの段落スタイルに設定されている正規表現スタイル（GREPスタイル）を収集し、一覧表示します。
段落スタイル／適用される文字スタイル／正規表現を列で確認でき、任意の列でソート可能です。
一覧全体と、重複を除いた正規表現リストをテキストとして書き出せます。
UIおよびエラーメッセージは、実行環境のロケール（日本語／英語）に応じて切り替わります。

Summary

Collects GREP styles from paragraph styles in the active document and displays them in a list.
You can review paragraph styles, applied character styles, and GREP expressions, and sort by any column.
Exports both the full list and a deduplicated list of GREP expressions as a text file.
UI and error messages switch based on the current locale (Japanese/English).
*/

(function () {

    // =========================================
    // バージョンとローカライズ / Version and localization
    // =========================================

    var SCRIPT_VERSION = "v1.0.0";

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: {
            ja: "正規表現スタイル一覧",
            en: "GREP Style Inspector"
        },
        sortPanelTitle: {
            ja: "ソート基準",
            en: "Sort By"
        },
        paragraphStyle: {
            ja: "段落スタイル",
            en: "Paragraph Style"
        },
        characterStyle: {
            ja: "文字スタイル",
            en: "Character Style"
        },
        grepExpression: {
            ja: "正規表現",
            en: "GREP Expression"
        },
        exportTextButton: {
            ja: "テキストに書き出し…",
            en: "Export to Text..."
        },
        closeButton: {
            ja: "閉じる",
            en: "Close"
        },
        exportSectionAll: {
            ja: "■ 一覧",
            en: "■ List"
        },
        exportSectionUniqueExpressions: {
            ja: "■ 正規表現一覧（重複なし・{count}件）",
            en: "■ GREP Expressions (unique: {count})"
        },
        exportFilePrefix: {
            ja: "正規表現スタイル一覧",
            en: "GREPStyleInspector"
        },
        unavailable: {
            ja: "取得不可",
            en: "Unavailable"
        },
        errNoDocument: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        errNoGrepStyles: {
            ja: "正規表現スタイルは見つかりませんでした。",
            en: "No GREP styles were found."
        },
        exportComplete: {
            ja: "正規表現スタイルの一覧を書き出しました。",
            en: "The GREP style list has been exported."
        },
        errExportFailed: {
            ja: "テキストファイルを書き出せませんでした。",
            en: "The text file could not be exported."
        },
        errOpenExportedFileFailed: {
            ja: "書き出したファイルを開けませんでした。",
            en: "The exported file could not be opened."
        }
    };

    function L(key) {
        if (!LABELS[key]) {
            return key;
        }
        return LABELS[key][lang] || LABELS[key].en || key;
    }


    // =========================================
    // メイン処理 / Main process
    // =========================================

    if (app.documents.length === 0) {
        alert(L("errNoDocument"));
        return;
    }

    var activeDocument = app.activeDocument;
    var grepStyleRows = collectGrepStyleRows(activeDocument);

    if (grepStyleRows.length === 0) {
        alert(L("errNoGrepStyles"));
        return;
    }

    showResultDialog(grepStyleRows);

    // =========================================
    // 正規表現スタイルの収集 / GREP style collection
    // =========================================

    /* ドキュメント内の正規表現スタイル情報を収集 / Collect GREP style rows from the document */
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

    /* 正規表現スタイルに適用された文字スタイル名を取得 / Get the character style name applied by a GREP style */
    function getNestedGrepCharacterStyleName(nestedGrepStyle) {
        try {
            return getStylePath(nestedGrepStyle.appliedCharacterStyle);
        } catch (characterStyleError) {
            return L("unavailable");
        }
    }

    /* 正規表現文字列を取得 / Get the GREP expression string */
    function getNestedGrepExpression(nestedGrepStyle) {
        try {
            return nestedGrepStyle.grepExpression;
        } catch (grepExpressionError) {
            return L("unavailable");
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    function showResultDialog(grepStyleRows) {

        var countText = (lang === "ja")
            ? grepStyleRows.length + "件"
            : grepStyleRows.length + " items";

        var dialog = new Window(
            "dialog",
            L("dialogTitle") + " " + SCRIPT_VERSION + "（" + countText + "）"
        );

        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "fill"];
        dialog.margins = 16;
        dialog.spacing = 10;
        dialog.resizeable = true;

        var sortPanel = dialog.add("panel", undefined, L("sortPanelTitle"));
        sortPanel.orientation = "row";
        sortPanel.alignChildren = ["left", "center"];
        sortPanel.alignment = "fill";
        sortPanel.margins = [15, 20, 15, 10];
        sortPanel.spacing = 12;

        var sortRadioButtons = [
            sortPanel.add("radiobutton", undefined, L("paragraphStyle")),
            sortPanel.add("radiobutton", undefined, L("characterStyle")),
            sortPanel.add("radiobutton", undefined, L("grepExpression"))
        ];
        sortRadioButtons[0].value = true;

        var resultListBox = dialog.add("listbox", undefined, "", {
            numberOfColumns: 3,
            showHeaders: true,
            columnTitles: [L("paragraphStyle"), L("characterStyle"), L("grepExpression")],
            columnWidths: [160, 160, 300]
        });
        resultListBox.preferredSize = [660, 480];
        resultListBox.minimumSize = [360, 200];
        resultListBox.alignment = ["fill", "fill"];

        function populateResultList(sortColumnIndex) {
            var sortedRows = sortRowsByColumn(grepStyleRows, sortColumnIndex);
            resultListBox.removeAll();
            for (var rowIndex = 0; rowIndex < sortedRows.length; rowIndex++) {
                var listItem = resultListBox.add("item", sortedRows[rowIndex][0]);
                listItem.subItems[0].text = sortedRows[rowIndex][1];
                listItem.subItems[1].text = sortedRows[rowIndex][2];
            }
        }

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

        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = ["right", "bottom"];
        buttonGroup.spacing = 8;

        var exportTextButton = buttonGroup.add("button", undefined, L("exportTextButton"));
        buttonGroup.add("button", undefined, L("closeButton"), { name: "ok" });

        exportTextButton.onClick = function () {
            exportGrepStyleRows(activeDocument, grepStyleRows);
        };

        dialog.onResizing = dialog.onResize = function () {
            this.layout.resize();
        };

        dialog.show();
    }

    /* 指定列を基準に行データをソート / Sort row data by the specified column */
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

    /* 正規表現スタイル一覧をテキストファイルに書き出す / Export GREP style rows to a text file */
    function exportGrepStyleRows(activeDocument, grepStyleRows) {
        var exportFile = null;
        try {
            var exportLines = buildExportLines(grepStyleRows);
            exportFile = createExportFile(activeDocument);
            exportFile.encoding = "UTF-8";
            if (!exportFile.open("w")) {
                throw new Error(exportFile.error || L("errExportFailed"));
            }
            exportFile.write(exportLines.join("\r"));
            exportFile.close();
        } catch (exportError) {
            try { if (exportFile) exportFile.close(); } catch (closeError) { }
            alert(L("errExportFailed") + "\n\n" + (exportError && exportError.message ? exportError.message : exportError));
            return;
        }

        alert(L("exportComplete") + "\n\n" + exportFile.fsName);

        try {
            exportFile.execute();
        } catch (openError) {
            alert(L("errOpenExportedFileFailed") + "\n\n" + openError.message);
        }
    }

    /* 書き出し用の行データを作成 / Build text lines for export */
    function buildExportLines(grepStyleRows) {
        var exportLines = [L("exportSectionAll"), L("paragraphStyle") + "\t" + L("characterStyle") + "\t" + L("grepExpression")];
        for (var exportRowIndex = 0; exportRowIndex < grepStyleRows.length; exportRowIndex++) {
            exportLines.push(grepStyleRows[exportRowIndex].join("\t"));
        }

        var uniqueExpressions = getUniqueExpressions(grepStyleRows);
        exportLines.push("");
        exportLines.push(L("exportSectionUniqueExpressions").replace("{count}", uniqueExpressions.length));
        for (var uniqueExpressionIndex = 0; uniqueExpressionIndex < uniqueExpressions.length; uniqueExpressionIndex++) {
            exportLines.push(uniqueExpressions[uniqueExpressionIndex]);
        }

        return exportLines;
    }

    /* 重複しない正規表現リストを取得 / Get deduplicated GREP expressions */
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

    /* タイムスタンプを作成 / Create timestamp string */
    function createTimestamp() {
        var exportDate = new Date();
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

    /* 書き出し先ファイルを作成 / Create export file object */
    function createExportFile(activeDocument) {
        var documentName = sanitizeFileName(activeDocument.name.replace(/\.indd$/i, ""));
        var fileName = L("exportFilePrefix") + "-" + documentName + "-" + createTimestamp() + ".txt";
        return File(Folder.desktop + "/" + encodeURI(fileName));
    }

    /* ファイル名に使えない文字を置換 / Replace characters that cannot be used in file names */
    function sanitizeFileName(fileName) {
        return fileName.replace(/[\\\/:\*\?"<>\|]/g, "_");
    }

    // =========================================
    // スタイル名処理 / Style name utilities
    // =========================================

    /* スタイルグループを含めたパス名を取得 / Get style path including style groups */
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