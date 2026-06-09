#target indesign

    /*
    
    ### 概要
    
    - 選択中の表に対して、行の高さを設定します。
    - ［範囲］で、現在の選択範囲・選択中の表が属するストーリー・ドキュメント全体のいずれを対象にするかを切り替えられます。ストーリー／ドキュメントを選ぶと、複数の表へまとめて適用します。
    - 表全体、見出し行を除く全体、または選択した行のみを対象にできます。「選択した行のみ」は範囲が選択範囲のときだけ有効です。
    - 選択した行は行オブジェクトではなく行インデックスで保持し、適用時に再解決します。
    - 複数の表が混在する選択はエラーにします。
    - すべてのセルを選択している場合は、選択行ではなく表全体として扱います。
    - 行の高さは手動入力で指定でき、初期値には現在の行高を使います。対象行がすべて同じ高さならその値を表示し、そうでない場合は見出し行を除く平均値を表示します。入力値はドキュメントの縦方向単位で表示・入力し、内部では pt に変換して適用します。
    - プレビューは常に有効で、結果を確認しながら調整できます。表全体を対象にしているときのみ、ダイアログ表示中の選択ハイライトを解除し、閉じた後に元の選択を復元します。
    - プレビュー適用と最終適用は共通ロジックに寄せ、ダイアログは設定状態を返し、適用はメイン処理側で行います。
    - 左下には画面モード切り替えボタンと、親テキストフレームを内容に合わせてフィットする［親フレームの調整］ボタンを配置しています。
    - 画面モード切り替えボタンは動作ベースの文言（Enter Preview / Exit Preview）を使います。
    
    ### Overview
    
    - Sets row heights for the selected table.
    - The Scope setting switches whether to target the current selection, the story that contains the selected table, or the whole document. Choosing Story or Document applies the change to multiple tables at once.
    - You can target the whole table, the whole table except header rows, or only the selected rows. "Selected Rows" is available only when the scope is Selection.
    - Selected rows are stored as row indices rather than row objects and are resolved again when applied.
    - Selections that mix multiple tables are treated as an error.
    - If all cells are selected, the script treats the target as the whole table rather than selected rows.
    - Row height can be set manually, and the initial value is based on the current row height. If all target rows share the same height, that value is shown; otherwise, the average of the non-header rows is shown. The input value is shown and entered in the document's vertical units, then converted internally to points.
    - Preview is always enabled so you can adjust the setting while seeing the result. Only when the whole table is targeted, the script clears the selection highlight while the dialog is open and restores the original selection after closing.
    - Preview apply and final apply share the same apply logic, while the dialog returns state and the main flow performs the actual apply.
    - At the lower left, the dialog includes a screen-mode toggle button and a Fit Parent Frame button that fits the parent text frame to its content.
    - The screen-mode toggle button uses action-oriented labels (Enter Preview / Exit Preview).
    
    ### 更新履歴 / Changelog
    
    - v1.0.0 (20260420) : Initial version
    - v1.2.1 (20260420) : Clear selection highlight while the dialog is open and restore the original selection after closing
    - v1.3.1 (20260609) : Wire up the Scope panel (Document / Story / Selection) so it actually selects which tables are targeted
    
    */

    (function () {

        // =========================================
        // バージョン / Version
        // =========================================

        var SCRIPT_VERSION = "v1.3.1";

        // =========================================
        // ユーザー設定 / User Settings
        // =========================================

        /* 行の高さの下限（pt）。InDesign の最小行高（約 0.0139 inch ≒ 1pt）に合わせる / Minimum row height in points (InDesign's own minimum is about 0.0139 inch ≈ 1pt) */
        var MIN_ROW_HEIGHT_PT = 1.0008;

        // =========================================
        // ローカライズ / Localization
        // =========================================

        function getCurrentLang() {
            return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
        }
        var currentLanguage = getCurrentLang();

        /* 日英ラベル定義 / Japanese-English label definitions */
        var LABELS = {
            dialog: {
                title: { ja: "行の高さを設定", en: "Set Row Height" }
            },
            scope: {
                panel: { ja: "範囲", en: "Scope" },
                document: { ja: "ドキュメント", en: "Document" },
                story: { ja: "ストーリー", en: "Story" },
                selection: { ja: "選択範囲", en: "Selection" }
            },
            target: {
                panel: { ja: "対象", en: "Target" },
                whole: { ja: "表全体", en: "Whole Table" },
                wholeNoHead: { ja: "表全体（見出し行を除く）", en: "Whole Table (Except Header Rows)" },
                selected: { ja: "選択した行のみ", en: "Selected Rows" }
            },
            mode: {
                panel: { ja: "行の高さ", en: "Row Height" },
                minimum: { ja: "最小限度", en: "Minimum" },
                specified: { ja: "指定値を使用", en: "Use Specified Value" }
            },
            options: {
                panel: { ja: "オプション", en: "Options" }
            },
            button: {
                enterPreview: { ja: "プレビュー", en: "Enter Preview" },
                exitPreview: { ja: "標準モード", en: "Exit Preview" },
                fitParentFrame: { ja: "親フレームの調整", en: "Fit Parent Frame" },
                ok: { ja: "OK", en: "OK" },
                cancel: { ja: "キャンセル", en: "Cancel" }
            },
            tooltip: {
                scopeDocument: {
                    ja: "ドキュメント内のすべての表を対象にします。",
                    en: "Targets all tables in the document."
                },
                scopeStory: {
                    ja: "選択中のテキスト（ストーリー）に含まれる表を対象にします。",
                    en: "Targets tables in the selected story."
                },
                scopeSelection: {
                    ja: "現在の選択範囲に含まれる表を対象にします。",
                    en: "Targets tables in the current selection."
                },
                rowHeightInput: {
                    ja: "↑↓キーで増減（Shiftで整数値にスナップ）。単位はドキュメントの縦方向単位です。",
                    en: "Use Up/Down to adjust (Shift snaps to whole numbers). Unit follows the document's vertical units."
                },
                screenModeToggle: {
                    ja: "プレビュー画面モードを切り替えて、ガイドや選択ハイライトの表示を消します。",
                    en: "Toggles the preview screen mode to hide guides and selection highlights."
                },
                fitParentFrame: {
                    ja: "表を含む親テキストフレームの高さを内容に合わせて調整します。",
                    en: "Fits the height of the parent text frame to its content."
                },
                modeMinimum: {
                    ja: "各行を最小の高さ（約1pt）に設定し、内容に合わせて自動的に伸びるようにします。",
                    en: "Sets each row to the minimum height (about 1 pt) so it grows to fit its content."
                },
                modeSpecified: {
                    ja: "入力した値を行の高さとして設定します。",
                    en: "Uses the entered value as the row height."
                }
            },
            error: {
                noTable: {
                    ja: "表、セル、または表を含むテキストフレームを選択してください。",
                    en: "Please select a table, cell, or a text frame containing a table."
                },
                multipleTables: {
                    ja: "複数の表が選択されています。1つの表だけを選択してください。",
                    en: "Multiple tables are selected. Please select only one table."
                },
                invalidNumber: {
                    ja: "正の数値を入力してください。",
                    en: "Please enter a positive number."
                }
            },
            progress: {
                title: { ja: "行の高さを適用中", en: "Applying Row Heights" }
            }
        };

        /* ドット区切りパスでラベルを取得し、{slash} を / に置換 / Resolve a dotted label path and replace {slash} with / */
        function L(path) {
            var parts = path.split(".");
            var node = LABELS;
            for (var i = 0; i < parts.length; i++) {
                if (node && node.hasOwnProperty(parts[i])) {
                    node = node[parts[i]];
                } else {
                    return path;
                }
            }
            var text = (node && node[currentLanguage]) ? node[currentLanguage] : path;
            return text.replace(/\{slash\}/g, "/");
        }

        // =========================================
        // 単位 / Units
        // =========================================

        /* 環境設定の単位ラベルを取得 / Get unit label from preferences */
        function getUnitLabel() {
            try {
                var unit = app.activeDocument.viewPreferences.verticalMeasurementUnits;
                switch (unit) {
                    case MeasurementUnits.MILLIMETERS: return "mm";
                    case MeasurementUnits.CENTIMETERS: return "cm";
                    case MeasurementUnits.INCHES:
                    case MeasurementUnits.INCHES_DECIMAL: return "inch";
                    case MeasurementUnits.POINTS:
                    case MeasurementUnits.AMERICAN_POINTS: return "pt";
                    case MeasurementUnits.PICAS: return "pica";
                    case MeasurementUnits.PIXELS: return "px";
                    case MeasurementUnits.AGATES: return "ag";
                    case MeasurementUnits.Q: return "Q";
                    case MeasurementUnits.HA: return "H";
                    case MeasurementUnits.CICEROS: return "cicero";
                    case MeasurementUnits.BAI: return "倍";
                    case MeasurementUnits.U: return "U";
                    default: return "";
                }
            } catch (e) {
                return "";
            }
        }

        /* 単位ラベルから MeasurementUnits を推定 / Resolve MeasurementUnits from document preferences */
        function getDocumentVerticalUnit() {
            try {
                return app.activeDocument.viewPreferences.verticalMeasurementUnits;
            } catch (e) {
                return MeasurementUnits.POINTS;
            }
        }

        /* 1単位あたりの pt 数を取得 / Get points per one unit */
        function getPointsPerUnit(unit) {
            switch (unit) {
                case MeasurementUnits.MILLIMETERS: return 72 / 25.4;
                case MeasurementUnits.CENTIMETERS: return 72 / 2.54;
                case MeasurementUnits.INCHES:
                case MeasurementUnits.INCHES_DECIMAL: return 72;
                case MeasurementUnits.POINTS:
                case MeasurementUnits.AMERICAN_POINTS: return 1;
                case MeasurementUnits.PICAS: return 12;
                case MeasurementUnits.AGATES: return 5.5;
                case MeasurementUnits.Q:
                case MeasurementUnits.HA: return 72 / 25.4 * 0.25;
                default: return 1;
            }
        }

        /* 表示単位の値を pt に変換 / Convert a value in display units to points */
        function convertToPointsByUnit(value, unit) {
            return value * getPointsPerUnit(unit);
        }

        /* pt を表示単位の値に変換 / Convert points to a value in display units */
        function convertFromPointsByUnit(value, unit) {
            return value / getPointsPerUnit(unit);
        }

        /* 表示用の数値を整形 / Format display value */
        function formatDisplayValue(value) {
            if (Math.abs(value - Math.round(value)) < 0.0001) return String(Math.round(value));
            return String(Math.round(value * 1000) / 1000);
        }

        // =========================================
        // ユーティリティ / Utilities
        // =========================================

        /* 選択ノードから表・セル・行へさかのぼる / Walk up from a selected node to its table, cell, or row */
        function walkUpToTable(start) {
            var node = start;
            var foundCell = null;
            var foundRow = null;
            var foundTable = null;
            while (node) {
                try {
                    if (node instanceof Cell) { foundCell = node; foundTable = node.parent; break; }
                    if (node instanceof Row) { foundRow = node; foundTable = node.parent; break; }
                    if (node instanceof Table) { foundTable = node; break; }
                    if (node instanceof TextFrame) {
                        if (node.tables.length > 0) { foundTable = node.tables[0]; break; }
                    }
                    node = node.parent;
                } catch (e) { break; }
            }
            return { table: foundTable, cell: foundCell, row: foundRow, node: start };
        }

        /* 行インデックスを重複なく集める収集オブジェクト / A collector that gathers unique row indices */
        function createRowIndexCollector() {
            var rowMap = {};
            var hasSpecificRows = false;

            function addRowIndex(index) {
                if (index === undefined || index === null) return;
                rowMap[index] = true;
                hasSpecificRows = true;
            }

            function addRowsFromRows(rows) {
                if (!rows || rows.length === 0) return false;
                var added = false;
                for (var r = 0; r < rows.length; r++) {
                    try {
                        addRowIndex(rows[r].index);
                        added = true;
                    } catch (e) { }
                }
                return added;
            }

            function addRowsFromCells(cells) {
                if (!cells || cells.length === 0) return false;
                var added = false;
                for (var c = 0; c < cells.length; c++) {
                    try {
                        addRowIndex(cells[c].parentRow.index);
                        added = true;
                    } catch (e) { }
                }
                return added;
            }

            function toSortedIndices() {
                var rowIndices = [];
                for (var k in rowMap) {
                    if (rowMap.hasOwnProperty(k)) rowIndices.push(parseInt(k, 10));
                }
                rowIndices.sort(function (a, b) { return a - b; });
                return rowIndices;
            }

            return {
                addRowIndex: addRowIndex,
                addRowsFromRows: addRowsFromRows,
                addRowsFromCells: addRowsFromCells,
                toSortedIndices: toSortedIndices,
                hasSpecificRows: function () { return hasSpecificRows; }
            };
        }

        /* 選択 1 件分の行インデックスを収集オブジェクトへ集める / Collect the row indices for a single selection item into the collector */
        function collectRowIndicesFromItem(walkResult, collector) {
            var node = walkResult.node;

            /* 選択ノードが持つ行・セルのコレクションから行インデックスを集める / Collect row indices from any row/cell collections the node exposes */
            var rangeSources = [
                { prop: "parentRows", add: collector.addRowsFromRows },
                { prop: "rows", add: collector.addRowsFromRows },
                { prop: "parentCells", add: collector.addRowsFromCells },
                { prop: "cells", add: collector.addRowsFromCells }
            ];
            var addedFromRange = false;
            for (var s = 0; s < rangeSources.length; s++) {
                try {
                    var collection = node && node[rangeSources[s].prop];
                    if (collection && collection.length > 0) {
                        addedFromRange = rangeSources[s].add(collection.everyItem().getElements()) || addedFromRange;
                    }
                } catch (e) { }
            }
            if (addedFromRange) return;

            if (walkResult.row) { collector.addRowIndex(walkResult.row.index); return; }
            if (walkResult.cell) { collector.addRowIndex(walkResult.cell.parentRow.index); return; }

            try {
                if (node && node.parentRow) collector.addRowIndex(node.parentRow.index);
            } catch (e) { }
        }

        function resolveTargetFromSelection(selection) {
            if (!selection || selection.length === 0) return null;

            var table = null;
            var collector = createRowIndexCollector();

            for (var i = 0; i < selection.length; i++) {
                var walkResult = walkUpToTable(selection[i]);
                if (!walkResult.table) continue;

                if (!table) {
                    table = walkResult.table;
                } else if (table !== walkResult.table) {
                    return { error: "multipleTables" };
                }

                collectRowIndicesFromItem(walkResult, collector);
            }

            if (!table) return null;

            var rowIndices = null;
            if (collector.hasSpecificRows()) {
                rowIndices = collector.toSortedIndices();

                /* すべての行が選択されている場合は表全体として扱う / Treat as whole table if all rows are selected */
                if (rowIndices.length === table.rows.length) {
                    rowIndices = null;
                }
            }

            return { table: table, rowIndices: rowIndices };
        }

        /* ドキュメント内のすべての表を取得 / Collect all tables in the document */
        function collectDocumentTables() {
            var tables = [];
            try {
                var stories = app.activeDocument.stories;
                for (var s = 0; s < stories.length; s++) {
                    var storyTables = stories[s].tables;
                    for (var t = 0; t < storyTables.length; t++) tables.push(storyTables[t]);
                }
            } catch (e) { }
            return tables;
        }

        /* 指定した表が属するストーリー内のすべての表を取得 / Collect all tables in the story that contains the given table */
        function collectStoryTables(baseTable) {
            try {
                var story = baseTable.parentStory;
                var storyTables = story.tables;
                var tables = [];
                for (var t = 0; t < storyTables.length; t++) tables.push(storyTables[t]);
                return tables;
            } catch (e) {
                return [baseTable];
            }
        }

        /* 範囲設定に応じて対象となる表の配列を返す / Resolve the array of target tables for the current scope */
        function resolveScopeTables(scope, baseTable) {
            if (scope === "document") return collectDocumentTables();
            if (scope === "story") return collectStoryTables(baseTable);
            return [baseTable];
        }

        /* 触れた表の元の行高を遅延記憶し、まとめて復元する / Lazily snapshot original row heights of touched tables and restore them together */
        function createRowSnapshotStore() {
            var entries = [];
            var seen = [];

            function ensure(table) {
                for (var i = 0; i < seen.length; i++) {
                    if (seen[i] === table) return;
                }
                var rows = getAllRows(table);
                entries.push({ rows: rows, heights: getOriginalHeights(rows) });
                seen.push(table);
            }

            function restoreAll() {
                for (var i = 0; i < entries.length; i++) {
                    restoreRowHeights(entries[i].rows, entries[i].heights);
                }
            }

            return { ensure: ensure, restoreAll: restoreAll };
        }



        /* 表全体の行を取得 / Get all rows of the table */
        function getAllRows(table) {
            var all = [];
            for (var i = 0; i < table.rows.length; i++) all.push(table.rows[i]);
            return all;
        }

        /* 保存した行インデックスから行配列を再解決 / Re-resolve rows from saved row indices */
        function getRowsByIndices(table, rowIndices) {
            if (!rowIndices || rowIndices.length === 0) return [];
            var rows = [];
            for (var i = 0; i < rowIndices.length; i++) {
                var idx = rowIndices[i];
                if (idx >= 0 && idx < table.rows.length) {
                    rows.push(table.rows[idx]);
                }
            }
            return rows;
        }

        /* 見出し行を除く全行を取得 / Get all rows excluding header rows */
        function getBodyRows(table) {
            var headerCount = 0;
            try { headerCount = table.headerRowCount || 0; } catch (e) { }
            var rows = [];
            for (var i = headerCount; i < table.rows.length; i++) {
                rows.push(table.rows[i]);
            }
            return rows;
        }

        /* すべての対象行が同じ高さか確認し、一致していればその値を返す / Return the common row height if all target rows match */
        function getCommonRowHeight(rows) {
            if (!rows || rows.length === 0) return null;
            var first = rows[0].height;
            for (var i = 1; i < rows.length; i++) {
                if (Math.abs(rows[i].height - first) > 0.01) return null;
            }
            return first;
        }

        /* 行高の平均値を返す / Return the average row height */
        function getAverageRowHeight(rows) {
            if (!rows || rows.length === 0) return null;
            var total = 0;
            for (var i = 0; i < rows.length; i++) {
                total += rows[i].height;
            }
            return total / rows.length;
        }

        /* ダイアログ初期値の行高を求める / Resolve the initial row height for the dialog */
        function getInitialRowHeightForDialog(table, selectedRowIndices) {
            var targetRows = resolveTargetRows(table, selectedRowIndices, selectedRowIndices && selectedRowIndices.length > 0 ? "selected" : "whole");
            var commonHeight = getCommonRowHeight(targetRows);
            if (commonHeight !== null) return commonHeight;

            var bodyRows = getBodyRows(table);
            var averageHeight = getAverageRowHeight(bodyRows);
            if (averageHeight !== null) return averageHeight;

            return 15;
        }

        // =========================================
        // プログレス / Progress
        // =========================================

        /* 適用中の進捗を表示する palette を作る / Build a palette that shows apply progress */
        function createProgressBar(title, maxValue) {
            var win = new Window("palette", title);
            win.orientation = "column";
            win.alignChildren = "fill";
            win.margins = 16;

            var label = win.add("statictext", undefined, "");
            label.preferredSize.width = 320;

            var bar = win.add("progressbar", undefined, 0, maxValue);
            bar.preferredSize = [320, 12];

            win.show();

            return {
                update: function (value, text) {
                    bar.value = value;
                    if (text !== undefined) label.text = text;
                    win.update();
                },
                close: function () {
                    try { win.close(); } catch (e) { }
                }
            };
        }

        // =========================================
        // スクリーンモード切り替え / Screen mode toggle
        // =========================================

        function isInPreviewScreenMode() {
            try {
                return app.activeWindow && app.activeWindow.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE;
            } catch (e) {
                return false;
            }
        }

        function toggleScreenPreviewMode() {
            try {
                var activeWindow = app.activeWindow;
                if (!activeWindow) return;
                if (activeWindow.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE) {
                    activeWindow.screenMode = ScreenModeOptions.PREVIEW_OFF;
                } else {
                    activeWindow.screenMode = ScreenModeOptions.PREVIEW_TO_PAGE;
                }
            } catch (e) { /* ignore */ }
        }

        function getScreenModeToggleButtonLabel() {
            return isInPreviewScreenMode() ? L('button.exitPreview') : L('button.enterPreview');
        }

        function updateScreenModeToggleButtonLabel(button) {
            button.text = getScreenModeToggleButtonLabel();
        }

        // =========================================
        // キーボード操作 / Keyboard interaction
        // =========================================

        /**
         * ↑↓ で値を増減。通常は±0.1、shift では整数値へスナップ
         * Arrow keys adjust value: normal ±0.1, shift snaps to whole numbers
         */
        function changeValueByArrowKey(editText, onUpdate) {
            editText.addEventListener("keydown", function (event) {
                var keyName = event.keyName;
                var isArrowUp = (keyName === "Up" || keyName === "PageUp");
                var isArrowDown = (keyName === "Down" || keyName === "PageDown");
                if (!isArrowUp && !isArrowDown) return;

                /* 先に既定動作を抑止し、修飾キー状態は keyboardState を優先して取得 / Prevent default early and prefer keyboardState for modifier keys */
                event.preventDefault();

                var value = Number(editText.text);
                if (isNaN(value)) value = 0;

                var keyboard = ScriptUI.environment.keyboardState;
                var isShift = keyboard && keyboard.shiftKey ? true : event.shiftKey;
                var isUp = isArrowUp;
                if (isShift) {
                    if (isUp) {
                        value = Math.ceil(value + 0.0001);
                    } else {
                        value = Math.floor(value - 0.0001);
                    }
                } else {
                    value += isUp ? 0.1 : -0.1;
                    value = Math.round(value * 10) / 10;
                }

                if (value < 0) value = 0;

                editText.text = value;
                if (typeof onUpdate === "function") onUpdate();
            });
        }

        // =========================================
        // 行の高さ操作 / Row height operations
        // =========================================

        function applyRowHeight(rows, height) {
            for (var i = 0; i < rows.length; i++) {
                rows[i].height = height;
            }
        }

        function getOriginalHeights(rows) {
            var heights = [];
            for (var i = 0; i < rows.length; i++) {
                heights.push(rows[i].height);
            }
            return heights;
        }

        function restoreRowHeights(rows, heights) {
            for (var i = 0; i < rows.length; i++) {
                rows[i].height = heights[i];
            }
        }

        /* 現在の対象行を解決 / Resolve current target rows */
        function resolveTargetRows(table, selectedRowIndices, targetMode) {
            var hasSelection = !!(selectedRowIndices && selectedRowIndices.length > 0);
            if (targetMode === "selected" && hasSelection) return getRowsByIndices(table, selectedRowIndices);
            if (targetMode === "body") return getBodyRows(table);
            return getAllRows(table);
        }

        /* 現在の選択を保存 / Save current selection */
        function getSelectionSnapshot() {
            var snapshot = [];
            try {
                var sel = app.selection;
                for (var i = 0; i < sel.length; i++) snapshot.push(sel[i]);
            } catch (e) { }
            return snapshot;
        }

        /* 保存した選択を復元 / Restore saved selection */
        function restoreSelectionSnapshot(snapshot) {
            try {
                if (snapshot && snapshot.length > 0) {
                    app.select(snapshot);
                } else {
                    app.select(NothingEnum.NOTHING);
                }
            } catch (e) { }
        }

        function fitParentFrameToContent(table) {
            try {
                var parentFrame = null;
                var story = null;

                try {
                    story = table.parentStory;
                } catch (e) { }

                try {
                    if (table.storyOffset && table.storyOffset.parentTextFrames && table.storyOffset.parentTextFrames.length > 0) {
                        parentFrame = table.storyOffset.parentTextFrames[0];
                    }
                } catch (e) { }

                if (!parentFrame) {
                    try {
                        if (table.parent && table.parent.parentTextFrames && table.parent.parentTextFrames.length > 0) {
                            parentFrame = table.parent.parentTextFrames[0];
                        }
                    } catch (e) { }
                }

                if (!parentFrame && story) {
                    var containers = story.textContainers;
                    if (containers && containers.length > 0) {
                        parentFrame = containers[0];
                    }
                }

                if (!parentFrame) return false;

                try {
                    if (story) story.recompose();
                } catch (e) { }

                parentFrame.fit(FitOptions.FRAME_TO_CONTENT);

                try {
                    if (story) story.recompose();
                } catch (e) { }

                return true;
            } catch (e) {
                return false;
            }
        }

        // =========================================
        // ダイアログ / Dialog
        // =========================================

        function showRowHeightDialog(table, selectedRowIndices, defaultValuePt) {
            /* 触れた表の元の高さを遅延記憶（範囲切り替えに追従） / Lazily remember original heights of touched tables (follows scope changes) */
            var snapshot = createRowSnapshotStore();
            var hasSelection = !!(selectedRowIndices && selectedRowIndices.length > 0);
            var documentUnit = getDocumentVerticalUnit();
            var defaultValueDisplay = convertFromPointsByUnit(defaultValuePt, documentUnit);

            var dlg = new Window("dialog", L('dialog.title') + ' ' + SCRIPT_VERSION);
            dlg.orientation = "column";
            dlg.alignChildren = "fill";

            var contentGroup = dlg.add("group");
            contentGroup.orientation = "row";
            contentGroup.alignChildren = ["fill", "top"];
            contentGroup.alignment = "fill";

            var leftColumn = contentGroup.add("group");
            leftColumn.orientation = "column";
            leftColumn.alignChildren = "fill";
            leftColumn.alignment = ["fill", "fill"];

            var rightColumn = contentGroup.add("group");
            rightColumn.orientation = "column";
            rightColumn.alignChildren = ["fill", "top"];
            rightColumn.alignment = ["right", "fill"];

            /* 範囲選択パネル / Scope selection panel */
            var scopeGroup = leftColumn.add("panel", undefined, L('scope.panel'));
            scopeGroup.orientation = "column";
            scopeGroup.alignChildren = "left";
            scopeGroup.margins = [15, 20, 15, 10];
            var scopeDocument = scopeGroup.add("radiobutton", undefined, L('scope.document'));
            var scopeStory = scopeGroup.add("radiobutton", undefined, L('scope.story'));
            var scopeSelection = scopeGroup.add("radiobutton", undefined, L('scope.selection'));
            scopeDocument.helpTip = L('tooltip.scopeDocument');
            scopeStory.helpTip = L('tooltip.scopeStory');
            scopeSelection.helpTip = L('tooltip.scopeSelection');
            scopeSelection.value = true;

            scopeDocument.onClick = function () { syncTargetEnabled(); updatePreview(); };
            scopeStory.onClick = function () { syncTargetEnabled(); updatePreview(); };
            scopeSelection.onClick = function () { syncTargetEnabled(); updatePreview(); };

            /* 対象選択パネル / Target selection panel */
            var targetGroup = leftColumn.add("panel", undefined, L('target.panel'));
            targetGroup.orientation = "column";
            targetGroup.alignChildren = "left";
            targetGroup.margins = [15, 20, 15, 10];
            var targetWhole = targetGroup.add("radiobutton", undefined, L('target.whole'));
            var targetBody = targetGroup.add("radiobutton", undefined, L('target.wholeNoHead'));
            var targetSel = targetGroup.add("radiobutton", undefined, L('target.selected'));
            targetSel.enabled = hasSelection;
            if (hasSelection) {
                targetSel.value = true;
            } else {
                targetWhole.value = true;
            }

            /* モード選択パネル / Mode selection panel */
            var modeGroup = leftColumn.add("panel", undefined, L('mode.panel'));
            modeGroup.orientation = "column";
            modeGroup.alignChildren = "left";
            modeGroup.margins = [15, 20, 15, 10];

            var modeMinimum = modeGroup.add("radiobutton", undefined, L('mode.minimum'));
            var modeSpecified = modeGroup.add("radiobutton", undefined, L('mode.specified'));
            modeMinimum.helpTip = L('tooltip.modeMinimum');
            modeSpecified.helpTip = L('tooltip.modeSpecified');
            modeMinimum.value = true;

            var manualRow = modeGroup.add("group");
            manualRow.orientation = "row";
            manualRow.alignChildren = ["left", "center"];
            var input = manualRow.add("edittext", undefined, formatDisplayValue(defaultValueDisplay));
            input.characters = 5;
            input.helpTip = L('tooltip.rowHeightInput');
            manualRow.add("statictext", undefined, getUnitLabel());

            /* 指定値モードのときだけ入力欄を有効化 / Enable the input only in the specified-value mode */
            function syncInputEnabled() {
                input.enabled = modeSpecified.value;
            }
            syncInputEnabled();

            modeMinimum.onClick = function () { syncInputEnabled(); updatePreview(); };
            modeSpecified.onClick = function () {
                syncInputEnabled();
                if (modeSpecified.value) input.active = true;
                updatePreview();
            };

            targetWhole.onClick = function () { updatePreview(); };
            targetBody.onClick = function () { updatePreview(); };
            targetSel.onClick = function () { updatePreview(); };


            /* 右カラムの実行ボタン / Action buttons in the right column */
            var actionGroup = rightColumn.add("group");
            actionGroup.orientation = "column";
            actionGroup.alignChildren = ["fill", "top"];
            actionGroup.alignment = ["fill", "top"];
            var okBtn = actionGroup.add("button", undefined, L('button.ok'), { name: "ok" });
            var cancelBtn = actionGroup.add("button", undefined, L('button.cancel'), { name: "cancel" });

            /* OK/Cancel とプレビューボタンの間を伸ばすスペーサー / Spacer that pushes the preview button to the bottom */
            var rightSpacer = rightColumn.add("group");
            rightSpacer.alignment = ["fill", "fill"];

            /* 右カラム下部の画面モード切り替え / Screen mode toggle at the bottom of the right column */
            var screenModeGroup = rightColumn.add("group");
            screenModeGroup.orientation = "row";
            screenModeGroup.alignChildren = ["fill", "center"];
            screenModeGroup.alignment = ["fill", "bottom"];
            var screenModeToggleBtn = screenModeGroup.add("button", undefined, getScreenModeToggleButtonLabel());
            screenModeToggleBtn.alignment = ["fill", "center"];
            screenModeToggleBtn.helpTip = L('tooltip.screenModeToggle');

            screenModeToggleBtn.onClick = function () {
                toggleScreenPreviewMode();
                updateScreenModeToggleButtonLabel(screenModeToggleBtn);
            };

            /* 左カラム下部のオプションパネル / Options panel at the bottom of the left column */
            var optionsGroup = leftColumn.add("panel", undefined, L('options.panel'));
            optionsGroup.orientation = "row";
            optionsGroup.alignChildren = ["left", "center"];
            optionsGroup.margins = [15, 20, 15, 10];
            var fitParentFrameBtn = optionsGroup.add("button", undefined, L('button.fitParentFrame'));
            fitParentFrameBtn.helpTip = L('tooltip.fitParentFrame');

            fitParentFrameBtn.onClick = function () {
                snapshot.restoreAll();
                updatePreview();
                fitParentFrameToContent(table);
            };


            function parseValue() {
                var v = parseFloat(input.text);
                if (isNaN(v) || v <= 0) return null;
                return convertToPointsByUnit(v, documentUnit);
            }

            /* 現在のモードに応じた行高（pt）を取得 / Resolve the row height (pt) for the current mode */
            function getRowHeightValue() {
                if (modeMinimum.value) return MIN_ROW_HEIGHT_PT;
                return parseValue();
            }

            function getCurrentScope() {
                if (scopeDocument.value) return "document";
                if (scopeStory.value) return "story";
                return "selection";
            }

            /* 選択行モードは「選択範囲」のときだけ有効 / The selected-rows mode is valid only in the Selection scope */
            function syncTargetEnabled() {
                var allowSelected = (getCurrentScope() === "selection") && hasSelection;
                targetSel.enabled = allowSelected;
                if (!allowSelected && targetSel.value) {
                    targetWhole.value = true;
                }
            }

            function getCurrentTargetMode() {
                if (targetSel.value && targetSel.enabled && hasSelection) return "selected";
                if (targetBody.value) return "body";
                return "whole";
            }

            function updatePreview() {
                /* 毎回すべて元に戻してから対象表に適用 / Always restore then apply to the current target tables */
                snapshot.restoreAll();
                var scope = getCurrentScope();
                var targetMode = getCurrentTargetMode();
                var tables = resolveScopeTables(scope, table);
                var value = getRowHeightValue();
                if (value === null) return;
                for (var i = 0; i < tables.length; i++) {
                    snapshot.ensure(tables[i]);
                }
                for (var j = 0; j < tables.length; j++) {
                    var rowIndicesForTable = (tables[j] === table) ? selectedRowIndices : null;
                    var targetRows = resolveTargetRows(tables[j], rowIndicesForTable, targetMode);
                    applyRowHeight(targetRows, value);
                }
            }

            input.onChanging = updatePreview;
            changeValueByArrowKey(input, updatePreview);

            cancelBtn.onClick = function () {
                snapshot.restoreAll();
                dlg.close(0);
            };

            /* 初期状態を範囲に同期してプレビュー / Sync to scope and show the initial preview */
            syncTargetEnabled();
            updatePreview();

            if (dlg.show() !== 1) return null;

            var finalScope = getCurrentScope();
            var finalTargetMode = getCurrentTargetMode();
            var value = getRowHeightValue();
            snapshot.restoreAll();

            if (value === null) {
                alert(L('error.invalidNumber'));
                return null;
            }
            return { height: value, targetMode: finalTargetMode, scope: finalScope };
        }

        // =========================================
        // メイン処理 / Main
        // =========================================

        var originalSelection = getSelectionSnapshot();
        var target = resolveTargetFromSelection(app.selection);

        if (!target) {
            alert(L('error.noTable'));
        } else if (target.error === "multipleTables") {
            alert(L('error.multipleTables'));
        } else {
            var table = target.table;
            var initialHeightPt = getInitialRowHeightForDialog(table, target.rowIndices);

            var didClearSelection = false;

            /* 表全体のときのみハイライトをオフ / Clear selection highlight only when whole table is targeted */
            if (target.rowIndices === null) {
                try {
                    app.select(NothingEnum.NOTHING);
                    didClearSelection = true;
                } catch (e) { }
            }

            var result = showRowHeightDialog(table, target.rowIndices, initialHeightPt);

            if (didClearSelection) {
                restoreSelectionSnapshot(originalSelection);
            }

            if (result !== null) {
                var targetTables = resolveScopeTables(result.scope, table);

                /* 先に対象行をすべて解決して総数を数える / Resolve all target rows up front and count the total */
                var jobRowSets = [];
                var totalRows = 0;
                for (var ti = 0; ti < targetTables.length; ti++) {
                    var rowIndicesForTable = (targetTables[ti] === table) ? target.rowIndices : null;
                    var rows = resolveTargetRows(targetTables[ti], rowIndicesForTable, result.targetMode);
                    jobRowSets.push(rows);
                    totalRows += rows.length;
                }

                var progress = createProgressBar(L('progress.title') + ' ' + SCRIPT_VERSION, totalRows);
                try {
                    var doneRows = 0;
                    for (var ji = 0; ji < jobRowSets.length; ji++) {
                        var jobRows = jobRowSets[ji];
                        for (var ri = 0; ri < jobRows.length; ri++) {
                            jobRows[ri].height = result.height;
                            doneRows++;
                            /* 行ごとの更新は重いので間引く / Throttle updates since per-row refresh is costly */
                            if (doneRows === totalRows || doneRows % 10 === 0) {
                                progress.update(doneRows, doneRows + " / " + totalRows);
                            }
                        }
                    }
                } finally {
                    progress.close();
                }
            }
        }

    })();
