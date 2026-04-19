#target indesign

/*
概要 / Overview
選択中の表に対して、行の高さを設定します。
表全体、見出し行を除く全体、または選択した行のみを対象にできます。
選択した行は行オブジェクトではなく行インデックスで保持し、適用時に再解決します。
複数の表が混在する選択はエラーにします。
すべてのセルを選択している場合は、選択行ではなく表全体として扱います。
行の高さは手動入力で指定でき、初期値には現在の行高を使います。対象行がすべて同じ高さならその値を表示し、そうでない場合は見出し行を除く平均値を表示します。入力値はドキュメントの縦方向単位で表示・入力し、内部では pt に変換して適用します。
プレビューは常に有効で、結果を確認しながら調整できます。表全体を対象にしているときのみ、ダイアログ表示中の選択ハイライトを解除し、閉じた後に元の選択を復元します。
プレビュー適用と最終適用は共通ロジックに寄せ、ダイアログは設定状態を返し、適用はメイン処理側で行います。
左下には画面モード切り替えボタンと、親テキストフレームを内容に合わせてフィットする［親フレームの調整］ボタンを配置しています。
画面モード切り替えボタンは動作ベースの文言（Enter Preview / Exit Preview）を使います。

Sets row heights for the selected table.
You can target the whole table, the whole table except header rows, or only the selected rows.
Selected rows are stored as row indices rather than row objects and are resolved again when applied.
Selections that mix multiple tables are treated as an error.
If all cells are selected, the script treats the target as the whole table rather than selected rows.
Row height can be set manually, and the initial value is based on the current row height. If all target rows share the same height, that value is shown; otherwise, the average of the non-header rows is shown. The input value is shown and entered in the document's vertical units, then converted internally to points.
Preview is always enabled so you can adjust the setting while seeing the result. Only when the whole table is targeted, the script clears the selection highlight while the dialog is open and restores the original selection after closing.
Preview apply and final apply share the same apply logic, while the dialog returns state and the main flow performs the actual apply.
At the lower left, the dialog includes a screen-mode toggle button and a Fit Parent Frame button that fits the parent text frame to its content.
The screen-mode toggle button uses action-oriented labels (Enter Preview / Exit Preview).

更新履歴 / Changelog
v1.0.0 (20260420) : Initial version
v1.1.0 (20260420) : Reject mixed multi-table selections, store selected rows by index, and normalize row-height input by document vertical units
v1.1.7 (20260420) : Initialize the row-height field from current values, using the common height when all target rows match or the average of non-header rows otherwise
v1.2.0 (20260420) : Removed the Auto Fit UI and logic, keeping manual row-height input only
v1.2.1 (20260420) : Clear selection highlight while the dialog is open and restore the original selection after closing
*/

    (function () {

        // =========================================
        // バージョンとローカライズ / Version and Localization
        // =========================================

        var SCRIPT_VERSION = "v1.2.1";

        function getCurrentLang() {
            return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
        }
        var lang = getCurrentLang();

        /* 日英ラベル定義 / Japanese-English label definitions */
        var LABELS = {
            dialogTitle: { ja: "行の高さを設定", en: "Set Row Height" },
            targetPanel: { ja: "対象", en: "Target" },
            targetWhole: { ja: "表全体", en: "Whole Table" },
            targetWholeNoHead: { ja: "表全体（見出し行を除く）", en: "Whole Table (Except Header Rows)" },
            targetSelected: { ja: "選択した行のみ", en: "Selected Rows" },
            modePanel: { ja: "行の高さ", en: "Row Height" },
            enterPreview: { ja: "プレビュー", en: "Enter Preview" },
            exitPreview: { ja: "標準モード", en: "Exit Preview" },
            fitParentFrame: { ja: "親フレームの調整", en: "Fit Parent Frame" },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            errorNoTable: { ja: "表、セル、または表を含むテキストフレームを選択してください。", en: "Please select a table, cell, or a text frame containing a table." },
            errorMultipleTables: { ja: "複数の表が選択されています。1つの表だけを選択してください。", en: "Multiple tables are selected. Please select only one table." },
            errorInvalidNumber: { ja: "正の数値を入力してください。", en: "Please enter a positive number." }
        };

        function L(key) {
            return (LABELS[key] && LABELS[key][lang]) ? LABELS[key][lang] : key;
        }

        // =========================================
        // ユーティリティ / Utilities
        // =========================================

        function resolveTargetFromSelection(selection) {
            if (!selection || selection.length === 0) return null;

            var table = null;
            var rowMap = {};
            var hasSpecificRows = false;

            function addRowIndex(index) {
                if (index === undefined || index === null) return;
                rowMap[index] = true;
                hasSpecificRows = true;
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

            function walkUp(start) {
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

            for (var i = 0; i < selection.length; i++) {
                var r = walkUp(selection[i]);
                if (!r.table) continue;

                if (!table) {
                    table = r.table;
                } else if (table !== r.table) {
                    return { error: "multipleTables" };
                }

                var node = r.node;
                var addedFromRange = false;

                try {
                    if (node && node.parentRows && node.parentRows.length > 0) {
                        addedFromRange = addRowsFromRows(node.parentRows.everyItem().getElements()) || addedFromRange;
                    }
                } catch (e) { }

                try {
                    if (node && node.rows && node.rows.length > 0) {
                        addedFromRange = addRowsFromRows(node.rows.everyItem().getElements()) || addedFromRange;
                    }
                } catch (e) { }

                try {
                    if (node && node.parentCells && node.parentCells.length > 0) {
                        addedFromRange = addRowsFromCells(node.parentCells.everyItem().getElements()) || addedFromRange;
                    }
                } catch (e) { }

                try {
                    if (node && node.cells && node.cells.length > 0) {
                        addedFromRange = addRowsFromCells(node.cells.everyItem().getElements()) || addedFromRange;
                    }
                } catch (e) { }

                if (addedFromRange) continue;

                if (r.row) {
                    addRowIndex(r.row.index);
                    continue;
                }

                if (r.cell) {
                    addRowIndex(r.cell.parentRow.index);
                    continue;
                }

                try {
                    if (node && node.parentRow) {
                        addRowIndex(node.parentRow.index);
                        continue;
                    }
                } catch (e) { }
            }

            if (!table) return null;

            var rowIndices = null;
            if (hasSpecificRows) {
                rowIndices = [];
                for (var k in rowMap) {
                    if (rowMap.hasOwnProperty(k)) rowIndices.push(parseInt(k, 10));
                }
                rowIndices.sort(function (a, b) { return a - b; });

                /* すべての行が選択されている場合は表全体として扱う / Treat as whole table if all rows are selected */
                try {
                    if (rowIndices.length === table.rows.length) {
                        rowIndices = null;
                    }
                } catch (e) { }
            }

            return { table: table, rowIndices: rowIndices };
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

        /* 表示単位の値を pt に変換 / Convert a value in display units to points */
        function convertToPointsByUnit(value, unit) {
            switch (unit) {
                case MeasurementUnits.MILLIMETERS: return value * 72 / 25.4;
                case MeasurementUnits.CENTIMETERS: return value * 72 / 2.54;
                case MeasurementUnits.INCHES:
                case MeasurementUnits.INCHES_DECIMAL: return value * 72;
                case MeasurementUnits.POINTS:
                case MeasurementUnits.AMERICAN_POINTS: return value;
                case MeasurementUnits.PICAS: return value * 12;
                case MeasurementUnits.AGATES: return value * 5.5;
                case MeasurementUnits.Q: return value * 72 / 25.4 * 0.25;
                case MeasurementUnits.HA: return value * 72 / 25.4 * 0.25;
                default: return value;
            }
        }

        /* pt を表示単位の値に変換 / Convert points to a value in display units */
        function convertFromPointsByUnit(value, unit) {
            switch (unit) {
                case MeasurementUnits.MILLIMETERS: return value * 25.4 / 72;
                case MeasurementUnits.CENTIMETERS: return value * 2.54 / 72;
                case MeasurementUnits.INCHES:
                case MeasurementUnits.INCHES_DECIMAL: return value / 72;
                case MeasurementUnits.POINTS:
                case MeasurementUnits.AMERICAN_POINTS: return value;
                case MeasurementUnits.PICAS: return value / 12;
                case MeasurementUnits.AGATES: return value / 5.5;
                case MeasurementUnits.Q: return value / (72 / 25.4 * 0.25);
                case MeasurementUnits.HA: return value / (72 / 25.4 * 0.25);
                default: return value;
            }
        }

        /* 表示用の数値を整形 / Format display value */
        function formatDisplayValue(value) {
            if (Math.abs(value - Math.round(value)) < 0.0001) return String(Math.round(value));
            return String(Math.round(value * 1000) / 1000);
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
            return isInPreviewScreenMode() ? L('exitPreview') : L('enterPreview');
        }

        function updateScreenModeToggleButtonLabel(button) {
            button.text = getScreenModeToggleButtonLabel();
        }

        // =========================================
        // キーボード操作 / Keyboard interaction
        // =========================================

        /**
         * ↑↓ で値を増減。通常は±0.1、shift では整数値へスナップ、option/alt で±0.1
         * Arrow keys adjust value: normal ±0.1, shift snaps to whole numbers, option/alt ±0.1
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
                var isAlt = keyboard && keyboard.altKey ? true : event.altKey;
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

        /* 対象行に行高を適用 / Apply row height to target rows */
        function applyRowHeightByMode(rows, height) {
            applyRowHeight(rows, height);
            return true;
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
            /* 全行の元の高さを保存 / Save original heights of all rows */
            var allRows = getAllRows(table);
            var originalAllHeights = getOriginalHeights(allRows);
            var hasSelection = !!(selectedRowIndices && selectedRowIndices.length > 0);
            var documentUnit = getDocumentVerticalUnit();
            var defaultValueDisplay = convertFromPointsByUnit(defaultValuePt, documentUnit);

            var dlg = new Window("dialog", L('dialogTitle') + ' ' + SCRIPT_VERSION);
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
            rightColumn.alignment = ["right", "top"];

            /* 対象選択パネル / Target selection panel */
            var targetGroup = leftColumn.add("panel", undefined, L('targetPanel'));
            targetGroup.orientation = "column";
            targetGroup.alignChildren = "left";
            targetGroup.margins = [15, 20, 15, 10];
            var targetWhole = targetGroup.add("radiobutton", undefined, L('targetWhole'));
            var targetBody = targetGroup.add("radiobutton", undefined, L('targetWholeNoHead'));
            var targetSel = targetGroup.add("radiobutton", undefined, L('targetSelected'));
            targetSel.enabled = hasSelection;
            if (hasSelection) {
                targetSel.value = true;
            } else {
                targetWhole.value = true;
            }

            /* モード選択パネル / Mode selection panel */
            var modeGroup = leftColumn.add("panel", undefined, L('modePanel'));
            modeGroup.orientation = "column";
            modeGroup.alignChildren = "left";
            modeGroup.margins = [15, 20, 15, 10];

            var manualRow = modeGroup.add("group");
            manualRow.orientation = "row";
            manualRow.alignChildren = ["left", "center"];
            var input = manualRow.add("edittext", undefined, formatDisplayValue(defaultValueDisplay));
            input.characters = 5;
            input.active = true;
            manualRow.add("statictext", undefined, getUnitLabel());

            targetWhole.onClick = function () { updatePreview(); };
            targetBody.onClick = function () { updatePreview(); };
            targetSel.onClick = function () { updatePreview(); };


            /* 右カラムの実行ボタン / Action buttons in the right column */
            var actionGroup = rightColumn.add("group");
            actionGroup.orientation = "column";
            actionGroup.alignChildren = ["fill", "top"];
            actionGroup.alignment = ["fill", "top"];
            var okBtn = actionGroup.add("button", undefined, L('ok'), { name: "ok" });
            var cancelBtn = actionGroup.add("button", undefined, L('cancel'), { name: "cancel" });

            /* 左カラム下部の画面モード切り替え / Screen mode toggle at the bottom of the left column */
            var screenModeGroup = leftColumn.add("group");
            screenModeGroup.orientation = "row";
            screenModeGroup.alignChildren = ["left", "center"];
            screenModeGroup.alignment = ["left", "center"];
            var screenModeToggleBtn = screenModeGroup.add("button", undefined, getScreenModeToggleButtonLabel());
            screenModeToggleBtn.alignment = ["left", "center"];

            screenModeToggleBtn.onClick = function () {
                toggleScreenPreviewMode();
                updateScreenModeToggleButtonLabel(screenModeToggleBtn);
            };

            var fitParentFrameBtn = screenModeGroup.add("button", undefined, L('fitParentFrame'));

            fitParentFrameBtn.onClick = function () {
                restoreRowHeights(allRows, originalAllHeights);
                updatePreview();
                fitParentFrameToContent(table);
            };


            function parseValue() {
                var v = parseFloat(input.text);
                if (isNaN(v) || v <= 0) return null;
                return convertToPointsByUnit(v, documentUnit);
            }

            function getCurrentTargetMode() {
                if (targetSel.value && hasSelection) return "selected";
                if (targetBody.value) return "body";
                return "whole";
            }

            function updatePreview() {
                /* 毎回すべて元に戻してから対象行に適用 / Always restore then apply to current target */
                restoreRowHeights(allRows, originalAllHeights);
                var targetMode = getCurrentTargetMode();
                var targetRows = resolveTargetRows(table, selectedRowIndices, targetMode);
                var value = parseValue();
                if (value === null) return;
                applyRowHeightByMode(targetRows, value);
            }

            input.onChanging = updatePreview;
            changeValueByArrowKey(input, updatePreview);

            cancelBtn.onClick = function () {
                restoreRowHeights(allRows, originalAllHeights);
                dlg.close(0);
            };

            /* 初期プレビュー / Initial preview */
            updatePreview();

            if (dlg.show() !== 1) return null;

            var finalTargetMode = getCurrentTargetMode();
            var value = parseValue();
            restoreRowHeights(allRows, originalAllHeights);

            if (value === null) {
                alert(L('errorInvalidNumber'));
                return null;
            }
            return { height: value, targetMode: finalTargetMode };
        }

        // =========================================
        // メイン処理 / Main
        // =========================================

        var originalSelection = getSelectionSnapshot();
        var target = resolveTargetFromSelection(app.selection);

        if (!target) {
            alert(L('errorNoTable'));
        } else if (target.error === "multipleTables") {
            alert(L('errorMultipleTables'));
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
                var rows = resolveTargetRows(table, target.rowIndices, result.targetMode);
                applyRowHeightByMode(rows, result.height);
            }
        }

    })();
