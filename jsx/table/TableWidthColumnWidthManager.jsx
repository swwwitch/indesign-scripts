#target indesign
    (function () {
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;

        /*
         * 概要 / Overview
         * テーブル内の選択位置から対象テーブルを特定し、表全体の幅と列の幅をまとめて調整するスクリプトです。
         * ダイアログでは、表全体の幅を「変更しない」「自動調整」「親フレームいっぱいに」「指定」から選択でき、
         * 列の幅を「均等に」「自動調整」「自動調整を維持」「最終列のみ調整」「指定」から選択できます。
         * 常時プレビュー、選択復元、画面のプレビュー表示切り替えに対応し、指定値入力では
         * 定規の単位に応じた値入力と、↑↓キーでの増減（Shift: ±10、Option: ±0.1）も行えます。
         * 列の幅で「指定」を選択した場合は、表全体の幅より列の幅の指定を優先し、
         * 表全体の幅の指定値は「列の幅 × 列数」で自動更新されます。
         * 「自動調整を維持」では、自動調整で得た各列幅を基準に、表全体の幅の増減差分を列数で均等配分します。
         * 「最終列のみ調整」では、最終列以外の現在の幅を維持したまま、最終列だけを調整します。
         *
         * This script resolves the target table from the current selection inside a table
         * and adjusts the table width and column width together.
         * In the dialog, Table Width can be set to Do Not Change, Auto, Fit to Parent Frame,
         * or Custom, and Column Width can be set to Equal Widths, Fit to Content,
         * Fit to Content+, Adjust Last Column, or Custom.
         * It supports always-on preview, restoring the original selection, and toggling the screen preview mode.
         * For Custom inputs, the value uses the current ruler unit and can be changed with the arrow keys
         * (Shift: ±10, Option: ±0.1).
         * When Column Width is set to Custom, the column width takes precedence over Table Width,
         * and the Table Width Custom field is updated automatically as column width × number of columns.
         * In Fit to Content+, the fitted column widths are used as a base, and the table-width difference
         * is distributed evenly across all columns.
         * In Adjust Last Column, all non-last columns keep their current widths, and only the last column is adjusted.
         */

        // =========================================
        // バージョンとローカライズ / Version and localization
        // =========================================
        var SCRIPT_VERSION = "v1.0";

        function getCurrentLang() {
            return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
        }
        var lang = getCurrentLang();

        /* 日英ラベル定義 / Japanese-English label definitions */
        var LABELS = {
            dialogTitle: {
                ja: "表全体の幅、列の幅の調整",
                en: "Adjust Table and Column Widths"
            },
            errorSelectTable: {
                ja: "テーブル内にカーソルを置いてください。",
                en: "Place the cursor inside a table."
            },
            errorTextFrameNotFound: {
                ja: "テーブルがテキストフレーム内に見つかりません。",
                en: "The table was not found inside a text frame."
            },
            panelWidth: {
                ja: "表全体の幅",
                en: "Table Width"
            },
            panelCell: {
                ja: "列の幅",
                en: "Column Width"
            },
            widthKeep: {
                ja: "変更しない",
                en: "Do Not Change"
            },
            widthAuto: {
                ja: "自動調整",
                en: "Auto"
            },
            widthFit: {
                ja: "親フレームいっぱいに",
                en: "Fit to Parent Frame"
            },
            widthCustom: {
                ja: "指定",
                en: "Custom"
            },
            cellEqual: {
                ja: "均等に",
                en: "Equal Widths"
            },
            cellNatural: {
                ja: "自動調整",
                en: "Fit to Content"
            },
            cellNaturalPlus: {
                ja: "自動調整を維持",
                en: "Fit to Content+"
            },
            cellLast: {
                ja: "最終列のみ調整",
                en: "Adjust Last Column"
            },
            cellCustom: {
                ja: "指定",
                en: "Custom"
            },
            previewMode: {
                ja: "プレビュー",
                en: "Preview"
            },
            normalMode: {
                ja: "標準モード",
                en: "Normal Mode"
            },
            cancel: {
                ja: "キャンセル",
                en: "Cancel"
            },
            ok: {
                ja: "OK",
                en: "OK"
            }
        };

        function L(key) {
            return (LABELS[key] && LABELS[key][lang]) ? LABELS[key][lang] : key;
        }


        function getCurrentRulerUnitInfo() {
            var doc = app.activeDocument;
            var unit = doc.viewPreferences.horizontalMeasurementUnits;

            switch (unit) {
                case MeasurementUnits.POINTS:
                    return { label: "pt", unitValue: "pt" };
                case MeasurementUnits.PICAS:
                    return { label: "pica", unitValue: "pc" };
                case MeasurementUnits.INCHES:
                case MeasurementUnits.INCHES_DECIMAL:
                    return { label: "in", unitValue: "in" };
                case MeasurementUnits.MILLIMETERS:
                    return { label: "mm", unitValue: "mm" };
                case MeasurementUnits.CENTIMETERS:
                    return { label: "cm", unitValue: "cm" };
                default:
                    return { label: "pt", unitValue: "pt" };
            }
        }

        function convertToPointsByRulerUnit(value) {
            var unitInfo = getCurrentRulerUnitInfo();
            return UnitValue(value, unitInfo.unitValue).as("pt");
        }

        function roundToOneDecimal(value) {
            return Math.round(value * 10) / 10;
        }

        function getTotalTableWidth(table) {
            var totalWidth = 0;
            for (var columnIndex = 0; columnIndex < table.columns.length; columnIndex++) {
                totalWidth += table.columns[columnIndex].width;
            }
            return totalWidth;
        }

        function getFirstColumnWidth(table) {
            if (!table || !table.columns || table.columns.length === 0) return 0;
            return table.columns[0].width;
        }

        /*
         * メイン処理 / Main process
         */
        function runAdjustTableWidth() {
            var selection = app.selection;

            if (selection.length == 0) {
                alert(L("errorSelectTable"));
                return;
            }

            var selectionTarget = selection[0];
            var targetTable = null;
            var selectionType = selectionTarget.constructor.name;

            // テーブルを含むセルを見つける
            if (selectionType === "Cell" || selectionType === "Row" || selectionType === "Column") {
                targetTable = selectionTarget.parent;
            } else if (selectionType === "Table") {
                targetTable = selectionTarget;
            } else if (selectionType === "InsertionPoint") {
                if (selectionTarget.parent.constructor.name === "Cell") {
                    targetTable = selectionTarget.parent.parent;
                }
            } else if (selectionType === "Text" || selectionTarget.hasOwnProperty('baseline')) {
                if (selectionTarget.parentTextFrames.length > 0) {
                    var frame = selectionTarget.parentTextFrames[0];
                    if (frame.tables.length > 0) {
                        for (var i = 0; i < frame.tables.length; i++) {
                            var table = frame.tables[i];
                            if (table.storyOffset.index <= selectionTarget.index && selectionTarget.index <= table.storyOffset.index + table.characters.length) {
                                targetTable = table;
                                break;
                            }
                        }
                    }
                }
            }

            if (!targetTable) {
                alert(L("errorSelectTable"));
                return;
            }

            // 選択セルを記憶し、ハイライト表示を消す(プレビューが見やすくなるように)
            var savedSelection = [];
            for (var selectionIndex = 0; selectionIndex < selection.length; selectionIndex++) {
                savedSelection.push(selection[selectionIndex]);
            }
            try {
                app.select(NothingEnum.NOTHING);
            } catch (e) { /* 無視 */ }

            function restoreSelection() {
                try {
                    if (savedSelection.length > 0) {
                        app.selection = savedSelection;
                    }
                } catch (e) { /* 無視 */ }
            }

            // 元の状態をスナップショット(プレビュー復元用)
            var originalTableWidth = getTotalTableWidth(targetTable);
            var originalColumnWidths = [];
            var originalLeftInsets = [];
            var originalRightInsets = [];
            for (var columnIndex = 0; columnIndex < targetTable.columns.length; columnIndex++) {
                originalColumnWidths.push(targetTable.columns[columnIndex].width);
                originalLeftInsets.push(targetTable.columns[columnIndex].leftInset);
                originalRightInsets.push(targetTable.columns[columnIndex].rightInset);
            }

            function revert() {
                targetTable.width = originalTableWidth;
                for (var columnIndex = 0; columnIndex < targetTable.columns.length; columnIndex++) {
                    targetTable.columns[columnIndex].width = originalColumnWidths[columnIndex];
                    targetTable.columns[columnIndex].leftInset = originalLeftInsets[columnIndex];
                    targetTable.columns[columnIndex].rightInset = originalRightInsets[columnIndex];
                }
            }

            function applyTableWidthResult(result) {
                var currentPreviewColumnWidths = null;
                if (result.columnWidth && result.columnWidth.value === "last") {
                    currentPreviewColumnWidths = [];
                    for (var columnIndex = 0; columnIndex < targetTable.columns.length; columnIndex++) {
                        currentPreviewColumnWidths.push(targetTable.columns[columnIndex].width);
                    }
                }
                revert();

                // 1. 目標の表幅を算出(この時点では代入しない)
                var widthMode = result.tableWidth.value;
                var columnWidthMode = result.columnWidth.value;
                var targetTableWidth = getTotalTableWidth(targetTable); // default: keep
                var autoWidth = false;
                var originalAutoFitColumnWidths = null;
                var originalAutoFitTableWidth = null;

                if (columnWidthMode === "custom") {
                    if (isNaN(result.columnWidth.input) || result.columnWidth.input <= 0) return;
                    var customColumnWidth = convertToPointsByRulerUnit(result.columnWidth.input);
                    for (var columnIndex = 0; columnIndex < targetTable.columns.length; columnIndex++) {
                        targetTable.columns[columnIndex].width = customColumnWidth;
                    }
                    return;
                }

                if (columnWidthMode === "naturalPlus") {
                    fitColumnsToContent(targetTable, 2);
                    originalAutoFitColumnWidths = [];
                    for (var columnIndex = 0; columnIndex < targetTable.columns.length; columnIndex++) {
                        originalAutoFitColumnWidths.push(targetTable.columns[columnIndex].width);
                    }
                    originalAutoFitTableWidth = getTotalTableWidth(targetTable);
                }

                if (widthMode === "fit") {
                    var textFrame = targetTable.parent;
                    while (textFrame.constructor.name !== "TextFrame" && textFrame.constructor.name !== "Story") {
                        textFrame = textFrame.parent;
                    }
                    if (textFrame.constructor.name !== "TextFrame") {
                        alert(L("errorTextFrameNotFound"));
                        return;
                    }
                    targetTableWidth = textFrame.geometricBounds[3] - textFrame.geometricBounds[1];
                } else if (widthMode === "custom") {
                    if (isNaN(result.tableWidth.input) || result.tableWidth.input <= 0) return;
                    targetTableWidth = convertToPointsByRulerUnit(result.tableWidth.input);
                } else if (widthMode === "auto") {
                    autoWidth = true;
                }

                // 2. 列幅「最終列のみ調整」は特別処理
                // 非最終列は現在の幅を維持し、最終列だけで表全体の幅に合わせます。
                if (columnWidthMode === "last") {
                    if (widthMode === "fit" || widthMode === "custom") {
                        var baseColumnWidths = currentPreviewColumnWidths || originalColumnWidths;
                        var totalOtherColumnsWidth = 0;
                        for (var columnIndex = 0; columnIndex < targetTable.columns.length - 1; columnIndex++) {
                            totalOtherColumnsWidth += baseColumnWidths[columnIndex];
                        }

                        var newLastColumnWidth = targetTableWidth - totalOtherColumnsWidth;
                        if (newLastColumnWidth <= 0) return;

                        for (var columnIndex = 0; columnIndex < targetTable.columns.length - 1; columnIndex++) {
                            targetTable.columns[columnIndex].width = baseColumnWidths[columnIndex];
                        }
                        targetTable.columns[targetTable.columns.length - 1].width = newLastColumnWidth;
                    }
                    return;
                }

                // 3. 通常フロー: 幅を適用
                if (widthMode === "fit" || widthMode === "custom") {
                    targetTable.width = targetTableWidth;
                } else if (autoWidth && columnWidthMode !== "naturalPlus") {
                    fitColumnsToContent(targetTable, 2);
                }

                // 4. 列幅処理
                if (columnWidthMode === "equal") {
                    equalizeColumns(targetTable);
                } else if (columnWidthMode === "natural") {
                    fitColumnsToContent(targetTable, 2);
                } else if (columnWidthMode === "naturalPlus") {
                    var widthDelta = getTotalTableWidth(targetTable) - originalAutoFitTableWidth;
                    var deltaPerColumn = widthDelta / targetTable.columns.length;
                    for (var columnIndex = 0; columnIndex < targetTable.columns.length; columnIndex++) {
                        targetTable.columns[columnIndex].width = originalAutoFitColumnWidths[columnIndex] + deltaPerColumn;
                    }
                }
            }

            var rulerUnitInfo = getCurrentRulerUnitInfo();
            var currentColumnWidthValue = roundToOneDecimal(UnitValue(getFirstColumnWidth(targetTable), "pt").as(rulerUnitInfo.unitValue));
            var currentTableWidthValue = roundToOneDecimal(currentColumnWidthValue * targetTable.columns.length);
            var result = showMultiPanelOptionDialog(L("dialogTitle"), [
                {
                    key: "tableWidth",
                    label: L("panelWidth"),
                    options: [
                        { value: "keep", label: L("widthKeep"), defaultSelected: true },
                        { value: "auto", label: L("widthAuto") },
                        { value: "fit", label: L("widthFit") },
                        { value: "custom", label: L("widthCustom"), input: { suffix: rulerUnitInfo.label, defaultValue: currentTableWidthValue } }
                    ]
                },
                {
                    key: "columnWidth",
                    label: L("panelCell"),
                    options: [
                        { value: "equal", label: L("cellEqual"), defaultSelected: true },
                        { value: "natural", label: L("cellNatural"), linkedSelection: { panel: "tableWidth", value: "auto" } },
                        { value: "naturalPlus", label: L("cellNaturalPlus") },
                        { value: "last", label: L("cellLast"), linkedSelection: { panel: "tableWidth", value: "fit" } },
                        { value: "custom", label: L("cellCustom"), input: { suffix: rulerUnitInfo.label, defaultValue: currentColumnWidthValue } }
                    ]
                }
            ], applyTableWidthResult, revert, targetTable);

            if (result === null) {
                revert();
                restoreSelection();
                return;
            }

            applyTableWidthResult(result);
            restoreSelection();
        }

        /*
         * 列幅を均等化 / Equalize column widths
         */
        function equalizeColumns(table) {
            var avg = getTotalTableWidth(table) / table.columns.length;
            for (var j = 0; j < table.columns.length; j++) {
                table.columns[j].width = avg;
            }
        }


        /*
         * 内容に合わせて列幅を調整 / Fit column widths to content
         */
        function fitColumnsToContent(table, margin) {
            if (margin === undefined) margin = 2;
            var columns = table.columns;
            for (var columnIndex = 0, columnCount = columns.length; columnIndex < columnCount; columnIndex++) {
                var columnCells = columns[columnIndex].cells;
                var contentWidths = [];
                for (var cellIndex = 0, cellCount = columnCells.length; cellIndex < cellCount; cellIndex++) {
                    if (columnCells[cellIndex].texts[0].contents === "") continue;
                    if (columnCells[cellIndex].overflows) {
                        while (columnCells[cellIndex].overflows) {
                            columnCells[cellIndex].width += 1;
                            if (columnCells[cellIndex].properties['lines'] !== undefined) break;
                        }
                    }
                    var contentStartOffset = columnCells[cellIndex].lines[0].insertionPoints[0].horizontalOffset;
                    var contentEndOffset = columnCells[cellIndex].lines[0].insertionPoints[-1].horizontalOffset;
                    contentWidths.push(contentEndOffset - contentStartOffset);
                }
                columns[columnIndex].rightInset = columns[columnIndex].leftInset = margin * 1.0;
                var padding = columns[columnIndex].rightInset + columns[columnIndex].leftInset;
                var lineWeight = columns[columnIndex].rightEdgeStrokeWeight * 0.5 + columns[columnIndex].leftEdgeStrokeWeight * 0.5;
                if (contentWidths.length > 0) {
                    columns[columnIndex].width = contentWidths.sort(function (a, b) { return b > a; })[0] + padding + lineWeight;
                }
            }
        }

        /*
         * 画面プレビューモード補助 / Screen preview mode helpers
         * 他のスクリプトでも使い回せるように、現在の画面モード判定、
         * プレビュー表示と標準表示の切り替え、ボタンラベル更新をまとめています。
         *
         * Reusable helpers for checking the current screen mode,
         * toggling between Preview and Normal Mode,
         * and updating the toggle button label.
         */
        function isPreviewScreenMode() {
            try {
                return app.activeWindow && app.activeWindow.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE;
            } catch (e) {
                return false;
            }
        }

        function togglePreviewScreenMode() {
            try {
                var w = app.activeWindow;
                if (!w) return;
                if (w.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE) {
                    w.screenMode = ScreenModeOptions.PREVIEW_OFF;
                } else {
                    w.screenMode = ScreenModeOptions.PREVIEW_TO_PAGE;
                }
            } catch (e) { /* 無視 */ }
        }

        /*
         * 再利用用ラベル取得 / Get reusable toggle button label
         */
        function getPreviewToggleButtonLabel() {
            return isPreviewScreenMode() ? L("previewMode") : L("normalMode");
        }

        /*
         * ボタンラベル更新 / Update toggle button label
         */
        function updatePreviewToggleButtonLabel(btn) {
            btn.text = getPreviewToggleButtonLabel();
        }

        /*
         * 概要 / Overview
         * 複数パネルのラジオボタンUIを構築するダイアログです。
         * 各パネルの選択状態と入力値を収集し、常時プレビューの適用・復元、
         * linkedSelection による他パネル連動、画面のプレビュー表示切り替え、
         * 定規の単位に応じた入力欄表示、↑↓キーによる値変更、
         * 列の幅の指定時に表全体の幅より列の幅を優先する制御、
         * および依存オプションの有効・無効制御に対応します。
         * キャンセル時は null を返し、OK時は各パネルの結果をオブジェクトで返します。
         *
         * Builds a dialog that contains multiple radio-button panels.
         * It collects selected values and input values, supports applying and reverting always-on preview,
         * linked selections across panels, toggling the screen preview mode,
         * showing input units based on the current ruler unit, changing input values with the arrow keys,
         * prioritizing Column Width Custom input over Table Width,
         * and enabling or disabling dependent options as needed.
         * It returns null on cancel, or an object containing the selected results on OK.
         */
        function showMultiPanelOptionDialog(title, panels, onApply, onRevert, targetTable) {
            var dlg = new Window("dialog", title + " " + SCRIPT_VERSION);
            dlg.alignChildren = "fill";

            var previewToggleCount = 0;
            var panelStates = [];

            function buildPanelsRow(parent, panels) {
                var row = parent.add("group");
                row.orientation = "row";
                row.alignChildren = ["fill", "top"];
                row.alignment = ["fill", "top"];
                return row;
            }

            function createPanelState(panelDef) {
                return {
                    key: panelDef.key,
                    options: panelDef.options,
                    radios: [],
                    inputs: [],
                    rows: []
                };
            }

            function changeValueByArrowKey(editText, state, optionIndex) {
                editText.addEventListener("keydown", function (event) {
                    var value = Number(editText.text);
                    if (isNaN(value)) return;

                    var keyboard = ScriptUI.environment.keyboardState;
                    var delta = 1;

                    if (keyboard.shiftKey) {
                        delta = 10;
                        if (event.keyName == "Up") {
                            value = Math.ceil((value + 1) / delta) * delta;
                            event.preventDefault();
                        } else if (event.keyName == "Down") {
                            value = Math.floor((value - 1) / delta) * delta;
                            if (value < 0) value = 0;
                            event.preventDefault();
                        }
                    } else if (keyboard.altKey) {
                        delta = 0.1;
                        if (event.keyName == "Up") {
                            value += delta;
                            event.preventDefault();
                        } else if (event.keyName == "Down") {
                            value -= delta;
                            if (value < 0) value = 0;
                            event.preventDefault();
                        }
                    } else {
                        delta = 1;
                        if (event.keyName == "Up") {
                            value += delta;
                            event.preventDefault();
                        } else if (event.keyName == "Down") {
                            value -= delta;
                            if (value < 0) value = 0;
                            event.preventDefault();
                        }
                    }

                    if (keyboard.altKey) {
                        value = Math.round(value * 10) / 10;
                    } else {
                        value = Math.round(value);
                    }

                    selectRadio(state, optionIndex);
                    applyLinkedSelection(state.options[optionIndex]);
                    updateDependentUI();
                    editText.text = value;
                    syncTableWidthCustomInputFromColumnWidth();
                    if (typeof editText.onChange === "function") {
                        editText.onChange();
                    }
                });
            }

            function buildOptionRow(panel, state, option) {
                var row = panel.add("group");
                row.alignChildren = ["left", "center"];
                state.rows.push(row);

                var rb = row.add("radiobutton", undefined, option.label);
                state.radios.push(rb);

                if (option.input) {
                    var et = row.add("edittext", undefined, String(option.input.defaultValue));
                    changeValueByArrowKey(et, state, state.radios.length - 1);
                    et.characters = 4;
                    row.add("statictext", undefined, option.input.suffix || "");
                    state.inputs.push(et);
                } else {
                    state.inputs.push(null);
                }
            }

            function buildPanel(parent, panelDef) {
                var pan = parent.add("panel", undefined, panelDef.label);
                pan.alignChildren = "left";
                pan.alignment = ["fill", "top"];
                pan.margins = [15, 20, 15, 10];

                var state = createPanelState(panelDef);
                var defaultIndex = 0;

                for (var i = 0; i < panelDef.options.length; i++) {
                    buildOptionRow(pan, state, panelDef.options[i]);
                    if (panelDef.options[i].defaultSelected) defaultIndex = i;
                }

                state.radios[defaultIndex].value = true;
                return state;
            }

            function buildAllPanels(parent, panels) {
                var row = buildPanelsRow(parent, panels);
                for (var panelIndex = 0; panelIndex < panels.length; panelIndex++) {
                    panelStates.push(buildPanel(row, panels[panelIndex]));
                }
            }


            function collectDialogResult() {
                var result = {};
                for (var panelIndex = 0; panelIndex < panelStates.length; panelIndex++) {
                    var panelState = panelStates[panelIndex];
                    for (var optionIndex = 0; optionIndex < panelState.radios.length; optionIndex++) {
                        if (panelState.radios[optionIndex].value) {
                            var optionResult = { value: panelState.options[optionIndex].value };
                            if (panelState.inputs[optionIndex]) optionResult.input = parseFloat(panelState.inputs[optionIndex].text);
                            result[panelState.key] = optionResult;
                            break;
                        }
                    }
                }
                return result;
            }

            function applyPreviewFromDialog() {
                if (onApply) onApply(collectDialogResult());
            }

            function selectRadio(state, idx) {
                for (var radioIndex = 0; radioIndex < state.radios.length; radioIndex++) {
                    state.radios[radioIndex].value = (radioIndex === idx);
                }
            }

            function findPanelState(panelKey) {
                for (var i = 0; i < panelStates.length; i++) {
                    if (panelStates[i].key === panelKey) return panelStates[i];
                }
                return null;
            }

            function selectInPanel(panelKey, targetValue) {
                var targetState = findPanelState(panelKey);
                if (!targetState) return;

                for (var i = 0; i < targetState.options.length; i++) {
                    if (targetState.options[i].value === targetValue) {
                        selectRadio(targetState, i);
                        return;
                    }
                }
            }

            function applyLinkedSelection(option) {
                if (option && option.linkedSelection) {
                    selectInPanel(option.linkedSelection.panel, option.linkedSelection.value);
                }
            }

            function setOptionEnabled(state, optionValue, isEnabled) {
                if (!state) return;

                for (var optionIndex = 0; optionIndex < state.options.length; optionIndex++) {
                    if (state.options[optionIndex].value === optionValue) {
                        if (state.radios[optionIndex]) state.radios[optionIndex].enabled = isEnabled;
                        if (state.inputs[optionIndex]) state.inputs[optionIndex].enabled = isEnabled;
                        if (state.inputs[optionIndex]) state.inputs[optionIndex].readonly = !isEnabled;
                        if (state.rows[optionIndex]) state.rows[optionIndex].enabled = isEnabled;
                        return;
                    }
                }
            }

            function getSelectedOptionValue(state) {
                if (!state) return null;
                for (var optionIndex = 0; optionIndex < state.radios.length; optionIndex++) {
                    if (state.radios[optionIndex].value) {
                        return state.options[optionIndex].value;
                    }
                }
                return null;
            }

            function resetDependentUI(tableWidthState, columnWidthState) {
                setOptionEnabled(tableWidthState, "keep", true);
                setOptionEnabled(tableWidthState, "auto", true);
                setOptionEnabled(tableWidthState, "fit", true);
                setOptionEnabled(tableWidthState, "custom", true);
                setOptionEnabled(columnWidthState, "equal", true);
                setOptionEnabled(columnWidthState, "natural", true);
                setOptionEnabled(columnWidthState, "naturalPlus", true);
                setOptionEnabled(columnWidthState, "last", true);
                setOptionEnabled(columnWidthState, "custom", true);
            }

            function applyTableWidthDependencies(selectedTableWidthValue, selectedColumnWidthValue, columnWidthState) {
            }

            function applyColumnWidthDependencies(selectedColumnWidthValue, tableWidthState) {
                if (selectedColumnWidthValue === "equal") {
                    setOptionEnabled(tableWidthState, "auto", false);
                }

                if (selectedColumnWidthValue === "last") {
                    selectInPanel("tableWidth", "fit");
                    setOptionEnabled(tableWidthState, "keep", false);
                    setOptionEnabled(tableWidthState, "auto", false);
                    setOptionEnabled(tableWidthState, "custom", false);
                }

                if (selectedColumnWidthValue === "custom") {
                    setOptionEnabled(tableWidthState, "keep", false);
                    setOptionEnabled(tableWidthState, "auto", false);
                    setOptionEnabled(tableWidthState, "fit", false);
                    setOptionEnabled(tableWidthState, "custom", false);
                }
            }

            function syncTableWidthCustomInputFromColumnWidth() {
                var tableWidthState = findPanelState("tableWidth");
                var columnWidthState = findPanelState("columnWidth");
                if (!tableWidthState || !columnWidthState) return;

                var tableWidthCustomIndex = -1;
                for (var optionIndex = 0; optionIndex < tableWidthState.options.length; optionIndex++) {
                    if (tableWidthState.options[optionIndex].value === "custom") {
                        tableWidthCustomIndex = optionIndex;
                        break;
                    }
                }

                var columnWidthCustomIndex = -1;
                for (var optionIndex = 0; optionIndex < columnWidthState.options.length; optionIndex++) {
                    if (columnWidthState.options[optionIndex].value === "custom") {
                        columnWidthCustomIndex = optionIndex;
                        break;
                    }
                }

                if (tableWidthCustomIndex < 0 || columnWidthCustomIndex < 0) return;
                if (!tableWidthState.inputs[tableWidthCustomIndex] || !columnWidthState.inputs[columnWidthCustomIndex]) return;
                if (!columnWidthState.radios[columnWidthCustomIndex].value) return;
                if (!targetTable || !targetTable.columns || targetTable.columns.length <= 0) return;

                var columnWidthValue = Number(columnWidthState.inputs[columnWidthCustomIndex].text);
                if (isNaN(columnWidthValue) || columnWidthValue <= 0) return;

                var syncedTableWidthValue = roundToOneDecimal(columnWidthValue * targetTable.columns.length);
                tableWidthState.inputs[tableWidthCustomIndex].text = String(syncedTableWidthValue);
                try {
                    dlg.layout.layout(true);
                    dlg.update();
                } catch (e) { }
            }

            function updateDependentUI() {
                var tableWidthState = findPanelState("tableWidth");
                var columnWidthState = findPanelState("columnWidth");
                if (!tableWidthState || !columnWidthState) return;

                var selectedTableWidthValue = getSelectedOptionValue(tableWidthState);
                var selectedColumnWidthValue = getSelectedOptionValue(columnWidthState);

                resetDependentUI(tableWidthState, columnWidthState);
                applyTableWidthDependencies(selectedTableWidthValue, selectedColumnWidthValue, columnWidthState);
                applyColumnWidthDependencies(selectedColumnWidthValue, tableWidthState);
            }

            function bindStateEvents(state) {
                for (var radioIndex = 0; radioIndex < state.radios.length; radioIndex++) {
                    (function (optionIndex) {
                        state.radios[optionIndex].onClick = function () {
                            selectRadio(state, optionIndex);
                            applyLinkedSelection(state.options[optionIndex]);
                            updateDependentUI();
                            syncTableWidthCustomInputFromColumnWidth();
                            applyPreviewFromDialog();
                        };
                    })(radioIndex);
                }

                for (var inputIndex = 0; inputIndex < state.inputs.length; inputIndex++) {
                    if (state.inputs[inputIndex]) {
                        (function (optionIndex) {
                            state.inputs[optionIndex].onChange = function () {
                                updateDependentUI();
                                syncTableWidthCustomInputFromColumnWidth();
                                applyPreviewFromDialog();
                            };
                            state.inputs[optionIndex].onChanging = function () {
                                selectRadio(state, optionIndex);
                                applyLinkedSelection(state.options[optionIndex]);
                                updateDependentUI();
                                syncTableWidthCustomInputFromColumnWidth();
                            };
                        })(inputIndex);
                    }
                }
            }

            function bindAllStateEvents() {
                for (var panelIndex = 0; panelIndex < panelStates.length; panelIndex++) {
                    bindStateEvents(panelStates[panelIndex]);
                }
            }

            function createBottomButtons(parent) {
                var bottomGroup = parent.add("group");
                bottomGroup.orientation = "row";
                bottomGroup.alignChildren = "fill";
                bottomGroup.alignment = ["fill", "top"];
                bottomGroup.margins = [0, 8, 0, 0];

                var btnPreviewToggle = bottomGroup.add("button", undefined, getPreviewToggleButtonLabel());
                btnPreviewToggle.alignment = ["left", "center"];

                var spacer = bottomGroup.add("group");
                spacer.alignment = ["fill", "top"];

                var buttonGroup = bottomGroup.add("group");
                buttonGroup.orientation = "row";
                buttonGroup.alignment = ["right", "top"];
                buttonGroup.alignChildren = "right";
                var cancelBtn = buttonGroup.add("button", undefined, L("cancel"), { name: "cancel" });
                var okBtn = buttonGroup.add("button", undefined, L("ok"), { name: "ok" });

                return {
                    previewToggleButton: btnPreviewToggle,
                    cancelButton: cancelBtn,
                    okButton: okBtn
                };
            }

            function bindPreviewToggleButton(btn) {
                btn.onClick = function () {
                    togglePreviewScreenMode();
                    previewToggleCount++;
                    updatePreviewToggleButtonLabel(btn);
                };
            }


            buildAllPanels(dlg, panels);

            bindAllStateEvents();
            updateDependentUI();
            syncTableWidthCustomInputFromColumnWidth();

            var buttons = createBottomButtons(dlg);
            bindPreviewToggleButton(buttons.previewToggleButton);

            applyPreviewFromDialog();

            var shown = dlg.show();

            if (previewToggleCount % 2 === 1) {
                togglePreviewScreenMode();
            }

            if (shown !== 1) return null;
            return collectDialogResult();
        }

        runAdjustTableWidth();
    })();