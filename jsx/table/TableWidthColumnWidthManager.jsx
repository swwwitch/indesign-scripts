#target indesign

/*
 * TableWidthColumnWidthManager.jsx
 *
 * 選択位置から対象の表を特定し、表全体の幅と列の幅をプレビュー付きでまとめて調整します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TableWidthColumnWidthManager"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-05-05";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/TableWidthColumnWidthManager.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/TableWidthColumnWidthManager.md

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
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;

        // =========================================
        // ユーザー設定 / User settings
        // =========================================

        /* ↑↓キーでの増減幅（通常 / Shift / Option）/ Arrow-key steps (normal, Shift, Option) */
        var ARROW_STEP_NORMAL = 1;
        var ARROW_STEP_SHIFT  = 10;
        var ARROW_STEP_OPTION = 0.1;

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
                title: { ja: "表全体の幅、列の幅の調整", en: "Adjust Table and Column Widths" }
            },
            panel: {
                tableWidth:  { ja: "表全体の幅", en: "Table Width" },
                columnWidth: { ja: "列の幅", en: "Column Width" }
            },
            radio: {
                widthKeep:       { ja: "変更しない", en: "Do Not Change" },
                widthAuto:       { ja: "自動調整", en: "Auto" },
                widthFit:        { ja: "親フレームいっぱいに", en: "Fit to Parent Frame" },
                widthCustom:     { ja: "指定", en: "Custom" },
                cellEqual:       { ja: "均等に", en: "Equal Widths" },
                cellNatural:     { ja: "自動調整", en: "Fit to Content" },
                cellNaturalPlus: { ja: "自動調整を維持", en: "Fit to Content+" },
                cellLast:        { ja: "最終列のみ調整", en: "Adjust Last Column" },
                cellCustom:      { ja: "指定", en: "Custom" }
            },
            button: {
                ok:          { ja: "OK", en: "OK" },
                cancel:      { ja: "キャンセル", en: "Cancel" },
                previewMode: { ja: "プレビュー", en: "Preview" },
                normalMode:  { ja: "標準モード", en: "Normal Mode" }
            },
            undo: {
                applyWidths: { ja: "表全体の幅、列の幅の調整", en: "Adjust Table and Column Widths" }
            },
            error: {
                selectTable:        { ja: "テーブル内にカーソルを置いてください。", en: "Place the cursor inside a table." },
                textFrameNotFound:  { ja: "テーブルがテキストフレーム内に見つかりません。", en: "The table was not found inside a text frame." }
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

        /**
         * 現在の定規単位のラベルと単位記号を取得する
         * @returns {{label: string, unitValue: string}} 単位の情報
         */
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

        /**
         * 小数第 1 位に丸める
         * @param {number} value 対象の数値
         * @returns {number} 丸めた数値
         */
        function roundToOneDecimal(value) {
            return Math.round(value * 10) / 10;
        }

        /**
         * 表の 1 列目の幅を取得する
         * @param {Table} table 対象の表
         * @returns {number} 列幅。取得できない場合は 0
         */
        function getFirstColumnWidth(table) {
            if (!table || !table.columns || table.columns.length === 0) return 0;
            return table.columns[0].width;
        }

        /**
         * 選択から表を特定し、幅調整ダイアログを表示して適用する
         * @returns {void}
         */
        function runAdjustTableWidth() {
            var selection = app.selection;

            if (selection.length == 0) {
                alert(getLabel("error.selectTable"));
                return;
            }

            var selectionTarget = selection[0];
            var targetTable = null;
            var selectionType = selectionTarget.constructor.name;

            // 選択範囲から表オブジェクトを特定する
            if (selectionType === "TextFrame") {
                if (selectionTarget.tables.length > 0) {
                    targetTable = selectionTarget.tables[0];
                }
            } else if (selectionType === "Cell" || selectionType === "Row" || selectionType === "Column" || selectionType === "InsertionPoint" || selectionType === "Text") {
                try {
                    targetTable = selectionTarget.parent;
                    while (targetTable && targetTable.constructor.name !== "Table") {
                        targetTable = targetTable.parent;
                    }
                } catch (e) {
                    targetTable = null;
                }
            } else if (selectionType === "Table") {
                targetTable = selectionTarget;
            }

            if (!targetTable) {
                alert(getLabel("error.selectTable"));
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

            /**
             * 実行前の選択状態を復元する
             * @returns {void}
             */
            function restoreSelection() {
                try {
                    if (savedSelection.length > 0) {
                        app.selection = savedSelection;
                    }
                } catch (e) { /* 無視 */ }
            }

            // 元の状態をスナップショット(プレビュー復元用)
            var originalTableWidth = targetTable.width;
            var originalColumnWidths = [];
            var originalLeftInsets = [];
            var originalRightInsets = [];
            for (var columnIndex = 0; columnIndex < targetTable.columns.length; columnIndex++) {
                originalColumnWidths.push(targetTable.columns[columnIndex].width);
                originalLeftInsets.push(targetTable.columns[columnIndex].leftInset);
                originalRightInsets.push(targetTable.columns[columnIndex].rightInset);
            }

            /**
             * プレビューで加えた変更を元に戻す
             * @returns {void}
             */
            function revert() {
                targetTable.width = originalTableWidth;
                for (var columnIndex = 0; columnIndex < targetTable.columns.length; columnIndex++) {
                    targetTable.columns[columnIndex].width = originalColumnWidths[columnIndex];
                    targetTable.columns[columnIndex].leftInset = originalLeftInsets[columnIndex];
                    targetTable.columns[columnIndex].rightInset = originalRightInsets[columnIndex];
                }
            }

            /**
             * ダイアログの結果に従って表と列の幅を適用する
             * @param {object} result ダイアログが返した設定
             * @returns {void}
             */
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
                var targetTableWidth = targetTable.width; // default: keep
                var autoWidth = false;
                var originalAutoFitColumnWidths = null;
                var originalAutoFitTableWidth = null;

                if (columnWidthMode === "custom") {
                    if (isNaN(result.columnWidth.input) || result.columnWidth.input <= 0) return;
                    var customColumnWidth = result.columnWidth.input;
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
                    originalAutoFitTableWidth = targetTable.width;
                }

                if (widthMode === "fit") {
                    var textFrame = targetTable.parent;
                    while (textFrame.constructor.name !== "TextFrame" && textFrame.constructor.name !== "Story") {
                        textFrame = textFrame.parent;
                    }
                    if (textFrame.constructor.name !== "TextFrame") {
                        alert(getLabel("error.textFrameNotFound"));
                        return;
                    }
                    targetTableWidth = textFrame.geometricBounds[3] - textFrame.geometricBounds[1];
                } else if (widthMode === "custom") {
                    if (isNaN(result.tableWidth.input) || result.tableWidth.input <= 0) return;
                    targetTableWidth = result.tableWidth.input;
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
                    var widthDelta = targetTable.width - originalAutoFitTableWidth;
                    var deltaPerColumn = widthDelta / targetTable.columns.length;
                    for (var columnIndex = 0; columnIndex < targetTable.columns.length; columnIndex++) {
                        targetTable.columns[columnIndex].width = originalAutoFitColumnWidths[columnIndex] + deltaPerColumn;
                    }
                }
            }

            var rulerUnitInfo = getCurrentRulerUnitInfo();
            var currentColumnWidthValue = roundToOneDecimal(getFirstColumnWidth(targetTable));
            var currentTableWidthValue = roundToOneDecimal(targetTable.width);
            var result = showMultiPanelOptionDialog(getLabel("dialog.title"), [
                {
                    key: "tableWidth",
                    label: getLabel("panel.tableWidth"),
                    options: [
                        { value: "keep", label: getLabel("radio.widthKeep"), defaultSelected: true },
                        { value: "auto", label: getLabel("radio.widthAuto") },
                        { value: "fit", label: getLabel("radio.widthFit") },
                        { value: "custom", label: getLabel("radio.widthCustom"), input: { suffix: rulerUnitInfo.label, defaultValue: currentTableWidthValue } }
                    ]
                },
                {
                    key: "columnWidth",
                    label: getLabel("panel.columnWidth"),
                    options: [
                        { value: "equal", label: getLabel("radio.cellEqual"), defaultSelected: true },
                        { value: "natural", label: getLabel("radio.cellNatural"), linkedSelection: { panel: "tableWidth", value: "auto" } },
                        { value: "naturalPlus", label: getLabel("radio.cellNaturalPlus") },
                        { value: "last", label: getLabel("radio.cellLast"), linkedSelection: { panel: "tableWidth", value: "fit" } },
                        { value: "custom", label: getLabel("radio.cellCustom"), input: { suffix: rulerUnitInfo.label, defaultValue: currentColumnWidthValue } }
                    ]
                }
            ], applyTableWidthResult, revert, targetTable);

            if (result === null) {
                revert();
                restoreSelection();
                return;
            }

            /* プレビューを一度戻してから、確定分だけを 1 つの取り消しにまとめる
               / Revert the preview first so only the confirmed change forms the undo step */
            revert();
            app.doScript(function () {
                applyTableWidthResult(result);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.applyWidths"));
            restoreSelection();
        }

        /**
         * すべての列幅を均等にする
         * @param {Table} table 対象の表
         * @returns {void}
         */
        function equalizeColumns(table) {
            var avg = table.width / table.columns.length;
            for (var j = 0; j < table.columns.length; j++) {
                table.columns[j].width = avg;
            }
        }


        /**
         * 組版後の 1 行の幅を求める
         * @param {Line} line 対象の行
         * @returns {number} 行の幅
         */
        function getComposedLineWidth(line) {
            if (!line) return 0;
            try {
                var lineStart = line.insertionPoints[0].horizontalOffset;
                var lineEnd = line.insertionPoints[-1].horizontalOffset;
                return lineEnd - lineStart;
            } catch (e) {
                return 0;
            }
        }

        /**
         * セル内で最も長い行の幅を求める
         * @param {Cell} cell 対象のセル
         * @returns {number} 最大の行幅
         */
        function getMaxComposedLineWidthInCell(cell) {
            if (!cell || !cell.lines || cell.lines.length === 0) return 0;

            var cellLines = cell.lines;
            var maxContentWidth = 0;
            for (var lineIndex = 0; lineIndex < cellLines.length; lineIndex++) {
                var lineWidth = getComposedLineWidth(cellLines[lineIndex]);
                if (lineWidth > maxContentWidth) maxContentWidth = lineWidth;
            }
            return maxContentWidth;
        }

        /**
         * 各列の幅を内容に合わせて調整する
         * @param {Table} table 対象の表
         * @param {number} margin 内容に加える余白
         * @returns {Array<number>} 調整後の列幅
         */
        function fitColumnsToContent(table, margin) {
            if (margin === undefined) margin = 2;

            // 計測前に表を親フレームいっぱいまで広げ、列を均等化して余裕を作る
            var textFrame = table.parent;
            while (textFrame && textFrame.constructor.name !== "TextFrame" && textFrame.constructor.name !== "Story") {
                textFrame = textFrame.parent;
            }
            if (textFrame && textFrame.constructor.name === "TextFrame") {
                var parentFrameWidth = textFrame.geometricBounds[3] - textFrame.geometricBounds[1];
                if (parentFrameWidth > 0) {
                    table.width = parentFrameWidth;
                    var evenColumnWidth = parentFrameWidth / table.columns.length;
                    for (var equalIndex = 0; equalIndex < table.columns.length; equalIndex++) {
                        table.columns[equalIndex].width = evenColumnWidth;
                    }
                }
            }

            var columns = table.columns;
            for (var columnIndex = 0, columnCount = columns.length; columnIndex < columnCount; columnIndex++) {
                var columnCells = columns[columnIndex].cells;
                var contentWidths = [];
                for (var cellIndex = 0, cellCount = columnCells.length; cellIndex < cellCount; cellIndex++) {
                    if (columnCells[cellIndex].texts[0].contents === "") continue;

                    // 複数行に渡るセル（ハードリターンや折り返し）も想定し、全行中の最大幅を採用
                    contentWidths.push(getMaxComposedLineWidthInCell(columnCells[cellIndex]));
                }
                columns[columnIndex].rightInset = columns[columnIndex].leftInset = margin * 1.0;
                var padding = columns[columnIndex].rightInset + columns[columnIndex].leftInset;
                var lineWeight = columns[columnIndex].rightEdgeStrokeWeight * 0.5 + columns[columnIndex].leftEdgeStrokeWeight * 0.5;
                if (contentWidths.length > 0) {
                    columns[columnIndex].width = contentWidths.sort(function (a, b) { return b - a; })[0] + padding + lineWeight;
                }
            }
        }

        /**
         * 現在プレビュー表示になっているかを判定する
         * @returns {boolean} プレビュー表示なら true
         */
        function isPreviewScreenMode() {
            try {
                return app.activeWindow && app.activeWindow.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE;
            } catch (e) {
                return false;
            }
        }

        /**
         * 標準表示とプレビュー表示を切り替える
         * @returns {void}
         */
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

        /**
         * 画面モードに応じたトグルボタンのラベルを返す
         * @returns {string} ボタンに表示する文字列
         */
        function getPreviewToggleButtonLabel() {
            return isPreviewScreenMode() ? getLabel("button.previewMode") : getLabel("button.normalMode");
        }

        /**
         * トグルボタンのラベルを現在の画面モードに合わせて更新する
         * @param {Button} btn 対象のボタン
         * @returns {void}
         */
        function updatePreviewToggleButtonLabel(btn) {
            btn.text = getPreviewToggleButtonLabel();
        }

        /**
         * 複数のラジオボタンパネルを持つ設定ダイアログを表示する
         * @param {string} title ダイアログのタイトル
         * @param {Array<object>} panels パネルの定義
         * @param {function} onApply プレビュー適用時に呼ぶ処理
         * @param {function} onRevert プレビュー復元時に呼ぶ処理
         * @param {Table} targetTable 対象の表
         * @returns {object|null} 各パネルの選択結果。キャンセル時は null
         */
        function showMultiPanelOptionDialog(title, panels, onApply, onRevert, targetTable) {
            var dlg = new Window("dialog", title + " " + SCRIPT_VERSION);
            setupWindow(dlg, 10);

            var previewToggleCount = 0;
            var panelStates = [];

            /**
             * パネルを横に並べる行グループを作る
             * @param {object} parent 追加先のコンテナ
             * @param {Array<object>} panels パネルの定義
             * @returns {Group} 行グループ
             */
            function buildPanelsRow(parent, panels) {
                var row = parent.add("group");
                setupRow(row, "fill", COLUMN_SPACING);
                row.alignChildren = ["fill", "top"];
                return row;
            }

            /**
             * 1 パネル分の状態オブジェクトを作る
             * @param {object} panelDef パネルの定義
             * @returns {object} パネルの状態
             */
            function createPanelState(panelDef) {
                return {
                    key: panelDef.key,
                    options: panelDef.options,
                    radios: [],
                    inputs: [],
                    rows: []
                };
            }

            /**
             * 入力欄に上下キーでの増減操作を追加する
             * @param {EditText} editText 対象の入力欄
             * @param {object} state パネルの状態
             * @param {number} optionIndex 対象オプションの位置
             * @returns {void}
             */
            function changeValueByArrowKey(editText, state, optionIndex) {
                editText.addEventListener("keydown", function (event) {
                    var value = Number(editText.text);
                    if (isNaN(value)) return;

                    var keyboard = ScriptUI.environment.keyboardState;
                    var delta = ARROW_STEP_NORMAL;

                    if (keyboard.shiftKey) {
                        delta = ARROW_STEP_SHIFT;
                        if (event.keyName == "Up") {
                            value = Math.ceil((value + 1) / delta) * delta;
                            event.preventDefault();
                        } else if (event.keyName == "Down") {
                            value = Math.floor((value - 1) / delta) * delta;
                            if (value < 0) value = 0;
                            event.preventDefault();
                        }
                    } else if (keyboard.altKey) {
                        delta = ARROW_STEP_OPTION;
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

            /**
             * パネル内のオプション 1 行を組み立てる
             * @param {Panel} panel 追加先のパネル
             * @param {object} state パネルの状態
             * @param {object} option オプションの定義
             * @returns {void}
             */
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

            /**
             * パネル 1 つを組み立てる
             * @param {object} parent 追加先のコンテナ
             * @param {object} panelDef パネルの定義
             * @returns {object} パネルの状態
             */
            function buildPanel(parent, panelDef) {
                var optionPanel = parent.add("panel", undefined, panelDef.label);
                setupPanel(optionPanel, 6);
                optionPanel.alignChildren = "left";

                var state = createPanelState(panelDef);
                var defaultIndex = 0;

                for (var i = 0; i < panelDef.options.length; i++) {
                    buildOptionRow(optionPanel, state, panelDef.options[i]);
                    if (panelDef.options[i].defaultSelected) defaultIndex = i;
                }

                state.radios[defaultIndex].value = true;
                return state;
            }

            /**
             * すべてのパネルを組み立てる
             * @param {object} parent 追加先のコンテナ
             * @param {Array<object>} panels パネルの定義
             * @returns {void}
             */
            function buildAllPanels(parent, panels) {
                var row = buildPanelsRow(parent, panels);
                for (var panelIndex = 0; panelIndex < panels.length; panelIndex++) {
                    panelStates.push(buildPanel(row, panels[panelIndex]));
                }
            }


            /**
             * 各パネルの選択結果をまとめて返す
             * @returns {object} パネルごとの選択結果
             */
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

            /**
             * 現在のダイアログ設定でプレビューを適用する
             * @returns {void}
             */
            function applyPreviewFromDialog() {
                if (onApply) onApply(collectDialogResult());
            }

            /**
             * 指定した位置のラジオボタンを選択状態にする
             * @param {object} state パネルの状態
             * @param {number} idx 選択する位置
             * @returns {void}
             */
            function selectRadio(state, idx) {
                for (var radioIndex = 0; radioIndex < state.radios.length; radioIndex++) {
                    state.radios[radioIndex].value = (radioIndex === idx);
                }
            }

            /**
             * キーからパネルの状態を探す
             * @param {string} panelKey パネルのキー
             * @returns {object|null} パネルの状態。見つからない場合は null
             */
            function findPanelState(panelKey) {
                for (var i = 0; i < panelStates.length; i++) {
                    if (panelStates[i].key === panelKey) return panelStates[i];
                }
                return null;
            }

            /**
             * 指定したパネルで、値に対応するオプションを選択する
             * @param {string} panelKey パネルのキー
             * @param {string} targetValue 選択したい値
             * @returns {void}
             */
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

            /**
             * 連動指定のあるオプションを他パネルへ反映する
             * @param {object} option 選択されたオプション
             * @returns {void}
             */
            function applyLinkedSelection(option) {
                if (option && option.linkedSelection) {
                    selectInPanel(option.linkedSelection.panel, option.linkedSelection.value);
                }
            }

            /**
             * 指定したオプションの有効／無効を切り替える
             * @param {object} state パネルの状態
             * @param {string} optionValue 対象の値
             * @param {boolean} isEnabled 有効にするなら true
             * @returns {void}
             */
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

            /**
             * パネルで選択中のオプション値を取得する
             * @param {object} state パネルの状態
             * @returns {string|null} 選択中の値。なければ null
             */
            function getSelectedOptionValue(state) {
                if (!state) return null;
                for (var optionIndex = 0; optionIndex < state.radios.length; optionIndex++) {
                    if (state.radios[optionIndex].value) {
                        return state.options[optionIndex].value;
                    }
                }
                return null;
            }

            /**
             * 依存関係で無効化したオプションを既定状態に戻す
             * @param {object} tableWidthState 表全体の幅パネルの状態
             * @param {object} columnWidthState 列の幅パネルの状態
             * @returns {void}
             */
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

            /**
             * 表全体の幅の選択に応じて列の幅パネルを制御する
             * @param {string} selectedTableWidthValue 表全体の幅の選択値
             * @param {string} selectedColumnWidthValue 列の幅の選択値
             * @param {object} columnWidthState 列の幅パネルの状態
             * @returns {void}
             */
            function applyTableWidthDependencies(selectedTableWidthValue, selectedColumnWidthValue, columnWidthState) {
            }

            /**
             * 列の幅の選択に応じて表全体の幅パネルを制御する
             * @param {string} selectedColumnWidthValue 列の幅の選択値
             * @param {object} tableWidthState 表全体の幅パネルの状態
             * @returns {void}
             */
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

            /**
             * 列の幅の指定値から、表全体の幅の指定値を更新する
             * @returns {void}
             */
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

            /**
             * パネル間の依存関係をまとめて反映する
             * @returns {void}
             */
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

            /**
             * 1 パネル分のイベントを結び付ける
             * @param {object} state パネルの状態
             * @returns {void}
             */
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

            /**
             * すべてのパネルのイベントを結び付ける
             * @returns {void}
             */
            function bindAllStateEvents() {
                for (var panelIndex = 0; panelIndex < panelStates.length; panelIndex++) {
                    bindStateEvents(panelStates[panelIndex]);
                }
            }

            /**
             * ダイアログ下部のボタン行を組み立てる
             * @param {object} parent 追加先のコンテナ
             * @returns {object} 生成したボタン
             */
            function createBottomButtons(parent) {
                var bottomGroup = parent.add("group");
                setupRow(bottomGroup, "fill", 8);
                bottomGroup.alignChildren = "fill";
                bottomGroup.margins = [0, 8, 0, 0];

                var btnPreviewToggle = bottomGroup.add("button", undefined, getPreviewToggleButtonLabel());
                btnPreviewToggle.alignment = ["left", "center"];

                var spacer = bottomGroup.add("group");
                spacer.alignment = ["fill", "top"];

                /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
                var buttonGroup = bottomGroup.add("group");
                setupRow(buttonGroup, "right", 8);
                buttonGroup.alignChildren = "right";
                var cancelBtn = buttonGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
                var okBtn = buttonGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

                return {
                    previewToggleButton: btnPreviewToggle,
                    cancelButton: cancelBtn,
                    okButton: okBtn
                };
            }

            /**
             * 画面モード切り替えボタンにイベントを結び付ける
             * @param {Button} btn 対象のボタン
             * @returns {void}
             */
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