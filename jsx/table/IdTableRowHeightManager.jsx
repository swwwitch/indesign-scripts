#target indesign

/*
 * IdTableRowHeightManager.jsx
 *
 * 選択した表の行の高さを、範囲（選択範囲／ストーリー／ドキュメント）と対象行を指定しながらプレビュー付きで設定します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTableRowHeightManager";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-06-09";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableRowHeightManager.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableRowHeightManager.md

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
        // ユーザー設定 / User settings
        // =========================================

        /* 行の高さの下限（pt）。InDesign の最小行高（約 0.0139 inch ≒ 1pt）に合わせる
           / Minimum row height in points; InDesign's own minimum is about 0.0139 inch (≈ 1pt) */
        var MIN_ROW_HEIGHT_PT = 1.0008;

        // =========================================
        // レイアウト設定 / Layout settings
        // =========================================

        /* 行高入力欄の文字数 / Character width of the row-height field */
        var ROW_HEIGHT_INPUT_CHARACTERS = 5;

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
        var currentLanguage = getCurrentLang();

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
            },
            undo: {
                applyRowHeight: { ja: "行の高さを設定", en: "Set Row Height" }
            }
        };

        /**
         * ドット区切りキーでラベルを取得する（{slash} は / に置換）
         * @param {string} labelKey 例: "dialog.title"
         * @returns {string} 現在の言語のラベル文字列。見つからない場合はキーをそのまま返す
         */
        function getLabel(labelKey) {
            var keyParts = labelKey.split(".");
            var node = LABELS;
            for (var i = 0; i < keyParts.length; i++) {
                if (node && node.hasOwnProperty(keyParts[i])) {
                    node = node[keyParts[i]];
                } else {
                    return labelKey;
                }
            }
            var labelText = (node && node[currentLanguage]) ? node[currentLanguage] : labelKey;
            return labelText.replace(/\{slash\}/g, "/");
        }

        // =========================================
        // 単位 / Units
        // =========================================

        /**
         * 環境設定の縦方向単位のラベルを取得する
         * @returns {string} 単位のラベル
         */
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

        /**
         * ドキュメントの縦方向の単位を取得する
         * @returns {MeasurementUnits} 縦方向の単位
         */
        function getDocumentVerticalUnit() {
            try {
                return app.activeDocument.viewPreferences.verticalMeasurementUnits;
            } catch (e) {
                return MeasurementUnits.POINTS;
            }
        }

        /**
         * 1 単位あたりのポイント数を返す
         * @param {MeasurementUnits} unit 対象の単位
         * @returns {number} 1 単位あたりのポイント数
         */
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

        /**
         * 指定単位の数値をポイントへ換算する
         * @param {number} value 換算元の数値
         * @param {MeasurementUnits} unit 換算元の単位
         * @returns {number} ポイント値
         */
        function convertToPointsByUnit(value, unit) {
            return value * getPointsPerUnit(unit);
        }

        /**
         * ポイント値を指定単位の数値へ換算する
         * @param {number} value ポイント値
         * @param {MeasurementUnits} unit 換算先の単位
         * @returns {number} 指定単位での数値
         */
        function convertFromPointsByUnit(value, unit) {
            return value / getPointsPerUnit(unit);
        }

        /**
         * 入力欄に表示する数値を整形する
         * @param {number} value 表示する数値
         * @returns {string} 整形した文字列
         */
        function formatDisplayValue(value) {
            if (Math.abs(value - Math.round(value)) < 0.0001) return String(Math.round(value));
            return String(Math.round(value * 1000) / 1000);
        }

        // =========================================
        // ユーティリティ / Utilities
        // =========================================

        /**
         * 親方向にたどって表とその経路情報を求める
         * @param {object} start 起点となるオブジェクト
         * @returns {object|null} 表と経路の情報。見つからない場合は null
         */
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

        /**
         * 行インデックスを重複なく集めるコレクタを作る
         * @returns {object} 行インデックスを追加・取得するオブジェクト
         */
        function createRowIndexCollector() {
            var rowMap = {};
            var hasSpecificRows = false;

            /**
             * 行インデックスを 1 つ追加する
             * @param {number} index 行インデックス
             * @returns {void}
             */
            function addRowIndex(index) {
                if (index === undefined || index === null) return;
                rowMap[index] = true;
                hasSpecificRows = true;
            }

            /**
             * 行オブジェクトの配列からインデックスを追加する
             * @param {Array<Row>} rows 行の配列
             * @returns {void}
             */
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

            /**
             * セルの配列から行インデックスを追加する
             * @param {Array<Cell>} cells セルの配列
             * @returns {void}
             */
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

            /**
             * 集めた行インデックスを昇順の配列で返す
             * @returns {Array<number>} 昇順の行インデックス
             */
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

        /**
         * 選択項目から行インデックスを収集する
         * @param {object} walkResult walkUpToTable の結果
         * @param {object} collector 行インデックスのコレクタ
         * @returns {void}
         */
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

        /**
         * 選択から対象の表と選択行インデックスを特定する
         * @param {Array} selection 選択オブジェクトの配列
         * @returns {object|null} 表と選択行の情報。特定できない場合は null
         */
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

        /**
         * ドキュメント内のすべての表を集める
         * @returns {Array<Table>} 表の配列
         */
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

        /**
         * 基準の表が属するストーリー内の表を集める
         * @param {Table} baseTable 基準となる表
         * @returns {Array<Table>} 表の配列
         */
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

        /**
         * 指定した範囲に含まれる表を求める
         * @param {string} scope "document" / "story" / "selection"
         * @param {Table} baseTable 基準となる表
         * @returns {Array<Table>} 対象の表の配列
         */
        function resolveScopeTables(scope, baseTable) {
            if (scope === "document") return collectDocumentTables();
            if (scope === "story") return collectStoryTables(baseTable);
            return [baseTable];
        }

        /**
         * 行の高さを一時保存して復元できるストアを作る
         * @returns {object} 保存と復元を行うオブジェクト
         */
        function createRowSnapshotStore() {
            var entries = [];
            var seen = [];

            /**
             * まだ保存していない表の行高を控える
             * @param {Table} table 対象の表
             * @returns {void}
             */
            function ensure(table) {
                for (var i = 0; i < seen.length; i++) {
                    if (seen[i] === table) return;
                }
                var rows = getAllRows(table);
                entries.push({ rows: rows, heights: getOriginalHeights(rows) });
                seen.push(table);
            }

            /**
             * 控えておいたすべての行高を元に戻す
             * @returns {void}
             */
            function restoreAll() {
                for (var i = 0; i < entries.length; i++) {
                    restoreRowHeights(entries[i].rows, entries[i].heights);
                }
            }

            return { ensure: ensure, restoreAll: restoreAll };
        }



        /**
         * 表のすべての行を取得する
         * @param {Table} table 対象の表
         * @returns {Array<Row>} 行の配列
         */
        function getAllRows(table) {
            var all = [];
            for (var i = 0; i < table.rows.length; i++) all.push(table.rows[i]);
            return all;
        }

        /**
         * 指定したインデックスの行を取得する
         * @param {Table} table 対象の表
         * @param {Array<number>} rowIndices 行インデックスの配列
         * @returns {Array<Row>} 行の配列
         */
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

        /**
         * 見出し行を除いた本文行を取得する
         * @param {Table} table 対象の表
         * @returns {Array<Row>} 行の配列
         */
        function getBodyRows(table) {
            var headerCount = 0;
            try { headerCount = table.headerRowCount || 0; } catch (e) { }
            var rows = [];
            for (var i = headerCount; i < table.rows.length; i++) {
                rows.push(table.rows[i]);
            }
            return rows;
        }

        /**
         * すべての行が同じ高さならその値を返す
         * @param {Array<Row>} rows 対象の行
         * @returns {number|null} 共通の高さ。異なる場合は null
         */
        function getCommonRowHeight(rows) {
            if (!rows || rows.length === 0) return null;
            var first = rows[0].height;
            for (var i = 1; i < rows.length; i++) {
                if (Math.abs(rows[i].height - first) > 0.01) return null;
            }
            return first;
        }

        /**
         * 行の高さの平均を求める
         * @param {Array<Row>} rows 対象の行
         * @returns {number} 平均の高さ
         */
        function getAverageRowHeight(rows) {
            if (!rows || rows.length === 0) return null;
            var total = 0;
            for (var i = 0; i < rows.length; i++) {
                total += rows[i].height;
            }
            return total / rows.length;
        }

        /**
         * ダイアログの初期値に使う行の高さを求める
         * @param {Table} table 対象の表
         * @param {Array<number>} selectedRowIndices 選択行のインデックス
         * @returns {number} 初期値（pt）
         */
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

        /**
         * 進捗表示用のパレットを作る
         * @param {string} title タイトル
         * @param {number} maxValue 進捗の最大値
         * @returns {object} 更新と終了を行うオブジェクト
         */
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

        /**
         * 現在プレビュー表示になっているかを判定する
         * @returns {boolean} プレビュー表示なら true
         */
        function isInPreviewScreenMode() {
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

        /**
         * 画面モードに応じたトグルボタンのラベルを返す
         * @returns {string} ボタンに表示する文字列
         */
        function getScreenModeToggleButtonLabel() {
            return isInPreviewScreenMode() ? getLabel('button.exitPreview') : getLabel('button.enterPreview');
        }

        /**
         * トグルボタンのラベルを現在の画面モードに合わせて更新する
         * @param {Button} button 対象のボタン
         * @returns {void}
         */
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

        /**
         * 指定した行に高さを適用する
         * @param {Array<Row>} rows 対象の行
         * @param {number} height 行の高さ（pt）
         * @returns {void}
         */
        function applyRowHeight(rows, height) {
            for (var i = 0; i < rows.length; i++) {
                rows[i].height = height;
            }
        }

        /**
         * 行の現在の高さを控える
         * @param {Array<Row>} rows 対象の行
         * @returns {Array<number>} 各行の高さ
         */
        function getOriginalHeights(rows) {
            var heights = [];
            for (var i = 0; i < rows.length; i++) {
                heights.push(rows[i].height);
            }
            return heights;
        }

        /**
         * 控えておいた行の高さを戻す
         * @param {Array<Row>} rows 対象の行
         * @param {Array<number>} heights 戻す高さの配列
         * @returns {void}
         */
        function restoreRowHeights(rows, heights) {
            for (var i = 0; i < rows.length; i++) {
                rows[i].height = heights[i];
            }
        }

        /**
         * 対象モードに応じて処理する行を求める
         * @param {Table} table 対象の表
         * @param {Array<number>|null} selectedRowIndices 選択行のインデックス
         * @param {string} targetMode "whole" / "body" / "selected"
         * @returns {Array<Row>} 対象の行
         */
        function resolveTargetRows(table, selectedRowIndices, targetMode) {
            var hasSelection = !!(selectedRowIndices && selectedRowIndices.length > 0);
            if (targetMode === "selected" && hasSelection) return getRowsByIndices(table, selectedRowIndices);
            if (targetMode === "body") return getBodyRows(table);
            return getAllRows(table);
        }

        /**
         * 実行前の選択状態を控えておく
         * @returns {Array} 復元用の選択情報
         */
        function getSelectionSnapshot() {
            var snapshot = [];
            try {
                var sel = app.selection;
                for (var i = 0; i < sel.length; i++) snapshot.push(sel[i]);
            } catch (e) { }
            return snapshot;
        }

        /**
         * 控えておいた選択状態を復元する
         * @param {Array} snapshot 復元用の選択情報
         * @returns {void}
         */
        function restoreSelectionSnapshot(snapshot) {
            try {
                if (snapshot && snapshot.length > 0) {
                    app.select(snapshot);
                } else {
                    app.select(NothingEnum.NOTHING);
                }
            } catch (e) { }
        }

        /**
         * 表を含む親テキストフレームを内容に合わせる
         * @param {Table} table 対象の表
         * @returns {void}
         */
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

        /**
         * 行の高さを設定するダイアログを表示する
         * @param {Table} table 対象の表
         * @param {Array<number>} selectedRowIndices 選択行のインデックス
         * @param {number} defaultValuePt 初期値（pt）
         * @returns {object|null} 設定内容。キャンセル時は null
         */
        function showRowHeightDialog(table, selectedRowIndices, defaultValuePt) {
            /* 触れた表の元の高さを遅延記憶（範囲切り替えに追従） / Lazily remember original heights of touched tables (follows scope changes) */
            var snapshot = createRowSnapshotStore();
            var hasSelection = !!(selectedRowIndices && selectedRowIndices.length > 0);
            var documentUnit = getDocumentVerticalUnit();
            var defaultValueDisplay = convertFromPointsByUnit(defaultValuePt, documentUnit);

            var dlg = new Window("dialog", getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
            setupWindow(dlg, 10);

            var contentGroup = dlg.add("group");
            setupRow(contentGroup, "fill", COLUMN_SPACING);
            contentGroup.alignChildren = ["fill", "top"];

            var leftColumn = contentGroup.add("group");
            leftColumn.orientation = "column";
            leftColumn.alignChildren = "fill";
            leftColumn.alignment = ["fill", "fill"];

            var rightColumn = contentGroup.add("group");
            rightColumn.orientation = "column";
            rightColumn.alignChildren = ["fill", "top"];
            rightColumn.alignment = ["right", "fill"];

            /* 範囲選択パネル / Scope selection panel */
            var scopeGroup = leftColumn.add("panel", undefined, getLabel('scope.panel'));

            setupPanel(scopeGroup, 6);

            scopeGroup.alignChildren = "left";
            var scopeDocument = scopeGroup.add("radiobutton", undefined, getLabel('scope.document'));
            var scopeStory = scopeGroup.add("radiobutton", undefined, getLabel('scope.story'));
            var scopeSelection = scopeGroup.add("radiobutton", undefined, getLabel('scope.selection'));
            scopeDocument.helpTip = getLabel('tooltip.scopeDocument');
            scopeStory.helpTip = getLabel('tooltip.scopeStory');
            scopeSelection.helpTip = getLabel('tooltip.scopeSelection');
            scopeSelection.value = true;

            scopeDocument.onClick = function () { syncTargetEnabled(); updatePreview(); };
            scopeStory.onClick = function () { syncTargetEnabled(); updatePreview(); };
            scopeSelection.onClick = function () { syncTargetEnabled(); updatePreview(); };

            /* 対象選択パネル / Target selection panel */
            var targetGroup = leftColumn.add("panel", undefined, getLabel('target.panel'));

            setupPanel(targetGroup, 6);

            targetGroup.alignChildren = "left";
            var targetWhole = targetGroup.add("radiobutton", undefined, getLabel('target.whole'));
            var targetBody = targetGroup.add("radiobutton", undefined, getLabel('target.wholeNoHead'));
            var targetSel = targetGroup.add("radiobutton", undefined, getLabel('target.selected'));
            targetSel.enabled = hasSelection;
            if (hasSelection) {
                targetSel.value = true;
            } else {
                targetWhole.value = true;
            }

            /* モード選択パネル / Mode selection panel */
            var modeGroup = leftColumn.add("panel", undefined, getLabel('mode.panel'));

            setupPanel(modeGroup, 6);

            modeGroup.alignChildren = "left";

            var modeMinimum = modeGroup.add("radiobutton", undefined, getLabel('mode.minimum'));
            var modeSpecified = modeGroup.add("radiobutton", undefined, getLabel('mode.specified'));
            modeMinimum.helpTip = getLabel('tooltip.modeMinimum');
            modeSpecified.helpTip = getLabel('tooltip.modeSpecified');
            modeMinimum.value = true;

            var manualRow = modeGroup.add("group");
            setupRow(manualRow, "left", 6);
            manualRow.alignChildren = ["left", "center"];
            var input = manualRow.add("edittext", undefined, formatDisplayValue(defaultValueDisplay));
            input.characters = ROW_HEIGHT_INPUT_CHARACTERS;
            input.helpTip = getLabel('tooltip.rowHeightInput');
            manualRow.add("statictext", undefined, getUnitLabel());

            /**
             * 指定値モードのときだけ入力欄を有効にする
             * @returns {void}
             */
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
            var okBtn = actionGroup.add("button", undefined, getLabel('button.ok'), { name: "ok" });
            var cancelBtn = actionGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });

            /* OK/Cancel とプレビューボタンの間を伸ばすスペーサー / Spacer that pushes the preview button to the bottom */
            var rightSpacer = rightColumn.add("group");
            rightSpacer.alignment = ["fill", "fill"];

            /* 右カラム下部の画面モード切り替え / Screen mode toggle at the bottom of the right column */
            var screenModeGroup = rightColumn.add("group");
            setupRow(screenModeGroup, "fill", 8);
            screenModeGroup.alignChildren = ["fill", "center"];
            var screenModeToggleBtn = screenModeGroup.add("button", undefined, getScreenModeToggleButtonLabel());
            screenModeToggleBtn.alignment = ["fill", "center"];
            screenModeToggleBtn.helpTip = getLabel('tooltip.screenModeToggle');

            screenModeToggleBtn.onClick = function () {
                toggleScreenPreviewMode();
                updateScreenModeToggleButtonLabel(screenModeToggleBtn);
            };

            /* 左カラム下部のオプションパネル / Options panel at the bottom of the left column */
            var optionsGroup = leftColumn.add("panel", undefined, getLabel('options.panel'));

            setupPanel(optionsGroup, 6);

            optionsGroup.alignChildren = ["left", "center"];
            var fitParentFrameBtn = optionsGroup.add("button", undefined, getLabel('button.fitParentFrame'));
            fitParentFrameBtn.helpTip = getLabel('tooltip.fitParentFrame');

            fitParentFrameBtn.onClick = function () {
                snapshot.restoreAll();
                updatePreview();
                fitParentFrameToContent(table);
            };


            /**
             * 入力欄の値をポイント値に変換する
             * @returns {number|null} ポイント値。無効な場合は null
             */
            function parseValue() {
                var v = parseFloat(input.text);
                if (isNaN(v) || v <= 0) return null;
                return convertToPointsByUnit(v, documentUnit);
            }

            /**
             * 現在のモードに応じた行の高さ（pt）を求める
             * @returns {number|null} 行の高さ。無効な場合は null
             */
            function getRowHeightValue() {
                if (modeMinimum.value) return MIN_ROW_HEIGHT_PT;
                return parseValue();
            }

            /**
             * 選択中の範囲を取得する
             * @returns {string} "document" / "story" / "selection"
             */
            function getCurrentScope() {
                if (scopeDocument.value) return "document";
                if (scopeStory.value) return "story";
                return "selection";
            }

            /**
             * 範囲に応じて「選択した行のみ」の有効／無効を切り替える
             * @returns {void}
             */
            function syncTargetEnabled() {
                var allowSelected = (getCurrentScope() === "selection") && hasSelection;
                targetSel.enabled = allowSelected;
                if (!allowSelected && targetSel.value) {
                    targetWhole.value = true;
                }
            }

            /**
             * 選択中の対象モードを取得する
             * @returns {string} "whole" / "body" / "selected"
             */
            function getCurrentTargetMode() {
                if (targetSel.value && targetSel.enabled && hasSelection) return "selected";
                if (targetBody.value) return "body";
                return "whole";
            }

            /**
             * 現在の設定で行の高さのプレビューを描き直す
             * @returns {void}
             */
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
                alert(getLabel('error.invalidNumber'));
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
            alert(getLabel('error.noTable'));
        } else if (target.error === "multipleTables") {
            alert(getLabel('error.multipleTables'));
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

                var progress = createProgressBar(getLabel('progress.title') + ' ' + SCRIPT_VERSION, totalRows);

                /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
                app.doScript(function () {
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
                }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel('undo.applyRowHeight'));
            }
        }

    })();
