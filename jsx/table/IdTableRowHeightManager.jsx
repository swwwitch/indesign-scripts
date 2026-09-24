#target indesign

/*

### 概要

選択した表の行の高さを、範囲（選択範囲／ストーリー／ドキュメント）と対象行を指定しながらプレビュー付きで設定します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableRowHeightManager.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n9f95f8e98db6

### Overview

Sets the row heights of the selected table with a live preview, choosing the scope (selection, story or document) and which rows to target.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableRowHeightManager.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTableRowHeightManager";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-25";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableRowHeightManager.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableRowHeightManager.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n9f95f8e98db6"; /* 紹介記事 / article URL */

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

/* ボタンエリア / Button area */
var BUTTON_ROW_TOP_MARGIN = 10;          /* ボタン列の上余白 / top margin above the button row */
var BUTTON_SPACING        = 10;          /* ボタンの間隔 / gap between buttons */

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

        /* 初期値を行から求められないときの行の高さ（pt） / Fallback row height (pt) when none can be read from the rows */
        var FALLBACK_ROW_HEIGHT_PT = 15;

        // =========================================
        // レイアウト設定 / Layout settings
        // =========================================

        /* 行高入力欄の文字数 / Character width of the row-height field */
        var ROW_HEIGHT_INPUT_CHARACTERS = 5;

        /* 進捗バーの幅 / Width of the progress bar */
        var PROGRESS_BAR_WIDTH = 320;

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
        var uiLang = getCurrentLang();

        var LABELS = {
            dialog: {
                title: { ja: "行の高さを設定", en: "Set Row Height" }
            },
            panel: {
                scope: { ja: "範囲", en: "Scope" },
                target: { ja: "対象", en: "Target" },
                rowHeight: { ja: "行の高さ", en: "Row Height" },
                options: { ja: "オプション", en: "Options" }
            },
            radio: {
                scopeDocument: { ja: "ドキュメント", en: "Document" },
                scopeStory: { ja: "ストーリー", en: "Story" },
                scopeSelection: { ja: "選択範囲", en: "Selection" },
                targetWhole: { ja: "表全体", en: "Whole Table" },
                targetBody: { ja: "表全体（ヘッダー行を除く）", en: "Whole Table (Except Header Rows)" },
                targetSelected: { ja: "選択した行のみ", en: "Selected Rows Only" },
                heightMinimum: { ja: "最小限度", en: "At Least" },
                heightSpecified: { ja: "指定値を使用", en: "Exactly" }
            },
            checkbox: {
                fitFrameToContent: { ja: "フレームを内容に合わせる", en: "Fit Frame to Content" }
            },
            button: {
                enterPreviewMode: { ja: "プレビューモード", en: "Preview Mode" },
                exitPreviewMode: { ja: "標準モード", en: "Normal Mode" },
                ok: { ja: "OK", en: "OK" },
                cancel: { ja: "キャンセル", en: "Cancel" }
            },
            tooltip: {
                scopeDocument: {
                    ja: "ドキュメント内のすべての表を対象にします。",
                    en: "Targets all tables in the document."
                },
                scopeStory: {
                    ja: "選択した表と同じストーリーにあるすべての表を対象にします。",
                    en: "Targets all tables in the same story as the selected table."
                },
                scopeSelection: {
                    ja: "選択した表だけを対象にします。",
                    en: "Targets only the selected table."
                },
                targetBody: {
                    ja: "ヘッダー行を除いた行を対象にします。",
                    en: "Targets all rows except header rows."
                },
                targetSelected: {
                    ja: "選択した行だけを対象にします。範囲が「選択範囲」で、行を選択しているときに使えます。",
                    en: "Targets only the selected rows. Available when the scope is Selection and rows are selected."
                },
                heightMinimum: {
                    ja: "各行を最小の高さ（約1pt）に設定し、内容に合わせて自動的に伸びるようにします。",
                    en: "Sets each row to the minimum height (about 1 pt) so it grows to fit its content."
                },
                heightSpecified: {
                    ja: "入力した値を行の高さとして設定します。",
                    en: "Uses the entered value as the row height."
                },
                heightInput: {
                    ja: "↑↓キーで0.1ずつ増減、Shift＋↑↓で整数値に揃えます。単位はドキュメントの縦方向の単位です。",
                    en: "Up/Down changes the value by 0.1; Shift+Up/Down snaps to whole numbers. Unit follows the document's vertical units."
                },
                screenMode: {
                    ja: "画面モードを標準とプレビューで切り替えます。ガイドや枠線を隠して仕上がりを確認できます。",
                    en: "Switches the screen mode between Normal and Preview to check the result without guides and frame edges."
                },
                fitFrameToContent: {
                    ja: "行の高さを変えたあと、表を含むテキストフレームの高さを内容に合わせます。",
                    en: "After changing the row heights, fits the height of the text frame containing the table to its content."
                }
            },
            alert: {
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
            var labelNode = LABELS;
            for (var i = 0; i < keyParts.length; i++) {
                if (labelNode && labelNode.hasOwnProperty(keyParts[i])) {
                    labelNode = labelNode[keyParts[i]];
                } else {
                    return labelKey;
                }
            }
            var labelString = (labelNode && labelNode[uiLang]) ? labelNode[uiLang] : labelKey;
            return labelString.replace(/\{slash\}/g, "/");
        }

        // =========================================
        // 単位 / Units
        // =========================================

        /**
         * 単位の表示ラベルと 1 単位あたりのポイント数を返す
         * @param {MeasurementUnits} unit 対象の単位
         * @returns {object} { label: string, pointsPerUnit: number }
         */
        function getUnitInfo(unit) {
            switch (unit) {
                case MeasurementUnits.MILLIMETERS: return { label: "mm", pointsPerUnit: 72 / 25.4 };
                case MeasurementUnits.CENTIMETERS: return { label: "cm", pointsPerUnit: 72 / 2.54 };
                case MeasurementUnits.INCHES:
                case MeasurementUnits.INCHES_DECIMAL: return { label: "inch", pointsPerUnit: 72 };
                case MeasurementUnits.POINTS:
                case MeasurementUnits.AMERICAN_POINTS: return { label: "pt", pointsPerUnit: 1 };
                case MeasurementUnits.PICAS: return { label: "pica", pointsPerUnit: 12 };
                case MeasurementUnits.PIXELS: return { label: "px", pointsPerUnit: 1 };
                case MeasurementUnits.AGATES: return { label: "ag", pointsPerUnit: 5.5 };
                case MeasurementUnits.Q: return { label: "Q", pointsPerUnit: 72 / 25.4 * 0.25 };
                case MeasurementUnits.HA: return { label: "H", pointsPerUnit: 72 / 25.4 * 0.25 };
                case MeasurementUnits.CICEROS: return { label: "cicero", pointsPerUnit: 1 };
                case MeasurementUnits.BAI: return { label: "倍", pointsPerUnit: 1 };
                case MeasurementUnits.U: return { label: "U", pointsPerUnit: 1 };
                default: return { label: "", pointsPerUnit: 1 };
            }
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
         * コレクションや選択を通常の配列に写す
         * @param {object} collection length と添字アクセスを持つオブジェクト
         * @param {number} [startIndex] 写し始める位置。省略時は 0
         * @returns {Array} 要素の配列
         */
        function collectionToArray(collection, startIndex) {
            var items = [];
            for (var i = startIndex || 0; i < collection.length; i++) items.push(collection[i]);
            return items;
        }

        /**
         * 親方向にたどって表とその経路情報を求める
         * @param {object} startItem 起点となるオブジェクト
         * @returns {object} { table, cell, row, sourceItem }。表が無ければ table は null
         */
        function walkUpToTable(startItem) {
            var currentNode = startItem;
            var foundCell = null;
            var foundRow = null;
            var foundTable = null;
            while (currentNode) {
                /* Application まで遡ると parent が無く例外になる / Reaching Application throws on .parent */
                try {
                    if (currentNode instanceof Cell) { foundCell = currentNode; foundTable = currentNode.parent; break; }
                    if (currentNode instanceof Row) { foundRow = currentNode; foundTable = currentNode.parent; break; }
                    if (currentNode instanceof Table) { foundTable = currentNode; break; }
                    if (currentNode instanceof TextFrame && currentNode.tables.length > 0) { foundTable = currentNode.tables[0]; break; }
                    currentNode = currentNode.parent;
                } catch (e) { break; }
            }
            return { table: foundTable, cell: foundCell, row: foundRow, sourceItem: startItem };
        }

        /**
         * 行インデックスを重複なく集めるコレクタを作る
         * @returns {object} 行インデックスを追加・取得するオブジェクト
         */
        function createRowIndexCollector() {
            var rowIndexMap = {};
            var hasSpecificRows = false;

            /**
             * 行インデックスを 1 つ追加する
             * @param {number} rowIndex 行インデックス
             * @returns {void}
             */
            function addRowIndex(rowIndex) {
                if (rowIndex === undefined || rowIndex === null) return;
                rowIndexMap[rowIndex] = true;
                hasSpecificRows = true;
            }

            /**
             * 要素ごとに行を求めて、そのインデックスを追加する
             * @param {Array} items 行またはセルの配列
             * @param {Function} getRow 要素から行を返す関数
             * @returns {boolean} 1 つでも追加できたら true
             */
            function addRowsFrom(items, getRow) {
                if (!items || items.length === 0) return false;
                var added = false;
                for (var i = 0; i < items.length; i++) {
                    /* 無効になった参照は読めないので飛ばす / Skip references that became invalid */
                    try {
                        addRowIndex(getRow(items[i]).index);
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
                for (var k in rowIndexMap) {
                    if (rowIndexMap.hasOwnProperty(k)) rowIndices.push(parseInt(k, 10));
                }
                rowIndices.sort(function (a, b) { return a - b; });
                return rowIndices;
            }

            return {
                addRowIndex: addRowIndex,
                addRowsFromRows: function (rows) { return addRowsFrom(rows, function (row) { return row; }); },
                addRowsFromCells: function (cells) { return addRowsFrom(cells, function (cell) { return cell.parentRow; }); },
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
            var sourceItem = walkResult.sourceItem;

            /* 選択項目が持つ行・セルのコレクションから行インデックスを集める / Collect row indices from any row/cell collections the item exposes */
            var rangeSources = [
                { propName: "parentRows", addRows: collector.addRowsFromRows },
                { propName: "rows", addRows: collector.addRowsFromRows },
                { propName: "parentCells", addRows: collector.addRowsFromCells },
                { propName: "cells", addRows: collector.addRowsFromCells }
            ];
            var addedFromRange = false;
            for (var i = 0; i < rangeSources.length; i++) {
                /* 選択の型によってはプロパティ自体が無く例外になる / Some selection types lack the property and throw */
                try {
                    var collection = sourceItem && sourceItem[rangeSources[i].propName];
                    if (collection && collection.length > 0) {
                        addedFromRange = rangeSources[i].addRows(collection.everyItem().getElements()) || addedFromRange;
                    }
                } catch (e) { }
            }
            if (addedFromRange) return;

            if (walkResult.row) { collector.addRowIndex(walkResult.row.index); return; }
            if (walkResult.cell) { collector.addRowIndex(walkResult.cell.parentRow.index); return; }

            /* parentRow を持たない選択もある / Not every selection has parentRow */
            try {
                if (sourceItem && sourceItem.parentRow) collector.addRowIndex(sourceItem.parentRow.index);
            } catch (e) { }
        }

        /**
         * 選択から対象の表と選択行インデックスを特定する
         * @param {Array} selection 選択オブジェクトの配列
         * @returns {object|null} { table, rowIndices } または { error }。特定できない場合は null
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
            var stories = app.activeDocument.stories;
            for (var i = 0; i < stories.length; i++) {
                tables = tables.concat(collectionToArray(stories[i].tables));
            }
            return tables;
        }

        /**
         * 指定した範囲に含まれる表を求める
         * @param {string} scope "document" / "story" / "selection"
         * @param {Table} baseTable 基準となる表
         * @returns {Array<Table>} 対象の表の配列
         */
        function resolveScopeTables(scope, baseTable) {
            if (scope === "document") return collectDocumentTables();
            if (scope === "story") return collectionToArray(baseTable.parentStory.tables);
            return [baseTable];
        }

        // =========================================
        // 行の取得 / Row lookup
        // =========================================

        /**
         * 指定したインデックスの行を取得する
         * @param {Table} table 対象の表
         * @param {Array<number>} rowIndices 行インデックスの配列
         * @returns {Array<Row>} 行の配列
         */
        function getRowsByIndices(table, rowIndices) {
            var rows = [];
            for (var i = 0; i < rowIndices.length; i++) {
                var rowIndex = rowIndices[i];
                if (rowIndex >= 0 && rowIndex < table.rows.length) rows.push(table.rows[rowIndex]);
            }
            return rows;
        }

        /**
         * 対象モードに応じて処理する行を求める
         * @param {Table} table 対象の表
         * @param {Array<number>|null} selectedRowIndices 選択行のインデックス
         * @param {string} targetMode "whole" / "body" / "selected"
         * @returns {Array<Row>} 対象の行
         */
        function resolveTargetRows(table, selectedRowIndices, targetMode) {
            var hasSelectedRows = !!(selectedRowIndices && selectedRowIndices.length > 0);
            if (targetMode === "selected" && hasSelectedRows) return getRowsByIndices(table, selectedRowIndices);
            if (targetMode === "body") return collectionToArray(table.rows, table.headerRowCount);
            return collectionToArray(table.rows);
        }

        /**
         * 範囲と対象モードから、表ごとの対象行をまとめて求める
         * 選択行の指定は基準の表にだけ効き、ほかの表は表全体として扱う
         * @param {Array<Table>} tables 対象の表
         * @param {Table} baseTable 基準となる表
         * @param {Array<number>|null} selectedRowIndices 基準の表の選択行
         * @param {string} targetMode "whole" / "body" / "selected"
         * @returns {Array<Array<Row>>} 表ごとの行の配列
         */
        function collectTargetRowSets(tables, baseTable, selectedRowIndices, targetMode) {
            var rowSets = [];
            for (var i = 0; i < tables.length; i++) {
                var rowIndicesForTable = (tables[i] === baseTable) ? selectedRowIndices : null;
                rowSets.push(resolveTargetRows(tables[i], rowIndicesForTable, targetMode));
            }
            return rowSets;
        }

        /**
         * ダイアログの初期値に使う行の高さを求める
         * 対象行がそろっていればその値、ばらばらなら本文行の平均
         * @param {Table} table 対象の表
         * @param {Array<number>|null} selectedRowIndices 選択行のインデックス
         * @returns {number} 初期値（pt）
         */
        function getInitialRowHeight(table, selectedRowIndices) {
            var targetRows = resolveTargetRows(table, selectedRowIndices, "selected");
            if (targetRows.length > 0) {
                var firstHeight = targetRows[0].height;
                var isUniform = true;
                for (var i = 1; i < targetRows.length; i++) {
                    if (Math.abs(targetRows[i].height - firstHeight) > 0.01) { isUniform = false; break; }
                }
                if (isUniform) return firstHeight;
            }

            var bodyRows = resolveTargetRows(table, null, "body");
            if (bodyRows.length === 0) return FALLBACK_ROW_HEIGHT_PT;
            var totalHeight = 0;
            for (var j = 0; j < bodyRows.length; j++) totalHeight += bodyRows[j].height;
            return totalHeight / bodyRows.length;
        }

        // =========================================
        // 行の高さ操作 / Row height operations
        // =========================================

        /**
         * 行の高さを一時保存して復元できるストアを作る
         * @returns {object} 保存と復元を行うオブジェクト
         */
        function createRowHeightSnapshotStore() {
            var entries = [];
            var frameEntries = [];

            /**
             * まだ控えていない表の行高を控える
             * @param {Table} table 対象の表
             * @returns {void}
             */
            function rememberTable(table) {
                for (var i = 0; i < entries.length; i++) {
                    if (entries[i].table === table) return;
                }
                var rows = collectionToArray(table.rows);
                var heights = [];
                for (var j = 0; j < rows.length; j++) heights.push(rows[j].height);
                entries.push({ table: table, rows: rows, heights: heights });
            }

            /**
             * まだ控えていないフレームの寸法を控える
             * @param {TextFrame} frame 対象のフレーム
             * @returns {void}
             */
            function rememberFrame(frame) {
                for (var i = 0; i < frameEntries.length; i++) {
                    if (frameEntries[i].frame === frame) return;
                }
                frameEntries.push({ frame: frame, bounds: frame.geometricBounds });
            }

            /**
             * 控えておいたすべての行高とフレームの寸法を元に戻す
             * @returns {void}
             */
            function restoreAll() {
                for (var i = 0; i < entries.length; i++) {
                    for (var j = 0; j < entries[i].rows.length; j++) {
                        entries[i].rows[j].height = entries[i].heights[j];
                    }
                }
                for (var k = 0; k < frameEntries.length; k++) {
                    frameEntries[k].frame.geometricBounds = frameEntries[k].bounds;
                }
            }

            return { rememberTable: rememberTable, rememberFrame: rememberFrame, restoreAll: restoreAll };
        }

        /**
         * 表を含むテキストフレームを求める
         * @param {Table} table 対象の表
         * @returns {TextFrame|null} テキストフレーム。パス上のテキストなどでは null
         */
        function getParentTextFrame(table) {
            var frames = table.storyOffset.parentTextFrames;
            var parentFrame = (frames.length > 0) ? frames[0] : table.parentStory.textContainers[0];
            return (parentFrame instanceof TextFrame) ? parentFrame : null;
        }

        /**
         * テキストフレームを内容に合わせる
         * @param {TextFrame} frame 対象のフレーム
         * @returns {void}
         */
        function fitFrameToContent(frame) {
            var story = frame.parentStory;
            story.recompose();
            /* ロックされたフレームなどは fit が例外になる / fit throws on locked frames and the like */
            try {
                frame.fit(FitOptions.FRAME_TO_CONTENT);
            } catch (e) { return; }
            story.recompose();
        }

        // =========================================
        // 選択の控えと復元 / Selection snapshot
        // =========================================

        /**
         * 控えておいた選択状態を復元する
         * @param {Array} selectionItems 復元する選択項目
         * @returns {void}
         */
        function restoreSelection(selectionItems) {
            /* 控えた項目が無効になっていると select が例外になる / select throws if a remembered item became invalid */
            try {
                app.select(selectionItems.length > 0 ? selectionItems : NothingEnum.NOTHING);
            } catch (e) { }
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
            var progressWin = new Window("palette", title);
            setupWindow(progressWin);

            var progressLabel = progressWin.add("statictext", undefined, "");
            progressLabel.preferredSize.width = PROGRESS_BAR_WIDTH;

            var progressControl = progressWin.add("progressbar", undefined, 0, maxValue);
            progressControl.preferredSize = [PROGRESS_BAR_WIDTH, 12];

            progressWin.show();

            return {
                update: function (value, text) {
                    progressControl.value = value;
                    if (text !== undefined) progressLabel.text = text;
                    progressWin.update();
                },
                close: function () {
                    progressWin.close();
                }
            };
        }

        // =========================================
        // 画面モード切り替え / Screen mode toggle
        // =========================================

        /**
         * 現在プレビュー表示になっているかを判定する
         * @returns {boolean} プレビュー表示なら true
         */
        function isInPreviewScreenMode() {
            /* ストーリーエディターのウィンドウには screenMode が無い / Story editor windows have no screenMode */
            try {
                return app.activeWindow.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE;
            } catch (e) {
                return false;
            }
        }

        /**
         * 標準表示とプレビュー表示を切り替える
         * @returns {void}
         */
        function toggleScreenPreviewMode() {
            var nextMode = isInPreviewScreenMode() ? ScreenModeOptions.PREVIEW_OFF : ScreenModeOptions.PREVIEW_TO_PAGE;
            /* ストーリーエディターのウィンドウには screenMode が無い / Story editor windows have no screenMode */
            try {
                app.activeWindow.screenMode = nextMode;
            } catch (e) { }
        }

        /**
         * 画面モードに応じた切り替えボタンのラベルを返す
         * @returns {string} ボタンに表示する文字列
         */
        function getScreenModeButtonLabel() {
            return isInPreviewScreenMode() ? getLabel('button.exitPreviewMode') : getLabel('button.enterPreviewMode');
        }

        // =========================================
        // キーボード操作 / Keyboard interaction
        // =========================================

        /**
         * ↑↓ で値を増減する。通常は±0.1、Shift では整数値へスナップ
         * @param {EditText} editText 対象の入力欄
         * @param {Function} [onUpdate] 値を変えた後に呼ぶ関数
         * @returns {void}
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

                var keyboardState = ScriptUI.environment.keyboardState;
                var isShift = (keyboardState && keyboardState.shiftKey) ? true : event.shiftKey;
                if (isShift) {
                    value = isArrowUp ? Math.ceil(value + 0.0001) : Math.floor(value - 0.0001);
                } else {
                    value = Math.round((value + (isArrowUp ? 0.1 : -0.1)) * 10) / 10;
                }

                if (value < 0) value = 0;

                editText.text = value;
                if (typeof onUpdate === "function") onUpdate();
            });
        }

        // =========================================
        // ダイアログ部品 / Dialog parts
        // =========================================

        /**
         * パネルにラジオボタンを並べる
         * @param {Panel} panel 追加先のパネル
         * @param {Array<object>} radioDefs { labelKey, tooltipKey } の配列。tooltipKey は省略可
         * @returns {Array<RadioButton>} 追加したラジオボタン
         */
        function addRadioButtons(panel, radioDefs) {
            var radios = [];
            for (var i = 0; i < radioDefs.length; i++) {
                var radio = panel.add("radiobutton", undefined, getLabel(radioDefs[i].labelKey));
                if (radioDefs[i].tooltipKey) radio.helpTip = getLabel(radioDefs[i].tooltipKey);
                radios.push(radio);
            }
            return radios;
        }

        /**
         * ラジオボタンだけを縦に並べるパネルを作る
         * @param {Group} parentGroup 追加先
         * @param {string} titleKey パネル見出しのラベルキー
         * @returns {Panel} 作ったパネル
         */
        function addRadioPanel(parentGroup, titleKey) {
            var radioPanel = parentGroup.add("panel", undefined, getLabel(titleKey));
            setupPanel(radioPanel, 6);
            radioPanel.alignChildren = "left";
            return radioPanel;
        }

        /**
         * 「範囲」パネルを作る
         * @param {Group} parentGroup 追加先
         * @returns {object} { documentRadio, storyRadio, selectionRadio }
         */
        function addScopePanel(parentGroup) {
            var radios = addRadioButtons(addRadioPanel(parentGroup, 'panel.scope'), [
                { labelKey: 'radio.scopeDocument', tooltipKey: 'tooltip.scopeDocument' },
                { labelKey: 'radio.scopeStory', tooltipKey: 'tooltip.scopeStory' },
                { labelKey: 'radio.scopeSelection', tooltipKey: 'tooltip.scopeSelection' }
            ]);
            radios[2].value = true;
            return { documentRadio: radios[0], storyRadio: radios[1], selectionRadio: radios[2] };
        }

        /**
         * 「対象」パネルを作る
         * @param {Group} parentGroup 追加先
         * @param {boolean} hasSelectedRows 行が選択されているか
         * @returns {object} { wholeRadio, bodyRadio, selectedRadio }
         */
        function addTargetPanel(parentGroup, hasSelectedRows) {
            var radios = addRadioButtons(addRadioPanel(parentGroup, 'panel.target'), [
                { labelKey: 'radio.targetWhole' },
                { labelKey: 'radio.targetBody', tooltipKey: 'tooltip.targetBody' },
                { labelKey: 'radio.targetSelected', tooltipKey: 'tooltip.targetSelected' }
            ]);
            radios[2].enabled = hasSelectedRows;
            radios[hasSelectedRows ? 2 : 0].value = true;
            return { wholeRadio: radios[0], bodyRadio: radios[1], selectedRadio: radios[2] };
        }

        /**
         * 「行の高さ」パネルを作る
         * @param {Group} parentGroup 追加先
         * @param {string} initialText 入力欄の初期値
         * @param {string} unitLabel 単位の表示
         * @returns {object} { minimumRadio, specifiedRadio, heightInput }
         */
        function addRowHeightPanel(parentGroup, initialText, unitLabel) {
            var rowHeightPanel = addRadioPanel(parentGroup, 'panel.rowHeight');
            var radios = addRadioButtons(rowHeightPanel, [
                { labelKey: 'radio.heightMinimum', tooltipKey: 'tooltip.heightMinimum' },
                { labelKey: 'radio.heightSpecified', tooltipKey: 'tooltip.heightSpecified' }
            ]);
            radios[0].value = true;

            var heightInputRow = rowHeightPanel.add("group");
            setupRow(heightInputRow, "left", 6);
            heightInputRow.alignChildren = ["left", "center"];
            var heightInput = heightInputRow.add("edittext", undefined, initialText);
            heightInput.characters = ROW_HEIGHT_INPUT_CHARACTERS;
            heightInput.helpTip = getLabel('tooltip.heightInput');
            heightInputRow.add("statictext", undefined, unitLabel);

            return { minimumRadio: radios[0], specifiedRadio: radios[1], heightInput: heightInput };
        }

        /**
         * 「オプション」パネルを作る
         * @param {Group} parentGroup 追加先
         * @param {boolean} canFitFrame フレームを内容に合わせられるか
         * @returns {Checkbox} ［フレームを内容に合わせる］チェックボックス
         */
        function addOptionsPanel(parentGroup, canFitFrame) {
            var optionsPanel = parentGroup.add("panel", undefined, getLabel('panel.options'));
            setupPanel(optionsPanel, 6);
            optionsPanel.alignChildren = ["left", "center"];
            var fitFrameCheckbox = optionsPanel.add("checkbox", undefined, getLabel('checkbox.fitFrameToContent'));
            fitFrameCheckbox.helpTip = getLabel('tooltip.fitFrameToContent');
            fitFrameCheckbox.value = canFitFrame;
            fitFrameCheckbox.enabled = canFitFrame;
            return fitFrameCheckbox;
        }

        /**
         * ボタンエリア（左：画面モード／スペーサー／右：キャンセル・OK）を作る
         * @param {Window} dialogWin 追加先のダイアログ
         * @returns {object} { btnScreenMode, btnCancel, btnOK }
         */
        function addButtonRow(dialogWin) {
            var btnRowGroup = dialogWin.add("group");
            btnRowGroup.orientation = "row";
            btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
            btnRowGroup.alignment = ["fill", "bottom"];

            var btnLeftGroup = btnRowGroup.add("group");
            btnLeftGroup.alignChildren = ["left", "center"];
            var btnScreenMode = btnLeftGroup.add("button", undefined, getScreenModeButtonLabel());
            btnScreenMode.helpTip = getLabel('tooltip.screenMode');

            var spacer = btnRowGroup.add("group");
            spacer.alignment = ["fill", "fill"];
            spacer.minimumSize.width = 0;

            var btnRightGroup = btnRowGroup.add("group");
            btnRightGroup.alignChildren = ["right", "center"];
            btnRightGroup.spacing = BUTTON_SPACING;
            var btnCancel = btnRightGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
            var btnOK = btnRightGroup.add("button", undefined, getLabel('button.ok'), { name: "ok" });

            return { btnScreenMode: btnScreenMode, btnCancel: btnCancel, btnOK: btnOK };
        }

        // =========================================
        // ダイアログ / Dialog
        // =========================================

        /**
         * 行の高さを設定するダイアログを表示する
         * @param {Table} baseTable 対象の表
         * @param {Array<number>|null} selectedRowIndices 選択行のインデックス
         * @param {number} initialHeightPt 初期値（pt）
         * @returns {object|null} { height, targetMode, scope, fitFrame }。キャンセル時は null
         */
        function showRowHeightDialog(baseTable, selectedRowIndices, initialHeightPt) {
            /* 触れた表の元の高さを遅延記憶（範囲切り替えに追従） / Lazily remember original heights of touched tables (follows scope changes) */
            var snapshotStore = createRowHeightSnapshotStore();
            var hasSelectedRows = !!(selectedRowIndices && selectedRowIndices.length > 0);
            var parentFrame = getParentTextFrame(baseTable);
            var unitInfo = getUnitInfo(app.activeDocument.viewPreferences.verticalMeasurementUnits);

            var rowHeightDialog = new Window("dialog", getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
            setupWindow(rowHeightDialog, 10);

            var columnsGroup = rowHeightDialog.add("group");
            setupRow(columnsGroup, "fill", COLUMN_SPACING);
            columnsGroup.alignChildren = ["fill", "top"];

            var leftColumn = columnsGroup.add("group");
            leftColumn.orientation = "column";
            leftColumn.alignChildren = "fill";
            leftColumn.alignment = ["fill", "fill"];

            var rightColumn = columnsGroup.add("group");
            rightColumn.orientation = "column";
            rightColumn.alignChildren = "fill";
            rightColumn.alignment = ["fill", "fill"];

            var scopeControls = addScopePanel(leftColumn);
            var targetControls = addTargetPanel(leftColumn, hasSelectedRows);
            var heightControls = addRowHeightPanel(rightColumn, formatDisplayValue(initialHeightPt / unitInfo.pointsPerUnit), unitInfo.label);
            var fitFrameCheckbox = addOptionsPanel(rightColumn, parentFrame !== null);
            var buttons = addButtonRow(rowHeightDialog);

            /**
             * 選択中の範囲を取得する
             * @returns {string} "document" / "story" / "selection"
             */
            function getCurrentScope() {
                if (scopeControls.documentRadio.value) return "document";
                if (scopeControls.storyRadio.value) return "story";
                return "selection";
            }

            /**
             * 選択中の対象モードを取得する
             * @returns {string} "whole" / "body" / "selected"
             */
            function getCurrentTargetMode() {
                if (targetControls.selectedRadio.value && targetControls.selectedRadio.enabled) return "selected";
                if (targetControls.bodyRadio.value) return "body";
                return "whole";
            }

            /**
             * 現在のモードに応じた行の高さ（pt）を求める
             * @returns {number|null} 行の高さ。入力が正の数でなければ null
             */
            function getRowHeightValue() {
                if (heightControls.minimumRadio.value) return MIN_ROW_HEIGHT_PT;
                var inputValue = parseFloat(heightControls.heightInput.text);
                if (isNaN(inputValue) || inputValue <= 0) return null;
                return inputValue * unitInfo.pointsPerUnit;
            }

            /**
             * 範囲に応じて「選択した行のみ」の有効／無効を切り替える
             * @returns {void}
             */
            function syncTargetEnabled() {
                var allowSelected = (getCurrentScope() === "selection") && hasSelectedRows;
                targetControls.selectedRadio.enabled = allowSelected;
                if (!allowSelected && targetControls.selectedRadio.value) {
                    targetControls.wholeRadio.value = true;
                }
            }

            /**
             * 指定値モードのときだけ入力欄を有効にする
             * @returns {void}
             */
            function syncInputEnabled() {
                heightControls.heightInput.enabled = heightControls.specifiedRadio.value;
            }

            /**
             * 現在の設定で行の高さのプレビューを描き直す
             * @returns {void}
             */
            function updatePreview() {
                /* 毎回すべて元に戻してから対象表に適用 / Always restore then apply to the current target tables */
                snapshotStore.restoreAll();
                var rowHeight = getRowHeightValue();
                if (rowHeight === null) return;
                var tables = resolveScopeTables(getCurrentScope(), baseTable);
                for (var i = 0; i < tables.length; i++) snapshotStore.rememberTable(tables[i]);
                var rowSets = collectTargetRowSets(tables, baseTable, selectedRowIndices, getCurrentTargetMode());
                for (var j = 0; j < rowSets.length; j++) {
                    for (var k = 0; k < rowSets[j].length; k++) rowSets[j][k].height = rowHeight;
                }
                if (fitFrameCheckbox.value) {
                    snapshotStore.rememberFrame(parentFrame);
                    fitFrameToContent(parentFrame);
                }
            }

            scopeControls.documentRadio.onClick =
                scopeControls.storyRadio.onClick =
                scopeControls.selectionRadio.onClick = function () { syncTargetEnabled(); updatePreview(); };

            targetControls.wholeRadio.onClick =
                targetControls.bodyRadio.onClick =
                targetControls.selectedRadio.onClick = updatePreview;

            heightControls.minimumRadio.onClick = function () { syncInputEnabled(); updatePreview(); };
            heightControls.specifiedRadio.onClick = function () {
                syncInputEnabled();
                if (heightControls.specifiedRadio.value) heightControls.heightInput.active = true;
                updatePreview();
            };
            heightControls.heightInput.onChanging = updatePreview;
            changeValueByArrowKey(heightControls.heightInput, updatePreview);

            fitFrameCheckbox.onClick = updatePreview;

            buttons.btnScreenMode.onClick = function () {
                toggleScreenPreviewMode();
                buttons.btnScreenMode.text = getScreenModeButtonLabel();
            };

            buttons.btnCancel.onClick = function () {
                snapshotStore.restoreAll();
                rowHeightDialog.close(0);
            };

            /* 初期状態を範囲に同期してプレビュー / Sync to scope and show the initial preview */
            syncInputEnabled();
            syncTargetEnabled();
            updatePreview();

            if (rowHeightDialog.show() !== 1) return null;

            var dialogResult = {
                height: getRowHeightValue(),
                targetMode: getCurrentTargetMode(),
                scope: getCurrentScope(),
                fitFrame: fitFrameCheckbox.value ? parentFrame : null
            };
            snapshotStore.restoreAll();

            if (dialogResult.height === null) {
                alert(getLabel('alert.invalidNumber'));
                return null;
            }
            return dialogResult;
        }

        // =========================================
        // 適用 / Apply
        // =========================================

        /**
         * 表ごとの対象行に行の高さを適用する（進捗表示つき、1 回の取り消しにまとめる）
         * @param {Array<Array<Row>>} rowSets 表ごとの行の配列
         * @param {number} rowHeight 行の高さ（pt）
         * @param {TextFrame|null} frameToFit 適用後に内容に合わせるフレーム。合わせない場合は null
         * @returns {void}
         */
        function applyRowHeightToRowSets(rowSets, rowHeight, frameToFit) {
            var totalRows = 0;
            for (var i = 0; i < rowSets.length; i++) totalRows += rowSets[i].length;

            var progress = createProgressBar(getLabel('progress.title') + ' ' + SCRIPT_VERSION, totalRows);

            /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
            app.doScript(function () {
                try {
                    var doneRows = 0;
                    for (var j = 0; j < rowSets.length; j++) {
                        for (var k = 0; k < rowSets[j].length; k++) {
                            rowSets[j][k].height = rowHeight;
                            doneRows++;
                            /* 行ごとの更新は重いので間引く / Throttle updates since per-row refresh is costly */
                            if (doneRows === totalRows || doneRows % 10 === 0) {
                                progress.update(doneRows, doneRows + " / " + totalRows);
                            }
                        }
                    }
                    if (frameToFit) fitFrameToContent(frameToFit);
                } finally {
                    progress.close();
                }
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel('undo.applyRowHeight'));
        }

        // =========================================
        // メイン処理 / Main
        // =========================================

        /**
         * 選択から表を特定し、ダイアログの設定で行の高さを適用する
         * @returns {void}
         */
        function main() {
            var originalSelection = collectionToArray(app.selection);
            var selectionTarget = resolveTargetFromSelection(app.selection);

            if (!selectionTarget) {
                alert(getLabel('alert.noTable'));
                return;
            }
            if (selectionTarget.error === "multipleTables") {
                alert(getLabel('alert.multipleTables'));
                return;
            }

            var baseTable = selectionTarget.table;
            var initialHeightPt = getInitialRowHeight(baseTable, selectionTarget.rowIndices);

            /* 表全体のときのみハイライトをオフ / Clear selection highlight only when whole table is targeted */
            var didClearSelection = (selectionTarget.rowIndices === null);
            if (didClearSelection) app.select(NothingEnum.NOTHING);

            var dialogResult = showRowHeightDialog(baseTable, selectionTarget.rowIndices, initialHeightPt);

            if (didClearSelection) restoreSelection(originalSelection);
            if (dialogResult === null) return;

            var targetTables = resolveScopeTables(dialogResult.scope, baseTable);
            var rowSets = collectTargetRowSets(targetTables, baseTable, selectionTarget.rowIndices, dialogResult.targetMode);
            applyRowHeightToRowSets(rowSets, dialogResult.height, dialogResult.fitFrame);
        }

        main();

    })();
