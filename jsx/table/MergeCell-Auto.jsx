/*
    自動でセル結合
    説明: 選択した表について、同一内容のセルを水平方向・垂直方向に自動結合します。
    作成日: 2025-12-05
    更新日: 2025-12-05
*/

#target "indesign"

var SCRIPT_NAME    = "自動でセル結合";
var SCRIPT_VERSION = "v0.2";

main();

function main() {
    if (app.documents.length === 0) {
        return;
    }

    // --- 選択チェックと対象テーブルの取得 ---
    if (app.selection.length === 0) {
        alert("表、または表の中のセルを選択してください。");
        return;
    }

    var targetTable = null;
    var sel = app.selection[0];

    // 表の取得判定
    if (sel.constructor.name === "Table") {
        targetTable = sel;
    } else if (sel.parent && sel.parent.constructor.name === "Table") {
        targetTable = sel.parent;
    } else if (sel.tables && sel.tables.length > 0) {
        targetTable = sel.tables[0];
    } else if (sel.parent && sel.parent.parent && sel.parent.parent.constructor.name === "Table") {
        targetTable = sel.parent.parent;
    }

    if (!targetTable) {
        alert("表が見つかりません。");
        return;
    }

    // DIALOG
    // ======
    var dialog = new Window("dialog", SCRIPT_NAME + " " + SCRIPT_VERSION);
    dialog.orientation = "row";
    dialog.alignChildren = ["center","top"];
    dialog.spacing = 10;
    dialog.margins = 20;

    // GROUP1
    // ======
    var group1 = dialog.add("group", undefined, {name: "group1"});
    group1.orientation = "column";
    group1.alignChildren = ["fill","center"];
    group1.spacing = 10;
    group1.margins = 0;

    // PANEL1: 結合
    // =============
    var panel1 = group1.add("panel", undefined, undefined, {name: "panel1"});
    panel1.text = "結合";
    panel1.orientation = "column";
    panel1.alignChildren = ["left","top"];
    panel1.spacing = 10;
    panel1.margins = [15,20,15,10];

    var rbSingle = panel1.add("radiobutton", undefined, undefined, {name: "radiobutton1"});
    rbSingle.text = "単一方向";
    rbSingle.value = false;

    var rbBoth = panel1.add("radiobutton", undefined, undefined, {name: "radiobutton2"});
    rbBoth.text = "両方向（優先付き）";
    rbBoth.value = true; // デフォルト（両方向）

    // PANEL2: 方向
    // ============
    var panel2 = group1.add("panel", undefined, undefined, {name: "panel2"});
    panel2.text = "方向";
    panel2.orientation = "column";
    panel2.alignChildren = ["left","top"];
    panel2.spacing = 10;
    panel2.margins = [15,20,15,10];

    var rbHorizontal = panel2.add("radiobutton", undefined, undefined, {name: "radiobutton3"});
    rbHorizontal.text = "水平";
    rbHorizontal.value = true; // デフォルト

    var rbVertical = panel2.add("radiobutton", undefined, undefined, {name: "radiobutton4"});
    rbVertical.text = "垂直";

    // PANEL3: 対象
    // ==========
    var panel3 = group1.add("panel", undefined, undefined, {name: "panel3"});
    panel3.text = "対象";
    panel3.orientation = "column";
    panel3.alignChildren = ["left","top"];
    panel3.spacing = 10;
    panel3.margins = [15,20,15,10];

    var rbTargetAll = panel3.add("radiobutton", undefined, undefined, {name: "radiobutton5"});
    rbTargetAll.text = "表全体";
    rbTargetAll.value = true; // デフォルト

    var rbTargetSelection = panel3.add("radiobutton", undefined, undefined, {name: "radiobutton6"});
    rbTargetSelection.text = "選択セルのみ";

    // GROUP2: ボタン
    // ==============
    var group2 = dialog.add("group", undefined, {name: "group2"});
    group2.orientation = "column";
    group2.alignChildren = ["fill","center"];
    group2.spacing = 10;
    group2.margins = 0;

    var btnOk = group2.add("button", undefined, undefined, {name: "button1"});
    btnOk.text = "OK";
    btnOk.onClick = function () {
        dialog.close(1);
    };

    var btnCancel = group2.add("button", undefined, undefined, {name: "button2"});
    btnCancel.text = "キャンセル";
    btnCancel.onClick = function () {
        dialog.close(0);
    };

    // ダイアログ表示
    var result = dialog.show();

    if (result === 1) {
        // ユーザーの選択を記録
        var mode;       // "horizontalOnly", "verticalOnly", "horizontalFirst", "verticalFirst"
        var cellRange;  // 対象セル範囲（null の場合は表全体）

        // 対象範囲の決定
        if (rbTargetAll.value) {
            // 表全体
            cellRange = null;
        } else {
            // 選択セルの矩形範囲のみ
            cellRange = getSelectedCellRange(targetTable);
            if (!cellRange) {
                alert("「選択セルの矩形範囲のみ」が選択されていますが、有効なセル選択が見つかりませんでした。\\n表全体を対象にするか、複数セルを選択してください。");
                return;
            }
        }

        // 単一方向か両方向か
        if (rbSingle.value) { // 「単一方向」が選択されている場合
            if (rbHorizontal.value) {
                mode = "horizontalOnly";
            } else {
                mode = "verticalOnly";
            }
        } else { // rbBoth.value （「両方向（優先付き）」が選択されている場合）
            if (rbHorizontal.value) {
                mode = "horizontalFirst";
            } else {
                mode = "verticalFirst";
            }
        }

        // 処理実行（元スクリプトのロジック）
        app.doScript(function () {
            switch (mode) {
                case "horizontalOnly":
                    mergeHorizontalSafe(targetTable, cellRange);
                    break;
                case "verticalOnly":
                    mergeVerticalSafe(targetTable, cellRange);
                    break;
                case "horizontalFirst":
                    mergeHorizontalSafe(targetTable, cellRange);
                    mergeVerticalSafe(targetTable, cellRange);
                    break;
                case "verticalFirst":
                    mergeVerticalSafe(targetTable, cellRange);
                    mergeHorizontalSafe(targetTable, cellRange);
                    break;
            }
        }, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, SCRIPT_NAME);
    }

    // ScriptUI ダイアログは show() 後、自動的に破棄されるので明示的 destroy() は不要
}

// ===== テキスト処理ユーティリティ =====

// --- カスタムトリム関数（String.prototype.trim() が使えない環境用） ---
function customTrim(str) {
    if (typeof str !== "string") {
        str = String(str); // 文字列であることを保証
    }
    // 正規表現で前後の空白文字（スペース、タブ、改行など）を削除
    return str.replace(/^\s+|\s+$/g, "");
}

// --- テキスト取得用（強化版）：空白文字をすべて除去して比較 ---
function getText(cell) {
    if (!cell || !cell.isValid) return ""; // 無効なセルからのアクセスを防止

    var content = cell.contents;

    // cell.contentsが配列で返る場合に対応
    if (content instanceof Array) {
        content = content.join(""); // 配列を結合して単一の文字列にする
    }

    // customTrim 関数を使用して前後の空白を除去
    return customTrim(content);
}

// --- 選択セルの矩形範囲を取得 ---
function getSelectedCellRange(table) {
    var cells = table.cells;
    if (!cells || cells.length === 0) return null;

    var minRow = Number.MAX_VALUE;
    var maxRow = -1;
    var minCol = Number.MAX_VALUE;
    var maxCol = -1;

    for (var i = 0; i < cells.length; i++) {
        var cell = cells[i];
        if (!cell || !cell.isValid) continue;

        // セルが選択されているか判定
        // 一般的な InDesign DOM では many pageItems have a 'selected' プロパティを持つ
        // 持たない場合は常に false になるので安全
        if (!cell.selected) continue;

        var r = cell.row.index;
        var c = cell.column.index;

        if (r < minRow) minRow = r;
        if (r > maxRow) maxRow = r;
        if (c < minCol) minCol = c;
        if (c > maxCol) maxCol = c;
    }

    if (maxRow < 0 || maxCol < 0) {
        // 選択セルが見つからなかった
        return null;
    }

    return {
        rowStart: minRow,
        rowEnd:   maxRow,
        colStart: minCol,
        colEnd:   maxCol
    };
}

// ===== セル結合ロジック =====

// --- 水平方向（左右）の結合：安全版 ---
function mergeHorizontalSafe(table, range) {
    var rows = table.rows;
    var numRows = rows.length;

    for (var r = 0; r < numRows; r++) {
        var currentRow = rows[r];
        if (!currentRow.isValid) continue;

        // 行範囲の制限（range が指定されている場合）
        if (range && (r < range.rowStart || r > range.rowEnd)) {
            continue;
        }

        var cells = currentRow.cells;
        var cellIndex = 0;

        while (cellIndex < cells.length - 1) {
            var cellA = cells[cellIndex];
            var cellB = cells[cellIndex + 1];

            if (!cellA || !cellA.isValid || !cellB || !cellB.isValid) {
                cellIndex++;
                continue;
            }

            // 列範囲の制限（range が指定されている場合）
            if (range) {
                var colIndexA = cellA.column.index;
                var colIndexB = cellB.column.index;
                if (colIndexA < range.colStart || colIndexB > range.colEnd) {
                    cellIndex++;
                    continue;
                }
            }

            var t1 = getText(cellA);
            var t2 = getText(cellB);

            if (t1 === t2) {
                try {
                    cellA.merge(cellB);
                    cellA.contents = t1;
                    cells = currentRow.cells; // 結合後にセル配列を更新
                } catch (e) {
                    cellIndex++;
                }
            } else {
                cellIndex++;
            }
        }
    }
}

// --- 垂直方向（上下）の結合：安全版 ---
function mergeVerticalSafe(table, range) {
    var numCols = table.columns.length;

    for (var c = 0; c < numCols; c++) {
        var currentColumn = table.columns[c];
        if (!currentColumn.isValid) continue;

        // 列範囲の制限（range が指定されている場合）
        if (range && (c < range.colStart || c > range.colEnd)) {
            continue;
        }

        var cells = currentColumn.cells;
        var cellIndex = 0;

        while (cellIndex < cells.length - 1) {
            var cellA = cells[cellIndex];
            var cellB = cells[cellIndex + 1];

            if (!cellA || !cellA.isValid || !cellB || !cellB.isValid) {
                cellIndex++;
                continue;
            }

            // 行範囲の制限（range が指定されている場合）
            if (range) {
                var rowIndexA = cellA.row.index;
                var rowIndexB = cellB.row.index;
                if (rowIndexA < range.rowStart || rowIndexB > range.rowEnd) {
                    cellIndex++;
                    continue;
                }
            }

            var t1 = getText(cellA);
            var t2 = getText(cellB);

            if (t1 === t2) {
                try {
                    cellA.merge(cellB);
                    cellA.contents = t1;
                    cells = currentColumn.cells; // 結合後にセル配列を更新
                } catch (e) {
                    cellIndex++;
                }
            } else {
                cellIndex++;
            }
        }
    }
}