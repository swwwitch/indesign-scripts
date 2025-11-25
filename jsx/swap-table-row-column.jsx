#target indesign

// Table Transpose (modified for robustness)
// Original: Table Transpose v1.0 by Iain Anderson
// Modified: selection handling & error-proofing

//$.level = 0;
//debugger;

(function () {
    var SCRIPT_VERSION = "v1.0";
    // --- ローカライズヘルパー：日本語 / 英語を自動切り替え ---
    function getLocaleLabels() {
        var isJa = false;
        try {
            // InDesign の UI ロケールが日本語かどうか
            if (app.locale === Locale.JAPANESE) {
                isJa = true;
            }
        } catch (e1) {}
        try {
            // $.locale も補助的にチェック
            if (!isJa && String($.locale).indexOf("ja") === 0) {
                isJa = true;
            }
        } catch (e2) {}

        if (isJa) {
            return {
                dlgTitle: "行と列を入れ替え",
                msgNoDocument: "ドキュメントが開いていません。",
                msgNoSelection: "表、または表を含むテキストフレームを選択してから実行してください。",
                msgNoTable: "選択範囲から表を特定できませんでした。\n表、セル、または表内のテキストを選択して再度お試しください。",
                msgHasMergeStop: "セル結合があるため、処理を中止しました。\nセル結合の扱いを変更して再度お試しください。",
                labelHeaderCheckbox: "ヘッダー行を対象にする",
                panelMergeTitle: "セル結合",
                rbMergeNone: "しない（終了）",
                rbUnmergeAndTranspose: "転置前にセル結合を解除",
                btnOk: "OK",
                btnCancel: "キャンセル"
            };
        } else {
            return {
                dlgTitle: "Transpose Rows and Columns",
                msgNoDocument: "No document is open.",
                msgNoSelection: "Please select a table or a text frame containing a table, then run this script.",
                msgNoTable: "Could not find a table from the selection.\nPlease select a table, cell, or text inside a table and try again.",
                msgHasMergeStop: "The table contains merged cells. The operation has been cancelled.\nChange how merged cells are handled and try again.",
                labelHeaderCheckbox: "Include header rows",
                panelMergeTitle: "Merged cells",
                rbMergeNone: "Do nothing (cancel)",
                rbUnmergeAndTranspose: "Unmerge before transposing",
                btnOk: "OK",
                btnCancel: "Cancel"
            };
        }
    }

    var L = getLocaleLabels();

    if (app.documents.length === 0) {
        alert(L.msgNoDocument);
        return;
    }

    if (app.selection.length === 0) {
        alert(L.msgNoSelection);
        return;
    }

    var sel = app.selection[0];
    var myTable = resolveTableFromSelection(sel);

    if (!myTable) {
        alert(L.msgNoTable);
        return;
    }

    var hasMergeForDialog = hasMergedCells(myTable);
    var hasHeaderForDialog = (myTable.headerRowCount > 0);

    var includeHeader = true;
    var mergeMode = 3; // 1: しない（終了）, 3: 解除してから転置

    // --- 条件によってダイアログボックスをスキップ ---
    // ヘッダー行がなく、かつセル結合もない場合は、
    // 選択肢の余地がないためダイアログを表示せずに既定値で処理を続行
    if (!(hasHeaderForDialog === false && hasMergeForDialog === false)) {
        // --- ダイアログボックス：ヘッダー行の扱いを指定 ---
        // Dialog: choose whether to treat header rows specially
        var dlg = new Window("dialog", L.dlgTitle + "  " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = "left";

        var cbHeader = dlg.add("checkbox", undefined, L.labelHeaderCheckbox);
        cbHeader.value = true; // デフォルト：ヘッダー行も対象にする（現状の挙動）
        if (!hasHeaderForDialog) {
            cbHeader.enabled = false;
            cbHeader.value = false;
        }

        // --- セル結合の扱いパネル ---
        var pnlMerge = dlg.add("panel", undefined, L.panelMergeTitle);
        pnlMerge.orientation = "column";
        pnlMerge.alignChildren = "left";
        pnlMerge.margins = [15, 20, 15, 10];

        var rbMergeNone = pnlMerge.add("radiobutton", undefined, L.rbMergeNone);
        var rbUnmergeAndTranspose = pnlMerge.add("radiobutton", undefined, L.rbUnmergeAndTranspose);
        rbUnmergeAndTranspose.value = true; // デフォルト：現在の挙動（解除してから転置）
        if (!hasMergeForDialog) {
            rbMergeNone.enabled = false;
            rbUnmergeAndTranspose.enabled = false;
        }

        var btnGroup = dlg.add("group");
        btnGroup.alignment = "right";
        var cancelBtn = btnGroup.add("button", undefined, L.btnCancel, { name: "cancel" });
        var okBtn = btnGroup.add("button", undefined, L.btnOk, { name: "ok" });

        if (dlg.show() !== 1) {
            // キャンセルされたときは何もせず終了
            return;
        }

        includeHeader = cbHeader.value;
        if (rbMergeNone.value) {
            mergeMode = 1;
        } else if (rbUnmergeAndTranspose.value) {
            mergeMode = 3;
        } else {
            mergeMode = 3;
        }
    } else {
        // ヘッダーもセル結合もないので、
        // includeHeader = false（意味的には無関係だが明示）として、
        // mergeMode は既定値(3)のままダイアログなしで続行
        includeHeader = false;
        mergeMode = 3;
    }


    // --- 選択から Table を取得するヘルパー ---
    function resolveTableFromSelection(obj) {
        if (!obj) return null;

        // 直接 Table を選択している
        if (obj.constructor.name === "Table") {
            return obj;
        }

        // Cell を選択している
        if (obj.constructor.name === "Cell") {
            return obj.parent; // parent は Table
        }

        // テキスト内にカーソル・範囲選択がある
        try {
            if (obj.tables && obj.tables.length > 0) {
                return obj.tables[0];
            }
        } catch (e) {}

        // テキストフレームや他のオブジェクトの parent に Table がぶら下がっているケース
        try {
            if (obj.parent && obj.parent.tables && obj.parent.tables.length > 0) {
                return obj.parent.tables[0];
            }
        } catch (e2) {}

        return null;
    }

    // --- セル結合が存在するかどうかのチェック ---
    function hasMergedCells(tbl) {
        try {
            var cells = tbl.cells;
            for (var i = 0; i < cells.length; i++) {
                if (cells[i].rowSpan > 1 || cells[i].columnSpan > 1) {
                    return true;
                }
            }
        } catch (e) {}
        return false;
    }

    // --- セル結合の領域情報を収集（最大矩形として扱う前提） ---
    function collectMergedRegions(tbl) {
        var regions = [];
        try {
            var rows = tbl.rows.length;
            var cols = tbl.columns.length;
            for (var r = 0; r < rows; r++) {
                for (var c = 0; c < cols; c++) {
                    var cell = tbl.rows[r].cells[c];
                    var rs = cell.rowSpan;
                    var cs = cell.columnSpan;
                    if (rs > 1 || cs > 1) {
                        // すでに登録済みの領域に含まれていないかチェック
                        var covered = false;
                        for (var i = 0; i < regions.length && !covered; i++) {
                            var reg = regions[i];
                            if (
                                r >= reg.row && r < reg.row + reg.rowSpan &&
                                c >= reg.col && c < reg.col + reg.colSpan
                            ) {
                                covered = true;
                            }
                        }
                        if (!covered) {
                            regions.push({
                                row: r,
                                col: c,
                                rowSpan: rs,
                                colSpan: cs
                            });
                        }
                    }
                }
            }
        } catch (e) {}
        return regions;
    }


    // 元のヘッダー行数・フッター行数を記憶
    // Remember original header/footer row counts
    // includeHeader が false のときはヘッダー行を保持せず、ボディ行として扱う
    var originalHeaderRowCount = includeHeader ? myTable.headerRowCount : 0;
    var originalFooterRowCount = myTable.footerRowCount;

    // --- ここからオリジナルロジックをベースに転置処理 ---

    // セル結合の扱いを mergeMode に応じて制御
    var hasMerge = false;
    try {
        hasMerge = hasMergedCells(myTable);
    } catch (e3) {}

    if (mergeMode === 1) {
        // しない（終了）: セル結合があれば処理を中止
        if (hasMerge) {
            alert(L.msgHasMergeStop);
            return;
        }
        // セル結合がなければ、そのまま転置処理を続行
    } else if (mergeMode === 3) {
        // 転置前にセル結合を解除してから転置
        if (hasMerge) {
            try {
                myTable.unmerge();
            } catch (e5) {
                // unmerge できなくても、とりあえず続行（単純な表なら問題ない）
            }
        }
    }

    var myRows = myTable.rows.length;      // 全行数（ヘッダー/フッター含む）
    var myColumns = myTable.columnCount;   // 列数
    var myOriginalSize = 0;
    var myExtraMode = 0; // 0: 変更なし, 1: 列を増やした, 2: 行を増やした

    // 行数 > 列数 → 列を追加して正方形に
    if (myRows > myColumns) {
        for (var extraCols = myColumns; extraCols < myRows; extraCols++) {
            myTable.columns.add(LocationOptions.atEnd);
        }
        myOriginalSize = myColumns; // 元の列数
        myExtraMode = 1;
    }
    // 行数 < 列数 → 行を追加して正方形に
    else if (myRows < myColumns) {
        for (var extraRows = myRows; extraRows < myColumns; extraRows++) {
            myTable.rows.add(LocationOptions.atEnd);
        }
        myOriginalSize = myRows; // 元の行数
        myExtraMode = 2;
    }

    // サイズ再取得（全行数で再取得）
    myRows = myTable.rows.length;
    myColumns = myTable.columnCount;

    // --- 空セルにスペースを入れてパラグラフを最低1つは作る ---
    var cellCount = myRows * myColumns;
    for (var i = 0; i < cellCount; i++) {
        var c = myTable.cells.item(i);
        try {
            if (c.contents === "") {
                c.contents = " ";
            }
        } catch (e4) {}
    }

    // セルの最初の段落を取得（ない場合は null）
    function getFirstParagraph(cell) {
        try {
            if (cell.paragraphs.length > 0) {
                return cell.paragraphs.item(0);
            }
            if (cell.texts && cell.texts.length > 0 && cell.texts[0].paragraphs.length > 0) {
                return cell.texts[0].paragraphs.item(0);
            }
        } catch (e) {}
        return null;
    }

    // --- 転置：上三角と下三角を入れ替える ---
    for (var row = 0; row < myRows; row++) {
        for (var col = row + 1; col < myColumns; col++) {

            var indexA = col + (row * myColumns);
            var indexB = row + (col * myColumns);

            var cellA = myTable.cells.item(indexA);
            var cellB = myTable.cells.item(indexB);

            // --- テキスト内容 ---
            var tmpContents = cellA.contents;
            cellA.contents = cellB.contents;
            cellB.contents = tmpContents;

            var paraA = getFirstParagraph(cellA);
            var paraB = getFirstParagraph(cellB);

            // --- 段落の pointSize ---
            if (paraA && paraB) {
                try {
                    var tmpSize = paraA.pointSize;
                    paraA.pointSize = paraB.pointSize;
                    paraB.pointSize = tmpSize;
                } catch (e5) {}
            }

            // --- フォントとスタイル ---
            if (paraA && paraB) {
                try {
                    var tmpFont = paraA.appliedFont;
                    var tmpStyle = paraA.fontStyle;

                    paraA.appliedFont = paraB.appliedFont;
                    paraA.fontStyle = paraB.fontStyle;

                    paraB.appliedFont = tmpFont;
                    paraB.fontStyle = tmpStyle;
                } catch (e6) {}
            }

            // --- テキストの塗り色 ---
            if (paraA && paraB) {
                try {
                    var tmpFill = paraA.fillColor;
                    paraA.fillColor = paraB.fillColor;
                    paraB.fillColor = tmpFill;
                } catch (e7) {}
            }

            // --- セルの塗り色 ---
            try {
                var tmpCellFill = cellA.fillColor;
                cellA.fillColor = cellB.fillColor;
                cellB.fillColor = tmpCellFill;
            } catch (e8) {}

            // --- セルのティント ---
            try {
                var tmpTint = cellA.fillTint;
                cellA.fillTint = cellB.fillTint;
                cellB.fillTint = tmpTint;
            } catch (e9) {}
        }
    }

    // --- 余分に増やした行・列を元に戻す（安全版） ---
    if (myExtraMode === 1) {
        // 行が多かった → 列を増やした → 転置後 行数を「元の列数」に合わせる
        while (myTable.rows.length > myOriginalSize) {
            try {
                myTable.rows.lastItem().remove();
            } catch (er1) {
                break;
            }
        }
    } else if (myExtraMode === 2) {
        // 列が多かった → 行を増やした → 転置後 列数を「元の行数」に合わせる
        while (myTable.columnCount > myOriginalSize) {
            try {
                myTable.columns.lastItem().remove();
            } catch (er2) {
                break;
            }
        }
    }

    // --- ヘッダー／フッター行を元の設定に近い形で復元 ---
    // Restore header/footer rows based on original settings
    try {
        var totalRows = myTable.rows.length;
        var newHeaderCount = Math.min(originalHeaderRowCount, totalRows);
        var newFooterCount = Math.min(
            originalFooterRowCount,
            Math.max(0, totalRows - newHeaderCount)
        );
        myTable.headerRowCount = newHeaderCount;
        myTable.footerRowCount = newFooterCount;
    } catch (eRestore) {}

})();