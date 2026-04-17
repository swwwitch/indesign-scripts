#target indesign

// メイン関数
function main() {
    var doc = app.activeDocument;
    var sel = app.selection;

    if (sel.length == 0 || !(sel[0].hasOwnProperty('baseline')) || !(sel[0].constructor.name === "InsertionPoint" || sel[0].constructor.name === "Text")) {
        alert("テーブル内にカーソルを置いてください。");
        return;
    }

    var cursor = sel[0];
    var parentTable = null;

    // テーブルを含むセルを見つける
    if (cursor.constructor.name === "InsertionPoint") {
        if (cursor.parent.constructor.name === "Cell") {
            parentTable = cursor.parent.parent;
        }
    } else if (cursor.constructor.name === "Text") {
        if (cursor.parentTextFrames.length > 0) {
            var frame = cursor.parentTextFrames[0];
            if (frame.tables.length > 0) {
                for (var i = 0; i < frame.tables.length; i++) {
                    if (frame.tables[i].storyOffset <= cursor.index && frame.tables[i].storyOffset + frame.tables[i].characters.length >= cursor.index) {
                        parentTable = frame.tables[i];
                        break;
                    }
                }
            }
        }
    }

    if (!parentTable) {
        alert("テーブル内にカーソルを置いてください。");
        return;
    }

    // テーブルが含まれているテキストフレームの幅を取得
    var textFrameWidth = parentTable.parent.geometricBounds[3] - parentTable.parent.geometricBounds[1];

    // テーブルの現在の幅を取得
    var currentTableWidth = parentTable.width;

    // 最終列を除くすべての列幅の合計を計算
    var totalOtherColumnsWidth = 0;
    for (var col = 0; col < parentTable.columns.length - 1; col++) {
        totalOtherColumnsWidth += parentTable.columns[col].width;
    }

    // 最終列の幅を計算して設定
    var newLastColumnWidth = textFrameWidth - totalOtherColumnsWidth;
    parentTable.columns[parentTable.columns.length - 1].width = newLastColumnWidth;
}

main();