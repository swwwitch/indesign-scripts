#target indesign

// メイン関数
function main() {
    var doc = app.activeDocument;
    var sel = app.selection;

    if (sel.length == 0 || !(sel[0].hasOwnProperty('baseline'))) {
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
                    var table = frame.tables[i];
                    if (table.storyOffset.index <= cursor.index && cursor.index <= table.storyOffset.index + table.characters.length) {
                        parentTable = table;
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

    // テキストフレームの幅を取得
    var textFrame = parentTable.parent;
    while (textFrame.constructor.name !== "TextFrame" && textFrame.constructor.name !== "Story") {
        textFrame = textFrame.parent;
    }

    if (textFrame.constructor.name === "TextFrame") {
        var frameWidth = textFrame.geometricBounds[3] - textFrame.geometricBounds[1];

        // テーブルの幅をテキストフレームの幅に設定
        parentTable.width = frameWidth;
    } else {
        alert("テーブルがテキストフレーム内に見つかりません。");
    }
}

main();