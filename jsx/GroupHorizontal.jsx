#target indesign

function main() {
    if (app.selection.length === 0) {
        alert("アイテムを選択してください。");
        return;
    }

    var doc = app.activeDocument;
    var sel = app.selection;
    var items = [];

    // 選択されたアイテムの情報を取得
    for (var i = 0; i < sel.length; i++) {
        var obj = sel[i];
        // geometricBoundsは [y1, x1, y2, x2] (上、左、下、右)
        var bounds = obj.geometricBounds;
        // Y座標の中心を基準にする（上端にしたい場合は bounds[0] にしてください）
        var centerY = (bounds[0] + bounds[2]) / 2;
        
        items.push({
            obj: obj,
            y: centerY
        });
    }

    // Y座標で昇順にソート（上から下へ）
    items.sort(function(a, b) {
        return a.y - b.y;
    });

    var rows = [];
    var currentRow = [];
    // どの程度のY座標のズレまでを「同じ横並び」とみなすかの許容値（単位は現在のルーラー単位）
    var tolerance = 5; 

    // 横に並んでいるものをグループ分けする
    for (var i = 0; i < items.length; i++) {
        if (currentRow.length === 0) {
            currentRow.push(items[i]);
        } else {
            // 現在の行の最初のアイテムとY座標を比較
            if (Math.abs(items[i].y - currentRow[0].y) <= tolerance) {
                currentRow.push(items[i]);
            } else {
                rows.push(currentRow);
                currentRow = [items[i]];
            }
        }
    }
    if (currentRow.length > 0) {
        rows.push(currentRow);
    }

    // 各行ごとにグループ化を実行
    var groupCount = 0;
    for (var j = 0; j < rows.length; j++) {
        var row = rows[j];
        // グループ化には2つ以上のアイテムが必要
        if (row.length > 1) {
            var itemsToGroup = [];
            for (var k = 0; k < row.length; k++) {
                itemsToGroup.push(row[k].obj);
            }
            doc.groups.add(itemsToGroup);
            groupCount++;
        }
    }

    alert(groupCount + " 個のグループを作成しました。");
}

// 処理をまとめて取り消せるようにする
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "横並びのアイテムをグループ化");
