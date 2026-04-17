#target indesign

var SCRIPT_VERSION = "v1.0";

// ローカライズ設定
// InDesign のロケールや ExtendScript のロケールから日本語かどうかを判定
var IS_JAPANESE_UI = (function() {
    try {
        if (app.locale && app.locale === Locale.JAPANESE) return true;
    } catch (e) {}
    try {
        if ($.locale && $.locale.toString().indexOf("ja") === 0) return true;
    } catch (e2) {}
    return false;
})();

// 簡易ローカライズ関数：_(ja, en)
function _(ja, en) {
    return IS_JAPANESE_UI ? ja : en;
}

// セルの結合解除スクリプト
// ・結合したセルを解除
// ・「テキストを分配」オンのとき、解除後のすべてのセルに同じテキストを複製
// ・「表全体」 / 「選択したセルのみ」 対応

main();

function main() {
    if (app.documents.length === 0) {
        alert(_("ドキュメントが開かれていません。", "No document is open."));
        return;
    }
    if (!app.selection || app.selection.length === 0) {
        alert(_("表のセルを選択してください。", "Please select table cells."));
        return;
    }

    var ui = createDialog();
    var dlg = ui.window;
    var result = dlg.show();

    // キャンセル時は何もしない
    if (result !== 1) {
        return;
    }

    var distributeText = ui.rbDistribute.value; // 「テキストを分配」
    var wholeTable = ui.rbWholeTable.value; // 「表全体」

    app.doScript(
        function() {
            runUnmerge(distributeText, wholeTable);
        },
        ScriptLanguage.JAVASCRIPT,
        undefined,
        UndoModes.ENTIRE_SCRIPT,
        _("セルの結合解除", "Unmerge Cells")
    );
}

// ダイアログ生成
function createDialog() {
    // DIALOG
    var dialog = new Window("dialog");
    dialog.text = _("セルの結合解除", "Unmerge Cells") + " " + SCRIPT_VERSION;
    dialog.orientation = "column";
    dialog.alignChildren = ["fill", "top"];
    dialog.spacing = 10;
    dialog.margins = 20;

    // PANEL1: 結合
    var panel1 = dialog.add("panel", undefined, undefined, {
        name: "panel1"
    });
    panel1.text = _("結合", "Merge");
    panel1.orientation = "column";
    panel1.alignChildren = ["left", "top"];
    panel1.spacing = 10;
    panel1.margins = [15, 20, 15, 10];

    var radiobutton1 = panel1.add("radiobutton", undefined, undefined, {
        name: "radiobutton1"
    });
    radiobutton1.text = _("デフォルト（分配なし）", "Default (no distribution)");

    var radiobutton2 = panel1.add("radiobutton", undefined, undefined, {
        name: "radiobutton2"
    });
    radiobutton2.text = _("テキストを分配", "Distribute text");
    radiobutton2.value = true; // デフォルト

    // PANEL2: 対象
    var panel2 = dialog.add("panel", undefined, undefined, {
        name: "panel2"
    });
    panel2.text = _("対象", "Scope");
    panel2.orientation = "column";
    panel2.alignChildren = ["fill", "top"];
    panel2.spacing = 9;
    panel2.margins = [15, 20, 15, 10];

    var radiobutton3 = panel2.add("radiobutton", undefined, undefined, {
        name: "radiobutton3"
    });
    radiobutton3.text = _("表全体", "Whole table");
    radiobutton3.value = true; // デフォルト

    var radiobutton4 = panel2.add("radiobutton", undefined, undefined, {
        name: "radiobutton4"
    });
    radiobutton4.text = _("選択したセルのみ", "Selected cells only");

    // GROUP1: ボタン行
    var group1 = dialog.add("group", undefined, {
        name: "group1"
    });
    group1.orientation = "row";
    group1.alignChildren = ["center", "center"];
    group1.spacing = 10;
    group1.margins = [0, 10, 0, 0];

    var buttonCancel = group1.add("button", undefined, undefined, {
        name: "button1"
    });
    buttonCancel.text = _("キャンセル", "Cancel");

    var buttonOk = group1.add("button", undefined, undefined, {
        name: "button2"
    });
    buttonOk.text = _("OK", "OK");

    // ボタン動作
    buttonCancel.onClick = function() {
        dialog.close(0);
    };
    buttonOk.onClick = function() {
        dialog.close(1);
    };

    return {
        window: dialog,
        rbDistribute: radiobutton2,
        rbWholeTable: radiobutton3
    };
}

// 結合解除本体
// distributeText: true のとき、解除後のセルすべてに同じテキストを複製
// wholeTable: true なら表全体、false なら選択セルのみ
function runUnmerge(distributeText, wholeTable) {
    var sel = app.selection;
    if (!sel || sel.length === 0) {
        alert(_("表のセルを選択してください。", "Please select table cells."));
        return;
    }

    var table = getTableFromSelection(sel[0]);
    if (!table) {
        alert(_("表のセルを選択してください。", "Please select table cells."));
        return;
    }

    var cellsToCheck;

    if (wholeTable) {
        // 表全体
        cellsToCheck = table.cells.everyItem().getElements();
    } else {
        // 選択したセルのみ
        cellsToCheck = getSelectedCells(sel, table);
        if (!cellsToCheck || cellsToCheck.length === 0) {
            alert(_("対象となるセルが選択されていません。", "No target cells are selected."));
            return;
        }
    }

    cellsToCheck = uniqueCells(cellsToCheck);

    for (var i = 0; i < cellsToCheck.length; i++) {
        var c = cellsToCheck[i];
        if (!c || !c.isValid) {
            continue;
        }

        // 結合セルかどうか（行方向 or 列方向に 2 つ以上を跨いでいる）
        if (c.rowSpan > 1 || c.columnSpan > 1) {
            var originalContents = c.contents;

            // unmerge() は解除後のセル配列を返す
            var newCells;
            try {
                newCells = c.unmerge();
            } catch (e) {
                continue;
            }

            if (distributeText && originalContents !== "") {
                for (var j = 0; j < newCells.length; j++) {
                    if (!newCells[j] || !newCells[j].isValid) {
                        continue;
                    }
                    try {
                        newCells[j].contents = originalContents;
                    } catch (e2) {
                        // 何かあってもスクリプトごと止まらないように握りつぶす
                    }
                }
            }
        }
    }
}

// 選択状態から表を取得
function getTableFromSelection(obj) {
    if (!obj) {
        return null;
    }

    var name = obj.constructor && obj.constructor.name;

    if (name === "Table") {
        return obj;
    }

    if (name === "Cell") {
        return getParentTable(obj);
    }

    if (name === "Cells") {
        if (obj.length > 0) {
            return getParentTable(obj[0]);
        }
        return null;
    }

    return getParentTable(obj);
}

// 任意のオブジェクトから親方向に辿って Table を探す
function getParentTable(obj) {
    var p = obj;
    while (p) {
        if (p.constructor && p.constructor.name === "Table") {
            return p;
        }
        if (!p.parent || p.parent === p) {
            break;
        }
        p = p.parent;
    }
    return null;
}

// 選択範囲から対象のセル配列を取得
function getSelectedCells(selection, table) {
    var result = [];

    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        var name = item.constructor && item.constructor.name;

        if (name === "Cell") {
            if (getParentTable(item) === table) {
                result.push(item);
            }
        } else if (name === "Cells") {
            var arr = item.getElements();
            for (var j = 0; j < arr.length; j++) {
                if (getParentTable(arr[j]) === table) {
                    result.push(arr[j]);
                }
            }
        } else {
            // テキストオブジェクトなど、セル内部が選択されているケース
            var p = item;
            while (p && p !== p.parent) {
                if (p.constructor && p.constructor.name === "Cell") {
                    if (getParentTable(p) === table) {
                        result.push(p);
                    }
                    break;
                }
                p = p.parent;
            }
        }
    }

    return result;
}

// セル配列の重複排除（id ベース）
function uniqueCells(cells) {
    var result = [];
    var seen = {};

    for (var i = 0; i < cells.length; i++) {
        var c = cells[i];
        if (!c || !c.isValid) {
            continue;
        }
        var id = c.id;
        if (!seen[id]) {
            seen[id] = true;
            result.push(c);
        }
    }

    return result;
}