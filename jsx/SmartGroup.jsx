#target indesign

/*
 * SmartGroup.jsx
 *
 * 選択中のアイテムを「水平方向」または「垂直方向」の近接距離でグループ化するスクリプト。
 *
 * 【動作フロー】
 *  1. ダイアログで方向（水平 / 垂直）と許容値（アイテム間の隙間）を設定
 *  2. ダイアログ表示中、グループ化される範囲を赤いプレビューフレームでリアルタイム表示
 *     - 水平方向：右端 → 次のアイテムの左端の隙間が許容値以内なら同グループ
 *     - 垂直方向：下端 → 次のアイテムの上端の隙間が許容値以内なら同グループ
 *  3. OK でプレビューフレームを削除してグループ化を実行
 *     キャンセルでプレビューフレームを削除して中止
 *
 * 【プレビューレイヤー】
 *  "SmartGroup Preview"（非印刷）レイヤーに一時描画。スクリプト終了時に自動削除。
 */

// プレビューフレームの参照を保持するグローバル配列
var previewFrames = [];

// ── ユーティリティ ────────────────────────────────────────

function getRedColor(doc) {
    var name = "SmartGroup_Preview_Red";
    for (var i = 0; i < doc.colors.length; i++) {
        if (doc.colors[i].name === name) return doc.colors[i];
    }
    return doc.colors.add({
        name: name,
        model: ColorModel.PROCESS,
        space: ColorSpace.CMYK,
        colorValue: [0, 100, 100, 0]
    });
}

function getPreviewLayer(doc) {
    var name = "SmartGroup Preview";
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === name) return doc.layers[i];
    }
    return doc.layers.add({ name: name, printable: false });
}

// ── グループ計算 ──────────────────────────────────────────
// 水平方向（横並び）：Y中心が近いもの同士＝同じ行。許容値＝Y方向のズレ量
// 垂直方向（縦並び）：X中心が近いもの同士＝同じ列。許容値＝X方向のズレ量

function computeGroups(items, direction, tolerance) {
    var sorted = items.slice();
    if (direction === "horizontal") {
        // Y中心でソートして、Y差が tolerance 以内なら同じ行
        sorted.sort(function (a, b) {
            var ay = (a.top + a.bottom) / 2;
            var by = (b.top + b.bottom) / 2;
            return ay - by;
        });
    } else {
        // X中心でソートして、X差が tolerance 以内なら同じ列
        sorted.sort(function (a, b) {
            var ax = (a.left + a.right) / 2;
            var bx = (b.left + b.right) / 2;
            return ax - bx;
        });
    }

    var groups = [];
    var current = [];

    for (var i = 0; i < sorted.length; i++) {
        if (current.length === 0) {
            current.push(sorted[i]);
        } else {
            var prev = current[current.length - 1];
            var delta;
            if (direction === "horizontal") {
                var cy = (sorted[i].top + sorted[i].bottom) / 2;
                var py = (prev.top + prev.bottom) / 2;
                delta = Math.abs(cy - py); // Y方向のズレ
            } else {
                var cx = (sorted[i].left + sorted[i].right) / 2;
                var px = (prev.left + prev.right) / 2;
                delta = Math.abs(cx - px); // X方向のズレ
            }
            if (delta <= tolerance) {
                current.push(sorted[i]);
            } else {
                groups.push(current);
                current = [sorted[i]];
            }
        }
    }
    if (current.length > 0) groups.push(current);
    return groups;
}

function getGroupBounds(group) {
    var top = group[0].top, left = group[0].left,
        bottom = group[0].bottom, right = group[0].right;
    for (var i = 1; i < group.length; i++) {
        top = Math.min(top, group[i].top);
        left = Math.min(left, group[i].left);
        bottom = Math.max(bottom, group[i].bottom);
        right = Math.max(right, group[i].right);
    }
    return [top, left, bottom, right];
}

// ── プレビューフレーム ────────────────────────────────────

// プレビューレイヤー上の全アイテムを削除（配列への参照ズレを防ぐため層ごと掃除）
function clearPreview(doc) {
    var name = "SmartGroup Preview";
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === name) {
            try { doc.layers[i].remove(); } catch (e) {}
            break;
        }
    }
    previewFrames = [];
    try { app.redraw(); } catch (e) {}
}

function updatePreview(doc, items, direction, tolerance) {
    // まずレイヤー上の全フレームを削除してから描き直す
    clearPreview(doc);

    var groups = computeGroups(items, direction, tolerance);
    var layer  = getPreviewLayer(doc);
    var red    = getRedColor(doc);
    var none   = doc.swatches.itemByName("[None]");

    for (var j = 0; j < groups.length; j++) {
        if (groups[j].length < 2) continue;
        var page = groups[j][0].obj.parentPage;
        if (!page) continue;
        var b    = getGroupBounds(groups[j]);
        var rect = page.rectangles.add(layer);
        rect.geometricBounds = b;
        rect.fillColor       = red;
        rect.strokeColor     = none;
        rect.opacity         = 40;
        previewFrames.push(rect);
    }
    try { app.redraw(); } catch (e) {}
}

// ── ダイアログ ────────────────────────────────────────────

function showDialog(doc, items) {
    var dlg = new Window("dialog", "グループ化の設定");
    dlg.orientation = "column";
    dlg.alignChildren = "left";
    dlg.margins = 20;
    dlg.spacing = 14;

    // パネル：グループ化する方向
    var dirPanel = dlg.add("panel", undefined, "グループ化する方向");
    dirPanel.orientation = "column";
    dirPanel.alignChildren = "left";
    dirPanel.margins = [10, 15, 10, 10];
    dirPanel.spacing = 8;

    var rbHorizontal = dirPanel.add("radiobutton", undefined, "水平方向（横並び）");
    var rbVertical = dirPanel.add("radiobutton", undefined, "垂直方向（縦並び）");
    rbHorizontal.value = true;

    // パネル：許容値
    var tolPanel = dlg.add("panel", undefined, "許容値");
    tolPanel.orientation = "column";
    tolPanel.alignChildren = "left";
    tolPanel.margins = [10, 15, 10, 10];

    var sliderGroup = tolPanel.add("group");
    sliderGroup.orientation = "row";
    sliderGroup.alignChildren = "center";

    var slider = sliderGroup.add("slider", undefined, 5, 0, 50);
    slider.preferredSize.width = 180;
    var valueLabel = sliderGroup.add("statictext", undefined, "5");
    valueLabel.preferredSize.width = 30;

    function refresh() {
        var dir = rbHorizontal.value ? "horizontal" : "vertical";
        var tol = Math.round(slider.value);
        valueLabel.text = String(tol);
        updatePreview(doc, items, dir, tol);
    }

    slider.onChanging = refresh;
    rbHorizontal.onClick = refresh;
    rbVertical.onClick = refresh;
    dlg.onShow = function () { refresh(); };

    var btnGroup = dlg.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignment = "center";
    btnGroup.add("button", undefined, "キャンセル", { name: "cancel" });
    btnGroup.add("button", undefined, "OK", { name: "ok" });

    var ok = dlg.show() === 1;

    // ダイアログを閉じたらプレビューフレームを削除
    clearPreview(doc);

    if (!ok) return null;
    return {
        direction: rbHorizontal.value ? "horizontal" : "vertical",
        tolerance: Math.round(slider.value)
    };
}

// ── メイン ────────────────────────────────────────────────

function main() {
    if (app.selection.length === 0) {
        alert("アイテムを選択してください。");
        return;
    }

    var doc = app.activeDocument;
    var sel = app.selection;
    var items = [];

    // geometricBounds = [top, left, bottom, right]
    for (var i = 0; i < sel.length; i++) {
        var b = sel[i].geometricBounds;
        items.push({ obj: sel[i], top: b[0], left: b[1], bottom: b[2], right: b[3] });
    }

    var result = showDialog(doc, items);
    if (result === null) return; // キャンセル

    var groups = computeGroups(items, result.direction, result.tolerance);
    var groupCount = 0;

    for (var j = 0; j < groups.length; j++) {
        var row = groups[j];
        if (row.length > 1) {
            app.select(row[0].obj);
            for (var k = 1; k < row.length; k++) {
                app.select(row[k].obj, SelectionOptions.ADD_TO);
            }
            doc.groups.add(app.selection);
            groupCount++;
        }
    }

    var label = result.direction === "horizontal" ? "水平方向" : "垂直方向";
    alert(label + "で " + groupCount + " 個のグループを作成しました。");
}

// 処理をまとめて取り消せるようにする
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "スマートグループ化");
