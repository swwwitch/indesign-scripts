#target "InDesign"

/*
 * スクリプトの概要：
 * 選択中の表セルに対して、罫線の描画・消去を切り替えながら適用できるスクリプトです。
 * 「すべて」「境界線のみ」「内部のみ」「水平線のみ」「垂直線のみ」「すべて消去」に対応しています。
 *
 * 線幅は mm 固定のドロップダウンから選択し、カラーはドキュメントのスウォッチから選択できます。
 * プレビューを確認しながら適用でき、選択したスウォッチは罫線色として反映されます。
 *
 * 「描画前に消去」を OFF にすると、既存の罫線を残したまま上書きできるため、
 * 例：すべてを細線 → 境界線のみを太線、のような段階的な組み合わせ調整が可能です。
 *
 * 主な機能：
 * - 選択中の表セルに対する罫線の描画／消去
 * - 外枠・内部・水平・垂直の各モード切り替え
 * - 「描画前に消去」の ON/OFF 切り替え
 * - mm 固定の線幅指定
 * - スウォッチによる罫線カラー指定
 * - プレビュー表示による確認
 * - 日本語／英語UI対応
 */

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "罫線の設定", en: "Border Settings" },
    panelDraw: { ja: "罫線の描画", en: "Draw Borders" },
    panelErase: { ja: "オプション", en: "Settings" },
    clearFirst: { ja: "描画前に消去", en: "Clear Before Draw" },
    all: { ja: "すべて", en: "All" },
    outer: { ja: "境界線のみ", en: "Outer Only" },
    inner: { ja: "内部のみ", en: "Inner Only" },
    horizontal: { ja: "水平線のみ", en: "Horizontal Only" },
    vertical: { ja: "垂直線のみ", en: "Vertical Only" },
    allOff: { ja: "すべて消去", en: "Clear All" },
    outerOnly: { ja: "境界線を消去", en: "Clear Outer Border" },
    lineWidth: { ja: "線幅：", en: "Stroke Weight:" },
    lineWidthUnit: { ja: "mm", en: "mm" },
    color: { ja: "カラー：", en: "Stroke Color:" },
    preview: { ja: "プレビュー", en: "Preview" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    alertSelect: { ja: "表のセルを選択してください。", en: "Please select table cells." },
    alertWeight: { ja: "線幅には0以上の数値を入力してください。", en: "Enter a value of 0 or greater for weight." }
};

function L(key) {
    return LABELS[key] ? LABELS[key][lang] : key;
}

(function () {
    var cells = getSelectedCellsFromApp();
    if (cells.length === 0) {
        alert(L('alertSelect'));
        return;
    }

    var state = {
        cells: cells,
        previewed: false
    };

    app.selection = NothingEnum.NOTHING;

    var ui = buildDialog();
    bindDialogEvents(ui, state);

    var result = ui.dlg.show();
    if (result != 1) {
        clearPreview(state);
        return;
    }

    applyFinalFromDialog(ui, state);

    // =========================================
    // UI構築 / Build UI
    // =========================================
    function buildDialog() {
        var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = "fill";

        var columns = dlg.add("group");
        columns.orientation = "row";
        columns.alignChildren = ["fill", "top"];

        var panelDraw = columns.add("panel", undefined, L('panelDraw'));
        panelDraw.orientation = "column";
        panelDraw.alignChildren = "left";
        panelDraw.margins = [15, 20, 15, 10];

        var rbAll = panelDraw.add("radiobutton", undefined, L('all'));
        var rbOuter = panelDraw.add("radiobutton", undefined, L('outer'));
        var rbInnerOnly = panelDraw.add("radiobutton", undefined, L('inner'));
        var rbHorzOnly = panelDraw.add("radiobutton", undefined, L('horizontal'));
        var rbVertOnly = panelDraw.add("radiobutton", undefined, L('vertical'));
        var rbAllOff = panelDraw.add("radiobutton", undefined, L('allOff'));

        var panelErase = columns.add("panel", undefined, L('panelErase'));
        panelErase.orientation = "column";
        panelErase.alignChildren = "left";
        panelErase.margins = [15, 20, 15, 10];

        var cbClearFirst = panelErase.add("checkbox", undefined, L('clearFirst'));
        cbClearFirst.value = true;

        var weightRow = panelErase.add("group");
        weightRow.orientation = "row";
        weightRow.alignChildren = ["left", "center"];
        weightRow.alignment = ["fill", "top"];
        weightRow.spacing = 8;

        weightRow.add("statictext", undefined, L('lineWidth'));
        var weightValues = ["0", "0.1", "0.25", "0.35", "0.5", "0.75", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "15", "20", "25", "30"];
        var weightDropdown = weightRow.add("dropdownlist", undefined, weightValues);
        weightDropdown.selection = getDefaultWeightIndex(weightValues);
        weightDropdown.minimumSize.width = 60;
        weightRow.add("statictext", undefined, L('lineWidthUnit'));

        panelErase.add("statictext", undefined, L('color'));

        var colorNames = getSwatchNames();
        var colorDropdown = panelErase.add("dropdownlist", undefined, colorNames);
        colorDropdown.minimumSize.height = 22;
        colorDropdown.minimumSize.width = 160;
        colorDropdown.preferredSize.width = 140;
        colorDropdown.selection = getDefaultColorIndex(colorNames);

        rbAll.value = true;


        var btnRowGroup = dlg.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["left", "center"];
        btnRowGroup.alignment = "fill";

        var cbPreview = btnRowGroup.add("checkbox", undefined, L('preview'));
        cbPreview.value = true;

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 40;

        var btnCancel = btnRowGroup.add("button", undefined, L('cancel'), { name: "cancel" });
        var btnOk = btnRowGroup.add("button", undefined, L('ok'), { name: "ok" });

        return {
            dlg: dlg,
            rbAll: rbAll,
            rbOuter: rbOuter,
            rbInnerOnly: rbInnerOnly,
            rbHorzOnly: rbHorzOnly,
            rbVertOnly: rbVertOnly,
            rbAllOff: rbAllOff,
            weightDropdown: weightDropdown,
            colorDropdown: colorDropdown,

            cbClearFirst: cbClearFirst,
            cbPreview: cbPreview,
            btnCancel: btnCancel,
            btnOk: btnOk,
            drawButtons: [rbAll, rbOuter, rbInnerOnly, rbHorzOnly, rbVertOnly, rbAllOff],
            eraseButtons: []
        };
    }

    // =========================================
    // イベント / Events
    // =========================================
    function bindDialogEvents(ui, state) {
        var di;
        var ei;

        for (di = 0; di < ui.drawButtons.length; di++) {
            ui.drawButtons[di].onClick = function () {
                onRadioClick(true, ui, state);
            };
        }
        for (ei = 0; ei < ui.eraseButtons.length; ei++) {
            ui.eraseButtons[ei].onClick = function () {
                onRadioClick(false, ui, state);
            };
        }

        ui.cbPreview.onClick = function () {
            if (ui.cbPreview.value) {
                doPreview(ui, state);
            } else {
                clearPreview(state);
            }
        };

        ui.weightDropdown.onChange = function () {
            doPreview(ui, state);
        };

        ui.colorDropdown.onChange = function () {
            doPreview(ui, state);
        };

        ui.cbClearFirst.onClick = function () {
            doPreview(ui, state);
        };

        ui.dlg.onShow = function () {
            doPreview(ui, state);
        };
    }

    function onRadioClick(isDrawBtn, ui, state) {
        var i;
        if (isDrawBtn) {
            for (i = 0; i < ui.eraseButtons.length; i++) ui.eraseButtons[i].value = false;
        } else {
            for (i = 0; i < ui.drawButtons.length; i++) ui.drawButtons[i].value = false;
        }
        doPreview(ui, state);
    }

    // =========================================
    // プレビューと確定 / Preview & Apply
    // =========================================
    function doPreview(ui, state) {
        var weight;
        var swatch;
        if (!ui.cbPreview.value) return;

        weight = parseLineWeight(getSelectedWeightText(ui));
        if (isNaN(weight) || weight < 0) return;

        swatch = getSelectedSwatch(ui);
        if (!swatch) return;

        clearPreview(state);

        try {
            app.doScript(function () {
                applyBorders(state.cells, getMode(ui), weight, ui.cbClearFirst.value, swatch);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "罫線プレビュー");
            state.previewed = true;
            app.activeDocument.recompose();
        } catch (e) {
            state.previewed = false;
        }
    }

    function clearPreview(state) {
        if (!state.previewed) return;
        try {
            app.undo();
        } catch (e) { }
        state.previewed = false;
        app.activeDocument.recompose();
    }

    function applyFinalFromDialog(ui, state) {
        var mode = getMode(ui);
        var weight = parseLineWeight(getSelectedWeightText(ui));
        var swatch = getSelectedSwatch(ui);

        if (isNaN(weight) || weight < 0) {
            clearPreview(state);
            alert(L('alertWeight'));
            return;
        }

        if (state.previewed) return;

        if (!swatch) return;

        app.doScript(function () {
            applyBorders(state.cells, mode, weight, ui.cbClearFirst.value, swatch);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "罫線の設定");
    }

    // =========================================
    // UI値の取得 / Read UI values
    // =========================================
    function getMode(ui) {
        if (ui.rbAll.value) return "all";
        if (ui.rbOuter.value) return "outer";
        if (ui.rbInnerOnly.value) return "innerOnly";
        if (ui.rbHorzOnly.value) return "horizontal";
        if (ui.rbVertOnly.value) return "vertical";
        if (ui.rbAllOff.value) return "allOff";

        return "";
    }


    function getSelectedWeightText(ui) {
        if (!ui.weightDropdown || !ui.weightDropdown.selection) return "0.1";
        return String(ui.weightDropdown.selection.text);
    }

    function getDefaultWeightIndex(weightValues) {
        var i;
        for (i = 0; i < weightValues.length; i++) {
            if (weightValues[i] === "0.1") return i;
        }
        return 0;
    }

    function getSwatchNames() {
        var names = [];
        var i;
        try {
            for (i = 0; i < app.activeDocument.swatches.length; i++) {
                names.push(String(app.activeDocument.swatches[i].name));
            }
        } catch (e) { }
        if (names.length === 0) names.push("Black");
        return names;
    }

    function getDefaultColorIndex(colorNames) {
        var preferred = ["Black", "[Black]", "ブラック", "黒"];
        var i, j;

        for (i = 0; i < preferred.length; i++) {
            for (j = 0; j < colorNames.length; j++) {
                if (colorNames[j] === preferred[i]) return j;
            }
        }
        return 0;
    }

    function getSelectedColorName(ui) {
        if (!ui.colorDropdown || !ui.colorDropdown.selection) return "";
        return String(ui.colorDropdown.selection.text);
    }

    function getSelectedSwatch(ui) {
        var swatchName = getSelectedColorName(ui);
        try {
            if (!swatchName) return null;
            return app.activeDocument.swatches.itemByName(swatchName);
        } catch (e) {
            return null;
        }
    }

    // =========================================
    // 値変換 / Value conversion
    // =========================================
    function convertMmToPoints(value) {
        return value * 2.834645669291339;
    }

    function parseLineWeight(text) {
        var value = parseFloat(text);
        if (isNaN(value)) return NaN;
        return convertMmToPoints(value);
    }

    // =========================================
    // 選択取得 / Selection
    // =========================================
    function getSelectedCellsFromApp() {
        if (app.selection.length === 0) return [];
        return getSelectedCells(app.selection[0]);
    }

    function getSelectedCells(sel) {
        var result = [];
        var i;
        try {
            if (sel.constructor.name === "Cell") {
                result.push(sel);
            } else if (sel.hasOwnProperty("cells") && sel.cells.length > 0) {
                for (i = 0; i < sel.cells.length; i++) {
                    result.push(sel.cells[i]);
                }
            } else if (sel.parent && sel.parent.constructor.name === "Cell") {
                result.push(sel.parent);
            }
        } catch (e) { }
        return result;
    }

    // =========================================
    // 罫線適用 / Apply borders
    // =========================================
    function applyBorders(cells, mode, weight, clearFirst, swatch) {
        if (cells.length === 0) return;

        var bounds = getBounds(cells);

        if (mode === "allOff") {
            applyAllOff(cells);
            return;
        }


        if (mode === "all") {
            applyAll(cells, weight, swatch);
            return;
        }

        if (mode === "outer") {
            applyOuter(cells, bounds, weight, clearFirst, swatch);
            return;
        }

        if (mode === "innerOnly") {
            applyInnerOnly(cells, bounds, weight, clearFirst, swatch);
            return;
        }

        if (mode === "horizontal") {
            applyHorizontal(cells, weight, clearFirst, swatch);
            return;
        }

        if (mode === "vertical") {
            applyVertical(cells, weight, clearFirst, swatch);
            return;
        }
    }

    function applyAllOff(cells) {
        var i, cell;
        for (i = 0; i < cells.length; i++) {
            cell = cells[i];
            // reset weights
            setCellEdges(cell, 0, 0, 0, 0);
            // reset colors to Nothing
            try {
                cell.topEdgeStrokeColor = NothingEnum.NOTHING;
                cell.bottomEdgeStrokeColor = NothingEnum.NOTHING;
                cell.leftEdgeStrokeColor = NothingEnum.NOTHING;
                cell.rightEdgeStrokeColor = NothingEnum.NOTHING;
            } catch (e) {}
        }
    }

    function applyAll(cells, weight, swatch) {
        var i;
        for (i = 0; i < cells.length; i++) {
            setCellEdges(cells[i], weight, weight, weight, weight);
            setCellEdgeColors(cells[i], swatch, swatch, swatch, swatch);
        }
    }

    function applyOuter(cells, bounds, weight, clearFirst, swatch) {
        if (clearFirst) clearAllEdges(cells);

        var i, c, edgeFlags;
        for (i = 0; i < cells.length; i++) {
            c = cells[i];
            edgeFlags = getCellEdgeFlags(c, bounds);

            if (edgeFlags.top) { c.topEdgeStrokeWeight = weight; c.topEdgeStrokeColor = swatch; }
            if (edgeFlags.bottom) { c.bottomEdgeStrokeWeight = weight; c.bottomEdgeStrokeColor = swatch; }
            if (edgeFlags.left) { c.leftEdgeStrokeWeight = weight; c.leftEdgeStrokeColor = swatch; }
            if (edgeFlags.right) { c.rightEdgeStrokeWeight = weight; c.rightEdgeStrokeColor = swatch; }
        }
    }

    function applyInnerOnly(cells, bounds, weight, clearFirst, swatch) {
        if (clearFirst) clearAllEdges(cells);

        var i, c, edgeFlags;
        for (i = 0; i < cells.length; i++) {
            c = cells[i];
            edgeFlags = getCellEdgeFlags(c, bounds);

            if (!edgeFlags.top) { c.topEdgeStrokeWeight = weight; c.topEdgeStrokeColor = swatch; }
            if (!edgeFlags.bottom) { c.bottomEdgeStrokeWeight = weight; c.bottomEdgeStrokeColor = swatch; }
            if (!edgeFlags.left) { c.leftEdgeStrokeWeight = weight; c.leftEdgeStrokeColor = swatch; }
            if (!edgeFlags.right) { c.rightEdgeStrokeWeight = weight; c.rightEdgeStrokeColor = swatch; }
        }
    }

    function applyHorizontal(cells, weight, clearFirst, swatch) {
        var i, cell;
        if (clearFirst) {
            clearAllEdges(cells);
            for (i = 0; i < cells.length; i++) {
                setCellEdges(cells[i], weight, weight, 0, 0);
                setCellEdgeColors(cells[i], swatch, swatch, null, null);
            }
            return;
        }

        for (i = 0; i < cells.length; i++) {
            cell = cells[i];
            cell.topEdgeStrokeWeight = weight;
            cell.bottomEdgeStrokeWeight = weight;
            setCellEdgeColors(cell, swatch, swatch, null, null);
        }
    }

    function applyVertical(cells, weight, clearFirst, swatch) {
        var i, cell;
        if (clearFirst) {
            clearAllEdges(cells);
            for (i = 0; i < cells.length; i++) {
                setCellEdges(cells[i], 0, 0, weight, weight);
                setCellEdgeColors(cells[i], null, null, swatch, swatch);
            }
            return;
        }

        for (i = 0; i < cells.length; i++) {
            cell = cells[i];
            cell.leftEdgeStrokeWeight = weight;
            cell.rightEdgeStrokeWeight = weight;
            setCellEdgeColors(cell, null, null, swatch, swatch);
        }
    }

    function clearAllEdges(cells) {
        var i;
        for (i = 0; i < cells.length; i++) {
            setCellEdges(cells[i], 0, 0, 0, 0);
        }
    }

    function getCellEdgeFlags(cell, bounds) {
        var rowIndex = cell.parentRow.index;
        var colIndex = cell.parentColumn.index;

        return {
            top: rowIndex === bounds.minRow,
            bottom: rowIndex === bounds.maxRow,
            left: colIndex === bounds.minCol,
            right: colIndex === bounds.maxCol
        };
    }

    function getBounds(cells) {
        var minRow = 999999;
        var maxRow = -1;
        var minCol = 999999;
        var maxCol = -1;
        var i, c, r, col;

        for (i = 0; i < cells.length; i++) {
            c = cells[i];
            r = c.parentRow.index;
            col = c.parentColumn.index;

            if (r < minRow) minRow = r;
            if (r > maxRow) maxRow = r;
            if (col < minCol) minCol = col;
            if (col > maxCol) maxCol = col;
        }

        return {
            minRow: minRow,
            maxRow: maxRow,
            minCol: minCol,
            maxCol: maxCol
        };
    }

    function setCellEdges(cell, top, bottom, left, right) {
        cell.topEdgeStrokeWeight = top;
        cell.bottomEdgeStrokeWeight = bottom;
        cell.leftEdgeStrokeWeight = left;
        cell.rightEdgeStrokeWeight = right;
    }

    function setCellEdgeColors(cell, top, bottom, left, right) {
        if (top != null) cell.topEdgeStrokeColor = top;
        if (bottom != null) cell.bottomEdgeStrokeColor = bottom;
        if (left != null) cell.leftEdgeStrokeColor = left;
        if (right != null) cell.rightEdgeStrokeColor = right;
    }
})();
