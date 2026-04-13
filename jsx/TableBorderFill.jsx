#target "InDesign"

/*
 * =========================================
 * 罫線の調整 / Borders
 * =========================================
 *
 * 【概要 / Overview】
 * 選択中の表セルに対して、罫線をインタラクティブに調整できるツールです。
 * 複数の適用範囲（すべて／境界線のみ／内部のみ／水平線のみ／垂直線のみ／上下端／左右端／左右の罫線を消去／すべて消去）に対応し、
 * 線幅・カラー・濃淡を一括で設定できます。
 *
 * This tool lets you interactively adjust borders for selected table cells.
 * It supports multiple scope modes (All, Outer Borders, Inner Borders, Horizontal Borders,
 * Vertical Borders, Top & Bottom, Left & Right, Remove Side Borders, Clear All)
 * and lets you set stroke weight, color, and tint together.
 *
 * 【挙動 / Behavior】
 * すべての適用範囲は、選択セルの外接矩形（bounds）を基準に判定されます。
 * 結合セル（rowSpan / columnSpan）は考慮されます。
 *
 * All scope modes are evaluated against the bounding rectangle (bounds) of the selection.
 * Merged cells (rowSpan / columnSpan) are respected.
 *
 * 【注意 / Notes】
 * L字や飛び飛びなどの非矩形選択では、見た目どおりの結果にならない場合があります（矩形基準のため）。
 *
 * Non-rectangular selections (L-shaped, disjoint) may not match visual expectations
 * because processing is strictly bounds-based.
 *
 * 【UI構成 / UI Structure】
 * 左カラム：適用範囲 + 適用オプション（適用前に消去）
 * 右カラム：スタイル（線幅・カラー・濃淡）
 *
 * Left column: Border scope + apply options (Clear Before Apply)
 * Right column: Style (stroke weight, color, and tint)
 *
 * 【特徴 / Features】
 * - 常時プレビューによるインタラクティブ操作
 * - ［プレビュー／標準モード］ボタンによる画面表示切替
 * - ダイアログを開く前の画面モードに復帰
 * - スウォッチ連動のカラー適用
 * - 線カラーの濃淡指定
 * - 結合セルを考慮した境界処理（矩形基準内）
 * - 適用範囲ボタンの Option+クリックで［適用前に消去］を切り替え
 *
 * - Always-on interactive preview
 * - Screen mode toggle (Preview / Standard)
 * - Restores the previous screen mode when the dialog closes
 * - Swatch-based color application
 * - Stroke tint control
 * - Handles merged cells within bounds-based logic
 * - Option+Click on scope buttons toggles "Clear Before Apply"
 *
 * =========================================
 */

var SCRIPT_VERSION = "v1.5.1";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "罫線の調整", en: "Borders" },

    modePanel: { ja: "適用範囲", en: "Border Scope" },
    stylePanel: { ja: "スタイル", en: "Style" },
    clearFirst: { ja: "適用前に消去", en: "Clear Before Apply" },

    all: { ja: "すべて", en: "All" },
    outer: { ja: "境界線のみ", en: "Outer Borders" },
    inner: { ja: "内部のみ", en: "Inner Borders" },
    horizontal: { ja: "水平線のみ", en: "Horizontal Borders" },
    vertical: { ja: "垂直線のみ", en: "Vertical Borders" },
    headerRow: { ja: "上下端", en: "Top & Bottom" },
    headerColumn: { ja: "左右端", en: "Left & Right" },
    clearLeftRight: { ja: "左右の罫線を消去", en: "Remove Side Borders" },
    allOff: { ja: "すべて消去", en: "Clear All" },

    lineWidthLabel: { ja: "線幅：", en: "Stroke Weight:" },
    lineWidthUnitMm: "mm",
    lineWidthUnitPt: "pt",
    lineWidthUnitCm: "cm",
    lineWidthUnitIn: "in",
    lineWidthUnitPica: "pica",
    lineWidthUnitQ: "Q",

    lineWidthPresetNone: { ja: "なし", en: "None" },
    lineWidthPreset01: "0.1",
    lineWidthPreset02: "0.2",
    lineWidthPreset025: "0.25",
    lineWidthPreset035: "0.35",
    lineWidthPreset05: "0.5",

    colorLabel: { ja: "カラー：", en: "Color:" },
    tintLabel: { ja: "濃淡：", en: "Tint:" },

    swatchBlack: { ja: "黒", en: "Black" },
    swatchPaper: { ja: "紙色", en: "Paper" },
    swatchNone: { ja: "なし", en: "None" },

    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" },

    alertSelect: { ja: "表のセルを選択してください。", en: "Please select table cells." },
    alertWeight: { ja: "線幅には0以上の数値を入力してください。", en: "Enter a stroke weight of 0 or greater." },

    undoPreview: { ja: "罫線プレビュー", en: "Border Preview" },
    undoApply: { ja: "罫線の設定", en: "Apply Borders" },

    tipAll: { ja: "ショートカット: A", en: "Shortcut: A" },
    tipOuter: { ja: "ショートカット: E", en: "Shortcut: E" },
    tipInner: { ja: "ショートカット: I", en: "Shortcut: I" },
    tipHorizontal: { ja: "ショートカット: H", en: "Shortcut: H" },
    tipVertical: { ja: "ショートカット: V", en: "Shortcut: V" },
    tipHeaderRow: { ja: "ショートカット: U", en: "Shortcut: U" },
    tipHeaderColumn: { ja: "ショートカット: L", en: "Shortcut: L" },
    tipClearLeftRight: { ja: "ショートカット: R", en: "Shortcut: R" },
    tipAllOff: { ja: "ショートカット: C", en: "Shortcut: C" }
};

function L(key) {
    var v = LABELS[key];
    if (v == null) return key;
    if (typeof v === "string") return v;
    return v[lang] || v.en || key;
}

(function () {
    var cells = normalizeSelectedCells(getSelectedCellsFromApp());
    if (cells.length === 0) {
        alert(L('alertSelect'));
        return;
    }

    var fullTable = isFullTableSelected(cells);

    var state = {
        cells: cells,
        previewed: false,
        initialScreenMode: getCurrentScreenModeSafe(),
        previewToggleCount: 0,
        isFullTable: fullTable
    };

    app.selection = NothingEnum.NOTHING;

    var ui = buildDialog();
    bindDialogEvents(ui, state);

    var result = ui.dlg.show();
    if (result != 1) {
        clearPreview(state);
        restoreScreenModeFromState(state);
        return;
    }

    applyFinalFromDialog(ui, state);
    restoreScreenModeFromState(state);

    function normalizeSelectedCells(cells) {
        var result = [];
        var seen = {};
        var i, key;

        if (!cells || cells.length === 0) return result;

        for (i = 0; i < cells.length; i++) {
            key = getCellKey(cells[i]);
            if (!seen[key]) {
                seen[key] = true;
                result.push(cells[i]);
            }
        }

        return result;
    }

    function getCellKey(cell) {
        var range;
        try {
            range = getCellRange(cell);
            return [
                cell.parent && cell.parent.id,
                range.startRow,
                range.endRow,
                range.startCol,
                range.endCol
            ].join(":");
        } catch (e) {
            return String(cell);
        }
    }

    // =========================================
    // UI構築 / Build UI
    // =========================================
    function buildDialog() {
        var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = "fill";

        var settingsColumns = dlg.add("group");
        settingsColumns.orientation = "row";
        settingsColumns.alignChildren = ["fill", "top"];
        settingsColumns.alignment = ["fill", "top"];
        settingsColumns.spacing = 15;

        var leftColumn = settingsColumns.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];
        leftColumn.alignment = ["fill", "top"];
        leftColumn.spacing = 10;

        var panelMode = leftColumn.add("panel", undefined, L('modePanel'));
        panelMode.orientation = "column";
        panelMode.alignChildren = "left";
        panelMode.alignment = ["fill", "top"];
        panelMode.margins = [15, 20, 15, 10];

        var rbAll = panelMode.add("radiobutton", undefined, L('all'));
        rbAll.helpTip = L('tipAll');
        var rbOuter = panelMode.add("radiobutton", undefined, L('outer'));
        rbOuter.helpTip = L('tipOuter');
        var rbInnerOnly = panelMode.add("radiobutton", undefined, L('inner'));
        rbInnerOnly.helpTip = L('tipInner');
        var rbHorzOnly = panelMode.add("radiobutton", undefined, L('horizontal'));
        rbHorzOnly.helpTip = L('tipHorizontal');
        var rbVertOnly = panelMode.add("radiobutton", undefined, L('vertical'));
        rbVertOnly.helpTip = L('tipVertical');
        var rbHeaderRow = panelMode.add("radiobutton", undefined, L('headerRow'));
        rbHeaderRow.helpTip = L('tipHeaderRow');
        var rbHeaderColumn = panelMode.add("radiobutton", undefined, L('headerColumn'));
        rbHeaderColumn.helpTip = L('tipHeaderColumn');
        var rbClearLeftRight = panelMode.add("radiobutton", undefined, L('clearLeftRight'));
        rbClearLeftRight.helpTip = L('tipClearLeftRight');
        var rbAllOff = panelMode.add("radiobutton", undefined, L('allOff'));
        rbAllOff.helpTip = L('tipAllOff');

        // 適用オプションパネル / Apply options panel
        var panelDrawingOptions = leftColumn.add("group");
        panelDrawingOptions.orientation = "column";
        panelDrawingOptions.alignChildren = "left";
        panelDrawingOptions.alignment = ["fill", "top"];
        panelDrawingOptions.margins = [15, 5, 0, 0];

        var cbClearFirst = panelDrawingOptions.add("checkbox", undefined, L('clearFirst'));
        cbClearFirst.value = true;

        var rightColumn = settingsColumns.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];
        rightColumn.alignment = ["fill", "top"];
        rightColumn.spacing = 10;

        var panelStyle = rightColumn.add("panel", undefined, L('stylePanel'));
        panelStyle.orientation = "column";
        panelStyle.alignChildren = ["fill", "top"];
        panelStyle.alignment = ["fill", "top"];
        panelStyle.margins = [15, 20, 15, 10];

        var panelWeight = panelStyle.add("group");
        panelWeight.orientation = "column";
        panelWeight.alignChildren = ["fill", "top"];
        panelWeight.alignment = ["fill", "top"];

        var weightRow = panelWeight.add("group");
        weightRow.orientation = "row";
        weightRow.alignChildren = ["left", "center"];
        weightRow.alignment = ["left", "center"];
        weightRow.spacing = 8;

        weightRow.add("statictext", undefined, L('lineWidthLabel'));
        var weightInput = weightRow.add("edittext", undefined, getDefaultLineWidthText());
        weightInput.characters = 5;

        weightRow.add("statictext", undefined, getCurrentLineWidthUnitLabel());

        var weightPresetGroup = panelWeight.add("group");
        weightPresetGroup.orientation = "column";
        weightPresetGroup.alignChildren = ["left", "center"];
        weightPresetGroup.alignment = ["left", "top"];
        weightPresetGroup.spacing = 4;
        weightPresetGroup.margins = [15, 5, 15, 10];

        var rbWeightNone = weightPresetGroup.add("radiobutton", undefined, L('lineWidthPresetNone'));
        var rbWeight01 = weightPresetGroup.add("radiobutton", undefined, L('lineWidthPreset01'));
        var rbWeight02 = weightPresetGroup.add("radiobutton", undefined, L('lineWidthPreset02'));
        var rbWeight025 = weightPresetGroup.add("radiobutton", undefined, L('lineWidthPreset025'));
        var rbWeight035 = weightPresetGroup.add("radiobutton", undefined, L('lineWidthPreset035'));
        var rbWeight05 = weightPresetGroup.add("radiobutton", undefined, L('lineWidthPreset05'));

        syncWeightPresetFromTextValue({
            rbWeightNone: rbWeightNone,
            rbWeight01: rbWeight01,
            rbWeight02: rbWeight02,
            rbWeight025: rbWeight025,
            rbWeight035: rbWeight035,
            rbWeight05: rbWeight05
        }, getDefaultLineWidthText());

        var panelColor = panelStyle.add("group");
        panelColor.orientation = "column";
        panelColor.alignChildren = ["left", "top"];
        panelColor.alignment = ["fill", "top"];

        var colorLabelRow = panelColor.add("group");
        colorLabelRow.orientation = "row";
        colorLabelRow.alignChildren = ["left", "center"];
        colorLabelRow.add("statictext", undefined, L('colorLabel'));
        var colorPreviewBox = createSwatchPreviewBox(colorLabelRow);

        var swatchEntries = getSwatchEntries();
        var colorDropdown = createSwatchDropdown(panelColor, swatchEntries, getDefaultColorIndex(swatchEntries));

        var tintGroup = panelColor.add("group");
        tintGroup.orientation = "column";
        tintGroup.alignChildren = "left";
        tintGroup.margins = [0, 10, 0, 10];
        var tintLabelRow = tintGroup.add("group");
        tintLabelRow.orientation = "row";
        tintLabelRow.alignChildren = ["left", "center"];
        tintLabelRow.add("statictext", undefined, L('tintLabel'));
        var tintText = tintLabelRow.add("edittext", undefined, "100");
        tintText.preferredSize.width = 50;
        tintLabelRow.add("statictext", undefined, "%");
        var tintSlider = tintGroup.add("slider", undefined, 100, 0, 100);
        tintSlider.preferredSize.width = 150;
        tintGroup.enabled = true;


        rbAll.value = true;

        if (!state.isFullTable) {
            rbHeaderRow.enabled = false;
            rbHeaderColumn.enabled = false;
        }

        // カラープレビューの初期表示
        updateSwatchPreview(colorPreviewBox, colorDropdown, dlg);

        var btnArea = dlg.add("group");
        btnArea.orientation = "row";
        btnArea.alignChildren = ["fill", "fill"];
        btnArea.alignment = ["fill", "bottom"];
        btnArea.margins = [0, 8, 0, 0];

        var btnLeftGroup = btnArea.add("group");
        btnLeftGroup.orientation = "column";
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.alignment = ["left", "fill"];

        var btnPreviewToggle = btnLeftGroup.add("button", undefined, getPreviewToggleButtonLabel());
        btnPreviewToggle.alignment = ["left", "center"];

        var btnCenterGroup = btnArea.add("group");
        btnCenterGroup.orientation = "column";
        btnCenterGroup.alignChildren = ["fill", "fill"];
        btnCenterGroup.alignment = ["fill", "fill"];

        var spacer = btnCenterGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 40;

        var btnRightGroup = btnArea.add("group");
        btnRightGroup.orientation = "column";
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.alignment = ["right", "fill"];

        var buttonRow = btnRightGroup.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignChildren = ["right", "center"];
        buttonRow.alignment = ["right", "center"];
        buttonRow.spacing = 8;

        var btnCancel = buttonRow.add("button", undefined, L('cancel'), { name: "cancel" });
        var btnOk = buttonRow.add("button", undefined, L('ok'), { name: "ok" });

        dlg.layout.layout(true);
        dlg.layout.resize();
        dlg.onResizing = dlg.onResize = function () { this.layout.resize(); };
        return {
            dlg: dlg,
            rbAll: rbAll,
            rbOuter: rbOuter,
            rbInnerOnly: rbInnerOnly,
            rbHorzOnly: rbHorzOnly,
            rbVertOnly: rbVertOnly,
            rbHeaderRow: rbHeaderRow,
            rbHeaderColumn: rbHeaderColumn,
            rbClearLeftRight: rbClearLeftRight,
            rbAllOff: rbAllOff,
            weightInput: weightInput,
            rbWeightNone: rbWeightNone,
            rbWeight01: rbWeight01,
            rbWeight02: rbWeight02,
            rbWeight025: rbWeight025,
            rbWeight035: rbWeight035,
            rbWeight05: rbWeight05,
            colorDropdown: colorDropdown,
            colorPreviewBox: colorPreviewBox,
            tintSlider: tintSlider,
            tintText: tintText,
            tintGroup: tintGroup,

            cbClearFirst: cbClearFirst,
            btnCancel: btnCancel,
            btnOk: btnOk,
            drawButtons: [rbAll, rbOuter, rbInnerOnly, rbHorzOnly, rbVertOnly, rbHeaderRow, rbHeaderColumn, rbClearLeftRight, rbAllOff],
            btnPreviewToggle: btnPreviewToggle
        };
    }

    // =========================================
    // イベント / Events
    // =========================================
    function bindDialogEvents(ui, state) {
        var di;

        for (di = 0; di < ui.drawButtons.length; di++) {
            ui.drawButtons[di].onClick = function (event) {
                onRadioClick(ui, state, event);
            };
            if (typeof ui.drawButtons[di].helpTip !== "string") ui.drawButtons[di].helpTip = "";
            ui.drawButtons[di].helpTip += "\nOption+Click: Toggle Clear Before Apply";
        }

        ui.rbWeightNone.onClick = function () { applyWeightPreset(ui, "0", state); };
        ui.rbWeight01.onClick = function () { applyWeightPreset(ui, "0.1", state); };
        ui.rbWeight02.onClick = function () { applyWeightPreset(ui, "0.2", state); };
        ui.rbWeight025.onClick = function () { applyWeightPreset(ui, "0.25", state); };
        ui.rbWeight035.onClick = function () { applyWeightPreset(ui, "0.35", state); };
        ui.rbWeight05.onClick = function () { applyWeightPreset(ui, "0.5", state); };


        ui.weightInput.onChange = function () {
            syncWeightPresetFromInput(ui);
            doPreview(ui, state);
        };

        ui.colorDropdown.onChange = function () {
            updateSwatchPreview(ui.colorPreviewBox, ui.colorDropdown, ui.dlg);
            updateBorderTintEnabled(ui);
            if (ui.colorDropdown.selection && isNoneSwatchName(String(ui.colorDropdown.selection._swatchName || ""))) {
                ui.weightInput.text = "0";
                syncWeightPresetFromInput(ui);
            }
            doPreview(ui, state);
        };

        ui.cbClearFirst.onClick = function () {
            doPreview(ui, state);
        };

        changeValueByArrowKey(ui.weightInput, false, function () {
            syncWeightPresetFromInput(ui);
            doPreview(ui, state);
        });

        ui.tintSlider.onChanging = function () {
            var v = adjustFillTintValue(this.value);
            this.value = v;
            ui.tintText.text = String(v);
            doPreview(ui, state);
        };
        ui.tintText.onChange = function () {
            var v = adjustFillTintValue(parseFloat(this.text));
            this.text = String(v);
            ui.tintSlider.value = v;
            doPreview(ui, state);
        };

        changeValueByArrowKey(ui.tintText, false, function () {
            var v = adjustFillTintValue(parseFloat(ui.tintText.text));
            ui.tintText.text = String(v);
            ui.tintSlider.value = v;
            doPreview(ui, state);
        });

        addDrawingOptionKeyHandler(ui.dlg, ui, state);
        addModeShortcutKeyHandler(ui.dlg, ui, state);

        ui.dlg.onShow = function () {
            updateBorderTintEnabled(ui);
            doPreview(ui, state);
        };

        ui.btnPreviewToggle.onClick = function () {
            togglePreviewScreenMode();
            state.previewToggleCount++;
            updatePreviewToggleButtonLabel(ui.btnPreviewToggle);
        };
    }

    function getCurrentScreenModeSafe() {
        try {
            return app.activeWindow ? app.activeWindow.screenMode : null;
        } catch (e) {
            return null;
        }
    }

    function setScreenModeSafe(mode) {
        try {
            if (app.activeWindow && mode != null) {
                app.activeWindow.screenMode = mode;
            }
        } catch (e) { }
    }

    function restoreScreenModeFromState(state) {
        if (!state) return;
        if ((state.previewToggleCount % 2) === 0) return;
        setScreenModeSafe(state.initialScreenMode);
    }

    function getPreviewToggleButtonLabel() {
        return isPreviewScreenMode() ? "プレビュー" : "標準モード";
    }

    function updatePreviewToggleButtonLabel(button) {
        if (!button) return;
        button.text = getPreviewToggleButtonLabel();
    }

    function isPreviewScreenMode() {
        try {
            return app.activeWindow && app.activeWindow.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE;
        } catch (e) {
            return false;
        }
    }

    function togglePreviewScreenMode() {
        try {
            var w = app.activeWindow;
            if (!w) return;
            if (w.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE) {
                w.screenMode = ScreenModeOptions.PREVIEW_OFF;
            } else {
                w.screenMode = ScreenModeOptions.PREVIEW_TO_PAGE;
            }
        } catch (e) { }
    }

    function onRadioClick(ui, state, event) {
        var ks = ScriptUI.environment.keyboardState;
        var isAlt = !!(ks && ks.altKey);
        if (isAlt && ui && ui.cbClearFirst) {
            ui.cbClearFirst.value = !ui.cbClearFirst.value;
        }
        doPreview(ui, state);
    }

    function applyWeightPreset(ui, value, state) {
        if (!ui || !ui.weightInput) return;

        ui.weightInput.text = String(value);
        syncWeightPresetFromInput(ui);
        doPreview(ui, state);
    }

    function syncWeightPresetFromInput(ui) {
        if (!ui) return;
        syncWeightPresetFromTextValue(ui, getSelectedWeightText(ui));
    }

    function changeValueByArrowKey(editText, allowNegative, onAfterChange) {
        editText.addEventListener("keydown", function (event) {
            if (applyArrowStepToEditText(editText, allowNegative, event, onAfterChange)) {
                event.preventDefault();
            }
        });
    }

    function applyArrowStepToEditText(editText, allowNegative, event, onAfterChange) {
        var value = Number(editText.text);
        var keyboard = ScriptUI.environment.keyboardState;
        var keyName = normalizeArrowKeyName(event ? event.keyName : "");
        var isShift = !!(keyboard.shiftKey || (event && event.shiftKey));

        if (isNaN(value)) return false;
        if (keyName !== "Up" && keyName !== "Down") return false;

        if (isShift) {
            if (keyName == "Up") {
                value = Math.floor(value) + 1;
            } else {
                value = Math.ceil(value) - 1;
            }
        } else {
            if (keyName == "Up") {
                value += 0.1;
            } else {
                value -= 0.1;
            }
        }

        if (!allowNegative && value < 0) value = 0;

        value = Math.round(value * 10) / 10;
        editText.text = String(value.toFixed(1).replace(/\.0$/, ""));

        if (typeof onAfterChange === "function") onAfterChange();
        return true;
    }

    function normalizeArrowKeyName(keyName) {
        keyName = String(keyName);
        if (keyName === "Up" || keyName === "Down") return keyName;
        if (keyName === "UpArrow") return "Up";
        if (keyName === "DownArrow") return "Down";
        if (keyName === "PageUp") return "Up";
        if (keyName === "PageDown") return "Down";
        return keyName;
    }

    function addDrawingOptionKeyHandler(dialog, ui, state) {
        dialog.addEventListener("keydown", function (event) {
            var ks = ScriptUI.environment.keyboardState;
            if (ks.metaKey || ks.ctrlKey || ks.altKey) return;

            if (event.keyName == "M") {
                ui.cbClearFirst.value = !ui.cbClearFirst.value;
                event.preventDefault();
                doPreview(ui, state);
            }
        });
    }

    function addModeShortcutKeyHandler(dialog, ui, state) {
        dialog.addEventListener("keydown", function (event) {
            var ks = ScriptUI.environment.keyboardState;
            if (ks.metaKey || ks.ctrlKey || ks.altKey) return;

            var keyName = String(event.keyName);
            var handled = true;

            if (keyName == "A") {
                ui.rbAll.value = true;
            } else if (keyName == "E") {
                ui.rbOuter.value = true;
            } else if (keyName == "I") {
                ui.rbInnerOnly.value = true;
            } else if (keyName == "H") {
                ui.rbHorzOnly.value = true;
            } else if (keyName == "V") {
                ui.rbVertOnly.value = true;
            } else if (keyName == "U" && ui.rbHeaderRow.enabled) {
                ui.rbHeaderRow.value = true;
            } else if (keyName == "L" && ui.rbHeaderColumn.enabled) {
                ui.rbHeaderColumn.value = true;
            } else if (keyName == "R") {
                ui.rbClearLeftRight.value = true;
            } else if (keyName == "C") {
                ui.rbAllOff.value = true;
            } else {
                handled = false;
            }

            if (handled) {
                event.preventDefault();
                doPreview(ui, state);
            }
        });
    }

    // =========================================
    // プレビューと確定 / Preview & Apply
    // =========================================
    function doPreview(ui, state) {
        clearPreview(state);

        var weight = parseLineWeight(getSelectedWeightText(ui));
        if (!isValidLineWeight(weight)) {
            return;
        }

        var swatch = getSelectedSwatch(ui);
        var tint = getBorderTintFromUI(ui);
        if (!swatch) {
            return;
        }

        try {
            app.doScript(function () {
                applyBorders(state.cells, getMode(ui), weight, ui.cbClearFirst.value, swatch, tint);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, L('undoPreview'));
            state.previewed = true;
            app.activeDocument.recompose();
        } catch (e) {
            state.previewed = false;
            try { $.writeln("[TableBorderFill] Preview error: " + e); } catch (_) { }
        }
    }

    function clearPreview(state) {
        if (!state.previewed) return;
        try {
            app.undo();
        } catch (e) {
            try { $.writeln("[TableBorderFill] Undo preview error: " + e); } catch (_) { }
        }
        state.previewed = false;
        app.activeDocument.recompose();
    }

    function applyFinalFromDialog(ui, state) {
        if (state.previewed) {
            clearPreview(state);
        }

        var mode = getMode(ui);
        var weight = parseLineWeight(getSelectedWeightText(ui));
        var swatch = getSelectedSwatch(ui);
        var tint = getBorderTintFromUI(ui);

        if (!isValidLineWeight(weight)) {
            alert(L('alertWeight'));
            return;
        }

        if (!swatch) return;

        try {
            app.doScript(function () {
                applyBorders(state.cells, mode, weight, ui.cbClearFirst.value, swatch, tint);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, L('undoApply'));
        } catch (e) {
            try { $.writeln("[TableBorderFill] Apply error: " + e); } catch (_) { }
            throw e;
        }
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
        if (ui.rbHeaderRow.value) return "headerRow";
        if (ui.rbHeaderColumn.value) return "headerColumn";
        if (ui.rbClearLeftRight.value) return "clearLeftRight";
        if (ui.rbAllOff.value) return "allOff";

        return "";
    }

    function getSelectedWeightText(ui) {
        var text = "";

        if (ui.weightInput && ui.weightInput.text != null) {
            text = String(ui.weightInput.text).replace(/^\s+|\s+$/g, "");
            if (text !== "") return text;
        }

        return getDefaultLineWidthText();
    }

    function getSwatchEntries() {
        var entries = [];
        var i;
        var swatch;
        var actualName;
        try {
            for (i = 0; i < app.activeDocument.swatches.length; i++) {
                swatch = app.activeDocument.swatches[i];
                actualName = String(swatch.name);
                if (isRegistrationSwatchName(actualName)) continue;
                entries.push({
                    displayName: getDisplaySwatchName(actualName),
                    actualName: actualName
                });
            }
        } catch (e) { }
        if (entries.length === 0) {
            entries.push({ displayName: L('swatchBlack'), actualName: 'Black' });
        }
        return entries;
    }

    function getDisplaySwatchName(name) {
        var kind = getLocalizedSwatchDisplayKind(name);
        if (kind === "none") return L('swatchNone');
        if (kind === "black") return L('swatchBlack');
        if (kind === "paper") return L('swatchPaper');
        return String(name);
    }

    function getLocalizedSwatchDisplayKind(name) {
        if (isNoneSwatchName(name)) return "none";
        if (isBlackSwatchName(name)) return "black";
        if (isPaperSwatchName(name)) return "paper";
        return "";
    }


    function normalizeSwatchName(name) {
        if (name == null) return "";
        var n = String(name);

        // remove brackets like [Black]
        n = n.replace(/^\[|\]$/g, "");

        // trim and lowercase for comparison
        n = n.replace(/^\s+|\s+$/g, "").toLowerCase();

        return n;
    }

    function isRegistrationSwatchName(name) {
        var n = normalizeSwatchName(name);
        return n === "registration" || n === "レジストレーション";
    }

    function getDefaultColorIndex(swatchEntries) {
        var i;
        for (i = 0; i < swatchEntries.length; i++) {
            if (isBlackSwatchName(String(swatchEntries[i].actualName))) return i;
        }
        return 0;
    }

    function getSelectedColorName(ui) {
        return getSelectedSwatchNameFromDropdown(ui ? ui.colorDropdown : null);
    }

    function getSelectedSwatchNameFromDropdown(dropdown) {
        if (!dropdown || !dropdown.selection) return "";
        if (dropdown.selection._swatchName != null) return String(dropdown.selection._swatchName);
        return String(dropdown.selection.text);
    }

    function getSelectedSwatch(ui) {
        return getSwatchByName(getSelectedColorName(ui));
    }

    function getBorderTintFromUI(ui) {
        if (!ui || !ui.tintText) return 100;
        return clampTintValue(parseFloat(ui.tintText.text));
    }
    // =========================================
    // スウォッチUIとプレビュー / Swatch UI and preview helpers
    // =========================================
    function createSwatchPreviewBox(parent) {
        var previewBox = parent.add("group");
        previewBox.preferredSize = [18, 18];
        previewBox.minimumSize = [18, 18];
        previewBox.maximumSize = [18, 18];
        return previewBox;
    }

    function createSwatchDropdown(parent, swatchEntries, defaultIndex) {
        var dropdown;
        var displayNames = [];
        var i;

        swatchEntries = swatchEntries || [];
        for (i = 0; i < swatchEntries.length; i++) {
            displayNames.push(String(swatchEntries[i].displayName));
        }

        dropdown = parent.add("dropdownlist", undefined, displayNames);
        dropdown.minimumSize.height = 22;
        dropdown.minimumSize.width = 130;
        dropdown.preferredSize.width = 130;

        for (i = 0; i < dropdown.items.length && i < swatchEntries.length; i++) {
            dropdown.items[i]._swatchName = String(swatchEntries[i].actualName);
        }

        if (dropdown.items.length > 0) {
            if (typeof defaultIndex === "number" && defaultIndex >= 0 && defaultIndex < dropdown.items.length) {
                dropdown.selection = defaultIndex;
            } else {
                dropdown.selection = 0;
            }
        }

        return dropdown;
    }

    function updateSwatchPreview(previewBox, dropdown, dlg) {
        var swatch;
        if (!previewBox || !dropdown || !dropdown.selection) return;

        swatch = getSwatchByName(getSelectedSwatchNameFromDropdown(dropdown));
        if (!swatch) return;

        previewBox.graphics.backgroundColor = previewBox.graphics.newBrush(
            previewBox.graphics.BrushType.SOLID_COLOR,
            getSwatchPreviewRGBAFromSwatch(swatch)
        );
        if (dlg) dlg.update();
    }

    function getSwatchByName(swatchName) {
        var swatch;
        try {
            if (!swatchName) return null;
            swatch = app.activeDocument.swatches.itemByName(String(swatchName));
            if (!swatch || !swatch.isValid) return null;
            return swatch;
        } catch (e) {
            return null;
        }
    }

    function getSwatchPreviewRGBAFromSwatch(swatch) {
        var rgb = convertSwatchToPreviewRGB(swatch);
        return [rgb[0], rgb[1], rgb[2], 1];
    }

    function convertSwatchToPreviewRGB(swatch) {
        var vals;
        var c, m, y, k;

        if (!swatch) return [0.5, 0.5, 0.5];

        if (isWhitePreviewSwatch(swatch)) return [1, 1, 1];
        if (isBlackPreviewSwatch(swatch)) return [0, 0, 0];

        try {
            if (swatch.hasOwnProperty("colorValue")) {
                vals = swatch.colorValue;
                if (swatch.space === ColorSpace.RGB) {
                    return [vals[0] / 255, vals[1] / 255, vals[2] / 255];
                }
                if (swatch.space === ColorSpace.CMYK) {
                    c = vals[0] / 100;
                    m = vals[1] / 100;
                    y = vals[2] / 100;
                    k = vals[3] / 100;
                    return [(1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k)];
                }
            }
        } catch (e) { }

        return [0.5, 0.5, 0.5];
    }

    function isWhitePreviewSwatch(swatch) {
        var kind = getPreviewSwatchKind(swatch);
        return kind === "none" || kind === "paper";
    }

    function isBlackPreviewSwatch(swatch) {
        var kind = getPreviewSwatchKind(swatch);
        return kind === "registration" || kind === "black";
    }

    function getPreviewSwatchKind(swatch) {
        var name = swatch && swatch.name != null ? String(swatch.name) : "";
        if (isNoneSwatchName(name)) return "none";
        if (isPaperSwatchName(name)) return "paper";
        if (isRegistrationSwatchName(name)) return "registration";
        if (isBlackSwatchName(name)) return "black";
        return "";
    }

    function getCurrentMeasurementUnit() {
        try {
            return app.activeDocument.viewPreferences.strokeMeasurementUnits;
        } catch (e) {
            return MeasurementUnits.POINTS;
        }
    }

    function getCurrentLineWidthUnitLabel() {
        switch (getCurrentMeasurementUnit()) {
            case MeasurementUnits.MILLIMETERS:
                return L('lineWidthUnitMm');
            case MeasurementUnits.POINTS:
                return L('lineWidthUnitPt');
            case MeasurementUnits.CENTIMETERS:
                return L('lineWidthUnitCm');
            case MeasurementUnits.INCHES:
                return L('lineWidthUnitIn');
            case MeasurementUnits.PICAS:
                return L('lineWidthUnitPica');
            case MeasurementUnits.Q:
                return L('lineWidthUnitQ');
            default:
                return L('lineWidthUnitPt');
        }
    }

    function getDefaultLineWidthText() {
        return getCurrentMeasurementUnit() === MeasurementUnits.POINTS ? "0.25" : "0.1";
    }

    function getCurrentLineWidthUnitSuffix() {
        switch (getCurrentMeasurementUnit()) {
            case MeasurementUnits.MILLIMETERS:
                return "mm";
            case MeasurementUnits.POINTS:
                return "pt";
            case MeasurementUnits.CENTIMETERS:
                return "cm";
            case MeasurementUnits.INCHES:
                return "in";
            case MeasurementUnits.PICAS:
                return "p";
            case MeasurementUnits.Q:
                return "q";
            default:
                return "pt";
        }
    }

    function syncWeightPresetFromTextValue(target, textValue) {
        var value = parseFloat(textValue);
        if (!target || isNaN(value)) return;

        if (target.rbWeightNone) target.rbWeightNone.value = (value === 0);
        if (target.rbWeight01) target.rbWeight01.value = (value === 0.1);
        if (target.rbWeight02) target.rbWeight02.value = (value === 0.2);
        if (target.rbWeight025) target.rbWeight025.value = (value === 0.25);
        if (target.rbWeight035) target.rbWeight035.value = (value === 0.35);
        if (target.rbWeight05) target.rbWeight05.value = (value === 0.5);
    }

    function updateBorderTintEnabled(ui) {
        var swatchName;
        var enabled;

        if (!ui || !ui.colorDropdown || !ui.tintGroup) return;

        swatchName = getSelectedSwatchNameFromDropdown(ui.colorDropdown);
        enabled = !(isNoneSwatchName(swatchName) || isPaperSwatchName(swatchName));

        ui.tintGroup.enabled = enabled;
    }

    // =========================================
    // 値変換 / Value conversion
    // =========================================
    function parseLineWeight(text) {
        var value = parseFloat(text);
        var suffix;

        if (isNaN(value)) return NaN;
        if (value === 0) return 0;

        suffix = getCurrentLineWidthUnitSuffix();
        if (suffix) {
            return String(value) + suffix;
        }

        return value;
    }

    // =========================================
    // 線幅バリデーション / Line weight validation
    function extractLineWeightNumber(weight) {
        if (typeof weight === "number") return weight;
        if (typeof weight === "string") return parseFloat(weight);
        return NaN;
    }

    function isValidLineWeight(weight) {
        var numericWeight = extractLineWeightNumber(weight);
        return !isNaN(numericWeight) && numericWeight >= 0;
    }

    // =========================================
    // 選択取得 / Selection
    // =========================================
    function getSelectedCellsFromApp() {
        var result = [];
        var i;
        var cells;
        var j;

        if (app.selection.length === 0) return [];

        for (i = 0; i < app.selection.length; i++) {
            cells = getSelectedCells(app.selection[i]);
            for (j = 0; j < cells.length; j++) {
                result.push(cells[j]);
            }
        }

        return result;
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

    function isFullTableSelected(cells) {
        if (!cells || cells.length === 0) return false;
        try {
            var table = cells[0].parentRow.parent;
            var bounds = getBounds(cells);
            return bounds.minRow === 0
                && bounds.maxRow === table.rows.length - 1
                && bounds.minCol === 0
                && bounds.maxCol === table.columns.length - 1;
        } catch (e) {
            return false;
        }
    }


    function clampTintValue(v) {
        if (isNaN(v)) v = 100;
        if (v < 0) v = 0;
        if (v > 100) v = 100;
        return v;
    }

    function adjustFillTintValue(v) {
        v = clampTintValue(v);
        var ks = ScriptUI.environment.keyboardState;
        var step = ks.altKey ? 1 : (ks.shiftKey ? 10 : 5);
        v = Math.round(v / step) * step;
        if (v < 0) v = 0;
        if (v > 100) v = 100;
        return v;
    }

    // =========================================
    // 罫線適用 / Apply borders
    // =========================================
    function applyBorders(cells, mode, weight, clearFirst, swatch, tint) {
        if (cells.length === 0) return;

        var bounds = getBounds(cells);

        if (mode === "allOff") {
            applyAllOff(cells);
            return;
        }


        if (mode === "all") {
            applyAll(cells, weight, clearFirst, swatch, tint);
            return;
        }

        if (mode === "outer") {
            applyOuter(cells, bounds, weight, clearFirst, swatch, tint);
            return;
        }

        if (mode === "innerOnly") {
            applyInnerOnly(cells, bounds, weight, clearFirst, swatch, tint);
            return;
        }

        if (mode === "horizontal") {
            applyHorizontal(cells, bounds, weight, clearFirst, swatch, tint);
            return;
        }

        if (mode === "vertical") {
            applyVertical(cells, bounds, weight, clearFirst, swatch, tint);
            return;
        }

        if (mode === "headerRow") {
            applyTopAndBottomRows(cells, bounds, weight, clearFirst, swatch, tint);
            return;
        }

        if (mode === "headerColumn") {
            applyLeftAndRightColumns(cells, bounds, weight, clearFirst, swatch, tint);
            return;
        }

        if (mode === "clearLeftRight") {
            applyClearLeftRight(cells, bounds);
            return;
        }
    }

    function applyTopAndBottomRows(cells, bounds, weight, clearFirst, swatch, tint) {
        var i, cell, range;
        var isFirstRow, isLastRow;

        if (clearFirst) clearAllEdges(cells);

        for (i = 0; i < cells.length; i++) {
            cell = cells[i];
            range = getCellRange(cell);

            isFirstRow = (range.startRow === bounds.minRow);
            isLastRow = (range.endRow === bounds.maxRow);

            if (isFirstRow) {
                cell.topEdgeStrokeWeight = weight;
                cell.bottomEdgeStrokeWeight = weight;
                setCellEdgeColors(cell, swatch, swatch, null, null);
                applyCellEdgeTints(cell, tint, tint, null, null, swatch, swatch, null, null);
            }

            if (isLastRow) {
                cell.bottomEdgeStrokeWeight = weight;
                setCellEdgeColors(cell, null, swatch, null, null);
                applyCellEdgeTints(cell, null, tint, null, null, null, swatch, null, null);
            }
        }
    }

    function applyAllOff(cells) {
        var i;
        for (i = 0; i < cells.length; i++) {
            clearCellEdgesAndSelectedOpposites(cells[i], cells);
        }
    }

    function applyAll(cells, weight, clearFirst, swatch, tint) {
        var i, cell;
        if (clearFirst) {
            applyAllOff(cells);
        }
        for (i = 0; i < cells.length; i++) {
            cell = cells[i];
            applyEdgeWithSelectedOpposite(cell, cells, "top", weight, swatch, tint);
            applyEdgeWithSelectedOpposite(cell, cells, "bottom", weight, swatch, tint);
            applyEdgeWithSelectedOpposite(cell, cells, "left", weight, swatch, tint);
            applyEdgeWithSelectedOpposite(cell, cells, "right", weight, swatch, tint);
        }
    }

    function applyOuter(cells, bounds, weight, clearFirst, swatch, tint) {
        if (clearFirst) clearAllEdges(cells);

        var i, c, edgeFlags;
        for (i = 0; i < cells.length; i++) {
            c = cells[i];
            edgeFlags = getCellEdgeFlags(c, bounds);

            if (edgeFlags.top) {
                c.topEdgeStrokeWeight = weight;
                c.topEdgeStrokeColor = swatch;
                applyCellEdgeTints(c, tint, null, null, null, swatch, null, null, null);
            }
            if (edgeFlags.bottom) {
                c.bottomEdgeStrokeWeight = weight;
                c.bottomEdgeStrokeColor = swatch;
                applyCellEdgeTints(c, null, tint, null, null, null, swatch, null, null);
            }
            if (edgeFlags.left) {
                c.leftEdgeStrokeWeight = weight;
                c.leftEdgeStrokeColor = swatch;
                applyCellEdgeTints(c, null, null, tint, null, null, null, swatch, null);
            }
            if (edgeFlags.right) {
                c.rightEdgeStrokeWeight = weight;
                c.rightEdgeStrokeColor = swatch;
                applyCellEdgeTints(c, null, null, null, tint, null, null, null, swatch);
            }
        }
    }

    function applyInnerOnly(cells, bounds, weight, clearFirst, swatch, tint) {
        if (clearFirst) {
            applyAllOff(cells);
        }

        var i, c, edgeFlags;
        for (i = 0; i < cells.length; i++) {
            c = cells[i];
            edgeFlags = getCellEdgeFlags(c, bounds);

            if (!edgeFlags.top) {
                applyEdgeWithSelectedOpposite(c, cells, "top", weight, swatch, tint);
            }
            if (!edgeFlags.bottom) {
                applyEdgeWithSelectedOpposite(c, cells, "bottom", weight, swatch, tint);
            }
            if (!edgeFlags.left) {
                applyEdgeWithSelectedOpposite(c, cells, "left", weight, swatch, tint);
            }
            if (!edgeFlags.right) {
                applyEdgeWithSelectedOpposite(c, cells, "right", weight, swatch, tint);
            }
        }
    }

    function applyHorizontal(cells, bounds, weight, clearFirst, swatch, tint) {
        if (clearFirst) {
            applyAllOff(cells);
        }

        var i, cell, edgeFlags;
        for (i = 0; i < cells.length; i++) {
            cell = cells[i];
            edgeFlags = getCellEdgeFlags(cell, bounds);

            if (!edgeFlags.top) {
                applyEdgeWithSelectedOpposite(cell, cells, "top", weight, swatch, tint);
            }
            if (!edgeFlags.bottom) {
                applyEdgeWithSelectedOpposite(cell, cells, "bottom", weight, swatch, tint);
            }
        }
    }

    function applyVertical(cells, bounds, weight, clearFirst, swatch, tint) {
        if (clearFirst) {
            applyAllOff(cells);
        }

        var i, cell, edgeFlags;
        for (i = 0; i < cells.length; i++) {
            cell = cells[i];
            edgeFlags = getCellEdgeFlags(cell, bounds);

            if (!edgeFlags.left) {
                applyEdgeWithSelectedOpposite(cell, cells, "left", weight, swatch, tint);
            }
            if (!edgeFlags.right) {
                applyEdgeWithSelectedOpposite(cell, cells, "right", weight, swatch, tint);
            }
        }
    }

    function applyClearLeftRight(cells, bounds) {
        var i, cell, edgeFlags;

        for (i = 0; i < cells.length; i++) {
            cell = cells[i];
            edgeFlags = getCellEdgeFlags(cell, bounds);

            if (edgeFlags.left) {
                cell.leftEdgeStrokeWeight = 0;
                try {
                    cell.leftEdgeStrokeColor = NothingEnum.NOTHING;
                } catch (e) { }
            }

            if (edgeFlags.right) {
                cell.rightEdgeStrokeWeight = 0;
                try {
                    cell.rightEdgeStrokeColor = NothingEnum.NOTHING;
                } catch (e) { }
            }
        }
    }

    function applyLeftAndRightColumns(cells, bounds, weight, clearFirst, swatch, tint) {
        var i, cell, range;
        var isFirstCol, isLastCol;

        if (clearFirst) clearAllEdges(cells);

        for (i = 0; i < cells.length; i++) {
            cell = cells[i];
            range = getCellRange(cell);

            isFirstCol = (range.startCol === bounds.minCol);
            isLastCol = (range.endCol === bounds.maxCol);

            if (isFirstCol) {
                cell.leftEdgeStrokeWeight = weight;
                cell.rightEdgeStrokeWeight = weight;
                setCellEdgeColors(cell, null, null, swatch, swatch);
                applyCellEdgeTints(cell, null, null, tint, tint, null, null, swatch, swatch);
            }

            if (isLastCol) {
                cell.rightEdgeStrokeWeight = weight;
                setCellEdgeColors(cell, null, null, null, swatch);
                applyCellEdgeTints(cell, null, null, null, tint, null, null, null, swatch);
            }
        }
    }

    
    function applyEdgeWithSelectedOpposite(cell, selectedCells, side, weight, swatch, tint) {
        var adjacent = findAdjacentSelectedCell(cell, selectedCells, side);
        setSingleCellEdge(cell, side, weight, swatch, tint);
        if (adjacent) {
            setSingleCellEdge(adjacent, getOppositeSide(side), weight, swatch, tint);
        }
    }

    function clearCellEdgesAndSelectedOpposites(cell, selectedCells) {
        clearSingleCellEdge(cell, "top");
        clearSingleCellEdge(cell, "bottom");
        clearSingleCellEdge(cell, "left");
        clearSingleCellEdge(cell, "right");

        var topCell = findAdjacentSelectedCell(cell, selectedCells, "top");
        var bottomCell = findAdjacentSelectedCell(cell, selectedCells, "bottom");
        var leftCell = findAdjacentSelectedCell(cell, selectedCells, "left");
        var rightCell = findAdjacentSelectedCell(cell, selectedCells, "right");

        if (topCell) clearSingleCellEdge(topCell, "bottom");
        if (bottomCell) clearSingleCellEdge(bottomCell, "top");
        if (leftCell) clearSingleCellEdge(leftCell, "right");
        if (rightCell) clearSingleCellEdge(rightCell, "left");
    }

    function findAdjacentSelectedCell(cell, selectedCells, side) {
        var baseRange = getCellRange(cell);
        var i, other, otherRange;

        for (i = 0; i < selectedCells.length; i++) {
            other = selectedCells[i];
            if (other === cell) continue;
            otherRange = getCellRange(other);

            if (side === "top") {
                if (otherRange.endRow + 1 === baseRange.startRow && rangesOverlapHorizontally(baseRange, otherRange)) return other;
            } else if (side === "bottom") {
                if (baseRange.endRow + 1 === otherRange.startRow && rangesOverlapHorizontally(baseRange, otherRange)) return other;
            } else if (side === "left") {
                if (otherRange.endCol + 1 === baseRange.startCol && rangesOverlapVertically(baseRange, otherRange)) return other;
            } else if (side === "right") {
                if (baseRange.endCol + 1 === otherRange.startCol && rangesOverlapVertically(baseRange, otherRange)) return other;
            }
        }

        return null;
    }

    function getOppositeSide(side) {
        if (side === "top") return "bottom";
        if (side === "bottom") return "top";
        if (side === "left") return "right";
        return "left";
    }

    function rangesOverlapHorizontally(a, b) {
        return !(a.endCol < b.startCol || b.endCol < a.startCol);
    }

    function rangesOverlapVertically(a, b) {
        return !(a.endRow < b.startRow || b.endRow < a.startRow);
    }

    function setSingleCellEdge(cell, side, weight, swatch, tint) {
        try {
            if (side === "top") {
                cell.topEdgeStrokeWeight = weight;
                cell.topEdgeStrokeColor = swatch;
                applyCellEdgeTints(cell, tint, null, null, null, swatch, null, null, null);
            } else if (side === "bottom") {
                cell.bottomEdgeStrokeWeight = weight;
                cell.bottomEdgeStrokeColor = swatch;
                applyCellEdgeTints(cell, null, tint, null, null, null, swatch, null, null);
            } else if (side === "left") {
                cell.leftEdgeStrokeWeight = weight;
                cell.leftEdgeStrokeColor = swatch;
                applyCellEdgeTints(cell, null, null, tint, null, null, null, swatch, null);
            } else if (side === "right") {
                cell.rightEdgeStrokeWeight = weight;
                cell.rightEdgeStrokeColor = swatch;
                applyCellEdgeTints(cell, null, null, null, tint, null, null, null, swatch);
            }
        } catch (e) { }
    }

    function clearSingleCellEdge(cell, side) {
        try {
            if (side === "top") {
                cell.topEdgeStrokeWeight = 0;
                cell.topEdgeStrokeColor = NothingEnum.NOTHING;
            } else if (side === "bottom") {
                cell.bottomEdgeStrokeWeight = 0;
                cell.bottomEdgeStrokeColor = NothingEnum.NOTHING;
            } else if (side === "left") {
                cell.leftEdgeStrokeWeight = 0;
                cell.leftEdgeStrokeColor = NothingEnum.NOTHING;
            } else if (side === "right") {
                cell.rightEdgeStrokeWeight = 0;
                cell.rightEdgeStrokeColor = NothingEnum.NOTHING;
            }
        } catch (e) { }
    }



    function clearAllEdges(cells) {
        var i, cell;
        for (i = 0; i < cells.length; i++) {
            cell = cells[i];
            setCellEdges(cell, 0, 0, 0, 0);
            try {
                cell.topEdgeStrokeColor = NothingEnum.NOTHING;
                cell.bottomEdgeStrokeColor = NothingEnum.NOTHING;
                cell.leftEdgeStrokeColor = NothingEnum.NOTHING;
                cell.rightEdgeStrokeColor = NothingEnum.NOTHING;
            } catch (e) { }
        }
    }

    function getCellEdgeFlags(cell, bounds) {
        var range = getCellRange(cell);

        return {
            top: range.startRow === bounds.minRow,
            bottom: range.endRow === bounds.maxRow,
            left: range.startCol === bounds.minCol,
            right: range.endCol === bounds.maxCol
        };
    }

    function getBounds(cells) {
        var minRow = 999999;
        var maxRow = -1;
        var minCol = 999999;
        var maxCol = -1;
        var i, range;

        for (i = 0; i < cells.length; i++) {
            range = getCellRange(cells[i]);

            if (range.startRow < minRow) minRow = range.startRow;
            if (range.endRow > maxRow) maxRow = range.endRow;
            if (range.startCol < minCol) minCol = range.startCol;
            if (range.endCol > maxCol) maxCol = range.endCol;
        }

        return {
            minRow: minRow,
            maxRow: maxRow,
            minCol: minCol,
            maxCol: maxCol
        };
    }

    function getCellRange(cell) {
        var startRow = 0;
        var endRow = 0;
        var startCol = 0;
        var endCol = 0;
        var rowSpan = 1;
        var colSpan = 1;

        startRow = cell.parentRow.index;
        startCol = cell.parentColumn.index;

        try {
            if (cell.rowSpan != null && !isNaN(Number(cell.rowSpan))) {
                rowSpan = Math.max(1, Number(cell.rowSpan));
            }
        } catch (e) { }

        try {
            if (cell.columnSpan != null && !isNaN(Number(cell.columnSpan))) {
                colSpan = Math.max(1, Number(cell.columnSpan));
            }
        } catch (e) { }

        endRow = startRow + rowSpan - 1;
        endCol = startCol + colSpan - 1;

        return {
            startRow: startRow,
            endRow: endRow,
            startCol: startCol,
            endCol: endCol
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

    function applyCellEdgeTints(cell, topTint, bottomTint, leftTint, rightTint, topSwatch, bottomSwatch, leftSwatch, rightSwatch) {
        try {
            if (topTint != null && topSwatch != null && !isNoneSwatchName(topSwatch.name) && !isPaperSwatchName(topSwatch.name)) {
                cell.topEdgeStrokeTint = topTint;
            }
        } catch (e) { }
        try {
            if (bottomTint != null && bottomSwatch != null && !isNoneSwatchName(bottomSwatch.name) && !isPaperSwatchName(bottomSwatch.name)) {
                cell.bottomEdgeStrokeTint = bottomTint;
            }
        } catch (e) { }
        try {
            if (leftTint != null && leftSwatch != null && !isNoneSwatchName(leftSwatch.name) && !isPaperSwatchName(leftSwatch.name)) {
                cell.leftEdgeStrokeTint = leftTint;
            }
        } catch (e) { }
        try {
            if (rightTint != null && rightSwatch != null && !isNoneSwatchName(rightSwatch.name) && !isPaperSwatchName(rightSwatch.name)) {
                cell.rightEdgeStrokeTint = rightTint;
            }
        } catch (e) { }
    }

    function isNoneSwatchName(name) {
        var n = normalizeSwatchName(name);
        return n === "none" || n === "なし";
    }

    function isBlackSwatchName(name) {
        var n = normalizeSwatchName(name);
        return n === "black" || n === "ブラック" || n === "黒";
    }

    function isPaperSwatchName(name) {
        var n = normalizeSwatchName(name);
        return n === "paper" || n === "紙色";
    }
})();