#target "InDesign"

/*
 * スクリプトの概要：
 * 選択中の表セルに対して、罫線の描画・消去を切り替えながら適用できるスクリプトです。
 * 「すべて」「境界線のみ」「内部のみ」「水平線のみ」「垂直線のみ」「見出し行」「見出し列」「左右の境界線を消去」「すべて消去」に対応しています。
 *
 * 左カラムではモードと描画オプションを設定し、右カラムでは線幅とカラーを設定します。
 * 線幅は mm 単位で入力でき、プリセットのラジオボタンから「なし」「0.1」「0.2」「0.25」「0.35」「0.5」を素早く選択できます。
 * 線幅入力欄では ↑↓ で 0.1 単位、shift + ↑↓ で 1 単位の増減が可能です。
 *
 * カラーはドキュメントのスウォッチから選択でき、選択したスウォッチは罫線色として反映されます。
 * スウォッチ名はUI表示用に整形され、日本語UIでは Black→黒、Paper→紙色、None→なし と表示されます。
 * Registration / レジストレーションはカラー候補に表示しません。
 *
 * 見出し行は、選択範囲の1行目に上下の罫線を描画し、最終行に下の罫線を描画します。
 * 見出し列は、選択範囲の1列目に左右の罫線を描画し、最終列に右の罫線を描画します。
 * 左右の境界線を消去では、左右に隣接する選択セル間の線を無視し、選択ブロックの左端・右端にある線だけを消去します。
 * 結合セルはセル範囲として扱い、境界判定や重複除外に反映します。
 *
 * プレビューを確認しながら適用でき、「描画前に消去」を OFF にすると既存の罫線を残したまま上書きできます。
 * 例：すべてを細線 → 境界線のみを太線、のような段階的な組み合わせ調整が可能です。
 *
 * 主な機能：
 * - 選択中の表セルに対する罫線の描画／消去
 * - 外枠・内部・水平・垂直・見出し行・見出し列・左右の境界線を消去・すべて消去の各モード切り替え
 * - モード切り替え用ショートカットキー対応（A/E/I/H/V/U/L/R/C）
 * - 「描画前に消去」の ON/OFF 切り替え（Mキーで切り替え）
 * - mm 単位の線幅入力
 * - 線幅プリセットのラジオボタン選択
 * - 線幅入力欄でのキー操作による値変更
 * - スウォッチによる罫線カラー指定
 * - スウォッチプレビュー表示
 * - プレビュー表示による確認
 * - 結合セルを考慮した境界判定
 * - 日本語／英語UI対応
 */

var SCRIPT_VERSION = "v1.4.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "罫線の設定", en: "Border Settings" },
    panelDrawingOptions: { ja: "描画オプション", en: "Drawing Options" },
    modePanel: { ja: "モード", en: "Mode" },
    stylePanel: { ja: "スタイル", en: "Style" },
    clearFirst: { ja: "描画前に消去", en: "Clear Existing Borders First" },
    all: { ja: "すべて", en: "All" },
    outer: { ja: "境界線のみ", en: "Outer Borders Only" },
    inner: { ja: "内部のみ", en: "Inner Borders Only" },
    horizontal: { ja: "水平線のみ", en: "Horizontal Borders Only" },
    vertical: { ja: "垂直線のみ", en: "Vertical Borders Only" },
    headerRow: { ja: "見出し行", en: "Header Row" },
    headerColumn: { ja: "見出し列", en: "Header Column" },
    clearLeftRight: { ja: "左右の境界線を消去", en: "Clear Left/Right Borders" },
    allOff: { ja: "すべて消去", en: "Clear All Borders" },
    lineWidthPanel: { ja: "線幅", en: "Border Weight" },
    lineWidthUnit: { ja: "mm", en: "mm" },
    lineWidthPresetNone: { ja: "なし", en: "None" },
    lineWidthPreset01: { ja: "0.1", en: "0.1" },
    lineWidthPreset02: { ja: "0.2", en: "0.2" },
    lineWidthPreset025: { ja: "0.25", en: "0.25" },
    lineWidthPreset035: { ja: "0.35", en: "0.35" },
    lineWidthPreset05: { ja: "0.5", en: "0.5" },
    colorPanel: { ja: "カラー", en: "Border Color" },
    swatchBlack: { ja: "黒", en: "Black" },
    swatchPaper: { ja: "紙色", en: "Paper" },
    swatchNone: { ja: "なし", en: "None" },
    preview: { ja: "プレビュー", en: "Preview" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    alertSelect: { ja: "表のセルを選択してください。", en: "Please select table cells." },
    alertWeight: { ja: "線幅には0以上の数値を入力してください。", en: "Enter a value of 0 or greater for weight." },
    undoPreview: { ja: "罫線プレビュー", en: "Border Preview" },
    undoApply: { ja: "罫線の設定", en: "Apply Border Settings" },
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

        var settingsColumns = dlg.add("group");
        settingsColumns.orientation = "row";
        settingsColumns.alignChildren = ["fill", "top"];
        settingsColumns.alignment = ["fill", "top"];
        settingsColumns.spacing = 10;

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
        var panelDrawingOptions = leftColumn.add("panel", undefined, L('panelDrawingOptions'));
        panelDrawingOptions.orientation = "column";
        panelDrawingOptions.alignChildren = "left";
        panelDrawingOptions.alignment = ["fill", "top"];
        panelDrawingOptions.margins = [15, 20, 15, 10];

        var cbClearFirst = panelDrawingOptions.add("checkbox", undefined, L('clearFirst'));
        cbClearFirst.value = true;


        var panelStyle = settingsColumns.add("panel", undefined, L('stylePanel'));
        panelStyle.orientation = "column";
        panelStyle.alignChildren = ["fill", "top"];
        panelStyle.alignment = ["fill", "top"];
        panelStyle.margins = [15, 20, 15, 10];

        var panelWeight = panelStyle.add("panel", undefined, L('lineWidthPanel'));
        panelWeight.orientation = "column";
        panelWeight.alignChildren = ["fill", "top"];
        panelWeight.alignment = ["fill", "top"];
        panelWeight.margins = [15, 20, 15, 10];

        var weightGroup = panelWeight.add("group");
        weightGroup.orientation = "column";
        weightGroup.alignChildren = ["left", "top"];
        weightGroup.alignment = ["fill", "top"];
        weightGroup.spacing = 8;

        var weightRow = weightGroup.add("group");
        weightRow.orientation = "row";
        weightRow.alignChildren = ["left", "center"];
        weightRow.alignment = ["left", "center"];
        weightRow.spacing = 8;

        var weightInput = weightRow.add("edittext", undefined, "0.1");
        weightInput.characters = 6;
        weightInput.minimumSize.width = 60;

        weightRow.add("statictext", undefined, L('lineWidthUnit'));

        var weightPresetContainer = weightGroup.add("panel", undefined);
        weightPresetContainer.orientation = "column";
        weightPresetContainer.alignChildren = ["left", "center"];
        weightPresetContainer.alignment = ["fill", "top"];
        weightPresetContainer.margins = [15, 10, 15, 0];

        var weightPresetGroup = weightPresetContainer.add("group");
        weightPresetGroup.orientation = "column";
        weightPresetGroup.alignChildren = ["left", "center"];
        weightPresetGroup.alignment = ["left", "top"];
        weightPresetGroup.spacing = 4;

        var rbWeightNone = weightPresetGroup.add("radiobutton", undefined, L('lineWidthPresetNone'));
        var rbWeight01 = weightPresetGroup.add("radiobutton", undefined, L('lineWidthPreset01'));
        var rbWeight02 = weightPresetGroup.add("radiobutton", undefined, L('lineWidthPreset02'));
        var rbWeight025 = weightPresetGroup.add("radiobutton", undefined, L('lineWidthPreset025'));
        var rbWeight035 = weightPresetGroup.add("radiobutton", undefined, L('lineWidthPreset035'));
        var rbWeight05 = weightPresetGroup.add("radiobutton", undefined, L('lineWidthPreset05'));

        rbWeight01.value = true;

        var panelColor = panelStyle.add("panel", undefined, L('colorPanel'));
        panelColor.orientation = "column";
        panelColor.alignChildren = ["left", "top"];
        panelColor.alignment = ["fill", "top"];
        panelColor.margins = [15, 20, 15, 10];

        var swatchEntries = getSwatchEntries();
        var colorPicker = createSwatchDropdownWithPreview(panelColor, swatchEntries, getDefaultColorIndex(swatchEntries));
        var colorPreviewBox = colorPicker.previewBox;
        var colorDropdown = colorPicker.dropdown;

        rbAll.value = true;

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

        var cbPreview = btnLeftGroup.add("checkbox", undefined, L('preview'));
        cbPreview.value = true;

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

            cbClearFirst: cbClearFirst,
            cbPreview: cbPreview,
            btnCancel: btnCancel,
            btnOk: btnOk,
            drawButtons: [rbAll, rbOuter, rbInnerOnly, rbHorzOnly, rbVertOnly, rbHeaderRow, rbHeaderColumn, rbClearLeftRight, rbAllOff]
        };
    }

    // =========================================
    // イベント / Events
    // =========================================
    function bindDialogEvents(ui, state) {
        var di;

        for (di = 0; di < ui.drawButtons.length; di++) {
            ui.drawButtons[di].onClick = function () {
                onRadioClick(ui, state);
            };
        }

        ui.rbWeightNone.onClick = function () { applyWeightPreset(ui, "0", state); };
        ui.rbWeight01.onClick = function () { applyWeightPreset(ui, "0.1", state); };
        ui.rbWeight02.onClick = function () { applyWeightPreset(ui, "0.2", state); };
        ui.rbWeight025.onClick = function () { applyWeightPreset(ui, "0.25", state); };
        ui.rbWeight035.onClick = function () { applyWeightPreset(ui, "0.35", state); };
        ui.rbWeight05.onClick = function () { applyWeightPreset(ui, "0.5", state); };

        ui.cbPreview.onClick = function () {
            if (ui.cbPreview.value) {
                doPreview(ui, state);
            } else {
                clearPreview(state);
            }
        };

        ui.weightInput.onChange = function () {
            syncWeightPresetFromInput(ui);
            doPreview(ui, state);
        };

        ui.colorDropdown.onChange = function () {
            updateSwatchPreview(ui.colorPreviewBox, ui.colorDropdown, ui.dlg);
            doPreview(ui, state);
        };

        ui.cbClearFirst.onClick = function () {
            doPreview(ui, state);
        };

        changeValueByArrowKey(ui.weightInput, false, function () {
            syncWeightPresetFromInput(ui);
            doPreview(ui, state);
        });

        addDrawingOptionKeyHandler(ui.dlg, ui, state);
        addModeShortcutKeyHandler(ui.dlg, ui, state);

        ui.dlg.onShow = function () {
            doPreview(ui, state);
        };
    }

    function onRadioClick(ui, state) {
        doPreview(ui, state);
    }

    function applyWeightPreset(ui, value, state) {
        if (!ui || !ui.weightInput) return;

        ui.weightInput.text = String(value);
        syncWeightPresetFromInput(ui);
        doPreview(ui, state);
    }

    function syncWeightPresetFromInput(ui) {
        var value = parseFloat(getSelectedWeightText(ui));
        if (!ui || isNaN(value)) return;

        if (ui.rbWeightNone) ui.rbWeightNone.value = (value === 0);
        if (ui.rbWeight01) ui.rbWeight01.value = (value === 0.1);
        if (ui.rbWeight02) ui.rbWeight02.value = (value === 0.2);
        if (ui.rbWeight025) ui.rbWeight025.value = (value === 0.25);
        if (ui.rbWeight035) ui.rbWeight035.value = (value === 0.35);
        if (ui.rbWeight05) ui.rbWeight05.value = (value === 0.5);
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
            if (event.keyName == "M") {
                ui.cbClearFirst.value = !ui.cbClearFirst.value;
                event.preventDefault();
                doPreview(ui, state);
            }
        });
    }

    function addModeShortcutKeyHandler(dialog, ui, state) {
        dialog.addEventListener("keydown", function (event) {
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
            } else if (keyName == "U") {
                ui.rbHeaderRow.value = true;
            } else if (keyName == "L") {
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
        var weight;
        var swatch;
        if (!ui.cbPreview.value) return;

        weight = parseLineWeight(getSelectedWeightText(ui));
        if (isNaN(weight) || weight < 0) {
            clearPreview(state);
            return;
        }

        swatch = getSelectedSwatch(ui);
        if (!swatch) {
            clearPreview(state);
            return;
        }

        clearPreview(state);

        try {
            app.doScript(function () {
                applyBorders(state.cells, getMode(ui), weight, ui.cbClearFirst.value, swatch);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, L('undoPreview'));
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

        if (!swatch) return;

        if (state.previewed) {
            clearPreview(state);
        }

        app.doScript(function () {
            applyBorders(state.cells, mode, weight, ui.cbClearFirst.value, swatch);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, L('undoApply'));
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

        return "0.1";
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
        if (isNoneSwatchName(name)) return L('swatchNone');
        if (isBlackSwatchName(name)) return L('swatchBlack');
        if (isPaperSwatchName(name)) return L('swatchPaper');
        return String(name);
    }


    function isRegistrationSwatchName(name) {
        return name === "Registration" || name === "[Registration]" || name === "レジストレーション" || name === "[レジストレーション]";
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

    // =========================================
    // スウォッチUIヘルパー / Swatch UI helpers
    // =========================================
    function createSwatchDropdownWithPreview(parent, swatchEntries, defaultIndex) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignChildren = ["left", "center"];
        row.spacing = 6;

        var previewBox = createSwatchPreviewBox(row);
        var dropdown = createSwatchDropdown(row, swatchEntries, defaultIndex);

        return {
            row: row,
            previewBox: previewBox,
            dropdown: dropdown
        };
    }

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
        dropdown.minimumSize.width = 90;
        dropdown.preferredSize.width = 90;

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
        try {
            if (!swatchName) return null;
            return app.activeDocument.swatches.itemByName(String(swatchName));
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

        if (isWhiteLikeSwatch(swatch)) return [1, 1, 1];
        if (isBlackLikeSwatch(swatch)) return [0, 0, 0];

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

    function isWhiteLikeSwatch(swatch) {
        var name = swatch && swatch.name != null ? String(swatch.name) : "";
        return isNoneSwatchName(name) || isPaperSwatchName(name);
    }

    function isBlackLikeSwatch(swatch) {
        var name = swatch && swatch.name != null ? String(swatch.name) : "";
        return isRegistrationSwatchName(name) || isBlackSwatchName(name);
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
        var result = [];
        var seen = {};
        var i;
        var cells;
        var j;
        var key;

        if (app.selection.length === 0) return [];

        for (i = 0; i < app.selection.length; i++) {
            cells = getSelectedCells(app.selection[i]);
            for (j = 0; j < cells.length; j++) {
                key = getCellKey(cells[j]);
                if (!seen[key]) {
                    seen[key] = true;
                    result.push(cells[j]);
                }
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

        if (mode === "headerRow") {
            applyHeaderRow(cells, bounds, weight, clearFirst, swatch);
            return;
        }

        if (mode === "headerColumn") {
            applyHeaderColumn(cells, bounds, weight, clearFirst, swatch);
            return;
        }

        if (mode === "clearLeftRight") {
            applyClearLeftRight(cells);
            return;
        }
    }

    function applyHeaderRow(cells, bounds, weight, clearFirst, swatch) {
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
            }

            if (isLastRow) {
                cell.bottomEdgeStrokeWeight = weight;
                setCellEdgeColors(cell, null, swatch, null, null);
            }
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
            } catch (e) { }
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
    function applyClearLeftRight(cells) {
        var i, cell;
        var hasLeftNeighbor, hasRightNeighbor;

        for (i = 0; i < cells.length; i++) {
            cell = cells[i];
            hasLeftNeighbor = hasAdjacentSelectedCellOnLeft(cell, cells);
            hasRightNeighbor = hasAdjacentSelectedCellOnRight(cell, cells);

            if (!hasLeftNeighbor) {
                cell.leftEdgeStrokeWeight = 0;
                try {
                    cell.leftEdgeStrokeColor = NothingEnum.NOTHING;
                } catch (e) { }
            }

            if (!hasRightNeighbor) {
                cell.rightEdgeStrokeWeight = 0;
                try {
                    cell.rightEdgeStrokeColor = NothingEnum.NOTHING;
                } catch (e) { }
            }
        }
    }

    function applyHeaderColumn(cells, bounds, weight, clearFirst, swatch) {
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
            }

            if (isLastCol) {
                cell.rightEdgeStrokeWeight = weight;
                setCellEdgeColors(cell, null, null, null, swatch);
            }
        }
    }

    function hasAdjacentSelectedCellOnLeft(cell, cells) {
        return hasAdjacentSelectedCell(cell, cells, -1);
    }

    function hasAdjacentSelectedCellOnRight(cell, cells) {
        return hasAdjacentSelectedCell(cell, cells, 1);
    }

    function hasAdjacentSelectedCell(cell, cells, direction) {
        var baseRange = getCellRange(cell);
        var i, other, otherRange;

        for (i = 0; i < cells.length; i++) {
            other = cells[i];
            if (other === cell) continue;

            otherRange = getCellRange(other);

            if (!rangesOverlapVertically(baseRange, otherRange)) continue;

            if (direction < 0) {
                if (otherRange.endCol + 1 === baseRange.startCol) return true;
            } else {
                if (baseRange.endCol + 1 === otherRange.startCol) return true;
            }
        }

        return false;
    }

    function rangesOverlapVertically(a, b) {
        return !(a.endRow < b.startRow || b.endRow < a.startRow);
    }

    function clearAllEdges(cells) {
        var i;
        for (i = 0; i < cells.length; i++) {
            setCellEdges(cells[i], 0, 0, 0, 0);
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

    function isNoneSwatchName(name) {
        return name === "None" || name === "[None]" || name === "なし" || name === "[なし]";
    }

    function isBlackSwatchName(name) {
        return name === "Black" || name === "[Black]" || name === "ブラック" || name === "黒";
    }

    function isPaperSwatchName(name) {
        return name === "Paper" || name === "[Paper]" || name === "紙色" || name === "[紙色]";
    }
})();