#target indesign

/*
 * SmartBorderBuilder.jsx
 *
 * 選択した表セルに対して、モード・線幅・カラー・濃淡を指定しながら罫線をプレビュー付きで描画・消去します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartBorderBuilder";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-13";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/SmartBorderBuilder.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/SmartBorderBuilder.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 線幅プリセットの候補（現在の線幅単位で解釈）/ Border-weight presets, read in the current weight unit */
var WEIGHT_PRESET_VALUES = ["0.1", "0.2", "0.25", "0.35", "0.5"];

/* 濃淡（Tint）の初期値と範囲 / Initial value and range of the tint control */
var TINT_DEFAULT = 100;
var TINT_MIN     = 0;
var TINT_MAX     = 100;

// =========================================
// レイアウト設定 / Layout settings
// =========================================

/* 線幅入力欄・濃淡入力欄の文字数と最小幅（px）/ Character width and minimum width of the weight and tint fields (px) */
var WEIGHT_INPUT_CHARACTERS = 6;
var WEIGHT_INPUT_MIN_WIDTH  = 60;
var TINT_INPUT_CHARACTERS   = 4;
var TINT_INPUT_MIN_WIDTH    = 45;

/* ボタン列の左右を分けるスペーサーの最小幅（px）/ Minimum width of the spacer between the button clusters (px) */
var BUTTON_ROW_SPACER_MIN_WIDTH = 40;

// ==============================
// UIレイアウトの共通設定 / Shared UI layout
// ==============================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

/**
 * ウィンドウの共通設定を適用する
 * @param {Window} win 対象ウィンドウ
 * @param {number} [spacing] 要素間隔。省略時は WINDOW_SPACING
 * @returns {void}
 */
function setupWindow(win, spacing) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルの共通設定を適用する
 * @param {Panel} panel 対象パネル
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * 行グループの共通設定を適用する（ボタン列など）
 * @param {Group} group 対象グループ
 * @param {string} [alignment] 配置。省略時は "left"
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// =========================================
// ラベル定義 / Labels
// =========================================

/**
 * UI 言語を判定する
 * @returns {string} "ja" または "en"
 */
function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var currentLang = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "罫線の設定", en: "Border Settings" }
    },
    panel: {
        mode:      { ja: "モード", en: "Mode" },
        style:     { ja: "スタイル", en: "Style" },
        lineWidth: { ja: "線幅", en: "Border Weight" },
        color:     { ja: "カラー", en: "Border Color" }
    },
    radio: {
        all:            { ja: "すべて", en: "All" },
        outer:          { ja: "境界線のみ", en: "Outer Borders Only" },
        inner:          { ja: "内部のみ", en: "Inner Borders Only" },
        horizontal:     { ja: "水平線のみ", en: "Horizontal Borders Only" },
        vertical:       { ja: "垂直線のみ", en: "Vertical Borders Only" },
        headerRow:      { ja: "見出し行", en: "Header Row" },
        headerColumn:   { ja: "見出し列", en: "Header Column" },
        clearLeftRight: { ja: "左右の境界線を消去", en: "Clear Left/Right Borders" },
        allOff:         { ja: "すべて消去", en: "Clear All Borders" },
        weightNone:     { ja: "なし", en: "None" }
    },
    checkbox: {
        clearFirst: { ja: "描画前に消去", en: "Clear Existing Borders First" }
    },
    field: {
        tint: { ja: "濃淡：", en: "Tint:" }
    },
    swatch: {
        black: { ja: "黒", en: "Black" },
        paper: { ja: "紙色", en: "Paper" },
        none:  { ja: "なし", en: "None" }
    },
    button: {
        ok:           { ja: "OK", en: "OK" },
        cancel:       { ja: "キャンセル", en: "Cancel" },
        standardMode: { ja: "標準モード", en: "Standard Mode" },
        previewMode:  { ja: "プレビュー", en: "Preview" }
    },
    tooltip: {
        tintSlider:     { ja: "0〜100", en: "0–100" },
        all:            { ja: "ショートカット: A / Option+クリックで消去ON/OFF", en: "Shortcut: A / Option-click toggles Clear Existing Borders First" },
        outer:          { ja: "ショートカット: E / Option+クリックで消去ON/OFF", en: "Shortcut: E / Option-click toggles Clear Existing Borders First" },
        inner:          { ja: "ショートカット: I / Option+クリックで消去ON/OFF", en: "Shortcut: I / Option-click toggles Clear Existing Borders First" },
        horizontal:     { ja: "ショートカット: H / Option+クリックで消去ON/OFF", en: "Shortcut: H / Option-click toggles Clear Existing Borders First" },
        vertical:       { ja: "ショートカット: V / Option+クリックで消去ON/OFF", en: "Shortcut: V / Option-click toggles Clear Existing Borders First" },
        headerRow:      { ja: "ショートカット: U / Option+クリックで消去ON/OFF", en: "Shortcut: U / Option-click toggles Clear Existing Borders First" },
        headerColumn:   { ja: "ショートカット: L / Option+クリックで消去ON/OFF", en: "Shortcut: L / Option-click toggles Clear Existing Borders First" },
        clearLeftRight: { ja: "ショートカット: R / Option+クリックで消去ON/OFF", en: "Shortcut: R / Option-click toggles Clear Existing Borders First" },
        allOff:         { ja: "ショートカット: C / Option+クリックで消去ON/OFF", en: "Shortcut: C / Option-click toggles Clear Existing Borders First" }
    },
    alert: {
        select: { ja: "表のセルを選択してください。", en: "Please select table cells." },
        weight: { ja: "線幅には0以上の数値を入力してください。", en: "Enter a value of 0 or greater for weight." },
        tint:   { ja: "濃淡には0〜100の数値を入力してください。", en: "Enter a value between 0 and 100 for tint." }
    },
    undo: {
        preview: { ja: "罫線プレビュー", en: "Border Preview" },
        apply:   { ja: "罫線の設定", en: "Apply Border Settings" }
    }
};

/**
 * ドット区切りキーでラベルを取得する
 * @param {string} labelKey 例: "dialog.title"
 * @returns {string} 現在の言語のラベル文字列。見つからない場合はキーをそのまま返す
 */
function getLabel(labelKey) {
    var node = LABELS;
    var keyParts = labelKey.split(".");
    for (var i = 0; i < keyParts.length; i++) {
        node = node[keyParts[i]];
        if (!node) return labelKey;
    }
    return node[currentLang] || node.en || labelKey;
}

(function () {
    var cells = getSelectedCellsFromApp();
    if (cells.length === 0) {
        alert(getLabel('alert.select'));
        return;
    }

    var originalSelection = snapshotSelection(app.selection);
    var state = {
        cells: cells,
        previewed: false,
        originalSelection: originalSelection
    };
    state.headerModesEnabled = isFullTableSelection(cells);

    app.selection = NothingEnum.NOTHING;

    var ui = buildDialog();
    bindDialogEvents(ui, state);

    var result = ui.dlg.show();
    if (result != 1) {
        clearPreview(state);
        restoreSelection(state.originalSelection);
        return;
    }

    applyFinalFromDialog(ui, state);
    restoreSelection(state.originalSelection);

    // =========================================
    // UI構築 / Build UI
    // =========================================
    /**
     * 罫線設定ダイアログを組み立てる
     * @returns {object} ダイアログとコントロールをまとめたオブジェクト
     */
    function buildDialog() {
        var dlg = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        setupWindow(dlg, 10);

        var settingsColumns = dlg.add("group");
        setupRow(settingsColumns, "fill", COLUMN_SPACING);
        settingsColumns.alignChildren = ["fill", "top"];

        var leftColumn = settingsColumns.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];
        leftColumn.alignment = ["fill", "top"];
        leftColumn.spacing = PANEL_SPACING;

        var panelMode = leftColumn.add("panel", undefined, getLabel('panel.mode'));
        setupPanel(panelMode, 6);
        panelMode.alignChildren = "left";

        var rbAll = panelMode.add("radiobutton", undefined, getLabel('radio.all'));
        rbAll.helpTip = getLabel('tooltip.all');
        var rbOuter = panelMode.add("radiobutton", undefined, getLabel('radio.outer'));
        rbOuter.helpTip = getLabel('tooltip.outer');
        var rbInnerOnly = panelMode.add("radiobutton", undefined, getLabel('radio.inner'));
        rbInnerOnly.helpTip = getLabel('tooltip.inner');
        var rbHorzOnly = panelMode.add("radiobutton", undefined, getLabel('radio.horizontal'));
        rbHorzOnly.helpTip = getLabel('tooltip.horizontal');
        var rbVertOnly = panelMode.add("radiobutton", undefined, getLabel('radio.vertical'));
        rbVertOnly.helpTip = getLabel('tooltip.vertical');
        var rbHeaderRow = panelMode.add("radiobutton", undefined, getLabel('radio.headerRow'));
        rbHeaderRow.helpTip = getLabel('tooltip.headerRow');
        var rbHeaderColumn = panelMode.add("radiobutton", undefined, getLabel('radio.headerColumn'));
        rbHeaderColumn.helpTip = getLabel('tooltip.headerColumn');
        rbHeaderRow.enabled = state.headerModesEnabled;
        rbHeaderColumn.enabled = state.headerModesEnabled;
        var rbClearLeftRight = panelMode.add("radiobutton", undefined, getLabel('radio.clearLeftRight'));
        rbClearLeftRight.helpTip = getLabel('tooltip.clearLeftRight');
        var rbAllOff = panelMode.add("radiobutton", undefined, getLabel('radio.allOff'));
        rbAllOff.helpTip = getLabel('tooltip.allOff');

        /* 枠線なしのグループで描画オプションを並べる / Drawing options sit in a borderless group */
        var panelDrawingOptions = leftColumn.add("group");
        panelDrawingOptions.orientation = "column";
        panelDrawingOptions.alignChildren = "left";
        panelDrawingOptions.alignment = ["fill", "top"];
        panelDrawingOptions.margins = [PANEL_MARGINS[0], 10, PANEL_MARGINS[2], 10];


        var cbClearFirst = panelDrawingOptions.add("checkbox", undefined, getLabel('checkbox.clearFirst'));
        cbClearFirst.value = true;


        var panelStyle = settingsColumns.add("panel", undefined, getLabel('panel.style'));
        setupPanel(panelStyle, 8);

        var panelWeight = panelStyle.add("panel", undefined, getLabel('panel.lineWidth'));
        setupPanel(panelWeight, 6);

        var weightGroup = panelWeight.add("group");
        weightGroup.orientation = "column";
        weightGroup.alignChildren = ["left", "top"];
        weightGroup.alignment = ["fill", "top"];
        weightGroup.spacing = 8;

        var weightRow = weightGroup.add("group");
        setupRow(weightRow, "left", 8);
        weightRow.alignChildren = ["left", "center"];

        var weightInput = weightRow.add("edittext", undefined, getDefaultLineWidthText());
        weightInput.characters = WEIGHT_INPUT_CHARACTERS;
        weightInput.minimumSize.width = WEIGHT_INPUT_MIN_WIDTH;

        weightRow.add("statictext", undefined, getCurrentLineWidthUnitLabel());

        var weightPresetContainer = weightGroup.add("group");
        weightPresetContainer.orientation = "column";
        weightPresetContainer.alignChildren = ["left", "center"];
        weightPresetContainer.alignment = ["fill", "top"];
        weightPresetContainer.margins = [PANEL_MARGINS[0], 10, PANEL_MARGINS[2], 10];

        var weightPresetGroup = weightPresetContainer.add("group");
        weightPresetGroup.orientation = "column";
        weightPresetGroup.alignChildren = ["left", "center"];
        weightPresetGroup.alignment = ["left", "top"];
        weightPresetGroup.spacing = 4;

        var rbWeightNone = weightPresetGroup.add("radiobutton", undefined, getLabel('radio.weightNone'));
        var rbWeight01 = weightPresetGroup.add("radiobutton", undefined, WEIGHT_PRESET_VALUES[0]);
        var rbWeight02 = weightPresetGroup.add("radiobutton", undefined, WEIGHT_PRESET_VALUES[1]);
        var rbWeight025 = weightPresetGroup.add("radiobutton", undefined, WEIGHT_PRESET_VALUES[2]);
        var rbWeight035 = weightPresetGroup.add("radiobutton", undefined, WEIGHT_PRESET_VALUES[3]);
        var rbWeight05 = weightPresetGroup.add("radiobutton", undefined, WEIGHT_PRESET_VALUES[4]);

        syncWeightPresetFromTextValue({
            rbWeightNone: rbWeightNone,
            rbWeight01: rbWeight01,
            rbWeight02: rbWeight02,
            rbWeight025: rbWeight025,
            rbWeight035: rbWeight035,
            rbWeight05: rbWeight05
        }, getDefaultLineWidthText());

        var panelColor = panelStyle.add("panel", undefined, getLabel('panel.color'));
        setupPanel(panelColor, 6);
        panelColor.alignChildren = ["left", "top"];

        var swatchEntries = getSwatchEntries();
        var colorPicker = createSwatchDropdownWithPreview(panelColor, swatchEntries, getDefaultColorIndex(swatchEntries));
        var colorPreviewBox = colorPicker.previewBox;
        var colorDropdown = colorPicker.dropdown;

        var panelTint = panelColor.add("group");
        panelTint.orientation = "column";
        panelTint.alignChildren = ["fill", "top"];
        panelTint.alignment = ["fill", "top"];

        var tintRow = panelTint.add("group");
        setupRow(tintRow, "fill", 8);
        tintRow.alignChildren = ["left", "center"];

        tintRow.add("statictext", undefined, getLabel('field.tint'));

        /* 濃淡の初期値は TINT_DEFAULT / The tint starts at TINT_DEFAULT */
        var tintInput = tintRow.add("edittext", undefined, String(TINT_DEFAULT));
        tintInput.characters = TINT_INPUT_CHARACTERS;
        tintInput.minimumSize.width = TINT_INPUT_MIN_WIDTH;

        var tintSlider = panelTint.add("slider", undefined, TINT_DEFAULT, TINT_MIN, TINT_MAX);
        tintSlider.helpTip = getLabel('tooltip.tintSlider');

        rbAll.value = true;
        if (!state.headerModesEnabled) {
            if (rbHeaderRow.value) rbHeaderRow.value = false;
            if (rbHeaderColumn.value) rbHeaderColumn.value = false;
            rbAll.value = true;
        }
        updateTintControlsEnabledState(colorDropdown, tintInput, tintSlider);

        /* カラープレビューの初期表示 / Paint the initial colour preview */
        updateSwatchPreview(colorPreviewBox, colorDropdown, dlg);

        var btnArea = dlg.add("group");
        setupRow(btnArea, "fill", 8);
        btnArea.alignChildren = ["fill", "fill"];
        btnArea.margins = [0, 8, 0, 0];

        var btnLeftGroup = btnArea.add("group");
        btnLeftGroup.orientation = "column";
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.alignment = ["left", "fill"];

        var btnStandardMode = btnLeftGroup.add("button", undefined, "");
        updatePreviewToggleButtonLabel(btnStandardMode);

        var btnCenterGroup = btnArea.add("group");
        btnCenterGroup.orientation = "column";
        btnCenterGroup.alignChildren = ["fill", "fill"];
        btnCenterGroup.alignment = ["fill", "fill"];

        var spacer = btnCenterGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = BUTTON_ROW_SPACER_MIN_WIDTH;

        var btnRightGroup = btnArea.add("group");
        btnRightGroup.orientation = "column";
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.alignment = ["right", "fill"];

        /* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
        var buttonRow = btnRightGroup.add("group");
        setupRow(buttonRow, "right", 8);
        buttonRow.alignChildren = ["right", "center"];

        var btnCancel = buttonRow.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
        var btnOk = buttonRow.add("button", undefined, getLabel('button.ok'), { name: "ok" });

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
            tintInput: tintInput,
            tintSlider: tintSlider,

            cbClearFirst: cbClearFirst,
            btnStandardMode: btnStandardMode,
            btnCancel: btnCancel,
            btnOk: btnOk,
            drawButtons: [rbAll, rbOuter, rbInnerOnly, rbHorzOnly, rbVertOnly, rbHeaderRow, rbHeaderColumn, rbClearLeftRight, rbAllOff]
        };
    }

    // =========================================
    // イベント / Events
    // =========================================
    /**
     * ダイアログのコントロールにイベントを結び付ける
     * @param {object} ui buildDialog が返した UI オブジェクト
     * @param {object} state 選択セルやプレビュー状態を保持するオブジェクト
     * @returns {void}
     */
    function bindDialogEvents(ui, state) {
        var di;

        for (di = 0; di < ui.drawButtons.length; di++) {
            ui.drawButtons[di].onClick = function () {
                var keyboard = ScriptUI.environment.keyboardState;

                // Option (Alt) + click → toggle "描画前に消去"
                if (keyboard && keyboard.altKey) {
                    ui.cbClearFirst.value = !ui.cbClearFirst.value;
                }

                onRadioClick(ui, state);
            };
        }

        ui.rbWeightNone.onClick = function () { applyWeightPreset(ui, "0", state); };
        ui.rbWeight01.onClick = function () { applyWeightPreset(ui, "0.1", state); };
        ui.rbWeight02.onClick = function () { applyWeightPreset(ui, "0.2", state); };
        ui.rbWeight025.onClick = function () { applyWeightPreset(ui, "0.25", state); };
        ui.rbWeight035.onClick = function () { applyWeightPreset(ui, "0.35", state); };
        ui.rbWeight05.onClick = function () { applyWeightPreset(ui, "0.5", state); };

        ui.btnStandardMode.onClick = function () {
            togglePreviewScreenMode();
            updatePreviewToggleButtonLabel(ui.btnStandardMode);
        };

        ui.weightInput.onChange = function () {
            syncWeightPresetFromInput(ui);
            doPreview(ui, state);
        };

        ui.colorDropdown.onChange = function () {
            updateSwatchPreview(ui.colorPreviewBox, ui.colorDropdown, ui.dlg);
            updateTintControlsEnabledState(ui.colorDropdown, ui.tintInput, ui.tintSlider);
            doPreview(ui, state);
        };
        ui.tintInput.onChange = function () {
            clampTintInput(ui.tintInput);
            syncTintSliderFromInput(ui);
            doPreview(ui, state);
        };

        ui.tintSlider.onChanging = function () {
            var keyboard = ScriptUI.environment.keyboardState;
            var value = ui.tintSlider.value;

            if (keyboard && keyboard.shiftKey) {
                value = Math.round(value / 10) * 10;
                ui.tintSlider.value = value;
            }

            syncTintInputFromSlider(ui);
        };

        ui.tintSlider.onChange = function () {
            var keyboard = ScriptUI.environment.keyboardState;
            var value = ui.tintSlider.value;

            if (keyboard && keyboard.shiftKey) {
                value = Math.round(value / 10) * 10;
                ui.tintSlider.value = value;
            } else {
                value = Math.round(value);
                ui.tintSlider.value = value;
            }

            syncTintInputFromSlider(ui);
            doPreview(ui, state);
        };

        ui.cbClearFirst.onClick = function () {
            doPreview(ui, state);
        };

        changeValueByArrowKey(ui.weightInput, false, function () {
            syncWeightPresetFromInput(ui);
            doPreview(ui, state);
        });

        changeValueByArrowKey(ui.tintInput, false, function () {
            clampTintInput(ui.tintInput);
            syncTintSliderFromInput(ui);
            doPreview(ui, state);
        });

        addDrawingOptionKeyHandler(ui.dlg, ui, state);
        addModeShortcutKeyHandler(ui.dlg, ui, state);

        ui.dlg.onShow = function () {
            updatePreviewToggleButtonLabel(ui.btnStandardMode);
            doPreview(ui, state);
        };
    }

    /**
     * モードのラジオボタンが押されたときにプレビューを更新する
     * @param {object} ui UI オブジェクト
     * @param {object} state 状態オブジェクト
     * @returns {void}
     */
    function onRadioClick(ui, state) {
        doPreview(ui, state);
    }

    /**
     * 線幅プリセットの値を入力欄へ反映してプレビューを更新する
     * @param {object} ui UI オブジェクト
     * @param {string} value プリセットの線幅文字列
     * @param {object} state 状態オブジェクト
     * @returns {void}
     */
    function applyWeightPreset(ui, value, state) {
        if (!ui || !ui.weightInput) return;

        ui.weightInput.text = String(value);
        syncWeightPresetFromInput(ui);
        doPreview(ui, state);
    }

    /**
     * 線幅入力欄の値に合わせてプリセットの選択状態を揃える
     * @param {object} ui UI オブジェクト
     * @returns {void}
     */
    function syncWeightPresetFromInput(ui) {
        if (!ui) return;
        syncWeightPresetFromTextValue(ui, getSelectedWeightText(ui));
    }

    /**
     * 入力欄に上下キーでの増減操作を追加する
     * @param {EditText} editText 対象の入力欄
     * @param {boolean} allowNegative 負の値を許可するか
     * @param {function} onAfterChange 値の変更後に呼ぶ処理
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onAfterChange) {
        editText.addEventListener("keydown", function (event) {
            if (applyArrowStepToEditText(editText, allowNegative, event, onAfterChange)) {
                event.preventDefault();
            }
        });
    }

    /**
     * 押されたキーに応じて入力欄の数値を増減する
     * @param {EditText} editText 対象の入力欄
     * @param {boolean} allowNegative 負の値を許可するか
     * @param {object} event キーイベント
     * @param {function} onAfterChange 値の変更後に呼ぶ処理
     * @returns {void}
     */
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

    /**
     * 環境差のあるキー名を Up / Down に正規化する
     * @param {string} keyName イベントから得たキー名
     * @returns {string} "Up"、"Down"、または空文字
     */
    function normalizeArrowKeyName(keyName) {
        keyName = String(keyName);
        if (keyName === "Up" || keyName === "Down") return keyName;
        if (keyName === "UpArrow") return "Up";
        if (keyName === "DownArrow") return "Down";
        if (keyName === "PageUp") return "Up";
        if (keyName === "PageDown") return "Down";
        return keyName;
    }

    /**
     * 描画オプションを切り替えるキー操作を登録する
     * @param {Window} dialog 対象のダイアログ
     * @param {object} ui UI オブジェクト
     * @param {object} state 状態オブジェクト
     * @returns {void}
     */
    function addDrawingOptionKeyHandler(dialog, ui, state) {
        dialog.addEventListener("keydown", function (event) {
            if (event.keyName == "M") {
                ui.cbClearFirst.value = !ui.cbClearFirst.value;
                event.preventDefault();
                doPreview(ui, state);
            }
        });
    }

    /**
     * モード切り替えのショートカットキーを登録する
     * @param {Window} dialog 対象のダイアログ
     * @param {object} ui UI オブジェクト
     * @param {object} state 状態オブジェクト
     * @returns {void}
     */
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
    /**
     * 現在の設定で罫線のプレビューを描画する
     * @param {object} ui UI オブジェクト
     * @param {object} state 状態オブジェクト
     * @returns {void}
     */
    function doPreview(ui, state) {
        var weight;
        var swatch;
        var tint;

        weight = parseLineWeight(getSelectedWeightText(ui));
        if (!isValidLineWeight(weight)) {
            clearPreview(state);
            return;
        }

        swatch = getSelectedSwatch(ui);
        if (!swatch) {
            clearPreview(state);
            return;
        }

        tint = getSelectedTint(ui);
        if (!isValidTint(tint)) {
            clearPreview(state);
            return;
        }

        clearPreview(state);

        try {
            app.doScript(function () {
                applyBorders(state.cells, getMode(ui), weight, ui.cbClearFirst.value, swatch, tint);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel('undo.preview'));
            state.previewed = true;
            app.activeDocument.recompose();
        } catch (e) {
            state.previewed = false;
        }
    }

    /**
     * プレビューとして適用した罫線を取り消す
     * @param {object} state 状態オブジェクト
     * @returns {void}
     */
    function clearPreview(state) {
        if (!state.previewed) return;
        try {
            app.undo();
        } catch (e) { }
        state.previewed = false;
        app.activeDocument.recompose();
    }

    /**
     * ダイアログの設定を確定して罫線を適用する
     * @param {object} ui UI オブジェクト
     * @param {object} state 状態オブジェクト
     * @returns {void}
     */
    function applyFinalFromDialog(ui, state) {
        var mode = getMode(ui);
        var weight = parseLineWeight(getSelectedWeightText(ui));
        var swatch = getSelectedSwatch(ui);
        var tint = getSelectedTint(ui);

        if (!isValidLineWeight(weight)) {
            clearPreview(state);
            alert(getLabel('alert.weight'));
            return;
        }

        if (!isValidTint(tint)) {
            clearPreview(state);
            alert(getLabel('alert.tint'));
            return;
        }

        if (!swatch) return;

        if (state.previewed) {
            clearPreview(state);
        }

        app.doScript(function () {
            applyBorders(state.cells, mode, weight, ui.cbClearFirst.value, swatch, tint);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel('undo.apply'));
    }

    // =========================================
    // UI値の取得 / Read UI values
    // =========================================
    /**
     * 選択中のモードを取得する
     * @param {object} ui UI オブジェクト
     * @returns {string} モードを表す識別子
     */
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

    /**
     * 入力欄またはプリセットから線幅の文字列を取得する
     * @param {object} ui UI オブジェクト
     * @returns {string} 線幅を表す文字列
     */
    function getSelectedWeightText(ui) {
        var text = "";

        if (ui.weightInput && ui.weightInput.text != null) {
            text = String(ui.weightInput.text).replace(/^\s+|\s+$/g, "");
            if (text !== "") return text;
        }

        return getDefaultLineWidthText();
    }

    /**
     * カラー候補として表示するスウォッチ一覧を作る
     * @returns {Array<object>} スウォッチ名と表示名の配列
     */
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
            entries.push({ displayName: getLabel('swatch.black'), actualName: 'Black' });
        }
        return entries;
    }

    /**
     * スウォッチ名を UI 表示用の名前に変換する
     * @param {string} name スウォッチ名
     * @returns {string} 表示用の名前
     */
    function getDisplaySwatchName(name) {
        if (isNoneSwatchName(name)) return getLabel('swatch.none');
        if (isBlackSwatchName(name)) return getLabel('swatch.black');
        if (isPaperSwatchName(name)) return getLabel('swatch.paper');
        return String(name);
    }


    /**
     * レジストレーションのスウォッチ名かどうかを判定する
     * @param {string} name スウォッチ名
     * @returns {boolean} レジストレーションなら true
     */
    function isRegistrationSwatchName(name) {
        return name === "Registration" || name === "[Registration]" || name === "レジストレーション" || name === "[レジストレーション]";
    }

    /**
     * 既定で選択するカラーの位置を求める
     * @param {Array<object>} swatchEntries スウォッチ一覧
     * @returns {number} 既定で選ぶ位置
     */
    function getDefaultColorIndex(swatchEntries) {
        var i;
        for (i = 0; i < swatchEntries.length; i++) {
            if (isBlackSwatchName(String(swatchEntries[i].actualName))) return i;
        }
        return 0;
    }

    /**
     * 選択中のカラー名を取得する
     * @param {object} ui UI オブジェクト
     * @returns {string} スウォッチ名
     */
    function getSelectedColorName(ui) {
        return getSelectedSwatchNameFromDropdown(ui ? ui.colorDropdown : null);
    }

    /**
     * ドロップダウンで選択中のスウォッチ名を取得する
     * @param {DropDownList} dropdown 対象のドロップダウン
     * @returns {string} スウォッチ名
     */
    function getSelectedSwatchNameFromDropdown(dropdown) {
        if (!dropdown || !dropdown.selection) return "";
        if (dropdown.selection._swatchName != null) return String(dropdown.selection._swatchName);
        return String(dropdown.selection.text);
    }

    /**
     * 入力欄から濃淡の値を取得する
     * @param {object} ui UI オブジェクト
     * @returns {number} 濃淡の値
     */
    function getSelectedTint(ui) {
        if (!ui || !ui.tintInput) return 100;
        return clampTintValue(parseFloat(String(ui.tintInput.text).replace(/^\s+|\s+$/g, "")));
    }

    /**
     * 濃淡の値を有効範囲に収める
     * @param {number} value 濃淡の値
     * @returns {number} 範囲内に収めた値
     */
    function clampTintValue(value) {
        if (isNaN(value)) return NaN;
        if (value < 0) return 0;
        if (value > 100) return 100;
        return value;
    }

    /**
     * 濃淡の値が有効範囲内かどうかを判定する
     * @param {number} value 濃淡の値
     * @returns {boolean} 有効なら true
     */
    function isValidTint(value) {
        return !isNaN(value) && value >= 0 && value <= 100;
    }

    /**
     * 濃淡入力欄の値を有効範囲に収めて書き戻す
     * @param {EditText} editText 濃淡の入力欄
     * @returns {number} 収めたあとの値
     */
    function clampTintInput(editText) {
        var value;
        if (!editText) return;
        value = clampTintValue(parseFloat(String(editText.text).replace(/^\s+|\s+$/g, "")));
        if (isNaN(value)) return;
        editText.text = String(Math.round(value));
    }

    /**
     * 入力欄の値をスライダーへ反映する
     * @param {object} ui UI オブジェクト
     * @returns {void}
     */
    function syncTintSliderFromInput(ui) {
        var tint;
        if (!ui || !ui.tintInput || !ui.tintSlider) return;
        clampTintInput(ui.tintInput);
        tint = getSelectedTint(ui);
        if (isNaN(tint)) return;
        ui.tintSlider.value = tint;
    }

    /**
     * スライダーの値を入力欄へ反映する
     * @param {object} ui UI オブジェクト
     * @returns {void}
     */
    function syncTintInputFromSlider(ui) {
        if (!ui || !ui.tintInput || !ui.tintSlider) return;
        ui.tintInput.text = String(Math.round(ui.tintSlider.value));
    }

    /**
     * そのスウォッチで濃淡を調整できるかどうかを判定する
     * @param {string} swatchName スウォッチ名
     * @returns {boolean} 調整できるなら true
     */
    function shouldEnableTintControlsBySwatchName(swatchName) {
        return !(isNoneSwatchName(swatchName) || isPaperSwatchName(swatchName));
    }

    /**
     * 選択中のスウォッチに応じて濃淡コントロールの有効／無効を切り替える
     * @param {DropDownList} dropdown カラーのドロップダウン
     * @param {EditText} tintInput 濃淡の入力欄
     * @param {Slider} tintSlider 濃淡のスライダー
     * @returns {void}
     */
    function updateTintControlsEnabledState(dropdown, tintInput, tintSlider) {
        var swatchName = getSelectedSwatchNameFromDropdown(dropdown);
        var enabled = shouldEnableTintControlsBySwatchName(swatchName);

        if (tintInput) tintInput.enabled = enabled;
        if (tintSlider) tintSlider.enabled = enabled;
    }

    // --- Preview/Standard Mode toggle helpers ---
    /**
     * 画面モードに応じたトグルボタンのラベルを返す
     * @returns {string} ボタンに表示する文字列
     */
    function getPreviewToggleButtonLabel() {
        return isPreviewScreenMode() ? getLabel('button.standardMode') : getLabel('button.previewMode');
    }

    /**
     * トグルボタンのラベルを現在の画面モードに合わせて更新する
     * @param {Button} button 対象のボタン
     * @returns {void}
     */
    function updatePreviewToggleButtonLabel(button) {
        if (!button) return;
        button.text = getPreviewToggleButtonLabel();
    }

    /**
     * 現在プレビュー表示になっているかを判定する
     * @returns {boolean} プレビュー表示なら true
     */
    function isPreviewScreenMode() {
        try {
            return app.activeWindow && app.activeWindow.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE;
        } catch (e) {
            return false;
        }
    }

    /**
     * 標準表示とプレビュー表示を切り替える
     * @returns {void}
     */
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

    /**
     * 選択中のスウォッチを取得する
     * @param {object} ui UI オブジェクト
     * @returns {Swatch|null} スウォッチ。取得できない場合は null
     */
    function getSelectedSwatch(ui) {
        return getSwatchByName(getSelectedColorName(ui));
    }

    // =========================================
    // スウォッチUIヘルパー / Swatch UI helpers
    // =========================================
    /**
     * 色見本つきのスウォッチ選択コントロールを作る
     * @param {object} parent 追加先のコンテナ
     * @param {Array<object>} swatchEntries スウォッチ一覧
     * @param {number} defaultIndex 既定で選ぶ位置
     * @returns {{previewBox: Group, dropdown: DropDownList}} 生成したコントロール
     */
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

    /**
     * スウォッチの色見本を表示する枠を作る
     * @param {object} parent 追加先のコンテナ
     * @returns {Group} 色見本用のグループ
     */
    function createSwatchPreviewBox(parent) {
        var previewBox = parent.add("group");
        previewBox.preferredSize = [18, 18];
        previewBox.minimumSize = [18, 18];
        previewBox.maximumSize = [18, 18];
        return previewBox;
    }

    /**
     * スウォッチ選択のドロップダウンを作る
     * @param {object} parent 追加先のコンテナ
     * @param {Array<object>} swatchEntries スウォッチ一覧
     * @param {number} defaultIndex 既定で選ぶ位置
     * @returns {DropDownList} 生成したドロップダウン
     */
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

    /**
     * 選択中のスウォッチに合わせて色見本を塗り直す
     * @param {Group} previewBox 色見本のグループ
     * @param {DropDownList} dropdown カラーのドロップダウン
     * @param {Window} dlg 対象のダイアログ
     * @returns {void}
     */
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

    /**
     * 名前からスウォッチを取得する
     * @param {string} swatchName スウォッチ名
     * @returns {Swatch|null} スウォッチ。見つからない場合は null
     */
    function getSwatchByName(swatchName) {
        try {
            if (!swatchName) return null;
            return app.activeDocument.swatches.itemByName(String(swatchName));
        } catch (e) {
            return null;
        }
    }

    /**
     * 色見本の描画に使う RGB 値を取得する
     * @param {Swatch} swatch 対象のスウォッチ
     * @returns {Array<number>|null} RGB 値の配列。取得できない場合は null
     */
    function getSwatchPreviewRGBAFromSwatch(swatch) {
        var rgb = convertSwatchToPreviewRGB(swatch);
        return [rgb[0], rgb[1], rgb[2], 1];
    }

    /**
     * スウォッチのカラー値を表示用の RGB に変換する
     * @param {Swatch} swatch 対象のスウォッチ
     * @returns {Array<number>|null} RGB 値の配列。変換できない場合は null
     */
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

    /**
     * 白に近いスウォッチかどうかを判定する
     * @param {Swatch} swatch 対象のスウォッチ
     * @returns {boolean} 白に近ければ true
     */
    function isWhiteLikeSwatch(swatch) {
        var name = swatch && swatch.name != null ? String(swatch.name) : "";
        return isNoneSwatchName(name) || isPaperSwatchName(name);
    }

    /**
     * 黒に近いスウォッチかどうかを判定する
     * @param {Swatch} swatch 対象のスウォッチ
     * @returns {boolean} 黒に近ければ true
     */
    function isBlackLikeSwatch(swatch) {
        var name = swatch && swatch.name != null ? String(swatch.name) : "";
        return isRegistrationSwatchName(name) || isBlackSwatchName(name);
    }

    /**
     * ドキュメントの線幅単位を取得する
     * @returns {MeasurementUnits|null} 線幅の単位。取得できない場合は null
     */
    function getCurrentMeasurementUnit() {
        try {
            return app.activeDocument.viewPreferences.strokeMeasurementUnits;
        } catch (e) {
            return MeasurementUnits.POINTS;
        }
    }

    /**
     * 線幅単位の表示ラベルを取得する
     * @returns {string} 単位のラベル
     */
    function getCurrentLineWidthUnitLabel() {
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
                return "pica";
            case MeasurementUnits.Q:
                return "Q";
            default:
                return "pt";
        }
    }

    /**
     * 線幅入力欄の初期値を取得する
     * @returns {string} 初期値の文字列
     */
    function getDefaultLineWidthText() {
        return getCurrentMeasurementUnit() === MeasurementUnits.POINTS ? "0.25" : "0.1";
    }

    /**
     * 線幅に付ける単位のサフィックスを取得する
     * @returns {string} 単位のサフィックス
     */
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

    /**
     * 線幅の文字列に一致するプリセットを選択状態にする
     * @param {object} target プリセットのラジオボタンをまとめたオブジェクト
     * @param {string} textValue 線幅の文字列
     * @returns {void}
     */
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

    // =========================================
    // 値変換 / Value conversion
    // =========================================
    /**
     * 線幅の入力文字列を単位付きの値に変換する
     * @param {string} text 線幅の入力文字列
     * @returns {string|number|null} 適用できる線幅。無効な場合は null
     */
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
    /**
     * 単位付きの線幅から数値部分を取り出す
     * @param {string|number} weight 線幅
     * @returns {number} 線幅の数値
     */
    function extractLineWeightNumber(weight) {
        if (typeof weight === "number") return weight;
        if (typeof weight === "string") return parseFloat(weight);
        return NaN;
    }

    /**
     * 線幅として有効な値かどうかを判定する
     * @param {string|number} weight 線幅
     * @returns {boolean} 有効なら true
     */
    function isValidLineWeight(weight) {
        var numericWeight = extractLineWeightNumber(weight);
        return !isNaN(numericWeight) && numericWeight >= 0;
    }

    // =========================================
    // 選択取得 / Selection
    // =========================================
    /**
     * 現在の選択から対象の表セルを取得する
     * @returns {Array<Cell>} 選択されたセルの配列
     */
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

    /**
     * 実行前の選択状態を控えておく
     * @param {Array} selectionItems 現在の選択
     * @returns {Array} 復元用の選択情報
     */
    function snapshotSelection(selectionItems) {
        var result = [];
        var i;

        if (!selectionItems || selectionItems.length == null) return result;

        for (i = 0; i < selectionItems.length; i++) {
            try {
                result.push(selectionItems[i]);
            } catch (e) { }
        }

        return result;
    }

    /**
     * 控えておいた選択状態を復元する
     * @param {Array} selectionItems 復元用の選択情報
     * @returns {void}
     */
    function restoreSelection(selectionItems) {
        var restorable = [];
        var i, item;

        if (!selectionItems || selectionItems.length === 0) return;

        for (i = 0; i < selectionItems.length; i++) {
            item = selectionItems[i];
            try {
                if (item && item.isValid !== false) {
                    restorable.push(item);
                }
            } catch (e) { }
        }

        if (restorable.length === 0) return;

        try {
            app.select(restorable);
        } catch (e) {
            try {
                app.selection = restorable;
            } catch (e2) { }
        }
    }

    /**
     * 重複除外のためのセル識別キーを作る
     * @param {Cell} cell 対象のセル
     * @returns {string} 識別キー
     */
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

    /**
     * 選択オブジェクトから表セルの配列を取り出す
     * @param {Array} sel 選択オブジェクトの配列
     * @returns {Array<Cell>} 表セルの配列
     */
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
    /**
     * 指定したモードで罫線を適用する
     * @param {Array<Cell>} cells 対象のセル
     * @param {string} mode 適用モード
     * @param {string|number} weight 線幅
     * @param {boolean} clearFirst 描画前に既存の罫線を消すか
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
    function applyBorders(cells, mode, weight, clearFirst, swatch, tint) {
        if (cells.length === 0) return;

        var bounds = getBounds(cells);

        if (mode === "allOff") {
            applyAllOff(cells);
            return;
        }

        if (mode === "all") {
            applyAll(cells, weight, swatch, tint);
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
            applyHeaderRow(cells, bounds, weight, clearFirst, swatch, tint);
            return;
        }

        if (mode === "headerColumn") {
            applyHeaderColumn(cells, bounds, weight, clearFirst, swatch, tint);
            return;
        }

        if (mode === "clearLeftRight") {
            applyClearLeftRight(cells);
            return;
        }
    }

    /**
     * 見出し行として上下の罫線を引く
     * @param {Array<Cell>} cells 対象のセル
     * @param {object} bounds 選択範囲の境界情報
     * @param {string|number} weight 線幅
     * @param {boolean} clearFirst 描画前に既存の罫線を消すか
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
    function applyHeaderRow(cells, bounds, weight, clearFirst, swatch, tint) {
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
                setCellEdgeTints(cell, tint, tint, null, null);
            }

            if (isLastRow) {
                cell.bottomEdgeStrokeWeight = weight;
                setCellEdgeColors(cell, null, swatch, null, null);
                setCellEdgeTints(cell, null, tint, null, null);
            }
        }
    }

    /**
     * すべての罫線を消去する
     * @param {Array<Cell>} cells 対象のセル
     * @returns {void}
     */
    function applyAllOff(cells) {
        var bounds, rectCells;
        var i, cell;

        if (!cells || cells.length === 0) return;

        bounds = getBounds(cells);
        rectCells = getRectangularCellsFromBounds(cells[0], bounds);

        for (i = 0; i < rectCells.length; i++) {
            cell = rectCells[i];
            clearCellTopEdge(cell);
            clearCellBottomEdge(cell);
            clearCellLeftEdge(cell);
            clearCellRightEdge(cell);
        }
    }

    /**
     * すべての罫線を引く
     * @param {Array<Cell>} cells 対象のセル
     * @param {string|number} weight 線幅
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
    function applyAll(cells, weight, swatch, tint) {
        var bounds, rectCells;
        var i, cell;

        if (!cells || cells.length === 0) return;

        bounds = getBounds(cells);
        rectCells = getRectangularCellsFromBounds(cells[0], bounds);

        for (i = 0; i < rectCells.length; i++) {
            cell = rectCells[i];
            setCellEdges(cell, weight, weight, weight, weight);
            setCellEdgeColors(cell, swatch, swatch, swatch, swatch);
            setCellEdgeTints(cell, tint, tint, tint, tint);
        }
    }

    /**
     * 選択範囲の外周だけに罫線を引く
     * @param {Array<Cell>} cells 対象のセル
     * @param {object} bounds 選択範囲の境界情報
     * @param {string|number} weight 線幅
     * @param {boolean} clearFirst 描画前に既存の罫線を消すか
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
    function applyOuter(cells, bounds, weight, clearFirst, swatch, tint) {
        var rectCells;
        var i, c, edgeFlags;

        if (!cells || cells.length === 0) return;

        rectCells = getRectangularCellsFromBounds(cells[0], bounds);

        if (clearFirst) clearAllEdges(rectCells);

        for (i = 0; i < rectCells.length; i++) {
            c = rectCells[i];
            edgeFlags = getCellEdgeFlags(c, bounds);

            if (edgeFlags.top) { c.topEdgeStrokeWeight = weight; c.topEdgeStrokeColor = swatch; c.topEdgeStrokeTint = tint; }
            if (edgeFlags.bottom) { c.bottomEdgeStrokeWeight = weight; c.bottomEdgeStrokeColor = swatch; c.bottomEdgeStrokeTint = tint; }
            if (edgeFlags.left) { c.leftEdgeStrokeWeight = weight; c.leftEdgeStrokeColor = swatch; c.leftEdgeStrokeTint = tint; }
            if (edgeFlags.right) { c.rightEdgeStrokeWeight = weight; c.rightEdgeStrokeColor = swatch; c.rightEdgeStrokeTint = tint; }
        }
    }

    /**
     * 選択範囲の内側だけに罫線を引く
     * @param {Array<Cell>} cells 対象のセル
     * @param {object} bounds 選択範囲の境界情報
     * @param {string|number} weight 線幅
     * @param {boolean} clearFirst 描画前に既存の罫線を消すか
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
    function applyInnerOnly(cells, bounds, weight, clearFirst, swatch, tint) {
        var rectCells;
        var i, cell;
        var hasBottomNeighbor, hasRightNeighbor;

        if (!cells || cells.length === 0) return;

        rectCells = getRectangularCellsFromBounds(cells[0], bounds);

        if (clearFirst) clearAllEdges(rectCells);

        for (i = 0; i < rectCells.length; i++) {
            cell = rectCells[i];
            hasBottomNeighbor = hasAdjacentSelectedCellOnBottom(cell, rectCells);
            hasRightNeighbor = hasAdjacentSelectedCellOnRight(cell, rectCells);

            if (hasBottomNeighbor) {
                cell.bottomEdgeStrokeWeight = weight;
                cell.bottomEdgeStrokeColor = swatch;
                cell.bottomEdgeStrokeTint = tint;
            }
            if (hasRightNeighbor) {
                cell.rightEdgeStrokeWeight = weight;
                cell.rightEdgeStrokeColor = swatch;
                cell.rightEdgeStrokeTint = tint;
            }
        }
    }

    /**
     * 水平方向の罫線だけを引く
     * @param {Array<Cell>} cells 対象のセル
     * @param {object} bounds 選択範囲の境界情報
     * @param {string|number} weight 線幅
     * @param {boolean} clearFirst 描画前に既存の罫線を消すか
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
    function applyHorizontal(cells, bounds, weight, clearFirst, swatch, tint) {
        var rectCells;
        var i, cell;
        var range;
        var hasBottomNeighbor;
        var isTopBoundary, isBottomBoundary;

        if (!cells || cells.length === 0) return;

        rectCells = getRectangularCellsFromBounds(cells[0], bounds);

        if (clearFirst) clearAllEdges(rectCells);

        for (i = 0; i < rectCells.length; i++) {
            cell = rectCells[i];
            range = getCellRange(cell);
            hasBottomNeighbor = hasAdjacentSelectedCellOnBottom(cell, rectCells);
            isTopBoundary = (range.startRow === bounds.minRow);
            isBottomBoundary = (range.endRow === bounds.maxRow);

            if (isTopBoundary) {
                cell.topEdgeStrokeWeight = weight;
                cell.topEdgeStrokeColor = swatch;
                cell.topEdgeStrokeTint = tint;
            }
            if (hasBottomNeighbor || isBottomBoundary) {
                cell.bottomEdgeStrokeWeight = weight;
                cell.bottomEdgeStrokeColor = swatch;
                cell.bottomEdgeStrokeTint = tint;
            }
        }
    }

    /**
     * 垂直方向の罫線だけを引く
     * @param {Array<Cell>} cells 対象のセル
     * @param {object} bounds 選択範囲の境界情報
     * @param {string|number} weight 線幅
     * @param {boolean} clearFirst 描画前に既存の罫線を消すか
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
    function applyVertical(cells, bounds, weight, clearFirst, swatch, tint) {
        var rectCells;
        var i, cell;
        var range;
        var hasRightNeighbor;
        var isLeftBoundary, isRightBoundary;

        if (!cells || cells.length === 0) return;

        rectCells = getRectangularCellsFromBounds(cells[0], bounds);

        if (clearFirst) clearAllEdges(rectCells);

        for (i = 0; i < rectCells.length; i++) {
            cell = rectCells[i];
            range = getCellRange(cell);
            hasRightNeighbor = hasAdjacentSelectedCellOnRight(cell, rectCells);
            isLeftBoundary = (range.startCol === bounds.minCol);
            isRightBoundary = (range.endCol === bounds.maxCol);

            if (isLeftBoundary) {
                cell.leftEdgeStrokeWeight = weight;
                cell.leftEdgeStrokeColor = swatch;
                cell.leftEdgeStrokeTint = tint;
            }
            if (hasRightNeighbor || isRightBoundary) {
                cell.rightEdgeStrokeWeight = weight;
                cell.rightEdgeStrokeColor = swatch;
                cell.rightEdgeStrokeTint = tint;
            }
        }
    }
    /**
     * 選択ブロックの左端と右端の罫線だけを消去する
     * @param {Array<Cell>} cells 対象のセル
     * @returns {void}
     */
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

    /**
     * 見出し列として左右の罫線を引く
     * @param {Array<Cell>} cells 対象のセル
     * @param {object} bounds 選択範囲の境界情報
     * @param {string|number} weight 線幅
     * @param {boolean} clearFirst 描画前に既存の罫線を消すか
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
    function applyHeaderColumn(cells, bounds, weight, clearFirst, swatch, tint) {
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
                setCellEdgeTints(cell, null, null, tint, tint);
            }

            if (isLastCol) {
                cell.rightEdgeStrokeWeight = weight;
                setCellEdgeColors(cell, null, null, null, swatch);
                setCellEdgeTints(cell, null, null, null, tint);
            }
        }
    }

    /**
     * 左隣に選択セルがあるかを判定する
     * @param {Cell} cell 対象のセル
     * @param {Array<Cell>} cells 選択セルの配列
     * @returns {boolean} 左隣に選択セルがあれば true
     */
    function hasAdjacentSelectedCellOnLeft(cell, cells) {
        return hasAdjacentSelectedCell(cell, cells, -1);
    }

    /**
     * 右隣に選択セルがあるかを判定する
     * @param {Cell} cell 対象のセル
     * @param {Array<Cell>} cells 選択セルの配列
     * @returns {boolean} 右隣に選択セルがあれば true
     */
    function hasAdjacentSelectedCellOnRight(cell, cells) {
        return hasAdjacentSelectedCell(cell, cells, 1);
    }

    /**
     * 指定した方向の隣に選択セルがあるかを判定する
     * @param {Cell} cell 対象のセル
     * @param {Array<Cell>} cells 選択セルの配列
     * @param {string} direction 判定する方向
     * @returns {boolean} 隣に選択セルがあれば true
     */
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

    /**
     * 2 つのセル範囲が縦方向に重なるかを判定する
     * @param {object} a セル範囲
     * @param {object} b セル範囲
     * @returns {boolean} 重なっていれば true
     */
    function rangesOverlapVertically(a, b) {
        return !(a.endRow < b.startRow || b.endRow < a.startRow);
    }

    /**
     * セルの四辺の罫線を消去する
     * @param {Array<Cell>} cells 対象のセル
     * @returns {void}
     */
    function clearAllEdges(cells) {
        var i;
        for (i = 0; i < cells.length; i++) {
            setCellEdges(cells[i], 0, 0, 0, 0);
        }
    }

    /**
     * セルが選択範囲のどの辺に接しているかを求める
     * @param {Cell} cell 対象のセル
     * @param {object} bounds 選択範囲の境界情報
     * @returns {object} 各辺に接しているかを示すフラグ
     */
    function getCellEdgeFlags(cell, bounds) {
        var range = getCellRange(cell);

        return {
            top: range.startRow === bounds.minRow,
            bottom: range.endRow === bounds.maxRow,
            left: range.startCol === bounds.minCol,
            right: range.endCol === bounds.maxCol
        };
    }

    /**
     * 選択セル全体の行・列の範囲を求める
     * @param {Array<Cell>} cells 対象のセル
     * @returns {object} 行と列の範囲
     */
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

    /**
     * 境界情報から矩形範囲のセルを再構築する
     * @param {Cell} seedCell 基準となるセル
     * @param {object} bounds 選択範囲の境界情報
     * @returns {Array<Cell>} 矩形範囲のセル
     */
    function getRectangularCellsFromBounds(seedCell, bounds) {
        var table = getParentTableFromCell(seedCell);
        var result = [];
        var seen = {};
        var rowIndex, colIndex, cell, key;

        if (!table) return result;

        for (rowIndex = bounds.minRow; rowIndex <= bounds.maxRow; rowIndex++) {
            for (colIndex = bounds.minCol; colIndex <= bounds.maxCol; colIndex++) {
                cell = getTableCellCoveringCoordinate(table, rowIndex, colIndex);
                if (!cell) continue;

                key = getCellKey(cell);
                if (!seen[key]) {
                    seen[key] = true;
                    result.push(cell);
                }
            }
        }

        return result;
    }

    /**
     * 指定した行・列を覆っているセルを探す
     * @param {Table} table 対象の表
     * @param {number} rowIndex 行番号
     * @param {number} colIndex 列番号
     * @returns {Cell|null} 該当するセル。見つからない場合は null
     */
    function getTableCellCoveringCoordinate(table, rowIndex, colIndex) {
        var i, cell, range;

        try {
            for (i = 0; i < table.cells.length; i++) {
                cell = table.cells[i];
                range = getCellRange(cell);

                if (rowIndex >= range.startRow && rowIndex <= range.endRow &&
                    colIndex >= range.startCol && colIndex <= range.endCol) {
                    return cell;
                }
            }
        } catch (e) { }

        return null;
    }

    /**
     * 結合を考慮したセルの占有範囲を求める
     * @param {Cell} cell 対象のセル
     * @returns {object} 行と列の占有範囲
     */
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

    /**
     * セルの各辺に線幅を設定する
     * @param {Cell} cell 対象のセル
     * @param {string|number|null} top 上辺の線幅
     * @param {string|number|null} bottom 下辺の線幅
     * @param {string|number|null} left 左辺の線幅
     * @param {string|number|null} right 右辺の線幅
     * @returns {void}
     */
    function setCellEdges(cell, top, bottom, left, right) {
        cell.topEdgeStrokeWeight = top;
        cell.bottomEdgeStrokeWeight = bottom;
        cell.leftEdgeStrokeWeight = left;
        cell.rightEdgeStrokeWeight = right;
    }

    /**
     * セルの各辺にカラーを設定する
     * @param {Cell} cell 対象のセル
     * @param {Swatch|null} top 上辺のカラー
     * @param {Swatch|null} bottom 下辺のカラー
     * @param {Swatch|null} left 左辺のカラー
     * @param {Swatch|null} right 右辺のカラー
     * @returns {void}
     */
    function setCellEdgeColors(cell, top, bottom, left, right) {
        if (top != null) cell.topEdgeStrokeColor = top;
        if (bottom != null) cell.bottomEdgeStrokeColor = bottom;
        if (left != null) cell.leftEdgeStrokeColor = left;
        if (right != null) cell.rightEdgeStrokeColor = right;
    }

    /**
     * セルの各辺に濃淡を設定する
     * @param {Cell} cell 対象のセル
     * @param {number|null} top 上辺の濃淡
     * @param {number|null} bottom 下辺の濃淡
     * @param {number|null} left 左辺の濃淡
     * @param {number|null} right 右辺の濃淡
     * @returns {void}
     */
    function setCellEdgeTints(cell, top, bottom, left, right) {
        if (top != null) cell.topEdgeStrokeTint = top;
        if (bottom != null) cell.bottomEdgeStrokeTint = bottom;
        if (left != null) cell.leftEdgeStrokeTint = left;
        if (right != null) cell.rightEdgeStrokeTint = right;
    }

    /**
     * セルの上辺の罫線を消去する
     * @param {Cell} cell 対象のセル
     * @returns {void}
     */
    function clearCellTopEdge(cell) {
        cell.topEdgeStrokeWeight = 0;
        try {
            cell.topEdgeStrokeColor = NothingEnum.NOTHING;
            cell.topEdgeStrokeTint = 100;
        } catch (e) { }
    }

    /**
     * セルの下辺の罫線を消去する
     * @param {Cell} cell 対象のセル
     * @returns {void}
     */
    function clearCellBottomEdge(cell) {
        cell.bottomEdgeStrokeWeight = 0;
        try {
            cell.bottomEdgeStrokeColor = NothingEnum.NOTHING;
            cell.bottomEdgeStrokeTint = 100;
        } catch (e) { }
    }

    /**
     * セルの左辺の罫線を消去する
     * @param {Cell} cell 対象のセル
     * @returns {void}
     */
    function clearCellLeftEdge(cell) {
        cell.leftEdgeStrokeWeight = 0;
        try {
            cell.leftEdgeStrokeColor = NothingEnum.NOTHING;
            cell.leftEdgeStrokeTint = 100;
        } catch (e) { }
    }

    /**
     * セルの右辺の罫線を消去する
     * @param {Cell} cell 対象のセル
     * @returns {void}
     */
    function clearCellRightEdge(cell) {
        cell.rightEdgeStrokeWeight = 0;
        try {
            cell.rightEdgeStrokeColor = NothingEnum.NOTHING;
            cell.rightEdgeStrokeTint = 100;
        } catch (e) { }
    }

    /**
     * 「なし」のスウォッチ名かどうかを判定する
     * @param {string} name スウォッチ名
     * @returns {boolean} 「なし」なら true
     */
    function isNoneSwatchName(name) {
        return name === "None" || name === "[None]" || name === "なし" || name === "[なし]";
    }

    /**
     * 黒のスウォッチ名かどうかを判定する
     * @param {string} name スウォッチ名
     * @returns {boolean} 黒なら true
     */
    function isBlackSwatchName(name) {
        return name === "Black" || name === "[Black]" || name === "ブラック" || name === "黒";
    }

    /**
     * 紙色のスウォッチ名かどうかを判定する
     * @param {string} name スウォッチ名
     * @returns {boolean} 紙色なら true
     */
    function isPaperSwatchName(name) {
        return name === "Paper" || name === "[Paper]" || name === "紙色" || name === "[紙色]";
    }

    // =========================================
    // Selection helper: check if selection is full table
    // =========================================
    /**
     * 表全体が選択されているかを判定する
     * @param {Array<Cell>} cells 選択セルの配列
     * @returns {boolean} 表全体なら true
     */
    function isFullTableSelection(cells) {
        var table, totalCellCount;
        if (!cells || cells.length === 0) return false;

        table = getParentTableFromCell(cells[0]);
        if (!table) return false;

        totalCellCount = getAllTableCells(table).length;
        return totalCellCount > 0 && cells.length === totalCellCount;
    }

    /**
     * セルが属する表を取得する
     * @param {Cell} cell 対象のセル
     * @returns {Table|null} 表。取得できない場合は null
     */
    function getParentTableFromCell(cell) {
        try {
            return cell.parent;
        } catch (e) {
            return null;
        }
    }

    /**
     * 表に含まれるすべてのセルを取得する
     * @param {Table} table 対象の表
     * @returns {Array<Cell>} セルの配列
     */
    function getAllTableCells(table) {
        var result = [];
        var i;
        try {
            for (i = 0; i < table.cells.length; i++) {
                result.push(table.cells[i]);
            }
        } catch (e) { }
        return result;
    }

    /**
     * 上隣に選択セルがあるかを判定する
     * @param {Cell} cell 対象のセル
     * @param {Array<Cell>} cells 選択セルの配列
     * @returns {boolean} 上隣に選択セルがあれば true
     */
    function hasAdjacentSelectedCellOnTop(cell, cells) {
        var baseRange = getCellRange(cell);
        var i, other, otherRange;
        for (i = 0; i < cells.length; i++) {
            other = cells[i];
            if (other === cell) continue;
            otherRange = getCellRange(other);
            if (!rangesOverlapHorizontally(baseRange, otherRange)) continue;
            if (otherRange.endRow + 1 === baseRange.startRow) return true;
        }
        return false;
    }

    /**
     * 下隣に選択セルがあるかを判定する
     * @param {Cell} cell 対象のセル
     * @param {Array<Cell>} cells 選択セルの配列
     * @returns {boolean} 下隣に選択セルがあれば true
     */
    function hasAdjacentSelectedCellOnBottom(cell, cells) {
        var baseRange = getCellRange(cell);
        var i, other, otherRange;
        for (i = 0; i < cells.length; i++) {
            other = cells[i];
            if (other === cell) continue;
            otherRange = getCellRange(other);
            if (!rangesOverlapHorizontally(baseRange, otherRange)) continue;
            if (baseRange.endRow + 1 === otherRange.startRow) return true;
        }
        return false;
    }

    /**
     * 2 つのセル範囲が横方向に重なるかを判定する
     * @param {object} a セル範囲
     * @param {object} b セル範囲
     * @returns {boolean} 重なっていれば true
     */
    function rangesOverlapHorizontally(a, b) {
        return !(a.endCol < b.startCol || b.endCol < a.startCol);
    }
})();