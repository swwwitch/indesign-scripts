#target indesign

/*
 * IdTableBorderFill.jsx
 *
 * 選択した表セルの罫線を、適用範囲・線幅・カラー・濃淡を指定しながらプレビュー付きで調整します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTableBorderFill";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-13";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableBorderFill.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableBorderFill.md

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

/* 線幅入力欄の文字数、濃淡入力欄とスライダーの幅（px）/ Character width of the weight field, widths of the tint field and slider (px) */
var WEIGHT_INPUT_CHARACTERS = 5;
var TINT_INPUT_WIDTH        = 50;
var TINT_SLIDER_WIDTH       = 150;

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
        title: { ja: "罫線の調整", en: "Borders" }
    },
    panel: {
        mode:  { ja: "適用範囲", en: "Border Scope" },
        style: { ja: "スタイル", en: "Style" }
    },
    radio: {
        all:            { ja: "すべて", en: "All" },
        outer:          { ja: "境界線のみ", en: "Outer Borders" },
        inner:          { ja: "内部のみ", en: "Inner Borders" },
        horizontal:     { ja: "水平線のみ", en: "Horizontal Borders" },
        vertical:       { ja: "垂直線のみ", en: "Vertical Borders" },
        headerRow:      { ja: "上下端", en: "Top & Bottom" },
        headerColumn:   { ja: "左右端", en: "Left & Right" },
        clearLeftRight: { ja: "左右の罫線を消去", en: "Remove Side Borders" },
        allOff:         { ja: "すべて消去", en: "Clear All" },
        weightNone:     { ja: "なし", en: "None" }
    },
    checkbox: {
        clearFirst: { ja: "適用前に消去", en: "Clear Before Apply" }
    },
    field: {
        lineWidth: { ja: "線幅：", en: "Stroke Weight:" },
        color:     { ja: "カラー：", en: "Color:" },
        tint:      { ja: "濃淡：", en: "Tint:" }
    },
    unit: {
        mm:   { ja: "mm", en: "mm" },
        pt:   { ja: "pt", en: "pt" },
        cm:   { ja: "cm", en: "cm" },
        inch: { ja: "in", en: "in" },
        pica: { ja: "pica", en: "pica" },
        q:    { ja: "Q", en: "Q" }
    },
    swatch: {
        black: { ja: "黒", en: "Black" },
        paper: { ja: "紙色", en: "Paper" },
        none:  { ja: "なし", en: "None" }
    },
    button: {
        ok:     { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    tooltip: {
        all:            { ja: "ショートカット: A", en: "Shortcut: A" },
        outer:          { ja: "ショートカット: E", en: "Shortcut: E" },
        inner:          { ja: "ショートカット: I", en: "Shortcut: I" },
        horizontal:     { ja: "ショートカット: H", en: "Shortcut: H" },
        vertical:       { ja: "ショートカット: V", en: "Shortcut: V" },
        headerRow:      { ja: "ショートカット: U", en: "Shortcut: U" },
        headerColumn:   { ja: "ショートカット: L", en: "Shortcut: L" },
        clearLeftRight: { ja: "ショートカット: R", en: "Shortcut: R" },
        allOff:         { ja: "ショートカット: C", en: "Shortcut: C" }
    },
    alert: {
        select: { ja: "表のセルを選択してください。", en: "Please select table cells." },
        weight: { ja: "線幅には0以上の数値を入力してください。", en: "Enter a stroke weight of 0 or greater." }
    },
    undo: {
        preview: { ja: "罫線プレビュー", en: "Border Preview" },
        apply:   { ja: "罫線の設定", en: "Apply Borders" }
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
    var cells = normalizeSelectedCells(getSelectedCellsFromApp());
    if (cells.length === 0) {
        alert(getLabel('alert.select'));
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

    /**
     * 重複を取り除いた選択セルの配列を作る
     * @param {Array<Cell>} cells 選択セルの配列
     * @returns {Array<Cell>} 重複を除いたセルの配列
     */
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

    // =========================================
    // UI構築 / Build UI
    // =========================================
    /**
     * 罫線調整ダイアログを組み立てる
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
        var rbClearLeftRight = panelMode.add("radiobutton", undefined, getLabel('radio.clearLeftRight'));
        rbClearLeftRight.helpTip = getLabel('tooltip.clearLeftRight');
        var rbAllOff = panelMode.add("radiobutton", undefined, getLabel('radio.allOff'));
        rbAllOff.helpTip = getLabel('tooltip.allOff');

        /* 枠線なしのグループで適用オプションを並べる / Apply options sit in a borderless group */
        var panelDrawingOptions = leftColumn.add("group");
        panelDrawingOptions.orientation = "column";
        panelDrawingOptions.alignChildren = "left";
        panelDrawingOptions.alignment = ["fill", "top"];
        panelDrawingOptions.margins = [PANEL_MARGINS[0], 5, 0, 0];

        var cbClearFirst = panelDrawingOptions.add("checkbox", undefined, getLabel('checkbox.clearFirst'));
        cbClearFirst.value = true;

        var rightColumn = settingsColumns.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];
        rightColumn.alignment = ["fill", "top"];
        rightColumn.spacing = PANEL_SPACING;

        var panelStyle = rightColumn.add("panel", undefined, getLabel('panel.style'));
        setupPanel(panelStyle, 8);

        var panelWeight = panelStyle.add("group");
        panelWeight.orientation = "column";
        panelWeight.alignChildren = ["fill", "top"];
        panelWeight.alignment = ["fill", "top"];

        var weightRow = panelWeight.add("group");
        setupRow(weightRow, "left", 8);
        weightRow.alignChildren = ["left", "center"];

        weightRow.add("statictext", undefined, getLabel('field.lineWidth'));
        var weightInput = weightRow.add("edittext", undefined, getDefaultLineWidthText());
        weightInput.characters = WEIGHT_INPUT_CHARACTERS;

        weightRow.add("statictext", undefined, getCurrentLineWidthUnitLabel());

        var weightPresetGroup = panelWeight.add("group");
        weightPresetGroup.orientation = "column";
        weightPresetGroup.alignChildren = ["left", "center"];
        weightPresetGroup.alignment = ["left", "top"];
        weightPresetGroup.spacing = 4;
        weightPresetGroup.margins = [PANEL_MARGINS[0], 5, PANEL_MARGINS[2], 10];

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

        var panelColor = panelStyle.add("group");
        panelColor.orientation = "column";
        panelColor.alignChildren = ["left", "top"];
        panelColor.alignment = ["fill", "top"];

        var colorLabelRow = panelColor.add("group");
        setupRow(colorLabelRow, "left", 6);
        colorLabelRow.alignChildren = ["left", "center"];
        colorLabelRow.add("statictext", undefined, getLabel('field.color'));
        var colorPreviewBox = createSwatchPreviewBox(colorLabelRow);

        var swatchEntries = getSwatchEntries();
        var colorDropdown = createSwatchDropdown(panelColor, swatchEntries, getDefaultColorIndex(swatchEntries));

        var tintGroup = panelColor.add("group");
        tintGroup.orientation = "column";
        tintGroup.alignChildren = "left";
        tintGroup.margins = [0, 10, 0, 10];
        var tintLabelRow = tintGroup.add("group");
        setupRow(tintLabelRow, "left", 6);
        tintLabelRow.alignChildren = ["left", "center"];
        tintLabelRow.add("statictext", undefined, getLabel('field.tint'));

        /* 濃淡の初期値は TINT_DEFAULT / The tint starts at TINT_DEFAULT */
        var tintText = tintLabelRow.add("edittext", undefined, String(TINT_DEFAULT));
        tintText.preferredSize.width = TINT_INPUT_WIDTH;
        tintLabelRow.add("statictext", undefined, "%");
        var tintSlider = tintGroup.add("slider", undefined, TINT_DEFAULT, TINT_MIN, TINT_MAX);
        tintSlider.preferredSize.width = TINT_SLIDER_WIDTH;
        tintGroup.enabled = true;


        rbAll.value = true;

        if (!state.isFullTable) {
            rbHeaderRow.enabled = false;
            rbHeaderColumn.enabled = false;
        }

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

        var btnPreviewToggle = btnLeftGroup.add("button", undefined, getPreviewToggleButtonLabel());
        btnPreviewToggle.alignment = ["left", "center"];

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
    /**
     * ダイアログのコントロールにイベントを結び付ける
     * @param {object} ui buildDialog が返した UI オブジェクト
     * @param {object} state 選択セルやプレビュー状態を保持するオブジェクト
     * @returns {void}
     */
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

    /**
     * 現在の画面モードを安全に取得する
     * @returns {ScreenModeOptions|null} 画面モード。取得できない場合は null
     */
    function getCurrentScreenModeSafe() {
        try {
            return app.activeWindow ? app.activeWindow.screenMode : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * 画面モードを安全に切り替える
     * @param {ScreenModeOptions} mode 設定する画面モード
     * @returns {void}
     */
    function setScreenModeSafe(mode) {
        try {
            if (app.activeWindow && mode != null) {
                app.activeWindow.screenMode = mode;
            }
        } catch (e) { }
    }

    /**
     * 実行前の画面モードに戻す
     * @param {object} state 状態オブジェクト
     * @returns {void}
     */
    function restoreScreenModeFromState(state) {
        if (!state) return;
        if ((state.previewToggleCount % 2) === 0) return;
        setScreenModeSafe(state.initialScreenMode);
    }

    /**
     * 画面モードに応じたトグルボタンのラベルを返す
     * @returns {string} ボタンに表示する文字列
     */
    function getPreviewToggleButtonLabel() {
        return isPreviewScreenMode() ? "プレビュー" : "標準モード";
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
     * 適用範囲のラジオボタンが押されたときにプレビューを更新する
     * @param {object} ui UI オブジェクト
     * @param {object} state 状態オブジェクト
     * @param {object} event クリックイベント
     * @returns {void}
     */
    function onRadioClick(ui, state, event) {
        var ks = ScriptUI.environment.keyboardState;
        var isAlt = !!(ks && ks.altKey);
        if (isAlt && ui && ui.cbClearFirst) {
            ui.cbClearFirst.value = !ui.cbClearFirst.value;
        }
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
     * 適用オプションを切り替えるキー操作を登録する
     * @param {Window} dialog 対象のダイアログ
     * @param {object} ui UI オブジェクト
     * @param {object} state 状態オブジェクト
     * @returns {void}
     */
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

    /**
     * 適用範囲を切り替えるショートカットキーを登録する
     * @param {Window} dialog 対象のダイアログ
     * @param {object} ui UI オブジェクト
     * @param {object} state 状態オブジェクト
     * @returns {void}
     */
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
    /**
     * 現在の設定で罫線のプレビューを描画する
     * @param {object} ui UI オブジェクト
     * @param {object} state 状態オブジェクト
     * @returns {void}
     */
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
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel('undo.preview'));
            state.previewed = true;
            app.activeDocument.recompose();
        } catch (e) {
            state.previewed = false;
            try { $.writeln("[IdTableBorderFill] Preview error: " + e); } catch (_) { }
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
        } catch (e) {
            try { $.writeln("[IdTableBorderFill] Undo preview error: " + e); } catch (_) { }
        }
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
        if (state.previewed) {
            clearPreview(state);
        }

        var mode = getMode(ui);
        var weight = parseLineWeight(getSelectedWeightText(ui));
        var swatch = getSelectedSwatch(ui);
        var tint = getBorderTintFromUI(ui);

        if (!isValidLineWeight(weight)) {
            alert(getLabel('alert.weight'));
            return;
        }

        if (!swatch) return;

        try {
            app.doScript(function () {
                applyBorders(state.cells, mode, weight, ui.cbClearFirst.value, swatch, tint);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel('undo.apply'));
        } catch (e) {
            try { $.writeln("[IdTableBorderFill] Apply error: " + e); } catch (_) { }
            throw e;
        }
    }

    // =========================================
    // UI値の取得 / Read UI values
    // =========================================
    /**
     * 選択中の適用範囲を取得する
     * @param {object} ui UI オブジェクト
     * @returns {string} 適用範囲を表す識別子
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
        var kind = getLocalizedSwatchDisplayKind(name);
        if (kind === "none") return getLabel('swatch.none');
        if (kind === "black") return getLabel('swatch.black');
        if (kind === "paper") return getLabel('swatch.paper');
        return String(name);
    }

    /**
     * スウォッチ名から表示種別（黒・紙色・なし）を判定する
     * @param {string} name スウォッチ名
     * @returns {string} 表示種別を表す識別子
     */
    function getLocalizedSwatchDisplayKind(name) {
        if (isNoneSwatchName(name)) return "none";
        if (isBlackSwatchName(name)) return "black";
        if (isPaperSwatchName(name)) return "paper";
        return "";
    }


    /**
     * 比較しやすいようにスウォッチ名を正規化する
     * @param {string} name スウォッチ名
     * @returns {string} 正規化した名前
     */
    function normalizeSwatchName(name) {
        if (name == null) return "";
        var n = String(name);

        // remove brackets like [Black]
        n = n.replace(/^\[|\]$/g, "");

        // trim and lowercase for comparison
        n = n.replace(/^\s+|\s+$/g, "").toLowerCase();

        return n;
    }

    /**
     * レジストレーションのスウォッチ名かどうかを判定する
     * @param {string} name スウォッチ名
     * @returns {boolean} レジストレーションなら true
     */
    function isRegistrationSwatchName(name) {
        var n = normalizeSwatchName(name);
        return n === "registration" || n === "レジストレーション";
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
     * 選択中のスウォッチを取得する
     * @param {object} ui UI オブジェクト
     * @returns {Swatch|null} スウォッチ。取得できない場合は null
     */
    function getSelectedSwatch(ui) {
        return getSwatchByName(getSelectedColorName(ui));
    }

    /**
     * 入力欄から罫線の濃淡を取得する
     * @param {object} ui UI オブジェクト
     * @returns {number} 濃淡の値
     */
    function getBorderTintFromUI(ui) {
        if (!ui || !ui.tintText) return 100;
        return clampTintValue(parseFloat(ui.tintText.text));
    }
    // =========================================
    // スウォッチUIとプレビュー / Swatch UI and preview helpers
    // =========================================
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

    /**
     * 白として表示すべきスウォッチかどうかを判定する
     * @param {Swatch} swatch 対象のスウォッチ
     * @returns {boolean} 白として扱うなら true
     */
    function isWhitePreviewSwatch(swatch) {
        var kind = getPreviewSwatchKind(swatch);
        return kind === "none" || kind === "paper";
    }

    /**
     * 黒として表示すべきスウォッチかどうかを判定する
     * @param {Swatch} swatch 対象のスウォッチ
     * @returns {boolean} 黒として扱うなら true
     */
    function isBlackPreviewSwatch(swatch) {
        var kind = getPreviewSwatchKind(swatch);
        return kind === "registration" || kind === "black";
    }

    /**
     * 色見本の描画に使う種別を判定する
     * @param {Swatch} swatch 対象のスウォッチ
     * @returns {string} 種別を表す識別子
     */
    function getPreviewSwatchKind(swatch) {
        var name = swatch && swatch.name != null ? String(swatch.name) : "";
        if (isNoneSwatchName(name)) return "none";
        if (isPaperSwatchName(name)) return "paper";
        if (isRegistrationSwatchName(name)) return "registration";
        if (isBlackSwatchName(name)) return "black";
        return "";
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
                return getLabel('unit.mm');
            case MeasurementUnits.POINTS:
                return getLabel('unit.pt');
            case MeasurementUnits.CENTIMETERS:
                return getLabel('unit.cm');
            case MeasurementUnits.INCHES:
                return getLabel('unit.inch');
            case MeasurementUnits.PICAS:
                return getLabel('unit.pica');
            case MeasurementUnits.Q:
                return getLabel('unit.q');
            default:
                return getLabel('unit.pt');
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

    /**
     * 選択中のスウォッチに応じて濃淡コントロールの有効／無効を切り替える
     * @param {object} ui UI オブジェクト
     * @returns {void}
     */
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

    /**
     * 表全体が選択されているかを判定する
     * @param {Array<Cell>} cells 選択セルの配列
     * @returns {boolean} 表全体なら true
     */
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


    /**
     * 濃淡の値を有効範囲に収める
     * @param {number} v 濃淡の値
     * @returns {number} 範囲内に収めた値
     */
    function clampTintValue(v) {
        if (isNaN(v)) v = 100;
        if (v < 0) v = 0;
        if (v > 100) v = 100;
        return v;
    }

    /**
     * 塗りの濃淡として適用できる値に整える
     * @param {number} v 濃淡の値
     * @returns {number} 調整した値
     */
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
    /**
     * 指定した適用範囲で罫線を適用する
     * @param {Array<Cell>} cells 対象のセル
     * @param {string} mode 適用範囲
     * @param {string|number} weight 線幅
     * @param {boolean} clearFirst 適用前に既存の罫線を消すか
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

    /**
     * 選択範囲の上端と下端の行に罫線を引く
     * @param {Array<Cell>} cells 対象のセル
     * @param {object} bounds 選択範囲の境界情報
     * @param {string|number} weight 線幅
     * @param {boolean} clearFirst 適用前に既存の罫線を消すか
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
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

    /**
     * すべての罫線を消去する
     * @param {Array<Cell>} cells 対象のセル
     * @returns {void}
     */
    function applyAllOff(cells) {
        var i;
        for (i = 0; i < cells.length; i++) {
            clearCellEdgesAndSelectedOpposites(cells[i], cells);
        }
    }

    /**
     * すべての罫線を引く
     * @param {Array<Cell>} cells 対象のセル
     * @param {string|number} weight 線幅
     * @param {boolean} clearFirst 適用前に既存の罫線を消すか
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
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

    /**
     * 選択範囲の外周だけに罫線を引く
     * @param {Array<Cell>} cells 対象のセル
     * @param {object} bounds 選択範囲の境界情報
     * @param {string|number} weight 線幅
     * @param {boolean} clearFirst 適用前に既存の罫線を消すか
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
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

    /**
     * 選択範囲の内側だけに罫線を引く
     * @param {Array<Cell>} cells 対象のセル
     * @param {object} bounds 選択範囲の境界情報
     * @param {string|number} weight 線幅
     * @param {boolean} clearFirst 適用前に既存の罫線を消すか
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
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

    /**
     * 水平方向の罫線だけを引く
     * @param {Array<Cell>} cells 対象のセル
     * @param {object} bounds 選択範囲の境界情報
     * @param {string|number} weight 線幅
     * @param {boolean} clearFirst 適用前に既存の罫線を消すか
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
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

    /**
     * 垂直方向の罫線だけを引く
     * @param {Array<Cell>} cells 対象のセル
     * @param {object} bounds 選択範囲の境界情報
     * @param {string|number} weight 線幅
     * @param {boolean} clearFirst 適用前に既存の罫線を消すか
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
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

    /**
     * 選択ブロックの左端と右端の罫線だけを消去する
     * @param {Array<Cell>} cells 対象のセル
     * @param {object} bounds 選択範囲の境界情報
     * @returns {void}
     */
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

    /**
     * 選択範囲の左端と右端の列に罫線を引く
     * @param {Array<Cell>} cells 対象のセル
     * @param {object} bounds 選択範囲の境界情報
     * @param {string|number} weight 線幅
     * @param {boolean} clearFirst 適用前に既存の罫線を消すか
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
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

    
    /**
     * 隣接する選択セルの向かい合う辺も揃えて罫線を引く
     * @param {Cell} cell 対象のセル
     * @param {Array<Cell>} selectedCells 選択セルの配列
     * @param {string} side 対象の辺
     * @param {string|number} weight 線幅
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
    function applyEdgeWithSelectedOpposite(cell, selectedCells, side, weight, swatch, tint) {
        var adjacent = findAdjacentSelectedCell(cell, selectedCells, side);
        setSingleCellEdge(cell, side, weight, swatch, tint);
        if (adjacent) {
            setSingleCellEdge(adjacent, getOppositeSide(side), weight, swatch, tint);
        }
    }

    /**
     * セルの四辺と、隣接する選択セルの向かい合う辺を消去する
     * @param {Cell} cell 対象のセル
     * @param {Array<Cell>} selectedCells 選択セルの配列
     * @returns {void}
     */
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

    /**
     * 指定した辺の隣にある選択セルを探す
     * @param {Cell} cell 対象のセル
     * @param {Array<Cell>} selectedCells 選択セルの配列
     * @param {string} side 対象の辺
     * @returns {Cell|null} 隣接する選択セル。なければ null
     */
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

    /**
     * 向かい合う辺の名前を返す
     * @param {string} side 辺の名前
     * @returns {string} 向かい合う辺の名前
     */
    function getOppositeSide(side) {
        if (side === "top") return "bottom";
        if (side === "bottom") return "top";
        if (side === "left") return "right";
        return "left";
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
     * セルの 1 辺に線幅・カラー・濃淡を設定する
     * @param {Cell} cell 対象のセル
     * @param {string} side 対象の辺
     * @param {string|number} weight 線幅
     * @param {Swatch} swatch 罫線のカラー
     * @param {number} tint 濃淡
     * @returns {void}
     */
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

    /**
     * セルの 1 辺の罫線を消去する
     * @param {Cell} cell 対象のセル
     * @param {string} side 対象の辺
     * @returns {void}
     */
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



    /**
     * セルの四辺の罫線を消去する
     * @param {Array<Cell>} cells 対象のセル
     * @returns {void}
     */
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
     * 「なし」「紙色」以外のスウォッチが指定されている辺だけに濃淡を設定する
     * @param {Cell} cell 対象のセル
     * @param {number|null} topTint 上辺の濃淡
     * @param {number|null} bottomTint 下辺の濃淡
     * @param {number|null} leftTint 左辺の濃淡
     * @param {number|null} rightTint 右辺の濃淡
     * @param {Swatch|null} topSwatch 上辺のカラー
     * @param {Swatch|null} bottomSwatch 下辺のカラー
     * @param {Swatch|null} leftSwatch 左辺のカラー
     * @param {Swatch|null} rightSwatch 右辺のカラー
     * @returns {void}
     */
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

    /**
     * 「なし」のスウォッチ名かどうかを判定する
     * @param {string} name スウォッチ名
     * @returns {boolean} 「なし」なら true
     */
    function isNoneSwatchName(name) {
        var n = normalizeSwatchName(name);
        return n === "none" || n === "なし";
    }

    /**
     * 黒のスウォッチ名かどうかを判定する
     * @param {string} name スウォッチ名
     * @returns {boolean} 黒なら true
     */
    function isBlackSwatchName(name) {
        var n = normalizeSwatchName(name);
        return n === "black" || n === "ブラック" || n === "黒";
    }

    /**
     * 紙色のスウォッチ名かどうかを判定する
     * @param {string} name スウォッチ名
     * @returns {boolean} 紙色なら true
     */
    function isPaperSwatchName(name) {
        var n = normalizeSwatchName(name);
        return n === "paper" || n === "紙色";
    }
})();