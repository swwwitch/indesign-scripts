#target "InDesign"

/*
 * =========================================
 * スクリプトの概要（日本語）
 * =========================================
 * 選択中の表セルに対して、罫線の描画・消去を切り替えながら適用できるスクリプトです。
 * 「すべて」「境界線のみ」「内部のみ」「水平線のみ」「垂直線のみ」「上下端」「左右端」「左右の罫線を消去」「すべて消去」に対応しています。
 * 左カラムでは適用範囲と適用オプションを設定し、右カラムではスタイルを設定します。
 * 線幅はドキュメントの線幅単位の環境設定に合わせて入力でき、UI表示・入力値・適用時の単位指定を一致させています。
 * プリセットから「なし」「0.1」「0.2」「0.25」「0.35」「0.5」を素早く選択できます。
 * カラーはドキュメントのスウォッチから選択でき、選択したスウォッチが罫線色として反映されます。
 * 結合セルはセル範囲として扱い、境界判定や重複除外に反映します。
 * 常にプレビューを確認しながら適用でき、「適用前に消去」をOFFにすると既存の罫線を残したまま上書きできます。
 * ボタンエリア左の［プレビュー／標準モード］ボタンで画面表示モードを切り替えられます。
 * =========================================
 * Script Overview (English)
 * =========================================
 * This script lets you apply or remove table borders on selected cells using a range of border modes.
 * Available modes include All, Outer Borders, Inner Borders, Horizontal Borders, Vertical Borders, Top & Bottom, Left & Right, Remove Side Borders, and Clear All.
 * The left column contains border scope and apply options, while the right column contains style settings.
 * Stroke weight follows the document's measurement units so the UI, input values, and applied results stay consistent.
 * Preset buttons let you quickly choose common stroke weights.
 * You can choose a color from the document swatches and apply it to the borders.
 * Merged cells are treated as cell ranges so boundary detection and duplicate handling remain accurate.
 * The result is always previewed interactively before applying. When "Clear Before Apply" is off, new borders are applied without removing existing ones.
 * Use the "Preview / Standard" button to toggle the screen display mode.
 */

var SCRIPT_VERSION = "v1.5.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "罫線・塗りの調整", en: "Borders & Fill" },

    modePanel: { ja: "適用範囲", en: "Border Scope" },
    stylePanel: { ja: "スタイル", en: "Style" },
    drawingOptionsPanel: { ja: "適用オプション", en: "Apply Options" },
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

    swatchBlack: { ja: "黒", en: "Black" },
    swatchPaper: { ja: "紙色", en: "Paper" },
    swatchNone: { ja: "なし", en: "None" },

    previewMode: { ja: "プレビュー", en: "Preview" },
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
    tipAllOff: { ja: "ショートカット: C", en: "Shortcut: C" },

    tabBorder: { ja: "罫線", en: "Borders" },
    tabFill: { ja: "塗り", en: "Cell Fill" },

    fillOddPanel: { ja: "奇数行", en: "Odd Rows" },
    fillEvenPanel: { ja: "偶数行", en: "Even Rows" },
    fillColorLabel: { ja: "カラー：", en: "Color:" },
    fillTintLabel: { ja: "濃淡：", en: "Tint:" },
    fillSwapButton: { ja: "交換", en: "Swap" },
    fillSkipHeader: { ja: "ヘッダー行を無視", en: "Skip Header Row" },
    fillSkipFooter: { ja: "フッター行を無視", en: "Skip Footer Row" },

    undoFillPreview: { ja: "塗りプレビュー", en: "Fill Preview" },
    undoFillApply: { ja: "塗りの設定", en: "Apply Fill Colors" }
};

function L(key) {
    var v = LABELS[key];
    if (v == null) return key;
    if (typeof v === "string") return v;
    return v[lang] || v.en || key;
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

        var tabbedPanel = dlg.add("tabbedpanel");
        tabbedPanel.alignChildren = ["fill", "fill"];
        tabbedPanel.alignment = ["fill", "fill"];

        var tabBorder = tabbedPanel.add("tab", undefined, L('tabBorder'));
        tabBorder.orientation = "column";
        tabBorder.alignChildren = "fill";
        tabBorder.margins = [15, 20, 5, 10];

        var tabFill = tabbedPanel.add("tab", undefined, L('tabFill'));
        tabFill.orientation = "column";
        tabFill.alignChildren = "fill";
        tabFill.margins = [15, 20, 5, 10];

        var settingsColumns = tabBorder.add("group");
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


        rbAll.value = true;

        // カラープレビューの初期表示
        updateSwatchPreview(colorPreviewBox, colorDropdown, dlg);

        // =========================================
        // 塗りタブの内容 / Fill tab content
        // =========================================
        var fillDefaults = getDefaultFillSettings(state.cells);
        var fillSwatchEntries = getSwatchEntries();

        var fillMainGroup = tabFill.add("group");
        fillMainGroup.orientation = "row";
        fillMainGroup.alignChildren = "top";
        fillMainGroup.spacing = 15;

        // 奇数行パネル / Odd rows panel
        var fillOddPanel = fillMainGroup.add("panel", undefined, L('fillOddPanel'));
        fillOddPanel.orientation = "column";
        fillOddPanel.alignChildren = "left";
        fillOddPanel.spacing = 10;
        fillOddPanel.margins = [15, 20, 15, 10];

        var fillOddColorGroup = fillOddPanel.add("group");
        fillOddColorGroup.orientation = "column";
        fillOddColorGroup.alignChildren = "left";
        fillOddColorGroup.margins = [0, 10, 0, 10];
        var fillOddColorLabelRow = fillOddColorGroup.add("group");
        fillOddColorLabelRow.orientation = "row";
        fillOddColorLabelRow.alignChildren = ["left", "center"];
        fillOddColorLabelRow.add("statictext", undefined, L('fillColorLabel'));
        var fillOddColorPreviewBox = createSwatchPreviewBox(fillOddColorLabelRow);
        var fillOddColorDropdown = createSwatchDropdown(fillOddColorGroup, fillSwatchEntries, getFillDefaultColorIndex(fillSwatchEntries, fillDefaults.oddColorName));

        var fillOddTintGroup = fillOddPanel.add("group");
        fillOddTintGroup.orientation = "column";
        fillOddTintGroup.alignChildren = "left";
        fillOddTintGroup.margins = [0, 10, 0, 10];
        var fillOddTintLabelRow = fillOddTintGroup.add("group");
        fillOddTintLabelRow.orientation = "row";
        fillOddTintLabelRow.alignChildren = ["left", "center"];
        fillOddTintLabelRow.add("statictext", undefined, L('fillTintLabel'));
        var fillOddTintText = fillOddTintLabelRow.add("edittext", undefined, String(fillDefaults.oddTint));
        fillOddTintText.preferredSize.width = 50;
        var fillOddSlider = fillOddTintGroup.add("slider", undefined, fillDefaults.oddTint, 0, 100);
        fillOddSlider.preferredSize.width = 150;

        // 偶数行パネル / Even rows panel
        var fillEvenPanel = fillMainGroup.add("panel", undefined, L('fillEvenPanel'));
        fillEvenPanel.orientation = "column";
        fillEvenPanel.alignChildren = "left";
        fillEvenPanel.spacing = 10;
        fillEvenPanel.margins = [15, 20, 15, 10];

        var fillEvenColorGroup = fillEvenPanel.add("group");
        fillEvenColorGroup.orientation = "column";
        fillEvenColorGroup.alignChildren = "left";
        fillEvenColorGroup.margins = [0, 10, 0, 10];
        var fillEvenColorLabelRow = fillEvenColorGroup.add("group");
        fillEvenColorLabelRow.orientation = "row";
        fillEvenColorLabelRow.alignChildren = ["left", "center"];
        fillEvenColorLabelRow.add("statictext", undefined, L('fillColorLabel'));
        var fillEvenColorPreviewBox = createSwatchPreviewBox(fillEvenColorLabelRow);
        var fillEvenColorDropdown = createSwatchDropdown(fillEvenColorGroup, fillSwatchEntries, getFillDefaultColorIndex(fillSwatchEntries, fillDefaults.evenColorName));

        var fillEvenTintGroup = fillEvenPanel.add("group");
        fillEvenTintGroup.orientation = "column";
        fillEvenTintGroup.alignChildren = "left";
        fillEvenTintGroup.margins = [0, 10, 0, 10];
        var fillEvenTintLabelRow = fillEvenTintGroup.add("group");
        fillEvenTintLabelRow.orientation = "row";
        fillEvenTintLabelRow.alignChildren = ["left", "center"];
        fillEvenTintLabelRow.add("statictext", undefined, L('fillTintLabel'));
        var fillEvenTintText = fillEvenTintLabelRow.add("edittext", undefined, String(fillDefaults.evenTint));
        fillEvenTintText.preferredSize.width = 50;
        var fillEvenSlider = fillEvenTintGroup.add("slider", undefined, fillDefaults.evenTint, 0, 100);
        fillEvenSlider.preferredSize.width = 150;

        // 交換チェックボックス / Swap checkbox
        var fillSwapWrapper = tabFill.add("group");
        fillSwapWrapper.orientation = "column";
        fillSwapWrapper.alignChildren = ["left", "top"];
        fillSwapWrapper.alignment = ["left", "top"];
        fillSwapWrapper.margins = [0, 10, 0, 0];

        var fillSwapCheckbox = fillSwapWrapper.add("checkbox", undefined, L('fillSwapButton'));
        var fillSkipHeaderCheckbox = fillSwapWrapper.add("checkbox", undefined, L('fillSkipHeader'));
        var fillSkipFooterCheckbox = fillSwapWrapper.add("checkbox", undefined, L('fillSkipFooter'));

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

            cbClearFirst: cbClearFirst,
            btnCancel: btnCancel,
            btnOk: btnOk,
            drawButtons: [rbAll, rbOuter, rbInnerOnly, rbHorzOnly, rbVertOnly, rbHeaderRow, rbHeaderColumn, rbClearLeftRight, rbAllOff],
            tabbedPanel: tabbedPanel,
            tabBorder: tabBorder,
            tabFill: tabFill,
            fillOddColorDropdown: fillOddColorDropdown,
            fillOddColorPreviewBox: fillOddColorPreviewBox,
            fillEvenColorDropdown: fillEvenColorDropdown,
            fillEvenColorPreviewBox: fillEvenColorPreviewBox,
            fillOddSlider: fillOddSlider,
            fillOddTintText: fillOddTintText,
            fillEvenSlider: fillEvenSlider,
            fillEvenTintText: fillEvenTintText,
            fillSwapCheckbox: fillSwapCheckbox,
            fillSkipHeaderCheckbox: fillSkipHeaderCheckbox,
            fillSkipFooterCheckbox: fillSkipFooterCheckbox,
            btnPreviewToggle: btnPreviewToggle
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


        ui.weightInput.onChange = function () {
            syncWeightPresetFromInput(ui);
            doPreview(ui, state);
        };

        ui.colorDropdown.onChange = function () {
            updateSwatchPreview(ui.colorPreviewBox, ui.colorDropdown, ui.dlg);
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

        addDrawingOptionKeyHandler(ui.dlg, ui, state);
        addModeShortcutKeyHandler(ui.dlg, ui, state);

        // 塗りタブ イベント / Fill tab events
        ui.tabbedPanel.onChange = function () {
            doPreview(ui, state);
        };

        ui.fillOddColorDropdown.onChange = function () {
            updateSwatchPreview(ui.fillOddColorPreviewBox, ui.fillOddColorDropdown, ui.dlg);
            updateFillTintUI(ui, true);
            doPreview(ui, state);
        };
        ui.fillEvenColorDropdown.onChange = function () {
            updateSwatchPreview(ui.fillEvenColorPreviewBox, ui.fillEvenColorDropdown, ui.dlg);
            updateFillTintUI(ui, false);
            doPreview(ui, state);
        };

        ui.fillOddSlider.onChanging = function () {
            var v = adjustFillTintValue(this.value);
            this.value = v;
            ui.fillOddTintText.text = String(v);
            doPreview(ui, state);
        };
        ui.fillOddTintText.onChange = function () {
            var v = adjustFillTintValue(parseFloat(this.text));
            this.text = String(v);
            ui.fillOddSlider.value = v;
            doPreview(ui, state);
        };

        ui.fillEvenSlider.onChanging = function () {
            var v = adjustFillTintValue(this.value);
            this.value = v;
            ui.fillEvenTintText.text = String(v);
            doPreview(ui, state);
        };
        ui.fillEvenTintText.onChange = function () {
            var v = adjustFillTintValue(parseFloat(this.text));
            this.text = String(v);
            ui.fillEvenSlider.value = v;
            doPreview(ui, state);
        };

        ui.fillSwapCheckbox.onClick = function () {
            doPreview(ui, state);
        };
        ui.fillSkipHeaderCheckbox.onClick = function () {
            doPreview(ui, state);
        };
        ui.fillSkipFooterCheckbox.onClick = function () {
            doPreview(ui, state);
        };

        // 初期状態で濃淡UIを更新
        updateFillTintUI(ui, true);
        updateFillTintUI(ui, false);

        // 塗りカラープレビューの初期表示
        updateSwatchPreview(ui.fillOddColorPreviewBox, ui.fillOddColorDropdown, ui.dlg);
        updateSwatchPreview(ui.fillEvenColorPreviewBox, ui.fillEvenColorDropdown, ui.dlg);

        ui.dlg.onShow = function () {
            doPreview(ui, state);
        };

        ui.btnPreviewToggle.onClick = function () {
            togglePreviewScreenMode();
            updatePreviewToggleButtonLabel(ui.btnPreviewToggle);
        };
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
    function isFillTabActive(ui) {
        return ui.tabbedPanel.selection === ui.tabFill;
    }

    function doPreview(ui, state) {
        clearPreview(state);

        if (isFillTabActive(ui)) {
            var fillSettings = getFillSettingsFromUI(ui);
            try {
                app.doScript(function () {
                    applyFillZebra(state.cells, fillSettings.oddColorName, fillSettings.evenColorName, fillSettings.oddTint, fillSettings.evenTint, fillSettings.skipHeader, fillSettings.skipFooter);
                }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, L('undoFillPreview'));
                state.previewed = true;
                app.activeDocument.recompose();
            } catch (e) {
                state.previewed = false;
            }
            return;
        }

        var weight = parseLineWeight(getSelectedWeightText(ui));
        if (!isValidLineWeight(weight)) {
            return;
        }

        var swatch = getSelectedSwatch(ui);
        if (!swatch) {
            return;
        }

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
        if (state.previewed) {
            clearPreview(state);
        }

        if (isFillTabActive(ui)) {
            var fillSettings = getFillSettingsFromUI(ui);
            app.doScript(function () {
                applyFillZebra(state.cells, fillSettings.oddColorName, fillSettings.evenColorName, fillSettings.oddTint, fillSettings.evenTint, fillSettings.skipHeader, fillSettings.skipFooter);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, L('undoFillApply'));
            return;
        }

        var mode = getMode(ui);
        var weight = parseLineWeight(getSelectedWeightText(ui));
        var swatch = getSelectedSwatch(ui);

        if (!isValidLineWeight(weight)) {
            alert(L('alertWeight'));
            return;
        }

        if (!swatch) return;

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

    function getFillDefaultColorIndex(swatchEntries, colorName) {
        var i;
        for (i = 0; i < swatchEntries.length; i++) {
            if (swatchEntries[i].actualName === colorName) return i;
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
    // 塗りヘルパー / Fill helpers
    // =========================================
    function getDefaultFillSettings(cells) {
        var result = { oddColorName: "Black", evenColorName: "Black", oddTint: 100, evenTint: 100 };
        if (!cells || cells.length === 0) return result;

        try {
            var comboCounts = {};
            var combos = [];
            var i, c, colorName, tint, key, k;

            for (i = 0; i < cells.length; i++) {
                c = cells[i];
                colorName = (c.fillColor && c.fillColor.name) ? c.fillColor.name : "Black";
                tint = (!isNaN(c.fillTint) && c.fillTint >= 0) ? c.fillTint : 100;
                key = colorName + "||" + tint;
                if (!comboCounts[key]) {
                    comboCounts[key] = { count: 0, colorName: colorName, tint: tint };
                }
                comboCounts[key].count++;
            }
            for (k in comboCounts) {
                combos.push(comboCounts[k]);
            }
            combos.sort(function (a, b) { return b.count - a.count; });

            if (combos.length >= 1) {
                result.oddColorName = combos[0].colorName;
                result.oddTint = combos[0].tint;
            }
            if (combos.length >= 2) {
                result.evenColorName = combos[1].colorName;
                result.evenTint = combos[1].tint;
            } else {
                result.evenColorName = result.oddColorName;
                result.evenTint = result.oddTint;
            }
        } catch (e) { }
        return result;
    }

    function getSelectedFillColorName(dropdown) {
        return getSelectedSwatchNameFromDropdown(dropdown);
    }

    function getFillSettingsFromUI(ui) {
        var oddColorName = getSelectedFillColorName(ui.fillOddColorDropdown);
        var evenColorName = getSelectedFillColorName(ui.fillEvenColorDropdown);
        var oddTint = clampTintValue(parseFloat(ui.fillOddTintText.text));
        var evenTint = clampTintValue(parseFloat(ui.fillEvenTintText.text));
        var skipHeader = !!(ui.fillSkipHeaderCheckbox && ui.fillSkipHeaderCheckbox.value);
        var skipFooter = !!(ui.fillSkipFooterCheckbox && ui.fillSkipFooterCheckbox.value);
        if (ui.fillSwapCheckbox && ui.fillSwapCheckbox.value) {
            return { oddColorName: evenColorName, evenColorName: oddColorName, oddTint: evenTint, evenTint: oddTint, skipHeader: skipHeader, skipFooter: skipFooter };
        }
        return { oddColorName: oddColorName, evenColorName: evenColorName, oddTint: oddTint, evenTint: evenTint, skipHeader: skipHeader, skipFooter: skipFooter };
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

    function updateFillTintUI(ui, isOdd) {
        var name = isOdd
            ? getSelectedFillColorName(ui.fillOddColorDropdown)
            : getSelectedFillColorName(ui.fillEvenColorDropdown);
        var disabled = isNoneSwatchName(name) || isPaperSwatchName(name);

        if (isOdd) {
            ui.fillOddSlider.enabled = !disabled;
            ui.fillOddTintText.enabled = !disabled;
        } else {
            ui.fillEvenSlider.enabled = !disabled;
            ui.fillEvenTintText.enabled = !disabled;
        }
    }

    function applyFillZebra(cells, oddColorName, evenColorName, oddTint, evenTint, skipHeader, skipFooter) {
        if (!cells || cells.length === 0) return;

        var doc = app.activeDocument;
        var oddSwatch = doc.swatches.item(oddColorName);
        var evenSwatch = doc.swatches.item(evenColorName);

        var rowIndexMap = {};
        var selectedRowIndices = [];
        var i, rIndex, n, cell, localRowOrder, rowType;

        for (i = 0; i < cells.length; i++) {
            rowType = null;
            try { rowType = cells[i].parentRow.rowType; } catch (e) { }
            if (skipHeader && rowType === RowTypes.HEADER_ROW) continue;
            if (skipFooter && rowType === RowTypes.FOOTER_ROW) continue;
            rIndex = cells[i].parentRow.index;
            if (rowIndexMap[rIndex] === undefined) {
                rowIndexMap[rIndex] = selectedRowIndices.length;
                selectedRowIndices.push(rIndex);
            }
        }

        for (n = 0; n < cells.length; n++) {
            cell = cells[n];
            rowType = null;
            try { rowType = cell.parentRow.rowType; } catch (e) { }
            if (skipHeader && rowType === RowTypes.HEADER_ROW) continue;
            if (skipFooter && rowType === RowTypes.FOOTER_ROW) continue;
            localRowOrder = rowIndexMap[cell.parentRow.index];

            if (localRowOrder % 2 === 0) {
                cell.fillColor = oddSwatch;
                if (!isNoneSwatchName(oddColorName) && !isPaperSwatchName(oddColorName)) {
                    try { cell.fillTint = oddTint; } catch (e) { }
                }
            } else {
                cell.fillColor = evenSwatch;
                if (!isNoneSwatchName(evenColorName) && !isPaperSwatchName(evenColorName)) {
                    try { cell.fillTint = evenTint; } catch (e) { }
                }
            }
        }
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
            applyTopAndBottomRows(cells, bounds, weight, clearFirst, swatch);
            return;
        }

        if (mode === "headerColumn") {
            applyLeftAndRightColumns(cells, bounds, weight, clearFirst, swatch);
            return;
        }

        if (mode === "clearLeftRight") {
            applyClearLeftRight(cells);
            return;
        }
    }

    function applyTopAndBottomRows(cells, bounds, weight, clearFirst, swatch) {
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
        clearAllEdges(cells);
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

    function applyLeftAndRightColumns(cells, bounds, weight, clearFirst, swatch) {
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