#target "InDesign"

/*
 * スクリプトの概要：
 * 選択中の表セルに対して、罫線の描画・消去を切り替えながら適用できるスクリプトです。
 * 「すべて」「境界線のみ」「内部のみ」「水平線のみ」「垂直線のみ」「見出し行」「見出し列」「左右の境界線を消去」「すべて消去」に対応しています。
 *
 * 左カラムではモードと描画オプションを設定し、右カラムでは線幅とカラーを設定します。
 * 線幅はドキュメントの線幅単位の環境設定に合わせて入力でき、UI表示・入力値・適用時の単位指定を一致させています。
 * 線幅は内部で単位付きの値として適用し、ミリメートル環境では 0.1 を 0.1mm、ポイント環境では 0.1 を 0.1pt としてそのまま反映します。
 * プリセットのラジオボタンから「なし」「0.1」「0.2」「0.25」「0.35」「0.5」を素早く選択でき、ポイント単位のときは初期値に 0.25pt を使用します。
 * 線幅入力欄では ↑↓ で 0.1 単位、shift + ↑↓ で 1 単位の増減が可能です。
 *
 * カラーはドキュメントのスウォッチから選択でき、選択したスウォッチは罫線色として反映されます。
 * カラー内に濃淡（Tint）の数値入力とスライダーを備え、デフォルト値 100 を基準に 0〜100 の範囲で罫線の濃淡を調整できます。なし・紙色を選択した場合は濃淡UIを自動的にディム表示します。
 * 表の一部を選択した場合でも、選択範囲の矩形（bounds）からセルを再構築して処理するため、内部罫線・水平線・垂直線・すべて消去が安定して適用されます。結合セルがある場合も、各座標を実際に覆っているセルを探索して再構築します。
 * スウォッチ名はUI表示用に整形され、日本語UIでは Black→黒、Paper→紙色、None→なし と表示されます。
 * Registration / レジストレーションはカラー候補に表示しません。
 *
 * 見出し行は、選択範囲の1行目に上下の罫線を描画し、最終行に下の罫線を描画します。
 * 見出し列は、選択範囲の1列目に左右の罫線を描画し、最終列に右の罫線を描画します。
 * 左右の境界線を消去では、左右に隣接する選択セル間の線を無視し、選択ブロックの左端・右端にある線だけを消去します。
 * 結合セルはセル範囲として扱い、境界判定や重複除外に反映します。
 *
 * ダイアログ表示中は常にプレビューが有効で、「描画前に消去」を OFF にすると既存の罫線を残したまま上書きできます。
 * ボタンエリアには［標準モード］／［プレビュー］のトグルボタンを備え、現在の画面モードに応じてラベルが切り替わります。クリックで標準表示とプレビュー表示を相互に切り替えられます。
 * ダイアログ終了後は、実行前の選択状態を復元します。
 *
 * 主な機能：
 * - 選択中の表セルに対する罫線の描画／消去
 * - 外枠・内部・水平・垂直・見出し行・見出し列・左右の境界線を消去・すべて消去の各モード切り替え
 * - モード切り替え用ショートカットキー対応（A/E/I/H/V/U/L/R/C）
 * - 「描画前に消去」の ON/OFF 切り替え（Mキーで切り替え）
 * - ドキュメントの線幅単位の環境設定に追従した線幅表示・入力・適用
 * - 線幅プリセットのラジオボタン選択
 * - 線幅入力欄でのキー操作による値変更
 * - スウォッチによる罫線カラー指定
 * - 濃淡（Tint）をデフォルト値 100 から数値入力とスライダーで調整（なし／紙色時はディム表示、Shiftで10%刻み）
 * - スウォッチプレビュー表示
 * - 常時プレビューによる確認
 * - ［標準モード］／［プレビュー］トグルボタンによる画面モード切り替え（現在状態に応じてラベルが変化）
 * - 表の部分選択でも矩形再構築により安定した罫線適用（内部／水平／垂直／すべて消去）
 * - 実行前の選択状態を復元
 * - 結合セルを考慮した境界判定と、部分選択時の矩形セル再構築
 * - 日本語／英語UI対応
 */

var SCRIPT_VERSION = "v1.6.5";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "罫線の設定", en: "Border Settings" },
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
    colorPanel: { ja: "カラー", en: "Border Color" },
    tintLabel: { ja: "濃淡：", en: "Tint:" },
    tintSliderTip: { ja: "0〜100", en: "0–100" },
    swatchBlack: { ja: "黒", en: "Black" },
    swatchPaper: { ja: "紙色", en: "Paper" },
    swatchNone: { ja: "なし", en: "None" },
    standardMode: { ja: "標準モード", en: "Standard Mode" },
    previewMode: { ja: "プレビュー", en: "Preview" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    alertSelect: { ja: "表のセルを選択してください。", en: "Please select table cells." },
    alertWeight: { ja: "線幅には0以上の数値を入力してください。", en: "Enter a value of 0 or greater for weight." },
    alertTint: { ja: "濃淡には0〜100の数値を入力してください。", en: "Enter a value between 0 and 100 for tint." },
    undoPreview: { ja: "罫線プレビュー", en: "Border Preview" },
    undoApply: { ja: "罫線の設定", en: "Apply Border Settings" },
    tipAll: { ja: "ショートカット: A / Option+クリックで消去ON/OFF", en: "Shortcut: A / Option-click toggles Clear Existing Borders First" },
    tipOuter: { ja: "ショートカット: E / Option+クリックで消去ON/OFF", en: "Shortcut: E / Option-click toggles Clear Existing Borders First" },
    tipInner: { ja: "ショートカット: I / Option+クリックで消去ON/OFF", en: "Shortcut: I / Option-click toggles Clear Existing Borders First" },
    tipHorizontal: { ja: "ショートカット: H / Option+クリックで消去ON/OFF", en: "Shortcut: H / Option-click toggles Clear Existing Borders First" },
    tipVertical: { ja: "ショートカット: V / Option+クリックで消去ON/OFF", en: "Shortcut: V / Option-click toggles Clear Existing Borders First" },
    tipHeaderRow: { ja: "ショートカット: U / Option+クリックで消去ON/OFF", en: "Shortcut: U / Option-click toggles Clear Existing Borders First" },
    tipHeaderColumn: { ja: "ショートカット: L / Option+クリックで消去ON/OFF", en: "Shortcut: L / Option-click toggles Clear Existing Borders First" },
    tipClearLeftRight: { ja: "ショートカット: R / Option+クリックで消去ON/OFF", en: "Shortcut: R / Option-click toggles Clear Existing Borders First" },
    tipAllOff: { ja: "ショートカット: C / Option+クリックで消去ON/OFF", en: "Shortcut: C / Option-click toggles Clear Existing Borders First" }
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
        rbHeaderRow.enabled = state.headerModesEnabled;
        rbHeaderColumn.enabled = state.headerModesEnabled;
        var rbClearLeftRight = panelMode.add("radiobutton", undefined, L('clearLeftRight'));
        rbClearLeftRight.helpTip = L('tipClearLeftRight');
        var rbAllOff = panelMode.add("radiobutton", undefined, L('allOff'));
        rbAllOff.helpTip = L('tipAllOff');

        // Replace panel with group and add visible label
        var panelDrawingOptions = leftColumn.add("group");
        panelDrawingOptions.orientation = "column";
        panelDrawingOptions.alignChildren = "left";
        panelDrawingOptions.alignment = ["fill", "top"];
        panelDrawingOptions.margins = [15, 10, 15, 10];


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

        var weightInput = weightRow.add("edittext", undefined, getDefaultLineWidthText());
        weightInput.characters = 6;
        weightInput.minimumSize.width = 60;

        weightRow.add("statictext", undefined, getCurrentLineWidthUnitLabel());

        var weightPresetContainer = weightGroup.add("group");
        weightPresetContainer.orientation = "column";
        weightPresetContainer.alignChildren = ["left", "center"];
        weightPresetContainer.alignment = ["fill", "top"];
        weightPresetContainer.margins = [15, 10, 15, 10];

        var weightPresetGroup = weightPresetContainer.add("group");
        weightPresetGroup.orientation = "column";
        weightPresetGroup.alignChildren = ["left", "center"];
        weightPresetGroup.alignment = ["left", "top"];
        weightPresetGroup.spacing = 4;

        var rbWeightNone = weightPresetGroup.add("radiobutton", undefined, lang === "ja" ? "なし" : "None");
        var rbWeight01 = weightPresetGroup.add("radiobutton", undefined, "0.1");
        var rbWeight02 = weightPresetGroup.add("radiobutton", undefined, "0.2");
        var rbWeight025 = weightPresetGroup.add("radiobutton", undefined, "0.25");
        var rbWeight035 = weightPresetGroup.add("radiobutton", undefined, "0.35");
        var rbWeight05 = weightPresetGroup.add("radiobutton", undefined, "0.5");

        syncWeightPresetFromTextValue({
            rbWeightNone: rbWeightNone,
            rbWeight01: rbWeight01,
            rbWeight02: rbWeight02,
            rbWeight025: rbWeight025,
            rbWeight035: rbWeight035,
            rbWeight05: rbWeight05
        }, getDefaultLineWidthText());

        var panelColor = panelStyle.add("panel", undefined, L('colorPanel'));
        panelColor.orientation = "column";
        panelColor.alignChildren = ["left", "top"];
        panelColor.alignment = ["fill", "top"];
        panelColor.margins = [15, 20, 15, 10];

        var swatchEntries = getSwatchEntries();
        var colorPicker = createSwatchDropdownWithPreview(panelColor, swatchEntries, getDefaultColorIndex(swatchEntries));
        var colorPreviewBox = colorPicker.previewBox;
        var colorDropdown = colorPicker.dropdown;

        var panelTint = panelColor.add("group");
        panelTint.orientation = "column";
        panelTint.alignChildren = ["fill", "top"];
        panelTint.alignment = ["fill", "top"];

        var tintRow = panelTint.add("group");
        tintRow.orientation = "row";
        tintRow.alignChildren = ["left", "center"];
        tintRow.alignment = ["fill", "center"];
        tintRow.spacing = 8;

        tintRow.add("statictext", undefined, L('tintLabel'));

        // Tint default is 100 / 濃淡のデフォルト値は100
        var tintInput = tintRow.add("edittext", undefined, "100");
        tintInput.characters = 4;
        tintInput.minimumSize.width = 45;

        // Tint default is 100 / 濃淡のデフォルト値は100
        var tintSlider = panelTint.add("slider", undefined, 100, 0, 100);
        tintSlider.helpTip = L('tintSliderTip');

        rbAll.value = true;
        if (!state.headerModesEnabled) {
            if (rbHeaderRow.value) rbHeaderRow.value = false;
            if (rbHeaderColumn.value) rbHeaderColumn.value = false;
            rbAll.value = true;
        }
        updateTintControlsEnabledState(colorDropdown, tintInput, tintSlider);

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

        var btnStandardMode = btnLeftGroup.add("button", undefined, "");
        updatePreviewToggleButtonLabel(btnStandardMode);

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
        var tint = getSelectedTint(ui);

        if (!isValidLineWeight(weight)) {
            clearPreview(state);
            alert(L('alertWeight'));
            return;
        }

        if (!isValidTint(tint)) {
            clearPreview(state);
            alert(L('alertTint'));
            return;
        }

        if (!swatch) return;

        if (state.previewed) {
            clearPreview(state);
        }

        app.doScript(function () {
            applyBorders(state.cells, mode, weight, ui.cbClearFirst.value, swatch, tint);
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

    function getSelectedTint(ui) {
        if (!ui || !ui.tintInput) return 100;
        return clampTintValue(parseFloat(String(ui.tintInput.text).replace(/^\s+|\s+$/g, "")));
    }

    function clampTintValue(value) {
        if (isNaN(value)) return NaN;
        if (value < 0) return 0;
        if (value > 100) return 100;
        return value;
    }

    function isValidTint(value) {
        return !isNaN(value) && value >= 0 && value <= 100;
    }

    function clampTintInput(editText) {
        var value;
        if (!editText) return;
        value = clampTintValue(parseFloat(String(editText.text).replace(/^\s+|\s+$/g, "")));
        if (isNaN(value)) return;
        editText.text = String(Math.round(value));
    }

    function syncTintSliderFromInput(ui) {
        var tint;
        if (!ui || !ui.tintInput || !ui.tintSlider) return;
        clampTintInput(ui.tintInput);
        tint = getSelectedTint(ui);
        if (isNaN(tint)) return;
        ui.tintSlider.value = tint;
    }

    function syncTintInputFromSlider(ui) {
        if (!ui || !ui.tintInput || !ui.tintSlider) return;
        ui.tintInput.text = String(Math.round(ui.tintSlider.value));
    }

    function shouldEnableTintControlsBySwatchName(swatchName) {
        return !(isNoneSwatchName(swatchName) || isPaperSwatchName(swatchName));
    }

    function updateTintControlsEnabledState(dropdown, tintInput, tintSlider) {
        var swatchName = getSelectedSwatchNameFromDropdown(dropdown);
        var enabled = shouldEnableTintControlsBySwatchName(swatchName);

        if (tintInput) tintInput.enabled = enabled;
        if (tintSlider) tintSlider.enabled = enabled;
    }

    // --- Preview/Standard Mode toggle helpers ---
    function getPreviewToggleButtonLabel() {
        return isPreviewScreenMode() ? L('standardMode') : L('previewMode');
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

    function setCellEdgeTints(cell, top, bottom, left, right) {
        if (top != null) cell.topEdgeStrokeTint = top;
        if (bottom != null) cell.bottomEdgeStrokeTint = bottom;
        if (left != null) cell.leftEdgeStrokeTint = left;
        if (right != null) cell.rightEdgeStrokeTint = right;
    }

    function clearCellTopEdge(cell) {
        cell.topEdgeStrokeWeight = 0;
        try {
            cell.topEdgeStrokeColor = NothingEnum.NOTHING;
            cell.topEdgeStrokeTint = 100;
        } catch (e) { }
    }

    function clearCellBottomEdge(cell) {
        cell.bottomEdgeStrokeWeight = 0;
        try {
            cell.bottomEdgeStrokeColor = NothingEnum.NOTHING;
            cell.bottomEdgeStrokeTint = 100;
        } catch (e) { }
    }

    function clearCellLeftEdge(cell) {
        cell.leftEdgeStrokeWeight = 0;
        try {
            cell.leftEdgeStrokeColor = NothingEnum.NOTHING;
            cell.leftEdgeStrokeTint = 100;
        } catch (e) { }
    }

    function clearCellRightEdge(cell) {
        cell.rightEdgeStrokeWeight = 0;
        try {
            cell.rightEdgeStrokeColor = NothingEnum.NOTHING;
            cell.rightEdgeStrokeTint = 100;
        } catch (e) { }
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

    // =========================================
    // Selection helper: check if selection is full table
    // =========================================
    function isFullTableSelection(cells) {
        var table, totalCellCount;
        if (!cells || cells.length === 0) return false;

        table = getParentTableFromCell(cells[0]);
        if (!table) return false;

        totalCellCount = getAllTableCells(table).length;
        return totalCellCount > 0 && cells.length === totalCellCount;
    }

    function getParentTableFromCell(cell) {
        try {
            return cell.parent;
        } catch (e) {
            return null;
        }
    }

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

    function rangesOverlapHorizontally(a, b) {
        return !(a.endCol < b.startCol || b.endCol < a.startCol);
    }
})();