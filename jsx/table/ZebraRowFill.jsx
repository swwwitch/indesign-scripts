#target indesign

// 概要：
//
// InDesign Script: 選択したセルに対して、選択範囲内の行方向の並びを基準に交互の塗りを適用します。
// 選択範囲内で見た奇数行・偶数行それぞれにカラーと濃淡（0〜100%）を設定でき、
// スワップ、行のスキップ、列のスキップに対応しています。
//
// 【値の変更】
//   テキストフィールドは次のキー操作で値を増減できます：
//     ・↑↓キー：±1
//     ・Shift + ↑↓キー：±10
//
// 【スライダーの挙動】
//   濃淡スライダーはキー操作に応じて刻み幅が変わります：
//     ・通常：1% 刻み（細かく調整）
//     ・Shift：10% 刻み
//     ・Option（Alt）：5% 刻み
//
// 【現在の仕様】
//   ・選択セルのうち、行方向の並び順を基準に交互の塗りを適用します。
//   ・行のスキップは、選択範囲内で上から指定行をカラーリング対象外にし、ダイアログ起動時のカラーに戻します。
//   ・列のスキップは、選択範囲内で左から指定列をカラーリング対象外にし、ダイアログ起動時のカラーに戻します。
//   ・カラーリストでは Registration を常に非表示とし、Paper を「紙色」、None を「なし」と表示します。
//   ・「なし」「紙色」のときは濃淡 UI を無効化します。
//   ・「なし」「紙色」以外のスウォッチもすべてリストに表示します（グラデーション等も含む）。
//   ・成功時のメッセージは表示せず、エラー時のみアラート表示します。
//   ・プレビューは doScript + app.undo により、キャンセル時に変更を自動で巻き戻します。
//
// 【v1.2.0】
//   ・ローカライズを LABELS + L() 形式に整理
//   ・ダイアログタイトルを現仕様に合わせて更新
//   ・行／列スキップの文言とコメントを選択範囲基準に修正
//   ・スキップ対象のセルをダイアログ起動時のカラーに戻すように変更
//   ・行／列スキップ、濃淡入力のキー操作を整理
//   ・不要なダイアログ位置調整処理を削除
//
// オリジナルアイデア：KK sawaさん
// 更新日：2026-04-17

var SCRIPT_VERSION = "v1.2.0";

// ローカライズ設定
var IS_JAPANESE_UI = (function () {
    try {
        if (app.locale && app.locale === Locale.JAPANESE) return true;
    } catch (e) { }
    try {
        if ($.locale && $.locale.toString().indexOf("ja") === 0) return true;
    } catch (e2) { }
    return false;
})();

var lang = IS_JAPANESE_UI ? "ja" : "en";

var LABELS = {
    alertOpenDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
    alertSelectCells: { ja: "セルを選択してください。", en: "Please select table cells." },
    alertNoUsableColors: { ja: "使用可能なカラーがありません。", en: "No usable colors are available." },
    swatchPaper: { ja: "紙色", en: "Paper Color" },
    swatchNone: { ja: "なし", en: "None" },
    swatchBlack: { ja: "黒", en: "Black" },
    dialogTitle: { ja: "セルの塗りを交互に反復", en: "Repeat Cell Fills Every Other Row" },
    oddRows: { ja: "奇数行", en: "Odd Rows" },
    evenRows: { ja: "偶数行", en: "Even Rows" },
    colorLabel: { ja: "カラー：", en: "Color:" },
    tintLabel: { ja: "濃淡：", en: "Tint:" },
    swap: { ja: "交換", en: "Swap" },
    skipRows: { ja: "行のスキップ", en: "Skip top rows" },
    rowUnit: { ja: "行", en: "rows" },
    skipColumns: { ja: "列のスキップ", en: "Skip left columns" },
    columnUnit: { ja: "列", en: "columns" },
    preview: { ja: "プレビュー", en: "Preview" },
    standardMode: { ja: "標準モード", en: "Standard" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" },
    undoFillPreview: { ja: "塗りプレビュー", en: "Fill Preview" },
    undoApplyFill: { ja: "塗りの設定", en: "Apply Fill" }
};

function L(key) {
    return LABELS[key][lang];
}

function main() {
    // -------------------------------------------------------------
    // 前提チェック
    // -------------------------------------------------------------
    if (app.documents.length === 0) {
        alert(L("alertOpenDocument"));
        return;
    }

    var doc = app.activeDocument;

    // 選択セルを保持し、プレビュー中は選択ハイライトを一旦消して見やすくする
    var baseTargetCells = [];
    var selCells;
    try {
        selCells = app.selection[0].cells;
    } catch (e) {
        alert(L("alertSelectCells"));
        return;
    }
    if (!selCells || selCells.length === 0) {
        alert(L("alertSelectCells"));
        return;
    }
    for (var bi = 0; bi < selCells.length; bi++) {
        baseTargetCells.push(selCells[bi]);
    }
    // 元の塗り状態を保持（スキップ時に復元するため）
    var originalCellStyles = [];
    for (var oi = 0; oi < baseTargetCells.length; oi++) {
        var oc = baseTargetCells[oi];
        originalCellStyles.push({
            color: oc.fillColor,
            tint: oc.fillTint
        });
    }
    // プレビュー時にセルの青いハイライトが邪魔になるため、選択は一旦解除
    app.selection = null;

    // -------------------------------------------------------------
    // スウォッチ取得（Registration を除外）
    // -------------------------------------------------------------
    var REG_NAMES = ["Registration", "[Registration]", "トンボ用"];
    var PAPER_NAMES = ["Paper", "[Paper]", "紙色", "[紙色]"];
    var NONE_NAMES = ["None", "[None]", "なし", "[なし]"];

    function isInList(name, list) {
        for (var i = 0; i < list.length; i++) {
            if (name === list[i]) return true;
        }
        return false;
    }

    function isRegistrationName(name) {
        if (!name) return false;
        if (isInList(name, REG_NAMES)) return true;
        return name.toLowerCase() === "registration";
    }

    function isPaperName(name) {
        return isInList(name, PAPER_NAMES);
    }

    function isNoneName(name) {
        return isInList(name, NONE_NAMES);
    }

    var swatches = [];
    for (var i = 0; i < doc.swatches.length; i++) {
        var sw = doc.swatches[i];
        var name = sw.name;

        if (isRegistrationName(name)) continue;

        var label = name;
        if (isPaperName(name)) {
            label = L("swatchPaper");
        } else if (isNoneName(name)) {
            label = L("swatchNone");
        }
        if (name === "Black") label = L("swatchBlack");

        swatches.push({
            name: name,
            label: label
        });
    }

    if (swatches.length === 0) {
        alert(L("alertNoUsableColors"));
        return;
    }

    // -------------------------------------------------------------
    // プレビュー管理ステート
    // -------------------------------------------------------------
    var state = { previewed: false };

    // -------------------------------------------------------------
    // スウォッチプレビュー用ヘルパー
    // -------------------------------------------------------------
    function normalizeSwatchName(n) {
        if (n == null) return "";
        n = String(n);
        n = n.replace(/^\[|\]$/g, "");
        n = n.replace(/^\s+|\s+$/g, "").toLowerCase();
        return n;
    }

    function isBlackPreviewName(n) {
        var x = normalizeSwatchName(n);
        return x === "black" || x === "ブラック" || x === "黒" || x === "registration";
    }

    function isWhitePreviewName(n) {
        var x = normalizeSwatchName(n);
        return x === "paper" || x === "紙色" || x === "none" || x === "なし";
    }

    function convertSwatchToPreviewRGB(swatch) {
        if (!swatch) return [0.5, 0.5, 0.5];
        var sname = swatch.name != null ? String(swatch.name) : "";
        if (isWhitePreviewName(sname)) return [1, 1, 1];
        if (isBlackPreviewName(sname)) return [0, 0, 0];
        try {
            if (swatch.hasOwnProperty("colorValue")) {
                var vals = swatch.colorValue;
                if (swatch.space === ColorSpace.RGB) {
                    return [vals[0] / 255, vals[1] / 255, vals[2] / 255];
                }
                if (swatch.space === ColorSpace.CMYK) {
                    var c = vals[0] / 100,
                        m = vals[1] / 100,
                        y = vals[2] / 100,
                        k = vals[3] / 100;
                    return [(1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k)];
                }
            }
        } catch (e) { }
        return [0.5, 0.5, 0.5];
    }

    function createSwatchPreviewBox(parent) {
        var box = parent.add("group");
        box.preferredSize = [18, 18];
        box.minimumSize = [18, 18];
        box.maximumSize = [18, 18];
        return box;
    }

    function updateSwatchPreview(previewBox, swatchName) {
        if (!previewBox) return;
        var swatch = null;
        try {
            swatch = doc.swatches.itemByName(String(swatchName));
            if (!swatch || !swatch.isValid) return;
        } catch (e) {
            return;
        }
        var rgb = convertSwatchToPreviewRGB(swatch);
        try {
            previewBox.graphics.backgroundColor = previewBox.graphics.newBrush(
                previewBox.graphics.BrushType.SOLID_COLOR,
                [rgb[0], rgb[1], rgb[2], 1]
            );
            if (dialog) dialog.update();
        } catch (e2) { }
    }

    // -------------------------------------------------------------
    // プレビュー表示モード切替ヘルパー
    // -------------------------------------------------------------
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

    function getPreviewToggleButtonLabel() {
        return isPreviewScreenMode()
            ? L("preview")
            : L("standardMode");
    }

    function updatePreviewToggleButtonLabel(button) {
        if (!button) return;
        button.text = getPreviewToggleButtonLabel();
    }

    // -------------------------------------------------------------
    // デフォルトのカラー／濃淡（選択セル内で使われている設定をもとに決定）
    // -------------------------------------------------------------
    var defaultOddColorName = swatches[0].name;
    var defaultEvenColorName = swatches[0].name;
    var defaultOddTint = 100;
    var defaultEvenTint = 100;

    try {
        var firstCell = baseTargetCells[0];
        var fillColor = firstCell.fillColor;
        var fillTint = firstCell.fillTint;

        if (fillColor && fillColor.name) {
            defaultOddColorName = fillColor.name;
            defaultEvenColorName = fillColor.name;
        }
        if (!isNaN(fillTint)) {
            defaultOddTint = fillTint;
            defaultEvenTint = fillTint;
        }

        var comboCounts = {};

        for (var ci = 0; ci < baseTargetCells.length; ci++) {
            var c = baseTargetCells[ci];
            var cFillColor = c.fillColor;
            var cFillTint = c.fillTint;

            var cColorName = (cFillColor && cFillColor.name) ? cFillColor.name : defaultOddColorName;
            var cTint = (!isNaN(cFillTint)) ? cFillTint : defaultOddTint;

            var key = cColorName + "||" + cTint;
            if (!comboCounts[key]) {
                comboCounts[key] = {
                    count: 0,
                    colorName: cColorName,
                    tint: cTint
                };
            }
            comboCounts[key].count++;
        }

        var combos = [];
        for (var kCombo in comboCounts) {
            combos.push(comboCounts[kCombo]);
        }
        combos.sort(function (a, b) {
            return b.count - a.count;
        });

        if (combos.length >= 1) {
            defaultOddColorName = combos[0].colorName;
            defaultOddTint = combos[0].tint;
        }
        if (combos.length >= 2) {
            defaultEvenColorName = combos[1].colorName;
            defaultEvenTint = combos[1].tint;
        } else {
            defaultEvenColorName = defaultOddColorName;
            defaultEvenTint = defaultOddTint;
        }
    } catch (e) { }

    // -------------------------------------------------------------
    // ダイアログボックス作成
    // -------------------------------------------------------------
    var dialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    var dialogOpacity = 0.97;


    dialog.opacity = dialogOpacity;
    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    // メインパネル（左右2カラム）
    var mainGroup = dialog.add("group");
    mainGroup.orientation = "row";
    mainGroup.alignChildren = "top";
    mainGroup.spacing = 20;

    // -------------------------------------------------------------
    // 左カラム：奇数行パネル
    // -------------------------------------------------------------
    var oddPanel = mainGroup.add("panel", undefined, L("oddRows"));
    oddPanel.orientation = "column";
    oddPanel.alignChildren = "left";
    oddPanel.spacing = 10;
    oddPanel.margins = [15, 20, 15, 10];

    var oddColorPanel = oddPanel.add("group");
    oddColorPanel.orientation = "column";
    oddColorPanel.alignChildren = "left";

    var oddColorLabelRow = oddColorPanel.add("group");
    oddColorLabelRow.orientation = "row";
    oddColorLabelRow.alignChildren = ["left", "center"];
    oddColorLabelRow.add("statictext", undefined, L("colorLabel"));
    var oddPreviewBox = createSwatchPreviewBox(oddColorLabelRow);

    var oddSwatchLabels = [];
    for (var j = 0; j < swatches.length; j++) {
        oddSwatchLabels.push(swatches[j].label);
    }
    var oddColorDropdown = oddColorPanel.add("dropdownlist", undefined, oddSwatchLabels);
    oddColorDropdown.preferredSize.width = 140;

    var oddDefaultIdx = 0;
    for (var j2 = 0; j2 < swatches.length; j2++) {
        if (swatches[j2].name === defaultOddColorName) {
            oddDefaultIdx = j2;
            break;
        }
    }
    oddColorDropdown.selection = oddDefaultIdx;

    oddColorDropdown.onChange = function () {
        updateTintUI(true);
        updateSwatchPreview(oddPreviewBox, getOddSelectedColorName());
        updatePreview();
    };

    // 濃淡設定（奇数）
    var oddTintPanel = oddPanel.add("group");
    oddTintPanel.orientation = "column";
    oddTintPanel.alignChildren = "left";
    oddTintPanel.margins = [0, 20, 0, 10];

    var oddTintLabelGroup = oddTintPanel.add("group");
    oddTintLabelGroup.orientation = "row";
    oddTintLabelGroup.add("statictext", undefined, L("tintLabel"));
    var oddTintText = oddTintLabelGroup.add("edittext", undefined, String(defaultOddTint));
    oddTintText.preferredSize.width = 50;

    var oddSlider = oddTintPanel.add("slider", undefined, defaultOddTint, 0, 100);
    oddSlider.preferredSize.width = 150;

    // -------------------------------------------------------------
    // 右カラム：偶数行パネル
    // -------------------------------------------------------------
    var evenPanel = mainGroup.add("panel", undefined, L("evenRows"));
    evenPanel.orientation = "column";
    evenPanel.alignChildren = "left";
    evenPanel.spacing = 10;
    evenPanel.margins = [15, 20, 15, 10];

    var evenColorPanel = evenPanel.add("group");
    evenColorPanel.orientation = "column";
    evenColorPanel.alignChildren = "left";

    var evenColorLabelRow = evenColorPanel.add("group");
    evenColorLabelRow.orientation = "row";
    evenColorLabelRow.alignChildren = ["left", "center"];
    evenColorLabelRow.add("statictext", undefined, L("colorLabel"));
    var evenPreviewBox = createSwatchPreviewBox(evenColorLabelRow);

    var evenSwatchLabels = [];
    for (var k = 0; k < swatches.length; k++) {
        evenSwatchLabels.push(swatches[k].label);
    }
    var evenColorDropdown = evenColorPanel.add("dropdownlist", undefined, evenSwatchLabels);
    evenColorDropdown.preferredSize.width = 140;

    var evenDefaultIdx = 0;
    for (var k2 = 0; k2 < swatches.length; k2++) {
        if (swatches[k2].name === defaultEvenColorName) {
            evenDefaultIdx = k2;
            break;
        }
    }
    evenColorDropdown.selection = evenDefaultIdx;

    evenColorDropdown.onChange = function () {
        updateTintUI(false);
        updateSwatchPreview(evenPreviewBox, getEvenSelectedColorName());
        updatePreview();
    };

    // 濃淡設定（偶数）
    var evenTintPanel = evenPanel.add("group");
    evenTintPanel.orientation = "column";
    evenTintPanel.alignChildren = "left";
    evenTintPanel.margins = [0, 20, 0, 10];

    var evenTintLabelGroup = evenTintPanel.add("group");
    evenTintLabelGroup.orientation = "row";
    evenTintLabelGroup.add("statictext", undefined, L("tintLabel"));
    var evenTintText = evenTintLabelGroup.add("edittext", undefined, String(defaultEvenTint));
    evenTintText.preferredSize.width = 50;

    var evenSlider = evenTintPanel.add("slider", undefined, defaultEvenTint, 0, 100);
    evenSlider.preferredSize.width = 150;

    // -------------------------------------------------------------
    // オプションエリア（交換／スキップ設定）
    // -------------------------------------------------------------
    var optionGroup = dialog.add("group");
    optionGroup.orientation = "column";
    optionGroup.alignChildren = "left";
    optionGroup.alignment = ["left", "top"];

    var swapCheckbox = optionGroup.add("checkbox", undefined, L("swap"));
    swapCheckbox.value = false;
    swapCheckbox.onClick = function () {
        updatePreview();
    };

    // 行のスキップ：選択範囲内で、上から指定行のみカラーリングしない
    var rowSkipGroup = optionGroup.add("group");
    rowSkipGroup.orientation = "row";
    var rowSkipCheckbox = rowSkipGroup.add("checkbox", undefined, L("skipRows"));
    var rowSkipText = rowSkipGroup.add("edittext", undefined, "1");
    rowSkipText.preferredSize.width = 30;
    rowSkipGroup.add("statictext", undefined, L("rowUnit"));

    // 列のスキップ：選択範囲内で、左から指定列のみカラーリングしない
    var colSkipGroup = optionGroup.add("group");
    colSkipGroup.orientation = "row";
    var colSkipCheckbox = colSkipGroup.add("checkbox", undefined, L("skipColumns"));
    var colSkipText = colSkipGroup.add("edittext", undefined, "1");
    colSkipText.preferredSize.width = 30;
    colSkipGroup.add("statictext", undefined, L("columnUnit"));

    function adjustSkipValue(v) {
        if (isNaN(v)) v = 0;
        if (v < 0) v = 0;
        return Math.round(v);
    }

    function formatArrowValue(value) {
        return String(Math.round(value));
    }

    function bindArrowKeyValueControl(editText, allowNegative, onAfterChange, options) {
        options = options || {};
        var minValue = (options.min != null) ? options.min : (allowNegative ? null : 0);
        var maxValue = (options.max != null) ? options.max : null;

        editText.addEventListener("keydown", function (event) {
            var key = event.keyName;
            if (key !== "Up" && key !== "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = keyboard.shiftKey ? 10 : 1;

            if (keyboard.shiftKey) {
                if (key === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                } else {
                    value = Math.floor((value - 1) / delta) * delta;
                }
            } else {
                if (key === "Up") {
                    value += delta;
                } else {
                    value -= delta;
                }
            }

            value = Math.round(value);

            if (!allowNegative && value < 0) value = 0;
            if (minValue != null && value < minValue) value = minValue;
            if (maxValue != null && value > maxValue) value = maxValue;

            editText.text = formatArrowValue(value);

            if (typeof onAfterChange === "function") {
                onAfterChange(value, editText);
            }

            try { event.preventDefault(); } catch (e) { }
        });
    }

    function updateSkipUI() {
        rowSkipText.enabled = rowSkipCheckbox.value;
        colSkipText.enabled = colSkipCheckbox.value;
    }
    updateSkipUI();

    rowSkipCheckbox.onClick = function () {
        updateSkipUI();
        updatePreview();
    };
    colSkipCheckbox.onClick = function () {
        updateSkipUI();
        updatePreview();
    };
    rowSkipText.onChange = function () {
        var v = adjustSkipValue(parseInt(this.text, 10));
        this.text = String(v);
        updatePreview();
    };
    colSkipText.onChange = function () {
        var v = adjustSkipValue(parseInt(this.text, 10));
        this.text = String(v);
        updatePreview();
    };
    bindArrowKeyValueControl(rowSkipText, false, function () {
        var v = adjustSkipValue(parseInt(rowSkipText.text, 10));
        rowSkipText.text = String(v);
        updatePreview();
    }, { min: 0 });

    bindArrowKeyValueControl(colSkipText, false, function () {
        var v = adjustSkipValue(parseInt(colSkipText.text, 10));
        colSkipText.text = String(v);
        updatePreview();
    }, { min: 0 });

    // -------------------------------------------------------------
    // ボタングループ（左：プレビューモード切替／右：キャンセル・OK）
    // -------------------------------------------------------------
    var bottomGroup = dialog.add("group");
    bottomGroup.orientation = "row";
    bottomGroup.alignChildren = "fill";
    bottomGroup.alignment = ["fill", "top"];
    bottomGroup.margins = [0, 8, 0, 0];

    var btnPreviewToggle = bottomGroup.add("button", undefined, getPreviewToggleButtonLabel());
    btnPreviewToggle.alignment = ["left", "center"];
    btnPreviewToggle.onClick = function () {
        togglePreviewScreenMode();
        updatePreviewToggleButtonLabel(btnPreviewToggle);
    };

    var spacer = bottomGroup.add("group");
    spacer.alignment = ["fill", "top"];

    var buttonGroup = bottomGroup.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = ["right", "top"];
    buttonGroup.alignChildren = "right";

    var cancelButton = buttonGroup.add("button", undefined, L("cancel"), {
        name: "cancel"
    });
    var okButton = buttonGroup.add("button", undefined, L("ok"), {
        name: "ok"
    });

    // -------------------------------------------------------------
    // 濃淡ディム制御用ヘルパー
    // -------------------------------------------------------------
    function getOddSelectedColorName() {
        if (oddColorDropdown.selection) {
            return swatches[oddColorDropdown.selection.index].name;
        }
        return swatches[0].name;
    }

    function getEvenSelectedColorName() {
        if (evenColorDropdown.selection) {
            return swatches[evenColorDropdown.selection.index].name;
        }
        return swatches[0].name;
    }

    function updateTintUI(isOdd) {
        var name = isOdd ? getOddSelectedColorName() : getEvenSelectedColorName();
        var disabled = isNoneName(name) || isPaperName(name);

        if (isOdd) {
            oddSlider.enabled = !disabled;
            oddTintText.enabled = !disabled;
        } else {
            evenSlider.enabled = !disabled;
            evenTintText.enabled = !disabled;
        }
    }

    updateTintUI(true);
    updateTintUI(false);
    updateSwatchPreview(oddPreviewBox, getOddSelectedColorName());
    updateSwatchPreview(evenPreviewBox, getEvenSelectedColorName());

    // -------------------------------------------------------------
    // プレビュー制御（doScript + app.undo）
    // -------------------------------------------------------------
    function clearPreview() {
        if (!state.previewed) return;
        try { app.undo(); } catch (e) { }
        state.previewed = false;
        try { doc.recompose(); } catch (e2) { }
    }

    function updatePreview() {
        clearPreview();
        var oddColorName = getOddSelectedColorName();
        var evenColorName = getEvenSelectedColorName();
        var oddTint = adjustTintTextValue(parseFloat(oddTintText.text));
        var evenTint = adjustTintTextValue(parseFloat(evenTintText.text));
        var swap = swapCheckbox.value;
        var rowSkip = rowSkipCheckbox.value ? Math.max(0, parseInt(rowSkipText.text, 10) || 0) : 0;
        var colSkip = colSkipCheckbox.value ? Math.max(0, parseInt(colSkipText.text, 10) || 0) : 0;
        try {
            app.doScript(function () {
                applyZebra(oddColorName, evenColorName, oddTint, evenTint, swap, rowSkip, colSkip);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, L("undoFillPreview"));
            state.previewed = true;
            try { doc.recompose(); } catch (e) { }
        } catch (e2) {
            state.previewed = false;
        }
    }

    // -------------------------------------------------------------
    // 濃淡入力ヘルパー：スライダーは 1 / 10 / 5 % 刻み、テキストは ↑↓ / Shift+↑↓ に対応
    // -------------------------------------------------------------
    function getTintSliderStep() {
        var ks = ScriptUI.environment.keyboardState;
        if (ks.shiftKey) return 10;
        if (ks.altKey) return 5;
        return 1;
    }

    function adjustTintSliderValue(v) {
        if (isNaN(v)) v = 100;
        if (v < 0) v = 0;
        if (v > 100) v = 100;
        var step = getTintSliderStep();
        v = Math.round(v / step) * step;
        if (v < 0) v = 0;
        if (v > 100) v = 100;
        return v;
    }

    function adjustTintTextValue(v) {
        if (isNaN(v)) v = 100;
        if (v < 0) v = 0;
        if (v > 100) v = 100;
        return Math.round(v);
    }

    oddSlider.onChanging = function () {
        var v = adjustTintSliderValue(this.value);
        this.value = v;
        oddTintText.text = formatArrowValue(v);
        updatePreview();
    };

    oddTintText.onChange = function () {
        var v = adjustTintTextValue(parseFloat(this.text));
        this.text = formatArrowValue(v);
        oddSlider.value = v;
        updatePreview();
    };

    bindArrowKeyValueControl(oddTintText, false, function () {
        var v = adjustTintTextValue(parseFloat(oddTintText.text));
        oddTintText.text = formatArrowValue(v);
        oddSlider.value = v;
        updatePreview();
    }, { min: 0, max: 100 });

    evenSlider.onChanging = function () {
        var v = adjustTintSliderValue(this.value);
        this.value = v;
        evenTintText.text = formatArrowValue(v);
        updatePreview();
    };

    evenTintText.onChange = function () {
        var v = adjustTintTextValue(parseFloat(this.text));
        this.text = formatArrowValue(v);
        evenSlider.value = v;
        updatePreview();
    };

    bindArrowKeyValueControl(evenTintText, false, function () {
        var v = adjustTintTextValue(parseFloat(evenTintText.text));
        evenTintText.text = formatArrowValue(v);
        evenSlider.value = v;
        updatePreview();
    }, { min: 0, max: 100 });

    // -------------------------------------------------------------
    // 実行処理
    // -------------------------------------------------------------
    function applyZebra(oddColorName, evenColorName, oddTint, evenTint, swap, rowSkip, colSkip) {
        if (swap) {
            var tmpName = oddColorName;
            oddColorName = evenColorName;
            evenColorName = tmpName;
            var tmpTint = oddTint;
            oddTint = evenTint;
            evenTint = tmpTint;
        }

        rowSkip = rowSkip || 0;
        colSkip = colSkip || 0;

        var oddSwatch = doc.swatches.item(oddColorName);
        var evenSwatch = doc.swatches.item(evenColorName);

        if (!baseTargetCells || baseTargetCells.length === 0) return;

        var targetCells = [];
        for (var i = 0; i < baseTargetCells.length; i++) {
            targetCells.push(baseTargetCells[i]);
        }

        // 選択範囲内の行・列インデックスを昇順で収集
        var uniqueRowIndices = [];
        var uniqueColIndices = [];
        var seenRows = {};
        var seenCols = {};
        for (var uri = 0; uri < targetCells.length; uri++) {
            var rIdx = targetCells[uri].parentRow.index;
            var cIdx = targetCells[uri].parentColumn.index;
            if (!seenRows[rIdx]) { seenRows[rIdx] = true; uniqueRowIndices.push(rIdx); }
            if (!seenCols[cIdx]) { seenCols[cIdx] = true; uniqueColIndices.push(cIdx); }
        }
        uniqueRowIndices.sort(function (a, b) { return a - b; });
        uniqueColIndices.sort(function (a, b) { return a - b; });

        // 選択範囲内で、上から rowSkip 行・左から colSkip 列をカラーリング対象外にする
        var rowsToSkip = {};
        for (var rs = 0; rs < rowSkip && rs < uniqueRowIndices.length; rs++) {
            rowsToSkip[String(uniqueRowIndices[rs])] = true;
        }
        var colsToSkip = {};
        for (var cs = 0; cs < colSkip && cs < uniqueColIndices.length; cs++) {
            colsToSkip[String(uniqueColIndices[cs])] = true;
        }

        // スキップされていない行に、選択範囲内で上から 0, 1, 2, ... の順番を振る
        var rowIndexMap = {};
        var orderCounter = 0;
        for (var uri2 = 0; uri2 < uniqueRowIndices.length; uri2++) {
            var ridx = uniqueRowIndices[uri2];
            if (rowsToSkip[String(ridx)]) continue;
            rowIndexMap[ridx] = orderCounter++;
        }

        for (var n = 0; n < targetCells.length; n++) {
            var cell = targetCells[n];
            var tableRowIndex = cell.parentRow.index;
            var tableColIndex = cell.parentColumn.index;

            // スキップ対象は元のカラーに戻す
            if (rowsToSkip[String(tableRowIndex)] || colsToSkip[String(tableColIndex)]) {
                try {
                    var original = originalCellStyles[n];
                    if (original) {
                        cell.fillColor = original.color;
                        if (!isNaN(original.tint)) {
                            cell.fillTint = original.tint;
                        }
                    }
                } catch (e) { }
                continue;
            }

            var localRowOrder = rowIndexMap[tableRowIndex];
            if (localRowOrder === undefined) continue;

            if (localRowOrder % 2 === 0) {
                cell.fillColor = oddSwatch;
                if (!isNoneName(oddColorName) && !isPaperName(oddColorName)) {
                    try { cell.fillTint = oddTint; } catch (e) { }
                }
            } else {
                cell.fillColor = evenSwatch;
                if (!isNoneName(evenColorName) && !isPaperName(evenColorName)) {
                    try { cell.fillTint = evenTint; } catch (e) { }
                }
            }
        }
    }

    // 初期プレビュー
    updatePreview();

    var result = dialog.show();

    if (result == 1) {
        clearPreview();
        var finalOddColorName = getOddSelectedColorName();
        var finalEvenColorName = getEvenSelectedColorName();
        var finalOddTint = adjustTintTextValue(parseFloat(oddTintText.text));
        var finalEvenTint = adjustTintTextValue(parseFloat(evenTintText.text));
        var finalSwap = swapCheckbox.value;
        var finalRowSkip = rowSkipCheckbox.value ? Math.max(0, parseInt(rowSkipText.text, 10) || 0) : 0;
        var finalColSkip = colSkipCheckbox.value ? Math.max(0, parseInt(colSkipText.text, 10) || 0) : 0;
        app.doScript(function () {
            applyZebra(finalOddColorName, finalEvenColorName, finalOddTint, finalEvenTint,
                finalSwap, finalRowSkip, finalColSkip);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, L("undoApplyFill"));
    } else {
        clearPreview();
    }
}

main();