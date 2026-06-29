#target indesign

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

/*

### 概要

- 段落スタイルの「同一スタイル間の段落間隔」をスタイル定義ごと上書きします。
- グループ内の段落スタイルも再帰的に収集し、ドロップダウンから選択できます。
- 間隔は「無視」「0」「数値指定」から選択でき、数値は環境設定の表示単位に連動します。
- 数値入力欄は ↑↓ キーで増減できます（Shift＝10 の倍数にスナップ、Option＝0.1 刻み）。
- ダイアログ起動時、テキスト選択中なら適用中の段落スタイルを初期選択にし、現在値も反映します。

### Overview

- Overrides a paragraph style's "Spacing (Same Style)" at the style-definition level.
- Collects paragraph styles recursively, including those inside style groups.
- Spacing can be set to "Ignore", "0", or a numeric value linked to the document display units.
- The value field supports arrow-key stepping (Shift = snap to multiples of 10, Option = 0.1 steps).
- On launch, preselects the paragraph style of the current text selection and reflects its current value.

*/

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 既定で選択する段落スタイル名 / Default paragraph style to preselect */
var DEFAULT_STYLE_NAME = "p";

/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;
var COLUMN_SPACING = 12; /* 2カラムの間隔 / Gap between the two columns */

// =========================================
// ローカライズ / Localization
// =========================================

/* 現在の言語を判定 / Detect current language */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

/* ラベル定義 / Labels */
var LABELS = {
    dialog: {
        title: { ja: "段落スタイルの段落間隔を上書き", en: "Override Paragraph Style Spacing" }
    },
    panel: {
        style:   { ja: "段落スタイル", en: "Paragraph Style" },
        spacing: { ja: "同一設定段落の間隔", en: "Spacing (Same Style)" }
    },
    radio: {
        ignore:   { ja: "無視", en: "Ignore" },
        zero:     { ja: "0", en: "0" },
        useValue: { ja: "数値指定", en: "Use value" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    alert: {
        noDoc:   { ja: "ドキュメントを開いてください。", en: "Please open a document." },
        noStyle: { ja: "段落スタイルがありません。", en: "No paragraph styles found." }
    }
};

/* ドット区切りキーからラベルを取得 / Look up a label by dot-separated key */
function L(key) {
    var parts = key.split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) node = node[parts[i]];
    return node[currentLanguage];
}

// =========================================
// 単位 / Units
// =========================================

/* MeasurementUnits を単位表記へ / Map MeasurementUnits to a label */
function unitLabel(unit) {
    if (unit === MeasurementUnits.POINTS) return "pt";
    if (unit === MeasurementUnits.PICAS) return "pica";
    if (unit === MeasurementUnits.MILLIMETERS) return "mm";
    if (unit === MeasurementUnits.CENTIMETERS) return "cm";
    if (unit === MeasurementUnits.INCHES || unit === MeasurementUnits.INCHES_DECIMAL) return "inch";
    if (unit === MeasurementUnits.CICEROS) return "cicero";
    if (unit === MeasurementUnits.AGATES) return "agate";
    if (unit === MeasurementUnits.PIXELS) return "px";
    if (unit === MeasurementUnits.Q) return "Q";
    if (unit === MeasurementUnits.HA) return "H";
    if (unit === MeasurementUnits.AMERICAN_POINTS) return "pt(US)";
    if (unit === MeasurementUnits.BAI) return "bai";
    if (unit === MeasurementUnits.MILS) return "mils";
    if (unit === MeasurementUnits.U) return "u";
    return String(unit);
}

// =========================================
// メイン処理 / Main
// =========================================

main();

/* ドキュメント確認・単位整合・ダイアログ表示・スタイル上書きを統括 / Orchestrate the whole flow */
function main() {
    if (app.documents.length === 0) {
        alert(L("alert.noDoc"));
        return;
    }

    /* 環境設定の単位を参照（段落間隔は垂直方向）/ Reference preference units (vertical) */
    var unit = app.activeDocument.viewPreferences.verticalMeasurementUnits;
    var unitText = unitLabel(unit);

    /* スクリプトの単位を表示単位に合わせ、読み書きを一致させる / Match script units to display */
    var savedUnit = app.scriptPreferences.measurementUnit;
    app.scriptPreferences.measurementUnit = unit;

    try {
        var styleEntries = [];
        collectParagraphStyles(app.activeDocument, "", styleEntries);
        if (styleEntries.length === 0) {
            alert(L("alert.noStyle"));
            return;
        }

        /* 選択中の段落スタイルを初期選択に使う / Preselect the current selection's style */
        var selectedStyle = getSelectedParagraphStyle();

        var settings = showDialog(styleEntries, unitText, selectedStyle);
        if (settings === null || settings.style === null) return;

        /* 選択に依存せず、スタイル定義そのものを上書き / Override the style definition itself */
        if (settings.ignore) {
            settings.style.sameParaStyleSpacing = Spacing.SETIGNORE;
        } else {
            settings.style.sameParaStyleSpacing = settings.value;
        }
    } finally {
        app.scriptPreferences.measurementUnit = savedUnit;
    }
}

/* 段落スタイルを再帰収集（グループはパス付き）/ Collect paragraph styles recursively */
/* 名前は everyItem() で一括取得し、実体は同コレクションの同インデックス参照で必ず一致させる / Names batched, objects by same index */
function collectParagraphStyles(container, prefix, list) {
    var styles = container.paragraphStyles;
    var styleNames = [].concat(styles.everyItem().name);
    for (var i = 0; i < styleNames.length; i++) {
        list.push({ name: prefix + styleNames[i], style: styles[i] });
    }
    var groups = container.paragraphStyleGroups;
    var groupNames = [].concat(groups.everyItem().name);
    for (var g = 0; g < groupNames.length; g++) {
        collectParagraphStyles(groups[g], prefix + groupNames[g] + "/", list);
    }
}

// =========================================
// UI / Dialog
// =========================================

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* ↑↓キーで値を増減（Shift=10刻み・Option=0.1刻み）/ Increment value with arrow keys */
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function (event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
            /* Shift 押下時は 10 の倍数にスナップ / Snap to multiples of 10 */
            if (event.keyName === "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName === "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
            delta = 0.1;
            /* Option 押下時は 0.1 単位で増減 / Step by 0.1 */
            if (event.keyName === "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName === "Down") {
                value -= delta;
                event.preventDefault();
            }
        } else {
            delta = 1;
            if (event.keyName === "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName === "Down") {
                value -= delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        }

        if (keyboard.altKey) {
            /* 小数第1位までに丸め / Round to 1 decimal */
            value = Math.round(value * 10) / 10;
        } else {
            /* 整数に丸め / Round to integer */
            value = Math.round(value);
        }

        editText.text = value;
    });
}

/* スタイル選択と段落間隔を入力するダイアログ / Dialog to pick a style and its spacing */
function showDialog(styleEntries, unitText, selectedStyle) {
    var dialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
    dialog.alignChildren = "fill";

    /* --- 段落スタイル / Paragraph style --- */
    var stylePanel = dialog.add("panel", undefined, L("panel.style"));
    setupPanel(stylePanel);

    var names = [];
    for (var i = 0; i < styleEntries.length; i++) names.push(styleEntries[i].name);
    var styleDropdown = stylePanel.add("dropdownlist", undefined, names);

    /* 既定選択：選択中スタイル → DEFAULT_STYLE_NAME → 先頭 / Default selection */
    var defaultIndex = indexOfStyle(styleEntries, selectedStyle);
    if (defaultIndex < 0) defaultIndex = indexOfName(styleEntries, DEFAULT_STYLE_NAME);
    if (defaultIndex < 0) defaultIndex = 0;
    styleDropdown.selection = defaultIndex;

    /* --- 段落間隔 / Spacing --- */
    var spacingPanel = dialog.add("panel", undefined, L("panel.spacing"));
    setupPanel(spacingPanel);

    var ignoreRadio = spacingPanel.add("radiobutton", undefined, L("radio.ignore"));
    var zeroRadio = spacingPanel.add("radiobutton", undefined, L("radio.zero"));

    var valueGroup = spacingPanel.add("group");
    var valueRadio = valueGroup.add("radiobutton", undefined, L("radio.useValue"));
    var valueInput = valueGroup.add("edittext", undefined, "0");
    valueInput.characters = 6;
    changeValueByArrowKey(valueInput);
    valueGroup.add("statictext", undefined, unitText);

    /* 別コンテナのため手動で排他化（"ignore" / "zero" / "value"）/ Manual exclusivity */
    function setSpacingMode(mode) {
        ignoreRadio.value = (mode === "ignore");
        zeroRadio.value = (mode === "zero");
        valueRadio.value = (mode === "value");
        valueInput.enabled = (mode === "value");
    }
    ignoreRadio.onClick = function () { setSpacingMode("ignore"); };
    zeroRadio.onClick = function () { setSpacingMode("zero"); };
    valueRadio.onClick = function () { setSpacingMode("value"); };

    /* 選択スタイルの現在値を UI に反映 / Reflect the selected style's value */
    function refreshSpacingFromStyle(style) {
        var current = style.sameParaStyleSpacing;
        if (isIgnoreValue(current)) {
            setSpacingMode("ignore");
        } else if (current === 0) {
            setSpacingMode("zero");
            valueInput.text = "0";
        } else {
            setSpacingMode("value");
            valueInput.text = String(current);
        }
    }
    refreshSpacingFromStyle(styleEntries[defaultIndex].style);

    styleDropdown.onChange = function () {
        if (styleDropdown.selection !== null) {
            refreshSpacingFromStyle(styleEntries[styleDropdown.selection.index].style);
        }
    };

    /* --- ボタン / Buttons（Mac 規約：Cancel → OK）--- */
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    var cancelButton = buttonGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
    var okButton = buttonGroup.add("button", undefined, "OK", { name: "ok" });

    if (dialog.show() !== 1) return null;

    var selectedEntry = (styleDropdown.selection !== null)
        ? styleEntries[styleDropdown.selection.index]
        : null;

    return {
        style: selectedEntry ? selectedEntry.style : null,
        ignore: ignoreRadio.value,
        value: zeroRadio.value ? 0 : (parseFloat(valueInput.text) || 0)
    };
}

// =========================================
// ヘルパー / Helpers
// =========================================

/* 選択中の段落から適用中の段落スタイルを取得 / Applied paragraph style of the current selection */
function getSelectedParagraphStyle() {
    var sel = app.selection;
    if (!sel || sel.length === 0) return null;
    try {
        var paragraphs = sel[0].paragraphs;
        if (paragraphs && paragraphs.length > 0) {
            return paragraphs[0].appliedParagraphStyle;
        }
    } catch (e) {}
    return null;
}

/* スタイル名からインデックスを取得 / Find index by name */
function indexOfName(entries, name) {
    if (!name) return -1;
    for (var i = 0; i < entries.length; i++) {
        if (entries[i].name === name) return i;
    }
    return -1;
}

/* スタイル実体（id 一致）からインデックスを取得 / Find index by style id */
function indexOfStyle(entries, style) {
    if (!style) return -1;
    for (var i = 0; i < entries.length; i++) {
        if (entries[i].style.id === style.id) return i;
    }
    return -1;
}

/* 「無視」判定 / Detect ignore state */
function isIgnoreValue(value) {
    try { return (value === Spacing.SETIGNORE); } catch (e) { return false; }
}

})();
