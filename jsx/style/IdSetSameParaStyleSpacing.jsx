#target indesign

/*
 * IdSetSameParaStyleSpacing.jsx
 *
 * 段落スタイルの「同一スタイル間の段落間隔」を、スタイル定義そのものに対して設定します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdSetSameParaStyleSpacing";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-06-30";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSetSameParaStyleSpacing.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSetSameParaStyleSpacing.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

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

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 既定で選択する段落スタイル名 / Paragraph style preselected on launch */
var DEFAULT_STYLE_NAME = "p";

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
        title: { ja: "同一設定段落の間隔設定", en: "Set space between paragraphs with the same style" }
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
    },
    undo: {
        setSpacing: { ja: "同一設定段落の間隔を設定", en: "Set Spacing Between Same-Style Paragraphs" }
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

// =========================================
// 単位 / Units
// =========================================

/**
 * 単位の列挙値から表示用のラベルを返す
 * @param {MeasurementUnits} unit 対象の単位
 * @returns {string} 単位のラベル
 */
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

/**
 * 段落スタイルを集めてダイアログを表示し、間隔を適用する
 * @returns {void}
 */
function main() {
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDoc"));
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
            alert(getLabel("alert.noStyle"));
            return;
        }

        /* 選択中の段落スタイルを初期選択に使う / Preselect the current selection's style */
        var selectedStyle = getSelectedParagraphStyle();

        var settings = showDialog(styleEntries, unitText, selectedStyle);
        if (settings === null || settings.style === null) return;

        /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
        app.doScript(function () {
            /* 選択に依存せず、スタイル定義そのものを上書き / Override the style definition itself */
            if (settings.ignore) {
                settings.style.sameParaStyleSpacing = Spacing.SETIGNORE;
            } else {
                settings.style.sameParaStyleSpacing = settings.value;
            }
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.setSpacing"));
    } finally {
        app.scriptPreferences.measurementUnit = savedUnit;
    }
}

/* 段落スタイルを再帰収集（グループはパス付き）/ Collect paragraph styles recursively */
/**
 * グループを含めて段落スタイルを再帰的に集める
 * @param {object} container 段落スタイルまたはグループのコンテナ
 * @param {string} prefix グループ名の接頭辞
 * @param {Array<object>} list 収集先の配列
 * @returns {void}
 */
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

/**
 * 入力欄に上下キーでの増減操作を追加する
 * @param {EditText} editText 対象の入力欄
 * @returns {void}
 */
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

/**
 * 間隔を指定するダイアログを表示する
 * @param {Array<object>} styleEntries 段落スタイルの一覧
 * @param {string} unitText 表示する単位
 * @param {ParagraphStyle} selectedStyle 初期選択する段落スタイル
 * @returns {object|null} 設定内容。キャンセル時は null
 */
function showDialog(styleEntries, unitText, selectedStyle) {
    var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(dialog, 10);

    /* --- 段落スタイル / Paragraph style --- */
    var stylePanel = dialog.add("panel", undefined, getLabel("panel.style"));
    setupPanel(stylePanel, 6);

    var names = [];
    for (var i = 0; i < styleEntries.length; i++) names.push(styleEntries[i].name);
    var styleDropdown = stylePanel.add("dropdownlist", undefined, names);

    /* 既定選択：選択中スタイル → DEFAULT_STYLE_NAME → 先頭 / Default selection */
    var defaultIndex = indexOfStyle(styleEntries, selectedStyle);
    if (defaultIndex < 0) defaultIndex = indexOfName(styleEntries, DEFAULT_STYLE_NAME);
    if (defaultIndex < 0) defaultIndex = 0;
    styleDropdown.selection = defaultIndex;

    /* --- 段落間隔 / Spacing --- */
    var spacingPanel = dialog.add("panel", undefined, getLabel("panel.spacing"));
    setupPanel(spacingPanel, 6);

    var ignoreRadio = spacingPanel.add("radiobutton", undefined, getLabel("radio.ignore"));
    var zeroRadio = spacingPanel.add("radiobutton", undefined, getLabel("radio.zero"));

    var valueGroup = spacingPanel.add("group");
    var valueRadio = valueGroup.add("radiobutton", undefined, getLabel("radio.useValue"));
    var valueInput = valueGroup.add("edittext", undefined, "0");
    valueInput.characters = 6;
    changeValueByArrowKey(valueInput);
    valueGroup.add("statictext", undefined, unitText);

    /**
     * 間隔の指定方法を切り替える
     * @param {string} mode "ignore" / "zero" / "value"
     * @returns {void}
     */
    function setSpacingMode(mode) {
        ignoreRadio.value = (mode === "ignore");
        zeroRadio.value = (mode === "zero");
        valueRadio.value = (mode === "value");
        valueInput.enabled = (mode === "value");
    }
    ignoreRadio.onClick = function () { setSpacingMode("ignore"); };
    zeroRadio.onClick = function () { setSpacingMode("zero"); };
    valueRadio.onClick = function () { setSpacingMode("value"); };

    /**
     * 選択した段落スタイルの現在値を UI に反映する
     * @param {ParagraphStyle} style 対象の段落スタイル
     * @returns {void}
     */
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

    /* ボタン行（Mac 規約に合わせて Cancel → OK。幅いっぱいには広げない）
       / Button row: Cancel then OK per macOS convention, never stretched to full width */
    var buttonGroup = dialog.add("group");
    setupRow(buttonGroup, "center", 8);
    var cancelButton = buttonGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
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

/**
 * テキスト選択中の段落スタイルを取得する
 * @returns {ParagraphStyle|null} 段落スタイル。取得できない場合は null
 */
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

/**
 * 名前から一覧内の位置を探す
 * @param {Array<object>} entries 段落スタイルの一覧
 * @param {string} name 探す名前
 * @returns {number} 見つかった位置。なければ -1
 */
function indexOfName(entries, name) {
    if (!name) return -1;
    for (var i = 0; i < entries.length; i++) {
        if (entries[i].name === name) return i;
    }
    return -1;
}

/**
 * 段落スタイルから一覧内の位置を探す
 * @param {Array<object>} entries 段落スタイルの一覧
 * @param {ParagraphStyle} style 探す段落スタイル
 * @returns {number} 見つかった位置。なければ -1
 */
function indexOfStyle(entries, style) {
    if (!style) return -1;
    for (var i = 0; i < entries.length; i++) {
        if (entries[i].style.id === style.id) return i;
    }
    return -1;
}

/**
 * 「無視」を表す値かどうかを判定する
 * @param {*} value 判定する値
 * @returns {boolean} 「無視」なら true
 */
function isIgnoreValue(value) {
    try { return (value === Spacing.SETIGNORE); } catch (e) { return false; }
}

})();
