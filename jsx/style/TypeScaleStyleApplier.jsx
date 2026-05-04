#target indesign

/*
概要 / Overview

【日本語】
タイプスケールを使って、見出し・本文・キャプション用の段落スタイルを一括設定するInDesign用スクリプトです。
基準サイズ（本文）、倍率、見出しレベル数、サイズの丸め、行送り（%指定）、段落後のアキ、カーニング（本文／見出し別）を設定できます。

フォント指定オプションでは、次の3つのモードを選択できます：
・本文と見出しで共通
・本文と見出しで別々に指定
・フォントを変更しない

フォントファミリーを選択すると、そのフォントで使用可能なスタイル（ウエイト）を自動取得し、各段落スタイルに適用できます。
共通指定の場合は本文のフォント設定を見出しにも適用し、別々に指定する場合は見出し専用のフォント設定を使用します。

段落スタイルとサイズプレビューでは、各レベル（h1〜h6）、本文、キャプションのフォントサイズと行送りを確認できます。
ライブプレビューに対応しており、ダイアログ操作中に結果をリアルタイムで確認できます。

見出しは左揃え、本文・キャプションは均等配置（最終行左）に自動設定されます。

ダイアログを閉じる際には、オーバーライドをクリアして最終状態を適用します。

【English】
This InDesign script batch-configures paragraph styles for headings, body text, and captions using a type scale.
You can define the base body size, scale ratio, number of heading levels, size rounding, leading (percentage-based), space after, and kerning (separately for body and headings).

The font options panel provides three modes:
- Use same font for body and headings
- Specify separately for body and headings
- Do not change fonts

When a font family is selected, available font styles (weights) are automatically collected and applied.
When using shared mode, heading styles inherit the body font settings; when using separate mode, headings use their own font settings.

The preview panel displays font size and leading for each level (h1–h6), body, and caption.
Live preview allows real-time confirmation while adjusting settings.

Headings are set to left alignment, while body text and captions use left-justified alignment.

When the dialog is closed, overrides are cleared and final settings are applied.
*/


/* =========================================

   ユーザー設定 / User Settings

   ========================================= */

var DEFAULT_BASE_SIZE_PT = 9; // 基準サイズ（本文, pt） / Base body size (pt)
var DEFAULT_BASE_SIZE_Q = 13; // 基準サイズ（本文, Q） / Base body size (Q)
var DEFAULT_RATIO = 1.309; // スケール倍率 / Scale ratio
var DEFAULT_LEVEL_COUNT = 4; // 見出しレベル数 / Heading levels

var DEFAULT_BODY_LEADING_PERCENT = 160; // 本文の行送り(%)
var DEFAULT_HEADING_LEADING_PERCENT = 115; // 見出しの行送り(%)
var DEFAULT_SPACE_AFTER_PERCENT = 10; // 段落後のアキ(%)

/* =========================================

    バージョンとローカライズ / Version and Localization

   ========================================= */

var SCRIPT_VERSION = "v1.0.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "タイプスケールで一括設定", en: "Type Scale Settings" },
    errNoDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    errInvalidBaseSize: { ja: "基準サイズ（本文）には正の数値を入力してください。", en: "Enter a positive number for Base Size (Body)." },
    errMissingParagraphStyle: { ja: "段落スタイル「%1」が見つかりません。", en: "Paragraph style \"%1\" was not found." },
    noFontChange: { ja: "（変更しない）", en: "(No change)" },
    textSettingsPanel: { ja: "テキスト設定", en: "Text Settings" },
    fontSettingsPanel: { ja: "フォント指定オプション", en: "Font Options" },
    disableFontSelection: { ja: "フォントを変更しない", en: "Do not change fonts" },
    useSameFontForBodyAndHeading: { ja: "本文と見出しで共通", en: "Use same font for body and headings" },
    separateFontForBodyAndHeading: { ja: "本文と見出しで別々に指定", en: "Specify separately" },
    bodyTextPanel: { ja: "本文", en: "Body" },
    headingTextPanel: { ja: "見出し", en: "Headings" },
    baseSizeBody: { ja: "基準サイズ", en: "Base Size (Body)" },
    font: { ja: "フォント", en: "Font" },
    fontStyle: { ja: "スタイル", en: "Style" },
    bodyLeadingRatio: { ja: "行送り", en: "Leading Ratio (Body)" },
    headingLeadingRatio: { ja: "行送り", en: "Leading Ratio (Headings)" },
    spaceAfterRatio: { ja: "段落後のアキ", en: "Space After" },
    kerningMethod: { ja: "カーニング", en: "Kerning Method" },
    kerningJapaneseMono: { ja: "和文等幅", en: "Japanese Mono" },
    kerningMetrics: { ja: "メトリクス", en: "Metrics" },
    kerningOptical: { ja: "オプティカル", en: "Optical" },
    scaleSettingsPanel: { ja: "スケール設定", en: "Scale Settings" },
    scaleRatio: { ja: "倍率", en: "Scale Ratio" },
    headingLevelCount: { ja: "見出しレベル数", en: "Heading Levels" },
    sizeRounding: { ja: "サイズの丸め", en: "Size Rounding" },
    roundInteger: { ja: "整数", en: "Integer" },
    fontStyleHeader: { ja: "スタイル", en: "Style" },
    roundFirstDecimal: { ja: "小数点第1位", en: "1 decimal place" },
    roundSecondDecimal: { ja: "小数点第2位", en: "2 decimal places" },
    previewPanel: { ja: "段落スタイルとサイズプレビュー", en: "Paragraph Styles & Size Preview" },
    levelHeader: { ja: "レベル", en: "Level" },
    fontSizeHeader: { ja: "フォントサイズ", en: "Font Size" },
    leadingHeader: { ja: "行送り", en: "Leading" },
    paragraphStyleHeader: { ja: "段落スタイル", en: "Paragraph Style" },
    levelPrefix: { ja: "レベル", en: "Level " },
    baseBodyPreview: { ja: "基準（本文）", en: "Base (Body)" },
    captionPreview: { ja: "キャプション", en: "Caption" },
    livePreview: { ja: "ライブプレビュー", en: "Live Preview" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" },
    loadingTitle: { ja: "処理中", en: "Processing" },
    loadingFonts: { ja: "フォント情報を読み込んでいます…", en: "Loading font information..." }
};

function L(key) {
    return (LABELS[key] && LABELS[key][lang]) ? LABELS[key][lang] : key;
}

/* コロン付きラベル（日本語は全角、英語は半角） / Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return L(key) + (lang === "ja" ? "：" : ":");
}

function formatLabel(key, value) {
    return L(key).replace("%1", value);
}

function createProgressPalette(message) {
    var palette = new Window("palette", L("loadingTitle"));
    palette.orientation = "column";
    palette.alignChildren = ["fill", "center"];
    palette.margins = [20, 16, 20, 16];
    palette.spacing = 10;

    var messageText = palette.add("statictext", undefined, message);
    messageText.preferredSize.width = 260;

    var progressBar = palette.add("progressbar", undefined, 0, 100);
    progressBar.preferredSize.width = 260;
    progressBar.value = 35;

    palette.show();
    try { palette.update(); } catch (e) { }
    return palette;
}

function closeProgressPalette(palette) {
    if (!palette) return;
    try { palette.close(); } catch (e) { }
}

(function () {

    if (app.documents.length === 0) {
        alert(L("errNoDocument"));
        return;
    }

    var targetDocument = app.activeDocument;
    var unit = getTextSizeUnit();

    var defaultBaseSize = DEFAULT_BASE_SIZE_PT;
    try {
        if (unit === MeasurementUnits.POINTS) {
            defaultBaseSize = DEFAULT_BASE_SIZE_PT;
        } else if (unit === MeasurementUnits.Q) {
            defaultBaseSize = DEFAULT_BASE_SIZE_Q;
        } else if (unit === MeasurementUnits.MILLIMETERS) {
            defaultBaseSize = DEFAULT_BASE_SIZE_Q * 0.25; // Q→mm
        }
    } catch (e) { }

    var typescaleSettings = showTypescaleDialog(targetDocument, defaultBaseSize, DEFAULT_RATIO, DEFAULT_LEVEL_COUNT, unit);

    if (typescaleSettings !== null) {
        applyTypescaleSettings(targetDocument, typescaleSettings, false, unit);
    }

    function computeSizes(base, ratio, levelCount) {
        /* 本文サイズを基準に、タイプスケールで見出しとキャプションのサイズを算出 / Calculate heading and caption sizes from the body size using the type scale */
        /* h1 が最大、h<levelCount> が最小の見出しサイズになる / h1 is the largest heading, and h<levelCount> is the smallest heading */
        /* レベル k = base × ratio^(levelCount - k + 1) / Level k = base × ratio^(levelCount - k + 1) */
        /* キャプション = base / ratio / Caption = base / ratio */
        var headingSizes = [];
        for (var levelIndex = 1; levelIndex <= levelCount; levelIndex++) {
            headingSizes.push(base * Math.pow(ratio, levelCount - levelIndex + 1));
        }
        return { headingSizes: headingSizes, base: base, caption: base / ratio };
    }

    function applyTypescaleSettings(targetDocument, typescaleSettings, silent, unit) {
        var computedSizes = computeSizes(typescaleSettings.baseSize, typescaleSettings.ratio, typescaleSettings.levelCount);
        // 無効値時はデフォルトにフォールバック
        var spaceAfterPercent = (typeof typescaleSettings.spaceAfterPercent === "number" && typescaleSettings.spaceAfterPercent >= 0)
            ? typescaleSettings.spaceAfterPercent
            : DEFAULT_SPACE_AFTER_PERCENT;
        function applyParagraphStyleSettings(styleName, sizeInUnit, leadingMult, isHeading, fontFamilyName, fontStyleName) {
            if (!styleName) return;
            var rounded = roundTo(sizeInUnit, typescaleSettings.roundDigits);
            var sizePt = toPoints(rounded, unit);
            var leadingPt = (typeof leadingMult === "number") ? sizePt * leadingMult : null;
            var spaceAfterPt = isHeading ? sizePt * spaceAfterPercent / 100 : null;
            // フォントファミリー＋スタイルで解決。未指定時はファミリー内の推奨スタイルを使用
            var fontToUse = null;
            if (fontFamilyName) {
                fontToUse = fontStyleName ? findFontByFamilyAndStyle(fontFamilyName, fontStyleName) : findFontInFamily(fontFamilyName);
            }
            var kerningMethod = isHeading ? typescaleSettings.headingKerningMethod : typescaleSettings.bodyKerningMethod;
            setParagraphStyleProps(targetDocument, styleName, sizePt, fontToUse, leadingPt, spaceAfterPt, kerningMethod, silent);
        }
        applyParagraphStyleSettings(typescaleSettings.baseStyleName, computedSizes.base, typescaleSettings.bodyLeading, false, typescaleSettings.fontFamily, typescaleSettings.baseFontStyleName);
        applyParagraphStyleSettings(typescaleSettings.captionStyleName, computedSizes.caption, typescaleSettings.bodyLeading, false, typescaleSettings.fontFamily, typescaleSettings.captionFontStyleName);
        for (var levelNumber = 1; levelNumber <= typescaleSettings.levelCount; levelNumber++) {
            var name = typescaleSettings.levelStyleNames && typescaleSettings.levelStyleNames[levelNumber - 1];
            var fontStyleName = typescaleSettings.levelFontStyleNames && typescaleSettings.levelFontStyleNames[levelNumber - 1];
            applyParagraphStyleSettings(name, computedSizes.headingSizes[levelNumber - 1], typescaleSettings.headingLeading, true, typescaleSettings.headingFontFamily, fontStyleName);
        }
    }

    function getTextSizeUnit() {
        // テキスト単位優先 → 定規単位 → ptにフォールバック
        try {
            var textSizeUnit = app.viewPreferences.textSizeMeasurementUnits;
            if (textSizeUnit !== undefined && textSizeUnit !== null) return textSizeUnit;
        } catch (e) { }
        try {
            return app.viewPreferences.horizontalMeasurementUnits;
        } catch (e) { }
        return MeasurementUnits.POINTS;
    }

    function getMeasurementUnitValue(unitName) {
        try {
            return MeasurementUnits[unitName];
        } catch (e) { }
        return null;
    }

    function unitSymbol(unit) {
        var unitLabels = [
            { name: "POINTS", label: "pt" },
            { name: "MILLIMETERS", label: "mm" },
            { name: "CENTIMETERS", label: "cm" },
            { name: "INCHES", label: "inch" },
            { name: "INCHES_DECIMAL", label: "inch" },
            { name: "PICAS", label: "pica" },
            { name: "CICEROS", label: "c" },
            { name: "AGATES", label: "ag" },
            { name: "PIXELS", label: "px" },
            { name: "Q", label: "Q" },
            { name: "HA", label: "H" }
        ];

        for (var unitIndex = 0; unitIndex < unitLabels.length; unitIndex++) {
            if (unit === getMeasurementUnitValue(unitLabels[unitIndex].name)) {
                return unitLabels[unitIndex].label;
            }
        }
        return "pt";
    }

    function toPoints(value, unit) {
        // 各種単位をポイントに変換（内部計算はpt基準）
        var pointConverters = [
            { name: "POINTS", factor: 1 },
            { name: "MILLIMETERS", factor: 2.834645669 },
            { name: "CENTIMETERS", factor: 28.34645669 },
            { name: "INCHES", factor: 72 },
            { name: "INCHES_DECIMAL", factor: 72 },
            { name: "PICAS", factor: 12 },
            { name: "CICEROS", factor: 12.7896 },
            { name: "PIXELS", factor: 1 },
            { name: "Q", factor: 0.708661417 }, /* 1Q = 0.25mm / 1Q = 0.25mm */
            { name: "HA", factor: 0.708661417 }
        ];

        for (var unitIndex = 0; unitIndex < pointConverters.length; unitIndex++) {
            if (unit === getMeasurementUnitValue(pointConverters[unitIndex].name)) {
                return value * pointConverters[unitIndex].factor;
            }
        }
        return value;
    }

    function roundTo(num, places) {
        var factor = Math.pow(10, places);
        return Math.round(num * factor) / factor;
    }

    function arrayContains(array, value) {
        for (var itemIndex = 0; itemIndex < array.length; itemIndex++) {
            if (array[itemIndex] === value) return true;
        }
        return false;
    }

    function getParagraphStyleNames(targetDocument) {
        var names = [];
        var styles = targetDocument.allParagraphStyles;
        for (var styleIndex = 0; styleIndex < styles.length; styleIndex++) {
            names.push(styles[styleIndex].name);
        }
        return names;
    }

    var _fontInfo = null;

    function getFontInfo() {
        if (_fontInfo) return _fontInfo;

        var fonts = [];
        var fontMap = {};
        var fontObjectMap = {};
        var families = [];

        try { fonts = app.fonts.everyItem().getElements(); } catch (e1) { fonts = []; }

        for (var fontIndex = 0; fontIndex < fonts.length; fontIndex++) {
            try {
                var fontItem = fonts[fontIndex];
                var familyName = fontItem.fontFamily;
                var styleName = fontItem.fontStyleName;

                if (!familyName || !styleName) continue;

                if (!fontMap[familyName]) {
                    fontMap[familyName] = [];
                    families.push(familyName);
                }

                if (!arrayContains(fontMap[familyName], styleName)) {
                    fontMap[familyName].push(styleName);
                }

                fontObjectMap[familyName + "\t" + styleName] = fontItem;
            } catch (e2) { }
        }

        families.sort();

        for (var familyIndex = 0; familyIndex < families.length; familyIndex++) {
            fontMap[families[familyIndex]].sort();
        }

        _fontInfo = {
            families: families,
            fontMap: fontMap,
            fontObjectMap: fontObjectMap
        };
        return _fontInfo;
    }

    function getFontFamilyNames() {
        return getFontInfo().families.slice(0);
    }

    function getPreferredFontStyleName(styleNames) {
        if (!styleNames || styleNames.length === 0) return null;

        var preferredStyleNames = [
            "Thin",
            "ExtraLight",
            "Extra Light",
            "UltraLight",
            "Ultra Light",
            "Light",
            "EL",
            "L",
            "W1",
            "W2",
            "100",
            "200",
            "細",
            "極細",
            "ライト",
            "Regular",
            "Roman",
            "Book",
            "Normal",
            "Medium",
            "R",
            "レギュラー",
            "標準",
            "中"
        ];

        for (var preferredIndex = 0; preferredIndex < preferredStyleNames.length; preferredIndex++) {
            for (var styleIndex = 0; styleIndex < styleNames.length; styleIndex++) {
                if (styleNames[styleIndex] === preferredStyleNames[preferredIndex]) {
                    return styleNames[styleIndex];
                }
            }
        }

        return styleNames[0];
    }

    function getFontFullName(familyName, styleName) {
        if (!familyName || !styleName) return null;
        return familyName + "\t" + styleName;
    }

    function findFontInFamily(familyName) {
        if (!familyName) return null;
        var styleNames = getFontStylesInFamily(familyName);
        var preferredStyleName = getPreferredFontStyleName(styleNames);
        if (!preferredStyleName) return null;
        return findFontByFamilyAndStyle(familyName, preferredStyleName);
    }

    function getFontStylesInFamily(familyName) {
        if (!familyName) return [];
        var fontInfo = getFontInfo();
        if (!fontInfo.fontMap[familyName]) return [];
        return fontInfo.fontMap[familyName].slice(0);
    }

    function findFontByFamilyAndStyle(familyName, styleName) {
        var fontFullName = getFontFullName(familyName, styleName);
        if (!fontFullName) return null;

        var fontInfo = getFontInfo();
        if (fontInfo.fontObjectMap[fontFullName]) {
            return fontInfo.fontObjectMap[fontFullName];
        }
        return null;
    }

    function selectDropdownByText(dropdownList, text) {
        for (var itemIndex = 0; itemIndex < dropdownList.items.length; itemIndex++) {
            if (dropdownList.items[itemIndex].text === text) {
                dropdownList.selection = itemIndex;
                return true;
            }
        }
        if (dropdownList.items.length > 0) dropdownList.selection = 0;
        return false;
    }

    function getDropdownText(dropdownList) {
        return dropdownList.selection ? dropdownList.selection.text : null;
    }

    function showTypescaleDialog(targetDocument, defaultBase, defaultRatio, defaultLevelCount, unit) {
        var unitSym = unitSymbol(unit);
        var styleNames = getParagraphStyleNames(targetDocument);
        var loadingPalette = createProgressPalette(L("loadingFonts"));
        var fontFamilies = getFontFamilyNames();
        closeProgressPalette(loadingPalette);
        var fontOptions = [L("noFontChange")].concat(fontFamilies);
        var roundOptions = [
            { label: L("roundInteger"), digits: 0 },
            { label: L("roundFirstDecimal"), digits: 1 },
            { label: L("roundSecondDecimal"), digits: 2 }
        ];
        var defaultRoundDigits = 0;
        var kerningOptions = [
            { label: L("kerningJapaneseMono"), value: "和文等幅" },
            { label: L("kerningMetrics"), value: "メトリクス" },
            { label: L("kerningOptical"), value: "オプティカル" }
        ];
        var PANEL_MARGINS = [15, 20, 15, 10];

        function setupPanel(panel, spacing) {
            panel.orientation = "column";
            panel.alignChildren = "left";
            panel.alignment = "fill";
            panel.margins = PANEL_MARGINS;
            if (typeof spacing === "number") {
                panel.spacing = spacing;
            }
        }

        function addFixedWidthLabel(parent, text, width) {
            var label = parent.add("statictext", undefined, text);
            label.preferredSize.width = width;
            return label;
        }

        function addLabeledGroup(panel, labelText, labelWidth) {
            var group = panel.add("group");
            group.orientation = "row";
            group.alignChildren = ["left", "center"];
            addFixedWidthLabel(group, labelText, labelWidth);
            return group;
        }

        function changeValueByArrowKey(editText, onValueChange) {
            editText.addEventListener("keydown", function (event) {
                var value = Number(editText.text);
                if (isNaN(value)) return;

                var keyboard = ScriptUI.environment.keyboardState;
                var delta = 1;

                if (keyboard.shiftKey) {
                    delta = 10;
                    if (event.keyName === "Up") {
                        value = Math.ceil((value + 1) / delta) * delta;
                    } else if (event.keyName === "Down") {
                        value = Math.floor((value - 1) / delta) * delta;
                    } else {
                        return;
                    }
                } else if (keyboard.altKey) {
                    delta = 0.1;
                    if (event.keyName === "Up") {
                        value += delta;
                    } else if (event.keyName === "Down") {
                        value -= delta;
                    } else {
                        return;
                    }
                } else {
                    delta = 1;
                    if (event.keyName === "Up") {
                        value += delta;
                    } else if (event.keyName === "Down") {
                        value -= delta;
                    } else {
                        return;
                    }
                }

                if (keyboard.altKey) {
                    value = Math.round(value * 10) / 10;
                } else {
                    value = Math.round(value);
                }

                if (value < 0) value = 0;

                event.preventDefault();
                editText.text = String(value);

                if (typeof onValueChange === "function") {
                    onValueChange();
                }
            });
        }

        var ratioOptions = [
            { name: "Minor Second", value: 1.067 },
            { name: "Major Second", value: 1.125 },
            { name: "Minor Third", value: 1.2 },
            { name: "Major Third", value: 1.25 },
            { name: "Golden Ratio: ½", value: 1.309 },
            { name: "Perfect Fourth", value: 1.333 },
            { name: "Augmented Fourth", value: 1.414 },
            { name: "Golden Ratio", value: 1.618 }
        ];
        var levelOptions = [3, 4, 5, 6];

        function getKerningOptionLabels() {
            var labels = [];
            for (var kerningOptionIndex = 0; kerningOptionIndex < kerningOptions.length; kerningOptionIndex++) {
                labels.push(kerningOptions[kerningOptionIndex].label);
            }
            return labels;
        }

        function selectKerningDropdownByValue(dropdownList, value) {
            for (var kerningOptionIndex = 0; kerningOptionIndex < kerningOptions.length; kerningOptionIndex++) {
                if (kerningOptions[kerningOptionIndex].value === value) {
                    dropdownList.selection = kerningOptionIndex;
                    return;
                }
            }
            dropdownList.selection = 0;
        }


        function createBodyTextPanel(parent, labelWidth) {
            var bodyPanel = parent.add("panel", undefined, L("bodyTextPanel"));
            setupPanel(bodyPanel, 6);
            bodyPanel.alignment = ["fill", "top"];

            var fontGrp = addLabeledGroup(bodyPanel, labelText("font"), labelWidth);
            var fontDD = fontGrp.add("dropdownlist", undefined, fontOptions);
            fontDD.preferredSize.width = 180;
            fontDD.selection = 0;

            var fontStyleGrp = addLabeledGroup(bodyPanel, labelText("fontStyle"), labelWidth);
            var fontStyleDD = fontStyleGrp.add("dropdownlist", undefined, [L("noFontChange")]);
            fontStyleDD.preferredSize.width = 130;
            fontStyleDD.selection = 0;

            var leadingBodyGrp = addLabeledGroup(bodyPanel, labelText("bodyLeadingRatio"), labelWidth);
            var leadingBodyInput = leadingBodyGrp.add("edittext", undefined, String(DEFAULT_BODY_LEADING_PERCENT));
            leadingBodyInput.characters = 4;
            leadingBodyGrp.add("statictext", undefined, "%");

            var bodyKerningGrp = addLabeledGroup(bodyPanel, labelText("kerningMethod"), labelWidth);
            var bodyKerningDD = bodyKerningGrp.add("dropdownlist", undefined, getKerningOptionLabels());
            bodyKerningDD.preferredSize.width = 110;
            selectKerningDropdownByValue(bodyKerningDD, "和文等幅");

            return {
                fontDD: fontDD,
                fontStyleDD: fontStyleDD,
                leadingBodyInput: leadingBodyInput,
                bodyKerningDD: bodyKerningDD
            };
        }

        function createHeadingTextPanel(parent, labelWidth) {
            var headingPanel = parent.add("panel", undefined, L("headingTextPanel"));
            setupPanel(headingPanel, 4);
            headingPanel.alignment = ["fill", "top"];

            // Font controls (like body panel)
            var headingFontGrp = addLabeledGroup(headingPanel, labelText("font"), labelWidth);
            var headingFontDD = headingFontGrp.add("dropdownlist", undefined, fontOptions);
            headingFontDD.preferredSize.width = 180;
            headingFontDD.selection = 0;

            var headingFontStyleGrp = addLabeledGroup(headingPanel, labelText("fontStyle"), labelWidth);
            var headingFontStyleDD = headingFontStyleGrp.add("dropdownlist", undefined, [L("noFontChange")]);
            headingFontStyleDD.preferredSize.width = 130;
            headingFontStyleDD.selection = 0;

            var leadingHeadingGrp = addLabeledGroup(headingPanel, labelText("headingLeadingRatio"), labelWidth);
            var leadingHeadingInput = leadingHeadingGrp.add("edittext", undefined, String(DEFAULT_HEADING_LEADING_PERCENT));
            leadingHeadingInput.characters = 4;
            leadingHeadingGrp.add("statictext", undefined, "%");

            var headingKerningGrp = addLabeledGroup(headingPanel, labelText("kerningMethod"), labelWidth);
            var headingKerningDD = headingKerningGrp.add("dropdownlist", undefined, getKerningOptionLabels());
            headingKerningDD.preferredSize.width = 110;
            selectKerningDropdownByValue(headingKerningDD, "メトリクス");

            var spaceAfterGrp = addLabeledGroup(headingPanel, labelText("spaceAfterRatio"), labelWidth);
            var spaceAfterInput = spaceAfterGrp.add("edittext", undefined, String(DEFAULT_SPACE_AFTER_PERCENT));
            spaceAfterInput.characters = 4;
            spaceAfterGrp.add("statictext", undefined, "%");

            return {
                headingFontDD: headingFontDD,
                headingFontStyleDD: headingFontStyleDD,
                leadingHeadingInput: leadingHeadingInput,
                spaceAfterInput: spaceAfterInput,
                headingKerningDD: headingKerningDD
            };
        }

        function createTextSettingsPanel(dialog) {
            var basicPanel = dialog.add("group");
            basicPanel.orientation = "column";
            basicPanel.alignChildren = "fill";
            basicPanel.alignment = "fill";
            basicPanel.spacing = 6;

            var BODY_LABEL_WIDTH = 80;
            var HEADING_LABEL_WIDTH = 94;

            var textColumnGroup = basicPanel.add("group");
            textColumnGroup.orientation = "row";
            textColumnGroup.alignChildren = ["fill", "top"];
            textColumnGroup.alignment = "fill";
            textColumnGroup.spacing = 10;

            var bodyUi = createBodyTextPanel(textColumnGroup, BODY_LABEL_WIDTH);
            var headingUi = createHeadingTextPanel(textColumnGroup, HEADING_LABEL_WIDTH);

            return {
                fontDD: bodyUi.fontDD,
                fontStyleDD: bodyUi.fontStyleDD,
                headingFontDD: headingUi.headingFontDD,
                headingFontStyleDD: headingUi.headingFontStyleDD,
                leadingBodyInput: bodyUi.leadingBodyInput,
                leadingHeadingInput: headingUi.leadingHeadingInput,
                spaceAfterInput: headingUi.spaceAfterInput,
                bodyKerningDD: bodyUi.bodyKerningDD,
                headingKerningDD: headingUi.headingKerningDD
            };
        }

        function createFontSettingsPanel(dialog) {
            var fontPanel = dialog.add("panel", undefined, L("fontSettingsPanel"));
            fontPanel.alignment = ["fill", "top"];
            setupPanel(fontPanel, 6);

            var modeGroup = fontPanel.add("group");
            modeGroup.orientation = "column";
            modeGroup.alignChildren = ["left", "center"];

            var useSameFontRadio = modeGroup.add("radiobutton", undefined, L("useSameFontForBodyAndHeading"));
            var separateFontRadio = modeGroup.add("radiobutton", undefined, L("separateFontForBodyAndHeading"));
            var disableFontRadio = modeGroup.add("radiobutton", undefined, L("disableFontSelection"));

            useSameFontRadio.value = true;
            separateFontRadio.value = false;
            disableFontRadio.value = false;

            return {
                disableFontRadio: disableFontRadio,
                useSameFontRadio: useSameFontRadio,
                separateFontRadio: separateFontRadio
            };
        }

        function createScaleSettingsPanel(dialog) {
            var optionsPanel = dialog.add("panel", undefined, L("scaleSettingsPanel"));
            optionsPanel.alignment = ["fill", "top"];
            setupPanel(optionsPanel, 6);

            var OPTIONS_LABEL_WIDTH = 110;

            var baseGrp = addLabeledGroup(optionsPanel, labelText("baseSizeBody"), OPTIONS_LABEL_WIDTH);
            var baseInput = baseGrp.add("edittext", undefined, String(defaultBase));
            baseInput.characters = 4;
            baseGrp.add("statictext", undefined, unitSym);

            var ratioGrp = addLabeledGroup(optionsPanel, labelText("scaleRatio"), OPTIONS_LABEL_WIDTH);
            var ratioLabels = [];
            for (var ratioIndex = 0; ratioIndex < ratioOptions.length; ratioIndex++) {
                ratioLabels.push(ratioOptions[ratioIndex].name + "  " + ratioOptions[ratioIndex].value);
            }
            var ratioDD = ratioGrp.add("dropdownlist", undefined, ratioLabels);
            ratioDD.preferredSize.width = 200;
            for (var selectedRatioIndex = 0; selectedRatioIndex < ratioOptions.length; selectedRatioIndex++) {
                if (ratioOptions[selectedRatioIndex].value === defaultRatio) { ratioDD.selection = selectedRatioIndex; break; }
            }
            if (!ratioDD.selection) ratioDD.selection = 0;

            var levelGrp = addLabeledGroup(optionsPanel, labelText("headingLevelCount"), OPTIONS_LABEL_WIDTH);
            levelGrp.alignChildren = ["left", "bottom"];
            var levelRadios = [];
            for (var levelOptionIndex = 0; levelOptionIndex < levelOptions.length; levelOptionIndex++) {
                var levelRadio = levelGrp.add("radiobutton", undefined, String(levelOptions[levelOptionIndex]));
                levelRadio.alignment = ["left", "bottom"];
                if (levelOptions[levelOptionIndex] === defaultLevelCount) levelRadio.value = true;
                levelRadios.push(levelRadio);
            }
            if (!getSelectedRadioValue(levelRadios, levelOptions, null, null)) {
                levelRadios[0].value = true;
            }

            var roundGrp = addLabeledGroup(optionsPanel, labelText("sizeRounding"), OPTIONS_LABEL_WIDTH);
            roundGrp.alignChildren = ["left", "bottom"];
            var roundRadios = [];
            for (var roundOptionIndex = 0; roundOptionIndex < roundOptions.length; roundOptionIndex++) {
                var roundRadio = roundGrp.add("radiobutton", undefined, roundOptions[roundOptionIndex].label);
                roundRadio.alignment = ["left", "bottom"];
                if (roundOptions[roundOptionIndex].digits === defaultRoundDigits) roundRadio.value = true;
                roundRadios.push(roundRadio);
            }

            return {
                baseInput: baseInput,
                ratioDD: ratioDD,
                levelRadios: levelRadios,
                roundRadios: roundRadios
            };
        }

        function createPreviewPanel(dialog) {
            var previewPanel = dialog.add("panel", undefined, L("previewPanel"));
            setupPanel(previewPanel, 2);

            var PREVIEW_LABEL_WIDTH = 100;
            var PREVIEW_SIZE_WIDTH = 110;
            var PREVIEW_LEADING_WIDTH = 70;

            var headerRow = previewPanel.add("group");
            headerRow.orientation = "row";
            headerRow.alignChildren = "left";
            headerRow.add("statictext", undefined, labelText("levelHeader")).preferredSize.width = PREVIEW_LABEL_WIDTH;
            headerRow.add("statictext", undefined, labelText("fontSizeHeader")).preferredSize.width = PREVIEW_SIZE_WIDTH;
            headerRow.add("statictext", undefined, labelText("leadingHeader")).preferredSize.width = PREVIEW_LEADING_WIDTH;
            headerRow.add("statictext", undefined, labelText("paragraphStyleHeader")).characters = 14;
            var fontStyleHeader = headerRow.add("statictext", undefined, labelText("fontStyleHeader"));
            fontStyleHeader.characters = 12;
            fontStyleHeader.enabled = true;

            var previewHeaderSpacer = previewPanel.add("group");
            previewHeaderSpacer.preferredSize.height = 4;

            function createPreviewRow(parent, label, defaultStyleName) {
                var row = parent.add("group");
                row.orientation = "row";
                row.alignChildren = "center";
                var labelText = row.add("statictext", undefined, label);
                labelText.preferredSize.width = PREVIEW_LABEL_WIDTH;
                var sizeText = row.add("statictext", undefined, "");
                sizeText.preferredSize.width = PREVIEW_SIZE_WIDTH;
                var leadingText = row.add("statictext", undefined, "");
                leadingText.preferredSize.width = PREVIEW_LEADING_WIDTH;
                var styleDD = row.add("dropdownlist", undefined, styleNames);
                styleDD.preferredSize.width = 140;
                selectDropdownByText(styleDD, defaultStyleName);

                var fontStyleDD = row.add("dropdownlist", undefined, [L("noFontChange")]);
                fontStyleDD.preferredSize.width = 130;
                fontStyleDD.selection = 0;
                fontStyleDD.enabled = true;

                return {
                    lbl: labelText,
                    sizeText: sizeText,
                    leadingText: leadingText,
                    styleDD: styleDD,
                    fontStyleDD: fontStyleDD
                };
            }

            var levelRows = [];
            for (var levelNumber = 1; levelNumber <= 6; levelNumber++) {
                levelRows.push(createPreviewRow(previewPanel, L("levelPrefix") + levelNumber, "h" + levelNumber));
            }

            return {
                levelRows: levelRows,
                baseRow: createPreviewRow(previewPanel, L("baseBodyPreview"), "p"),
                captionRow: createPreviewRow(previewPanel, L("captionPreview"), "p.caption")
            };
        }

        function createButtonRow(dialog) {
            var bottomRow = dialog.add("group");
            bottomRow.margins = [0, 10, 0, 0];
            bottomRow.orientation = "row";
            bottomRow.alignment = "fill";
            bottomRow.alignChildren = ["fill", "center"];

            var leftButtonColumn = bottomRow.add("group");
            leftButtonColumn.orientation = "row";
            leftButtonColumn.alignChildren = ["left", "center"];

            var previewCheck = leftButtonColumn.add("checkbox", undefined, L("livePreview"));
            previewCheck.value = true;

            var centerButtonColumn = bottomRow.add("group");
            centerButtonColumn.alignment = ["fill", "fill"];
            centerButtonColumn.minimumSize.width = 0;

            var rightButtonColumn = bottomRow.add("group");
            rightButtonColumn.orientation = "row";
            rightButtonColumn.alignChildren = ["right", "center"];
            rightButtonColumn.alignment = ["right", "center"];

            rightButtonColumn.add("button", undefined, L("cancel"), { name: "cancel" });
            rightButtonColumn.add("button", undefined, L("ok"), { name: "ok" });

            return { previewCheck: previewCheck };
        }

        function createTypescaleDialog() {
            var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
            dlg.orientation = "column";
            dlg.alignChildren = "fill";
            dlg.margins = 16;
            dlg.spacing = 10;

            var textSettingsUi = createTextSettingsPanel(dlg);

            var optionColumnGroup = dlg.add("group");
            optionColumnGroup.orientation = "row";
            optionColumnGroup.alignChildren = ["fill", "top"];
            optionColumnGroup.alignment = "fill";
            optionColumnGroup.spacing = 10;

            var scaleSettingsUi = createScaleSettingsPanel(optionColumnGroup);
            var fontSettingsUi = createFontSettingsPanel(optionColumnGroup);
            var previewUi = createPreviewPanel(dlg);

            var buttonUi = createButtonRow(dlg);

            return {
                dialog: dlg,
                baseInput: scaleSettingsUi.baseInput,
                fontDD: textSettingsUi.fontDD,
                fontStyleDD: textSettingsUi.fontStyleDD,
                headingFontDD: textSettingsUi.headingFontDD,
                headingFontStyleDD: textSettingsUi.headingFontStyleDD,
                leadingBodyInput: textSettingsUi.leadingBodyInput,
                leadingHeadingInput: textSettingsUi.leadingHeadingInput,
                spaceAfterInput: textSettingsUi.spaceAfterInput,
                bodyKerningDD: textSettingsUi.bodyKerningDD,
                headingKerningDD: textSettingsUi.headingKerningDD,
                ratioDD: scaleSettingsUi.ratioDD,
                levelRadios: scaleSettingsUi.levelRadios,
                roundRadios: scaleSettingsUi.roundRadios,
                levelRows: previewUi.levelRows,
                baseRow: previewUi.baseRow,
                captionRow: previewUi.captionRow,
                previewCheck: buttonUi.previewCheck,
                disableFontRadio: fontSettingsUi.disableFontRadio,
                useSameFontRadio: fontSettingsUi.useSameFontRadio,
                separateFontRadio: fontSettingsUi.separateFontRadio,
            };
        }

        var dialogUi = createTypescaleDialog();

        function parsePositiveNumber(text, fallbackValue) {
            var value = parseFloat(text);
            return (isNaN(value) || value <= 0) ? fallbackValue : value;
        }

        function parseNonNegativeNumber(text, fallbackValue) {
            var value = parseFloat(text);
            return (isNaN(value) || value < 0) ? fallbackValue : value;
        }

        function parsePositivePercentMultiplier(text, fallbackValue) {
            var value = parseFloat(text);
            return (isNaN(value) || value <= 0) ? fallbackValue : value / 100;
        }

        function getSelectedRadioValue(radioButtons, options, valueKey, fallbackValue) {
            for (var radioIndex = 0; radioIndex < radioButtons.length; radioIndex++) {
                if (radioButtons[radioIndex].value) {
                    return valueKey ? options[radioIndex][valueKey] : options[radioIndex];
                }
            }
            return fallbackValue;
        }

        function getSelectedDropdownOptionValue(dropdownList, options, valueKey, fallbackValue) {
            if (!dropdownList.selection) return fallbackValue;
            return valueKey ? options[dropdownList.selection.index][valueKey] : options[dropdownList.selection.index];
        }

        function getCurrentRatio(dialogUi) {
            if (!dialogUi.ratioDD.selection) return defaultRatio;
            return ratioOptions[dialogUi.ratioDD.selection.index].value;
        }

        function getCurrentLevelCount(dialogUi) {
            return getSelectedRadioValue(dialogUi.levelRadios, levelOptions, null, defaultLevelCount);
        }

        function getCurrentRoundDigits(dialogUi) {
            return getSelectedRadioValue(dialogUi.roundRadios, roundOptions, "digits", defaultRoundDigits);
        }

        function getSelectedFontFamily(dialogUi) {
            if (dialogUi.disableFontRadio && dialogUi.disableFontRadio.value) return null;
            if (!dialogUi.fontDD.selection || dialogUi.fontDD.selection.index === 0) return null;
            return dialogUi.fontDD.selection.text;
        }

        function getSelectedHeadingFontFamily(dialogUi) {
            if (dialogUi.disableFontRadio && dialogUi.disableFontRadio.value) return null;
            if (dialogUi.useSameFontRadio && dialogUi.useSameFontRadio.value) return getSelectedFontFamily(dialogUi);
            if (!dialogUi.headingFontDD.selection || dialogUi.headingFontDD.selection.index === 0) return null;
            return dialogUi.headingFontDD.selection.text;
        }

        function getHeadingFontStyleName(dialogUi, previewRow) {
            if (dialogUi.disableFontRadio && dialogUi.disableFontRadio.value) return null;
            return getFontStyleDropdownValue(previewRow.fontStyleDD);
        }

        function getBodyLeadingMultiplier(dialogUi) {
            return parsePositivePercentMultiplier(dialogUi.leadingBodyInput.text, DEFAULT_BODY_LEADING_PERCENT / 100);
        }

        function getHeadingLeadingMultiplier(dialogUi) {
            return parsePositivePercentMultiplier(dialogUi.leadingHeadingInput.text, DEFAULT_HEADING_LEADING_PERCENT / 100);
        }

        function getBodyKerningMethod(dialogUi) {
            return getSelectedDropdownOptionValue(dialogUi.bodyKerningDD, kerningOptions, "value", "和文等幅");
        }

        function getHeadingKerningMethod(dialogUi) {
            return getSelectedDropdownOptionValue(dialogUi.headingKerningDD, kerningOptions, "value", "メトリクス");
        }

        function getSpaceAfterPercent(dialogUi) {
            return parseNonNegativeNumber(dialogUi.spaceAfterInput.text, DEFAULT_SPACE_AFTER_PERCENT);
        }

        function getBaseSize(dialogUi) {
            return parsePositiveNumber(dialogUi.baseInput.text, null);
        }

        function formatLeadingValue(value, roundDigits) {
            if (typeof value !== "number" || isNaN(value)) return "—";
            return String(roundTo(value, roundDigits)) + unitSym;
        }

        function getLevelStyleNames(dialogUi) {
            var levelStyleNames = [];
            for (var levelRowIndex = 0; levelRowIndex < dialogUi.levelRows.length; levelRowIndex++) {
                levelStyleNames.push(getDropdownText(dialogUi.levelRows[levelRowIndex].styleDD));
            }
            return levelStyleNames;
        }

        function getFontStyleDropdownValue(dropdownList) {
            if (!dropdownList.selection) return null;
            if (dropdownList.selection.text === L("noFontChange")) return null;
            return dropdownList.selection.text;
        }

        function setRowEnabled(previewRow, enabled) {
            previewRow.lbl.enabled = enabled;
            previewRow.sizeText.enabled = enabled;
            previewRow.leadingText.enabled = enabled;
            previewRow.styleDD.enabled = enabled;
            previewRow.fontStyleDD.enabled = enabled;
        }

        function getAllPreviewRows(dialogUi) {
            var previewRows = [];
            for (var levelRowIndex = 0; levelRowIndex < dialogUi.levelRows.length; levelRowIndex++) {
                previewRows.push(dialogUi.levelRows[levelRowIndex]);
            }
            previewRows.push(dialogUi.baseRow);
            previewRows.push(dialogUi.captionRow);
            return previewRows;
        }

        function resetDropdownItems(dropdownList, items) {
            var currentText = getDropdownText(dropdownList);

            while (dropdownList.items.length > 0) {
                dropdownList.remove(dropdownList.items[0]);
            }

            for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
                dropdownList.add("item", items[itemIndex]);
            }

            if (!selectDropdownByText(dropdownList, currentText) && dropdownList.items.length > 0) {
                dropdownList.selection = 0;
            }
        }

        function selectAllPreviewFontStyleDropdowns(dialogUi, styleName) {
            var previewRows = getAllPreviewRows(dialogUi);
            for (var previewRowIndex = 0; previewRowIndex < previewRows.length; previewRowIndex++) {
                selectDropdownByText(previewRows[previewRowIndex].fontStyleDD, styleName);
            }
        }

        function syncPreviewFontStylesFromTextSettings(dialogUi) {
            var selectedFontStyleName = getDropdownText(dialogUi.fontStyleDD);
            if (!selectedFontStyleName || selectedFontStyleName === L("noFontChange")) return;
            selectAllPreviewFontStyleDropdowns(dialogUi, selectedFontStyleName);
        }

        function syncHeadingPreviewFontStylesFromTextSettings(dialogUi) {
            var sourceFontStyleDropdown = (dialogUi.useSameFontRadio && dialogUi.useSameFontRadio.value)
                ? dialogUi.fontStyleDD
                : dialogUi.headingFontStyleDD;
            var selectedFontStyleName = getDropdownText(sourceFontStyleDropdown);
            if (!selectedFontStyleName || selectedFontStyleName === L("noFontChange")) return;
            for (var levelRowIndex = 0; levelRowIndex < dialogUi.levelRows.length; levelRowIndex++) {
                selectDropdownByText(dialogUi.levelRows[levelRowIndex].fontStyleDD, selectedFontStyleName);
            }
        }

        function syncFontSelectionEnabled(dialogUi) {
            var enabled = !(dialogUi.disableFontRadio && dialogUi.disableFontRadio.value);
            var headingFontEnabled = enabled && !(dialogUi.useSameFontRadio && dialogUi.useSameFontRadio.value);
            dialogUi.fontDD.enabled = enabled;
            dialogUi.fontStyleDD.enabled = enabled;
            dialogUi.headingFontDD.enabled = headingFontEnabled;
            dialogUi.headingFontStyleDD.enabled = headingFontEnabled;
            if (dialogUi.useSameFontRadio) dialogUi.useSameFontRadio.enabled = true;
            if (dialogUi.separateFontRadio) dialogUi.separateFontRadio.enabled = true;
            if (dialogUi.disableFontRadio) dialogUi.disableFontRadio.enabled = true;
            if (!enabled) {
                selectDropdownByText(dialogUi.fontDD, L("noFontChange"));
                resetDropdownItems(dialogUi.fontStyleDD, [L("noFontChange")]);
                selectDropdownByText(dialogUi.headingFontDD, L("noFontChange"));
                resetDropdownItems(dialogUi.headingFontStyleDD, [L("noFontChange")]);
            }
            var previewRows = getAllPreviewRows(dialogUi);
            for (var previewRowIndex = 0; previewRowIndex < previewRows.length; previewRowIndex++) {
                previewRows[previewRowIndex].fontStyleDD.enabled = enabled;
                if (!enabled) {
                    resetDropdownItems(previewRows[previewRowIndex].fontStyleDD, [L("noFontChange")]);
                }
            }
        }

        function updateFontStyleDropdowns(dialogUi) {
            if (dialogUi.disableFontRadio && dialogUi.disableFontRadio.value) {
                syncFontSelectionEnabled(dialogUi);
                return;
            }
            var selectedFontFamily = getSelectedFontFamily(dialogUi);
            var fontStyleOptions = selectedFontFamily ? getFontStylesInFamily(selectedFontFamily) : [L("noFontChange")];
            if (fontStyleOptions.length === 0) fontStyleOptions = [L("noFontChange")];

            resetDropdownItems(dialogUi.fontStyleDD, fontStyleOptions);
            dialogUi.fontStyleDD.enabled = true;

            var selectedMasterFontStyleName = getDropdownText(dialogUi.fontStyleDD);
            var previewRows = getAllPreviewRows(dialogUi);
            for (var previewRowIndex = 0; previewRowIndex < previewRows.length; previewRowIndex++) {
                resetDropdownItems(previewRows[previewRowIndex].fontStyleDD, fontStyleOptions);
                previewRows[previewRowIndex].fontStyleDD.enabled = true;
            }
            if (selectedMasterFontStyleName && selectedMasterFontStyleName !== L("noFontChange")) {
                selectAllPreviewFontStyleDropdowns(dialogUi, selectedMasterFontStyleName);
            }
        }

        function updateHeadingFontStyleDropdowns(dialogUi) {
            if (dialogUi.disableFontRadio && dialogUi.disableFontRadio.value) {
                syncFontSelectionEnabled(dialogUi);
                return;
            }
            if (dialogUi.useSameFontRadio && dialogUi.useSameFontRadio.value) {
                syncFontSelectionEnabled(dialogUi);
                return;
            }
            var selectedFontFamily = getSelectedHeadingFontFamily(dialogUi);
            var fontStyleOptions = selectedFontFamily ? getFontStylesInFamily(selectedFontFamily) : [L("noFontChange")];
            if (fontStyleOptions.length === 0) fontStyleOptions = [L("noFontChange")];

            resetDropdownItems(dialogUi.headingFontStyleDD, fontStyleOptions);
            dialogUi.headingFontStyleDD.enabled = true;

            var selectedHeadingFontStyleName = getDropdownText(dialogUi.headingFontStyleDD);
            for (var levelRowIndex = 0; levelRowIndex < dialogUi.levelRows.length; levelRowIndex++) {
                resetDropdownItems(dialogUi.levelRows[levelRowIndex].fontStyleDD, fontStyleOptions);
                dialogUi.levelRows[levelRowIndex].fontStyleDD.enabled = true;
            }
            if (selectedHeadingFontStyleName && selectedHeadingFontStyleName !== L("noFontChange")) {
                syncHeadingPreviewFontStylesFromTextSettings(dialogUi);
            }
        }

        function getHeadingLevelFontStyleNames(dialogUi) {
            var levelFontStyleNames = [];
            for (var levelRowIndex = 0; levelRowIndex < dialogUi.levelRows.length; levelRowIndex++) {
                levelFontStyleNames.push(getHeadingFontStyleName(dialogUi, dialogUi.levelRows[levelRowIndex]));
            }
            return levelFontStyleNames;
        }

        function collectTypescaleSettings(dialogUi) {
            var baseSize = getBaseSize(dialogUi);
            return {
                baseSize: baseSize,
                ratio: getCurrentRatio(dialogUi),
                levelCount: getCurrentLevelCount(dialogUi),
                levelStyleNames: getLevelStyleNames(dialogUi),
                levelFontStyleNames: getHeadingLevelFontStyleNames(dialogUi),
                baseStyleName: getDropdownText(dialogUi.baseRow.styleDD),
                captionStyleName: getDropdownText(dialogUi.captionRow.styleDD),
                baseFontStyleName: getFontStyleDropdownValue(dialogUi.baseRow.fontStyleDD),
                captionFontStyleName: getFontStyleDropdownValue(dialogUi.captionRow.fontStyleDD),
                fontFamily: getSelectedFontFamily(dialogUi),

                headingFontFamily: getSelectedHeadingFontFamily(dialogUi),

                roundDigits: getCurrentRoundDigits(dialogUi),
                bodyLeading: getBodyLeadingMultiplier(dialogUi),
                headingLeading: getHeadingLeadingMultiplier(dialogUi),
                bodyKerningMethod: getBodyKerningMethod(dialogUi),
                headingKerningMethod: getHeadingKerningMethod(dialogUi),
                spaceAfterPercent: getSpaceAfterPercent(dialogUi)
            };
        }

        function updateTypescalePreview(dialogUi) {
            var baseSize = getBaseSize(dialogUi);
            var ratio = getCurrentRatio(dialogUi);
            var levelCount = getCurrentLevelCount(dialogUi);
            var roundDigits = getCurrentRoundDigits(dialogUi);

            if (baseSize === null) {
                for (var levelRowIndex = 0; levelRowIndex < dialogUi.levelRows.length; levelRowIndex++) {
                    dialogUi.levelRows[levelRowIndex].sizeText.text = "—";
                    dialogUi.levelRows[levelRowIndex].leadingText.text = "—";
                    setRowEnabled(dialogUi.levelRows[levelRowIndex], false);
                }
                dialogUi.baseRow.sizeText.text = "—";
                dialogUi.baseRow.leadingText.text = "—";
                dialogUi.captionRow.sizeText.text = "—";
                dialogUi.captionRow.leadingText.text = "—";
                return;
            }

            var computedSizes = computeSizes(baseSize, ratio, levelCount);
            for (var levelNumber = 1; levelNumber <= dialogUi.levelRows.length; levelNumber++) {
                if (levelNumber <= levelCount) {
                    var headingSize = computedSizes.headingSizes[levelNumber - 1];
                    dialogUi.levelRows[levelNumber - 1].sizeText.text = roundTo(headingSize, roundDigits) + " " + unitSym;
                    dialogUi.levelRows[levelNumber - 1].leadingText.text = formatLeadingValue(headingSize * getHeadingLeadingMultiplier(dialogUi), roundDigits);
                    setRowEnabled(dialogUi.levelRows[levelNumber - 1], true);
                } else {
                    dialogUi.levelRows[levelNumber - 1].sizeText.text = "—";
                    dialogUi.levelRows[levelNumber - 1].leadingText.text = "—";
                    setRowEnabled(dialogUi.levelRows[levelNumber - 1], false);
                }
            }
            dialogUi.baseRow.sizeText.text = roundTo(computedSizes.base, roundDigits) + " " + unitSym;
            dialogUi.baseRow.leadingText.text = formatLeadingValue(computedSizes.base * getBodyLeadingMultiplier(dialogUi), roundDigits);
            dialogUi.captionRow.sizeText.text = roundTo(computedSizes.caption, roundDigits) + " " + unitSym;
            dialogUi.captionRow.leadingText.text = formatLeadingValue(computedSizes.caption * getBodyLeadingMultiplier(dialogUi), roundDigits);

            if (dialogUi.previewCheck.value) {
                applyTypescaleSettings(targetDocument, collectTypescaleSettings(dialogUi), true, unit);
            }
        }

        function clearTextOverridesInSelection() {
            var selectionItems = app.selection;
            if (!selectionItems || selectionItems.length === 0) return;

            for (var selectionIndex = 0; selectionIndex < selectionItems.length; selectionIndex++) {
                var selectedItem = selectionItems[selectionIndex];
                try { selectedItem.clearOverrides(OverrideType.ALL); } catch (e1) { }
                try { selectedItem.texts[0].clearOverrides(OverrideType.ALL); } catch (e2) { }
                try { selectedItem.paragraphs.everyItem().clearOverrides(OverrideType.ALL); } catch (e3) { }
            }
        }

        function clearOverridesIfActive(dialogUi, forceClear) {
            // プレビュー時のみオーバーライドをクリア（破壊的操作回避）
            if (!forceClear && !dialogUi.previewCheck.value) return;
            clearTextOverridesInSelection();
            try { app.menuActions.itemByID(8489).invoke(); } catch (e) { }
            try { app.redraw(); } catch (e2) { }
        }

        function refreshPreviewAfterFontChange(dialogUi) {
            // プレビューを一度OFF→ONして再描画を強制
            var wasPreviewEnabled = dialogUi.previewCheck.value;
            if (wasPreviewEnabled) {
                dialogUi.previewCheck.value = false;
                clearOverridesIfActive(dialogUi, true);
                dialogUi.previewCheck.value = true;
            }
            updateTypescalePreview(dialogUi);
        }

        function setFontOptionMode(dialogUi, mode) {
            dialogUi.useSameFontRadio.value = (mode === "same");
            dialogUi.separateFontRadio.value = (mode === "separate");
            dialogUi.disableFontRadio.value = (mode === "disable");
        }

        function bindTypescaleDialogEvents(dialogUi) {
            changeValueByArrowKey(dialogUi.baseInput, function () {
                updateTypescalePreview(dialogUi);
            });
            changeValueByArrowKey(dialogUi.leadingBodyInput, function () { updateTypescalePreview(dialogUi); });
            changeValueByArrowKey(dialogUi.leadingHeadingInput, function () { updateTypescalePreview(dialogUi); });
            changeValueByArrowKey(dialogUi.spaceAfterInput, function () { updateTypescalePreview(dialogUi); });
            dialogUi.baseInput.onChanging = function () { updateTypescalePreview(dialogUi); };
            dialogUi.baseInput.onChange = function () {
                updateTypescalePreview(dialogUi);
            };
            dialogUi.ratioDD.onChange = function () { updateTypescalePreview(dialogUi); };
            for (var levelRadioIndex = 0; levelRadioIndex < dialogUi.levelRadios.length; levelRadioIndex++) {
                dialogUi.levelRadios[levelRadioIndex].onClick = function () { updateTypescalePreview(dialogUi); };
            }
            for (var levelRowIndex = 0; levelRowIndex < dialogUi.levelRows.length; levelRowIndex++) {
                dialogUi.levelRows[levelRowIndex].styleDD.onChange = function () { updateTypescalePreview(dialogUi); };
            }
            for (var levelFontStyleRowIndex = 0; levelFontStyleRowIndex < dialogUi.levelRows.length; levelFontStyleRowIndex++) {
                dialogUi.levelRows[levelFontStyleRowIndex].fontStyleDD.onChange = function () { updateTypescalePreview(dialogUi); };
            }
            dialogUi.baseRow.styleDD.onChange = function () { updateTypescalePreview(dialogUi); };
            dialogUi.captionRow.styleDD.onChange = function () { updateTypescalePreview(dialogUi); };
            dialogUi.baseRow.fontStyleDD.onChange = function () { updateTypescalePreview(dialogUi); };
            dialogUi.captionRow.fontStyleDD.onChange = function () { updateTypescalePreview(dialogUi); };
            dialogUi.disableFontRadio.onClick = function () {
                setFontOptionMode(dialogUi, "disable");
                syncFontSelectionEnabled(dialogUi);
                updateFontStyleDropdowns(dialogUi);
                updateHeadingFontStyleDropdowns(dialogUi);
                updateTypescalePreview(dialogUi);
            };
            dialogUi.useSameFontRadio.onClick = function () {
                setFontOptionMode(dialogUi, "same");
                syncFontSelectionEnabled(dialogUi);
                updateFontStyleDropdowns(dialogUi);
                updateHeadingFontStyleDropdowns(dialogUi);
                syncPreviewFontStylesFromTextSettings(dialogUi);
                updateTypescalePreview(dialogUi);
            };
            dialogUi.separateFontRadio.onClick = function () {
                setFontOptionMode(dialogUi, "separate");
                syncFontSelectionEnabled(dialogUi);
                updateFontStyleDropdowns(dialogUi);
                updateHeadingFontStyleDropdowns(dialogUi);
                updateTypescalePreview(dialogUi);
            };
            dialogUi.fontDD.onChange = function () {
                updateFontStyleDropdowns(dialogUi);
                if (dialogUi.useSameFontRadio && dialogUi.useSameFontRadio.value) updateHeadingFontStyleDropdowns(dialogUi);
                syncPreviewFontStylesFromTextSettings(dialogUi);
                refreshPreviewAfterFontChange(dialogUi);
            };
            dialogUi.fontStyleDD.onChange = function () {
                syncPreviewFontStylesFromTextSettings(dialogUi);
                if (dialogUi.useSameFontRadio && dialogUi.useSameFontRadio.value) syncHeadingPreviewFontStylesFromTextSettings(dialogUi);
                refreshPreviewAfterFontChange(dialogUi);
            };
            dialogUi.headingFontDD.onChange = function () {
                updateHeadingFontStyleDropdowns(dialogUi);
                syncHeadingPreviewFontStylesFromTextSettings(dialogUi);
                refreshPreviewAfterFontChange(dialogUi);
            };
            dialogUi.headingFontStyleDD.onChange = function () {
                syncHeadingPreviewFontStylesFromTextSettings(dialogUi);
                refreshPreviewAfterFontChange(dialogUi);
            };
            for (var roundRadioIndex = 0; roundRadioIndex < dialogUi.roundRadios.length; roundRadioIndex++) {
                dialogUi.roundRadios[roundRadioIndex].onClick = function () { updateTypescalePreview(dialogUi); };
            }
            dialogUi.leadingBodyInput.onChanging = function () { updateTypescalePreview(dialogUi); };
            dialogUi.leadingBodyInput.onChange = function () { updateTypescalePreview(dialogUi); };
            dialogUi.leadingHeadingInput.onChanging = function () { updateTypescalePreview(dialogUi); };
            dialogUi.leadingHeadingInput.onChange = function () { updateTypescalePreview(dialogUi); };
            dialogUi.bodyKerningDD.onChange = function () { updateTypescalePreview(dialogUi); };
            dialogUi.headingKerningDD.onChange = function () { updateTypescalePreview(dialogUi); };
            dialogUi.spaceAfterInput.onChanging = function () { updateTypescalePreview(dialogUi); };
            dialogUi.spaceAfterInput.onChange = function () { updateTypescalePreview(dialogUi); };
            dialogUi.previewCheck.onClick = function () { updateTypescalePreview(dialogUi); };
        }

        bindTypescaleDialogEvents(dialogUi);
        updateFontStyleDropdowns(dialogUi);
        updateHeadingFontStyleDropdowns(dialogUi);
        updateTypescalePreview(dialogUi);
        syncFontSelectionEnabled(dialogUi);

        var dialogResult = dialogUi.dialog.show();
        clearOverridesIfActive(dialogUi);

        if (dialogResult !== 1) {
            return null;
        }

        if (getBaseSize(dialogUi) === null) {
            alert(L("errInvalidBaseSize"));
            return null;
        }
        return collectTypescaleSettings(dialogUi);
    }

    function setParagraphStyleProps(targetDocument, styleName, size, font, leading, spaceAfter, kerningMethod, silent) {
        var style = findParagraphStyle(targetDocument, styleName);
        if (style === null) {
            if (!silent) alert(formatLabel("errMissingParagraphStyle", styleName));
            return false;
        }
        // 環境差異に対応するため、複数パターンでフォントを設定（fullName → object → name）
        if (font) {
            var assigned = false;
            var fontFullName = getFontFullName(font.fontFamily, font.fontStyleName);
            if (fontFullName) {
                try {
                    style.appliedFont = fontFullName;
                    assigned = true;
                } catch (e1) { }
            }
            if (!assigned) {
                try {
                    style.appliedFont = font;
                    assigned = true;
                } catch (e2) { }
            }
            if (!assigned && font.name) {
                try {
                    style.appliedFont = font.name;
                    assigned = true;
                } catch (e3) { }
            }
            if (assigned && font.fontStyleName) {
                try { style.fontStyle = font.fontStyleName; } catch (e4) { }
            }
        }
        style.pointSize = size;
        if (typeof leading === "number" && leading > 0) {
            style.leading = leading;
        }
        if (typeof spaceAfter === "number" && spaceAfter >= 0) {
            style.spaceAfter = spaceAfter;
        }
        // フォントによっては設定不可のため、安全に無視
        if (kerningMethod) {
            try { style.kerningMethod = kerningMethod; } catch (ke) { }
        }
        try {
            if (spaceAfter !== null) {
                // 見出しは左揃え / Headings are left aligned
                style.justification = Justification.LEFT_ALIGN;
            } else {
                // 本文・キャプションは均等配置（最終行左） / Body and captions are left-justified
                style.justification = Justification.LEFT_JUSTIFIED;
            }
        } catch (e) { }
        return true;
    }

    function findParagraphStyle(targetDocument, styleName) {
        var styles = targetDocument.allParagraphStyles;
        for (var styleIndex = 0; styleIndex < styles.length; styleIndex++) {
            if (styles[styleIndex].name === styleName) {
                return styles[styleIndex];
            }
        }
        return null;
    }

})();