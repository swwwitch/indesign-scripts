#target indesign

/*
 * InDesignTypescale.jsx
 *
 * 基準サイズとスケール倍率からタイプスケールを組み立て、本文・見出し・キャプションの段落スタイルへ一括適用します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "InDesignTypescale";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-05-05";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/InDesignTypescale.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/InDesignTypescale.md

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

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 基準サイズとスケール / Base size and scale */
var DEFAULT_BASE_SIZE_PT = 9;      /* 基準サイズ（本文, pt）/ base body size (pt) */
var DEFAULT_BASE_SIZE_Q  = 13;     /* 基準サイズ（本文, Q）/ base body size (Q) */
var DEFAULT_RATIO        = 1.309;  /* スケール倍率 / scale ratio */
var DEFAULT_LEVEL_COUNT  = 4;      /* 見出しレベル数 / number of heading levels */

/* 行送りとアキの既定値（%）/ Default leading and spacing (%) */
var DEFAULT_BODY_LEADING_PERCENT    = 160;  /* 本文の行送り / body leading */
var DEFAULT_HEADING_LEADING_PERCENT = 115;  /* 見出しの行送り / heading leading */
var DEFAULT_SPACE_AFTER_PERCENT     = 10;   /* 段落後のアキ / space after */

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
        title: { ja: "タイプスケールで一括設定", en: "Type Scale Settings" }
    },
    panel: {
        textSettings:  { ja: "テキスト設定", en: "Text Settings" },
        fontSettings:  { ja: "フォント指定オプション", en: "Font Options" },
        bodyText:      { ja: "本文", en: "Body" },
        headingText:   { ja: "見出し", en: "Headings" },
        scaleSettings: { ja: "スケール設定", en: "Scale Settings" },
        preview:       { ja: "段落スタイルとサイズプレビュー", en: "Paragraph Styles & Size Preview" }
    },
    field: {
        baseSizeBody:        { ja: "基準サイズ", en: "Base Size (Body)" },
        font:                { ja: "フォント", en: "Font" },
        fontStyle:           { ja: "スタイル", en: "Style" },
        bodyLeadingRatio:    { ja: "行送り", en: "Leading Ratio (Body)" },
        headingLeadingRatio: { ja: "行送り", en: "Leading Ratio (Headings)" },
        spaceAfterRatio:     { ja: "段落後のアキ", en: "Space After" },
        kerningMethod:       { ja: "カーニング", en: "Kerning Method" },
        scaleRatio:          { ja: "倍率", en: "Scale Ratio" },
        headingLevelCount:   { ja: "見出しレベル数", en: "Heading Levels" },
        sizeRounding:        { ja: "サイズの丸め", en: "Size Rounding" }
    },
    radio: {
        disableFontSelection:          { ja: "フォントを変更しない", en: "Do not change fonts" },
        useSameFontForBodyAndHeading:  { ja: "本文と見出しで共通", en: "Use same font for body and headings" },
        separateFontForBodyAndHeading: { ja: "本文と見出しで別々に指定", en: "Specify separately" }
    },
    kerning: {
        japaneseMono: { ja: "和文等幅", en: "Japanese Mono" },
        metrics:      { ja: "メトリクス", en: "Metrics" },
        optical:      { ja: "オプティカル", en: "Optical" }
    },
    rounding: {
        integer:       { ja: "整数", en: "Integer" },
        firstDecimal:  { ja: "小数点第1位", en: "1 decimal place" },
        secondDecimal: { ja: "小数点第2位", en: "2 decimal places" }
    },
    header: {
        level:          { ja: "レベル", en: "Level" },
        fontSize:       { ja: "フォントサイズ", en: "Font Size" },
        leading:        { ja: "行送り", en: "Leading" },
        paragraphStyle: { ja: "段落スタイル", en: "Paragraph Style" },
        fontStyle:      { ja: "スタイル", en: "Style" }
    },
    row: {
        levelPrefix: { ja: "レベル", en: "Level " },
        baseBody:    { ja: "基準（本文）", en: "Base (Body)" },
        caption:     { ja: "キャプション", en: "Caption" }
    },
    option: {
        noFontChange: { ja: "（変更しない）", en: "(No change)" }
    },
    checkbox: {
        livePreview: { ja: "ライブプレビュー", en: "Live Preview" }
    },
    button: {
        ok:     { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    progress: {
        title:        { ja: "処理中", en: "Processing" },
        loadingFonts: { ja: "フォント情報を読み込んでいます…", en: "Loading font information..." }
    },
    undo: {
        applyTypescale: { ja: "タイプスケールを適用", en: "Apply Type Scale" }
    },
    error: {
        noDocument:           { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        invalidBaseSize:      { ja: "基準サイズ（本文）には正の数値を入力してください。", en: "Enter a positive number for Base Size (Body)." },
        missingParagraphStyle:{ ja: "段落スタイル「%1」が見つかりません。", en: "Paragraph style \"%1\" was not found." }
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

/**
 * コロン付きラベルを取得する（日本語は全角コロン、英語は半角コロン）
 * @param {string} labelKey 例: "field.font"
 * @returns {string} コロンを付与したラベル文字列
 */
function getLabelWithColon(labelKey) {
    return getLabel(labelKey) + (currentLang === "ja" ? "：" : ":");
}

/**
 * ラベル内の %1 を値で置き換える
 * @param {string} labelKey 置換対象のラベルキー
 * @param {*} value 埋め込む値
 * @returns {string} 置換後の文字列
 */
function formatLabel(labelKey, value) {
    return getLabel(labelKey).replace("%1", value);
}

/**
 * 進捗表示用のパレットを作る
 * @param {string} message 最初に表示するメッセージ
 * @returns {Window|null} パレット。作成できない場合は null
 */
function createProgressPalette(message) {
    var palette = new Window("palette", getLabel("progress.title"));
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

/**
 * 進捗パレットを閉じる
 * @param {Window} palette 対象のパレット
 * @returns {void}
 */
function closeProgressPalette(palette) {
    if (!palette) return;
    try { palette.close(); } catch (e) { }
}

(function () {

    if (app.documents.length === 0) {
        alert(getLabel("error.noDocument"));
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
        /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
        app.doScript(function () {
            applyTypescaleSettings(targetDocument, typescaleSettings, false, unit);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.applyTypescale"));
    }

    /**
     * 基準サイズとスケール倍率から各レベルのサイズを計算する
     * @param {number} baseSize 基準サイズ
     * @param {number} ratio スケール倍率
     * @param {number} levelCount 見出しレベル数
     * @param {number} roundDigits 丸め桁数
     * @returns {Array<number>} 各レベルのサイズ
     */
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

    /**
     * 確定したタイプスケール設定を段落スタイルへ適用する
     * @param {object} settings ダイアログで確定した設定
     * @returns {void}
     */
    function applyTypescaleSettings(targetDocument, typescaleSettings, silent, unit) {
        var computedSizes = computeSizes(typescaleSettings.baseSize, typescaleSettings.ratio, typescaleSettings.levelCount);
        // 無効値時はデフォルトにフォールバック
        var spaceAfterPercent = (typeof typescaleSettings.spaceAfterPercent === "number" && typescaleSettings.spaceAfterPercent >= 0)
            ? typescaleSettings.spaceAfterPercent
            : DEFAULT_SPACE_AFTER_PERCENT;
        /**
         * 1 つの段落スタイルへサイズ・行送り・アキなどを適用する
         * @param {ParagraphStyle} paragraphStyle 対象の段落スタイル
         * @param {object} styleSettings 適用する設定
         * @returns {void}
         */
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

    /**
     * 環境設定のテキストサイズ単位を取得する
     * @returns {MeasurementUnits} テキストサイズの単位
     */
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

    /**
     * 単位の列挙値を数値として取得する
     * @param {MeasurementUnits} unit 対象の単位
     * @returns {number} 単位を表す数値
     */
    function getMeasurementUnitValue(unitName) {
        try {
            return MeasurementUnits[unitName];
        } catch (e) { }
        return null;
    }

    /**
     * 単位の表示記号を返す
     * @returns {string} 単位の記号
     */
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

    /**
     * 現在の単位の数値をポイントへ換算する
     * @param {number} value 現在の単位での数値
     * @returns {number} ポイント値
     */
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

    /**
     * 指定桁数で丸める
     * @param {number} value 対象の数値
     * @param {number} digits 小数点以下の桁数
     * @returns {number} 丸めた数値
     */
    function roundTo(num, places) {
        var factor = Math.pow(10, places);
        return Math.round(num * factor) / factor;
    }

    /**
     * 配列に値が含まれるかを判定する
     * @param {Array} list 対象の配列
     * @param {*} value 探す値
     * @returns {boolean} 含まれていれば true
     */
    function arrayContains(array, value) {
        for (var itemIndex = 0; itemIndex < array.length; itemIndex++) {
            if (array[itemIndex] === value) return true;
        }
        return false;
    }

    /**
     * ドキュメント内の段落スタイル名を集める
     * @param {Document} doc 対象ドキュメント
     * @returns {Array<string>} 段落スタイル名の配列
     */
    function getParagraphStyleNames(targetDocument) {
        var names = [];
        var styles = targetDocument.allParagraphStyles;
        for (var styleIndex = 0; styleIndex < styles.length; styleIndex++) {
            names.push(styles[styleIndex].name);
        }
        return names;
    }

    var _fontInfo = null;

    /**
     * 利用できるフォントファミリーとスタイルの一覧を作る
     * @returns {object} フォント情報
     */
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

    /**
     * フォントファミリー名の一覧を取得する
     * @returns {Array<string>} ファミリー名の配列
     */
    function getFontFamilyNames() {
        return getFontInfo().families.slice(0);
    }

    /**
     * ファミリー内で優先して選ぶスタイル名を求める
     * @param {string} familyName フォントファミリー名
     * @returns {string} スタイル名
     */
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

    /**
     * ファミリーとスタイルからフォントのフルネームを作る
     * @param {string} familyName フォントファミリー名
     * @param {string} styleName フォントスタイル名
     * @returns {string} フォントのフルネーム
     */
    function getFontFullName(familyName, styleName) {
        if (!familyName || !styleName) return null;
        return familyName + "\t" + styleName;
    }

    /**
     * ファミリー内のフォントを探す
     * @param {string} familyName フォントファミリー名
     * @returns {Font|null} フォント。見つからない場合は null
     */
    function findFontInFamily(familyName) {
        if (!familyName) return null;
        var styleNames = getFontStylesInFamily(familyName);
        var preferredStyleName = getPreferredFontStyleName(styleNames);
        if (!preferredStyleName) return null;
        return findFontByFamilyAndStyle(familyName, preferredStyleName);
    }

    /**
     * ファミリー内のスタイル名を集める
     * @param {string} familyName フォントファミリー名
     * @returns {Array<string>} スタイル名の配列
     */
    function getFontStylesInFamily(familyName) {
        if (!familyName) return [];
        var fontInfo = getFontInfo();
        if (!fontInfo.fontMap[familyName]) return [];
        return fontInfo.fontMap[familyName].slice(0);
    }

    /**
     * ファミリーとスタイルからフォントを探す
     * @param {string} familyName フォントファミリー名
     * @param {string} styleName フォントスタイル名
     * @returns {Font|null} フォント。見つからない場合は null
     */
    function findFontByFamilyAndStyle(familyName, styleName) {
        var fontFullName = getFontFullName(familyName, styleName);
        if (!fontFullName) return null;

        var fontInfo = getFontInfo();
        if (fontInfo.fontObjectMap[fontFullName]) {
            return fontInfo.fontObjectMap[fontFullName];
        }
        return null;
    }

    /**
     * 表示名でドロップダウンの項目を選択する
     * @param {DropDownList} dropdown 対象のドロップダウン
     * @param {string} text 選択したい項目名
     * @returns {boolean} 選択できたら true
     */
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

    /**
     * ドロップダウンで選択中の表示名を取得する
     * @param {DropDownList} dropdown 対象のドロップダウン
     * @returns {string} 選択中の表示名
     */
    function getDropdownText(dropdownList) {
        return dropdownList.selection ? dropdownList.selection.text : null;
    }

    /**
     * タイプスケール設定ダイアログを表示する
     * @param {Document} doc 対象ドキュメント
     * @returns {object|null} 設定内容。キャンセル時は null
     */
    function showTypescaleDialog(targetDocument, defaultBase, defaultRatio, defaultLevelCount, unit) {
        var unitSym = unitSymbol(unit);
        var styleNames = getParagraphStyleNames(targetDocument);
        var loadingPalette = createProgressPalette(getLabel("progress.loadingFonts"));
        var fontFamilies = getFontFamilyNames();
        closeProgressPalette(loadingPalette);
        var fontOptions = [getLabel("option.noFontChange")].concat(fontFamilies);
        var roundOptions = [
            { label: getLabel("rounding.integer"), digits: 0 },
            { label: getLabel("rounding.firstDecimal"), digits: 1 },
            { label: getLabel("rounding.secondDecimal"), digits: 2 }
        ];
        var defaultRoundDigits = 0;
        var kerningOptions = [
            { label: getLabel("kerning.japaneseMono"), value: "和文等幅" },
            { label: getLabel("kerning.metrics"), value: "メトリクス" },
            { label: getLabel("kerning.optical"), value: "オプティカル" }
        ];
        var PANEL_MARGINS = [15, 20, 15, 10];


        /**
         * 幅を固定したラベルを追加する
         * @param {object} parent 追加先のコンテナ
         * @param {string} text 表示する文字列
         * @param {number} width ラベルの幅（px）
         * @returns {StaticText} 追加したラベル
         */
        function addFixedWidthLabel(parent, text, width) {
            var label = parent.add("statictext", undefined, text);
            label.preferredSize.width = width;
            return label;
        }

        /**
         * ラベル付きの行グループを追加する
         * @param {object} parent 追加先のコンテナ
         * @param {string} labelText ラベルの文字列
         * @param {number} labelWidth ラベルの幅（px）
         * @returns {Group} 追加したグループ
         */
        function addLabeledGroup(panel, labelText, labelWidth) {
            var group = panel.add("group");
            group.orientation = "row";
            group.alignChildren = ["left", "center"];
            addFixedWidthLabel(group, labelText, labelWidth);
            return group;
        }

        /**
         * 入力欄に上下キーでの増減操作を追加する
         * @param {EditText} editText 対象の入力欄
         * @param {function} onAfterChange 値の変更後に呼ぶ処理
         * @returns {void}
         */
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

        /**
         * カーニング方式の表示名を取得する
         * @returns {Array<string>} 表示名の配列
         */
        function getKerningOptionLabels() {
            var labels = [];
            for (var kerningOptionIndex = 0; kerningOptionIndex < kerningOptions.length; kerningOptionIndex++) {
                labels.push(kerningOptions[kerningOptionIndex].label);
            }
            return labels;
        }

        /**
         * カーニング方式の値でドロップダウンを選択する
         * @param {DropDownList} dropdown 対象のドロップダウン
         * @param {string} value 選択したい値
         * @returns {void}
         */
        function selectKerningDropdownByValue(dropdownList, value) {
            for (var kerningOptionIndex = 0; kerningOptionIndex < kerningOptions.length; kerningOptionIndex++) {
                if (kerningOptions[kerningOptionIndex].value === value) {
                    dropdownList.selection = kerningOptionIndex;
                    return;
                }
            }
            dropdownList.selection = 0;
        }


        /**
         * 本文設定パネルを組み立てる
         * @param {object} parent 追加先のコンテナ
         * @param {Array<string>} fontOptions フォントの選択肢
         * @param {number} labelWidth ラベルの幅（px）
         * @returns {object} パネル内のコントロール
         */
        function createBodyTextPanel(parent, labelWidth) {
            var bodyPanel = parent.add("panel", undefined, getLabel("panel.bodyText"));
            setupPanel(bodyPanel, 6);
            bodyPanel.alignment = ["fill", "top"];

            var fontGrp = addLabeledGroup(bodyPanel, getLabel("field.font"), labelWidth);
            var fontDD = fontGrp.add("dropdownlist", undefined, fontOptions);
            fontDD.preferredSize.width = 180;
            fontDD.selection = 0;

            var fontStyleGrp = addLabeledGroup(bodyPanel, getLabel("field.fontStyle"), labelWidth);
            var fontStyleDD = fontStyleGrp.add("dropdownlist", undefined, [getLabel("option.noFontChange")]);
            fontStyleDD.preferredSize.width = 130;
            fontStyleDD.selection = 0;

            var leadingBodyGrp = addLabeledGroup(bodyPanel, getLabel("field.bodyLeadingRatio"), labelWidth);
            var leadingBodyInput = leadingBodyGrp.add("edittext", undefined, String(DEFAULT_BODY_LEADING_PERCENT));
            leadingBodyInput.characters = 4;
            leadingBodyGrp.add("statictext", undefined, "%");

            var bodyKerningGrp = addLabeledGroup(bodyPanel, getLabel("field.kerningMethod"), labelWidth);
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

        /**
         * 見出し設定パネルを組み立てる
         * @param {object} parent 追加先のコンテナ
         * @param {Array<string>} fontOptions フォントの選択肢
         * @param {number} labelWidth ラベルの幅（px）
         * @returns {object} パネル内のコントロール
         */
        function createHeadingTextPanel(parent, labelWidth) {
            var headingPanel = parent.add("panel", undefined, getLabel("panel.headingText"));
            setupPanel(headingPanel, 4);
            headingPanel.alignment = ["fill", "top"];

            // Font controls (like body panel)
            var headingFontGrp = addLabeledGroup(headingPanel, getLabel("field.font"), labelWidth);
            var headingFontDD = headingFontGrp.add("dropdownlist", undefined, fontOptions);
            headingFontDD.preferredSize.width = 180;
            headingFontDD.selection = 0;

            var headingFontStyleGrp = addLabeledGroup(headingPanel, getLabel("field.fontStyle"), labelWidth);
            var headingFontStyleDD = headingFontStyleGrp.add("dropdownlist", undefined, [getLabel("option.noFontChange")]);
            headingFontStyleDD.preferredSize.width = 130;
            headingFontStyleDD.selection = 0;

            var leadingHeadingGrp = addLabeledGroup(headingPanel, getLabel("field.headingLeadingRatio"), labelWidth);
            var leadingHeadingInput = leadingHeadingGrp.add("edittext", undefined, String(DEFAULT_HEADING_LEADING_PERCENT));
            leadingHeadingInput.characters = 4;
            leadingHeadingGrp.add("statictext", undefined, "%");

            var headingKerningGrp = addLabeledGroup(headingPanel, getLabel("field.kerningMethod"), labelWidth);
            var headingKerningDD = headingKerningGrp.add("dropdownlist", undefined, getKerningOptionLabels());
            headingKerningDD.preferredSize.width = 110;
            selectKerningDropdownByValue(headingKerningDD, "メトリクス");

            var spaceAfterGrp = addLabeledGroup(headingPanel, getLabel("field.spaceAfterRatio"), labelWidth);
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

        /**
         * 本文・見出しをまとめたテキスト設定パネルを組み立てる
         * @param {Window} dialog 対象のダイアログ
         * @param {Array<string>} fontOptions フォントの選択肢
         * @returns {object} パネル内のコントロール
         */
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

        /**
         * フォント指定オプションのパネルを組み立てる
         * @param {Window} dialog 対象のダイアログ
         * @returns {object} パネル内のコントロール
         */
        function createFontSettingsPanel(dialog) {
            var fontPanel = dialog.add("panel", undefined, getLabel("panel.fontSettings"));
            fontPanel.alignment = ["fill", "top"];
            setupPanel(fontPanel, 6);

            var modeGroup = fontPanel.add("group");
            modeGroup.orientation = "column";
            modeGroup.alignChildren = ["left", "center"];

            var useSameFontRadio = modeGroup.add("radiobutton", undefined, getLabel("radio.useSameFontForBodyAndHeading"));
            var separateFontRadio = modeGroup.add("radiobutton", undefined, getLabel("radio.separateFontForBodyAndHeading"));
            var disableFontRadio = modeGroup.add("radiobutton", undefined, getLabel("radio.disableFontSelection"));

            useSameFontRadio.value = true;
            separateFontRadio.value = false;
            disableFontRadio.value = false;

            return {
                disableFontRadio: disableFontRadio,
                useSameFontRadio: useSameFontRadio,
                separateFontRadio: separateFontRadio
            };
        }

        /**
         * スケール設定パネルを組み立てる
         * @param {Window} dialog 対象のダイアログ
         * @returns {object} パネル内のコントロール
         */
        function createScaleSettingsPanel(dialog) {
            var optionsPanel = dialog.add("panel", undefined, getLabel("panel.scaleSettings"));
            optionsPanel.alignment = ["fill", "top"];
            setupPanel(optionsPanel, 6);

            var OPTIONS_LABEL_WIDTH = 110;

            var baseGrp = addLabeledGroup(optionsPanel, getLabel("field.baseSizeBody"), OPTIONS_LABEL_WIDTH);
            var baseInput = baseGrp.add("edittext", undefined, String(defaultBase));
            baseInput.characters = 4;
            baseGrp.add("statictext", undefined, unitSym);

            var ratioGrp = addLabeledGroup(optionsPanel, getLabel("field.scaleRatio"), OPTIONS_LABEL_WIDTH);
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

            var levelGrp = addLabeledGroup(optionsPanel, getLabel("field.headingLevelCount"), OPTIONS_LABEL_WIDTH);
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

            var roundGrp = addLabeledGroup(optionsPanel, getLabel("field.sizeRounding"), OPTIONS_LABEL_WIDTH);
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

        /**
         * プレビューパネルを組み立てる
         * @param {Window} dialog 対象のダイアログ
         * @param {Array<string>} fontOptions フォントの選択肢
         * @returns {object} パネル内のコントロール
         */
        function createPreviewPanel(dialog) {
            var previewPanel = dialog.add("panel", undefined, getLabel("panel.preview"));
            setupPanel(previewPanel, 2);

            var PREVIEW_LABEL_WIDTH = 100;
            var PREVIEW_SIZE_WIDTH = 110;
            var PREVIEW_LEADING_WIDTH = 70;

            var headerRow = previewPanel.add("group");
            headerRow.orientation = "row";
            headerRow.alignChildren = "left";
            headerRow.add("statictext", undefined, getLabel("header.level")).preferredSize.width = PREVIEW_LABEL_WIDTH;
            headerRow.add("statictext", undefined, getLabel("header.fontSize")).preferredSize.width = PREVIEW_SIZE_WIDTH;
            headerRow.add("statictext", undefined, getLabel("header.leading")).preferredSize.width = PREVIEW_LEADING_WIDTH;
            headerRow.add("statictext", undefined, getLabel("header.paragraphStyle")).characters = 14;
            var fontStyleHeader = headerRow.add("statictext", undefined, getLabel("header.fontStyle"));
            fontStyleHeader.characters = 12;
            fontStyleHeader.enabled = true;

            var previewHeaderSpacer = previewPanel.add("group");
            previewHeaderSpacer.preferredSize.height = 4;

            /**
             * プレビューの 1 行を組み立てる
             * @param {object} parent 追加先のコンテナ
             * @param {string} rowLabel 行のラベル
             * @param {Array<string>} fontOptions フォントの選択肢
             * @returns {object} 行のコントロール
             */
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

                var fontStyleDD = row.add("dropdownlist", undefined, [getLabel("option.noFontChange")]);
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
                levelRows.push(createPreviewRow(previewPanel, getLabel("row.levelPrefix") + levelNumber, "h" + levelNumber));
            }

            return {
                levelRows: levelRows,
                baseRow: createPreviewRow(previewPanel, getLabel("row.baseBody"), "p"),
                captionRow: createPreviewRow(previewPanel, getLabel("row.caption"), "p.caption")
            };
        }

        /**
         * ダイアログ下部のボタン行を組み立てる
         * @param {Window} dialog 対象のダイアログ
         * @returns {object} 生成したボタン
         */
        function createButtonRow(dialog) {
            var bottomRow = dialog.add("group");
            bottomRow.margins = [0, 10, 0, 0];
            bottomRow.orientation = "row";
            bottomRow.alignment = "fill";
            bottomRow.alignChildren = ["fill", "center"];

            var leftButtonColumn = bottomRow.add("group");
            leftButtonColumn.orientation = "row";
            leftButtonColumn.alignChildren = ["left", "center"];

            var previewCheck = leftButtonColumn.add("checkbox", undefined, getLabel("checkbox.livePreview"));
            previewCheck.value = true;

            var centerButtonColumn = bottomRow.add("group");
            centerButtonColumn.alignment = ["fill", "fill"];
            centerButtonColumn.minimumSize.width = 0;

            var rightButtonColumn = bottomRow.add("group");
            rightButtonColumn.orientation = "row";
            rightButtonColumn.alignChildren = ["right", "center"];
            rightButtonColumn.alignment = ["right", "center"];

            rightButtonColumn.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
            rightButtonColumn.add("button", undefined, getLabel("button.ok"), { name: "ok" });

            return { previewCheck: previewCheck };
        }

        /**
         * タイプスケールダイアログ全体を組み立てる
         * @param {Document} doc 対象ドキュメント
         * @param {Array<string>} fontOptions フォントの選択肢
         * @returns {object} ダイアログとコントロール
         */
        function createTypescaleDialog() {
            var dlg = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
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

        /**
         * 正の数値として解釈する
         * @param {string} text 入力文字列
         * @param {number} fallback 解釈できない場合の値
         * @returns {number} 数値
         */
        function parsePositiveNumber(text, fallbackValue) {
            var value = parseFloat(text);
            return (isNaN(value) || value <= 0) ? fallbackValue : value;
        }

        /**
         * 0 以上の数値として解釈する
         * @param {string} text 入力文字列
         * @param {number} fallback 解釈できない場合の値
         * @returns {number} 数値
         */
        function parseNonNegativeNumber(text, fallbackValue) {
            var value = parseFloat(text);
            return (isNaN(value) || value < 0) ? fallbackValue : value;
        }

        /**
         * パーセント入力を倍率として解釈する
         * @param {string} text 入力文字列
         * @param {number} fallback 解釈できない場合の倍率
         * @returns {number} 倍率
         */
        function parsePositivePercentMultiplier(text, fallbackValue) {
            var value = parseFloat(text);
            return (isNaN(value) || value <= 0) ? fallbackValue : value / 100;
        }

        /**
         * ラジオボタン群で選択中の値を取得する
         * @param {Array<RadioButton>} radios ラジオボタンの配列
         * @param {Array<string>} values 対応する値
         * @returns {string|null} 選択中の値
         */
        function getSelectedRadioValue(radioButtons, options, valueKey, fallbackValue) {
            for (var radioIndex = 0; radioIndex < radioButtons.length; radioIndex++) {
                if (radioButtons[radioIndex].value) {
                    return valueKey ? options[radioIndex][valueKey] : options[radioIndex];
                }
            }
            return fallbackValue;
        }

        /**
         * ドロップダウンで選択中の値を取得する
         * @param {DropDownList} dropdown 対象のドロップダウン
         * @param {Array<string>} values 対応する値
         * @returns {string|null} 選択中の値
         */
        function getSelectedDropdownOptionValue(dropdownList, options, valueKey, fallbackValue) {
            if (!dropdownList.selection) return fallbackValue;
            return valueKey ? options[dropdownList.selection.index][valueKey] : options[dropdownList.selection.index];
        }

        /**
         * 現在のスケール倍率を取得する
         * @returns {number} スケール倍率
         */
        function getCurrentRatio(dialogUi) {
            if (!dialogUi.ratioDD.selection) return defaultRatio;
            return ratioOptions[dialogUi.ratioDD.selection.index].value;
        }

        /**
         * 現在の見出しレベル数を取得する
         * @returns {number} 見出しレベル数
         */
        function getCurrentLevelCount(dialogUi) {
            return getSelectedRadioValue(dialogUi.levelRadios, levelOptions, null, defaultLevelCount);
        }

        /**
         * 現在のサイズ丸め桁数を取得する
         * @returns {number} 丸め桁数
         */
        function getCurrentRoundDigits(dialogUi) {
            return getSelectedRadioValue(dialogUi.roundRadios, roundOptions, "digits", defaultRoundDigits);
        }

        /**
         * 本文で選択中のフォントファミリーを取得する
         * @returns {string} フォントファミリー名
         */
        function getSelectedFontFamily(dialogUi) {
            if (dialogUi.disableFontRadio && dialogUi.disableFontRadio.value) return null;
            if (!dialogUi.fontDD.selection || dialogUi.fontDD.selection.index === 0) return null;
            return dialogUi.fontDD.selection.text;
        }

        /**
         * 見出しで選択中のフォントファミリーを取得する
         * @returns {string} フォントファミリー名
         */
        function getSelectedHeadingFontFamily(dialogUi) {
            if (dialogUi.disableFontRadio && dialogUi.disableFontRadio.value) return null;
            if (dialogUi.useSameFontRadio && dialogUi.useSameFontRadio.value) return getSelectedFontFamily(dialogUi);
            if (!dialogUi.headingFontDD.selection || dialogUi.headingFontDD.selection.index === 0) return null;
            return dialogUi.headingFontDD.selection.text;
        }

        /**
         * 見出しで選択中のフォントスタイル名を取得する
         * @returns {string} フォントスタイル名
         */
        function getHeadingFontStyleName(dialogUi, previewRow) {
            if (dialogUi.disableFontRadio && dialogUi.disableFontRadio.value) return null;
            return getFontStyleDropdownValue(previewRow.fontStyleDD);
        }

        /**
         * 本文の行送り倍率を取得する
         * @returns {number} 行送り倍率
         */
        function getBodyLeadingMultiplier(dialogUi) {
            return parsePositivePercentMultiplier(dialogUi.leadingBodyInput.text, DEFAULT_BODY_LEADING_PERCENT / 100);
        }

        /**
         * 見出しの行送り倍率を取得する
         * @returns {number} 行送り倍率
         */
        function getHeadingLeadingMultiplier(dialogUi) {
            return parsePositivePercentMultiplier(dialogUi.leadingHeadingInput.text, DEFAULT_HEADING_LEADING_PERCENT / 100);
        }

        /**
         * 本文のカーニング方式を取得する
         * @returns {string} カーニング方式
         */
        function getBodyKerningMethod(dialogUi) {
            return getSelectedDropdownOptionValue(dialogUi.bodyKerningDD, kerningOptions, "value", "和文等幅");
        }

        /**
         * 見出しのカーニング方式を取得する
         * @returns {string} カーニング方式
         */
        function getHeadingKerningMethod(dialogUi) {
            return getSelectedDropdownOptionValue(dialogUi.headingKerningDD, kerningOptions, "value", "メトリクス");
        }

        /**
         * 段落後アキ（%）を取得する
         * @returns {number} 段落後のアキ（%）
         */
        function getSpaceAfterPercent(dialogUi) {
            return parseNonNegativeNumber(dialogUi.spaceAfterInput.text, DEFAULT_SPACE_AFTER_PERCENT);
        }

        /**
         * 入力されている基準サイズを取得する
         * @returns {number} 基準サイズ
         */
        function getBaseSize(dialogUi) {
            return parsePositiveNumber(dialogUi.baseInput.text, null);
        }

        /**
         * 行送りの表示用文字列を作る
         * @param {number} value 行送りの値
         * @returns {string} 表示する文字列
         */
        function formatLeadingValue(value, roundDigits) {
            if (typeof value !== "number" || isNaN(value)) return "—";
            return String(roundTo(value, roundDigits)) + unitSym;
        }

        /**
         * 見出しレベルに対応する段落スタイル名を返す
         * @param {number} levelCount 見出しレベル数
         * @returns {Array<string>} 段落スタイル名の配列
         */
        function getLevelStyleNames(dialogUi) {
            var levelStyleNames = [];
            for (var levelRowIndex = 0; levelRowIndex < dialogUi.levelRows.length; levelRowIndex++) {
                levelStyleNames.push(getDropdownText(dialogUi.levelRows[levelRowIndex].styleDD));
            }
            return levelStyleNames;
        }

        /**
         * フォントスタイルのドロップダウンから値を取得する
         * @param {DropDownList} dropdown 対象のドロップダウン
         * @returns {string} フォントスタイル名
         */
        function getFontStyleDropdownValue(dropdownList) {
            if (!dropdownList.selection) return null;
            if (dropdownList.selection.text === getLabel("option.noFontChange")) return null;
            return dropdownList.selection.text;
        }

        /**
         * プレビュー行の有効／無効を切り替える
         * @param {object} row 対象の行
         * @param {boolean} enabled 有効にするなら true
         * @returns {void}
         */
        function setRowEnabled(previewRow, enabled) {
            previewRow.lbl.enabled = enabled;
            previewRow.sizeText.enabled = enabled;
            previewRow.leadingText.enabled = enabled;
            previewRow.styleDD.enabled = enabled;
            previewRow.fontStyleDD.enabled = enabled;
        }

        /**
         * プレビューの全行を取得する
         * @returns {Array<object>} プレビュー行の配列
         */
        function getAllPreviewRows(dialogUi) {
            var previewRows = [];
            for (var levelRowIndex = 0; levelRowIndex < dialogUi.levelRows.length; levelRowIndex++) {
                previewRows.push(dialogUi.levelRows[levelRowIndex]);
            }
            previewRows.push(dialogUi.baseRow);
            previewRows.push(dialogUi.captionRow);
            return previewRows;
        }

        /**
         * ドロップダウンの項目を入れ替える
         * @param {DropDownList} dropdown 対象のドロップダウン
         * @param {Array<string>} items 新しい項目
         * @returns {void}
         */
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

        /**
         * プレビュー各行のフォントスタイルをまとめて選択する
         * @param {string} styleName 選択するスタイル名
         * @returns {void}
         */
        function selectAllPreviewFontStyleDropdowns(dialogUi, styleName) {
            var previewRows = getAllPreviewRows(dialogUi);
            for (var previewRowIndex = 0; previewRowIndex < previewRows.length; previewRowIndex++) {
                selectDropdownByText(previewRows[previewRowIndex].fontStyleDD, styleName);
            }
        }

        /**
         * テキスト設定のフォントスタイルをプレビューへ反映する
         * @returns {void}
         */
        function syncPreviewFontStylesFromTextSettings(dialogUi) {
            var selectedFontStyleName = getDropdownText(dialogUi.fontStyleDD);
            if (!selectedFontStyleName || selectedFontStyleName === getLabel("option.noFontChange")) return;
            selectAllPreviewFontStyleDropdowns(dialogUi, selectedFontStyleName);
        }

        /**
         * 見出しのフォントスタイルをプレビューへ反映する
         * @returns {void}
         */
        function syncHeadingPreviewFontStylesFromTextSettings(dialogUi) {
            var sourceFontStyleDropdown = (dialogUi.useSameFontRadio && dialogUi.useSameFontRadio.value)
                ? dialogUi.fontStyleDD
                : dialogUi.headingFontStyleDD;
            var selectedFontStyleName = getDropdownText(sourceFontStyleDropdown);
            if (!selectedFontStyleName || selectedFontStyleName === getLabel("option.noFontChange")) return;
            for (var levelRowIndex = 0; levelRowIndex < dialogUi.levelRows.length; levelRowIndex++) {
                selectDropdownByText(dialogUi.levelRows[levelRowIndex].fontStyleDD, selectedFontStyleName);
            }
        }

        /**
         * フォント指定モードに応じてコントロールの有効／無効を切り替える
         * @returns {void}
         */
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
                selectDropdownByText(dialogUi.fontDD, getLabel("option.noFontChange"));
                resetDropdownItems(dialogUi.fontStyleDD, [getLabel("option.noFontChange")]);
                selectDropdownByText(dialogUi.headingFontDD, getLabel("option.noFontChange"));
                resetDropdownItems(dialogUi.headingFontStyleDD, [getLabel("option.noFontChange")]);
            }
            var previewRows = getAllPreviewRows(dialogUi);
            for (var previewRowIndex = 0; previewRowIndex < previewRows.length; previewRowIndex++) {
                previewRows[previewRowIndex].fontStyleDD.enabled = enabled;
                if (!enabled) {
                    resetDropdownItems(previewRows[previewRowIndex].fontStyleDD, [getLabel("option.noFontChange")]);
                }
            }
        }

        /**
         * 本文のフォントスタイル候補を更新する
         * @returns {void}
         */
        function updateFontStyleDropdowns(dialogUi) {
            if (dialogUi.disableFontRadio && dialogUi.disableFontRadio.value) {
                syncFontSelectionEnabled(dialogUi);
                return;
            }
            var selectedFontFamily = getSelectedFontFamily(dialogUi);
            var fontStyleOptions = selectedFontFamily ? getFontStylesInFamily(selectedFontFamily) : [getLabel("option.noFontChange")];
            if (fontStyleOptions.length === 0) fontStyleOptions = [getLabel("option.noFontChange")];

            resetDropdownItems(dialogUi.fontStyleDD, fontStyleOptions);
            dialogUi.fontStyleDD.enabled = true;

            var selectedMasterFontStyleName = getDropdownText(dialogUi.fontStyleDD);
            var previewRows = getAllPreviewRows(dialogUi);
            for (var previewRowIndex = 0; previewRowIndex < previewRows.length; previewRowIndex++) {
                resetDropdownItems(previewRows[previewRowIndex].fontStyleDD, fontStyleOptions);
                previewRows[previewRowIndex].fontStyleDD.enabled = true;
            }
            if (selectedMasterFontStyleName && selectedMasterFontStyleName !== getLabel("option.noFontChange")) {
                selectAllPreviewFontStyleDropdowns(dialogUi, selectedMasterFontStyleName);
            }
        }

        /**
         * 見出しのフォントスタイル候補を更新する
         * @returns {void}
         */
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
            var fontStyleOptions = selectedFontFamily ? getFontStylesInFamily(selectedFontFamily) : [getLabel("option.noFontChange")];
            if (fontStyleOptions.length === 0) fontStyleOptions = [getLabel("option.noFontChange")];

            resetDropdownItems(dialogUi.headingFontStyleDD, fontStyleOptions);
            dialogUi.headingFontStyleDD.enabled = true;

            var selectedHeadingFontStyleName = getDropdownText(dialogUi.headingFontStyleDD);
            for (var levelRowIndex = 0; levelRowIndex < dialogUi.levelRows.length; levelRowIndex++) {
                resetDropdownItems(dialogUi.levelRows[levelRowIndex].fontStyleDD, fontStyleOptions);
                dialogUi.levelRows[levelRowIndex].fontStyleDD.enabled = true;
            }
            if (selectedHeadingFontStyleName && selectedHeadingFontStyleName !== getLabel("option.noFontChange")) {
                syncHeadingPreviewFontStylesFromTextSettings(dialogUi);
            }
        }

        /**
         * 見出しレベルごとのフォントスタイル名を集める
         * @returns {Array<string>} フォントスタイル名の配列
         */
        function getHeadingLevelFontStyleNames(dialogUi) {
            var levelFontStyleNames = [];
            for (var levelRowIndex = 0; levelRowIndex < dialogUi.levelRows.length; levelRowIndex++) {
                levelFontStyleNames.push(getHeadingFontStyleName(dialogUi, dialogUi.levelRows[levelRowIndex]));
            }
            return levelFontStyleNames;
        }

        /**
         * ダイアログの入力内容を設定オブジェクトにまとめる
         * @returns {object} 適用に使う設定
         */
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

        /**
         * 現在の設定でプレビューを描き直す
         * @returns {void}
         */
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

        /**
         * 選択範囲の文字オーバーライドを消去する
         * @returns {void}
         */
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

        /**
         * 選択があればオーバーライドを消去する
         * @returns {void}
         */
        function clearOverridesIfActive(dialogUi, forceClear) {
            // プレビュー時のみオーバーライドをクリア（破壊的操作回避）
            if (!forceClear && !dialogUi.previewCheck.value) return;
            clearTextOverridesInSelection();
            try { app.menuActions.itemByID(8489).invoke(); } catch (e) { }
            try { app.redraw(); } catch (e2) { }
        }

        /**
         * フォント変更後にプレビューを更新する
         * @returns {void}
         */
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

        /**
         * フォント指定モードを切り替える
         * @param {string} mode "same" / "separate" / "none"
         * @returns {void}
         */
        function setFontOptionMode(dialogUi, mode) {
            dialogUi.useSameFontRadio.value = (mode === "same");
            dialogUi.separateFontRadio.value = (mode === "separate");
            dialogUi.disableFontRadio.value = (mode === "disable");
        }

        /**
         * ダイアログ全体のイベントを結び付ける
         * @returns {void}
         */
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
            alert(getLabel("error.invalidBaseSize"));
            return null;
        }
        return collectTypescaleSettings(dialogUi);
    }

    /**
     * 段落スタイルへ各プロパティを適用する
     * @param {ParagraphStyle} paragraphStyle 対象の段落スタイル
     * @param {object} props 適用するプロパティ
     * @returns {void}
     */
    function setParagraphStyleProps(targetDocument, styleName, size, font, leading, spaceAfter, kerningMethod, silent) {
        var style = findParagraphStyle(targetDocument, styleName);
        if (style === null) {
            if (!silent) alert(formatLabel("error.missingParagraphStyle", styleName));
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

    /**
     * 名前から段落スタイルを探す
     * @param {Document} doc 対象ドキュメント
     * @param {string} styleName 段落スタイル名
     * @param {boolean} silent 見つからないときに警告を出さないなら true
     * @returns {ParagraphStyle|null} 段落スタイル。見つからない場合は null
     */
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