#target indesign

/*
概要 / Overview

【日本語】
概要：
InDesign の段落スタイルに対して、本文サイズを基準にタイプスケールで見出し／本文／キャプションのサイズを一括設定します。
単純倍率のスケールと、Browser Default のようなプリセットスケールを scaleOptions で管理します。
本文／見出しのフォント指定は、共通指定・別指定・変更しない、から選択できます。
見出し側は本文フォントを参照しつつ、ウエイト／スタイルのみ変更できます。
行送り、カーニング、段落前後のアキ、サイズの丸め単位を指定できます。
本文の段落前後のアキの初期値は 15% です。
プレビュー上でサイズ／段落前後のアキを確認し、必要に応じて個別調整できます。
フォント一覧はキャッシュして高速化し、必要に応じてキャッシュをクリアして再読み込みできます。

【English】
This InDesign script batch-configures paragraph styles for headings, body text, and captions using a type scale.
You can define the base body size, scale ratio, number of heading levels, size rounding, leading (percentage-based), space before/after, and kerning (separately for body and headings).

The font options panel provides three modes:
- Use same font for body and headings
- Specify separately for body and headings
- Do not change fonts

When a font family is selected, available font styles (weights) are automatically collected and applied.
In separate mode, headings reference the body font by default and can optionally override it.

The preview panel displays font size and space before for each level (h1–h6), body, and caption.
Live preview allows real-time confirmation while adjusting settings.

Headings are set to left alignment, while body text and captions use left-justified alignment.
Space before and space after can be specified independently for headings and body text.

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
var DEFAULT_SPACE_BEFORE_PERCENT = 10; // 見出しの段落前のアキ(%)
var DEFAULT_SPACE_AFTER_PERCENT = 10; // 見出しの段落後のアキ(%)
var DEFAULT_BODY_SPACE_BEFORE_PERCENT = 15; // 本文の段落前のアキ(%)
var DEFAULT_BODY_SPACE_AFTER_PERCENT = 15; // 本文の段落後のアキ(%)

/* =========================================

    バージョンとローカライズ / Version and Localization

   ========================================= */

var SCRIPT_VERSION = "v1.1.5";

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
    fontSettingsPanel: { ja: "フォント指定オプション", en: "Font Assignment" },
    disableFontSelection: { ja: "フォントを変更しない", en: "Do not change fonts" },
    useSameFontForBodyAndHeading: { ja: "本文と見出しで共通", en: "Use same font for body and headings" },
    separateFontForBodyAndHeading: {
        ja: "本文と見出しで別々に指定", en: "Specify body and heading fonts separately"
    },
    refBodyFont: {
        ja: "本文のフォントを参照", en: "Reference body font"
    },
    bodyTextPanel: { ja: "本文", en: "Body" },
    headingTextPanel: { ja: "見出し", en: "Headings" },
    baseSizeBody: { ja: "基準サイズ", en: "Base Size (Body)" },
    font: { ja: "フォント", en: "Font" },
    fontStyle: { ja: "スタイル", en: "Style" },
    bodyLeadingRatio: { ja: "行送り", en: "Leading Ratio (Body)" },
    headingLeadingRatio: { ja: "行送り", en: "Leading Ratio (Headings)" },
    spaceBeforeRatio: { ja: "段落前のアキ", en: "Space Before" },
    spaceAfterRatio: { ja: "段落後のアキ", en: "Space After" },
    kerningMethod: { ja: "カーニング", en: "Kerning Method" },
    kerningJapaneseMono: { ja: "和文等幅", en: "Japanese Mono" },
    kerningMetrics: { ja: "メトリクス", en: "Metrics" },
    kerningOptical: { ja: "オプティカル", en: "Optical" },
    scaleSettingsPanel: { ja: "スケール設定", en: "Scale Settings" },
    scaleRatio: { ja: "スケール方式", en: "Scale Method" },
    headingLevelCount: { ja: "見出しレベル数", en: "Heading Levels" },
    sizeRounding: { ja: "サイズの丸め", en: "Size Rounding" },
    roundInteger: { ja: "整数", en: "Integer" },
    fontStyleHeader: { ja: "スタイル", en: "Style" },
    roundFirstDecimal: { ja: "小数点第1位", en: "1 decimal place" },
    roundSecondDecimal: { ja: "小数点第2位", en: "2 decimal places" },
    previewPanel: { ja: "段落スタイルとサイズプレビュー", en: "Paragraph Styles & Size Preview" },
    levelHeader: { ja: "レベル", en: "Level" },
    fontSizeHeader: { ja: "サイズ", en: "Size" },
    spaceBeforeHeader: { ja: "段落前アキ", en: "Space Before" },
    spaceAfterHeader: { ja: "段落後のアキ", en: "Space After" },
    paragraphStyleHeader: { ja: "段落スタイル", en: "Paragraph Style" },
    levelPrefix: { ja: "レベル", en: "Level " },
    baseBodyPreview: { ja: "基準（本文）", en: "Base (Body)" },
    captionPreview: { ja: "キャプション", en: "Caption" },
    clearFontCache: { ja: "キャッシュをクリアして再読み込み", en: "Clear cache & reload" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" },
    loadingTitle: { ja: "処理中", en: "Processing" },
    loadingFonts: { ja: "フォント情報を読み込んでいます…", en: "Loading font information..." },
    buildingCacheNote: { ja: "キャッシュを作成しています。（初回は時間がかかります）", en: "Building cache (first run may take longer)" },
    notAvailable: { ja: "—", en: "—" },
    ratioMinorSecond: { ja: "短2度", en: "Minor Second" },
    ratioMajorSecond: { ja: "長2度", en: "Major Second" },
    ratioMinorThird: { ja: "短3度", en: "Minor Third" },
    ratioMajorThird: { ja: "長3度", en: "Major Third" },
    ratioGoldenHalf: { ja: "黄金比：1/2", en: "Golden Ratio: ½" },
    ratioPerfectFourth: { ja: "完全4度", en: "Perfect Fourth" },
    ratioAugmentedFourth: { ja: "増4度", en: "Augmented Fourth" },
    ratioGolden: { ja: "黄金比", en: "Golden Ratio" },
    ratioBrowserDefault: { ja: "ブラウザー既定", en: "Browser Default" }
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
    var subMessageText = palette.add("statictext", undefined, L("buildingCacheNote"));
    subMessageText.preferredSize.width = 260;

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

/* フォント一覧のキャッシュ：app.fonts.everyItem().getElements() が遅いので、ファミリー名・スタイル名のテーブルをディスクに保存して再利用する */
var FONT_CACHE_VERSION = "v1";
var FONT_CACHE_FILENAME = "TypeScaleStyleApplier_fontcache.txt";

function getFontCacheFile() {
    try {
        var dir = new Folder(Folder.userData.fsName + "/TypeScaleStyleApplier");
        if (!dir.exists) dir.create();
        return new File(dir.fsName + "/" + FONT_CACHE_FILENAME);
    } catch (e) {
        return null;
    }
}

function loadFontCache(currentFontCount) {
    var file = getFontCacheFile();
    if (!file || !file.exists) return null;
    if (!file.open("r")) return null;
    file.encoding = "UTF-8";
    try {
        var headerLine = file.readln();
        if (!headerLine) return null;
        var headerParts = headerLine.split("\t");
        if (headerParts[0] !== FONT_CACHE_VERSION) return null;
        var cachedCount = parseInt(headerParts[1], 10);
        if (isNaN(cachedCount) || cachedCount !== currentFontCount) return null;

        var families = [];
        var fontMap = {};
        while (!file.eof) {
            var line = file.readln();
            if (!line) continue;
            var parts = line.split("\t");
            var family = parts[0];
            if (!family) continue;
            var styles = [];
            for (var i = 1; i < parts.length; i++) {
                if (parts[i] && parts[i].length > 0) styles.push(parts[i]);
            }
            families.push(family);
            fontMap[family] = styles;
        }
        return { families: families, fontMap: fontMap };
    } catch (e) {
        return null;
    } finally {
        try { file.close(); } catch (e2) { }
    }
}

function saveFontCache(fontCount, families, fontMap) {
    var file = getFontCacheFile();
    if (!file) return false;
    if (!file.open("w")) return false;
    file.encoding = "UTF-8";
    try {
        file.writeln(FONT_CACHE_VERSION + "\t" + fontCount);
        for (var i = 0; i < families.length; i++) {
            var family = families[i];
            var styles = fontMap[family] || [];
            file.writeln(family + "\t" + styles.join("\t"));
        }
        return true;
    } catch (e) {
        return false;
    } finally {
        try { file.close(); } catch (e2) { }
    }
}

/* ウエイトランクテーブル（TypefaceSampler.jsx より移植） */
/* IIFE 外に置くのは、ExtendScript の hoisting で IIFE 内 var が初期化前に参照されるのを避けるため */
var WEIGHT_GROUPS = [
    ["hairline"],                                                                  // 0
    ["ultra thin", "ultrathin", "ut"],                                              // 1
    ["thin", "th"],                                                                 // 2
    ["default"],                                                                    // 3
    ["ultralight", "ultra light", "ultlt", "ul"],                                   // 4
    ["extralight", "extra light", "el", "xlight", "xl"],                            // 5
    ["lightsemi"],                                                                  // 6
    ["light", "lt", "lite", "l", "ライト"],                                          // 7
    ["lb"],                                                                         // 8
    ["book", "bk"],                                                                 // 9
    ["n"],                                                                          // 10
    ["middle"],                                                                     // 11
    ["regular", "roman", "normal", "r", "レギュラー", "標準", "中"],                 // 12
    ["rb"],                                                                         // 13
    ["medium", "md", "ミディアム", "m"],                                            // 14
    ["semibold", "semi bold", "sb"],                                                // 15
    ["demibold", "demi bold", "db", "デミボールド", "demi", "d", "demixtra"],        // 16
    ["bold", "bd", "ボールド", "b"],                                                 // 17
    ["extrabold", "extra bold", "xbold", "エクストラボールド", "e", "eb", "xb"],     // 18
    ["heavy", "h"],                                                                 // 19
    ["black"],                                                                      // 20
    ["xblack", "extra black", "extrablack"],                                        // 21
    ["ultra", "u", "ub", "ultra black", "ultrablack"]                               // 22
];
var WEIGHT_REGULAR_INDEX = 12;
var WEIGHT_REGULAR_SINGLES = [
    "display", "compressed", "comp", "compact", "expanded", "extended", "semiextended",
    "ultracondensed", "extracondensed", "semicondensed", "cond", "condensed", "wide",
    "headline", "text", "low", "micro", "extra compressed", "semi expanded", "semiexpanded"
];

function getStyleWeightBaseRank(style, family) {
    var familyLower = (family || "").toLowerCase();
    var words = style.split(/\s+/);

    var wMatch = style.match(/^w(\d)$/);
    if (wMatch !== null) return parseInt(wMatch[1], 10);
    var w3Match = style.match(/^w(\d{3})$/);
    if (w3Match !== null) return parseInt(w3Match[1], 10);
    var headNumMatch = style.match(/^(\d{1,3})(?=\D|$)/);
    if (headNumMatch) return parseInt(headNumMatch[1], 10);

    // 特例：Helvetica Neue / Tazugane / Univers Next + Ultra Light → 999
    if (
        (/helveticaneue/.test(familyLower) || /tazugane/.test(familyLower) || /universnextpro/.test(familyLower)) &&
        /ultralight|ultra light|ultlt/.test(style)
    ) {
        return 999;
    }

    if (words.length === 1 && /^(italic|oblique|it|wide)$/.test(words[0])) {
        return 1000 + WEIGHT_REGULAR_INDEX;
    }
    if (words.length === 1) {
        for (var sIndex = 0; sIndex < WEIGHT_REGULAR_SINGLES.length; sIndex++) {
            if (words[0] === WEIGHT_REGULAR_SINGLES[sIndex]) return 1000 + WEIGHT_REGULAR_INDEX;
        }
    }

    for (var groupIndex = 0; groupIndex < WEIGHT_GROUPS.length; groupIndex++) {
        for (var termIndex = 0; termIndex < WEIGHT_GROUPS[groupIndex].length; termIndex++) {
            if (style === WEIGHT_GROUPS[groupIndex][termIndex]) return 1000 + groupIndex;
        }
    }

    var allTerms = [];
    for (var gi = 0; gi < WEIGHT_GROUPS.length; gi++) {
        for (var ti = 0; ti < WEIGHT_GROUPS[gi].length; ti++) {
            allTerms.push({ term: WEIGHT_GROUPS[gi][ti], index: gi });
        }
    }
    allTerms.sort(function (a, b) { return b.term.length - a.term.length; });
    for (var aIndex = 0; aIndex < allTerms.length; aIndex++) {
        var term = allTerms[aIndex].term;
        var pattern = new RegExp("\\b" + term.replace(/[-\/\\^$*+?.()|[\]{}]/g, "\\$&") + "\\b");
        if (pattern.test(style)) return 1000 + allTerms[aIndex].index;
    }

    return 1000 + WEIGHT_REGULAR_INDEX;
}

function getStyleWeightRank(styleName, familyName) {
    var raw = (styleName || "").toString();
    var style = raw.toLowerCase().replace(/[_\-]+/g, " ").replace(/^\s+|\s+$/g, "");
    var words = style.split(/\s+/);
    var baseRank = getStyleWeightBaseRank(style, familyName);
    var offset = 0;

    var flags = {
        hasText: false, hasHeadline: false, hasCondensed: false, hasCn: false,
        hasExpanded: false, hasExtended: false, hasUltraCondensed: false,
        hasExtraCondensed: false, hasSemiCondensed: false, hasCompressed: false,
        hasExtraCompressed: false, hasCompact: false, hasDisplay: false,
        hasMicro: false, hasLow: false, hasWide: false
    };
    for (var w = 0; w < words.length; w++) {
        var word = words[w], next = words[w + 1];
        if (word === "text") flags.hasText = true;
        if (word === "headline") flags.hasHeadline = true;
        if (word === "cond" || word === "condensed") flags.hasCondensed = true;
        if (word === "cn") flags.hasCn = true;
        if (word === "expanded") flags.hasExpanded = true;
        if (word === "extended") flags.hasExtended = true;
        if (word === "semiextended" || (word === "semi" && next === "extended")) flags.hasExtended = true;
        if (word === "semiexpanded" || (word === "semi" && next === "expanded")) flags.hasExpanded = true;
        if (word === "ultracondensed" || (word === "ultra" && next === "condensed")) flags.hasUltraCondensed = true;
        if (word === "extracondensed" || (word === "extra" && next === "condensed")) flags.hasExtraCondensed = true;
        if (word === "semicondensed" || (word === "semi" && next === "condensed")) flags.hasSemiCondensed = true;
        if (word === "compressed" || word === "comp") flags.hasCompressed = true;
        if (word === "extra" && next === "compressed") flags.hasExtraCompressed = true;
        if (word === "compact") flags.hasCompact = true;
        if (word === "display") flags.hasDisplay = true;
        if (word === "micro") flags.hasMicro = true;
        if (word === "low") flags.hasLow = true;
        if (word === "wide") flags.hasWide = true;
    }
    var isItalic = /italic|oblique|slanted|inclined|kursiv|\bit\b/.test(style);

    if (flags.hasDisplay) offset += 100;
    if (flags.hasCompressed) offset += 200;
    if (flags.hasCompact) offset += 300;
    if (flags.hasExpanded) offset += 400;
    if (flags.hasExtended) offset += 500;
    if (flags.hasUltraCondensed) offset += 600;
    if (flags.hasExtraCondensed) offset += 700;
    if (flags.hasSemiCondensed) offset += 850;
    if (flags.hasCondensed || flags.hasCn || flags.hasWide || flags.hasSemiCondensed || flags.hasExtraCompressed) offset += 900;
    if (flags.hasHeadline) offset += 1000;
    if (flags.hasText) offset += 1100;
    if (flags.hasLow) offset += 1200;
    if (flags.hasMicro) offset += 1250;
    if (flags.hasWide) offset += 1275;
    if (flags.hasExtraCompressed) offset += 150;
    if (isItalic) offset += 1300;

    return baseRank + offset;
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

    function computeSizes(base, ratio, levelCount, multipliers, captionMultiplier) {
        /* 本文サイズを基準に、タイプスケールで見出しとキャプションのサイズを算出 / Calculate heading and caption sizes from the body size using the type scale */
        /* multipliers が与えられればレベルごとの倍率を直接使用、なければ ratio の累乗で算出 */
        var headingSizes = [];
        if (multipliers && multipliers.length > 0) {
            var fixedCount = Math.min(levelCount, multipliers.length);
            for (var fixedIndex = 1; fixedIndex <= fixedCount; fixedIndex++) {
                headingSizes.push(base * multipliers[fixedIndex - 1]);
            }
            var captionSize = (typeof captionMultiplier === "number") ? base * captionMultiplier : base;
            return { headingSizes: headingSizes, base: base, caption: captionSize };
        }
        for (var levelIndex = 1; levelIndex <= levelCount; levelIndex++) {
            headingSizes.push(base * Math.pow(ratio, levelCount - levelIndex + 1));
        }
        return { headingSizes: headingSizes, base: base, caption: base / ratio };
    }

    function applyTypescaleSettings(targetDocument, typescaleSettings, silent, unit) {
        var computedSizes = computeSizes(typescaleSettings.baseSize, typescaleSettings.ratio, typescaleSettings.levelCount, typescaleSettings.scaleMultipliers, typescaleSettings.captionMultiplier);
        // 無効値時はデフォルトにフォールバック
        var headingSpaceAfterPercent = (typeof typescaleSettings.headingSpaceAfterPercent === "number" && typescaleSettings.headingSpaceAfterPercent >= 0)
            ? typescaleSettings.headingSpaceAfterPercent
            : DEFAULT_SPACE_AFTER_PERCENT;
        var headingSpaceBeforePercent = (typeof typescaleSettings.headingSpaceBeforePercent === "number" && typescaleSettings.headingSpaceBeforePercent >= 0)
            ? typescaleSettings.headingSpaceBeforePercent
            : DEFAULT_SPACE_BEFORE_PERCENT;
        var bodySpaceAfterPercent = (typeof typescaleSettings.bodySpaceAfterPercent === "number" && typescaleSettings.bodySpaceAfterPercent >= 0)
            ? typescaleSettings.bodySpaceAfterPercent
            : DEFAULT_BODY_SPACE_AFTER_PERCENT;
        var bodySpaceBeforePercent = (typeof typescaleSettings.bodySpaceBeforePercent === "number" && typescaleSettings.bodySpaceBeforePercent >= 0)
            ? typescaleSettings.bodySpaceBeforePercent
            : DEFAULT_BODY_SPACE_BEFORE_PERCENT;
        function applyParagraphStyleSettings(styleName, sizeInUnit, leadingMult, isHeading, fontFamilyName, fontStyleName, sizeOverrideInUnit, spaceBeforeOverrideInUnit, spaceAfterOverrideInUnit) {
            if (!styleName) return;
            var effectiveSize = (typeof sizeOverrideInUnit === "number") ? sizeOverrideInUnit : sizeInUnit;
            var rounded = roundTo(effectiveSize, typescaleSettings.roundDigits);
            var sizePt = toPoints(rounded, unit);
            var leadingPt = (typeof leadingMult === "number") ? sizePt * leadingMult : null;
            var spaceBeforePercent = isHeading ? headingSpaceBeforePercent : bodySpaceBeforePercent;
            var spaceAfterPercent = isHeading ? headingSpaceAfterPercent : bodySpaceAfterPercent;
            var spaceBeforePt;
            if (typeof spaceBeforeOverrideInUnit === "number") {
                spaceBeforePt = toPoints(roundTo(spaceBeforeOverrideInUnit, typescaleSettings.roundDigits), unit);
            } else {
                spaceBeforePt = sizePt * spaceBeforePercent / 100;
            }
            var spaceAfterPt;
            if (typeof spaceAfterOverrideInUnit === "number") {
                spaceAfterPt = toPoints(roundTo(spaceAfterOverrideInUnit, typescaleSettings.roundDigits), unit);
            } else {
                spaceAfterPt = sizePt * spaceAfterPercent / 100;
            }
            // フォントファミリー＋スタイルで解決。未指定時はファミリー内の推奨スタイルを使用
            var fontToUse = null;
            if (fontFamilyName) {
                fontToUse = fontStyleName ? findFontByFamilyAndStyle(fontFamilyName, fontStyleName) : findFontInFamily(fontFamilyName);
            }
            var kerningMethod = isHeading ? typescaleSettings.headingKerningMethod : typescaleSettings.bodyKerningMethod;
            setParagraphStyleProps(targetDocument, styleName, sizePt, fontToUse, leadingPt, spaceAfterPt, spaceBeforePt, kerningMethod, isHeading, silent);
        }
        applyParagraphStyleSettings(typescaleSettings.baseStyleName, computedSizes.base, typescaleSettings.bodyLeading, false, typescaleSettings.fontFamily, typescaleSettings.baseFontStyleName, typescaleSettings.baseSizeOverride, typescaleSettings.baseSpaceBeforeOverride, typescaleSettings.baseSpaceAfterOverride);
        applyParagraphStyleSettings(typescaleSettings.captionStyleName, computedSizes.caption, typescaleSettings.bodyLeading, false, typescaleSettings.fontFamily, typescaleSettings.captionFontStyleName, typescaleSettings.captionSizeOverride, typescaleSettings.captionSpaceBeforeOverride, typescaleSettings.captionSpaceAfterOverride);
        for (var levelNumber = 1; levelNumber <= typescaleSettings.levelCount; levelNumber++) {
            var name = typescaleSettings.levelStyleNames && typescaleSettings.levelStyleNames[levelNumber - 1];
            var fontStyleName = typescaleSettings.levelFontStyleNames && typescaleSettings.levelFontStyleNames[levelNumber - 1];
            var levelSizeOverride = typescaleSettings.levelSizeOverrides ? typescaleSettings.levelSizeOverrides[levelNumber - 1] : null;
            var levelSpaceBeforeOverride = typescaleSettings.levelSpaceBeforeOverrides ? typescaleSettings.levelSpaceBeforeOverrides[levelNumber - 1] : null;
            var levelSpaceAfterOverride = typescaleSettings.levelSpaceAfterOverrides ? typescaleSettings.levelSpaceAfterOverrides[levelNumber - 1] : null;
            applyParagraphStyleSettings(name, computedSizes.headingSizes[levelNumber - 1], typescaleSettings.headingLeading, true, typescaleSettings.headingFontFamily, fontStyleName, levelSizeOverride, levelSpaceBeforeOverride, levelSpaceAfterOverride);
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

        var currentFontCount = 0;
        try { currentFontCount = app.fonts.length; } catch (eCount) { currentFontCount = 0; }

        // ディスクキャッシュが有効ならそれを使う（Font オブジェクトは findFontByFamilyAndStyle で遅延解決）
        var cached = loadFontCache(currentFontCount);
        if (cached) {
            _fontInfo = {
                families: cached.families,
                fontMap: cached.fontMap,
                fontObjectMap: {}
            };
            return _fontInfo;
        }

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
            var sortFamilyName = families[familyIndex];
            fontMap[sortFamilyName].sort(function (a, b) {
                return getStyleWeightRank(a, sortFamilyName) - getStyleWeightRank(b, sortFamilyName);
            });
        }

        _fontInfo = {
            families: families,
            fontMap: fontMap,
            fontObjectMap: fontObjectMap
        };
        saveFontCache(currentFontCount > 0 ? currentFontCount : fonts.length, families, fontMap);
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
        // キャッシュ経由で fontObjectMap が空の場合は app.fonts.itemByName で遅延解決
        try {
            var resolved = app.fonts.itemByName(fontFullName);
            if (resolved && resolved.isValid) {
                fontInfo.fontObjectMap[fontFullName] = resolved;
                return resolved;
            }
        } catch (e) { }
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
        var defaultRoundDigits = 1;
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

        var scaleOptions = [
            { type: "scale", key: "ratioMinorSecond", ratio: 1.067 },
            { type: "scale", key: "ratioMajorSecond", ratio: 1.125 },
            { type: "scale", key: "ratioMinorThird", ratio: 1.2 },
            { type: "scale", key: "ratioMajorThird", ratio: 1.25 },
            { type: "scale", key: "ratioGoldenHalf", ratio: 1.309 },
            { type: "scale", key: "ratioPerfectFourth", ratio: 1.333 },
            { type: "scale", key: "ratioAugmentedFourth", ratio: 1.414 },
            { type: "scale", key: "ratioGolden", ratio: 1.618 },
            {
                type: "preset",
                key: "ratioBrowserDefault",
                multipliers: [2.00, 1.50, 1.17, 1.00],
                captionMultiplier: 0.83,
                forcedLevelCount: 4
            }
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
            setupPanel(bodyPanel, 4);
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
            var headingFontOptions = [L("refBodyFont")].concat(fontFamilies);
            var headingFontDD = headingFontGrp.add("dropdownlist", undefined, headingFontOptions);
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

            return {
                headingFontDD: headingFontDD,
                headingFontStyleDD: headingFontStyleDD,
                leadingHeadingInput: leadingHeadingInput,
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

            var scaleGrp = addLabeledGroup(optionsPanel, labelText("scaleRatio"), OPTIONS_LABEL_WIDTH);
            var scaleLabels = [];
            for (var scaleOptionIndex = 0; scaleOptionIndex < scaleOptions.length; scaleOptionIndex++) {
                var scaleLabel = L(scaleOptions[scaleOptionIndex].key);
                if (scaleOptions[scaleOptionIndex].type === "scale") {
                    scaleLabel += "  " + scaleOptions[scaleOptionIndex].ratio;
                }
                scaleLabels.push(scaleLabel);
            }
            var scaleDD = scaleGrp.add("dropdownlist", undefined, scaleLabels);
            scaleDD.preferredSize.width = 200;
            for (var selectedScaleOptionIndex = 0; selectedScaleOptionIndex < scaleOptions.length; selectedScaleOptionIndex++) {
                if (scaleOptions[selectedScaleOptionIndex].type === "scale" && scaleOptions[selectedScaleOptionIndex].ratio === defaultRatio) {
                    scaleDD.selection = selectedScaleOptionIndex;
                    break;
                }
            }
            if (!scaleDD.selection) scaleDD.selection = 0;

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
                scaleDD: scaleDD,
                levelRadios: levelRadios,
                roundRadios: roundRadios
            };
        }

        function createPreviewPanel(dialog) {
            var previewPanel = dialog.add("panel", undefined, L("previewPanel"));
            setupPanel(previewPanel, 2);

            var PREVIEW_LABEL_WIDTH = 84;
            var PREVIEW_SIZE_WIDTH = 80;
            var PREVIEW_SPACE_BEFORE_WIDTH = 80;
            var PREVIEW_SPACE_AFTER_WIDTH = 80;

            var headerRow = previewPanel.add("group");
            headerRow.orientation = "row";
            headerRow.alignChildren = "left";
            headerRow.add("statictext", undefined, labelText("levelHeader")).preferredSize.width = PREVIEW_LABEL_WIDTH;
            headerRow.add("statictext", undefined, labelText("fontSizeHeader")).preferredSize.width = PREVIEW_SIZE_WIDTH;
            headerRow.add("statictext", undefined, labelText("spaceBeforeHeader")).preferredSize.width = PREVIEW_SPACE_BEFORE_WIDTH;
            headerRow.add("statictext", undefined, labelText("spaceAfterHeader")).preferredSize.width = PREVIEW_SPACE_AFTER_WIDTH;
            headerRow.add("statictext", undefined, labelText("paragraphStyleHeader")).preferredSize.width = 100;
            var fontStyleHeader = headerRow.add("statictext", undefined, labelText("fontStyleHeader"));
            fontStyleHeader.preferredSize.width = 100;
            fontStyleHeader.enabled = true;

            var previewHeaderSpacer = previewPanel.add("group");
            previewHeaderSpacer.preferredSize.height = 4;

            function selectDropdownByCandidates(dropdownList, candidates) {
                for (var candidateIndex = 0; candidateIndex < candidates.length; candidateIndex++) {
                    for (var itemIndex = 0; itemIndex < dropdownList.items.length; itemIndex++) {
                        if (dropdownList.items[itemIndex].text === candidates[candidateIndex]) {
                            dropdownList.selection = itemIndex;
                            return true;
                        }
                    }
                }
                if (dropdownList.items.length > 0) dropdownList.selection = 0;
                return false;
            }

            function createPreviewRow(parent, label, defaultStyleNames, isHeading) {
                var row = parent.add("group");
                row.orientation = "row";
                row.alignChildren = "center";
                var labelText = row.add("statictext", undefined, label);
                labelText.preferredSize.width = PREVIEW_LABEL_WIDTH;
                var sizeText = row.add("edittext", undefined, "");
                sizeText.preferredSize.width = PREVIEW_SIZE_WIDTH;
                var spaceBeforeText = row.add("edittext", undefined, "");
                spaceBeforeText.preferredSize.width = PREVIEW_SPACE_BEFORE_WIDTH;
                var spaceAfterText = row.add("edittext", undefined, "");
                spaceAfterText.preferredSize.width = PREVIEW_SPACE_AFTER_WIDTH;
                var styleDD = row.add("dropdownlist", undefined, styleNames);
                styleDD.preferredSize.width = 100;
                var candidateList = (defaultStyleNames instanceof Array) ? defaultStyleNames : [defaultStyleNames];
                selectDropdownByCandidates(styleDD, candidateList);

                var fontStyleDD = row.add("dropdownlist", undefined, [L("noFontChange")]);
                fontStyleDD.preferredSize.width = 100;
                fontStyleDD.selection = 0;
                fontStyleDD.enabled = true;

                return {
                    lbl: labelText,
                    sizeText: sizeText,
                    spaceBeforeText: spaceBeforeText,
                    spaceAfterText: spaceAfterText,
                    styleDD: styleDD,
                    fontStyleDD: fontStyleDD,
                    isHeading: !!isHeading,
                    sizeOverride: null,
                    spaceBeforeOverride: null,
                    spaceAfterOverride: null
                };
            }

            var levelRows = [];
            for (var levelNumber = 1; levelNumber <= 6; levelNumber++) {
                var levelCandidates = ["h" + levelNumber, "Heading " + levelNumber];
                levelRows.push(createPreviewRow(previewPanel, L("levelPrefix") + levelNumber, levelCandidates, true));
            }

            return {
                levelRows: levelRows,
                baseRow: createPreviewRow(previewPanel, L("baseBodyPreview"), ["p", "Normal"], false),
                captionRow: createPreviewRow(previewPanel, L("captionPreview"), "p.caption", false)
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

            var clearCacheButton = leftButtonColumn.add("button", undefined, L("clearFontCache"));

            var centerButtonColumn = bottomRow.add("group");
            centerButtonColumn.alignment = ["fill", "fill"];
            centerButtonColumn.minimumSize.width = 0;

            var rightButtonColumn = bottomRow.add("group");
            rightButtonColumn.orientation = "row";
            rightButtonColumn.alignChildren = ["right", "center"];
            rightButtonColumn.alignment = ["right", "center"];

            rightButtonColumn.add("button", undefined, L("cancel"), { name: "cancel" });
            rightButtonColumn.add("button", undefined, L("ok"), { name: "ok" });

            return { clearCacheButton: clearCacheButton };
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
                bodyKerningDD: textSettingsUi.bodyKerningDD,
                headingKerningDD: textSettingsUi.headingKerningDD,
                scaleDD: scaleSettingsUi.scaleDD,
                levelRadios: scaleSettingsUi.levelRadios,
                roundRadios: scaleSettingsUi.roundRadios,
                levelRows: previewUi.levelRows,
                baseRow: previewUi.baseRow,
                captionRow: previewUi.captionRow,
                clearCacheButton: buttonUi.clearCacheButton,
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

        function getCurrentScaleRatio(dialogUi) {
            if (!dialogUi.scaleDD.selection) return defaultRatio;
            var scaleOption = scaleOptions[dialogUi.scaleDD.selection.index];
            return (scaleOption && scaleOption.type === "scale") ? scaleOption.ratio : defaultRatio;
        }

        function getCurrentScaleOption(dialogUi) {
            if (!dialogUi.scaleDD.selection) return null;
            return scaleOptions[dialogUi.scaleDD.selection.index];
        }

        function getCurrentLevelCount(dialogUi) {
            var scaleOption = getCurrentScaleOption(dialogUi);
            if (scaleOption && scaleOption.type === "preset" && typeof scaleOption.forcedLevelCount === "number") {
                return scaleOption.forcedLevelCount;
            }
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

            // 常に本文フォントを参照（ウエイトのみ変更できるようにする）
            var bodyFont = getSelectedFontFamily(dialogUi);

            // 見出し側で別フォントが明示指定されている場合のみそれを優先
            if (dialogUi.separateFontRadio && dialogUi.separateFontRadio.value) {
                if (dialogUi.headingFontDD.selection && dialogUi.headingFontDD.selection.index !== 0) {
                    return dialogUi.headingFontDD.selection.text;
                }
            }

            return bodyFont;
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

        function getBodySpaceBeforePercent(dialogUi) {
            return DEFAULT_BODY_SPACE_BEFORE_PERCENT;
        }

        function getBodySpaceAfterPercent(dialogUi) {
            return DEFAULT_BODY_SPACE_AFTER_PERCENT;
        }

        function getHeadingSpaceBeforePercent(dialogUi) {
            return DEFAULT_SPACE_BEFORE_PERCENT;
        }

        function getHeadingSpaceAfterPercent(dialogUi) {
            return DEFAULT_SPACE_AFTER_PERCENT;
        }

        function getBaseSize(dialogUi) {
            return parsePositiveNumber(dialogUi.baseInput.text, null);
        }

        function formatLeadingValue(value, roundDigits) {
            if (typeof value !== "number" || isNaN(value)) return L("notAvailable");
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
            previewRow.spaceBeforeText.enabled = enabled;
            previewRow.spaceAfterText.enabled = enabled;
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
                selectDropdownByText(dialogUi.headingFontDD, L("refBodyFont"));
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

            // 共通指定時のみ本文フォントを参照表示にする
            if (dialogUi.useSameFontRadio && dialogUi.useSameFontRadio.value) {
                dialogUi.headingFontDD.enabled = false;
                dialogUi.headingFontDD.helpTip = L("refBodyFont");
            } else {
                dialogUi.headingFontDD.enabled = true;
                dialogUi.headingFontDD.helpTip = "";
            }

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

        function collectLevelSizeOverrides(dialogUi) {
            var arr = [];
            for (var i = 0; i < dialogUi.levelRows.length; i++) {
                arr.push(dialogUi.levelRows[i].sizeOverride);
            }
            return arr;
        }

        function collectLevelSpaceBeforeOverrides(dialogUi) {
            var arr = [];
            for (var i = 0; i < dialogUi.levelRows.length; i++) {
                arr.push(dialogUi.levelRows[i].spaceBeforeOverride);
            }
            return arr;
        }

        function collectLevelSpaceAfterOverrides(dialogUi) {
            var arr = [];
            for (var i = 0; i < dialogUi.levelRows.length; i++) {
                arr.push(dialogUi.levelRows[i].spaceAfterOverride);
            }
            return arr;
        }

        function collectTypescaleSettings(dialogUi) {
            var baseSize = getBaseSize(dialogUi);
            var scaleOption = getCurrentScaleOption(dialogUi);
            return {
                baseSize: baseSize,
                ratio: getCurrentScaleRatio(dialogUi),
                scaleMultipliers: (scaleOption && scaleOption.type === "preset") ? scaleOption.multipliers : null,
                captionMultiplier: (scaleOption && scaleOption.type === "preset") ? scaleOption.captionMultiplier : null,
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
                bodySpaceBeforePercent: getBodySpaceBeforePercent(dialogUi),
                bodySpaceAfterPercent: getBodySpaceAfterPercent(dialogUi),
                headingSpaceBeforePercent: getHeadingSpaceBeforePercent(dialogUi),
                headingSpaceAfterPercent: getHeadingSpaceAfterPercent(dialogUi),
                levelSizeOverrides: collectLevelSizeOverrides(dialogUi),
                levelSpaceBeforeOverrides: collectLevelSpaceBeforeOverrides(dialogUi),
                levelSpaceAfterOverrides: collectLevelSpaceAfterOverrides(dialogUi),
                baseSizeOverride: dialogUi.baseRow.sizeOverride,
                baseSpaceBeforeOverride: dialogUi.baseRow.spaceBeforeOverride,
                baseSpaceAfterOverride: dialogUi.baseRow.spaceAfterOverride,
                captionSizeOverride: dialogUi.captionRow.sizeOverride,
                captionSpaceBeforeOverride: dialogUi.captionRow.spaceBeforeOverride,
                captionSpaceAfterOverride: dialogUi.captionRow.spaceAfterOverride
            };
        }

        function updateTypescalePreview(dialogUi) {
            var baseSize = getBaseSize(dialogUi);
            var ratio = getCurrentScaleRatio(dialogUi);
            var levelCount = getCurrentLevelCount(dialogUi);
            var roundDigits = getCurrentRoundDigits(dialogUi);

            if (baseSize === null) {
                for (var levelRowIndex = 0; levelRowIndex < dialogUi.levelRows.length; levelRowIndex++) {
                    dialogUi.levelRows[levelRowIndex].sizeText.text = L("notAvailable");
                    dialogUi.levelRows[levelRowIndex].spaceBeforeText.text = L("notAvailable");
                    dialogUi.levelRows[levelRowIndex].spaceAfterText.text = L("notAvailable");
                    setRowEnabled(dialogUi.levelRows[levelRowIndex], false);
                }
                dialogUi.baseRow.sizeText.text = L("notAvailable");
                dialogUi.baseRow.spaceBeforeText.text = L("notAvailable");
                dialogUi.baseRow.spaceAfterText.text = L("notAvailable");
                dialogUi.captionRow.sizeText.text = L("notAvailable");
                dialogUi.captionRow.spaceBeforeText.text = L("notAvailable");
                dialogUi.captionRow.spaceAfterText.text = L("notAvailable");
                return;
            }

            var scaleOption = getCurrentScaleOption(dialogUi);
            var multipliers = (scaleOption && scaleOption.type === "preset") ? scaleOption.multipliers : null;
            var captionMultiplier = (scaleOption && scaleOption.type === "preset") ? scaleOption.captionMultiplier : null;
            var computedSizes = computeSizes(baseSize, ratio, levelCount, multipliers, captionMultiplier);
            var headingSpaceBeforeRatio = getHeadingSpaceBeforePercent(dialogUi) / 100;
            var bodySpaceBeforeRatio = getBodySpaceBeforePercent(dialogUi) / 100;
            var headingSpaceAfterRatio = getHeadingSpaceAfterPercent(dialogUi) / 100;
            var bodySpaceAfterRatio = getBodySpaceAfterPercent(dialogUi) / 100;
            for (var levelNumber = 1; levelNumber <= dialogUi.levelRows.length; levelNumber++) {
                var levelRow = dialogUi.levelRows[levelNumber - 1];
                if (levelNumber <= levelCount) {
                    var computedHeadingSize = computedSizes.headingSizes[levelNumber - 1];
                    var effectiveHeadingSize = (typeof levelRow.sizeOverride === "number") ? levelRow.sizeOverride : computedHeadingSize;
                    levelRow.sizeText.text = roundTo(effectiveHeadingSize, roundDigits) + " " + unitSym;
                    var effectiveHeadingSpaceBefore = (typeof levelRow.spaceBeforeOverride === "number") ? levelRow.spaceBeforeOverride : effectiveHeadingSize * headingSpaceBeforeRatio;
                    levelRow.spaceBeforeText.text = formatLeadingValue(effectiveHeadingSpaceBefore, roundDigits);
                    var effectiveHeadingSpaceAfter = (typeof levelRow.spaceAfterOverride === "number") ? levelRow.spaceAfterOverride : effectiveHeadingSize * headingSpaceAfterRatio;
                    levelRow.spaceAfterText.text = formatLeadingValue(effectiveHeadingSpaceAfter, roundDigits);
                    setRowEnabled(levelRow, true);
                } else {
                    levelRow.sizeText.text = L("notAvailable");
                    levelRow.spaceBeforeText.text = L("notAvailable");
                    levelRow.spaceAfterText.text = L("notAvailable");
                    setRowEnabled(levelRow, false);
                }
            }
            var effectiveBaseSize = (typeof dialogUi.baseRow.sizeOverride === "number") ? dialogUi.baseRow.sizeOverride : computedSizes.base;
            dialogUi.baseRow.sizeText.text = roundTo(effectiveBaseSize, roundDigits) + " " + unitSym;
            var effectiveBaseSpaceBefore = (typeof dialogUi.baseRow.spaceBeforeOverride === "number") ? dialogUi.baseRow.spaceBeforeOverride : effectiveBaseSize * bodySpaceBeforeRatio;
            dialogUi.baseRow.spaceBeforeText.text = formatLeadingValue(effectiveBaseSpaceBefore, roundDigits);
            var effectiveBaseSpaceAfter = (typeof dialogUi.baseRow.spaceAfterOverride === "number") ? dialogUi.baseRow.spaceAfterOverride : effectiveBaseSize * bodySpaceAfterRatio;
            dialogUi.baseRow.spaceAfterText.text = formatLeadingValue(effectiveBaseSpaceAfter, roundDigits);
            var effectiveCaptionSize = (typeof dialogUi.captionRow.sizeOverride === "number") ? dialogUi.captionRow.sizeOverride : computedSizes.caption;
            dialogUi.captionRow.sizeText.text = roundTo(effectiveCaptionSize, roundDigits) + " " + unitSym;
            var effectiveCaptionSpaceBefore = (typeof dialogUi.captionRow.spaceBeforeOverride === "number") ? dialogUi.captionRow.spaceBeforeOverride : effectiveCaptionSize * bodySpaceBeforeRatio;
            dialogUi.captionRow.spaceBeforeText.text = formatLeadingValue(effectiveCaptionSpaceBefore, roundDigits);
            var effectiveCaptionSpaceAfter = (typeof dialogUi.captionRow.spaceAfterOverride === "number") ? dialogUi.captionRow.spaceAfterOverride : effectiveCaptionSize * bodySpaceAfterRatio;
            dialogUi.captionRow.spaceAfterText.text = formatLeadingValue(effectiveCaptionSpaceAfter, roundDigits);

            applyTypescaleSettings(targetDocument, collectTypescaleSettings(dialogUi), true, unit);
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

        function clearOverridesIfActive(dialogUi) {
            clearTextOverridesInSelection();
            try { app.menuActions.itemByID(8489).invoke(); } catch (e) { }
            try { app.redraw(); } catch (e2) { }
        }

        function refreshPreviewAfterFontChange(dialogUi) {
            clearOverridesIfActive(dialogUi);
            updateTypescalePreview(dialogUi);
        }

        function clearFontCacheAndReload(dialogUi) {
            // ディスクキャッシュとメモリキャッシュをクリア
            try {
                var cacheFile = getFontCacheFile();
                if (cacheFile && cacheFile.exists) cacheFile.remove();
            } catch (eRemove) { }
            _fontInfo = null;

            // ファミリー／スタイル一覧を再取得してドロップダウンを差し替え
            var refreshedFamilies = getFontFamilyNames();
            var refreshedFontOptions = [L("noFontChange")].concat(refreshedFamilies);
            var refreshedHeadingFontOptions = [L("refBodyFont")].concat(refreshedFamilies);
            var previousFontText = getDropdownText(dialogUi.fontDD);
            var previousHeadingFontText = getDropdownText(dialogUi.headingFontDD);
            resetDropdownItems(dialogUi.fontDD, refreshedFontOptions);
            resetDropdownItems(dialogUi.headingFontDD, refreshedHeadingFontOptions);
            if (previousFontText) selectDropdownByText(dialogUi.fontDD, previousFontText);
            if (previousHeadingFontText) selectDropdownByText(dialogUi.headingFontDD, previousHeadingFontText);

            updateFontStyleDropdowns(dialogUi);
            updateHeadingFontStyleDropdowns(dialogUi);
            refreshPreviewAfterFontChange(dialogUi);
        }

        function setFontOptionMode(dialogUi, mode) {
            dialogUi.useSameFontRadio.value = (mode === "same");
            dialogUi.separateFontRadio.value = (mode === "separate");
            dialogUi.disableFontRadio.value = (mode === "disable");
        }

        function syncLevelRadiosWithScaleOption(dialogUi) {
            var scaleOption = getCurrentScaleOption(dialogUi);
            var forcedLevelCount = (scaleOption && scaleOption.type === "preset" && typeof scaleOption.forcedLevelCount === "number")
                ? scaleOption.forcedLevelCount
                : null;

            for (var radioIndex = 0; radioIndex < dialogUi.levelRadios.length; radioIndex++) {
                if (forcedLevelCount !== null) {
                    dialogUi.levelRadios[radioIndex].value = (levelOptions[radioIndex] === forcedLevelCount);
                    dialogUi.levelRadios[radioIndex].enabled = false;
                } else {
                    dialogUi.levelRadios[radioIndex].enabled = true;
                }
            }
        }

        function clearAllPreviewOverrides(dialogUi) {
            for (var levelClearIndex = 0; levelClearIndex < dialogUi.levelRows.length; levelClearIndex++) {
                dialogUi.levelRows[levelClearIndex].sizeOverride = null;
                dialogUi.levelRows[levelClearIndex].spaceBeforeOverride = null;
                dialogUi.levelRows[levelClearIndex].spaceAfterOverride = null;
            }
            dialogUi.baseRow.sizeOverride = null;
            dialogUi.baseRow.spaceBeforeOverride = null;
            dialogUi.baseRow.spaceAfterOverride = null;
            dialogUi.captionRow.sizeOverride = null;
            dialogUi.captionRow.spaceBeforeOverride = null;
            dialogUi.captionRow.spaceAfterOverride = null;
        }

        function changePreviewValueByArrowKey(editText, allowZero, applyValue) {
            editText.addEventListener("keydown", function (event) {
                if (event.keyName !== "Up" && event.keyName !== "Down") return;
                var value = parseFloat(editText.text);
                if (isNaN(value)) return;
                var direction = (event.keyName === "Up") ? 1 : -1;
                var keyboard = ScriptUI.environment.keyboardState;
                if (keyboard.shiftKey) {
                    var stepShift = 10;
                    if (direction > 0) {
                        value = Math.ceil((value + 1) / stepShift) * stepShift;
                    } else {
                        value = Math.floor((value - 1) / stepShift) * stepShift;
                    }
                } else if (keyboard.altKey) {
                    value += 0.1 * direction;
                } else {
                    value += 1 * direction;
                }
                if (keyboard.altKey) {
                    value = Math.round(value * 10) / 10;
                } else {
                    value = Math.round(value);
                }
                if (allowZero) {
                    if (value < 0) value = 0;
                } else {
                    if (value < 1) value = 1;
                }
                event.preventDefault();
                applyValue(value);
            });
        }

        function bindPreviewCellEdit(dialogUi, previewRow, sizeCallback) {
            var defaultSizeCallback = function (parsedValue) {
                previewRow.sizeOverride = parsedValue;
                updateTypescalePreview(dialogUi);
            };
            var resolvedSizeCallback = sizeCallback || defaultSizeCallback;
            previewRow.sizeText.onChange = function () {
                resolvedSizeCallback(parsePositiveNumber(previewRow.sizeText.text, null));
            };
            previewRow.spaceBeforeText.onChange = function () {
                var parsed = parseNonNegativeNumber(previewRow.spaceBeforeText.text, null);
                previewRow.spaceBeforeOverride = parsed;
                updateTypescalePreview(dialogUi);
            };
            previewRow.spaceAfterText.onChange = function () {
                var parsed = parseNonNegativeNumber(previewRow.spaceAfterText.text, null);
                previewRow.spaceAfterOverride = parsed;
                updateTypescalePreview(dialogUi);
            };
            changePreviewValueByArrowKey(previewRow.sizeText, false, resolvedSizeCallback);
            changePreviewValueByArrowKey(previewRow.spaceBeforeText, true, function (newValue) {
                previewRow.spaceBeforeOverride = newValue;
                updateTypescalePreview(dialogUi);
            });
            changePreviewValueByArrowKey(previewRow.spaceAfterText, true, function (newValue) {
                previewRow.spaceAfterOverride = newValue;
                updateTypescalePreview(dialogUi);
            });
        }

        function bindTypescaleDialogEvents(dialogUi) {
            changeValueByArrowKey(dialogUi.baseInput, function () {
                clearAllPreviewOverrides(dialogUi);
                updateTypescalePreview(dialogUi);
            });
            changeValueByArrowKey(dialogUi.leadingBodyInput, function () { updateTypescalePreview(dialogUi); });
            changeValueByArrowKey(dialogUi.leadingHeadingInput, function () { updateTypescalePreview(dialogUi); });
            dialogUi.baseInput.onChanging = function () {
                clearAllPreviewOverrides(dialogUi);
                updateTypescalePreview(dialogUi);
            };
            dialogUi.baseInput.onChange = function () {
                clearAllPreviewOverrides(dialogUi);
                updateTypescalePreview(dialogUi);
            };
            dialogUi.ratioDD.onChange = function () {
                clearAllPreviewOverrides(dialogUi);
                syncLevelRadiosWithScaleOption(dialogUi);
                updateTypescalePreview(dialogUi);
            };
            for (var levelRowEditIndex = 0; levelRowEditIndex < dialogUi.levelRows.length; levelRowEditIndex++) {
                bindPreviewCellEdit(dialogUi, dialogUi.levelRows[levelRowEditIndex]);
            }
            bindPreviewCellEdit(dialogUi, dialogUi.baseRow, function (parsedValue) {
                if (parsedValue !== null) {
                    dialogUi.baseInput.text = String(parsedValue);
                }
                clearAllPreviewOverrides(dialogUi);
                updateTypescalePreview(dialogUi);
            });
            bindPreviewCellEdit(dialogUi, dialogUi.captionRow);
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
            dialogUi.clearCacheButton.onClick = function () { clearFontCacheAndReload(dialogUi); };
        }

        bindTypescaleDialogEvents(dialogUi);
        updateFontStyleDropdowns(dialogUi);
        updateHeadingFontStyleDropdowns(dialogUi);
        syncLevelRadiosWithScaleOption(dialogUi);
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

    function setParagraphStyleProps(targetDocument, styleName, size, font, leading, spaceAfter, spaceBefore, kerningMethod, isHeading, silent) {
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
        if (typeof spaceBefore === "number" && spaceBefore >= 0) {
            style.spaceBefore = spaceBefore;
        }
        if (typeof spaceAfter === "number" && spaceAfter >= 0) {
            style.spaceAfter = spaceAfter;
        }
        // フォントによっては設定不可のため、安全に無視
        if (kerningMethod) {
            try { style.kerningMethod = kerningMethod; } catch (ke) { }
        }
        try {
            if (isHeading) {
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