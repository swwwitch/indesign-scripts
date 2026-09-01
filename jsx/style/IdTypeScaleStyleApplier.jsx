#target indesign

/*

### 概要

基準サイズとスケール倍率からタイプスケールを組み立て、本文・見出し・リスト・表の段落スタイルへ一括適用します。

詳細は README を参照してください。

### Overview

Builds a type scale from a base size and a scale ratio, then applies it to the body, heading, list and table paragraph styles at once.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTypeScaleStyleApplier";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-06-30";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTypeScaleStyleApplier.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTypeScaleStyleApplier.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n4f9b0666db66"; /* 紹介記事 / article URL */

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
 * @param {number} spacing 要素間隔。省略時は WINDOW_SPACING
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
 * @param {number} spacing 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 基準サイズとスケール / Base size and scale */
var DEFAULT_BASE_SIZE_PT  = 9;      /* 基準サイズ（本文, pt）/ base body size (pt) */
var DEFAULT_BASE_SIZE_Q   = 13;     /* 基準サイズ（本文, Q）/ base body size (Q) */
var DEFAULT_RATIO         = 1.309;  /* スケール倍率 / scale ratio */
var DEFAULT_LEVEL_COUNT   = 4;      /* 見出しレベル数 / number of heading levels */

/* 行送りとアキの既定値（%）/ Default leading and spacing (%) */
var DEFAULT_BODY_LEADING_PERCENT      = 160;  /* 本文の行送り / body leading */
var DEFAULT_HEADING_LEADING_PERCENT   = 115;  /* 見出しの行送り / heading leading */
var DEFAULT_SPACE_BEFORE_PERCENT      = 10;   /* 見出しの段落前のアキ / space before headings */
var DEFAULT_SPACE_AFTER_PERCENT       = 10;   /* 見出しの段落後のアキ / space after headings */
var DEFAULT_BODY_SPACE_BEFORE_PERCENT = 15;   /* 本文の段落前のアキ / space before body */
var DEFAULT_BODY_SPACE_AFTER_PERCENT  = 15;   /* 本文の段落後のアキ / space after body */

/* 本文サイズに対するリスト／表の文字サイズ比率（%）/ List and table size as a percentage of the body size */
var DEFAULT_LIST_SIZE_PERCENT  = 100;  /* ul-li / ol-li / List Paragraph */
var DEFAULT_TABLE_SIZE_PERCENT = 94;   /* td / th */

/* 段落前後のアキを段落スタイルへ適用するか / Whether to write spaceBefore and spaceAfter to the styles */
var ENABLE_SPACE_BEFORE = true;
var ENABLE_SPACE_AFTER  = true;

/* 「同じスタイルの段落間隔」を適用するか。スタイル名ごとのルールは resolveSameStyleSpacing を参照
   / Whether to write sameParaStyleSpacing; see resolveSameStyleSpacing for the per-style rules
     ul-li      → 0（箇条書きの項目どうしは詰める / bullet items are kept tight）
     p / ol-li  → 段落前のアキと同じ値 / same value as spaceBefore
     それ以外   → 変更しない / left unchanged */
var ENABLE_SAME_STYLE_SPACING = true;

/* 字揃えを強制適用するか。false なら元の字揃えを保持する
   / Whether to force the justification; false keeps the original justification */
var ENABLE_JUSTIFICATION = false;

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

/* 日英ラベル定義（カテゴリ別） / Japanese-English labels by category */
var LABELS = {
    dialog: {
        title: { ja: "タイプスケールで一括設定", en: "Type Scale Settings" }
    },
    undo: {
        applyTypeScale: { ja: "タイプスケールで一括設定", en: "Apply Type Scale" }
    },
    error: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        invalidBaseSize: {
            ja: "基準サイズ（本文）には正の数値を入力してください。",
            en: "Enter a positive number for Base Size (Body)."
        },
        missingStyle: {
            ja: "段落スタイル「%1」が見つかりません。",
            en: "Paragraph style \"%1\" was not found."
        }
    },
    progress: {
        title: { ja: "処理中", en: "Processing" },
        loadingFonts: {
            ja: "フォント情報を読み込んでいます…",
            en: "Loading font information..."
        },
        preparingDialog: {
            ja: "ダイアログを準備しています…",
            en: "Preparing dialog..."
        }
    },
    panel: {
        bodyText: { ja: "本文", en: "Body" },
        headingText: { ja: "見出し", en: "Headings" },
        fontAssignment: { ja: "フォント指定オプション", en: "Font Assignment" },
        scaleSettings: { ja: "スケール設定", en: "Scale Settings" },
        preview: { ja: "段落スタイルとサイズプレビュー", en: "Paragraph Styles & Size Preview" }
    },
    fontMode: {
        same: {
            ja: "本文と見出しで共通",
            en: "Use same font for body and headings"
        },
        separate: {
            ja: "本文と見出しで別々に指定",
            en: "Specify body and heading fonts separately"
        },
        none: { ja: "フォントを変更しない", en: "Do not change fonts" }
    },
    checkbox: {
        sizeOnly: { ja: "サイズのみ", en: "Size only" }
    },
    field: {
        baseSize: { ja: "基準サイズ", en: "Base Size (Body)" },
        font: { ja: "フォント", en: "Font" },
        fontStyle: { ja: "スタイル", en: "Style" },
        bodyLeading: { ja: "行送り", en: "Leading Ratio (Body)" },
        headingLeading: { ja: "行送り", en: "Leading Ratio (Headings)" },
        kerning: { ja: "カーニング", en: "Kerning Method" },
        scaleMethod: { ja: "スケール方式", en: "Scale Method" },
        headingLevels: { ja: "見出しレベル数", en: "Heading Levels" },
        sizeRounding: { ja: "サイズの丸め", en: "Size Rounding" }
    },
    kerning: {
        japaneseMono: { ja: "和文等幅", en: "Japanese Mono" },
        metrics: { ja: "メトリクス", en: "Metrics" },
        optical: { ja: "オプティカル", en: "Optical" }
    },
    rounding: {
        integer: { ja: "整数", en: "Integer" },
        firstDecimal: { ja: "小数点第1位", en: "1 decimal place" },
        secondDecimal: { ja: "小数点第2位", en: "2 decimal places" }
    },
    column: {
        level: { ja: "レベル", en: "Level" },
        size: { ja: "サイズ", en: "Size" },
        spaceBefore: { ja: "段落前アキ", en: "Space Before" },
        spaceAfter: { ja: "段落後のアキ", en: "Space After" },
        paragraphStyle: { ja: "段落スタイル", en: "Paragraph Style" },
        fontStyle: { ja: "ウエイト", en: "Weight" }
    },
    row: {
        levelPrefix: { ja: "レベル", en: "Level " },
        baseBody: { ja: "★基準（本文）", en: "★Base (Body)" },
        list: { ja: "リスト", en: "List" },
        table: { ja: "テーブル", en: "Table" },
        caption: { ja: "キャプション", en: "Caption" }
    },
    option: {
        noFontChange: { ja: "（変更しない）", en: "(No change)" },
        refBodyFont: { ja: "本文のフォントを参照", en: "Reference body font" },
        notAvailable: { ja: "—", en: "—" }
    },
    button: {
        includeFonts: {
            ja: "フォント、スタイルを含める",
            en: "Include fonts & styles"
        },
        cancel: { ja: "キャンセル", en: "Cancel" },
        ok: { ja: "OK", en: "OK" }
    },
    ratio: {
        minorSecond: { ja: "短2度", en: "Minor Second" },
        majorSecond: { ja: "長2度", en: "Major Second" },
        minorThird: { ja: "短3度", en: "Minor Third" },
        majorThird: { ja: "長3度", en: "Major Third" },
        goldenHalf: { ja: "黄金比：1/2", en: "Golden Ratio: ½" },
        perfectFourth: { ja: "完全4度", en: "Perfect Fourth" },
        augmentedFourth: { ja: "増4度", en: "Augmented Fourth" },
        golden: { ja: "黄金比", en: "Golden Ratio" },
        browserDefault: { ja: "ブラウザー既定", en: "Browser Default" }
    }
};

/* ドット区切りのキーで多言語ラベルを取得（例：getLabel("dialog.title")） / Resolve a dotted-path label key */
/**
 * ドット区切りキーでラベルを取得する
 * @param {string} labelKey 例: "dialog.title"
 * @returns {string} 現在の言語のラベル文字列。見つからない場合はキーをそのまま返す
 */
function getLabel(labelKey) {
    var keyParts = labelKey.split(".");
    var node = LABELS;
    for (var i = 0; i < keyParts.length; i++) {
        if (node && typeof node[keyParts[i]] !== "undefined") {
            node = node[keyParts[i]];
        } else {
            return labelKey;
        }
    }
    /* currentLang → en → キーの順でフォールバック / Fall back to en, then to the key itself */
    if (node) {
        if (node[currentLang]) return node[currentLang];
        if (node.en) return node.en;
    }
    return labelKey;
}

/* コロン付きラベル（日本語は全角、英語は半角） / Label with colon (full-width JA, half-width EN) */
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
    palette.margins = WINDOW_MARGINS;
    palette.spacing = 10;

    var messageText = palette.add("statictext", undefined, message);
    messageText.preferredSize.width = 260;

    var progressBar = palette.add("progressbar", undefined, 0, 100);
    progressBar.preferredSize.width = 260;
    progressBar.preferredSize.height = 6; // バーを細く
    progressBar.value = 20;

    // 後からメッセージ・進捗を更新できるよう参照を保持 / Keep references so the message/value can be updated later
    palette.messageText = messageText;
    palette.progressBar = progressBar;

    palette.show();
    try { palette.update(); } catch (e) { }
    return palette;
}

/**
 * 進捗パレットのメッセージを更新する
 * @param {Window} palette 対象のパレット
 * @param {string} message 表示するメッセージ
 * @param {number} value 進捗バーの値（0〜100）
 * @returns {void}
 */
function updateProgressPalette(palette, message, value) {
    if (!palette) return;
    try {
        if (message && palette.messageText) palette.messageText.text = message;
        if (typeof value === "number" && palette.progressBar) palette.progressBar.value = value;
        palette.update();
    } catch (e) { }
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

// =========================================
// 単位 / Units
// =========================================

/* ポイント換算係数（pt = 値 × factor） / Point conversion factor per unit */
var UNIT_POINT_CONVERTERS = [
    { name: "POINTS", label: "pt", factor: 1 },
    { name: "MILLIMETERS", label: "mm", factor: 2.834645669 },
    { name: "CENTIMETERS", label: "cm", factor: 28.34645669 },
    { name: "INCHES", label: "inch", factor: 72 },
    { name: "INCHES_DECIMAL", label: "inch", factor: 72 },
    { name: "PICAS", label: "pica", factor: 12 },
    { name: "CICEROS", label: "c", factor: 12.7896 },
    { name: "AGATES", label: "ag", factor: 1 },
    { name: "PIXELS", label: "px", factor: 1 },
    { name: "Q", label: "Q", factor: 0.708661417 }, /* 1Q = 0.25mm / 1Q = 0.25mm */
    { name: "HA", label: "H", factor: 0.708661417 }
];

/**
 * 単位の列挙値を数値として取得する
 * @param {MeasurementUnits} unitName 対象の単位
 * @returns {number} 単位を表す数値
 */
function getMeasurementUnitValue(unitName) {
    try {
        return MeasurementUnits[unitName];
    } catch (e) { }
    return null;
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
 * 単位の表示記号を返す
 * @param {MeasurementUnits} unit 対象の単位
 * @returns {string} 単位の記号
 */
function unitSymbol(unit) {
    for (var unitIndex = 0; unitIndex < UNIT_POINT_CONVERTERS.length; unitIndex++) {
        if (unit === getMeasurementUnitValue(UNIT_POINT_CONVERTERS[unitIndex].name)) {
            return UNIT_POINT_CONVERTERS[unitIndex].label;
        }
    }
    return "pt";
}

/**
 * 現在の単位の数値をポイントへ換算する
 * @param {number} value 現在の単位での数値
 * @param {MeasurementUnits} unit 対象の単位
 * @returns {number} ポイント値
 */
function toPoints(value, unit) {
    // 各種単位をポイントに変換（内部計算はpt基準）
    for (var unitIndex = 0; unitIndex < UNIT_POINT_CONVERTERS.length; unitIndex++) {
        if (unit === getMeasurementUnitValue(UNIT_POINT_CONVERTERS[unitIndex].name)) {
            return value * UNIT_POINT_CONVERTERS[unitIndex].factor;
        }
    }
    return value;
}

/**
 * ポイント値を現在の単位へ換算する
 * @param {number} valueInPoints ポイント値
 * @param {MeasurementUnits} unit 対象の単位
 * @returns {number} 現在の単位での数値
 */
function fromPoints(valueInPoints, unit) {
    // ポイントを表示単位へ逆変換（pt → 表示単位）
    for (var unitIndex = 0; unitIndex < UNIT_POINT_CONVERTERS.length; unitIndex++) {
        if (unit === getMeasurementUnitValue(UNIT_POINT_CONVERTERS[unitIndex].name)) {
            return valueInPoints / UNIT_POINT_CONVERTERS[unitIndex].factor;
        }
    }
    return valueInPoints;
}

/* フォント一覧のキャッシュ：app.fonts.everyItem().getElements() が遅いので、ファミリー名・スタイル名のテーブルをディスクに保存して再利用する */
var FONT_CACHE_VERSION = "v1";
var FONT_CACHE_FILENAME = "TypeScaleStyleApplier_fontcache.txt";

/**
 * フォント一覧キャッシュのファイルを取得する
 * @returns {File} キャッシュファイル
 */
function getFontCacheFile() {
    try {
        var dir = new Folder(Folder.userData.fsName + "/TypeScaleStyleApplier");
        if (!dir.exists) dir.create();
        return new File(dir.fsName + "/" + FONT_CACHE_FILENAME);
    } catch (e) {
        return null;
    }
}

/**
 * ディスクからフォント一覧キャッシュを読み込む
 * @param {number} currentFontCount 現在のフォント数。キャッシュの有効判定に使う
 * @returns {object|null} キャッシュ。読み込めない場合は null
 */
function loadFontCache(currentFontCount) {
    var file = getFontCacheFile();
    if (!file || !file.exists) return null;
    file.encoding = "UTF-8";
    if (!file.open("r")) return null;
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

/**
 * フォント一覧をキャッシュとして書き出す
 * @param {number} fontCount 保存時のフォント数
 * @param {Array<string>} families フォントファミリー名の配列
 * @param {object} fontMap ファミリー名をキーにしたスタイル名の配列
 * @returns {void}
 */
function saveFontCache(fontCount, families, fontMap) {
    var file = getFontCacheFile();
    if (!file) return false;
    file.encoding = "UTF-8";
    if (!file.open("w")) return false;
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

/* 用語 → ランクの照合テーブル。長い用語を先に照合したいので長さの降順に並べ、
   正規表現もここで 1 度だけ作る（フォント数ぶん呼ばれるため、都度の生成は避ける）
   / Term-to-rank table, longest first; the regexes are built once */
var WEIGHT_TERM_PATTERNS = (function () {
    var terms = [];
    for (var groupIndex = 0; groupIndex < WEIGHT_GROUPS.length; groupIndex++) {
        for (var termIndex = 0; termIndex < WEIGHT_GROUPS[groupIndex].length; termIndex++) {
            terms.push({ term: WEIGHT_GROUPS[groupIndex][termIndex], index: groupIndex });
        }
    }
    terms.sort(function (a, b) { return b.term.length - a.term.length; });
    for (var patternIndex = 0; patternIndex < terms.length; patternIndex++) {
        terms[patternIndex].pattern = new RegExp(
            "\\b" + terms[patternIndex].term.replace(/[-\/\\^$*+?.()|[\]{}]/g, "\\$&") + "\\b"
        );
    }
    return terms;
})();

/**
 * フォントスタイル名から太さの基準ランクを求める
 * @param {string} style 小文字化・正規化済みのフォントスタイル名
 * @param {string} family フォントファミリー名
 * @returns {number} 基準ランク
 */
function getStyleWeightBaseRank(style, family) {
    /* 空白を除いてから判定する（"Helvetica Neue" → "helveticaneue"）/ Strip spaces before matching */
    var familyLower = (family || "").toLowerCase().replace(/\s+/g, "");
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

    for (var patternIndex = 0; patternIndex < WEIGHT_TERM_PATTERNS.length; patternIndex++) {
        if (WEIGHT_TERM_PATTERNS[patternIndex].pattern.test(style)) {
            return 1000 + WEIGHT_TERM_PATTERNS[patternIndex].index;
        }
    }

    return 1000 + WEIGHT_REGULAR_INDEX;
}

/**
 * フォントスタイル名から並べ替え用の太さランクを求める
 * @param {string} styleName フォントスタイル名
 * @param {string} familyName フォントファミリー名
 * @returns {number} 太さランク
 */
function getStyleWeightRank(styleName, familyName) {
    var rawStyleName = (styleName || "").toString();
    var style = rawStyleName.toLowerCase().replace(/[_\-]+/g, " ").replace(/^\s+|\s+$/g, "");
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
    /* hasSemiCondensed をここにも含めるのは、1 語の "semicondensed"（hasCondensed が立たない）を
       2 語の "semi condensed"（hasCondensed も立つ）と同じランクに揃えるため。二重加算ではない
       / semiCondensed is listed here so the one-word and two-word spellings land on the same rank */
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
        alert(getLabel("error.noDocument"));
        return;
    }

    var targetDocument = app.activeDocument;
    var unit = getTextSizeUnit();

    // ライブプレビュー／適用で書き換えた段落スタイル名（キャンセル時の復元対象）
    // showTypescaleDialog 内のライブプレビューより前に初期化しておく
    var _previewModifiedStyles = {};

    // 基準サイズの初期値は段落スタイル "p" → "Normal" の本文サイズを参照する
    var defaultBaseSize = getBodyStyleBaseSize(targetDocument, unit);
    if (defaultBaseSize === null) {
        // "p" / "Normal" が無ければ単位ごとの既定値にフォールバック
        defaultBaseSize = DEFAULT_BASE_SIZE_PT;
        try {
            if (unit === MeasurementUnits.Q) {
                defaultBaseSize = DEFAULT_BASE_SIZE_Q;
            } else if (unit === MeasurementUnits.MILLIMETERS) {
                defaultBaseSize = DEFAULT_BASE_SIZE_Q * 0.25; // Q→mm
            }
        } catch (e) { }
    }

    // 段落スタイル "p" → "Normal" の本文サイズを表示単位で返す（無ければ null）
    /**
     * 本文スタイルの基準サイズを取得する
     * @param {Document} targetDocument 対象ドキュメント
     * @param {MeasurementUnits} unit 表示単位
     * @returns {number|null} 基準サイズ。取得できない場合は null
     */
    function getBodyStyleBaseSize(targetDocument, unit) {
        var candidateNames = ["p", "Normal"];
        for (var nameIndex = 0; nameIndex < candidateNames.length; nameIndex++) {
            var style = findParagraphStyle(targetDocument, candidateNames[nameIndex]);
            if (style === null) continue;
            var sizePt = null;
            try { sizePt = style.pointSize; } catch (eSize) { sizePt = null; }
            if (typeof sizePt !== "number" || isNaN(sizePt) || sizePt <= 0) continue;
            return roundTo(fromPoints(sizePt, unit), 2);
        }
        return null;
    }

    var typescaleSettings = showTypescaleDialog(targetDocument, defaultBaseSize, DEFAULT_RATIO, DEFAULT_LEVEL_COUNT, unit);

    if (typescaleSettings !== null) {
        // ライブプレビューは doScript の外でスタイルを直接書き換えているため、
        // そのままだと doScript のステップを取り消してもプレビュー時の値が残る。
        // 最終適用の前に起動時の状態へ戻し、doScript の 1 ステップで起動時 → 適用後にする
        restoreParagraphStyleProps(targetDocument, typescaleSettings.originalStyleProps, _previewModifiedStyles);
        // OK 時の最終適用は 1 つの undo ステップにまとめる（Edit > 取り消し で一括に戻せる）
        app.doScript(function () {
            applyTypescaleSettings(targetDocument, typescaleSettings, false, unit);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.applyTypeScale"));
    }

    /**
     * 基準サイズとスケール倍率から各レベルのサイズを計算する
     * @param {number} base 基準サイズ（本文）
     * @param {number} ratio スケール倍率
     * @param {number} levelCount 見出しレベル数
     * @param {Array<number>} [multipliers] レベルごとの倍率。省略時は ratio の累乗で算出
     * @param {number} [captionMultiplier] キャプションの倍率
     * @returns {object} headingSizes・base・caption を持つオブジェクト
     */
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

    /**
     * 確定したタイプスケール設定を段落スタイルへ適用する
     * @param {Document} targetDocument 対象ドキュメント
     * @param {object} typescaleSettings ダイアログで確定した設定
     * @param {boolean} silent スタイル未検出時に警告を出さないなら true
     * @param {MeasurementUnits} unit 表示単位
     * @returns {void}
     */
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
        /**
         * 1 つの段落スタイルへサイズ・行送り・アキなどを適用する
         * @param {object} rowSettings 適用する行の設定
         * @param {string} rowSettings.styleName 対象の段落スタイル名
         * @param {number} rowSettings.sizeInUnit 表示単位での文字サイズ
         * @param {number} rowSettings.leadingMult 文字サイズに対する行送りの倍率
         * @param {boolean} rowSettings.isHeading 見出しなら true
         * @param {string} [rowSettings.fontFamilyName] フォントファミリー名
         * @param {string} [rowSettings.fontStyleName] フォントスタイル名
         * @param {number} [rowSettings.sizeOverride] 表示単位での文字サイズの個別指定
         * @param {number} [rowSettings.spaceBeforeOverride] 表示単位での段落前アキの個別指定
         * @param {number} [rowSettings.spaceAfterOverride] 表示単位での段落後アキの個別指定
         * @returns {void}
         */
        function applyParagraphStyleSettings(rowSettings) {
            var styleName = rowSettings.styleName;
            if (!styleName) return;
            var isHeading = !!rowSettings.isHeading;
            var roundDigits = typescaleSettings.roundDigits;

            var effectiveSize = (typeof rowSettings.sizeOverride === "number") ? rowSettings.sizeOverride : rowSettings.sizeInUnit;
            var sizePt = toPoints(roundTo(effectiveSize, roundDigits), unit);

            /* 段落前後のアキ：個別指定があればその値、なければ文字サイズに対する割合
               / Space before/after: the per-row override if given, otherwise a ratio of the size */
            var spaceBeforePt = (typeof rowSettings.spaceBeforeOverride === "number")
                ? toPoints(roundTo(rowSettings.spaceBeforeOverride, roundDigits), unit)
                : sizePt * (isHeading ? headingSpaceBeforePercent : bodySpaceBeforePercent) / 100;
            var spaceAfterPt = (typeof rowSettings.spaceAfterOverride === "number")
                ? toPoints(roundTo(rowSettings.spaceAfterOverride, roundDigits), unit)
                : sizePt * (isHeading ? headingSpaceAfterPercent : bodySpaceAfterPercent) / 100;

            // フォントファミリー＋スタイルで解決。未指定時はファミリー内の推奨スタイルを使用
            var fontToUse = null;
            if (rowSettings.fontFamilyName) {
                fontToUse = rowSettings.fontStyleName
                    ? findFontByFamilyAndStyle(rowSettings.fontFamilyName, rowSettings.fontStyleName)
                    : findFontInFamily(rowSettings.fontFamilyName);
            }

            setParagraphStyleProps(targetDocument, styleName, {
                size: sizePt,
                font: fontToUse,
                leading: (typeof rowSettings.leadingMult === "number") ? sizePt * rowSettings.leadingMult : null,
                spaceBefore: spaceBeforePt,
                spaceAfter: spaceAfterPt,
                kerningMethod: isHeading ? typescaleSettings.headingKerningMethod : typescaleSettings.bodyKerningMethod,
                isHeading: isHeading,
                silent: silent,
                sizeOnly: !!typescaleSettings.sizeOnly,
                originalProps: (typescaleSettings.originalStyleProps && typescaleSettings.originalStyleProps[styleName]) || null
            });
        }
        applyParagraphStyleSettings({
            styleName: typescaleSettings.baseStyleName,
            sizeInUnit: computedSizes.base,
            leadingMult: typescaleSettings.bodyLeading,
            isHeading: false,
            fontFamilyName: typescaleSettings.fontFamily,
            fontStyleName: typescaleSettings.baseFontStyleName,
            sizeOverride: typescaleSettings.baseSizeOverride,
            spaceBeforeOverride: typescaleSettings.baseSpaceBeforeOverride,
            spaceAfterOverride: typescaleSettings.baseSpaceAfterOverride
        });
        applyParagraphStyleSettings({
            styleName: typescaleSettings.captionStyleName,
            sizeInUnit: computedSizes.caption,
            leadingMult: typescaleSettings.bodyLeading,
            isHeading: false,
            fontFamilyName: typescaleSettings.fontFamily,
            fontStyleName: typescaleSettings.captionFontStyleName,
            sizeOverride: typescaleSettings.captionSizeOverride,
            spaceBeforeOverride: typescaleSettings.captionSpaceBeforeOverride,
            spaceAfterOverride: typescaleSettings.captionSpaceAfterOverride
        });
        // 本文派生行（リスト・テーブル）: 文字サイズは基準（本文）×係数。行送り・カーニング・フォントは本文と同じ
        var bodyDerivedRows = typescaleSettings.bodyDerivedRows || [];
        for (var bodyRowIndex = 0; bodyRowIndex < bodyDerivedRows.length; bodyRowIndex++) {
            var bodyDerivedRow = bodyDerivedRows[bodyRowIndex];
            applyParagraphStyleSettings({
                styleName: bodyDerivedRow.styleName,
                sizeInUnit: computedSizes.base * bodyDerivedRow.sizeFactor,
                leadingMult: typescaleSettings.bodyLeading,
                isHeading: false,
                fontFamilyName: typescaleSettings.fontFamily,
                fontStyleName: bodyDerivedRow.fontStyleName,
                sizeOverride: bodyDerivedRow.sizeOverride,
                spaceBeforeOverride: bodyDerivedRow.spaceBeforeOverride,
                spaceAfterOverride: bodyDerivedRow.spaceAfterOverride
            });
        }
        for (var levelNumber = 1; levelNumber <= typescaleSettings.levelCount; levelNumber++) {
            var levelStyleName = typescaleSettings.levelStyleNames && typescaleSettings.levelStyleNames[levelNumber - 1];
            var fontStyleName = typescaleSettings.levelFontStyleNames && typescaleSettings.levelFontStyleNames[levelNumber - 1];
            var levelSizeOverride = typescaleSettings.levelSizeOverrides ? typescaleSettings.levelSizeOverrides[levelNumber - 1] : null;
            var levelSpaceBeforeOverride = typescaleSettings.levelSpaceBeforeOverrides ? typescaleSettings.levelSpaceBeforeOverrides[levelNumber - 1] : null;
            var levelSpaceAfterOverride = typescaleSettings.levelSpaceAfterOverrides ? typescaleSettings.levelSpaceAfterOverrides[levelNumber - 1] : null;
            applyParagraphStyleSettings({
                styleName: levelStyleName,
                sizeInUnit: computedSizes.headingSizes[levelNumber - 1],
                leadingMult: typescaleSettings.headingLeading,
                isHeading: true,
                fontFamilyName: typescaleSettings.headingFontFamily,
                fontStyleName: fontStyleName,
                sizeOverride: levelSizeOverride,
                spaceBeforeOverride: levelSpaceBeforeOverride,
                spaceAfterOverride: levelSpaceAfterOverride
            });
        }
    }

    /**
     * 指定桁数で丸める
     * @param {number} num 対象の数値
     * @param {number} places 小数点以下の桁数
     * @returns {number} 丸めた数値
     */
    function roundTo(num, places) {
        var factor = Math.pow(10, places);
        return Math.round(num * factor) / factor;
    }

    /**
     * 配列に値が含まれるかを判定する
     * @param {Array} array 対象の配列
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
     * @param {Document} targetDocument 対象ドキュメント
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

    /**
     * フォントファミリー名の一覧を取得する
     * @returns {Array<string>} ファミリー名の配列
     */
    function getFontFamilyNames() {
        return getFontInfo().families.slice(0);
    }

    /**
     * ファミリー内で優先して選ぶスタイル名を求める
     * @param {string} styleNames フォントファミリー名
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

    /**
     * 表示名でドロップダウンの項目を選択する
     * @param {DropDownList} dropdownList 対象のドロップダウン
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
     * @param {DropDownList} dropdownList 対象のドロップダウン
     * @returns {string} 選択中の表示名
     */
    function getDropdownText(dropdownList) {
        return dropdownList.selection ? dropdownList.selection.text : null;
    }

    /**
     * タイプスケール設定ダイアログを表示する
     * @param {Document} targetDocument 対象ドキュメント
     * @param {number} defaultBase 基準サイズの初期値
     * @param {number} defaultRatio スケール倍率の初期値
     * @param {number} defaultLevelCount 見出しレベル数の初期値
     * @param {MeasurementUnits} unit 表示単位
     * @returns {object|null} 設定内容。キャンセル時は null
     */
    function showTypescaleDialog(targetDocument, defaultBase, defaultRatio, defaultLevelCount, unit) {
        var unitSym = unitSymbol(unit);
        var styleNames = getParagraphStyleNames(targetDocument);
        // ダイアログ表示中のライブプレビューが値を書き換えるため、元値を退避する（取得はダイアログ構築直前に実行）
        var originalStyleProps = null;
        var _fontSectionActive = false;  // 「フォント、スタイルを含める」でフォント・スタイル指定を有効化したか

        // フォント一覧は UI（パレット／ダイアログ Window）を組み立てる前に読み込む。
        // ダイアログ Window 生成後やモーダル表示中に重いフォント列挙を行うと InDesign が落ちるため、
        // 「パレット表示 → 読み込み → … → ダイアログ表示直前にパレットを閉じる」の順を厳守する。
        // パレットはフォント読み込み後も閉じず、ダイアログ構築の間も表示し続けて空白時間をなくす。
        var loadingPalette = createProgressPalette(getLabel("progress.loadingFonts"));
        var fontFamilies = getFontFamilyNames();
        updateProgressPalette(loadingPalette, getLabel("progress.preparingDialog"), 60);
        var fontOptions = [getLabel("option.noFontChange")].concat(fontFamilies);
        var roundOptions = [
            { label: getLabel("rounding.integer"), digits: 0 },
            { label: getLabel("rounding.firstDecimal"), digits: 1 },
            { label: getLabel("rounding.secondDecimal"), digits: 2 }
        ];
        var defaultRoundDigits = 1;
        var kerningOptions = [
            { label: getLabel("kerning.japaneseMono"), value: "和文等幅" },
            { label: getLabel("kerning.metrics"), value: "メトリクス" },
            { label: getLabel("kerning.optical"), value: "オプティカル" }
        ];


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
         * @param {object} panel 追加先のコンテナ
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
         * @param {function} onValueChange 値の変更後に呼ぶ処理
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

        var scaleOptions = [
            { type: "scale", key: "ratio.minorSecond", ratio: 1.067 },
            { type: "scale", key: "ratio.majorSecond", ratio: 1.125 },
            { type: "scale", key: "ratio.minorThird", ratio: 1.2 },
            { type: "scale", key: "ratio.majorThird", ratio: 1.25 },
            { type: "scale", key: "ratio.goldenHalf", ratio: 1.309 },
            { type: "scale", key: "ratio.perfectFourth", ratio: 1.333 },
            { type: "scale", key: "ratio.augmentedFourth", ratio: 1.414 },
            { type: "scale", key: "ratio.golden", ratio: 1.618 },
            {
                type: "preset",
                key: "ratio.browserDefault",
                multipliers: [2.00, 1.50, 1.17, 1.00],
                captionMultiplier: 0.83,
                forcedLevelCount: 4
            }
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
         * @param {DropDownList} dropdownList 対象のドロップダウン
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
         * @param {number} labelWidth ラベルの幅（px）
         * @returns {object} パネル内のコントロール
         */
        function createBodyTextPanel(parent, labelWidth) {
            var bodyPanel = parent.add("panel", undefined, getLabel("panel.bodyText"));
            setupPanel(bodyPanel, 4);
            bodyPanel.alignment = ["fill", "top"];

            var fontGrp = addLabeledGroup(bodyPanel, getLabelWithColon("field.font"), labelWidth);
            var fontDD = fontGrp.add("dropdownlist", undefined, fontOptions);
            fontDD.preferredSize.width = 180;
            fontDD.selection = 0;

            var fontStyleGrp = addLabeledGroup(bodyPanel, getLabelWithColon("field.fontStyle"), labelWidth);
            var fontStyleDD = fontStyleGrp.add("dropdownlist", undefined, [getLabel("option.noFontChange")]);
            fontStyleDD.preferredSize.width = 130;
            fontStyleDD.selection = 0;

            var leadingBodyGrp = addLabeledGroup(bodyPanel, getLabelWithColon("field.bodyLeading"), labelWidth);
            var leadingBodyInput = leadingBodyGrp.add("edittext", undefined, String(DEFAULT_BODY_LEADING_PERCENT));
            leadingBodyInput.characters = 4;
            leadingBodyGrp.add("statictext", undefined, "%");

            var bodyKerningGrp = addLabeledGroup(bodyPanel, getLabelWithColon("field.kerning"), labelWidth);
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
         * @param {number} labelWidth ラベルの幅（px）
         * @returns {object} パネル内のコントロール
         */
        function createHeadingTextPanel(parent, labelWidth) {
            var headingPanel = parent.add("panel", undefined, getLabel("panel.headingText"));
            setupPanel(headingPanel, 4);
            headingPanel.alignment = ["fill", "top"];

            // Font controls (like body panel)
            var headingFontGrp = addLabeledGroup(headingPanel, getLabelWithColon("field.font"), labelWidth);
            var headingFontOptions = [getLabel("option.refBodyFont")].concat(fontFamilies);
            var headingFontDD = headingFontGrp.add("dropdownlist", undefined, headingFontOptions);
            headingFontDD.preferredSize.width = 180;
            headingFontDD.selection = 0;

            var headingFontStyleGrp = addLabeledGroup(headingPanel, getLabelWithColon("field.fontStyle"), labelWidth);
            var headingFontStyleDD = headingFontStyleGrp.add("dropdownlist", undefined, [getLabel("option.noFontChange")]);
            headingFontStyleDD.preferredSize.width = 130;
            headingFontStyleDD.selection = 0;

            var leadingHeadingGrp = addLabeledGroup(headingPanel, getLabelWithColon("field.headingLeading"), labelWidth);
            var leadingHeadingInput = leadingHeadingGrp.add("edittext", undefined, String(DEFAULT_HEADING_LEADING_PERCENT));
            leadingHeadingInput.characters = 4;
            leadingHeadingGrp.add("statictext", undefined, "%");

            var headingKerningGrp = addLabeledGroup(headingPanel, getLabelWithColon("field.kerning"), labelWidth);
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

        /**
         * 本文・見出しをまとめたテキスト設定パネルを組み立てる
         * @param {Window} dialog 対象のダイアログ
         * @returns {object} パネル内のコントロール
         */
        function createTextSettingsPanel(dialog) {
            var textSettingsGroup = dialog.add("group");
            textSettingsGroup.orientation = "column";
            textSettingsGroup.alignChildren = "fill";
            textSettingsGroup.alignment = "fill";
            textSettingsGroup.spacing = 6;

            var BODY_LABEL_WIDTH = 80;
            var HEADING_LABEL_WIDTH = 94;

            var textColumnGroup = textSettingsGroup.add("group");
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

        /**
         * フォント指定オプションのパネルを組み立てる
         * @param {Window} dialog 対象のダイアログ
         * @returns {object} パネル内のコントロール
         */
        function createFontSettingsPanel(dialog) {
            var fontPanel = dialog.add("panel", undefined, getLabel("panel.fontAssignment"));
            fontPanel.alignment = ["fill", "top"];
            setupPanel(fontPanel, 6);

            var fontModeGroup = fontPanel.add("group");
            fontModeGroup.orientation = "column";
            fontModeGroup.alignChildren = ["left", "center"];

            var useSameFontRadio = fontModeGroup.add("radiobutton", undefined, getLabel("fontMode.same"));
            var separateFontRadio = fontModeGroup.add("radiobutton", undefined, getLabel("fontMode.separate"));
            var disableFontRadio = fontModeGroup.add("radiobutton", undefined, getLabel("fontMode.none"));

            // 初期状態はフォント・スタイル設定を不可（フォントを読み込まないため）
            useSameFontRadio.value = false;
            separateFontRadio.value = false;
            disableFontRadio.value = true;

            // 「サイズのみ」: ON にするとサイズだけ更新し、フォント／行送り／アキ等は元の値を保持する
            // （保持の実体は setParagraphStyleProps の sizeOnly 分岐。ON 時にフォント指定を「変更しない」へ切り替えるのは onClick ハンドラ側）
            // 既定 ON：初回は非破壊的にサイズだけ適用する（フォント指定の既定 disable とも整合）
            var sizeOnlyCheckbox = fontPanel.add("checkbox", undefined, getLabel("checkbox.sizeOnly"));
            sizeOnlyCheckbox.value = true;

            return {
                disableFontRadio: disableFontRadio,
                useSameFontRadio: useSameFontRadio,
                separateFontRadio: separateFontRadio,
                sizeOnlyCheckbox: sizeOnlyCheckbox
            };
        }

        /**
         * スケール設定パネルを組み立てる
         * @param {Window} dialog 対象のダイアログ
         * @returns {object} パネル内のコントロール
         */
        function createScaleSettingsPanel(dialog) {
            var scaleSettingsPanel = dialog.add("panel", undefined, getLabel("panel.scaleSettings"));
            scaleSettingsPanel.alignment = ["fill", "top"];
            setupPanel(scaleSettingsPanel, 6);

            var OPTIONS_LABEL_WIDTH = 110;

            var baseGrp = addLabeledGroup(scaleSettingsPanel, getLabelWithColon("field.baseSize"), OPTIONS_LABEL_WIDTH);
            var baseInput = baseGrp.add("edittext", undefined, String(defaultBase));
            baseInput.characters = 4;
            baseGrp.add("statictext", undefined, unitSym);

            var scaleGrp = addLabeledGroup(scaleSettingsPanel, getLabelWithColon("field.scaleMethod"), OPTIONS_LABEL_WIDTH);
            var scaleLabels = [];
            for (var scaleOptionIndex = 0; scaleOptionIndex < scaleOptions.length; scaleOptionIndex++) {
                var scaleLabel = getLabel(scaleOptions[scaleOptionIndex].key);
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

            var levelGrp = addLabeledGroup(scaleSettingsPanel, getLabelWithColon("field.headingLevels"), OPTIONS_LABEL_WIDTH);
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

            var roundGrp = addLabeledGroup(scaleSettingsPanel, getLabelWithColon("field.sizeRounding"), OPTIONS_LABEL_WIDTH);
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

        /**
         * プレビューパネルを組み立てる
         * @param {Window} dialog 対象のダイアログ
         * @returns {object} パネル内のコントロール
         */
        function createPreviewPanel(dialog) {
            var previewPanel = dialog.add("panel", undefined, getLabel("panel.preview"));
            setupPanel(previewPanel, 2);

            var PREVIEW_LABEL_WIDTH = 96; // 1列目（レベル名）を約1文字分広げる（★基準（本文）などに対応）
            var PREVIEW_SIZE_WIDTH = 80;
            var PREVIEW_SPACE_BEFORE_WIDTH = 80;
            var PREVIEW_SPACE_AFTER_WIDTH = 80;

            var headerRow = previewPanel.add("group");
            headerRow.orientation = "row";
            headerRow.alignChildren = "left";
            headerRow.add("statictext", undefined, getLabelWithColon("column.level")).preferredSize.width = PREVIEW_LABEL_WIDTH;
            headerRow.add("statictext", undefined, getLabelWithColon("column.size")).preferredSize.width = PREVIEW_SIZE_WIDTH;
            var spaceBeforeHeader = headerRow.add("statictext", undefined, getLabelWithColon("column.spaceBefore"));
            spaceBeforeHeader.preferredSize.width = PREVIEW_SPACE_BEFORE_WIDTH;
            var spaceAfterHeader = headerRow.add("statictext", undefined, getLabelWithColon("column.spaceAfter"));
            spaceAfterHeader.preferredSize.width = PREVIEW_SPACE_AFTER_WIDTH;
            headerRow.add("statictext", undefined, getLabelWithColon("column.paragraphStyle")).preferredSize.width = 100;
            var fontStyleHeader = headerRow.add("statictext", undefined, getLabelWithColon("column.fontStyle"));
            fontStyleHeader.preferredSize.width = 100;
            fontStyleHeader.enabled = true;

            var previewHeaderSpacer = previewPanel.add("group");
            previewHeaderSpacer.preferredSize.height = 4;

            /**
             * 候補名を順に試してドロップダウンを選択する
             * @param {DropDownList} dropdownList 対象のドロップダウン
             * @param {Array<string>} candidates 候補となる項目名
             * @returns {boolean} 選択できたら true
             */
            function selectDropdownByCandidates(dropdownList, candidates) {
                for (var candidateIndex = 0; candidateIndex < candidates.length; candidateIndex++) {
                    for (var itemIndex = 0; itemIndex < dropdownList.items.length; itemIndex++) {
                        if (dropdownList.items[itemIndex].text === candidates[candidateIndex]) {
                            dropdownList.selection = itemIndex;
                            return true;
                        }
                    }
                }
                // 候補が見つからない場合、ルート／組み込みスタイル（[...]）への誤適用（error 516）を避けるため
                // 先頭の通常スタイルを選ぶ。通常スタイルが無ければ未選択のままにする。
                for (var fallbackIndex = 0; fallbackIndex < dropdownList.items.length; fallbackIndex++) {
                    if (dropdownList.items[fallbackIndex].text.charAt(0) !== "[") {
                        dropdownList.selection = fallbackIndex;
                        return false;
                    }
                }
                dropdownList.selection = null;
                return false;
            }

            /**
             * プレビューの 1 行を組み立てる
             * @param {object} parent 追加先のコンテナ
             * @param {string} label 行ラベル
             * @param {string|Array<string>} defaultStyleNames 既定で選ぶ段落スタイル名の候補
             * @param {boolean} isHeading 見出し行なら true
             * @returns {object} 行のコントロール
             */
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
                var candidateList = (typeof defaultStyleNames === "string") ? [defaultStyleNames] : defaultStyleNames;
                selectDropdownByCandidates(styleDD, candidateList);

                var fontStyleDD = row.add("dropdownlist", undefined, [getLabel("option.noFontChange")]);
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
                levelRows.push(createPreviewRow(previewPanel, getLabel("row.levelPrefix") + levelNumber, levelCandidates, true));
            }

            // 基準（本文）の下に並ぶ、本文サイズ×係数で算出する行（係数はユーザー設定の比率(%)から）
            var baseRow = createPreviewRow(previewPanel, getLabel("row.baseBody"), ["p", "Normal"], false);
            var listRow = createPreviewRow(previewPanel, getLabel("row.list"), ["ul-li", "ol-li", "List Paragraph"], false);
            listRow.bodySizeFactor = DEFAULT_LIST_SIZE_PERCENT / 100;
            var tableRow = createPreviewRow(previewPanel, getLabel("row.table"), ["td", "th"], false);
            tableRow.bodySizeFactor = DEFAULT_TABLE_SIZE_PERCENT / 100;
            var captionRow = createPreviewRow(previewPanel, getLabel("row.caption"), "p.caption", false);

            return {
                levelRows: levelRows,
                baseRow: baseRow,
                bodyDerivedRows: [listRow, tableRow],
                captionRow: captionRow,
                spaceBeforeHeader: spaceBeforeHeader,
                spaceAfterHeader: spaceAfterHeader,
                fontStyleHeader: fontStyleHeader
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

            var includeFontsButton = leftButtonColumn.add("button", undefined, getLabel("button.includeFonts"));

            var centerButtonColumn = bottomRow.add("group");
            centerButtonColumn.alignment = ["fill", "fill"];
            centerButtonColumn.minimumSize.width = 0;
            // 「フォント、スタイルを含める」適用中の進捗バー。
            // モーダル表示中は別ウィンドウ（パレット）を出すと落ちるため、ダイアログ内に置く
            var applyProgressBar = centerButtonColumn.add("progressbar", undefined, 0, 100);
            applyProgressBar.alignment = ["fill", "center"];
            applyProgressBar.preferredSize.height = 6;
            applyProgressBar.visible = false;

            var rightButtonColumn = bottomRow.add("group");
            rightButtonColumn.orientation = "row";
            rightButtonColumn.alignChildren = ["right", "center"];
            rightButtonColumn.alignment = ["right", "center"];

            rightButtonColumn.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
            rightButtonColumn.add("button", undefined, getLabel("button.ok"), { name: "ok" });

            return { includeFontsButton: includeFontsButton, applyProgressBar: applyProgressBar };
        }

        /**
         * タイプスケールダイアログ全体を組み立てる
         * @returns {object} ダイアログとコントロール
         */
        function createTypescaleDialog() {
            var dialogWindow = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
            setupWindow(dialogWindow, 10);

            var textSettingsUi = createTextSettingsPanel(dialogWindow);

            var optionColumnGroup = dialogWindow.add("group");
            optionColumnGroup.orientation = "row";
            optionColumnGroup.alignChildren = ["fill", "top"];
            optionColumnGroup.alignment = "fill";
            optionColumnGroup.spacing = 10;

            var scaleSettingsUi = createScaleSettingsPanel(optionColumnGroup);
            var fontSettingsUi = createFontSettingsPanel(optionColumnGroup);
            var previewUi = createPreviewPanel(dialogWindow);

            var buttonUi = createButtonRow(dialogWindow);

            return {
                dialog: dialogWindow,
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
                bodyDerivedRows: previewUi.bodyDerivedRows,
                captionRow: previewUi.captionRow,
                spaceBeforeHeader: previewUi.spaceBeforeHeader,
                spaceAfterHeader: previewUi.spaceAfterHeader,
                fontStyleHeader: previewUi.fontStyleHeader,
                includeFontsButton: buttonUi.includeFontsButton,
                applyProgressBar: buttonUi.applyProgressBar,
                disableFontRadio: fontSettingsUi.disableFontRadio,
                useSameFontRadio: fontSettingsUi.useSameFontRadio,
                separateFontRadio: fontSettingsUi.separateFontRadio,
                sizeOnlyCheckbox: fontSettingsUi.sizeOnlyCheckbox
            };
        }

        var dialogUi = createTypescaleDialog();

        /**
         * 正の数値として解釈する
         * @param {string} text 入力文字列
         * @param {number} fallbackValue 解釈できない場合の値
         * @returns {number} 数値
         */
        function parsePositiveNumber(text, fallbackValue) {
            var value = parseFloat(text);
            return (isNaN(value) || value <= 0) ? fallbackValue : value;
        }

        /**
         * 0 以上の数値として解釈する
         * @param {string} text 入力文字列
         * @param {number} fallbackValue 解釈できない場合の値
         * @returns {number} 数値
         */
        function parseNonNegativeNumber(text, fallbackValue) {
            var value = parseFloat(text);
            return (isNaN(value) || value < 0) ? fallbackValue : value;
        }

        /**
         * パーセント入力を倍率として解釈する
         * @param {string} text 入力文字列
         * @param {number} fallbackValue 解釈できない場合の倍率
         * @returns {number} 倍率
         */
        function parsePositivePercentMultiplier(text, fallbackValue) {
            var value = parseFloat(text);
            return (isNaN(value) || value <= 0) ? fallbackValue : value / 100;
        }

        /**
         * ラジオボタン群で選択中の値を取得する
         * @param {Array} radioButtons 対象のラジオボタン
         * @param {Array} options 対応する選択肢
         * @param {string} [valueKey] 選択肢から取り出すキー。省略時は選択肢そのもの
         * @param {*} fallbackValue どれも選ばれていない場合の値
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
         * @param {DropDownList} dropdownList 対象のドロップダウン
         * @param {Array} options 対応する選択肢
         * @param {string} [valueKey] 選択肢から取り出すキー。省略時は選択肢そのもの
         * @param {*} fallbackValue 未選択の場合の値
         * @returns {string|null} 選択中の値
         */
        function getSelectedDropdownOptionValue(dropdownList, options, valueKey, fallbackValue) {
            if (!dropdownList.selection) return fallbackValue;
            return valueKey ? options[dropdownList.selection.index][valueKey] : options[dropdownList.selection.index];
        }

        /**
         * 現在のスケール倍率を取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {number} スケール倍率
         */
        function getCurrentScaleRatio(dialogUi) {
            if (!dialogUi.scaleDD.selection) return defaultRatio;
            var scaleOption = scaleOptions[dialogUi.scaleDD.selection.index];
            return (scaleOption && scaleOption.type === "scale") ? scaleOption.ratio : defaultRatio;
        }

        /**
         * 現在選択中のスケールプリセットを取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {string} プリセットを表す識別子
         */
        function getCurrentScaleOption(dialogUi) {
            if (!dialogUi.scaleDD.selection) return null;
            return scaleOptions[dialogUi.scaleDD.selection.index];
        }

        /**
         * 現在の見出しレベル数を取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {number} 見出しレベル数
         */
        function getCurrentLevelCount(dialogUi) {
            var scaleOption = getCurrentScaleOption(dialogUi);
            if (scaleOption && scaleOption.type === "preset" && typeof scaleOption.forcedLevelCount === "number") {
                return scaleOption.forcedLevelCount;
            }
            return getSelectedRadioValue(dialogUi.levelRadios, levelOptions, null, defaultLevelCount);
        }

        /**
         * 現在のサイズ丸め桁数を取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {number} 丸め桁数
         */
        function getCurrentRoundDigits(dialogUi) {
            return getSelectedRadioValue(dialogUi.roundRadios, roundOptions, "digits", defaultRoundDigits);
        }

        /**
         * 本文で選択中のフォントファミリーを取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {string} フォントファミリー名
         */
        function getSelectedFontFamily(dialogUi) {
            if (dialogUi.disableFontRadio && dialogUi.disableFontRadio.value) return null;
            if (!dialogUi.fontDD.selection || dialogUi.fontDD.selection.index === 0) return null;
            return dialogUi.fontDD.selection.text;
        }

        /**
         * 見出しで選択中のフォントファミリーを取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {string} フォントファミリー名
         */
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

        /**
         * 見出しで選択中のフォントスタイル名を取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @param {object} previewRow 対象のプレビュー行
         * @returns {string} フォントスタイル名
         */
        function getHeadingFontStyleName(dialogUi, previewRow) {
            if (dialogUi.disableFontRadio && dialogUi.disableFontRadio.value) return null;
            return getFontStyleDropdownValue(previewRow.fontStyleDD);
        }

        /**
         * 本文の行送り倍率を取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {number} 行送り倍率
         */
        function getBodyLeadingMultiplier(dialogUi) {
            return parsePositivePercentMultiplier(dialogUi.leadingBodyInput.text, DEFAULT_BODY_LEADING_PERCENT / 100);
        }

        /**
         * 見出しの行送り倍率を取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {number} 行送り倍率
         */
        function getHeadingLeadingMultiplier(dialogUi) {
            return parsePositivePercentMultiplier(dialogUi.leadingHeadingInput.text, DEFAULT_HEADING_LEADING_PERCENT / 100);
        }

        /**
         * 本文のカーニング方式を取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {string} カーニング方式
         */
        function getBodyKerningMethod(dialogUi) {
            return getSelectedDropdownOptionValue(dialogUi.bodyKerningDD, kerningOptions, "value", "和文等幅");
        }

        /**
         * 見出しのカーニング方式を取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {string} カーニング方式
         */
        function getHeadingKerningMethod(dialogUi) {
            return getSelectedDropdownOptionValue(dialogUi.headingKerningDD, kerningOptions, "value", "メトリクス");
        }

        /**
         * 本文の段落前アキ（%）を取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {number} 段落前のアキ（%）
         */
        function getBodySpaceBeforePercent(dialogUi) {
            return DEFAULT_BODY_SPACE_BEFORE_PERCENT;
        }

        /**
         * 本文の段落後アキ（%）を取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {number} 段落後のアキ（%）
         */
        function getBodySpaceAfterPercent(dialogUi) {
            return DEFAULT_BODY_SPACE_AFTER_PERCENT;
        }

        /**
         * 見出しの段落前アキ（%）を取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {number} 段落前のアキ（%）
         */
        function getHeadingSpaceBeforePercent(dialogUi) {
            return DEFAULT_SPACE_BEFORE_PERCENT;
        }

        /**
         * 見出しの段落後アキ（%）を取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {number} 段落後のアキ（%）
         */
        function getHeadingSpaceAfterPercent(dialogUi) {
            return DEFAULT_SPACE_AFTER_PERCENT;
        }

        /**
         * 入力されている基準サイズを取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {number} 基準サイズ
         */
        function getBaseSize(dialogUi) {
            return parsePositiveNumber(dialogUi.baseInput.text, null);
        }

        /**
         * 行送りの表示用文字列を作る
         * @param {number} value 対象の数値
         * @param {number} roundDigits 小数点以下の桁数
         * @returns {string} 表示する文字列
         */
        function formatLeadingValue(value, roundDigits) {
            if (typeof value !== "number" || isNaN(value)) return getLabel("option.notAvailable");
            return String(roundTo(value, roundDigits)) + unitSym;
        }

        /**
         * 見出しレベルに対応する段落スタイル名を返す
         * @param {number} dialogUi 見出しレベル数
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
         * @param {DropDownList} dropdownList 対象のドロップダウン
         * @returns {string} フォントスタイル名
         */
        function getFontStyleDropdownValue(dropdownList) {
            if (!dropdownList.selection) return null;
            if (dropdownList.selection.text === getLabel("option.noFontChange")) return null;
            return dropdownList.selection.text;
        }

        /**
         * プレビュー行の有効／無効を切り替える
         * @param {object} previewRow 対象の行
         * @param {boolean} enabled 有効にするなら true
         * @returns {void}
         */
        function setRowEnabled(previewRow, enabled) {
            previewRow.lbl.enabled = enabled;
            previewRow.sizeText.enabled = enabled;
            previewRow.spaceBeforeText.enabled = enabled;
            previewRow.spaceAfterText.enabled = enabled;
            previewRow.styleDD.enabled = enabled;
            previewRow.fontStyleDD.enabled = enabled;
        }

        /**
         * プレビューの全行を取得する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {Array<object>} プレビュー行の配列
         */
        function getAllPreviewRows(dialogUi) {
            var previewRows = [];
            for (var levelRowIndex = 0; levelRowIndex < dialogUi.levelRows.length; levelRowIndex++) {
                previewRows.push(dialogUi.levelRows[levelRowIndex]);
            }
            previewRows.push(dialogUi.baseRow);
            for (var bodyRowIndex = 0; bodyRowIndex < dialogUi.bodyDerivedRows.length; bodyRowIndex++) {
                previewRows.push(dialogUi.bodyDerivedRows[bodyRowIndex]);
            }
            previewRows.push(dialogUi.captionRow);
            return previewRows;
        }

        /**
         * ドロップダウンの項目を入れ替える
         * @param {DropDownList} dropdownList 対象のドロップダウン
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
         * @param {object} dialogUi ダイアログのコントロール一式
         * @param {string} styleName 選択したいフォントスタイル名
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
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {void}
         */
        function syncPreviewFontStylesFromTextSettings(dialogUi) {
            var selectedFontStyleName = getDropdownText(dialogUi.fontStyleDD);
            if (!selectedFontStyleName || selectedFontStyleName === getLabel("option.noFontChange")) return;
            selectAllPreviewFontStyleDropdowns(dialogUi, selectedFontStyleName);
        }

        /**
         * 見出しのフォントスタイルをプレビューへ反映する
         * @param {object} dialogUi ダイアログのコントロール一式
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
         * 「サイズのみ」設定に応じて列のディム表示を切り替える
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {void}
         */
        function syncSizeOnlyColumnDimming(dialogUi) {
            // 「サイズのみ」に連動して、段落前後のアキ列だけを淡色化する
            // （「スタイル」列のヘッダーは syncFontSelectionEnabled 側で制御）
            var enabled = !(dialogUi.sizeOnlyCheckbox && dialogUi.sizeOnlyCheckbox.value);
            if (dialogUi.spaceBeforeHeader) dialogUi.spaceBeforeHeader.enabled = enabled;
            if (dialogUi.spaceAfterHeader) dialogUi.spaceAfterHeader.enabled = enabled;
            var previewRows = getAllPreviewRows(dialogUi);
            for (var previewRowIndex = 0; previewRowIndex < previewRows.length; previewRowIndex++) {
                previewRows[previewRowIndex].spaceBeforeText.enabled = enabled;
                previewRows[previewRowIndex].spaceAfterText.enabled = enabled;
            }
        }

        /**
         * フォント指定モードに応じてコントロールの有効／無効を切り替える
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {void}
         */
        function syncFontSelectionEnabled(dialogUi) {
            var enabled = !(dialogUi.disableFontRadio && dialogUi.disableFontRadio.value);
            var headingFontEnabled = enabled && !(dialogUi.useSameFontRadio && dialogUi.useSameFontRadio.value);
            dialogUi.fontDD.enabled = enabled;
            dialogUi.fontStyleDD.enabled = enabled;
            dialogUi.headingFontDD.enabled = headingFontEnabled;
            dialogUi.headingFontStyleDD.enabled = headingFontEnabled;
            // 「スタイル（ウエイト）」列は、フォント指定が有効なときだけ操作可能にする（未アクティブ／変更しない時はディム）
            if (dialogUi.fontStyleHeader) dialogUi.fontStyleHeader.enabled = enabled;
            // フォント指定モードのラジオは「フォント、スタイルを含める」で有効化するまで停止状態
            var radiosEnabled = !!_fontSectionActive;
            if (dialogUi.useSameFontRadio) dialogUi.useSameFontRadio.enabled = radiosEnabled;
            if (dialogUi.separateFontRadio) dialogUi.separateFontRadio.enabled = radiosEnabled;
            if (dialogUi.disableFontRadio) dialogUi.disableFontRadio.enabled = radiosEnabled;
            if (!enabled) {
                selectDropdownByText(dialogUi.fontDD, getLabel("option.noFontChange"));
                resetDropdownItems(dialogUi.fontStyleDD, [getLabel("option.noFontChange")]);
                selectDropdownByText(dialogUi.headingFontDD, getLabel("option.refBodyFont"));
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
         * @param {object} dialogUi ダイアログのコントロール一式
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
         * @param {object} dialogUi ダイアログのコントロール一式
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

            // ここに到達するのは「別指定」モードのみ（共通指定は冒頭で return 済み）。見出しフォントは選択可能
            dialogUi.headingFontDD.enabled = true;
            dialogUi.headingFontDD.helpTip = "";

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
         * @param {object} dialogUi ダイアログのコントロール一式
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
         * プレビューで上書きされたサイズを集める
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {object} レベルごとのサイズ
         */
        function collectLevelSizeOverrides(dialogUi) {
            var levelOverrides = [];
            for (var i = 0; i < dialogUi.levelRows.length; i++) {
                levelOverrides.push(dialogUi.levelRows[i].sizeOverride);
            }
            return levelOverrides;
        }

        /**
         * プレビューで上書きされた段落前アキを集める
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {object} レベルごとの段落前アキ
         */
        function collectLevelSpaceBeforeOverrides(dialogUi) {
            var levelOverrides = [];
            for (var i = 0; i < dialogUi.levelRows.length; i++) {
                levelOverrides.push(dialogUi.levelRows[i].spaceBeforeOverride);
            }
            return levelOverrides;
        }

        /**
         * プレビューで上書きされた段落後アキを集める
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {object} レベルごとの段落後アキ
         */
        function collectLevelSpaceAfterOverrides(dialogUi) {
            var levelOverrides = [];
            for (var i = 0; i < dialogUi.levelRows.length; i++) {
                levelOverrides.push(dialogUi.levelRows[i].spaceAfterOverride);
            }
            return levelOverrides;
        }

        // 本文派生行（リスト・テーブル）の適用情報を収集する
        /**
         * 本文サイズから派生する行（リスト・表など）の設定を集める
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {Array<object>} 派生行の設定
         */
        function collectBodyDerivedRows(dialogUi) {
            var rows = [];
            for (var bodyRowIndex = 0; bodyRowIndex < dialogUi.bodyDerivedRows.length; bodyRowIndex++) {
                var bodyRow = dialogUi.bodyDerivedRows[bodyRowIndex];
                rows.push({
                    styleName: getDropdownText(bodyRow.styleDD),
                    fontStyleName: getFontStyleDropdownValue(bodyRow.fontStyleDD),
                    sizeFactor: bodyRow.bodySizeFactor,
                    sizeOverride: bodyRow.sizeOverride,
                    spaceBeforeOverride: bodyRow.spaceBeforeOverride,
                    spaceAfterOverride: bodyRow.spaceAfterOverride
                });
            }
            return rows;
        }

        /**
         * ダイアログの入力内容を設定オブジェクトにまとめる
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {object} 適用に使う設定
         */
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
                captionSpaceAfterOverride: dialogUi.captionRow.spaceAfterOverride,
                bodyDerivedRows: collectBodyDerivedRows(dialogUi),
                sizeOnly: !!(dialogUi.sizeOnlyCheckbox && dialogUi.sizeOnlyCheckbox.value),
                originalStyleProps: originalStyleProps
            };
        }

        /**
         * 現在の設定でプレビューを描き直す
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {void}
         */
        function updateTypescalePreview(dialogUi) {
            var baseSize = getBaseSize(dialogUi);
            var ratio = getCurrentScaleRatio(dialogUi);
            var levelCount = getCurrentLevelCount(dialogUi);
            var roundDigits = getCurrentRoundDigits(dialogUi);

            if (baseSize === null) {
                for (var levelRowIndex = 0; levelRowIndex < dialogUi.levelRows.length; levelRowIndex++) {
                    dialogUi.levelRows[levelRowIndex].sizeText.text = getLabel("option.notAvailable");
                    dialogUi.levelRows[levelRowIndex].spaceBeforeText.text = getLabel("option.notAvailable");
                    dialogUi.levelRows[levelRowIndex].spaceAfterText.text = getLabel("option.notAvailable");
                    setRowEnabled(dialogUi.levelRows[levelRowIndex], false);
                }
                dialogUi.baseRow.sizeText.text = getLabel("option.notAvailable");
                dialogUi.baseRow.spaceBeforeText.text = getLabel("option.notAvailable");
                dialogUi.baseRow.spaceAfterText.text = getLabel("option.notAvailable");
                for (var bodyRowIndex = 0; bodyRowIndex < dialogUi.bodyDerivedRows.length; bodyRowIndex++) {
                    dialogUi.bodyDerivedRows[bodyRowIndex].sizeText.text = getLabel("option.notAvailable");
                    dialogUi.bodyDerivedRows[bodyRowIndex].spaceBeforeText.text = getLabel("option.notAvailable");
                    dialogUi.bodyDerivedRows[bodyRowIndex].spaceAfterText.text = getLabel("option.notAvailable");
                }
                dialogUi.captionRow.sizeText.text = getLabel("option.notAvailable");
                dialogUi.captionRow.spaceBeforeText.text = getLabel("option.notAvailable");
                dialogUi.captionRow.spaceAfterText.text = getLabel("option.notAvailable");
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
                    levelRow.sizeText.text = getLabel("option.notAvailable");
                    levelRow.spaceBeforeText.text = getLabel("option.notAvailable");
                    levelRow.spaceAfterText.text = getLabel("option.notAvailable");
                    setRowEnabled(levelRow, false);
                }
            }
            var effectiveBaseSize = (typeof dialogUi.baseRow.sizeOverride === "number") ? dialogUi.baseRow.sizeOverride : computedSizes.base;
            dialogUi.baseRow.sizeText.text = roundTo(effectiveBaseSize, roundDigits) + " " + unitSym;
            var effectiveBaseSpaceBefore = (typeof dialogUi.baseRow.spaceBeforeOverride === "number") ? dialogUi.baseRow.spaceBeforeOverride : effectiveBaseSize * bodySpaceBeforeRatio;
            dialogUi.baseRow.spaceBeforeText.text = formatLeadingValue(effectiveBaseSpaceBefore, roundDigits);
            var effectiveBaseSpaceAfter = (typeof dialogUi.baseRow.spaceAfterOverride === "number") ? dialogUi.baseRow.spaceAfterOverride : effectiveBaseSize * bodySpaceAfterRatio;
            dialogUi.baseRow.spaceAfterText.text = formatLeadingValue(effectiveBaseSpaceAfter, roundDigits);
            // 本文派生行（リスト・テーブル）: 文字サイズは基準（本文）×係数
            for (var bodyRowIndex = 0; bodyRowIndex < dialogUi.bodyDerivedRows.length; bodyRowIndex++) {
                var bodyRow = dialogUi.bodyDerivedRows[bodyRowIndex];
                var bodyRowSize = computedSizes.base * bodyRow.bodySizeFactor;
                var effectiveBodyRowSize = (typeof bodyRow.sizeOverride === "number") ? bodyRow.sizeOverride : bodyRowSize;
                bodyRow.sizeText.text = roundTo(effectiveBodyRowSize, roundDigits) + " " + unitSym;
                var effectiveBodyRowSpaceBefore = (typeof bodyRow.spaceBeforeOverride === "number") ? bodyRow.spaceBeforeOverride : effectiveBodyRowSize * bodySpaceBeforeRatio;
                bodyRow.spaceBeforeText.text = formatLeadingValue(effectiveBodyRowSpaceBefore, roundDigits);
                var effectiveBodyRowSpaceAfter = (typeof bodyRow.spaceAfterOverride === "number") ? bodyRow.spaceAfterOverride : effectiveBodyRowSize * bodySpaceAfterRatio;
                bodyRow.spaceAfterText.text = formatLeadingValue(effectiveBodyRowSpaceAfter, roundDigits);
            }
            var effectiveCaptionSize = (typeof dialogUi.captionRow.sizeOverride === "number") ? dialogUi.captionRow.sizeOverride : computedSizes.caption;
            dialogUi.captionRow.sizeText.text = roundTo(effectiveCaptionSize, roundDigits) + " " + unitSym;
            var effectiveCaptionSpaceBefore = (typeof dialogUi.captionRow.spaceBeforeOverride === "number") ? dialogUi.captionRow.spaceBeforeOverride : effectiveCaptionSize * bodySpaceBeforeRatio;
            dialogUi.captionRow.spaceBeforeText.text = formatLeadingValue(effectiveCaptionSpaceBefore, roundDigits);
            var effectiveCaptionSpaceAfter = (typeof dialogUi.captionRow.spaceAfterOverride === "number") ? dialogUi.captionRow.spaceAfterOverride : effectiveCaptionSize * bodySpaceAfterRatio;
            dialogUi.captionRow.spaceAfterText.text = formatLeadingValue(effectiveCaptionSpaceAfter, roundDigits);

            applyTypescaleSettings(targetDocument, collectTypescaleSettings(dialogUi), true, unit);
            syncSizeOnlyColumnDimming(dialogUi);
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
        function clearOverridesIfActive() {
            clearTextOverridesInSelection();
            try { app.menuActions.itemByID(8489).invoke(); } catch (e) { }
            try { app.redraw(); } catch (e2) { }
        }

        /**
         * フォント変更後にプレビューを更新する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {void}
         */
        function refreshPreviewAfterFontChange(dialogUi) {
            clearOverridesIfActive();
            updateTypescalePreview(dialogUi);
        }

        // 「フォント、スタイルを含める」: 停止中のフォント・スタイル指定を有効化する
        // （フォント自体は起動時に読み込み済みなので、ここではモードの切り替えと有効化のみ）
        /**
         * 停止中のフォント・スタイル指定を有効にし、プレビューへ反映する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {void}
         */
        function includeFontsAndStyles(dialogUi) {
            _fontSectionActive = true;
            // 本文と見出しで共通の指定からアクティブ化する
            setFontOptionMode(dialogUi, "same");
            if (dialogUi.sizeOnlyCheckbox) dialogUi.sizeOnlyCheckbox.value = false;
            if (dialogUi.includeFontsButton) dialogUi.includeFontsButton.enabled = false;
            // 初回のフォント適用（フォントオブジェクト解決＋段落スタイル反映）は時間がかかるため、
            // ダイアログ内の進捗バーを表示する（モーダル中は別ウィンドウを出せないため）
            if (dialogUi.applyProgressBar) {
                dialogUi.applyProgressBar.value = 40;
                dialogUi.applyProgressBar.visible = true;
                try { dialogUi.dialog.update(); } catch (eShowProgress) { }
            }
            syncFontSelectionEnabled(dialogUi);
            updateFontStyleDropdowns(dialogUi);
            updateHeadingFontStyleDropdowns(dialogUi);
            syncPreviewFontStylesFromTextSettings(dialogUi);
            updateTypescalePreview(dialogUi);
            if (dialogUi.applyProgressBar) {
                dialogUi.applyProgressBar.value = 100;
                try { dialogUi.dialog.update(); } catch (eDoneProgress) { }
                dialogUi.applyProgressBar.visible = false;
                try { dialogUi.dialog.update(); } catch (eHideProgress) { }
            }
        }

        /**
         * フォント指定モードを切り替える
         * @param {object} dialogUi ダイアログのコントロール一式
         * @param {string} mode "same" / "separate" / "disable" のいずれか
         * @returns {void}
         */
        function setFontOptionMode(dialogUi, mode) {
            dialogUi.useSameFontRadio.value = (mode === "same");
            dialogUi.separateFontRadio.value = (mode === "separate");
            dialogUi.disableFontRadio.value = (mode === "disable");
        }

        /**
         * スケールプリセットに合わせて見出しレベルの選択を揃える
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {void}
         */
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

        /**
         * プレビューの上書き入力をすべて消去する
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {void}
         */
        function clearAllPreviewOverrides(dialogUi) {
            for (var levelClearIndex = 0; levelClearIndex < dialogUi.levelRows.length; levelClearIndex++) {
                dialogUi.levelRows[levelClearIndex].sizeOverride = null;
                dialogUi.levelRows[levelClearIndex].spaceBeforeOverride = null;
                dialogUi.levelRows[levelClearIndex].spaceAfterOverride = null;
            }
            dialogUi.baseRow.sizeOverride = null;
            dialogUi.baseRow.spaceBeforeOverride = null;
            dialogUi.baseRow.spaceAfterOverride = null;
            for (var bodyClearIndex = 0; bodyClearIndex < dialogUi.bodyDerivedRows.length; bodyClearIndex++) {
                dialogUi.bodyDerivedRows[bodyClearIndex].sizeOverride = null;
                dialogUi.bodyDerivedRows[bodyClearIndex].spaceBeforeOverride = null;
                dialogUi.bodyDerivedRows[bodyClearIndex].spaceAfterOverride = null;
            }
            dialogUi.captionRow.sizeOverride = null;
            dialogUi.captionRow.spaceBeforeOverride = null;
            dialogUi.captionRow.spaceAfterOverride = null;
        }

        /**
         * プレビューの入力欄に上下キーでの増減操作を追加する
         * @param {EditText} editText 対象の入力欄
         * @param {boolean} allowZero 0 を許可するなら true
         * @param {function} applyValue 変更後の値を適用する処理
         * @returns {void}
         */
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

        /**
         * プレビューのセル編集にイベントを結び付ける
         * @param {object} dialogUi ダイアログのコントロール一式
         * @param {object} previewRow 対象のプレビュー行
         * @param {function} [sizeCallback] 文字サイズ変更時の処理。省略時は行の個別指定を更新する
         * @returns {void}
         */
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

        /**
         * ダイアログ全体のイベントを結び付ける
         * @param {object} dialogUi ダイアログのコントロール一式
         * @returns {void}
         */
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
            dialogUi.scaleDD.onChange = function () {
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
            for (var bodyBindIndex = 0; bodyBindIndex < dialogUi.bodyDerivedRows.length; bodyBindIndex++) {
                bindPreviewCellEdit(dialogUi, dialogUi.bodyDerivedRows[bodyBindIndex]);
                dialogUi.bodyDerivedRows[bodyBindIndex].styleDD.onChange = function () { updateTypescalePreview(dialogUi); };
                dialogUi.bodyDerivedRows[bodyBindIndex].fontStyleDD.onChange = function () { updateTypescalePreview(dialogUi); };
            }
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
                if (dialogUi.sizeOnlyCheckbox) dialogUi.sizeOnlyCheckbox.value = false;
                syncFontSelectionEnabled(dialogUi);
                updateFontStyleDropdowns(dialogUi);
                updateHeadingFontStyleDropdowns(dialogUi);
                syncPreviewFontStylesFromTextSettings(dialogUi);
                updateTypescalePreview(dialogUi);
            };
            dialogUi.separateFontRadio.onClick = function () {
                setFontOptionMode(dialogUi, "separate");
                if (dialogUi.sizeOnlyCheckbox) dialogUi.sizeOnlyCheckbox.value = false;
                syncFontSelectionEnabled(dialogUi);
                updateFontStyleDropdowns(dialogUi);
                updateHeadingFontStyleDropdowns(dialogUi);
                updateTypescalePreview(dialogUi);
            };
            dialogUi.sizeOnlyCheckbox.onClick = function () {
                if (dialogUi.sizeOnlyCheckbox.value) {
                    setFontOptionMode(dialogUi, "disable");
                    syncFontSelectionEnabled(dialogUi);
                    updateFontStyleDropdowns(dialogUi);
                    updateHeadingFontStyleDropdowns(dialogUi);
                }
                syncSizeOnlyColumnDimming(dialogUi);
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
            dialogUi.includeFontsButton.onClick = function () { includeFontsAndStyles(dialogUi); };
        }

        bindTypescaleDialogEvents(dialogUi);
        // フォントドロップダウンはファミリー一覧を反映済みの状態で構築されている（先頭でロード済み）
        updateFontStyleDropdowns(dialogUi);
        updateHeadingFontStyleDropdowns(dialogUi);
        syncLevelRadiosWithScaleOption(dialogUi);
        syncFontSelectionEnabled(dialogUi);
        // ライブプレビューが書き換える前の段落スタイルを退避し、初回プレビューを反映してから開く
        originalStyleProps = snapshotParagraphStyleProps(targetDocument);
        updateTypescalePreview(dialogUi);

        // モーダルダイアログを表示する直前にパレットを閉じる（モーダル表示中は表示しない）
        closeProgressPalette(loadingPalette);
        var dialogResult = dialogUi.dialog.show();
        clearOverridesIfActive();

        if (dialogResult !== 1) {
            // キャンセル：ライブプレビューで書き換えた段落スタイルを起動時の状態へ戻す
            restoreParagraphStyleProps(targetDocument, originalStyleProps, _previewModifiedStyles);
            return null;
        }

        if (getBaseSize(dialogUi) === null) {
            alert(getLabel("error.invalidBaseSize"));
            return null;
        }
        return collectTypescaleSettings(dialogUi);
    }

    // スタイル名に応じた「同じスタイルの段落間隔」を返す（適用対象外は null） / Resolve same-style spacing by style name (null = leave unchanged)
    /**
     * スタイル名に応じた「同じスタイルの段落間隔」を決める
     * @param {string} styleName 段落スタイル名
     * @param {number} spaceBeforePt 段落前のアキ
     * @returns {number|null} 設定する間隔。変更しない場合は null
     */
    function resolveSameStyleSpacing(styleName, spaceBeforePt) {
        if (styleName === "ul-li") return 0;
        if (styleName === "ol-li" || styleName === "p") {
            return (typeof spaceBeforePt === "number" && spaceBeforePt >= 0) ? spaceBeforePt : null;
        }
        return null;
    }

    /* 控えた値を段落スタイルへ書き戻す。プロパティごとに設定できない場合があるので個別に握りつぶす
       / Write a snapshot back to a style; each property is guarded because some may be read-only */
    /**
     * 控えておいた値を段落スタイルへ書き戻す
     * @param {ParagraphStyle} style 対象の段落スタイル
     * @param {object} snapshot 控えたプロパティ
     * @param {boolean} includePointSize 文字サイズも戻すなら true
     * @returns {void}
     */
    function assignSnapshotToStyle(style, snapshot, includePointSize) {
        if (!style || !snapshot) return;
        if (includePointSize && typeof snapshot.pointSize !== "undefined") {
            try { style.pointSize = snapshot.pointSize; } catch (ePS) { }
        }
        if (snapshot.appliedFont) { try { style.appliedFont = snapshot.appliedFont; } catch (eF) { } }
        if (snapshot.fontStyle) { try { style.fontStyle = snapshot.fontStyle; } catch (eFS) { } }
        if (typeof snapshot.leading !== "undefined") { try { style.leading = snapshot.leading; } catch (eL) { } }
        if (typeof snapshot.spaceBefore !== "undefined") { try { style.spaceBefore = snapshot.spaceBefore; } catch (eSB) { } }
        if (typeof snapshot.spaceAfter !== "undefined") { try { style.spaceAfter = snapshot.spaceAfter; } catch (eSA) { } }
        if (typeof snapshot.kerningMethod !== "undefined") { try { style.kerningMethod = snapshot.kerningMethod; } catch (eK) { } }
        if (typeof snapshot.justification !== "undefined") { try { style.justification = snapshot.justification; } catch (eJ) { } }
        if (typeof snapshot.sameParaStyleSpacing !== "undefined") {
            try { style.sameParaStyleSpacing = snapshot.sameParaStyleSpacing; } catch (eSSS) { }
        }
    }

    /**
     * 段落スタイルへ各プロパティを適用する
     * @param {Document} targetDocument 対象ドキュメント
     * @param {string} styleName 対象の段落スタイル名
     * @param {object} styleProps 適用するプロパティ
     * @param {number} styleProps.size 文字サイズ（pt）
     * @param {Font} [styleProps.font] 適用するフォント
     * @param {number} [styleProps.leading] 行送り（pt）
     * @param {number} [styleProps.spaceBefore] 段落前のアキ（pt）
     * @param {number} [styleProps.spaceAfter] 段落後のアキ（pt）
     * @param {string} [styleProps.kerningMethod] カーニング方式
     * @param {boolean} styleProps.isHeading 見出しなら true
     * @param {boolean} styleProps.silent スタイル未検出時に警告を出さないなら true
     * @param {boolean} styleProps.sizeOnly 文字サイズだけを更新するなら true
     * @param {object} [styleProps.originalProps] ダイアログ起動時に控えた値
     * @returns {boolean} 適用できたら true
     */
    function setParagraphStyleProps(targetDocument, styleName, styleProps) {
        var size = styleProps.size;
        var font = styleProps.font;
        var leading = styleProps.leading;
        var spaceBefore = styleProps.spaceBefore;
        var spaceAfter = styleProps.spaceAfter;
        var kerningMethod = styleProps.kerningMethod;
        var isHeading = styleProps.isHeading;
        var silent = styleProps.silent;
        var originalProps = styleProps.originalProps;

        var style = findParagraphStyle(targetDocument, styleName);
        if (style === null) {
            if (!silent) alert(formatLabel("error.missingStyle", styleName));
            return false;
        }
        // ルート／組み込みスタイル（[基本段落] / [Basic Paragraph] / [段落スタイルなし] など）は
        // pointSize 等の設定で error 516「ルートスタイルに対する無効な要求」になるためスキップ
        if (isRootParagraphStyle(style)) return false;
        // 書き換えるスタイルを記録（キャンセル時にこのスタイルだけ起動時状態へ戻す）
        _previewModifiedStyles[styleName] = true;
        // サイズのみモード: サイズだけを更新し、その他はダイアログ起動時の値に戻す
        if (styleProps.sizeOnly) {
            style.pointSize = size;
            assignSnapshotToStyle(style, originalProps, false);
            return true;
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
        if (ENABLE_SPACE_BEFORE && typeof spaceBefore === "number" && spaceBefore >= 0) {
            style.spaceBefore = spaceBefore;
        }
        if (ENABLE_SPACE_AFTER && typeof spaceAfter === "number" && spaceAfter >= 0) {
            style.spaceAfter = spaceAfter;
        }
        // 同じスタイルの段落間隔（sameParaStyleSpacing）をスタイル名ごとのルールで設定
        if (ENABLE_SAME_STYLE_SPACING) {
            var sameStyleSpacing = resolveSameStyleSpacing(styleName, spaceBefore);
            if (sameStyleSpacing !== null) {
                try { style.sameParaStyleSpacing = sameStyleSpacing; } catch (eSSS) { }
            }
        }
        // フォントによっては設定不可のため、安全に無視
        if (kerningMethod) {
            try { style.kerningMethod = kerningMethod; } catch (ke) { }
        }
        // 字揃えは既定では変更しない（元の字揃えを保持）。ENABLE_JUSTIFICATION を true にすると強制適用する
        if (ENABLE_JUSTIFICATION) {
            try {
                if (isHeading) {
                    // 見出しは左揃え / Headings are left aligned
                    style.justification = Justification.LEFT_ALIGN;
                } else {
                    // 本文・キャプションは均等配置（最終行左） / Body and captions are left-justified
                    style.justification = Justification.LEFT_JUSTIFIED;
                }
            } catch (e) { }
        }
        return true;
    }

    // 「サイズのみ」復元・キャンセル復元用に、ダイアログ起動時点の段落スタイルのプロパティを記録する
    /**
     * すべての段落スタイルの現在値を控える
     * @param {Document} targetDocument 対象ドキュメント
     * @returns {object} 段落スタイル名をキーにしたプロパティの控え
     */
    function snapshotParagraphStyleProps(targetDocument) {
        var snapshot = {};
        var styles = targetDocument.allParagraphStyles;
        for (var snapshotIndex = 0; snapshotIndex < styles.length; snapshotIndex++) {
            var paragraphStyle = styles[snapshotIndex];
            var styleSnapshot = {};
            try { styleSnapshot.pointSize = paragraphStyle.pointSize; } catch (ePS) { }
            try { styleSnapshot.leading = paragraphStyle.leading; } catch (eL) { }
            try { styleSnapshot.spaceBefore = paragraphStyle.spaceBefore; } catch (eSB) { }
            try { styleSnapshot.spaceAfter = paragraphStyle.spaceAfter; } catch (eSA) { }
            try { styleSnapshot.justification = paragraphStyle.justification; } catch (eJ) { }
            try { styleSnapshot.appliedFont = paragraphStyle.appliedFont; } catch (eF) { }
            try { styleSnapshot.fontStyle = paragraphStyle.fontStyle; } catch (eFS) { }
            try { styleSnapshot.kerningMethod = paragraphStyle.kerningMethod; } catch (eK) { }
            try { styleSnapshot.sameParaStyleSpacing = paragraphStyle.sameParaStyleSpacing; } catch (eSSS) { }
            snapshot[paragraphStyle.name] = styleSnapshot;
        }
        return snapshot;
    }

    // キャンセル時：プレビューで書き換えたスタイルだけを起動時の状態へ戻す
    /**
     * 書き換えた段落スタイルを控えておいた値に戻す
     * @param {Document} targetDocument 対象ドキュメント
     * @param {object} originalStyleProps 控えたプロパティ。段落スタイル名がキー
     * @param {object} modifiedStyleNames 書き換えた段落スタイル名の集合
     * @returns {void}
     */
    function restoreParagraphStyleProps(targetDocument, originalStyleProps, modifiedStyleNames) {
        if (!originalStyleProps || !modifiedStyleNames) return;
        for (var styleName in modifiedStyleNames) {
            if (!modifiedStyleNames.hasOwnProperty(styleName)) continue;
            var snap = originalStyleProps[styleName];
            if (!snap) continue;
            var style = findParagraphStyle(targetDocument, styleName);
            if (style === null) continue;
            assignSnapshotToStyle(style, snap, true);
        }
    }

    /* 段落スタイル名 → スタイルの対応表。ライブプレビューは 1 キーストロークごとに
       行数ぶんスタイルを引くため、allParagraphStyles の走査を毎回行わずキャッシュする。
       ダイアログはモーダルで、表示中にスタイルが増減することはない
       / Cached name-to-style table; the live preview looks styles up on every keystroke */
    var _paragraphStyleMap = null;

    /**
     * 段落スタイル名から段落スタイルを引くための対応表を返す
     * @param {Document} targetDocument 対象ドキュメント
     * @returns {object} スタイル名をキーにした対応表
     */
    function getParagraphStyleMap(targetDocument) {
        if (_paragraphStyleMap !== null) return _paragraphStyleMap;
        var map = {};
        var styles = targetDocument.allParagraphStyles;
        for (var styleIndex = 0; styleIndex < styles.length; styleIndex++) {
            var style = styles[styleIndex];
            /* 同名スタイルが複数ある場合は最初のものを採用（従来の探索と同じ）
               / Keep the first match when names collide, as the previous scan did */
            if (!map.hasOwnProperty(style.name)) map[style.name] = style;
        }
        _paragraphStyleMap = map;
        return map;
    }

    /**
     * 名前から段落スタイルを探す
     * @param {Document} targetDocument 対象ドキュメント
     * @param {string} styleName 段落スタイル名
     * @returns {ParagraphStyle|null} 段落スタイル。見つからない場合は null
     */
    function findParagraphStyle(targetDocument, styleName) {
        var map = getParagraphStyleMap(targetDocument);
        return map.hasOwnProperty(styleName) ? map[styleName] : null;
    }

    // ルート／組み込みスタイルかどうか（名前が "[" で始まる：[基本段落] / [Basic Paragraph] / [段落スタイルなし] など）
    /**
     * 組み込みスタイル（名前が [ で始まる）かどうかを判定する
     * @param {ParagraphStyle} style 対象の段落スタイル
     * @returns {boolean} 組み込みスタイルなら true
     */
    function isRootParagraphStyle(style) {
        try {
            return !!(style && style.name && style.name.charAt(0) === "[");
        } catch (e) {
            return false;
        }
    }

})();