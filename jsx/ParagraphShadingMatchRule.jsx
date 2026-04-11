#target indesign

/*
 * スクリプト概要：
 * 選択したテキスト内の各段落に対して「段落境界線（前境界線）」を設定します。
 * 段落背景色の上オフセットとフォントサイズをもとに、
 * 段落境界線の線幅を計算し、背景と一致する見た目を再現します。
 * 罫線カラーは「なし（None）」に設定され、実質的にレイアウト調整用の不可視オブジェクトとして機能します。
 * ※重要：このスクリプトは見た目の線を描くためのものではなく、
 * 「太い透明の段落境界線」をスペーサーとして使うことで背景領域の高さを疑似的に再現する設計です。
 *
 * 主な処理内容 / Key Features:
 * - Validates that a text object is selected
 * - Normalizes the selection to a text container (insertion point, word, paragraph, etc.)
 * - 各段落ごとに以下を設定
 *   ・段落境界線の前境界線を有効化
 *   ・段落境界線の線幅を「段落背景色の上オフセット + フォントサイズ」で算出
 *   ・段落境界線のカラーを「なし」に設定
 *   ・フレーム内に罫線を保持
 *
 * 注意点 / Notes:
 * - InDesign の数値は基本的に pt 単位で返るため、単位変換は行わず pt のまま計算しています。
 * - 本スクリプトは「位置調整」ではなく「太い透明の段落境界線による擬似的な背景制御」を行う設計です。
 * - 段落背景色の上マージンが未設定の場合は 0 として扱います。
 *
 * Script Overview / スクリプト概要:
 * Applies a "Rule Above" to each paragraph in the selected text.
 * The rule weight is calculated from the paragraph shading top offset
 * and the font size, effectively matching the height of the shaded area.
 * The rule color is set to "None", so the rule acts as an invisible spacer
 * for layout control rather than a visible line.
 * Important:
 * This script is not intended to draw a visible line.
 * It uses a thick invisible paragraph rule as a spacer to approximate the height of the shaded area.
 *
 * Key Features / 主な処理内容:
 * - Validates that a text object is selected
 * - Normalizes the selection to a text container (insertion point, word, paragraph, etc.)
 * - For each paragraph:
 *   • Enables rule above
 *   • Calculates rule weight (top offset + font size)
 *   • Sets rule color to None
 *   • Keeps the rule inside the text frame
 *
 * Notes / 注意点:
 * - All numeric values are handled in points (pt) to avoid unit inconsistency.
 * - This script does NOT reposition objects; it uses a thick invisible rule as a layout spacer.
 * - Missing shading offsets are treated as 0.
 */

// =========================================
// バージョンとローカライズ
// =========================================
// LABELS は ui / error / process の3系統で管理し、キーは「用途.対象.意味」の階層で命名する（例: error.selectText.message / process.undo.label）
// - ui: UI表示用（ダイアログ・ボタンなど）
// - error: エラーメッセージ
// - process: 内部処理名（Undo名など）

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    error: {
        selectText: {
            ja: "テキストを選択してください。",
            en: "Please select text."
        }
    },
    process: {
        undo: {
            label: {
                ja: "段落ルール設定",
                en: "Set Paragraph Rules"
            }
        }
    }
};

function L(path) {
    var keys = path.split(".");
    var value = LABELS;
    var i;

    for (i = 0; i < keys.length; i++) {
        if (!value || !value.hasOwnProperty(keys[i])) {
            throw new Error("Missing label path: " + path);
        }
        value = value[keys[i]];
    }

    if (!value || !value.hasOwnProperty(lang)) {
        throw new Error("Missing label language: " + path + " / " + lang);
    }

    return value[lang];
}

function toNumberOrZero(value) {
    value = Number(value);
    return isNaN(value) ? 0 : value;
}

function getTargetTextFromSelection(selectionItem) {
    var textTypeNames = {
        Text: true,
        InsertionPoint: true,
        Word: true,
        Line: true,
        TextStyleRange: true,
        Paragraph: true
    };

    if (!selectionItem) {
        return null;
    }

    if (selectionItem.hasOwnProperty("baseline")) {
        return selectionItem;
    }

    if (textTypeNames[selectionItem.constructor.name]) {
        return selectionItem;
    }

    if (selectionItem.hasOwnProperty("parentStory")) {
        return selectionItem.parentStory;
    }

    return null;
}

app.doScript(function () {
    if (app.selection.length === 0) {
        alert(L("error.selectText"));
        return;
    }

    var targetText = getTargetTextFromSelection(app.selection[0]);

    if (!targetText || !targetText.paragraphs || targetText.paragraphs.length === 0) {
        alert(L("error.selectText"));
        return;
    }

    var noneSwatch = app.activeDocument.swatches.itemByName("None");
    var paragraphs = targetText.paragraphs;

    for (var i = 0; i < paragraphs.length; i++) {
        var p = paragraphs[i];

        // 段落シェーディングの上オフセット（pt）とフォントサイズ（pt）を取得
        // この2つを合算して、シェーディング領域の高さに近い値を作る
        // InDesign は基本的に pt 単位で値を返すため、単位変換は行わず pt のまま扱う
        var topOffsetPt = toNumberOrZero(p.paragraphShadingTopOffset);
        var fontSizePt = toNumberOrZero(p.pointSize);

        // ---- 段落罫線（上）を透明スペーサーとして使う設定 ----
        // 見た目の線を描くのではなく、レイアウト調整用の不可視領域を作る
        p.ruleAbove = true;                    // 段落境界線の前境界線を有効化
        p.ruleAboveLineWeight = topOffsetPt + fontSizePt;    // シェーディング領域の高さを疑似的に再現するため、透明罫線の太さを「オフセット + フォントサイズ」で設定
        p.ruleAboveColor = noneSwatch;         // 段落境界線のカラーを「なし」（不可視）に設定
        p.keepRuleAboveInFrame = true;         // テキストフレーム内に罫線を保持
    }

}, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, L("process.undo.label") + " " + SCRIPT_VERSION);