#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdAddPagesUsingCurrentMaster
Name=現在の親ページを引き継いでページを挿入
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.4.1
Release-Date=2025-06-26
Description-BEGIN
現在のページに適用されている親（マスター）ページを引き継いだまま、指定した枚数のページを直後に挿入します。

Inserts the requested number of pages right after the current page, keeping the parent (master) page that the current page uses.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

現在のページに適用されている親（マスター）ページを引き継いだまま、指定した枚数のページを直後に挿入します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAddPagesUsingCurrentMaster.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n2d1a76097a39

### Overview

Inserts the requested number of pages right after the current page, keeping the parent (master) page that the current page uses.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAddPagesUsingCurrentMaster.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdAddPagesUsingCurrentMaster"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "現在の親ページを引き継いでページを挿入"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAddPagesUsingCurrentMaster.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAddPagesUsingCurrentMaster.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n2d1a76097a39"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/page/IdAddPagesUsingCurrentMaster.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* ダイアログの初期ページ数 / Default page count in the dialog */
var DEFAULT_PAGE_COUNT = 2;

// ==============================
// UIレイアウトの共通設定 / Shared UI layout
// ==============================

// UIレイアウト（再利用パーツ） / UI layout (reusable)

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

/**
 * ウィンドウの共通設定
 * @param {Window} targetWindow - 対象のウィンドウ
 * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
 * @returns {void}
 */
function setupWindow(targetWindow, spacing) {
    targetWindow.orientation = "column";
    targetWindow.alignChildren = "fill";
    targetWindow.margins = WINDOW_MARGINS;
    targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
 * @param {Panel} targetPanel - 対象のパネル
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
 * @returns {void}
 */
function setupPanel(targetPanel, spacing) {
    targetPanel.orientation = "column";
    targetPanel.alignChildren = ["fill", "top"];
    targetPanel.alignment = "fill";
    targetPanel.margins = PANEL_MARGINS;
    targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * タブの共通設定
 * @param {Tab} targetTab - 対象のタブ
 * @param {number} [spacing] - 要素間隔（省略時は変えない）
 * @returns {void}
 */
function setupTab(targetTab, spacing) {
    targetTab.orientation = "column";
    targetTab.alignChildren = "fill";
    targetTab.margins = TAB_MARGINS;
    if (typeof spacing === "number") targetTab.spacing = spacing;
}

/**
 * 横並びの行グループの共通設定（ボタン列など）。
 * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
 * @param {Group} rowGroup - 対象のグループ
 * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
 * @returns {void}
 */
function setupRow(rowGroup, rowAlignment, spacing) {
    rowGroup.orientation = "row";
    rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
    rowGroup.alignChildren = ["left", "center"];
    rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
 * @param {Button} targetButton - 対象のボタン
 * @param {number} trimPixels - 詰める量（px）
 * @returns {void}
 */
function trimButtonHeight(targetButton, trimPixels) {
    /* レイアウト前は size が無い / size is not set until the layout runs */
    if (!targetButton.size) return;
    targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
}

// UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

/* 挿入後の見開きの縮小図 / Thumbnail of the spreads after insertion */
var THUMBNAIL_SIZE = [220, 76];            /* 描画域（幅は最小幅。ダイアログの幅いっぱいに伸ばす）/ canvas (minimum width; fills the dialog width) */
var THUMBNAIL_PADDING = 8;                 /* 描画域の内側の余白 / inner padding */
var THUMBNAIL_PAGE_HEIGHT = 40;            /* ページの高さの上限（px）/ maximum page height (px) */
var THUMBNAIL_SPREAD_GAP = 8;              /* 見開きどうしの間隔（px）/ gap between spreads (px) */
var THUMBNAIL_ELLIPSIS_GAP = 14;           /* 省いたページの位置にあける間隔（px）/ gap where pages are left out (px) */
var THUMBNAIL_MAX_INSERTED_PAGES = 16;     /* これより多く挿入するときは、前後8ページずつだけ描く / beyond this, draw only the first and last 8 inserted pages */

// UI の明暗（再利用パーツ） / UI theme (reusable)

/**
 * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
 * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
 */
function isDarkUI() {
    try {
        if (app.preferences && app.preferences.getRealPreference) {
            return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
        }
        return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
    } catch (e) {
        return false;
    }
}

// UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

// ステップボタン（再利用パーツ） / Stepper buttons (reusable)

// -----------------------------------------
// ステップボタンの寸法・増減量 / Stepper metrics and steps
// -----------------------------------------
var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

/* 単位の換算は UnitValue に任せる（in / ft / yd / mm / cm / m / pt / pc / px ほか、単数形・複数形も可）。
   UnitValue に無い単位だけ、ここで UnitValue の単位に読み替える（値は「1単位＝何 unit か」）。
   「p」は「1p6」（1パイカ6ポイント）の形にも使う
   Units UnitValue lacks, mapped onto UnitValue units (how many of `unit` make one) */
var STEPPER_UNIT_ALIASES = {
    "q": { unit: "mm", amount: 0.25 }, /* 級 / Q */
    "h": { unit: "mm", amount: 0.25 }, /* 歯 / H */
    "p": { unit: "pc", amount: 1 }     /* パイカ / pica */
};

// -----------------------------------------
// ステップボタンの配色 / Stepper colors
// -----------------------------------------
var STEPPER_UI_DARK           = isDarkUI();
/* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
   ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
   UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
   dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

// -----------------------------------------
// 数値欄を作る（外から呼ぶ関数） / Public API
// -----------------------------------------
/**
 * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
 * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
 * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す。
 * 四則演算（+ - * / と括弧）を入れると、確定時に計算した値にする。欄と異なる単位で入れた値は欄の単位へ換算する（mm の欄に「1 in」→「25.4 mm」）
 * @param {Group|Panel} parent - 追加先
 * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
 *     step / min / max / integer（true で整数のみ）/ unit / onStep
 * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
 */
function addSteppedField(parent, fieldOptions) {
    var fieldRowGroup = parent.add("group");
    fieldRowGroup.orientation = "row";
    fieldRowGroup.alignChildren = ["left", "center"];
    fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

    var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
    if (fieldOptions.labelWidth) {
        fieldLabel.preferredSize.width = fieldOptions.labelWidth;
        fieldLabel.justify = "right";
    }

    /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
    var stepperInputGroup = fieldRowGroup.add("group");
    stepperInputGroup.orientation = "row";
    stepperInputGroup.alignChildren = ["left", "center"];
    stepperInputGroup.spacing = 0;
    stepperInputGroup.margins = 0;

    var numberInput;
    var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
    numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
    numberInput.characters = fieldOptions.characters || 6;
    numberInput.fieldLabel = fieldLabel;
    numberInput.stepperGroup = stepperGroup;

    /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
    bindSteppedArrowKeys(numberInput, stepperGroup);

    /* 直接入力をそろえる。計算式は計算し、数値でなければ直前の値に戻す / normalize typed values; evaluate arithmetic, revert non-numbers */
    numberInput.lastValidText = numberInput.text;
    numberInput.onChange = function () {
        var value = evaluateArithmetic(numberInput.text, fieldOptions.unit);
        if (isNaN(value)) {
            numberInput.text = numberInput.lastValidText;
            return;
        }
        writeSteppedValue(numberInput, value, fieldOptions);
    };
    return numberInput;
}

/**
 * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
 * @param {EditText} numberInput - addSteppedField() で作った入力欄
 * @param {boolean} isEnabled - 有効にするなら true
 * @returns {void}
 */
function setSteppedFieldEnabled(numberInput, isEnabled) {
    numberInput.enabled = isEnabled;
    numberInput.fieldLabel.enabled = isEnabled;
    numberInput.stepperGroup.enabled = isEnabled;
    /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
    for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
        redrawStepperGroup(numberInput.stepperGroup.children[i]);
    }
}

/**
 * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
 * @param {Group|Panel} parent - 追加先
 * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
 * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
 * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
 */
function addStepper(parent, getNumberInput, stepOptions) {
    var stepperGroup = parent.add("group");
    stepperGroup.orientation = "column";
    stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
    stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
    stepperGroup.alignment = ["left", "center"];

    /**
     * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
     * @param {number} direction - 増やすなら 1、減らすなら -1
     * @returns {void}
     */
    function stepBy(direction) {
        var numberInput = getNumberInput();
        if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
        var value = evaluateArithmetic(numberInput.text, stepOptions.unit); /* 確定前の計算式も計算してから増減 / evaluate an uncommitted expression first */
        if (isNaN(value)) value = parseFloat(numberInput.text); /* 計算できなければ従来どおり先頭の数値 / fall back to the leading number */
        if (isNaN(value)) value = 0;
        writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
        if (stepOptions.onStep) stepOptions.onStep(numberInput);
    }

    /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
    var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
    var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
    makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
    makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
    stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
    stepperGroup.stepOptions = stepOptions; /* 確定時の計算で欄の単位を引けるよう公開 / lets the commit-time evaluation find the unit */
    return stepperGroup;
}

/**
 * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し。
 * あわせて、確定時に計算式・単位付きの値を計算して書き戻す（各スクリプトの onChange より先に呼ばれるので、onChange は計算後の値を読む）
 * @param {EditText} numberInput - 対象の入力欄
 * @param {Group} stepperGroup - addStepper() で作った∧∨
 * @returns {void}
 */
function bindSteppedArrowKeys(numberInput, stepperGroup) {
    numberInput.addEventListener("keydown", function (event) {
        if (event.keyName !== "Up" && event.keyName !== "Down") return;
        stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
        event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
    });
    numberInput.addEventListener("change", function () {
        var fieldUnit = stepperGroup.stepOptions ? stepperGroup.stepOptions.unit : undefined;
        var value = evaluateArithmetic(numberInput.text, fieldUnit);
        if (isNaN(value)) return; /* 計算できなければ各スクリプトの処理に任せる / leave it to the script's own handler */
        /* 式か、換算で値が変わったときだけ書き戻す（ただの数値は書式を崩さない） / rewrite only expressions and converted values */
        var hasOperator = /[*\/()\u00D7\u00F7\uFF0A\uFF0F\uFF08\uFF09]|[\d.\uFF10-\uFF19][^\d.\uFF10-\uFF19]*[+\-\u2212\uFF0B\uFF0D]/.test(numberInput.text);
        if (!hasOperator && value === parseFloat(numberInput.text)) return;
        numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
    });
}

// -----------------------------------------
// 値の計算 / Value helpers
// -----------------------------------------
/**
 * 押された修飾キーに応じて、1回分増減した値を返す
 * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
 * 整数の欄では option を無視して step の倍数へ）
 * @param {number} value - 元の値
 * @param {number} direction - 増やすなら 1、減らすなら -1
 * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
 * @returns {number} 増減した値（下限・上限は未適用）
 */
function computeSteppedValue(value, direction, stepOptions) {
    var keyState = ScriptUI.environment.keyboardState;
    if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
    if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
    return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
}

/**
 * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
 * @param {number} value - 元の値
 * @param {number} multiple - 倍数の単位（例 10）
 * @param {number} direction - 上げるなら 1、下げるなら -1
 * @returns {number} 移した値
 */
function snapStepperToNextMultiple(value, multiple, direction) {
    /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
       round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
    var quotient = Math.round(value / multiple * 1e6) / 1e6;
    if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
    return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
}

/**
 * 値を下限・上限の範囲に収める
 * @param {number} value - 数値
 * @param {Object} rangeOptions - min / max（どちらも省略可）
 * @returns {number} 範囲に収めた値
 */
function clampSteppedValue(value, rangeOptions) {
    if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
    if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
    return value;
}

/**
 * 入力欄の文字列を四則演算（+ - * / と括弧）として計算する。eval は使わない。
 * 数値の後ろの単位は UnitValue で欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
 * 全角の数字・記号と × ÷ は半角に直す
 * @param {string} text - 入力欄の文字列
 * @param {string} [fieldUnit] - 欄の単位（例 " mm"。前後の空白は無視）
 * @returns {number} 欄の単位での計算結果（式として読めない・換算できない単位・0で割ったときは NaN）
 */
function evaluateArithmetic(text, fieldUnit) {
    var source = String(text)
        .replace(/[！-～]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
        .replace(/×/g, "*")
        .replace(/÷/g, "/")
        .replace(/[−–—]/g, "-")
        .replace(/\s/g, "");
    if (source === "") return NaN;
    var fieldUnitKey = String(fieldUnit || "").replace(/^\s+|\s+$/g, "").toLowerCase();
    var fieldUnitValue = createStepperUnitValue(1, fieldUnitKey); /* 欄の単位の1単位（換算できない欄は null） / one field unit */
    var position = 0;

    /**
     * 加減算の並び（項 ± 項 …）を読む
     * @returns {number} 値（読めなければ NaN）
     */
    function readSum() {
        var total = readProduct();
        while (position < source.length && (source.charAt(position) === "+" || source.charAt(position) === "-")) {
            var operator = source.charAt(position++);
            var operand = readProduct();
            total = (operator === "+") ? total + operand : total - operand;
        }
        return total;
    }

    /**
     * 乗除算の並び（因子 × 因子 …）を読む
     * @returns {number} 値（読めなければ NaN）
     */
    function readProduct() {
        var total = readFactor();
        while (position < source.length && (source.charAt(position) === "*" || source.charAt(position) === "/")) {
            var operator = source.charAt(position++);
            var operand = readFactor();
            if (operator === "/" && operand === 0) return NaN;
            total = (operator === "*") ? total * operand : total / operand;
        }
        return total;
    }

    /**
     * 符号付きの数値（単位付きなら欄の単位へ換算）か、括弧で囲んだ式を読む
     * @returns {number} 値（読めなければ NaN）
     */
    function readFactor() {
        var ch = source.charAt(position);
        if (ch === "+" || ch === "-") {
            position++;
            var signedValue = readFactor();
            return (ch === "-") ? -signedValue : signedValue;
        }
        if (ch === "(") {
            position++;
            var innerValue = readSum();
            if (source.charAt(position) !== ")") return NaN;
            position++;
            return innerValue;
        }
        var numberMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
        if (!numberMatch) return NaN;
        position += numberMatch[0].length;
        return readUnitSuffix(parseFloat(numberMatch[0]));
    }

    /**
     * 数値の直後の単位を読み、欄の単位へ換算する
     * @param {number} value - 単位の前の数値
     * @returns {number} 欄の単位での値（換算できない単位なら NaN）
     */
    function readUnitSuffix(value) {
        var unitMatch = /^([A-Za-z]+|%|°)/.exec(source.substring(position));
        if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
        position += unitMatch[0].length;
        var unitKey = unitMatch[0].toLowerCase();
        if (unitKey === fieldUnitKey) return value;
        var typedValue = createStepperUnitValue(value, unitKey);
        if (!typedValue || !fieldUnitValue) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
        var points = typedValue.as("pt");
        /* 「1p6」＝1パイカ6ポイント / pica-point notation */
        if (unitKey === "p") {
            var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
            if (pointMatch) {
                position += pointMatch[0].length;
                points += parseFloat(pointMatch[0]);
            }
        }
        return points / fieldUnitValue.as("pt");
    }

    var result = readSum();
    if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
    return result;
}

/**
 * 数値と単位から UnitValue を作る。Q・H・p は STEPPER_UNIT_ALIASES で UnitValue の単位に読み替える。
 * %（percent）は基準の長さが無いと換算できないので扱わない
 * @param {number} value - 数値
 * @param {string} unitKey - 単位（小文字。例 "mm"、"inches"、"q"）
 * @returns {UnitValue|null} UnitValue（UnitValue が知らない単位・空・% なら null）
 */
function createStepperUnitValue(value, unitKey) {
    if (unitKey === "" || unitKey === "%") return null;
    var alias = STEPPER_UNIT_ALIASES[unitKey];
    var unitValue = alias ? new UnitValue(value * alias.amount, alias.unit) : new UnitValue(value, unitKey);
    if (unitValue.type === "?" || unitValue.type === "%") return null; /* 知らない単位は例外にならず "?" になる。"percent" も除く / unknown units become "?" */
    return unitValue;
}

/**
 * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
 * @param {EditText} numberInput - 書き込む入力欄
 * @param {number} value - 数値
 * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
 * @returns {void}
 */
function writeSteppedValue(numberInput, value, valueOptions) {
    numberInput.text = formatSteppedValue(value, valueOptions);
    numberInput.lastValidText = numberInput.text;
}

/**
 * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
 * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
 * @param {number} value - 数値
 * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
 * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
 */
function formatSteppedValue(value, valueOptions) {
    if (valueOptions.integer) value = Math.round(value);
    return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
}

/**
 * 小数第2位で丸めた数値を文字列で返す
 * @param {number} value - 数値
 * @returns {string} 表示用の数値文字列
 */
function formatStepperNumber(value) {
    return String(Math.round(value * 100) / 100);
}

// -----------------------------------------
// ∧∨ボタンの描画 / Drawing
// -----------------------------------------
/**
 * 山形（∧／∨）の極小ボタンを作成する。
 * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
 * 継ぎ目に線は引かない
 * @param {Group|Panel} parent - 追加先
 * @param {string} direction - "up" または "down"
 * @param {Function} onClickFn - クリック時の処理
 * @returns {Group} ボタンとして使う group
 */
function makeStepperChevronButton(parent, direction, onClickFn) {
    var buttonWidth = STEPPER_BUTTON_WIDTH;
    var buttonHeight = STEPPER_BUTTON_HEIGHT;
    var isUp = (direction === "up");
    var chevronBox = parent.add("group");
    chevronBox.margins = 0;
    chevronBox.spacing = 0;
    chevronBox.preferredSize = [buttonWidth, buttonHeight];
    chevronBox.minimumSize = [buttonWidth, buttonHeight];
    chevronBox.maximumSize = [buttonWidth, buttonHeight];
    chevronBox.isPressed = false;
    chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

    chevronBox.onDraw = function () {
        var boxGraphics = chevronBox.graphics;
        /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
           Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
        var isDimmed = !isStepperEnabledInTree(chevronBox);

        /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
        var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
        boxGraphics.newPath();
        boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
        boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

        drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
        drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
    };

    /**
     * 押下状態を変えて描き直す
     * @param {boolean} isPressed - 押下中なら true
     * @returns {void}
     */
    function repaint(isPressed) {
        if (chevronBox.isPressed === isPressed) return;
        chevronBox.isPressed = isPressed;
        redrawStepperGroup(chevronBox);
    }
    chevronBox.addEventListener("mousedown", function () {
        if (!isStepperEnabledInTree(chevronBox)) return;
        repaint(true);
        if (onClickFn) onClickFn();
    });
    chevronBox.addEventListener("mouseup", function () { repaint(false); });
    /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
    chevronBox.addEventListener("mouseout", function () { repaint(false); });
    return chevronBox;
}

/**
 * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
 * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
 * @param {ScriptUIGraphics} boxGraphics - 描画先
 * @param {number} boxWidth - ボタンの幅
 * @param {number} boxHeight - ボタンの高さ
 * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
 * @param {number[]} frameColor - [r, g, b, a]
 * @returns {void}
 */
function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
    var frameLeft = 0.5;
    var frameRight = boxWidth - 0.5;
    var outerY = isUp ? 0.5 : boxHeight - 0.5;
    var seamY = isUp ? boxHeight : 0;
    var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
    var radius = STEPPER_CORNER_RADIUS;
    var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
    var angle, k;

    boxGraphics.newPath();
    boxGraphics.moveTo(frameLeft, seamY);
    /* 左の角丸 / left corner */
    for (k = 0; k <= arcSteps; k++) {
        angle = (Math.PI / 2) * k / arcSteps;
        boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
    }
    /* 右の角丸 / right corner */
    for (k = 0; k <= arcSteps; k++) {
        angle = (Math.PI / 2) * k / arcSteps;
        boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
    }
    boxGraphics.lineTo(frameRight, seamY);
    boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
}

/**
 * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
 * @param {ScriptUIGraphics} boxGraphics - 描画先
 * @param {number} boxWidth - ボタンの幅
 * @param {number} boxHeight - ボタンの高さ
 * @param {boolean} isUp - ∧なら true、∨なら false
 * @param {number[]} chevronColor - [r, g, b, a]
 * @returns {void}
 */
function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
    var centerX = boxWidth / 2;
    var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
    var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
    var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
    boxGraphics.newPath();
    boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
    boxGraphics.lineTo(centerX, centerY + tipOffsetY);
    boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
    boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
}

/**
 * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
 * @param {Object} control - 対象のコントロール
 * @returns {boolean} すべて有効なら true
 */
function isStepperEnabledInTree(control) {
    for (var node = control; node; node = node.parent) {
        if (!node.enabled) return false;
    }
    return true;
}

/**
 * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
 * @param {Object} container - 行・グループ・パネルなど
 * @returns {void}
 */
function redrawSteppersIn(container) {
    if (!container.children) return;
    for (var i = 0; i < container.children.length; i++) {
        var child = container.children[i];
        if (child.isStepperButton) redrawStepperGroup(child);
        else redrawSteppersIn(child);
    }
}

/**
 * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
 * @param {Group} targetGroup - 描き直す group
 * @returns {void}
 */
function redrawStepperGroup(targetGroup) {
    targetGroup.hide();
    targetGroup.show();
}

// ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

// =========================================
// ラベル定義 / Labels
// =========================================

// ローカライズ（再利用パーツ） / Localization (reusable)

/**
 * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
 * @returns {string} "ja" または "en"
 */
function getCurrentLang() {
    return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
}

var uiLang = getCurrentLang();

/**
 * LABELS から今の UI 言語の文言を取り出す。
 * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
 * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
 * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
 */
function getLabel(labelRef, placeholderValues) {
    var labelEntry = labelRef;
    if (typeof labelRef === "string") {
        var labelPathKeys = labelRef.split(".");
        labelEntry = LABELS;
        for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
            labelEntry = labelEntry[labelPathKeys[i]];
        }
    }
    var labelString;
    if (typeof labelEntry === "string") labelString = labelEntry;
    else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
    else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
    else return (typeof labelRef === "string") ? labelRef : "";
    return fillLabelPlaceholders(String(labelString), placeholderValues);
}

/**
 * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
 * @param {string|Object} labelRef - getLabel と同じ
 * @param {Object|Array} [placeholderValues] - getLabel と同じ
 * @returns {string} コロン付きの文言
 */
function labelText(labelRef, placeholderValues) {
    return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
}

/**
 * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
 * @param {string|Object} labelRef - getLabel と同じ
 * @param {string|number} value - コロンのあとに続ける値
 * @returns {string} 項目名と値をつないだ文字列
 */
function labelValueText(labelRef, value) {
    return labelText(labelRef) + " " + value;
}

/**
 * 文言の {name} や %1 に値を差し込む
 * @param {string} labelString - 文言
 * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
 * @returns {string} 差し込んだ文言
 */
function fillLabelPlaceholders(labelString, placeholderValues) {
    if (placeholderValues == null) return labelString;
    if (placeholderValues instanceof Array) {
        /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
        for (var i = placeholderValues.length; i >= 1; i--) {
            labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
        }
        return labelString;
    }
    for (var placeholderKey in placeholderValues) {
        if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
        labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
    }
    return labelString;
}

// ローカライズ（再利用パーツ）ここまで / End of the reusable localization

// ボタン行（再利用パーツ） / Button row (reusable)

var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

/**
 * ダイアログ下部のボタン行を作る。
 * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
 * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
 * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
 * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
 */
function addButtonRow(parent, rowOptions) {
    var isCentered = !!(rowOptions && rowOptions.centered);
    var btnRowGroup = parent.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
    btnRowGroup.spacing = BUTTON_ROW_SPACING;

    if (isCentered) {
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
    }

    btnRowGroup.alignment = ["fill", "bottom"];

    var btnLeftGroup = btnRowGroup.add("group");
    btnLeftGroup.alignChildren = ["left", "center"];
    btnLeftGroup.spacing = BUTTON_ROW_SPACING;

    /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
    var spacer = btnRowGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 0;

    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.alignChildren = ["right", "center"];
    btnRightGroup.spacing = BUTTON_ROW_SPACING;

    return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
}

/**
 * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
 * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
 * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
 * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
 * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
 * @returns {void}
 */
function alignRightOnlyButtonRow(buttonRow) {
    if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
    var dialogWindow = buttonRow.rowGroup.window;
    dialogWindow.addEventListener("show", function () {
        if (!buttonRow.leftGroup) return;
        var btnRowGroup = buttonRow.rowGroup;
        /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
        if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
        /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
        btnRowGroup.remove(buttonRow.leftGroup);
        btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        buttonRow.leftGroup = null;
        dialogWindow.layout.layout(true);
    });
}

// ボタン行（再利用パーツ）ここまで / End of the reusable button row

// バージョン表示（再利用パーツ） / Version footer (reusable)

var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

/* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
   9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
   / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
     nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
var VERSION_FOOTER_GLYPHS = {
    "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
    "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
    "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
    "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
    "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
    "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
    "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
    "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
    "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
    "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
    "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
    "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
    "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
    "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
    "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
    "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
    "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
    "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
    "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
    "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
    "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
    "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
    "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
    "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
    "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
    "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
    "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
    "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
    "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
    "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
    "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
    "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
    "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
    "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
    "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
    "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
    "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
    "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
    "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
    "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
    "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
    "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
    "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
    "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
    "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
    "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
    "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
    "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
    "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
    "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
    "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
    "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
    "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
    "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
    "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
    "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
    "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
    "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
    "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
    "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
    "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
    "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
    ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
    "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
    "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
    "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
    " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
};

/**
 * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPT_VERSION）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
 * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
 * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
 * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
 * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
 */
function addVersionFooter(targetWindow) {
    if (targetWindow.versionFooter) return null;

    /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
    var lastChild = targetWindow.children[targetWindow.children.length - 1];
    if (lastChild && lastChild.type === "group") {
        var rowMargins = lastChild.margins;
        lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
    }
    var windowMargins = targetWindow.margins;
    targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

    var footerGroup = targetWindow.add("group");
    footerGroup.orientation = "row";
    footerGroup.alignment = ["fill", "bottom"];
    footerGroup.alignChildren = ["left", "center"];
    footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

    /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
       左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
       / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
         or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
    var versionText = String(SCRIPT_VERSION);
    var leftCandidates = [versionText];
    if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
    var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
    versionCanvas.alignment = ["fill", "center"];
    bindVersionInfoClick(versionCanvas); /* クリックで「このスクリプトについて」を開く / click opens About This Script */
    var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
    brandCanvas.alignment = ["right", "center"];
    brandCanvas.maximumSize = brandCanvas.preferredSize;
    targetWindow.versionFooter = footerGroup;
    return versionCanvas;
}

/**
 * 文字列のすべての文字に字形があるかを返す
 * @param {string} dotText - 文字列
 * @returns {boolean} すべてあれば true
 */
function hasVersionFooterGlyphs(dotText) {
    for (var i = 0; i < dotText.length; i++) {
        if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
    }
    return dotText.length > 0;
}

/**
 * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
 * @param {Group} parentGroup - 足す先
 * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
 * @returns {Group} 描くグループ
 */
function addVersionFooterCanvas(parentGroup, dotTexts) {
    var dotCanvas = parentGroup.add("group");
    var shortestText = dotTexts[dotTexts.length - 1];
    dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
    dotCanvas.helpTip = dotTexts[0];
    dotCanvas.onDraw = function () {
        var availableWidth = this.size ? this.size.width : 0;
        var drawText = shortestText;
        for (var i = 0; i < dotTexts.length; i++) {
            if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                drawText = dotTexts[i];
                break;
            }
        }
        drawVersionFooterDots(this.graphics, drawText);
    };
    return dotCanvas;
}

/**
 * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
 * @param {string} versionText - 文字列
 * @returns {number} 幅（ドット）
 */
function measureVersionFooterDots(versionText) {
    var dotCount = 0;
    for (var i = 0; i < versionText.length; i++) {
        var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
        if (!glyphRows) continue;
        if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
        dotCount += glyphRows[0].length;
    }
    return dotCount;
}

/**
 * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
 * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
 * @param {string} versionText - 文字列
 * @returns {void}
 */
function drawVersionFooterDots(canvasGraphics, versionText) {
    var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
    var dotSize = VERSION_FOOTER_PIXEL_SIZE;
    var cursorDots = 0;
    canvasGraphics.newPath();
    for (var i = 0; i < versionText.length; i++) {
        var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
        if (!glyphRows) continue;
        if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
        for (var row = 0; row < glyphRows.length; row++) {
            var rowText = glyphRows[row];
            var runStart = -1;
            for (var col = 0; col <= rowText.length; col++) {
                var isDot = col < rowText.length && rowText.charAt(col) === "#";
                if (isDot && runStart < 0) runStart = col;
                if (!isDot && runStart >= 0) {
                    canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                    runStart = -1;
                }
            }
        }
        cursorDots += glyphRows[0].length;
    }
    canvasGraphics.fillPath(footerBrush);
}

// -----------------------------------------
// このスクリプトについて / About this script
// -----------------------------------------
/* 左下のスクリプト名・バージョンをクリックすると開く。基本情報ブロックの SCRIPT_* を読み、
   この部品だけで動くよう、文言（VERSION_INFO_LABELS）とレイアウトの値も部品の中に持つ
   Opened by clicking the name and version at the bottom left. Reads the SCRIPT_* basic info,
   and keeps its own labels and layout values so the part works on its own */

var VERSION_INFO_LANG = (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en"; /* 文言の言語 / label language */
var VERSION_INFO_IS_INDESIGN = (String(app.name).indexOf("InDesign") >= 0);         /* InDesign で動いているか / running in InDesign */
/* 実行中のスクリプトのファイル。ボタンを押すころには $.fileName が別のものを指すことがあるので、読み込んだ時点で控える
   The running script file, captured at load time; $.fileName may point elsewhere by the time the button is clicked */
var VERSION_INFO_SCRIPT_FILE = new File($.fileName);
var VERSION_INFO_REVEAL_APP_PATH = "/Applications/RevealInFinder.app";   /* Finder で選択表示する補助アプリ / helper app that reveals a file in the Finder */
var VERSION_INFO_REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";  /* 補助アプリが読むパス（アプリの旧名のまま。変えると読めない）/ path the helper reads (keeps the app's old name) */

var VERSION_INFO_WINDOW_MARGINS = 16;               /* ウィンドウの余白 / window margin */
var VERSION_INFO_WINDOW_SPACING = 12;               /* ウィンドウ内の間隔 / window spacing */
var VERSION_INFO_PANEL_MARGINS = [16, 20, 16, 12];  /* パネルの余白 [左,上,右,下] / panel margins */
var VERSION_INFO_ROW_SPACING = 6;                   /* 項目名と値の間隔 / gap between a label and its value */
var VERSION_INFO_BUTTON_SPACING = 10;               /* ボタンどうしの間隔 / gap between buttons */
var VERSION_INFO_BUTTON_ROW_MARGINS = [0, 5, 0, 14]; /* ボタン行の余白（Illustrator 標準のダイアログに合わせる）/ button row margins */
var VERSION_INFO_COPY_BUTTON_SIZE = 22;             /* コピーボタンの一辺 / copy button size */
/* コピーボタンの配色。∧∨（ステップボタン）と同じく、背景に重ねる半透明 / Copy button colors, translucent overlays like the stepper */
var VERSION_INFO_COPY_FILL_COLOR    = isDarkUI() ? [0, 0, 0, 0.10] : [1, 1, 1, 0.50]; /* 地 / background */
var VERSION_INFO_COPY_FRAME_COLOR   = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.15]; /* 枠線 / frame */
var VERSION_INFO_COPY_PRESSED_COLOR = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 押下中 / pressed */
var VERSION_INFO_COPY_ICON_COLOR    = isDarkUI() ? [1, 1, 1, 0.85] : [0, 0, 0, 0.70]; /* アイコン / icon */

var VERSION_INFO_LABELS = {
    dialog: {
        title: { ja: "このスクリプトについて", en: "About This Script" }
    },
    panel: {
        generalInfo: { ja: "基本情報", en: "General" },
        links: { ja: "リンク", en: "Links" }
    },
    fieldLabel: {
        scriptName: { ja: "スクリプト名", en: "Script Name" },
        scriptFeature: { ja: "機能", en: "Feature" },
        scriptVersion: { ja: "バージョン", en: "Version" }
    },
    button: {
        openReadme: { ja: "README", en: "README" },
        openSourceCode: { ja: "ソースコード", en: "Source Code" },
        openArticle: { ja: "記事（note）", en: "Article (note)" },
        revealInFinder: { ja: "Finder に表示", en: "Show in Finder" },
        revealInExplorer: { ja: "エクスプローラーに表示", en: "Show in Explorer" },
        close: { ja: "閉じる", en: "Close" }
    },
    tooltip: {
        versionInfo: { ja: "クリックで、このスクリプトについての情報を表示", en: "Click for info about this script" },
        copyScriptName: { ja: "スクリプト名をコピー", en: "Copy Script Name" }
    },
    alert: {
        copyFailed: { ja: "スクリプト名をコピーできませんでした。", en: "Could not copy the script name." }
    }
};

/**
 * VERSION_INFO_LABELS の文言を今の言語で返す
 * @param {Object} labelEntry - { ja, en }
 * @returns {string} 文言
 */
function getVersionInfoLabel(labelEntry) {
    return labelEntry[VERSION_INFO_LANG] || labelEntry.en;
}

/**
 * 基本情報ブロックの値を返す（変数が無い・文字列でなければ空文字）。typeof で読むので、未定義でも例外にならない
 * @param {string} variableKey - "name" / "version" / "feature" / "readme" / "source" / "article"
 * @returns {string} 値
 */
function getVersionInfoValue(variableKey) {
    var infoValue;
    if (variableKey === "name") infoValue = (typeof SCRIPT_NAME !== "undefined") ? SCRIPT_NAME : "";
    else if (variableKey === "version") infoValue = (typeof SCRIPT_VERSION !== "undefined") ? SCRIPT_VERSION : "";
    else if (variableKey === "feature") infoValue = (typeof SCRIPT_DESCRIPTION_JA !== "undefined") ? SCRIPT_DESCRIPTION_JA : "";
    else if (variableKey === "readme") {
        if (VERSION_INFO_LANG === "ja") infoValue = (typeof SCRIPT_README_JA !== "undefined") ? SCRIPT_README_JA : "";
        else infoValue = (typeof SCRIPT_README_EN !== "undefined") ? SCRIPT_README_EN : "";
    }
    else if (variableKey === "source") infoValue = (typeof SCRIPT_SOURCE_URL !== "undefined") ? SCRIPT_SOURCE_URL : "";
    else if (variableKey === "article") infoValue = (typeof SCRIPT_ARTICLE_URL !== "undefined") ? SCRIPT_ARTICLE_URL : "";
    return (infoValue === undefined || infoValue === null) ? "" : String(infoValue);
}

/**
 * 左下のスクリプト名・バージョンをクリックしたら、「このスクリプトについて」のダイアログを開くようにする
 * @param {Group} versionCanvas - addVersionFooterCanvas() で作った左の欄
 * @returns {void}
 */
function bindVersionInfoClick(versionCanvas) {
    versionCanvas.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.versionInfo);
    /* mousedown で開くと、離したときのイベントが閉じたあとの欄に届くので mouseup で開く
       Open on mouseup so the release does not land on the footer after the dialog closes */
    versionCanvas.addEventListener("mouseup", function () {
        showVersionInfoDialog(versionCanvas.window.type === "palette");
    });
}

/**
 * 「このスクリプトについて」のダイアログを開く。スクリプト名（コピーボタン付き）・機能・バージョンの［基本情報］パネル、
 * README・ソースコード・記事を開く［リンク］パネル、［Finder に表示］［閉じる］を並べる
 * @param {boolean} isFromPalette - 常駐パレットから開いたなら true（Illustrator のコピーを BridgeTalk 経由にする）
 * @returns {void}
 */
function showVersionInfoDialog(isFromPalette) {
    var infoDialog = new Window("dialog", getVersionInfoLabel(VERSION_INFO_LABELS.dialog.title));
    infoDialog.orientation = "column";
    infoDialog.alignChildren = "fill";
    infoDialog.margins = VERSION_INFO_WINDOW_MARGINS;
    infoDialog.spacing = VERSION_INFO_WINDOW_SPACING;

    var infoPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.generalInfo);
    /* コピーボタンはスクリプト名だけ / Only the script name gets a copy button */
    addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptName, getVersionInfoValue("name"), isFromPalette);
    /* 機能は README の見出しで日本語だけなので、日本語の画面だけに出す / The feature is the Japanese README title, shown only in the Japanese UI */
    if (VERSION_INFO_LANG === "ja" && getVersionInfoValue("feature") !== "") {
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptFeature, getVersionInfoValue("feature"), null);
    }
    addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptVersion, getVersionInfoValue("version"), null);
    alignVersionInfoLabels(infoPanel);

    /* URL が無いボタンはディム表示にして、並びはスクリプトごとに変えない / Dim buttons without a URL; keep the same buttons everywhere */
    var linksPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.links);
    var linkButtonRow = linksPanel.add("group");
    linkButtonRow.orientation = "row";
    linkButtonRow.alignment = ["left", "center"];
    linkButtonRow.alignChildren = ["left", "center"];
    linkButtonRow.spacing = VERSION_INFO_BUTTON_SPACING;
    addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openReadme, getVersionInfoValue("readme"));
    addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openSourceCode, getVersionInfoValue("source"));
    addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openArticle, getVersionInfoValue("article"));

    /* ボタン行：左に［Finder に表示］、右に［閉じる］ / Button row: Show in Finder at the left, Close at the right */
    var buttonRow = infoDialog.add("group");
    buttonRow.orientation = "row";
    buttonRow.alignment = ["fill", "bottom"];
    buttonRow.alignChildren = ["left", "center"];
    buttonRow.margins = VERSION_INFO_BUTTON_ROW_MARGINS;
    buttonRow.spacing = VERSION_INFO_BUTTON_SPACING;
    var revealLabel = (File.fs === "Windows") ? VERSION_INFO_LABELS.button.revealInExplorer : VERSION_INFO_LABELS.button.revealInFinder;
    var revealButton = buttonRow.add("button", undefined, getVersionInfoLabel(revealLabel));
    if (VERSION_INFO_SCRIPT_FILE.exists) {
        revealButton.helpTip = VERSION_INFO_SCRIPT_FILE.fsName;
        revealButton.onClick = function () {
            revealVersionInfoScriptFile(VERSION_INFO_SCRIPT_FILE);
            infoDialog.close();
        };
    } else {
        revealButton.enabled = false;
    }
    /* 余りの幅を吸って、［閉じる］を右端に寄せる / Absorbs the extra width so Close sits at the right edge */
    var buttonSpacer = buttonRow.add("group");
    buttonSpacer.alignment = ["fill", "fill"];
    buttonSpacer.minimumSize.width = 0;
    var closeButton = buttonRow.add("button", undefined, getVersionInfoLabel(VERSION_INFO_LABELS.button.close), { name: "ok" });
    closeButton.alignment = ["right", "center"];
    /* return でも esc でも閉じる / Close with both Return and Esc */
    infoDialog.defaultElement = closeButton;
    infoDialog.cancelElement = closeButton;
    infoDialog.show();
}

/**
 * 「このスクリプトについて」のダイアログにパネルを足す
 * @param {Window} infoDialog - 足す先のダイアログ
 * @param {Object} titleEntry - パネル名の文言 { ja, en }
 * @returns {Panel} パネル
 */
function addVersionInfoPanel(infoDialog, titleEntry) {
    var infoPanel = infoDialog.add("panel", undefined, getVersionInfoLabel(titleEntry));
    infoPanel.orientation = "column";
    infoPanel.alignChildren = ["fill", "top"];
    infoPanel.alignment = "fill";
    infoPanel.margins = VERSION_INFO_PANEL_MARGINS;
    infoPanel.spacing = VERSION_INFO_WINDOW_SPACING;
    return infoPanel;
}

/**
 * ［基本情報］パネルに「項目名 : 値」の1行を足す（値は見せるだけなので statictext）
 * @param {Panel} infoPanel - 足す先のパネル
 * @param {Object} labelEntry - 項目名の文言 { ja, en }
 * @param {string} valueText - 表示する値
 * @param {boolean|null} copyFromPalette - 値の右にコピーボタンを置くなら true / false（常駐パレットから開いたか）、置かないなら null
 * @returns {StaticText} 値の statictext
 */
function addVersionInfoRow(infoPanel, labelEntry, valueText, copyFromPalette) {
    var infoRow = infoPanel.add("group");
    infoRow.orientation = "row";
    infoRow.alignment = ["fill", "center"];
    infoRow.alignChildren = ["left", "center"];
    infoRow.spacing = VERSION_INFO_ROW_SPACING;
    /* 項目名のコロンは日本語が「 :」、英語が「:」 / Colon: " :" in Japanese, ":" in English */
    var rowLabel = infoRow.add("statictext", undefined, getVersionInfoLabel(labelEntry) + (VERSION_INFO_LANG === "ja" ? " :" : ":"));
    rowLabel.justify = "right";
    var valueLabel = infoRow.add("statictext", undefined, valueText);
    /* 余りの幅を取り、コピーボタンを右端に置く / Take the extra width so the copy button sits at the right */
    valueLabel.alignment = ["fill", "center"];
    if (copyFromPalette !== null) addVersionInfoCopyButton(infoRow, valueLabel, copyFromPalette);
    return valueLabel;
}

/**
 * ［基本情報］パネルの項目名を、いちばん長いものの幅にそろえる
 * @param {Panel} infoPanel - 対象のパネル
 * @returns {void}
 */
function alignVersionInfoLabels(infoPanel) {
    var rowLabels = [];
    var labelWidth = 0;
    for (var i = 0; i < infoPanel.children.length; i++) {
        var rowLabel = infoPanel.children[i].children[0]; /* 行の先頭は項目名 / each row starts with its label */
        rowLabels.push(rowLabel);
        labelWidth = Math.max(labelWidth, rowLabel.graphics.measureString(rowLabel.text)[0]);
    }
    for (var j = 0; j < rowLabels.length; j++) {
        rowLabels[j].preferredSize.width = Math.ceil(labelWidth) + 2;
    }
}

/**
 * URL を開くボタンを足す。URL が無ければディム表示にする。開いたらダイアログは閉じる
 * @param {Window} infoDialog - ボタンを置くダイアログ
 * @param {Group} buttonGroup - 足す先のグループ
 * @param {Object} labelEntry - ボタンの文言 { ja, en }
 * @param {string} targetUrl - 開く URL（空ならディム表示）
 * @returns {void}
 */
function addVersionInfoLinkButton(infoDialog, buttonGroup, labelEntry, targetUrl) {
    var linkButton = buttonGroup.add("button", undefined, getVersionInfoLabel(labelEntry));
    if (!targetUrl) {
        linkButton.enabled = false;
        return;
    }
    linkButton.helpTip = targetUrl;
    linkButton.onClick = function () {
        openVersionInfoUrl(targetUrl);
        infoDialog.close();
    };
}

/**
 * 値をクリップボードへコピーするアイコンのボタンを足す。コピーしたらアイコンをチェックマークに変え、ポインタが外れたら戻す
 * @param {Group} infoRow - 足す先の行
 * @param {StaticText} valueLabel - コピーする値の statictext
 * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
 * @returns {Group} ボタンとして使う group
 */
function addVersionInfoCopyButton(infoRow, valueLabel, isFromPalette) {
    var buttonSize = VERSION_INFO_COPY_BUTTON_SIZE;
    var copyButton = infoRow.add("group");
    copyButton.preferredSize = [buttonSize, buttonSize];
    copyButton.minimumSize = [buttonSize, buttonSize];
    copyButton.maximumSize = [buttonSize, buttonSize];
    copyButton.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.copyScriptName);
    copyButton.isPressed = false;
    copyButton.isCopied = false;

    copyButton.onDraw = function () {
        var buttonGraphics = copyButton.graphics;
        var fillColor = copyButton.isPressed ? VERSION_INFO_COPY_PRESSED_COLOR : VERSION_INFO_COPY_FILL_COLOR;
        buttonGraphics.newPath();
        buttonGraphics.rectPath(0, 0, buttonSize, buttonSize);
        buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, fillColor));
        buttonGraphics.newPath();
        buttonGraphics.rectPath(0.5, 0.5, buttonSize - 1, buttonSize - 1);
        buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_FRAME_COLOR, 1));
        if (copyButton.isCopied) drawVersionInfoCheckIcon(buttonGraphics, buttonSize);
        else drawVersionInfoCopyIcon(buttonGraphics, buttonSize);
    };

    /**
     * 状態を変えて描き直す（group には notify() が無いので隠して再表示する）
     * @param {boolean} isPressed - 押下中なら true
     * @param {boolean} isCopied - コピー済みなら true
     * @returns {void}
     */
    function repaint(isPressed, isCopied) {
        if (copyButton.isPressed === isPressed && copyButton.isCopied === isCopied) return;
        copyButton.isPressed = isPressed;
        copyButton.isCopied = isCopied;
        copyButton.hide();
        copyButton.show();
    }
    copyButton.addEventListener("mousedown", function () { repaint(true, false); });
    copyButton.addEventListener("mouseup", function () {
        if (!copyButton.isPressed) return;
        var hasCopied = copyVersionInfoText(valueLabel.text, isFromPalette);
        repaint(false, hasCopied);
        if (!hasCopied) alert(getVersionInfoLabel(VERSION_INFO_LABELS.alert.copyFailed));
    });
    /* 押したまま外へ出たとき・コピーのあと離れたときは元の見た目に戻す / Reset when the pointer leaves */
    copyButton.addEventListener("mouseout", function () { repaint(false, false); });
    return copyButton;
}

/**
 * コピーのアイコン（後ろの四角に前の四角を重ねた形）を描く
 * @param {ScriptUIGraphics} buttonGraphics - 描画先
 * @param {number} buttonSize - ボタンの一辺
 * @returns {void}
 */
function drawVersionInfoCopyIcon(buttonGraphics, buttonSize) {
    var iconPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1);
    var sheetSize = 9;
    var offset = 3;
    var frontLeft = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
    var frontTop = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
    var backLeft = frontLeft - offset;
    var backTop = frontTop - offset;
    /* 後ろの四角は、前の四角に隠れる部分を描かない（上と左の辺だけ）/ Draw only the visible top and left of the back sheet */
    buttonGraphics.newPath();
    buttonGraphics.moveTo(frontLeft, backTop + sheetSize);
    buttonGraphics.lineTo(backLeft, backTop + sheetSize);
    buttonGraphics.lineTo(backLeft, backTop);
    buttonGraphics.lineTo(backLeft + sheetSize, backTop);
    buttonGraphics.lineTo(backLeft + sheetSize, frontTop);
    buttonGraphics.strokePath(iconPen);
    buttonGraphics.newPath();
    buttonGraphics.rectPath(frontLeft, frontTop, sheetSize, sheetSize);
    buttonGraphics.strokePath(iconPen);
}

/**
 * コピーしたあとのチェックマークを描く
 * @param {ScriptUIGraphics} buttonGraphics - 描画先
 * @param {number} buttonSize - ボタンの一辺
 * @returns {void}
 */
function drawVersionInfoCheckIcon(buttonGraphics, buttonSize) {
    var center = buttonSize / 2;
    buttonGraphics.newPath();
    buttonGraphics.moveTo(center - 4.5, center);
    buttonGraphics.lineTo(center - 1.5, center + 3);
    buttonGraphics.lineTo(center + 4.5, center - 3.5);
    buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1.5));
}

/**
 * 文字列をクリップボードにコピーする。
 * InDesign は AppleScript（Windows は VBScript）の doScript で書き込む。
 * Illustrator は文字列を直接送る API が無いので、一時レイヤーのテキストフレームを選んでコピーする（プレーンテキストも入る）。
 * Illustrator の常駐パレットからは DOM に触れないので、同じ処理を BridgeTalk でメインエンジンに渡す
 * @param {string} copyText - コピーする文字列
 * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
 * @returns {boolean} コピーの操作まで進めたら true
 */
function copyVersionInfoText(copyText, isFromPalette) {
    try {
        if (VERSION_INFO_IS_INDESIGN) {
            if (File.fs === "Windows") {
                app.doScript('CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""') + '"', ScriptLanguage.VISUAL_BASIC);
            } else {
                app.doScript('set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"', ScriptLanguage.APPLESCRIPT_LANGUAGE);
            }
            return true;
        }
        var copyScript = buildVersionInfoCopyScript(copyText);
        if (!isFromPalette) return eval(copyScript) === true;
        var copyMessage = new BridgeTalk();
        copyMessage.target = BridgeTalk.appSpecifier; /* 複数バージョンを起動していても、今のバージョンに送る / the running version, even with several installed */
        copyMessage.body = copyScript;
        copyMessage.send();
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * Illustrator でコピーする処理を文字列で組み立てる（BridgeTalk でも送れるよう、関数の toString() は使わない）。
 * 一時レイヤーにテキストフレームを作り、フレームを選んでメニューコマンドでコピーして、レイヤーごと消す。
 * 文字として選択（textRange.select()）してコピーすると「オブジェクトをコピーできません」の警告が出るので使わない。
 * 選択と作業レイヤーは元に戻す。ドキュメントが無ければ一時ドキュメントを作って閉じる
 * @param {string} copyText - コピーする文字列
 * @returns {string} 実行すると true / false を返すスクリプト
 */
function buildVersionInfoCopyScript(copyText) {
    return [
        "(function (copyText) {",
        "    var usingTempDoc = (app.documents.length === 0);",
        "    var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;",
        "    var savedSelection = (!usingTempDoc && targetDoc.selection instanceof Array) ? targetDoc.selection : [];",
        "    var savedActiveLayer = usingTempDoc ? null : targetDoc.activeLayer;",
        "    var tempLayer = null;",
        "    var hasCopied = false;",
        "    try {",
        "        tempLayer = targetDoc.layers.add();",
        "        var tempFrame = tempLayer.textFrames.add();",
        "        tempFrame.contents = copyText;",
        "        app.redraw();",
        "        app.executeMenuCommand('deselectall');",
        "        tempFrame.selected = true;",
        "        app.redraw();",
        "        app.executeMenuCommand('copy');",
        "        app.redraw();",
        "        hasCopied = true;",
        "    } catch (e) { }",
        "    if (tempLayer) { try { tempLayer.remove(); } catch (e) { } }",
        "    if (usingTempDoc) { targetDoc.close(SaveOptions.DONOTSAVECHANGES); return hasCopied; }",
        "    try { targetDoc.activeLayer = savedActiveLayer; } catch (e) { }",
        "    targetDoc.selection = null;",
        "    for (var i = 0; i < savedSelection.length; i++) { try { savedSelection[i].selected = true; } catch (e) { } }",
        "    app.redraw();",
        "    return hasCopied;",
        "})(" + quoteVersionInfoString(copyText) + ");"
    ].join(" "); /* 行コメントを含まないので1行につなぐ（BridgeTalk の転送で改行が化けない）/ joined on one line; there are no line comments */
}

/**
 * 文字列を、スクリプトに埋め込める二重引用符のリテラルにする
 * @param {string} sourceText - 文字列
 * @returns {string} "…" の形のリテラル
 */
function quoteVersionInfoString(sourceText) {
    return '"' + String(sourceText).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r/g, "\\r").replace(/\n/g, "\\n") + '"';
}

/**
 * ファイルを Finder で表示する。RevealInFinder.app があればファイルを選択した状態で表示し、無ければ入っているフォルダーを開く
 * @param {File} scriptFile - 表示するファイル
 * @returns {void}
 */
function revealVersionInfoScriptFile(scriptFile) {
    /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
    var hasRevealApp = (File.fs === "Macintosh")
        && (new Folder(VERSION_INFO_REVEAL_APP_PATH).exists || new File(VERSION_INFO_REVEAL_APP_PATH).exists);
    /* アプリは一時ファイルからパスを読むので、書き出してから起動する / The app reads the path from a temp file */
    if (hasRevealApp && writeVersionInfoRevealPath(scriptFile.fsName) && new File(VERSION_INFO_REVEAL_APP_PATH).execute()) return;
    /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
    scriptFile.parent.execute();
}

/**
 * RevealInFinder.app に渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
 * @param {string} targetPath - 絶対パス
 * @returns {boolean} 書き出せたら true
 */
function writeVersionInfoRevealPath(targetPath) {
    var pathFile = new File(VERSION_INFO_REVEAL_PATH_FILE);
    pathFile.encoding = "UTF-8";
    pathFile.lineFeed = "Unix";
    if (!pathFile.open("w")) return false;
    var hasWritten = pathFile.write(targetPath);
    pathFile.close();
    return hasWritten;
}

/**
 * URL を既定のブラウザーで開く（一時的なショートカットファイルを作って開く）
 * @param {string} targetUrl - 開く URL
 * @returns {void}
 */
function openVersionInfoUrl(targetUrl) {
    var isMac = (File.fs === "Macintosh");
    var shortcutFile = new File(Folder.temp.fsName + "/VersionInfoLink" + (isMac ? ".webloc" : ".url"));
    shortcutFile.encoding = "UTF-8";
    shortcutFile.open("w");
    if (isMac) {
        var escapedUrl = targetUrl.replace(/&/g, "&amp;").replace(/</g, "&lt;");
        shortcutFile.write('<?xml version="1.0" encoding="UTF-8"?>\n'
            + '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
            + '<plist version="1.0"><dict><key>URL</key><string>' + escapedUrl + '</string></dict></plist>\n');
    } else {
        shortcutFile.write("[InternetShortcut]\r\nURL=" + targetUrl + "\r\n");
    }
    shortcutFile.close();
    shortcutFile.execute();
}

// バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

var LABELS = {
    dialog: {
        title: { ja: "ページを挿入", en: "Insert Pages" }
    },
    field: {
        currentPage:    { ja: "現在のページ", en: "Current page" },
        appliedMaster:  { ja: "親（マスター）", en: "Parent page" },
        pageCount:      { ja: "挿入するページ数", en: "Number of pages to insert" }
    },
    button: {
        ok:     { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    tooltip: {
        insertAfter: { ja: "指定ページの次に挿入", en: "Insert after the specified page" },
        stepUp: {
            ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
            en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
        },
        stepDown: {
            ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
            en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
        },
        stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
        stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
        thumbnail: {
            ja: "挿入後の見開きの縮小図です。青が挿入するページ、太枠が現在のページ、ページ内の文字は親ページの接頭辞です。",
            en: "Thumbnail of the spreads after insertion. Blue pages are inserted, the thick frame is the current page, and the letters are the parent page prefixes."
        }
    },
    undo: {
        insertPages: { ja: "ページを挿入", en: "Insert Pages" }
    },
    alert: {
        invalidNumber: { ja: "1以上の数値を入力してください。", en: "Please enter a number greater than 0." },
        noDocument:    { ja: "ドキュメントが開いていません。", en: "No document is open." },
        noLayoutWindow: {
            ja: "レイアウトウィンドウがアクティブではありません。ドキュメントのページを表示してから実行してください。",
            en: "No layout window is active. Switch to a document layout view and try again."
        },
        masterPageActive: {
            ja: "親（マスター）ページが表示されています。ドキュメントのページを表示してから実行してください。",
            en: "A parent (master) page is displayed. Switch to a document page and try again."
        }
    }
};

// =========================================
// 実行前のチェック / Preconditions
// =========================================

/**
 * レイアウトウィンドウがアクティブかどうかを判定する
 * @returns {boolean} レイアウトウィンドウなら true
 */
function isLayoutWindowActive() {
    try {
        return (app.activeWindow.constructor.name === "LayoutWindow");
    } catch (e) {
        return false;
    }
}

/**
 * 指定ページが親（マスター）ページかどうかを判定する
 * @param {Page} targetPage 判定するページ
 * @returns {boolean} 親ページなら true
 */
function isMasterPage(targetPage) {
    try {
        return (targetPage.parent instanceof MasterSpread);
    } catch (e) {
        return false;
    }
}

// =========================================
// 縮小図 / Spread thumbnail
// =========================================

var THUMBNAIL_UI_DARK = isDarkUI();
var THUMBNAIL_COLORS = {
    background:  THUMBNAIL_UI_DARK ? [0.16, 0.16, 0.16, 1] : [0.90, 0.90, 0.90, 1], /* 地 / background */
    page:        [1, 1, 1, 1],                                                     /* 既存のページ / existing page */
    pageFrame:   THUMBNAIL_UI_DARK ? [0.55, 0.55, 0.55, 1] : [0.50, 0.50, 0.50, 1], /* ページの枠 / page frame */
    inserted:    [0.74, 0.84, 0.97, 1],                                            /* 挿入するページ / inserted page */
    accent:      THUMBNAIL_UI_DARK ? [0.40, 0.65, 1.00, 1] : [0.10, 0.40, 0.85, 1], /* 現在のページの枠・挿入するページの枠 / current and inserted page frames */
    prefix:      [0.40, 0.40, 0.40, 1],                                            /* 親ページの接頭辞 / parent page prefix */
    label:       THUMBNAIL_UI_DARK ? [0.85, 0.85, 0.85, 1] : [0.20, 0.20, 0.20, 1]  /* 現在のページ名 / current page name */
};

// 縮小図（再利用パーツ） / Thumbnail canvas (reusable)

var thumbnailFontCache = {}; /* 文字サイズごとの書体 / fonts per size */

/**
 * 縮小図を自作描画する描画域（group）を追加する
 * @param {Window|Panel|Group} parent - 追加先
 * @param {number[]} canvasSize - [幅, 高さ]（px）。fillWidth のときの幅は最小幅
 * @param {Function} drawFn - 描くたびに drawFn(canvasGraphics, canvasWidth, canvasHeight) で呼ぶ関数
 * @param {boolean} [fillWidth] - true なら親の幅いっぱいに伸ばす（ダイアログの幅は広げない）
 * @returns {Group} 描画域
 */
function addThumbnailCanvas(parent, canvasSize, drawFn, fillWidth) {
    var thumbnailCanvas = parent.add("group");
    thumbnailCanvas.preferredSize = [canvasSize[0], canvasSize[1]];
    thumbnailCanvas.minimumSize = [canvasSize[0], canvasSize[1]];
    thumbnailCanvas.maximumSize = [fillWidth ? 10000 : canvasSize[0], canvasSize[1]];
    if (fillWidth) thumbnailCanvas.alignment = ["fill", "top"];
    thumbnailCanvas.onDraw = function () {
        var drawnSize = thumbnailCanvas.size;
        drawFn(thumbnailCanvas.graphics, drawnSize.width, drawnSize.height);
    };
    return thumbnailCanvas;
}

/**
 * 縮小図を描き直す（group には notify() が無いので、隠して再表示する）
 * @param {Group} thumbnailCanvas - addThumbnailCanvas() の戻り値
 * @returns {void}
 */
function redrawThumbnailCanvas(thumbnailCanvas) {
    thumbnailCanvas.hide();
    thumbnailCanvas.show();
}

/**
 * 縮小図に矩形を塗る（rectPath の前に newPath を呼ばないとパスが重なって線画になる）
 * @param {ScriptUIGraphics} canvasGraphics - 描画先
 * @param {number} left - 左（px）
 * @param {number} top - 上（px）
 * @param {number} width - 幅（px）
 * @param {number} height - 高さ（px）
 * @param {number[]} color - RGBA
 * @returns {void}
 */
function fillThumbnailRect(canvasGraphics, left, top, width, height, color) {
    canvasGraphics.newPath();
    canvasGraphics.rectPath(left, top, width, height);
    canvasGraphics.fillPath(canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, color));
}

/**
 * 縮小図に矩形の枠を描く（1px の線が内側に収まるよう 0.5px 内に寄せる）
 * @param {ScriptUIGraphics} canvasGraphics - 描画先
 * @param {number} left - 左（px）
 * @param {number} top - 上（px）
 * @param {number} width - 幅（px）
 * @param {number} height - 高さ（px）
 * @param {number[]} color - RGBA
 * @returns {void}
 */
function strokeThumbnailRect(canvasGraphics, left, top, width, height, color) {
    canvasGraphics.newPath();
    canvasGraphics.rectPath(left + 0.5, top + 0.5, Math.max(width - 1, 0), Math.max(height - 1, 0));
    canvasGraphics.strokePath(canvasGraphics.newPen(canvasGraphics.PenType.SOLID_COLOR, color, 1));
}

/**
 * 縮小図の文字の書体を返す（作れない環境ではコントロールの書体を使う）
 * @param {ScriptUIGraphics} canvasGraphics - 描画先
 * @param {number} fontSize - 文字サイズ（px。InDesign では無視される）
 * @returns {ScriptUIFont} 書体
 */
function getThumbnailFont(canvasGraphics, fontSize) {
    if (thumbnailFontCache[fontSize]) return thumbnailFontCache[fontSize];
    try {
        thumbnailFontCache[fontSize] = ScriptUI.newFont(canvasGraphics.font.name, "REGULAR", fontSize);
    } catch (e) {
        thumbnailFontCache[fontSize] = canvasGraphics.font;
    }
    return thumbnailFontCache[fontSize];
}

/**
 * 文字列の幅と高さを測る（measureString は配列か {width, height} で返る）
 * @param {ScriptUIGraphics} canvasGraphics - 描画先
 * @param {string} text - 文字列
 * @param {ScriptUIFont} font - 書体
 * @returns {number[]} [幅, 高さ]（px）
 */
function measureThumbnailText(canvasGraphics, text, font) {
    var measured = canvasGraphics.measureString(text, font);
    if (measured.width !== undefined) return [measured.width, measured.height];
    return [measured[0], measured[1]];
}

// 縮小図（再利用パーツ）ここまで / End of the reusable thumbnail canvas

/**
 * 親ページの接頭辞を返す（［なし］や読めないときは空文字）
 * @param {Page} targetPage - ページ
 * @returns {string} 接頭辞
 */
function getAppliedMasterPrefix(targetPage) {
    try {
        var appliedMaster = targetPage.appliedMaster;
        return (appliedMaster && appliedMaster.isValid) ? appliedMaster.namePrefix : "";
    } catch (e) {
        return "";
    }
}

/**
 * 縮小図に使う、ドキュメントのページ構成を読む（ダイアログの前に1回だけ。ページごとの接頭辞は描くときに読んで控える）
 * @param {Document} targetDocument - ドキュメント
 * @param {Page} currentPage - 現在のページ
 * @param {MasterSpread|null} appliedMaster - 挿入するページに適用する親ページ
 * @returns {Object} ページ数・現在のページの位置・見開きと綴じ方・ページの縦横比・接頭辞の控え
 */
function collectSpreadContext(targetDocument, currentPage, appliedMaster) {
    var documentPreferences = targetDocument.documentPreferences;
    var isRightToLeft = (documentPreferences.pageBinding === PageBindingOptions.RIGHT_TO_LEFT);
    var isFacing = documentPreferences.facingPages;
    /* 綴じ方の側から見て表（奇数）側に来るページから始まるときは、1ページ目が単独になる
       When the first page sits on the recto side, it stands alone in the first spread */
    var rectoSide = isRightToLeft ? PageSideOptions.LEFT_HAND : PageSideOptions.RIGHT_HAND;
    var firstPageIsRecto = isFacing && targetDocument.pages[0].side === rectoSide;
    return {
        documentPages: targetDocument.pages,
        pageCount: targetDocument.pages.length,
        currentIndex: currentPage.documentOffset,
        currentName: currentPage.name,
        isFacing: isFacing,
        isRightToLeft: isRightToLeft,
        spreadOffset: firstPageIsRecto ? 1 : 0,
        pageAspect: documentPreferences.pageWidth / documentPreferences.pageHeight,
        insertedPrefix: appliedMaster ? appliedMaster.namePrefix : "",
        prefixCache: {}
    };
}

/**
 * 挿入後のページ番号（0始まり）が何番目の見開きに入るか（見開きでなければページごと）
 * @param {Object} spreadContext - collectSpreadContext() の戻り値
 * @param {number} resultIndex - 挿入後のページの位置（0始まり）
 * @returns {number} 見開きの番号
 */
function getResultSpreadIndex(spreadContext, resultIndex) {
    if (!spreadContext.isFacing) return resultIndex;
    return Math.floor((resultIndex + spreadContext.spreadOffset) / 2);
}

/**
 * 挿入後の並びのうち、現在のページの見開きから、挿入したページの次のページの見開きまでを、描く順のページの一覧にする。
 * 挿入するページが多いときは、前後 THUMBNAIL_MAX_INSERTED_PAGES / 2 ページずつだけ残す
 * @param {Object} spreadContext - collectSpreadContext() の戻り値
 * @param {number} insertCount - 挿入するページ数（0 なら今の並び）
 * @returns {Object[]} ページ（resultIndex / spreadIndex / kind: "current" | "inserted" | "existing" / prefix）
 */
function buildInsertThumbnail(spreadContext, insertCount) {
    var currentIndex = spreadContext.currentIndex;
    var resultCount = spreadContext.pageCount + insertCount;
    var firstSpread = getResultSpreadIndex(spreadContext, currentIndex);
    var lastSpread = getResultSpreadIndex(spreadContext, Math.min(currentIndex + insertCount + 1, resultCount - 1));
    var keepEachSide = THUMBNAIL_MAX_INSERTED_PAGES / 2;

    var thumbnailPages = [];
    for (var resultIndex = 0; resultIndex < resultCount; resultIndex++) {
        var spreadIndex = getResultSpreadIndex(spreadContext, resultIndex);
        if (spreadIndex < firstSpread) continue;
        if (spreadIndex > lastSpread) break;

        var insertedOrdinal = resultIndex - currentIndex; /* 1 から insertCount までが挿入するページ / 1..insertCount are inserted */
        if (insertedOrdinal >= 1 && insertedOrdinal <= insertCount) {
            if (insertCount > THUMBNAIL_MAX_INSERTED_PAGES
                    && insertedOrdinal > keepEachSide && insertedOrdinal <= insertCount - keepEachSide) continue;
            thumbnailPages.push({ resultIndex: resultIndex, spreadIndex: spreadIndex, kind: "inserted", prefix: spreadContext.insertedPrefix });
            continue;
        }
        var documentIndex = (resultIndex <= currentIndex) ? resultIndex : resultIndex - insertCount;
        if (spreadContext.prefixCache[documentIndex] === undefined) {
            spreadContext.prefixCache[documentIndex] = getAppliedMasterPrefix(spreadContext.documentPages[documentIndex]);
        }
        thumbnailPages.push({
            resultIndex: resultIndex,
            spreadIndex: spreadIndex,
            kind: (resultIndex === currentIndex) ? "current" : "existing",
            prefix: spreadContext.prefixCache[documentIndex]
        });
    }
    return thumbnailPages;
}

/**
 * 挿入後の見開きを縮小して描く。挿入するページは青、現在のページは太枠、省いたページの位置は点3つで示す。
 * 右綴じは見開きもページも右から左へ並べる
 * @param {ScriptUIGraphics} canvasGraphics - 描画先
 * @param {number} canvasWidth - 描画域の幅（px）
 * @param {number} canvasHeight - 描画域の高さ（px）
 * @param {Object} spreadContext - collectSpreadContext() の戻り値
 * @param {Object[]} thumbnailPages - buildInsertThumbnail() の戻り値
 * @returns {void}
 */
function drawInsertThumbnail(canvasGraphics, canvasWidth, canvasHeight, spreadContext, thumbnailPages) {
    fillThumbnailRect(canvasGraphics, 0, 0, canvasWidth, canvasHeight, THUMBNAIL_COLORS.background);
    if (!thumbnailPages || thumbnailPages.length === 0) return;

    /* ページの左端を、ページ幅を単位とした数と固定の間隔（px）に分けて積む / accumulate page slots and fixed gaps separately */
    var slotPositions = [], ellipsisPositions = [], gapPixels = 0;
    for (var i = 0; i < thumbnailPages.length; i++) {
        if (i > 0) {
            var previousPage = thumbnailPages[i - 1];
            if (thumbnailPages[i].resultIndex - previousPage.resultIndex > 1) {
                ellipsisPositions.push({ slots: i, gap: gapPixels });
                gapPixels += THUMBNAIL_ELLIPSIS_GAP;
            } else if (thumbnailPages[i].spreadIndex !== previousPage.spreadIndex) {
                gapPixels += THUMBNAIL_SPREAD_GAP;
            }
        }
        slotPositions.push({ slots: i, gap: gapPixels });
    }

    var labelFont = getThumbnailFont(canvasGraphics, 9);
    var labelHeight = measureThumbnailText(canvasGraphics, "0", labelFont)[1];
    var availableWidth = canvasWidth - THUMBNAIL_PADDING * 2;
    var availableHeight = canvasHeight - THUMBNAIL_PADDING * 2 - labelHeight - 2;
    var pageHeight = Math.min(THUMBNAIL_PAGE_HEIGHT, availableHeight);
    var pageWidth = pageHeight * spreadContext.pageAspect;
    if (pageWidth * thumbnailPages.length + gapPixels > availableWidth) {
        pageWidth = Math.max(2, (availableWidth - gapPixels) / thumbnailPages.length);
        pageHeight = pageWidth / spreadContext.pageAspect;
    }
    var totalWidth = pageWidth * thumbnailPages.length + gapPixels;
    var stripLeft = (canvasWidth - totalWidth) / 2;
    var stripTop = THUMBNAIL_PADDING + Math.max(0, (availableHeight - pageHeight) / 2);

    /**
     * 並びの位置から描く左端を求める（右綴じは左右を反転）
     * @param {{slots: number, gap: number}} position - ページ幅の数と間隔
     * @param {number} itemWidth - 描くものの幅（px）
     * @returns {number} 左端（px）
     */
    function toCanvasLeft(position, itemWidth) {
        var offset = position.slots * pageWidth + position.gap;
        return Math.round(spreadContext.isRightToLeft ? stripLeft + totalWidth - offset - itemWidth : stripLeft + offset);
    }

    var prefixPen = canvasGraphics.newPen(canvasGraphics.PenType.SOLID_COLOR, THUMBNAIL_COLORS.prefix, 1);
    var labelPen = canvasGraphics.newPen(canvasGraphics.PenType.SOLID_COLOR, THUMBNAIL_COLORS.label, 1);
    var roundedWidth = Math.max(2, Math.round(pageWidth));
    var roundedHeight = Math.max(2, Math.round(pageHeight));
    var roundedTop = Math.round(stripTop);
    for (var j = 0; j < thumbnailPages.length; j++) {
        var thumbnailPage = thumbnailPages[j];
        var pageLeft = toCanvasLeft(slotPositions[j], pageWidth);
        var isInserted = (thumbnailPage.kind === "inserted");
        fillThumbnailRect(canvasGraphics, pageLeft, roundedTop, roundedWidth, roundedHeight,
            isInserted ? THUMBNAIL_COLORS.inserted : THUMBNAIL_COLORS.page);
        strokeThumbnailRect(canvasGraphics, pageLeft, roundedTop, roundedWidth, roundedHeight,
            isInserted ? THUMBNAIL_COLORS.accent : THUMBNAIL_COLORS.pageFrame);

        if (thumbnailPage.prefix) {
            var prefixSize = measureThumbnailText(canvasGraphics, thumbnailPage.prefix, labelFont);
            if (prefixSize[0] + 4 <= roundedWidth && prefixSize[1] + 2 <= roundedHeight) {
                canvasGraphics.drawString(thumbnailPage.prefix, prefixPen,
                    Math.round(pageLeft + (roundedWidth - prefixSize[0]) / 2),
                    Math.round(roundedTop + (roundedHeight - prefixSize[1]) / 2), labelFont);
            }
        }

        if (thumbnailPage.kind === "current") {
            /* 太枠は2重の1px枠で描く / draw the thick frame as two 1px frames */
            strokeThumbnailRect(canvasGraphics, pageLeft, roundedTop, roundedWidth, roundedHeight, THUMBNAIL_COLORS.accent);
            strokeThumbnailRect(canvasGraphics, pageLeft + 1, roundedTop + 1, roundedWidth - 2, roundedHeight - 2, THUMBNAIL_COLORS.accent);
            var nameSize = measureThumbnailText(canvasGraphics, spreadContext.currentName, labelFont);
            var nameLeft = Math.max(0, Math.min(canvasWidth - nameSize[0], pageLeft + (roundedWidth - nameSize[0]) / 2));
            canvasGraphics.drawString(spreadContext.currentName, labelPen, Math.round(nameLeft), roundedTop + roundedHeight + 2, labelFont);
        }
    }

    /* 省いたページの位置に点を3つ / three dots where pages are left out */
    for (var k = 0; k < ellipsisPositions.length; k++) {
        var gapLeft = toCanvasLeft(ellipsisPositions[k], THUMBNAIL_ELLIPSIS_GAP);
        var dotTop = Math.round(stripTop + pageHeight / 2) - 1;
        for (var dotIndex = 0; dotIndex < 3; dotIndex++) {
            fillThumbnailRect(canvasGraphics, gapLeft + 2 + dotIndex * 4, dotTop, 2, 2, THUMBNAIL_COLORS.accent);
        }
    }
}

// =========================================
// ダイアログ / Dialog
// =========================================

/**
 * 全角数字を半角に変換する
 * @param {string} inputText 変換する文字列
 * @returns {string} 全角数字を半角に置き換えた文字列
 */
function toHalfWidthDigits(inputText) {
    return String(inputText).replace(/[０-９]/g, function (fullWidthDigit) {
        return String.fromCharCode(fullWidthDigit.charCodeAt(0) - 0xFEE0);
    });
}

/**
 * 入力文字列を挿入ページ数として解釈する
 * @param {string} inputText 入力欄の文字列
 * @returns {number} 1以上の整数。解釈できない場合は 0
 */
function parsePageCount(inputText) {
    var normalizedText = toHalfWidthDigits(inputText).replace(/^\s+|\s+$/g, "");
    if (!/^[0-9]+$/.test(normalizedText)) return 0;
    var parsedCount = parseInt(normalizedText, 10);
    return (parsedCount > 0) ? parsedCount : 0;
}

/**
 * 各行の左側ラベルの幅を最大値に揃える
 * @param {Array<StaticText>} rowLabelControls 幅を揃える StaticText の配列
 * @returns {void}
 */
function matchLabelWidthsToWidest(rowLabelControls) {
    var maxLabelWidth = 0;
    for (var i = 0; i < rowLabelControls.length; i++) {
        var labelWidth = rowLabelControls[i].preferredSize.width;
        if (labelWidth > maxLabelWidth) maxLabelWidth = labelWidth;
    }
    for (var j = 0; j < rowLabelControls.length; j++) {
        rowLabelControls[j].preferredSize.width = maxLabelWidth;
    }
}

/**
 * 挿入ページ数を尋ねるダイアログを表示する
 * @param {number} defaultPageCount 入力欄の初期値
 * @returns {{pageCount: number, insertAfterPage: Page, appliedMaster: MasterSpread}|null} 入力結果。キャンセル時は null
 */
function showPageCountDialog(defaultPageCount) {
    var pageInsertDialog = new Window("dialog", getLabel("dialog.title"));
    setupWindow(pageInsertDialog, 8);
    pageInsertDialog.alignChildren = ["left", "top"];

    var currentPage          = app.activeWindow.activePage;
    var currentAppliedMaster = currentPage.appliedMaster;

    var rowLabelControls = [];

    /* 現在のページ表示行 / Row showing the current page */
    var currentPageRow = pageInsertDialog.add("group");
    setupRow(currentPageRow, "left", 4);

    var currentPageLabel = currentPageRow.add("statictext", undefined, labelText("field.currentPage"));
    rowLabelControls.push(currentPageLabel);
    currentPageRow.add("statictext", undefined, currentPage.name);

    /* 親ページ名表示行 / Row showing the applied parent page */
    var appliedMasterRow = pageInsertDialog.add("group");
    setupRow(appliedMasterRow, "left", 4);

    var appliedMasterLabel = appliedMasterRow.add("statictext", undefined, labelText("field.appliedMaster"));
    rowLabelControls.push(appliedMasterLabel);
    appliedMasterRow.add("statictext", undefined, currentAppliedMaster ? currentAppliedMaster.name : "-");

    /* ページ数入力行 / Row for the page-count field */
    var pageCountRow = pageInsertDialog.add("group");
    setupRow(pageCountRow, "left", 4);

    var pageCountLabel = pageCountRow.add("statictext", undefined, labelText("field.pageCount"));
    rowLabelControls.push(pageCountLabel);
    /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
    var pageCountStepperGroup = pageCountRow.add("group");
    pageCountStepperGroup.orientation = "row";
    pageCountStepperGroup.alignChildren = ["left", "center"];
    pageCountStepperGroup.spacing = 0;
    pageCountStepperGroup.margins = 0;
    var pageCountInput;
    var pageCountStepper = addStepper(pageCountStepperGroup, function () { return pageCountInput; },
        { step: 1, min: 1, integer: true, onStep: function () { updateThumbnail(); } });
    pageCountInput = pageCountStepperGroup.add("edittext", undefined, defaultPageCount.toString());
    pageCountInput.characters = 5;
    bindSteppedArrowKeys(pageCountInput, pageCountStepper);
    pageCountInput.active = true;
    pageCountInput.helpTip = getLabel("tooltip.insertAfter");
    pageCountLabel.helpTip = getLabel("tooltip.insertAfter");

    matchLabelWidthsToWidest(rowLabelControls);

    /* 挿入後の見開きの縮小図（ダイアログの幅いっぱい）/ Thumbnail of the spreads after insertion (full dialog width) */
    var spreadContext = collectSpreadContext(app.activeDocument, currentPage, currentAppliedMaster);
    var thumbnailPages = null;
    var thumbnailCanvas = addThumbnailCanvas(pageInsertDialog, THUMBNAIL_SIZE, function (canvasGraphics, canvasWidth, canvasHeight) {
        drawInsertThumbnail(canvasGraphics, canvasWidth, canvasHeight, spreadContext, thumbnailPages);
    }, true);
    thumbnailCanvas.helpTip = getLabel("tooltip.thumbnail");

    /**
     * 入力中のページ数で縮小図を作り直して描き直す（数値でなければ挿入なしの並び）
     * @returns {void}
     */
    function updateThumbnail() {
        thumbnailPages = buildInsertThumbnail(spreadContext, parsePageCount(pageCountInput.text));
        redrawThumbnailCanvas(thumbnailCanvas);
    }
    pageCountInput.onChanging = updateThumbnail;
    pageCountInput.onChange = updateThumbnail;
    updateThumbnail();

    /* ボタン行（左が空なので中央に並ぶ）/ Button row (centered because the left group is empty) */
    var buttonRow = addButtonRow(pageInsertDialog);
    /* キャンセルは name: "cancel" の既定動作（クリック・ESC で閉じる）に任せる / Cancel relies on the built-in behavior of name: "cancel" */
    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    alignRightOnlyButtonRow(buttonRow);

    var pageInsertSettings = null;

    btnOK.onClick = function () {
        var enteredPageCount = parsePageCount(pageCountInput.text);
        if (enteredPageCount > 0) {
            pageInsertSettings = {
                pageCount: enteredPageCount,
                insertAfterPage: currentPage,
                appliedMaster: currentAppliedMaster
            };
            pageInsertDialog.close();
        } else {
            alert(getLabel("alert.invalidNumber"));
        }
    };

    addVersionFooter(pageInsertDialog);
    pageInsertDialog.show();
    return pageInsertSettings;
}

// =========================================
// ページ挿入 / Page insertion
// =========================================

/**
 * ダイアログの入力に従い、現在ページの直後にページを追加して親ページを適用する
 * @returns {void}
 */
function insertPagesAfterCurrentPage() {
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    /* 挿入位置はアクティブなレイアウトウィンドウのページから決まる / The insertion point comes from the active layout window */
    if (!isLayoutWindowActive()) {
        alert(getLabel("alert.noLayoutWindow"));
        return;
    }

    /* 親ページは document.pages のメンバーではないため挿入位置に使えない / A parent page is not a member of document.pages */
    if (isMasterPage(app.activeWindow.activePage)) {
        alert(getLabel("alert.masterPageActive"));
        return;
    }

    var pageInsertSettings = showPageCountDialog(DEFAULT_PAGE_COUNT);
    if (!pageInsertSettings) return; /* キャンセル時は何もしない / Do nothing when cancelled */

    var pageCount     = pageInsertSettings.pageCount;
    var appliedMaster = pageInsertSettings.appliedMaster;

    /* 一括 undo になるよう doScript でラップ / Wrap in doScript so the whole insertion is a single undo step */
    app.doScript(function () {
        var activeDocument  = app.activeDocument;
        var insertAfterPage = pageInsertSettings.insertAfterPage;

        /* ドキュメント基準で追加し、スプレッドへ正しく流し込む / Add at document level so pages reflow into proper spreads */
        for (var i = 0; i < pageCount; i++) {
            var insertedPage = activeDocument.pages.add(LocationOptions.AFTER, insertAfterPage);
            /* 現在のページが［なし］なら挿入したページも［なし］にそろえる / Match [None] as well, not just a named parent */
            insertedPage.appliedMaster = appliedMaster ? appliedMaster : NothingEnum.nothing;
            insertAfterPage = insertedPage;
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.insertPages"));
}

// =========================================
// 実行 / Run
// =========================================
insertPagesAfterCurrentPage();

})();
