#target indesign

/*

### 概要

アクティブページに版面・タイトルエリア・額縁・列行グリッド・区切り線などを、プレビューを見ながら一括作成します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdLayoutGridBuilder.md

### Overview

Builds the type area, title area, page frame, column and row grids and dividers on the active page in one pass, with a live preview.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdLayoutGridBuilder.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdLayoutGridBuilder";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdLayoutGridBuilder.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdLayoutGridBuilder.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

var PREVIEW_LAYER_NAME   = "__QuickLayoutPreview__";  /* プレビュー用レイヤー名 / preview layer name */
var TEMP_GRID_LAYER_NAME = "Temp Grid";               /* 仮グリッドを残すレイヤー名 / layer that keeps the temp grid */
var PAGE_FRAME_BLEED_MM  = 3;                         /* 額縁を裁ち落としへ広げる量（mm）/ page frame bleed in mm */
var DEFAULT_FONT_SIZE    = 9.5;                       /* 文字サイズが読めないときの値（pt）/ fallback font size in pt */

/* 長さの初期値（mm。定規の単位に換算して表示）/ Default lengths in mm, shown in ruler units */
var DEFAULT_LENGTHS_MM = {
    footerHeight: 30,  /* コラムエリアの高さ / footer column area height */
    offset:       10,  /* 実コンテンツ領域のオフセット / content region offset */
    columnGap:    10   /* 列の間隔 / column gap */
};

/* 区切り線の線種（ドキュメントの線種名）/ Stroke style names for dividers */
var DASHED_STROKE_STYLE_NAME = "破線 (3 & 2)";
var DOTTED_STROKE_STYLE_NAME = "点線 (1 & 1)";

/* サンプル文に使うフォントの PostScript 名（先に見つかったもの）/ PostScript names for the sample text, first match wins */
var SAMPLE_FONT_POSTSCRIPT_NAMES = ["HiraKakuProN-W3", "HiraKakuPro-W3", "HiraginoSans-W3"];

/* 作成するカラー / Colors created in the document */
var LAYOUT_COLORS = {
    titleFill:  { name: "K25", cmyk: [0, 0, 0, 25] },
    footerFill: { name: "K40", cmyk: [0, 0, 0, 40] },
    pageFrame:  { name: "K30", cmyk: [0, 0, 0, 30] },
    cellFill:   { name: "K10", cmyk: [0, 0, 0, 10] },
    tempGrid:   { name: "LayoutGrid", cmyk: [100, 0, 0, 0] }
};

// =========================================
// 単位換算 / Unit conversion
// =========================================

var MM_PER_POINT = 25.4 / 72;           /* 1pt = 0.352777…mm / millimeters per point */
var Q_PER_POINT  = MM_PER_POINT / 0.25;  /* 1Q（1H）= 0.25mm なので 1pt = 1.41111…Q / Q (and H) per point */

// =========================================
// レイアウト / Layout
// =========================================

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

var AUTO_BUTTON_SIZE = [70, 22];           /* ［自動調整］ボタンの寸法 / auto-adjust button size */

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

// リンクアイコン（再利用パーツ） / Link toggle (reusable)

// -----------------------------------------
// リンクアイコンの寸法 / Link toggle metrics
// -----------------------------------------
var LINK_ICON_SIZE          = [22, 22]; /* アイコンの大きさ / icon size */
var LINK_ICON_STROKE        = 1.5;      /* 線幅 / stroke width */
var LINK_CUT_DIRECTION      = [1, 0];   /* 連動中の左辺の切れ目の向き（水平）/ direction of the left-leg cut when linked (horizontal) */
var LINK_HOOK_CUT_DIRECTION = [0, 1];   /* 連動中の巻き込みの切れ目の向き（垂直）/ direction of the hook cut when linked (vertical) */
var LINK_STRAND_COUNT       = 4;        /* 切れ目の向きをそろえるための細い線の本数 / strands used to shape the cuts */
var LINK_SLASH_CLEARANCE    = 2.2;      /* 連動OFFの斜線とフックの間（22px 基準）/ gap between the slash and the hooks when unlinked */

// -----------------------------------------
// リンクアイコンの配色 / Link toggle colors
// -----------------------------------------
var LINK_UI_DARK = isDarkUI();
/* ダイアログの地に重ねる半透明の黒・白（UIの明るさの段階に追従する）。値はステップボタンの配色と同じ
   Translucent overlays that follow the dialog background; same values as the stepper buttons */
var LINK_PRESSED_COLOR  = LINK_UI_DARK ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 連動中の地 / background while linked */
var LINK_FRAME_COLOR    = LINK_UI_DARK ? [1, 1, 1, 0.07] : [0, 0, 0, 0.10]; /* 連動中の枠 / frame while linked */
var LINK_ICON_COLOR     = LINK_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* アイコンの線 / icon strokes */
var LINK_DIM_ICON_COLOR = LINK_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の線 / strokes when disabled */

// -----------------------------------------
// アイコンを作る・切り替える（外から呼ぶ関数） / Public API
// -----------------------------------------
/**
 * 連動の ON／OFF を切り替えるリンクアイコンを追加する（onDraw で自作描画）。
 * クリックで切り替わる。連動中は押し込んだボタンのように地と枠を描く。
 * @param {Group} parent - 追加先
 * @param {boolean} initialValue - 連動の初期値
 * @param {Function} onToggle - 切り替えたあとに呼ぶ関数
 * @returns {Group} アイコン（.value で連動中かを読む）
 */
function addLinkToggle(parent, initialValue, onToggle) {
    var linkToggle = parent.add("group");
    linkToggle.preferredSize = LINK_ICON_SIZE;
    linkToggle.minimumSize = LINK_ICON_SIZE;
    linkToggle.maximumSize = LINK_ICON_SIZE;
    linkToggle.value = initialValue;

    linkToggle.onDraw = function () {
        var iconGraphics = linkToggle.graphics;
        var iconWidth = LINK_ICON_SIZE[0];
        var iconHeight = LINK_ICON_SIZE[1];
        /* 自作描画は自動でディムにならないため、親もたどって判定する / Custom drawing is not dimmed automatically */
        var isDimmed = !isLinkToggleEnabledInTree(linkToggle);
        /* 連動中は押し込んだボタンのように地と枠を描く / While linked, draw it like a pressed button */
        if (linkToggle.value && !isDimmed) {
            iconGraphics.newPath();
            iconGraphics.rectPath(0, 0, iconWidth, iconHeight);
            iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, LINK_PRESSED_COLOR));
            iconGraphics.newPath();
            iconGraphics.rectPath(0.5, 0.5, iconWidth - 1, iconHeight - 1);
            iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, LINK_FRAME_COLOR, 1));
        }
        drawLinkIcon(iconGraphics, iconWidth, iconHeight, linkToggle.value, isDimmed ? LINK_DIM_ICON_COLOR : LINK_ICON_COLOR);
    };

    linkToggle.addEventListener("mousedown", function () {
        if (!isLinkToggleEnabledInTree(linkToggle)) return;
        linkToggle.value = !linkToggle.value;
        redrawLinkToggle(linkToggle);
        if (onToggle) onToggle();
    });
    return linkToggle;
}

/**
 * 連動の状態をコードから変えて描き直す（onToggle は呼ばない）
 * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
 * @param {boolean} isLinked - 連動にするなら true
 * @returns {void}
 */
function setLinkToggleValue(linkToggle, isLinked) {
    if (linkToggle.value === isLinked) return;
    linkToggle.value = isLinked;
    redrawLinkToggle(linkToggle);
}

/**
 * アイコンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
 * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
 * @param {boolean} isEnabled - 有効にするなら true
 * @returns {void}
 */
function setLinkToggleEnabled(linkToggle, isEnabled) {
    if (linkToggle.enabled === isEnabled) return;
    linkToggle.enabled = isEnabled;
    redrawLinkToggle(linkToggle);
}

/**
 * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
 * @param {Object} control - 判定するコントロール
 * @returns {boolean} すべて有効なら true
 */
function isLinkToggleEnabledInTree(control) {
    for (var node = control; node; node = node.parent) {
        if (!node.enabled) return false;
    }
    return true;
}

/**
 * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
 * @param {Group} linkToggle - 描き直すアイコン
 * @returns {void}
 */
function redrawLinkToggle(linkToggle) {
    linkToggle.hide();
    linkToggle.show();
}

// -----------------------------------------
// アイコンの形 / Icon geometry
// -----------------------------------------
/**
 * 連動アイコンを描く。Illustrator の［縦横比を固定］に合わせ、連動中は縦につながったチェーン、
 * 連動していないときは上下に分かれたチェーンに斜線を重ねる。座標は 22px 四方を基準に拡大縮小する。
 * @param {ScriptUIGraphics} iconGraphics - 描画先
 * @param {number} iconWidth - 描画範囲の幅
 * @param {number} iconHeight - 描画範囲の高さ
 * @param {boolean} isLinked - 連動中なら true
 * @param {number[]} iconColor - [r, g, b, a]
 * @returns {void}
 */
function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor) {
    var iconScale = Math.min(iconWidth, iconHeight) / 22;
    var offsetX = (iconWidth - 22 * iconScale) / 2;
    var offsetY = (iconHeight - 22 * iconScale) / 2;
    var strokes = isLinked ? buildLinkedChainStrokes() : buildUnlinkedChainStrokes();
    for (var i = 0; i < strokes.length; i++) {
        var strokePoints = strokes[i].points;
        /* newPath() を呼ばないとパスが前の描画に積み重なる / Without newPath() the paths accumulate */
        iconGraphics.newPath();
        for (var j = 0; j < strokePoints.length; j++) {
            var pointX = offsetX + strokePoints[j][0] * iconScale;
            var pointY = offsetY + strokePoints[j][1] * iconScale;
            if (j === 0) iconGraphics.moveTo(pointX, pointY);
            else iconGraphics.lineTo(pointX, pointY);
        }
        iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, strokes[i].width * iconScale));
    }
}

/**
 * 連動中のチェーン（縦に組み合った2つの輪）の線を返す。
 * 上の輪は左辺の途中から上端を回って右辺を下り、下端で内側へ巻き込む。下の輪はそれを180度回したもの。
 * 切れ目の向きをそろえるため、輪を細い線の束にし、両端を延ばしてから直線で切る（左辺は水平、巻き込みは垂直）
 * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
 */
function buildLinkedChainStrokes() {
    /* 左辺は上端の丸みだけ残して短く切り、下の輪の巻き込みとの間を空ける
       Keep only a stub on the left so it stays clear of the lower ring's hook */
    var upperRing = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360)
        .concat([[14.5, 11.2]])
        .concat(buildArcPoints(11, 11.2, 3.5, 2.3, 0, 115)));
    var ringStart = upperRing[0];
    var ringEnd = upperRing[upperRing.length - 1];
    var extendedRing = extendPolylineEnds(upperRing, LINK_ICON_STROKE);
    /* 延ばした先がどちら側かで、切り捨てる側を決める / The extended tips tell which side to cut away */
    var startOutsideSign = sideOfLine(extendedRing[0], ringStart, LINK_CUT_DIRECTION);
    var endOutsideSign = sideOfLine(extendedRing[extendedRing.length - 1], ringEnd, LINK_HOOK_CUT_DIRECTION);

    var upperStrands = buildStrandStrokes(extendedRing, function (strandPoints) {
        var trimmed = trimPolylineTail(strandPoints, ringEnd, LINK_HOOK_CUT_DIRECTION, endOutsideSign);
        trimmed = trimPolylineTail(trimmed.reverse(), ringStart, LINK_CUT_DIRECTION, startOutsideSign).reverse();
        return [trimmed];
    });
    var strokes = [];
    for (var i = 0; i < upperStrands.length; i++) {
        strokes.push(upperStrands[i]);
        strokes.push({ points: rotatePointsHalfTurn(upperStrands[i].points), width: upperStrands[i].width });
    }
    return strokes;
}

/**
 * 中心線を線幅の中で等分した細い線に分け、clipStrand で切った結果を線として返す。
 * @param {Array<number[]>} centerline - 中心線の点列
 * @param {Function} clipStrand - 細い線の点列を受け取り、残す点列の配列を返す関数
 * @returns {Array<{points: Array<number[]>, width: number}>} 細い線ごとの点列と線幅
 */
function buildStrandStrokes(centerline, clipStrand) {
    var strandWidth = LINK_ICON_STROKE / LINK_STRAND_COUNT;
    var strokes = [];
    for (var k = 0; k < LINK_STRAND_COUNT; k++) {
        /* 線幅の中を等分した位置に細い線を並べる / Lay the strands evenly across the stroke width */
        var strandOffset = -LINK_ICON_STROKE / 2 + strandWidth * (k + 0.5);
        var strandPieces = clipStrand(offsetPolyline(centerline, strandOffset));
        for (var j = 0; j < strandPieces.length; j++) {
            /* 隣の線と少し重ねて隙間を埋める / Overlap neighbours slightly so no seams show */
            if (strandPieces[j].length > 1) strokes.push({ points: strandPieces[j], width: strandWidth * 1.4 });
        }
    }
    return strokes;
}

/**
 * 点列の両端を、端の向きのまま length だけ延ばす。
 * @param {Array<number[]>} points - 点列
 * @param {number} length - 延ばす長さ
 * @returns {Array<number[]>} 延ばした点列
 */
function extendPolylineEnds(points, length) {
    /* from から to の向きへ、to から length 先の点 / point length beyond to, heading from from to to */
    function extendBeyond(from, to) {
        var dx = to[0] - from[0];
        var dy = to[1] - from[1];
        var segmentLength = Math.sqrt(dx * dx + dy * dy) || 1;
        return [to[0] + dx / segmentLength * length, to[1] + dy / segmentLength * length];
    }
    var lastIndex = points.length - 1;
    return [extendBeyond(points[1], points[0])].concat(points, [extendBeyond(points[lastIndex - 1], points[lastIndex])]);
}

/**
 * 点が直線のどちら側にあるかを符号で返す。
 * @param {number[]} point - 点
 * @param {number[]} linePoint - 直線上の1点
 * @param {number[]} direction - 直線の向き
 * @returns {number} 正・負で側を表す値
 */
function sideOfLine(point, linePoint, direction) {
    return direction[0] * (point[1] - linePoint[1]) - direction[1] * (point[0] - linePoint[0]);
}

/**
 * 点列の終わり側で、直線より outsideSign の側にはみ出した部分を切り、直線との交点で止める。
 * 輪の別の場所が同じ直線をまたいでも切らないよう、終わりから数点の範囲だけを見る。
 * @param {Array<number[]>} points - 点列
 * @param {number[]} cutPoint - 切る直線上の1点
 * @param {number[]} direction - 切る直線の向き
 * @param {number} outsideSign - 切り捨てる側の符号
 * @returns {Array<number[]>} 切った点列
 */
function trimPolylineTail(points, cutPoint, direction, outsideSign) {
    var lastIndex = points.length - 1;
    var searchLimit = Math.max(0, lastIndex - 12);
    var index = lastIndex;
    while (index > searchLimit && sideOfLine(points[index], cutPoint, direction) * outsideSign > 0) index--;
    if (index === lastIndex) return points.slice(0);
    var inside = points[index];
    var outside = points[index + 1];
    var insideSide = sideOfLine(inside, cutPoint, direction);
    var ratio = insideSide / (insideSide - sideOfLine(outside, cutPoint, direction));
    return points.slice(0, index + 1).concat([[inside[0] + (outside[0] - inside[0]) * ratio, inside[1] + (outside[1] - inside[1]) * ratio]]);
}

/**
 * 連動していないときのチェーン（上下に分かれた輪と斜線）の線を返す。
 * フックは斜線の近くで切る。線の端は進む向きに直角にしか切れないため、フックを細い線の束にして
 * 1本ずつ斜線と平行な境界で切り、切り口が斜線に沿って見えるようにする。
 * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
 */
function buildUnlinkedChainStrokes() {
    var slashStart = [3.5, 3.5];
    var slashEnd = [18.5, 18.5];
    var upperHook = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360).concat([[14.5, 11.5]]));
    var hooks = [upperHook, rotatePointsHalfTurn(upperHook)];

    /* 斜線の近くの帯を切り取る / Cut away the band around the slash */
    function clipAroundSlash(strandPoints) {
        return clipOutsideBand(strandPoints, slashStart, slashEnd, LINK_SLASH_CLEARANCE);
    }
    var strokes = buildStrandStrokes(hooks[0], clipAroundSlash).concat(buildStrandStrokes(hooks[1], clipAroundSlash));
    strokes.push({ points: [slashStart, slashEnd], width: LINK_ICON_STROKE });
    return strokes;
}

/**
 * 点の間隔が 0.5 以下になるよう、線分の間に点を足す。
 * @param {Array<number[]>} points - 点列
 * @returns {Array<number[]>} 細かくした点列
 */
function densifyPoints(points) {
    var densePoints = [points[0]];
    for (var i = 1; i < points.length; i++) {
        var from = points[i - 1];
        var to = points[i];
        var steps = Math.max(1, Math.ceil(Math.sqrt(Math.pow(to[0] - from[0], 2) + Math.pow(to[1] - from[1], 2)) / 0.5));
        for (var j = 1; j <= steps; j++) {
            densePoints.push([from[0] + (to[0] - from[0]) * j / steps, from[1] + (to[1] - from[1]) * j / steps]);
        }
    }
    return densePoints;
}

/**
 * 点列を、進む向きの左側へ offset だけずらした点列を返す（負の値なら右側）。
 * @param {Array<number[]>} points - 点列
 * @param {number} offset - ずらす距離
 * @returns {Array<number[]>} ずらした点列
 */
function offsetPolyline(points, offset) {
    var shifted = [];
    for (var i = 0; i < points.length; i++) {
        var before = points[Math.max(0, i - 1)];
        var after = points[Math.min(points.length - 1, i + 1)];
        var tangentX = after[0] - before[0];
        var tangentY = after[1] - before[1];
        var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY) || 1;
        shifted.push([points[i][0] - tangentY / tangentLength * offset, points[i][1] + tangentX / tangentLength * offset]);
    }
    return shifted;
}

/**
 * 直線（線分を延長したもの）から clearance 未満の帯に入る部分を切り取り、残りを点列に分けて返す。
 * 帯の境界で線分を補間して切るので、切り口は直線と平行にそろう。
 * @param {Array<number[]>} points - 点列
 * @param {number[]} lineStart - 直線上の1点
 * @param {number[]} lineEnd - 直線上のもう1点
 * @param {number} clearance - 空ける距離
 * @returns {Array<Array<number[]>>} 帯の外側に残った点列（2点未満のものは除く）
 */
function clipOutsideBand(points, lineStart, lineEnd, clearance) {
    var directionX = lineEnd[0] - lineStart[0];
    var directionY = lineEnd[1] - lineStart[1];
    var directionLength = Math.sqrt(directionX * directionX + directionY * directionY);

    /* 直線からの符号付き距離 / signed distance from the line */
    function signedDistance(point) {
        return (directionX * (point[1] - lineStart[1]) - directionY * (point[0] - lineStart[0])) / directionLength;
    }
    /* 2点の間で、距離が boundary になる点 / point between two points where the distance equals boundary */
    function interpolateAt(from, to, fromDistance, toDistance, boundary) {
        var ratio = (boundary - fromDistance) / (toDistance - fromDistance);
        return [from[0] + (to[0] - from[0]) * ratio, from[1] + (to[1] - from[1]) * ratio];
    }

    var pieces = [];
    var currentPiece = [];
    for (var i = 0; i < points.length; i++) {
        var distance = signedDistance(points[i]);
        var isOutside = Math.abs(distance) >= clearance;
        if (i > 0) {
            var previousDistance = signedDistance(points[i - 1]);
            var wasOutside = Math.abs(previousDistance) >= clearance;
            if (wasOutside && !isOutside) {
                /* 帯に入る: 境界で止める / entering the band: stop at the boundary */
                currentPiece.push(interpolateAt(points[i - 1], points[i], previousDistance, distance, previousDistance > 0 ? clearance : -clearance));
                if (currentPiece.length > 1) pieces.push(currentPiece);
                currentPiece = [];
            } else if (!wasOutside && isOutside) {
                /* 帯から出る: 境界から始める / leaving the band: start at the boundary */
                currentPiece = [interpolateAt(points[i - 1], points[i], previousDistance, distance, distance > 0 ? clearance : -clearance)];
            }
        }
        if (isOutside) currentPiece.push(points[i]);
    }
    if (currentPiece.length > 1) pieces.push(currentPiece);
    return pieces;
}

/**
 * 楕円弧の点列を返す（角度は右が0度、下が90度の画面座標）。
 * @param {number} centerX - 中心X
 * @param {number} centerY - 中心Y
 * @param {number} radiusX - 横の半径
 * @param {number} radiusY - 縦の半径
 * @param {number} startDegrees - 開始角度
 * @param {number} endDegrees - 終了角度
 * @returns {Array<number[]>} 点列
 */
function buildArcPoints(centerX, centerY, radiusX, radiusY, startDegrees, endDegrees) {
    var arcSteps = 12;
    var arcPoints = [];
    for (var i = 0; i <= arcSteps; i++) {
        var angle = (startDegrees + (endDegrees - startDegrees) * i / arcSteps) * Math.PI / 180;
        arcPoints.push([centerX + radiusX * Math.cos(angle), centerY + radiusY * Math.sin(angle)]);
    }
    return arcPoints;
}

/**
 * 点列を 22px 四方の中心で180度回す。
 * @param {Array<number[]>} points - 点列
 * @returns {Array<number[]>} 回した点列
 */
function rotatePointsHalfTurn(points) {
    var rotated = [];
    for (var i = 0; i < points.length; i++) {
        rotated.push([22 - points[i][0], 22 - points[i][1]]);
    }
    return rotated;
}

// リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle

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
var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
    btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
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

var LABELS = {
    dialog: {
        title: { ja: "版面とグリッドの作成", en: "Layout Grid Builder" }
    },
    tab: {
        page:          { ja: "ページ", en: "Page" },
        areas:         { ja: "額縁・エリア", en: "Frame & Areas" },
        contentRegion: { ja: "実コンテンツ領域", en: "Content Region" }
    },
    panel: {
        units:         { ja: "単位（定規／線／文字／行送り）", en: "Units (Ruler / Stroke / Text / Leading)" },
        baseText:      { ja: "基本テキスト", en: "Base Text" },
        margin:        { ja: "マージン", en: "Margins" },
        pageFrame:     { ja: "額縁", en: "Page Frame" },
        typeArea:      { ja: "版面の罫線", en: "Type Area Border" },
        titleArea:     { ja: "タイトルエリア", en: "Title Area" },
        footerArea:    { ja: "フッターのコラムエリア", en: "Footer Column Area" },
        offset:        { ja: "オフセット", en: "Offset" },
        rowCol:        { ja: "列・行", en: "Columns / Rows" },
        cells:         { ja: "セル", en: "Cells" },
        divider:       { ja: "区切り線", en: "Dividers" }
    },
    fieldLabel: {
        cornerRadius: { ja: "角丸", en: "Corner Radius" },
        extension:    { ja: "伸縮", en: "Extension" },
        capStyle:     { ja: "線端", en: "Cap" },
        position:     { ja: "位置", en: "Position" },
        relative:     { ja: "相対", en: "Relative" },
        columnCount:  { ja: "列数", en: "Columns" },
        rowCount:     { ja: "行数", en: "Rows" },
        gap:          { ja: "間隔", en: "Gap" },
        fontSize:     { ja: "サイズ", en: "Size" },
        leading:      { ja: "行送り", en: "Leading" },
        height:       { ja: "高さ", en: "Height" },
        footerGap:    { ja: "アキ", en: "Gap" },
        tempGrid:     { ja: "仮グリッド", en: "Temp Grid" }
    },
    side: {
        top:    { ja: "天", en: "Top" },
        bottom: { ja: "地", en: "Bottom" },
        left:   { ja: "左", en: "Left" },
        right:  { ja: "右", en: "Right" }
    },
    checkbox: {
        border:        { ja: "罫線", en: "Border" },
        fill:          { ja: "塗り", en: "Fill" },
        drawBorder:    { ja: "罫線を描画", en: "Draw border" },
        drawPageFrame: { ja: "額縁を描画", en: "Draw page frame" },
        applyMargins:  { ja: "ページのマージンにも反映", en: "Apply to page margins" },
        bleed:         { ja: "裁ち落とし", en: "Use bleed" },
        frameCorner:   { ja: "角丸", en: "Corner radius" },
        drawDividers:  { ja: "区切り線を描画", en: "Draw dividers" },
        showTempGrid:  { ja: "表示", en: "Show" },
        keepTempGrid:  { ja: "残す", en: "Keep" },
        threadFrames:  { ja: "フレームを連結", en: "Thread frames" }
    },
    radio: {
        capNone:        { ja: "なし", en: "None" },
        capRound:       { ja: "丸型", en: "Round" },
        capProjecting:  { ja: "突出", en: "Projecting" },
        positionTop:    { ja: "上", en: "Top" },
        positionBottom: { ja: "下", en: "Bottom" },
        positionLeft:   { ja: "左", en: "Left" },
        positionRight:  { ja: "右", en: "Right" },
        cellFill:       { ja: "塗り", en: "Fill" },
        cellTextFrame:  { ja: "テキストフレーム", en: "Text Frame" },
        sampleNone:     { ja: "なし", en: "None" },
        sampleProse:    { ja: "サンプル文", en: "Sample text" },
        sampleDummy:    { ja: "ダミー文字", en: "Placeholder" },
        lineSolid:      { ja: "実線", en: "Solid" },
        lineDashed:     { ja: "破線", en: "Dashed" },
        lineDotted:     { ja: "点線", en: "Dotted" }
    },
    unit: {
        characters: { ja: "文字", en: "chars" }
    },
    button: {
        ok:            { ja: "OK", en: "OK" },
        cancel:        { ja: "キャンセル", en: "Cancel" },
        autoAdjust:    { ja: "自動調整", en: "Auto" },
        autoAdjustAll: { ja: "すべて自動調整", en: "Auto All" }
    },
    tooltip: {
        units: {
            ja: "線幅と基本テキストを入力する単位を切り替えます。定規の単位はドキュメントの設定に従います。",
            en: "Switches the units for stroke weights and base text. The ruler unit follows the document setting."
        },
        borderWeight: { ja: "版面とタイトルエリアの罫線の線幅です。", en: "Stroke weight for the type area and title area borders." },
        cornerRadius: {
            ja: "版面の罫線（［伸縮］が 0 のとき）とタイトルエリアの塗りに使います。",
            en: "Used for the type area border (when Extension is 0) and the title area fill."
        },
        extension: {
            ja: "0 のときは長方形で描きます。正の値で四辺の罫線を角から外へ伸ばし、負の値で内へ縮めます。",
            en: "At 0 the border is drawn as a rectangle. Positive values extend each side past the corners; negative values shorten them."
        },
        capStyle:    { ja: "［伸縮］が 0 以外のときに使えます。", en: "Available when Extension is not 0." },
        titleStroke: { ja: "線幅は［版面の罫線］と同じです。", en: "Uses the same weight as Type Area Border." },
        titleLength: {
            ja: "［位置］が左・右のときは、タイトルエリアの幅として使います。",
            en: "When Position is Left or Right, this is used as the width of the title area."
        },
        titleExtension:  { ja: "タイトルエリアの罫線を両端から伸ばす量です。", en: "How far the title area border extends past both ends." },
        autoTitleLength: { ja: "仮グリッドの線に合うように［高さ］を丸めます。", en: "Rounds Height to the nearest temp grid line." },
        footerGap:       { ja: "実コンテンツ領域とコラムエリアのあいだのアキです。", en: "Space between the content region and the column area." },
        autoFooterHeight: {
            ja: "コラムエリアの上端が仮グリッドの線に合うように［高さ］を調整します。",
            en: "Adjusts Height so the top of the column area sits on a temp grid line."
        },
        showTempGrid: {
            ja: "基本テキストの文字サイズと行送りから求めた行の位置に、水色の線をプレビューで表示します。",
            en: "Previews cyan lines at the line positions given by the base text size and leading."
        },
        keepTempGrid: {
            ja: "［OK］のあと、印刷されない「Temp Grid」レイヤーに仮グリッドを残します。",
            en: "After OK, keeps the temp grid on a non-printing \"Temp Grid\" layer."
        },
        relative: {
            ja: "入力した値の増減分を、天地左右のマージンにまとめて加えます。",
            en: "Adds the change in this value to all four margins."
        },
        drawPageFrame: { ja: "ページの周囲を K30 の塗りで囲みます。", en: "Surrounds the page with a K30 fill." },
        applyMargins: {
            ja: "［OK］のとき、このページの「マージン・段組」のマージンも同じ値に変更します。",
            en: "On OK, also sets this page's Margins and Columns margins to these values."
        },
        frameCorner: {
            ja: "額縁の内側（くり抜いた部分）の角を丸めます。外側の角は丸めません。",
            en: "Rounds the corners of the frame opening. The outer corners stay square."
        },
        bleed: {
            ja: "額縁の外側を 3mm 外へ広げます。見開きの内側には広げません。",
            en: "Extends the outer edge of the frame by 3 mm. The spine side of a spread is not extended."
        },
        linkSides: { ja: "天地左右を同じ値にそろえます。", en: "Keeps all four sides at the same value." },
        autoOffset: {
            ja: "左右は文字サイズの倍数に、天地は仮グリッドの線に合わせて調整します。",
            en: "Rounds left and right to multiples of the font size, and top and bottom to temp grid lines."
        },
        characterCount: {
            ja: "1 列に入る文字数です。変更すると列の間隔を計算し直します。",
            en: "Characters per column. Changing it recalculates the column gap."
        },
        autoColumnGap: { ja: "［文字］の値に合わせて列の間隔を計算し直します。", en: "Recalculates the column gap from the character count." },
        linkGaps:      { ja: "列と行の間隔を同じ値にそろえます。", en: "Keeps the column and row gaps the same." },
        threadFrames:  { ja: "作成したテキストフレームを順に連結します。", en: "Threads the new text frames in order." },
        sampleText: {
            ja: "［OK］のあと、最初のフレームに流し込みます（プレビューでは流し込みません）。",
            en: "Placed into the first frame after OK (not shown in the preview)."
        },
        dividers: {
            ja: "列・行の間隔の中央に線を引きます。間隔が 0 のときは使えません。",
            en: "Draws lines centered in the column and row gaps. Unavailable when both gaps are 0."
        },
        dividerStyle: {
            ja: "ドキュメントの線種「破線 (3 & 2)」「点線 (1 & 1)」を使います。",
            en: "Uses the document stroke styles \"破線 (3 & 2)\" and \"点線 (1 & 1)\"."
        },
        autoAdjustAll: { ja: "各パネルの［自動調整］をまとめて実行します。", en: "Runs every Auto button in the dialog." },
        stepUp: {
            ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
            en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
        },
        stepDown: {
            ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
            en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
        },
        stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
        stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
    },
    alert: {
        noDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document before running this script." },
        strokeStyleMissing: {
            ja: "線種「{name}」がドキュメントにありません。実線で描画します。",
            en: "The stroke style \"{name}\" was not found in the document. Solid lines will be drawn instead."
        }
    }
};

/**
 * ラベルの末尾に単位を括弧付きで添える
 * @param {string} labelKey 例: "panel.margin"
 * @param {string} unitLabel 表示する単位名
 * @returns {string} 単位付きのラベル
 */
function withUnit(labelKey, unitLabel) {
    return getLabel(labelKey) + " (" + unitLabel + ")";
}

// =========================================
// サンプル文 / Sample text
// =========================================

var SAMPLE_PROSE_TEXT = "朝、目が覚めると、枕元の端末が静かに光っていた。\r「おはようございます。昨日の記憶を同期しますか？」\r\r　私はしばらくその表示を見つめた。\r　同期ボタンは、もう三日間押していない。\r\r　窓の外には、相変わらず同じ街が広がっている。\r　高層ビルの壁面には、朝のニュースが流れていた。\r\r「政府は本日、記憶バックアップ制度の利用率が国民の92%に達したと発表しました」\r\r　人々は、もうほとんど忘れない。\r　毎晩、脳内の記憶はクラウドに保存される。事故でも病気でも、バックアップから復元できる。\r\r　昨日までの自分を、正確に続きから生きられる。\r\r　便利な世界だ。\r\r　私は端末を伏せて、キッチンへ向かった。\r　コーヒーを淹れていると、壁のディスプレイが自動で点灯する。\r\r「未同期の記憶があります」\r\r　分かっている。\r\r　その記憶のせいだ。\r\r　昨日、私は一人の老人に会った。\r\r　河川敷のベンチで、古い紙の本を読んでいた。\r　今どき珍しい。\r\r「それ、オフラインの本ですか？」\r\r　私が声をかけると、老人は少し笑った。\r\r「そうだよ。記録に残らないものが好きでね」\r\r　意味が分からなかった。\r\r　記録に残らない？\r　そんなもの、価値があるのだろうか。\r\r「今の時代、全部残せるじゃないですか」\r\r　私が言うと、老人は本を閉じて言った。\r\r「だから残らないものが必要なんだ」\r\r　風が吹いた。\r　河川敷の草が揺れる。\r\r「人はね、本当は忘れる生き物なんだよ」\r\r　私は黙っていた。\r\r「忘れるから、また会いたくなる。忘れるから、思い出になる」\r\r　老人は空を見上げた。\r\r「全部残るなら、人生はただのログだ」\r\r　ログ。\r\r　その言葉が、妙に頭に残った。\r\r　家に帰ってから、私は同期を押せなかった。\r\r　もし同期すれば、この会話は永久に保存される。\r　政府のサーバーにも、医療記録にも、私の人生ログにも。\r\r　そしてきっと、忘れられなくなる。\r\r　私は端末をもう一度見る。\r\r「記憶同期を実行しますか？」\r\r　画面の下に、小さく表示されている。\r\r「同期しない記憶は、時間とともに消失する可能性があります」\r\r　それでいい。\r\r　私は河川敷の風を思い出す。\r　老人の声を思い出す。\r\r　でも、きっと少しずつ薄れていく。\r\r　声の高さも。\r　顔の皺も。\r　本の色も。\r\r　いつか曖昧になる。\r\r　それでいいのだと思う。\r\r　私は端末の通知を閉じた。\r\r　しばらくして、端末が静かに言う。\r\r「未同期記憶の自動削除まで、残り23時間」\r\r　窓の外では、ドローンが郵便物を運んでいた。\r　街は今日も、正確に記録されている。\r\r　私はコーヒーを飲みながら、ふと思う。\r\r　もしかしたら、あの老人の顔も。\r　もう、はっきり思い出せない。\r\r　でも、不思議と安心していた。\r\r　その記憶は、私の中だけにある。\r\r　サーバーにも、政府にも、誰のログにも残らない。\r\r　ただ、私の人生のどこかに、少しだけ影響して。\r　そして、静かに消えていく。\r\r　端末の光が消える。\r\r　私は窓を開けた。\r\r　春の風が、部屋に入ってきた。";

/**
 * □□□□○□□□□● を繰り返したダミー文字を作る（5〜10 回ごとに改行）
 * @returns {string} ダミー文字列
 */
function buildDummyText() {
    var unitPattern = "□□□□○□□□□●";
    var dummyText = "";
    var repeatsInLine = 0;
    var repeatsUntilBreak = Math.floor(Math.random() * 6) + 5;
    for (var i = 0; i < 200; i++) {
        dummyText += unitPattern;
        repeatsInLine++;
        if (repeatsInLine >= repeatsUntilBreak) {
            dummyText += "\r";
            repeatsInLine = 0;
            repeatsUntilBreak = Math.floor(Math.random() * 6) + 5;
        }
    }
    return dummyText;
}

// =========================================
// 共通ユーティリティ / Utilities
// =========================================

/**
 * 定規の単位の表示名を返す
 * @param {MeasurementUnits} measurementUnit 定規の単位
 * @returns {string} 単位名
 */
function getRulerUnitLabel(measurementUnit) {
    switch (measurementUnit) {
        case MeasurementUnits.MILLIMETERS:     return "mm";
        case MeasurementUnits.CENTIMETERS:     return "cm";
        case MeasurementUnits.INCHES:
        case MeasurementUnits.INCHES_DECIMAL:  return "in";
        case MeasurementUnits.PICAS:           return "p";
        case MeasurementUnits.PIXELS:          return "px";
        case MeasurementUnits.CICEROS:         return "c";
        case MeasurementUnits.AGATES:          return "ag";
        case MeasurementUnits.AMERICAN_POINTS: return "ap";
        case MeasurementUnits.Q:               return "Q";
        case MeasurementUnits.HA:              return "H";
        default:                               return "pt";
    }
}

/**
 * ページの pt 寸法と定規の単位での寸法を比べて、1 単位あたりのポイント数を実測する
 * （換算表を持たないので、どの単位でもドキュメントと食い違わない）
 * @param {Page} targetPage 対象ページ
 * @returns {{horizontal: number, vertical: number}} 横・縦それぞれの 1 単位あたりのポイント数
 */
function measurePointsPerUnit(targetPage) {
    var topLeft = targetPage.resolve(AnchorPoint.TOP_LEFT_ANCHOR, CoordinateSpaces.SPREAD_COORDINATES)[0];
    var bottomRight = targetPage.resolve(AnchorPoint.BOTTOM_RIGHT_ANCHOR, CoordinateSpaces.SPREAD_COORDINATES)[0];
    var bounds = targetPage.bounds;  /* [上, 左, 下, 右]。縦は縦の単位、横は横の単位 / vertical and horizontal units */
    return {
        horizontal: (bottomRight[0] - topLeft[0]) / (bounds[3] - bounds[1]),
        vertical: (bottomRight[1] - topLeft[1]) / (bounds[2] - bounds[0])
    };
}

/**
 * mm の長さを定規の単位に換算する
 * @param {number} millimeters 長さ（mm）
 * @param {number} pointsPerUnit 1 単位あたりのポイント数
 * @returns {number} 定規の単位での長さ
 */
function millimetersToUnits(millimeters, pointsPerUnit) {
    return millimeters / MM_PER_POINT / pointsPerUnit;
}

/**
 * ポイント値を単位付き文字列にする（素の数値はドキュメントの単位で解釈されるため）
 * @param {number} points ポイント値
 * @returns {string} 例: "0.3pt"
 */
function toPointString(points) {
    return String(points) + "pt";
}

/**
 * 小数点以下の桁数を指定して丸める
 * @param {number} value 対象の値
 * @param {number} digits 小数点以下の桁数
 * @returns {number} 丸めた値
 */
function roundTo(value, digits) {
    var scale = Math.pow(10, digits);
    return Math.round(value * scale) / scale;
}

/**
 * 文字列を数値にし、読めなければ代わりの値を返す（0 はそのまま 0）
 * @param {string} text 入力文字列
 * @param {number} fallback 数値にならないときの値
 * @returns {number} 数値
 */
function parseNumberOr(text, fallback) {
    var value = parseFloat(text);
    return isNaN(value) ? fallback : value;
}

/**
 * 配列に値が含まれるか調べる（ES3 に indexOf がないため）
 * @param {Array<string>} values 配列
 * @param {string} target 探す値
 * @returns {boolean} 含まれていれば true
 */
function arrayContains(values, target) {
    for (var i = 0; i < values.length; i++) {
        if (values[i] === target) return true;
    }
    return false;
}

/**
 * 範囲 [上, 左, 下, 右] を内側へ縮める
 * @param {Array<number>} bounds 元の範囲
 * @param {number} top 上から縮める量
 * @param {number} left 左から縮める量
 * @param {number} bottom 下から縮める量
 * @param {number} right 右から縮める量
 * @returns {Array<number>} 縮めた範囲
 */
function insetBounds(bounds, top, left, bottom, right) {
    return [bounds[0] + top, bounds[1] + left, bounds[2] - bottom, bounds[3] - right];
}

/**
 * 範囲からタイトルエリアを差し引く
 * @param {Array<number>} bounds 版面 [上, 左, 下, 右]
 * @param {boolean} titleOn タイトルエリアを描くか
 * @param {number} titleLength タイトルエリアの長さ
 * @param {string} titlePosition "top" / "bottom" / "left" / "right"
 * @returns {Array<number>} 差し引いた範囲
 */
function subtractTitleArea(bounds, titleOn, titleLength, titlePosition) {
    var result = bounds.slice(0);
    if (!titleOn || !(titleLength > 0)) return result;
    if (titlePosition === "top") result[0] += titleLength;
    else if (titlePosition === "bottom") result[2] -= titleLength;
    else if (titlePosition === "left") result[1] += titleLength;
    else if (titlePosition === "right") result[3] -= titleLength;
    return result;
}

/**
 * 範囲の下からフッターのコラムエリア（高さ＋アキ）を差し引く
 * @param {Array<number>} bounds 元の範囲 [上, 左, 下, 右]
 * @param {boolean} footerOn コラムエリアを描くか
 * @param {number} footerHeight コラムエリアの高さ
 * @param {number} footerGap 実コンテンツ領域とのアキ
 * @returns {Array<number>} 差し引いた範囲
 */
function subtractFooterArea(bounds, footerOn, footerHeight, footerGap) {
    var result = bounds.slice(0);
    if (footerOn && footerHeight > 0) result[2] -= (footerHeight + footerGap);
    return result;
}

/**
 * 版面上端からの距離を、最も近い仮グリッドの線の位置に丸める
 * 1 本目は上端から文字サイズぶん下、以降は行送りの間隔
 * @param {number} distance 版面上端からの距離
 * @param {{fontSize: number, leading: number}} fontMetrics 文字サイズと行送り（定規の単位）
 * @returns {number} 丸めた距離
 */
function snapToTempGrid(distance, fontMetrics) {
    var lineIndex = Math.round((distance - fontMetrics.fontSize) / fontMetrics.leading);
    if (lineIndex < 0) lineIndex = 0;
    return fontMetrics.fontSize + lineIndex * fontMetrics.leading;
}

// =========================================
// レイヤーとカラー / Layers and colors
// =========================================

/**
 * 印刷しない作業用レイヤーを取得する（なければ作成）
 * @param {Document} doc 対象ドキュメント
 * @param {string} layerName レイヤー名
 * @returns {Layer} 表示・ロック解除済みのレイヤー
 */
function getOrCreateWorkLayer(doc, layerName) {
    var layer = doc.layers.itemByName(layerName);
    if (!layer.isValid) layer = doc.layers.add({ name: layerName });
    layer.visible = true;
    layer.locked = false;
    layer.printable = false;
    return layer;
}

/**
 * 作業用レイヤーがアクティブなら、ほかのレイヤーに切り替える
 * @param {Document} doc 対象ドキュメント
 * @param {Array<string>} excludedNames アクティブにしないレイヤー名
 * @returns {void}
 */
function activateOtherLayer(doc, excludedNames) {
    if (!arrayContains(excludedNames, doc.activeLayer.name)) return;
    for (var i = 0; i < doc.layers.length; i++) {
        var layer = doc.layers[i];
        if (arrayContains(excludedNames, layer.name)) continue;
        try {
            doc.activeLayer = layer;
        } catch (e) {
            /* 切り替えられなくても描画は続ける / Keep drawing even if the switch fails */
        }
        return;
    }
}

/**
 * プレビューレイヤー上のオブジェクトをすべて削除する
 * @param {Document} doc 対象ドキュメント
 * @returns {void}
 */
function clearPreviewLayer(doc) {
    var layer = doc.layers.itemByName(PREVIEW_LAYER_NAME);
    if (!layer.isValid) return;
    layer.locked = false;
    layer.visible = true;
    for (var i = layer.pageItems.length - 1; i >= 0; i--) {
        layer.pageItems[i].remove();
    }
}

/**
 * プレビューレイヤーを削除する
 * @param {Document} doc 対象ドキュメント
 * @returns {void}
 */
function removePreviewLayer(doc) {
    var layer = doc.layers.itemByName(PREVIEW_LAYER_NAME);
    if (!layer.isValid) return;
    activateOtherLayer(doc, [PREVIEW_LAYER_NAME]);
    layer.remove();
}

/**
 * 名前付きのカラーを取得する（なければ作成）
 * @param {Document} doc 対象ドキュメント
 * @param {{name: string, cmyk: Array<number>}} colorDef カラーの定義
 * @returns {Color} カラー
 */
function getOrCreateColor(doc, colorDef) {
    var color = doc.colors.itemByName(colorDef.name);
    if (color.isValid) return color;
    return doc.colors.add({
        name: colorDef.name,
        model: ColorModel.PROCESS,
        space: ColorSpace.CMYK,
        colorValue: colorDef.cmyk
    });
}

// =========================================
// 描画 / Drawing
// =========================================

var ALL_CORNERS = ["topLeft", "topRight", "bottomLeft", "bottomRight"];

/* タイトルエリアの位置ごとに角丸にする角 / Corners rounded for each title position */
var TITLE_AREA_CORNERS = {
    top:    ["topLeft", "topRight"],
    bottom: ["bottomLeft", "bottomRight"],
    left:   ["topLeft", "bottomLeft"],
    right:  ["topRight", "bottomRight"]
};

/**
 * 設定値に従ってレイアウト要素を描画する
 * 縦の単位を横にそろえ、原点をスプレッドにしてから描き、最後に元へ戻す
 * @param {object} settings readSettings() が返す設定値
 * @returns {void}
 */
function drawLayout(settings) {
    var doc = app.activeDocument;
    var viewPrefs = doc.viewPreferences;
    var savedVerticalUnits = viewPrefs.verticalMeasurementUnits;
    var savedRulerOrigin = viewPrefs.rulerOrigin;
    var savedZeroPoint = doc.zeroPoint;
    var savedActiveLayer = null;

    viewPrefs.verticalMeasurementUnits = viewPrefs.horizontalMeasurementUnits;
    viewPrefs.rulerOrigin = RulerOrigin.SPREAD_ORIGIN;  /* 見開きの右ページにも対応 / Works on right-hand pages too */
    doc.zeroPoint = [0, 0];

    try {
        if (settings.targetLayer && settings.targetLayer.isValid) {
            savedActiveLayer = doc.activeLayer;
            doc.activeLayer = settings.targetLayer;
        }

        var page = app.activeWindow.activePage;
        var drawing = {
            doc: doc,
            page: page,
            settings: settings,
            regions: computeDrawRegions(page.bounds, settings),
            blackSwatch: doc.swatches.item("Black"),
            noneSwatch: doc.swatches.item("None")
        };

        if (!settings.gridOnly) {
            if (settings.typeAreaBorder) drawTypeAreaBorder(drawing);
            if (settings.titleLength > 0) drawTitleArea(drawing);
            if (drawing.regions.footerBounds) drawFooterArea(drawing);
            if (settings.pageFrame) drawPageFrame(drawing);
            if (settings.cellFill || settings.cellTextFrame) drawCells(drawing);
            if (settings.dividers && (settings.colCount > 1 || settings.rowCount > 1)) drawDividers(drawing);
        }
        if (settings.showTempGrid && (settings.targetLayer || settings.gridOnly)) drawTempGrid(drawing);
    } finally {
        if (savedActiveLayer && savedActiveLayer.isValid) {
            try {
                doc.activeLayer = savedActiveLayer;
            } catch (e) {
                /* 戻せなくても単位と原点の復元は続ける / Still restore units and origin */
            }
        }
        viewPrefs.rulerOrigin = savedRulerOrigin;
        doc.zeroPoint = savedZeroPoint;
        viewPrefs.verticalMeasurementUnits = savedVerticalUnits;
    }
}

/**
 * 描画に使う各領域を求める
 * @param {Array<number>} pageBounds ページ [上, 左, 下, 右]
 * @param {object} settings 設定値
 * @returns {object} page / typeArea / content / footerBounds / borderBottom / grid
 */
function computeDrawRegions(pageBounds, settings) {
    var typeArea = insetBounds(pageBounds, settings.marginTop, settings.marginLeft, settings.marginBottom, settings.marginRight);
    var footerOn = (settings.footerFill || settings.footerStroke) && settings.footerHeight > 0;
    var afterTitle = subtractTitleArea(typeArea, settings.titleFill || settings.titleStroke, settings.titleLength, settings.titlePosition);
    var content = subtractFooterArea(afterTitle, footerOn, settings.footerHeight, settings.footerGap);
    if (content[2] < content[0]) content[2] = content[0];

    /* フッターのコラムエリア（実コンテンツ領域の下端＋アキから）/ Footer column area below the content region */
    var footerBounds = null;
    if (footerOn) {
        var footerTop = content[2] + settings.footerGap;
        var footerBottom = footerTop + settings.footerHeight;
        if (footerTop < content[2]) footerTop = content[2];
        if (footerBottom > afterTitle[2]) footerBottom = afterTitle[2];
        if (footerBottom > footerTop && content[3] > content[1]) {
            footerBounds = [footerTop, content[1], footerBottom, content[3]];
        }
    }

    return {
        page: pageBounds,
        typeArea: typeArea,
        content: content,
        footerBounds: footerBounds,
        /* 版面の罫線はコラムエリア（高さ＋アキ）を除いた範囲 / The border stops above the footer */
        borderBottom: footerOn ? typeArea[2] - settings.footerHeight - settings.footerGap : typeArea[2],
        grid: insetBounds(content, settings.offsetTop, settings.offsetLeft, settings.offsetBottom, settings.offsetRight)
    };
}

/**
 * 線を 1 本引く（既定は黒）
 * @param {object} drawing 描画コンテキスト
 * @param {Array<Array<number>>} points 始点と終点 [[x, y], [x, y]]
 * @param {number} weightPt 線幅（pt）
 * @param {object} [lineOptions] endCap / strokeType / strokeColor
 * @returns {GraphicLine} 作成した線
 */
function addLine(drawing, points, weightPt, lineOptions) {
    var opts = lineOptions || {};
    var line = drawing.page.graphicLines.add();
    line.paths[0].entirePath = points;
    line.strokeWeight = toPointString(weightPt);
    line.strokeColor = opts.strokeColor || drawing.blackSwatch;
    if (opts.endCap) line.endCap = opts.endCap;
    if (opts.strokeType) line.strokeType = opts.strokeType;
    return line;
}

/**
 * 塗りだけの長方形を作り、最背面へ送る
 * @param {object} drawing 描画コンテキスト
 * @param {Array<number>} bounds [上, 左, 下, 右]
 * @param {{name: string, cmyk: Array<number>}} colorDef 塗りのカラー
 * @returns {Rectangle} 作成した長方形
 */
function addFilledRectangle(drawing, bounds, colorDef) {
    var rect = drawing.page.rectangles.add({
        geometricBounds: bounds,
        strokeWeight: 0,
        strokeColor: drawing.noneSwatch,
        fillColor: getOrCreateColor(drawing.doc, colorDef)
    });
    rect.sendToBack();
    return rect;
}

/**
 * 黒の罫線だけの長方形を作る
 * @param {object} drawing 描画コンテキスト
 * @param {Array<number>} bounds [上, 左, 下, 右]
 * @param {number} weightPt 線幅（pt）
 * @returns {Rectangle} 作成した長方形
 */
function addStrokedRectangle(drawing, bounds, weightPt) {
    return drawing.page.rectangles.add({
        geometricBounds: bounds,
        strokeWeight: toPointString(weightPt),
        strokeColor: drawing.blackSwatch,
        fillColor: drawing.noneSwatch
    });
}

/**
 * 長方形の指定した角を角丸にする
 * @param {Rectangle} rect 対象の長方形
 * @param {number} radius 角丸の半径（定規の単位）。0 以下なら何もしない
 * @param {Array<string>} cornerNames "topLeft" などの角の名前
 * @param {number} pointsPerUnit 1 単位あたりのポイント数
 * @returns {void}
 */
function roundCorners(rect, radius, cornerNames, pointsPerUnit) {
    if (!(radius > 0)) return;
    var radiusText = toPointString(radius * pointsPerUnit);
    for (var i = 0; i < cornerNames.length; i++) {
        rect[cornerNames[i] + "CornerOption"] = CornerOptions.ROUNDED_CORNER;
        rect[cornerNames[i] + "CornerRadius"] = radiusText;
    }
}

/**
 * 線端の識別子を EndCap に変換する
 * @param {string} capStyle "none" / "round" / "project"
 * @returns {EndCap} 線端
 */
function toEndCap(capStyle) {
    if (capStyle === "round") return EndCap.ROUND_END_CAP;
    if (capStyle === "project") return EndCap.PROJECTING_END_CAP;
    return EndCap.BUTT_END_CAP;
}

/**
 * 版面の罫線を描く。［伸縮］が 0 なら長方形、それ以外は四辺を別々の線で描く
 * @param {object} drawing 描画コンテキスト
 * @returns {void}
 */
function drawTypeAreaBorder(drawing) {
    var settings = drawing.settings;
    var typeArea = drawing.regions.typeArea;
    var top = typeArea[0];
    var left = typeArea[1];
    var right = typeArea[3];
    var bottom = drawing.regions.borderBottom;
    var ext = settings.borderExtension;

    if (ext === 0) {
        var rect = addStrokedRectangle(drawing, [top, left, bottom, right], settings.borderWeight);
        roundCorners(rect, settings.borderCornerRadius, ALL_CORNERS, settings.pointsPerUnit);
        return;
    }

    /* 正なら角から外へ伸ばし、負なら内へ縮める / Positive extends past the corners, negative leaves gaps */
    var lineOptions = { endCap: toEndCap(settings.capStyle) };
    addLine(drawing, [[left - ext, top], [right + ext, top]], settings.borderWeight, lineOptions);
    addLine(drawing, [[left - ext, bottom], [right + ext, bottom]], settings.borderWeight, lineOptions);
    addLine(drawing, [[left, top - ext], [left, bottom + ext]], settings.borderWeight, lineOptions);
    addLine(drawing, [[right, top - ext], [right, bottom + ext]], settings.borderWeight, lineOptions);
}

/**
 * タイトルエリアの塗りと罫線を描く
 * @param {object} drawing 描画コンテキスト
 * @returns {void}
 */
function drawTitleArea(drawing) {
    var settings = drawing.settings;
    var typeArea = drawing.regions.typeArea;
    var top = typeArea[0];
    var left = typeArea[1];
    var bottom = typeArea[2];
    var right = typeArea[3];
    var length = settings.titleLength;
    var ext = settings.titleExtension;
    var position = settings.titlePosition;

    /* 塗り：版面の端から長さぶん。外側の 2 角だけ角丸 / Fill with the two outer corners rounded */
    if (settings.titleFill) {
        var fillBounds = {
            top:    [top, left, top + length, right],
            bottom: [bottom - length, left, bottom, right],
            left:   [top, left, bottom, left + length],
            right:  [top, right - length, bottom, right]
        }[position];
        var rect = addFilledRectangle(drawing, fillBounds, LAYOUT_COLORS.titleFill);
        roundCorners(rect, settings.titleCornerRadius, TITLE_AREA_CORNERS[position], settings.pointsPerUnit);
    }

    /* 罫線：タイトルエリアの内側の辺に 1 本 / One line along the inner edge */
    if (settings.titleStroke) {
        var points;
        if (position === "top" || position === "bottom") {
            var lineY = (position === "top") ? top + length : bottom - length;
            points = [[left - ext, lineY], [right + ext, lineY]];
        } else {
            var lineX = (position === "left") ? left + length : right - length;
            points = [[lineX, top - ext], [lineX, bottom + ext]];
        }
        addLine(drawing, points, settings.borderWeight);
    }
}

/**
 * フッターのコラムエリアの塗りと罫線を描く
 * @param {object} drawing 描画コンテキスト
 * @returns {void}
 */
function drawFooterArea(drawing) {
    var settings = drawing.settings;
    var bounds = drawing.regions.footerBounds;
    if (settings.footerFill) {
        roundCorners(addFilledRectangle(drawing, bounds, LAYOUT_COLORS.footerFill), settings.footerCornerRadius, ALL_CORNERS, settings.pointsPerUnit);
    }
    if (settings.footerStroke) {
        roundCorners(addStrokedRectangle(drawing, bounds, settings.footerWeight), settings.footerCornerRadius, ALL_CORNERS, settings.pointsPerUnit);
    }
}

/**
 * ページの周囲を囲む額縁を、穴あきの多角形で描く
 * @param {object} drawing 描画コンテキスト
 * @returns {void}
 */
function drawPageFrame(drawing) {
    var settings = drawing.settings;
    var pageBounds = drawing.regions.page;
    var outer = pageBounds.slice(0);

    /* 裁ち落とし：天地は常に、左右は見開きの外側だけ / Bleed on the outside edges only */
    if (settings.pageFrameBleed) {
        var pageSide = drawing.page.side;
        var bleed = millimetersToUnits(PAGE_FRAME_BLEED_MM, settings.pointsPerUnit);
        outer[0] -= bleed;
        outer[2] += bleed;
        if (pageSide !== PageSideOptions.RIGHT_HAND) outer[1] -= bleed;
        if (pageSide !== PageSideOptions.LEFT_HAND) outer[3] += bleed;
    }
    var inner = insetBounds(pageBounds, settings.pageFrameTop, settings.pageFrameLeft, settings.pageFrameBottom, settings.pageFrameRight);

    /* 外側を順回り、内側を逆回りにして型抜きにする / Reverse the inner path to cut a hole */
    var framePolygon = drawing.page.polygons.add();
    framePolygon.paths[0].entirePath = [[outer[1], outer[0]], [outer[3], outer[0]], [outer[3], outer[2]], [outer[1], outer[2]]];
    var openingPath = framePolygon.paths.add();
    openingPath.entirePath = buildOpeningPath(inner, settings.pageFrameCornerRadiusPt / settings.pointsPerUnit);
    /* 足したパスは開いたままなので閉じる（閉じないと最後の角が直線でつながり面取りになる）/ Close it, or the last corner becomes a chamfer */
    openingPath.pathType = PathType.CLOSED_PATH;
    framePolygon.fillColor = getOrCreateColor(drawing.doc, LAYOUT_COLORS.pageFrame);
    framePolygon.strokeColor = drawing.noneSwatch;
    framePolygon.strokeWeight = 0;
    framePolygon.sendToBack();
}

/**
 * 額縁のくり抜き部分のパスを、外側と逆回り（左上から下へ）で作る。半径があれば角を曲線で丸める
 * @param {Array<number>} bounds くり抜く範囲 [上, 左, 下, 右]
 * @param {number} radius 角丸の半径（定規の単位）。幅・高さの半分までに抑える
 * @returns {Array} entirePath に渡す点の配列（角丸ありは [入り方向, 基準点, 出方向] の組）
 */
function buildOpeningPath(bounds, radius) {
    var top = bounds[0];
    var left = bounds[1];
    var bottom = bounds[2];
    var right = bounds[3];
    var r = Math.min(radius, (right - left) / 2, (bottom - top) / 2);
    if (!(r > 0)) return [[left, top], [left, bottom], [right, bottom], [right, top]];

    /* 4 分の 1 円をベジェで近似する係数 / Bezier handle length for a quarter circle */
    var handle = r * 0.5522847498;
    /**
     * 直線と曲線のつなぎ目の点を作る
     * @param {number} x 基準点の x
     * @param {number} y 基準点の y
     * @param {Array<number>} inHandle 入り方向の点
     * @param {Array<number>} outHandle 出方向の点
     * @returns {Array<Array<number>>} [入り方向, 基準点, 出方向]
     */
    function point(x, y, inHandle, outHandle) {
        return [inHandle || [x, y], [x, y], outHandle || [x, y]];
    }
    return [
        point(left, top + r, [left, top + r - handle], null),           /* 左辺の上端 / left edge, top */
        point(left, bottom - r, null, [left, bottom - r + handle]),     /* 左辺の下端 / left edge, bottom */
        point(left + r, bottom, [left + r - handle, bottom], null),     /* 下辺の左端 / bottom edge, left */
        point(right - r, bottom, null, [right - r + handle, bottom]),   /* 下辺の右端 / bottom edge, right */
        point(right, bottom - r, [right, bottom - r + handle], null),   /* 右辺の下端 / right edge, bottom */
        point(right, top + r, null, [right, top + r - handle]),         /* 右辺の上端 / right edge, top */
        point(right - r, top, [right - r + handle, top], null),         /* 上辺の右端 / top edge, right */
        point(left + r, top, null, [left + r - handle, top])            /* 上辺の左端 / top edge, left */
    ];
}

/**
 * グリッドを列×行のセルに分けた範囲を求める（列ごとに上から下の順）
 * @param {Array<number>} grid グリッドの範囲 [上, 左, 下, 右]
 * @param {object} settings 設定値
 * @returns {Array<Array<number>>} セルの範囲の配列
 */
function computeCellBounds(grid, settings) {
    var cellWidth = ((grid[3] - grid[1]) - settings.colGap * (settings.colCount - 1)) / settings.colCount;
    var cellHeight = ((grid[2] - grid[0]) - settings.rowGap * (settings.rowCount - 1)) / settings.rowCount;
    var cells = [];
    for (var i = 0; i < settings.colCount; i++) {
        for (var j = 0; j < settings.rowCount; j++) {
            var cellLeft = grid[1] + i * (cellWidth + settings.colGap);
            var cellTop = grid[0] + j * (cellHeight + settings.rowGap);
            cells.push([cellTop, cellLeft, cellTop + cellHeight, cellLeft + cellWidth]);
        }
    }
    return cells;
}

/**
 * セルを塗り、またはテキストフレームで埋める
 * @param {object} drawing 描画コンテキスト
 * @returns {void}
 */
function drawCells(drawing) {
    var settings = drawing.settings;
    var cells = computeCellBounds(drawing.regions.grid, settings);
    var i;

    if (settings.cellFill) {
        for (i = 0; i < cells.length; i++) {
            addFilledRectangle(drawing, cells[i], LAYOUT_COLORS.cellFill);
        }
        return;
    }

    var textFrames = [];
    for (i = 0; i < cells.length; i++) {
        textFrames.push(drawing.page.textFrames.add({
            geometricBounds: cells[i],
            strokeWeight: 0,
            strokeColor: drawing.noneSwatch,
            fillColor: drawing.noneSwatch
        }));
    }
    if (settings.threadFrames) {
        for (i = 0; i < textFrames.length - 1; i++) {
            textFrames[i].nextTextFrame = textFrames[i + 1];
        }
    }
    /* プレビューでは流し込まない / Skip in the preview */
    if ((settings.sampleProse || settings.sampleDummy) && !settings.targetLayer) {
        placeSampleText(textFrames[0], settings);
    }
}

/**
 * テキストフレームにサンプル文かダミー文字を流し込み、基本テキストの書式を当てる
 * @param {TextFrame} textFrame 流し込み先
 * @param {object} settings 設定値
 * @returns {void}
 */
function placeSampleText(textFrame, settings) {
    textFrame.contents = settings.sampleDummy ? buildDummyText() : SAMPLE_PROSE_TEXT;

    var story = textFrame.parentStory;
    if (settings.fontSizePt > 0) story.pointSize = toPointString(settings.fontSizePt);
    if (settings.leading === "auto" || settings.leading === "") {
        story.leading = Leading.AUTO;
    } else {
        var leadingPt = parseFloat(settings.leading);
        if (leadingPt > 0) story.leading = toPointString(leadingPt);
    }

    var sampleFont = findFontByPostScriptName(SAMPLE_FONT_POSTSCRIPT_NAMES);
    if (sampleFont) story.appliedFont = sampleFont;
}

/**
 * PostScript 名の候補から、インストールされているフォントを探す
 * （app.fonts の名前は「ファミリー名＋タブ＋スタイル名」で言語によって変わるため、PostScript 名で照合する）
 * @param {Array<string>} postScriptNames 候補の PostScript 名（先頭ほど優先）
 * @returns {Font|null} 見つかったフォント。なければ null
 */
function findFontByPostScriptName(postScriptNames) {
    var installedNames = app.fonts.everyItem().postscriptName;
    for (var i = 0; i < postScriptNames.length; i++) {
        for (var j = 0; j < installedNames.length; j++) {
            if (installedNames[j] === postScriptNames[i]) return app.fonts[j];
        }
    }
    return null;
}

/**
 * 区切り線の線種を取得する。ドキュメントになければ実線（null）にする
 * @param {Document} doc 対象ドキュメント
 * @param {string} lineType "solid" / "dashed" / "dotted"
 * @param {boolean} shouldWarn 見つからないときに警告するか
 * @returns {StrokeStyle|null} 線種。実線なら null
 */
function findDividerStrokeStyle(doc, lineType, shouldWarn) {
    var styleName = null;
    if (lineType === "dashed") styleName = DASHED_STROKE_STYLE_NAME;
    else if (lineType === "dotted") styleName = DOTTED_STROKE_STYLE_NAME;
    if (!styleName) return null;

    var strokeStyle = doc.strokeStyles.itemByName(styleName);
    if (strokeStyle.isValid) return strokeStyle;
    if (shouldWarn) alert(getLabel("alert.strokeStyleMissing").replace("{name}", styleName));
    return null;
}

/**
 * 列・行の間隔の中央に区切り線を引く
 * @param {object} drawing 描画コンテキスト
 * @returns {void}
 */
function drawDividers(drawing) {
    var settings = drawing.settings;
    var grid = drawing.regions.grid;
    var cellWidth = ((grid[3] - grid[1]) - settings.colGap * (settings.colCount - 1)) / settings.colCount;
    var cellHeight = ((grid[2] - grid[0]) - settings.rowGap * (settings.rowCount - 1)) / settings.rowCount;
    var lineOptions = {
        endCap: EndCap.BUTT_END_CAP,
        /* 警告は確定時だけ（プレビューのたびに出さない）/ Warn only on the final run */
        strokeType: findDividerStrokeStyle(drawing.doc, settings.dividerLineType, !settings.targetLayer)
    };

    /* 列の区切り（縦線）/ Column dividers */
    for (var i = 1; i < settings.colCount; i++) {
        var lineX = grid[1] + i * cellWidth + (i - 0.5) * settings.colGap;
        addLine(drawing, [[lineX, grid[0]], [lineX, grid[2]]], settings.dividerWeight, lineOptions);
    }
    /* 行の区切り（横線）/ Row dividers */
    for (var j = 1; j < settings.rowCount; j++) {
        var lineY = grid[0] + j * cellHeight + (j - 0.5) * settings.rowGap;
        addLine(drawing, [[grid[1], lineY], [grid[3], lineY]], settings.dividerWeight, lineOptions);
    }
}

/**
 * 仮グリッド（行の位置の横線）をページ全幅に描く
 * 版面上端から文字サイズぶん下が 1 本目、以降は行送りの間隔
 * @param {object} drawing 描画コンテキスト
 * @returns {void}
 */
function drawTempGrid(drawing) {
    var settings = drawing.settings;
    var pageBounds = drawing.regions.page;
    var fontSizePt = (settings.fontSizePt > 0) ? settings.fontSizePt : DEFAULT_FONT_SIZE;
    var leadingPt = parseFloat(settings.leading);
    if (isNaN(leadingPt) || leadingPt <= 0) leadingPt = fontSizePt * 1.5;

    var lineOptions = { strokeColor: getOrCreateColor(drawing.doc, LAYOUT_COLORS.tempGrid) };
    var lineStep = leadingPt / settings.pointsPerUnit;
    var gridLines = [];
    for (var lineY = drawing.regions.typeArea[0] + fontSizePt / settings.pointsPerUnit; lineY <= pageBounds[2]; lineY += lineStep) {
        gridLines.push(addLine(drawing, [[pageBounds[1], lineY], [pageBounds[3], lineY]], 0.1, lineOptions));
    }
    if (gridLines.length > 1) drawing.page.groups.add(gridLines);
}

// =========================================
// ダイアログの構築 / Dialog construction
// =========================================

/**
 * 設定ダイアログを組み立てる
 * @param {object} context ドキュメント・ページ・単位・初期値
 * @returns {object} ダイアログとコントロール、入力単位の状態をまとめたオブジェクト
 */
function buildDialog(context) {
    var ui = {
        context: context,
        fontUnitIsQ: false,     /* 基本テキストを Q/H で入力中か / Base text entered in Q/H */
        strokeUnitIsMm: false,  /* 線幅を mm で入力中か / Stroke weights entered in mm */
        lastRelativeValue: 0    /* ［相対］の前回値 / Previous Relative value */
    };

    var dlg = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(dlg);
    ui.dlg = dlg;

    var settingsTabs = dlg.add("tabbedpanel");
    settingsTabs.alignChildren = ["fill", "top"];

    var pageTab = addTab(settingsTabs, getLabel("tab.page"));
    buildPageTab(ui, pageTab);

    var areasTab = addTab(settingsTabs, getLabel("tab.areas"));
    buildPageFramePanel(ui, areasTab);
    buildTitleAreaPanel(ui, areasTab);
    buildFooterAreaPanel(ui, areasTab);

    buildContentRegionTab(ui, addTab(settingsTabs, getLabel("tab.contentRegion")));
    settingsTabs.selection = pageTab;

    buildButtonRow(ui, dlg);
    return ui;
}

/**
 * 共通設定を当てたタブを追加する
 * @param {TabbedPanel} parent タブパネル
 * @param {string} title タブの見出し
 * @returns {Tab} タブ
 */
function addTab(parent, title) {
    var tab = parent.add("tab", undefined, title);
    setupTab(tab);
    return tab;
}

/**
 * 共通設定を当てたパネルを追加する
 * @param {Group|Panel} parent 親
 * @param {string} title パネルの見出し
 * @returns {Panel} パネル
 */
function addPanel(parent, title) {
    var panel = parent.add("panel", undefined, title);
    setupPanel(panel);
    return panel;
}

/**
 * 共通設定を当てた行グループを追加する
 * @param {Group|Panel} parent 親
 * @param {string} [alignment] 横方向の配置
 * @returns {Group} 行グループ
 */
function addRow(parent, alignment) {
    var row = parent.add("group");
    setupRow(row, alignment);
    return row;
}

/**
 * コロン付きの項目名を右揃えで追加する
 * @param {Group} parent 親グループ
 * @param {string} labelKey LABELS のキー
 * @returns {StaticText} 項目名
 */
function addFieldLabel(parent, labelKey) {
    var label = parent.add("statictext", undefined, labelText(labelKey));
    label.justify = "right";
    return label;
}

/**
 * ∧∨と↑↓キーで増減できる数値入力欄を追加する（∧∨は入力欄の左に隙間なく置く）。
 * 増減したあとは onChange を呼び、直接入力の確定時は最小値を下回らないようにする
 * @param {Group} parent 親グループ
 * @param {number|string} defaultValue 初期値
 * @param {number} characters 表示幅（文字数）
 * @param {number} [minValue] 最小値。省略時は下限なし（負の値も可）
 * @param {boolean} [isInteger] 整数のみなら true（列数・行数・文字数）
 * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できる）
 */
function addNumberInput(parent, defaultValue, characters, minValue, isInteger) {
    var stepperInputGroup = parent.add("group");
    stepperInputGroup.orientation = "row";
    stepperInputGroup.alignChildren = ["left", "center"];
    stepperInputGroup.spacing = 0;
    stepperInputGroup.margins = 0;

    var input;
    var stepperGroup = addStepper(stepperInputGroup, function () { return input; }, {
        step: 1,
        min: minValue,
        integer: isInteger === true,
        onStep: function (numberInput) { numberInput.notify("onChange"); }
    });
    input = stepperInputGroup.add("edittext", undefined, String(defaultValue));
    input.characters = characters;
    input.stepperGroup = stepperGroup;
    bindSteppedArrowKeys(input, stepperGroup);

    if (minValue !== undefined) {
        input.addEventListener("change", function () {
            var value = parseFloat(input.text);
            if (isNaN(value) || value < minValue) input.text = String(minValue);
        });
    }
    return input;
}

/**
 * 数値入力欄と∧∨の有効／無効をまとめて切り替える
 * @param {EditText} input addNumberInput() で作った入力欄
 * @param {boolean} isEnabled 有効にするなら true
 * @returns {void}
 */
function setNumberInputEnabled(input, isEnabled) {
    input.enabled = isEnabled;
    input.stepperGroup.enabled = isEnabled;
}

/**
 * 有効／無効が変わった∧∨だけを描き直す（入力のたびに全部を描き直さない）
 * @param {Object} container ダイアログ・パネルなど
 * @returns {void}
 */
function redrawChangedSteppers(container) {
    if (!container.children) return;
    for (var i = 0; i < container.children.length; i++) {
        var child = container.children[i];
        if (!child.isStepperButton) {
            redrawChangedSteppers(child);
            continue;
        }
        var isEnabled = isStepperEnabledInTree(child);
        if (child.drawnEnabled === isEnabled) continue;
        child.drawnEnabled = isEnabled;
        redrawStepperGroup(child);
    }
}

/**
 * 項目名と入力欄を並べた小さなグループを追加する（行の設定は当てない）
 * @param {Group} parent 親グループ
 * @param {string} labelKey LABELS のキー
 * @param {number|string} defaultValue 初期値
 * @param {number} characters 表示幅（文字数）
 * @returns {EditText} 入力欄
 */
function addLabeledInput(parent, labelKey, defaultValue, characters) {
    var labeledGroup = parent.add("group");
    addFieldLabel(labeledGroup, labelKey);
    return addNumberInput(labeledGroup, defaultValue, characters);
}

/**
 * 項目名・入力欄・単位を 1 行に並べる
 * @param {Group|Panel} parent 親
 * @param {string} labelKey LABELS のキー
 * @param {number|string} defaultValue 初期値
 * @param {number} characters 表示幅（文字数）
 * @param {string|null} unitText 単位の表記。null なら付けない
 * @param {number} [minValue] 最小値
 * @param {boolean} [isInteger] 整数のみなら true
 * @returns {{row: Group, input: EditText, unitLabel: StaticText}} 作成したコントロール
 */
function addNumberRow(parent, labelKey, defaultValue, characters, unitText, minValue, isInteger) {
    var row = addRow(parent);
    addFieldLabel(row, labelKey);
    var input = addNumberInput(row, defaultValue, characters, minValue, isInteger);
    var unitLabel = unitText ? row.add("statictext", undefined, unitText) : null;
    return { row: row, input: input, unitLabel: unitLabel };
}

/**
 * ［自動調整］ボタンを追加する
 * @param {Group|Panel} parent 親
 * @param {string} tooltipKey ツールチップのキー
 * @returns {Button} ボタン
 */
function addAutoButton(parent, tooltipKey) {
    var button = parent.add("button", undefined, getLabel("button.autoAdjust"));
    button.preferredSize = AUTO_BUTTON_SIZE;
    button.helpTip = getLabel(tooltipKey);
    return button;
}

/**
 * ラベルキーの一覧からラジオボタンを並べ、先頭を選択しておく
 * @param {Group} parent 親グループ（同じ親の中で排他になる）
 * @param {Array<string>} labelKeys LABELS のキー
 * @param {string} [tooltipKey] すべてに付けるツールチップのキー
 * @returns {Array<RadioButton>} ラジオボタン
 */
function addRadioButtons(parent, labelKeys, tooltipKey) {
    var radios = [];
    for (var i = 0; i < labelKeys.length; i++) {
        var radio = parent.add("radiobutton", undefined, getLabel(labelKeys[i]));
        if (tooltipKey) radio.helpTip = getLabel(tooltipKey);
        radios.push(radio);
    }
    radios[0].value = true;
    return radios;
}

/**
 * チェックボックスを追加する
 * @param {Group|Panel} parent 親
 * @param {string} labelKey LABELS のキー
 * @param {boolean} checked 初期状態
 * @param {string} [tooltipKey] ツールチップのキー
 * @returns {Checkbox} チェックボックス
 */
function addCheckbox(parent, labelKey, checked, tooltipKey) {
    var checkbox = parent.add("checkbox", undefined, getLabel(labelKey));
    checkbox.value = checked;
    if (tooltipKey) checkbox.helpTip = getLabel(tooltipKey);
    return checkbox;
}

/**
 * 天地左右の入力欄を「左｜天・連動アイコン・地｜右」の形に並べる
 * @param {Panel} parent 親パネル
 * @param {number} defaultValue 初期値
 * @param {number} characters 表示幅（文字数）
 * @returns {{row: Group, top: EditText, bottom: EditText, left: EditText, right: EditText, linkToggle: Group}} 作成したコントロール
 */
function addLinkedSideInputs(parent, defaultValue, characters) {
    var sidesRow = addRow(parent, "center");
    var left = addLabeledInput(sidesRow, "side.left", defaultValue, characters);

    var topBottomGroup = sidesRow.add("group");
    topBottomGroup.orientation = "column";
    topBottomGroup.alignChildren = "center";
    var top = addLabeledInput(topBottomGroup, "side.top", defaultValue, characters);
    var linkToggle = addLinkToggle(topBottomGroup, true);
    linkToggle.helpTip = getLabel("tooltip.linkSides");
    var bottom = addLabeledInput(topBottomGroup, "side.bottom", defaultValue, characters);

    var right = addLabeledInput(sidesRow, "side.right", defaultValue, characters);
    return { row: sidesRow, top: top, bottom: bottom, left: left, right: right, linkToggle: linkToggle };
}

/**
 * ［ページ］タブ（単位・基本テキスト・マージン・版面の罫線）を組み立てる
 * @param {object} ui UI オブジェクト
 * @param {Tab} pagePanel ［ページ］タブ
 * @returns {void}
 */
function buildPageTab(ui, pagePanel) {
    var unitLabel = ui.context.unitLabel;
    var defaults = ui.context.defaults;

    /* 単位：先頭は定規の単位（ドキュメントの設定）/ Units: the first part is the document ruler unit */
    var unitsPanel = addPanel(pagePanel, getLabel("panel.units"));
    var unitRadios = [
        unitsPanel.add("radiobutton", undefined, unitLabel + "/pt/pt/pt"),
        unitsPanel.add("radiobutton", undefined, unitLabel + "/mm/pt/pt"),
        unitsPanel.add("radiobutton", undefined, unitLabel + "/mm/Q/H")
    ];
    for (var i = 0; i < unitRadios.length; i++) unitRadios[i].helpTip = getLabel("tooltip.units");
    unitRadios[0].value = true;
    ui.rbUnitAllPt = unitRadios[0];
    ui.rbUnitStrokeMm = unitRadios[1];
    ui.rbUnitTextQ = unitRadios[2];

    /* 基本テキスト / Base text */
    var baseTextPanel = addPanel(pagePanel, getLabel("panel.baseText"));
    var fontSizeRow = addNumberRow(baseTextPanel, "fieldLabel.fontSize", DEFAULT_FONT_SIZE, 5, "pt");
    ui.baseFontSizeInput = fontSizeRow.input;
    ui.fontSizeUnitLabel = fontSizeRow.unitLabel;
    var leadingRow = addNumberRow(baseTextPanel, "fieldLabel.leading", 16, 5, "pt");
    ui.leadingInput = leadingRow.input;
    ui.leadingUnitLabel = leadingRow.unitLabel;
    var tempGridRow = addRow(baseTextPanel);
    addFieldLabel(tempGridRow, "fieldLabel.tempGrid");
    ui.showTempGridCheck = addCheckbox(tempGridRow, "checkbox.showTempGrid", true, "tooltip.showTempGrid");
    ui.keepTempGridCheck = addCheckbox(tempGridRow, "checkbox.keepTempGrid", true, "tooltip.keepTempGrid");

    /* マージン：額縁と同じ並び＋相対＋ページへの反映 / Margins: same layout as the page frame, plus Relative */
    var marginPanel = addPanel(pagePanel, withUnit("panel.margin", unitLabel));
    ui.marginSides = addLinkedSideInputs(marginPanel, 0, 4);
    ui.marginTopInput = ui.marginSides.top;
    ui.marginBottomInput = ui.marginSides.bottom;
    ui.marginLeftInput = ui.marginSides.left;
    ui.marginRightInput = ui.marginSides.right;
    ui.marginTopInput.text = String(defaults.marginTop);
    ui.marginBottomInput.text = String(defaults.marginBottom);
    ui.marginLeftInput.text = String(defaults.marginLeft);
    ui.marginRightInput.text = String(defaults.marginRight);
    /* 連動は四辺がそろっているときだけ初期オン / Link starts on only when all four sides match */
    setLinkToggleValue(ui.marginSides.linkToggle, defaults.marginTop === defaults.marginBottom
        && defaults.marginTop === defaults.marginLeft && defaults.marginTop === defaults.marginRight);
    var relativeRow = addNumberRow(marginPanel, "fieldLabel.relative", 0, 4, unitLabel);
    relativeRow.row.alignment = ["center", "center"];
    relativeRow.input.helpTip = getLabel("tooltip.relative");
    ui.relativeInput = relativeRow.input;
    ui.applyMarginsCheck = addCheckbox(marginPanel, "checkbox.applyMargins", false, "tooltip.applyMargins");

    buildTypeAreaPanel(ui, pagePanel);
}

/**
 * ［額縁］パネルを組み立てる
 * @param {object} ui UI オブジェクト
 * @param {Tab} parent 親タブ
 * @returns {void}
 */
function buildPageFramePanel(ui, parent) {
    var pageFramePanel = addPanel(parent, withUnit("panel.pageFrame", ui.context.unitLabel));
    var pageFrameCheckRow = addRow(pageFramePanel);
    ui.pageFrameCheck = addCheckbox(pageFrameCheckRow, "checkbox.drawPageFrame", false, "tooltip.drawPageFrame");
    ui.pageFrameBleedCheck = addCheckbox(pageFrameCheckRow, "checkbox.bleed", false, "tooltip.bleed");
    ui.pageFrameSides = addLinkedSideInputs(pageFramePanel, 0, 4);
    ui.pageFrameCornerRow = addRow(pageFramePanel);
    ui.pageFrameCornerCheck = addCheckbox(ui.pageFrameCornerRow, "checkbox.frameCorner", false, "tooltip.frameCorner");
    ui.pageFrameCornerInput = addNumberInput(ui.pageFrameCornerRow, 10, 4, 0);
    ui.pageFrameCornerUnitLabel = ui.pageFrameCornerRow.add("statictext", undefined, "pt");
}

/**
 * ［版面の罫線］パネルを組み立てる
 * @param {object} ui UI オブジェクト
 * @param {Tab} parent 親タブ
 * @returns {void}
 */
function buildTypeAreaPanel(ui, parent) {
    var typeAreaPanel = addPanel(parent, getLabel("panel.typeArea"));
    var unitLabel = ui.context.unitLabel;

    var borderRow = addRow(typeAreaPanel);
    ui.borderCheck = addCheckbox(borderRow, "checkbox.drawBorder", false);
    ui.borderWeightInput = addNumberInput(borderRow, 0.3, 5);
    ui.borderWeightInput.helpTip = getLabel("tooltip.borderWeight");
    ui.borderWeightUnitLabel = borderRow.add("statictext", undefined, "pt");

    var cornerRow = addNumberRow(typeAreaPanel, "fieldLabel.cornerRadius", 0, 4, unitLabel);
    cornerRow.input.helpTip = getLabel("tooltip.cornerRadius");
    ui.cornerRadiusRow = cornerRow.row;
    ui.cornerRadiusInput = cornerRow.input;

    var extensionRow = addNumberRow(typeAreaPanel, "fieldLabel.extension", 0, 4, unitLabel);
    extensionRow.input.helpTip = getLabel("tooltip.extension");
    ui.extensionRow = extensionRow.row;
    ui.extensionInput = extensionRow.input;

    ui.capStyleRow = addRow(typeAreaPanel);
    addFieldLabel(ui.capStyleRow, "fieldLabel.capStyle");
    var capRadios = addRadioButtons(ui.capStyleRow, ["radio.capNone", "radio.capRound", "radio.capProjecting"], "tooltip.capStyle");
    ui.rbCapNone = capRadios[0];
    ui.rbCapRound = capRadios[1];
    ui.rbCapProjecting = capRadios[2];
}

/**
 * ［タイトルエリア］パネルを組み立てる
 * @param {object} ui UI オブジェクト
 * @param {Tab} parent 親タブ
 * @returns {void}
 */
function buildTitleAreaPanel(ui, parent) {
    var titleAreaPanel = addPanel(parent, getLabel("panel.titleArea"));
    var unitLabel = ui.context.unitLabel;

    var titleCheckRow = addRow(titleAreaPanel);
    ui.titleFillCheck = addCheckbox(titleCheckRow, "checkbox.fill", false);
    ui.titleStrokeCheck = addCheckbox(titleCheckRow, "checkbox.border", false, "tooltip.titleStroke");

    var lengthRow = addNumberRow(titleAreaPanel, "fieldLabel.height", ui.context.defaults.titleLength, 4, unitLabel);
    lengthRow.input.helpTip = getLabel("tooltip.titleLength");
    ui.titleLengthRow = lengthRow.row;
    ui.titleLengthInput = lengthRow.input;
    ui.btnAutoTitleLength = addAutoButton(lengthRow.row, "tooltip.autoTitleLength");

    ui.titlePositionRow = addRow(titleAreaPanel);
    addFieldLabel(ui.titlePositionRow, "fieldLabel.position");
    var positionRadios = addRadioButtons(ui.titlePositionRow, ["radio.positionTop", "radio.positionBottom", "radio.positionLeft", "radio.positionRight"]);
    ui.rbTitleTop = positionRadios[0];
    ui.rbTitleBottom = positionRadios[1];
    ui.rbTitleLeft = positionRadios[2];
    ui.rbTitleRight = positionRadios[3];

    var extensionRow = addNumberRow(titleAreaPanel, "fieldLabel.extension", 0, 4, unitLabel);
    extensionRow.input.helpTip = getLabel("tooltip.titleExtension");
    ui.titleExtensionRow = extensionRow.row;
    ui.titleExtensionInput = extensionRow.input;
}

/**
 * ［フッターのコラムエリア］パネルを組み立てる
 * @param {object} ui UI オブジェクト
 * @param {Tab} parent 親タブ
 * @returns {void}
 */
function buildFooterAreaPanel(ui, parent) {
    var footerPanel = addPanel(parent, getLabel("panel.footerArea"));
    var unitLabel = ui.context.unitLabel;

    var footerCheckRow = addRow(footerPanel);
    ui.footerFillCheck = addCheckbox(footerCheckRow, "checkbox.fill", false);
    ui.footerStrokeCheck = addCheckbox(footerCheckRow, "checkbox.border", false);
    ui.footerWeightInput = addNumberInput(footerCheckRow, 0.3, 5);
    ui.footerWeightUnitLabel = footerCheckRow.add("statictext", undefined, "pt");

    var heightRow = addNumberRow(footerPanel, "fieldLabel.height", ui.context.defaults.footerHeight, 4, unitLabel);
    ui.footerHeightRow = heightRow.row;
    ui.footerHeightInput = heightRow.input;
    ui.btnAutoFooterHeight = addAutoButton(heightRow.row, "tooltip.autoFooterHeight");

    var gapRow = addNumberRow(footerPanel, "fieldLabel.footerGap", 0, 4, unitLabel, 0);
    gapRow.input.helpTip = getLabel("tooltip.footerGap");
    ui.footerGapRow = gapRow.row;
    ui.footerGapInput = gapRow.input;

    var cornerRow = addNumberRow(footerPanel, "fieldLabel.cornerRadius", 0, 4, unitLabel);
    ui.footerCornerRow = cornerRow.row;
    ui.footerCornerInput = cornerRow.input;
}

/**
 * ［実コンテンツ領域］タブ（オフセット・列行・セル・区切り線）を組み立てる
 * @param {object} ui UI オブジェクト
 * @param {Tab} contentPanel ［実コンテンツ領域］タブ
 * @returns {void}
 */
function buildContentRegionTab(ui, contentPanel) {
    var unitLabel = ui.context.unitLabel;

    /* オフセット / Offset */
    var offsetPanel = addPanel(contentPanel, withUnit("panel.offset", unitLabel));
    ui.offsetSides = addLinkedSideInputs(offsetPanel, 0, 4);
    var offsetFields = [ui.offsetSides.top, ui.offsetSides.bottom, ui.offsetSides.left, ui.offsetSides.right];
    for (var i = 0; i < offsetFields.length; i++) setRoundedDisplayValue(offsetFields[i], ui.context.defaults.offset);
    ui.btnAutoOffset = addAutoButton(offsetPanel, "tooltip.autoOffset");
    ui.btnAutoOffset.alignment = "center";

    /* 列・行 / Columns and rows */
    var rowColPanel = addPanel(contentPanel, getLabel("panel.rowCol"));
    var colCountRow = addNumberRow(rowColPanel, "fieldLabel.columnCount", 2, 5, null, 1, true);
    ui.colCountInput = colCountRow.input;
    ui.charCountInput = addNumberInput(colCountRow.row, 0, 4, 1, true);
    ui.charCountInput.helpTip = getLabel("tooltip.characterCount");
    colCountRow.row.add("statictext", undefined, getLabel("unit.characters"));

    var colGapRow = addNumberRow(rowColPanel, "fieldLabel.gap", ui.context.defaults.columnGap, 5, unitLabel, 0);
    ui.colGapInput = colGapRow.input;
    ui.btnAutoColumnGap = addAutoButton(colGapRow.row, "tooltip.autoColumnGap");

    ui.rowCountInput = addNumberRow(rowColPanel, "fieldLabel.rowCount", 1, 5, null, 1, true).input;

    var rowGapRow = addNumberRow(rowColPanel, "fieldLabel.gap", 0, 5, unitLabel, 0);
    ui.rowGapInput = rowGapRow.input;
    ui.gapLinkToggle = addLinkToggle(rowGapRow.row, true);
    ui.gapLinkToggle.helpTip = getLabel("tooltip.linkGaps");

    /* セル / Cells */
    var cellsPanel = addPanel(contentPanel, getLabel("panel.cells"));
    var cellTypeRadios = addRadioButtons(addRow(cellsPanel), ["radio.cellFill", "radio.cellTextFrame"]);
    ui.rbCellFill = cellTypeRadios[0];
    ui.rbCellTextFrame = cellTypeRadios[1];
    ui.threadCheck = addCheckbox(cellsPanel, "checkbox.threadFrames", false, "tooltip.threadFrames");
    ui.sampleTextRow = addRow(cellsPanel);
    var sampleRadios = addRadioButtons(ui.sampleTextRow, ["radio.sampleNone", "radio.sampleProse", "radio.sampleDummy"], "tooltip.sampleText");
    ui.rbSampleNone = sampleRadios[0];
    ui.rbSampleProse = sampleRadios[1];
    ui.rbSampleDummy = sampleRadios[2];

    /* 区切り線 / Dividers */
    var dividerPanel = addPanel(contentPanel, getLabel("panel.divider"));
    var dividerCheckRow = addRow(dividerPanel);
    ui.dividerCheck = addCheckbox(dividerCheckRow, "checkbox.drawDividers", true, "tooltip.dividers");
    ui.dividerWeightInput = addNumberInput(dividerCheckRow, 0.3, 5);
    ui.dividerWeightUnitLabel = dividerCheckRow.add("statictext", undefined, "pt");
    ui.dividerLineTypeRow = addRow(dividerPanel);
    var lineTypeRadios = addRadioButtons(ui.dividerLineTypeRow, ["radio.lineSolid", "radio.lineDashed", "radio.lineDotted"]);
    ui.rbLineSolid = lineTypeRadios[0];
    ui.rbLineDashed = lineTypeRadios[1];
    ui.rbLineDotted = lineTypeRadios[2];
    ui.rbLineDashed.helpTip = ui.rbLineDotted.helpTip = getLabel("tooltip.dividerStyle");
}

/**
 * ボタンエリア（左：一括自動調整、右：キャンセル・OK）を組み立てる
 * @param {object} ui UI オブジェクト
 * @param {Window} dlg ダイアログ
 * @returns {void}
 */
function buildButtonRow(ui, dlg) {
    var buttonRow = addButtonRow(dlg);
    ui.btnAutoAdjustAll = buttonRow.leftGroup.add("button", undefined, getLabel("button.autoAdjustAll"));
    ui.btnAutoAdjustAll.helpTip = getLabel("tooltip.autoAdjustAll");
    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    alignRightOnlyButtonRow(buttonRow);
}

// =========================================
// ダイアログの値 / Dialog values
// =========================================

/**
 * 入力欄を数値で読む（0・空欄・読めない値は代わりの値）
 * @param {EditText} editText 入力欄
 * @param {number} fallback 代わりの値
 * @returns {number} 数値
 */
function readNumber(editText, fallback) {
    return parseFloat(editText.text) || fallback;
}

/**
 * 値を丸めずに持ったまま、欄には小数点以下 1 桁までを表示する
 * @param {EditText} editText 入力欄
 * @param {number} value 丸める前の値
 * @returns {void}
 */
function setRoundedDisplayValue(editText, value) {
    editText.text = String(roundTo(value, 1));
    editText.exactValue = value;
    editText.exactValueText = editText.text;  /* 表示が書き換えられたかの目印 / Detects user edits */
}

/**
 * setRoundedDisplayValue() で入れた欄を読む
 * 表示がそのままなら丸める前の値、書き換えられていれば入力値（0・読めない値は 0）
 * @param {EditText} editText 入力欄
 * @returns {number} 数値
 */
function readExactValue(editText) {
    if (editText.exactValueText !== undefined && editText.text === editText.exactValueText) return editText.exactValue;
    return parseFloat(editText.text) || 0;
}

/**
 * 列数・行数を読む（1 未満や読めない値は 1）
 * @param {EditText} editText 入力欄
 * @returns {number} 1 以上の整数
 */
function readCount(editText) {
    var count = parseInt(editText.text, 10);
    return (isNaN(count) || count < 1) ? 1 : count;
}

/**
 * 線幅を pt で読む（mm 入力中なら換算、読めなければ 0.3pt）
 * @param {object} ui UI オブジェクト
 * @param {EditText} editText 入力欄
 * @returns {number} 線幅（pt）
 */
function readStrokeWeight(ui, editText) {
    var weight = parseFloat(editText.text);
    if (isNaN(weight)) return 0.3;
    return ui.strokeUnitIsMm ? weight / MM_PER_POINT : weight;
}

/**
 * タイトルエリアを描くか
 * @param {object} ui UI オブジェクト
 * @returns {boolean} 塗りか罫線がオンなら true
 */
function isTitleOn(ui) {
    return ui.titleFillCheck.value || ui.titleStrokeCheck.value;
}

/**
 * フッターのコラムエリアを描くか
 * @param {object} ui UI オブジェクト
 * @returns {boolean} 塗りか罫線がオンなら true
 */
function isFooterOn(ui) {
    return ui.footerFillCheck.value || ui.footerStrokeCheck.value;
}

/**
 * 選択中のタイトルエリアの位置を取得する
 * @param {object} ui UI オブジェクト
 * @returns {string} "top" / "bottom" / "left" / "right"
 */
function getSelectedTitlePosition(ui) {
    if (ui.rbTitleBottom.value) return "bottom";
    if (ui.rbTitleLeft.value) return "left";
    if (ui.rbTitleRight.value) return "right";
    return "top";
}

/**
 * 選択中の線端を取得する
 * @param {object} ui UI オブジェクト
 * @returns {string} "none" / "round" / "project"
 */
function getSelectedCapStyle(ui) {
    if (ui.rbCapRound.value) return "round";
    if (ui.rbCapProjecting.value) return "project";
    return "none";
}

/**
 * 選択中の区切り線の種類を取得する
 * @param {object} ui UI オブジェクト
 * @returns {string} "solid" / "dashed" / "dotted"
 */
function getSelectedDividerLineType(ui) {
    if (ui.rbLineDashed.value) return "dashed";
    if (ui.rbLineDotted.value) return "dotted";
    return "solid";
}

/**
 * 基本テキストの入力値を pt にする（Q/H 入力中なら換算）
 * @param {object} ui UI オブジェクト
 * @param {number} value 入力値
 * @returns {number} ポイント値
 */
function textInputToPoints(ui, value) {
    return ui.rbUnitTextQ.value ? value / Q_PER_POINT : value;
}

/**
 * 文字サイズと行送りを定規の単位で読む（行送りが読めなければ文字サイズの 1.5 倍）
 * @param {object} ui UI オブジェクト
 * @returns {{fontSize: number, leading: number}} 文字サイズと行送り
 */
function readFontMetrics(ui) {
    var fontSize = readNumber(ui.baseFontSizeInput, DEFAULT_FONT_SIZE);
    var leading = parseFloat(ui.leadingInput.text);
    if (isNaN(leading) || leading <= 0) leading = fontSize * 1.5;
    var pointsPerUnit = ui.context.pointsPerUnit;
    return {
        fontSize: textInputToPoints(ui, fontSize) / pointsPerUnit,
        leading: textInputToPoints(ui, leading) / pointsPerUnit
    };
}

/**
 * 入力中の値から版面の範囲を求める（スプレッド座標・定規の単位）
 * @param {object} ui UI オブジェクト
 * @returns {Array<number>} [上, 左, 下, 右]
 */
function readTypeAreaBounds(ui) {
    return insetBounds(ui.context.pageBounds,
        readNumber(ui.marginTopInput, 0), readNumber(ui.marginLeftInput, 0),
        readNumber(ui.marginBottomInput, 0), readNumber(ui.marginRightInput, 0));
}

/**
 * 入力中の値から実コンテンツ領域を求める
 * @param {object} ui UI オブジェクト
 * @returns {Array<number>} [上, 左, 下, 右]
 */
function readContentBounds(ui) {
    var afterTitle = subtractTitleArea(readTypeAreaBounds(ui), isTitleOn(ui), readNumber(ui.titleLengthInput, 0), getSelectedTitlePosition(ui));
    return subtractFooterArea(afterTitle, isFooterOn(ui), readNumber(ui.footerHeightInput, 0), readNumber(ui.footerGapInput, 0));
}

/**
 * 入力中の値から、オフセットを差し引いたグリッドの範囲を求める
 * @param {object} ui UI オブジェクト
 * @returns {Array<number>} [上, 左, 下, 右]
 */
function readGridBounds(ui) {
    var sides = ui.offsetSides;
    return insetBounds(readContentBounds(ui),
        readExactValue(sides.top), readExactValue(sides.left), readExactValue(sides.bottom), readExactValue(sides.right));
}

/**
 * ダイアログの入力値を描画用の設定値にまとめる
 * @param {object} ui UI オブジェクト
 * @returns {object} 描画に使う設定値（長さは定規の単位、線幅・文字サイズは pt）
 */
function readSettings(ui) {
    var isQ = ui.rbUnitTextQ.value;
    var marginTop = parseNumberOr(ui.marginTopInput.text, 0);
    var cornerRadius = parseNumberOr(ui.cornerRadiusInput.text, 0);
    var fontSizeInput = parseFloat(ui.baseFontSizeInput.text);
    var leadingText = ui.leadingInput.text;
    var leadingInput = parseFloat(leadingText);
    var frameSides = ui.pageFrameSides;
    var offsetSides = ui.offsetSides;

    return {
        marginTop: marginTop,
        marginBottom: parseNumberOr(ui.marginBottomInput.text, marginTop),
        marginLeft: parseNumberOr(ui.marginLeftInput.text, marginTop),
        marginRight: parseNumberOr(ui.marginRightInput.text, marginTop),

        applyPageMargins: ui.applyMarginsCheck.value,

        typeAreaBorder: ui.borderCheck.value,
        borderWeight: readStrokeWeight(ui, ui.borderWeightInput),
        borderCornerRadius: cornerRadius,
        borderExtension: parseNumberOr(ui.extensionInput.text, 0),
        capStyle: getSelectedCapStyle(ui),

        titleFill: ui.titleFillCheck.value,
        titleStroke: ui.titleStrokeCheck.value,
        titleLength: parseNumberOr(ui.titleLengthInput.text, 0),
        titlePosition: getSelectedTitlePosition(ui),
        titleExtension: parseNumberOr(ui.titleExtensionInput.text, 0),
        titleCornerRadius: cornerRadius,

        footerFill: ui.footerFillCheck.value,
        footerStroke: ui.footerStrokeCheck.value,
        footerWeight: readStrokeWeight(ui, ui.footerWeightInput),
        footerHeight: parseNumberOr(ui.footerHeightInput.text, 0),
        footerGap: parseNumberOr(ui.footerGapInput.text, 0),
        footerCornerRadius: parseNumberOr(ui.footerCornerInput.text, 0),

        pageFrame: ui.pageFrameCheck.value,
        pageFrameBleed: ui.pageFrameBleedCheck.value,
        pageFrameTop: parseNumberOr(frameSides.top.text, 0),
        pageFrameBottom: parseNumberOr(frameSides.bottom.text, 0),
        pageFrameLeft: parseNumberOr(frameSides.left.text, 0),
        pageFrameRight: parseNumberOr(frameSides.right.text, 0),
        /* 額縁の内側の角丸（pt）。オフなら 0 / Opening corner radius in pt, 0 when off */
        pageFrameCornerRadiusPt: ui.pageFrameCornerCheck.value ? parseNumberOr(ui.pageFrameCornerInput.text, 0) : 0,

        offsetTop: readExactValue(offsetSides.top),
        offsetBottom: readExactValue(offsetSides.bottom),
        offsetLeft: readExactValue(offsetSides.left),
        offsetRight: readExactValue(offsetSides.right),
        colCount: readCount(ui.colCountInput),
        colGap: parseNumberOr(ui.colGapInput.text, 0),
        rowCount: readCount(ui.rowCountInput),
        rowGap: parseNumberOr(ui.rowGapInput.text, 0),

        cellFill: ui.rbCellFill.value,
        cellTextFrame: ui.rbCellTextFrame.value,
        threadFrames: ui.threadCheck.value,
        sampleProse: ui.rbSampleProse.value,
        sampleDummy: ui.rbSampleDummy.value,

        dividers: ui.dividerCheck.value,
        dividerLineType: getSelectedDividerLineType(ui),
        dividerWeight: readStrokeWeight(ui, ui.dividerWeightInput),

        fontSizePt: isQ ? fontSizeInput / Q_PER_POINT : fontSizeInput,
        /* 行送りは文字列のまま（"auto" を通す）/ Keep leading as text so "auto" passes through */
        leading: (!isNaN(leadingInput) && leadingInput > 0 && isQ) ? String(leadingInput / Q_PER_POINT) : leadingText,
        showTempGrid: ui.showTempGridCheck.value,
        pointsPerUnit: ui.context.pointsPerUnit,

        targetLayer: null,  /* 描画先レイヤー。null ならアクティブレイヤー / Target layer, null for the active layer */
        gridOnly: false     /* 仮グリッドだけを描く / Draw the temp grid only */
    };
}

// =========================================
// ダイアログの更新 / Dialog updates
// =========================================

/**
 * コントロールの有効／無効をまとめて切り替える
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function updateEnabledStates(ui) {
    /* 版面：線幅と角丸はタイトルエリアでも使う / Weight and radius are shared with the title area */
    var borderOn = ui.borderCheck.value;
    var extensionIsZero = readNumber(ui.extensionInput, 0) === 0;
    var borderWeightOn = borderOn || ui.titleStrokeCheck.value;
    setNumberInputEnabled(ui.borderWeightInput, borderWeightOn);
    ui.borderWeightUnitLabel.enabled = borderWeightOn;
    ui.cornerRadiusRow.enabled = (borderOn && extensionIsZero) || ui.titleFillCheck.value;
    ui.extensionRow.enabled = borderOn;
    ui.capStyleRow.enabled = borderOn && !extensionIsZero;

    /* タイトルエリア / Title area */
    var titleOn = isTitleOn(ui);
    ui.titleLengthRow.enabled = titleOn;
    ui.titlePositionRow.enabled = titleOn;
    ui.titleExtensionRow.enabled = ui.titleStrokeCheck.value;

    /* フッターのコラムエリア / Footer column area */
    var footerOn = isFooterOn(ui);
    setNumberInputEnabled(ui.footerWeightInput, ui.footerStrokeCheck.value);
    ui.footerWeightUnitLabel.enabled = ui.footerStrokeCheck.value;
    ui.footerHeightRow.enabled = footerOn;
    ui.footerGapRow.enabled = footerOn;
    ui.footerCornerRow.enabled = footerOn;

    /* 額縁 / Page frame */
    ui.pageFrameBleedCheck.enabled = ui.pageFrameCheck.value;
    ui.pageFrameSides.row.enabled = ui.pageFrameCheck.value;
    ui.pageFrameCornerRow.enabled = ui.pageFrameCheck.value;
    setNumberInputEnabled(ui.pageFrameCornerInput, ui.pageFrameCornerCheck.value);
    ui.pageFrameCornerUnitLabel.enabled = ui.pageFrameCornerCheck.value;

    /* 間隔：列数・行数が 1 のときは無効 / Gaps are unavailable for a single column or row */
    var hasColumns = readCount(ui.colCountInput) > 1;
    var hasRows = readCount(ui.rowCountInput) > 1;
    setNumberInputEnabled(ui.colGapInput, hasColumns);
    setNumberInputEnabled(ui.rowGapInput, hasRows);
    setLinkToggleEnabled(ui.gapLinkToggle, hasColumns && hasRows);

    /* 区切り線：両方の間隔が 0 のときは無効 / Dividers need a gap */
    var hasGap = readNumber(ui.colGapInput, 0) > 0 || readNumber(ui.rowGapInput, 0) > 0;
    var dividersOn = hasGap && ui.dividerCheck.value;
    ui.dividerCheck.enabled = hasGap;
    setNumberInputEnabled(ui.dividerWeightInput, dividersOn);
    ui.dividerWeightUnitLabel.enabled = dividersOn;
    ui.dividerLineTypeRow.enabled = dividersOn;

    /* セル / Cells */
    ui.threadCheck.enabled = ui.rbCellTextFrame.value;
    ui.sampleTextRow.enabled = ui.rbCellTextFrame.value;

    /* ∧∨は自作描画なので、有効／無効が変わったものを描き直す / redraw steppers whose state changed */
    redrawChangedSteppers(ui.dlg);
}

/**
 * 文字サイズと列幅から、1 列に入る文字数を計算して表示する
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function updateCharCount(ui) {
    var grid = readGridBounds(ui);
    var gridWidth = grid[3] - grid[1];
    if (gridWidth <= 0 || grid[2] - grid[0] <= 0) {
        ui.charCountInput.text = "0";
        return;
    }
    var colCount = parseInt(ui.colCountInput.text, 10) || 1;
    var cellWidth = (gridWidth - readNumber(ui.colGapInput, 0) * (colCount - 1)) / colCount;
    ui.charCountInput.text = String(Math.floor(cellWidth / readFontMetrics(ui).fontSize));
}

/**
 * プレビューを描き直す（プレビューは常にオン）
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function updatePreview(ui) {
    var doc = ui.context.doc;
    clearPreviewLayer(doc);
    var settings = readSettings(ui);
    settings.targetLayer = getOrCreateWorkLayer(doc, PREVIEW_LAYER_NAME);
    drawLayout(settings);
}

/**
 * 値が変わったあとの共通処理（有効／無効・文字数・プレビュー）
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function handleChange(ui) {
    updateEnabledStates(ui);
    updateCharCount(ui);
    updatePreview(ui);
}

/**
 * 基本テキストの入力単位を pt と Q/H で切り替え、入力値を換算する
 * @param {object} ui UI オブジェクト
 * @param {boolean} toQ Q/H に切り替えるなら true
 * @returns {void}
 */
function switchFontUnit(ui, toQ) {
    if (toQ === ui.fontUnitIsQ) return;
    var factor = toQ ? Q_PER_POINT : (1 / Q_PER_POINT);
    var fields = [ui.baseFontSizeInput, ui.leadingInput];
    for (var i = 0; i < fields.length; i++) {
        var value = parseFloat(fields[i].text);
        if (value > 0) fields[i].text = String(roundTo(value * factor, 2));
    }
    ui.fontSizeUnitLabel.text = toQ ? "Q" : "pt";
    ui.leadingUnitLabel.text = toQ ? "H" : "pt";
    ui.fontUnitIsQ = toQ;
}

/**
 * 線幅の入力単位を pt と mm で切り替え、入力値を換算する
 * @param {object} ui UI オブジェクト
 * @param {boolean} toMm mm に切り替えるなら true
 * @returns {void}
 */
function switchStrokeUnit(ui, toMm) {
    if (toMm === ui.strokeUnitIsMm) return;
    var factor = toMm ? MM_PER_POINT : (1 / MM_PER_POINT);
    var unitText = toMm ? "mm" : "pt";
    var fields = [ui.borderWeightInput, ui.footerWeightInput, ui.dividerWeightInput];
    for (var i = 0; i < fields.length; i++) {
        var value = parseFloat(fields[i].text);
        if (value > 0) fields[i].text = String(roundTo(value * factor, 3));
    }
    ui.borderWeightUnitLabel.text = unitText;
    ui.footerWeightUnitLabel.text = unitText;
    ui.dividerWeightUnitLabel.text = unitText;
    ui.strokeUnitIsMm = toMm;
}

/**
 * 1 列の文字数から列の間隔を求めて入力する（連動中なら行の間隔も）
 * @param {object} ui UI オブジェクト
 * @param {number} charCount 1 列の文字数
 * @returns {void}
 */
function applyGapForCharCount(ui, charCount) {
    var grid = readGridBounds(ui);
    var colCount = parseInt(ui.colCountInput.text, 10) || 1;
    var newGap = 0;
    if (colCount > 1) {
        var cellWidth = charCount * readFontMetrics(ui).fontSize;
        newGap = roundTo(((grid[3] - grid[1]) - cellWidth * colCount) / (colCount - 1), 3);
        if (newGap < 0) newGap = 0;
    }
    ui.colGapInput.text = String(newGap);
    if (ui.gapLinkToggle.value) ui.rowGapInput.text = ui.colGapInput.text;
}

// =========================================
// 自動調整 / Auto adjust
// =========================================

/**
 * タイトルエリアの長さを仮グリッドの線に合わせる
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function autoAdjustTitleLength(ui) {
    var titleLength = readNumber(ui.titleLengthInput, 0);
    if (titleLength <= 0) return;
    ui.titleLengthInput.text = String(roundTo(snapToTempGrid(titleLength, readFontMetrics(ui)), 3));
}

/**
 * コラムエリアの上端が仮グリッドの線に乗るように高さを調整する
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function autoAdjustFooterHeight(ui) {
    var footerHeight = readNumber(ui.footerHeightInput, 0);
    if (footerHeight <= 0) return;
    var fontMetrics = readFontMetrics(ui);
    var typeArea = readTypeAreaBounds(ui);
    var footerGap = readNumber(ui.footerGapInput, 0);

    var footerTop = typeArea[2] - footerGap - footerHeight;
    var snappedFooterTop = typeArea[0] + snapToTempGrid(footerTop - typeArea[0], fontMetrics);
    var newHeight = typeArea[2] - footerGap - snappedFooterTop;
    if (newHeight <= 0) newHeight = fontMetrics.fontSize;
    ui.footerHeightInput.text = String(roundTo(newHeight, 3));
}

/**
 * オフセットを調整する。左右は文字サイズの倍数、天地は仮グリッドの線に合わせる
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function autoAdjustOffsets(ui) {
    var sides = ui.offsetSides;
    var fontMetrics = readFontMetrics(ui);

    /* 左右 / Left and right */
    var horizontalInputs = [sides.left, sides.right];
    for (var i = 0; i < horizontalInputs.length; i++) {
        var value = readExactValue(horizontalInputs[i]);
        if (value === 0) continue;
        var charCount = Math.round(value / fontMetrics.fontSize);
        if (charCount < 1) charCount = 1;
        setRoundedDisplayValue(horizontalInputs[i], charCount * fontMetrics.fontSize);
    }

    /* 天地：実コンテンツ領域の端＋オフセットが仮グリッドの線に乗るように / Snap top and bottom to temp grid lines */
    var typeAreaTop = readTypeAreaBounds(ui)[0];
    var content = readContentBounds(ui);

    var offsetTop = readExactValue(sides.top);
    if (offsetTop > 0) {
        var snappedTop = typeAreaTop + snapToTempGrid(content[0] + offsetTop - typeAreaTop, fontMetrics);
        setRoundedDisplayValue(sides.top, Math.max(snappedTop - content[0], 0));
    }

    var offsetBottom = readExactValue(sides.bottom);
    if (offsetBottom > 0) {
        var snappedBottom = typeAreaTop + snapToTempGrid(content[2] - offsetBottom - typeAreaTop, fontMetrics);
        setRoundedDisplayValue(sides.bottom, Math.max(content[2] - snappedBottom, 0));
    }
}

/**
 * ［文字］の値から列の間隔を計算し直す
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function autoAdjustColumnGap(ui) {
    applyGapForCharCount(ui, readCount(ui.charCountInput));
}

/**
 * 有効な［自動調整］をまとめて実行する
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function autoAdjustAll(ui) {
    if (isTitleOn(ui)) autoAdjustTitleLength(ui);
    if (isFooterOn(ui)) autoAdjustFooterHeight(ui);
    autoAdjustOffsets(ui);
    updateCharCount(ui);  /* 列の間隔は調整後の文字数から / Gap uses the updated character count */
    autoAdjustColumnGap(ui);
}

// =========================================
// イベント / Events
// =========================================

/**
 * 入力中と確定時の両方に同じ処理を結び付ける
 * @param {EditText} editText 入力欄
 * @param {Function} handler 処理
 * @returns {void}
 */
function bindEditHandler(editText, handler) {
    editText.onChanging = handler;
    editText.onChange = handler;
}

/**
 * 天地左右の入力欄を連動させる
 * @param {object} sides addLinkedSideInputs() の戻り値
 * @param {Function} afterChange 連動後の処理
 * @returns {void}
 */
function bindLinkedSides(sides, afterChange) {
    var fields = [sides.top, sides.bottom, sides.left, sides.right];
    for (var i = 0; i < fields.length; i++) {
        (function (source) {
            bindEditHandler(source, function () {
                if (sides.linkToggle.value) {
                    for (var j = 0; j < fields.length; j++) {
                        /* 丸める前の値も一緒に写す / Copy the unrounded value too */
                        fields[j].text = source.text;
                        fields[j].exactValue = source.exactValue;
                        fields[j].exactValueText = source.exactValueText;
                    }
                }
                afterChange();
            });
        })(fields[i]);
    }
}

/**
 * ダイアログのコントロールにイベントを結び付ける
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function bindDialogEvents(ui) {
    var refresh = function () { handleChange(ui); };
    var i;

    var clickControls = [
        ui.showTempGridCheck,
        ui.borderCheck, ui.rbCapNone, ui.rbCapRound, ui.rbCapProjecting,
        ui.titleFillCheck, ui.titleStrokeCheck, ui.rbTitleTop, ui.rbTitleBottom, ui.rbTitleLeft, ui.rbTitleRight,
        ui.footerFillCheck, ui.footerStrokeCheck,
        ui.pageFrameCheck, ui.pageFrameBleedCheck, ui.pageFrameCornerCheck,
        ui.threadCheck, ui.dividerCheck, ui.rbLineSolid, ui.rbLineDashed, ui.rbLineDotted
    ];
    for (i = 0; i < clickControls.length; i++) clickControls[i].onClick = refresh;

    var editControls = [
        ui.baseFontSizeInput, ui.leadingInput,
        ui.borderWeightInput, ui.cornerRadiusInput, ui.extensionInput,
        ui.titleLengthInput, ui.titleExtensionInput,
        ui.footerWeightInput, ui.footerHeightInput, ui.footerGapInput, ui.footerCornerInput,
        ui.colCountInput, ui.rowCountInput, ui.dividerWeightInput, ui.pageFrameCornerInput
    ];
    for (i = 0; i < editControls.length; i++) bindEditHandler(editControls[i], refresh);

    bindLinkedSides(ui.marginSides, refresh);
    bindLinkedSides(ui.pageFrameSides, refresh);
    bindLinkedSides(ui.offsetSides, refresh);

    /* 入力単位 / Input units */
    ui.rbUnitAllPt.onClick = function () { switchFontUnit(ui, false); switchStrokeUnit(ui, false); refresh(); };
    ui.rbUnitStrokeMm.onClick = function () { switchFontUnit(ui, false); switchStrokeUnit(ui, true); refresh(); };
    ui.rbUnitTextQ.onClick = function () { switchFontUnit(ui, true); switchStrokeUnit(ui, true); refresh(); };

    /* テキストフレームを選んだら連結とサンプル文をオンに / Text frames default to threaded with sample text */
    ui.rbCellFill.onClick = ui.rbCellTextFrame.onClick = function () {
        if (ui.rbCellTextFrame.value) {
            ui.threadCheck.value = true;
            ui.rbSampleProse.value = true;
            ui.rbSampleNone.value = false;
            ui.rbSampleDummy.value = false;
        }
        refresh();
    };

    /* 列と行の間隔の連動 / Linked column and row gaps */
    bindEditHandler(ui.colGapInput, function () {
        if (ui.gapLinkToggle.value) ui.rowGapInput.text = ui.colGapInput.text;
        refresh();
    });
    bindEditHandler(ui.rowGapInput, function () {
        if (ui.gapLinkToggle.value) ui.colGapInput.text = ui.rowGapInput.text;
        refresh();
    });

    /* 文字数から間隔を逆算（入力中の文字数は書き戻さない）/ Derive the gap without rewriting the count being typed */
    bindEditHandler(ui.charCountInput, function () {
        var charCount = parseInt(ui.charCountInput.text, 10);
        if (isNaN(charCount) || charCount < 1) return;
        applyGapForCharCount(ui, charCount);
        updateEnabledStates(ui);
        updatePreview(ui);
    });

    /* 相対：前回からの増減分を四辺のマージンに加える / Relative: add the change to all margins */
    bindEditHandler(ui.relativeInput, function () {
        var relativeValue = readNumber(ui.relativeInput, 0);
        var delta = relativeValue - ui.lastRelativeValue;
        ui.lastRelativeValue = relativeValue;
        var marginInputs = [ui.marginTopInput, ui.marginBottomInput, ui.marginLeftInput, ui.marginRightInput];
        for (var j = 0; j < marginInputs.length; j++) {
            marginInputs[j].text = String(roundTo(readNumber(marginInputs[j], 0) + delta, 2));
        }
        refresh();
    });

    /* 自動調整 / Auto adjust */
    ui.btnAutoTitleLength.onClick = function () { autoAdjustTitleLength(ui); refresh(); };
    ui.btnAutoFooterHeight.onClick = function () { autoAdjustFooterHeight(ui); refresh(); };
    ui.btnAutoOffset.onClick = function () { autoAdjustOffsets(ui); refresh(); };
    ui.btnAutoColumnGap.onClick = function () { autoAdjustColumnGap(ui); refresh(); };
    ui.btnAutoAdjustAll.onClick = function () { autoAdjustAll(ui); refresh(); };
}

// =========================================
// メイン / Main
// =========================================

/**
 * スプレッド座標系でのページ境界を定規の単位で取得する
 * @param {Page} targetPage 対象ページ
 * @param {number} pointsPerUnit 1 単位あたりのポイント数
 * @returns {Array<number>} [上, 左, 下, 右]
 */
function getPageBoundsOnSpread(targetPage, pointsPerUnit) {
    var topLeft = targetPage.resolve(AnchorPoint.TOP_LEFT_ANCHOR, CoordinateSpaces.SPREAD_COORDINATES)[0];
    var bottomRight = targetPage.resolve(AnchorPoint.BOTTOM_RIGHT_ANCHOR, CoordinateSpaces.SPREAD_COORDINATES)[0];
    return [topLeft[1] / pointsPerUnit, topLeft[0] / pointsPerUnit, bottomRight[1] / pointsPerUnit, bottomRight[0] / pointsPerUnit];
}

/**
 * 入力欄の初期値を求める（すべて横の定規の単位）
 * @param {Page} targetPage 対象ページ
 * @param {Array<number>} pageBounds ページ [上, 左, 下, 右]
 * @param {{horizontal: number, vertical: number}} unitScale 1 単位あたりのポイント数
 * @returns {object} マージン・タイトルエリアの長さ・コラムエリアの高さ・オフセット・列の間隔
 */
function getDefaultInputs(targetPage, pageBounds, unitScale) {
    var marginPrefs = targetPage.marginPreferences;
    /* 天地のマージンは縦の単位なので横の単位に換算 / Top and bottom margins are in vertical units */
    var verticalToHorizontal = unitScale.vertical / unitScale.horizontal;
    var marginTop = roundTo(marginPrefs.top * verticalToHorizontal, 3);
    var marginBottom = roundTo(marginPrefs.bottom * verticalToHorizontal, 3);
    /* 左ページは内側（left）が右に来るので入れ替える / Swap inside and outside on left-hand pages */
    var isLeftPage = (targetPage.side === PageSideOptions.LEFT_HAND);
    var pointsPerUnit = unitScale.horizontal;
    return {
        marginTop: marginTop,
        marginBottom: marginBottom,
        marginLeft: roundTo(isLeftPage ? marginPrefs.right : marginPrefs.left, 3),
        marginRight: roundTo(isLeftPage ? marginPrefs.left : marginPrefs.right, 3),
        /* タイトルエリアの長さ：版面の高さの 1/5 / Title length: one fifth of the type area height */
        titleLength: roundTo(((pageBounds[2] - pageBounds[0]) - marginTop - marginBottom) / 5, 2),
        footerHeight: roundTo(millimetersToUnits(DEFAULT_LENGTHS_MM.footerHeight, pointsPerUnit), 2),
        offset: millimetersToUnits(DEFAULT_LENGTHS_MM.offset, pointsPerUnit),  /* 丸めずに持つ / Kept unrounded */
        columnGap: roundTo(millimetersToUnits(DEFAULT_LENGTHS_MM.columnGap, pointsPerUnit), 2)
    };
}

/**
 * ページの「マージン・段組」のマージンを書き換える
 * 左ページは left が内側（画面の右）なので入れ替え、単位の食い違いを避けて pt で入れる
 * @param {Page} targetPage 対象ページ
 * @param {object} settings 設定値（マージンは横の定規の単位）
 * @returns {void}
 */
function applyPageMargins(targetPage, settings) {
    var marginPrefs = targetPage.marginPreferences;
    var isLeftPage = (targetPage.side === PageSideOptions.LEFT_HAND);
    var toPoints = function (value) { return toPointString(value * settings.pointsPerUnit); };
    marginPrefs.top = toPoints(settings.marginTop);
    marginPrefs.bottom = toPoints(settings.marginBottom);
    marginPrefs.left = toPoints(isLeftPage ? settings.marginRight : settings.marginLeft);
    marginPrefs.right = toPoints(isLeftPage ? settings.marginLeft : settings.marginRight);
}

/**
 * 確定した設定で描画する（必要ならページのマージンも変更し、仮グリッドを残すなら専用レイヤーにも描く）
 * @param {Document} doc 対象ドキュメント
 * @param {object} settings 設定値
 * @param {boolean} keepTempGrid 仮グリッドを残すか
 * @returns {void}
 */
function drawFinalLayout(doc, settings, keepTempGrid) {
    if (settings.applyPageMargins) applyPageMargins(app.activeWindow.activePage, settings);
    activateOtherLayer(doc, [TEMP_GRID_LAYER_NAME, PREVIEW_LAYER_NAME]);
    drawLayout(settings);
    if (!keepTempGrid) return;

    settings.targetLayer = getOrCreateWorkLayer(doc, TEMP_GRID_LAYER_NAME);
    settings.gridOnly = true;
    drawLayout(settings);
}

/**
 * ドキュメントを確認し、設定ダイアログを表示してレイアウトを作成する
 * @returns {void}
 */
function main() {
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var doc = app.activeDocument;
    var page = app.activeWindow.activePage;
    var unitScale = measurePointsPerUnit(page);
    var pageBounds = getPageBoundsOnSpread(page, unitScale.horizontal);
    var ui = buildDialog({
        doc: doc,
        page: page,
        unitLabel: getRulerUnitLabel(doc.viewPreferences.horizontalMeasurementUnits),
        pointsPerUnit: unitScale.horizontal,
        pageBounds: pageBounds,
        defaults: getDefaultInputs(page, pageBounds, unitScale)
    });

    bindDialogEvents(ui);
    handleChange(ui);

    var confirmed = (ui.dlg.show() === 1);
    removePreviewLayer(doc);
    if (!confirmed) return;

    var settings = readSettings(ui);
    var keepTempGrid = ui.keepTempGridCheck.value && ui.showTempGridCheck.value;
    app.doScript(function () {
        drawFinalLayout(doc, settings, keepTempGrid);
    }, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, getLabel("dialog.title"));
}

main();
