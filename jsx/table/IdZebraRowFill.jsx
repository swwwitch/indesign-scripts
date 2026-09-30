#target indesign

/*

### 概要

選択した表セルに、選択範囲内の行の並びを基準に交互の塗り（縞模様）をプレビュー付きで適用します。
上から／左から指定した数の行・列は、塗りの対象外にできます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdZebraRowFill.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n20ff60f6b508

### Overview

Applies alternating fills (zebra striping) to the selected table cells, based on the row order within the selection, with a live preview.
A given number of rows from the top, and columns from the left, can be left out of the fill.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdZebraRowFill.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdZebraRowFill";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdZebraRowFill.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdZebraRowFill.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n20ff60f6b508"; /* 紹介記事 / article URL */

// Original idea
// KK sawa

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 濃淡（Tint）の初期値と範囲 / Initial value and range of the tint control */
var TINT_DEFAULT = 100;
var TINT_MIN     = 0;
var TINT_MAX     = 100;

/* 濃淡スライダーの刻み幅（通常 / Shift / Option）/ Tint slider steps (normal, Shift, Option) */
var TINT_STEP_NORMAL = 1;
var TINT_STEP_SHIFT  = 10;
var TINT_STEP_OPTION = 5;

/* ［行のスキップ］［列のスキップ］の初期値 / Initial number of rows and columns to skip */
var SKIP_COUNT_DEFAULT = 1;

// =========================================
// レイアウト設定 / Layout settings
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

/* 行パネル内の間隔と、濃淡グループの余白 [左,上,右,下] / Spacing inside a row panel and margins of its tint group */
var ROW_PANEL_SPACING  = 6;
var TINT_GROUP_MARGINS = [0, 20, 0, 10];

/* 入力欄の行・ボタン同士の間隔と、項目名と色見本の間隔 / Spacing inside control rows and next to the swatch chip */
var CONTROL_SPACING    = 8;
var SWATCH_ROW_SPACING = 6;

/* カラードロップダウン・濃淡入力欄・スライダー・スキップ入力欄の幅（px）
   / Widths of the color dropdown, tint field, slider and skip fields (px) */
var COLOR_DROPDOWN_WIDTH = 140;
var TINT_INPUT_WIDTH     = 50;
var TINT_SLIDER_WIDTH    = 150;
var SKIP_INPUT_WIDTH     = 30;

/* スキップのチェックボックスの幅。行と列で入力欄の位置をそろえる（px）
   / Width of the skip checkboxes so both fields line up (px) */
var SKIP_CHECKBOX_WIDTH = 80;

/* 色見本の一辺（px）/ Size of the swatch chip (px) */
var SWATCH_CHIP_SIZE = 18;

/* ダイアログの不透明度 / Dialog opacity */
var DIALOG_OPACITY = 0.97;

/**
 * 行に「∧∨＋入力欄」を隙間0で突き合わせる group を足し、∧∨を置く。入力欄は戻り値の .parent に続けて追加する
 * @param {Group} parentRow 追加先の行
 * @param {function} getNumberInput 対象の入力欄を返す関数
 * @param {object} stepOptions addStepper() に渡す増減の設定
 * @returns {Group} ∧∨をまとめた group
 */
function addStepperInputPair(parentRow, getNumberInput, stepOptions) {
    var pairGroup = parentRow.add("group");
    pairGroup.orientation = "row";
    pairGroup.alignChildren = ["left", "center"];
    pairGroup.spacing = 0; /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
    pairGroup.margins = 0;
    return addStepper(pairGroup, getNumberInput, stepOptions);
}

/**
 * ∧∨で増減したあと、手入力と同じ onChange を通す
 * @param {EditText} numberInput 増減した入力欄
 * @returns {void}
 */
function runInputOnChange(numberInput) {
    if (numberInput.onChange) numberInput.onChange();
}

/**
 * 入力欄と∧∨の有効／無効をまとめて切り替え、∧∨を描き直す
 * @param {EditText} numberInput 対象の入力欄（.stepperGroup に∧∨を控えておく）
 * @param {boolean} isEnabled 有効にするなら true
 * @returns {void}
 */
function setStepperInputEnabled(numberInput, isEnabled) {
    numberInput.enabled = isEnabled;
    numberInput.stepperGroup.enabled = isEnabled;
    /* 自作描画なのでディム表示を描き直す。表示前は最初の描画で反映されるので描き直さない
       redraw the custom-drawn dimming; before the window is shown the first draw picks it up */
    if (numberInput.window && numberInput.window.visible) redrawSteppersIn(numberInput.stepperGroup);
}

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
 * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
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

    /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
    numberInput.lastValidText = numberInput.text;
    numberInput.onChange = function () {
        var value = parseFloat(numberInput.text);
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
        var value = parseFloat(numberInput.text);
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
    return stepperGroup;
}

/**
 * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
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
 * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
 * @param {string|Object} labelRef - getLabel と同じ
 * @param {Object|Array} [placeholderValues] - getLabel と同じ
 * @returns {string} コロン付きの文言
 */
function labelText(labelRef, placeholderValues) {
    return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
}

/**
 * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
 * @param {string|Object} labelRef - getLabel と同じ
 * @param {string|number} value - コロンのあとに続ける値
 * @returns {string} 項目名と値をつないだ文字列
 */
function labelValueText(labelRef, value) {
    return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
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
 * 左のグループにボタンが無い（右のボタンだけの）とき、行を左右中央に並べ直す。
 * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
 * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
 * @returns {void}
 */
function centerButtonRowIfRightOnly(buttonRow) {
    if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
    var btnRowGroup = buttonRow.rowGroup;
    /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
    btnRowGroup.remove(buttonRow.leftGroup);
    btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
    btnRowGroup.alignment = ["center", "bottom"];
    btnRowGroup.alignChildren = ["center", "center"];
    buttonRow.leftGroup = null;
}

// ボタン行（再利用パーツ）ここまで / End of the reusable button row

// プレビュー画面モード（再利用パーツ） / Preview screen mode (reusable)

/**
 * 作業中のウィンドウがプレビュー画面モードかどうかを返す
 * @returns {boolean} プレビューなら true。ストーリーエディターなど screenMode の無いウィンドウでは false
 */
function isPreviewScreenMode() {
    try {
        var docWindow = app.activeWindow;
        return !!(docWindow && docWindow.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE);
    } catch (e) {
        /* ストーリーエディターのウィンドウには screenMode が無い / Story editor windows have no screenMode */
        return false;
    }
}

/**
 * 作業中のウィンドウの画面モードを、標準モードとプレビューで切り替える
 * @returns {void}
 */
function togglePreviewScreenMode() {
    try {
        var docWindow = app.activeWindow;
        if (!docWindow) return;
        docWindow.screenMode = isPreviewScreenMode() ? ScreenModeOptions.PREVIEW_OFF : ScreenModeOptions.PREVIEW_TO_PAGE;
    } catch (e) {
        /* screenMode の無いウィンドウでは何もしない / Nothing to do for windows without a screenMode */
    }
}

/**
 * 画面モードの切り替えボタンの文字（押すと切り替わる先）を返す
 * @returns {string} プレビュー中は「標準モード」、それ以外は「プレビュー」
 */
function getScreenModeButtonLabel() {
    return getLabel(isPreviewScreenMode() ? LABELS.button.screenModeNormal : LABELS.button.screenModePreview);
}

/**
 * 画面モードの切り替えボタンを足す
 * @param {Group} parent - ボタンを足す先（ふつうはボタン行の左のグループ）
 * @param {Function} [onToggle] - 切り替えたあとに呼ぶ関数
 * @returns {Button} 足したボタン
 */
function addScreenModeButton(parent, onToggle) {
    var btnScreenMode = parent.add("button", undefined, getScreenModeButtonLabel());
    btnScreenMode.helpTip = getLabel(LABELS.tooltip.screenMode);
    btnScreenMode.onClick = function () {
        togglePreviewScreenMode();
        btnScreenMode.text = getScreenModeButtonLabel();
        if (onToggle) onToggle();
    };
    return btnScreenMode;
}

// プレビュー画面モード（再利用パーツ）ここまで / End of the reusable preview screen mode

// 表の選択（再利用パーツ） / Table selection (reusable)

var TABLE_PARENT_LOOKUP_LIMIT = 20; /* 親をたどる上限（無限ループよけ） / max parent hops (guards against loops) */

/**
 * 親をたどって、いちばん近い表を返す（セル・行・列・セル内のテキストや挿入点に対応）
 * @param {Object} startItem - たどり始めるオブジェクト
 * @returns {Table|null} 見つかった表。表の中でなければ null
 */
function findParentTable(startItem) {
    var node = startItem;
    for (var i = 0; i < TABLE_PARENT_LOOKUP_LIMIT && node; i++) {
        try {
            var typeName = node.constructor.name;
            if (typeName === "Table") return node;
            if (typeName === "Document" || typeName === "Application") return null;
            node = node.parent;
        } catch (e) {
            return null;
        }
    }
    return null;
}

/**
 * 選択から対象の表を返す。表の中の選択ならその表、表を含むテキストフレームやテキストなら最初の表
 * @param {Object} selectionItem - app.selection[0] など
 * @returns {Table|null} 対象の表。見つからなければ null
 */
function getTableFromSelection(selectionItem) {
    if (!selectionItem) return null;
    var parentTable = findParentTable(selectionItem);
    if (parentTable) return parentTable;
    try {
        if (selectionItem.tables && selectionItem.tables.length > 0) return selectionItem.tables[0];
    } catch (e) {
        /* tables を持たない選択（画像など） / Selections without tables (images, etc.) */
    }
    return null;
}

/**
 * 選択からセルを1つずつの配列にして返す。セル選択・表の選択はその全セル、セル内のテキストや挿入点はそのセル
 * @param {Object} selectionItem - app.selection[0] など
 * @returns {Cell[]} セルの配列。表の外なら空配列
 */
function getSelectedCells(selectionItem) {
    var selectedCells = [];
    if (!selectionItem) return selectedCells;
    try {
        var typeName = selectionItem.constructor.name;
        if (typeName === "Cell" || typeName === "Table") {
            /* 複数セルの選択も1つの Cell で返るので .cells で展開する / A multi-cell selection is one Cell; expand it via .cells */
            var cellCollection = selectionItem.cells;
            for (var i = 0; i < cellCollection.length; i++) selectedCells.push(cellCollection[i]);
            return selectedCells;
        }
        var node = selectionItem;
        for (var j = 0; j < TABLE_PARENT_LOOKUP_LIMIT && node; j++) {
            var nodeType = node.constructor.name;
            if (nodeType === "Cell") {
                selectedCells.push(node);
                break;
            }
            if (nodeType === "Table" || nodeType === "Document" || nodeType === "Application") break;
            node = node.parent;
        }
    } catch (e) {
        /* 親をたどれない選択は空のまま / Leave empty when the parent chain cannot be followed */
    }
    return selectedCells;
}

/**
 * 2つの表が同じ表かどうかを返す
 * @param {Table} tableA - 表
 * @param {Table} tableB - 表
 * @returns {boolean} 同じ表なら true
 */
function isSameTable(tableA, tableB) {
    if (!tableA || !tableB) return false;
    try {
        return tableA.id === tableB.id && tableA.parent.id === tableB.parent.id;
    } catch (e) {
        return false;
    }
}

// 表の選択（再利用パーツ）ここまで / End of the reusable table selection


var LABELS = {
    dialog: {
        title: { ja: "行の塗りを交互に設定", en: "Alternating Row Fill" }
    },
    panel: {
        oddRows:  { ja: "奇数番目の行", en: "Odd-numbered Rows" },
        evenRows: { ja: "偶数番目の行", en: "Even-numbered Rows" },
        options:  { ja: "オプション", en: "Options" },
        skip:     { ja: "スキップ", en: "Skip" }
    },
    fieldLabel: {
        color: { ja: "カラー", en: "Color" },
        tint:  { ja: "濃淡", en: "Tint" }
    },
    checkbox: {
        swapRows:    { ja: "奇数行と偶数行を入れ替え", en: "Swap odd and even" },
        skipRows:    { ja: "行", en: "Rows" },
        skipColumns: { ja: "列", en: "Columns" }
    },
    swatch: {
        black: { ja: "黒", en: "Black" },
        paper: { ja: "紙色", en: "Paper" },
        none:  { ja: "なし", en: "None" }
    },
    button: {
        ok:                { ja: "OK", en: "OK" },
        cancel:            { ja: "キャンセル", en: "Cancel" },
        screenModePreview: { ja: "プレビュー", en: "Preview" },
        screenModeNormal:  { ja: "標準モード", en: "Normal Mode" }
    },
    tooltip: {
        oddRows: {
            ja: "選択範囲の上から数えて1・3・5…番目の行です。除外した行は数に入りません。",
            en: "The 1st, 3rd, 5th and so on row counted from the top of the selection. Skipped rows are not counted."
        },
        evenRows: {
            ja: "選択範囲の上から数えて2・4・6…番目の行です。除外した行は数に入りません。",
            en: "The 2nd, 4th, 6th and so on row counted from the top of the selection. Skipped rows are not counted."
        },
        colorDropdown: {
            ja: "「なし」と「紙色」では濃淡を指定できません。",
            en: "Tint is not available for None or Paper."
        },
        tintInput: {
            ja: "↑↓キーで1ずつ、Shift+↑↓キーで10ずつ増減します。",
            en: "Up/Down arrow keys change the value by 1, or by 10 with Shift."
        },
        tintSlider: {
            ja: "Shift+ドラッグで10%刻み、Option（Alt）+ドラッグで5%刻みになります。",
            en: "Shift-drag snaps to 10% steps, Option (Alt)-drag to 5% steps."
        },
        /* ステップボタン用 / for the stepper */
        stepUp: {
            ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
            en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
        },
        stepDown: {
            ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
            en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
        },
        stepUpInteger:   { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
        stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
        swapRows: {
            ja: "奇数行と偶数行のカラー・濃淡を入れ替えます。",
            en: "Exchanges the color and tint between the odd and even rows."
        },
        skipRows: {
            ja: "選択範囲の上から指定した行数を塗りの対象から外し、実行前のカラーに戻します。",
            en: "Leaves the given number of rows at the top of the selection out of the fill, restoring the color they had before the script ran."
        },
        skipColumns: {
            ja: "選択範囲の左から指定した列数を塗りの対象から外し、実行前のカラーに戻します。",
            en: "Leaves the given number of columns at the left of the selection out of the fill, restoring the color they had before the script ran."
        },
        screenMode: {
            ja: "ドキュメントウィンドウの画面モードを、標準モードとプレビューで切り替えます。",
            en: "Switches the document window between the Normal and Preview screen modes."
        }
    },
    alert: {
        noDocument:      { ja: "ドキュメントを開いてください。", en: "Please open a document." },
        noCellSelection: { ja: "セルを選択してください。", en: "Please select table cells." },
        noUsableSwatch:  { ja: "使用可能なカラーがありません。", en: "No usable colors are available." }
    },
    undo: {
        previewFills: { ja: "塗りプレビュー", en: "Fill Preview" },
        applyFills:   { ja: "セルの塗りを交互に設定", en: "Apply Alternating Fills" }
    }
};

// =========================================
// スウォッチ / Swatches
// =========================================

/* 特別扱いするスウォッチ名（角括弧・大文字小文字は無視して照合）/ Special swatch names (brackets and case are ignored) */
var SPECIAL_SWATCH_NAMES = {
    none:         ["None", "なし"],
    black:        ["Black", "ブラック", "黒"],
    paper:        ["Paper", "紙色", "紙"],
    registration: ["Registration", "レジストレーション", "トンボ用"]
};

/* 色見本に使う RGB（0〜1）/ RGB values (0-1) used for the swatch chip */
var CHIP_RGB_BLACK    = [0, 0, 0];
var CHIP_RGB_WHITE    = [1, 1, 1];
var CHIP_RGB_FALLBACK = [0.5, 0.5, 0.5];

/**
 * @typedef {object} SwatchChoice
 * @property {string} swatchName ドキュメント上のスウォッチ名
 * @property {string} displayName ドロップダウンに表示する名前
 * @property {string} swatchKind 特別なスウォッチの種別。通常のカラーは空文字
 */

/**
 * 比較しやすいようにスウォッチ名を正規化する（角括弧・前後の空白・大文字小文字を落とす）
 * @param {string} swatchName スウォッチ名
 * @returns {string} 正規化した名前
 */
function normalizeSwatchName(swatchName) {
    if (swatchName == null) return "";
    return String(swatchName).replace(/^\[|\]$/g, "").replace(/^\s+|\s+$/g, "").toLowerCase();
}

/**
 * スウォッチ名が特別なスウォッチのどれに当たるかを調べる
 * @param {string} swatchName スウォッチ名
 * @returns {string} "none" / "black" / "paper" / "registration"。どれでもなければ空文字
 */
function getSpecialSwatchKind(swatchName) {
    var normalizedName = normalizeSwatchName(swatchName);
    var swatchKind, names, i;
    for (swatchKind in SPECIAL_SWATCH_NAMES) {
        if (!SPECIAL_SWATCH_NAMES.hasOwnProperty(swatchKind)) continue;
        names = SPECIAL_SWATCH_NAMES[swatchKind];
        for (i = 0; i < names.length; i++) {
            if (normalizeSwatchName(names[i]) === normalizedName) return swatchKind;
        }
    }
    return "";
}

/**
 * 濃淡を指定できないスウォッチかどうかを判定する
 * @param {string} swatchKind getSpecialSwatchKind() が返した種別
 * @returns {boolean} 濃淡を指定できないなら true
 */
function isTintDisabledKind(swatchKind) {
    return swatchKind === "none" || swatchKind === "paper";
}

/**
 * カラー候補として表示するスウォッチ一覧を作る（レジストレーションは除く）
 * @returns {Array<SwatchChoice>} スウォッチ一覧
 */
function getSwatchChoices() {
    var swatches = app.activeDocument.swatches;
    var swatchChoices = [];
    var i, swatchName, swatchKind;

    for (i = 0; i < swatches.length; i++) {
        swatchName = String(swatches[i].name);
        swatchKind = getSpecialSwatchKind(swatchName);
        if (swatchKind === "registration") continue;
        swatchChoices.push({
            swatchName: swatchName,
            displayName: swatchKind ? getLabel("swatch." + swatchKind) : swatchName,
            swatchKind: swatchKind
        });
    }
    return swatchChoices;
}

/**
 * 名前からスウォッチを取得する
 * @param {string} swatchName スウォッチ名
 * @returns {Swatch|null} 見つかったスウォッチ。無ければ null
 */
function findSwatch(swatchName) {
    var swatch = app.activeDocument.swatches.itemByName(String(swatchName));
    return (swatch && swatch.isValid) ? swatch : null;
}

/**
 * 一覧の中からスウォッチ名に一致する位置を求める
 * @param {Array<SwatchChoice>} swatchChoices スウォッチ一覧
 * @param {string} swatchName 探すスウォッチ名
 * @returns {number} 見つかった位置。無ければ 0
 */
function findSwatchChoiceIndex(swatchChoices, swatchName) {
    for (var i = 0; i < swatchChoices.length; i++) {
        if (swatchChoices[i].swatchName === swatchName) return i;
    }
    return 0;
}

/**
 * スウォッチのカラー値を色見本用の RGB に変換する
 * @param {Swatch|null} swatch 対象のスウォッチ
 * @returns {Array<number>} 0〜1 の RGB 値
 */
function getSwatchChipRGB(swatch) {
    if (!swatch || !swatch.isValid) return CHIP_RGB_FALLBACK;

    var swatchKind = getSpecialSwatchKind(String(swatch.name));
    if (swatchKind === "paper" || swatchKind === "none") return CHIP_RGB_WHITE;
    if (swatchKind === "black" || swatchKind === "registration") return CHIP_RGB_BLACK;

    /* グラデーションや混合インキは色値を読めない / Gradients and mixed inks have no readable color value */
    try {
        var colorValues = swatch.colorValue;
        if (swatch.space === ColorSpace.RGB) {
            return [colorValues[0] / 255, colorValues[1] / 255, colorValues[2] / 255];
        }
        if (swatch.space === ColorSpace.CMYK) {
            var cyanValue    = colorValues[0] / 100;
            var magentaValue = colorValues[1] / 100;
            var yellowValue  = colorValues[2] / 100;
            var blackValue   = colorValues[3] / 100;
            return [
                (1 - cyanValue) * (1 - blackValue),
                (1 - magentaValue) * (1 - blackValue),
                (1 - yellowValue) * (1 - blackValue)
            ];
        }
    } catch (e) { }
    return CHIP_RGB_FALLBACK;
}

/**
 * 色見本をスウォッチのカラーで塗り直す
 * @param {Group} swatchChip 色見本のグループ
 * @param {Swatch|null} swatch 表示するスウォッチ
 * @returns {void}
 */
function paintSwatchChip(swatchChip, swatch) {
    var rgb = getSwatchChipRGB(swatch);
    swatchChip.graphics.backgroundColor = swatchChip.graphics.newBrush(
        swatchChip.graphics.BrushType.SOLID_COLOR,
        [rgb[0], rgb[1], rgb[2], 1]
    );
    if (swatchChip.window) swatchChip.window.update();
}

// =========================================
// 数値の正規化とキー操作 / Value handling
// =========================================

/**
 * 修飾キーに応じた濃淡スライダーの刻み幅を返す
 * @returns {number} 刻み幅
 */
function getTintSliderStep() {
    var keyboard = ScriptUI.environment.keyboardState;
    if (keyboard.shiftKey) return TINT_STEP_SHIFT;
    if (keyboard.altKey) return TINT_STEP_OPTION;
    return TINT_STEP_NORMAL;
}

/**
 * 濃淡を有効範囲の整数に収める
 * @param {number} value 入力された値
 * @returns {number} 整えた値
 */
function normalizeTint(value) {
    if (isNaN(value)) value = TINT_DEFAULT;
    if (value < TINT_MIN) value = TINT_MIN;
    if (value > TINT_MAX) value = TINT_MAX;
    return Math.round(value);
}

/**
 * 濃淡スライダーの値を、修飾キーに応じた刻み幅に合わせる
 * @param {number} value スライダーの値
 * @returns {number} 整えた値
 */
function snapTintToSliderStep(value) {
    var step = getTintSliderStep();
    return normalizeTint(Math.round(normalizeTint(value) / step) * step);
}

/**
 * スキップ数を 0 以上の整数に整える
 * @param {number} value 入力された値
 * @returns {number} 整えた値
 */
function normalizeSkipCount(value) {
    if (isNaN(value) || value < 0) return 0;
    return Math.round(value);
}

// =========================================
// 選択セルと初期値 / Selected cells & defaults
// =========================================

/**
 * @typedef {object} CellFill
 * @property {Swatch|null} color 塗りのカラー
 * @property {number|null} tint 濃淡。指定しない場合は null
 */

/**
 * セルの現在の塗りを控える（スキップ時に戻すため）
 * @param {Cell} cell 対象のセル
 * @returns {CellFill} 控えた塗り
 */
function captureCellFill(cell) {
    var fillColor = cell.fillColor;
    var fillTint = cell.fillTint;
    var swatchKind = (fillColor && fillColor.name) ? getSpecialSwatchKind(String(fillColor.name)) : "";
    var hasTint = !isTintDisabledKind(swatchKind) && !isNaN(fillTint) &&
        fillTint >= TINT_MIN && fillTint <= TINT_MAX;

    return { color: fillColor, tint: hasTint ? fillTint : null };
}

/**
 * 選択セルすべての塗りを控える
 * @param {Array<Cell>} cells 対象のセル
 * @returns {Array<CellFill>} セルと同じ並びの塗り
 */
function captureCellFills(cells) {
    var fills = [];
    for (var i = 0; i < cells.length; i++) {
        fills.push(captureCellFill(cells[i]));
    }
    return fills;
}

/**
 * 選択セルの塗り（カラー＋濃淡）を、使われている数の多い順に並べる
 * @param {Array<Cell>} cells 対象のセル
 * @param {{colorName: string, tint: number}} fallbackFill 塗りを読めなかったときに使う値
 * @returns {Array<{colorName: string, tint: number, count: number}>} 多い順の組み合わせ
 */
function countFillUsages(cells, fallbackFill) {
    var usageByKey = {};
    var usages = [];
    var i, fillColor, fillTint, colorName, tint, usageKey;

    for (i = 0; i < cells.length; i++) {
        fillColor = cells[i].fillColor;
        fillTint = cells[i].fillTint;
        colorName = (fillColor && fillColor.name) ? String(fillColor.name) : fallbackFill.colorName;
        tint = (!isNaN(fillTint) && fillTint >= TINT_MIN && fillTint <= TINT_MAX) ? fillTint : fallbackFill.tint;

        usageKey = colorName + "||" + tint;
        if (!usageByKey[usageKey]) {
            usageByKey[usageKey] = { colorName: colorName, tint: tint, count: 0 };
            usages.push(usageByKey[usageKey]);
        }
        usageByKey[usageKey].count++;
    }

    usages.sort(function (a, b) { return b.count - a.count; });
    return usages;
}

/**
 * 選択セルでよく使われている塗りから、奇数行・偶数行の初期値を決める
 * @param {Array<Cell>} cells 対象のセル
 * @param {string} fallbackColorName 塗りを読めなかったときに使うスウォッチ名
 * @returns {{odd: {colorName: string, tint: number}, even: {colorName: string, tint: number}}} 初期値
 */
function detectDefaultFills(cells, fallbackColorName) {
    var fallbackFill = { colorName: fallbackColorName, tint: TINT_DEFAULT };
    var usages = countFillUsages(cells, fallbackFill);
    var mostUsed = (usages.length > 0) ? usages[0] : fallbackFill;

    return {
        odd: mostUsed,
        even: (usages.length > 1) ? usages[1] : mostUsed
    };
}

// =========================================
// ダイアログの組み立て / Dialog
// =========================================

/**
 * @typedef {object} RowFillControls
 * @property {Panel} panel パネル本体
 * @property {DropDownList} colorDropdown カラーのドロップダウン
 * @property {Group} swatchChip 選択中のカラーを示す色見本
 * @property {EditText} tintInput 濃淡の入力欄
 * @property {Slider} tintSlider 濃淡のスライダー
 * @property {Array<SwatchChoice>} swatchChoices ドロップダウンと同じ並びのスウォッチ一覧
 */

/**
 * 行のカラーと濃淡を設定するパネルを作る（奇数行・偶数行で共通）
 * @param {Group} parent 追加先のグループ
 * @param {string} titleKey パネル名のラベルキー
 * @param {string} tooltipKey パネルに付けるツールチップのキー
 * @param {Array<SwatchChoice>} swatchChoices スウォッチ一覧
 * @param {{colorName: string, tint: number}} defaultFill カラーと濃淡の初期値
 * @param {function} onChange 値が変わったときに呼ぶ処理
 * @returns {RowFillControls} 作成したコントロール一式
 */
function createRowFillPanel(parent, titleKey, tooltipKey, swatchChoices, defaultFill, onChange) {
    var rowFillPanel = parent.add("panel", undefined, getLabel(titleKey));
    setupPanel(rowFillPanel, ROW_PANEL_SPACING);
    rowFillPanel.alignChildren = ["left", "top"];
    rowFillPanel.helpTip = getLabel(tooltipKey);

    var colorGroup = rowFillPanel.add("group");
    colorGroup.orientation = "column";
    colorGroup.alignChildren = ["left", "top"];

    var colorLabelRow = colorGroup.add("group");
    setupRow(colorLabelRow, "left", SWATCH_ROW_SPACING);
    colorLabelRow.add("statictext", undefined, labelText("fieldLabel.color"));

    var swatchChip = colorLabelRow.add("group");
    swatchChip.preferredSize = [SWATCH_CHIP_SIZE, SWATCH_CHIP_SIZE];
    swatchChip.minimumSize   = [SWATCH_CHIP_SIZE, SWATCH_CHIP_SIZE];
    swatchChip.maximumSize   = [SWATCH_CHIP_SIZE, SWATCH_CHIP_SIZE];

    var displayNames = [];
    for (var i = 0; i < swatchChoices.length; i++) {
        displayNames.push(swatchChoices[i].displayName);
    }
    var colorDropdown = colorGroup.add("dropdownlist", undefined, displayNames);
    colorDropdown.preferredSize.width = COLOR_DROPDOWN_WIDTH;
    colorDropdown.helpTip = getLabel("tooltip.colorDropdown");
    colorDropdown.selection = findSwatchChoiceIndex(swatchChoices, defaultFill.colorName);

    var tintGroup = rowFillPanel.add("group");
    tintGroup.orientation = "column";
    tintGroup.alignChildren = ["left", "top"];
    tintGroup.margins = TINT_GROUP_MARGINS;

    var tintInputRow = tintGroup.add("group");
    setupRow(tintInputRow, "left", CONTROL_SPACING);
    tintInputRow.add("statictext", undefined, labelText("fieldLabel.tint"));

    /* 増減後は手入力と同じ onChange（スライダーの同期とプレビュー）を通す / after stepping, run the same onChange as typing */
    var tintInput;
    var tintStepperGroup = addStepperInputPair(tintInputRow, function () { return tintInput; }, {
        step: 1,
        min: TINT_MIN,
        max: TINT_MAX,
        integer: true,
        onStep: runInputOnChange
    });
    tintInput = tintStepperGroup.parent.add("edittext", undefined, String(defaultFill.tint));
    tintInput.preferredSize.width = TINT_INPUT_WIDTH;
    tintInput.helpTip = getLabel("tooltip.tintInput");
    tintInput.stepperGroup = tintStepperGroup;
    bindSteppedArrowKeys(tintInput, tintStepperGroup);

    var tintSlider = tintGroup.add("slider", undefined, defaultFill.tint, TINT_MIN, TINT_MAX);
    tintSlider.preferredSize.width = TINT_SLIDER_WIDTH;
    tintSlider.helpTip = getLabel("tooltip.tintSlider");

    var rowFill = {
        panel: rowFillPanel,
        colorDropdown: colorDropdown,
        swatchChip: swatchChip,
        tintInput: tintInput,
        tintSlider: tintSlider,
        swatchChoices: swatchChoices
    };

    /* ハンドラーは初期値を入れ終えてから付ける（selection への代入で発火するため）
       / Attach the handlers after the initial values are in place */
    colorDropdown.onChange = function () {
        refreshRowFill(rowFill);
        onChange();
    };
    tintSlider.onChanging = function () {
        var tint = snapTintToSliderStep(this.value);
        this.value = tint;
        tintInput.text = String(tint);
        onChange();
    };
    tintInput.onChange = function () {
        setRowFillTint(rowFill, normalizeTint(parseFloat(this.text)));
        onChange();
    };

    return rowFill;
}

/**
 * パネルで選択中のスウォッチを取得する
 * @param {RowFillControls} rowFill 対象のパネル
 * @returns {SwatchChoice} 選択中のスウォッチ
 */
function getRowFillChoice(rowFill) {
    var selectedIndex = rowFill.colorDropdown.selection ? rowFill.colorDropdown.selection.index : 0;
    return rowFill.swatchChoices[selectedIndex];
}

/**
 * パネルの濃淡を入力欄とスライダーの両方に反映する
 * @param {RowFillControls} rowFill 対象のパネル
 * @param {number} tint 設定する濃淡
 * @returns {void}
 */
function setRowFillTint(rowFill, tint) {
    rowFill.tintInput.text = String(tint);
    rowFill.tintSlider.value = tint;
}

/**
 * 選択中のカラーに合わせて、色見本と濃淡コントロールの状態を更新する
 * @param {RowFillControls} rowFill 対象のパネル
 * @returns {void}
 */
function refreshRowFill(rowFill) {
    var swatchChoice = getRowFillChoice(rowFill);
    var tintEnabled = !isTintDisabledKind(swatchChoice.swatchKind);

    setStepperInputEnabled(rowFill.tintInput, tintEnabled);
    rowFill.tintSlider.enabled = tintEnabled;
    paintSwatchChip(rowFill.swatchChip, findSwatch(swatchChoice.swatchName));
}

/**
 * オプションのパネルとスキップのパネルを追加する
 * @param {Window} dialog 対象のダイアログ
 * @param {object} ui UI オブジェクト
 * @param {function} onChange 値が変わったときに呼ぶ処理
 * @returns {void}
 */
function addOptions(dialog, ui, onChange) {
    /* 左にオプション、右にスキップの2カラム / Two columns: options on the left, skip on the right */
    var optionPanelsGroup = dialog.add("group");
    setupRow(optionPanelsGroup, "fill", COLUMN_SPACING);
    optionPanelsGroup.alignChildren = ["left", "top"];

    var optionsPanel = optionPanelsGroup.add("panel", undefined, getLabel("panel.options"));
    setupPanel(optionsPanel, ROW_PANEL_SPACING);
    optionsPanel.alignChildren = ["left", "top"];

    ui.swapCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.swapRows"));
    ui.swapCheckbox.helpTip = getLabel("tooltip.swapRows");
    ui.swapCheckbox.value = false;
    ui.swapCheckbox.onClick = onChange;

    /* 選択範囲の上から／左から、指定数を塗りの対象外にする / Leave the given rows and columns out of the fill */
    var skipPanel = optionPanelsGroup.add("panel", undefined, getLabel("panel.skip"));
    setupPanel(skipPanel, ROW_PANEL_SPACING);
    skipPanel.alignChildren = ["left", "top"];

    var skipRowControls = addSkipRow(skipPanel, "checkbox.skipRows", "tooltip.skipRows", onChange);
    ui.skipRowCheckbox = skipRowControls.checkbox;
    ui.skipRowInput = skipRowControls.input;

    var skipColumnControls = addSkipRow(skipPanel, "checkbox.skipColumns", "tooltip.skipColumns", onChange);
    ui.skipColumnCheckbox = skipColumnControls.checkbox;
    ui.skipColumnInput = skipColumnControls.input;
}

/**
 * スキップ数の1行（チェックボックス＋入力欄）を追加する
 * @param {Panel} parent 追加先のパネル
 * @param {string} labelKey チェックボックスのラベルキー
 * @param {string} tooltipKey ツールチップのキー
 * @param {function} onChange 値が変わったときに呼ぶ処理
 * @returns {{checkbox: Checkbox, input: EditText}} 作成したコントロール
 */
function addSkipRow(parent, labelKey, tooltipKey, onChange) {
    var skipGroup = parent.add("group");
    setupRow(skipGroup, "left", CONTROL_SPACING);

    var skipCheckbox = skipGroup.add("checkbox", undefined, labelText(labelKey));
    skipCheckbox.preferredSize.width = SKIP_CHECKBOX_WIDTH;
    /* 増減後は手入力と同じ onChange（整数化とプレビュー）を通す / after stepping, run the same onChange as typing */
    var skipInput;
    var skipStepperGroup = addStepperInputPair(skipGroup, function () { return skipInput; }, {
        step: 1,
        min: 0,
        integer: true,
        onStep: runInputOnChange
    });
    skipInput = skipStepperGroup.parent.add("edittext", undefined, String(SKIP_COUNT_DEFAULT));
    skipInput.preferredSize.width = SKIP_INPUT_WIDTH;
    skipInput.stepperGroup = skipStepperGroup;
    bindSteppedArrowKeys(skipInput, skipStepperGroup);

    skipCheckbox.helpTip = getLabel(tooltipKey);
    skipInput.helpTip = getLabel(tooltipKey);
    setStepperInputEnabled(skipInput, skipCheckbox.value);

    skipCheckbox.onClick = function () {
        setStepperInputEnabled(skipInput, skipCheckbox.value);
        onChange();
    };
    skipInput.onChange = function () {
        this.text = String(normalizeSkipCount(parseInt(this.text, 10)));
        onChange();
    };

    return { checkbox: skipCheckbox, input: skipInput };
}

/**
 * ダイアログ下部のボタン列を追加する（左：画面モード／右：キャンセル・OK）
 * @param {Window} dialog 対象のダイアログ
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function addDialogButtons(dialog, ui) {
    var buttonRow = addButtonRow(dialog);
    ui.btnScreenMode = addScreenModeButton(buttonRow.leftGroup);
    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    centerButtonRowIfRightOnly(buttonRow);
}

/**
 * ダイアログを組み立てる
 * @param {Array<SwatchChoice>} swatchChoices スウォッチ一覧
 * @param {{odd: object, even: object}} defaultFills 奇数行・偶数行の初期値
 * @param {function} onChange 値が変わったときに呼ぶ処理
 * @returns {object} ダイアログとコントロールをまとめた UI オブジェクト
 */
function buildDialog(swatchChoices, defaultFills, onChange) {
    var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    dialog.opacity = DIALOG_OPACITY;
    setupWindow(dialog);

    /* 左に奇数行、右に偶数行の2カラム / Two columns: odd rows on the left, even rows on the right */
    var rowPanelsGroup = dialog.add("group");
    setupRow(rowPanelsGroup, "fill", COLUMN_SPACING);
    rowPanelsGroup.alignChildren = ["left", "top"];

    var ui = { dialog: dialog };
    ui.oddRowFill = createRowFillPanel(rowPanelsGroup, "panel.oddRows", "tooltip.oddRows",
        swatchChoices, defaultFills.odd, onChange);
    ui.evenRowFill = createRowFillPanel(rowPanelsGroup, "panel.evenRows", "tooltip.evenRows",
        swatchChoices, defaultFills.even, onChange);

    addOptions(dialog, ui, onChange);
    addDialogButtons(dialog, ui);

    refreshRowFill(ui.oddRowFill);
    refreshRowFill(ui.evenRowFill);
    return ui;
}

// =========================================
// UIの値の読み取り / Read UI values
// =========================================

/**
 * @typedef {object} ZebraSettings
 * @property {CellFill} oddFill 奇数行の塗り
 * @property {CellFill} evenFill 偶数行の塗り
 * @property {boolean} swapRows 奇数行と偶数行を入れ替えるか
 * @property {number} skipRowCount 上からスキップする行数
 * @property {number} skipColumnCount 左からスキップする列数
 */

/**
 * パネルの設定を、セルに適用する形の塗りとして読み取る
 * @param {RowFillControls} rowFill 対象のパネル
 * @returns {CellFill} 読み取った塗り
 */
function readRowFill(rowFill) {
    var swatchChoice = getRowFillChoice(rowFill);
    return {
        color: findSwatch(swatchChoice.swatchName),
        tint: isTintDisabledKind(swatchChoice.swatchKind) ? null : normalizeTint(parseFloat(rowFill.tintInput.text))
    };
}

/**
 * スキップ数を読み取る（チェックが外れていれば 0）
 * @param {Checkbox} skipCheckbox スキップのチェックボックス
 * @param {EditText} skipInput スキップ数の入力欄
 * @returns {number} スキップする数
 */
function readSkipCount(skipCheckbox, skipInput) {
    return skipCheckbox.value ? normalizeSkipCount(parseInt(skipInput.text, 10)) : 0;
}

/**
 * ダイアログの設定をまとめて読み取る
 * @param {object} ui UI オブジェクト
 * @returns {ZebraSettings} 読み取った設定
 */
function readZebraSettings(ui) {
    return {
        oddFill: readRowFill(ui.oddRowFill),
        evenFill: readRowFill(ui.evenRowFill),
        swapRows: ui.swapCheckbox.value,
        skipRowCount: readSkipCount(ui.skipRowCheckbox, ui.skipRowInput),
        skipColumnCount: readSkipCount(ui.skipColumnCheckbox, ui.skipColumnInput)
    };
}

// =========================================
// 塗りの適用 / Apply fills
// =========================================

/**
 * @typedef {object} ZebraContext
 * @property {Array<Cell>} targetCells 対象のセル
 * @property {Array<CellFill>} originalFills 実行前の塗り（セルと同じ並び）
 * @property {boolean} previewApplied プレビューを当てたままかどうか
 */

/**
 * 選択セルが使っている行または列の位置を、重複なく昇順で集める
 * @param {Array<Cell>} cells 対象のセル
 * @param {string} parentName "parentRow" または "parentColumn"
 * @returns {Array<number>} 昇順に並べた位置
 */
function collectSortedIndices(cells, parentName) {
    var sortedIndices = [];
    var seen = {};
    for (var i = 0; i < cells.length; i++) {
        var parentIndex = cells[i][parentName].index;
        if (seen[parentIndex]) continue;
        seen[parentIndex] = true;
        sortedIndices.push(parentIndex);
    }
    sortedIndices.sort(function (a, b) { return a - b; });
    return sortedIndices;
}

/**
 * 先頭から指定数ぶんを「スキップする位置」として控える
 * @param {Array<number>} sortedIndices 昇順に並べた位置
 * @param {number} skipCount スキップする数
 * @returns {object} スキップする位置をキーに持つオブジェクト
 */
function buildSkipLookup(sortedIndices, skipCount) {
    var skipped = {};
    for (var i = 0; i < skipCount && i < sortedIndices.length; i++) {
        skipped[sortedIndices[i]] = true;
    }
    return skipped;
}

/**
 * スキップしていない行に、選択範囲の上から 0, 1, 2, … の順番を振る
 * @param {Array<number>} rowIndices 昇順に並べた行の位置
 * @param {object} skippedRows スキップする行
 * @returns {object} 行の位置をキー、順番を値に持つオブジェクト
 */
function buildRowOrderMap(rowIndices, skippedRows) {
    var rowOrderByIndex = {};
    var rowOrder = 0;
    for (var i = 0; i < rowIndices.length; i++) {
        if (skippedRows[rowIndices[i]]) continue;
        rowOrderByIndex[rowIndices[i]] = rowOrder++;
    }
    return rowOrderByIndex;
}

/**
 * セルに塗りを設定する
 * @param {Cell} cell 対象のセル
 * @param {CellFill} fill 設定する塗り。tint が null のときは濃淡を触らない
 * @returns {void}
 */
function applyCellFill(cell, fill) {
    if (!fill || !fill.color) return;
    cell.fillColor = fill.color;
    if (fill.tint === null) return;

    /* グラデーションなど濃淡を持たない塗りでは無視する / Fills without a tint, such as gradients, are left alone */
    try {
        cell.fillTint = fill.tint;
    } catch (e) { }
}

/**
 * 選択範囲内の行の並びを基準に、交互の塗りを適用する
 * @param {ZebraContext} context 対象のセルと実行前の塗り
 * @param {ZebraSettings} settings ダイアログの設定
 * @returns {void}
 */
function applyAlternatingFills(context, settings) {
    var targetCells = context.targetCells;
    if (targetCells.length === 0) return;

    var oddFill = settings.swapRows ? settings.evenFill : settings.oddFill;
    var evenFill = settings.swapRows ? settings.oddFill : settings.evenFill;

    var rowIndices = collectSortedIndices(targetCells, "parentRow");
    var columnIndices = collectSortedIndices(targetCells, "parentColumn");
    var skippedRows = buildSkipLookup(rowIndices, settings.skipRowCount);
    var skippedColumns = buildSkipLookup(columnIndices, settings.skipColumnCount);
    var rowOrderByIndex = buildRowOrderMap(rowIndices, skippedRows);

    for (var i = 0; i < targetCells.length; i++) {
        var cell = targetCells[i];
        var rowIndex = cell.parentRow.index;
        var columnIndex = cell.parentColumn.index;

        /* スキップした行・列は実行前の塗りに戻す / Skipped rows and columns go back to their previous fill */
        if (skippedRows[rowIndex] || skippedColumns[columnIndex]) {
            applyCellFill(cell, context.originalFills[i]);
            continue;
        }

        var rowOrder = rowOrderByIndex[rowIndex];
        if (rowOrder === undefined) continue;
        applyCellFill(cell, (rowOrder % 2 === 0) ? oddFill : evenFill);
    }
}

// =========================================
// プレビューと確定 / Preview & apply
// =========================================

/**
 * プレビューとして適用した塗りを取り消す
 * @param {ZebraContext} context 対象のセルと実行前の塗り
 * @returns {void}
 */
function clearPreview(context) {
    if (!context.previewApplied) return;
    context.previewApplied = false;

    /* プレビュー1回分の取り消し。失敗しても続行する / Undo the single preview step; keep going if it fails */
    try {
        app.undo();
        app.activeDocument.recompose();
    } catch (e) { }
}

/**
 * 現在の設定で塗りのプレビューを描き直す
 * @param {ZebraContext} context 対象のセルと実行前の塗り
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function refreshPreview(context, ui) {
    clearPreview(context);
    var settings = readZebraSettings(ui);

    /* 取り消しを1ステップにまとめる / Keep the preview to a single undo step */
    try {
        app.doScript(function () {
            applyAlternatingFills(context, settings);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.previewFills"));
        context.previewApplied = true;
        app.activeDocument.recompose();
    } catch (e) {
        context.previewApplied = false;
    }
}

/**
 * 選択セルに交互の塗りを適用するダイアログを表示して実行する
 * @returns {void}
 */
function main() {
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var targetCells = getSelectedCells(app.selection[0]);
    if (targetCells.length === 0) {
        alert(getLabel("alert.noCellSelection"));
        return;
    }

    var swatchChoices = getSwatchChoices();
    if (swatchChoices.length === 0) {
        alert(getLabel("alert.noUsableSwatch"));
        return;
    }

    var context = {
        targetCells: targetCells,
        originalFills: captureCellFills(targetCells),
        previewApplied: false
    };
    var defaultFills = detectDefaultFills(targetCells, swatchChoices[0].swatchName);

    /* セルの青い選択ハイライトがプレビューの邪魔になるため、選択は一旦解除する
       / The blue selection highlight hides the preview, so clear the selection */
    app.selection = null;

    var ui = null;
    ui = buildDialog(swatchChoices, defaultFills, function () {
        refreshPreview(context, ui);
    });

    refreshPreview(context, ui);

    if (ui.dialog.show() === 1) {
        clearPreview(context);
        var settings = readZebraSettings(ui);
        app.doScript(function () {
            applyAlternatingFills(context, settings);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.applyFills"));
    } else {
        clearPreview(context);
    }
}

main();

})();
