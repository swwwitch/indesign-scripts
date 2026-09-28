#target indesign

/*

### 概要

段落スタイルの「同一スタイル間の段落間隔」を、スタイル定義そのものに対して設定します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSetSameParaStyleSpacing.md

### Overview

Sets "space between paragraphs using same style" on the paragraph style definition itself.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSetSameParaStyleSpacing.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdSetSameParaStyleSpacing";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSetSameParaStyleSpacing.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSetSameParaStyleSpacing.md"; /* README (English) */

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

// ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
// ステップボタン（再利用パーツ） / Stepper buttons (reusable)
//
// 【移植手順 / How to port】
// 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
//    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
// 2. コピー先の LABELS.tooltip に stepUp / stepDown / stepUpInteger / stepDownInteger を足す（このファイルの LABELS から写す）。
//    getLabel() と uiLang はコピー先のものをそのまま使う
// 3. 数値欄を addSteppedField() で作る。項目名・∧∨・入力欄がひと組で入り、↑↓キーも∧∨と同じ処理で増減する
//      var widthInput = addSteppedField(parentPanel, {
//          label: labelText(LABELS.fieldLabel.width), labelWidth: 60,
//          text: "210 mm", characters: 8, step: 1, min: 1, unit: " mm",
//          onStep: function (numberInput) { updatePreview(); }
//      });
//    値の種類は options で切り分ける:
//      小数あり（幅・位置など）   … 指定なし（option＋クリックで0.1ずつ）
//      整数・1以上（段数・個数など）… integer: true, min: 1（0・小数・負数は受け付けず、option＋クリックも1ずつ）
//      整数・0以上（間隔の数など）  … integer: true, min: 0
//      範囲つき（％など）           … min: 0, max: 100, unit: "%"
// 4. 有効／無効は setSteppedFieldEnabled(widthInput, isEnabled)（∧∨のディム表示も切り替わる）。
//    行・パネルなど親の enabled を切り替えたときは、そのあとで redrawSteppersIn(親) を呼んで∧∨を描き直す
//    （∧∨は親をたどって無効を判定し、無効の間はクリックも↑↓キーも効かない）
// 5. 値は parseFloat(widthInput.text) で読む（unit 付きの欄は「210 mm」の形で入っている）
// 6. この欄に別の↑↓キー処理を付けない（↑↓キーが二重に効く）
// 既存の edittext をそのまま使うときは、同じ行の group（spacing 0）に addStepper() → edittext の順で置き、
// bindSteppedArrowKeys(edittext, stepperGroup) を呼ぶ
// ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
/**
 * UIがダークテーマかどうかを判定する（Illustrator・InDesign の両方に対応）
 * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
 */
function isDarkStepperUI() {
    try {
        if (app.preferences && app.preferences.getRealPreference) {
            return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
        }
        return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
    } catch (e) {
        return false;
    }
}

var STEPPER_UI_DARK           = isDarkStepperUI();
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
    var upTooltip = stepOptions.integer ? "tooltip.stepUpInteger" : "tooltip.stepUp";
    var downTooltip = stepOptions.integer ? "tooltip.stepDownInteger" : "tooltip.stepDown";
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

// ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
// ステップボタン（再利用パーツ）ここまで / End of the reusable stepper
// ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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
    /* ステップボタン用 / for the stepper */
    tooltip: {
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
    /* ∧∨と入力欄は隙間0で突き合わせる。下限0 / butt the stepper against the field; minimum 0 */
    var valueStepperInputGroup = valueGroup.add("group");
    valueStepperInputGroup.orientation = "row";
    valueStepperInputGroup.alignChildren = ["left", "center"];
    valueStepperInputGroup.spacing = 0;
    valueStepperInputGroup.margins = 0;
    var valueStepper = addStepper(valueStepperInputGroup, function () { return valueInput; }, { step: 1, min: 0 });
    var valueInput = valueStepperInputGroup.add("edittext", undefined, "0");
    valueInput.characters = 6;
    bindSteppedArrowKeys(valueInput, valueStepper);
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
        /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn stepper to update the dimming */
        valueStepper.enabled = (mode === "value");
        redrawSteppersIn(valueStepper);
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
