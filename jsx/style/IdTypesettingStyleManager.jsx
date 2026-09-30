#target indesign

/*

### 概要

段落スタイルの文字組版設定（禁則・文字組み・グリッド揃え・ハイフネーションなど）をダイアログでまとめて設定します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTypesettingStyleManager.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n7f67e8da571f

### Overview

Configures the composition settings of paragraph styles (kinsoku, mojikumi, grid alignment, hyphenation and more) from a single dialog.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTypesettingStyleManager.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTypesettingStyleManager";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTypesettingStyleManager.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTypesettingStyleManager.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n7f67e8da571f"; /* 紹介記事 / article URL */

// Original idea
// 欧文組版でのハイフネーション設定：コンさん
// https://typesetterkon.blogspot.com/2011/06/indesign5.html

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php
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
        title:            { ja: "文字組版設定（一括）", en: "Typesetting Settings (Batch)" },
        stylePicker:      { ja: "段落スタイルの選択", en: "Select Paragraph Styles" },
        presetNameInput:  { ja: "プリセット名の入力", en: "Enter Preset Name" }
    },
    panel: {
        targetStyles:     { ja: "対象の段落スタイル", en: "Target Paragraph Styles" },
        preset:           { ja: "プリセット", en: "Preset" },
        basicSettings:    { ja: "基本設定", en: "Basic Settings" },
        quotes:           { ja: "引用符（環境設定）", en: "Quotes (Preferences)" },
        units:            { ja: "単位（環境設定）", en: "Units (Preferences)" },
        japaneseTypeset:  { ja: "日本語文字組版", en: "Japanese Typesetting" },
        hyphenation:      { ja: "ハイフネーション", en: "Hyphenation" },
        hyphenateBreak:   { ja: "ハイフンで区切る", en: "Hyphenate" }
    },
    field: {
        autoKerning:      { ja: "自動カーニング", en: "Kerning" },
        autoLeading:      { ja: "自動行送り", en: "Auto Leading" },
        characterAlign:   { ja: "文字揃え", en: "Character Alignment" },
        leadingModel:     { ja: "行送りの基準位置", en: "Leading Model" },
        gridAlignment:    { ja: "グリッド揃え", en: "Align to Grid" },
        composer:         { ja: "コンポーザー", en: "Composer" },
        doubleQuote:      { ja: "二重引用符", en: "Double Quotes" },
        singleQuote:      { ja: "引用符", en: "Quotes" },
        textSize:         { ja: "テキストサイズ", en: "Text Size" },
        typography:       { ja: "組版", en: "Typography" },
        language:         { ja: "言語", en: "Language" },
        kinsokuSet:       { ja: "禁則処理セット", en: "Kinsoku Set" },
        kinsokuType:      { ja: "禁則調整方式", en: "Kinsoku Adjustment" },
        kinsokuHangType:  { ja: "ぶら下がり方法", en: "Hanging Punctuation" },
        mojikumi:         { ja: "文字組みアキ量", en: "Mojikumi" },
        minWordLength:    { ja: "単語の最初文字数", en: "Shortest Word" },
        afterFirst:       { ja: "先頭の後", en: "After First" },
        beforeLast:       { ja: "最後の前", en: "Before Last" },
        maxHyphens:       { ja: "最大のハイフン数", en: "Hyphen Limit" },
        hyphenationZone:  { ja: "領域", en: "Hyphenation Zone" },
        presetName:       { ja: "プリセット名（書き出しファイル名にも使用）", en: "Preset name (also used as the export filename)" },
        targetCount:      { ja: "対象", en: "Targets" },
        fileName:         { ja: "ファイル名", en: "File name" }
    },
    radio: {
        targetSelection:  { ja: "選択中", en: "Selection" },
        targetAll:        { ja: "すべて", en: "All" },
        targetSpecified:  { ja: "指定", en: "Specified" },
        languageJapanese: { ja: "日本語", en: "Japanese" },
        languageEnglish:  { ja: "英語", en: "English" },
        languageNone:     { ja: "なし", en: "None" }
    },
    checkbox: {
        typographersQuotes: { ja: "英文引用符を使用", en: "Use Typographer's Quotes" },
        noBreak:            { ja: "分離禁止処理", en: "No Break" },
        digitsRotation:     { ja: "連数字処理", en: "Tatechuyoko" },
        rotateInVertical:   { ja: "縦組み中の文字回転", en: "Rotate Characters in Vertical Text" },
        absorbTrailingSpace:{ ja: "全角スペースを行末吸収", en: "Absorb Trailing Full-width Space" },
        arbitraryHyphen:    { ja: "欧文泣き別れ", en: "Allow Arbitrary Hyphenation" },
        capitalizedWords:   { ja: "大文字の単語", en: "Capitalized Words" },
        acrossColumns:      { ja: "段間、フレームにわたる単語", en: "Words Across Columns and Frames" },
        lastWord:           { ja: "段落末尾の単語", en: "Last Word" },
        ligatures:          { ja: "欧文合字", en: "Ligatures" },
        hyphenation:        { ja: "ハイフネーション", en: "Hyphenation" }
    },
    button: {
        ok:        { ja: "OK", en: "OK" },
        cancel:    { ja: "キャンセル", en: "Cancel" },
        select:    { ja: "選択", en: "Select" },
        selectAll: { ja: "全選択", en: "Select All" },
        clearAll:  { ja: "全解除", en: "Clear All" },
        export:    { ja: "書き出し", en: "Export" }
    },
    unit: {
        item:      { ja: " 件", en: " items" },
        character: { ja: "文字", en: "characters" },
        hyphen:    { ja: "ハイフン", en: "hyphens" }
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
    hint: {
        multiSelect: { ja: "Shift / Cmd（Ctrl）+ クリックで複数選択", en: "Shift / Cmd (Ctrl) + click to select multiple" }
    },
    export: {
        codeHeader:      { ja: "// PRESETS マップに以下を追加してください（プリセットドロップダウン項目への追加もお忘れなく）", en: "// Add the following to the PRESETS map (and to the preset dropdown items as well)" },
        overwritePrefix: { ja: "「", en: "\"" },
        overwriteSuffix: { ja: ".jsx」は既にデスクトップに存在します。上書きしますか？", en: ".jsx\" already exists on the Desktop. Overwrite it?" },
        savedPrefix:     { ja: "プリセット「", en: "Preset \"" },
        savedSuffix:     { ja: "」をデスクトップに書き出しました。", en: "\" was exported to the Desktop." }
    },
    alert: {
        openFileFailed:       { ja: "ファイルを開けませんでした。", en: "The file could not be opened." },
        exportErrorPrefix:    { ja: "書き出しエラー: ", en: "Export error: " },
        partialFailurePrefix: { ja: "適用しましたが、", en: "Applied, but " },
        partialFailureSuffix: { ja: " 件の段落スタイルでエラーが発生しました。\n\n", en: " paragraph style(s) reported an error.\n\n" },
        noDocument:           { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document before running." },
        noKinsokuTables:      { ja: "このドキュメントには禁則処理セットがありません。", en: "This document has no kinsoku tables." },
        noParagraphStyles:    { ja: "適用可能な段落スタイルがありません。", en: "There are no applicable paragraph styles." },
        noTargetStyles:       { ja: "適用対象の段落スタイルが見つかりません。選択範囲、または指定した段落スタイルを確認してください。", en: "No target paragraph style was found. Check the selection or the specified styles." }
    },
    undo: {
        applyTypesetting: { ja: "段落スタイルに組版設定を適用", en: "Apply Typesetting Settings to Paragraph Styles" }
    }
};

// =========================================
// ユーザー設定 / User settings
// =========================================

// 言語の候補（先頭から順に試行） / Language candidates (tried in order)
var LANGUAGE_CANDIDATES = {
    "ja": ["日本語", "Japanese"],
    "en": ["英語：米国", "English: USA"],
    "none": ["[言語なし]", "[No Language]"]
};

// 二重引用符の候補 / Double quote options
var DOUBLE_QUOTE_OPTIONS = [
    "“”",
    "«»",
    "„“",
    "『』",
    "「」",
    "\"\""
];

// 引用符の候補 / Single quote options
var SINGLE_QUOTE_OPTIONS = [
    "‘’",
    "‹›",
    "‚‘",
    "〈〉",
    "''"
];

/* 対象から除外する段落スタイルグループの接頭辞 / Prefix marking style groups to skip */
var EXCLUDED_STYLE_GROUP_PREFIX = "_";

// ドロップダウン幅 / Dropdown widths
var W_DROP = 130;

// =========================================
// 文字組みアキ量プリセット定義 / Mojikumi preset definitions
// =========================================

/* 文字組み「なし」の表示名（プリセットや読み取り結果と名前で照合する） / Display name of "no mojikumi" (matched by name against presets and read values) */
var MOJIKUMI_NONE_NAME = "なし";

/*
組み込みの文字組みアキ量 preset enum と表示名の対応表。
InDesign の enum 名から UI 表示用ラベルへ変換するために使用。

Mapping table between built-in mojikumi preset enums and display labels.
Used to convert internal InDesign enum names into user-facing labels.
*/
var MOJIKUMI_LABELS = {
    "LINE_END_ALL_ONE_HALF_EM_ENUM": "行末約物半角",
    "ONE_EM_INDENT_LINE_END_UKE_ONE_HALF_EM_ENUM": "行末受け約物半角・段落1字下げ（起こし全角）",
    "ONE_OR_ONE_HALF_EM_INDENT_LINE_END_UKE_ONE_HALF_EM_ENUM": "行末受け約物半角・段落1字下げ（起こし食い込み）",
    "ONE_OR_ONE_HALF_EM_INDENT_LINE_END_ALL_ONE_EM_ENUM": "約物全角・段落1字下げ",
    "ONE_EM_INDENT_LINE_END_ALL_ONE_EM_ENUM": "約物全角・段落1字下げ（起こし全角）",
    "ONE_EM_INDENT_LINE_END_ALL_NO_FLOAT_ENUM": "行末約物全角/半角・段落1字下げ",
    "ONE_EM_INDENT_LINE_END_UKE_NO_FLOAT_ENUM": "行末受け約物全角／半角・段落1字下げ（起こし全角）",
    "ONE_OR_ONE_HALF_EM_INDENT_LINE_END_UKE_NO_FLOAT_ENUM": "行末受け約物全角／半角・段落1字下げ（起こし食い込み）",
    "ONE_EM_INDENT_LINE_END_ALL_ONE_HALF_EM_ENUM": "行末約物半角・段落1字下げ",
    "LINE_END_ALL_ONE_EM_ENUM": "約物全角",
    "LINE_END_UKE_NO_FLOAT_ENUM": "行末受け約物全角／半角",
    "ONE_OR_ONE_HALF_EM_INDENT_LINE_END_PERIOD_ONE_EM_ENUM": "行末句点全角・段落1字下げ",
    "ONE_EM_INDENT_LINE_END_PERIOD_ONE_EM_ENUM": "行末句点全角・段落1字下げ（起こし全角）",
    "LINE_END_PERIOD_ONE_EM_ENUM": "行末句点全角",
    "TRAD_CHINESE_DEFAULT": "繁体字中国語デフォルト",
    "SIMP_CHINESE_DEFAULT": "簡体字中国語デフォルト"
};

// =========================================
// 配列の検索 / Array lookup
// =========================================

/**
 * 配列の中で値が一致する（===）最初の位置を探す
 * @param {Array} items 探す配列（表示名・列挙値・オブジェクト参照など）
 * @param {*} target 探す値
 * @returns {number} 見つかった位置。なければ -1
 */
function findIndexInArray(items, target) {
    for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
        if (items[itemIndex] === target) return itemIndex;
    }
    return -1;
}

/**
 * 既定値の名前から選択位置を求める（見つからなければ先頭）
 * @param {Array<string>} names 表示名の一覧
 * @param {string} defaultName 既定値の名前
 * @returns {number} 選択する位置
 */
function getDefaultIndexByName(names, defaultName) {
    var foundIndex = findIndexInArray(names, defaultName);
    return foundIndex >= 0 ? foundIndex : 0;
}

// =========================================
// ドキュメント情報の取得 / Document data collection
// =========================================

/**
 * ドキュメント内の禁則処理セットを集める
 * @param {Document} documentObject 対象ドキュメント
 * @returns {{tables: Array, names: Array<string>}} 禁則処理セットと表示名
 */
function collectKinsokuTables(documentObject) {
    var tables = [];
    var names = [];
    for (var kinsokuTableIndex = 0; kinsokuTableIndex < documentObject.kinsokuTables.length; kinsokuTableIndex++) {
        var kinsokuTable = documentObject.kinsokuTables.item(kinsokuTableIndex);
        tables.push(kinsokuTable);
        names.push(kinsokuTable.name);
    }
    return { tables: tables, names: names };
}

/**
 * 禁則調整方式の選択肢を作る
 * @returns {{values: Array, names: Array<string>}} 調整方式と表示名
 */
function createKinsokuTypeOptions() {
    return {
        names: ["追い込み優先", "追い出し優先", "追い出しのみ", "調整量を優先"],
        values: [
            KinsokuType.KINSOKU_PUSH_IN_FIRST,
            KinsokuType.KINSOKU_PUSH_OUT_FIRST,
            KinsokuType.KINSOKU_PUSH_OUT_ONLY,
            KinsokuType.KINSOKU_PRIORITIZE_ADJUSTMENT_AMOUNT
        ]
    };
}

/**
 * ドキュメント内の文字組みアキ量設定を集める（先頭は「なし」）
 * @param {Document} documentObject 対象ドキュメント
 * @returns {{tables: Array, names: Array<string>}} 文字組み設定と表示名（「なし」の設定は null）
 */
function collectMojikumiTables(documentObject) {
    var tables = [null];
    var names = [MOJIKUMI_NONE_NAME];
    for (var mojikumiTableIndex = 0; mojikumiTableIndex < documentObject.mojikumiTables.length; mojikumiTableIndex++) {
        var mojikumiTable = documentObject.mojikumiTables.item(mojikumiTableIndex);
        tables.push(mojikumiTable);
        names.push(mojikumiTable.name);
    }
    return { tables: tables, names: names };
}

/**
 * 対象にする段落スタイルをグループ込みで集める
 * @param {Document} documentObject 対象ドキュメント
 * @returns {{styles: Array<ParagraphStyle>, names: Array<string>}} 段落スタイルと表示名
 */
function collectTargetParagraphStyles(documentObject) {
    var styles = [];
    var names = [];

    /**
     * スタイルグループを再帰的にたどって段落スタイルを集める
     * @param {object} container 段落スタイルのコンテナ
     * @param {string} prefix グループ名の接頭辞
     * @returns {void}
     */
    function walk(container, prefix) {
        for (var paragraphStyleIndex = 0; paragraphStyleIndex < container.paragraphStyles.length; paragraphStyleIndex++) {
            var paragraphStyle = container.paragraphStyles.item(paragraphStyleIndex);
            var paragraphStyleName = paragraphStyle.name;
            if (paragraphStyleName === "[段落スタイルなし]" || paragraphStyleName === "[No Paragraph Style]") continue;
            if (paragraphStyleName === "[基本段落]" || paragraphStyleName === "[Basic Paragraph]") continue;
            styles.push(paragraphStyle);
            names.push(prefix + paragraphStyleName);
        }
        for (var styleGroupIndex = 0; styleGroupIndex < container.paragraphStyleGroups.length; styleGroupIndex++) {
            var styleGroup = container.paragraphStyleGroups.item(styleGroupIndex);
            if (styleGroup.name.charAt(0) === EXCLUDED_STYLE_GROUP_PREFIX) continue;
            walk(styleGroup, prefix + styleGroup.name + " / ");
        }
    }

    walk(documentObject, "");
    return { styles: styles, names: names };
}

/**
 * コンポーザーの選択肢と適用用エイリアスを作る
 * @returns {{names: Array<string>, aliases: Array<Array<string>>}} 表示名と、ロケール・バージョン違いの名前の候補
 */
function createComposerOptions() {
    var entries = [
        { name: "日本語段落コンポーザー", aliases: ["Adobe 日本語段落コンポーザー", "Adobe Japanese Paragraph Composer"] },
        { name: "日本語単数行コンポーザー", aliases: ["Adobe 日本語単数行コンポーザー", "Adobe Japanese Single-line Composer"] },
        { name: "多言語対応段落コンポーザー", aliases: ["$ID/HL Composer Optyca", "Adobe World-Ready Paragraph Composer", "Adobe 多言語対応段落コンポーザー", "Adobe World-Ready 段落コンポーザー"] },
        { name: "多言語対応単数行コンポーザー", aliases: ["$ID/HL Single Optyca", "Adobe World-Ready Single-line Composer", "Adobe 多言語対応単数行コンポーザー", "Adobe World-Ready 単数行コンポーザー"] },
        { name: "欧文段落コンポーザー", aliases: ["$ID/HL Composer", "Adobe Paragraph Composer", "Adobe 欧文段落コンポーザー", "Adobe 段落コンポーザー"] },
        { name: "欧文単数行コンポーザー", aliases: ["$ID/HL Single", "Adobe Single-line Composer", "Adobe 欧文単数行コンポーザー", "Adobe 単数行コンポーザー"] }
    ];
    var names = [];
    var aliases = [];
    for (var entryIndex = 0; entryIndex < entries.length; entryIndex++) {
        names.push(entries[entryIndex].name);
        aliases.push(entries[entryIndex].aliases);
    }
    return { names: names, aliases: aliases };
}

// =========================================
// 設定値と選択肢の照合 / Matching setting values to options
// =========================================

/**
 * コンポーザー名（エイリアスのどれか）から選択位置を探す
 * @param {Array<Array<string>>} aliasesList createComposerOptions() の aliases
 * @param {string} target 探すコンポーザー名
 * @returns {number} 見つかった位置。なければ -1
 */
function findIndexByComposerAliases(aliasesList, target) {
    if (!target) return -1;
    for (var listIndex = 0; listIndex < aliasesList.length; listIndex++) {
        if (findIndexInArray(aliasesList[listIndex], target) >= 0) return listIndex;
    }
    return -1;
}

/**
 * エイリアスを順に試して段落スタイルにコンポーザーを設定する
 * @param {ParagraphStyle} targetParagraphStyle 対象の段落スタイル
 * @param {Array<string>} aliases コンポーザー名の候補
 * @returns {boolean} 設定できたら true
 */
function applyComposerAliases(targetParagraphStyle, aliases) {
    if (!aliases) return false;
    for (var aliasIndex = 0; aliasIndex < aliases.length; aliasIndex++) {
        try {
            targetParagraphStyle.composer = aliases[aliasIndex];
            return true;
        } catch (composerAliasError) { }
    }
    return false;
}

/**
 * 禁則処理セットの設定値から選択位置を探す（参照・ID・名前の順に照合）
 * @param {*} kinsokuValue 禁則処理セットの設定値
 * @param {Array} kinsokuTables 禁則処理セットの一覧
 * @param {Array<string>} kinsokuNames 禁則処理セットの表示名
 * @returns {number} 見つかった位置。なければ -1
 */
function findKinsokuIndexFromValue(kinsokuValue, kinsokuTables, kinsokuNames) {
    if (kinsokuValue === null || kinsokuValue === undefined) return -1;
    for (var refIndex = 0; refIndex < kinsokuTables.length; refIndex++) {
        if (kinsokuTables[refIndex] === kinsokuValue) return refIndex;
        try {
            if (kinsokuTables[refIndex].id !== undefined && kinsokuValue.id !== undefined && kinsokuTables[refIndex].id === kinsokuValue.id) return refIndex;
        } catch (eKinsokuIdCompare) { }
    }
    var kinsokuValueName = null;
    try { kinsokuValueName = kinsokuValue.name; } catch (eKinsokuValueName) { }
    if (typeof kinsokuValueName === "string" && kinsokuValueName.length > 0) {
        return findIndexInArray(kinsokuNames, kinsokuValueName);
    }
    return -1;
}

/**
 * 文字組みアキ量設定の表示名を求める
 * @param {*} mojikumiValue 文字組みの設定値（MojikumiTable・文字列・NothingEnum・組み込みプリセットの列挙値）
 * @returns {string} 表示名。求められなければ空文字
 */
function resolveMojikumiName(mojikumiValue) {
    if (mojikumiValue === null || mojikumiValue === undefined || mojikumiValue === NothingEnum.NOTHING) {
        return MOJIKUMI_NONE_NAME;
    }
    if (typeof mojikumiValue === "string") {
        return mojikumiValue;
    }

    /* MojikumiTable は .name を持つ。プリセットの列挙値は持たないため toString() で照合 / MojikumiTable has .name; preset enums are matched via toString() */
    try {
        if (mojikumiValue.isValid && typeof mojikumiValue.name === "string" && mojikumiValue.name.length > 0) {
            return mojikumiValue.name;
        }
    } catch (mojikumiNameError) { }

    var mojikumiKey = "";
    try { mojikumiKey = mojikumiValue.toString(); } catch (mojikumiStringError) { }
    for (var enumKey in MOJIKUMI_LABELS) {
        if (mojikumiKey.indexOf(enumKey) !== -1) {
            return MOJIKUMI_LABELS[enumKey];
        }
    }

    return "";
}

// =========================================
// 選択肢の定義 / Option definitions
// =========================================

/**
 * ぶら下がり方法の選択肢を作る
 * @returns {object} ぶら下がり方法と表示名
 */
function createKinsokuHangTypeOptions() {
    return {
        names: ["なし", "標準", "強制"],
        values: [
            KinsokuHangTypes.NONE,
            KinsokuHangTypes.KINSOKU_HANG_REGULAR,
            KinsokuHangTypes.KINSOKU_HANG_FORCE
        ]
    };
}

/**
 * 行送りの基準位置の選択肢を作る
 * @returns {object} 基準位置と表示名
 */
function createLeadingModelOptions() {
    return {
        names: ["仮想ボディの上/右", "仮想ボディの中央", "欧文ベースライン", "仮想ボディの下/左"],
        values: [
            LeadingModel.LEADING_MODEL_AKI_BELOW,
            LeadingModel.LEADING_MODEL_CENTER,
            LeadingModel.LEADING_MODEL_ROMAN,
            LeadingModel.LEADING_MODEL_AKI_ABOVE
        ]
    };
}

/**
 * 自動カーニング方式の選択肢を作る
 * @returns {object} カーニング方式と表示名
 */
function createKerningMethodOptions() {
    var values = ["メトリクス", "オプティカル", "和文等幅", "0"];
    return { names: values, values: values };
}

/**
 * グリッド揃えの選択肢を作る
 * @returns {object} グリッド揃えと表示名
 */
function createGridAlignmentOptions() {
    return {
        names: [
            "なし",
            "欧文ベースライン",
            "仮想ボディの上/右",
            "仮想ボディの中央",
            "仮想ボディの下/左",
            "平均字面の上/右",
            "平均字面の下/左"
        ],
        values: [
            GridAlignment.NONE,
            GridAlignment.ALIGN_BASELINE,
            GridAlignment.ALIGN_EM_TOP,
            GridAlignment.ALIGN_EM_CENTER,
            GridAlignment.ALIGN_EM_BOTTOM,
            GridAlignment.ALIGN_ICF_TOP,
            GridAlignment.ALIGN_ICF_BOTTOM
        ]
    };
}

/**
 * 文字揃えの選択肢を作る
 * @returns {object} 文字揃えと表示名
 */
function createCharacterAlignmentOptions() {
    return {
        names: [
            "欧文ベースライン",
            "仮想ボディの上/右",
            "仮想ボディの中央",
            "仮想ボディの下/左",
            "平均字面の上/右",
            "平均字面の下/左"
        ],
        values: [
            CharacterAlignment.ALIGN_BASELINE,
            CharacterAlignment.ALIGN_EM_TOP,
            CharacterAlignment.ALIGN_EM_CENTER,
            CharacterAlignment.ALIGN_EM_BOTTOM,
            CharacterAlignment.ALIGN_ICF_TOP,
            CharacterAlignment.ALIGN_ICF_BOTTOM
        ]
    };
}

// =========================================
// ダイアログ UI / Dialog UI
// =========================================

/**
 * ラベル付きドロップダウンの行を追加する
 * @param {object} parent 追加先のコンテナ
 * @param {string} labelString ラベルの文字列（コロン付き）
 * @param {Array<string>} items 選択肢
 * @param {number} selectionIndex 既定の選択位置
 * @returns {DropDownList} 追加したドロップダウン
 */
function addDropdownRow(parent, labelString, items, selectionIndex) {
    var row = parent.add("group");
    setupRow(row, "left", 8);

    var label = row.add("statictext", undefined, labelString);
    label.preferredSize.width = 120;

    var dropdown = row.add("dropdownlist", undefined, items);
    dropdown.selection = selectionIndex;
    dropdown.preferredSize.width = W_DROP;

    return dropdown;
}

/**
 * ラベル付き数値入力の行を追加する
 * @param {object} parent 追加先のコンテナ
 * @param {string} labelString ラベルの文字列（コロン付き）
 * @param {*} defaultValue 初期値
 * @param {string} suffixText 単位の文字列
 * @param {object} stepOptions ∧∨の増減の設定（addStepper() に渡す step / min / max / integer）
 * @returns {EditText} 追加した入力欄
 */
function addNumberRow(parent, labelString, defaultValue, suffixText, stepOptions) {
    var row = parent.add("group");
    setupRow(row, "left", 8);

    var label = row.add("statictext", undefined, labelString);
    label.preferredSize.width = 120;

    /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
    var stepperInputGroup = row.add("group");
    stepperInputGroup.orientation = "row";
    stepperInputGroup.alignChildren = ["left", "center"];
    stepperInputGroup.spacing = 0;
    stepperInputGroup.margins = 0;
    var stepperGroup = addStepper(stepperInputGroup, function () { return input; }, stepOptions);

    var input = stepperInputGroup.add("edittext", undefined, String(defaultValue));
    input.preferredSize.width = 50;
    input.justify = "right";
    bindSteppedArrowKeys(input, stepperGroup);
    input.fieldRow = row; /* 行ごと有効／無効を切り替えるため / for toggling the whole row */

    if (typeof suffixText === "string" && suffixText.length > 0) {
        row.add("statictext", undefined, suffixText);
    }

    return input;
}

/**
 * 段落スタイルを複数選択するダイアログを表示する
 * @param {Array<string>} paragraphStyleNames 段落スタイルの表示名
 * @param {Array<string>} selectedNames すでに選択済みの名前
 * @returns {Array<string>|null} 選択した名前。キャンセル時は null
 */
function showParagraphStylePicker(paragraphStyleNames, currentSelectedIndexes) {
    var picker = new Window("dialog", getLabel("dialog.stylePicker"));
    setupWindow(picker, 10);

    var targetCountText = picker.add("statictext", undefined, labelValueText("field.targetCount", paragraphStyleNames.length + getLabel("unit.item")));
    targetCountText.alignment = "left";

    // listbox はスタイル数が多くても自動でスクロールバーが付く /
    // listbox shows a scrollbar automatically when items overflow
    var styleListbox = picker.add("listbox", undefined, paragraphStyleNames, { multiselect: true });
    styleListbox.preferredSize = [360, 320];

    /**
     * その段落スタイルが選択済みかを判定する
     * @param {string} styleName 段落スタイル名
     * @returns {boolean} 選択済みなら true
     */
    function isCurrentlySelected(index) {
        if (!currentSelectedIndexes) return true;
        return findIndexInArray(currentSelectedIndexes, index) >= 0;
    }

    var initialSelectionIndexes = [];
    for (var nameIndex = 0; nameIndex < paragraphStyleNames.length; nameIndex++) {
        if (isCurrentlySelected(nameIndex)) initialSelectionIndexes.push(nameIndex);
    }
    if (initialSelectionIndexes.length > 0) styleListbox.selection = initialSelectionIndexes;

    var toolRow = picker.add("group");
    setupRow(toolRow, "left", 6);
    var btnSelectAll = toolRow.add("button", undefined, getLabel("button.selectAll"));
    var btnClearAll = toolRow.add("button", undefined, getLabel("button.clearAll"));

    btnSelectAll.onClick = function () {
        var allIndexes = [];
        for (var allIndex = 0; allIndex < styleListbox.items.length; allIndex++) allIndexes.push(allIndex);
        styleListbox.selection = allIndexes;
    };
    btnClearAll.onClick = function () {
        styleListbox.selection = null;
    };

    var hint = picker.add("statictext", undefined, getLabel("hint.multiSelect"));
    hint.alignment = "left";

    var buttonRow = addButtonRow(picker);
    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    alignRightOnlyButtonRow(buttonRow);

    if (picker.show() !== 1) return null;

    var selectedIndexes = [];
    if (styleListbox.selection) {
        for (var selectionItemIndex = 0; selectionItemIndex < styleListbox.selection.length; selectionItemIndex++) {
            selectedIndexes.push(styleListbox.selection[selectionItemIndex].index);
        }
    }
    return selectedIndexes;
}

/**
 * 列挙値からドロップダウンを安全に選択する
 * @param {DropDownList} dropdown 対象のドロップダウン
 * @param {Array} values 選択肢の列挙値
 * @param {*} target 設定する列挙値
 * @returns {boolean} 選択できたら true
 */
function safeAssignDropdownFromEnum(dropdown, values, target) {
    var foundIndex = findIndexInArray(values, target);
    if (foundIndex < 0) return false;
    dropdown.selection = foundIndex;
    return true;
}

/**
 * 名前からドロップダウンを安全に選択する
 * @param {DropDownList} dropdown 対象のドロップダウン
 * @param {Array<string>} names 表示名の一覧
 * @param {string} targetName 設定する名前
 * @returns {boolean} 選択できたら true
 */
function safeAssignDropdownFromName(dropdown, names, targetName) {
    if (!targetName) return false;
    var foundIndex = findIndexInArray(names, targetName);
    if (foundIndex < 0) return false;
    dropdown.selection = foundIndex;
    return true;
}

/**
 * チェックボックスへ値を安全に設定する
 * @param {Checkbox} checkbox 対象のチェックボックス
 * @param {*} value 設定する値
 * @returns {void}
 */
function safeAssignCheckbox(checkbox, value) {
    checkbox.value = !!value;
}

/**
 * 数値入力欄へ値を安全に設定する
 * @param {EditText} input 対象の入力欄
 * @param {*} value 設定する値
 * @returns {void}
 */
function safeAssignNumberInput(input, value) {
    if (typeof value === "number") input.text = String(value);
}

/**
 * 存在しない場合もあるプロパティを安全に設定する
 * @param {object} targetObject 設定先のオブジェクト
 * @param {string} propertyName プロパティ名
 * @param {*} value 設定する値
 * @returns {boolean} 設定できたら true
 */
function safeSetProperty(targetObject, propertyName, value) {
    try {
        targetObject[propertyName] = value;
        return true;
    } catch (propertySetError) { }
    return false;
}

/* プリセット定義（段落スタイル設定と環境設定を分離） / Preset definitions split into paragraph style settings and app preferences */

var PRESETS = {
    "欧文組版": {
        styleSettings: {
            kinsoku: "弱い禁則",
            kinsokuType: "調整量を優先",
            kinsokuHangType: "なし",
            bunriKinshi: true,
            mojikumi: "なし",
            leadingModel: "欧文ベースライン",
            rensuuji: true,
            rotateSingleByte: false,
            absorbLineEndIdeographicSpace: true,
            latinWordBreak: false,
            kerningMethod: "メトリクス",
            autoLeading: 120,
            characterAlignment: "欧文ベースライン",
            gridAlignment: "なし",
            composer: "欧文段落コンポーザー",
            hyphenation: true,
            hyphenateWordsLongerThan: 6,
            hyphenateAfterFirst: 3,
            hyphenateBeforeLast: 3,
            hyphenateLadderLimit: 2,
            hyphenationZone: 6,
            hyphenateCapitalizedWords: false,
            hyphenateAcrossColumns: false,
            hyphenateLastWord: false,
            ligatures: true,
            language: "en"
        },
        appPreferences: {
            useSmartQuotes: true,
            doubleQuotes: "“”",
            singleQuotes: "‘’",
            textSizeUnit: "ポイント",
            compositionUnit: "ポイント"
        }
    },
    "グリッド優先": {
        styleSettings: {
            kinsoku: "強い禁則",
            kinsokuType: "追い込み優先",
            kinsokuHangType: "なし",
            bunriKinshi: true,
            mojikumi: "行末約物半角",
            leadingModel: "仮想ボディの中央",
            rensuuji: true,
            rotateSingleByte: false,
            absorbLineEndIdeographicSpace: true,
            latinWordBreak: false,
            kerningMethod: "和文等幅",
            autoLeading: 100,
            characterAlignment: "仮想ボディの中央",
            gridAlignment: "仮想ボディの中央",
            composer: "日本語単数行コンポーザー",
            hyphenation: false,
            hyphenateWordsLongerThan: 6,
            hyphenateAfterFirst: 3,
            hyphenateBeforeLast: 3,
            hyphenateLadderLimit: 2,
            hyphenationZone: 10,
            hyphenateCapitalizedWords: false,
            hyphenateAcrossColumns: false,
            hyphenateLastWord: false,
            ligatures: true,
            language: "ja"
        },
        appPreferences: {
            useSmartQuotes: false,
            doubleQuotes: "“”",
            singleQuotes: "‘’",
            textSizeUnit: "級",
            compositionUnit: "歯"
        }
    },
    "グリッド無視": {
        styleSettings: {
            kinsoku: "弱い禁則",
            kinsokuType: "調整量を優先",
            kinsokuHangType: "なし",
            bunriKinshi: true,
            mojikumi: "行末約物半角",
            leadingModel: "仮想ボディの中央",
            rensuuji: true,
            rotateSingleByte: false,
            absorbLineEndIdeographicSpace: true,
            latinWordBreak: false,
            kerningMethod: "メトリクス",
            autoLeading: 175,
            characterAlignment: "仮想ボディの中央",
            gridAlignment: "なし",
            composer: "日本語段落コンポーザー",
            hyphenation: false,
            hyphenateWordsLongerThan: 6,
            hyphenateAfterFirst: 3,
            hyphenateBeforeLast: 3,
            hyphenateLadderLimit: 2,
            hyphenationZone: 10,
            hyphenateCapitalizedWords: false,
            hyphenateAcrossColumns: false,
            hyphenateLastWord: false,
            ligatures: true,
            language: "ja"
        },
        appPreferences: {
            useSmartQuotes: true,
            doubleQuotes: "“”",
            singleQuotes: "‘’",
            textSizeUnit: "ポイント",
            compositionUnit: "ポイント"
        }
    },
    "ソースコード": {
        styleSettings: {
            kinsoku: "弱い禁則",
            kinsokuType: "追い込み優先",
            kinsokuHangType: "なし",
            bunriKinshi: false,
            mojikumi: "なし",
            leadingModel: "欧文ベースライン",
            rensuuji: false,
            rotateSingleByte: false,
            absorbLineEndIdeographicSpace: false,
            latinWordBreak: false,
            kerningMethod: "0",
            autoLeading: 120,
            characterAlignment: "欧文ベースライン",
            gridAlignment: "なし",
            composer: "欧文単数行コンポーザー",
            hyphenation: false,
            hyphenateWordsLongerThan: 6,
            hyphenateAfterFirst: 3,
            hyphenateBeforeLast: 3,
            hyphenateLadderLimit: 2,
            hyphenationZone: 1.25,
            hyphenateCapitalizedWords: false,
            hyphenateAcrossColumns: false,
            hyphenateLastWord: false,
            ligatures: false,
            language: "none"
        },
        appPreferences: {
            useSmartQuotes: false,
            doubleQuotes: "\"\"",
            singleQuotes: "''",
            textSizeUnit: "ポイント",
            compositionUnit: "ポイント"
        }
    },
    "InDesignのデフォルト": {
        styleSettings: {
            kinsoku: "強い禁則",
            kinsokuType: "追い込み優先",
            kinsokuHangType: "なし",
            bunriKinshi: true,
            mojikumi: "行末約物半角",
            leadingModel: "仮想ボディの上/右",
            rensuuji: true,
            rotateSingleByte: false,
            absorbLineEndIdeographicSpace: true,
            latinWordBreak: false,
            kerningMethod: "和文等幅",
            autoLeading: 175,
            characterAlignment: "仮想ボディの中央",
            gridAlignment: "なし",
            composer: "日本語段落コンポーザー",
            hyphenation: true,
            hyphenateWordsLongerThan: 5,
            hyphenateAfterFirst: 2,
            hyphenateBeforeLast: 2,
            hyphenateLadderLimit: 3,
            hyphenationZone: 10,
            hyphenateCapitalizedWords: true,
            hyphenateAcrossColumns: true,
            hyphenateLastWord: true,
            ligatures: true,
            language: "ja"
        },
        appPreferences: {
            useSmartQuotes: false,
            doubleQuotes: "“”",
            singleQuotes: "‘’",
            textSizeUnit: "ポイント",
            compositionUnit: "ポイント"
        }
    }
};

/**
 * 段落スタイル向けプリセット項目の定義を作る
 * @returns {object} プリセット項目の定義
 */
function createStylePresetFields(dialogUi, dialogData) {
    return [
        { key: "kinsoku", type: "dd", control: dialogUi.kinsokuDropdown, names: dialogData.kinsokuNames },
        { key: "kinsokuType", type: "dd", control: dialogUi.kinsokuTypeDropdown, names: dialogData.kinsokuTypeNames },
        { key: "kinsokuHangType", type: "dd", control: dialogUi.kinsokuHangTypeDropdown, names: dialogData.kinsokuHangTypeNames },
        { key: "bunriKinshi", type: "cb", control: dialogUi.bunriKinshiCheckbox },
        { key: "mojikumi", type: "dd", control: dialogUi.mojikumiDropdown, names: dialogData.mojikumiNames },
        { key: "leadingModel", type: "dd", control: dialogUi.leadingModelDropdown, names: dialogData.leadingModelNames },
        { key: "rensuuji", type: "cb", control: dialogUi.rensuujiCheckbox },
        { key: "rotateSingleByte", type: "cb", control: dialogUi.rotateSingleByteCheckbox },
        { key: "absorbLineEndIdeographicSpace", type: "cb", control: dialogUi.absorbLineEndIdeographicSpaceCheckbox },
        { key: "latinWordBreak", type: "cb", control: dialogUi.latinWordBreakCheckbox },
        { key: "kerningMethod", type: "dd", control: dialogUi.kerningMethodDropdown, names: dialogData.kerningMethodNames },
        { key: "autoLeading", type: "in", control: dialogUi.autoLeadingInput },
        { key: "characterAlignment", type: "dd", control: dialogUi.characterAlignmentDropdown, names: dialogData.characterAlignmentNames },
        { key: "gridAlignment", type: "dd", control: dialogUi.gridAlignmentDropdown, names: dialogData.gridAlignmentNames },
        { key: "composer", type: "dd", control: dialogUi.composerDropdown, names: dialogData.composerNames },
        { key: "hyphenation", type: "cb", control: dialogUi.hyphenationCheckbox },
        { key: "hyphenateWordsLongerThan", type: "in", control: dialogUi.hyphenateWordsLongerThanInput },
        { key: "hyphenateAfterFirst", type: "in", control: dialogUi.hyphenateAfterFirstInput },
        { key: "hyphenateBeforeLast", type: "in", control: dialogUi.hyphenateBeforeLastInput },
        { key: "hyphenateLadderLimit", type: "in", control: dialogUi.hyphenateLadderLimitInput },
        { key: "hyphenationZone", type: "in", control: dialogUi.hyphenationZoneInput },
        { key: "hyphenateCapitalizedWords", type: "cb", control: dialogUi.hyphenateCapitalizedWordsCheckbox },
        { key: "hyphenateAcrossColumns", type: "cb", control: dialogUi.hyphenateAcrossColumnsCheckbox },
        { key: "hyphenateLastWord", type: "cb", control: dialogUi.hyphenateLastWordCheckbox },
        { key: "ligatures", type: "cb", control: dialogUi.ligaturesCheckbox }
    ];
}

/**
 * 環境設定向けプリセット項目の定義を作る
 * @returns {object} プリセット項目の定義
 */
function createAppPreferencePresetFields(dialogUi) {
    return [
        { key: "textSizeUnit", type: "dd", control: dialogUi.textSizeUnitDropdown, names: ["ポイント", "級", "アメリカ式ポイント"] },
        { key: "compositionUnit", type: "dd", control: dialogUi.compositionUnitDropdown, names: ["ポイント", "歯", "U", "倍", "ミルス", "アメリカ式ポイント"] }
    ];
}

/**
 * プリセット項目の定義をまとめて作る
 * @returns {object} プリセット項目の定義
 */
function createPresetFields(dialogUi, dialogData) {
    return {
        styleFields: createStylePresetFields(dialogUi, dialogData),
        appPreferenceFields: createAppPreferencePresetFields(dialogUi)
    };
}

/**
 * addNumberRow() で作った行（項目名・∧∨・入力欄）の有効／無効を切り替え、∧∨を描き直す
 * @param {EditText} numberInput addNumberRow() で作った入力欄
 * @param {boolean} isEnabled 有効にするなら true
 * @returns {void}
 */
function setNumberRowEnabled(numberInput, isEnabled) {
    numberInput.fieldRow.enabled = isEnabled;
    redrawSteppersIn(numberInput.fieldRow);
}

/**
 * ハイフネーションの ON/OFF に応じて関連項目を切り替える
 * @returns {void}
 */
function updateHyphenationControlsEnabled(dialogUi) {
    var isEnabled = dialogUi.hyphenationCheckbox.value;
    setNumberRowEnabled(dialogUi.hyphenateWordsLongerThanInput, isEnabled);
    setNumberRowEnabled(dialogUi.hyphenateAfterFirstInput, isEnabled);
    setNumberRowEnabled(dialogUi.hyphenateBeforeLastInput, isEnabled);
    setNumberRowEnabled(dialogUi.hyphenateLadderLimitInput, isEnabled);
    setNumberRowEnabled(dialogUi.hyphenationZoneInput, isEnabled);
    dialogUi.hyphenateCapitalizedWordsCheckbox.enabled = isEnabled;
    dialogUi.hyphenateAcrossColumnsCheckbox.enabled = isEnabled;
    dialogUi.hyphenateLastWordCheckbox.enabled = isEnabled;
}

/**
 * 対象範囲のラジオボタンを選択状態にする
 * @param {string} targetValue 選択する対象範囲
 * @returns {void}
 */
function activateTargetRadio(dialogUi, activeRadio) {
    var targetRadios = [dialogUi.targetAllRadio, dialogUi.targetSelectionRadio, dialogUi.targetSelectedParagraphsRadio];
    for (var radioIndex = 0; radioIndex < targetRadios.length; radioIndex++) {
        targetRadios[radioIndex].value = (targetRadios[radioIndex] === activeRadio);
    }
}

/**
 * 言語のラジオボタンを選択状態にする
 * @param {string} languageKey 選択する言語のキー
 * @returns {void}
 */
function activateLanguageRadio(dialogUi, activeRadio) {
    var languageRadios = [dialogUi.languageJapaneseRadio, dialogUi.languageEnglishRadio, dialogUi.languageNoneRadio];
    for (var radioIndex = 0; radioIndex < languageRadios.length; radioIndex++) {
        languageRadios[radioIndex].value = (languageRadios[radioIndex] === activeRadio);
    }
}

/**
 * 選択中の言語キーを取得する
 * @returns {string} 言語のキー
 */
function getLanguageSelection(dialogUi) {
    if (dialogUi.languageEnglishRadio.value) return "en";
    if (dialogUi.languageNoneRadio.value) return "none";
    return "ja";
}

/**
 * 選択から最初の段落を取り出す
 * @returns {Paragraph|null} 段落。取得できない場合は null
 */
function getFirstParagraphFromSelection() {
    try {
        var selectionItems = app.selection;
        if (!selectionItems || selectionItems.length === 0) return null;
        var firstSelectedItem = selectionItems[0];
        if (firstSelectedItem.paragraphs && firstSelectedItem.paragraphs.length > 0) {
            return firstSelectedItem.paragraphs.firstItem();
        }
        if (firstSelectedItem.parentStory && firstSelectedItem.parentStory.paragraphs && firstSelectedItem.parentStory.paragraphs.length > 0) {
            return firstSelectedItem.parentStory.paragraphs.firstItem();
        }
    } catch (selectionParagraphReadError) { }
    return null;
}

/**
 * 言語オブジェクトから言語キーを求める
 * @param {*} languageValue 言語の設定値
 * @returns {string} 言語のキー
 */
function resolveLanguageKey(appliedLanguage) {
    var languageName = "";
    if (appliedLanguage === null || appliedLanguage === undefined) {
        languageName = "[言語なし]";
    } else if (typeof appliedLanguage === "string") {
        languageName = appliedLanguage;
    } else {
        try { languageName = appliedLanguage.name; } catch (languageNameError) { }
    }

    for (var languageKey in LANGUAGE_CANDIDATES) {
        if (findIndexInArray(LANGUAGE_CANDIDATES[languageKey], languageName) >= 0) return languageKey;
    }

    return null;
}

/**
 * 選択段落から禁則関連の設定を読み取る
 * @param {Paragraph} paragraph 対象の段落
 * @param {object} lookupTables 選択肢の参照表
 * @returns {object} 読み取った設定
 */
function loadKinsokuSettingsFromParagraph(paragraphObject, appliedParagraphStyle, dialogUi, dialogData, dialogLookupTables) {
    var kinsokuIndex = -1;
    if (appliedParagraphStyle) {
        try { kinsokuIndex = findKinsokuIndexFromValue(appliedParagraphStyle.kinsokuSet, dialogLookupTables.kinsokuTables, dialogData.kinsokuNames); } catch (eKinsokuFromStyle) { }
    }
    if (kinsokuIndex < 0) {
        try { kinsokuIndex = findKinsokuIndexFromValue(paragraphObject.kinsokuSet, dialogLookupTables.kinsokuTables, dialogData.kinsokuNames); } catch (eKinsokuFromParagraph) { }
    }
    if (kinsokuIndex >= 0) {
        dialogUi.kinsokuDropdown.selection = kinsokuIndex;
    }

    var assignedKinsokuType = false;
    try {
        assignedKinsokuType = safeAssignDropdownFromEnum(dialogUi.kinsokuTypeDropdown, dialogLookupTables.kinsokuTypeValues, paragraphObject.kinsokuType);
    } catch (eKinsokuTypeFromParagraph) { }
    if (!assignedKinsokuType && appliedParagraphStyle) {
        try { safeAssignDropdownFromEnum(dialogUi.kinsokuTypeDropdown, dialogLookupTables.kinsokuTypeValues, appliedParagraphStyle.kinsokuType); } catch (eKinsokuTypeFromStyle) { }
    }

    var assignedKinsokuHangType = false;
    try {
        assignedKinsokuHangType = safeAssignDropdownFromEnum(dialogUi.kinsokuHangTypeDropdown, dialogLookupTables.kinsokuHangTypeValues, paragraphObject.kinsokuHangType);
    } catch (eKinsokuHangTypeFromParagraph) { }
    if (!assignedKinsokuHangType && appliedParagraphStyle) {
        try { safeAssignDropdownFromEnum(dialogUi.kinsokuHangTypeDropdown, dialogLookupTables.kinsokuHangTypeValues, appliedParagraphStyle.kinsokuHangType); } catch (eKinsokuHangTypeFromStyle) { }
    }
}

/**
 * 選択段落からハイフネーション設定を読み取る
 * @param {Paragraph} paragraph 対象の段落
 * @returns {object} 読み取った設定
 */
function loadHyphenationSettingsFromParagraph(paragraphObject, dialogUi) {
    safeAssignCheckbox(dialogUi.hyphenationCheckbox, paragraphObject.hyphenation);
    safeAssignNumberInput(dialogUi.hyphenateWordsLongerThanInput, paragraphObject.hyphenateWordsLongerThan);
    safeAssignNumberInput(dialogUi.hyphenateAfterFirstInput, paragraphObject.hyphenateAfterFirst);
    safeAssignNumberInput(dialogUi.hyphenateBeforeLastInput, paragraphObject.hyphenateBeforeLast);
    try { safeAssignNumberInput(dialogUi.hyphenateLadderLimitInput, paragraphObject.hyphenateLadderLimit); } catch (hyphenateLadderLimitReadError) { }

    try {
        var savedMeasurementUnit = app.scriptPreferences.measurementUnit;
        app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
        var hyphenationZonePt;
        try {
            hyphenationZonePt = paragraphObject.hyphenationZone;
        } catch (eHyphenationZoneRead) {
            hyphenationZonePt = null;
        }
        app.scriptPreferences.measurementUnit = savedMeasurementUnit;
        if (typeof hyphenationZonePt === "number") {
            var hyphenationZoneMm = hyphenationZonePt / 2.834645669;
            dialogUi.hyphenationZoneInput.text = String(Math.round(hyphenationZoneMm * 100) / 100);
        }
    } catch (eHyphenationZone) { }

    try { safeAssignCheckbox(dialogUi.hyphenateCapitalizedWordsCheckbox, paragraphObject.hyphenateCapitalizedWords); } catch (eHyphenateCapitalizedWords) { }
    try { safeAssignCheckbox(dialogUi.hyphenateAcrossColumnsCheckbox, paragraphObject.hyphenateAcrossColumns); } catch (eHyphenateAcrossColumns) { }
    try { safeAssignCheckbox(dialogUi.hyphenateLastWordCheckbox, paragraphObject.hyphenateLastWord); } catch (eHyphenateLastWord) { }
}

/**
 * 選択段落から組版設定をまとめて読み取る
 * @param {Paragraph} paragraph 対象の段落
 * @param {object} lookupTables 選択肢の参照表
 * @returns {object} 読み取った設定
 */
function loadSettingsFromParagraph(paragraphObject, dialogUi, dialogData) {
    if (!paragraphObject || !dialogData.lookupTables) return;

    var dialogLookupTables = dialogData.lookupTables;
    var appliedParagraphStyle = null;
    try {
        if (paragraphObject.appliedParagraphStyle && paragraphObject.appliedParagraphStyle.isValid) {
            appliedParagraphStyle = paragraphObject.appliedParagraphStyle;
        }
    } catch (eAppliedStyle) { }
    loadKinsokuSettingsFromParagraph(paragraphObject, appliedParagraphStyle, dialogUi, dialogData, dialogLookupTables);
    try {
        safeAssignDropdownFromName(dialogUi.mojikumiDropdown, dialogData.mojikumiNames, resolveMojikumiName(paragraphObject.mojikumi));
    } catch (eMojikumi) { }

    safeAssignDropdownFromEnum(dialogUi.leadingModelDropdown, dialogLookupTables.leadingModelValues, paragraphObject.leadingModel);
    safeAssignDropdownFromEnum(dialogUi.characterAlignmentDropdown, dialogLookupTables.characterAlignmentValues, paragraphObject.characterAlignment);
    safeAssignDropdownFromEnum(dialogUi.gridAlignmentDropdown, dialogLookupTables.gridAlignmentValues, paragraphObject.gridAlignment);
    safeAssignDropdownFromEnum(dialogUi.kerningMethodDropdown, dialogLookupTables.kerningMethodValues, paragraphObject.kerningMethod);
    safeAssignNumberInput(dialogUi.autoLeadingInput, paragraphObject.autoLeading);
    try {
        var matchedComposerIndex = findIndexByComposerAliases(dialogLookupTables.composerAliases, paragraphObject.composer);
        if (matchedComposerIndex >= 0) dialogUi.composerDropdown.selection = matchedComposerIndex;
    } catch (eComposer) { }
    safeAssignCheckbox(dialogUi.bunriKinshiCheckbox, paragraphObject.bunriKinshi);
    safeAssignCheckbox(dialogUi.rensuujiCheckbox, paragraphObject.rensuuji);
    safeAssignCheckbox(dialogUi.rotateSingleByteCheckbox, paragraphObject.rotateSingleByteCharacters);
    safeAssignCheckbox(dialogUi.absorbLineEndIdeographicSpaceCheckbox, paragraphObject.treatIdeographicSpaceAsSpace);
    try {
        safeAssignCheckbox(dialogUi.latinWordBreakCheckbox, paragraphObject.allowArbitraryHyphenation);
    } catch (eLatinWordBreakFromParagraph) {
        if (appliedParagraphStyle) {
            try { safeAssignCheckbox(dialogUi.latinWordBreakCheckbox, appliedParagraphStyle.allowArbitraryHyphenation); } catch (eLatinWordBreakFromStyle) { }
        }
    }
    try { safeAssignCheckbox(dialogUi.ligaturesCheckbox, paragraphObject.ligatures); } catch (eLigatures) { }
    loadHyphenationSettingsFromParagraph(paragraphObject, dialogUi);

    try {
        var matchedLanguageKey = resolveLanguageKey(paragraphObject.appliedLanguage);
        if (matchedLanguageKey === "ja") activateLanguageRadio(dialogUi, dialogUi.languageJapaneseRadio);
        else if (matchedLanguageKey === "en") activateLanguageRadio(dialogUi, dialogUi.languageEnglishRadio);
        else if (matchedLanguageKey === "none") activateLanguageRadio(dialogUi, dialogUi.languageNoneRadio);
    } catch (eLang) { }

    updateHyphenationControlsEnabled(dialogUi);
}

/**
 * プリセットの値をダイアログへ反映する
 * @param {string} presetKey プリセットのキー
 * @returns {void}
 */
function applyPreset(presetName, dialogUi, presetFields) {
    var preset = PRESETS[presetName];
    if (!preset) return;
    var presetStyleSettings = preset.styleSettings || preset;
    var presetAppPreferences = preset.appPreferences || preset;

    var mergedPresetFields = presetFields.styleFields.concat(presetFields.appPreferenceFields);

    for (var fieldIndex = 0; fieldIndex < mergedPresetFields.length; fieldIndex++) {
        var presetField = mergedPresetFields[fieldIndex];
        if (presetStyleSettings[presetField.key] === undefined && presetAppPreferences[presetField.key] === undefined) continue;
        var presetValue = presetStyleSettings[presetField.key] !== undefined ? presetStyleSettings[presetField.key] : presetAppPreferences[presetField.key];
        if (presetField.type === "dd") {
            safeAssignDropdownFromName(presetField.control, presetField.names, presetValue);
        } else if (presetField.type === "cb") {
            presetField.control.value = !!presetValue;
        } else if (presetField.type === "in") {
            presetField.control.text = String(presetValue);
        }
    }

    if (presetAppPreferences.useSmartQuotes !== undefined) {
        dialogUi.useTypographersQuotesCheckbox.value = !!presetAppPreferences.useSmartQuotes;
    }
    if (presetAppPreferences.doubleQuotes !== undefined) {
        safeAssignDropdownFromName(dialogUi.smartQuoteDropdown, DOUBLE_QUOTE_OPTIONS, presetAppPreferences.doubleQuotes);
    }
    if (presetAppPreferences.singleQuotes !== undefined) {
        safeAssignDropdownFromName(dialogUi.smartSingleQuoteDropdown, SINGLE_QUOTE_OPTIONS, presetAppPreferences.singleQuotes);
    }
    if (presetStyleSettings.language !== undefined) {
        if (presetStyleSettings.language === "en") activateLanguageRadio(dialogUi, dialogUi.languageEnglishRadio);
        else if (presetStyleSettings.language === "none") activateLanguageRadio(dialogUi, dialogUi.languageNoneRadio);
        else activateLanguageRadio(dialogUi, dialogUi.languageJapaneseRadio);
    }
    updateHyphenationControlsEnabled(dialogUi);
}

/**
 * プリセット名を入力するダイアログを表示する
 * @returns {string|null} 入力された名前。キャンセル時は null
 */
function showPresetNameInputDialog() {
    var nameDialog = new Window("dialog", getLabel("dialog.presetNameInput"));
    setupWindow(nameDialog, 10);

    nameDialog.add("statictext", undefined, labelText("field.presetName"));
    var nameInput = nameDialog.add("edittext", undefined, "");
    nameInput.preferredSize = [320, -1];
    nameInput.active = true;

    var buttonRow = addButtonRow(nameDialog);
    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    alignRightOnlyButtonRow(buttonRow);

    if (nameDialog.show() !== 1) return null;
    var presetName = nameInput.text;
    if (!presetName) return null;
    return presetName.replace(/^\s+|\s+$/g, "");
}

/**
 * 現在の設定をプリセット定義のコードとして組み立てる
 * @param {string} presetName プリセット名
 * @returns {string} 書き出すコード
 */
function buildPresetCodeSnippet(presetName, presetFields, dialogUi) {
    var appPreferenceKeys = {
        useSmartQuotes: true,
        doubleQuotes: true,
        singleQuotes: true,
        textSizeUnit: true,
        compositionUnit: true
    };

    var styleSettingLines = [];
    var appPreferenceLines = [];

    styleSettingLines.push("        language: \"" + getLanguageSelection(dialogUi) + "\"");
    appPreferenceLines.push("        useSmartQuotes: " + (dialogUi.useTypographersQuotesCheckbox.value ? "true" : "false"));
    appPreferenceLines.push("        doubleQuotes: \"" + (dialogUi.smartQuoteDropdown.selection ? dialogUi.smartQuoteDropdown.selection.text : "") + "\"");
    appPreferenceLines.push("        singleQuotes: \"" + (dialogUi.smartSingleQuoteDropdown.selection ? dialogUi.smartSingleQuoteDropdown.selection.text : "") + "\"");

    var mergedPresetFields = presetFields.styleFields.concat(presetFields.appPreferenceFields);

    for (var fieldIndex = 0; fieldIndex < mergedPresetFields.length; fieldIndex++) {
        var presetField = mergedPresetFields[fieldIndex];
        var valueText;

        if (presetField.type === "dd") {
            valueText = "\"" + presetField.control.selection.text + "\"";
        } else if (presetField.type === "cb") {
            valueText = presetField.control.value ? "true" : "false";
        } else {
            valueText = presetField.control.text;
            if (isNaN(parseFloat(valueText)) || String(parseFloat(valueText)) !== valueText) {
                valueText = "\"" + valueText + "\"";
            }
        }

        if (appPreferenceKeys[presetField.key]) {
            appPreferenceLines.push("        " + presetField.key + ": " + valueText);
        } else {
            styleSettingLines.push("        " + presetField.key + ": " + valueText);
        }
    }

    var lines = [];
    lines.push(getLabel("export.codeHeader"));
    lines.push("PRESETS[\"" + presetName + "\"] = {");
    lines.push("    styleSettings: {");
    lines.push(styleSettingLines.join(",\n"));
    lines.push("    },");
    lines.push("    appPreferences: {");
    lines.push(appPreferenceLines.join(",\n"));
    lines.push("    }");
    lines.push("};");
    return lines.join("\n");
}

/**
 * ファイル名に使えない文字を置き換える
 * @param {string} fileName 元のファイル名
 * @returns {string} 安全なファイル名
 */
function sanitizeFileName(name) {
    return name.replace(/[\/\\:*?"<>|]/g, "_");
}

/**
 * 現在の設定をプリセットコードとしてデスクトップへ書き出す
 * @returns {void}
 */
function exportPresetCode(presetFields, dialogUi) {
    var presetName = showPresetNameInputDialog();
    if (!presetName) return;

    var code = buildPresetCodeSnippet(presetName, presetFields, dialogUi);
    var safeFileName = sanitizeFileName(presetName);

    try {
        var file = File(Folder.desktop + "/" + encodeURI(safeFileName) + ".jsx");
        if (file.exists) {
            if (!confirm(getLabel("export.overwritePrefix") + safeFileName + getLabel("export.overwriteSuffix"))) return;
        }
        file.encoding = "UTF-8";
        if (file.open("w")) {
            file.write(code);
            file.close();
            var savedMessage = getLabel("export.savedPrefix") + presetName + getLabel("export.savedSuffix");
            if (safeFileName !== presetName) {
                savedMessage += "\n" + labelValueText("field.fileName", safeFileName + ".jsx");
            }
            alert(savedMessage);
        } else {
            alert(getLabel("alert.openFileFailed"));
        }
    } catch (eExport) {
        alert(getLabel("alert.exportErrorPrefix") + eExport.message);
    }
}

/**
 * 文字組版設定ダイアログを組み立てる
 * @param {object} lookupTables 選択肢の参照表
 * @param {object} initialSettings 初期値
 * @returns {object} ダイアログとコントロール
 */
function createDialogUI(dialogData) {
    var defaultIndexes = dialogData.defaultIndexes;
    var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(dialog, 10);

    var topColumnsGroup = dialog.add("group");
    topColumnsGroup.orientation = "row";
    topColumnsGroup.alignChildren = ["fill", "top"];
    topColumnsGroup.spacing = COLUMN_SPACING;

    var targetPanel = topColumnsGroup.add("panel", undefined, getLabel("panel.targetStyles"));
    setupPanel(targetPanel, 8);
    targetPanel.orientation = "row";
    targetPanel.alignChildren = ["left", "center"];

    var targetSelectedParagraphsRadio = targetPanel.add("radiobutton", undefined, getLabel("radio.targetSelection"));
    var targetAllRadio = targetPanel.add("radiobutton", undefined, getLabel("radio.targetAll"));
    var targetSelectionRadio = targetPanel.add("radiobutton", undefined, getLabel("radio.targetSpecified"));
    var btnSelectStyles = targetPanel.add("button", undefined, getLabel("button.select"));
    targetSelectedParagraphsRadio.value = true;

    var presetPanel = topColumnsGroup.add("panel", undefined, getLabel("panel.preset"));
    setupPanel(presetPanel, 8);
    presetPanel.orientation = "row";

    var presetDropdown = presetPanel.add("dropdownlist", undefined, ["欧文組版", "グリッド優先", "グリッド無視", "ソースコード", "InDesignのデフォルト"]);
    presetDropdown.selection = null;
    presetDropdown.preferredSize.width = W_DROP;
    var btnExportPreset = presetPanel.add("button", undefined, getLabel("button.export"));

    var columnsGroup = dialog.add("group");
    columnsGroup.orientation = "row";
    columnsGroup.alignChildren = ["fill", "top"];
    columnsGroup.spacing = COLUMN_SPACING;

    var leftColumn = columnsGroup.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = "fill";
    leftColumn.spacing = 10;

    var rightColumn = columnsGroup.add("group");
    rightColumn.orientation = "column";
    rightColumn.alignChildren = "fill";
    rightColumn.spacing = 10;

    var compositionExtraPanel = leftColumn.add("panel", undefined, getLabel("panel.basicSettings"));
    setupPanel(compositionExtraPanel, 8);

    var kerningMethodDropdown = addDropdownRow(compositionExtraPanel, labelText("field.autoKerning"), dialogData.kerningMethodNames, defaultIndexes.kerningMethodIndex);
    var autoLeadingInput = addNumberRow(compositionExtraPanel, labelText("field.autoLeading"), defaultIndexes.autoLeadingPercent, "%", { step: 1, min: 0, max: 500 });
    var characterAlignmentDropdown = addDropdownRow(compositionExtraPanel, labelText("field.characterAlign"), dialogData.characterAlignmentNames, defaultIndexes.characterAlignmentIndex);
    var leadingModelDropdown = addDropdownRow(compositionExtraPanel, labelText("field.leadingModel"), dialogData.leadingModelNames, defaultIndexes.leadingModelIndex);
    var gridAlignmentDropdown = addDropdownRow(compositionExtraPanel, labelText("field.gridAlignment"), dialogData.gridAlignmentNames, defaultIndexes.gridAlignmentIndex);
    var composerDropdown = addDropdownRow(compositionExtraPanel, labelText("field.composer"), dialogData.composerNames, defaultIndexes.composerIndex);

    var compositionOptionalPanel = rightColumn.add("panel", undefined, getLabel("panel.quotes"));
    setupPanel(compositionOptionalPanel, 8);

    var useTypographersQuotesInitial;

    try {
        useTypographersQuotesInitial = !!app.textPreferences.typographersQuotes;
    } catch (typographersQuotesReadError) {
        useTypographersQuotesInitial = true;
    }

    var useTypographersQuotesCheckbox = compositionOptionalPanel.add("checkbox", undefined, getLabel("checkbox.typographersQuotes"));

    useTypographersQuotesCheckbox.value = useTypographersQuotesInitial;

    var smartQuoteDropdown = addDropdownRow(
        compositionOptionalPanel,
        labelText("field.doubleQuote"),
        DOUBLE_QUOTE_OPTIONS,
        0
    );

    var smartSingleQuoteDropdown = addDropdownRow(
        compositionOptionalPanel,
        labelText("field.singleQuote"),
        SINGLE_QUOTE_OPTIONS,
        0
    );

    var unitsPanel = rightColumn.add("panel", undefined, getLabel("panel.units"));
    setupPanel(unitsPanel, 8);

    var textSizeUnitNames = ["ポイント", "級", "アメリカ式ポイント"];
    var textSizeUnitDropdown = addDropdownRow(unitsPanel, labelText("field.textSize"), textSizeUnitNames, 0);
    textSizeUnitDropdown.preferredSize.width = W_DROP;

    var compositionUnitNames = ["ポイント", "歯", "U", "倍", "ミルス", "アメリカ式ポイント"];
    var compositionUnitDropdown = addDropdownRow(unitsPanel, labelText("field.typography"), compositionUnitNames, 0);
    compositionUnitDropdown.preferredSize.width = W_DROP;

    var languageRow = compositionExtraPanel.add("group");
    setupRow(languageRow, "left", 8);

    var languageLabel = languageRow.add("statictext", undefined, labelText("field.language"));
    languageLabel.preferredSize.width = 120;

    var languageJapaneseRadio = languageRow.add("radiobutton", undefined, getLabel("radio.languageJapanese"));
    var languageEnglishRadio = languageRow.add("radiobutton", undefined, getLabel("radio.languageEnglish"));
    var languageNoneRadio = languageRow.add("radiobutton", undefined, getLabel("radio.languageNone"));
    languageJapaneseRadio.value = (defaultIndexes.language !== "en" && defaultIndexes.language !== "none");
    languageEnglishRadio.value = (defaultIndexes.language === "en");
    languageNoneRadio.value = (defaultIndexes.language === "none");

    var ligaturesCheckbox = compositionExtraPanel.add("checkbox", undefined, getLabel("checkbox.ligatures"));
    ligaturesCheckbox.value = defaultIndexes.ligatures;

    var compositionPanel = leftColumn.add("panel", undefined, getLabel("panel.japaneseTypeset"));
    setupPanel(compositionPanel, 8);

    var kinsokuDropdown = addDropdownRow(compositionPanel, labelText("field.kinsokuSet"), dialogData.kinsokuNames, defaultIndexes.kinsokuIndex);
    var kinsokuTypeDropdown = addDropdownRow(compositionPanel, labelText("field.kinsokuType"), dialogData.kinsokuTypeNames, defaultIndexes.kinsokuTypeIndex);
    var kinsokuHangTypeDropdown = addDropdownRow(compositionPanel, labelText("field.kinsokuHangType"), dialogData.kinsokuHangTypeNames, defaultIndexes.kinsokuHangTypeIndex);

    var bunriKinshiCheckbox = compositionPanel.add("checkbox", undefined, getLabel("checkbox.noBreak"));
    bunriKinshiCheckbox.value = defaultIndexes.bunriKinshi;

    var mojikumiDropdown = addDropdownRow(compositionPanel, labelText("field.mojikumi"), dialogData.mojikumiNames, defaultIndexes.mojikumiIndex);

    var compositionCheckboxesGroup = compositionPanel.add("group");
    compositionCheckboxesGroup.orientation = "row";
    compositionCheckboxesGroup.alignChildren = ["fill", "top"];
    compositionCheckboxesGroup.spacing = 16;
    compositionCheckboxesGroup.margins = [0, 10, 0, 0];

    var compositionCheckboxesLeft = compositionCheckboxesGroup.add("group");
    compositionCheckboxesLeft.orientation = "column";
    compositionCheckboxesLeft.alignChildren = "left";
    compositionCheckboxesLeft.spacing = 4;

    var compositionCheckboxesRight = compositionCheckboxesGroup.add("group");
    compositionCheckboxesRight.orientation = "column";
    compositionCheckboxesRight.alignChildren = "left";
    compositionCheckboxesRight.spacing = 4;

    var rensuujiCheckbox = compositionCheckboxesLeft.add("checkbox", undefined, getLabel("checkbox.digitsRotation"));
    rensuujiCheckbox.value = defaultIndexes.rensuuji;
    var rotateSingleByteCheckbox = compositionCheckboxesLeft.add("checkbox", undefined, getLabel("checkbox.rotateInVertical"));
    rotateSingleByteCheckbox.value = defaultIndexes.rotateSingleByte;
    var absorbLineEndIdeographicSpaceCheckbox = compositionCheckboxesRight.add("checkbox", undefined, getLabel("checkbox.absorbTrailingSpace"));
    absorbLineEndIdeographicSpaceCheckbox.value = defaultIndexes.absorbLineEndIdeographicSpace;
    var latinWordBreakCheckbox = compositionCheckboxesRight.add("checkbox", undefined, getLabel("checkbox.arbitraryHyphen"));
    latinWordBreakCheckbox.value = defaultIndexes.latinWordBreak;

    var hyphenationPanel = rightColumn.add("panel", undefined, getLabel("panel.hyphenation"));
    setupPanel(hyphenationPanel, 8);
    var hyphenationCheckbox = hyphenationPanel.add("checkbox", undefined, getLabel("checkbox.hyphenation"));
    hyphenationCheckbox.value = defaultIndexes.hyphenation;
    var hyphenateWordsLongerThanInput = addNumberRow(hyphenationPanel, labelText("field.minWordLength"), defaultIndexes.hyphenateWordsLongerThan, getLabel("unit.character"), { step: 1, min: 3, integer: true });
    var hyphenateAfterFirstInput = addNumberRow(hyphenationPanel, labelText("field.afterFirst"), defaultIndexes.hyphenateAfterFirst, getLabel("unit.character"), { step: 1, min: 1, integer: true });
    var hyphenateBeforeLastInput = addNumberRow(hyphenationPanel, labelText("field.beforeLast"), defaultIndexes.hyphenateBeforeLast, getLabel("unit.character"), { step: 1, min: 1, integer: true });
    var hyphenateLadderLimitInput = addNumberRow(hyphenationPanel, labelText("field.maxHyphens"), defaultIndexes.hyphenateLadderLimit, getLabel("unit.hyphen"), { step: 1, min: 0, integer: true });
    var hyphenationZoneInput = addNumberRow(hyphenationPanel, labelText("field.hyphenationZone"), defaultIndexes.hyphenationZoneMm, "mm", { step: 1, min: 0 });

    var hyphenateBreakPanel = hyphenationPanel.add("panel", undefined, getLabel("panel.hyphenateBreak"));
    setupPanel(hyphenateBreakPanel, 8);
    hyphenateBreakPanel.alignChildren = "left";
    var hyphenateCapitalizedWordsCheckbox = hyphenateBreakPanel.add("checkbox", undefined, getLabel("checkbox.capitalizedWords"));
    hyphenateCapitalizedWordsCheckbox.value = defaultIndexes.hyphenateCapitalizedWords;
    var hyphenateAcrossColumnsCheckbox = hyphenateBreakPanel.add("checkbox", undefined, getLabel("checkbox.acrossColumns"));
    hyphenateAcrossColumnsCheckbox.value = defaultIndexes.hyphenateAcrossColumns;
    var hyphenateLastWordCheckbox = hyphenateBreakPanel.add("checkbox", undefined, getLabel("checkbox.lastWord"));
    hyphenateLastWordCheckbox.value = defaultIndexes.hyphenateLastWord;

    var buttonRow = addButtonRow(dialog);
    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    alignRightOnlyButtonRow(buttonRow);

    return {
        dialog: dialog,
        targetSelectedParagraphsRadio: targetSelectedParagraphsRadio,
        targetAllRadio: targetAllRadio,
        targetSelectionRadio: targetSelectionRadio,
        btnSelectStyles: btnSelectStyles,
        presetDropdown: presetDropdown,
        btnExportPreset: btnExportPreset,
        kinsokuDropdown: kinsokuDropdown,
        kinsokuTypeDropdown: kinsokuTypeDropdown,
        kinsokuHangTypeDropdown: kinsokuHangTypeDropdown,
        bunriKinshiCheckbox: bunriKinshiCheckbox,
        mojikumiDropdown: mojikumiDropdown,
        leadingModelDropdown: leadingModelDropdown,
        rensuujiCheckbox: rensuujiCheckbox,
        rotateSingleByteCheckbox: rotateSingleByteCheckbox,
        absorbLineEndIdeographicSpaceCheckbox: absorbLineEndIdeographicSpaceCheckbox,
        latinWordBreakCheckbox: latinWordBreakCheckbox,
        kerningMethodDropdown: kerningMethodDropdown,
        autoLeadingInput: autoLeadingInput,
        characterAlignmentDropdown: characterAlignmentDropdown,
        gridAlignmentDropdown: gridAlignmentDropdown,
        languageJapaneseRadio: languageJapaneseRadio,
        languageEnglishRadio: languageEnglishRadio,
        languageNoneRadio: languageNoneRadio,
        composerDropdown: composerDropdown,
        ligaturesCheckbox: ligaturesCheckbox,
        hyphenationCheckbox: hyphenationCheckbox,
        hyphenateWordsLongerThanInput: hyphenateWordsLongerThanInput,
        hyphenateAfterFirstInput: hyphenateAfterFirstInput,
        hyphenateBeforeLastInput: hyphenateBeforeLastInput,
        hyphenateLadderLimitInput: hyphenateLadderLimitInput,
        hyphenationZoneInput: hyphenationZoneInput,
        hyphenateCapitalizedWordsCheckbox: hyphenateCapitalizedWordsCheckbox,
        hyphenateAcrossColumnsCheckbox: hyphenateAcrossColumnsCheckbox,
        hyphenateLastWordCheckbox: hyphenateLastWordCheckbox,
        useTypographersQuotesCheckbox: useTypographersQuotesCheckbox,
        smartQuoteDropdown: smartQuoteDropdown,
        smartSingleQuoteDropdown: smartSingleQuoteDropdown,
        textSizeUnitDropdown: textSizeUnitDropdown,
        compositionUnitDropdown: compositionUnitDropdown,
        selectedStyleIndexes: null
    };
}

/**
 * ダイアログのコントロールにイベントを結び付ける
 * @returns {void}
 */
function bindDialogEvents(dialogUi, dialogData, presetFields) {
    dialogUi.targetAllRadio.onClick = function () {
        activateTargetRadio(dialogUi, dialogUi.targetAllRadio);
    };
    dialogUi.targetSelectionRadio.onClick = function () {
        activateTargetRadio(dialogUi, dialogUi.targetSelectionRadio);
    };
    dialogUi.targetSelectedParagraphsRadio.onClick = function () {
        activateTargetRadio(dialogUi, dialogUi.targetSelectedParagraphsRadio);
        loadSettingsFromParagraph(getFirstParagraphFromSelection(), dialogUi, dialogData);
    };

    dialogUi.btnSelectStyles.onClick = function () {
        var pickerResult = showParagraphStylePicker(dialogData.paragraphStyleNames, dialogUi.selectedStyleIndexes);
        if (pickerResult !== null) {
            dialogUi.selectedStyleIndexes = pickerResult;
            activateTargetRadio(dialogUi, dialogUi.targetSelectionRadio);
        }
    };

    dialogUi.languageJapaneseRadio.onClick = function () { activateLanguageRadio(dialogUi, dialogUi.languageJapaneseRadio); };
    dialogUi.languageEnglishRadio.onClick = function () { activateLanguageRadio(dialogUi, dialogUi.languageEnglishRadio); };
    dialogUi.languageNoneRadio.onClick = function () { activateLanguageRadio(dialogUi, dialogUi.languageNoneRadio); };

    dialogUi.hyphenationCheckbox.onClick = function () {
        updateHyphenationControlsEnabled(dialogUi);
    };
    dialogUi.presetDropdown.onChange = function () {
        if (!dialogUi.presetDropdown.selection) return;
        applyPreset(dialogUi.presetDropdown.selection.text, dialogUi, presetFields);
    };
    dialogUi.btnExportPreset.onClick = function () {
        exportPresetCode(presetFields, dialogUi);
    };
}

/**
 * 段落スタイルへ適用する設定を組み立てる
 * @returns {object} 適用する設定
 */
function buildStyleSettingsResult(dialogUi) {
    return {
        targetMode: dialogUi.targetAllRadio.value ? "all" : (dialogUi.targetSelectionRadio.value ? "specified" : "selectedParagraphs"),
        selectedStyleIndexes: dialogUi.selectedStyleIndexes,
        kinsokuIndex: dialogUi.kinsokuDropdown.selection.index,
        kinsokuTypeIndex: dialogUi.kinsokuTypeDropdown.selection.index,
        kinsokuHangTypeIndex: dialogUi.kinsokuHangTypeDropdown.selection.index,
        mojikumiIndex: dialogUi.mojikumiDropdown.selection.index,
        leadingModelIndex: dialogUi.leadingModelDropdown.selection.index,
        characterAlignmentIndex: dialogUi.characterAlignmentDropdown.selection.index,
        gridAlignmentIndex: dialogUi.gridAlignmentDropdown.selection.index,
        kerningMethodIndex: dialogUi.kerningMethodDropdown.selection.index,
        autoLeadingPercent: parseFloat(dialogUi.autoLeadingInput.text),
        composerIndex: dialogUi.composerDropdown.selection.index,
        hyphenation: dialogUi.hyphenationCheckbox.value,
        bunriKinshi: dialogUi.bunriKinshiCheckbox.value,
        rensuuji: dialogUi.rensuujiCheckbox.value,
        rotateSingleByte: dialogUi.rotateSingleByteCheckbox.value,
        absorbLineEndIdeographicSpace: dialogUi.absorbLineEndIdeographicSpaceCheckbox.value,
        latinWordBreak: dialogUi.latinWordBreakCheckbox.value,
        hyphenateWordsLongerThan: parseInt(dialogUi.hyphenateWordsLongerThanInput.text, 10),
        hyphenateAfterFirst: parseInt(dialogUi.hyphenateAfterFirstInput.text, 10),
        hyphenateBeforeLast: parseInt(dialogUi.hyphenateBeforeLastInput.text, 10),
        hyphenateLadderLimit: parseInt(dialogUi.hyphenateLadderLimitInput.text, 10),
        hyphenationZoneMm: parseFloat(dialogUi.hyphenationZoneInput.text),
        hyphenateCapitalizedWords: dialogUi.hyphenateCapitalizedWordsCheckbox.value,
        hyphenateAcrossColumns: dialogUi.hyphenateAcrossColumnsCheckbox.value,
        hyphenateLastWord: dialogUi.hyphenateLastWordCheckbox.value,
        ligatures: dialogUi.ligaturesCheckbox.value,
        language: getLanguageSelection(dialogUi)
    };
}

/**
 * 環境設定へ適用する設定を組み立てる
 * @returns {object} 適用する設定
 */
function buildAppPreferencesResult(dialogUi) {
    return {
        useSmartQuotes: dialogUi.useTypographersQuotesCheckbox.value,
        doubleQuotes: dialogUi.smartQuoteDropdown.selection
            ? dialogUi.smartQuoteDropdown.selection.text
            : null,
        singleQuotes: dialogUi.smartSingleQuoteDropdown.selection
            ? dialogUi.smartSingleQuoteDropdown.selection.text
            : null,
        textSizeUnit: dialogUi.textSizeUnitDropdown.selection
            ? dialogUi.textSizeUnitDropdown.selection.text
            : null,
        compositionUnit: dialogUi.compositionUnitDropdown.selection
            ? dialogUi.compositionUnitDropdown.selection.text
            : null
    };
}

/**
 * 段落スタイル用と環境設定用の結果をまとめる
 * @param {object} styleSettings 段落スタイル向けの設定
 * @param {object} appPreferences 環境設定向けの設定
 * @returns {object} まとめた設定
 */
function mergeDialogResults(styleSettingsResult, appPreferencesResult) {
    for (var appPreferenceKey in appPreferencesResult) {
        styleSettingsResult[appPreferenceKey] = appPreferencesResult[appPreferenceKey];
    }
    return styleSettingsResult;
}

/**
 * ダイアログの入力内容を結果オブジェクトにまとめる
 * @returns {object} 適用に使う設定
 */
function buildDialogResult(dialogUi) {
    return mergeDialogResults(
        buildStyleSettingsResult(dialogUi),
        buildAppPreferencesResult(dialogUi)
    );
}

/**
 * 文字組版設定ダイアログを表示する
 * @param {object} lookupTables 選択肢の参照表
 * @param {object} initialSettings 初期値
 * @returns {object|null} 設定内容。キャンセル時は null
 */
function showTypesettingSettingsDialog(kinsokuNames, kinsokuTypeNames, kinsokuHangTypeNames, mojikumiNames, leadingModelNames, characterAlignmentNames, gridAlignmentNames, kerningMethodNames, composerNames, paragraphStyleNames, defaultIndexes, lookupTables) {
    var dialogData = {
        kinsokuNames: kinsokuNames,
        kinsokuTypeNames: kinsokuTypeNames,
        kinsokuHangTypeNames: kinsokuHangTypeNames,
        mojikumiNames: mojikumiNames,
        leadingModelNames: leadingModelNames,
        characterAlignmentNames: characterAlignmentNames,
        gridAlignmentNames: gridAlignmentNames,
        kerningMethodNames: kerningMethodNames,
        composerNames: composerNames,
        paragraphStyleNames: paragraphStyleNames,
        defaultIndexes: defaultIndexes,
        lookupTables: lookupTables
    };

    var dialogUi = createDialogUI(dialogData);
    var presetFields = createPresetFields(dialogUi, dialogData);
    bindDialogEvents(dialogUi, dialogData, presetFields);
    updateHyphenationControlsEnabled(dialogUi);

    if (dialogUi.targetSelectedParagraphsRadio.value) {
        var selectedParagraph = getFirstParagraphFromSelection();

        if (selectedParagraph) {
            loadSettingsFromParagraph(selectedParagraph, dialogUi, dialogData);
        } else {
            activateTargetRadio(dialogUi, dialogUi.targetAllRadio);
        }
    }

    if (dialogUi.dialog.show() !== 1) return null;
    return buildDialogResult(dialogUi);
}

// =========================================
// オーバーライドの消去 / Clear overrides
// =========================================

/**
 * 選択範囲の文字オーバーライドを消去する
 * @returns {void}
 */
function clearTextOverridesInSelection() {
    var selectionItems = app.selection;
    if (!selectionItems || selectionItems.length === 0) return;

    for (var selectionIndex = 0; selectionIndex < selectionItems.length; selectionIndex++) {
        var selectedItem = selectionItems[selectionIndex];
        try { selectedItem.clearOverrides(OverrideType.ALL); } catch (clearItemOverridesError) { }
        try { selectedItem.texts[0].clearOverrides(OverrideType.ALL); } catch (clearTextOverridesError) { }
        try { selectedItem.paragraphs.everyItem().clearOverrides(OverrideType.ALL); } catch (clearParagraphOverridesError) { }
    }
}

/**
 * 選択があればオーバーライドを消去する
 * @returns {void}
 */
function clearOverridesIfActive() {
    clearTextOverridesInSelection();
    try { app.menuActions.itemByID(8489).invoke(); } catch (clearOverridesMenuActionError) { }
    try { app.redraw(); } catch (redrawError) { }
}

/**
 * 適用対象なら選択中の段落スタイルを追加する
 * @param {Array<ParagraphStyle>} styles 収集先の配列
 * @param {ParagraphStyle} paragraphStyle 追加する段落スタイル
 * @returns {void}
 */
function addSelectedParagraphStyleIfApplicable(resultStyles, allParagraphStyles, paragraphStyle) {
    if (!paragraphStyle || !paragraphStyle.isValid) return;
    if (findIndexInArray(resultStyles, paragraphStyle) >= 0) return;
    if (findIndexInArray(allParagraphStyles, paragraphStyle) < 0) return;
    resultStyles.push(paragraphStyle);
}

/**
 * 選択から対象の段落スタイルを集める
 * @returns {Array<ParagraphStyle>} 段落スタイルの配列
 */
function collectParagraphStylesFromSelection(allParagraphStyles) {
    var resultStyles = [];
    var selectionItems = app.selection;
    if (!selectionItems || selectionItems.length === 0) return resultStyles;

    for (var selectionIndex = 0; selectionIndex < selectionItems.length; selectionIndex++) {
        var selectedItem = selectionItems[selectionIndex];
        try {
            if (selectedItem.paragraphs && selectedItem.paragraphs.length > 0) {
                for (var paragraphIndex = 0; paragraphIndex < selectedItem.paragraphs.length; paragraphIndex++) {
                    addSelectedParagraphStyleIfApplicable(resultStyles, allParagraphStyles, selectedItem.paragraphs.item(paragraphIndex).appliedParagraphStyle);
                }
                continue;
            }
        } catch (eParagraphs) { }

        try {
            if (selectedItem.parentStory && selectedItem.parentStory.paragraphs && selectedItem.parentStory.paragraphs.length > 0) {
                for (var storyParagraphIndex = 0; storyParagraphIndex < selectedItem.parentStory.paragraphs.length; storyParagraphIndex++) {
                    addSelectedParagraphStyleIfApplicable(resultStyles, allParagraphStyles, selectedItem.parentStory.paragraphs.item(storyParagraphIndex).appliedParagraphStyle);
                }
            }
        } catch (eStory) { }
    }

    return resultStyles;
}

/**
 * 対象範囲の指定に応じて適用先の段落スタイルを求める
 * @param {string} targetMode 対象範囲を表す識別子
 * @param {Array<ParagraphStyle>} allStyles すべての段落スタイル
 * @param {Array<string>} specifiedNames 指定した段落スタイル名
 * @returns {Array<ParagraphStyle>} 適用先の段落スタイル
 */
function resolveTargetParagraphStyles(allParagraphStyles, dialogResult) {
    if (!dialogResult || dialogResult.targetMode === "all") {
        return allParagraphStyles;
    }

    if (dialogResult.targetMode === "specified") {
        var selectedStyles = [];
        var selectedIndexes = dialogResult.selectedStyleIndexes;
        if (!selectedIndexes || selectedIndexes.length === 0) return selectedStyles;

        for (var indexPosition = 0; indexPosition < selectedIndexes.length; indexPosition++) {
            var styleIndex = selectedIndexes[indexPosition];
            if (styleIndex >= 0 && styleIndex < allParagraphStyles.length) {
                selectedStyles.push(allParagraphStyles[styleIndex]);
            }
        }
        return selectedStyles;
    }

    if (dialogResult.targetMode === "selectedParagraphs") {
        return collectParagraphStylesFromSelection(allParagraphStyles);
    }

    return allParagraphStyles;
}

// =========================================
// 設定の適用 / Apply settings
// =========================================

/**
 * 辞書の引用符設定を適用する
 * @param {object} settings 適用する設定
 * @returns {void}
 */
function applyDictionaryQuoteSettings(dialogResult) {
    try {
        if (dialogResult.doubleQuotes) {
            app.languagesWithVendors.everyItem().doubleQuotes =
                dialogResult.doubleQuotes;
        }

        if (dialogResult.singleQuotes) {
            app.languagesWithVendors.everyItem().singleQuotes =
                dialogResult.singleQuotes;
        }
    } catch (dictionaryQuotesApplyError) { }
}

/**
 * 環境設定側の項目を適用する
 * @param {object} settings 適用する設定
 * @returns {void}
 */
function applyAppPreferenceSettings(dialogResult) {
    try {
        app.textPreferences.typographersQuotes =
            !!dialogResult.useSmartQuotes;
    } catch (typographersQuotesApplyError) { }

    applyDictionaryQuoteSettings(dialogResult);

    // テキストサイズの単位 / Text size measurement units
    try {
        if (dialogResult.textSizeUnit === "ポイント") app.viewPreferences.textSizeMeasurementUnits = TextSizeMeasurementUnits.POINTS;
        else if (dialogResult.textSizeUnit === "級") app.viewPreferences.textSizeMeasurementUnits = TextSizeMeasurementUnits.Q;
        else if (dialogResult.textSizeUnit === "アメリカ式ポイント") app.viewPreferences.textSizeMeasurementUnits = TextSizeMeasurementUnits.AMERICAN_POINTS;
    } catch (textSizeUnitApplyError) { }

    // 組版の単位 / Composition (vertical) measurement units
    try {
        if (dialogResult.compositionUnit === "ポイント") app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
        else if (dialogResult.compositionUnit === "歯") app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.HA;
        else if (dialogResult.compositionUnit === "U") app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.U;
        else if (dialogResult.compositionUnit === "倍") app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.BAI;
        else if (dialogResult.compositionUnit === "ミルス") app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILS;
        else if (dialogResult.compositionUnit === "アメリカ式ポイント") app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.AMERICAN_POINTS;
    } catch (compositionUnitApplyError) { }
}

/**
 * 確定した設定を対象の段落スタイルへ適用する
 * @param {Array<ParagraphStyle>} targetParagraphStyles 適用先の段落スタイル
 * @param {object} settings 適用する設定
 * @param {object} lookupTables 選択肢の参照表
 * @returns {void}
 */
function applyTypesettingSettingsToAll(targetParagraphStyles, dialogResult, lookupTables) {
    app.doScript(
        function () {
            var kinsokuTable = lookupTables.kinsokuTables[dialogResult.kinsokuIndex];
            var kinsokuTypeValue = lookupTables.kinsokuTypeValues[dialogResult.kinsokuTypeIndex];
            var kinsokuHangTypeValue = lookupTables.kinsokuHangTypeValues[dialogResult.kinsokuHangTypeIndex];
            var mojikumiTable = lookupTables.mojikumiTables[dialogResult.mojikumiIndex];
            var leadingModelValue = lookupTables.leadingModelValues[dialogResult.leadingModelIndex];
            var characterAlignmentValue = lookupTables.characterAlignmentValues[dialogResult.characterAlignmentIndex];
            var gridAlignmentValue = lookupTables.gridAlignmentValues[dialogResult.gridAlignmentIndex];
            var kerningMethodValue = lookupTables.kerningMethodValues[dialogResult.kerningMethodIndex];
            var autoLeadingPercent = dialogResult.autoLeadingPercent;
            var composerAliasList = lookupTables.composerAliases[dialogResult.composerIndex];
            var hyphenationValue = dialogResult.hyphenation;
            var bunriKinshiValue = dialogResult.bunriKinshi;

            var skipped = 0;
            var errorDetails = [];

            for (var styleIndex = 0; styleIndex < targetParagraphStyles.length; styleIndex++) {
                var paragraphStyle = targetParagraphStyles[styleIndex];
                try {
                    paragraphStyle.kinsokuSet = kinsokuTable;
                    paragraphStyle.kinsokuType = kinsokuTypeValue;
                    paragraphStyle.kinsokuHangType = kinsokuHangTypeValue;
                    paragraphStyle.mojikumi = mojikumiTable === null ? NothingEnum.NOTHING : mojikumiTable;
                    paragraphStyle.leadingModel = leadingModelValue;
                    paragraphStyle.characterAlignment = characterAlignmentValue;
                    paragraphStyle.gridAlignment = gridAlignmentValue;
                    paragraphStyle.kerningMethod = kerningMethodValue;
                    if (!isNaN(dialogResult.autoLeadingPercent)) {
                        paragraphStyle.autoLeading = dialogResult.autoLeadingPercent;
                    }
                    paragraphStyle.hyphenation = hyphenationValue;
                    paragraphStyle.bunriKinshi = bunriKinshiValue;
                } catch (applyError) {
                    skipped++;
                    var styleNameForError = "(unknown)";
                    try { styleNameForError = paragraphStyle.name; } catch (eName) { }
                    errorDetails.push("[" + styleNameForError + "] " + applyError);
                }

                // composer はロケールやバージョンで受理名が異なるため alias を順に試行 /
                // Composer name varies by locale/version; try aliases in order
                applyComposerAliases(paragraphStyle, composerAliasList);

                // プロパティ名が不確実なものは安全代入で適用 / Apply uncertain properties via safe assignment
                safeSetProperty(paragraphStyle, "rensuuji", dialogResult.rensuuji);
                safeSetProperty(paragraphStyle, "rotateSingleByteCharacters", dialogResult.rotateSingleByte);
                safeSetProperty(paragraphStyle, "treatIdeographicSpaceAsSpace", dialogResult.absorbLineEndIdeographicSpace);
                // 欧文泣き別れ / Latin word break (allowArbitraryHyphenation)
                safeSetProperty(paragraphStyle, "allowArbitraryHyphenation", dialogResult.latinWordBreak);

                // ハイフネーション詳細設定 / Hyphenation detail settings
                if (!isNaN(dialogResult.hyphenateWordsLongerThan)) {
                    safeSetProperty(paragraphStyle, "hyphenateWordsLongerThan", dialogResult.hyphenateWordsLongerThan);
                }
                if (!isNaN(dialogResult.hyphenateAfterFirst)) {
                    safeSetProperty(paragraphStyle, "hyphenateAfterFirst", dialogResult.hyphenateAfterFirst);
                }
                if (!isNaN(dialogResult.hyphenateBeforeLast)) {
                    safeSetProperty(paragraphStyle, "hyphenateBeforeLast", dialogResult.hyphenateBeforeLast);
                }
                if (!isNaN(dialogResult.hyphenateLadderLimit)) {
                    safeSetProperty(paragraphStyle, "hyphenateLadderLimit", dialogResult.hyphenateLadderLimit);
                }
                if (!isNaN(dialogResult.hyphenationZoneMm)) {
                    safeSetProperty(paragraphStyle, "hyphenationZone", dialogResult.hyphenationZoneMm + "mm");
                }

                safeSetProperty(paragraphStyle, "hyphenateCapitalizedWords", dialogResult.hyphenateCapitalizedWords);
                safeSetProperty(paragraphStyle, "hyphenateAcrossColumns", dialogResult.hyphenateAcrossColumns);
                safeSetProperty(paragraphStyle, "hyphenateLastWord", dialogResult.hyphenateLastWord);

                safeSetProperty(paragraphStyle, "ligatures", dialogResult.ligatures);

                if (dialogResult.language && LANGUAGE_CANDIDATES[dialogResult.language]) {
                    var languageCandidates = LANGUAGE_CANDIDATES[dialogResult.language];
                    for (var languageCandidateIndex = 0; languageCandidateIndex < languageCandidates.length; languageCandidateIndex++) {
                        try {
                            var candidateLanguage = app.languagesWithVendors.itemByName(languageCandidates[languageCandidateIndex]);
                            if (candidateLanguage.isValid) {
                                paragraphStyle.appliedLanguage = candidateLanguage;
                                break;
                            }
                        } catch (eLangApply) { }
                    }
                }
            }

            applyAppPreferenceSettings(dialogResult);

            if (skipped > 0) {
                alert(getLabel("alert.partialFailurePrefix") + skipped + getLabel("alert.partialFailureSuffix") + errorDetails.join("\n"));
            }
        },
        ScriptLanguage.JAVASCRIPT,
        undefined,
        UndoModes.ENTIRE_SCRIPT,
        getLabel("undo.applyTypesetting")
    );
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var activeDocument = app.activeDocument;

    var kinsokuTableData = collectKinsokuTables(activeDocument);
    if (kinsokuTableData.tables.length === 0) {
        alert(getLabel("alert.noKinsokuTables"));
        return;
    }

    var kinsokuTypeOptions = createKinsokuTypeOptions();
    var kinsokuHangTypeOptions = createKinsokuHangTypeOptions();
    var mojikumiTableData = collectMojikumiTables(activeDocument);
    var targetParagraphStyleData = collectTargetParagraphStyles(activeDocument);
    var targetParagraphStyles = targetParagraphStyleData.styles;

    if (targetParagraphStyles.length === 0) {
        alert(getLabel("alert.noParagraphStyles"));
        return;
    }

    var leadingModelOptions = createLeadingModelOptions();
    var characterAlignmentOptions = createCharacterAlignmentOptions();
    var gridAlignmentOptions = createGridAlignmentOptions();
    var kerningMethodOptions = createKerningMethodOptions();
    var composerOptions = createComposerOptions();
    var defaultPreset = PRESETS["InDesignのデフォルト"].styleSettings;
    var defaultIndexes = {
        kinsokuIndex: getDefaultIndexByName(kinsokuTableData.names, defaultPreset.kinsoku),
        kinsokuTypeIndex: getDefaultIndexByName(kinsokuTypeOptions.names, defaultPreset.kinsokuType),
        kinsokuHangTypeIndex: getDefaultIndexByName(kinsokuHangTypeOptions.names, defaultPreset.kinsokuHangType),
        mojikumiIndex: getDefaultIndexByName(mojikumiTableData.names, defaultPreset.mojikumi),
        leadingModelIndex: getDefaultIndexByName(leadingModelOptions.names, defaultPreset.leadingModel),
        characterAlignmentIndex: getDefaultIndexByName(characterAlignmentOptions.names, defaultPreset.characterAlignment),
        gridAlignmentIndex: getDefaultIndexByName(gridAlignmentOptions.names, defaultPreset.gridAlignment),
        kerningMethodIndex: getDefaultIndexByName(kerningMethodOptions.names, defaultPreset.kerningMethod),
        autoLeadingPercent: defaultPreset.autoLeading,
        composerIndex: (function () {
            var aliasMatch = findIndexByComposerAliases(composerOptions.aliases, defaultPreset.composer);
            return aliasMatch >= 0 ? aliasMatch : getDefaultIndexByName(composerOptions.names, defaultPreset.composer);
        })(),
        hyphenation: defaultPreset.hyphenation,
        bunriKinshi: defaultPreset.bunriKinshi,
        rensuuji: defaultPreset.rensuuji,
        rotateSingleByte: defaultPreset.rotateSingleByte,
        absorbLineEndIdeographicSpace: defaultPreset.absorbLineEndIdeographicSpace,
        latinWordBreak: defaultPreset.latinWordBreak,
        hyphenateWordsLongerThan: defaultPreset.hyphenateWordsLongerThan,
        hyphenateAfterFirst: defaultPreset.hyphenateAfterFirst,
        hyphenateBeforeLast: defaultPreset.hyphenateBeforeLast,
        hyphenateLadderLimit: defaultPreset.hyphenateLadderLimit,
        hyphenationZoneMm: defaultPreset.hyphenationZone,
        hyphenateCapitalizedWords: defaultPreset.hyphenateCapitalizedWords,
        hyphenateAcrossColumns: defaultPreset.hyphenateAcrossColumns,
        hyphenateLastWord: defaultPreset.hyphenateLastWord,
        ligatures: defaultPreset.ligatures,
        language: defaultPreset.language
    };

    var lookupTables = {
        kinsokuTables: kinsokuTableData.tables,
        kinsokuTypeValues: kinsokuTypeOptions.values,
        kinsokuHangTypeValues: kinsokuHangTypeOptions.values,
        mojikumiTables: mojikumiTableData.tables,
        leadingModelValues: leadingModelOptions.values,
        characterAlignmentValues: characterAlignmentOptions.values,
        gridAlignmentValues: gridAlignmentOptions.values,
        kerningMethodValues: kerningMethodOptions.values,
        composerAliases: composerOptions.aliases
    };

    var dialogResult = showTypesettingSettingsDialog(
        kinsokuTableData.names,
        kinsokuTypeOptions.names,
        kinsokuHangTypeOptions.names,
        mojikumiTableData.names,
        leadingModelOptions.names,
        characterAlignmentOptions.names,
        gridAlignmentOptions.names,
        kerningMethodOptions.names,
        composerOptions.names,
        targetParagraphStyleData.names,
        defaultIndexes,
        lookupTables
    );
    if (dialogResult === null) return;

    var resolvedTargetParagraphStyles = resolveTargetParagraphStyles(targetParagraphStyles, dialogResult);
    if (resolvedTargetParagraphStyles.length === 0) {
        alert(getLabel("alert.noTargetStyles"));
        return;
    }

    applyTypesettingSettingsToAll(resolvedTargetParagraphStyles, dialogResult, lookupTables);

    clearOverridesIfActive();

})();