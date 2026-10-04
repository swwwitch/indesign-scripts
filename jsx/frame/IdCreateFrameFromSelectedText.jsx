#target indesign

/*

### 概要

選択したテキストのサイズを基準に、インライン（アンカー付き）またはページ上へグラフィックフレームを作成します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCreateFrameFromSelectedText.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ndd1b7c5246a3

### Overview

Creates a graphic frame sized from the selected text, either inline (anchored) or placed on the page.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCreateFrameFromSelectedText.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdCreateFrameFromSelectedText"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.7.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCreateFrameFromSelectedText.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCreateFrameFromSelectedText.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ndd1b7c5246a3"; /* 紹介記事 / article URL */

// Original idea
// DTP Script note
// https://note.com/yosi2631/n/ned2dbc1cb79d

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 既定で選ばれる段落スタイル名の手がかり / Hints used to preselect a paragraph style */
var DEFAULT_PARA_STYLE_NAME = "p.img";
var DEFAULT_PARA_STYLE_HINT = ".img";

/* 「なし」オブジェクトスタイルの名前（日英）/ Names of the "None" object style (JA/EN) */
var NONE_OBJECT_STYLE_NAMES = ["[なし]", "[None]"];

/* 高さ入力の初期値 / Default values of the height field */
var DEFAULT_HEIGHT_IN_LINES = "1";
var DEFAULT_HEIGHT_IN_MM    = "40";

/* 高さの∧∨・↑↓キーで下げられる下限（行数は1行未満を1行として扱う） / lowest value the height stepper reaches */
var MIN_HEIGHT_IN_LINES = 1;
var MIN_HEIGHT_IN_MM    = 0.1;

/* 行送りを取得できないときのフォールバック（pt）/ Fallback used when leading cannot be read (pt) */
var FALLBACK_POINT_SIZE = 14;
var FALLBACK_LEADING    = 14;

/* mm をポイントへ換算する係数 / Millimeter-to-point conversion factor */
var MM_TO_POINTS = 2.834645669;

// =========================================
// レイアウト設定 / Layout settings
// =========================================

/* 高さ入力欄の文字数 / Character width of the height field */
var HEIGHT_INPUT_CHARACTERS = 8;

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
        title:            { ja: "選択した文字列からフレーム作成", en: "Create Frame from Selected Text" },
        titleNoSelection: { ja: "グラフィックフレーム作成", en: "Create Graphic Frame" }
    },
    panel: {
        method:         { ja: "配置方法", en: "Placement" },
        frameSize:      { ja: "フレームサイズ", en: "Frame Size" },
        width:          { ja: "幅", en: "Width" },
        height:         { ja: "高さ", en: "Height" },
        textSettings:   { ja: "テキスト設定", en: "Text Settings" },
        objectSettings: { ja: "オブジェクト設定", en: "Object Settings" }
    },
    radio: {
        graphicFrame: { ja: "グラフィックフレーム", en: "Graphic Frame" },
        inlineFrame:  { ja: "インライン（アンカー付き）", en: "Inline (Anchored)" },
        widthText:    { ja: "選択したテキスト", en: "Selected Text" },
        widthColumn:  { ja: "カラム幅", en: "Column Width" },
        widthFrame:   { ja: "親フレーム", en: "Parent Frame" },
        widthMargin:  { ja: "ページのマージン", en: "Page Margins" },
        heightLines:  { ja: "行数", en: "Line Count" },
        heightSize:   { ja: "サイズ指定", en: "Specify Size" }
    },
    field: {
        paraStyle: { ja: "挿入行の段落スタイル:", en: "Paragraph Style for Inserted Line:" },
        objStyle:  { ja: "オブジェクトスタイル:", en: "Object Style:" },
        wrap:      { ja: "テキストの回り込み:", en: "Text Wrap:" }
    },
    checkbox: {
        autoLeading: { ja: "行送り：自動", en: "Leading: Auto" }
    },
    wrap: {
        none:        { ja: "なし", en: "None" },
        boundingBox: { ja: "境界線ボックスで回り込む", en: "Wrap Around Bounding Box" },
        contour:     { ja: "オブジェクトのシェイプで回り込む", en: "Wrap Around Object Shape" },
        jumpObject:  { ja: "オブジェクトを挟んで回り込む", en: "Jump Object" },
        nextColumn:  { ja: "次の段へテキストを送る", en: "Jump to Next Column" }
    },
    unit: {
        lines: { ja: "行", en: "lines" },
        mm:    { ja: "mm", en: "mm" }
    },
    button: {
        ok:     { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    tooltip: {
        heightLines: {
            ja: "行数モードの高さは概算です。1行目の文字サイズ＋残り行数×行送りで計算します。",
            en: "Line Count height is approximate. It is calculated as first-line point size plus leading for the remaining lines."
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
        stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
        stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
    },
    alert: {
        noDocument:          { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        selectText:          { ja: "テキストを選択してください。", en: "Please select text." },
        selectTextBeforeRun: { ja: "テキスト項目を1つだけ選択してから実行してください。", en: "Select only one text item before running the script." },
        boundsError:         { ja: "選択したテキストの座標を取得できませんでした。", en: "Could not get the bounds of the selected text." },
        parentPageError:     { ja: "配置先のページを取得できませんでした。", en: "Could not determine the destination page." },
        parentFrameError:    { ja: "挿入位置の親テキストフレームを取得できませんでした。", en: "Could not get the parent text frame for the insertion point." }
    },
    undo: {
        createFrame: { ja: "フレーム作成", en: "Create Frame" }
    }
};

// =========================================
// テキスト判定と座標取得 / Text detection and bounds
// =========================================

/**
 * 選択オブジェクトがテキストかどうかを判定する
 * @param {object} selectionItem 選択オブジェクト
 * @returns {boolean} テキストなら true
 */
function isTextSelection(selectionItem) {
    if (!selectionItem) return false;

    try {
        if (selectionItem.hasOwnProperty("baseline")) return true;
    } catch (e) {}

    try {
        if (selectionItem.hasOwnProperty("characters") && selectionItem.characters && selectionItem.characters.length > 0) {
            return true;
        }
    } catch (e) {}

    try {
        if (selectionItem.constructor && selectionItem.constructor.name) {
            var typeName = String(selectionItem.constructor.name);
            if (typeName === "Text" || typeName === "Word" || typeName === "Line" ||
                typeName === "Paragraph" || typeName === "TextStyleRange" || typeName === "Character") {
                return true;
            }
        }
    } catch (e) {}

    return false;
}

/**
 * 2つの外接矩形を合わせた外接矩形を返す
 * @param {Array<number>|null} baseBounds [上, 左, 下, 右]。まだ無ければ null
 * @param {Array<number>} addedBounds 合わせる [上, 左, 下, 右]
 * @returns {Array<number>} 合わせた [上, 左, 下, 右]
 */
function unionBounds(baseBounds, addedBounds) {
    if (baseBounds === null) return [addedBounds[0], addedBounds[1], addedBounds[2], addedBounds[3]];
    return [
        Math.min(baseBounds[0], addedBounds[0]),
        Math.min(baseBounds[1], addedBounds[1]),
        Math.max(baseBounds[2], addedBounds[2]),
        Math.max(baseBounds[3], addedBounds[3])
    ];
}

/**
 * アウトライン化した一時オブジェクトから外接矩形を求め、そのオブジェクトを削除する
 * @param {Array} outlineItems アウトライン化で得られたオブジェクトの配列
 * @returns {Array<number>|null} [上, 左, 下, 右]。求められない場合は null
 */
function getAndRemoveOutlineBounds(outlineItems) {
    var mergedBounds = null;

    try {
        for (var i = 0; i < outlineItems.length; i++) {
            mergedBounds = unionBounds(mergedBounds, outlineItems[i].geometricBounds);
        }
    } finally {
        for (var j = outlineItems.length - 1; j >= 0; j--) {
            try { outlineItems[j].remove(); } catch (e) {}
        }
    }

    return mergedBounds;
}

/**
 * テキスト（選択全体または1文字）を一度にアウトライン化して外接矩形を求める
 * @param {object} textObject 対象のテキスト
 * @returns {Array<number>|null} [上, 左, 下, 右]。求められない場合は null
 */
function getOutlineBounds(textObject) {
    try {
        if (!textObject || !textObject.characters || textObject.characters.length === 0) return null;
        var outlineItems = textObject.createOutlines(false);
        if (!outlineItems || outlineItems.length === 0) return null;
        return getAndRemoveOutlineBounds(outlineItems);
    } catch (e) {
        return null;
    }
}

/**
 * 1 文字ずつアウトライン化して外接矩形を合成する
 * @param {object} textObject 対象のテキスト
 * @returns {Array<number>|null} [上, 左, 下, 右]。求められない場合は null
 */
function getTextSelectionBoundsPerCharacter(textObject) {
    try {
        var characterList = textObject.characters;
        if (!characterList || characterList.length === 0) return null;

        var mergedBounds = null;

        for (var i = 0; i < characterList.length; i++) {
            var character = characterList[i];

            try {
                /* 改行類はアウトライン化できないので除外 / Skip break characters, which cannot be outlined */
                if (character.contents === "\r" || character.contents === "\n" || character.contents === "\u0003") continue;
            } catch (e) {}

            var characterBounds = getOutlineBounds(character);
            if (!characterBounds) continue;
            mergedBounds = unionBounds(mergedBounds, characterBounds);
        }

        if (mergedBounds === null) {
            try {
                return textObject.parentTextFrames[0].geometricBounds;
            } catch (e) {
                return null;
            }
        }

        return mergedBounds;
    } catch (e) {
        return null;
    }
}

/**
 * 選択テキストの外接矩形を求める
 * @param {object} textObject 対象のテキスト
 * @returns {Array<number>|null} [上, 左, 下, 右]。求められない場合は null
 */
function getTextSelectionBounds(textObject) {
    return getOutlineBounds(textObject) || getTextSelectionBoundsPerCharacter(textObject);
}

/**
 * 挿入ポイントが属するテキストフレームを取得する
 * @param {InsertionPoint} insertionPoint 対象の挿入ポイント
 * @returns {TextFrame|null} テキストフレーム。取得できない場合は null
 */
function getInsertionPointTextFrame(insertionPoint) {
    try {
        if (insertionPoint && insertionPoint.parentTextFrames && insertionPoint.parentTextFrames.length > 0) {
            return insertionPoint.parentTextFrames[0];
        }
    } catch (e) {}
    return null;
}

/**
 * テキストが配置されているページを取得する
 * @param {object} textObject 対象のテキスト
 * @returns {Page|null} ページ。取得できない場合は null
 */
function getParentPage(textObject) {
    try {
        if (textObject.parentTextFrames.length > 0) return textObject.parentTextFrames[0].parentPage;
    } catch (e) {}
    return null;
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {

    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var selectedText     = null;
    var isInsertionPoint = false;
    var hasNoSelection   = false;

    if (app.selection.length === 0) {
        hasNoSelection = true;
    } else if (app.selection.length !== 1) {
        alert(getLabel("alert.selectTextBeforeRun"));
        return;
    } else {
        selectedText = app.selection[0];
        if (selectedText.constructor && selectedText.constructor.name === "InsertionPoint") {
            isInsertionPoint = true;
        } else if (!isTextSelection(selectedText)) {
            alert(getLabel("alert.selectText"));
            return;
        }
    }

    var activeDoc = app.activeDocument;

    /* 選択テキストの外接矩形（挿入ポイント・未選択時は取得しない）/ Bounds of the selection (skipped for insertion points and empty selections) */
    var selectionBounds = null;
    var hasTextBounds   = false;
    if (!isInsertionPoint && !hasNoSelection) {
        selectionBounds = getTextSelectionBounds(selectedText);
        if (!selectionBounds) {
            alert(getLabel("alert.boundsError"));
            return;
        }
        hasTextBounds = true;
    }

    var selectionWidth  = hasTextBounds ? (selectionBounds[3] - selectionBounds[1]) : 0;
    var selectionHeight = hasTextBounds ? (selectionBounds[2] - selectionBounds[0]) : 0;

    /* オブジェクトスタイル一覧 / Object style list */
    var objectStyleNames = [];
    for (var objectStyleIndex = 0; objectStyleIndex < activeDoc.objectStyles.length; objectStyleIndex++) {
        objectStyleNames.push(activeDoc.objectStyles[objectStyleIndex].name);
    }

    /* 段落スタイル一覧と既定選択位置 / Paragraph style list and the default selection */
    var paragraphStyleNames = [];
    var defaultParaStyleIndex = 0;
    for (var paraStyleIndex = 0; paraStyleIndex < activeDoc.paragraphStyles.length; paraStyleIndex++) {
        var paragraphStyleName = activeDoc.paragraphStyles[paraStyleIndex].name;
        paragraphStyleNames.push(paragraphStyleName);
        if (paragraphStyleName === DEFAULT_PARA_STYLE_NAME || paragraphStyleName.indexOf(DEFAULT_PARA_STYLE_HINT) !== -1) {
            defaultParaStyleIndex = paraStyleIndex;
        }
    }

    /* 回り込みの選択肢 / Text wrap options */
    var wrapOptionLabels = [
        getLabel("wrap.none"),
        getLabel("wrap.boundingBox"),
        getLabel("wrap.contour"),
        getLabel("wrap.jumpObject"),
        getLabel("wrap.nextColumn")
    ];
    var wrapOptionModes = [
        TextWrapModes.NONE,
        TextWrapModes.BOUNDING_BOX_TEXT_WRAP,
        TextWrapModes.CONTOUR,
        TextWrapModes.JUMP_OBJECT_TEXT_WRAP,
        TextWrapModes.NEXT_COLUMN_TEXT_WRAP
    ];

    /* 「なし」オブジェクトスタイルの位置 / Index of the "None" object style */
    var noneObjectStyleIndex = 0;
    for (var i = 0; i < objectStyleNames.length; i++) {
        if (objectStyleNames[i] === NONE_OBJECT_STYLE_NAMES[0] || objectStyleNames[i] === NONE_OBJECT_STYLE_NAMES[1]) {
            noneObjectStyleIndex = i;
            break;
        }
    }

    // ---------------------------------------
    // ダイアログ / Dialog
    // ---------------------------------------
    var dialogTitle = (isInsertionPoint || hasNoSelection)
        ? getLabel("dialog.titleNoSelection")
        : getLabel("dialog.title");

    var createFrameDialog = new Window("dialog", dialogTitle + " " + SCRIPT_VERSION);
    setupWindow(createFrameDialog);

    /* 配置方法パネル / Placement panel */
    var placementPanel = createFrameDialog.add("panel", undefined, getLabel("panel.method"));
    setupPanel(placementPanel, 6);
    placementPanel.alignChildren = ["left", "top"];

    var graphicFrameRadio = placementPanel.add("radiobutton", undefined, getLabel("radio.graphicFrame"));
    var inlineFrameRadio  = placementPanel.add("radiobutton", undefined, getLabel("radio.inlineFrame"));
    if (isInsertionPoint || hasNoSelection) {
        graphicFrameRadio.value = true;
        /* 未選択時は挿入位置がないためインラインを無効化 / Inline needs an insertion point, so disable it when nothing is selected */
        if (hasNoSelection) inlineFrameRadio.enabled = false;
    } else {
        inlineFrameRadio.value = true;
    }

    /* 2カラムレイアウト / Two-column layout */
    var mainColumnsRow = createFrameDialog.add("group");
    setupRow(mainColumnsRow, "fill", COLUMN_SPACING);
    mainColumnsRow.alignChildren = ["fill", "top"];

    var leftColumn = mainColumnsRow.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = ["fill", "top"];
    leftColumn.spacing = PANEL_SPACING;

    var frameSizePanel = leftColumn.add("panel", undefined, getLabel("panel.frameSize"));
    setupPanel(frameSizePanel, 8);

    /* 幅パネル / Width panel */
    var widthPanel = frameSizePanel.add("panel", undefined, getLabel("panel.width"));
    setupPanel(widthPanel, 6);
    widthPanel.alignChildren = ["left", "top"];

    var widthFromTextRadio        = widthPanel.add("radiobutton", undefined, getLabel("radio.widthText"));
    var widthFromColumnRadio      = widthPanel.add("radiobutton", undefined, getLabel("radio.widthColumn"));
    var widthFromParentFrameRadio = widthPanel.add("radiobutton", undefined, getLabel("radio.widthFrame"));
    var widthFromMarginRadio      = widthPanel.add("radiobutton", undefined, getLabel("radio.widthMargin"));

    if (isInsertionPoint || hasNoSelection) {
        widthFromMarginRadio.value = true;
        widthFromTextRadio.enabled = false;
        if (hasNoSelection) {
            widthFromColumnRadio.enabled = false;
            widthFromParentFrameRadio.enabled = false;
        }
    } else {
        widthFromColumnRadio.value = true;
    }

    /* 高さパネル（テキスト外接矩形が使えないときのみ）/ Height panel (only when text bounds are unavailable) */
    var heightInput      = null;
    var heightLinesRadio = null;
    var heightSizeRadio  = null;
    if (!hasTextBounds) {
        var heightPanel = frameSizePanel.add("panel", undefined, getLabel("panel.height"));
        setupPanel(heightPanel, 6);
        heightPanel.alignChildren = ["left", "top"];

        heightLinesRadio = heightPanel.add("radiobutton", undefined, getLabel("radio.heightLines"));
        heightLinesRadio.helpTip = getLabel("tooltip.heightLines");
        heightSizeRadio = heightPanel.add("radiobutton", undefined, getLabel("radio.heightSize"));

        var heightInputRow = heightPanel.add("group");
        setupRow(heightInputRow, "left", 6);

        /* ∧∨と入力欄は隙間0で突き合わせる。下限は行数か mm かで切り替える
           butt the stepper against the field; the minimum follows lines or mm */
        var heightStepperGroup = heightInputRow.add("group");
        heightStepperGroup.orientation = "row";
        heightStepperGroup.alignChildren = ["left", "center"];
        heightStepperGroup.spacing = 0;
        heightStepperGroup.margins = 0;
        var heightStepOptions = { step: 1, min: hasNoSelection ? MIN_HEIGHT_IN_MM : MIN_HEIGHT_IN_LINES };
        var heightStepper = addStepper(heightStepperGroup, function () { return heightInput; }, heightStepOptions);
        heightInput = heightStepperGroup.add("edittext", undefined, hasNoSelection ? DEFAULT_HEIGHT_IN_MM : DEFAULT_HEIGHT_IN_LINES);
        heightInput.characters = HEIGHT_INPUT_CHARACTERS;
        bindSteppedArrowKeys(heightInput, heightStepper);
        if (!hasNoSelection) heightInput.helpTip = getLabel("tooltip.heightLines");

        var heightUnitLabel = heightInputRow.add("statictext", undefined,
            hasNoSelection ? getLabel("unit.mm") : getLabel("unit.lines"));

        /* 未選択時は行送りを参照できないため行数モードを無効化 / Without a selection there is no leading to read, so disable line-count mode */
        if (hasNoSelection) {
            heightSizeRadio.value = true;
            heightLinesRadio.enabled = false;
        } else {
            heightLinesRadio.value = true;
        }

        heightLinesRadio.onClick = function () {
            heightInput.text = DEFAULT_HEIGHT_IN_LINES;
            heightStepOptions.min = MIN_HEIGHT_IN_LINES;
            heightUnitLabel.text = getLabel("unit.lines");
        };
        heightSizeRadio.onClick = function () {
            heightInput.text = DEFAULT_HEIGHT_IN_MM;
            heightStepOptions.min = MIN_HEIGHT_IN_MM;
            heightUnitLabel.text = getLabel("unit.mm");
        };
    }

    /* 右カラム / Right column */
    var rightColumn = mainColumnsRow.add("group");
    rightColumn.orientation = "column";
    rightColumn.alignChildren = ["fill", "top"];
    rightColumn.spacing = PANEL_SPACING;

    var textSettingsPanel = rightColumn.add("panel", undefined, getLabel("panel.textSettings"));
    setupPanel(textSettingsPanel, 6);

    var paraStyleLabel    = textSettingsPanel.add("statictext", undefined, getLabel("field.paraStyle"));
    var paraStyleDropdown = textSettingsPanel.add("dropdownlist", undefined, paragraphStyleNames);
    paraStyleDropdown.selection = defaultParaStyleIndex;

    var autoLeadingCheckbox = textSettingsPanel.add("checkbox", undefined, getLabel("checkbox.autoLeading"));
    autoLeadingCheckbox.value = true;

    var objectSettingsPanel = rightColumn.add("panel", undefined, getLabel("panel.objectSettings"));
    setupPanel(objectSettingsPanel, 6);

    objectSettingsPanel.add("statictext", undefined, getLabel("field.objStyle"));
    var objectStyleDropdown = objectSettingsPanel.add("dropdownlist", undefined, objectStyleNames);
    objectStyleDropdown.selection = noneObjectStyleIndex;

    var wrapLabel    = objectSettingsPanel.add("statictext", undefined, getLabel("field.wrap"));
    var wrapDropdown = objectSettingsPanel.add("dropdownlist", undefined, wrapOptionLabels);
    wrapDropdown.selection = 0;

    /**
     * 現在の選択状態に合わせてコントロールの有効／無効を切り替える
     * @returns {void}
     */
    function updateDialogState() {
        var isInlinePlacement = inlineFrameRadio.value;
        var isNoneObjectStyle = (objectStyleDropdown.selection.index === noneObjectStyleIndex);

        /* グラフィックフレームでは挿入行の段落スタイルを使わない / The inserted-line paragraph style only applies to inline placement */
        paraStyleLabel.enabled      = isInlinePlacement;
        paraStyleDropdown.enabled   = isInlinePlacement;
        autoLeadingCheckbox.enabled = isInlinePlacement;

        var disableWidthFromText        = !hasTextBounds;
        var disableWidthFromColumn      = hasNoSelection;
        var disableWidthFromParentFrame = isInlinePlacement || hasNoSelection;
        var disableWidthFromMargin      = isInlinePlacement;

        widthFromTextRadio.enabled        = !disableWidthFromText;
        widthFromColumnRadio.enabled      = !disableWidthFromColumn;
        widthFromParentFrameRadio.enabled = !disableWidthFromParentFrame;
        widthFromMarginRadio.enabled      = !disableWidthFromMargin;

        /* 選択中の項目が無効になったら、有効な項目へ移す / Move the selection to an enabled option when the current one is disabled */
        if ((widthFromTextRadio.value && disableWidthFromText) ||
            (widthFromColumnRadio.value && disableWidthFromColumn) ||
            (widthFromParentFrameRadio.value && disableWidthFromParentFrame) ||
            (widthFromMarginRadio.value && disableWidthFromMargin)) {
            if (!disableWidthFromText) {
                widthFromTextRadio.value = true;
            } else if (!disableWidthFromColumn) {
                widthFromColumnRadio.value = true;
            } else if (!disableWidthFromParentFrame) {
                widthFromParentFrameRadio.value = true;
            } else if (!disableWidthFromMargin) {
                widthFromMarginRadio.value = true;
            }
        }

        /* 回り込みはページ配置かつオブジェクトスタイルが「なし」のときだけ有効 / Text wrap applies only to page placement with the "None" object style */
        wrapLabel.enabled    = !isInlinePlacement && isNoneObjectStyle;
        wrapDropdown.enabled = !isInlinePlacement && isNoneObjectStyle;
    }

    updateDialogState();
    graphicFrameRadio.onClick     = updateDialogState;
    inlineFrameRadio.onClick      = updateDialogState;
    objectStyleDropdown.onChange  = updateDialogState;

    var buttonRow = addButtonRow(createFrameDialog);
    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    alignRightOnlyButtonRow(buttonRow);

    if (createFrameDialog.show() !== 1) return;

    // ---------------------------------------
    // 入力値の確定 / Resolve settings
    // ---------------------------------------
    var useInlinePlacement    = inlineFrameRadio.value;
    var selectedObjStyleIndex = objectStyleDropdown.selection.index;
    var selectedWrapMode      = wrapOptionModes[wrapDropdown.selection.index];
    var selectedParaStyleIdx  = paraStyleDropdown.selection.index;
    var useAutoLeading        = autoLeadingCheckbox.value;

    /* フレームの高さ / Frame height */
    var frameHeight = selectionHeight;
    if (heightLinesRadio && heightSizeRadio) {
        if (heightLinesRadio.value) {
            /* 行数モード: 1行目の文字サイズ + 残り行数 × 行送り / Line-count mode: first-line size plus leading for the remaining lines */
            var lineCount = parseFloat(heightInput.text) || 1;
            if (lineCount < 1) lineCount = 1;
            try {
                var referenceFrame = isInsertionPoint
                    ? getInsertionPointTextFrame(selectedText)
                    : selectedText.parentTextFrames[0];
                if (!referenceFrame) throw new Error("No parent text frame");

                var referenceStory = referenceFrame.parentStory;
                var insertionIndex = isInsertionPoint ? selectedText.index : selectedText.characters[0].index;
                var referenceCharacter = (insertionIndex < referenceStory.characters.length)
                    ? referenceStory.characters[insertionIndex]
                    : referenceStory.characters[referenceStory.characters.length - 1];

                var referencePointSize = referenceCharacter.pointSize;
                var referenceLeading   = referenceCharacter.leading;
                if (referenceLeading === Leading.AUTO) {
                    referenceLeading = referencePointSize * (referenceCharacter.autoLeading / 100);
                }
                frameHeight = referencePointSize + Math.max(0, lineCount - 1) * referenceLeading;
            } catch (e) {
                frameHeight = FALLBACK_POINT_SIZE + Math.max(0, lineCount - 1) * FALLBACK_LEADING;
            }
        } else {
            /* サイズ指定モード: mm 入力をポイントへ変換 / Size mode: convert the millimeter input to points */
            var heightInMillimeters = parseFloat(heightInput.text) || parseFloat(DEFAULT_HEIGHT_IN_MM);
            frameHeight = heightInMillimeters * MM_TO_POINTS;
        }
    }

    /* フレームの幅 / Frame width */
    var frameWidth = selectionWidth;
    if (widthFromColumnRadio.value) {
        try {
            var columnSourceFrame = selectedText.parentTextFrames[0];
            var columnWidth = columnSourceFrame.textFramePreferences.textColumnFixedWidth;
            if (columnWidth <= 0) {
                var columnFrameBounds = columnSourceFrame.geometricBounds;
                var textFrameWidth = columnFrameBounds[3] - columnFrameBounds[1];
                var columnCount    = columnSourceFrame.textFramePreferences.textColumnCount;
                var columnGutter   = columnSourceFrame.textFramePreferences.textColumnGutter;
                var insetLeft      = columnSourceFrame.textFramePreferences.insetSpacing[1];
                var insetRight     = columnSourceFrame.textFramePreferences.insetSpacing[3];
                columnWidth = (textFrameWidth - insetLeft - insetRight - columnGutter * (columnCount - 1)) / columnCount;
            }
            frameWidth = columnWidth;
        } catch (e) {}
    } else if (widthFromParentFrameRadio.value) {
        try {
            var parentFrameBounds = selectedText.parentTextFrames[0].geometricBounds;
            frameWidth = parentFrameBounds[3] - parentFrameBounds[1];
        } catch (e) {}
    } else if (widthFromMarginRadio.value) {
        try {
            var marginPage;
            if (hasNoSelection) {
                marginPage = app.activeWindow.activePage;
            } else if (isInsertionPoint) {
                var insertionFrame = getInsertionPointTextFrame(selectedText);
                marginPage = insertionFrame ? insertionFrame.parentPage : null;
            } else {
                marginPage = getParentPage(selectedText);
            }
            if (marginPage) {
                var marginPageBounds = marginPage.bounds;
                var pageMargins = marginPage.marginPreferences;
                frameWidth = (marginPageBounds[3] - marginPageBounds[1]) - pageMargins.left - pageMargins.right;
            }
        } catch (e) {}
    }

    // ---------------------------------------
    // フレーム作成 / Frame creation
    // ---------------------------------------

    /**
     * インライン（アンカー付き）でフレームを挿入する
     * @returns {void}
     */
    function createInlineFrame() {
        var anchorFrame = getInsertionPointTextFrame(selectedText);
        if (!anchorFrame) {
            alert(getLabel("alert.parentFrameError"));
            return;
        }

        var anchorStory = anchorFrame.parentStory;
        var charIndex = isInsertionPoint ? selectedText.index : selectedText.characters[0].index;

        /* 元のテキストの段落スタイルを控える / Remember the original paragraph style */
        var originalParaStyle = null;
        var originalParagraph = null;
        if (!isInsertionPoint) {
            originalParaStyle = selectedText.characters[0].appliedParagraphStyle;
            originalParagraph = selectedText.characters[0].paragraphs[0];
        }

        /* 直前が改行でなければ改行を挿入 / Insert a return unless the previous character already is one */
        var needsLeadingReturn = (charIndex > 0) && (anchorStory.characters[charIndex - 1].contents !== "\r");
        if (needsLeadingReturn) {
            anchorStory.insertionPoints[charIndex].contents = "\r";
            charIndex = charIndex + 1;
        }

        var anchorInsertionPoint = anchorStory.insertionPoints[charIndex];
        var anchoredRectangle = anchorInsertionPoint.rectangles.add();
        anchoredRectangle.geometricBounds = [0, 0, frameHeight, frameWidth];
        anchoredRectangle.contentType = ContentType.GRAPHIC_TYPE;
        anchoredRectangle.anchoredObjectSettings.anchoredPosition = AnchorPosition.ANCHORED;
        anchoredRectangle.appliedObjectStyle = activeDoc.objectStyles[selectedObjStyleIndex];

        /* フレーム挿入で文字が 1 つ増えるため、その次の位置に改行を入れる / The inline frame adds one character, so the return goes after it */
        anchorStory.insertionPoints[charIndex + 1].contents = "\r";

        var anchorParagraph = anchorInsertionPoint.paragraphs[0];
        anchorParagraph.appliedParagraphStyle = activeDoc.paragraphStyles[selectedParaStyleIdx];

        if (useAutoLeading) {
            anchorParagraph.autoLeading = 100;
            anchorParagraph.leading = Leading.AUTO;
        }

        if (originalParagraph && originalParaStyle) {
            try {
                originalParagraph.appliedParagraphStyle = originalParaStyle;
            } catch (e) {}
        }
    }

    /**
     * ページ上にグラフィックフレームを作成する
     * @returns {void}
     */
    function createGraphicFrameOnPage() {
        var destinationPage;
        if (hasNoSelection) {
            destinationPage = app.activeWindow.activePage;
        } else if (isInsertionPoint) {
            var insertionFrame = getInsertionPointTextFrame(selectedText);
            destinationPage = insertionFrame ? insertionFrame.parentPage : null;
        } else {
            destinationPage = getParentPage(selectedText);
        }

        if (!destinationPage) {
            alert(getLabel("alert.parentPageError"));
            return;
        }

        var frameTop, frameLeft;
        if (hasTextBounds) {
            frameTop  = selectionBounds[0];
            frameLeft = selectionBounds[1];
        } else {
            /* 挿入ポイント・未選択時はマージン左上を基準にする / Anchor to the top-left margin for insertion points and empty selections */
            var pageMargins = destinationPage.marginPreferences;
            var destinationBounds = destinationPage.bounds;
            frameTop  = destinationBounds[0] + pageMargins.top;
            frameLeft = destinationBounds[1] + pageMargins.left;
        }

        var graphicRectangle = destinationPage.rectangles.add();
        graphicRectangle.geometricBounds = [frameTop, frameLeft, frameTop + frameHeight, frameLeft + frameWidth];
        graphicRectangle.contentType = ContentType.GRAPHIC_TYPE;
        graphicRectangle.appliedObjectStyle = activeDoc.objectStyles[selectedObjStyleIndex];

        /* 回り込みはオブジェクトスタイルが「なし」のときだけ設定 / Apply text wrap only with the "None" object style */
        if (selectedObjStyleIndex === noneObjectStyleIndex) {
            graphicRectangle.textWrapPreferences.textWrapMode = selectedWrapMode;
        }
    }

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(function () {
        if (useInlinePlacement) {
            createInlineFrame();
        } else {
            createGraphicFrameOnPage();
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.createFrame"));

})();
