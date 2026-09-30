#target indesign

/*

### 概要

選択したセルや表の列幅を、列ごとの個別指定または一括入力でまとめて調整します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableColumnWidthAdjuster.md

### Overview

Adjusts the column widths of the selected cells or table, either per column or with a single value applied to all.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableColumnWidthAdjuster.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTableColumnWidthAdjuster";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableColumnWidthAdjuster.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableColumnWidthAdjuster.md"; /* README (English) */

// Original idea
// AdjColWidth_221003a.jsx by 照山裕爾（mottainaiDTP）
// https://mottainaidtp.seesaa.net/article/492096133.html

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


(function () {

	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;

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

	/**
	 * 行に「∧∨＋入力欄」を隙間0で突き合わせる group を足し、∧∨と入力欄を置く
	 * @param {Group} parentRow 追加先の行
	 * @param {string} initialText 入力欄の初期値
	 * @param {object} stepOptions addStepper() に渡す増減の設定
	 * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できる）
	 */
	function addSteppedInput(parentRow, initialText, stepOptions) {
		var pairGroup = parentRow.add("group");
		pairGroup.orientation = "row";
		pairGroup.alignChildren = ["left", "center"];
		pairGroup.spacing = 0; /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
		pairGroup.margins = 0;
		var numberInput;
		var stepperGroup = addStepper(pairGroup, function () { return numberInput; }, stepOptions);
		numberInput = pairGroup.add("edittext", undefined, initialText);
		numberInput.stepperGroup = stepperGroup;
		bindSteppedArrowKeys(numberInput, stepperGroup);
		return numberInput;
	}

	/**
	 * ∧∨・↑↓キーで増減したあと、手入力と同じ onChanging（連動欄の更新）と onChange（検証・プレビュー）を通す
	 * @param {EditText} numberInput 増減した入力欄
	 * @returns {void}
	 */
	function runInputHandlersAfterStep(numberInput) {
		if (typeof numberInput.onChanging === "function") numberInput.onChanging();
		if (typeof numberInput.onChange === "function") numberInput.onChange();
	}

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
			title: { ja: "列幅の調整", en: "Adjust Column Widths" }
		},
		panel: {
			basic:        { ja: "指定方法", en: "Sizing Method" },
			perColumn:    { ja: "個別に設定", en: "Per-Column Settings" },
			columnWidths: { ja: "各列の設定", en: "Column Settings" },
			batch:        { ja: "一括入力", en: "Batch Input" }
		},
		field: {
			calculationBasis: { ja: "指定方法", en: "Sizing Method" },
			inputMethod:      { ja: "入力方法", en: "Input Method" }
		},
		radio: {
			inputMethodPerColumn: { ja: "個別に設定", en: "Per-Column Input" },
			inputMethodBatch:     { ja: "一括入力", en: "Batch Input" },
			modeAbsolute:         { ja: "幅で指定", en: "Set by Width" },
			modeCharacterBased:   { ja: "文字数で指定", en: "Set by Character Count" }
		},
		checkbox: {
			unify:   { ja: "全列に適用", en: "Apply to All Columns" },
			preview: { ja: "プレビュー", en: "Preview" }
		},
		header: {
			column:    { ja: "列", en: "Col" },
			width:     { ja: "幅", en: "Width" },
			charCount: { ja: "文字数", en: "Character Count" },
			inset:     { ja: "左右の余白", en: "Left/Right Inset" },
			autoFit:   { ja: "自動調整", en: "Auto Fit" }
		},
		unit: {
			mm:         { ja: "mm", en: "mm" },
			characters: { ja: "文字", en: "chars" },
			columnSuffix: { ja: "列目", en: "Col" }
		},
		button: {
			ok:     { ja: "OK", en: "OK" },
			cancel: { ja: "キャンセル", en: "Cancel" },
			apply:  { ja: "適用", en: "Apply" },
			screenModePreview: { ja: "プレビュー", en: "Preview" },
			screenModeNormal:  { ja: "標準モード", en: "Normal Mode" }
		},
		hint: {
			batch: {
				ja: "入力形式：15 20 34 10 または 15, 20, 34, 10",
				en: "Example: Enter column widths as 15 20 34 10 or 15, 20, 34, 10"
			}
		},
		alert: {
			selectCellOrTable:   { ja: "セルまたは表を選択してから実行してください", en: "Select a cell or table before running this script." },
			batchInvalidValue:   { ja: "一括入力に無効な値が含まれています", en: "Batch input contains invalid values." },
			batchTooManyValues:  { ja: "入力数が列数を超えています", en: "Too many values for the number of columns." },
			invalidNumber:       { ja: "数値を入力してください", en: "Enter a valid number." },
			negativeWidth:       { ja: "幅には 0 以上の数値を入力してください", en: "Width must be 0 or greater." },
			negativeCharCount:   { ja: "文字数には 0 以上の数値を入力してください", en: "Character count must be 0 or greater." },
			negativeInset:       { ja: "左右の余白には 0 以上の数値を入力してください", en: "Left/right inset must be 0 or greater." },
			insetTooLarge:       { ja: "左右の余白が大きすぎます。内容幅が 0 以下になります", en: "Left/right inset is too large. Content width would become 0 or less." }
		},
		tooltip: {
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
			screenMode: {
				ja: "ドキュメントウィンドウの画面モードを、標準モードとプレビューで切り替えます。",
				en: "Switches the document window between the Normal and Preview screen modes."
			}
		},
		undo: {
			adjustColumnWidths: { ja: "列幅の調整", en: "Adjust Column Widths" }
		}
	};

	/* 実行とエラー対策 / Execution and error handling */
	main();
	/**
	 * 列幅調整の処理を開始する
	 * @returns {void}
	 */
	function main() {
		/* 文字列で渡すとグローバルスコープで評価され IIFE 内の関数を見つけられないため関数で渡す
		   / Pass a function: a string would be evaluated in the global scope and could not see functions inside the IIFE */
		app.doScript(adjustColumnWidths, ScriptLanguage.JAVASCRIPT, undefined,
			UndoModes.ENTIRE_SCRIPT, getLabel("undo.adjustColumnWidths"));
	}

	/**
	 * ダイアログを表示して列幅の調整を実行する
	 * @returns {void}
	 */
	function adjustColumnWidths() {
		var selection = app.activeDocument.selection;
		var targetTable = getTableFromSelection(selection[0]);
		if (!targetTable) {
			alert(getLabel("alert.selectCellOrTable"));
			return;
		}

		// 選択を記憶し、ダイアログ中はハイライトを消す / Save selection and hide highlight while dialog is open
		var savedSelection = [];
		for (var selectionIndex = 0; selectionIndex < selection.length; selectionIndex++) {
			savedSelection.push(selection[selectionIndex]);
		}
		try {
			app.select(NothingEnum.NOTHING);
		} catch (e) { }

		var columnCount = targetTable.columns.length;
		var originalWidths = getColumnWidths(targetTable);
		var originalInsets = getOriginalInsets(targetTable);

		var dominantFontSizePt = findDominantFontSize(targetTable);
		fontSizePtForValidationCache = dominantFontSizePt;

		var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
		setupWindow(dialog, 10);

		var inputMethodOuterGroup = dialog.add("group");
		inputMethodOuterGroup.orientation = "row";
		inputMethodOuterGroup.alignment = "center";
		inputMethodOuterGroup.alignChildren = ["center", "center"];
		inputMethodOuterGroup.margins = [0, 3, 0, 10];

		var inputMethodGroup = inputMethodOuterGroup.add("group");
		setupRow(inputMethodGroup, "left", COLUMN_SPACING);
		var perColumnRadio = inputMethodGroup.add("radiobutton", undefined, getLabel("radio.inputMethodPerColumn"));
		var batchModeRadio = inputMethodGroup.add("radiobutton", undefined, getLabel("radio.inputMethodBatch"));
		perColumnRadio.value = true;

		var panelUnitLabel = getRulerUnitString();
		var perColumnPanel = dialog.add("panel", undefined, getLabel("panel.perColumn"));
		setupPanel(perColumnPanel, 10);

		var basicSettingsPanel = perColumnPanel.add("panel", undefined, getLabel("panel.basic"));
		setupPanel(basicSettingsPanel, 6);
		basicSettingsPanel.alignChildren = "left";

		var modeGroup = basicSettingsPanel.add("group");
		modeGroup.orientation = "column";
		modeGroup.alignChildren = "left";
		modeGroup.spacing = 4;
		var absoluteRadio = modeGroup.add("radiobutton", undefined, getLabel("radio.modeAbsolute"));
		var characterBasedRadio = modeGroup.add("radiobutton", undefined, getLabel("radio.modeCharacterBased"));
		absoluteRadio.value = true;

		var columnSettingsPanel = perColumnPanel.add(
			"panel",
			undefined,
			getLabel("panel.columnWidths") + (uiLang === "ja" ? "（" + panelUnitLabel + "）" : " (" + panelUnitLabel + ")")
		);
		setupPanel(columnSettingsPanel, 6);
		columnSettingsPanel.alignChildren = "left";
		// 全列に適用 / Apply to all columns
		var unifyCheckbox = columnSettingsPanel.add("checkbox", undefined, getLabel("checkbox.unify"));

		var initialInsetValuesUi = insetsToInitialInsetValues(originalInsets);
		var columnSettingsControls = buildColumnSettingsControls(columnSettingsPanel, columnCount, originalWidths, initialInsetValuesUi, dominantFontSizePt);
		var widthInputs = columnSettingsControls.widthInputs;
		var charCountInputs = columnSettingsControls.charCountInputs;
		var sideInsetInputs = columnSettingsControls.sideInsetInputs;
		var autoFitCheckboxes = columnSettingsControls.autoFitCheckboxes;

		var columnStates = [];

		for (var i = 0; i < columnCount; i++) {
			columnStates.push({
				mode: "manual", // "manual" or "autofit"
				width: rulerValueToInputUnit(originalWidths[i]),
				inset: initialInsetValuesUi[i],
				lockedWidth: null
			});
		}

		var rowLabels = columnSettingsControls.rowLabels;
		var headerLabels = columnSettingsControls.headerLabels;
		// Fallback normalizations for missing/malformed rowLabels/headerLabels
		if (!rowLabels || !(rowLabels instanceof Array)) rowLabels = [];
		if (!headerLabels || !(headerLabels instanceof Array)) headerLabels = [];

		var batchInputPanel = dialog.add("panel", undefined, getLabel("panel.batch"));
		setupPanel(batchInputPanel, 6);

		var batchRow = batchInputPanel.add("group");
		batchRow.orientation = "row";
		batchRow.alignment = "fill";
		batchRow.alignChildren = ["fill", "center"];
		batchRow.spacing = 8;

		var batchLeftGroup = batchRow.add("group");
		batchLeftGroup.orientation = "column";
		batchLeftGroup.alignment = ["fill", "center"];
		batchLeftGroup.alignChildren = ["fill", "center"];

		var batchInput = batchLeftGroup.add("edittext", undefined, "");
		batchInput.alignment = ["fill", "center"];

		var batchRightGroup = batchRow.add("group");
		batchRightGroup.orientation = "column";
		batchRightGroup.alignment = ["right", "center"];
		batchRightGroup.alignChildren = ["right", "center"];

		var btnBatchApply = batchRightGroup.add("button", undefined, getLabel("button.apply"));

		var batchHintText = batchInputPanel.add("statictext", undefined, getLabel("hint.batch"));

		var buttonRow = addButtonRow(dialog);
		addScreenModeButton(buttonRow.leftGroup);
		var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
		var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
		centerButtonRowIfRightOnly(buttonRow);

		var primaryInputMode = "absolute";
		var currentInputMethod = "perColumn";
		var isPreviewCurrentlyApplied = false;

		/**
		 * 指定方法（幅／文字数）を切り替える
		 * @param {string} mode "absolute" または "character"
		 * @returns {void}
		 */
		function setInputMode(mode) {
			primaryInputMode = mode;
			refreshControlStates();
		}

		/**
		 * 入力方法（個別／一括）を切り替える
		 * @param {string} method "perColumn" または "batch"
		 * @returns {void}
		 */
		function setInputMethod(method) {
			currentInputMethod = method;
			refreshControlStates();
		}

		/**
		 * 現在のモードに応じてコントロールの有効／無効を更新する
		 * @returns {void}
		 */
		function refreshControlStates() {
			var isPerColumnMode = (currentInputMethod == "perColumn");
			columnSettingsPanel.visible = true;
			batchInputPanel.visible = true;

			setControlEnabled(batchInput, !isPerColumnMode);
			setControlEnabled(btnBatchApply, !isPerColumnMode);
			setControlEnabled(batchHintText, !isPerColumnMode);

			setControlEnabled(absoluteRadio, isPerColumnMode);
			setControlEnabled(characterBasedRadio, isPerColumnMode);
			setControlEnabled(unifyCheckbox, isPerColumnMode);

			for (var h = 0; h < headerLabels.length; h++) {
				if (headerLabels[h]) setLabelDimmed(headerLabels[h], !isPerColumnMode);
			}

			for (var i = 0; i < columnCount; i++) {
				var isAuto = autoFitCheckboxes[i].value;
				var isNonFirstUnified = unifyCheckbox.value && i > 0 && !isAuto;
				var enableWidth = isPerColumnMode && !((primaryInputMode == "characterBased") || isNonFirstUnified);
				var enableChar = isPerColumnMode && !((primaryInputMode == "absolute") || isNonFirstUnified);
				var enableInset = isPerColumnMode && !isNonFirstUnified;

				setControlEnabled(widthInputs[i], enableWidth);
				setControlEnabled(charCountInputs[i], enableChar);
				setControlEnabled(sideInsetInputs[i], enableInset);
				setControlEnabled(autoFitCheckboxes[i], isPerColumnMode);

				var currentRowLabels = (rowLabels[i] && rowLabels[i] instanceof Array) ? rowLabels[i] : [];
				for (var j = 0; j < currentRowLabels.length; j++) {
					if (currentRowLabels[j]) setLabelDimmed(currentRowLabels[j], !isPerColumnMode || isNonFirstUnified);
				}
			}
			// 個別設定パネル全体の見た目を切り替え / Update the appearance of the entire per-column settings panel
			try {
				var panelDimRgb = isPerColumnMode ? [0, 0, 0] : [0.55, 0.55, 0.55];
				perColumnPanel.graphics.foregroundColor = perColumnPanel.graphics.newPen(perColumnPanel.graphics.PenType.SOLID_COLOR, panelDimRgb, 1);
				basicSettingsPanel.graphics.foregroundColor = basicSettingsPanel.graphics.newPen(basicSettingsPanel.graphics.PenType.SOLID_COLOR, panelDimRgb, 1);
				columnSettingsPanel.graphics.foregroundColor = columnSettingsPanel.graphics.newPen(columnSettingsPanel.graphics.PenType.SOLID_COLOR, panelDimRgb, 1);
			} catch (e) { }

			// 一括入力パネル全体の見た目を切り替え / Update the appearance of the entire batch input panel
			try {
				var batchDimRgb = isPerColumnMode ? [0.55, 0.55, 0.55] : [0, 0, 0];
				batchInputPanel.graphics.foregroundColor = batchInputPanel.graphics.newPen(batchInputPanel.graphics.PenType.SOLID_COLOR, batchDimRgb, 1);
				batchRow.graphics.foregroundColor = batchRow.graphics.newPen(batchRow.graphics.PenType.SOLID_COLOR, batchDimRgb, 1);
				batchLeftGroup.graphics.foregroundColor = batchLeftGroup.graphics.newPen(batchLeftGroup.graphics.PenType.SOLID_COLOR, batchDimRgb, 1);
				batchRightGroup.graphics.foregroundColor = batchRightGroup.graphics.newPen(batchRightGroup.graphics.PenType.SOLID_COLOR, batchDimRgb, 1);
				batchHintText.graphics.foregroundColor = batchHintText.graphics.newPen(batchHintText.graphics.PenType.SOLID_COLOR, batchDimRgb, 1);
			} catch (e) { }
		}

		/**
		 * 入力値を列幅と余白へ反映する
		 * @returns {void}
		 */
		function applyColumnSettings() {
			// columnStates を唯一の truth として適用する
			// Apply everything from columnStates as the single source of truth
			for (var i = 0; i < columnStates.length; i++) {
				try {
					var state = columnStates[i];
					var widthToApply = null;

					if (state.mode === "autofit") {
						widthToApply = state.lockedWidth;
					} else {
						widthToApply = state.width;
					}

					if (widthToApply != null && !isNaN(widthToApply)) {
						targetTable.columns[i].width = inputUnitToRulerValue(widthToApply);
					}
				} catch (e) { }
			}

			applyColumnInsetsFromStates(targetTable, columnStates);
			isPreviewCurrentlyApplied = true;
		}

		/**
		 * 列幅と余白を実行前の状態に戻す
		 * @returns {void}
		 */
		function restoreOriginalColumnSettings() {
			restoreColumnWidths(targetTable, originalWidths);
			restoreColumnInsets(targetTable, originalInsets);
			isPreviewCurrentlyApplied = false;
		}

		// 高速自動調整: 文字数推定による列内容幅計算
		/**
		 * 文字数を基準に列の内容幅を見積もる
		 * @param {number} colIdx 列の位置
		 * @param {number} fontSizePt 基準の文字サイズ（pt）
		 * @returns {number} 見積もった内容幅
		 */
		function estimateColumnContentWidthByChars(colIdx, fontSizePt) {
			var col = targetTable.columns[colIdx];
			var cells = col.cells;
			var maxChars = 0;

			for (var i = 0; i < cells.length; i++) {
				var cell = cells[i];
				if (!cell.texts || cell.texts.length === 0) continue;
				try {
					var lines = cell.texts[0].lines;
					for (var j = 0; j < lines.length; j++) {
						var len = lines[j].characters.length;
						if (len > maxChars) maxChars = len;
					}
				} catch (e) { }
			}

			return ptToInputUnit(maxChars * fontSizePt);
		}

		/**
		 * 列の内容が収まる幅を実測する
		 * @param {number} colIdx 列の位置
		 * @returns {number} 実測した内容幅
		 */
		function measureColumnContentWidth(colIdx) {
			var table = targetTable;
			var col = table.columns[colIdx];
			var cells = col.cells;
			var savedWidths = [];
			for (var i = 0; i < table.columns.length; i++) {
				try {
					savedWidths.push(table.columns[i].width);
				} catch (e) {
					savedWidths.push(null);
				}
			}

			/**
			 * 列内のいずれかのセルに 2 行目があるかを判定する
			 * @returns {boolean} 2 行目があれば true
			 */
			function hasSecondLineInAnyCell() {
				for (var k = 0; k < cells.length; k++) {
					var cell = cells[k];
					if (!cell.texts || cell.texts.length === 0) continue;
					try {
						if (cell.texts[0].lines.length >= 2) return true;
					} catch (e) { }
				}
				return false;
			}

			// 2行目がない場合は現在幅をそのまま使わず、内容推定に切り替える
			if (!hasSecondLineInAnyCell()) {
				return estimateColumnContentWidthByChars(colIdx, dominantFontSizePt);
			}

			var measuredWidth = col.width;
			var coarseStep = inputUnitToRulerValue(10);
			if (coarseStep <= 0 || isNaN(coarseStep)) coarseStep = 10;
			var fineStep = inputUnitToRulerValue(1);
			if (fineStep <= 0 || isNaN(fineStep)) fineStep = 1;
			var maxIterations = 200;
			var count = 0;

			try {
				// 1) 粗く広げて、2行目が消える幅まで到達する
				while (hasSecondLineInAnyCell() && count < maxIterations) {
					measuredWidth += coarseStep;
					try { col.width = measuredWidth; } catch (e) { break; }
					count++;
				}

				// 2) 少し戻して、細かい刻みで詰める
				var refineStart = measuredWidth - coarseStep;
				if (refineStart < 0) refineStart = 0;
				try { col.width = refineStart; } catch (e) { }
				measuredWidth = refineStart;

				count = 0;
				while (hasSecondLineInAnyCell() && count < maxIterations) {
					measuredWidth += fineStep;
					try { col.width = measuredWidth; } catch (e) { break; }
					count++;
				}
			} finally {
				for (var j = 0; j < savedWidths.length; j++) {
					if (savedWidths[j] == null) continue;
					try { table.columns[j].width = savedWidths[j]; } catch (e) { }
				}
			}

			return rulerValueToInputUnit(measuredWidth);
		}

		/**
		 * 指定した列に自動調整を適用する
		 * @param {number} colIdx 列の位置
		 * @returns {void}
		 */
		function applyAutoFitToColumn(colIdx) {
			var contentW = measureColumnContentWidth(colIdx);
			var inset = parseFloat(sideInsetInputs[colIdx].text);
			if (isNaN(inset)) inset = 0;
			var newWidth = contentW + 2 * inset;
			var widthText = formatNumber(newWidth);
			var cc = calculateCharCount(newWidth, inset, dominantFontSizePt);
			var charText = formatNumber(cc);

			columnStates[colIdx].mode = "autofit";
			columnStates[colIdx].lockedWidth = newWidth;
			columnStates[colIdx].width = newWidth;
			columnStates[colIdx].inset = inset;
			widthInputs[colIdx].text = widthText;
			charCountInputs[colIdx].text = charText;

			// ScriptUI の描画更新を強める / Force ScriptUI to repaint updated edittexts
			try { widthInputs[colIdx].text = ""; widthInputs[colIdx].text = widthText; } catch (e) { }
			try { charCountInputs[colIdx].text = ""; charCountInputs[colIdx].text = charText; } catch (e) { }
			try { widthInputs[colIdx].parent.parent.layout.layout(true); } catch (e) { }
			try { dialog.layout.layout(true); } catch (e) { }
			try { dialog.update(); } catch (e) { }

			widthInputs[colIdx].text = formatNumber(columnStates[colIdx].width);
		}

		absoluteRadio.onClick = function () {
			if (absoluteRadio.value) setInputMode("absolute");
		};
		characterBasedRadio.onClick = function () {
			if (characterBasedRadio.value) setInputMode("characterBased");
		};

		perColumnRadio.onClick = function () {
			if (perColumnRadio.value) setInputMethod("perColumn");
		};
		batchModeRadio.onClick = function () {
			if (batchModeRadio.value) setInputMethod("batch");
		};

		unifyCheckbox.onClick = function () {
			// 全列に適用を ON にしたときは、自動調整を全列で解除する
			// When Apply to All Columns is turned on, disable auto-fit for all columns
			if (unifyCheckbox.value) {
				for (var i = 0; i < autoFitCheckboxes.length; i++) {
					autoFitCheckboxes[i].value = false;
				}
			}
			// ここでは全列を即時上書きしない / Do not overwrite all columns immediately here
			// 実際の同期は編集中に行う / Actual synchronization happens during editing
			refreshControlStates();
		};

		for (var i = 0; i < columnCount; i++) {
			(function (idx) {
				widthInputs[idx].onChanging = function () {
					updateCharCountFromWidth(widthInputs[idx], charCountInputs[idx], sideInsetInputs[idx], dominantFontSizePt);
					if (unifyCheckbox.value) syncUnifiedInputsFromSource(widthInputs, charCountInputs, sideInsetInputs, idx, dominantFontSizePt, "absolute", autoFitCheckboxes, columnStates);
				};
				widthInputs[idx].onChange = function () {
					var validation = validatePerColumnRow(widthInputs[idx], charCountInputs[idx], sideInsetInputs[idx], "absolute");
					if (!validation.ok) {
						alert(validation.message);
						try { validation.focus.active = true; } catch (e) { }
						updateCharCountFromWidth(widthInputs[idx], charCountInputs[idx], sideInsetInputs[idx], dominantFontSizePt);
						return;
					}
					var manualWidth = parseFloat(widthInputs[idx].text);
					var manualInset = parseFloat(sideInsetInputs[idx].text);
					columnStates[idx].mode = "manual";
					columnStates[idx].lockedWidth = null;
					columnStates[idx].width = isNaN(manualWidth) ? columnStates[idx].width : manualWidth;
					if (!isNaN(manualInset)) columnStates[idx].inset = manualInset;
					if (autoFitCheckboxes[idx].value) {
						autoFitCheckboxes[idx].value = false;
						refreshControlStates();
					}
					applyColumnSettings();
				};
				charCountInputs[idx].onChanging = function () {
					updateWidthFromCharCount(widthInputs[idx], charCountInputs[idx], sideInsetInputs[idx], dominantFontSizePt);
					if (unifyCheckbox.value) syncUnifiedInputsFromSource(widthInputs, charCountInputs, sideInsetInputs, idx, dominantFontSizePt, "characterBased", autoFitCheckboxes, columnStates);
				};
				charCountInputs[idx].onChange = function () {
					var validation = validatePerColumnRow(widthInputs[idx], charCountInputs[idx], sideInsetInputs[idx], "characterBased");
					if (!validation.ok) {
						alert(validation.message);
						try { validation.focus.active = true; } catch (e) { }
						updateWidthFromCharCount(widthInputs[idx], charCountInputs[idx], sideInsetInputs[idx], dominantFontSizePt);
						return;
					}
					var manualWidth = parseFloat(widthInputs[idx].text);
					var manualInset = parseFloat(sideInsetInputs[idx].text);
					columnStates[idx].mode = "manual";
					columnStates[idx].lockedWidth = null;
					columnStates[idx].width = isNaN(manualWidth) ? columnStates[idx].width : manualWidth;
					if (!isNaN(manualInset)) columnStates[idx].inset = manualInset;
					if (autoFitCheckboxes[idx].value) {
						autoFitCheckboxes[idx].value = false;
						refreshControlStates();
					}
					applyColumnSettings();
				};
				sideInsetInputs[idx].onChanging = function () {
					if (autoFitCheckboxes[idx].value) {
						applyAutoFitToColumn(idx);
					}
					else {
						syncWidthAndCharCountFromInset(widthInputs[idx], charCountInputs[idx], sideInsetInputs[idx], dominantFontSizePt, primaryInputMode);
					}
					if (unifyCheckbox.value) syncUnifiedInputsFromSource(widthInputs, charCountInputs, sideInsetInputs, idx, dominantFontSizePt, primaryInputMode, autoFitCheckboxes, columnStates);
				};
				sideInsetInputs[idx].onChange = function () {
					var validation = validatePerColumnRow(widthInputs[idx], charCountInputs[idx], sideInsetInputs[idx], primaryInputMode);
					if (!validation.ok) {
						alert(validation.message);
						try { validation.focus.active = true; } catch (e) { }
						if (autoFitCheckboxes[idx].value) {
							applyAutoFitToColumn(idx);
						}
						else {
							syncWidthAndCharCountFromInset(widthInputs[idx], charCountInputs[idx], sideInsetInputs[idx], dominantFontSizePt, primaryInputMode);
						}
						return;
					}

					var updatedInset = parseFloat(sideInsetInputs[idx].text);
					if (!isNaN(updatedInset)) columnStates[idx].inset = updatedInset;

					if (autoFitCheckboxes[idx].value) {
						applyAutoFitToColumn(idx);
					}
					else {
						var manualWidth = parseFloat(widthInputs[idx].text);
						columnStates[idx].mode = "manual";
						columnStates[idx].lockedWidth = null;
						columnStates[idx].width = isNaN(manualWidth) ? columnStates[idx].width : manualWidth;
					}

					applyColumnSettings();
				};
				/**
				 * 自動調整チェックボックスの切り替えを処理する
				 * @returns {void}
				 */
				function handleAutoFitToggle() {
					var isAutoFitOn = !!autoFitCheckboxes[idx].value;
					if (isAutoFitOn) {
						applyAutoFitToColumn(idx);
					}
					refreshControlStates();
					applyColumnSettings();
				}
				autoFitCheckboxes[idx].onClick = handleAutoFitToggle;
				autoFitCheckboxes[idx].onChange = handleAutoFitToggle;
			})(i);
		}

		/**
		 * 一括入力の値を各列へ反映する
		 * @returns {void}
		 */
		function applyBatchInput() {
			var values = parseBatchInput(batchInput.text);
			if (values.length == 0) return;

			// 入力検証：無効な値があれば中断 / Validation: reject if any invalid value exists
			for (var i = 0; i < values.length; i++) {
				if (values[i] == null) {
					alert(getLabel("alert.batchInvalidValue"));
					return;
				}
			}

			// 入力検証：列数を超える場合は中断 / Validation: reject if too many values
			if (values.length > columnCount) {
				alert(getLabel("alert.batchTooManyValues"));
				return;
			}

			// 入力された列数ぶんだけ反映 / Apply only to the provided columns
			for (var i = 0; i < values.length; i++) {
				widthInputs[i].text = formatNumber(values[i]);
				updateCharCountFromWidth(widthInputs[i], charCountInputs[i], sideInsetInputs[i], dominantFontSizePt);
				columnStates[i].mode = "manual";
				columnStates[i].lockedWidth = null;
				columnStates[i].width = values[i];
				if (autoFitCheckboxes[i].value) autoFitCheckboxes[i].value = false;
			}

			applyColumnSettings();
		}

		// 自動適用は無効 / Auto-apply disabled
		// batchInput.onChange = applyBatchInput;
		btnBatchApply.onClick = applyBatchInput;

		setInputMode("absolute");
		setInputMethod("perColumn");
		widthInputs[0].active = true;

		var dialogResult = dialog.show();

		if (dialogResult != 1) {
			if (isPreviewCurrentlyApplied) restoreOriginalColumnSettings();
			restoreSelection(savedSelection);
			return;
		}

		applyColumnSettings();
		restoreSelection(savedSelection);
	}

	// =========================================
	// 表・選択ヘルパー / Table and selection helpers
	// =========================================

	/**
	 * 控えておいた選択を復元する。無効になった項目は除き、何も残らなければ選択は変えない
	 * @param {Array} selectionItems 実行前に控えた選択
	 * @returns {void}
	 */
	function restoreSelection(selectionItems) {
		var validItems = [];
		for (var i = 0; i < selectionItems.length; i++) {
			if (selectionItems[i] && selectionItems[i].isValid !== false) validItems.push(selectionItems[i]);
		}
		if (validItems.length === 0) return;
		/* 復元できない選択（削除済みの文字など）は諦める / Give up on selections that can no longer be restored */
		try {
			app.select(validItems);
		} catch (e) { }
	}

	/**
	 * 各列の現在の幅を取得する
	 * @param {Table} table 対象の表
	 * @returns {Array<number>} 列幅の配列
	 */
	function getColumnWidths(table) {
		var widths = [];
		for (var i = 0; i < table.columns.length; i++) {
			widths.push(table.columns[i].width);
		}
		return widths;
	}

	/**
	 * 各列の現在の左右余白を控える
	 * @param {Table} table 対象の表
	 * @returns {Array<object>} 列ごとの余白情報
	 */
	function getOriginalInsets(table) {
		var perColumn = [];
		for (var i = 0; i < table.columns.length; i++) {
			var cells = table.columns[i].cells;
			var cellInsets = [];
			for (var j = 0; j < cells.length; j++) {
				cellInsets.push({ left: cells[j].leftInset, right: cells[j].rightInset });
			}
			perColumn.push(cellInsets);
		}
		return perColumn;
	}

	/**
	 * 控えた余白から入力欄の初期値を作る
	 * @param {Array<object>} originalInsets 列ごとの余白情報
	 * @returns {Array<number>} 入力欄の初期値
	 */
	function insetsToInitialInsetValues(originalInsets) {
		var insetValues = [];
		for (var i = 0; i < originalInsets.length; i++) {
			var leftInset = originalInsets[i].length > 0 ? originalInsets[i][0].left : 0;
			insetValues.push(rulerValueToInputUnit(leftInset));
		}
		return insetValues;
	}

	/**
	 * 入力値に従って列幅を設定する
	 * @param {Table} table 対象の表
	 * @param {Array<EditText>} widthInputs 幅の入力欄
	 * @param {Array<number>} originalWidths 元の列幅
	 * @returns {void}
	 */
	function applyColumnWidths(table, widthInputs, originalWidths) {
		for (var i = widthInputs.length - 1; i >= 0; i--) {
			try {
				var text = widthInputs[i].text;
				if (text == "") {
					table.columns[i].width = originalWidths[i];
					continue;
				}
				table.columns[i].width = inputUnitToRulerValue(text * 1);
			} catch (e) { }
		}
	}

	/**
	 * 入力値に従って列の左右余白を設定する
	 * @param {Table} table 対象の表
	 * @param {Array<EditText>} sideInsetInputs 余白の入力欄
	 * @returns {void}
	 */
	function applyColumnInsets(table, sideInsetInputs) {
		for (var i = 0; i < sideInsetInputs.length; i++) {
			var text = sideInsetInputs[i].text;
			if (text == "") continue;
			var value = parseFloat(text);
			if (isNaN(value)) continue;
			var rulerValue = inputUnitToRulerValue(value);
			var cells = table.columns[i].cells;
			for (var j = 0; j < cells.length; j++) {
				try {
					cells[j].leftInset = rulerValue;
					cells[j].rightInset = rulerValue;
				} catch (e) { }
			}
		}
	}

	/**
	 * 列の状態オブジェクトから左右余白を設定する
	 * @param {Table} table 対象の表
	 * @param {Array<object>} columnStates 列ごとの状態
	 * @returns {void}
	 */
	function applyColumnInsetsFromStates(table, columnStates) {
		for (var i = 0; i < columnStates.length; i++) {
			var insetValue = columnStates[i].inset;
			if (insetValue == null || isNaN(insetValue)) continue;
			var rulerValue = inputUnitToRulerValue(insetValue);
			var cells = table.columns[i].cells;
			for (var j = 0; j < cells.length; j++) {
				try {
					cells[j].leftInset = rulerValue;
					cells[j].rightInset = rulerValue;
				} catch (e) { }
			}
		}
	}

	/**
	 * 控えておいた列幅を戻す
	 * @param {Table} table 対象の表
	 * @param {Array<number>} originalWidths 元の列幅
	 * @returns {void}
	 */
	function restoreColumnWidths(table, originalWidths) {
		for (var i = 0; i < originalWidths.length; i++) {
			try { table.columns[i].width = originalWidths[i]; } catch (e) { }
		}
	}

	/**
	 * 控えておいた左右余白を戻す
	 * @param {Table} table 対象の表
	 * @param {Array<object>} originalInsets 元の余白情報
	 * @returns {void}
	 */
	function restoreColumnInsets(table, originalInsets) {
		for (var i = 0; i < originalInsets.length; i++) {
			var cells = table.columns[i].cells;
			for (var j = 0; j < cells.length && j < originalInsets[i].length; j++) {
				try {
					cells[j].leftInset = originalInsets[i][j].left;
					cells[j].rightInset = originalInsets[i][j].right;
				} catch (e) { }
			}
		}
	}

	// =========================================
	// UI ヘルパー / UI helpers
	// =========================================

	/**
	 * 列ごとの設定コントロールを組み立てる
	 * @param {object} parent 追加先のコンテナ
	 * @param {number} columnCount 列数
	 * @param {Array<number>} originalWidths 元の列幅
	 * @param {Array<number>} initialInsetValues 余白の初期値
	 * @param {number} fontSizePt 基準の文字サイズ（pt）
	 * @returns {object} 生成したコントロール
	 */
	function buildColumnSettingsControls(parent, columnCount, originalWidths, initialInsetValues, fontSizePt) {
		var widthInputFields = [];
		var charCountInputFields = [];
		var sideInsetInputFields = [];
		var autoFitToggleCheckboxes = [];
		var rowLabelControls = [];
		var headerLabelControls = [];

		/* 各入力欄の左に∧∨が付くぶん、見出しの幅を広げて中央をそろえる / widen headers by the stepper so they stay centered over the fields */
		var stepperWidth = STEPPER_BUTTON_WIDTH + STEPPER_SIDE_MARGIN;
		/* 増減後は手入力と同じ onChanging・onChange を通す。下限は既存の↑↓キーと同じ0 / after stepping, run the same handlers as typing */
		var columnStepOptions = { min: 0, onStep: runInputHandlersAfterStep };

		// ヘッダー行を追加 / Insert header row
		var headerRow = parent.add("group");
		headerRow.orientation = "row";
		headerRow.spacing = 6;
		headerRow.margins = [0, 2, 0, 6];

		// 「列」
		var headerColumn = headerRow.add("statictext", undefined, getLabel("header.column"));
		headerColumn.justify = "center";
		headerColumn.preferredSize.width = 14;
		headerLabelControls.push(headerColumn);

		//　「幅」
		var headerWidth = headerRow.add("statictext", undefined, getLabel("header.width"));
		headerWidth.justify = "center";
		headerWidth.preferredSize.width = 70 + stepperWidth;
		headerLabelControls.push(headerWidth);

		// var headerMidSpacer = headerRow.add("group");
		// headerMidSpacer.preferredSize.width = 10;

		// 「文字数」
		var headerCharCount = headerRow.add("statictext", undefined, getLabel("header.charCount"));
		headerCharCount.justify = "center";
		headerCharCount.preferredSize.width = 73 + stepperWidth;
		headerLabelControls.push(headerCharCount);

		// var headerSpacer = headerRow.add("group");
		// headerSpacer.preferredSize.width = 16;

		// 「左右の余白」
		var headerInset = headerRow.add("statictext", undefined, getLabel("header.inset"));
		headerInset.justify = "center";
		headerInset.preferredSize.width = 70 + stepperWidth;
		headerLabelControls.push(headerInset);

		// var headerAutoFitSpacer = headerRow.add("group");
		// headerAutoFitSpacer.preferredSize.width = 16;

		// 「自動調整」
		var headerAutoFit = headerRow.add("statictext", undefined, getLabel("header.autoFit"));
		headerAutoFit.justify = "center";
		headerAutoFit.preferredSize.width = 60;
		headerLabelControls.push(headerAutoFit);

		for (var i = 0; i < columnCount; i++) {
			var inputRow = parent.add("group");
			inputRow.orientation = "row";
			inputRow.spacing = 6;

			var rowLabelRow = [];
			var columnLabel = inputRow.add("statictext", undefined, String(i + 1));
			columnLabel.preferredSize.width = 20;
			rowLabelRow.push(columnLabel);

			var widthValue = rulerValueToInputUnit(originalWidths[i]);
			var widthInput = addSteppedInput(inputRow, formatNumber(widthValue), columnStepOptions);
			widthInput.characters = 5;
			// 行ごとの幅単位ラベルは表示しない / Per-row width unit label is omitted
			// rowLabelRow.push(inputRow.add("statictext", undefined, getLabel("unit.mm")));

			var midSpacer = inputRow.add("group");
			midSpacer.preferredSize.width = 10;

			var charCount = calculateCharCount(widthValue, initialInsetValues[i], fontSizePt);
			var charCountInput = addSteppedInput(inputRow, formatNumber(charCount), columnStepOptions);
			charCountInput.characters = 5;
			// 行ごとの文字数単位ラベルは表示しない / Per-row character-count unit label is omitted
			// rowLabelRow.push(inputRow.add("statictext", undefined, getLabel("unit.characters")));

			var rowSpacer = inputRow.add("group");
			rowSpacer.preferredSize.width = 16;

			var sideInsetInput = addSteppedInput(inputRow, formatNumber(initialInsetValues[i]), columnStepOptions);
			sideInsetInput.characters = 5;

			var autoFitSpacer = inputRow.add("group");
			autoFitSpacer.preferredSize.width = 10;

			var autoFitGroup = inputRow.add("group");
			autoFitGroup.preferredSize.width = 20;
			autoFitGroup.alignChildren = ["center", "center"];
			var autoFitCheckbox = autoFitGroup.add("checkbox", undefined, "");

			widthInputFields.push(widthInput);
			charCountInputFields.push(charCountInput);
			sideInsetInputFields.push(sideInsetInput);
			autoFitToggleCheckboxes.push(autoFitCheckbox);
			rowLabelControls.push(rowLabelRow);
		}
		return {
			widthInputs: widthInputFields,
			charCountInputs: charCountInputFields,
			sideInsetInputs: sideInsetInputFields,
			autoFitCheckboxes: autoFitToggleCheckboxes,
			rowLabels: rowLabelControls,
			headerLabels: headerLabelControls
		};
	}

	/**
	 * 幅の入力値から文字数の表示を更新する
	 * @param {EditText} widthInput 幅の入力欄
	 * @param {EditText} charCountInput 文字数の入力欄
	 * @param {EditText} sideInsetInput 余白の入力欄
	 * @param {number} fontSizePt 基準の文字サイズ（pt）
	 * @returns {void}
	 */
	function updateCharCountFromWidth(widthInput, charCountInput, sideInsetInput, fontSizePt) {
		var widthValue = parseFloat(widthInput.text);
		var insetValue = parseFloat(sideInsetInput.text);
		var cc = calculateCharCount(widthValue, insetValue, fontSizePt);
		charCountInput.text = formatNumber(cc);
	}

	/**
	 * 文字数の入力値から幅の表示を更新する
	 * @param {EditText} widthInput 幅の入力欄
	 * @param {EditText} charCountInput 文字数の入力欄
	 * @param {EditText} sideInsetInput 余白の入力欄
	 * @param {number} fontSizePt 基準の文字サイズ（pt）
	 * @returns {void}
	 */
	function updateWidthFromCharCount(widthInput, charCountInput, sideInsetInput, fontSizePt) {
		var cc = parseFloat(charCountInput.text);
		var insetValue = parseFloat(sideInsetInput.text);
		var widthValue = calculateWidthFromCharCount(cc, insetValue, fontSizePt);
		widthInput.text = formatNumber(widthValue);
	}

	/**
	 * 余白の変更に合わせて幅と文字数を揃える
	 * @param {EditText} widthInput 幅の入力欄
	 * @param {EditText} charCountInput 文字数の入力欄
	 * @param {EditText} sideInsetInput 余白の入力欄
	 * @param {number} fontSizePt 基準の文字サイズ（pt）
	 * @param {string} primaryInputMode 現在の指定方法
	 * @returns {void}
	 */
	function syncWidthAndCharCountFromInset(widthInput, charCountInput, sideInsetInput, fontSizePt, primaryInputMode) {
		if (primaryInputMode == "absolute") {
			updateCharCountFromWidth(widthInput, charCountInput, sideInsetInput, fontSizePt);
		}
		else {
			updateWidthFromCharCount(widthInput, charCountInput, sideInsetInput, fontSizePt);
		}
	}

	/**
	 * 「全列に適用」で 1 列の値を他の列へ反映する
	 * @param {Array<EditText>} widthInputs 幅の入力欄
	 * @param {Array<EditText>} charCountInputs 文字数の入力欄
	 * @param {Array<EditText>} sideInsetInputs 余白の入力欄
	 * @param {number} sourceIdx 基準にする列の位置
	 * @param {number} fontSizePt 基準の文字サイズ（pt）
	 * @param {string} primaryInputMode 現在の指定方法
	 * @param {Array<Checkbox>} autoFitCheckboxes 自動調整のチェックボックス
	 * @param {Array<Object>} columnStates 列ごとの状態（書き換える）
	 * @returns {void}
	 */
	function syncUnifiedInputsFromSource(widthInputs, charCountInputs, sideInsetInputs, sourceIdx, fontSizePt, primaryInputMode, autoFitCheckboxes, columnStates) {
		var sourceWidth = parseFloat(widthInputs[sourceIdx].text);
		var sourceChar = parseFloat(charCountInputs[sourceIdx].text);
		var sourceInset = parseFloat(sideInsetInputs[sourceIdx].text);

		var hasWidth = !isNaN(sourceWidth);
		var hasChar = !isNaN(sourceChar);
		var hasInset = !isNaN(sourceInset);

		for (var i = 0; i < widthInputs.length; i++) {
			if (i == sourceIdx) continue;
			if (autoFitCheckboxes && autoFitCheckboxes[i] && autoFitCheckboxes[i].value) continue;

			if (hasInset) {
				sideInsetInputs[i].text = formatNumber(sourceInset);
				columnStates[i].inset = sourceInset;
			}
			else {
				sideInsetInputs[i].text = "";
				columnStates[i].inset = null;
			}

			if (primaryInputMode == "characterBased") {
				if (hasChar) {
					charCountInputs[i].text = formatNumber(sourceChar);
					var recalculatedWidth = calculateWidthFromCharCount(sourceChar, hasInset ? sourceInset : 0, fontSizePt);
					widthInputs[i].text = formatNumber(recalculatedWidth);
					columnStates[i].mode = "manual";
					columnStates[i].lockedWidth = null;
					columnStates[i].width = recalculatedWidth;
				}
				else {
					charCountInputs[i].text = "";
					widthInputs[i].text = "";
					columnStates[i].mode = "manual";
					columnStates[i].lockedWidth = null;
				}
			}
			else {
				if (hasWidth) {
					widthInputs[i].text = formatNumber(sourceWidth);
					var recalculatedChar = calculateCharCount(sourceWidth, hasInset ? sourceInset : 0, fontSizePt);
					charCountInputs[i].text = formatNumber(recalculatedChar);
					columnStates[i].mode = "manual";
					columnStates[i].lockedWidth = null;
					columnStates[i].width = sourceWidth;
				}
				else {
					widthInputs[i].text = "";
					charCountInputs[i].text = "";
					columnStates[i].mode = "manual";
					columnStates[i].lockedWidth = null;
				}
			}
		}
	}

	/**
	 * コントロールの有効／無効を切り替える
	 * @param {object} control 対象のコントロール
	 * @param {boolean} enabled 有効にするなら true
	 * @returns {void}
	 */
	function setControlEnabled(control, enabled) {
		try {
			control.enabled = enabled;
		} catch (e) { }
		setLabelDimmed(control, !enabled);
		/* ∧∨付きの入力欄は∧∨もそろえ、自作描画なので描き直す（表示前は最初の描画で反映される）
		   keep the stepper in step and redraw it; before the window is shown the first draw picks it up */
		if (control.stepperGroup) {
			control.stepperGroup.enabled = enabled;
			if (control.window && control.window.visible) redrawSteppersIn(control.stepperGroup);
		}
	}

	/**
	 * ラベルのディム表示を切り替える
	 * @param {object} control 対象のコントロール
	 * @param {boolean} dim ディム表示にするなら true
	 * @returns {void}
	 */
	function setLabelDimmed(control, dim) {
		try {
			var g = control.graphics;
			var rgb = dim ? [0.55, 0.55, 0.55] : [0, 0, 0];
			g.foregroundColor = g.newPen(g.PenType.SOLID_COLOR, rgb, 1);
		} catch (e) { }
	}

	// =========================================
	// フォントサイズ / Font size
	// =========================================

	var fontSizePtForValidationCache = null;

	/**
	 * 表内で最も支配的な文字サイズを求める
	 * @param {Table} table 対象の表
	 * @returns {number} 文字サイズ（pt）
	 */
	function findDominantFontSize(table) {
		var sizeCharCounts = {};
		var allCells = table.cells;
		for (var i = 0; i < allCells.length; i++) {
			try {
				var styleRanges = allCells[i].textStyleRanges;
				for (var j = 0; j < styleRanges.length; j++) {
					var size = styleRanges[j].pointSize;
					var charLength = styleRanges[j].characters.length;
					if (charLength == 0) continue;
					if (typeof size != "number") continue;
					var key = String(size);
					if (!sizeCharCounts[key]) sizeCharCounts[key] = 0;
					sizeCharCounts[key] += charLength;
				}
			} catch (e) { }
		}

		var dominantSize = null;
		var maxCount = 0;
		for (var key in sizeCharCounts) {
			if (sizeCharCounts[key] > maxCount) {
				maxCount = sizeCharCounts[key];
				dominantSize = Number(key);
			}
		}
		return dominantSize;
	}

	// =========================================
	// 単位変換 / Unit conversion
	// =========================================

	/**
	 * ポイントを Q に換算する
	 * @param {number} pt ポイント値
	 * @returns {number} Q 値
	 */
	function ptToQ(pt) { return pt * 25.4 / 18; }
	/**
	 * ポイントをミリメートルに換算する
	 * @param {number} pt ポイント値
	 * @returns {number} ミリメートル値
	 */
	function ptToMm(pt) { return pt * 25.4 / 72; }
	/**
	 * ミリメートルをポイントに換算する
	 * @param {number} mm ミリメートル値
	 * @returns {number} ポイント値
	 */
	function mmToPt(mm) { return mm * 72 / 25.4; }

	/**
	 * 現在の定規単位の表示文字列を取得する
	 * @returns {string} 単位の文字列
	 */
	function getRulerUnitString() {
		var u = app.activeDocument.viewPreferences.horizontalMeasurementUnits;
		if (u == MeasurementUnits.MILLIMETERS) return "mm";
		if (u == MeasurementUnits.CENTIMETERS) return "cm";
		if (u == MeasurementUnits.POINTS) return "pt";
		if (u == MeasurementUnits.INCHES || u == MeasurementUnits.INCHES_DECIMAL) return "in";
		if (u == MeasurementUnits.PICAS) return "pc";
		return "pt";
	}

	/**
	 * 定規単位の値を入力欄の単位へ変換する
	 * @param {number} value 定規単位での値
	 * @returns {number} 入力欄での値
	 */
	function rulerValueToInputUnit(value) {
		var unit = getRulerUnitString();
		try { return new UnitValue(value, unit).as(unit); } catch (e) { return value; }
	}

	/**
	 * 入力欄の値を定規単位へ変換する
	 * @param {number} value 入力欄での値
	 * @returns {number} 定規単位での値
	 */
	function inputUnitToRulerValue(value) {
		return value;
	}

	/**
	 * 入力欄の値をポイントへ変換する
	 * @param {number} value 入力欄での値
	 * @returns {number} ポイント値
	 */
	function inputUnitToPt(value) {
		var unit = getRulerUnitString();
		if (unit == "pt") return value;
		try { return new UnitValue(value, unit).as("pt"); } catch (e) { return value; }
	}

	/**
	 * ポイント値を入力欄の単位へ変換する
	 * @param {number} value ポイント値
	 * @returns {number} 入力欄での値
	 */
	function ptToInputUnit(value) {
		var unit = getRulerUnitString();
		if (unit == "pt") return value;
		try { return new UnitValue(value, "pt").as(unit); } catch (e) { return value; }
	}

	/**
	 * 幅と余白から収まる文字数を求める
	 * @param {number} widthValue 列幅
	 * @param {number} insetValue 左右の余白
	 * @param {number} fontSizePt 基準の文字サイズ（pt）
	 * @returns {number} 文字数
	 */
	function calculateCharCount(widthValue, insetValue, fontSizePt) {
		if (!fontSizePt || widthValue == null || isNaN(widthValue)) return null;
		var inset = (insetValue == null || isNaN(insetValue)) ? 0 : insetValue;
		var contentValue = widthValue - 2 * inset;
		if (contentValue <= 0) return null;
		return inputUnitToPt(contentValue) / fontSizePt;
	}

	/**
	 * 文字数と余白から必要な列幅を求める
	 * @param {number} charCount 文字数
	 * @param {number} insetValue 左右の余白
	 * @param {number} fontSizePt 基準の文字サイズ（pt）
	 * @returns {number} 列幅
	 */
	function calculateWidthFromCharCount(charCount, insetValue, fontSizePt) {
		if (!fontSizePt || charCount == null || isNaN(charCount)) return null;
		if (charCount < 0) return null;
		var inset = (insetValue == null || isNaN(insetValue)) ? 0 : insetValue;
		return ptToInputUnit(charCount * fontSizePt) + 2 * inset;
	}

	/**
	 * 入力欄に表示する数値を整形する
	 * @param {number} value 表示する数値
	 * @returns {string} 整形した文字列
	 */
	function formatNumber(value) {
		if (value == null || isNaN(value)) return "";
		return String(Math.round(value * 100) / 100);
	}

	/**
	 * 列ごとの入力値を検証する
	 * @param {EditText} widthInput 幅の入力欄
	 * @param {EditText} charCountInput 文字数の入力欄
	 * @param {EditText} sideInsetInput 余白の入力欄
	 * @param {string} primaryInputMode 現在の指定方法
	 * @returns {boolean} すべて有効なら true
	 */
	function validatePerColumnRow(widthInput, charCountInput, sideInsetInput, primaryInputMode) {
		var widthValue = parseFloat(widthInput.text);
		var charValue = parseFloat(charCountInput.text);
		var insetValue = parseFloat(sideInsetInput.text);

		if (sideInsetInput.text !== "") {
			if (isNaN(insetValue)) return { ok: false, message: getLabel("alert.invalidNumber"), focus: sideInsetInput };
			if (insetValue < 0) return { ok: false, message: getLabel("alert.negativeInset"), focus: sideInsetInput };
		}

		if (primaryInputMode == "absolute") {
			if (widthInput.text !== "") {
				if (isNaN(widthValue)) return { ok: false, message: getLabel("alert.invalidNumber"), focus: widthInput };
				if (widthValue < 0) return { ok: false, message: getLabel("alert.negativeWidth"), focus: widthInput };
			}
			if (!isNaN(widthValue) && !isNaN(insetValue) && (widthValue - 2 * insetValue) <= 0) {
				return { ok: false, message: getLabel("alert.insetTooLarge"), focus: sideInsetInput };
			}
		}
		else {
			if (charCountInput.text !== "") {
				if (isNaN(charValue)) return { ok: false, message: getLabel("alert.invalidNumber"), focus: charCountInput };
				if (charValue < 0) return { ok: false, message: getLabel("alert.negativeCharCount"), focus: charCountInput };
			}
			if (!isNaN(charValue) && !isNaN(insetValue)) {
				var calculatedWidth = calculateWidthFromCharCount(charValue, insetValue, fontSizePtForValidationCache);
				if (calculatedWidth == null || (calculatedWidth - 2 * insetValue) <= 0) {
					return { ok: false, message: getLabel("alert.insetTooLarge"), focus: sideInsetInput };
				}
			}
		}

		return { ok: true };
	}

	/**
	 * 一括入力の文字列を数値の配列に変換する
	 * @param {string} text 入力された文字列
	 * @returns {Array<number>|null} 数値の配列。無効な場合は null
	 */
	function parseBatchInput(text) {
		if (text == null) return [];
		var trimmed = text.replace(/^\s+|\s+$/g, "");
		if (trimmed == "") return [];
		// 空白またはカンマで分割 / Split by whitespace or commas
		var parts = trimmed.split(/[\s,]+/);
		var values = [];
		for (var i = 0; i < parts.length; i++) {
			// 空要素は無効値として扱う / Treat empty parts as invalid values
			if (parts[i] == "") { values.push(null); continue; }
			var n = parseFloat(parts[i]);
			// 数値化できない要素は無効値として扱う / Treat non-numeric parts as invalid values
			values.push(isNaN(n) ? null : n);
		}
		return values;
	}

})();