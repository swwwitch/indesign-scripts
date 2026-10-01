#target indesign

/*

### 概要

選択したセルや表の列幅を、列ごとの個別指定または一括入力でまとめて調整します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableColumnWidthAdjuster.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n20d58c1bc003

### Overview

Adjusts the column widths of the selected cells or table, either per column or with a single value applied to all.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableColumnWidthAdjuster.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTableColumnWidthAdjuster";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableColumnWidthAdjuster.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableColumnWidthAdjuster.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n20d58c1bc003"; /* 紹介記事 / article URL */

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
			sizingMethod: { ja: "指定方法", en: "Sizing Method" },
			perColumn:    { ja: "個別に設定", en: "Per-Column Settings" },
			columnWidths: { ja: "各列の設定", en: "Column Settings" },
			batch:        { ja: "一括入力", en: "Batch Input" }
		},
		radio: {
			inputMethodPerColumn: { ja: "個別に設定", en: "Per-Column Input" },
			inputMethodBatch:     { ja: "一括入力", en: "Batch Input" },
			modeAbsolute:         { ja: "幅で指定", en: "Set by Width" },
			modeCharacterBased:   { ja: "文字数で指定", en: "Set by Character Count" }
		},
		checkbox: {
			applyToAll: { ja: "全列に適用", en: "Apply to All Columns" }
		},
		header: {
			column:    { ja: "列", en: "Col" },
			width:     { ja: "幅", en: "Width" },
			charCount: { ja: "文字数", en: "Characters" },
			inset:     { ja: "左右の余白", en: "L/R Inset" },
			autoFit:   { ja: "自動調整", en: "Auto Fit" }
		},
		button: {
			ok:     { ja: "OK", en: "OK" },
			cancel: { ja: "キャンセル", en: "Cancel" },
			apply:  { ja: "適用", en: "Apply" },
			screenModePreview: { ja: "プレビュー", en: "Preview" },
			screenModeNormal:  { ja: "標準モード", en: "Normal Mode" }
		},
		hint: {
			batch: { ja: "入力形式：15 20 34 10 または 15, 20, 34, 10", en: "Format: 15 20 34 10 or 15, 20, 34, 10" }
		},
		alert: {
			selectCellOrTable:  { ja: "セルまたは表を選択してから実行してください", en: "Select a cell or table before running this script." },
			batchInvalidValue:  { ja: "一括入力に無効な値が含まれています", en: "Batch input contains invalid values." },
			batchTooManyValues: { ja: "入力数が列数を超えています", en: "Too many values for the number of columns." },
			invalidNumber:      { ja: "数値を入力してください", en: "Enter a valid number." },
			negativeWidth:      { ja: "幅には 0 以上の数値を入力してください", en: "Width must be 0 or greater." },
			negativeCharCount:  { ja: "文字数には 0 以上の数値を入力してください", en: "Character count must be 0 or greater." },
			negativeInset:      { ja: "左右の余白には 0 以上の数値を入力してください", en: "Left/right inset must be 0 or greater." },
			insetTooLarge:      { ja: "左右の余白が大きすぎます。内容幅が 0 以下になります", en: "Left/right inset is too large. Content width would become 0 or less." }
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
			},
			modeCharacterBased: {
				ja: "表内でいちばん多く使われている文字サイズを1文字として、幅を文字数で指定します。",
				en: "Specifies widths as a character count, based on the most common font size in the table."
			},
			applyToAll: {
				ja: "1列目の値を、ほかの列（自動調整の列を除く）にも反映します。オンにすると自動調整は解除されます。",
				en: "Copies the first column's values to the other columns (except auto-fit columns). Turning this on clears Auto Fit."
			},
			inset: {
				ja: "列内のすべてのセルの左右の余白に、同じ値を設定します。",
				en: "Sets the same left and right inset for every cell in the column."
			},
			autoFit: {
				ja: "列の内容が改行されずに収まる幅に合わせます。",
				en: "Fits the column to the width its contents need without wrapping."
			},
			batchInput: {
				ja: "左の列から順に幅を入力します。入力した数の列にだけ反映します。",
				en: "Enter widths from the leftmost column. Only as many columns as values are changed."
			}
		},
		undo: {
			adjustColumnWidths: { ja: "列幅の調整", en: "Adjust Column Widths" }
		}
	};

	// =========================================
	// 文字色 / Text colors
	// =========================================
	/* 無効な部分をディム表示する色。通常色はダークUIで黒にならないよう明暗で切り替える
	   Color for dimmed parts; the normal color follows the UI brightness so it is not black on a dark UI */
	var TEXT_NORMAL_COLOR = isDarkUI() ? [0.86, 0.86, 0.86] : [0, 0, 0];
	var TEXT_DIMMED_COLOR = [0.55, 0.55, 0.55];

	// =========================================
	// 各列の一覧のレイアウト / Column list layout
	// =========================================
	var COLUMN_ROW_SPACING         = 6;   /* 行内の要素間隔 / spacing within a row */
	var COLUMN_NUMBER_WIDTH        = 20;  /* 列番号の幅 / column number width */
	var COLUMN_FIELD_CHARACTERS    = 5;   /* 入力欄の文字数 / characters of each field */
	var GAP_AFTER_WIDTH_FIELD      = 10;  /* 幅と文字数の間 / gap after the width field */
	var GAP_AFTER_CHAR_COUNT_FIELD = 16;  /* 文字数と余白の間 / gap after the character-count field */
	var GAP_BEFORE_AUTOFIT         = 10;  /* 余白と自動調整の間 / gap before the auto-fit checkbox */
	var AUTOFIT_BOX_WIDTH          = 20;  /* 自動調整のチェックボックスの枠 / auto-fit checkbox box */
	var HEADER_WIDTHS = { column: 14, width: 70, charCount: 73, inset: 70, autoFit: 60 }; /* 見出しの幅（入力欄の見出しは∧∨の幅を足す） / header widths */

	// =========================================
	// 自動調整 / Auto fit
	// =========================================
	var AUTOFIT_COARSE_STEP = 10;  /* 粗く広げる刻み（定規の単位） / coarse step in ruler units */
	var AUTOFIT_FINE_STEP   = 1;   /* 細かく詰める刻み（定規の単位） / fine step in ruler units */
	var AUTOFIT_MAX_STEPS   = 200; /* 1段階あたりの最大回数 / max steps per pass */

	// =========================================
	// メイン / Main
	// =========================================

	/**
	 * 列幅調整の処理を、1回で取り消せるようにして開始する
	 * @returns {void}
	 */
	function main() {
		/* 文字列で渡すとグローバルスコープで評価され IIFE 内の関数を見つけられないため関数で渡す
		   / Pass a function: a string would be evaluated in the global scope and could not see functions inside the IIFE */
		app.doScript(adjustColumnWidths, ScriptLanguage.JAVASCRIPT, undefined,
			UndoModes.ENTIRE_SCRIPT, getLabel("undo.adjustColumnWidths"));
	}

	/**
	 * 選択から表を取り出し、ダイアログを表示して列幅を調整する
	 * @returns {void}
	 */
	function adjustColumnWidths() {
		if (app.documents.length === 0) {
			alert(getLabel("alert.selectCellOrTable"));
			return;
		}
		var savedSelection = app.activeDocument.selection;
		var targetTable = getTableFromSelection(savedSelection[0]);
		if (!targetTable) {
			alert(getLabel("alert.selectCellOrTable"));
			return;
		}

		/* ダイアログ中はセルのハイライトを消し、閉じたら選択を戻す / hide the highlight while the dialog is open */
		app.select(NothingEnum.NOTHING);
		runColumnWidthDialog(targetTable);
		restoreSelection(savedSelection);
	}

	/**
	 * 列幅のダイアログを表示する。編集のたびに表へ反映し、キャンセルなら元に戻す
	 * @param {Table} targetTable 対象の表
	 * @returns {void}
	 */
	function runColumnWidthDialog(targetTable) {
		var columnCount = targetTable.columns.length;
		var originalWidths = getColumnWidths(targetTable);
		var originalInsets = getOriginalInsets(targetTable);
		var rulerUnit = getRulerUnitInfo();
		/* 文字数と幅の換算に使う値 / values for converting between widths and character counts */
		var textMetrics = { fontSizePt: findDominantFontSize(targetTable), pointsPerUnit: rulerUnit.pointsPerUnit };

		/* 表へ反映する値（定規の単位）。入力欄より先にこちらを更新し、反映はここからだけ行う
		   Values applied to the table (ruler units); this is the single source for applying */
		var columnStates = [];
		for (var i = 0; i < columnCount; i++) {
			var firstCellInsets = originalInsets[i][0];
			columnStates.push({ width: originalWidths[i], inset: firstCellInsets ? firstCellInsets.left : 0 });
		}

		var adjusterDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
		setupWindow(adjusterDialog, 10);

		var inputMethodControls = buildInputMethodRow(adjusterDialog);
		var perColumnControls = buildPerColumnPanel(adjusterDialog, columnStates, textMetrics, rulerUnit.label);
		var batchControls = buildBatchInputPanel(adjusterDialog);
		var columnRows = perColumnControls.columnRows;

		var buttonRow = addButtonRow(adjusterDialog);
		addScreenModeButton(buttonRow.leftGroup);
		buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
		buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
		alignRightOnlyButtonRow(buttonRow);

		var sizingMode = "absolute";    /* "absolute"（幅で指定）または "characterBased"（文字数で指定） */
		var inputMethod = "perColumn";  /* "perColumn"（個別に設定）または "batch"（一括入力） */
		var isPreviewApplied = false;

		/**
		 * 入力方法・指定方法・全列に適用・自動調整に合わせて、各コントロールの有効／無効とディム表示を更新する
		 * @returns {void}
		 */
		function refreshControlStates() {
			var isPerColumn = (inputMethod === "perColumn");

			setControlEnabled(batchControls.batchInput, !isPerColumn);
			setControlEnabled(batchControls.applyButton, !isPerColumn);
			setControlEnabled(batchControls.hintText, !isPerColumn);
			setControlEnabled(perColumnControls.absoluteRadio, isPerColumn);
			setControlEnabled(perColumnControls.characterBasedRadio, isPerColumn);
			setControlEnabled(perColumnControls.applyToAllCheckbox, isPerColumn);

			for (var h = 0; h < perColumnControls.headerLabels.length; h++) {
				setLabelDimmed(perColumnControls.headerLabels[h], !isPerColumn);
			}

			for (var i = 0; i < columnRows.length; i++) {
				var row = columnRows[i];
				/* 全列に適用の間は1列目だけを編集する（自動調整の列は対象外） / only the first column is edited while Apply to All is on */
				var followsFirstColumn = perColumnControls.applyToAllCheckbox.value && i > 0 && !row.autoFitCheckbox.value;
				var isEditable = isPerColumn && !followsFirstColumn;
				setControlEnabled(row.widthInput, isEditable && sizingMode === "absolute");
				setControlEnabled(row.charCountInput, isEditable && sizingMode === "characterBased");
				setControlEnabled(row.insetInput, isEditable);
				setControlEnabled(row.autoFitCheckbox, isPerColumn);
				setLabelDimmed(row.numberLabel, !isEditable);
			}

			/* パネルの見出しもディム表示にそろえる / dim the panel titles as well */
			setLabelDimmed(perColumnControls.panel, !isPerColumn);
			setLabelDimmed(perColumnControls.sizingMethodPanel, !isPerColumn);
			setLabelDimmed(perColumnControls.columnSettingsPanel, !isPerColumn);
			setLabelDimmed(batchControls.panel, isPerColumn);
		}

		/**
		 * columnStates の値を表の列幅と左右の余白へ反映する
		 * @returns {void}
		 */
		function applyColumnSettings() {
			for (var i = 0; i < columnStates.length; i++) {
				var columnWidth = columnStates[i].width;
				if (columnWidth == null || isNaN(columnWidth)) continue;
				/* 列幅の上下限を超える値は例外になるので、その列は変えない / out-of-range widths throw; leave that column as is */
				try {
					targetTable.columns[i].width = columnWidth;
				} catch (e) { }
			}
			applyInsetsFromStates(targetTable, columnStates);
			isPreviewApplied = true;
		}

		/**
		 * 列に自動調整の幅を設定し、入力欄と columnStates を更新する
		 * @param {number} columnIndex 列の位置
		 * @returns {void}
		 */
		function applyAutoFitToColumn(columnIndex) {
			var row = columnRows[columnIndex];
			var insetValue = parseFloat(row.insetInput.text);
			if (isNaN(insetValue)) insetValue = 0;
			var fittedWidth = measureColumnContentWidth(targetTable, columnIndex, textMetrics) + 2 * insetValue;

			columnStates[columnIndex].width = fittedWidth;
			columnStates[columnIndex].inset = insetValue;
			/* 空にしてから入れ直し、ScriptUI に描き直させる / clear first so ScriptUI repaints the fields */
			setFieldTextAndRepaint(row.widthInput, formatNumber(fittedWidth));
			setFieldTextAndRepaint(row.charCountInput, formatNumber(calculateCharCount(fittedWidth, insetValue, textMetrics)));
			adjusterDialog.layout.layout(true);
			adjusterDialog.update();
		}

		/**
		 * 余白の変更を幅・文字数へ反映する（自動調整の列は測り直す）
		 * @param {number} columnIndex 列の位置
		 * @returns {void}
		 */
		function refreshColumnFromInset(columnIndex) {
			if (columnRows[columnIndex].autoFitCheckbox.value) applyAutoFitToColumn(columnIndex);
			else syncWidthAndCharCountFromInset(columnRows[columnIndex], textMetrics, sizingMode);
		}

		/**
		 * 全列に適用がオンなら、編集した列の値をほかの列へ写す
		 * @param {number} sourceIndex 編集した列の位置
		 * @param {string} sourceMode 写す基準の値（"absolute" は幅、"characterBased" は文字数）
		 * @returns {void}
		 */
		function syncOtherColumnsIfApplyToAll(sourceIndex, sourceMode) {
			if (!perColumnControls.applyToAllCheckbox.value) return;
			copyColumnValuesToOthers(columnRows, columnStates, sourceIndex, sourceMode, textMetrics);
		}

		/**
		 * 幅・文字数の入力を確定する。不正な値なら知らせて表示を戻し、正しければ手入力の幅として反映する
		 * @param {number} columnIndex 列の位置
		 * @param {string} editedMode 編集した欄（"absolute" は幅、"characterBased" は文字数）
		 * @returns {void}
		 */
		function commitWidthOrCharCount(columnIndex, editedMode) {
			var row = columnRows[columnIndex];
			var validation = validateColumnRow(row, editedMode, textMetrics);
			if (!validation.ok) {
				alert(validation.message);
				validation.focus.active = true;
				if (editedMode === "absolute") updateCharCountFromWidth(row, textMetrics);
				else updateWidthFromCharCount(row, textMetrics);
				return;
			}
			storeTypedValues(columnIndex);
			if (row.autoFitCheckbox.value) {
				row.autoFitCheckbox.value = false;
				refreshControlStates();
			}
			applyColumnSettings();
		}

		/**
		 * 左右の余白の入力を確定する。自動調整の列は測り直す
		 * @param {number} columnIndex 列の位置
		 * @returns {void}
		 */
		function commitInset(columnIndex) {
			var row = columnRows[columnIndex];
			var validation = validateColumnRow(row, sizingMode, textMetrics);
			if (!validation.ok) {
				alert(validation.message);
				validation.focus.active = true;
				refreshColumnFromInset(columnIndex);
				return;
			}
			if (row.autoFitCheckbox.value) applyAutoFitToColumn(columnIndex);
			else storeTypedValues(columnIndex);
			applyColumnSettings();
		}

		/**
		 * 入力欄の幅・余白を columnStates に控える（数値でない欄は前の値のまま）
		 * @param {number} columnIndex 列の位置
		 * @returns {void}
		 */
		function storeTypedValues(columnIndex) {
			var typedWidth = parseFloat(columnRows[columnIndex].widthInput.text);
			var typedInset = parseFloat(columnRows[columnIndex].insetInput.text);
			if (!isNaN(typedWidth)) columnStates[columnIndex].width = typedWidth;
			if (!isNaN(typedInset)) columnStates[columnIndex].inset = typedInset;
		}

		/**
		 * 1列分の入力欄・チェックボックスにイベントを付ける
		 * @param {number} columnIndex 列の位置
		 * @returns {void}
		 */
		function bindColumnRowEvents(columnIndex) {
			var row = columnRows[columnIndex];
			row.widthInput.onChanging = function () {
				updateCharCountFromWidth(row, textMetrics);
				syncOtherColumnsIfApplyToAll(columnIndex, "absolute");
			};
			row.widthInput.onChange = function () { commitWidthOrCharCount(columnIndex, "absolute"); };
			row.charCountInput.onChanging = function () {
				updateWidthFromCharCount(row, textMetrics);
				syncOtherColumnsIfApplyToAll(columnIndex, "characterBased");
			};
			row.charCountInput.onChange = function () { commitWidthOrCharCount(columnIndex, "characterBased"); };
			row.insetInput.onChanging = function () {
				refreshColumnFromInset(columnIndex);
				syncOtherColumnsIfApplyToAll(columnIndex, sizingMode);
			};
			row.insetInput.onChange = function () { commitInset(columnIndex); };
			row.autoFitCheckbox.onClick = function () {
				if (row.autoFitCheckbox.value) applyAutoFitToColumn(columnIndex);
				refreshControlStates();
				applyColumnSettings();
			};
		}

		/**
		 * 一括入力の値を左の列から順に反映する
		 * @returns {void}
		 */
		function applyBatchInput() {
			var batchValues = parseBatchInput(batchControls.batchInput.text);
			if (batchValues.length === 0) return;
			for (var i = 0; i < batchValues.length; i++) {
				if (batchValues[i] == null) {
					alert(getLabel("alert.batchInvalidValue"));
					return;
				}
			}
			if (batchValues.length > columnCount) {
				alert(getLabel("alert.batchTooManyValues"));
				return;
			}
			for (var j = 0; j < batchValues.length; j++) {
				columnRows[j].widthInput.text = formatNumber(batchValues[j]);
				updateCharCountFromWidth(columnRows[j], textMetrics);
				columnStates[j].width = batchValues[j];
				columnRows[j].autoFitCheckbox.value = false;
			}
			applyColumnSettings();
		}

		inputMethodControls.perColumnRadio.onClick = function () {
			inputMethod = "perColumn";
			refreshControlStates();
		};
		inputMethodControls.batchRadio.onClick = function () {
			inputMethod = "batch";
			refreshControlStates();
		};
		perColumnControls.absoluteRadio.onClick = function () {
			sizingMode = "absolute";
			refreshControlStates();
		};
		perColumnControls.characterBasedRadio.onClick = function () {
			sizingMode = "characterBased";
			refreshControlStates();
		};
		perColumnControls.applyToAllCheckbox.onClick = function () {
			/* オンにしたときは自動調整を全列で解除する。値を写すのは次に編集したとき
			   Turning this on clears Auto Fit on every column; values are copied on the next edit */
			if (perColumnControls.applyToAllCheckbox.value) {
				for (var i = 0; i < columnRows.length; i++) columnRows[i].autoFitCheckbox.value = false;
			}
			refreshControlStates();
		};
		for (var columnIndex = 0; columnIndex < columnCount; columnIndex++) {
			bindColumnRowEvents(columnIndex);
		}
		batchControls.applyButton.onClick = applyBatchInput;

		refreshControlStates();
		columnRows[0].widthInput.active = true;

		if (adjusterDialog.show() === 1) {
			applyColumnSettings();
			return;
		}
		if (isPreviewApplied) {
			restoreColumnInsets(targetTable, originalInsets);
			restoreColumnWidths(targetTable, originalWidths);
		}
	}

	// =========================================
	// ダイアログの組み立て / Dialog construction
	// =========================================

	/**
	 * 入力方法（個別に設定／一括入力）のラジオボタンの行を作る
	 * @param {Window} parentWindow 追加先のダイアログ
	 * @returns {{perColumnRadio: RadioButton, batchRadio: RadioButton}} ラジオボタン
	 */
	function buildInputMethodRow(parentWindow) {
		var inputMethodGroup = parentWindow.add("group");
		setupRow(inputMethodGroup, "center", COLUMN_SPACING);
		inputMethodGroup.margins = [0, 3, 0, 10];
		var perColumnRadio = inputMethodGroup.add("radiobutton", undefined, getLabel("radio.inputMethodPerColumn"));
		var batchRadio = inputMethodGroup.add("radiobutton", undefined, getLabel("radio.inputMethodBatch"));
		perColumnRadio.value = true;
		return { perColumnRadio: perColumnRadio, batchRadio: batchRadio };
	}

	/**
	 * 「個別に設定」のパネル（指定方法・各列の設定）を作る
	 * @param {Window} parentWindow 追加先のダイアログ
	 * @param {Array<Object>} columnStates 列ごとの初期値 { width, inset }
	 * @param {Object} textMetrics 文字数の換算に使う値 { fontSizePt, pointsPerUnit }
	 * @param {string} unitLabel 定規の単位の表示名
	 * @returns {Object} パネルとコントロール（columnRows は列ごとの入力欄）
	 */
	function buildPerColumnPanel(parentWindow, columnStates, textMetrics, unitLabel) {
		var perColumnPanel = parentWindow.add("panel", undefined, getLabel("panel.perColumn"));
		setupPanel(perColumnPanel, 10);

		var sizingMethodPanel = perColumnPanel.add("panel", undefined, getLabel("panel.sizingMethod"));
		setupPanel(sizingMethodPanel);
		var sizingMethodRow = sizingMethodPanel.add("group");
		setupRow(sizingMethodRow, "left", COLUMN_SPACING);
		var absoluteRadio = sizingMethodRow.add("radiobutton", undefined, getLabel("radio.modeAbsolute"));
		var characterBasedRadio = sizingMethodRow.add("radiobutton", undefined, getLabel("radio.modeCharacterBased"));
		characterBasedRadio.helpTip = getLabel("tooltip.modeCharacterBased");
		absoluteRadio.value = true;

		var unitSuffix = (uiLang === "ja") ? "（" + unitLabel + "）" : " (" + unitLabel + ")";
		var columnSettingsPanel = perColumnPanel.add("panel", undefined, getLabel("panel.columnWidths") + unitSuffix);
		setupPanel(columnSettingsPanel, 6);
		columnSettingsPanel.alignChildren = "left";
		var applyToAllCheckbox = columnSettingsPanel.add("checkbox", undefined, getLabel("checkbox.applyToAll"));
		applyToAllCheckbox.helpTip = getLabel("tooltip.applyToAll");

		var headerLabels = buildColumnHeaderRow(columnSettingsPanel);
		var columnRows = [];
		for (var i = 0; i < columnStates.length; i++) {
			columnRows.push(buildColumnRow(columnSettingsPanel, i, columnStates[i], textMetrics));
		}

		return {
			panel: perColumnPanel,
			sizingMethodPanel: sizingMethodPanel,
			absoluteRadio: absoluteRadio,
			characterBasedRadio: characterBasedRadio,
			columnSettingsPanel: columnSettingsPanel,
			applyToAllCheckbox: applyToAllCheckbox,
			headerLabels: headerLabels,
			columnRows: columnRows
		};
	}

	/**
	 * 各列の一覧の見出し行を作る
	 * @param {Panel} parent 追加先のパネル
	 * @returns {Array<StaticText>} 見出しの statictext
	 */
	function buildColumnHeaderRow(parent) {
		var headerRow = parent.add("group");
		headerRow.orientation = "row";
		headerRow.spacing = COLUMN_ROW_SPACING;
		headerRow.margins = [0, 2, 0, 6];

		/* 各入力欄の左に∧∨が付くぶん、見出しの幅を広げて中央をそろえる / widen headers by the stepper so they stay centered over the fields */
		var stepperWidth = STEPPER_BUTTON_WIDTH + STEPPER_SIDE_MARGIN;
		var headerSpecs = [
			{ key: "column", width: HEADER_WIDTHS.column },
			{ key: "width", width: HEADER_WIDTHS.width + stepperWidth },
			{ key: "charCount", width: HEADER_WIDTHS.charCount + stepperWidth },
			{ key: "inset", width: HEADER_WIDTHS.inset + stepperWidth, tooltip: "tooltip.inset" },
			{ key: "autoFit", width: HEADER_WIDTHS.autoFit, tooltip: "tooltip.autoFit" }
		];
		var headerLabels = [];
		for (var i = 0; i < headerSpecs.length; i++) {
			var headerLabel = headerRow.add("statictext", undefined, getLabel("header." + headerSpecs[i].key));
			headerLabel.justify = "center";
			headerLabel.preferredSize.width = headerSpecs[i].width;
			if (headerSpecs[i].tooltip) headerLabel.helpTip = getLabel(headerSpecs[i].tooltip);
			headerLabels.push(headerLabel);
		}
		return headerLabels;
	}

	/**
	 * 1列分の行（列番号・幅・文字数・左右の余白・自動調整）を作る
	 * @param {Panel} parent 追加先のパネル
	 * @param {number} columnIndex 列の位置
	 * @param {Object} columnState 列の初期値 { width, inset }
	 * @param {Object} textMetrics 文字数の換算に使う値 { fontSizePt, pointsPerUnit }
	 * @returns {Object} 行のコントロール { numberLabel, widthInput, charCountInput, insetInput, autoFitCheckbox }
	 */
	function buildColumnRow(parent, columnIndex, columnState, textMetrics) {
		/* 増減後は手入力と同じ onChanging・onChange を通す / after stepping, run the same handlers as typing */
		var columnStepOptions = { min: 0, onStep: runInputHandlersAfterStep };

		var columnRow = parent.add("group");
		columnRow.orientation = "row";
		columnRow.spacing = COLUMN_ROW_SPACING;

		var numberLabel = columnRow.add("statictext", undefined, String(columnIndex + 1));
		numberLabel.preferredSize.width = COLUMN_NUMBER_WIDTH;

		var widthInput = addSteppedInput(columnRow, formatNumber(columnState.width), columnStepOptions);
		widthInput.characters = COLUMN_FIELD_CHARACTERS;
		addFixedGap(columnRow, GAP_AFTER_WIDTH_FIELD);

		var charCount = calculateCharCount(columnState.width, columnState.inset, textMetrics);
		var charCountInput = addSteppedInput(columnRow, formatNumber(charCount), columnStepOptions);
		charCountInput.characters = COLUMN_FIELD_CHARACTERS;
		addFixedGap(columnRow, GAP_AFTER_CHAR_COUNT_FIELD);

		var insetInput = addSteppedInput(columnRow, formatNumber(columnState.inset), columnStepOptions);
		insetInput.characters = COLUMN_FIELD_CHARACTERS;
		insetInput.helpTip = getLabel("tooltip.inset");
		addFixedGap(columnRow, GAP_BEFORE_AUTOFIT);

		var autoFitGroup = columnRow.add("group");
		autoFitGroup.preferredSize.width = AUTOFIT_BOX_WIDTH;
		autoFitGroup.alignChildren = ["center", "center"];
		var autoFitCheckbox = autoFitGroup.add("checkbox", undefined, "");
		autoFitCheckbox.helpTip = getLabel("tooltip.autoFit");

		return {
			numberLabel: numberLabel,
			widthInput: widthInput,
			charCountInput: charCountInput,
			insetInput: insetInput,
			autoFitCheckbox: autoFitCheckbox
		};
	}

	/**
	 * 行に固定幅の空きを足す
	 * @param {Group} parentRow 追加先の行
	 * @param {number} gapWidth 空きの幅（px）
	 * @returns {void}
	 */
	function addFixedGap(parentRow, gapWidth) {
		parentRow.add("group").preferredSize.width = gapWidth;
	}

	/**
	 * 「一括入力」のパネルを作る
	 * @param {Window} parentWindow 追加先のダイアログ
	 * @returns {{panel: Panel, batchInput: EditText, applyButton: Button, hintText: StaticText}} パネルとコントロール
	 */
	function buildBatchInputPanel(parentWindow) {
		var batchInputPanel = parentWindow.add("panel", undefined, getLabel("panel.batch"));
		setupPanel(batchInputPanel, 6);

		var batchRow = batchInputPanel.add("group");
		batchRow.orientation = "row";
		batchRow.alignChildren = ["fill", "center"];
		batchRow.spacing = 8;

		var batchInput = batchRow.add("edittext", undefined, "");
		batchInput.alignment = ["fill", "center"];
		batchInput.helpTip = getLabel("tooltip.batchInput");

		var applyButton = batchRow.add("button", undefined, getLabel("button.apply"));
		applyButton.alignment = ["right", "center"];

		var hintText = batchInputPanel.add("statictext", undefined, getLabel("hint.batch"));

		return { panel: batchInputPanel, batchInput: batchInput, applyButton: applyButton, hintText: hintText };
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
	 * 各列の全セルの左右の余白を控える
	 * @param {Table} table 対象の表
	 * @returns {Array<Array<Object>>} 列ごと・セルごとの { left, right }
	 */
	function getOriginalInsets(table) {
		var insetsByColumn = [];
		for (var i = 0; i < table.columns.length; i++) {
			var columnCells = table.columns[i].cells;
			var cellInsets = [];
			for (var j = 0; j < columnCells.length; j++) {
				cellInsets.push({ left: columnCells[j].leftInset, right: columnCells[j].rightInset });
			}
			insetsByColumn.push(cellInsets);
		}
		return insetsByColumn;
	}

	/**
	 * 列の状態から、列内の全セルの左右の余白を設定する
	 * @param {Table} table 対象の表
	 * @param {Array<Object>} columnStates 列ごとの状態 { width, inset }
	 * @returns {void}
	 */
	function applyInsetsFromStates(table, columnStates) {
		for (var i = 0; i < columnStates.length; i++) {
			var insetValue = columnStates[i].inset;
			if (insetValue == null || isNaN(insetValue)) continue;
			var columnCells = table.columns[i].cells;
			for (var j = 0; j < columnCells.length; j++) {
				/* セル幅に収まらない余白は例外になるので、そのセルは変えない / insets wider than the cell throw; skip that cell */
				try {
					columnCells[j].leftInset = insetValue;
					columnCells[j].rightInset = insetValue;
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
			/* 戻す途中で幅と余白がぶつかると例外になるので、その列は飛ばす / skip a column when the width clashes with the current insets */
			try {
				table.columns[i].width = originalWidths[i];
			} catch (e) { }
		}
	}

	/**
	 * 控えておいた左右の余白を戻す
	 * @param {Table} table 対象の表
	 * @param {Array<Array<Object>>} originalInsets 元の余白
	 * @returns {void}
	 */
	function restoreColumnInsets(table, originalInsets) {
		for (var i = 0; i < originalInsets.length; i++) {
			var columnCells = table.columns[i].cells;
			for (var j = 0; j < columnCells.length && j < originalInsets[i].length; j++) {
				/* セル幅に収まらない余白は例外になるので、そのセルは飛ばす / skip a cell when the inset does not fit */
				try {
					columnCells[j].leftInset = originalInsets[i][j].left;
					columnCells[j].rightInset = originalInsets[i][j].right;
				} catch (e) { }
			}
		}
	}

	// =========================================
	// 自動調整の測定 / Auto-fit measurement
	// =========================================

	/**
	 * セルの行を返す（自動の折り返しも1行に数える）
	 * @param {Cell} cell 対象のセル
	 * @returns {Lines|null} 行のコレクション。読めないセルは null
	 */
	function getCellLines(cell) {
		/* 結合で隠れたセルなど、文字を読めないセルは空として扱う / treat cells whose text cannot be read as empty */
		try {
			return cell.texts[0].lines;
		} catch (e) {
			return null;
		}
	}

	/**
	 * 列内のいずれかのセルが2行以上になっているかを返す
	 * @param {Cells} columnCells 列のセル
	 * @returns {boolean} 2行目があれば true
	 */
	function hasWrappedCell(columnCells) {
		for (var i = 0; i < columnCells.length; i++) {
			var cellLines = getCellLines(columnCells[i]);
			if (cellLines && cellLines.length >= 2) return true;
		}
		return false;
	}

	/**
	 * いちばん長い行の文字数から、列の内容幅を見積もる
	 * @param {Cells} columnCells 列のセル
	 * @param {Object} textMetrics 文字数の換算に使う値 { fontSizePt, pointsPerUnit }
	 * @returns {number} 見積もった内容幅（定規の単位）
	 */
	function estimateContentWidthByChars(columnCells, textMetrics) {
		var maxLineLength = 0;
		for (var i = 0; i < columnCells.length; i++) {
			var cellLines = getCellLines(columnCells[i]);
			if (!cellLines) continue;
			for (var j = 0; j < cellLines.length; j++) {
				maxLineLength = Math.max(maxLineLength, cellLines[j].characters.length);
			}
		}
		return maxLineLength * textMetrics.fontSizePt / textMetrics.pointsPerUnit;
	}

	/**
	 * 列の内容が改行されずに収まる幅（左右の余白を除く）を求める。
	 * 折り返しのある列は実際に幅を広げて測り（粗く広げてから細かく詰める）、測ったあと元の幅に戻す。
	 * 折り返しの無い列は文字数から見積もる
	 * @param {Table} table 対象の表
	 * @param {number} columnIndex 列の位置
	 * @param {Object} textMetrics 文字数の換算に使う値 { fontSizePt, pointsPerUnit }
	 * @returns {number} 内容幅（定規の単位）
	 */
	function measureColumnContentWidth(table, columnIndex, textMetrics) {
		var targetColumn = table.columns[columnIndex];
		var columnCells = targetColumn.cells;
		if (!hasWrappedCell(columnCells)) return estimateContentWidthByChars(columnCells, textMetrics);

		var savedWidth = targetColumn.width;
		var measuredWidth;
		try {
			/* 1) 粗く広げて、2行目が消える幅まで到達する / widen coarsely until no cell wraps */
			var refineStartWidth = Math.max(0, widenColumnUntilUnwrapped(targetColumn, savedWidth, AUTOFIT_COARSE_STEP) - AUTOFIT_COARSE_STEP);
			/* 2) 1段階戻して、細かい刻みで詰める。下限を割る幅は例外になるので、そのときは今の幅から詰める
			   step back once and refine; a width below the minimum throws, so refine from the current width then */
			try {
				targetColumn.width = refineStartWidth;
			} catch (e) { }
			measuredWidth = widenColumnUntilUnwrapped(targetColumn, refineStartWidth, AUTOFIT_FINE_STEP);
		} finally {
			targetColumn.width = savedWidth;
		}
		/* 測った列幅には今の左右の余白が含まれるので除く / the measured width includes the current insets */
		return measuredWidth - columnCells[0].leftInset - columnCells[0].rightInset;
	}

	/**
	 * 折り返しが無くなるまで列幅を step ずつ広げる（AUTOFIT_MAX_STEPS 回まで。列幅の上限に当たったらそこで止める）
	 * @param {Column} targetColumn 対象の列
	 * @param {number} startWidth 広げ始める幅
	 * @param {number} step 刻み（定規の単位）
	 * @returns {number} 広げたあとの幅
	 */
	function widenColumnUntilUnwrapped(targetColumn, startWidth, step) {
		var columnWidth = startWidth;
		for (var i = 0; i < AUTOFIT_MAX_STEPS && hasWrappedCell(targetColumn.cells); i++) {
			/* 上限を超える幅は例外になる / widths above the maximum throw */
			try {
				targetColumn.width = columnWidth + step;
			} catch (e) {
				break;
			}
			columnWidth += step;
		}
		return columnWidth;
	}

	// =========================================
	// 入力欄の連動 / Field synchronization
	// =========================================

	/**
	 * 幅の入力値から文字数の表示を更新する
	 * @param {Object} columnRow buildColumnRow() の戻り値
	 * @param {Object} textMetrics 文字数の換算に使う値
	 * @returns {void}
	 */
	function updateCharCountFromWidth(columnRow, textMetrics) {
		var charCount = calculateCharCount(parseFloat(columnRow.widthInput.text), parseFloat(columnRow.insetInput.text), textMetrics);
		columnRow.charCountInput.text = formatNumber(charCount);
	}

	/**
	 * 文字数の入力値から幅の表示を更新する
	 * @param {Object} columnRow buildColumnRow() の戻り値
	 * @param {Object} textMetrics 文字数の換算に使う値
	 * @returns {void}
	 */
	function updateWidthFromCharCount(columnRow, textMetrics) {
		var columnWidth = calculateWidthFromCharCount(parseFloat(columnRow.charCountInput.text), parseFloat(columnRow.insetInput.text), textMetrics);
		columnRow.widthInput.text = formatNumber(columnWidth);
	}

	/**
	 * 余白の変更に合わせて、指定方法で固定していないほうの値を計算し直す
	 * @param {Object} columnRow buildColumnRow() の戻り値
	 * @param {Object} textMetrics 文字数の換算に使う値
	 * @param {string} sizingMode "absolute"（幅を固定）または "characterBased"（文字数を固定）
	 * @returns {void}
	 */
	function syncWidthAndCharCountFromInset(columnRow, textMetrics, sizingMode) {
		if (sizingMode === "absolute") updateCharCountFromWidth(columnRow, textMetrics);
		else updateWidthFromCharCount(columnRow, textMetrics);
	}

	/**
	 * 1列の値をほかの列（自動調整の列を除く）へ写し、columnStates も更新する
	 * @param {Array<Object>} columnRows 列ごとの入力欄
	 * @param {Array<Object>} columnStates 列ごとの状態（書き換える）
	 * @param {number} sourceIndex 写し元の列の位置
	 * @param {string} sourceMode 写す基準の値（"absolute" は幅、"characterBased" は文字数）
	 * @param {Object} textMetrics 文字数の換算に使う値
	 * @returns {void}
	 */
	function copyColumnValuesToOthers(columnRows, columnStates, sourceIndex, sourceMode, textMetrics) {
		var sourceRow = columnRows[sourceIndex];
		var sourceWidth = parseFloat(sourceRow.widthInput.text);
		var sourceCharCount = parseFloat(sourceRow.charCountInput.text);
		var sourceInset = parseFloat(sourceRow.insetInput.text);
		var hasInset = !isNaN(sourceInset);
		var insetForCalc = hasInset ? sourceInset : 0;

		for (var i = 0; i < columnRows.length; i++) {
			var row = columnRows[i];
			if (i === sourceIndex || row.autoFitCheckbox.value) continue;

			row.insetInput.text = hasInset ? formatNumber(sourceInset) : "";
			columnStates[i].inset = hasInset ? sourceInset : null;

			var sourceValue = (sourceMode === "characterBased") ? sourceCharCount : sourceWidth;
			if (isNaN(sourceValue)) {
				/* 写し元が空なら空にする（表の列幅はそのまま） / mirror an empty source; the table width stays */
				row.widthInput.text = "";
				row.charCountInput.text = "";
				continue;
			}
			var columnWidth = sourceWidth;
			var charCount = calculateCharCount(sourceWidth, insetForCalc, textMetrics);
			if (sourceMode === "characterBased") {
				columnWidth = calculateWidthFromCharCount(sourceCharCount, insetForCalc, textMetrics);
				charCount = sourceCharCount;
			}
			row.widthInput.text = formatNumber(columnWidth);
			row.charCountInput.text = formatNumber(charCount);
			columnStates[i].width = columnWidth;
		}
	}

	/**
	 * 1列分の入力値を検証する
	 * @param {Object} columnRow buildColumnRow() の戻り値
	 * @param {string} sizingMode "absolute"（幅を検証）または "characterBased"（文字数を検証）
	 * @param {Object} textMetrics 文字数の換算に使う値
	 * @returns {{ok: boolean, message: string, focus: EditText}} 結果。不正なら message と戻すフォーカス先が入る
	 */
	function validateColumnRow(columnRow, sizingMode, textMetrics) {
		var widthValue = parseFloat(columnRow.widthInput.text);
		var charValue = parseFloat(columnRow.charCountInput.text);
		var insetValue = parseFloat(columnRow.insetInput.text);

		if (columnRow.insetInput.text !== "") {
			if (isNaN(insetValue)) return { ok: false, message: getLabel("alert.invalidNumber"), focus: columnRow.insetInput };
			if (insetValue < 0) return { ok: false, message: getLabel("alert.negativeInset"), focus: columnRow.insetInput };
		}

		if (sizingMode === "absolute") {
			if (columnRow.widthInput.text !== "") {
				if (isNaN(widthValue)) return { ok: false, message: getLabel("alert.invalidNumber"), focus: columnRow.widthInput };
				if (widthValue < 0) return { ok: false, message: getLabel("alert.negativeWidth"), focus: columnRow.widthInput };
			}
			if (!isNaN(widthValue) && !isNaN(insetValue) && (widthValue - 2 * insetValue) <= 0) {
				return { ok: false, message: getLabel("alert.insetTooLarge"), focus: columnRow.insetInput };
			}
			return { ok: true };
		}

		if (columnRow.charCountInput.text !== "") {
			if (isNaN(charValue)) return { ok: false, message: getLabel("alert.invalidNumber"), focus: columnRow.charCountInput };
			if (charValue < 0) return { ok: false, message: getLabel("alert.negativeCharCount"), focus: columnRow.charCountInput };
		}
		if (!isNaN(charValue) && !isNaN(insetValue)) {
			var calculatedWidth = calculateWidthFromCharCount(charValue, insetValue, textMetrics);
			if (calculatedWidth == null || (calculatedWidth - 2 * insetValue) <= 0) {
				return { ok: false, message: getLabel("alert.insetTooLarge"), focus: columnRow.insetInput };
			}
		}
		return { ok: true };
	}

	// =========================================
	// UI ヘルパー / UI helpers
	// =========================================

	/**
	 * コントロールの有効／無効とディム表示を切り替える（∧∨付きの入力欄は∧∨もそろえる）
	 * @param {Object} control 対象のコントロール
	 * @param {boolean} isEnabled 有効にするなら true
	 * @returns {void}
	 */
	function setControlEnabled(control, isEnabled) {
		control.enabled = isEnabled;
		setLabelDimmed(control, !isEnabled);
		/* ∧∨は自作描画なので描き直す（表示前は最初の描画で反映される）
		   the stepper is custom-drawn, so redraw it; before the window is shown the first draw picks it up */
		if (control.stepperGroup) {
			control.stepperGroup.enabled = isEnabled;
			if (control.window && control.window.visible) redrawSteppersIn(control.stepperGroup);
		}
	}

	/**
	 * 文字（パネルは見出し）の色をディム表示と通常色で切り替える
	 * @param {Object} control 対象のコントロール
	 * @param {boolean} isDimmed ディム表示にするなら true
	 * @returns {void}
	 */
	function setLabelDimmed(control, isDimmed) {
		/* 文字色を持たないコントロールでは例外になるので、色は変えない / controls without a text color throw; leave them */
		try {
			var controlGraphics = control.graphics;
			controlGraphics.foregroundColor = controlGraphics.newPen(controlGraphics.PenType.SOLID_COLOR,
				isDimmed ? TEXT_DIMMED_COLOR : TEXT_NORMAL_COLOR, 1);
		} catch (e) { }
	}

	/**
	 * 入力欄の文字を、いったん空にしてから入れ直す（ScriptUI に確実に描き直させる）
	 * @param {EditText} targetField 対象の入力欄
	 * @param {string} fieldText 入れる文字
	 * @returns {void}
	 */
	function setFieldTextAndRepaint(targetField, fieldText) {
		targetField.text = "";
		targetField.text = fieldText;
	}

	// =========================================
	// 文字サイズ・単位・数値 / Font size, units and numbers
	// =========================================

	/**
	 * 表内でいちばん多くの文字に使われている文字サイズを求める
	 * @param {Table} table 対象の表
	 * @returns {number|null} 文字サイズ（pt）。文字が無ければ null
	 */
	function findDominantFontSize(table) {
		var charCountBySize = {};
		var allCells = table.cells;
		for (var i = 0; i < allCells.length; i++) {
			/* 結合で隠れたセルなど、書式を読めないセルは飛ばす / skip cells whose formatting cannot be read */
			try {
				var styleRanges = allCells[i].textStyleRanges;
				for (var j = 0; j < styleRanges.length; j++) {
					var pointSize = styleRanges[j].pointSize;
					var rangeLength = styleRanges[j].characters.length;
					if (rangeLength === 0 || typeof pointSize !== "number") continue;
					var sizeKey = String(pointSize);
					charCountBySize[sizeKey] = (charCountBySize[sizeKey] || 0) + rangeLength;
				}
			} catch (e) { }
		}

		var dominantSize = null;
		var maxCharCount = 0;
		for (var sizeKey2 in charCountBySize) {
			if (charCountBySize[sizeKey2] > maxCharCount) {
				maxCharCount = charCountBySize[sizeKey2];
				dominantSize = Number(sizeKey2);
			}
		}
		return dominantSize;
	}

	/**
	 * 横方向の定規の単位について、表示名と 1 単位あたりのポイント数を返す
	 * @returns {{label: string, pointsPerUnit: number}} 単位の情報
	 */
	function getRulerUnitInfo() {
		switch (app.activeDocument.viewPreferences.horizontalMeasurementUnits) {
			case MeasurementUnits.MILLIMETERS: return { label: "mm", pointsPerUnit: 72 / 25.4 };
			case MeasurementUnits.CENTIMETERS: return { label: "cm", pointsPerUnit: 72 / 2.54 };
			case MeasurementUnits.INCHES:
			case MeasurementUnits.INCHES_DECIMAL: return { label: "in", pointsPerUnit: 72 };
			case MeasurementUnits.PICAS: return { label: "pc", pointsPerUnit: 12 };
			case MeasurementUnits.AGATES: return { label: "ag", pointsPerUnit: 5.5 };
			case MeasurementUnits.Q: return { label: "Q", pointsPerUnit: 72 / 25.4 * 0.25 };
			case MeasurementUnits.HA: return { label: "H", pointsPerUnit: 72 / 25.4 * 0.25 };
			case MeasurementUnits.CICEROS: return { label: "c", pointsPerUnit: 12.7883 };
			case MeasurementUnits.PIXELS: return { label: "px", pointsPerUnit: 1 };
			default: return { label: "pt", pointsPerUnit: 1 };
		}
	}

	/**
	 * 列幅と余白から、内容幅に収まる文字数を求める
	 * @param {number} widthValue 列幅（定規の単位）
	 * @param {number} insetValue 左右の余白（定規の単位、未入力なら 0 扱い）
	 * @param {Object} textMetrics 文字数の換算に使う値 { fontSizePt, pointsPerUnit }
	 * @returns {number|null} 文字数。求められなければ null
	 */
	function calculateCharCount(widthValue, insetValue, textMetrics) {
		if (!textMetrics.fontSizePt || widthValue == null || isNaN(widthValue)) return null;
		var inset = (insetValue == null || isNaN(insetValue)) ? 0 : insetValue;
		var contentWidth = widthValue - 2 * inset;
		if (contentWidth <= 0) return null;
		return contentWidth * textMetrics.pointsPerUnit / textMetrics.fontSizePt;
	}

	/**
	 * 文字数と余白から、必要な列幅を求める
	 * @param {number} charCount 文字数
	 * @param {number} insetValue 左右の余白（定規の単位、未入力なら 0 扱い）
	 * @param {Object} textMetrics 文字数の換算に使う値 { fontSizePt, pointsPerUnit }
	 * @returns {number|null} 列幅（定規の単位）。求められなければ null
	 */
	function calculateWidthFromCharCount(charCount, insetValue, textMetrics) {
		if (!textMetrics.fontSizePt || charCount == null || isNaN(charCount) || charCount < 0) return null;
		var inset = (insetValue == null || isNaN(insetValue)) ? 0 : insetValue;
		return charCount * textMetrics.fontSizePt / textMetrics.pointsPerUnit + 2 * inset;
	}

	/**
	 * 入力欄に表示する数値を小数第2位で丸める
	 * @param {number|null} value 表示する数値
	 * @returns {string} 整形した文字列。数値でなければ空文字
	 */
	function formatNumber(value) {
		if (value == null || isNaN(value)) return "";
		return String(Math.round(value * 100) / 100);
	}

	/**
	 * 一括入力の文字列を数値の配列に変換する（空白またはカンマ区切り）
	 * @param {string} batchText 入力された文字列
	 * @returns {Array<number|null>} 数値の配列。数値にできない要素は null
	 */
	function parseBatchInput(batchText) {
		var trimmedText = String(batchText || "").replace(/^\s+|\s+$/g, "");
		if (trimmedText === "") return [];
		var parts = trimmedText.split(/[\s,]+/);
		var batchValues = [];
		for (var i = 0; i < parts.length; i++) {
			var parsedValue = parseFloat(parts[i]);
			batchValues.push(isNaN(parsedValue) ? null : parsedValue);
		}
		return batchValues;
	}

	main();

})();
