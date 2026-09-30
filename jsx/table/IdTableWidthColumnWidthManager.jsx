#target indesign

/*

### 概要

選択位置から対象の表を特定し、表全体の幅と列の幅をプレビュー付きでまとめて調整します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableWidthColumnWidthManager.md

### Overview

Finds the target table from the current selection and adjusts the overall table width and the column widths together, with a live preview.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableWidthColumnWidthManager.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTableWidthColumnWidthManager"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableWidthColumnWidthManager.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableWidthColumnWidthManager.md"; /* README (English) */

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
                title: { ja: "表全体の幅、列の幅の調整", en: "Adjust Table and Column Widths" }
            },
            panel: {
                tableWidth:  { ja: "表全体の幅", en: "Table Width" },
                columnWidth: { ja: "列の幅", en: "Column Width" }
            },
            radio: {
                widthKeep:       { ja: "変更しない", en: "Do Not Change" },
                widthAuto:       { ja: "自動調整", en: "Auto" },
                widthFit:        { ja: "親フレームいっぱいに", en: "Fit to Parent Frame" },
                widthCustom:     { ja: "指定", en: "Custom" },
                cellEqual:       { ja: "均等に", en: "Equal Widths" },
                cellNatural:     { ja: "自動調整", en: "Fit to Content" },
                cellNaturalPlus: { ja: "自動調整を維持", en: "Fit to Content+" },
                cellLast:        { ja: "最終列のみ調整", en: "Adjust Last Column" },
                cellCustom:      { ja: "指定", en: "Custom" }
            },
            button: {
                ok:          { ja: "OK", en: "OK" },
                cancel:      { ja: "キャンセル", en: "Cancel" },
                screenModePreview: { ja: "プレビュー", en: "Preview" },
                screenModeNormal:  { ja: "標準モード", en: "Normal Mode" }
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
                applyWidths: { ja: "表全体の幅、列の幅の調整", en: "Adjust Table and Column Widths" }
            },
            error: {
                selectTable:        { ja: "テーブル内にカーソルを置いてください。", en: "Place the cursor inside a table." },
                textFrameNotFound:  { ja: "テーブルがテキストフレーム内に見つかりません。", en: "The table was not found inside a text frame." }
            }
        };

        /**
         * 現在の定規単位のラベルと単位記号を取得する
         * @returns {{label: string, unitValue: string}} 単位の情報
         */
        function getCurrentRulerUnitInfo() {
            var doc = app.activeDocument;
            var unit = doc.viewPreferences.horizontalMeasurementUnits;

            switch (unit) {
                case MeasurementUnits.POINTS:
                    return { label: "pt", unitValue: "pt" };
                case MeasurementUnits.PICAS:
                    return { label: "pica", unitValue: "pc" };
                case MeasurementUnits.INCHES:
                case MeasurementUnits.INCHES_DECIMAL:
                    return { label: "in", unitValue: "in" };
                case MeasurementUnits.MILLIMETERS:
                    return { label: "mm", unitValue: "mm" };
                case MeasurementUnits.CENTIMETERS:
                    return { label: "cm", unitValue: "cm" };
                default:
                    return { label: "pt", unitValue: "pt" };
            }
        }

        /**
         * 小数第 1 位に丸める
         * @param {number} value 対象の数値
         * @returns {number} 丸めた数値
         */
        function roundToOneDecimal(value) {
            return Math.round(value * 10) / 10;
        }

        /**
         * 表の 1 列目の幅を取得する
         * @param {Table} table 対象の表
         * @returns {number} 列幅。取得できない場合は 0
         */
        function getFirstColumnWidth(table) {
            if (!table || !table.columns || table.columns.length === 0) return 0;
            return table.columns[0].width;
        }

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
         * 選択から表を特定し、幅調整ダイアログを表示して適用する
         * @returns {void}
         */
        function runAdjustTableWidth() {
            var selection = app.selection;

            if (selection.length == 0) {
                alert(getLabel("error.selectTable"));
                return;
            }

            var targetTable = getTableFromSelection(selection[0]);
            if (!targetTable) {
                alert(getLabel("error.selectTable"));
                return;
            }

            // 選択セルを記憶し、ハイライト表示を消す(プレビューが見やすくなるように)
            var savedSelection = [];
            for (var selectionIndex = 0; selectionIndex < selection.length; selectionIndex++) {
                savedSelection.push(selection[selectionIndex]);
            }
            try {
                app.select(NothingEnum.NOTHING);
            } catch (e) { /* 無視 */ }

            // 元の状態をスナップショット(プレビュー復元用)
            var originalTableWidth = targetTable.width;
            var originalColumnWidths = [];
            var originalLeftInsets = [];
            var originalRightInsets = [];
            for (var columnIndex = 0; columnIndex < targetTable.columns.length; columnIndex++) {
                originalColumnWidths.push(targetTable.columns[columnIndex].width);
                originalLeftInsets.push(targetTable.columns[columnIndex].leftInset);
                originalRightInsets.push(targetTable.columns[columnIndex].rightInset);
            }

            /**
             * プレビューで加えた変更を元に戻す
             * @returns {void}
             */
            function revert() {
                targetTable.width = originalTableWidth;
                for (var columnIndex = 0; columnIndex < targetTable.columns.length; columnIndex++) {
                    targetTable.columns[columnIndex].width = originalColumnWidths[columnIndex];
                    targetTable.columns[columnIndex].leftInset = originalLeftInsets[columnIndex];
                    targetTable.columns[columnIndex].rightInset = originalRightInsets[columnIndex];
                }
            }

            /**
             * ダイアログの結果に従って表と列の幅を適用する
             * @param {object} result ダイアログが返した設定
             * @returns {void}
             */
            function applyTableWidthResult(result) {
                var currentPreviewColumnWidths = null;
                if (result.columnWidth && result.columnWidth.value === "last") {
                    currentPreviewColumnWidths = [];
                    for (var columnIndex = 0; columnIndex < targetTable.columns.length; columnIndex++) {
                        currentPreviewColumnWidths.push(targetTable.columns[columnIndex].width);
                    }
                }
                revert();

                // 1. 目標の表幅を算出(この時点では代入しない)
                var widthMode = result.tableWidth.value;
                var columnWidthMode = result.columnWidth.value;
                var targetTableWidth = targetTable.width; // default: keep
                var autoWidth = false;
                var originalAutoFitColumnWidths = null;
                var originalAutoFitTableWidth = null;

                if (columnWidthMode === "custom") {
                    if (isNaN(result.columnWidth.input) || result.columnWidth.input <= 0) return;
                    var customColumnWidth = result.columnWidth.input;
                    for (var columnIndex = 0; columnIndex < targetTable.columns.length; columnIndex++) {
                        targetTable.columns[columnIndex].width = customColumnWidth;
                    }
                    return;
                }

                if (columnWidthMode === "naturalPlus") {
                    fitColumnsToContent(targetTable, 2);
                    originalAutoFitColumnWidths = [];
                    for (var columnIndex = 0; columnIndex < targetTable.columns.length; columnIndex++) {
                        originalAutoFitColumnWidths.push(targetTable.columns[columnIndex].width);
                    }
                    originalAutoFitTableWidth = targetTable.width;
                }

                if (widthMode === "fit") {
                    var textFrame = targetTable.parent;
                    while (textFrame.constructor.name !== "TextFrame" && textFrame.constructor.name !== "Story") {
                        textFrame = textFrame.parent;
                    }
                    if (textFrame.constructor.name !== "TextFrame") {
                        alert(getLabel("error.textFrameNotFound"));
                        return;
                    }
                    targetTableWidth = textFrame.geometricBounds[3] - textFrame.geometricBounds[1];
                } else if (widthMode === "custom") {
                    if (isNaN(result.tableWidth.input) || result.tableWidth.input <= 0) return;
                    targetTableWidth = result.tableWidth.input;
                } else if (widthMode === "auto") {
                    autoWidth = true;
                }

                // 2. 列幅「最終列のみ調整」は特別処理
                // 非最終列は現在の幅を維持し、最終列だけで表全体の幅に合わせます。
                if (columnWidthMode === "last") {
                    if (widthMode === "fit" || widthMode === "custom") {
                        var baseColumnWidths = currentPreviewColumnWidths || originalColumnWidths;
                        var totalOtherColumnsWidth = 0;
                        for (var columnIndex = 0; columnIndex < targetTable.columns.length - 1; columnIndex++) {
                            totalOtherColumnsWidth += baseColumnWidths[columnIndex];
                        }

                        var newLastColumnWidth = targetTableWidth - totalOtherColumnsWidth;
                        if (newLastColumnWidth <= 0) return;

                        for (var columnIndex = 0; columnIndex < targetTable.columns.length - 1; columnIndex++) {
                            targetTable.columns[columnIndex].width = baseColumnWidths[columnIndex];
                        }
                        targetTable.columns[targetTable.columns.length - 1].width = newLastColumnWidth;
                    }
                    return;
                }

                // 3. 通常フロー: 幅を適用
                if (widthMode === "fit" || widthMode === "custom") {
                    targetTable.width = targetTableWidth;
                } else if (autoWidth && columnWidthMode !== "naturalPlus") {
                    fitColumnsToContent(targetTable, 2);
                }

                // 4. 列幅処理
                if (columnWidthMode === "equal") {
                    equalizeColumns(targetTable);
                } else if (columnWidthMode === "natural") {
                    fitColumnsToContent(targetTable, 2);
                } else if (columnWidthMode === "naturalPlus") {
                    var widthDelta = targetTable.width - originalAutoFitTableWidth;
                    var deltaPerColumn = widthDelta / targetTable.columns.length;
                    for (var columnIndex = 0; columnIndex < targetTable.columns.length; columnIndex++) {
                        targetTable.columns[columnIndex].width = originalAutoFitColumnWidths[columnIndex] + deltaPerColumn;
                    }
                }
            }

            var rulerUnitInfo = getCurrentRulerUnitInfo();
            var currentColumnWidthValue = roundToOneDecimal(getFirstColumnWidth(targetTable));
            var currentTableWidthValue = roundToOneDecimal(targetTable.width);
            var result = showMultiPanelOptionDialog(getLabel("dialog.title"), [
                {
                    key: "tableWidth",
                    label: getLabel("panel.tableWidth"),
                    options: [
                        { value: "keep", label: getLabel("radio.widthKeep"), defaultSelected: true },
                        { value: "auto", label: getLabel("radio.widthAuto") },
                        { value: "fit", label: getLabel("radio.widthFit") },
                        { value: "custom", label: getLabel("radio.widthCustom"), input: { suffix: rulerUnitInfo.label, defaultValue: currentTableWidthValue } }
                    ]
                },
                {
                    key: "columnWidth",
                    label: getLabel("panel.columnWidth"),
                    options: [
                        { value: "equal", label: getLabel("radio.cellEqual"), defaultSelected: true },
                        { value: "natural", label: getLabel("radio.cellNatural"), linkedSelection: { panel: "tableWidth", value: "auto" } },
                        { value: "naturalPlus", label: getLabel("radio.cellNaturalPlus") },
                        { value: "last", label: getLabel("radio.cellLast"), linkedSelection: { panel: "tableWidth", value: "fit" } },
                        { value: "custom", label: getLabel("radio.cellCustom"), input: { suffix: rulerUnitInfo.label, defaultValue: currentColumnWidthValue } }
                    ]
                }
            ], applyTableWidthResult, revert, targetTable);

            if (result === null) {
                revert();
                restoreSelection(savedSelection);
                return;
            }

            /* プレビューを一度戻してから、確定分だけを 1 つの取り消しにまとめる
               / Revert the preview first so only the confirmed change forms the undo step */
            revert();
            app.doScript(function () {
                applyTableWidthResult(result);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel("undo.applyWidths"));
            restoreSelection(savedSelection);
        }

        /**
         * すべての列幅を均等にする
         * @param {Table} table 対象の表
         * @returns {void}
         */
        function equalizeColumns(table) {
            var avg = table.width / table.columns.length;
            for (var j = 0; j < table.columns.length; j++) {
                table.columns[j].width = avg;
            }
        }

        /**
         * 組版後の 1 行の幅を求める
         * @param {Line} line 対象の行
         * @returns {number} 行の幅
         */
        function getComposedLineWidth(line) {
            if (!line) return 0;
            try {
                var lineStart = line.insertionPoints[0].horizontalOffset;
                var lineEnd = line.insertionPoints[-1].horizontalOffset;
                return lineEnd - lineStart;
            } catch (e) {
                return 0;
            }
        }

        /**
         * セル内で最も長い行の幅を求める
         * @param {Cell} cell 対象のセル
         * @returns {number} 最大の行幅
         */
        function getMaxComposedLineWidthInCell(cell) {
            if (!cell || !cell.lines || cell.lines.length === 0) return 0;

            var cellLines = cell.lines;
            var maxContentWidth = 0;
            for (var lineIndex = 0; lineIndex < cellLines.length; lineIndex++) {
                var lineWidth = getComposedLineWidth(cellLines[lineIndex]);
                if (lineWidth > maxContentWidth) maxContentWidth = lineWidth;
            }
            return maxContentWidth;
        }

        /**
         * 各列の幅を内容に合わせて調整する
         * @param {Table} table 対象の表
         * @param {number} margin 内容に加える余白
         * @returns {Array<number>} 調整後の列幅
         */
        function fitColumnsToContent(table, margin) {
            if (margin === undefined) margin = 2;

            // 計測前に表を親フレームいっぱいまで広げ、列を均等化して余裕を作る
            var textFrame = table.parent;
            while (textFrame && textFrame.constructor.name !== "TextFrame" && textFrame.constructor.name !== "Story") {
                textFrame = textFrame.parent;
            }
            if (textFrame && textFrame.constructor.name === "TextFrame") {
                var parentFrameWidth = textFrame.geometricBounds[3] - textFrame.geometricBounds[1];
                if (parentFrameWidth > 0) {
                    table.width = parentFrameWidth;
                    var evenColumnWidth = parentFrameWidth / table.columns.length;
                    for (var equalIndex = 0; equalIndex < table.columns.length; equalIndex++) {
                        table.columns[equalIndex].width = evenColumnWidth;
                    }
                }
            }

            var columns = table.columns;
            for (var columnIndex = 0, columnCount = columns.length; columnIndex < columnCount; columnIndex++) {
                var columnCells = columns[columnIndex].cells;
                var contentWidths = [];
                for (var cellIndex = 0, cellCount = columnCells.length; cellIndex < cellCount; cellIndex++) {
                    if (columnCells[cellIndex].texts[0].contents === "") continue;

                    // 複数行に渡るセル（ハードリターンや折り返し）も想定し、全行中の最大幅を採用
                    contentWidths.push(getMaxComposedLineWidthInCell(columnCells[cellIndex]));
                }
                columns[columnIndex].rightInset = columns[columnIndex].leftInset = margin * 1.0;
                var padding = columns[columnIndex].rightInset + columns[columnIndex].leftInset;
                var lineWeight = columns[columnIndex].rightEdgeStrokeWeight * 0.5 + columns[columnIndex].leftEdgeStrokeWeight * 0.5;
                if (contentWidths.length > 0) {
                    columns[columnIndex].width = contentWidths.sort(function (a, b) { return b - a; })[0] + padding + lineWeight;
                }
            }
        }

        /**
         * 複数のラジオボタンパネルを持つ設定ダイアログを表示する
         * @param {string} title ダイアログのタイトル
         * @param {Array<object>} panels パネルの定義
         * @param {function} onApply プレビュー適用時に呼ぶ処理
         * @param {function} onRevert プレビュー復元時に呼ぶ処理
         * @param {Table} targetTable 対象の表
         * @returns {object|null} 各パネルの選択結果。キャンセル時は null
         */
        function showMultiPanelOptionDialog(title, panels, onApply, onRevert, targetTable) {
            var dlg = new Window("dialog", title + " " + SCRIPT_VERSION);
            setupWindow(dlg, 10);

            var previewToggleCount = 0;
            var panelStates = [];

            /**
             * パネルを横に並べる行グループを作る
             * @param {object} parent 追加先のコンテナ
             * @param {Array<object>} panels パネルの定義
             * @returns {Group} 行グループ
             */
            function buildPanelsRow(parent, panels) {
                var row = parent.add("group");
                setupRow(row, "fill", COLUMN_SPACING);
                row.alignChildren = ["fill", "top"];
                return row;
            }

            /**
             * 1 パネル分の状態オブジェクトを作る
             * @param {object} panelDef パネルの定義
             * @returns {object} パネルの状態
             */
            function createPanelState(panelDef) {
                return {
                    key: panelDef.key,
                    options: panelDef.options,
                    radios: [],
                    inputs: [],
                    rows: []
                };
            }

            /**
             * ∧∨・↑↓キーで増減したあとの処理。そのオプションのラジオを選び、手入力の確定と同じ onChange を通す
             * @param {EditText} editText 増減した入力欄
             * @param {object} state パネルの状態
             * @param {number} optionIndex 対象オプションの位置
             * @returns {void}
             */
            function onInputStepped(editText, state, optionIndex) {
                selectRadio(state, optionIndex);
                applyLinkedSelection(state.options[optionIndex]);
                if (typeof editText.onChange === "function") {
                    editText.onChange(); /* 依存関係の反映・表全体の幅の同期・プレビュー / dependencies, width sync and preview */
                }
            }

            /**
             * パネル内のオプション 1 行を組み立てる
             * @param {Panel} panel 追加先のパネル
             * @param {object} state パネルの状態
             * @param {object} option オプションの定義
             * @returns {void}
             */
            function buildOptionRow(panel, state, option) {
                var row = panel.add("group");
                row.alignChildren = ["left", "center"];
                state.rows.push(row);

                var rb = row.add("radiobutton", undefined, option.label);
                state.radios.push(rb);

                if (option.input) {
                    var optionIndex = state.radios.length - 1;
                    var et;
                    var stepperGroup = addStepperInputPair(row, function () { return et; }, {
                        min: 0,
                        onStep: function (numberInput) { onInputStepped(numberInput, state, optionIndex); }
                    });
                    et = stepperGroup.parent.add("edittext", undefined, String(option.input.defaultValue));
                    bindSteppedArrowKeys(et, stepperGroup);
                    et.characters = 4;
                    row.add("statictext", undefined, option.input.suffix || "");
                    state.inputs.push(et);
                } else {
                    state.inputs.push(null);
                }
            }

            /**
             * パネル 1 つを組み立てる
             * @param {object} parent 追加先のコンテナ
             * @param {object} panelDef パネルの定義
             * @returns {object} パネルの状態
             */
            function buildPanel(parent, panelDef) {
                var optionPanel = parent.add("panel", undefined, panelDef.label);
                setupPanel(optionPanel, 6);
                optionPanel.alignChildren = "left";

                var state = createPanelState(panelDef);
                var defaultIndex = 0;

                for (var i = 0; i < panelDef.options.length; i++) {
                    buildOptionRow(optionPanel, state, panelDef.options[i]);
                    if (panelDef.options[i].defaultSelected) defaultIndex = i;
                }

                state.radios[defaultIndex].value = true;
                return state;
            }

            /**
             * すべてのパネルを組み立てる
             * @param {object} parent 追加先のコンテナ
             * @param {Array<object>} panels パネルの定義
             * @returns {void}
             */
            function buildAllPanels(parent, panels) {
                var row = buildPanelsRow(parent, panels);
                for (var panelIndex = 0; panelIndex < panels.length; panelIndex++) {
                    panelStates.push(buildPanel(row, panels[panelIndex]));
                }
            }

            /**
             * 各パネルの選択結果をまとめて返す
             * @returns {object} パネルごとの選択結果
             */
            function collectDialogResult() {
                var result = {};
                for (var panelIndex = 0; panelIndex < panelStates.length; panelIndex++) {
                    var panelState = panelStates[panelIndex];
                    for (var optionIndex = 0; optionIndex < panelState.radios.length; optionIndex++) {
                        if (panelState.radios[optionIndex].value) {
                            var optionResult = { value: panelState.options[optionIndex].value };
                            if (panelState.inputs[optionIndex]) optionResult.input = parseFloat(panelState.inputs[optionIndex].text);
                            result[panelState.key] = optionResult;
                            break;
                        }
                    }
                }
                return result;
            }

            /**
             * 現在のダイアログ設定でプレビューを適用する
             * @returns {void}
             */
            function applyPreviewFromDialog() {
                if (onApply) onApply(collectDialogResult());
            }

            /**
             * 指定した位置のラジオボタンを選択状態にする
             * @param {object} state パネルの状態
             * @param {number} idx 選択する位置
             * @returns {void}
             */
            function selectRadio(state, idx) {
                for (var radioIndex = 0; radioIndex < state.radios.length; radioIndex++) {
                    state.radios[radioIndex].value = (radioIndex === idx);
                }
            }

            /**
             * キーからパネルの状態を探す
             * @param {string} panelKey パネルのキー
             * @returns {object|null} パネルの状態。見つからない場合は null
             */
            function findPanelState(panelKey) {
                for (var i = 0; i < panelStates.length; i++) {
                    if (panelStates[i].key === panelKey) return panelStates[i];
                }
                return null;
            }

            /**
             * 指定したパネルで、値に対応するオプションを選択する
             * @param {string} panelKey パネルのキー
             * @param {string} targetValue 選択したい値
             * @returns {void}
             */
            function selectInPanel(panelKey, targetValue) {
                var targetState = findPanelState(panelKey);
                if (!targetState) return;

                for (var i = 0; i < targetState.options.length; i++) {
                    if (targetState.options[i].value === targetValue) {
                        selectRadio(targetState, i);
                        return;
                    }
                }
            }

            /**
             * 連動指定のあるオプションを他パネルへ反映する
             * @param {object} option 選択されたオプション
             * @returns {void}
             */
            function applyLinkedSelection(option) {
                if (option && option.linkedSelection) {
                    selectInPanel(option.linkedSelection.panel, option.linkedSelection.value);
                }
            }

            /**
             * 指定したオプションの有効／無効を切り替える
             * @param {object} state パネルの状態
             * @param {string} optionValue 対象の値
             * @param {boolean} isEnabled 有効にするなら true
             * @returns {void}
             */
            function setOptionEnabled(state, optionValue, isEnabled) {
                if (!state) return;

                for (var optionIndex = 0; optionIndex < state.options.length; optionIndex++) {
                    if (state.options[optionIndex].value === optionValue) {
                        if (state.radios[optionIndex]) state.radios[optionIndex].enabled = isEnabled;
                        if (state.inputs[optionIndex]) state.inputs[optionIndex].enabled = isEnabled;
                        if (state.inputs[optionIndex]) state.inputs[optionIndex].readonly = !isEnabled;
                        if (state.rows[optionIndex] && state.rows[optionIndex].enabled !== isEnabled) {
                            state.rows[optionIndex].enabled = isEnabled;
                            /* 行の無効化は∧∨の自作描画に届かないので描き直す / redraw the custom-drawn stepper for the row's state */
                            if (dlg.visible) redrawSteppersIn(state.rows[optionIndex]);
                        }
                        return;
                    }
                }
            }

            /**
             * パネルで選択中のオプション値を取得する
             * @param {object} state パネルの状態
             * @returns {string|null} 選択中の値。なければ null
             */
            function getSelectedOptionValue(state) {
                if (!state) return null;
                for (var optionIndex = 0; optionIndex < state.radios.length; optionIndex++) {
                    if (state.radios[optionIndex].value) {
                        return state.options[optionIndex].value;
                    }
                }
                return null;
            }

            /**
             * 依存関係で無効化したオプションを既定状態に戻す
             * @param {object} tableWidthState 表全体の幅パネルの状態
             * @param {object} columnWidthState 列の幅パネルの状態
             * @returns {void}
             */
            function resetDependentUI(tableWidthState, columnWidthState) {
                setOptionEnabled(tableWidthState, "keep", true);
                setOptionEnabled(tableWidthState, "auto", true);
                setOptionEnabled(tableWidthState, "fit", true);
                setOptionEnabled(tableWidthState, "custom", true);
                setOptionEnabled(columnWidthState, "equal", true);
                setOptionEnabled(columnWidthState, "natural", true);
                setOptionEnabled(columnWidthState, "naturalPlus", true);
                setOptionEnabled(columnWidthState, "last", true);
                setOptionEnabled(columnWidthState, "custom", true);
            }

            /**
             * 表全体の幅の選択に応じて列の幅パネルを制御する
             * @param {string} selectedTableWidthValue 表全体の幅の選択値
             * @param {string} selectedColumnWidthValue 列の幅の選択値
             * @param {object} columnWidthState 列の幅パネルの状態
             * @returns {void}
             */
            function applyTableWidthDependencies(selectedTableWidthValue, selectedColumnWidthValue, columnWidthState) {
            }

            /**
             * 列の幅の選択に応じて表全体の幅パネルを制御する
             * @param {string} selectedColumnWidthValue 列の幅の選択値
             * @param {object} tableWidthState 表全体の幅パネルの状態
             * @returns {void}
             */
            function applyColumnWidthDependencies(selectedColumnWidthValue, tableWidthState) {
                if (selectedColumnWidthValue === "equal") {
                    setOptionEnabled(tableWidthState, "auto", false);
                }

                if (selectedColumnWidthValue === "last") {
                    selectInPanel("tableWidth", "fit");
                    setOptionEnabled(tableWidthState, "keep", false);
                    setOptionEnabled(tableWidthState, "auto", false);
                    setOptionEnabled(tableWidthState, "custom", false);
                }

                if (selectedColumnWidthValue === "custom") {
                    setOptionEnabled(tableWidthState, "keep", false);
                    setOptionEnabled(tableWidthState, "auto", false);
                    setOptionEnabled(tableWidthState, "fit", false);
                    setOptionEnabled(tableWidthState, "custom", false);
                }
            }

            /**
             * 列の幅の指定値から、表全体の幅の指定値を更新する
             * @returns {void}
             */
            function syncTableWidthCustomInputFromColumnWidth() {
                var tableWidthState = findPanelState("tableWidth");
                var columnWidthState = findPanelState("columnWidth");
                if (!tableWidthState || !columnWidthState) return;

                var tableWidthCustomIndex = -1;
                for (var optionIndex = 0; optionIndex < tableWidthState.options.length; optionIndex++) {
                    if (tableWidthState.options[optionIndex].value === "custom") {
                        tableWidthCustomIndex = optionIndex;
                        break;
                    }
                }

                var columnWidthCustomIndex = -1;
                for (var optionIndex = 0; optionIndex < columnWidthState.options.length; optionIndex++) {
                    if (columnWidthState.options[optionIndex].value === "custom") {
                        columnWidthCustomIndex = optionIndex;
                        break;
                    }
                }

                if (tableWidthCustomIndex < 0 || columnWidthCustomIndex < 0) return;
                if (!tableWidthState.inputs[tableWidthCustomIndex] || !columnWidthState.inputs[columnWidthCustomIndex]) return;
                if (!columnWidthState.radios[columnWidthCustomIndex].value) return;
                if (!targetTable || !targetTable.columns || targetTable.columns.length <= 0) return;

                var columnWidthValue = Number(columnWidthState.inputs[columnWidthCustomIndex].text);
                if (isNaN(columnWidthValue) || columnWidthValue <= 0) return;

                var syncedTableWidthValue = roundToOneDecimal(columnWidthValue * targetTable.columns.length);
                tableWidthState.inputs[tableWidthCustomIndex].text = String(syncedTableWidthValue);
                try {
                    dlg.layout.layout(true);
                    dlg.update();
                } catch (e) { }
            }

            /**
             * パネル間の依存関係をまとめて反映する
             * @returns {void}
             */
            function updateDependentUI() {
                var tableWidthState = findPanelState("tableWidth");
                var columnWidthState = findPanelState("columnWidth");
                if (!tableWidthState || !columnWidthState) return;

                var selectedTableWidthValue = getSelectedOptionValue(tableWidthState);
                var selectedColumnWidthValue = getSelectedOptionValue(columnWidthState);

                resetDependentUI(tableWidthState, columnWidthState);
                applyTableWidthDependencies(selectedTableWidthValue, selectedColumnWidthValue, columnWidthState);
                applyColumnWidthDependencies(selectedColumnWidthValue, tableWidthState);
            }

            /**
             * 1 パネル分のイベントを結び付ける
             * @param {object} state パネルの状態
             * @returns {void}
             */
            function bindStateEvents(state) {
                for (var radioIndex = 0; radioIndex < state.radios.length; radioIndex++) {
                    (function (optionIndex) {
                        state.radios[optionIndex].onClick = function () {
                            selectRadio(state, optionIndex);
                            applyLinkedSelection(state.options[optionIndex]);
                            updateDependentUI();
                            syncTableWidthCustomInputFromColumnWidth();
                            applyPreviewFromDialog();
                        };
                    })(radioIndex);
                }

                for (var inputIndex = 0; inputIndex < state.inputs.length; inputIndex++) {
                    if (state.inputs[inputIndex]) {
                        (function (optionIndex) {
                            state.inputs[optionIndex].onChange = function () {
                                updateDependentUI();
                                syncTableWidthCustomInputFromColumnWidth();
                                applyPreviewFromDialog();
                            };
                            state.inputs[optionIndex].onChanging = function () {
                                selectRadio(state, optionIndex);
                                applyLinkedSelection(state.options[optionIndex]);
                                updateDependentUI();
                                syncTableWidthCustomInputFromColumnWidth();
                            };
                        })(inputIndex);
                    }
                }
            }

            /**
             * すべてのパネルのイベントを結び付ける
             * @returns {void}
             */
            function bindAllStateEvents() {
                for (var panelIndex = 0; panelIndex < panelStates.length; panelIndex++) {
                    bindStateEvents(panelStates[panelIndex]);
                }
            }

            /**
             * ダイアログ下部のボタン行を組み立てる（左：画面モード／右：キャンセル・OK）
             * 画面モードを切り替えた回数を数え、閉じたあとで元のモードに戻せるようにする
             * @param {object} parent 追加先のコンテナ
             * @returns {void}
             */
            function addDialogButtons(parent) {
                var buttonRow = addButtonRow(parent);
                addScreenModeButton(buttonRow.leftGroup, function () {
                    previewToggleCount++;
                });
                var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
                var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
                centerButtonRowIfRightOnly(buttonRow);
            }

            buildAllPanels(dlg, panels);

            bindAllStateEvents();
            updateDependentUI();
            syncTableWidthCustomInputFromColumnWidth();

            addDialogButtons(dlg);

            applyPreviewFromDialog();

            var shown = dlg.show();

            if (previewToggleCount % 2 === 1) {
                togglePreviewScreenMode();
            }

            if (shown !== 1) return null;
            return collectDialogResult();
        }

        runAdjustTableWidth();
    })();