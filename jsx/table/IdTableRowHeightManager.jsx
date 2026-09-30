#target indesign

/*

### 概要

選択した表の行の高さを、範囲（選択範囲／ストーリー／ドキュメント）と対象行を指定しながらプレビュー付きで設定します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableRowHeightManager.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n9f95f8e98db6

### Overview

Sets the row heights of the selected table with a live preview, choosing the scope (selection, story or document) and which rows to target.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableRowHeightManager.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTableRowHeightManager";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableRowHeightManager.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableRowHeightManager.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n9f95f8e98db6"; /* 紹介記事 / article URL */

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

        // =========================================
        // ユーザー設定 / User settings
        // =========================================

        /* 行の高さの下限（pt）。InDesign の最小行高（約 0.0139 inch ≒ 1pt）に合わせる
           / Minimum row height in points; InDesign's own minimum is about 0.0139 inch (≈ 1pt) */
        var MIN_ROW_HEIGHT_PT = 1.0008;

        /* 初期値を行から求められないときの行の高さ（pt） / Fallback row height (pt) when none can be read from the rows */
        var FALLBACK_ROW_HEIGHT_PT = 15;

        // =========================================
        // レイアウト設定 / Layout settings
        // =========================================

        /* 行高入力欄の文字数 / Character width of the row-height field */
        var ROW_HEIGHT_INPUT_CHARACTERS = 5;

        /* 進捗バーの幅 / Width of the progress bar */
        var PROGRESS_BAR_WIDTH = 320;

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
                title: { ja: "行の高さを設定", en: "Set Row Height" }
            },
            panel: {
                scope: { ja: "範囲", en: "Scope" },
                target: { ja: "対象", en: "Target" },
                rowHeight: { ja: "行の高さ", en: "Row Height" },
                options: { ja: "オプション", en: "Options" }
            },
            radio: {
                scopeDocument: { ja: "ドキュメント", en: "Document" },
                scopeStory: { ja: "ストーリー", en: "Story" },
                scopeSelection: { ja: "選択範囲", en: "Selection" },
                targetWhole: { ja: "表全体", en: "Whole Table" },
                targetBody: { ja: "表全体（ヘッダー行を除く）", en: "Whole Table (Except Header Rows)" },
                targetSelected: { ja: "選択した行のみ", en: "Selected Rows Only" },
                heightMinimum: { ja: "最小限度", en: "At Least" },
                heightSpecified: { ja: "指定値を使用", en: "Exactly" }
            },
            checkbox: {
                fitFrameToContent: { ja: "フレームを内容に合わせる", en: "Fit Frame to Content" }
            },
            button: {
                screenModePreview: { ja: "プレビュー", en: "Preview" },
                screenModeNormal: { ja: "標準モード", en: "Normal Mode" },
                ok: { ja: "OK", en: "OK" },
                cancel: { ja: "キャンセル", en: "Cancel" }
            },
            tooltip: {
                scopeDocument: {
                    ja: "ドキュメント内のすべての表を対象にします。",
                    en: "Targets all tables in the document."
                },
                scopeStory: {
                    ja: "選択した表と同じストーリーにあるすべての表を対象にします。",
                    en: "Targets all tables in the same story as the selected table."
                },
                scopeSelection: {
                    ja: "選択した表だけを対象にします。",
                    en: "Targets only the selected table."
                },
                targetBody: {
                    ja: "ヘッダー行を除いた行を対象にします。",
                    en: "Targets all rows except header rows."
                },
                targetSelected: {
                    ja: "選択した行だけを対象にします。範囲が「選択範囲」で、行を選択しているときに使えます。",
                    en: "Targets only the selected rows. Available when the scope is Selection and rows are selected."
                },
                heightMinimum: {
                    ja: "各行を最小の高さ（約1pt）に設定し、内容に合わせて自動的に伸びるようにします。",
                    en: "Sets each row to the minimum height (about 1 pt) so it grows to fit its content."
                },
                heightSpecified: {
                    ja: "入力した値を行の高さとして設定します。",
                    en: "Uses the entered value as the row height."
                },
                heightInput: {
                    ja: "↑↓キーで次の整数へ、Shift＋↑↓で10の倍数へ、Option＋↑↓で0.1ずつ増減します。単位はドキュメントの縦方向の単位です。",
                    en: "Up/Down steps to the next whole number, Shift+Up/Down to multiples of 10, Option+Up/Down by 0.1. Unit follows the document's vertical units."
                },
                screenMode: {
                    ja: "ドキュメントウィンドウの画面モードを、標準モードとプレビューで切り替えます。",
                    en: "Switches the document window between the Normal and Preview screen modes."
                },
                fitFrameToContent: {
                    ja: "行の高さを変えたあと、表を含むテキストフレームの高さを内容に合わせます。",
                    en: "After changing the row heights, fits the height of the text frame containing the table to its content."
                },
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
                noTable: {
                    ja: "表、セル、または表を含むテキストフレームを選択してください。",
                    en: "Please select a table, cell, or a text frame containing a table."
                },
                multipleTables: {
                    ja: "複数の表が選択されています。1つの表だけを選択してください。",
                    en: "Multiple tables are selected. Please select only one table."
                },
                invalidNumber: {
                    ja: "正の数値を入力してください。",
                    en: "Please enter a positive number."
                }
            },
            progress: {
                title: { ja: "行の高さを適用中", en: "Applying Row Heights" }
            },
            undo: {
                applyRowHeight: { ja: "行の高さを設定", en: "Set Row Height" }
            }
        };

        // =========================================
        // 単位 / Units
        // =========================================

        /**
         * 単位の表示ラベルと 1 単位あたりのポイント数を返す
         * @param {MeasurementUnits} unit 対象の単位
         * @returns {object} { label: string, pointsPerUnit: number }
         */
        function getUnitInfo(unit) {
            switch (unit) {
                case MeasurementUnits.MILLIMETERS: return { label: "mm", pointsPerUnit: 72 / 25.4 };
                case MeasurementUnits.CENTIMETERS: return { label: "cm", pointsPerUnit: 72 / 2.54 };
                case MeasurementUnits.INCHES:
                case MeasurementUnits.INCHES_DECIMAL: return { label: "inch", pointsPerUnit: 72 };
                case MeasurementUnits.POINTS:
                case MeasurementUnits.AMERICAN_POINTS: return { label: "pt", pointsPerUnit: 1 };
                case MeasurementUnits.PICAS: return { label: "pica", pointsPerUnit: 12 };
                case MeasurementUnits.PIXELS: return { label: "px", pointsPerUnit: 1 };
                case MeasurementUnits.AGATES: return { label: "ag", pointsPerUnit: 5.5 };
                case MeasurementUnits.Q: return { label: "Q", pointsPerUnit: 72 / 25.4 * 0.25 };
                case MeasurementUnits.HA: return { label: "H", pointsPerUnit: 72 / 25.4 * 0.25 };
                case MeasurementUnits.CICEROS: return { label: "cicero", pointsPerUnit: 1 };
                case MeasurementUnits.BAI: return { label: "倍", pointsPerUnit: 1 };
                case MeasurementUnits.U: return { label: "U", pointsPerUnit: 1 };
                default: return { label: "", pointsPerUnit: 1 };
            }
        }

        /**
         * 入力欄に表示する数値を整形する
         * @param {number} value 表示する数値
         * @returns {string} 整形した文字列
         */
        function formatDisplayValue(value) {
            if (Math.abs(value - Math.round(value)) < 0.0001) return String(Math.round(value));
            return String(Math.round(value * 1000) / 1000);
        }

        // =========================================
        // ユーティリティ / Utilities
        // =========================================

        /**
         * コレクションや選択を通常の配列に写す
         * @param {object} collection length と添字アクセスを持つオブジェクト
         * @param {number} [startIndex] 写し始める位置。省略時は 0
         * @returns {Array} 要素の配列
         */
        function collectionToArray(collection, startIndex) {
            var items = [];
            for (var i = startIndex || 0; i < collection.length; i++) items.push(collection[i]);
            return items;
        }

        /**
         * 親方向にたどって表とその経路情報を求める
         * @param {object} startItem 起点となるオブジェクト
         * @returns {object} { table, cell, row, sourceItem }。表が無ければ table は null
         */
        function walkUpToTable(startItem) {
            var currentNode = startItem;
            var foundCell = null;
            var foundRow = null;
            var foundTable = null;
            while (currentNode) {
                /* Application まで遡ると parent が無く例外になる / Reaching Application throws on .parent */
                try {
                    if (currentNode instanceof Cell) { foundCell = currentNode; foundTable = currentNode.parent; break; }
                    if (currentNode instanceof Row) { foundRow = currentNode; foundTable = currentNode.parent; break; }
                    if (currentNode instanceof Table) { foundTable = currentNode; break; }
                    if (currentNode instanceof TextFrame && currentNode.tables.length > 0) { foundTable = currentNode.tables[0]; break; }
                    currentNode = currentNode.parent;
                } catch (e) { break; }
            }
            return { table: foundTable, cell: foundCell, row: foundRow, sourceItem: startItem };
        }

        /**
         * 行インデックスを重複なく集めるコレクタを作る
         * @returns {object} 行インデックスを追加・取得するオブジェクト
         */
        function createRowIndexCollector() {
            var rowIndexMap = {};
            var hasSpecificRows = false;

            /**
             * 行インデックスを 1 つ追加する
             * @param {number} rowIndex 行インデックス
             * @returns {void}
             */
            function addRowIndex(rowIndex) {
                if (rowIndex === undefined || rowIndex === null) return;
                rowIndexMap[rowIndex] = true;
                hasSpecificRows = true;
            }

            /**
             * 要素ごとに行を求めて、そのインデックスを追加する
             * @param {Array} items 行またはセルの配列
             * @param {Function} getRow 要素から行を返す関数
             * @returns {boolean} 1 つでも追加できたら true
             */
            function addRowsFrom(items, getRow) {
                if (!items || items.length === 0) return false;
                var added = false;
                for (var i = 0; i < items.length; i++) {
                    /* 無効になった参照は読めないので飛ばす / Skip references that became invalid */
                    try {
                        addRowIndex(getRow(items[i]).index);
                        added = true;
                    } catch (e) { }
                }
                return added;
            }

            /**
             * 集めた行インデックスを昇順の配列で返す
             * @returns {Array<number>} 昇順の行インデックス
             */
            function toSortedIndices() {
                var rowIndices = [];
                for (var k in rowIndexMap) {
                    if (rowIndexMap.hasOwnProperty(k)) rowIndices.push(parseInt(k, 10));
                }
                rowIndices.sort(function (a, b) { return a - b; });
                return rowIndices;
            }

            return {
                addRowIndex: addRowIndex,
                addRowsFromRows: function (rows) { return addRowsFrom(rows, function (row) { return row; }); },
                addRowsFromCells: function (cells) { return addRowsFrom(cells, function (cell) { return cell.parentRow; }); },
                toSortedIndices: toSortedIndices,
                hasSpecificRows: function () { return hasSpecificRows; }
            };
        }

        /**
         * 選択項目から行インデックスを収集する
         * @param {object} walkResult walkUpToTable の結果
         * @param {object} collector 行インデックスのコレクタ
         * @returns {void}
         */
        function collectRowIndicesFromItem(walkResult, collector) {
            var sourceItem = walkResult.sourceItem;

            /* 選択項目が持つ行・セルのコレクションから行インデックスを集める / Collect row indices from any row/cell collections the item exposes */
            var rangeSources = [
                { propName: "parentRows", addRows: collector.addRowsFromRows },
                { propName: "rows", addRows: collector.addRowsFromRows },
                { propName: "parentCells", addRows: collector.addRowsFromCells },
                { propName: "cells", addRows: collector.addRowsFromCells }
            ];
            var addedFromRange = false;
            for (var i = 0; i < rangeSources.length; i++) {
                /* 選択の型によってはプロパティ自体が無く例外になる / Some selection types lack the property and throw */
                try {
                    var collection = sourceItem && sourceItem[rangeSources[i].propName];
                    if (collection && collection.length > 0) {
                        addedFromRange = rangeSources[i].addRows(collection.everyItem().getElements()) || addedFromRange;
                    }
                } catch (e) { }
            }
            if (addedFromRange) return;

            if (walkResult.row) { collector.addRowIndex(walkResult.row.index); return; }
            if (walkResult.cell) { collector.addRowIndex(walkResult.cell.parentRow.index); return; }

            /* parentRow を持たない選択もある / Not every selection has parentRow */
            try {
                if (sourceItem && sourceItem.parentRow) collector.addRowIndex(sourceItem.parentRow.index);
            } catch (e) { }
        }

        /**
         * 選択から対象の表と選択行インデックスを特定する
         * @param {Array} selection 選択オブジェクトの配列
         * @returns {object|null} { table, rowIndices } または { error }。特定できない場合は null
         */
        function resolveTargetFromSelection(selection) {
            if (!selection || selection.length === 0) return null;

            var table = null;
            var collector = createRowIndexCollector();

            for (var i = 0; i < selection.length; i++) {
                var walkResult = walkUpToTable(selection[i]);
                if (!walkResult.table) continue;

                if (!table) {
                    table = walkResult.table;
                } else if (!isSameTable(table, walkResult.table)) {
                    return { error: "multipleTables" };
                }

                collectRowIndicesFromItem(walkResult, collector);
            }

            if (!table) return null;

            var rowIndices = null;
            if (collector.hasSpecificRows()) {
                rowIndices = collector.toSortedIndices();

                /* すべての行が選択されている場合は表全体として扱う / Treat as whole table if all rows are selected */
                if (rowIndices.length === table.rows.length) {
                    rowIndices = null;
                }
            }

            return { table: table, rowIndices: rowIndices };
        }

        /**
         * ドキュメント内のすべての表を集める
         * @returns {Array<Table>} 表の配列
         */
        function collectDocumentTables() {
            var tables = [];
            var stories = app.activeDocument.stories;
            for (var i = 0; i < stories.length; i++) {
                tables = tables.concat(collectionToArray(stories[i].tables));
            }
            return tables;
        }

        /**
         * 指定した範囲に含まれる表を求める
         * @param {string} scope "document" / "story" / "selection"
         * @param {Table} baseTable 基準となる表
         * @returns {Array<Table>} 対象の表の配列
         */
        function resolveScopeTables(scope, baseTable) {
            if (scope === "document") return collectDocumentTables();
            if (scope === "story") return collectionToArray(baseTable.parentStory.tables);
            return [baseTable];
        }

        // =========================================
        // 行の取得 / Row lookup
        // =========================================

        /**
         * 指定したインデックスの行を取得する
         * @param {Table} table 対象の表
         * @param {Array<number>} rowIndices 行インデックスの配列
         * @returns {Array<Row>} 行の配列
         */
        function getRowsByIndices(table, rowIndices) {
            var rows = [];
            for (var i = 0; i < rowIndices.length; i++) {
                var rowIndex = rowIndices[i];
                if (rowIndex >= 0 && rowIndex < table.rows.length) rows.push(table.rows[rowIndex]);
            }
            return rows;
        }

        /**
         * 対象モードに応じて処理する行を求める
         * @param {Table} table 対象の表
         * @param {Array<number>|null} selectedRowIndices 選択行のインデックス
         * @param {string} targetMode "whole" / "body" / "selected"
         * @returns {Array<Row>} 対象の行
         */
        function resolveTargetRows(table, selectedRowIndices, targetMode) {
            var hasSelectedRows = !!(selectedRowIndices && selectedRowIndices.length > 0);
            if (targetMode === "selected" && hasSelectedRows) return getRowsByIndices(table, selectedRowIndices);
            if (targetMode === "body") return collectionToArray(table.rows, table.headerRowCount);
            return collectionToArray(table.rows);
        }

        /**
         * 範囲と対象モードから、表ごとの対象行をまとめて求める
         * 選択行の指定は基準の表にだけ効き、ほかの表は表全体として扱う
         * @param {Array<Table>} tables 対象の表
         * @param {Table} baseTable 基準となる表
         * @param {Array<number>|null} selectedRowIndices 基準の表の選択行
         * @param {string} targetMode "whole" / "body" / "selected"
         * @returns {Array<Array<Row>>} 表ごとの行の配列
         */
        function collectTargetRowSets(tables, baseTable, selectedRowIndices, targetMode) {
            var rowSets = [];
            for (var i = 0; i < tables.length; i++) {
                var rowIndicesForTable = isSameTable(tables[i], baseTable) ? selectedRowIndices : null;
                rowSets.push(resolveTargetRows(tables[i], rowIndicesForTable, targetMode));
            }
            return rowSets;
        }

        /**
         * ダイアログの初期値に使う行の高さを求める
         * 対象行がそろっていればその値、ばらばらなら本文行の平均
         * @param {Table} table 対象の表
         * @param {Array<number>|null} selectedRowIndices 選択行のインデックス
         * @returns {number} 初期値（pt）
         */
        function getInitialRowHeight(table, selectedRowIndices) {
            var targetRows = resolveTargetRows(table, selectedRowIndices, "selected");
            if (targetRows.length > 0) {
                var firstHeight = targetRows[0].height;
                var isUniform = true;
                for (var i = 1; i < targetRows.length; i++) {
                    if (Math.abs(targetRows[i].height - firstHeight) > 0.01) { isUniform = false; break; }
                }
                if (isUniform) return firstHeight;
            }

            var bodyRows = resolveTargetRows(table, null, "body");
            if (bodyRows.length === 0) return FALLBACK_ROW_HEIGHT_PT;
            var totalHeight = 0;
            for (var j = 0; j < bodyRows.length; j++) totalHeight += bodyRows[j].height;
            return totalHeight / bodyRows.length;
        }

        // =========================================
        // 行の高さ操作 / Row height operations
        // =========================================

        /**
         * 行の高さを一時保存して復元できるストアを作る
         * @returns {object} 保存と復元を行うオブジェクト
         */
        function createRowHeightSnapshotStore() {
            var entries = [];
            var frameEntries = [];

            /**
             * まだ控えていない表の行高を控える
             * @param {Table} table 対象の表
             * @returns {void}
             */
            function rememberTable(table) {
                for (var i = 0; i < entries.length; i++) {
                    if (entries[i].table === table) return;
                }
                var rows = collectionToArray(table.rows);
                var heights = [];
                for (var j = 0; j < rows.length; j++) heights.push(rows[j].height);
                entries.push({ table: table, rows: rows, heights: heights });
            }

            /**
             * まだ控えていないフレームの寸法を控える
             * @param {TextFrame} frame 対象のフレーム
             * @returns {void}
             */
            function rememberFrame(frame) {
                for (var i = 0; i < frameEntries.length; i++) {
                    if (frameEntries[i].frame === frame) return;
                }
                frameEntries.push({ frame: frame, bounds: frame.geometricBounds });
            }

            /**
             * 控えておいたすべての行高とフレームの寸法を元に戻す
             * @returns {void}
             */
            function restoreAll() {
                for (var i = 0; i < entries.length; i++) {
                    for (var j = 0; j < entries[i].rows.length; j++) {
                        entries[i].rows[j].height = entries[i].heights[j];
                    }
                }
                for (var k = 0; k < frameEntries.length; k++) {
                    frameEntries[k].frame.geometricBounds = frameEntries[k].bounds;
                }
            }

            return { rememberTable: rememberTable, rememberFrame: rememberFrame, restoreAll: restoreAll };
        }

        /**
         * 表を含むテキストフレームを求める
         * @param {Table} table 対象の表
         * @returns {TextFrame|null} テキストフレーム。パス上のテキストなどでは null
         */
        function getParentTextFrame(table) {
            var frames = table.storyOffset.parentTextFrames;
            var parentFrame = (frames.length > 0) ? frames[0] : table.parentStory.textContainers[0];
            return (parentFrame instanceof TextFrame) ? parentFrame : null;
        }

        /**
         * テキストフレームを内容に合わせる
         * @param {TextFrame} frame 対象のフレーム
         * @returns {void}
         */
        function fitFrameToContent(frame) {
            var story = frame.parentStory;
            story.recompose();
            /* ロックされたフレームなどは fit が例外になる / fit throws on locked frames and the like */
            try {
                frame.fit(FitOptions.FRAME_TO_CONTENT);
            } catch (e) { return; }
            story.recompose();
        }

        // =========================================
        // 選択の控えと復元 / Selection snapshot
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

        // =========================================
        // プログレス / Progress
        // =========================================

        /**
         * 進捗表示用のパレットを作る
         * @param {string} title タイトル
         * @param {number} maxValue 進捗の最大値
         * @returns {object} 更新と終了を行うオブジェクト
         */
        function createProgressBar(title, maxValue) {
            var progressWin = new Window("palette", title);
            setupWindow(progressWin);

            var progressLabel = progressWin.add("statictext", undefined, "");
            progressLabel.preferredSize.width = PROGRESS_BAR_WIDTH;

            var progressControl = progressWin.add("progressbar", undefined, 0, maxValue);
            progressControl.preferredSize = [PROGRESS_BAR_WIDTH, 12];

            progressWin.show();

            return {
                update: function (value, text) {
                    progressControl.value = value;
                    if (text !== undefined) progressLabel.text = text;
                    progressWin.update();
                },
                close: function () {
                    progressWin.close();
                }
            };
        }

        // =========================================
        // ダイアログ部品 / Dialog parts
        // =========================================

        /**
         * パネルにラジオボタンを並べる
         * @param {Panel} panel 追加先のパネル
         * @param {Array<object>} radioDefs { labelKey, tooltipKey } の配列。tooltipKey は省略可
         * @returns {Array<RadioButton>} 追加したラジオボタン
         */
        function addRadioButtons(panel, radioDefs) {
            var radios = [];
            for (var i = 0; i < radioDefs.length; i++) {
                var radio = panel.add("radiobutton", undefined, getLabel(radioDefs[i].labelKey));
                if (radioDefs[i].tooltipKey) radio.helpTip = getLabel(radioDefs[i].tooltipKey);
                radios.push(radio);
            }
            return radios;
        }

        /**
         * ラジオボタンだけを縦に並べるパネルを作る
         * @param {Group} parentGroup 追加先
         * @param {string} titleKey パネル見出しのラベルキー
         * @returns {Panel} 作ったパネル
         */
        function addRadioPanel(parentGroup, titleKey) {
            var radioPanel = parentGroup.add("panel", undefined, getLabel(titleKey));
            setupPanel(radioPanel, 6);
            radioPanel.alignChildren = "left";
            return radioPanel;
        }

        /**
         * 「範囲」パネルを作る
         * @param {Group} parentGroup 追加先
         * @returns {object} { documentRadio, storyRadio, selectionRadio }
         */
        function addScopePanel(parentGroup) {
            var radios = addRadioButtons(addRadioPanel(parentGroup, 'panel.scope'), [
                { labelKey: 'radio.scopeDocument', tooltipKey: 'tooltip.scopeDocument' },
                { labelKey: 'radio.scopeStory', tooltipKey: 'tooltip.scopeStory' },
                { labelKey: 'radio.scopeSelection', tooltipKey: 'tooltip.scopeSelection' }
            ]);
            radios[2].value = true;
            return { documentRadio: radios[0], storyRadio: radios[1], selectionRadio: radios[2] };
        }

        /**
         * 「対象」パネルを作る
         * @param {Group} parentGroup 追加先
         * @param {boolean} hasSelectedRows 行が選択されているか
         * @returns {object} { wholeRadio, bodyRadio, selectedRadio }
         */
        function addTargetPanel(parentGroup, hasSelectedRows) {
            var radios = addRadioButtons(addRadioPanel(parentGroup, 'panel.target'), [
                { labelKey: 'radio.targetWhole' },
                { labelKey: 'radio.targetBody', tooltipKey: 'tooltip.targetBody' },
                { labelKey: 'radio.targetSelected', tooltipKey: 'tooltip.targetSelected' }
            ]);
            radios[2].enabled = hasSelectedRows;
            radios[hasSelectedRows ? 2 : 0].value = true;
            return { wholeRadio: radios[0], bodyRadio: radios[1], selectedRadio: radios[2] };
        }

        /**
         * 「行の高さ」パネルを作る
         * @param {Group} parentGroup 追加先
         * @param {string} initialText 入力欄の初期値
         * @param {string} unitLabel 単位の表示
         * @param {object} stepOptions ∧∨・↑↓キーの増減設定（addStepper() に渡す）
         * @returns {object} { minimumRadio, specifiedRadio, heightInput, heightStepper }
         */
        function addRowHeightPanel(parentGroup, initialText, unitLabel, stepOptions) {
            var rowHeightPanel = addRadioPanel(parentGroup, 'panel.rowHeight');
            var radios = addRadioButtons(rowHeightPanel, [
                { labelKey: 'radio.heightMinimum', tooltipKey: 'tooltip.heightMinimum' },
                { labelKey: 'radio.heightSpecified', tooltipKey: 'tooltip.heightSpecified' }
            ]);
            radios[0].value = true;

            var heightInputRow = rowHeightPanel.add("group");
            setupRow(heightInputRow, "left", 6);
            heightInputRow.alignChildren = ["left", "center"];
            /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
            var stepperInputGroup = heightInputRow.add("group");
            stepperInputGroup.orientation = "row";
            stepperInputGroup.alignChildren = ["left", "center"];
            stepperInputGroup.spacing = 0;
            stepperInputGroup.margins = 0;
            var heightInput;
            var heightStepper = addStepper(stepperInputGroup, function () { return heightInput; }, stepOptions);
            heightInput = stepperInputGroup.add("edittext", undefined, initialText);
            heightInput.characters = ROW_HEIGHT_INPUT_CHARACTERS;
            heightInput.helpTip = getLabel('tooltip.heightInput');
            bindSteppedArrowKeys(heightInput, heightStepper);
            heightInputRow.add("statictext", undefined, unitLabel);

            return { minimumRadio: radios[0], specifiedRadio: radios[1], heightInput: heightInput, heightStepper: heightStepper };
        }

        /**
         * 「オプション」パネルを作る
         * @param {Group} parentGroup 追加先
         * @param {boolean} canFitFrame フレームを内容に合わせられるか
         * @returns {Checkbox} ［フレームを内容に合わせる］チェックボックス
         */
        function addOptionsPanel(parentGroup, canFitFrame) {
            var optionsPanel = parentGroup.add("panel", undefined, getLabel('panel.options'));
            setupPanel(optionsPanel, 6);
            optionsPanel.alignChildren = ["left", "center"];
            var fitFrameCheckbox = optionsPanel.add("checkbox", undefined, getLabel('checkbox.fitFrameToContent'));
            fitFrameCheckbox.helpTip = getLabel('tooltip.fitFrameToContent');
            fitFrameCheckbox.value = canFitFrame;
            fitFrameCheckbox.enabled = canFitFrame;
            return fitFrameCheckbox;
        }

        /**
         * ボタンエリア（左：画面モード／右：キャンセル・OK）を作る
         * @param {Window} dialogWin 追加先のダイアログ
         * @returns {object} { btnScreenMode, btnCancel, btnOK }
         */
        function addDialogButtons(dialogWin) {
            var buttonRow = addButtonRow(dialogWin);
            var btnScreenMode = addScreenModeButton(buttonRow.leftGroup);
            var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
            var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel('button.ok'), { name: "ok" });
            centerButtonRowIfRightOnly(buttonRow);
            return { btnScreenMode: btnScreenMode, btnCancel: btnCancel, btnOK: btnOK };
        }

        // =========================================
        // ダイアログ / Dialog
        // =========================================

        /**
         * 行の高さを設定するダイアログを表示する
         * @param {Table} baseTable 対象の表
         * @param {Array<number>|null} selectedRowIndices 選択行のインデックス
         * @param {number} initialHeightPt 初期値（pt）
         * @returns {object|null} { height, targetMode, scope, fitFrame }。キャンセル時は null
         */
        function showRowHeightDialog(baseTable, selectedRowIndices, initialHeightPt) {
            /* 触れた表の元の高さを遅延記憶（範囲切り替えに追従） / Lazily remember original heights of touched tables (follows scope changes) */
            var snapshotStore = createRowHeightSnapshotStore();
            var hasSelectedRows = !!(selectedRowIndices && selectedRowIndices.length > 0);
            var parentFrame = getParentTextFrame(baseTable);
            var unitInfo = getUnitInfo(app.activeDocument.viewPreferences.verticalMeasurementUnits);

            var rowHeightDialog = new Window("dialog", getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
            setupWindow(rowHeightDialog, 10);

            var columnsGroup = rowHeightDialog.add("group");
            setupRow(columnsGroup, "fill", COLUMN_SPACING);
            columnsGroup.alignChildren = ["fill", "top"];

            var leftColumn = columnsGroup.add("group");
            leftColumn.orientation = "column";
            leftColumn.alignChildren = "fill";
            leftColumn.alignment = ["fill", "fill"];

            var rightColumn = columnsGroup.add("group");
            rightColumn.orientation = "column";
            rightColumn.alignChildren = "fill";
            rightColumn.alignment = ["fill", "fill"];

            var scopeControls = addScopePanel(leftColumn);
            var targetControls = addTargetPanel(leftColumn, hasSelectedRows);
            /* ∧∨・↑↓キーでも手入力（onChanging）と同じくプレビューを描き直す / refresh the preview as typing does */
            var heightStepOptions = { min: 0, onStep: function () { updatePreview(); } };
            var heightControls = addRowHeightPanel(rightColumn, formatDisplayValue(initialHeightPt / unitInfo.pointsPerUnit), unitInfo.label, heightStepOptions);
            var fitFrameCheckbox = addOptionsPanel(rightColumn, parentFrame !== null);
            var buttons = addDialogButtons(rowHeightDialog);

            /**
             * 選択中の範囲を取得する
             * @returns {string} "document" / "story" / "selection"
             */
            function getCurrentScope() {
                if (scopeControls.documentRadio.value) return "document";
                if (scopeControls.storyRadio.value) return "story";
                return "selection";
            }

            /**
             * 選択中の対象モードを取得する
             * @returns {string} "whole" / "body" / "selected"
             */
            function getCurrentTargetMode() {
                if (targetControls.selectedRadio.value && targetControls.selectedRadio.enabled) return "selected";
                if (targetControls.bodyRadio.value) return "body";
                return "whole";
            }

            /**
             * 現在のモードに応じた行の高さ（pt）を求める
             * @returns {number|null} 行の高さ。入力が正の数でなければ null
             */
            function getRowHeightValue() {
                if (heightControls.minimumRadio.value) return MIN_ROW_HEIGHT_PT;
                var inputValue = parseFloat(heightControls.heightInput.text);
                if (isNaN(inputValue) || inputValue <= 0) return null;
                return inputValue * unitInfo.pointsPerUnit;
            }

            /**
             * 範囲に応じて「選択した行のみ」の有効／無効を切り替える
             * @returns {void}
             */
            function syncTargetEnabled() {
                var allowSelected = (getCurrentScope() === "selection") && hasSelectedRows;
                targetControls.selectedRadio.enabled = allowSelected;
                if (!allowSelected && targetControls.selectedRadio.value) {
                    targetControls.wholeRadio.value = true;
                }
            }

            /**
             * 指定値モードのときだけ入力欄を有効にする
             * @returns {void}
             */
            function syncInputEnabled() {
                heightControls.heightInput.enabled = heightControls.specifiedRadio.value;
                heightControls.heightStepper.enabled = heightControls.specifiedRadio.value;
                redrawSteppersIn(heightControls.heightStepper); /* ∧∨のディム表示を切り替える / update the stepper's dimming */
            }

            /**
             * 現在の設定で行の高さのプレビューを描き直す
             * @returns {void}
             */
            function updatePreview() {
                /* 毎回すべて元に戻してから対象表に適用 / Always restore then apply to the current target tables */
                snapshotStore.restoreAll();
                var rowHeight = getRowHeightValue();
                if (rowHeight === null) return;
                var tables = resolveScopeTables(getCurrentScope(), baseTable);
                for (var i = 0; i < tables.length; i++) snapshotStore.rememberTable(tables[i]);
                var rowSets = collectTargetRowSets(tables, baseTable, selectedRowIndices, getCurrentTargetMode());
                for (var j = 0; j < rowSets.length; j++) {
                    for (var k = 0; k < rowSets[j].length; k++) rowSets[j][k].height = rowHeight;
                }
                if (fitFrameCheckbox.value) {
                    snapshotStore.rememberFrame(parentFrame);
                    fitFrameToContent(parentFrame);
                }
            }

            scopeControls.documentRadio.onClick =
                scopeControls.storyRadio.onClick =
                scopeControls.selectionRadio.onClick = function () { syncTargetEnabled(); updatePreview(); };

            targetControls.wholeRadio.onClick =
                targetControls.bodyRadio.onClick =
                targetControls.selectedRadio.onClick = updatePreview;

            heightControls.minimumRadio.onClick = function () { syncInputEnabled(); updatePreview(); };
            heightControls.specifiedRadio.onClick = function () {
                syncInputEnabled();
                if (heightControls.specifiedRadio.value) heightControls.heightInput.active = true;
                updatePreview();
            };
            heightControls.heightInput.onChanging = updatePreview;

            fitFrameCheckbox.onClick = updatePreview;

            buttons.btnCancel.onClick = function () {
                snapshotStore.restoreAll();
                rowHeightDialog.close(0);
            };

            /* 初期状態を範囲に同期してプレビュー / Sync to scope and show the initial preview */
            syncInputEnabled();
            syncTargetEnabled();
            updatePreview();

            if (rowHeightDialog.show() !== 1) return null;

            var dialogResult = {
                height: getRowHeightValue(),
                targetMode: getCurrentTargetMode(),
                scope: getCurrentScope(),
                fitFrame: fitFrameCheckbox.value ? parentFrame : null
            };
            snapshotStore.restoreAll();

            if (dialogResult.height === null) {
                alert(getLabel('alert.invalidNumber'));
                return null;
            }
            return dialogResult;
        }

        // =========================================
        // 適用 / Apply
        // =========================================

        /**
         * 表ごとの対象行に行の高さを適用する（進捗表示つき、1 回の取り消しにまとめる）
         * @param {Array<Array<Row>>} rowSets 表ごとの行の配列
         * @param {number} rowHeight 行の高さ（pt）
         * @param {TextFrame|null} frameToFit 適用後に内容に合わせるフレーム。合わせない場合は null
         * @returns {void}
         */
        function applyRowHeightToRowSets(rowSets, rowHeight, frameToFit) {
            var totalRows = 0;
            for (var i = 0; i < rowSets.length; i++) totalRows += rowSets[i].length;

            var progress = createProgressBar(getLabel('progress.title') + ' ' + SCRIPT_VERSION, totalRows);

            /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
            app.doScript(function () {
                try {
                    var doneRows = 0;
                    for (var j = 0; j < rowSets.length; j++) {
                        for (var k = 0; k < rowSets[j].length; k++) {
                            rowSets[j][k].height = rowHeight;
                            doneRows++;
                            /* 行ごとの更新は重いので間引く / Throttle updates since per-row refresh is costly */
                            if (doneRows === totalRows || doneRows % 10 === 0) {
                                progress.update(doneRows, doneRows + " / " + totalRows);
                            }
                        }
                    }
                    if (frameToFit) fitFrameToContent(frameToFit);
                } finally {
                    progress.close();
                }
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel('undo.applyRowHeight'));
        }

        // =========================================
        // メイン処理 / Main
        // =========================================

        /**
         * 選択から表を特定し、ダイアログの設定で行の高さを適用する
         * @returns {void}
         */
        function main() {
            var originalSelection = collectionToArray(app.selection);
            var selectionTarget = resolveTargetFromSelection(app.selection);

            if (!selectionTarget) {
                alert(getLabel('alert.noTable'));
                return;
            }
            if (selectionTarget.error === "multipleTables") {
                alert(getLabel('alert.multipleTables'));
                return;
            }

            var baseTable = selectionTarget.table;
            var initialHeightPt = getInitialRowHeight(baseTable, selectionTarget.rowIndices);

            /* 表全体のときのみハイライトをオフ / Clear selection highlight only when whole table is targeted */
            var didClearSelection = (selectionTarget.rowIndices === null);
            if (didClearSelection) app.select(NothingEnum.NOTHING);

            var dialogResult = showRowHeightDialog(baseTable, selectionTarget.rowIndices, initialHeightPt);

            if (didClearSelection) restoreSelection(originalSelection);
            if (dialogResult === null) return;

            var targetTables = resolveScopeTables(dialogResult.scope, baseTable);
            var rowSets = collectTargetRowSets(targetTables, baseTable, selectionTarget.rowIndices, dialogResult.targetMode);
            applyRowHeightToRowSets(rowSets, dialogResult.height, dialogResult.fitFrame);
        }

        main();

    })();
