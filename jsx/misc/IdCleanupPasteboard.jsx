#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdCleanupPasteboard
Name=ページ外（ペーストボード上）のオブジェクトを削除
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.0.1
Release-Date=2026-10-29
Description-BEGIN
現在の見開き・ドキュメント・開いているすべてのドキュメントで、ページの外側（ペーストボード）に置かれたオブジェクトをまとめて削除します。
残す範囲は、裁ち落とし・印刷可能領域・ページからの距離から選べます。

Deletes objects placed outside the pages (on the pasteboard) in the current spread, the active document, or all open documents.
Choose the area to keep: the bleed, the slug, or a distance from the pages.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

現在の見開き・ドキュメント・開いているすべてのドキュメントで、ページの外側（ペーストボード）に置かれたオブジェクトをまとめて削除します。
残す範囲は、裁ち落とし・印刷可能領域・ページからの距離から選べます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCleanupPasteboard.md

### Overview

Deletes objects placed outside the pages (on the pasteboard) in the current spread, the active document, or all open documents.
Choose the area to keep: the bleed, the slug, or a distance from the pages.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCleanupPasteboard.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdCleanupPasteboard";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-29";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-05";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "ページ外（ペーストボード上）のオブジェクトを削除"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCleanupPasteboard.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCleanupPasteboard.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/misc/IdCleanupPasteboard.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @discussion 参考 / Reference
 * 同じ目的のスクリプト：CleanupPasteboard（Marc Autret）。コードは使わず、仕様から書き起こしている
 * A script with the same purpose: CleanupPasteboard by Marc Autret. Written from a spec, without its code
 * http://indiscripts.com
 */

(function () {

    // =========================================
    // 基本設定 / Settings
    // =========================================

    var DEFAULT_SCOPE                 = "document"; /* 対象の初期値（"spread" 現在の見開きページ / "document" ドキュメント全体 / "allDocuments" 開いているすべてのドキュメント） / default scope */
    var DEFAULT_RANGE_MODE            = "distance"; /* 範囲の初期値（"bleed" 裁ち落とし / "slug" 印刷可能領域 / "distance" ページからの距離） / default area */
    var DEFAULT_DISTANCE              = 0;          /* 距離の初期値（DISTANCE_UNIT の単位） / default distance, in DISTANCE_UNIT */
    var DISTANCE_UNIT                 = "mm";       /* 距離の単位（mm / cm / in / pt / pc / px / Q / H） / distance unit */
    var DISTANCE_STEP                 = 1;          /* ∧∨・↑↓キーの増減量 / step for the stepper and arrow keys */
    var DISTANCE_MIN                  = 0;          /* 距離の下限 / minimum distance */
    var DEFAULT_INCLUDE_PARENT_PAGES  = false;      /* 「親ページも対象にする」の初期値 / default for "Include parent pages" */
    var DEFAULT_KEEP_THREADED_FRAMES  = true;       /* 「連結されたテキストフレームは残す」の初期値 / default for "Keep threaded text frames" */
    var DEFAULT_KEEP_HIDDEN_ITEMS     = true;       /* 「非表示のオブジェクトは残す」の初期値 / default for "Keep hidden objects" */
    var DEFAULT_DELETE_LOCKED_ITEMS   = false;      /* 「ロックされたオブジェクトを削除」の初期値 / default for "Delete locked objects" */
    var DEFAULT_CONFIRM_BEFORE_DELETE = true;       /* 「削除する前に確認」の初期値 / default for "Confirm before deleting" */
    /* 2回目以降は、前回 OK を押したときの設定が初期値になる / from the second run, the settings of the last OK are used instead */

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

    var DISTANCE_FIELD_INDENT     = 20; /* 距離の欄の字下げ（px） / indent of the distance field */
    var DISTANCE_FIELD_CHARACTERS = 7;  /* 距離の欄の幅（文字数） / width of the distance field in characters */
    var LAYER_LIST_SIZE           = [200, 110]; /* 残すレイヤーのリストの大きさ [幅, 高さ] / size of the layer list */

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
        "q": { unit: "mm", amount: 0.25 },    /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 },    /* 歯 / H */
        "p": { unit: "pc", amount: 1 },       /* パイカ / pica */
        "ft/in": { unit: "ft", amount: 1 },   /* Illustrator の単位コード7の表示 / Illustrator unit code 7 */
        "c": { unit: "ci", amount: 1 },       /* シセロ（InDesign の表示） / ciceros as InDesign shows them */
        "ag": { unit: "in", amount: 1 / 14 }, /* アゲート / agates */
        "ap": { unit: "tpt", amount: 1 }      /* アメリカンポイント / American points */
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
     * 数値欄の単位を差し替える（単位の設定やドロップダウンを切り替えたとき用）。
     * shouldConvert が true なら値を新しい単位へ換算し（10 mm → 28.35 pt）、false なら数値はそのままで単位だけ付け替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄、または bindSteppedArrowKeys() を呼んだ入力欄
     * @param {string} unit - 新しい単位（例 " pt"。単位なしは ""）
     * @param {boolean} [shouldConvert] - 値も換算するなら true
     * @returns {void}
     */
    function setSteppedFieldUnit(numberInput, unit, shouldConvert) {
        var stepOptions = numberInput.stepperGroup.stepOptions;
        var oldUnit = stepOptions.unit || "";
        var value = parseFloat(numberInput.text);
        stepOptions.unit = unit;
        if (isNaN(value)) return;
        if (shouldConvert) {
            var converted = evaluateArithmetic(String(value) + oldUnit, unit);
            if (!isNaN(converted)) value = converted;
        }
        numberInput.text = formatStepperNumber(value) + unit;
        numberInput.lastValidText = numberInput.text;
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
            if (!hasOperator && value === parseFloat(numberInput.text)) {
                /* 単位を省いて入れた数値には、欄の単位だけ付け足す（桁は丸めない） / append the field unit to a bare number */
                var trimmedText = numberInput.text.replace(/^\s+|\s+$/g, "");
                if (fieldUnit && /[\d.]$/.test(trimmedText)) numberInput.text = trimmedText + fieldUnit;
                return;
            }
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
        });
        numberInput.stepperGroup = stepperGroup; /* setSteppedFieldUnit() から∧∨の設定を引けるようにする / lets setSteppedFieldUnit() find the options */
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
            var unitMatch = /^(ft\/in|[A-Za-z]+|%|°)/i.exec(source.substring(position));
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

    /**
     * ウィンドウの下端にバージョン（SCRIPT_VERSION）を薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * 文字サイズは変えられない（InDesign の ScriptUI は無視する）ので、無効にしてディム表示の薄さにする
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {StaticText|null} バージョンの文字。すでに足してあれば null
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
        var versionText = footerGroup.add("statictext", undefined, SCRIPT_VERSION);
        versionText.enabled = false;
        targetWindow.versionFooter = versionText;
        return versionText;
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

    var LABELS = {
        dialog: {
            title:   { ja: "ページ外のオブジェクトを削除", en: "Delete Objects Outside Pages" },
            confirm: { ja: "削除の確認", en: "Confirm Deletion" }
        },
        panel: {
            scope:       { ja: "対象", en: "Scope" },
            range:       { ja: "この範囲の外側を削除", en: "Delete Outside This Area" },
            options:     { ja: "オプション", en: "Options" },
            keptLayers:  { ja: "残すレイヤー", en: "Layers to Keep" }
        },
        radio: {
            spread:       { ja: "現在の見開きページ", en: "Current spread" },
            document:     { ja: "ドキュメント全体", en: "Entire document" },
            allDocuments: { ja: "開いているすべてのドキュメント", en: "All open documents" },
            bleed:        { ja: "裁ち落とし", en: "Bleed" },
            slug:         { ja: "印刷可能領域", en: "Slug" },
            distance:     { ja: "ページからの距離", en: "Distance from pages" }
        },
        fieldLabel: {
            distance: { ja: "距離", en: "Distance" }
        },
        checkbox: {
            includeParentPages:  { ja: "親ページも対象にする", en: "Include parent pages" },
            keepThreadedFrames:  { ja: "連結されたテキストフレームは残す", en: "Keep threaded text frames" },
            keepHiddenItems:     { ja: "非表示のオブジェクトは残す", en: "Keep hidden objects" },
            deleteLockedItems:   { ja: "ロックされたオブジェクトを削除", en: "Delete locked objects" },
            confirmBeforeDelete: { ja: "削除する前に確認", en: "Confirm before deleting" }
        },
        tooltip: {
            spread: {
                ja: "ウインドウに表示中の見開きだけを処理します（親ページを編集中ならその親ページ）。",
                en: "Processes only the spread shown in the window (the parent spread, if you are editing one)."
            },
            document: {
                ja: "現在のドキュメントのすべての見開きを処理します。",
                en: "Processes every spread in the current document."
            },
            allDocuments: {
                ja: "開いているすべてのドキュメントの、すべての見開きを処理します。",
                en: "Processes every spread in every open document."
            },
            bleed: {
                ja: "［ドキュメント設定］の裁ち落とし（天・地・ノド・小口）より外にあるオブジェクトを削除します。",
                en: "Deletes objects beyond the bleed (top, bottom, inside, outside) set in Document Setup."
            },
            slug: {
                ja: "［ドキュメント設定］の印刷可能領域（天・地・ノド・小口）より外にあるオブジェクトを削除します。",
                en: "Deletes objects beyond the slug (top, bottom, inside, outside) set in Document Setup."
            },
            distance: {
                ja: "ページの4辺から、下の距離より外にあるオブジェクトを削除します。",
                en: "Deletes objects farther than the distance below from all four edges of the pages."
            },
            distanceField: {
                ja: "4辺に同じ距離を使います。ほかの単位で入れると %1 に換算します。",
                en: "The same distance is used on all four sides. Values in other units are converted to %1."
            },
            includeParentPages: {
                ja: "ドキュメントページのスプレッドに加えて、親ページのスプレッドも処理します（［現在の見開きページ］では使えません）。",
                en: "Processes parent page spreads as well as document spreads (not available with Current spread)."
            },
            keepThreadedFrames: {
                ja: "ほかのフレームと連結されたテキストフレームと、それを含むグループを削除しません。",
                en: "Keeps text frames threaded to other frames, and groups that contain them."
            },
            keepHiddenItems: {
                ja: "非表示のオブジェクトと、非表示のレイヤー上のオブジェクトを削除しません。",
                en: "Keeps hidden objects and objects on hidden layers."
            },
            deleteLockedItems: {
                ja: "ロックされたオブジェクトもロックを解除して削除します。ロックされたレイヤー上のオブジェクトは削除しません。",
                en: "Unlocks and deletes locked objects too. Objects on locked layers are never deleted."
            },
            confirmBeforeDelete: {
                ja: "削除する前に件数を表示し、削除するか、候補を選択して表示するかを選べます。",
                en: "Shows the number of objects first, so you can delete them or select and show them instead."
            },
            keptLayers: {
                ja: "選んだレイヤーのオブジェクトは削除しません（⌘／Ctrl＋クリックで複数選択）。ほかのドキュメントでは同じ名前のレイヤーを残します。",
                en: "Objects on the selected layers are kept (Cmd/Ctrl-click to select several). In other documents, layers with the same names are kept."
            },
            selectCandidates: {
                ja: "削除せずに、候補がある最初の見開きを表示して、その見開きの候補を選択します。",
                en: "Instead of deleting, shows the first spread with objects to delete and selects them."
            },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger:   { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        button: {
            cancel:           { ja: "キャンセル", en: "Cancel" },
            ok:               { ja: "OK", en: "OK" },
            deleteItems:      { ja: "削除", en: "Delete" },
            selectCandidates: { ja: "選択して表示", en: "Select and Show" }
        },
        message: {
            candidateCount: { ja: "削除するオブジェクト", en: "Objects to delete" },
            documentCount:  { ja: "ドキュメント", en: "Documents" }
        },
        alert: {
            noDocument:      { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            invalidDistance: { ja: "距離には 0 以上の数値を入力してください。", en: "Enter a number of 0 or greater for the distance." },
            noCandidates:    { ja: "削除するオブジェクトはありません。", en: "There are no objects to delete." },
            deletedCount:    { ja: "削除したオブジェクト", en: "Deleted objects" }
        },
        undoName: { ja: "ページ外のオブジェクトを削除", en: "Delete Objects Outside Pages" }
    };

    // =========================================
    // 型の定義 / Type definitions
    // =========================================

    /**
     * ダイアログで選んだ内容（次回の初期値としても記憶する）
     * @typedef {Object} CleanupSettings
     * @property {string} scope - "spread"（現在の見開きページ）/ "document"（ドキュメント全体）/ "allDocuments"（開いているすべてのドキュメント）
     * @property {string} rangeMode - "bleed"（裁ち落とし）/ "slug"（印刷可能領域）/ "distance"（ページからの距離）
     * @property {number} distanceValue - ページからの距離（DISTANCE_UNIT）
     * @property {number} distancePoints - ページからの距離（pt）
     * @property {boolean} includeParentPages - 親ページのスプレッドも処理するなら true（scope が "spread" のときは使わない）
     * @property {boolean} keepThreadedFrames - 連結されたテキストフレームとそれを含むグループを残すなら true
     * @property {boolean} keepHiddenItems - 非表示のオブジェクトと非表示のレイヤー上のオブジェクトを残すなら true
     * @property {boolean} deleteLockedItems - ロックされたオブジェクトも削除するなら true
     * @property {boolean} confirmBeforeDelete - 削除する前に確認するなら true
     * @property {string[]} keptLayerNames - 残すレイヤーの名前
     */

    /**
     * 設定ダイアログに渡す、ドキュメントの状態
     * @typedef {Object} DialogContext
     * @property {boolean} canUseActiveSpread - 表示中の見開きを取れたなら true
     * @property {string[]} layerNames - 現在のドキュメントのレイヤー名（上から順）
     * @property {CleanupSettings} savedSettings - 記憶していた設定（無ければ初期値）
     */

    /**
     * 削除の候補
     * @typedef {Object} DeletionCandidate
     * @property {PageItem} pageItem - 候補のオブジェクト
     * @property {Spread|MasterSpread} spread - 載っているスプレッド
     * @property {number} spreadNumber - 処理したスプレッドの通し番号（同じスプレッドかの判定用）
     * @property {number} documentNumber - 処理したドキュメントの通し番号
     */

    /**
     * ページの外へ広げる量（pt）
     * @typedef {Object} RangeMargins
     * @property {number} top - 天
     * @property {number} bottom - 地
     * @property {number} inside - ノド
     * @property {number} outside - 小口
     */

    /**
     * スプレッド座標の長方形（pt。y は下向き）
     * @typedef {Object} SpreadRect
     * @property {number} top - 天
     * @property {number} left - 左端
     * @property {number} bottom - 地
     * @property {number} right - 右端
     */

    /**
     * スプレッド内の全ページを合わせた長方形と、左端・右端にあるページ
     * @typedef {Object} PagesExtent
     * @property {SpreadRect} rect - 全ページの境界を合わせた長方形
     * @property {Page} leftmostPage - 左端にあるページ
     * @property {Page} rightmostPage - 右端にあるページ
     */

    // =========================================
    // 設定の記憶 / Saved settings
    // =========================================

    /* InDesign には任意の値を残す環境設定 API が無いので、設定ファイルに key=value で書き出す
       / InDesign has no scriptable preference store, so settings go to a key=value file */
    var PREFS_FILE_NAME = "IdCleanupPasteboard-prefs.txt";

    /**
     * 設定ファイルを返す
     * @returns {File} 設定ファイル
     */
    function getPrefsFile() {
        return File(Folder.userData.fsName + "/" + PREFS_FILE_NAME);
    }

    /**
     * テキストファイルを読み込む
     * @param {File} textFile - 読み込むファイル
     * @returns {string} 中身。読めなければ空文字
     */
    function readTextFile(textFile) {
        if (!textFile.exists) return "";
        textFile.encoding = "UTF-8";
        if (!textFile.open("r")) return "";
        var text = textFile.read();
        textFile.close();
        return text;
    }

    /**
     * テキストファイルへ書き出す。書き出せなくても処理は続ける（設定が記憶されないだけ）
     * @param {File} textFile - 書き出すファイル
     * @param {string} text - 書き出す内容
     * @returns {void}
     */
    function writeTextFile(textFile, text) {
        textFile.encoding = "UTF-8";
        textFile.lineFeed = "Unix";
        if (!textFile.open("w")) return;
        textFile.write(text);
        textFile.close();
    }

    /**
     * key=value の行を読み、キーと文字列の対応にする
     * @param {string} prefsText - 設定ファイルの中身
     * @returns {Object<string, string>} キーと値の対応
     */
    function parsePrefsText(prefsText) {
        var prefs = {};
        var prefsLines = String(prefsText).split(/\r\n|\r|\n/);
        for (var i = 0; i < prefsLines.length; i++) {
            /* 値にも = が入りうるので最初の = だけで区切る / split on the first = only */
            var separatorIndex = prefsLines[i].indexOf("=");
            if (separatorIndex > 0) prefs[prefsLines[i].substring(0, separatorIndex)] = prefsLines[i].substring(separatorIndex + 1);
        }
        return prefs;
    }

    /**
     * 記憶した真偽値を読む
     * @param {string|undefined} storedText - 記憶した値（"true" / "false"）
     * @param {boolean} defaultValue - 記憶が無いときの値
     * @returns {boolean} 真偽値
     */
    function readStoredBoolean(storedText, defaultValue) {
        if (storedText === "true") return true;
        if (storedText === "false") return false;
        return defaultValue;
    }

    /**
     * 記憶した値が候補のどれかなら返し、そうでなければ既定値を返す
     * @param {string|undefined} storedText - 記憶した値
     * @param {string[]} choices - 候補
     * @param {string} defaultValue - 候補に無いときの値
     * @returns {string} 値
     */
    function readStoredChoice(storedText, choices, defaultValue) {
        for (var i = 0; i < choices.length; i++) {
            if (choices[i] === storedText) return storedText;
        }
        return defaultValue;
    }

    /**
     * 前回の設定を読み出す。記録が無い項目は初期値にする
     * @returns {CleanupSettings} 設定
     */
    function loadSavedSettings() {
        var prefs = parsePrefsText(readTextFile(getPrefsFile()));
        var storedDistance = parseFloat(prefs.distanceValue);
        var distanceValue = (storedDistance >= DISTANCE_MIN) ? storedDistance : DEFAULT_DISTANCE;
        var keptLayerNames = [];
        if (prefs.keptLayerNames) {
            var encodedNames = prefs.keptLayerNames.split(",");
            for (var i = 0; i < encodedNames.length; i++) keptLayerNames.push(decodeURIComponent(encodedNames[i]));
        }
        return {
            scope: readStoredChoice(prefs.scope, SCOPES, DEFAULT_SCOPE),
            rangeMode: readStoredChoice(prefs.rangeMode, RANGE_MODES, DEFAULT_RANGE_MODE),
            distanceValue: distanceValue,
            distancePoints: convertDistanceToPoints(distanceValue),
            includeParentPages: readStoredBoolean(prefs.includeParentPages, DEFAULT_INCLUDE_PARENT_PAGES),
            keepThreadedFrames: readStoredBoolean(prefs.keepThreadedFrames, DEFAULT_KEEP_THREADED_FRAMES),
            keepHiddenItems: readStoredBoolean(prefs.keepHiddenItems, DEFAULT_KEEP_HIDDEN_ITEMS),
            deleteLockedItems: readStoredBoolean(prefs.deleteLockedItems, DEFAULT_DELETE_LOCKED_ITEMS),
            confirmBeforeDelete: readStoredBoolean(prefs.confirmBeforeDelete, DEFAULT_CONFIRM_BEFORE_DELETE),
            keptLayerNames: keptLayerNames
        };
    }

    /**
     * 設定を次回の初期値として書き出す
     * @param {CleanupSettings} cleanupSettings - 設定
     * @returns {void}
     */
    function saveSettings(cleanupSettings) {
        /* レイヤー名は , や改行を含みうるので URI エンコードして並べる / layer names may contain commas or line breaks */
        var encodedNames = [];
        for (var i = 0; i < cleanupSettings.keptLayerNames.length; i++) encodedNames.push(encodeURIComponent(cleanupSettings.keptLayerNames[i]));
        var prefsLines = [
            "scope=" + cleanupSettings.scope,
            "rangeMode=" + cleanupSettings.rangeMode,
            "distanceValue=" + cleanupSettings.distanceValue,
            "includeParentPages=" + cleanupSettings.includeParentPages,
            "keepThreadedFrames=" + cleanupSettings.keepThreadedFrames,
            "keepHiddenItems=" + cleanupSettings.keepHiddenItems,
            "deleteLockedItems=" + cleanupSettings.deleteLockedItems,
            "confirmBeforeDelete=" + cleanupSettings.confirmBeforeDelete,
            "keptLayerNames=" + encodedNames.join(",")
        ];
        writeTextFile(getPrefsFile(), prefsLines.join("\n"));
    }

    // =========================================
    // 設定ダイアログ / Settings dialog
    // =========================================

    /* 対象の選び方（ラジオボタンの並び順） / Scopes, in radio button order */
    var SCOPES = ["spread", "document", "allDocuments"];

    /* 範囲の選び方（ラジオボタンの並び順） / Area modes, in radio button order */
    var RANGE_MODES = ["bleed", "slug", "distance"];

    /**
     * ラジオボタンを縦に並べ、指定した値のものを選んでおく
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string[]} choices - 値の並び（LABELS.radio と LABELS.tooltip のキー）
     * @param {string} selectedChoice - 選んでおく値
     * @returns {Object<string, RadioButton>} 値ごとのラジオボタン
     */
    function addRadioButtons(parentPanel, choices, selectedChoice) {
        var radioButtons = {};
        for (var i = 0; i < choices.length; i++) {
            var radioButton = parentPanel.add("radiobutton", undefined, getLabel(LABELS.radio[choices[i]]));
            radioButton.value = (choices[i] === selectedChoice);
            radioButton.helpTip = getLabel(LABELS.tooltip[choices[i]]);
            radioButtons[choices[i]] = radioButton;
        }
        return radioButtons;
    }

    /**
     * 選ばれているラジオボタンの値を返す
     * @param {Object<string, RadioButton>} radioButtons - 値ごとのラジオボタン
     * @param {string[]} choices - 値の並び
     * @returns {string} 選ばれている値（どれも選ばれていなければ先頭）
     */
    function getSelectedChoice(radioButtons, choices) {
        for (var i = 0; i < choices.length; i++) {
            if (radioButtons[choices[i]].value) return choices[i];
        }
        return choices[0];
    }

    /**
     * 「対象」パネルを作る。表示中の見開きを取れないとき（ストーリーエディターが前面など）は、
     * ［現在の見開きページ］をディム表示にする
     * @param {Group} parent - 追加先
     * @param {string} initialScope - 選んでおく対象
     * @param {boolean} canUseActiveSpread - 表示中の見開きを取れたなら true
     * @returns {Object<string, RadioButton>} 対象ごとのラジオボタン
     */
    function addScopePanel(parent, initialScope, canUseActiveSpread) {
        var scopePanel = parent.add("panel", undefined, getLabel(LABELS.panel.scope));
        setupPanel(scopePanel, 6);
        var scopeRadios = addRadioButtons(scopePanel, SCOPES, initialScope);
        scopeRadios.spread.enabled = canUseActiveSpread;
        return scopeRadios;
    }

    /**
     * 「この範囲の外側を削除」パネル（ラジオボタン3つと、字下げした距離の欄）を作る
     * @param {Group} parent - 追加先
     * @param {CleanupSettings} initialSettings - 初期値
     * @returns {{radios: Object<string, RadioButton>, distanceInput: EditText}} 範囲ごとのラジオボタンと距離の欄
     */
    function addRangePanel(parent, initialSettings) {
        var rangePanel = parent.add("panel", undefined, getLabel(LABELS.panel.range));
        setupPanel(rangePanel, 6);
        var rangeRadios = addRadioButtons(rangePanel, RANGE_MODES, initialSettings.rangeMode);

        /* ラジオボタンの下に字下げして置く / indented below the radio buttons */
        var distanceRowGroup = rangePanel.add("group");
        distanceRowGroup.orientation = "row";
        distanceRowGroup.alignChildren = ["left", "center"];
        distanceRowGroup.margins = [DISTANCE_FIELD_INDENT, 0, 0, 0];

        /* ∧∨と↑↓キーは同じ処理で増減する / the stepper and the arrow keys share one step logic */
        var distanceInput = addSteppedField(distanceRowGroup, {
            label: labelText(LABELS.fieldLabel.distance),
            text: formatSteppedValue(initialSettings.distanceValue, { min: DISTANCE_MIN, unit: " " + DISTANCE_UNIT }),
            characters: DISTANCE_FIELD_CHARACTERS,
            step: DISTANCE_STEP,
            min: DISTANCE_MIN,
            unit: " " + DISTANCE_UNIT
        });
        distanceInput.helpTip = getLabel(LABELS.tooltip.distanceField, [DISTANCE_UNIT]);

        return { radios: rangeRadios, distanceInput: distanceInput };
    }

    /**
     * チェックボックスを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelKey - LABELS.checkbox と LABELS.tooltip のキー
     * @param {boolean} isChecked - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, labelKey, isChecked) {
        var optionCheckbox = parentPanel.add("checkbox", undefined, getLabel(LABELS.checkbox[labelKey]));
        optionCheckbox.value = isChecked;
        optionCheckbox.helpTip = getLabel(LABELS.tooltip[labelKey]);
        return optionCheckbox;
    }

    /* オプションの並び順（LABELS.checkbox と CleanupSettings のキー） / Option order */
    var OPTION_KEYS = ["includeParentPages", "keepThreadedFrames", "keepHiddenItems", "deleteLockedItems", "confirmBeforeDelete"];

    /**
     * 「オプション」パネルを作る
     * @param {Group} parent - 追加先
     * @param {CleanupSettings} initialSettings - 初期値
     * @returns {Object<string, Checkbox>} オプションごとのチェックボックス
     */
    function addOptionPanel(parent, initialSettings) {
        var optionPanel = parent.add("panel", undefined, getLabel(LABELS.panel.options));
        setupPanel(optionPanel, 6);
        var optionCheckboxes = {};
        for (var i = 0; i < OPTION_KEYS.length; i++) {
            optionCheckboxes[OPTION_KEYS[i]] = addOptionCheckbox(optionPanel, OPTION_KEYS[i], initialSettings[OPTION_KEYS[i]]);
        }
        return optionCheckboxes;
    }

    /**
     * 「残すレイヤー」パネル（複数選択できるリスト）を作る。記憶したレイヤー名と同じ名前を選んでおく
     * @param {Group} parent - 追加先
     * @param {string[]} layerNames - 現在のドキュメントのレイヤー名
     * @param {string[]} keptLayerNames - 選んでおくレイヤー名
     * @returns {ListBox} レイヤーのリスト
     */
    function addKeptLayerPanel(parent, layerNames, keptLayerNames) {
        var keptLayerPanel = parent.add("panel", undefined, getLabel(LABELS.panel.keptLayers));
        setupPanel(keptLayerPanel, 6);
        var layerList = keptLayerPanel.add("listbox", undefined, layerNames, { multiselect: true });
        layerList.preferredSize = LAYER_LIST_SIZE;
        layerList.alignment = ["fill", "fill"];
        layerList.helpTip = getLabel(LABELS.tooltip.keptLayers);
        var keptLayerLookup = createLayerLookup(keptLayerNames);
        for (var i = 0; i < layerList.items.length; i++) {
            if (isLayerNameKept(keptLayerLookup, layerList.items[i].text)) layerList.items[i].selected = true;
        }
        return layerList;
    }

    /**
     * リストで選ばれているレイヤー名を返す
     * @param {ListBox} layerList - レイヤーのリスト
     * @returns {string[]} 選ばれているレイヤー名
     */
    function getSelectedLayerNames(layerList) {
        var selectedNames = [];
        for (var i = 0; i < layerList.items.length; i++) {
            if (layerList.items[i].selected) selectedNames.push(layerList.items[i].text);
        }
        return selectedNames;
    }

    /**
     * 距離の欄の値を DISTANCE_UNIT の数値で読む。
     * 確定前の計算式や別の単位（「1 in」など）も、欄の確定時と同じ計算で換算する
     * @param {EditText} distanceInput - 距離の欄（「5 mm」の形）
     * @returns {number} 距離（読めなければ NaN）
     */
    function readDistanceValue(distanceInput) {
        var distanceValue = evaluateArithmetic(distanceInput.text, " " + DISTANCE_UNIT);
        return isNaN(distanceValue) ? parseFloat(distanceInput.text) : distanceValue;
    }

    /**
     * 距離（DISTANCE_UNIT の単位）を pt に換算する
     * @param {number} distanceValue - 距離
     * @returns {number} pt での距離
     */
    function convertDistanceToPoints(distanceValue) {
        return createStepperUnitValue(distanceValue, DISTANCE_UNIT.toLowerCase()).as("pt");
    }

    /**
     * 閉じたダイアログのコントロールから、設定を読み取る
     * @param {{scopeRadios: Object<string, RadioButton>, rangeControls: Object, optionCheckboxes: Object<string, Checkbox>, layerList: ListBox}} dialogControls - ダイアログのコントロール
     * @returns {CleanupSettings} 設定
     */
    function readDialogSettings(dialogControls) {
        var distanceValue = readDistanceValue(dialogControls.rangeControls.distanceInput);
        var cleanupSettings = {
            scope: getSelectedChoice(dialogControls.scopeRadios, SCOPES),
            rangeMode: getSelectedChoice(dialogControls.rangeControls.radios, RANGE_MODES),
            distanceValue: distanceValue,
            distancePoints: convertDistanceToPoints(distanceValue),
            keptLayerNames: getSelectedLayerNames(dialogControls.layerList)
        };
        for (var i = 0; i < OPTION_KEYS.length; i++) {
            cleanupSettings[OPTION_KEYS[i]] = dialogControls.optionCheckboxes[OPTION_KEYS[i]].value;
        }
        return cleanupSettings;
    }

    /**
     * 設定ダイアログを表示して、対象・削除の範囲・オプション・残すレイヤーを選ばせる
     * @param {DialogContext} dialogContext - ドキュメントの状態と記憶した設定
     * @returns {CleanupSettings|null} 選んだ内容。キャンセル時は null
     */
    function showSettingsDialog(dialogContext) {
        var savedSettings = dialogContext.savedSettings;
        /* 記憶した対象が使えなければドキュメント全体にする / fall back when the saved scope is unavailable */
        var initialScope = (savedSettings.scope === "spread" && !dialogContext.canUseActiveSpread) ? "document" : savedSettings.scope;

        var settingsDialog = new Window("dialog", getLabel(LABELS.dialog.title));
        setupWindow(settingsDialog);

        /* 左に対象と範囲、右にオプションと残すレイヤー / scope and area on the left, options and layers on the right */
        var columnsGroup = settingsDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;
        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];
        leftColumn.spacing = WINDOW_SPACING;
        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];
        rightColumn.spacing = WINDOW_SPACING;

        var dialogControls = {
            scopeRadios: addScopePanel(leftColumn, initialScope, dialogContext.canUseActiveSpread),
            rangeControls: addRangePanel(leftColumn, savedSettings),
            optionCheckboxes: addOptionPanel(rightColumn, savedSettings),
            layerList: addKeptLayerPanel(rightColumn, dialogContext.layerNames, savedSettings.keptLayerNames)
        };

        var buttonRow = addButtonRow(settingsDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        bindDialogDimming(dialogControls, initialScope, savedSettings.rangeMode);

        btnOK.onClick = function () {
            var rangeControls = dialogControls.rangeControls;
            if (getSelectedChoice(rangeControls.radios, RANGE_MODES) === "distance" && !(readDistanceValue(rangeControls.distanceInput) >= DISTANCE_MIN)) {
                alert(getLabel(LABELS.alert.invalidDistance), getLabel(LABELS.dialog.title));
                rangeControls.distanceInput.active = true;
                return;
            }
            settingsDialog.close(1);
        };

        addVersionFooter(settingsDialog);
        if (settingsDialog.show() !== 1) return null;
        return readDialogSettings(dialogControls);
    }

    /**
     * 選んだ内容に合わせて、使えない欄をディム表示にする処理をつなぐ。
     * 表示前のラジオボタンの value は読み戻せないので、最初の状態は初期値から決める
     * @param {{scopeRadios: Object<string, RadioButton>, rangeControls: Object, optionCheckboxes: Object<string, Checkbox>}} dialogControls - ダイアログのコントロール
     * @param {string} initialScope - 最初の対象
     * @param {string} initialRangeMode - 最初の範囲
     * @returns {void}
     */
    function bindDialogDimming(dialogControls, initialScope, initialRangeMode) {
        var rangeRadios = dialogControls.rangeControls.radios;
        var distanceInput = dialogControls.rangeControls.distanceInput;
        var includeParentPagesCheckbox = dialogControls.optionCheckboxes.includeParentPages;

        /* 距離の欄は［ページからの距離］のときだけ / the distance field only with Distance from pages */
        for (var i = 0; i < RANGE_MODES.length; i++) {
            rangeRadios[RANGE_MODES[i]].onClick = function () {
                setSteppedFieldEnabled(distanceInput, rangeRadios.distance.value);
            };
        }
        setSteppedFieldEnabled(distanceInput, initialRangeMode === "distance");

        /* ［親ページも対象にする］は［現在の見開きページ］では使えない / not available with Current spread */
        for (var j = 0; j < SCOPES.length; j++) {
            dialogControls.scopeRadios[SCOPES[j]].onClick = function () {
                includeParentPagesCheckbox.enabled = !dialogControls.scopeRadios.spread.value;
            };
        }
        includeParentPagesCheckbox.enabled = (initialScope !== "spread");
    }

    // =========================================
    // 確認ダイアログ / Confirmation dialog
    // =========================================

    /**
     * 削除する候補の件数を表示し、削除・選択して表示・キャンセルを選ばせる
     * @param {number} candidateCount - 候補の数
     * @param {number} documentCount - 候補があるドキュメントの数
     * @param {boolean} canSelectCandidates - ［選択して表示］を使えるなら true（すべてのドキュメントが対象のときは使えない）
     * @returns {string} "delete" / "select" / "cancel"
     */
    function showConfirmDialog(candidateCount, documentCount, canSelectCandidates) {
        var confirmDialog = new Window("dialog", getLabel(LABELS.dialog.confirm));
        setupWindow(confirmDialog);
        confirmDialog.add("statictext", undefined, labelValueText(LABELS.message.candidateCount, candidateCount));
        if (documentCount > 1) confirmDialog.add("statictext", undefined, labelValueText(LABELS.message.documentCount, documentCount));

        var buttonRow = addButtonRow(confirmDialog);
        var btnSelect = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.selectCandidates));
        btnSelect.helpTip = getLabel(LABELS.tooltip.selectCandidates);
        btnSelect.enabled = canSelectCandidates;
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnDelete = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.deleteItems), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        btnSelect.onClick = function () {
            confirmDialog.close(2);
        };

        var dialogResult = confirmDialog.show();
        if (dialogResult === 1) return "delete";
        if (dialogResult === 2) return "select";
        return "cancel";
    }

    // =========================================
    // 削除の範囲 / Deletion area
    // =========================================

    /* 範囲の辺との差がこれ以下なら「接している」とみなす（pt）。座標の計算誤差で接しているだけのものを消さないため
       Gaps up to this (pt) count as touching, so float error does not delete items that only touch the edge */
    var EDGE_TOLERANCE_POINTS = 0.001;

    /**
     * 選んだ範囲の、ページの外へ広げる量を返す（定規の単位を pt にしてから呼ぶ）
     * @param {Document} doc - 対象ドキュメント
     * @param {CleanupSettings} cleanupSettings - 設定
     * @returns {RangeMargins} 天・地・ノド・小口の量（pt）
     */
    function getRangeMargins(doc, cleanupSettings) {
        var documentPrefs = doc.documentPreferences;
        if (cleanupSettings.rangeMode === "bleed") {
            return {
                top: documentPrefs.documentBleedTopOffset,
                bottom: documentPrefs.documentBleedBottomOffset,
                inside: documentPrefs.documentBleedInsideOrLeftOffset,
                outside: documentPrefs.documentBleedOutsideOrRightOffset
            };
        }
        if (cleanupSettings.rangeMode === "slug") {
            /* 印刷可能領域は語順が裁ち落としと逆（RightOrOutside） / the slug property names swap the word order */
            return {
                top: documentPrefs.slugTopOffset,
                bottom: documentPrefs.slugBottomOffset,
                inside: documentPrefs.slugInsideOrLeftOffset,
                outside: documentPrefs.slugRightOrOutsideOffset
            };
        }
        var distancePoints = cleanupSettings.distancePoints;
        return { top: distancePoints, bottom: distancePoints, inside: distancePoints, outside: distancePoints };
    }

    /**
     * スプレッド内の全ページの境界を合わせた長方形と、左端・右端にあるページを求める。
     * 最初と最後のページではなく外側の端で決めるので、右綴じやサイズの異なるページが混ざる見開きでも正しい
     * @param {Page[]} spreadPages - スプレッド内のページ（1つ以上）
     * @returns {PagesExtent} 全ページの長方形と左端・右端のページ
     */
    function getPagesExtent(spreadPages) {
        var extent = null;
        for (var i = 0; i < spreadPages.length; i++) {
            var pageBounds = spreadPages[i].bounds; /* [天, 左, 地, 右] / [top, left, bottom, right] */
            if (!extent) {
                extent = {
                    rect: { top: pageBounds[0], left: pageBounds[1], bottom: pageBounds[2], right: pageBounds[3] },
                    leftmostPage: spreadPages[i],
                    rightmostPage: spreadPages[i]
                };
                continue;
            }
            if (pageBounds[0] < extent.rect.top) extent.rect.top = pageBounds[0];
            if (pageBounds[2] > extent.rect.bottom) extent.rect.bottom = pageBounds[2];
            if (pageBounds[1] < extent.rect.left) {
                extent.rect.left = pageBounds[1];
                extent.leftmostPage = spreadPages[i];
            }
            if (pageBounds[3] > extent.rect.right) {
                extent.rect.right = pageBounds[3];
                extent.rightmostPage = spreadPages[i];
            }
        }
        return extent;
    }

    /**
     * スプレッドの削除の範囲（この外側にあるものを削除する）を求める。
     * 左端は左ページなら小口・それ以外はノド、右端は左ページならノド・それ以外は小口だけ広げる
     * （単ページのドキュメントは「ノド＝左」「小口＝右」になる）
     * @param {Page[]} spreadPages - スプレッド内のページ（1つ以上）
     * @param {RangeMargins} rangeMargins - 広げる量（pt）
     * @returns {SpreadRect} 削除の範囲
     */
    function getDeletionRange(spreadPages, rangeMargins) {
        var extent = getPagesExtent(spreadPages);
        var isLeftmostLeftHand = (extent.leftmostPage.side === PageSideOptions.LEFT_HAND);
        var isRightmostLeftHand = (extent.rightmostPage.side === PageSideOptions.LEFT_HAND);
        return {
            top: extent.rect.top - rangeMargins.top,
            bottom: extent.rect.bottom + rangeMargins.bottom,
            left: extent.rect.left - (isLeftmostLeftHand ? rangeMargins.outside : rangeMargins.inside),
            right: extent.rect.right + (isRightmostLeftHand ? rangeMargins.inside : rangeMargins.outside)
        };
    }

    /**
     * 境界が範囲とまったく重ならないかを返す（辺が接しているだけなら重なりとみなす）
     * @param {number[]} itemBounds - オブジェクトの境界 [天, 左, 地, 右]
     * @param {SpreadRect} deletionRange - 削除の範囲
     * @returns {boolean} 範囲の完全に外なら true
     */
    function isOutsideRange(itemBounds, deletionRange) {
        return itemBounds[2] < deletionRange.top - EDGE_TOLERANCE_POINTS ||
            itemBounds[0] > deletionRange.bottom + EDGE_TOLERANCE_POINTS ||
            itemBounds[3] < deletionRange.left - EDGE_TOLERANCE_POINTS ||
            itemBounds[1] > deletionRange.right + EDGE_TOLERANCE_POINTS;
    }

    // =========================================
    // 残すオブジェクトの判定 / Items to keep
    // =========================================

    /**
     * レイヤー名の引き当て表を作る（"toString" などの名前が Object の組み込みとぶつからないよう接頭辞を付ける）
     * @param {string[]} layerNames - レイヤー名
     * @returns {Object<string, boolean>} 引き当て表
     */
    function createLayerLookup(layerNames) {
        var layerLookup = {};
        for (var i = 0; i < layerNames.length; i++) layerLookup["layer:" + layerNames[i]] = true;
        return layerLookup;
    }

    /**
     * レイヤー名が引き当て表にあるかを返す
     * @param {Object<string, boolean>} layerLookup - createLayerLookup() の戻り値
     * @param {string} layerName - レイヤー名
     * @returns {boolean} あれば true
     */
    function isLayerNameKept(layerLookup, layerName) {
        return layerLookup.hasOwnProperty("layer:" + layerName);
    }

    /**
     * ロックのために残すかを返す。ロックされたレイヤー上のものは常に、
     * ロックされたオブジェクトは deleteLockedItems がオフのときに残す
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {boolean} deleteLockedItems - ロックされたオブジェクトも削除するなら true
     * @returns {boolean} 残すなら true
     */
    function isKeptByLock(pageItem, deleteLockedItems) {
        if (pageItem.itemLayer.locked) return true;
        return pageItem.locked && !deleteLockedItems;
    }

    /**
     * オブジェクトか、それが載っているレイヤーが非表示かを返す
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {boolean} 非表示なら true
     */
    function isHiddenItem(pageItem) {
        return !pageItem.visible || !pageItem.itemLayer.visible;
    }

    /**
     * ほかのフレームと連結されたテキストフレームかを返す
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {boolean} 連結されたテキストフレームなら true
     */
    function isThreadedTextFrame(pageItem) {
        return pageItem.constructor.name === "TextFrame" && pageItem.parentStory.textContainers.length > 1;
    }

    /**
     * オブジェクト自身か、その中（グループの中身など）に連結されたテキストフレームがあるかを返す
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {boolean} 連結されたテキストフレームを含めば true
     */
    function containsThreadedTextFrame(pageItem) {
        if (isThreadedTextFrame(pageItem)) return true;
        var nestedItems = pageItem.allPageItems;
        for (var i = 0; i < nestedItems.length; i++) {
            if (isThreadedTextFrame(nestedItems[i])) return true;
        }
        return false;
    }

    /**
     * 範囲の外にあっても残すオブジェクトかを返す（残すレイヤー・ロック・非表示・連結フレーム）
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {CleanupSettings} cleanupSettings - 設定
     * @param {Object<string, boolean>} keptLayerLookup - 残すレイヤーの引き当て表
     * @returns {boolean} 残すなら true
     */
    function shouldKeepItem(pageItem, cleanupSettings, keptLayerLookup) {
        if (isLayerNameKept(keptLayerLookup, pageItem.itemLayer.name)) return true;
        if (isKeptByLock(pageItem, cleanupSettings.deleteLockedItems)) return true;
        if (cleanupSettings.keepHiddenItems && isHiddenItem(pageItem)) return true;
        return cleanupSettings.keepThreadedFrames && containsThreadedTextFrame(pageItem);
    }

    // =========================================
    // 候補の収集 / Collecting candidates
    // =========================================

    /**
     * ウインドウに表示中の見開き（親ページを編集中なら親ページのスプレッド）を返す
     * @returns {Spread|MasterSpread|null} 表示中の見開き。レイアウトウインドウが前面でなければ null
     */
    function getActiveSpread() {
        var activeWindow = app.activeWindow;
        if (!activeWindow || activeWindow.constructor.name !== "LayoutWindow") return null; /* ストーリーエディターには activeSpread が無い / story windows have no activeSpread */
        return activeWindow.activeSpread;
    }

    /**
     * ドキュメントのレイヤー名を上から順に返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {string[]} レイヤー名
     */
    function getLayerNames(doc) {
        var layerNames = [];
        var docLayers = doc.layers.everyItem().getElements();
        for (var i = 0; i < docLayers.length; i++) layerNames.push(docLayers[i].name);
        return layerNames;
    }

    /**
     * 処理するドキュメントを返す
     * @param {string} scope - 対象
     * @returns {Document[]} 処理するドキュメント
     */
    function getTargetDocuments(scope) {
        if (scope === "allDocuments") return app.documents.everyItem().getElements();
        return [app.activeDocument];
    }

    /**
     * ドキュメントの中で処理するスプレッドを返す
     * @param {Document} doc - 対象ドキュメント
     * @param {CleanupSettings} cleanupSettings - 設定
     * @param {Spread|MasterSpread|null} activeSpread - 表示中の見開き
     * @returns {Array<Spread|MasterSpread>} 処理するスプレッド
     */
    function getTargetSpreads(doc, cleanupSettings, activeSpread) {
        if (cleanupSettings.scope === "spread") return [activeSpread];
        var targetSpreads = doc.spreads.everyItem().getElements();
        if (cleanupSettings.includeParentPages && doc.masterSpreads.length > 0) {
            targetSpreads = targetSpreads.concat(doc.masterSpreads.everyItem().getElements());
        }
        return targetSpreads;
    }

    /**
     * スプレッドの最上位のオブジェクトのうち、削除の範囲の外にあって残す理由の無いものを集める
     * @param {Spread|MasterSpread} spread - 対象のスプレッド
     * @param {SpreadRect} deletionRange - 削除の範囲
     * @param {CleanupSettings} cleanupSettings - 設定
     * @param {Object<string, boolean>} keptLayerLookup - 残すレイヤーの引き当て表
     * @returns {PageItem[]} 削除するオブジェクト
     */
    function collectItemsOutside(spread, deletionRange, cleanupSettings, keptLayerLookup) {
        var itemsOutside = [];
        if (spread.pageItems.length === 0) return itemsOutside;
        var topLevelItems = spread.pageItems.everyItem().getElements();
        for (var i = 0; i < topLevelItems.length; i++) {
            var pageItem = topLevelItems[i];
            if (!isOutsideRange(pageItem.visibleBounds, deletionRange)) continue;
            if (shouldKeepItem(pageItem, cleanupSettings, keptLayerLookup)) continue;
            itemsOutside.push(pageItem);
        }
        return itemsOutside;
    }

    /**
     * 定規の単位を縦横とも pt、原点をスプレッドにして処理を実行し、終わったら（失敗しても）元に戻す
     * @param {Document} doc - 対象ドキュメント
     * @param {Function} task - 実行する処理（引数なし）
     * @returns {*} task の戻り値
     */
    function runInPointsWithSpreadOrigin(doc, task) {
        var viewPrefs = doc.viewPreferences;
        var savedHorizontalUnits = viewPrefs.horizontalMeasurementUnits;
        var savedVerticalUnits = viewPrefs.verticalMeasurementUnits;
        var savedRulerOrigin = viewPrefs.rulerOrigin;
        viewPrefs.horizontalMeasurementUnits = MeasurementUnits.POINTS;
        viewPrefs.verticalMeasurementUnits = MeasurementUnits.POINTS;
        viewPrefs.rulerOrigin = RulerOrigin.SPREAD_ORIGIN;
        try {
            return task();
        } finally {
            viewPrefs.horizontalMeasurementUnits = savedHorizontalUnits;
            viewPrefs.verticalMeasurementUnits = savedVerticalUnits;
            viewPrefs.rulerOrigin = savedRulerOrigin;
        }
    }

    /**
     * すべての対象ドキュメント・スプレッドから、削除の候補を集める。
     * 定規の切り替えと復元は取り消しの履歴に残さない（AUTO_UNDO で直前の操作にまとめる）
     * @param {Document[]} targetDocuments - 処理するドキュメント
     * @param {CleanupSettings} cleanupSettings - 設定
     * @param {Spread|MasterSpread|null} activeSpread - 表示中の見開き
     * @returns {DeletionCandidate[]} 削除の候補
     */
    function collectCandidates(targetDocuments, cleanupSettings, activeSpread) {
        var candidates = [];
        var keptLayerLookup = createLayerLookup(cleanupSettings.keptLayerNames);
        var spreadNumber = 0;
        app.doScript(function () {
            for (var i = 0; i < targetDocuments.length; i++) {
                var doc = targetDocuments[i];
                runInPointsWithSpreadOrigin(doc, function () {
                    var rangeMargins = getRangeMargins(doc, cleanupSettings);
                    var targetSpreads = getTargetSpreads(doc, cleanupSettings, activeSpread);
                    for (var j = 0; j < targetSpreads.length; j++) {
                        spreadNumber++;
                        if (targetSpreads[j].pages.length === 0) continue;
                        var deletionRange = getDeletionRange(targetSpreads[j].pages.everyItem().getElements(), rangeMargins);
                        var itemsOutside = collectItemsOutside(targetSpreads[j], deletionRange, cleanupSettings, keptLayerLookup);
                        for (var k = 0; k < itemsOutside.length; k++) {
                            candidates.push({ pageItem: itemsOutside[k], spread: targetSpreads[j], spreadNumber: spreadNumber, documentNumber: i });
                        }
                    }
                });
            }
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.AUTO_UNDO, getLabel(LABELS.undoName));
        return candidates;
    }

    /**
     * 候補があるドキュメントの数を返す
     * @param {DeletionCandidate[]} candidates - 削除の候補
     * @returns {number} ドキュメントの数
     */
    function countCandidateDocuments(candidates) {
        var documentCount = 0;
        var lastDocumentNumber = -1;
        for (var i = 0; i < candidates.length; i++) {
            /* 候補はドキュメント順に並んでいる / candidates are grouped by document */
            if (candidates[i].documentNumber !== lastDocumentNumber) {
                documentCount++;
                lastDocumentNumber = candidates[i].documentNumber;
            }
        }
        return documentCount;
    }

    // =========================================
    // 削除・選択 / Deleting and selecting
    // =========================================

    /**
     * 候補を削除し、削除できた数を返す（ロックされたものはロックを解除してから削除する）
     * @param {DeletionCandidate[]} candidates - 削除の候補
     * @returns {number} 削除できた数
     */
    function removeCandidates(candidates) {
        var removedCount = 0;
        for (var i = 0; i < candidates.length; i++) {
            var pageItem = candidates[i].pageItem;
            try {
                if (pageItem.locked) pageItem.locked = false; /* ロックされたままでは削除できない / locked items cannot be removed */
                pageItem.remove();
                removedCount++;
            } catch (e) {
                /* 削除を受け付けないものは残して数えない / items that refuse removal are left and not counted */
            }
        }
        return removedCount;
    }

    /**
     * 候補の削除を1回の取り消し単位で実行する
     * @param {DeletionCandidate[]} candidates - 削除の候補
     * @returns {number} 削除したオブジェクトの数
     */
    function removeCandidatesWithUndo(candidates) {
        var removedCount = 0;
        app.doScript(function () {
            removedCount = removeCandidates(candidates);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undoName));
        return removedCount;
    }

    /**
     * 候補がある最初の見開きを表示し、その見開きの候補を選択する（現在のドキュメントの候補だけを渡す）。
     * 選択できないもの（ロック・非表示）は選択に含めない
     * @param {DeletionCandidate[]} candidates - 削除の候補
     * @returns {void}
     */
    function selectCandidatesOnFirstSpread(candidates) {
        var firstSpreadNumber = candidates[0].spreadNumber;
        var selectableItems = [];
        for (var i = 0; i < candidates.length; i++) {
            if (candidates[i].spreadNumber !== firstSpreadNumber) continue;
            var pageItem = candidates[i].pageItem;
            if (!pageItem.locked && !isHiddenItem(pageItem)) selectableItems.push(pageItem);
        }
        var activeWindow = app.activeWindow;
        if (activeWindow && activeWindow.constructor.name === "LayoutWindow") activeWindow.activeSpread = candidates[0].spread;
        app.activeDocument.select(selectableItems.length > 0 ? selectableItems : NothingEnum.NOTHING, SelectionOptions.REPLACE_WITH);
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * 設定ダイアログ → 候補の収集 → 確認 → 削除（または選択） → 結果表示
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument), getLabel(LABELS.dialog.title));
            return;
        }

        /* ダイアログを開く前に控える / captured before the dialog opens */
        var activeSpread = getActiveSpread();

        var cleanupSettings = showSettingsDialog({
            canUseActiveSpread: activeSpread !== null,
            layerNames: getLayerNames(app.activeDocument),
            savedSettings: loadSavedSettings()
        });
        if (!cleanupSettings) return;
        saveSettings(cleanupSettings);

        var candidates = collectCandidates(getTargetDocuments(cleanupSettings.scope), cleanupSettings, activeSpread);
        if (candidates.length === 0) {
            alert(getLabel(LABELS.alert.noCandidates), getLabel(LABELS.dialog.title));
            return;
        }

        if (cleanupSettings.confirmBeforeDelete) {
            var confirmChoice = showConfirmDialog(candidates.length, countCandidateDocuments(candidates), cleanupSettings.scope !== "allDocuments");
            if (confirmChoice === "select") {
                selectCandidatesOnFirstSpread(candidates);
                return;
            }
            if (confirmChoice !== "delete") return;
        }

        var removedCount = removeCandidatesWithUndo(candidates);
        alert(labelValueText(LABELS.alert.deletedCount, removedCount), getLabel(LABELS.dialog.title));
    }

    main();

})();
