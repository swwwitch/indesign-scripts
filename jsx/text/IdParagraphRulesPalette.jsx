#target indesign
#targetengine "IdParagraphRulesPalette"

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdParagraphRulesPalette
Name=IdParagraphRulesPalette
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.4.0
Release-Date=2026-10-19
Description-BEGIN
InDesign の［段落境界線］ダイアログボックスを常駐パレットで再現し、選択したテキスト・テキストフレーム・表のセルの段落に前境界線・後境界線をプレビューしながら設定します。
変更した項目だけを書き込むので、ほかの設定や複数段落の個別の値はそのまま残ります。

Recreates InDesign's Paragraph Rules dialog box as a palette and sets the rule above and rule below on the paragraphs of the selected text, text frames or table cells with a live preview.
Only the items you change are written, so other settings and per-paragraph values stay as they are.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

InDesign の［段落境界線］ダイアログボックスを常駐パレットで再現し、選択したテキスト・テキストフレーム・表のセルの段落に前境界線・後境界線をプレビューしながら設定します。
変更した項目だけを書き込むので、ほかの設定や複数段落の個別の値はそのまま残ります。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdParagraphRulesPalette.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n4f45fb194ac1

### Overview

Recreates InDesign's Paragraph Rules dialog box as a palette and sets the rule above and rule below on the paragraphs of the selected text, text frames or table cells with a live preview.
Only the items you change are written, so other settings and per-paragraph values stay as they are.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdParagraphRulesPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdParagraphRulesPalette";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-20";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdParagraphRulesPalette.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdParagraphRulesPalette.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n4f45fb194ac1"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 濃淡の範囲と∧∨の増減量 / Tint range and stepper step */
    var TINT_MIN  = 0;
    var TINT_MAX  = 100;
    var TINT_STEP = 10;

    /* 濃淡が「自動」（-1）のときに表示する値 / Tint shown when the rule's tint is automatic (-1) */
    var TINT_AUTO_DISPLAY = 100;

    /* プリセット。Above（前境界線）・Below（後境界線）に書いた境界線はオン、書かなかった境界線はオフにする。
       書いた境界線の省いた項目は PRESET_RULE_DEFAULTS の値になる。
       長さ（weight / offset / leftIndent / rightIndent）は数値なら pt、"0.2em" の形なら先頭の段落の文字サイズに対する倍率。
       color・gapColor は "textColor"（テキストカラー）・"none"・"paper"・"black"・"registration" かスウォッチ名、
       type は "solid"（実線）か線種の名前、width は "column"（段）か "text"（テキスト）。見つからないカラー・線種は今の値のまま
       / Presets. A rule listed under Above or Below is turned on, an unlisted one off; omitted items take PRESET_RULE_DEFAULTS.
       Lengths are pt as numbers, or relative to the first paragraph's point size as "0.2em" */
    var RULE_PRESETS = [
        {
            name: { ja: "下線（テキスト幅）", en: "Underline (Text)" },
            Below: { weight: 0.5, width: "text", offset: "0.15em" }
        },
        {
            name: { ja: "見出しの罫（段幅）", en: "Heading Rule (Column)" },
            Below: { weight: 1, offset: "0.4em" }
        },
        {
            name: { ja: "上下の罫（段幅）", en: "Above & Below (Column)" },
            Above: { weight: 0.5, offset: "1.1em" },
            Below: { weight: 0.5, offset: "0.4em" }
        },
        {
            name: { ja: "帯（背景）", en: "Background Band" },
            Above: { weight: "1.5em", color: "black", tint: 15, width: "text", offset: "-0.35em", leftIndent: "-0.25em", rightIndent: "-0.25em" }
        },
        {
            name: { ja: "マーカー", en: "Highlighter" },
            Above: { weight: "0.45em", color: "black", tint: 20, width: "text", offset: "-0.1em" }
        },
        {
            name: { ja: "境界線なし", en: "No Rules" }
        }
    ];

    /* プリセットで省いた項目の値 / Values for the items a preset leaves out */
    var PRESET_RULE_DEFAULTS = {
        weight: 1,
        type: "solid",
        color: "textColor",
        tint: 100,
        overprint: false,
        gapColor: "none",
        gapTint: 100,
        gapOverprint: false,
        width: "column",
        offset: 0,
        leftIndent: 0,
        rightIndent: 0,
        keepInFrame: false
    };

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

    /* 項目名の幅（左の列・右の列）/ Label widths of the left and right columns */
    var LEFT_LABEL_WIDTH  = 120;
    var RIGHT_LABEL_WIDTH = 100;

    /* 左の列の幅。右の列の位置をそろえる / Width of the left column so the right column lines up */
    var LEFT_CELL_WIDTH = 310;

    /* 入力欄・ドロップダウン・［種類］の幅（px）/ Widths of the numeric fields, dropdowns and the Type dropdown (px) */
    var NUMBER_FIELD_WIDTH  = 110;
    var DROPDOWN_WIDTH      = 160;
    var TYPE_DROPDOWN_WIDTH = 120;

    /* 項目名と入力の間隔、行どうしの間隔 / Spacing between a label and its control, and between rows */
    var FIELD_SPACING = 6;
    var ROW_SPACING   = 6;

    /* 色見本の一辺と、ドロップダウンとの間隔（px）/ Swatch chip size and its gap to the dropdown (px) */
    var SWATCH_CHIP_SIZE    = 16;
    var SWATCH_CHIP_SPACING = 4;

    /* ラジオボタンの文字を項目名の高さにそろえるため、下に足す余白（px。行の中で中央に置くので、半分だけ上がる）
       / Bottom margin that lifts the radio buttons to the label's text line (px; centered in the row, so they rise by half) */
    var RADIO_BASELINE_MARGIN = 4;

    /* オーバープリントなど、入力の列にそろえて置くチェックボックスの字下げ
       / Indent of the checkboxes that line up with the control column */
    var CHECKBOX_INDENT = LEFT_LABEL_WIDTH + FIELD_SPACING;

    /* ［プリセットを保存］の名前の欄を、ラジオボタンの文字にそろえる字下げ / Indent of the name field in the Save Preset dialog box, lined up with the radio button text */
    var PRESET_NAME_INDENT = 20;

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

    var LABELS = {
        palette: {
            title: { ja: "段落境界線", en: "Paragraph Rules" }
        },
        panel: {
            ruleAbove: { ja: "前境界線", en: "Rule Above" },
            ruleBelow: { ja: "後境界線", en: "Rule Below" },
            indent:    { ja: "インデント", en: "Indents" }
        },
        fieldLabel: {
            preset:      { ja: "プリセット", en: "Preset" },
            presetName:  { ja: "名前", en: "Name" },
            weight:      { ja: "線幅", en: "Weight" },
            type:        { ja: "種類", en: "Type" },
            color:       { ja: "カラー", en: "Color" },
            tint:        { ja: "濃淡", en: "Tint" },
            gapColor:    { ja: "間隔カラー", en: "Gap Color" },
            gapTint:     { ja: "間隔の濃淡", en: "Gap Tint" },
            width:       { ja: "幅", en: "Width" },
            offset:      { ja: "オフセット", en: "Offset" },
            leftIndent:   { ja: "左インデント", en: "Left Indent" },
            rightIndent:  { ja: "右インデント", en: "Right Indent" },
            topIndent:    { ja: "上インデント", en: "Top Indent" },
            bottomIndent: { ja: "下インデント", en: "Bottom Indent" }
        },
        checkbox: {
            ruleOn:       { ja: "境界線を挿入", en: "Rule On" },
            overprint:    { ja: "線オーバープリント", en: "Overprint Stroke" },
            gapOverprint: { ja: "間隔オーバープリント", en: "Overprint Gap" },
            keepInFrame:  { ja: "フレームに収める", en: "Keep In Frame" },
            preview:      { ja: "プレビュー", en: "Preview" }
        },
        ruleWidth: {
            column: { ja: "段", en: "Column" },
            text:   { ja: "テキスト", en: "Text" }
        },
        swatch: {
            textColor:    { ja: "(テキストカラー)", en: "(Text Color)" },
            none:         { ja: "[なし]", en: "[None]" },
            paper:        { ja: "[紙色]", en: "[Paper]" },
            black:        { ja: "[黒]", en: "[Black]" },
            registration: { ja: "[レジストレーション]", en: "[Registration]" }
        },
        button: {
            ok:           { ja: "OK", en: "OK" },
            cancel:       { ja: "キャンセル", en: "Cancel" },
            save:         { ja: "保存", en: "Save" },
            savePreset:   { ja: "保存...", en: "Save..." },
            deletePreset: { ja: "削除", en: "Delete" }
        },
        presetDialog: {
            title:          { ja: "プリセットを保存", en: "Save Preset" },
            overwrite:      { ja: "上書き保存", en: "Overwrite" },
            overwriteNamed: { ja: "上書き保存（%1）", en: "Overwrite (%1)" },
            saveAs:         { ja: "名前を付けて保存", en: "Save As" }
        },
        confirm: {
            replacePreset: { ja: "プリセット「%1」はすでにあります。置き換えますか？", en: "A preset named \"%1\" already exists. Replace it?" },
            deletePreset:  { ja: "プリセット「%1」を削除しますか？", en: "Delete the preset \"%1\"?" }
        },
        tooltip: {
            preset: {
                ja: "選んだ設定を前境界線・後境界線の欄に入れます。",
                en: "Fills in the rule above and rule below with the chosen settings."
            },
            savePreset: {
                ja: "今の設定をプリセットとして保存します（上書き保存・名前を付けて保存）。",
                en: "Saves the current settings as a preset (overwrite or save as)."
            },
            deletePreset: {
                ja: "選んでいる保存済みのプリセットを削除します。",
                en: "Deletes the selected saved preset."
            },
            offset: {
                ja: "前境界線はベースラインから上へ、後境界線は下へずらす距離です。",
                en: "Distance from the baseline: upward for the rule above, downward for the rule below."
            },
            indent: {
                ja: "境界線の端を内側へ詰める量です。",
                en: "How far to pull in the end of the rule."
            },
            linkIndents: {
                ja: "段落の2つのインデントを同じ値で連動させます。",
                en: "Keeps the paragraph's two indents at the same value."
            },
            keepInFrame: {
                ja: "前境界線をフレームの上端からはみ出さないようにします。",
                en: "Keeps the rule above inside the top of the frame."
            },
            preview: {
                ja: "設定をドキュメントに仮に適用して表示します。［キャンセル］・Esc・クローズボックスで閉じると元に戻ります。",
                en: "Shows the settings applied to the document. Closing with Cancel, Esc or the close box restores the original."
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
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectText: { ja: "テキストまたはテキストフレームを選択してください。", en: "Please select text or a text frame." },
            previewFailed: {
                ja: "プレビューできませんでした。\n%1",
                en: "Could not show the preview.\n%1"
            },
            emptyPresetName: { ja: "名前を入力してください。", en: "Please enter a name." },
            savePresetFailed: {
                ja: "プリセットを保存できませんでした。\n%1",
                en: "Could not save the presets.\n%1"
            },
            applyFailed: {
                ja: "段落境界線を設定できませんでした。\n%1",
                en: "Could not set the paragraph rules.\n%1"
            }
        },
        undo: {
            preview: { ja: "段落境界線のプレビュー", en: "Paragraph Rules Preview" },
            apply:   { ja: "段落境界線", en: "Paragraph Rules" }
        }
    };

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

    // =========================================
    // 単位 / Units
    // =========================================

    /**
     * @typedef {object} UnitInfo
     * @property {string} unit 入力欄に付ける単位（先頭に空白。例 " mm"）
     * @property {number} pointsPerUnit 1 単位あたりのポイント数
     * @property {number} step 長さの欄の∧∨の増減量
     * @property {number} weightStep 線幅の欄の∧∨の増減量
     */

    /**
     * InDesign の単位の設定から、入力欄の単位と換算係数を返す
     * @param {MeasurementUnits} measurementUnit viewPreferences の単位
     * @returns {UnitInfo} 単位の情報
     */
    function getUnitInfo(measurementUnit) {
        switch (measurementUnit) {
            case MeasurementUnits.MILLIMETERS:
                return { unit: " mm", pointsPerUnit: 72 / 25.4, step: 1, weightStep: 0.05 };
            case MeasurementUnits.CENTIMETERS:
                return { unit: " cm", pointsPerUnit: 72 / 2.54, step: 0.1, weightStep: 0.01 };
            case MeasurementUnits.INCHES:
            case MeasurementUnits.INCHES_DECIMAL:
                return { unit: " in", pointsPerUnit: 72, step: 0.1, weightStep: 0.01 };
            case MeasurementUnits.PICAS:
                return { unit: " p", pointsPerUnit: 12, step: 1, weightStep: 0.1 };
            case MeasurementUnits.Q:
                return { unit: " Q", pointsPerUnit: 72 / 25.4 / 4, step: 1, weightStep: 0.25 };
            case MeasurementUnits.HA:
                return { unit: " H", pointsPerUnit: 72 / 25.4 / 4, step: 1, weightStep: 0.25 };
            case MeasurementUnits.CICEROS:
                return { unit: " c", pointsPerUnit: 12.7878, step: 1, weightStep: 0.1 };
            case MeasurementUnits.AGATES:
                return { unit: " ag", pointsPerUnit: 72 / 14, step: 1, weightStep: 0.1 };
            default:
                return { unit: " pt", pointsPerUnit: 1, step: 1, weightStep: 0.25 };
        }
    }

    /**
     * ポイントの値を、入力欄に入れる単位付きの文字列にする
     * @param {number} points ポイントの値
     * @param {UnitInfo} unitInfo 表示する単位
     * @returns {string} 例 "0.25 mm"
     */
    function formatLength(points, unitInfo) {
        return formatStepperNumber(points / unitInfo.pointsPerUnit) + unitInfo.unit;
    }

    /**
     * 入力欄の文字列を、DOM に代入するポイントの文字列にする（素の数値は現在の単位で解釈されるため）
     * @param {string} fieldText 入力欄の文字列（例 "0.25 mm"）
     * @param {UnitInfo} unitInfo 入力欄の単位
     * @returns {string} 例 "0.7087pt"
     */
    function toPointString(fieldText, unitInfo) {
        var points = parseFloat(fieldText) * unitInfo.pointsPerUnit;
        return String(Math.round(points * 10000) / 10000) + "pt";
    }

    // =========================================
    // スウォッチ・線の種類・幅の候補 / Choices
    // =========================================

    /* カラーの候補で「テキストカラー」を表すキー / Key that stands for the text color in the color choices */
    var TEXT_COLOR_KEY = "\u0000textColor";

    /* InDesign が「テキストカラー」に使う値 / Value InDesign uses for the text color */
    var TEXT_COLOR_VALUE = "Text Color";

    /* 特別扱いするスウォッチ名（角括弧・大文字小文字は無視して照合）/ Special swatch names (brackets and case are ignored) */
    var SPECIAL_SWATCH_NAMES = {
        none:         ["None", "なし"],
        paper:        ["Paper", "紙色", "紙"],
        black:        ["Black", "ブラック", "黒"],
        registration: ["Registration", "レジストレーション", "トンボ用"]
    };

    /* 実線の線の種類の名前（間隔カラーを使えない）。表示言語の名前も足す / Names of the solid stroke type, which has no gap, including the UI-language name */
    var SOLID_STROKE_NAMES = ["Solid", "実線", app.translateKeyString("$ID/Solid")];

    /* 色見本に使う RGB（0〜1）/ RGB values (0-1) used for the swatch chip */
    var CHIP_RGB_BLACK    = [0, 0, 0];
    var CHIP_RGB_WHITE    = [1, 1, 1];
    var CHIP_RGB_FALLBACK = [0.5, 0.5, 0.5];
    var CHIP_FRAME_COLOR  = [0.5, 0.5, 0.5, 1];
    var CHIP_NONE_COLOR   = [0.9, 0.1, 0.1, 1];

    /**
     * @typedef {object} ColorChoice
     * @property {string} key 選択の照合に使うキー（スウォッチ名、またはテキストカラーのキー）
     * @property {string} displayName ドロップダウンに表示する名前
     * @property {string} swatchKind 特別なスウォッチの種別（"none" / "paper" / "black" / "registration" / "textColor"）。通常のカラーは空文字
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
     * @returns {string} "none" / "paper" / "black" / "registration"。どれでもなければ空文字
     */
    function getSpecialSwatchKind(swatchName) {
        var normalizedName = normalizeSwatchName(swatchName);
        for (var swatchKind in SPECIAL_SWATCH_NAMES) {
            if (!SPECIAL_SWATCH_NAMES.hasOwnProperty(swatchKind)) continue;
            var names = SPECIAL_SWATCH_NAMES[swatchKind];
            for (var i = 0; i < names.length; i++) {
                if (normalizeSwatchName(names[i]) === normalizedName) return swatchKind;
            }
        }
        return "";
    }

    /**
     * ドキュメントのスウォッチからカラーの候補を作る
     * @param {Document} targetDoc 対象のドキュメント
     * @param {boolean} includeTextColor 先頭に「テキストカラー」を入れるなら true（間隔カラーには無い）
     * @returns {Array<ColorChoice>} カラーの候補
     */
    function buildColorChoices(targetDoc, includeTextColor) {
        var colorChoices = [];
        if (includeTextColor) {
            colorChoices.push({ key: TEXT_COLOR_KEY, displayName: getLabel("swatch.textColor"), swatchKind: "textColor" });
        }
        var swatches = targetDoc.swatches;
        for (var i = 0; i < swatches.length; i++) {
            var swatchName = String(swatches[i].name);
            var swatchKind = getSpecialSwatchKind(swatchName);
            colorChoices.push({
                key: swatchName,
                displayName: swatchKind ? getLabel("swatch." + swatchKind) : swatchName,
                swatchKind: swatchKind
            });
        }
        return colorChoices;
    }

    /**
     * 候補の中からキーに一致する位置を求める
     * @param {Array<{key: string}>} choices 候補
     * @param {string} key 探すキー
     * @returns {number} 見つかった位置。無ければ -1
     */
    function findChoiceIndex(choices, key) {
        for (var i = 0; i < choices.length; i++) {
            if (choices[i].key === key) return i;
        }
        return -1;
    }

    /**
     * 線の種類の候補を作る（ドキュメントの線種）
     * @param {Document} targetDoc 対象のドキュメント
     * @returns {Array<{key: string, displayName: string}>} 線の種類の候補
     */
    function buildStrokeTypeChoices(targetDoc) {
        var strokeTypeChoices = [];
        var strokeStyles = targetDoc.strokeStyles;
        for (var i = 0; i < strokeStyles.length; i++) {
            var strokeName = String(strokeStyles[i].name);
            strokeTypeChoices.push({ key: strokeName, displayName: strokeName });
        }
        return strokeTypeChoices;
    }

    /**
     * 線の種類が実線かどうかを判定する（実線には間隔が無い）
     * @param {string} strokeName 線の種類の名前
     * @returns {boolean} 実線なら true
     */
    function isSolidStroke(strokeName) {
        for (var i = 0; i < SOLID_STROKE_NAMES.length; i++) {
            if (SOLID_STROKE_NAMES[i] === strokeName) return true;
        }
        return false;
    }

    /**
     * 幅の候補を作る（段・テキスト）
     * @returns {Array<{key: string, displayName: string}>} 幅の候補
     */
    function buildRuleWidthChoices() {
        return [
            { key: "COLUMN_WIDTH", displayName: getLabel("ruleWidth.column") },
            { key: "TEXT_WIDTH", displayName: getLabel("ruleWidth.text") }
        ];
    }

    /**
     * RuleWidth の値から幅の候補のキーを求める
     * @param {RuleWidth} ruleWidth 段落の幅の値
     * @returns {string} 候補のキー
     */
    function getRuleWidthKey(ruleWidth) {
        return (ruleWidth === RuleWidth.TEXT_WIDTH) ? "TEXT_WIDTH" : "COLUMN_WIDTH";
    }

    // =========================================
    // 色見本 / Swatch chip
    // =========================================

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

        /* 濃淡スウォッチは元のカラーで見せる。グラデーションや混合インキは色値を読めない
           / Tint swatches show their base color; gradients and mixed inks have no readable color value */
        try {
            var colorSource = (swatch.hasOwnProperty("baseColor") && swatch.baseColor) ? swatch.baseColor : swatch;
            var colorValues = colorSource.colorValue;
            if (colorSource.space === ColorSpace.RGB) {
                return [colorValues[0] / 255, colorValues[1] / 255, colorValues[2] / 255];
            }
            if (colorSource.space === ColorSpace.CMYK) {
                var blackValue = colorValues[3] / 100;
                return [
                    (1 - colorValues[0] / 100) * (1 - blackValue),
                    (1 - colorValues[1] / 100) * (1 - blackValue),
                    (1 - colorValues[2] / 100) * (1 - blackValue)
                ];
            }
        } catch (e) { }
        return CHIP_RGB_FALLBACK;
    }

    /**
     * ドロップダウンの左に置く色見本を作る（onDraw で描く。［なし］は赤い斜線）
     * @param {Group} parent 追加先の行
     * @returns {Group} 色見本（.chipRGB と .isNoneSwatch を描画に使う）
     */
    function addSwatchChip(parent) {
        var swatchChip = parent.add("group");
        swatchChip.preferredSize = [SWATCH_CHIP_SIZE, SWATCH_CHIP_SIZE];
        swatchChip.minimumSize   = [SWATCH_CHIP_SIZE, SWATCH_CHIP_SIZE];
        swatchChip.maximumSize   = [SWATCH_CHIP_SIZE, SWATCH_CHIP_SIZE];
        swatchChip.chipRGB = CHIP_RGB_FALLBACK;
        swatchChip.isNoneSwatch = false;

        swatchChip.onDraw = function () {
            var chipGraphics = this.graphics;
            var chipWidth = this.size[0];
            var chipHeight = this.size[1];
            /* 無効の間は薄く描く / draw faintly while disabled */
            var chipAlpha = isStepperEnabledInTree(this) ? 1 : 0.35;
            var chipRGB = this.chipRGB;

            /* newPath() を呼ばないとパスが前回の描画に足されていく / without newPath() the paths accumulate */
            chipGraphics.newPath();
            chipGraphics.rectPath(0, 0, chipWidth, chipHeight);
            chipGraphics.fillPath(chipGraphics.newBrush(chipGraphics.BrushType.SOLID_COLOR,
                [chipRGB[0], chipRGB[1], chipRGB[2], chipAlpha]));

            chipGraphics.newPath();
            chipGraphics.rectPath(0.5, 0.5, chipWidth - 1, chipHeight - 1);
            chipGraphics.strokePath(chipGraphics.newPen(chipGraphics.PenType.SOLID_COLOR,
                [CHIP_FRAME_COLOR[0], CHIP_FRAME_COLOR[1], CHIP_FRAME_COLOR[2], chipAlpha], 1));

            if (this.isNoneSwatch) {
                chipGraphics.newPath();
                chipGraphics.moveTo(chipWidth - 1, 1);
                chipGraphics.lineTo(1, chipHeight - 1);
                chipGraphics.strokePath(chipGraphics.newPen(chipGraphics.PenType.SOLID_COLOR,
                    [CHIP_NONE_COLOR[0], CHIP_NONE_COLOR[1], CHIP_NONE_COLOR[2], chipAlpha], 2));
            }
        };
        return swatchChip;
    }

    /**
     * 色見本を選択中のカラーで描き直す
     * @param {Group} swatchChip 色見本
     * @param {ColorChoice|null} colorChoice 選択中のカラー
     * @param {Swatch|null} textColorSwatch テキストカラーとして見せるスウォッチ
     * @returns {void}
     */
    function paintSwatchChip(swatchChip, colorChoice, textColorSwatch) {
        if (!colorChoice) {
            swatchChip.chipRGB = CHIP_RGB_FALLBACK;
            swatchChip.isNoneSwatch = false;
        } else {
            var swatch = (colorChoice.swatchKind === "textColor") ? textColorSwatch : findSwatch(colorChoice.key);
            swatchChip.chipRGB = getSwatchChipRGB(swatch);
            swatchChip.isNoneSwatch = (colorChoice.swatchKind === "none");
        }
        /* group には notify() が無いので、隠して再表示して描き直す / groups have no notify(), so hide and show to redraw */
        swatchChip.hide();
        swatchChip.show();
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

    // =========================================
    // 選択範囲と境界線の読み取り / Selection & reading rules
    // =========================================

    /**
     * 選択から、段落境界線を設定するテキストを集める
     * （テキスト・挿入点はそのまま、テキストフレームはその中のテキスト、セルはセルのテキスト）
     * @param {Array} selectionItems app.selection
     * @returns {Array<Text>} 対象のテキスト
     */
    function collectTextTargets(selectionItems) {
        var textTargets = [];
        for (var i = 0; i < selectionItems.length; i++) {
            var selectionItem = selectionItems[i];
            var typeName = selectionItem.constructor.name;
            if (typeName === "TextFrame") {
                textTargets.push(selectionItem.texts[0]);
            } else if (typeName === "Cell" || typeName === "Table") {
                /* 複数セルの選択も Cell 1つで返るので、.cells で1セルずつに分ける / a multi-cell selection comes back as one Cell, so split it with .cells */
                var cells = selectionItem.cells;
                for (var j = 0; j < cells.length; j++) {
                    textTargets.push(cells[j].texts[0]);
                }
            } else if (selectionItem.hasOwnProperty("ruleAbove") && selectionItem.hasOwnProperty("paragraphs")) {
                textTargets.push(selectionItem);
            }
        }
        return textTargets;
    }

    /**
     * @typedef {object} RuleModel
     * @property {boolean} ruleOn 境界線を挿入するか
     * @property {string} weightText 線幅（単位付き）
     * @property {string} typeKey 線の種類の名前
     * @property {string} colorKey カラーのキー
     * @property {string} tintText 濃淡（%付き）
     * @property {boolean} overprint 線オーバープリント
     * @property {string} gapColorKey 間隔カラーのキー
     * @property {string} gapTintText 間隔の濃淡（%付き）
     * @property {boolean} gapOverprint 間隔オーバープリント
     * @property {string} widthKey 幅（RuleWidth のキー）
     * @property {string} offsetText オフセット（単位付き）
     * @property {string} leftIndentText 左 / 上インデント（単位付き）
     * @property {string} rightIndentText 右 / 下インデント（単位付き）
     * @property {boolean} keepInFrame フレームに収める（前境界線だけ）
     */

    /**
     * @typedef {object} IndentModel 段落のインデント
     * @property {string} leftIndentText 左 / 上インデント（単位付き）
     * @property {string} rightIndentText 右 / 下インデント（単位付き）
     */

    /**
     * 段落の濃淡を、入力欄に入れる文字列にする（自動の -1 は 100%）
     * @param {number} tint 段落の濃淡
     * @returns {string} 例 "100%"
     */
    function formatTint(tint) {
        if (isNaN(tint) || tint < 0) tint = TINT_AUTO_DISPLAY;
        return formatStepperNumber(tint) + "%";
    }

    /**
     * カラーの値を候補のキーにする
     * @param {Swatch|string} colorValue 段落のカラー（テキストカラーは文字列）
     * @returns {string} 候補のキー
     */
    function getColorKey(colorValue) {
        if (typeof colorValue === "string") return TEXT_COLOR_KEY;
        return (colorValue && colorValue.isValid) ? String(colorValue.name) : "";
    }

    /**
     * 段落から前境界線または後境界線の設定を読み取る（測定値は pt で読む）
     * @param {Paragraph} paragraph 読み取る段落
     * @param {string} ruleSide "Above" または "Below"
     * @param {{stroke: UnitInfo, horizontal: UnitInfo, vertical: UnitInfo}} fieldUnits 入力欄の単位
     * @returns {RuleModel} 読み取った設定
     */
    function readRuleModel(paragraph, ruleSide, fieldUnits) {
        var prefix = "rule" + ruleSide;
        var strokeType = paragraph[prefix + "Type"];
        return {
            ruleOn: paragraph[prefix] === true,
            weightText: formatLength(paragraph[prefix + "LineWeight"], fieldUnits.stroke),
            typeKey: (strokeType && strokeType.isValid) ? String(strokeType.name) : "",
            colorKey: getColorKey(paragraph[prefix + "Color"]),
            tintText: formatTint(paragraph[prefix + "Tint"]),
            overprint: paragraph[prefix + "Overprint"] === true,
            gapColorKey: getColorKey(paragraph[prefix + "GapColor"]),
            gapTintText: formatTint(paragraph[prefix + "GapTint"]),
            gapOverprint: paragraph[prefix + "GapOverprint"] === true,
            widthKey: getRuleWidthKey(paragraph[prefix + "Width"]),
            offsetText: formatLength(paragraph[prefix + "Offset"], fieldUnits.vertical),
            leftIndentText: formatLength(paragraph[prefix + "LeftIndent"], fieldUnits.horizontal),
            rightIndentText: formatLength(paragraph[prefix + "RightIndent"], fieldUnits.horizontal),
            keepInFrame: (ruleSide === "Above") ? paragraph.keepRuleAboveInFrame === true : false
        };
    }

    /* 文字サイズを数値で読めなかったときに 1em とみなす大きさ（pt）/ Point size assumed for 1em when it cannot be read as a number */
    var FALLBACK_POINT_SIZE = 12;

    /**
     * 段落の文字サイズを pt の数値で読む（数値で読めなければ FALLBACK_POINT_SIZE）
     * @param {Paragraph} paragraph 読み取る段落
     * @returns {number} 文字サイズ（pt）
     */
    function readPointSize(paragraph) {
        var pointSize = parseFloat(paragraph.pointSize);
        return (isFinite(pointSize) && pointSize > 0) ? pointSize : FALLBACK_POINT_SIZE;
    }

    /**
     * 先頭の段落から前境界線・後境界線の設定、段落のインデント、文字サイズを読み取る。
     * 測定値は単位の設定によって返る単位が変わるので、読み取る間だけスクリプトの単位を pt にする
     * @param {Paragraph} paragraph 読み取る段落
     * @param {{stroke: UnitInfo, horizontal: UnitInfo, vertical: UnitInfo}} fieldUnits 入力欄の単位
     * @returns {{models: {Above: RuleModel, Below: RuleModel}, indentModel: IndentModel, pointSize: number}} 読み取った設定・段落のインデント・文字サイズ（pt）
     */
    function readParagraphSettings(paragraph, fieldUnits) {
        var originalUnit = app.scriptPreferences.measurementUnit;
        app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
        try {
            return {
                models: {
                    Above: readRuleModel(paragraph, "Above", fieldUnits),
                    Below: readRuleModel(paragraph, "Below", fieldUnits)
                },
                indentModel: {
                    leftIndentText: formatLength(paragraph.leftIndent, fieldUnits.horizontal),
                    rightIndentText: formatLength(paragraph.rightIndent, fieldUnits.horizontal)
                },
                pointSize: readPointSize(paragraph)
            };
        } finally {
            app.scriptPreferences.measurementUnit = originalUnit;
        }
    }

    /**
     * 設定を複製する（変更の有無を比べる元として控えるため）
     * @param {RuleModel} ruleModel 複製する設定
     * @returns {RuleModel} 複製
     */
    function cloneRuleModel(ruleModel) {
        var copiedModel = {};
        for (var fieldKey in ruleModel) {
            if (ruleModel.hasOwnProperty(fieldKey)) copiedModel[fieldKey] = ruleModel[fieldKey];
        }
        return copiedModel;
    }

    // =========================================
    // 境界線の書き込み / Writing rules
    // =========================================

    /* 境界線の向き。プロパティ名の "rule" に続く部分 / Rule sides, the part of the property names after "rule" */
    var RULE_SIDES = ["Above", "Below"];

    /**
     * 値をそのまま返す（チェックボックスの値など）
     * @param {*} value 設定の値
     * @returns {*} 同じ値
     */
    function keepValue(value) {
        return value;
    }

    /**
     * 「100%」の形の濃淡を数値にする
     * @param {string} tintText 濃淡の文字列
     * @returns {number} 濃淡
     */
    function toTintNumber(tintText) {
        return parseFloat(tintText);
    }

    /**
     * 長さの欄を、指定した向きの単位で pt の文字列にする変換を作る
     * @param {string} unitAxis "stroke" / "horizontal" / "vertical"
     * @returns {function(string, object): string} 入力欄の文字列と単位から pt の文字列を返す関数
     */
    function makeLengthConverter(unitAxis) {
        return function (fieldText, fieldUnits) {
            return toPointString(fieldText, fieldUnits[unitAxis]);
        };
    }

    /**
     * 設定の項目と、書き込む段落のプロパティ・値の作り方の対応
     * suffix は "rule" + Above/Below に続くプロパティ名の後半。keepInFrame は前境界線だけなので別に扱う
     */
    var RULE_FIELD_WRITERS = [
        { field: "ruleOn",          suffix: "",             toValue: keepValue },
        { field: "weightText",      suffix: "LineWeight",   toValue: makeLengthConverter("stroke") },
        { field: "typeKey",         suffix: "Type",         toValue: function (strokeName) { return app.activeDocument.strokeStyles.itemByName(strokeName); } },
        { field: "colorKey",        suffix: "Color",        toValue: toColorValue },
        { field: "tintText",        suffix: "Tint",         toValue: toTintNumber },
        { field: "overprint",       suffix: "Overprint",    toValue: keepValue },
        { field: "gapColorKey",     suffix: "GapColor",     toValue: toColorValue },
        { field: "gapTintText",     suffix: "GapTint",      toValue: toTintNumber },
        { field: "gapOverprint",    suffix: "GapOverprint", toValue: keepValue },
        { field: "widthKey",        suffix: "Width",        toValue: function (widthKey) { return RuleWidth[widthKey]; } },
        { field: "offsetText",      suffix: "Offset",       toValue: makeLengthConverter("vertical") },
        { field: "leftIndentText",  suffix: "LeftIndent",   toValue: makeLengthConverter("horizontal") },
        { field: "rightIndentText", suffix: "RightIndent",  toValue: makeLengthConverter("horizontal") }
    ];

    /* 段落のインデントの項目と、書き込む段落のプロパティの対応 / Paragraph indent items and the properties they write */
    var INDENT_FIELDS = [
        { field: "leftIndentText",  property: "leftIndent" },
        { field: "rightIndentText", property: "rightIndent" }
    ];

    /**
     * カラーのキーを、段落に代入する値にする
     * @param {string} colorKey カラーのキー
     * @returns {Swatch|string} スウォッチ、またはテキストカラーの文字列
     */
    function toColorValue(colorKey) {
        if (colorKey === TEXT_COLOR_KEY) return TEXT_COLOR_VALUE;
        return app.activeDocument.swatches.itemByName(colorKey);
    }

    /**
     * 読み取ったときから変わった項目だけを、段落に代入するプロパティにまとめる
     * @param {object} paletteState 状態オブジェクト（models / initialModels / indentModel / initialIndentModel / fieldUnits）
     * @returns {object} プロパティ名と値。変更が無ければ空のオブジェクト
     */
    function buildChangedProperties(paletteState) {
        var changedProperties = {};
        for (var i = 0; i < RULE_SIDES.length; i++) {
            var ruleSide = RULE_SIDES[i];
            var ruleModel = paletteState.models[ruleSide];
            var initialModel = paletteState.initialModels[ruleSide];
            for (var j = 0; j < RULE_FIELD_WRITERS.length; j++) {
                var fieldWriter = RULE_FIELD_WRITERS[j];
                if (ruleModel[fieldWriter.field] === initialModel[fieldWriter.field]) continue;
                changedProperties["rule" + ruleSide + fieldWriter.suffix] = fieldWriter.toValue(ruleModel[fieldWriter.field], paletteState.fieldUnits);
            }
        }
        if (paletteState.models.Above.keepInFrame !== paletteState.initialModels.Above.keepInFrame) {
            changedProperties.keepRuleAboveInFrame = paletteState.models.Above.keepInFrame;
        }
        for (var k = 0; k < INDENT_FIELDS.length; k++) {
            var indentField = INDENT_FIELDS[k];
            if (paletteState.indentModel[indentField.field] === paletteState.initialIndentModel[indentField.field]) continue;
            changedProperties[indentField.property] = toPointString(paletteState.indentModel[indentField.field], paletteState.fieldUnits.horizontal);
        }
        return changedProperties;
    }

    /**
     * プロパティのまとまりが空かどうかを判定する
     * @param {object} properties プロパティ名と値
     * @returns {boolean} 空なら true
     */
    function isEmptyProperties(properties) {
        for (var propertyName in properties) {
            if (properties.hasOwnProperty(propertyName)) return false;
        }
        return true;
    }

    /**
     * 対象のテキストすべてに、変わった項目を1つの取り消し単位で書き込む
     * @param {Array<Text>} textTargets 対象のテキスト
     * @param {object} changedProperties プロパティ名と値
     * @param {string} undoName 取り消し履歴に出す名前
     * @returns {void}
     */
    function applyRuleProperties(textTargets, changedProperties, undoName) {
        app.doScript(function () {
            for (var i = 0; i < textTargets.length; i++) {
                textTargets[i].properties = changedProperties;
            }
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.FAST_ENTIRE_SCRIPT, undoName);
    }

    // =========================================
    // プリセット / Presets
    // =========================================

    /* プリセットの width と、幅の候補のキーの対応 / Preset width values mapped to rule width keys */
    var PRESET_WIDTH_KEYS = { column: "COLUMN_WIDTH", text: "TEXT_WIDTH" };

    /**
     * プリセットの長さを、入力欄に入れる単位付きの文字列にする
     * @param {number|string} presetLength 数値なら pt、"0.2em" の形なら文字サイズに対する倍率
     * @param {UnitInfo} unitInfo 入力欄の単位
     * @param {number} emPoints 1em の大きさ（pt）
     * @returns {string} 例 "0.18 mm"
     */
    function formatPresetLength(presetLength, unitInfo, emPoints) {
        var lengthText = String(presetLength);
        var isEmLength = /em$/i.test(lengthText);
        /* ExtendScript の parseFloat は「1.1em」の e を指数と読んで NaN を返すので、em を外してから読む
           / ExtendScript's parseFloat reads the "e" of "1.1em" as an exponent and returns NaN, so strip "em" first */
        var amount = parseFloat(lengthText.replace(/em$/i, ""));
        return formatLength(isEmLength ? amount * emPoints : amount, unitInfo);
    }

    /**
     * プリセットのカラーを、カラーの候補のキーにする
     * @param {Array<ColorChoice>} colorChoices カラーの候補
     * @param {string} presetColor "textColor" / "none" / "paper" / "black" / "registration" またはスウォッチ名
     * @returns {string|null} 候補のキー。候補に無ければ null
     */
    function resolvePresetColorKey(colorChoices, presetColor) {
        for (var i = 0; i < colorChoices.length; i++) {
            if (colorChoices[i].swatchKind === presetColor) return colorChoices[i].key;
        }
        return (findChoiceIndex(colorChoices, presetColor) >= 0) ? presetColor : null;
    }

    /**
     * プリセットの線の種類を、線の種類の候補のキーにする
     * @param {Array<{key: string}>} strokeTypeChoices 線の種類の候補
     * @param {string} presetType "solid" または線種の名前
     * @returns {string|null} 候補のキー。候補に無ければ null
     */
    function resolvePresetTypeKey(strokeTypeChoices, presetType) {
        for (var i = 0; i < strokeTypeChoices.length; i++) {
            var strokeName = strokeTypeChoices[i].key;
            if (presetType === "solid" ? isSolidStroke(strokeName) : strokeName === presetType) return strokeName;
        }
        return null;
    }

    /**
     * プリセットの1本分を、既定値で補って設定に書き込む（見つからないカラー・線種は今の値のまま）
     * @param {object} paletteState 状態オブジェクト
     * @param {RuleModel} ruleModel 書き込む先の設定
     * @param {object} presetRule プリセットの Above または Below
     * @returns {void}
     */
    function applyPresetRule(paletteState, ruleModel, presetRule) {
        var spec = {};
        var key;
        for (key in PRESET_RULE_DEFAULTS) {
            if (PRESET_RULE_DEFAULTS.hasOwnProperty(key)) spec[key] = PRESET_RULE_DEFAULTS[key];
        }
        for (key in presetRule) {
            if (presetRule.hasOwnProperty(key)) spec[key] = presetRule[key];
        }

        var fieldUnits = paletteState.fieldUnits;
        var emPoints = paletteState.emPoints;
        var typeKey = resolvePresetTypeKey(paletteState.strokeTypeChoices, spec.type);
        var colorKey = resolvePresetColorKey(paletteState.colorChoices, spec.color);
        var gapColorKey = resolvePresetColorKey(paletteState.gapColorChoices, spec.gapColor);

        ruleModel.ruleOn = true;
        ruleModel.weightText = formatPresetLength(spec.weight, fieldUnits.stroke, emPoints);
        if (typeKey !== null) ruleModel.typeKey = typeKey;
        if (colorKey !== null) ruleModel.colorKey = colorKey;
        ruleModel.tintText = formatTint(spec.tint);
        ruleModel.overprint = spec.overprint;
        if (gapColorKey !== null) ruleModel.gapColorKey = gapColorKey;
        ruleModel.gapTintText = formatTint(spec.gapTint);
        ruleModel.gapOverprint = spec.gapOverprint;
        ruleModel.widthKey = PRESET_WIDTH_KEYS[spec.width] || PRESET_WIDTH_KEYS.column;
        ruleModel.offsetText = formatPresetLength(spec.offset, fieldUnits.vertical, emPoints);
        ruleModel.leftIndentText = formatPresetLength(spec.leftIndent, fieldUnits.horizontal, emPoints);
        ruleModel.rightIndentText = formatPresetLength(spec.rightIndent, fieldUnits.horizontal, emPoints);
        ruleModel.keepInFrame = spec.keepInFrame;
    }

    /**
     * プリセットを前境界線・後境界線の設定に入れ、パネルとプレビューを更新する
     * （プリセットに無い境界線は［境界線を挿入］だけオフにし、ほかの項目は今の値のまま）
     * @param {object} paletteState 状態オブジェクト
     * @param {object} preset RULE_PRESETS の1項目
     * @returns {void}
     */
    function applyPreset(paletteState, preset) {
        for (var i = 0; i < RULE_SIDES.length; i++) {
            var ruleSide = RULE_SIDES[i];
            var ruleModel = paletteState.models[ruleSide];
            if (preset[ruleSide]) applyPresetRule(paletteState, ruleModel, preset[ruleSide]);
            else ruleModel.ruleOn = false;
            loadRuleModel(paletteState, ruleSide);
        }
        refreshPreview(paletteState);
    }

    /**
     * プリセットの表示名を返す（組み込みは { ja, en }、保存したものは文字列）
     * @param {object} preset プリセット
     * @returns {string} 表示名
     */
    function getPresetDisplayName(preset) {
        return (typeof preset.name === "string") ? preset.name : getLabel(preset.name);
    }

    /**
     * プリセットのドロップダウンを、組み込み・区切り線・保存したプリセットの順で作り直す
     * @param {object} paletteState 状態オブジェクト
     * @param {object|null} selectedPreset 選んだ状態にするプリセット（null なら未選択）
     * @returns {void}
     */
    function fillPresetDropdown(paletteState, selectedPreset) {
        var presetDropdown = paletteState.ui.presetDropdown;
        var userPresets = paletteState.userPresets;
        var selectedIndex = -1;
        var i;

        /**
         * プリセットを1つ足す
         * @param {object} preset プリセット
         * @param {boolean} isUser 保存したプリセットなら true
         * @returns {void}
         */
        function addPresetItem(preset, isUser) {
            if (preset === selectedPreset) selectedIndex = presetDropdown.items.length;
            presetDropdown.add("item", getPresetDisplayName(preset));
            presetDropdown.entries.push({ preset: preset, isUser: isUser });
        }

        /* 作り直しや選択の代入で onChange が呼ばれるので、その間はプリセットを適用しない / hold back onChange while rebuilding */
        paletteState.isLoading = true;
        try {
            presetDropdown.removeAll();
            presetDropdown.entries = [];
            for (i = 0; i < RULE_PRESETS.length; i++) addPresetItem(RULE_PRESETS[i], false);
            if (userPresets.length > 0) {
                presetDropdown.add("separator");
                presetDropdown.entries.push(null);
            }
            for (i = 0; i < userPresets.length; i++) addPresetItem(userPresets[i], true);
            presetDropdown.selection = (selectedIndex >= 0) ? selectedIndex : null;
        } finally {
            paletteState.isLoading = false;
        }
        refreshPresetButtons(paletteState);
    }

    /**
     * ドロップダウンで選んでいるプリセットを返す
     * @param {DropDownList} presetDropdown プリセットのドロップダウン
     * @returns {{preset: object, isUser: boolean}|null} 選んでいるプリセット。未選択なら null
     */
    function getSelectedPresetEntry(presetDropdown) {
        return presetDropdown.selection ? presetDropdown.entries[presetDropdown.selection.index] : null;
    }

    /**
     * ［削除］を、保存したプリセットを選んでいるときだけ有効にする
     * @param {object} paletteState 状態オブジェクト
     * @returns {void}
     */
    function refreshPresetButtons(paletteState) {
        var selectedEntry = getSelectedPresetEntry(paletteState.ui.presetDropdown);
        paletteState.ui.btnDeletePreset.enabled = !!(selectedEntry && selectedEntry.isUser);
    }

    /**
     * パレットの最上部にプリセットの行（ドロップダウン・［保存...］・［削除］）を左右中央に追加する
     * @param {Window} rulePalette 追加先のパレット
     * @param {object} paletteState 状態オブジェクト
     * @returns {void}
     */
    function addPresetRow(rulePalette, paletteState) {
        var ui = paletteState.ui;
        var presetRow = rulePalette.add("group");
        setupRow(presetRow, "center", FIELD_SPACING);

        presetRow.add("statictext", undefined, labelText("fieldLabel.preset"));
        ui.presetDropdown = presetRow.add("dropdownlist");
        ui.presetDropdown.preferredSize.width = DROPDOWN_WIDTH;
        ui.presetDropdown.helpTip = getLabel("tooltip.preset");
        ui.btnSavePreset = presetRow.add("button", undefined, getLabel("button.savePreset"));
        ui.btnSavePreset.helpTip = getLabel("tooltip.savePreset");
        ui.btnDeletePreset = presetRow.add("button", undefined, getLabel("button.deletePreset"));
        ui.btnDeletePreset.helpTip = getLabel("tooltip.deletePreset");

        ui.presetDropdown.onChange = function () {
            if (paletteState.isLoading) return;
            refreshPresetButtons(paletteState);
            var selectedEntry = getSelectedPresetEntry(ui.presetDropdown);
            if (!selectedEntry) return;
            /* 上書き保存の対象は、最後に選んだ保存済みのプリセット / the last chosen saved preset is the one to overwrite */
            paletteState.activeUserPreset = selectedEntry.isUser ? selectedEntry.preset : null;
            applyPreset(paletteState, selectedEntry.preset);
        };
        ui.btnSavePreset.onClick = function () { saveCurrentAsPreset(paletteState); };
        ui.btnDeletePreset.onClick = function () { deleteSelectedPreset(paletteState); };

        fillPresetDropdown(paletteState, null);
    }

    /**
     * プリセットの選択を外す（各項目を手で変えたあと、プリセットどおりではなくなったことを示す）。
     * 上書き保存の対象（activeUserPreset）はそのまま残す
     * @param {object} paletteState 状態オブジェクト
     * @returns {void}
     */
    function clearPresetSelection(paletteState) {
        var presetDropdown = paletteState.ui.presetDropdown;
        if (!presetDropdown.selection) return;
        paletteState.isLoading = true;
        try {
            presetDropdown.selection = null;
        } finally {
            paletteState.isLoading = false;
        }
        refreshPresetButtons(paletteState);
    }

    // -----------------------------------------
    // プリセットの保存 / Saving presets
    // -----------------------------------------

    /* 保存したプリセットのファイル（toSource() の形で書き出す）/ File holding the saved presets, written with toSource() */
    var PRESETS_FILE_NAME = "IdParagraphRulesPalette-presets.txt";

    /**
     * 保存したプリセットのファイルを返す
     * @returns {File} プリセットのファイル
     */
    function getPresetsFile() {
        return File(Folder.userData.fsName + "/" + PRESETS_FILE_NAME);
    }

    /**
     * 保存したプリセットを読み込む
     * @returns {Array<object>} 保存したプリセット。無い・読めなければ空の配列
     */
    function loadUserPresets() {
        var presetsFile = getPresetsFile();
        if (!presetsFile.exists) return [];
        presetsFile.encoding = "UTF-8";
        if (!presetsFile.open("r")) return [];
        try {
            var loadedPresets = eval(presetsFile.read());
            return (loadedPresets instanceof Array) ? loadedPresets : [];
        } catch (e) {
            return [];
        } finally {
            presetsFile.close();
        }
    }

    /**
     * 保存したプリセットをファイルに書き出す
     * @param {Array<object>} userPresets 保存したプリセット
     * @returns {boolean} 書き出せたら true
     */
    function saveUserPresets(userPresets) {
        var presetsFile = getPresetsFile();
        presetsFile.encoding = "UTF-8";
        presetsFile.lineFeed = "Unix";
        if (!presetsFile.open("w")) return false;
        try {
            return presetsFile.write(userPresets.toSource());
        } finally {
            presetsFile.close();
        }
    }

    /**
     * 長さの欄の文字列を、プリセットに残す pt の数値にする
     * @param {string} fieldText 入力欄の文字列（例 "0.25 mm"）
     * @param {UnitInfo} unitInfo 入力欄の単位
     * @returns {number} pt の値
     */
    function toPresetPoints(fieldText, unitInfo) {
        return Math.round(parseFloat(fieldText) * unitInfo.pointsPerUnit * 10000) / 10000;
    }

    /**
     * カラーのキーを、プリセットに残す値にする（特別なスウォッチは種別で残し、表示言語が変わっても引けるようにする）
     * @param {Array<ColorChoice>} colorChoices カラーの候補
     * @param {string} colorKey カラーのキー
     * @returns {string} "textColor" などの種別、またはスウォッチ名
     */
    function toPresetColor(colorChoices, colorKey) {
        var choiceIndex = findChoiceIndex(colorChoices, colorKey);
        if (choiceIndex >= 0 && colorChoices[choiceIndex].swatchKind) return colorChoices[choiceIndex].swatchKind;
        return colorKey;
    }

    /**
     * 境界線1本分の設定を、プリセットの形にする（長さは pt）
     * @param {object} paletteState 状態オブジェクト
     * @param {RuleModel} ruleModel 境界線の設定
     * @param {string} ruleSide "Above" または "Below"
     * @returns {object} プリセットの Above または Below
     */
    function toPresetRule(paletteState, ruleModel, ruleSide) {
        var fieldUnits = paletteState.fieldUnits;
        var presetRule = {
            weight: toPresetPoints(ruleModel.weightText, fieldUnits.stroke),
            type: isSolidStroke(ruleModel.typeKey) ? "solid" : ruleModel.typeKey,
            color: toPresetColor(paletteState.colorChoices, ruleModel.colorKey),
            tint: parseFloat(ruleModel.tintText),
            overprint: ruleModel.overprint,
            gapColor: toPresetColor(paletteState.gapColorChoices, ruleModel.gapColorKey),
            gapTint: parseFloat(ruleModel.gapTintText),
            gapOverprint: ruleModel.gapOverprint,
            width: (ruleModel.widthKey === PRESET_WIDTH_KEYS.text) ? "text" : "column",
            offset: toPresetPoints(ruleModel.offsetText, fieldUnits.vertical),
            leftIndent: toPresetPoints(ruleModel.leftIndentText, fieldUnits.horizontal),
            rightIndent: toPresetPoints(ruleModel.rightIndentText, fieldUnits.horizontal)
        };
        if (ruleSide === "Above") presetRule.keepInFrame = ruleModel.keepInFrame;
        return presetRule;
    }

    /**
     * 今の設定をプリセットに書き込む（［境界線を挿入］がオフの境界線は書かない）
     * @param {object} paletteState 状態オブジェクト
     * @param {object} preset 書き込む先のプリセット
     * @returns {void}
     */
    function writeModelsToPreset(paletteState, preset) {
        for (var i = 0; i < RULE_SIDES.length; i++) {
            var ruleSide = RULE_SIDES[i];
            var ruleModel = paletteState.models[ruleSide];
            if (ruleModel.ruleOn) preset[ruleSide] = toPresetRule(paletteState, ruleModel, ruleSide);
            else delete preset[ruleSide];
        }
    }

    /**
     * 保存したプリセットを名前で探す
     * @param {Array<object>} userPresets 保存したプリセット
     * @param {string} presetName 名前
     * @returns {object|null} 見つかったプリセット。無ければ null
     */
    function findUserPreset(userPresets, presetName) {
        for (var i = 0; i < userPresets.length; i++) {
            if (userPresets[i].name === presetName) return userPresets[i];
        }
        return null;
    }

    /**
     * 保存の方法と名前を聞くダイアログボックスを表示する
     * @param {string} overwriteName 上書きできるプリセットの名前（無ければ空文字。そのときは［上書き保存］を選べない）
     * @returns {{overwrite: boolean, name: string}|null} 選んだ方法と名前。キャンセルなら null
     */
    function showSavePresetDialog(overwriteName) {
        var saveDialog = new Window("dialog", getLabel("presetDialog.title"));
        setupWindow(saveDialog);

        /* ラジオボタンは同じ親の中だけで排他になるので、2つを1つの group に入れる / radio buttons are exclusive only within one parent */
        var modeGroup = saveDialog.add("group");
        modeGroup.orientation = "column";
        modeGroup.alignChildren = ["left", "center"];
        modeGroup.spacing = ROW_SPACING;
        var overwriteRadio = modeGroup.add("radiobutton", undefined,
            overwriteName ? getLabel("presetDialog.overwriteNamed", [overwriteName]) : getLabel("presetDialog.overwrite"));
        var saveAsRadio = modeGroup.add("radiobutton", undefined, getLabel("presetDialog.saveAs"));

        var nameRow = saveDialog.add("group");
        setupRow(nameRow, "left", FIELD_SPACING);
        nameRow.margins = [PRESET_NAME_INDENT, 0, 0, 0];
        var nameLabel = nameRow.add("statictext", undefined, labelText("fieldLabel.presetName"));
        var nameInput = nameRow.add("edittext", undefined, overwriteName);
        nameInput.preferredSize.width = DROPDOWN_WIDTH;

        /* 表示前に代入した value は読み戻せないので、選んだ方法は変数で持つ / a value assigned before show() cannot be read back */
        var isOverwrite = !!overwriteName;

        /**
         * 選んだ方法に合わせて、ラジオボタンと名前の欄を切り替える
         * @returns {void}
         */
        function refreshMode() {
            overwriteRadio.value = isOverwrite;
            saveAsRadio.value = !isOverwrite;
            nameLabel.enabled = !isOverwrite;
            nameInput.enabled = !isOverwrite;
        }
        overwriteRadio.enabled = !!overwriteName;
        overwriteRadio.onClick = function () { isOverwrite = true; refreshMode(); };
        saveAsRadio.onClick = function () {
            isOverwrite = false;
            refreshMode();
            nameInput.active = true;
        };
        refreshMode();

        var buttonRow = addButtonRow(saveDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnSave = buttonRow.rightGroup.add("button", undefined, getLabel("button.save"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        btnSave.onClick = function () {
            if (!isOverwrite && trimText(nameInput.text) === "") {
                alert(getLabel("alert.emptyPresetName"));
                nameInput.active = true;
                return;
            }
            saveDialog.close(1);
        };
        if (!isOverwrite) saveDialog.onShow = function () { nameInput.active = true; };

        if (saveDialog.show() !== 1) return null;
        return { overwrite: isOverwrite, name: isOverwrite ? overwriteName : trimText(nameInput.text) };
    }

    /**
     * 文字列の前後の空白を取る
     * @param {string} text 文字列
     * @returns {string} 前後の空白を取った文字列
     */
    function trimText(text) {
        return String(text).replace(/^\s+|\s+$/g, "");
    }

    /**
     * 今の設定をプリセットとして保存する（上書き保存、または名前を付けて保存）
     * @param {object} paletteState 状態オブジェクト
     * @returns {void}
     */
    function saveCurrentAsPreset(paletteState) {
        var userPresets = paletteState.userPresets;
        var activePreset = paletteState.activeUserPreset;
        var saveAnswer = showSavePresetDialog(activePreset ? activePreset.name : "");
        if (!saveAnswer) return;

        var targetPreset = saveAnswer.overwrite ? activePreset : findUserPreset(userPresets, saveAnswer.name);
        if (!saveAnswer.overwrite && targetPreset && !confirm(getLabel("confirm.replacePreset", [saveAnswer.name]), true)) return;
        if (!targetPreset) {
            targetPreset = { name: saveAnswer.name };
            userPresets.push(targetPreset);
        }
        writeModelsToPreset(paletteState, targetPreset);
        if (!saveUserPresets(userPresets)) alert(getLabel("alert.savePresetFailed", [getPresetsFile().fsName]));

        paletteState.activeUserPreset = targetPreset;
        fillPresetDropdown(paletteState, targetPreset);
    }

    /**
     * 選んでいる保存済みのプリセットを削除する
     * @param {object} paletteState 状態オブジェクト
     * @returns {void}
     */
    function deleteSelectedPreset(paletteState) {
        var selectedEntry = getSelectedPresetEntry(paletteState.ui.presetDropdown);
        if (!selectedEntry || !selectedEntry.isUser) return;
        var targetPreset = selectedEntry.preset;
        if (!confirm(getLabel("confirm.deletePreset", [targetPreset.name]), true)) return;

        var userPresets = paletteState.userPresets;
        for (var i = userPresets.length - 1; i >= 0; i--) {
            if (userPresets[i] === targetPreset) userPresets.splice(i, 1);
        }
        if (paletteState.activeUserPreset === targetPreset) paletteState.activeUserPreset = null;
        if (!saveUserPresets(userPresets)) alert(getLabel("alert.savePresetFailed", [getPresetsFile().fsName]));
        fillPresetDropdown(paletteState, null);
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューとして書き込んだ設定を取り消す。
     * 直前の取り消しがプレビューのときだけ戻す（ドキュメントが閉じられたり、別の操作が挟まったりしたら戻さない）
     * @param {object} paletteState 状態オブジェクト
     * @returns {void}
     */
    function clearPreview(paletteState) {
        if (!paletteState.previewed) return;
        paletteState.previewed = false;
        var targetDoc = paletteState.targetDoc;
        if (targetDoc.isValid && targetDoc.undoName === getLabel("undo.preview")) targetDoc.undo();
    }

    /**
     * 現在の設定でプレビューを描き直す（プレビューがオフなら取り消すだけ）
     * @param {object} paletteState 状態オブジェクト
     * @returns {void}
     */
    function refreshPreview(paletteState) {
        clearPreview(paletteState);
        if (!paletteState.ui.previewCheckbox.value) return;

        var changedProperties = buildChangedProperties(paletteState);
        /* 変更が無いと取り消しの段が増えず、次の undo がユーザーの操作を戻してしまう
           / with no change there is no undo step, and the next undo would revert the user's own work */
        if (isEmptyProperties(changedProperties)) return;

        /* パレット表示中に対象のテキストが消されると書き込めない / the target text may be deleted while the palette is open */
        try {
            applyRuleProperties(paletteState.textTargets, changedProperties, getLabel("undo.preview"));
        } catch (e) {
            paletteState.ui.previewCheckbox.value = false;
            alert(getLabel("alert.previewFailed", [e.message]));
            return;
        }
        paletteState.previewed = true;
    }

    // =========================================
    // パレット / Palette
    // =========================================

    /**
     * 左右2列の1行を作る
     * @param {Group|Panel} parent 追加先
     * @returns {{left: Group, right: Group}} 左の列と右の列
     */
    function addTwoColumnRow(parent) {
        var twoColumnRow = parent.add("group");
        setupRow(twoColumnRow, "fill", 0);

        var leftCell = twoColumnRow.add("group");
        setupRow(leftCell, "left", FIELD_SPACING);
        leftCell.preferredSize.width = LEFT_CELL_WIDTH;

        var rightCell = twoColumnRow.add("group");
        setupRow(rightCell, "left", FIELD_SPACING);
        return { left: leftCell, right: rightCell };
    }

    /**
     * 右揃えの項目名を追加する
     * @param {Group} parent 追加先
     * @param {string} labelKey 項目名のラベルキー
     * @param {number} labelWidth 項目名の幅
     * @returns {StaticText} 項目名
     */
    function addFieldLabel(parent, labelKey, labelWidth) {
        var fieldLabel = parent.add("statictext", undefined, labelText(labelKey));
        fieldLabel.preferredSize.width = labelWidth;
        fieldLabel.justify = "right";
        return fieldLabel;
    }

    /**
     * 項目名つきのドロップダウンを追加する
     * @param {Group} parent 追加先
     * @param {string} labelKey 項目名のラベルキー
     * @param {number} labelWidth 項目名の幅
     * @param {Array<{key: string, displayName: string}>} choices 候補
     * @param {boolean} withSwatchChip 左に色見本を置くなら true
     * @returns {DropDownList} ドロップダウン（.choices・.fieldLabel・.swatchChip を持つ）
     */
    function addChoiceDropdown(parent, labelKey, labelWidth, choices, withSwatchChip) {
        var fieldLabel = addFieldLabel(parent, labelKey, labelWidth);

        var controlGroup = parent.add("group");
        setupRow(controlGroup, "left", SWATCH_CHIP_SPACING);
        var swatchChip = withSwatchChip ? addSwatchChip(controlGroup) : null;

        var displayNames = [];
        for (var i = 0; i < choices.length; i++) {
            displayNames.push(choices[i].displayName);
        }
        var choiceDropdown = controlGroup.add("dropdownlist", undefined, displayNames);
        choiceDropdown.preferredSize.width = withSwatchChip ? DROPDOWN_WIDTH - SWATCH_CHIP_SIZE - SWATCH_CHIP_SPACING : DROPDOWN_WIDTH;
        choiceDropdown.choices = choices;
        choiceDropdown.fieldLabel = fieldLabel;
        choiceDropdown.swatchChip = swatchChip;
        return choiceDropdown;
    }

    /**
     * 項目名・∧∨つきの数値欄を追加する
     * @param {Group} parent 追加先
     * @param {string} labelKey 項目名のラベルキー
     * @param {number} labelWidth 項目名の幅
     * @param {object} stepOptions step / min / max / unit
     * @param {function(EditText): void} onValueChange 値が変わったときに呼ぶ処理
     * @returns {EditText} 入力欄
     */
    function addNumberField(parent, labelKey, labelWidth, stepOptions, onValueChange) {
        stepOptions.label = labelText(labelKey);
        stepOptions.labelWidth = labelWidth;
        stepOptions.onStep = onValueChange;
        var numberInput = addSteppedField(parent, stepOptions);
        numberInput.preferredSize.width = NUMBER_FIELD_WIDTH;
        numberInput.fieldLabel.justify = "right";

        /* 部品の onChange（値の整え）のあとに、設定への書き戻しを続ける / chain the model update after the part's own normalization */
        var normalizeOnChange = numberInput.onChange;
        numberInput.onChange = function () {
            normalizeOnChange();
            onValueChange(numberInput);
        };
        return numberInput;
    }

    /**
     * 項目名つきの横並びのラジオボタンを追加する
     * @param {Group} parent 追加先
     * @param {string} labelKey 項目名のラベルキー
     * @param {number} labelWidth 項目名の幅
     * @param {Array<{key: string, displayName: string}>} choices 候補
     * @returns {Group} ラジオボタンを並べた group（.radioButtons・.choices を持つ）
     */
    function addChoiceRadioGroup(parent, labelKey, labelWidth, choices) {
        addFieldLabel(parent, labelKey, labelWidth);

        /* ラジオボタンは同じ親の中だけで排他になるので、1つの group にまとめる / radio buttons are exclusive only within one parent */
        var radioGroup = parent.add("group");
        setupRow(radioGroup, "left", FIELD_SPACING * 2);
        radioGroup.margins = [0, 0, 0, RADIO_BASELINE_MARGIN];
        radioGroup.radioButtons = [];
        radioGroup.choices = choices;
        for (var i = 0; i < choices.length; i++) {
            radioGroup.radioButtons.push(radioGroup.add("radiobutton", undefined, choices[i].displayName));
        }
        return radioGroup;
    }

    /**
     * 入力の列にそろえたチェックボックスを追加する
     * @param {Group|Panel} parent 追加先
     * @param {string} labelKey ラベルキー
     * @returns {Checkbox} チェックボックス
     */
    function addIndentedCheckbox(parent, labelKey) {
        var checkboxRow = parent.add("group");
        setupRow(checkboxRow, "left", 0);
        checkboxRow.margins = [CHECKBOX_INDENT, 0, 0, 0];
        return checkboxRow.add("checkbox", undefined, getLabel(labelKey));
    }

    /**
     * 前境界線または後境界線のパネルを作る
     * @param {Window} rulePalette 追加先のパレット
     * @param {object} paletteState 状態オブジェクト
     * @param {string} ruleSide "Above" または "Below"
     * @returns {object} パネルのコントロール一式
     */
    function buildRuleSection(rulePalette, paletteState, ruleSide) {
        var fieldUnits = paletteState.fieldUnits;
        var ruleSection = { ruleSide: ruleSide };

        var rulePanel = rulePalette.add("panel", undefined, getLabel(ruleSide === "Above" ? "panel.ruleAbove" : "panel.ruleBelow"));
        setupPanel(rulePanel, ROW_SPACING);

        ruleSection.ruleOnCheckbox = rulePanel.add("checkbox", undefined, getLabel("checkbox.ruleOn"));

        /* ［境界線を挿入］で有効／無効を切り替える部分 / The part enabled by Rule On */
        var detailGroup = rulePanel.add("group");
        detailGroup.orientation = "column";
        detailGroup.alignChildren = ["fill", "top"];
        detailGroup.spacing = ROW_SPACING;
        ruleSection.detailGroup = detailGroup;

        var onFieldChange = function (numberInput) { handleFieldChange(paletteState, numberInput); };

        var weightRow = addTwoColumnRow(detailGroup);
        ruleSection.weightInput = addNumberField(weightRow.left, "fieldLabel.weight", LEFT_LABEL_WIDTH,
            { step: fieldUnits.stroke.weightStep, min: 0, unit: fieldUnits.stroke.unit }, onFieldChange);
        ruleSection.typeDropdown = addChoiceDropdown(weightRow.right, "fieldLabel.type", RIGHT_LABEL_WIDTH, paletteState.strokeTypeChoices, false);
        ruleSection.typeDropdown.preferredSize.width = TYPE_DROPDOWN_WIDTH;

        var colorRow = addTwoColumnRow(detailGroup);
        ruleSection.colorDropdown = addChoiceDropdown(colorRow.left, "fieldLabel.color", LEFT_LABEL_WIDTH, paletteState.colorChoices, true);
        ruleSection.tintInput = addNumberField(colorRow.right, "fieldLabel.tint", RIGHT_LABEL_WIDTH,
            { step: TINT_STEP, min: TINT_MIN, max: TINT_MAX, unit: "%" }, onFieldChange);
        ruleSection.overprintCheckbox = addIndentedCheckbox(detailGroup, "checkbox.overprint");

        var gapRow = addTwoColumnRow(detailGroup);
        ruleSection.gapColorDropdown = addChoiceDropdown(gapRow.left, "fieldLabel.gapColor", LEFT_LABEL_WIDTH, paletteState.gapColorChoices, true);
        ruleSection.gapTintInput = addNumberField(gapRow.right, "fieldLabel.gapTint", RIGHT_LABEL_WIDTH,
            { step: TINT_STEP, min: TINT_MIN, max: TINT_MAX, unit: "%" }, onFieldChange);
        ruleSection.gapOverprintCheckbox = addIndentedCheckbox(detailGroup, "checkbox.gapOverprint");

        var widthRow = addTwoColumnRow(detailGroup);
        ruleSection.widthRadioGroup = addChoiceRadioGroup(widthRow.left, "fieldLabel.width", LEFT_LABEL_WIDTH, paletteState.ruleWidthChoices);
        ruleSection.offsetInput = addNumberField(widthRow.right, "fieldLabel.offset", RIGHT_LABEL_WIDTH,
            { step: fieldUnits.vertical.step, unit: fieldUnits.vertical.unit }, onFieldChange);
        ruleSection.offsetInput.helpTip = getLabel("tooltip.offset");

        var indentRow = addTwoColumnRow(detailGroup);
        /* 縦組みでは左右のインデントが上下になる / in vertical text the left and right indents become top and bottom */
        var startIndentKey = paletteState.isVerticalStory ? "fieldLabel.topIndent" : "fieldLabel.leftIndent";
        var endIndentKey = paletteState.isVerticalStory ? "fieldLabel.bottomIndent" : "fieldLabel.rightIndent";
        ruleSection.leftIndentInput = addNumberField(indentRow.left, startIndentKey, LEFT_LABEL_WIDTH,
            { step: fieldUnits.horizontal.step, unit: fieldUnits.horizontal.unit }, onFieldChange);
        ruleSection.rightIndentInput = addNumberField(indentRow.right, endIndentKey, RIGHT_LABEL_WIDTH,
            { step: fieldUnits.horizontal.step, unit: fieldUnits.horizontal.unit }, onFieldChange);
        ruleSection.leftIndentInput.helpTip = getLabel("tooltip.indent");
        ruleSection.rightIndentInput.helpTip = getLabel("tooltip.indent");

        /* ［フレームに収める］は前境界線だけ / Keep In Frame exists only for the rule above */
        ruleSection.keepInFrameCheckbox = null;
        if (ruleSide === "Above") {
            var keepInFrameRow = detailGroup.add("group");
            setupRow(keepInFrameRow, "left", 0);
            keepInFrameRow.margins = [0, ROW_SPACING, 0, 0];
            ruleSection.keepInFrameCheckbox = keepInFrameRow.add("checkbox", undefined, getLabel("checkbox.keepInFrame"));
            ruleSection.keepInFrameCheckbox.helpTip = getLabel("tooltip.keepInFrame");
        }

        bindRuleSectionControls(paletteState, ruleSection);
        return ruleSection;
    }

    /**
     * パネルの各コントロールに、編集する設定の項目と境界線の向きを持たせ、ハンドラーを付ける
     * （数値欄のハンドラーは addNumberField で付けてある）
     * @param {object} paletteState 状態オブジェクト
     * @param {object} ruleSection パネルのコントロール一式
     * @returns {void}
     */
    function bindRuleSectionControls(paletteState, ruleSection) {
        var controlsByField = {
            ruleOn: ruleSection.ruleOnCheckbox,
            weightText: ruleSection.weightInput,
            typeKey: ruleSection.typeDropdown,
            colorKey: ruleSection.colorDropdown,
            tintText: ruleSection.tintInput,
            overprint: ruleSection.overprintCheckbox,
            gapColorKey: ruleSection.gapColorDropdown,
            gapTintText: ruleSection.gapTintInput,
            gapOverprint: ruleSection.gapOverprintCheckbox,
            widthKey: ruleSection.widthRadioGroup,
            offsetText: ruleSection.offsetInput,
            leftIndentText: ruleSection.leftIndentInput,
            rightIndentText: ruleSection.rightIndentInput,
            keepInFrame: ruleSection.keepInFrameCheckbox
        };
        var onControlChange = function () { handleFieldChange(paletteState, this); };
        for (var modelField in controlsByField) {
            if (!controlsByField.hasOwnProperty(modelField) || !controlsByField[modelField]) continue;
            var fieldControl = controlsByField[modelField];
            fieldControl.modelField = modelField;
            fieldControl.ruleSide = ruleSection.ruleSide;
            if (fieldControl.type === "dropdownlist") fieldControl.onChange = onControlChange;
            else if (fieldControl.type === "checkbox") fieldControl.onClick = onControlChange;
            else if (fieldControl.radioButtons) bindRadioGroup(paletteState, fieldControl);
        }
    }

    /**
     * ラジオボタンのクリックを、group 単位の変更として受ける
     * @param {object} paletteState 状態オブジェクト
     * @param {Group} radioGroup addChoiceRadioGroup() で作った group
     * @returns {void}
     */
    function bindRadioGroup(paletteState, radioGroup) {
        var onRadioClick = function () { handleFieldChange(paletteState, radioGroup); };
        for (var i = 0; i < radioGroup.radioButtons.length; i++) {
            radioGroup.radioButtons[i].onClick = onRadioClick;
        }
    }

    /**
     * 段落のインデントのパネル（左・右の欄と、左右を連動させるリンクアイコン）を作る
     * @param {Window} rulePalette 追加先のパレット
     * @param {object} paletteState 状態オブジェクト
     * @returns {void}
     */
    function buildIndentSection(rulePalette, paletteState) {
        var ui = paletteState.ui;
        var horizontalUnit = paletteState.fieldUnits.horizontal;
        var indentModel = paletteState.indentModel;

        var indentPanel = rulePalette.add("panel", undefined, getLabel("panel.indent"));
        setupPanel(indentPanel, ROW_SPACING);

        var onIndentChange = function (numberInput) { handleIndentChange(paletteState, numberInput); };
        var indentRow = addTwoColumnRow(indentPanel);
        /* 縦組みでは左右のインデントが上下になる / in vertical text the left and right indents become top and bottom */
        var startIndentKey = paletteState.isVerticalStory ? "fieldLabel.topIndent" : "fieldLabel.leftIndent";
        var endIndentKey = paletteState.isVerticalStory ? "fieldLabel.bottomIndent" : "fieldLabel.rightIndent";
        ui.paragraphLeftIndentInput = addNumberField(indentRow.left, startIndentKey, LEFT_LABEL_WIDTH,
            { step: horizontalUnit.step, min: 0, unit: horizontalUnit.unit, text: indentModel.leftIndentText }, onIndentChange);
        ui.paragraphRightIndentInput = addNumberField(indentRow.right, endIndentKey, RIGHT_LABEL_WIDTH,
            { step: horizontalUnit.step, min: 0, unit: horizontalUnit.unit, text: indentModel.rightIndentText }, onIndentChange);
        ui.paragraphLeftIndentInput.modelField = "leftIndentText";
        ui.paragraphRightIndentInput.modelField = "rightIndentText";
        ui.paragraphLeftIndentInput.pairedInput = ui.paragraphRightIndentInput;
        ui.paragraphRightIndentInput.pairedInput = ui.paragraphLeftIndentInput;

        /* 開いたときに左右が同じなら連動から始める。連動にしたら右を左にそろえる
           / start linked when both indents match; linking copies the left value to the right */
        ui.indentLinkToggle = addLinkToggle(indentRow.right,
            indentModel.leftIndentText === indentModel.rightIndentText,
            function () {
                if (ui.indentLinkToggle.value) handleIndentChange(paletteState, ui.paragraphLeftIndentInput);
            });
        ui.indentLinkToggle.helpTip = getLabel("tooltip.linkIndents");
    }

    /**
     * 段落のインデントの欄の値を設定へ書き戻し、連動中ならもう一方の欄もそろえてプレビューを更新する
     * @param {object} paletteState 状態オブジェクト
     * @param {EditText} numberInput 値が変わった欄（.modelField と .pairedInput を持つ）
     * @returns {void}
     */
    function handleIndentChange(paletteState, numberInput) {
        var indentModel = paletteState.indentModel;
        indentModel[numberInput.modelField] = numberInput.text;
        if (paletteState.ui.indentLinkToggle && paletteState.ui.indentLinkToggle.value) {
            var pairedInput = numberInput.pairedInput;
            writeFieldText(pairedInput, numberInput.text);
            indentModel[pairedInput.modelField] = numberInput.text;
        }
        refreshPreview(paletteState);
    }

    /* パレットの参照を残す $.global のキー（再実行で前回のパレットを閉じるため）
       / $.global key holding the palette, so a rerun can close the previous one */
    var PALETTE_GLOBAL_KEY = "__IdParagraphRulesPalette";

    /**
     * パレットを組み立てる（上からプリセット・前境界線・後境界線・インデント、いちばん下にプレビューとボタン）
     * @param {object} paletteState 状態オブジェクト
     * @returns {object} パレットとコントロールをまとめた UI オブジェクト
     */
    function buildPalette(paletteState) {
        var rulePalette = new Window("palette", getLabel("palette.title") + " " + SCRIPT_VERSION);
        setupWindow(rulePalette);

        var ui = { palette: rulePalette, sections: {} };
        paletteState.ui = ui;
        addPresetRow(rulePalette, paletteState);
        for (var i = 0; i < RULE_SIDES.length; i++) {
            ui.sections[RULE_SIDES[i]] = buildRuleSection(rulePalette, paletteState, RULE_SIDES[i]);
        }
        buildIndentSection(rulePalette, paletteState);

        /* ボタン行：左にプレビュー、右にキャンセル・OK / Button row: Preview on the left, Cancel and OK on the right */
        var buttonRow = addButtonRow(rulePalette);
        ui.previewCheckbox = buttonRow.leftGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        ui.previewCheckbox.value = true;
        ui.previewCheckbox.helpTip = getLabel("tooltip.preview");
        ui.btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"));
        ui.btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"));
        alignRightOnlyButtonRow(buttonRow);

        bindPaletteEvents(paletteState);
        return ui;
    }

    /**
     * プレビュー・ボタン・Esc・閉じたときの後片付けを付ける
     * @param {object} paletteState 状態オブジェクト
     * @returns {void}
     */
    function bindPaletteEvents(paletteState) {
        var ui = paletteState.ui;
        var rulePalette = ui.palette;

        ui.previewCheckbox.onClick = function () { refreshPreview(paletteState); };

        /* ［OK］はプレビューを戻してから1回の取り消しで書き込む / OK undoes the preview, then writes everything as one undo step */
        ui.btnOK.onClick = function () {
            clearPreview(paletteState);
            commitRuleChanges(paletteState);
            rulePalette.close();
        };
        ui.btnCancel.onClick = function () { rulePalette.close(); };

        /* パレットは Esc で閉じないので、［キャンセル］と同じにする / palettes do not close on Esc, so treat it as Cancel */
        rulePalette.addEventListener("keydown", function (event) {
            if (event.keyName !== "Escape") return;
            event.preventDefault();
            rulePalette.close();
        });

        /* ［OK］以外で閉じたらプレビューを戻す（［OK］のあとは戻すものが無い） / Closing without OK reverts the preview */
        rulePalette.onClose = function () {
            clearPreview(paletteState);
            $.global[PALETTE_GLOBAL_KEY] = null;
        };
    }

    /**
     * 変わった項目を対象のテキストに書き込む
     * @param {object} paletteState 状態オブジェクト
     * @returns {void}
     */
    function commitRuleChanges(paletteState) {
        var changedProperties = buildChangedProperties(paletteState);
        if (isEmptyProperties(changedProperties)) return;
        /* パレット表示中に対象のテキストが消されると書き込めない / the target text may be deleted while the palette is open */
        try {
            applyRuleProperties(paletteState.textTargets, changedProperties, getLabel("undo.apply"));
        } catch (e) {
            alert(getLabel("alert.applyFailed", [e.message]));
        }
    }

    /**
     * コントロールの値を、その境界線の設定へ書き戻して有効／無効とプレビューを更新する
     * @param {object} paletteState 状態オブジェクト
     * @param {Control} changedControl 値が変わったコントロール（.modelField と .ruleSide を持つ）
     * @returns {void}
     */
    function handleFieldChange(paletteState, changedControl) {
        if (paletteState.isLoading || !changedControl.modelField) return;
        var ruleModel = paletteState.models[changedControl.ruleSide];

        if (changedControl.type === "dropdownlist") {
            if (!changedControl.selection) return;
            ruleModel[changedControl.modelField] = changedControl.choices[changedControl.selection.index].key;
        } else if (changedControl.type === "checkbox") {
            ruleModel[changedControl.modelField] = changedControl.value;
        } else if (changedControl.radioButtons) {
            var checkedIndex = getCheckedRadioIndex(changedControl);
            if (checkedIndex < 0) return;
            ruleModel[changedControl.modelField] = changedControl.choices[checkedIndex].key;
        } else {
            ruleModel[changedControl.modelField] = changedControl.text;
        }

        clearPresetSelection(paletteState);
        refreshControlStates(paletteState, changedControl.ruleSide);
        refreshPreview(paletteState);
    }

    /**
     * ドロップダウンを候補のキーで選ぶ。候補に無ければ未選択にする（複数の値が混在する表示を兼ねる）
     * @param {DropDownList} choiceDropdown 対象のドロップダウン
     * @param {string} key 選ぶキー
     * @returns {void}
     */
    function selectChoice(choiceDropdown, key) {
        var choiceIndex = findChoiceIndex(choiceDropdown.choices, key);
        choiceDropdown.selection = (choiceIndex >= 0) ? choiceIndex : null;
    }

    /**
     * ラジオボタンの group で選ばれている位置を返す
     * @param {Group} radioGroup addChoiceRadioGroup() で作った group
     * @returns {number} 選ばれている位置。無ければ -1
     */
    function getCheckedRadioIndex(radioGroup) {
        for (var i = 0; i < radioGroup.radioButtons.length; i++) {
            if (radioGroup.radioButtons[i].value) return i;
        }
        return -1;
    }

    /**
     * ラジオボタンを候補のキーで選ぶ。候補に無ければどれも選ばない
     * @param {Group} radioGroup addChoiceRadioGroup() で作った group
     * @param {string} key 選ぶキー
     * @returns {void}
     */
    function selectRadioChoice(radioGroup, key) {
        var choiceIndex = findChoiceIndex(radioGroup.choices, key);
        for (var i = 0; i < radioGroup.radioButtons.length; i++) {
            radioGroup.radioButtons[i].value = (i === choiceIndex);
        }
    }

    /**
     * 前境界線または後境界線の設定をパネルのコントロールに入れる
     * @param {object} paletteState 状態オブジェクト
     * @param {string} ruleSide "Above" または "Below"
     * @returns {void}
     */
    function loadRuleModel(paletteState, ruleSide) {
        var ruleSection = paletteState.ui.sections[ruleSide];
        var ruleModel = paletteState.models[ruleSide];

        /* selection への代入で onChange が呼ばれるので、その間は書き戻さない / assigning a selection fires onChange, so hold back */
        paletteState.isLoading = true;
        try {
            ruleSection.ruleOnCheckbox.value = ruleModel.ruleOn;
            writeFieldText(ruleSection.weightInput, ruleModel.weightText);
            selectChoice(ruleSection.typeDropdown, ruleModel.typeKey);
            selectChoice(ruleSection.colorDropdown, ruleModel.colorKey);
            writeFieldText(ruleSection.tintInput, ruleModel.tintText);
            ruleSection.overprintCheckbox.value = ruleModel.overprint;
            selectChoice(ruleSection.gapColorDropdown, ruleModel.gapColorKey);
            writeFieldText(ruleSection.gapTintInput, ruleModel.gapTintText);
            ruleSection.gapOverprintCheckbox.value = ruleModel.gapOverprint;
            selectRadioChoice(ruleSection.widthRadioGroup, ruleModel.widthKey);
            writeFieldText(ruleSection.offsetInput, ruleModel.offsetText);
            writeFieldText(ruleSection.leftIndentInput, ruleModel.leftIndentText);
            writeFieldText(ruleSection.rightIndentInput, ruleModel.rightIndentText);
            if (ruleSection.keepInFrameCheckbox) ruleSection.keepInFrameCheckbox.value = ruleModel.keepInFrame;
        } finally {
            paletteState.isLoading = false;
        }
        refreshControlStates(paletteState, ruleSide);
    }

    /**
     * 数値欄に文字列を入れ、直前の正しい値としても控える（数値でない入力を戻す先）
     * @param {EditText} numberInput 対象の入力欄
     * @param {string} fieldText 入れる文字列
     * @returns {void}
     */
    function writeFieldText(numberInput, fieldText) {
        numberInput.text = fieldText;
        numberInput.lastValidText = fieldText;
    }

    /**
     * 選択中のカラーの候補を返す
     * @param {DropDownList} colorDropdown カラーのドロップダウン
     * @returns {ColorChoice|null} 選択中の候補。未選択なら null
     */
    function getSelectedColorChoice(colorDropdown) {
        return colorDropdown.selection ? colorDropdown.choices[colorDropdown.selection.index] : null;
    }

    /**
     * 濃淡・オーバープリントを指定できるカラーかどうかを判定する（テキストカラー・なし・紙色は不可）
     * @param {ColorChoice|null} colorChoice カラーの候補
     * @returns {boolean} 指定できるなら true
     */
    function canSetTintAndOverprint(colorChoice) {
        if (!colorChoice) return true;
        var swatchKind = colorChoice.swatchKind;
        return swatchKind !== "textColor" && swatchKind !== "none" && swatchKind !== "paper";
    }

    /**
     * ドロップダウンと項目名・色見本の有効／無効をまとめて切り替える
     * @param {DropDownList} choiceDropdown 対象のドロップダウン
     * @param {boolean} isEnabled 有効にするなら true
     * @returns {void}
     */
    function setDropdownEnabled(choiceDropdown, isEnabled) {
        choiceDropdown.enabled = isEnabled;
        choiceDropdown.fieldLabel.enabled = isEnabled;
    }

    /**
     * 設定の組み合わせに合わせて、パネルのコントロールの有効／無効と色見本を更新する
     * （［境界線を挿入］がオフなら全体、実線なら間隔、テキストカラー・なし・紙色なら濃淡とオーバープリント）
     * @param {object} paletteState 状態オブジェクト
     * @param {string} ruleSide "Above" または "Below"
     * @returns {void}
     */
    function refreshControlStates(paletteState, ruleSide) {
        var ruleSection = paletteState.ui.sections[ruleSide];
        ruleSection.detailGroup.enabled = ruleSection.ruleOnCheckbox.value;

        var colorChoice = getSelectedColorChoice(ruleSection.colorDropdown);
        var colorAllowsTint = canSetTintAndOverprint(colorChoice);
        setSteppedFieldEnabled(ruleSection.tintInput, colorAllowsTint);
        ruleSection.overprintCheckbox.enabled = colorAllowsTint;

        var hasGap = !isSolidStroke(paletteState.models[ruleSide].typeKey);
        var gapColorChoice = getSelectedColorChoice(ruleSection.gapColorDropdown);
        var gapAllowsTint = hasGap && canSetTintAndOverprint(gapColorChoice);
        setDropdownEnabled(ruleSection.gapColorDropdown, hasGap);
        setSteppedFieldEnabled(ruleSection.gapTintInput, gapAllowsTint);
        ruleSection.gapOverprintCheckbox.enabled = gapAllowsTint;

        paintSwatchChip(ruleSection.colorDropdown.swatchChip, colorChoice, paletteState.textColorSwatch);
        paintSwatchChip(ruleSection.gapColorDropdown.swatchChip, gapColorChoice, null);
        /* 親の enabled を切り替えたので、自作描画の∧∨を描き直す / the parent's enabled changed, so redraw the custom-drawn steppers */
        redrawSteppersIn(ruleSection.detailGroup);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 段落のテキストカラーとして色見本に見せるスウォッチを返す
     * @param {Paragraph} paragraph 対象の段落
     * @returns {Swatch|null} 文字の塗りのスウォッチ
     */
    function getTextColorSwatch(paragraph) {
        var fillColor = paragraph.fillColor;
        return (fillColor && fillColor.isValid) ? fillColor : null;
    }

    /**
     * テキストが縦組みのストーリーにあるかどうかを判定する
     * @param {Text} textTarget 対象のテキスト
     * @returns {boolean} 縦組みなら true
     */
    function isVerticalStory(textTarget) {
        return textTarget.parentStory.storyPreferences.storyOrientation === StoryHorizontalOrVertical.VERTICAL;
    }

    /**
     * 選択したテキストの段落境界線をパレットで設定する
     * @returns {void}
     */
    function main() {
        /* 前回のパレットが開いていれば閉じる（プレビューは戻る） / Close a palette left open (its preview is reverted) */
        if ($.global[PALETTE_GLOBAL_KEY]) $.global[PALETTE_GLOBAL_KEY].close();

        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var textTargets = collectTextTargets(app.selection);
        if (textTargets.length === 0) {
            alert(getLabel("alert.selectText"));
            return;
        }

        var targetDoc = app.activeDocument;
        var viewPreferences = targetDoc.viewPreferences;
        var fieldUnits = {
            stroke: getUnitInfo(viewPreferences.strokeMeasurementUnits),
            horizontal: getUnitInfo(viewPreferences.horizontalMeasurementUnits),
            vertical: getUnitInfo(viewPreferences.verticalMeasurementUnits)
        };

        /* 表示する値は先頭の段落から読む。空のフレームやセルでも段落を取れるよう挿入点から引く
           / Values shown come from the first paragraph, reached through an insertion point so empty frames and cells work too */
        var firstParagraph = textTargets[0].insertionPoints[0].paragraphs[0];
        var paragraphSettings = readParagraphSettings(firstParagraph, fieldUnits);
        var models = paragraphSettings.models;
        var indentModel = paragraphSettings.indentModel;

        var paletteState = {
            emPoints: paragraphSettings.pointSize,
            userPresets: loadUserPresets(),
            activeUserPreset: null,
            textTargets: textTargets,
            fieldUnits: fieldUnits,
            models: models,
            initialModels: { Above: cloneRuleModel(models.Above), Below: cloneRuleModel(models.Below) },
            indentModel: indentModel,
            initialIndentModel: cloneRuleModel(indentModel),
            colorChoices: buildColorChoices(targetDoc, true),
            gapColorChoices: buildColorChoices(targetDoc, false),
            strokeTypeChoices: buildStrokeTypeChoices(targetDoc),
            ruleWidthChoices: buildRuleWidthChoices(),
            textColorSwatch: getTextColorSwatch(firstParagraph),
            targetDoc: targetDoc,
            isVerticalStory: isVerticalStory(textTargets[0]),
            isLoading: false,
            previewed: false,
            ui: null
        };

        var ui = buildPalette(paletteState);
        for (var i = 0; i < RULE_SIDES.length; i++) {
            loadRuleModel(paletteState, RULE_SIDES[i]);
        }

        /* パレットは show() から即座に戻るので、参照を $.global に残す / palettes return from show() at once, so keep a reference */
        $.global[PALETTE_GLOBAL_KEY] = ui.palette;
        ui.palette.show();
    }

    main();

})();
