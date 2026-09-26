#target indesign
#targetengine "IdSetImageScale"

/*

### 概要

選択したフレーム内の画像の縮尺率を、パレットで数値・プリセット・マージン幅・親フレームの段幅から指定して変更し、フレームを画像に合わせます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSetImageScale.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n91c6a628b7ed

### Overview

Changes the scale of the image in each selected frame, set from a palette by value, preset, margin width or parent column width, and fits the frame to the image.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSetImageScale.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdSetImageScale";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-26";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSetImageScale.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSetImageScale.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n91c6a628b7ed"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 基本設定 / Settings
    // =========================================
    var DEFAULT_SCALE_PERCENT = 10;                        /* 縮尺率の初期値（現在値が混在のとき）（%） / default scale when current values differ (%) */
    var SCALE_STEPS   = [-10, 10, -5, 5, -1, 1];           /* 増減ボタンの量（2列で −／＋ の対に並ぶ順） / step buttons, in −/+ pairs */
    var SCALE_PRESETS = [10, 15, 20, 25, 30, 35, 40, 45, 50, 100, 200, 300, 400, 500];   /* プリセットの縮尺率（%） / preset scales (%) */

    /* 幅に合わせるときの基準 / Width references for fitting */
    var FIT_MODE = {
        margin:    "margin",      /* ページのマージン幅 / page margin width */
        textFrame: "textFrame"    /* 親テキストフレームの段幅 / column width of the parent text frame */
    };

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PALETTE_MARGINS        = 16;   /* パレット外周の余白 / palette margin */
    var PALETTE_SPACING        = 12;   /* パレット内の要素間隔 / palette spacing */
    var SCALE_FIELD_CHARS      = 5;    /* 縮尺率欄の文字数 / scale field width in characters */
    var ROW_LABEL_WIDTH        = 110;  /* 行ラベルの幅 / row label width */
    var COMPACT_BUTTON_WIDTH   = 44;   /* 小さいボタンの幅（増減・プリセット） / compact button width */
    var COMPACT_BUTTON_HEIGHT  = 20;   /* 小さいボタンの高さ / compact button height */
    var COMPACT_FONT_REDUCTION = 2;    /* 小さいボタンの文字サイズの縮小量（pt） / font size reduction for compact buttons */
    var CURRENT_SCALE_CHARS    = 12;   /* 現在の縮尺率・PPI の表示幅（文字数）。更新で値が長くなっても切れないように / width of the current scale & PPI text */
    var INFO_BOTTOM_GAP        = 10;   /* ピクセル寸法の下に足す余白（ボタン類との区切り） / extra gap below the pixel size row */
    var COMPACT_BUTTON_SPACING = 4;    /* 小さいボタンの間隔 / gap between compact buttons */
    var STEPS_PER_ROW          = 2;    /* 増減ボタンの1行の個数 / step buttons per row */
    var PRESETS_PER_ROW        = 2;    /* プリセットボタンの1行の個数 / preset buttons per row */
    var COLUMN_SPACING         = 10;   /* ボタン類の2カラムの間隔 / gap between the button columns */
    var COLUMN_PANEL_MARGINS   = [10, 16, 10, 10];   /* カラムのパネル余白 [左,上,右,下] / column panel margins */

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

    var uiLang = getCurrentLang();

    var LABELS = {
        palette: {
            title: { ja: "画像の縮尺率とフレームの合わせ", en: "Scale Images and Fit Frames" }
        },
        field: {
            scale:        { ja: "変更後の縮尺率", en: "New scale" },
            currentScale: { ja: "現在の縮尺率", en: "Current scale" },
            actualPpi:    { ja: "元のPPI", en: "Actual PPI" },
            effectivePpi: { ja: "編集後のPPI", en: "Effective PPI" },
            pixelSize:    { ja: "ピクセル寸法", en: "Pixel size" },
            step:         { ja: "増減", en: "Adjust" },
            preset:       { ja: "プリセット", en: "Presets" },
            fitWidth:     { ja: "幅に合わせる", en: "Fit width" }
        },
        value: {
            mixed:    { ja: "混在", en: "Mixed" },
            perFrame: { ja: "個別", en: "Varies" },
            none:     { ja: "—", en: "—" }
        },
        button: {
            cancel:       { ja: "キャンセル", en: "Cancel" },
            refresh:      { ja: "更新", en: "Refresh" },
            fitMargin:    { ja: "マージン幅", en: "Margins" },
            fitTextFrame: { ja: "親フレーム", en: "Parent Frame" }
        },
        tooltip: {
            scale: {
                ja: "画像の縦横の縮尺率。適用後、フレームを画像のサイズに合わせます。\n↑↓で1、shift+↑↓で10ずつ増減します。",
                en: "Horizontal and vertical scale of the image. Frames are then fitted to the image.\nUse Up/Down to change by 1, or shift+Up/Down by 10."
            },
            currentScale: {
                ja: "選択中の画像の縮尺率（水平／垂直）。値が画像ごとに異なる場合は「混在」と表示します。",
                en: "Scale of the selected images (horizontal / vertical). Shows \"Mixed\" when the images differ."
            },
            actualPpi: {
                ja: "画像そのものの解像度（リンクパネルの「実際の PPI」）。縦横が異なる場合は「水平 × 垂直」、画像ごとに異なる場合は「混在」と表示します。PDF にはありません。",
                en: "Resolution of the image file itself (Actual PPI in the Links panel). Shows \"horizontal × vertical\" when they differ, or \"Mixed\" when the images differ. Not available for PDF."
            },
            effectivePpi: {
                ja: "縮尺率を反映した解像度（リンクパネルの「有効 PPI」）。値を変えるたびに更新します。",
                en: "Resolution after scaling (Effective PPI in the Links panel). Updates as you change the value."
            },
            pixelSize: {
                ja: "画像の幅 × 高さ（px）。縮尺率を変えても変わりません。画像ごとに異なる場合は「混在」と表示します。PDF にはありません。",
                en: "Width × height of the image in pixels. Does not change with the scale. Shows \"Mixed\" when the images differ. Not available for PDF."
            },
            refresh: {
                ja: "ここまでの変更を確定し、いま選択しているフレームを読み込み直します。",
                en: "Keeps the changes so far and reloads the currently selected frames."
            },
            cancel: {
                ja: "最後に［更新］してから（または開いてから）の変更を取り消して閉じます。\nウィンドウの閉じるボタンで閉じると、変更は残ります。",
                en: "Reverts the changes since the last Refresh (or since opening) and closes.\nClosing the window with its close button keeps the changes."
            },
            step:   { ja: "「変更後の縮尺率」の値を増減します。", en: "Changes the value in the New scale field." },
            preset: { ja: "「変更後の縮尺率」にこの値を入れます。", en: "Sets this value in the New scale field." },
            fitMargin: {
                ja: "画像の幅がページのマージン幅に収まる縮尺率（整数・切り捨て）をフレームごとに求め、フレームの左端を左マージンにそろえます。\nアンカー付きのフレームは位置を動かしません。",
                en: "Scales each image to the largest whole percentage that fits the page margin width, and moves the frame's left edge to the left margin.\nAnchored frames are not moved."
            },
            fitTextFrame: {
                ja: "テキストにアンカーされた画像を、親テキストフレームの段幅（インセットと段落のインデントを除く）に収まる縮尺率（整数・切り捨て）にします。\nアンカーされていないフレームは変更しません。",
                en: "Scales each anchored image to the largest whole percentage that fits the column width of its text frame (excluding insets and paragraph indents).\nFrames that are not anchored are left unchanged."
            }
        },
        alert: {
            noDocument:    { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTargetFrame: {
                ja: "選択範囲に対象のフレームがありません。\n画像（またはPDF）を1点だけ配置した長方形フレームを選択してください。",
                en: "No applicable frames in the selection.\nSelect rectangle frames that each contain a single image (or PDF)."
            }
        }
    };

    /**
     * ラベルを現在の言語で取得する
     * @param {object} labelSet ja / en を持つラベルオブジェクト
     * @returns {string} 現在の言語のラベル文字列
     */
    function getLabel(labelSet) {
        return labelSet[uiLang] || labelSet.en;
    }

    /**
     * 言語別のコロンを付けたラベルを返す（日本語は全角、英語は半角）
     * @param {object} labelSet ja / en を持つラベルオブジェクト
     * @returns {string} コロン付きラベル
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // パレット / Palette
    // =========================================

    /**
     * 縮尺率を増減する（±1 はそのまま、±5・±10 はその倍数へ寄せる。下限は 1）
     * @param {number} value 現在の値
     * @param {number} delta 増減量
     * @returns {number} 増減後の値
     */
    function stepScaleValue(value, delta) {
        var stepSize = Math.abs(delta);
        var nextValue = (stepSize === 1)
            ? Math.round(value + delta)
            : (delta > 0 ? Math.ceil((value + 1) / stepSize) : Math.floor((value - 1) / stepSize)) * stepSize;
        return Math.max(nextValue, 1);
    }

    /**
     * 行ラベル（右揃え・固定幅）を追加する
     * @param {Group} rowGroup 追加先の行
     * @param {object} labelSet ja / en を持つラベルオブジェクト
     * @returns {StaticText} 追加したラベル
     */
    function addRowLabel(rowGroup, labelSet) {
        var rowLabel = rowGroup.add("statictext", undefined, labelText(labelSet));
        rowLabel.preferredSize.width = ROW_LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowLabel;
    }

    /**
     * 行グループを追加する（ラベル付き）
     * @param {Container} parent 追加先
     * @param {object} labelSet ja / en を持つラベルオブジェクト
     * @returns {Group} 追加した行
     */
    function addLabeledRow(parent, labelSet) {
        var rowGroup = parent.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        addRowLabel(rowGroup, labelSet);
        return rowGroup;
    }

    /**
     * ボタン類のカラムになるパネルを追加する
     * @param {Group} parent 追加先のカラム
     * @param {object} labelSet パネル見出しのラベル
     * @returns {Panel} 追加したパネル
     */
    function addColumnPanel(parent, labelSet) {
        var columnPanel = parent.add("panel", undefined, getLabel(labelSet));
        columnPanel.orientation = "column";
        columnPanel.alignChildren = ["fill", "top"];
        columnPanel.margins = COLUMN_PANEL_MARGINS;
        columnPanel.spacing = COMPACT_BUTTON_SPACING;
        return columnPanel;
    }

    /**
     * ボタンをひとまわり小さくする（文字サイズと高さを詰める）
     * @param {Button} targetButton 対象のボタン
     * @param {number} [buttonWidth] 幅。省略時は文字に合わせる
     * @returns {void}
     */
    function makeCompactButton(targetButton, buttonWidth) {
        var baseFont = targetButton.graphics.font;
        targetButton.graphics.font = ScriptUI.newFont(baseFont.name, "REGULAR", baseFont.size - COMPACT_FONT_REDUCTION);
        targetButton.preferredSize.height = COMPACT_BUTTON_HEIGHT;
        if (buttonWidth) targetButton.preferredSize.width = buttonWidth;
    }

    /**
     * 値ボタンを並べる（行ごとに perRow 個）
     * @param {Panel} parent 追加先（カラムのパネル）
     * @param {Array} buttonValues ボタンに割り当てる値
     * @param {function} formatValue 値からボタン名を作る処理
     * @param {string} tooltipText ボタンの helpTip
     * @param {function} onValueClick 押されたときに値を受け取る処理
     * @param {number} perRow 1行の個数
     * @returns {void}
     */
    function addValueButtons(parent, buttonValues, formatValue, tooltipText, onValueClick, perRow) {
        var buttonColumnGroup = parent.add("group");
        buttonColumnGroup.orientation = "column";
        buttonColumnGroup.alignChildren = ["left", "center"];
        buttonColumnGroup.spacing = COMPACT_BUTTON_SPACING;

        var buttonRowGroup = null;
        for (var i = 0; i < buttonValues.length; i++) {
            if (i % perRow === 0) {
                buttonRowGroup = buttonColumnGroup.add("group");
                buttonRowGroup.orientation = "row";
                buttonRowGroup.spacing = COMPACT_BUTTON_SPACING;
            }
            var valueButton = buttonRowGroup.add("button", undefined, formatValue(buttonValues[i]));
            makeCompactButton(valueButton, COMPACT_BUTTON_WIDTH);
            valueButton.helpTip = tooltipText;
            valueButton.buttonValue = buttonValues[i];
            valueButton.onClick = function () { onValueClick(this.buttonValue); };
        }
    }

    /**
     * 入力欄に上下キーでの増減操作を追加する（shift で 10 の倍数へ）
     * @param {EditText} editText 対象の入力欄
     * @param {function} onValueChange 値の変更後に呼ぶ処理
     * @returns {void}
     */
    function changeValueByArrowKey(editText, onValueChange) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var stepSize = ScriptUI.environment.keyboardState.shiftKey ? 10 : 1;
            event.preventDefault();
            editText.text = String(stepScaleValue(value, (event.keyName === "Up") ? stepSize : -stepSize));
            onValueChange();
        });
    }

    /**
     * 縮尺率を表示用に丸める（小数第2位まで）
     * @param {number} scalePercent 縮尺率（%）
     * @returns {string} 表示用の文字列
     */
    function formatPercent(scalePercent) {
        return String(Math.round(scalePercent * 100) / 100);
    }

    /**
     * 画像の PPI を表示用の文字列にする
     * @param {Array} scaleTargets collectScaleTargets() の戻り値
     * @param {string} ppiProperty "actualPpi" または "effectivePpi"
     * @returns {string} 例："300"、"300 × 350"、"混在"。PPI を持つ画像が無ければ "—"
     */
    function describePpi(scaleTargets, ppiProperty) {
        var ppiText = null;
        for (var i = 0; i < scaleTargets.length; i++) {
            var placedGraphic = scaleTargets[i].graphic;
            /* PDF は PPI を持たない / PDF has no PPI */
            if (!placedGraphic.isValid || placedGraphic.reflect.name !== "Image") continue;
            var ppiPair = placedGraphic[ppiProperty];
            var horizontalPpi = Math.round(ppiPair[0]);
            var verticalPpi = Math.round(ppiPair[1]);
            var itemText = (horizontalPpi === verticalPpi) ? String(horizontalPpi) : horizontalPpi + " \u00D7 " + verticalPpi;
            if (ppiText !== null && itemText !== ppiText) return getLabel(LABELS.value.mixed);
            ppiText = itemText;
        }
        return (ppiText === null) ? getLabel(LABELS.value.none) : ppiText;
    }

    /**
     * 画像の基準点をペーストボードの座標（pt）で返す（回転していても画像の角を追う）
     * @param {Image} placedImage 対象の画像
     * @param {AnchorPoint} anchorPoint 基準点
     * @returns {Array} [x, y]（pt）
     */
    function resolveImageCorner(placedImage, anchorPoint) {
        return placedImage.resolve(
            [anchorPoint, BoundingBoxLimits.GEOMETRIC_PATH_BOUNDS, CoordinateSpaces.INNER_COORDINATES],
            CoordinateSpaces.PASTEBOARD_COORDINATES
        )[0];
    }

    /**
     * 2点間の距離を返す
     * @param {Array} startPoint [x, y]
     * @param {Array} endPoint [x, y]
     * @returns {number} 距離
     */
    function getDistance(startPoint, endPoint) {
        var dx = endPoint[0] - startPoint[0];
        var dy = endPoint[1] - startPoint[1];
        return Math.sqrt(dx * dx + dy * dy);
    }

    /**
     * 画像のピクセル寸法を表示用の文字列にする（表示サイズ pt × 有効 PPI ÷ 72）
     * @param {Array} scaleTargets collectScaleTargets() の戻り値
     * @returns {string} 例："3000 × 2000 px"、"混在"。画像が無ければ "—"
     */
    function describePixelSize(scaleTargets) {
        var sizeText = null;
        for (var i = 0; i < scaleTargets.length; i++) {
            var placedGraphic = scaleTargets[i].graphic;
            /* PDF はピクセルを持たない / PDF has no pixels */
            if (!placedGraphic.isValid || placedGraphic.reflect.name !== "Image") continue;
            var topLeft = resolveImageCorner(placedGraphic, AnchorPoint.TOP_LEFT_ANCHOR);
            var topRight = resolveImageCorner(placedGraphic, AnchorPoint.TOP_RIGHT_ANCHOR);
            var bottomLeft = resolveImageCorner(placedGraphic, AnchorPoint.BOTTOM_LEFT_ANCHOR);
            var effectivePpi = placedGraphic.effectivePpi;
            var pixelWidth = Math.round(getDistance(topLeft, topRight) / 72 * effectivePpi[0]);
            var pixelHeight = Math.round(getDistance(topLeft, bottomLeft) / 72 * effectivePpi[1]);
            var itemText = pixelWidth + " \u00D7 " + pixelHeight + " px";
            if (sizeText !== null && itemText !== sizeText) return getLabel(LABELS.value.mixed);
            sizeText = itemText;
        }
        return (sizeText === null) ? getLabel(LABELS.value.none) : sizeText;
    }

    /**
     * 現在の縮尺率を表示用の文字列にする
     * @param {object|null} currentScale { horizontal, vertical }。画像ごとに異なる場合は null
     * @returns {string} 例："25%"、"25% / 30%"、"混在"
     */
    function describeCurrentScale(currentScale) {
        if (currentScale === null) return getLabel(LABELS.value.mixed);
        var horizontalText = formatPercent(currentScale.horizontal) + "%";
        if (currentScale.horizontal === currentScale.vertical) return horizontalText;
        return horizontalText + " / " + formatPercent(currentScale.vertical) + "%";
    }

    /**
     * 入力欄の縮尺率を読む
     * @param {EditText} scaleInput 縮尺率の入力欄
     * @returns {number|null} 0 より大きい数値。無効なら null
     */
    function readScalePercent(scaleInput) {
        var value = Number(scaleInput.text);
        return (value > 0) ? value : null;
    }

    /**
     * 幅に合わせたときの縮尺率を表示用の文字列にする
     * @param {Array} scaleTargets 対象の配列
     * @param {string} fitMode FIT_MODE の値
     * @returns {string} 全フレーム共通なら数値、異なれば「個別」
     */
    function describeFitPercent(scaleTargets, fitMode) {
        var commonPercent = null;
        for (var i = 0; i < scaleTargets.length; i++) {
            var fitPercent = scaleTargets[i].fitPercents[fitMode];
            if (fitPercent === null) continue;
            if (commonPercent !== null && fitPercent !== commonPercent) return getLabel(LABELS.value.perFrame);
            commonPercent = fitPercent;
        }
        return String(commonPercent);
    }

    /**
     * いずれかのフレームで幅に合わせられるか
     * @param {Array} scaleTargets 対象の配列
     * @param {string} fitMode FIT_MODE の値
     * @returns {boolean} 合わせられるフレームがあれば true
     */
    function canFitAny(scaleTargets, fitMode) {
        for (var i = 0; i < scaleTargets.length; i++) {
            if (scaleTargets[i].fitPercents[fitMode] !== null) return true;
        }
        return false;
    }

    /**
     * 縮尺率を入力するパレットを表示する（変更はその場でドキュメントに反映）
     * InDesign はモーダルダイアログの表示中にドキュメントを変更できない（Error 30486）ため、パレットにしている
     * @param {object} initialState readSelectionState() の戻り値
     * @returns {void}
     */
    function showScalePalette(initialState) {
        var selectionState = initialState;

        var scalePalette = new Window("palette", getLabel(LABELS.palette.title) + "  " + SCRIPT_VERSION);
        scalePalette.orientation = "column";
        scalePalette.alignChildren = "fill";
        scalePalette.margins = PALETTE_MARGINS;
        scalePalette.spacing = PALETTE_SPACING;

        var currentScaleRowGroup = addLabeledRow(scalePalette, LABELS.field.currentScale);
        var currentScaleText = currentScaleRowGroup.add("statictext", undefined, "");
        currentScaleText.characters = CURRENT_SCALE_CHARS;
        currentScaleText.helpTip = getLabel(LABELS.tooltip.currentScale);

        var scaleRowGroup = addLabeledRow(scalePalette, LABELS.field.scale);
        var scaleInput = scaleRowGroup.add("edittext", undefined, "");
        scaleInput.characters = SCALE_FIELD_CHARS;
        scaleInput.helpTip = getLabel(LABELS.tooltip.scale);
        scaleRowGroup.add("statictext", undefined, "%");

        var actualPpiRowGroup = addLabeledRow(scalePalette, LABELS.field.actualPpi);
        var actualPpiText = actualPpiRowGroup.add("statictext", undefined, "");
        actualPpiText.characters = CURRENT_SCALE_CHARS;
        actualPpiText.helpTip = getLabel(LABELS.tooltip.actualPpi);

        var effectivePpiRowGroup = addLabeledRow(scalePalette, LABELS.field.effectivePpi);
        var effectivePpiText = effectivePpiRowGroup.add("statictext", undefined, "");
        effectivePpiText.characters = CURRENT_SCALE_CHARS;
        effectivePpiText.helpTip = getLabel(LABELS.tooltip.effectivePpi);

        var pixelSizeRowGroup = addLabeledRow(scalePalette, LABELS.field.pixelSize);
        pixelSizeRowGroup.margins = [0, 0, 0, INFO_BOTTOM_GAP];
        var pixelSizeText = pixelSizeRowGroup.add("statictext", undefined, "");
        pixelSizeText.characters = CURRENT_SCALE_CHARS;
        pixelSizeText.alignment = ["fill", "center"];   /* 「3000 × 2000 px」でも切れないよう行の残り幅まで伸ばす / stretch to the row width */
        pixelSizeText.helpTip = getLabel(LABELS.tooltip.pixelSize);

        /* ボタン類は2カラム（左：増減＋幅に合わせる／右：プリセット）/ Two columns: steps + fit width, presets */
        var buttonColumnsGroup = scalePalette.add("group");
        buttonColumnsGroup.orientation = "row";
        buttonColumnsGroup.alignChildren = ["fill", "top"];
        buttonColumnsGroup.spacing = COLUMN_SPACING;
        var leftColumnGroup = buttonColumnsGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = ["fill", "top"];
        leftColumnGroup.spacing = COLUMN_SPACING;
        var stepPanel = addColumnPanel(leftColumnGroup, LABELS.field.step);
        var fitPanel = addColumnPanel(leftColumnGroup, LABELS.field.fitWidth);
        var presetPanel = addColumnPanel(buttonColumnsGroup, LABELS.field.preset);

        var btnFitMargin = fitPanel.add("button", undefined, getLabel(LABELS.button.fitMargin));
        var btnFitTextFrame = fitPanel.add("button", undefined, getLabel(LABELS.button.fitTextFrame));
        makeCompactButton(btnFitMargin);
        makeCompactButton(btnFitTextFrame);
        btnFitMargin.helpTip = getLabel(LABELS.tooltip.fitMargin);
        btnFitTextFrame.helpTip = getLabel(LABELS.tooltip.fitTextFrame);

        /* ボタン（左：キャンセル／右：更新）/ Buttons (left: Cancel / right: Refresh) */
        var btnRowGroup = scalePalette.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["fill", "center"];
        btnRowGroup.alignChildren = ["left", "center"];
        var btnCancel = btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnCancel.helpTip = getLabel(LABELS.tooltip.cancel);
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "center"];
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignment = ["right", "center"];
        var btnRefresh = btnRightGroup.add("button", undefined, getLabel(LABELS.button.refresh));
        btnRefresh.helpTip = getLabel(LABELS.tooltip.refresh);

        var scalePreview = null;
        var activeFitMode = null;   /* 幅に合わせるボタンを押した状態。数値を入れ直すと解除 / set by fit buttons, cleared by typing */

        /**
         * 読み込んだ選択状態をパレットの表示に反映する
         * @returns {void}
         */
        function loadSelectionState() {
            var scaleTargets = selectionState.scaleTargets;
            scalePreview = createScalePreview(selectionState.targetDocument, scaleTargets);
            activeFitMode = null;
            currentScaleText.text = (scaleTargets.length > 0)
                ? describeCurrentScale(selectionState.currentScale)
                : getLabel(LABELS.value.none);
            actualPpiText.text = describePpi(scaleTargets, "actualPpi");
            effectivePpiText.text = describePpi(scaleTargets, "effectivePpi");
            pixelSizeText.text = describePixelSize(scaleTargets);
            scaleInput.text = String(selectionState.initialPercent);
            btnFitMargin.enabled = canFitAny(scaleTargets, FIT_MODE.margin);
            btnFitTextFrame.enabled = canFitAny(scaleTargets, FIT_MODE.textFrame);
        }

        /**
         * 入力内容から適用内容を読む
         * @returns {object|null} { fitMode, scalePercent }。数値が無効なら null
         */
        function readScaleSpec() {
            if (activeFitMode !== null) return { fitMode: activeFitMode, scalePercent: null };
            var scalePercent = readScalePercent(scaleInput);
            return (scalePercent === null) ? null : { fitMode: null, scalePercent: scalePercent };
        }

        /**
         * 入力内容に合わせて、直前の変更を戻してから掛け直す
         * @returns {void}
         */
        function refreshPreview() {
            scalePreview.clear();
            var scaleSpec = readScaleSpec();
            if (scaleSpec !== null) scalePreview.apply(scaleSpec);
            effectivePpiText.text = describePpi(selectionState.scaleTargets, "effectivePpi");
        }

        /**
         * 縮尺率の欄に値を入れて掛け直す
         * @param {number} scalePercent 縮尺率（%）
         * @returns {void}
         */
        function setScaleValue(scalePercent) {
            activeFitMode = null;
            scaleInput.text = String(scalePercent);
            refreshPreview();
        }

        /**
         * 幅に合わせる基準を選び、掛け直す
         * @param {string} fitMode FIT_MODE の値
         * @returns {void}
         */
        function selectFitMode(fitMode) {
            activeFitMode = fitMode;
            scaleInput.text = describeFitPercent(selectionState.scaleTargets, fitMode);
            refreshPreview();
        }

        btnFitMargin.onClick = function () { selectFitMode(FIT_MODE.margin); };
        btnFitTextFrame.onClick = function () { selectFitMode(FIT_MODE.textFrame); };

        addValueButtons(stepPanel, SCALE_STEPS,
            function (delta) { return (delta > 0 ? "+" : "\u2212") + Math.abs(delta); },
            getLabel(LABELS.tooltip.step),
            function (delta) {
                var value = Number(scaleInput.text);
                setScaleValue(stepScaleValue(isNaN(value) ? selectionState.initialPercent : value, delta));
            },
            STEPS_PER_ROW);
        addValueButtons(presetPanel, SCALE_PRESETS,
            function (presetPercent) { return presetPercent + "%"; },
            getLabel(LABELS.tooltip.preset),
            setScaleValue,
            PRESETS_PER_ROW);

        scaleInput.onChanging = function () {
            activeFitMode = null;
            refreshPreview();
        };
        changeValueByArrowKey(scaleInput, function () {
            activeFitMode = null;
            refreshPreview();
        });

        /* ここまでの変更を残し、選択を読み込み直す / Keep the changes and reload the selection */
        btnRefresh.onClick = function () {
            scalePreview.keep();
            selectionState = readSelectionState();
            if (selectionState.alertLabel) alert(getLabel(selectionState.alertLabel));
            loadSelectionState();
        };

        var isCancelled = false;
        btnCancel.onClick = function () {
            isCancelled = true;
            scalePalette.close();
        };

        /* ［キャンセル］だけが戻す。閉じるボタンでは変更を残す / Only Cancel reverts; the close box keeps the changes */
        scalePalette.onClose = function () {
            if (isCancelled) scalePreview.clear();
            $.global.idSetImageScalePalette = null;
        };

        scalePalette.onShow = function () { scaleInput.active = true; };
        scalePalette.cancelElement = btnCancel;

        loadSelectionState();

        /* パレットは show() から即座に戻るので、参照を $.global に残す / Keep a reference: palettes return from show() at once */
        $.global.idSetImageScalePalette = scalePalette;
        scalePalette.show();
    }

    // =========================================
    // 処理 / Processing
    // =========================================

    /**
     * 縮尺変更の対象になる画像を返す（長方形フレームに画像または PDF が1点だけ入っている場合）
     * @param {PageItem} pageItem 選択中のオブジェクト
     * @returns {Graphic|null} 対象の画像。対象外なら null
     */
    function getScalableGraphic(pageItem) {
        if (pageItem.reflect.name !== "Rectangle" || pageItem.allGraphics.length !== 1) return null;
        var placedGraphic = pageItem.allGraphics[0];
        var graphicType = placedGraphic.constructor.name;
        return (graphicType === "Image" || graphicType === "PDF") ? placedGraphic : null;
    }

    /**
     * フレームがテキストにアンカーされているか
     * @param {PageItem} graphicFrame 対象のフレーム
     * @returns {boolean} アンカー付き（インライン含む）なら true
     */
    function isAnchoredFrame(graphicFrame) {
        return graphicFrame.parent.reflect.name === "Character";
    }

    /**
     * ページの左右マージンの X 座標を返す（見開きの左ページは内側・外側を入れ替える）
     * @param {Page} targetPage 対象ページ
     * @returns {object} { left, right }
     */
    function getMarginEdges(targetPage) {
        var pageBounds = targetPage.bounds;
        var marginPrefs = targetPage.marginPreferences;
        var isLeftHand = (targetPage.side === PageSideOptions.LEFT_HAND);
        return {
            left:  pageBounds[1] + (isLeftHand ? marginPrefs.right : marginPrefs.left),
            right: pageBounds[3] - (isLeftHand ? marginPrefs.left : marginPrefs.right)
        };
    }

    /**
     * テキストフレームの左右インセットの合計を返す
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @returns {number} 左右インセットの合計
     */
    function getHorizontalInsets(textFrame) {
        var insetSpacing = textFrame.textFramePreferences.insetSpacing;
        return (insetSpacing instanceof Array) ? insetSpacing[1] + insetSpacing[3] : insetSpacing * 2;
    }

    /**
     * アンカー付きフレームが入る段の幅を返す（インセット・段間・段落の左右インデントを除く）
     * @param {PageItem} graphicFrame 対象のフレーム
     * @returns {number|null} 段の幅。アンカーされていない・親が見つからない場合は null
     */
    function getParentColumnWidth(graphicFrame) {
        if (!isAnchoredFrame(graphicFrame)) return null;
        var anchorCharacter = graphicFrame.parent;
        var parentTextFrames = anchorCharacter.parentTextFrames;
        if (parentTextFrames.length === 0) return null;   /* オーバーセット / overset */

        var textFrame = parentTextFrames[0];
        var framePrefs = textFrame.textFramePreferences;
        var frameBounds = textFrame.geometricBounds;
        var innerWidth = frameBounds[3] - frameBounds[1] - getHorizontalInsets(textFrame);
        var columnCount = framePrefs.textColumnCount;
        var columnWidth = (innerWidth - framePrefs.textColumnGutter * (columnCount - 1)) / columnCount;

        var anchorParagraph = anchorCharacter.paragraphs[0];
        return columnWidth - anchorParagraph.leftIndent - anchorParagraph.rightIndent;
    }

    /**
     * 指定の幅に収まる縮尺率（整数・切り捨て）を返す
     * @param {Graphic} placedGraphic 対象の画像
     * @param {number|null} targetWidth 収めたい幅
     * @returns {number|null} 縮尺率（%）。求められなければ null
     */
    function getFitPercent(placedGraphic, targetWidth) {
        if (targetWidth === null || targetWidth <= 0) return null;
        var graphicBounds = placedGraphic.geometricBounds;
        var graphicWidth = graphicBounds[3] - graphicBounds[1];
        if (graphicWidth <= 0) return null;
        /* 幅は水平の縮尺率に比例する。浮動小数の誤差で 1 下がらないよう少し足す / width is proportional to horizontalScale */
        var fitPercent = Math.floor(placedGraphic.horizontalScale * targetWidth / graphicWidth + 1e-6);
        return (fitPercent >= 1) ? fitPercent : null;
    }

    /**
     * 縮尺変更の対象を1件作る（幅に合わせるときの縮尺率は変更前の状態で求めておく）
     * @param {PageItem} graphicFrame 対象のフレーム
     * @param {Graphic} placedGraphic フレーム内の画像
     * @returns {object} { frame, graphic, marginLeft, fitPercents }
     */
    function createScaleTarget(graphicFrame, placedGraphic) {
        var targetPage = graphicFrame.parentPage;
        var marginEdges = targetPage ? getMarginEdges(targetPage) : null;
        return {
            frame: graphicFrame,
            graphic: placedGraphic,
            /* アンカー付きは位置を動かさない / anchored frames are not moved */
            marginLeft: (marginEdges && !isAnchoredFrame(graphicFrame)) ? marginEdges.left : null,
            fitPercents: {
                margin:    getFitPercent(placedGraphic, marginEdges ? marginEdges.right - marginEdges.left : null),
                textFrame: getFitPercent(placedGraphic, getParentColumnWidth(graphicFrame))
            }
        };
    }

    /**
     * 選択範囲から縮尺変更の対象を集める
     * @param {Array} selectedItems 選択中のオブジェクト
     * @returns {Array} createScaleTarget() の戻り値の配列
     */
    function collectScaleTargets(selectedItems) {
        var scaleTargets = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var placedGraphic = getScalableGraphic(selectedItems[i]);
            if (placedGraphic) scaleTargets.push(createScaleTarget(selectedItems[i], placedGraphic));
        }
        return scaleTargets;
    }

    /**
     * 対象の画像に共通する現在の縮尺率を返す
     * @param {Array} scaleTargets collectScaleTargets() の戻り値
     * @returns {object|null} { horizontal, vertical }。画像ごとに異なる場合は null
     */
    function getCommonScale(scaleTargets) {
        var firstGraphic = scaleTargets[0].graphic;
        var commonScale = { horizontal: firstGraphic.horizontalScale, vertical: firstGraphic.verticalScale };
        for (var i = 1; i < scaleTargets.length; i++) {
            var placedGraphic = scaleTargets[i].graphic;
            if (placedGraphic.horizontalScale !== commonScale.horizontal || placedGraphic.verticalScale !== commonScale.vertical) return null;
        }
        return commonScale;
    }

    /**
     * 画像を縮尺変更し、フレームを画像に合わせる（マージン幅のときは左端を左マージンへ）
     * @param {Array} scaleTargets collectScaleTargets() の戻り値
     * @param {object} scaleSpec { fitMode, scalePercent }
     * @returns {void}
     */
    function scaleGraphicsAndFitFrames(scaleTargets, scaleSpec) {
        for (var i = 0; i < scaleTargets.length; i++) {
            var scaleTarget = scaleTargets[i];
            if (!scaleTarget.frame.isValid) continue;   /* パレット表示中に削除された / deleted while the palette is open */
            var scalePercent = scaleSpec.fitMode ? scaleTarget.fitPercents[scaleSpec.fitMode] : scaleSpec.scalePercent;
            if (scalePercent === null) continue;   /* 幅に合わせられないフレームは変えない / skip frames that cannot fit */

            scaleTarget.graphic.horizontalScale = scalePercent;
            scaleTarget.graphic.verticalScale = scalePercent;
            scaleTarget.frame.fit(FitOptions.FRAME_TO_CONTENT);

            if (scaleSpec.fitMode === FIT_MODE.margin && scaleTarget.marginLeft !== null) {
                var dx = scaleTarget.marginLeft - scaleTarget.frame.geometricBounds[1];
                scaleTarget.frame.move(undefined, [dx, 0]);
            }
        }
    }

    /**
     * 縮尺変更を 1 段の取り消しとして適用する
     * @param {Array} scaleTargets collectScaleTargets() の戻り値
     * @param {object} scaleSpec { fitMode, scalePercent }
     * @param {string} undoName 取り消しの名前
     * @returns {void}
     */
    function applyScaleAsOneUndo(scaleTargets, scaleSpec, undoName) {
        app.doScript(
            function () { scaleGraphicsAndFitFrames(scaleTargets, scaleSpec); },
            ScriptLanguage.JAVASCRIPT,
            undefined,
            UndoModes.ENTIRE_SCRIPT,
            undoName
        );
    }

    /**
     * 変更の適用・取り消し・確定を受け持つオブジェクトを作る
     * 変更は常に 1 段の取り消しにまとめ、掛け直すときは直前の 1 段を戻してから適用する
     * @param {Document|null} targetDocument 対象ドキュメント
     * @param {Array} scaleTargets collectScaleTargets() の戻り値
     * @returns {object} apply(scaleSpec)・clear()・keep() を持つオブジェクト
     */
    function createScalePreview(targetDocument, scaleTargets) {
        var previewUndoName = getLabel(LABELS.palette.title);
        var isApplied = false;
        return {
            apply: function (scaleSpec) {
                if (scaleTargets.length === 0 || !targetDocument.isValid) return;
                applyScaleAsOneUndo(scaleTargets, scaleSpec, previewUndoName);
                isApplied = true;
            },
            clear: function () {
                /* 直前の取り消しが自分の変更のときだけ戻す。ドキュメントが閉じられたり別の操作が挟まったりしたら戻さない */
                /* Undo only when the last step is ours; skip when the document was closed or another step intervened */
                if (isApplied && targetDocument.isValid && targetDocument.undoName === previewUndoName) targetDocument.undo();
                isApplied = false;
            },
            keep: function () {
                isApplied = false;
            }
        };
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * 現在の選択から対象と初期値を読み込む
     * @returns {object} { targetDocument, scaleTargets, currentScale, initialPercent, alertLabel }。対象が無いときは alertLabel に理由
     */
    function readSelectionState() {
        var selectionState = {
            targetDocument: null,
            scaleTargets: [],
            currentScale: null,
            initialPercent: DEFAULT_SCALE_PERCENT,
            alertLabel: null
        };
        if (app.documents.length === 0) {
            selectionState.alertLabel = LABELS.alert.noDocument;
            return selectionState;
        }
        selectionState.targetDocument = app.activeDocument;
        var selectedItems = selectionState.targetDocument.selection;
        /* 未選択は警告せず、空のまま開く（選択してから［更新］） / No selection: open empty without a warning */
        if (selectedItems.length === 0) return selectionState;
        selectionState.scaleTargets = collectScaleTargets(selectedItems);
        if (selectionState.scaleTargets.length === 0) {
            selectionState.alertLabel = LABELS.alert.noTargetFrame;
            return selectionState;
        }
        /* 縦横がそろった共通の縮尺率があれば初期値に使う / Start from the common scale when uniform */
        var currentScale = getCommonScale(selectionState.scaleTargets);
        selectionState.currentScale = currentScale;
        if (currentScale !== null && currentScale.horizontal === currentScale.vertical) {
            selectionState.initialPercent = Number(formatPercent(currentScale.horizontal));
        }
        return selectionState;
    }

    /* 前回のパレットが開いていれば閉じる（変更は残る） / Close a palette left open (its changes are kept) */
    if ($.global.idSetImageScalePalette) $.global.idSetImageScalePalette.close();

    /* 対象が無ければパレットを出さずに終える / Skip the palette when nothing applies */
    var initialState = readSelectionState();
    if (initialState.alertLabel) {
        alert(getLabel(initialState.alertLabel));
        return;
    }

    showScalePalette(initialState);

})();
