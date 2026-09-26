#target indesign
#targetengine "IdSetImageScale"

/*

### 概要

選択したフレーム内の画像の縮尺率を、パレットで数値・幅・高さ・PPI・プリセット・マージン幅・親フレームの段幅から指定して変更し、フレームを画像に合わせます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSetImageScale.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n91c6a628b7ed

### Overview

Changes the scale of the image in each selected frame, set from a palette by value, width, height, PPI, preset, margin width or parent column width, and fits the frame to the image.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSetImageScale.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdSetImageScale";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
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
    var SCALE_PRESETS = [10, 15, 20, 25, 30, 35, 40, 50, 100, 200, 300];   /* プリセットの縮尺率（%） / preset scales (%) */

    /* 幅に合わせるときの基準 / Width references for fitting */
    var FIT_MODE = {
        margin:    "margin",      /* ページのマージン幅 / page margin width */
        textFrame: "textFrame",   /* 親テキストフレームの段幅 / column width of the parent text frame */
        width:     "width",       /* 幅の欄に入力した幅 / width typed in the Width field */
        height:    "height",      /* 高さの欄に入力した高さ / height typed in the Height field */
        ppi:       "ppi"          /* PPI の欄に入力した有効 PPI / effective PPI typed in the PPI field */
    };

    /* 幅・高さで指定するときに比べる、縮尺率 100% のときの寸法 / natural size used by the Width / Height fields */
    var NATURAL_LENGTH_KEY = {
        width:  "naturalWidthPt",
        height: "naturalHeightPt"
    };

    var POINTS_PER_MM = 72 / 25.4;   /* 1mm あたりの pt / points per millimeter */

    var DEFAULT_ANCHOR_INDEX = 0;    /* 基準点の初期値（0..8 を行優先。0=左上, 4=中央, 8=右下）/ default reference point, row-major 0..8 */

    /* キーボードショートカットの初期値（機能 ID → 「修飾キー+キー」）/ default shortcuts: action ID → "modifiers+key" */
    var DEFAULT_SHORTCUTS = {
        "preset:100": "option+1",
        "fit:margin": "option+0"
    };
    /* 管理画面で選べる修飾キー。修飾キーなしは入力欄への文字入力とぶつかるので選べない / selectable modifiers (none would clash with typing) */
    var SHORTCUT_MODIFIERS = ["option", "control", "control+option"];
    var SHORTCUT_KEYS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");   /* 管理画面で選べるキー / selectable keys */

    /* InDesign には任意の値を残す環境設定 API が無いので、設定ファイルに key=value で書き出す
       / InDesign has no scriptable preference store, so settings go to a key=value file */
    var PREFS_FILE_NAME      = "IdSetImageScale-prefs.txt";
    var PREF_SHORTCUT_PREFIX = "shortcut.";

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PALETTE_MARGINS        = 16;   /* パレット外周の余白 / palette margin */
    var PALETTE_SPACING        = 12;   /* パレット内の要素間隔 / palette spacing */
    var SCALE_FIELD_CHARS      = 5;    /* 縮尺率欄の文字数 / scale field width in characters */
    var LENGTH_FIELD_CHARS     = 7;    /* 幅・高さの欄の文字数（小数の mm が入る）/ width & height fields in characters (fit decimal mm) */
    var ROW_LABEL_WIDTH        = 96;   /* 行ラベルの幅 / row label width */
    var COMPACT_BUTTON_WIDTH   = 44;   /* 小さいボタンの幅（増減・プリセット） / compact button width */
    var COMPACT_BUTTON_HEIGHT  = 20;   /* 小さいボタンの高さ / compact button height */
    var COMPACT_FONT_REDUCTION = 2;    /* 小さいボタンの文字サイズの縮小量（pt） / font size reduction for compact buttons */
    var INFO_TEXT_CHARS        = 12;   /* ピクセル寸法の表示幅（文字数）。更新で値が長くなっても切れないように / width of the pixel size text */
    var INFO_BOTTOM_GAP        = 10;   /* ピクセル寸法の下に足す余白（ボタン類との区切り） / extra gap below the pixel size row */
    var COMPACT_BUTTON_SPACING = 4;    /* 小さいボタンの間隔 / gap between compact buttons */
    var STEPS_PER_ROW          = 2;    /* 増減ボタンの1行の個数 / step buttons per row */
    var ANCHOR_WIDGET_SIZE     = 66;   /* 9軸ウィジェット全体の大きさ / overall size of the 9-axis widget */
    var ANCHOR_CELL_SIZE       = 9;    /* 9軸の□1個のサイズ / size of one anchor square */
    var ANCHOR_CELL_GAP        = 7.5;  /* 9軸の□どうしの間隔 / gap between anchor squares */
    var PRESETS_PER_ROW        = 1;    /* プリセットボタンの1行の個数 / preset buttons per row */
    var COLUMN_SPACING         = 10;   /* ボタン類の2カラムの間隔 / gap between the button columns */
    var COLUMN_PANEL_MARGINS   = [10, 16, 10, 10];   /* カラムのパネル余白 [左,上,右,下] / column panel margins */
    var ANCHOR_PANEL_MARGINS   = [4, 10, 4, 0];      /* 基準点パネルの余白 [左,上,右,下]。9軸自体に余白があるので詰める / anchor panel margins; the widget has its own padding */
    var SHORTCUT_LABEL_WIDTH   = 80;   /* 管理画面の機能名の幅 / action label width in the Manage dialog */
    var BUTTON_ROW_TOP_MARGIN  = 4;    /* 管理画面のボタン行の上余白 / top margin of the Manage dialog button row */

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
        manageDialog: {
            title: { ja: "管理：キーボードショートカット", en: "Manage: Keyboard Shortcuts" }
        },
        field: {
            scale:        { ja: "縮尺率", en: "Scale" },
            width:        { ja: "幅", en: "Width" },
            height:       { ja: "高さ", en: "Height" },
            ppi:          { ja: "PPI", en: "PPI" },
            pixelSize:    { ja: "ピクセル寸法", en: "Pixel size" },
            step:         { ja: "増減", en: "Adjust" },
            preset:       { ja: "プリセット", en: "Presets" },
            fitWidth:     { ja: "幅に合わせる", en: "Fit width" },
            anchor:       { ja: "基準点", en: "Reference point" }
        },
        value: {
            mixed:    { ja: "混在", en: "Mixed" },
            perFrame: { ja: "個別", en: "Varies" },
            none:     { ja: "—", en: "—" },
            noModifier: { ja: "なし", en: "None" }
        },
        checkbox: {
            roundToInteger: { ja: "縮尺率を整数に丸める", en: "Round scale to whole numbers" }
        },
        button: {
            fitMargin:    { ja: "マージン幅", en: "Margins" },
            fitTextFrame: { ja: "親フレーム", en: "Parent Frame" },
            manage:       { ja: "管理", en: "Manage" },
            cancel:       { ja: "キャンセル", en: "Cancel" },
            ok:           { ja: "OK", en: "OK" }
        },
        tooltip: {
            manage: { ja: "キーボードショートカットを設定する管理画面を表示します。", en: "Opens the window for setting keyboard shortcuts." },
            shortcutModifier: { ja: "修飾キー。「なし」でショートカットを外します。", en: "Modifier keys. Choose None to remove the shortcut." },
            shortcutKey: { ja: "組み合わせるキー", en: "Key to combine with the modifiers" },
            scale: {
                ja: "画像の縦横の縮尺率。適用後、フレームを画像のサイズに合わせます。\n↑↓で1、shift+↑↓で10ずつ増減します。",
                en: "Horizontal and vertical scale of the image. Frames are then fitted to the image.\nUse Up/Down to change by 1, or shift+Up/Down by 10."
            },
            currentScale: {
                ja: "選択中の画像の今の縮尺率（水平／垂直）。値が画像ごとに異なる場合は「混在」と表示します。",
                en: "Scale of the selected images (horizontal / vertical). Shows \"Mixed\" when the images differ."
            },
            actualPpi: {
                ja: "画像そのものの解像度（リンクパネルの「実際の PPI」）。縦横が異なる場合は「水平 × 垂直」、画像ごとに異なる場合は「混在」と表示します。PDF にはありません。",
                en: "Resolution of the image file itself (Actual PPI in the Links panel). Shows \"horizontal × vertical\" when they differ, or \"Mixed\" when the images differ. Not available for PDF."
            },
            effectivePpi: {
                ja: "縮尺率を反映した解像度（リンクパネルの「有効 PPI」、水平方向）。縮尺率を変えると更新します。\n入力すると、その PPI になる縮尺率をフレームごとに求めます。［縮尺率を整数に丸める］がオンなら、入力した PPI を下回らないよう縮尺率を切り捨てます。\n↑↓で1、shift+↑↓で10ずつ増減します。PDF にはありません。",
                en: "Resolution after scaling (Effective PPI in the Links panel, horizontal). Updates as the scale changes.\nType a PPI to scale each image to that resolution. With Round scale to whole numbers on, the scale is rounded down so the PPI never falls below the value.\nUse Up/Down to change by 1, or shift+Up/Down by 10. Not available for PDF."
            },
            pixelSize: {
                ja: "画像の幅 × 高さ（px）。縮尺率を変えても変わりません。画像ごとに異なる場合は「混在」と表示します。PDF にはありません。",
                en: "Width × height of the image in pixels. Does not change with the scale. Shows \"Mixed\" when the images differ. Not available for PDF."
            },
            step:   { ja: "縮尺率の入力欄の値を増減します。", en: "Changes the value in the Scale field." },
            preset: { ja: "縮尺率の入力欄にこの値を入れます。", en: "Sets this value in the Scale field." },
            width: {
                ja: "画像の幅（mm）。縮尺率を変えると更新します。\n入力すると、画像がその幅になる縮尺率をフレームごとに求めます。［縮尺率を整数に丸める］がオンなら、はみ出さないよう切り捨てます。\n↑↓で1mm、shift+↑↓で10mmずつ増減します。",
                en: "Width of the image in mm. Updates as the scale changes.\nType a width to scale each image to that width. With Round scale to whole numbers on, the scale is rounded down so the image never overflows.\nUse Up/Down to change by 1 mm, or shift+Up/Down by 10 mm."
            },
            height: {
                ja: "画像の高さ（mm）。縮尺率を変えると更新します。\n入力すると、画像がその高さになる縮尺率をフレームごとに求めます。［縮尺率を整数に丸める］がオンなら、はみ出さないよう切り捨てます。\n↑↓で1mm、shift+↑↓で10mmずつ増減します。",
                en: "Height of the image in mm. Updates as the scale changes.\nType a height to scale each image to that height. With Round scale to whole numbers on, the scale is rounded down so the image never overflows.\nUse Up/Down to change by 1 mm, or shift+Up/Down by 10 mm."
            },
            roundToInteger: {
                ja: "オンにすると、適用する縮尺率を整数にします。入力した値は四捨五入、［マージン幅］［親フレーム］ははみ出さないよう切り捨て、増減の ±5・±10 はその倍数へ寄せます。\nオフにすると丸めません。入力した値や増減の結果をそのまま使い、［マージン幅］［親フレーム］では幅にぴったりの縮尺率にします。",
                en: "When on, the applied scale is a whole number: typed values are rounded, Margins / Parent Frame round down so the image never overflows, and the ±5 / ±10 steps snap to multiples.\nWhen off, nothing is rounded: typed and stepped values are used as is, and Margins / Parent Frame use the exact scale that fits the width."
            },
            anchor: {
                ja: "縮尺を変えたあと、フレームのこの位置が元の場所に残るように配置します。\nアンカー付きのフレームは位置を動かしません。［マージン幅］では、左右の位置は左マージンが優先されます。",
                en: "After scaling, the frame is placed so this point stays where it was.\nAnchored frames are not moved. With Margins, the horizontal position follows the left margin."
            },
            fitMargin: {
                ja: "画像の幅がページのマージン幅に収まる縮尺率をフレームごとに求め、フレームの左端を左マージンにそろえます。\nアンカー付きのフレームは位置を動かしません。",
                en: "Scales each image to fit the page margin width, and moves the frame's left edge to the left margin.\nAnchored frames are not moved."
            },
            fitTextFrame: {
                ja: "テキストにアンカーされた画像を、親テキストフレームの段幅（インセットと段落のインデントを除く）に収まる縮尺率にします。\nアンカーされていないフレームは変更しません。",
                en: "Scales each anchored image to fit the column width of its text frame (excluding insets and paragraph indents).\nFrames that are not anchored are left unchanged."
            }
        },
        alert: {
            duplicateShortcut: {
                ja: "同じショートカットが複数の機能に割り当てられています：",
                en: "The same shortcut is assigned to more than one action:"
            },
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
     * 縮尺率を増減する
     * 丸めるときは ±1 を整数に、±5・±10 をその倍数へ寄せる（下限 1）。丸めないときは増減量をそのまま足す（0 以下になるなら据え置き）
     * @param {number} value 現在の値
     * @param {number} delta 増減量
     * @param {boolean} roundToInteger 整数に丸めるか
     * @returns {number} 増減後の値
     */
    function stepScaleValue(value, delta, roundToInteger) {
        if (!roundToInteger) {
            /* 小数の足し算の誤差を表示前に落とす / drop float noise from the addition */
            var exactValue = Number(formatNumber(value + delta));
            return (exactValue > 0) ? exactValue : value;
        }
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
     * 適用後の有効 PPI（水平方向・整数）を表示用の文字列にする
     * @param {Array} scaleTargets 対象の配列
     * @param {object|null} scaleSpec { fitMode, scalePercent, targetLengthPt, targetPpi, roundToInteger, anchorIndex }
     * @returns {string} 全フレーム共通なら数値、異なれば「個別」、PPI を持つ画像が無ければ ""
     */
    function describeEffectivePpi(scaleTargets, scaleSpec) {
        if (scaleSpec === null) return "";
        return describeCommonValue(scaleTargets, function (scaleTarget) {
            var scalePercent = resolveScalePercent(scaleTarget, scaleSpec);
            if (scalePercent === null || scaleTarget.actualPpi === null) return null;
            return Math.round(scaleTarget.actualPpi * 100 / scalePercent);
        });
    }

    /**
     * 幅・高さの入力行（ラベル＋入力欄＋mm）を追加する
     * @param {Container} parent 追加先
     * @param {object} labelSet 項目名のラベル
     * @param {object} tooltipSet ツールチップのラベル
     * @returns {EditText} 追加した入力欄
     */
    function addLengthRow(parent, labelSet, tooltipSet) {
        var lengthRowGroup = addLabeledRow(parent, labelSet);
        var lengthInput = lengthRowGroup.add("edittext", undefined, "");
        lengthInput.characters = LENGTH_FIELD_CHARS;
        lengthInput.helpTip = getLabel(tooltipSet);
        lengthRowGroup.add("statictext", undefined, "mm");
        return lengthInput;
    }

    /**
     * UI が明るいテーマか（InDesign は generalPreferences.uiBrightnessPreference、明るいと 1）
     * @returns {boolean} 明るいテーマなら true
     */
    function isLightUI() {
        return app.generalPreferences.uiBrightnessPreference > 0.5;
    }

    /**
     * 9軸の□を描くパスを作る（rectPath の前に newPath しないとパスが累積する）
     * @param {ScriptUIGraphics} graphics 描画先
     * @param {number} x 左端
     * @param {number} y 上端
     * @param {number} size 一辺
     * @returns {void}
     */
    function squarePath(graphics, x, y, size) {
        graphics.newPath();
        graphics.rectPath(x, y, size, size);
    }

    /* 中央(4)を除く外周の□どうしをつなぐケイ線の組み合わせ / Pairs of outer squares (center 4 excluded) joined by rules */
    var ANCHOR_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

    /**
     * 9軸ウィジェットを描画する（外周の□をケイ線でつなぐ・中央は独立。選択中の□だけ塗る）
     * @param {Button} anchorWidget 対象のウィジェット（selectedAnchorIndex を持つ）
     * @returns {void}
     */
    function drawAnchorWidget(anchorWidget) {
        var graphics = anchorWidget.graphics;
        var lightUI = isLightUI();
        var lineColor = lightUI ? [0.6, 0.6, 0.6, 1] : [0.55, 0.55, 0.55, 1];
        var selectedFill = lightUI ? [0.4, 0.4, 0.4, 1] : [0.8, 0.8, 0.8, 1];
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1);

        var cellStep = ANCHOR_CELL_SIZE + ANCHOR_CELL_GAP;
        var gridSize = ANCHOR_CELL_SIZE * 3 + ANCHOR_CELL_GAP * 2;
        var originX = Math.round((anchorWidget.size[0] - gridSize) / 2);
        var originY = Math.round((anchorWidget.size[1] - gridSize) / 2);
        var cellPositions = [];
        for (var cellIndex = 0; cellIndex < 9; cellIndex++) {
            cellPositions.push([originX + (cellIndex % 3) * cellStep, originY + Math.floor(cellIndex / 3) * cellStep]);
        }

        for (var i = 0; i < ANCHOR_CONNECTIONS.length; i++) {
            var cellA = cellPositions[ANCHOR_CONNECTIONS[i][0]];
            var cellB = cellPositions[ANCHOR_CONNECTIONS[i][1]];
            graphics.newPath();
            if (ANCHOR_CONNECTIONS[i][1] - ANCHOR_CONNECTIONS[i][0] === 1) {
                /* 横方向：右隣の□へ / Horizontal: to the square on the right */
                graphics.moveTo(cellA[0] + ANCHOR_CELL_SIZE, cellA[1] + ANCHOR_CELL_SIZE / 2);
                graphics.lineTo(cellB[0], cellB[1] + ANCHOR_CELL_SIZE / 2);
            } else {
                /* 縦方向：下の□へ / Vertical: to the square below */
                graphics.moveTo(cellA[0] + ANCHOR_CELL_SIZE / 2, cellA[1] + ANCHOR_CELL_SIZE);
                graphics.lineTo(cellB[0] + ANCHOR_CELL_SIZE / 2, cellB[1]);
            }
            graphics.strokePath(linePen);
        }

        for (var j = 0; j < cellPositions.length; j++) {
            /* 枠を上に描くので塗りを先に行う / Fill first so the border draws on top */
            if (j === anchorWidget.selectedAnchorIndex) {
                squarePath(graphics, cellPositions[j][0], cellPositions[j][1], ANCHOR_CELL_SIZE);
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, selectedFill));
            }
            squarePath(graphics, cellPositions[j][0], cellPositions[j][1], ANCHOR_CELL_SIZE);
            graphics.strokePath(linePen);
        }
    }

    /**
     * 9軸（3×3）の基準点ウィジェットを追加する（button を onDraw で描き、クリックしたセルを選ぶ）
     * @param {Container} parent 追加先
     * @param {function} onAnchorChange 基準点が変わったときに呼ぶ処理
     * @returns {Button} 追加したウィジェット（selectedAnchorIndex に 0..8）
     */
    function addAnchorWidget(parent, onAnchorChange) {
        var anchorWidget = parent.add("button", undefined, "");
        anchorWidget.alignment = ["center", "top"];
        anchorWidget.minimumSize = anchorWidget.preferredSize = anchorWidget.maximumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.helpTip = getLabel(LABELS.tooltip.anchor);
        anchorWidget.selectedAnchorIndex = DEFAULT_ANCHOR_INDEX;
        anchorWidget.onDraw = function () { drawAnchorWidget(this); };
        /* クリック座標（コントロール基準）を 3 分割してセルを決める / Hit-test the clicked cell by splitting the control into thirds */
        anchorWidget.addEventListener("mousedown", function (event) {
            var col = Math.min(2, Math.max(0, Math.floor(event.clientX / (anchorWidget.size[0] / 3))));
            var row = Math.min(2, Math.max(0, Math.floor(event.clientY / (anchorWidget.size[1] / 3))));
            var nextIndex = row * 3 + col;
            if (nextIndex === anchorWidget.selectedAnchorIndex) return;
            anchorWidget.selectedAnchorIndex = nextIndex;
            anchorWidget.notify("onDraw");
            onAnchorChange();
        });
        return anchorWidget;
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
        /* パネルの中で左右中央に置く / center horizontally in the panel */
        buttonColumnGroup.alignment = ["center", "top"];
        buttonColumnGroup.alignChildren = ["center", "center"];
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
     * 入力欄に上下キーでの増減操作を追加する（shift で 10 ずつ）
     * @param {EditText} editText 対象の入力欄
     * @param {function} stepValue (値, 増減量) から増減後の値を返す処理
     * @param {function} onValueChange 値の変更後に呼ぶ処理
     * @returns {void}
     */
    function changeValueByArrowKey(editText, stepValue, onValueChange) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var stepSize = ScriptUI.environment.keyboardState.shiftKey ? 10 : 1;
            event.preventDefault();
            editText.text = String(stepValue(value, (event.keyName === "Up") ? stepSize : -stepSize));
            onValueChange();
        });
    }

    /**
     * 数値を表示用に丸める（小数第2位まで）
     * @param {number} value 縮尺率（%）や幅（mm）
     * @returns {string} 表示用の文字列
     */
    function formatNumber(value) {
        return String(Math.round(value * 100) / 100);
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
     * @param {Graphic} placedGraphic 対象の画像
     * @param {AnchorPoint} anchorPoint 基準点
     * @returns {Array} [x, y]（pt）
     */
    function resolveGraphicCorner(placedGraphic, anchorPoint) {
        return placedGraphic.resolve(
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
            var topLeft = resolveGraphicCorner(placedGraphic, AnchorPoint.TOP_LEFT_ANCHOR);
            var topRight = resolveGraphicCorner(placedGraphic, AnchorPoint.TOP_RIGHT_ANCHOR);
            var bottomLeft = resolveGraphicCorner(placedGraphic, AnchorPoint.BOTTOM_LEFT_ANCHOR);
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
        var horizontalText = formatNumber(currentScale.horizontal) + "%";
        if (currentScale.horizontal === currentScale.vertical) return horizontalText;
        return horizontalText + " / " + formatNumber(currentScale.vertical) + "%";
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
     * 対象ごとの値をまとめて表示用の文字列にする
     * @param {Array} scaleTargets 対象の配列
     * @param {function} getItemValue 対象から数値（求められなければ null）を返す処理
     * @returns {string} 全フレーム共通なら数値、異なれば「個別」、1つも無ければ ""
     */
    function describeCommonValue(scaleTargets, getItemValue) {
        var commonText = null;
        for (var i = 0; i < scaleTargets.length; i++) {
            var itemValue = getItemValue(scaleTargets[i]);
            if (itemValue === null) continue;
            var itemText = formatNumber(itemValue);
            if (commonText !== null && itemText !== commonText) return getLabel(LABELS.value.perFrame);
            commonText = itemText;
        }
        return (commonText === null) ? "" : commonText;
    }

    /**
     * 幅に合わせたときの縮尺率を表示用の文字列にする
     * @param {Array} scaleTargets 対象の配列
     * @param {object} scaleSpec { fitMode, scalePercent, targetLengthPt, targetPpi, roundToInteger, anchorIndex }
     * @returns {string} 全フレーム共通なら数値、異なれば「個別」
     */
    function describeFitPercent(scaleTargets, scaleSpec) {
        return describeCommonValue(scaleTargets, function (scaleTarget) {
            return resolveScalePercent(scaleTarget, scaleSpec);
        });
    }

    /**
     * 適用後の画像の幅または高さ（mm）を表示用の文字列にする
     * @param {Array} scaleTargets 対象の配列
     * @param {object|null} scaleSpec { fitMode, scalePercent, targetLengthPt, targetPpi, roundToInteger, anchorIndex }
     * @param {string} lengthMode FIT_MODE.width または FIT_MODE.height
     * @returns {string} 全フレーム共通なら数値、異なれば「個別」
     */
    function describeLength(scaleTargets, scaleSpec, lengthMode) {
        if (scaleSpec === null) return "";
        var naturalKey = NATURAL_LENGTH_KEY[lengthMode];
        return describeCommonValue(scaleTargets, function (scaleTarget) {
            var scalePercent = resolveScalePercent(scaleTarget, scaleSpec);
            if (scalePercent === null || scaleTarget[naturalKey] === null) return null;
            return scaleTarget[naturalKey] * scalePercent / 100 / POINTS_PER_MM;
        });
    }

    /**
     * 読み込んだ対象を見分けるキーを返す（ドキュメント名＋フレームの ID）
     * @param {object} selectionState readSelectionState() の戻り値
     * @returns {string} 対象が同じなら同じ文字列
     */
    function getSelectionKey(selectionState) {
        var scaleTargets = selectionState.scaleTargets;
        if (scaleTargets.length === 0) return "";
        var frameIds = [];
        for (var i = 0; i < scaleTargets.length; i++) {
            /* 削除されたフレームは ID を読めない / deleted frames have no readable ID */
            frameIds.push(scaleTargets[i].frame.isValid ? scaleTargets[i].frame.id : "-");
        }
        return selectionState.targetDocument.name + ":" + frameIds.join(",");
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
     * ショートカットを割り当てられる機能の一覧を返す（増減・幅に合わせる・プリセット）
     * @returns {Array} { id, kind, value, labelText, groupLabel } の配列
     */
    function getShortcutActions() {
        var shortcutActions = [];

        /**
         * 機能を一覧に加える
         * @param {string} kind "step"／"fit"／"preset"
         * @param {number|string} value 増減量・FIT_MODE の値・縮尺率
         * @param {string} labelString 管理画面に出す機能名
         * @param {object} groupLabel 所属するパネルのラベル
         * @returns {void}
         */
        function addAction(kind, value, labelString, groupLabel) {
            shortcutActions.push({ id: kind + ":" + value, kind: kind, value: value, labelText: labelString, groupLabel: groupLabel });
        }
        for (var i = 0; i < SCALE_STEPS.length; i++) {
            addAction("step", SCALE_STEPS[i], (SCALE_STEPS[i] > 0 ? "+" : "\u2212") + Math.abs(SCALE_STEPS[i]), LABELS.field.step);
        }
        addAction("fit", FIT_MODE.margin, getLabel(LABELS.button.fitMargin), LABELS.field.fitWidth);
        addAction("fit", FIT_MODE.textFrame, getLabel(LABELS.button.fitTextFrame), LABELS.field.fitWidth);
        for (var j = 0; j < SCALE_PRESETS.length; j++) {
            addAction("preset", SCALE_PRESETS[j], SCALE_PRESETS[j] + "%", LABELS.field.preset);
        }
        return shortcutActions;
    }

    /**
     * 設定ファイルを返す
     * @returns {File} 設定ファイル
     */
    function getPrefsFile() {
        return File(Folder.userData.fsName + "/" + PREFS_FILE_NAME);
    }

    /**
     * ショートカットの割り当てを読み出す。記録の無い機能は初期値にする
     * @returns {Object<string, string>} 機能 ID → 「修飾キー+キー」（未割り当ては ""）
     */
    function loadShortcuts() {
        var shortcuts = {};
        for (var actionId in DEFAULT_SHORTCUTS) {
            if (DEFAULT_SHORTCUTS.hasOwnProperty(actionId)) shortcuts[actionId] = DEFAULT_SHORTCUTS[actionId];
        }

        var prefsFile = getPrefsFile();
        prefsFile.encoding = "UTF-8";
        if (!prefsFile.exists || !prefsFile.open("r")) return shortcuts;
        var lines = prefsFile.read().split("\n");
        prefsFile.close();

        for (var i = 0; i < lines.length; i++) {
            var separatorIndex = lines[i].indexOf("=");
            if (separatorIndex < 0 || lines[i].indexOf(PREF_SHORTCUT_PREFIX) !== 0) continue;
            shortcuts[lines[i].substring(PREF_SHORTCUT_PREFIX.length, separatorIndex)] = lines[i].substring(separatorIndex + 1);
        }
        return shortcuts;
    }

    /**
     * ショートカットの割り当てを設定ファイルに書き出す（未割り当ても "" で残し、初期値に戻らないようにする）
     * @param {Object<string, string>} shortcuts 機能 ID → 「修飾キー+キー」
     * @returns {void}
     */
    function saveShortcuts(shortcuts) {
        var shortcutActions = getShortcutActions();
        var lines = [];
        for (var i = 0; i < shortcutActions.length; i++) {
            lines.push(PREF_SHORTCUT_PREFIX + shortcutActions[i].id + "=" + (shortcuts[shortcutActions[i].id] || ""));
        }

        /* 保存できなくてもこのセッションでは使えるので、書き出せたかどうかは見ない / A failed save still works for this session */
        var prefsFile = getPrefsFile();
        prefsFile.encoding = "UTF-8";
        if (!prefsFile.open("w")) return;
        prefsFile.write(lines.join("\n"));
        prefsFile.close();
    }

    /**
     * 押されたキーを「修飾キー+キー」の文字列にする
     * @param {KeyboardEvent} event keydown のイベント
     * @returns {string} 例："option+1"。修飾キーが無ければ ""
     */
    function getPressedShortcut(event) {
        var keyboardState = ScriptUI.environment.keyboardState;
        var modifiers = [];
        if (keyboardState.ctrlKey) modifiers.push("control");
        if (keyboardState.altKey) modifiers.push("option");
        if (modifiers.length === 0 || !event.keyName) return "";
        return modifiers.join("+") + "+" + String(event.keyName).toUpperCase();
    }

    /**
     * 押されたキーに割り当てた機能を探す
     * @param {Object<string, string>} shortcuts 機能 ID → 「修飾キー+キー」
     * @param {string} pressedShortcut getPressedShortcut() の戻り値
     * @returns {object|null} getShortcutActions() の要素。割り当てが無ければ null
     */
    function findShortcutAction(shortcuts, pressedShortcut) {
        if (pressedShortcut === "") return null;
        var shortcutActions = getShortcutActions();
        for (var i = 0; i < shortcutActions.length; i++) {
            if (shortcuts[shortcutActions[i].id] === pressedShortcut) return shortcutActions[i];
        }
        return null;
    }

    /**
     * 管理画面の1行（機能名・修飾キー・キー）を追加する
     * @param {Panel} parent 追加先のパネル
     * @param {object} shortcutAction getShortcutActions() の要素
     * @param {string} shortcut 今の割り当て（未割り当ては ""）
     * @returns {{modifierList: DropDownList, keyList: DropDownList}} 追加した選択欄
     */
    function addShortcutRow(parent, shortcutAction, shortcut) {
        var rowGroup = parent.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        var actionLabel = rowGroup.add("statictext", undefined, shortcutAction.labelText);
        actionLabel.preferredSize.width = SHORTCUT_LABEL_WIDTH;
        actionLabel.justify = "right";

        var modifierList = rowGroup.add("dropdownlist", undefined, [getLabel(LABELS.value.noModifier)].concat(SHORTCUT_MODIFIERS));
        modifierList.helpTip = getLabel(LABELS.tooltip.shortcutModifier);
        var keyList = rowGroup.add("dropdownlist", undefined, SHORTCUT_KEYS);
        keyList.helpTip = getLabel(LABELS.tooltip.shortcutKey);

        /* 「修飾キー+キー」を分ける（キーは最後の + の後）/ the key follows the last "+" */
        var separatorIndex = shortcut.lastIndexOf("+");
        var modifierIndex = -1;
        var keyIndex = -1;
        if (separatorIndex > 0) {
            var modifierText = shortcut.substring(0, separatorIndex);
            var keyText = shortcut.substring(separatorIndex + 1);
            for (var i = 0; i < SHORTCUT_MODIFIERS.length; i++) {
                if (SHORTCUT_MODIFIERS[i] === modifierText) modifierIndex = i;
            }
            for (var j = 0; j < SHORTCUT_KEYS.length; j++) {
                if (SHORTCUT_KEYS[j] === keyText) keyIndex = j;
            }
        }
        var isAssigned = (modifierIndex >= 0 && keyIndex >= 0);
        modifierList.selection = isAssigned ? modifierIndex + 1 : 0;
        keyList.selection = isAssigned ? keyIndex : 0;
        keyList.enabled = isAssigned;
        modifierList.onChange = function () { keyList.enabled = (modifierList.selection.index > 0); };

        return { modifierList: modifierList, keyList: keyList };
    }

    /**
     * キーボードショートカットを設定する管理画面を表示する
     * @param {Object<string, string>} shortcuts 今の割り当て（機能 ID → 「修飾キー+キー」）
     * @returns {Object<string, string>|null} OK なら新しい割り当て、キャンセルなら null
     */
    function showManageDialog(shortcuts) {
        var shortcutActions = getShortcutActions();

        var manageDialog = new Window("dialog", getLabel(LABELS.manageDialog.title));
        manageDialog.orientation = "column";
        manageDialog.alignChildren = "fill";
        manageDialog.margins = PALETTE_MARGINS;
        manageDialog.spacing = PALETTE_SPACING;

        /* パレットと同じ並び（左：増減＋幅に合わせる／右：プリセット）/ Same layout as the palette */
        var columnsGroup = manageDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;
        var leftColumnGroup = columnsGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = ["fill", "top"];
        leftColumnGroup.spacing = COLUMN_SPACING;
        var groupPanels = {};
        groupPanels.step = addColumnPanel(leftColumnGroup, LABELS.field.step);
        groupPanels.fit = addColumnPanel(leftColumnGroup, LABELS.field.fitWidth);
        groupPanels.preset = addColumnPanel(columnsGroup, LABELS.field.preset);

        var shortcutRows = [];
        for (var i = 0; i < shortcutActions.length; i++) {
            shortcutRows.push(addShortcutRow(groupPanels[shortcutActions[i].kind], shortcutActions[i], shortcuts[shortcutActions[i].id] || ""));
        }

        var btnRowGroup = manageDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        var nextShortcuts = null;

        /**
         * 選択欄から割り当てを読む
         * @returns {Object<string, string>} 機能 ID → 「修飾キー+キー」
         */
        function readShortcutRows() {
            var readShortcuts = {};
            for (var k = 0; k < shortcutActions.length; k++) {
                var modifierIndex = shortcutRows[k].modifierList.selection.index;
                readShortcuts[shortcutActions[k].id] = (modifierIndex > 0)
                    ? SHORTCUT_MODIFIERS[modifierIndex - 1] + "+" + shortcutRows[k].keyList.selection.text
                    : "";
            }
            return readShortcuts;
        }

        /* 同じキーの重複があれば閉じずに知らせる / Keep the dialog open on duplicates */
        btnOK.onClick = function () {
            var readShortcuts = readShortcutRows();
            var ownerLabels = {};
            var duplicateLines = [];
            for (var k = 0; k < shortcutActions.length; k++) {
                var shortcut = readShortcuts[shortcutActions[k].id];
                if (shortcut === "") continue;
                var actionName = getLabel(shortcutActions[k].groupLabel) + " " + shortcutActions[k].labelText;
                if (ownerLabels.hasOwnProperty(shortcut)) {
                    duplicateLines.push(shortcut + "\uFF1A" + ownerLabels[shortcut] + " / " + actionName);
                } else {
                    ownerLabels[shortcut] = actionName;
                }
            }
            if (duplicateLines.length > 0) {
                alert(getLabel(LABELS.alert.duplicateShortcut) + "\n" + duplicateLines.join("\n"));
                return;
            }
            nextShortcuts = readShortcuts;
            manageDialog.close(1);
        };

        manageDialog.show();
        return nextShortcuts;
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

        /* 縮尺率：今の値 → ［入力欄］% / Scale: current → [input] % */
        var scaleRowGroup = addLabeledRow(scalePalette, LABELS.field.scale);
        var currentScaleText = scaleRowGroup.add("statictext", undefined, "");
        currentScaleText.helpTip = getLabel(LABELS.tooltip.currentScale);
        scaleRowGroup.add("statictext", undefined, "\u2192");
        var scaleInput = scaleRowGroup.add("edittext", undefined, "");
        scaleInput.characters = SCALE_FIELD_CHARS;
        scaleInput.helpTip = getLabel(LABELS.tooltip.scale);
        scaleRowGroup.add("statictext", undefined, "%");

        var widthInput = addLengthRow(scalePalette, LABELS.field.width, LABELS.tooltip.width);
        var heightInput = addLengthRow(scalePalette, LABELS.field.height, LABELS.tooltip.height);


        /* PPI：元の PPI → ［有効 PPI の入力欄］/ PPI: actual → [effective PPI field] */
        var ppiRowGroup = addLabeledRow(scalePalette, LABELS.field.ppi);
        var actualPpiText = ppiRowGroup.add("statictext", undefined, "");
        actualPpiText.helpTip = getLabel(LABELS.tooltip.actualPpi);
        ppiRowGroup.add("statictext", undefined, "\u2192");
        var ppiInput = ppiRowGroup.add("edittext", undefined, "");
        ppiInput.characters = SCALE_FIELD_CHARS;
        ppiInput.helpTip = getLabel(LABELS.tooltip.effectivePpi);

        var pixelSizeRowGroup = addLabeledRow(scalePalette, LABELS.field.pixelSize);
        pixelSizeRowGroup.margins = [0, 0, 0, INFO_BOTTOM_GAP];
        var pixelSizeText = pixelSizeRowGroup.add("statictext", undefined, "");
        pixelSizeText.characters = INFO_TEXT_CHARS;
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
        var anchorPanel = addColumnPanel(leftColumnGroup, LABELS.field.anchor);
        anchorPanel.margins = ANCHOR_PANEL_MARGINS;
        var presetPanel = addColumnPanel(buttonColumnsGroup, LABELS.field.preset);

        var btnFitMargin = fitPanel.add("button", undefined, getLabel(LABELS.button.fitMargin));
        var btnFitTextFrame = fitPanel.add("button", undefined, getLabel(LABELS.button.fitTextFrame));
        makeCompactButton(btnFitMargin);
        makeCompactButton(btnFitTextFrame);
        /* パネル幅まで伸ばさず、文字に合わせた幅で中央に置く / natural width, centered (not filled) */
        btnFitMargin.alignment = ["center", "top"];
        btnFitTextFrame.alignment = ["center", "top"];
        btnFitMargin.helpTip = getLabel(LABELS.tooltip.fitMargin);
        btnFitTextFrame.helpTip = getLabel(LABELS.tooltip.fitTextFrame);

        /* パレットの一番下（左：丸め／右：管理）/ At the bottom: rounding on the left, Manage on the right */
        var btnRowGroup = scalePalette.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["fill", "bottom"];
        btnRowGroup.alignChildren = ["left", "center"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var roundCheckbox = btnLeftGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.roundToInteger));
        roundCheckbox.value = true;
        roundCheckbox.helpTip = getLabel(LABELS.tooltip.roundToInteger);

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        var btnManage = btnRightGroup.add("button", undefined, getLabel(LABELS.button.manage));
        btnManage.helpTip = getLabel(LABELS.tooltip.manage);
        btnManage.onClick = function () {
            var nextShortcuts = showManageDialog(shortcuts);
            if (nextShortcuts === null) return;
            shortcuts = nextShortcuts;
            saveShortcuts(shortcuts);
        };

        var shortcuts = loadShortcuts();
        var scalePreview = null;
        var isChangingDocument = false;   /* 変更・取り消しの最中は選択の変更イベントを無視する / ignore selection events while applying or undoing */
        var anchorWidget = addAnchorWidget(anchorPanel, function () { refreshPreview(); });
        var activeFitMode = null;   /* 幅に合わせるボタン・幅の欄を使った状態。縮尺率を入れ直すと解除 / set by fit buttons or the Width field, cleared by typing a scale */

        /**
         * 文字列の幅ぴったりに表示し、行を詰め直す（ラベルと矢印の間に空きを作らない）
         * @param {StaticText} textControl 対象の表示
         * @param {string} displayText 表示する文字列
         * @returns {void}
         */
        function setTextFitWidth(textControl, displayText) {
            textControl.text = displayText;
            var measuredSize = textControl.graphics.measureString(displayText);
            var fitSize = [Math.ceil(measuredSize[0]), Math.ceil(measuredSize[1])];
            /* 一度配置されたあとは preferredSize だけでは幅が変わらない（0 のまま残る）ので size も入れる */
            /* After the first layout, preferredSize alone does not resize the control (it stays 0), so set size too */
            textControl.preferredSize = fitSize;
            textControl.size = fitSize;
            scalePalette.layout.layout(true);
        }

        /**
         * 読み込んだ選択状態をパレットの表示に反映する
         * @returns {void}
         */
        function loadSelectionState() {
            var scaleTargets = selectionState.scaleTargets;
            scalePreview = createScalePreview(selectionState.targetDocument, scaleTargets);
            activeFitMode = null;
            setTextFitWidth(currentScaleText, (scaleTargets.length > 0)
                ? describeCurrentScale(selectionState.currentScale)
                : getLabel(LABELS.value.none));
            setTextFitWidth(actualPpiText, describePpi(scaleTargets, "actualPpi"));
            pixelSizeText.text = describePixelSize(scaleTargets);
            scaleInput.text = String(selectionState.initialPercent);
            /* 表示前はチェックボックスを読めないので、丸めずに今の幅を出す / checkbox is unreadable before show; show the current width unrounded */
            var currentSpec = { fitMode: null, scalePercent: selectionState.initialPercent, roundToInteger: false };
            widthInput.text = describeLength(scaleTargets, currentSpec, FIT_MODE.width);
            heightInput.text = describeLength(scaleTargets, currentSpec, FIT_MODE.height);
            ppiInput.text = describeEffectivePpi(scaleTargets, currentSpec);
            btnFitMargin.enabled = canFitAny(scaleTargets, FIT_MODE.margin);
            btnFitTextFrame.enabled = canFitAny(scaleTargets, FIT_MODE.textFrame);
        }

        /**
         * 入力内容から適用内容を読む
         * @returns {object|null} { fitMode, scalePercent, targetLengthPt, targetPpi, roundToInteger, anchorIndex }。数値が無効なら null
         */
        function readScaleSpec() {
            var scaleSpec = {
                fitMode: activeFitMode,
                scalePercent: null,
                targetLengthPt: null,
                targetPpi: null,
                roundToInteger: roundCheckbox.value,
                anchorIndex: anchorWidget.selectedAnchorIndex
            };
            if (activeFitMode === FIT_MODE.width || activeFitMode === FIT_MODE.height) {
                /* 0 より大きい数値を読む処理は縮尺率と共通 / same positive-number check as the scale */
                var lengthMm = readScalePercent(activeFitMode === FIT_MODE.width ? widthInput : heightInput);
                if (lengthMm === null) return null;
                scaleSpec.targetLengthPt = lengthMm * POINTS_PER_MM;
            } else if (activeFitMode === FIT_MODE.ppi) {
                scaleSpec.targetPpi = readScalePercent(ppiInput);
                if (scaleSpec.targetPpi === null) return null;
            } else if (activeFitMode === null) {
                scaleSpec.scalePercent = readScalePercent(scaleInput);
                if (scaleSpec.scalePercent === null) return null;
            }
            return scaleSpec;
        }

        /**
         * 入力内容に合わせて、直前の変更を戻してから掛け直す
         * @returns {void}
         */
        function refreshPreview() {
            isChangingDocument = true;
            scalePreview.clear();
            var scaleSpec = readScaleSpec();
            if (scaleSpec !== null) scalePreview.apply(scaleSpec);
            isChangingDocument = false;
            /* 入力中の欄（幅・高さ・PPI）は書き換えない / leave the field being typed in alone */
            if (activeFitMode !== FIT_MODE.width) widthInput.text = describeLength(selectionState.scaleTargets, scaleSpec, FIT_MODE.width);
            if (activeFitMode !== FIT_MODE.height) heightInput.text = describeLength(selectionState.scaleTargets, scaleSpec, FIT_MODE.height);
            if (activeFitMode !== FIT_MODE.ppi) ppiInput.text = describeEffectivePpi(selectionState.scaleTargets, scaleSpec);
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
         * 縮尺率を増減して掛け直す（欄が空・無効なら今の縮尺率から）
         * @param {number} delta 増減量
         * @returns {void}
         */
        function stepScale(delta) {
            var value = Number(scaleInput.text);
            setScaleValue(stepScaleValue(isNaN(value) ? selectionState.initialPercent : value, delta, roundCheckbox.value));
        }

        /**
         * 幅に合わせる基準を選び、掛け直す
         * @param {string} fitMode FIT_MODE の値
         * @returns {void}
         */
        function selectFitMode(fitMode) {
            activeFitMode = fitMode;
            showFitPercent();
            refreshPreview();
        }

        /**
         * 幅に合わせているときの縮尺率を縮尺率の入力欄に表示する
         * @returns {void}
         */
        function showFitPercent() {
            var scaleSpec = readScaleSpec();
            scaleInput.text = (scaleSpec === null) ? "" : describeFitPercent(selectionState.scaleTargets, scaleSpec);
        }

        /* 幅に合わせている間は、丸め方を変えると欄の値も変わる / In fit mode the shown value follows the rounding */
        roundCheckbox.onClick = function () {
            if (activeFitMode !== null) showFitPercent();
            refreshPreview();
        };

        /**
         * 幅・高さ・PPI の欄の入力から縮尺率を求めて掛け直す
         * @param {string} targetMode FIT_MODE.width／height／ppi
         * @returns {void}
         */
        function applyTargetInput(targetMode) {
            activeFitMode = targetMode;
            showFitPercent();
            refreshPreview();
        }

        /**
         * 幅・高さ・PPI の欄に入力と↑↓キーの処理を付ける
         * @param {EditText} targetInput 対象の欄
         * @param {string} targetMode FIT_MODE.width／height／ppi
         * @returns {void}
         */
        function bindTargetInput(targetInput, targetMode) {
            targetInput.onChanging = function () { applyTargetInput(targetMode); };
            changeValueByArrowKey(targetInput, function (value, delta) {
                var nextValue = Number(formatNumber(value + delta));
                return (nextValue > 0) ? nextValue : value;
            }, function () { applyTargetInput(targetMode); });
        }

        bindTargetInput(widthInput, FIT_MODE.width);
        bindTargetInput(heightInput, FIT_MODE.height);
        bindTargetInput(ppiInput, FIT_MODE.ppi);

        btnFitMargin.onClick = function () { selectFitMode(FIT_MODE.margin); };
        btnFitTextFrame.onClick = function () { selectFitMode(FIT_MODE.textFrame); };

        addValueButtons(stepPanel, SCALE_STEPS,
            function (delta) { return (delta > 0 ? "+" : "\u2212") + Math.abs(delta); },
            getLabel(LABELS.tooltip.step),
            stepScale,
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
        changeValueByArrowKey(scaleInput, function (value, delta) {
            return stepScaleValue(value, delta, roundCheckbox.value);
        }, function () {
            activeFitMode = null;
            refreshPreview();
        });

        /* 管理画面で割り当てたショートカット（欄に「¡」などが入らないよう既定の入力は止める） */
        /* Shortcuts set in the Manage dialog (suppress the typed character such as "¡") */
        scalePalette.addEventListener("keydown", function (event) {
            var shortcutAction = findShortcutAction(shortcuts, getPressedShortcut(event));
            if (shortcutAction === null) return;
            event.preventDefault();
            if (shortcutAction.kind === "step") {
                stepScale(shortcutAction.value);
            } else if (shortcutAction.kind === "fit") {
                /* ボタンが使えない選択では何もしない / skip when the fit button is disabled */
                if (!canFitAny(selectionState.scaleTargets, shortcutAction.value)) return;
                selectFitMode(shortcutAction.value);
            } else {
                setScaleValue(shortcutAction.value);
            }
        });

        /* 選択が変わったら自動で読み込み直す（対象が同じなら何もしない。警告も出さない） */
        /* Follow selection changes (skip when the targets are unchanged; no warnings) */
        var selectionListener = app.addEventListener("afterSelectionChanged", function () {
            if (isChangingDocument) return;
            var nextState = readSelectionState();
            if (getSelectionKey(nextState) === getSelectionKey(selectionState)) return;
            scalePreview.keep();   /* ここまでの変更は残す / keep the changes so far */
            selectionState = nextState;
            loadSelectionState();
        });

        /* 閉じても変更は残す（元に戻すのは Cmd+Z）/ Closing keeps the changes (undo with Cmd+Z) */
        scalePalette.onClose = function () {
            selectionListener.remove();
            $.global.idSetImageScalePalette = null;
        };

        scalePalette.onShow = function () { scaleInput.active = true; };

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
     * 指定の幅にぴったり合う縮尺率を返す（整数への丸めは適用時に行う）
     * @param {Graphic} placedGraphic 対象の画像
     * @param {number|null} targetWidth 収めたい幅
     * @returns {number|null} 縮尺率（%）。求められなければ null
     */
    function getFitPercent(placedGraphic, targetWidth) {
        if (targetWidth === null || targetWidth <= 0) return null;
        var graphicBounds = placedGraphic.geometricBounds;
        var graphicWidth = graphicBounds[3] - graphicBounds[1];
        if (graphicWidth <= 0) return null;
        /* 幅は水平の縮尺率に比例する / width is proportional to horizontalScale */
        return placedGraphic.horizontalScale * targetWidth / graphicWidth;
    }

    /**
     * 縮尺率 100% のときの画像の幅・高さ（pt）を返す（回転していても画像の辺に沿って測る）
     * @param {Graphic} placedGraphic 対象の画像
     * @returns {object} { width, height }（pt）。求められない値は null
     */
    function getNaturalSize(placedGraphic) {
        var horizontalScale = placedGraphic.horizontalScale;
        var verticalScale = placedGraphic.verticalScale;
        var topLeft = resolveGraphicCorner(placedGraphic, AnchorPoint.TOP_LEFT_ANCHOR);
        return {
            width: (horizontalScale > 0)
                ? getDistance(topLeft, resolveGraphicCorner(placedGraphic, AnchorPoint.TOP_RIGHT_ANCHOR)) / horizontalScale * 100
                : null,
            height: (verticalScale > 0)
                ? getDistance(topLeft, resolveGraphicCorner(placedGraphic, AnchorPoint.BOTTOM_LEFT_ANCHOR)) / verticalScale * 100
                : null
        };
    }

    /**
     * 縮尺変更の対象を1件作る（幅に合わせるときの縮尺率は変更前の状態で求めておく）
     * @param {PageItem} graphicFrame 対象のフレーム
     * @param {Graphic} placedGraphic フレーム内の画像
     * @returns {object} { frame, graphic, naturalWidthPt, naturalHeightPt, actualPpi, marginLeft, fitPercents }
     */
    function createScaleTarget(graphicFrame, placedGraphic) {
        var targetPage = graphicFrame.parentPage;
        var marginEdges = targetPage ? getMarginEdges(targetPage) : null;
        var naturalSize = getNaturalSize(placedGraphic);
        return {
            frame: graphicFrame,
            graphic: placedGraphic,
            naturalWidthPt: naturalSize.width,
            naturalHeightPt: naturalSize.height,
            /* PDF は PPI を持たない / PDF has no PPI */
            actualPpi: (placedGraphic.reflect.name === "Image") ? placedGraphic.actualPpi[0] : null,
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
     * 対象ごとに適用する縮尺率を決める（整数に丸めるときは、入力値は四捨五入、幅に合わせるときははみ出さないよう切り捨て）
     * @param {object} scaleTarget createScaleTarget() の戻り値
     * @param {object} scaleSpec { fitMode, scalePercent, targetLengthPt, targetPpi, roundToInteger, anchorIndex }
     * @returns {number|null} 縮尺率（%）。求められなければ null
     */
    function resolveScalePercent(scaleTarget, scaleSpec) {
        var scalePercent;
        var naturalKey = NATURAL_LENGTH_KEY[scaleSpec.fitMode];
        if (naturalKey) {
            var naturalLength = scaleTarget[naturalKey];
            scalePercent = (naturalLength > 0) ? scaleSpec.targetLengthPt / naturalLength * 100 : null;
        } else if (scaleSpec.fitMode === FIT_MODE.ppi) {
            /* 有効 PPI ＝ 元の PPI ÷ 縮尺率。PPI を上げるほど縮尺率は小さくなる / effective PPI = actual PPI / scale */
            scalePercent = (scaleTarget.actualPpi > 0) ? scaleTarget.actualPpi / scaleSpec.targetPpi * 100 : null;
        } else {
            scalePercent = scaleSpec.fitMode ? scaleTarget.fitPercents[scaleSpec.fitMode] : scaleSpec.scalePercent;
        }
        if (scalePercent === null) return null;
        if (scaleSpec.roundToInteger) {
            /* 浮動小数の誤差で 1 下がらないよう少し足す / nudge so float error does not drop a whole percent */
            scalePercent = scaleSpec.fitMode ? Math.floor(scalePercent + 1e-6) : Math.round(scalePercent);
        }
        return (scalePercent > 0) ? scalePercent : null;
    }

    /**
     * 境界の基準点の座標を返す
     * @param {Array} bounds geometricBounds [上, 左, 下, 右]
     * @param {number} anchorIndex 基準点（0..8 を行優先）
     * @returns {Array} [x, y]
     */
    function getAnchorPosition(bounds, anchorIndex) {
        var col = anchorIndex % 3;
        var row = Math.floor(anchorIndex / 3);
        return [
            bounds[1] + (bounds[3] - bounds[1]) * col / 2,
            bounds[0] + (bounds[2] - bounds[0]) * row / 2
        ];
    }

    /**
     * 画像を縮尺変更し、フレームを画像に合わせる（基準点の位置を保つ。マージン幅のときは左端を左マージンへ）
     * @param {Array} scaleTargets collectScaleTargets() の戻り値
     * @param {object} scaleSpec { fitMode, scalePercent, targetLengthPt, targetPpi, roundToInteger, anchorIndex }
     * @returns {void}
     */
    function scaleGraphicsAndFitFrames(scaleTargets, scaleSpec) {
        for (var i = 0; i < scaleTargets.length; i++) {
            var scaleTarget = scaleTargets[i];
            if (!scaleTarget.frame.isValid) continue;   /* パレット表示中に削除された / deleted while the palette is open */
            var scalePercent = resolveScalePercent(scaleTarget, scaleSpec);
            if (scalePercent === null) continue;   /* 幅に合わせられないフレームは変えない / skip frames that cannot fit */

            /* 基準点の位置を控えてから変形し、あとで同じ位置へ戻す（アンカー付きは動かせない）/ keep the reference point in place (not for anchored frames) */
            var canMove = !isAnchoredFrame(scaleTarget.frame);
            var anchorBefore = canMove ? getAnchorPosition(scaleTarget.frame.geometricBounds, scaleSpec.anchorIndex) : null;

            scaleTarget.graphic.horizontalScale = scalePercent;
            scaleTarget.graphic.verticalScale = scalePercent;
            scaleTarget.frame.fit(FitOptions.FRAME_TO_CONTENT);

            if (canMove) {
                var anchorAfter = getAnchorPosition(scaleTarget.frame.geometricBounds, scaleSpec.anchorIndex);
                scaleTarget.frame.move(undefined, [anchorBefore[0] - anchorAfter[0], anchorBefore[1] - anchorAfter[1]]);
            }

            if (scaleSpec.fitMode === FIT_MODE.margin && scaleTarget.marginLeft !== null) {
                var dx = scaleTarget.marginLeft - scaleTarget.frame.geometricBounds[1];
                scaleTarget.frame.move(undefined, [dx, 0]);
            }
        }
    }

    /**
     * 縮尺変更を 1 段の取り消しとして適用する
     * @param {Array} scaleTargets collectScaleTargets() の戻り値
     * @param {object} scaleSpec { fitMode, scalePercent, targetLengthPt, targetPpi, roundToInteger, anchorIndex }
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
        /* 未選択は警告せず、空のまま開く（選択すると読み込む） / No selection: open empty without a warning */
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
            selectionState.initialPercent = Number(formatNumber(currentScale.horizontal));
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
