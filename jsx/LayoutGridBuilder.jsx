#target indesign

    (function () {
        //
        // 簡易レイアウト
        // アクティブページに罫線・フレーム・タイトルエリア・塗り・区切り線などの
        // レイアウト要素をダイアログで設定し、一括作成するスクリプト。
        //
        // バージョン: v0.1.4
        // 作成日: 2026-03-12
        // 更新日: 2026-03-13
        // 日英ローカライズ対応
        // 外側エリアと実コンテンツ領域を分離し、タイトルエリアとカラムエリア（＋アキ）を実コンテンツ領域から差し引くよう調整
        // ローカライズの直書き箇所（自動調整・仮グリッド・表示/残す・表示パネル名）をLABELSへ集約
        // ページのマージンUIを2カラム構成へ変更し、左右連動チェックを削除
        // Autoボタン幅を統一し、列幅の文字数ラベルをLABELSへ集約
        // 単位パネル名を具体化し、単位ラジオボタンを縦並びと詳細表記へ変更
        // 見開き対応のため、自動調整系で使うページ境界をスプレッド座標へ統一

        var SCRIPT_VERSION = "v0.1.3";

        // 言語判定
        function getCurrentLang() {
            return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
        }
        var lang = getCurrentLang();

        // ラベル定義
        var LABELS = {
            dialogTitle: { ja: "簡易レイアウト", en: "Quick Layout" },
            alertOpenDoc: { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document before running this script." },
            panelPage: { ja: "ページ", en: "Page" },
            panelDisplay: { ja: "単位（定規、線、テキスト、組版）", en: "Units (Ruler, Stroke, Text, Typesetting)" },
            panelText: { ja: "基本テキスト", en: "Base Text" },
            panelOuter: { ja: "版面", en: "Content Area" },
            panelTitle: { ja: "タイトルエリア", en: "Title Area" },
            panelMargin: { ja: "マージン", en: "Margins" },
            panelFrame: { ja: "フレーム", en: "Frame" },
            panelGrid: { ja: "実コンテンツ領域", en: "Content Region" },
            panelOffset: { ja: "オフセット", en: "Offset" },
            panelFill: { ja: "塗り", en: "Fill" },
            panelRowCol: { ja: "列・行", en: "Columns / Rows" },
            panelDivider: { ja: "区切り線", en: "Dividers" },
            preview: { ja: "プレビュー", en: "Preview" },
            autoAdjust: { ja: "自動調整", en: "Auto" },
            tempGrid: { ja: "仮グリッド", en: "Temp grid" },
            show: { ja: "表示", en: "Show" },
            keep: { ja: "残す", en: "Keep" },
            charUnit: { ja: "文字", en: "chars" },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },

            outerLine: { ja: "罫線", en: "Border" },
            cornerRadius: { ja: "角丸:", en: "Corner:" },
            extension: { ja: "伸縮:", en: "Extension:" },
            capStyle: { ja: "線端:", en: "Cap:" },
            capNone: { ja: "なし", en: "None" },
            capRound: { ja: "丸型", en: "Round" },
            capProject: { ja: "突出", en: "Projecting" },

            titleFill: { ja: "塗り", en: "Fill" },
            titleStroke: { ja: "罫線", en: "Border" },
            titleLength: { ja: "長さ:", en: "Length:" },
            position: { ja: "位置:", en: "Position:" },
            posTop: { ja: "上", en: "Top" },
            posBottom: { ja: "下", en: "Bottom" },
            posLeft: { ja: "左", en: "Left" },
            posRight: { ja: "右", en: "Right" },

            sideTop: { ja: "天", en: "Top" },
            sideBottom: { ja: "地", en: "Bottom" },
            sideLeft: { ja: "左", en: "Left" },
            sideRight: { ja: "右", en: "Right" },
            link: { ja: "連動", en: "Link" },
            relative: { ja: "相対:", en: "Relative:" },

            frameEnable: { ja: "フレームを描画", en: "Draw frame" },
            frameBleed: { ja: "裁ち落とし", en: "Use bleed" },

            colCount: { ja: "列数:", en: "Cols:" },
            rowCount: { ja: "行数:", en: "Rows:" },
            gap: { ja: "間隔:", en: "Gap:" },
            gapLink: { ja: "間隔を連動", en: "Link gaps" },

            dividerEnable: { ja: "区切り線を描画", en: "Draw dividers" },
            lineSolid: { ja: "実線", en: "Solid" },
            lineDashed: { ja: "点線", en: "Dashed" },
            lineDotted: { ja: "ドット点線", en: "Dotted" },
            lineWeight: { ja: "線幅:", en: "Weight:" },
            enabled: { ja: "有効", en: "Enable" },
            baseFontSize: { ja: "サイズ:", en: "Size:" },
            leading: { ja: "行送り:", en: "Leading:" },
            panelColumnArea: { ja: "フッターの「コラム」エリア", en: "Footer Column Area" },
            columnFill: { ja: "塗り", en: "Fill" },
            columnStroke: { ja: "罫線", en: "Border" },
            columnHeight: { ja: "高さ:", en: "Height:" },
            columnMargin: { ja: "アキ:", en: "Spacing:" },
            textFrame: { ja: "テキストフレーム", en: "Text Frame" },
            threadText: { ja: "スレッドテキスト", en: "Thread Text" },
            sampleTextNone: { ja: "なし", en: "None" },
            sampleTextSample: { ja: "サンプル", en: "Sample" },
            sampleTextSquareCircle: { ja: "ダミー文字", en: "Dummy" }
        };

        function main() {
            // ドキュメントが開かれているか確認
            if (app.documents.length === 0) {
                alert(LABELS.alertOpenDoc[lang]);
                return;
            }
            // ドキュメントの定規単位を取得
            var doc = app.activeDocument;
            var rulerUnit = doc.viewPreferences.horizontalMeasurementUnits;

            // 単位名の取得
            function getUnitName(unit) {
                switch (unit) {
                    case MeasurementUnits.MILLIMETERS: return "mm";
                    case MeasurementUnits.POINTS: return "pt";
                    case MeasurementUnits.INCHES: return "in";
                    case MeasurementUnits.INCHES_DECIMAL: return "in";
                    case MeasurementUnits.CENTIMETERS: return "cm";
                    case MeasurementUnits.PICAS: return "p";
                    case MeasurementUnits.PIXELS: return "px";
                    case MeasurementUnits.CICEROS: return "c";
                    case MeasurementUnits.Q: return "Q";
                    case MeasurementUnits.HA: return "H";
                    default: return "pt";
                }
            }

            function withUnit(labelKey, unit) {
                return LABELS[labelKey][lang] + " (" + unit + ")";
            }

            function getPageBoundsOnSpread(targetPage) {
                var tl = targetPage.resolve(AnchorPoint.TOP_LEFT_ANCHOR, CoordinateSpaces.SPREAD_COORDINATES)[0];
                var br = targetPage.resolve(AnchorPoint.BOTTOM_RIGHT_ANCHOR, CoordinateSpaces.SPREAD_COORDINATES)[0];
                // スプレッド座標はpt単位なので、ドキュメントの定規単位に変換
                return [tl[1] / ptPerUnit, tl[0] / ptPerUnit, br[1] / ptPerUnit, br[0] / ptPerUnit];
            }

            function getPageBleedOffsets(targetPage) {
                var dp = doc.documentPreferences;
                var top = dp.documentBleedTopOffset || 0;
                var bottom = dp.documentBleedBottomOffset || 0;
                var insideOrLeft = dp.documentBleedInsideOrLeftOffset || 0;
                var outsideOrRight = dp.documentBleedOutsideOrRightOffset || 0;
                var side = targetPage.side;

                if (side === PageSideOptions.LEFT_HAND) {
                    return {
                        top: top,
                        bottom: bottom,
                        left: outsideOrRight,
                        right: 0
                    };
                }
                if (side === PageSideOptions.RIGHT_HAND) {
                    return {
                        top: top,
                        bottom: bottom,
                        left: 0,
                        right: outsideOrRight
                    };
                }
                return {
                    top: top,
                    bottom: bottom,
                    left: insideOrLeft,
                    right: outsideOrRight
                };
            }

            var unitName = getUnitName(rulerUnit);

            // 単位からポイントへの変換係数
            function unitToPt(unit) {
                switch (unit) {
                    case MeasurementUnits.MILLIMETERS: return 2.834645669;
                    case MeasurementUnits.POINTS: return 1;
                    case MeasurementUnits.INCHES: return 72;
                    case MeasurementUnits.INCHES_DECIMAL: return 72;
                    case MeasurementUnits.CENTIMETERS: return 28.34645669;
                    case MeasurementUnits.PICAS: return 12;
                    case MeasurementUnits.PIXELS: return 1;
                    case MeasurementUnits.CICEROS: return 12.7878;
                    case MeasurementUnits.Q: return 0.708661417;
                    case MeasurementUnits.HA: return 0.708661417;
                    default: return 1;
                }
            }
            var ptPerUnit = unitToPt(rulerUnit);

            // 現在のページのマージン値を取得
            var page = app.activeWindow.activePage;
            var mp = page.marginPreferences;
            var defMarginTop = mp.top;
            var defMarginBottom = mp.bottom;
            var isLeftPage = (page.side === PageSideOptions.LEFT_HAND);
            var defMarginLeft = isLeftPage ? mp.right : mp.left;
            var defMarginRight = isLeftPage ? mp.left : mp.right;

            function buildDialogUI() {
                // ダイアログの作成
                var dlg = new Window("dialog", LABELS.dialogTitle[lang] + " v" + SCRIPT_VERSION);
                dlg.orientation = "column";
                dlg.alignChildren = "fill";

                // ↑↓キーで値を増減する機能を追加
                function addArrowKeySupport(editField, minValue) {
                    editField.addEventListener("keydown", function (e) {
                        if (e.keyName === "Up" || e.keyName === "Down") {
                            e.preventDefault();
                            var val = parseFloat(editField.text);
                            if (isNaN(val)) val = 0;
                            var step = e.shiftKey ? 10 : 1;
                            if (e.keyName === "Up") {
                                val = Math.ceil((val + 0.001) / step) * step;
                            } else {
                                val = Math.floor((val - 0.001) / step) * step;
                            }
                            if (minValue !== undefined && val < minValue) val = minValue;
                            editField.text = String(val);
                            editField.notify("onChange");
                        }
                    });
                    if (minValue !== undefined) {
                        editField.addEventListener("change", function () {
                            var val = parseFloat(editField.text);
                            if (isNaN(val) || val < minValue) {
                                editField.text = String(minValue);
                            }
                        });
                    }
                }

                // 3カラムレイアウト
                var mainRow = dlg.add("group");
                mainRow.orientation = "row";
                mainRow.alignChildren = ["fill", "top"];

                var leftCol = mainRow.add("group");
                leftCol.orientation = "column";
                leftCol.alignChildren = "fill";

                var centerCol = mainRow.add("group");
                centerCol.orientation = "column";
                centerCol.alignChildren = "fill";

                var rightCol = mainRow.add("group");
                rightCol.orientation = "column";
                rightCol.alignChildren = "fill";

                // 外側エリア panel（中央カラム）
                var pnl1 = centerCol.add("panel", undefined, LABELS.panelOuter[lang]);
                pnl1.orientation = "column";
                pnl1.alignChildren = "left";
                pnl1.margins = [15, 20, 15, 10];

                var outerLineGroup = pnl1.add("group");
                var outerLineCheck = outerLineGroup.add("checkbox", undefined, LABELS.outerLine[lang]);
                outerLineCheck.value = false;
                var strokeWeightInput = outerLineGroup.add("edittext", undefined, "0.3");
                strokeWeightInput.characters = 5;
                addArrowKeySupport(strokeWeightInput);
                var strokeUnitLabel = outerLineGroup.add("statictext", undefined, "pt");

                var row2b = pnl1.add("group");
                row2b.add("statictext", undefined, LABELS.cornerRadius[lang], { justify: "right" });
                var cornerInput = row2b.add("edittext", undefined, "0");
                cornerInput.characters = 4;
                addArrowKeySupport(cornerInput);
                row2b.add("statictext", undefined, unitName);

                var row2 = pnl1.add("group");
                row2.add("statictext", undefined, LABELS.extension[lang], { justify: "right" });
                var extensionInput = row2.add("edittext", undefined, "0");
                extensionInput.characters = 4;
                addArrowKeySupport(extensionInput);
                row2.add("statictext", undefined, unitName);

                // 線端のラジオボタン
                var capGroup = pnl1.add("group");
                capGroup.add("statictext", undefined, LABELS.capStyle[lang], { justify: "right" });
                var rbCapNone = capGroup.add("radiobutton", undefined, LABELS.capNone[lang]);
                var rbCapRound = capGroup.add("radiobutton", undefined, LABELS.capRound[lang]);
                var rbCapProject = capGroup.add("radiobutton", undefined, LABELS.capProject[lang]);
                rbCapNone.value = true;

                // タイトルエリア panel（中央カラム）
                var pnl2 = centerCol.add("panel", undefined, LABELS.panelTitle[lang]);
                pnl2.orientation = "column";
                pnl2.alignChildren = "left";
                pnl2.margins = [15, 20, 15, 10];

                var chkGroup = pnl2.add("group");
                var titleFillCheck = chkGroup.add("checkbox", undefined, LABELS.titleFill[lang]);
                var titleStrokeCheck = chkGroup.add("checkbox", undefined, LABELS.titleStroke[lang]);

                // 長さのデフォルト: 外側エリア高さの1/5
                var pageBoundsOnSpread = getPageBoundsOnSpread(page);
                var defTitleLength = Math.round(((pageBoundsOnSpread[2] - pageBoundsOnSpread[0]) - defMarginTop - defMarginBottom) / 5 * 100) / 100;

                var row3 = pnl2.add("group");
                row3.add("statictext", undefined, LABELS.titleLength[lang], { justify: "right" });
                var titleLengthInput = row3.add("edittext", undefined, String(defTitleLength));
                titleLengthInput.characters = 4;
                addArrowKeySupport(titleLengthInput);
                row3.add("statictext", undefined, unitName);
                var titleLengthAutoBtn = row3.add("button", undefined, LABELS.autoAdjust[lang]);
                titleLengthAutoBtn.preferredSize = [70, 22];

                // 位置のラジオボタン
                var radioGroup = pnl2.add("group");
                radioGroup.add("statictext", undefined, LABELS.position[lang], { justify: "right" });
                var rbTop = radioGroup.add("radiobutton", undefined, LABELS.posTop[lang]);
                var rbBottom = radioGroup.add("radiobutton", undefined, LABELS.posBottom[lang]);
                var rbLeft = radioGroup.add("radiobutton", undefined, LABELS.posLeft[lang]);
                var rbRight = radioGroup.add("radiobutton", undefined, LABELS.posRight[lang]);
                rbTop.value = true; // デフォルトは「上」

                var row4 = pnl2.add("group");
                row4.add("statictext", undefined, LABELS.extension[lang], { justify: "right" });
                var titleExtensionInput = row4.add("edittext", undefined, "0");
                titleExtensionInput.characters = 4;
                addArrowKeySupport(titleExtensionInput);
                row4.add("statictext", undefined, unitName);

                // フッターの「コラム」エリア panel（中央カラム）
                var pnlColumn = centerCol.add("panel", undefined, LABELS.panelColumnArea[lang]);
                pnlColumn.orientation = "column";
                pnlColumn.alignChildren = "left";
                pnlColumn.margins = [15, 20, 15, 10];

                var colAreaChkGroup = pnlColumn.add("group");
                var columnFillCheck = colAreaChkGroup.add("checkbox", undefined, LABELS.columnFill[lang]);
                var columnStrokeCheck = colAreaChkGroup.add("checkbox", undefined, LABELS.columnStroke[lang]);
                var columnWeightInput = colAreaChkGroup.add("edittext", undefined, "0.3");
                columnWeightInput.characters = 5;
                addArrowKeySupport(columnWeightInput);
                var columnWeightUnitLabel = colAreaChkGroup.add("statictext", undefined, "pt");

                var rowColHeight = pnlColumn.add("group");
                rowColHeight.add("statictext", undefined, LABELS.columnHeight[lang], { justify: "right" });
                var columnHeightInput = rowColHeight.add("edittext", undefined, "30");
                columnHeightInput.characters = 4;
                addArrowKeySupport(columnHeightInput);
                rowColHeight.add("statictext", undefined, unitName);
                var columnHeightAutoBtn = rowColHeight.add("button", undefined, LABELS.autoAdjust[lang]);
                columnHeightAutoBtn.preferredSize = [70, 22];

                var rowColMargin = pnlColumn.add("group");
                rowColMargin.add("statictext", undefined, LABELS.columnMargin[lang], { justify: "right" });
                var columnMarginInput = rowColMargin.add("edittext", undefined, "0");
                columnMarginInput.characters = 4;
                addArrowKeySupport(columnMarginInput, 0);
                rowColMargin.add("statictext", undefined, unitName);

                var rowColCorner = pnlColumn.add("group");
                rowColCorner.add("statictext", undefined, LABELS.cornerRadius[lang], { justify: "right" });
                var columnCornerInput = rowColCorner.add("edittext", undefined, "0");
                columnCornerInput.characters = 4;
                addArrowKeySupport(columnCornerInput);
                rowColCorner.add("statictext", undefined, unitName);

                // ページ panel（左カラム）
                var pnlPage = leftCol.add("panel", undefined, LABELS.panelPage[lang]);
                pnlPage.orientation = "column";
                pnlPage.alignChildren = "fill";
                pnlPage.margins = [15, 20, 15, 10];

                // 単位 panel（ページ内）
                var pnlUnit = pnlPage.add("panel", undefined, LABELS.panelDisplay[lang]);
                pnlUnit.orientation = "column";
                pnlUnit.alignChildren = "left";
                pnlUnit.margins = [15, 20, 15, 10];
                var rbUnitMmPt = pnlUnit.add("radiobutton", undefined, "mm/pt/pt/pt");
                var rbUnitMmMmPt = pnlUnit.add("radiobutton", undefined, "mm/mm/pt/pt");
                var rbUnitMmQ = pnlUnit.add("radiobutton", undefined, "mm/mm/Q/H");
                rbUnitMmPt.value = true;

                // 基本テキスト panel（ページ内）
                var pnlText = pnlPage.add("panel", undefined, LABELS.panelText[lang]);
                pnlText.orientation = "column";
                pnlText.alignChildren = "left";
                pnlText.margins = [15, 20, 15, 10];

                var fontGroup = pnlText.add("group");
                fontGroup.orientation = "row";
                fontGroup.alignChildren = ["left", "center"];
                fontGroup.add("statictext", undefined, LABELS.baseFontSize[lang]);
                var baseFontSizeInput = fontGroup.add("edittext", undefined, "9.5");
                baseFontSizeInput.characters = 5;
                addArrowKeySupport(baseFontSizeInput);
                var fontUnitLabel = fontGroup.add("statictext", undefined, "pt");

                var leadingGroup = pnlText.add("group");
                leadingGroup.orientation = "row";
                leadingGroup.alignChildren = ["left", "center"];
                leadingGroup.add("statictext", undefined, LABELS.leading[lang]);
                var leadingInput = leadingGroup.add("edittext", undefined, "16");
                leadingInput.characters = 5;
                addArrowKeySupport(leadingInput);
                var leadingUnitLabel = leadingGroup.add("statictext", undefined, "pt");

                var gridGroup = pnlText.add("group");
                gridGroup.orientation = "row";
                gridGroup.add("statictext", undefined, LABELS.tempGrid[lang]);
                var showLayoutGridCheck = gridGroup.add("checkbox", undefined, LABELS.show[lang]);
                showLayoutGridCheck.value = true;
                var keepGridCheck = gridGroup.add("checkbox", undefined, LABELS.keep[lang]);
                keepGridCheck.value = true;

                // マージン panel（ページ内）
                var pnlMargin = pnlPage.add("panel", undefined, withUnit("panelMargin", unitName));
                pnlMargin.orientation = "column";
                pnlMargin.alignChildren = "center";
                pnlMargin.margins = [15, 20, 15, 10];

                var marginGroup = pnlMargin.add("group");
                marginGroup.orientation = "row";
                marginGroup.alignChildren = ["left", "top"];
                marginGroup.alignment = "center";
                marginGroup.spacing = 16;

                var marginColLeft = marginGroup.add("group");
                marginColLeft.orientation = "column";
                marginColLeft.alignChildren = ["left", "center"];

                var rowMT = marginColLeft.add("group");
                rowMT.add("statictext", undefined, LABELS.sideTop[lang], { justify: "right" });
                var marginTopInput = rowMT.add("edittext", undefined, String(defMarginTop));
                marginTopInput.characters = 4;
                addArrowKeySupport(marginTopInput);

                var rowMB = marginColLeft.add("group");
                rowMB.add("statictext", undefined, LABELS.sideBottom[lang], { justify: "right" });
                var marginBottomInput = rowMB.add("edittext", undefined, String(defMarginBottom));
                marginBottomInput.characters = 4;
                addArrowKeySupport(marginBottomInput);

                var marginColRight = marginGroup.add("group");
                marginColRight.orientation = "column";
                marginColRight.alignChildren = ["left", "center"];

                var rowML = marginColRight.add("group");
                rowML.add("statictext", undefined, LABELS.sideLeft[lang], { justify: "right" });
                var marginLeftInput = rowML.add("edittext", undefined, String(defMarginLeft));
                marginLeftInput.characters = 4;
                addArrowKeySupport(marginLeftInput);

                var rowMR = marginColRight.add("group");
                rowMR.add("statictext", undefined, LABELS.sideRight[lang], { justify: "right" });
                var marginRightInput = rowMR.add("edittext", undefined, String(defMarginRight));
                marginRightInput.characters = 4;
                addArrowKeySupport(marginRightInput);

                var relativeGroup = pnlMargin.add("group");
                relativeGroup.margins = [0, 10, 0, 0];
                relativeGroup.add("statictext", undefined, LABELS.relative[lang], { justify: "right" });
                var relativeInput = relativeGroup.add("edittext", undefined, "0");
                relativeInput.characters = 4;
                addArrowKeySupport(relativeInput);
                relativeGroup.add("statictext", undefined, unitName);

                // フレーム panel（ページ内）
                var pnl3 = pnlPage.add("panel", undefined, withUnit("panelFrame", unitName));
                pnl3.orientation = "column";
                pnl3.alignChildren = "left";
                pnl3.margins = [15, 20, 15, 10];

                var frameChkGroup = pnl3.add("group");
                var frameEnableCheck = frameChkGroup.add("checkbox", undefined, LABELS.frameEnable[lang]);
                frameEnableCheck.value = false;
                var frameBleedCheck = frameChkGroup.add("checkbox", undefined, LABELS.frameBleed[lang]);

                // 幅（3列レイアウト: 左=左、中央=天地+連動、右=右）
                var frameMarginGroup = pnl3.add("group");
                frameMarginGroup.orientation = "row";
                frameMarginGroup.alignment = "center";

                // 1列目：左
                var fColLeft = frameMarginGroup.add("group");
                fColLeft.alignment = "center";
                fColLeft.add("statictext", undefined, LABELS.sideLeft[lang], { justify: "right" });
                var frameLeftInput = fColLeft.add("edittext", undefined, "0");
                frameLeftInput.characters = 4;
                addArrowKeySupport(frameLeftInput);

                // 2列目：天・地 + 連動
                var fColCenter = frameMarginGroup.add("group");
                fColCenter.orientation = "column";
                fColCenter.alignChildren = "center";
                var rowFT = fColCenter.add("group");
                rowFT.add("statictext", undefined, LABELS.sideTop[lang], { justify: "right" });
                var frameTopInput = rowFT.add("edittext", undefined, "0");
                frameTopInput.characters = 4;
                addArrowKeySupport(frameTopInput);
                var frameLinkCheck = fColCenter.add("checkbox", undefined, LABELS.link[lang]);
                frameLinkCheck.value = true;
                var rowFB = fColCenter.add("group");
                rowFB.add("statictext", undefined, LABELS.sideBottom[lang], { justify: "right" });
                var frameBottomInput = rowFB.add("edittext", undefined, "0");
                frameBottomInput.characters = 4;
                addArrowKeySupport(frameBottomInput);

                // 3列目：右
                var fColRight = frameMarginGroup.add("group");
                fColRight.alignment = "center";
                fColRight.add("statictext", undefined, LABELS.sideRight[lang], { justify: "right" });
                var frameRightInput = fColRight.add("edittext", undefined, "0");
                frameRightInput.characters = 4;
                addArrowKeySupport(frameRightInput);

                // グリッド panel（右カラム）
                var pnl4 = rightCol.add("panel", undefined, LABELS.panelGrid[lang]);
                pnl4.orientation = "column";
                pnl4.alignChildren = "fill";
                pnl4.margins = [15, 20, 15, 10];

                // オフセット sub-panel
                var pnlOffset = pnl4.add("panel", undefined, withUnit("panelOffset", unitName));
                pnlOffset.orientation = "column";
                pnlOffset.alignChildren = "center";
                pnlOffset.margins = [15, 20, 15, 10];

                var innerMarginGroup = pnlOffset.add("group");
                innerMarginGroup.orientation = "row";
                innerMarginGroup.alignment = "center";

                // 1列目：左
                var iColLeft = innerMarginGroup.add("group");
                iColLeft.alignment = "center";
                iColLeft.add("statictext", undefined, LABELS.sideLeft[lang], { justify: "right" });
                var innerLeftInput = iColLeft.add("edittext", undefined, "10");
                innerLeftInput.characters = 6;
                addArrowKeySupport(innerLeftInput);

                // 2列目：天・地 + 連動
                var iColCenter = innerMarginGroup.add("group");
                iColCenter.orientation = "column";
                iColCenter.alignChildren = "center";
                var rowIT = iColCenter.add("group");
                rowIT.add("statictext", undefined, LABELS.sideTop[lang], { justify: "right" });
                var innerTopInput = rowIT.add("edittext", undefined, "10");
                innerTopInput.characters = 6;
                addArrowKeySupport(innerTopInput);
                var innerLinkCheck = iColCenter.add("checkbox", undefined, LABELS.link[lang]);
                innerLinkCheck.value = true;
                var rowIB = iColCenter.add("group");
                rowIB.add("statictext", undefined, LABELS.sideBottom[lang], { justify: "right" });
                var innerBottomInput = rowIB.add("edittext", undefined, "10");
                innerBottomInput.characters = 6;
                addArrowKeySupport(innerBottomInput);

                // 3列目：右
                var iColRight = innerMarginGroup.add("group");
                iColRight.alignment = "center";
                iColRight.add("statictext", undefined, LABELS.sideRight[lang], { justify: "right" });
                var innerRightInput = iColRight.add("edittext", undefined, "10");
                innerRightInput.characters = 6;
                addArrowKeySupport(innerRightInput);

                var offsetAutoBtn = pnlOffset.add("button", undefined, LABELS.autoAdjust[lang]);
                offsetAutoBtn.preferredSize = [70, 22];

                // 列・行 sub-panel（グリッドpanel内）
                var pnlRowCol = pnl4.add("panel", undefined, LABELS.panelRowCol[lang]);
                pnlRowCol.orientation = "column";
                pnlRowCol.alignChildren = "left";
                pnlRowCol.margins = [15, 20, 15, 10];

                var rowColCount = pnlRowCol.add("group");
                rowColCount.add("statictext", undefined, LABELS.colCount[lang], { justify: "right" });
                var colCountInput = rowColCount.add("edittext", undefined, "2");
                colCountInput.characters = 5;
                addArrowKeySupport(colCountInput, 1);
                var colCharInput = rowColCount.add("edittext", undefined, "0");
                colCharInput.characters = 4;
                addArrowKeySupport(colCharInput, 1);
                var colCharUnit = rowColCount.add("statictext", undefined, LABELS.charUnit[lang]);

                var rowColGap = pnlRowCol.add("group");
                rowColGap.add("statictext", undefined, LABELS.gap[lang], { justify: "right" });
                var colGapInput = rowColGap.add("edittext", undefined, "10");
                colGapInput.characters = 5;
                addArrowKeySupport(colGapInput, 0);
                rowColGap.add("statictext", undefined, unitName);
                var colGapAutoBtn = rowColGap.add("button", undefined, LABELS.autoAdjust[lang]);
                colGapAutoBtn.preferredSize = [70, 22];

                var rowRowCount = pnlRowCol.add("group");
                rowRowCount.add("statictext", undefined, LABELS.rowCount[lang], { justify: "right" });
                var rowCountInput = rowRowCount.add("edittext", undefined, "1");
                rowCountInput.characters = 5;
                addArrowKeySupport(rowCountInput, 1);

                var rowRowGap = pnlRowCol.add("group");
                rowRowGap.add("statictext", undefined, LABELS.gap[lang], { justify: "right" });
                var rowGapInput = rowRowGap.add("edittext", undefined, "0");
                rowGapInput.characters = 5;
                addArrowKeySupport(rowGapInput, 0);
                rowRowGap.add("statictext", undefined, unitName);
                var gapLinkCheck = rowRowGap.add("checkbox", undefined, LABELS.link[lang]);
                gapLinkCheck.value = true;

                // 塗り sub-panel
                var pnlFill = pnl4.add("panel", undefined, LABELS.panelFill[lang]);
                pnlFill.orientation = "column";
                pnlFill.alignChildren = "left";
                pnlFill.margins = [15, 20, 15, 10];
                var fillGroup = pnlFill.add("group");
                fillGroup.alignment = "left";
                var rbFillColor = fillGroup.add("radiobutton", undefined, LABELS.panelFill[lang]);
                var rbFillTextFrame = fillGroup.add("radiobutton", undefined, LABELS.textFrame[lang]);
                rbFillColor.value = true;
                var threadCheck = pnlFill.add("checkbox", undefined, LABELS.threadText[lang]);
                var sampleTextGroup = pnlFill.add("group");
                sampleTextGroup.alignment = "left";
                var rbSampleNone = sampleTextGroup.add("radiobutton", undefined, LABELS.sampleTextNone[lang]);
                var rbSampleText = sampleTextGroup.add("radiobutton", undefined, LABELS.sampleTextSample[lang]);
                var rbSampleSquareCircle = sampleTextGroup.add("radiobutton", undefined, LABELS.sampleTextSquareCircle[lang]);
                rbSampleNone.value = true;

                // 区切り線 sub-panel（グリッドpanel内）
                var pnlDivider = pnl4.add("panel", undefined, LABELS.panelDivider[lang]);
                pnlDivider.orientation = "column";
                pnlDivider.alignChildren = "left";
                pnlDivider.margins = [15, 20, 15, 10];

                var divHeaderGroup = pnlDivider.add("group");
                var innerStrokeCheck = divHeaderGroup.add("checkbox", undefined, LABELS.dividerEnable[lang]);
                innerStrokeCheck.value = true;
                var divWeightInput = divHeaderGroup.add("edittext", undefined, "0.3");
                divWeightInput.characters = 5;
                addArrowKeySupport(divWeightInput);
                var divWeightUnitLabel = divHeaderGroup.add("statictext", undefined, "pt");

                var lineTypeGroup = pnlDivider.add("group");
                var rbLineSolid = lineTypeGroup.add("radiobutton", undefined, LABELS.lineSolid[lang]);
                var rbLineDashed = lineTypeGroup.add("radiobutton", undefined, LABELS.lineDashed[lang]);
                var rbLineDotted = lineTypeGroup.add("radiobutton", undefined, LABELS.lineDotted[lang]);
                rbLineSolid.value = true;

                // ボタンエリア（左:プレビュー、中央:スペーサー、右:ボタン）
                var btnGroup = dlg.add("group");
                btnGroup.alignment = "fill";
                btnGroup.alignChildren = ["left", "center"];
                var previewCheck = btnGroup.add("checkbox", undefined, LABELS.preview[lang]);
                previewCheck.value = true;
                var allAutoBtn = btnGroup.add("button", undefined, LABELS.autoAdjust[lang]);
                allAutoBtn.preferredSize = [70, 22];
                var btnSpacer = btnGroup.add("group");
                btnSpacer.alignment = ["fill", "center"];
                btnSpacer.preferredSize.width = -1;
                var btnRight = btnGroup.add("group");
                btnRight.alignment = ["right", "center"];
                btnRight.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });
                btnRight.add("button", undefined, LABELS.ok[lang], { name: "ok" });

                return {
                    dlg: dlg,
                    rbUnitMmPt: rbUnitMmPt,
                    rbUnitMmMmPt: rbUnitMmMmPt,
                    rbUnitMmQ: rbUnitMmQ,
                    baseFontSizeInput: baseFontSizeInput,
                    leadingInput: leadingInput,
                    fontUnitLabel: fontUnitLabel,
                    leadingUnitLabel: leadingUnitLabel,
                    showLayoutGridCheck: showLayoutGridCheck,
                    keepGridCheck: keepGridCheck,
                    previewCheck: previewCheck,
                    allAutoBtn: allAutoBtn,
                    outerLineCheck: outerLineCheck,
                    strokeWeightInput: strokeWeightInput,
                    strokeUnitLabel: strokeUnitLabel,
                    columnWeightUnitLabel: columnWeightUnitLabel,
                    divWeightUnitLabel: divWeightUnitLabel,
                    marginTopInput: marginTopInput,
                    marginBottomInput: marginBottomInput,
                    marginLeftInput: marginLeftInput,
                    marginRightInput: marginRightInput,
                    extensionInput: extensionInput,
                    cornerInput: cornerInput,
                    titleLengthInput: titleLengthInput,
                    titleLengthAutoBtn: titleLengthAutoBtn,
                    titleExtensionInput: titleExtensionInput,
                    rbTop: rbTop,
                    rbBottom: rbBottom,
                    rbLeft: rbLeft,
                    rbRight: rbRight,
                    titleFillCheck: titleFillCheck,
                    titleStrokeCheck: titleStrokeCheck,
                    rbCapNone: rbCapNone,
                    rbCapRound: rbCapRound,
                    rbCapProject: rbCapProject,
                    relativeInput: relativeInput,
                    frameEnableCheck: frameEnableCheck,
                    frameTopInput: frameTopInput,
                    frameBottomInput: frameBottomInput,
                    frameLeftInput: frameLeftInput,
                    frameRightInput: frameRightInput,
                    frameLinkCheck: frameLinkCheck,
                    frameBleedCheck: frameBleedCheck,
                    columnFillCheck: columnFillCheck,
                    columnStrokeCheck: columnStrokeCheck,
                    columnHeightInput: columnHeightInput,
                    columnHeightAutoBtn: columnHeightAutoBtn,
                    columnMarginInput: columnMarginInput,
                    columnCornerInput: columnCornerInput,
                    columnWeightInput: columnWeightInput,
                    innerTopInput: innerTopInput,
                    innerBottomInput: innerBottomInput,
                    innerLeftInput: innerLeftInput,
                    innerRightInput: innerRightInput,
                    innerLinkCheck: innerLinkCheck,
                    offsetAutoBtn: offsetAutoBtn,
                    rbFillColor: rbFillColor,
                    rbFillTextFrame: rbFillTextFrame,
                    threadCheck: threadCheck,
                    rbSampleNone: rbSampleNone,
                    rbSampleText: rbSampleText,
                    rbSampleSquareCircle: rbSampleSquareCircle,
                    colCountInput: colCountInput,
                    colCharInput: colCharInput,
                    colGapInput: colGapInput,
                    colGapAutoBtn: colGapAutoBtn,
                    rowCountInput: rowCountInput,
                    rowGapInput: rowGapInput,
                    gapLinkCheck: gapLinkCheck,
                    innerStrokeCheck: innerStrokeCheck,
                    rbLineSolid: rbLineSolid,
                    rbLineDashed: rbLineDashed,
                    rbLineDotted: rbLineDotted,
                    divWeightInput: divWeightInput,
                };
            }

            var ui = buildDialogUI();
            var previewItems = [];
            var PREVIEW_LAYER_NAME = "__QuickLayoutPreview__";
            var previewLayer = null;
            function getOrCreatePreviewLayer() {
                if (previewLayer && previewLayer.isValid) return previewLayer;

                try {
                    previewLayer = doc.layers.itemByName(PREVIEW_LAYER_NAME);
                    previewLayer.name;
                } catch (e) {
                    previewLayer = doc.layers.add({ name: PREVIEW_LAYER_NAME });
                }

                try { previewLayer.visible = true; } catch (e1) { }
                try { previewLayer.locked = false; } catch (e2) { }
                try { previewLayer.printable = false; } catch (e3) { }

                return previewLayer;
            }

            function clearPreviewLayer() {
                var lyr = previewLayer;
                if (!lyr || !lyr.isValid) return;

                try { lyr.locked = false; } catch (e1) { }
                try { lyr.visible = true; } catch (e2) { }

                for (var i = lyr.pageItems.length - 1; i >= 0; i--) {
                    try { lyr.pageItems[i].remove(); } catch (e) { }
                }
            }

            function destroyPreviewLayer() {
                var lyr = previewLayer;
                if (!lyr || !lyr.isValid) {
                    previewLayer = null;
                    return;
                }

                clearPreviewLayer();

                try {
                    if (doc.activeLayer === lyr) {
                        for (var li = 0; li < doc.layers.length; li++) {
                            if (doc.layers[li] !== lyr) {
                                doc.activeLayer = doc.layers[li];
                                break;
                            }
                        }
                    }
                } catch (e4) { }

                try { lyr.remove(); } catch (e5) { }
                previewLayer = null;
            }

            // 選択されている線端を取得
            function getSelectedCap(ui) {
                if (ui.rbCapRound.value) return "round";
                if (ui.rbCapProject.value) return "project";
                return "none";
            }

            // 選択されている位置を取得
            function getSelectedPosition(ui) {
                if (ui.rbTop.value) return "top";
                if (ui.rbBottom.value) return "bottom";
                if (ui.rbLeft.value) return "left";
                if (ui.rbRight.value) return "right";
                return "top";
            }

            // 選択されている罫線の種類を取得
            function getSelectedLineType(ui) {
                if (ui.rbLineDashed.value) return "dashed";
                if (ui.rbLineDotted.value) return "dotted";
                return "solid";
            }

            // すべてのUI値をまとめて取得するヘルパー
            function getCurrentUIValues(ui) {
                var values = {};

                values.outerLine = ui.outerLineCheck.value;
                values.marginTop = parseFloat(ui.marginTopInput.text);
                values.marginBottom = parseFloat(ui.marginBottomInput.text);
                values.marginLeft = parseFloat(ui.marginLeftInput.text);
                values.marginRight = parseFloat(ui.marginRightInput.text);
                values.extension = parseFloat(ui.extensionInput.text);
                values.cornerRadius = parseFloat(ui.cornerInput.text);
                values.titleLength = parseFloat(ui.titleLengthInput.text);
                values.titleExtension = parseFloat(ui.titleExtensionInput.text);
                values.titleCornerRadius = values.cornerRadius;
                values.titlePosition = getSelectedPosition(ui);
                values.titleFill = ui.titleFillCheck.value;
                values.titleStroke = ui.titleStrokeCheck.value;
                values.columnFill = ui.columnFillCheck.value;
                values.columnStroke = ui.columnStrokeCheck.value;
                values.columnHeight = parseFloat(ui.columnHeightInput.text);
                values.columnMargin = parseFloat(ui.columnMarginInput.text);
                values.columnCornerRadius = parseFloat(ui.columnCornerInput.text);
                var isMmStroke = currentStrokeUnitIsMm;
                values.columnWeight = parseFloat(ui.columnWeightInput.text);
                if (isMmStroke) values.columnWeight = values.columnWeight / PT_TO_MM;
                values.capStyle = getSelectedCap(ui);
                values.lineWeight = parseFloat(ui.strokeWeightInput.text);
                if (isMmStroke) values.lineWeight = values.lineWeight / PT_TO_MM;
                values.frameEnable = ui.frameEnableCheck.value;
                values.frameTop = parseFloat(ui.frameTopInput.text);
                values.frameBottom = parseFloat(ui.frameBottomInput.text);
                values.frameLeft = parseFloat(ui.frameLeftInput.text);
                values.frameRight = parseFloat(ui.frameRightInput.text);
                values.frameBleed = ui.frameBleedCheck.value;
                values.innerFill = ui.rbFillColor.value;
                values.innerTextFrame = ui.rbFillTextFrame.value;
                values.threadText = ui.threadCheck.value;
                values.sampleText = ui.rbSampleText.value;
                values.sampleSquareCircle = ui.rbSampleSquareCircle.value;
                var isQ = ui.rbUnitMmQ.value;
                var fontVal = parseFloat(ui.baseFontSizeInput.text);
                values.baseFontSize = isQ ? fontVal / PT_TO_Q : fontVal;
                var leadVal = ui.leadingInput.text;
                var leadNum = parseFloat(leadVal);
                values.leading = (!isNaN(leadNum) && leadNum > 0 && isQ) ? String(leadNum / PT_TO_Q) : leadVal;
                values.innerTop = parseFloat(ui.innerTopInput.text);
                values.innerBottom = parseFloat(ui.innerBottomInput.text);
                values.innerLeft = parseFloat(ui.innerLeftInput.text);
                values.innerRight = parseFloat(ui.innerRightInput.text);
                values.colCount = parseInt(ui.colCountInput.text, 10);
                values.colGap = parseFloat(ui.colGapInput.text);
                values.rowCount = parseInt(ui.rowCountInput.text, 10);
                values.rowGap = parseFloat(ui.rowGapInput.text);
                values.innerStroke = ui.innerStrokeCheck.value;
                values.lineType = getSelectedLineType(ui);
                values.divWeight = parseFloat(ui.divWeightInput.text);
                if (isMmStroke) values.divWeight = values.divWeight / PT_TO_MM;
                values.showLayoutGrid = ui.showLayoutGridCheck.value;
                values.ptPerUnit = ptPerUnit;

                if (isNaN(values.marginTop)) values.marginTop = 0;
                if (isNaN(values.marginBottom)) values.marginBottom = values.marginTop;
                if (isNaN(values.marginLeft)) values.marginLeft = values.marginTop;
                if (isNaN(values.marginRight)) values.marginRight = values.marginTop;
                if (isNaN(values.extension)) values.extension = 0;
                if (isNaN(values.cornerRadius)) values.cornerRadius = 0;
                if (isNaN(values.titleLength)) values.titleLength = 0;
                if (isNaN(values.titleExtension)) values.titleExtension = 0;
                if (isNaN(values.columnHeight)) values.columnHeight = 0;
                if (isNaN(values.columnMargin)) values.columnMargin = 0;
                if (isNaN(values.columnCornerRadius)) values.columnCornerRadius = 0;
                if (isNaN(values.columnWeight)) values.columnWeight = 0.3;
                if (isNaN(values.lineWeight)) values.lineWeight = 0.3;
                if (isNaN(values.frameTop)) values.frameTop = 0;
                if (isNaN(values.frameBottom)) values.frameBottom = 0;
                if (isNaN(values.frameLeft)) values.frameLeft = 0;
                if (isNaN(values.frameRight)) values.frameRight = 0;
                if (isNaN(values.innerTop)) values.innerTop = 0;
                if (isNaN(values.innerBottom)) values.innerBottom = 0;
                if (isNaN(values.innerLeft)) values.innerLeft = 0;
                if (isNaN(values.innerRight)) values.innerRight = 0;
                if (isNaN(values.colCount) || values.colCount < 1) values.colCount = 1;
                if (isNaN(values.colGap)) values.colGap = 0;
                if (isNaN(values.rowCount) || values.rowCount < 1) values.rowCount = 1;
                if (isNaN(values.rowGap)) values.rowGap = 0;
                if (isNaN(values.divWeight)) values.divWeight = 0.3;

                return values;
            }

            // プレビューの作成・削除
            function removePreview() {
                clearPreviewLayer();
                previewItems = [];
            }

            function updatePreview() {
                removePreview();
                if (!ui.previewCheck.value) return;

                var previewValues = getCurrentUIValues(ui);
                previewValues.targetLayer = getOrCreatePreviewLayer();
                previewItems = createLines(previewValues);
            }



            function syncFrameWidths(ui, source) {
                if (ui.frameLinkCheck.value) {
                    ui.frameTopInput.text = source.text;
                    ui.frameBottomInput.text = source.text;
                    ui.frameLeftInput.text = source.text;
                    ui.frameRightInput.text = source.text;
                }
            }

            function syncInnerMargins(ui, source) {
                if (ui.innerLinkCheck.value) {
                    ui.innerTopInput.text = source.text;
                    ui.innerBottomInput.text = source.text;
                    ui.innerLeftInput.text = source.text;
                    ui.innerRightInput.text = source.text;
                }
            }

            function updateFillEnabled(ui) {
                var isTextFrame = ui.rbFillTextFrame.value;
                ui.threadCheck.enabled = isTextFrame;
                ui.rbSampleNone.enabled = isTextFrame;
                ui.rbSampleText.enabled = isTextFrame;
                ui.rbSampleSquareCircle.enabled = isTextFrame;
                if (isTextFrame) {
                    ui.threadCheck.value = true;
                    ui.rbSampleNone.value = false;
                    ui.rbSampleText.value = true;
                    ui.rbSampleSquareCircle.value = false;
                }
            }

            function updateCharCount(ui) {
                try {
                    var bounds = getPageBoundsOnSpread(page);
                    var pgLeft = bounds[1];
                    var pgRight = bounds[3];

                    var mTop = parseFloat(ui.marginTopInput.text) || 0;
                    var mBottom = parseFloat(ui.marginBottomInput.text) || 0;
                    var mLeft = parseFloat(ui.marginLeftInput.text) || 0;
                    var mRight = parseFloat(ui.marginRightInput.text) || 0;

                    var outerTop = bounds[0] + mTop;
                    var outerBottom = bounds[2] - mBottom;
                    var outerLeft = pgLeft + mLeft;
                    var outerRight = pgRight - mRight;

                    var contentTop = outerTop;
                    var contentBottom = outerBottom;
                    var contentLeft = outerLeft;
                    var contentRight = outerRight;

                    // タイトルエリアを実コンテンツ領域から差し引く
                    var titleOn = ui.titleFillCheck.value || ui.titleStrokeCheck.value;
                    var titleLen = parseFloat(ui.titleLengthInput.text) || 0;
                    if (titleOn && titleLen > 0) {
                        var pos = getSelectedPosition(ui);
                        if (pos === "top") contentTop += titleLen;
                        else if (pos === "bottom") contentBottom -= titleLen;
                        else if (pos === "left") contentLeft += titleLen;
                        else if (pos === "right") contentRight -= titleLen;
                    }

                    // フッターの「コラム」エリア（＋アキ）を実コンテンツ領域から差し引く
                    var columnOn = ui.columnFillCheck.value || ui.columnStrokeCheck.value;
                    var columnHeight = parseFloat(ui.columnHeightInput.text) || 0;
                    var columnMargin = parseFloat(ui.columnMarginInput.text) || 0;
                    if (columnOn && columnHeight > 0) {
                        contentBottom -= (columnHeight + columnMargin);
                    }

                    // オフセット
                    var innerT = parseFloat(ui.innerTopInput.text) || 0;
                    var innerB = parseFloat(ui.innerBottomInput.text) || 0;
                    var innerL = parseFloat(ui.innerLeftInput.text) || 0;
                    var innerR = parseFloat(ui.innerRightInput.text) || 0;

                    var gridTop = contentTop + innerT;
                    var gridBottom = contentBottom - innerB;
                    var gridLeft = contentLeft + innerL;
                    var gridRight = contentRight - innerR;

                    var totalWidth = gridRight - gridLeft;
                    var totalHeight = gridBottom - gridTop;

                    if (totalWidth <= 0 || totalHeight <= 0) {
                        ui.colCharInput.text = "0";
                        return;
                    }

                    var colCount = parseInt(ui.colCountInput.text, 10) || 1;
                    var colGap = parseFloat(ui.colGapInput.text) || 0;
                    var cellWidth = (totalWidth - colGap * (colCount - 1)) / colCount;

                    // 列幅を基本フォントサイズ（ドキュメント単位）で割る
                    var fontSizeRaw = parseFloat(ui.baseFontSizeInput.text) || 9.5;
                    var fontSize = ui.rbUnitMmQ.value ? fontSizeRaw / PT_TO_Q : fontSizeRaw;
                    var fontSizeInUnit = fontSize / ptPerUnit;
                    var charCount = Math.floor(cellWidth / fontSizeInUnit);

                    ui.colCharInput.text = String(charCount);
                } catch (e) { }
            }

            function updateFrameEnabled(ui) {
                var on = ui.frameEnableCheck.value;
                ui.frameBleedCheck.enabled = on;
                ui.frameTopInput.enabled = on;
                ui.frameBottomInput.enabled = on;
                ui.frameLeftInput.enabled = on;
                ui.frameRightInput.enabled = on;
                ui.frameLinkCheck.enabled = on;
            }

            function updateTitleEnabled(ui) {
                var on = ui.titleFillCheck.value || ui.titleStrokeCheck.value;
                ui.titleLengthInput.enabled = on;
                ui.titleLengthAutoBtn.enabled = on;
                ui.rbTop.enabled = on;
                ui.rbBottom.enabled = on;
                ui.rbLeft.enabled = on;
                ui.rbRight.enabled = on;
                ui.titleExtensionInput.enabled = on;
            }

            function updateOuterEnabled(ui) {
                var on = ui.outerLineCheck.value;
                ui.strokeWeightInput.enabled = on;
                var extZero = (function () { var v = parseFloat(ui.extensionInput.text); return isNaN(v) || v === 0; })();
                ui.cornerInput.enabled = on && extZero;
                ui.extensionInput.enabled = on;
                ui.rbCapNone.enabled = on && !extZero;
                ui.rbCapRound.enabled = on && !extZero;
                ui.rbCapProject.enabled = on && !extZero;
                ui.columnMarginInput.enabled = on;
            }

            // 区切り線のディム表示（両方の間隔が0のとき無効化）
            function updateDividerEnabled(ui) {
                var colGap = parseFloat(ui.colGapInput.text) || 0;
                var rowGap = parseFloat(ui.rowGapInput.text) || 0;
                var on = colGap > 0 || rowGap > 0;
                ui.innerStrokeCheck.enabled = on;
                ui.rbLineSolid.enabled = on && ui.innerStrokeCheck.value;
                ui.rbLineDashed.enabled = on && ui.innerStrokeCheck.value;
                ui.rbLineDotted.enabled = on && ui.innerStrokeCheck.value;
                ui.divWeightInput.enabled = on && ui.innerStrokeCheck.value;
            }

            // 間隔のディム表示（対応する数が1のとき無効化）
            function updateGapEnabled(ui) {
                var colCount = parseInt(ui.colCountInput.text) || 1;
                var rowCount = parseInt(ui.rowCountInput.text) || 1;
                ui.colGapInput.enabled = colCount > 1;
                ui.rowGapInput.enabled = rowCount > 1;
                ui.gapLinkCheck.enabled = colCount > 1 && rowCount > 1;
            }

            // pt <-> Q 変換定数: 1pt = 0.3528mm, 1Q = 0.25mm, 1pt = 1.41102Q
            var PT_TO_Q = 1.41102;

            var currentFontUnitIsQ = false;
            var currentStrokeUnitIsMm = false;
            var PT_TO_MM = 0.3528;

            function switchFontUnit(ui, toQ) {
                if (toQ === currentFontUnitIsQ) return;
                var factor = toQ ? PT_TO_Q : (1 / PT_TO_Q);
                var unitStr = toQ ? "Q" : "pt";
                var leadingUnitStr = toQ ? "H" : "pt";
                var fields = [ui.baseFontSizeInput, ui.leadingInput];
                for (var i = 0; i < fields.length; i++) {
                    var val = parseFloat(fields[i].text);
                    if (!isNaN(val) && val > 0) {
                        fields[i].text = String(Math.round(val * factor * 100) / 100);
                    }
                }
                ui.fontUnitLabel.text = unitStr;
                ui.leadingUnitLabel.text = leadingUnitStr;
                currentFontUnitIsQ = toQ;
            }

            function switchStrokeUnit(ui, toMm) {
                if (toMm === currentStrokeUnitIsMm) return;
                var factor = toMm ? PT_TO_MM : (1 / PT_TO_MM);
                var unitStr = toMm ? "mm" : "pt";
                var fields = [ui.strokeWeightInput, ui.columnWeightInput, ui.divWeightInput];
                for (var i = 0; i < fields.length; i++) {
                    var val = parseFloat(fields[i].text);
                    if (!isNaN(val) && val > 0) {
                        fields[i].text = String(Math.round(val * factor * 1000) / 1000);
                    }
                }
                ui.strokeUnitLabel.text = unitStr;
                ui.columnWeightUnitLabel.text = unitStr;
                ui.divWeightUnitLabel.text = unitStr;
                currentStrokeUnitIsMm = toMm;
            }

            var prevRelativeValue = 0;
            function bindDialogEvents() {
                ui.rbUnitMmPt.onClick = function () { switchFontUnit(ui, false); switchStrokeUnit(ui, false); updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.rbUnitMmMmPt.onClick = function () { switchFontUnit(ui, false); switchStrokeUnit(ui, true); updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.rbUnitMmQ.onClick = function () { switchFontUnit(ui, true); switchStrokeUnit(ui, true); updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.baseFontSizeInput.onChanging = ui.baseFontSizeInput.onChange = function () { updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.leadingInput.onChanging = ui.leadingInput.onChange = function () { if (ui.previewCheck.value) updatePreview(); };
                ui.showLayoutGridCheck.onClick = function () { if (ui.previewCheck.value) updatePreview(); };
                ui.outerLineCheck.onClick = function () { updateOuterEnabled(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.strokeWeightInput.onChanging = ui.strokeWeightInput.onChange = function () { if (ui.previewCheck.value) updatePreview(); };

                ui.marginTopInput.onChanging = ui.marginTopInput.onChange = function () { updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.marginBottomInput.onChanging = ui.marginBottomInput.onChange = function () { updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.marginLeftInput.onChanging = ui.marginLeftInput.onChange = function () { updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.marginRightInput.onChanging = ui.marginRightInput.onChange = function () { updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };

                ui.extensionInput.onChanging = ui.extensionInput.onChange = function () { updateOuterEnabled(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.cornerInput.onChanging = ui.cornerInput.onChange = function () { if (ui.previewCheck.value) updatePreview(); };
                ui.titleLengthInput.onChanging = ui.titleLengthInput.onChange = function () { updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.titleLengthAutoBtn.onClick = function () {
                    try {
                        var titleVal = parseFloat(ui.titleLengthInput.text) || 0;
                        if (titleVal <= 0) return;
                        var fontSizeRaw = parseFloat(ui.baseFontSizeInput.text) || 9.5;
                        var leadingStr = ui.leadingInput.text;
                        var leadingVal = parseFloat(leadingStr);
                        if (isNaN(leadingVal) || leadingVal <= 0) leadingVal = fontSizeRaw * 1.5;
                        var leadingPt = ui.rbUnitMmQ.value ? leadingVal / PT_TO_Q : leadingVal;
                        var leadingInUnit = leadingPt / ptPerUnit;
                        var fontSizePt = ui.rbUnitMmQ.value ? fontSizeRaw / PT_TO_Q : fontSizeRaw;
                        var fontSizeInUnit = fontSizePt / ptPerUnit;
                        // 仮グリッド線位置: fontSize + n * leading
                        var n = Math.round((titleVal - fontSizeInUnit) / leadingInUnit);
                        if (n < 0) n = 0;
                        var newVal = fontSizeInUnit + n * leadingInUnit;
                        ui.titleLengthInput.text = String(Math.round(newVal * 1000) / 1000);
                        updateCharCount(ui);
                        if (ui.previewCheck.value) updatePreview();
                    } catch (e) { }
                };
                ui.titleExtensionInput.onChanging = ui.titleExtensionInput.onChange = function () { if (ui.previewCheck.value) updatePreview(); };
                ui.rbTop.onClick = ui.rbBottom.onClick = ui.rbLeft.onClick = ui.rbRight.onClick = function () { updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.titleFillCheck.onClick = ui.titleStrokeCheck.onClick = function () { updateTitleEnabled(ui); updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };

                ui.columnFillCheck.onClick = ui.columnStrokeCheck.onClick = function () {
                    var colOn = ui.columnFillCheck.value || ui.columnStrokeCheck.value;
                    ui.columnHeightAutoBtn.enabled = colOn;
                    updateCharCount(ui);
                    if (ui.previewCheck.value) updatePreview();
                };
                ui.columnHeightInput.onChanging = ui.columnHeightInput.onChange = function () {
                    updateCharCount(ui);
                    if (ui.previewCheck.value) updatePreview();
                };
                ui.columnMarginInput.onChanging = ui.columnMarginInput.onChange = function () {
                    updateCharCount(ui);
                    if (ui.previewCheck.value) updatePreview();
                };
                ui.columnHeightAutoBtn.onClick = function () {
                    try {
                        var colHVal = parseFloat(ui.columnHeightInput.text) || 0;
                        if (colHVal <= 0) return;
                        var fontSizeRaw = parseFloat(ui.baseFontSizeInput.text) || 9.5;
                        var leadingStr = ui.leadingInput.text;
                        var leadingVal = parseFloat(leadingStr);
                        if (isNaN(leadingVal) || leadingVal <= 0) leadingVal = fontSizeRaw * 1.5;
                        var leadingPt = ui.rbUnitMmQ.value ? leadingVal / PT_TO_Q : leadingVal;
                        var leadingInUnit = leadingPt / ptPerUnit;
                        var fontSizePt = ui.rbUnitMmQ.value ? fontSizeRaw / PT_TO_Q : fontSizeRaw;
                        var fontSizeInUnit = fontSizePt / ptPerUnit;

                        // outer上端からグリッド線位置を計算し、フッター上端がグリッド線に揃うように調整
                        var bounds = getPageBoundsOnSpread(page);
                        var oTop = bounds[0] + (parseFloat(ui.marginTopInput.text) || 0);
                        var oBottom = bounds[2] - (parseFloat(ui.marginBottomInput.text) || 0);
                        var colMargin = parseFloat(ui.columnMarginInput.text) || 0;

                        // 現在のフッター上端位置（outer下端から逆算）
                        var footerTop = oBottom - colMargin - colHVal;
                        // outer上端からの距離をグリッド線位置に丸める
                        var distFromTop = footerTop - oTop;
                        var n = Math.round((distFromTop - fontSizeInUnit) / leadingInUnit);
                        if (n < 0) n = 0;
                        var snappedFooterTop = oTop + fontSizeInUnit + n * leadingInUnit;
                        var newVal = oBottom - colMargin - snappedFooterTop;
                        if (newVal <= 0) newVal = fontSizeInUnit;
                        ui.columnHeightInput.text = String(Math.round(newVal * 1000) / 1000);
                        updateCharCount(ui);
                        if (ui.previewCheck.value) updatePreview();
                    } catch (e) { }
                };

                ui.columnCornerInput.onChanging = ui.columnCornerInput.onChange = function () { if (ui.previewCheck.value) updatePreview(); };
                ui.columnWeightInput.onChanging = ui.columnWeightInput.onChange = function () { if (ui.previewCheck.value) updatePreview(); };
                ui.rbCapNone.onClick = ui.rbCapRound.onClick = ui.rbCapProject.onClick = function () { if (ui.previewCheck.value) updatePreview(); };
                ui.frameEnableCheck.onClick = function () { updateFrameEnabled(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.frameTopInput.onChanging = ui.frameTopInput.onChange = function () { syncFrameWidths(ui, ui.frameTopInput); if (ui.previewCheck.value) updatePreview(); };
                ui.frameBottomInput.onChanging = ui.frameBottomInput.onChange = function () { syncFrameWidths(ui, ui.frameBottomInput); if (ui.previewCheck.value) updatePreview(); };
                ui.frameLeftInput.onChanging = ui.frameLeftInput.onChange = function () { syncFrameWidths(ui, ui.frameLeftInput); if (ui.previewCheck.value) updatePreview(); };
                ui.frameRightInput.onChanging = ui.frameRightInput.onChange = function () { syncFrameWidths(ui, ui.frameRightInput); if (ui.previewCheck.value) updatePreview(); };
                ui.frameBleedCheck.onClick = function () { if (ui.previewCheck.value) updatePreview(); };
                ui.innerTopInput.onChanging = ui.innerTopInput.onChange = function () { syncInnerMargins(ui, ui.innerTopInput); updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.innerBottomInput.onChanging = ui.innerBottomInput.onChange = function () { syncInnerMargins(ui, ui.innerBottomInput); updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.innerLeftInput.onChanging = ui.innerLeftInput.onChange = function () { syncInnerMargins(ui, ui.innerLeftInput); updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.innerRightInput.onChanging = ui.innerRightInput.onChange = function () { syncInnerMargins(ui, ui.innerRightInput); updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.rbFillColor.onClick = ui.rbFillTextFrame.onClick = function () { updateFillEnabled(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.threadCheck.onClick = function () { if (ui.previewCheck.value) updatePreview(); };
                ui.relativeInput.onChanging = ui.relativeInput.onChange = function () {
                    var newVal = parseFloat(ui.relativeInput.text) || 0;
                    var delta = newVal - prevRelativeValue;
                    prevRelativeValue = newVal;
                    var inputs = [ui.marginTopInput, ui.marginBottomInput, ui.marginLeftInput, ui.marginRightInput];
                    for (var i = 0; i < inputs.length; i++) {
                        var v = parseFloat(inputs[i].text) || 0;
                        inputs[i].text = String(Math.round((v + delta) * 100) / 100);
                    }
                    updateCharCount(ui);
                    if (ui.previewCheck.value) updatePreview();
                };
                ui.colCountInput.onChanging = ui.colCountInput.onChange = function () { updateGapEnabled(ui); updateDividerEnabled(ui); updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.colGapInput.onChanging = ui.colGapInput.onChange = function () { if (ui.gapLinkCheck.value) ui.rowGapInput.text = ui.colGapInput.text; updateDividerEnabled(ui); updateCharCount(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.colCharInput.onChanging = ui.colCharInput.onChange = function () {
                    try {
                        var charCount = parseInt(ui.colCharInput.text, 10);
                        if (isNaN(charCount) || charCount < 1) return;

                        var bounds = getPageBoundsOnSpread(page);
                        var mLeft = parseFloat(ui.marginLeftInput.text) || 0;
                        var mRight = parseFloat(ui.marginRightInput.text) || 0;
                        var outerLeft = bounds[1] + mLeft;
                        var outerRight = bounds[3] - mRight;
                        var contentLeft = outerLeft;
                        var contentRight = outerRight;

                        var titleOn = ui.titleFillCheck.value || ui.titleStrokeCheck.value;
                        var titleLen = parseFloat(ui.titleLengthInput.text) || 0;
                        if (titleOn && titleLen > 0) {
                            var pos = getSelectedPosition(ui);
                            if (pos === "left") contentLeft += titleLen;
                            else if (pos === "right") contentRight -= titleLen;
                        }

                        var innerL = parseFloat(ui.innerLeftInput.text) || 0;
                        var innerR = parseFloat(ui.innerRightInput.text) || 0;
                        var totalWidth = (contentRight - innerR) - (contentLeft + innerL);

                        var colCount = parseInt(ui.colCountInput.text) || 1;
                        var fontSizeRaw = parseFloat(ui.baseFontSizeInput.text) || 9.5;
                        var fontSize = ui.rbUnitMmQ.value ? fontSizeRaw / PT_TO_Q : fontSizeRaw;
                        var fontSizeInUnit = fontSize / ptPerUnit;

                        var targetCellWidth = charCount * fontSizeInUnit;
                        var newGap = (totalWidth - targetCellWidth * colCount) / (colCount - 1);
                        if (colCount <= 1) newGap = 0;
                        newGap = Math.round(newGap * 1000) / 1000;
                        if (newGap < 0) newGap = 0;

                        ui.colGapInput.text = String(newGap);
                        if (ui.gapLinkCheck.value) ui.rowGapInput.text = ui.colGapInput.text;
                        updateDividerEnabled(ui);
                        if (ui.previewCheck.value) updatePreview();
                    } catch (e) { }
                };
                ui.colGapAutoBtn.onClick = function () {
                    try {
                        var bounds = getPageBoundsOnSpread(page);
                        var pgLeft = bounds[1];
                        var pgRight = bounds[3];

                        var mTop = parseFloat(ui.marginTopInput.text) || 0;
                        var mBottom = parseFloat(ui.marginBottomInput.text) || 0;
                        var mLeft = parseFloat(ui.marginLeftInput.text) || 0;
                        var mRight = parseFloat(ui.marginRightInput.text) || 0;

                        var outerTop = bounds[0] + mTop;
                        var outerBottom = bounds[2] - mBottom;
                        var outerLeft = pgLeft + mLeft;
                        var outerRight = pgRight - mRight;

                        var contentTop = outerTop;
                        var contentBottom = outerBottom;
                        var contentLeft = outerLeft;
                        var contentRight = outerRight;

                        var titleOn = ui.titleFillCheck.value || ui.titleStrokeCheck.value;
                        var titleLen = parseFloat(ui.titleLengthInput.text) || 0;
                        if (titleOn && titleLen > 0) {
                            var pos = getSelectedPosition(ui);
                            if (pos === "top") contentTop += titleLen;
                            else if (pos === "bottom") contentBottom -= titleLen;
                            else if (pos === "left") contentLeft += titleLen;
                            else if (pos === "right") contentRight -= titleLen;
                        }

                        var columnOn = ui.columnFillCheck.value || ui.columnStrokeCheck.value;
                        var columnHeight = parseFloat(ui.columnHeightInput.text) || 0;
                        var columnMargin = parseFloat(ui.columnMarginInput.text) || 0;
                        if (columnOn && columnHeight > 0) {
                            contentBottom -= (columnHeight + columnMargin);
                        }

                        var innerL = parseFloat(ui.innerLeftInput.text) || 0;
                        var innerR = parseFloat(ui.innerRightInput.text) || 0;
                        var totalWidth = (contentRight - innerR) - (contentLeft + innerL);

                        var colCount = parseInt(ui.colCountInput.text) || 1;
                        var fontSizeRaw = parseFloat(ui.baseFontSizeInput.text) || 9.5;
                        var fontSize = ui.rbUnitMmQ.value ? fontSizeRaw / PT_TO_Q : fontSizeRaw;
                        var fontSizeInUnit = fontSize / ptPerUnit;

                        var charCount = parseInt(ui.colCharInput.text, 10);
                        if (isNaN(charCount) || charCount < 1) charCount = 1;

                        var targetCellWidth = charCount * fontSizeInUnit;
                        var newGap = (totalWidth - targetCellWidth * colCount) / (colCount - 1);
                        if (colCount <= 1) newGap = 0;
                        newGap = Math.round(newGap * 1000) / 1000;
                        if (newGap < 0) newGap = 0;

                        ui.colGapInput.text = String(newGap);
                        if (ui.gapLinkCheck.value) ui.rowGapInput.text = ui.colGapInput.text;
                        updateDividerEnabled(ui);
                        updateCharCount(ui);
                        if (ui.previewCheck.value) updatePreview();
                    } catch (e) { }
                };
                ui.offsetAutoBtn.onClick = function () {
                    try {
                        var fontSizeRaw = parseFloat(ui.baseFontSizeInput.text) || 9.5;
                        var fontSize = ui.rbUnitMmQ.value ? fontSizeRaw / PT_TO_Q : fontSizeRaw;
                        var fontSizeInUnit = fontSize / ptPerUnit;
                        // 左右：文字サイズの倍数に調整
                        var lrInputs = [ui.innerLeftInput, ui.innerRightInput];
                        for (var i = 0; i < lrInputs.length; i++) {
                            var val = parseFloat(lrInputs[i].text) || 0;
                            if (val === 0) continue;
                            var charCount = Math.round(val / fontSizeInUnit);
                            if (charCount < 1) charCount = 1;
                            lrInputs[i].text = String(Math.round(charCount * fontSizeInUnit * 1000) / 1000);
                        }
                        // 天地：outerTop基準の仮グリッド線にぴったり合うように調整
                        var leadingStr = ui.leadingInput.text;
                        var leadingVal = parseFloat(leadingStr);
                        if (isNaN(leadingVal) || leadingVal <= 0) {
                            leadingVal = fontSizeRaw * 1.5;
                        }
                        var leadingPt = ui.rbUnitMmQ.value ? leadingVal / PT_TO_Q : leadingVal;
                        var leadingInUnit = leadingPt / ptPerUnit;

                        var bounds = getPageBoundsOnSpread(page);
                        var oTop = bounds[0] + (parseFloat(ui.marginTopInput.text) || 0);
                        var oBottom = bounds[2] - (parseFloat(ui.marginBottomInput.text) || 0);

                        // タイトルエリア分を考慮したコンテンツ上端・下端
                        var cTop = oTop;
                        var cBottom = oBottom;
                        var titleOn = ui.titleFillCheck.value || ui.titleStrokeCheck.value;
                        var titleLen = parseFloat(ui.titleLengthInput.text) || 0;
                        if (titleOn && titleLen > 0) {
                            var pos = getSelectedPosition(ui);
                            if (pos === "top") cTop += titleLen;
                            else if (pos === "bottom") cBottom -= titleLen;
                        }
                        // フッター分を考慮
                        var colOn = ui.columnFillCheck.value || ui.columnStrokeCheck.value;
                        var colH = parseFloat(ui.columnHeightInput.text) || 0;
                        var colM = parseFloat(ui.columnMarginInput.text) || 0;
                        if (colOn && colH > 0) cBottom -= (colH + colM);

                        // 天：コンテンツ上端 + innerTop がグリッド線に乗るように
                        var innerTopVal = parseFloat(ui.innerTopInput.text) || 0;
                        if (innerTopVal > 0) {
                            var absTop = cTop + innerTopVal;
                            var nTop = Math.round((absTop - oTop - fontSizeInUnit) / leadingInUnit);
                            if (nTop < 0) nTop = 0;
                            var snappedTop = oTop + fontSizeInUnit + nTop * leadingInUnit;
                            var newInnerTop = snappedTop - cTop;
                            if (newInnerTop < 0) newInnerTop = 0;
                            ui.innerTopInput.text = String(Math.round(newInnerTop * 1000) / 1000);
                        }

                        // 地：コンテンツ下端 - innerBottom がグリッド線に乗るように
                        var innerBottomVal = parseFloat(ui.innerBottomInput.text) || 0;
                        if (innerBottomVal > 0) {
                            var absBottom = cBottom - innerBottomVal;
                            var nBottom = Math.round((absBottom - oTop - fontSizeInUnit) / leadingInUnit);
                            if (nBottom < 0) nBottom = 0;
                            var snappedBottom = oTop + fontSizeInUnit + nBottom * leadingInUnit;
                            var newInnerBottom = cBottom - snappedBottom;
                            if (newInnerBottom < 0) newInnerBottom = 0;
                            ui.innerBottomInput.text = String(Math.round(newInnerBottom * 1000) / 1000);
                        }
                        updateCharCount(ui);
                        if (ui.previewCheck.value) updatePreview();
                    } catch (e) { }
                };
                ui.allAutoBtn.onClick = function () {
                    ui.titleLengthAutoBtn.notify("onClick");
                    ui.columnHeightAutoBtn.notify("onClick");
                    ui.offsetAutoBtn.notify("onClick");
                    ui.colGapAutoBtn.notify("onClick");
                };
                ui.rowCountInput.onChanging = ui.rowCountInput.onChange = function () { updateGapEnabled(ui); updateDividerEnabled(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.rowGapInput.onChanging = ui.rowGapInput.onChange = function () { if (ui.gapLinkCheck.value) ui.colGapInput.text = ui.rowGapInput.text; updateDividerEnabled(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.innerStrokeCheck.onClick = function () { updateDividerEnabled(ui); if (ui.previewCheck.value) updatePreview(); };
                ui.rbLineSolid.onClick = ui.rbLineDashed.onClick = ui.rbLineDotted.onClick = function () { if (ui.previewCheck.value) updatePreview(); };
                ui.divWeightInput.onChanging = ui.divWeightInput.onChange = function () { if (ui.previewCheck.value) updatePreview(); };
            }

            function initializeDialogState() {
                updateFrameEnabled(ui);
                updateTitleEnabled(ui);
                updateOuterEnabled(ui);
                updateFillEnabled(ui);
                updateGapEnabled(ui);
                updateDividerEnabled(ui);
                ui.columnHeightAutoBtn.enabled = ui.columnFillCheck.value || ui.columnStrokeCheck.value;
                updateCharCount(ui);

                ui.previewCheck.onClick = function () { updatePreview(); };

                updatePreview();
            }

            bindDialogEvents();
            initializeDialogState();

            // ダイアログを表示
            if (ui.dlg.show() === 1) {
                destroyPreviewLayer();

                var finalValues = getCurrentUIValues(ui);
                var keepGrid = ui.keepGridCheck.value && ui.showLayoutGridCheck.value;

                app.doScript(function () {
                    // アクティブレイヤーが "Temp Grid" なら別レイヤーに切り替え
                    try {
                        if (doc.activeLayer.name === "Temp Grid") {
                            for (var li = 0; li < doc.layers.length; li++) {
                                var lyr = doc.layers[li];
                                if (lyr.name !== "Temp Grid" && lyr.name !== PREVIEW_LAYER_NAME) {
                                    doc.activeLayer = lyr;
                                    break;
                                }
                            }
                        }
                    } catch (e) { }
                    createLines(finalValues);
                    // 「残す」がONなら新規レイヤーに仮グリッドのみ描画
                    if (keepGrid) {
                        var gridLayerName = "Temp Grid";
                        var gridLayer;
                        try {
                            gridLayer = doc.layers.itemByName(gridLayerName);
                            gridLayer.name;
                        } catch (e) {
                            gridLayer = doc.layers.add({ name: gridLayerName });
                        }
                        try { gridLayer.printable = false; } catch (e) { }
                        var gridValues = getCurrentUIValues(ui);
                        gridValues.targetLayer = gridLayer;
                        gridValues.gridOnly = true;
                        createLines(gridValues);
                    }
                }, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, LABELS.dialogTitle[lang]);
            } else {
                destroyPreviewLayer();
            }
        }

        main();

        // カラーを取得または作成するヘルパー関数
        function getOrCreateColor(doc, name, cmykValues) {
            try {
                var c = doc.colors.item(name);
                c.name;
                return c;
            } catch (e) {
                return doc.colors.add({
                    name: name,
                    model: ColorModel.PROCESS,
                    space: ColorSpace.CMYK,
                    colorValue: cmykValues
                });
            }
        }

        // 罫線を作成する関数（作成したオブジェクトの配列を返す）
        function createLines(opts) {
            var outerLine = opts.outerLine, marginTop = opts.marginTop, marginBottom = opts.marginBottom;
            var marginLeft = opts.marginLeft, marginRight = opts.marginRight;
            var extension = opts.extension, cornerRadius = opts.cornerRadius;
            var titleLength = opts.titleLength, titleExtension = opts.titleExtension;
            var titleCornerRadius = opts.titleCornerRadius, titlePosition = opts.titlePosition;
            var titleFill = opts.titleFill, titleStroke = opts.titleStroke;
            var capStyle = opts.capStyle, lineWeight = opts.lineWeight;
            var frameEnable = opts.frameEnable, frameTop = opts.frameTop, frameBottom = opts.frameBottom;
            var frameLeft = opts.frameLeft, frameRight = opts.frameRight;
            var frameBleed = opts.frameBleed;
            var innerFill = opts.innerFill, innerTextFrame = opts.innerTextFrame;
            var threadText = opts.threadText;
            var sampleText = opts.sampleText;
            var sampleSquareCircle = opts.sampleSquareCircle;
            var baseFontSize = opts.baseFontSize;
            var leadingValue = opts.leading;
            var showLayoutGrid = opts.showLayoutGrid;
            var ptPerUnit = opts.ptPerUnit || 1;
            var innerTop = opts.innerTop, innerBottom = opts.innerBottom;
            var innerLeft = opts.innerLeft, innerRight = opts.innerRight;
            var colCount = opts.colCount, colGap = opts.colGap;
            var rowCount = opts.rowCount, rowGap = opts.rowGap;
            var innerStroke = opts.innerStroke, lineType = opts.lineType;
            var divWeight = opts.divWeight;
            var columnFill = opts.columnFill, columnStroke = opts.columnStroke;
            var columnHeight = opts.columnHeight, columnMargin = opts.columnMargin;
            var columnCornerRadius = opts.columnCornerRadius;
            var columnWeight = opts.columnWeight;
            var targetLayer = opts.targetLayer || null;
            var gridOnly = opts.gridOnly || false;

            var doc = app.activeDocument;
            var win = app.activeWindow;
            var items = [];
            var prevActiveLayer = null;

            // 現在の単位設定を保存し、垂直単位を水平に統一
            var oldYUnits = doc.viewPreferences.verticalMeasurementUnits;
            doc.viewPreferences.verticalMeasurementUnits = doc.viewPreferences.horizontalMeasurementUnits;

            // ルーラー原点とゼロポイントを保存し、スプレッド原点に統一（見開き右ページ対応）
            var oldRulerOrigin = doc.viewPreferences.rulerOrigin;
            var oldZeroPoint = doc.zeroPoint;
            doc.viewPreferences.rulerOrigin = RulerOrigin.SPREAD_ORIGIN;
            doc.zeroPoint = [0, 0];

            try {
                // 現在アクティブな（選択されている）ページを取得
                var page = win.activePage;

                // ページのサイズ（座標）を取得: [上, 左, 下, 右]
                var bounds = page.bounds;
                var top = bounds[0];
                var left = bounds[1];
                var bottom = bounds[2];
                var right = bounds[3];

                var createParent = page;
                if (targetLayer && targetLayer.isValid) {
                    try {
                        prevActiveLayer = doc.activeLayer;
                    } catch (e0) {
                        prevActiveLayer = null;
                    }
                    try {
                        doc.activeLayer = targetLayer;
                    } catch (e00) { }
                }

                // マージン位置を計算（上下左右個別）
                var mTop = top + marginTop;
                var mBottom = bottom - marginBottom;
                var mLeft = left + marginLeft;
                var mRight = right - marginRight;

                // 外側エリアと実コンテンツ領域を分離
                var outerTop = mTop;
                var outerBottom = mBottom;
                var outerLeft = mLeft;
                var outerRight = mRight;

                var contentTop = outerTop;
                var contentBottom = outerBottom;
                var contentLeft = outerLeft;
                var contentRight = outerRight;

                if ((titleFill || titleStroke) && titleLength > 0) {
                    if (titlePosition === "top") contentTop += titleLength;
                    else if (titlePosition === "bottom") contentBottom -= titleLength;
                    else if (titlePosition === "left") contentLeft += titleLength;
                    else if (titlePosition === "right") contentRight -= titleLength;
                }

                // フッターの「コラム」エリア用に、content 下端の元位置を保持
                var contentBottomBeforeFooter = contentBottom;

                // フッターの「コラム」エリア（＋アキ）を実コンテンツ領域から差し引く
                if ((columnFill || columnStroke) && columnHeight > 0) {
                    contentBottom -= (columnHeight + columnMargin);
                }

                if (contentBottom < contentTop) contentBottom = contentTop;

                // フッターの「コラム」エリアの描画用バウンズ（content 側基準）
                var colBounds = null;
                if ((columnFill || columnStroke) && columnHeight > 0) {
                    var colTop = contentBottom + columnMargin;
                    var colBottom = colTop + columnHeight;
                    var colLeft = contentLeft;
                    var colRight = contentRight;

                    if (colTop < contentBottom) colTop = contentBottom;
                    if (colBottom > contentBottomBeforeFooter) colBottom = contentBottomBeforeFooter;

                    if (colBottom > colTop && colRight > colLeft) {
                        colBounds = [colTop, colLeft, colBottom, colRight];
                    }
                }

                // 版面の罫線はフッターエリア（高さ＋アキ）を除いた範囲で描画
                var strokeMBottom = mBottom;
                if ((columnFill || columnStroke) && columnHeight > 0) {
                    strokeMBottom = mBottom - columnHeight - columnMargin;
                }

                var ext = extension;
                var strokeWeight = lineWeight;

                // 線端の設定
                var endCap = EndCap.BUTT_END_CAP;
                if (capStyle === "round") endCap = EndCap.ROUND_END_CAP;
                else if (capStyle === "project") endCap = EndCap.PROJECTING_END_CAP;

                if (!gridOnly && outerLine) {
                    if (ext === 0) {
                        // 伸縮が0の場合、長方形として作成
                        var rect = createParent.rectangles.add({
                            geometricBounds: [mTop, mLeft, strokeMBottom, mRight],
                            strokeWeight: strokeWeight,
                            strokeColor: doc.swatches.item("Black"),
                            fillColor: doc.swatches.item("None")
                        });
                        if (cornerRadius > 0) {
                            rect.topLeftCornerOption = CornerOptions.ROUNDED_CORNER;
                            rect.topRightCornerOption = CornerOptions.ROUNDED_CORNER;
                            rect.bottomLeftCornerOption = CornerOptions.ROUNDED_CORNER;
                            rect.bottomRightCornerOption = CornerOptions.ROUNDED_CORNER;
                            rect.topLeftCornerRadius = cornerRadius;
                            rect.topRightCornerRadius = cornerRadius;
                            rect.bottomLeftCornerRadius = cornerRadius;
                            rect.bottomRightCornerRadius = cornerRadius;
                        }
                        items.push(rect);
                    } else if (ext > 0) {
                        // 正の伸縮：各辺を外側に伸ばす
                        var line;

                        line = createParent.graphicLines.add();
                        line.paths[0].entirePath = [[mLeft - ext, mTop], [mRight + ext, mTop]];
                        line.strokeWeight = strokeWeight;
                        line.strokeColor = doc.swatches.item("Black");
                        line.endCap = endCap;
                        items.push(line);

                        line = createParent.graphicLines.add();
                        line.paths[0].entirePath = [[mLeft - ext, strokeMBottom], [mRight + ext, strokeMBottom]];
                        line.strokeWeight = strokeWeight;
                        line.strokeColor = doc.swatches.item("Black");
                        line.endCap = endCap;
                        items.push(line);

                        line = createParent.graphicLines.add();
                        line.paths[0].entirePath = [[mLeft, mTop - ext], [mLeft, strokeMBottom + ext]];
                        line.strokeWeight = strokeWeight;
                        line.strokeColor = doc.swatches.item("Black");
                        line.endCap = endCap;
                        items.push(line);

                        line = createParent.graphicLines.add();
                        line.paths[0].entirePath = [[mRight, mTop - ext], [mRight, strokeMBottom + ext]];
                        line.strokeWeight = strokeWeight;
                        line.strokeColor = doc.swatches.item("Black");
                        line.endCap = endCap;
                        items.push(line);
                    } else {
                        // 負の伸縮：各辺を内側に縮める（角に隙間ができる）
                        var shrink = Math.abs(ext);
                        var line;

                        // 上の罫線
                        line = createParent.graphicLines.add();
                        line.paths[0].entirePath = [[mLeft + shrink, mTop], [mRight - shrink, mTop]];
                        line.strokeWeight = strokeWeight;
                        line.strokeColor = doc.swatches.item("Black");
                        line.endCap = endCap;
                        items.push(line);

                        // 下の罫線
                        line = createParent.graphicLines.add();
                        line.paths[0].entirePath = [[mLeft + shrink, strokeMBottom], [mRight - shrink, strokeMBottom]];
                        line.strokeWeight = strokeWeight;
                        line.strokeColor = doc.swatches.item("Black");
                        line.endCap = endCap;
                        items.push(line);

                        // 左の罫線
                        line = createParent.graphicLines.add();
                        line.paths[0].entirePath = [[mLeft, mTop + shrink], [mLeft, strokeMBottom - shrink]];
                        line.strokeWeight = strokeWeight;
                        line.strokeColor = doc.swatches.item("Black");
                        line.endCap = endCap;
                        items.push(line);

                        // 右の罫線
                        line = createParent.graphicLines.add();
                        line.paths[0].entirePath = [[mRight, mTop + shrink], [mRight, strokeMBottom - shrink]];
                        line.strokeWeight = strokeWeight;
                        line.strokeColor = doc.swatches.item("Black");
                        line.endCap = endCap;
                        items.push(line);
                    }
                } // outerLine

                // タイトルエリア
                if (!gridOnly && titleLength > 0) {
                    var titleLen = titleLength;
                    var titleExt = titleExtension;

                    // 塗り（独立）: K30長方形、位置に応じた2角だけ角丸
                    if (titleFill) {
                        var tBounds;
                        if (titlePosition === "top") {
                            tBounds = [mTop, mLeft, mTop + titleLen, mRight];
                        } else if (titlePosition === "bottom") {
                            tBounds = [mBottom - titleLen, mLeft, mBottom, mRight];
                        } else if (titlePosition === "left") {
                            tBounds = [mTop, mLeft, mBottom, mLeft + titleLen];
                        } else if (titlePosition === "right") {
                            tBounds = [mTop, mRight - titleLen, mBottom, mRight];
                        }
                        var tRect = createParent.rectangles.add({
                            geometricBounds: tBounds,
                            strokeWeight: 0,
                            strokeColor: doc.swatches.item("None"),
                            fillColor: getOrCreateColor(doc, "K25", [0, 0, 0, 25])
                        });
                        if (titleCornerRadius > 0) {
                            if (titlePosition === "top") {
                                tRect.topLeftCornerOption = CornerOptions.ROUNDED_CORNER;
                                tRect.topRightCornerOption = CornerOptions.ROUNDED_CORNER;
                                tRect.topLeftCornerRadius = titleCornerRadius;
                                tRect.topRightCornerRadius = titleCornerRadius;
                            } else if (titlePosition === "bottom") {
                                tRect.bottomLeftCornerOption = CornerOptions.ROUNDED_CORNER;
                                tRect.bottomRightCornerOption = CornerOptions.ROUNDED_CORNER;
                                tRect.bottomLeftCornerRadius = titleCornerRadius;
                                tRect.bottomRightCornerRadius = titleCornerRadius;
                            } else if (titlePosition === "left") {
                                tRect.topLeftCornerOption = CornerOptions.ROUNDED_CORNER;
                                tRect.bottomLeftCornerOption = CornerOptions.ROUNDED_CORNER;
                                tRect.topLeftCornerRadius = titleCornerRadius;
                                tRect.bottomLeftCornerRadius = titleCornerRadius;
                            } else if (titlePosition === "right") {
                                tRect.topRightCornerOption = CornerOptions.ROUNDED_CORNER;
                                tRect.bottomRightCornerOption = CornerOptions.ROUNDED_CORNER;
                                tRect.topRightCornerRadius = titleCornerRadius;
                                tRect.bottomRightCornerRadius = titleCornerRadius;
                            }
                        }
                        tRect.sendToBack();
                        items.push(tRect);
                    }

                    // 罫線（独立）: 常に内側に1本
                    if (titleStroke) {
                        var titleLine = createParent.graphicLines.add();
                        if (titlePosition === "top") {
                            var titleY = mTop + titleLen;
                            titleLine.paths[0].entirePath = [[mLeft - titleExt, titleY], [mRight + titleExt, titleY]];
                        } else if (titlePosition === "bottom") {
                            var titleY = mBottom - titleLen;
                            titleLine.paths[0].entirePath = [[mLeft - titleExt, titleY], [mRight + titleExt, titleY]];
                        } else if (titlePosition === "left") {
                            var titleX = mLeft + titleLen;
                            titleLine.paths[0].entirePath = [[titleX, mTop - titleExt], [titleX, mBottom + titleExt]];
                        } else if (titlePosition === "right") {
                            var titleX = mRight - titleLen;
                            titleLine.paths[0].entirePath = [[titleX, mTop - titleExt], [titleX, mBottom + titleExt]];
                        }
                        titleLine.strokeWeight = strokeWeight;
                        titleLine.strokeColor = doc.swatches.item("Black");
                        items.push(titleLine);
                    }
                }

                // カラムエリアの作成（外側エリアの下にタイトルエリアの「下」固定版）
                if (!gridOnly && (columnFill || columnStroke) && columnHeight > 0 && colBounds) {
                    var colAreaTop = mBottom + columnMargin;
                    var colAreaBottom = colAreaTop + columnHeight;
                    var colAreaLeft = mLeft;
                    var colAreaRight = mRight;

                    function applyCornerRadius(rect, radius) {
                        if (radius > 0) {
                            rect.topLeftCornerOption = CornerOptions.ROUNDED_CORNER;
                            rect.topRightCornerOption = CornerOptions.ROUNDED_CORNER;
                            rect.bottomLeftCornerOption = CornerOptions.ROUNDED_CORNER;
                            rect.bottomRightCornerOption = CornerOptions.ROUNDED_CORNER;
                            rect.topLeftCornerRadius = radius;
                            rect.topRightCornerRadius = radius;
                            rect.bottomLeftCornerRadius = radius;
                            rect.bottomRightCornerRadius = radius;
                        }
                    }

                    if (columnFill) {
                        var colAreaRect = createParent.rectangles.add({
                            geometricBounds: colBounds,
                            strokeWeight: 0,
                            strokeColor: doc.swatches.item("None"),
                            fillColor: getOrCreateColor(doc, "K40", [0, 0, 0, 40])
                        });
                        applyCornerRadius(colAreaRect, columnCornerRadius);
                        colAreaRect.sendToBack();
                        items.push(colAreaRect);
                    }

                    if (columnStroke) {
                        var colAreaStroke = createParent.rectangles.add({
                            geometricBounds: colBounds,
                            strokeWeight: columnWeight,
                            strokeColor: doc.swatches.item("Black"),
                            fillColor: doc.swatches.item("None")
                        });
                        applyCornerRadius(colAreaStroke, columnCornerRadius);
                        items.push(colAreaStroke);
                    }
                }

                // フレームの作成（額縁状の塗り図形：ポリゴン穴あき）
                if (!gridOnly && frameEnable) {
                    // 外側：ページサイズ（裁ち落としONなら外側方向のみ拡張。見開きの内側には伸ばさない）
                    var fOuterTop = top;
                    var fOuterBottom = bottom;
                    var fOuterLeft = left;
                    var fOuterRight = right;

                    if (frameBleed) {
                        var bleedOffset = 3;

                        // 上下は常に裁ち落とし
                        fOuterTop -= bleedOffset;
                        fOuterBottom += bleedOffset;

                        // 左右はページ側に応じて外側のみ拡張
                        if (page.side === PageSideOptions.LEFT_HAND) {
                            // 左ページ：外側は左
                            fOuterLeft -= bleedOffset;
                        } else if (page.side === PageSideOptions.RIGHT_HAND) {
                            // 右ページ：外側は右
                            fOuterRight += bleedOffset;
                        } else {
                            // 単ページ
                            fOuterLeft -= bleedOffset;
                            fOuterRight += bleedOffset;
                        }
                    }

                    // 内側（くり抜き）：ページサイズから天地左右だけ小さくする
                    var fInnerTop = top + frameTop;
                    var fInnerBottom = bottom - frameBottom;
                    var fInnerLeft = left + frameLeft;
                    var fInnerRight = right - frameRight;

                    var fk30Color = getOrCreateColor(doc, "K30", [0, 0, 0, 30]);

                    // ポリゴンで2パス構成（外側順回り＋内側逆回り＝型抜き）
                    var fPoly = createParent.polygons.add();
                    fPoly.paths[0].entirePath = [
                        [fOuterLeft, fOuterTop],
                        [fOuterRight, fOuterTop],
                        [fOuterRight, fOuterBottom],
                        [fOuterLeft, fOuterBottom]
                    ];
                    var cutPath = fPoly.paths.add();
                    cutPath.entirePath = [
                        [fInnerLeft, fInnerTop],
                        [fInnerLeft, fInnerBottom],
                        [fInnerRight, fInnerBottom],
                        [fInnerRight, fInnerTop]
                    ];
                    fPoly.fillColor = fk30Color;
                    fPoly.strokeColor = doc.swatches.item("None");
                    fPoly.strokeWeight = 0;
                    fPoly.sendToBack();
                    items.push(fPoly);
                }

                // グリッドの基準領域（タイトルエリア＋フッターカラムエリアを除外＋オフセット）
                var iBaseTop = contentTop;
                var iBaseBottom = contentBottom;
                var iBaseLeft = contentLeft;
                var iBaseRight = contentRight;
                var iRectTop = iBaseTop + innerTop;
                var iRectBottom = iBaseBottom - innerBottom;
                var iRectLeft = iBaseLeft + innerLeft;
                var iRectRight = iBaseRight - innerRight;

                // グリッドの塗り／テキストフレーム
                if (!gridOnly && (innerFill || innerTextFrame)) {

                    // 列×行のグリッドに分割して描画
                    var totalWidth = iRectRight - iRectLeft;
                    var totalHeight = iRectBottom - iRectTop;
                    var cellWidth = (totalWidth - colGap * (colCount - 1)) / colCount;
                    var cellHeight = (totalHeight - rowGap * (rowCount - 1)) / rowCount;

                    if (innerFill) {
                        var fillSwatchColor = getOrCreateColor(doc, "K10", [0, 0, 0, 10]);
                        for (var ci = 0; ci < colCount; ci++) {
                            for (var ri = 0; ri < rowCount; ri++) {
                                var cellLeft = iRectLeft + ci * (cellWidth + colGap);
                                var cellTop = iRectTop + ri * (cellHeight + rowGap);
                                var cellRight = cellLeft + cellWidth;
                                var cellBottom = cellTop + cellHeight;
                                var iRect = createParent.rectangles.add({
                                    geometricBounds: [cellTop, cellLeft, cellBottom, cellRight],
                                    strokeWeight: 0,
                                    strokeColor: doc.swatches.item("None"),
                                    fillColor: fillSwatchColor
                                });
                                iRect.sendToBack();
                                items.push(iRect);
                            }
                        }
                    } else if (innerTextFrame) {
                        var textFrames = [];
                        for (var ci = 0; ci < colCount; ci++) {
                            for (var ri = 0; ri < rowCount; ri++) {
                                var cellLeft = iRectLeft + ci * (cellWidth + colGap);
                                var cellTop = iRectTop + ri * (cellHeight + rowGap);
                                var cellRight = cellLeft + cellWidth;
                                var cellBottom = cellTop + cellHeight;
                                var tf = createParent.textFrames.add({
                                    geometricBounds: [cellTop, cellLeft, cellBottom, cellRight],
                                    strokeWeight: 0,
                                    strokeColor: doc.swatches.item("None"),
                                    fillColor: doc.swatches.item("None")
                                });
                                textFrames.push(tf);
                                items.push(tf);
                            }
                        }
                        // スレッド化：テキストフレームを順番にリンク
                        if (threadText && textFrames.length > 1) {
                            for (var ti = 0; ti < textFrames.length - 1; ti++) {
                                textFrames[ti].nextTextFrame = textFrames[ti + 1];
                            }
                        }
                        // サンプルテキストを流し込み（プレビューレイヤーではスキップ）
                        if ((sampleText || sampleSquareCircle) && textFrames.length > 0 && !targetLayer) {
                            if (sampleSquareCircle) {
                                // □□□□○□□□□●パターンを繰り返し生成
                                var sqUnit = "□□□□○□□□□●";
                                var sqPattern = "";
                                var sqLineCount = 0;
                                var sqNextBreak = Math.floor(Math.random() * 6) + 5; // 5-10
                                for (var sp = 0; sp < 200; sp++) {
                                    sqPattern += sqUnit;
                                    sqLineCount++;
                                    if (sqLineCount >= sqNextBreak) {
                                        sqPattern += "\r";
                                        sqLineCount = 0;
                                        sqNextBreak = Math.floor(Math.random() * 6) + 5;
                                    }
                                }
                                textFrames[0].contents = sqPattern;
                            } else {
                                textFrames[0].contents = "朝、目が覚めると、枕元の端末が静かに光っていた。\r「おはようございます。昨日の記憶を同期しますか？」\r\r　私はしばらくその表示を見つめた。\r　同期ボタンは、もう三日間押していない。\r\r　窓の外には、相変わらず同じ街が広がっている。\r　高層ビルの壁面には、朝のニュースが流れていた。\r\r「政府は本日、記憶バックアップ制度の利用率が国民の92%に達したと発表しました」\r\r　人々は、もうほとんど忘れない。\r　毎晩、脳内の記憶はクラウドに保存される。事故でも病気でも、バックアップから復元できる。\r\r　昨日までの自分を、正確に続きから生きられる。\r\r　便利な世界だ。\r\r　私は端末を伏せて、キッチンへ向かった。\r　コーヒーを淹れていると、壁のディスプレイが自動で点灯する。\r\r「未同期の記憶があります」\r\r　分かっている。\r\r　その記憶のせいだ。\r\r　昨日、私は一人の老人に会った。\r\r　河川敷のベンチで、古い紙の本を読んでいた。\r　今どき珍しい。\r\r「それ、オフラインの本ですか？」\r\r　私が声をかけると、老人は少し笑った。\r\r「そうだよ。記録に残らないものが好きでね」\r\r　意味が分からなかった。\r\r　記録に残らない？\r　そんなもの、価値があるのだろうか。\r\r「今の時代、全部残せるじゃないですか」\r\r　私が言うと、老人は本を閉じて言った。\r\r「だから残らないものが必要なんだ」\r\r　風が吹いた。\r　河川敷の草が揺れる。\r\r「人はね、本当は忘れる生き物なんだよ」\r\r　私は黙っていた。\r\r「忘れるから、また会いたくなる。忘れるから、思い出になる」\r\r　老人は空を見上げた。\r\r「全部残るなら、人生はただのログだ」\r\r　ログ。\r\r　その言葉が、妙に頭に残った。\r\r　家に帰ってから、私は同期を押せなかった。\r\r　もし同期すれば、この会話は永久に保存される。\r　政府のサーバーにも、医療記録にも、私の人生ログにも。\r\r　そしてきっと、忘れられなくなる。\r\r　私は端末をもう一度見る。\r\r「記憶同期を実行しますか？」\r\r　画面の下に、小さく表示されている。\r\r「同期しない記憶は、時間とともに消失する可能性があります」\r\r　それでいい。\r\r　私は河川敷の風を思い出す。\r　老人の声を思い出す。\r\r　でも、きっと少しずつ薄れていく。\r\r　声の高さも。\r　顔の皺も。\r　本の色も。\r\r　いつか曖昧になる。\r\r　それでいいのだと思う。\r\r　私は端末の通知を閉じた。\r\r　しばらくして、端末が静かに言う。\r\r「未同期記憶の自動削除まで、残り23時間」\r\r　窓の外では、ドローンが郵便物を運んでいた。\r　街は今日も、正確に記録されている。\r\r　私はコーヒーを飲みながら、ふと思う。\r\r　もしかしたら、あの老人の顔も。\r　もう、はっきり思い出せない。\r\r　でも、不思議と安心していた。\r\r　その記憶は、私の中だけにある。\r\r　サーバーにも、政府にも、誰のログにも残らない。\r\r　ただ、私の人生のどこかに、少しだけ影響して。\r　そして、静かに消えていく。\r\r　端末の光が消える。\r\r　私は窓を開けた。\r\r　春の風が、部屋に入ってきた。";
                            }
                            // フォントサイズと行送りを適用
                            var story = textFrames[0].parentStory;
                            if (!isNaN(baseFontSize) && baseFontSize > 0) {
                                story.pointSize = baseFontSize;
                            }
                            if (leadingValue === "auto" || leadingValue === "") {
                                story.leading = Leading.AUTO;
                            } else {
                                var lv = parseFloat(leadingValue);
                                if (!isNaN(lv) && lv > 0) {
                                    story.leading = lv;
                                }
                            }
                            try {
                                story.appliedFont = app.fonts.item("ヒラギノ角ゴ Pro W3");
                            } catch (e) {
                                try {
                                    story.appliedFont = app.fonts.item("HiraginoSans-W3");
                                } catch (e2) { }
                            }
                        }
                    }
                }

                // 区切り線の描画（ガター中央に罫線）
                if (!gridOnly && innerStroke && (colCount > 1 || rowCount > 1)) {
                    var dTotalWidth = iRectRight - iRectLeft;
                    var dTotalHeight = iRectBottom - iRectTop;
                    var dCellWidth = (dTotalWidth - colGap * (colCount - 1)) / colCount;
                    var dCellHeight = (dTotalHeight - rowGap * (rowCount - 1)) / rowCount;

                    // 罫線の種類（InDesign組み込みストロークスタイルを使用）
                    var divStrokeStyle = null;
                    var divEndCap = EndCap.BUTT_END_CAP;
                    if (lineType === "dashed") {
                        var dashedNames = ["Dashed (3 and 2)", "破線（3 - 2）", "Dashed (4 and 4)", "破線（4 - 4）"];
                        for (var dn = 0; dn < dashedNames.length; dn++) {
                            try { var tmp = doc.strokeStyles.item(dashedNames[dn]); tmp.name; divStrokeStyle = tmp; break; } catch (e) { }
                        }
                    } else if (lineType === "dotted") {
                        // ドット点線: 線幅と同じ間隔の丸型線端で表現
                        divEndCap = EndCap.ROUND_END_CAP;
                        var dottedNames = ["Japanese Dots", "ドット", "Dotted"];
                        for (var dn2 = 0; dn2 < dottedNames.length; dn2++) {
                            try { var tmp2 = doc.strokeStyles.item(dottedNames[dn2]); tmp2.name; divStrokeStyle = tmp2; break; } catch (e) { }
                        }
                    }

                    // 列の区切り線（縦線）
                    for (var di = 1; di < colCount; di++) {
                        var divX = iRectLeft + di * dCellWidth + (di - 0.5) * colGap;
                        var divLine = createParent.graphicLines.add();
                        divLine.paths[0].entirePath = [[divX, iRectTop], [divX, iRectBottom]];
                        divLine.strokeWeight = divWeight;
                        divLine.strokeColor = doc.swatches.item("Black");
                        if (divStrokeStyle) divLine.strokeType = divStrokeStyle;
                        divLine.endCap = divEndCap;
                        items.push(divLine);
                    }

                    // 行の区切り線（横線）
                    for (var dj = 1; dj < rowCount; dj++) {
                        var divY = iRectTop + dj * dCellHeight + (dj - 0.5) * rowGap;
                        var divLine = createParent.graphicLines.add();
                        divLine.paths[0].entirePath = [[iRectLeft, divY], [iRectRight, divY]];
                        divLine.strokeWeight = divWeight;
                        divLine.strokeColor = doc.swatches.item("Black");
                        if (divStrokeStyle) divLine.strokeType = divStrokeStyle;
                        divLine.endCap = divEndCap;
                        items.push(divLine);
                    }
                }

                // 擬似レイアウトグリッド描画（行位置に横線をページ全幅で描画）
                // 基準: コンテンツ上端 + fontSize で1本目、以降 leading 間隔
                if (showLayoutGrid && (targetLayer || gridOnly)) {
                    var fontSizePt = (!isNaN(baseFontSize) && baseFontSize > 0) ? baseFontSize : 9.5;
                    var leadingPt = parseFloat(leadingValue);
                    if (isNaN(leadingPt) || leadingPt <= 0) leadingPt = fontSizePt * 1.5;
                    // ptをドキュメント単位に変換
                    var charW = fontSizePt / ptPerUnit;
                    var lineH = leadingPt / ptPerUnit;

                    var gridColor = getOrCreateColor(doc, "LayoutGrid", [100, 0, 0, 0]);

                    // ページ全幅に描画、outer上端基準で下端まで
                    var gridLines = [];
                    var gy = outerTop + charW;
                    while (gy <= bottom) {
                        var gLine = page.graphicLines.add();
                        gLine.paths[0].entirePath = [[left, gy], [right, gy]];
                        gLine.strokeColor = gridColor;
                        gLine.strokeWeight = 0.1;
                        gridLines.push(gLine);
                        gy += lineH;
                    }
                    if (gridLines.length > 1) {
                        var gridGroup = page.groups.add(gridLines);
                        items.push(gridGroup);
                    } else if (gridLines.length === 1) {
                        items.push(gridLines[0]);
                    }

                }

                return items;
            } finally {
                if (prevActiveLayer && prevActiveLayer.isValid) {
                    try { doc.activeLayer = prevActiveLayer; } catch (e1) { }
                }
                // ルーラー原点とゼロポイントを元に戻す
                doc.viewPreferences.rulerOrigin = oldRulerOrigin;
                doc.zeroPoint = oldZeroPoint;
                // 垂直単位を元の状態に戻す
                doc.viewPreferences.verticalMeasurementUnits = oldYUnits;
            }
        }

    })();