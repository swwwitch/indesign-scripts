#target indesign
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;

(function () {

	/* 
	概要 / Overview
	選択中のセル、セル内テキスト、表、またはテーブルを含むテキストフレームを対象に、列幅をまとめて調整する InDesign スクリプトです。
	入力方法は「個別に設定」と「一括入力」を切り替えできます。

	【個別に設定】
	- 指定方法は「幅で指定」または「文字数で指定」を選択可能
	- 列ごとに、幅／文字数換算／左右の余白／自動調整を設定
	- 自動調整は、2行目がある列では2行目が消えるまで列幅を広げ、2行目がない列では文字数ベースで幅を推定
	- 「全列に適用」で、編集中の値を他列へ連動反映
	- 自動調整を ON にした列は、その時点の算出結果を固定し、ユーザーが変更するまで維持

	【一括入力】
	- 幅指定専用
	- 入力形式：スペース区切りまたはカンマ区切り（例：30 50 70 70 / 30, 50, 70, 70）

	【共通仕様】
	- 列幅と左右の余白はドキュメントの単位設定に従う
	- 文字数換算は、表内で最も支配的なフォントサイズを基準に計算
	- 左右の余白は各列内すべてのセルに適用
	- 変更は常に即時反映し、キャンセル時は元に戻す

	*/

	/*
	 * 参考 / References
	 * Based on: AdjColWidth_221003a.jsx by 照山裕爾（mottainaiDTP）
	 * https://mottainaidtp.seesaa.net/article/492096133.html
	 * Error-handling pattern reference: http://cat.adodtp.com/2016/08/23/?p=638
	 * UI ダイアログ、ローカライズ、文字数指定モード、プレビューなどを追加して再構成。
	 * Reworked with UI dialog, localization, character-count mode, preview support, and related enhancements.
	 */

	// =========================================
	// バージョンとローカライズ / Version and localization
	// =========================================

	var SCRIPT_VERSION = "v1.1.0";

	function getCurrentLang() {
		return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
	}
	var lang = getCurrentLang();

	/* 日英ラベル定義 / Japanese-English label definitions */
	var LABELS = {
		dialogTitle: {
			ja: "列幅の調整",
			en: "Adjust Column Widths"
		},
		panelBasic: {
			ja: "指定方法",
			en: "Sizing Method"
		},
		panelPerColumn: {
			ja: "個別に設定",
			en: "Per-Column Settings"
		},
		labelCalculationBasis: {
			ja: "指定方法",
			en: "Sizing Method"
		},
		labelInputMethod: {
			ja: "入力方法",
			en: "Input Method"
		},
		inputMethodPerColumn: {
			ja: "個別に設定",
			en: "Per-Column Input"
		},
		inputMethodBatch: {
			ja: "一括入力",
			en: "Batch Input"
		},
		modeAbsolute: {
			ja: "幅で指定",
			en: "Set by Width"
		},
		modeCharacterBased: {
			ja: "文字数で指定",
			en: "Set by Character Count"
		},
		checkboxUnify: {
			ja: "全列に適用",
			en: "Apply to All Columns"
		},
		panelColumnWidths: {
			ja: "各列の設定",
			en: "Column Settings"
		},
		unitMm: {
			ja: "mm",
			en: "mm"
		},
		unitCharacters: {
			ja: "文字",
			en: "chars"
		},
		columnSuffix: {
			ja: "列目",
			en: "Col"
		},
		headerColumn: {
			ja: "列",
			en: "Col"
		},
		headerWidth: {
			ja: "幅",
			en: "Width"
		},
		headerCharCount: {
			ja: "文字数",
			en: "Character Count"
		},
		headerInset: {
			ja: "左右の余白",
			en: "Left/Right Inset"
		},
		headerAutoFit: {
			ja: "自動調整",
			en: "Auto Fit"
		},
		panelBatch: {
			ja: "一括入力",
			en: "Batch Input"
		},
		buttonApply: {
			ja: "適用",
			en: "Apply"
		},
		batchHint: {
			ja: "入力形式：15 20 34 10 または 15, 20, 34, 10",
			en: "Example: Enter column widths as 15 20 34 10 or 15, 20, 34, 10"
		},
		checkboxPreview: {
			ja: "プレビュー",
			en: "Preview"
		},
		buttonCancel: {
			ja: "キャンセル",
			en: "Cancel"
		},
		buttonOk: {
			ja: "OK",
			en: "OK"
		},
		alertSelectCellOrTable: {
			ja: "セルまたは表を選択してから実行してください",
			en: "Select a cell or table before running this script."
		},
		alertBatchInvalidValue: {
			ja: "一括入力に無効な値が含まれています",
			en: "Batch input contains invalid values."
		},
		alertBatchTooManyValues: {
			ja: "入力数が列数を超えています",
			en: "Too many values for the number of columns."
		},
		alertInvalidNumber: {
			ja: "数値を入力してください",
			en: "Enter a valid number."
		},
		alertNegativeWidth: {
			ja: "幅には 0 以上の数値を入力してください",
			en: "Width must be 0 or greater."
		},
		alertNegativeCharCount: {
			ja: "文字数には 0 以上の数値を入力してください",
			en: "Character count must be 0 or greater."
		},
		alertNegativeInset: {
			ja: "左右の余白には 0 以上の数値を入力してください",
			en: "Left/right inset must be 0 or greater."
		},
		alertInsetTooLarge: {
			ja: "左右の余白が大きすぎます。内容幅が 0 以下になります",
			en: "Left/right inset is too large. Content width would become 0 or less."
		}
	};

	function L(key) {
		return (LABELS[key] && LABELS[key][lang]) ? LABELS[key][lang] : key;
	}

	function labelText(key) {
		return L(key) + (lang === 'ja' ? '：' : ':');
	}

	function isPreviewScreenMode() {
		try {
			return app.activeWindow && app.activeWindow.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE;
		} catch (e) {
			return false;
		}
	}

	function togglePreviewScreenMode() {
		try {
			var w = app.activeWindow;
			if (!w) return;
			if (w.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE) {
				w.screenMode = ScreenModeOptions.PREVIEW_OFF;
			} else {
				w.screenMode = ScreenModeOptions.PREVIEW_TO_PAGE;
			}
		} catch (e) { }
	}

	function getPreviewToggleButtonLabel(lang) {
		return isPreviewScreenMode()
			? (lang === "ja" ? "プレビュー" : "Preview")
			: (lang === "ja" ? "標準モード" : "Normal Mode");
	}

	function updatePreviewToggleButtonLabel(btn, lang) {
		btn.text = getPreviewToggleButtonLabel(lang);
	}

	/* 実行とエラー対策 / Execution and error handling */
	main();
	function main() {
		app.doScript("adjustColumnWidths()", ScriptLanguage.JAVASCRIPT, [], UndoModes.fastEntireScript);
	}

	function adjustColumnWidths() {
		var selection = app.activeDocument.selection;
		var targetTable = resolveTableFromSelection(selection);
		if (!targetTable) return;

		// 選択を記憶し、ダイアログ中はハイライトを消す / Save selection and hide highlight while dialog is open
		var savedSelection = [];
		for (var selectionIndex = 0; selectionIndex < selection.length; selectionIndex++) {
			savedSelection.push(selection[selectionIndex]);
		}
		try {
			app.select(NothingEnum.NOTHING);
		} catch (e) { }

		function restoreSelection() {
			try {
				if (savedSelection.length > 0) {
					app.selection = savedSelection;
				}
			} catch (e) { }
		}

		var columnCount = targetTable.columns.length;
		var originalWidths = getColumnWidths(targetTable);
		var originalInsets = getOriginalInsets(targetTable);

		var dominantFontSizePt = findDominantFontSize(targetTable);
		fontSizePtForValidationCache = dominantFontSizePt;

		var dialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
		dialog.orientation = "column";
		dialog.alignChildren = "fill";
		dialog.margins = 15;
		dialog.spacing = 10;

		var inputMethodOuterGroup = dialog.add("group");
		inputMethodOuterGroup.orientation = "row";
		inputMethodOuterGroup.alignment = "center";
		inputMethodOuterGroup.alignChildren = ["center", "center"];
		inputMethodOuterGroup.margins = [0, 3, 0, 10];

		var inputMethodGroup = inputMethodOuterGroup.add("group");
		inputMethodGroup.orientation = "row";
		inputMethodGroup.spacing = 12;
		var perColumnRadio = inputMethodGroup.add("radiobutton", undefined, L("inputMethodPerColumn"));
		var batchModeRadio = inputMethodGroup.add("radiobutton", undefined, L("inputMethodBatch"));
		perColumnRadio.value = true;


		var panelUnitLabel = getRulerUnitString();
		var perColumnPanel = dialog.add("panel", undefined, L("panelPerColumn"));
		perColumnPanel.orientation = "column";
		perColumnPanel.alignChildren = "fill";
		perColumnPanel.margins = [15, 20, 15, 10];
		perColumnPanel.spacing = 10;

		var basicSettingsPanel = perColumnPanel.add("panel", undefined, L("panelBasic"));
		basicSettingsPanel.orientation = "column";
		basicSettingsPanel.alignChildren = "left";
		basicSettingsPanel.margins = [15, 20, 15, 10];
		// basicSettingsPanel.spacing = 15; // 必要に応じて有効化 / Enable if needed

		var modeGroup = basicSettingsPanel.add("group");
		modeGroup.orientation = "column";
		modeGroup.alignChildren = "left";
		modeGroup.spacing = 4;
		var absoluteRadio = modeGroup.add("radiobutton", undefined, L("modeAbsolute"));
		var characterBasedRadio = modeGroup.add("radiobutton", undefined, L("modeCharacterBased"));
		absoluteRadio.value = true;


		var columnSettingsPanel = perColumnPanel.add(
			"panel",
			undefined,
			L("panelColumnWidths") + (lang === "ja" ? "（" + panelUnitLabel + "）" : " (" + panelUnitLabel + ")")
		);
		columnSettingsPanel.orientation = "column";
		columnSettingsPanel.alignChildren = "left";
		columnSettingsPanel.margins = [15, 20, 15, 10];
		columnSettingsPanel.spacing = 6;
		// 全列に適用 / Apply to all columns
		var unifyCheckbox = columnSettingsPanel.add("checkbox", undefined, L("checkboxUnify"));

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

		var batchInputPanel = dialog.add("panel", undefined, L("panelBatch"));
		batchInputPanel.orientation = "column";
		batchInputPanel.alignChildren = "fill";
		batchInputPanel.margins = [15, 20, 15, 10];
		batchInputPanel.spacing = 6;

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

		var batchApplyButton = batchRightGroup.add("button", undefined, L("buttonApply"));

		var batchHintText = batchInputPanel.add("statictext", undefined, L("batchHint"));

		var footerGroup = dialog.add("group");
		footerGroup.orientation = "row";
		footerGroup.alignment = "fill";
		footerGroup.alignChildren = ["left", "center"];
		footerGroup.margins = [0, 10, 0, 0];

		// 左：プレビューモード切替
		var footerLeft = footerGroup.add("group");
		footerLeft.orientation = "row";
		footerLeft.alignment = ["left", "center"];
		footerLeft.alignChildren = ["left", "center"];

		var previewModeButton = footerLeft.add("button", undefined, getPreviewToggleButtonLabel(lang));

		// 中央：spacer
		var footerCenter = footerGroup.add("group");
		footerCenter.alignment = ["fill", "center"];

		// 右：ボタン
		var footerRight = footerGroup.add("group");
		footerRight.orientation = "row";
		footerRight.alignment = ["right", "center"];
		footerRight.alignChildren = ["right", "center"];

		footerRight.add("button", undefined, L("buttonCancel"), { name: "cancel" });
		footerRight.add("button", undefined, L("buttonOk"), { name: "ok" });

		var primaryInputMode = "absolute";
		var currentInputMethod = "perColumn";
		var isPreviewCurrentlyApplied = false;

		function setInputMode(mode) {
			primaryInputMode = mode;
			refreshControlStates();
		}

		function setInputMethod(method) {
			currentInputMethod = method;
			refreshControlStates();
		}

		function refreshControlStates() {
			var isPerColumnMode = (currentInputMethod == "perColumn");
			columnSettingsPanel.visible = true;
			batchInputPanel.visible = true;

			setControlEnabled(batchInput, !isPerColumnMode);
			setControlEnabled(batchApplyButton, !isPerColumnMode);
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

		function restoreOriginalColumnSettings() {
			restoreColumnWidths(targetTable, originalWidths);
			restoreColumnInsets(targetTable, originalInsets);
			isPreviewCurrentlyApplied = false;
		}

		// 高速自動調整: 文字数推定による列内容幅計算
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
			try { widthInputs[colIdx].parent.layout.layout(true); } catch (e) { }
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

		previewModeButton.onClick = function () {
			togglePreviewScreenMode();
			updatePreviewToggleButtonLabel(previewModeButton, lang);
		};

		for (var i = 0; i < columnCount; i++) {
			(function (idx) {
				widthInputs[idx].onChanging = function () {
					updateCharCountFromWidth(widthInputs[idx], charCountInputs[idx], sideInsetInputs[idx], dominantFontSizePt);
					if (unifyCheckbox.value) syncUnifiedInputsFromSource(widthInputs, charCountInputs, sideInsetInputs, idx, dominantFontSizePt, "absolute", autoFitCheckboxes);
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
					if (unifyCheckbox.value) syncUnifiedInputsFromSource(widthInputs, charCountInputs, sideInsetInputs, idx, dominantFontSizePt, "characterBased", autoFitCheckboxes);
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
					if (unifyCheckbox.value) syncUnifiedInputsFromSource(widthInputs, charCountInputs, sideInsetInputs, idx, dominantFontSizePt, primaryInputMode, autoFitCheckboxes);
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

				attachArrowKeyStepper(widthInputs[idx], { forceInteger: false });
				attachArrowKeyStepper(charCountInputs[idx], { forceInteger: true });
				attachArrowKeyStepper(sideInsetInputs[idx], { forceInteger: false });
			})(i);
		}

		function applyBatchInput() {
			var values = parseBatchInput(batchInput.text);
			if (values.length == 0) return;

			// 入力検証：無効な値があれば中断 / Validation: reject if any invalid value exists
			for (var i = 0; i < values.length; i++) {
				if (values[i] == null) {
					alert(L("alertBatchInvalidValue"));
					return;
				}
			}

			// 入力検証：列数を超える場合は中断 / Validation: reject if too many values
			if (values.length > columnCount) {
				alert(L("alertBatchTooManyValues"));
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
		batchApplyButton.onClick = applyBatchInput;

		setInputMode("absolute");
		setInputMethod("perColumn");
		widthInputs[0].active = true;

		var dialogResult = dialog.show();

		if (dialogResult != 1) {
			if (isPreviewCurrentlyApplied) restoreOriginalColumnSettings();
			restoreSelection();
			return;
		}

		applyColumnSettings();
		restoreSelection();
	}

	// =========================================
	// 表・選択ヘルパー / Table and selection helpers
	// =========================================

	function resolveTableFromSelection(selection) {
		if (!selection || selection.length === 0) {
			alert(L("alertSelectCellOrTable"));
			return null;
		}

		var node = selection[0];
		if (node == null) {
			alert(L("alertSelectCellOrTable"));
			return null;
		}

		while (node) {
			try {
				if (node instanceof Table) return node;
				if (node instanceof Cell) return node.parent;
				if (node instanceof TextFrame) {
					if (node.tables.length > 0) {
						return node.tables[0];
					}
				}
				node = node.parent;
			} catch (e) {
				break;
			}
		}

		alert(L("alertSelectCellOrTable"));
		return null;
	}

	function getColumnWidths(table) {
		var widths = [];
		for (var i = 0; i < table.columns.length; i++) {
			widths.push(table.columns[i].width);
		}
		return widths;
	}

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

	function insetsToInitialInsetValues(originalInsets) {
		var insetValues = [];
		for (var i = 0; i < originalInsets.length; i++) {
			var leftInset = originalInsets[i].length > 0 ? originalInsets[i][0].left : 0;
			insetValues.push(rulerValueToInputUnit(leftInset));
		}
		return insetValues;
	}

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

	function restoreColumnWidths(table, originalWidths) {
		for (var i = 0; i < originalWidths.length; i++) {
			try { table.columns[i].width = originalWidths[i]; } catch (e) { }
		}
	}

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

	function buildColumnSettingsControls(parent, columnCount, originalWidths, initialInsetValues, fontSizePt) {
		var widthInputFields = [];
		var charCountInputFields = [];
		var sideInsetInputFields = [];
		var autoFitToggleCheckboxes = [];
		var rowLabelControls = [];
		var headerLabelControls = [];

		// ヘッダー行を追加 / Insert header row
		var headerRow = parent.add("group");
		headerRow.orientation = "row";
		headerRow.spacing = 6;
		headerRow.margins = [0, 2, 0, 6];

		// 「列」
		var headerColumn = headerRow.add("statictext", undefined, L("headerColumn"));
		headerColumn.justify = "center";
		headerColumn.preferredSize.width = 14;
		headerLabelControls.push(headerColumn);

		//　「幅」
		var headerWidth = headerRow.add("statictext", undefined, L("headerWidth"));
		headerWidth.justify = "center";
		headerWidth.preferredSize.width = 70;
		headerLabelControls.push(headerWidth);

		// var headerMidSpacer = headerRow.add("group");
		// headerMidSpacer.preferredSize.width = 10;

		// 「文字数」
		var headerCharCount = headerRow.add("statictext", undefined, L("headerCharCount"));
		headerCharCount.justify = "center";
		headerCharCount.preferredSize.width = 73;
		headerLabelControls.push(headerCharCount);

		// var headerSpacer = headerRow.add("group");
		// headerSpacer.preferredSize.width = 16;

		// 「左右の余白」
		var headerInset = headerRow.add("statictext", undefined, L("headerInset"));
		headerInset.justify = "center";
		headerInset.preferredSize.width = 70;
		headerLabelControls.push(headerInset);

		// var headerAutoFitSpacer = headerRow.add("group");
		// headerAutoFitSpacer.preferredSize.width = 16;

		// 「自動調整」
		var headerAutoFit = headerRow.add("statictext", undefined, L("headerAutoFit"));
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
			var widthInput = inputRow.add("edittext", undefined, formatNumber(widthValue));
			widthInput.characters = 5;
			// 行ごとの幅単位ラベルは表示しない / Per-row width unit label is omitted
			// rowLabelRow.push(inputRow.add("statictext", undefined, L("unitMm")));

			var midSpacer = inputRow.add("group");
			midSpacer.preferredSize.width = 10;

			var charCount = calculateCharCount(widthValue, initialInsetValues[i], fontSizePt);
			var charCountInput = inputRow.add("edittext", undefined, formatNumber(charCount));
			charCountInput.characters = 5;
			// 行ごとの文字数単位ラベルは表示しない / Per-row character-count unit label is omitted
			// rowLabelRow.push(inputRow.add("statictext", undefined, L("unitCharacters")));

			var rowSpacer = inputRow.add("group");
			rowSpacer.preferredSize.width = 16;

			var sideInsetInput = inputRow.add("edittext", undefined, formatNumber(initialInsetValues[i]));
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

	function updateCharCountFromWidth(widthInput, charCountInput, sideInsetInput, fontSizePt) {
		var widthValue = parseFloat(widthInput.text);
		var insetValue = parseFloat(sideInsetInput.text);
		var cc = calculateCharCount(widthValue, insetValue, fontSizePt);
		charCountInput.text = formatNumber(cc);
	}

	function updateWidthFromCharCount(widthInput, charCountInput, sideInsetInput, fontSizePt) {
		var cc = parseFloat(charCountInput.text);
		var insetValue = parseFloat(sideInsetInput.text);
		var widthValue = calculateWidthFromCharCount(cc, insetValue, fontSizePt);
		widthInput.text = formatNumber(widthValue);
	}

	function syncWidthAndCharCountFromInset(widthInput, charCountInput, sideInsetInput, fontSizePt, primaryInputMode) {
		if (primaryInputMode == "absolute") {
			updateCharCountFromWidth(widthInput, charCountInput, sideInsetInput, fontSizePt);
		}
		else {
			updateWidthFromCharCount(widthInput, charCountInput, sideInsetInput, fontSizePt);
		}
	}

	function syncUnifiedInputsFromSource(widthInputs, charCountInputs, sideInsetInputs, sourceIdx, fontSizePt, primaryInputMode, autoFitCheckboxes) {
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


	function setControlEnabled(control, enabled) {
		try {
			control.enabled = enabled;
		} catch (e) { }
		setLabelDimmed(control, !enabled);
	}

	function setLabelDimmed(control, dim) {
		try {
			var g = control.graphics;
			var rgb = dim ? [0.55, 0.55, 0.55] : [0, 0, 0];
			g.foregroundColor = g.newPen(g.PenType.SOLID_COLOR, rgb, 1);
		} catch (e) { }
	}

	function attachArrowKeyStepper(editText, options) {
		// ↑↓キーと修飾キーで数値を増減 / Adjust numeric values with arrow keys and modifier keys
		editText.addEventListener("keydown", function (event) {
			var value = Number(editText.text);
			if (isNaN(value)) return;

			var keyboard = ScriptUI.environment.keyboardState;
			var delta = 1;
			var forceInteger = options && options.forceInteger;

			if (keyboard.shiftKey) {
				delta = 10;
				if (event.keyName == "Up") {
					value = Math.ceil((value + 1) / delta) * delta;
					event.preventDefault();
				} else if (event.keyName == "Down") {
					value = Math.floor((value - 1) / delta) * delta;
					if (value < 0) value = 0;
					event.preventDefault();
				}
			} else if (keyboard.altKey) {
				delta = 0.1;
				if (event.keyName == "Up") {
					value += delta;
					event.preventDefault();
				} else if (event.keyName == "Down") {
					value -= delta;
					event.preventDefault();
				}
			} else {
				delta = 1;
				if (event.keyName == "Up") {
					value += delta;
					event.preventDefault();
				} else if (event.keyName == "Down") {
					value -= delta;
					if (value < 0) value = 0;
					event.preventDefault();
				}
			}

			var hadDecimal = (String(editText.text).indexOf(".") !== -1);
			if (forceInteger && !keyboard.altKey) {
				value = Math.round(value);
			} else if (keyboard.altKey || hadDecimal || value % 1 !== 0) {
				value = Math.round(value * 10) / 10;
			} else {
				value = Math.round(value);
			}

			editText.text = value;
			if (typeof editText.onChanging === "function") editText.onChanging();
			if (typeof editText.onChange === "function") editText.onChange();
		});
	}

	// =========================================
	// フォントサイズ / Font size
	// =========================================

	var fontSizePtForValidationCache = null;

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

	function ptToQ(pt) { return pt * 25.4 / 18; }
	function ptToMm(pt) { return pt * 25.4 / 72; }
	function mmToPt(mm) { return mm * 72 / 25.4; }

	function getRulerUnitString() {
		var u = app.activeDocument.viewPreferences.horizontalMeasurementUnits;
		if (u == MeasurementUnits.MILLIMETERS) return "mm";
		if (u == MeasurementUnits.CENTIMETERS) return "cm";
		if (u == MeasurementUnits.POINTS) return "pt";
		if (u == MeasurementUnits.INCHES || u == MeasurementUnits.INCHES_DECIMAL) return "in";
		if (u == MeasurementUnits.PICAS) return "pc";
		return "pt";
	}

	function rulerValueToInputUnit(value) {
		var unit = getRulerUnitString();
		try { return new UnitValue(value, unit).as(unit); } catch (e) { return value; }
	}

	function inputUnitToRulerValue(value) {
		return value;
	}

	function inputUnitToPt(value) {
		var unit = getRulerUnitString();
		if (unit == "pt") return value;
		try { return new UnitValue(value, unit).as("pt"); } catch (e) { return value; }
	}

	function ptToInputUnit(value) {
		var unit = getRulerUnitString();
		if (unit == "pt") return value;
		try { return new UnitValue(value, "pt").as(unit); } catch (e) { return value; }
	}

	function calculateCharCount(widthValue, insetValue, fontSizePt) {
		if (!fontSizePt || widthValue == null || isNaN(widthValue)) return null;
		var inset = (insetValue == null || isNaN(insetValue)) ? 0 : insetValue;
		var contentValue = widthValue - 2 * inset;
		if (contentValue <= 0) return null;
		return inputUnitToPt(contentValue) / fontSizePt;
	}

	function calculateWidthFromCharCount(charCount, insetValue, fontSizePt) {
		if (!fontSizePt || charCount == null || isNaN(charCount)) return null;
		if (charCount < 0) return null;
		var inset = (insetValue == null || isNaN(insetValue)) ? 0 : insetValue;
		return ptToInputUnit(charCount * fontSizePt) + 2 * inset;
	}

	function formatNumber(value) {
		if (value == null || isNaN(value)) return "";
		return String(Math.round(value * 100) / 100);
	}

	function validatePerColumnRow(widthInput, charCountInput, sideInsetInput, primaryInputMode) {
		var widthValue = parseFloat(widthInput.text);
		var charValue = parseFloat(charCountInput.text);
		var insetValue = parseFloat(sideInsetInput.text);

		if (sideInsetInput.text !== "") {
			if (isNaN(insetValue)) return { ok: false, message: L("alertInvalidNumber"), focus: sideInsetInput };
			if (insetValue < 0) return { ok: false, message: L("alertNegativeInset"), focus: sideInsetInput };
		}

		if (primaryInputMode == "absolute") {
			if (widthInput.text !== "") {
				if (isNaN(widthValue)) return { ok: false, message: L("alertInvalidNumber"), focus: widthInput };
				if (widthValue < 0) return { ok: false, message: L("alertNegativeWidth"), focus: widthInput };
			}
			if (!isNaN(widthValue) && !isNaN(insetValue) && (widthValue - 2 * insetValue) <= 0) {
				return { ok: false, message: L("alertInsetTooLarge"), focus: sideInsetInput };
			}
		}
		else {
			if (charCountInput.text !== "") {
				if (isNaN(charValue)) return { ok: false, message: L("alertInvalidNumber"), focus: charCountInput };
				if (charValue < 0) return { ok: false, message: L("alertNegativeCharCount"), focus: charCountInput };
			}
			if (!isNaN(charValue) && !isNaN(insetValue)) {
				var calculatedWidth = calculateWidthFromCharCount(charValue, insetValue, fontSizePtForValidationCache);
				if (calculatedWidth == null || (calculatedWidth - 2 * insetValue) <= 0) {
					return { ok: false, message: L("alertInsetTooLarge"), focus: sideInsetInput };
				}
			}
		}

		return { ok: true };
	}

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