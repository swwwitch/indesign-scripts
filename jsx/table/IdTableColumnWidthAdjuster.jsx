#target indesign

/*
 * IdTableColumnWidthAdjuster.jsx
 *
 * 選択したセルや表の列幅を、列ごとの個別指定または一括入力でまとめて調整します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdTableColumnWidthAdjuster";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-19";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdTableColumnWidthAdjuster.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableColumnWidthAdjuster.md

// Original idea
// AdjColWidth_221003a.jsx by 照山裕爾（mottainaiDTP）
// https://mottainaidtp.seesaa.net/article/492096133.html

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// ==============================
// UIレイアウトの共通設定 / Shared UI layout
// ==============================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

/**
 * ウィンドウの共通設定を適用する
 * @param {Window} win 対象ウィンドウ
 * @param {number} [spacing] 要素間隔。省略時は WINDOW_SPACING
 * @returns {void}
 */
function setupWindow(win, spacing) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルの共通設定を適用する
 * @param {Panel} panel 対象パネル
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * 行グループの共通設定を適用する（ボタン列など）
 * @param {Group} group 対象グループ
 * @param {string} [alignment] 配置。省略時は "left"
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

(function () {

	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;

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
	var currentLang = getCurrentLang();

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
			apply:  { ja: "適用", en: "Apply" }
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
		}
	};

	/**
	 * ドット区切りキーでラベルを取得する
	 * @param {string} labelKey 例: "dialog.title"
	 * @returns {string} 現在の言語のラベル文字列。見つからない場合はキーをそのまま返す
	 */
	function getLabel(labelKey) {
		var node = LABELS;
		var keyParts = labelKey.split(".");
		for (var i = 0; i < keyParts.length; i++) {
			node = node[keyParts[i]];
			if (!node) return labelKey;
		}
		return node[currentLang] || node.en || labelKey;
	}

	/**
	 * コロン付きラベルを取得する（日本語は全角コロン、英語は半角コロン）
	 * @param {string} labelKey 例: "field.inputMethod"
	 * @returns {string} コロンを付与したラベル文字列
	 */
	function getLabelWithColon(labelKey) {
		return getLabel(labelKey) + (currentLang === "ja" ? "：" : ":");
	}

	/**
	 * 現在プレビュー表示になっているかを判定する
	 * @returns {boolean} プレビュー表示なら true
	 */
	function isPreviewScreenMode() {
		try {
			return app.activeWindow && app.activeWindow.screenMode === ScreenModeOptions.PREVIEW_TO_PAGE;
		} catch (e) {
			return false;
		}
	}

	/**
	 * 標準表示とプレビュー表示を切り替える
	 * @returns {void}
	 */
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

	/**
	 * 画面モードに応じたトグルボタンのラベルを返す
	 * @param {string} currentLang UI 言語
	 * @returns {string} ボタンに表示する文字列
	 */
	function getPreviewToggleButtonLabel(currentLang) {
		return isPreviewScreenMode()
			? (currentLang === "ja" ? "プレビュー" : "Preview")
			: (currentLang === "ja" ? "標準モード" : "Normal Mode");
	}

	/**
	 * トグルボタンのラベルを現在の画面モードに合わせて更新する
	 * @param {Button} btn 対象のボタン
	 * @param {string} currentLang UI 言語
	 * @returns {void}
	 */
	function updatePreviewToggleButtonLabel(btn, currentLang) {
		btn.text = getPreviewToggleButtonLabel(currentLang);
	}

	/* 実行とエラー対策 / Execution and error handling */
	main();
	/**
	 * 列幅調整の処理を開始する
	 * @returns {void}
	 */
	function main() {
		app.doScript("adjustColumnWidths()", ScriptLanguage.JAVASCRIPT, [], UndoModes.fastEntireScript);
	}

	/**
	 * ダイアログを表示して列幅の調整を実行する
	 * @returns {void}
	 */
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

		/**
		 * 実行前の選択状態を復元する
		 * @returns {void}
		 */
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
			getLabel("panel.columnWidths") + (currentLang === "ja" ? "（" + panelUnitLabel + "）" : " (" + panelUnitLabel + ")")
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

		var batchApplyButton = batchRightGroup.add("button", undefined, getLabel("button.apply"));

		var batchHintText = batchInputPanel.add("statictext", undefined, getLabel("hint.batch"));

		var footerGroup = dialog.add("group");
		setupRow(footerGroup, "fill", 8);
		footerGroup.alignChildren = ["left", "center"];
		footerGroup.margins = [0, 10, 0, 0];

		// 左：プレビューモード切替
		var footerLeft = footerGroup.add("group");
		footerLeft.orientation = "row";
		footerLeft.alignment = ["left", "center"];
		footerLeft.alignChildren = ["left", "center"];

		var previewModeButton = footerLeft.add("button", undefined, getPreviewToggleButtonLabel(currentLang));

		// 中央：spacer
		var footerCenter = footerGroup.add("group");
		footerCenter.alignment = ["fill", "center"];

		// 右：ボタン
		/* ボタン行（幅いっぱいには広げない）/ Button row (never stretched to full width) */
		var footerRight = footerGroup.add("group");
		setupRow(footerRight, "right", 8);
		footerRight.alignChildren = ["right", "center"];

		footerRight.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
		footerRight.add("button", undefined, getLabel("button.ok"), { name: "ok" });

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
			updatePreviewToggleButtonLabel(previewModeButton, currentLang);
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

				attachArrowKeyStepper(widthInputs[idx], { forceInteger: false });
				attachArrowKeyStepper(charCountInputs[idx], { forceInteger: true });
				attachArrowKeyStepper(sideInsetInputs[idx], { forceInteger: false });
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

	/**
	 * 選択から対象の表を特定する
	 * @param {Array} selection 選択オブジェクトの配列
	 * @returns {Table|null} 対象の表。特定できない場合は null
	 */
	function resolveTableFromSelection(selection) {
		if (!selection || selection.length === 0) {
			alert(getLabel("alert.selectCellOrTable"));
			return null;
		}

		var node = selection[0];
		if (node == null) {
			alert(getLabel("alert.selectCellOrTable"));
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

		alert(getLabel("alert.selectCellOrTable"));
		return null;
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
		headerWidth.preferredSize.width = 70;
		headerLabelControls.push(headerWidth);

		// var headerMidSpacer = headerRow.add("group");
		// headerMidSpacer.preferredSize.width = 10;

		// 「文字数」
		var headerCharCount = headerRow.add("statictext", undefined, getLabel("header.charCount"));
		headerCharCount.justify = "center";
		headerCharCount.preferredSize.width = 73;
		headerLabelControls.push(headerCharCount);

		// var headerSpacer = headerRow.add("group");
		// headerSpacer.preferredSize.width = 16;

		// 「左右の余白」
		var headerInset = headerRow.add("statictext", undefined, getLabel("header.inset"));
		headerInset.justify = "center";
		headerInset.preferredSize.width = 70;
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
			var widthInput = inputRow.add("edittext", undefined, formatNumber(widthValue));
			widthInput.characters = 5;
			// 行ごとの幅単位ラベルは表示しない / Per-row width unit label is omitted
			// rowLabelRow.push(inputRow.add("statictext", undefined, getLabel("unit.mm")));

			var midSpacer = inputRow.add("group");
			midSpacer.preferredSize.width = 10;

			var charCount = calculateCharCount(widthValue, initialInsetValues[i], fontSizePt);
			var charCountInput = inputRow.add("edittext", undefined, formatNumber(charCount));
			charCountInput.characters = 5;
			// 行ごとの文字数単位ラベルは表示しない / Per-row character-count unit label is omitted
			// rowLabelRow.push(inputRow.add("statictext", undefined, getLabel("unit.characters")));

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
	 * @returns {void}
	 */
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

	/**
	 * 入力欄に上下キーでの増減操作を追加する
	 * @param {EditText} editText 対象の入力欄
	 * @param {object} options 最小値などの追加設定
	 * @returns {void}
	 */
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