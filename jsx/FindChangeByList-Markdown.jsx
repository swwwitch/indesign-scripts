#target indesign

// FindChangeByList-Markdown.jsx
// Adobe InDesign に付属する FindChangeByList.jsx をベースに、Markdown記法の検索と段落スタイル適用向けに改変したスクリプト。
// Original script is based on Adobe InDesign's bundled FindChangeByList.jsx.
// memo https://note.com/dtp_tranist/n/n8c0211d92c96

(function () {

var SCRIPT_VERSION = "v1.0.1";

function getCurrentLang() {
	return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
	dialogTitle: { ja: "Markdown → 段落スタイル変換", en: "Markdown to Paragraph Styles" },
	progressTitle: { ja: "変換中…", en: "Converting…" },
	progressLabel: { ja: "Markdown → 段落スタイル変換", en: "Markdown to Paragraph Styles" },
	undoName: { ja: "Markdown → 段落スタイル変換", en: "Markdown to Paragraph Styles" },
	// スコープ
	scopeLabel: { ja: "スコープ", en: "Scope" },
	scopeDoc: { ja: "ドキュメント", en: "Document" },
	scopeStory: { ja: "ストーリー", en: "Story" },
	scopeSelection: { ja: "選択範囲", en: "Selection" },
	// ソース形式
	sourceLabel: { ja: "形式", en: "Format" },
	sourceHTML: { ja: "HTML", en: "HTML" },
	sourceMSWord: { ja: "MS Word", en: "MS Word" },
	// パネル
	panelHeadings: { ja: "基本段落と見出し", en: "Paragraph & Headings" },
	panelList: { ja: "リスト", en: "List" },
	panelOptions: { ja: "オプション", en: "Options" },
	panelCode: { ja: "ソースコード", en: "Source Code" },
	panelImage: { ja: "画像", en: "Image" },
	panelTable: { ja: "表組み", en: "Table" },
	panelCharStyle: { ja: "文字スタイル", en: "Character Styles" },
	panelSpace: { ja: "スペース", en: "Space" },
	panelCleanup: { ja: "クリーンアップ", en: "Cleanup" },
	// 基本段落と見出し
	labelParagraph: { ja: "段落", en: "Paragraph" },
	labelHeading1: { ja: "見出し1", en: "Heading 1" },
	labelHeading2: { ja: "見出し2", en: "Heading 2" },
	labelHeading3: { ja: "見出し3", en: "Heading 3" },
	labelHeading4: { ja: "見出し4", en: "Heading 4" },
	labelHeading5: { ja: "見出し5", en: "Heading 5" },
	labelHeading6: { ja: "見出し6", en: "Heading 6" },
	// リスト
	labelBullet: { ja: "箇条書き", en: "Bullet List" },
	labelBulletSub: { ja: "箇条書き（サブ）", en: "Bullet List (Sub)" },
	labelNumbered: { ja: "番号リスト", en: "Numbered List" },
	// オプション
	labelBlockquote: { ja: "引用", en: "Blockquote" },
	labelCaption: { ja: "キャプション", en: "Caption" },
	// ソースコード
	labelCodeBlock: { ja: "コードブロック", en: "Code Block" },
	labelCodeInline: { ja: "インライン", en: "Inline" },
	// 画像
	labelImage: { ja: "画像", en: "Image" },
	// 表組み
	labelCell: { ja: "セル", en: "Cell" },
	labelHeaderCell: { ja: "見出しセル", en: "Header Cell" },
	// 文字スタイル
	labelBold: { ja: "太字", en: "Bold" },
	labelEmphasis: { ja: "強調", en: "Emphasis" },
	labelLink: { ja: "リンク", en: "Link" },
	// スペース
	labelTrimLead: { ja: "行頭のスペースを削除", en: "Remove leading spaces" },
	labelTrimTrail: { ja: "行末のスペースを削除", en: "Remove trailing spaces" },
	labelNbsp: { ja: "&nbsp; を半角スペースに変換", en: "Convert &nbsp; to space" },
	// クリーンアップ
	labelTrimReturn: { ja: "連続する空行を削除", en: "Remove consecutive blank lines" },
	labelComment: { ja: "コメントアウトを削除", en: "Remove HTML comments" },
	labelHr: { ja: "水平線を削除", en: "Remove horizontal rules" },
	labelGeta: { ja: "〓を削除", en: "Remove 〓 marks" },
	// ボタン
	btnCancel: { ja: "キャンセル", en: "Cancel" },
	btnRun: { ja: "実行", en: "Run" },
	// メッセージ
	msgNoDoc: { ja: "ドキュメントが開かれていません。", en: "No documents are open." },
	msgNoSelection: { ja: "ストーリーまたは選択範囲で実行するには、テキストを選択してください。", en: "To run on Story or Selection, select some text first." },
	msgInvalidSelection: { ja: "選択範囲モードではテキストを選択してください。ストーリーモードではテキストまたはテキストフレームを選択してください。", en: "For Selection mode, select text. For Story mode, select text or a text frame." },
	msgNoTarget: { ja: "変換対象が選択されていません。", en: "No conversion targets selected." },
	msgNoMatch: { ja: "該当するMarkdown記法は見つかりませんでした。", en: "No matching Markdown syntax was found." },
	// 追加ラベル
	resultSep: { ja: "：", en: ": " },
	resultCountSuffix: { ja: " 件", en: " items" }
};

function L(key) {
	return LABELS[key][lang] || LABELS[key]["en"];
}

function main() {
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
	if (app.documents.length == 0) {
		alert(L("msgNoDoc"));
		return;
	}
	showDialog();
}

// ドロップダウンの項目を名前で選択するヘルパー
function selectDropdownByName(dd, name) {
	for (var i = 0; i < dd.items.length; i++) {
		if (dd.items[i].text == name) {
			dd.selection = i;
			return;
		}
	}
}

// ソース形式ごとのデフォルトスタイル名
var STYLE_PRESETS = {
	html: {
		paragraph: "p", h1: "h1", h2: "h2", h3: "h3", h4: "h4", h5: "h5", h6: "h6",
		bullet: "ul-li", bulletSub: "ul-ul-li", numbered: "ol-li",
		blockquote: "p.blockquote", caption: "p.caption",
		codeBlock: "p.code", codeInline: "code",
		image: "p.img", cell: "td", headerCell: "th",
		bold: "strong-bold", emphasis: "em-i-marker", link: "link"
	},
	msword: {
		paragraph: "Paragraph", h1: "Header1", h2: "Header2", h3: "Header3", h4: "Header4", h5: "Header4", h6: "Header4",
		bullet: "BulList > first", bulletSub: "BulList > BulList > first", numbered: "NumList > first",
		blockquote: "Blockquote > Paragraph", caption: "Caption",
		codeBlock: "CodeBlock", codeInline: "Code",
		image: "Figure", cell: "TablePar", headerCell: "TablePar > TableHeader",
		bold: "Bold", emphasis: "Italic", link: "Link"
	}
};

// ドロップダウン生成ヘルパー：preferred＋全プリセット候補を先頭にまとめたリストを返す
// presetKey: STYLE_PRESETS内のキー名（例: "paragraph", "h1"）。省略可。
function addStyleDropdown(parent, names, preferred, width, presetKey) {
	// preferred＋各プリセットの候補名を収集（重複排除）
	var candidates = [preferred];
	if (presetKey) {
		for (var p in STYLE_PRESETS) {
			if (STYLE_PRESETS.hasOwnProperty(p) && STYLE_PRESETS[p][presetKey]) {
				var cand = STYLE_PRESETS[p][presetKey];
				var dup = false;
				for (var c = 0; c < candidates.length; c++) {
					if (candidates[c] == cand) { dup = true; break; }
				}
				if (!dup) candidates.push(cand);
			}
		}
	}

	// candidatesを先頭に、残りのドキュメントスタイルを後ろに
	var sorted = [];
	for (var c = 0; c < candidates.length; c++) {
		sorted.push(candidates[c]);
	}
	for (var i = 0; i < names.length; i++) {
		var exists = false;
		for (var c = 0; c < candidates.length; c++) {
			if (names[i] == candidates[c]) { exists = true; break; }
		}
		if (!exists) sorted.push(names[i]);
	}

	var dd = parent.add("dropdownlist", [0, 0, width, 24], sorted);
	dd.selection = 0;
	return dd;
}

function canUseAsTextTarget(item) {
	if (!item) return false;
	try {
		if (typeof item.changeGrep === "function" && typeof item.changeText === "function") return true;
	} catch (e) { }
	return false;
}

function resolveStoryTargetFromSelection(item) {
	if (!item) return null;
	try {
		if (item.hasOwnProperty("parentStory") && item.parentStory != null) return item.parentStory;
	} catch (e) { }
	try {
		if (item.hasOwnProperty("texts") && item.texts.length > 0 && item.texts[0].parentStory != null) return item.texts[0].parentStory;
	} catch (e) { }
	return null;
}

function showDialog() {
	var doc = app.documents.item(0);

	// ドキュメントのスタイル名一覧を取得
	var paraStyleNames = [];
	var allPara = doc.allParagraphStyles;
	for (var i = 0; i < allPara.length; i++) {
		paraStyleNames.push(allPara[i].name);
	}
	var charStyleNames = [];
	var allChar = doc.allCharacterStyles;
	for (var i = 0; i < allChar.length; i++) {
		charStyleNames.push(allChar[i].name);
	}

	var w = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
	w.orientation = "column";
	w.alignChildren = ["fill", "top"];

	// ── スコープ・ソース形式（中央配置、内部左揃え） ──
	var topOuter = w.add("group");
	topOuter.alignment = ["center", "top"];
	topOuter.orientation = "column";
	topOuter.alignChildren = ["left", "top"];

	var labelWidth = 80;

	var scopeGroup = topOuter.add("group");
	scopeGroup.add("statictext", [0, 0, labelWidth, 20], L("scopeLabel") + L("resultSep"));
	var rbDoc = scopeGroup.add("radiobutton", undefined, L("scopeDoc"));
	var rbStory = scopeGroup.add("radiobutton", undefined, L("scopeStory"));
	var rbSel = scopeGroup.add("radiobutton", undefined, L("scopeSelection"));
	rbStory.value = true;

	var sourceGroup = topOuter.add("group");
	sourceGroup.add("statictext", [0, 0, labelWidth, 20], L("sourceLabel") + L("resultSep"));
	var rbHTML = sourceGroup.add("radiobutton", undefined, L("sourceHTML"));
	var rbMSWord = sourceGroup.add("radiobutton", undefined, L("sourceMSWord"));
	rbHTML.value = true;

	// ── 2カラム ──
	var columns = w.add("group");
	columns.orientation = "row";
	columns.alignChildren = ["fill", "top"];

	var colL = columns.add("group");
	colL.orientation = "column";
	colL.alignChildren = ["fill", "top"];

	var colR = columns.add("group");
	colR.orientation = "column";
	colR.alignChildren = ["fill", "top"];

	// ━━ 左列 ━━

	// ── 基本段落と見出し ──
	var headPanel = colL.add("panel", undefined, L("panelHeadings"));
	headPanel.alignChildren = ["fill", "top"];
	headPanel.margins = [15, 20, 15, 10];

	// 段落
	var rowP = headPanel.add("group");
	var chkP = rowP.add("checkbox", undefined, "");
	chkP.value = true;
	rowP.add("statictext", [0, 0, 120, 20], L("labelParagraph"));
	var styleP = addStyleDropdown(rowP, paraStyleNames, "p", 160, "paragraph");

	var headings = [
		{ label: L("labelHeading1"), md: "#", style: "h1", key: "h1" },
		{ label: L("labelHeading2"), md: "##", style: "h2", key: "h2" },
		{ label: L("labelHeading3"), md: "###", style: "h3", key: "h3" },
		{ label: L("labelHeading4"), md: "####", style: "h4", key: "h4" },
		{ label: L("labelHeading5"), md: "#####", style: "h5", key: "h5" },
		{ label: L("labelHeading6"), md: "######", style: "h6", key: "h6" }
	];

	var chkH = [], styleH = [];
	for (var i = 0; i < headings.length; i++) {
		var row = headPanel.add("group");
		chkH[i] = row.add("checkbox", undefined, "");
		chkH[i].value = true;
		row.add("statictext", [0, 0, 120, 20], headings[i].label);
		styleH[i] = addStyleDropdown(row, paraStyleNames, headings[i].style, 160, headings[i].key);
	}

	// ── リスト ──
	var listPanel = colL.add("panel", undefined, L("panelList"));
	listPanel.alignChildren = ["fill", "top"];
	listPanel.margins = [15, 20, 15, 10];

	var rowUl = listPanel.add("group");
	var chkUl = rowUl.add("checkbox", undefined, "");
	chkUl.value = true;
	rowUl.add("statictext", [0, 0, 120, 20], L("labelBullet"));
	var styleUl = addStyleDropdown(rowUl, paraStyleNames, "ul-li", 160, "bullet");

	var rowUl2 = listPanel.add("group");
	var chkUl2 = rowUl2.add("checkbox", undefined, "");
	chkUl2.value = true;
	rowUl2.add("statictext", [0, 0, 120, 20], L("labelBulletSub"));
	var styleUl2 = addStyleDropdown(rowUl2, paraStyleNames, "ul-ul-li", 160, "bulletSub");

	var rowOl = listPanel.add("group");
	var chkOl = rowOl.add("checkbox", undefined, "");
	chkOl.value = true;
	rowOl.add("statictext", [0, 0, 120, 20], L("labelNumbered"));
	var styleOl = addStyleDropdown(rowOl, paraStyleNames, "ol-li", 160, "numbered");

	// ── オプション ──
	var optPanel2 = colL.add("panel", undefined, L("panelOptions"));
	optPanel2.alignChildren = ["fill", "top"];
	optPanel2.margins = [15, 20, 15, 10];

	var rowBq = optPanel2.add("group");
	var chkBq = rowBq.add("checkbox", undefined, "");
	chkBq.value = true;
	rowBq.add("statictext", [0, 0, 120, 20], L("labelBlockquote"));
	var styleBq = addStyleDropdown(rowBq, paraStyleNames, "p.blockquote", 160, "blockquote");

	var rowCaption = optPanel2.add("group");
	var chkCaption = rowCaption.add("checkbox", undefined, "");
	chkCaption.value = true;
	rowCaption.add("statictext", [0, 0, 120, 20], L("labelCaption"));
	var styleCaption = addStyleDropdown(rowCaption, paraStyleNames, "p.caption", 160, "caption");

	// ── ソースコード ──
	var codePanel = colL.add("panel", undefined, L("panelCode"));
	codePanel.alignChildren = ["fill", "top"];
	codePanel.margins = [15, 20, 15, 10];

	var rowCodeBlock = codePanel.add("group");
	var chkCodeBlock = rowCodeBlock.add("checkbox", undefined, "");
	chkCodeBlock.value = true;
	rowCodeBlock.add("statictext", [0, 0, 120, 20], L("labelCodeBlock"));
	var styleCodeBlock = addStyleDropdown(rowCodeBlock, paraStyleNames, "p.code", 160, "codeBlock");

	var rowCodeInline = codePanel.add("group");
	var chkCodeInline = rowCodeInline.add("checkbox", undefined, "");
	chkCodeInline.value = true;
	rowCodeInline.add("statictext", [0, 0, 120, 20], L("labelCodeInline"));
	var styleCodeInline = addStyleDropdown(rowCodeInline, charStyleNames, "code", 160, "codeInline");

	// ━━ 右列 ━━

	// ── 画像 ──
	var imgPanel = colR.add("panel", undefined, L("panelImage"));
	imgPanel.alignChildren = ["fill", "top"];
	imgPanel.margins = [15, 20, 15, 10];

	var rowImg = imgPanel.add("group");
	var chkImg = rowImg.add("checkbox", undefined, "");
	chkImg.value = true;
	rowImg.add("statictext", [0, 0, 120, 20], L("labelImage"));
	var styleImg = addStyleDropdown(rowImg, paraStyleNames, "p.img", 160, "image");

	// ── 表組み ──
	var tablePanel = colR.add("panel", undefined, L("panelTable"));
	tablePanel.alignChildren = ["fill", "top"];
	tablePanel.margins = [15, 20, 15, 10];

	var rowTd = tablePanel.add("group");
	var chkTable = rowTd.add("checkbox", undefined, "");
	chkTable.value = true;
	rowTd.add("statictext", [0, 0, 120, 20], L("labelCell"));
	var styleTable = addStyleDropdown(rowTd, paraStyleNames, "td", 160, "cell");

	var rowTh = tablePanel.add("group");
	var chkTh = rowTh.add("checkbox", undefined, "");
	chkTh.value = true;
	rowTh.add("statictext", [0, 0, 120, 20], L("labelHeaderCell"));
	var styleTh = addStyleDropdown(rowTh, paraStyleNames, "th", 160, "headerCell");

	// ── 文字スタイル ──
	var charPanel = colR.add("panel", undefined, L("panelCharStyle"));
	charPanel.alignChildren = ["fill", "top"];
	charPanel.margins = [15, 20, 15, 10];

	var charItems = [
		{ label: L("labelBold"), find: "\\*\\*([^*\\r\\n]+)\\*\\*", style: "strong-bold", key: "bold" },
		{ label: L("labelEmphasis"), find: "\\*([^*\\r\\n]+)\\*", style: "em-i-marker", key: "emphasis" },
		{ label: L("labelLink"), find: "\\[([^\\]\\r\\n]+)\\]\\(([^)\\r\\n]+)\\)", style: "link", changeTo: "$1 <$2>", key: "link" }
	];

	var chkCh = [], styleCh = [];
	for (var i = 0; i < charItems.length; i++) {
		var row = charPanel.add("group");
		chkCh[i] = row.add("checkbox", undefined, "");
		chkCh[i].value = true;
		row.add("statictext", [0, 0, 120, 20], charItems[i].label);
		styleCh[i] = addStyleDropdown(row, charStyleNames, charItems[i].style, 160, charItems[i].key);
	}

	// ── スペース ──
	var spacePanel = colR.add("panel", undefined, L("panelSpace"));
	spacePanel.alignChildren = ["left", "top"];
	spacePanel.margins = [15, 20, 15, 10];
	var chkTrimSpace = spacePanel.add("checkbox", undefined, L("labelTrimLead"));
	chkTrimSpace.value = true;
	var chkTrailSpace = spacePanel.add("checkbox", undefined, L("labelTrimTrail"));
	chkTrailSpace.value = true;
	var chkNbsp = spacePanel.add("checkbox", undefined, L("labelNbsp"));
	chkNbsp.value = true;

	// ── クリーンアップ ──
	var optPanel = colR.add("panel", undefined, L("panelCleanup"));
	optPanel.alignChildren = ["left", "top"];
	optPanel.margins = [15, 20, 15, 10];
	var chkTrimReturn = optPanel.add("checkbox", undefined, L("labelTrimReturn"));
	chkTrimReturn.value = true;
	var chkComment = optPanel.add("checkbox", undefined, L("labelComment"));
	chkComment.value = true;
	var chkHr = optPanel.add("checkbox", undefined, L("labelHr"));
	chkHr.value = true;
	var chkGeta = optPanel.add("checkbox", undefined, L("labelGeta"));
	chkGeta.value = false;

	// ── ソース形式プリセット切替 ──
	var ddMap = {
		paragraph: styleP, h1: styleH[0], h2: styleH[1], h3: styleH[2], h4: styleH[3], h5: styleH[4], h6: styleH[5],
		bullet: styleUl, bulletSub: styleUl2, numbered: styleOl,
		blockquote: styleBq, caption: styleCaption,
		codeBlock: styleCodeBlock, codeInline: styleCodeInline,
		image: styleImg, cell: styleTable, headerCell: styleTh,
		bold: styleCh[0], emphasis: styleCh[1], link: styleCh[2]
	};

	function applyPreset(presetName) {
		var preset = STYLE_PRESETS[presetName];
		if (!preset) return;
		for (var key in preset) {
			if (preset.hasOwnProperty(key) && ddMap[key]) {
				selectDropdownByName(ddMap[key], preset[key]);
			}
		}
	}

	rbHTML.onClick = function () { applyPreset("html"); };
	rbMSWord.onClick = function () { applyPreset("msword"); };

	// ── ボタン ──
	var btnGroup = w.add("group");
	btnGroup.alignment = ["right", "top"];
	btnGroup.add("button", undefined, L("btnCancel"), { name: "cancel" });
	btnGroup.add("button", undefined, L("btnRun"), { name: "ok" });

	if (w.show() != 1) return;

	// ドロップダウンから選択値を取得するヘルパー
	function sel(dd) { return dd.selection ? dd.selection.text : null; }

	// 検索対象の決定
	var target;
	if (rbSel.value) {
		if (app.selection.length == 0) { alert(L("msgNoSelection")); return; }
		target = app.selection[0];
		if (!canUseAsTextTarget(target)) {
			alert(L("msgInvalidSelection"));
			return;
		}
	} else if (rbStory.value) {
		if (app.selection.length == 0) { alert(L("msgNoSelection")); return; }
		target = resolveStoryTargetFromSelection(app.selection[0]);
		if (target == null) {
			alert(L("msgInvalidSelection"));
			return;
		}
	} else {
		target = app.documents.item(0);
	}

	// 変換リストを作成（元リストの順番に準拠）
	var operations = [];
	var pStyle = chkP.value ? sel(styleP) : null;

	// 0. 全体に段落スタイルp、文字スタイルなしを強制適用
	operations.push({ find: ".+", changeTo: "$0", style: "p", charStyle: "[None]" });

	// 0b. グラフィックフレームを作成してクリップボードにコピー（画像挿入用）
	if (chkImg.value) {
		operations.push({ type: "createFrame" });
	}

	// 1. フェンスドコードブロック（```）を先に処理（中身を保護）
	//    開始```～終了```を一括マッチし、中身のみ残してスタイル適用
	if (chkCodeBlock.value) {
		var cbs0 = sel(styleCodeBlock);
		operations.push({ find: "^```[^\\r]*\\r([\\s\\S]+?)\\r```", changeTo: "$1", style: cbs0 });
	}

	// 1b. 空行削除・スペース+改行（前処理）
	if (chkTrimReturn.value) {
		operations.push({ find: "\\r\\r+", changeTo: "\\r", style: null });
		operations.push({ find: " \\r", changeTo: "\\r", style: null });
	}

	// 2. 番号リスト（見出しより前に処理し、### 1. 等の誤マッチを防ぐ）
	if (chkOl.value) {
		operations.push({ find: "^\\d+\\.\\s(.+)", changeTo: "$1", style: sel(styleOl) });
	}

	// 3. 見出し（###### → # の順で処理）
	for (var i = headings.length - 1; i >= 0; i--) {
		if (chkH[i].value) {
			operations.push({ find: "^" + headings[i].md + " (.+)", changeTo: "$1", style: sel(styleH[i]) });
		}
	}

	// 3. 箇条書き（リンク付き）
	if (chkUl.value) {
		var ulStyle = sel(styleUl);
		operations.push({ find: "- \\[([^\\]\\r\\n]+)\\]\\(([^)\\r\\n]+)\\)", changeTo: "$1\r$2", style: ulStyle });
	}

	// 4. 画像
	if (chkImg.value) {
		var imgStyle = sel(styleImg);
		operations.push({ find: "^\\s*<img [^>]*src=\"(.+?)\".+>", changeTo: "~C\\r$1", style: imgStyle });
		operations.push({ find: "^!\\[\\]\\((.+)\\)", changeTo: "~C\\r$1", style: imgStyle });
	}

	// 5. 箇条書き
	if (chkUl.value) {
		var ulStyle2 = sel(styleUl);
		operations.push({ find: "^- (.+)", changeTo: "$1", style: ulStyle2 });
		operations.push({ find: "^\\* (.+)", changeTo: "$1", style: ulStyle2 });
		operations.push({ find: "^\\+ (.+)", changeTo: "$1", style: ulStyle2 });
		operations.push({ find: "^・(.+)", changeTo: "$1", style: ulStyle2 });
	}
	if (chkUl2.value) {
		operations.push({ find: "^\\t- (.+)", changeTo: "$1", style: sel(styleUl2) });
	}

	// 6. 引用
	if (chkBq.value) operations.push({ find: "^>\\s?(.+)", changeTo: "$1", style: sel(styleBq) });

	// 7. 行頭スペース削除
	if (chkTrimSpace.value) {
		operations.push({ find: "^[\\s|　]", changeTo: "", style: null });
		operations.push({ find: "\\s*\\t", changeTo: "\\t", style: null });
	}

	// 8. 文字スタイル（**bold** → *italic* → link の順）
	for (var i = 0; i < charItems.length; i++) {
		if (chkCh[i].value) {
			var ct = charItems[i].changeTo || "$1";
			operations.push({ find: charItems[i].find, changeTo: ct, charStyle: sel(styleCh[i]), style: null });
		}
	}

	// 8b. リンク
	if (chkCh.length > 0) {
		var linkStyle = (charItems.length > 2 && chkCh[2].value) ? sel(styleCh[2]) : null;
		if (linkStyle) {
			operations.push({ find: "〓(.+?)〓", changeTo: "<$1>", charStyle: linkStyle, style: null });
		}
	}

	// 9. ソースコード（コードブロック）
	if (chkCodeBlock.value) {
		var cbs = sel(styleCodeBlock);
		operations.push({ find: "^<pre><code>(.+)<\\/code><\\/pre>", changeTo: "$1", style: cbs });
		operations.push({ find: "^<pre>(.+)<\\/pre>", changeTo: "$1", style: cbs });
		operations.push({ find: "^<pre>\\r(.+)\\r<\\/pre>", changeTo: "$1", style: cbs });
		operations.push({ find: "^<pre>\\r", changeTo: "<pre>", style: cbs });
		operations.push({ find: "\\r<\\/pre>(.+)", changeTo: "</pre>", style: cbs });
		operations.push({ find: "^<code>(.+)<\\/code>", changeTo: "$1", style: cbs });
	}
	// 10. ソースコード（インライン）
	if (chkCodeInline.value) {
		operations.push({ find: "`([^`\\r\\n]+)`", changeTo: "$1", charStyle: sel(styleCodeInline), style: null });
	}

	// 11. キャプション
	if (chkCaption.value) {
		var capStyle = sel(styleCaption);
		operations.push({ find: "^[\\[［]?キャプション[\\]］：:]?(.+)", changeTo: "$1", style: capStyle });
		operations.push({ find: "^<!-- キャプション -->(.+)<!-- \\/キャプション -->", changeTo: "$1", style: capStyle });
		operations.push({ find: "^▲(.+)", changeTo: "$1", style: capStyle });
		operations.push({ find: "^caption[：:]\\s?(.+)", changeTo: "$1", style: capStyle });
	}

	// 12. テーブル
	var thStyleName = chkTh.value ? sel(styleTh) : null;
	var tblStyle = chkTable.value ? sel(styleTable) : null;

	// 12a. ヘッダー行を検出しthを適用（区切り行の直前の行）
	if (chkTh.value) {
		operations.push({ find: "^(.*\\|.*)$\\r?\\n^[-:| ]+$", changeTo: "$1", style: thStyleName });
	}
	// 12b. 区切り行（|------|------| 等）を削除
	if (chkTable.value || chkTh.value) {
		operations.push({ find: "^\\|?(\\s*:?-{3,}:?\\s*\\|)+\\s*:?-{3,}:?\\s*\\|?\\r?", changeTo: "", style: null });
	}
	// 12c. td行：行頭 | 削除 → 行末 | 削除 → 行中 | をタブに
	if (chkTable.value) {
		operations.push({ find: "^\\|\\s*", changeTo: "", style: tblStyle });
		operations.push({ find: "\\s*\\|\\s*$", changeTo: "", style: tblStyle });
		operations.push({ find: "\\s*\\|\\s*", changeTo: "\\t", style: tblStyle });
	}
	// 12d. th行：| を整形（thスタイルを保持、findStyleでth行のみ対象）
	if (chkTh.value) {
		operations.push({ find: "^\\|\\s*", changeTo: "", style: null, findStyle: thStyleName });
		operations.push({ find: "\\s*\\|\\s*$", changeTo: "", style: null, findStyle: thStyleName });
		operations.push({ find: "\\s*\\|\\s*", changeTo: "\\t", style: null, findStyle: thStyleName });
	}


	// 14. クリーンアップ（コメントアウト、水平線）
	if (chkComment.value) operations.push({ find: "<!-- (.+) -->", changeTo: "\\r", style: pStyle });
	if (chkHr.value) operations.push({ find: "^-+", changeTo: "\\r", style: pStyle });

	// 15. 空行削除（後処理）
	if (chkTrimReturn.value) {
		operations.push({ find: "\\r\\r+", changeTo: "\\r", style: null });
	}

	// 16. スペース
	if (chkNbsp.value) operations.push({ find: "&nbsp;", changeTo: " ", style: null, type: "text" });
	if (chkTrailSpace.value) operations.push({ find: "\\s+$", changeTo: " ", style: null });

	// 17. 〓を削除
	if (chkGeta.value) {
		operations.push({ find: "〓", changeTo: "", style: null });
	}

	// 18. 画像（念のため：行頭以外もマッチ）
	if (chkImg.value) {
		var imgStyle2 = sel(styleImg);
		operations.push({ find: "<img [^>]*src=\"(.+?)\".+>", changeTo: "~C\\r$1", style: imgStyle2 });
	}

	if (operations.length == 0) {
		alert(L("msgNoTarget"));
		return;
	}

	// プログレスバー
	var prog = new Window("palette", L("progressTitle"));
	prog.add("statictext", undefined, L("progressLabel"));
	var bar = prog.add("progressbar", [0, 0, 300, 20], 0, operations.length);
	var progLabel = prog.add("statictext", [0, 0, 300, 20], "");
	prog.show();

	// 実行（1つのUndo単位にまとめる）
	var totalCount = 0;
	var details = [];
	app.doScript(function () {
		for (var i = 0; i < operations.length; i++) {
			var op = operations[i];
			bar.value = i + 1;
			progLabel.text = (i + 1) + " / " + operations.length;
			prog.update();

			if (op.type == "createFrame") {
				// グラフィックフレームを作成してクリップボードにコピー
				var currentPage = doc.layoutWindows[0].activePage;
				var pageMargins = currentPage.marginPreferences;
				var pageBounds = currentPage.bounds;
				var marginWidth = (pageBounds[3] - pageBounds[1]) - pageMargins.left - pageMargins.right;
				var graphicFrame = currentPage.rectangles.add({
					geometricBounds: [pageMargins.top, pageMargins.left, pageMargins.top + 50, pageMargins.left + marginWidth],
					contentType: ContentType.GRAPHIC_TYPE
				});
				graphicFrame.fillColor = doc.swatches.itemByName("None");
				graphicFrame.strokeColor = doc.swatches.itemByName("None");
				app.select(graphicFrame);
				app.copy();
				graphicFrame.remove();
				continue;
			}

			var count = (op.type == "text")
				? findChangeText(target, op.find, op.changeTo)
				: findChangeGrep(target, op.find, op.changeTo, op.style, op.charStyle, op.findStyle);
			if (count > 0) {
				var label = op.style != null ? op.style : op.find;
				details.push(label + L("resultSep") + count + L("resultCountSuffix"));
			}
			totalCount += count;
		}
	}, ScriptLanguage.javascript, undefined, UndoModes.ENTIRE_SCRIPT, L("undoName"));

	prog.close();

	// 結果表示（該当なしの場合のみ）
	if (totalCount == 0) {
		alert(L("msgNoMatch"));
	}
}

function findChangeGrep(target, findWhat, changeTo, styleName, charStyleName, findStyleName) {
	app.findGrepPreferences = NothingEnum.nothing;
	app.changeGrepPreferences = NothingEnum.nothing;

	app.findGrepPreferences.findWhat = findWhat;
	app.changeGrepPreferences.changeTo = changeTo;

	// 検索対象の段落スタイルを絞り込む
	if (findStyleName != null && findStyleName != undefined) {
		try {
			app.findGrepPreferences.appliedParagraphStyle = findStyleName;
		} catch (e) { /* スタイルが見つからない場合は絞り込まない */ }
	}

	if (styleName != null) {
		try {
			app.changeGrepPreferences.appliedParagraphStyle = styleName;
		} catch (e) {
			app.findGrepPreferences = NothingEnum.nothing;
			app.changeGrepPreferences = NothingEnum.nothing;
			return 0;
		}
	}

	if (charStyleName != null && charStyleName != undefined) {
		try {
			if (charStyleName == "[None]" || charStyleName == "[なし]") {
				app.changeGrepPreferences.appliedCharacterStyle = app.activeDocument.characterStyles[0];
			} else {
				app.changeGrepPreferences.appliedCharacterStyle = charStyleName;
			}
		} catch (e) {
			app.findGrepPreferences = NothingEnum.nothing;
			app.changeGrepPreferences = NothingEnum.nothing;
			return 0;
		}
	}

	var found = target.changeGrep();

	app.findGrepPreferences = NothingEnum.nothing;
	app.changeGrepPreferences = NothingEnum.nothing;

	return found.length;
}

function findChangeText(target, findWhat, changeTo) {
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;

	app.findTextPreferences.findWhat = findWhat;
	app.changeTextPreferences.changeTo = changeTo;

	var found = target.changeText();

	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;

	return found.length;
}

main();

})();
