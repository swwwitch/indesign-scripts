#target indesign

/*
 * FileNameManager.jsx
 *
 * アクティブなドキュメントのファイル名を、ベース・日付・タイトル・バージョンのセグメント単位で編集し、リネーム／別名保存／コピー保存を行います。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FileNameManager";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-05-27";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/FileNameManager.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/FileNameManager.md

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

        /* ファイルの保存・リネーム・削除はドキュメント編集ではなく InDesign の取り消し対象外のため、
           doScript でのラップは行わない
           / Saving, renaming and deleting files are file-system operations outside InDesign's undo,
             so this script is not wrapped in doScript */

        // =========================================
        // バージョン / Version
        // =========================================

        var SCRIPT_VERSION = "v1.0.0";

        // =========================================
        // ユーザー設定 / User Settings
        // =========================================

        /* 出力時のセグメント順序。base / title / timestamp / version から選んで並べ替え可能
           / Output segment order. Pick from base / title / timestamp / version */
        var SEGMENT_ORDER = ['base', 'title', 'timestamp', 'version'];

        var NEW_NAME_FIELD_WIDTH = 250;
        var PANEL_SPACING = 8;

        var DIALOG_OPACITY = 0.98;

        // =========================================
        // ローカライズ / Localization
        // =========================================

        var lang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

        /* 日英ラベル定義 / Japanese-English label definitions */

        var LABELS = {
            dialog: {
                title: { ja: "ファイル名の管理", en: "File Name Manager" }
            },
            panel: {
                mode: { ja: "動作", en: "Mode" },
                filename: { ja: "ファイル名", en: "File Name" },
                options: { ja: "オプション", en: "Options" }
            },
            radio: {
                rename: { ja: "ファイル名の変更", en: "Rename" },
                saveAs: { ja: "別名で保存", en: "Save As" },
                saveCopy: { ja: "コピーを保存", en: "Save a Copy" },
                noChange: { ja: "変更しない", en: "No Change" },
                titleNone: { ja: "つけない", en: "None" },
                titleParent: { ja: "親フォルダー", en: "Parent Folder" },
                titleCustom: { ja: "指定", en: "Custom" },
                timestampNone: { ja: "つけない", en: "None" },
                timestampDate: { ja: "YYYYMMDD", en: "YYYYMMDD" },
                versionNone: { ja: "つけない", en: "None" },
                versionShort: { ja: "-vN", en: "-vN" },
                versionPadded: { ja: "-v0N", en: "-v0N" }
            },
            label: {
                currentName: { ja: "現在", en: "Current" },
                finalName: { ja: "変更後", en: "Final" },
                base: { ja: "ベース", en: "Base" },
                title: { ja: "タイトル", en: "Title" },
                timestamp: { ja: "タイムスタンプ", en: "Timestamp" },
                version: { ja: "バージョン番号", en: "Version" },
                separator: { ja: "区切り記号", en: "Separator" }
            },
            tip: {
                rename: {
                    ja: "新しい名前で保存したあと、元のファイルを削除します。",
                    en: "Saves with the new name, then removes the original file."
                },
                saveAs: {
                    ja: "新しい名前で保存します（アクティブドキュメントは新ファイルに切り替わる）。",
                    en: "Saves with the new name (the active document switches to the new file)."
                },
                saveCopy: {
                    ja: "元のファイルに上書き保存したうえで、別名のコピーを作成します。アクティブドキュメントは元のまま。",
                    en: "Saves the original, then creates a copy with the new name. The active document is unchanged."
                },
                title: {
                    ja: "ファイル名のタイトル部分の扱いを選択します。「指定」で入力欄の文字列を使用します。",
                    en: "Choose how to set the title part of the filename. With \"Custom\", the entered text is used."
                },
                timestamp: {
                    ja: "タイムスタンプの形式を選択。「つけない」で元の日付があっても削除します。",
                    en: "Choose timestamp format. \"None\" removes any existing date."
                },
                version: {
                    ja: "バージョン番号の形式。-vN はパディング無し、-v0N はゼロ埋め。「つけない」で削除。",
                    en: "Version format. -vN has no padding, -v0N is zero-padded. \"None\" removes."
                },
                separator: {
                    ja: "ファイル名全体の区切り記号の扱いを選択します。",
                    en: "Choose how separators in the filename are handled."
                }
            },
            button: {
                cancel: { ja: "キャンセル", en: "Cancel" }
            },
            message: {
                noDoc: { ja: "ドキュメントが開かれていません", en: "No document is open." },
                emptyName: { ja: "ファイル名が空です", en: "File name is empty." },
                chooseDestination: { ja: "保存先フォルダを指定", en: "Choose destination folder" },
                confirmOverwrite: {
                    ja: "同名ファイルが存在します。上書きしますか？",
                    en: "A file with the same name exists. Overwrite?"
                },
                saveFailed: { ja: "保存に失敗しました", en: "Failed to save" }
            }
        };

        /**
         * ドット区切りキーでラベルを取得する
         * @param {string} path 例: "dialog.title"
         * @returns {string} 現在の言語のラベル文字列。見つからない場合はキーをそのまま返す
         */
        function getLabel(path) {
            var parts = String(path).split('.');
            var entry = LABELS;
            for (var i = 0; i < parts.length; i++) {
                entry = entry && entry[parts[i]];
                if (!entry) return path;
            }
            return entry[lang] || entry.en || path;
        }

        /**
         * コロン付きラベルを取得する（日本語は全角コロン、英語は半角コロン）
         * @param {string} path 例: "label.separator"
         * @returns {string} コロンを付与したラベル文字列
         */
        function getLabelWithColon(path) {
            return L(path) + (lang === 'ja' ? '：' : ':');
        }

        // =========================================
        // ヘルパー / Helpers
        // =========================================

        var VERSION_TOKEN_RE = /^[vV]\d+$/;   // v123 / V123

        /**
         * 月日として妥当な組み合わせかを判定する
         * @param {number} month 月
         * @param {number} day 日
         * @returns {boolean} 妥当なら true
         */
        function isValidMonthDay(monthStr, dayStr) {
            var m = parseInt(monthStr, 10);
            var d = parseInt(dayStr, 10);
            return m >= 1 && m <= 12 && d >= 1 && d <= 31;
        }

        /**
         * 日付セグメントとして解釈できる文字列かを判定する
         * @param {string} token 判定する文字列
         * @returns {boolean} 日付なら true
         */
        function isDateToken(token) {
            var s = String(token);
            if (/^\d{8}$/.test(s)) return isValidMonthDay(s.substring(4, 6), s.substring(6, 8));
            if (/^\d{6}$/.test(s)) return isValidMonthDay(s.substring(2, 4), s.substring(4, 6));
            return false;
        }

        /**
         * バージョンセグメントとして解釈できる文字列かを判定する
         * @param {string} token 判定する文字列
         * @returns {boolean} バージョンなら true
         */
        function isVersionToken(token) {
            return VERSION_TOKEN_RE.test(String(token));
        }

        /**
         * ファイル名をセグメントの配列へ分解する
         * @param {string} fileName 分解するファイル名
         * @returns {Array<object>} セグメントの配列
         */
        function parseFileName(name) {
            var split = String(name).split(/([-_])/); // ["handout","-","Adobe","-","20260422"]
            var segments = [];
            var textBuffer = [];      // token と区切りを交互に蓄積（join('') で結合）
            var textLeadingSep = '';  // テキスト segment の直前に置く区切り
            var hasDate = false, hasVersion = false;
            var currentSep = '';

            /**
             * 読み取り中のテキストを 1 セグメントとして確定する
             * @returns {void}
             */
            function flushText() {
                if (!textBuffer.length) return;
                segments.push({
                    kind: 'text',
                    value: textBuffer.join(''),
                    sep: textLeadingSep
                });
                textBuffer = [];
                textLeadingSep = '';
            }

            var i = 0;
            while (i < split.length) {
                if (i % 2 === 1) {
                    currentSep = split[i];
                    i++;
                    continue;
                }
                var token = split[i];

                // YYYY-MM-DD / YYYY_MM_DD（3 トークン + 同じ区切り 2 つ）の日付パターン
                if (!hasDate && i > 0
                    && /^\d{4}$/.test(token)
                    && i + 4 < split.length
                    && (split[i + 1] === '-' || split[i + 1] === '_')
                    && /^\d{2}$/.test(split[i + 2])
                    && split[i + 3] === split[i + 1]
                    && /^\d{2}$/.test(split[i + 4])
                    && isValidMonthDay(split[i + 2], split[i + 4])) {
                    flushText();
                    var compositeDate = token + split[i + 1] + split[i + 2] + split[i + 3] + split[i + 4];
                    segments.push({ kind: 'date', value: compositeDate, sep: currentSep });
                    hasDate = true;
                    i += 5; // 3 tokens + 2 separators
                    continue;
                }

                if (i === 0) {
                    segments.push({ kind: 'base', value: token, sep: '' });
                } else if (!hasDate && isDateToken(token)) {
                    flushText();
                    segments.push({ kind: 'date', value: token, sep: currentSep });
                    hasDate = true;
                } else if (!hasVersion && isVersionToken(token)) {
                    flushText();
                    segments.push({ kind: 'version', value: token, sep: currentSep });
                    hasVersion = true;
                } else {
                    if (!textBuffer.length) {
                        textLeadingSep = currentSep;
                    } else {
                        textBuffer.push(currentSep);
                    }
                    textBuffer.push(token);
                }
                i++;
            }
            flushText();
            return segments;
        }

        /**
         * 指定した種別の最初のセグメント値を取得する
         * @param {Array<object>} segments セグメントの配列
         * @param {string} kind セグメントの種別
         * @returns {string} セグメントの値。なければ空文字
         */
        function getFirstSegmentValue(segments, kind) {
            for (var i = 0; i < segments.length; i++) {
                if (segments[i].kind === kind) return segments[i].value;
            }
            return '';
        }

        /**
         * 左側をゼロ埋めして桁を揃える
         * @param {*} value 対象の値
         * @param {number} length 揃える桁数
         * @returns {string} ゼロ埋めした文字列
         */
        function padLeft(str, width) {
            while (str.length < width) str = '0' + str;
            return str;
        }

        /**
         * 今日の日付からタイムスタンプ文字列を作る
         * @param {string} format 日付の書式
         * @returns {string} タイムスタンプ
         */
        function todayTimestamp() {
            var d = new Date();
            return String(d.getFullYear()) +
                padLeft(String(d.getMonth() + 1), 2) +
                padLeft(String(d.getDate()), 2);
        }

        /**
         * バージョン番号を指定の書式へ整形する
         * @param {number} versionNumber バージョン番号
         * @param {string} format バージョンの書式
         * @returns {string} 整形した文字列
         */
        function formatVersion(originalVersion, mode) {
            var match = String(originalVersion || '').match(/^([vV])(\d+)$/);
            var letter = match ? match[1] : 'v';
            var nextNum = match ? (parseInt(match[2], 10) + 1) : 2;
            if (mode === 'padded') {
                var width = match ? Math.max(match[2].length, 2) : 2;
                return letter + padLeft(String(nextNum), width);
            }
            return letter + String(nextNum);
        }

        /**
         * パーセントエンコードされた文字列を復号する
         * @param {string} text 対象の文字列
         * @returns {string} 復号した文字列
         */
        function decodePercentEncoded(str) {
            str = String(str);
            if (str.indexOf('%') === -1) return str;
            try {
                return decodeURIComponent(str);
            } catch (e) {
                return str;
            }
        }

        /**
         * ファイル名から拡張子を取り除く
         * @param {string} fileName 対象のファイル名
         * @returns {string} 拡張子を除いた名前
         */
        function stripExtension(name) {
            var dot = name.lastIndexOf('.');
            return (dot > 0) ? name.substring(0, dot) : name;
        }

        /**
         * ファイル名に使えない文字を置き換える
         * @param {string} fileName 元のファイル名
         * @returns {string} 安全なファイル名
         */
        function sanitizeFilename(str) {
            return String(str).replace(/[\\\/:*?"<>|]/g, '').replace(/^\s+|\s+$/g, '');
        }

        /**
         * 保存先フォルダをユーザーに選ばせる
         * @returns {Folder|null} 選択したフォルダ。キャンセル時は null
         */
        function pickInddDestination(promptLabel) {
            var chosen = File.saveDialog(promptLabel, '*.indd');
            if (!chosen) return null;
            var file = File(chosen.fsName);
            if (!/\.indd$/i.test(file.name)) {
                file = File(file.parent.fsName + '/' + stripExtension(file.name) + '.indd');
            }
            return file;
        }

        /**
         * 設定を保存するファイルを取得する
         * @returns {File} 設定ファイル
         */
        function getPrefsFile() {
            return File(Folder.userData.fsName + '/FileNameManager-prefs.txt');
        }

        /**
         * 保存しておいた設定を読み込む
         * @returns {object} 読み込んだ設定
         */
        function loadPrefs() {
            var file = getPrefsFile();
            if (!file.exists) return {};
            file.encoding = 'UTF-8';
            file.open('r');
            var raw = file.read();
            file.close();
            var prefs = {};
            var lines = String(raw).split('\n');
            for (var i = 0; i < lines.length; i++) {
                var line = lines[i];
                var eq = line.indexOf('=');
                if (eq > 0) prefs[line.substring(0, eq)] = line.substring(eq + 1);
            }
            return prefs;
        }

        /**
         * 現在の設定を保存する
         * @param {object} prefs 保存する設定
         * @returns {void}
         */
        function savePrefs(prefs) {
            try {
                var file = getPrefsFile();
                file.encoding = 'UTF-8';
                file.open('w');
                var lines = [];
                for (var key in prefs) {
                    if (prefs.hasOwnProperty(key)) lines.push(key + '=' + prefs[key]);
                }
                file.write(lines.join('\n'));
                file.close();
            } catch (e) { /* 保存できなくても処理は止めない */ }
        }

        /**
         * アクティブドキュメントのファイル情報を集める
         * @returns {object|null} ファイル情報。取得できない場合は null
         */
        function gatherDocumentInfo(doc) {
            var fullName = doc.fullName;
            var currentName = decodePercentEncoded(fullName ? fullName.name : doc.name);
            var parentFolderName = '';
            if (fullName && fullName.parent) {
                parentFolderName = decodePercentEncoded(fullName.parent.name);
            }
            return {
                currentName: currentName,
                baseName: stripExtension(currentName),
                fsPath: fullName ? fullName.fsName : null,
                folder: fullName ? fullName.parent : null,
                parentFolderName: parentFolderName
            };
        }

        /**
         * 保存先フォルダを確定する（未保存なら選択させる）
         * @param {object} documentInfo ファイル情報
         * @returns {Folder|null} 保存先フォルダ。決まらない場合は null
         */
        function ensureTargetFolder(folder) {
            if (folder) return folder;
            var picked = pickInddDestination(getLabel('message.chooseDestination'));
            return picked ? picked.parent : null;
        }

        /**
         * 同名ファイルがある場合に上書きの可否を確認する
         * @param {File} targetFile 保存先のファイル
         * @returns {boolean} 続行してよければ true
         */
        function confirmOverwriteIfExists(destFile, originalFsPath) {
            if (!destFile.exists) return true;
            if (destFile.fsName === originalFsPath) return true;
            return confirm(getLabel('message.confirmOverwrite') + '\n\n' + destFile.fsName);
        }

        /**
         * リネーム時に元ファイルを削除する
         * @param {File} originalFile 元のファイル
         * @returns {void}
         */
        function removeOriginalFile(originalFsPath, destFsPath) {
            if (!originalFsPath || originalFsPath === destFsPath) return;
            var file = File(originalFsPath);
            if (!file.exists) return;
            try { file.remove(); } catch (e) { /* 削除できない場合は黙って継続 */ }
        }

        /**
         * 選択したモードでリネーム・別名保存・コピー保存を実行する
         * @param {string} mode 動作モード
         * @param {File} targetFile 保存先のファイル
         * @param {object} documentInfo ファイル情報
         * @returns {boolean} 成功したら true
         */
        function executeOutput(doc, destFile, mode, originalFsPath) {
            if (mode === 'copy' && originalFsPath) {
                // 現在の変更を元ファイルへ保存してから、物理ファイルとしてコピー
                if (!doc.saved) doc.save();
                File(originalFsPath).copy(destFile);
                return;
            }
            // rename / saveAs / 未保存ドキュメントの copy: 新名で保存
            doc.save(destFile);
            if (mode === 'rename') {
                removeOriginalFile(originalFsPath, destFile.fsName);
            }
        }

        /**
         * UI の状態とセグメントから最終的なファイル名を組み立てる
         * @param {object} uiState UI の状態
         * @param {Array<object>} segments セグメントの配列
         * @returns {string} 最終的なファイル名
         */
        function buildFinalName(segments, uiState) {
            // 区切り記号: 明示選択があればそれを、無ければ元のファイル名で優勢な区切りを使う
            var defaultSep;
            if (uiState.separator === '-' || uiState.separator === '_') {
                defaultSep = uiState.separator;
            } else {
                defaultSep = dominantSeparator(segments);
            }

            /**
             * セグメント種別ごとに出力する値を決める
             * @param {string} kind セグメントの種別
             * @param {object} uiState UI の状態
             * @param {Array<object>} segments セグメントの配列
             * @returns {string} 出力する値
             */
            function valueForKind(kind) {
                if (kind === 'base') {
                    return getFirstSegmentValue(segments, 'base');
                }
                if (kind === 'title') {
                    if (uiState.titleMode === 'none') return '';
                    if (uiState.titleMode === 'parent') return sanitizeFilename(uiState.parentFolderName);
                    return sanitizeFilename(uiState.titleText);
                }
                if (kind === 'timestamp') {
                    return (uiState.timestamp === 'date') ? todayTimestamp() : '';
                }
                if (kind === 'version') {
                    if (uiState.version === 'none') return '';
                    return formatVersion(getFirstSegmentValue(segments, 'version'), uiState.version);
                }
                return '';
            }

            var pieces = [];
            for (var i = 0; i < SEGMENT_ORDER.length; i++) {
                var value = valueForKind(SEGMENT_ORDER[i]);
                if (value) pieces.push(value);
            }
            var result = pieces.join(defaultSep);

            // 「変更しない」以外は内部の "-" / "_" も統一（タイトル等に既存の区切りが残る場合に備えて）
            if (uiState.separator === '-') {
                result = result.replace(/_/g, '-');
            } else if (uiState.separator === '_') {
                result = result.replace(/-/g, '_');
            }

            return result;
        }

        /**
         * ファイル名で最も多く使われている区切り記号を求める
         * @param {string} fileName 対象のファイル名
         * @returns {string} 区切り記号
         */
        function dominantSeparator(segments) {
            var dashes = 0, underscores = 0;
            for (var i = 0; i < segments.length; i++) {
                if (segments[i].sep === '-') dashes++;
                else if (segments[i].sep === '_') underscores++;
            }
            return (underscores > dashes) ? '_' : '-';
        }

        // =========================================
        // ダイアログビルダー / Dialog builders
        // =========================================


        /**
         * 動作モードのパネルを組み立てる
         * @param {Window} dialog 対象のダイアログ
         * @param {object} prefs 保存しておいた設定
         * @returns {object} パネル内のコントロール
         */
        function buildModePanel(parent, prefs) {
            var panel = parent.add('panel', undefined, getLabel('panel.mode'));
            setupPanel(panel);
            var renameRadio = panel.add('radiobutton', undefined, getLabel('radio.rename'));
            renameRadio.helpTip = getLabel('tip.rename');
            var saveAsRadio = panel.add('radiobutton', undefined, getLabel('radio.saveAs'));
            saveAsRadio.helpTip = getLabel('tip.saveAs');
            var saveCopyRadio = panel.add('radiobutton', undefined, getLabel('radio.saveCopy'));
            saveCopyRadio.helpTip = getLabel('tip.saveCopy');
            // 初期選択（プリセットがあれば優先、無ければ「別名で保存」）
            var initialMode = (prefs && (prefs.mode === 'rename' || prefs.mode === 'copy')) ? prefs.mode : 'saveAs';
            renameRadio.value = (initialMode === 'rename');
            saveAsRadio.value = (initialMode === 'saveAs');
            saveCopyRadio.value = (initialMode === 'copy');
            return {
                panel: panel,
                renameRadio: renameRadio,
                saveAsRadio: saveAsRadio,
                saveCopyRadio: saveCopyRadio,
                /* 現在選択中のモード（'rename' / 'saveAs' / 'copy'） / Currently selected mode */
                getMode: function () {
                    if (renameRadio.value) return 'rename';
                    if (saveCopyRadio.value) return 'copy';
                    return 'saveAs';
                }
            };
        }

        /**
         * ファイル名のパネルを組み立てる
         * @param {Window} dialog 対象のダイアログ
         * @param {string} currentName 現在のファイル名
         * @returns {object} パネル内のコントロール
         */
        function buildFilenamePanel(parent, currentName) {
            var panel = parent.add('panel', undefined, getLabel('panel.filename'));
            setupPanel(panel);

            var currentNameRow = panel.add('group');
            currentNameRow.orientation = 'row';
            var currentNameLabel = currentNameRow.add('statictext', undefined, getLabelWithColon('label.currentName'));
            currentNameRow.add('statictext', undefined, currentName);

            var finalNameRow = panel.add('group');
            finalNameRow.orientation = 'row';
            var finalNameLabel = finalNameRow.add('statictext', undefined, getLabelWithColon('label.finalName'));
            // 「変更後：」は statictext のためレイアウト後にサイズ固定。
            // 現在のファイル名と「入力フィールド + 余白」の大きい方を確保しておく
            var finalNameValue = finalNameRow.add('statictext', undefined, currentName + '.indd');
            var currentNameWidth = panel.graphics.measureString(currentName + '.indd').width;
            finalNameValue.preferredSize.width = Math.max(currentNameWidth + 20, NEW_NAME_FIELD_WIDTH + 150);

            alignLabelWidths(panel, [
                getLabelWithColon('label.currentName'),
                getLabelWithColon('label.finalName')
            ], [currentNameLabel, finalNameLabel]);

            return {
                panel: panel,
                finalNameValue: finalNameValue
            };
        }

        /**
         * セグメントごとのオプションパネルを組み立てる
         * @param {object} parent 追加先のコンテナ
         * @param {Array<object>} segments セグメントの配列
         * @param {object} prefs 保存しておいた設定
         * @returns {object} パネル内のコントロール
         */
        function buildOptionsPanel(parent, segments, prefs, parentFolderName) {
            var panel = parent.add('panel', undefined, getLabel('panel.options'));
            setupPanel(panel);

            var baseValue = getFirstSegmentValue(segments, 'base');
            var initialText = getFirstSegmentValue(segments, 'text');

            // ベース（読み取り専用表示）
            var baseRow = panel.add('group');
            baseRow.orientation = 'row';
            var baseLabel = baseRow.add('statictext', undefined, getLabelWithColon('label.base'));
            baseRow.add('statictext', undefined, baseValue);

            // タイトル: 1 行目 = ラベル + 3 ラジオ、2 行目 = 「指定」用の入力欄
            var titleSection = panel.add('group');
            titleSection.orientation = 'column';
            titleSection.alignChildren = ['fill', 'top'];
            titleSection.spacing = 4;

            var titleRow = titleSection.add('group');
            titleRow.orientation = 'row';
            titleRow.alignment = ['left', 'top'];
            var titleLabel = titleRow.add('statictext', undefined, getLabelWithColon('label.title'));
            titleLabel.helpTip = getLabel('tip.title');
            var titleNoneRadio = titleRow.add('radiobutton', undefined, getLabel('radio.titleNone'));
            titleNoneRadio.helpTip = getLabel('tip.title');
            var titleParentRadio = titleRow.add('radiobutton', undefined, getLabel('radio.titleParent'));
            titleParentRadio.helpTip = parentFolderName
                ? getLabel('tip.title') + ' (' + parentFolderName + ')'
                : getLabel('tip.title');
            if (!parentFolderName) titleParentRadio.enabled = false;
            var titleCustomRadio = titleRow.add('radiobutton', undefined, getLabel('radio.titleCustom'));
            titleCustomRadio.helpTip = getLabel('tip.title');

            // 「指定」用の入力欄は次の行（ラベル列幅だけ左に余白を入れて radios に揃える）
            var titleFieldRow = titleSection.add('group');
            titleFieldRow.orientation = 'row';
            titleFieldRow.alignment = ['left', 'top'];
            var titleFieldSpacer = titleFieldRow.add('statictext', undefined, '');
            var titleField = titleFieldRow.add('edittext', undefined, initialText);
            titleField.preferredSize.width = NEW_NAME_FIELD_WIDTH;
            titleField.helpTip = getLabel('tip.title');

            // 初期モード: プリセット優先、無ければ初期テキスト有無で 'custom' / 'none'
            var initialTitleMode;
            if (prefs && (prefs.titleMode === 'none' || prefs.titleMode === 'parent' || prefs.titleMode === 'custom')) {
                initialTitleMode = prefs.titleMode;
            } else {
                initialTitleMode = initialText ? 'custom' : 'none';
            }
            if (initialTitleMode === 'parent' && !parentFolderName) {
                initialTitleMode = initialText ? 'custom' : 'none';
            }
            titleNoneRadio.value = (initialTitleMode === 'none');
            titleParentRadio.value = (initialTitleMode === 'parent');
            titleCustomRadio.value = (initialTitleMode === 'custom');
            titleField.enabled = (initialTitleMode === 'custom');

            /**
             * タイトルの指定方法に応じて入力欄の有効／無効を切り替える
             * @returns {void}
             */
            function syncTitleFieldEnabled() {
                titleField.enabled = titleCustomRadio.value;
            }

            // タイムスタンプ（つけない / YYYYMMDD。デフォルト YYYYMMDD）
            var timestampRow = panel.add('group');
            timestampRow.orientation = 'row';
            var timestampLabel = timestampRow.add('statictext', undefined, getLabelWithColon('label.timestamp'));
            timestampLabel.helpTip = getLabel('tip.timestamp');
            var timestampNoneRadio = timestampRow.add('radiobutton', undefined, getLabel('radio.timestampNone'));
            timestampNoneRadio.helpTip = getLabel('tip.timestamp');
            var timestampDateRadio = timestampRow.add('radiobutton', undefined, getLabel('radio.timestampDate'));
            timestampDateRadio.helpTip = getLabel('tip.timestamp');
            var initialTimestamp = (prefs && prefs.timestamp === 'none') ? 'none' : 'date';
            timestampNoneRadio.value = (initialTimestamp === 'none');
            timestampDateRadio.value = (initialTimestamp === 'date');

            // バージョン番号（つけない / -vN / -v0N。デフォルト -vN）
            var versionRow = panel.add('group');
            versionRow.orientation = 'row';
            var versionLabel = versionRow.add('statictext', undefined, getLabelWithColon('label.version'));
            versionLabel.helpTip = getLabel('tip.version');
            var versionNoneRadio = versionRow.add('radiobutton', undefined, getLabel('radio.versionNone'));
            versionNoneRadio.helpTip = getLabel('tip.version');
            var versionShortRadio = versionRow.add('radiobutton', undefined, getLabel('radio.versionShort'));
            versionShortRadio.helpTip = getLabel('tip.version');
            var versionPaddedRadio = versionRow.add('radiobutton', undefined, getLabel('radio.versionPadded'));
            versionPaddedRadio.helpTip = getLabel('tip.version');
            var initialVersion = (prefs && (prefs.version === 'none' || prefs.version === 'padded')) ? prefs.version : 'short';
            versionNoneRadio.value = (initialVersion === 'none');
            versionShortRadio.value = (initialVersion === 'short');
            versionPaddedRadio.value = (initialVersion === 'padded');

            // 区切り記号（変更しない / - / _ 横並び。デフォルト "-"）
            var separatorRow = panel.add('group');
            separatorRow.orientation = 'row';
            var separatorLabel = separatorRow.add('statictext', undefined, getLabel('label.separator'));
            separatorLabel.helpTip = getLabel('tip.separator');
            var noChangeRadio = separatorRow.add('radiobutton', undefined, getLabel('radio.noChange'));
            noChangeRadio.helpTip = getLabel('tip.separator');
            var dashRadio = separatorRow.add('radiobutton', undefined, '-');
            dashRadio.helpTip = getLabel('tip.separator');
            var underscoreRadio = separatorRow.add('radiobutton', undefined, '_');
            underscoreRadio.helpTip = getLabel('tip.separator');
            var initialSeparator = (prefs && (prefs.separator === '' || prefs.separator === '_')) ? prefs.separator : '-';
            noChangeRadio.value = (initialSeparator === '');
            dashRadio.value = (initialSeparator === '-');
            underscoreRadio.value = (initialSeparator === '_');

            // 5 行のラベル幅を統一（区切り記号のみコロン無し）
            alignLabelWidths(panel, [
                getLabelWithColon('label.base'),
                getLabelWithColon('label.title'),
                getLabelWithColon('label.timestamp'),
                getLabelWithColon('label.version'),
                getLabel('label.separator')
            ], [baseLabel, titleLabel, timestampLabel, versionLabel, separatorLabel]);

            // タイトル 2 行目「指定」入力欄の左余白をラベル列幅と一致させる
            titleFieldSpacer.preferredSize = [titleLabel.preferredSize.width, 1];

            return {
                panel: panel,
                titleNoneRadio: titleNoneRadio,
                titleParentRadio: titleParentRadio,
                titleCustomRadio: titleCustomRadio,
                titleField: titleField,
                syncTitleFieldEnabled: syncTitleFieldEnabled,
                timestampNoneRadio: timestampNoneRadio,
                timestampDateRadio: timestampDateRadio,
                versionNoneRadio: versionNoneRadio,
                versionShortRadio: versionShortRadio,
                versionPaddedRadio: versionPaddedRadio,
                noChangeRadio: noChangeRadio,
                dashRadio: dashRadio,
                underscoreRadio: underscoreRadio,
                /* '' = 変更しない、'-' / '_' = 統一 */
                getSeparator: function () {
                    if (noChangeRadio.value) return '';
                    if (dashRadio.value) return '-';
                    return '_';
                },
                /* 'none' / 'date' */
                getTimestamp: function () {
                    return timestampNoneRadio.value ? 'none' : 'date';
                },
                /* 'none' / 'short' / 'padded' */
                getVersion: function () {
                    if (versionNoneRadio.value) return 'none';
                    if (versionShortRadio.value) return 'short';
                    return 'padded';
                },
                /* 'none' / 'parent' / 'custom' */
                getTitleMode: function () {
                    if (titleNoneRadio.value) return 'none';
                    if (titleParentRadio.value) return 'parent';
                    return 'custom';
                }
            };
        }

        /**
         * ラベル群の幅を最大値に揃える
         * @param {Array<StaticText>} labelControls 幅を揃えるラベル
         * @returns {void}
         */
        function alignLabelWidths(panel, labelTexts, controls) {
            var graphics = panel.graphics;
            var maxWidth = 0;
            for (var i = 0; i < labelTexts.length; i++) {
                var width = graphics.measureString(labelTexts[i]).width;
                if (width > maxWidth) maxWidth = width;
            }
            maxWidth += 12;
            var needsIndicator = false;
            for (var k = 0; k < controls.length; k++) {
                if (controls[k].type === 'checkbox' || controls[k].type === 'radiobutton') {
                    needsIndicator = true;
                    break;
                }
            }
            if (needsIndicator) maxWidth += 20;
            for (var j = 0; j < controls.length; j++) {
                controls[j].preferredSize = [maxWidth, controls[j].preferredSize.height || 20];
            }
        }

        /**
         * ファイル名変更のダイアログを組み立てる
         * @param {object} documentInfo ファイル情報
         * @param {Array<object>} segments セグメントの配列
         * @param {object} prefs 保存しておいた設定
         * @returns {object} ダイアログとコントロール
         */
        function createDialog(segments, currentName, prefs, parentFolderName) {
            var dialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
            dialog.opacity = DIALOG_OPACITY;
            setupWindow(dialog, 10);

            var mode = buildModePanel(dialog, prefs);
            var filename = buildFilenamePanel(dialog, currentName);
            // ファイル名行とオプションパネルの間に余白を入れる
            var optionsSpacer = filename.panel.add('group');
            optionsSpacer.preferredSize = [-1, 5];
            var options = buildOptionsPanel(filename.panel, segments, prefs, parentFolderName);

            // ---- ライブプレビュー ----
            /**
             * ダイアログの現在の状態を取得する
             * @returns {object} UI の状態
             */
            function currentUIState() {
                return {
                    titleMode: options.getTitleMode(),
                    titleText: options.titleField.text,
                    parentFolderName: parentFolderName,
                    timestamp: options.getTimestamp(),
                    version: options.getVersion(),
                    separator: options.getSeparator()
                };
            }

            /**
             * 現在の入力内容でファイル名のプレビューを更新する
             * @returns {void}
             */
            function refreshPreviews() {
                options.syncTitleFieldEnabled();
                var finalBase = buildFinalName(segments, currentUIState());
                filename.finalNameValue.text = finalBase + '.indd';
            }

            options.titleNoneRadio.onClick = refreshPreviews;
            options.titleParentRadio.onClick = refreshPreviews;
            options.titleCustomRadio.onClick = refreshPreviews;
            options.titleField.onChanging = refreshPreviews;
            options.timestampNoneRadio.onClick = refreshPreviews;
            options.timestampDateRadio.onClick = refreshPreviews;
            options.versionNoneRadio.onClick = refreshPreviews;
            options.versionShortRadio.onClick = refreshPreviews;
            options.versionPaddedRadio.onClick = refreshPreviews;
            options.noChangeRadio.onClick = refreshPreviews;
            options.dashRadio.onClick = refreshPreviews;
            options.underscoreRadio.onClick = refreshPreviews;
            refreshPreviews();

            // ---- ボタン ----
            var buttonGroup = dialog.add('group');
            buttonGroup.alignment = 'center';
            buttonGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
            buttonGroup.add('button', undefined, 'OK', { name: 'ok' });

            return {
                dialog: dialog,
                getUIState: currentUIState,
                getMode: mode.getMode
            };
        }

        // =========================================
        // メイン / Main
        // =========================================

        /**
         * ドキュメントを検証し、ダイアログの指定に従って出力する
         * @returns {void}
         */
        function main() {
            if (app.documents.length === 0) {
                alert(getLabel('message.noDoc'));
                return;
            }

            var doc = app.activeDocument;
            var info = gatherDocumentInfo(doc);
            var segments = parseFileName(info.baseName);
            var prefs = loadPrefs();

            var ui = createDialog(segments, info.currentName, prefs, info.parentFolderName);
            if (ui.dialog.show() !== 1) return; // キャンセル

            var uiState = ui.getUIState();
            var newBaseName = buildFinalName(segments, uiState);
            if (!newBaseName) {
                alert(getLabel('message.emptyName'));
                return;
            }

            var targetFolder = ensureTargetFolder(info.folder);
            if (!targetFolder) return; // キャンセル

            var destFile = File(targetFolder.fsName + '/' + newBaseName + '.indd');
            if (!confirmOverwriteIfExists(destFile, info.fsPath)) return;

            try {
                executeOutput(doc, destFile, ui.getMode(), info.fsPath);
                // 成功したら今回の選択をプリセットとして保存
                savePrefs({
                    mode: ui.getMode(),
                    titleMode: uiState.titleMode,
                    timestamp: uiState.timestamp,
                    version: uiState.version,
                    separator: uiState.separator
                });
            } catch (e) {
                alert(getLabel('message.saveFailed') + '\n' + e);
            }
        }

        main();

    })();
