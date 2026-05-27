#target indesign

    /*
    
    ### スクリプト名：
    
    FileNameManager.jsx — 現在のドキュメントのファイル名を変更／別名で保存
    
    ### 概要：
    
    - アクティブな InDesign ドキュメントのファイル名を変更（実ファイルのリネーム）／別名で保存／コピーを保存します。
    - 保存先は常に現在のファイルと同じフォルダ。未保存ドキュメントは保存先フォルダを選択して保存します。
    - ファイル名を base / date / text / version の各セグメントに分解し、元の出現順序を保持したまま編集できます。
      - base: 「ベース」として読み取り専用表示（先頭の固定部分）
      - date: 「タイムスタンプ」ラジオで「つけない / YYYYMMDD」を選択（デフォルト YYYYMMDD）
      - text: 「タイトル」フィールドで編集
      - version: 「バージョン番号」ラジオで「つけない / -vN / -v0N」を選択（デフォルト -vN）
    - 動作モード: 「ファイル名の変更」「別名で保存」（デフォルト）「コピーを保存」（元ファイル維持＋別名コピー）
    - 区切り記号: 「変更しない / `-` / `_`」から選択。`-` または `_` を選ぶとファイル名全体で統一。
    - UI は「動作」「ファイル名」パネル（内部にオプションサブパネル）で構成。各要素に tooltip を付与。
    
    ### 主な機能：
    
    - ファイル名の変更（新名で保存後に元ファイルを削除して実質的にリネーム）
    - 別名で保存（元ファイルは保持。アクティブドキュメントは新ファイルに切り替わる。デフォルト）
    - コピーを保存（元ファイルに上書き保存後、別名のコピーを物理ファイルとして作成。アクティブドキュメントは元のまま）
    - セグメント分解による独立編集（base / date / text / version。順序維持）
    - タイムスタンプ／バージョン番号は「つけない」で削除、「YYYYMMDD」「-vN」等で形式を切替
    - バージョンは選択時に元バージョン +1（無ければ新規付与）
    - モードに応じて OK ボタン文言を「リネーム / 保存 / コピーを保存」に動的切替
    
    ### 処理の流れ：
    
    1) ドキュメント検証 → gatherDocumentInfo() でファイル情報を取得
    2) parseFileName() でセグメント配列に分解（区切り文字も保持）
    3) createDialog() でダイアログを構築（buildModePanel + buildFilenamePanel + buildOptionsPanel）
    4) buildFinalName() で UI 状態と segments から最終ファイル名を構築
    5) ensureTargetFolder() で未保存ドキュメントは保存先フォルダを取得
    6) confirmOverwriteIfExists() で既存同名の上書きを確認
    7) executeOutput() で選択モード（rename / saveAs / copy）に応じた出力を実行
    
    ### 更新履歴：
    
    - v1.0 (2026-05-27) : 初期バージョン
    
    ---
    
    ### Script name:
    
    FileNameManager.jsx — Rename or Save As the Current Document
    
    ### Overview:
    
    - Renames the active InDesign document (true rename), saves as a new file, or saves a copy.
    - Destination is always the same folder as the current file; for unsaved docs, a destination folder is chosen.
    - Decomposes the filename into base / date / text / version segments, preserving their original order.
      - base: displayed read-only as "Base" (the fixed prefix)
      - date: "Timestamp" radio selects "None / YYYYMMDD" (default YYYYMMDD)
      - text: editable via the "Title" field
      - version: "Version" radio selects "None / -vN / -v0N" (default -vN)
    - Modes: "Rename", "Save As" (default), and "Save a Copy".
    - Separator: choose from "No Change", `-`, or `_`. Choosing `-` or `_` unifies separators across the filename.
    - UI is composed of "Mode" and "File Name" panels (with an "Options" sub-panel inside) and tooltips on each element.
    
    ### Key features:
    
    - Rename (saves with the new name, then removes the original)
    - Save As (keeps the original; the active document switches to the new file; default)
    - Save a Copy (keeps the original, creates a separate copy; the active document is unchanged)
    - Segment-based editing with order preservation (base / date / text / version)
    - Timestamp / version radios switch the format; "None" removes the segment
    - Version radios bump the existing version by +1 (or create a new one if missing)
    - OK button label switches between Rename / Save / Save a Copy based on the selected mode
    
    ### Flow:
    
    1) gatherDocumentInfo() collects file info
    2) parseFileName() splits the name into ordered segments (separators preserved)
    3) createDialog() composes the dialog (buildModePanel + buildFilenamePanel + buildOptionsPanel)
    4) buildFinalName() composes the final filename from UI state and segments
    5) ensureTargetFolder() prompts for a destination folder if the doc is unsaved
    6) confirmOverwriteIfExists() confirms before overwriting an existing file
    7) executeOutput() runs the action for the selected mode (rename / saveAs / copy)
    
    ### Changelog:
    
    - v1.0 (2026-05-27): Initial release
    
    */

    (function () {

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
        var PANEL_MARGINS = [15, 20, 15, 10];
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
                timestampNone: { ja: "つけない", en: "None" },
                timestampDate: { ja: "YYYYMMDD", en: "YYYYMMDD" },
                versionNone: { ja: "つけない", en: "None" },
                versionShort: { ja: "-vN", en: "-vN" },
                versionPadded: { ja: "-v0N", en: "-v0N" }
            },
            checkbox: {
                newName: { ja: "タイトル：", en: "Title:" }
            },
            label: {
                currentName: { ja: "現在：", en: "Current:" },
                finalName: { ja: "変更後：", en: "Final:" },
                base: { ja: "ベース：", en: "Base:" },
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
                newName: {
                    ja: "ファイル名のタイトル部分を入力した値に置換します。",
                    en: "Replaces the title part of the filename with this value."
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

        /* ドット記法でローカライズ済み文字列を取得 / Get the localized string by dotted key (e.g. L('dialog.title')) */
        function L(path) {
            var parts = String(path).split('.');
            var entry = LABELS;
            for (var i = 0; i < parts.length; i++) {
                entry = entry && entry[parts[i]];
                if (!entry) return path;
            }
            return entry[lang] || entry.en || path;
        }

        // =========================================
        // ヘルパー / Helpers
        // =========================================

        var VERSION_TOKEN_RE = /^[vV]\d+$/;   // v123 / V123

        /* 月（1-12）と日（1-31）の妥当性 / Validate month (1-12) and day (1-31) */
        function isValidMonthDay(monthStr, dayStr) {
            var m = parseInt(monthStr, 10);
            var d = parseInt(dayStr, 10);
            return m >= 1 && m <= 12 && d >= 1 && d <= 31;
        }

        /* YYYYMMDD（8 桁）または YYMMDD（6 桁）の日付トークンか / Whether a token is a YYYYMMDD or YYMMDD date */
        function isDateToken(token) {
            var s = String(token);
            if (/^\d{8}$/.test(s)) return isValidMonthDay(s.substring(4, 6), s.substring(6, 8));
            if (/^\d{6}$/.test(s)) return isValidMonthDay(s.substring(2, 4), s.substring(4, 6));
            return false;
        }

        /* "v123" 形式のバージョントークンか / Whether a token is a v+digits version */
        function isVersionToken(token) {
            return VERSION_TOKEN_RE.test(String(token));
        }

        /* ファイル名を順序付きセグメント配列に分解。各 segment は前にあった区切り文字も保持
           / Decompose into ordered segments; each segment stores the preceding separator */
        function parseFileName(name) {
            var split = String(name).split(/([-_])/); // ["handout","-","Adobe","-","20260422"]
            var segments = [];
            var textBuffer = [];      // token と区切りを交互に蓄積（join('') で結合）
            var textLeadingSep = '';  // テキスト segment の直前に置く区切り
            var hasDate = false, hasVersion = false;
            var currentSep = '';

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

        /* segments から最初に出現する kind の value を取得 / Get the value of the first segment of the given kind */
        function getFirstSegmentValue(segments, kind) {
            for (var i = 0; i < segments.length; i++) {
                if (segments[i].kind === kind) return segments[i].value;
            }
            return '';
        }

        /* 文字列を左 0 パディング / Left-pad a string with zeros to the given width */
        function padLeft(str, width) {
            while (str.length < width) str = '0' + str;
            return str;
        }

        /* 今日の日付を YYYYMMDD 形式で / Today's date as YYYYMMDD */
        function todayTimestamp() {
            var d = new Date();
            return String(d.getFullYear()) +
                padLeft(String(d.getMonth() + 1), 2) +
                padLeft(String(d.getDate()), 2);
        }

        /* バージョン文字列を +1。mode='padded' でゼロ埋め（最低 2 桁）。元バージョンが無ければ新規付与（v2 / v02）
           / Bump the version string by +1. 'padded' zero-pads to at least 2 digits. Returns v2 / v02 if no original */
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

        /* % エンコードを 1 回デコード / Decode a percent-encoded string once (best-effort) */
        function decodePercentEncoded(str) {
            str = String(str);
            if (str.indexOf('%') === -1) return str;
            try {
                return decodeURIComponent(str);
            } catch (e) {
                return str;
            }
        }

        /* 拡張子を除去 / Strip a trailing file extension */
        function stripExtension(name) {
            var dot = name.lastIndexOf('.');
            return (dot > 0) ? name.substring(0, dot) : name;
        }

        /* ファイル名として無効な文字を除去し前後空白をトリム / Strip invalid filename chars and trim whitespace */
        function sanitizeFilename(str) {
            return String(str).replace(/[\\\/:*?"<>|]/g, '').replace(/^\s+|\s+$/g, '');
        }

        /* 保存先ダイアログ。.indd 拡張子を補正して File を返す。キャンセルは null / Show save dialog, normalize to .indd; null on cancel */
        function pickInddDestination(promptLabel) {
            var chosen = File.saveDialog(promptLabel, '*.indd');
            if (!chosen) return null;
            var file = File(chosen.fsName);
            if (!/\.indd$/i.test(file.name)) {
                file = File(file.parent.fsName + '/' + stripExtension(file.name) + '.indd');
            }
            return file;
        }

        /* プリセット保存用ファイル / Preferences file (key=value lines) */
        function getPrefsFile() {
            return File(Folder.userData.fsName + '/FileNameManager-prefs.txt');
        }

        /* 前回保存したプリセットを読み込み（無ければ空オブジェクト） / Load previously saved prefs (or empty if none) */
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

        /* プリセットを保存（失敗時は黙って継続） / Save prefs (silently ignore failure) */
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

        /* ドキュメントから現在のファイル情報を収集 / Collect file info from the active document */
        function gatherDocumentInfo(doc) {
            var fullName = doc.fullName;
            var currentName = decodePercentEncoded(fullName ? fullName.name : doc.name);
            return {
                currentName: currentName,
                baseName: stripExtension(currentName),
                fsPath: fullName ? fullName.fsName : null,
                folder: fullName ? fullName.parent : null
            };
        }

        /* 保存先フォルダを確定（未保存なら saveDialog で取得）。キャンセル時は null / Resolve the target folder (prompt if needed); null on cancel */
        function ensureTargetFolder(folder) {
            if (folder) return folder;
            var picked = pickInddDestination(L('message.chooseDestination'));
            return picked ? picked.parent : null;
        }

        /* 既存同名（自分自身を除く）の上書き確認。OK なら true / Confirm overwrite for an existing file (excluding self); true if approved */
        function confirmOverwriteIfExists(destFile, originalFsPath) {
            if (!destFile.exists) return true;
            if (destFile.fsName === originalFsPath) return true;
            return confirm(L('message.confirmOverwrite') + '\n\n' + destFile.fsName);
        }

        /* 旧ファイルを削除（失敗時は黙って継続） / Remove the original file (silently ignore failure) */
        function removeOriginalFile(originalFsPath, destFsPath) {
            if (!originalFsPath || originalFsPath === destFsPath) return;
            var file = File(originalFsPath);
            if (!file.exists) return;
            try { file.remove(); } catch (e) { /* 削除できない場合は黙って継続 */ }
        }

        /* モード別の出力処理（rename / saveAs / copy） / Execute the output according to the selected mode */
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

        /* segments と UI 状態から最終ファイル名（拡張子なし）を構築。SEGMENT_ORDER に従う
           / Build the final filename from segments and UI state, following SEGMENT_ORDER */
        function buildFinalName(segments, uiState) {
            // 区切り記号: 明示選択があればそれを、無ければ元のファイル名で優勢な区切りを使う
            var defaultSep;
            if (uiState.separator === '-' || uiState.separator === '_') {
                defaultSep = uiState.separator;
            } else {
                defaultSep = dominantSeparator(segments);
            }

            function valueForKind(kind) {
                if (kind === 'base') {
                    return getFirstSegmentValue(segments, 'base');
                }
                if (kind === 'title') {
                    return uiState.useNewName ? sanitizeFilename(uiState.newNameText) : '';
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

        /* segments から優勢な区切り記号を返す。同数なら "-" / Dominant separator across segments (defaults to "-") */
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

        /* パネルの共通設定（向き・整列・余白・子要素間隔） / Apply shared panel settings (orientation, alignment, margins, spacing) */
        function setupPanel(panel, spacing) {
            panel.orientation = 'column';
            panel.alignChildren = ['fill', 'top'];
            panel.alignment = 'fill';
            panel.margins = PANEL_MARGINS;
            panel.spacing = (typeof spacing === 'number') ? spacing : PANEL_SPACING;
        }

        /* モード選択パネルを構築（リネーム / 別名で保存 / コピーを保存） / Build the mode panel (Rename / Save As / Save a Copy) */
        function buildModePanel(parent, prefs) {
            var panel = parent.add('panel', undefined, L('panel.mode'));
            setupPanel(panel);
            var renameRadio = panel.add('radiobutton', undefined, L('radio.rename'));
            renameRadio.helpTip = L('tip.rename');
            var saveAsRadio = panel.add('radiobutton', undefined, L('radio.saveAs'));
            saveAsRadio.helpTip = L('tip.saveAs');
            var saveCopyRadio = panel.add('radiobutton', undefined, L('radio.saveCopy'));
            saveCopyRadio.helpTip = L('tip.saveCopy');
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

        /* ファイル名パネルを構築（現在名・変更後名のみ） / Build the file-name panel (current / final only) */
        function buildFilenamePanel(parent, currentName) {
            var panel = parent.add('panel', undefined, L('panel.filename'));
            setupPanel(panel);

            var currentNameRow = panel.add('group');
            currentNameRow.orientation = 'row';
            var currentNameLabel = currentNameRow.add('statictext', undefined, L('label.currentName'));
            currentNameRow.add('statictext', undefined, currentName);

            var finalNameRow = panel.add('group');
            finalNameRow.orientation = 'row';
            var finalNameLabel = finalNameRow.add('statictext', undefined, L('label.finalName'));
            // 「変更後：」は statictext のためレイアウト後にサイズ固定。
            // 現在のファイル名と「入力フィールド + 余白」の大きい方を確保しておく
            var finalNameValue = finalNameRow.add('statictext', undefined, currentName + '.indd');
            var currentNameWidth = panel.graphics.measureString(currentName + '.indd').width;
            finalNameValue.preferredSize.width = Math.max(currentNameWidth + 20, NEW_NAME_FIELD_WIDTH + 150);

            alignLabelWidths(panel, [
                L('label.currentName'),
                L('label.finalName')
            ], [currentNameLabel, finalNameLabel]);

            return {
                panel: panel,
                finalNameValue: finalNameValue
            };
        }

        /* オプションパネルを構築（ベース表示・タイトル編集・タイムスタンプ・バージョン番号・区切り） / Build the options panel */
        function buildOptionsPanel(parent, segments, prefs) {
            var panel = parent.add('panel', undefined, L('panel.options'));
            setupPanel(panel);

            var baseValue = getFirstSegmentValue(segments, 'base');
            var initialText = getFirstSegmentValue(segments, 'text');

            // ベース（読み取り専用表示）
            var baseRow = panel.add('group');
            baseRow.orientation = 'row';
            var baseLabel = baseRow.add('statictext', undefined, L('label.base'));
            baseRow.add('statictext', undefined, baseValue);

            // タイトル（チェックボックス + 編集フィールド）
            var newNameRow = panel.add('group');
            newNameRow.orientation = 'row';
            var newNameCheckbox = newNameRow.add('checkbox', undefined, L('checkbox.newName'));
            newNameCheckbox.helpTip = L('tip.newName');
            var newNameField = newNameRow.add('edittext', undefined, initialText);
            newNameField.preferredSize.width = NEW_NAME_FIELD_WIDTH;
            newNameField.helpTip = L('tip.newName');
            newNameCheckbox.value = !!initialText;
            if (!initialText) {
                newNameCheckbox.enabled = false;
                newNameField.enabled = false;
            }

            // タイムスタンプ（つけない / YYYYMMDD。デフォルト YYYYMMDD）
            var timestampRow = panel.add('group');
            timestampRow.orientation = 'row';
            var timestampLabel = timestampRow.add('statictext', undefined, L('label.timestamp'));
            timestampLabel.helpTip = L('tip.timestamp');
            var timestampNoneRadio = timestampRow.add('radiobutton', undefined, L('radio.timestampNone'));
            timestampNoneRadio.helpTip = L('tip.timestamp');
            var timestampDateRadio = timestampRow.add('radiobutton', undefined, L('radio.timestampDate'));
            timestampDateRadio.helpTip = L('tip.timestamp');
            var initialTimestamp = (prefs && prefs.timestamp === 'none') ? 'none' : 'date';
            timestampNoneRadio.value = (initialTimestamp === 'none');
            timestampDateRadio.value = (initialTimestamp === 'date');

            // バージョン番号（つけない / -vN / -v0N。デフォルト -vN）
            var versionRow = panel.add('group');
            versionRow.orientation = 'row';
            var versionLabel = versionRow.add('statictext', undefined, L('label.version'));
            versionLabel.helpTip = L('tip.version');
            var versionNoneRadio = versionRow.add('radiobutton', undefined, L('radio.versionNone'));
            versionNoneRadio.helpTip = L('tip.version');
            var versionShortRadio = versionRow.add('radiobutton', undefined, L('radio.versionShort'));
            versionShortRadio.helpTip = L('tip.version');
            var versionPaddedRadio = versionRow.add('radiobutton', undefined, L('radio.versionPadded'));
            versionPaddedRadio.helpTip = L('tip.version');
            var initialVersion = (prefs && (prefs.version === 'none' || prefs.version === 'padded')) ? prefs.version : 'short';
            versionNoneRadio.value = (initialVersion === 'none');
            versionShortRadio.value = (initialVersion === 'short');
            versionPaddedRadio.value = (initialVersion === 'padded');

            // 区切り記号（変更しない / - / _ 横並び。デフォルト "-"）
            var separatorRow = panel.add('group');
            separatorRow.orientation = 'row';
            var separatorLabel = separatorRow.add('statictext', undefined, L('label.separator'));
            separatorLabel.helpTip = L('tip.separator');
            var noChangeRadio = separatorRow.add('radiobutton', undefined, L('radio.noChange'));
            noChangeRadio.helpTip = L('tip.separator');
            var dashRadio = separatorRow.add('radiobutton', undefined, '-');
            dashRadio.helpTip = L('tip.separator');
            var underscoreRadio = separatorRow.add('radiobutton', undefined, '_');
            underscoreRadio.helpTip = L('tip.separator');
            var initialSeparator = (prefs && (prefs.separator === '' || prefs.separator === '_')) ? prefs.separator : '-';
            noChangeRadio.value = (initialSeparator === '');
            dashRadio.value = (initialSeparator === '-');
            underscoreRadio.value = (initialSeparator === '_');

            alignLabelWidths(panel, [
                L('label.base'),
                L('checkbox.newName')
            ], [baseLabel, newNameCheckbox]);

            return {
                panel: panel,
                newNameCheckbox: newNameCheckbox,
                newNameField: newNameField,
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
                }
            };
        }

        /* labelTexts の最大幅に controls の幅を揃える。checkbox/radiobutton が含まれていれば
           インジケーター分（+20px）を全コントロールに加算して右端を揃える
           / Align controls' widths to the widest label. If any control is a checkbox/radiobutton,
           add indicator width (+20px) to all so their right edges align */
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

        /* ダイアログ全体を組み立て、イベント配線とプレビューを行う / Compose the full dialog, wire events, and run live preview */
        function createDialog(segments, currentName, prefs) {
            var dialog = new Window('dialog', L('dialog.title') + ' ' + SCRIPT_VERSION);
            dialog.opacity = DIALOG_OPACITY;
            dialog.orientation = 'column';
            dialog.alignChildren = 'fill';

            var mode = buildModePanel(dialog, prefs);
            var filename = buildFilenamePanel(dialog, currentName);
            // ファイル名行とオプションパネルの間に余白を入れる
            var optionsSpacer = filename.panel.add('group');
            optionsSpacer.preferredSize = [-1, 5];
            var options = buildOptionsPanel(filename.panel, segments, prefs);

            // ---- ライブプレビュー ----
            function currentUIState() {
                return {
                    useNewName: options.newNameCheckbox.value,
                    newNameText: options.newNameField.text,
                    timestamp: options.getTimestamp(),
                    version: options.getVersion(),
                    separator: options.getSeparator()
                };
            }

            function refreshPreviews() {
                var finalBase = buildFinalName(segments, currentUIState());
                filename.finalNameValue.text = finalBase + '.indd';
            }

            options.newNameCheckbox.onClick = refreshPreviews;
            options.newNameField.onChanging = refreshPreviews;
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
            buttonGroup.add('button', undefined, L('button.cancel'), { name: 'cancel' });
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

        /* エントリポイント。ダイアログを開き、選択モードに応じた出力を実行 / Entry point: open the dialog and execute the selected mode */
        function main() {
            if (app.documents.length === 0) {
                alert(L('message.noDoc'));
                return;
            }

            var doc = app.activeDocument;
            var info = gatherDocumentInfo(doc);
            var segments = parseFileName(info.baseName);
            var prefs = loadPrefs();

            var ui = createDialog(segments, info.currentName, prefs);
            if (ui.dialog.show() !== 1) return; // キャンセル

            var uiState = ui.getUIState();
            var newBaseName = buildFinalName(segments, uiState);
            if (!newBaseName) {
                alert(L('message.emptyName'));
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
                    timestamp: uiState.timestamp,
                    version: uiState.version,
                    separator: uiState.separator
                });
            } catch (e) {
                alert(L('message.saveFailed') + '\n' + e);
            }
        }

        main();

    })();
