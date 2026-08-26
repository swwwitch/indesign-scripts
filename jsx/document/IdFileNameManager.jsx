#target indesign

/*

### 概要

アクティブなドキュメントのファイル名を、ベース・サブテキスト・ステータス・日付・連番・バージョンのセグメント単位で編集し、リネーム／別名保存／コピー保存を行います。

詳細は README を参照してください。

### Overview

Edits the active document's file name segment by segment - base, subtext, status, date, sequence number and version - then renames it, saves it under a new name, or saves a copy.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdFileNameManager";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-05-29";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdFileNameManager.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdFileNameManager.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nc88dd887eb1c"; /* 紹介記事 / article URL */

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

        var SCRIPT_VERSION = "v1.3.1";

        // =========================================
        // ユーザー設定 / User Settings
        // =========================================

        /* 整形機能のオン/オフ。false にすると該当 UI とロジックを丸ごとスキップ
           / Feature switches. Set to false to omit the UI and skip the logic */
        var FEATURE_STATUS = true;        // ステータス dropdown と検出 / Status dropdown + detection
        var FEATURE_SORT = true;          // ソートパネル + 並び順カスタマイズ / Sort panel + custom segment order
        var FEATURE_SEPARATOR = '-';      // 区切り記号統一: '-' / '_' で有効化＋既定値、false で無効 / '-' or '_' enables with that default; false disables
        var FEATURE_DOT_NORMALIZE = true; // "." を区切り記号にあわせて置換（要 FEATURE_SEPARATOR） / "." normalization with the chosen separator
        var FEATURE_NFC = true;           // 濁点・半濁点の NFC 結合 / Combine separated dakuten/handakuten (NFC)
        var FEATURE_CLEAN = true;         // クリーンなファイル名（OS 禁止文字 + 絵文字・機種依存文字の処理） / Clean filename (OS-invalid + emoji + platform-dependent)
        var FEATURE_TRANSLITERATE = true; // 法人略記・丸数字などを ASCII 相当に変換 / Transliterate corp abbrev / circled numbers / etc.
        var FEATURE_HALFWIDTH_KANA = true; // 半角カナ → 全角カナ変換（クリーンが '-' / '_' のときのみ有効） / Halfwidth-kana → fullwidth (active only when Clean is '-' or '_')
        var FEATURE_PAGE = true;          // 連番（pageNN）セグメント / Page-number segment (pageNN)

        /* ファイル名（拡張子込み）の UTF-8 バイト長の上限。超過時は確認ダイアログを出して続行可
           / Max UTF-8 byte length of the basename + extension; exceeding it triggers a confirm dialog */
        var FEATURE_MAX_FILENAME_BYTES = 240;

        /* rename モードで元ファイルを削除するときに、即削除ではなく ~/.Trash に移動
           / In rename mode, move the original to ~/.Trash instead of removing it outright */
        var FEATURE_USE_TRASH = true;

        /* 文字 → 置換文字列のマップ。法人略記、丸数字（白・黒・括弧）、ローマ数字、略記号、単位、ダッシュ類
           / Char → replacement map: corporate abbrev, circled (white/black/parenthesized), Roman numerals, abbrev symbols, units, dashes */
        var TRANSLITERATE_MAP = {
            // 法人略記
            '㈱': '株', '㈲': '有', '㈹': '代', '㈳': '社',
            '㈵': '特', '㈶': '財', '㈻': '学', '㍿': '株式会社',
            // 白丸数字 ①-⑳ (U+2460-U+2473)
            '①': '1', '②': '2', '③': '3', '④': '4', '⑤': '5',
            '⑥': '6', '⑦': '7', '⑧': '8', '⑨': '9', '⑩': '10',
            '⑪': '11', '⑫': '12', '⑬': '13', '⑭': '14', '⑮': '15',
            '⑯': '16', '⑰': '17', '⑱': '18', '⑲': '19', '⑳': '20',
            // 黒丸数字 ❶-❿ (U+2776-U+277F) + ⓫-⓴ (U+24EB-U+24F4)
            '❶': '1', '❷': '2', '❸': '3', '❹': '4', '❺': '5',
            '❻': '6', '❼': '7', '❽': '8', '❾': '9', '❿': '10',
            '⓫': '11', '⓬': '12', '⓭': '13', '⓮': '14', '⓯': '15',
            '⓰': '16', '⓱': '17', '⓲': '18', '⓳': '19', '⓴': '20',
            // 括弧数字 ⑴-⒇ (U+2474-U+2487)
            '⑴': '1', '⑵': '2', '⑶': '3', '⑷': '4', '⑸': '5',
            '⑹': '6', '⑺': '7', '⑻': '8', '⑼': '9', '⑽': '10',
            '⑾': '11', '⑿': '12', '⒀': '13', '⒁': '14', '⒂': '15',
            '⒃': '16', '⒄': '17', '⒅': '18', '⒆': '19', '⒇': '20',
            // ローマ数字 大文字 Ⅰ-Ⅻ (U+2160-U+216B)
            'Ⅰ': '1', 'Ⅱ': '2', 'Ⅲ': '3', 'Ⅳ': '4', 'Ⅴ': '5',
            'Ⅵ': '6', 'Ⅶ': '7', 'Ⅷ': '8', 'Ⅸ': '9', 'Ⅹ': '10',
            'Ⅺ': '11', 'Ⅻ': '12',
            // ローマ数字 小文字 ⅰ-ⅻ (U+2170-U+217B)
            'ⅰ': '1', 'ⅱ': '2', 'ⅲ': '3', 'ⅳ': '4', 'ⅴ': '5',
            'ⅵ': '6', 'ⅶ': '7', 'ⅷ': '8', 'ⅸ': '9', 'ⅹ': '10',
            'ⅺ': '11', 'ⅻ': '12',
            // 略記号
            '℡': 'TEL', '№': 'No',
            // 単位記号
            '㎜': 'mm', '㎝': 'cm', '㎞': 'km',
            '㎎': 'mg', '㎏': 'kg',
            '㎡': 'm2', '㎥': 'm3',
            // チルダ類
            '〜': '~', '～': '~',  // 〜 WAVE DASH / ～ FULLWIDTH TILDE
            // ハイフン・ダッシュ類
            '－': '-',  // － FULLWIDTH HYPHEN-MINUS
            '‐': '-',  // ‐ HYPHEN
            '‑': '-',  // ‑ NON-BREAKING HYPHEN
            '‒': '-',  // ‒ FIGURE DASH
            '–': '-',  // – EN DASH
            '—': '-',  // — EM DASH
            '―': '-',  // ― HORIZONTAL BAR
            '−': '-'   // − MINUS SIGN
        };

        /* 出力時のセグメント順序。base / title / status / timestamp / page / version。
           FEATURE_STATUS=false なら status を、FEATURE_PAGE=false なら page を除外
           / Output segment order; "status" / "page" are dropped when their FEATURE_* flags are false */
        var SEGMENT_ORDER = (function () {
            var order = ['base', 'title'];
            if (FEATURE_STATUS) order.push('status');
            order.push('timestamp');
            if (FEATURE_PAGE) order.push('page');
            order.push('version');
            return order;
        })();

        /* ステータス選択肢。value がファイル名に入り、ja / en が UI 表示。先頭は「なし」。
           ja が '---' の項目は dropdown 上の区切り線として表示（選択不可）
           / Status choices: `value` is what enters the filename; `ja` / `en` are UI labels.
           Items with `ja === '---'` render as an unselectable divider. */
        var STATUS_ITEMS = [
            { value: '', ja: 'なし', en: 'None' },
            { value: 'wip', ja: 'wip：作業中・仕掛かり中', en: 'wip: Work in progress' },
            { value: 'draft', ja: 'draft：下書き・ラフ・素案', en: 'draft: Draft / Rough' },
            { value: 'review', ja: 'review：レビュー待ち', en: 'review: Awaiting review' },
            { value: 'revised', ja: 'revised：改訂版・修正反映版', en: 'revised: Revised' },
            { value: 'updated', ja: 'updated：更新版', en: 'updated: Updated' },
            { value: 'fixed', ja: 'fixed：修正完了', en: 'fixed: Fix completed' },
            { value: 'approved', ja: 'approved：承認済み・確定', en: 'approved: Approved / Final' },
            { value: 'rejected', ja: 'rejected：ボツ・不採用案', en: 'rejected: Rejected' },
            { value: 'archived', ja: 'archived：保管用', en: 'archived: Archive' },
            { value: '', ja: '---', en: '---' },
            { value: 'flattened', ja: 'flattened：レイヤー結合済み（.psd）', en: 'flattened: Flattened (.psd)' },
            { value: 'outlined', ja: 'outlined：アウトライン済み（.ai）', en: 'outlined: Outlined (.ai)' }
        ];

        /**
         * ステータス表記の区切り文字かどうかを判定する
         * @param {string} character 判定する文字
         * @returns {boolean} 区切り文字なら true
         */
        function isStatusDivider(item) {
            return item && item.ja === '---';
        }

        var NEW_NAME_FIELD_WIDTH = 250;
        var PANEL_SPACING = 8;

        var DIALOG_OPACITY = 0.98;

        // =========================================
        // ローカライズ / Localization
        // =========================================

        var currentLanguage = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

        /* 日英ラベル定義 / Japanese-English label definitions */

        var LABELS = {
            dialog: {
                title: { ja: "ファイル名を変更して保存", en: "Rename and Save" }
            },
            panel: {
                mode: { ja: "動作", en: "Mode" },
                opMode: { ja: "モード", en: "Scope" },
                filename: { ja: "ファイル名プレビュー", en: "File Name Preview" },
                options: { ja: "ファイル名の設定", en: "Filename Settings" },
                sort: { ja: "構成要素の順序", en: "Segment Order" }
            },
            radio: {
                rename: { ja: "元ファイルをリネーム", en: "Rename Original" },
                saveAs: { ja: "別名で保存", en: "Save As" },
                saveCopy: { ja: "コピー（複製）を保存", en: "Save a Copy" },
                opVersionOnly: { ja: "バージョンのみ", en: "Version Only" },
                opFull: { ja: "全体", en: "Full" },
                noChange: { ja: "変更しない", en: "No Change" },
                titleNone: { ja: "なし", en: "None" },
                titleParent: { ja: "親フォルダー", en: "Parent Folder" },
                titleGrandparent: { ja: "2 階層上のフォルダー", en: "Grandparent Folder" },
                titleCustom: { ja: "指定", en: "Custom" },
                timestampNone: { ja: "なし", en: "None" },
                timestampDate: { ja: "YYYYMMDD", en: "YYYYMMDD" },
                timestampDateDash: { ja: "YYYY-MM-DD", en: "YYYY-MM-DD" },
                timestampWithTime: { ja: "時刻も付与", en: "Append HHMM" },
                versionNone: { ja: "なし", en: "None" },
                versionShort: { ja: "v1, v2…", en: "v1, v2…" },
                versionPadded: { ja: "v01, v02…", en: "v01, v02…" },
                versionPaddedWide: { ja: "v001, v002…", en: "v001, v002…" },
                changeToDash: { ja: "-に変更", en: "Change to -" },
                changeToUnderscore: { ja: "_に変更", en: "Change to _" },
                nfcCombine: { ja: "結合する", en: "Combine" },
                cleanRemove: { ja: "削除する", en: "Remove" },
                translitRemove: { ja: "削除する", en: "Remove" },
                translitConvert: { ja: "変換する", en: "Convert" },
                halfwidthKanaConvert: { ja: "半角カナ → 全角", en: "Half-width → Full-width Kana" },
                pageEnable: { ja: "連番を付与", en: "Append Sequence" },
                pagePad2: { ja: "01, 02…", en: "01, 02…" },
                pagePad3: { ja: "001, 002…", en: "001, 002…" },
                sortOff: { ja: "標準順", en: "Default" },
                sortCurrent: { ja: "現在のファイル名に準じる", en: "Match Current" },
                sortOn: { ja: "カスタム順", en: "Custom" }
            },
            label: {
                currentName: { ja: "現在", en: "Current" },
                finalName: { ja: "保存後の名前", en: "Saved Name" },
                base: { ja: "ベース", en: "Base" },
                title: { ja: "サブテキスト", en: "Project Name" },
                status: { ja: "ステータス", en: "Status" },
                timestamp: { ja: "タイムスタンプ", en: "Timestamp" },
                version: { ja: "バージョン番号", en: "Version" },
                page: { ja: "連番", en: "Sequence" },
                separator: { ja: "区切り記号", en: "Separator" },
                nfc: { ja: "濁点・半濁点の正規化", en: "NFC Normalization" },
                clean: { ja: "クリーンなファイル名", en: "Clean Filename" },
                translit: { ja: "丸数字や法人略記など", en: "Symbols" }
            },
            tip: {
                rename: {
                    ja: "新しい名前で保存したあと、元ファイルをゴミ箱に移します（実質的にリネーム）。元ファイルを参照している他のドキュメント（配置 .ai／InDesign のリンクなど）はリンク切れになります。",
                    en: "Saves with the new name, then moves the original to the Trash (effectively a rename). Documents that reference the original file (placed .ai or InDesign links) will lose the link."
                },
                saveAs: {
                    ja: "新しい名前で保存します。元ファイルは残り、作業中のドキュメントが新ファイルに切り替わります。",
                    en: "Saves with the new name. The original file is kept, and the active document switches to the new file."
                },
                saveCopy: {
                    ja: "元ファイルに上書き保存したうえで、別名のコピーを作成します。作業中のドキュメントは元ファイルのまま残ります。",
                    en: "Saves the original, then creates a copy with the new name. The active document remains the original file."
                },
                base: {
                    ja: "ファイル名の先頭部分。空欄にすると省略されます。",
                    en: "Leading part of the filename. Leave empty to omit."
                },
                title: {
                    ja: "ファイル名に追加する案件名・補足テキスト。「指定」で入力欄の文字列を使用します。",
                    en: "Choose how to set the project name in the filename. With \"Custom\", the entered text is used."
                },
                status: {
                    ja: "ファイル名に挿入する制作ステータスを選択します。「：」より前の文字列が入ります。",
                    en: "Choose a production status to insert. Only the text before \":\" is used in the filename."
                },
                timestamp: {
                    ja: "タイムスタンプの形式を選択。「なし」で元の日付があっても削除します。",
                    en: "Choose timestamp format. \"None\" removes any existing date."
                },
                timestampWithTime: {
                    ja: "タイムスタンプの末尾に時刻（HHMM）を付加します。1 日に複数版を出すときに便利。",
                    en: "Append the current time (HHMM) to the timestamp. Useful for multiple versions per day."
                },
                version: {
                    ja: "バージョン番号の形式。v1/v2 はパディング無し、v01/v02 は 2 桁、v001/v002 は 3 桁ゼロ埋め。既存の v 番号は +1、無い場合は v1/v01/v001 を付与。「なし」で削除。",
                    en: "Version format. v1/v2 has no padding, v01/v02 is 2-digit, v001/v002 is 3-digit zero-padded. An existing v-number is bumped by +1; otherwise v1/v01/v001 is added. \"None\" removes."
                },
                page: {
                    ja: "「page01」「page001」のような連番をファイル名に追加します。プレフィックス（page 等）と桁数（01 / 001）を選択。保存時に同フォルダ内の最大連番 +1 へ自動繰り上げ。",
                    en: "Append a sequence such as page01 / page001. Configure prefix (e.g. page) and width (01 / 001). On save, bumps to (max in folder) + 1 if collisions exist."
                },
                separator: {
                    ja: "ファイル名全体の区切り記号の扱いを選択します。「-」「_」「.」が対象。YYYY-MM-DD のタイムスタンプは保護されます。",
                    en: "Choose how separators in the filename are handled. Targets `-`, `_`, and `.`. YYYY-MM-DD timestamps are preserved."
                },
                sort: {
                    ja: "並び順を選択。「現在のファイル名に準じる」は検出された要素のみ、「カスタム順」は［編集］で並び替え。",
                    en: "Choose order. \"Match Current\" uses only detected segments; \"Custom\" enables [Edit]."
                },
                nfc: {
                    ja: "ファイル名に分離した濁点・半濁点が含まれる場合、結合済み文字（NFC）に正規化します。",
                    en: "Normalize separated dakuten/handakuten to combined characters (NFC)."
                },
                clean: {
                    ja: "OS で使えない文字（\\ / : * ? \" < > |）、絵文字・機種依存文字（㈱・①・㌔ など）、スペースの扱いをまとめて選択します。連続するスペースは 1 つにまとめてから処理。",
                    en: "How to handle OS-invalid characters (\\ / : * ? \" < > |), emoji, platform-dependent chars (㈱, ①, ㌔), and spaces. Consecutive spaces collapse to one before replacement."
                },
                translit: {
                    ja: "法人略記（㈱→株）、丸数字（①→1）、ローマ数字（Ⅰ→1）、TEL/No、単位記号（㎜→mm）、ダッシュ類を ASCII 相当に変換します。",
                    en: "Convert corporate abbreviations (㈱→株), circled numbers (①→1), Roman numerals, TEL/No, units (㎜→mm), and dashes to ASCII equivalents."
                },
                halfwidthKana: {
                    ja: "「-」または「_」を選んだとき、半角カタカナを全角カタカナに変換します（濁点・半濁点付きは結合）。",
                    en: "When `-` or `_` is selected, convert half-width katakana to full-width (voiced/semi-voiced marks are merged)."
                }
            },
            button: {
                cancel: { ja: "キャンセル", en: "Cancel" },
                sort: { ja: "編集", en: "Edit" }
            },
            sort: {
                title: { ja: "ファイル名の並び順", en: "Filename Order" },
                hint: { ja: "選択中の項目を ↑↓ で並び替え", en: "Select an item and use ↑↓ to reorder" }
            },
            message: {
                noDoc: { ja: "ドキュメントが開かれていません", en: "No document is open." },
                emptyName: { ja: "ファイル名が空です", en: "File name is empty." },
                chooseDestination: { ja: "保存先フォルダを指定", en: "Choose destination folder" },
                confirmOverwrite: {
                    ja: "同名ファイルが存在します。上書きしますか？",
                    en: "A file with the same name exists. Overwrite?"
                },
                confirmTooLong: {
                    ja: "ファイル名が長すぎる可能性があります（{bytes} バイト / 上限 {limit} バイト）。このまま続行しますか？",
                    en: "The filename may be too long ({bytes} bytes / limit {limit}). Continue anyway?"
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
            return entry[currentLanguage] || entry.en || path;
        }

        /**
         * コロン付きラベルを取得する（日本語は全角コロン、英語は半角コロン）
         * @param {string} path 例: "label.separator"
         * @returns {string} コロンを付与したラベル文字列
         */
        function getLabelWithColon(path) {
            return L(path) + (currentLanguage === 'ja' ? '：' : ':');
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
         * ステータスセグメントとして解釈できるかを判定する
         * @param {string} token 判定する文字列
         * @returns {string|null} ステータス名。該当しない場合は null
         */
        function matchStatusToken(token) {
            var s = String(token).toLowerCase();
            for (var i = 1; i < STATUS_ITEMS.length; i++) {
                if (isStatusDivider(STATUS_ITEMS[i])) continue;
                if (STATUS_ITEMS[i].value === s) return STATUS_ITEMS[i].value;
            }
            return '';
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
            var hasDate = false, hasVersion = false, hasStatus = false;
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
                } else if (FEATURE_STATUS && !hasStatus && matchStatusToken(token)) {
                    flushText();
                    segments.push({ kind: 'status', value: matchStatusToken(token), sep: currentSep });
                    hasStatus = true;
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
         * 分断されたテキストセグメントをつなぎ直す
         * @param {Array<object>} segments セグメントの配列
         * @returns {Array<object>} 統合後のセグメント
         */
        function mergeFragmentedText(segments) {
            var firstText = -1, lastText = -1;
            for (var i = 0; i < segments.length; i++) {
                if (segments[i].kind === 'text') {
                    if (firstText === -1) firstText = i;
                    lastText = i;
                }
            }
            if (firstText === -1 || firstText === lastText) return segments;
            var combined = segments[firstText].value;
            for (var j = firstText + 1; j <= lastText; j++) {
                combined += segments[j].sep + segments[j].value;
            }
            var merged = { kind: 'text', value: combined, sep: segments[firstText].sep };
            var result = segments.slice(0, firstText);
            result.push(merged);
            for (var k = lastText + 1; k < segments.length; k++) result.push(segments[k]);
            return result;
        }

        /**
         * 既存のファイル名からセグメントの並び順を推定する
         * @param {Array<object>} segments セグメントの配列
         * @returns {Array<string>} セグメント種別の並び
         */
        function deriveOrderFromSegments(segments) {
            var kindToOrder = {
                base: 'base',
                text: 'title',
                status: 'status',
                date: 'timestamp',
                version: 'version'
            };
            var order = [];
            var seen = {};
            for (var i = 0; i < segments.length; i++) {
                var k = kindToOrder[segments[i].kind];
                if (!k || seen[k]) continue;
                if (k === 'status' && !FEATURE_STATUS) continue;
                order.push(k);
                seen[k] = true;
            }
            return order;
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
        function todayTimestamp(sep, withTime) {
            sep = sep || '';
            var d = new Date();
            var date = String(d.getFullYear()) + sep +
                padLeft(String(d.getMonth() + 1), 2) + sep +
                padLeft(String(d.getDate()), 2);
            if (!withTime) return date;
            return date + '-' + padLeft(String(d.getHours()), 2) + padLeft(String(d.getMinutes()), 2);
        }

        /**
         * バージョン番号を 1 つ繰り上げる
         * @param {string} versionText 現在のバージョン表記
         * @returns {string} 繰り上げ後の表記
         */
        function bumpVersionInPlace(currentBaseName) {
            var match = String(currentBaseName).match(/([vV])(\d+)/);
            if (match) {
                var letter = match[1];
                var digits = match[2];
                var nextNum = parseInt(digits, 10) + 1;
                var newDigits = padLeft(String(nextNum), digits.length);
                return currentBaseName.replace(/([vV])\d+/, letter + newDigits);
            }
            return currentBaseName + '-v2';
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
            var nextNum = match ? (parseInt(match[2], 10) + 1) : 1;
            if (mode === 'padded') {
                var width = match ? Math.max(match[2].length, 2) : 2;
                return letter + padLeft(String(nextNum), width);
            }
            if (mode === 'paddedWide') {
                var widthWide = match ? Math.max(match[2].length, 3) : 3;
                return letter + padLeft(String(nextNum), widthWide);
            }
            return letter + String(nextNum);
        }

        /**
         * バージョン表記を接頭辞と数値へ分解する
         * @param {string} versionText バージョン表記
         * @returns {object|null} 分解結果。該当しない場合は null
         */
        function extractVersionParts(baseName) {
            var m = String(baseName).match(/^(.*?)([vV])(\d+)(.*)$/);
            if (!m) return null;
            return { prefix: m[1], letter: m[2], digits: m[3], suffix: m[4] };
        }

        /**
         * 連番表記を接頭辞と数値へ分解する
         * @param {string} pageText 連番表記
         * @returns {object|null} 分解結果。該当しない場合は null
         */
        function extractPageParts(baseName, prefix) {
            var p = String(prefix || '');
            if (!p) return null;
            var re = new RegExp('^(.*?)(' + escapeRegExp(p) + ')(\\d+)(.*)$', 'i');
            var m = String(baseName).match(re);
            if (!m) return null;
            return { prefix: m[1], digits: m[3], suffix: m[4] };
        }

        /**
         * 正規表現で特別な意味を持つ文字をエスケープする
         * @param {string} text 対象の文字列
         * @returns {string} エスケープした文字列
         */
        function escapeRegExp(s) {
            return String(s).replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
        }

        /**
         * 同じフォルダ内で使われている最大のバージョン番号を探す
         * @param {Folder} folder 対象のフォルダ
         * @param {string} baseName 照合するベース名
         * @returns {number} 最大のバージョン番号
         */
        function findMaxVersionInFolder(baseName, folder, ext) {
            if (!folder) return null;
            var parts = extractVersionParts(baseName);
            if (!parts) return null;
            var re = new RegExp(
                '^' + escapeRegExp(parts.prefix) + '[vV](\\d+)'
                + escapeRegExp(parts.suffix) + escapeRegExp(ext) + '$',
                'i'
            );
            var files;
            try { files = folder.getFiles(); } catch (e) { return null; }
            var maxNum = -1;
            for (var i = 0; i < files.length; i++) {
                if (!(files[i] instanceof File)) continue;
                var fname = decodePercentEncoded(files[i].name);
                var fm = fname.match(re);
                if (!fm) continue;
                var num = parseInt(fm[1], 10);
                if (num > maxNum) maxNum = num;
            }
            return maxNum >= 0 ? maxNum : null;
        }

        /**
         * 重複しない次のバージョン付きファイル名を求める
         * @param {Folder} folder 対象のフォルダ
         * @param {string} fileName 元のファイル名
         * @returns {string} 次に使えるファイル名
         */
        function nextAvailableVersionName(baseName, folder, ext) {
            var maxNum = findMaxVersionInFolder(baseName, folder, ext);
            if (maxNum === null) return baseName;
            var parts = extractVersionParts(baseName);
            if (!parts) return baseName;
            var current = parseInt(parts.digits, 10);
            var target = maxNum + 1;
            if (target <= current) return baseName;
            var width = Math.max(parts.digits.length, String(target).length);
            return parts.prefix + parts.letter + padLeft(String(target), width) + parts.suffix;
        }

        /**
         * 同じフォルダ内で使われている最大の連番を探す
         * @param {Folder} folder 対象のフォルダ
         * @param {string} baseName 照合するベース名
         * @returns {number} 最大の連番
         */
        function findMaxPageInFolder(baseName, folder, ext, prefix) {
            if (!folder) return null;
            var p = String(prefix || '');
            if (!p) return null;
            var parts = extractPageParts(baseName, p);
            if (!parts) return null;
            var re = new RegExp(
                '^' + escapeRegExp(parts.prefix) + escapeRegExp(p) + '(\\d+)'
                + escapeRegExp(parts.suffix) + escapeRegExp(ext) + '$',
                'i'
            );
            var files;
            try { files = folder.getFiles(); } catch (e) { return null; }
            var maxNum = -1;
            for (var i = 0; i < files.length; i++) {
                if (!(files[i] instanceof File)) continue;
                var fname = decodePercentEncoded(files[i].name);
                var fm = fname.match(re);
                if (!fm) continue;
                var num = parseInt(fm[1], 10);
                if (num > maxNum) maxNum = num;
            }
            return maxNum >= 0 ? maxNum : null;
        }

        /**
         * 重複しない次の連番付きファイル名を求める
         * @param {Folder} folder 対象のフォルダ
         * @param {string} fileName 元のファイル名
         * @returns {string} 次に使えるファイル名
         */
        function nextAvailablePageName(baseName, folder, ext, prefix) {
            var p = String(prefix || '');
            if (!p) return baseName;
            var maxNum = findMaxPageInFolder(baseName, folder, ext, p);
            if (maxNum === null) return baseName;
            var parts = extractPageParts(baseName, p);
            if (!parts) return baseName;
            var current = parseInt(parts.digits, 10);
            var target = maxNum + 1;
            if (target <= current) return baseName;
            var width = Math.max(parts.digits.length, String(target).length);
            return parts.prefix + p + padLeft(String(target), width) + parts.suffix;
        }

        /* Windows 予約名（拡張子の有無を問わず使用不可）。一致したら末尾に "_" を足してエスケープ
           / Windows-reserved basenames (regardless of extension); append "_" to escape them */
        var WINDOWS_RESERVED_NAMES = {
            CON: 1, PRN: 1, AUX: 1, NUL: 1,
            COM1: 1, COM2: 1, COM3: 1, COM4: 1, COM5: 1,
            COM6: 1, COM7: 1, COM8: 1, COM9: 1,
            LPT1: 1, LPT2: 1, LPT3: 1, LPT4: 1, LPT5: 1,
            LPT6: 1, LPT7: 1, LPT8: 1, LPT9: 1
        };

        /**
         * Windows の予約語にあたる名前を安全な形へ変える
         * @param {string} fileName 対象のファイル名
         * @returns {string} 安全なファイル名
         */
        function escapeWindowsReserved(baseName) {
            var key = String(baseName).toUpperCase();
            return (WINDOWS_RESERVED_NAMES[key] === 1) ? baseName + '_' : baseName;
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
         * 連続する区切り記号をまとめ、前後の区切りを取り除く
         * @param {string} fileName 対象のファイル名
         * @returns {string} 整えたファイル名
         */
        function collapseAndTrimSeparators(str, sep) {
            var s = String(str);
            if (sep === '-' || sep === '_') {
                // 異種の混在も含めて連続した区切りを sep 1 つに統一（YYYY-MM-DD 等は呼び出し前に通過済み想定）
                s = s.replace(/[-_.]{2,}/g, sep);
            } else {
                // 区切り未指定でも同種連続だけは圧縮（-- → -, __ → _, .. → .）
                s = s.replace(/--+/g, '-').replace(/__+/g, '_').replace(/\.\.+/g, '.');
            }
            // 先頭末尾の区切り・空白を除去
            s = s.replace(/^[-_.\s]+|[-_.\s]+$/g, '');
            return s;
        }

        /**
         * UTF-8 でのバイト数を数える
         * @param {string} text 対象の文字列
         * @returns {number} バイト数
         */
        function byteLengthUTF8(str) {
            var s = String(str);
            var n = 0;
            for (var i = 0; i < s.length; i++) {
                var code = s.charCodeAt(i);
                if (code < 0x80) n += 1;
                else if (code < 0x800) n += 2;
                else if (code >= 0xD800 && code <= 0xDBFF) { n += 4; i++; } // サロゲートペア（4 バイト）
                else n += 3;
            }
            return n;
        }

        /**
         * ファイル名に使えない文字を置き換える
         * @param {string} fileName 元のファイル名
         * @returns {string} 安全なファイル名
         */
        function sanitizeFilename(str) {
            var trimmed = String(str).replace(/^\s+|\s+$/g, '');
            if (FEATURE_CLEAN) return trimmed;
            return trimmed.replace(/[\\\/:*?"<>|]/g, '');
        }

        /**
         * 文字列を NFC 正規化する
         * @param {string} text 対象の文字列
         * @returns {string} 正規化した文字列
         */
        function normalizeNFC(str) {
            return String(str).replace(/(.)([゙゚])/g, function (_, base, mark) {
                var voiced = (mark === '゙');
                if (voiced) {
                    if (base === 'う') return 'ゔ';
                    if (base === 'ウ') return 'ヴ';
                    if (base === 'ワ') return 'ヷ';
                    if (base === 'ヰ') return 'ヸ';
                    if (base === 'ヱ') return 'ヹ';
                    if (base === 'ヲ') return 'ヺ';
                }
                var code = base.charCodeAt(0);
                var isHiragana = code >= 0x3041 && code <= 0x3093;
                var isKatakana = code >= 0x30A1 && code <= 0x30F3;
                if (isHiragana || isKatakana) {
                    return String.fromCharCode(code + (voiced ? 1 : 2));
                }
                return base + mark;
            });
        }

        /**
         * ファイル名にそのまま使える文字かを判定する
         * @param {string} character 判定する文字
         * @returns {boolean} 使えるなら true
         */
        function isStandardFilenameChar(code) {
            // OS 禁止文字: \ 5C / 2F : 3A * 2A ? 3F " 22 < 3C > 3E | 7C
            if (code === 0x5C || code === 0x2F || code === 0x3A || code === 0x2A
                || code === 0x3F || code === 0x22 || code === 0x3C
                || code === 0x3E || code === 0x7C) return false;
            // 半角・全角スペース
            if (code === 0x20 || code === 0x3000) return false;
            if (code >= 0x21 && code <= 0x7E) return true;     // ASCII printable（space を除く）
            if (code >= 0x3001 && code <= 0x303F) return true; // CJK 記号と句読点（、。「」など / 全角スペースを除く）
            if (code >= 0x3040 && code <= 0x309F) return true; // ひらがな
            if (code >= 0x30A0 && code <= 0x30FF) return true; // カタカナ
            if (code >= 0x3400 && code <= 0x4DBF) return true; // CJK 拡張 A
            if (code >= 0x4E00 && code <= 0x9FFF) return true; // CJK 基本
            if (code >= 0xAC00 && code <= 0xD7A3) return true; // ハングル音節
            if (code >= 0xFF00 && code <= 0xFF5E) return true; // 全角 ASCII
            if (code >= 0xFF61 && code <= 0xFF9F) return true; // 半角カタカナ
            return false;
        }

        /**
         * ファイル名に不向きな文字を置き換える
         * @param {string} fileName 対象のファイル名
         * @returns {string} 整えたファイル名
         */
        function cleanFilenameChars(str, mode) {
            var replacement = (mode === 'dash') ? '-' : (mode === 'underscore') ? '_' : '';
            // 連続する空白を 1 つにまとめてから per-char で置換（mode='dash' で "  " が "--" にならないように）
            var s = String(str).replace(/\s+/g, ' ');
            var result = '';
            var i = 0;
            while (i < s.length) {
                var code = s.charCodeAt(i);
                if (code >= 0xD800 && code <= 0xDBFF && i + 1 < s.length) {
                    // 高サロゲート + 低サロゲートは常に非標準（BMP 外の絵文字など）
                    result += replacement;
                    i += 2;
                    continue;
                }
                if (isStandardFilenameChar(code)) {
                    result += s.charAt(i);
                } else {
                    result += replacement;
                }
                i++;
            }
            return result;
        }

        /* 半角カタカナ → 全角カタカナのマップ（清音）/ Half-width → full-width katakana (unvoiced) */
        var HALFWIDTH_KANA_MAP = {
            'ｦ': 'ヲ', 'ｧ': 'ァ', 'ｨ': 'ィ', 'ｩ': 'ゥ', 'ｪ': 'ェ', 'ｫ': 'ォ',
            'ｬ': 'ャ', 'ｭ': 'ュ', 'ｮ': 'ョ', 'ｯ': 'ッ', 'ｰ': 'ー',
            'ｱ': 'ア', 'ｲ': 'イ', 'ｳ': 'ウ', 'ｴ': 'エ', 'ｵ': 'オ',
            'ｶ': 'カ', 'ｷ': 'キ', 'ｸ': 'ク', 'ｹ': 'ケ', 'ｺ': 'コ',
            'ｻ': 'サ', 'ｼ': 'シ', 'ｽ': 'ス', 'ｾ': 'セ', 'ｿ': 'ソ',
            'ﾀ': 'タ', 'ﾁ': 'チ', 'ﾂ': 'ツ', 'ﾃ': 'テ', 'ﾄ': 'ト',
            'ﾅ': 'ナ', 'ﾆ': 'ニ', 'ﾇ': 'ヌ', 'ﾈ': 'ネ', 'ﾉ': 'ノ',
            'ﾊ': 'ハ', 'ﾋ': 'ヒ', 'ﾌ': 'フ', 'ﾍ': 'ヘ', 'ﾎ': 'ホ',
            'ﾏ': 'マ', 'ﾐ': 'ミ', 'ﾑ': 'ム', 'ﾒ': 'メ', 'ﾓ': 'モ',
            'ﾔ': 'ヤ', 'ﾕ': 'ユ', 'ﾖ': 'ヨ',
            'ﾗ': 'ラ', 'ﾘ': 'リ', 'ﾙ': 'ル', 'ﾚ': 'レ', 'ﾛ': 'ロ',
            'ﾜ': 'ワ', 'ﾝ': 'ン'
        };

        /* 半角カナ + ﾞ の結合マップ（濁音）/ Half-width + dakuten → voiced full-width */
        var HALFWIDTH_KANA_VOICED = {
            'ｶ': 'ガ', 'ｷ': 'ギ', 'ｸ': 'グ', 'ｹ': 'ゲ', 'ｺ': 'ゴ',
            'ｻ': 'ザ', 'ｼ': 'ジ', 'ｽ': 'ズ', 'ｾ': 'ゼ', 'ｿ': 'ゾ',
            'ﾀ': 'ダ', 'ﾁ': 'ヂ', 'ﾂ': 'ヅ', 'ﾃ': 'デ', 'ﾄ': 'ド',
            'ﾊ': 'バ', 'ﾋ': 'ビ', 'ﾌ': 'ブ', 'ﾍ': 'ベ', 'ﾎ': 'ボ',
            'ｳ': 'ヴ'
        };

        /* 半角カナ + ﾟ の結合マップ（半濁音）/ Half-width + handakuten → semi-voiced full-width */
        var HALFWIDTH_KANA_SEMI_VOICED = {
            'ﾊ': 'パ', 'ﾋ': 'ピ', 'ﾌ': 'プ', 'ﾍ': 'ペ', 'ﾎ': 'ポ'
        };

        /**
         * 半角カナを全角カナへ変換する
         * @param {string} text 対象の文字列
         * @returns {string} 変換した文字列
         */
        function convertHalfwidthKana(str) {
            var s = String(str);
            var result = '';
            var i = 0;
            while (i < s.length) {
                var ch = s.charAt(i);
                var next = (i + 1 < s.length) ? s.charAt(i + 1) : '';
                if (next === 'ﾞ' && typeof HALFWIDTH_KANA_VOICED[ch] === 'string') {
                    result += HALFWIDTH_KANA_VOICED[ch];
                    i += 2;
                    continue;
                }
                if (next === 'ﾟ' && typeof HALFWIDTH_KANA_SEMI_VOICED[ch] === 'string') {
                    result += HALFWIDTH_KANA_SEMI_VOICED[ch];
                    i += 2;
                    continue;
                }
                var mapped = HALFWIDTH_KANA_MAP[ch];
                if (typeof mapped === 'string') {
                    result += mapped;
                } else if (ch === 'ﾞ') {
                    result += '゛';
                } else if (ch === 'ﾟ') {
                    result += '゜';
                } else {
                    result += ch;
                }
                i++;
            }
            return result;
        }

        /**
         * 丸数字や法人略記などを ASCII 表記へ置き換える
         * @param {string} text 対象の文字列
         * @returns {string} 置き換えた文字列
         */
        function transliterate(str, mode) {
            if (mode === 'keep') return String(str);
            var s = String(str);
            var result = '';
            for (var i = 0; i < s.length; i++) {
                var ch = s.charAt(i);
                var mapped = TRANSLITERATE_MAP[ch];
                if (typeof mapped === 'string') {
                    result += (mode === 'remove') ? '' : mapped;
                } else {
                    result += ch;
                }
            }
            return result;
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
            var grandparentFolderName = '';
            if (fullName && fullName.parent) {
                parentFolderName = decodePercentEncoded(fullName.parent.name);
                if (fullName.parent.parent) {
                    grandparentFolderName = decodePercentEncoded(fullName.parent.parent.name);
                }
            }
            return {
                currentName: currentName,
                baseName: stripExtension(currentName),
                fsPath: fullName ? fullName.fsName : null,
                folder: fullName ? fullName.parent : null,
                parentFolderName: parentFolderName,
                grandparentFolderName: grandparentFolderName
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
         * ファイルをゴミ箱へ移動する
         * @param {File} targetFile 対象のファイル
         * @returns {boolean} 移動できたら true
         */
        function moveToTrash(file) {
            if (!file || !file.exists) return false;
            var trash = Folder("~/.Trash");
            if (!trash.exists) return false;
            var origName = file.name;
            var dot = origName.lastIndexOf('.');
            var stem = (dot > 0) ? origName.substring(0, dot) : origName;
            var ext = (dot > 0) ? origName.substring(dot) : '';
            var dest = File(trash.fsName + '/' + origName);
            var counter = 1;
            while (dest.exists) {
                dest = File(trash.fsName + '/' + stem + ' ' + counter + ext);
                counter++;
                if (counter > 9999) return false; // 暴走防止
            }
            try {
                if (!file.copy(dest)) return false;
                return file.remove();
            } catch (e) {
                return false;
            }
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
            if (FEATURE_USE_TRASH && moveToTrash(file)) return;
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
                var originalFile = File(originalFsPath);
                if (!originalFile.exists || !originalFile.copy(destFile)) {
                    throw new Error(getLabel('message.saveFailed') + '\n' + destFile.fsName);
                }
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
                    return sanitizeFilename(uiState.baseText || '');
                }
                if (kind === 'title') {
                    if (uiState.titleMode === 'none') return '';
                    if (uiState.titleMode === 'parent') return sanitizeFilename(uiState.parentFolderName);
                    if (uiState.titleMode === 'grandparent') return sanitizeFilename(uiState.grandparentFolderName);
                    return sanitizeFilename(uiState.titleText);
                }
                if (kind === 'status') {
                    return uiState.status || '';
                }
                if (kind === 'timestamp') {
                    var withTime = (uiState.timestampTime === 'hhmm');
                    if (uiState.timestamp === 'date') return todayTimestamp('', withTime);
                    if (uiState.timestamp === 'dateDash') return todayTimestamp('-', withTime);
                    return '';
                }
                if (kind === 'version') {
                    if (uiState.version === 'none') return '';
                    return formatVersion(getFirstSegmentValue(segments, 'version'), uiState.version);
                }
                if (kind === 'page') {
                    if (uiState.pageEnable !== 'yes') return '';
                    var pageWidth = (uiState.pagePad === '3') ? 3 : 2;
                    return (uiState.pagePrefix || '') + padLeft('1', pageWidth);
                }
                return '';
            }

            // ピースごとに区切り統一（YYYY-MM-DD のタイムスタンプだけは内部 "-" を保護）
            var order = (uiState.segmentOrder && uiState.segmentOrder.length) ? uiState.segmentOrder : SEGMENT_ORDER;
            var pieces = [];
            for (var i = 0; i < order.length; i++) {
                var kind = order[i];
                var value = valueForKind(kind);
                if (!value) continue;
                // YYYY-MM-DD（および YYYY-MM-DD-HHMM、YYYYMMDD-HHMM）は内部の "-" を保護
                var isProtectedDate = (kind === 'timestamp')
                    && (uiState.timestamp === 'dateDash' || uiState.timestampTime === 'hhmm');
                if (!isProtectedDate && FEATURE_SEPARATOR) {
                    // 区切り記号統一。FEATURE_DOT_NORMALIZE が true なら "." も対象
                    if (uiState.separator === '-') {
                        value = FEATURE_DOT_NORMALIZE ? value.replace(/[_.]/g, '-') : value.replace(/_/g, '-');
                    } else if (uiState.separator === '_') {
                        value = FEATURE_DOT_NORMALIZE ? value.replace(/[-.]/g, '_') : value.replace(/-/g, '_');
                    }
                }
                pieces.push(value);
            }
            var result = pieces.join(defaultSep);
            // スペース（半角・全角・タブ・改行）の処理は後段の cleanFilenameChars に集約
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
         * セグメントの並び替えパネルを組み立てる
         * @param {Window} dialog 対象のダイアログ
         * @param {object} prefs 保存しておいた設定
         * @returns {object} パネル内のコントロール
         */
        function buildSortPanel(parent, currentOrderAvailable) {
            var panel = parent.add('panel', undefined, getLabel('panel.sort'));
            setupPanel(panel);
            var sortOffRadio = panel.add('radiobutton', undefined, getLabel('radio.sortOff'));
            sortOffRadio.helpTip = getLabel('tip.sort');
            var sortCurrentRadio = panel.add('radiobutton', undefined, getLabel('radio.sortCurrent'));
            sortCurrentRadio.helpTip = getLabel('tip.sort');
            if (!currentOrderAvailable) sortCurrentRadio.enabled = false;
            // 「カスタム順」ラジオと「編集」ボタンを同じ行に並べる
            var customRow = panel.add('group');
            customRow.orientation = 'row';
            customRow.alignment = ['fill', 'top'];
            customRow.alignChildren = ['left', 'center'];
            customRow.spacing = 8;
            var sortOnRadio = customRow.add('radiobutton', undefined, getLabel('radio.sortOn'));
            sortOnRadio.helpTip = getLabel('tip.sort');
            var sortButton = customRow.add('button', undefined, getLabel('button.sort'));
            // 並び順の初期値は prefs を見ず、常に「現在のファイル名に準じる」（不可なら「標準順」）に固定
            var initialSort = currentOrderAvailable ? 'current' : 'off';
            sortOffRadio.value = (initialSort === 'off');
            sortCurrentRadio.value = (initialSort === 'current');
            sortOnRadio.value = (initialSort === 'on');
            sortButton.enabled = (initialSort === 'on');
            /**
             * 並び替えボタンの有効／無効を切り替える
             * @returns {void}
             */
            function syncSortButtonEnabled() {
                sortButton.enabled = sortOnRadio.value;
            }
            // onClick は呼び出し側で wire（refreshPreviews と組み合わせるため）
            return {
                panel: panel,
                sortOffRadio: sortOffRadio,
                sortCurrentRadio: sortCurrentRadio,
                sortOnRadio: sortOnRadio,
                sortButton: sortButton,
                /* 'off' / 'current' / 'on' */
                getSortMode: function () {
                    if (sortOnRadio.value) return 'on';
                    if (sortCurrentRadio.value) return 'current';
                    return 'off';
                },
                isSortOn: function () { return sortOnRadio.value; },
                syncSortButtonEnabled: syncSortButtonEnabled
            };
        }

        /**
         * 動作モード（リネーム／別名保存／コピー）のパネルを組み立てる
         * @param {Window} dialog 対象のダイアログ
         * @param {object} prefs 保存しておいた設定
         * @returns {object} パネル内のコントロール
         */
        function buildOpModePanel(parent) {
            var panel = parent.add('panel', undefined, getLabel('panel.opMode'));
            setupPanel(panel);
            var versionOnlyRadio = panel.add('radiobutton', undefined, getLabel('radio.opVersionOnly'));
            var fullRadio = panel.add('radiobutton', undefined, getLabel('radio.opFull'));
            versionOnlyRadio.value = false;
            fullRadio.value = true;
            return {
                panel: panel,
                versionOnlyRadio: versionOnlyRadio,
                fullRadio: fullRadio,
                isVersionOnly: function () { return versionOnlyRadio.value; },
                getOpMode: function () { return versionOnlyRadio.value ? 'versionOnly' : 'full'; }
            };
        }

        /**
         * 動作モードのパネルを組み立てる
         * @param {Window} dialog 対象のダイアログ
         * @param {object} prefs 保存しておいた設定
         * @returns {object} パネル内のコントロール
         */
        function buildModePanel(parent) {
            var panel = parent.add('panel', undefined, getLabel('panel.mode'));
            setupPanel(panel);
            var renameRadio = panel.add('radiobutton', undefined, getLabel('radio.rename'));
            renameRadio.helpTip = getLabel('tip.rename');
            var saveAsRadio = panel.add('radiobutton', undefined, getLabel('radio.saveAs'));
            saveAsRadio.helpTip = getLabel('tip.saveAs');
            var saveCopyRadio = panel.add('radiobutton', undefined, getLabel('radio.saveCopy'));
            saveCopyRadio.helpTip = getLabel('tip.saveCopy');
            // 初期選択は常に「別名で保存」
            renameRadio.value = false;
            saveAsRadio.value = true;
            saveCopyRadio.value = false;
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
            var currentNameLabel = currentNameRow.add('statictext', undefined, getLabelWithColon('label.currentName'), { justify: 'right' });
            currentNameRow.add('statictext', undefined, currentName);

            var finalNameRow = panel.add('group');
            finalNameRow.orientation = 'row';
            var finalNameLabel = finalNameRow.add('statictext', undefined, getLabelWithColon('label.finalName'), { justify: 'right' });
            // 「変更後：」は statictext のためレイアウト後にサイズ固定。
            // 現在のファイル名と「入力フィールド + 余白」の大きい方を確保しておく
            var finalNameValue = finalNameRow.add('statictext', undefined, currentName + '.indd');
            var currentNameWidth = panel.graphics.measureString(currentName + '.indd').width;
            finalNameValue.preferredSize.width = Math.max(currentNameWidth + 20, 340);

            // 個別整列はせず、ラベル参照を呼び出し側に返し、後段で全パネル統一整列する
            return {
                panel: panel,
                finalNameValue: finalNameValue,
                labels: [currentNameLabel, finalNameLabel],
                labelTexts: [getLabelWithColon('label.currentName'), getLabelWithColon('label.finalName')]
            };
        }

        /**
         * セグメントごとのオプションパネルを組み立てる
         * @param {object} parent 追加先のコンテナ
         * @param {Array<object>} segments セグメントの配列
         * @param {object} prefs 保存しておいた設定
         * @returns {object} パネル内のコントロール
         */
        function buildOptionsPanel(parent, segments, prefs, parentFolderName, grandparentFolderName) {
            var panel = parent.add('panel', undefined, getLabel('panel.options'));
            setupPanel(panel);

            // ベース: 検出値を初期表示する入力欄。空欄可、prefs には保存しない
            var detectedBase = getFirstSegmentValue(segments, 'base');
            var baseRow = panel.add('group');
            baseRow.orientation = 'row';
            var baseLabel = baseRow.add('statictext', undefined, getLabelWithColon('label.base'));
            baseLabel.helpTip = getLabel('tip.base');
            var baseField = baseRow.add('edittext', undefined, detectedBase);
            baseField.preferredSize.width = NEW_NAME_FIELD_WIDTH;
            baseField.helpTip = getLabel('tip.base');

            // 元ファイル名のテキスト部（base / status / date / version 以外）をサブテキストとして検出
            var detectedTitle = getFirstSegmentValue(segments, 'text');

            // サブテキスト: 1 行目 = ラベル + 3 ラジオ、2 行目 = 「指定」用の入力欄
            var titleSection = panel.add('group');
            titleSection.orientation = 'column';
            titleSection.alignChildren = ['fill', 'top'];
            titleSection.spacing = 4;

            var titleRow = addRadioRow(titleSection, 'label.title', 'tip.title', [
                { key: 'none', text: getLabel('radio.titleNone') },
                { key: 'custom', text: getLabel('radio.titleCustom') },
                { key: 'parent', text: getLabel('radio.titleParent') },
                { key: 'grandparent', text: getLabel('radio.titleGrandparent') }
            ]);
            titleRow.group.alignment = ['left', 'top'];
            // 親/2 階層上のフォルダー名がある場合は helpTip にフォルダ名を追記、無ければ無効化
            if (parentFolderName) titleRow.radios.parent.helpTip = getLabel('tip.title') + ' (' + parentFolderName + ')';
            else titleRow.radios.parent.enabled = false;
            if (grandparentFolderName) titleRow.radios.grandparent.helpTip = getLabel('tip.title') + ' (' + grandparentFolderName + ')';
            else titleRow.radios.grandparent.enabled = false;

            // 「指定」用の入力欄は次の行（ラベル列幅だけ左に余白を入れて radios に揃える）
            var titleFieldRow = titleSection.add('group');
            titleFieldRow.orientation = 'row';
            titleFieldRow.alignment = ['left', 'top'];
            var titleFieldSpacer = titleFieldRow.add('statictext', undefined, '');
            var titleField = titleFieldRow.add('edittext', undefined, detectedTitle);
            titleField.preferredSize.width = NEW_NAME_FIELD_WIDTH;
            titleField.helpTip = getLabel('tip.title');

            // 初期モード: 元ファイル名にサブテキストが無ければ「なし」を強制。
            // 検出できた場合は prefs.titleMode を優先、無ければ 'custom'（parent/grandparent は対応フォルダ名必須）
            var initialTitleMode = !detectedTitle ? 'none' :
                pickPref(prefs, 'titleMode', ['none', 'parent', 'grandparent', 'custom'], 'custom');
            if (initialTitleMode === 'parent' && !parentFolderName) {
                initialTitleMode = detectedTitle ? 'custom' : 'none';
            }
            if (initialTitleMode === 'grandparent' && !grandparentFolderName) {
                initialTitleMode = detectedTitle ? 'custom' : 'none';
            }
            titleRow.radios.none.value = (initialTitleMode === 'none');
            titleRow.radios.parent.value = (initialTitleMode === 'parent');
            titleRow.radios.grandparent.value = (initialTitleMode === 'grandparent');
            titleRow.radios.custom.value = (initialTitleMode === 'custom');
            titleField.enabled = (initialTitleMode === 'custom');

            /**
             * タイトルの指定方法に応じて入力欄の有効／無効を切り替える
             * @returns {void}
             */
            function syncTitleFieldEnabled() {
                titleField.enabled = titleRow.radios.custom.value;
            }

            // ステータス（dropdown。先頭は「なし」）
            var statusLabel = null, statusDropdown = null;
            if (FEATURE_STATUS) {
                var statusRow = panel.add('group');
                statusRow.orientation = 'row';
                statusLabel = statusRow.add('statictext', undefined, getLabelWithColon('label.status'));
                statusLabel.helpTip = getLabel('tip.status');
                statusDropdown = statusRow.add('dropdownlist', undefined, undefined);
                statusDropdown.helpTip = getLabel('tip.status');
                var parsedStatus = getFirstSegmentValue(segments, 'status');
                var initialStatus = parsedStatus || ((prefs && prefs.status) ? prefs.status : '');
                var initialStatusIdx = -1;
                for (var siItem = 0; siItem < STATUS_ITEMS.length; siItem++) {
                    var sItem = STATUS_ITEMS[siItem];
                    var ddItem = statusDropdown.add('item', (currentLanguage === 'ja' ? sItem.ja : sItem.en));
                    if (isStatusDivider(sItem)) {
                        ddItem.enabled = false;
                        continue;
                    }
                    if (initialStatusIdx === -1 && sItem.value === initialStatus) {
                        initialStatusIdx = siItem;
                    }
                }
                if (initialStatusIdx === -1) initialStatusIdx = 0;
                statusDropdown.selection = initialStatusIdx;
            }

            // タイムスタンプ（なし / YYYYMMDD / YYYY-MM-DD。デフォルト YYYYMMDD）
            // 末尾に「時刻も付与」(HHMM) チェックボックスを同居
            var timestampRow = addRadioRow(panel, 'label.timestamp', 'tip.timestamp', [
                { key: 'none', text: getLabel('radio.timestampNone') },
                { key: 'date', text: getLabel('radio.timestampDate') },
                { key: 'dateDash', text: getLabel('radio.timestampDateDash') }
            ]);
            var initialTimestamp = pickPref(prefs, 'timestamp', ['none', 'date', 'dateDash'], 'date');
            timestampRow.radios.none.value = (initialTimestamp === 'none');
            timestampRow.radios.date.value = (initialTimestamp === 'date');
            timestampRow.radios.dateDash.value = (initialTimestamp === 'dateDash');

            var timestampHHMMCheckbox = timestampRow.group.add('checkbox', undefined, getLabel('radio.timestampWithTime'));
            timestampHHMMCheckbox.helpTip = getLabel('tip.timestampWithTime');
            timestampHHMMCheckbox.value = pickPref(prefs, 'timestampTime', ['no', 'hhmm'], 'no') === 'hhmm';

            /**
             * タイムスタンプの書式に応じて時刻入力の有効／無効を切り替える
             * @returns {void}
             */
            function syncTimestampHHMMEnabled() {
                timestampHHMMCheckbox.enabled = !timestampRow.radios.none.value;
            }
            syncTimestampHHMMEnabled();

            // 連番（チェックボックス + プレフィックス入力 + 桁数ラジオ）。デフォルト OFF / "page" / 2 桁
            // SEGMENT_ORDER 内では timestamp の後・version の前に配置
            var pageRow = null;
            var pageCheckbox = null;
            var pagePrefixField = null;
            var pagePadRadio2 = null;
            var pagePadRadio3 = null;
            if (FEATURE_PAGE) {
                pageRow = panel.add('group');
                pageRow.orientation = 'row';
                var pageLabelCtrl = pageRow.add('statictext', undefined, getLabelWithColon('label.page'));
                pageLabelCtrl.helpTip = getLabel('tip.page');
                pageCheckbox = pageRow.add('checkbox', undefined, getLabel('radio.pageEnable'));
                pageCheckbox.helpTip = getLabel('tip.page');
                pageCheckbox.value = pickPref(prefs, 'pageEnable', ['no', 'yes'], 'no') === 'yes';
                pagePrefixField = pageRow.add('edittext', undefined, '');
                pagePrefixField.preferredSize.width = 60;
                pagePrefixField.helpTip = getLabel('tip.page');
                pagePrefixField.text = (prefs && typeof prefs.pagePrefix === 'string') ? prefs.pagePrefix : 'page';
                pagePadRadio2 = pageRow.add('radiobutton', undefined, getLabel('radio.pagePad2'));
                pagePadRadio2.helpTip = getLabel('tip.page');
                pagePadRadio3 = pageRow.add('radiobutton', undefined, getLabel('radio.pagePad3'));
                pagePadRadio3.helpTip = getLabel('tip.page');
                var initialPagePad = pickPref(prefs, 'pagePad', ['2', '3'], '2');
                pagePadRadio2.value = (initialPagePad === '2');
                pagePadRadio3.value = (initialPagePad === '3');
                pageRow.label = pageLabelCtrl; // alignLabelWidths 用に統一形にしておく
            }

            /**
             * 連番の ON/OFF に応じて関連項目を切り替える
             * @returns {void}
             */
            function syncPageControlsEnabled() {
                if (!pageCheckbox) return;
                var enabled = pageCheckbox.value;
                if (pagePrefixField) pagePrefixField.enabled = enabled;
                if (pagePadRadio2) pagePadRadio2.enabled = enabled;
                if (pagePadRadio3) pagePadRadio3.enabled = enabled;
            }
            syncPageControlsEnabled();

            // バージョン番号（なし / v1, v2… / v01, v02… / v001, v002…。デフォルト v1, v2…）
            // ES3 で 'short' は予約語のため、ラジオキーは short_ にする（pref 値 'short' とは別物）
            var versionRow = addRadioRow(panel, 'label.version', 'tip.version', [
                { key: 'none', text: getLabel('radio.versionNone') },
                { key: 'short_', text: getLabel('radio.versionShort') },
                { key: 'padded', text: getLabel('radio.versionPadded') },
                { key: 'paddedWide', text: getLabel('radio.versionPaddedWide') }
            ]);
            // デフォルトは「なし」。元ファイルに v 番号がある場合は prefs を優先（無ければ「なし」）
            var hasOriginalVersion = !!getFirstSegmentValue(segments, 'version');
            var initialVersion = hasOriginalVersion
                ? pickPref(prefs, 'version', ['none', 'short', 'padded', 'paddedWide'], 'none')
                : 'none';
            versionRow.radios.none.value = (initialVersion === 'none');
            versionRow.radios.short_.value = (initialVersion === 'short');
            versionRow.radios.padded.value = (initialVersion === 'padded');
            versionRow.radios.paddedWide.value = (initialVersion === 'paddedWide');

            // 区切り記号（変更しない / - / _ 横並び。デフォルト "-"）
            var separatorRow = null;
            if (FEATURE_SEPARATOR) {
                separatorRow = addRadioRow(panel, 'label.separator', 'tip.separator', [
                    { key: 'noChange', text: getLabel('radio.noChange') },
                    { key: 'dash', text: '-' },
                    { key: 'underscore', text: '_' }
                ]);
                // prefs があればそれを優先、無ければ FEATURE_SEPARATOR の指定文字（'-' or '_'）を初期値に
                var defaultSeparator = (FEATURE_SEPARATOR === '_') ? '_' : '-';
                var initialSeparator = pickPref(prefs, 'separator', ['', '-', '_'], defaultSeparator);
                separatorRow.radios.noChange.value = (initialSeparator === '');
                separatorRow.radios.dash.value = (initialSeparator === '-');
                separatorRow.radios.underscore.value = (initialSeparator === '_');
            }

            // 濁点処理（変更しない / 結合する。デフォルト "結合する"）
            var nfcRow = null;
            if (FEATURE_NFC) {
                nfcRow = addRadioRow(panel, 'label.nfc', 'tip.nfc', [
                    { key: 'keep', text: getLabel('radio.noChange') },
                    { key: 'combine', text: getLabel('radio.nfcCombine') }
                ]);
                var initialNfc = pickPref(prefs, 'nfc', ['keep', 'combine'], 'combine');
                nfcRow.radios.keep.value = (initialNfc === 'keep');
                nfcRow.radios.combine.value = (initialNfc === 'combine');
            }

            // クリーンなファイル名（削除する / -に変更 / _に変更。デフォルト "-に変更"）
            // OS 禁止文字（\/:*?"<>|）と絵文字・機種依存文字をまとめて処理。
            // 末尾に「半角カナ → 全角」チェックボックスを同居（'-' / '_' のときだけ有効）
            var cleanRow = null;
            var halfwidthKanaCheckbox = null;
            if (FEATURE_CLEAN) {
                cleanRow = addRadioRow(panel, 'label.clean', 'tip.clean', [
                    { key: 'remove', text: getLabel('radio.cleanRemove') },
                    { key: 'dash', text: getLabel('radio.changeToDash') },
                    { key: 'underscore', text: getLabel('radio.changeToUnderscore') }
                ]);
                var initialClean = pickPref(prefs, 'clean', ['remove', 'dash', 'underscore'], 'dash');
                cleanRow.radios.remove.value = (initialClean === 'remove');
                cleanRow.radios.dash.value = (initialClean === 'dash');
                cleanRow.radios.underscore.value = (initialClean === 'underscore');

                if (FEATURE_HALFWIDTH_KANA) {
                    halfwidthKanaCheckbox = cleanRow.group.add('checkbox', undefined, getLabel('radio.halfwidthKanaConvert'));
                    halfwidthKanaCheckbox.helpTip = getLabel('tip.halfwidthKana');
                    halfwidthKanaCheckbox.value = pickPref(prefs, 'halfwidthKana', ['keep', 'convert'], 'convert') === 'convert';
                }
            }

            // 法人略記・丸数字（変更しない / 削除する / 変換する。デフォルト "変換する"）
            var translitRow = null;
            if (FEATURE_TRANSLITERATE) {
                translitRow = addRadioRow(panel, 'label.translit', 'tip.translit', [
                    { key: 'keep', text: getLabel('radio.noChange') },
                    { key: 'remove', text: getLabel('radio.translitRemove') },
                    { key: 'convert', text: getLabel('radio.translitConvert') }
                ]);
                var initialTranslit = pickPref(prefs, 'translit', ['keep', 'remove', 'convert'], 'convert');
                translitRow.radios.keep.value = (initialTranslit === 'keep');
                translitRow.radios.remove.value = (initialTranslit === 'remove');
                translitRow.radios.convert.value = (initialTranslit === 'convert');
            }

            // クリーンが '-' / '_' のときだけ「半角カナ → 全角」を有効化
            /**
             * 半角カナ変換の ON/OFF に応じて関連項目を切り替える
             * @returns {void}
             */
            function syncHalfwidthKanaEnabled() {
                if (!halfwidthKanaCheckbox || !cleanRow) return;
                halfwidthKanaCheckbox.enabled = (cleanRow.radios.dash.value || cleanRow.radios.underscore.value);
            }
            syncHalfwidthKanaEnabled();

            // ラベル幅を統一（FEATURE で UI 非表示の行は除外）
            var labelTexts = [getLabelWithColon('label.base'), getLabelWithColon('label.title')];
            var labelControls = [baseLabel, titleRow.label];
            if (FEATURE_STATUS) { labelTexts.push(getLabelWithColon('label.status')); labelControls.push(statusLabel); }
            labelTexts.push(getLabelWithColon('label.timestamp')); labelControls.push(timestampRow.label);
            if (FEATURE_PAGE) { labelTexts.push(getLabelWithColon('label.page')); labelControls.push(pageRow.label); }
            labelTexts.push(getLabelWithColon('label.version')); labelControls.push(versionRow.label);
            if (FEATURE_SEPARATOR) { labelTexts.push(getLabelWithColon('label.separator')); labelControls.push(separatorRow.label); }
            if (FEATURE_NFC) { labelTexts.push(getLabelWithColon('label.nfc')); labelControls.push(nfcRow.label); }
            if (FEATURE_CLEAN) { labelTexts.push(getLabelWithColon('label.clean')); labelControls.push(cleanRow.label); }
            if (FEATURE_TRANSLITERATE) { labelTexts.push(getLabelWithColon('label.translit')); labelControls.push(translitRow.label); }
            // 個別整列はせず、呼び出し側に渡してファイル名パネルと統一整列する
            // titleFieldSpacer の幅は createDialog 側で整列後に設定する

            return {
                panel: panel,
                baseField: baseField,
                titleRow: titleRow,
                titleField: titleField,
                titleFieldSpacer: titleFieldSpacer,
                syncTitleFieldEnabled: syncTitleFieldEnabled,
                labels: labelControls,
                labelTexts: labelTexts,
                statusDropdown: statusDropdown,
                timestampRow: timestampRow,
                timestampHHMMCheckbox: timestampHHMMCheckbox,
                syncTimestampHHMMEnabled: syncTimestampHHMMEnabled,
                pageCheckbox: pageCheckbox,
                pagePrefixField: pagePrefixField,
                pagePadRadio2: pagePadRadio2,
                pagePadRadio3: pagePadRadio3,
                syncPageControlsEnabled: syncPageControlsEnabled,
                versionRow: versionRow,
                separatorRow: separatorRow,
                nfcRow: nfcRow,
                cleanRow: cleanRow,
                halfwidthKanaCheckbox: halfwidthKanaCheckbox,
                translitRow: translitRow,
                syncHalfwidthKanaEnabled: syncHalfwidthKanaEnabled,
                /* '' = 変更しない、'-' / '_' = 統一。FEATURE_SEPARATOR=false なら常に '' */
                getSeparator: function () {
                    if (!separatorRow) return '';
                    if (separatorRow.radios.noChange.value) return '';
                    if (separatorRow.radios.dash.value) return '-';
                    return '_';
                },
                /* 'keep' = そのまま、'combine' = 濁点・半濁点を NFC 結合。FEATURE_NFC=false なら常に 'keep' */
                getNfc: function () {
                    if (!nfcRow) return 'keep';
                    return nfcRow.radios.keep.value ? 'keep' : 'combine';
                },
                /* 'remove' / 'dash' / 'underscore'。FEATURE_CLEAN=false なら常に '' を返す（無処理） */
                getClean: function () {
                    if (!cleanRow) return '';
                    if (cleanRow.radios.dash.value) return 'dash';
                    if (cleanRow.radios.underscore.value) return 'underscore';
                    return 'remove';
                },
                /* 'keep' = そのまま、'remove' = 削除、'convert' = TRANSLITERATE_MAP で変換。FEATURE_TRANSLITERATE=false なら常に 'keep' */
                getTranslit: function () {
                    if (!translitRow) return 'keep';
                    if (translitRow.radios.convert.value) return 'convert';
                    if (translitRow.radios.remove.value) return 'remove';
                    return 'keep';
                },
                /* 'keep' = そのまま、'convert' = 半角カナを全角カナに。FEATURE_HALFWIDTH_KANA=false なら常に 'keep' */
                getHalfwidthKana: function () {
                    if (!halfwidthKanaCheckbox) return 'keep';
                    return halfwidthKanaCheckbox.value ? 'convert' : 'keep';
                },
                /* 'none' / 'date' / 'dateDash' */
                getTimestamp: function () {
                    if (timestampRow.radios.none.value) return 'none';
                    if (timestampRow.radios.dateDash.value) return 'dateDash';
                    return 'date';
                },
                /* 'no' / 'hhmm'（時刻 HHMM をタイムスタンプ末尾に付与するか）*/
                getTimestampTime: function () {
                    return timestampHHMMCheckbox.value ? 'hhmm' : 'no';
                },
                /* 'no' / 'yes'（ページ番号セグメントを付与するか）。FEATURE_PAGE=false なら常に 'no' */
                getPageEnable: function () {
                    if (!pageCheckbox) return 'no';
                    return pageCheckbox.value ? 'yes' : 'no';
                },
                /* '2' / '3'（連番のゼロ埋め桁数）。FEATURE_PAGE=false なら '2' */
                getPagePad: function () {
                    if (!pagePadRadio3) return '2';
                    return pagePadRadio3.value ? '3' : '2';
                },
                /* 連番のプレフィックス（例: "page"）。FEATURE_PAGE=false なら '' */
                getPagePrefix: function () {
                    if (!pagePrefixField) return '';
                    return pagePrefixField.text;
                },
                /* STATUS_ITEMS の value（'' = なし）。FEATURE_STATUS=false なら常に '' */
                getStatus: function () {
                    if (!FEATURE_STATUS || !statusDropdown) return '';
                    var idx = statusDropdown.selection ? statusDropdown.selection.index : 0;
                    var item = STATUS_ITEMS[idx];
                    if (!item || isStatusDivider(item)) return '';
                    return item.value;
                },
                /* 'none' / 'short' / 'padded' / 'paddedWide' */
                getVersion: function () {
                    if (versionRow.radios.none.value) return 'none';
                    if (versionRow.radios.short_.value) return 'short';
                    if (versionRow.radios.padded.value) return 'padded';
                    return 'paddedWide';
                },
                /* 'none' / 'parent' / 'grandparent' / 'custom' */
                getTitleMode: function () {
                    if (titleRow.radios.none.value) return 'none';
                    if (titleRow.radios.parent.value) return 'parent';
                    if (titleRow.radios.grandparent.value) return 'grandparent';
                    return 'custom';
                },
                /* ベース入力欄の現在値 / Current value of the base input */
                getBase: function () {
                    return baseField.text;
                }
            };
        }

        /**
         * セグメントの並び替えダイアログを開く
         * @returns {Array<string>|null} 並び順。キャンセル時は null
         */
        function openSortDialog(initialOrder) {
            var dlg = new Window('dialog', getLabel('sort.title'));
            dlg.opacity = DIALOG_OPACITY;
            setupWindow(dlg, 10);

            var hint = dlg.add('statictext', undefined, getLabel('sort.hint'));
            hint.alignment = 'left';

            var body = dlg.add('group');
            body.orientation = 'row';
            body.alignChildren = ['fill', 'fill'];

            var order = initialOrder.slice();
            var list = body.add('listbox', undefined, []);
            list.preferredSize = [200, 140];
            for (var i = 0; i < order.length; i++) {
                list.add('item', getLabel('label.' + order[i]));
            }
            list.selection = 0;

            var btnCol = body.add('group');
            btnCol.orientation = 'column';
            btnCol.alignChildren = 'fill';
            var upBtn = btnCol.add('button', undefined, '↑');
            upBtn.preferredSize = [36, 24];
            var downBtn = btnCol.add('button', undefined, '↓');
            downBtn.preferredSize = [36, 24];

            /**
             * 並び替えダイアログのリスト表示を更新する
             * @returns {void}
             */
            function refreshList(newIdx) {
                for (var i = 0; i < order.length; i++) {
                    list.items[i].text = getLabel('label.' + order[i]);
                }
                list.selection = newIdx;
            }
            upBtn.onClick = function () {
                var idx = list.selection ? list.selection.index : -1;
                if (idx <= 0) return;
                var tmp = order[idx - 1];
                order[idx - 1] = order[idx];
                order[idx] = tmp;
                refreshList(idx - 1);
            };
            downBtn.onClick = function () {
                var idx = list.selection ? list.selection.index : -1;
                if (idx < 0 || idx >= order.length - 1) return;
                var tmp = order[idx + 1];
                order[idx + 1] = order[idx];
                order[idx] = tmp;
                refreshList(idx + 1);
            };

            var buttons = dlg.add('group');
            buttons.alignment = ['center', 'top'];
            buttons.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
            buttons.add('button', undefined, 'OK', { name: 'ok' });

            if (dlg.show() !== 1) return null;
            return order;
        }

        /**
         * 配列に値が含まれるかを判定する
         * @param {Array} list 対象の配列
         * @param {*} value 探す値
         * @returns {boolean} 含まれていれば true
         */
        function arrayContains(arr, val) {
            for (var i = 0; i < arr.length; i++) {
                if (arr[i] === val) return true;
            }
            return false;
        }

        /**
         * 保存された設定から値を取り出す（無ければ既定値）
         * @param {object} prefs 保存しておいた設定
         * @param {string} key 設定のキー
         * @param {*} fallback 既定値
         * @returns {*} 設定の値
         */
        function pickPref(prefs, key, validValues, fallback) {
            var v = prefs && prefs[key];
            return arrayContains(validValues, v) ? v : fallback;
        }

        /**
         * 保存されたセグメント順の設定を配列へ変換する
         * @param {*} orderValue 保存されている値
         * @returns {Array<string>} セグメント種別の並び
         */
        function parseSegmentOrderPref(value) {
            if (!value) return null;
            var parts = String(value).split(',');
            if (parts.length !== SEGMENT_ORDER.length) return null;
            var seen = {};
            for (var i = 0; i < parts.length; i++) {
                var kind = parts[i];
                if (!arrayContains(SEGMENT_ORDER, kind) || seen[kind]) return null;
                seen[kind] = true;
            }
            return parts;
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
         * ラベル付きのラジオボタン行を追加する
         * @param {object} parent 追加先のコンテナ
         * @param {string} labelText ラベルの文字列
         * @param {Array<string>} optionLabels 選択肢の表示名
         * @returns {object} 追加したコントロール
         */
        function addRadioRow(panel, labelKey, tipKey, radioDefs) {
            var row = panel.add('group');
            row.orientation = 'row';
            var tip = L(tipKey);
            var label = row.add('statictext', undefined, getLabelWithColon(labelKey));
            label.helpTip = tip;
            var radios = {};
            for (var i = 0; i < radioDefs.length; i++) {
                var def = radioDefs[i];
                var r = row.add('radiobutton', undefined, def.text);
                r.helpTip = tip;
                radios[def.key] = r;
            }
            return { group: row, label: label, radios: radios };
        }

        /**
         * コントロールの変更時にプレビューを更新するよう結び付ける
         * @param {Array} controls 対象のコントロール
         * @returns {void}
         */
        function wireRefresh(callback, controls) {
            for (var i = 0; i < controls.length; i++) {
                var c = controls[i];
                if (!c) continue;
                var t = c.type;
                if (t === 'radiobutton' || t === 'checkbox') c.onClick = callback;
                else if (t === 'edittext') c.onChanging = callback;
                else if (t === 'dropdownlist') c.onChange = callback;
            }
        }

        /**
         * ファイル名変更のダイアログを組み立てる
         * @param {object} documentInfo ファイル情報
         * @param {Array<object>} segments セグメントの配列
         * @param {object} prefs 保存しておいた設定
         * @returns {object} ダイアログとコントロール
         */
        function createDialog(segments, currentName, prefs, parentFolderName, grandparentFolderName, folder) {
            var dialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
            dialog.opacity = DIALOG_OPACITY;
            setupWindow(dialog, 10);

            // 上部: 動作パネル + モードパネル + ソートパネル（FEATURE_SORT=false ならソートパネル無し）
            var topRow = dialog.add('group');
            topRow.orientation = 'row';
            topRow.alignChildren = ['fill', 'top'];
            var mode = buildModePanel(topRow, prefs);
            var opMode = buildOpModePanel(topRow);
            // 現在のファイル名から検出した並び順（要素ゼロなら「現在に準じる」は無効化）
            var currentOrder = FEATURE_SORT ? deriveOrderFromSegments(segments) : [];
            var sort = FEATURE_SORT ? buildSortPanel(topRow, prefs, currentOrder.length > 0) : null;

            // 並び順カスタム値（prefs に保存されたものを採用、不正・未保存ならデフォルト）
            var customOrder = FEATURE_SORT
                ? (parseSegmentOrderPref(prefs && prefs.segmentOrder) || SEGMENT_ORDER.slice())
                : SEGMENT_ORDER.slice();

            var filename = buildFilenamePanel(dialog, currentName);
            var options = buildOptionsPanel(dialog, segments, prefs, parentFolderName, grandparentFolderName);

            // 2 つのパネルのラベル幅を統一整列（測定基準は options.panel）
            var combinedTexts = filename.labelTexts.concat(options.labelTexts);
            var combinedLabels = filename.labels.concat(options.labels);
            alignLabelWidths(options.panel, combinedTexts, combinedLabels);
            // サブテキスト 2 行目「指定」入力欄の左余白を統一後のラベル列幅に合わせる
            options.titleFieldSpacer.preferredSize = [options.titleRow.label.preferredSize.width, 1];

            // ---- ライブプレビュー ----
            /**
             * ダイアログの現在の状態を取得する
             * @returns {object} UI の状態
             */
            function currentUIState() {
                var sortMode = (FEATURE_SORT && sort) ? sort.getSortMode() : 'off';
                var segmentOrder;
                if (sortMode === 'on') segmentOrder = customOrder;
                else if (sortMode === 'current') segmentOrder = currentOrder;
                else segmentOrder = SEGMENT_ORDER;
                return {
                    opMode: opMode.getOpMode(),
                    baseText: options.getBase(),
                    titleMode: options.getTitleMode(),
                    titleText: options.titleField.text,
                    parentFolderName: parentFolderName,
                    grandparentFolderName: grandparentFolderName,
                    status: options.getStatus(),
                    timestamp: options.getTimestamp(),
                    timestampTime: options.getTimestampTime(),
                    pageEnable: options.getPageEnable(),
                    pagePad: options.getPagePad(),
                    pagePrefix: options.getPagePrefix(),
                    version: options.getVersion(),
                    separator: options.getSeparator(),
                    nfc: options.getNfc(),
                    translit: options.getTranslit(),
                    clean: options.getClean(),
                    halfwidthKana: options.getHalfwidthKana(),
                    sort: sortMode,
                    customSegmentOrder: customOrder.slice(),
                    segmentOrder: segmentOrder
                };
            }

            // 「バージョン番号のみ」モード時に隠す UI（ソート + ファイル名の設定）
            /**
             * 動作モードに応じて関連コントロールの表示を切り替える
             * @returns {void}
             */
            function syncOpModeVisibility() {
                var versionOnly = opMode.isVersionOnly();
                if (sort) sort.panel.visible = !versionOnly;
                options.panel.visible = !versionOnly;
                dialog.layout.layout(true);
                dialog.layout.resize();
            }

            /**
             * 現在の入力内容でファイル名のプレビューを更新する
             * @returns {void}
             */
            function refreshPreviews() {
                options.syncTitleFieldEnabled();
                options.syncHalfwidthKanaEnabled();
                options.syncTimestampHHMMEnabled();
                options.syncPageControlsEnabled();
                // 「バージョンのみ」モードでは UI 整形を一切かけず、元ファイル名の v 番号だけ更新
                if (opMode.isVersionOnly()) {
                    var versionOnlyBase = bumpVersionInPlace(stripExtension(currentName));
                    versionOnlyBase = nextAvailableVersionName(versionOnlyBase, folder, '.indd');
                    filename.finalNameValue.text = versionOnlyBase + '.indd';
                    return;
                }
                var st = currentUIState();
                var finalBase = buildFinalName(segments, st);
                if (st.version === 'short' || st.version === 'padded' || st.version === 'paddedWide') {
                    finalBase = nextAvailableVersionName(finalBase, folder, '.indd');
                }
                if (st.pageEnable === 'yes') {
                    finalBase = nextAvailablePageName(finalBase, folder, '.indd', st.pagePrefix);
                }
                if (FEATURE_NFC && options.getNfc() === 'combine') {
                    finalBase = normalizeNFC(finalBase);
                }
                // 半角カナ → 全角カナ（clean が '-' / '_' のときだけ）。translit より先に行う
                if (FEATURE_HALFWIDTH_KANA && options.getHalfwidthKana() === 'convert') {
                    var cleanMode = options.getClean();
                    if (cleanMode === 'dash' || cleanMode === 'underscore') {
                        finalBase = convertHalfwidthKana(finalBase);
                    }
                }
                // translit はクリーンより先に行う（㈱→株 などを残すため）
                if (FEATURE_TRANSLITERATE) {
                    finalBase = transliterate(finalBase, options.getTranslit());
                }
                if (FEATURE_CLEAN) {
                    finalBase = cleanFilenameChars(finalBase, options.getClean());
                }
                finalBase = collapseAndTrimSeparators(finalBase, options.getSeparator());
                finalBase = escapeWindowsReserved(finalBase);
                filename.finalNameValue.text = finalBase + '.indd';
            }

            opMode.versionOnlyRadio.onClick = function () {
                syncOpModeVisibility();
                refreshPreviews();
            };
            opMode.fullRadio.onClick = function () {
                syncOpModeVisibility();
                refreshPreviews();
            };

            // ファイル名設定パネル内のすべての入力（無効化中の FEATURE は row が null で来るので skip される）
            var tr = options.titleRow, ts = options.timestampRow, vr = options.versionRow;
            var sr = options.separatorRow, nr = options.nfcRow, cr = options.cleanRow, lr = options.translitRow;

            // 構成要素（ベース／サブテキスト／ステータス／タイムスタンプ／バージョン）を変更したら、
            // 「現在のファイル名に準じる」は前提が崩れるので「標準順」に降格させる
            /**
             * 並び順の指定を既定へ戻す
             * @returns {void}
             */
            function demoteSortToDefault() {
                if (!sort) return;
                if (!sort.sortCurrentRadio.value) return;
                sort.sortCurrentRadio.value = false;
                sort.sortOffRadio.value = true;
                sort.syncSortButtonEnabled();
            }
            /**
             * プレビューを更新し、必要なら並び順を既定へ戻す
             * @returns {void}
             */
            function refreshAndDemoteSort() {
                demoteSortToDefault();
                refreshPreviews();
            }

            // 構成要素を変える操作（「現在のファイル名に準じる」は解除）
            wireRefresh(refreshAndDemoteSort, [
                tr.radios.none, tr.radios.parent, tr.radios.grandparent, tr.radios.custom,
                options.baseField, options.titleField,
                options.statusDropdown,
                ts.radios.none, ts.radios.date, ts.radios.dateDash,
                options.timestampHHMMCheckbox,
                options.pageCheckbox, options.pagePrefixField, options.pagePadRadio2, options.pagePadRadio3,
                vr.radios.none, vr.radios.short_, vr.radios.padded, vr.radios.paddedWide
            ]);
            // 整形のみ変える操作（並び順には影響しないので「現在のファイル名に準じる」を維持）
            wireRefresh(refreshPreviews, [
                sr && sr.radios.noChange, sr && sr.radios.dash, sr && sr.radios.underscore,
                nr && nr.radios.keep, nr && nr.radios.combine,
                cr && cr.radios.remove, cr && cr.radios.dash, cr && cr.radios.underscore,
                options.halfwidthKanaCheckbox,
                lr && lr.radios.keep, lr && lr.radios.remove, lr && lr.radios.convert
            ]);

            // ソートパネルの ON/OFF とサブダイアログ起動（FEATURE_SORT のとき）
            if (FEATURE_SORT && sort) {
                var onSortToggle = function () {
                    sort.syncSortButtonEnabled();
                    refreshPreviews();
                };
                sort.sortOffRadio.onClick = onSortToggle;
                sort.sortCurrentRadio.onClick = onSortToggle;
                sort.sortOnRadio.onClick = onSortToggle;
                sort.sortButton.onClick = function () {
                    var newOrder = openSortDialog(customOrder);
                    if (newOrder) {
                        customOrder = newOrder;
                        refreshPreviews();
                    }
                };
            }

            refreshPreviews();

            // ---- ボタン（右寄せ Cancel / OK） ----
            var buttonGroup = dialog.add('group');
            buttonGroup.alignment = ['right', 'top'];
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
            var segments = mergeFragmentedText(parseFileName(info.baseName));
            var prefs = loadPrefs();

            var ui = createDialog(segments, info.currentName, prefs, info.parentFolderName, info.grandparentFolderName, info.folder);
            if (ui.dialog.show() !== 1) return; // キャンセル

            var uiState = ui.getUIState();
            var newBaseName = (uiState.opMode === 'versionOnly')
                ? bumpVersionInPlace(info.baseName)
                : buildFinalName(segments, uiState);
            if (!newBaseName) {
                alert(getLabel('message.emptyName'));
                return;
            }

            var targetFolder = ensureTargetFolder(info.folder);
            if (!targetFolder) return; // キャンセル

            if (uiState.opMode === 'versionOnly' || uiState.version === 'short' || uiState.version === 'padded' || uiState.version === 'paddedWide') {
                newBaseName = nextAvailableVersionName(newBaseName, targetFolder, '.indd');
            }
            if (uiState.opMode !== 'versionOnly' && uiState.pageEnable === 'yes') {
                newBaseName = nextAvailablePageName(newBaseName, targetFolder, '.indd', uiState.pagePrefix);
            }

            // 「バージョンのみ」モードでは UI 整形をスキップして元ファイル名をそのまま尊重
            if (uiState.opMode !== 'versionOnly') {
                if (FEATURE_NFC && uiState.nfc === 'combine') {
                    newBaseName = normalizeNFC(newBaseName);
                }
                if (FEATURE_HALFWIDTH_KANA && uiState.halfwidthKana === 'convert'
                    && (uiState.clean === 'dash' || uiState.clean === 'underscore')) {
                    newBaseName = convertHalfwidthKana(newBaseName);
                }
                if (FEATURE_TRANSLITERATE) {
                    newBaseName = transliterate(newBaseName, uiState.translit);
                }
                if (FEATURE_CLEAN) {
                    newBaseName = cleanFilenameChars(newBaseName, uiState.clean);
                }
                newBaseName = collapseAndTrimSeparators(newBaseName, uiState.separator);
                newBaseName = escapeWindowsReserved(newBaseName);
            }

            var destFile = File(targetFolder.fsName + '/' + newBaseName + '.indd');

            // 長さチェック（拡張子込み）: 上限超過なら確認ダイアログを出して続行可
            var fullByteLength = byteLengthUTF8(newBaseName + '.indd');
            if (fullByteLength > FEATURE_MAX_FILENAME_BYTES) {
                var msg = getLabel('message.confirmTooLong')
                    .replace('{bytes}', String(fullByteLength))
                    .replace('{limit}', String(FEATURE_MAX_FILENAME_BYTES));
                if (!confirm(msg + '\n\n' + newBaseName + '.indd')) return;
            }

            if (!confirmOverwriteIfExists(destFile, info.fsPath)) return;

            try {
                executeOutput(doc, destFile, ui.getMode(), info.fsPath);
                // 成功したら今回の選択をプリセットとして保存（versionOnly モードでは保存しない）
                if (uiState.opMode !== 'versionOnly') {
                    var prefsToSave = {
                        titleMode: uiState.titleMode,
                        timestamp: uiState.timestamp,
                        timestampTime: uiState.timestampTime,
                        version: uiState.version
                    };
                    if (FEATURE_PAGE) {
                        prefsToSave.pageEnable = uiState.pageEnable;
                        prefsToSave.pagePad = uiState.pagePad;
                        prefsToSave.pagePrefix = uiState.pagePrefix;
                    }
                    if (FEATURE_STATUS) prefsToSave.status = uiState.status;
                    if (FEATURE_SEPARATOR) prefsToSave.separator = uiState.separator;
                    if (FEATURE_NFC) prefsToSave.nfc = uiState.nfc;
                    if (FEATURE_CLEAN) prefsToSave.clean = uiState.clean;
                    if (FEATURE_HALFWIDTH_KANA && FEATURE_CLEAN) prefsToSave.halfwidthKana = uiState.halfwidthKana;
                    if (FEATURE_TRANSLITERATE) prefsToSave.translit = uiState.translit;
                    if (FEATURE_SORT) {
                        // sort モード自体は復元しない（毎回「現在のファイル名に準じる」を初期値）
                        prefsToSave.segmentOrder = uiState.customSegmentOrder.join(',');
                    }
                    savePrefs(prefsToSave);
                }
            } catch (e) {
                alert(getLabel('message.saveFailed') + '\n' + e);
            }
        }

        main();

    })();