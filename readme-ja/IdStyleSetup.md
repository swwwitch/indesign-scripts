# 段落スタイル・文字スタイルを一括登録

[![Direct](https://img.shields.io/badge/Direct%20Link-IdStyleSetup.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdStyleSetup.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdStyleSetup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

段落スタイル・文字スタイルとそのグループ、継承関係、正規表現スタイルまでを一括で登録します。フォントは `base-font`・`base-heading`・`base-text` で一括変更でき、スタイル名は HTML と Word 対応から選べます。

### 主な機能

- 実行時のダイアログでスタイル名の体系を選択（HTML／Word 対応）
- スタイルとスタイルグループの作成、属性の適用、正規表現スタイルの設定、並び替えを 4 段階で実行
- basedOn による継承関係を先に確定させてから属性を適用し、共通の組版設定は基準スタイルへ集約
- フォントは基準スタイルで一括変更：`base-font` はすべて、`base-heading` は見出しと目次タイトル、`base-text` は本文・リスト・表セル・目次項目に反映
- 正規表現スタイルは用途別に設定：太字は見出し・表の見出しセル、欧文は見出し・p・リスト、行末分離禁止は p・リスト、アンカーは p・リスト・表セル、ラベルは ul-li のみ
- 表セルは `base-table`（水平・垂直比率 92%）を基準に、本文セル `td-left`（その下に td-justify / td-justify-all / td-center / td-right / 箇条書きの td-left ul-li）と見出しセル `th-left`（その下に th-center / 紙色文字の th-center-W）に分けて登録
- 各正規表現の先頭に `(?#欧文)` などの説明を付けて登録
- 処理中はプログレスバーを表示

### スタイルの構成

フォントはどのスタイルにも設定しません。`base-font`・`base-heading`・`base-text` のどれかを変えると、その下のスタイルにまとめて反映されます。

```
［基本段落］
└─ base-font               すべての段落スタイルの親。ここでフォントを変えると全体に反映
   ├─ base-heading         見出し系の親。メトリクス、左揃え、次の段落と連動（2行）、行はすべて分離禁止、ハイフネーションなし
   │  │                    正規表現：太字、欧文
   │  ├─ h1〜h6            次のスタイル：p
   │  └─ toc-title         目次タイトル
   ├─ base-text            本文系の親。和文等幅、均等配置（最終行左揃え）、行はすべて分離禁止、ハイフネーションなし
   │  ├─ p                 分離禁止をすべてオフ
   │  │  │                 正規表現：欧文、行末分離禁止、アンカー
   │  │  └─ p.table        表組み用（p の正規表現を継承）
   │  ├─ p.caption         前の段落と連動、次のスタイル：p
   │  ├─ p.code            言語なし、合字オフ、左揃え
   │  ├─ p.img             中央揃え
   │  ├─ ul-li             箇条書き（記号は li-bullet）、前の段落と連動、同じスタイル間のスペース 0、タブ位置：文字サイズ
   │  │                    正規表現：欧文、行末分離禁止、アンカー、ラベル
   │  ├─ ol-li             番号付き（番号は li-num）
   │  │                    正規表現：欧文、行末分離禁止、アンカー
   │  ├─ base-table        表全体の共通設定。水平・垂直比率 92%
   │  │  │                 正規表現：アンカー
   │  │  ├─ td-left        本文セルの親。左揃え
   │  │  │  ├─ td-justify      均等配置（最終行左揃え）
   │  │  │  ├─ td-justify-all  両端揃え
   │  │  │  ├─ td-center       中央揃え
   │  │  │  ├─ td-right        右揃え
   │  │  │  └─ td-left ul-li   表セル内の箇条書き（記号は li-bullet）、前の段落と連動、同じスタイル間のスペース 0、タブ位置：文字サイズ
   │  │  └─ th-left        見出しセルの親。左揃え
   │  │     │              正規表現：太字、アンカー
   │  │     ├─ th-center       中央揃え
   │  │     └─ th-center-W     中央揃え、文字色：紙色
   │  └─ base-toc          目次項目の親。左揃え、次の段落と連動（2行）、行はすべて分離禁止、ハイフネーションなし
   │     ├─ toc-h1
   │     ├─ toc-h2
   │     └─ toc-h3         次の段落と連動しない
   ├─ page-number          ノンブル。小口揃え
   ├─ running-head         柱
   └─ thumb-index          ツメ
```

### スタイル名の対応

| HTML | Word 対応 |
| --- | --- |
| `h1`〜`h6` | `Heading 1`〜`Heading 6` |
| `p` | `Normal` |
| `ul-li` | `List Paragraph` |
| `ol-li` | `List Number` |
| `strong-bold` | `Strong` |
| `em-italic` | `Emphasis` |

［Word 対応］を選ぶと、Word ファイルを「テキストと表のスタイルおよびフォーマットを保持」で配置したときにスタイルのマッピングが不要になります。上の表に無いスタイル（`p.caption`、`p.code`、`link` や、グループ内の `base-font`、`td-left`、`toc-h1` など）は、どちらを選んでも名前は変わりません。

### 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する
3. ダイアログでスタイル名の体系を選び、［OK］をクリック

### 制限事項・メモ

- 既定では既存の同名スタイルには触れません。`OVERWRITE_EXISTING_STYLES` を true にすると全属性を再適用し、GREP を付け直します。Word 原稿を配置済みのドキュメントで［Word 対応］を実行した場合、`Heading 1` などは既存スタイルとして扱われます。v1.4.3 以前で登録したドキュメントに実行すると、基準スタイルは新しい名前で別に作られ、正規表現スタイルも重複します。置き換えを有効にして実行してください。
- 箇条書き（`ul-li`）と番号付きリスト（`ol-li`）は Word ではどちらも「リスト段落」ですが、InDesign では同名のスタイルを2つ作れないため、番号付きリストは `List Number` に割り当てています。対応表はスクリプト冒頭の `WORD_STYLE_NAMES` で変更できます。
- フォント・太さなどスクリプトが扱わない属性は変更しません。色は `th-center-W` の文字色（紙色）だけを設定します。
- 全処理が 1 回の取り消しにまとまります。
- カーニング方式は UI 言語で表示名が変わるため、候補を順に試します。どれも該当しない環境では設定を据え置きます。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdStyleSetup.jsx` |
| バージョン | v1.5.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-05-03 |
| 最終更新 | 2026-10-01 |
| 紹介記事 | https://note.com/dtp_tranist/n/nfe87ec253780 |

### 更新履歴

#### v1.5.1（2026-10-01）

- 表セル内の箇条書き用に `td-left ul-li`（`td-left` を継承）を追加
- 文字スタイル `td-bold` の継承元を `strong-bold` に設定
- `base-heading` と `th-left` に正規表現スタイル `(?#太字)(.+)`（strong-bold）を追加
- `ul-li` と `td-left ul-li` のタブ位置を文字サイズ（1字分）に設定

#### v1.5.0（2026-10-01）

- 継承関係を再編。`base-font` をすべての親にし、見出し系は `base-heading`、本文系（p.code・p.img・表セル・目次項目を含む）は `base-text` を経由するように変更。基準スタイルの名前を変更：`base-regex` → `base-font`、`body-text` → `base-text`、`heading` → `base-heading`
- 正規表現スタイルを基準スタイルから外し、欧文は見出し・p・リスト、行末分離禁止は p・リスト、アンカーは p・リスト・表セルに設定
- 表セルのスタイルを再編。`base-table`（水平・垂直比率 92%）の下に `td-left`・`th-left` を置き、`td-justify`（均等配置（最終行左揃え））・`td-justify-all`（両端揃え）・`td-center`・`td-right`、`th-center`・`th-center-W`（文字色：紙色）を追加。`th`・`td`・`th-right` は廃止
- 各正規表現の先頭に `(?#欧文)` などの説明を付けるように変更
- ノンブル用の段落スタイル `page-number` の行揃えを小口揃えに設定

#### v1.4.3（2026-10-01）

- ダイアログの不透明度を 0.98 に統一し、進捗パレットの作り方をほかのスクリプトと同じ形に整理

#### v1.4.2（2026-10-01）

- 右側のボタンだけの行は、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更

#### v1.4.1（2026-09-30）

- ボタン行・余白・ローカライズの処理を共通部品に統一。ボタンを中央に配置

#### v1.4.0（2026-09-19）

- 実行時にダイアログを表示し、スタイル名の体系を［HTML］と［Word 対応］から選べるように変更
- ［Word 対応］では `Heading 1`〜`Heading 6`、`Normal`、`List Paragraph`、`List Number`、`Strong`、`Emphasis` の名前で登録。Word 原稿を配置するときのスタイルのマッピングが不要になります
- 対応表はスクリプト冒頭の `WORD_STYLE_NAMES` で変更可能

#### v1.3.3（2026-09-01）

- 表組み用の段落スタイル `p.table` を追加。`p` を基準スタイル（basedOn）に設定

#### v1.3.2（2026-09-01）

- `p` の分離禁止オプションを OFF にする設定などが効かないことがある問題を修正。属性を設定してから basedOn を張っていたため、親（body-text）側の設定を継承し直して打ち消される可能性がありました。継承関係を先に確定させてから属性を適用するように変更
- `basestyle` グループに `body-text` がないと、`heading` と h1〜h6 の継承関係まで設定されなくなる問題を修正
- 日本語版以外の InDesign でカーニング方式の設定に失敗し、スクリプト全体が取り消される問題を修正。表示名の候補（`和文等幅` / `Japanese Mojikumi` など）を順に試すように変更
- 未使用の UI ヘルパーを削除し、JSDoc の戻り値型とコメントの記述を実装に合わせて整理

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
