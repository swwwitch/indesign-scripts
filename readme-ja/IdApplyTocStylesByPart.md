# 目次の項目を部ごとの派生スタイルに当て直す

[![Direct](https://img.shields.io/badge/Direct%20Link-IdApplyTocStylesByPart.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdApplyTocStylesByPart.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdApplyTocStylesByPart.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

目次を更新したあとに実行し、目次の項目を部ごとの派生スタイル（「toc-h1.blue」など）に当て直します。本文の段落スタイルは変えません。目次を更新するたびに実行します。

本文の見出しを派生スタイルに分けて、［目次の更新］だけで済むようにしたいときは、[IdSplitHeadingStylesByPart](IdSplitHeadingStylesByPart.md) を使います。

### 主な機能

- **部の判定**：本文で選んだ段落スタイル（初期値は目次スタイルで最上位レベルの段落スタイル、例「h0」）の段落から、次の同じスタイルの段落の手前までを1つの部とみなします
- **目次との対応**：目次の部の見出し（例「toc-h0」の段落）を本文の部と文字列で照合し、次の部の見出しまでの項目をその部の派生スタイルにします。文字列が一致しないときは順番で対応させ、結果に表示します
- **部ごとの派生スタイル**：部ごとに接尾辞（blue・pink など）を選びます。初期値は、部の中ですでに使われている派生スタイル（「h2.blue」など）の多数決です
- 目次の書式「toc-h1」の段落は、部の接尾辞を付けた「toc-h1.blue」にします。そのスタイルが無いときは元の書式のままにし、結果に表示します
- **目次の強制改行を削除**：見出しの中の強制改行を目次から削除します。目次スタイルの［強制改行を削除］がオンでも目次に残るため、その代わりです
- 実行は［取り消し］1回で元に戻せます

### 使い方

1. 目次の書式の段落スタイルに、部ごとの派生スタイルを作っておく（例「toc-h1」をもとにした「toc-h1.blue」）
2. 目次を更新する
3. スクリプトを実行し、部ごとの接尾辞を確かめて［OK］

### 制限事項・メモ

- 目次は、目次スタイルの書式の段落スタイルを使ったストーリーから探します。
- 最初の部の見出しより前の項目（「はじめに」など）は元の書式にします。
- 段落スタイルのグループに入ったスタイルは、名前だけで照合します。
- 区切りの「.」は、スクリプト冒頭の `VARIANT_SEPARATOR` で変更できます。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdApplyTocStylesByPart.jsx` |
| バージョン | v1.0.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-10-30 |
| 最終更新 | 2026-11-05 |

### 更新履歴

- v1.0.1（2026-11-05）バージョンをタイトルバーから外し、ダイアログボックスの左下に小さく薄く表示（共通部品の追加）
- v1.0.0（2026-10-30）：初版

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
