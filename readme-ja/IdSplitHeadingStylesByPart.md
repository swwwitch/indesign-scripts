# 部ごとに見出しの派生スタイルを当てて、目次の書式を分ける

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSplitHeadingStylesByPart.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdSplitHeadingStylesByPart.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSplitHeadingStylesByPart.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

InDesign の目次は段落スタイル単位で項目を拾うため、同じ「h1」は部が変わっても同じ書式になります。このスクリプトは、目次に載せる見出しを部ごとの派生スタイル（「h1.blue」など）に振り分け、目次スタイルにも派生スタイルごとの書式（「toc-h1.blue」など）を登録します。以後は［目次の更新］だけで、部ごとに異なる書式の目次ができます。

目次を作ったあとに書式だけを当て直したいときは、[IdApplyTocStylesByPart](IdApplyTocStylesByPart.md) を使います。

### 主な機能

- **部の判定**：選んだ段落スタイル（初期値は目次スタイルで最上位レベルの段落スタイル、例「h0」）の段落から、次の同じスタイルの段落の手前までを1つの部とみなします。ページの並び順で判定します
- **派生スタイルの対応**：本文の「h1.blue」は、目次スタイルで「h1」に設定した書式（例「toc-h1」）に同じ接尾辞を付けた「toc-h1.blue」へ対応させます
- **部ごとの派生スタイル**：部ごとに接尾辞（blue・pink など）を選びます。初期値は、部の中ですでに使われている派生スタイル（「h2.blue」など）の多数決です。選べる接尾辞は、目次の書式の段落スタイルに「.接尾辞」を付けたスタイルから集めます
- **分ける見出し**：目次スタイルに登録された見出しのうち、振り分けるものを選びます
- 本文に派生スタイルが無ければ、元のスタイルを基準にして作ります（見た目は変わりません）。目次の書式の派生スタイルが無い場合も、目次の書式を基準にして作ります
- 目次スタイルに派生スタイルの項目を追加します。レベル・ノンブルの位置・区切り文字などは元の項目の設定を写します
- **目次を更新**：既存の目次を、選んだ目次スタイルで作り直します
- **目次の強制改行を削除**：見出しの中の強制改行を目次から削除します。目次スタイルの［強制改行を削除］がオンでも目次に残るため、その代わりです
- 実行は［取り消し］1回で元に戻せます

### 使い方

1. 目次の書式の段落スタイルに、部ごとの派生スタイルを作っておく（例「toc-h0」をもとにした「toc-h0.blue」）。少なくとも1つあれば、ほかのレベルの派生スタイルは自動で作ります
2. スクリプトを実行し、目次スタイル・部の段落スタイルを選ぶ
3. 部ごとに接尾辞を選び、［OK］

### 制限事項・メモ

- 部の外（最初の部より前）の見出しは変更しません。部の接尾辞を［なし］にすると、その部の見出しは元のスタイルに戻します。
- 接尾辞の候補に無い派生スタイル（例「h0.lead」）の段落は変更しません。
- 段落スタイルのグループに入ったスタイルは、名前だけで照合します。同じ名前のスタイルが複数のグループにあると、先に見つかったものを使います。
- 親ページ・ペーストボード上・オーバーセットの段落は対象外です。表のセル内の見出しも対象外です。
- 区切りの「.」は、スクリプト冒頭の `VARIANT_SEPARATOR` で変更できます。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdSplitHeadingStylesByPart.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-10-30 |
| 最終更新 | 2026-10-30 |

### 更新履歴

- v1.0.0（2026-10-30）：初版

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
