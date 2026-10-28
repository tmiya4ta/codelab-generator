# codelab-generator

Markdown から **Google Codelab 形式の DOCX と HTML** を生成する CLI ツール。
ClojureScript (shadow-cljs) で実装され、Node.js 上で動作します。

同じ Markdown から、DOCX (配布・印刷用) と、ブラウザで読む Codelab の HTML の両方を作れます。
HTML は [claat](https://github.com/googlecodelabs/tools) が出すものと同じ形 (`google-codelab` の Web Components) で、claat は要りません。

## 特徴

- Markdown 1 ファイルから Codelab 形式の DOCX を生成
- 同じ Markdown から Codelab の HTML を生成 (`-f html`)。日本語で読みやすい見た目 (本文の幅・字・表・画像・コード) にしてある
- FrontMatter (YAML) でメタデータ（`id` / `summary` / `categories` など）を指定
- ステップ分割・所要時間・インフォボックス・コンソール表示・テーブル・チャート・画像に対応
- 画像は [sharp](https://sharp.pixelplumbing.com/) で余白を自動トリム
- ` ```chart ` ブロックから [chart.js](https://www.chartjs.org/) で図を生成

## インストール / ビルド

```bash
npm install
npm run release   # :advanced 最適化ビルド（単発なら npm run build）
```

ビルド成果物は `dist/codelab.js` に出力されます。

## 使い方

```bash
node dist/codelab.js <input.md> [-f docx|html] [-o <output>]

# npm link 後は codelab-gen コマンドで実行できます
codelab-gen my-codelab.md -o my-codelab.docx
codelab-gen my-codelab.md -f html -o out/my-codelab
```

| オプション | 説明 |
|--|--|
| `-f`, `--format` | 出力の形式。`docx` (既定) か `html` |
| `-o`, `--output` | DOCX: 出力ファイル（省略時は `<input>.docx`）。HTML: 出力ディレクトリ（省略時は FrontMatter の `id`） |
| `-h`, `--help` | ヘルプ表示 |

### HTML の出力

出力ディレクトリに次のものを作ります。ディレクトリごと Web サーバーに置けば、そのまま読めます (`index.html` をブラウザで直接開いても読めます)。

```
<output>/
├── index.html   本文 (google-codelab / google-codelab-step)
├── img/         画像 (中身のハッシュの名前でコピー)。チャートは PNG にして置く
└── elements/    codelab-elements の JS と CSS
```

- 見た目は `resources/html/template.html` で決めています。claat の html テンプレートに、日本語で読みやすくする CSS と、表の見出し行・画像の説明と拡大・コードのコピーボタンを付ける JS を足したものです。
- codelab-elements は本家では `storage.googleapis.com/claat-public/` から読み込みますが、ここは 403 になるため `resources/html/elements/` に同梱し、出力にコピーしています。
- `> [!NOTE]` は `aside.special` (緑)、`> [!WARNING]` は `aside.warning` (オレンジ) になります。番号付きリストは、DOCX と同じく、見出しやステップが変わるまで番号を続けます。

## Markdown フォーマット

### FrontMatter (YAML)

ファイル先頭に `---` で囲んでメタデータを書きます。

````markdown
---
id: my-codelab
summary: Learn the basics
categories: MuleSoft
environments: Web
status: Published
feedback_link: https://example.com
authors: Your Name
---
````

### 本文

````markdown
# Codelab タイトル

## ステップ 1
Duration: 5:00
Environment: Web, Mobile

通常の段落・**太字**・`インラインコード`・リストが使えます。

- 箇条書き
1. 番号付きリスト

```console
$ npm install
```

```javascript
console.log("通常のコードブロック");
```

> [!NOTE]
> 補足のインフォボックス（肯定）

> [!WARNING]
> 注意のインフォボックス

| 項目 | 値 |
|------|-----|
| Name | value |

![代替テキスト](./image.png)

```chart
{ "type": "bar", "data": { "labels": ["A","B"], "datasets": [{ "data": [1,2] }] } }
```
````

### 記法一覧

| 記法 | 意味 |
|--|--|
| `# タイトル` | Codelab タイトル |
| `## タイトル` | ステップ区切り |
| `Duration: 5:00` | そのステップの所要時間 |
| `Environment: Web, Mobile` | 対象環境 |
| ` ```console ` | ターミナル表示（Consolas フォント） |
| ` ```chart ` | chart.js の設定 (JSON) からチャートを生成 |
| ` ```<lang> ` | 通常のコードブロック |
| `> [!NOTE]` / `> [!WARNING]` | インフォボックス |
| `\| ... \|` | テーブル（GitHub 形式） |
| `![alt](path)` | 画像（sharp で余白を自動トリム） |

## アーキテクチャ

```
src/codelab/
├── core.cljs   # CLI エントリーポイント・引数パース
├── meta.cljs   # FrontMatter (YAML) パーサー
├── parser.cljs # Markdown → AST 変換
├── docx.cljs   # AST → DOCX 変換（docx.js 使用）
└── html.cljs   # AST → HTML 変換（resources/html/ のテンプレートと codelab-elements を使う）
resources/html/
├── template.html  # HTML のテンプレート
└── elements/      # codelab-elements (出力にコピーする)
```

処理フロー: ファイル読み込み → `meta` + `parser` で AST 生成 → `docx` で DOCX 出力、または `html` で HTML 出力。
AST のノードタイプ: `:step` `:heading` `:paragraph` `:code-block` `:terminal` `:list` `:table` `:image-block` `:chart` `:infobox-positive` `:infobox-negative` など。

## 主要な依存関係

- [docx](https://www.npmjs.com/package/docx) — DOCX 生成
- [sharp](https://sharp.pixelplumbing.com/) / [image-size](https://www.npmjs.com/package/image-size) — 画像処理・サイズ取得
- [chart.js](https://www.chartjs.org/) / [chartjs-node-canvas](https://www.npmjs.com/package/chartjs-node-canvas) — チャート生成

## ライセンス

MIT

`resources/html/template.html` は claat のテンプレート、`resources/html/elements/` は [codelab-elements](https://github.com/googlecodelabs/codelab-elements) で、どちらも Google の Apache License 2.0 です。
