# codelab-generator

Markdown から **Google Codelab 形式の DOCX** を生成する CLI ツール。
ClojureScript (shadow-cljs) で実装され、Node.js 上で動作します。

生成した DOCX は [claat](https://github.com/googlecodelabs/tools) などに通すことで Codelab の HTML に変換できます。

## 特徴

- Markdown 1 ファイルから Codelab 形式の DOCX を生成
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
node dist/codelab.js <input.md> [-o <output.docx>]

# npm link 後は codelab-gen コマンドで実行できます
codelab-gen my-codelab.md -o my-codelab.docx
```

| オプション | 説明 |
|--|--|
| `-o`, `--output` | 出力 DOCX ファイル（省略時は `<input>.docx`） |
| `-h`, `--help` | ヘルプ表示 |

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
└── docx.cljs   # AST → DOCX 変換（docx.js 使用）
```

処理フロー: ファイル読み込み → `meta` + `parser` で AST 生成 → `docx` で DOCX 出力。
AST のノードタイプ: `:step` `:heading` `:paragraph` `:code-block` `:terminal` `:list` `:table` `:image-block` `:chart` `:infobox-positive` `:infobox-negative` など。

## 主要な依存関係

- [docx](https://www.npmjs.com/package/docx) — DOCX 生成
- [sharp](https://sharp.pixelplumbing.com/) / [image-size](https://www.npmjs.com/package/image-size) — 画像処理・サイズ取得
- [chart.js](https://www.chartjs.org/) / [chartjs-node-canvas](https://www.npmjs.com/package/chartjs-node-canvas) — チャート生成

## ライセンス

MIT
