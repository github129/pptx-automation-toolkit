# PPTX Automation Toolkit

**Power Automate + Office Scripts で、標準コネクタだけを使って PowerPoint（.pptx）を自動生成する** サンプル集です。
Premium コネクタも外部サービスも使いません。Office Script の中で PPTX（ZIP + XML）を組み立て、OneDrive に保存します。

▶ 解説動画: https://youtu.be/CDGWKGvRCKc

![全スライド見本](docs/images/all-types-preview.png)

## できること

- Excel に **1 行 = 1 スライド** で書くだけで、デザインの整った PowerPoint を作成（かんたんモード）
- 表紙・目次・カード・KPI・ステップ・タイムライン・比較・表・グラフ（ネイティブ）など **21 種類** のスライド
- **8 種類のデザイン**（ビジネス / コーポレート / ナチュラル / モダン / ポップ / ミニマル / エレガント / テック）
- **社内テンプレート（.pptx / .potx）** のマスター・レイアウトを使った生成、既存スライドの複製と `{{差し込み}}`
- 構成 JSON で座標・図形・表・グラフ・変数・繰り返しまで細かく指定（詳細モード）

## ファイル構成

```
├─ scripts/
│  ├─ BuildPptx.ts          入門: タイトル + 表 1 枚の PPTX を作る最小サンプル (約 260 行)
│  └─ BuildDeck.ts          本格: PowerPoint 生成エンジン (かんたん / 詳細モード、テンプレート対応)
├─ templates/
│  └─ スライド入力.xlsx     かんたんモードの入力シート (「Power Apps 概要」の見本入り)
├─ samples/
│  ├─ all-types.json        かんたんモード: 全 21 種類の見本
│  ├─ simple-deck.json      かんたんモード: 最小の例
│  ├─ sample-deck.json      詳細モード: 組み込みデザインで主要機能を使う例
│  ├─ template-deck.json    詳細モード: 社内テンプレートを使う例
│  ├─ table-data.json       BuildPptx 用の表データ
│  └─ images/               サンプル画像 (image.png / logo.png)
├─ docs/
│  ├─ flow-setup.md         Power Automate フローの作り方
│  └─ DeckSpec.md           入力の書き方 (かんたんモード / 構成 JSON リファレンス)
└─ tools/                   ローカル動作確認用 (配布先での利用は任意)
```

## はじめ方

1. このリポジトリをダウンロード（**Code → Download ZIP**）
2. OneDrive に `/PowerPoint/` フォルダーを作り、`スライド入力.xlsx`、`samples/images/*.png`（→ `/PowerPoint/images/`）、`samples/*.json`（→ `/PowerPoint/decks/`）をアップロード
3. Excel for the web の **［自動化］→［新しいスクリプト］** に `scripts/BuildDeck.ts`（と `BuildPptx.ts`）を貼り付けて保存
4. フロー「PowerPoint作成」を作成（作り方の動画: https://youtu.be/QpKEk2xwoOs / 手順: [docs/flow-setup.md](docs/flow-setup.md#2-本格フロー-powerpoint作成builddeck)）
5. フローを実行（ファイルのパス: `/PowerPoint/スライド入力.xlsx`）→ `/PowerPoint/output/` に PowerPoint が保存されます

まずは仕組みを知りたい場合は、[入門フロー（BuildPptx）](docs/flow-setup.md#1-入門フロー-表-1-枚の-powerpointbuildpptx) から試すのがおすすめです。

## 入力の書き方

| モード | 入力 | サンプル |
|---|---|---|
| かんたんモード（Excel） | `スライド入力.xlsx` の「設定」「スライド」シート | `templates/スライド入力.xlsx` |
| かんたんモード（JSON） | `{"デザイン": "モダン", "slides": [{"種類": "…", "タイトル": "…", "内容": "…"}]}` | `samples/simple-deck.json`, `samples/all-types.json` |
| 詳細モード（JSON） | レイアウト・座標・テンプレートまで指定する構成 JSON | `samples/sample-deck.json`, `samples/template-deck.json` |

書き方の詳細は [docs/DeckSpec.md](docs/DeckSpec.md) を参照してください。

> `samples/template-deck.json` は、`/PowerPoint/templates/corporate.pptx` に **ご自身の社内テンプレート** を置いて使う例です
> （テンプレートは同梱していません）。1・2 枚目に `{{customer}}` や `{{address}}` を書いたスライドがある想定です。

## 必要な環境

- Microsoft 365 の職場・学校アカウント（Office Scripts と Power Automate が使えること）
- Power Automate の標準コネクタ: Excel Online (Business)、OneDrive for Business

## 制限事項

- 「スクリプトの実行」は 1 回 120 秒、引数の合計は約 28MB、1 日 1,600 回まで
- 画像は PNG / JPEG / GIF / BMP（SVG は PNG に変換）
- テキストの自動縮小は文字数からの推定です（PowerPoint で編集すると再計算されます）

## ローカルでの動作確認（開発者向け）

Node.js があれば、Power Automate を使わずに手元で PPTX を生成して確認できます。

```bash
npm install
```

```bash
npm run build:all-types
```

`out/` に PowerPoint が出力されます。ほかのサンプルは `build:simple` / `build:sample` / `build:table`、
任意の JSON は `node tools/build-local.js <構成JSON> [--template <.pptx>]` で試せます。

## 免責事項

- このリポジトリは **個人が製作したサンプル** です。Microsoft およびその他の企業・団体とは関係ありません。
- 個人・商用を問わず、**自由に使用・改変・再配布して構いません**。連絡や許可は不要です。
- 内容は現状のまま提供します。動作や正確性は保証しません。
- このリポジトリのスクリプト・フロー・資料を使ったことで生じたいかなる損害・トラブルについても、製作者は **一切の責任を負わず、補償もしません**。ご自身の責任でお使いください。
- 業務で使う場合は、事前にテスト環境で動作を確認し、所属組織のルール（Power Platform の利用ポリシーなど）に従ってください。
- 個別のサポートや不具合修正のお約束はできません。
