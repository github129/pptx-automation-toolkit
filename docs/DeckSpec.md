# PowerPoint 生成エンジン — 使い方

フロー「PowerPoint作成」は、次の 2 つのモードで PowerPoint を作ります。

| モード | 入力 | 向いている人 |
|---|---|---|
| **かんたんモード**（おすすめ） | Excel（[`templates/スライド入力.xlsx`](../templates/スライド入力.xlsx)）または短い JSON | 市民開発者・資料を作る人 |
| 詳細モード | 構成 JSON（このページの §1 以降） | レイアウトや座標まで作り込みたい人 |

## かんたんモード

1. `スライド入力.xlsx` をコピーして OneDrive に保存する
2. 「設定」シートで **デザイン** を選ぶ（ビジネス / コーポレート / ナチュラル / モダン / ポップ / ミニマル / エレガント / テック）
3. 「スライド」シートに **1 行 = 1 スライド** で「種類・タイトル・内容・画像・ノート」を書く
4. フローを実行して Excel ファイルのパスを指定する

種類は 21 種類: 表紙 / 目次 / 章 / 箇条書き / カード / 数字 / ステップ / タイムライン / 比較 / 表 /
棒グラフ / 積み上げ棒グラフ / 横棒グラフ / 折れ線グラフ / 面グラフ / 円グラフ / ドーナツグラフ / 画像 / 引用 / メッセージ / 終わり。
全種類の見本は [`samples/all-types.json`](../samples/all-types.json)です。
図形の配置・色・フォントはデザインが自動で決めます。「内容」の書き方は Excel の「使い方」シートを参照してください。

- 項目は `|` で区切る（例: `48.2億円 | 売上高 | 前年比 +12%`）
- 行頭の全角スペースで 2 段目の箇条書き、行頭の「・」「- 」は自動で除去
- 色を変えたいときは「メインカラー」に `#E11D48` のように入力

JSON で渡す場合も同じ項目名です（[`samples/simple-deck.json`](../samples/simple-deck.json)）。

```json
{ "デザイン": "モダン", "フッター": "社外秘",
  "slides": [
    { "種類": "表紙", "タイトル": "新サービスのご紹介", "内容": "2026/10/27" },
    { "種類": "数字", "タイトル": "実績", "内容": "120社 | 導入企業
98% | 継続率" }
  ] }
```

かんたんモードの行に `"layout"` を書くと、その行だけ詳細モードのスライドとして扱います。

### 社内テンプレートを使う

「設定」シートの **テンプレート** に、社内テンプレート（`.potx` / `.pptx`）の OneDrive パスを入力します。

| 項目 | テンプレートなし | テンプレートあり |
|---|---|---|
| 背景・飾り・ロゴ・帯 | デザインのもの | **テンプレートのスライドマスター** |
| 配色・フォント | デザインのもの | **テンプレートのテーマ**（「メインカラー」「フォント」を書いた場合だけ上書き） |
| タイトル・本文の位置と書式 | デザインのもの | **テンプレートのレイアウト** |
| カード・数字・表などの図形の形 | デザインのもの | デザインのもの（色はテンプレートの配色） |
| 図形を並べる範囲 | 16:9 の既定 | テンプレートの「タイトルとコンテンツ」の本文枠に合わせる（4:3 も可） |

- 使うレイアウト: タイトル スライド（表紙・終わり）、セクション見出し（章）、タイトルとコンテンツ、2 つのコンテンツ、比較、タイトルのみ、白紙。
  名前ではなく PowerPoint のレイアウトの種類で探すので、社内向けに名前を変えたレイアウトでも使えます。
- テンプレートにないレイアウトは近いものに置き換え、結果の警告に表示します（例: 「比較」がなければ「2 つのコンテンツ」）。
- テンプレート内の既存スライドは使いません（削除して作ります）。

---

# 詳細モード — 構成 JSON リファレンス

JSON を書き換えるだけで、スライド枚数・レイアウト・デザイン・データを自由に変えられます。

- エンジン本体: [`scripts/BuildDeck.ts`](../scripts/BuildDeck.ts)（Office Script）
- サンプル: [`samples/sample-deck.json`](../samples/sample-deck.json)（組み込みデザイン・主要機能） / [`samples/template-deck.json`](../samples/template-deck.json)（テンプレート利用）
- OneDrive 上の配置: 構成 JSON は `/PowerPoint/decks/`、画像は `/PowerPoint/images/`、テンプレートは `/PowerPoint/templates/`、出力先の既定は `/PowerPoint/output/`
- フローの作り方: [flow-setup.md](flow-setup.md)

---

## 1. 全体構造

```json
{
  "template": "",                       // テンプレート (.pptx/.potx) の OneDrive パス。空なら組み込みデザイン
  "templateSlides": "remove",           // テンプレート内の既存スライド: remove(削除) / keep(残す)
  "assets": [ { "name": "logo", "path": "/logo.png" } ],   // 画像 (OneDrive パス) に名前を付ける
  "output": { "folder": "/PowerPoint/output", "fileName": "報告書_{timestamp}" },

  "size": "16:9",                       // 16:9 / 4:3 / 16:10 / A4 / A4portrait / {"width":..,"height":..}
  "unit": "cm",                         // 位置・サイズの単位: cm / mm / in / pt / px / emu
  "properties": { "title": "文書タイトル", "author": "作成者", "subject": "件名" },
  "vars": { "company": "タクマル工房" },  // {{company}} で参照できる変数

  "theme":   { ... },                   // 配色・フォント (§5)
  "master":  { ... },                   // スライドマスター (§6)
  "styles":  { "名前": { ...書式... } }, // 再利用できる名前付き書式 (§8)
  "defaults": { "font": "Meiryo", "size": 16 },  // テキストボックスの既定書式

  "footer": "社外秘", "slideNumber": true, "date": true,  // 全スライド共通のフッター
  "transition": "fade",                 // 全スライド共通の画面切り替え

  "slides": [ { ... }, { ... } ]        // スライド (§2)
}
```

`template` / `assets` / `output` はフローが読み、それ以外は Office Script が読みます。
`output.fileName` の `{timestamp}` は実行日時（日本時間 `yyyyMMdd_HHmmss`）に置き換わります。

---

## 2. スライド

```json
{ "layout": "content", "title": "タイトル", "body": ["箇条書き 1", "箇条書き 2"], "notes": "発表者ノート" }
```

| キー | 内容 |
|---|---|
| `layout` | レイアウトの名前・キー・番号（§3）。省略時は内容から自動選択 |
| `title` / `subtitle` / `body` / `left` … | **スロット**。レイアウトのプレースホルダーに差し込む内容（§4） |
| `elements` | 自由配置の図形・テキスト・画像・表・グラフ（§7） |
| `notes` | 発表者ノート（文字列または配列） |
| `background` | このスライドだけの背景（§9 の塗り） |
| `section` | ここから始まるセクション名 |
| `transition` | 画面切り替え: `fade` `push` `wipe` `split` `cover` `pull` `cut` `dissolve` `zoom` `random` `circle` `wheel` `blinds` `checker` `none`、または `{ "type": "push", "dir": "l", "duration": 0.5 }` |
| `hidden` | `true` で非表示スライド |
| `footer` / `slideNumber` / `date` | このスライドだけのフッター設定（`false` で消す） |
| `master` | マスターが複数あるテンプレートで、使うマスターの番号または名前 |
| `base` / `replace` | テンプレートの既存スライドを複製（§10） |
| `each` | データの件数分スライドを繰り返し生成（§11） |

---

## 3. レイアウト

### 組み込みデザイン（`template` が空のとき）

| `layout` | 用途 | スロット |
|---|---|---|
| `title` | 表紙 | `title`, `subtitle` |
| `section` | セクション見出し | `title`, `subtitle` |
| `content` | タイトルとコンテンツ | `title`, `body` |
| `twoContent` | 2 カラム | `title`, `left`, `right` |
| `comparison` | 比較 | `title`, `leftHeading`, `left`, `rightHeading`, `right` |
| `threeContent` | 3 カラム | `title`, `left`, `center`, `right` |
| `titleOnly` | タイトルのみ（自由配置用） | `title` |
| `blank` | 白紙 | なし |
| `contentCaption` | 説明付きコンテンツ | `title`, `caption`, `body` |
| `pictureCaption` | 説明付きの図 | `title`, `picture`, `caption` |
| `imageRight` / `imageLeft` | 画像と箇条書き | `title`, `body`, `picture` |
| `quote` | 引用 | `quote`, `author` |
| `closing` | 結び | `title`, `subtitle` |

### テンプレート使用時

- **レイアウト名**（例 `"タイトルとコンテンツ"`）、**番号**（例 `6`）、または上表の **キー**（標準レイアウトなら `content` `comparison` などがそのまま使えます）で指定します。
- 実行結果の **「レイアウト一覧」** に、テンプレートのレイアウト名とスロット名が出力されます。初めて使うテンプレートは、一度実行して確認すると便利です。

---

## 4. スロット（プレースホルダーへの差し込み）

スロットに渡せる値:

| 値 | 結果 |
|---|---|
| `"文字列"` | 1 段落（改行 `\n` で段落を分割）。箇条書きの記号は付きません |
| `["A", "B"]` | 箇条書き。**入れ子の配列で下の階層**になります: `["A", ["A-1", "A-2"], "B"]` |
| `{ "text": ..., "size": 18, "color": "accent1" }` | 書式付きのテキスト（§8 の書式キーが使えます） |
| `{ "type": "table", ... }` / `{ "type": "chart", ... }` / `{ "type": "image", ... }` | プレースホルダーの位置に表・グラフ・画像を配置 |
| `"logo"`（`picture` スロット） | 画像アセット名。枠いっぱいに切り抜いて配置 |

スロット名の解決ルール（テンプレートでも共通）:
1. プレースホルダー名と一致（大文字・小文字・空白は無視）
2. `title` / `subtitle` / `picture` / `caption` / `left` / `right` / `leftHeading` / `rightHeading`
3. `body`, `body2`, `body3` … ＝ タイトル以外のプレースホルダーの順番
4. `#3` ＝ プレースホルダーの idx 番号を直接指定

文字が多い場合は、枠に収まるようにフォントを自動で縮小します（`"autofit": "none"` で無効化）。

### テキストの書き方

```json
"body": [
  "通常の文字。**太字** や [リンク](https://example.com) も書けます",
  { "text": "色・サイズ指定", "color": "accent2", "size": 24 },
  { "runs": ["一部だけ", { "text": "赤字", "color": "red", "bold": true }, "にする"] },
  { "text": "番号付き", "bullet": "number" },
  { "text": "記号なし", "bullet": false }
]
```

---

## 5. テーマ（配色・フォント）

```json
"theme": {
  "colors": { "accent1": "2563EB", "accent2": "0EA5A4", "dk2": "0F2A4A", "lt2": "E8EEF5" },
  "fonts":  { "major": "Aptos Display", "minor": "Aptos", "majorEa": "Yu Gothic", "minorEa": "Yu Gothic" }
}
```

- 色名: `dk1`(文字) `lt1`(背景) `dk2` `lt2` `accent1`〜`accent6` `hlink` `folHlink`
- テンプレート使用時も上書きできます（テンプレートの配色を一部だけ変える、など）。

---

## 6. スライドマスター

```json
"master": {
  "background": "FFFFFF",                       // マスターの背景 (§9)
  "titleStyle": { "size": 30, "color": "tx2", "bold": true, "font": "Meiryo" },  // 組み込みのみ
  "bodyStyle":  { "size": 20, "color": "tx1", "bulletColor": "accent1" },       // 組み込みのみ
  "bullets": ["•", "–", "▪"],                   // 箇条書き記号 (階層順)。組み込みのみ
  "titleBar": "accent1",                        // タイトル下のアクセントバー。false で非表示。組み込みのみ
  "logo": { "image": "logo", "position": "topRight", "h": 1.0 },   // 全スライドにロゴ
  "elements": [ { "type": "shape", ... } ],     // 全スライド共通の図形 ("front": true で前面)
  "layoutOptions": {                            // レイアウト単位の変更 (テンプレートでも可)
    "title":   { "background": { "gradient": ["dk2", "accent1"], "angle": 45 }, "elements": [ ... ] },
    "section": { "showMaster": false }
  },
  "layouts": [                                  // 独自レイアウトの追加 / 組み込みレイアウトの置き換え (組み込みのみ)
    { "key": "kpi4", "name": "KPI 4 分割", "background": "lt2", "placeholders": [
      { "name": "Title", "type": "title",   "x": 1.5, "y": 1,   "w": 30.8, "h": 2.5 },
      { "name": "A",     "type": "content", "x": 1.5, "y": 4.5, "w": 15,   "h": 6.5 },
      { "name": "B",     "type": "content", "x": 17.3,"y": 4.5, "w": 15,   "h": 6.5, "size": 16 }
    ] }
  ]
}
```

- `logo.position`: `topRight` `topLeft` `bottomRight` `bottomLeft` `topCenter` `bottomCenter`
- 独自レイアウトのプレースホルダー `type`: `title` / `content`（箇条書き） / `text`（記号なし） / `subtitle` / `picture`
- 独自レイアウトは `"layout": "kpi4"` で使い、スロットはプレースホルダー名（例 `"a"`, `"b"`）で指定します。

---

## 7. 要素（elements）— 自由配置

すべての要素で `x` `y` `w` `h` を指定します（単位は `unit`）。

- `"50%"` はスライド幅（高さ）に対する割合です。
- `"2in"` や `"30mm"` のように単位付きでも書けます。
- `x: "center"` は中央揃えです。
- 負の値は右端（下端）からの距離を表します。

| `type` | 主なキー |
|---|---|
| `text` | `text`（§4 と同じ書き方）、書式キー |
| `shape` | `shape`（`rect` `roundRect` `ellipse` `triangle` `rightArrow` `chevron` `homePlate` `wedgeRoundRectCallout` `star5` `flowChartDecision` など）、`fill` `line` `radius` `shadow` `rotation` `text` |
| `line` | `x` `y` `x2` `y2`、`line: { "color", "width", "dash" }`、`head` / `tail`（`arrow` `stealth` `oval` `diamond`） |
| `image` | `image`（アセット名）、`fit`（`contain` 収める / `cover` 切り抜き / `stretch`）、`shape: "ellipse"` で丸抜き、`w` か `h` だけなら縦横比を維持 |
| `table` | §7.1 |
| `chart` | §7.2 |

### 7.1 表

```json
{ "type": "table", "x": 2, "y": 4, "w": 30,
  "columns": [ { "key": "name", "label": "製品", "width": 8 }, { "key": "sales", "label": "売上", "format": "#,##0" } ],
  "data": [ { "name": "A", "sales": 1200 }, { "name": "B", "sales": 980 } ],
  "tableStyle": "medium2accent1", "headerFill": "dk2", "headerColor": "FFFFFF", "bandFill": "F3F4F6", "borderColor": "D1D5DB",
  "size": 14, "rowsPerSlide": 10 }
```

- `data`（オブジェクト配列。Power Automate の「選択」の出力をそのまま渡せます）または `rows`（2 次元配列。既定では 1 行目が見出し）
- `format`: `#,##0` / `0.0` / `0.0%` / `¥#,##0` など
- セルをオブジェクトにすると個別の書式や結合を指定できます: `{ "text": "合計", "bold": true, "fill": "FDE68A", "colspan": 2, "rowspan": 1 }`
- `tableStyle`: `medium2accent1`〜`6`、`medium2`、`light1accent1`、`grid`、`none`、または表スタイルの GUID
- `rowsPerSlide`: 行が多いときに自動で複数スライドに分割します（2 枚目以降のタイトルに「（続き）」が付きます。文言は `continuedSuffix` で変更可）

### 7.2 グラフ（ネイティブグラフ。PowerPoint で「データの編集」も可能）

```json
{ "type": "chart", "chartType": "column", "title": "売上推移",
  "data": [ { "月": "4月", "売上": 120, "目標": 110 }, { "月": "5月", "売上": 135, "目標": 120 } ],
  "dataLabels": true, "legend": "bottom", "numberFormat": "#,##0" }
```

- `chartType`: `column` `bar` `line` `area` `pie` `doughnut`
- データの指定方法は 2 通りです。
  - `data`: 1 列目が項目、数値の列が系列になります。`categoryKey` と `valueKeys` で明示もできます。
  - `categories` と `series: [{ "name", "values", "color" }]`
- `stacked`: `true`（積み上げ）または `"percent"`（100% 積み上げ）
- `dataLabels`: `true` で値を表示、`"percent"` で割合を表示（円グラフ・ドーナツ）
- その他: `colors`、`smooth`、`markers`、`gridlines`、`axisMin` / `axisMax`、`hole`（ドーナツの穴の大きさ %）、`size`（文字サイズ）

---

## 8. 書式キー（テキスト・図形・表・スロット共通）

| キー | 例 |
|---|---|
| `font` / `fontEa` | `"Meiryo"`（`fontEa` は日本語用。省略時は `font` と同じ） |
| `size` | `18`（pt） |
| `color` | `"1F4E79"` `"#1F4E79"` `"accent1"` `"accent1+40"`（40% 明るく）`"accent1-25"`（25% 暗く）`"red"` |
| `bold` `italic` `underline` `strike` | `true` |
| `highlight` | `"FFFF00"` |
| `align` / `valign` | `left` `center` `right` `justify` / `top` `middle` `bottom` |
| `lineSpacing` / `spaceBefore` / `spaceAfter` | `1.2`（倍） / `6`（pt） |
| `bullet` / `bulletColor` | `true` `false` `"number"` `"★"` / `"accent2"` |
| `fill` / `line` | §9 |
| `margin` | 内側の余白 `0.3` または `[左, 上, 右, 下]` |
| `autofit` | `shrink`（既定。はみ出す場合は縮小） / `resize`（枠を広げる） / `none` |
| `wrap` | `false` で折り返さない |
| `rotation` `shadow` `radius` `link` | `15` / `true` / `10`（角丸の %） / `"https://..."` |
| `style` | `"card"` や `"card kpi"` のように `styles` に定義した名前を参照します（複数可） |

---

## 9. 塗り（`fill` / `background`）と線（`line`）

```json
"fill": "accent1"                                         // 単色
"fill": { "color": "000000", "transparency": 40 }         // 透過 (%)
"fill": { "gradient": ["accent1", "accent5"], "angle": 45 } // グラデーション
"fill": { "image": "photo" }                              // 画像で塗りつぶし
"fill": "none"
"line": "BFBFBF"  /  "line": { "color": "tx1", "width": 2, "dash": "dash" }  /  "line": "none"
```

---

## 10. テンプレートの既存スライドを複製する（`base`）

テンプレートに `{{customer}}` のような差し込み用の文字を入れたスライドを作っておくと、そのスライドを複製して値を差し込めます。

```json
"vars": { "customer": "株式会社サンプル" },
"slides": [
  { "base": 1 },                                                         // 1 枚目を複製 (vars は自動で差し込み)
  { "base": 2, "replace": { "{{address}}": "東京都…" }, "body": ["…"] }  // 置換＋スロット差し込みも可
]
```

- 書式が混在していて、トークンが複数のランに分かれている場合も置換されます（その段落は先頭の書式にそろいます）。
- 複製したスライドにも `elements` / `notes` / `background` / `transition` を追加できます。

---

## 11. 変数と繰り返し

- `vars` の値は、構成 JSON 内のすべての文字列で `{{名前}}` として使えます。値そのもの（`"{{list}}"`）を書くと、配列や数値をそのまま渡せます。
- `each` を書くと、配列の要素ごとにスライドが 1 枚ずつ作られます。要素の各キーと `{{index}}`（1 から始まる番号）が使えます。

```json
{ "each": [ { "name": "東日本", "sales": 2210 }, { "name": "西日本", "sales": 1730 } ],
  "layout": "content", "title": "{{index}}. {{name}}", "body": ["売上: {{sales}} 百万円"] }
```

---

## 12. 他のフローからデータを流し込むには

1. SharePoint や Excel などからデータを取得し、「選択」アクションで `[{...}]` の形に整えます。
2. 「作成」アクションで構成 JSON を組み立てます。表の `data` やグラフの `data` に、手順 1 の出力を差し込みます。
3. 「ファイルの作成」(OneDrive) で `/PowerPoint/decks/xxx.json` に保存します。
4. このフローを、そのパスを指定して実行します。

フローを毎回手動で実行したくない場合は、トリガーを「ファイルが作成されたとき (OneDrive)」に変更すると、`decks` フォルダーに JSON を置くだけで自動生成できます（標準コネクタの範囲で可能です）。

---

## 13. 制限事項

- 画像は PNG / JPEG / GIF / BMP に対応しています（SVG は PNG に変換してください）。
- 「スクリプトの実行」の制限: 1 回あたり 120 秒、引数の合計は約 28MB、1 日 1,600 回までです。テンプレートと画像の合計は数 MB 以内が目安です。
- テキストの自動縮小は文字数からの推定です。PowerPoint でテキストを編集すると、縮小率は再計算されます。
