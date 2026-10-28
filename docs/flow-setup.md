# Power Automate フローの作り方

このリポジトリの Office Script を Power Automate から呼び出して、PowerPoint を OneDrive に保存するまでの手順です。
使うのは **標準コネクタだけ**（Excel Online (Business) / OneDrive for Business）です。Premium ライセンスは不要です。

- [0. 準備](#0-準備)
- [1. 入門フロー: 表 1 枚の PowerPoint（BuildPptx）](#1-入門フロー-表-1-枚の-powerpointbuildpptx)
- [2. 本格フロー: PowerPoint作成（BuildDeck）](#2-本格フロー-powerpoint作成builddeck)
- [3. よくあるエラー](#3-よくあるエラー)

---

## 0. 準備

### 0-1. OneDrive にフォルダーとファイルを置く

```
OneDrive
└─ PowerPoint/
   ├─ スライド入力.xlsx        ← templates/スライド入力.xlsx をアップロード
   ├─ decks/                   ← 構成 JSON (samples/*.json) を置く
   ├─ images/                  ← samples/images/*.png をアップロード
   ├─ templates/               ← 社内テンプレート (.pptx / .potx) を使う場合だけ
   └─ output/                  ← 生成した PowerPoint の保存先
```

### 0-2. Office Script を登録する

1. Excel for the web で `/PowerPoint/スライド入力.xlsx` を開く
2. **［自動化］** タブ → **［新しいスクリプト］**
3. 中身をすべて消して [`scripts/BuildDeck.ts`](../scripts/BuildDeck.ts) を貼り付け、スクリプト名を **BuildDeck** にして保存
4. 同じように [`scripts/BuildPptx.ts`](../scripts/BuildPptx.ts) を **BuildPptx** という名前で保存

> スクリプトは OneDrive の `/ドキュメント/Office Scripts/` に `.osts` として保存されます。
> 自動化タブが表示されない場合は、Microsoft 365 管理センターで Office Scripts が有効か確認してください。

---

## 1. 入門フロー: 表 1 枚の PowerPoint（BuildPptx）

データ（JSON 配列）から、タイトル + 表 1 枚の PowerPoint を作る最小構成です。Office Script で PPTX（ZIP）を組み立てる仕組みを理解するのに向いています。

| # | アクション | 設定 |
|---|---|---|
| 1 | **手動でフローをトリガーします** | — |
| 2 | **作成**（名前: `データ`） | [`samples/table-data.json`](../samples/table-data.json) の中身を貼り付け |
| 3 | Excel Online (Business) **スクリプトの実行**（名前: `PPTX生成`） | 場所: OneDrive for Business / ファイル: `/PowerPoint/スライド入力.xlsx` / スクリプト: `BuildPptx`<br>title: `売上一覧`<br>tableJson: `string(outputs('データ'))` |
| 4 | OneDrive for Business **ファイルの作成** | フォルダーのパス: `/PowerPoint/output`<br>ファイル名: `売上一覧.pptx`<br>ファイル コンテンツ: `base64ToBinary(outputs('PPTX生成')?['body/result'])` |

実際のデータを使う場合は、手順 2 を SharePoint の「複数の項目の取得」や Excel の「表内に存在する行を一覧表示」に替え、
**選択** アクションで列を絞ってから `string(body('選択'))` を `tableJson` に渡します。

> 数値の列は右寄せ・桁区切りになります。Excel から読んだ値は文字列になるため、
> 選択アクションで `int(item()?['売上'])` のように数値に変換してください。

---

## 2. 本格フロー: PowerPoint作成（BuildDeck）

Excel（かんたんモード）と構成 JSON（かんたん / 詳細モード）のどちらからでも PowerPoint を作れるフローです。

### 2-1. 作り方（動画）

▶ フローの作り方の動画: https://youtu.be/QpKEk2xwoOs

動画を見ながら、下の [2-3. フローの中身](#2-3-フローの中身) の表のとおりにアクションを追加してください。
アクション名は式の中で参照するので、表の名前のとおりに変更してください（各アクションの「…」→「名前の変更」）。

2 つの「スクリプトの実行」（「必要なファイルを確認」と「PowerPointを生成」）は、次のように設定します。

| 項目 | 設定値 |
|---|---|
| 場所 | OneDrive for Business |
| ドキュメント ライブラリ | OneDrive |
| ファイル | 動的なコンテンツ `対象ブック`（変数） |
| スクリプト | `BuildDeck` |
| deckJson | `構成JSON`（変数） |
| templateBase64 | 「必要なファイルを確認」は空 / 「PowerPointを生成」は `テンプレート`（変数） |
| assetsJson | 「必要なファイルを確認」は空 / 「PowerPointを生成」は式 `string(variables('画像'))` |
| listFilesOnly | 「必要なファイルを確認」は **はい** / 「PowerPointを生成」は **いいえ** |

> 「変数_対象ブック」の初期値には、構成 JSON（.json）を指定したときにスクリプトを実行する Excel のファイル ID を入れます
> （ブックの中身は読まないので `/PowerPoint/スライド入力.xlsx` で構いません）。Excel だけを使う場合は空のままで動きます。
> ファイル ID は、一度 `ファイルのパス` に `/PowerPoint/スライド入力.xlsx` を指定して実行し、
> 実行履歴の「対象ブックを設定」に表示される値をコピーするのが簡単です。

### 2-2. 実行する

| 入力 | 内容 |
|---|---|
| ファイルのパス（必須） | Excel（例 `/PowerPoint/スライド入力.xlsx`）または構成 JSON（例 `/PowerPoint/decks/all-types.json`）の OneDrive パス |
| テンプレートのパス（任意） | 社内テンプレート（.potx / .pptx）の OneDrive パス。空なら Excel の「設定」や JSON の `template` を使う |

`保存先`（既定 `/PowerPoint/output`）に PowerPoint が保存されます。最後の「結果」アクションで、保存先・スライド数・警告を確認できます。

### 2-3. フローの中身

スクリプトを **2 回** 呼ぶのがポイントです。

1. 1 回目（`listFilesOnly = はい`）: 使う画像とテンプレートのパスを調べる
2. フローが OneDrive からそれらのファイルを取得して Base64 で集める
3. 2 回目: 画像・テンプレートを渡して PPTX を生成する

| # | アクション名 | 種類 | 内容 |
|---|---|---|---|
| 1 | 手動でフローをトリガーします | トリガー | 入力 `ファイルのパス`（`text`）と `テンプレートのパス`（`text_1`） |
| 2 | 変数_対象ブック / 変数_構成JSON / 変数_テンプレート / 変数_画像 | 変数を初期化する | 文字列 / 文字列 / 文字列 / アレイ |
| 3 | JSONファイルか | 条件 | `ファイルのパス` が `.json` で終わるか |
| 3-はい | JSONを取得 → 構成JSONを設定 | パスを使用したファイル コンテンツの取得 → 変数の設定 | JSON の中身を `構成JSON` に入れる |
| 3-いいえ | Excelの情報を取得 → 対象ブックを設定 | パスによるファイル メタデータの取得 → 変数の設定 | 指定した Excel の ID を `対象ブック` に入れる |
| 4 | 必要なファイルを確認 | スクリプトの実行 | 1 回目（`listFilesOnly = はい`） |
| 5 | テンプレートのパス | 作成 | トリガーの入力を優先し、空ならスクリプトが返した `template` |
| 6 | テンプレートあり → テンプレートを取得 → テンプレートを設定 | 条件 → ファイル コンテンツの取得 → 変数の設定 | テンプレートを Base64 で `テンプレート` に入れる |
| 7 | 画像ごとに取得 → 画像を取得 → 画像を追加 | Apply to each → ファイル コンテンツの取得 → 配列変数に追加 | `{ "name", "data" }` を `画像` に追加（同時実行は 1） |
| 8 | PowerPointを生成 | スクリプトの実行 | 2 回目 |
| 9 | OneDriveに保存 | ファイルの作成 | `{timestamp}` を日本時間に置き換えて `.pptx` で保存 |
| 10 | 結果 | 作成 | 保存先・スライド数・警告 |

主な式:

```
構成JSONを設定:  if(contains(string(body('JSONを取得')), '$content-type'), base64ToString(body('JSONを取得')?['$content']), string(body('JSONを取得')))
対象ブックを設定: last(split(body('Excelの情報を取得')?['Id'], '.'))
OneDriveに保存 (ファイル名):  @{replace(body('PowerPointを生成')?['result']?['fileName'], '{timestamp}', formatDateTime(convertFromUtc(utcNow(), 'Tokyo Standard Time'), 'yyyyMMdd_HHmmss'))}.pptx
OneDriveに保存 (コンテンツ):  base64ToBinary(body('PowerPointを生成')?['result']?['base64'])
```

### スクリプトの戻り値

| キー | 内容 |
|---|---|
| `base64` | 生成した PPTX（Base64）。`listFilesOnly` のときは空 |
| `slideCount` | スライド枚数 |
| `fileName` / `folder` | 保存するファイル名（拡張子なし）と保存先フォルダー |
| `files` | 取得が必要な画像 `[{ "name", "path" }]` |
| `template` | 取得が必要なテンプレートのパス |
| `warnings` | 警告（改行区切り）。エラーにはならないが確認したほうがよい点 |
| `layouts` | 使えるレイアウトとスロット名の一覧（テンプレート利用時に便利） |

### 応用

- **JSON を置いたら自動生成**: トリガーを OneDrive for Business「ファイルが作成されたとき」（フォルダー `/PowerPoint/decks`）に替え、
  `ファイルのパス` の代わりにトリガーのファイル パスを使います。
- **他システムのデータを流し込む**: 詳細は [DeckSpec.md §12](DeckSpec.md#12-他のフローからデータを流し込むには)。

---

## 3. よくあるエラー

| 症状 | 原因と対処 |
|---|---|
| `ブックに「スライド」シートが見つかりません` | `ファイルのパス` に指定した Excel に「スライド」シートがない |
| スクリプトの実行で「ファイルが見つかりません」 | JSON 入力のとき「変数_対象ブック」に Excel のファイル ID が入っていない（[2-1](#2-1-作り方動画) の注記） |
| `JSON` の解析エラー | 構成 JSON の文法ミス（末尾のカンマ、`"` の閉じ忘れなど）。VS Code などで確認してください |
| 画像が表示されない | `images` に置いたパスと、Excel / JSON に書いたパスが一致しているか確認 |
| `レイアウト「…」が見つかりません` | テンプレートのレイアウト名と違う。エラーに出る「使用可能」の名前かキー（`content` など）を使う |
| タイムアウト | 「スクリプトの実行」は 1 回 120 秒まで。スライド枚数・画像サイズを減らす |
| 生成した PPTX が開けない | 「OneDriveに保存」のファイル コンテンツが `base64ToBinary(...)` になっているか確認 |
