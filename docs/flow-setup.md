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
スクリプトを **2 回** 呼ぶのがポイントです。

1. 1 回目（`listFilesOnly = true`）: 使う画像とテンプレートのパスを調べる
2. フローが OneDrive からそれらのファイルを取得して Base64 で集める
3. 2 回目: 画像・テンプレートを渡して PPTX を生成する

```
[トリガー: 構成JSONのパス (空なら Excel)]
  ├─ 変数の初期化 deckJson / assets / templateBase64
  ├─ 条件: パスが空でない → 構成JSONの取得 → deckJson に設定
  ├─ スクリプトの実行 (ファイル一覧)  listFilesOnly = はい
  ├─ Apply to each (files) → 画像の取得 → assets に追加
  ├─ 条件: template が空でない → テンプレートの取得 → templateBase64 に設定
  ├─ スクリプトの実行 (PPTX生成)
  └─ ファイルの作成 (OneDrive)
```

### 手順

アクション名は式の中で参照するため、表の **名前** のとおりに変更してください（名前の変更は各アクションの「…」→「名前の変更」）。

| # | アクション（名前） | 設定 |
|---|---|---|
| 1 | **手動でフローをトリガーします** | テキスト入力を 1 つ追加: `構成JSONのパス`（任意。例 `/PowerPoint/decks/all-types.json`。空なら Excel を読む） |
| 2 | **変数を初期化する** | 名前 `deckJson` / 種類 文字列 / 値 なし |
| 3 | **変数を初期化する** | 名前 `assets` / 種類 アレイ / 値 `[]` |
| 4 | **変数を初期化する** | 名前 `templateBase64` / 種類 文字列 / 値 なし |
| 5 | **条件** | `triggerBody()?['text']` が `次の値に等しくない`（空） |
| 5-はい-1 | OneDrive for Business **パスを使用したファイル コンテンツの取得**（`構成JSONの取得`） | ファイル パス: `triggerBody()?['text']`<br>詳細パラメーター「コンテンツ タイプの推測」: **いいえ** |
| 5-はい-2 | **変数の設定** | `deckJson` = `base64ToString(body('構成JSONの取得')?['$content'])` |
| 6 | Excel Online (Business) **スクリプトの実行**（`ファイル一覧`） | ファイル: `/PowerPoint/スライド入力.xlsx` / スクリプト: `BuildDeck`<br>deckJson: `variables('deckJson')`<br>listFilesOnly: **はい** |
| 7 | **Apply to each**（`画像ごと`） | 出力の選択: `outputs('ファイル一覧')?['body/result/files']` |
| 7-1 | OneDrive for Business **パスを使用したファイル コンテンツの取得**（`画像の取得`） | ファイル パス: `items('画像ごと')?['path']`<br>「コンテンツ タイプの推測」: **いいえ** |
| 7-2 | **配列変数に追加** | `assets` に `{"name": "@{items('画像ごと')?['name']}", "data": "@{body('画像の取得')?['$content']}"}` |
| 8 | **条件** | `outputs('ファイル一覧')?['body/result/template']` が `次の値に等しくない`（空） |
| 8-はい-1 | OneDrive for Business **パスを使用したファイル コンテンツの取得**（`テンプレートの取得`） | ファイル パス: `outputs('ファイル一覧')?['body/result/template']`<br>「コンテンツ タイプの推測」: **いいえ** |
| 8-はい-2 | **変数の設定** | `templateBase64` = `body('テンプレートの取得')?['$content']` |
| 9 | Excel Online (Business) **スクリプトの実行**（`PPTX生成`） | ファイル: `/PowerPoint/スライド入力.xlsx` / スクリプト: `BuildDeck`<br>deckJson: `variables('deckJson')`<br>templateBase64: `variables('templateBase64')`<br>assetsJson: `string(variables('assets'))`<br>listFilesOnly: **いいえ** |
| 10 | OneDrive for Business **ファイルの作成** | フォルダーのパス: `outputs('PPTX生成')?['body/result/folder']`<br>ファイル名: 下記の式<br>ファイル コンテンツ: `base64ToBinary(outputs('PPTX生成')?['body/result/base64'])` |
| 11 | （任意）**作成**（`結果`） | `outputs('PPTX生成')?['body/result/warnings']` と `['layouts']` を表示。警告や、テンプレートのレイアウト一覧を確認できます |

手順 10 のファイル名（`{timestamp}` を日本時間の `yyyyMMdd_HHmmss` に置き換える）:

```
concat(replace(outputs('PPTX生成')?['body/result/fileName'], '{timestamp}', convertFromUtc(utcNow(), 'Tokyo Standard Time', 'yyyyMMdd_HHmmss')), '.pptx')
```

> - Apply to each の **同時実行制御はオフ（既定）のまま** にしてください。オンにすると配列変数への追加が不安定になります。
> - 7-2 の `name` には、スクリプトが返した名前（かんたんモードではパスそのもの、詳細モードでは `assets` の `name`）をそのまま使います。
> - スクリプトを実行するブックは何でも構いません。構成 JSON を渡したときは、ブックの中身は読みません。

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
  `構成JSONの取得` を「ファイル コンテンツの取得」（ファイル識別子 = トリガーの `x-ms-file-id`）にします。
- **他システムのデータを流し込む**: 詳細は [DeckSpec.md §12](DeckSpec.md#12-他のフローからデータを流し込むには)。

---

## 3. よくあるエラー

| 症状 | 原因と対処 |
|---|---|
| `ブックに「スライド」シートが見つかりません` | 構成 JSON のパスが空で、実行対象のブックが `スライド入力.xlsx` ではない |
| `JSON` の解析エラー | 構成 JSON の文法ミス（末尾のカンマ、`"` の閉じ忘れなど）。VS Code などで確認してください |
| 画像が表示されない | `images` に置いたパスと、Excel / JSON に書いたパスが一致しているか確認 |
| `レイアウト「…」が見つかりません` | テンプレートのレイアウト名と違う。エラーに出る「使用可能」の名前かキー（`content` など）を使う |
| タイムアウト | 「スクリプトの実行」は 1 回 120 秒まで。スライド枚数・画像サイズを減らす |
| 生成した PPTX が開けない | 手順 10 のファイル コンテンツが `base64ToBinary(...)` になっているか確認 |
