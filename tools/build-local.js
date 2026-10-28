/**
 * Office Script をローカル (Node.js) で動かして PPTX を作る動作確認用ツール。
 * Power Automate を使わずに、構成 JSON の書き方やエンジンの改修結果を手元で確認できます。
 *
 *   npm install
 *   node tools/build-local.js samples/all-types.json                 # → out/全スライド見本_xxxx.pptx
 *   node tools/build-local.js samples/template-deck.json --template corporate.pptx
 *   node tools/build-local.js --pptx samples/table-data.json "売上一覧"  # BuildPptx.ts (表 1 枚)
 *
 * 画像はファイル名 (パスの最後の部分) が同じものを samples/images/ から探します。
 * 見つからない場合は 1x1 のダミー画像で代用します。
 */
const fs = require("fs");
const path = require("path");
const ts = require("typescript");

const ROOT = path.join(__dirname, "..");
const OUT = path.join(ROOT, "out");
const IMAGES = path.join(ROOT, "samples", "images");
const DUMMY_PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

function loadScript(file) {
  const src = fs.readFileSync(path.join(ROOT, "scripts", file), "utf8");
  const js = ts.transpileModule(src, { compilerOptions: { target: ts.ScriptTarget.ES2019 } }).outputText;
  return new Function(`${js}\nreturn main;`)();
}

function timestamp() {
  const d = new Date(Date.now() + 9 * 3600 * 1000); // 日本時間
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getUTCFullYear()}${p(d.getUTCMonth() + 1)}${p(d.getUTCDate())}_${p(d.getUTCHours())}${p(d.getUTCMinutes())}${p(d.getUTCSeconds())}`;
}

function save(name, base64) {
  fs.mkdirSync(OUT, { recursive: true });
  const file = path.join(OUT, `${name.replace("{timestamp}", timestamp())}.pptx`);
  fs.writeFileSync(file, Buffer.from(base64, "base64"));
  return file;
}

function imageBase64(onedrivePath) {
  const local = path.join(IMAGES, path.basename(onedrivePath));
  if (fs.existsSync(local)) return fs.readFileSync(local).toString("base64");
  console.warn(`画像が見つからないためダミーを使います: ${onedrivePath}`);
  return DUMMY_PNG;
}

function buildDeck(deckPath, templatePath) {
  const main = loadScript("BuildDeck.ts");
  const deckJson = fs.readFileSync(deckPath, "utf8");
  // フローと同じく 2 回呼ぶ: 1 回目で必要な画像・テンプレートを調べ、2 回目で PPTX を作る
  const list = main({}, deckJson, "", "", true);
  const assets = list.files.map((f) => ({ name: f.name, data: imageBase64(f.path) }));
  let templateBase64 = "";
  if (list.template) {
    if (!templatePath) throw new Error(`テンプレート「${list.template}」を使う JSON です。--template <ローカルの .pptx> を指定してください`);
    templateBase64 = fs.readFileSync(templatePath).toString("base64");
  }
  const r = main({}, deckJson, templateBase64, JSON.stringify(assets), false);
  console.log(`スライド数: ${r.slideCount}`);
  if (r.warnings) console.log(`警告:\n${r.warnings}`);
  if (templateBase64 && r.layouts) console.log(`レイアウト一覧:\n${r.layouts}`);
  console.log(`出力: ${save(r.fileName, r.base64)}`);
}

function buildPptx(dataPath, title) {
  const main = loadScript("BuildPptx.ts");
  const base64 = main({}, title ?? "売上一覧", fs.readFileSync(dataPath, "utf8"));
  console.log(`出力: ${save("BuildPptx_{timestamp}", base64)}`);
}

const args = process.argv.slice(2);
if (args.length === 0) {
  console.log("使い方: node tools/build-local.js <構成JSON> [--template <.pptx>]  |  --pptx <表データJSON> [タイトル]");
  process.exit(1);
}
if (args[0] === "--pptx") {
  buildPptx(args[1], args[2]);
} else {
  const t = args.indexOf("--template");
  buildDeck(args[0], t >= 0 ? args[t + 1] : "");
}
