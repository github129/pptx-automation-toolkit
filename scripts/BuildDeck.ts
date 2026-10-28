/**
 * Office Script: BuildDeck — PowerPoint (PPTX) 生成エンジン
 *
 * Power Automate の Excel Online (Business)「スクリプトの実行」から呼び出す (Premium コネクタ不要)。
 * 入力は次の 3 通り。
 *   1. かんたんモード (Excel) : deckJson を空にすると、実行対象ブックの「スライド」「設定」シートを読み込む
 *   2. かんたんモード (JSON)  : {"デザイン":"ビジネス","slides":[{"種類":"箇条書き","タイトル":"…","内容":"…"}]}
 *   3. 詳細モード (JSON)      : レイアウト・座標・テンプレートまで指定できる構成 JSON (DeckSpec.md 参照)
 *
 * deckJson       : 任意。構成 JSON。空なら実行対象ブックの表を読み込む
 * templateBase64 : 任意。テンプレート PPTX/POTX の Base64 (詳細モードのみ)
 * assetsJson     : 任意。画像 [{"name":"/images/a.png","data":"<Base64>"}] の JSON
 * listFilesOnly  : true なら PPTX を作らず、フローが取得すべき画像とテンプレートのパスだけを返す
 */
function main(workbook: ExcelScript.Workbook, deckJson?: string, templateBase64?: string, assetsJson?: string, listFilesOnly?: boolean): BuildResult {
  const text = (deckJson ?? "").trim() !== "" ? deckJson ?? "" : JSON.stringify(readWorkbookInput(workbook));
  return listFilesOnly ? listFiles(text) : buildDeck(text, templateBase64 ?? "", assetsJson ?? "");
}

// =====================================================================
// 型定義 (構成 JSON)
// =====================================================================
interface FileRef { name: string; path: string }
interface BuildResult {
  base64: string; slideCount: number; fileName: string; folder: string; layouts: string; warnings: string;
  files: FileRef[]; template: string;
}

type Len = number | string;
type FillSpec = string | FillObj;
interface FillObj { color?: string; transparency?: number; gradient?: string[]; angle?: number; image?: string }
type LineSpec = string | LineObj;
interface LineObj { color?: string; width?: number; dash?: string; transparency?: number }

interface StyleSpec {
  style?: string; font?: string; fontEa?: string; size?: number; color?: string;
  bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; highlight?: string;
  align?: string; valign?: string; lineSpacing?: number; spaceBefore?: number; spaceAfter?: number;
  bullet?: boolean | string; bulletColor?: string;
  fill?: FillSpec; line?: LineSpec; margin?: Len | Len[]; autofit?: string; wrap?: boolean;
  rotation?: number; shadow?: boolean; radius?: number; link?: string; transparency?: number;
}
type TextValue = string | number | TextNode | TextValue[];
interface TextNode extends StyleSpec { text?: TextValue; runs?: (string | TextNode)[]; level?: number }

type CellValue = string | number | boolean | null | CellObj;
interface CellObj extends StyleSpec { text?: string | number | null; colspan?: number; rowspan?: number }
interface DataRow { [key: string]: CellValue }
interface ColumnSpec extends StyleSpec { key?: string; label?: string; width?: Len; format?: string }
interface SeriesSpec { name?: string; values?: (number | null)[]; color?: string }

interface ElementSpec extends TextNode {
  type?: string; name?: string; x?: Len; y?: Len; w?: Len; h?: Len; x2?: Len; y2?: Len;
  shape?: string; image?: string; fit?: string; alt?: string; head?: string; tail?: string;
  rows?: CellValue[][]; data?: DataRow[]; columns?: ColumnSpec[]; header?: boolean; tableStyle?: string;
  headerFill?: FillSpec; headerColor?: string; bandFill?: FillSpec; borderColor?: string; borderWidth?: number;
  rowHeight?: Len; colWidths?: Len[]; bandRow?: boolean; rowsPerSlide?: number; continuedSuffix?: string;
  chartType?: string; categories?: (string | number)[]; series?: SeriesSpec[]; title?: string;
  legend?: string | boolean; dataLabels?: boolean | string; stacked?: boolean | string; colors?: string[];
  numberFormat?: string; categoryKey?: string; valueKeys?: string[]; smooth?: boolean; markers?: boolean;
  gridlines?: boolean; axisMin?: number; axisMax?: number; hole?: number; front?: boolean;
}

type SlotValue = TextValue | ElementSpec;
interface TransitionObj { type?: string; duration?: number; dir?: string }
interface SlideSpec {
  layout?: string | number; master?: string | number; notes?: TextValue; background?: FillSpec;
  elements?: ElementSpec[]; transition?: string | TransitionObj; hidden?: boolean; section?: string;
  footer?: string | boolean; slideNumber?: boolean; date?: string | boolean;
  base?: number; replace?: { [token: string]: string };
  [slot: string]: SlotValue | ElementSpec[] | TransitionObj | FillSpec | boolean | { [token: string]: string } | undefined;
}

interface SizeObj { width: Len; height: Len }
interface ThemeSpec { name?: string; colors?: { [name: string]: string }; fonts?: FontsSpec }
interface FontsSpec { major?: string; minor?: string; majorEa?: string; minorEa?: string }
interface LogoSpec { image: string; position?: string; w?: Len; h?: Len; margin?: Len }
interface LayoutOptionSpec { background?: FillSpec; showMaster?: boolean; elements?: ElementSpec[] }
interface CustomPhSpec extends StyleSpec { name: string; type?: string; x: Len; y: Len; w: Len; h: Len; bullets?: boolean }
interface CustomLayoutSpec {
  key: string; name?: string; type?: string; background?: FillSpec; showMaster?: boolean;
  placeholders?: CustomPhSpec[]; elements?: ElementSpec[];
}
interface MasterSpec {
  background?: FillSpec; titleStyle?: StyleSpec; bodyStyle?: StyleSpec; bullets?: string[];
  titleBar?: boolean | string; logo?: LogoSpec; elements?: ElementSpec[];
  layouts?: CustomLayoutSpec[]; layoutOptions?: { [layout: string]: LayoutOptionSpec };
}
interface PropsSpec { title?: string; author?: string; subject?: string }
interface DeckSpec {
  size?: string | SizeObj; unit?: string; templateSlides?: string;
  theme?: ThemeSpec; master?: MasterSpec; styles?: { [name: string]: StyleSpec }; defaults?: StyleSpec;
  footer?: string; slideNumber?: boolean; date?: string | boolean; transition?: string | TransitionObj;
  properties?: PropsSpec; slides?: SlideSpec[];
  output?: { folder?: string; fileName?: string };  // フロー側で使用 ({{変数}} 展開後の値を返す)
}

type Json = string | number | boolean | null | Json[] | JsonObj;
interface JsonObj { [key: string]: Json }

// =====================================================================
// 内部の型
// =====================================================================
interface RawEntry { method: number; crc: number; size: number; comp: Uint8Array }
interface Part { raw?: RawEntry; bytes?: Uint8Array; text?: string; dirty: boolean }
interface Rel { id: string; type: string; target: string; external: boolean }
interface Pkg {
  parts: { [name: string]: Part };
  rels: { [owner: string]: Rel[] };
  ctDefaults: { [ext: string]: string };
  ctOverrides: { [partName: string]: string };
}
interface Rect { x: number; y: number; w: number; h: number }
interface PhInfo { type: string; idx: number; name: string; rect: Rect | null; sizes: (number | null)[]; block: string }
interface LayoutInfo { part: string; name: string; type: string; keys: string[]; phs: PhInfo[]; master: MasterInfo; showMaster: boolean; label: string }
interface MasterInfo { part: string; name: string; theme: string; phs: PhInfo[]; layouts: LayoutInfo[]; titleSizes: (number | null)[]; bodySizes: (number | null)[] }
interface MediaInfo { ext: string; w: number; h: number }
interface DeckCtx {
  pkg: Pkg; spec: DeckSpec; W: number; H: number; unit: number; scale: number;
  assets: { [name: string]: string }; media: { [key: string]: string }; mediaInfo: { [part: string]: MediaInfo };
  warnings: string[]; masters: MasterInfo[]; styles: { [name: string]: StyleSpec };
  presPart: string; builtin: boolean; needNotesMaster: boolean; notesMaster: string;
  layoutKeys: { [part: string]: string }; tokens: { [token: string]: string };
}
interface PartCtx { deck: DeckCtx; part: string; nextId: number }
interface ParaOut { xml: string; plain: string; size: number; indent: number; spcBef: number; lnSpc: number }
interface NewSlide { part: string; section?: string }

// =====================================================================
// 定数
// =====================================================================
const XML_HEAD = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
const NS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT_P = "application/vnd.openxmlformats-officedocument.presentationml";
const R_SLIDE = `${REL}/slide`;
const R_LAYOUT = `${REL}/slideLayout`;
const R_MASTER = `${REL}/slideMaster`;
const R_THEME = `${REL}/theme`;
const R_IMAGE = `${REL}/image`;
const R_NOTES = `${REL}/notesSlide`;
const R_NOTES_MASTER = `${REL}/notesMaster`;
const R_CHART = `${REL}/chart`;
const R_LINK = `${REL}/hyperlink`;
const R_PACKAGE = `${REL}/package`;
const R_CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const UNIT_EMU: { [u: string]: number } = { emu: 1, in: 914400, cm: 360000, mm: 36000, pt: 12700, px: 9525 };
const SCHEME_COLORS = ["dk1", "lt1", "dk2", "lt2", "tx1", "tx2", "bg1", "bg2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
const NAMED_COLORS: { [k: string]: string } = {
  white: "FFFFFF", black: "000000", gray: "808080", grey: "808080", lightgray: "D9D9D9", darkgray: "404040",
  red: "E53935", blue: "1E88E5", green: "43A047", yellow: "FDD835", orange: "FB8C00", purple: "8E24AA",
};
const DEFAULT_COLORS: { [k: string]: string } = {
  dk1: "1E293B", lt1: "FFFFFF", dk2: "0F2A4A", lt2: "E8EEF5", accent1: "2563EB", accent2: "0EA5A4",
  accent3: "F59E0B", accent4: "E11D48", accent5: "7C3AED", accent6: "16A34A", hlink: "2563EB", folHlink: "7C3AED",
};
const TITLE_TYPES = ["title", "ctrTitle"];
const FOOTER_TYPES = ["dt", "ftr", "sldNum", "hdr", "sldImg"];
const RESERVED_SLIDE_KEYS = ["layout", "master", "notes", "background", "elements", "transition", "hidden", "section",
  "footer", "slidenumber", "date", "base", "replace", "each", "comment"];
const LAYOUT_TYPE_KEYS: { [type: string]: string[] } = {
  title: ["title", "titleslide"], obj: ["content", "titleandcontent"], secHead: ["section", "sectionheader"],
  twoObj: ["twocontent"], twoTxTwoObj: ["comparison"], titleOnly: ["titleonly"], blank: ["blank"],
  objTx: ["contentcaption", "contentwithcaption"], picTx: ["picturecaption", "picturewithcaption"], tx: ["text"],
};
const LAYOUT_FALLBACKS: { [key: string]: string[] } = {
  closing: ["title"], section: ["title"], quote: ["blank", "titleOnly"], imageright: ["twoContent", "content"],
  threecontent: ["twoContent", "content"], comparison: ["twoContent", "content"], twocontent: ["content"],
  titleonly: ["content"], blank: ["titleOnly", "content"],
};
const TABLE_STYLES: { [k: string]: string } = {
  medium2: "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}",
  medium2accent1: "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", medium2accent2: "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE8}",
  medium2accent3: "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}", medium2accent4: "{00A15C55-8517-42AA-B614-E9B94910E393}",
  medium2accent5: "{7DF18680-E054-41AD-8BC1-D1AEF772440D}", medium2accent6: "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}",
  light1accent1: "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}", grid: "{5940675A-B579-460E-94D1-54222C63F5DA}",
  none: "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
};
const PRESET_SHAPES = ["rect", "roundRect", "ellipse", "triangle", "rtTriangle", "diamond", "parallelogram", "trapezoid",
  "pentagon", "hexagon", "octagon", "star4", "star5", "star6", "star8", "star12", "rightArrow", "leftArrow", "upArrow",
  "downArrow", "leftRightArrow", "upDownArrow", "chevron", "homePlate", "notchedRightArrow", "bentArrow", "uturnArrow",
  "circularArrow", "plus", "mathPlus", "mathMinus", "mathMultiply", "heart", "cloud", "donut", "noSmoking", "blockArc",
  "arc", "pie", "chord", "frame", "halfFrame", "corner", "plaque", "can", "cube", "bevel", "foldedCorner", "smileyFace",
  "sun", "moon", "lightningBolt", "snip1Rect", "snip2SameRect", "round1Rect", "round2SameRect", "wedgeRectCallout",
  "wedgeRoundRectCallout", "wedgeEllipseCallout", "cloudCallout", "leftBracket", "rightBracket", "leftBrace", "rightBrace",
  "flowChartProcess", "flowChartDecision", "flowChartTerminator", "flowChartDocument", "flowChartAlternateProcess",
  "flowChartInputOutput", "flowChartPredefinedProcess", "flowChartConnector", "flowChartManualInput", "flowChartPreparation",
  "teardrop", "decagon", "dodecagon", "ribbon", "ribbon2", "wave", "doubleWave", "horizontalScroll", "verticalScroll"];

// =====================================================================
// メイン処理
// =====================================================================
function buildDeck(deckJson: string, templateBase64: string, assetsJson: string): BuildResult {
  const warnings: string[] = [];
  const useTemplate = templateBase64.trim() !== "";
  const pkg = useTemplate ? openTemplate(base64Decode(templateBase64)) : newPkg();
  const root = parseRoot(deckJson, warnings, useTemplate ? templateFrame(pkg) : undefined);
  const spec = JSON.parse(JSON.stringify(expandDeck(root))) as DeckSpec;
  const slides = spec.slides ?? [];
  if (slides.length === 0) throw new Error("slides が空です。1 枚以上のスライドを指定してください");

  const unitName = (spec.unit ?? "cm").toLowerCase();
  if (UNIT_EMU[unitName] === undefined) throw new Error(`unit「${spec.unit}」は未対応です (cm / mm / in / pt / px / emu)`);
  const deck: DeckCtx = {
    pkg,
    spec, W: 12192000, H: 6858000, unit: UNIT_EMU[unitName], scale: 1,
    assets: parseAssets(assetsJson), media: {}, mediaInfo: {}, warnings, masters: [],
    styles: spec.styles ?? {}, presPart: "ppt/presentation.xml", builtin: !useTemplate,
    needNotesMaster: false, notesMaster: "", layoutKeys: {}, tokens: {},
  };
  // vars の値はテンプレートを複製したスライド内の {{名前}} にも自動で差し込む
  const vars = root.vars;
  if (isObj(vars)) {
    for (const k of Object.keys(vars)) {
      const v = vars[k];
      if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") deck.tokens[`{{${k}}}`] = String(v);
    }
  }

  if (useTemplate) {
    deck.presPart = templatePresPart(pkg);
    const sz = /<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/.exec(getText(deck.pkg, deck.presPart));
    if (sz) { deck.W = +sz[1]; deck.H = +sz[2]; }
  } else {
    setSlideSize(deck);
  }
  deck.scale = deck.H / 6858000;
  if (!useTemplate) createBuiltinPackage(deck);

  deck.masters = analyzeMasters(deck);
  if (useTemplate) applyThemeOverrides(deck);
  applyMasterExtras(deck);

  const presXml = getText(deck.pkg, deck.presPart);
  const presRels = getRels(deck.pkg, deck.presPart);
  const origSlides: { id: number; part: string }[] = allMatches(/<p:sldId\b[^>]*\bid="(\d+)"[^>]*\br:id="([^"]+)"/g, presXml)
    .map(m => ({ id: +m[1], part: resolveTarget(deck.presPart, (presRels.filter(r => r.id === m[2])[0] ?? { target: "" }).target) }));
  const keepOriginal = (spec.templateSlides ?? "remove").toLowerCase() === "keep";
  const offset = keepOriginal ? origSlides.length : 0;

  const built: NewSlide[] = [];
  slides.forEach((s, i) => {
    try {
      built.push(s.base !== undefined
        ? copyTemplateSlide(deck, s, origSlides.map(o => o.part), offset + i + 1)
        : buildSlide(deck, s, offset + i + 1));
    } catch (e) {
      throw new Error(`スライド ${i + 1}: ${(e as Error).message}`);
    }
  });

  finalizePresentation(deck, origSlides, built, keepOriginal);
  const bytes = writePackage(deck.pkg);
  return {
    base64: base64Encode(bytes),
    slideCount: built.length + offset,
    fileName: spec.output?.fileName || "Deck_{timestamp}",
    folder: spec.output?.folder || "/PowerPoint/output",
    layouts: describeLayouts(deck),
    warnings: warnings.join("\n"),
    files: fileRefsOf(root), template: typeof root.template === "string" ? root.template : "",
  };
}

function parseRoot(deckJson: string, warnings: string[], frame?: SimpleFrame): JsonObj {
  const root = parseJson(deckJson, "構成JSON (deckJson)");
  if (!isObj(root)) throw new Error("構成JSON は { ... } 形式のオブジェクトである必要があります");
  return isSimpleInput(root) ? simpleToDeck(root, warnings, frame) : root;
}

function templatePresPart(pkg: Pkg): string {
  const mainRel = getRels(pkg, "").filter(r => r.type === `${REL}/officeDocument`)[0];
  if (!mainRel) throw new Error("テンプレートに presentation.xml が見つかりません (PPTX/POTX を指定してください)");
  return resolveTarget("", mainRel.target);
}

// テンプレートのレイアウトを役割 (表紙 / 章 / 本文 / タイトルのみ / 白紙 …) に振り分け、
// スライドサイズと本文枠 (cm) を、かんたんモードの配置領域にする
function templateFrame(pkg: Pkg): SimpleFrame {
  const presPart = templatePresPart(pkg);
  const sz = /<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/.exec(getText(pkg, presPart));
  const W = sz ? +sz[1] : 12192000, H = sz ? +sz[2] : 6858000;
  const m = analyzeMasters({ pkg, presPart, layoutKeys: {} } as unknown as DeckCtx)[0];
  const titleOf = (l: LayoutInfo): PhInfo | undefined => l.phs.filter(p => TITLE_TYPES.indexOf(p.type) >= 0)[0];
  const bigBody = (l: LayoutInfo): PhInfo | undefined =>
    contentPhs(l.phs).filter(p => p.type !== "pic" && p.rect && p.rect.w * p.rect.h > W * H * 0.2)[0];

  // 1. レイアウトの種類 → 2. レイアウト名 → 3. プレースホルダーの配置 の順で役割を決める (同じ役割は先頭のレイアウト)
  const roles: { [role: string]: LayoutInfo } = {};
  const set = (role: string, l: LayoutInfo): void => { if (role && !roles[role]) roles[role] = l; };
  const assigned = (l: LayoutInfo): boolean => Object.keys(roles).some(k => roles[k] === l);
  for (const l of m.layouts) set(TEMPLATE_TYPE_ROLES[l.type] ?? "", l);
  for (const l of m.layouts) {
    if (l.type !== "cust" || assigned(l)) continue;
    set((TEMPLATE_NAME_ROLES.filter(r => r[1].test(l.name))[0] ?? [""])[0], l);
  }
  for (const l of m.layouts) {
    if (l.type !== "cust" || assigned(l)) continue;
    const t = titleOf(l);
    if (!t) { if (contentPhs(l.phs).length === 0) set("blank", l); continue; }
    if (bigBody(l)) set("content", l);
    else if (t.rect && t.rect.y + t.rect.h / 2 < H * 0.25) set("titleOnly", l);
    else set(roles.cover ? "section" : "cover", l);
  }

  const cm = (v: number): number => r2(v / 360000);
  const f: SimpleFrame = { sw: cm(W), sh: cm(H), cl: cm(W * 0.055), cw: cm(W * 0.89), ct: cm(H * 0.24), ch: cm(H * 0.63), layouts: {} };
  const ph = (roles.content ? bigBody(roles.content) : undefined) ?? m.phs.filter(p => p.type === "body")[0];
  const rect = ph ? ph.rect : null;
  if (rect && rect.w > W * 0.3 && rect.h > H * 0.2) {
    f.cl = cm(rect.x); f.cw = cm(rect.w); f.ct = cm(rect.y); f.ch = cm(Math.min(rect.h, H * 0.92 - rect.y));
  }
  // 「タイトルのみ」のタイトルと重ならないようにする
  const top = roles.titleOnly ? titleOf(roles.titleOnly) : undefined;
  if (top && top.rect && cm(top.rect.y + top.rect.h) + 0.5 > f.ct) {
    const bottom = f.ct + f.ch;
    f.ct = r2(cm(top.rect.y + top.rect.h) + 0.5);
    f.ch = r2(bottom - f.ct);
  }

  for (const role of Object.keys(roles)) {
    const l = roles[role];
    const rest = contentPhs(l.phs).slice();
    const pick = (test: (p: PhInfo) => boolean): number | undefined => {
      const p = rest.filter(test)[0];
      if (!p) return undefined;
      rest.splice(rest.indexOf(p), 1);
      return p.idx;
    };
    const info: TemplateLayout = { name: l.name, subTitleType: l.phs.some(p => p.type === "subTitle") };
    info.number = pick(p => /番号|number|^no\b/i.test(p.name));
    info.subtitle = pick(p => p.type === "subTitle" || /サブ|副題|sub/i.test(p.name));
    info.body = pick(p => p.type !== "pic");
    (f.layouts as { [role: string]: TemplateLayout })[role] = info;
  }
  return f;
}

// 組み込みのレイアウト名 (かんたんモードが使うもの) → テンプレート内の役割
const TEMPLATE_TYPE_ROLES: { [type: string]: string } = {
  title: "cover", secHead: "section", obj: "content", titleOnly: "titleOnly", blank: "blank", twoObj: "twoContent", twoTxTwoObj: "comparison",
};
const TEMPLATE_NAME_ROLES: [string, RegExp][] = [
  ["cover", /表紙|タイトル\s*スライド|title\s*slide|cover/i], ["section", /章|セクション|section|扉|区切り|divider/i],
  ["titleOnly", /タイトルのみ|title\s*only/i], ["blank", /白紙|blank/i], ["content", /コンテンツ|本文|content/i],
];
const SIMPLE_LAYOUT_ROLES: { [layout: string]: string } = {
  title: "cover", closing: "cover", section: "section", content: "content", titleOnly: "titleOnly", blank: "blank",
  twoContent: "twoContent", comparison: "comparison",
};

// かんたんモードのスライドを、テンプレートにあるレイアウトとプレースホルダーに合わせて組み替える
function adaptToTemplate(s: JsonObj, f: SimpleFrame, chapterNum: string): JsonObj {
  const L = f.layouts ?? {};
  let role = SIMPLE_LAYOUT_ROLES[String(s.layout)] ?? "";
  if (!role) return s;
  // 本文用のレイアウトがないテンプレートは「タイトルのみ」に表・グラフ・箇条書きを図形として置く
  if ((role === "content" || role === "twoContent" || role === "comparison") && !L[role]) {
    s = slotsToElements(s, role);
    role = "titleOnly";
  }
  if (role === "section" && !L.section) role = "cover";
  if (role === "blank" && !L.blank) role = "titleOnly";
  if (role === "titleOnly" && !L.titleOnly && L.content) role = "content";
  const l = L[role];
  if (!l) return s;
  const out: JsonObj = Object.assign({}, s, { layout: l.name });
  // 表紙・章のページはフッターとページ番号を出さない (テンプレートの絵柄と重なりやすいため)
  if (role === "cover" || role === "section") { out.footer = false; out.slideNumber = false; }
  const move = (key: string, idx: number | undefined): void => {
    if (out[key] === undefined) return;
    if (idx !== undefined) out[`#${idx}`] = out[key];
    delete out[key];
  };
  if (!l.subTitleType) move("subtitle", l.subtitle ?? (role === "cover" || role === "section" ? l.body : undefined));
  if (role === "section" && l.number !== undefined) out[`#${l.number}`] = chapterNum;
  if (role === "content" || role === "titleOnly") move("body", l.body);
  return out;
}

function slotsToElements(s: JsonObj, role: string): JsonObj {
  const out: JsonObj = {};
  for (const k of Object.keys(s)) if (["body", "left", "right", "leftHeading", "rightHeading"].indexOf(k) < 0) out[k] = s[k];
  out.layout = "titleOnly";
  const put = (v: Json, x: number, y: number, w: number, h: number): Json =>
    isObj(v) && (v.type === "table" || v.type === "chart")
      ? Object.assign({}, v, { x: r2(x), y: r2(y), w: r2(w), h: r2(h) })
      : { type: "text", text: v, x: r2(x), y: r2(y), w: r2(w), h: r2(h), size: 20, color: "tx1", bullet: true, spaceAfter: 6, valign: "top" };
  const els: Json[] = [];
  const half = (CW - 1.2) / 2;
  if (role === "content") els.push(put(s.body ?? "", CL, CT, CW, CH));
  else if (role === "twoContent") {
    els.push(put(s.left ?? "", CL, CT, half, CH), put(s.right ?? "", CL + half + 1.2, CT, half, CH));
  } else {
    [["leftHeading", "left"], ["rightHeading", "right"]].forEach((keys, i) => {
      const x = CL + i * (half + 1.2);
      els.push({ type: "text", text: s[keys[0]] ?? "", x: r2(x), y: CT, w: r2(half), h: 1.4, size: 22, bold: true, color: "accent1", valign: "bottom" });
      els.push({ type: "line", x: r2(x), y: r2(CT + 1.5), x2: r2(x + half), y2: r2(CT + 1.5), line: { color: "accent1", width: 1.5 } });
      els.push(put(s[keys[1]] ?? "", x, CT + 1.8, half, CH - 1.8));
    });
  }
  out.elements = els.concat(Array.isArray(s.elements) ? s.elements : []);
  return out;
}

// フローが OneDrive から取得して assetsJson / templateBase64 に渡すファイルの一覧
function listFiles(deckJson: string): BuildResult {
  const warnings: string[] = [];
  const root = parseRoot(deckJson, warnings);
  const out = isObj(root.output) ? root.output : {};
  return {
    base64: "", slideCount: Array.isArray(root.slides) ? root.slides.length : 0,
    fileName: typeof out.fileName === "string" ? out.fileName : "", folder: typeof out.folder === "string" ? out.folder : "",
    layouts: "", warnings: warnings.join("\n"),
    files: fileRefsOf(root), template: typeof root.template === "string" ? root.template : "",
  };
}

function fileRefsOf(root: JsonObj): FileRef[] {
  const list = Array.isArray(root.assets) ? root.assets : [];
  const out: FileRef[] = [];
  for (const a of list) {
    if (isObj(a) && typeof a.name === "string" && typeof a.path === "string" && a.path !== "") out.push({ name: a.name, path: a.path });
  }
  return out;
}

function setSlideSize(deck: DeckCtx): void {
  const s = deck.spec.size ?? "16:9";
  const presets: { [k: string]: number[] } = {
    "16:9": [12192000, 6858000], "4:3": [9144000, 6858000], "16:10": [12192000, 7620000],
    "a4": [10692000, 7560000], "a4portrait": [7560000, 10692000], "widescreen": [12192000, 6858000],
  };
  if (typeof s === "string") {
    const p = presets[s.toLowerCase().replace(/\s/g, "")];
    if (!p) throw new Error(`size「${s}」は未対応です (16:9 / 4:3 / 16:10 / A4 / A4portrait / {width,height})`);
    deck.W = p[0]; deck.H = p[1];
  } else {
    deck.W = lenEmu(deck, s.width, "x") ?? 12192000;
    deck.H = lenEmu(deck, s.height, "y") ?? 6858000;
  }
}

function parseJson(text: string, label: string): Json {
  const t = (text ?? "").replace(/^﻿/, "").trim();
  if (t === "") throw new Error(`${label} が空です`);
  try {
    return JSON.parse(t) as Json;
  } catch (e) {
    throw new Error(`${label} の JSON が不正です: ${(e as Error).message}`);
  }
}

function parseAssets(text: string): { [name: string]: string } {
  const out: { [name: string]: string } = {};
  if (!text || text.trim() === "") return out;
  const j = parseJson(text, "assetsJson");
  if (Array.isArray(j)) {
    for (const a of j) {
      if (isObj(a) && typeof a.name === "string") {
        const d = a.data ?? a.content ?? a["$content"];
        if (typeof d === "string") out[a.name] = d;
        else if (isObj(d) && typeof d["$content"] === "string") out[a.name] = d["$content"] as string;
      }
    }
  } else if (isObj(j)) {
    for (const k of Object.keys(j)) {
      const d = j[k];
      if (typeof d === "string") out[k] = d;
      else if (isObj(d) && typeof d["$content"] === "string") out[k] = d["$content"] as string;
    }
  }
  return out;
}

// =====================================================================
// かんたんモード (Excel の表 / シンプルな JSON → 構成 JSON)
//   1 行 = 1 スライド。「種類」「タイトル」「内容」「画像」「ノート」を書くだけで、
//   選んだデザインに合わせて図形の配置・色・フォントを自動で決める。
// =====================================================================
interface DesignText { title?: string; sub?: string; num?: string }
interface DesignDef {
  desc: string; colors: { [k: string]: string }; font: string; titleFont?: string;
  master: JsonObj; cover: DesignText; section: DesignText; closing: DesignText;
  card: JsonObj; table: string; numShape: string;
}
interface SItem { head: string; fields: string[]; body: string[] }

// スライドの幅・高さと、図形を並べるコンテンツ領域 (cm)。既定は組み込みデザイン (16:9) 用で、
// 社内テンプレートを使うときは simpleToDeck でテンプレートの本文枠に置き換える
interface SimpleFrame { sw: number; sh: number; cl: number; cw: number; ct: number; ch: number; layouts?: { [role: string]: TemplateLayout } }
// テンプレート内の、役割ごとのレイアウトと差し込み先プレースホルダー (idx)
interface TemplateLayout { name: string; subTitleType: boolean; number?: number; subtitle?: number; body?: number }
const DEFAULT_FRAME: SimpleFrame = { sw: 33.867, sh: 19.05, cl: 1.86, cw: 30.14, ct: 4.6, ch: 12.0 };
let SW = 33.867, SH = 19.05, CL = 1.86, CW = 30.14, CT = 4.6, CH = 12.0;

const SIMPLE_TYPES: { [k: string]: string } = {
  "表紙": "cover", "タイトル": "cover", "cover": "cover", "title": "cover",
  "目次": "agenda", "アジェンダ": "agenda", "agenda": "agenda",
  "章": "section", "セクション": "section", "中表紙": "section", "section": "section",
  "箇条書き": "bullets", "本文": "bullets", "テキスト": "bullets", "まとめ": "bullets", "bullets": "bullets", "content": "bullets",
  "比較": "compare", "2列": "compare", "２列": "compare", "compare": "compare",
  "カード": "cards", "ポイント": "cards", "cards": "cards",
  "数字": "kpi", "数値": "kpi", "kpi": "kpi",
  "ステップ": "steps", "手順": "steps", "流れ": "steps", "プロセス": "steps", "steps": "steps",
  "タイムライン": "timeline", "スケジュール": "timeline", "年表": "timeline", "timeline": "timeline",
  "表": "table", "table": "table",
  "棒グラフ": "chart:column", "縦棒グラフ": "chart:column", "グラフ": "chart:column", "column": "chart:column",
  "横棒グラフ": "chart:bar", "bar": "chart:bar", "折れ線グラフ": "chart:line", "line": "chart:line",
  "円グラフ": "chart:pie", "pie": "chart:pie", "ドーナツグラフ": "chart:doughnut", "doughnut": "chart:doughnut",
  "面グラフ": "chart:area", "area": "chart:area", "積み上げ棒グラフ": "chart:stacked",
  "画像": "image", "写真": "image", "image": "image",
  "引用": "quote", "quote": "quote", "メッセージ": "message", "message": "message",
  "終わり": "closing", "おわり": "closing", "エンディング": "closing", "クロージング": "closing", "closing": "closing",
};

function designPresets(): { [name: string]: DesignDef } {
  // 位置の負の値は「右端 / 下端からの距離」と解釈されるため、スライド外にはみ出す図形は右・下基準の値に変換する
  const edge = (v: Len, size: Len, full: number): Len => typeof v === "number" && v < 0 && typeof size === "number" ? r2(v - full + size) : v;
  const R = (x: Len, y: Len, w: Len, h: Len, fill: Json, extra: JsonObj = {}): JsonObj =>
    Object.assign({ type: "shape", shape: "rect", x: edge(x, w, SW), y: edge(y, h, 19.05), w, h, fill, line: "none" } as JsonObj, extra);
  const O = (x: number, y: number, d: number, color: string, transparency: number): JsonObj =>
    R(x, y, d, d, { color, transparency }, { shape: "ellipse" });
  const grad = (a: string, b: string, angle: number): JsonObj => ({ gradient: [a, b], angle });
  const frame = (): Json[] => [
    R(1.0, 1.0, 31.87, 17.05, "none", { line: { color: "accent2", width: 1.25 } }),
    R(1.3, 1.3, 31.27, 16.45, "none", { line: { color: "accent2", width: 0.5 } }),
  ];
  const coverBar = (fill: string): JsonObj => R(2.54, 10.38, 2.71, 0.15, fill, { front: true });
  const base = (titleColor = "tx2"): JsonObj => ({ titleStyle: { size: 28, color: titleColor }, bodyStyle: { size: 20 } });
  const pal = (dk1: string, dk2: string, lt2: string, a: string[]): { [k: string]: string } => ({
    dk1, lt1: "FFFFFF", dk2, lt2, accent1: a[0], accent2: a[1], accent3: a[2], accent4: a[3], accent5: a[4], accent6: a[5], hlink: a[0], folHlink: a[4],
  });
  return {
    "ビジネス": {
      desc: "紺と青の定番デザイン。社内報告・提案など幅広く使える",
      colors: pal("1E293B", "0F2A4A", "E8EEF5", ["2563EB", "0EA5A4", "F59E0B", "E11D48", "7C3AED", "16A34A"]), font: "Yu Gothic",
      master: Object.assign(base(), { layoutOptions: {
        title: { background: "dk2", elements: [O(23, -6, 17, "accent1", 70), O(28, 11, 9, "accent2", 65)] },
        section: { background: "accent1", elements: [O(25, -5, 14, "lt1", 88)] },
        closing: { background: "dk2", elements: [O(-4, 12, 10, "accent1", 70), O(29, -3, 8, "accent2", 70)] },
      } }),
      cover: {}, section: { num: "accent1+40" }, closing: {},
      card: { shape: "roundRect", fill: "lt2", line: "none", radius: 6 }, table: "medium2accent1", numShape: "roundRect",
    },
    "コーポレート": {
      desc: "紺とオレンジの帯で引き締めた、堅めの社外向けデザイン",
      colors: pal("262626", "1F3864", "EDF1F7", ["2F5597", "ED7D31", "70AD47", "FFC000", "5B9BD5", "7F7F7F"]), font: "Yu Gothic",
      master: Object.assign(base(), { titleBar: "accent2", elements: [R(0, 0, "100%", 0.35, "dk2"), R(0, 0.35, "100%", 0.08, "accent2")], layoutOptions: {
        title: { background: "lt1", elements: [R(0, 0, "100%", 3.2, "dk2"), R(0, 3.2, "100%", 0.15, "accent2"), R(0, 16.6, "100%", 2.45, "dk2")] },
        section: { background: "dk2", elements: [R(0, 17.9, "100%", 0.3, "accent2")] },
        closing: { background: "dk2", elements: [R(0, 17.9, "100%", 0.3, "accent2")] },
      } }),
      cover: { title: "tx2", sub: "tx1+25" }, section: { num: "accent2" }, closing: {},
      card: { shape: "rect", fill: "lt1", line: { color: "tx1+75", width: 0.75 } }, table: "medium2accent1", numShape: "rect",
    },
    "ナチュラル": {
      desc: "緑とアースカラーのやわらかいデザイン。研修・社内イベント向け",
      colors: pal("2F3A2F", "2F5233", "EEF2E6", ["5B8C3A", "C08A3E", "7FA88A", "D9A441", "8C6E54", "5B8C85"]), font: "Yu Gothic",
      master: Object.assign(base(), { background: "F8F6F0", titleBar: "accent2", layoutOptions: {
        title: { background: grad("accent1", "dk2", 45), elements: [O(24, 9, 14, "lt1", 85), O(20.5, 13, 8, "accent4", 65)] },
        section: { background: "accent1", elements: [O(26, -4, 12, "lt1", 85)] },
        closing: { background: grad("accent1", "dk2", 45), elements: [O(-3, -3, 9, "lt1", 85)] },
      } }),
      cover: {}, section: { num: "accent1+40" }, closing: {},
      card: { shape: "roundRect", fill: "lt1", line: "none", radius: 10, shadow: true }, table: "medium2accent1", numShape: "ellipse",
    },
    "モダン": {
      desc: "黒地にシアン～紫のグラデーション。IT・新サービスの紹介向け",
      colors: pal("1F2937", "111827", "F1F5F9", ["0891B2", "7C3AED", "F43F5E", "F59E0B", "10B981", "3B82F6"]), font: "Yu Gothic",
      master: Object.assign(base(), { titleBar: false, elements: [R(0, 1.15, 0.6, 2.1, grad("accent1", "accent2", 90))], layoutOptions: {
        title: { background: "dk2", elements: [R(30.5, 0, 3.37, "100%", grad("accent1", "accent2", 90))] },
        section: { background: "dk2", elements: [R(0, 0, 0.6, "100%", grad("accent1", "accent2", 90))] },
        closing: { background: "dk2", elements: [R(0, 18.65, "100%", 0.4, grad("accent1", "accent2", 0))] },
      } }),
      cover: {}, section: { num: "accent1" }, closing: {},
      card: { shape: "rect", fill: "lt2", line: "none" }, table: "medium2accent1", numShape: "rect",
    },
    "ポップ": {
      desc: "オレンジとピンクの明るいデザイン。イベント・社内告知向け",
      colors: pal("3F3D56", "3F3D56", "FFEFD9", ["F97316", "EC4899", "14B8A6", "FACC15", "8B5CF6", "22C55E"]), font: "Yu Gothic",
      master: Object.assign(base(), { background: "FFFBF5", titleBar: "accent2", elements: [O(31, -1.8, 4, "accent4", 40)], layoutOptions: {
        title: { background: "accent1", elements: [O(23, -6, 18, "lt1", 85), O(26, 11, 10, "accent4", 30), O(-3, 14, 8, "accent2", 40), coverBar("lt1")] },
        section: { background: "accent2", elements: [O(25, -5, 14, "lt1", 85), O(-2, 15, 6, "accent4", 40)] },
        closing: { background: "accent1", elements: [O(-4, -4, 12, "lt1", 85), O(27, 12, 9, "accent4", 30)] },
      } }),
      cover: {}, section: { num: "lt1" }, closing: {},
      card: { shape: "roundRect", fill: "lt1", line: { color: "accent1+60", width: 1.5 }, radius: 14, shadow: true }, table: "medium2accent1", numShape: "ellipse",
    },
    "ミニマル": {
      desc: "白地に黒と赤だけのシンプルなデザイン。文字中心の資料向け",
      colors: pal("222222", "111111", "F3F4F6", ["E11D48", "374151", "9CA3AF", "F59E0B", "0EA5E9", "6B7280"]), font: "Yu Gothic",
      master: Object.assign(base(), { elements: [R(1.86, 17.35, 30.14, 0.03, "tx1+75")], layoutOptions: {
        title: { background: "lt1", elements: [R(2.54, 10.43, 28.8, 0.03, "tx1+75")] },
        section: { background: "lt2" },
        closing: { background: "lt1", elements: [R(15.43, 10.45, 3, 0.12, "accent1")] },
      } }),
      cover: { title: "tx2", sub: "tx1+40" }, section: { title: "tx2", sub: "tx1+40", num: "accent1" }, closing: { title: "tx2", sub: "tx1+40" },
      card: { shape: "rect", fill: "lt2", line: "none" }, table: "medium2accent2", numShape: "none",
    },
    "エレガント": {
      desc: "深い紫と金、明朝体の見出しで上品に。式典・役員向け",
      colors: pal("2B2B2B", "3B1F4F", "F4EFF6", ["6B3FA0", "B8902E", "A35D8C", "4E6E9E", "7A9E7E", "C46B5B"]), font: "Yu Gothic", titleFont: "Yu Mincho",
      master: Object.assign(base(), { titleBar: "accent2", elements: [R(0, 18.75, "100%", 0.3, "dk2")], layoutOptions: {
        title: { background: "dk2", elements: frame().concat([coverBar("accent2")]) },
        section: { background: "dk2", elements: frame() },
        closing: { background: "dk2", elements: frame() },
      } }),
      cover: {}, section: { num: "accent2" }, closing: {},
      card: { shape: "rect", fill: "lt1", line: { color: "accent2", width: 0.75 } }, table: "medium2accent1", numShape: "ellipse",
    },
    "テック": {
      desc: "藍色～水色のグラデーション。DX・技術系の発表向け",
      colors: pal("0F172A", "1E1B4B", "EEF2FF", ["4F46E5", "0EA5E9", "14B8A6", "F472B6", "A855F7", "F59E0B"]), font: "Yu Gothic",
      master: Object.assign(base(), { titleBar: "accent2", elements: [R(0, 0, "100%", 0.25, grad("accent1", "accent2", 0))], layoutOptions: {
        title: { background: grad("dk2", "accent1", 45), elements: [
          R(24, 2, 11, 11, "none", { shape: "ellipse", line: { color: "accent2", width: 1.5 } }),
          R(27.5, 7.5, 8, 8, "none", { shape: "ellipse", line: { color: "lt1", width: 0.75, transparency: 50 } }),
          R(0, 18.75, "100%", 0.3, grad("accent2", "accent3", 0)), coverBar("accent2")] },
        section: { background: grad("accent1", "accent2", 45), elements: [R(25, -4, 12, 12, "none", { shape: "ellipse", line: { color: "lt1", width: 1, transparency: 40 } })] },
        closing: { background: grad("dk2", "accent1", 45), elements: [R(0, 18.75, "100%", 0.3, grad("accent2", "accent3", 0))] },
      } }),
      cover: {}, section: { num: "lt1" }, closing: {},
      card: { shape: "roundRect", fill: "lt1", line: { color: "accent1+60", width: 1 }, radius: 6 }, table: "medium2accent1", numShape: "roundRect",
    },
  };
}

function readWorkbookInput(wb: ExcelScript.Workbook): JsonObj {
  const rows = readSheetRows(wb, "スライド");
  if (!rows) throw new Error("ブックに「スライド」シートが見つかりません。入力用テンプレート (スライド入力.xlsx) を使うか、構成 JSON を指定してください");
  const root: JsonObj = {};
  for (const r of readSheetRows(wb, "設定") ?? []) {
    const k = cellText(r["項目"]);
    if (k) root[k] = r["値"] === undefined ? "" : r["値"];
  }
  root.slides = rows.filter(r => ["種類", "タイトル", "内容", "画像"].some(k => cellText(r[k]) !== ""));
  return root;
}

// テーブル (名前 = シート名) を優先し、なければシートの使用範囲を 1 行目を見出しとして読む
function readSheetRows(wb: ExcelScript.Workbook, name: string): JsonObj[] | null {
  let values: (string | number | boolean)[][] = [];
  const table = wb.getTable(name);
  if (table) {
    values = table.getHeaderRowRange().getValues();
    if (table.getRowCount() > 0) values = values.concat(table.getRangeBetweenHeaderAndTotal().getValues());
  } else {
    const ws = wb.getWorksheet(name);
    if (!ws) return null;
    const used = ws.getUsedRange(true);
    if (used) values = used.getValues();
  }
  if (values.length === 0) return [];
  const head = values[0].map(h => String(h).trim());
  return values.slice(1).map(row => {
    const o: JsonObj = {};
    head.forEach((h, i) => { if (h) o[h] = row[i] as Json; });
    return o;
  });
}

function cellText(v: Json | undefined): string {
  if (v === undefined || v === null) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  return "";
}

function field(o: JsonObj, names: string[]): Json | undefined {
  for (const n of names) {
    const v = o[n];
    if (v !== undefined && v !== null && !(typeof v === "string" && v.trim() === "")) return v;
  }
  return undefined;
}

function isSimpleInput(root: JsonObj): boolean {
  if (field(root, ["design", "デザイン"]) !== undefined) return true;
  const ss = field(root, ["slides", "スライド"]);
  return Array.isArray(ss) && ss.some(s => isObj(s) && s.layout === undefined && field(s, ["type", "種類"]) !== undefined);
}

function yesNo(v: Json | undefined, def: boolean): boolean {
  if (v === undefined) return def;
  if (typeof v === "boolean") return v;
  const s = String(v).trim().toLowerCase();
  if (["はい", "あり", "有", "表示", "する", "on", "true", "yes", "1", "○", "〇"].indexOf(s) >= 0) return true;
  if (["いいえ", "なし", "無", "非表示", "しない", "off", "false", "no", "0", "×", "-"].indexOf(s) >= 0) return false;
  return def;
}

// Excel の日付シリアル値も yyyy/MM/dd にする
function dateText(v: Json | undefined): string {
  if (typeof v === "number" && v > 20000 && v < 80000) {
    const d = new Date(Math.round((v - 25569) * 86400000));
    return `${d.getUTCFullYear()}/${String(d.getUTCMonth() + 1).padStart(2, "0")}/${String(d.getUTCDate()).padStart(2, "0")}`;
  }
  return cellText(v);
}

function linesOf(v: Json | undefined): string[] {
  const s = Array.isArray(v) ? v.map(x => cellText(x as Json)).join("\n") : v === undefined || v === null ? "" : typeof v === "string" ? v : cellText(v);
  return s.replace(/\r\n?/g, "\n").split("\n").map(l => l.replace(/\s+$/, ""));
}

const MARKER = /^(?:[・•●◆■□◇▪]\s*|[-*]\s+)/;
const INDENT = /^(?: {2,}|\t|　|>)/;

// 行頭が空白 (半角 2 つ / 全角 / タブ) または「>」の行を 2 段目にする。行頭の「・」「- 」などは取り除く
function bulletTree(lines: string[]): Json[] {
  const out: Json[] = [];
  for (const raw of lines) {
    if (raw.trim() === "") continue;
    const sub = INDENT.test(raw);
    const text = raw.replace(/^[\s　>]+/, "").replace(MARKER, "");
    const last = out[out.length - 1];
    if (sub && out.length > 0) {
      if (Array.isArray(last)) last.push(text); else out.push([text]);
    } else out.push(text);
  }
  return out;
}

function splitFields(s: string): string[] {
  return s.split(/\s*[|｜\t]\s*/).map(t => t.trim());
}

// 「見出し | 説明」の 1 行 = 1 項目。行頭が「・」「- 」や字下げの行は直前の項目の説明に追加する。
// 「---」だけの行があれば、それで区切ったまとまり (1 行目 = 見出し) を 1 項目とする
function parseItems(lines: string[]): SItem[] {
  const items: SItem[] = [];
  const sep = /^\s*[-ー─]{3,}\s*$/;
  if (lines.some(l => sep.test(l))) {
    let cur: SItem | null = null;
    for (const l of lines) {
      if (sep.test(l)) { cur = null; continue; }
      if (l.trim() === "") continue;
      if (!cur) {
        const f = splitFields(l.replace(MARKER, ""));
        cur = { head: f[0], fields: f.slice(1).filter(x => x !== ""), body: [] };
        items.push(cur);
      } else cur.body.push(l.trim().replace(MARKER, ""));
    }
    return items;
  }
  for (const l of lines) {
    if (l.trim() === "") continue;
    if ((MARKER.test(l.trim()) || INDENT.test(l)) && items.length > 0) {
      items[items.length - 1].body.push(l.replace(/^[\s　>]+/, "").replace(MARKER, ""));
      continue;
    }
    const f = splitFields(l);
    items.push({ head: f[0], fields: f.slice(1).filter(x => x !== ""), body: [] });
  }
  return items;
}

function descOf(it: SItem): string[] {
  return it.fields.concat(it.body);
}

function toNumber(s: string): number | null {
  const t = s.replace(/[,，\s円人件%％]/g, "");
  return /^-?\d+(\.\d+)?$/.test(t) ? +t : null;
}

function simpleToDeck(root: JsonObj, warnings: string[], frame?: SimpleFrame): JsonObj {
  const get = (names: string[]): string => cellText(field(root, names));
  // 社内テンプレート指定時は、背景・飾り・色・フォントをテンプレート (スライドマスター) に任せ、
  // デザインはカードや表などの図形の形だけに使う
  const template = get(["template", "テンプレート"]);
  const tpl = template !== "" || frame !== undefined;
  const fr = frame ?? DEFAULT_FRAME;
  SW = fr.sw; SH = fr.sh; CL = fr.cl; CW = fr.cw; CT = fr.ct; CH = fr.ch;
  const presets = designPresets();
  const names = Object.keys(presets);
  let designName = get(["design", "デザイン"]) || "ビジネス";
  if (!presets[designName]) {
    const hit = names.filter(n => norm(n) === norm(designName))[0];
    if (!hit) warnings.push(`デザイン「${designName}」はありません。ビジネスで作成しました (選べるデザイン: ${names.join(" / ")})`);
    designName = hit ?? "ビジネス";
  }
  const d = tpl ? Object.assign({}, presets[designName], { cover: {}, section: {}, closing: {} }) : presets[designName];
  const colors: { [k: string]: string } = tpl ? {} : Object.assign({}, d.colors);
  const setColor = (key: string, v: string, label: string): void => {
    if (!v) return;
    try { colors[key] = hexOf(v); } catch (e) { warnings.push(`${label}「${v}」は色として読み取れません (例: #2563EB)`); }
  };
  setColor("accent1", get(["color", "mainColor", "メインカラー"]), "メインカラー");
  setColor("accent2", get(["subColor", "サブカラー"]), "サブカラー");
  const font = get(["font", "フォント"]);
  const author = get(["author", "作成者"]);
  const date = dateText(field(root, ["date", "日付"]));
  const logo = get(["logo", "ロゴ"]);
  const files: string[] = [];
  const useFile = (p: string): string => { if (files.indexOf(p) < 0) files.push(p); return p; };

  const master = tpl ? {} as JsonObj : cloneJson(d.master) as JsonObj;
  if (logo) master.logo = { image: useFile(logo), position: "topRight", h: 1.0 };
  const rows = (field(root, ["slides", "スライド"]) ?? []) as Json[];
  const slidesIn = (Array.isArray(rows) ? rows : []).filter(isObj);
  const typeOf = (s: JsonObj): string => {
    const t = cellText(field(s, ["type", "種類"]));
    if (t === "") return "bullets";
    const key = t.replace(/[\s　]/g, "");
    return SIMPLE_TYPES[key] ?? SIMPLE_TYPES[key.toLowerCase()] ?? "";
  };
  const chapters = slidesIn.filter(s => typeOf(s) === "section").map(s => cellText(field(s, ["title", "タイトル"])));
  const firstTitle = slidesIn.filter(s => typeOf(s) === "cover").map(s => cellText(field(s, ["title", "タイトル"])))[0] ?? "";

  const out: Json[] = [];
  let chapter = 0;
  slidesIn.forEach((s, i) => {
    if (s.layout !== undefined) { out.push(s); return; }   // 詳細モードのスライドはそのまま使う
    let type = typeOf(s);
    if (type === "") {
      warnings.push(`${i + 1} 行目: 種類「${cellText(field(s, ["type", "種類"]))}」は使えないため箇条書きにしました`);
      type = "bullets";
    }
    const title = cellText(field(s, ["title", "タイトル"]));
    const lines = linesOf(field(s, ["content", "内容", "body"]));
    const image = cellText(field(s, ["image", "画像"]));
    let slide: JsonObj;
    try {
      if (type === "section") chapter++;
      slide = simpleSlide(type, title, lines, image ? useFile(image) : "", d, { author, date, chapters, chapter, tpl });
    } catch (e) {
      throw new Error(`${i + 1} 行目 (${title || type}): ${(e as Error).message}`);
    }
    if (frame && frame.layouts) slide = adaptToTemplate(slide, frame, String(chapter).padStart(2, "0"));
    const notes = linesOf(field(s, ["notes", "ノート", "メモ"])).join("\n").trim();
    if (notes) slide.notes = notes;
    out.push(slide);
  });

  const trans = get(["transition", "切り替え", "画面切り替え"]).toLowerCase();
  const transMap: { [k: string]: string } = { "フェード": "fade", "プッシュ": "push", "ワイプ": "wipe", "カバー": "cover", "fade": "fade", "push": "push", "wipe": "wipe" };
  const fileName = get(["fileName", "ファイル名"]) || `${(firstTitle || "プレゼンテーション").replace(/[\\/:*?"<>|\r\n]/g, "")}_{timestamp}`;
  const fonts: JsonObj = tpl ? (font ? { majorEa: font, minorEa: font } : {}) : { majorEa: font || d.titleFont || d.font, minorEa: font || d.font };
  const deck: JsonObj = {
    size: "16:9", unit: "cm",
    properties: { title: firstTitle, author },
    theme: tpl ? { colors, fonts } : { name: designName, colors, fonts },
    master,
    footer: get(["footer", "フッター"]),
    slideNumber: yesNo(field(root, ["slideNumber", "ページ番号"]), true),
    output: { folder: get(["folder", "保存先"]) || "/PowerPoint/output", fileName },
    assets: files.map(p => ({ name: p, path: p })),
    slides: out,
  };
  if (transMap[trans]) deck.transition = transMap[trans];
  if (template) deck.template = template;
  if (out.length === 0) warnings.push("スライドが 1 行もありません");
  return deck;
}

interface SimpleCtx { author: string; date: string; chapters: string[]; chapter: number; tpl: boolean }

function simpleSlide(type: string, title: string, lines: string[], image: string, d: DesignDef, c: SimpleCtx): JsonObj {
  const text = lines.filter(l => l.trim() !== "");
  const colored = (t: string, color?: string): Json => color ? { text: t, color } : t;
  const shape = (x: number, y: number, w: number, h: number, extra: JsonObj): JsonObj =>
    Object.assign({ type: "shape", x: r2(x), y: r2(y), w: r2(w), h: r2(h) } as JsonObj, extra);
  const card = (x: number, y: number, w: number, h: number, extra: JsonObj): JsonObj =>
    shape(x, y, w, h, Object.assign({}, d.card, { color: "tx1", align: "left", valign: "top", margin: [0.5, 0.45, 0.5, 0.4] }, extra));
  // 左に箇条書き、右に画像。組み込みデザインは専用レイアウト、テンプレートは「タイトルのみ」に配置する
  const imageRight = (): JsonObj => {
    if (!c.tpl) return { layout: "imageRight", title, body: bulletTree(lines), picture: image };
    const tw = r2(CW * 0.5);
    return { layout: "titleOnly", title, elements: [
      { type: "text", text: bulletTree(lines), x: CL, y: CT, w: tw, h: CH, size: 18, color: "tx1", bullet: true, spaceAfter: 6, valign: "top" },
      { type: "image", image, x: r2(CL + tw + 0.8), y: CT, w: r2(CW - tw - 0.8), h: CH, fit: "contain" },
    ] };
  };

  switch (type) {
    case "cover": {
      const sub = text.length ? text.join("\n") : [c.author, c.date].filter(x => x).join(" ｜ ");
      return { layout: "title", title: colored(title, d.cover.title), subtitle: colored(sub, d.cover.sub) };
    }
    case "section": {
      const s: JsonObj = { layout: "section", title: colored(title, d.section.title), section: title };
      if (text.length) s.subtitle = colored(text.join("\n"), d.section.sub);
      if (!c.tpl) s.elements = [{ type: "text", text: String(c.chapter).padStart(2, "0"), x: CL, y: 2.6, w: 10, h: 3, size: 54, bold: true, color: d.section.num ?? "lt1", autofit: "none" }];
      return s;
    }
    case "closing": {
      const s: JsonObj = { layout: c.tpl ? "title" : "closing", title: colored(title || "ご清聴ありがとうございました", d.closing.title) };
      if (text.length) s.subtitle = colored(text.join("\n"), d.closing.sub);
      return s;
    }
    case "bullets": {
      if (image) return imageRight();
      return { layout: "content", title, body: bulletTree(lines) };
    }
    case "image": {
      if (!image) throw new Error("「画像」列に画像ファイルのパスを入力してください");
      if (text.length) return imageRight();
      return { layout: "titleOnly", title, elements: [{ type: "image", image, x: CL, y: CT, w: CW, h: CH, fit: "contain" }] };
    }
    case "quote": {
      const author = text.filter(l => /^[—―\-ー]/.test(l.trim())).map(l => l.trim())[0] ?? (title ? `— ${title}` : "");
      const q = text.filter(l => !/^[—―\-ー]/.test(l.trim())).join("\n");
      if (!c.tpl) return { layout: "quote", quote: q || title, author };
      // テンプレートには引用用レイアウトがないため、白紙レイアウトに文字を置く
      const top = SH * 0.2, qh = SH * 0.45;
      const els: Json[] = [
        { type: "text", text: "“", x: r2(CL), y: r2(top - 2.2), w: 4, h: 4, size: 96, bold: true, color: "accent1", autofit: "none" },
        { type: "text", text: q || title, x: r2(CL + 2), y: r2(top), w: r2(CW - 4), h: r2(qh), size: 28, color: "tx1", align: "center", valign: "middle" },
      ];
      if (author) els.push({ type: "text", text: author, x: r2(CL + 2), y: r2(top + qh + 0.4), w: r2(CW - 4), h: 1.4, size: 18, color: "tx1+40", align: "right" });
      return { layout: "blank", elements: els };
    }
    case "message": {
      const els: Json[] = [
        { type: "text", text: title, x: CL + 1, y: 4.2, w: CW - 2, h: 5.6, size: 36, bold: true, color: "tx2", align: "center", valign: "b" },
        shape(SW / 2 - 1.5, 10.3, 3, 0.18, { fill: "accent1", line: "none" }),
      ];
      if (text.length) els.push({ type: "text", text: text.join("\n"), x: CL + 2, y: 11.1, w: CW - 4, h: 4.6, size: 18, color: "tx1+30", align: "center" });
      return { layout: "blank", elements: els };
    }
    case "agenda": {
      const items = (text.length ? text.map(l => l.replace(MARKER, "")) : c.chapters).filter(x => x);
      if (items.length === 0) throw new Error("目次にする「章」のスライドがありません。「内容」に項目を 1 行ずつ入力してください");
      const n = Math.min(items.length, 12);
      const cols = n > 6 ? 2 : 1;
      const per = Math.ceil(n / cols);
      const colW = (CW - (cols - 1) * 1.2) / cols;
      const rowH = Math.min(2.3, CH / per);
      const box = Math.min(rowH - 0.6, 1.5);
      const els: Json[] = [];
      for (let i = 0; i < n; i++) {
        const x = CL + Math.floor(i / per) * (colW + 1.2);
        const y = CT + (i % per) * rowH + (CH - per * rowH) / 2;
        const num = String(i + 1).padStart(2, "0");
        if (d.numShape === "none") els.push({ type: "text", text: num, x: r2(x), y: r2(y), w: r2(box + 0.4), h: r2(rowH - 0.2), size: 24, bold: true, color: "accent1", valign: "middle", autofit: "none" });
        else els.push(shape(x, y + (rowH - 0.2 - box) / 2, box, box, { shape: d.numShape, fill: "accent1", line: "none", text: num, size: 16, bold: true, radius: 18, margin: 0, wrap: false }));
        els.push({ type: "text", text: items[i], x: r2(x + box + 0.6), y: r2(y), w: r2(colW - box - 0.6), h: r2(rowH - 0.2), size: cols > 1 ? 18 : 20, color: "tx1", valign: "middle" });
        els.push({ type: "line", x: r2(x), y: r2(y + rowH - 0.1), x2: r2(x + colW), y2: r2(y + rowH - 0.1), line: { color: "tx1+80", width: 0.75 } });
      }
      return { layout: "titleOnly", title: title || "目次", elements: els };
    }
    case "compare": {
      const items = parseItems(lines);
      if (items.length >= 2 && items.length <= 3 && items.some(it => descOf(it).length > 0)) {
        if (items.length === 2) {
          return { layout: "comparison", title, leftHeading: items[0].head, left: descOf(items[0]), rightHeading: items[1].head, right: descOf(items[1]) };
        }
        if (c.tpl) return simpleSlide("cards", title, lines, image, d, c);
        const col = (it: SItem): Json => ({ text: [{ text: it.head, bold: true, bullet: false, color: "accent1", size: 22 } as Json].concat(descOf(it)) });
        return { layout: "threeContent", title, left: col(items[0]), center: col(items[1]), right: col(items[2]) };
      }
      if (items.length > 3) return simpleSlide("cards", title, lines, image, d, c);
      const half = Math.ceil(text.length / 2);
      return { layout: "twoContent", title, left: bulletTree(text.slice(0, half)), right: bulletTree(text.slice(half)) };
    }
    case "cards": {
      const items = parseItems(lines).slice(0, 8);
      if (items.length === 0) throw new Error("「内容」に「見出し | 説明」を 1 行ずつ入力してください");
      const n = items.length;
      const cols = n <= 4 ? n : Math.ceil(n / 2), rows = n <= 4 ? 1 : 2, gap = 0.6;
      const w = (CW - gap * (cols - 1)) / cols;
      const need = Math.max(...items.map(it => textHeightCm([[it.head, cols > 3 ? 17 : 20] as [string, number]].concat(descOf(it).map(t => [t, cols > 3 ? 13 : 15] as [string, number])), w - 1.0)));
      const h = rows === 1 ? Math.min(9.6, Math.max(5.0, need + 1.7)) : (CH - gap) / 2;
      const y0 = CT + (rows === 1 ? Math.max(0, (CH - h) / 2 - 0.4) : 0);
      const els: Json[] = [];
      items.forEach((it, i) => {
        const x = CL + (i % cols) * (w + gap), y = y0 + Math.floor(i / cols) * (h + gap);
        const desc = descOf(it);
        const body: Json[] = [{ text: it.head, size: cols > 3 ? 17 : 20, bold: true, color: "accent1", spaceAfter: 6 }];
        for (const t of desc) body.push({ text: t, size: cols > 3 ? 13 : 15, color: "tx1", bullet: desc.length > 1, spaceAfter: 3 });
        els.push(card(x, y, w, h, { text: body, margin: [0.5, 0.95, 0.5, 0.4] }));
        els.push(shape(x + 0.5, y + 0.5, 1.4, 0.14, { fill: "accent1", line: "none" }));
      });
      return { layout: "titleOnly", title, elements: els };
    }
    case "kpi": {
      const items = parseItems(lines).slice(0, 6);
      if (items.length === 0) throw new Error("「内容」に「数値 | 説明」を 1 行ずつ入力してください (例: 48.2億円 | 売上高 | 前年比 +12%)");
      const n = items.length;
      const cols = n <= 4 ? n : 3, rows = n <= 4 ? 1 : 2, gap = 0.6;
      const w = (CW - gap * (cols - 1)) / cols;
      const h = rows === 1 ? 6.8 : (CH - gap) / 2;
      const y0 = CT + (rows === 1 ? (CH - h) / 2 - 0.4 : 0);
      const els: Json[] = items.map((it, i) => {
        const x = CL + (i % cols) * (w + gap), y = y0 + Math.floor(i / cols) * (h + gap);
        const body: Json[] = [{ text: it.head, size: cols >= 4 ? 34 : 40, bold: true, color: "accent1" }];
        if (it.fields[0]) body.push({ text: it.fields[0], size: 16, bold: true, color: "tx1", spaceBefore: 4 });
        for (const t of it.fields.slice(1).concat(it.body)) body.push({ text: t, size: 12, color: "tx1+40", spaceBefore: 2 });
        return card(x, y, w, h, { text: body, align: "center", valign: "middle", margin: 0.4 });
      });
      return { layout: "titleOnly", title, elements: els };
    }
    case "steps": {
      const items = parseItems(lines).slice(0, 6);
      if (items.length === 0) throw new Error("「内容」に「ステップ名 | 説明」を 1 行ずつ入力してください");
      const n = items.length;
      const stepW = CW / n;
      const size = n > 4 ? 13 : 15;
      const need = Math.max(...items.map(it => textHeightCm(descOf(it).map(t => [t, size] as [string, number]), stepW - 1.5)));
      const dh = Math.min(CH - 2.9, Math.max(3.5, need + 1.0));
      const top = CT + Math.max(0, (CH - 2.9 - dh) / 2 - 0.4);
      const els: Json[] = [];
      items.forEach((it, i) => {
        const x = CL + i * stepW;
        els.push(shape(x, top, stepW - 0.12, 2.4, {
          shape: i === 0 ? "homePlate" : "chevron", fill: `accent1-${Math.min(50, i * 9)}`, line: "none",
          text: [{ text: `STEP ${i + 1}`, size: 11, bold: true }, { text: it.head, size: n > 4 ? 14 : 16, bold: true }], margin: [0.6, 0.1, 0.3, 0.1],
        }));
        const desc = descOf(it);
        if (desc.length) {
          els.push(card(x, top + 2.9, stepW - 0.4, dh, {
            text: desc.map(t => ({ text: t, bullet: desc.length > 1, size, spaceAfter: 4 } as Json)),
          }));
        }
      });
      return { layout: "titleOnly", title, elements: els };
    }
    case "timeline": {
      const items = parseItems(lines).slice(0, 7);
      if (items.length === 0) throw new Error("「内容」に「時期 | 出来事」を 1 行ずつ入力してください (例: 10月 | 新プラン発表)");
      const n = items.length;
      const colW = CW / n, size = n > 5 ? 13 : 15;
      const need = Math.max(...items.map(it => textHeightCm(descOf(it).map(t => [t, size] as [string, number]), colW - 1.1)));
      const dh = Math.min(CH - 4.9, Math.max(2.5, need + 0.9));
      const lineY = CT + 3.6 + Math.max(0, (CH - 4.8 - dh) / 2 - 0.4);
      const els: Json[] = [{ type: "line", x: CL, y: r2(lineY), x2: CL + CW, y2: r2(lineY), line: { color: "tx1+70", width: 2.5 }, tail: "arrow" }];
      items.forEach((it, i) => {
        const cx = CL + (i + 0.5) * colW;
        els.push({ type: "text", text: it.head, x: r2(cx - colW / 2), y: r2(lineY - 2.8), w: r2(colW), h: 1.8, size: n > 5 ? 16 : 18, bold: true, color: "accent1", align: "center", valign: "bottom" });
        els.push(shape(cx - 0.4, lineY - 0.4, 0.8, 0.8, { shape: "ellipse", fill: "accent1", line: { color: "lt1", width: 2 } }));
        const desc = descOf(it);
        if (desc.length) {
          els.push(card(cx - colW / 2 + 0.2, lineY + 1.2, colW - 0.4, dh, {
            text: desc.map(t => ({ text: t, size, bullet: desc.length > 1, spaceAfter: 3 } as Json)), margin: 0.35,
          }));
        }
      });
      return { layout: "titleOnly", title, elements: els };
    }
    case "table": {
      const rows: Json[] = text.map(l => splitFields(l).map(cell => {
        const num = /^-?[1-9][\d,]*(\.\d+)?$|^-?0(\.\d+)?$/.test(cell) ? +cell.replace(/,/g, "") : null;
        return num === null ? cell : num;
      }));
      if (rows.length === 0) throw new Error("「内容」に表の行を「列1 | 列2 | …」の形で入力してください (1 行目は見出し)");
      const tbl: JsonObj = { type: "table", rows, tableStyle: d.table };
      if (rows.length > 11) tbl.rowsPerSlide = 10;
      return { layout: "content", title, body: tbl };
    }
    default: {
      if (type.indexOf("chart:") === 0) return chartSlide(type.slice(6), title, lines);
      throw new Error(`種類「${type}」は未対応です`);
    }
  }
}

function chartSlide(kind: string, title: string, lines: string[]): JsonObj {
  const sep = lines.map(l => /^\s*[-ー─]{3,}\s*$/.test(l)).indexOf(true);
  const dataLines = (sep >= 0 ? lines.slice(0, sep) : lines).filter(l => l.trim() !== "");
  const notes = sep >= 0 ? lines.slice(sep + 1) : [];
  const rows = dataLines.map(splitFields);
  if (rows.length === 0) throw new Error("「内容」にグラフのデータを「項目 | 値」の形で 1 行ずつ入力してください");
  const header = rows.length > 1 && rows[0].slice(1).some(x => x !== "" && toNumber(x) === null) ? rows.shift() as string[] : null;
  const nSeries = Math.max(1, ...rows.map(r => r.length - 1));
  const isPie = kind === "pie" || kind === "doughnut";
  const series: Json[] = range(isPie ? 1 : nSeries).map(j => ({
    name: header?.[j + 1] || (nSeries === 1 ? title || "値" : `系列${j + 1}`),
    values: rows.map(r => toNumber(r[j + 1] ?? "")),
  }));
  const chart: JsonObj = {
    type: "chart", chartType: kind === "stacked" ? "column" : kind, categories: rows.map(r => r[0]), series,
    dataLabels: isPie ? "percent" : nSeries === 1, legend: isPie || nSeries > 1 ? "bottom" : false,
  };
  if (kind === "stacked") chart.stacked = true;
  if (kind === "doughnut") chart.hole = 55;
  const text = notes.filter(l => l.trim() !== "");
  if (text.length) return { layout: "twoContent", title, left: chart, right: bulletTree(text) };
  return { layout: "content", title, body: chart };
}

// 文字列 [テキスト, pt] を幅 w (cm) に流し込んだときのおおよその高さ (cm)
function textHeightCm(paras: [string, number][], w: number): number {
  let h = 0;
  for (const [t, size] of paras) {
    const em = size * 0.0353;
    const perLine = Math.max(1, Math.floor(w / em));
    h += Math.max(1, Math.ceil(textWidth(t) / 2 / perLine)) * em * 1.3 + 0.12;
  }
  return h;
}

function r2(v: number): number {
  return Math.round(v * 100) / 100;
}

// =====================================================================
// 構成 JSON の前処理 (変数 {{var}}・each による繰り返し・表の自動改ページ)
// =====================================================================
function isObj(v: Json | undefined): v is JsonObj {
  return v !== null && v !== undefined && typeof v === "object" && !Array.isArray(v);
}

function cloneJson(v: Json): Json {
  return JSON.parse(JSON.stringify(v)) as Json;
}

function expandDeck(root: JsonObj): JsonObj {
  const vars: JsonObj = isObj(root.vars) ? root.vars : {};
  const slides = Array.isArray(root.slides) ? root.slides : [];
  const out: Json[] = [];
  for (const s of slides) {
    if (!isObj(s)) continue;
    const eachVal = s.each === undefined ? undefined : substitute(s.each, vars);
    if (Array.isArray(eachVal)) {
      eachVal.forEach((row, i) => {
        const local: JsonObj = Object.assign({}, vars, isObj(row) ? row : { value: row }, { index: i + 1 });
        const copy = cloneJson(s) as JsonObj;
        delete copy.each;
        for (const p of paginate(substitute(copy, local) as JsonObj)) out.push(p);
      });
    } else {
      for (const p of paginate(substitute(s, vars) as JsonObj)) out.push(p);
    }
  }
  const rest = Object.assign({}, root);
  delete rest.slides;
  const result = substitute(rest, vars) as JsonObj;
  result.slides = out;
  return result;
}

function lookupVar(vars: JsonObj, path: string): Json | undefined {
  let cur: Json | undefined = vars;
  for (const seg of path.split(".")) {
    if (Array.isArray(cur) && /^\d+$/.test(seg)) cur = cur[+seg];
    else if (isObj(cur)) cur = cur[seg];
    else return undefined;
    if (cur === undefined) return undefined;
  }
  return cur;
}

function substitute(v: Json, vars: JsonObj): Json {
  if (typeof v === "string") {
    const whole = /^\{\{\s*([\w.\-]+)\s*\}\}$/.exec(v);
    if (whole) {
      const r = lookupVar(vars, whole[1]);
      return r === undefined ? v : cloneJson(r);
    }
    return v.replace(/\{\{\s*([\w.\-]+)\s*\}\}/g, (m: string, k: string) => {
      const r = lookupVar(vars, k);
      if (r === undefined) return m;
      return typeof r === "string" ? r : typeof r === "number" || typeof r === "boolean" ? String(r) : JSON.stringify(r);
    });
  }
  if (Array.isArray(v)) return v.map(x => substitute(x, vars));
  if (isObj(v)) {
    const o: JsonObj = {};
    for (const k of Object.keys(v)) o[k] = substitute(v[k], vars);
    return o;
  }
  return v;
}

function paginate(slide: JsonObj): JsonObj[] {
  let table: JsonObj | null = null;
  let holderKey = "";
  let elIndex = -1;
  for (const k of Object.keys(slide)) {
    const v = slide[k];
    if (isObj(v) && typeof v.rowsPerSlide === "number") { table = v; holderKey = k; break; }
  }
  const els = slide.elements;
  if (!table && Array.isArray(els)) {
    for (let i = 0; i < els.length; i++) {
      const e = els[i];
      if (isObj(e) && typeof e.rowsPerSlide === "number") { table = e; elIndex = i; break; }
    }
  }
  if (!table) return [slide];
  const per = Math.max(1, Math.floor(table.rowsPerSlide as number));
  const useData = Array.isArray(table.data);
  const all = (useData ? table.data : table.rows) as Json[];
  if (!Array.isArray(all)) return [slide];
  const header = !useData && table.header !== false ? all.slice(0, 1) : [];
  const body = !useData && table.header !== false ? all.slice(1) : all;
  const suffix = typeof table.continuedSuffix === "string" ? table.continuedSuffix : "（続き）";
  const pages: JsonObj[] = [];
  for (let i = 0; i < Math.max(1, body.length); i += per) {
    const copy = cloneJson(slide) as JsonObj;
    const t = (elIndex >= 0 ? (copy.elements as Json[])[elIndex] : copy[holderKey]) as JsonObj;
    delete t.rowsPerSlide;
    delete t.continuedSuffix;
    const chunk = body.slice(i, i + per);
    if (useData) t.data = chunk; else t.rows = header.concat(chunk);
    if (i > 0 && typeof copy.title === "string") copy.title = copy.title + suffix;
    pages.push(copy);
  }
  return pages;
}

// =====================================================================
// 組み込みスライドマスター／レイアウト
// =====================================================================
interface PhDef { name: string; type: string; idx: number; rect: number[]; st: StyleSpec; bullets?: boolean }
interface LayoutDef {
  key: string; name: string; type?: string; bg?: FillSpec; showMaster?: boolean;
  phs: PhDef[]; elements?: ElementSpec[];
}

function pct(v: number): string {
  return `${Math.round(v * 100000) / 1000}%`;
}

function builtinGeometry(W: number, H: number): { [k: string]: number[] } {
  const mx = Math.round(W * 0.055);
  const cw = W - mx * 2;
  return {
    title: [mx, Math.round(H * 0.05), cw, Math.round(H * 0.13)],
    body: [mx, Math.round(H * 0.24), cw, Math.round(H * 0.64)],
    dt: [mx, Math.round(H * 0.925), Math.round(W * 0.25), Math.round(H * 0.05)],
    ftr: [Math.round(W * 0.3), Math.round(H * 0.925), Math.round(W * 0.4), Math.round(H * 0.05)],
    sldNum: [W - mx - Math.round(W * 0.15), Math.round(H * 0.925), Math.round(W * 0.15), Math.round(H * 0.05)],
    misc: [mx, cw, Math.round(W * 0.025)],
  };
}

function builtinLayouts(deck: DeckCtx): LayoutDef[] {
  const W = deck.W, H = deck.H;
  const g = builtinGeometry(W, H);
  const [mx, cw, gap] = g.misc;
  const T = g.title, B = g.body;
  const half = Math.round((cw - gap) / 2);
  const third = Math.round((cw - gap * 2) / 3);
  const r = (x: number, y: number, w: number, h: number): number[] => [Math.round(x), Math.round(y), Math.round(w), Math.round(h)];
  const title: PhDef = { name: "Title", type: "title", idx: 0, rect: T, st: {} };
  const noBullet = (name: string, idx: number, rect: number[], st: StyleSpec): PhDef => ({ name, type: "body", idx, rect, st, bullets: false });
  const obj = (name: string, idx: number, rect: number[]): PhDef => ({ name, type: "obj", idx, rect, st: {} });
  const pic = (idx: number, rect: number[]): PhDef => ({ name: "Picture", type: "pic", idx, rect, st: {} });
  const defs: LayoutDef[] = [
    {
      key: "title", name: "タイトル スライド", type: "title", bg: "dk2", showMaster: false,
      phs: [
        { name: "Title", type: "ctrTitle", idx: 0, rect: r(mx + W * 0.02, H * 0.26, W * 0.82, H * 0.26), st: { size: 40, bold: true, color: "lt1", valign: "b" } },
        { name: "Subtitle", type: "subTitle", idx: 1, rect: r(mx + W * 0.02, H * 0.58, W * 0.82, H * 0.16), st: { size: 20, color: "lt2" }, bullets: false },
      ],
      elements: [{ type: "shape", name: "Accent Bar", x: pct((mx + W * 0.02) / W), y: "54.5%", w: "8%", h: "0.8%", fill: "accent1", line: "none" }],
    },
    {
      key: "section", name: "セクション見出し", type: "secHead", bg: "accent1", showMaster: false,
      phs: [
        { name: "Title", type: "title", idx: 0, rect: r(mx, H * 0.3, W * 0.85, H * 0.24), st: { size: 36, bold: true, color: "lt1", valign: "b" } },
        noBullet("Subtitle", 1, r(mx, H * 0.58, W * 0.85, H * 0.14), { size: 18, color: "lt1" }),
      ],
      elements: [{ type: "shape", name: "Accent Bar", x: pct(mx / W), y: "55.5%", w: "6%", h: "0.6%", fill: "lt1", line: "none" }],
    },
    { key: "content", name: "タイトルとコンテンツ", type: "obj", phs: [title, obj("Body", 1, B)] },
    {
      key: "twoContent", name: "2 つのコンテンツ", type: "twoObj",
      phs: [title, obj("Left", 1, r(mx, B[1], half, B[3])), obj("Right", 2, r(mx + half + gap, B[1], half, B[3]))],
    },
    {
      key: "comparison", name: "比較", type: "twoTxTwoObj",
      phs: [
        title,
        noBullet("Left Heading", 1, r(mx, H * 0.24, half, H * 0.08), { size: 20, bold: true, color: "accent1", valign: "b" }),
        obj("Left", 2, r(mx, H * 0.33, half, H * 0.55)),
        noBullet("Right Heading", 3, r(mx + half + gap, H * 0.24, half, H * 0.08), { size: 20, bold: true, color: "accent1", valign: "b" }),
        obj("Right", 4, r(mx + half + gap, H * 0.33, half, H * 0.55)),
      ],
    },
    {
      key: "threeContent", name: "3 つのコンテンツ",
      phs: [title, obj("Left", 1, r(mx, B[1], third, B[3])), obj("Center", 2, r(mx + third + gap, B[1], third, B[3])),
        obj("Right", 3, r(mx + (third + gap) * 2, B[1], third, B[3]))],
    },
    { key: "titleOnly", name: "タイトルのみ", type: "titleOnly", phs: [title] },
    { key: "blank", name: "白紙", type: "blank", phs: [] },
    {
      key: "contentCaption", name: "タイトル付きのコンテンツ", type: "objTx",
      phs: [title, obj("Body", 1, r(mx + cw * 0.36 + gap, B[1], cw * 0.64 - gap, B[3])),
        noBullet("Caption", 2, r(mx, B[1], cw * 0.36, B[3]), { size: 16 })],
    },
    {
      key: "pictureCaption", name: "タイトル付きの図", type: "picTx",
      phs: [title, pic(1, r(mx, B[1], cw * 0.6, B[3])), noBullet("Caption", 2, r(mx + cw * 0.6 + gap, B[1], cw * 0.4 - gap, B[3]), { size: 16 })],
    },
    { key: "imageRight", name: "画像 (右)", phs: [title, obj("Body", 1, r(mx, B[1], half, B[3])), pic(2, r(mx + half + gap, B[1], half, B[3]))] },
    { key: "imageLeft", name: "画像 (左)", phs: [title, pic(2, r(mx, B[1], half, B[3])), obj("Body", 1, r(mx + half + gap, B[1], half, B[3]))] },
    {
      key: "quote", name: "引用", bg: "lt2", showMaster: false,
      phs: [
        noBullet("Quote", 1, r(W * 0.12, H * 0.22, W * 0.76, H * 0.42), { size: 28, align: "center", valign: "ctr" }),
        noBullet("Author", 2, r(W * 0.12, H * 0.68, W * 0.76, H * 0.08), { size: 16, align: "right", color: "accent1" }),
      ],
      elements: [{ type: "text", name: "Quote Mark", text: "“", x: "7%", y: "6%", w: "12%", h: "22%", size: 120, color: "accent1", autofit: "none" }],
    },
    {
      key: "closing", name: "クロージング", bg: "dk2", showMaster: false,
      phs: [
        { name: "Title", type: "title", idx: 0, rect: r(mx, H * 0.32, cw, H * 0.22), st: { size: 40, bold: true, color: "lt1", align: "center", valign: "ctr" } },
        noBullet("Subtitle", 1, r(mx, H * 0.56, cw, H * 0.12), { size: 18, color: "lt2", align: "center" }),
      ],
    },
  ];
  // 構成 JSON の master.layouts でレイアウトを追加 / 同じ key の組み込みレイアウトを置き換え
  for (const c of deck.spec.master?.layouts ?? []) {
    let idx = 1;
    const phs: PhDef[] = (c.placeholders ?? []).map(p => {
      const t = (p.type ?? "content").toLowerCase();
      const type = t === "title" ? "title" : t === "picture" || t === "image" ? "pic" : t === "content" ? "obj" : "body";
      const rect = [lenEmu(deck, p.x, "x") ?? 0, lenEmu(deck, p.y, "y") ?? 0, lenEmu(deck, p.w, "x") ?? W, lenEmu(deck, p.h, "y") ?? H];
      const bullets = p.bullets ?? (type === "obj" ? undefined : false);
      return { name: p.name, type, idx: type === "title" ? 0 : idx++, rect, st: p, bullets };
    });
    const def: LayoutDef = { key: c.key, name: c.name ?? c.key, type: c.type, bg: c.background, showMaster: c.showMaster, phs, elements: c.elements };
    const at = defs.map(d => norm(d.key)).indexOf(norm(c.key));
    if (at >= 0) defs[at] = def; else defs.push(def);
  }
  return defs;
}

function createBuiltinPackage(deck: DeckCtx): void {
  const pkg = deck.pkg;
  const spec = deck.spec;
  const ms = spec.master ?? {};
  const W = deck.W, H = deck.H;
  const g = builtinGeometry(W, H);
  const sc = deck.scale;

  setText(pkg, "ppt/theme/theme1.xml", themeXml(spec.theme), "application/vnd.openxmlformats-officedocument.theme+xml");
  setText(pkg, "ppt/presProps.xml", XML_HEAD + `<p:presentationPr ${NS}/>`, `${CT_P}.presProps+xml`);
  setText(pkg, "ppt/viewProps.xml", XML_HEAD + `<p:viewPr ${NS}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>`, `${CT_P}.viewProps+xml`);
  setText(pkg, "ppt/tableStyles.xml", XML_HEAD + `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`, `${CT_P}.tableStyles+xml`);

  // --- スライドマスター ---
  const master = "ppt/slideMasters/slideMaster1.xml";
  const ts = Object.assign({ size: 32, bold: true, color: "tx2" }, ms.titleStyle ?? {});
  const bs = Object.assign({ size: 20, color: "tx1" }, ms.bodyStyle ?? {});
  const bullets = ms.bullets ?? ["•", "–", "▪", "–", "▪"];
  const bulletColor = bs.bulletColor ?? "accent1";
  const phSp = (id: number, name: string, ph: string, rect: number[], body: string, lst: string, bodyPr: string): string =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${rect[0]}" y="${rect[1]}"/><a:ext cx="${rect[2]}" cy="${rect[3]}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody>${bodyPr}<a:lstStyle>${lst}</a:lstStyle>${body}</p:txBody></p:sp>`;
  const sample = (t: string, lvl = 0): string => `<a:p>${lvl ? `<a:pPr lvl="${lvl}"/>` : ""}<a:r><a:rPr lang="ja-JP" altLang="en-US"/><a:t>${t}</a:t></a:r></a:p>`;
  const gray = `<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="50000"/><a:lumOff val="50000"/></a:schemeClr></a:solidFill>`;
  const footLst = (algn: string): string => `<a:lvl1pPr algn="${algn}"><a:defRPr sz="${Math.round(1000 * sc)}">${gray}</a:defRPr></a:lvl1pPr>`;
  const masterPhs =
    phSp(100, "Title Placeholder", `<p:ph type="title"/>`, g.title, sample("マスター タイトルの書式設定"), "",
      `<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr>`) +
    phSp(101, "Text Placeholder", `<p:ph type="body" idx="1"/>`, g.body,
      sample("マスター テキストの書式設定") + sample("第 2 レベル", 1) + sample("第 3 レベル", 2) + sample("第 4 レベル", 3) + sample("第 5 レベル", 4), "",
      `<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr>`) +
    phSp(102, "Date Placeholder", `<p:ph type="dt" sz="half" idx="10"/>`, g.dt, `<a:p><a:endParaRPr lang="ja-JP"/></a:p>`, footLst("l"), `<a:bodyPr anchor="ctr"/>`) +
    phSp(103, "Footer Placeholder", `<p:ph type="ftr" sz="quarter" idx="11"/>`, g.ftr, `<a:p><a:endParaRPr lang="ja-JP"/></a:p>`, footLst("ctr"), `<a:bodyPr anchor="ctr"/>`) +
    phSp(104, "Slide Number Placeholder", `<p:ph type="sldNum" sz="quarter" idx="12"/>`, g.sldNum,
      `<a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="ja-JP"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="ja-JP"/></a:p>`, footLst("r"), `<a:bodyPr anchor="ctr"/>`);

  const lvlPPr = (n: number, st: StyleSpec, bullet: string | null, sz: number): string => {
    const marL = bullet === null ? 0 : Math.round((n * 342900 + 285750) * sc);
    const bu = bullet === null ? `<a:buNone/>` :
      `<a:buClr>${colorXml(bulletColor)}</a:buClr><a:buFont typeface="Arial"/><a:buChar char="${esc(bullet)}"/>`;
    return `<a:lvl${n + 1}pPr marL="${marL}" indent="${bullet === null ? 0 : -Math.round(285750 * sc)}" algn="${alignVal(st.align) ?? "l"}" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">` +
      `<a:lnSpc><a:spcPct val="${Math.round((st.lineSpacing ?? (bullet === null ? 0.9 : 1.1)) * 100000)}"/></a:lnSpc>` +
      `<a:spcBef><a:spcPts val="${Math.round((st.spaceBefore ?? (bullet === null ? 0 : 6)) * 100)}"/></a:spcBef>${bu}` +
      `<a:defRPr sz="${Math.round(sz * 100)}" kern="1200"${st.bold ? ` b="1"` : ""}${st.italic ? ` i="1"` : ""}>` +
      `<a:solidFill>${colorXml(st.color ?? "tx1")}</a:solidFill>${fontXml(st.font, st.fontEa, bullet === null ? "major" : "minor")}</a:defRPr></a:lvl${n + 1}pPr>`;
  };
  const bodySizes = [0, 1, 2, 3, 4].map(n => Math.max(10, ((bs.size as number) - n * 2)) * sc);
  let bodyStyle = "";
  for (let n = 0; n < 5; n++) bodyStyle += lvlPPr(n, bs, bullets[n] ?? bullets[bullets.length - 1], bodySizes[n]);
  let otherStyle = `<a:defPPr><a:defRPr lang="ja-JP"/></a:defPPr>`;
  for (let n = 0; n < 5; n++) {
    otherStyle += `<a:lvl${n + 1}pPr marL="${n * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${n + 1}pPr>`;
  }

  const defs = builtinLayouts(deck);
  const layoutIds = defs.map((d, i) =>
    `<p:sldLayoutId id="${2147483649 + i}" r:id="${addRel(pkg, master, R_LAYOUT, `../slideLayouts/slideLayout${i + 1}.xml`)}"/>`).join("");
  addRel(pkg, master, R_THEME, "../theme/theme1.xml");
  setText(pkg, master, XML_HEAD + `<p:sldMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>` +
    `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    masterPhs + `</p:spTree></p:cSld>` +
    `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
    `<p:sldLayoutIdLst>${layoutIds}</p:sldLayoutIdLst>` +
    `<p:txStyles><p:titleStyle>${lvlPPr(0, ts, null, (ts.size as number) * sc)}</p:titleStyle><p:bodyStyle>${bodyStyle}</p:bodyStyle><p:otherStyle>${otherStyle}</p:otherStyle></p:txStyles></p:sldMaster>`,
    `${CT_P}.slideMaster+xml`);

  // --- スライドレイアウト ---
  defs.forEach((d, i) => {
    const part = `ppt/slideLayouts/slideLayout${i + 1}.xml`;
    addRel(pkg, part, R_MASTER, "../slideMasters/slideMaster1.xml");
    const lctx: PartCtx = { deck, part, nextId: 2 };
    const els = (d.elements ?? []).map(e => renderElement(lctx, e));
    // タイトル下のアクセントバー (タイトルがあり、マスターを表示するレイアウトのみ)
    const tph = d.phs.filter(p => p.type === "title")[0];
    if (tph && d.showMaster !== false && ms.titleBar !== false) {
      els.unshift(renderElement(lctx, {
        type: "shape", name: "Title Bar", x: `${tph.rect[0] + 91440}emu`, y: `${tph.rect[1] + tph.rect[3] + Math.round(H * 0.006)}emu`,
        w: "5%", h: "0.7%", fill: typeof ms.titleBar === "string" ? ms.titleBar : "accent1", line: "none",
      }));
    }
    const phs = d.phs.map(p => layoutPhXml(lctx, p, sc)).join("");
    const foot = [["dt", 10, g.dt, "Date"], ["ftr", 11, g.ftr, "Footer"], ["sldNum", 12, g.sldNum, "Slide Number"]].map(f =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${lctx.nextId++}" name="${f[3]} Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${f[0]}" sz="quarter" idx="${f[1]}"/></p:nvPr></p:nvSpPr><p:spPr/>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/>${f[0] === "sldNum" ? `<a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="ja-JP"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="ja-JP"/></a:p>` : `<a:p><a:endParaRPr lang="ja-JP"/></a:p>`}</p:txBody></p:sp>`).join("");
    const bg = d.bg !== undefined ? bgXml(lctx, d.bg) : "";
    setText(pkg, part, XML_HEAD + `<p:sldLayout ${NS}${d.type ? ` type="${d.type}"` : ""} preserve="1"${d.showMaster === false ? ` showMasterSp="0"` : ""}>` +
      `<p:cSld name="${esc(d.name)}">${bg}${spTreeOpen()}${els.join("")}${phs}${foot}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      `${CT_P}.slideLayout+xml`);
    deck.layoutKeys[part] = d.key;
  });

  // --- presentation.xml ---
  addRel(pkg, deck.presPart, R_MASTER, "slideMasters/slideMaster1.xml");
  addRel(pkg, deck.presPart, `${REL}/presProps`, "presProps.xml");
  addRel(pkg, deck.presPart, `${REL}/viewProps`, "viewProps.xml");
  addRel(pkg, deck.presPart, R_THEME, "theme/theme1.xml");
  addRel(pkg, deck.presPart, `${REL}/tableStyles`, "tableStyles.xml");
  const masterRid = getRels(pkg, deck.presPart)[0].id;
  setText(pkg, deck.presPart, XML_HEAD + `<p:presentation ${NS} saveSubsetFonts="1">` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="${masterRid}"/></p:sldMasterIdLst>` +
    `<p:sldSz cx="${W}" cy="${H}"/><p:notesSz cx="6858000" cy="9144000"/>` +
    `<p:defaultTextStyle>${otherStyle}</p:defaultTextStyle></p:presentation>`, `${CT_P}.presentation.main+xml`);
  addRel(pkg, "", `${REL}/officeDocument`, "ppt/presentation.xml");
}

function layoutPhXml(ctx: PartCtx, d: PhDef, sc: number): string {
  const st = d.st;
  const phAttr = (d.type === "obj" ? "" : ` type="${d.type}"`) + (d.idx ? ` idx="${d.idx}"` : "");
  const hasLvl = st.size !== undefined || st.bold !== undefined || st.color !== undefined || st.align !== undefined ||
    st.italic !== undefined || st.font !== undefined || d.bullets === false;
  const lvl = hasLvl
    ? `<a:lvl1pPr${d.bullets === false ? ` marL="0" indent="0"` : ""}${alignVal(st.align) ? ` algn="${alignVal(st.align)}"` : ""}>${d.bullets === false ? "<a:buNone/>" : ""}` +
    `<a:defRPr${st.size !== undefined ? ` sz="${Math.round(st.size * 100 * sc)}"` : ""}${st.bold !== undefined ? ` b="${st.bold ? 1 : 0}"` : ""}${st.italic ? ` i="1"` : ""}>` +
    `${st.color ? `<a:solidFill>${colorXml(st.color)}</a:solidFill>` : ""}${fontXml(st.font, st.fontEa, "")}</a:defRPr></a:lvl1pPr>`
    : "";
  const anchor = anchorVal(st.valign);
  const prompt = TITLE_TYPES.indexOf(d.type) >= 0 ? "タイトルを入力" : d.type === "pic" ? "図を挿入" : d.type === "subTitle" ? "サブタイトルを入力" : "テキストを入力";
  return `<p:sp><p:nvSpPr><p:cNvPr id="${ctx.nextId++}" name="${esc(d.name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${phAttr}/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${d.rect[0]}" y="${d.rect[1]}"/><a:ext cx="${d.rect[2]}" cy="${d.rect[3]}"/></a:xfrm></p:spPr>` +
    `<p:txBody><a:bodyPr${anchor ? ` anchor="${anchor}"` : ""}><a:normAutofit/></a:bodyPr><a:lstStyle>${lvl}</a:lstStyle><a:p><a:r><a:rPr lang="ja-JP" altLang="en-US"/><a:t>${prompt}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

function themeXml(t?: ThemeSpec): string {
  const c: { [k: string]: string } = Object.assign({}, DEFAULT_COLORS);
  for (const k of Object.keys(t?.colors ?? {})) c[themeColorKey(k)] = hexOf((t?.colors ?? {})[k]);
  const f = t?.fonts ?? {};
  const clr = (n: string): string => `<a:${n}><a:srgbClr val="${c[n]}"/></a:${n}>`;
  const font = (latin: string, ea: string): string =>
    `<a:latin typeface="${esc(latin)}"/><a:ea typeface="${esc(ea)}"/><a:cs typeface=""/><a:font script="Jpan" typeface="${esc(ea)}"/>`;
  const fill = `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`;
  const ln = (w: number): string => `<a:ln w="${w}"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>`;
  return XML_HEAD + `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${esc(t?.name ?? "Deck Theme")}"><a:themeElements>` +
    `<a:clrScheme name="${esc(t?.name ?? "Deck")}">${["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"].map(clr).join("")}</a:clrScheme>` +
    `<a:fontScheme name="${esc(t?.name ?? "Deck")}"><a:majorFont>${font(f.major ?? "Aptos Display", f.majorEa ?? "Yu Gothic")}</a:majorFont><a:minorFont>${font(f.minor ?? "Aptos", f.minorEa ?? "Yu Gothic")}</a:minorFont></a:fontScheme>` +
    `<a:fmtScheme name="Office"><a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst><a:lnStyleLst>${ln(6350)}${ln(12700)}${ln(19050)}</a:lnStyleLst>` +
    `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
    `<a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

function themeColorKey(k: string): string {
  const m: { [k: string]: string } = { text: "dk1", background: "lt1", text2: "dk2", background2: "lt2", tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2", folhlink: "folHlink" };
  return m[k.toLowerCase()] ?? k;
}

function applyThemeOverrides(deck: DeckCtx): void {
  const t = deck.spec.theme;
  if (!t) return;
  const done: string[] = [];
  for (const m of deck.masters) {
    if (!m.theme || done.indexOf(m.theme) >= 0) continue;
    done.push(m.theme);
    let xml = getText(deck.pkg, m.theme);
    for (const k of Object.keys(t.colors ?? {})) {
      const name = themeColorKey(k);
      xml = xml.replace(new RegExp(`<a:${name}>[\\s\\S]*?</a:${name}>`), `<a:${name}><a:srgbClr val="${hexOf((t.colors ?? {})[k])}"/></a:${name}>`);
    }
    const f = t.fonts ?? {};
    const setFont = (x: string, block: string, tag: string, v: string | undefined): string => v === undefined ? x :
      x.replace(new RegExp(`(<a:${block}>[\\s\\S]*?<a:${tag} typeface=")[^"]*"`), `$1${esc(v)}"`);
    xml = setFont(xml, "majorFont", "latin", f.major);
    xml = setFont(xml, "minorFont", "latin", f.minor);
    xml = setFont(xml, "majorFont", "ea", f.majorEa);
    xml = setFont(xml, "minorFont", "ea", f.minorEa);
    if (f.majorEa) xml = xml.replace(/(<a:majorFont>[\s\S]*?<a:font script="Jpan" typeface=")[^"]*"/, `$1${esc(f.majorEa)}"`);
    if (f.minorEa) xml = xml.replace(/(<a:minorFont>[\s\S]*?<a:font script="Jpan" typeface=")[^"]*"/, `$1${esc(f.minorEa)}"`);
    setText(deck.pkg, m.theme, xml);
  }
}

// master.background / master.logo / master.elements / master.layoutOptions を (組み込み・テンプレート共通で) 適用
function applyMasterExtras(deck: DeckCtx): void {
  const ms = deck.spec.master;
  if (!ms) return;
  for (const m of deck.masters) {
    let xml = getText(deck.pkg, m.part);
    const ctx: PartCtx = { deck, part: m.part, nextId: maxShapeId(xml) + 1 };
    if (ms.background !== undefined) xml = replaceBg(xml, bgXml(ctx, ms.background));
    const back: string[] = [];
    const front: string[] = [];
    for (const e of ms.elements ?? []) (e.front ? front : back).push(renderElement(ctx, e));
    if (ms.logo) front.push(renderElement(ctx, logoElement(deck, ms.logo)));
    xml = insertShapes(xml, back.join(""), front.join(""));
    setText(deck.pkg, m.part, xml);
    for (const key of Object.keys(ms.layoutOptions ?? {})) {
      const opt = (ms.layoutOptions ?? {})[key];
      const layouts = m.layouts.filter(l => layoutMatches(l, key));
      if (layouts.length === 0 && m === deck.masters[0]) deck.warnings.push(`master.layoutOptions: レイアウト「${key}」が見つかりません`);
      for (const l of layouts) {
        let lx = getText(deck.pkg, l.part);
        const lctx: PartCtx = { deck, part: l.part, nextId: maxShapeId(lx) + 1 };
        if (opt.background !== undefined) lx = replaceBg(lx, bgXml(lctx, opt.background));
        if (opt.showMaster !== undefined) {
          lx = lx.replace(/<p:sldLayout\b[^>]*>/, tag => setAttr(tag, "showMasterSp", opt.showMaster ? "1" : "0"));
        }
        const lb: string[] = [];
        const lf: string[] = [];
        for (const e of opt.elements ?? []) (e.front ? lf : lb).push(renderElement(lctx, e));
        lx = insertShapes(lx, lb.join(""), lf.join(""));
        setText(deck.pkg, l.part, lx);
      }
    }
  }
}

function logoElement(deck: DeckCtx, logo: LogoSpec): ElementSpec {
  const pos = (logo.position ?? "topRight").toLowerCase();
  const m = lenEmu(deck, logo.margin, "x") ?? Math.round(deck.W * 0.02);
  const w = lenEmu(deck, logo.w, "x");
  const h = lenEmu(deck, logo.h, "y") ?? (w === undefined ? Math.round(deck.H * 0.07) : undefined);
  const info = mediaInfoForRef(deck, logo.image);
  const ew = w ?? Math.round((h ?? 0) * info.w / info.h);
  const eh = h ?? Math.round(ew * info.h / info.w);
  const x = pos.indexOf("left") >= 0 ? m : pos.indexOf("center") >= 0 ? Math.round((deck.W - ew) / 2) : deck.W - m - ew;
  const y = pos.indexOf("bottom") >= 0 ? deck.H - m - eh : m;
  return { type: "image", name: "Logo", image: logo.image, x: `${x}emu`, y: `${y}emu`, w: `${ew}emu`, h: `${eh}emu`, fit: "stretch" };
}

// =====================================================================
// マスター／レイアウトの解析 (組み込み・テンプレート共通)
// =====================================================================
function analyzeMasters(deck: DeckCtx): MasterInfo[] {
  const pkg = deck.pkg;
  const pres = getText(pkg, deck.presPart);
  const presRels = getRels(pkg, deck.presPart);
  const masters: MasterInfo[] = [];
  for (const mm of allMatches(/<p:sldMasterId\b[^>]*\br:id="([^"]+)"/g, pres)) {
    const rel = presRels.filter(r => r.id === mm[1])[0];
    if (!rel) continue;
    const part = resolveTarget(deck.presPart, rel.target);
    const xml = getText(pkg, part);
    const mrels = getRels(pkg, part);
    const themeRel = mrels.filter(r => r.type === R_THEME)[0];
    const theme = themeRel ? resolveTarget(part, themeRel.target) : "";
    const themeName = theme ? (/<a:theme\b[^>]*\bname="([^"]*)"/.exec(getText(pkg, theme)) ?? ["", ""])[1] : "";
    const titleStyle = (/<p:titleStyle>([\s\S]*?)<\/p:titleStyle>/.exec(xml) ?? ["", ""])[1];
    const bodyStyle = (/<p:bodyStyle>([\s\S]*?)<\/p:bodyStyle>/.exec(xml) ?? ["", ""])[1];
    const info: MasterInfo = {
      part, name: unesc(themeName || `Master ${masters.length + 1}`), theme, phs: parsePhs(xml), layouts: [],
      titleSizes: lvlSizes(titleStyle), bodySizes: lvlSizes(bodyStyle),
    };
    for (const lm of allMatches(/<p:sldLayoutId\b[^>]*\br:id="([^"]+)"/g, xml)) {
      const lrel = mrels.filter(r => r.id === lm[1])[0];
      if (!lrel) continue;
      const lpart = resolveTarget(part, lrel.target);
      const lxml = getText(pkg, lpart);
      const rootTag = (/<p:sldLayout\b[^>]*>/.exec(lxml) ?? [""])[0];
      const type = attrOf(rootTag, "type") ?? "cust";
      const name = unesc((/<p:cSld\b[^>]*\bname="([^"]*)"/.exec(lxml) ?? ["", ""])[1]);
      const keys = (LAYOUT_TYPE_KEYS[type] ?? []).slice();
      const bkey = deck.layoutKeys[lpart];
      if (bkey) keys.unshift(norm(bkey));
      const phs = parsePhs(lxml).map(p => inheritPh(p, info));
      info.layouts.push({
        part: lpart, name, type, keys, phs, master: info,
        showMaster: attrOf(rootTag, "showMasterSp") !== "0", label: bkey ?? keys[0] ?? "",
      });
    }
    masters.push(info);
  }
  if (masters.length === 0) throw new Error("スライドマスターが見つかりません");
  return masters;
}

function parsePhs(xml: string): PhInfo[] {
  const out: PhInfo[] = [];
  for (const m of allMatches(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g, xml)) {
    const block = m[0];
    const ph = /<p:ph\b([^>]*?)\/?>/.exec(block);
    if (!ph) continue;
    const type = attrOf(ph[0], "type") ?? "obj";
    const idx = +(attrOf(ph[0], "idx") ?? "0");
    const name = unesc((/<p:cNvPr\b[^>]*\bname="([^"]*)"/.exec(block) ?? ["", ""])[1]);
    const x = /<a:xfrm\b[^>]*>\s*<a:off x="(-?\d+)" y="(-?\d+)"\s*\/>\s*<a:ext cx="(\d+)" cy="(\d+)"\s*\/>/.exec(block);
    const rect = x ? { x: +x[1], y: +x[2], w: +x[3], h: +x[4] } : null;
    const lst = (/<a:lstStyle>([\s\S]*?)<\/a:lstStyle>/.exec(block) ?? ["", ""])[1];
    out.push({ type, idx, name, rect, sizes: lvlSizes(lst), block });
  }
  return out;
}

function lvlSizes(xml: string): (number | null)[] {
  const out: (number | null)[] = [];
  for (let n = 1; n <= 5; n++) {
    const b = new RegExp(`<a:lvl${n}pPr\\b[^>]*?(?:/>|>([\\s\\S]*?)</a:lvl${n}pPr>)`).exec(xml);
    const sz = b && b[1] ? /<a:defRPr\b[^>]*\bsz="(\d+)"/.exec(b[1]) : null;
    out.push(sz ? +sz[1] / 100 : null);
  }
  return out;
}

function masterPhFor(m: MasterInfo, type: string): PhInfo | undefined {
  const t = TITLE_TYPES.indexOf(type) >= 0 ? "title" : FOOTER_TYPES.indexOf(type) >= 0 ? type : "body";
  return m.phs.filter(p => p.type === t)[0];
}

function inheritPh(p: PhInfo, m: MasterInfo): PhInfo {
  const mp = masterPhFor(m, p.type);
  const isTitle = TITLE_TYPES.indexOf(p.type) >= 0;
  const style = isTitle ? m.titleSizes : m.bodySizes;
  const sizes = [0, 1, 2, 3, 4].map(i => p.sizes[i] ?? (mp ? mp.sizes[i] : null) ?? style[i] ?? (isTitle ? 44 : [28, 24, 20, 20, 20][i]));
  return { type: p.type, idx: p.idx, name: p.name, rect: p.rect ?? (mp ? mp.rect : null), sizes, block: p.block };
}

function describeLayouts(deck: DeckCtx): string {
  const lines: string[] = [];
  deck.masters.forEach((m, mi) => {
    lines.push(`マスター ${mi + 1}: ${m.name}`);
    m.layouts.forEach((l, li) => {
      const content = contentPhs(l.phs);
      const builtin = deck.layoutKeys[l.part] !== undefined;
      const slots = l.phs.filter(p => FOOTER_TYPES.indexOf(p.type) < 0).map(p => {
        if (builtin) return p.name.replace(/\s+(\w)/g, (m: string, c: string) => c.toUpperCase()).replace(/^\w/, c => c.toLowerCase());
        const generic = TITLE_TYPES.indexOf(p.type) >= 0 ? "title" : p.type === "subTitle" ? "subtitle" : p.type === "pic" ? "picture" :
          (content.indexOf(p) === 0 ? "body" : `body${content.indexOf(p) + 1}`);
        return `${generic}「${p.name}」`;
      });
      lines.push(`  [${li + 1}] ${l.name.replace(/\s*\n\s*/g, " ")}${l.label ? ` (layout: "${l.label}")` : ""} … ${slots.join(", ") || "(プレースホルダーなし)"}`);
    });
  });
  return lines.join("\n");
}

function contentPhs(phs: PhInfo[]): PhInfo[] {
  return phs.filter(p => TITLE_TYPES.indexOf(p.type) < 0 && p.type !== "subTitle" && FOOTER_TYPES.indexOf(p.type) < 0)
    .sort((a, b) => a.idx - b.idx);
}

function layoutMatches(l: LayoutInfo, ref: string): boolean {
  const n = norm(ref);
  return norm(l.name) === n || l.keys.indexOf(n) >= 0;
}

function findLayout(deck: DeckCtx, s: SlideSpec, slotKeys: string[]): LayoutInfo {
  let masters = deck.masters;
  if (s.master !== undefined) {
    const m = typeof s.master === "number" ? deck.masters[s.master - 1] : deck.masters.filter(x => norm(x.name) === norm(String(s.master)))[0];
    if (!m) throw new Error(`master「${s.master}」が見つかりません`);
    masters = [m];
  }
  const all: LayoutInfo[] = [];
  for (const m of masters) for (const l of m.layouts) all.push(l);
  let ref = s.layout;
  if (ref === undefined) {
    const others = slotKeys.filter(k => norm(k) !== "title");
    ref = others.length > 0 ? "content" : slotKeys.length > 0 ? "titleOnly" : "blank";
    if (!all.some(l => layoutMatches(l, String(ref)))) ref = 1;
  }
  if (typeof ref === "number") {
    const l = all[ref - 1];
    if (!l) throw new Error(`layout 番号 ${ref} は範囲外です (1〜${all.length})`);
    return l;
  }
  const match = (r: string): LayoutInfo | undefined => all.filter(l => norm(l.name) === norm(r))[0] ?? all.filter(l => l.keys.indexOf(norm(r)) >= 0)[0];
  let hit = match(String(ref));
  // テンプレートにないレイアウトは近いものに置き換える (かんたんモードを社内テンプレートで使うときなど)
  const alt = LAYOUT_FALLBACKS[norm(String(ref))];
  if (!hit && !deck.builtin && alt) {
    for (const a of alt) { hit = match(a); if (hit) break; }
    if (hit) {
      const msg = `テンプレートに「${ref}」レイアウトがないため「${hit.name}」で作成しました`;
      if (deck.warnings.indexOf(msg) < 0) deck.warnings.push(msg);
    }
  }
  if (!hit) {
    throw new Error(`レイアウト「${ref}」が見つかりません。使用可能: ${all.map(l => l.label ? `${l.label}(${l.name})` : l.name).join(", ")}`);
  }
  return hit;
}

// スロット名 (title / subtitle / body / left / caption / picture / #idx / プレースホルダー名) → プレースホルダー
function resolveSlots(type: string, phs: PhInfo[], keys: string[], warnings: string[]): { [key: string]: PhInfo } {
  const result: { [key: string]: PhInfo } = {};
  const used: PhInfo[] = [];
  const content = contentPhs(phs);
  const free = (p: PhInfo | undefined): PhInfo | undefined => (p && used.indexOf(p) < 0 ? p : undefined);
  const claim = (k: string, p: PhInfo | undefined): boolean => { if (!p) return false; result[k] = p; used.push(p); return true; };
  const typeAlias: { [t: string]: { [slot: string]: number } } = {
    twoObj: { left: 0, right: 1 },
    twoTxTwoObj: { leftheading: 0, left: 1, rightheading: 2, right: 3 },
  };
  const pending: string[] = [];
  for (const k of keys) {
    const n = norm(k);
    let p = free(phs.filter(x => norm(x.name) === n && FOOTER_TYPES.indexOf(x.type) < 0)[0]);
    const im = /^#(\d+)$/.exec(k.trim());
    if (!p && im) p = free(phs.filter(x => x.idx === +im[1])[0]);
    if (!p && n === "title") p = free(phs.filter(x => TITLE_TYPES.indexOf(x.type) >= 0)[0]);
    if (!p && n === "subtitle") p = free(phs.filter(x => x.type === "subTitle")[0]);
    if (!p && (n === "picture" || n === "image")) p = free(phs.filter(x => x.type === "pic")[0]);
    if (!p && n === "caption" && content.some(x => x.type !== "body")) p = free(content.filter(x => x.type === "body")[0]);
    if (!p && typeAlias[type] && typeAlias[type][n] !== undefined) p = free(content[typeAlias[type][n]]);
    if (!claim(k, p)) pending.push(k);
  }
  for (const k of pending) {
    const n = norm(k);
    const m = /^(?:body|content)(\d*)$/.exec(n);
    let p: PhInfo | undefined;
    if (m && m[1]) p = free(content[+m[1] - 1]);
    else if (m || ["subtitle", "left", "right", "center", "caption", "leftheading", "rightheading", "text", "picture", "image"].indexOf(n) >= 0) {
      p = content.filter(x => used.indexOf(x) < 0 && (n === "picture" || n === "image" || x.type !== "pic"))[0];
    }
    if (!claim(k, p)) warnings.push(`「${k}」に対応するプレースホルダーがレイアウトにありません (無視しました)`);
  }
  return result;
}

// =====================================================================
// スライド生成
// =====================================================================
function slotKeysOf(s: SlideSpec): string[] {
  return Object.keys(s).filter(k => RESERVED_SLIDE_KEYS.indexOf(k.toLowerCase()) < 0 && !/^(_|\/\/)/.test(k) && s[k] !== undefined && s[k] !== null);
}

function buildSlide(deck: DeckCtx, s: SlideSpec, number: number): NewSlide {
  const pkg = deck.pkg;
  const keys = slotKeysOf(s);
  const layout = findLayout(deck, s, keys);
  const part = nextPartName(pkg, "ppt/slides/slide", ".xml");
  addRel(pkg, part, R_LAYOUT, relTarget(part, layout.part));
  const ctx: PartCtx = { deck, part, nextId: 2 };
  const shapes: string[] = [];
  const slots = resolveSlots(layout.type, layout.phs, keys, deck.warnings);
  for (const k of keys) {
    const ph = slots[k];
    if (!ph) continue;
    shapes.push(fillPlaceholder(ctx, ph, s[k] as SlotValue));
  }
  shapes.push(footerShapes(ctx, s, layout, number));
  for (const e of s.elements ?? []) shapes.push(renderElement(ctx, e));
  const bg = s.background !== undefined ? bgXml(ctx, s.background) : "";
  const tr = transitionXml(s.transition ?? deck.spec.transition);
  setText(pkg, part, XML_HEAD + `<p:sld ${NS}${s.hidden ? ` show="0"` : ""}><p:cSld>${bg}${spTreeOpen()}${shapes.join("")}</p:spTree></p:cSld>` +
    `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${tr}</p:sld>`, `${CT_P}.slide+xml`);
  if (s.notes !== undefined) addNotes(deck, part, s.notes);
  return { part, section: s.section };
}

function isElementValue(v: SlotValue): v is ElementSpec {
  if (v === null || typeof v !== "object" || Array.isArray(v)) return false;
  const e = v as ElementSpec;
  return e.type !== undefined || e.image !== undefined || e.chartType !== undefined || e.rows !== undefined || e.data !== undefined || e.shape !== undefined;
}

function fillPlaceholder(ctx: PartCtx, ph: PhInfo, v: SlotValue): string {
  const deck = ctx.deck;
  const rect = ph.rect ?? { x: Math.round(deck.W * 0.055), y: Math.round(deck.H * 0.24), w: Math.round(deck.W * 0.89), h: Math.round(deck.H * 0.64) };
  if (ph.type === "pic" && typeof v === "string") v = { type: "image", image: v };
  if (isElementValue(v)) {
    const el = Object.assign({}, v);
    if (elementType(el) === "image" && el.fit === undefined) el.fit = ph.type === "pic" ? "cover" : "contain";
    return renderElement(ctx, el, rect);
  }
  const node: TextNode = typeof v === "object" && !Array.isArray(v) ? v as TextNode : {};
  const own = typeof v === "object" && !Array.isArray(v) ? (v as ElementSpec) : {};
  let r = rect;
  if (own.x !== undefined || own.y !== undefined || own.w !== undefined || own.h !== undefined) r = rectOf(deck, own, rect);
  const isTitle = TITLE_TYPES.indexOf(ph.type) >= 0;
  const paras = buildParas(ctx, v as TextValue, withStyle(deck, node as ElementSpec), {
    placeholder: true, defaultSize: (lvl: number) => ph.sizes[Math.min(lvl, 4)] ?? (isTitle ? 44 : 24),
  });
  const xfrm = r !== rect ? `<a:xfrm><a:off x="${r.x}" y="${r.y}"/><a:ext cx="${r.w}" cy="${r.h}"/></a:xfrm>` : "";
  const st = withStyle(deck, node as ElementSpec);
  const fill = st.fill !== undefined ? fillXml(ctx, st.fill) : "";
  const ln = st.line !== undefined ? lineXml(st.line) : "";
  const phAttr = (ph.type === "obj" ? "" : ` type="${ph.type}"`) + (ph.idx ? ` idx="${ph.idx}"` : "");
  const id = ctx.nextId++;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(ph.name || "Placeholder " + id)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${phAttr}/></p:nvPr></p:nvSpPr>` +
    `<p:spPr>${xfrm}${fill}${ln}</p:spPr><p:txBody>${bodyPrXml(deck, st, r, paras, undefined, true)}<a:lstStyle/>${parasXml(paras)}</p:txBody></p:sp>`;
}

function footerShapes(ctx: PartCtx, s: SlideSpec, layout: LayoutInfo, number: number): string {
  const deck = ctx.deck;
  // デッキ全体のフッター設定は、タイトルスライドとマスター非表示のレイアウトには適用しない (スライド個別の指定は常に優先)
  const isTitle = layout.type === "title" || !layout.showMaster;
  const pick = <T>(slideVal: T | undefined, deckVal: T | undefined): T | undefined => slideVal !== undefined ? slideVal : isTitle ? undefined : deckVal;
  const footer = pick(s.footer, deck.spec.footer);
  const num = pick(s.slideNumber, deck.spec.slideNumber);
  const date = pick(s.date, deck.spec.date);
  const out: string[] = [];
  const ph = (type: string): PhInfo | undefined => layout.phs.filter(p => p.type === type)[0] ?? masterPhFor(layout.master, type);
  const add = (type: string, inner: string, label: string): void => {
    const p = ph(type);
    if (!p) { deck.warnings.push(`レイアウト「${layout.name}」に${label}のプレースホルダーがないため表示できません`); return; }
    out.push(`<p:sp><p:nvSpPr><p:cNvPr id="${ctx.nextId++}" name="${label}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${type}" sz="quarter" idx="${p.idx}"/></p:nvPr></p:nvSpPr><p:spPr/>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${inner}</a:p></p:txBody></p:sp>`);
  };
  if (typeof footer === "string" && footer !== "") add("ftr", `<a:r><a:rPr lang="ja-JP" altLang="en-US"/><a:t>${esc(footer)}</a:t></a:r>`, "フッター");
  if (num) add("sldNum", `<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="ja-JP"/><a:t>${number}</a:t></a:fld>`, "スライド番号");
  if (date) {
    const text = typeof date === "string" ? date : new Date(Date.now() + 9 * 3600000).toISOString().slice(0, 10).replace(/-/g, "/");
    add("dt", `<a:r><a:rPr lang="ja-JP" altLang="en-US"/><a:t>${esc(text)}</a:t></a:r>`, "日付");
  }
  return out.join("");
}

function transitionXml(t: string | TransitionObj | undefined): string {
  if (t === undefined) return "";
  const o: TransitionObj = typeof t === "string" ? { type: t } : t;
  const type = (o.type ?? "fade").toLowerCase();
  if (type === "none") return "";
  const dur = o.duration;
  const spd = dur === undefined ? "med" : dur <= 0.5 ? "fast" : dur <= 1 ? "med" : "slow";
  const dir = o.dir;
  const map: { [k: string]: string } = {
    fade: "<p:fade/>", push: `<p:push dir="${dir ?? "u"}"/>`, wipe: `<p:wipe dir="${dir ?? "r"}"/>`,
    cover: `<p:cover dir="${dir ?? "l"}"/>`, pull: `<p:pull dir="${dir ?? "l"}"/>`, split: `<p:split orient="horz" dir="out"/>`,
    cut: "<p:cut/>", dissolve: "<p:dissolve/>", zoom: "<p:zoom/>", random: "<p:random/>", circle: "<p:circle/>",
    wheel: `<p:wheel spokes="4"/>`, blinds: `<p:blinds dir="horz"/>`, checker: `<p:checker dir="horz"/>`,
  };
  const inner = map[type];
  if (!inner) return "";
  return `<p:transition spd="${spd}">${inner}</p:transition>`;
}

function addNotes(deck: DeckCtx, slidePart: string, notes: TextValue): void {
  const pkg = deck.pkg;
  ensureNotesMaster(deck);
  const part = nextPartName(pkg, "ppt/notesSlides/notesSlide", ".xml");
  addRel(pkg, part, R_NOTES_MASTER, relTarget(part, deck.notesMaster));
  addRel(pkg, part, R_SLIDE, relTarget(part, slidePart));
  addRel(pkg, slidePart, R_NOTES, relTarget(slidePart, part));
  const ctx: PartCtx = { deck, part, nextId: 4 };
  const paras = buildParas(ctx, notes, {}, { placeholder: true, defaultSize: () => 12 });
  setText(pkg, part, XML_HEAD + `<p:notes ${NS}><p:cSld>${spTreeOpen()}` +
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/>${parasXml(paras)}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`,
    `${CT_P}.notesSlide+xml`);
}

function ensureNotesMaster(deck: DeckCtx): void {
  if (deck.notesMaster) return;
  const pkg = deck.pkg;
  const existing = getRels(pkg, deck.presPart).filter(r => r.type === R_NOTES_MASTER)[0];
  if (existing) { deck.notesMaster = resolveTarget(deck.presPart, existing.target); return; }
  const part = nextPartName(pkg, "ppt/notesMasters/notesMaster", ".xml");
  const theme = nextPartName(pkg, "ppt/theme/theme", ".xml");
  setText(pkg, theme, themeXml(deck.builtin ? deck.spec.theme : undefined), "application/vnd.openxmlformats-officedocument.theme+xml");
  addRel(pkg, part, R_THEME, relTarget(part, theme));
  const sp = (id: number, name: string, ph: string, x: number, y: number, w: number, h: number, body: string): string =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>${body}</p:sp>`;
  const imgH = Math.round(5486400 * Math.min(1, deck.H / deck.W * 1.6));
  setText(pkg, part, XML_HEAD + `<p:notesMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>${spTreeOpen()}` +
    sp(2, "Slide Image Placeholder", `<p:ph type="sldImg" idx="2"/>`, 685800, 685800, 5486400, Math.round(5486400 * deck.H / deck.W), "") +
    sp(3, "Notes Placeholder", `<p:ph type="body" sz="quarter" idx="3"/>`, 685800, 685800 + imgH + 228600, 5486400, 9144000 - imgH - 685800 * 2 - 228600,
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ja-JP"/></a:p></p:txBody>`) +
    `</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
    `<p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>`,
    `${CT_P}.notesMaster+xml`);
  deck.notesMaster = part;
  deck.needNotesMaster = true;
}

// テンプレート内のスライドを複製して {{トークン}} 置換・スロット差し込み・要素追加を行う
function copyTemplateSlide(deck: DeckCtx, s: SlideSpec, orig: string[], number: number): NewSlide {
  if (deck.builtin) throw new Error("base (テンプレートのスライド複製) はテンプレート指定時のみ使えます");
  const src = orig[(s.base ?? 0) - 1];
  if (!src) throw new Error(`base: ${s.base} のスライドがテンプレートにありません (1〜${orig.length})`);
  const pkg = deck.pkg;
  const part = nextPartName(pkg, "ppt/slides/slide", ".xml");
  for (const r of getRels(pkg, src)) {
    if (r.type === R_NOTES) continue;
    if (r.type === R_CHART && !r.external) {
      addRel(pkg, part, r.type, relTarget(part, duplicateChart(deck, resolveTarget(src, r.target))), false, r.id);
    } else {
      addRel(pkg, part, r.type, r.external ? r.target : relTarget(part, resolveTarget(src, r.target)), r.external, r.id);
    }
  }
  let xml = getText(pkg, src);
  xml = replaceTokens(xml, Object.assign({}, deck.tokens, s.replace ?? {}));
  const ctx: PartCtx = { deck, part, nextId: maxShapeId(xml) + 1 };
  const layoutRel = getRels(pkg, part).filter(r => r.type === R_LAYOUT)[0];
  const layout = deck.masters.map(m => m.layouts.filter(l => layoutRel && l.part === resolveTarget(part, layoutRel.target))[0]).filter(l => !!l)[0];
  const slidePhs = parsePhs(xml).map(p => {
    const lp = layout ? layout.phs.filter(x => x.idx === p.idx && (x.type === p.type || p.type === "obj"))[0] : undefined;
    return lp ? { type: p.type, idx: p.idx, name: p.name, rect: p.rect ?? lp.rect, sizes: lp.sizes, block: p.block } : p;
  });
  const keys = slotKeysOf(s);
  const slots = resolveSlots(layout ? layout.type : "cust", slidePhs, keys, deck.warnings);
  const extra: string[] = [];
  for (const k of keys) {
    const ph = slots[k];
    if (!ph) continue;
    const filled = fillPlaceholder(ctx, ph, s[k] as SlotValue);
    if (/<p:ph\b/.test(filled)) {
      const body = /<p:txBody>[\s\S]*<\/p:txBody>/.exec(filled);
      xml = xml.replace(ph.block, body ? (/<p:txBody>/.test(ph.block) ? ph.block.replace(/<p:txBody>[\s\S]*<\/p:txBody>/, body[0]) : ph.block.replace("</p:sp>", body[0] + "</p:sp>")) : ph.block);
    } else {
      xml = xml.replace(ph.block, "");
      extra.push(filled);
    }
  }
  if (layout) extra.push(footerShapes(ctx, s, layout, number));
  for (const e of s.elements ?? []) extra.push(renderElement(ctx, e));
  xml = insertShapes(xml, "", extra.join(""));
  if (s.background !== undefined) xml = replaceBg(xml, bgXml(ctx, s.background));
  if (s.hidden !== undefined) xml = xml.replace(/<p:sld\b[^>]*>/, tag => setAttr(tag, "show", s.hidden ? "0" : "1"));
  const tr = transitionXml(s.transition ?? deck.spec.transition);
  if (tr) {
    xml = xml.replace(/<p:transition\b[\s\S]*?<\/p:transition>|<p:transition\b[^>]*\/>/, "");
    xml = xml.replace(/(<\/p:clrMapOvr>)/, `$1${tr}`);
  }
  setText(pkg, part, xml, `${CT_P}.slide+xml`);
  if (s.notes !== undefined) addNotes(deck, part, s.notes);
  return { part, section: s.section };
}

function duplicateChart(deck: DeckCtx, chartPart: string): string {
  const pkg = deck.pkg;
  const part = nextPartName(pkg, "ppt/charts/chart", ".xml");
  const p = pkg.parts[chartPart];
  if (!p) return chartPart;
  pkg.parts[part] = { raw: p.raw, bytes: p.bytes, text: p.text, dirty: p.dirty };
  pkg.ctOverrides["/" + part] = pkg.ctOverrides["/" + chartPart] ?? "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
  for (const r of getRels(pkg, chartPart)) {
    if (r.external) { addRel(pkg, part, r.type, r.target, true, r.id); continue; }
    const target = resolveTarget(chartPart, r.target);
    const dot = target.lastIndexOf(".");
    const copy = nextPartName(pkg, target.slice(0, dot).replace(/\d+$/, ""), target.slice(dot));
    const tp = pkg.parts[target];
    if (tp) {
      pkg.parts[copy] = { raw: tp.raw, bytes: tp.bytes, text: tp.text, dirty: tp.dirty };
      if (pkg.ctOverrides["/" + target]) pkg.ctOverrides["/" + copy] = pkg.ctOverrides["/" + target];
    }
    addRel(pkg, part, r.type, relTarget(part, tp ? copy : target), false, r.id);
  }
  return part;
}

function replaceTokens(xml: string, map: { [token: string]: string }): string {
  const keys = Object.keys(map);
  if (keys.length === 0) return xml;
  return xml.replace(/<a:p(?:\s[^>]*)?>[\s\S]*?<\/a:p>/g, (para: string) => {
    const texts = allMatches(/<a:t>([\s\S]*?)<\/a:t>/g, para).map(m => unesc(m[1]));
    const joined = texts.join("");
    if (!keys.some(k => joined.indexOf(k) >= 0)) return para;
    const rep = (t: string): string => { for (const k of keys) t = t.split(k).join(map[k] ?? ""); return t; };
    const inSingle = keys.every(k => joined.indexOf(k) < 0 || texts.some(t => t.indexOf(k) >= 0));
    if (inSingle) return para.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (m: string, t: string) => `<a:t>${esc(rep(unesc(t)))}</a:t>`);
    let first = true;
    return para.replace(/<a:t>([\s\S]*?)<\/a:t>/g, () => {
      const v = first ? esc(rep(joined)) : "";
      first = false;
      return `<a:t>${v}</a:t>`;
    });
  });
}

// =====================================================================
// 要素 (text / shape / image / table / chart / line)
// =====================================================================
function elementType(el: ElementSpec): string {
  if (el.type) return el.type.toLowerCase();
  if (el.image !== undefined) return "image";
  if (el.chartType !== undefined) return "chart";
  if (el.rows !== undefined || el.data !== undefined) return "table";
  if (el.shape !== undefined) return "shape";
  return "text";
}

function withStyle(deck: DeckCtx, el: ElementSpec): ElementSpec {
  if (!el.style) return el;
  const names = el.style.split(/[\s,]+/).filter(n => n);
  const merged: ElementSpec = {};
  for (const n of names) {
    const st = deck.styles[n];
    if (!st) { deck.warnings.push(`style「${n}」が styles に定義されていません`); continue; }
    Object.assign(merged, withStyle(deck, st as ElementSpec));
  }
  const out = Object.assign(merged, el);
  delete out.style;
  return out;
}

function renderElement(ctx: PartCtx, el0: ElementSpec, def?: Rect): string {
  const el = withStyle(ctx.deck, el0);
  const type = elementType(el);
  switch (type) {
    case "text": return shapeXml(ctx, el, rectOf(ctx.deck, el, def), true);
    case "shape": return shapeXml(ctx, el, rectOf(ctx.deck, el, def), false);
    case "image": return imageXml(ctx, el, def);
    case "table": return tableXml(ctx, el, rectOf(ctx.deck, el, def));
    case "chart": return chartFrameXml(ctx, el, rectOf(ctx.deck, el, def));
    case "line": return lineElementXml(ctx, el);
    default:
      ctx.deck.warnings.push(`要素 type「${el.type}」は未対応です (text / shape / image / table / chart / line)`);
      return "";
  }
}

function lenEmu(deck: DeckCtx, v: Len | undefined, axis: string): number | undefined {
  if (v === undefined || v === null) return undefined;
  if (typeof v === "number") return Math.round(v * deck.unit);
  const s = String(v).trim();
  let m = /^(-?[\d.]+)\s*%$/.exec(s);
  if (m) return Math.round(+m[1] / 100 * (axis === "x" ? deck.W : deck.H));
  m = /^(-?[\d.]+)\s*(emu|in|cm|mm|pt|px)$/i.exec(s);
  if (m) return Math.round(+m[1] * UNIT_EMU[m[2].toLowerCase()]);
  if (/^-?[\d.]+$/.test(s)) return Math.round(+s * deck.unit);
  throw new Error(`位置・サイズの指定「${v}」が不正です (例: 2.5 / "50%" / "1in")`);
}

function rectOf(deck: DeckCtx, el: ElementSpec, def?: Rect, wFix?: number, hFix?: number): Rect {
  const w = wFix ?? lenEmu(deck, el.w, "x") ?? def?.w ?? Math.round(deck.W * 0.4);
  const h = hFix ?? lenEmu(deck, el.h, "y") ?? def?.h ?? Math.round(deck.H * 0.15);
  const pos = (v: Len | undefined, axis: string, size: number, full: number, d: number | undefined): number => {
    if (v === "center") return Math.round((full - size) / 2);
    const n = lenEmu(deck, v, axis);
    if (n === undefined) return d ?? Math.round((full - size) / 2);
    return n < 0 ? full + n - size : n;
  };
  return { x: pos(el.x, "x", w, deck.W, def?.x), y: pos(el.y, "y", h, deck.H, def?.y), w, h };
}

function xfrmXml(r: Rect, rot?: number, extra = ""): string {
  return `<a:xfrm${rot ? ` rot="${Math.round(rot * 60000)}"` : ""}${extra}><a:off x="${r.x}" y="${r.y}"/><a:ext cx="${Math.max(0, r.w)}" cy="${Math.max(0, r.h)}"/></a:xfrm>`;
}

const SHADOW = `<a:effectLst><a:outerShdw blurRad="76200" dist="25400" dir="5400000" algn="t" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="28000"/></a:srgbClr></a:outerShdw></a:effectLst>`;

function shapeXml(ctx: PartCtx, el: ElementSpec, r: Rect, isText: boolean): string {
  const deck = ctx.deck;
  const id = ctx.nextId++;
  let geom = isText ? "rect" : (el.shape ?? "rect");
  if (PRESET_SHAPES.indexOf(geom) < 0) { deck.warnings.push(`shape「${geom}」は未対応のため rect にしました`); geom = "rect"; }
  const fill = el.fill !== undefined ? fillXml(ctx, el.fill) : isText ? "<a:noFill/>" : `<a:solidFill>${colorXml("accent1")}</a:solidFill>`;
  const ln = el.line !== undefined ? lineXml(el.line) : isText ? "" : "<a:ln><a:noFill/></a:ln>";
  // radius (角の丸み) は adj ガイドを 1 つだけ持つ角丸系の図形にのみ適用する (楕円などに付けると PowerPoint が開けなくなる)
  const av = el.radius !== undefined && ["roundRect", "round1Rect", "snip1Rect", "plaque", "foldedCorner", "bevel", "frame", "can", "cube"].indexOf(geom) >= 0
    ? `<a:gd name="adj" fmla="val ${Math.round(el.radius * 1000)}"/>` : "";
  const link = el.link && el.text === undefined && el.runs === undefined ? hlinkXml(ctx, el.link) : "";
  const base: ElementSpec = Object.assign({}, deck.spec.defaults ?? {},
    isText ? {} : { align: "center", color: autoTextColor(el.fill) }, el);
  if (base.color === undefined) delete base.color;
  const paras = el.text !== undefined || el.runs !== undefined
    ? buildParas(ctx, el.runs !== undefined ? { runs: el.runs } : el.text as TextValue, base, { placeholder: false, defaultSize: () => (base.size ?? 18) })
    : [];
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(el.name ?? `${isText ? "TextBox" : "Shape"} ${id}`)}"${el.alt ? ` descr="${esc(el.alt)}"` : ""}>${link}</p:cNvPr>` +
    `<p:cNvSpPr${isText ? ` txBox="1"` : ""}/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrmXml(r, el.rotation)}<a:prstGeom prst="${geom}"><a:avLst>${av}</a:avLst></a:prstGeom>${fill}${ln}${el.shadow ? SHADOW : ""}</p:spPr>` +
    `<p:txBody>${bodyPrXml(deck, base, r, paras, isText ? "t" : "ctr", false)}<a:lstStyle/>${parasXml(paras)}</p:txBody></p:sp>`;
}

function lineElementXml(ctx: PartCtx, el: ElementSpec): string {
  const deck = ctx.deck;
  const x1 = lenEmu(deck, el.x, "x") ?? 0;
  const y1 = lenEmu(deck, el.y, "y") ?? 0;
  const x2 = el.x2 !== undefined ? (lenEmu(deck, el.x2, "x") ?? 0) : x1 + (lenEmu(deck, el.w, "x") ?? 0);
  const y2 = el.y2 !== undefined ? (lenEmu(deck, el.y2, "y") ?? 0) : y1 + (lenEmu(deck, el.h, "y") ?? 0);
  const r: Rect = { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1) };
  const flip = `${x2 < x1 ? ` flipH="1"` : ""}${y2 < y1 ? ` flipV="1"` : ""}`;
  const lo: LineObj = typeof el.line === "object" ? el.line : { color: typeof el.line === "string" ? el.line : el.color };
  const ends = (el.head ? `<a:headEnd type="${arrowType(el.head)}"/>` : "") + (el.tail ? `<a:tailEnd type="${arrowType(el.tail)}"/>` : "");
  const id = ctx.nextId++;
  return `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="${esc(el.name ?? `Line ${id}`)}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
    `<p:spPr>${xfrmXml(r, undefined, flip)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>${lineXml(Object.assign({ color: "tx1", width: 1.5 }, lo), ends)}</p:spPr></p:cxnSp>`;
}

function arrowType(v: string): string {
  const m: { [k: string]: string } = { arrow: "triangle", triangle: "triangle", stealth: "stealth", open: "arrow", oval: "oval", diamond: "diamond", none: "none" };
  return m[v.toLowerCase()] ?? "triangle";
}

function hlinkXml(ctx: PartCtx, link: string): string {
  return `<a:hlinkClick r:id="${addRel(ctx.deck.pkg, ctx.part, R_LINK, link, true)}"/>`;
}

// ---------- 画像 ----------
function mediaInfoForRef(deck: DeckCtx, ref: string): MediaInfo {
  return deck.mediaInfo[mediaPartFor(deck, ref)];
}

function mediaPartFor(deck: DeckCtx, ref: string): string {
  let b64 = deck.assets[ref];
  if (b64 === undefined) {
    if (/^data:/.test(ref) || ref.length > 200) b64 = ref;
    else throw new Error(`画像アセット「${ref}」が見つかりません (構成 JSON の assets に name="${ref}" を追加してください)`);
  }
  const key = "asset:" + (b64 === ref ? ref.slice(0, 64) + ref.length : ref);
  if (deck.media[key]) return deck.media[key];
  const bytes = base64Decode(b64);
  const info = imageInfo(bytes, ref);
  const part = nextPartName(deck.pkg, "ppt/media/image", "." + info.ext);
  setBytes(deck.pkg, part, bytes);
  const mime: { [e: string]: string } = { png: "image/png", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp" };
  deck.pkg.ctDefaults[info.ext] = mime[info.ext];
  deck.media[key] = part;
  deck.mediaInfo[part] = info;
  return part;
}

function embedImage(ctx: PartCtx, ref: string): { rid: string; info: MediaInfo } {
  const part = mediaPartFor(ctx.deck, ref);
  return { rid: addRel(ctx.deck.pkg, ctx.part, R_IMAGE, relTarget(ctx.part, part)), info: ctx.deck.mediaInfo[part] };
}

function imageInfo(b: Uint8Array, ref: string): MediaInfo {
  const be16 = (i: number): number => (b[i] << 8) | b[i + 1];
  const be32 = (i: number): number => ((b[i] << 24) | (b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3]) >>> 0;
  const le16 = (i: number): number => b[i] | (b[i + 1] << 8);
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return { ext: "png", w: be32(16), h: be32(20) };
  if (b[0] === 0xff && b[1] === 0xd8) {
    let i = 2;
    while (i < b.length - 9) {
      if (b[i] !== 0xff) { i++; continue; }
      const mk = b[i + 1];
      if (mk >= 0xc0 && mk <= 0xcf && mk !== 0xc4 && mk !== 0xc8 && mk !== 0xcc) return { ext: "jpeg", w: be16(i + 7), h: be16(i + 5) };
      if (mk === 0xd8 || mk === 0x01 || (mk >= 0xd0 && mk <= 0xd7)) { i += 2; continue; }
      i += 2 + be16(i + 2);
    }
    return { ext: "jpeg", w: 800, h: 600 };
  }
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return { ext: "gif", w: le16(6), h: le16(8) };
  if (b[0] === 0x42 && b[1] === 0x4d) return { ext: "bmp", w: le16(18) | (le16(20) << 16), h: Math.abs((le16(22) | (le16(24) << 16)) << 0) };
  if (b[0] === 0x3c || (b[0] === 0xef && b[3] === 0x3c)) throw new Error(`画像「${ref.slice(0, 40)}」は SVG です。PNG または JPEG に変換してください`);
  throw new Error(`画像「${ref.slice(0, 40)}」の形式を判別できません (PNG / JPEG / GIF / BMP に対応)`);
}

function imageXml(ctx: PartCtx, el: ElementSpec, def?: Rect): string {
  const deck = ctx.deck;
  if (!el.image) { deck.warnings.push("image 要素に image (アセット名) がありません"); return ""; }
  const img = embedImage(ctx, el.image);
  const iw = img.info.w || 1, ih = img.info.h || 1;
  let w = lenEmu(deck, el.w, "x");
  let h = lenEmu(deck, el.h, "y");
  const boxed = (w !== undefined && h !== undefined) || (w === undefined && h === undefined && def !== undefined);
  if (w === undefined && h === undefined && def) { w = def.w; h = def.h; }
  if (w === undefined && h === undefined) { w = Math.min(iw * 9525, Math.round(deck.W * 0.5)); h = Math.round(w * ih / iw); }
  else if (w === undefined) w = Math.round((h ?? 0) * iw / ih);
  else if (h === undefined) h = Math.round(w * ih / iw);
  let r = rectOf(deck, el, def, w, h);
  const fit = (el.fit ?? (boxed ? "contain" : "stretch")).toLowerCase();
  let src = "";
  if (fit === "contain") {
    const s = Math.min(r.w / iw, r.h / ih);
    const nw = Math.round(iw * s), nh = Math.round(ih * s);
    r = { x: r.x + Math.round((r.w - nw) / 2), y: r.y + Math.round((r.h - nh) / 2), w: nw, h: nh };
  } else if (fit === "cover") {
    const boxA = r.w / r.h, imgA = iw / ih;
    if (imgA > boxA) { const c = Math.round((1 - boxA / imgA) / 2 * 100000); src = `<a:srcRect l="${c}" r="${c}"/>`; }
    else if (imgA < boxA) { const c = Math.round((1 - imgA / boxA) / 2 * 100000); src = `<a:srcRect t="${c}" b="${c}"/>`; }
  }
  let geom = el.shape ?? "rect";
  if (PRESET_SHAPES.indexOf(geom) < 0) geom = "rect";
  const id = ctx.nextId++;
  const link = el.link ? hlinkXml(ctx, el.link) : "";
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(el.name ?? `Picture ${id}`)}" descr="${esc(el.alt ?? "")}">${link}</p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${img.rid}"/>${src}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${xfrmXml(r, el.rotation)}<a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom>${el.line !== undefined ? lineXml(el.line) : ""}${el.shadow ? SHADOW : ""}</p:spPr></p:pic>`;
}

// ---------- 表 ----------
interface TCell { text: string; st: StyleSpec; colspan: number; rowspan: number; num: boolean }

function cellOf(v: CellValue | undefined, col: ColumnSpec | undefined): TCell {
  const base: StyleSpec = {};
  if (v !== null && v !== undefined && typeof v === "object") {
    const o = v as CellObj;
    const t = o.text;
    const num = typeof t === "number";
    return { text: num ? formatNumber(t as number, col?.format) : t === null || t === undefined ? "" : String(t), st: o, colspan: o.colspan ?? 1, rowspan: o.rowspan ?? 1, num };
  }
  if (typeof v === "number") return { text: formatNumber(v, col?.format), st: base, colspan: 1, rowspan: 1, num: true };
  if (typeof v === "boolean") return { text: v ? "✓" : "", st: base, colspan: 1, rowspan: 1, num: false };
  return { text: v === null || v === undefined ? "" : String(v), st: base, colspan: 1, rowspan: 1, num: false };
}

function formatNumber(n: number, fmt?: string): string {
  if (!fmt) return Number.isInteger(n) ? groupDigits(String(n)) : groupDigits(String(Math.round(n * 100) / 100));
  const m = /^([^#0]*)([#,0]*)(?:\.(0+))?(%?)(.*)$/.exec(fmt);
  if (!m) return String(n);
  const dec = m[3] ? m[3].length : 0;
  const v = m[4] ? n * 100 : n;
  let s = Math.abs(v).toFixed(dec);
  if (m[2].indexOf(",") >= 0) {
    const [i, f] = s.split(".");
    s = groupDigits(i) + (f ? "." + f : "");
  }
  return (v < 0 ? "-" : "") + m[1] + s + m[4] + m[5];
}

function groupDigits(s: string): string {
  const neg = s.charAt(0) === "-";
  const [i, f] = (neg ? s.slice(1) : s).split(".");
  return (neg ? "-" : "") + i.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + (f ? "." + f : "");
}

function tableXml(ctx: PartCtx, el: ElementSpec, r: Rect): string {
  const deck = ctx.deck;
  let cols: ColumnSpec[] = el.columns ?? [];
  let header: TCell[] | null = null;
  let body: TCell[][] = [];
  if (el.data) {
    if (el.data.length === 0 && cols.length === 0) { deck.warnings.push("table の data が空です"); return ""; }
    if (cols.length === 0) cols = Object.keys(el.data[0]).map(k => ({ key: k }));
    header = el.header === false ? null : cols.map(c => cellOf(c.label ?? c.key ?? "", undefined));
    body = el.data.map(row => cols.map(c => cellOf(row[c.key ?? ""], c)));
  } else if (el.rows) {
    const m = el.rows.map(row => row.map((v, i) => cellOf(v, cols[i])));
    if (el.header !== false && m.length > 0) { header = m[0]; body = m.slice(1); } else body = m;
  } else {
    return "";
  }
  const allRows = header ? [header].concat(body) : body;
  const nCols = Math.max(cols.length, ...allRows.map(row => row.reduce((a, c) => a + c.colspan, 0)));
  if (nCols === 0) return "";
  // セル結合を考慮したグリッド配置
  const grid: (TCell | string | null)[][] = allRows.map(() => new Array<TCell | string | null>(nCols).fill(null));
  allRows.forEach((row, ri) => {
    let c = 0;
    for (const cell of row) {
      while (c < nCols && grid[ri][c] !== null) c++;
      if (c >= nCols) break;
      grid[ri][c] = cell;
      for (let dr = 0; dr < cell.rowspan && ri + dr < allRows.length; dr++) {
        for (let dc = 0; dc < cell.colspan && c + dc < nCols; dc++) {
          if (dr === 0 && dc === 0) continue;
          grid[ri + dr][c + dc] = dr > 0 && dc > 0 ? "hv" : dr > 0 ? "v" : "h";
        }
      }
      c += cell.colspan;
    }
  });
  const bodyStart = header ? 1 : 0;
  const numericCol = range(nCols).map(ci => {
    let any = false;
    for (let ri = bodyStart; ri < allRows.length; ri++) {
      const g = grid[ri][ci];
      if (g === null || typeof g === "string") continue;
      if (g.num) any = true; else if (g.text !== "") return false;
    }
    return any;
  });

  const sc = deck.scale;
  const rowCount = allRows.length;
  const fontSz = el.size ?? Math.round((rowCount > 14 ? 10 : rowCount > 10 ? 11 : rowCount > 7 ? 12 : 14) * sc);
  const rowH = lenEmu(deck, el.rowHeight, "y") ?? Math.min(Math.round(457200 * sc), Math.floor(r.h / rowCount));
  let colWs: number[];
  // 列幅: colWidths / columns[].width の指定を優先し、残りの幅を文字量に比例して配分
  const fixed: (number | undefined)[] = range(nCols).map(ci => lenEmu(deck, el.colWidths ? el.colWidths[ci] : cols[ci]?.width, "x"));
  const weights = range(nCols).map(ci => Math.min(40, Math.max(4, ...allRows.map((row, ri) => {
    const g = grid[ri][ci];
    return g && typeof g !== "string" && g.colspan === 1 ? Math.max(...g.text.split("\n").map(textWidth)) : 0;
  }))));
  const rest = Math.max(0, r.w - fixed.reduce((a: number, b) => a + (b ?? 0), 0));
  const freeWeight = weights.reduce((a, w, i) => a + (fixed[i] === undefined ? w : 0), 0);
  colWs = range(nCols).map(ci => fixed[ci] ?? Math.floor(rest * weights[ci] / (freeWeight || 1)));
  const styleId = el.tableStyle ? (TABLE_STYLES[norm(el.tableStyle)] ?? (/^\{[0-9A-F-]+\}$/i.test(el.tableStyle) ? el.tableStyle : TABLE_STYLES.medium2accent1)) : TABLE_STYLES.medium2accent1;
  const border = el.borderColor ? (side: string): string => `<a:${side} w="${Math.round((el.borderWidth ?? 0.75) * 12700)}"><a:solidFill>${colorXml(el.borderColor as string)}</a:solidFill></a:${side}>` : null;
  const rowsXml = allRows.map((_, ri) => {
    const isHeader = header !== null && ri === 0;
    const bodyIndex = header ? ri - 1 : ri;
    const tcs = range(nCols).map(ci => {
      const g = grid[ri][ci];
      const borders = border ? border("lnL") + border("lnR") + border("lnT") + border("lnB") : "";
      if (g === null || typeof g === "string") {
        const attr = g === "h" ? ` hMerge="1"` : g === "v" ? ` vMerge="1"` : g === "hv" ? ` hMerge="1" vMerge="1"` : "";
        return `<a:tc${attr}><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ja-JP" sz="${fontSz * 100}"/></a:p></a:txBody><a:tcPr anchor="ctr">${borders}</a:tcPr></a:tc>`;
      }
      const col = cols[ci] ?? {};
      const st: StyleSpec = Object.assign({ size: fontSz }, deck.spec.defaults ?? {}, el.font ? { font: el.font } : {},
        isHeader ? (el.headerColor ? { color: el.headerColor } : {}) : (el.color ? { color: el.color } : {}),
        isHeader ? {} : col, isHeader ? { bold: true } : {}, g.st);
      if (st.align === undefined && (numericCol[ci] && !isHeader)) st.align = "right";
      if (st.align === undefined && isHeader && numericCol[ci]) st.align = "right";
      const fill = g.st.fill !== undefined ? fillXml(ctx, g.st.fill)
        : isHeader && el.headerFill !== undefined ? fillXml(ctx, el.headerFill)
          : !isHeader && el.bandFill !== undefined && bodyIndex % 2 === 1 ? fillXml(ctx, el.bandFill)
            : !isHeader && el.fill !== undefined ? fillXml(ctx, el.fill) : "";
      const paras = g.text.split(/\r?\n/).map(line => {
        const runs = parseInline(line).map(seg => runXml(ctx, seg.text, Object.assign({}, st, seg.st))).join("");
        return `<a:p>${alignVal(st.align) ? `<a:pPr algn="${alignVal(st.align)}"/>` : ""}${runs || `<a:endParaRPr lang="ja-JP" sz="${Math.round((st.size ?? fontSz) * 100)}"/>`}</a:p>`;
      }).join("");
      const span = (g.colspan > 1 ? ` gridSpan="${g.colspan}"` : "") + (g.rowspan > 1 ? ` rowSpan="${g.rowspan}"` : "");
      const anchor = anchorVal(st.valign) ?? "ctr";
      return `<a:tc${span}><a:txBody><a:bodyPr/><a:lstStyle/>${paras}</a:txBody><a:tcPr marL="91440" marR="91440" marT="45720" marB="45720" anchor="${anchor}">${borders}${fill}</a:tcPr></a:tc>`;
    }).join("");
    return `<a:tr h="${rowH}">${tcs}</a:tr>`;
  }).join("");
  const id = ctx.nextId++;
  const tw = colWs.reduce((a, b) => a + b, 0);
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(el.name ?? `Table ${id}`)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${r.x}" y="${r.y}"/><a:ext cx="${tw}" cy="${rowH * rowCount}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
    `<a:tbl><a:tblPr${header ? ` firstRow="1"` : ""}${el.bandRow === false ? "" : ` bandRow="1"`}><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr>` +
    `<a:tblGrid>${colWs.map(w => `<a:gridCol w="${w}"/>`).join("")}</a:tblGrid>${rowsXml}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

// ---------- グラフ ----------
function chartFrameXml(ctx: PartCtx, el: ElementSpec, r: Rect): string {
  const deck = ctx.deck;
  const pkg = deck.pkg;
  const data = chartData(el);
  if (data.cats.length === 0 || data.series.length === 0) { deck.warnings.push("chart のデータが空です"); return ""; }
  const part = nextPartName(pkg, "ppt/charts/chart", ".xml");
  const xlsx = nextPartName(pkg, "ppt/embeddings/Microsoft_Excel_Worksheet", ".xlsx");
  setBytes(pkg, xlsx, chartWorkbook(data.cats, data.series));
  pkg.ctDefaults.xlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const xr = addRel(pkg, part, R_PACKAGE, relTarget(part, xlsx));
  setText(pkg, part, chartXml(el, data.cats, data.series, xr, deck.scale), "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
  const rid = addRel(pkg, ctx.part, R_CHART, relTarget(ctx.part, part));
  const id = ctx.nextId++;
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(el.name ?? `Chart ${id}`)}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${r.x}" y="${r.y}"/><a:ext cx="${r.w}" cy="${r.h}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
    `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}

interface ChartSeries { name: string; values: (number | null)[]; color?: string }

function plainValue(v: CellValue | undefined): string | number | null {
  if (v === null || v === undefined) return null;
  if (typeof v === "object") { const t = (v as CellObj).text; return t === undefined ? null : t; }
  if (typeof v === "boolean") return v ? 1 : 0;
  return v;
}

function chartData(el: ElementSpec): { cats: string[]; series: ChartSeries[] } {
  if (el.data && el.data.length > 0) {
    const keys = Object.keys(el.data[0]);
    const catKey = el.categoryKey ?? keys[0];
    const valKeys = el.valueKeys ?? keys.filter(k => k !== catKey && (el.data ?? []).every(row => { const p = plainValue(row[k]); return p === null || typeof p === "number"; }));
    return {
      cats: el.data.map(row => String(plainValue(row[catKey]) ?? "")),
      series: valKeys.map((k, i) => ({ name: k, values: (el.data ?? []).map(row => { const p = plainValue(row[k]); return typeof p === "number" ? p : null; }), color: el.colors?.[i] })),
    };
  }
  return {
    cats: (el.categories ?? []).map(c => String(c)),
    series: (el.series ?? []).map((s, i) => ({ name: s.name ?? `系列${i + 1}`, values: s.values ?? [], color: s.color ?? el.colors?.[i] })),
  };
}

function range(n: number): number[] {
  const out: number[] = [];
  for (let i = 0; i < n; i++) out.push(i);
  return out;
}

function colName(i: number): string {
  let s = "";
  for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}

function chartWorkbook(cats: string[], series: ChartSeries[]): Uint8Array {
  const cell = (ref: string, v: string | number | null): string =>
    v === null ? "" : typeof v === "number" ? `<c r="${ref}"><v>${v}</v></c>` : `<c r="${ref}" t="inlineStr"><is><t>${esc(v)}</t></is></c>`;
  let rows = `<row r="1">${cell("A1", "")}${series.map((s, i) => cell(`${colName(i + 1)}1`, s.name)).join("")}</row>`;
  cats.forEach((c, ri) => {
    rows += `<row r="${ri + 2}">${cell(`A${ri + 2}`, c)}${series.map((s, i) => cell(`${colName(i + 1)}${ri + 2}`, s.values[ri] ?? null)).join("")}</row>`;
  });
  const S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const files: ZipOut[] = [
    { name: "[Content_Types].xml", data: utf8(XML_HEAD + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`) },
    { name: "_rels/.rels", data: utf8(relsXml([{ id: "rId1", type: `${REL}/officeDocument`, target: "xl/workbook.xml", external: false }])) },
    { name: "xl/workbook.xml", data: utf8(XML_HEAD + `<workbook xmlns="${S}" xmlns:r="${REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`) },
    { name: "xl/_rels/workbook.xml.rels", data: utf8(relsXml([{ id: "rId1", type: `${REL}/worksheet`, target: "worksheets/sheet1.xml", external: false }])) },
    { name: "xl/worksheets/sheet1.xml", data: utf8(XML_HEAD + `<worksheet xmlns="${S}"><sheetData>${rows}</sheetData></worksheet>`) },
  ];
  return zipWrite(files);
}

function chartXml(el: ElementSpec, cats: string[], series: ChartSeries[], xlsxRid: string, sc: number): string {
  const type = (el.chartType ?? "column").toLowerCase();
  const isPie = type === "pie" || type === "doughnut";
  const n = cats.length;
  const fmt = el.numberFormat ?? "General";
  const palette = (i: number): string => colorXml(el.colors?.[i] ?? `accent${(i % 6) + 1}`);
  const strCache = (vals: string[]): string => `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join("")}</c:strCache>`;
  const catRef = `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${n + 1}</c:f>${strCache(cats)}</c:strRef></c:cat>`;
  const valRef = (s: ChartSeries, i: number): string => {
    const col = colName(i + 1);
    return `<c:val><c:numRef><c:f>Sheet1!$${col}$2:$${col}$${n + 1}</c:f><c:numCache><c:formatCode>${esc(fmt)}</c:formatCode><c:ptCount val="${n}"/>` +
      s.values.slice(0, n).map((v, j) => v === null || v === undefined ? "" : `<c:pt idx="${j}"><c:v>${v}</c:v></c:pt>`).join("") + `</c:numCache></c:numRef></c:val>`;
  };
  const txRef = (s: ChartSeries, i: number): string => `<c:tx><c:strRef><c:f>Sheet1!$${colName(i + 1)}$1</c:f>${strCache([s.name])}</c:strRef></c:tx>`;
  const showLabels = el.dataLabels === true || typeof el.dataLabels === "string";
  const labelPct = el.dataLabels === "percent";
  const dLbls = showLabels
    ? `<c:dLbls><c:numFmt formatCode="${esc(labelPct ? "0%" : fmt)}" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>` +
    (isPie ? `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr b="1"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr>` : "") +
    `<c:showLegendKey val="0"/><c:showVal val="${labelPct ? 0 : 1}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${labelPct ? 1 : 0}"/><c:showBubbleSize val="0"/>${isPie ? `<c:showLeaderLines val="1"/>` : ""}</c:dLbls>`
    : "";
  const stacked = el.stacked === "percent" ? "percentStacked" : el.stacked ? "stacked" : "";
  const fillSp = (i: number, s: ChartSeries): string => `<c:spPr><a:solidFill>${s.color ? colorXml(s.color) : palette(i)}</a:solidFill></c:spPr>`;
  let plot = "";
  const AX1 = 50010001, AX2 = 50010002;
  if (type === "column" || type === "bar") {
    plot = `<c:barChart><c:barDir val="${type === "bar" ? "bar" : "col"}"/><c:grouping val="${stacked || "clustered"}"/><c:varyColors val="0"/>` +
      series.map((s, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${txRef(s, i)}${fillSp(i, s)}<c:invertIfNegative val="0"/>${dLbls}${catRef}${valRef(s, i)}</c:ser>`).join("") +
      `<c:gapWidth val="80"/>${stacked ? `<c:overlap val="100"/>` : ""}<c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:barChart>`;
  } else if (type === "line") {
    const marker = (clr: string): string => el.markers === false ? `<c:marker><c:symbol val="none"/></c:marker>`
      : `<c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill>${clr}</a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:marker>`;
    plot = `<c:lineChart><c:grouping val="${stacked || "standard"}"/><c:varyColors val="0"/>` +
      series.map((s, i) => {
        const clr = s.color ? colorXml(s.color) : palette(i);
        return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${txRef(s, i)}<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill>${clr}</a:solidFill><a:round/></a:ln></c:spPr>` +
          `${marker(clr)}${dLbls}${catRef}${valRef(s, i)}<c:smooth val="${el.smooth ? 1 : 0}"/></c:ser>`;
      }).join("") +
      `<c:marker val="1"/><c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:lineChart>`;
  } else if (type === "area") {
    plot = `<c:areaChart><c:grouping val="${stacked || "standard"}"/><c:varyColors val="0"/>` +
      series.map((s, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${txRef(s, i)}${fillSp(i, s)}${dLbls}${catRef}${valRef(s, i)}</c:ser>`).join("") +
      `<c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:areaChart>`;
  } else if (isPie) {
    const s = series[0];
    const dpts = cats.map((_, j) => `<c:dPt><c:idx val="${j}"/><c:bubble3D val="0"/><c:spPr><a:solidFill>${palette(j)}</a:solidFill><a:ln w="19050"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join("");
    const ser = `<c:ser><c:idx val="0"/><c:order val="0"/>${txRef(s, 0)}${dpts}${dLbls}${catRef}${valRef(s, 0)}</c:ser>`;
    plot = type === "pie"
      ? `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`
      : `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/><c:holeSize val="${Math.min(90, Math.max(10, el.hole ?? 55))}"/></c:doughnutChart>`;
  } else {
    throw new Error(`chartType「${el.chartType}」は未対応です (column / bar / line / area / pie / doughnut)`);
  }
  const gray = (c: string): string => `<a:ln w="9525"><a:solidFill><a:srgbClr val="${c}"/></a:solidFill></a:ln>`;
  const isBar = type === "bar";
  const axes = isPie ? "" :
    `<c:catAx><c:axId val="${AX1}"/><c:scaling><c:orientation val="${isBar ? "maxMin" : "minMax"}"/></c:scaling><c:delete val="0"/><c:axPos val="${isBar ? "l" : "b"}"/>` +
    `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr>${gray("BFBFBF")}</c:spPr>` +
    `<c:crossAx val="${AX2}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
    `<c:valAx><c:axId val="${AX2}"/><c:scaling><c:orientation val="minMax"/>${el.axisMax !== undefined ? `<c:max val="${el.axisMax}"/>` : ""}${el.axisMin !== undefined ? `<c:min val="${el.axisMin}"/>` : ""}</c:scaling>` +
    `<c:delete val="0"/><c:axPos val="${isBar ? "t" : "l"}"/>${el.gridlines === false ? "" : `<c:majorGridlines><c:spPr>${gray("E5E5E5")}</c:spPr></c:majorGridlines>`}` +
    `<c:numFmt formatCode="${esc(stacked === "percentStacked" ? "0%" : fmt)}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>` +
    `<c:crossAx val="${AX1}"/><c:crosses val="${isBar ? "max" : "autoZero"}"/><c:crossBetween val="between"/></c:valAx>`;
  const legendOn = el.legend === undefined ? (series.length > 1 || isPie) : el.legend !== false && el.legend !== "none";
  const lp: { [k: string]: string } = { bottom: "b", top: "t", left: "l", right: "r" };
  const legend = legendOn ? `<c:legend><c:legendPos val="${lp[String(el.legend).toLowerCase()] ?? "b"}"/><c:overlay val="0"/></c:legend>` : "";
  const fs = Math.round((el.size ?? 12) * 100 * (el.size !== undefined ? 1 : sc));
  const title = el.title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${Math.round(fs * 1.2)}" b="1"/></a:pPr><a:r><a:rPr lang="ja-JP" altLang="en-US" sz="${Math.round(fs * 1.2)}" b="1"/><a:t>${esc(el.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : `<c:autoTitleDeleted val="1"/>`;
  return XML_HEAD + `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL}">` +
    `<c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}${axes}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>` +
    `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${fs}"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr>` +
    `<c:externalData r:id="${xlsxRid}"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`;
}

// =====================================================================
// テキスト (段落・ラン・箇条書き・自動縮小)
// =====================================================================
interface TextMode { placeholder: boolean; defaultSize: (level: number) => number }
interface FlatPara { st: StyleSpec; level: number; inList: boolean; runs: { text: string; st: StyleSpec }[] }

const TEXT_KEYS = ["font", "fontEa", "size", "color", "bold", "italic", "underline", "strike", "highlight", "align",
  "lineSpacing", "spaceBefore", "spaceAfter", "bullet", "bulletColor", "link"];

function textStyleOf(n: StyleSpec): StyleSpec {
  const o: StyleSpec = {};
  const src = n as { [k: string]: string | number | boolean | FillSpec | LineSpec | Len[] | undefined };
  const dst = o as { [k: string]: string | number | boolean | FillSpec | LineSpec | Len[] | undefined };
  for (const k of TEXT_KEYS) if (src[k] !== undefined) dst[k] = src[k];
  return o;
}

function parseInline(s: string): { text: string; st: StyleSpec }[] {
  const out: { text: string; st: StyleSpec }[] = [];
  const re = /\*\*([^*]+)\*\*|\[([^\]]+)\]\(([^)\s]+)\)/g;
  let last = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(s)) !== null) {
    if (m.index > last) out.push({ text: s.slice(last, m.index), st: {} });
    if (m[1] !== undefined) out.push({ text: m[1], st: { bold: true } });
    else out.push({ text: m[2], st: { link: m[3] } });
    last = m.index + m[0].length;
  }
  if (last < s.length) out.push({ text: s.slice(last), st: {} });
  return out;
}

function flattenText(v: TextValue | undefined, level: number, inherited: StyleSpec, inList: boolean, out: FlatPara[]): void {
  if (v === undefined || v === null) return;
  if (typeof v === "string" || typeof v === "number") {
    for (const line of String(v).split(/\r?\n/)) out.push({ st: inherited, level, inList, runs: parseInline(line) });
    return;
  }
  if (Array.isArray(v)) {
    for (const item of v) {
      if (Array.isArray(item)) flattenText(item, level + 1, inherited, true, out);
      else flattenText(item, level, inherited, true, out);
    }
    return;
  }
  const node = v as TextNode;
  const st = Object.assign({}, inherited, textStyleOf(node));
  const lvl = node.level ?? level;
  if (node.runs) {
    out.push({
      st, level: lvl, inList, runs: node.runs.map(r => typeof r === "string" ? { text: r, st: {} } : { text: String(r.text ?? ""), st: textStyleOf(r) }),
    });
  } else if (Array.isArray(node.text)) {
    flattenText(node.text, lvl, st, true, out);
  } else if (node.text !== undefined) {
    flattenText(String(node.text), lvl, st, inList, out);
  }
}

function buildParas(ctx: PartCtx, v: TextValue | undefined, base: StyleSpec, mode: TextMode): ParaOut[] {
  const flat: FlatPara[] = [];
  const baseText = textStyleOf(base);
  if (typeof v === "object" && v !== null && !Array.isArray(v)) {
    const node = v as TextNode;
    flattenText(node.runs ? { runs: node.runs, level: node.level } : node.text, node.level ?? 0, baseText, false, flat);
  } else {
    flattenText(v, 0, baseText, false, flat);
  }
  const sc = ctx.deck.scale;
  return flat.map(p => {
    const st = p.st;
    const size = st.size ?? mode.defaultSize(p.level);
    const lvl = Math.max(0, Math.min(8, p.level));
    let attrs = lvl > 0 ? ` lvl="${lvl}"` : "";
    let bu = "";
    let indent = 0;
    const explicit = st.bullet !== undefined && st.bullet !== false;
    if (mode.placeholder) {
      if (st.bullet === false || (!explicit && !p.inList)) { attrs += ` marL="0" indent="0"`; bu = "<a:buNone/>"; }
      else indent = Math.round((lvl * 342900 + 285750) * sc);
    } else if (!explicit && lvl > 0) {
      indent = Math.round(lvl * 342900 * sc);
      attrs += ` marL="${indent}" indent="0"`;
    }
    if (explicit) {
      indent = Math.round((lvl * 342900 + 285750) * sc);
      attrs += ` marL="${indent}" indent="${-Math.round(285750 * sc)}"`;
      const clr = st.bulletColor ? `<a:buClr>${colorXml(st.bulletColor)}</a:buClr>` : "";
      if (st.bullet === "number" || st.bullet === "1.") bu = `${clr}<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>`;
      else {
        const ch = typeof st.bullet === "string" && st.bullet.length <= 2 ? st.bullet : ["•", "–", "▪", "–", "▪"][lvl % 5];
        bu = `${clr}<a:buFont typeface="Arial"/><a:buChar char="${esc(ch)}"/>`;
      }
    }
    const algn = alignVal(st.align);
    if (algn) attrs += ` algn="${algn}"`;
    const lnSpc = st.lineSpacing !== undefined ? `<a:lnSpc><a:spcPct val="${Math.round(st.lineSpacing * 100000)}"/></a:lnSpc>` : "";
    const spcBef = st.spaceBefore !== undefined ? `<a:spcBef><a:spcPts val="${Math.round(st.spaceBefore * 100)}"/></a:spcBef>` : "";
    const spcAft = st.spaceAfter !== undefined ? `<a:spcAft><a:spcPts val="${Math.round(st.spaceAfter * 100)}"/></a:spcAft>` : "";
    const inner = lnSpc + spcBef + spcAft + bu;
    const pPr = attrs || inner ? `<a:pPr${attrs}>${inner}</a:pPr>` : "";
    const runs = p.runs.filter(r => r.text !== "").map(r => runXml(ctx, r.text, Object.assign({}, st, r.st))).join("");
    const end = `<a:endParaRPr lang="ja-JP" altLang="en-US"${st.size !== undefined ? ` sz="${Math.round(st.size * 100)}"` : ""}/>`;
    const plain = p.runs.map(r => r.text).join("");
    return {
      xml: `<a:p>${pPr}${runs || end}</a:p>`, plain, size, indent,
      spcBef: st.spaceBefore ?? (mode.placeholder && p.inList ? 6 : 0), lnSpc: st.lineSpacing ?? (mode.placeholder && p.inList ? 1.1 : 1.0),
    };
  });
}

function parasXml(paras: ParaOut[]): string {
  return paras.length ? paras.map(p => p.xml).join("") : `<a:p><a:endParaRPr lang="ja-JP" altLang="en-US"/></a:p>`;
}

function runXml(ctx: PartCtx, text: string, st: StyleSpec): string {
  return `<a:r>${rPrXml(ctx, st)}<a:t>${esc(text)}</a:t></a:r>`;
}

function rPrXml(ctx: PartCtx, st: StyleSpec): string {
  let a = ` lang="ja-JP" altLang="en-US"`;
  if (st.size !== undefined) a += ` sz="${Math.round(st.size * 100)}"`;
  if (st.bold !== undefined) a += ` b="${st.bold ? 1 : 0}"`;
  if (st.italic !== undefined) a += ` i="${st.italic ? 1 : 0}"`;
  if (st.underline) a += ` u="sng"`;
  if (st.strike) a += ` strike="sngStrike"`;
  const fill = st.color ? `<a:solidFill>${colorXml(st.color)}</a:solidFill>` : "";
  const hl = st.highlight ? `<a:highlight>${colorXml(st.highlight)}</a:highlight>` : "";
  const link = st.link ? hlinkXml(ctx, st.link) : "";
  return `<a:rPr${a}>${fill}${hl}${fontXml(st.font, st.fontEa, "")}${link}</a:rPr>`;
}

// kind: "major" / "minor" はフォント未指定時にテーマのフォントを明示、"" は未指定なら何も出力しない
function fontXml(font: string | undefined, fontEa: string | undefined, kind: string): string {
  const t = kind === "major" ? "mj" : "mn";
  if (font === undefined && fontEa === undefined) return kind ? `<a:latin typeface="+${t}-lt"/><a:ea typeface="+${t}-ea"/><a:cs typeface="+${t}-cs"/>` : "";
  const lat = font ?? `+${t}-lt`;
  const ea = fontEa ?? font ?? `+${t}-ea`;
  return `<a:latin typeface="${esc(lat)}"/><a:ea typeface="${esc(ea)}"/><a:cs typeface="${esc(lat)}"/>`;
}

function alignVal(a: string | undefined): string | undefined {
  if (!a) return undefined;
  const m: { [k: string]: string } = { left: "l", l: "l", center: "ctr", centre: "ctr", ctr: "ctr", right: "r", r: "r", justify: "just", just: "just", distributed: "dist", dist: "dist" };
  return m[a.toLowerCase()];
}

function anchorVal(v: string | undefined): string | undefined {
  if (!v) return undefined;
  const m: { [k: string]: string } = { top: "t", t: "t", middle: "ctr", center: "ctr", ctr: "ctr", bottom: "b", b: "b" };
  return m[v.toLowerCase()];
}

function marginOf(deck: DeckCtx, m: Len | Len[] | undefined): number[] {
  if (m === undefined) return [91440, 45720, 91440, 45720];
  const arr = Array.isArray(m) ? m : [m];
  const v = arr.map((x, i) => lenEmu(deck, x, i % 2 === 0 ? "y" : "x") ?? 0);
  if (v.length === 1) return [v[0], v[0], v[0], v[0]];
  if (v.length === 2) return [v[1], v[0], v[1], v[0]];
  return [v[0], v[1], v[2] ?? v[0], v[3] ?? v[1]];
}

function textWidth(s: string): number {
  let w = 0;
  for (let i = 0; i < s.length; i++) w += s.charCodeAt(i) < 0x80 ? 1 : 2;
  return w;
}

// 文字量からおおよそのフォント縮小率 (%) を推定する。PowerPoint は保存済みの fontScale で描画するため事前計算が必要
function fitScale(paras: ParaOut[], w: number, h: number): number {
  if (paras.length === 0 || w <= 0 || h <= 0) return 100;
  for (let s = 100; s >= 40; s -= 5) {
    let total = 0;
    for (const p of paras) {
      const sz = p.size * 12700 * s / 100;
      const avail = Math.max(w - p.indent, w * 0.3);
      const lines = Math.max(1, Math.ceil(textWidth(p.plain) * 0.52 * sz / avail));
      total += lines * sz * 1.2 * p.lnSpc + p.spcBef * 12700 * s / 100;
    }
    if (total <= h * 0.95) return s;
  }
  return 40;
}

function bodyPrXml(deck: DeckCtx, st: ElementSpec, r: Rect, paras: ParaOut[], defaultAnchor: string | undefined, isPh: boolean): string {
  let a = "";
  if (st.wrap === false) a += ` wrap="none"`;
  else if (!isPh) a += ` wrap="square"`;
  const mg = st.margin !== undefined ? marginOf(deck, st.margin) : null;
  if (mg) a += ` lIns="${mg[0]}" tIns="${mg[1]}" rIns="${mg[2]}" bIns="${mg[3]}"`;
  const anchor = anchorVal(st.valign) ?? defaultAnchor;
  if (anchor) a += ` anchor="${anchor}"`;
  if (!isPh) a += ` rtlCol="0"`;
  const ins = mg ?? [91440, 45720, 91440, 45720];
  const mode = (st.autofit ?? "shrink").toLowerCase();
  let fit = "";
  if (mode === "resize") fit = "<a:spAutoFit/>";
  else if (mode === "none") fit = isPh ? "<a:noAutofit/>" : "";
  else {
    const s = fitScale(paras, r.w - ins[0] - ins[2], r.h - ins[1] - ins[3]);
    fit = s < 100 ? `<a:normAutofit fontScale="${s * 1000}"${s <= 85 ? ` lnSpcReduction="${s <= 70 ? 20000 : 10000}"` : ""}/>` : (isPh ? "" : "<a:normAutofit/>");
  }
  return `<a:bodyPr${a}>${fit}</a:bodyPr>`;
}

// =====================================================================
// 色・塗り・線・背景
// =====================================================================
function hexOf(c: string): string {
  let s = c.trim().replace(/^#/, "");
  if (NAMED_COLORS[s.toLowerCase()]) s = NAMED_COLORS[s.toLowerCase()];
  if (/^[0-9a-f]{3}$/i.test(s)) s = s.split("").map(ch => ch + ch).join("");
  if (!/^[0-9a-f]{6}$/i.test(s)) throw new Error(`色の指定「${c}」が不正です (例: "1F4E79" / "#1F4E79" / "accent1" / "accent1+40")`);
  return s.toUpperCase();
}

function colorXml(c: string, transparency?: number): string {
  const alpha = transparency !== undefined && transparency > 0 ? `<a:alpha val="${Math.round((100 - transparency) * 1000)}"/>` : "";
  const m = /^\s*([a-zA-Z]+\d?)\s*([+-]\d{1,2})?\s*$/.exec(c);
  if (m && SCHEME_COLORS.map(x => x.toLowerCase()).indexOf(m[1].toLowerCase()) >= 0) {
    const name = SCHEME_COLORS.filter(x => x.toLowerCase() === m[1].toLowerCase())[0];
    let mods = "";
    if (m[2]) {
      const v = +m[2];
      mods = v > 0 ? `<a:lumMod val="${(100 - v) * 1000}"/><a:lumOff val="${v * 1000}"/>` : `<a:lumMod val="${(100 + v) * 1000}"/>`;
    }
    return `<a:schemeClr val="${name}">${mods}${alpha}</a:schemeClr>`;
  }
  return `<a:srgbClr val="${hexOf(c)}">${alpha}</a:srgbClr>`;
}

function autoTextColor(fill: FillSpec | undefined): string | undefined {
  const c = fill === undefined ? "accent1" : typeof fill === "string" ? fill : fill.color ?? (fill.gradient ? fill.gradient[0] : undefined);
  if (!c || c === "none") return undefined;
  const name = c.replace(/[+-]\d+$/, "").toLowerCase();
  if (["lt1", "lt2", "bg1", "bg2"].indexOf(name) >= 0) return "tx1";
  if (SCHEME_COLORS.map(x => x.toLowerCase()).indexOf(name) >= 0) return /\+[4-9]\d$/.test(c) ? "tx1" : "lt1";
  try {
    const hx = hexOf(c);
    const lum = (0.299 * parseInt(hx.slice(0, 2), 16) + 0.587 * parseInt(hx.slice(2, 4), 16) + 0.114 * parseInt(hx.slice(4, 6), 16)) / 255;
    return lum > 0.6 ? "tx1" : "lt1";
  } catch (e) {
    return undefined;
  }
}

function fillXml(ctx: PartCtx, f: FillSpec): string {
  if (typeof f === "string") return f.toLowerCase() === "none" ? "<a:noFill/>" : `<a:solidFill>${colorXml(f)}</a:solidFill>`;
  if (f.image) {
    const img = embedImage(ctx, f.image);
    return `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${img.rid}">${f.transparency ? `<a:alphaModFix amt="${Math.round((100 - f.transparency) * 1000)}"/>` : ""}</a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;
  }
  if (f.gradient && f.gradient.length > 0) {
    const g = f.gradient.length === 1 ? [f.gradient[0], f.gradient[0]] : f.gradient;
    const stops = g.map((c, i) => `<a:gs pos="${Math.round(i / (g.length - 1) * 100000)}">${colorXml(c, f.transparency)}</a:gs>`).join("");
    return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:lin ang="${Math.round((f.angle ?? 90) * 60000)}" scaled="0"/></a:gradFill>`;
  }
  if (f.color) return f.color.toLowerCase() === "none" ? "<a:noFill/>" : `<a:solidFill>${colorXml(f.color, f.transparency)}</a:solidFill>`;
  return "";
}

function lineXml(l: LineSpec, ends = ""): string {
  if (typeof l === "string") {
    if (l.toLowerCase() === "none") return "<a:ln><a:noFill/></a:ln>";
    return `<a:ln w="12700"><a:solidFill>${colorXml(l)}</a:solidFill>${ends}</a:ln>`;
  }
  const dashes: { [k: string]: string } = { solid: "solid", dash: "dash", dot: "sysDot", dashdot: "dashDot", longdash: "lgDash", sysdash: "sysDash" };
  const dash = l.dash ? `<a:prstDash val="${dashes[l.dash.toLowerCase()] ?? "dash"}"/>` : "";
  if (l.color && l.color.toLowerCase() === "none") return "<a:ln><a:noFill/></a:ln>";
  return `<a:ln w="${Math.round((l.width ?? 1) * 12700)}"><a:solidFill>${colorXml(l.color ?? "tx1", l.transparency)}</a:solidFill>${dash}${ends}</a:ln>`;
}

function bgXml(ctx: PartCtx, f: FillSpec): string {
  const fill = fillXml(ctx, f);
  return fill ? `<p:bg><p:bgPr>${fill}<a:effectLst/></p:bgPr></p:bg>` : "";
}

function replaceBg(xml: string, bg: string): string {
  if (!bg) return xml;
  if (/<p:bg>[\s\S]*?<\/p:bg>/.test(xml)) return xml.replace(/<p:bg>[\s\S]*?<\/p:bg>/, bg);
  return xml.replace(/(<p:cSld\b[^>]*>)/, `$1${bg}`);
}

// =====================================================================
// XML ユーティリティ
// =====================================================================
function spTreeOpen(): string {
  return `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`;
}

function insertShapes(xml: string, back: string, front: string): string {
  if (back) xml = xml.replace(/(<p:grpSpPr\s*\/>|<p:grpSpPr>[\s\S]*?<\/p:grpSpPr>)/, `$1${back}`);
  if (front) {
    const i = xml.lastIndexOf("</p:spTree>");
    if (i >= 0) xml = xml.slice(0, i) + front + xml.slice(i);
  }
  return xml;
}

function maxShapeId(xml: string): number {
  let max = 1;
  for (const m of allMatches(/<p:cNvPr\b[^>]*\bid="(\d+)"/g, xml)) max = Math.max(max, +m[1]);
  return max;
}

function allMatches(re: RegExp, s: string): RegExpExecArray[] {
  const out: RegExpExecArray[] = [];
  const g = new RegExp(re.source, re.flags.indexOf("g") >= 0 ? re.flags : re.flags + "g");
  let m: RegExpExecArray | null;
  while ((m = g.exec(s)) !== null) {
    out.push(m);
    if (m[0] === "") g.lastIndex++;
  }
  return out;
}

function attrOf(tag: string, name: string): string | undefined {
  const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag);
  return m ? m[1] : undefined;
}

function setAttr(tag: string, name: string, value: string): string {
  if (new RegExp(`\\s${name}="[^"]*"`).test(tag)) return tag.replace(new RegExp(`\\s${name}="[^"]*"`), ` ${name}="${value}"`);
  return tag.replace(/\s*(\/?)>$/, ` ${name}="${value}"$1>`);
}

function esc(s: string): string {
  return String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function unesc(s: string): string {
  return s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, "\"").replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (m: string, h: string) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (m: string, d: string) => String.fromCodePoint(+d)).replace(/&amp;/g, "&");
}

function norm(s: string): string {
  return String(s).toLowerCase().replace(/[\s_\-・]/g, "");
}

// =====================================================================
// パッケージ (OPC) 操作
// =====================================================================
function newPkg(): Pkg {
  return {
    parts: {}, rels: {},
    ctDefaults: { rels: "application/vnd.openxmlformats-package.relationships+xml", xml: "application/xml" },
    ctOverrides: {},
  };
}

function openTemplate(bytes: Uint8Array): Pkg {
  const pkg = newPkg();
  const entries = unzip(bytes);
  for (const name of Object.keys(entries)) pkg.parts[name] = { raw: entries[name], dirty: false };
  const ct = pkg.parts["[Content_Types].xml"];
  if (!ct) throw new Error("テンプレートに [Content_Types].xml がありません (PPTX/POTX ではない可能性があります)");
  const ctXml = getText(pkg, "[Content_Types].xml");
  for (const m of allMatches(/<Default\b[^>]*>/g, ctXml)) {
    const e = attrOf(m[0], "Extension"), c = attrOf(m[0], "ContentType");
    if (e && c) pkg.ctDefaults[e.toLowerCase()] = c;
  }
  for (const m of allMatches(/<Override\b[^>]*>/g, ctXml)) {
    const p = attrOf(m[0], "PartName"), c = attrOf(m[0], "ContentType");
    if (p && c) pkg.ctOverrides[p] = c.replace(".template.main+xml", ".presentation.main+xml").replace(".slideshow.main+xml", ".presentation.main+xml");
  }
  delete pkg.parts["[Content_Types].xml"];
  return pkg;
}

function getBytes(pkg: Pkg, name: string): Uint8Array {
  const p = pkg.parts[name];
  if (!p) throw new Error(`パーツ ${name} がありません`);
  if (p.bytes) return p.bytes;
  if (p.text !== undefined) return utf8(p.text);
  const raw = p.raw as RawEntry;
  p.bytes = raw.method === 0 ? raw.comp : raw.method === 8 ? inflate(raw.comp, raw.size) : new Uint8Array(0);
  if (raw.method !== 0 && raw.method !== 8) throw new Error(`${name} の圧縮形式 (${raw.method}) は未対応です`);
  return p.bytes;
}

function getText(pkg: Pkg, name: string): string {
  const p = pkg.parts[name];
  if (!p) throw new Error(`パーツ ${name} がありません`);
  if (p.text === undefined) p.text = utf8Decode(getBytes(pkg, name));
  return p.text;
}

function setText(pkg: Pkg, name: string, text: string, contentType?: string): void {
  pkg.parts[name] = { text, dirty: true };
  if (contentType) pkg.ctOverrides["/" + name] = contentType;
}

function setBytes(pkg: Pkg, name: string, bytes: Uint8Array): void {
  pkg.parts[name] = { bytes, dirty: true };
}

function removePart(pkg: Pkg, name: string): void {
  delete pkg.parts[name];
  delete pkg.ctOverrides["/" + name];
  delete pkg.rels[name];
  delete pkg.parts[relsName(name)];
}

function nextPartName(pkg: Pkg, prefix: string, ext: string): string {
  let n = 1;
  while (pkg.parts[`${prefix}${n}${ext}`]) n++;
  const name = `${prefix}${n}${ext}`;
  pkg.parts[name] = { text: "", dirty: true };
  return name;
}

function relsName(owner: string): string {
  if (owner === "") return "_rels/.rels";
  const i = owner.lastIndexOf("/");
  return owner.slice(0, i + 1) + "_rels/" + owner.slice(i + 1) + ".rels";
}

function getRels(pkg: Pkg, owner: string): Rel[] {
  if (pkg.rels[owner]) return pkg.rels[owner];
  const list: Rel[] = [];
  const rn = relsName(owner);
  if (pkg.parts[rn]) {
    for (const m of allMatches(/<Relationship\b[^>]*>/g, getText(pkg, rn))) {
      list.push({
        id: attrOf(m[0], "Id") ?? "", type: attrOf(m[0], "Type") ?? "", target: unesc(attrOf(m[0], "Target") ?? ""),
        external: attrOf(m[0], "TargetMode") === "External",
      });
    }
  }
  pkg.rels[owner] = list;
  return list;
}

function addRel(pkg: Pkg, owner: string, type: string, target: string, external = false, preferId?: string): string {
  const list = getRels(pkg, owner);
  const same = list.filter(r => r.type === type && r.target === target && r.external === external)[0];
  if (same && !preferId) return same.id;
  let id = preferId && !list.some(r => r.id === preferId) ? preferId : "";
  if (!id) {
    let n = list.length + 1;
    while (list.some(r => r.id === `rId${n}`)) n++;
    id = `rId${n}`;
  }
  list.push({ id, type, target, external });
  return id;
}

function relsXml(list: Rel[]): string {
  return XML_HEAD + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    list.map(r => `<Relationship Id="${esc(r.id)}" Type="${esc(r.type)}" Target="${esc(r.target)}"${r.external ? ` TargetMode="External"` : ""}/>`).join("") +
    `</Relationships>`;
}

function resolveTarget(from: string, target: string): string {
  if (target.charAt(0) === "/") return target.slice(1);
  const parts = from.split("/");
  parts.pop();
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== "." && seg !== "") parts.push(seg);
  }
  return parts.join("/");
}

function relTarget(from: string, to: string): string {
  const f = from.split("/");
  f.pop();
  const t = to.split("/");
  let i = 0;
  while (i < f.length && i < t.length - 1 && f[i] === t[i]) i++;
  let up = "";
  for (let k = i; k < f.length; k++) up += "../";
  return up + t.slice(i).join("/");
}

function finalizePresentation(deck: DeckCtx, orig: { id: number; part: string }[], built: NewSlide[], keep: boolean): void {
  const pkg = deck.pkg;
  const pres = deck.presPart;
  let xml = getText(pkg, pres);
  const presRels = getRels(pkg, pres);
  let nextId = Math.max(255, ...orig.map(o => o.id)) + 1;
  const entries: { id: number; rid: string; section?: string }[] = [];
  if (keep) {
    for (const o of orig) {
      const rel = presRels.filter(r => r.type === R_SLIDE && resolveTarget(pres, r.target) === o.part)[0];
      if (rel) entries.push({ id: o.id, rid: rel.id });
    }
  } else {
    for (const o of orig) {
      const rels = getRels(pkg, o.part);
      for (const r of rels) if (r.type === R_NOTES) removePart(pkg, resolveTarget(o.part, r.target));
      removePart(pkg, o.part);
      for (let i = presRels.length - 1; i >= 0; i--) {
        if (presRels[i].type === R_SLIDE && resolveTarget(pres, presRels[i].target) === o.part) presRels.splice(i, 1);
      }
    }
  }
  for (const b of built) entries.push({ id: nextId++, rid: addRel(pkg, pres, R_SLIDE, relTarget(pres, b.part)), section: b.section });

  xml = xml.replace(/<p:sldIdLst\s*\/>|<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, "");
  xml = xml.replace(/<p:custShowLst>[\s\S]*?<\/p:custShowLst>/, "");
  xml = xml.replace(/<p:ext uri="\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\}">[\s\S]*?<\/p:ext>/, "");
  xml = xml.replace(/<p:extLst>\s*<\/p:extLst>/, "");
  const sld = `<p:sldIdLst>${entries.map(e => `<p:sldId id="${e.id}" r:id="${e.rid}"/>`).join("")}</p:sldIdLst>`;
  xml = xml.replace(/<p:sldSz\b/, `${sld}<p:sldSz`);
  if (deck.needNotesMaster && !/<p:notesMasterIdLst>/.test(xml)) {
    const rid = addRel(pkg, pres, R_NOTES_MASTER, relTarget(pres, deck.notesMaster));
    xml = xml.replace(/(<\/p:sldMasterIdLst>)/, `$1<p:notesMasterIdLst><p:notesMasterId r:id="${rid}"/></p:notesMasterIdLst>`);
  }
  if (!/<p:notesSz\b/.test(xml)) xml = xml.replace(/(<p:sldSz\b[^>]*\/>)/, `$1<p:notesSz cx="6858000" cy="9144000"/>`);
  if (entries.some(e => e.section)) {
    const groups: { name: string; ids: number[] }[] = [];
    for (const e of entries) {
      if (e.section || groups.length === 0) groups.push({ name: e.section ?? "既定のセクション", ids: [] });
      groups[groups.length - 1].ids.push(e.id);
    }
    const ext = `<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">` +
      groups.map(g => `<p14:section name="${esc(g.name)}" id="${guid()}"><p14:sldIdLst>${g.ids.map(id => `<p14:sldId id="${id}"/>`).join("")}</p14:sldIdLst></p14:section>`).join("") +
      `</p14:sectionLst></p:ext>`;
    if (/<p:extLst>/.test(xml)) xml = xml.replace(/<p:extLst>/, `<p:extLst>${ext}`);
    else xml = xml.replace(/<\/p:presentation>\s*$/, `<p:extLst>${ext}</p:extLst></p:presentation>`);
  }
  setText(pkg, pres, xml);

  // 文書プロパティ
  const props = deck.spec.properties;
  const core = getRels(pkg, "").filter(r => r.type === R_CORE)[0];
  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  if (!core) {
    setText(pkg, "docProps/core.xml", XML_HEAD + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `<dc:title>${esc(props?.title ?? "")}</dc:title><dc:subject>${esc(props?.subject ?? "")}</dc:subject><dc:creator>${esc(props?.author ?? "")}</dc:creator>` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
      "application/vnd.openxmlformats-package.core-properties+xml");
    addRel(pkg, "", R_CORE, "docProps/core.xml");
  } else if (props) {
    const cp = resolveTarget("", core.target);
    let cx = getText(pkg, cp);
    const put = (tag: string, v: string | undefined): void => {
      if (v === undefined) return;
      const re = new RegExp(`<${tag}\\s*/>|<${tag}>[\\s\\S]*?</${tag}>`);
      cx = re.test(cx) ? cx.replace(re, `<${tag}>${esc(v)}</${tag}>`) : cx.replace(/<\/cp:coreProperties>/, `<${tag}>${esc(v)}</${tag}></cp:coreProperties>`);
    };
    put("dc:title", props.title);
    put("dc:subject", props.subject);
    put("dc:creator", props.author);
    setText(pkg, cp, cx);
  }
}

function guid(): string {
  const h = (n: number): string => { let s = ""; for (let i = 0; i < n; i++) s += "0123456789ABCDEF".charAt(Math.floor(Math.random() * 16)); return s; };
  return `{${h(8)}-${h(4)}-4${h(3)}-A${h(3)}-${h(12)}}`;
}

interface ZipOut { name: string; data?: Uint8Array; raw?: RawEntry }

function writePackage(pkg: Pkg): Uint8Array {
  for (const owner of Object.keys(pkg.rels)) {
    if (owner !== "" && !pkg.parts[owner]) continue;
    const list = pkg.rels[owner];
    const rn = relsName(owner);
    if (list.length === 0) { delete pkg.parts[rn]; continue; }
    pkg.parts[rn] = { text: relsXml(list), dirty: true };
  }
  for (const name of Object.keys(pkg.parts)) {
    const p = pkg.parts[name];
    if (p.dirty && p.text === "" && !p.bytes) delete pkg.parts[name];
  }
  const names = Object.keys(pkg.parts);
  let ct = XML_HEAD + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`;
  for (const e of Object.keys(pkg.ctDefaults)) ct += `<Default Extension="${esc(e)}" ContentType="${esc(pkg.ctDefaults[e])}"/>`;
  for (const p of Object.keys(pkg.ctOverrides)) if (pkg.parts[p.slice(1)]) ct += `<Override PartName="${esc(p)}" ContentType="${esc(pkg.ctOverrides[p])}"/>`;
  ct += `</Types>`;
  const files: ZipOut[] = [{ name: "[Content_Types].xml", data: utf8(ct) }];
  const order = (n: string): number => n === "_rels/.rels" ? 0 : n === "docProps/core.xml" ? 1 : n.indexOf("ppt/presentation.xml") === 0 ? 2 : 3;
  names.sort((a, b) => order(a) - order(b));
  for (const name of names) {
    const p = pkg.parts[name];
    if (p.dirty || !p.raw) files.push({ name, data: p.text !== undefined && (p.dirty || !p.bytes) ? utf8(p.text) : p.bytes ?? new Uint8Array(0) });
    else files.push({ name, raw: p.raw });
  }
  return zipWrite(files);
}

// =====================================================================
// バイナリ処理 (UTF-8 / Base64 / CRC32 / ZIP 読み書き / Inflate)
// =====================================================================
function utf8(s: string): Uint8Array {
  const out = new Uint8Array(s.length * 3);
  let o = 0;
  for (let i = 0; i < s.length; i++) {
    let c = s.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff && i + 1 < s.length) c = 0x10000 + ((c - 0xd800) << 10) + (s.charCodeAt(++i) - 0xdc00);
    if (c < 0x80) out[o++] = c;
    else if (c < 0x800) { out[o++] = 0xc0 | (c >> 6); out[o++] = 0x80 | (c & 63); }
    else if (c < 0x10000) { out[o++] = 0xe0 | (c >> 12); out[o++] = 0x80 | ((c >> 6) & 63); out[o++] = 0x80 | (c & 63); }
    else { out[o++] = 0xf0 | (c >> 18); out[o++] = 0x80 | ((c >> 12) & 63); out[o++] = 0x80 | ((c >> 6) & 63); out[o++] = 0x80 | (c & 63); }
  }
  return out.slice(0, o);
}

function utf8Decode(b: Uint8Array): string {
  const parts: string[] = [];
  let chunk: number[] = [];
  let i = b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf ? 3 : 0;
  while (i < b.length) {
    const c = b[i++];
    let cp = c;
    if (c >= 0xf0) { cp = ((c & 7) << 18) | ((b[i] & 63) << 12) | ((b[i + 1] & 63) << 6) | (b[i + 2] & 63); i += 3; }
    else if (c >= 0xe0) { cp = ((c & 15) << 12) | ((b[i] & 63) << 6) | (b[i + 1] & 63); i += 2; }
    else if (c >= 0xc0) { cp = ((c & 31) << 6) | (b[i] & 63); i += 1; }
    if (cp >= 0x10000) { cp -= 0x10000; chunk.push(0xd800 + (cp >> 10), 0xdc00 + (cp & 1023)); }
    else chunk.push(cp);
    if (chunk.length >= 8192) { parts.push(String.fromCharCode.apply(null, chunk)); chunk = []; }
  }
  if (chunk.length) parts.push(String.fromCharCode.apply(null, chunk));
  return parts.join("");
}

const B64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

function base64Encode(b: Uint8Array): string {
  const parts: string[] = [];
  let s = "";
  for (let i = 0; i < b.length; i += 3) {
    const n = (b[i] << 16) | ((i + 1 < b.length ? b[i + 1] : 0) << 8) | (i + 2 < b.length ? b[i + 2] : 0);
    s += B64.charAt((n >> 18) & 63) + B64.charAt((n >> 12) & 63) +
      (i + 1 < b.length ? B64.charAt((n >> 6) & 63) : "=") + (i + 2 < b.length ? B64.charAt(n & 63) : "=");
    if (s.length >= 16384) { parts.push(s); s = ""; }
  }
  parts.push(s);
  return parts.join("");
}

function base64Decode(s: string): Uint8Array {
  const clean = s.replace(/^data:[^,]*,/, "").replace(/-/g, "+").replace(/_/g, "/").replace(/[^A-Za-z0-9+/]/g, "");
  const lookup = new Uint8Array(128);
  for (let i = 0; i < B64.length; i++) lookup[B64.charCodeAt(i)] = i;
  const out = new Uint8Array(Math.floor(clean.length * 3 / 4));
  let o = 0;
  for (let i = 0; i < clean.length; i += 4) {
    const a = lookup[clean.charCodeAt(i)], b = lookup[clean.charCodeAt(i + 1)];
    const c = i + 2 < clean.length ? lookup[clean.charCodeAt(i + 2)] : 0;
    const d = i + 3 < clean.length ? lookup[clean.charCodeAt(i + 3)] : 0;
    const n = (a << 18) | (b << 12) | (c << 6) | d;
    if (o < out.length) out[o++] = (n >> 16) & 255;
    if (i + 2 < clean.length && o < out.length) out[o++] = (n >> 8) & 255;
    if (i + 3 < clean.length && o < out.length) out[o++] = n & 255;
  }
  return out.slice(0, o);
}

const CRC_TABLE: number[] = (() => {
  const t: number[] = [];
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t.push(c >>> 0);
  }
  return t;
})();

function crc32(b: Uint8Array): number {
  let c = 0xffffffff;
  for (let i = 0; i < b.length; i++) c = CRC_TABLE[(c ^ b[i]) & 255] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function zipWrite(files: ZipOut[]): Uint8Array {
  const items = files.map(f => {
    const name = utf8(f.name);
    if (f.raw) return { name, method: f.raw.method, crc: f.raw.crc, size: f.raw.size, comp: f.raw.comp };
    const d = f.data ?? new Uint8Array(0);
    return { name, method: 0, crc: crc32(d), size: d.length, comp: d };
  });
  let total = 22;
  for (const it of items) total += 30 + it.name.length + it.comp.length + 46 + it.name.length;
  const out = new Uint8Array(total);
  let o = 0;
  const u16 = (v: number): void => { out[o++] = v & 255; out[o++] = (v >>> 8) & 255; };
  const u32 = (v: number): void => { out[o++] = v & 255; out[o++] = (v >>> 8) & 255; out[o++] = (v >>> 16) & 255; out[o++] = (v >>> 24) & 255; };
  const DOS_DATE = ((2026 - 1980) << 9) | (1 << 5) | 1;
  const offsets: number[] = [];
  for (const it of items) {
    offsets.push(o);
    u32(0x04034b50); u16(20); u16(0x0800); u16(it.method); u16(0); u16(DOS_DATE);
    u32(it.crc); u32(it.comp.length); u32(it.size); u16(it.name.length); u16(0);
    out.set(it.name, o); o += it.name.length;
    out.set(it.comp, o); o += it.comp.length;
  }
  const cd = o;
  items.forEach((it, i) => {
    u32(0x02014b50); u16(20); u16(20); u16(0x0800); u16(it.method); u16(0); u16(DOS_DATE);
    u32(it.crc); u32(it.comp.length); u32(it.size); u16(it.name.length); u16(0); u16(0); u16(0); u16(0); u32(0); u32(offsets[i]);
    out.set(it.name, o); o += it.name.length;
  });
  const cdSize = o - cd;
  u32(0x06054b50); u16(0); u16(0); u16(items.length); u16(items.length); u32(cdSize); u32(cd); u16(0);
  return out;
}

function unzip(buf: Uint8Array): { [name: string]: RawEntry } {
  const u16 = (i: number): number => buf[i] | (buf[i + 1] << 8);
  const u32 = (i: number): number => (buf[i] | (buf[i + 1] << 8) | (buf[i + 2] << 16) | (buf[i + 3] << 24)) >>> 0;
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (u32(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("テンプレートが ZIP (PPTX/POTX) 形式ではありません。Base64 が正しいか確認してください");
  const count = u16(eocd + 10);
  let p = u32(eocd + 16);
  const out: { [name: string]: RawEntry } = {};
  for (let n = 0; n < count; n++) {
    if (u32(p) !== 0x02014b50) throw new Error("テンプレートの ZIP 構造が壊れています");
    const method = u16(p + 10), crc = u32(p + 16), compSize = u32(p + 20), size = u32(p + 24);
    const nl = u16(p + 28), xl = u16(p + 30), cl = u16(p + 32), off = u32(p + 42);
    const name = utf8Decode(buf.subarray(p + 46, p + 46 + nl)).replace(/\\/g, "/").replace(/^\//, "");
    p += 46 + nl + xl + cl;
    if (name.charAt(name.length - 1) === "/") continue;
    const start = off + 30 + u16(off + 26) + u16(off + 28);
    out[name] = { method, crc, size, comp: buf.subarray(start, start + compSize) };
  }
  return out;
}

interface Huff { counts: Uint16Array; symbols: Uint16Array }

function buildHuff(lengths: Uint8Array, n: number): Huff {
  const counts = new Uint16Array(16);
  const symbols = new Uint16Array(n);
  for (let i = 0; i < n; i++) counts[lengths[i]]++;
  counts[0] = 0;
  const offs = new Uint16Array(16);
  for (let i = 1; i < 16; i++) offs[i] = offs[i - 1] + counts[i - 1];
  for (let i = 0; i < n; i++) if (lengths[i]) symbols[offs[lengths[i]]++] = i;
  return { counts, symbols };
}

const LBASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const LEXT = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
const DBASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
const DEXT = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];

function inflate(src: Uint8Array, size: number): Uint8Array {
  let out = new Uint8Array(Math.max(size, 1024));
  let op = 0, ip = 0, bitBuf = 0, bitCnt = 0;
  const need = (n: number): void => {
    if (op + n <= out.length) return;
    const bigger = new Uint8Array(Math.max(out.length * 2, op + n));
    bigger.set(out);
    out = bigger;
  };
  const bits = (n: number): number => {
    while (bitCnt < n) { bitBuf |= (ip < src.length ? src[ip] : 0) << bitCnt; ip++; bitCnt += 8; }
    const v = bitBuf & ((1 << n) - 1);
    bitBuf >>>= n;
    bitCnt -= n;
    return v;
  };
  const decode = (h: Huff): number => {
    let code = 0, first = 0, index = 0;
    for (let len = 1; len < 16; len++) {
      code |= bits(1);
      const count = h.counts[len];
      if (code - count < first) return h.symbols[index + (code - first)];
      index += count;
      first += count;
      first <<= 1;
      code <<= 1;
    }
    throw new Error("Deflate データが壊れています");
  };
  let fixedLit: Huff | null = null, fixedDist: Huff | null = null;
  let last = 0;
  while (!last) {
    last = bits(1);
    const type = bits(2);
    if (type === 0) {
      bitBuf = 0; bitCnt = 0;
      const len = src[ip] | (src[ip + 1] << 8);
      ip += 4;
      need(len);
      out.set(src.subarray(ip, ip + len), op);
      op += len; ip += len;
      continue;
    }
    let lit: Huff, dist: Huff;
    if (type === 1) {
      if (!fixedLit || !fixedDist) {
        const l = new Uint8Array(288);
        for (let i = 0; i < 288; i++) l[i] = i < 144 ? 8 : i < 256 ? 9 : i < 280 ? 7 : 8;
        fixedLit = buildHuff(l, 288);
        const d = new Uint8Array(30).fill(5);
        fixedDist = buildHuff(d, 30);
      }
      lit = fixedLit; dist = fixedDist;
    } else if (type === 2) {
      const hlit = bits(5) + 257, hdist = bits(5) + 1, hclen = bits(4) + 4;
      const order = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
      const cl = new Uint8Array(19);
      for (let i = 0; i < hclen; i++) cl[order[i]] = bits(3);
      const clh = buildHuff(cl, 19);
      const lens = new Uint8Array(hlit + hdist);
      for (let i = 0; i < hlit + hdist;) {
        const sym = decode(clh);
        if (sym < 16) lens[i++] = sym;
        else if (sym === 16) { const prev = lens[i - 1]; for (let r = 3 + bits(2); r > 0; r--) lens[i++] = prev; }
        else if (sym === 17) { for (let r = 3 + bits(3); r > 0; r--) lens[i++] = 0; }
        else { for (let r = 11 + bits(7); r > 0; r--) lens[i++] = 0; }
      }
      lit = buildHuff(lens.subarray(0, hlit), hlit);
      dist = buildHuff(lens.subarray(hlit), hdist);
    } else {
      throw new Error("Deflate ブロック形式が不正です");
    }
    for (;;) {
      const sym = decode(lit);
      if (sym < 256) { need(1); out[op++] = sym; continue; }
      if (sym === 256) break;
      const li = sym - 257;
      const len = LBASE[li] + bits(LEXT[li]);
      const di = decode(dist);
      const d = DBASE[di] + bits(DEXT[di]);
      need(len);
      for (let k = 0; k < len; k++, op++) out[op] = out[op - d];
    }
  }
  return op === out.length ? out : out.slice(0, op);
}
