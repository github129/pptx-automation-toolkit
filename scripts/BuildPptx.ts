/**
 * Office Script: BuildPptx
 * Power Automate から受け取ったタイトルと表データ(JSON)で 1 枚スライドの PPTX を生成し、
 * Base64 文字列として返す。Premium コネクタ不要 (Excel Online (Business) の「スクリプトの実行」で呼ぶ)。
 *
 * title     : スライドタイトル
 * tableJson : オブジェクト配列の JSON 文字列。例 [{"商品名":"A","売上":100}, ...]
 *             1 行目のキーの順番が列見出しになる (Select アクションの出力をそのまま渡せる)
 */
function main(workbook: ExcelScript.Workbook, title: string, tableJson: string): string {
  const rows = JSON.parse(tableJson) as CellRecord[];
  if (!Array.isArray(rows) || rows.length === 0) {
    throw new Error("tableJson は 1 件以上のオブジェクト配列である必要があります");
  }
  const headers = Object.keys(rows[0]);
  const cells: string[][] = rows.map(r => headers.map(h => toText(r[h])));
  const numeric: boolean[] = headers.map(h => rows.every(r => typeof r[h] === "number"));

  const files: ZipEntry[] = [
    { name: "[Content_Types].xml", data: utf8(contentTypesXml()) },
    { name: "_rels/.rels", data: utf8(rootRelsXml()) },
    { name: "ppt/presentation.xml", data: utf8(presentationXml()) },
    { name: "ppt/_rels/presentation.xml.rels", data: utf8(presentationRelsXml()) },
    { name: "ppt/presProps.xml", data: utf8(presPropsXml()) },
    { name: "ppt/viewProps.xml", data: utf8(viewPropsXml()) },
    { name: "ppt/tableStyles.xml", data: utf8(tableStylesXml()) },
    { name: "ppt/theme/theme1.xml", data: utf8(themeXml()) },
    { name: "ppt/slideMasters/slideMaster1.xml", data: utf8(slideMasterXml()) },
    { name: "ppt/slideMasters/_rels/slideMaster1.xml.rels", data: utf8(slideMasterRelsXml()) },
    { name: "ppt/slideLayouts/slideLayout1.xml", data: utf8(slideLayoutXml()) },
    { name: "ppt/slideLayouts/_rels/slideLayout1.xml.rels", data: utf8(slideLayoutRelsXml()) },
    { name: "ppt/slides/slide1.xml", data: utf8(slideXml(title, headers, cells, numeric)) },
    { name: "ppt/slides/_rels/slide1.xml.rels", data: utf8(slideRelsXml()) },
  ];
  return base64(zip(files));
}

type CellValue = string | number | boolean | null;
interface CellRecord { [key: string]: CellValue }
interface ZipEntry { name: string; data: number[] }

function toText(v: CellValue): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return v.toLocaleString("ja-JP");
  return String(v);
}

function textWidth(s: string): number {
  let w = 0;
  for (let i = 0; i < s.length; i++) w += s.charCodeAt(i) < 0x80 ? 1 : 2;
  return w;
}

function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// ---------- スライド ----------
const SLIDE_W = 12192000; // 16:9
const SLIDE_H = 6858000;
const MARGIN = 457200;

function slideXml(title: string, headers: string[], cells: string[][], numeric: boolean[]): string {
  const tableTop = 1447800;
  const tableW = SLIDE_W - MARGIN * 2;
  const rowCount = cells.length + 1;
  const rowH = Math.min(457200, Math.floor((SLIDE_H - tableTop - MARGIN) / rowCount));
  const fontSz = rowCount > 12 ? 1000 : rowCount > 8 ? 1200 : 1400;
  // 列幅は見出し・セルの最大表示幅 (全角=2, 半角=1) に比例させる (4〜40 で丸める)
  const weights = headers.map((h, i) =>
    Math.min(40, Math.max(4, textWidth(h), ...cells.map(r => textWidth(r[i])))));
  const totalWeight = weights.reduce((a, b) => a + b, 0);
  const colWs = weights.map(w => Math.floor(tableW * w / totalWeight));

  const cell = (text: string, alignRight: boolean): string =>
    `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>${alignRight ? '<a:pPr algn="r"/>' : ""}` +
    (text === ""
      ? `<a:endParaRPr lang="ja-JP" sz="${fontSz}"/>`
      : `<a:r><a:rPr lang="ja-JP" sz="${fontSz}"/><a:t>${esc(text)}</a:t></a:r>`) +
    `</a:p></a:txBody><a:tcPr anchor="ctr"/></a:tc>`;

  const headerRow = `<a:tr h="${rowH}">${headers.map((h, i) => cell(h, numeric[i])).join("")}</a:tr>`;
  const bodyRows = cells.map(r => `<a:tr h="${rowH}">${r.map((c, i) => cell(c, numeric[i])).join("")}</a:tr>`).join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${MARGIN}" y="381000"/><a:ext cx="${tableW}" cy="838200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" anchor="b"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP" sz="3200" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/></a:rPr><a:t>${esc(title)}</a:t></a:r></a:p></p:txBody></p:sp>` +
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${MARGIN}" y="${tableTop}"/><a:ext cx="${colWs.reduce((a, b) => a + b, 0)}" cy="${rowH * rowCount}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid>${colWs.map(w => `<a:gridCol w="${w}"/>`).join("")}</a:tblGrid>${headerRow}${bodyRows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>` +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

// ---------- 固定パーツ ----------
const XML_HEAD = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
const NS_A = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;
const NS_R = `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
const NS_P = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT = "application/vnd.openxmlformats-officedocument.presentationml";
const EMPTY_TREE = `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree>`;

function rels(items: string[][]): string {
  return XML_HEAD + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    items.map(i => `<Relationship Id="${i[0]}" Type="${i[1]}" Target="${i[2]}"/>`).join("") + `</Relationships>`;
}

function contentTypesXml(): string {
  return XML_HEAD + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/ppt/presentation.xml" ContentType="${CT}.presentation.main+xml"/>` +
    `<Override PartName="/ppt/presProps.xml" ContentType="${CT}.presProps+xml"/>` +
    `<Override PartName="/ppt/viewProps.xml" ContentType="${CT}.viewProps+xml"/>` +
    `<Override PartName="/ppt/tableStyles.xml" ContentType="${CT}.tableStyles+xml"/>` +
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${CT}.slideMaster+xml"/>` +
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${CT}.slideLayout+xml"/>` +
    `<Override PartName="/ppt/slides/slide1.xml" ContentType="${CT}.slide+xml"/>` +
    `</Types>`;
}

function rootRelsXml(): string {
  return rels([["rId1", `${REL}/officeDocument`, "ppt/presentation.xml"]]);
}

function presentationXml(): string {
  return XML_HEAD + `<p:presentation ${NS_A} ${NS_R} ${NS_P} saveSubsetFonts="1">` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>` +
    `<p:sldSz cx="${SLIDE_W}" cy="${SLIDE_H}"/><p:notesSz cx="6858000" cy="9144000"/>` +
    `</p:presentation>`;
}

function presentationRelsXml(): string {
  return rels([
    ["rId1", `${REL}/slideMaster`, "slideMasters/slideMaster1.xml"],
    ["rId2", `${REL}/slide`, "slides/slide1.xml"],
    ["rId3", `${REL}/presProps`, "presProps.xml"],
    ["rId4", `${REL}/viewProps`, "viewProps.xml"],
    ["rId5", `${REL}/theme`, "theme/theme1.xml"],
    ["rId6", `${REL}/tableStyles`, "tableStyles.xml"],
  ]);
}

function presPropsXml(): string {
  return XML_HEAD + `<p:presentationPr ${NS_A} ${NS_R} ${NS_P}/>`;
}

function viewPropsXml(): string {
  return XML_HEAD + `<p:viewPr ${NS_A} ${NS_R} ${NS_P}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>`;
}

function tableStylesXml(): string {
  return XML_HEAD + `<a:tblStyleLst ${NS_A} def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;
}

function slideMasterXml(): string {
  return XML_HEAD + `<p:sldMaster ${NS_A} ${NS_R} ${NS_P}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>${EMPTY_TREE}</p:cSld>` +
    `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
    `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>` +
    `<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/></a:defRPr></a:lvl1pPr></p:titleStyle>` +
    `<p:bodyStyle><a:lvl1pPr><a:defRPr sz="2800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></p:bodyStyle>` +
    `<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles>` +
    `</p:sldMaster>`;
}

function slideMasterRelsXml(): string {
  return rels([
    ["rId1", `${REL}/slideLayout`, "../slideLayouts/slideLayout1.xml"],
    ["rId2", `${REL}/theme`, "../theme/theme1.xml"],
  ]);
}

function slideLayoutXml(): string {
  return XML_HEAD + `<p:sldLayout ${NS_A} ${NS_R} ${NS_P} type="blank" preserve="1"><p:cSld name="Blank">${EMPTY_TREE}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

function slideLayoutRelsXml(): string {
  return rels([["rId1", `${REL}/slideMaster`, "../slideMasters/slideMaster1.xml"]]);
}

function slideRelsXml(): string {
  return rels([["rId1", `${REL}/slideLayout`, "../slideLayouts/slideLayout1.xml"]]);
}

function themeXml(): string {
  const clr = (name: string, hex: string): string => `<a:${name}><a:srgbClr val="${hex}"/></a:${name}>`;
  const fill = `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`;
  const ln = (w: number): string => `<a:ln w="${w}"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>`;
  const font = (latin: string): string => `<a:latin typeface="${latin}"/><a:ea typeface="Yu Gothic"/><a:cs typeface=""/>`;
  return XML_HEAD + `<a:theme ${NS_A} name="Office Theme"><a:themeElements>` +
    `<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>` +
    clr("dk2", "0E2841") + clr("lt2", "E8E8E8") + clr("accent1", "156082") + clr("accent2", "E97132") + clr("accent3", "196B24") +
    clr("accent4", "0F9ED5") + clr("accent5", "A02B93") + clr("accent6", "4EA72E") + clr("hlink", "467886") + clr("folHlink", "96607D") +
    `</a:clrScheme>` +
    `<a:fontScheme name="Office"><a:majorFont>${font("Aptos Display")}</a:majorFont><a:minorFont>${font("Aptos")}</a:minorFont></a:fontScheme>` +
    `<a:fmtScheme name="Office"><a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst>` +
    `<a:lnStyleLst>${ln(6350)}${ln(12700)}${ln(19050)}</a:lnStyleLst>` +
    `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
    `<a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst></a:fmtScheme>` +
    `</a:themeElements></a:theme>`;
}

// ---------- UTF-8 / CRC32 / ZIP (無圧縮) / Base64 ----------
function utf8(s: string): number[] {
  const out: number[] = [];
  for (let i = 0; i < s.length; i++) {
    let c = s.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff && i + 1 < s.length) {
      c = 0x10000 + ((c - 0xd800) << 10) + (s.charCodeAt(++i) - 0xdc00);
    }
    if (c < 0x80) out.push(c);
    else if (c < 0x800) out.push(0xc0 | (c >> 6), 0x80 | (c & 63));
    else if (c < 0x10000) out.push(0xe0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
    else out.push(0xf0 | (c >> 18), 0x80 | ((c >> 12) & 63), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
  }
  return out;
}

function crc32(data: number[]): number {
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i++) {
    crc ^= data[i];
    for (let k = 0; k < 8; k++) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function zip(entries: ZipEntry[]): number[] {
  const out: number[] = [];
  const central: number[] = [];
  const u16 = (a: number[], v: number): void => { a.push(v & 0xff, (v >>> 8) & 0xff); };
  const u32 = (a: number[], v: number): void => { a.push(v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff); };
  const DOS_TIME = 0;
  const DOS_DATE = (2026 - 1980) << 9 | 1 << 5 | 1;
  for (const e of entries) {
    const name = utf8(e.name);
    const crc = crc32(e.data);
    const offset = out.length;
    u32(out, 0x04034b50); u16(out, 20); u16(out, 0x0800); u16(out, 0); u16(out, DOS_TIME); u16(out, DOS_DATE);
    u32(out, crc); u32(out, e.data.length); u32(out, e.data.length); u16(out, name.length); u16(out, 0);
    for (const b of name) out.push(b);
    for (const b of e.data) out.push(b);
    u32(central, 0x02014b50); u16(central, 20); u16(central, 20); u16(central, 0x0800); u16(central, 0); u16(central, DOS_TIME); u16(central, DOS_DATE);
    u32(central, crc); u32(central, e.data.length); u32(central, e.data.length); u16(central, name.length);
    u16(central, 0); u16(central, 0); u16(central, 0); u16(central, 0); u32(central, 0); u32(central, offset);
    for (const b of name) central.push(b);
  }
  const cdOffset = out.length;
  for (const b of central) out.push(b);
  u32(out, 0x06054b50); u16(out, 0); u16(out, 0); u16(out, entries.length); u16(out, entries.length);
  u32(out, central.length); u32(out, cdOffset); u16(out, 0);
  return out;
}

function base64(bytes: number[]): string {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  let s = "";
  for (let i = 0; i < bytes.length; i += 3) {
    const n = (bytes[i] << 16) | ((bytes[i + 1] ?? 0) << 8) | (bytes[i + 2] ?? 0);
    s += chars[(n >> 18) & 63] + chars[(n >> 12) & 63] +
      (i + 1 < bytes.length ? chars[(n >> 6) & 63] : "=") +
      (i + 2 < bytes.length ? chars[n & 63] : "=");
  }
  return s;
}
