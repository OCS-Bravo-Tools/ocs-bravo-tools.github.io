/**
 * Minimal .xlsx writer and reader, so no library is needed.
 * Writer: inline strings, numbers, a bold header row, yellow cells, column widths, in a
 * stored (uncompressed) zip; for the paper form also formulas, merged cells, dropdowns,
 * row heights and the original sheet's colors (STYLE); for the feedback, horizontal bar
 * charts (sheet.charts). Reader: every tab's cells as text, read in this browser.
 */
import { inflate, fromUtf8 } from './codec.js';

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

/** The paper form's cell styles (a cell { v, f, s }): the original sheet's olive bands and header, thin borders. */
export const STYLE = { band: 3, title: 4, head: 5, label: 6, ask: 7, val: 8, mid: 9, box: 10, note: 11, rater: 12, bar: 13 };

/** A stored (uncompressed) zip of { name, data } files. */
export function zip(files) {
  const enc = new TextEncoder();
  const chunks = [];
  const central = [];
  let offset = 0;
  files.forEach((f) => {
    const name = enc.encode(f.name);
    const data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
    const crc = crc32(data);
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint16(6, 0x0800, true);
    local.setUint16(8, 0, true);
    local.setUint16(10, 0, true);
    local.setUint16(12, 0x21, true);
    local.setUint32(14, crc, true);
    local.setUint32(18, data.length, true);
    local.setUint32(22, data.length, true);
    local.setUint16(26, name.length, true);
    local.setUint16(28, 0, true);
    chunks.push(new Uint8Array(local.buffer), name, data);
    const cen = new DataView(new ArrayBuffer(46));
    cen.setUint32(0, 0x02014b50, true);
    cen.setUint16(4, 20, true);
    cen.setUint16(6, 20, true);
    cen.setUint16(8, 0x0800, true);
    cen.setUint16(10, 0, true);
    cen.setUint16(12, 0, true);
    cen.setUint16(14, 0x21, true);
    cen.setUint32(16, crc, true);
    cen.setUint32(20, data.length, true);
    cen.setUint32(24, data.length, true);
    cen.setUint16(28, name.length, true);
    cen.setUint32(42, offset, true);
    central.push(new Uint8Array(cen.buffer), name);
    offset += 30 + name.length + data.length;
  });
  const cenSize = central.reduce((s, c) => s + c.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, files.length, true);
  end.setUint16(10, files.length, true);
  end.setUint32(12, cenSize, true);
  end.setUint32(16, offset, true);
  const all = chunks.concat(central, [new Uint8Array(end.buffer)]);
  const out = new Uint8Array(all.reduce((s, c) => s + c.length, 0));
  let p = 0;
  all.forEach((c) => {
    out.set(c, p);
    p += c.length;
  });
  return out;
}

const esc = (s) => String(s)
  .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;');

function colName(i) {
  let s = '';
  i += 1;
  while (i > 0) {
    const m = (i - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    i = Math.floor((i - 1) / 26);
  }
  return s;
}

const isCell = (x) => x !== null && typeof x === 'object'; // { v, f: a formula, s: a STYLE }

function sheetXml(sheet) {
  const cols = (sheet.widths || []).map((w, i) => '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>').join('');
  const ht = sheet.heights || {};
  const rows = sheet.rows.map((row, r) => {
    const cells = row.map((x, c) => {
      const ref = colName(c) + (r + 1);
      const o = isCell(x) ? x : { v: x };
      const v = o.v;
      const yellow = sheet.fill && sheet.fill.has(r + ',' + c);
      const s = o.s !== undefined ? o.s : yellow ? 2 : (r === 0 && sheet.header !== false) || (sheet.bold || []).indexOf(r) >= 0 ? 1 : 0;
      const style = s ? ' s="' + s + '"' : '';
      if (o.f && typeof v === 'string' && v !== '') return '<c r="' + ref + '"' + style + ' t="str"><f>' + esc(o.f) + '</f><v>' + esc(v) + '</v></c>'; // a text result, cached too
      if (o.f) return '<c r="' + ref + '"' + style + '><f>' + esc(o.f) + '</f>' + (typeof v === 'number' && isFinite(v) ? '<v>' + v + '</v>' : '') + '</c>'; // v: the result worked out here, for previews that don't calculate
      if (v === null || v === undefined || v === '') return yellow || o.s ? '<c r="' + ref + '"' + style + '/>' : '';
      if (typeof v === 'number' && isFinite(v)) return '<c r="' + ref + '"' + style + '><v>' + v + '</v></c>';
      return '<c r="' + ref + '"' + style + ' t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
    }).join('');
    return '<row r="' + (r + 1) + '"' + (ht[r] ? ' ht="' + ht[r] + '" customHeight="1"' : '') + '>' + cells + '</row>';
  }).join('');
  const merge = (sheet.merge || []).map((m) => '<mergeCell ref="' + m + '"/>').join('');
  const lists = (sheet.lists || []).filter(Boolean).map((l) => '<dataValidation type="list" allowBlank="1" showErrorMessage="1"' + (l.prompt ? ' showInputMessage="1" promptTitle="' + esc(l.promptTitle || '') + '" prompt="' + esc(l.prompt) + '"' : '') + ' sqref="' + l.ref + '"><formula1>' + esc(l.range || '"' + l.items.join(',') + '"') + '</formula1></dataValidation>').join('');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    (sheet.landscape ? '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' : '') +
    (sheet.rowHeight ? '<sheetFormatPr defaultRowHeight="' + sheet.rowHeight + '" customHeight="1"/>' : '') + // fixed rows as in the original (a preview blows up auto rows)
    (cols ? '<cols>' + cols + '</cols>' : '') +
    '<sheetData>' + rows + '</sheetData>' +
    (merge ? '<mergeCells count="' + sheet.merge.length + '">' + merge + '</mergeCells>' : '') +
    (sheet.dupes || []).map((ref, i) => { // a value twice in the range: red (dxf 0); Google Sheets drops duplicateValues, so the same as a formula too
      const [a, b] = ref.split(':');
      const abs = (c) => c.replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');
      return '<conditionalFormatting sqref="' + ref + '"><cfRule type="duplicateValues" dxfId="0" priority="' + (2 * i + 1) + '"/>' +
        '<cfRule type="expression" dxfId="0" priority="' + (2 * i + 2) + '"><formula>' + esc('AND(' + a + '<>"",COUNTIF(' + abs(a) + ':' + abs(b) + ',' + a + ')>1)') + '</formula></cfRule></conditionalFormatting>';
    }).join('') +
    (lists ? '<dataValidations count="' + sheet.lists.filter(Boolean).length + '">' + lists + '</dataValidations>' : '') +
    (sheet.landscape ? '<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>' : '') +
    (sheet.charts && sheet.charts.length ? '<drawing r:id="rId1"/>' : '') +
    '</worksheet>';
}

/* ---------- charts: { title, from: [col, row], to: [col, row], max, cats, catRef, series: [{ name, ref, values, color, labels }] } ---------- */
const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A horizontal bar chart: categories top to bottom, one bar per series, values 0 to max. */
function chartXml(ch) {
  const strCache = (vals) => '<c:strCache><c:ptCount val="' + vals.length + '"/>' + vals.map((v, i) => '<c:pt idx="' + i + '"><c:v>' + esc(v) + '</c:v></c:pt>').join('') + '</c:strCache>';
  const numCache = (vals) => '<c:numCache><c:formatCode>0.0</c:formatCode><c:ptCount val="' + vals.length + '"/>' + vals.map((v, i) => (typeof v === 'number' && isFinite(v) ? '<c:pt idx="' + i + '"><c:v>' + v + '</c:v></c:pt>' : '')).join('') + '</c:numCache>';
  const txt = (t, sz, b) => '<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="' + sz + '" b="' + (b ? 1 : 0) + '"/></a:pPr><a:r><a:rPr lang="en-US" sz="' + sz + '" b="' + (b ? 1 : 0) + '"/><a:t>' + esc(t) + '</a:t></a:r></a:p></c:rich></c:tx>';
  const ser = ch.series.map((s, i) => '<c:ser><c:idx val="' + i + '"/><c:order val="' + i + '"/>' +
    '<c:tx><c:v>' + esc(s.name) + '</c:v></c:tx>' +
    '<c:spPr><a:solidFill><a:srgbClr val="' + s.color + '"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>' +
    (s.labels ? '<c:dLbls><c:numFmt formatCode="0.0" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="1"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>' : '') +
    '<c:cat><c:strRef><c:f>' + esc(ch.catRef) + '</c:f>' + strCache(ch.cats) + '</c:strRef></c:cat>' +
    '<c:val><c:numRef><c:f>' + esc(s.ref) + '</c:f>' + numCache(s.values) + '</c:numRef></c:val></c:ser>').join('');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<c:chartSpace xmlns:c="' + NS_C + '" xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '"><c:roundedCorners val="0"/>' +
    '<c:chart><c:title>' + txt(ch.title, 1200, true) + '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>' +
    '<c:plotArea><c:layout/><c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:varyColors val="0"/>' + ser +
    '<c:gapWidth val="50"/><c:axId val="5001"/><c:axId val="5002"/></c:barChart>' +
    '<c:catAx><c:axId val="5001"/><c:scaling><c:orientation val="maxMin"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' +
    '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:crossAx val="5002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
    '<c:valAx><c:axId val="5002"/><c:scaling><c:orientation val="minMax"/><c:max val="' + ch.max + '"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="E4E5DC"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="0" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' +
    '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:crossAx val="5001"/><c:crosses val="max"/><c:crossBetween val="between"/><c:majorUnit val="1"/></c:valAx>' +
    '<c:spPr><a:noFill/></c:spPr></c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' +
    '<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="E4E5DC"/></a:solidFill></a:ln></c:spPr></c:chartSpace>';
}

/** The drawing that places charts on a sheet: anchors by [col, row] (0-based) from and to; chart k is rId(k+1) in the drawing's rels. */
function drawingXml(charts) {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="' + NS_A + '">' +
    charts.map((ch, k) => '<xdr:twoCellAnchor><xdr:from><xdr:col>' + ch.from[0] + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + ch.from[1] + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
      '<xdr:to><xdr:col>' + ch.to[0] + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + ch.to[1] + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
      '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="' + (k + 2) + '" name="Chart ' + (k + 1) + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
      '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="' + NS_C + '"><c:chart xmlns:c="' + NS_C + '" xmlns:r="' + NS_R + '" r:id="rId' + (k + 1) + '"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>').join('') +
    '</xdr:wsDr>';
}


/** One cell style: font, fill, border, alignment. */
const xf = (font, fill, h, v, wrap) => '<xf numFmtId="0" fontId="' + font + '" fillId="' + fill + '" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="' + h + '" vertical="' + v + '"' + (wrap ? ' wrapText="1"' : '') + '/></xf>';
const font = (b, sz, rgb) => '<font>' + (b ? '<b/>' : '') + '<sz val="' + sz + '"/>' + (rgb ? '<color rgb="' + rgb + '"/>' : '') + '<name val="Calibri"/></font>';
const solid = (rgb) => '<fill><patternFill patternType="solid"><fgColor rgb="' + rgb + '"/><bgColor indexed="64"/></patternFill></fill>';
const thin = ['left', 'right', 'top', 'bottom'].map((e) => '<' + e + ' style="thin"><color auto="1"/></' + e + '>').join('');

/** Excel sheet names: 31 characters, none of []:*?/\ , no ' at either end, and unique (ignoring case). */
function safeNames(names) {
  const used = {};
  return names.map((n) => {
    let s = String(n).replace(/[\[\]:*?/\\]/g, ' ').slice(0, 31).replace(/^[\s']+|[\s']+$/g, '') || 'Sheet';
    let k = s;
    let i = 2;
    while (used[k.toLowerCase()]) k = s.slice(0, 28) + ' ' + i++;
    used[k.toLowerCase()] = true;
    return k;
  });
}

/**
 * @param {{name: string, rows: any[][], widths?: number[], header?: boolean, bold?: number[], fill?: Set<string>, rowHeight?: number, heights?: object, merge?: string[], lists?: object[], landscape?: boolean}[]} sheets
 *   header: false leaves row 1 plain; bold: more rows (0-based) in bold; fill: "row,col" cells in yellow;
 *   a cell may be { v, f, s } (value, formula, STYLE); heights: row (0-based) -> points, rowHeight the rest; merge: "D1:J1" ranges;
 *   lists: dropdowns { ref: "E7 G7", items: ["Yes", "No"] } or { ref, range: "'Tab'!$B$4:$B$9" }; landscape: print on one page wide
 * @returns {Uint8Array} the .xlsx file
 */
export function buildXlsx(sheets) {
  const names = safeNames(sheets.map((s) => s.name));
  const files = [
    {
      name: '[Content_Types].xml',
      data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        sheets.map((s, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') +
        '</Types>'
    },
    {
      name: '_rels/.rels',
      data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>'
    },
    {
      name: 'xl/workbook.xml',
      data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
        names.map((n, i) => '<sheet name="' + esc(n) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('') +
        '</sheets>' + (sheets.some((s) => s.rows.some((row) => row.some((x) => isCell(x) && x.f))) ? '<calcPr fullCalcOnLoad="1"/>' : '') + '</workbook>'
    },
    {
      name: 'xl/_rels/workbook.xml.rels',
      data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        sheets.map((s, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') +
        '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>'
    },
    {
      name: 'xl/styles.xml',
      data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts count="7">' + font(0, 11) + font(1, 11) + font(1, 12, 'FF3B4426') + font(1, 16, 'FF3B4426') + font(1, 11, 'FFFFFFFF') + font(0, 9) + font(0, 11, 'FF5B7A4B') + '</fonts>' +
        '<fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' + solid('FFFFF2B3') + solid('FFC2D69B') + solid('FF3B4426') + solid('FFFFC000') + '</fills>' +
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border>' + thin + '<diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="14"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>' +
        xf(2, 3, 'center', 'center') + xf(3, 3, 'center', 'center') + xf(4, 4, 'center', 'center', 1) + xf(1, 0, 'right', 'center', 1) + xf(1, 0, 'center', 'center', 1) +
        xf(0, 0, 'left', 'center', 1) + xf(0, 0, 'center', 'center', 1) + xf(0, 0, 'left', 'top', 1) + xf(5, 0, 'center', 'center', 1) + xf(0, 5, 'right', 'center') + xf(6, 0, 'left', 'center') + '</cellXfs>' +
        '<dxfs count="1"><dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>' + // conditional formats: 0 = a repeat, red
        '</styleSheet>'
    }
  ];
  sheets.forEach((s, i) => files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: sheetXml(s) }));
  let nChart = 0;
  const types = [];
  sheets.forEach((s, i) => { // a sheet's charts: its drawing, the drawing's charts, and their content types
    if (!s.charts || !s.charts.length) return;
    const d = i + 1;
    files.push({ name: 'xl/worksheets/_rels/sheet' + d + '.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' + d + '.xml"/></Relationships>' });
    files.push({ name: 'xl/drawings/drawing' + d + '.xml', data: drawingXml(s.charts) });
    types.push('<Override PartName="/xl/drawings/drawing' + d + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>');
    const rels = s.charts.map((ch, k) => {
      nChart += 1;
      files.push({ name: 'xl/charts/chart' + nChart + '.xml', data: chartXml(ch) });
      types.push('<Override PartName="/xl/charts/chart' + nChart + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>');
      return '<Relationship Id="rId' + (k + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart' + nChart + '.xml"/>';
    });
    files.push({ name: 'xl/drawings/_rels/drawing' + d + '.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels.join('') + '</Relationships>' });
  });
  if (types.length) files[0].data = files[0].data.replace('</Types>', types.join('') + '</Types>');
  return zip(files);
}

/* ---------- reading ---------- */

const unescape = (s) => String(s)
  .replace(/&#x([0-9a-f]+);/gi, (m, h) => String.fromCodePoint(parseInt(h, 16)))
  .replace(/&#(\d+);/g, (m, d) => String.fromCodePoint(Number(d)))
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** The text of every <t> in a piece of XML (rich text runs joined; phonetic runs left out). */
const texts = (xml) => unescape((String(xml).replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').match(/<t\b[^>]*?\/>|<t\b[^>]*>[\s\S]*?<\/t>/g) || []).map((t) => (/\/>$/.test(t) && !/<\/t>$/.test(t) ? '' : t.replace(/^<t\b[^>]*>|<\/t>$/g, ''))).join(''));

/** The files inside a zip, by name; deflated ones are opened on reading. */
function unzip(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) {
      end = i;
      break;
    }
  }
  if (end < 0) throw new Error('notzip');
  const files = {};
  let p = dv.getUint32(end + 16, true);
  for (let n = dv.getUint16(end + 10, true); n > 0 && dv.getUint32(p, true) === 0x02014b50; n--) {
    const method = dv.getUint16(p + 10, true);
    const size = dv.getUint32(p + 20, true);
    const nameLen = dv.getUint16(p + 28, true);
    const at = dv.getUint32(p + 42, true);
    const name = fromUtf8(bytes.subarray(p + 46, p + 46 + nameLen));
    const start = at + 30 + dv.getUint16(at + 26, true) + dv.getUint16(at + 28, true);
    const tm = dv.getUint16(p + 12, true);
    const dt = dv.getUint16(p + 14, true);
    const time = dt ? new Date(1980 + (dt >> 9), ((dt >> 5) & 15) - 1, dt & 31, tm >> 11, (tm >> 5) & 63, (tm & 31) * 2).getTime() : 0; // the entry's date (17m: the newest form wins)
    files[name] = { method, data: bytes.subarray(start, start + size), time };
    p += 46 + nameLen + dv.getUint16(p + 30, true) + dv.getUint16(p + 32, true);
  }
  return files;
}

/** The .xlsx / .xlsm files inside a zip (a Classroom or Drive download, folders and all): [{ name, data }]. */
export async function xlsxInZip(bytes) {
  const files = unzip(bytes);
  const out = [];
  for (const [name, f] of Object.entries(files)) {
    if (!/\.xls[xm]$/i.test(name) || /(^|\/)(__MACOSX\/|\._)/.test(name)) continue;
    out.push({ name: name.split('/').pop(), data: f.method === 8 ? await inflate(f.data) : f.data.slice(), time: f.time || 0 });
  }
  return out;
}

/**
 * Every tab of an .xlsx, in workbook order, as rows of text cells (what a copy from Excel
 * gives: the value of a formula, a merged cell's text in its first cell only). Read here,
 * never sent anywhere.
 * @param {Uint8Array} bytes
 * @returns {Promise<{name: string, cells: string[][]}[]>}
 * @throws {Error} 'locked' for an old .xls or a password-protected file, 'notxlsx' for anything else
 */
export async function readXlsx(bytes) {
  if (bytes.length > 8 && bytes[0] === 0xD0 && bytes[1] === 0xCF && bytes[2] === 0x11 && bytes[3] === 0xE0) throw new Error('locked');
  let files;
  try {
    files = unzip(bytes);
  } catch (err) {
    throw new Error('notxlsx');
  }
  const read = async (name) => {
    const f = files[name.replace(/^\//, '')];
    if (!f) return '';
    return fromUtf8(f.method === 8 ? await inflate(f.data) : f.data);
  };
  const book = await read('xl/workbook.xml');
  if (!book) throw new Error('notxlsx');
  const rels = {};
  ((await read('xl/_rels/workbook.xml.rels')).match(/<Relationship\b[^>]*>/g) || []).forEach((r) => {
    const id = (/\bId="([^"]*)"/.exec(r) || [])[1];
    const target = (/\bTarget="([^"]*)"/.exec(r) || [])[1] || '';
    if (id) rels[id] = target.charAt(0) === '/' ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
  });
  const shared = ((await read('xl/sharedStrings.xml')).match(/<si\b[\s\S]*?<\/si>/g) || []).map(texts);
  const out = [];
  for (const tag of book.match(/<sheet\b[^>]*>/g) || []) {
    const name = unescape((/\bname="([^"]*)"/.exec(tag) || [])[1] || 'Sheet');
    const rid = (/\s[A-Za-z0-9]+:id="([^"]*)"/.exec(tag) || [])[1];
    const xml = rels[rid] ? await read(rels[rid]) : '';
    const cells = [];
    let r = -1;
    for (const row of xml.match(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g) || []) {
      const rn = (/^<row\b[^>]*?\br="(\d+)"/.exec(row) || [])[1];
      r = rn ? Number(rn) - 1 : r + 1;
      const line = (cells[r] = cells[r] || []);
      let col = -1;
      for (const m of row.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const ref = (/\br="([A-Z]+)\d+"/.exec(m[1]) || [])[1];
        col = ref ? ref.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 : col + 1;
        const type = (/\bt="([^"]*)"/.exec(m[1]) || [])[1] || '';
        const inner = m[2] || '';
        const v = (/<v\b[^>]*>([\s\S]*?)<\/v>/.exec(inner) || [])[1];
        let text = '';
        if (type === 's') text = shared[Number(v)] || '';
        else if (type === 'inlineStr') text = texts(inner);
        else if (type === 'b') text = v === '1' ? 'TRUE' : v === '0' ? 'FALSE' : '';
        else text = v === undefined ? '' : unescape(v);
        line[col] = text;
      }
    }
    for (let i = 0; i < cells.length; i++) cells[i] = Array.from(cells[i] || [], (x) => (x === undefined ? '' : x));
    out.push({ name, cells });
  }
  return out;
}
