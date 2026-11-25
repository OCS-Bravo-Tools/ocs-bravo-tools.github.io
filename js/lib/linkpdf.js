/**
 * The link list: one PDF with every OC's personal link, for the admin to upload once
 * (Teams, a shared drive, the group chat). Each OC finds their name and taps it.
 * Plain PDF 1.4 (Helvetica, one link annotation per row, nothing compressed), so any
 * phone or computer opens it and the links work. No DOM.
 */
const PAGE_W = 612;
const PAGE_H = 792;
const LEFT = 54;
const RIGHT = 558;
const TOP = 738;
const BOTTOM = 72;
const ROW = 24;

/** Text the standard fonts can print (WinAnsi / Latin-1); anything else becomes '?'. */
const latin1 = (s) => String(s).replace(/[^\x20-\x7E\xA0-\xFF]/g, '?');
const str = (s) => '(' + latin1(s).replace(/[\\()]/g, (c) => '\\' + c) + ')';

/**
 * @param {{title: string, notes: string[], groups: {label: string, rows: {name: string, link: string}[]}[]}} doc
 *   notes: lines printed under the title, each short enough for one line
 * @returns {Uint8Array} the PDF file
 */
export function buildLinkPdf(doc) {
  const pages = [];
  let page = null;
  let y = 0;
  const newPage = () => {
    page = { ops: [], links: [] };
    pages.push(page);
    y = TOP;
  };
  const text = (x, at, size, bold, words, rgb) => {
    page.ops.push((rgb || '0.09 0.1 0.11') + ' rg BT /' + (bold ? 'F2' : 'F1') + ' ' + size + ' Tf ' + x + ' ' + at + ' Td ' + str(words) + ' Tj ET');
  };
  newPage();
  text(LEFT, y, 18, true, doc.title);
  y -= 24;
  doc.notes.forEach((line) => {
    text(LEFT, y, 10.5, false, line, '0.29 0.31 0.34');
    y -= 15;
  });
  y -= 12;
  doc.groups.forEach((g) => {
    if (y - 22 - ROW < BOTTOM) newPage(); // a squad's heading stays with its first name
    text(LEFT, y, 13, true, g.label);
    y -= 22;
    g.rows.forEach((r) => {
      if (y < BOTTOM) {
        newPage();
        text(LEFT, y, 13, true, g.label + ' (continued)');
        y -= 22;
      }
      text(LEFT, y, 12, true, r.name);
      text(330, y, 12, false, 'Open my evaluation', '0.15 0.29 0.53');
      page.ops.push('0.81 0.8 0.75 RG 0.5 w ' + LEFT + ' ' + (y - 8) + ' m ' + RIGHT + ' ' + (y - 8) + ' l S');
      page.links.push({ rect: [LEFT, y - 8, RIGHT, y + 16], uri: r.link });
      y -= ROW;
    });
    y -= 12;
  });
  pages.forEach((p, i) => {
    page = p;
    text(LEFT, 40, 9, false, 'Page ' + (i + 1) + ' of ' + pages.length + '  ·  ' + doc.title, '0.29 0.31 0.34');
  });

  // Objects: 1 catalog, 2 page tree, 3-4 fonts, 5 info, then per page: page, contents, links.
  const objs = [];
  const add = (body) => objs.push(body) + 0;
  add('<< /Type /Catalog /Pages 2 0 R >>');
  add(null); // the page tree, filled in once the page numbers are known
  add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>');
  add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>');
  add('<< /Title ' + str(doc.title) + ' /Producer (Peer Eval) >>');
  const kids = [];
  pages.forEach((p) => {
    const content = p.ops.join('\n');
    const pageNo = objs.length + 1;
    const contentNo = pageNo + 1;
    const annots = p.links.map((l, i) => contentNo + 1 + i);
    kids.push(pageNo + ' 0 R');
    add('<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ' + PAGE_W + ' ' + PAGE_H + '] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents ' + contentNo + ' 0 R' +
      (annots.length ? ' /Annots [' + annots.map((n) => n + ' 0 R').join(' ') + ']' : '') + ' >>');
    add('<< /Length ' + content.length + ' >>\nstream\n' + content + '\nendstream');
    p.links.forEach((l) => add('<< /Type /Annot /Subtype /Link /Rect [' + l.rect.join(' ') + '] /Border [0 0 0] /A << /Type /Action /S /URI /URI ' + str(l.uri) + ' >> >>'));
  });
  objs[1] = '<< /Type /Pages /Kids [' + kids.join(' ') + '] /Count ' + pages.length + ' >>';

  let out = '%PDF-1.4\n%\xE2\xE3\xCF\xD3\n';
  const offsets = [];
  objs.forEach((body, i) => {
    offsets.push(out.length);
    out += (i + 1) + ' 0 obj\n' + body + '\nendobj\n';
  });
  const xref = out.length;
  out += 'xref\n0 ' + (objs.length + 1) + '\n0000000000 65535 f \n' + offsets.map((o) => String(o).padStart(10, '0') + ' 00000 n \n').join('');
  out += 'trailer\n<< /Size ' + (objs.length + 1) + ' /Root 1 0 R /Info 5 0 R >>\nstartxref\n' + xref + '\n%%EOF\n';
  return Uint8Array.from(out, (c) => c.charCodeAt(0));
}
