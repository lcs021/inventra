/* Controle de Avaria — gerador da planilha Excel (.xlsx) corporativa.
   Uso: AvariaPlanilha.gerar(linhas, {usuario, cd, prazo}) -> Blob
   linhas: [{ref, rel:'AAAA-MM-DD', origem, sku, produto, ent:'AAAA-MM-DD'|'', por}] */
(function (G) {
  'use strict';
  // ---------------------------------------------------------------- utilidades
  const xe = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const colL = n => { let s = ''; n++; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const colN = s => { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  const parse = a => { const m = /^([A-Z]+)(\d+)$/.exec(a); return [colN(m[1]), +m[2] - 1]; };
  const A = (c, r) => colL(c) + (r + 1);
  const serial = s => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(s || '')); return m ? (Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 864e5 : null; };
  const isoLocal = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  // funções do Excel 365 precisam do prefixo de compatibilidade no arquivo
  const DYN = /\b(FILTER|SORTBY|SORT|UNIQUE|SEQUENCE|XLOOKUP|LET)\(/;
  const px = f => String(f).replace(/\bXLOOKUP\(/g, '_xlfn.XLOOKUP(').replace(/\bSORTBY\(/g, '_xlfn.SORTBY(').replace(/\bSORT\(/g, '_xlfn._xlws.SORT(').replace(/\bFILTER\(/g, '_xlfn._xlws.FILTER(').replace(/\bUNIQUE\(/g, '_xlfn.UNIQUE(').replace(/\bSEQUENCE\(/g, '_xlfn.SEQUENCE(').replace(/\bLET\(/g, '_xlfn.LET(').replace(/\bMAXIFS\(/g, '_xlfn.MAXIFS(').replace(/\bMINIFS\(/g, '_xlfn.MINIFS(');
  const MES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];

  // ---------------------------------------------------------------- zip (sem compressão)
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  function zip(files) {
    const enc = new TextEncoder(), parts = [], cen = []; let off = 0;
    const u16 = v => [v & 255, (v >>> 8) & 255], u32 = v => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255];
    for (const [name, txt] of files) {
      const nb = enc.encode(name), data = enc.encode(txt), c = crc32(data);
      const lh = new Uint8Array([...u32(0x04034b50), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(c), ...u32(data.length), ...u32(data.length), ...u16(nb.length), ...u16(0)]);
      parts.push(lh, nb, data);
      cen.push(new Uint8Array([...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(c), ...u32(data.length), ...u32(data.length), ...u16(nb.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(off)]), nb);
      off += lh.length + nb.length + data.length;
    }
    const cs = cen.reduce((a, x) => a + x.length, 0);
    const end = new Uint8Array([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length), ...u32(cs), ...u32(off), ...u16(0)]);
    return new Blob([...parts, ...cen, end], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  // ---------------------------------------------------------------- paleta e estilos
  const C = { azul: '1E5799', azulEsc: '173F6E', azulCl: 'DCE8F7', azulBg: 'F2F6FB', tinta: '1F2937', tinta2: '374151', mudo: '6B7280', fraco: '9CA3AF', linha: 'E2E8F0', linha2: 'CBD5E1', auto: 'F1F5F9', branco: 'FFFFFF', ambar: '92400E', ambarBg: 'FEF3C7', verm: 'B91C1C', vermBg: 'FEE2E2', input: 'FFFDF5' };
  const FONTE = 'Calibri';
  function Estilos() {
    const fonts = [], fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'], fmts = [], xfs = [], dxfs = [];
    const idx = (arr, xml) => { let i = arr.indexOf(xml); if (i < 0) { arr.push(xml); i = arr.length - 1; } return i; };
    const font = ({ sz = 10, b, i, cor = C.tinta } = {}) => idx(fonts, `<font>${b ? '<b/>' : ''}${i ? '<i/>' : ''}<sz val="${sz}"/><color rgb="FF${cor}"/><name val="${FONTE}"/><family val="2"/></font>`);
    font({});
    const fill = cor => cor ? idx(fills, `<fill><patternFill patternType="solid"><fgColor rgb="FF${cor}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
    const lado = (n, v) => v ? `<${n} style="${v[0]}"><color rgb="FF${v[1]}"/></${n}>` : `<${n}/>`;
    const border = b => b ? idx(borders, `<border>${lado('left', b.l)}${lado('right', b.r)}${lado('top', b.t)}${lado('bottom', b.b)}<diagonal/></border>`) : 0;
    const fmt = code => { if (!code || code === 'General') return 0; const fixo = { '0': 1, '0.00': 2, '#,##0': 3 }; if (fixo[code]) return fixo[code]; let i = fmts.indexOf(code); if (i < 0) { fmts.push(code); i = fmts.length - 1; } return 164 + i; };
    const cache = new Map();
    const xf = (o = {}) => {
      const k = JSON.stringify(o); if (cache.has(k)) return cache.get(k);
      const f = font(o.font || {}), fi = fill(o.fill), bo = border(o.border), nf = fmt(o.fmt);
      const al = o.al || {}, alXml = Object.keys(al).length ? `<alignment${al.h ? ` horizontal="${al.h}"` : ''}${al.v !== false ? ` vertical="${al.v || 'center'}"` : ''}${al.wrap ? ' wrapText="1"' : ''}${al.ind ? ` indent="${al.ind}"` : ''}/>` : '<alignment vertical="center"/>';
      const pr = o.livre ? '<protection locked="0"/>' : '';
      xfs.push(`<xf numFmtId="${nf}" fontId="${f}" fillId="${fi}" borderId="${bo}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"${pr ? ' applyProtection="1"' : ''}>${alXml}${pr}</xf>`);
      cache.set(k, xfs.length - 1); return xfs.length - 1;
    };
    xf({});
    const dxf = ({ cor, bg, b, borda } = {}) => { dxfs.push(`<dxf>${cor || b ? `<font>${b ? '<b/>' : ''}${cor ? `<color rgb="FF${cor}"/>` : ''}</font>` : ''}${bg ? `<fill><patternFill patternType="solid"><bgColor rgb="FF${bg}"/></patternFill></fill>` : ''}${borda ? `<border><bottom style="thin"><color rgb="FF${borda}"/></bottom><horizontal style="thin"><color rgb="FF${borda}"/></horizontal></border>` : ''}</dxf>`); return dxfs.length - 1; };
    const xml = tabela => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${fmts.length ? `<numFmts count="${fmts.length}">${fmts.map((c, i) => `<numFmt numFmtId="${164 + i}" formatCode="${xe(c)}"/>`).join('')}</numFmts>` : ''}<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs><tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"><tableStyle name="AvariaCorporativo" pivot="0" count="3"><tableStyleElement type="wholeTable" dxfId="${tabela.inteira}"/><tableStyleElement type="headerRow" dxfId="${tabela.cab}"/><tableStyleElement type="firstRowStripe" dxfId="${tabela.faixa}"/></tableStyle></tableStyles></styleSheet>`;
    return { xf, dxf, xml };
  }

  // ---------------------------------------------------------------- planilha (folha)
  function Folha(nome, est, sst) {
    const cel = new Map(); const S = { nome, links: [], altPag: 0, cols: [], alt: {}, merges: [], cf: [], dv: [], protegida: false, freeze: null, grade: false, tab: null, drawing: null, tabela: null, zoom: 100, paisagem: true, ocultas: new Set() };
    const put = (a, v, s = 0) => { const [c, r] = parse(a); if (!cel.has(r)) cel.set(r, new Map()); cel.get(r).set(c, { v, s }); };
    S.put = put;
    S.f = (a, f, s = 0, v) => put(a, { f, v }, s);
    S.get = a => { const [c, r] = parse(a); return cel.has(r) && cel.get(r).get(c); };
    // fórmula de matriz dinâmica: grava a fórmula na 1ª célula e o resultado (cache) nas vizinhas
    S.dyn = (a, f, vals, s) => {
      const [c0, r0] = parse(a); const m = Array.isArray(vals) ? vals : [[vals]];
      const h = Math.max(1, m.length), w = Math.max(1, ...m.map(x => x.length));
      for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) {
        const ad = A(c0 + j, r0 + i), old = S.get(ad), st = old ? old.s : (s || 0), val = (m[i] || [])[j];
        if (i === 0 && j === 0) put(ad, { f, v: val ?? '', dyn: `${a}:${A(c0 + w - 1, r0 + h - 1)}` }, s != null ? s : st);
        else put(ad, val ?? '', st);
      }
    };
    S.link = (a, loc, tip) => S.links.push({ a, loc, tip });
    S.merge = (a, b, v, s) => { if (v !== undefined) put(a, v, s); else if (s != null) put(a, '', s); if (s != null) { const [c1, r1] = parse(a), [c2, r2] = parse(b); for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (r !== r1 || c !== c1) put(A(c, r), '', s); } S.merges.push(`${a}:${b}`); };
    S.xml = () => {
      const rows = [...cel.keys()].sort((a, b) => a - b);
      let maxC = 0; rows.forEach(r => cel.get(r).forEach((_, c) => { if (c > maxC) maxC = c; }));
      const cellXml = (r, c, { v: v0, s }) => { let v = v0;
        const a = A(c, r), st = s ? ` s="${s}"` : '';
        if (v && typeof v === 'object' && 'f' in v) {
          const val = v.v;
          if (v.dyn) {
            const ref = v.dyn || a, fx = `<f t="array" ref="${ref}">${xe(px(v.f))}</f>`;
            if (typeof val === 'number') return `<c r="${a}"${st} cm="1">${fx}<v>${Number.isFinite(val) ? val : 0}</v></c>`;
            return `<c r="${a}"${st} cm="1" t="str">${fx}<v>${xe(val ?? '')}</v></c>`;
          }
          v = { ...v, f: px(v.f) };
          if (val === undefined || val === null || val === '') return `<c r="${a}"${st}${typeof val === 'string' ? ' t="str"' : ''}><f>${xe(v.f)}</f>${val === '' ? '<v></v>' : ''}</c>`;
          if (typeof val === 'number') return `<c r="${a}"${st}><f>${xe(v.f)}</f><v>${Number.isFinite(val) ? val : 0}</v></c>`;
          if (typeof val === 'boolean') return `<c r="${a}"${st} t="b"><f>${xe(v.f)}</f><v>${val ? 1 : 0}</v></c>`;
          return `<c r="${a}"${st} t="str"><f>${xe(v.f)}</f><v>${xe(val)}</v></c>`;
        }
        if (v === '' || v == null) return `<c r="${a}"${st}/>`;
        if (typeof v === 'number') return `<c r="${a}"${st}><v>${v}</v></c>`;
        return `<c r="${a}"${st} t="s"><v>${sst(v)}</v></c>`;
      };
      const pane = S.freeze ? (() => { const [c, r] = parse(S.freeze); return `<pane${c ? ` xSplit="${c}"` : ''}${r ? ` ySplit="${r}"` : ''} topLeftCell="${S.freeze}" activePane="${c && r ? 'bottomRight' : r ? 'bottomLeft' : 'topRight'}" state="frozen"/><selection pane="${c && r ? 'bottomRight' : r ? 'bottomLeft' : 'topRight'}" activeCell="${S.freeze}" sqref="${S.freeze}"/>`; })() : '';
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr>${S.tab ? `<tabColor rgb="FF${S.tab}"/>` : ''}<pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:${A(Math.max(maxC, 0), Math.max(rows[rows.length - 1] || 0, 0))}"/><sheetViews><sheetView workbookViewId="0"${S.grade ? '' : ' showGridLines="0"'} zoomScale="${S.zoom}" zoomScaleNormal="${S.zoom}">${pane || '<selection activeCell="A1" sqref="A1"/>'}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="18" customHeight="1"/>${S.cols.length ? `<cols>${S.cols.map((w, i) => w ? `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"${S.ocultas.has(i) ? ' hidden="1"' : ''}/>` : '').join('')}</cols>` : ''}<sheetData>${rows.map(r => { const m = cel.get(r); return `<row r="${r + 1}"${S.alt[r] ? ` ht="${S.alt[r]}" customHeight="1"` : ''}>${[...m.keys()].sort((a, b) => a - b).map(c => cellXml(r, c, m.get(c))).join('')}</row>`; }).join('')}</sheetData>${S.protegida ? '<sheetProtection sheet="1" objects="1" scenarios="1" formatColumns="0" formatRows="0" sort="0" autoFilter="0"/>' : ''}${S.merges.length ? `<mergeCells count="${S.merges.length}">${S.merges.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : ''}${S.cf.map((c, i) => `<conditionalFormatting sqref="${c.ref}">${c.rules.map((ru, j) => ru.type === 'dataBar' ? `<cfRule type="dataBar" priority="${i * 10 + j + 1}"><dataBar><cfvo type="num" val="0"/><cfvo type="num" val="1"/><color rgb="FF${ru.cor}"/></dataBar></cfRule>` : `<cfRule type="${ru.type}" dxfId="${ru.dxf}" priority="${i * 10 + j + 1}"${ru.op ? ` operator="${ru.op}"` : ''}${ru.text ? ` text="${xe(ru.text)}"` : ''}${ru.stop ? ' stopIfTrue="1"' : ''}>${(ru.f || []).map(f => `<formula>${xe(f)}</formula>`).join('')}</cfRule>`).join('')}</conditionalFormatting>`).join('')}${S.dv.length ? `<dataValidations count="${S.dv.length}">${S.dv.map(d => `<dataValidation type="${d.type}"${d.op ? ` operator="${d.op}"` : ''} allowBlank="1" showInputMessage="1" showErrorMessage="1"${d.err ? ` errorStyle="${d.estilo || 'stop'}" errorTitle="${xe(d.errT || 'Valor inválido')}" error="${xe(d.err)}"` : ''}${d.msg ? ` promptTitle="${xe(d.msgT || '')}" prompt="${xe(d.msg)}"` : ''} sqref="${d.ref}">${d.f1 != null ? `<formula1>${xe(d.f1)}</formula1>` : ''}${d.f2 != null ? `<formula2>${xe(d.f2)}</formula2>` : ''}</dataValidation>`).join('')}</dataValidations>` : ''}${S.links.length ? `<hyperlinks>${S.links.map(l => `<hyperlink ref="${l.a}" location="${xe(l.loc)}" display="${xe(l.tip || '')}"/>`).join('')}</hyperlinks>` : ''}<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.55" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="${S.paisagem ? 'landscape' : 'portrait'}" fitToWidth="1" fitToHeight="${S.altPag}"/><headerFooter><oddFooter>&amp;L&amp;8Controle de Avaria · ${xe(nome)}&amp;C&amp;8Criado por Lucas Carvalho&amp;R&amp;8Página &amp;P de &amp;N</oddFooter></headerFooter>${S.drawing ? '<drawing r:id="rIdD"/>' : ''}${S.tabela ? '<tableParts count="1"><tablePart r:id="rIdT"/></tableParts>' : ''}</worksheet>`;
    };
    return S;
  }

  // ---------------------------------------------------------------- gráficos
  const txPr = (sz = 900, cor = C.mudo, b = false) => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}" b="${b ? 1 : 0}"><a:solidFill><a:srgbClr val="${cor}"/></a:solidFill><a:latin typeface="${FONTE}"/></a:defRPr></a:pPr><a:endParaRPr lang="pt-BR"/></a:p></c:txPr>`;
  const titulo = t => `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1050" b="1"><a:solidFill><a:srgbClr val="${C.tinta}"/></a:solidFill><a:latin typeface="${FONTE}"/></a:defRPr></a:pPr><a:r><a:rPr lang="pt-BR" sz="1050" b="1"><a:solidFill><a:srgbClr val="${C.tinta}"/></a:solidFill><a:latin typeface="${FONTE}"/></a:rPr><a:t>${xe(t)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`;
  const strRef = (f, vals) => `<c:strRef><c:f>${xe(f)}</c:f><c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${xe(v)}</c:v></c:pt>`).join('')}</c:strCache></c:strRef>`;
  const numRef = (f, vals) => `<c:numRef><c:f>${xe(f)}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${+v || 0}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`;
  const solid = c => `<c:spPr><a:solidFill><a:srgbClr val="${c}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>`;
  const casca = (corpo, legenda = true, soVisivel = true) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="pt-BR"/><c:roundedCorners val="0"/><c:chart>${corpo}${legenda ? `<c:legend><c:legendPos val="t"/><c:overlay val="0"/>${txPr(900, C.tinta2)}</c:legend>` : ''}<c:plotVisOnly val="${soVisivel ? 1 : 0}"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>${txPr(900, C.mudo)}<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`;
  const eixos = (dir = 'col') => `<c:catAx><c:axId val="5001"/><c:scaling><c:orientation val="${dir === 'bar' ? 'maxMin' : 'minMax'}"/></c:scaling><c:delete val="0"/><c:axPos val="${dir === 'bar' ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="${C.linha2}"/></a:solidFill></a:ln></c:spPr>${txPr(850, C.mudo)}<c:crossAx val="5002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx><c:valAx><c:axId val="5002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${dir === 'bar' ? 't' : 'l'}"/><c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="${C.linha}"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="0" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${txPr(850, C.fraco)}<c:crossAx val="5001"/><c:crosses val="${dir === 'bar' ? 'max' : 'autoZero'}"/><c:crossBetween val="between"/></c:valAx>`;
  function grafBarras({ titulo: t, dir = 'col', grupo = 'clustered', cat, series }) {
    const ser = series.map((s, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${strRef(s.nomeRef, [s.nome])}</c:tx>${solid(s.cor)}<c:invertIfNegative val="0"/><c:cat>${strRef(cat.ref, cat.vals)}</c:cat><c:val>${numRef(s.ref, s.vals)}</c:val></c:ser>`).join('');
    return casca(`${titulo(t)}<c:plotArea><c:layout/><c:barChart><c:barDir val="${dir}"/><c:grouping val="${grupo}"/><c:varyColors val="0"/>${ser}<c:gapWidth val="${grupo === 'stacked' ? 55 : 70}"/>${grupo === 'stacked' ? '<c:overlap val="100"/>' : '<c:overlap val="-8"/>'}<c:axId val="5001"/><c:axId val="5002"/></c:barChart>${eixos(dir)}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>`);
  }
  function grafRosca({ titulo: t, cat, serie, cores }) {
    const pts = cores.map((c, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${c}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join('');
    return casca(`${titulo(t)}<c:plotArea><c:layout/><c:doughnutChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strRef(serie.nomeRef, [serie.nome])}</c:tx>${pts}<c:dLbls><c:numFmt formatCode="0%" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${txPr(900, 'FFFFFF', true)}<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/><c:showLeaderLines val="0"/></c:dLbls><c:cat>${strRef(cat.ref, cat.vals)}</c:cat><c:val>${numRef(serie.ref, serie.vals)}</c:val></c:ser><c:firstSliceAng val="0"/><c:holeSize val="62"/></c:doughnutChart><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>`, true, false);
  }
  const ancora = (i, [c1, r1, c2, r2]) => `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${c2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r2}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i + 2}" name="Gráfico ${i + 1}"/><xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdC${i + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;

  // ================================================================= GERAR (v2 · Excel 365)
  function gerar(linhas, opt = {}) {
    const agora = new Date(), hojeIso = isoLocal(agora), HOJE = serial(hojeIso);
    const cdNome = opt.cd || 'Centro de Distribuição', usuario = opt.usuario || '', demo = !!opt.demo;
    const FX = opt.faixas || [3, 7, 15, 30];
    const L = [...linhas].sort((a, b) => (a.ref - b.ref) || String(a.rel).localeCompare(String(b.rel)));
    const ORIG = ['Mercado Livre', 'Shopee', 'Avaria CD'];
    L.forEach(o => { if (o.origem && !ORIG.includes(o.origem)) ORIG.push(o.origem); });
    const RESP = [...new Set([usuario, ...L.map(o => o.por)].filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR')).slice(0, 30);
    const HIST = [...(opt.historico || [])].sort((a, b) => String(b.t).localeCompare(String(a.t)));

    const sstArr = [], sstMap = new Map();
    const sst = s => { s = String(s); if (!sstMap.has(s)) { sstMap.set(s, sstArr.length); sstArr.push(s); } return sstMap.get(s); };
    const E = Estilos();
    const NAVY = '0F2A47';
    const bd = (cor = C.linha) => ({ l: ['thin', cor], r: ['thin', cor], t: ['thin', cor], b: ['thin', cor] });
    const bb = (cor = C.linha) => ({ b: ['thin', cor] });
    const s = {
      band: E.xf({ font: { sz: 10, b: 1, cor: C.branco }, fill: NAVY }),
      bandR: E.xf({ font: { sz: 9, cor: 'C7D4E5' }, fill: NAVY, al: { h: 'right' } }),
      bandDemo: E.xf({ font: { sz: 9, b: 1, cor: '7C2D12' }, fill: 'FCD34D', al: { h: 'center' } }),
      nav: E.xf({ font: { sz: 9, cor: '334155' }, fill: 'F1F5F9', border: { b: ['thin', C.linha] }, al: { h: 'center' } }),
      navOn: E.xf({ font: { sz: 9, b: 1, cor: C.azul }, fill: C.branco, border: { b: ['thick', C.azul] }, al: { h: 'center' } }),
      tit: E.xf({ font: { sz: 18, b: 1, cor: C.tinta } }),
      sub: E.xf({ font: { sz: 10, cor: C.mudo } }),
      cdR: E.xf({ font: { sz: 10, b: 1, cor: C.tinta2 }, al: { h: 'right' } }),
      dataR: E.xf({ font: { sz: 9, cor: C.mudo }, fmt: '"Atualizado em "dd/mm/yyyy', al: { h: 'right' } }),
      emitido: E.xf({ font: { sz: 9, cor: C.mudo }, fmt: '"Emitido em "dd/mm/yyyy', al: { h: 'right' } }),
      secao: E.xf({ font: { sz: 10, b: 1, cor: C.azulEsc }, border: { b: ['medium', C.azul] }, al: { v: 'bottom' } }),
      secaoL: E.xf({ border: { b: ['medium', C.azul] } }),
      rot: E.xf({ font: { sz: 8, b: 1, cor: C.mudo }, al: { v: 'bottom' } }),
      rotC: E.xf({ font: { sz: 8, b: 1, cor: C.mudo }, al: { h: 'center' } }),
      inData: E.xf({ font: { sz: 10, b: 1, cor: C.tinta }, fill: C.input, border: bd('94A3B8'), fmt: 'dd/mm/yyyy', al: { h: 'center' }, livre: 1 }),
      inTxt: E.xf({ font: { sz: 10, b: 1, cor: C.tinta }, fill: C.input, border: bd('94A3B8'), al: { h: 'left', ind: 1 }, livre: 1 }),
      inNum: E.xf({ font: { sz: 10, b: 1, cor: C.tinta }, fill: C.input, border: bd('94A3B8'), fmt: '0', al: { h: 'center' }, livre: 1 }),
      nota: E.xf({ font: { sz: 8, i: 1, cor: C.fraco } }),
      link: E.xf({ font: { sz: 9, b: 1, cor: C.azul }, al: { h: 'right' } }),
      kRot: E.xf({ font: { sz: 8, b: 1, cor: C.mudo }, fill: C.branco, border: { l: ['thick', C.azul], t: ['thin', C.linha] }, al: { h: 'left', ind: 1, v: 'bottom' } }),
      kRotR: E.xf({ fill: C.branco, border: { t: ['thin', C.linha], r: ['thin', C.linha] } }),
      kVal: E.xf({ font: { sz: 22, b: 1, cor: C.tinta }, fill: C.branco, border: { l: ['thick', C.azul] }, fmt: '#,##0', al: { h: 'left', ind: 1 } }),
      kValP: E.xf({ font: { sz: 22, b: 1, cor: C.tinta }, fill: C.branco, border: { l: ['thick', C.azul] }, fmt: '0%', al: { h: 'left', ind: 1 } }),
      kValD: E.xf({ font: { sz: 22, b: 1, cor: C.tinta }, fill: C.branco, border: { l: ['thick', C.azul] }, fmt: '0.0', al: { h: 'left', ind: 1 } }),
      kValR: E.xf({ fill: C.branco, border: { r: ['thin', C.linha] } }),
      kNota: E.xf({ font: { sz: 8, cor: C.mudo }, fill: C.branco, border: { l: ['thick', C.azul], b: ['thin', C.linha] }, al: { h: 'left', ind: 1, v: 'top' } }),
      kNotaR: E.xf({ fill: C.branco, border: { b: ['thin', C.linha], r: ['thin', C.linha] } }),
      cab: E.xf({ font: { sz: 9, b: 1, cor: C.branco }, fill: C.azul, border: bd(C.azul), al: { h: 'center', wrap: 1 } }),
      cabE: E.xf({ font: { sz: 9, b: 1, cor: C.branco }, fill: C.azul, border: bd(C.azul), al: { h: 'left', ind: 1, wrap: 1 } }),
      tTx: E.xf({ border: bb(), al: { h: 'left', ind: 1 } }),
      tTxB: E.xf({ font: { b: 1 }, border: bb(), al: { h: 'left', ind: 1 } }),
      tNum: E.xf({ border: bb(), fmt: '#,##0', al: { h: 'center' } }),
      tNumB: E.xf({ font: { b: 1 }, border: bb(), fmt: '#,##0', al: { h: 'center' } }),
      tDec: E.xf({ border: bb(), fmt: '0.0', al: { h: 'center' } }),
      tPct: E.xf({ border: bb(), fmt: '0%', al: { h: 'center' } }),
      tPctE: E.xf({ border: bb(), fmt: '0%', al: { h: 'left', ind: 1 } }),
      tVar: E.xf({ font: { b: 1 }, border: bb(), fmt: '+0%;-0%;0%', al: { h: 'center' } }),
      tPP: E.xf({ font: { b: 1 }, border: bb(), fmt: '+0" p.p.";-0" p.p.";"0 p.p."', al: { h: 'center' } }),
      tVarD: E.xf({ font: { b: 1 }, border: bb(), fmt: '+0.0;-0.0;0.0', al: { h: 'center' } }),
      tData: E.xf({ border: bb(), fmt: 'dd/mm/yyyy', al: { h: 'center' } }),
      tCen: E.xf({ border: bb(), al: { h: 'center' } }),
      tTot: E.xf({ font: { b: 1, cor: C.azulEsc }, fill: C.azulBg, border: { t: ['medium', C.azul] }, al: { h: 'left', ind: 1 } }),
      tTotN: E.xf({ font: { b: 1, cor: C.azulEsc }, fill: C.azulBg, border: { t: ['medium', C.azul] }, fmt: '#,##0', al: { h: 'center' } }),
      tTotP: E.xf({ font: { b: 1, cor: C.azulEsc }, fill: C.azulBg, border: { t: ['medium', C.azul] }, fmt: '0%', al: { h: 'center' } }),
      tTotD: E.xf({ font: { b: 1, cor: C.azulEsc }, fill: C.azulBg, border: { t: ['medium', C.azul] }, fmt: '0.0', al: { h: 'center' } }),
      // dentro das tabelas (sem preenchimento: respeita as faixas da tabela)
      lNum: E.xf({ font: { b: 1 }, fmt: '0', al: { h: 'center' }, livre: 1 }),
      lData: E.xf({ fmt: 'dd/mm/yyyy', al: { h: 'center' }, livre: 1 }),
      lDH: E.xf({ fmt: 'dd/mm/yyyy hh:mm', al: { h: 'center' }, livre: 1 }),
      lTx: E.xf({ al: { h: 'left' }, livre: 1 }),
      lCen: E.xf({ al: { h: 'center' }, livre: 1 }),
      aTx: E.xf({ font: { cor: '475569' }, al: { h: 'center' } }),
      aNum: E.xf({ font: { cor: '475569' }, fmt: '0', al: { h: 'center' } }),
      aMes: E.xf({ font: { cor: '475569' }, fmt: 'mm/yyyy', al: { h: 'center' } }),
      aAl: E.xf({ font: { cor: '475569', sz: 9 }, al: { h: 'left' } }),
      tagIn: E.xf({ font: { sz: 8, b: 1, cor: C.ambar }, fill: C.ambarBg, al: { h: 'center' } }),
      tagOp: E.xf({ font: { sz: 8, b: 1, cor: C.mudo }, fill: C.auto, al: { h: 'center' } }),
      tagAu: E.xf({ font: { sz: 8, b: 1, cor: C.azul }, fill: C.azulCl, al: { h: 'center' } }),
      // áreas de resultado dinâmico (bordas aparecem só nas linhas com dado, via formatação condicional)
      dNum: E.xf({ font: { b: 1 }, fmt: '0;-0;;@', al: { h: 'left', ind: 1 } }),
      dData: E.xf({ fmt: 'dd/mm/yyyy;;;@', al: { h: 'center' } }),
      dCen: E.xf({ fmt: '0;-0;;@', al: { h: 'center' } }),
      dTx: E.xf({ fmt: '0;-0;;@', al: { h: 'left', ind: 1 } }),
      dDias: E.xf({ fmt: '0', al: { h: 'center' } }),
      dMsg: E.xf({ font: { i: 1, cor: C.mudo }, al: { h: 'left', ind: 1 } }),
      txt: E.xf({ font: { sz: 10, cor: C.tinta2 }, al: { wrap: 1, v: 'top' } }),
      txtB: E.xf({ font: { sz: 10, b: 1, cor: C.tinta }, al: { wrap: 1, v: 'top' } }),
      resumo: E.xf({ font: { sz: 10.5, cor: C.tinta }, fill: C.azulBg, border: { l: ['thick', C.azul] }, al: { wrap: 1, v: 'center', ind: 1 } }),
      auto: E.xf({ font: { sz: 10, cor: '475569' }, fill: C.auto, border: bd(), al: { h: 'center' } }),
      autoD: E.xf({ font: { sz: 10, cor: '475569' }, fill: C.auto, border: bd(), fmt: 'dd/mm/yyyy', al: { h: 'center' } }),
      autoE: E.xf({ font: { sz: 10, cor: '475569' }, fill: C.auto, border: bd(), al: { h: 'left', ind: 1 } }),
      chipIn: E.xf({ font: { sz: 9, b: 1, cor: C.tinta }, fill: C.input, border: bd('94A3B8'), al: { h: 'center' } }),
      chipAu: E.xf({ font: { sz: 9, b: 1, cor: '475569' }, fill: C.auto, border: bd(), al: { h: 'center' } }),
      chipPr: E.xf({ font: { sz: 9, b: 1, cor: C.branco }, fill: C.azul, border: bd(C.azul), al: { h: 'center' } }),
    };
    const dx = {
      tabInteira: E.dxf({ borda: C.linha }), tabCab: E.dxf({ cor: C.branco, bg: C.azul, b: 1 }), tabFaixa: E.dxf({ bg: 'F7F9FC' }),
      ok: E.dxf({ cor: C.azul, bg: C.azulCl, b: 1 }), pend: E.dxf({ cor: C.ambar, bg: C.ambarBg, b: 1 }), ruim: E.dxf({ cor: C.verm, bg: C.vermBg, b: 1 }),
      ambarTx: E.dxf({ cor: C.ambar, b: 1 }), vermTx: E.dxf({ cor: C.verm, b: 1 }), azulTx: E.dxf({ cor: C.azul, b: 1 }),
      linhaDado: E.dxf({ borda: C.linha }), apagado: E.dxf({ cor: 'CBD5E1' }),
    };

    // ---------------------------------------------------------------- base e cálculos (cache)
    const R0 = 10; // 1ª linha de dados em LANÇAMENTOS
    const cont = {}; L.forEach(o => { if (o.ref !== '' && o.ref != null) cont[o.ref] = (cont[o.ref] || 0) + 1; });
    const D = L.map((o, i) => {
      const rel = serial(o.rel), ent = serial(o.ent), pend = ent == null;
      const dias = rel == null ? '' : pend ? HOJE - rel : ent - rel;
      let alerta = '';
      if (o.ref !== '' && cont[o.ref] > 1) alerta = 'Ref. duplicada';
      else if (o.ref === '' || o.ref == null || rel == null || !o.origem || !o.sku || !o.produto) alerta = 'Campo obrigatório vazio';
      else if (rel > HOJE) alerta = 'Data do relatório no futuro';
      else if (!pend && ent < rel) alerta = 'Entrega antes do relatório';
      else if (!pend && ent > HOJE) alerta = 'Data de entrega no futuro';
      else if (!pend && !o.por) alerta = 'Entregue sem responsável';
      return { ref: o.ref, relS: rel, origem: o.origem || '', sku: String(o.sku ?? ''), produto: o.produto || '', entS: ent, por: o.por || '', obs: o.obs || '', pend, sit: pend ? 'Não entregue' : 'Entregue', dias, mes: rel == null ? '' : serial(String(o.rel).slice(0, 7) + '-01'), alerta, i };
    });
    const media = a => a.length ? a.reduce((x, y) => x + y, 0) / a.length : '';
    const mediaEnt = a => media(a.filter(d => !d.pend && d.dias !== '').map(d => d.dias));
    const entre = (v, a, b) => v != null && v !== '' && v >= a && v <= b;

    // nomes da tabela
    const T = 'tbAvarias', last = R0 + Math.max(D.length, 1) - 1, MAXR = Math.max(last + 1000, 3000);
    const col = n => `${T}[${n}]`, eu = n => `${T}[[#This Row],[${n}]]`;
    const COLS = [
      { n: 'Ref.', w: 10, tipo: 'in' }, { n: 'Data do relatório', w: 14, tipo: 'in' }, { n: 'Origem', w: 15, tipo: 'in' }, { n: 'Código', w: 13, tipo: 'in' }, { n: 'Descrição do produto', w: 42, tipo: 'in' },
      { n: 'Data de entrega', w: 14, tipo: 'op' }, { n: 'Entregue por', w: 20, tipo: 'op' }, { n: 'Observação', w: 26, tipo: 'op' },
      { n: 'Situação', w: 14, tipo: 'au', f: `IF(AND(${eu('Ref.')}="",${eu('Descrição do produto')}=""),"",IF(${eu('Data de entrega')}="","Não entregue","Entregue"))` },
      { n: 'Dias', w: 12, tipo: 'au', f: `IF(NOT(ISNUMBER(${eu('Data do relatório')})),"",IF(${eu('Data de entrega')}="",cfgHoje-${eu('Data do relatório')},${eu('Data de entrega')}-${eu('Data do relatório')}))` },
      { n: 'Mês', w: 12, tipo: 'au', f: `IF(ISNUMBER(${eu('Data do relatório')}),DATE(YEAR(${eu('Data do relatório')}),MONTH(${eu('Data do relatório')}),1),"")` },
      { n: 'Alerta', w: 28, tipo: 'au', f: `IF(AND(${eu('Ref.')}="",${eu('Data do relatório')}="",${eu('Código')}="",${eu('Descrição do produto')}=""),"",IF(AND(${eu('Ref.')}<>"",COUNTIF(${col('Ref.')},${eu('Ref.')})>1),"Ref. duplicada",IF(OR(${eu('Ref.')}="",${eu('Data do relatório')}="",${eu('Origem')}="",${eu('Código')}="",${eu('Descrição do produto')}=""),"Campo obrigatório vazio",IF(${eu('Data do relatório')}>cfgHoje,"Data do relatório no futuro",IF(AND(${eu('Data de entrega')}<>"",${eu('Data de entrega')}<${eu('Data do relatório')}),"Entrega antes do relatório",IF(AND(${eu('Data de entrega')}<>"",${eu('Data de entrega')}>cfgHoje),"Data de entrega no futuro",IF(AND(${eu('Data de entrega')}<>"",${eu('Entregue por')}=""),"Entregue sem responsável","")))))))` },
    ];
    const ci = n => COLS.findIndex(c => c.n === n) + 1; // +1: coluna A é margem
    const cl = n => colL(ci(n));
    const RB = col('Data do relatório'), RE = col('Data de entrega'), RO = col('Origem'), RS = col('Situação'), RD = col('Dias'), RP = col('Entregue por'), RA = col('Alerta'), RR = col('Ref.');
    const pick = idx => `INDEX(${T},SEQUENCE(ROWS(${T})),{${idx.join(',')}})`; // escolhe colunas da tabela
    const IDX = n => COLS.findIndex(c => c.n === n) + 1;

    // ---------------------------------------------------------------- moldura comum (faixa + navegação)
    const ABAS = [['PAINEL', 'Painel'], ['CENTRAL', 'Central'], ['LANÇAMENTOS', 'Lançamentos'], ['HISTÓRICO', 'Histórico'], ['RELATÓRIO', 'Relatório'], ['INDICADORES', 'Indicadores'], ['CONFIGURAÇÕES', 'Config.'], ['AJUDA', 'Ajuda']];
    function moldura(F, ultimaCol, titulo, sub) {
      F.alt[0] = 26; F.alt[1] = 22; F.alt[2] = 8; F.alt[3] = 30; F.alt[4] = 18; F.alt[5] = 8;
      for (let c = 1; c <= ultimaCol; c++) F.put(A(c, 0), '', s.band);
      F.put('B1', 'CONTROLE DE AVARIA', s.band);
      if (demo) F.merge(A(ultimaCol - 2, 0), A(ultimaCol, 0), 'DADOS DEMONSTRATIVOS', s.bandDemo);
      else F.merge(A(ultimaCol - 2, 0), A(ultimaCol, 0), { f: 'cfgCD', v: cdNome }, s.bandR);
      ABAS.forEach(([nome, rot], i) => { const a = A(1 + i, 1); F.put(a, rot, nome === F.nome ? s.navOn : s.nav); if (nome !== F.nome) F.link(a, `'${nome}'!A1`, `Ir para ${rot}`); });
      for (let c = 1 + ABAS.length; c <= ultimaCol; c++) F.put(A(c, 1), '', s.nav);
      F.put('B4', titulo, s.tit); F.put('B5', sub, s.sub);
    }
    const secao = (F, a, txt, ate) => { const [c, r] = parse(a); F.put(a, txt, s.secao); for (let k = c + 1; k <= colN(ate); k++) F.put(A(k, r), '', s.secaoL); };
    const cabec = (F, linha, inicio, titulos, esq = []) => titulos.forEach((t, i) => F.put(A(colN(inicio) + i, linha - 1), t, esq.includes(i) ? s.cabE : s.cab));
    const dataBR = x => `TEXT(DAY(${x}),"00")&"/"&TEXT(MONTH(${x}),"00")&"/"&YEAR(${x})`; // independente do idioma do Excel
    const br = n => { if (n == null || n === '') return ''; const d = new Date(Date.UTC(1899, 11, 30) + n * 864e5); return `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`; };
    const pct = (a, b) => b ? a / b : 0;
    const sinal = v => (v > 0 ? '+' : v < 0 ? '-' : '') + Math.round(Math.abs(v) * 100) + '%';

    // ================================================================ CONFIGURAÇÕES
    const cfg = Folha('CONFIGURAÇÕES', E, sst); cfg.tab = '64748B'; cfg.protegida = true;
    cfg.cols = [2, 30, 22, 26, 16, 20, 32, 14, 12];
    moldura(cfg, 8, 'Configurações', 'Parâmetros e listas usados em todo o sistema. Edite apenas as células em creme.');
    secao(cfg, 'B7', 'PARÂMETROS', 'E');
    cfg.put('B8', 'Nome do CD / unidade', s.tTxB); cfg.put('C8', cdNome, s.inTxt); cfg.put('D8', 'Aparece no topo de todas as abas.', s.nota);
    cfg.put('B9', 'Data de referência (hoje)', s.tTxB); cfg.f('C9', 'TODAY()', s.autoD, HOJE); cfg.put('D9', 'Automática.', s.nota);
    secao(cfg, 'B11', 'FAIXAS DE TEMPO EM ABERTO (dias)', 'E');
    cfg.put('B12', 'Faixa', s.cabE); cfg.put('C12', 'Até (dias)', s.cab); cfg.put('D12', 'Rótulo (automático)', s.cabE);
    const rotFx = [`"0 a "&C13&" dias"`, `C13+1&" a "&C14&" dias"`, `C14+1&" a "&C15&" dias"`, `C15+1&" a "&C16&" dias"`, `"Mais de "&C16&" dias"`];
    const rotFxV = [`0 a ${FX[0]} dias`, `${FX[0] + 1} a ${FX[1]} dias`, `${FX[1] + 1} a ${FX[2]} dias`, `${FX[2] + 1} a ${FX[3]} dias`, `Mais de ${FX[3]} dias`];
    for (let i = 0; i < 5; i++) {
      const r = 13 + i; cfg.put(`B${r}`, `Faixa ${i + 1}`, s.tTx);
      if (i < 4) cfg.put(`C${r}`, FX[i], s.inNum); else cfg.put(`C${r}`, '—', s.auto);
      cfg.f(`D${r}`, rotFx[i], s.autoE, rotFxV[i]);
    }
    cfg.put('E13', 'Só organizam os gráficos e relatórios. Não são prazos.', s.nota);
    cfg.dv.push({ type: 'whole', op: 'between', f1: '1', f2: '3650', ref: 'C13:C16', err: 'Use números inteiros em ordem crescente (ex.: 3, 7, 15, 30).', msgT: 'Faixa', msg: 'Limite superior da faixa, em dias.' });
    secao(cfg, 'B19', 'LISTAS SUSPENSAS', 'F');
    cfg.put('B20', 'Origens', s.cabE); cfg.put('C20', '', s.cab); cfg.put('D20', 'Responsáveis pela entrega', s.cabE); cfg.put('E20', 'Filtro: origem', s.cab); cfg.put('F20', 'Filtro: responsável', s.cab);
    const LN = 30;
    for (let i = 0; i < LN; i++) {
      const r = 21 + i;
      cfg.put(`B${r}`, ORIG[i] || '', s.inTxt); cfg.put(`D${r}`, RESP[i] || '', s.inTxt);
      cfg.f(`E${22 + i}`, `IF(B${r}="","",B${r})`, s.auto, ORIG[i] || ''); cfg.f(`F${22 + i}`, `IF(D${r}="","",D${r})`, s.auto, RESP[i] || '');
    }
    cfg.put('E21', 'Todas', s.auto); cfg.put('F21', 'Todas', s.auto);
    secao(cfg, 'B53', 'QUALIDADE DAS LISTAS (automático)', 'F');
    cfg.put('B54', 'Origens usadas na base e fora da lista', s.cabE); cfg.put('D54', 'Responsáveis usados na base e fora da lista', s.cabE);
    const foraO = [...new Set(D.map(d => d.origem).filter(Boolean))].filter(o => !ORIG.includes(o));
    const foraR = [...new Set(D.map(d => d.por).filter(Boolean))].filter(o => !RESP.includes(o));
    cfg.dyn('B55', `LET(_xlpm.u,UNIQUE(FILTER(${RO},${RO}<>"","")),_xlpm.f,FILTER(_xlpm.u,COUNTIF($B$21:$B$50,_xlpm.u)=0,""),IF(_xlpm.f="","Nenhuma — lista completa.",_xlpm.f))`, foraO.length ? foraO.map(x => [x]) : [['Nenhuma — lista completa.']], s.dTx);
    cfg.dyn('D55', `LET(_xlpm.u,UNIQUE(FILTER(${RP},${RP}<>"","")),_xlpm.f,FILTER(_xlpm.u,COUNTIF($D$21:$D$50,_xlpm.u)=0,""),IF(_xlpm.f="","Nenhum — lista completa.",_xlpm.f))`, foraR.length ? foraR.map(x => [x]) : [['Nenhum — lista completa.']], s.dTx);
    secao(cfg, 'G7', 'CÁLCULOS INTERNOS — não editar', 'H');
    const PI = `'PAINEL'!$C$8`, PF = `'PAINEL'!$F$8`, PO = `'PAINEL'!$I$8`, PR = `'PAINEL'!$L$8`, QI = `'RELATÓRIO'!$C$7`, QF = `'RELATÓRIO'!$E$7`;
    const mesIni = serial(hojeIso.slice(0, 7) + '-01');
    const repDe = mesIni, repAte = HOJE, repLen = repAte - repDe + 1;
    const internos = [
      ['fDe', 'Painel: data inicial', `IF(ISNUMBER(${PI}),${PI},1)`, 1, 1],
      ['fAte', 'Painel: data final', `IF(ISNUMBER(${PF}),${PF},2958465)`, 2958465, 1],
      ['fOrig', 'Painel: origem (critério)', `IF(OR(${PO}="",${PO}="Todas"),"*",${PO})`, '*', 0],
      ['fResp', 'Painel: responsável (critério)', `IF(OR(${PR}="",${PR}="Todas"),"<>#todos#",${PR})`, '<>#todos#', 0],
      ['fTem', 'Painel: período definido?', `AND(ISNUMBER(${PI}),ISNUMBER(${PF}))`, false, 0],
      ['fDe0', 'Painel: anterior — início', `IF(fTem,fDe-(fAte-fDe+1),1)`, 1, 1],
      ['fAte0', 'Painel: anterior — fim', `IF(fTem,fDe-1,0)`, 0, 1],
      ['rDe', 'Relatório: data inicial', `IF(ISNUMBER(${QI}),${QI},1)`, repDe, 1],
      ['rAte', 'Relatório: data final', `IF(ISNUMBER(${QF}),${QF},2958465)`, repAte, 1],
      ['rTem', 'Relatório: período definido?', `AND(ISNUMBER(${QI}),ISNUMBER(${QF}))`, true, 0],
      ['rDe0', 'Relatório: anterior — início', `IF(rTem,rDe-(rAte-rDe+1),1)`, repDe - repLen, 1],
      ['rAte0', 'Relatório: anterior — fim', `IF(rTem,rDe-1,0)`, repDe - 1, 1],
    ];
    internos.forEach(([n, t, f, v, d], i) => { cfg.put(`G${8 + i}`, t, s.tTx); cfg.f(`H${8 + i}`, f, d ? s.autoD : s.auto, v); });
    const names = [
      ['cfgCD', `'CONFIGURAÇÕES'!$C$8`], ['cfgHoje', `'CONFIGURAÇÕES'!$C$9`],
      ['fx1', `'CONFIGURAÇÕES'!$C$13`], ['fx2', `'CONFIGURAÇÕES'!$C$14`], ['fx3', `'CONFIGURAÇÕES'!$C$15`], ['fx4', `'CONFIGURAÇÕES'!$C$16`],
      ['lstOrigem', `OFFSET('CONFIGURAÇÕES'!$B$21,0,0,MAX(1,COUNTIF('CONFIGURAÇÕES'!$B$21:$B$50,"?*")),1)`],
      ['lstResp', `OFFSET('CONFIGURAÇÕES'!$D$21,0,0,MAX(1,COUNTIF('CONFIGURAÇÕES'!$D$21:$D$50,"?*")),1)`],
      ['lstOrigemF', `OFFSET('CONFIGURAÇÕES'!$E$21,0,0,1+COUNTIF('CONFIGURAÇÕES'!$B$21:$B$50,"?*"),1)`],
      ['lstRespF', `OFFSET('CONFIGURAÇÕES'!$F$21,0,0,1+COUNTIF('CONFIGURAÇÕES'!$D$21:$D$50,"?*"),1)`],
      ...internos.map(([n], i) => [n, `'CONFIGURAÇÕES'!$H$${8 + i}`]),
    ];

    // ================================================================ LANÇAMENTOS
    const lan = Folha('LANÇAMENTOS', E, sst); lan.tab = C.tinta; lan.freeze = 'C10';
    lan.cols = [2, ...COLS.map(c => c.w)];
    moldura(lan, COLS.length, 'Lançamentos', 'Banco de dados: uma linha por avaria. Para cadastrar, digite na primeira linha vazia logo abaixo da tabela — ela cresce sozinha.');
    lan.alt[6] = 20; lan.alt[7] = 16; lan.alt[8] = 30;
    lan.put('B7', 'Editável', s.chipIn); lan.put('C7', 'Automático', s.chipAu); lan.merge('D7', 'F7', '', 0);
    lan.put('D7', 'Campos com OBRIGATÓRIO precisam ser preenchidos. Os AUTOMÁTICOS não devem ser digitados.', s.nota);
    if (demo) lan.put('J7', 'ATENÇÃO: dados demonstrativos. Apague as linhas e lance os seus.', E.xf({ font: { sz: 9, b: 1, cor: '7C2D12' } }));
    COLS.forEach((c, i) => lan.put(A(1 + i, 7), c.tipo === 'in' ? 'OBRIGATÓRIO' : c.tipo === 'op' ? 'OPCIONAL' : 'AUTOMÁTICO', c.tipo === 'in' ? s.tagIn : c.tipo === 'op' ? s.tagOp : s.tagAu));
    COLS.forEach((c, i) => lan.put(A(1 + i, 8), c.n, s.cab));
    (D.length ? D : [null]).forEach((d, i) => {
      const r = R0 + i, p = (n, v, st) => lan.put(`${cl(n)}${r}`, v, st);
      p('Ref.', d ? d.ref : '', s.lNum); p('Data do relatório', d && d.relS != null ? d.relS : '', s.lData); p('Origem', d ? d.origem : '', s.lCen); p('Código', d ? d.sku : '', s.lCen);
      p('Descrição do produto', d ? d.produto : '', s.lTx); p('Data de entrega', d && d.entS != null ? d.entS : '', s.lData); p('Entregue por', d ? d.por : '', s.lTx); p('Observação', d ? d.obs : '', s.lTx);
      const v = d ? { 'Situação': d.sit, 'Dias': d.dias, 'Mês': d.mes, 'Alerta': d.alerta } : { 'Situação': '', 'Dias': '', 'Mês': '', 'Alerta': '' };
      COLS.filter(c => c.f).forEach(c => lan.f(`${cl(c.n)}${r}`, c.f, c.n === 'Situação' ? s.aTx : c.n === 'Dias' ? s.aNum : c.n === 'Mês' ? s.aMes : s.aAl, v[c.n]));
    });
    const tabRef = `B9:${colL(COLS.length)}${last}`;
    const tabela = (id, nome, ref, cols) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${nome}" displayName="${nome}" ref="${ref}" totalsRowShown="0"><autoFilter ref="${ref}"/><tableColumns count="${cols.length}">${cols.map((c, i) => `<tableColumn id="${i + 1}" name="${xe(c.n)}"${c.f ? `><calculatedColumnFormula>${xe(px(c.f))}</calculatedColumnFormula></tableColumn>` : '/>'}`).join('')}</tableColumns><tableStyleInfo name="AvariaCorporativo" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
    const tab1 = tabela(1, T, tabRef, COLS); lan.tabela = true;
    const lr = n => `${cl(n)}${R0}:${cl(n)}${MAXR}`, c0 = n => `${cl(n)}${R0}`;
    lan.cf.push({ ref: lr('Situação'), rules: [{ type: 'cellIs', op: 'equal', f: ['"Entregue"'], dxf: dx.ok }, { type: 'cellIs', op: 'equal', f: ['"Não entregue"'], dxf: dx.pend }] });
    lan.cf.push({ ref: lr('Alerta'), rules: [
      { type: 'expression', f: [`$${c0('Alerta')}="Entregue sem responsável"`], dxf: dx.pend, stop: 1 },
      { type: 'expression', f: [`$${c0('Alerta')}<>""`], dxf: dx.ruim }] });
    lan.cf.push({ ref: lr('Ref.'), rules: [{ type: 'expression', f: [`AND($${c0('Ref.')}<>"",COUNTIF($${cl('Ref.')}$${R0}:$${cl('Ref.')}$${MAXR},$${c0('Ref.')})>1)`], dxf: dx.ruim }] });
    ['Ref.', 'Data do relatório', 'Origem', 'Código', 'Descrição do produto'].forEach(n => lan.cf.push({ ref: lr(n), rules: [{ type: 'expression', f: [`AND($${c0('Alerta')}="Campo obrigatório vazio",${c0(n)}="")`], dxf: E.dxf({ bg: 'FDECEC' }) }] }));
    lan.dv.push({ type: 'custom', f1: `AND(ISNUMBER(${c0('Ref.')}),${c0('Ref.')}>=1,INT(${c0('Ref.')})=${c0('Ref.')},COUNTIF($${cl('Ref.')}$${R0}:$${cl('Ref.')}$${MAXR},${c0('Ref.')})=1)`, ref: lr('Ref.'), err: 'A Ref. deve ser um número inteiro maior que zero e não pode repetir uma Ref. já lançada.', errT: 'Ref. inválida ou repetida', msgT: 'Ref. (obrigatório)', msg: 'Número da referência. Não pode repetir.' });
    lan.dv.push({ type: 'custom', f1: `AND(ISNUMBER(${c0('Data do relatório')}),${c0('Data do relatório')}>=DATE(2020,1,1),${c0('Data do relatório')}<=cfgHoje)`, ref: lr('Data do relatório'), err: 'Digite uma data válida (dd/mm/aaaa), a partir de 2020 e não posterior a hoje.', msgT: 'Data do relatório (obrigatório)', msg: 'Formato dd/mm/aaaa.' });
    lan.dv.push({ type: 'list', f1: 'lstOrigem', ref: lr('Origem'), err: 'Escolha uma origem da lista. Para incluir outra, use a aba CONFIGURAÇÕES.', msgT: 'Origem (obrigatório)', msg: 'Escolha na lista.' });
    lan.dv.push({ type: 'textLength', op: 'between', f1: '1', f2: '40', ref: lr('Código'), err: 'O código deve ter entre 1 e 40 caracteres.', msgT: 'Código (obrigatório)', msg: 'Código do produto no WMS.' });
    lan.dv.push({ type: 'textLength', op: 'between', f1: '2', f2: '250', ref: lr('Descrição do produto'), err: 'A descrição deve ter entre 2 e 250 caracteres.', msgT: 'Descrição (obrigatório)', msg: 'Descrição do produto avariado.' });
    lan.dv.push({ type: 'custom', f1: `OR(${c0('Data de entrega')}="",AND(ISNUMBER(${c0('Data de entrega')}),${c0('Data de entrega')}>=${c0('Data do relatório')},${c0('Data de entrega')}<=cfgHoje))`, ref: lr('Data de entrega'), err: 'A data de entrega deve ser igual ou posterior à data do relatório e não pode ser futura.', msgT: 'Data de entrega (opcional)', msg: 'Preencha quando o relatório for entregue. A situação muda sozinha.' });
    lan.dv.push({ type: 'list', f1: 'lstResp', ref: lr('Entregue por'), estilo: 'warning', err: 'Esse nome não está na lista de responsáveis (aba CONFIGURAÇÕES). Deseja continuar?', errT: 'Responsável fora da lista', msgT: 'Entregue por (opcional)', msg: 'Quem entregou o relatório.' });
    lan.dv.push({ type: 'textLength', op: 'lessThanOrEqual', f1: '300', ref: lr('Observação'), err: 'Use até 300 caracteres.', msgT: 'Observação (opcional)', msg: 'Informação complementar.' });

    // ================================================================ HISTÓRICO
    const his = Folha('HISTÓRICO', E, sst); his.tab = C.tinta; his.freeze = 'B10';
    const HC = [{ n: 'Data e hora', w: 18 }, { n: 'Ref.', w: 10 }, { n: 'Ação', w: 22 }, { n: 'Detalhe', w: 28 }, { n: 'Responsável', w: 22 }, { n: 'Fonte', w: 12 }, { n: 'Produto (automático)', w: 40, f: `IF(${'tbHistorico'}[[#This Row],[Ref.]]="","",XLOOKUP(${'tbHistorico'}[[#This Row],[Ref.]],${RR},${col('Descrição do produto')},"Ref. não encontrada"))` }];
    his.cols = [2, ...HC.map(c => c.w), 11];
    moldura(his, 8, 'Histórico de alterações', 'Registro de quem marcou ou trocou a data de entrega e quando. As linhas vindas do sistema são preenchidas automaticamente.');
    his.alt[6] = 30;
    his.merge('B7', 'H7', 'O Excel não registra alterações sozinho sem macros. Ao lançar ou corrigir uma entrega diretamente na planilha, registre aqui (Fonte = Manual) para manter o histórico completo.', E.xf({ font: { sz: 9, cor: C.ambar }, fill: C.ambarBg, al: { wrap: 1, ind: 1 } }));
    HC.forEach((c, i) => his.put(A(1 + i, 8), c.n, s.cab));
    const descDe = {}; D.forEach(d => { if (!(d.ref in descDe)) descDe[d.ref] = d.produto; });
    (HIST.length ? HIST : [null]).forEach((h, i) => {
      const r = 10 + i;
      his.put(`B${r}`, h ? h.serie : '', s.lDH); his.put(`C${r}`, h ? h.ref : '', s.lNum); his.put(`D${r}`, h ? h.acao : '', s.lTx); his.put(`E${r}`, h ? h.detalhe : '', s.lTx); his.put(`F${r}`, h ? h.por : '', s.lTx); his.put(`G${r}`, h ? 'Sistema' : '', s.lCen);
      his.f(`H${r}`, HC[6].f, s.aAl, h ? (h.ref in descDe ? descDe[h.ref] : 'Ref. não encontrada') : '');
    });
    const tab2 = tabela(2, 'tbHistorico', `B9:H${10 + Math.max(HIST.length, 1) - 1}`, HC); his.tabela = true;
    his.dv.push({ type: 'list', f1: '"Sistema,Manual"', ref: 'G10:G5000', err: 'Use Sistema ou Manual.', msgT: 'Fonte', msg: 'Manual = registrado direto na planilha.' });
    his.dv.push({ type: 'custom', f1: 'OR(B10="",ISNUMBER(B10))', ref: 'B10:B5000', err: 'Digite data e hora (dd/mm/aaaa hh:mm).', msgT: 'Data e hora', msg: 'Ex.: 08/11/2026 14:30' });

    // ================================================================ INDICADORES
    const ind = Folha('INDICADORES', E, sst); ind.tab = '475569'; ind.protegida = true;
    ind.cols = [2, 30, 14, 14, 14, 14, 14, 14, 14, 3];
    moldura(ind, 8, 'Central de indicadores', 'Base de cálculo do Painel: comparativos, evolução, origens, tempo em aberto e produtividade.');
    const base = `${RB},">="&fDe,${RB},"<="&fAte,${RO},fOrig,${RP},fResp`;
    const base0 = `${RB},">="&fDe0,${RB},"<="&fAte0,${RO},fOrig,${RP},fResp`;
    const fil = { de: 1, ate: 2958465 }; // cache: filtros do Painel vazios
    const noPer = D.filter(d => entre(d.relS, fil.de, fil.ate));
    const nTot = noPer.length, nEnt = noPer.filter(d => !d.pend).length, nPen = nTot - nEnt;
    const mEnt = mediaEnt(noPer), maxPend = Math.max(0, ...noPer.filter(d => d.pend && d.dias !== '').map(d => d.dias));
    // comparativo do período do Painel
    secao(ind, 'B7', 'PERÍODO DO PAINEL × PERÍODO ANTERIOR', 'I');
    cabec(ind, 8, 'B', ['Indicador', 'Período', 'Período anterior', 'Variação'], [0]);
    ind.f('F8', `IF(fTem,"Período: "&${dataBR('fDe')}&" a "&${dataBR('fAte')},"Período: todo o histórico (defina datas no Painel para comparar)")`, s.nota, 'Período: todo o histórico (defina datas no Painel para comparar)');
    const comp = [
      ['Avarias registradas', `COUNTIFS(${base})`, `COUNTIFS(${base0})`, nTot, 0, s.tNum, 'var'],
      ['Entregues', `COUNTIFS(${base},${RS},"Entregue")`, `COUNTIFS(${base0},${RS},"Entregue")`, nEnt, 0, s.tNum, 'var'],
      ['Pendentes', `COUNTIFS(${base},${RS},"Não entregue")`, `COUNTIFS(${base0},${RS},"Não entregue")`, nPen, 0, s.tNum, 'var'],
      ['Taxa de entrega', `IF(C9=0,0,C10/C9)`, `IF(D9=0,0,D10/D9)`, pct(nEnt, nTot), 0, s.tPct, 'pp'],
      ['Tempo médio até a entrega (dias)', `IFERROR(AVERAGEIFS(${RD},${base},${RS},"Entregue"),0)`, `IFERROR(AVERAGEIFS(${RD},${base0},${RS},"Entregue"),0)`, mEnt === '' ? 0 : mEnt, 0, s.tDec, 'd'],
      ['Pendência mais antiga (dias)', `MAXIFS(${RD},${base},${RS},"Não entregue")`, `""`, maxPend, '', s.tNum, 'none'],
    ];
    comp.forEach(([t, f1, f0, v1, v0, st, tipo], i) => {
      const r = 9 + i; ind.put(`B${r}`, t, s.tTxB); ind.f(`C${r}`, f1, st, v1); ind.f(`D${r}`, f0, st, v0);
      if (tipo === 'var') ind.f(`E${r}`, `IF(OR(NOT(fTem),D${r}=0),"",C${r}/D${r}-1)`, s.tVar, '');
      else if (tipo === 'pp') ind.f(`E${r}`, `IF(NOT(fTem),"",(C${r}-D${r})*100)`, s.tPP, '');
      else if (tipo === 'd') ind.f(`E${r}`, `IF(OR(NOT(fTem),D${r}=0),"",C${r}-D${r})`, s.tVarD, '');
      else ind.put(`E${r}`, '', s.tCen);
    });
    ind.cf.push({ ref: 'E9:E10 E12', rules: [{ type: 'cellIs', op: 'greaterThan', f: ['0'], dxf: dx.azulTx }, { type: 'cellIs', op: 'lessThan', f: ['0'], dxf: dx.ambarTx }] });
    ind.cf.push({ ref: 'E11 E13', rules: [{ type: 'cellIs', op: 'greaterThan', f: ['0'], dxf: dx.ambarTx }, { type: 'cellIs', op: 'lessThan', f: ['0'], dxf: dx.azulTx }] });
    // evolução mensal
    secao(ind, 'B17', 'EVOLUÇÃO MENSAL — ÚLTIMOS 12 MESES (filtros de origem e responsável do Painel)', 'I');
    cabec(ind, 18, 'B', ['Mês', 'Registradas', 'Entregues no mês', 'Pendentes (do mês)', '% resolvido', 'Tempo médio (dias)', 'Início do mês'], [0]);
    ind.alt[17] = 30;
    const meses = []; const hj = new Date(agora.getFullYear(), agora.getMonth(), 1);
    for (let i = 11; i >= 0; i--) { const d = new Date(hj.getFullYear(), hj.getMonth() - i, 1); meses.push({ ini: serial(isoLocal(d)), fim: serial(isoLocal(new Date(d.getFullYear(), d.getMonth() + 1, 1))), rot: MES[d.getMonth()] + '/' + String(d.getFullYear()).slice(2) }); }
    const fob = `${RO},fOrig,${RP},fResp`;
    const mc = meses.map(m => { const reg = D.filter(d => d.relS >= m.ini && d.relS < m.fim), en = D.filter(d => d.entS != null && d.entS >= m.ini && d.entS < m.fim); return { reg: reg.length, ent: en.length, pen: reg.filter(d => d.pend).length, med: mediaEnt(en) }; });
    meses.forEach((m, i) => {
      const r = 19 + i, c = mc[i];
      ind.f(`H${r}`, `DATE(YEAR(cfgHoje),MONTH(cfgHoje)-${11 - i},1)`, s.autoD, m.ini);
      ind.f(`B${r}`, `CHOOSE(MONTH($H${r}),"jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez")&"/"&RIGHT(YEAR($H${r}),2)`, s.tTxB, m.rot);
      ind.f(`C${r}`, `COUNTIFS(${RB},">="&$H${r},${RB},"<"&EDATE($H${r},1),${fob})`, s.tNum, c.reg);
      ind.f(`D${r}`, `COUNTIFS(${RE},">="&$H${r},${RE},"<"&EDATE($H${r},1),${fob})`, s.tNum, c.ent);
      ind.f(`E${r}`, `COUNTIFS(${RB},">="&$H${r},${RB},"<"&EDATE($H${r},1),${fob},${RS},"Não entregue")`, s.tNum, c.pen);
      ind.f(`F${r}`, `IF(C${r}=0,"",(C${r}-E${r})/C${r})`, s.tPct, c.reg ? (c.reg - c.pen) / c.reg : '');
      ind.f(`G${r}`, `IFERROR(AVERAGEIFS(${RD},${RE},">="&$H${r},${RE},"<"&EDATE($H${r},1),${fob}),"")`, s.tDec, c.med);
    });
    const sm = k => mc.reduce((a, c) => a + c[k], 0);
    ind.put('B31', 'Total 12 meses', s.tTot);
    ind.f('C31', 'SUM(C19:C30)', s.tTotN, sm('reg')); ind.f('D31', 'SUM(D19:D30)', s.tTotN, sm('ent')); ind.f('E31', 'SUM(E19:E30)', s.tTotN, sm('pen'));
    ind.f('F31', 'IF(C31=0,"",(C31-E31)/C31)', s.tTotP, sm('reg') ? (sm('reg') - sm('pen')) / sm('reg') : '');
    ind.f('G31', `IFERROR(AVERAGEIFS(${RD},${RE},">="&H19,${fob}),"")`, s.tTotD, mediaEnt(D.filter(d => d.entS != null && d.entS >= meses[0].ini))); ind.put('H31', '', s.tTot);
    // comparativos fixos
    secao(ind, 'B33', 'COMPARATIVOS FIXOS', 'I');
    cabec(ind, 34, 'B', ['Indicador', 'Mês atual', 'Mês anterior', 'Variação', 'Últimos 30 dias', '30 dias anteriores', 'Variação'], [0]);
    ind.alt[33] = 30;
    const jan = (a, b, k) => D.filter(d => d[k] != null && d[k] >= a && d[k] <= b);
    const fx = [
      ['Avarias registradas', 'C30', 'C29', mc[11].reg, mc[10].reg, `COUNTIFS(${RB},">="&(cfgHoje-29),${RB},"<="&cfgHoje,${fob})`, `COUNTIFS(${RB},">="&(cfgHoje-59),${RB},"<="&(cfgHoje-30),${fob})`, jan(HOJE - 29, HOJE, 'relS').length, jan(HOJE - 59, HOJE - 30, 'relS').length, s.tNum],
      ['Entregas realizadas', 'D30', 'D29', mc[11].ent, mc[10].ent, `COUNTIFS(${RE},">="&(cfgHoje-29),${RE},"<="&cfgHoje,${fob})`, `COUNTIFS(${RE},">="&(cfgHoje-59),${RE},"<="&(cfgHoje-30),${fob})`, jan(HOJE - 29, HOJE, 'entS').length, jan(HOJE - 59, HOJE - 30, 'entS').length, s.tNum],
      ['Tempo médio até a entrega (dias)', 'G30', 'G29', mc[11].med, mc[10].med, `IFERROR(AVERAGEIFS(${RD},${RE},">="&(cfgHoje-29),${RE},"<="&cfgHoje,${fob}),"")`, `IFERROR(AVERAGEIFS(${RD},${RE},">="&(cfgHoje-59),${RE},"<="&(cfgHoje-30),${fob}),"")`, mediaEnt(jan(HOJE - 29, HOJE, 'entS')), mediaEnt(jan(HOJE - 59, HOJE - 30, 'entS')), s.tDec],
    ];
    fx.forEach(([t, a1, a2, v1, v2, f1, f2, w1, w2, st], i) => {
      const r = 35 + i, var_ = (x, y) => (typeof x === 'number' && typeof y === 'number' && y) ? x / y - 1 : '';
      ind.put(`B${r}`, t, s.tTxB); ind.f(`C${r}`, a1, st, v1); ind.f(`D${r}`, a2, st, v2);
      ind.f(`E${r}`, `IF(OR(C${r}="",D${r}="",D${r}=0),"",C${r}/D${r}-1)`, s.tVar, var_(v1, v2));
      ind.f(`F${r}`, f1, st, w1); ind.f(`G${r}`, f2, st, w2);
      ind.f(`H${r}`, `IF(OR(F${r}="",G${r}="",G${r}=0),"",F${r}/G${r}-1)`, s.tVar, var_(w1, w2));
    });
    ind.cf.push({ ref: 'E35:E36 H35:H36', rules: [{ type: 'cellIs', op: 'greaterThan', f: ['0'], dxf: dx.azulTx }, { type: 'cellIs', op: 'lessThan', f: ['0'], dxf: dx.ambarTx }] });
    ind.cf.push({ ref: 'E37 H37', rules: [{ type: 'cellIs', op: 'greaterThan', f: ['0'], dxf: dx.ambarTx }, { type: 'cellIs', op: 'lessThan', f: ['0'], dxf: dx.azulTx }] });
    // por origem
    secao(ind, 'B39', 'POR ORIGEM (período e responsável do Painel)', 'I');
    cabec(ind, 40, 'B', ['Origem', 'Avarias', 'Entregues', 'Pendentes', '% entregue', 'Tempo médio (dias)', 'Pendência mais antiga'], [0]);
    ind.alt[39] = 30;
    const perO = `${RB},">="&fDe,${RB},"<="&fAte,${RP},fResp`;
    const oc = ORIG.map(g => { const a = noPer.filter(d => d.origem === g); return { n: a.length, e: a.filter(d => !d.pend).length, p: a.filter(d => d.pend).length, m: mediaEnt(a), x: Math.max(0, ...a.filter(d => d.pend && d.dias !== '').map(d => d.dias)) }; });
    ORIG.forEach((g, i) => {
      const r = 41 + i, c = oc[i];
      ind.f(`B${r}`, `IF('CONFIGURAÇÕES'!B${21 + i}="","—",'CONFIGURAÇÕES'!B${21 + i})`, s.tTxB, g);
      ind.f(`C${r}`, `COUNTIFS(${RO},$B${r},${perO})`, s.tNum, c.n);
      ind.f(`D${r}`, `COUNTIFS(${RO},$B${r},${perO},${RS},"Entregue")`, s.tNum, c.e);
      ind.f(`E${r}`, `COUNTIFS(${RO},$B${r},${perO},${RS},"Não entregue")`, s.tNum, c.p);
      ind.f(`F${r}`, `IF(C${r}=0,"",D${r}/C${r})`, s.tPct, c.n ? c.e / c.n : '');
      ind.f(`G${r}`, `IFERROR(AVERAGEIFS(${RD},${RO},$B${r},${perO},${RS},"Entregue"),"")`, s.tDec, c.m);
      ind.f(`H${r}`, `MAXIFS(${RD},${RO},$B${r},${perO},${RS},"Não entregue")`, s.tNum, c.x);
    });
    const rOT = 41 + ORIG.length, oE = 41 + ORIG.length - 1;
    ind.put(`B${rOT}`, 'Total', s.tTot);
    ind.f(`C${rOT}`, `SUM(C41:C${oE})`, s.tTotN, oc.reduce((a, c) => a + c.n, 0)); ind.f(`D${rOT}`, `SUM(D41:D${oE})`, s.tTotN, oc.reduce((a, c) => a + c.e, 0)); ind.f(`E${rOT}`, `SUM(E41:E${oE})`, s.tTotN, oc.reduce((a, c) => a + c.p, 0));
    const toN = oc.reduce((a, c) => a + c.n, 0), toE = oc.reduce((a, c) => a + c.e, 0);
    ind.f(`F${rOT}`, `IF(C${rOT}=0,"",D${rOT}/C${rOT})`, s.tTotP, toN ? toE / toN : ''); ind.f(`G${rOT}`, `IFERROR(AVERAGEIFS(${RD},${perO},${RS},"Entregue"),"")`, s.tTotD, mediaEnt(noPer));
    ind.f(`H${rOT}`, `MAX(H41:H${oE})`, s.tTotN, Math.max(0, ...oc.map(c => c.x)));
    // tempo em aberto
    const rA = rOT + 2;
    secao(ind, `B${rA}`, 'TEMPO EM ABERTO DAS PENDÊNCIAS (origem do Painel)', 'I');
    cabec(ind, rA + 1, 'B', ['Faixa', 'Pendências', '% das pendências'], [0]);
    const lim = [[0, FX[0]], [FX[0] + 1, FX[1]], [FX[1] + 1, FX[2]], [FX[2] + 1, FX[3]], [FX[3] + 1, 1e9]];
    const limF = [['0', 'fx1'], ['fx1+1', 'fx2'], ['fx2+1', 'fx3'], ['fx3+1', 'fx4'], ['fx4+1', '1E+9']];
    const pendAll = D.filter(d => d.pend && d.dias !== '');
    const aging = lim.map(([a, b]) => pendAll.filter(d => d.dias >= a && d.dias <= b).length);
    lim.forEach((_, i) => {
      const r = rA + 2 + i;
      ind.f(`B${r}`, `'CONFIGURAÇÕES'!D${13 + i}`, s.tTxB, rotFxV[i]);
      ind.f(`C${r}`, `COUNTIFS(${RS},"Não entregue",${RO},fOrig,${RD},">="&(${limF[i][0]}),${RD},"<="&(${limF[i][1]}))`, s.tNum, aging[i]);
      ind.f(`D${r}`, `IF(SUM($C$${rA + 2}:$C$${rA + 6})=0,"",C${r}/SUM($C$${rA + 2}:$C$${rA + 6}))`, s.tPct, pendAll.length ? aging[i] / pendAll.length : '');
    });
    // produtividade
    const rP = rA + 9;
    secao(ind, `B${rP}`, 'PRODUTIVIDADE POR RESPONSÁVEL (entregas no período do Painel, pela data de entrega)', 'I');
    cabec(ind, rP + 1, 'B', ['Responsável', 'Entregas', 'Tempo médio (dias)', 'Participação'], [0]);
    const perE = `${RE},">="&fDe,${RE},"<="&fAte,${RO},fOrig`;
    const respRows = [...RESP, '(sem responsável)'], fimR = rP + 1 + respRows.length;
    const entPer = D.filter(d => !d.pend && entre(d.entS, fil.de, fil.ate));
    const pc = respRows.map(p => { const a = entPer.filter(d => p === '(sem responsável)' ? !d.por : d.por === p); return { n: a.length, m: media(a.map(d => d.dias).filter(v => v !== '')) }; });
    respRows.forEach((p, i) => {
      const r = rP + 2 + i, sem = p === '(sem responsável)', crit = sem ? `${RP},""` : `${RP},$B${r}`;
      if (sem) ind.put(`B${r}`, p, s.tTx); else ind.f(`B${r}`, `IF('CONFIGURAÇÕES'!D${21 + i}="","—",'CONFIGURAÇÕES'!D${21 + i})`, s.tTxB, p);
      ind.f(`C${r}`, `COUNTIFS(${crit},${perE})`, s.tNumB, pc[i].n);
      ind.f(`D${r}`, `IFERROR(AVERAGEIFS(${RD},${crit},${perE}),"")`, s.tDec, pc[i].m);
      ind.f(`E${r}`, `IF(SUM($C$${rP + 2}:$C$${fimR})=0,0,C${r}/SUM($C$${rP + 2}:$C$${fimR}))`, s.tPctE, entPer.length ? pc[i].n / entPer.length : 0);
    });
    ind.cf.push({ ref: `E${rP + 2}:E${fimR}`, rules: [{ type: 'dataBar', cor: '93B4DC' }] });
    ind.put(`B${fimR + 1}`, 'Total', s.tTot); ind.f(`C${fimR + 1}`, `SUM(C${rP + 2}:C${fimR})`, s.tTotN, entPer.length);
    ind.f(`D${fimR + 1}`, `IFERROR(AVERAGEIFS(${RD},${perE}),"")`, s.tTotD, media(entPer.map(d => d.dias).filter(v => v !== ''))); ind.f(`E${fimR + 1}`, `IF(C${fimR + 1}=0,0,1)`, s.tTotP, entPer.length ? 1 : 0);

    // ================================================================ PAINEL
    const pn = Folha('PAINEL', E, sst); pn.tab = C.azul; pn.protegida = true; pn.zoom = 90; pn.altPag = 1;
    pn.cols = [2, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 2];
    moldura(pn, 12, 'Painel executivo', 'Indicadores calculados a partir da aba LANÇAMENTOS. Use os filtros em creme; deixe em branco para ver tudo.');
    pn.merge('K4', 'M4', { f: 'cfgHoje', v: HOJE }, s.dataR);
    pn.alt[6] = 14; pn.alt[7] = 24; pn.alt[8] = 10; pn.alt[9] = 22; pn.alt[10] = 40; pn.alt[11] = 22;
    pn.put('C7', 'PERÍODO DO RELATÓRIO — DE', s.rot); pn.put('F7', 'ATÉ', s.rot); pn.put('I7', 'ORIGEM', s.rot); pn.put('L7', 'RESPONSÁVEL PELA ENTREGA', s.rot);
    pn.put('B8', 'FILTROS', E.xf({ font: { sz: 9, b: 1, cor: C.azul } }));
    pn.merge('C8', 'D8', '', s.inData); pn.merge('F8', 'G8', '', s.inData); pn.merge('I8', 'J8', 'Todas', s.inTxt); pn.merge('L8', 'M8', 'Todas', s.inTxt);
    pn.dv.push({ type: 'custom', f1: 'OR(C8="",ISNUMBER(C8))', ref: 'C8 F8', err: 'Digite uma data válida (dd/mm/aaaa) ou deixe em branco.', msgT: 'Período', msg: 'Filtra pela data do relatório. Em branco = todo o histórico.' });
    pn.dv.push({ type: 'list', f1: 'lstOrigemF', ref: 'I8', err: 'Escolha uma opção da lista.', msgT: 'Origem', msg: 'Escolha "Todas" ou uma origem.' });
    pn.dv.push({ type: 'list', f1: 'lstRespF', ref: 'L8', err: 'Escolha uma opção da lista.', msgT: 'Responsável', msg: 'Mostra só as avarias entregues por esta pessoa.' });
    const I_ = r => `'INDICADORES'!$C$${r}`, I0 = r => `'INDICADORES'!$D$${r}`, IV = r => `'INDICADORES'!$E$${r}`;
    const oldest = D.filter(d => d.pend && d.dias !== '').sort((a, b) => b.dias - a.dias);
    const condPend = `(${RS}="Não entregue")*(${RB}>=fDe)*(${RB}<=fAte)*((fOrig="*")+(${RO}=fOrig)>0)*((fResp="<>#todos#")+(${RP}=fResp)>0)`;
    const kpis = [
      ['AVARIAS REGISTRADAS', I_(9), nTot, s.kVal, `IF(NOT(fTem),"todo o histórico · defina datas para comparar",IF(${I0(9)}=0,"período anterior sem registros","vs período anterior: "&TEXT(${IV(9)},"+0%;-0%;0%")))`, 'todo o histórico · defina datas para comparar'],
      ['ENTREGUES', I_(10), nEnt, s.kVal, `IF(${I_(9)}=0,"—",TEXT(${I_(10)}/${I_(9)},"0%")&" das avarias do período")`, nTot ? Math.round(nEnt / nTot * 100) + '% das avarias do período' : '—'],
      ['PENDENTES', I_(11), nPen, s.kVal, `"aguardando entrega do relatório"`, 'aguardando entrega do relatório'],
      ['TAXA DE ENTREGA', I_(12), pct(nEnt, nTot), s.kValP, `IF(fTem,"vs período anterior: "&TEXT(${IV(12)},"+0;-0;0")&" p.p.","entregues ÷ registradas")`, 'entregues ÷ registradas'],
      ['TEMPO MÉDIO ATÉ A ENTREGA', I_(13), mEnt === '' ? 0 : mEnt, s.kValD, `"dias, do relatório até a entrega"`, 'dias, do relatório até a entrega'],
      ['PENDÊNCIA MAIS ANTIGA', I_(14), maxPend, s.kVal, `LET(_xlpm.c,${condPend},IF(SUM(_xlpm.c)=0,"sem pendências","dias em aberto · Ref. "&INDEX(SORTBY(FILTER(${RR},_xlpm.c),FILTER(${RD},_xlpm.c),-1),1)))`, oldest.length ? 'dias em aberto · Ref. ' + oldest[0].ref : 'sem pendências'],
    ];
    kpis.forEach(([t, f, v, st, nf, nv], i) => {
      const a = colL(1 + i * 2), b = colL(2 + i * 2);
      pn.put(`${a}10`, t, s.kRot); pn.put(`${b}10`, '', s.kRotR);
      pn.f(`${a}11`, f, st, v); pn.put(`${b}11`, '', s.kValR);
      pn.put(`${b}12`, '', s.kNotaR);
      if (DYN.test(nf)) { pn.dyn(`${a}12`, nf, nv, s.kNota); pn.merges.push(`${a}10:${b}10`, `${a}11:${b}11`); }
      else { pn.f(`${a}12`, nf, s.kNota, nv); pn.merges.push(`${a}10:${b}10`, `${a}11:${b}11`, `${a}12:${b}12`); }
    });
    pn.cf.push({ ref: 'F11', rules: [{ type: 'cellIs', op: 'greaterThan', f: ['0'], dxf: dx.ambarTx }] });
    secao(pn, 'B14', 'EVOLUÇÃO MENSAL', 'H'); secao(pn, 'J14', 'PONTOS DE ATENÇÃO', 'M');
    cabec(pn, 15, 'J', ['Situação', '', '', 'Qtd.'], [0]); pn.merges.push('J15:L15');
    const cnt = t => D.filter(d => d.alerta === t).length;
    const aten = [
      ['Refs. duplicadas', `COUNTIF(${RA},"Ref. duplicada")`, cnt('Ref. duplicada')],
      ['Campos obrigatórios vazios', `COUNTIF(${RA},"Campo obrigatório vazio")`, cnt('Campo obrigatório vazio')],
      ['Datas inconsistentes', `COUNTIF(${RA},"Data do relatório no futuro")+COUNTIF(${RA},"Entrega antes do relatório")+COUNTIF(${RA},"Data de entrega no futuro")`, cnt('Data do relatório no futuro') + cnt('Entrega antes do relatório') + cnt('Data de entrega no futuro')],
      ['Entregues sem responsável', `COUNTIF(${RA},"Entregue sem responsável")`, cnt('Entregue sem responsável')],
      [{ f: `"Pendentes há mais de "&fx4&" dias"`, v: `Pendentes há mais de ${FX[3]} dias` }, `COUNTIFS(${RS},"Não entregue",${RD},">"&fx4)`, pendAll.filter(d => d.dias > FX[3]).length],
    ];
    aten.forEach(([t, f, v], i) => { const r = 16 + i; pn.merge(`J${r}`, `L${r}`, t, s.tTx); pn.f(`M${r}`, f, s.tNumB, v); });
    pn.cf.push({ ref: 'M16:M18', rules: [{ type: 'cellIs', op: 'greaterThan', f: ['0'], dxf: dx.ruim }] });
    pn.cf.push({ ref: 'M19:M20', rules: [{ type: 'cellIs', op: 'greaterThan', f: ['0'], dxf: dx.pend }] });
    pn.merge('J21', 'M21', 'Ver os registros na Central  →', s.link); pn.link('J21', `'CENTRAL'!A1`, 'Abrir a Central de operações');
    secao(pn, 'J23', 'ATIVIDADE RECENTE (sem filtros)', 'M');
    cabec(pn, 24, 'J', ['Período', '', 'Registradas', 'Entregues'], [0]); pn.merges.push('J24:K24');
    const ativ = [['Hoje', 0], ['Últimos 7 dias', 6], ['Últimos 30 dias', 29]];
    ativ.forEach(([t, k], i) => {
      const r = 25 + i; pn.merge(`J${r}`, `K${r}`, t, s.tTxB);
      pn.f(`L${r}`, `COUNTIFS(${RB},">="&(cfgHoje-${k}),${RB},"<="&cfgHoje)`, s.tNum, jan(HOJE - k, HOJE, 'relS').length);
      pn.f(`M${r}`, `COUNTIFS(${RE},">="&(cfgHoje-${k}),${RE},"<="&cfgHoje)`, s.tNum, jan(HOJE - k, HOJE, 'entS').length);
    });
    secao(pn, 'B33', 'TEMPO EM ABERTO DAS PENDÊNCIAS', 'F'); secao(pn, 'H33', 'POR ORIGEM', 'M');
    secao(pn, 'B51', 'PRODUTIVIDADE POR RESPONSÁVEL', 'F'); secao(pn, 'H51', 'PENDÊNCIAS MAIS ANTIGAS', 'M');
    cabec(pn, 52, 'H', ['Ref.', 'Origem', 'Dias', 'Descrição do produto', '', ''], [3]); pn.merges.push('K52:M52');
    for (let r = 53; r <= 62; r++) { pn.put(`H${r}`, '', s.dNum); pn.put(`I${r}`, '', s.dCen); pn.put(`J${r}`, '', s.dDias); pn.put(`K${r}`, '', s.dTx); }
    const top = D.filter(d => d.pend && d.dias !== '').sort((a, b) => b.dias - a.dias).slice(0, 10).map(d => [d.ref, d.origem, d.dias, d.produto]);
    pn.dyn('H53', `LET(_xlpm.c,(${RS}="Não entregue")*((fOrig="*")+(${RO}=fOrig)>0),_xlpm.n,SUM(_xlpm.c),IF(_xlpm.n=0,"Nenhuma pendência.",INDEX(SORTBY(FILTER(${pick([IDX('Ref.'), IDX('Origem'), IDX('Dias'), IDX('Descrição do produto')])},_xlpm.c),FILTER(${RD},_xlpm.c),-1),SEQUENCE(MIN(10,_xlpm.n)),SEQUENCE(1,4))))`, top.length ? top : [['Nenhuma pendência.']], null);
    pn.cf.push({ ref: 'H53:K62', rules: [{ type: 'expression', f: ['$H53<>""'], dxf: dx.linhaDado }] });
    pn.put('H64', 'Lista completa na aba CENTRAL.', s.nota); pn.link('H64', `'CENTRAL'!A1`, 'Abrir a Central');
    pn.drawing = true;

    // ================================================================ CENTRAL
    const ce = Folha('CENTRAL', E, sst); ce.tab = C.azul; ce.protegida = true;
    ce.cols = [2, 10, 12, 14, 12, 38, 13, 12, 18, 9, 3, 10, 28, 34];
    moldura(ce, 13, 'Central de operações', 'Consulta rápida, pendências e alertas de qualidade — tudo atualizado automaticamente.');
    secao(ce, 'B7', 'CONSULTA RÁPIDA', 'J');
    ce.put('B8', 'Buscar', E.xf({ font: { sz: 9, b: 1, cor: C.azul } }));
    ce.merge('C8', 'F8', '', s.inTxt);
    ce.merge('G8', 'J8', 'Ref., código, descrição ou responsável · até 15 resultados', s.nota);
    ce.dv.push({ type: 'textLength', op: 'lessThanOrEqual', f1: '60', ref: 'C8', msgT: 'Buscar', msg: 'Digite parte da Ref., do código, da descrição ou o nome de quem entregou.', err: 'Use até 60 caracteres.' });
    cabec(ce, 10, 'B', ['Ref.', 'Relatório', 'Origem', 'Código', 'Descrição do produto', 'Situação', 'Entrega', 'Entregue por', 'Dias'], [4, 7]);
    const estCols = [s.dNum, s.dData, s.dCen, s.dCen, s.dTx, s.dCen, s.dData, s.dTx, s.dDias];
    for (let r = 11; r <= 25; r++) estCols.forEach((st, j) => ce.put(A(1 + j, r - 1), '', st));
    ce.dyn('B11', `LET(_xlpm.t,TRIM($C$8),_xlpm.c,ISNUMBER(SEARCH(_xlpm.t,${RR}&" "&${col('Código')}&" "&${col('Descrição do produto')}&" "&${RP})),_xlpm.n,SUM(--_xlpm.c),IF(_xlpm.t="","Digite um termo no campo Buscar.",IF(_xlpm.n=0,"Nenhum registro encontrado.",INDEX(SORTBY(FILTER(${pick([1, 2, 3, 4, 5, 9, 6, 7, 10])},_xlpm.c),FILTER(${RR},_xlpm.c),-1),SEQUENCE(MIN(15,_xlpm.n)),SEQUENCE(1,9)))))`, 'Digite um termo no campo Buscar.', null);
    ce.cf.push({ ref: 'B11:J25', rules: [{ type: 'expression', f: ['AND(ISNUMBER($B11),$G11="Entregue",COLUMN()=7)'], dxf: dx.ok }, { type: 'expression', f: ['AND(ISNUMBER($B11),$G11="Não entregue",COLUMN()=7)'], dxf: dx.pend }, { type: 'expression', f: ['ISNUMBER($B11)'], dxf: dx.linhaDado }] });
    secao(ce, 'B28', 'PENDÊNCIAS — MAIS ANTIGAS PRIMEIRO', 'G'); secao(ce, 'L28', 'ALERTAS DE QUALIDADE', 'N');
    cabec(ce, 29, 'B', ['Ref.', 'Relatório', 'Origem', 'Código', 'Descrição do produto', 'Dias em aberto'], [4]);
    cabec(ce, 29, 'L', ['Ref.', 'Alerta', 'Descrição do produto'], [1, 2]);
    const RES = Math.max(300, D.length + 50);
    const estP = [s.dNum, s.dData, s.dCen, s.dCen, s.dTx, s.dDias], estA = [s.dNum, s.dTx, s.dTx];
    for (let r = 30; r < 30 + RES; r++) { estP.forEach((st, j) => ce.put(A(1 + j, r - 1), '', st)); estA.forEach((st, j) => ce.put(A(11 + j, r - 1), '', st)); }
    const pendL = D.filter(d => d.pend && d.dias !== '').sort((a, b) => b.dias - a.dias).map(d => [d.ref, d.relS, d.origem, d.sku, d.produto, d.dias]);
    ce.dyn('B30', `LET(_xlpm.c,(${RS}="Não entregue"),IF(SUM(--_xlpm.c)=0,"Nenhuma pendência no momento.",SORTBY(FILTER(${pick([1, 2, 3, 4, 5, 10])},_xlpm.c),FILTER(${RD},_xlpm.c),-1)))`, pendL.length ? pendL : [['Nenhuma pendência no momento.']], null);
    const alL = D.filter(d => d.alerta).map(d => [d.ref === '' ? '' : d.ref, d.alerta, d.produto]);
    ce.dyn('L30', `LET(_xlpm.c,(${RA}<>""),IF(SUM(--_xlpm.c)=0,"Nenhum alerta. Base consistente.",FILTER(${pick([1, 12, 5])},_xlpm.c)))`, alL.length ? alL : [['Nenhum alerta. Base consistente.']], null);
    ce.cf.push({ ref: `B30:G${29 + RES}`, rules: [{ type: 'expression', f: ['ISNUMBER($B30)'], dxf: dx.linhaDado }] });
    ce.cf.push({ ref: `L30:N${29 + RES}`, rules: [{ type: 'expression', f: ['AND($M30="Entregue sem responsável",COLUMN()=13)'], dxf: dx.pend, stop: 1 }, { type: 'expression', f: ['AND($M30<>"",$M30<>"Nenhum alerta. Base consistente.",COLUMN()=13)'], dxf: dx.ruim }, { type: 'expression', f: ['$M30<>""'], dxf: dx.linhaDado }] });

    // ================================================================ RELATÓRIO
    const rl = Folha('RELATÓRIO', E, sst); rl.tab = '475569'; rl.protegida = true; rl.paisagem = false;
    rl.cols = [2, 10, 12, 14, 12, 34, 13, 12, 10];
    moldura(rl, 8, 'Relatório de Avarias', 'Relatório consolidado do período, pronto para imprimir ou salvar em PDF (Arquivo › Exportar).');
    rl.merge('G5', 'I5', { f: 'cfgHoje', v: HOJE }, s.emitido);
    rl.put('B7', 'PERÍODO', E.xf({ font: { sz: 9, b: 1, cor: C.azul } })); rl.put('C7', repDe, s.inData); rl.put('D7', 'até', s.rotC); rl.put('E7', repAte, s.inData);
    rl.merge('F7', 'I7', 'Em branco = todo o histórico. Altere as datas para outro período.', s.nota);
    rl.dv.push({ type: 'custom', f1: 'OR(C7="",ISNUMBER(C7))', ref: 'C7 E7', err: 'Digite uma data válida (dd/mm/aaaa) ou deixe em branco.', msgT: 'Período do relatório', msg: 'Filtra pela data do relatório de avaria.' });
    const rp = D.filter(d => entre(d.relS, repDe, repAte)), rp0 = D.filter(d => entre(d.relS, repDe - repLen, repDe - 1));
    const rk = a => ({ n: a.length, e: a.filter(d => !d.pend).length, p: a.filter(d => d.pend).length, m: mediaEnt(a), x: Math.max(0, ...a.filter(d => d.pend && d.dias !== '').map(d => d.dias)) });
    const k1 = rk(rp), k0 = rk(rp0);
    const rb = `${RB},">="&rDe,${RB},"<="&rAte`, rb0 = `${RB},">="&rDe0,${RB},"<="&rAte0`;
    secao(rl, 'B9', 'RESUMO EXECUTIVO', 'I');
    rl.alt[9] = 22; rl.alt[10] = 22; rl.alt[11] = 22;
    const perTxt = `IF(rTem,${dataBR('rDe')}&" a "&${dataBR('rAte')},"todo o histórico")`;
    const resumoF = `"Período analisado: "&${perTxt}&". Foram registradas "&B16&" avarias; "&D16&" já foram entregues ("&TEXT(IF(B16=0,0,D16/B16),"0%")&") e "&F16&" seguem pendentes."&IF(F16>0," A pendência mais antiga do período está em aberto há "&MAXIFS(${RD},${rb},${RS},"Não entregue")&" dias.","")&IF(AND(rTem,E22>0)," Em relação ao período anterior, os registros "&IF(D22>=E22,"aumentaram ","diminuíram ")&TEXT(ABS(D22/E22-1),"0%")&".","")`;
    const resumoV = `Período analisado: ${br(repDe)} a ${br(repAte)}. Foram registradas ${k1.n} avarias; ${k1.e} já foram entregues (${Math.round(pct(k1.e, k1.n) * 100)}%) e ${k1.p} seguem pendentes.` + (k1.p > 0 ? ` A pendência mais antiga do período está em aberto há ${k1.x} dias.` : '') + (k0.n > 0 ? ` Em relação ao período anterior, os registros ${k1.n >= k0.n ? 'aumentaram' : 'diminuíram'} ${Math.round(Math.abs(k1.n / k0.n - 1) * 100)}%.` : '');
    rl.merge('B10', 'I12', { f: resumoF, v: resumoV }, s.resumo);
    secao(rl, 'B14', 'INDICADORES DO PERÍODO', 'I');
    rl.alt[14] = 20; rl.alt[15] = 34; rl.alt[16] = 18;
    const rcards = [['REGISTRADAS', `COUNTIFS(${rb})`, k1.n, s.kVal], ['ENTREGUES', `COUNTIFS(${rb},${RS},"Entregue")`, k1.e, s.kVal], ['PENDENTES', `COUNTIFS(${rb},${RS},"Não entregue")`, k1.p, s.kVal], ['TEMPO MÉDIO (dias)', `IFERROR(AVERAGEIFS(${RD},${rb},${RS},"Entregue"),0)`, k1.m === '' ? 0 : k1.m, s.kValD]];
    rcards.forEach(([t, f, v, st], i) => {
      const a = colL(1 + i * 2), b = colL(2 + i * 2);
      rl.put(`${a}15`, t, s.kRot); rl.put(`${b}15`, '', s.kRotR); rl.f(`${a}16`, f, st, v); rl.put(`${b}16`, '', s.kValR); rl.put(`${a}17`, '', s.kNota); rl.put(`${b}17`, '', s.kNotaR);
      rl.merges.push(`${a}15:${b}15`, `${a}16:${b}16`, `${a}17:${b}17`);
    });
    rl.f('B17', `IF(rTem,"anterior: "&E22,"")`, s.kNota, `anterior: ${k0.n}`); rl.f('D17', `IF(rTem,"anterior: "&E23,"")`, s.kNota, `anterior: ${k0.e}`);
    rl.f('F17', `IF(rTem,"anterior: "&E24,"")`, s.kNota, `anterior: ${k0.p}`); rl.f('H17', `IF(rTem,"anterior: "&TEXT(E26,"0.0"),"")`, s.kNota, `anterior: ${(k0.m === '' ? 0 : k0.m).toFixed(1)}`);
    // D16 = registradas, F16 = entregues, H16 = pendentes (cards em B,D,F,H) — o resumo usa estas células
    secao(rl, 'B20', 'COMPARATIVO COM O PERÍODO ANTERIOR', 'I');
    rl.put('B21', 'Indicador', s.cabE); rl.put('C21', '', s.cabE); rl.merges.push('B21:C21'); cabec(rl, 21, 'D', ['Período', 'Anterior', 'Variação']);
    rl.merge('G21', 'I21', 'Anterior = mesmo nº de dias imediatamente antes', s.nota);
    const rcomp = [['Registradas', `COUNTIFS(${rb})`, `COUNTIFS(${rb0})`, k1.n, k0.n, s.tNum, 'var'], ['Entregues', `COUNTIFS(${rb},${RS},"Entregue")`, `COUNTIFS(${rb0},${RS},"Entregue")`, k1.e, k0.e, s.tNum, 'var'], ['Pendentes', `COUNTIFS(${rb},${RS},"Não entregue")`, `COUNTIFS(${rb0},${RS},"Não entregue")`, k1.p, k0.p, s.tNum, 'var'], ['Taxa de entrega', `IF(D22=0,0,D23/D22)`, `IF(E22=0,0,E23/E22)`, pct(k1.e, k1.n), pct(k0.e, k0.n), s.tPct, 'pp'], ['Tempo médio (dias)', `IFERROR(AVERAGEIFS(${RD},${rb},${RS},"Entregue"),0)`, `IFERROR(AVERAGEIFS(${RD},${rb0},${RS},"Entregue"),0)`, k1.m === '' ? 0 : k1.m, k0.m === '' ? 0 : k0.m, s.tDec, 'd']];
    rcomp.forEach(([t, f1, f0, v1, v0, st, tipo], i) => {
      const r = 22 + i; rl.merge(`B${r}`, `C${r}`, t, s.tTxB); rl.f(`D${r}`, f1, st, v1); rl.f(`E${r}`, f0, st, v0);
      if (tipo === 'var') rl.f(`F${r}`, `IF(OR(NOT(rTem),E${r}=0),"",D${r}/E${r}-1)`, s.tVar, v0 ? v1 / v0 - 1 : '');
      else if (tipo === 'pp') rl.f(`F${r}`, `IF(NOT(rTem),"",(D${r}-E${r})*100)`, s.tPP, (v1 - v0) * 100);
      else rl.f(`F${r}`, `IF(OR(NOT(rTem),E${r}=0),"",D${r}-E${r})`, s.tVarD, v0 ? v1 - v0 : '');
    });
    secao(rl, 'B28', 'POR ORIGEM', 'I');
    rl.put('B29', 'Origem', s.cabE); rl.put('C29', '', s.cabE); rl.merges.push('B29:C29'); cabec(rl, 29, 'D', ['Avarias', 'Entregues', 'Pendentes', '% entregue', 'Tempo médio']);
    const roc = ORIG.map(g => rk(rp.filter(d => d.origem === g)));
    ORIG.forEach((g, i) => {
      const r = 30 + i, c = roc[i];
      rl.merge(`B${r}`, `C${r}`, { f: `IF('CONFIGURAÇÕES'!B${21 + i}="","—",'CONFIGURAÇÕES'!B${21 + i})`, v: g }, s.tTxB);
      rl.f(`D${r}`, `COUNTIFS(${RO},$B${r},${rb})`, s.tNum, c.n); rl.f(`E${r}`, `COUNTIFS(${RO},$B${r},${rb},${RS},"Entregue")`, s.tNum, c.e); rl.f(`F${r}`, `COUNTIFS(${RO},$B${r},${rb},${RS},"Não entregue")`, s.tNum, c.p);
      rl.f(`G${r}`, `IF(D${r}=0,"",E${r}/D${r})`, s.tPct, c.n ? c.e / c.n : ''); rl.f(`H${r}`, `IFERROR(AVERAGEIFS(${RD},${RO},$B${r},${rb},${RS},"Entregue"),"")`, s.tDec, c.m);
    });
    const rRT = 30 + ORIG.length;
    rl.merge(`B${rRT}`, `C${rRT}`, 'Total', s.tTot); rl.f(`D${rRT}`, `SUM(D30:D${rRT - 1})`, s.tTotN, k1.n); rl.f(`E${rRT}`, `SUM(E30:E${rRT - 1})`, s.tTotN, k1.e); rl.f(`F${rRT}`, `SUM(F30:F${rRT - 1})`, s.tTotN, k1.p);
    rl.f(`G${rRT}`, `IF(D${rRT}=0,"",E${rRT}/D${rRT})`, s.tTotP, k1.n ? k1.e / k1.n : ''); rl.f(`H${rRT}`, `IFERROR(AVERAGEIFS(${RD},${rb},${RS},"Entregue"),"")`, s.tTotD, k1.m);
    const rG = rRT + 2; secao(rl, `B${rG}`, 'SITUAÇÃO POR ORIGEM NO PERÍODO', 'I');
    const rDt = rG + 17; secao(rl, `B${rDt}`, 'DETALHAMENTO DAS AVARIAS DO PERÍODO', 'I');
    cabec(rl, rDt + 1, 'B', ['Ref.', 'Relatório', 'Origem', 'Código', 'Descrição do produto', 'Situação', 'Entrega', 'Dias'], [4]);
    const RESr = Math.max(200, D.length + 30), estR = [s.dNum, s.dData, s.dCen, s.dCen, s.dTx, s.dCen, s.dData, s.dDias];
    for (let r = rDt + 2; r < rDt + 2 + RESr; r++) estR.forEach((st, j) => rl.put(A(1 + j, r - 1), '', st));
    const detL = [...rp].sort((a, b) => a.ref - b.ref).map(d => [d.ref, d.relS, d.origem, d.sku, d.produto, d.sit, d.entS ?? '', d.dias]);
    rl.dyn(`B${rDt + 2}`, `LET(_xlpm.c,(${RB}>=rDe)*(${RB}<=rAte)*(${RR}<>""),IF(SUM(_xlpm.c)=0,"Nenhuma avaria no período.",SORTBY(FILTER(${pick([1, 2, 3, 4, 5, 9, 6, 10])},_xlpm.c),FILTER(${RR},_xlpm.c),1)))`, detL.length ? detL : [['Nenhuma avaria no período.']], null);
    rl.cf.push({ ref: `B${rDt + 2}:I${rDt + 1 + RESr}`, rules: [{ type: 'expression', f: [`AND($G${rDt + 2}="Entregue",COLUMN()=7)`], dxf: dx.ok }, { type: 'expression', f: [`AND($G${rDt + 2}="Não entregue",COLUMN()=7)`], dxf: dx.pend }, { type: 'expression', f: [`ISNUMBER($B${rDt + 2})`], dxf: dx.linhaDado }] });
    rl.drawing = true;

    // ================================================================ AJUDA
    const aj = Folha('AJUDA', E, sst); aj.tab = '94A3B8'; aj.protegida = true; aj.paisagem = false;
    aj.cols = [2, 20, 12, 12, 12, 12, 12, 12, 12, 12];
    moldura(aj, 8, 'Como usar o sistema', 'Guia rápido de operação. Clique no nome de uma aba para ir até ela.');
    const blocos = [
      ['MÓDULOS', [
        ['PAINEL', 'Visão executiva: indicadores, gráficos, pontos de atenção e as pendências mais antigas. Filtros de período, origem e responsável no topo.', 'PAINEL'],
        ['CENTRAL', 'Operação do dia a dia: busca rápida, lista de pendências (mais antigas primeiro) e alertas de qualidade dos dados.', 'CENTRAL'],
        ['LANÇAMENTOS', 'Banco de dados e cadastro. Uma linha por avaria. É aqui que se lança uma avaria nova e a data de entrega.', 'LANÇAMENTOS'],
        ['HISTÓRICO', 'Quem registrou ou trocou a data de entrega e quando. Linhas do sistema vêm prontas; correções feitas na planilha devem ser anotadas aqui.', 'HISTÓRICO'],
        ['RELATÓRIO', 'Relatório do período com resumo executivo, indicadores, comparativo, gráfico e detalhamento. Pronto para PDF.', 'RELATÓRIO'],
        ['INDICADORES', 'Tabelas que alimentam o Painel: comparativos, evolução mensal, origens, tempo em aberto e produtividade.', 'INDICADORES'],
        ['CONFIGURAÇÕES', 'Nome do CD, faixas de tempo em aberto e as listas suspensas (origens e responsáveis).', 'CONFIGURAÇÕES'],
      ]],
      ['COMO LANÇAR UMA AVARIA', [
        ['1. Abrir', 'Vá em LANÇAMENTOS e clique na primeira linha vazia logo abaixo da tabela.'],
        ['2. Preencher', 'Ref., Data do relatório, Origem (lista), Código e Descrição. Esses campos são obrigatórios.'],
        ['3. Conferir', 'Situação, Dias, Mês e Alerta são calculados sozinhos. Se aparecer um alerta, corrija o campo indicado.'],
        ['4. Entregar', 'Quando o relatório for entregue, preencha Data de entrega e Entregue por. A situação muda para Entregue.'],
      ]],
      ['CORES E CAMPOS', [
        ['Creme', 'Campo editável (filtros, parâmetros, listas e a busca).'],
        ['Cinza', 'Campo automático, calculado por fórmula e protegido.'],
        ['Azul-escuro', 'Cabeçalho de tabela. Nas abas protegidas não pode ser alterado.'],
        ['Vermelho', 'Erro nos dados: Ref. duplicada, campo obrigatório vazio ou datas inconsistentes.'],
        ['Âmbar', 'Atenção: pendente, ou entregue sem responsável informado.'],
      ]],
      ['REGRAS DE VALIDAÇÃO', [
        ['Ref.', 'Número inteiro maior que zero, sem repetir uma Ref. já lançada.'],
        ['Datas', 'Data do relatório entre 01/01/2020 e hoje. Data de entrega igual ou posterior ao relatório e não futura.'],
        ['Listas', 'Origem só aceita itens da lista. Entregue por avisa quando o nome não está na lista.'],
        ['Sem prazo fixo', 'O sistema não considera nenhuma avaria “atrasada”. As faixas de tempo só organizam os relatórios.'],
      ]],
      ['ATUALIZAÇÃO E SEGURANÇA', [
        ['Automático', 'Todos os indicadores, listas e gráficos se atualizam ao lançar ou alterar uma linha.'],
        ['Pelo sistema', 'No Controle de Avaria, aba Planilha › Baixar Excel gera este arquivo com todas as avarias do INVENTRA.'],
        ['Proteção', 'Painel, Central, Relatório, Indicadores, Configurações e Ajuda estão protegidas sem senha. Para alterar a estrutura: Revisão › Desproteger Planilha.'],
        ['Requisito', 'Microsoft Excel 365 ou 2021+ (usa FILTRO, ORDENARPOR, ÚNICO, PROCX e SEQUÊNCIA).'],
      ]],
    ];
    let r = 7;
    blocos.forEach(([tt, itens]) => {
      secao(aj, `B${r}`, tt, 'I'); r++;
      itens.forEach(([a, b, alvo]) => {
        const leg = { Creme: s.chipIn, Cinza: s.chipAu, 'Azul-escuro': s.chipPr, Vermelho: E.xf({ font: { sz: 9, b: 1, cor: C.verm }, fill: C.vermBg, border: bd(), al: { h: 'center' } }), 'Âmbar': E.xf({ font: { sz: 9, b: 1, cor: C.ambar }, fill: C.ambarBg, border: bd(), al: { h: 'center' } }) }[a];
        aj.put(`B${r}`, a, leg || (alvo ? E.xf({ font: { sz: 10, b: 1, cor: C.azul }, al: { v: 'top' } }) : s.txtB));
        if (alvo) aj.link(`B${r}`, `'${alvo}'!A1`, `Ir para ${a}`);
        aj.merge(`C${r}`, `I${r}`, b, s.txt); aj.alt[r - 1] = 30; r++;
      });
      r++;
    });
    aj.put(`B${r}`, 'Controle de Avaria · criado por Lucas Carvalho', s.nota);

    // ================================================================ GRÁFICOS
    const ORIGr = ORIG.length;
    const g1 = grafBarras({ titulo: 'Avarias registradas × entregues por mês', cat: { ref: `'INDICADORES'!$B$19:$B$30`, vals: meses.map(m => m.rot) }, series: [
      { nome: 'Registradas', nomeRef: `'INDICADORES'!$C$18`, ref: `'INDICADORES'!$C$19:$C$30`, vals: mc.map(c => c.reg), cor: C.azul },
      { nome: 'Entregues no mês', nomeRef: `'INDICADORES'!$D$18`, ref: `'INDICADORES'!$D$19:$D$30`, vals: mc.map(c => c.ent), cor: '93B4DC' }] });
    const g2 = grafBarras({ titulo: 'Pendências por tempo em aberto', dir: 'bar', cat: { ref: `'INDICADORES'!$B$${rA + 2}:$B$${rA + 6}`, vals: rotFxV }, series: [
      { nome: 'Pendências', nomeRef: `'INDICADORES'!$C$${rA + 1}`, ref: `'INDICADORES'!$C$${rA + 2}:$C$${rA + 6}`, vals: aging, cor: 'D9A441' }] });
    const g3 = grafBarras({ titulo: 'Entregues e pendentes por origem', dir: 'bar', grupo: 'stacked', cat: { ref: `'INDICADORES'!$B$41:$B$${oE}`, vals: ORIG }, series: [
      { nome: 'Entregues', nomeRef: `'INDICADORES'!$D$40`, ref: `'INDICADORES'!$D$41:$D$${oE}`, vals: oc.map(c => c.e), cor: C.azul },
      { nome: 'Pendentes', nomeRef: `'INDICADORES'!$E$40`, ref: `'INDICADORES'!$E$41:$E$${oE}`, vals: oc.map(c => c.p), cor: 'D9A441' }] });
    const g4 = grafBarras({ titulo: 'Entregas por responsável', dir: 'bar', cat: { ref: `'INDICADORES'!$B$${rP + 2}:$B$${fimR}`, vals: respRows }, series: [
      { nome: 'Entregas', nomeRef: `'INDICADORES'!$C$${rP + 1}`, ref: `'INDICADORES'!$C$${rP + 2}:$C$${fimR}`, vals: pc.map(c => c.n), cor: C.azul }] });
    const g5 = grafBarras({ titulo: 'Situação por origem no período', dir: 'bar', grupo: 'stacked', cat: { ref: `'RELATÓRIO'!$B$30:$B$${rRT - 1}`, vals: ORIG }, series: [
      { nome: 'Entregues', nomeRef: `'RELATÓRIO'!$E$29`, ref: `'RELATÓRIO'!$E$30:$E$${rRT - 1}`, vals: roc.map(c => c.e), cor: C.azul },
      { nome: 'Pendentes', nomeRef: `'RELATÓRIO'!$F$29`, ref: `'RELATÓRIO'!$F$30:$F$${rRT - 1}`, vals: roc.map(c => c.p), cor: 'D9A441' }] });
    const wsDr = anc => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anc.map((a, i) => ancora(i, a)).join('')}</xdr:wsDr>`;
    const draw1 = wsDr([[1, 14, 8, 31], [1, 33, 6, 49], [7, 33, 13, 49], [1, 51, 6, 67]]);
    void ORIGr;
    const draw2 = wsDr([[1, rG, 9, rG + 15]]);

    // ================================================================ PACOTE
    const folhas = [pn, ce, lan, his, rl, ind, cfg, aj];
    const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const sheetXml = folhas.map(f => f.xml());
    const idx = n => folhas.findIndex(f => f.nome === n);
    const definedNames = names.map(([n, f]) => `<definedName name="${n}">${xe(f)}</definedName>`).join('') +
      `<definedName name="_xlnm.Print_Titles" localSheetId="${idx('LANÇAMENTOS')}">'LANÇAMENTOS'!$9:$9</definedName><definedName name="_xlnm.Print_Titles" localSheetId="${idx('RELATÓRIO')}">'RELATÓRIO'!$${rDt + 1}:$${rDt + 1}</definedName><definedName name="_xlnm.Print_Titles" localSheetId="${idx('HISTÓRICO')}">'HISTÓRICO'!$9:$9</definedName>`;
    const wb = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"><workbookPr defaultThemeVersion="164011"/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15000" activeTab="0"/></bookViews><sheets>${folhas.map((f, i) => `<sheet name="${xe(f.nome)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets><definedNames>${definedNames}</definedNames><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`;
    const n = folhas.length;
    const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${folhas.map((f, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${n + 1}" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId${n + 2}" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId${n + 3}" Type="${REL}/sheetMetadata" Target="metadata.xml"/></Relationships>`;
    const metadata = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`;
    const sstXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sstArr.length}" uniqueCount="${sstArr.length}">${sstArr.map(t => `<si><t xml:space="preserve">${xe(t)}</t></si>`).join('')}</sst>`;
    const charts = [g1, g2, g3, g4, g5];
    const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${folhas.map((f, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>${charts.map((_, i) => `<Override PartName="/xl/charts/chart${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join('')}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
    const relX = (pairs) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${pairs.map(([id, t, tg]) => `<Relationship Id="${id}" Type="${REL}/${t}" Target="${tg}"/>`).join('')}</Relationships>`;
    const sn = nome => idx(nome) + 1;
    const files = [
      ['[Content_Types].xml', ct],
      ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`],
      ['docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Controle de Avaria — Sistema em Excel</dc:title><dc:creator>Lucas Carvalho</dc:creator><cp:lastModifiedBy>${xe(usuario || 'Controle de Avaria')}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${agora.toISOString().slice(0, 19)}Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${agora.toISOString().slice(0, 19)}Z</dcterms:modified></cp:coreProperties>`],
      ['docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application><Company>Controle de Avaria</Company></Properties>`],
      ['xl/workbook.xml', wb], ['xl/_rels/workbook.xml.rels', rels],
      ['xl/styles.xml', E.xml({ inteira: dx.tabInteira, cab: dx.tabCab, faixa: dx.tabFaixa })], ['xl/sharedStrings.xml', sstXml], ['xl/metadata.xml', metadata],
      ...sheetXml.map((x, i) => [`xl/worksheets/sheet${i + 1}.xml`, x]),
      [`xl/worksheets/_rels/sheet${sn('PAINEL')}.xml.rels`, relX([['rIdD', 'drawing', '../drawings/drawing1.xml']])],
      [`xl/worksheets/_rels/sheet${sn('RELATÓRIO')}.xml.rels`, relX([['rIdD', 'drawing', '../drawings/drawing2.xml']])],
      [`xl/worksheets/_rels/sheet${sn('LANÇAMENTOS')}.xml.rels`, relX([['rIdT', 'table', '../tables/table1.xml']])],
      [`xl/worksheets/_rels/sheet${sn('HISTÓRICO')}.xml.rels`, relX([['rIdT', 'table', '../tables/table2.xml']])],
      ['xl/tables/table1.xml', tab1], ['xl/tables/table2.xml', tab2],
      ['xl/drawings/drawing1.xml', draw1], ['xl/drawings/drawing2.xml', draw2],
      ['xl/drawings/_rels/drawing1.xml.rels', relX([1, 2, 3, 4].map(i => [`rIdC${i}`, 'chart', `../charts/chart${i}.xml`]))],
      ['xl/drawings/_rels/drawing2.xml.rels', relX([['rIdC1', 'chart', '../charts/chart5.xml']])],
      ...charts.map((g, i) => [`xl/charts/chart${i + 1}.xml`, g]),
    ];
    return zip(files);
  }
  G.AvariaPlanilha = { gerar, versao: 2 };
})(typeof window !== 'undefined' ? window : globalThis);
