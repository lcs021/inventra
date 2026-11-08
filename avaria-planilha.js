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
    const cel = new Map(); const S = { nome, altPag: 0, cols: [], alt: {}, merges: [], cf: [], dv: [], protegida: false, freeze: null, grade: false, tab: null, drawing: null, tabela: null, zoom: 100, paisagem: true, ocultas: new Set() };
    const put = (a, v, s = 0) => { const [c, r] = parse(a); if (!cel.has(r)) cel.set(r, new Map()); cel.get(r).set(c, { v, s }); };
    S.put = put;
    S.f = (a, f, s = 0, v) => put(a, { f, v }, s);
    S.merge = (a, b, v, s) => { if (v !== undefined) put(a, v, s); else if (s != null) put(a, '', s); if (s != null) { const [c1, r1] = parse(a), [c2, r2] = parse(b); for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (r !== r1 || c !== c1) put(A(c, r), '', s); } S.merges.push(`${a}:${b}`); };
    S.xml = () => {
      const rows = [...cel.keys()].sort((a, b) => a - b);
      let maxC = 0; rows.forEach(r => cel.get(r).forEach((_, c) => { if (c > maxC) maxC = c; }));
      const cellXml = (r, c, { v, s }) => {
        const a = A(c, r), st = s ? ` s="${s}"` : '';
        if (v && typeof v === 'object' && 'f' in v) {
          const val = v.v;
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
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr>${S.tab ? `<tabColor rgb="FF${S.tab}"/>` : ''}<pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:${A(Math.max(maxC, 0), Math.max(rows[rows.length - 1] || 0, 0))}"/><sheetViews><sheetView workbookViewId="0"${S.grade ? '' : ' showGridLines="0"'} zoomScale="${S.zoom}" zoomScaleNormal="${S.zoom}">${pane || '<selection activeCell="A1" sqref="A1"/>'}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="18" customHeight="1"/>${S.cols.length ? `<cols>${S.cols.map((w, i) => w ? `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"${S.ocultas.has(i) ? ' hidden="1"' : ''}/>` : '').join('')}</cols>` : ''}<sheetData>${rows.map(r => { const m = cel.get(r); return `<row r="${r + 1}"${S.alt[r] ? ` ht="${S.alt[r]}" customHeight="1"` : ''}>${[...m.keys()].sort((a, b) => a - b).map(c => cellXml(r, c, m.get(c))).join('')}</row>`; }).join('')}</sheetData>${S.protegida ? '<sheetProtection sheet="1" objects="1" scenarios="1" formatColumns="0" formatRows="0" sort="0" autoFilter="0"/>' : ''}${S.merges.length ? `<mergeCells count="${S.merges.length}">${S.merges.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : ''}${S.cf.map((c, i) => `<conditionalFormatting sqref="${c.ref}">${c.rules.map((ru, j) => ru.type === 'dataBar' ? `<cfRule type="dataBar" priority="${i * 10 + j + 1}"><dataBar><cfvo type="num" val="0"/><cfvo type="num" val="1"/><color rgb="FF${ru.cor}"/></dataBar></cfRule>` : `<cfRule type="${ru.type}" dxfId="${ru.dxf}" priority="${i * 10 + j + 1}"${ru.op ? ` operator="${ru.op}"` : ''}${ru.text ? ` text="${xe(ru.text)}"` : ''}${ru.stop ? ' stopIfTrue="1"' : ''}>${(ru.f || []).map(f => `<formula>${xe(f)}</formula>`).join('')}</cfRule>`).join('')}</conditionalFormatting>`).join('')}${S.dv.length ? `<dataValidations count="${S.dv.length}">${S.dv.map(d => `<dataValidation type="${d.type}"${d.op ? ` operator="${d.op}"` : ''} allowBlank="1" showInputMessage="1" showErrorMessage="1"${d.err ? ` errorStyle="${d.estilo || 'stop'}" errorTitle="${xe(d.errT || 'Valor inválido')}" error="${xe(d.err)}"` : ''}${d.msg ? ` promptTitle="${xe(d.msgT || '')}" prompt="${xe(d.msg)}"` : ''} sqref="${d.ref}">${d.f1 != null ? `<formula1>${xe(d.f1)}</formula1>` : ''}${d.f2 != null ? `<formula2>${xe(d.f2)}</formula2>` : ''}</dataValidation>`).join('')}</dataValidations>` : ''}<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.55" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="${S.paisagem ? 'landscape' : 'portrait'}" fitToWidth="1" fitToHeight="${S.altPag}"/><headerFooter><oddFooter>&amp;L&amp;8Controle de Avaria · ${xe(nome)}&amp;R&amp;8Página &amp;P de &amp;N</oddFooter></headerFooter>${S.drawing ? '<drawing r:id="rIdD"/>' : ''}${S.tabela ? '<tableParts count="1"><tablePart r:id="rIdT"/></tableParts>' : ''}</worksheet>`;
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

  // ================================================================= GERAR
  function gerar(linhas, opt = {}) {
    const agora = new Date(), hojeIso = isoLocal(agora), HOJE = serial(hojeIso);
    const prazo = opt.prazo || 7, cdNome = opt.cd || 'Centro de Distribuição', usuario = opt.usuario || '';
    const L = [...linhas].sort((a, b) => (a.ref - b.ref) || String(a.rel).localeCompare(String(b.rel)));
    const ORIG = ['Mercado Livre', 'Shopee', 'Avaria CD'];
    L.forEach(o => { if (o.origem && !ORIG.includes(o.origem)) ORIG.push(o.origem); });
    const RESP = [...new Set([usuario, ...L.map(o => o.por)].filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));

    // shared strings
    const sstArr = [], sstMap = new Map();
    const sst = s => { s = String(s); if (!sstMap.has(s)) { sstMap.set(s, sstArr.length); sstArr.push(s); } return sstMap.get(s); };
    const E = Estilos();
    // estilos de célula
    const bd = (cor = C.linha) => ({ l: ['thin', cor], r: ['thin', cor], t: ['thin', cor], b: ['thin', cor] });
    const s = {
      tit: E.xf({ font: { sz: 18, b: 1, cor: C.tinta }, al: { v: 'center' } }),
      sub: E.xf({ font: { sz: 10, cor: C.mudo } }),
      marca: E.xf({ font: { sz: 9, b: 1, cor: C.azul }, al: { h: 'left' } }),
      cdR: E.xf({ font: { sz: 10, b: 1, cor: C.tinta2 }, al: { h: 'right' } }),
      dataR: E.xf({ font: { sz: 9, cor: C.mudo }, fmt: '"Atualizado em "dd/mm/yyyy', al: { h: 'right' } }),
      secao: E.xf({ font: { sz: 11, b: 1, cor: C.azulEsc }, border: { b: ['medium', C.azul] }, al: { v: 'bottom' } }),
      secaoL: E.xf({ border: { b: ['medium', C.azul] } }),
      rotFiltro: E.xf({ font: { sz: 8, b: 1, cor: C.mudo }, al: { v: 'bottom' } }),
      inData: E.xf({ font: { sz: 10, b: 1, cor: C.tinta }, fill: C.input, border: bd('94A3B8'), fmt: 'dd/mm/yyyy', al: { h: 'center' }, livre: 1 }),
      inTxt: E.xf({ font: { sz: 10, b: 1, cor: C.tinta }, fill: C.input, border: bd('94A3B8'), al: { h: 'left', ind: 1 }, livre: 1 }),
      inNum: E.xf({ font: { sz: 10, b: 1, cor: C.tinta }, fill: C.input, border: bd('94A3B8'), fmt: '0', al: { h: 'center' }, livre: 1 }),
      dica: E.xf({ font: { sz: 8, i: 1, cor: C.fraco }, al: { wrap: 1, v: 'center' } }),
      nota: E.xf({ font: { sz: 8, i: 1, cor: C.fraco }, al: { h: 'left', v: 'center' } }),
      kRot: E.xf({ font: { sz: 8, b: 1, cor: C.mudo }, fill: C.branco, border: { l: ['thick', C.azul], t: ['thin', C.linha], r: ['thin', C.linha] }, al: { h: 'left', ind: 1, v: 'bottom' } }),
      kRotR: E.xf({ fill: C.branco, border: { t: ['thin', C.linha], r: ['thin', C.linha] } }),
      kVal: E.xf({ font: { sz: 22, b: 1, cor: C.tinta }, fill: C.branco, border: { l: ['thick', C.azul], r: ['thin', C.linha] }, fmt: '#,##0', al: { h: 'left', ind: 1 } }),
      kValP: E.xf({ font: { sz: 22, b: 1, cor: C.tinta }, fill: C.branco, border: { l: ['thick', C.azul], r: ['thin', C.linha] }, fmt: '0%', al: { h: 'left', ind: 1 } }),
      kValD: E.xf({ font: { sz: 22, b: 1, cor: C.tinta }, fill: C.branco, border: { l: ['thick', C.azul], r: ['thin', C.linha] }, fmt: '0.0', al: { h: 'left', ind: 1 } }),
      kNota: E.xf({ font: { sz: 8, cor: C.mudo }, fill: C.branco, border: { l: ['thick', C.azul], b: ['thin', C.linha], r: ['thin', C.linha] }, al: { h: 'left', ind: 1, v: 'top' } }),
      kNotaR: E.xf({ fill: C.branco, border: { b: ['thin', C.linha], r: ['thin', C.linha] } }),
      kValR: E.xf({ fill: C.branco, border: { r: ['thin', C.linha] } }),
      cab: E.xf({ font: { sz: 9, b: 1, cor: C.branco }, fill: C.azul, border: bd(C.azul), al: { h: 'center', wrap: 1 } }),
      cabE: E.xf({ font: { sz: 9, b: 1, cor: C.branco }, fill: C.azul, border: bd(C.azul), al: { h: 'left', ind: 1, wrap: 1 } }),
      tTx: E.xf({ border: bd(), al: { h: 'left', ind: 1 } }),
      tTxB: E.xf({ font: { b: 1 }, border: bd(), al: { h: 'left', ind: 1 } }),
      tNum: E.xf({ border: bd(), fmt: '#,##0', al: { h: 'center' } }),
      tNumB: E.xf({ font: { b: 1 }, border: bd(), fmt: '#,##0', al: { h: 'center' } }),
      tDec: E.xf({ border: bd(), fmt: '0.0', al: { h: 'center' } }),
      tPct: E.xf({ border: bd(), fmt: '0%', al: { h: 'center' } }),
      tPctE: E.xf({ border: bd(), fmt: '0%', al: { h: 'left', ind: 1 } }),
      tVar: E.xf({ font: { b: 1 }, border: bd(), fmt: '+0%;-0%;0%', al: { h: 'center' } }),
      tData: E.xf({ border: bd(), fmt: 'dd/mm/yyyy', al: { h: 'center' } }),
      tCen: E.xf({ border: bd(), al: { h: 'center' } }),
      tTot: E.xf({ font: { b: 1, cor: C.azulEsc }, fill: C.azulBg, border: { t: ['medium', C.azul], b: ['thin', C.linha], l: ['thin', C.linha], r: ['thin', C.linha] }, al: { h: 'left', ind: 1 } }),
      tTotN: E.xf({ font: { b: 1, cor: C.azulEsc }, fill: C.azulBg, border: { t: ['medium', C.azul], b: ['thin', C.linha], l: ['thin', C.linha], r: ['thin', C.linha] }, fmt: '#,##0', al: { h: 'center' } }),
      tTotP: E.xf({ font: { b: 1, cor: C.azulEsc }, fill: C.azulBg, border: { t: ['medium', C.azul], b: ['thin', C.linha], l: ['thin', C.linha], r: ['thin', C.linha] }, fmt: '0%', al: { h: 'center' } }),
      tTotD: E.xf({ font: { b: 1, cor: C.azulEsc }, fill: C.azulBg, border: { t: ['medium', C.azul], b: ['thin', C.linha], l: ['thin', C.linha], r: ['thin', C.linha] }, fmt: '0.0', al: { h: 'center' } }),
      // lançamentos (dentro da tabela: sem preenchimento para respeitar as faixas)
      lNum: E.xf({ font: { b: 1 }, fmt: '0', al: { h: 'center' }, livre: 1 }),
      lData: E.xf({ fmt: 'dd/mm/yyyy', al: { h: 'center' }, livre: 1 }),
      lTx: E.xf({ al: { h: 'left' }, livre: 1 }),
      lCen: E.xf({ al: { h: 'center' }, livre: 1 }),
      aTx: E.xf({ font: { cor: '475569' }, al: { h: 'center' } }),
      aNum: E.xf({ font: { cor: '475569' }, fmt: '0', al: { h: 'center' } }),
      aMes: E.xf({ font: { cor: '475569' }, fmt: 'mm/yyyy', al: { h: 'center' } }),
      aAlerta: E.xf({ font: { cor: '475569', sz: 9 }, al: { h: 'left' } }),
      tagIn: E.xf({ font: { sz: 8, b: 1, cor: C.ambar }, fill: C.ambarBg, al: { h: 'center' } }),
      tagOp: E.xf({ font: { sz: 8, b: 1, cor: C.mudo }, fill: C.auto, al: { h: 'center' } }),
      tagAu: E.xf({ font: { sz: 8, b: 1, cor: C.azul }, fill: C.azulCl, al: { h: 'center' } }),
      txt: E.xf({ font: { sz: 10, cor: C.tinta2 }, al: { wrap: 1, v: 'top' } }),
      txtB: E.xf({ font: { sz: 10, b: 1, cor: C.tinta }, al: { wrap: 1, v: 'top' } }),
      auto: E.xf({ font: { sz: 10, cor: '475569' }, fill: C.auto, border: bd(), al: { h: 'center' } }),
      autoD: E.xf({ font: { sz: 10, cor: '475569' }, fill: C.auto, border: bd(), fmt: 'dd/mm/yyyy', al: { h: 'center' } }),
      leg1: E.xf({ fill: C.input, border: bd('94A3B8') }), leg2: E.xf({ fill: C.auto, border: bd() }), leg3: E.xf({ fill: C.azul, border: bd(C.azul) }),
    };
    const dx = {
      tabInteira: E.dxf({ borda: C.linha }), tabCab: E.dxf({ cor: C.branco, bg: C.azul, b: 1 }), tabFaixa: E.dxf({ bg: 'F7F9FC' }),
      ok: E.dxf({ cor: C.azul, bg: C.azulCl, b: 1 }), pend: E.dxf({ cor: C.ambar, bg: C.ambarBg, b: 1 }), ruim: E.dxf({ cor: C.verm, bg: C.vermBg, b: 1 }),
      ambarTx: E.dxf({ cor: C.ambar, b: 1 }), vermTx: E.dxf({ cor: C.verm, b: 1 }), azulTx: E.dxf({ cor: C.azul, b: 1 }),
    };

    // ---------------------------------------------------------------- cálculo prévio (valores em cache)
    const R0 = 6; // primeira linha de dados (1-based) em LANÇAMENTOS
    const cont = {}; L.forEach(o => { cont[o.ref] = (cont[o.ref] || 0) + 1; });
    const D = L.map((o, i) => {
      const rel = serial(o.rel), ent = serial(o.ent), pend = ent == null;
      const dias = rel == null ? '' : pend ? HOJE - rel : ent - rel;
      const row = R0 + i;
      let alerta = '';
      if (cont[o.ref] > 1) alerta = 'Ref. duplicada';
      else if (!pend && rel != null && ent < rel) alerta = 'Entrega antes do relatório';
      else if (pend && dias > prazo) alerta = `Pendente há ${dias} dias`;
      else if (!o.ref || rel == null || !o.origem || !o.sku || !o.produto) alerta = 'Campo obrigatório vazio';
      const mes = rel == null ? '' : serial(String(o.rel).slice(0, 7) + '-01');
      return { ...o, relS: rel, entS: ent, pend, dias, mes, alerta, ordem: pend && dias !== '' ? dias + row / 100000 : '', row };
    });
    const nTot = D.length, nEnt = D.filter(d => !d.pend).length, nPen = nTot - nEnt;
    const media = a => a.length ? a.reduce((x, y) => x + y, 0) / a.length : '';
    const mediaEnt = media(D.filter(d => !d.pend && d.dias !== '').map(d => d.dias));
    const atrasadas = D.filter(d => d.pend && d.dias > prazo).length;

    // nomes de faixa da tabela
    const T = 'tbAvarias', last = R0 + Math.max(nTot, 1) - 1, MAXR = Math.max(last + 500, 2000);
    const col = n => `${T}[${n}]`, eu = n => `${T}[[#This Row],[${n}]]`;
    const COLS = [
      { n: 'Ref.', w: 11, tipo: 'in' }, { n: 'Data do relatório', w: 14, tipo: 'in' }, { n: 'Origem', w: 15, tipo: 'in' }, { n: 'Código', w: 12, tipo: 'in' }, { n: 'Descrição do produto', w: 44, tipo: 'in' },
      { n: 'Data de entrega', w: 14, tipo: 'op' }, { n: 'Entregue por', w: 20, tipo: 'op' }, { n: 'Observação', w: 26, tipo: 'op' },
      { n: 'Situação', w: 14, tipo: 'au', f: `IF(${eu('Data de entrega')}="","Não entregue","Entregue")` },
      { n: 'Dias', w: 11, tipo: 'au', f: `IF(${eu('Data do relatório')}="","",IF(${eu('Data de entrega')}="",cfgHoje-${eu('Data do relatório')},${eu('Data de entrega')}-${eu('Data do relatório')}))` },
      { n: 'Mês', w: 11, tipo: 'au', f: `IF(${eu('Data do relatório')}="","",DATE(YEAR(${eu('Data do relatório')}),MONTH(${eu('Data do relatório')}),1))` },
      { n: 'Alerta', w: 26, tipo: 'au', f: `IF(${eu('Ref.')}="","",IF(COUNTIF(${col('Ref.')},${eu('Ref.')})>1,"Ref. duplicada",IF(AND(${eu('Data de entrega')}<>"",${eu('Data de entrega')}<${eu('Data do relatório')}),"Entrega antes do relatório",IF(AND(${eu('Situação')}="Não entregue",N(${eu('Dias')})>cfgPrazo),"Pendente há "&${eu('Dias')}&" dias",IF(OR(${eu('Data do relatório')}="",${eu('Origem')}="",${eu('Código')}="",${eu('Descrição do produto')}=""),"Campo obrigatório vazio","")))))` },
      { n: 'Ordem pendência', w: 10, tipo: 'au', f: `IF(AND(${eu('Situação')}="Não entregue",${eu('Data do relatório')}<>""),${eu('Dias')}+ROW()/100000,"")`, oculta: true },
    ];
    // a fórmula de "Alerta" para linha vazia: ajusta a regra "obrigatório" para a ref. vazia
    COLS[11].f = `IF(AND(${eu('Ref.')}="",${eu('Descrição do produto')}=""),"",IF(COUNTIF(${col('Ref.')},${eu('Ref.')})>1,"Ref. duplicada",IF(AND(${eu('Data de entrega')}<>"",${eu('Data de entrega')}<${eu('Data do relatório')}),"Entrega antes do relatório",IF(AND(${eu('Situação')}="Não entregue",N(${eu('Dias')})>cfgPrazo),"Pendente há "&${eu('Dias')}&" dias",IF(OR(${eu('Ref.')}="",${eu('Data do relatório')}="",${eu('Origem')}="",${eu('Código')}="",${eu('Descrição do produto')}=""),"Campo obrigatório vazio","")))))`;
    const ci = n => COLS.findIndex(c => c.n === n), cl = n => colL(ci(n));
    const rg = n => `'LANÇAMENTOS'!$${cl(n)}$${R0}:$${cl(n)}$${MAXR}`; // faixa fixa (para validação/CF)

    // ================================================================ CONFIGURAÇÕES
    const cfg = Folha('CONFIGURAÇÕES', E, sst); cfg.tab = '64748B'; cfg.protegida = true;
    cfg.cols = [2, 40, 22, 4, 26, 4, 18, 4, 22];
    cfg.alt[1] = 30;
    cfg.put('B2', 'Configurações', s.tit); cfg.put('B3', 'Parâmetros e listas usados em todo o arquivo. Edite apenas as células em creme.', s.sub);
    cfg.put('B5', 'PARÂMETROS', s.secao); ['C5', 'D5', 'E5', 'F5', 'G5', 'H5', 'I5'].forEach(a => cfg.put(a, '', s.secaoL));
    cfg.put('B6', 'Nome do CD / unidade', s.tTxB); cfg.put('C6', cdNome, s.inTxt);
    cfg.put('B7', 'Prazo para alerta de pendência (dias)', s.tTxB); cfg.put('C7', prazo, s.inNum);
    cfg.put('B8', 'Data de referência (hoje)', s.tTxB); cfg.f('C8', 'TODAY()', s.autoD, HOJE);
    cfg.put('E6', 'Campos em creme podem ser editados.', s.nota); cfg.put('E7', 'Pendências acima deste prazo ficam em alerta.', s.nota); cfg.put('E8', 'Automático.', s.nota);
    cfg.put('B10', 'LISTAS SUSPENSAS', s.secao); ['C10', 'D10', 'E10', 'F10', 'G10', 'H10', 'I10'].forEach(a => cfg.put(a, '', s.secaoL));
    cfg.put('B11', 'Origens', s.cabE); cfg.put('E11', 'Responsáveis pela entrega', s.cabE); cfg.put('G11', 'Filtro: origem', s.cab); cfg.put('I11', 'Filtro: responsável', s.cab);
    const LN = 30; // linhas disponíveis nas listas
    for (let i = 0; i < LN; i++) {
      const r = 12 + i;
      cfg.put(`B${r}`, ORIG[i] || '', s.inTxt); cfg.put(`E${r}`, RESP[i] || '', s.inTxt);
    }
    cfg.put('G12', 'Todas', s.auto); cfg.put('I12', 'Todas', s.auto);
    for (let i = 0; i < LN; i++) { cfg.f(`G${13 + i}`, `IF(B${12 + i}="","",B${12 + i})`, s.auto, ORIG[i] || ''); cfg.f(`I${13 + i}`, `IF(E${12 + i}="","",E${12 + i})`, s.auto, RESP[i] || ''); }
    cfg.put('B44', 'CÁLCULOS INTERNOS — não editar', s.secao); ['C44', 'D44', 'E44', 'F44', 'G44', 'H44', 'I44'].forEach(a => cfg.put(a, '', s.secaoL));
    const internos = [
      ['Painel: data inicial', `IF(ISNUMBER('DASHBOARD'!$C$6),'DASHBOARD'!$C$6,1)`, 1, 1],
      ['Painel: data final', `IF(ISNUMBER('DASHBOARD'!$E$6),'DASHBOARD'!$E$6,2958465)`, 2958465, 1],
      ['Painel: origem (critério)', `IF(OR('DASHBOARD'!$G$6="",'DASHBOARD'!$G$6="Todas"),"*",'DASHBOARD'!$G$6)`, '*', 0],
      ['Painel: responsável (critério)', `IF(OR('DASHBOARD'!$J$6="",'DASHBOARD'!$J$6="Todas"),"<>#todos#",'DASHBOARD'!$J$6)`, '<>#todos#', 0],
      ['Relatório: data inicial', `IF(ISNUMBER('RELATÓRIOS'!$C$6),'RELATÓRIOS'!$C$6,1)`, 1, 1],
      ['Relatório: data final', `IF(ISNUMBER('RELATÓRIOS'!$E$6),'RELATÓRIOS'!$E$6,2958465)`, 2958465, 1],
    ];
    internos.forEach(([t, f, v, d], i) => { cfg.put(`B${45 + i}`, t, s.tTx); cfg.f(`C${45 + i}`, f, d ? s.autoD : s.auto, v); });
    cfg.dv.push({ type: 'whole', op: 'between', f1: '1', f2: '365', ref: 'C7', err: 'Informe um número de dias entre 1 e 365.', msgT: 'Prazo', msg: 'Dias até a pendência entrar em alerta.' });
    const names = [
      ['cfgCD', `'CONFIGURAÇÕES'!$C$6`], ['cfgPrazo', `'CONFIGURAÇÕES'!$C$7`], ['cfgHoje', `'CONFIGURAÇÕES'!$C$8`],
      ['lstOrigem', `OFFSET('CONFIGURAÇÕES'!$B$12,0,0,MAX(1,COUNTIF('CONFIGURAÇÕES'!$B$12:$B$41,"?*")),1)`],
      ['lstResp', `OFFSET('CONFIGURAÇÕES'!$E$12,0,0,MAX(1,COUNTIF('CONFIGURAÇÕES'!$E$12:$E$41,"?*")),1)`],
      ['lstOrigemF', `OFFSET('CONFIGURAÇÕES'!$G$12,0,0,1+COUNTIF('CONFIGURAÇÕES'!$B$12:$B$41,"?*"),1)`],
      ['lstRespF', `OFFSET('CONFIGURAÇÕES'!$I$12,0,0,1+COUNTIF('CONFIGURAÇÕES'!$E$12:$E$41,"?*"),1)`],
      ['fDe', `'CONFIGURAÇÕES'!$C$45`], ['fAte', `'CONFIGURAÇÕES'!$C$46`], ['fOrig', `'CONFIGURAÇÕES'!$C$47`], ['fResp', `'CONFIGURAÇÕES'!$C$48`],
      ['rDe', `'CONFIGURAÇÕES'!$C$49`], ['rAte', `'CONFIGURAÇÕES'!$C$50`],
    ];

    // ================================================================ LANÇAMENTOS
    const lan = Folha('LANÇAMENTOS', E, sst); lan.tab = C.tinta; lan.freeze = 'B6';
    lan.cols = COLS.map(c => c.w); COLS.forEach((c, i) => { if (c.oculta) lan.ocultas.add(i); });
    lan.alt[0] = 30; lan.alt[3] = 16; lan.alt[4] = 30;
    lan.put('A1', 'Lançamentos de avarias', s.tit);
    lan.put('A2', 'Base de dados do arquivo. Cada linha é uma avaria. Para incluir uma nova, digite na primeira linha vazia logo abaixo da tabela — ela se expande sozinha e as colunas azuis são calculadas.', s.sub);
    lan.put('A3', `Exportado do sistema Controle de Avaria em ${agora.toLocaleDateString('pt-BR')} às ${agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}${usuario ? ' por ' + usuario : ''}.`, s.nota);
    COLS.forEach((c, i) => lan.put(A(i, 3), c.tipo === 'in' ? 'OBRIGATÓRIO' : c.tipo === 'op' ? 'OPCIONAL' : 'AUTOMÁTICO', c.tipo === 'in' ? s.tagIn : c.tipo === 'op' ? s.tagOp : s.tagAu));
    COLS.forEach((c, i) => lan.put(A(i, 4), c.n, s.cab));
    const linhasTab = D.length ? D : [null];
    linhasTab.forEach((d, i) => {
      const r = R0 + i;
      const p = (n, v, st) => lan.put(`${cl(n)}${r}`, v, st);
      p('Ref.', d ? d.ref : '', s.lNum); p('Data do relatório', d && d.relS != null ? d.relS : '', s.lData); p('Origem', d ? d.origem || '' : '', s.lCen); p('Código', d ? String(d.sku ?? '') : '', s.lCen);
      p('Descrição do produto', d ? d.produto || '' : '', s.lTx); p('Data de entrega', d && d.entS != null ? d.entS : '', s.lData); p('Entregue por', d ? d.por || '' : '', s.lTx); p('Observação', '', s.lTx);
      const vals = d ? { 'Situação': d.pend ? 'Não entregue' : 'Entregue', 'Dias': d.dias, 'Mês': d.mes, 'Alerta': d.alerta, 'Ordem pendência': d.ordem } : { 'Situação': 'Não entregue', 'Dias': '', 'Mês': '', 'Alerta': '', 'Ordem pendência': '' };
      COLS.filter(c => c.f).forEach(c => lan.f(`${cl(c.n)}${r}`, c.f, c.n === 'Situação' ? s.aTx : c.n === 'Dias' ? s.aNum : c.n === 'Mês' ? s.aMes : c.n === 'Alerta' ? s.aAlerta : s.aNum, vals[c.n]));
    });
    const tabRef = `A5:${colL(COLS.length - 1)}${last}`;
    const tabelaXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="${T}" displayName="${T}" ref="${tabRef}" totalsRowShown="0"><autoFilter ref="${tabRef}"/><tableColumns count="${COLS.length}">${COLS.map((c, i) => `<tableColumn id="${i + 1}" name="${xe(c.n)}"${c.f ? `><calculatedColumnFormula>${xe(c.f)}</calculatedColumnFormula></tableColumn>` : '/>'}`).join('')}</tableColumns><tableStyleInfo name="AvariaCorporativo" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
    lan.tabela = true;
    const fx = (n, r = R0) => `$${cl(n)}${r}`;
    lan.cf.push({ ref: `${cl('Situação')}${R0}:${cl('Situação')}${MAXR}`, rules: [{ type: 'cellIs', op: 'equal', f: ['"Entregue"'], dxf: dx.ok }, { type: 'cellIs', op: 'equal', f: ['"Não entregue"'], dxf: dx.pend }] });
    lan.cf.push({ ref: `${cl('Alerta')}${R0}:${cl('Alerta')}${MAXR}`, rules: [
      { type: 'containsText', op: 'containsText', text: 'duplicada', f: [`NOT(ISERROR(SEARCH("duplicada",${cl('Alerta')}${R0})))`], dxf: dx.ruim },
      { type: 'containsText', op: 'containsText', text: 'antes', f: [`NOT(ISERROR(SEARCH("antes",${cl('Alerta')}${R0})))`], dxf: dx.ruim },
      { type: 'containsText', op: 'containsText', text: 'Pendente', f: [`NOT(ISERROR(SEARCH("Pendente",${cl('Alerta')}${R0})))`], dxf: dx.pend },
      { type: 'containsText', op: 'containsText', text: 'obrigatório', f: [`NOT(ISERROR(SEARCH("obrigatório",${cl('Alerta')}${R0})))`], dxf: dx.pend }] });
    lan.cf.push({ ref: `A${R0}:A${MAXR}`, rules: [{ type: 'expression', f: [`AND($A${R0}<>"",COUNTIF($A$${R0}:$A$${MAXR},$A${R0})>1)`], dxf: dx.ruim }] });
    lan.cf.push({ ref: `${cl('Dias')}${R0}:${cl('Dias')}${MAXR}`, rules: [{ type: 'expression', f: [`AND(${fx('Situação')}="Não entregue",N(${fx('Dias')})>cfgPrazo)`], dxf: dx.vermTx }] });
    const lr = n => `${cl(n)}${R0}:${cl(n)}${MAXR}`;
    lan.dv.push({ type: 'whole', op: 'greaterThanOrEqual', f1: '1', ref: lr('Ref.'), err: 'A Ref. deve ser um número inteiro maior que zero.', msgT: 'Ref. (obrigatório)', msg: 'Número da referência da avaria.' });
    lan.dv.push({ type: 'date', op: 'between', f1: '43831', f2: '73050', ref: lr('Data do relatório'), err: 'Digite uma data válida no formato dd/mm/aaaa.', msgT: 'Data do relatório (obrigatório)', msg: 'Formato dd/mm/aaaa.' });
    lan.dv.push({ type: 'list', f1: 'lstOrigem', ref: lr('Origem'), err: 'Escolha uma origem da lista. Para criar outra, use a aba CONFIGURAÇÕES.', msgT: 'Origem (obrigatório)', msg: 'Escolha na lista.' });
    lan.dv.push({ type: 'textLength', op: 'between', f1: '1', f2: '40', ref: lr('Código'), err: 'O código deve ter entre 1 e 40 caracteres.', msgT: 'Código (obrigatório)', msg: 'Código do produto no WMS.' });
    lan.dv.push({ type: 'textLength', op: 'between', f1: '1', f2: '250', ref: lr('Descrição do produto'), err: 'A descrição deve ter até 250 caracteres.', msgT: 'Descrição (obrigatório)', msg: 'Descrição do produto avariado.' });
    lan.dv.push({ type: 'custom', f1: `OR(${cl('Data de entrega')}${R0}="",AND(ISNUMBER(${cl('Data de entrega')}${R0}),${cl('Data de entrega')}${R0}>=${cl('Data do relatório')}${R0}))`, ref: lr('Data de entrega'), err: 'A data de entrega deve ser uma data igual ou posterior à data do relatório.', msgT: 'Data de entrega (opcional)', msg: 'Preencha quando o relatório for entregue. A situação muda sozinha.' });
    lan.dv.push({ type: 'list', f1: 'lstResp', ref: lr('Entregue por'), estilo: 'warning', err: 'Esse nome não está na lista de responsáveis. Deseja continuar mesmo assim?', errT: 'Responsável fora da lista', msgT: 'Entregue por (opcional)', msg: 'Quem entregou o relatório.' });

    // ================================================================ ANÁLISES
    const ana = Folha('ANÁLISES', E, sst); ana.tab = '475569'; ana.protegida = true;
    ana.cols = [2, 26, 14, 14, 14, 14, 14, 14, 12];
    ana.alt[1] = 30;
    ana.put('B2', 'Análises', s.tit);
    ana.f('B3', `"Tendências e comparações · origem: "&IF(fOrig="*","todas",fOrig)&" (filtro do DASHBOARD)"`, s.sub, 'Tendências e comparações · origem: todas (filtro do DASHBOARD)');
    ana.put('B5', 'EVOLUÇÃO MENSAL — ÚLTIMOS 12 MESES', s.secao); ['C5', 'D5', 'E5', 'F5', 'G5', 'H5', 'I5'].forEach(a => ana.put(a, '', s.secaoL));
    ['Mês', 'Registradas', 'Entregues no mês', 'Pendentes (do mês)', '% resolvido', 'Média de dias', 'Acima do prazo', 'Início do mês'].forEach((t, i) => ana.put(A(1 + i, 5), t, i ? s.cab : s.cabE));
    ana.alt[5] = 30;
    const meses = []; const hj = new Date(agora.getFullYear(), agora.getMonth(), 1);
    for (let i = 11; i >= 0; i--) { const d = new Date(hj.getFullYear(), hj.getMonth() - i, 1); meses.push({ iso: isoLocal(d), prox: isoLocal(new Date(d.getFullYear(), d.getMonth() + 1, 1)), rot: MES[d.getMonth()] + '/' + String(d.getFullYear()).slice(2) }); }
    const RB = `${col('Data do relatório')}`, RE = `${col('Data de entrega')}`, RO = col('Origem'), RS = col('Situação'), RD = col('Dias'), RP = col('Entregue por');
    const mesCalc = meses.map(m => {
      const ini = serial(m.iso), fim = serial(m.prox);
      const reg = D.filter(d => d.relS >= ini && d.relS < fim), ent = D.filter(d => d.entS != null && d.entS >= ini && d.entS < fim);
      return { reg: reg.length, ent: ent.length, pen: reg.filter(d => d.pend).length, med: media(ent.map(d => d.dias).filter(v => v !== '')), atr: reg.filter(d => d.pend && d.dias > prazo).length };
    });
    meses.forEach((m, i) => {
      const r = 7 + i, c = mesCalc[i];
      ana.f(`I${r}`, `DATE(YEAR(cfgHoje),MONTH(cfgHoje)-${11 - i},1)`, s.autoD, serial(m.iso));
      ana.f(`B${r}`, `CHOOSE(MONTH($I${r}),"jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez")&"/"&RIGHT(YEAR($I${r}),2)`, s.tTxB, m.rot);
      ana.f(`C${r}`, `COUNTIFS(${RB},">="&$I${r},${RB},"<"&EDATE($I${r},1),${RO},fOrig)`, s.tNum, c.reg);
      ana.f(`D${r}`, `COUNTIFS(${RE},">="&$I${r},${RE},"<"&EDATE($I${r},1),${RO},fOrig)`, s.tNum, c.ent);
      ana.f(`E${r}`, `COUNTIFS(${RB},">="&$I${r},${RB},"<"&EDATE($I${r},1),${RO},fOrig,${RS},"Não entregue")`, s.tNum, c.pen);
      ana.f(`F${r}`, `IF(C${r}=0,"",(C${r}-E${r})/C${r})`, s.tPct, c.reg ? (c.reg - c.pen) / c.reg : '');
      ana.f(`G${r}`, `IFERROR(AVERAGEIFS(${RD},${RE},">="&$I${r},${RE},"<"&EDATE($I${r},1),${RO},fOrig),"")`, s.tDec, c.med);
      ana.f(`H${r}`, `COUNTIFS(${RB},">="&$I${r},${RB},"<"&EDATE($I${r},1),${RO},fOrig,${RS},"Não entregue",${RD},">"&cfgPrazo)`, s.tNum, c.atr);
    });
    const sm = k => mesCalc.reduce((a, c) => a + c[k], 0);
    ana.put('B19', 'Total do período', s.tTot);
    ana.f('C19', 'SUM(C7:C18)', s.tTotN, sm('reg')); ana.f('D19', 'SUM(D7:D18)', s.tTotN, sm('ent')); ana.f('E19', 'SUM(E7:E18)', s.tTotN, sm('pen'));
    ana.f('F19', 'IF(C19=0,"",(C19-E19)/C19)', s.tTotP, sm('reg') ? (sm('reg') - sm('pen')) / sm('reg') : '');
    const entJan = D.filter(d => d.entS != null && d.entS >= serial(meses[0].iso));
    ana.f('G19', `IFERROR(AVERAGEIFS(${RD},${RE},">="&I7,${RO},fOrig),"")`, s.tTotD, media(entJan.map(d => d.dias).filter(v => v !== '')));
    ana.f('H19', 'SUM(H7:H18)', s.tTotN, sm('atr')); ana.put('I19', '', s.tTot);
    ana.cf.push({ ref: 'F7:F19', rules: [{ type: 'cellIs', op: 'lessThan', f: ['0.5'], dxf: dx.ambarTx }] });
    ana.cf.push({ ref: 'H7:H19', rules: [{ type: 'cellIs', op: 'greaterThan', f: ['0'], dxf: dx.vermTx }] });

    ana.put('B21', 'COMPARATIVO DE PERÍODOS', s.secao); ['C21', 'D21', 'E21', 'F21', 'G21', 'H21', 'I21'].forEach(a => ana.put(a, '', s.secaoL));
    ['Indicador', 'Mês atual', 'Mês anterior', 'Variação', 'Últimos 30 dias', '30 dias anteriores', 'Variação'].forEach((t, i) => ana.put(A(1 + i, 21), t, i ? s.cab : s.cabE));
    ana.alt[21] = 30;
    const jan = (a, b, campo) => D.filter(d => d[campo] != null && d[campo] >= a && d[campo] <= b);
    const v30 = { reg: jan(HOJE - 29, HOJE, 'relS').length, reg0: jan(HOJE - 59, HOJE - 30, 'relS').length, ent: jan(HOJE - 29, HOJE, 'entS').length, ent0: jan(HOJE - 59, HOJE - 30, 'entS').length };
    const varc = (a, b) => b ? a / b - 1 : '';
    const cmp = [
      ['Avarias registradas', 'C18', 'C17', mesCalc[11].reg, mesCalc[10].reg, `COUNTIFS(${RB},">="&(cfgHoje-29),${RB},"<="&cfgHoje,${RO},fOrig)`, `COUNTIFS(${RB},">="&(cfgHoje-59),${RB},"<="&(cfgHoje-30),${RO},fOrig)`, v30.reg, v30.reg0, s.tNum],
      ['Entregas realizadas', 'D18', 'D17', mesCalc[11].ent, mesCalc[10].ent, `COUNTIFS(${RE},">="&(cfgHoje-29),${RE},"<="&cfgHoje,${RO},fOrig)`, `COUNTIFS(${RE},">="&(cfgHoje-59),${RE},"<="&(cfgHoje-30),${RO},fOrig)`, v30.ent, v30.ent0, s.tNum],
      ['Média de dias até a entrega', 'G18', 'G17', mesCalc[11].med, mesCalc[10].med, `IFERROR(AVERAGEIFS(${RD},${RE},">="&(cfgHoje-29),${RE},"<="&cfgHoje,${RO},fOrig),"")`, `IFERROR(AVERAGEIFS(${RD},${RE},">="&(cfgHoje-59),${RE},"<="&(cfgHoje-30),${RO},fOrig),"")`, media(jan(HOJE - 29, HOJE, 'entS').map(d => d.dias)), media(jan(HOJE - 59, HOJE - 30, 'entS').map(d => d.dias)), s.tDec],
    ];
    cmp.forEach(([t, a1, a2, v1, v2, f1, f2, w1, w2, st], i) => {
      const r = 23 + i;
      ana.put(`B${r}`, t, s.tTxB); ana.f(`C${r}`, a1, st, v1); ana.f(`D${r}`, a2, st, v2);
      ana.f(`E${r}`, `IF(OR(C${r}="",D${r}="",D${r}=0),"",C${r}/D${r}-1)`, s.tVar, typeof v1 === 'number' && v2 ? varc(v1, v2) : '');
      ana.f(`F${r}`, f1, st, w1); ana.f(`G${r}`, f2, st, w2);
      ana.f(`H${r}`, `IF(OR(F${r}="",G${r}="",G${r}=0),"",F${r}/G${r}-1)`, s.tVar, typeof w1 === 'number' && w2 ? varc(w1, w2) : '');
    });
    ana.cf.push({ ref: 'E23:E24 H23:H24', rules: [{ type: 'cellIs', op: 'greaterThan', f: ['0'], dxf: dx.azulTx }, { type: 'cellIs', op: 'lessThan', f: ['0'], dxf: dx.ambarTx }] });
    ana.cf.push({ ref: 'E25 H25', rules: [{ type: 'cellIs', op: 'greaterThan', f: ['0'], dxf: dx.ambarTx }, { type: 'cellIs', op: 'lessThan', f: ['0'], dxf: dx.azulTx }] });

    ana.put('B27', 'DESEMPENHO POR ORIGEM (período do DASHBOARD)', s.secao); ['C27', 'D27', 'E27', 'F27', 'G27', 'H27', 'I27'].forEach(a => ana.put(a, '', s.secaoL));
    ['Origem', 'Avarias', 'Entregues', 'Pendentes', '% entregue', 'Média de dias', 'Acima do prazo'].forEach((t, i) => ana.put(A(1 + i, 27), t, i ? s.cab : s.cabE));
    ana.alt[27] = 30;
    const per = `${RB},">="&fDe,${RB},"<="&fAte`;
    const oCalc = ORIG.map(g => { const a = D.filter(d => d.origem === g); return { n: a.length, e: a.filter(d => !d.pend).length, p: a.filter(d => d.pend).length, m: media(a.filter(d => !d.pend).map(d => d.dias).filter(v => v !== '')), x: a.filter(d => d.pend && d.dias > prazo).length }; });
    ORIG.forEach((g, i) => {
      const r = 29 + i, c = oCalc[i];
      ana.f(`B${r}`, `'CONFIGURAÇÕES'!B${12 + i}`, s.tTxB, g);
      ana.f(`C${r}`, `COUNTIFS(${RO},$B${r},${per})`, s.tNum, c.n);
      ana.f(`D${r}`, `COUNTIFS(${RO},$B${r},${per},${RS},"Entregue")`, s.tNum, c.e);
      ana.f(`E${r}`, `COUNTIFS(${RO},$B${r},${per},${RS},"Não entregue")`, s.tNum, c.p);
      ana.f(`F${r}`, `IF(C${r}=0,"",D${r}/C${r})`, s.tPct, c.n ? c.e / c.n : '');
      ana.f(`G${r}`, `IFERROR(AVERAGEIFS(${RD},${RO},$B${r},${per},${RS},"Entregue"),"")`, s.tDec, c.m);
      ana.f(`H${r}`, `COUNTIFS(${RO},$B${r},${per},${RS},"Não entregue",${RD},">"&cfgPrazo)`, s.tNum, c.x);
    });
    const rT = 29 + ORIG.length;
    ana.put(`B${rT}`, 'Total', s.tTot);
    ana.f(`C${rT}`, `SUM(C29:C${rT - 1})`, s.tTotN, nTot); ana.f(`D${rT}`, `SUM(D29:D${rT - 1})`, s.tTotN, nEnt); ana.f(`E${rT}`, `SUM(E29:E${rT - 1})`, s.tTotN, nPen);
    ana.f(`F${rT}`, `IF(C${rT}=0,"",D${rT}/C${rT})`, s.tTotP, nTot ? nEnt / nTot : ''); ana.f(`G${rT}`, `IFERROR(AVERAGEIFS(${RD},${per},${RS},"Entregue"),"")`, s.tTotD, mediaEnt); ana.f(`H${rT}`, `SUM(H29:H${rT - 1})`, s.tTotN, atrasadas);
    ana.cf.push({ ref: `H29:H${rT}`, rules: [{ type: 'cellIs', op: 'greaterThan', f: ['0'], dxf: dx.vermTx }] });

    // ================================================================ RELATÓRIOS
    const rel = Folha('RELATÓRIOS', E, sst); rel.tab = '475569'; rel.protegida = true; rel.freeze = null;
    rel.cols = [2, 7, 9, 13, 15, 12, 40, 12, 14];
    rel.alt[1] = 30;
    rel.put('B2', 'Relatórios', s.tit); rel.put('B3', 'Relatórios consolidados, gerados automaticamente a partir da aba LANÇAMENTOS.', s.sub);
    rel.put('B5', 'PERÍODO DE', s.rotFiltro); rel.put('E5', 'ATÉ', s.rotFiltro);
    rel.merge('C6', 'D6', '', s.inData); rel.merge('E6', 'F6', '', s.inData); rel.put('G6', 'Filtra o quadro de entregas por responsável. Em branco = todo o período.', s.nota);
    rel.put('B6', '', 0);
    rel.dv.push({ type: 'date', op: 'between', f1: '43831', f2: '73050', ref: 'C6 E6', err: 'Digite uma data válida (dd/mm/aaaa).', msgT: 'Período', msg: 'Data no formato dd/mm/aaaa.' });
    rel.put('B8', 'ENTREGAS POR RESPONSÁVEL', s.secao); ['C8', 'D8', 'E8', 'F8', 'G8', 'H8', 'I8'].forEach(a => rel.put(a, '', s.secaoL));
    rel.merge('B9', 'D9', 'Responsável', s.cabE);
    [['E9', 'Entregas'], ['F9', 'Média de dias'], ['G9', 'Participação nas entregas'], ['H9', 'Mercado Livre'], ['I9', 'Shopee']].forEach(([a, t]) => rel.put(a, t, s.cab));
    rel.alt[8] = 30; 
    const perR = `${RE},">="&rDe,${RE},"<="&rAte`;
    const respRows = [...RESP, '(sem responsável)'];
    const fimR = 9 + respRows.length;
    respRows.forEach((p, i) => {
      const r = 10 + i, sem = p === '(sem responsável)';
      const crit = sem ? `${RP},""` : `${RP},$B${r}`;
      const a = D.filter(d => !d.pend && (sem ? !d.por : d.por === p));
      if (sem) rel.merge(`B${r}`, `D${r}`, p, s.tTx); else rel.merge(`B${r}`, `D${r}`, { f: `IF('CONFIGURAÇÕES'!E${12 + i}="","—",'CONFIGURAÇÕES'!E${12 + i})`, v: p }, s.tTxB);
      rel.f(`E${r}`, `COUNTIFS(${crit},${perR})`, s.tNumB, a.length);
      rel.f(`F${r}`, `IFERROR(AVERAGEIFS(${RD},${crit},${perR}),"")`, s.tDec, media(a.map(d => d.dias).filter(v => v !== '')));
      rel.f(`G${r}`, `IF(SUM($E$10:$E$${fimR})=0,0,E${r}/SUM($E$10:$E$${fimR}))`, s.tPctE, nEnt ? a.length / nEnt : 0);
      rel.f(`H${r}`, `COUNTIFS(${crit},${perR},${RO},"Mercado Livre")`, s.tNum, a.filter(d => d.origem === 'Mercado Livre').length);
      rel.f(`I${r}`, `COUNTIFS(${crit},${perR},${RO},"Shopee")`, s.tNum, a.filter(d => d.origem === 'Shopee').length);
    });
    rel.cf.push({ ref: `G10:G${fimR}`, rules: [{ type: 'dataBar', cor: '93B4DC' }] });
    const rTR = 10 + respRows.length;
    rel.merge(`B${rTR}`, `D${rTR}`, 'Total', s.tTot);
    rel.f(`E${rTR}`, `SUM(E10:E${rTR - 1})`, s.tTotN, nEnt); rel.f(`F${rTR}`, `IFERROR(AVERAGEIFS(${RD},${perR}),"")`, s.tTotD, mediaEnt);
    rel.f(`G${rTR}`, `IF(E${rTR}=0,0,1)`, s.tTotP, nEnt ? 1 : 0);
    rel.f(`H${rTR}`, `SUM(H10:H${rTR - 1})`, s.tTotN, D.filter(d => !d.pend && d.origem === 'Mercado Livre').length); rel.f(`I${rTR}`, `SUM(I10:I${rTR - 1})`, s.tTotN, D.filter(d => !d.pend && d.origem === 'Shopee').length);

    const rF = rTR + 2;
    rel.put(`B${rF}`, 'PENDÊNCIAS POR TEMPO EM ABERTO', s.secao); ['C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(c => rel.put(`${c}${rF}`, '', s.secaoL));
    const faixas = [['Até 3 dias', 0, 3], ['4 a 7 dias', 4, 7], ['8 a 15 dias', 8, 15], ['16 a 30 dias', 16, 30], ['Mais de 30 dias', 31, 99999]];
    rel.merge(`B${rF + 1}`, `D${rF + 1}`, 'Faixa', s.cabE); rel.put(`E${rF + 1}`, 'Pendências', s.cab); rel.put(`F${rF + 1}`, '% do total', s.cab);
    faixas.forEach(([t, a, b], i) => {
      const r = rF + 2 + i, n = D.filter(d => d.pend && d.dias >= a && d.dias <= b).length;
      rel.merge(`B${r}`, `D${r}`, t, s.tTxB);
      rel.f(`E${r}`, `COUNTIFS(${RS},"Não entregue",${RD},">=${a}",${RD},"<=${b}")`, s.tNum, n);
      rel.f(`F${r}`, `IF(SUM($E$${rF + 2}:$E$${rF + 6})=0,"",E${r}/SUM($E$${rF + 2}:$E$${rF + 6}))`, s.tPct, nPen ? n / nPen : '');
    });
    rel.cf.push({ ref: `E${rF + 4}:E${rF + 6}`, rules: [{ type: 'cellIs', op: 'greaterThan', f: ['0'], dxf: dx.vermTx }] });

    const rL = rF + 9;
    rel.put(`B${rL}`, 'PENDÊNCIAS — DA MAIS ANTIGA PARA A MAIS RECENTE', s.secao); ['C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(c => rel.put(`${c}${rL}`, '', s.secaoL));
    ['Nº', 'Ref.', 'Relatório', 'Origem', 'Código', 'Descrição do produto', 'Dias em aberto', 'Status'].forEach((t, i) => rel.put(A(1 + i, rL), t, i === 5 ? s.cabE : s.cab));
    const pendOrd = D.filter(d => d.ordem !== '').sort((a, b) => b.ordem - a.ordem);
    const NL = Math.min(500, Math.max(30, pendOrd.length + 20));
    const OP = col('Ordem pendência');
    for (let k = 1; k <= NL; k++) {
      const r = rL + 1 + k, d = pendOrd[k - 1];
      const ix = c => `IFERROR(INDEX(${col(c)},MATCH(LARGE(${OP},$B${r}),${OP},0)),"")`;
      rel.put(`B${r}`, k, s.tCen);
      rel.f(`C${r}`, ix('Ref.'), s.tNumB, d ? d.ref : '');
      rel.f(`D${r}`, ix('Data do relatório'), s.tData, d ? d.relS : '');
      rel.f(`E${r}`, ix('Origem'), s.tCen, d ? d.origem : '');
      rel.f(`F${r}`, ix('Código'), s.tCen, d ? String(d.sku ?? '') : '');
      rel.f(`G${r}`, ix('Descrição do produto'), s.tTx, d ? d.produto : '');
      rel.f(`H${r}`, ix('Dias'), s.tNum, d ? d.dias : '');
      rel.f(`I${r}`, `IF(H${r}="","",IF(H${r}>cfgPrazo,"Atrasada","No prazo"))`, s.tCen, d ? (d.dias > prazo ? 'Atrasada' : 'No prazo') : '');
    }
    rel.cf.push({ ref: `I${rL + 2}:I${rL + 1 + NL}`, rules: [{ type: 'cellIs', op: 'equal', f: ['"Atrasada"'], dxf: dx.ruim }, { type: 'cellIs', op: 'equal', f: ['"No prazo"'], dxf: dx.ok }] });
    rel.cf.push({ ref: `B${rL + 2}:B${rL + 1 + NL}`, rules: [{ type: 'expression', f: [`$C${rL + 2}=""`], dxf: E.dxf({ cor: 'E5E7EB' }) }] });

    // ================================================================ DASHBOARD
    const dsh = Folha('DASHBOARD', E, sst); dsh.tab = C.azul; dsh.protegida = true; dsh.zoom = 90; dsh.altPag = 1;
    dsh.cols = [2, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 2];
    dsh.alt[0] = 8; dsh.alt[1] = 34; dsh.alt[2] = 18; dsh.alt[3] = 8; dsh.alt[4] = 16; dsh.alt[5] = 24; dsh.alt[6] = 10; dsh.alt[7] = 22; dsh.alt[8] = 40; dsh.alt[9] = 22;
    dsh.put('B2', 'Controle de Avaria · Painel Executivo', s.tit);
    dsh.merge('J2', 'M2', { f: 'cfgCD', v: cdNome }, s.cdR);
    dsh.put('B3', 'Indicadores calculados automaticamente a partir da aba LANÇAMENTOS. Altere os filtros nas células em creme.', s.sub);
    dsh.merge('J3', 'M3', { f: 'cfgHoje', v: HOJE }, s.dataR);
    dsh.put('B5', 'PERÍODO DO RELATÓRIO — DE', s.rotFiltro); dsh.put('E5', 'ATÉ', s.rotFiltro); dsh.put('G5', 'ORIGEM', s.rotFiltro); dsh.put('J5', 'RESPONSÁVEL PELA ENTREGA', s.rotFiltro);
    dsh.merge('C6', 'D6', '', s.inData); dsh.put('B6', 'Período', s.marca);
    dsh.merge('E6', 'F6', '', s.inData); dsh.merge('G6', 'I6', 'Todas', s.inTxt); dsh.merge('J6', 'L6', 'Todas', s.inTxt);
    dsh.put('B5', 'DE', s.rotFiltro); dsh.put('C5', 'PERÍODO — DE', s.rotFiltro); dsh.put('B5', '', 0); dsh.put('B6', '', 0);
    dsh.put('M6', 'Em branco = tudo', s.nota);
    dsh.dv.push({ type: 'date', op: 'between', f1: '43831', f2: '73050', ref: 'C6 E6', err: 'Digite uma data válida (dd/mm/aaaa).', msgT: 'Período', msg: 'Filtra pela data do relatório de avaria.' });
    dsh.dv.push({ type: 'list', f1: 'lstOrigemF', ref: 'G6', err: 'Escolha uma opção da lista.', msgT: 'Origem', msg: 'Escolha "Todas" ou uma origem.' });
    dsh.dv.push({ type: 'list', f1: 'lstRespF', ref: 'J6', err: 'Escolha uma opção da lista.', msgT: 'Responsável', msg: 'Filtra as avarias entregues por esta pessoa.' });
    const base = `${RB},">="&fDe,${RB},"<="&fAte,${RO},fOrig,${RP},fResp`;
    const kpis = [
      ['TOTAL DE AVARIAS', `COUNTIFS(${base})`, nTot, s.kVal, '"no período e filtros selecionados"', 'no período e filtros selecionados'],
      ['ENTREGUES', `COUNTIFS(${base},${RS},"Entregue")`, nEnt, s.kVal, `IF(B9=0,"—",TEXT(D9/B9*100,"0")&"% do total")`, nTot ? Math.round(nEnt / nTot * 100) + '% do total' : '—'],
      ['PENDENTES', `COUNTIFS(${base},${RS},"Não entregue")`, nPen, s.kVal, '"aguardando entrega do relatório"', 'aguardando entrega do relatório'],
      ['TAXA DE ENTREGA', `IF(B9=0,0,D9/B9)`, nTot ? nEnt / nTot : 0, s.kValP, `IF('ANÁLISES'!F18="","sem avarias no mês atual","mês atual: "&TEXT('ANÁLISES'!F18*100,"0")&"%")`, mesCalc[11].reg ? 'mês atual: ' + Math.round((mesCalc[11].reg - mesCalc[11].pen) / mesCalc[11].reg * 100) + '%' : 'sem avarias no mês atual'],
      ['MÉDIA DE DIAS ATÉ A ENTREGA', `IFERROR(AVERAGEIFS(${RD},${base},${RS},"Entregue"),0)`, mediaEnt === '' ? 0 : mediaEnt, s.kValD, '"do relatório até a entrega"', 'do relatório até a entrega'],
      ['PENDENTES ACIMA DO PRAZO', `COUNTIFS(${base},${RS},"Não entregue",${RD},">"&cfgPrazo)`, atrasadas, s.kVal, '"prazo configurado: "&cfgPrazo&" dias"', `prazo configurado: ${prazo} dias`],
    ];
    kpis.forEach(([t, f, v, st, nf, nv], i) => {
      const c1 = 1 + i * 2, a = colL(c1), b = colL(c1 + 1);
      dsh.merge(`${a}8`, `${b}8`, t, s.kRot); dsh.put(`${b}8`, '', s.kRotR);
      dsh.merge(`${a}9`, `${b}9`, { f, v }, st); dsh.put(`${b}9`, '', s.kValR);
      dsh.merge(`${a}10`, `${b}10`, { f: nf, v: nv }, s.kNota); dsh.put(`${b}10`, '', s.kNotaR);
    });
    // corrigir referências B9/D9 usadas nas notas (cards ocupam B,D,F,H,J,L)
    dsh.cf.push({ ref: 'L9', rules: [{ type: 'cellIs', op: 'greaterThan', f: ['0'], dxf: dx.vermTx }] });
    dsh.cf.push({ ref: 'F9', rules: [{ type: 'cellIs', op: 'greaterThan', f: ['0'], dxf: dx.ambarTx }] });
    dsh.put('B12', 'TENDÊNCIA MENSAL — ÚLTIMOS 12 MESES', s.secao); ['C12', 'D12', 'E12', 'F12', 'G12', 'H12'].forEach(a => dsh.put(a, '', s.secaoL));
    dsh.put('J12', 'SITUAÇÃO NO PERÍODO', s.secao); ['K12', 'L12', 'M12'].forEach(a => dsh.put(a, '', s.secaoL));
    dsh.put('B31', 'AVARIAS POR ORIGEM', s.secao); ['C31', 'D31', 'E31', 'F31', 'G31', 'H31'].forEach(a => dsh.put(a, '', s.secaoL));
    dsh.put('J31', 'PENDÊNCIAS MAIS ANTIGAS', s.secao); ['K31', 'L31', 'M31'].forEach(a => dsh.put(a, '', s.secaoL));
    dsh.put('J32', 'Ref.', s.cab); dsh.merge('K32', 'L32', 'Origem', s.cab); dsh.put('M32', 'Dias', s.cab);
    for (let k = 1; k <= 10; k++) {
      const r = 32 + k, d = pendOrd[k - 1];
      const ix = c => `IFERROR(INDEX(${col(c)},MATCH(LARGE(${OP},${k}),${OP},0)),"")`;
      dsh.f(`J${r}`, ix('Ref.'), s.tNumB, d ? d.ref : '');
      dsh.merge(`K${r}`, `L${r}`, { f: ix('Origem'), v: d ? d.origem : '' }, s.tCen);
      dsh.f(`M${r}`, ix('Dias'), s.tNum, d ? d.dias : '');
    }
    dsh.cf.push({ ref: 'M33:M42', rules: [{ type: 'expression', f: ['AND(ISNUMBER(M33),M33>cfgPrazo)'], dxf: dx.ruim }] });
    dsh.put('J44', 'Lista completa na aba RELATÓRIOS.', s.nota);
    dsh.put('B51', 'Como atualizar: baixe novamente pelo sistema Controle de Avaria (botão “Baixar Excel”) ou lance novas avarias na aba LANÇAMENTOS. Ajuda completa na aba AJUDA.', s.nota);
    // painel de situação para a rosca (células de apoio, ocultas pelo gráfico)
    dsh.put('O8', 'Entregues', s.dica); dsh.f('P8', 'D9', 0, nEnt); dsh.put('O9', 'Pendentes', s.dica); dsh.f('P9', 'F9', 0, nPen);
    dsh.cols[14] = 12; dsh.cols[15] = 8; dsh.ocultas.add(14); dsh.ocultas.add(15);
    // a nota do card "Entregues" usa B9 (total) e D9 (entregues)
    dsh.drawing = true;

    // ================================================================ AJUDA
    const aj = Folha('AJUDA', E, sst); aj.tab = '94A3B8'; aj.protegida = true; aj.paisagem = false;
    aj.cols = [2, 4, 30, 80];
    aj.alt[1] = 30;
    aj.put('B2', 'Como usar esta planilha', s.tit);
    aj.put('B3', 'Guia rápido para consultar, filtrar e atualizar as informações.', s.sub);
    const blocos = [
      ['ABAS', [
        ['DASHBOARD', 'Painel executivo: indicadores, gráficos e pendências mais antigas. Use os filtros de período, origem e responsável no topo.'],
        ['LANÇAMENTOS', 'Base de dados. Uma linha por avaria. É a única aba onde se digitam avarias.'],
        ['RELATÓRIOS', 'Entregas por responsável (com filtro de período), pendências por tempo em aberto e a lista completa de pendências.'],
        ['ANÁLISES', 'Evolução dos últimos 12 meses, comparativo mês atual × anterior e últimos 30 dias, e desempenho por origem.'],
        ['CONFIGURAÇÕES', 'Nome do CD, prazo de alerta e as listas suspensas (origens e responsáveis).'],
      ]],
      ['CORES E CAMPOS', [
        ['Célula creme', 'Você pode preencher ou alterar (filtros, parâmetros e listas).', s.leg1],
        ['Célula cinza', 'Automática. É calculada por fórmula e está protegida contra alteração.', s.leg2],
        ['OBRIGATÓRIO', 'Na aba LANÇAMENTOS: Ref., Data do relatório, Origem, Código e Descrição.'],
        ['OPCIONAL', 'Data de entrega, Entregue por e Observação. Ao preencher a data de entrega, a situação muda para “Entregue”.'],
        ['AUTOMÁTICO', 'Situação, Dias, Mês e Alerta são calculados sozinhos — não digite nessas colunas.'],
      ]],
      ['ALERTAS AUTOMÁTICOS', [
        ['Ref. duplicada', 'A mesma Ref. aparece em mais de uma linha. Fica em vermelho na coluna Ref. e em Alerta.'],
        ['Pendente há N dias', 'A avaria passou do prazo definido em CONFIGURAÇÕES e ainda não foi entregue.'],
        ['Entrega antes do relatório', 'A data de entrega é anterior à data do relatório — confira as datas.'],
        ['Campo obrigatório vazio', 'Falta preencher algum campo obrigatório da linha.'],
      ]],
      ['COMO ATUALIZAR', [
        ['Pelo sistema', 'No Controle de Avaria, abra Planilha e clique em “Baixar Excel”. O arquivo sai com todas as avarias do INVENTRA.'],
        ['Manualmente', 'Na aba LANÇAMENTOS, digite na primeira linha vazia logo abaixo da tabela. Todos os indicadores se atualizam sozinhos.'],
        ['Filtros', 'Na tabela LANÇAMENTOS use as setas do cabeçalho. No DASHBOARD use as células em creme.'],
        ['Proteção', 'As abas de painel e relatórios estão protegidas (sem senha) para evitar alterações acidentais. Para editar a estrutura: Revisão › Desproteger Planilha.'],
      ]],
    ];
    let r = 5;
    blocos.forEach(([tt, itens]) => {
      aj.put(`B${r}`, tt, s.secao); aj.put(`C${r}`, '', s.secaoL); aj.put(`D${r}`, '', s.secaoL); r++;
      itens.forEach(([a, b, leg]) => { if (leg) aj.put(`B${r}`, '', leg); aj.put(`C${r}`, a, s.txtB); aj.put(`D${r}`, b, s.txt); aj.alt[r - 1] = 30; r++; });
      r++;
    });
    aj.put(`C${r}`, 'Controle de Avaria · criado por Lucas Carvalho', s.nota);

    // ================================================================ GRÁFICOS (DASHBOARD)
    const rot = meses.map(m => m.rot);
    const g1 = grafBarras({ titulo: 'Avarias registradas × entregues por mês', cat: { ref: `'ANÁLISES'!$B$7:$B$18`, vals: rot }, series: [
      { nome: 'Registradas', nomeRef: `'ANÁLISES'!$C$6`, ref: `'ANÁLISES'!$C$7:$C$18`, vals: mesCalc.map(c => c.reg), cor: C.azul },
      { nome: 'Entregues no mês', nomeRef: `'ANÁLISES'!$D$6`, ref: `'ANÁLISES'!$D$7:$D$18`, vals: mesCalc.map(c => c.ent), cor: '93B4DC' }] });
    const g2 = grafRosca({ titulo: 'Entregues × pendentes', cat: { ref: `'DASHBOARD'!$O$8:$O$9`, vals: ['Entregues', 'Pendentes'] }, serie: { nome: 'Situação', nomeRef: `'DASHBOARD'!$B$8`, ref: `'DASHBOARD'!$P$8:$P$9`, vals: [nEnt, nPen] }, cores: [C.azul, 'D9A441'] });
    const g3 = grafBarras({ titulo: 'Entregues e pendentes por origem', dir: 'bar', grupo: 'stacked', cat: { ref: `'ANÁLISES'!$B$29:$B$${rT - 1}`, vals: ORIG }, series: [
      { nome: 'Entregues', nomeRef: `'ANÁLISES'!$D$28`, ref: `'ANÁLISES'!$D$29:$D$${rT - 1}`, vals: oCalc.map(c => c.e), cor: C.azul },
      { nome: 'Pendentes', nomeRef: `'ANÁLISES'!$E$28`, ref: `'ANÁLISES'!$E$29:$E$${rT - 1}`, vals: oCalc.map(c => c.p), cor: 'D9A441' }] });
    const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${ancora(0, [1, 12, 8, 29])}${ancora(1, [9, 12, 13, 29])}${ancora(2, [1, 31, 8, 49])}</xdr:wsDr>`;

    // ================================================================ PACOTE
    const folhas = [dsh, lan, rel, ana, cfg, aj];
    // DASHBOARD: notas referenciam B9 e D9 — os cards estão em B (total), D (entregues), F (pendentes)
    const definedNames = names.map(([n, f]) => `<definedName name="${n}">${xe(f)}</definedName>`).join('') +
      `<definedName name="_xlnm.Print_Titles" localSheetId="1">'LANÇAMENTOS'!$5:$5</definedName><definedName name="_xlnm.Print_Titles" localSheetId="2">'RELATÓRIOS'!$${rL + 1}:$${rL + 1}</definedName>`;
    const wb = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr defaultThemeVersion="164011"/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15000" activeTab="0"/></bookViews><sheets>${folhas.map((f, i) => `<sheet name="${xe(f.nome)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets><definedNames>${definedNames}</definedNames><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`;
    const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${folhas.map((f, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${folhas.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId${folhas.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`;
    const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const sheetXml = folhas.map(f => f.xml()); // gera shared strings
    const sstXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sstArr.length}" uniqueCount="${sstArr.length}">${sstArr.map(t => `<si><t xml:space="preserve">${xe(t)}</t></si>`).join('')}</sst>`;
    const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${folhas.map((f, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>${[1, 2, 3].map(i => `<Override PartName="/xl/charts/chart${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join('')}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
    const files = [
      ['[Content_Types].xml', ct],
      ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`],
      ['docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Controle de Avaria — Planilha corporativa</dc:title><dc:creator>${xe(usuario || 'Controle de Avaria')}</dc:creator><cp:lastModifiedBy>${xe(usuario || 'Controle de Avaria')}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${agora.toISOString().slice(0, 19)}Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${agora.toISOString().slice(0, 19)}Z</dcterms:modified></cp:coreProperties>`],
      ['docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application><Company>Controle de Avaria</Company></Properties>`],
      ['xl/workbook.xml', wb], ['xl/_rels/workbook.xml.rels', rels],
      ['xl/styles.xml', E.xml({ inteira: dx.tabInteira, cab: dx.tabCab, faixa: dx.tabFaixa })], ['xl/sharedStrings.xml', sstXml],
      ...sheetXml.map((x, i) => [`xl/worksheets/sheet${i + 1}.xml`, x]),
      ['xl/worksheets/_rels/sheet1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdD" Type="${REL}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`],
      ['xl/worksheets/_rels/sheet2.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdT" Type="${REL}/table" Target="../tables/table1.xml"/></Relationships>`],
      ['xl/tables/table1.xml', tabelaXml],
      ['xl/drawings/drawing1.xml', drawingXml],
      ['xl/drawings/_rels/drawing1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${[1, 2, 3].map(i => `<Relationship Id="rIdC${i}" Type="${REL}/chart" Target="../charts/chart${i}.xml"/>`).join('')}</Relationships>`],
      ['xl/charts/chart1.xml', g1], ['xl/charts/chart2.xml', g2], ['xl/charts/chart3.xml', g3],
    ];
    return zip(files);
  }
  G.AvariaPlanilha = { gerar };
})(typeof window !== 'undefined' ? window : globalThis);
