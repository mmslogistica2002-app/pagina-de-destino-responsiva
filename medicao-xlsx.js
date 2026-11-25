/* Preenche os modelos oficiais XLSX da MMS sem reconstruir o layout.
   A edição é feita diretamente nas células do arquivo original, preservando estilos, imagens, larguras e bordas. */
(function (global) {
  'use strict';
  const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const NS_VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
  const NS_XML = 'http://www.w3.org/XML/1998/namespace';
  const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  const LAYOUTS = {
    dof13: { template: 'dof13', startRow: 6, firstCol: 'A', lastCol: 'I', filename: 'DOF_FROTA_1_E_3' },
    tfmcDof2: { template: 'tfmcDof2', startRow: 6, firstCol: 'A', lastCol: 'J', filename: 'TFMC_DOF_FROTA_2' },
    constellation: { template: 'constellation', startRow: 6, firstCol: 'A', lastCol: 'I', filename: 'CONSTELLATION' },
    transporte: { template: 'transporte', startRow: 7, firstCol: 'B', lastCol: 'K', filename: 'TRANSPORTES' }
  };

  function parseXml(text, label) {
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    const parseErrors = doc.getElementsByTagName('parsererror');
    if (parseErrors && parseErrors.length) throw new Error('Não foi possível ler o XML do modelo (' + label + ').');
    return doc;
  }
  function serializeXml(doc) { return new XMLSerializer().serializeToString(doc); }
  function directChildren(node, ns, localName) {
    return Array.from(node.childNodes || []).filter(n => n.nodeType === 1 && n.namespaceURI === ns && n.localName === localName);
  }
  function elementText(node) { return node ? String(node.textContent || '') : ''; }
  function columnNumber(col) {
    let n = 0; for (const ch of String(col).toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n;
  }
  function columnLetters(n) {
    let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s;
  }
  function splitCellRef(ref) {
    const m = String(ref || '').match(/^([A-Z]+)(\d+)$/i);
    return m ? { col: m[1].toUpperCase(), row: Number(m[2]) } : null;
  }
  function findRow(sheetData, rowNum, doc) {
    const rows = directChildren(sheetData, NS_MAIN, 'row');
    let row = rows.find(r => Number(r.getAttribute('r')) === rowNum);
    if (row) return row;
    row = doc.createElementNS(NS_MAIN, 'row');
    row.setAttribute('r', String(rowNum));
    const next = rows.find(r => Number(r.getAttribute('r')) > rowNum);
    if (next) sheetData.insertBefore(row, next); else sheetData.appendChild(row);
    return row;
  }
  function getCell(row, col, rowNum, doc) {
    const ref = col + rowNum;
    const cells = directChildren(row, NS_MAIN, 'c');
    let cell = cells.find(c => c.getAttribute('r') === ref);
    if (cell) return cell;
    cell = doc.createElementNS(NS_MAIN, 'c'); cell.setAttribute('r', ref);
    const idx = columnNumber(col);
    const next = cells.find(c => { const p = splitCellRef(c.getAttribute('r')); return p && columnNumber(p.col) > idx; });
    if (next) row.insertBefore(cell, next); else row.appendChild(cell);
    return cell;
  }
  function clearCell(cell) {
    while (cell.firstChild) cell.removeChild(cell.firstChild);
    cell.removeAttribute('t');
  }
  function setCellText(doc, cell, value) {
    clearCell(cell); cell.setAttribute('t', 'inlineStr');
    const is = doc.createElementNS(NS_MAIN, 'is');
    const t = doc.createElementNS(NS_MAIN, 't');
    t.setAttributeNS(NS_XML, 'xml:space', 'preserve');
    t.appendChild(doc.createTextNode(String(value)));
    is.appendChild(t); cell.appendChild(is);
  }
  function setCellNumber(doc, cell, value) {
    clearCell(cell);
    const v = doc.createElementNS(NS_MAIN, 'v');
    v.appendChild(doc.createTextNode(String(value)));
    cell.appendChild(v);
  }
  function parseDatePt(value) {
    const s = String(value || '').trim();
    let m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (!m) m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/) && [null, s.slice(8,10), s.slice(5,7), s.slice(0,4)];
    if (!m) return null;
    const d = Number(m[1]), mo = Number(m[2]), y = Number(m[3]);
    if (!y || mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    const stamp = Date.UTC(y, mo - 1, d);
    const check = new Date(stamp);
    if (check.getUTCFullYear() !== y || check.getUTCMonth() !== mo - 1 || check.getUTCDate() !== d) return null;
    return Math.round(stamp / 86400000) + 25569;
  }
  function parseTime(value) {
    const s = String(value || '').trim().toUpperCase().replace(/\s/g, '').replace(/HRS?$/,'');
    const m = s.match(/^(\d{1,2})[:H.](\d{2})(?::(\d{2}))?$/);
    if (!m) return null;
    const h = Number(m[1]), min = Number(m[2]), sec = Number(m[3] || 0);
    if (h > 23 || min > 59 || sec > 59) return null;
    return (h * 3600 + min * 60 + sec) / 86400;
  }
  function isBlank(v) { return v === null || v === undefined || String(v).trim() === ''; }
  function typedValue(doc, cell, value, kind) {
    if (isBlank(value)) { clearCell(cell); return; }
    if (kind === 'date') {
      const n = parseDatePt(value);
      if (n !== null) setCellNumber(doc, cell, n); else setCellText(doc, cell, value);
      return;
    }
    if (kind === 'time') {
      const n = parseTime(value);
      if (n !== null) setCellNumber(doc, cell, n); else setCellText(doc, cell, value);
      return;
    }
    if (kind === 'number') {
      const n = Number(value);
      if (Number.isFinite(n)) setCellNumber(doc, cell, n); else setCellText(doc, cell, value);
      return;
    }
    setCellText(doc, cell, value);
  }
  function cellsForRow(tipo, row) {
    if (tipo === 'dof13') return [
      ['A', row.numero, 'number'], ['B', row.data, 'date'], ['C', row.unidade, 'text'], ['D', row.horario, 'time'],
      ['E', 'DOF', 'text'], ['F', row.qtd || 1, 'number'], ['G', row.preposto, 'text'], ['H', row.local, 'text'], ['I', row.obs, 'text']
    ];
    if (tipo === 'tfmcDof2') return [
      ['A', row.numero, 'number'], ['B', row.data, 'date'], ['C', row.unidade, 'text'], ['D', row.tipo, 'text'], ['E', row.horario, 'time'],
      ['F', row.empresa, 'text'], ['G', row.qtd || 1, 'number'], ['H', row.preposto, 'text'], ['I', row.local, 'text'], ['J', row.obs, 'text']
    ];
    if (tipo === 'constellation') return [
      ['A', row.numero, 'number'], ['B', row.data, 'date'], ['C', row.unidade, 'text'], ['D', row.horarioTT, 'time'],
      ['E', row.inicio, 'time'], ['F', row.fim, 'time'], ['G', row.qtd || 1, 'number'], ['H', row.preposto, 'text'], ['I', row.local, 'text']
    ];
    return [
      ['B', row.data, 'date'], ['C', row.diaSemana, 'text'], ['D', row.horario, 'time'], ['E', row.preposto, 'text'],
      ['F', 'RESIDÊNCIA', 'text'], ['G', row.destino, 'text'], ['H', row.empresa, 'text'], ['I', row.embarcacao, 'text'],
      ['J', row.localAtendimento, 'text'], ['K', row.obs, 'text']
    ];
  }
  function cleanDataRows(sheetDoc, startRow) {
    const sheetData = sheetDoc.getElementsByTagNameNS(NS_MAIN, 'sheetData')[0];
    if (!sheetData) throw new Error('O modelo não contém uma área de dados.');
    const rows = directChildren(sheetData, NS_MAIN, 'row');
    for (const row of rows) {
      const rn = Number(row.getAttribute('r'));
      if (rn < startRow) continue;
      directChildren(row, NS_MAIN, 'c').forEach(clearCell);
    }
    return sheetData;
  }
  function ensureRows(sheetDoc, sheetData, startRow, count) {
    const existing = directChildren(sheetData, NS_MAIN, 'row');
    const dataRows = existing.filter(r => Number(r.getAttribute('r')) >= startRow);
    const maxExisting = dataRows.reduce((m, r) => Math.max(m, Number(r.getAttribute('r'))), startRow - 1);
    if (count <= dataRows.length) return maxExisting;
    const templateRow = dataRows[dataRows.length - 1];
    if (!templateRow) throw new Error('O modelo não possui linha formatada para preenchimento.');
    let last = maxExisting;
    while (last < startRow + count - 1) {
      last += 1;
      const clone = templateRow.cloneNode(true);
      clone.setAttribute('r', String(last));
      directChildren(clone, NS_MAIN, 'c').forEach(cell => {
        const p = splitCellRef(cell.getAttribute('r'));
        if (p) cell.setAttribute('r', p.col + last);
        clearCell(cell);
      });
      sheetData.appendChild(clone);
    }
    return last;
  }
  function fillSheet(zip, tipo, rows) {
    const layout = LAYOUTS[tipo];
    const path = 'xl/worksheets/sheet1.xml';
    return zip.file(path).async('string').then(text => {
      const doc = parseXml(text, path);
      const sheetData = cleanDataRows(doc, layout.startRow);
      const lastRow = ensureRows(doc, sheetData, layout.startRow, rows.length);
      rows.forEach((row, i) => {
        const rowNum = layout.startRow + i;
        const rowEl = findRow(sheetData, rowNum, doc);
        cellsForRow(tipo, row).forEach(([col, value, kind]) => typedValue(doc, getCell(rowEl, col, rowNum, doc), value, kind));
      });
      const dim = doc.getElementsByTagNameNS(NS_MAIN, 'dimension')[0];
      if (dim) {
        const lastCol = layout.lastCol;
        const endRow = Math.max(layout.startRow, rows.length ? layout.startRow + rows.length - 1 : layout.startRow - 1);
        dim.setAttribute('ref', 'A1:' + lastCol + endRow);
      }
      zip.file(path, serializeXml(doc));
      return lastRow;
    });
  }
  function renameFirstSheet(zip, newName) {
    return zip.file('xl/workbook.xml').async('string').then(async text => {
      const doc = parseXml(text, 'xl/workbook.xml');
      const sheets = Array.from(doc.getElementsByTagNameNS(NS_MAIN, 'sheet'));
      if (!sheets.length) throw new Error('O modelo não possui aba de medição.');
      sheets[0].setAttribute('name', newName);
      const sheetsParent = sheets[0].parentNode;
      sheets.slice(1).forEach(s => sheetsParent.removeChild(s));
      const defined = doc.getElementsByTagNameNS(NS_MAIN, 'definedNames');
      Array.from(defined).forEach(n => n.parentNode.removeChild(n));
      zip.file('xl/workbook.xml', serializeXml(doc));

      const relPath = 'xl/_rels/workbook.xml.rels';
      const relText = await zip.file(relPath).async('string');
      const relDoc = parseXml(relText, relPath);
      Array.from(relDoc.getElementsByTagNameNS(NS_PKG, 'Relationship')).forEach(rel => {
        const type = rel.getAttribute('Type') || '';
        const target = rel.getAttribute('Target') || '';
        if (type.endsWith('/worksheet') && target !== 'worksheets/sheet1.xml') rel.parentNode.removeChild(rel);
      });
      zip.file(relPath, serializeXml(relDoc));

      const ctPath = '[Content_Types].xml';
      const ctText = await zip.file(ctPath).async('string');
      const ctDoc = parseXml(ctText, ctPath);
      Array.from(ctDoc.getElementsByTagNameNS(NS_CT, 'Override')).forEach(item => {
        const part = item.getAttribute('PartName') || '';
        if (/^\/xl\/worksheets\/sheet([2-9]|\d{2,})\.xml$/.test(part) || /^\/xl\/drawings\/drawing([2-9]|\d{2,})\.xml$/.test(part)) item.parentNode.removeChild(item);
      });
      zip.file(ctPath, serializeXml(ctDoc));

      const names = Object.keys(zip.files);
      names.forEach(name => {
        if (/^xl\/worksheets\/sheet([2-9]|\d{2,})\.xml$/.test(name) || /^xl\/worksheets\/_rels\/sheet([2-9]|\d{2,})\.xml\.rels$/.test(name) || /^xl\/drawings\/drawing([2-9]|\d{2,})\.xml$/.test(name) || /^xl\/drawings\/_rels\/drawing([2-9]|\d{2,})\.xml\.rels$/.test(name)) zip.remove(name);
      });

      const appPath = 'docProps/app.xml';
      if (zip.file(appPath)) {
        const appDoc = parseXml(await zip.file(appPath).async('string'), appPath);
        Array.from(appDoc.getElementsByTagNameNS(NS_VT, 'vector')).forEach(vec => {
          const parent = vec.parentNode;
          if (parent && parent.localName === 'TitlesOfParts') {
            const children = Array.from(vec.childNodes).filter(n => n.nodeType === 1);
            children.slice(1).forEach(n => vec.removeChild(n));
            vec.setAttribute('size', '1');
            const first = children[0];
            if (first) first.textContent = newName;
          }
          if (parent && parent.localName === 'HeadingPairs') {
            const i4 = vec.getElementsByTagNameNS(NS_VT, 'i4')[0];
            if (i4) i4.textContent = '1';
          }
        });
        zip.file(appPath, serializeXml(appDoc));
      }
    });
  }
  function filename(tipo, inicio, fim) {
    const slug = LAYOUTS[tipo].filename;
    return 'MMS_MEDICAO_' + slug + '_' + inicio + '_a_' + fim + '.xlsx';
  }
  async function gerarArquivo(resultado) {
    if (!resultado || !LAYOUTS[resultado.tipo]) throw new Error('Tipo de medição inválido.');
    if (!global.JSZip) throw new Error('A biblioteca de planilhas não foi carregada. Reabra o aplicativo e tente novamente.');
    const templates = global.MMS_MEDICAO_TEMPLATES || {};
    const base64 = templates[LAYOUTS[resultado.tipo].template];
    if (!base64) throw new Error('O modelo oficial desta medição não foi encontrado dentro do aplicativo.');
    const zip = await global.JSZip.loadAsync(base64, { base64: true });
    if (resultado.tipo === 'dof13') await renameFirstSheet(zip, 'MEDIÇÃO');
    else await renameFirstSheet(zip, 'MEDIÇÃO');
    await fillSheet(zip, resultado.tipo, Array.isArray(resultado.rows) ? resultado.rows : []);
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    return { bytes, filename: filename(resultado.tipo, resultado.inicio, resultado.fim), mime: MIME };
  }

  global.MMSMedicaoExcel = { gerarArquivo };
})(globalThis);
