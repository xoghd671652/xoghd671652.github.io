// 양식을 건드리지 않고 값만 써 넣기
// 1) SheetJS(읽기 전용, 빠름)로 읽어 ExcelJS 비슷한 모양으로 감싼 뒤 template.js 의 칸 찾기·채우기 규칙을 그대로 쓴다.
//    채우기 결과는 파일에 바로 쓰지 않고 "어느 칸에 무슨 값" 목록으로 모은다.
// 2) 확인을 받은 뒤 원본 파일(zip) 안의 해당 시트 XML에서 그 칸의 값만 바꾼다.
//    서식·병합·수식·다른 시트·그림은 원본 그대로 남는다.
(function (root) {
  const addr = (r, c) => { let s = '', n = c; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s + r; };
  const parseAddr = a => { const m = String(a).replace(/\$/g, '').match(/^([A-Z]+)(\d+)$/); if (!m) return null; return { r: +m[2], c: m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) }; };

  // SheetJS 워크북 → template.js 가 쓰는 최소 API (getCell, rowCount, columnCount, name, state, eachSheet)
  function wrap(XLSX, swb) {
    const writes = new Map();   // "시트!A1" → 값(null = 지움)
    const isDateFmt = z => z && XLSX.SSF && XLSX.SSF.is_date(z);
    const serialToDate = v => new Date(Date.UTC(1899, 11, 30) + Math.round(v * 86400000));
    const sheets = swb.SheetNames.map((name, i) => {
      const s = swb.Sheets[name];
      const range = s['!ref'] ? XLSX.utils.decode_range(s['!ref']) : { e: { r: -1, c: -1 } };
      const masterOf = new Map();   // 병합된 칸(첫 칸 제외) → 첫 칸 주소
      for (const m of s['!merges'] || []) for (let r = m.s.r; r <= m.e.r; r++) for (let c = m.s.c; c <= m.e.c; c++) if (r !== m.s.r || c !== m.s.c) masterOf.set(addr(r + 1, c + 1), addr(m.s.r + 1, m.s.c + 1));
      const hidden = ((swb.Workbook && swb.Workbook.Sheets) || [])[i];
      const ws = {
        name, state: hidden && hidden.Hidden ? 'hidden' : 'visible',
        rowCount: range.e.r + 1, columnCount: range.e.c + 1,
        getCell(r, c) {
          if (c === undefined) { const p = parseAddr(r); r = p.r; c = p.c; }
          const a = addr(r, c), key = name + '!' + a;
          const master = masterOf.get(a);
          const read = () => {
            if (writes.has(key)) return writes.get(key);
            const sc = s[a]; if (!sc) return null;
            let v = sc.v;
            if (sc.t === 'n' && isDateFmt(sc.z)) v = serialToDate(sc.v);
            if (sc.t === 'e' || sc.t === 'z') v = sc.t === 'z' ? undefined : null;
            if (sc.f) return { formula: sc.f, result: v === undefined ? undefined : v };
            return v === undefined ? null : v;
          };
          return {
            row: r, col: c, address: a,
            isMerged: !!master || (s['!merges'] || []).some(m => m.s.r === r - 1 && m.s.c === c - 1 && (m.e.r > m.s.r || m.e.c > m.s.c)),
            master: master ? { address: master } : { address: a },
            get value() { if (master) { const p = parseAddr(master); return ws.getCell(p.r, p.c).value; } return read(); },
            set value(v) { writes.set(key, v); },
          };
        },
      };
      return ws;
    });
    return { wb: { eachSheet(fn) { sheets.forEach(fn); }, calcProperties: {} }, writes };
  }

  // 미리보기: 칸 찾기·채우기 규칙을 돌려 보고 쓸 목록을 만든다(파일은 아직 그대로)
  function plan(deps, buf, result, opts) {
    const { XLSX, Tmpl } = deps;
    const swb = XLSX.read(buf, { type: 'array', cellFormula: true, cellNF: true, cellDates: false, sheetStubs: true });   // 결과값 없는 수식 칸도 읽기
    const { wb, writes } = wrap(XLSX, swb);
    const rep = Tmpl.fillTemplate(wb, result, opts);
    const own = [...writes.entries()].filter(([k]) => k.startsWith(rep.sheet + '!'));
    const original = k => { const a = k.slice(rep.sheet.length + 1); const sc = swb.Sheets[rep.sheet][a]; return sc ? sc.v : undefined; };
    rep.plan = {
      fill: own.filter(([, v]) => v !== null && v !== '').length,
      clear: own.filter(([k, v]) => (v === null || v === '') && original(k) !== undefined && original(k) !== '').length,
    };
    return { rep, writes: own.map(([k, v]) => [k.slice(rep.sheet.length + 1), v]), sheet: rep.sheet };
  }

  // 시트 XML 에 값만 쓰기
  async function apply(deps, buf, sheetName, writes) {
    const { JSZip, DOMParser, XMLSerializer } = deps;
    const arr = x => Array.prototype.slice.call(x);   // 브라우저·xmldom 공통
    const zip = await JSZip.loadAsync(buf);
    const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const parse = async p => new DOMParser().parseFromString(await zip.file(p).async('string'), 'application/xml');
    const wbXml = await parse('xl/workbook.xml');
    const sheetEl = arr(wbXml.getElementsByTagNameNS(NS_MAIN, 'sheet')).find(e => e.getAttribute('name') === sheetName);
    if (!sheetEl) throw new Error('시트를 찾지 못했습니다: ' + sheetName);
    const rid = sheetEl.getAttributeNS(NS_R, 'id') || sheetEl.getAttribute('r:id');
    const rels = await parse('xl/_rels/workbook.xml.rels');
    const rel = arr(rels.getElementsByTagName('Relationship')).find(e => e.getAttribute('Id') === rid);
    let target = rel.getAttribute('Target'); target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
    const doc = await parse(target);
    const sheetData = doc.getElementsByTagNameNS(NS_MAIN, 'sheetData')[0];
    const rowsByNum = new Map(); for (const r of arr(sheetData.getElementsByTagNameNS(NS_MAIN, 'row'))) rowsByNum.set(+r.getAttribute('r'), r);
    const el = (tag) => doc.createElementNS(NS_MAIN, tag);
    const colOf = a => parseAddr(a).c;
    for (const [a, v] of writes) {
      const { r } = parseAddr(a);
      let row = rowsByNum.get(r);
      if (!row) {   // 행이 없으면 순서에 맞게 새로 만든다
        if (v === null || v === '') continue;
        row = el('row'); row.setAttribute('r', String(r));
        const after = Array.from(rowsByNum.keys()).filter(n => n > r).sort((x, y) => x - y)[0];
        sheetData.insertBefore(row, after ? rowsByNum.get(after) : null); rowsByNum.set(r, row);
      }
      const cells = arr(row.getElementsByTagNameNS(NS_MAIN, 'c'));
      let c = cells.find(x => x.getAttribute('r') === a);
      if (!c) {
        if (v === null || v === '') continue;
        c = el('c'); c.setAttribute('r', a);
        const next = cells.find(x => colOf(x.getAttribute('r')) > colOf(a));
        row.insertBefore(c, next || null);
        if (!next) { const style = row.getAttribute('s'); if (style && row.getAttribute('customFormat') === '1') c.setAttribute('s', style); }
      }
      if (c.getElementsByTagNameNS(NS_MAIN, 'f').length) continue;   // 수식 칸은 건드리지 않는다
      for (const ch of arr(c.childNodes)) c.removeChild(ch);
      c.removeAttribute('t');
      if (v === null || v === '') continue;
      if (typeof v === 'number') { const ve = el('v'); ve.appendChild(doc.createTextNode(String(v))); c.appendChild(ve); }
      else { c.setAttribute('t', 'inlineStr'); const is = el('is'), t = el('t'); t.appendChild(doc.createTextNode(String(v))); is.appendChild(t); c.appendChild(is); }
    }
    zip.file(target, new XMLSerializer().serializeToString(doc));
    // 열 때 합계 수식을 다시 계산하도록
    let calc = wbXml.getElementsByTagNameNS(NS_MAIN, 'calcPr')[0];
    if (!calc) {   // 스키마 순서: calcPr 은 아래 요소들보다 앞
      calc = wbXml.createElementNS(NS_MAIN, 'calcPr');
      const after = arr(wbXml.documentElement.childNodes).find(n => n.nodeType === 1 && ['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'].includes(n.localName));
      wbXml.documentElement.insertBefore(calc, after || null);
    }
    calc.setAttribute('fullCalcOnLoad', '1');
    zip.file('xl/workbook.xml', new XMLSerializer().serializeToString(wbXml));
    return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
  }

  const api = { plan, apply, wrap };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.XFill = api;
})(typeof window !== 'undefined' ? window : globalThis);
