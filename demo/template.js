// 고객사 공수표 양식 채우기 (브라우저·node 공용, ExcelJS 통합문서를 받는다)
// - 올린 양식에서 성명 칸·날짜 칸·합계 칸을 찾아 정리 결과를 채운다. 서식·수식은 그대로 둔다.
// - 말로 설명한 칸 순서로 새 양식을 만든다.
// 금액·단가 칸은 건드리지 않는다.
(function (root) {
  const norm = s => String(s == null ? '' : s).replace(/\s+/g, '').trim();
  const NAME_RE = /^(성명|이름|근로자명|작업자|name|성명\(한글\))$/i;
  const TOTAL_RE = /^(합계|총공수|공수합계|월합계|공수계|계|총계)$/;
  const DAYS_RE = /^(출력일수|출역일수|근무일수|출근일수|일수)$/;
  const SITE_RE = /^(현장|현장명|site)$/i;   // '소속'은 업체명일 수 있어 채우지 않는다
  const JOB_RE = /^(직종|공종)$/;
  const NOTE_RE = /^(비고|특이사항)$/;
  const WORK_RE = /^(근무|주간|정상|공수)$/i;
  const OT_RE = /^(o\/?t|연장|잔업)$/i;
  const NIGHT_RE = /^야간$/;
  const STOP_RE = /^(합계|소계|일계|계|총계|총원)$/;

  // 셀 값 읽기: 수식이면 저장된 결과, 없으면 간단한 수식(DATE, 참조+1, 참조)을 직접 계산
  function cellValue(ws, cell, depth = 0) {
    let v = cell.value;
    if (v && typeof v === 'object' && !(v instanceof Date)) {
      if ('result' in v && v.result !== undefined && v.result !== null && !(typeof v.result === 'object' && v.result.error)) return v.result;
      if (v.richText) return v.richText.map(t => t.text).join('');
      if (v.formula || v.sharedFormula) return depth > 40 ? null : evalFormula(ws, cell, depth);
      if (v.text) return v.text;
      return null;
    }
    return v;
  }
  function refVal(ws, ref, depth) { return cellValue(ws, ws.getCell(ref.replace(/\$/g, '')), depth + 1); }
  function evalFormula(ws, cell, depth) {
    let f = cell.value.formula;
    if (!f && cell.value.sharedFormula) {   // 공유 수식: 기준 셀 수식을 행·열만큼 옮긴다
      const master = ws.getCell(cell.value.sharedFormula);
      const mf = master.value && master.value.formula; if (!mf) return null;
      const dr = cell.row - master.row, dc = cell.col - master.col;
      f = mf.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (m, a, c, b, r) => (a ? c : colName(colNum(c) + dc)) + (b ? r : String(+r + dr)));
    }
    f = String(f).replace(/^=/, '').replace(/^\+/, '').trim();
    let m;
    if ((m = f.match(/^DATE\(\s*([^,]+),\s*([^,]+),\s*([^)]+)\)$/i))) {
      const p = [m[1], m[2], m[3]].map(x => /^\$?[A-Z]{1,3}\$?\d+$/.test(x.trim()) ? Number(refVal(ws, x.trim(), depth)) : Number(x));
      if (p.some(x => !Number.isFinite(x))) return null;
      return new Date(Date.UTC(p[0], p[1] - 1, p[2]));
    }
    if ((m = f.match(/^\+?(\$?[A-Z]{1,3}\$?\d+)\s*\+\s*(\d+)$/))) {
      const base = refVal(ws, m[1], depth);
      if (base instanceof Date) return new Date(base.getTime() + Number(m[2]) * 86400000);
      if (typeof base === 'number') return base + Number(m[2]);
      return null;
    }
    if ((m = f.match(/^(\$?[A-Z]{1,3}\$?\d+)$/))) return refVal(ws, m[1], depth);
    return null;
  }
  function colNum(c) { return c.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0); }
  function colName(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
  const isFormula = cell => cell.value && typeof cell.value === 'object' && (cell.value.formula || cell.value.sharedFormula);
  const ownText = (ws, cell) => (cell.isMerged && cell.master && cell.master.address !== cell.address) ? '' : norm(cellValue(ws, cell));   // 병합된 칸은 첫 칸만 값으로 본다
  const formulaText = cell => { const v = cell.value; return v && typeof v === 'object' ? String(v.formula || '') : ''; };

  // 머리글 칸이 날짜(일)인지: Date, 1~31 숫자, "5", "5일", "11/5", "5(수)"
  function dayOf(v) {
    if (v instanceof Date) return { y: v.getUTCFullYear(), m: v.getUTCMonth() + 1, d: v.getUTCDate() };
    if (typeof v === 'number' && Number.isInteger(v) && v >= 1 && v <= 31) return { d: v };
    if (typeof v === 'number' && v > 40000 && v < 60000) { const dt = new Date(Date.UTC(1899, 11, 30) + v * 86400000); return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() }; }
    const s = norm(v); let m;
    if ((m = s.match(/^(\d{1,2})\/(\d{1,2})(\(.\))?$/))) return { m: +m[1], d: +m[2] };
    if ((m = s.match(/^(\d{1,2})일?(\(.\))?$/)) && +m[1] >= 1 && +m[1] <= 31) return { d: +m[1] };
    return null;
  }

  // 양식 구조 찾기
  function detectSheet(ws) {
    const maxR = Math.min(ws.rowCount, 40), maxC = Math.min(ws.columnCount, 200);
    for (let r = 1; r <= maxR; r++) {
      let nameCol = 0;   // 한글 '성명·이름' 칸을 영문 'Name' 칸보다 먼저
      for (const re of [/^(성명|이름|근로자명|작업자|성명\(한글\))$/, NAME_RE]) { for (let c = 1; c <= Math.min(maxC, 40) && !nameCol; c++) if (re.test(norm(cellValue(ws, ws.getCell(r, c))))) nameCol = c; if (nameCol) break; }
      if (!nameCol) continue;
      // 날짜 칸: 머리글 행과 아래 1행, 위 1행에서 찾는다
      let days = [], dayRow = 0;
      for (const rr of [r, r + 1, r - 1].filter(x => x >= 1)) {
        const found = [];
        for (let c = nameCol + 1; c <= maxC; c++) {
          const cell = ws.getCell(rr, c);
          if (cell.isMerged && cell.master && cell.master.address !== cell.address) continue;   // 병합된 날짜 칸(근무/O·T 위)은 한 번만
          const d = dayOf(cellValue(ws, cell)); if (!d) continue;
          const prev = found[found.length - 1];
          if (prev && !d.y) {   // 일 번호만 있으면 1씩 늘어나는 연속 칸만 (말일 뒤 1일은 다음 달)
            const next = prev.d + 1, wrap = prev.d >= 28 && d.d === 1;
            if (d.d !== next && !wrap) { if (found.length >= 5) break; found.length = 0; }
            d.seg = (prev ? prev.seg || 0 : 0) + (wrap ? 1 : 0);
          }
          found.push({ col: c, seg: 0, ...d });
        }
        if (found.length >= 5 && found.length > days.length) { days = found; dayRow = rr; }
      }
      if (!days.length) continue;
      // 날짜마다 두 칸(근무/O·T): 다음 행에 근무·O/T 라벨이 있으면
      const sub = r + 1 === dayRow ? r + 2 : r + 1;
      const labelAt = c => norm(cellValue(ws, ws.getCell(sub, c)));
      const night = days.slice(0, 3).some(d => NIGHT_RE.test(labelAt(d.col + 1)));   // 주간/야간 두 칸: 야간은 채우지 않는다
      const pairs = night || days.slice(0, 3).some(d => OT_RE.test(labelAt(d.col + 1)));   // 첫 칸 라벨이 비어 있는 양식도 있어 앞 3일을 본다
      const hasLabels = pairs || days.slice(0, 3).some(d => WORK_RE.test(labelAt(d.col)));
      const cols = { name: nameCol };
      for (let c = 1; c <= maxC; c++) {
        const h = norm(cellValue(ws, ws.getCell(r, c)));
        if (!cols.total && TOTAL_RE.test(h) && c > nameCol) cols.total = c;
        if (!cols.days && DAYS_RE.test(h)) cols.days = c;
        if (!cols.site && SITE_RE.test(h)) cols.site = c;
        if (!cols.job && JOB_RE.test(h)) cols.job = c;
        if (!cols.note && NOTE_RE.test(h)) cols.note = c;
      }
      // 합계 행(날짜 칸이 SUM 수식): 사람을 넣으면 안 되는 행
      const isTotalRow = rr => days.slice(0, 5).some(d => { const c = ws.getCell(rr, d.col); return isFormula(c) && (/SUM|SUBTOTAL/i.test(formulaText(c)) || (c.value.sharedFormula && /SUM/i.test(formulaText(ws.getCell(c.value.sharedFormula))))); });
      // 머리글 같은 행(요일·라벨이 적힌 행): 이름이 없고 날짜 칸 대부분이 글자
      const headerish = rr => !ownText(ws, ws.getCell(rr, nameCol)) && days.slice(0, 7).filter(d => { const v = cellValue(ws, ws.getCell(rr, d.col)); return typeof v === 'string' && v.trim() && !/^\d+(\.\d+)?$/.test(v.trim()); }).length >= 4;
      // 자료 시작 행: 머리글·날짜행·라벨행 다음
      let start = Math.max(r, dayRow) + 1;
      if (pairs || hasLabels) start = Math.max(start, sub + 1);
      while (start <= ws.rowCount && !ownText(ws, ws.getCell(start, nameCol)) && (isTotalRow(start) || headerish(start) || days.some(d => { const c = ws.getCell(start, d.col); return isFormula(c) && dayOf(cellValue(ws, c)); }))) start++;   // 요일·날짜 반복 행, 위쪽 합계 행 건너뛰기
      // 자료 끝 행: 합계·일계 행 전까지
      let end = start - 1;
      for (let rr = start; rr <= ws.rowCount; rr++) {
        const firstCells = [1, 2, 3, 4, nameCol].map(c => norm(cellValue(ws, ws.getCell(rr, c))));
        if (firstCells.some(t => STOP_RE.test(t)) || isTotalRow(rr)) break;
        end = rr;
      }
      return { ws, headerRow: r, dayRow, start, end, days, pairs, night, cols, headerish };
    }
    return null;
  }

  function pickSheet(wb, months) {
    const cands = [];
    wb.eachSheet(ws => { if (ws.state && ws.state !== 'visible') return; const d = detectSheet(ws); if (d) cands.push(d); });
    if (!cands.length) return null;
    if (months.length === 1) {   // 월별 시트(10월, 2026.10 등)면 자료의 달과 맞는 시트
      const [y, m] = months[0].split('-').map(Number);
      const byName = cands.find(c => new RegExp(`(^|[^0-9])0?${m}월|[.\\-_]0?${m}$|${y}[.\\-_]?0?${m}([^0-9]|$)`).test(c.ws.name));
      if (byName) return byName;
      const byHeader = cands.find(c => c.days.some(d => d.y === y && d.m === m));
      if (byHeader) return byHeader;
    }
    return cands[0];
  }

  // 정리 결과 → 하루 값
  function dayValues(row, unit, hoursPerMd) {
    if (unit === 'hours') return { work: row.base * hoursPerMd + row.otH, ot: row.otH, one: row.base * hoursPerMd + row.otH, workOnly: row.base * hoursPerMd };
    return { work: row.base, ot: row.otH, one: row.total, workOnly: row.base };
  }

  function fillTemplate(wb, result, opts = {}) {
    const unit = opts.unit || 'md', hoursPerMd = opts.hoursPerMd || 8, clear = opts.clear !== false;
    const months = [...new Set(result.dates.map(d => d.slice(0, 7)))];
    const t = pickSheet(wb, months);
    const report = { ok: false, sheet: null, placed: [], unplaced: [], skipped: [], warnings: [] };
    if (!t) { report.warnings.push('양식에서 성명 칸과 날짜 칸(1~31일)을 찾지 못했습니다.'); return report; }
    const { ws, days, pairs, night, cols } = t;
    report.sheet = ws.name; report.structure = { headerRow: t.headerRow, start: t.start, end: t.end, dayCount: days.length, pairs, cols: Object.fromEntries(Object.entries(cols).map(([k, c]) => [k, colName(c)])) };
    if (months.length > 1 && !days.some(d => d.m) && !days.some(d => d.seg > 0)) { report.warnings.push(`자료가 여러 달(${months.join(', ')})인데 양식 날짜 칸에 월이 없어 어느 달 칸인지 정할 수 없습니다. 한 달씩 나눠 정리한 뒤 채워 주세요.`); return report; }
    if (night) report.warnings.push('날짜마다 주간·야간 두 칸인 양식입니다. 야간 근무는 보고에서 구분하지 않으므로 주간 칸만 채웠습니다.');
    const [dy, dm] = (months[0] || '').split('-').map(Number);
    if (months.length === 1 && days.some(d => d.y) && !days.some(d => d.y === dy && d.m === dm)) { const h = days.find(d => d.y); report.warnings.push(`양식 날짜 칸이 ${h.y}-${String(h.m).padStart(2, '0')} 기준이라 ${months[0]} 자료를 넣을 칸이 없습니다. 양식의 기준 연·월 칸을 ${months[0]}로 바꿔 저장한 뒤 다시 올려 주세요.`); return report; }
    const wraps = days.some(x => x.seg > 0) && !days.some(x => x.m);   // 예: 5일~다음 달 4일
    const firstMonth = (() => { if (!wraps) return null; const ds = result.dates.slice().sort(); const f = ds[0]; const [fy, fm, fd] = f.split('-').map(Number); return fd >= days[0].d ? fy * 12 + fm - 1 : fy * 12 + fm - 2; })();
    const colFor = date => {
      const [y, m, d] = date.split('-').map(Number);
      if (wraps) { const seg = y * 12 + m - 1 - firstMonth; return (days.find(x => x.d === d && x.seg === seg) || {}).col; }
      return (days.find(x => x.d === d && (x.m ? x.m === m : true) && (x.y && months.length > 1 ? x.y === y : true)) || {}).col;
    };

    // 기존 값 지우기(지난달 파일을 양식으로 올린 경우): 날짜 칸의 수식 아닌 값만
    const rowsOfTable = []; for (let r = t.start; r <= Math.max(t.end, t.start); r++) rowsOfTable.push(r);
    if (clear) for (const r of rowsOfTable) { if (!ownText(ws, ws.getCell(r, cols.name)) || t.headerish(r)) continue; for (const d of days) for (const c of pairs ? [d.col, d.col + 1] : [d.col]) { const cell = ws.getCell(r, c); const v = cell.value; if (isFormula(cell) || v === null || v === '') continue; if (typeof v === 'number' || (typeof v === 'string' && v.trim().length <= 3)) cell.value = null; } }   // 이름 있는 줄의 숫자·짧은 표시(휴·결 등)만 지운다

    // 사람 행 찾기: 같은 이름 행, 없으면 빈 행
    const nameAt = r => ownText(ws, ws.getCell(r, cols.name));
    const used = new Set();
    const dupNames = new Set(); { const seen = new Set(); for (const r of rowsOfTable) { const n = nameAt(r); if (!n) continue; if (seen.has(n)) dupNames.add(n); seen.add(n); } }
    const dataNames = new Set(result.names.map(norm));
    // 자료에 없는 사람 줄: 일별 값을 지웠으니 고정 합계·출력일수도 비운다(수식이면 그대로)
    if (clear) for (const r of rowsOfTable) { const n = nameAt(r); if (!n || dataNames.has(n)) continue; for (const c of [cols.total, cols.days]) if (c && !isFormula(ws.getCell(r, c))) ws.getCell(r, c).value = null; }
    let written = 0;
    const findRow = n => { for (const r of rowsOfTable) if (nameAt(r) === norm(n)) { used.add(r); return r; } for (const r of rowsOfTable) if (!nameAt(r) && !used.has(r)) { used.add(r); ws.getCell(r, cols.name).value = n; return r; } return 0; };
    for (const n of result.names) {
      if (dupNames.has(norm(n))) { report.unplaced.push(n); report.warnings.push(`양식에 「${n}」이 두 줄 이상 있어 누구 줄인지 정할 수 없어 비워 두었습니다.`); continue; }
      const r = findRow(n);
      if (!r) { report.unplaced.push(n); continue; }
      let total = 0, count = 0;
      for (const row of result.rows.filter(x => x.name === n)) {
        const c = colFor(row.date);
        if (!c) { report.skipped.push(`${row.date} ${n}: 양식에 해당 날짜 칸 없음`); continue; }
        if (row.flag) { report.skipped.push(`${row.date} ${n}: 확인 필요라 비워 둠`); continue; }
        const v = dayValues(row, unit, hoursPerMd);
        const targets = pairs ? [[c, v.workOnly], [c + 1, night ? 0 : row.otH]] : [[c, v.one]];
        let blocked = false;
        for (const [cc, val] of targets) { if (!val) continue; const cell = ws.getCell(r, cc); if (isFormula(cell)) { blocked = true; continue; } cell.value = val; written++; }
        if (blocked) report.skipped.push(`${row.date} ${n}: 날짜 칸이 수식이라 덮어쓰지 않음`);
        if (night && row.otH) report.skipped.push(`${row.date} ${n}: 연장 ${row.otH}시간은 야간 칸에 넣지 않음`);
        total += v.one; if (v.one) count++;
      }
      if (cols.total && !isFormula(ws.getCell(r, cols.total))) ws.getCell(r, cols.total).value = total;
      if (cols.days && !isFormula(ws.getCell(r, cols.days))) ws.getCell(r, cols.days).value = count;
      if (cols.site && !norm(cellValue(ws, ws.getCell(r, cols.site)))) { const s = [...new Set(result.rows.filter(x => x.name === n && x.site).map(x => x.site))].join(', '); if (s) ws.getCell(r, cols.site).value = s; }
      report.placed.push(n);
    }
    if (report.unplaced.length) report.warnings.push(`양식에 빈 줄이 모자라 ${report.unplaced.length}명을 넣지 못했습니다: ${report.unplaced.join(', ')}. 양식 표 안(합계 행 위)에 빈 줄을 추가해 다시 올려 주세요.`);
    wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });   // 열 때 합계 수식 다시 계산
    report.written = written;
    report.otherSheets = []; wb.eachSheet(w => { if (w !== ws) report.otherSheets.push(w.name + (w.state && w.state !== 'visible' ? '(숨김)' : '')); });
    report.ok = report.placed.length > 0 && written > 0;
    if (report.placed.length && !written) report.warnings.push('채운 칸이 하나도 없습니다. 양식의 날짜·월이 자료와 맞는지 확인해 주세요.');
    return report;
  }

  // 말로 설명한 양식: "순번, 현장, 성명, 직종, 1~31, 합계, 출력일수, 비고" / 날짜 칸을 "1~31(근무/OT)"로 쓰면 두 칸
  function buildFromDescription(ExcelJS, desc, result, opts = {}) {
    const unit = opts.unit || 'md', hoursPerMd = opts.hoursPerMd || 8;
    const tokens = String(desc).split(/[,，\n]/).map(s => s.trim()).filter(Boolean);
    const months = [...new Set(result.dates.map(d => d.slice(0, 7)))];
    if (months.length > 1) return { error: `자료가 여러 달(${months.join(', ')})입니다. 한 달씩 나눠 정리한 뒤 만들어 주세요.` };
    const month = (result.dates[0] || '').slice(0, 7);
    const [y, m] = month.split('-').map(Number);
    const last = y ? new Date(y, m, 0).getDate() : 31;
    const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('공수표');
    const head = [], sub = [], plan = [];
    let pairs = false, hasDays = false;
    for (const tk of tokens) {
      const k = norm(tk);
      if (/^(1~31|1~말일|날짜|일자|일별|1일~말일|1~30)/.test(k)) {
        hasDays = true; pairs = /근무.*(ot|연장|야간)/i.test(k);
        for (let d = 1; d <= last; d++) { const dow = '일월화수목금토'[new Date(y, m - 1, d).getDay()]; if (pairs) { head.push(`${d}(${dow})`, ''); sub.push('근무', 'O/T'); plan.push({ day: d, part: 'work' }, { day: d, part: 'ot' }); } else { head.push(d); sub.push(dow); plan.push({ day: d }); } }
      } else { head.push(tk); sub.push(''); plan.push({ key: k }); }
    }
    if (!hasDays) return { error: '날짜 칸이 없습니다. "1~31"을 칸 목록에 넣어 주세요.' };
    if (!plan.some(p => p.key && NAME_RE.test(p.key))) return { error: '성명 칸이 없습니다. "성명"을 칸 목록에 넣어 주세요.' };
    ws.addRow([`${y}년 ${m}월 공수표`]); ws.addRow(head); ws.addRow(sub);
    result.names.forEach((n, i) => {
      const rs = result.rows.filter(x => x.name === n); let total = 0, count = 0;
      const cells = plan.map(p => {
        if (p.day) { const r = rs.find(x => +x.date.slice(8) === p.day && x.date.slice(0, 7) === month); if (!r || r.flag) return null; const v = dayValues(r, unit, hoursPerMd); if (p.part === 'ot') return r.otH || null; if (p.part === 'work') return v.workOnly || null; return v.one || null; }
        return null;
      });
      rs.filter(r => !r.flag && r.date.slice(0, 7) === month).forEach(r => { const v = dayValues(r, unit, hoursPerMd); total += v.one; if (v.one) count++; });
      plan.forEach((p, j) => {
        if (!p.key) return;
        if (NAME_RE.test(p.key)) cells[j] = n;
        else if (/^(순번|번호|no\.?|인원)$/i.test(p.key)) cells[j] = i + 1;
        else if (TOTAL_RE.test(p.key)) cells[j] = total;
        else if (DAYS_RE.test(p.key)) cells[j] = count;
        else if (SITE_RE.test(p.key)) cells[j] = [...new Set(rs.map(x => x.site).filter(Boolean))].join(', ') || null;
      });
      ws.addRow(cells);
    });
    ws.getRow(2).font = { bold: true }; ws.views = [{ state: 'frozen', ySplit: 3 }];
    return { wb, columns: head.length, pairs };
  }

  const api = { fillTemplate, buildFromDescription, detectSheet, dayOf };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Tmpl = api;
})(typeof window !== 'undefined' ? window : globalThis);
