// 양식 읽기·채우기를 화면과 따로 실행 (병합 칸이 아주 많은 파일이 화면을 멈추지 않게)
importScripts('exceljs.min.js', 'template.js');
onmessage = async e => {
  const { buf, result, opts, save } = e.data;
  try {
    const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf);
    const rep = Tmpl.fillTemplate(wb, result, opts);
    let out = null;
    if (save && rep.ok) { const u8 = new Uint8Array(await wb.xlsx.writeBuffer()); out = u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength); }   // 순수 ArrayBuffer로 넘긴다
    postMessage({ rep: JSON.parse(JSON.stringify(rep)), out }, out ? [out] : []);
  } catch (err) { postMessage({ error: err.message }); }
};
