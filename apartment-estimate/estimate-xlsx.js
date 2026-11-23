const encoder = new TextEncoder();
const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const spreadsheetNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const officeRelationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageRelationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
const xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function escapeXml(value) {
  return String(value).slice(0, 32767)
    .replace(/[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/gu, '')
    .replace(/_x[\da-f]{4}_/gi, match => '_x005F_' + match.slice(1))
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) {
      crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

// XLSX is a ZIP archive. These small XML parts can use ZIP's store method.
function zip(files) {
  const localParts = [];
  const directoryParts = [];
  let offset = 0;
  let directorySize = 0;
  for (const [path, xml] of files) {
    const name = encoder.encode(path);
    const data = encoder.encode(xmlHeader + xml);
    const crc = crc32(data);
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint16(12, 33, true);
    local.setUint32(14, crc, true);
    local.setUint32(18, data.length, true);
    local.setUint32(22, data.length, true);
    local.setUint16(26, name.length, true);
    localParts.push(local.buffer, name, data);

    const central = new DataView(new ArrayBuffer(46));
    central.setUint32(0, 0x02014b50, true);
    central.setUint16(4, 20, true);
    central.setUint16(6, 20, true);
    central.setUint16(14, 33, true);
    central.setUint32(16, crc, true);
    central.setUint32(20, data.length, true);
    central.setUint32(24, data.length, true);
    central.setUint16(28, name.length, true);
    central.setUint32(42, offset, true);
    directoryParts.push(central.buffer, name);
    offset += local.byteLength + name.length + data.length;
    directorySize += central.byteLength + name.length;
  }
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, files.length, true);
  end.setUint16(10, files.length, true);
  end.setUint32(12, directorySize, true);
  end.setUint32(16, offset, true);
  return new Blob([...localParts, ...directoryParts, end.buffer], { type: mimeType });
}

export function makeWorkbook(rows) {
  const sheetRows = rows.map(({ cells, heading, date }, index) => {
    const number = index + 1;
    const row = cells.map((value, column) => {
      const address = 'ABCDEF'[column] + number;
      if (Number.isFinite(value)) {
        const style = date && column === 1 ? 3 : column === 4 ? 1 : 0;
        return '<c r="' + address + '" s="' + style + '"><v>' + value + '</v></c>';
      }
      if (typeof value !== 'string' || !value) return '';
      return '<c r="' + address + '" t="inlineStr" s="' + (heading ? 2 : 0) + '">' +
        '<is><t xml:space="preserve">' + escapeXml(value) + '</t></is></c>';
    }).join('');
    return '<row r="' + number + '">' + row + '</row>';
  }).join('');
  const columns = [22, 40, 13, 10, 18, 57].map((width, index) =>
    '<col min="' + (index + 1) + '" max="' + (index + 1) + '" width="' + width + '" customWidth="1"/>'
  ).join('');
  const sheet = '<worksheet xmlns="' + spreadsheetNs + '"><cols>' + columns +
    '</cols><sheetData>' + sheetRows + '</sheetData></worksheet>';
  const styles = '<styleSheet xmlns="' + spreadsheetNs + '">' +
    '<fonts count="2"><font><sz val="11"/><name val="맑은 고딕"/></font>' +
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="맑은 고딕"/></font></fonts>' +
    '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF345C46"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
    '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  return zip([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'],
    ['_rels/.rels', '<Relationships xmlns="' + packageRelationships + '"><Relationship Id="rId1" Type="' +
      officeRelationships + '/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
    ['xl/workbook.xml', '<workbook xmlns="' + spreadsheetNs + '" xmlns:r="' + officeRelationships +
      '"><sheets><sheet name="예상 견적" sheetId="1" r:id="rId1"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels', '<Relationships xmlns="' + packageRelationships + '">' +
      '<Relationship Id="rId1" Type="' + officeRelationships + '/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="' + officeRelationships + '/styles" Target="styles.xml"/></Relationships>'],
    ['xl/styles.xml', styles],
    ['xl/worksheets/sheet1.xml', sheet],
  ]);
}

function makeEstimateRows(dateSerial) {
  const vatIncluded = $('vat').checked;
  const rows = [
    { cells: ['MADE 아파트 예상 견적'], heading: true },
    { cells: ['작성일', dateSerial], date: true },
    { cells: ['공급평수', area, '평'] },
    { cells: [] },
    { cells: ['견적 요약', '', '', '', '금액 (원)', '비고'], heading: true },
    { cells: ['선택 공사비', '', '', '', result.direct] },
    { cells: ['현장 공통 작업비', '', '', '', result.commonTotal, $('includeCommon').checked ? '포함' : '제외'] },
    { cells: ['간접공사비', '', '', '', result.fees] },
    { cells: ['공급가', '', '', '', result.supply] },
    { cells: ['부가세', '', '', '', result.tax, vatIncluded ? '합계에 포함' : '합계와 별도'] },
    { cells: [result.pending.length ? '확인된 항목 예상 소계' : '예상 합계', '', '', '', result.total,
      result.midway?.temporary ? '미드웨이 중간값 적용' : ''] },
    { cells: [] },
    { cells: ['공종', '선택 항목', '산정 수량', '단위', '견적금액 (원)', '비고'], heading: true },
  ];
  for (const product of products.filter(product => choices.has(product.id))) {
    const quote = item(product.id);
    const needsReview = !quote || pending(product.id);
    const note = product.id === 'midway' ? midwayText() : needsReview ? '금액 미포함 · 별도 확인' : '';
    const quantities = needsReview ? [] : quote.quantities ?? [];
    const [first, ...remaining] = quantities;
    rows.push({ cells: [product.group, product.label, first?.qty, first?.unit,
      needsReview ? null : quote.total, note] });
    for (const quantity of remaining) {
      rows.push({ cells: ['', '', quantity.qty, quantity.unit] });
    }
  }
  rows.push({ cells: [] }, { cells: ['확인사항'], heading: true });
  for (const note of [...result.pending.map(label => '금액 미포함 · ' + label), ...warnings(),
    '실측·제품 사양·현장 조건에 따라 최종 견적이 달라질 수 있습니다.']) {
    rows.push({ cells: ['', note] });
  }
  const memo = $('memo').value.trim();
  if (memo) {
    memo.split(/\r?\n/).forEach((line, index) => rows.push({ cells: [index ? '' : '추가 요청사항', line] }));
  }
  return rows;
}

function addDownloadButton() {
  const copy = document.getElementById('copy');
  const print = document.getElementById('print');
  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'secondary full';
  button.textContent = '엑셀 견적서 다운로드';
  button.disabled = true;
  copy.after(button);

  const update = () => {
    button.disabled = print.disabled || typeof result === 'undefined' || !result || choices.size === 0;
  };
  new MutationObserver(update).observe(print, { attributeFilter: ['disabled'] });
  update();
  button.addEventListener('click', () => {
    update();
    if (button.disabled) return;
    try {
      const date = new Intl.DateTimeFormat('sv-SE', { timeZone: 'Asia/Seoul',
        year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
      const dateSerial = Date.parse(date + 'T00:00:00Z') / 86400000 + 25569;
      const url = URL.createObjectURL(makeWorkbook(makeEstimateRows(dateSerial)));
      const link = document.createElement('a');
      link.href = url;
      link.download = 'MADE_아파트_예상견적_' + date + '.xlsx';
      document.body.append(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30000);
      $('copyStatus').textContent = '';
    } catch (error) {
      console.error('견적서 Excel 내보내기 오류', error);
      $('copyStatus').textContent = '엑셀 파일을 만들지 못했습니다. 다시 시도해주세요.';
    }
  });
}

if (typeof document !== 'undefined') {
  document.addEventListener('DOMContentLoaded', addDownloadButton, { once: true });
}
