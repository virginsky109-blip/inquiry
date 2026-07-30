// 견적서 계산·출력 검증 — 실제 발행 견적서(몽골 울란바토르 DDP)를 그대로 재현해
// ※ 공개 저장소이므로 고객사명·담당자·연락처는 가상 값으로 대체했습니다. 금액은 원본 그대로입니다.
// 조건별 누적 TOTAL이 원본과 원 단위까지 일치하는지, 그리고 원가·마진이 고객용 출력에 새지 않는지 확인한다.
//
//   npm install jsdom      (최초 1회)
//   node test_quote.js
//
// 금액이 틀리면 견적이 틀리는 것이므로, 요율 계산·누적 합계·마진 분리 로직을 고칠 때 반드시 이 파일을 돌린다.
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

const HTML = path.join(__dirname, 'index.html');

// 실제 발행 견적서 (울란바토르 By Truck 시트) 원본 값
const C = (name, section, amount, currency, unit, vat, note) => ({
  id: name.slice(0, 14) + Math.random().toString(36).slice(2, 6),
  name, section, amount, cost: amount, costSet: false, min: 0, vat, currency, unit,
  note: note || '', auto: false, conIdx: null
});

const QUOTE = {
  customer: '(주)샘플무역', attn: '홍길동 팀장님', item: '-',
  subject: "인천-몽골 울란바토르 FCL(40') DDP 운임 견적서",
  direction: 'export', mode: 'FCL', incoterm: 'DDP', fx: 1464.30, date: '2026-07-30',
  polCountry: 'KOREA', pol: 'INCHEON', podCountry: '몽골', pod: 'ULAANBAATAR',
  terminal: '', place: '화성시 팔탄면',
  containers: [{ type: 'DRY', size: '40HC', qty: 4 }], container: '40HC × 4', qty: 4,
  cbm: 0, ton: 0, grossKg: 0, airCbm: 0,
  remarks: '2026년 8월 운임\n부가세 별도입니다',
  remarkIds: [], validUntil: '',
  company: { name: '샘플로지스틱스', staff: '해상수출팀, 김 담당', email: 'sales@example.com', tel: '02-0000-0000', fax: '02-0000-0001', address: '' },
  charges: [
    C('TERMINAL HANDLING CHARGE(항구 터미널핸드링비용)', '국내발생비용', 210000, 'KRW', 'CNTR', false),
    C('WHARFAGE(부두사용료)', '국내발생비용', 8690, 'KRW', 'CNTR', false),
    C('SEAL CHARGE(컨테이너 씰 비용)', '국내발생비용', 8000, 'KRW', 'CNTR', false),
    C('DOC FEE(비엘비)', '국내발생비용', 40000, 'KRW', 'BL', false),
    C('PSF(PORT SAFETY MANAGEMENT FEE)', '국내발생비용', 518, 'KRW', 'CNTR', false),
    C('CNTR TRUCKING CHARGE(컨테이너 트럭킹 비용)', '국내발생비용', 388200, 'KRW', 'CNTR', true, '컨테이너 작업지 : 화성시 팔탄면 => 인천항'),
    C('TRUCKING HANDING CHARGE', '국내발생비용', 20000, 'KRW', 'CNTR', true),
    C('INSURANCE FEE', '국내발생비용', 14000, 'KRW', 'INV VALUE', false, 'C.INVOICE VALUE X 110% X 보험요율'),
    C('CUSTOMS CLEARANCE', '국내발생비용', 12000, 'KRW', 'INV VALUE', true, 'C.INVOICE VALUE X 0.01%'),
    C('AFS(ADVANCE FILING SURCHARGE)', '국내발생비용', 30, 'USD', 'BL', false),
    C('OCEAN FREIGHT(해상운임)', '해상운임', 4900, 'USD', 'CNTR', false, '인천항 => 울란바토르 (중국 신강에서 By Truck)'),
    C('UB CY CHARGE', '해외발생비용', 350, 'USD', 'CNTR', false, 'FREE 3 DAYS // STORAGE: $6/day'),
    C('CUSTOMS CLEARANCE FEE', '해외발생비용', 100, 'USD', 'SET', false),
    C('CUSTOM DECLARATION', '해외발생비용', 0, 'USD', 'SET', false, '$8/sheet & $6/item(from 2nd item)'),
    C('DELIVERY CHARGE', '해외발생비용', 500, 'USD', 'CNTR', false),
    C('DUTY&TAX', '수입관세', 0, 'USD', 'SET', false, '(INVOICE VALUE+FREIGHT) X 15.5% / 요율 참고 요망'),
    C('FEE FOR COLLECT CHARGE', '수입관세', 450, 'USD', 'SET', false, '세금 대납 수수료')
  ]
};

// 원본 견적서에 인쇄된 값
const EXPECTED = {
  'FOB TOTAL': 2651561,
  'CIF TOTAL': 31351841,     // 보험료 항목이 있으므로 CFR이 아니라 CIF로 표기되어야 한다
  '몽골 발생 수입부대비용 소계': 5125050,
  'DAP TOTAL': 36476891,
  '몽골 수입 관세 소계': 658935,
  'DDP TOTAL': 37135826
};

let failures = 0;
const check = (label, actual, expected) => {
  const ok = String(actual) === String(expected);
  if (!ok) failures++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}\n        기대: ${expected}\n        실제: ${actual}`);
};

const dom = new JSDOM(fs.readFileSync(HTML, 'utf8'), {
  runScripts: 'dangerously',
  url: 'http://localhost/',
  beforeParse(win) {
    // 저장 견적을 미리 심어두면 앱 초기화 시 db.quotes로 읽힌다 (loadQuote가 클로저 밖에서 호출 불가하므로 UI 경로를 그대로 탄다)
    win.localStorage.setItem('fq-quotes', JSON.stringify([{ id: 'sample', name: '검증용 재현 견적', savedAt: '', data: QUOTE }]));
    win.localStorage.setItem('fq-company', JSON.stringify(QUOTE.company));
    win.fetch = () => Promise.reject(new Error('offline in test'));
    if (!win.structuredClone) win.structuredClone = o => JSON.parse(JSON.stringify(o));
  }
});

const { window } = dom;
window.addEventListener('load', () => {
  const $ = id => window.document.getElementById(id);

  // 저장 견적 불러오기 → 견적서 렌더링 (실제 사용자 경로)
  $('loadQuote').click();
  const btn = window.document.querySelector('#savedList [data-load="0"]');
  if (!btn) { console.log('FAIL  저장 견적을 불러오지 못했습니다.'); process.exit(1); }
  btn.click();

  const rows = [...window.document.querySelectorAll('#preview .qmain tr')]
    .map(tr => [...tr.children].map(td => td.textContent.trim()));
  const findRow = label => rows.find(r => r[0] === label);

  console.log('--- 조건별 누적 TOTAL (원본 견적서 대조) ---');
  for (const [label, amount] of Object.entries(EXPECTED)) {
    const row = findRow(label);
    check(label, row ? row[1] : '(행 없음)', 'KRW ' + amount.toLocaleString('ko-KR'));
  }

  console.log('\n--- 표기 규칙 ---');
  const html = $('preview').innerHTML;
  const nameCell = t => rows.find(r => r.some(c => c.startsWith(t)));
  check('과세 항목에 ** 표시 (내륙운송료)', /CNTR TRUCKING CHARGE\(컨테이너 트럭킹 비용\)\*\*/.test(html.replace(/<[^>]+>/g, '')), true);
  check('영세율 항목에는 ** 없음 (THC)', /TERMINAL HANDLING CHARGE\(항구 터미널핸드링비용\)\*\*/.test(html.replace(/<[^>]+>/g, '')), false);
  check('VAT 각주 삽입', html.includes('부가세(VAT 10%) 별도입니다'), true);
  check('견적서 표에 VAT 열 없음', /VAT\(10%\)/.test(html), false);
  check('금액 0원 항목은 AT COST 표기', !!nameCell('CUSTOM DECLARATION') && nameCell('CUSTOM DECLARATION').includes('AT COST'), true);
  check('섹션명 자동 생성 (도착국가 반영)', html.includes('몽골 발생 수입부대비용'), true);
  check('한국발생 수출부대비용 표기', html.includes('한국발생 수출부대비용'), true);
  check('담당자 블록', html.includes('sales@example.com') && html.includes('해상수출팀, 김 담당'), true);

  console.log('\n--- 마진 유출 방어 (고객용 출력에 원가 부재) ---');
  // 원가를 판매가와 다르게 바꾼 뒤 견적서를 다시 그려도 원가·마진 값이 나타나면 안 된다
  const costInput = window.document.querySelector('#chargeRows tr[data-row="0"] [data-k="cost"]');
  costInput.value = '133331';
  costInput.dispatchEvent(new window.Event('input', { bubbles: true }));
  $('buildQuote').click();
  const after = $('preview').innerHTML;
  check('견적서에 원가 미노출', after.includes('133331') || after.includes('133,331'), false);
  check('견적서에 마진 미노출', after.includes('706,676'), false); // 840,000 - 133,331×1... 어떤 형태로도 나오면 안 됨
  check('내부 화면에는 마진 표시', $('marginTotal').textContent !== 'KRW 0', true);

  // DUMP_HTML=경로 로 실행하면 렌더링된 견적서를 독립 HTML로 저장한다 (브라우저로 양식 육안 확인용)
  if (process.env.DUMP_HTML) {
    $('buildQuote').click(); // 원가 변경분 되돌린 최종 상태로 다시 렌더
    const css = [...window.document.querySelectorAll('style')].map(s => s.textContent).join('\n');
    fs.writeFileSync(process.env.DUMP_HTML, `<!doctype html><meta charset="utf-8"><style>${css}</style><body style="background:#f4f7fa;padding:24px"><article class="preview">${$('preview').innerHTML}</article>`);
    console.log('\n견적서 HTML 저장: ' + process.env.DUMP_HTML);
  }

  console.log(`\n${failures ? `실패 ${failures}건` : '전체 통과'}`);
  process.exit(failures ? 1 : 0);
});
