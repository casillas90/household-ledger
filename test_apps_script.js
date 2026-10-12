// google_apps_script.js 단위 테스트 (가짜 스프레드시트로 실행)
// 실행: osascript -l JavaScript test_apps_script.js
ObjC.import('Foundation');

function readFile(path) {
  return $.NSString.stringWithContentsOfFileEncodingError(path, $.NSUTF8StringEncoding, null).js;
}

// ---- Apps Script 전역 객체 가짜 구현 ----
var scriptProps = {};
var PropertiesService = { getScriptProperties: function() { return { getProperty: function(k) { return scriptProps[k] || null; } }; } };
var LockService = { getScriptLock: function() { return { waitLock: function() {}, releaseLock: function() {} }; } };
var ContentService = {
  MimeType: { JSON: 'json' },
  createTextOutput: function(text) { return { text: text, setMimeType: function() { return this; } }; }
};

function makeSheet(name, values, formulas) {
  formulas = formulas || {};
  return {
    values: values,
    appended: [],
    getName: function() { return name; },
    getDataRange: function() { var v = this.values; return { getValues: function() { return v; } }; },
    getLastColumn: function() { return this.values.reduce(function(m, r) { return Math.max(m, r.length); }, 0); },
    getRange: function(r, c, nr, nc) {
      var self = this;
      if (nr) return { getValues: function() { return [self.values[r - 1].slice(c - 1, c - 1 + nc)]; } };
      return {
        getFormula: function() { return formulas[r + ',' + c] || ''; },
        setValue: function(v) { while (self.values[r - 1].length < c) self.values[r - 1].push(''); self.values[r - 1][c - 1] = v; }
      };
    },
    appendRow: function(row) { this.appended.push(row); this.values.push(row); }
  };
}

function makeSS(sheets) {
  return {
    getSheets: function() { return sheets; },
    getSheetByName: function(n) { return sheets.filter(function(s) { return s.getName() === n; })[0] || null; },
    insertSheet: function(n) { var s = makeSheet(n, []); sheets.push(s); return s; }
  };
}
var SpreadsheetApp = { getActiveSpreadsheet: function() { return currentSS; } };
var currentSS = null;

eval(readFile('google_apps_script.js'));

// ---- 테스트 도구 ----
var failures = 0, passes = 0;
function check(name, cond, info) {
  if (cond) { passes++; } else { failures++; console.log('FAIL: ' + name + (info !== undefined ? ' → ' + JSON.stringify(info) : '')); }
}

// 4. 선준: 금액 0 빈 행 제외
(function() {
  var ss = makeSS([makeSheet('선준 카드값', [
    ['카드', '항목', '금액', '구분', '비고', '월', '년'],
    ['하나카드', '병원', '', '', '', '', ''],
    ['하나카드', '휴대폰', 100000, '공통', '1/12', '10월', '2026년']
  ])]);
  var s = getSeonjunData(ss);
  check('선준: 빈 행 제외', s.length === 1 && s[0].item === '휴대폰' && s[0].note === '할부', s);
})();

// 5. 월 금액 쓰기: 연도별 같은 달 구분 + 수식 셀 보호
function mainSheet() {
  return makeSheet('가계부', [
    ['', '월', '급여', '급여-선준카드값', '다정남은금액', '남은금액', '전체비용', '적금'],
    ['2024년', '10월', 3000000, 2000000, 100000, 2100000, 1000000, 2000000],
    ['', '11월', 3000000, 2000000, 100000, 2100000, 1000000, 2000000],
    ['2026년', '10월', 4000000, 2200000, 300000, 2500000, 1800000, 2000000]
  ], { '4,4': '=C4-G4' });
}
(function() {
  var sheet = mainSheet();
  var ss = makeSS([sheet]);
  var res = postUpdateMonthField_(ss, { year: '2026년', month: '10월', field: 'salary', amount: 4100000 });
  check('월 수정: success', res.status === 'success', res);
  check('월 수정: 2026년 10월 행에 기록', sheet.values[3][2] === 4100000, sheet.values[3]);
  check('월 수정: 2024년 10월은 그대로', sheet.values[1][2] === 3000000, sheet.values[1]);

  var res2 = postUpdateMonthField_(ss, { year: '2026년', month: '10월', field: 'salaryMinusCard', amount: 1 });
  check('월 수정: 수식 셀은 덮어쓰지 않음', res2.status === 'error' && sheet.values[3][3] === 2200000, res2);

  var res3 = postUpdateMonthField_(ss, { year: '2026년', month: '12월', field: 'salary', amount: 1 });
  check('월 수정: 없는 달은 error', res3.status === 'error' && res3.message.indexOf('찾지 못했습니다') !== -1, res3);

  var res4 = postUpdateMonthField_(ss, { year: '2026년', month: '10월', field: 'pocketMoney', amount: 1 });
  check('월 수정: 열 없음 error', res4.status === 'error' && res4.message.indexOf('열 없음') !== -1, res4);

  var res5 = postAddMonth_(ss, { year: '2026년', month: '10월', salary: 4100000, savings: 1800000, pocketMoney: 100000 });
  check('월 추가: 일부만 저장되면 partial', res5.status === 'partial', res5);
})();

// 6. 카드 항목 추가: 헤더 순서대로 기록
(function() {
  var sj = makeSheet('선준 카드값', [['월', '년', '카드', '항목', '금액', '구분', '비고']]);
  var dj = makeSheet('다정 카드값', [[50000]]);
  var ss = makeSS([sj, dj]);
  var r1 = postAddCardItem_(ss, 'seonjun', { card: '하나카드', item: '커피', amount: 4500, month: '10월', year: '2026년' });
  check('선준 추가: success', r1.status === 'success', r1);
  check('선준 추가: 헤더 순서', JSON.stringify(sj.appended[0]) === JSON.stringify(['10월', '2026년', '하나카드', '커피', 4500, '공통', '일반']), sj.appended[0]);
  var r2 = postAddCardItem_(ss, 'dajeong', { item: '마트', amount: 1000 });
  check('다정 카드 추가는 지원 안 함', r2.status === 'error' && dj.appended.length === 0, r2);
  var r3 = postAddCardItem_(makeSS([]), 'seonjun', { item: 'x', amount: 1 });
  check('선준 시트 없음: error', r3.status === 'error', r3);
})();

// 7. 내집마련 자산 수정: 읽은 위치에 기록
(function() {
  var sheet = makeSheet('저축비용', [
    ['주식', '보증금', '주택청약', '월별 저축', '파킹통장', '저축합계'],
    [100, 30000000, 9000000, 38000000, 2000, 0]
  ]);
  var ss = makeSS([sheet]);
  var r = postUpdateHouseSavingsField_(ss, 'deposit', 31000000);
  check('내집마련: success', r.status === 'success' && sheet.values[1][1] === 31000000, [r, sheet.values[1]]);
  var r2 = postUpdateHouseSavingsField_(ss, 'stock', 1);
  check('내집마련 표 없으면 null', getHouseSavingsData(makeSS([makeSheet('기타', [['a']])])) === null);
  check('내집마련: 주식은 거부', r2.status === 'error' && sheet.values[1][0] === 100, r2);
})();

// 8. 접근 키 & 응답 형식
(function() {
  currentSS = makeSS([mainSheet()]);
  scriptProps = {};
  var open = JSON.parse(doGet({ parameter: {} }).text);
  check('키 미설정: 열림', open.status === 'success' && open.accessMode === 'open', open.status);
  check('응답에 _loc/debug 없음', JSON.stringify(open).indexOf('_loc') === -1 && !open.debugSheets && !open.debugMainSheet);
  check('scriptVersion v3.9', open.scriptVersion === 'v3.9', open.scriptVersion);
  check('키 미설정이면 dataKey 미포함', open.dataKey === undefined, open.dataKey);
  check('다정 카드 데이터 미포함', open.dajeong === undefined && open.summary.length > 0 && open.summary[0].dajeongRemain === 100000, open.summary && open.summary[0]);

  scriptProps = { DASHBOARD_KEY: 'secret', DATA_KEY: 'k'.repeat(64) };
  var denied = JSON.parse(doGet({ parameter: { key: 'wrong' } }).text);
  check('키 틀림: unauthorized', denied.status === 'unauthorized' && !denied.summary && !denied.dataKey, denied);
  var ok = JSON.parse(doGet({ parameter: { key: 'secret', debug: '1' } }).text);
  check('키 맞음: dataKey 전달', ok.dataKey === 'k'.repeat(64), ok.dataKey);
  check('키 맞음 + debug', ok.status === 'success' && ok.accessMode === 'key' && Array.isArray(ok.debugSheets), ok.status);

  var postDenied = JSON.parse(doPost({ postData: { contents: JSON.stringify({ action: 'updateMonthField', key: 'nope', data: {} }) } }).text);
  check('POST 키 틀림: unauthorized', postDenied.status === 'unauthorized', postDenied);
  var postUnknown = JSON.parse(doPost({ postData: { contents: JSON.stringify({ action: 'xyz', key: 'secret' }) } }).text);
  check('POST 알 수 없는 요청: error', postUnknown.status === 'error', postUnknown);
  scriptProps = {};
})();

console.log('test_apps_script: ' + passes + ' passed, ' + failures + ' failed');
if (failures) throw new Error(failures + ' test(s) failed');
