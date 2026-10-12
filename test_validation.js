// validation.js 단위 테스트
// 실행: osascript -l JavaScript test_validation.js
ObjC.import('Foundation');
eval($.NSString.stringWithContentsOfFileEncodingError('validation.js', $.NSUTF8StringEncoding, null).js);
var validate = globalThis.LedgerValidation.validateLedgerData;

var failures = 0, passes = 0;
function check(name, cond, info) {
  if (cond) { passes++; } else { failures++; console.log('FAIL: ' + name + (info !== undefined ? ' → ' + JSON.stringify(info) : '')); }
}
function titles(issues) { return issues.map(function(i) { return i.level + ':' + i.title; }); }
function has(issues, level, text) { return issues.some(function(i) { return i.level === level && i.title.indexOf(text) !== -1; }); }

function healthy() {
  return {
    status: 'success', scriptVersion: 'v3.9', accessMode: 'key',
    summary: [
      { year: '2026년', month: '9월', salary: 3000000, totalCost: 1000000, salaryMinusCard: 2000000, dajeongRemain: 100000, remain: 2100000, savings: 1800000, pocketMoney: 100000, livingExpense: 150000, emergencyFund: 50000 },
      { year: '2026년', month: '10월', salary: 4000000, totalCost: 1800000, salaryMinusCard: 2200000, dajeongRemain: 300000, remain: 2500000, savings: 1800000, pocketMoney: 200000, livingExpense: 300000, emergencyFund: 200000 }
    ],
    seonjun: [
      { item: '보험', amount: 1799500, month: '10월', year: '2026년' },
      { item: '문자', amount: 500, month: '10월', year: '2026년' }
    ],
    houseSavings: { stock: 1, deposit: 2, housingSubscription: 3, monthlySavings: 4, parkingAccount: 5, total: 15 }
  };
}

// 1. 정상 데이터 → 문제 없음
var ok = validate(healthy());
check('정상 데이터: 이슈 0건', ok.length === 0, titles(ok));

// 2. 응답 오류
check('null 응답', has(validate(null), 'error', '데이터 없음'));
check('Apps Script 오류', has(validate({ status: 'error', message: 'x' }), 'error', 'Apps Script 오류'));

// 3. 버전/공개 상태
var d = healthy(); d.scriptVersion = 'v3.6-debug'; d.accessMode = 'open';
var r = validate(d);
check('구버전 경고', has(r, 'warn', '최신 버전이 아님'), titles(r));
check('접근 키 미설정 경고', has(r, 'warn', '접근 키 미설정'), titles(r));

// 4. 카드 합계 불일치
d = healthy(); d.seonjun[0].amount = 1700000;
r = validate(d);
check('카드값 불일치', has(r, 'warn', '2026년 10월 선준 카드값 불일치'), titles(r));

// 5. 계산식 불일치 / 마이너스 비상금
d = healthy(); d.summary[1].salaryMinusCard = 2300000;
r = validate(d);
check('급여−카드값 불일치', has(r, 'warn', '급여−카드값'), titles(r));
check('잔금 계산 불일치도 함께', has(r, 'warn', '월 잔금 계산 불일치'), titles(r));

d = healthy(); d.summary[1].emergencyFund = -300000; d.summary[1].savings = 2300000;
r = validate(d);
check('비상금 마이너스', has(r, 'warn', '비상금이 마이너스'), titles(r));
check('배분 합계는 맞으면 경고 없음', !has(r, 'warn', '배분 합계'), titles(r));

d = healthy(); d.summary[1].savings = 2100000;
check('배분 합계 불일치', has(validate(d), 'warn', '잔금 배분 합계 불일치'));

// 6. 오래된 달은 계산식 점검 제외 (최근 3개월만)
d = healthy();
d.summary.unshift(Object.assign({}, d.summary[0], { month: '8월' }));
d.summary.unshift({ year: '2024년', month: '10월', salary: 100, totalCost: 50, salaryMinusCard: 1, remain: 5, savings: 999 });
check('과거 달 계산 불일치는 무시', validate(d).length === 0, titles(validate(d)));

// 7. 중복 달
d = healthy(); d.summary.push(JSON.parse(JSON.stringify(d.summary[1])));
check('중복 달', has(validate(d), 'warn', '중복된 달'));

// 8. 입력 누락
d = healthy();
d.seonjun.push({ item: '병원', amount: 0, month: '', year: '2026년' });
r = validate(d);
check('선준 월 누락', has(r, 'warn', '월이 비어 있는 행 1건'), titles(r));

// 9. 카드 내역만 있는 다음 달
d = healthy(); d.seonjun.push({ item: '커피', amount: 4500, month: '11월', year: '2026년' });
r = validate(d);
check('카드 내역만 있는 달', has(r, 'info', '카드 내역만 있는 달'));
d.scriptVersion = 'v3.7';
r = validate(d);
check('정렬: warn 이 info 보다 먼저', r[0].level === 'warn' && r[r.length - 1].level === 'info', titles(r));

// 10. 내집마련 합계
d = healthy(); d.houseSavings.total = 99;
check('내집마련 합계 불일치', has(validate(d), 'warn', '내집마련'));

// 11. 개인 데이터 잠김
d = healthy();
check('복호화 성공이면 경고 없음', validate(d, { privateDataLoaded: true }).length === 0);
check('DATA_KEY 없음', has(validate(d, { privateDataLoaded: false }), 'warn', '과거 기본 데이터 잠김'));
d.dataKey = 'abc';
check('키 불일치 안내', validate(d, { privateDataLoaded: false })[0].detail.indexOf('DATA_KEY 값') !== -1);

console.log('test_validation: ' + passes + ' passed, ' + failures + ' failed');
if (failures) throw new Error(failures + ' test(s) failed');
