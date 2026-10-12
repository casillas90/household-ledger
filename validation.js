/**
 * 가계부 데이터 점검 로직
 *
 * Apps Script(doGet) 응답을 받아 숫자가 서로 맞는지, 시트 입력이 빠졌는지 확인하고
 * [{ level: 'error' | 'warn' | 'info', title, detail }] 목록을 돌려준다.
 * 브라우저에서는 window.LedgerValidation, 테스트에서는 globalThis.LedgerValidation 으로 사용.
 */
(function (root) {
  var EXPECTED_SCRIPT_VERSION = 'v3.9';
  var TOLERANCE = 10;           // 원 단위 반올림 오차 허용
  var RECENT_MONTHS = 3;        // 금액 계산식은 최근 N개월만 점검 (과거 달은 시트 기록 그대로 둠)

  function ymNum(year, month) {
    var y = parseInt(String(year || '').replace(/[^0-9]/g, ''), 10) || 0;
    var m = parseInt(String(month || '').replace(/[^0-9]/g, ''), 10) || 0;
    return y * 100 + m;
  }

  function won(n) {
    return Math.round(n || 0).toLocaleString('ko-KR') + '원';
  }

  function isNumericOnly(str) {
    return /^[\s\d,.\-원₩]+$/.test(String(str || '')) && /\d/.test(String(str || ''));
  }

  function listPreview(rows, fmt, max) {
    max = max || 3;
    var shown = rows.slice(0, max).map(fmt).join(', ');
    return rows.length > max ? shown + ' 외 ' + (rows.length - max) + '건' : shown;
  }

  function validateLedgerData(raw, options) {
    options = options || {};
    var expectedVersion = options.expectedVersion || EXPECTED_SCRIPT_VERSION;
    var issues = [];
    var add = function (level, title, detail) { issues.push({ level: level, title: title, detail: detail || '' }); };

    if (!raw || typeof raw !== 'object') {
      add('error', '데이터 없음', 'Apps Script 응답이 비어 있습니다.');
      return issues;
    }
    if (raw.status && raw.status !== 'success') {
      add('error', 'Apps Script 오류', raw.message || raw.status);
      return issues;
    }

    // 1. 배포/보안 상태
    if (raw.scriptVersion !== expectedVersion) {
      add('warn', 'Apps Script가 최신 버전이 아님',
        '현재 ' + (raw.scriptVersion || '알 수 없음') + ', 필요 ' + expectedVersion + '. 시트의 Apps Script에 새 코드를 붙여넣고 [새 배포]를 해 주세요.');
    }
    if (raw.accessMode === 'open') {
      add('warn', '접근 키 미설정', '주소를 아는 누구나 가계부 데이터를 볼 수 있습니다. Apps Script 스크립트 속성에 DASHBOARD_KEY를 설정하세요.');
    }
    if (options.privateDataLoaded === false) {
      add('warn', '과거 기본 데이터 잠김',
        raw.dataKey ? '암호화 파일(private_data.enc.json)을 풀지 못했습니다. DATA_KEY 값이 private/data_key.txt 와 같은지 확인하세요.'
                    : 'Apps Script 스크립트 속성 DATA_KEY가 없어 내 집 마련 탭과 과거 기본값을 표시할 수 없습니다.');
    }

    var summary = Array.isArray(raw.summary) ? raw.summary : [];
    var seonjun = Array.isArray(raw.seonjun) ? raw.seonjun : [];

    if (!summary.length) {
      add('error', '월별 요약을 읽지 못함', '가계부 시트의 헤더(급여, 남은금액 등)를 찾지 못했습니다.');
    }

    // 2. 요약에 같은 년/월이 두 번 나오는지
    var seen = {};
    var dups = [];
    summary.forEach(function (s) {
      var k = s.year + ' ' + s.month;
      if (seen[k]) dups.push(k);
      seen[k] = true;
    });
    if (dups.length) {
      add('warn', '월별 요약에 중복된 달', dups.join(', ') + ' — 연도 표기가 빠진 행이 있는지 확인하세요.');
    }

    // 3. 최근 달 계산식 점검
    var seonjunSums = {};
    seonjun.forEach(function (r) {
      if (!r.month) return;
      var k = r.year + ' ' + r.month;
      seonjunSums[k] = (seonjunSums[k] || 0) + (Number(r.amount) || 0);
    });

    var recent = summary.slice().sort(function (a, b) {
      return ymNum(b.year, b.month) - ymNum(a.year, a.month);
    }).slice(0, RECENT_MONTHS);

    recent.forEach(function (s) {
      var label = s.year + ' ' + s.month;
      var cardSum = seonjunSums[label];

      if (cardSum !== undefined && s.totalCost > 0 && Math.abs(cardSum - s.totalCost) > TOLERANCE) {
        add('warn', label + ' 선준 카드값 불일치',
          '카드 내역 합계 ' + won(cardSum) + ' vs 가계부 ' + won(s.totalCost) + ' (차이 ' + won(cardSum - s.totalCost) + ')');
      }
      if (s.salary > 0 && s.totalCost > 0 && s.salaryMinusCard &&
          Math.abs(s.salary - s.totalCost - s.salaryMinusCard) > TOLERANCE) {
        add('warn', label + ' 급여−카드값 계산 불일치',
          won(s.salary) + ' − ' + won(s.totalCost) + ' = ' + won(s.salary - s.totalCost) + ' 인데 시트 값은 ' + won(s.salaryMinusCard));
      }
      if (s.remain > 0 && s.salaryMinusCard && Math.abs(s.salaryMinusCard + (s.dajeongRemain || 0) - s.remain) > TOLERANCE) {
        add('warn', label + ' 월 잔금 계산 불일치',
          '급여−카드값 ' + won(s.salaryMinusCard) + ' + 다정 ' + won(s.dajeongRemain) + ' = ' +
          won(s.salaryMinusCard + (s.dajeongRemain || 0)) + ' 인데 잔금은 ' + won(s.remain));
      }
      var allocSum = (s.savings || 0) + (s.pocketMoney || 0) + (s.livingExpense || 0) + (s.emergencyFund || 0);
      if (s.remain > 0 && allocSum > 0 && Math.abs(allocSum - s.remain) > TOLERANCE) {
        add('warn', label + ' 잔금 배분 합계 불일치',
          '적금+용돈+생활비+비상금 = ' + won(allocSum) + ' 인데 잔금은 ' + won(s.remain));
      }
      if (s.emergencyFund < 0) {
        add('warn', label + ' 비상금이 마이너스', won(s.emergencyFund) + ' — 잔금보다 배분 금액이 큽니다.');
      }
    });

    // 4. 입력 누락/형식 오류
    var seonjunNoMonth = seonjun.filter(function (r) { return !r.month; });
    if (seonjunNoMonth.length) {
      add('warn', '선준 카드값: 월이 비어 있는 행 ' + seonjunNoMonth.length + '건',
        listPreview(seonjunNoMonth, function (r) { return (r.item || '(항목 없음)') + ' ' + won(r.amount); }) + ' — 어느 달 통계에도 들어가지 않습니다.');
    }
    var seonjunNumericItems = seonjun.filter(function (r) { return isNumericOnly(r.item); });
    if (seonjunNumericItems.length) {
      add('warn', '선준 카드값: 항목 이름이 숫자인 행 ' + seonjunNumericItems.length + '건',
        listPreview(seonjunNumericItems, function (r) { return r.item; }) + ' — 열이 밀려 있을 수 있습니다.');
    }


    // 5. 카드 내역은 있는데 요약에 없는 달
    var summaryKeys = {};
    summary.forEach(function (s) { summaryKeys[s.year + ' ' + s.month] = true; });
    var latestSummary = recent.length ? ymNum(recent[0].year, recent[0].month) : 0;
    var orphanMonths = Object.keys(seonjunSums).filter(function (k) {
      var p = k.split(' ');
      return !summaryKeys[k] && ymNum(p[0], p[1]) > latestSummary;
    });
    if (orphanMonths.length) {
      add('info', '카드 내역만 있는 달', orphanMonths.join(', ') + ' — 가계부 시트에 해당 월 행을 추가하면 요약에 반영됩니다.');
    }

    // 6. 내집마련 자산 합계
    var h = raw.houseSavings;
    if (h && h.total) {
      var hSum = (h.stock || 0) + (h.deposit || 0) + (h.housingSubscription || 0) + (h.monthlySavings || 0) + (h.parkingAccount || 0);
      if (Math.abs(hSum - h.total) > TOLERANCE) {
        add('warn', '내집마련 자산 합계 불일치', '항목 합 ' + won(hSum) + ' vs 저축합계 ' + won(h.total));
      }
    }

    var rank = { error: 0, warn: 1, info: 2 };
    issues.sort(function (a, b) { return rank[a.level] - rank[b.level]; });
    return issues;
  }

  root.LedgerValidation = {
    validateLedgerData: validateLedgerData,
    EXPECTED_SCRIPT_VERSION: EXPECTED_SCRIPT_VERSION
  };
})(typeof window !== 'undefined' ? window : globalThis);
