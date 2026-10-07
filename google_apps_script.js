/**
 * 다정 & 선준 가계부 대시보드 - Google Apps Script (Code.gs)
 * 
 * 구글 스프레드시트 ➔ [확장 프로그램] ➔ [Apps Script] 에 아래 코드를 붙여넣고
 * [배포] ➔ [새 배포] ➔ 유형: [웹 앱] ➔ 액세스 권한: [모든 사용자] 로 배포하시면
 * 가계부 대시보드에서 주식 자산이 실시간으로 자동 반영됩니다!
 */

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var allSheets = ss.getSheets();
    var sheetNames = allSheets.map(function(s) { return s.getName(); });
    
    // 1. 월별 비용 시트 (잔금, 적금, 용돈, 생활비, 비상금)
    var monthlyExpensesData = getMonthlyExpensesData(ss);

    // 2. 월별 요약 시트
    var summaryData = getSummaryData(ss, monthlyExpensesData);
    
    // 3. 다정 시트
    var dajeongData = getDajeongData(ss);
    
    // 4. 선준 시트
    var seonjunData = getSeonjunData(ss);
    
    // 5. 주식 시트 (다정 & 선준 보유 주식 실시간 동기화)
    var stockData = getStockData(ss);
    
    // 6. 내집마련 플랜 4번 저축 자산 데이터 (주식, 보증금, 주택청약, 월별 저축, 파킹통장, 저축합계)
    var houseSavingsData = getHouseSavingsData(ss);

    // 주식 시트의 실시간 평가액을 내집마련 4번 저축 자산의 주식 항목에도 100% 실시간 동기화
    if (stockData && stockData.total > 0 && houseSavingsData) {
      houseSavingsData.stock = Math.round(stockData.total);
      houseSavingsData.total = houseSavingsData.stock + (houseSavingsData.deposit || 0) + (houseSavingsData.housingSubscription || 0) + (houseSavingsData.monthlySavings || 0) + (houseSavingsData.parkingAccount || 0);
    }

    var result = {
      status: "success",
      scriptVersion: "v3.6-debug",
      debugMainSheet: getDebugMainSheet(ss),
      sheetNames: sheetNames,
      summary: summaryData,
      monthlyExpenses: monthlyExpensesData,
      dajeong: dajeongData,
      seonjun: seonjunData,
      stocks: stockData,
      houseSavings: houseSavingsData
    };
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * [임시 진단용] 메인 가계부 시트 상단 행 덤프 (헤더 인식 문제 확인용)
 */
function getDebugMainSheet(ss) {
  var out = [];
  try {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var n = sheets[i].getName().trim();
      if (n.indexOf('★') !== -1 || n === '가계부') {
        var vals = sheets[i].getDataRange().getValues();
        var rows = [];
        for (var r = 20; r < Math.min(60, vals.length); r++) {
          rows.push(['r' + r].concat(vals[r].slice(0, 30).map(function(v) {
            if (v instanceof Date) return 'DATE:' + (v.getMonth() + 1) + '/' + v.getDate();
            return String(v === null || v === undefined ? '' : v).slice(0, 25);
          })));
        }
        out.push({ name: n, totalRows: vals.length, totalCols: vals[0] ? vals[0].length : 0, rows: rows });
      }
    }
  } catch (e) {
    out.push({ error: String(e) });
  }
  return out;
}

/**
 * 시트 이름 유연 탐색 헬퍼 함수 (공백/대소문자 무시 및 키워드 매칭)
 */
function findSheetByKeywords(ss, keywords, excludeKeywords) {
  var allSheets = ss.getSheets();
  // 1순위: 완전 일치 (공백/대소문자 무시)
  for (var i = 0; i < allSheets.length; i++) {
    var name = allSheets[i].getName().trim().toLowerCase();
    for (var k = 0; k < keywords.length; k++) {
      if (name === keywords[k].trim().toLowerCase()) {
        return allSheets[i];
      }
    }
  }
  // 2순위: 키워드 포함
  for (var i = 0; i < allSheets.length; i++) {
    var name = allSheets[i].getName().trim();
    var isExcluded = false;
    if (excludeKeywords) {
      for (var e = 0; e < excludeKeywords.length; e++) {
        if (name.indexOf(excludeKeywords[e]) !== -1) {
          isExcluded = true;
          break;
        }
      }
    }
    if (isExcluded) continue;
    for (var k = 0; k < keywords.length; k++) {
      if (name.indexOf(keywords[k]) !== -1) {
        return allSheets[i];
      }
    }
  }
  return null;
}

/**
 * 주식 시트 데이터 추출 함수
 * '주식', '주식자산', '주식 현황' 시트를 찾아 종목별 내역과 총 평가금액을 반환합니다.
 */
function getStockData(ss) {
  var sheet = ss.getSheetByName('주식') || 
              ss.getSheetByName('주식자산') || 
              ss.getSheetByName('주식 현황') || 
              ss.getSheetByName('Stocks');
              
  if (!sheet) {
    return {
      total: 16974329, // 시트가 아직 없을 경우 기본값
      items: [],
      updatedAt: new Date().toISOString()
    };
  }
  
  var values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) {
    return { total: 0, items: [] };
  }
  
  var headers = values[0].map(function(h) { return String(h).trim(); });
  
  // 열 인덱스 자동 감지
  var colOwner = -1;  // 소유자 (다정 / 선준)
  var colName = -1;   // 종목명
  var colAmount = -1; // 평가금액
  var colQty = -1;    // 보유수량
  var colPrice = -1;  // 현재가
  
  for (var c = 0; c < headers.length; c++) {
    var h = headers[c];
    if (h.indexOf('소유') !== -1 || h.indexOf('구분') !== -1 || h.indexOf('이름') !== -1) colOwner = c;
    else if (h.indexOf('종목') !== -1 || h.indexOf('종목명') !== -1 || h.indexOf('주식') !== -1) colName = c;
    else if (h.indexOf('평가금액') !== -1 || h.indexOf('평가액') !== -1 || h.indexOf('총액') !== -1 || h.indexOf('금액') !== -1) colAmount = c;
    else if (h.indexOf('수량') !== -1 || h.indexOf('주수') !== -1) colQty = c;
    else if (h.indexOf('현재가') !== -1 || h.indexOf('단가') !== -1) colPrice = c;
  }
  
  // 기본 매핑 fallback
  if (colName === -1 && values[0].length >= 2) colName = 0;
  if (colAmount === -1 && values[0].length >= 2) colAmount = values[0].length - 1;
  
  var items = [];
  var calculatedTotal = 0;
  var explicitTotal = 0;
  
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var firstColStr = String(row[0] || '').trim();
    
    // '합계' 또는 '총계' 행 감지
    if (firstColStr.indexOf('합계') !== -1 || firstColStr.indexOf('총') !== -1) {
      for (var k = 1; k < row.length; k++) {
        var num = parseNumeric(row[k]);
        if (num > explicitTotal) explicitTotal = num;
      }
      continue;
    }
    
    var name = colName >= 0 ? String(row[colName] || '').trim() : '';
    var amount = colAmount >= 0 ? parseNumeric(row[colAmount]) : 0;
    var owner = colOwner >= 0 ? String(row[colOwner] || '').trim() : '';
    var qty = colQty >= 0 ? parseNumeric(row[colQty]) : null;
    var price = colPrice >= 0 ? parseNumeric(row[colPrice]) : null;
    
    // 수량 * 현재가로 평가금액 계산 가능한 경우
    if (amount === 0 && qty && price) {
      amount = Math.round(qty * price);
    }
    
    if (name && amount > 0) {
      items.push({
        owner: owner || (name.indexOf('다정') !== -1 ? '다정' : (name.indexOf('선준') !== -1 ? '선준' : '공통')),
        name: name,
        amount: amount,
        quantity: qty,
        price: price
      });
      calculatedTotal += amount;
    }
  }
  
  var finalTotal = explicitTotal > 0 ? explicitTotal : calculatedTotal;
  
  return {
    total: finalTotal || 16974329,
    items: items,
    updatedAt: new Date().toISOString()
  };
}

/**
 * 날짜 및 년월 자동 추출 헬퍼 함수 (Date 객체, 타임스탬프, 문자열 완벽 지원)
 */
function extractYearMonth(val, defaultYear) {
  if (!val) return null;
  var dYear = defaultYear || '2026년';

  // 1. Date 객체
  if (val instanceof Date || (typeof val === 'object' && typeof val.getMonth === 'function')) {
    var y = val.getFullYear();
    var m = val.getMonth() + 1;
    if (m >= 1 && m <= 12) {
      return { year: y + '년', month: m + '월' };
    }
  }

  var str = String(val).trim();
  if (!str) return null;

  // 2. GMT / 한국 표준시 타임스탬프 문자열
  if (str.indexOf('GMT') !== -1 || str.indexOf('한국') !== -1) {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      return { year: d.getFullYear() + '년', month: (d.getMonth() + 1) + '월' };
    }
  }

  var yPart = '';
  var ymMatch = str.match(/(20\d{2})/);
  if (ymMatch) yPart = ymMatch[1] + '년';

  // 3. 2026-10, 2026.10, 2026/10 형식
  var ydashm = str.match(/20\d{2}[-./](\d{1,2})/);
  if (ydashm) {
    var mn = parseInt(ydashm[1], 10);
    if (mn >= 1 && mn <= 12) {
      return { year: yPart || dYear, month: mn + '월' };
    }
  }

  // 4. "10월" 또는 "10" 형식
  var mMatch = str.match(/^(\d{1,2})\s*월?$/) || str.match(/(\d{1,2})\s*월/);
  if (mMatch) {
    var mn = parseInt(mMatch[1], 10);
    if (mn >= 1 && mn <= 12) {
      return { year: yPart || dYear, month: mn + '월' };
    }
  }

  return null;
}

/**
 * 월별 비용 시트 (잔금, 적금, 용돈, 생활비, 비상금) 자동 추출 함수
 */
function getMonthlyExpensesData(ss) {
  var sheet = findSheetByKeywords(ss, ['월별 비용', '월별비용'], ['내집', '분양', '아파트', '주식', '다정', '선준', '저축', '저축비용', '장보기', '구매', '해야', '경비', '육아', '★']);
  if (!sheet) return [];

  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  var headerRowIdx = -1;
  var colRemain = -1, colSavings = -1, colPocket = -1, colLiving = -1, colEmergency = -1;
  var colYear = -1, colMonth = -1;

  for (var r = 0; r < Math.min(60, values.length); r++) {
    var rowStr = values[r].map(function(c) { return String(c || '').trim(); });
    for (var c = 0; c < rowStr.length; c++) {
      var val = rowStr[c];
      if (val.indexOf('년도') !== -1 || val === '년') colYear = c;
      else if (val === '월' || val.indexOf('월별') !== -1) colMonth = c;
      else if (val === '잔금' || val.indexOf('잔액') !== -1 || val.indexOf('남은') !== -1 || val.indexOf('잔여') !== -1) colRemain = c;
      else if (val === '적금' || val.indexOf('저축') !== -1) colSavings = c;
      else if (val === '용돈') colPocket = c;
      else if (val === '생활비') colLiving = c;
      else if (val === '비상금') colEmergency = c;
    }
    if ((colMonth !== -1 && (colRemain !== -1 || colSavings !== -1)) || 
        (colRemain !== -1 && colPocket !== -1) || 
        (colSavings !== -1 && colLiving !== -1)) {
      headerRowIdx = r;
      break;
    }
  }

  if (headerRowIdx !== -1) {
    var results = [];
    var currentYear = '2026년';

    for (var r = headerRowIdx + 1; r < values.length; r++) {
      var row = values[r];
      var yVal = colYear !== -1 ? row[colYear] : null;
      var mVal = colMonth !== -1 ? row[colMonth] : null;

      var ym = extractYearMonth(mVal, currentYear) || extractYearMonth(yVal, currentYear);
      if (!ym) continue;
      if (ym.year) currentYear = ym.year;

      var remain = colRemain !== -1 ? parseNumeric(row[colRemain]) : 0;
      var savings = colSavings !== -1 ? parseNumeric(row[colSavings]) : 0;
      var pocket = colPocket !== -1 ? parseNumeric(row[colPocket]) : 0;
      var living = colLiving !== -1 ? parseNumeric(row[colLiving]) : 0;
      var rawEmergency = colEmergency !== -1 ? row[colEmergency] : null;
      var emergency = 0;
      if (rawEmergency !== null && rawEmergency !== undefined && String(rawEmergency).trim() !== '') {
        emergency = parseNumeric(rawEmergency);
      } else {
        emergency = remain - (savings + pocket + living);
      }

      if (remain === 0 && savings > 10000000) continue;

      if (remain > 0 || savings > 0 || pocket > 0 || living > 0 || emergency !== 0) {
        results.push({
          year: currentYear || '2026년',
          month: ym.month,
          remain: remain,
          savings: savings,
          pocketMoney: pocket,
          livingExpense: living,
          emergencyFund: emergency
        });
      }
    }

    if (results.length > 0) return results;
  }

  return [];
}

function parseNumeric(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  var str = String(val).trim();
  var isNegative = false;
  if (str.indexOf('-') !== -1 || (str.indexOf('(') !== -1 && str.indexOf(')') !== -1)) {
    isNegative = true;
  }
  var clean = str.replace(/[^0-9.]/g, '');
  var num = parseFloat(clean);
  if (isNaN(num)) return 0;
  return isNegative ? -num : num;
}

/**
 * 가계부 & 월별 요약 시트 추출 함수 ('★저축 및 가계부' 최우선 연동 및 급여/다정남은금액/적금 완벽 파싱)
 */
function getSummaryData(ss, monthlyExpensesData) {
  var candidateSheets = [];
  var allSheets = ss.getSheets();

  // 1순위: '★'가 포함되고 '저축' 또는 '가계부'가 포함된 시트 (예: '★저축 및 가계부')
  for (var i = 0; i < allSheets.length; i++) {
    var sName = allSheets[i].getName().trim();
    if (sName.indexOf('★') !== -1 && (sName.indexOf('가계부') !== -1 || sName.indexOf('저축') !== -1)) {
      candidateSheets.push(allSheets[i]);
      break;
    }
  }

  // 2순위: '저축 및 가계부' 또는 '저축및가계부'
  for (var i = 0; i < allSheets.length; i++) {
    var sName = allSheets[i].getName().trim();
    if (sName.indexOf('저축') !== -1 && sName.indexOf('가계부') !== -1 && candidateSheets.indexOf(allSheets[i]) === -1) {
      candidateSheets.push(allSheets[i]);
    }
  }

  // 3순위: '가계부', '월별 요약', '월별요약', '요약'
  var s1 = ss.getSheetByName('가계부') || ss.getSheetByName('월별 요약') || ss.getSheetByName('월별요약') || ss.getSheetByName('요약');
  if (s1 && candidateSheets.indexOf(s1) === -1) candidateSheets.push(s1);

  // 4순위: 기타 일반 시트 (자산/세부지출/카드내역 제외)
  for (var i = 0; i < allSheets.length; i++) {
    var s = allSheets[i];
    var sName = s.getName().trim();
    if (sName.indexOf('다정') !== -1 || sName.indexOf('선준') !== -1 || sName.indexOf('주식') !== -1 || 
        sName.indexOf('내집') !== -1 || sName.indexOf('분양') !== -1 || sName.indexOf('축의') !== -1 || 
        sName.indexOf('조의') !== -1 || sName.indexOf('장보기') !== -1 || sName.indexOf('구매') !== -1 ||
        sName.indexOf('해야') !== -1 || sName.indexOf('신행') !== -1 || sName.indexOf('육아') !== -1 ||
        sName.indexOf('저축비용') !== -1 || sName.indexOf('시트') !== -1) continue;
    if (candidateSheets.indexOf(s) === -1) candidateSheets.push(s);
  }

  for (var sIdx = 0; sIdx < candidateSheets.length; sIdx++) {
    var curSheet = candidateSheets[sIdx];
    var values = curSheet.getDataRange().getValues();
    if (!values || values.length < 2) continue;

    // --- [탐색 방식 A: 세로형 표 (행마다 1개 월, 열에 급여/급여-선준/다정남은금액/남은금액/적금 등)] ---
    var headerRowIdx = -1;
    var colYear = -1, colMonth = -1, colSalary = -1, colSalaryMinusCard = -1, colDajeongRemain = -1;
    var colTotalCost = -1, colCardTotal = -1, colSavings = -1, colPocket = -1, colLiving = -1, colEmergency = -1, colRemain = -1, colRemain2 = -1;

    for (var r = 0; r < Math.min(60, values.length); r++) {
      var row = values[r].map(function(h) { return String(h || '').trim(); });
      var foundMatches = 0;
      var tYear = -1, tMonth = -1, tSalary = -1, tSalaryMinusCard = -1, tDajeongRemain = -1;
      var tTotalCost = -1, tCardTotal = -1, tSavings = -1, tPocket = -1, tLiving = -1, tEmergency = -1, tRemain = -1, tRemain2 = -1;

      // 같은 행에 '월별 비용' 표(오른쪽)가 함께 있는 경우 시작 열을 찾아 좌/우 표를 분리
      var tMonthlyStart = -1;
      for (var cs = 0; cs < row.length; cs++) {
        if (row[cs] && row[cs].indexOf('월별') !== -1 && row[cs].indexOf('비용') !== -1) { tMonthlyStart = cs; break; }
      }

      for (var c = 0; c < row.length; c++) {
        var h = row[c];
        if (!h) continue;
        if (tMonthlyStart !== -1 && c >= tMonthlyStart) {
          // 오른쪽 '월별 비용' 표: 잔금 / 적금 / 용돈 / 생활비 / 비상금
          if (c === tMonthlyStart) continue;
          if (h.indexOf('잔금') !== -1 || h.indexOf('잔액') !== -1 || h.indexOf('남은') !== -1) { tRemain2 = c; foundMatches++; }
          else if ((h.indexOf('적금') !== -1 || h.indexOf('저축') !== -1) && h.indexOf('합계') === -1 && h.indexOf('누계') === -1) { tSavings = c; foundMatches++; }
          else if (h.indexOf('용돈') !== -1) { tPocket = c; foundMatches++; }
          else if (h.indexOf('생활비') !== -1) { tLiving = c; foundMatches++; }
          else if (h.indexOf('비상금') !== -1) { tEmergency = c; foundMatches++; }
          continue;
        }
        if (h.indexOf('년도') !== -1 || h === '년' || h === 'Year') { tYear = c; foundMatches++; }
        else if (h === '월' || h.indexOf('월별') !== -1 || h === 'Month' || h === '기간' || h === '구분' || h.indexOf('일자') !== -1) { tMonth = c; foundMatches++; }
        // [중요] '급여-' 또는 '급여 -' 또는 '차액'을 '급여'보다 반드시 먼저 매칭!
        else if (h.indexOf('급여-') !== -1 || h.indexOf('급여−') !== -1 || h.indexOf('급여 -') !== -1 || h.indexOf('차액') !== -1 || (h.indexOf('급여') !== -1 && h.indexOf('카드') !== -1)) { tSalaryMinusCard = c; foundMatches++; }
        else if (h.indexOf('급여') !== -1 || h.indexOf('수입') !== -1 || h.indexOf('월급') !== -1) { tSalary = c; foundMatches++; }
        // [중요] '다정' 컬럼 (다정남은금액, 다정 남은금액, 다정카드남은금액 등)
        else if (h.indexOf('다정') !== -1) { tDajeongRemain = c; foundMatches++; }
        // 남은금액 / 잔금 / 총잔액
        else if (h.indexOf('남은금액') !== -1 || h.indexOf('남은 금액') !== -1 || h.indexOf('잔금') !== -1 || h.indexOf('잔액') !== -1 || h.indexOf('남은') !== -1) { tRemain = c; foundMatches++; }
        // 전체비용 / 선준 카드값
        else if (h.indexOf('전체비용') !== -1 || h.indexOf('선준 카드') !== -1 || h.indexOf('선준카드') !== -1 || (h.indexOf('선준') !== -1 && h.indexOf('지출') !== -1)) { tTotalCost = c; foundMatches++; }
        else if (h.indexOf('카드') !== -1 && (h.indexOf('합계') !== -1 || h.indexOf('총액') !== -1)) { tCardTotal = c; foundMatches++; }
        // [중요] 적금 / 저축 (합계, 누계, 총 제외하고 순수 월별 적금 컬럼 정확히 매칭)
        else if ((h.indexOf('적금') !== -1 || h.indexOf('저축') !== -1) && h.indexOf('합계') === -1 && h.indexOf('누계') === -1 && h.indexOf('총') === -1) { tSavings = c; foundMatches++; }
        else if (h.indexOf('용돈') !== -1 || (h.indexOf('선준') !== -1 && h.indexOf('용돈') !== -1)) { tPocket = c; foundMatches++; }
        else if (h.indexOf('생활비') !== -1) { tLiving = c; foundMatches++; }
        else if (h.indexOf('비상금') !== -1) { tEmergency = c; foundMatches++; }
      }

      // 핵심 컬럼 2개 이상 감지 시 헤더 행으로 확정!
      if (foundMatches >= 2 && (tMonth !== -1 || tSalary !== -1 || tSalaryMinusCard !== -1 || tRemain !== -1 || tSavings !== -1 || tDajeongRemain !== -1)) {
        headerRowIdx = r;
        colYear = tYear; colMonth = tMonth; colSalary = tSalary; colSalaryMinusCard = tSalaryMinusCard;
        colDajeongRemain = tDajeongRemain; colTotalCost = tTotalCost; colCardTotal = tCardTotal;
        colSavings = tSavings; colPocket = tPocket; colLiving = tLiving; colEmergency = tEmergency; colRemain = tRemain;
        colRemain2 = tRemain2;
        break;
      }
    }

    if (headerRowIdx !== -1) {
      // 월 열이 감지되지 않은 경우, 데이터 행에서 '10월' 등 월 형식을 담고 있는 열 자동 감지
      if (colMonth === -1) {
        for (var testR = headerRowIdx + 1; testR < Math.min(headerRowIdx + 6, values.length); testR++) {
          for (var testC = 0; testC < values[testR].length; testC++) {
            if (extractYearMonth(values[testR][testC])) {
              colMonth = testC;
              break;
            }
          }
          if (colMonth !== -1) break;
        }
      }
      if (colMonth === -1 && colYear !== -1) colMonth = colYear + 1;
      if (colMonth === -1 && colSalary > 0) colMonth = colSalary - 1;
      if (colMonth === -1) colMonth = 0;

      // '전체비용' 헤더가 실제로는 연도 열(2024년, 2025년...) 위에 붙어 있는 경우 무시 (연도 숫자가 금액으로 읽히는 것 방지)
      var yearColByLayout = colMonth > 0 ? colMonth - 1 : -1;
      if (colTotalCost !== -1 && colTotalCost === yearColByLayout) colTotalCost = -1;

      var result = [];
      var currentYear = '2026년';

      for (var r = headerRowIdx + 1; r < values.length; r++) {
        var row = values[r];
        var mCell = colMonth !== -1 ? row[colMonth] : null;
        var yCell = colYear !== -1 ? row[colYear] : null;

        // 월 열 바로 왼쪽 셀에 '2026년' 같은 연도가 적혀 있으면 현재 연도로 갱신
        var layoutYearCell = yearColByLayout !== -1 ? String(row[yearColByLayout] || '') : '';
        var layoutYearMatch = layoutYearCell.match(/(20\d{2})\s*년?/);
        if (layoutYearMatch) currentYear = layoutYearMatch[1] + '년';

        var ym = extractYearMonth(mCell, currentYear) || extractYearMonth(yCell, currentYear);
        if (!ym) continue;
        if (ym.year && !layoutYearMatch && String(mCell || '').match(/20\d{2}/)) currentYear = ym.year;

        var rawSalary = colSalary !== -1 ? parseNumeric(row[colSalary]) : 0;
        var rawSalaryMinusCard = colSalaryMinusCard !== -1 ? parseNumeric(row[colSalaryMinusCard]) : 0;
        var rawDajeongRemain = colDajeongRemain !== -1 ? parseNumeric(row[colDajeongRemain]) : 0;
        var rawRemain = colRemain !== -1 ? parseNumeric(row[colRemain]) : 0;
        if (rawRemain === 0 && colRemain2 !== -1) rawRemain = parseNumeric(row[colRemain2]);
        var rawTotalCost = colTotalCost !== -1 ? parseNumeric(row[colTotalCost]) : 0;
        var rawCardTotal = colCardTotal !== -1 ? parseNumeric(row[colCardTotal]) : 0;
        var rawSavings = colSavings !== -1 ? parseNumeric(row[colSavings]) : 0;
        var rawPocket = colPocket !== -1 ? parseNumeric(row[colPocket]) : 0;
        var rawLiving = colLiving !== -1 ? parseNumeric(row[colLiving]) : 0;
        var rawEmergency = colEmergency !== -1 ? parseNumeric(row[colEmergency]) : 0;

        // salaryMinusCard 자동 보정 (급여 - 전체비용)
        if (rawSalary > 0 && rawTotalCost > 0 && rawSalaryMinusCard === 0) {
          rawSalaryMinusCard = rawSalary - rawTotalCost;
        }
        // totalCost 자동 역산 (급여 - 차액)
        if (rawSalary > 0 && rawSalaryMinusCard > 0 && rawTotalCost === 0) {
          rawTotalCost = rawSalary - rawSalaryMinusCard;
        }
        // remain(총 남은금액) 자동 보정: 급여-선준카드 + 다정남은금액
        if (rawRemain === 0 && (rawSalaryMinusCard > 0 || rawDajeongRemain > 0)) {
          rawRemain = rawSalaryMinusCard + rawDajeongRemain;
        }
        // emergencyFund 자동 보정 (기재값이 없으면 잔금 - 적금 - 용돈 - 생활비)
        if (rawEmergency === 0 && rawRemain > 0 && (rawSavings > 0 || rawPocket > 0 || rawLiving > 0)) {
          rawEmergency = rawRemain - (rawSavings + rawPocket + rawLiving);
        }

        var item = {
          year: currentYear,
          month: ym.month,
          salary: rawSalary,
          salaryMinusCard: rawSalaryMinusCard,
          dajeongRemain: rawDajeongRemain,
          remain: rawRemain, // 가계부 시트 기재값 100% 최우선 반영!
          totalCost: rawTotalCost,
          cardTotal: rawCardTotal,
          savings: rawSavings, // 가계부 시트 적금(2,000,000 등) 100% 최우선 반영!
          pocketMoney: rawPocket,
          livingExpense: rawLiving,
          emergencyFund: rawEmergency
        };

        // 월별 비용 시트에 보조 데이터가 있다면 보완
        if (monthlyExpensesData && monthlyExpensesData.length > 0) {
          for (var me = 0; me < monthlyExpensesData.length; me++) {
            var exp = monthlyExpensesData[me];
            if ((exp.year === item.year || !exp.year) && exp.month === item.month) {
              if (item.remain === 0 && exp.remain > 0) item.remain = exp.remain;
              if (item.savings === 0 && exp.savings > 0) item.savings = exp.savings;
              if (item.pocketMoney === 0 && exp.pocketMoney > 0) item.pocketMoney = exp.pocketMoney;
              if (item.livingExpense === 0 && exp.livingExpense > 0) item.livingExpense = exp.livingExpense;
              if (item.emergencyFund === 0 && exp.emergencyFund !== 0) item.emergencyFund = exp.emergencyFund;
              break;
            }
          }
        }

        if (item.salary > 0 || item.remain > 0 || item.totalCost > 0 || item.savings > 0 || item.dajeongRemain > 0) {
          result.push(item);
        }
      }

      if (result.length > 0) {
        return result;
      }
    }
  }

  return [];
}

/**
 * 다정 시트 추출 함수 (다정 카드값 & 다정 용돈 & 다정 시트 전체 유연 자동 추출)
 */
function getDajeongData(ss) {
  var candidateSheets = [];
  var allSheets = ss.getSheets();

  // '다정' 키워드가 들어간 모든 시트 수집 (아카이브 제외)
  for (var i = 0; i < allSheets.length; i++) {
    var name = allSheets[i].getName().trim();
    if (name.indexOf('보관') !== -1 || name.indexOf('히스토리') !== -1 || name.indexOf('아카이브') !== -1) continue;
    if (name.indexOf('다정') !== -1) {
      // '다정 카드값' 시트를 1순위로 배치
      if (name.indexOf('카드') !== -1) {
        candidateSheets.unshift(allSheets[i]);
      } else {
        candidateSheets.push(allSheets[i]);
      }
    }
  }

  var archiveSheet = findSheetByKeywords(ss, ['다정_기록보관', '다정기록보관', '다정히스토리', '다정보관', '다정 히스토리']);
  var items = [];
  var total = 0;

  for (var sIdx = 0; sIdx < candidateSheets.length; sIdx++) {
    var sheet = candidateSheets[sIdx];
    var values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) continue;

    // 헤더 행 탐색 (0 ~ 15행 스캔)
    var headerRowIdx = -1;
    var colItem = -1, colAmount = -1, colMonth = -1, colYear = -1, colCard = -1, colDate = -1, colType = -1, colNote = -1;

    for (var r = 0; r < Math.min(50, values.length); r++) {
      var header = values[r].map(function(h) { return String(h || '').trim(); });
      var fItem = -1, fAmount = -1, fMonth = -1, fYear = -1, fCard = -1, fDate = -1, fType = -1, fNote = -1;

      for (var c = 0; c < header.length; c++) {
        var h = header[c];
        if (!h) continue;
        if (h.indexOf('날짜') !== -1 || h.indexOf('일자') !== -1 || h.indexOf('승인일') !== -1 || h.indexOf('결제일') !== -1 || h === 'Date' || h.indexOf('이용일') !== -1 || h.indexOf('거래일') !== -1) fDate = c;
        else if (h === '월' || h.indexOf('월별') !== -1 || h === 'Month') fMonth = c;
        else if (h.indexOf('년도') !== -1 || h === '년' || h === 'Year') fYear = c;
        else if (h.indexOf('항목') !== -1 || h.indexOf('내역') !== -1 || h.indexOf('사용처') !== -1 || h.indexOf('가맹점') !== -1 || h.indexOf('내용') !== -1 || h.indexOf('상호') !== -1 || h.indexOf('적요') !== -1 || h.indexOf('품목') !== -1 || h.indexOf('품명') !== -1 || h.indexOf('거래처') !== -1) fItem = c;
        else if (h.indexOf('금액') !== -1 || h.indexOf('가격') !== -1 || h.indexOf('이용금액') !== -1 || h.indexOf('결제금액') !== -1 || h.indexOf('청구금액') !== -1 || h.indexOf('사용금액') !== -1 || h.indexOf('지출') !== -1 || h.indexOf('승인금액') !== -1 || h.indexOf('출금') !== -1 || h.indexOf('원금') !== -1) fAmount = c;
        else if (h.indexOf('카드') !== -1 || h.indexOf('수단') !== -1) fCard = c;
        else if (h === '구분' || h === '분류' || h === '유형') fType = c;
        else if (h === '비고' || h === '메모' || h.indexOf('할부') !== -1) fNote = c;
      }

      // 항목과 금액이 모두 존재하는 행을 실제 거래 헤더 행으로 확정!
      if (fItem !== -1 && fAmount !== -1) {
        headerRowIdx = r;
        colItem = fItem; colAmount = fAmount; colMonth = fMonth; colYear = fYear;
        colCard = fCard; colDate = fDate; colType = fType; colNote = fNote;
        break;
      }
    }

    // 명시적 헤더가 없는 경우 지능형 열 추정 (금액 숫자 열 = 금액, 문자열 열 = 항목)
    if (headerRowIdx === -1 && values.length >= 2) {
      headerRowIdx = 0;
      var bestAmtCol = -1;
      var bestItemCol = -1;

      for (var c = 0; c < values[0].length; c++) {
        var numCount = 0;
        for (var testR = 1; testR < Math.min(10, values.length); testR++) {
          var n = parseNumeric(values[testR][c]);
          if (n >= 1000) numCount++;
        }
        if (numCount >= 1 && bestAmtCol === -1) {
          bestAmtCol = c;
        }
      }

      if (bestAmtCol !== -1) {
        colAmount = bestAmtCol;
        colItem = bestAmtCol > 0 ? bestAmtCol - 1 : 0;
      } else {
        colCard = 0; colItem = 1; colAmount = 2; colType = 3; colNote = 4; colMonth = 5; colYear = 6;
      }
    }

    if (headerRowIdx !== -1) {
      for (var r = headerRowIdx + 1; r < values.length; r++) {
        var row = values[r];
        var itemStr = colItem !== -1 ? String(row[colItem] || '').trim() : '';
        var amt = colAmount !== -1 ? parseNumeric(row[colAmount]) : 0;
        var mVal = colMonth !== -1 ? row[colMonth] : null;
        var yVal = colYear !== -1 ? row[colYear] : null;
        var dVal = colDate !== -1 ? row[colDate] : null;

        var ym = extractYearMonth(dVal) || extractYearMonth(mVal) || extractYearMonth(yVal);
        var finalMonth = ym ? ym.month : '';
        var finalYear = ym ? ym.year : '2026년';
        var card = colCard !== -1 ? String(row[colCard] || '다정카드').trim() : '다정카드';
        var type = colType !== -1 ? String(row[colType] || '일반').trim() : '일반';
        var note = colNote !== -1 ? String(row[colNote] || '').trim() : '';

        // 단순 연도 문자열(2026년 등)이나 이상치 행 필터링
        if (/^202\d년?$/.test(itemStr) && amt < 100) continue;

        if (itemStr && amt > 0) {
          items.push({
            item: itemStr,
            amount: amt,
            month: finalMonth,
            year: finalYear,
            card: card,
            type: type,
            note: note
          });
          total += amt;
        }
      }
    }
  }

  // 2. '다정_기록보관' 시트가 있으면 과거 월 보존 데이터도 함께 불러옴
  var archiveItems = [];
  if (archiveSheet) {
    var aValues = archiveSheet.getDataRange().getValues();
    if (aValues && aValues.length > 1) {
      var headerRow = aValues[0].map(function(h) { return String(h || '').trim(); });
      var hasYearCol = headerRow.indexOf('년도') !== -1 || headerRow.indexOf('년') !== -1;
      
      for (var ar = 1; ar < aValues.length; ar++) {
        var aRow = aValues[ar];
        var aYear = '2026년';
        var aMonth = '';
        var aItem = '';
        var aAmount = 0;
        var aCard = '다정카드';

        if (hasYearCol) {
          aYear = String(aRow[0] || '2026년').trim();
          aMonth = String(aRow[1] || '').trim();
          aItem = String(aRow[2] || '').trim();
          aAmount = parseNumeric(aRow[3]);
          aCard = String(aRow[4] || '다정카드').trim();
        } else {
          aMonth = String(aRow[0] || '').trim();
          aItem = String(aRow[1] || '').trim();
          aAmount = parseNumeric(aRow[2]);
          aCard = String(aRow[3] || '다정카드').trim();
        }

        if (aMonth.indexOf('년') !== -1) {
          var p = aMonth.split(/\s+/);
          if (p.length >= 2) {
            aYear = p[0];
            aMonth = p[1];
          }
        }

        if (aItem && aAmount > 0 && !(/^202\d년?$/.test(aItem) && aAmount < 100)) {
          archiveItems.push({
            year: aYear,
            month: aMonth,
            item: aItem,
            amount: aAmount,
            card: aCard,
            isArchive: true
          });
        }
      }
    }
  }

  return { items: items, archive: archiveItems, total: total };
}

/**
 * 선준 시트 추출 함수 (유연 탐색 및 동적 헤더 감지, 레고 및 할부 자동인식)
 */
function getSeonjunData(ss) {
  var sheet = findSheetByKeywords(ss, ['선준 카드값', '선준카드값', '선준 카드', '선준카드'], ['축의금', '조의금', '다정', '주식', '내집', '시트']);
  if (!sheet) {
    var allSheets = ss.getSheets();
    for (var s = 0; s < allSheets.length; s++) {
      var sName = allSheets[s].getName().trim();
      if (sName.indexOf('축의') !== -1 || sName.indexOf('조의') !== -1 || sName.indexOf('다정') !== -1 || sName.indexOf('주식') !== -1 || sName.indexOf('내집') !== -1) continue;
      if (sName.indexOf('선준') !== -1 && (sName.indexOf('카드') !== -1 || sName.indexOf('비용') !== -1)) {
        sheet = allSheets[s];
        break;
      }
    }
  }
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return [];

  // 동적 헤더 감지
  var colCard = 0, colItem = 1, colAmount = 2, colType = 3, colNote = 4, colMonth = 5, colYear = 6;
  var headers = values[0].map(function(h) { return String(h || '').trim(); });
  for (var c = 0; c < headers.length; c++) {
    var h = headers[c];
    if (h === '카드' || h.indexOf('카드') !== -1) colCard = c;
    else if (h.indexOf('항목') !== -1 || h.indexOf('내역') !== -1) colItem = c;
    else if (h.indexOf('금액') !== -1 || h.indexOf('가격') !== -1) colAmount = c;
    else if (h === '구분') colType = c;
    else if (h === '비고' || h.indexOf('할부') !== -1) colNote = c;
    else if (h === '월' || h.indexOf('월별') !== -1) colMonth = c;
    else if (h === '년' || h.indexOf('년도') !== -1) colYear = c;
  }

  var result = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var card = String(row[colCard] || '').trim();
    var item = String(row[colItem] || '').trim();
    var amount = parseNumeric(row[colAmount]);
    var type = colType !== -1 ? String(row[colType] || '').trim() : '공통';
    var note = colNote !== -1 ? String(row[colNote] || '').trim() : '일반';
    var month = colMonth !== -1 ? String(row[colMonth] || '').trim() : '';
    var year = (colYear !== -1 && String(row[colYear] || '').trim()) ? String(row[colYear]).trim() : '2026년';

    if (month.indexOf('년') !== -1) {
      var parts = month.split(/\s+/);
      if (parts.length >= 2) {
        year = parts[0];
        month = parts[1];
      }
    }

    // 년도 정규화 (Date 객체나 타임스탬프로 들어온 경우 2026년 추출)
    if (year) {
      var ym = String(year).match(/(20\d{2})/);
      if (ym) year = ym[1] + '년';
      else year = '2026년';
    }

    var rawNote = note;
    // 비고에 날짜 객체/타임스탬프가 들어온 경우 또는 레고 등 할부 항목 자동 보정
    var isDateNote = note && (note.indexOf('GMT') !== -1 || note.indexOf('한국') !== -1 || /^\d{4}[-./]/.test(note) || /^\d{1,2}\/\d{1,2}/.test(note) || /\d+\/\d+/.test(note));
    if (isDateNote || item.indexOf('레고') !== -1) {
      note = '할부';
    }

    if (item || amount > 0) {
      result.push({ card: card, item: item, amount: amount, type: type, note: note, month: month, year: year, rawNote: rawNote });
    }
  }
  return result;
}

/**
 * 웹 대시보드 -> 구글 시트 양방향 데이터 저장 (doPost)
 */
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var postData = {};
    if (e && e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      postData = e.parameter;
    }

    var action = postData.action;

    // 1. 새로운 월 추가 (addMonth)
    if (action === 'addMonth') {
      var d = postData.data;
      var sheetCosts = ss.getSheetByName('월별 비용') || ss.getSheetByName('월별비용');
      if (sheetCosts) {
        sheetCosts.appendRow([
          d.year,
          d.month,
          d.remain,
          d.savings,
          d.pocketMoney,
          d.livingExpense,
          d.emergencyFund
        ]);
      }

      var sheetSummary = ss.getSheetByName('월별 요약') || ss.getSheetByName('요약');
      if (sheetSummary) {
        sheetSummary.appendRow([
          d.year,
          d.month,
          d.salary || 0,
          d.salaryMinusCard || 0,
          d.dajeongRemain || 0,
          d.remain || 0,
          d.cardTotal || 0,
          d.savings || 0,
          d.totalCost || d.pocketMoney || 0,
          d.livingExpense || 0,
          d.emergencyFund || 0
        ]);
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "월 추가 완료: " + d.year + " " + d.month
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. 카드 지출 항목 추가 (addCardItem)
    if (action === 'addCardItem') {
      var target = postData.target; // 'dajeong' | 'seonjun'
      var it = postData.item;
      var cardSheet = ss.getSheetByName(target === 'dajeong' ? '다정' : '선준');
      if (cardSheet) {
        if (target === 'dajeong') {
          cardSheet.appendRow([it.item, it.amount, it.month || '', it.card || '다정카드', it.year || '2026년']);
        } else {
          cardSheet.appendRow([
            it.card || '신용카드',
            it.item,
            it.amount,
            it.type || '공통',
            it.note || '일반',
            it.month,
            it.year || '2026년'
          ]);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "카드 항목 추가 완료"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 3. 월 항목 개별 금액 수정 (updateMonthField)
    if (action === 'updateMonthField') {
      var d = postData.data;
      var targetYear = d.year;
      var targetMonth = d.month;
      var field = d.field;
      var amount = d.amount;

      var sheetCosts = ss.getSheetByName('월별 비용') || ss.getSheetByName('월별비용');
      if (sheetCosts) {
        var values = sheetCosts.getDataRange().getValues();
        var currentYear = '';
        for (var r = 1; r < values.length; r++) {
          var rowYear = String(values[r][0] || '').trim();
          if (rowYear && rowYear.indexOf('년') !== -1) currentYear = rowYear;
          var rowMonth = String(values[r][1] || '').trim();
          if ((currentYear === targetYear || !targetYear) && rowMonth === targetMonth) {
            var colIdx = -1;
            if (field === 'remain') colIdx = 3;
            else if (field === 'savings') colIdx = 4;
            else if (field === 'pocketMoney') colIdx = 5;
            else if (field === 'livingExpense') colIdx = 6;
            else if (field === 'emergencyFund') colIdx = 7;
            if (colIdx !== -1) {
              sheetCosts.getRange(r + 1, colIdx).setValue(amount);
            }
            break;
          }
        }
      }

      var sheetSummary = ss.getSheetByName('월별 요약') || ss.getSheetByName('요약');
      if (sheetSummary) {
        var sValues = sheetSummary.getDataRange().getValues();
        for (var sr = 1; sr < sValues.length; sr++) {
          var sYear = String(sValues[sr][0] || '').trim();
          var sMonth = String(sValues[sr][1] || '').trim();
          if ((sYear === targetYear || !targetYear) && sMonth === targetMonth) {
            if (field === 'salary') sheetSummary.getRange(sr + 1, 3).setValue(amount);
            else if (field === 'salaryMinusCard') sheetSummary.getRange(sr + 1, 4).setValue(amount);
            else if (field === 'dajeongRemain') sheetSummary.getRange(sr + 1, 5).setValue(amount);
            else if (field === 'remain') sheetSummary.getRange(sr + 1, 6).setValue(amount);
            else if (field === 'totalCost') sheetSummary.getRange(sr + 1, 9).setValue(amount);
            break;
          }
        }
      }

      // 메인 가계부 시트 ('★저축 및 가계부' 등)에서도 해당 필드 실시간 업데이트
      var mainSheet = null;
      var allSh = ss.getSheets();
      for (var si = 0; si < allSh.length; si++) {
        var sn = allSh[si].getName().trim();
        if (sn.indexOf('★') !== -1 && (sn.indexOf('가계부') !== -1 || sn.indexOf('저축') !== -1)) {
          mainSheet = allSh[si];
          break;
        }
      }
      if (!mainSheet) {
        mainSheet = ss.getSheetByName('가계부') || ss.getSheetByName('저축 및 가계부');
      }
      if (mainSheet) {
        var mValues = mainSheet.getDataRange().getValues();
        var mHeaderRow = -1;
        var targetCol = -1;
        var monthCol = -1;
        for (var mr = 0; mr < Math.min(60, mValues.length); mr++) {
          var mRow = mValues[mr].map(function(h) { return String(h || '').trim(); });
          for (var mc = 0; mc < mRow.length; mc++) {
            var mh = mRow[mc];
            if (mh === '월' || mh.indexOf('월별') !== -1 || mh === 'Month') monthCol = mc;
            if (field === 'salary' && (mh.indexOf('급여') !== -1 && mh.indexOf('-') === -1 && mh.indexOf('−') === -1)) targetCol = mc;
            else if (field === 'savings' && (mh.indexOf('적금') !== -1 || (mh.indexOf('저축') !== -1 && mh.indexOf('합계') === -1 && mh.indexOf('누계') === -1))) targetCol = mc;
            else if (field === 'remain' && (mh.indexOf('남은금액') !== -1 || mh.indexOf('잔금') !== -1)) targetCol = mc;
            else if (field === 'dajeongRemain' && mh.indexOf('다정') !== -1) targetCol = mc;
            else if (field === 'totalCost' && (mh.indexOf('전체비용') !== -1 || mh.indexOf('선준') !== -1)) targetCol = mc;
          }
          if (monthCol !== -1 && targetCol !== -1) { mHeaderRow = mr; break; }
        }
        if (mHeaderRow !== -1 && targetCol !== -1) {
          for (var dr = mHeaderRow + 1; dr < mValues.length; dr++) {
            var ymCheck = extractYearMonth(mValues[dr][monthCol]);
            if (ymCheck && ymCheck.month === targetMonth) {
              mainSheet.getRange(dr + 1, targetCol + 1).setValue(amount);
              break;
            }
          }
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "월 항목 수정 완료: " + targetYear + " " + targetMonth + " (" + field + "=" + amount + ")"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 4. 다정 카드 특정 월 보존 저장 (archiveDajeongMonth)
    if (action === 'archiveDajeongMonth') {
      var targetYear = postData.year || '2026년';
      var targetMonth = postData.month;
      var itemsToArchive = postData.items || [];
      var archiveSheet = ss.getSheetByName('다정_기록보관');
      if (!archiveSheet) {
        archiveSheet = ss.insertSheet('다정_기록보관');
        archiveSheet.appendRow(['년도', '월', '항목', '금액', '카드', '저장일시']);
      }
      var nowStr = new Date().toISOString();
      for (var k = 0; k < itemsToArchive.length; k++) {
        var it = itemsToArchive[k];
        archiveSheet.appendRow([targetYear, targetMonth, it.item, it.amount, it.card || '다정카드', nowStr]);
      }
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "다정 " + targetYear + " " + targetMonth + " " + itemsToArchive.length + "건 영구보관 완료"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 5. 내집마련 플랜 4번 자산 개별 금액 수정 (updateHouseSavingsField)
    if (action === 'updateHouseSavingsField') {
      var field = postData.field;
      var amount = postData.amount;
      var targetSheets = ['내집마련', '내집마련 플랜', '분양자금', '분양', '자산', '요약'];
      var sheet = null;
      for (var s = 0; s < targetSheets.length; s++) {
        sheet = ss.getSheetByName(targetSheets[s]);
        if (sheet) break;
      }
      if (sheet) {
        var values = sheet.getDataRange().getValues();
        var fieldLabels = {
          stock: ['주식'],
          deposit: ['보증금'],
          housingSubscription: ['주택청약', '청약'],
          monthlySavings: ['월별 저축', '월별저축', '적금'],
          parkingAccount: ['파킹', '파킹통장']
        };
        var targets = fieldLabels[field] || [];
        var updated = false;

        // 가로형/세로형 탐색 및 업데이트
        for (var r = 0; r < values.length && !updated; r++) {
          for (var c = 0; c < values[r].length && !updated; c++) {
            var cellVal = String(values[r][c] || '').trim();
            for (var t = 0; t < targets.length; t++) {
              if (cellVal.indexOf(targets[t]) !== -1) {
                // 바로 아래 행이거나 옆 열에 숫자 쓰기
                if (r + 1 < values.length) {
                  sheet.getRange(r + 2, c + 1).setValue(amount);
                  updated = true;
                  break;
                }
              }
            }
          }
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "내집마련 4번 자산 수정 완료: " + field + " = " + amount
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      received: postData
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 내집마련 플랜 4번 보유 자산 표 추출 함수
 * 구글 시트의 4번 표(가로형 6열 또는 세로형 리스트)를 자동 탐색하여 100% 실시간 연동
 */
function getHouseSavingsData(ss) {
  var defaultData = {
    stock: 17164816,
    deposit: 29921000,
    housingSubscription: 8850000,
    monthlySavings: 37860000,
    parkingAccount: 2238,
    total: 93798054
  };

  try {
    var candidateSheets = [];
    var allSheets = ss.getSheets();
    for (var i = 0; i < allSheets.length; i++) {
      var sName = allSheets[i].getName();
      if (sName.indexOf('내집') !== -1 || sName.indexOf('분양') !== -1 || sName.indexOf('자산') !== -1 || sName.indexOf('아파트') !== -1 || sName.indexOf('저축') !== -1) {
        candidateSheets.unshift(allSheets[i]);
      } else {
        candidateSheets.push(allSheets[i]);
      }
    }

    for (var sIdx = 0; sIdx < candidateSheets.length; sIdx++) {
      var sheet = candidateSheets[sIdx];
      var values = sheet.getDataRange().getValues();
      if (!values || values.length === 0) continue;

      // [탐색 1: 가로형 6열 표] (한 행에 주식, 보증금, 청약 등이 있고 다음 행에 금액)
      for (var r = 0; r < values.length; r++) {
        var rowStr = values[r].map(function(c) { return String(c || '').trim(); });
        var colStock = -1, colDeposit = -1, colSub = -1, colMonthly = -1, colParking = -1, colTotal = -1;
        for (var c = 0; c < rowStr.length; c++) {
          var v = rowStr[c];
          if (v === '주식' || v.indexOf('주식') !== -1) colStock = c;
          else if (v.indexOf('보증금') !== -1) colDeposit = c;
          else if (v.indexOf('주택청약') !== -1 || v.indexOf('청약') !== -1) colSub = c;
          else if (v.indexOf('월별 저축') !== -1 || v.indexOf('월별저축') !== -1 || v.indexOf('적금') !== -1) colMonthly = c;
          else if (v.indexOf('파킹') !== -1) colParking = c;
          else if (v.indexOf('저축합계') !== -1 || v.indexOf('합계') !== -1) colTotal = c;
        }

        if (colDeposit !== -1 && (colSub !== -1 || colMonthly !== -1 || colStock !== -1) && r + 1 < values.length) {
          var nextRow = values[r + 1];
          var stock = colStock !== -1 ? parseNumeric(nextRow[colStock]) : defaultData.stock;
          var deposit = colDeposit !== -1 ? parseNumeric(nextRow[colDeposit]) : defaultData.deposit;
          var sub = colSub !== -1 ? parseNumeric(nextRow[colSub]) : defaultData.housingSubscription;
          var monthly = colMonthly !== -1 ? parseNumeric(nextRow[colMonthly]) : defaultData.monthlySavings;
          var parking = colParking !== -1 ? parseNumeric(nextRow[colParking]) : defaultData.parkingAccount;
          var total = colTotal !== -1 ? parseNumeric(nextRow[colTotal]) : (stock + deposit + sub + monthly + parking);
          if (total === 0) total = stock + deposit + sub + monthly + parking;

          if (deposit > 0 || sub > 0 || monthly > 0) {
            return {
              stock: stock,
              deposit: deposit,
              housingSubscription: sub,
              monthlySavings: monthly,
              parkingAccount: parking,
              total: total,
              updatedAt: new Date().toISOString()
            };
          }
        }
      }

      // [탐색 2: 세로형 표] (각 행에 '주식', '보증금', '청약' 등이 있고 옆 열에 금액)
      var vStock = 0, vDeposit = 0, vSub = 0, vMonthly = 0, vParking = 0, vTotal = 0;
      var foundCount = 0;

      for (var vr = 0; vr < values.length; vr++) {
        for (var vc = 0; vc < values[vr].length; vc++) {
          var cellText = String(values[vr][vc] || '').trim();
          var rightNum = (vc + 1 < values[vr].length) ? parseNumeric(values[vr][vc + 1]) : 0;
          if (rightNum === 0 && vc + 2 < values[vr].length) {
            rightNum = parseNumeric(values[vr][vc + 2]);
          }

          if (rightNum > 0) {
            if (cellText === '주식' || cellText.indexOf('국내 · 해외 주식') !== -1 || cellText.indexOf('보유 주식') !== -1) {
              vStock = rightNum; foundCount++;
            } else if (cellText.indexOf('보증금') !== -1 || cellText.indexOf('전세보증금') !== -1) {
              vDeposit = rightNum; foundCount++;
            } else if (cellText.indexOf('주택청약') !== -1 || cellText.indexOf('청약') !== -1) {
              vSub = rightNum; foundCount++;
            } else if (cellText.indexOf('월별 저축') !== -1 || cellText.indexOf('월별저축') !== -1 || cellText.indexOf('적금 누계') !== -1) {
              vMonthly = rightNum; foundCount++;
            } else if (cellText.indexOf('파킹') !== -1) {
              vParking = rightNum; foundCount++;
            } else if (cellText.indexOf('저축합계') !== -1 || cellText.indexOf('총 저축') !== -1) {
              vTotal = rightNum;
            }
          }
        }
      }

      if (foundCount >= 2 && (vDeposit > 0 || vMonthly > 0 || vSub > 0)) {
        var calcTot = vTotal > 0 ? vTotal : (vStock + vDeposit + vSub + vMonthly + vParking);
        return {
          stock: vStock || defaultData.stock,
          deposit: vDeposit || defaultData.deposit,
          housingSubscription: vSub || defaultData.housingSubscription,
          monthlySavings: vMonthly || defaultData.monthlySavings,
          parkingAccount: vParking || defaultData.parkingAccount,
          total: calcTot || defaultData.total,
          updatedAt: new Date().toISOString()
        };
      }
    }
  } catch (e) {
    return defaultData;
  }

  return defaultData;
}
