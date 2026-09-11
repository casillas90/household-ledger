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
    
    var result = {
      status: "success",
      summary: summaryData,
      monthlyExpenses: monthlyExpensesData,
      dajeong: dajeongData,
      seonjun: seonjunData,
      stocks: stockData
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
 * 월별 비용 시트 (잔금, 적금, 용돈, 생활비, 비상금) 자동 추출 함수
 * '월별 비용', '월별비용', '비용' 등의 시트명이나 헤더를 검색하여 추출합니다.
 */
function getMonthlyExpensesData(ss) {
  var sheet = ss.getSheetByName('월별 비용') || 
              ss.getSheetByName('월별비용') || 
              ss.getSheetByName('비용') ||
              ss.getSheetByName('월별 요약') ||
              ss.getSheetByName('요약');
  
  var sheetsToSearch = sheet ? [sheet] : [];
  var allSheets = ss.getSheets();
  for (var s = 0; s < allSheets.length; s++) {
    if (!sheet || allSheets[s].getName() !== sheet.getName()) {
      sheetsToSearch.push(allSheets[s]);
    }
  }

  for (var i = 0; i < sheetsToSearch.length; i++) {
    var curSheet = sheetsToSearch[i];
    var values = curSheet.getDataRange().getValues();
    if (!values || values.length < 2) continue;

    var headerRowIdx = -1;
    var colRemain = -1, colSavings = -1, colPocket = -1, colLiving = -1, colEmergency = -1;
    var colYear = 0, colMonth = 1;

    for (var r = 0; r < Math.min(10, values.length); r++) {
      var rowStr = values[r].map(function(c) { return String(c || '').trim(); });
      for (var c = 0; c < rowStr.length; c++) {
        var val = rowStr[c];
        if (val === '잔금') colRemain = c;
        else if (val === '적금') colSavings = c;
        else if (val === '용돈') colPocket = c;
        else if (val === '생활비') colLiving = c;
        else if (val === '비상금') colEmergency = c;
      }
      if (colRemain !== -1 && colSavings !== -1 && colPocket !== -1) {
        headerRowIdx = r;
        break;
      }
    }

    if (headerRowIdx !== -1) {
      var results = [];
      var currentYear = '';

      // 잔금 열 위치를 기준으로 이전 열이 월, 그 앞 열이 연도
      if (colRemain > 1) {
        colMonth = colRemain - 1;
        colYear = colRemain - 2;
      }

      for (var r = headerRowIdx + 1; r < values.length; r++) {
        var row = values[r];
        var yVal = String(row[colYear] || '').trim();
        var mVal = String(row[colMonth] || '').trim();

        if (yVal && (yVal.indexOf('년') !== -1 || /^\d{4}$/.test(yVal))) {
          currentYear = yVal.indexOf('년') === -1 ? yVal + '년' : yVal;
        }

        if (!mVal || (mVal.indexOf('월') === -1 && isNaN(parseInt(mVal, 10)))) continue;
        var monthStr = mVal.indexOf('월') === -1 ? mVal + '월' : mVal;

        // 수식 없이 시트 셀에 입력된 순수 숫자만 추출
        var remain = colRemain !== -1 ? parseNumeric(row[colRemain]) : 0;
        var savings = colSavings !== -1 ? parseNumeric(row[colSavings]) : 0;
        var pocket = colPocket !== -1 ? parseNumeric(row[colPocket]) : 0;
        var living = colLiving !== -1 ? parseNumeric(row[colLiving]) : 0;
        var emergency = colEmergency !== -1 ? parseNumeric(row[colEmergency]) : 0;

        // 비상금 셀이 비어있거나 0일 경우: 잔금 - (적금 + 용돈 + 생활비)
        if (emergency === 0 && remain > 0 && (savings > 0 || pocket > 0 || living > 0)) {
          emergency = Math.max(0, remain - (savings + pocket + living));
        }

        if (remain > 0 || savings > 0 || pocket > 0 || living > 0) {
          results.push({
            year: currentYear || '2024년',
            month: monthStr,
            remain: remain,
            savings: savings,
            pocketMoney: pocket,
            livingExpense: living,
            emergencyFund: emergency
          });
        }
      }

      if (results.length > 0) {
        return results;
      }
    }
  }

  return [];
}

function parseNumeric(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  var str = String(val).replace(/[^0-9.-]/g, '');
  var num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

/**
 * 기존 시트 추출 함수 유지 (요약 시트)
 */
function getSummaryData(ss) {
  var sheet = ss.getSheetByName('월별 요약') || ss.getSheetByName('요약') || ss.getSheets()[0];
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  var result = [];
  
  // 헤더 검사 및 행 파싱
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!row[0] && !row[1]) continue;
    result.push({
      year: String(row[0] || ''),
      month: String(row[1] || ''),
      salary: parseNumeric(row[2]),
      salaryMinusCard: parseNumeric(row[3]),
      dajeongRemain: parseNumeric(row[4]),
      totalCost: parseNumeric(row[5]),
      cardTotal: parseNumeric(row[6]),
      savings: parseNumeric(row[7]),
      pocketMoney: parseNumeric(row[8]),
      livingExpense: parseNumeric(row[9]),
      emergencyFund: parseNumeric(row[10])
    });
  }
  return result;
}

/**
 * 다정 시트 추출 함수
 */
function getDajeongData(ss) {
  var sheet = ss.getSheetByName('다정');
  if (!sheet) return { items: [], total: 0 };
  var values = sheet.getDataRange().getValues();
  var items = [];
  var total = 0;
  for (var r = 1; r < values.length; r++) {
    var item = String(values[r][0] || '').trim();
    var amount = parseNumeric(values[r][1]);
    if (item && amount > 0) {
      items.push({ item: item, amount: amount });
      total += amount;
    }
  }
  return { items: items, total: total };
}

/**
 * 선준 시트 추출 함수
 */
function getSeonjunData(ss) {
  var sheet = ss.getSheetByName('선준');
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  var result = [];
  for (var r = 1; r < values.length; r++) {
    var card = String(values[r][0] || '').trim();
    var item = String(values[r][1] || '').trim();
    var amount = parseNumeric(values[r][2]);
    var type = String(values[r][3] || '').trim();
    var note = String(values[r][4] || '').trim();
    var month = String(values[r][5] || '').trim();
    if (item || amount > 0) {
      result.push({ card: card, item: item, amount: amount, type: type, note: note, month: month });
    }
  }
  return result;
}
