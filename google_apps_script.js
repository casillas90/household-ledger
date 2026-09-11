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
        var rawEmergency = colEmergency !== -1 ? row[colEmergency] : null;
        var emergency = 0;
        if (rawEmergency !== null && rawEmergency !== undefined && String(rawEmergency).trim() !== '') {
          emergency = parseNumeric(rawEmergency);
        } else {
          emergency = remain - (savings + pocket + living);
        }

        if (remain > 0 || savings > 0 || pocket > 0 || living > 0 || emergency !== 0) {
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
  var str = String(val).trim();
  // 회계 음수 표기 (1,000) 또는 -1,000 감지
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
 * 기존 시트 추출 함수 (요약 시트)
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
 * 다정 시트 추출 함수 (월별 지원 및 과거 월 영구 보존)
 */
function getDajeongData(ss) {
  var sheet = ss.getSheetByName('다정');
  var archiveSheet = ss.getSheetByName('다정_기록보관') || ss.getSheetByName('다정히스토리');
  if (!sheet && !archiveSheet) return { items: [], total: 0 };

  var items = [];
  var total = 0;

  // 1. 현재 '다정' 시트 읽기
  if (sheet) {
    var values = sheet.getDataRange().getValues();
    if (values && values.length > 0) {
      var colMonth = -1, colItem = 0, colAmount = 1, colCard = -1;
      var header = values[0].map(function(h) { return String(h || '').trim(); });
      for (var c = 0; c < header.length; c++) {
        if (header[c].indexOf('월') !== -1) colMonth = c;
        else if (header[c].indexOf('항목') !== -1 || header[c].indexOf('내역') !== -1) colItem = c;
        else if (header[c].indexOf('금액') !== -1 || header[c].indexOf('가격') !== -1) colAmount = c;
        else if (header[c].indexOf('카드') !== -1) colCard = c;
      }

      for (var r = 1; r < values.length; r++) {
        var item = String(values[r][colItem] || '').trim();
        var amount = parseNumeric(values[r][colAmount]);
        var month = colMonth !== -1 ? String(values[r][colMonth] || '').trim() : '';
        var card = colCard !== -1 ? String(values[r][colCard] || '').trim() : '다정카드';
        
        if (item && amount > 0) {
          items.push({ 
            item: item, 
            amount: amount, 
            month: month || '', 
            card: card 
          });
          total += amount;
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

        if (aItem && aAmount > 0) {
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
    var year = '2026년';

    if (values[r].length > 6 && String(values[r][6] || '').trim()) {
      year = String(values[r][6]).trim();
    } else if (month.indexOf('년') !== -1) {
      var parts = month.split(/\s+/);
      if (parts.length >= 2) {
        year = parts[0];
        month = parts[1];
      }
    }

    if (item || amount > 0) {
      result.push({ card: card, item: item, amount: amount, type: type, note: note, month: month, year: year });
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
          d.totalCost || 0,
          d.cardTotal || d.totalCost || 0,
          d.savings || 0,
          d.pocketMoney || 0,
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
            else if (field === 'totalCost') sheetSummary.getRange(sr + 1, 7).setValue(amount);
            else if (field === 'dajeongRemain') sheetSummary.getRange(sr + 1, 5).setValue(amount);
            break;
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
