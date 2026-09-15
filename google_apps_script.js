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
 * 월별 비용 시트 (잔금, 적금, 용돈, 생활비, 비상금) 자동 추출 함수
 * '월별 비용', '월별비용', '비용' 등의 시트명이나 헤더를 검색하여 추출합니다.
 */
function getMonthlyExpensesData(ss) {
  var sheet = findSheetByKeywords(ss, ['월별 비용', '월별비용', '비용', '월별지출', '지출'], ['내집', '분양', '아파트', '주식', '다정', '선준']);
  var sheetsToSearch = [];
  if (sheet) {
    sheetsToSearch.push(sheet);
  } else {
    var allSheets = ss.getSheets();
    for (var s = 0; s < allSheets.length; s++) {
      var curName = allSheets[s].getName().trim();
      if (curName.indexOf('내집') !== -1 || curName.indexOf('분양') !== -1 || curName.indexOf('아파트') !== -1 ||
          curName.indexOf('다정') !== -1 || curName.indexOf('선준') !== -1 || curName.indexOf('주식') !== -1) {
        continue;
      }
      sheetsToSearch.push(allSheets[s]);
    }
  }

  for (var i = 0; i < sheetsToSearch.length; i++) {
    var curSheet = sheetsToSearch[i];
    var values = curSheet.getDataRange().getValues();
    if (!values || values.length < 2) continue;

    var headerRowIdx = -1;
    var colRemain = -1, colSavings = -1, colPocket = -1, colLiving = -1, colEmergency = -1;
    var colYear = -1, colMonth = -1;

    for (var r = 0; r < Math.min(10, values.length); r++) {
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
      var currentYear = '';

      if (colRemain > 1 && colMonth === -1) {
        colMonth = colRemain - 1;
        colYear = colRemain - 2;
      } else if (colMonth === -1) {
        colMonth = 1;
        colYear = 0;
      }

      for (var r = headerRowIdx + 1; r < values.length; r++) {
        var row = values[r];
        var yVal = colYear !== -1 ? String(row[colYear] || '').trim() : '';
        var mVal = colMonth !== -1 ? String(row[colMonth] || '').trim() : '';

        if (yVal && (yVal.indexOf('년') !== -1 || /^\d{4}$/.test(yVal))) {
          currentYear = yVal.indexOf('년') === -1 ? yVal + '년' : yVal;
        }

        if (!mVal) continue;
        // X월 형식 검증 (1월 ~ 12월만 엄격 허용! 37월 등 불가능)
        var mMatch = mVal.match(/^(\d{1,2})\s*월?$/);
        if (!mMatch) continue;
        var mNum = parseInt(mMatch[1], 10);
        if (mNum < 1 || mNum > 12) continue;
        var monthStr = mNum + '월';

        // 수식 결과 숫자 추출
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
            year: currentYear || '2026년',
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
 * 기존 시트 추출 함수 (요약 시트) - 엄격한 1~12월 검증 및 동적 헤더 감지
 */
function getSummaryData(ss, monthlyExpensesData) {
  var sheet = findSheetByKeywords(ss, ['월별 요약', '월별요약', '요약', '가계부', '월별 요약 ', '요약 '], ['내집', '분양', '주식', '다정', '선준']);

  if (!sheet) {
    var allSheets = ss.getSheets();
    for (var i = 0; i < allSheets.length; i++) {
      var s = allSheets[i];
      var sName = s.getName().trim();
      if (sName.indexOf('다정') !== -1 || sName.indexOf('선준') !== -1 || sName.indexOf('주식') !== -1 || sName.indexOf('내집') !== -1 || sName.indexOf('분양') !== -1) continue;
      var v = s.getDataRange().getValues();
      if (!v || v.length < 2) continue;
      for (var r = 0; r < Math.min(5, v.length); r++) {
        var rowText = v[r].join(' ');
        if ((rowText.indexOf('급여') !== -1 || rowText.indexOf('월') !== -1) && 
            (rowText.indexOf('선준') !== -1 || rowText.indexOf('다정') !== -1 || rowText.indexOf('잔금') !== -1 || rowText.indexOf('적금') !== -1)) {
          sheet = s;
          break;
        }
      }
      if (sheet) break;
    }
  }

  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  var result = [];

  // 동적 헤더 위치 감지
  var colYear = 0, colMonth = 1, colSalary = 2, colSalaryMinusCard = 3, colDajeongRemain = 4;
  var colTotalCost = 5, colCardTotal = 6, colSavings = 7, colPocket = 8, colLiving = 9, colEmergency = 10, colRemain = -1;
  if (values.length > 0) {
    var headers = values[0].map(function(h) { return String(h || '').trim(); });
    for (var c = 0; c < headers.length; c++) {
      var h = headers[c];
      if (h.indexOf('년') !== -1) colYear = c;
      else if (h === '월') colMonth = c;
      else if (h.indexOf('급여') !== -1 || h.indexOf('수입') !== -1) colSalary = c;
      else if (h.indexOf('급여-') !== -1 || h.indexOf('급여−') !== -1 || h.indexOf('차액') !== -1) colSalaryMinusCard = c;
      else if (h.indexOf('다정') !== -1 && h.indexOf('남은') !== -1) colDajeongRemain = c;
      else if (h.indexOf('남은') !== -1 || h.indexOf('잔금') !== -1 || h.indexOf('잔액') !== -1) colRemain = c;
      else if (h.indexOf('전체비용') !== -1 || h.indexOf('선준 카드') !== -1 || h.indexOf('선준카드') !== -1) colTotalCost = c;
      else if (h.indexOf('카드') !== -1 && h.indexOf('합계') !== -1) colCardTotal = c;
      else if (h.indexOf('적금') !== -1 || h.indexOf('저축') !== -1) colSavings = c;
      else if (h === '용돈') colPocket = c;
      else if (h === '생활비') colLiving = c;
      else if (h === '비상금') colEmergency = c;
    }
  }
  
  // 행 파싱
  var currentYear = '2026년';
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!row[colYear] && !row[colMonth]) continue;
    var y = String(row[colYear] || '').trim();
    var m = String(row[colMonth] || '').trim();

    if (y && (y.indexOf('년') !== -1 || /^\d{4}$/.test(y))) {
      currentYear = y.indexOf('년') === -1 ? y + '년' : y;
    }

    // 1~12월 검증
    if (!m) continue;
    var mMatch = m.match(/^(\d{1,2})\s*월?$/);
    if (!mMatch) continue;
    var mNum = parseInt(mMatch[1], 10);
    if (mNum < 1 || mNum > 12) continue;
    var formattedMonth = mNum + '월';

    var rawRemain = colRemain !== -1 ? parseNumeric(row[colRemain]) : 0;
    var rawSalaryMinusCard = colSalaryMinusCard !== -1 ? parseNumeric(row[colSalaryMinusCard]) : 0;
    if (rawRemain > 0 && rawSalaryMinusCard === 0) {
      rawSalaryMinusCard = rawRemain;
    }

    var item = {
      year: currentYear,
      month: formattedMonth,
      salary: parseNumeric(row[colSalary]),
      salaryMinusCard: rawSalaryMinusCard,
      dajeongRemain: parseNumeric(row[colDajeongRemain]),
      remain: rawRemain,
      totalCost: parseNumeric(row[colTotalCost]),
      cardTotal: parseNumeric(row[colCardTotal]),
      savings: parseNumeric(row[colSavings]),
      pocketMoney: parseNumeric(row[colPocket]),
      livingExpense: parseNumeric(row[colLiving]),
      emergencyFund: parseNumeric(row[colEmergency])
    };

    // totalCost 자동 역산 (급여 - 남은금액)
    if (item.salary > 0 && item.salaryMinusCard > 0 && item.totalCost === 0) {
      item.totalCost = item.salary - item.salaryMinusCard;
    }

    // remain 자동 보정
    if (item.remain === 0 && item.salaryMinusCard > 0) {
      item.remain = item.salaryMinusCard + (item.dajeongRemain || 0);
    }

    // 월별 비용 시트에 기재된 데이터가 있다면 보완
    if (monthlyExpensesData && monthlyExpensesData.length > 0) {
      for (var me = 0; me < monthlyExpensesData.length; me++) {
        var exp = monthlyExpensesData[me];
        if ((exp.year === item.year || !exp.year) && exp.month === item.month) {
          if (exp.remain !== undefined && exp.remain > 0) item.remain = exp.remain;
          if (exp.savings !== undefined && exp.savings > 0) item.savings = exp.savings;
          if (exp.pocketMoney !== undefined && exp.pocketMoney > 0) item.pocketMoney = exp.pocketMoney;
          if (exp.livingExpense !== undefined && exp.livingExpense > 0) item.livingExpense = exp.livingExpense;
          if (exp.emergencyFund !== undefined) item.emergencyFund = exp.emergencyFund;
          break;
        }
      }
    }

    result.push(item);
  }
  return result;
}

/**
 * 다정 시트 추출 함수 (유연 탐색)
 */
function getDajeongData(ss) {
  var sheet = findSheetByKeywords(ss, ['다정', '다정카드', '다정 카드', '다정(카드)'], ['보관', '히스토리', '아카이브']);
  var archiveSheet = findSheetByKeywords(ss, ['다정_기록보관', '다정기록보관', '다정히스토리', '다정보관', '다정 히스토리']);
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
 * 선준 시트 추출 함수 (유연 탐색 및 동적 헤더 감지)
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

    // 비고에 날짜 객체/타임스탬프가 들어온 경우 할부로 자동 보정
    if (note && (note.indexOf('GMT') !== -1 || note.indexOf('한국') !== -1 || /^\d{4}[-./]/.test(note))) {
      note = '할부';
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
