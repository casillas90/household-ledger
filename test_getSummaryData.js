// Removed require

// The raw data from the curl output
const sheetData = {"name":"★저축 및 가계부","totalRows":50,"totalCols":25,"rows":[["","","","","","","","","","","","","","","","","","","",""],["","","","","","","","","","","","","","","","","","","",""],["","","","계약금(5%)","","중도금(60%)","","","","","","잔금(35%)","","","","","","","",""],["","유형","분양가","계약시","계약 후 30일이내","중도금1회\n25-04-25","중도금2회\n25-10-28","중도금3회\n26-04-27","중도금4회\n26-11-25","중도금5회\n27-04-27","중도금6회\n27-09-27","입주예정일","","","","","주식","보증금","주택청약","월별 저축"],["","71","629000000","10000000","21450000","62900000","62900000","62900000","62900000","62900000","62900000","220150000","","","","","17700844.138240002","29921000","8850000","37860000"],["","","","","","","","","","","","","","","","","","","",""],["","","구분","상품명","","세부내역","금액","","중도금","일수","이율","이자","","","","","","분양가-대출-계약금","이자","옵션"],["","","에어컨","","","","","","62900000","1010","0.0392","6822840.547945206","","","","","4억대출","197550000","22434310.356164385","16328000"],["","","","인피니티","선택2","거실+안방+침실1+침실2","11580000","","62900000","826","0.0392","5579867.616438356","","","","","5억대출","97550000","22434310.356164385","16328000"],["","","마감","와이드강마루","선택1","마일드 화이트","2000000","","62900000","645","0.0392","4357160.547945206","","","","","","","",""],["","","","반건식세탁실","","자동중문+우드패턴 타일\n+바닥 난방","2280000","","62900000","433","0.0392","2925039.5616438356","","","","","","","",""],["","","가구","수납강화 화장대","","","1160000","","62900000","280","0.0392","1891480.5479452056","","","","","","","",""],["","","냉장고","냉장고장","","","","","62900000","127","0.0392","857921.5342465754","","","","","","","",""],["","","전기/설비","거실조명","선택1","","1830000","","377400000","","","22434310.356164385","","","","","","","",""],["","","","인덕션","선택1","","870000","","","","","","","","","","","","",""],["","","","레인지후드","","","690000","","","","","","","","","","","","",""],["","","","","","","","","","","","","","","","","","","",""],["","","합계","","","","20410000","","","","","","","","","","","","",""],["","","","","","","16328000","","","","","","","","","","","","",""],["","","","","","","","","","","","","","","","","","","",""],["","","","","","","","","","","","","","","","","","","",""],["","","","","","","","","","","","","","","","","","","",""],["전체비용","","급여","급여-카드결제금액","다정남은금액","잔금","","","카드결제금액","","하나카드","롯데카드","신한카드","기업카드","국민카드","합계","","월별 비용","","잔금"],["2024년","10월","3463220","2465618","852641","3318259","","","2024년","10월","549477","169500","128000","31220","119405","997602","","2024년","10월","3320924"],["","11월","3463220","2374694","1024281","3398975","","","","11월","700741","272136","64000","9300","42349","1088526","","","11월","3198975"],["","12월","3463220","2393534","905933","3299467","","","","12월","686691","272173","64000","5500","41322","1069686","","","12월","3149467"],["2025년","1월","3463380","2228510","877153","3105663","","","2025년","1월","820937","271743","64000","11500","66690","1234870","","2025년","1월","3127458"],["","2월","3463220","2375314","665743","3041057","","","","2월","741352","184964","76000","48900","36690","1087906","","","2월","3072434"],["","3월","3775420","2760105","956341","3716446","","","","3월","740080","169545","64000","5000","36690","1015315","","","3월","3622378"],["","4월","3028660","2151877","909666","3061543","","","","4월","688593","79200","64000","8300","36690","876783","","","4월","3069506"],["","5월","3239940","2389525","785767","3175292","","","","5월","706015","79200","64000","1200","","850415","","","5월","3178516"],["","6월","3412020","2555957","691967","3247924","","","","6월","709263","79200","64000","3600","","856063","","","6월","3257113"],["","7월","3132500","2250782","661685","2912467","","","","7월","791618","18500","51500","20100","","881718","","","7월","2912467"],["","8월","3300930","2560766","624266","3185032","","","","8월","546144","121920","51500","20600","","740164","","","8월","3267201"],["","9월","3544060","2500128","808628","3308756","","","","9월","890332","92300","51500","9800","","1043932","","","9월","3309226"],["","10월","3544060","2734155","712028","3446183","","","","10월","635928","118277","51500","4200","","809905","","","10월","3451687"],["","11월","3544060","2666798","","2666798","","","","11월","652049","155213","51500","18500","","877262","","","11월","1782887"],["","12월","3544060","2239954","","2239954","","","","12월","904339","347967","51500","300","","1304106","","","12월","5711461"],["2026년","1월","3544060","2285463","925152","3210615","","","2026년","1월","850297","356200","51500","300","300","1258597","","2026년","1월","3210615"],["","2월","3939120","2677806","2340000","5017806","","","","2월","881814","311000","63500","5000","","1261314","","","2월","5017806"]]};

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

function extractYearMonth(val, defaultYear) {
  if (!val) return null;
  var dYear = defaultYear || '2026년';
  if (val instanceof Date || (typeof val === 'object' && typeof val.getMonth === 'function')) {
    var y = val.getFullYear();
    var m = val.getMonth() + 1;
    if (m >= 1 && m <= 12) {
      return { year: y + '년', month: m + '월' };
    }
  }
  var str = String(val).trim();
  if (!str) return null;
  if (str.indexOf('GMT') !== -1 || str.indexOf('한국') !== -1) {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      return { year: d.getFullYear() + '년', month: (d.getMonth() + 1) + '월' };
    }
  }
  var yPart = '';
  var ymMatch = str.match(/(20\d{2})/);
  if (ymMatch) yPart = ymMatch[1] + '년';
  var ydashm = str.match(/20\d{2}[-./](\d{1,2})/);
  if (ydashm) {
    var mn = parseInt(ydashm[1], 10);
    if (mn >= 1 && mn <= 12) {
      return { year: yPart || dYear, month: mn + '월' };
    }
  }
  var mMatch = str.match(/^(\d{1,2})\s*월?$/) || str.match(/(\d{1,2})\s*월/);
  if (mMatch) {
    var mn = parseInt(mMatch[1], 10);
    if (mn >= 1 && mn <= 12) {
      return { year: yPart || dYear, month: mn + '월' };
    }
  }
  return null;
}

function test() {
  var values = sheetData.rows;
  var headerRowIdx = -1;
  var colYear = -1, colMonth = -1, colSalary = -1, colSalaryMinusCard = -1, colDajeongRemain = -1;
  var colTotalCost = -1, colCardTotal = -1, colSavings = -1, colPocket = -1, colLiving = -1, colEmergency = -1, colRemain = -1, colRemain2 = -1;

  for (var r = 0; r < Math.min(30, values.length); r++) {
    var row = values[r].map(function(h) { return String(h || '').trim(); });
    var foundMatches = 0;
    var tYear = -1, tMonth = -1, tSalary = -1, tSalaryMinusCard = -1, tDajeongRemain = -1;
    var tTotalCost = -1, tCardTotal = -1, tSavings = -1, tPocket = -1, tLiving = -1, tEmergency = -1, tRemain = -1, tRemain2 = -1;

    var tMonthlyStart = -1;
    for (var cs = 0; cs < row.length; cs++) {
      if (row[cs] && row[cs].indexOf('월별') !== -1 && row[cs].indexOf('비용') !== -1) { tMonthlyStart = cs; break; }
    }

    for (var c = 0; c < row.length; c++) {
      var h = row[c];
      if (!h) continue;
      if (tMonthlyStart !== -1 && c >= tMonthlyStart) {
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
      else if (h.indexOf('급여-') !== -1 || h.indexOf('급여−') !== -1 || h.indexOf('급여 -') !== -1 || h.indexOf('차액') !== -1 || (h.indexOf('급여') !== -1 && h.indexOf('카드') !== -1)) { tSalaryMinusCard = c; foundMatches++; }
      else if (h.indexOf('급여') !== -1 || h.indexOf('수입') !== -1 || h.indexOf('월급') !== -1) { tSalary = c; foundMatches++; }
      else if (h.indexOf('다정') !== -1) { tDajeongRemain = c; foundMatches++; }
      else if (h.indexOf('남은금액') !== -1 || h.indexOf('남은 금액') !== -1 || h.indexOf('잔금') !== -1 || h.indexOf('잔액') !== -1 || h.indexOf('남은') !== -1) { tRemain = c; foundMatches++; }
      else if (h.indexOf('전체비용') !== -1 || h.indexOf('선준 카드') !== -1 || h.indexOf('선준카드') !== -1 || (h.indexOf('선준') !== -1 && h.indexOf('지출') !== -1)) { tTotalCost = c; foundMatches++; }
      else if (h.indexOf('카드') !== -1 && (h.indexOf('합계') !== -1 || h.indexOf('총액') !== -1)) { tCardTotal = c; foundMatches++; }
      else if ((h.indexOf('적금') !== -1 || h.indexOf('저축') !== -1) && h.indexOf('합계') === -1 && h.indexOf('누계') === -1 && h.indexOf('총') === -1) { tSavings = c; foundMatches++; }
      else if (h.indexOf('용돈') !== -1 || (h.indexOf('선준') !== -1 && h.indexOf('용돈') !== -1)) { tPocket = c; foundMatches++; }
      else if (h.indexOf('생활비') !== -1) { tLiving = c; foundMatches++; }
      else if (h.indexOf('비상금') !== -1) { tEmergency = c; foundMatches++; }
    }

    if (foundMatches >= 2 && (tMonth !== -1 || tSalary !== -1 || tSalaryMinusCard !== -1 || tRemain !== -1 || tSavings !== -1 || tDajeongRemain !== -1)) {
      console.log('Found Header at row: ' + r, 'foundMatches:', foundMatches);
      headerRowIdx = r;
      colYear = tYear; colMonth = tMonth; colSalary = tSalary; colSalaryMinusCard = tSalaryMinusCard;
      colDajeongRemain = tDajeongRemain; colTotalCost = tTotalCost; colCardTotal = tCardTotal;
      colSavings = tSavings; colPocket = tPocket; colLiving = tLiving; colEmergency = tEmergency; colRemain = tRemain;
      colRemain2 = tRemain2;
      break;
    }
  }

  console.log('headerRowIdx:', headerRowIdx);

  if (headerRowIdx !== -1) {
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

    var yearColByLayout = colMonth > 0 ? colMonth - 1 : -1;
    if (colTotalCost !== -1 && colTotalCost === yearColByLayout) colTotalCost = -1;

    console.log('colMonth:', colMonth, 'yearColByLayout:', yearColByLayout, 'colTotalCost:', colTotalCost);

    var result = [];
    var currentYear = '2026년';

    for (var r = headerRowIdx + 1; r < values.length; r++) {
      var row = values[r];
      var mCell = colMonth !== -1 ? row[colMonth] : null;
      var yCell = colYear !== -1 ? row[colYear] : null;

      var layoutYearCell = yearColByLayout !== -1 ? String(row[yearColByLayout] || '') : '';
      var layoutYearMatch = layoutYearCell.match(/(20\d{2})\s*년?/);
      if (layoutYearMatch) currentYear = layoutYearMatch[1] + '년';

      var ym = extractYearMonth(mCell, currentYear) || extractYearMonth(yCell, currentYear);
      if (!ym) {
        // console.log('Row ' + r + ': Skipped due to no ym. mCell=', mCell, 'currentYear=', currentYear);
        continue;
      }
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

      if (rawSalary > 0 && rawTotalCost > 0 && rawSalaryMinusCard === 0) {
        rawSalaryMinusCard = rawSalary - rawTotalCost;
      }
      if (rawSalary > 0 && rawSalaryMinusCard > 0 && rawTotalCost === 0) {
        rawTotalCost = rawSalary - rawSalaryMinusCard;
      }
      if (rawRemain === 0 && (rawSalaryMinusCard > 0 || rawDajeongRemain > 0)) {
        rawRemain = rawSalaryMinusCard + rawDajeongRemain;
      }
      if (rawEmergency === 0 && rawRemain > 0 && (rawSavings > 0 || rawPocket > 0 || rawLiving > 0)) {
        rawEmergency = rawRemain - (rawSavings + rawPocket + rawLiving);
      }

      var item = {
        year: currentYear,
        month: ym.month,
        salary: rawSalary,
        salaryMinusCard: rawSalaryMinusCard,
        dajeongRemain: rawDajeongRemain,
        remain: rawRemain,
        totalCost: rawTotalCost,
        cardTotal: rawCardTotal,
        savings: rawSavings,
        pocketMoney: rawPocket,
        livingExpense: rawLiving,
        emergencyFund: rawEmergency
      };

      if (item.salary > 0 || item.remain > 0 || item.totalCost > 0 || item.savings > 0 || item.dajeongRemain > 0) {
        result.push(item);
      }
    }
    console.log('Result length:', result.length);
    console.log('First 2 items:', result.slice(0, 2));
  }
}

test();
