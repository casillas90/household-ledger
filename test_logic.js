function run() {
    var app = Application.currentApplication();
    app.includeStandardAdditions = true;
    
    // This is the seonjun data I extracted from the log previously
    var seonjunData = [
        {"card":"하나카드","item":"아이폰","amount":127226,"type":"공통","note":"할부","month":"10월","year":"2026년","rawNote":"1/12"},
        {"card":"하나카드","item":"제습기","amount":96424,"type":"공통","note":"일반","month":"8월","year":"2026년","rawNote":"1/5"},
        {"card":"하나카드","item":"제습기","amount":96422,"type":"공통","note":"할부","month":"9월","year":"2026년","rawNote":"2/5"},
        {"card":"하나카드","item":"제습기","amount":96422,"type":"공통","note":"할부","month":"10월","year":"2026년","rawNote":"3/5"},
        {"card":"롯데카드","item":"피부과","amount":110000,"type":"공통","note":"","month":"5월","year":"2026년","rawNote":"1/6"},
        {"card":"롯데카드","item":"피부과","amount":110000,"type":"공통","note":"","month":"6월","year":"2026년","rawNote":"2/6"},
        {"card":"롯데카드","item":"피부과","amount":110000,"type":"공통","note":"","month":"7월","year":"2026년","rawNote":"3/6"},
        {"card":"롯데카드","item":"피부과","amount":110000,"type":"공통","note":"","month":"8월","year":"2026년","rawNote":"4/6"},
        {"card":"롯데카드","item":"피부과","amount":110000,"type":"공통","note":"할부","month":"9월","year":"2026년","rawNote":"5/6"},
        {"card":"롯데카드","item":"피부과","amount":110000,"type":"공통","note":"할부","month":"10월","year":"2026년","rawNote":"6/6"}
    ];

    var installmentRows = seonjunData; // assume all are installments
    var activeMonth = '10월';
    var targetMonthNum = 10;
    
    var groups = {};
    for (var i = 0; i < installmentRows.length; i++) {
        var row = installmentRows[i];
        var c = row.card || '신용카드';
        var itemClean = row.item ? row.item.trim() : '할부항목';
        var baseAmt = Math.round((row.amount || 0) / 100) * 100;
        var key = c + '__' + itemClean + '__' + baseAmt;
        if (!groups[key]) {
            groups[key] = { card: c, item: itemClean, type: row.type || '공통', records: [] };
        }
        groups[key].records.push(row);
    }
    
    var monthOrder = { '1월': 1, '2월': 2, '3월': 3, '4월': 4, '5월': 5, '6월': 6, '7월': 7, '8월': 8, '9월': 9, '10월': 10, '11월': 11, '12월': 12 };
    var getMonthNum = function(m) { return monthOrder[m] || parseInt(m, 10) || 0; };
    
    var processedItems = [];
    var groupKeys = Object.keys(groups);
    for (var i = 0; i < groupKeys.length; i++) {
        var g = groups[groupKeys[i]];
        g.records.sort(function(a, b) { return getMonthNum(a.month) - getMonthNum(b.month); });
        
        var currentAmount = 0, futureAmount = 0, pastAmount = 0;
        var currentActiveRec = null;
        for (var j = 0; j < g.records.length; j++) {
            var r = g.records[j];
            var mNum = getMonthNum(r.month);
            if (r.month === activeMonth) {
                currentAmount += r.amount;
                currentActiveRec = r;
            }
            if (mNum > targetMonthNum) futureAmount += r.amount;
            if (mNum < targetMonthNum) pastAmount += r.amount;
        }
        if (!currentActiveRec) currentActiveRec = g.records[g.records.length - 1];
        
        var currentRound = 0;
        var totalRounds = g.records.length;
        
        var rawNoteMatch = null;
        if (currentActiveRec && currentActiveRec.rawNote) {
            rawNoteMatch = currentActiveRec.rawNote.match(/(\d+)\s*\/\s*(\d+)/);
        }
        
        if (rawNoteMatch) {
            currentRound = parseInt(rawNoteMatch[1], 10);
            totalRounds = parseInt(rawNoteMatch[2], 10);
        }
        
        if (totalRounds > g.records.length && currentAmount > 0) {
            var remainingRounds = Math.max(0, totalRounds - currentRound);
            if (futureAmount === 0 && remainingRounds > 0) {
                futureAmount = currentAmount * remainingRounds;
            }
        }
        
        var totalAmount = (pastAmount + currentAmount + futureAmount) || (currentAmount * totalRounds) || 0;
        
        processedItems.push({
            item: g.item,
            currentRound: currentRound,
            totalRounds: totalRounds,
            currentAmount: currentAmount,
            remainingAmount: futureAmount,
            totalAmount: totalAmount
        });
    }
    
    return JSON.stringify(processedItems, null, 2);
}
