import os

path = r'd:\DHF\QLKV_WM\web_app\app.js'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

old1 = '''const weeklySales = new Map();
    const storeWeeklyDays = new Map();'''
new1 = '''const weeklySales = new Map();
    const storeWeeklyDays = new Map();
    const storeWeeklyGroupDays = new Map();'''
content = content.replace(old1, new1)

old2 = '''            let rawDate = String(row['calendarday'] || row['date'] || row['ngay'] || '').trim();
            if (rawDate && storeID) {
                if (!storeWeeklyDays.has(storeID)) storeWeeklyDays.set(storeID, new Set());
                storeWeeklyDays.get(storeID).add(rawDate);
            }

            if (isNaN(qty)) return;

            let prodStd = normalizeProductName(pr);
            if (!prodStd) {
                unmappedProducts.add(String(pr).trim());
                return;
            }
            let key = `${storeID}_${prodStd.toLowerCase()}`;
            
            if (!weeklySales.has(key)) {
                weeklySales.set(key, qty);
            } else {
                weeklySales.set(key, weeklySales.get(key) + qty);
            }'''

new2 = '''            let rawDate = String(row['calendarday'] || row['date'] || row['ngay'] || '').trim();
            let isWknd = false;
            if (rawDate && storeID) {
                if (!storeWeeklyDays.has(storeID)) storeWeeklyDays.set(storeID, new Set());
                storeWeeklyDays.get(storeID).add(rawDate);
                
                if (!storeWeeklyGroupDays.has(storeID)) {
                    storeWeeklyGroupDays.set(storeID, { weekdays: new Set(), weekends: new Set() });
                }
                let cDate = parseDateStrToTime(rawDate);
                let dayOfWeek = new Date(cDate).getDay();
                isWknd = (dayOfWeek === 5 || dayOfWeek === 6 || dayOfWeek === 0);
                if (isWknd) storeWeeklyGroupDays.get(storeID).weekends.add(rawDate);
                else storeWeeklyGroupDays.get(storeID).weekdays.add(rawDate);
            }

            if (isNaN(qty)) return;

            let prodStd = normalizeProductName(pr);
            if (!prodStd) {
                unmappedProducts.add(String(pr).trim());
                return;
            }
            let key = `${storeID}_${prodStd.toLowerCase()}`;
            
            if (!weeklySales.has(key)) {
                weeklySales.set(key, { totalQty: qty, weekdayQty: isWknd ? 0 : qty, weekendQty: isWknd ? qty : 0 });
            } else {
                let data = weeklySales.get(key);
                data.totalQty += qty;
                if (isWknd) data.weekendQty += qty;
                else data.weekdayQty += qty;
            }'''
content = content.replace(old2, new2)

old3 = '''        let wTotal = weeklySales.has(key) ? weeklySales.get(key) : 0;'''
new3 = '''        let wDataExt = weeklySales.get(key);
        let wTotal = wDataExt ? wDataExt.totalQty : 0;
        let wWeekdayQty = wDataExt ? wDataExt.weekdayQty : 0;
        let wWeekendQty = wDataExt ? wDataExt.weekendQty : 0;
        
        let wStoreGrps = storeWeeklyGroupDays.get(data.storeID);
        let wWeekdayDaysCount = wStoreGrps ? wStoreGrps.weekdays.size : 0;
        let wWeekendDaysCount = wStoreGrps ? wStoreGrps.weekends.size : 0;
        
        let wWeekdayAds = wWeekdayDaysCount > 0 ? wWeekdayQty / wWeekdayDaysCount : 0;
        let wWeekendAds = wWeekendDaysCount > 0 ? wWeekendQty / wWeekendDaysCount : 0;'''
content = content.replace(old3, new3)

old4 = '''        // THAY PHƯƠNG PHÁP TÍNH CỐ ĐỊNH BẰNG ĐỘNG THEO LOẠI NGÀY + NHÂN TREND FACTOR
        let basePeriodDemand = calculatePeriodDemand(invDate, totalLeadtime, weekdayAds, weekendAds);
        let totalDemand = basePeriodDemand * trendFactor;

        // --- NEW: Tăng trưởng theo Leadtime ---
        let leadtimeGrowth = 0;
        let growthHtml = '-';
        if (mAds > 0 && totalLeadtime > 0) {
            let periodAds = basePeriodDemand / totalLeadtime;
            leadtimeGrowth = ((periodAds - mAds) / mAds) * 100;
            if (leadtimeGrowth > 0) growthHtml = `<span style="color: var(--success)">+${leadtimeGrowth.toFixed(1)}%</span>`;
            else if (leadtimeGrowth < 0) growthHtml = `<span style="color: var(--danger)">${leadtimeGrowth.toFixed(1)}%</span>`;
            else growthHtml = `0%`;
        } else if (basePeriodDemand > 0) {
            growthHtml = `<span style="color: var(--success)">New</span>`;
        }'''

new4 = '''        // THAY PHƯƠNG PHÁP TÍNH CỐ ĐỊNH BẰNG ĐỘNG THEO LOẠI NGÀY + NHÂN TREND FACTOR
        let basePeriodDemand = calculatePeriodDemand(invDate, totalLeadtime, weekdayAds, weekendAds);
        let totalDemand = basePeriodDemand * trendFactor;

        // --- NEW: Tăng trưởng theo Leadtime (Weekly vs Monthly tương ứng) ---
        let leadtimeGrowth = 0;
        let growthHtml = '-';
        
        let periodAdsMonthly = basePeriodDemand / totalLeadtime; 
        let weeklyPeriodDemand = calculatePeriodDemand(invDate, totalLeadtime, wWeekdayAds, wWeekendAds);
        let periodAdsWeekly = weeklyPeriodDemand / totalLeadtime;

        if (periodAdsMonthly > 0 && totalLeadtime > 0 && periodAdsWeekly > 0) {
            leadtimeGrowth = ((periodAdsWeekly - periodAdsMonthly) / periodAdsMonthly) * 100;
            if (leadtimeGrowth > 0) growthHtml = `<span style="color: var(--success)">+${leadtimeGrowth.toFixed(1)}%</span>`;
            else if (leadtimeGrowth < 0) growthHtml = `<span style="color: var(--danger)">${Math.abs(leadtimeGrowth).toFixed(1)}%</span>`;
            else growthHtml = `0%`;
        } else if (periodAdsMonthly > 0) {
            growthHtml = `<span style="color: var(--success)">New</span>`;
        }'''
content = content.replace(old4, new4)

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)
