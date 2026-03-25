$path = "d:\DHF\QLKV_WM\web_app\app.js"
$content = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8)

$old1 = @"
    const weeklySales = new Map();
    const storeWeeklyDays = new Map();
"@
$new1 = @"
    const weeklySales = new Map();
    const storeWeeklyDays = new Map();
    const storeWeeklyGroupDays = new Map();
"@
$content = $content.Replace($old1, $new1)

$old2 = '            let rawDate = String(row[''calendarday''] || row[''date''] || row[''ngay''] || '''').trim();
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
            }'
            
$new2 = '            let rawDate = String(row[''calendarday''] || row[''date''] || row[''ngay''] || '''').trim();
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
            }'
$content = $content.Replace($old2.Replace("`r`n", "`n"), $new2.Replace("`r`n", "`n"))
$content = $content.Replace($old2, $new2)

$old3 = '        let wTotal = weeklySales.has(key) ? weeklySales.get(key) : 0;'
$new3 = '        let wDataExt = weeklySales.get(key);
        let wTotal = wDataExt ? wDataExt.totalQty : 0;
        let wWeekdayQty = wDataExt ? wDataExt.weekdayQty : 0;
        let wWeekendQty = wDataExt ? wDataExt.weekendQty : 0;
        
        let wStoreGrps = storeWeeklyGroupDays.get(data.storeID);
        let wWeekdayDaysCount = wStoreGrps ? wStoreGrps.weekdays.size : 0;
        let wWeekendDaysCount = wStoreGrps ? wStoreGrps.weekends.size : 0;
        
        let wWeekdayAds = wWeekdayDaysCount > 0 ? wWeekdayQty / wWeekdayDaysCount : 0;
        let wWeekendAds = wWeekendDaysCount > 0 ? wWeekendQty / wWeekendDaysCount : 0;'
$content = $content.Replace($old3.Replace("`r`n", "`n"), $new3.Replace("`r`n", "`n"))
$content = $content.Replace($old3, $new3)

$old4 = '        // --- NEW: Tăng trưởng theo Leadtime ---
        let leadtimeGrowth = 0;
        let growthHtml = ''-'';
        if (mAds > 0 && totalLeadtime > 0) {
            let periodAds = basePeriodDemand / totalLeadtime;
            leadtimeGrowth = ((periodAds - mAds) / mAds) * 100;'

$new4 = '        // --- NEW: Tăng trưởng theo Leadtime (Đối chiếu Weekly vs Monthly trên từng Thứ) ---
        let leadtimeGrowth = 0;
        let growthHtml = ''-'';
        
        let periodAdsMonthly = basePeriodDemand / totalLeadtime; 
        let weeklyPeriodDemand = calculatePeriodDemand(invDate, totalLeadtime, wWeekdayAds, wWeekendAds);
        let periodAdsWeekly = weeklyPeriodDemand / totalLeadtime;

        if (periodAdsMonthly > 0 && totalLeadtime > 0 && periodAdsWeekly > 0) {
            leadtimeGrowth = ((periodAdsWeekly - periodAdsMonthly) / periodAdsMonthly) * 100;'

$content = $content.Replace($old4.Replace("`r`n", "`n"), $new4.Replace("`r`n", "`n"))
$content = $content.Replace($old4, $new4)

[IO.File]::WriteAllText($path, $content, [Text.Encoding]::UTF8)
