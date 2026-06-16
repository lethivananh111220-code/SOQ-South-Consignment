// --- Cáº¤U HÃŒNH FIREBASE ---
// Báº¡n cáº§n láº¥y thÃ´ng tin nÃ y tá»« Firebase Console (https://console.firebase.google.com/)
const firebaseConfig = {
    apiKey: "AIzaSyBHG5WoQVon5lgoyZNZ7agIVYJDjyZdRrY",
    authDomain: "soq-south-consignment.firebaseapp.com",
    databaseURL: "https://soq-south-consignment-default-rtdb.asia-southeast1.firebasedatabase.app",
    projectId: "soq-south-consignment",
    storageBucket: "soq-south-consignment.firebasestorage.app",
    messagingSenderId: "491007756368",
    appId: "1:491007756368:web:8ea77f51a2a0f3b151a955",
    measurementId: "G-MSG7VKL5QQ"
};

// Khá»Ÿi táº¡o Firebase náº¿u thÆ° viá»‡n Ä‘Ã£ táº£i thÃ nh cÃ´ng
if (typeof firebase !== 'undefined') {
    firebase.initializeApp(firebaseConfig);
}

// Danh sÃ¡ch rau Äƒn lÃ¡/RTE (Tá»· lá»‡ há»§y > 30%)
const RTE_PRODUCTS = [
    "Cáº£i hoa há»“ng baby", "Cáº£i Kale xoÄƒn", "Cáº£i Kale khá»§ng long", "BÃ´ng cáº£i xanh baby",
    "XÃ  lÃ¡ch frisÃ©e xanh ngá»t", "XÃ  lÃ¡ch romaine xanh thÆ°á»£ng háº¡ng", "XÃ  lÃ¡ch frisÃ©e tÃ­m ngá»t",
    "XÃ  lÃ¡ch romaine tÃ­m thÆ°á»£ng háº¡ng", "XÃ  lÃ¡ch baby lollo", "XÃ  lÃ¡ch baby thá»§y tinh",
    "Cáº£i ngá»t giá»‘ng nháº­t", "Cáº£i bÃ³ xÃ´i", "XÃ  lÃ¡ch há»—n há»£p", "Asian Mix",
    "Gourmet Italian Mix", "Sweet Baby Lettuces", "Baby Spring Mix", "Chopped Kale",
    "Pure Rocket", "Cáº£i bÃ³ xÃ´i baby Äƒn liá»n"
];

// Danh sÃ¡ch sáº¯p xáº¿p hiá»ƒn thá»‹ máº·c Ä‘á»‹nh (theo yÃªu cáº§u ngÆ°á»i dÃ¹ng)
const CUSTOM_PRODUCT_ORDER = [
    "XÃ  lÃ¡ch há»—n há»£p loáº¡i Baby Spring Mix 100g",
    "XÃ  lÃ¡ch há»—n há»£p loáº¡i Sweet Baby Lettuces 100g",
    "XÃ  lÃ¡ch há»—n há»£p loáº¡i Gourmet Italian Mix 100g",
    "XÃ  lÃ¡ch há»—n há»£p loáº¡i Pure Rocket 100g",
    "XÃ  lÃ¡ch há»—n há»£p loáº¡i Chopped Kale 100g",
    "XÃ  lÃ¡ch há»—n há»£p loáº¡i Asian Mix 120g",
    "Cáº£i bÃ³ xÃ´i baby Äƒn liá»n 100g",
    "DÆ°a leo giá»‘ng nháº­t 450g",
    "CÃ  rá»‘t 400g",
    "HÃ nh tÃ¢y mini 350g",
    "Khoai tÃ¢y mini 400g",
    "Khoai tÃ¢y mini 400g (Baby)",
    "CÃ  chua ngá»t chÃ¹m 250g",
    "Äáº­u cove giá»‘ng nháº­t 200g",
    "Cáº§n tÃ¢y 350g",
    "CÃ  rá»‘t baby 250g",
    "XÃ  lÃ¡ch frisÃ©e xanh ngá»t 190g",
    "XÃ  lÃ¡ch frisÃ©e tÃ­m ngá»t 150g",
    "XÃ  lÃ¡ch romaine xanh thÆ°á»£ng háº¡ng 170g",
    "XÃ  lÃ¡ch romaine tÃ­m thÆ°á»£ng háº¡ng 130g",
    "XÃ  lÃ¡ch há»—n há»£p 200g",
    "CÃ  chua Roma 400g",
    "CÃ  chua Cherry ngá»t 250g",
    "Cáº£i Kale xoÄƒn 250g",
    "Cáº£i Kale khá»§ng long 250g",
    "Äáº­u ngá»t 200g",
    "BÃ­ vua HÃ n Quá»‘c 300g up",
    "BÃ´ng cáº£i xanh baby 250g",
    "BÃ´ng cáº£i xanh 200g",
    "BÃ´ng cáº£i xanh (RETAIL KG)",
    "Cáº£i bÃ³ xÃ´i 300g",
    "Cáº£i hoa há»“ng baby 200g",
    "XÃ  lÃ¡ch baby thá»§y tinh 200g",
    "XÃ  lÃ¡ch baby lollo 200g",
    "Cáº£i ngá»t giá»‘ng nháº­t 300g",
    "Cáº£i Kale xoÄƒn 250g (khuyáº¿n mÃ£i)",
    "CÃ  rá»‘t 500g (NTX)",
    "DÆ°a leo giá»‘ng nháº­t 600g (NTX)",
    "Äáº­u cove giá»‘ng nháº­t 500g (NTX)",
    "Cáº§n tÃ¢y 600g (NTX)",
    "Khoai tÃ¢y há»“ng 500g (NTX)",
    "Khoai tÃ¢y vÃ ng 500g (NTX)",
    "HÃ nh tÃ¢y tÃ­m mini 350g",
    "HÃ nh tÃ¢y tÃ­m 500g (NTX)",
    "HÃ nh tÃ¢y vÃ ng 500g (NTX)",
    "BÃ­ háº¡t dáº» (RETAIL KG)"
];

const datasets = {
    schedule: null,
    inventory: null,
    input: null,
    monthly: null,
    weekly: null,
    mapping_raw: null,
    template_headers: null,
    trend_report: null
};

let scheduleFileName = "SOQ_Calculated_Order"; // TÃªn máº·c Ä‘á»‹nh

let productWeightMap = new Map();
let globalStoreRegionMap = new Map();
let globalStoreNamesMap = new Map();
let globalStoreAliasesMap = new Map();
let globalReverseStoreNamesMap = new Map();
let globalMappingMap = new Map();
let globalStandardNamesSet = new Set();
let currentWeeklyReviewList = [];
let currentWeeklyFilteredList = [];

function buildMetadataMaps() {
    globalStoreRegionMap.clear();
    globalStoreNamesMap.clear();
    globalStoreAliasesMap.clear();
    globalReverseStoreNamesMap.clear();
    globalMappingMap.clear();
    globalStandardNamesSet.clear();
    
    if (datasets.mapping_raw && datasets.mapping_raw.length > 0) {
        let headerRow = datasets.mapping_raw[0] || [];
        let iOda = 1, iWm = 2;
        for (let c = 0; c < headerRow.length; c++) {
            let h = String(headerRow[c]).toUpperCase();
            if (h.includes('ODA')) iOda = c;
            else if (h.includes('WM')) iWm = c;
        }
        for (let i = 1; i < datasets.mapping_raw.length; i++) {
            let r = datasets.mapping_raw[i];
            if (!r || !Array.isArray(r)) continue;
            let odaName = r[iOda] ? String(r[iOda]).trim() : '';
            let wmName = r[iWm] ? String(r[iWm]).trim().toLowerCase() : '';
            if (!odaName && !wmName && r.length >= 2) {
                wmName = r[0] ? String(r[0]).trim().toLowerCase() : '';
                odaName = r[1] ? String(r[1]).trim() : '';
            }
            if (wmName && odaName) {
                globalMappingMap.set(wmName, odaName);
                globalStandardNamesSet.add(odaName.trim().toLowerCase());
            }
        }
    }
    
    if (datasets.schedule && datasets.schedule.length > 0) {
        datasets.schedule.forEach(row => {
            let store = row['sap'] || row['storekey'] || row['storecode'] || row['makho'] || row['mach'] || row['mÃ£khÃ¡chhÃ ng'] || row['mÃ£cá»­ahÃ ng'] || row['nickname'] || row['storename'] || row['store'];
            if (!store) return;
            let storeID = extractSAP(store);
            let region = String(row['khuvuc'] || row['khuvá»±c'] || row['region'] || 'KhÃ¡c').trim();
            globalStoreRegionMap.set(storeID, region);
            let sName = row['tencuahang'] || row['tncahng'] || row['storename'] || row['store'] || row['nickname'] || '';
            let nickname = row['nickname'] || '';
            if (sName) globalStoreNamesMap.set(storeID, String(sName).trim());
            if (!globalStoreAliasesMap.has(storeID)) globalStoreAliasesMap.set(storeID, new Set());
            if (sName) globalStoreAliasesMap.get(storeID).add(normalizeKey(sName));
            if (nickname) globalStoreAliasesMap.get(storeID).add(normalizeKey(nickname));
            globalStoreAliasesMap.get(storeID).add(normalizeKey(storeID));
        });
        globalStoreAliasesMap.forEach((aliases, id) => {
            aliases.forEach(alias => {
                globalReverseStoreNamesMap.set(alias, id);
            });
        });
        globalStoreNamesMap.forEach((name, id) => {
            globalReverseStoreNamesMap.set(normalizeKey(name), id);
            globalReverseStoreNamesMap.set(id, id);
        });
    }
}

function buildProductWeightMap() {
    productWeightMap.clear();
    if (!datasets.mapping_raw || datasets.mapping_raw.length === 0) return;
    let headerRow = datasets.mapping_raw[0] || [];
    let iOda = 1;
    let iWeight = -1;
    for (let c = 0; c < headerRow.length; c++) {
        let h = String(headerRow[c]).toUpperCase();
        if (h.includes('ODA')) iOda = c;
        if (h.includes('KHá»I LÆ¯á»¢NG') || h.includes('KHOI LUONG') || h.includes('WEIGHT') || h.includes('Tá»ŠNH') || h.includes('GRAM')) {
            iWeight = c;
        }
    }
    if (iWeight === -1 && headerRow.length > 4) {
        iWeight = 4;
    }
    for (let i = 1; i < datasets.mapping_raw.length; i++) {
        let r = datasets.mapping_raw[i];
        if (!r || !Array.isArray(r)) continue;
        let odaName = r[iOda] ? String(r[iOda]).trim() : '';
        if (odaName) {
            let weightVal = 1000;
            if (iWeight !== -1 && r[iWeight] !== undefined && r[iWeight] !== null) {
                let parsedWeight = parseFloat(String(r[iWeight]).replace(/,/g, ''));
                if (!isNaN(parsedWeight) && parsedWeight > 0) {
                    weightVal = parsedWeight;
                }
            }
            productWeightMap.set(odaName.toLowerCase(), weightVal);
        }
    }
}

function getProductWeightKG(productName) {
    if (!productName) return 1;
    let nameLower = productName.toLowerCase();
    if (productWeightMap.has(nameLower)) {
        return productWeightMap.get(nameLower) / 1000;
    }
    for (let [key, val] of productWeightMap.entries()) {
        if (nameLower.includes(key) || key.includes(nameLower)) {
            return val / 1000;
        }
    }
    let gMatch = productName.match(/(\d+(?:\.\d+)?)\s*g/i);
    if (gMatch) {
        return parseFloat(gMatch[1]) / 1000;
    }
    let kgMatch = productName.match(/(\d+(?:\.\d+)?)\s*kg/i);
    if (kgMatch) {
        return parseFloat(kgMatch[1]);
    }
    return 1;
}

function getGlobalNormalizedProduct(name) {
    if (!name) return '';
    let n = String(name).trim().toLowerCase();
    if (globalMappingMap.size === 0 && datasets.mapping_raw) {
        buildMetadataMaps();
    }
    if (globalMappingMap.has(n)) return String(globalMappingMap.get(n)).trim();
    if (globalStandardNamesSet.has(n)) return String(name).trim();
    return String(name).trim();
}

function formatDateStr(d) {
    let year = d.getFullYear();
    let month = String(d.getMonth() + 1).padStart(2, '0');
    let day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function formatDateDMY(d) {
    let day = String(d.getDate()).padStart(2, '0');
    let month = String(d.getMonth() + 1).padStart(2, '0');
    let year = d.getFullYear();
    return `${day}/${month}/${year}`;
}

function getMondayOfWeek(year, weekNum) {
    let simple = new Date(year, 0, 4);
    let dayOfWeek = simple.getDay();
    let dayDiff = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    let simpleMonday = new Date(year, 0, 4 + dayDiff);
    let targetMonday = new Date(simpleMonday.getTime() + (weekNum - 1) * 7 * 86400000);
    return targetMonday;
}

function getWeekNumber(d) {
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
    let yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
    let weekNo = Math.ceil(( ( (d - yearStart) / 86400000) + 1)/7);
    return { year: d.getUTCFullYear(), week: weekNo };
}

function getWeeksOfYear() {
    let now = new Date();
    let currentWeekInfo = getWeekNumber(now);
    let currentWeek = currentWeekInfo.week;
    let year = currentWeekInfo.year;
    let options = [];
    for (let w = currentWeek; w >= 1; w--) {
        let monday = getMondayOfWeek(year, w);
        let sunday = new Date(monday.getTime() + 6 * 86400000);
        let label = `Tuáº§n ${w} (${formatDateDMY(monday)} - ${formatDateDMY(sunday)})`;
        options.push({
            year: year,
            week: w,
            mondayStr: formatDateStr(monday),
            sundayStr: formatDateStr(sunday),
            label: label
        });
    }
    return options;
}

function getWeeklySalesMonday(weeklyData, fileName) {
    if (Array.isArray(weeklyData)) {
        for (let i = 0; i < Math.min(200, weeklyData.length); i++) {
            let row = weeklyData[i];
            if (!row) continue;
            let rawDate = String(row['calendarday'] || row['date'] || row['ngay'] || '').trim();
            if (rawDate) {
                let ts = parseDateStrToTime(rawDate);
                if (ts > 0) {
                    let d = new Date(ts);
                    let day = d.getDay();
                    let diff = d.getDate() - day + (day === 0 ? -6 : 1);
                    let mondayDate = new Date(d.getFullYear(), d.getMonth(), diff);
                    return formatDateStr(mondayDate);
                }
            }
        }
    }
    if (fileName) {
        let m = fileName.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
        if (m) {
            let ts = parseDateStrToTime(`${m[1]}-${m[2]}-${m[3]}`);
            if (ts > 0) {
                let d = new Date(ts);
                let day = d.getDay();
                let diff = d.getDate() - day + (day === 0 ? -6 : 1);
                let mondayDate = new Date(d.getFullYear(), d.getMonth(), diff);
                return formatDateStr(mondayDate);
            }
        }
        let m2 = fileName.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
        if (m2) {
            let ts = parseDateStrToTime(`${m2[3]}-${m2[2]}-${m2[1]}`);
            if (ts > 0) {
                let d = new Date(ts);
                let day = d.getDay();
                let diff = d.getDate() - day + (day === 0 ? -6 : 1);
                let mondayDate = new Date(d.getFullYear(), d.getMonth(), diff);
                return formatDateStr(mondayDate);
            }
        }
        let wMatch = fileName.match(/W(\d{1,2})/i);
        if (wMatch) {
            let weekNum = parseInt(wMatch[1], 10);
            let year = new Date().getFullYear();
            let mondayDate = getMondayOfWeek(year, weekNum);
            return formatDateStr(mondayDate);
        }
    }
    let d = new Date();
    let day = d.getDay();
    let diff = d.getDate() - day + (day === 0 ? -6 : 1);
    let mondayDate = new Date(d.getFullYear(), d.getMonth(), diff);
    return formatDateStr(mondayDate);
}

function getOverlapDays(mondayStr, startRangeStr, endRangeStr) {
    let mon = new Date(mondayStr);
    let sun = new Date(mon.getTime() + 6 * 86400000);
    let rangeStart = new Date(startRangeStr);
    let rangeEnd = new Date(endRangeStr);
    let overlapStart = new Date(Math.max(mon, rangeStart));
    let overlapEnd = new Date(Math.min(sun, rangeEnd));
    if (overlapStart <= overlapEnd) {
        return Math.round((overlapEnd - overlapStart) / 86400000) + 1;
    }
    return 0;
}

// Elements
const btnCalculate = document.getElementById('btn-calculate');
const btnExport = document.getElementById('btn-export');
const resultsSection = document.getElementById('results-section');
const tbody = document.getElementById('soq-tbody');
const inputUserName = document.getElementById('user-name');

function removeAccents(str) {
    if (!str) return '';
    return str.normalize('NFD')
              .replace(/[\u0300-\u036f]/g, '')
              .replace(/Ä‘/g, 'd').replace(/Ä/g, 'D');
}

// Helper to normalize column names
function normalizeKey(key) {
    if (!key) return '';
    let s = removeAccents(key.toString().toLowerCase());
    return s.replace(/[^a-z0-9]/g, '');
}

// HÃ m trÃ­ch xuáº¥t tá»± Ä‘á»™ng bá» qua cÃ¡c tiÃªu Ä‘á» bÃ¡o cÃ¡o rÃ¡c á»Ÿ file há»‡ thá»‘ng (Excel report info)
function extractJsonDataCleanly(worksheet) {
    let rawArr = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, dateNF: 'yyyy-mm-dd hh:mm:ss' });
    if (!rawArr || rawArr.length === 0) return [];

    let headerIdx = 0;
    // TÃ¬m dÃ²ng header thá»±c sá»± (ThÆ°á»ng cÃ³ chá»©a cÃ¡c chá»¯ khÃ³a nháº­n diá»‡n vÃ  > 3 cá»™t dá»¯ liá»‡u)
    for (let i = 0; i < Math.min(20, rawArr.length); i++) {
        let r = rawArr[i];
        if (!r) continue;
        let validCols = r.filter(c => typeof c === 'string' && c.trim() !== '');
        if (validCols.length >= 2 && r.some(c => typeof c === 'string' && (c.toUpperCase().includes('SAP') || c.toUpperCase().includes('STORE') || c.toUpperCase().includes('NICKNAME') || c.toUpperCase().includes('TÃŠN') || c.toUpperCase().includes('ARTICLE') || c.toUpperCase().includes('PRODUCT')))) {
            headerIdx = i;
            break;
        }
    }

    const headersRaw = rawArr[headerIdx] || [];
    const headersPrefix = headerIdx > 0 ? (rawArr[headerIdx - 1] || []) : [];
    
    let headers = headersRaw.map((h, j) => {
        let prefix = headersPrefix[j] ? String(headersPrefix[j]).trim() + '_' : '';
        return normalizeKey(prefix + h);
    });
    
    // FALLBACK: Náº¿u headers toÃ n lÃ  sá»‘ (Excel Serial Dates) -> CÃ³ thá»ƒ Ä‘Ã¢y lÃ  file matrix khÃ´ng cÃ³ label.
    let numericHeadersCount = headersRaw.filter(h => typeof h === 'number' && h > 40000).length;
    // TÄƒng cÆ°á»ng kiá»ƒm tra cáº£ headersPrefix náº¿u cÃ³
    if (headersPrefix.length > 0) numericHeadersCount += headersPrefix.filter(h => typeof h === 'number' && h > 40000).length;

    if (numericHeadersCount > 5) {
        // ÄÃ¢y lÃ  dáº¡ng file Lá»‹ch Matrix. Ã‰p cÃ¡c cá»™t cá»‘ Ä‘á»‹nh (0: Type, 1: SAP, 4: Name)
        // LÆ°u Ã½: Náº¿u cÃ³ prefix, headers[1] cÃ³ thá»ƒ lÃ  "v_sap". Ta rÃ  soÃ¡t index.
        headers[0] = 'type';
        headers[1] = 'sap';
        headers[2] = 'tier';
        headers[3] = 'function';
        headers[4] = 'storename';
    }

    const json = [];

    for (let i = headerIdx + 1; i < rawArr.length; i++) {
        let row = rawArr[i];
        if (!row || row.length === 0) continue;

        // Bá» qua dÃ²ng Total (DÃ²ng tá»•ng cá»™ng cá»§a SAP)
        if (row.some(cell => String(cell).toUpperCase().includes('RESULT') || String(cell).toUpperCase() === 'TOTAL')) continue;

        let obj = {};
        let hasData = false;
        for (let j = 0; j < headers.length; j++) {
            if (row[j] !== undefined && row[j] !== null && String(row[j]).trim() !== '') {
                obj[headers[j]] = row[j]; // Composite

                // KhÃ´i phá»¥c viá»‡c Ä‘á»c cÃ¡c cá»™t Ä‘Æ¡n giáº£n (sap, date...) Ä‘á»ƒ khÃ´ng bá»‹ hÆ° tÃªn do prefix cháº·n.
                // NgÄƒn cháº·n riÃªng biá»‡t lá»—i trÆ°á»£t/chá»“ng láº¯p lá»‹ch cÃ¡c ngÃ y trong tuáº§n (Ä‘Ã£ xá»­ lÃ½ á»Ÿ bÆ°á»›c trÆ°á»›c).
                let rawClean = normalizeKey(headersRaw[j]);
                const wDays = ['monday','tuesday','wednesday','thursday','friday','saturday','sunday'];
                if (!wDays.some(d => rawClean.includes(d))) {
                    obj[rawClean] = row[j];
                    obj[headersRaw[j]] = row[j];
                }
                
                hasData = true;
            }
        }
        if (hasData) json.push(obj);
    }
    return json;
}

// Convert Excel Serial Date or String to Timestamp
function parseDateStrToTime(val) {
    if (!val && val !== 0) return 0;
    if (typeof val === 'number') {
        // Fix for pure Excel serial numbers (brings it closer to local midnight)
        let utcDate = new Date(Math.round((val - 25569) * 86400 * 1000));
        return new Date(utcDate.getUTCFullYear(), utcDate.getUTCMonth(), utcDate.getUTCDate()).getTime();
    }
    let s = String(val).trim().split(' ')[0]; // Bá» time náº¿u cÃ³

    // Support YYYY-MM-DD formats natively returning local midnight
    let m2 = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (m2) {
        return new Date(parseInt(m2[1], 10), parseInt(m2[2], 10) - 1, parseInt(m2[3], 10)).getTime();
    }

    // Support DD/MM/YYYY or MM/DD/YYYY formats
    let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (m) {
        let p1 = parseInt(m[1], 10);
        let p2 = parseInt(m[2], 10);
        let yr = parseInt(m[3], 10);
        if (yr < 100) yr += 2000;

        if (typeof window.dateFmtDetected === 'undefined') window.dateFmtDetected = 0;

        let d, mth;
        if (p1 > 12) {
            d = p1; mth = p2; // Definitively DD/MM
            window.dateFmtDetected = 1;
        } else if (p2 > 12) {
            mth = p1; d = p2; // Definitively MM/DD
            window.dateFmtDetected = -1;
        } else {
            // Ambiguous date (e.g., 04/05/2026). Check majority format detected
            if (window.dateFmtDetected === -1) {
                mth = p1; d = p2;
            } else {
                d = p1; mth = p2; // Default to VN DD/MM
            }
        }
        return new Date(yr, mth - 1, d).getTime();
    }

    const parsed = new Date(s).getTime();
    if (!isNaN(parsed)) {
        // Ensure returning local midnight instead of UTC (fallback if parsed successfully)
        let d = new Date(parsed);
        return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
    }
    return 0; // fallback numerical value
}

// TÃ­nh tá»•ng nhu cáº§u Ä‘á»™ng dá»±a trÃªn loáº¡i ngÃ y (thá»© 2-5 hoáº·c thá»© 6-CN)
function calculatePeriodDemand(startTs, totalDays, adsWeekday, adsWeekend) {
    let total = 0;
    let roundedDays = Math.round(totalDays);
    if (roundedDays <= 0) return 0;

    for (let i = 0; i < roundedDays; i++) {
        let currentTs = startTs + (i * 86400 * 1000);
        let d = new Date(currentTs);
        let dw = d.getDay(); // 0:CN, 1:T2, ... 6:T7
        if (dw >= 1 && dw <= 5) {
            total += adsWeekday;
        } else {
            total += adsWeekend;
        }
    }
    return total;
}

// Universal File Reader 
function handleFileUpload(event, type) {
    const file = event.target.files[0];
    if (!file) return;

    const statusEl = document.getElementById(`status-${type}`);
    statusEl.textContent = `Äang Ä‘á»c ${file.name}...`;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        let firstSheetName = workbook.SheetNames[0];
        
        // Theo yÃªu cáº§u: Láº¥y dá»¯ liá»‡u tá»« sheet "Summary by Products" cho file Input
        if (type === 'input') {
            const desiredSheet = workbook.SheetNames.find(sheet => sheet.trim().toLowerCase() === 'summary by products');
            if (desiredSheet) {
                firstSheetName = desiredSheet;
            }
        }
        
        const worksheet = workbook.Sheets[firstSheetName];

        if (type === 'mapping') {
            const arr = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            datasets['mapping_raw'] = arr;
            saveToDB('mapping_raw', arr);
            buildMetadataMaps();
            buildProductWeightMap();
            statusEl.textContent = `ÄÃ£ táº£i & lÆ°u trá»¯: ${file.name} (${arr.length} dÃ²ng)`;
            statusEl.classList.add('success');
            checkReady();
            return;
        }

        try {
            if (type === 'template') {
                const arr = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                if (arr && arr.length > 0) {
                    let headerRow = arr.find(row => row && row.length > 0);
                    if (headerRow) {
                        datasets.template_headers = headerRow.map(h => String(h).trim());
                        saveToDB('template_headers', datasets.template_headers);
                        statusEl.textContent = `ÄÃ£ náº¡p Form Máº«u (${datasets.template_headers.length} cá»™t)`;
                        statusEl.classList.add('success');

                        if (typeof firebase !== 'undefined') {
                            firebase.database().ref('global_template').set({
                                headers: datasets.template_headers,
                                timestamp: Date.now()
                            }).then(() => console.log("ÄÃ£ cáº­p nháº­t Form Máº«u lÃªn Cloud."))
                              .catch(err => console.error("Lá»—i lÆ°u Form Máº«u lÃªn Cloud:", err));
                        }
                    } else {
                        statusEl.textContent = `Form Máº«u trá»‘ng!`;
                        statusEl.style.color = "var(--danger)";
                    }
                }
                return;
            }

            if (type === 'monthly' || type === 'trend_report') {
                let allJson = [];
                workbook.SheetNames.forEach(name => {
                    const ws = workbook.Sheets[name];
                    const json = extractJsonDataCleanly(ws);
                    if (json && json.length > 0) allJson = allJson.concat(json);
                });
                datasets[type] = allJson;
            } else {
                const json = extractJsonDataCleanly(worksheet);
                datasets[type] = json;
            }

            if (type === 'monthly' || type === 'weekly' || type === 'trend_report' || type === 'schedule' || type === 'inventory' || type === 'input') {
                saveToDB(type, datasets[type]);
                if (type === 'schedule') {
                    scheduleFileName = file.name.replace(/\.[^/.]+$/, "");
                    saveToDB('soq_latest_filename', scheduleFileName);
                    buildMetadataMaps();
                }
                if (type === 'weekly') {
                    let mondayStr = getWeeklySalesMonday(datasets['weekly'], file.name);
                    saveToDB('weekly_sales_archive_' + mondayStr, datasets['weekly']);
                    if (typeof firebase !== 'undefined') {
                        firebase.database().ref('archive_weekly_sales/' + mondayStr).set({
                            data: datasets['weekly'],
                            timestamp: Date.now(),
                            fileName: file.name
                        }).then(() => console.log(`ÄÃ£ lÆ°u trá»¯ bÃ¡o cÃ¡o doanh sá»‘ tuáº§n ngÃ y ${mondayStr} lÃªn Cloud.`))
                          .catch(err => console.error("Lá»—i lÆ°u doanh sá»‘ tuáº§n Cloud:", err));
                    }
                }
                statusEl.textContent = `ÄÃ£ táº£i & lÆ°u trá»¯: ${file.name} (${datasets[type].length} dÃ²ng)`;
            } else {
                statusEl.textContent = `ÄÃ£ táº£i: ${file.name} (${datasets[type].length} dÃ²ng)`;
            }
            statusEl.classList.add('success');
            checkReady();
        } catch (err) {
            console.error(err);
            statusEl.textContent = "Lá»—i xá»­ lÃ½ file: " + err.message;
            statusEl.style.color = "var(--danger)";
        }
    };
    reader.onerror = () => {
        statusEl.textContent = "Lá»—i Ä‘á»c file tá»« mÃ¡y tÃ­nh!";
        statusEl.style.color = "var(--danger)";
    };
    reader.readAsArrayBuffer(file);
}

// Bind events
document.getElementById('file-schedule').addEventListener('change', e => handleFileUpload(e, 'schedule'));
document.getElementById('file-inventory').addEventListener('change', e => handleFileUpload(e, 'inventory'));
document.getElementById('file-input').addEventListener('change', e => handleFileUpload(e, 'input'));
document.getElementById('file-monthly').addEventListener('change', e => handleFileUpload(e, 'monthly'));
document.getElementById('file-weekly').addEventListener('change', e => handleFileUpload(e, 'weekly'));
document.getElementById('file-trend_report').addEventListener('change', e => handleFileUpload(e, 'trend_report'));
document.getElementById('file-mapping').addEventListener('change', e => handleFileUpload(e, 'mapping'));
document.getElementById('file-template').addEventListener('change', e => handleFileUpload(e, 'template'));

// --- IndexedDB Caching cho cÃ¡c file cá»‘ Ä‘á»‹nh (Monthly, Weekly, Mapping) ---
const DB_NAME = "SOQ_V1";
function initDB() {
    return new Promise((resolve, reject) => {
        let req = indexedDB.open(DB_NAME, 1);
        req.onupgradeneeded = e => {
            let db = e.target.result;
            if (!db.objectStoreNames.contains('files')) db.createObjectStore('files');
        };
        req.onsuccess = e => resolve(e.target.result);
        req.onerror = e => reject(e.target.error);
    });
}
async function deleteFromDB(key) {
    try {
        let db = await initDB();
        let tx = db.transaction('files', 'readwrite');
        tx.objectStore('files').delete(key);
    } catch (e) { }
}

async function saveToDB(key, data) {
    try {
        let payload = { data: data, timestamp: Date.now() };
        let db = await initDB();
        let tx = db.transaction('files', 'readwrite');
        tx.objectStore('files').put(payload, key);
    } catch (e) { console.error('Lá»—i lÆ°u cache', e); }
}

function getWeekStart(date) {
    let d = new Date(date);
    let day = d.getDay();
    let diff = d.getDate() - day + (day === 0 ? -6 : 1);
    d.setDate(diff);
    d.setHours(0, 0, 0, 0);
    return d.getTime();
}

async function loadFromDB(key) {
    try {
        let db = await initDB();
        let raw = await new Promise((resolve) => {
            let tx = db.transaction('files', 'readonly');
            let req = tx.objectStore('files').get(key);
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => resolve(null);
        });

        if (!raw) return null;
        if (Array.isArray(raw)) return raw; // Cache cá»• Ä‘iá»ƒn 

        if (raw.timestamp && raw.data) {
            let dDate = new Date(raw.timestamp);
            let nDate = new Date();

            if (key === 'monthly' || key === 'trend_report') {
                if (dDate.getMonth() !== nDate.getMonth() || dDate.getFullYear() !== nDate.getFullYear()) {
                    // Náº¿u lÃ  file thÃ¡ng, chá»‰ xÃ³a náº¿u quÃ¡ 2 thÃ¡ng (Ä‘á»ƒ user dÃ¹ng Ä‘Æ°á»£c thÃ¡ng trÆ°á»›c + thÃ¡ng nÃ y)
                    let monthDiff = (nDate.getFullYear() - dDate.getFullYear()) * 12 + (nDate.getMonth() - dDate.getMonth());
                    if (monthDiff > 2) {
                        await deleteFromDB(key);
                        return { invalidated: true, reason: "dá»¯ liá»‡u quÃ¡ cÅ© (>2 thÃ¡ng)" };
                    }
                }
            } else if (key === 'weekly') {
                // Háº¡n sá»­ dá»¥ng cá»§a Doanh sá»‘ tuáº§n: Tá»« ngÃ y táº£i lÃªn (dDate) kÃ©o dÃ i Ä‘áº¿n Thá»© 3 cá»§a tuáº§n káº¿ tiáº¿p
                let expirationTime = getWeekStart(dDate) + 8 * 86400000; 
                if (nDate.getTime() >= expirationTime) {
                    await deleteFromDB(key);
                    return { invalidated: true, reason: "sang thá»© 3 tuáº§n má»›i" };
                }
            } else if (key === 'soq_latest_array') {
                // Háº¿t háº¡n bá»™ nhá»› táº¡m khi sang ngÃ y má»›i
                if (dDate.getDate() !== nDate.getDate() || dDate.getMonth() !== nDate.getMonth() || dDate.getFullYear() !== nDate.getFullYear()) {
                    await deleteFromDB(key);
                    return { invalidated: true, reason: "Ä‘Ã£ sang ngÃ y má»›i" };
                }
            }
            return raw.data;
        }
        return raw;
    } catch (e) { return null; }
}

window.addEventListener('DOMContentLoaded', async () => {
    // Tá»± sinh Dropdown Chá»n NgÃ y 1-31, máº·c Ä‘á»‹nh nháº£y vÃ o sá»‘ trÃ¹ng vá»›i HÃ´m Nay (Today)
    let dateSelect = document.getElementById('targetDeliveryDate');
    if (dateSelect) {
        let tomorrow = new Date();
        tomorrow.setDate(tomorrow.getDate() + 1);
        let targetDay = tomorrow.getDate();

        for (let i = 1; i <= 31; i++) {
            let opt = document.createElement('option');
            opt.value = i;
            opt.text = "NgÃ y " + i;
            if (i === targetDay) opt.selected = true;
            dateSelect.appendChild(opt);
        }
    }

    // Tá»± Ä‘á»™ng load láº¡i Cache cá»§a Monthly, Weekly, Mapping File náº¿u cÃ³
    let [cMonthly, cWeekly, cMapping, cTemplate, cTrendReport, cSchedule, cInventory, cInput, cDeliveryDate] = await Promise.all([
        loadFromDB('monthly'), loadFromDB('weekly'), loadFromDB('mapping_raw'), loadFromDB('template_headers'), loadFromDB('trend_report'),
        loadFromDB('schedule'), loadFromDB('inventory'), loadFromDB('input'), loadFromDB('soq_latest_delivery_date')
    ]);

    if (typeof firebase !== 'undefined') {
        try {
            let snapshot = await firebase.database().ref('global_template').once('value');
            let data = snapshot.val();
            if (data && data.headers && data.headers.length > 0) {
                cTemplate = data.headers;
                saveToDB('template_headers', cTemplate);
            }
        } catch (err) {
            console.error("Lá»—i láº¥y Form Máº«u tá»« Cloud, dÃ¹ng Local:", err);
        }
    }
    
    if (cDeliveryDate) {
        currentDeliveryDateStr = cDeliveryDate;
    }

    if (cTemplate && cTemplate.length > 0) {
        datasets.template_headers = cTemplate;
        let el = document.getElementById('status-template');
        if (el) { el.textContent = `ÄÃ£ náº¡p Form Máº«u (${cTemplate.length} cá»™t)`; el.classList.add('success'); }
    }

    if (cMonthly) {
        if (cMonthly.invalidated) {
            let el = document.getElementById('status-monthly');
            if (el) { el.innerHTML = `<span style="color: #ff9800; font-weight: bold;">LÆ°u Ã½: ÄÃ£ sang thÃ¡ng má»›i. Vui lÃ²ng Táº£i LÃªn file cáº­p nháº­t!</span>`; el.classList.remove('success'); }
        } else if (cMonthly.length > 0) {
            datasets.monthly = cMonthly;
            let el = document.getElementById('status-monthly');
            if (el) { el.textContent = `ÄÃ£ dÃ¹ng báº£n lÆ°u trÆ°á»›c (${cMonthly.length} dÃ²ng)`; el.classList.add('success'); }
        }
    }

    if(cWeekly) {
        if (cWeekly.invalidated) {
            let el = document.getElementById('status-weekly');
            if (el) { el.innerHTML = `<span style="color: #ff9800; font-weight: bold;">LÆ°u Ã½: Sang Thá»© 3 tuáº§n má»›i. Vui lÃ²ng táº£i sá»‘ bÃ¡o cÃ¡o tuáº§n má»›i!</span>`; el.classList.remove('success'); }
        } else if (cWeekly.length > 0) {
            datasets.weekly = cWeekly;
            let el = document.getElementById('status-weekly');
            if (el) { el.textContent = `ÄÃ£ dÃ¹ng báº£n lÆ°u trÆ°á»›c (${cWeekly.length} dÃ²ng)`; el.classList.add('success'); }
        }
    }
    
    if (cTrendReport) {
        if (cTrendReport.invalidated) {
            let el = document.getElementById('status-trend_report');
            if (el) { el.innerHTML = `<span style="color: #ff9800; font-weight: bold;">LÆ°u Ã½: ÄÃ£ sang thÃ¡ng má»›i. Vui lÃ²ng Táº£i LÃªn file xu hÆ°á»›ng bÃ¡n má»›i!</span>`; el.classList.remove('success'); }
        } else if (cTrendReport.length > 0) {
            datasets.trend_report = cTrendReport;
            let el = document.getElementById('status-trend_report');
            if (el) { el.textContent = `ÄÃ£ dÃ¹ng báº£n lÆ°u trÆ°á»›c (${cTrendReport.length} dÃ²ng)`; el.classList.add('success'); }
        }
    }

    if (cMapping && cMapping.length > 0) {
        datasets.mapping_raw = cMapping;
        let el = document.getElementById('status-mapping');
        if (el) { el.textContent = `ÄÃ£ dÃ¹ng báº£n lÆ°u trÆ°á»›c (${cMapping.length} dÃ²ng)`; el.classList.add('success'); }
    }

    if (cSchedule && cSchedule.length > 0) {
        datasets.schedule = cSchedule;
        let el = document.getElementById('status-schedule');
        if (el) {
            let savedName = await loadFromDB('soq_latest_filename');
            if (savedName) scheduleFileName = savedName;
            el.textContent = `ÄÃ£ dÃ¹ng báº£n lÆ°u trÆ°á»›c (${cSchedule.length} dÃ²ng)`;
            el.classList.add('success');
        }
    }

    if (cInventory && cInventory.length > 0) {
        datasets.inventory = cInventory;
        let el = document.getElementById('status-inventory');
        if (el) { el.textContent = `ÄÃ£ dÃ¹ng báº£n lÆ°u trÆ°á»›c (${cInventory.length} dÃ²ng)`; el.classList.add('success'); }
    }

    if (cInput && cInput.length > 0) {
        datasets.input = cInput;
        let el = document.getElementById('status-input');
        if (el) { el.textContent = `ÄÃ£ dÃ¹ng báº£n lÆ°u trÆ°á»›c (${cInput.length} dÃ²ng)`; el.classList.add('success'); }
    }

    buildMetadataMaps();
    buildProductWeightMap();
    checkReady();
});

function checkReady() {
    if (datasets.schedule && datasets.inventory && datasets.input && datasets.monthly && datasets.weekly) {
        btnCalculate.disabled = false;
        btnCalculate.textContent = "Tiáº¿n hÃ nh tÃ­nh SOQ";
    }
}

let finalResults = [];
let isHistoryView = false;
let isArchiveView = false;
let currentDeliveryDateStr = "";

function archiveTodayData() {
    if (!finalResults || finalResults.length === 0) return;
    
    let dateStr = currentDeliveryDateStr;
    if (!dateStr) {
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');
        dateStr = `${year}-${month}-${day}`;
    }

    const archivePayload = {
        filename: scheduleFileName,
        results: finalResults,
        timestamp: Date.now(),
        dateStr: dateStr
    };
    saveToDB('soq_archive_' + dateStr, archivePayload);

    if (typeof firebase !== 'undefined') {
        let userName = inputUserName ? inputUserName.value.trim() : "Há»‡ thá»‘ng";
        if (!userName) userName = "áº¨n danh";
        
        const cloudPayload = {
            filename: scheduleFileName,
            results: finalResults,
            timestamp: Date.now(),
            dateStr: dateStr,
                      deliveryDateStr: currentDeliveryDateStr,
            userName: userName
        };
        
        firebase.database().ref('archive_soq/' + dateStr).set(cloudPayload)
            .then(() => console.log(`ÄÃ£ lÆ°u trá»¯ dá»¯ liá»‡u ngÃ y ${dateStr} lÃªn Cloud.`))
            .catch(err => console.error("Lá»—i lÆ°u trá»¯ Cloud:", err));
    }
}

function extractSAP(str) {
    if (!str) return "";
    let s = String(str).trim();
    // Æ¯u tiÃªn: Náº¿u lÃ  chuá»—i sá»‘ Ä‘á»©ng Ä‘á»™c láº­p (cÃ³ thá»ƒ cÃ³ chá»¯ bao quanh bá»Ÿi dáº¥u cÃ¡ch) -> Láº¥y sá»‘
    let m = s.match(/\b\d+\b/);
    if (m) return Number(m[0]).toString();
    
    return s.toLowerCase();
}

btnCalculate.addEventListener('click', () => {
    try {
        tbody.innerHTML = "";
        finalResults = [];
        resultsSection.style.display = 'none';
        // --- KIá»‚M TRA Dá»® LIá»†U Äáº¦U VÃ€O ---
        if (!datasets.schedule || datasets.schedule.length === 0) {
            alert("Vui lÃ²ng táº£i file Lá»‹ch giao hÃ ng (Schedule)!");
            return;
        }
        if (!datasets.inventory || datasets.inventory.length === 0) {
            alert("Vui lÃ²ng táº£i file Tá»“n kho (Merchandiser)!");
            return;
        }
        if (!datasets.monthly || datasets.monthly.length === 0) {
            alert("Vui lÃ²ng táº£i file Doanh sá»‘ thÃ¡ng (Monthly Sales)!");
            return;
        }
        // Tip: Mapping lÃ  báº¯t buá»™c náº¿u muá»‘n dÃ¹ng tÃ­nh nÄƒng lá»c máº«u (strict mapping)
        if (!datasets.mapping_raw || datasets.mapping_raw.length === 0) {
            alert("LÆ°u Ã½: Báº¡n chÆ°a táº£i file Mapping. Há»‡ thá»‘ng sáº½ láº¥y tÃªn gá»‘c tá»« file doanh sá»‘.");
        }

        // --- TÃNH TOÃN NGÃ€Y GIAO HÃ€NG (WEEKEND HAY WEEKDAY) ---
        const getWeekdayIdxGlobal = (str) => {
            let s = String(str).trim().toLowerCase();
            const w = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
            let idx = w.indexOf(s);
            if (idx !== -1) return idx;
            if (s === 'cn' || s === 'chá»§ nháº­t' || s === 'sun') return 0;
            if (s === 't2' || s === 'thá»© 2' || s === 'thá»© hai' || s === 'mon') return 1;
            if (s === 't3' || s === 'thá»© 3' || s === 'thá»© ba' || s === 'tue') return 2;
            if (s === 't4' || s === 'thá»© 4' || s === 'thá»© tÆ°' || s === 'wed') return 3;
            if (s === 't5' || s === 'thá»© 5' || s === 'thá»© nÄƒm' || s === 'thu') return 4;
            if (s === 't6' || s === 'thá»© 6' || s === 'thá»© sÃ¡u' || s === 'fri') return 5;
            if (s === 't7' || s === 'thá»© 7' || s === 'thá»© báº£y' || s === 'sat') return 6;
            return -1;
        };

        let targetDateStr = document.getElementById('targetDeliveryDate') ? document.getElementById('targetDeliveryDate').value.trim() : "";
        let isWeekendDelivery = false;
        let targetTimestamp = 0; // Äá»ƒ tÃ­nh toÃ¡n Lead Time Arrival

        if (targetDateStr !== "") {
            let isTgtWkday = getWeekdayIdxGlobal(targetDateStr) !== -1;
            let tgtNum = isTgtWkday ? getWeekdayIdxGlobal(targetDateStr) : parseInt((targetDateStr.match(/^(\d{1,2})/) || [])[1] || 0);
            let finalWkday = -1;

            let dTarget = new Date();
            dTarget.setHours(0, 0, 0, 0);

            if (isTgtWkday) {
                finalWkday = tgtNum;
                // TÃ¬m ngÃ y gáº§n nháº¥t khá»›p vá»›i thá»© Ä‘Æ°á»£c chá»n (vÃ­ dá»¥ Thá»© 6 gáº§n nháº¥t)
                let diff = (tgtNum - dTarget.getDay() + 7) % 7;
                // Náº¿u diff = 0 thÃ¬ cÃ³ thá»ƒ lÃ  hÃ´m nay, nhÆ°ng thÆ°á»ng lÃ  Ä‘áº·t cho tuáº§n sau hoáº·c hÃ´m nay váº«n tÃ­nh sales?
                // Giá»¯ nguyÃªn logic cÅ© cho finalWkday nhÆ°ng tÃ­nh thÃªm timestamp
                dTarget.setDate(dTarget.getDate() + diff);
            } else if (tgtNum > 0) {
                // Náº¿u ngÃ y gÃµ < hÃ´m nay quÃ¡ nhiá»u (vÃ­ dá»¥ nay 28, gÃµ 2) -> Sang thÃ¡ng sau
                if (tgtNum < dTarget.getDate() - 7) {
                    dTarget.setMonth(dTarget.getMonth() + 1);
                }
                dTarget.setDate(tgtNum);
                finalWkday = dTarget.getDay();
            }

            targetTimestamp = dTarget.getTime();

            // LÆ¯U Láº I NGÃ€Y GIAO HÃ€NG Äá»‚ LÆ¯U TRá»®
            const year = dTarget.getFullYear();
            const month = String(dTarget.getMonth() + 1).padStart(2, '0');
            const day = String(dTarget.getDate()).padStart(2, '0');
            currentDeliveryDateStr = `${year}-${month}-${day}`;
            saveToDB('soq_latest_delivery_date', currentDeliveryDateStr);

            // Cuá»‘i tuáº§n: Thá»© 7 (6), Chá»§ nháº­t (0)
            if (finalWkday === 6 || finalWkday === 0) {
                isWeekendDelivery = true;
            }
        } else {
            let dTarget = new Date();
            const year = dTarget.getFullYear();
            const month = String(dTarget.getMonth() + 1).padStart(2, '0');
            const day = String(dTarget.getDate()).padStart(2, '0');
            currentDeliveryDateStr = `${year}-${month}-${day}`;
            saveToDB('soq_latest_delivery_date', currentDeliveryDateStr);
        }

        // ----------- 1. Map Rules (WM Name -> ODA Name) -----------
        const mappingMap = new Map();
        const standardNamesSet = new Set(); // LÆ°u danh sÃ¡ch TÃªn ODA chuáº©n
        const unmappedProducts = new Set(); // Tracking sáº£n pháº©m chÆ°a Ä‘Æ°á»£c mapping
        const reverseMappingKeys = new Set(); // DÃ¹ng Ä‘á»ƒ kiá»ƒm tra sáº£n pháº©m láº¡
        const productCategoryMap = new Map(); // LÆ°u nhÃ³m hÃ ng máº£ng Penalty

        if (datasets.mapping_raw && datasets.mapping_raw.length > 0) {
            let headerRow = datasets.mapping_raw[0] || [];
            let iOda = 1, iWm = 2, iCat = 3;

            // Nháº­n diá»‡n tá»± Ä‘á»™ng cá»™t báº±ng TÃªn Header
            for (let c = 0; c < headerRow.length; c++) {
                let h = String(headerRow[c]).toUpperCase();
                if (h.includes('ODA')) iOda = c;
                else if (h.includes('WM')) iWm = c;
                else if (h.includes('NHÃ“M')) iCat = c;
            }

            // Báº¯t Ä‘áº§u Ä‘á»c tá»« dÃ²ng sá»‘ 2 (Bá» qua Header)
            for (let i = 1; i < datasets.mapping_raw.length; i++) {
                let r = datasets.mapping_raw[i];
                if (!r || !Array.isArray(r)) continue;

                let odaName = r[iOda] ? String(r[iOda]).trim() : '';
                let wmName = r[iWm] ? String(r[iWm]).trim().toLowerCase() : '';
                let category = r[iCat] ? String(r[iCat]).trim().toUpperCase() : '';

                // Náº¿u ko cÃ³ Header (file trá»‘ng trÆ¡n 2 cá»™t), cháº¡y fallback truyá»n thá»‘ng
                if (!odaName && !wmName && r.length >= 2) {
                    wmName = r[0] ? String(r[0]).trim().toLowerCase() : '';
                    odaName = r[1] ? String(r[1]).trim() : '';
                }

                if (wmName && odaName && wmName !== 'tÃªn sáº£n pháº©m wm') {
                    mappingMap.set(wmName, odaName);
                    standardNamesSet.add(odaName.trim().toLowerCase());
                    reverseMappingKeys.add(wmName);

                    if (category && category !== 'NHÃ“M HÃ€NG') {
                        productCategoryMap.set(odaName.trim().toLowerCase(), category);
                    }
                }
            }
        }

        const normalizeProductName = (name) => {
            let n = String(name).trim().toLowerCase();
            // 1. Náº¿u lÃ  TÃªn WM -> Tráº£ vá» TÃªn ODA chuáº©n
            if (mappingMap.has(n)) return String(mappingMap.get(n)).trim();
            // 2. Náº¿u chÃ­nh nÃ³ Ä‘Ã£ lÃ  TÃªn ODA chuáº©n -> Tráº£ vá» chÃ­nh nÃ³
            if (standardNamesSet.has(n)) return String(name).trim();

            // Náº¿u cÃ³ náº¡p file mapping mÃ  khÃ´ng tháº¥y mÃ£ nÃ y -> Coi nhÆ° khÃ´ng há»£p lá»‡ (Tráº£ vá» null Ä‘á»ƒ lá»c bá»)
            if (datasets.mapping_raw && datasets.mapping_raw.length > 0) return null;
            return String(name).trim(); // Fallback náº¿u chÆ°a náº¡p mapping
        }

        // --- 2. Schedule Filter & Store Names ---
        const validSAPs = new Set();
        const storeNamesMap = new Map();
        const storeRegionMap = new Map();
        const storeAliasesMap = new Map(); // ID -> Set of normalized names/nicknames
        const scheduleLeadtimeMap = new Map();
        const storeTierMap = new Map();

        if (datasets.schedule && datasets.schedule.length > 0) {
            datasets.schedule.forEach(row => {
                let store = row['sap'] || row['storekey'] || row['storecode'] || row['makho'] || row['mach'] || row['mÃ£khÃ¡chhÃ ng'] || row['mÃ£cá»­ahÃ ng'] || row['nickname'] || row['storename'] || row['store'];
                if (!store) return;

                let storeID = extractSAP(store);
                let region = String(row['khuvuc'] || row['khuvá»±c'] || row['region'] || 'KhÃ¡c').trim();
                storeRegionMap.set(storeID, region);
                let hinhThuc = String(row['hinhthuc'] || row['HÃ¬nh thá»©c'] || row['type'] || '').toUpperCase();

                let dynamicLT = 0;
                const getWeekdayIdx = getWeekdayIdxGlobal;

                if (targetDateStr !== "") {
                    let hasDelivery = false;
                    let isTargetWeekday = getWeekdayIdx(targetDateStr) !== -1;
                    let impliedWeekdayIdx = new Date(targetTimestamp).getDay();
                    let currentTargetNum = isTargetWeekday ? getWeekdayIdx(targetDateStr) : new Date(targetTimestamp).getDate();

                    let possibleNextDeliveryTimestamps = [];

                    // Khá»Ÿi táº¡o biáº¿n kiá»ƒm tra Chá»©c nÄƒng (Function) cá»§a Store
                    let isMer = String(row['function'] || row['Function'] || row['chá»©c nÄƒng'] || row['loáº¡i'] || '').trim().toLowerCase() === 'mer';

                    for (const [key, val] of Object.entries(row)) {
                        let k = String(key).trim();
                        let match = false;
                        let headerTs = 0;

                        let headerWeekdayIdx = getWeekdayIdx(k);

                        // Náº¿u Header file Lá»‹ch lÃ  THá»¨ (VD: Friday, T2)
                        if (headerWeekdayIdx !== -1) {
                            if (isTargetWeekday) {
                                match = (headerWeekdayIdx === currentTargetNum);
                            } else if (impliedWeekdayIdx !== -1) {
                                match = (headerWeekdayIdx === impliedWeekdayIdx);
                            }
                        } else {
                        // Xá»­ lÃ½ Header phá»©c há»£p (vd: 01-Thg4_Wednesday) hoáº·c Header Ä‘Æ¡n thuáº§n
                        let kClean = k.toLowerCase();
                        
                        // Láº¥y sá»‘ ngÃ y cá»§a má»¥c tiÃªu (VD: 1 hoáº·c 01)
                        let tNum = new Date(targetTimestamp).getDate().toString();
                        let tPadded = tNum.padStart(2, '0');

                        // 1. So khá»›p Sá»‘ ngÃ y trá»±c tiáº¿p: "01", "1", "1-", "01-"
                        let dateMatch = kClean.startsWith(tNum + '-') || kClean.startsWith(tPadded + '-') || 
                                       kClean.includes('_' + tNum + '-') || kClean.includes('_' + tPadded + '-');
                        
                        // 2. So khá»›p Sá»‘ ngÃ y viáº¿t liá»n (VÃ­ dá»¥: 01thg4)
                        if (!dateMatch) {
                            let m = kClean.match(/^(\d{1,2})/);
                            if (m && (m[1] === tNum || m[1] === tPadded)) dateMatch = true;
                        }

                        // 3. So khá»›p Serial Date náº¿u cÃ³ trong Key
                        let serialMatch = false;
                        let serialInKey = kClean.match(/(\d{5})/);
                        if (serialInKey) {
                            headerTs = parseDateStrToTime(Number(serialInKey[1]));
                            if (targetTimestamp > 0 && headerTs > 0) {
                                let d1 = new Date(targetTimestamp);
                                let d2 = new Date(headerTs);
                                serialMatch = (d1.getFullYear() === d2.getFullYear() && d1.getMonth() === d2.getMonth() && d1.getDate() === d2.getDate());
                            }
                        }

                        // NEW: TrÃ­ch xuáº¥t Timestamp cho táº¥t cáº£ cÃ¡c cá»™t náº¿u cÃ³ Ä‘á»‹nh dáº¡ng ngÃ y (vd: 01-thg4)
                        if (headerTs === 0) {
                            // Thá»­ bÃ³c tÃ¡ch ngÃ y/thÃ¡ng tá»« chuá»—i "01-thg4"
                            let mDate = kClean.match(/^(\d{1,2})[^\d]+(\d{1,2})/);
                            if (mDate) {
                                let dd = parseInt(mDate[1]);
                                let mm = parseInt(mDate[2]) - 1;
                                let yyyy = new Date(targetTimestamp).getFullYear();
                                let dTemp = new Date(yyyy, mm, dd);
                                // Náº¿u ngÃ y quÃ¡ xa má»¥c tiÃªu (vd: thÃ¡ng 12 so vá»›i thÃ¡ng 1), lÃ¹i/tiáº¿n nÄƒm
                                headerTs = dTemp.getTime();
                            } else {
                                // Thá»­ bÃ³c tÃ¡ch ngÃ y Ä‘Æ¡n thuáº§n (vd: 01) -> Giáº£ Ä‘á»‹nh cÃ¹ng thÃ¡ng/nÄƒm vá»›i target
                                let mDay = kClean.match(/^(\d{1,2})/);
                                if (mDay) {
                                    let dd = parseInt(mDay[1]);
                                    let tDate = new Date(targetTimestamp);
                                    let dTemp = new Date(tDate.getFullYear(), tDate.getMonth(), dd);
                                    // Xá»­ lÃ½ rollover thÃ¡ng náº¿u cáº§n (vd: target lÃ  31/3, header lÃ  1)
                                    if (dd < tDate.getDate() - 15) dTemp.setMonth(dTemp.getMonth() + 1);
                                    if (dd > tDate.getDate() + 15) dTemp.setMonth(dTemp.getMonth() - 1);
                                    headerTs = dTemp.getTime();
                                }
                            }
                        }
                        
                        // Æ¯U TIÃŠN: Náº¿u Header chá»©a thÃ´ng tin NGÃ€Y Cá» Äá»ŠNH, nÃ³ sáº½ ghi Ä‘Ã¨ viá»‡c so khá»›p THá»¨ chung chung
                        if (dateMatch || serialMatch) {
                            match = true;
                        } else if (!isTargetWeekday && headerWeekdayIdx === -1) {
                            // Fallback náº¿u headers quÃ¡ Ä‘Æ¡n giáº£n (chá»‰ "1", "2")
                            match = (k === tNum || k === tPadded || k.startsWith(tNum + '/') || k.startsWith(tPadded + '/'));
                        }
                        }

                        let v = String(val).trim().toLowerCase().replace(/\s+/g, '');
                        let isDeliveryFound = false;

                        if (v && v !== '0' && v !== 'false' && v !== 'off' && !v.includes('nghá»‰')) {
                            if (isMer) {
                                // Rule Function Mer: Chá»‹u trÃ¡ch nhiá»‡m giao dá»‹ch náº¿u cÃ³ máº·t NVCH
                                // Tá»« chá»‘i nhá»¯ng CH Ä‘i thÄƒm (chá»‰ ghi "NVCH"). Pháº£i ghi "Shipper+NVCH" hoáº·c cÃ³ dáº¥u "+"
                                if ((v.includes('shipper') && v.includes('nvch')) || (v.includes('nvch') && v.includes('+')) || v.includes('giao')) {
                                    isDeliveryFound = true;
                                } else if (v === 'x' || v === 'yes' || v === 'true') {
                                    isDeliveryFound = true; // Fallback an toÃ n
                                }
                            } else {
                                // Náº¿u khÃ´ng pháº£i Function Mer (hoáº·c khÃ´ng cÃ³ cá»™t Function), má»i tÃ­n hiá»‡u nhÆ° Shipper, X Ä‘á»u tÃ­nh
                                isDeliveryFound = true;
                            }
                        }

                        if (isDeliveryFound) {
                            if (match) {
                                hasDelivery = true;
                            }
                            // Theo dÃµi táº¥t cáº£ cÃ¡c má»‘c cÃ³ giao hÃ ng tiáº¿p theo (Dáº¡ng Timestamp)
                            if (headerTs > 0) {
                                possibleNextDeliveryTimestamps.push(headerTs);
                            } else if (headerWeekdayIdx !== -1) {
                                // Náº¿u lÃ  THá»¨, quy Ä‘á»•i sang timestamp tÆ°Æ¡ng á»©ng trong tuáº§n Ä‘Ã³/tuáº§n sau
                                let dTarget = new Date(targetTimestamp);
                                let diff = (headerWeekdayIdx - dTarget.getDay() + 7) % 7;
                                let dNext = new Date(dTarget);
                                dNext.setDate(dNext.getDate() + diff);
                                possibleNextDeliveryTimestamps.push(dNext.getTime());
                            }
                        }
                    }

                    // Náº¿u khÃ´ng cÃ³ lá»‹ch giao -> Bá» qua
                    if (!hasDelivery) return;

                    // --- TÃNH TOÃN LEADTIME Äá»˜NG Tá»ª MA TRáº¬N Lá»ŠCH GIAO HÃ€NG (Dáº¡ng Timestamp) ---
                    let futureDates = possibleNextDeliveryTimestamps.filter(t => t > targetTimestamp + 3600000); // CÃ¡ch Ã­t nháº¥t 1h
                    if (futureDates.length > 0) {
                        let nextTS = Math.min(...futureDates);
                        dynamicLT = Math.round((nextTS - targetTimestamp) / 86400000);
                    }
                }

                // Máº·c Ä‘á»‹nh: Cháº¥p nháº­n Táº¤T Cáº¢ cÃ¡c mÃ£ cá»­a hÃ ng miá»…n lÃ  cÃ³ tÃªn trong file Lá»‹ch Giao HÃ ng
                if (storeID) {
                    validSAPs.add(storeID);

                    let sName = row['tencuahang'] || row['tncahng'] || row['storename'] || row['store'] || row['nickname'] || ''; 
                    let nickname = row['nickname'] || '';

                    if (sName) storeNamesMap.set(storeID, String(sName).trim());

                    // ÄÄƒng kÃ½ Alias
                    if (!storeAliasesMap.has(storeID)) storeAliasesMap.set(storeID, new Set());
                    if (sName) storeAliasesMap.get(storeID).add(normalizeKey(sName));
                    if (nickname) storeAliasesMap.get(storeID).add(normalizeKey(nickname));
                    storeAliasesMap.get(storeID).add(normalizeKey(storeID));

                    // LÆ¯U Cá»˜T TIER
                    let tierVal = String(row['tier'] || row['Tier'] || row['cáº¥pÄ‘á»™'] || row['phÃ¢nloáº¡i'] || '').trim().toUpperCase();
                    if (tierVal && tierVal !== 'UNDEFINED') storeTierMap.set(storeID, tierVal);

                    if (dynamicLT > 0) {
                        scheduleLeadtimeMap.set(storeID, dynamicLT); 
                    } else {
                        let lt = Number(row['leadtime'] || row['Leadtime'] || row['chu ká»³'] || row['chuká»³'] || 0);
                        if (lt > 0) scheduleLeadtimeMap.set(storeID, lt);
                    }
                }
            });
        }

        // Helper: BÃ³c tÃ¡ch Leadtime tá»« tÃªn file Lá»‹ch Giao HÃ ng (VD: Lá»‹ch 2003-2203 -> 3 ngÃ y)
        const extractLeadtimeFromFilename = (filename) => {
            let match = filename.match(/(\d{2})(\d{2})-(\d{2})(\d{2})/);
            if (match) {
                let d1 = parseInt(match[1], 10);
                let m1 = parseInt(match[2], 10) - 1;
                let d2 = parseInt(match[3], 10);
                let m2 = parseInt(match[4], 10) - 1;

                let y1 = new Date().getFullYear();
                let y2 = y1;
                if (m1 === 11 && m2 === 0) y2 = y1 + 1; // Wrap around year

                let date1 = new Date(y1, m1, d1);
                let date2 = new Date(y2, m2, d2);

                let diffTime = Math.abs(date2 - date1);
                let diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                return diffDays + 1; // Include both start and end days
            }
            return 1; // Fallback
        };

        // Táº O Báº¢N Äá»’ NGÆ¯á»¢C Sá»šM: TÃªn Store (Chuáº©n hÃ³a) / Nickname -> MÃ£ SAP Ä‘á»ƒ xá»­ lÃ½ Tá»“n Kho & Nháº­p
        const reverseStoreNamesMap = new Map();
        const buildReverseMap = () => {
            storeAliasesMap.forEach((aliases, id) => {
                aliases.forEach(alias => {
                    reverseStoreNamesMap.set(alias, id);
                });
            });
            storeNamesMap.forEach((name, id) => {
                reverseStoreNamesMap.set(normalizeKey(name), id);
                reverseStoreNamesMap.set(id, id);
            });
        };
        // Build láº§n 1: Láº¥y dá»¯ liá»‡u Alias tá»« file Lá»‹ch giao hÃ ng (Schedule) lÃ m gá»‘c
        buildReverseMap();

        const resolveStoreID = (rawSap, nick) => {
            let finalID = "";
            let extracted = extractSAP(rawSap);
            if (extracted && !isNaN(parseInt(extracted))) {
                finalID = extracted;
            } else {
                let nKey = normalizeKey(nick);
                let lookedUp = reverseStoreNamesMap.get(nKey);
                if (!lookedUp) {
                    for (let [alias, id] of reverseStoreNamesMap.entries()) {
                        if (alias && nKey && (alias.includes(nKey) || nKey.includes(alias))) {
                            if (alias.length > 5 || nKey.length > 5) { // TrÃ¡nh nháº§m láº«n chá»¯ táº¯t quÃ¡ ngáº¯n
                                lookedUp = id;
                                break;
                            }
                        }
                    }
                }
                finalID = lookedUp ? lookedUp : extractSAP(nick);
            }
            return finalID;
        }

        const trendReportMap = new Map();
        if (datasets.trend_report && datasets.trend_report.length > 0) {
            datasets.trend_report.forEach(row => {
                let st = row['sap'] || row['storecode'] || row['store'] || row['mach'] || row['tencuahang'] || row['tÃªncá»­ahÃ ng'];
                let pr = row['productnameprimarylanguage'] || row['productname'] || row['product'] || row['tensanpham'] || row['productnameprimarylanguage'];
                
                if (!pr) {
                    let foundKey = Object.keys(row).find(k => k.includes('product') || k.includes('name') || k.includes('sanpham'));
                    if (foundKey) pr = row[foundKey];
                }
                if (!st) {
                    let foundKey = Object.keys(row).find(k => k.includes('sap') || k.includes('store') || k.includes('mach'));
                    if (foundKey) st = row[foundKey];
                }
                
                let action = row['action'] || row['hanhdong'] || row['ghi_chu'] || row['ghichu'];
                if (!action) {
                    let foundKey = Object.keys(row).find(k => k.includes('action') || k.includes('hanhdong') || k.includes('xu_huong') || k.includes('xuhuong'));
                    if (foundKey) action = row[foundKey];
                }

                if (st && pr && action) {
                    let storeID = resolveStoreID(st, st);
                    let prodStd = normalizeProductName(pr);
                    if (prodStd) {
                        let key = `${storeID}_${prodStd.toLowerCase()}`;
                        trendReportMap.set(key, String(action).trim());
                    }
                }
            });
        }

        // --- BÆ¯á»šC 0: TÃŒM NGÃ€Y Lá»šN NHáº¤T Cá»¦A Tá»ªNG STORE LÃ€M Má»C (T) ---
        const storeMaxInvDateMap = new Map();
        const storeMaxOrderDateMap = new Map();

        if (datasets.inventory && datasets.inventory.length > 0) {
            datasets.inventory.forEach(row => {
                let store = row['sap'] || row['storecode'] || row['nickname'] || row['storename'] || row['store'] || row['mach'] || row['article'];
                if (!store) return;
                let storeID = extractSAP(store);
                if (storeID && isNaN(parseInt(storeID))) {
                    let lookedUp = reverseStoreNamesMap.get(normalizeKey(store));
                    if (lookedUp) storeID = lookedUp;
                }
                let rawDate = row['date'] || row['Date'] || row['ngay'] || row['ngÃ y'] || 0;
                let cDate = parseDateStrToTime(rawDate);
                if (cDate > 0) {
                    let currentMax = storeMaxInvDateMap.get(storeID) || 0;
                    if (cDate > currentMax) storeMaxInvDateMap.set(storeID, cDate);
                }
            });
        }

        if (datasets.input && datasets.input.length > 0) {
            datasets.input.forEach(row => {
                let rawSap = extractSAP(row['sap'] || row['storecode'] || row['mach'] || row['macuahang'] || row['sapcode']);
                let nick = row['nickname'] || row['storename'] || row['store'] || row['tencuahang'] || row['sap'] || '';
                let storeID = resolveStoreID(rawSap, nick);
                if (!storeID) return;
                let rawDate = row['orderdate'] || row['Order date'] || row['completeddate'] || row['Completed date'] || row['date'] || row['ngaydathang'] || row['ngay'] || row['ngaytao'] || row['createddate'] || 0;
                let cDate = parseDateStrToTime(rawDate);
                if (cDate > 0) {
                    let currentMax = storeMaxOrderDateMap.get(storeID) || 0;
                    if (cDate > currentMax) storeMaxOrderDateMap.set(storeID, cDate);
                }
            });
        }

        const storeMasterDateMap = new Map();
        for (let storeID of new Set([...storeMaxInvDateMap.keys(), ...storeMaxOrderDateMap.keys()])) {
            let sInvDate = storeMaxInvDateMap.get(storeID) || 0;
            let sOrderDate = storeMaxOrderDateMap.get(storeID) || 0;
            let sDeliveryDate = sOrderDate > 0 ? sOrderDate + 86400000 : 0;

            let T = 0;
            if (sInvDate > 0 && sDeliveryDate > 0) {
                T = Math.max(sInvDate, sDeliveryDate);
            } else if (sInvDate > 0) {
                T = sInvDate;
            } else {
                T = sDeliveryDate;
            }
            storeMasterDateMap.set(storeID, T);
        }

        // --- 3. Inventory Aggregation ---
        const inventoryMap = new Map();
        if (datasets.inventory && datasets.inventory.length > 0) {
            datasets.inventory.forEach(row => {
                let store = row['sap'] || row['storecode'] || row['nickname'] || row['storename'] || row['store'] || row['mach'];
                let prod = row['productname'] || row['listsnphm'] || row['tnsnphm'] || row['tensanphamwm'] || row['articlename'] || row['article'] || row['tensanpham'] || row['productname'];
                if (!store || !prod) return;

                let storeID = extractSAP(store);
                if (storeID && isNaN(parseInt(storeID))) {
                    let lookedUp = reverseStoreNamesMap.get(normalizeKey(store));
                    if (lookedUp) storeID = lookedUp;
                }
                let sName = row['tencuahang'] || row['tncahng'] || row['storename'] || row['store'];
                if (sName && !storeNamesMap.has(storeID)) storeNamesMap.set(storeID, String(sName).trim());

                let rawDate = row['date'] || row['Date'] || row['ngay'] || row['ngÃ y'] || 0;
                let cDate = parseDateStrToTime(rawDate);

                let T = storeMasterDateMap.get(storeID);
                if (!T || cDate > T) return;

                let prodStd = normalizeProductName(prod);
                if (!prodStd) {
                    unmappedProducts.add(String(prod).trim());
                    return;
                }
                let key = `${storeID}_${prodStd.toLowerCase()}`;

                // ... (Quy Ä‘á»•i kg)
                let inv = Number(String(row['tonkho'] || row['stock'] || row['ton'] || row['inventory'] || row['inventoryquantity'] || row['inventoryamount'] || row['stkinv'] || '0').replace(/,/g, ''));
                let disp = Number(String(row['huy'] || row['disposal'] || row['scrap'] || row['tonhuy'] || row['disposalquantity'] || row['disposalamount'] || '0').replace(/,/g, ''));

                if (prod && String(prod).toLowerCase().includes('retail kg')) {
                    inv = inv / 1000;
                    disp = disp / 1000;
                }

                if (!inventoryMap.has(key)) {
                    inventoryMap.set(key, {
                        currentInv: 0, currentDisp: 0,
                        prevInv: 0, prevInvDate: 0,
                        prodOrig: prodStd
                    });
                }

                let data = inventoryMap.get(key);
                if (cDate === T) {
                    data.currentInv += inv;
                    data.currentDisp += disp;
                } else if (cDate < T) {
                    // LÆ°u dá»¯ liá»‡u cá»§a ngÃ y gáº§n T nháº¥t
                    if (cDate > data.prevInvDate) {
                        data.prevInvDate = cDate;
                        data.prevInv = inv;
                    } else if (cDate === data.prevInvDate) {
                        data.prevInv += inv;
                    }
                }
            });
        }

        // ----------- 4. Input ODA Aggregation -----------
        const inputMap = new Map();
        const actualODA_Names = new Map(); // LÆ°u TÃªn ODA chuáº©n nháº¥t tá»« file váº­n hÃ nh

        if (datasets.input && datasets.input.length > 0) {
            datasets.input.forEach(row => {
                let prod = row['productnameprimarylanguage'] || row['productname'] || row['product'] || row['tensanphamwm'] || row['tensanpham'] || row['articlename'] || row['article'];
                let status = String(row['orderstatus'] || row['status'] || row['trangthai'] || '').toLowerCase();
                
                if (!prod) return;
                // Lá»c bá» hÃ ng Há»§y / ÄÃ£ hoÃ n (Chá»‰ láº¥y Completed)
                if (status && (status.includes('cancel') || status.includes('há»§y') || status.includes('reject'))) return;

                let rawSap = extractSAP(row['sap'] || row['storecode'] || row['mach'] || row['macuahang'] || row['sapcode']);
                let nick = row['nickname'] || row['storename'] || row['store'] || row['tencuahang'] || row['sap'] || '';
                let storeID = resolveStoreID(rawSap, nick);
                if (!storeID) return;
                let exactODAName = String(prod).trim();
                let prodStd = normalizeProductName(prod);
                if (!prodStd) {
                    unmappedProducts.add(String(prod).trim());
                    return;
                }

                let key = `${storeID}_${prodStd.toLowerCase()}`;
                actualODA_Names.set(prodStd.toLowerCase(), exactODAName);

                let dQty = row['deliveredqty'] !== undefined ? row['deliveredqty'] : row['slgiao'];
                let valStr = "";
                if (dQty !== undefined && String(dQty).trim() !== "") {
                    valStr = String(dQty);
                } else {
                    let oQty = row['orderedqty'] || row['orderqty'] || row['quantity'] || row['orderitemqty'] || row['quantityorder'] || row['sldat'] || row['sldathang'] || row['totalqty'] || row['soluong'] || row['soluongnhap'] || row['inputquantity'] || row['sum'] || row['total'] || row['qty'];
                    valStr = String(oQty || '0');
                }
                let qty = Number(valStr.replace(/,/g, ''));

                // TrÃ­ch xuáº¥t ngÃ y giao hÃ ng/nháº­p hÃ ng
                let rawDate = row['orderdate'] || row['Order date'] || row['completeddate'] || row['Completed date'] || row['date'] || row['ngaydathang'] || row['ngay'] || row['ngaytao'] || row['createddate'] || 0;
                let cOrderDate = parseDateStrToTime(rawDate);
                let cDeliveryDate = cOrderDate > 0 ? cOrderDate + 86400000 : 0; // Cá»™ng thÃªm 1 ngÃ y giao

                let T = storeMasterDateMap.get(storeID);
                if (!T || cDeliveryDate > T) return;

                if (!inputMap.has(key)) {
                    inputMap.set(key, {
                        currentInput: 0,
                        prevInput: 0, prevInputDate: 0,
                        prodOrig: exactODAName
                    });
                }

                let current = inputMap.get(key);
                if (cDeliveryDate === T) {
                    current.currentInput += qty;
                } else if (cDeliveryDate < T) {
                    if (cDeliveryDate > current.prevInputDate) {
                        current.prevInputDate = cDeliveryDate;
                        current.prevInput = qty;
                    } else if (cDeliveryDate === current.prevInputDate) {
                        current.prevInput += qty;
                    }
                }
            });
        }

        // HÃ m láº¥y láº¡i TÃªn Chuáº©n nháº¥t (Æ¯u tiÃªn ODA tháº­t > Mapping > Raw)
        const getBestAvailableName = (mappedName) => {
            if (!mappedName) return '';
            let k = String(mappedName).toLowerCase();
            return actualODA_Names.has(k) ? actualODA_Names.get(k) : mappedName;
        }

        // ----------- 5. Sales Data (Flat Transaction Aggregation) -----------
        // Trong file thá»±c táº¿: Dá»¯ liá»‡u doanh sá»‘ bÃ¡n náº±m tá»«ng dÃ²ng, cá»™t "POS Quantity"
        const monthlySales = new Map();
        const storeMonthlyDays = new Map(); // All days
        const storeGroupDays = new Map();  // storeID -> { weekdays: Set, weekends: Set }
        const globalMonthlyDays = new Set();
        const globalMonthlyGroupDays = { weekdays: new Set(), weekends: new Set() };
        let globalMonthlyMaxTs = 0;

        const processMonthlyData = (dataArr) => {
            if (!dataArr || dataArr.length === 0) return;
            dataArr.forEach(row => {
                let st = row['sap'] || row['storecode'] || row['sapcode'] || row['store'] || row['nickname'] || row['storename'] || row['tencuahang'];
                let pr = row['tnsnphmwm'] || row['tensanphamwm'] || row['tnsnphm'] || row['tensanpham'] || row['productname'] || row['articlename'] || row['article'];
                let qty = Number(String(row['posquantity'] || row['quantity'] || row['soluong'] || row['sum'] || '0').replace(/,/g, ''));
                if (pr && String(pr).toLowerCase().includes('retail kg')) qty /= 1000;

                let storeID = extractSAP(st);
                
                // Há»— trá»£ Fallback Lookup cho Monthly Sales y chang Weekly
                if (storeID && isNaN(parseInt(storeID))) {
                    let lookedUp = reverseStoreNamesMap.get(normalizeKey(st));
                    if (lookedUp) storeID = lookedUp;
                }

                let rawDate = String(row['calendarday'] || row['date'] || row['ngay'] || '').trim();

                if (rawDate) {
                    let cbDate = parseDateStrToTime(rawDate);
                    if (cbDate > 0) {
                        globalMonthlyDays.add(cbDate);
                        if (cbDate > globalMonthlyMaxTs) globalMonthlyMaxTs = cbDate;
                        let cbDayOfWeek = new Date(cbDate).getDay();
                        if (cbDayOfWeek === 6 || cbDayOfWeek === 0) {
                            globalMonthlyGroupDays.weekends.add(cbDate);
                        } else {
                            globalMonthlyGroupDays.weekdays.add(cbDate);
                        }
                    }
                }

                // ÄÄƒng kÃ½ TÃªn/Nickname tá»« file Doanh sá»‘ (ODA)
                if (storeID) {
                    let sName = row['storename'] || row['store'] || row['tÃªncá»­ahÃ ng'] || '';
                    let nickname = row['nickname'] || '';
                    if (!storeAliasesMap.has(storeID)) storeAliasesMap.set(storeID, new Set());
                    if (sName) storeAliasesMap.get(storeID).add(normalizeKey(sName));
                    if (nickname) storeAliasesMap.get(storeID).add(normalizeKey(nickname));
                    storeAliasesMap.get(storeID).add(normalizeKey(storeID));
                    
                    if (sName && !storeNamesMap.has(storeID)) storeNamesMap.set(storeID, sName);
                }

                if (rawDate && storeID) {
                    if (!storeMonthlyDays.has(storeID)) storeMonthlyDays.set(storeID, new Set());
                    storeMonthlyDays.get(storeID).add(rawDate);

                    if (!storeGroupDays.has(storeID)) {
                        storeGroupDays.set(storeID, { weekdays: new Set(), weekends: new Set() });
                    }
                    let cDate = parseDateStrToTime(rawDate);
                    let dayOfWeek = new Date(cDate).getDay();
                    let isWknd = (dayOfWeek === 6 || dayOfWeek === 0);
                    if (isWknd) storeGroupDays.get(storeID).weekends.add(rawDate);
                    else storeGroupDays.get(storeID).weekdays.add(rawDate);

                    if (!st || !pr || isNaN(qty)) return;

                    let prodStd = normalizeProductName(pr);
                    if (!prodStd) {
                        unmappedProducts.add(String(pr).trim());
                        return;
                    }
                    let key = `${storeID}_${prodStd.toLowerCase()}`;

                    if (!monthlySales.has(key)) {
                        monthlySales.set(key, {
                            storeOrig: st,
                            prodStd,
                            totalQty: qty,
                            weekdayQty: isWknd ? 0 : qty,
                            weekendQty: isWknd ? qty : 0,
                            minDateTs: qty > 0 ? cDate : 0
                        });
                    } else {
                        let data = monthlySales.get(key);
                        data.totalQty += qty;
                        if (isWknd) data.weekendQty += qty;
                        else data.weekdayQty += qty;
                        if (qty > 0 && cDate > 0 && (data.minDateTs === 0 || cDate < data.minDateTs)) {
                            data.minDateTs = cDate;
                        }
                    }
                }
            });
        };

        processMonthlyData(datasets.monthly);
        
        // Build láº§n 2: Bá»• sung thÃªm Alias náº¿u file Doanh Thu ThÃ¡ng cÃ³ ghi nháº­n tÃªn/nickname má»›i
        buildReverseMap();

        const weeklySales = new Map();
        const storeWeeklyDays = new Map();
        const storeWeeklyGroupDays = new Map();
        const globalWeeklyDays = new Set();
        const globalWeeklyGroupDays = { weekdays: new Set(), weekends: new Set() };
        let globalWeeklyMaxTs = 0;
        if (datasets.weekly && datasets.weekly.length > 0) {
            datasets.weekly.forEach(row => {
                // Kiá»ƒm tra xem Ä‘Ã¢y lÃ  file TRANSACTION (pháº³ng) hay MATRIX (ngang)
                let st = row['sap'] || row['storecode'] || row['nickname'] || row['storename'] || row['store'] || row['mach'] || row['tencuahang'];
                let pr = row['tnsnphmwm'] || row['tensanphamwm'] || row['tnsnphm'] || row['articlename'] || row['article'] || row['tensanpham'] || row['productname'];
                
                if (!pr) return;
                let prodStd = normalizeProductName(pr);
                if (!prodStd) return;

                if (st) {
                    // --- Dáº NG FILE PHáº²NG (TRANSACTION) ---
                    let storeID = extractSAP(st);
                    
                    // Fallback cá»±c máº¡nh cho ODA: Náº¿u Ã´ Name/Nickname khÃ´ng chá»©a MÃ£ SAP dáº¡ng sá»‘, ta sáº½ lookup tá»« thÆ° viá»‡n!
                    if (storeID && isNaN(parseInt(storeID))) {
                        let lookedUp = reverseStoreNamesMap.get(normalizeKey(st));
                        if (lookedUp) storeID = lookedUp;
                    }

                    let qty = Number(String(row['posquantity'] || row['sum'] || '0').replace(/,/g, ''));
                    if (pr && String(pr).toLowerCase().includes('retail kg')) qty /= 1000;

                    let rawDate = String(row['calendarday'] || row['date'] || row['ngay'] || '').trim();
                    let isWknd = false;

                    if (rawDate) {
                        let cbDate = parseDateStrToTime(rawDate);
                        if (cbDate > 0) {
                            globalWeeklyDays.add(cbDate);
                            if (cbDate > globalWeeklyMaxTs) globalWeeklyMaxTs = cbDate;
                            let cbDayOfWeek = new Date(cbDate).getDay();
                            isWknd = (cbDayOfWeek === 6 || cbDayOfWeek === 0);
                            if (isWknd) globalWeeklyGroupDays.weekends.add(cbDate);
                            else globalWeeklyGroupDays.weekdays.add(cbDate);
                        }
                    }

                    if (rawDate && storeID) {
                        if (!storeWeeklyDays.has(storeID)) storeWeeklyDays.set(storeID, new Set());
                        storeWeeklyDays.get(storeID).add(rawDate);

                        if (!storeWeeklyGroupDays.has(storeID)) {
                            storeWeeklyGroupDays.set(storeID, { weekdays: new Set(), weekends: new Set() });
                        }
                        if (isWknd) storeWeeklyGroupDays.get(storeID).weekends.add(rawDate);
                        else storeWeeklyGroupDays.get(storeID).weekdays.add(rawDate);
                    }

                    if (isNaN(qty)) return;
                    let key = `${storeID}_${prodStd.toLowerCase()}`;
                    let cDate = parseDateStrToTime(rawDate);

                    if (!weeklySales.has(key)) {
                        weeklySales.set(key, { totalQty: qty, weekdayQty: isWknd ? 0 : qty, weekendQty: isWknd ? qty : 0, minDateTs: qty > 0 ? cDate : 0 });
                    } else {
                        let data = weeklySales.get(key);
                        data.totalQty += qty;
                        if (isWknd) data.weekendQty += qty;
                        else data.weekdayQty += qty;
                        if (qty > 0 && cDate > 0 && (data.minDateTs === 0 || cDate < data.minDateTs)) {
                            data.minDateTs = cDate;
                        }
                    }
                } else {
                    // --- Dáº NG FILE MA TRáº¬N (MATRIX - TÃªn cá»­a hÃ ng á»Ÿ tiÃªu Ä‘á» cá»™t) ---
                    // Duyá»‡t tá»«ng cá»™t cá»§a dÃ²ng nÃ y
                    Object.entries(row).forEach(([colKey, qtyVal]) => {
                        let cKey = String(colKey).trim();
                        if (!cKey) return;

                        // Æ¯U TIÃŠN 1: TÃ¬m xem trong Header cÃ³ chá»©a MÃ£ SAP (4-5 sá»‘) khÃ´ng?
                        let sID = "";
                        let sapMatch = cKey.match(/(\d{4,5})/);
                        if (sapMatch && reverseStoreNamesMap.has(normalizeKey(sapMatch[1]))) {
                            sID = reverseStoreNamesMap.get(normalizeKey(sapMatch[1]));
                        } else {
                            // Æ¯U TIÃŠN 2: TÃ¬m theo TÃªn/Nickname Ä‘Ã£ normalize
                            sID = reverseStoreNamesMap.get(normalizeKey(cKey));
                        }

                        if (sID) {
                            let qty = Number(String(qtyVal || '0').replace(/,/g, ''));
                            if (pr && String(pr).toLowerCase().includes('retail kg')) qty /= 1000;
                            if (isNaN(qty) || qty === 0) return;

                            let key = `${sID}_${prodStd.toLowerCase()}`;
                            // Vá»›i file Matrix khÃ´ng cÃ³ ngÃ y, ta máº·c Ä‘á»‹nh chia Ä‘á»u tá»‰ lá»‡ 5/2 (5 ngÃ y thÆ°á»ng, 2 ngÃ y cuá»‘i tuáº§n)
                            let wQty = qty * (5 / 7);
                            let eQty = qty * (2 / 7);

                            if (!weeklySales.has(key)) {
                                weeklySales.set(key, { totalQty: qty, weekdayQty: wQty, weekendQty: eQty });
                            } else {
                                let d = weeklySales.get(key);
                                d.totalQty += qty;
                                d.weekdayQty += wQty;
                                d.weekendQty += eQty;
                            }
                            
                            // Giáº£ láº­p sá»‘ ngÃ y (5 ngÃ y thÆ°á»ng, 2 cuá»‘i tuáº§n) Ä‘á»ƒ denominator > 0
                            if (!storeWeeklyDays.has(sID)) storeWeeklyDays.set(sID, new Set(['dummy-w1','dummy-w2','dummy-w3','dummy-w4','dummy-w5','dummy-e1','dummy-e2']));
                            if (!storeWeeklyGroupDays.has(sID)) {
                                const dummyDays = { weekdays: new Set(['d1','d2','d3','d4','d5']), weekends: new Set(['d6','d7']) };
                                storeWeeklyGroupDays.set(sID, dummyDays);
                            }
                        }
                    });
                }
            });
        }


        // ----------- Cáº¢NH BÃO MAPPING LÃŠN MÃ€N HÃŒNH CHÃNH -----------
        const warningDiv = document.getElementById('mapping-warning-div');
        if (warningDiv) {
            if (unmappedProducts.size > 0 && datasets.mapping_raw && datasets.mapping_raw.length > 0) {
                warningDiv.innerHTML = `<strong style="color: #ff9800; font-size: 1.1em;"><i class="fas fa-exclamation-triangle"></i> Cáº­p nháº­p thÃªm sáº£n pháº©m: TÃŒM THáº¤Y ${unmappedProducts.size} Sáº¢N PHáº¨M Má»šI TRONG DOANH Sá» TUáº¦N!</strong><br>
            <span style="display:block; margin-top: 8px;">DÆ°á»›i Ä‘Ã¢y lÃ  cÃ¡c mÃ£ <b>CHÆ¯A ÄÆ¯á»¢C GHI NHáº¬N</b> trong Mapping vÃ  Ä‘Ã£ bá»‹ táº¡m áº©n khá»i báº£ng SOQ: <br>
            <i style="color: #fff; background: rgba(255,255,255,0.1); padding: 5px 8px; border-radius: 4px; display: inline-block; margin-top: 5px;">${Array.from(unmappedProducts).slice(0, 15).join(', ')}${unmappedProducts.size > 15 ? '...' : ''}</i></span>`;
                warningDiv.style.display = 'block';
            } else {
                warningDiv.style.display = 'none';
            }
        }


        // ----------- 6. Final Master Processing -----------
        const allItems = new Map();

        const registerKey = (key, storeID, storeOrig, rawProdStdName) => {
            if (!allItems.has(key)) {
                let bestName = getBestAvailableName(rawProdStdName);
                // ÄÃ£ bá»• sung prodStd Ä‘á»ƒ hÃ m lá»c nhÃ³m hÃ ng cÃ³ thá»ƒ map chÃ­nh xÃ¡c
                allItems.set(key, { storeID, storeOrig, bestName, prodStd: String(rawProdStdName || '') });
            }
        };

        // 2026-03-31: Äáº£m báº£o táº¥t cáº£ store trong lá»‹ch pháº£i Ä‘Æ°á»£c xuáº¥t hiá»‡n ká»ƒ cáº£ khi chÆ°a cÃ³ sá»‘ bÃ¡n/tá»“n
        let syncKeysSet = new Set([...monthlySales.keys(), ...inventoryMap.keys(), ...inputMap.keys()]);
        let hasScheduleUploaded = datasets.schedule && datasets.schedule.length > 0;

        if (hasScheduleUploaded) {
            // Láº¥y thÃªm cÃ¡c tá»• há»£p tá»« mapping hoáº·c cÃ¡c dá»¯ liá»‡u khÃ¡c náº¿u store Ä‘Ã³ cÃ³ trong schedule
            // Duyá»‡t qua mapping hoáº·c toÃ n bá»™ danh sÃ¡ch sáº£n pháº©m Ä‘Ã£ tá»«ng tháº¥y
            let anyProdStandards = new Set([...standardNamesSet]);
            // Náº¿u chÆ°a cÃ³ mapping, láº¥y tá»« Sales/Inventory
            if (anyProdStandards.size === 0) {
                monthlySales.forEach(v => anyProdStandards.add(v.prodStd.toLowerCase()));
                inventoryMap.forEach(v => anyProdStandards.add(v.prodOrig.toLowerCase()));
                inputMap.forEach(v => anyProdStandards.add(v.prodOrig.toLowerCase()));
            }

            validSAPs.forEach(sID => {
                anyProdStandards.forEach(prodName => {
                    syncKeysSet.add(`${sID}_${prodName.toLowerCase()}`);
                });
            });
        }
        
        let syncKeys = Array.from(syncKeysSet);

        syncKeys.forEach(k => {
            let parts = k.split('_');
            let storeID = parts[0];

            // Strict Filter Lá»‹ch Giao: Náº¿u cÃ³ táº£i file Lá»‹ch lÃªn, Báº®T BUá»˜C mÃ£ cá»­a hÃ ng pháº£i cÃ³ máº·t trong validSAPs (vá»«a check ngÃ y vá»«a check cÃ³ list)
            if (hasScheduleUploaded && !validSAPs.has(storeID)) return;

            let mData = monthlySales.get(k);
            let iData = inventoryMap.get(k);
            let inData = inputMap.get(k);

            let storeOrig = mData ? mData.storeOrig : (storeID);
            let rawProdStdName = mData ? mData.prodStd : (iData ? iData.prodOrig : (inData ? inData.prodOrig : parts[1]));

            registerKey(k, storeID, storeOrig, rawProdStdName);
        });

        // Block cáº£nh bÃ¡o mapping Ä‘Ã£ Ä‘Æ°á»£c dá»i lÃªn trÃªn Ä‘á»ƒ cháº¡y sá»›m hÆ¡n

        finalResults = [];
        tbody.innerHTML = '';

        const countPeriodDays = (startTs, globalDaysSet) => {
            if (startTs === 0 || !globalDaysSet || globalDaysSet.size === 0) return null;
            let total = 0, weekdays = 0, weekends = 0;
            globalDaysSet.forEach(ts => {
                if (ts >= startTs) {
                    total++;
                    let d = new Date(ts).getDay();
                    if (d === 0 || d === 6) weekends++;
                    else weekdays++;
                }
            });
            if (total === 0) return null;
            return { total, weekdays, weekends };
        };

        allItems.forEach((data, key) => {
            // Chá»‘t sá»‘ ngÃ y thá»±c táº¿ dá»±a trÃªn Tá»”NG Sá» NGÃ€Y GHI NHáº¬N Cá»¦A TOÃ€N Bá»˜ FILE (Thay vÃ¬ chia theo tá»«ng cá»­a hÃ ng)
            let mDaysCount = globalMonthlyDays.size > 0 ? globalMonthlyDays.size : 30;
            let wDaysCount = globalWeeklyDays.size > 0 ? globalWeeklyDays.size : 7;

            // Average Daily Sales
            let mDataExt = monthlySales.get(key);
            let mTotal = mDataExt ? mDataExt.totalQty : 0;
            let wDataExt = weeklySales.get(key);
            let wTotal = wDataExt ? wDataExt.totalQty : 0;
            let wWeekdayQty = wDataExt ? wDataExt.weekdayQty : 0;
            let wWeekendQty = wDataExt ? wDataExt.weekendQty : 0;

            let wWeekdayDaysCount = globalWeeklyGroupDays.weekdays.size > 0 ? globalWeeklyGroupDays.weekdays.size : 5;
            let wWeekendDaysCount = globalWeeklyGroupDays.weekends.size > 0 ? globalWeeklyGroupDays.weekends.size : 2;

            let wWeekdayAds = wWeekdayDaysCount > 0 ? wWeekdayQty / wWeekdayDaysCount : 0;
            let wWeekendAds = wWeekendDaysCount > 0 ? wWeekendQty / wWeekendDaysCount : 0;

            // --- NEW: PhÃ¢n tÃ­ch T2-T5 vs T6-CN ---
            let weekdayQty = mDataExt ? mDataExt.weekdayQty : 0;
            let weekendQty = mDataExt ? mDataExt.weekendQty : 0;
            let weekdayDaysCount = globalMonthlyGroupDays.weekdays.size > 0 ? globalMonthlyGroupDays.weekdays.size : (mDaysCount * 5/7);
            let weekendDaysCount = globalMonthlyGroupDays.weekends.size > 0 ? globalMonthlyGroupDays.weekends.size : (mDaysCount * 2/7);

            // Ghi Ä‘Ã¨ báº±ng tuá»•i thá» cÃ¡ nhÃ¢n náº¿u lÃ  hÃ ng má»›i (DÃ² theo tá»«ng cá»­a hÃ ng - tá»«ng sáº£n pháº©m)
            if (mDataExt && mDataExt.minDateTs > 0) {
                let lifeSpan = countPeriodDays(mDataExt.minDateTs, globalMonthlyDays);
                if (lifeSpan) {
                    mDaysCount = lifeSpan.total;
                    weekdayDaysCount = lifeSpan.weekdays;
                    weekendDaysCount = lifeSpan.weekends;
                }
            }

            if (wDataExt && wDataExt.minDateTs > 0) {
                let wLifeSpan = countPeriodDays(wDataExt.minDateTs, globalWeeklyDays);
                if (wLifeSpan) {
                    wDaysCount = wLifeSpan.total;
                    wWeekdayDaysCount = wLifeSpan.weekdays;
                    wWeekendDaysCount = wLifeSpan.weekends;
                }
            }

            let weekdayAds = weekdayDaysCount > 0 ? weekdayQty / weekdayDaysCount : 0;
            let weekendAds = weekendDaysCount > 0 ? weekendQty / weekendDaysCount : 0;

            let mAds = mTotal / mDaysCount;
            let wAds = wTotal / wDaysCount;

            let trend = 0;
            let trendHtml = '-';
            let trendExport = '0%';
            let trendFactor = 1;

            if (mAds > 0) {
                trend = ((wAds - mAds) / mAds) * 100;
                trendFactor = 1 + (trend / 100);
                trendExport = `${trend > 0 ? '+' : ''}${trend.toFixed(1)}%`;
                if (trend > 0) {
                    trendHtml = `<span style="color: var(--success)">â–² ${trend.toFixed(1)}%</span>`;
                } else if (trend < 0) {
                    trendHtml = `<span style="color: var(--danger)">â–¼ ${Math.abs(trend).toFixed(1)}%</span>`;
                } else {
                    trendHtml = `<span>0%</span>`;
                }
            } else if (wAds > 0) {
                trendExport = '100% (New)';
                trendHtml = `<span style="color: var(--success)">â–² Má»›i bÃ¡n</span>`;
                trendFactor = 1; // Máº·c Ä‘á»‹nh 1 cho hÃ ng má»›i
            }

            // Náº¿u Weekly ko cÃ³ thÃ¬ dÃ¹ng Monthly lÃ m gá»‘c Ä‘á»ƒ dá»± bÃ¡o, xu hÆ°á»›ng = N/A
            if (wTotal === 0 && mTotal > 0) {
                wAds = mAds;
                trendHtml = `<span style="color: var(--text-muted)">N/A (Tuáº§n 0)</span>`;
                trendExport = 'N/A';
                trendFactor = 1;
            }

            // Sá» TRUNG BÃŒNH BÃN NGÃ€Y HOÃ€N TOÃ€N Dá»°A VÃ€O THÃNG
            let forecastDay = mAds;
            
            if (mTotal === 0 && wTotal > 0) {
                // HÃ ng siÃªu má»›i chá»‰ cÃ³ trong tuáº§n
                forecastDay = wAds;
                weekdayAds = wWeekdayAds;
                weekendAds = wWeekendAds;
            }

            // --- TÃNH TOÃN LEAD TIME Tá»”NG Cá»˜NG ---
            // 1. Lead Time Arrival: Tá»« ngÃ y T (Master Date) Ä‘áº¿n ngÃ y Giao hÃ ng (Target Delivery)
            let T = storeMasterDateMap.get(data.storeID) || 0;
            let invData = inventoryMap.get(key) || { currentInv: 0, currentDisp: 0, prevInv: 0, prevInvDate: 0 };
            let inputData = inputMap.get(key) || { currentInput: 0, prevInput: 0, prevInputDate: 0 };

            let invDate = T > 0 ? T : new Date().setHours(0, 0, 0, 0);
            let leadTimeArrival = 0;
            if (targetTimestamp > 0) {
                leadTimeArrival = Math.max(0, (targetTimestamp - invDate) / (1000 * 60 * 60 * 24));
            }

            // 2. Coverage Leadtime: Khoáº£ng cÃ¡ch giá»¯a cÃ¡c Ä‘á»£t giao (láº¥y tá»« matrix lá»‹ch)
            let coverageLT = scheduleLeadtimeMap.has(data.storeID) ? scheduleLeadtimeMap.get(data.storeID) : extractLeadtimeFromFilename(scheduleFileName);

            let totalLeadtime = leadTimeArrival + coverageLT;
            
            let basePeriodDemand = calculatePeriodDemand(invDate, totalLeadtime, weekdayAds, weekendAds);
            
            // TÃ¡ch Demand dá»± kiáº¿n lÃºc chá» hÃ ng (trÃ¡nh Ã¢m kho dá»“n vÃ o SOQ gÃ¢y overstock)
            let leadTimeDemandBase = calculatePeriodDemand(invDate, leadTimeArrival, weekdayAds, weekendAds);
            let demandLeadTime = leadTimeDemandBase;

            // Demand ká»³ bÃ¡n SOQ (Chá»‰ tÃ­nh Coverage)
            let coverageStartDate = invDate + (leadTimeArrival * 24 * 60 * 60 * 1000);
            let coverageDemandBase = calculatePeriodDemand(coverageStartDate, coverageLT, weekdayAds, weekendAds);
            let totalDemand = coverageDemandBase;

            // --- NEW: TÄƒng trÆ°á»Ÿng theo Leadtime (Äá»‘i chiáº¿u Weekly vs Monthly trÃªn tá»«ng Thá»©) ---
            let leadtimeGrowth = 0;
            let growthHtml = '-';

            let periodAdsMonthly = basePeriodDemand / totalLeadtime;
            let weeklyPeriodDemand = calculatePeriodDemand(invDate, totalLeadtime, wWeekdayAds, wWeekendAds);
            let periodAdsWeekly = weeklyPeriodDemand / totalLeadtime;

            if (periodAdsMonthly > 0 && totalLeadtime > 0 && periodAdsWeekly > 0) {
                leadtimeGrowth = ((periodAdsWeekly - periodAdsMonthly) / periodAdsMonthly) * 100;
                if (leadtimeGrowth > 0) growthHtml = `<span style="color: var(--success)">+${leadtimeGrowth.toFixed(1)}%</span>`;
                else if (leadtimeGrowth < 0) growthHtml = `<span style="color: var(--danger)">${leadtimeGrowth.toFixed(1)}%</span>`;
                else growthHtml = `0%`;
            } else if (basePeriodDemand > 0) {
                growthHtml = `<span style="color: var(--success)">New</span>`;
            }

            // PhÃ¢n loáº¡i Tier Ä‘á»ƒ nhá»“i thÃªm Tá»“n Kho Tá»‘i Thiá»ƒu (Safety Stock)
            let tierLevel = 0;
            if (storeTierMap.has(data.storeID)) {
                let t = storeTierMap.get(data.storeID);
                if (t.includes('1') || t === 'T1' || t === 'TIER1' || t === 'TIER 1') {
                    tierLevel = 1;
                } else if (t.includes('2') || t === 'T2' || t === 'TIER 2' || t.includes('3') || t === 'T3' || t === 'TIER 3') {
                    tierLevel = 2; // Gá»™p Tier 2 vÃ  3 xÃ i chung rate
                }
            }

            let safetyStock = 0;
            if (forecastDay > 0) {
                if (tierLevel === 1) {
                    safetyStock = isWeekendDelivery ? (weekendAds * coverageLT * 0.30) : (weekdayAds * coverageLT * 0.15);
                } else if (tierLevel === 2) {
                    safetyStock = isWeekendDelivery ? (weekendAds * coverageLT * 0.20) : (weekdayAds * coverageLT * 0.10);
                }
                totalDemand += safetyStock;
            }

            // Sá»­ dá»¥ng má»‘c T Ä‘á»ƒ chuáº©n hÃ³a Tá»“n / Nháº­p Ä‘á»“ng bá»™ (Khá»Ÿi táº¡o á»Ÿ Ä‘áº§u vÃ²ng láº·p)
            let penaltyApplied = 0;
            let finalInv = invData.currentInv || 0;
            let finalDisp = invData.currentDisp || 0;
            let finalInput = inputData.currentInput || 0;

            let prevInv = invData.prevInv || 0;
            let prevInput = inputData.prevInput || 0;

            let formatDateStr = (ts) => {
                if (!ts) return 'N/A';
                let d = new Date(ts);
                return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}`;
            };

            let expectedInvAtArrival = Math.max(0, finalInv + finalInput - demandLeadTime);

            let strT = formatDateStr(T);
            let strPrevInv = formatDateStr(invData.prevInvDate);
            let strPrevInput = formatDateStr(inputData.prevInputDate);

            let invTooltip = `Tá»“n kho lÃºc T (${strT}): [ ${finalInv.toFixed(2)} ]\n- Trá»« nhu cáº§u bÃ¡n chá» hÃ ng (${leadTimeArrival.toFixed(1)} ngÃ y): -${demandLeadTime.toFixed(2)}\n=> Tá»“n dá»± kiáº¿n khi SOQ Ä‘áº¿n: ${expectedInvAtArrival.toFixed(2)}`;
            let inputTooltip = `Nháº­p/Giao hÃ ng lÃºc T (${strT}): [ ${finalInput.toFixed(2)} ]`;
            let disposalTooltip = `KHÃ”NG PHáº T Há»¦Y (Ratio quÃ¡ tháº¥p hoáº·c khÃ´ng Ä‘á»§ gá»‘c chia)`;

            let baseForDisposal = prevInv + prevInput;
            let disposalRatio = 0;

            if (baseForDisposal > 0) {
                disposalRatio = finalDisp / baseForDisposal; // Há»§y(T) / (Tá»“n(<T) + Nháº­p(<T))
            } else {
                disposalRatio = 0; // Bá»Ž QUA GIáº¢M TRá»ª Náº¾U KHÃ”NG TÃŒM THáº¤Y Lá»ŠCH Sá»¬ Dá»® LIá»†U
            }

            if (finalDisp > 0) {
                disposalTooltip = `CÃ´ng thá»©c: Há»§y(T) / (Tá»“n(<T) + Nháº­p(<T))\n`;
                disposalTooltip += `= ${finalDisp.toFixed(2)} / (${prevInv.toFixed(2)} + ${prevInput.toFixed(2)})\n`;
                if (baseForDisposal > 0) {
                    disposalTooltip += `= ${(disposalRatio * 100).toFixed(1)}%\n`;
                } else {
                    disposalTooltip += `=> Bá» qua pháº¡t giáº£m trá»« do thiáº¿u dá»¯ liá»‡u quÃ¡ khá»©\n`;
                }
                disposalTooltip += `(Ghi chÃº: Láº¥y Tá»“n cÅ©: ${strPrevInv}, Nháº­p cÅ©: ${strPrevInput})`;
            }

            if (finalDisp > 0) {
                let category = productCategoryMap.get(data.prodStd.toLowerCase()) || '';
                let isRTE_or_Leaf = category.includes('RTE') || category.includes('RAU LÃ');
                let isRoot = category.includes('Cá»¦');

                if (!category && RTE_PRODUCTS.some(p => data.bestName.toLowerCase().includes(p.toLowerCase()))) {
                    isRTE_or_Leaf = true;
                }

                let threshold = isRTE_or_Leaf ? 0.30 : (isRoot ? 0.15 : 0.15);

                if (disposalRatio > threshold) {
                    penaltyApplied = finalDisp * 0.5; // Giáº£m trá»« 50%
                    totalDemand -= penaltyApplied;
                    disposalTooltip += `\n\n--> KÃCH HOáº T PHáº T DO QUÃ NGÆ¯á» NG (${(threshold * 100).toFixed(0)}%)`;
                }
            }

            let soq = totalDemand - expectedInvAtArrival;
            soq = Math.max(Math.ceil(soq), 0);

            let itemKey = `${data.storeID}_${data.prodStd.toLowerCase()}`;
            let trendAction = trendReportMap.get(itemKey) || '';

            let xuHuongHtml = '<span>-</span>';
            if (trendAction) {
                let lowerAction = trendAction.toLowerCase();
                if (lowerAction.includes('tá»‘t') || lowerAction.includes('tot')) {
                    xuHuongHtml = `<span style="background: rgba(16, 185, 129, 0.15); color: #10b981; border: 1px solid rgba(16, 185, 129, 0.3); padding: 4px 8px; border-radius: 4px; font-weight: 600; font-size: 0.85em;">${trendAction}</span>`;
                } else if (lowerAction.includes('ngá»«ng') || lowerAction.includes('ngÆ°ng') || lowerAction.includes('ngung')) {
                    xuHuongHtml = `<span style="background: rgba(239, 68, 68, 0.15); color: #ef4444; border: 1px solid rgba(239, 68, 68, 0.3); padding: 4px 8px; border-radius: 4px; font-weight: 600; font-size: 0.85em;">${trendAction}</span>`;
                } else {
                    xuHuongHtml = `<span style="background: rgba(255, 255, 255, 0.05); color: var(--text-main); border: 1px solid var(--border); padding: 4px 8px; border-radius: 4px; font-size: 0.85em;">${trendAction}</span>`;
                }
            }

            // HIá»‚N THá»Š Äáº¦Y Äá»¦ SOQ Náº¾U CÃ“ Báº¤T Ká»² Ã NGHÄ¨A KINH DOANH NÃ€O
            // áº¨n dÃ²ng cÃ³ Táº¤T Cáº¢ = 0 (ÄÃ£ comment láº¡i theo yÃªu cáº§u Ä‘á»ƒ show Ä‘á»§ 46 mÃ£)
            // if (soq === 0 && totalDemand === 0 && finalInv === 0 && finalInput === 0 && finalDisp === 0) {
            //     return;
            // }

            let storeNameStr = storeNamesMap.get(data.storeID) || data.storeOrig;

            let totalDemandRaw = totalDemand + penaltyApplied;
            let breakdownTip = `CÃ´ng thá»©c: Demand (Nhu cáº§u gá»‘c) + SafetyStock. \n- Nhu cáº§u gá»‘c (Coverage): ${coverageDemandBase.toFixed(2)}\n- SafetyStock: +${safetyStock.toFixed(2)} \n- Penalty (Giáº£m trá»«): -${penaltyApplied.toFixed(2)}`;

            finalResults.push({
                'sap': data.storeID,
                'store': storeNameStr,
                'region': storeRegionMap.get(data.storeID) || 'KhÃ¡c',
                'product': data.bestName,
                'ads': forecastDay.toFixed(2),
                'trend': trendExport,
                'trendHtml': trendHtml,
                'ads_weekday': weekdayAds.toFixed(2),
                'ads_weekend': weekendAds.toFixed(2),
                'growth': mAds > 0 ? `${leadtimeGrowth.toFixed(1)}%` : (basePeriodDemand > 0 ? 'New' : '0%'),
                'growthHtml': growthHtml,
                'leadtime': coverageLT,
                'demand': (totalDemand + penaltyApplied).toFixed(2),
                'demandRaw': totalDemandRaw.toFixed(2),
                'inventory': Number(finalInv.toFixed(2)),
                'input': Number(finalInput.toFixed(2)),
                'penalty': penaltyApplied > 0 ? `-${penaltyApplied.toFixed(2)}` : '0',
                'soq': soq,
                'xu_huong': trendAction,
                'xu_huong_html': xuHuongHtml,
                // Tooltips
                'tip_ads': (mTotal === 0 && wTotal > 0) 
                           ? `[MÃƒ Má»šI Tá»ª FILE TUáº¦N] Sáº£n lÆ°á»£ng: ${wTotal.toFixed(1)} / ${Math.round(wDaysCount)} ngÃ y (VÃ²ng Ä‘á»i)\n=> Trung bÃ¬nh: ${forecastDay.toFixed(2)} SP/ngÃ y` 
                           : `Sáº£n lÆ°á»£ng gá»‘c: ${mTotal.toFixed(1)} / ${Math.round(mDaysCount)} ngÃ y (VÃ²ng Ä‘á»i)\n=> Trung bÃ¬nh: ${forecastDay.toFixed(2)} SP/ngÃ y`,
                'tip_trend': `BÃ¡n tuáº§n vá»«a qua: ${wAds.toFixed(2)}/ngÃ y\nBÃ¡n trung bÃ¬nh thÃ¡ng: ${mAds.toFixed(2)}/ngÃ y\n(Tá»· lá»‡ chÃªnh lá»‡ch: ${trendExport})`,
                'tip_growth': `Dá»± bÃ¡o ráº£i thá»±c táº¿ ngÃ y giao (Khá»›p T2-CN): ${forecastDay.toFixed(2)}/ngÃ y\n(Tá»· lá»‡ tÄƒng trÆ°á»Ÿng so vá»›i Trung bÃ¬nh ThÃ¡ng gá»‘c: ${mAds > 0 ? leadtimeGrowth.toFixed(1) : 0}%)`,
                'tip_weekday': `TÃ­nh tá»« gá»‘c ThÃ¡ng (Lifecycle): ${weekdayQty.toFixed(2)} / ${Math.round(weekdayDaysCount)} ngÃ y T2-T6`,
                'tip_weekend': `TÃ­nh tá»« gá»‘c ThÃ¡ng (Lifecycle): ${weekendQty.toFixed(2)} / ${Math.round(weekendDaysCount)} ngÃ y T7-CN`,
                'tip_leadtime': `Coverage: ${coverageLT} ngÃ y. (Chá»‰ tÃ­nh lÆ°á»£ng bÃ¡n ra trong ${coverageLT} ngÃ y giao hÃ ng, khÃ´ng tÃ­nh pháº§n thiáº¿u há»¥t trong ${leadTimeArrival.toFixed(1)} ngÃ y chá»)`,
                'tip_demand': breakdownTip,
                'tip_inventory': invTooltip,
                'tip_input': inputTooltip,
                'tip_penalty': disposalTooltip
            });
        });

        // Sáº¯p xáº¿p máº·c Ä‘á»‹nh: Theo MÃ£ SAP, sau Ä‘Ã³ theo TÃªn sáº£n pháº©m tÃ¹y chá»‰nh
        finalResults.sort((a, b) => {
            let sapCompare = String(a.sap).localeCompare(String(b.sap), undefined, { numeric: true });
            if (sapCompare !== 0) return sapCompare;
            
            let idxA = CUSTOM_PRODUCT_ORDER.findIndex(p => p.toLowerCase() === String(a.product).trim().toLowerCase());
            let idxB = CUSTOM_PRODUCT_ORDER.findIndex(p => p.toLowerCase() === String(b.product).trim().toLowerCase());
            
            idxA = idxA !== -1 ? idxA : 9999;
            idxB = idxB !== -1 ? idxB : 9999;
            
            if (idxA !== idxB) return idxA - idxB;
            
            return String(a.product).localeCompare(String(b.product), 'vi');
        });

        renderSOQTable(finalResults);
        populateRegionDropdown();

        if (finalResults.length === 0) {
            let monthlyKeys = (datasets.monthly && datasets.monthly.length > 0) ? Object.keys(datasets.monthly[0]).join(', ') : 'No data';
            let invKeys = (datasets.inventory && datasets.inventory.length > 0) ? Object.keys(datasets.inventory[0]).join(', ') : 'No data';
            let schedKeys = (datasets.schedule && datasets.schedule.length > 0) ? Object.keys(datasets.schedule[0]).join(', ') : 'No data';

            tbody.innerHTML = `<tr><td colspan="15" style="text-align:left; color: var(--danger); padding: 2rem;">
            <strong>KhÃ´ng tÃ¬m tháº¥y báº¥t ká»³ dá»¯ liá»‡u há»£p lá»‡ nÃ o. (Tá»“n kho, hÃ ng nháº­p vÃ  lá»‹ch giao khÃ´ng khá»›p ngÃ m dá»¯ liá»‡u, hoáº·c táº¥t cáº£ Ä‘á»u báº±ng 0).</strong><br/><br/>
            <div style="font-family: monospace; font-size:12px; color: var(--text-muted);">
                <strong>--- TRÃŒNH KIá»‚M TRA Lá»–I Ná»˜I Bá»˜ ---</strong><br/>
                - Schedule Headers: ${schedKeys}<br/>
                - Monthly Headers: ${monthlyKeys}<br/>
                - Inventory Headers: ${invKeys}<br/>
                - Lá»‹ch Giao HÃ ng quÃ©t Ä‘Æ°á»£c: ${validSAPs.size} mÃ£ há»£p lá»‡<br/>
                - Mapping File quÃ©t Ä‘Æ°á»£c: ${mappingMap ? mappingMap.size : 0} cáº·p quy Ä‘á»•i.<br/>
                - Master List Ä‘Äƒng kÃ½ Ä‘Æ°á»£c: ${allItems.size} mÃ£ sáº£n pháº©m.
            </div>
            <p>Vui lÃ²ng chá»¥p mÃ n hÃ¬nh Ä‘oáº¡n mÃ£ mÃ u xÃ¡m nÃ y vÃ  gá»­i láº¡i Ä‘á»ƒ ká»¹ sÆ° hoÃ n táº¥t cÄƒn chá»‰nh file.</p>
        </td></tr>`;
        }

        resultsSection.style.display = 'block';
        if (finalResults.length > 0) {
            btnExport.style.display = 'inline-block';
            
             // --- LÆ¯U Lá»ŠCH Sá»¬ TÃNH TOÃN NGAY Láº¬P Tá»¨C Äá»‚ XEM Láº I á»ž TAB "Lá»ŠCH Sá»¬ Táº¢I LÃŠN" (EXPIRES QUA ÄÃŠM) ---
             saveToDB('soq_latest_filename', scheduleFileName);
             saveToDB('soq_latest_html', tbody.innerHTML);
             saveToDB('soq_latest_array', finalResults);
             archiveTodayData();

             // --- LÆ¯U LÃŠN FIREBASE (CLOUD STORAGE) ---
             if (typeof firebase !== 'undefined') {
                 let userName = inputUserName ? inputUserName.value.trim() : "Há»‡ thá»‘ng";
                 if (!userName) userName = "áº¨n danh";

                 const now = new Date();
                 const dateStr = now.toISOString().split('T')[0]; // YYYY-MM-DD

                 const payload = {
                     results: finalResults,
                     filename: scheduleFileName,
                     timestamp: now.getTime(),
                     dateStr: dateStr,
                      deliveryDateStr: currentDeliveryDateStr,
                     userName: userName
                 };

                 firebase.database().ref('latest_soq').set(payload)
                     .then(() => console.log("ÄÃ£ cáº­p nháº­t SOQ má»›i nháº¥t lÃªn Cloud."))
                     .catch(err => console.error("Lá»—i lÆ°u Cloud:", err));
             }
        }
    } catch (err) {
        console.error("Lá»—i tÃ­nh toÃ¡n SOQ:", err);
        alert("Lá»—i tÃ­nh toÃ¡n: " + err.message + "\n\nBáº¡n hÃ£y kiá»ƒm tra xem cÃ¡c file Ä‘Ã£ Ä‘Æ°á»£c táº£i lÃªn Ä‘áº§y Ä‘á»§ chÆ°a nhÃ©!");
        btnCalculate.disabled = false;
        btnCalculate.textContent = "Tiáº¿n hÃ nh tÃ­nh SOQ";
    }
});

// HÃ m há»— trá»£ lÆ°u thay Ä‘á»•i lÃªn Cloud (Firebase Transaction)
function saveChangesToCloud() {
    return new Promise((resolve, reject) => {
        if (typeof firebase === 'undefined') {
            reject(new Error("Firebase chÆ°a Ä‘Æ°á»£c khá»Ÿi táº¡o."));
            return;
        }

        let userName = inputUserName ? inputUserName.value.trim() : "Há»‡ thá»‘ng";
        if (!userName) userName = "áº¨n danh";
        const now = new Date();
        const dateStr = now.toISOString().split('T')[0];

        const payload = {
            results: finalResults,
            filename: scheduleFileName,
            timestamp: now.getTime(),
            dateStr: dateStr,
            deliveryDateStr: currentDeliveryDateStr,
            userName: userName
        };

        firebase.database().ref('latest_soq').transaction((currentData) => {
            try {
                // Kiá»ƒm tra cÃ¹ng ngÃ y (dateStr) Ä‘á»ƒ gá»™p thay Ä‘á»•i cá»§a má»i ngÆ°á»i dÃ¹ng
                if (currentData && currentData.dateStr === dateStr) {
                    let cloudMap = {};
                    if (currentData.results) {
                        let resultsArr = Array.isArray(currentData.results) ? currentData.results : Object.values(currentData.results);
                        resultsArr.forEach(r => {
                            if (r) cloudMap[r.sap + '_' + r.product] = r;
                        });
                    }

                    let modified = false;
                    finalResults.forEach(localItem => {
                        if (localItem.is_dirty) {
                            modified = true;
                            let key = localItem.sap + '_' + localItem.product;
                            if (cloudMap[key]) {
                                cloudMap[key].final_order = localItem.final_order;
                                cloudMap[key].note = localItem.note;
                            } else {
                                if (!currentData.results) currentData.results = [];
                                
                                let cleanItem = JSON.parse(JSON.stringify(localItem));
                                delete cleanItem.is_dirty;

                                if (Array.isArray(currentData.results)) {
                                    currentData.results.push(cleanItem);
                                } else {
                                    let maxKey = Math.max(-1, ...Object.keys(currentData.results).map(Number).filter(n => !isNaN(n)));
                                    currentData.results[maxKey + 1] = cleanItem;
                                }
                            }
                        }
                    });

                    if (!modified) {
                        currentData.lastActive = now.getTime();
                    }

                    // Äá»“ng bá»™ metadata má»›i nháº¥t
                    currentData.filename = scheduleFileName;
                    currentData.timestamp = now.getTime();
                    currentData.userName = user        if (typeof firebase !== 'undefined') {
            btnSaveChanges.innerHTML = "â³ Äang lÆ°u...";
            saveChangesToCloud().then((committed) => {
                if (committed) {
                    btnSaveChanges.innerHTML = "âœ”ï¸ ÄÃ£ lÆ°u";
                } else {
                    btnSaveChanges.innerHTML = "âœ”ï¸ ÄÃ£ lÆ°u (KhÃ´ng Ä‘á»•i)";
                }
                setTimeout(() => { btnSaveChanges.innerHTML = "ðŸ’¾ LÆ°u Thay Äá»•i"; }, 2000);
            }).catch(err => {
                alert("Lá»—i khi lÆ°u lÃªn Cloud: " + err.message);
                btnSaveChanges.innerHTML = "ðŸ’¾ LÆ°u Thay Äá»•i";
            });
        } else {
            alert("Lá»—i: Firebase chÆ°a Ä‘Æ°á»£c khá»Ÿi táº¡o.");
        }
    });
}¶T" nhÆ°ng chÆ°a nháº­p Ä‘á»§ toÃ n bá»™ cÃ¡c mÃ£ sáº£n pháº©m:\n- ${missingStores.join('\n- ')}\n\nBáº¡n cÃ³ cháº¯c cháº¯n muá»‘n lÆ°u láº¡i khÃ´ng?`);
            if (!confirmSave) return;
        }

        if (typeof firebase !== 'undefined') {
            btnSaveChanges.innerHTML = "â³ Äang lÆ°u...";

                        if (!modified) {
                            // Cá»‘ tÃ¬nh sá»­a 1 trÆ°á»ng nhá» Ä‘á»ƒ Firebase báº¯t buá»™c nháº­n diá»‡n cÃ³ thay Ä‘á»•i (force commit)
                            currentData.lastActive = now.getTime();
                        }

                        currentData.timestamp = now.getTime();
                        currentData.userName = userName; 
                        
                        // Sanitize before returning to prevent Firebase SDK crash due to undefined properties
                        return JSON.parse(JSON.stringify(currentData));
                    }
                    
                    let newPayload = JSON.parse(JSON.stringify(payload));
                    if (Array.isArray(newPayload.results)) {
                        newPayload.results.forEach(r => { if(r) delete r.is_dirty; });
                    }
                    return newPayload;
                } catch (e) {
                    console.error("Lá»—i bÃªn trong transaction: ", e);
                    return; // Há»§y transaction
                }
            }).then((result) => {
                if (result.committed) {
                    let snapshotVal = result.snapshot.val();
                    if (snapshotVal && snapshotVal.results) {
                        let cloudRes = snapshotVal.results;
                        if (Array.isArray(cloudRes)) {
                            finalResults = cloudRes;
                        } else {
                            finalResults = Object.values(cloudRes);
                        }
                    }
                    if (Array.isArray(finalResults)) {
                        finalResults.forEach(r => { if(r) delete r.is_dirty; });
                    }
                    
                    btnSaveChanges.innerHTML = "âœ”ï¸ ÄÃ£ lÆ°u";
                    setTimeout(() => { btnSaveChanges.innerHTML = "ðŸ’¾ LÆ°u Thay Äá»•i"; }, 2000);
                    saveToDB('soq_latest_array', finalResults);
              archiveTodayData();
                    
                    renderSOQTable(finalResults);
                    populateRegionDropdown();
                } else {
                    if (Array.isArray(finalResults)) {
                        finalResults.forEach(r => { if(r) delete r.is_dirty; });
                    }
                    btnSaveChanges.innerHTML = "âœ”ï¸ ÄÃ£ lÆ°u (KhÃ´ng Ä‘á»•i)";
                    setTimeout(() => { btnSaveChanges.innerHTML = "ðŸ’¾ LÆ°u Thay Äá»•i"; }, 2000);
                }
            }).catch(err => {
                console.error("Lá»—i lÆ°u Cloud:", err);
                alert("Lá»—i khi lÆ°u lÃªn Cloud: " + err.message);
                btnSaveChanges.innerHTML = "ðŸ’¾ LÆ°u Thay Äá»•i";
            });
        } else {
            alert("Lá»—i: Firebase chÆ°a Ä‘Æ°á»£c khá»Ÿi táº¡o.");
        }
    });
}

// Export to Excel (Bypass Security Block for local file:///)
btnExport.addEventListener('click', () => {
    if (!isArchiveView) {
        archiveTodayData();
    }
    if (!datasets.template_headers || datasets.template_headers.length === 0) {
        alert("Vui lÃ²ng táº£i lÃªn 'Form Xuáº¥t Máº«u' á»Ÿ má»¥c 6 (Cáº¥u hÃ¬nh) trÆ°á»›c khi xuáº¥t Excel Ä‘á»ƒ Ä‘áº£m báº£o Ä‘Ãºng Ä‘á»‹nh dáº¡ng!");
        return;
    }

    const parseNum = (val) => {
        if (val === undefined || val === null || val === '') return '';
        if (typeof val === 'number') return val;
        let str = String(val).trim();
        if (str.includes('%')) return str;
        let num = Number(str);
        if (!isNaN(num)) return num;
        return str;
    };

    let stores = new Map();
    finalResults.forEach(item => {
        if (!stores.has(item.sap)) {
            stores.set(item.sap, {
                region: item.region || 'KhÃ¡c',
                buyerName: item.store || '',
                sap: item.sap || '',
                notes: new Set(),
                products: {}
            });
        }
        let s = stores.get(item.sap);
        let qty = (isHistoryView || isArchiveView) ? (item.final_order !== undefined ? item.final_order : '') : item.soq;
        
        if (qty !== '' && qty > 0) {
             let pKey = String(item.product).trim().toLowerCase();
             s.products[pKey] = parseNum(qty);
        }
        if (item.note && String(item.note).trim() !== '') {
             s.notes.add(String(item.note).trim());
        }
    });
    
    let storeArray = Array.from(stores.values());
    storeArray.sort((a, b) => a.region.localeCompare(b.region, 'vi'));
    
    let aoa = [];
    // DÃ²ng 1: TiÃªu Ä‘á» cá»™t giá»¯ nguyÃªn y há»‡t Form xuáº¥t máº«u
    aoa.push(datasets.template_headers);
    
    // CÃ¡c dÃ²ng dá»¯ liá»‡u
    storeArray.forEach(s => {
        let rowData = [];
        datasets.template_headers.forEach(header => {
            let hUpper = header.trim().toUpperCase();
            if (hUpper === 'KHU Vá»°C' || hUpper === 'KHU VUC' || hUpper === 'REGION') {
                rowData.push(s.region);
            } else if (hUpper === 'BUYER NAME' || hUpper.includes('BUYER')) {
                rowData.push(s.buyerName);
            } else if (hUpper === 'ORDER NOTE' || hUpper === 'GHI CHÃš' || hUpper.includes('NOTE')) {
                rowData.push(Array.from(s.notes).join(', '));
            } else if (hUpper === 'SAP' || hUpper === 'MÃƒ SAP') {
                rowData.push(s.sap);
            } else {
                let hKey = header.trim().toLowerCase();
                rowData.push(s.products[hKey] !== undefined ? s.products[hKey] : '');
            }
        });
        aoa.push(rowData);
    });

    const worksheet = XLSX.utils.aoa_to_sheet(aoa);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "SOQ_Results");

    // --- Táº¡o Sheet thá»© 2: Raw Data (Káº¿t Quáº£ Dá»± BÃ¡o) ---
    let rawAoa = [];
    let rawHeaders = [
        "MÃ£ SAP (Store)", "TÃªn Cá»­a HÃ ng", "Khu Vá»±c", "TÃªn Sáº£n Pháº©m", 
        "Trung BÃ¬nh BÃ¡n/NgÃ y", "Xu HÆ°á»›ng (%)", "ADS T2-T6", "ADS T7-CN", 
        "XU HÆ¯á»šNG GIAO (%)", "Leadtime", "Total Demand", "Tá»“n (Inv)", 
        "Nháº­p (Input)", "Giáº£m trá»«", "SOQ (Gá»¢I Ã)", "Xu hÆ°á»›ng", "SL Äáº¶T", "GHI CHÃš"
    ];
    rawAoa.push(rawHeaders);
    
    finalResults.forEach(item => {
        let trendText = String(item.trendHtml || '').replace(/<[^>]*>?/gm, '').trim();
        let growthText = String(item.growthHtml || '').replace(/<[^>]*>?/gm, '').trim();
        
        let slDat = (isHistoryView || isArchiveView) ? (item.final_order !== undefined ? item.final_order : '') : item.soq;
        let ghiChu = item.note || '';

        rawAoa.push([
            item.sap,
            item.store,
            item.region,
            item.product,
            item.ads,
            trendText,
            item.ads_weekday,
            item.ads_weekend,
            growthText,
            item.leadtime,
            item.demandRaw,
            item.inventory,
            item.input,
            item.penalty,
            item.soq,
            item.xu_huong || '',
            slDat,
            ghiChu
        ]);
    });
    const rawWorksheet = XLSX.utils.aoa_to_sheet(rawAoa);
    XLSX.utils.book_append_sheet(workbook, rawWorksheet, "Data_Chi_Tiet");
    // ------------------------------------------------

    // Khá»­ dáº¥u tiáº¿ng Viá»‡t vÃ  kÃ½ tá»± láº¡ Ä‘á»ƒ trÃ¡nh Browser cháº·n táº£i
    let safeName = String(scheduleFileName).normalize('NFD').replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9_\-]/g, "_");
    let exportName = `SOQ_Data_${safeName}.xlsx`;

    // Custom File downloader to bypass 'Cáº§n cÃ³ quyá»n táº£i xuá»‘ng' warning
    try {
        let wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });
        function s2ab(s) {
            let buf = new ArrayBuffer(s.length);
            let view = new Uint8Array(buf);
            for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
            return buf;
        }
        let blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
        let url = URL.createObjectURL(blob);
        let a = document.createElement("a");
        document.body.appendChild(a);
        a.href = url;
        a.download = exportName;
        a.click();
        setTimeout(() => {
            URL.revokeObjectURL(url);
            document.body.removeChild(a);
        }, 100);

    } catch (e) {
        alert("Lá»—i táº£i file: TrÃ¬nh duyá»‡t cá»§a báº¡n khÃ³a quyá»n táº£i cá»¥c bá»™. HÃ£y má»Ÿ trang nÃ y báº±ng Chrome nhÃ©!");
    }
});

// --- Bá»˜ Lá»ŒC TÃŒM KIáº¾M ---
const searchStoreInput = document.getElementById('search-store');
const searchProductInput = document.getElementById('search-product');
const filterRegionSelect = document.getElementById('filter-region');
const filterTrendActionSelect = document.getElementById('filter-trend-action');

function populateRegionDropdown() {
    if (!filterRegionSelect) return;
    while (filterRegionSelect.options.length > 1) {
        filterRegionSelect.remove(1);
    }
    const regions = new Set();
    finalResults.forEach(item => {
        if (item.region && item.region !== 'KhÃ¡c') regions.add(item.region);
    });
    let sortedRegions = Array.from(regions).sort();
    let hasKhac = finalResults.some(item => !item.region || item.region === 'KhÃ¡c');
    if (hasKhac) sortedRegions.push('KhÃ¡c');
    
    sortedRegions.forEach(r => {
        let opt = document.createElement('option');
        opt.value = r;
        opt.text = r;
        filterRegionSelect.appendChild(opt);
    });
}

function filterTable() {
    if (!searchStoreInput || !searchProductInput) return;
    const storeQuery = searchStoreInput.value.toLowerCase();
    const productQuery = searchProductInput.value.toLowerCase();
    const regionQuery = filterRegionSelect ? filterRegionSelect.value : "";
    const trendActionQuery = filterTrendActionSelect ? filterTrendActionSelect.value : "";
    const rows = document.querySelectorAll('#soq-tbody tr');

    rows.forEach(row => {
        if (row.cells.length < 3) return; // Skip special rows like empty data messages
        const sap = row.cells[0].textContent.toLowerCase();
        const storeName = row.cells[1].textContent.toLowerCase();
        const productName = row.cells[2].textContent.toLowerCase();
        const region = row.getAttribute('data-region') || 'KhÃ¡c';
        const xuHuong = row.getAttribute('data-xu-huong') || '';
        const lowerXuHuong = xuHuong.toLowerCase();

        const matchStore = sap.includes(storeQuery) || storeName.includes(storeQuery);
        const matchProduct = productName.includes(productQuery);
        const matchRegion = regionQuery === "" || region === regionQuery;
        
        let matchTrendAction = false;
        if (trendActionQuery === "") {
            matchTrendAction = true;
        } else if (trendActionQuery === "bantot") {
            matchTrendAction = lowerXuHuong.includes('tá»‘t') || lowerXuHuong.includes('tot');
        } else if (trendActionQuery === "ngunggiao") {
            matchTrendAction = lowerXuHuong.includes('ngá»«ng') || lowerXuHuong.includes('ngÆ°ng') || lowerXuHuong.includes('ngung');
        } else if (trendActionQuery === "trong") {
            matchTrendAction = xuHuong.trim() === "" || xuHuong === "-";
        }

        if (matchStore && matchProduct && matchRegion && matchTrendAction) {
            row.style.display = '';
        } else {
            row.style.display = 'none';
        }
    });
}

if (filterRegionSelect) {
    filterRegionSelect.addEventListener('change', filterTable);
}
if (filterTrendActionSelect) {
    filterTrendActionSelect.addEventListener('change', filterTable);
}

if (searchStoreInput && searchProductInput) {
    searchStoreInput.addEventListener('input', () => {
        // Khi gÃµ tÃ¬m kiáº¿m cá»­a hÃ ng, tá»± Ä‘á»™ng reset báº£ng vá» thá»© tá»± chuáº©n cá»§a riÃªng cá»­a hÃ ng Ä‘Ã³
        if (currentSort && currentSort.direction !== 0) {
            currentSort.column = null;
            currentSort.direction = 0;
            document.querySelectorAll('.sort-icon').forEach(icon => icon.textContent = '');
            
            if (typeof finalResults !== 'undefined' && finalResults) {
                finalResults.sort((a, b) => {
                    let sapCompare = String(a.sap).localeCompare(String(b.sap), undefined, { numeric: true });
                    if (sapCompare !== 0) return sapCompare;
                    
                    let idxA = CUSTOM_PRODUCT_ORDER.findIndex(p => p.toLowerCase() === String(a.product).trim().toLowerCase());
                    let idxB = CUSTOM_PRODUCT_ORDER.findIndex(p => p.toLowerCase() === String(b.product).trim().toLowerCase());
                    
                    idxA = idxA !== -1 ? idxA : 9999;
                    idxB = idxB !== -1 ? idxB : 9999;
                    
                    if (idxA !== idxB) return idxA - idxB;
                    
                    return String(a.product).localeCompare(String(b.product), 'vi');
                });
                renderSOQTable(finalResults);
        populateRegionDropdown();
            }
        }
        filterTable();
    });
    
    searchProductInput.addEventListener('input', filterTable);
}



// --- COLLAPSE SIDEBAR ---
const btnToggleSidebar = document.getElementById('btn-toggle-sidebar');
const btnToggleSidebarWeekly = document.getElementById('btn-toggle-sidebar-weekly');
const sidebar = document.querySelector('.sidebar');

function toggleSidebar() {
    const buttons = [
        document.getElementById('btn-toggle-sidebar'),
        document.getElementById('btn-toggle-sidebar-weekly')
    ];
    if (sidebar) {
        if (sidebar.style.display === 'none') {
            sidebar.style.display = 'flex';
            buttons.forEach(btn => {
                if (btn) btn.innerHTML = '<span>â—„</span> áº¨n Menu trÃ¡i';
            });
        } else {
            sidebar.style.display = 'none';
            buttons.forEach(btn => {
                if (btn) btn.innerHTML = '<span>â–º</span> Hiá»‡n Menu trÃ¡i';
            });
        }
    }
}

if (btnToggleSidebar) {
    btnToggleSidebar.addEventListener('click', toggleSidebar);
}
if (btnToggleSidebarWeekly) {
    btnToggleSidebarWeekly.addEventListener('click', toggleSidebar);
}
// --- ÄIá»€U CHUYá»‚N MENU TAB Lá»ŠCH Sá»¬ VÃ€ Báº¢NG TÃNH ---
const navDashboard = document.getElementById('nav-dashboard');
const navHistory = document.getElementById('nav-history');

if (navHistory && navDashboard) {
    navHistory.addEventListener('click', async (e) => {
        e.preventDefault();
        document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
        navHistory.classList.add('active');
        isHistoryView = true;
        isArchiveView = false;
        
        // áº¨n khu vá»±c táº£i file vÃ  báº£ng chá»n archive
        document.querySelector('.upload-section').style.display = 'none';
        document.getElementById('archive-selector-container').style.display = 'none';
        document.getElementById('weekly-review-container').style.display = 'none';
        document.querySelectorAll('.history-col').forEach(c => c.style.display = 'table-cell');
        let btnSave = document.getElementById('btn-save-changes');
        if (btnSave) btnSave.style.display = 'inline-block';
        
        let tbody = document.getElementById('soq-tbody');
        let titleSpan = document.querySelector('.results-section h2');
        let btnExport = document.getElementById('btn-export');

        // Hiá»‡n section káº¿t quáº£ trÆ°á»›c Ä‘á»ƒ ngÆ°á»i dÃ¹ng tháº¥y Ä‘ang load
        document.getElementById('results-section').style.display = 'block';
        tbody.innerHTML = `<tr><td colspan="17" style="text-align:center; padding: 2rem;">ðŸ”„ Äang táº£i lá»‹ch sá»­ tá»« Cloud...</td></tr>`;

        // 1. Kiá»ƒm tra Firebase trÆ°á»›c (Shared History)
        if (typeof firebase !== 'undefined') {
            firebase.database().ref('latest_soq').once('value').then(async (snapshot) => {
                const data = snapshot.val();
                const todayStr = new Date().toISOString().split('T')[0];

                if (data && data.dateStr === todayStr) {
                    if (data.deliveryDateStr) {
                        currentDeliveryDateStr = data.deliveryDateStr;
                        saveToDB('soq_latest_delivery_date', currentDeliveryDateStr);
                    }
                    // Dá»¯ liá»‡u há»£p lá»‡ (trong ngÃ y)
                    // Render báº£ng tá»« Array
                    finalResults = prepHistoricalData(data.results);
                    renderSOQTable(finalResults);
                    populateRegionDropdown();
                    btnExport.style.display = 'inline-block';

                    let timeStr = new Date(data.timestamp).toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' });
                    titleSpan.innerHTML = `Káº¿t Quáº£ Dá»± BÃ¡o <span style="font-size: 0.6em; background: rgba(76, 175, 80, 0.2); color: #4caf50; border: 1px solid #4caf50; padding: 4px 8px; border-radius: 4px; margin-left: 10px; vertical-align: middle;">Shared: ${data.userName} (${timeStr})</span>`;
                } else {
                    // KhÃ´ng cÃ³ dá»¯ liá»‡u Cloud hÃ´m nay -> Fallback vá» Local Cache cá»§a chÃ­nh mÃ¬nh
                    loadLocalHistoryFallback(tbody, titleSpan, btnExport);
                }
            }).catch(err => {
                console.error("Lá»—i táº£i Cloud:", err);
                loadLocalHistoryFallback(tbody, titleSpan, btnExport);
            });
        } else {
            loadLocalHistoryFallback(tbody, titleSpan, btnExport);
        }
    });

    function prepHistoricalData(arr) {
        return arr.map(item => {
            // 1. PhÃ¢n tÃ­ch Xu hÆ°á»›ng (Trend)
            let trendVal = String(item.trend || '-').trim();
            let trendNum = parseFloat(trendVal.replace(/[â–²â–¼+%\s]/g, ''));
            let trendHtml = `<span>${trendVal}</span>`;
            
            if (trendVal.toLowerCase().includes('new') || trendVal.toLowerCase().includes('má»›i')) {
                trendHtml = `<span style="color: var(--success)">â–² Má»›i bÃ¡n</span>`;
            } else if (!isNaN(trendNum)) {
                if (Math.abs(trendNum) < 1e-6) {
                    trendHtml = `<span>0.0%</span>`;
                } else if (trendNum > 0 || trendVal.includes('+') || trendVal.includes('â–²')) {
                    trendHtml = `<span style="color: var(--success)">â–² ${Math.abs(trendNum).toFixed(1)}%</span>`;
                } else if (trendNum < 0 || trendVal.includes('-') || trendVal.includes('â–¼')) {
                    trendHtml = `<span style="color: var(--danger)">â–¼ ${Math.abs(trendNum).toFixed(1)}%</span>`;
                }
            }

            // 2. PhÃ¢n tÃ­ch TÄƒng trÆ°á»Ÿng (Growth)
            let growthVal = String(item.growth || '-').trim();
            let growthNum = parseFloat(growthVal.replace(/[â–²â–¼+%\s]/g, ''));
            let growthHtml = `<span>${growthVal}</span>`;
            
            if (growthVal.toLowerCase().includes('new') || growthVal.toLowerCase().includes('má»›i')) {
                growthHtml = `<span style="color: var(--success)">${growthVal}</span>`;
            } else if (!isNaN(growthNum)) {
                if (growthNum > 1e-6) {
                    growthHtml = `<span style="color: var(--success)">+${growthNum.toFixed(1)}%</span>`;
                } else if (growthNum < -1e-6) {
                    growthHtml = `<span style="color: var(--danger)">-${Math.abs(growthNum).toFixed(1)}%</span>`;
                } else {
                    growthHtml = `<span>0.0%</span>`;
                }
            }

            item.trendHtml = trendHtml;
            item.growthHtml = growthHtml;
            item.demandRaw = item.demand || '0.00';
            
            // Clean undefineds to avoid "undefined" strings
            item.sap = item.sap || '';
            item.store = item.store || '';
            item.region = item.region || 'KhÃ¡c';
            item.product = item.product || '';
            item.ads = item.ads || '0.00';
            item.ads_weekday = item.ads_weekday || '0.00';
            item.ads_weekend = item.ads_weekend || '0.00';
            item.leadtime = item.leadtime || '';
            item.inventory = item.inventory || 0;
            item.input = item.input || 0;
            item.penalty = item.penalty || '0';
            item.soq = item.soq || 0;
            item.xu_huong = item.xu_huong || '';
            
            let xuHuongHtml = '<span>-</span>';
            if (item.xu_huong) {
                let lowerAction = item.xu_huong.toLowerCase();
                if (lowerAction.includes('tá»‘t') || lowerAction.includes('tot')) {
                    xuHuongHtml = `<span style="background: rgba(16, 185, 129, 0.15); color: #10b981; border: 1px solid rgba(16, 185, 129, 0.3); padding: 4px 8px; border-radius: 4px; font-weight: 600; font-size: 0.85em;">${item.xu_huong}</span>`;
                } else if (lowerAction.includes('ngá»«ng') || lowerAction.includes('ngÆ°ng') || lowerAction.includes('ngung')) {
                    xuHuongHtml = `<span style="background: rgba(239, 68, 68, 0.15); color: #ef4444; border: 1px solid rgba(239, 68, 68, 0.3); padding: 4px 8px; border-radius: 4px; font-weight: 600; font-size: 0.85em;">${item.xu_huong}</span>`;
                } else {
                    xuHuongHtml = `<span style="background: rgba(255, 255, 255, 0.05); color: var(--text-main); border: 1px solid var(--border); padding: 4px 8px; border-radius: 4px; font-size: 0.85em;">${item.xu_huong}</span>`;
                }
            }
            item.xu_huong_html = xuHuongHtml;
            
            return item;
        });
    }

    // HÃ m bá»• trá»£ Load Local
    async function loadLocalHistoryFallback(tbody, titleSpan, btnExport) {
        let histArr = await loadFromDB('soq_latest_array'); // KhÃ´ng dÃ¹ng histHtml tá»« Cache vÃ¬ cÃ³ thá»ƒ bá»‹ stale style
        let histName = await loadFromDB('soq_latest_filename');

        if (histArr && !histArr.invalidated) {
            finalResults = prepHistoricalData(histArr);
            renderSOQTable(finalResults);
        populateRegionDropdown(); // Render láº¡i tá»« máº£ng Ä‘á»ƒ Ã¡p dá»¥ng Style má»›i nháº¥t
            if (histName && !histName.invalidated) scheduleFileName = histName;
            btnExport.style.display = 'inline-block';
            titleSpan.innerHTML = `Káº¿t Quáº£ Dá»± BÃ¡o <span style="font-size: 0.6em; background: rgba(255,152,0,0.2); color: #ff9800; border: 1px solid #ff9800; padding: 4px 8px; border-radius: 4px; margin-left: 10px; vertical-align: middle;">Local: Báº£n lÆ°u mÃ¡y báº¡n</span>`;
        } else {
            btnExport.style.display = 'none';
            tbody.innerHTML = `<tr><td colspan="15" style="text-align:center; padding: 2.5rem; color: #ff9800; font-size: 1.1em;"><i class="fas fa-history" style="font-size: 2em; display: block; margin-bottom: 10px; opacity: 0.5;"></i>KhÃ´ng cÃ³ lá»‹ch sá»­ chia sáº» hoáº·c lá»‹ch sá»­ mÃ¡y báº¡n Ä‘Ã£ háº¿t háº¡n trong ngÃ y hÃ´m nay.</td></tr>`;
            titleSpan.innerHTML = `Káº¿t Quáº£ Dá»± BÃ¡o`;
        }
    }

    navDashboard.addEventListener('click', (e) => {
        e.preventDefault();
        document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
        navDashboard.classList.add('active');
        isHistoryView = false;
        isArchiveView = false;
        
        // Hiá»‡n láº¡i khu vá»±c Táº£i file
        document.querySelector('.upload-section').style.display = 'block';
        document.getElementById('archive-selector-container').style.display = 'none';
        document.getElementById('weekly-review-container').style.display = 'none';
        document.querySelectorAll('.history-col').forEach(c => c.style.display = 'none');
        let btnSave = document.getElementById('btn-save-changes');
        if (btnSave) btnSave.style.display = 'none';
        
        let titleSpan = document.querySelector('.results-section h2');
        if (titleSpan && titleSpan.querySelector('span')) { 
            // Dá»n dáº¹p View Lá»‹ch sá»­ (Ã‰p ngÆ°á»i dÃ¹ng báº¥n TÃ­nh SOQ láº¡i Ä‘á»ƒ táº£i láº¡i Live Data an toÃ n)
            titleSpan.innerHTML = `Káº¿t Quáº£ Dá»± BÃ¡o`;
            document.getElementById('soq-tbody').innerHTML = ''; 
            document.getElementById('results-section').style.display = 'none';
            finalResults = [];
        }
    });

    const navArchive = document.getElementById('nav-archive');
    const archiveDateInput = document.getElementById('archive-date-input');
    const btnLoadArchive = document.getElementById('btn-load-archive');

    function updateArchiveDayOfWeekBadge() {
        const badge = document.getElementById('archive-day-of-week-badge');
        if (badge && archiveDateInput) {
            const dateStr = archiveDateInput.value;
            if (dateStr) {
                const date = new Date(dateStr);
                if (!isNaN(date.getTime())) {
                    const days = ['Chá»§ Nháº­t', 'Thá»© Hai', 'Thá»© Ba', 'Thá»© TÆ°', 'Thá»© NÄƒm', 'Thá»© SÃ¡u', 'Thá»© Báº£y'];
                    const dayOfWeek = days[date.getDay()];
                    const parts = dateStr.split('-');
                    const formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
                    badge.textContent = `${dayOfWeek}, ${formattedDate}`;
                    badge.style.display = 'inline-block';
                    return;
                }
            }
            badge.style.display = 'none';
        }
    }

    if (navArchive) {
        navArchive.addEventListener('click', (e) => {
            e.preventDefault();
            document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
            navArchive.classList.add('active');
            isHistoryView = false;
            isArchiveView = true;
            
            // áº¨n khu vá»±c táº£i file, hiá»‡n khung chá»n lÆ°u trá»¯
            document.querySelector('.upload-section').style.display = 'none';
            document.getElementById('archive-selector-container').style.display = 'block';
            document.getElementById('weekly-review-container').style.display = 'none';
            
            // Hiá»‡n cÃ¡c cá»™t lá»‹ch sá»­ (cháº¿ Ä‘á»™ xem chá»‰ Ä‘á»c)
            document.querySelectorAll('.history-col').forEach(c => c.style.display = 'table-cell');
            
            // áº¨n nÃºt lÆ°u thay Ä‘á»•i (vÃ¬ lÃ  archive xem láº¡i chá»‰ Ä‘á»c)
            let btnSave = document.getElementById('btn-save-changes');
            if (btnSave) btnSave.style.display = 'none';
            
            // Dá»n dáº¹p báº£ng vÃ  káº¿t quáº£ cÅ©
            let tbody = document.getElementById('soq-tbody');
            tbody.innerHTML = '';
            document.getElementById('results-section').style.display = 'none';
            
            // Thiáº¿t láº­p máº·c Ä‘á»‹nh ngÃ y hÃ´m qua cho date input
            if (archiveDateInput) {
                const today = new Date();
                const yesterday = new Date(today);
                yesterday.setDate(today.getDate() - 1);
                const year = yesterday.getFullYear();
                const month = String(yesterday.getMonth() + 1).padStart(2, '0');
                const day = String(yesterday.getDate()).padStart(2, '0');
                const dateStr = `${year}-${month}-${day}`;
                archiveDateInput.value = dateStr;
                updateArchiveDayOfWeekBadge();
            }
            
            let titleSpan = document.querySelector('.results-section h2');
            if (titleSpan) titleSpan.innerHTML = `Káº¿t Quáº£ Dá»± BÃ¡o`;
        });
    }

    const loadArchiveData = async () => {
        const dateStr = archiveDateInput ? archiveDateInput.value : '';
        if (!dateStr) {
            alert("Vui lÃ²ng chá»n ngÃ y lÆ°u trá»¯!");
            return;
        }
        
        let tbody = document.getElementById('soq-tbody');
        let titleSpan = document.querySelector('.results-section h2');
        let btnExport = document.getElementById('btn-export');
        
        document.getElementById('results-section').style.display = 'block';
        tbody.innerHTML = `<tr><td colspan="17" style="text-align:center; padding: 2rem;">ðŸ”„ Äang táº£i dá»¯ liá»‡u lÆ°u trá»¯ ngÃ y ${dateStr.split('-').reverse().join('/')}...</td></tr>`;
        
        const renderArchive = (data, sourceName) => {
            finalResults = prepHistoricalData(data.results || data);
            renderSOQTable(finalResults);
            populateRegionDropdown();
            if (data.filename) scheduleFileName = data.filename;
            btnExport.style.display = 'inline-block';
            let formattedDate = dateStr.split('-').reverse().join('/');
            titleSpan.innerHTML = `Káº¿t Quáº£ Dá»± BÃ¡o <span style="font-size: 0.6em; background: rgba(33, 150, 243, 0.2); color: #2196f3; border: 1px solid #2196f3; padding: 4px 8px; border-radius: 4px; margin-left: 10px; vertical-align: middle;">LÆ°u trá»¯ ${sourceName}: ${formattedDate}</span>`;
        };

        // 1. Thá»­ táº£i tá»« Firebase trÆ°á»›c
        if (typeof firebase !== 'undefined') {
            try {
                let snapshot = await firebase.database().ref('archive_soq/' + dateStr).once('value');
                let data = snapshot.val();
                if (data && data.results) {
                    renderArchive(data, "Cloud");
                    return;
                }
            } catch (err) {
                console.error("Lá»—i táº£i lÆ°u trá»¯ tá»« Cloud:", err);
            }
        }
        
        // 2. Fallback táº£i tá»« IndexedDB local
        try {
            let localData = await loadFromDB('soq_archive_' + dateStr);
            if (localData && (localData.results || Array.isArray(localData))) {
                renderArchive(localData, "Local");
                return;
            }
        } catch (err) {
            console.error("Lá»—i táº£i lÆ°u trá»¯ Local:", err);
        }
        
        // 3. KhÃ´ng tÃ¬m tháº¥y dá»¯ liá»‡u
        tbody.innerHTML = `<tr><td colspan="17" style="text-align:center; padding: 2.5rem; color: #ff9800; font-size: 1.1em;"><i class="fas fa-exclamation-triangle" style="font-size: 2em; display: block; margin-bottom: 10px; opacity: 0.5;"></i>KhÃ´ng tÃ¬m tháº¥y dá»¯ liá»‡u lÆ°u trá»¯ ngÃ y ${dateStr.split('-').reverse().join('/')} trÃªn há»‡ thá»‘ng.</td></tr>`;
        btnExport.style.display = 'none';
    };

    if (btnLoadArchive) {
        btnLoadArchive.addEventListener('click', loadArchiveData);
    }
    if (archiveDateInput) {
        archiveDateInput.addEventListener('change', () => {
            updateArchiveDayOfWeekBadge();
            loadArchiveData();
        });
    }
}

// --- TABLE SORTING LOGIC ---
let currentSort = { column: null, direction: 1 };

function renderSOQTable(data) {
    const tbody = document.getElementById('soq-tbody');
    tbody.innerHTML = ``;
    data.forEach((item, index) => {
        let tr = document.createElement('tr');
        tr.setAttribute('data-region', item.region || 'KhÃ¡c');
        tr.setAttribute('data-xu-huong', item.xu_huong || '');
        
        let finalOrderTd = '';
        let noteTd = '';
        if (isHistoryView) {
            let finalVal = item.final_order !== undefined ? item.final_order : '';
            finalOrderTd = `<td><input type="number" class="final-order-input" data-index="${index}" value="${finalVal}" style="width: 80px; padding: 6px; text-align: center; border: 1px solid #ccc; border-radius: 4px; font-weight: bold; background: #fff; color: #333;" placeholder="-" min="0"></td>`;
            let noteVal = item.note !== undefined ? item.note : '';
            noteTd = `<td><input type="text" class="note-input" data-index="${index}" value="${noteVal}" style="width: 150px; padding: 6px; border: 1px solid #ccc; border-radius: 4px; background: #fff; color: #333;" placeholder="Ghi chÃº..."></td>`;
        } else if (isArchiveView) {
            let finalVal = item.final_order !== undefined ? item.final_order : '';
            finalOrderTd = `<td style="font-weight: bold; text-align: center; color: var(--primary);">${finalVal !== '' ? finalVal : '-'}</td>`;
            let noteVal = item.note !== undefined ? item.note : '';
            noteTd = `<td style="color: var(--text-muted); font-style: italic;">${noteVal !== '' ? noteVal : ''}</td>`;
        }

        tr.innerHTML = `
            <td>${item.sap}</td>
            <td>${item.store}</td>
            <td>${item.product}</td>
            <td title="${item.tip_ads}">${item.ads}</td>
            <td title="${item.tip_trend}"><b>${item.trendHtml}</b></td>
            <td title="${item.tip_weekday}">${item.ads_weekday}</td>
            <td title="${item.tip_weekend}">${item.ads_weekend}</td>
            <td title="${item.tip_growth}"><b>${item.growthHtml}</b></td>
            <td><span title="${item.tip_leadtime}">${item.leadtime}</span></td>
            <td title="${item.tip_demand}">${item.demandRaw}</td>
            <td class="warning" title="${item.tip_inventory}">${item.inventory}</td>
            <td class="highlight" title="${item.tip_input}">${item.input}</td>
            <td style="color:${item.penalty !== '0' ? 'var(--danger)' : ''}" title="${item.tip_penalty}">${item.penalty}</td>
            <td class="highlight">${item.soq}</td>
            <td>${item.xu_huong_html || '<span>-</span>'}</td>
            ${finalOrderTd}
            ${noteTd}
        `;
        tbody.appendChild(tr);
    });

    if (isHistoryView) {
        document.querySelectorAll('.final-order-input').forEach(input => {
            input.addEventListener('input', (e) => {
                let idx = e.target.getAttribute('data-index');
                if (data === finalResults) {
                    finalResults[idx].final_order = e.target.value;
                    finalResults[idx].is_dirty = true;
                } else {
                    data[idx].final_order = e.target.value;
                    data[idx].is_dirty = true;
                }
            });
            input.addEventListener('change', (e) => {
                saveToDB('soq_latest_array', finalResults);
                archiveTodayData();
            });
        });
        document.querySelectorAll('.note-input').forEach(input => {
            input.addEventListener('input', (e) => {
                let idx = e.target.getAttribute('data-index');
                if (data === finalResults) {
                    finalResults[idx].note = e.target.value;
                    finalResults[idx].is_dirty = true;
                } else {
                    data[idx].note = e.target.value;
                    data[idx].is_dirty = true;
                }
            });
            input.addEventListener('change', (e) => {
                saveToDB('soq_latest_array', finalResults);
                archiveTodayData();
            });
        });
    }
}

document.querySelectorAll('.sortable').forEach(th => {
    th.addEventListener('click', () => {
        let col = th.getAttribute('data-sort');
        
        if (col === 'product') {
            currentSort.column = null;
            currentSort.direction = 0;
        } else {
            if (currentSort.column === col) {
                if (currentSort.direction === 1) currentSort.direction = -1;
                else if (currentSort.direction === -1) currentSort.direction = 0;
                else currentSort.direction = 1;
            } else {
                currentSort.column = col;
                currentSort.direction = 1;
            }
        }

        document.querySelectorAll('.sort-icon').forEach(icon => icon.textContent = '');

        if (currentSort.direction === 0 || col === 'product') {
            finalResults.sort((a, b) => {
                let sapCompare = String(a.sap).localeCompare(String(b.sap), undefined, { numeric: true });
                if (sapCompare !== 0) return sapCompare;
                
                let idxA = CUSTOM_PRODUCT_ORDER.findIndex(p => p.toLowerCase() === String(a.product).trim().toLowerCase());
                let idxB = CUSTOM_PRODUCT_ORDER.findIndex(p => p.toLowerCase() === String(b.product).trim().toLowerCase());
                
                idxA = idxA !== -1 ? idxA : 9999;
                idxB = idxB !== -1 ? idxB : 9999;
                
                if (idxA !== idxB) return idxA - idxB;
                
                return String(a.product).localeCompare(String(b.product), 'vi');
            });
            renderSOQTable(finalResults);
            return;
        }

        finalResults.sort((a, b) => {
            let valA = a[col];
            let valB = b[col];

            if (typeof valA === 'string') valA = valA.toLowerCase();
            if (typeof valB === 'string') valB = valB.toLowerCase();

            let numA = parseFloat(String(valA).replace(/[^0-9.-]/g, ''));
            let numB = parseFloat(String(valB).replace(/[^0-9.-]/g, ''));

            if (!isNaN(numA) && !isNaN(numB) && String(valA).match(/\d/) && String(valB).match(/\d/)) {
                return (numA - numB) * currentSort.direction;
            }

            if (valA < valB) return -1 * currentSort.direction;
            if (valA > valB) return 1 * currentSort.direction;
            return 0;
        });

        renderSOQTable(finalResults);

        let targetIcon = th.querySelector('.sort-icon');
        if (targetIcon) targetIcon.textContent = currentSort.direction === 1 ? ' \u25BC' : ' \u25B2';
    });
});

// NgÄƒn lá»—i cuá»™n chuá»™t lÃ m thay Ä‘á»•i sá»‘ trong tháº» input type="number"
document.addEventListener('wheel', function(event) {
    if (document.activeElement.type === 'number') {
        document.activeElement.blur();
    }
});

// --- BÃO CÃO Tá»”NG Há»¢P TUáº¦N ---
async function fetchDailyOrderArchive(dateStr) {
    if (typeof firebase !== 'undefined') {
        try {
            let snapshot = await firebase.database().ref('archive_soq/' + dateStr).once('value');
            let data = snapshot.val();
            if (data && data.results) {
                return data.results;
            }
        } catch (e) {
            console.error(`Lá»—i táº£i archive Cloud ngÃ y ${dateStr}:`, e);
        }
    }
    try {
        let localData = await loadFromDB('soq_archive_' + dateStr);
        if (localData) {
            return localData.results || localData;
        }
    } catch (e) {
        console.error(`Lá»—i táº£i archive Local ngÃ y ${dateStr}:`, e);
    }
    return null;
}

async function fetchWeeklySalesArchive(mondayStr) {
    if (datasets.weekly && datasets.weekly.length > 0) {
        let activeMonday = getWeeklySalesMonday(datasets.weekly, "");
        if (activeMonday === mondayStr) {
            return datasets.weekly;
        }
    }
    try {
        let localData = await loadFromDB('weekly_sales_archive_' + mondayStr);
        if (localData) return localData;
    } catch (e) {
        console.error(`Lá»—i táº£i sales archive Local tuáº§n ${mondayStr}:`, e);
    }
    if (typeof firebase !== 'undefined') {
        try {
            let snapshot = await firebase.database().ref('archive_weekly_sales/' + mondayStr).once('value');
            let data = snapshot.val();
            if (data && data.data) {
                return data.data;
            }
        } catch (e) {
            console.error(`Lá»—i táº£i sales archive Cloud tuáº§n ${mondayStr}:`, e);
        }
    }
    return null;
}

async function loadWeeklyReview(startDateStr, endDateStr, filterMode) {
    const tbody = document.getElementById('weekly-review-tbody');
    tbody.innerHTML = `<tr><td colspan="15" style="text-align:center; padding: 2rem;">ðŸ”„ Äang tá»•ng há»£p dá»¯ liá»‡u...</td></tr>`;
    
    try {
        let dates = [];
        let curr = new Date(startDateStr);
        let end = new Date(endDateStr);
        while (curr <= end) {
            dates.push(formatDateStr(curr));
            curr.setDate(curr.getDate() + 1);
        }
        
        let orderArchives = await Promise.all(dates.map(d => fetchDailyOrderArchive(d)));
        
        let mondays = new Set();
        dates.forEach(dStr => {
            let d = new Date(dStr);
            let day = d.getDay();
            let diff = d.getDate() - day + (day === 0 ? -6 : 1);
            let mondayDate = new Date(d.getFullYear(), d.getMonth(), diff);
            mondays.add(formatDateStr(mondayDate));
        });
        
        let salesDataList = await Promise.all(Array.from(mondays).map(mondayStr => fetchWeeklySalesArchive(mondayStr)));
        
        buildMetadataMaps();
        buildProductWeightMap();
        
        let aggMap = new Map();
        
        // Populate orders
        for (let i = 0; i < dates.length; i++) {
            let dateStr = dates[i];
            let dayOfWeek = new Date(dateStr).getDay();
            let colIdx = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
            
            let dailyResults = orderArchives[i];
            if (dailyResults) {
                let list = Array.isArray(dailyResults) ? dailyResults : (dailyResults.results || []);
                list.forEach(item => {
                    let sap = extractSAP(item.sap);
                    if (!sap) return;
                    let region = item.region || globalStoreRegionMap.get(sap) || 'KhÃ¡c';
                    let storeName = item.store || globalStoreNamesMap.get(sap) || sap;
                    let productName = getGlobalNormalizedProduct(item.product);
                    if (!productName) return;
                    
                    let qty = getOrderedQty(item);
                    
                    let key = `${region}_${sap}_${storeName}_${productName}`;
                    if (!aggMap.has(key)) {
                        aggMap.set(key, {
                            region,
                            sap,
                            storeName,
                            productName,
                            orderDays: [0, 0, 0, 0, 0, 0, 0],
                            totalOrderPcs: 0,
                            totalOrderKg: 0,
                            totalSalesPcs: 0,
                            totalSalesKg: 0
                        });
                    }
                    
                    let data = aggMap.get(key);
                    data.orderDays[colIdx] += qty;
                    data.totalOrderPcs += qty;
                });
            }
        }
        
        function getOrderedQty(item) {
            if (item.final_order !== undefined && item.final_order !== null && item.final_order !== '') {
                return Number(item.final_order);
            }
            if (item.soq !== undefined && item.soq !== null && item.soq !== '') {
                let val = Number(item.soq);
                return isNaN(val) ? 0 : val;
            }
            return 0;
        }

        // Populate sales
        for (let idx = 0; idx < salesDataList.length; idx++) {
            let weeklyData = salesDataList[idx];
            let mondayStr = Array.from(mondays)[idx];
            if (!weeklyData || !Array.isArray(weeklyData)) continue;
            
            let overlapDays = 7;
            if (filterMode === 'date-range') {
                overlapDays = getOverlapDays(mondayStr, startDateStr, endDateStr);
            }
            if (overlapDays <= 0) continue;
            
            weeklyData.forEach(row => {
                let st = row['sap'] || row['storecode'] || row['nickname'] || row['storename'] || row['store'] || row['mach'] || row['tencuahang'];
                let pr = row['tnsnphmwm'] || row['tensanphamwm'] || row['tnsnphm'] || row['articlename'] || row['article'] || row['tensanpham'] || row['productname'];
                if (!pr) return;
                
                let productName = getGlobalNormalizedProduct(pr);
                if (!productName) return;
                
                if (st) {
                    let sap = extractSAP(st);
                    if (sap && isNaN(parseInt(sap))) {
                        let lookedUp = globalReverseStoreNamesMap.get(normalizeKey(st));
                        if (lookedUp) sap = lookedUp;
                    }
                    if (!sap) return;
                    
                    let rawDate = String(row['calendarday'] || row['date'] || row['ngay'] || '').trim();
                    let includeRow = false;
                    if (rawDate) {
                        let ts = parseDateStrToTime(rawDate);
                        if (ts > 0) {
                            let tsStart = new Date(startDateStr).getTime();
                            let tsEnd = new Date(endDateStr).getTime();
                            if (ts >= tsStart && ts <= tsEnd) {
                                includeRow = true;
                            }
                        }
                    } else {
                        includeRow = true;
                    }
                    
                    if (includeRow) {
                        let qty = Number(String(row['posquantity'] || row['sum'] || '0').replace(/,/g, ''));
                        if (pr && String(pr).toLowerCase().includes('retail kg')) qty /= 1000;
                        if (isNaN(qty)) return;
                        
                        let region = globalStoreRegionMap.get(sap) || 'KhÃ¡c';
                        let storeName = globalStoreNamesMap.get(sap) || sap;
                        let key = `${region}_${sap}_${storeName}_${productName}`;
                        
                        if (!aggMap.has(key)) {
                            aggMap.set(key, {
                                region,
                                sap,
                                storeName,
                                productName,
                                orderDays: [0, 0, 0, 0, 0, 0, 0],
                                totalOrderPcs: 0,
                                totalOrderKg: 0,
                                totalSalesPcs: 0,
                                totalSalesKg: 0
                            });
                        }
                        
                        let data = aggMap.get(key);
                        if (!rawDate && filterMode === 'date-range') {
                            data.totalSalesPcs += qty * (overlapDays / 7);
                        } else {
                            data.totalSalesPcs += qty;
                        }
                    }
                } else {
                    Object.entries(row).forEach(([colKey, qtyVal]) => {
                        let cKey = String(colKey).trim();
                        if (!cKey) return;
                        
                        let sap = "";
                        let sapMatch = cKey.match(/(\d{4,5})/);
                        if (sapMatch && globalReverseStoreNamesMap.has(normalizeKey(sapMatch[1]))) {
                            sap = globalReverseStoreNamesMap.get(normalizeKey(sapMatch[1]));
                        } else {
                            sap = globalReverseStoreNamesMap.get(normalizeKey(cKey));
                        }
                        
                        if (sap) {
                            let qty = Number(String(qtyVal || '0').replace(/,/g, ''));
                            if (pr && String(pr).toLowerCase().includes('retail kg')) qty /= 1000;
                            if (isNaN(qty) || qty === 0) return;
                            
                            let region = globalStoreRegionMap.get(sap) || 'KhÃ¡c';
                            let storeName = globalStoreNamesMap.get(sap) || sap;
                            let key = `${region}_${sap}_${storeName}_${productName}`;
                            
                            if (!aggMap.has(key)) {
                                aggMap.set(key, {
                                    region,
                                    sap,
                                    storeName,
                                    productName,
                                    orderDays: [0, 0, 0, 0, 0, 0, 0],
                                    totalOrderPcs: 0,
                                    totalOrderKg: 0,
                                    totalSalesPcs: 0,
                                    totalSalesKg: 0
                                });
                            }
                            
                            let data = aggMap.get(key);
                            data.totalSalesPcs += qty * (overlapDays / 7);
                        }
                    });
                }
            });
        }
        
        let list = Array.from(aggMap.values());
        list.forEach(item => {
            let weightKG = getProductWeightKG(item.productName);
            item.totalOrderKg = item.totalOrderPcs * weightKG;
            item.totalSalesKg = item.totalSalesPcs * weightKG;
        });
        
        currentWeeklyReviewList = list.filter(item => item.totalOrderPcs > 0 || item.totalSalesPcs > 0);
        
        // Sorting by SAP and CUSTOM_PRODUCT_ORDER
        currentWeeklyReviewList.sort((a, b) => {
            let sapCompare = String(a.sap).localeCompare(String(b.sap), undefined, { numeric: true });
            if (sapCompare !== 0) return sapCompare;
            
            let idxA = CUSTOM_PRODUCT_ORDER.findIndex(p => p.toLowerCase() === String(a.productName).trim().toLowerCase());
            let idxB = CUSTOM_PRODUCT_ORDER.findIndex(p => p.toLowerCase() === String(b.productName).trim().toLowerCase());
            
            idxA = idxA !== -1 ? idxA : 9999;
            idxB = idxB !== -1 ? idxB : 9999;
            
            if (idxA !== idxB) return idxA - idxB;
            return String(a.productName).localeCompare(String(b.productName), 'vi');
        });

        populateWeeklyRegionDropdown(currentWeeklyReviewList);
        filterWeeklyReviewTable();
        
    } catch (err) {
        console.error("Lá»—i tá»•ng há»£p bÃ¡o cÃ¡o tuáº§n:", err);
        tbody.innerHTML = `<tr><td colspan="15" style="text-align:center; padding: 2rem; color: var(--danger);">CÃ³ lá»—i xáº£y ra khi tá»•ng há»£p dá»¯ liá»‡u: ${err.message}</td></tr>`;
    }
}

function populateWeeklyRegionDropdown(dataList) {
    const selectEl = document.getElementById('weekly-filter-region');
    if (!selectEl) return;
    
    let selectedVal = selectEl.value;
    selectEl.innerHTML = '<option value="">ðŸ—ºï¸ Táº¥t cáº£ Khu vá»±c</option>';
    
    let regions = new Set();
    dataList.forEach(item => {
        if (item.region) regions.add(item.region);
    });
    
    let sortedRegions = Array.from(regions).sort((a, b) => a.localeCompare(b, 'vi'));
    sortedRegions.forEach(reg => {
        let opt = document.createElement('option');
        opt.value = reg;
        opt.text = reg;
        selectEl.appendChild(opt);
    });
    
    if (regions.has(selectedVal)) {
        selectEl.value = selectedVal;
    }
}

function renderWeeklyReviewTable(dataList) {
    const tbody = document.getElementById('weekly-review-tbody');
    tbody.innerHTML = '';
    
    if (dataList.length === 0) {
        tbody.innerHTML = `<tr><td colspan="16" style="text-align:center; padding: 2rem; color: var(--text-muted);">KhÃ´ng cÃ³ dá»¯ liá»‡u trong khoáº£ng thá»i gian Ä‘Æ°á»£c chá»n.</td></tr>`;
        document.getElementById('btn-export-weekly').style.display = 'none';
        return;
    }
    
    document.getElementById('btn-export-weekly').style.display = 'inline-block';
    
    // ThÃªm dÃ²ng Tá»”NG Cá»˜NG á»Ÿ Ä‘áº§u báº£ng review tuáº§n
    let totalOrderDays = [0, 0, 0, 0, 0, 0, 0];
    let totalOrderPcs = 0;
    let totalOrderKg = 0;
    let totalSalesPcs = 0;
    let totalSalesKg = 0;
    
    dataList.forEach(item => {
        for (let i = 0; i < 7; i++) {
            totalOrderDays[i] += item.orderDays[i] || 0;
        }
        totalOrderPcs += item.totalOrderPcs || 0;
        totalOrderKg += item.totalOrderKg || 0;
        totalSalesPcs += item.totalSalesPcs || 0;
        totalSalesKg += item.totalSalesKg || 0;
    });
    
    let totalDiffPcs = totalSalesPcs - totalOrderPcs;
    let totalDiffKg = totalSalesKg - totalOrderKg;
    
    let diffPcsColor = totalDiffPcs > 0 ? 'var(--success)' : (totalDiffPcs < 0 ? 'var(--danger)' : 'var(--text-muted)');
    let diffPcsText = totalDiffPcs > 0 ? '+' + totalDiffPcs.toLocaleString() : (totalDiffPcs < 0 ? totalDiffPcs.toLocaleString() : '-');
    
    let diffKgColor = totalDiffKg > 0 ? 'var(--success)' : (totalDiffKg < 0 ? 'var(--danger)' : 'var(--text-muted)');
    let diffKgText = totalDiffKg > 0 ? '+' + totalDiffKg.toFixed(2) : (totalDiffKg < 0 ? totalDiffKg.toFixed(2) : '-');
    
    let totalTr = document.createElement('tr');
    totalTr.style.fontWeight = 'bold';
    totalTr.style.background = 'rgba(255, 255, 255, 0.15)';
    totalTr.style.borderBottom = '2px solid var(--primary)';
    
    let totalOrderTds = '';
    for (let i = 0; i < 7; i++) {
        let val = totalOrderDays[i];
        totalOrderTds += `<td style="text-align: center; color: var(--primary); font-weight: bold;">${val > 0 ? val.toLocaleString() : '-'}</td>`;
    }
    
    totalTr.innerHTML = `
        <td colspan="3" style="text-align: center; color: var(--text-main); font-weight: bold; font-size: 1.05em;">Tá»”NG Cá»˜NG</td>
        ${totalOrderTds}
        <td style="text-align: center; color: var(--primary); font-weight: bold;">${totalOrderPcs > 0 ? totalOrderPcs.toLocaleString() : '-'}</td>
        <td style="text-align: center; color: var(--primary); font-weight: bold;">${totalOrderKg > 0 ? totalOrderKg.toFixed(2) : '-'}</td>
        <td style="text-align: center; color: var(--success); font-weight: bold;">${totalSalesPcs > 0 ? totalSalesPcs.toLocaleString() : '-'}</td>
        <td style="text-align: center; color: var(--success); font-weight: bold;">${totalSalesKg > 0 ? totalSalesKg.toFixed(2) : '-'}</td>
        <td style="text-align: center; color: ${diffPcsColor}; font-weight: bold;">${diffPcsText}</td>
        <td style="text-align: center; color: ${diffKgColor}; font-weight: bold;">${diffKgText}</td>
    `;
    tbody.appendChild(totalTr);

    // ThÃªm cÃ¡c dÃ²ng chi tiáº¿t phÃ­a sau
    dataList.forEach(item => {
        let tr = document.createElement('tr');
        
        let orderTds = '';
        for (let i = 0; i < 7; i++) {
            let val = item.orderDays[i];
            orderTds += `<td style="text-align: center;">${val > 0 ? val.toLocaleString() : '-'}</td>`;
        }
        
        let diffPcs = (item.totalSalesPcs || 0) - (item.totalOrderPcs || 0);
        let diffKg = (item.totalSalesKg || 0) - (item.totalOrderKg || 0);
        
        let diffPcsColor = diffPcs > 0 ? 'var(--success)' : (diffPcs < 0 ? 'var(--danger)' : 'var(--text-muted)');
        let diffPcsText = diffPcs > 0 ? '+' + diffPcs.toLocaleString() : (diffPcs < 0 ? diffPcs.toLocaleString() : '-');
        
        let diffKgColor = diffKg > 0 ? 'var(--success)' : (diffKg < 0 ? 'var(--danger)' : 'var(--text-muted)');
        let diffKgText = diffKg > 0 ? '+' + diffKg.toFixed(2) : (diffKg < 0 ? diffKg.toFixed(2) : '-');
        
        tr.innerHTML = `
            <td>${item.sap}</td>
            <td>${item.storeName}</td>
            <td>${item.productName}</td>
            ${orderTds}
            <td style="text-align: center; font-weight: bold; color: var(--primary);">${item.totalOrderPcs > 0 ? item.totalOrderPcs.toLocaleString() : '-'}</td>
            <td style="text-align: center; font-weight: bold; color: var(--primary);">${item.totalOrderKg > 0 ? item.totalOrderKg.toFixed(2) : '-'}</td>
            <td style="text-align: center; font-weight: bold; color: var(--success);">${item.totalSalesPcs > 0 ? item.totalSalesPcs.toLocaleString() : '-'}</td>
            <td style="text-align: center; font-weight: bold; color: var(--success);">${item.totalSalesKg > 0 ? item.totalSalesKg.toFixed(2) : '-'}</td>
            <td style="text-align: center; font-weight: bold; color: ${diffPcsColor};">${diffPcsText}</td>
            <td style="text-align: center; font-weight: bold; color: ${diffKgColor};">${diffKgText}</td>
        `;
        tbody.appendChild(tr);
    });
}

function filterWeeklyReviewTable() {
    let regionVal = document.getElementById('weekly-filter-region') ? document.getElementById('weekly-filter-region').value : '';
    let storeQuery = document.getElementById('weekly-search-store') ? removeAccents(document.getElementById('weekly-search-store').value.toLowerCase().trim()) : '';
    let productQuery = document.getElementById('weekly-search-product') ? removeAccents(document.getElementById('weekly-search-product').value.toLowerCase().trim()) : '';
    
    let filtered = currentWeeklyReviewList.filter(item => {
        if (regionVal && item.region !== regionVal) return false;
        
        if (storeQuery) {
            let sap = item.sap.toLowerCase();
            let store = removeAccents(item.storeName.toLowerCase());
            if (!sap.includes(storeQuery) && !store.includes(storeQuery)) return false;
        }
        
        if (productQuery) {
            let prod = removeAccents(item.productName.toLowerCase());
            if (!prod.includes(productQuery)) return false;
        }
        
        return true;
    });
    
    currentWeeklyFilteredList = filtered;
    renderWeeklyReviewTable(filtered);
}

function exportWeeklyReviewToExcel() {
    if (currentWeeklyFilteredList.length === 0) {
        alert("KhÃ´ng cÃ³ dá»¯ liá»‡u Ä‘á»ƒ xuáº¥t Excel!");
        return;
    }
    
    let aoa = [];
    aoa.push([
        "Khu Vá»±c", "MÃ£ SAP", "TÃªn Cá»­a HÃ ng", "TÃªn Sáº£n Pháº©m",
        "Thá»© 2", "Thá»© 3", "Thá»© 4", "Thá»© 5", "Thá»© 6", "Thá»© 7", "Chá»§ Nháº­t",
        "Tá»•ng Äáº·t (Pcs)", "Tá»•ng Äáº·t (Kg)", "Tá»•ng BÃ¡n (Pcs)", "Tá»•ng BÃ¡n (Kg)"
    ]);
    
    // TÃ­nh toÃ¡n dÃ²ng Tá»”NG Cá»˜NG trÆ°á»›c
    let totalOrderDays = [0, 0, 0, 0, 0, 0, 0];
    let totalOrderPcs = 0;
    let totalOrderKg = 0;
    let totalSalesPcs = 0;
    let totalSalesKg = 0;
    
    currentWeeklyFilteredList.forEach(item => {
        for (let i = 0; i < 7; i++) {
            totalOrderDays[i] += item.orderDays[i] || 0;
        }
        totalOrderPcs += item.totalOrderPcs || 0;
        totalOrderKg += item.totalOrderKg || 0;
        totalSalesPcs += item.totalSalesPcs || 0;
        totalSalesKg += item.totalSalesKg || 0;
    });
    
    aoa.push([
        "Tá»”NG Cá»˜NG",
        "",
        "",
        "",
        totalOrderDays[0] || 0,
        totalOrderDays[1] || 0,
        totalOrderDays[2] || 0,
        totalOrderDays[3] || 0,
        totalOrderDays[4] || 0,
        totalOrderDays[5] || 0,
        totalOrderDays[6] || 0,
        totalOrderPcs || 0,
        totalOrderKg || 0,
        totalSalesPcs || 0,
        totalSalesKg || 0
    ]);

    // ThÃªm cÃ¡c dÃ²ng chi tiáº¿t phÃ­a sau
    currentWeeklyFilteredList.forEach(item => {
        aoa.push([
            item.region,
            item.sap,
            item.storeName,
            item.productName,
            item.orderDays[0] || 0,
            item.orderDays[1] || 0,
            item.orderDays[2] || 0,
            item.orderDays[3] || 0,
            item.orderDays[4] || 0,
            item.orderDays[5] || 0,
            item.orderDays[6] || 0,
            item.totalOrderPcs || 0,
            item.totalOrderKg || 0,
            item.totalSalesPcs || 0,
            item.totalSalesKg || 0
        ]);
    });
    
    const worksheet = XLSX.utils.aoa_to_sheet(aoa);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Tong_Hop_Tuan");
    
    let titleStr = "Tong_Hop_Tuan";
    let isWeekMode = document.querySelector('input[name="filter-mode"]:checked').value === 'week';
    if (isWeekMode) {
        let selectEl = document.getElementById('weekly-review-select');
        if (selectEl && selectEl.value) {
            let label = selectEl.options[selectEl.selectedIndex].text;
            titleStr = "Tong_Hop_" + label.replace(/[^a-zA-Z0-9]/g, "_");
        }
    } else {
        let startStr = document.getElementById('weekly-start-date').value;
        let endStr = document.getElementById('weekly-end-date').value;
        if (startStr && endStr) {
            titleStr = `Tong_Hop_${startStr}_to_${endStr}`;
        }
    }
    
    let exportName = `${titleStr}.xlsx`;
    
    try {
        let wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });
        function s2ab(s) {
            let buf = new ArrayBuffer(s.length);
            let view = new Uint8Array(buf);
            for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
            return buf;
        }
        let blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
        let url = URL.createObjectURL(blob);
        let a = document.createElement("a");
        document.body.appendChild(a);
        a.href = url;
        a.download = exportName;
        a.click();
        setTimeout(() => {
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
        }, 0);
    } catch (e) {
        console.error("Lá»—i xuáº¥t Excel tuáº§n:", e);
        alert("Lá»—i xuáº¥t Excel: " + e.message);
    }
}

// Event Listeners for Weekly Review
const navWeeklyReview = document.getElementById('nav-weekly-review');
const weeklyReviewSelect = document.getElementById('weekly-review-select');
const btnLoadDateRange = document.getElementById('btn-load-date-range');
const btnExportWeekly = document.getElementById('btn-export-weekly');
const weeklyFilterRegion = document.getElementById('weekly-filter-region');
const weeklySearchStore = document.getElementById('weekly-search-store');
const weeklySearchProduct = document.getElementById('weekly-search-product');

if (navWeeklyReview) {
    navWeeklyReview.addEventListener('click', (e) => {
        e.preventDefault();
        document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
        navWeeklyReview.classList.add('active');
        isHistoryView = false;
        isArchiveView = false;
        
        document.querySelector('.upload-section').style.display = 'none';
        document.getElementById('archive-selector-container').style.display = 'none';
        document.getElementById('results-section').style.display = 'none';
        document.getElementById('weekly-review-container').style.display = 'block';
        
        // Initialize Default Values
        let today = new Date();
        let sevenDaysAgo = new Date();
        sevenDaysAgo.setDate(today.getDate() - 7);
        
        document.getElementById('weekly-start-date').value = formatDateStr(sevenDaysAgo);
        document.getElementById('weekly-end-date').value = formatDateStr(today);
        
        // Populate Weeks Dropdown
        populateWeeksDropdown();
        
        // Auto Load Current Week
        if (weeklyReviewSelect && weeklyReviewSelect.options.length > 1) {
            weeklyReviewSelect.selectedIndex = 1;
            let val = JSON.parse(weeklyReviewSelect.value);
            loadWeeklyReview(val.start, val.end, 'week');
        } else {
            let start = document.getElementById('weekly-start-date').value;
            let end = document.getElementById('weekly-end-date').value;
            loadWeeklyReview(start, end, 'date-range');
        }
    });
}

function populateWeeksDropdown() {
    const selectEl = document.getElementById('weekly-review-select');
    if (!selectEl) return;
    
    selectEl.innerHTML = '<option value="">-- Chá»n Tuáº§n --</option>';
    let weeks = getWeeksOfYear();
    weeks.forEach(w => {
        let opt = document.createElement('option');
        opt.value = JSON.stringify({ start: w.mondayStr, end: w.sundayStr });
        opt.text = w.label;
        selectEl.appendChild(opt);
    });
}

// Handle Mode Selector Switch
document.querySelectorAll('input[name="filter-mode"]').forEach(radio => {
    radio.addEventListener('change', (e) => {
        let mode = e.target.value;
        if (mode === 'week') {
            document.getElementById('week-selector-wrap').style.display = 'flex';
            document.getElementById('date-range-selector-wrap').style.display = 'none';
            if (weeklyReviewSelect && weeklyReviewSelect.value) {
                let val = JSON.parse(weeklyReviewSelect.value);
                loadWeeklyReview(val.start, val.end, 'week');
            }
        } else {
            document.getElementById('week-selector-wrap').style.display = 'none';
            document.getElementById('date-range-selector-wrap').style.display = 'flex';
            let start = document.getElementById('weekly-start-date').value;
            let end = document.getElementById('weekly-end-date').value;
            if (start && end) {
                loadWeeklyReview(start, end, 'date-range');
            }
        }
    });
});

if (weeklyReviewSelect) {
    weeklyReviewSelect.addEventListener('change', () => {
        if (weeklyReviewSelect.value) {
            let val = JSON.parse(weeklyReviewSelect.value);
            loadWeeklyReview(val.start, val.end, 'week');
        }
    });
}

if (btnLoadDateRange) {
    btnLoadDateRange.addEventListener('click', () => {
        let start = document.getElementById('weekly-start-date').value;
        let end = document.getElementById('weekly-end-date').value;
        if (!start || !end) {
            alert("Vui lÃ²ng nháº­p Ä‘áº§y Ä‘á»§ Tá»« ngÃ y vÃ  Äáº¿n ngÃ y!");
            return;
        }
        loadWeeklyReview(start, end, 'date-range');
    });
}

if (weeklyFilterRegion) {
    weeklyFilterRegion.addEventListener('change', filterWeeklyReviewTable);
}
if (weeklySearchStore) {
    weeklySearchStore.addEventListener('input', filterWeeklyReviewTable);
}
if (weeklySearchProduct) {
    weeklySearchProduct.addEventListener('input', filterWeeklyReviewTable);
}

if (btnExportWeekly) {
    btnExportWeekly.addEventListener('click', exportWeeklyReviewToExcel);
}
