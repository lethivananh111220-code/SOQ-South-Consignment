// Danh sách rau ăn lá/RTE (Tỷ lệ hủy > 30%)
const RTE_PRODUCTS = [
    "Cải hoa hồng baby", "Cải Kale xoăn", "Cải Kale khủng long", "Bông cải xanh baby",
    "Xà lách frisée xanh ngọt", "Xà lách romaine xanh thượng hạng", "Xà lách frisée tím ngọt",
    "Xà lách romaine tím thượng hạng", "Xà lách baby lollo", "Xà lách baby thủy tinh",
    "Cải ngọt giống nhật", "Cải bó xôi", "Xà lách hỗn hợp", "Asian Mix",
    "Gourmet Italian Mix", "Sweet Baby Lettuces", "Baby Spring Mix", "Chopped Kale",
    "Pure Rocket", "Cải bó xôi baby ăn liền"
];

const datasets = {
    schedule: null,
    inventory: null,
    input: null,
    monthly: null,
    weekly: null,
    mapping_raw: null
};

let scheduleFileName = "SOQ_Calculated_Order"; // Tên mặc định

// Elements
const btnCalculate = document.getElementById('btn-calculate');
const btnExport = document.getElementById('btn-export');
const resultsSection = document.getElementById('results-section');
const tbody = document.getElementById('soq-tbody');

// Helper to normalize column names
function normalizeKey(key) {
    if(!key) return '';
    return key.toString().toLowerCase().replace(/[^a-z0-9\/\-]/g, '');
}

// Hàm trích xuất tự động bỏ qua các tiêu đề báo cáo rác ở file hệ thống (Excel report info)
function extractJsonDataCleanly(worksheet) {
    let rawArr = XLSX.utils.sheet_to_json(worksheet, {header: 1, raw: false, dateNF: 'yyyy-mm-dd hh:mm:ss'});
    if (!rawArr || rawArr.length === 0) return [];
    
    let headerIdx = 0;
    // Tìm dòng header thực sự (Thường có chứa các chữ khóa nhận diện và > 3 cột dữ liệu)
    for (let i = 0; i < Math.min(20, rawArr.length); i++) {
        let r = rawArr[i];
        if (!r) continue;
        let validCols = r.filter(c => typeof c === 'string' && c.trim() !== '');
        if (validCols.length >= 2 && r.some(c => typeof c==='string' && (c.toUpperCase().includes('SAP') || c.toUpperCase().includes('STORE') || c.toUpperCase().includes('NICKNAME') || c.toUpperCase().includes('TÊN') || c.toUpperCase().includes('ARTICLE') || c.toUpperCase().includes('PRODUCT')))) {
            headerIdx = i;
            break;
        }
    }
    
    const headersRaw = rawArr[headerIdx] || [];
    const headers = headersRaw.map(h => normalizeKey(h));
    const json = [];
    
    for (let i = headerIdx + 1; i < rawArr.length; i++) {
        let row = rawArr[i];
        if (!row || row.length === 0) continue;
        
        // Bỏ qua dòng Total (Dòng tổng cộng của SAP)
        if (row.some(cell => String(cell).toUpperCase().includes('RESULT') || String(cell).toUpperCase() === 'TOTAL')) continue;

        let obj = {};
        let hasData = false;
        for (let j = 0; j < headers.length; j++) {
            if (row[j] !== undefined && row[j] !== null && String(row[j]).trim() !== '') {
                obj[headers[j]] = row[j];
                obj[headersRaw[j]] = row[j]; // Giữ nguyên key gốc để fallback
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
        return new Date(Math.round((val - 25569) * 86400 * 1000)).getTime();
    }
    let s = String(val).trim();
    // Support DD/MM/YYYY or DD-MM-YYYY formats commonly exported in VN
    let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (m) {
        return new Date(m[3], m[2] - 1, m[1]).getTime();
    }
    const parsed = new Date(s).getTime();
    if (!isNaN(parsed)) return parsed;
    return 0; // fallback numerical value
}

// Tính tổng nhu cầu động dựa trên loại ngày (thứ 2-5 hoặc thứ 6-CN)
function calculatePeriodDemand(startTs, totalDays, adsWeekday, adsWeekend) {
    let total = 0;
    let roundedDays = Math.round(totalDays);
    if (roundedDays <= 0) return 0;
    
    for (let i = 0; i < roundedDays; i++) {
        let currentTs = startTs + (i * 86400 * 1000);
        let d = new Date(currentTs);
        let dw = d.getDay(); // 0:CN, 1:T2, ... 6:T7
        if (dw >= 1 && dw <= 4) {
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
    statusEl.textContent = `Đang đọc ${file.name}...`;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array', cellDates: true});
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        if (type === 'mapping') {
            const arr = XLSX.utils.sheet_to_json(worksheet, {header: 1});
            datasets['mapping_raw'] = arr;
            saveToDB('mapping_raw', arr);
            statusEl.textContent = `Đã tải & lưu trữ: ${file.name} (${arr.length} dòng)`;
            statusEl.classList.add('success');
            checkReady();
            return;
        }

        try {
            const json = extractJsonDataCleanly(worksheet);
            datasets[type] = json;
            if (type === 'monthly' || type === 'weekly') {
                saveToDB(type, json);
                statusEl.textContent = `Đã tải & lưu trữ: ${file.name} (${json.length} dòng)`;
            } else {
                if (type === 'schedule') {
                    scheduleFileName = file.name.replace(/\.[^/.]+$/, ""); 
                }
                statusEl.textContent = `Đã tải: ${file.name} (${json.length} dòng)`;
            }
            statusEl.classList.add('success');
            checkReady();
        } catch (err) {
            console.error(err);
            statusEl.textContent = "Lỗi xử lý file: " + err.message;
            statusEl.style.color = "var(--danger)";
        }
    };
    reader.onerror = () => {
        statusEl.textContent = "Lỗi đọc file từ máy tính!";
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
document.getElementById('file-mapping').addEventListener('change', e => handleFileUpload(e, 'mapping'));

// --- IndexedDB Caching cho các file cố định (Monthly, Weekly, Mapping) ---
const DB_NAME = "SOQ_V1";
function initDB() {
    return new Promise((resolve, reject) => {
        let req = indexedDB.open(DB_NAME, 1);
        req.onupgradeneeded = e => {
            let db = e.target.result;
            if(!db.objectStoreNames.contains('files')) db.createObjectStore('files');
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
    } catch(e) {}
}

async function saveToDB(key, data) {
    try {
        let payload = { data: data, timestamp: Date.now() };
        let db = await initDB();
        let tx = db.transaction('files', 'readwrite');
        tx.objectStore('files').put(payload, key);
    } catch(e) { console.error('Lỗi lưu cache', e); }
}

function getWeekStart(date) {
    let d = new Date(date);
    let day = d.getDay(); 
    let diff = d.getDate() - day + (day === 0 ? -6 : 1);
    d.setDate(diff);
    d.setHours(0,0,0,0);
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
        if (Array.isArray(raw)) return raw; // Cache cổ điển 

        if (raw.timestamp && raw.data) {
            let dDate = new Date(raw.timestamp);
            let nDate = new Date();
            
            if (key === 'monthly') {
                if (dDate.getMonth() !== nDate.getMonth() || dDate.getFullYear() !== nDate.getFullYear()) {
                    await deleteFromDB(key);
                    return { invalidated: true, reason: "sang tháng mới" };
                }
            } else if (key === 'weekly') {
                if (getWeekStart(dDate) !== getWeekStart(nDate)) {
                    await deleteFromDB(key);
                    return { invalidated: true, reason: "sang tuần mới" };
                }
            }
            return raw.data;
        }
        return raw;
    } catch(e) { return null; }
}

window.addEventListener('DOMContentLoaded', async () => {
    // Tự sinh Dropdown Chọn Ngày 1-31, mặc định nhảy vào số trùng với Hôm Nay (Today)
    let dateSelect = document.getElementById('targetDeliveryDate');
    if (dateSelect) {
        let tomorrow = new Date();
        tomorrow.setDate(tomorrow.getDate() + 1);
        let targetDay = tomorrow.getDate();

        for (let i = 1; i <= 31; i++) {
            let opt = document.createElement('option');
            opt.value = i;
            opt.text = "Ngày " + i;
            if (i === targetDay) opt.selected = true;
            dateSelect.appendChild(opt);
        }
    }

    // Tự động load lại Cache của Monthly, Weekly, Mapping File nếu có
    let [cMonthly, cWeekly, cMapping] = await Promise.all([
        loadFromDB('monthly'), loadFromDB('weekly'), loadFromDB('mapping_raw')
    ]);

    if(cMonthly) {
        if (cMonthly.invalidated) {
            let el = document.getElementById('status-monthly');
            if (el) { el.innerHTML = `<span style="color: #ff9800; font-weight: bold;">Lưu ý: Đã sang tháng mới. Vui lòng Tải Lên file cập nhật!</span>`; el.classList.remove('success'); }
        } else if (cMonthly.length > 0) {
            datasets.monthly = cMonthly;
            let el = document.getElementById('status-monthly');
            if (el) { el.textContent = `Đã dùng bản lưu trước (${cMonthly.length} dòng)`; el.classList.add('success'); }
        }
    }
    
    if(cWeekly) {
        if (cWeekly.invalidated) {
            let el = document.getElementById('status-weekly');
            if (el) { el.innerHTML = `<span style="color: #ff9800; font-weight: bold;">Lưu ý: Đã sang tuần mới (T2). Vui lòng Tải Lên file cập nhật!</span>`; el.classList.remove('success'); }
        } else if (cWeekly.length > 0) {
            datasets.weekly = cWeekly;
            let el = document.getElementById('status-weekly');
            if (el) { el.textContent = `Đã dùng bản lưu trước (${cWeekly.length} dòng)`; el.classList.add('success'); }
        }
    }
    if(cMapping && cMapping.length > 0) {
        datasets.mapping_raw = cMapping;
        let el = document.getElementById('status-mapping');
        if (el) { el.textContent = `Đã dùng bản lưu trước (${cMapping.length} dòng)`; el.classList.add('success'); }
    }
    
    checkReady();
});

function checkReady() {
    if (datasets.schedule && datasets.inventory && datasets.input && datasets.monthly && datasets.weekly) {
        btnCalculate.disabled = false;
        btnCalculate.textContent = "Tiến hành tính SOQ";
    }
}

let finalResults = [];

function extractSAP(str) {
    if (!str) return "";
    let s = String(str).trim();
    let m = s.match(/\d+/);
    return m ? Number(m[0]).toString() : s.toLowerCase();
}

btnCalculate.addEventListener('click', () => {

    // --- TÍNH TOÁN NGÀY GIAO HÀNG (WEEKEND HAY WEEKDAY) ---
    const getWeekdayIdxGlobal = (str) => {
        let s = String(str).trim().toLowerCase();
        const w = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
        let idx = w.indexOf(s);
        if (idx !== -1) return idx;
        if (s === 'cn' || s === 'chủ nhật' || s === 'sun') return 0;
        if (s === 't2' || s === 'thứ 2' || s === 'thứ hai' || s === 'mon') return 1;
        if (s === 't3' || s === 'thứ 3' || s === 'thứ ba' || s === 'tue') return 2;
        if (s === 't4' || s === 'thứ 4' || s === 'thứ tư' || s === 'wed') return 3;
        if (s === 't5' || s === 'thứ 5' || s === 'thứ năm' || s === 'thu') return 4;
        if (s === 't6' || s === 'thứ 6' || s === 'thứ sáu' || s === 'fri') return 5;
        if (s === 't7' || s === 'thứ 7' || s === 'thứ bảy' || s === 'sat') return 6;
        return -1;
    };

    let targetDateStr = document.getElementById('targetDeliveryDate') ? document.getElementById('targetDeliveryDate').value.trim() : "";
    let isWeekendDelivery = false;
    let targetTimestamp = 0; // Để tính toán Lead Time Arrival

    if (targetDateStr !== "") {
        let isTgtWkday = getWeekdayIdxGlobal(targetDateStr) !== -1;
        let tgtNum = isTgtWkday ? getWeekdayIdxGlobal(targetDateStr) : parseInt((targetDateStr.match(/^(\d{1,2})/) || [])[1] || 0);
        let finalWkday = -1;

        let dTarget = new Date();
        dTarget.setHours(0, 0, 0, 0);

        if (isTgtWkday) {
            finalWkday = tgtNum;
            // Tìm ngày gần nhất khớp với thứ được chọn (ví dụ Thứ 6 gần nhất)
            let diff = (tgtNum - dTarget.getDay() + 7) % 7;
            // Nếu diff = 0 thì có thể là hôm nay, nhưng thường là đặt cho tuần sau hoặc hôm nay vẫn tính sales?
            // Giữ nguyên logic cũ cho finalWkday nhưng tính thêm timestamp
            dTarget.setDate(dTarget.getDate() + diff);
        } else if (tgtNum > 0) {
            // Nếu ngày gõ < hôm nay quá nhiều (ví dụ nay 28, gõ 2) -> Sang tháng sau
            if (tgtNum < dTarget.getDate() - 7) {
                dTarget.setMonth(dTarget.getMonth() + 1);
            }
            dTarget.setDate(tgtNum);
            finalWkday = dTarget.getDay();
        }
        
        targetTimestamp = dTarget.getTime();

        // Cuối tuần: Thứ 6 (5), Thứ 7 (6), Chủ nhật (0)
        if (finalWkday === 5 || finalWkday === 6 || finalWkday === 0) {
            isWeekendDelivery = true;
        }
    }

    // ----------- 1. Map Rules (WM Name -> ODA Name) -----------
    const mappingMap = new Map();
    const reverseMappingKeys = new Set(); // Dùng để kiểm tra sản phẩm lạ

    if (datasets.mapping_raw && datasets.mapping_raw.length > 0) {
        for (let i = 0; i < datasets.mapping_raw.length; i++) {
            let r = datasets.mapping_raw[i];
            if (!r || !Array.isArray(r)) continue;
            let validCells = r.filter(val => val !== null && val !== undefined && String(val).trim() !== '');
            if (validCells.length >= 3) {
                // Rule từ thực tế: Dòng có STT(0) | Tên ODA(1) | Tên WM(2)
                let odaName = String(validCells[1]).trim();
                let wmName = String(validCells[2]).trim().toLowerCase();
                if (wmName && odaName && wmName.toLowerCase() !== 'tên sản phẩm wm') {
                    mappingMap.set(wmName, odaName);
                    reverseMappingKeys.add(wmName);
                }
            } else if (validCells.length === 2) {
                let wmName = String(validCells[0]).trim().toLowerCase();
                mappingMap.set(wmName, String(validCells[1]).trim());
                reverseMappingKeys.add(wmName);
            }
        }
    }

    const normalizeProductName = (name) => {
        let n = String(name).trim().toLowerCase();
        return mappingMap.has(n) ? String(mappingMap.get(n)).trim() : String(name).trim();
    }

    // ----------- 2. Schedule Filter & Store Names -----------
    const validSAPs = new Set();
    const storeNamesMap = new Map(); // Lưu Tên Cửa Hàng
    const scheduleLeadtimeMap = new Map(); // Leadtime riêng theo CH
    const storeTierMap = new Map(); // Lưu Tier cửa hàng
    // targetDateStr đã được lấy ở trên cùng hàm btnCalculate
    datasets.schedule.forEach(row => {
        // Lưu ý: Các key trong row đã được normalize (viết thường, bỏ khoảng trắng)
        let store = row['sap'] || row['storekey'] || row['storecode'] || row['makho'] || row['mach'] || row['mãkháchhàng'] || row['mãcửahàng'];
        if (!store) return;
        
        let storeID = extractSAP(store);
        let hinhThuc = String(row['hinhthuc'] || row['Hình thức'] || row['type'] || '').toUpperCase();
        
        let dynamicLT = 0;

        // Hàm hỗ trợ đọc kiểu chữ (thứ trong tuần)
        const getWeekdayIdx = getWeekdayIdxGlobal;

        // --- BỘ LỌC NGÀY GIAO HÀNG (TARGET DELIVERY DATE) ---
        // Giờ đây Web sẽ TỰ ĐỘNG map Day (vd "20") thành Weekday ("Friday") nếu Lịch dùng Thứ!
        if (targetDateStr !== "") {
            let hasDelivery = false;
            
            let isTargetWeekday = getWeekdayIdx(targetDateStr) !== -1;
            let currentTargetNum = isTargetWeekday ? getWeekdayIdx(targetDateStr) : parseInt((targetDateStr.match(/^(\d{1,2})/) || [])[1] || 0);

            // Đoán Thứ bằng cách tính lịch (Nếu gõ Ngày 20 -> Gán thành Thứ Sáu (5) trong tháng hiện tại)
            let impliedWeekdayIdx = -1;
            if (!isTargetWeekday && currentTargetNum > 0) {
                 let d = new Date();
                 d.setDate(currentTargetNum);
                 impliedWeekdayIdx = d.getDay(); // 0(CN) -> 6(T7)
            }

            let possibleNextWeekdayIdx = [];
            let possibleNextDigitDays = [];

            for (const [key, val] of Object.entries(row)) {
                let k = String(key).trim();
                let match = false;
                
                let headerWeekdayIdx = getWeekdayIdx(k);

                // Nếu Header file Lịch là THỦ (VD: Friday, T2)
                if (headerWeekdayIdx !== -1) {
                    if (isTargetWeekday) {
                        match = (headerWeekdayIdx === currentTargetNum);
                    } else if (impliedWeekdayIdx !== -1) {
                        // Người dùng đang chọn NGÀY (VD 20), Web sẽ tự gán Thứ Sáu (5) = Friday (5)
                        match = (headerWeekdayIdx === impliedWeekdayIdx); 
                    }
                } else {
                    // Nếu Header file Lịch là SỐ NGÀY (VD: 20, 20/03)
                    match = (k === targetDateStr || k.startsWith(targetDateStr + '/') || k.startsWith(targetDateStr + '-'));
                }

                let v = String(val).trim().toLowerCase().replace(/\s+/g, ''); 
                if (v && v !== '0' && v !== 'false' && v !== 'off' && !v.includes('nghỉ') && v !== 'shipper') {
                    if (match) {
                        hasDelivery = true;
                    }
                    // Theo dõi tất cả các mốc có giao hàng tiếp theo
                    if (headerWeekdayIdx !== -1) {
                        possibleNextWeekdayIdx.push(headerWeekdayIdx);
                    } else {
                        let m = k.match(/^(\d{1,2})/);
                        if (m) possibleNextDigitDays.push(parseInt(m[1]));
                    }
                }
            }

            // Nếu không có lịch giao -> Bỏ qua
            if (!hasDelivery) return; 

            // --- TÍNH TOÁN LEADTIME ĐỘNG TỪ MA TRẬN LỊCH GIAO HÀNG ---
            if (possibleNextWeekdayIdx.length > 0) {
                // Ma trận đang dùng tên THỨ -> Tính khoảng cách bằng hệ tuần hoàn 7 ngày
                let anchorIdx = isTargetWeekday ? currentTargetNum : impliedWeekdayIdx;
                let future = possibleNextWeekdayIdx.filter(d => d !== anchorIdx && d > anchorIdx);
                if (future.length > 0) {
                    dynamicLT = Math.min(...future) - anchorIdx;
                } else {
                    let past = possibleNextWeekdayIdx.filter(d => d !== anchorIdx && d < anchorIdx);
                    // Rollover sang tuần mới
                    if (past.length > 0) dynamicLT = (7 - anchorIdx) + Math.min(...past);
                }
            } else if (possibleNextDigitDays.length > 0) {
                // Ma trận đang dùng số NGÀY -> Tính khoảng cách bằng toán thông thường
                let anchorNum = currentTargetNum;
                let future = possibleNextDigitDays.filter(d => d !== anchorNum && d > anchorNum);
                if (future.length > 0) {
                    dynamicLT = Math.min(...future) - anchorNum;
                } else {
                    let past = possibleNextDigitDays.filter(d => d !== anchorNum && d < anchorNum);
                    // Rollover sang tháng mới
                    if (past.length > 0) dynamicLT = (30 - anchorNum) + Math.min(...past);
                }
            }
        }

        // Mặc định: Chấp nhận TẤT CẢ các mã cửa hàng miễn là có tên trong file Lịch Giao Hàng
        validSAPs.add(storeID);

        let sName = row['tencuahang'] || row['tncahng'] || row['storename'] || row['store'];
        if (sName) storeNamesMap.set(storeID, String(sName).trim());

        // LƯU CỘT TIER
        let tierVal = String(row['tier'] || row['Tier'] || row['cấpđộ'] || row['phânloại'] || '').trim().toUpperCase();
        if (tierVal && tierVal !== 'UNDEFINED') storeTierMap.set(storeID, tierVal);

        if (dynamicLT > 0) {
            scheduleLeadtimeMap.set(storeID, dynamicLT); // Gắn chặt Leadtime từ matrix lịch
        } else {
            let lt = Number(row['leadtime'] || row['Leadtime'] || row['chu kỳ'] || row['chukỳ'] || 0);
            if (lt > 0) scheduleLeadtimeMap.set(storeID, lt);
        }
    });

    // Helper: Bóc tách Leadtime từ tên file Lịch Giao Hàng (VD: Lịch 2003-2203 -> 3 ngày)
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

    // ----------- 3. Inventory Aggregation -----------
    const inventoryMap = new Map(); 
    datasets.inventory.forEach(row => {
        let store = row['sap'] || row['storecode'];
        let prod = row['productname'] || row['listsnphm'] || row['tnsnphm'];
        if (!store || !prod) return;
        
        let storeID = extractSAP(store);
        let sName = row['tencuahang'] || row['tncahng'] || row['storename'] || row['store'];
        if (sName && !storeNamesMap.has(storeID)) storeNamesMap.set(storeID, String(sName).trim());

        let prodStd = normalizeProductName(prod); 
        let key = `${storeID}_${prodStd.toLowerCase()}`;
        
        let inv = Number(String(row['inventoryamount'] || row['Inventory Amount'] || row['inventoryquantity'] || row['Tồn kho'] || '0').replace(/,/g, ''));
        let disp = Number(String(row['disposalamount'] || row['Disposal Amount'] || row['disposalquantity'] || row['Hủy'] || '0').replace(/,/g, ''));
        
        // Cấu hình quy đổi riêng cho mã Bông cải xanh (RETAIL KG) từ Gam -> Kg
        if (String(prod).toLowerCase().includes('bông cải xanh (retail kg)')) {
            inv = inv / 1000;
            disp = disp / 1000;
        }

        let rawDate = row['date'] || row['Date'] || row['ngay'] || row['ngày'] || 0;
        let cDate = parseDateStrToTime(rawDate);

        // Chỉ lấy Tồn và Hủy của ngày mới nhất
        if (!inventoryMap.has(key)) {
            inventoryMap.set(key, {inv, disp, latestDate: cDate, prodOrig: prodStd});
        } else {
            let data = inventoryMap.get(key);
            if (cDate > data.latestDate) {
                // Nhận diện ngày mới hơn -> Thay thế (Ghi đè)
                data.latestDate = cDate;
                data.inv = inv;
                data.disp = disp;
            } else if (cDate === data.latestDate) {
                // Cùng 1 ngày (ví dụ cùng store tách list) -> Cộng dồn
                data.inv += inv;
                data.disp += disp;
            }
        }
    });

    // ----------- 4. Input ODA Aggregation -----------
    const inputMap = new Map();
    const actualODA_Names = new Map(); // Lưu Tên ODA chuẩn nhất từ file vận hành

    datasets.input.forEach(row => {
        let store = row['nickname'] || row['sap'];
        let prod = row['productnameprimarylanguage'] || row['productname'] || row['product'];
        let status = String(row['orderstatus'] || row['status'] || '').toLowerCase();
        
        if (!store || !prod) return;

        // Lọc bỏ hàng Hủy / Đã hoàn (Chỉ lấy Completed)
        if (status && (status.includes('cancel') || status.includes('hủy') || status.includes('reject'))) return;
        
        let storeID = extractSAP(store);
        let exactODAName = String(prod).trim();
        let prodStd = normalizeProductName(prod); 
        
        let key = `${storeID}_${prodStd.toLowerCase()}`;
        actualODA_Names.set(prodStd.toLowerCase(), exactODAName);

        let qty = Number(String(row['quantity'] || row['Quantity'] || '0').replace(/,/g, ''));
        
        // Trích xuất ngày giao hàng/nhập hàng
        let rawDate = row['completeddate'] || row['Completed date'] || row['orderdate'] || row['Order date'] || row['date'] || 0;
        let cDate = parseDateStrToTime(rawDate);
        
        if (!inputMap.has(key)) {
            inputMap.set(key, {qty, latestDate: cDate, prodOrig: exactODAName}); 
        } else {
            let current = inputMap.get(key);
            if (cDate > current.latestDate) {
                // Nhập mới nhất -> Thay thế hoàn toàn số lượng (Ghi đè)
                current.latestDate = cDate;
                current.qty = qty;
            } else if (cDate === current.latestDate) {
                // Nếu 2 dòng cùng 1 ngày (ví dụ bị tách bill) -> Cộng dồn
                current.qty += qty;
            }
            // Nếu ngày cũ hơn -> Bỏ qua
        }
    });

    // Hàm lấy lại Tên Chuẩn nhất (Ưu tiên ODA thật > Mapping > Raw)
    const getBestAvailableName = (mappedName) => {
        let k = mappedName.toLowerCase();
        return actualODA_Names.has(k) ? actualODA_Names.get(k) : mappedName;
    }

    // ----------- 5. Sales Data (Flat Transaction Aggregation) -----------
    // Trong file thực tế: Dữ liệu doanh số bán nằm từng dòng, cột "POS Quantity"
    const monthlySales = new Map();
    const storeMonthlyDays = new Map(); // All days
    const storeGroupDays = new Map();  // storeID -> { weekdays: Set, weekends: Set }

    datasets.monthly.forEach(row => {
        let st = row['sap'] || row['storecode'];
        let pr = row['tnsnphmwm'] || row['tnsnphm'] || row['articlename'] || row['article'] || row['tensanpham'] || row['productname'];
        let qty = Number(String(row['posquantity'] || row['sum'] || '0').replace(/,/g, ''));
        
        let storeID = extractSAP(st);
        let rawDate = String(row['calendarday'] || row['date'] || row['ngay'] || '').trim();
        
        if (rawDate && storeID) {
            if (!storeMonthlyDays.has(storeID)) storeMonthlyDays.set(storeID, new Set());
            storeMonthlyDays.get(storeID).add(rawDate); 

            // TRACK WEEKDAY vs WEEKEND
            if (!storeGroupDays.has(storeID)) {
                storeGroupDays.set(storeID, { weekdays: new Set(), weekends: new Set() });
            }
            let cDate = parseDateStrToTime(rawDate);
            let dayOfWeek = new Date(cDate).getDay();
            // Weekday: 1,2,3,4 (T2-T5) | Weekend: 5,6,0 (T6-CN)
            let isWknd = (dayOfWeek === 5 || dayOfWeek === 6 || dayOfWeek === 0);
            if (isWknd) storeGroupDays.get(storeID).weekends.add(rawDate);
            else storeGroupDays.get(storeID).weekdays.add(rawDate);

            if (!st || !pr || isNaN(qty)) return;

            let prodStd = normalizeProductName(pr);
            let key = `${storeID}_${prodStd.toLowerCase()}`;
            
            if (!monthlySales.has(key)) {
                monthlySales.set(key, { 
                    storeOrig: st, 
                    prodStd, 
                    totalQty: qty,
                    weekdayQty: isWknd ? 0 : qty,
                    weekendQty: isWknd ? qty : 0
                });
            } else {
                let data = monthlySales.get(key);
                data.totalQty += qty;
                if (isWknd) data.weekendQty += qty;
                else data.weekdayQty += qty;
            }
        }
    });

    const weeklySales = new Map();
    const storeWeeklyDays = new Map();
    const unmappedProducts = new Set(); // Thêm Tracking Lỗi Mapping

    datasets.weekly.forEach(row => {
        let st = row['sap'] || row['storecode'];
        let pr = row['tnsnphmwm'] || row['tnsnphm'] || row['articlename'] || row['article'] || row['tensanpham'] || row['productname'];
        
        // --- SCAN & CẢNH BÁO THIẾU MAPPING ---
        let pNameLower = String(pr).trim().toLowerCase();
        if (pNameLower && datasets.mapping_raw && datasets.mapping_raw.length > 0) {
            // Nếu phát hiện ra Name POS mới mà Chưa Có Trong Mapping File -> Cảnh Báo
            if (!mappingMap.has(pNameLower)) {
                unmappedProducts.add(String(pr).trim());
            }
        }

        let qty = Number(String(row['posquantity'] || row['sum'] || '0').replace(/,/g, ''));

        let storeID = extractSAP(st);
        let rawDate = String(row['calendarday'] || row['date'] || row['ngay'] || '').trim();
        
        if (rawDate && storeID) {
            if (!storeWeeklyDays.has(storeID)) storeWeeklyDays.set(storeID, new Set());
            storeWeeklyDays.get(storeID).add(rawDate);
        }

        if (!st || !pr || isNaN(qty)) return;

        let prodStd = normalizeProductName(pr);
        let key = `${storeID}_${prodStd.toLowerCase()}`;
        
        if (!weeklySales.has(key)) {
            weeklySales.set(key, qty);
        } else {
            weeklySales.set(key, weeklySales.get(key) + qty);
        }
    });


    // ----------- CẢNH BÁO MAPPING LÊN MÀN HÌNH CHÍNH -----------
    const warningDiv = document.getElementById('mapping-warning-div');
    if (warningDiv) {
        if (unmappedProducts.size > 0 && datasets.mapping_raw && datasets.mapping_raw.length > 0) {
            warningDiv.innerHTML = `<strong style="color: #ff9800; font-size: 1.1em;"><i class="fas fa-exclamation-triangle"></i> CẢNH BÁO MAPPING: TÌM THẤY ${unmappedProducts.size} SẢN PHẨM MỚI TRONG DOANH SỐ TUẦN!</strong><br>
            <span style="display:block; margin-top: 8px;">Hệ thống phát hiện các mã này đang phát sinh số lượng bán nhưng <b>CHƯA CÓ TRONG BẢNG MAPPING</b> hiên tại: <br>
            <i style="color: #fff; background: rgba(255,255,255,0.1); padding: 5px 8px; border-radius: 4px; display: inline-block; margin-top: 5px;">${Array.from(unmappedProducts).slice(0, 15).join(', ')}${unmappedProducts.size > 15 ? '...' : ''}</i></span>
            <small style="color: var(--text-muted); display: block; margin-top: 8px;">Web sẽ giữ nguyên Tên Gốc của Sản phẩm này vào Cột SOQ để tránh sót đơn. Tuy nhiên, xin vui lòng Cập Nhật File Mapping mới nhất nếu Tên đó bị sai mã nội bộ.</small>`;
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
            allItems.set(key, { storeID, storeOrig, bestName });
        }
    };

    // Đẩy tất cả các key từ các luồng vào list tổng (chỉ lấy cửa hàng thuộc Lịch Giao Hàng)
    let syncKeys = [...monthlySales.keys(), ...inventoryMap.keys(), ...inputMap.keys()];
    let hasScheduleUploaded = datasets.schedule && datasets.schedule.length > 0;
    
    syncKeys.forEach(k => {
        let parts = k.split('_');
        let storeID = parts[0];
        
        // Strict Filter Lịch Giao: Nếu có tải file Lịch lên, BẮT BUỘC mã cửa hàng phải có mặt trong validSAPs (vừa check ngày vừa check có list)
        if (hasScheduleUploaded && !validSAPs.has(storeID)) return; 
        
        let mData = monthlySales.get(k);
        let iData = inventoryMap.get(k);
        let inData = inputMap.get(k);

        let storeOrig = mData ? mData.storeOrig : (storeID);
        let rawProdStdName = mData ? mData.prodStd : (iData ? iData.prodOrig : (inData ? inData.prodOrig : parts[1]));

        registerKey(k, storeID, storeOrig, rawProdStdName);
    });

    // Block cảnh báo mapping đã được dời lên trên để chạy sớm hơn

    finalResults = [];
    tbody.innerHTML = '';

    allItems.forEach((data, key) => {
        // Chốt số ngày thực tế file doanh số bung qua THEO TỪNG CỬA HÀNG
        let mDaysCount = storeMonthlyDays.has(data.storeID) && storeMonthlyDays.get(data.storeID).size > 0 
                         ? storeMonthlyDays.get(data.storeID).size : 30;
        let wDaysCount = storeWeeklyDays.has(data.storeID) && storeWeeklyDays.get(data.storeID).size > 0 
                         ? storeWeeklyDays.get(data.storeID).size : 7;

        // Average Daily Sales
        let mDataExt = monthlySales.get(key);
        let mTotal = mDataExt ? mDataExt.totalQty : 0;
        let wTotal = weeklySales.has(key) ? weeklySales.get(key) : 0;

        // --- NEW: Phân tích T2-T5 vs T6-CN ---
        let weekdayQty = mDataExt ? mDataExt.weekdayQty : 0;
        let weekendQty = mDataExt ? mDataExt.weekendQty : 0;
        let storeGrps = storeGroupDays.get(data.storeID);
        let weekdayDaysCount = storeGrps ? storeGrps.weekdays.size : 0;
        let weekendDaysCount = storeGrps ? storeGrps.weekends.size : 0;

        let weekdayAds = weekdayDaysCount > 0 ? weekdayQty / weekdayDaysCount : 0;
        let weekendAds = weekendDaysCount > 0 ? weekendQty / weekendDaysCount : 0;
        let weekendGrowth = 0;
        let growthHtml = '-';
        if (weekdayAds > 0) {
            weekendGrowth = ((weekendAds - weekdayAds) / weekdayAds) * 100;
            if (weekendGrowth > 0) growthHtml = `<span style="color: var(--success)">+${weekendGrowth.toFixed(1)}%</span>`;
            else if (weekendGrowth < 0) growthHtml = `<span style="color: var(--danger)">${weekendGrowth.toFixed(1)}%</span>`;
            else growthHtml = `0%`;
        } else if (weekendAds > 0) {
            growthHtml = `<span style="color: var(--success)">Mới (CN)</span>`;
        }

        let mAds = mTotal / mDaysCount; 
        let wAds = wTotal / wDaysCount; 
        
        // Tính toán Xu Hướng Bán
        let trend = 0;
        let trendHtml = '-';
        let trendExport = '0%';
        if (mAds > 0) {
            trend = ((wAds - mAds) / mAds) * 100;
            trendExport = `${trend > 0 ? '+' : ''}${trend.toFixed(1)}%`;
            if (trend > 0) {
                trendHtml = `<span style="color: var(--success)">▲ ${trend.toFixed(1)}%</span>`;
            } else if (trend < 0) {
                trendHtml = `<span style="color: var(--danger)">▼ ${Math.abs(trend).toFixed(1)}%</span>`;
            } else {
                trendHtml = `<span>0%</span>`;
            }
        } else if (wAds > 0) {
            trendExport = '100% (New)';
            trendHtml = `<span style="color: var(--success)">▲ Mới bán</span>`;
        }

        // Nếu Weekly ko có thì dùng Monthly làm gốc để dự báo, xu hướng = N/A
        if (wTotal === 0 && mTotal > 0) {
            wAds = mAds; 
            trendHtml = `<span style="color: var(--text-muted)">N/A (Tuần 0)</span>`;
            trendExport = 'N/A';
        }
        
        // SỐ TRUNG BÌNH BÁN NGÀY HOÀN TOÀN DỰA VÀO THÁNG
        let forecastDay = mAds;
        
        // --- TÍNH TOÁN LEAD TIME TỔNG CỘNG ---
        // 1. Lead Time Arrival: Từ ngày Tồn kho (Inventory Date) đến ngày Giao hàng (Target Delivery)
        let invData = inventoryMap.get(key) || {inv: 0, disp: 0, latestDate: 0};
        let inputData = inputMap.get(key) || {qty: 0};

        let invDate = invData.latestDate || new Date().setHours(0,0,0,0);
        let leadTimeArrival = 0;
        if (targetTimestamp > 0) {
            leadTimeArrival = Math.max(0, (targetTimestamp - invDate) / (1000 * 60 * 60 * 24));
        }

        // 2. Coverage Leadtime: Khoảng cách giữa các đợt giao (lấy từ matrix lịch)
        let coverageLT = scheduleLeadtimeMap.has(data.storeID) ? scheduleLeadtimeMap.get(data.storeID) : extractLeadtimeFromFilename(scheduleFileName); 
        
        let totalLeadtime = leadTimeArrival + coverageLT;
        // THAY PHƯƠNG PHÁP TÍNH CỐ ĐỊNH BẰNG ĐỘNG THEO LOẠI NGÀY
        let totalDemand = calculatePeriodDemand(invDate, totalLeadtime, weekdayAds, weekendAds);

        // Phân loại Tier để nhồi thêm Tồn Kho Tối Thiểu (Safety Stock)
        let tierLevel = 0; 
        if (storeTierMap.has(data.storeID)) {
            let t = storeTierMap.get(data.storeID);
            if (t.includes('1') || t === 'T1' || t === 'TIER1' || t === 'TIER 1') {
                tierLevel = 1;
            } else if (t.includes('2') || t === 'T2' || t === 'TIER 2' || t.includes('3') || t === 'T3' || t === 'TIER 3') {
                tierLevel = 2; // Gộp Tier 2 và 3 xài chung rate
            }
        }
        
        let safetyStock = 0;
        if (forecastDay > 0) {
            if (tierLevel === 1) {
                safetyStock = isWeekendDelivery ? (weekendAds * 0.30) : (weekdayAds * 0.15); 
            } else if (tierLevel === 2) {
                safetyStock = isWeekendDelivery ? (weekendAds * 0.20) : (weekdayAds * 0.10); 
            }
            totalDemand += safetyStock;
        }

        // Sử dụng invData và inputData đã bóc tách ở đầu vòng lặp

        let penaltyApplied = 0;
        let shelfStock = invData.inv + invData.disp; 

        if (shelfStock > 0) {
            let disposalRatio = invData.disp / shelfStock;
            let isRTE = RTE_PRODUCTS.some(p => data.bestName.toLowerCase().includes(p.toLowerCase()));
            let threshold = isRTE ? 0.30 : 0.15;

            if (disposalRatio > threshold) {
                penaltyApplied = invData.disp; 
                totalDemand -= penaltyApplied;
            }
        }

        // CHỈ TRỪ HÀNG NHẬP (INPUT) NẾU NGÀY NHẬP > NGÀY TỒN KHO
        // Nếu ngày nhập <= ngày tồn, coi như hàng đã được đếm trong tồn kho rồi (số tồn trên kệ là tổng số tồn đang có)
        let actualInputQty = 0;
        if (inputData.latestDate > invData.latestDate) {
            actualInputQty = inputData.qty;
        }

        let soq = totalDemand - invData.inv - actualInputQty;
        soq = Math.max(Math.ceil(soq), 0);

        // HIỂN THỊ ĐẦY ĐỦ SOQ NẾU CÓ BẤT KỲ Ý NGHĨA KINH DOANH NÀO
        // Ẩn dòng có TẤT CẢ = 0
        if (soq === 0 && totalDemand === 0 && invData.inv === 0 && inputData.qty === 0 && invData.disp === 0) {
            return;
        }

        let storeNameStr = storeNamesMap.get(data.storeID) || data.storeOrig;

        finalResults.push({
            'Mã SAP (Store)': data.storeID,
            'Tên Cửa Hàng': storeNameStr,
            'Tên Sản Phẩm': data.bestName,
            'Trung Bình Bán/Ngày': forecastDay.toFixed(2),
            'Xu Hướng Bán (%)': trendExport,
            'ADS T2-T5': weekdayAds.toFixed(2),
            'ADS T6-CN': weekendAds.toFixed(2),
            'Tăng trưởng CT (%)': weekdayAds > 0 ? `${weekendGrowth.toFixed(1)}%` : (weekendAds > 0 ? 'New' : '0%'),
            'Leadtime': coverageLT,
            'Total Demand': (totalDemand + penaltyApplied).toFixed(2),
            'Tồn (Inv)': Number(invData.inv.toFixed(2)),
            'Nhập (Input)': Number(actualInputQty.toFixed(2)),
            'Giảm trừ (Penalty)': penaltyApplied > 0 ? `-${penaltyApplied.toFixed(2)}` : '0',
            'SOQ': soq
        });

        let tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${data.storeID}</td>
            <td>${storeNameStr}</td>
            <td>${data.bestName}</td>
            <td>${forecastDay.toFixed(2)}</td>
            <td><b>${trendHtml}</b></td>
            <td>${weekdayAds.toFixed(2)}</td>
            <td>${weekendAds.toFixed(2)}</td>
            <td><b>${growthHtml}</b></td>
            <td><span title="Coverage: ${coverageLT} ngày. (Tính cả Arrival +${leadTimeArrival.toFixed(1)} ngày thì Tổng Demand là ${(totalLeadtime).toFixed(1)} ngày bán)">${coverageLT}</span></td>
            <td>${(totalDemand + penaltyApplied).toFixed(2)}</td>
            <td class="warning">${Number(invData.inv.toFixed(2))}</td>
            <td class="highlight">${Number(actualInputQty.toFixed(2))}</td>
            <td style="color:${penaltyApplied > 0 ? 'var(--danger)' : ''}">${penaltyApplied > 0 ? `-${penaltyApplied.toFixed(2)}` : '0'}</td>
            <td class="highlight">${soq}</td>
        `;
        tbody.appendChild(tr);
    });

    if (finalResults.length === 0) {
        let monthlyKeys = datasets.monthly.length > 0 ? Object.keys(datasets.monthly[0]).join(', ') : 'No data';
        let invKeys = datasets.inventory.length > 0 ? Object.keys(datasets.inventory[0]).join(', ') : 'No data';
        let schedKeys = datasets.schedule.length > 0 ? Object.keys(datasets.schedule[0]).join(', ') : 'No data';
        
        tbody.innerHTML = `<tr><td colspan="9" style="text-align:left; color: var(--danger); padding: 2rem;">
            <strong>Không tìm thấy bất kỳ dữ liệu hợp lệ nào. (Tồn kho, hàng nhập và lịch giao không khớp ngàm dữ liệu, hoặc tất cả đều bằng 0).</strong><br/><br/>
            <div style="font-family: monospace; font-size:12px; color: var(--text-muted);">
                <strong>--- TRÌNH KIỂM TRA LỖI NỘI BỘ ---</strong><br/>
                - Schedule Headers: ${schedKeys}<br/>
                - Monthly Headers: ${monthlyKeys}<br/>
                - Inventory Headers: ${invKeys}<br/>
                - Lịch Giao Hàng quét được: ${validSAPs.size} mã hợp lệ (VD: ${Array.from(validSAPs).slice(0, 5).join(', ')})<br/>
                - Mapping File quét được: ${mappingMap.size} cặp quy đổi.<br/>
                - Master List đăng ký được: ${allItems.size} mã sản phẩm.
            </div>
            <p>Vui lòng chụp màn hình đoạn mã màu xám này và gửi lại để kỹ sư hoàn tất căn chỉnh file.</p>
        </td></tr>`;
    }

    resultsSection.style.display = 'block';
    if(finalResults.length > 0) btnExport.style.display = 'inline-block';
});

// Export to Excel (Bypass Security Block for local file:///)
btnExport.addEventListener('click', () => {
    if (finalResults.length === 0) return;
    const worksheet = XLSX.utils.json_to_sheet(finalResults);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "SOQ_Results");
    
    // Khử dấu tiếng Việt và ký tự lạ để tránh Browser chặn tải
    let safeName = String(scheduleFileName).normalize('NFD').replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9_\-]/g, "_");
    let exportName = `SOQ_Data_${safeName}.xlsx`;
                     
    // Custom File downloader to bypass 'Cần có quyền tải xuống' warning
    try {
        let wbout = XLSX.write(workbook, {bookType:'xlsx', type:'binary'});
        function s2ab(s) {
            let buf = new ArrayBuffer(s.length);
            let view = new Uint8Array(buf);
            for (let i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
            return buf;
        }
        let blob = new Blob([s2ab(wbout)], {type:"application/octet-stream"});
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
        alert("Lỗi tải file: Trình duyệt của bạn khóa quyền tải cục bộ. Hãy mở trang này bằng Chrome nhé!");
    }
});
