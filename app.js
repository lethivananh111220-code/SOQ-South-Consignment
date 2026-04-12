// --- CẤU HÌNH FIREBASE ---
// Bạn cần lấy thông tin này từ Firebase Console (https://console.firebase.google.com/)
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

// Khởi tạo Firebase nếu thư viện đã tải thành công
if (typeof firebase !== 'undefined') {
    firebase.initializeApp(firebaseConfig);
}

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
const inputUserName = document.getElementById('user-name');

function removeAccents(str) {
    if (!str) return '';
    return str.normalize('NFD')
              .replace(/[\u0300-\u036f]/g, '')
              .replace(/đ/g, 'd').replace(/Đ/g, 'D');
}

// Helper to normalize column names
function normalizeKey(key) {
    if (!key) return '';
    let s = removeAccents(key.toString().toLowerCase());
    return s.replace(/[^a-z0-9]/g, '');
}

// Hàm trích xuất tự động bỏ qua các tiêu đề báo cáo rác ở file hệ thống (Excel report info)
function extractJsonDataCleanly(worksheet) {
    let rawArr = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, dateNF: 'yyyy-mm-dd hh:mm:ss' });
    if (!rawArr || rawArr.length === 0) return [];

    let headerIdx = 0;
    // Tìm dòng header thực sự (Thường có chứa các chữ khóa nhận diện và > 3 cột dữ liệu)
    for (let i = 0; i < Math.min(20, rawArr.length); i++) {
        let r = rawArr[i];
        if (!r) continue;
        let validCols = r.filter(c => typeof c === 'string' && c.trim() !== '');
        if (validCols.length >= 2 && r.some(c => typeof c === 'string' && (c.toUpperCase().includes('SAP') || c.toUpperCase().includes('STORE') || c.toUpperCase().includes('NICKNAME') || c.toUpperCase().includes('TÊN') || c.toUpperCase().includes('ARTICLE') || c.toUpperCase().includes('PRODUCT')))) {
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
    
    // FALLBACK: Nếu headers toàn là số (Excel Serial Dates) -> Có thể đây là file matrix không có label.
    let numericHeadersCount = headersRaw.filter(h => typeof h === 'number' && h > 40000).length;
    // Tăng cường kiểm tra cả headersPrefix nếu có
    if (headersPrefix.length > 0) numericHeadersCount += headersPrefix.filter(h => typeof h === 'number' && h > 40000).length;

    if (numericHeadersCount > 5) {
        // Đây là dạng file Lịch Matrix. Ép các cột cố định (0: Type, 1: SAP, 4: Name)
        // Lưu ý: Nếu có prefix, headers[1] có thể là "v_sap". Ta rà soát index.
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

        // Bỏ qua dòng Total (Dòng tổng cộng của SAP)
        if (row.some(cell => String(cell).toUpperCase().includes('RESULT') || String(cell).toUpperCase() === 'TOTAL')) continue;

        let obj = {};
        let hasData = false;
        for (let j = 0; j < headers.length; j++) {
            if (row[j] !== undefined && row[j] !== null && String(row[j]).trim() !== '') {
                obj[headers[j]] = row[j]; // Composite

                // Khôi phục việc đọc các cột đơn giản (sap, date...) để không bị hư tên do prefix chặn.
                // Ngăn chặn riêng biệt lỗi trượt/chồng lắp lịch các ngày trong tuần (đã xử lý ở bước trước).
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
    let s = String(val).trim().split(' ')[0]; // Bỏ time nếu có

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

// Tính tổng nhu cầu động dựa trên loại ngày (thứ 2-5 hoặc thứ 6-CN)
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
    statusEl.textContent = `Đang đọc ${file.name}...`;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        let firstSheetName = workbook.SheetNames[0];
        
        // Theo yêu cầu: Lấy dữ liệu từ sheet "Summary by Products" cho file Input
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
            statusEl.textContent = `Đã tải & lưu trữ: ${file.name} (${arr.length} dòng)`;
            statusEl.classList.add('success');
            checkReady();
            return;
        }

        try {
            if (type === 'monthly') {
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

            if (type === 'monthly' || type === 'weekly') {
                saveToDB(type, datasets[type]);
                statusEl.textContent = `Đã tải & lưu trữ: ${file.name} (${datasets[type].length} dòng)`;
            } else {
                if (type === 'schedule') {
                    scheduleFileName = file.name.replace(/\.[^/.]+$/, "");
                }
                statusEl.textContent = `Đã tải: ${file.name} (${datasets[type].length} dòng)`;
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
    } catch (e) { console.error('Lỗi lưu cache', e); }
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
        if (Array.isArray(raw)) return raw; // Cache cổ điển 

        if (raw.timestamp && raw.data) {
            let dDate = new Date(raw.timestamp);
            let nDate = new Date();

            if (key === 'monthly') {
                if (dDate.getMonth() !== nDate.getMonth() || dDate.getFullYear() !== nDate.getFullYear()) {
                    // Nếu là file tháng, chỉ xóa nếu quá 2 tháng (để user dùng được tháng trước + tháng này)
                    let monthDiff = (nDate.getFullYear() - dDate.getFullYear()) * 12 + (nDate.getMonth() - dDate.getMonth());
                    if (monthDiff > 2) {
                        await deleteFromDB(key);
                        return { invalidated: true, reason: "dữ liệu quá cũ (>2 tháng)" };
                    }
                }
            } else if (key === 'weekly') {
                // Hạn sử dụng của Doanh số tuần: Từ ngày tải lên (dDate) kéo dài đến Thứ 3 của tuần kế tiếp
                let expirationTime = getWeekStart(dDate) + 8 * 86400000; 
                if (nDate.getTime() >= expirationTime) {
                    await deleteFromDB(key);
                    return { invalidated: true, reason: "sang thứ 3 tuần mới" };
                }
            } else if (key === 'schedule' || key === 'inventory' || key === 'input' || key.startsWith('soq_latest_')) {
                let vnDateStr = (date) => {
                    return date.toLocaleString('en-US', { timeZone: 'Asia/Ho_Chi_Minh', year: 'numeric', month: '2-digit', day: '2-digit' });
                };
                if (vnDateStr(dDate) !== vnDateStr(nDate)) {
                    await deleteFromDB(key);
                    return { invalidated: true, reason: "qua ngày mới" };
                }
            }
            return raw.data;
        }
        return raw;
    } catch (e) { return null; }
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

    if (cMonthly) {
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
            if (el) { el.innerHTML = `<span style="color: #ff9800; font-weight: bold;">Lưu ý: Sang Thứ 3 tuần mới. Vui lòng tải số báo cáo tuần mới!</span>`; el.classList.remove('success'); }
        } else if (cWeekly.length > 0) {
            datasets.weekly = cWeekly;
            let el = document.getElementById('status-weekly');
            if (el) { el.textContent = `Đã dùng bản lưu trước (${cWeekly.length} dòng)`; el.classList.add('success'); }
        }
    }
    if (cMapping && cMapping.length > 0) {
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
    // Ưu tiên: Nếu là chuỗi số đứng độc lập (có thể có chữ bao quanh bởi dấu cách) -> Lấy số
    let m = s.match(/\b\d+\b/);
    if (m) return Number(m[0]).toString();
    
    return s.toLowerCase();
}

btnCalculate.addEventListener('click', () => {
    try {
        tbody.innerHTML = "";
        finalResults = [];
        resultsSection.style.display = 'none';
        // --- KIỂM TRA DỮ LIỆU ĐẦU VÀO ---
        if (!datasets.schedule || datasets.schedule.length === 0) {
            alert("Vui lòng tải file Lịch giao hàng (Schedule)!");
            return;
        }
        if (!datasets.inventory || datasets.inventory.length === 0) {
            alert("Vui lòng tải file Tồn kho (Merchandiser)!");
            return;
        }
        if (!datasets.monthly || datasets.monthly.length === 0) {
            alert("Vui lòng tải file Doanh số tháng (Monthly Sales)!");
            return;
        }
        // Tip: Mapping là bắt buộc nếu muốn dùng tính năng lọc mẫu (strict mapping)
        if (!datasets.mapping_raw || datasets.mapping_raw.length === 0) {
            alert("Lưu ý: Bạn chưa tải file Mapping. Hệ thống sẽ lấy tên gốc từ file doanh số.");
        }

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

            // Cuối tuần: Thứ 7 (6), Chủ nhật (0)
            if (finalWkday === 6 || finalWkday === 0) {
                isWeekendDelivery = true;
            }
        }

        // ----------- 1. Map Rules (WM Name -> ODA Name) -----------
        const mappingMap = new Map();
        const standardNamesSet = new Set(); // Lưu danh sách Tên ODA chuẩn
        const unmappedProducts = new Set(); // Tracking sản phẩm chưa được mapping
        const reverseMappingKeys = new Set(); // Dùng để kiểm tra sản phẩm lạ
        const productCategoryMap = new Map(); // Lưu nhóm hàng mảng Penalty

        if (datasets.mapping_raw && datasets.mapping_raw.length > 0) {
            let headerRow = datasets.mapping_raw[0] || [];
            let iOda = 1, iWm = 2, iCat = 3;

            // Nhận diện tự động cột bằng Tên Header
            for (let c = 0; c < headerRow.length; c++) {
                let h = String(headerRow[c]).toUpperCase();
                if (h.includes('ODA')) iOda = c;
                else if (h.includes('WM')) iWm = c;
                else if (h.includes('NHÓM')) iCat = c;
            }

            // Bắt đầu đọc từ dòng số 2 (Bỏ qua Header)
            for (let i = 1; i < datasets.mapping_raw.length; i++) {
                let r = datasets.mapping_raw[i];
                if (!r || !Array.isArray(r)) continue;

                let odaName = r[iOda] ? String(r[iOda]).trim() : '';
                let wmName = r[iWm] ? String(r[iWm]).trim().toLowerCase() : '';
                let category = r[iCat] ? String(r[iCat]).trim().toUpperCase() : '';

                // Nếu ko có Header (file trống trơn 2 cột), chạy fallback truyền thống
                if (!odaName && !wmName && r.length >= 2) {
                    wmName = r[0] ? String(r[0]).trim().toLowerCase() : '';
                    odaName = r[1] ? String(r[1]).trim() : '';
                }

                if (wmName && odaName && wmName !== 'tên sản phẩm wm') {
                    mappingMap.set(wmName, odaName);
                    standardNamesSet.add(odaName.trim().toLowerCase());
                    reverseMappingKeys.add(wmName);

                    if (category && category !== 'NHÓM HÀNG') {
                        productCategoryMap.set(odaName.trim().toLowerCase(), category);
                    }
                }
            }
        }

        const normalizeProductName = (name) => {
            let n = String(name).trim().toLowerCase();
            // 1. Nếu là Tên WM -> Trả về Tên ODA chuẩn
            if (mappingMap.has(n)) return String(mappingMap.get(n)).trim();
            // 2. Nếu chính nó đã là Tên ODA chuẩn -> Trả về chính nó
            if (standardNamesSet.has(n)) return String(name).trim();

            // Nếu có nạp file mapping mà không thấy mã này -> Coi như không hợp lệ (Trả về null để lọc bỏ)
            if (datasets.mapping_raw && datasets.mapping_raw.length > 0) return null;
            return String(name).trim(); // Fallback nếu chưa nạp mapping
        }

        // --- 2. Schedule Filter & Store Names ---
        const validSAPs = new Set();
        const storeNamesMap = new Map();
        const storeAliasesMap = new Map(); // ID -> Set of normalized names/nicknames
        const scheduleLeadtimeMap = new Map();
        const storeTierMap = new Map();

        if (datasets.schedule && datasets.schedule.length > 0) {
            datasets.schedule.forEach(row => {
                let store = row['sap'] || row['storekey'] || row['storecode'] || row['makho'] || row['mach'] || row['mãkháchhàng'] || row['mãcửahàng'] || row['nickname'] || row['storename'] || row['store'];
                if (!store) return;

                let storeID = extractSAP(store);
                let hinhThuc = String(row['hinhthuc'] || row['Hình thức'] || row['type'] || '').toUpperCase();

                let dynamicLT = 0;
                const getWeekdayIdx = getWeekdayIdxGlobal;

                if (targetDateStr !== "") {
                    let hasDelivery = false;
                    let isTargetWeekday = getWeekdayIdx(targetDateStr) !== -1;
                    let impliedWeekdayIdx = new Date(targetTimestamp).getDay();
                    let currentTargetNum = isTargetWeekday ? getWeekdayIdx(targetDateStr) : new Date(targetTimestamp).getDate();

                    let possibleNextDeliveryTimestamps = [];

                    // Khởi tạo biến kiểm tra Chức năng (Function) của Store
                    let isMer = String(row['function'] || row['Function'] || row['chức năng'] || row['loại'] || '').trim().toLowerCase() === 'mer';

                    for (const [key, val] of Object.entries(row)) {
                        let k = String(key).trim();
                        let match = false;
                        let headerTs = 0;

                        let headerWeekdayIdx = getWeekdayIdx(k);

                        // Nếu Header file Lịch là THỨ (VD: Friday, T2)
                        if (headerWeekdayIdx !== -1) {
                            if (isTargetWeekday) {
                                match = (headerWeekdayIdx === currentTargetNum);
                            } else if (impliedWeekdayIdx !== -1) {
                                match = (headerWeekdayIdx === impliedWeekdayIdx);
                            }
                        } else {
                        // Xử lý Header phức hợp (vd: 01-Thg4_Wednesday) hoặc Header đơn thuần
                        let kClean = k.toLowerCase();
                        
                        // Lấy số ngày của mục tiêu (VD: 1 hoặc 01)
                        let tNum = new Date(targetTimestamp).getDate().toString();
                        let tPadded = tNum.padStart(2, '0');

                        // 1. So khớp Số ngày trực tiếp: "01", "1", "1-", "01-"
                        let dateMatch = kClean.startsWith(tNum + '-') || kClean.startsWith(tPadded + '-') || 
                                       kClean.includes('_' + tNum + '-') || kClean.includes('_' + tPadded + '-');
                        
                        // 2. So khớp Số ngày viết liền (Ví dụ: 01thg4)
                        if (!dateMatch) {
                            let m = kClean.match(/^(\d{1,2})/);
                            if (m && (m[1] === tNum || m[1] === tPadded)) dateMatch = true;
                        }

                        // 3. So khớp Serial Date nếu có trong Key
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

                        // NEW: Trích xuất Timestamp cho tất cả các cột nếu có định dạng ngày (vd: 01-thg4)
                        if (headerTs === 0) {
                            // Thử bóc tách ngày/tháng từ chuỗi "01-thg4"
                            let mDate = kClean.match(/^(\d{1,2})[^\d]+(\d{1,2})/);
                            if (mDate) {
                                let dd = parseInt(mDate[1]);
                                let mm = parseInt(mDate[2]) - 1;
                                let yyyy = new Date(targetTimestamp).getFullYear();
                                let dTemp = new Date(yyyy, mm, dd);
                                // Nếu ngày quá xa mục tiêu (vd: tháng 12 so với tháng 1), lùi/tiến năm
                                headerTs = dTemp.getTime();
                            } else {
                                // Thử bóc tách ngày đơn thuần (vd: 01) -> Giả định cùng tháng/năm với target
                                let mDay = kClean.match(/^(\d{1,2})/);
                                if (mDay) {
                                    let dd = parseInt(mDay[1]);
                                    let tDate = new Date(targetTimestamp);
                                    let dTemp = new Date(tDate.getFullYear(), tDate.getMonth(), dd);
                                    // Xử lý rollover tháng nếu cần (vd: target là 31/3, header là 1)
                                    if (dd < tDate.getDate() - 15) dTemp.setMonth(dTemp.getMonth() + 1);
                                    if (dd > tDate.getDate() + 15) dTemp.setMonth(dTemp.getMonth() - 1);
                                    headerTs = dTemp.getTime();
                                }
                            }
                        }
                        
                        // ƯU TIÊN: Nếu Header chứa thông tin NGÀY CỐ ĐỊNH, nó sẽ ghi đè việc so khớp THỨ chung chung
                        if (dateMatch || serialMatch) {
                            match = true;
                        } else if (!isTargetWeekday && headerWeekdayIdx === -1) {
                            // Fallback nếu headers quá đơn giản (chỉ "1", "2")
                            match = (k === tNum || k === tPadded || k.startsWith(tNum + '/') || k.startsWith(tPadded + '/'));
                        }
                        }

                        let v = String(val).trim().toLowerCase().replace(/\s+/g, '');
                        let isDeliveryFound = false;

                        if (v && v !== '0' && v !== 'false' && v !== 'off' && !v.includes('nghỉ')) {
                            if (isMer) {
                                // Rule Function Mer: Chịu trách nhiệm giao dịch nếu có mặt NVCH
                                // Từ chối những CH đi thăm (chỉ ghi "NVCH"). Phải ghi "Shipper+NVCH" hoặc có dấu "+"
                                if ((v.includes('shipper') && v.includes('nvch')) || (v.includes('nvch') && v.includes('+')) || v.includes('giao')) {
                                    isDeliveryFound = true;
                                } else if (v === 'x' || v === 'yes' || v === 'true') {
                                    isDeliveryFound = true; // Fallback an toàn
                                }
                            } else {
                                // Nếu không phải Function Mer (hoặc không có cột Function), mọi tín hiệu như Shipper, X đều tính
                                isDeliveryFound = true;
                            }
                        }

                        if (isDeliveryFound) {
                            if (match) {
                                hasDelivery = true;
                            }
                            // Theo dõi tất cả các mốc có giao hàng tiếp theo (Dạng Timestamp)
                            if (headerTs > 0) {
                                possibleNextDeliveryTimestamps.push(headerTs);
                            } else if (headerWeekdayIdx !== -1) {
                                // Nếu là THỨ, quy đổi sang timestamp tương ứng trong tuần đó/tuần sau
                                let dTarget = new Date(targetTimestamp);
                                let diff = (headerWeekdayIdx - dTarget.getDay() + 7) % 7;
                                let dNext = new Date(dTarget);
                                dNext.setDate(dNext.getDate() + diff);
                                possibleNextDeliveryTimestamps.push(dNext.getTime());
                            }
                        }
                    }

                    // Nếu không có lịch giao -> Bỏ qua
                    if (!hasDelivery) return;

                    // --- TÍNH TOÁN LEADTIME ĐỘNG TỪ MA TRẬN LỊCH GIAO HÀNG (Dạng Timestamp) ---
                    let futureDates = possibleNextDeliveryTimestamps.filter(t => t > targetTimestamp + 3600000); // Cách ít nhất 1h
                    if (futureDates.length > 0) {
                        let nextTS = Math.min(...futureDates);
                        dynamicLT = Math.round((nextTS - targetTimestamp) / 86400000);
                    }
                }

                // Mặc định: Chấp nhận TẤT CẢ các mã cửa hàng miễn là có tên trong file Lịch Giao Hàng
                if (storeID) {
                    validSAPs.add(storeID);

                    let sName = row['tencuahang'] || row['tncahng'] || row['storename'] || row['store'] || row['nickname'] || ''; 
                    let nickname = row['nickname'] || '';

                    if (sName) storeNamesMap.set(storeID, String(sName).trim());

                    // Đăng ký Alias
                    if (!storeAliasesMap.has(storeID)) storeAliasesMap.set(storeID, new Set());
                    if (sName) storeAliasesMap.get(storeID).add(normalizeKey(sName));
                    if (nickname) storeAliasesMap.get(storeID).add(normalizeKey(nickname));
                    storeAliasesMap.get(storeID).add(normalizeKey(storeID));

                    // LƯU CỘT TIER
                    let tierVal = String(row['tier'] || row['Tier'] || row['cấpđộ'] || row['phânloại'] || '').trim().toUpperCase();
                    if (tierVal && tierVal !== 'UNDEFINED') storeTierMap.set(storeID, tierVal);

                    if (dynamicLT > 0) {
                        scheduleLeadtimeMap.set(storeID, dynamicLT); 
                    } else {
                        let lt = Number(row['leadtime'] || row['Leadtime'] || row['chu kỳ'] || row['chukỳ'] || 0);
                        if (lt > 0) scheduleLeadtimeMap.set(storeID, lt);
                    }
                }
            });
        }

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

        // TẠO BẢN ĐỒ NGƯỢC SỚM: Tên Store (Chuẩn hóa) / Nickname -> Mã SAP để xử lý Tồn Kho & Nhập
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
        // Build lần 1: Lấy dữ liệu Alias từ file Lịch giao hàng (Schedule) làm gốc
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
                            if (alias.length > 5 || nKey.length > 5) { // Tránh nhầm lẫn chữ tắt quá ngắn
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

        // --- BƯỚC 0: TÌM NGÀY LỚN NHẤT CỦA TỪNG STORE LÀM MỐC (T) ---
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
                let rawDate = row['date'] || row['Date'] || row['ngay'] || row['ngày'] || 0;
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

                let rawDate = row['date'] || row['Date'] || row['ngay'] || row['ngày'] || 0;
                let cDate = parseDateStrToTime(rawDate);

                let T = storeMasterDateMap.get(storeID);
                if (!T || cDate > T) return;

                let prodStd = normalizeProductName(prod);
                if (!prodStd) {
                    unmappedProducts.add(String(prod).trim());
                    return;
                }
                let key = `${storeID}_${prodStd.toLowerCase()}`;

                // ... (Quy đổi kg)
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
                    // Lưu dữ liệu của ngày gần T nhất
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
        const actualODA_Names = new Map(); // Lưu Tên ODA chuẩn nhất từ file vận hành

        if (datasets.input && datasets.input.length > 0) {
            datasets.input.forEach(row => {
                let prod = row['productnameprimarylanguage'] || row['productname'] || row['product'] || row['tensanphamwm'] || row['tensanpham'] || row['articlename'] || row['article'];
                let status = String(row['orderstatus'] || row['status'] || row['trangthai'] || '').toLowerCase();
                
                if (!prod) return;
                // Lọc bỏ hàng Hủy / Đã hoàn (Chỉ lấy Completed)
                if (status && (status.includes('cancel') || status.includes('hủy') || status.includes('reject'))) return;

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

                // Trích xuất ngày giao hàng/nhập hàng
                let rawDate = row['orderdate'] || row['Order date'] || row['completeddate'] || row['Completed date'] || row['date'] || row['ngaydathang'] || row['ngay'] || row['ngaytao'] || row['createddate'] || 0;
                let cOrderDate = parseDateStrToTime(rawDate);
                let cDeliveryDate = cOrderDate > 0 ? cOrderDate + 86400000 : 0; // Cộng thêm 1 ngày giao

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

        // Hàm lấy lại Tên Chuẩn nhất (Ưu tiên ODA thật > Mapping > Raw)
        const getBestAvailableName = (mappedName) => {
            if (!mappedName) return '';
            let k = String(mappedName).toLowerCase();
            return actualODA_Names.has(k) ? actualODA_Names.get(k) : mappedName;
        }

        // ----------- 5. Sales Data (Flat Transaction Aggregation) -----------
        // Trong file thực tế: Dữ liệu doanh số bán nằm từng dòng, cột "POS Quantity"
        const monthlySales = new Map();
        const storeMonthlyDays = new Map(); // All days
        const storeGroupDays = new Map();  // storeID -> { weekdays: Set, weekends: Set }

        const processMonthlyData = (dataArr) => {
            if (!dataArr || dataArr.length === 0) return;
            dataArr.forEach(row => {
                let st = row['sap'] || row['storecode'] || row['sapcode'] || row['store'] || row['nickname'] || row['storename'] || row['tencuahang'];
                let pr = row['tnsnphmwm'] || row['tensanphamwm'] || row['tnsnphm'] || row['articlename'] || row['article'] || row['tensanpham'] || row['productname'];
                let qty = Number(String(row['posquantity'] || row['quantity'] || row['soluong'] || row['sum'] || '0').replace(/,/g, ''));
                if (pr && String(pr).toLowerCase().includes('retail kg')) qty /= 1000;

                let storeID = extractSAP(st);
                
                // Hỗ trợ Fallback Lookup cho Monthly Sales y chang Weekly
                if (storeID && isNaN(parseInt(storeID))) {
                    let lookedUp = reverseStoreNamesMap.get(normalizeKey(st));
                    if (lookedUp) storeID = lookedUp;
                }

                let rawDate = String(row['calendarday'] || row['date'] || row['ngay'] || '').trim();

                // Đăng ký Tên/Nickname từ file Doanh số (ODA)
                if (storeID) {
                    let sName = row['storename'] || row['store'] || row['têncửahàng'] || '';
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
        };

        processMonthlyData(datasets.monthly);
        
        // Build lần 2: Bổ sung thêm Alias nếu file Doanh Thu Tháng có ghi nhận tên/nickname mới
        buildReverseMap();

        const weeklySales = new Map();
        const storeWeeklyDays = new Map();
        const storeWeeklyGroupDays = new Map();
        if (datasets.weekly && datasets.weekly.length > 0) {
            datasets.weekly.forEach(row => {
                // Kiểm tra xem đây là file TRANSACTION (phẳng) hay MATRIX (ngang)
                let st = row['sap'] || row['storecode'] || row['nickname'] || row['storename'] || row['store'] || row['mach'] || row['tencuahang'];
                let pr = row['tnsnphmwm'] || row['tensanphamwm'] || row['tnsnphm'] || row['articlename'] || row['article'] || row['tensanpham'] || row['productname'];
                
                if (!pr) return;
                let prodStd = normalizeProductName(pr);
                if (!prodStd) return;

                if (st) {
                    // --- DẠNG FILE PHẲNG (TRANSACTION) ---
                    let storeID = extractSAP(st);
                    
                    // Fallback cực mạnh cho ODA: Nếu ô Name/Nickname không chứa Mã SAP dạng số, ta sẽ lookup từ thư viện!
                    if (storeID && isNaN(parseInt(storeID))) {
                        let lookedUp = reverseStoreNamesMap.get(normalizeKey(st));
                        if (lookedUp) storeID = lookedUp;
                    }

                    let qty = Number(String(row['posquantity'] || row['sum'] || '0').replace(/,/g, ''));
                    if (pr && String(pr).toLowerCase().includes('retail kg')) qty /= 1000;

                    let rawDate = String(row['calendarday'] || row['date'] || row['ngay'] || '').trim();
                    let isWknd = false;

                    if (rawDate && storeID) {
                        if (!storeWeeklyDays.has(storeID)) storeWeeklyDays.set(storeID, new Set());
                        storeWeeklyDays.get(storeID).add(rawDate);

                        if (!storeWeeklyGroupDays.has(storeID)) {
                            storeWeeklyGroupDays.set(storeID, { weekdays: new Set(), weekends: new Set() });
                        }
                        let cDate = parseDateStrToTime(rawDate);
                        let dayOfWeek = new Date(cDate).getDay();
                        isWknd = (dayOfWeek === 6 || dayOfWeek === 0);
                        if (isWknd) storeWeeklyGroupDays.get(storeID).weekends.add(rawDate);
                        else storeWeeklyGroupDays.get(storeID).weekdays.add(rawDate);
                    }

                    if (isNaN(qty)) return;
                    let key = `${storeID}_${prodStd.toLowerCase()}`;

                    if (!weeklySales.has(key)) {
                        weeklySales.set(key, { totalQty: qty, weekdayQty: isWknd ? 0 : qty, weekendQty: isWknd ? qty : 0 });
                    } else {
                        let data = weeklySales.get(key);
                        data.totalQty += qty;
                        if (isWknd) data.weekendQty += qty;
                        else data.weekdayQty += qty;
                    }
                } else {
                    // --- DẠNG FILE MA TRẬN (MATRIX - Tên cửa hàng ở tiêu đề cột) ---
                    // Duyệt từng cột của dòng này
                    Object.entries(row).forEach(([colKey, qtyVal]) => {
                        let cKey = String(colKey).trim();
                        if (!cKey) return;

                        // ƯU TIÊN 1: Tìm xem trong Header có chứa Mã SAP (4-5 số) không?
                        let sID = "";
                        let sapMatch = cKey.match(/(\d{4,5})/);
                        if (sapMatch && reverseStoreNamesMap.has(normalizeKey(sapMatch[1]))) {
                            sID = reverseStoreNamesMap.get(normalizeKey(sapMatch[1]));
                        } else {
                            // ƯU TIÊN 2: Tìm theo Tên/Nickname đã normalize
                            sID = reverseStoreNamesMap.get(normalizeKey(cKey));
                        }

                        if (sID) {
                            let qty = Number(String(qtyVal || '0').replace(/,/g, ''));
                            if (pr && String(pr).toLowerCase().includes('retail kg')) qty /= 1000;
                            if (isNaN(qty) || qty === 0) return;

                            let key = `${sID}_${prodStd.toLowerCase()}`;
                            // Với file Matrix không có ngày, ta mặc định chia đều tỉ lệ 5/2 (5 ngày thường, 2 ngày cuối tuần)
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
                            
                            // Giả lập số ngày (5 ngày thường, 2 cuối tuần) để denominator > 0
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


        // ----------- CẢNH BÁO MAPPING LÊN MÀN HÌNH CHÍNH -----------
        const warningDiv = document.getElementById('mapping-warning-div');
        if (warningDiv) {
            if (unmappedProducts.size > 0 && datasets.mapping_raw && datasets.mapping_raw.length > 0) {
                warningDiv.innerHTML = `<strong style="color: #ff9800; font-size: 1.1em;"><i class="fas fa-exclamation-triangle"></i> Cập nhập thêm sản phẩm: TÌM THẤY ${unmappedProducts.size} SẢN PHẨM MỚI TRONG DOANH SỐ TUẦN!</strong><br>
            <span style="display:block; margin-top: 8px;">Dưới đây là các mã <b>CHƯA ĐƯỢC GHI NHẬN</b> trong Mapping và đã bị tạm ẩn khỏi bảng SOQ: <br>
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
                // Đã bổ sung prodStd để hàm lọc nhóm hàng có thể map chính xác
                allItems.set(key, { storeID, storeOrig, bestName, prodStd: String(rawProdStdName || '') });
            }
        };

        // 2026-03-31: Đảm bảo tất cả store trong lịch phải được xuất hiện kể cả khi chưa có số bán/tồn
        let syncKeysSet = new Set([...monthlySales.keys(), ...inventoryMap.keys(), ...inputMap.keys()]);
        let hasScheduleUploaded = datasets.schedule && datasets.schedule.length > 0;

        if (hasScheduleUploaded) {
            // Lấy thêm các tổ hợp từ mapping hoặc các dữ liệu khác nếu store đó có trong schedule
            // Duyệt qua mapping hoặc toàn bộ danh sách sản phẩm đã từng thấy
            let anyProdStandards = new Set([...standardNamesSet]);
            // Nếu chưa có mapping, lấy từ Sales/Inventory
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
            let wDataExt = weeklySales.get(key);
            let wTotal = wDataExt ? wDataExt.totalQty : 0;
            let wWeekdayQty = wDataExt ? wDataExt.weekdayQty : 0;
            let wWeekendQty = wDataExt ? wDataExt.weekendQty : 0;

            let wStoreGrps = storeWeeklyGroupDays.get(data.storeID);
            let wWeekdayDaysCount = wStoreGrps ? wStoreGrps.weekdays.size : 0;
            let wWeekendDaysCount = wStoreGrps ? wStoreGrps.weekends.size : 0;

            let wWeekdayAds = wWeekdayDaysCount > 0 ? wWeekdayQty / wWeekdayDaysCount : 0;
            let wWeekendAds = wWeekendDaysCount > 0 ? wWeekendQty / wWeekendDaysCount : 0;

            // --- NEW: Phân tích T2-T5 vs T6-CN ---
            let weekdayQty = mDataExt ? mDataExt.weekdayQty : 0;
            let weekendQty = mDataExt ? mDataExt.weekendQty : 0;
            let storeGrps = storeGroupDays.get(data.storeID);
            let weekdayDaysCount = storeGrps ? storeGrps.weekdays.size : 0;
            let weekendDaysCount = storeGrps ? storeGrps.weekends.size : 0;


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
                    trendHtml = `<span style="color: var(--success)">▲ ${trend.toFixed(1)}%</span>`;
                } else if (trend < 0) {
                    trendHtml = `<span style="color: var(--danger)">▼ ${Math.abs(trend).toFixed(1)}%</span>`;
                } else {
                    trendHtml = `<span>0%</span>`;
                }
            } else if (wAds > 0) {
                trendExport = '100% (New)';
                trendHtml = `<span style="color: var(--success)">▲ Mới bán</span>`;
                trendFactor = 1; // Mặc định 1 cho hàng mới
            }

            // Nếu Weekly ko có thì dùng Monthly làm gốc để dự báo, xu hướng = N/A
            if (wTotal === 0 && mTotal > 0) {
                wAds = mAds;
                trendHtml = `<span style="color: var(--text-muted)">N/A (Tuần 0)</span>`;
                trendExport = 'N/A';
                trendFactor = 1;
            }

            // SỐ TRUNG BÌNH BÁN NGÀY HOÀN TOÀN DỰA VÀO THÁNG
            let forecastDay = mAds;

            // --- TÍNH TOÁN LEAD TIME TỔNG CỘNG ---
            // 1. Lead Time Arrival: Từ ngày T (Master Date) đến ngày Giao hàng (Target Delivery)
            let T = storeMasterDateMap.get(data.storeID) || 0;
            let invData = inventoryMap.get(key) || { currentInv: 0, currentDisp: 0, prevInv: 0, prevInvDate: 0 };
            let inputData = inputMap.get(key) || { currentInput: 0, prevInput: 0, prevInputDate: 0 };

            let invDate = T > 0 ? T : new Date().setHours(0, 0, 0, 0);
            let leadTimeArrival = 0;
            if (targetTimestamp > 0) {
                leadTimeArrival = Math.max(0, (targetTimestamp - invDate) / (1000 * 60 * 60 * 24));
            }

            // 2. Coverage Leadtime: Khoảng cách giữa các đợt giao (lấy từ matrix lịch)
            let coverageLT = scheduleLeadtimeMap.has(data.storeID) ? scheduleLeadtimeMap.get(data.storeID) : extractLeadtimeFromFilename(scheduleFileName);

            let totalLeadtime = leadTimeArrival + coverageLT;
            
            let basePeriodDemand = calculatePeriodDemand(invDate, totalLeadtime, weekdayAds, weekendAds);
            
            // Tách Demand dự kiến lúc chờ hàng (tránh âm kho dồn vào SOQ gây overstock)
            let leadTimeDemandBase = calculatePeriodDemand(invDate, leadTimeArrival, weekdayAds, weekendAds);
            let demandLeadTime = leadTimeDemandBase;

            // Demand kỳ bán SOQ (Chỉ tính Coverage)
            let coverageStartDate = invDate + (leadTimeArrival * 24 * 60 * 60 * 1000);
            let coverageDemandBase = calculatePeriodDemand(coverageStartDate, coverageLT, weekdayAds, weekendAds);
            let totalDemand = coverageDemandBase;

            // --- NEW: Tăng trưởng theo Leadtime (Đối chiếu Weekly vs Monthly trên từng Thứ) ---
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
                    safetyStock = isWeekendDelivery ? (weekendAds * coverageLT * 0.30) : (weekdayAds * coverageLT * 0.15);
                } else if (tierLevel === 2) {
                    safetyStock = isWeekendDelivery ? (weekendAds * coverageLT * 0.20) : (weekdayAds * coverageLT * 0.10);
                }
                totalDemand += safetyStock;
            }

            // Sử dụng mốc T để chuẩn hóa Tồn / Nhập đồng bộ (Khởi tạo ở đầu vòng lặp)
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

            let invTooltip = `Tồn kho lúc T (${strT}): [ ${finalInv.toFixed(2)} ]\n- Trừ nhu cầu bán chờ hàng (${leadTimeArrival.toFixed(1)} ngày): -${demandLeadTime.toFixed(2)}\n=> Tồn dự kiến khi SOQ đến: ${expectedInvAtArrival.toFixed(2)}`;
            let inputTooltip = `Nhập/Giao hàng lúc T (${strT}): [ ${finalInput.toFixed(2)} ]`;
            let disposalTooltip = `KHÔNG PHẠT HỦY (Ratio quá thấp hoặc không đủ gốc chia)`;

            let baseForDisposal = prevInv + prevInput;
            let disposalRatio = 0;

            if (baseForDisposal > 0) {
                disposalRatio = finalDisp / baseForDisposal; // Hủy(T) / (Tồn(<T) + Nhập(<T))
            } else {
                disposalRatio = 0; // BỎ QUA GIẢM TRỪ NẾU KHÔNG TÌM THẤY LỊCH SỬ DỮ LIỆU
            }

            if (finalDisp > 0) {
                disposalTooltip = `Công thức: Hủy(T) / (Tồn(<T) + Nhập(<T))\n`;
                disposalTooltip += `= ${finalDisp.toFixed(2)} / (${prevInv.toFixed(2)} + ${prevInput.toFixed(2)})\n`;
                if (baseForDisposal > 0) {
                    disposalTooltip += `= ${(disposalRatio * 100).toFixed(1)}%\n`;
                } else {
                    disposalTooltip += `=> Bỏ qua phạt giảm trừ do thiếu dữ liệu quá khứ\n`;
                }
                disposalTooltip += `(Ghi chú: Lấy Tồn cũ: ${strPrevInv}, Nhập cũ: ${strPrevInput})`;
            }

            if (finalDisp > 0) {
                let category = productCategoryMap.get(data.prodStd.toLowerCase()) || '';
                let isRTE_or_Leaf = category.includes('RTE') || category.includes('RAU LÁ');
                let isRoot = category.includes('CỦ');

                if (!category && RTE_PRODUCTS.some(p => data.bestName.toLowerCase().includes(p.toLowerCase()))) {
                    isRTE_or_Leaf = true;
                }

                let threshold = isRTE_or_Leaf ? 0.30 : (isRoot ? 0.15 : 0.15);

                if (disposalRatio > threshold) {
                    penaltyApplied = finalDisp * 0.5; // Giảm trừ 50%
                    totalDemand -= penaltyApplied;
                    disposalTooltip += `\n\n--> KÍCH HOẠT PHẠT DO QUÁ NGƯỠNG (${(threshold * 100).toFixed(0)}%)`;
                }
            }

            let soq = totalDemand - expectedInvAtArrival;
            soq = Math.max(Math.ceil(soq), 0);

            // HIỂN THỊ ĐẦY ĐỦ SOQ NẾU CÓ BẤT KỲ Ý NGHĨA KINH DOANH NÀO
            // Ẩn dòng có TẤT CẢ = 0
            if (soq === 0 && totalDemand === 0 && finalInv === 0 && finalInput === 0 && finalDisp === 0) {
                return;
            }

            let storeNameStr = storeNamesMap.get(data.storeID) || data.storeOrig;

            let totalDemandRaw = totalDemand + penaltyApplied;
            let breakdownTip = `Công thức: Demand (Nhu cầu gốc) + SafetyStock. \n- Nhu cầu gốc (Coverage): ${coverageDemandBase.toFixed(2)}\n- SafetyStock: +${safetyStock.toFixed(2)} \n- Penalty (Giảm trừ): -${penaltyApplied.toFixed(2)}`;

            finalResults.push({
                'sap': data.storeID,
                'store': storeNameStr,
                'product': data.bestName,
                'ads': forecastDay.toFixed(2),
                'trend': trendExport,
                'ads_weekday': wWeekdayAds.toFixed(2),
                'ads_weekend': wWeekendAds.toFixed(2),
                'growth': mAds > 0 ? `${leadtimeGrowth.toFixed(1)}%` : (basePeriodDemand > 0 ? 'New' : '0%'),
                'leadtime': coverageLT,
                'demand': (totalDemand + penaltyApplied).toFixed(2),
                'inventory': Number(finalInv.toFixed(2)),
                'input': Number(finalInput.toFixed(2)),
                'penalty': penaltyApplied > 0 ? `-${penaltyApplied.toFixed(2)}` : '0',
                'soq': soq,
                // Tooltips cho phần lịch sử
                'tip_weekday': `Tổng bán thực tế: ${wWeekdayQty.toFixed(2)} / ${wWeekdayDaysCount} ngày T2-T6 của Store`,
                'tip_weekend': `Tổng bán thực tế: ${wWeekendQty.toFixed(2)} / ${wWeekendDaysCount} ngày T7-CN của Store`,
                'tip_leadtime': `Coverage: ${coverageLT} ngày. (Chỉ tính lượng bán ra trong ${coverageLT} ngày giao hàng, không tính phần thiếu hụt trong ${leadTimeArrival.toFixed(1)} ngày chờ)`,
                'tip_demand': breakdownTip,
                'tip_inventory': invTooltip,
                'tip_input': inputTooltip,
                'tip_penalty': disposalTooltip
            });

            let tr = document.createElement('tr');
            tr.innerHTML = `
            <td>${data.storeID}</td>
            <td>${storeNameStr}</td>
            <td>${data.bestName}</td>
            <td>${forecastDay.toFixed(2)}</td>
            <td><b>${trendHtml}</b></td>
            <td title="Tổng bán thực tế: ${wWeekdayQty.toFixed(2)} / ${wWeekdayDaysCount} ngày T2-T6 của Store">${wWeekdayAds.toFixed(2)}</td>
            <td title="Tổng bán thực tế: ${wWeekendQty.toFixed(2)} / ${wWeekendDaysCount} ngày T7-CN của Store">${wWeekendAds.toFixed(2)}</td>
            <td><b>${growthHtml}</b></td>
            <td><span title="Coverage: ${coverageLT} ngày. (Chỉ tính lượng bán ra trong ${coverageLT} ngày giao hàng, không tính phần thiếu hụt trong ${leadTimeArrival.toFixed(1)} ngày chờ)">${coverageLT}</span></td>
            <td title="${breakdownTip}">${totalDemandRaw.toFixed(2)}</td>
            <td class="warning" title="${invTooltip}">${Number(finalInv.toFixed(2))}</td>
            <td class="highlight" title="${inputTooltip}">${Number(finalInput.toFixed(2))}</td>
            <td style="color:${penaltyApplied > 0 ? 'var(--danger)' : ''}" title="${disposalTooltip}">${penaltyApplied > 0 ? `-${penaltyApplied.toFixed(2)}` : '0'}</td>
            <td class="highlight">${soq}</td>
        `;
            tbody.appendChild(tr);
        });

        if (finalResults.length === 0) {
            let monthlyKeys = (datasets.monthly && datasets.monthly.length > 0) ? Object.keys(datasets.monthly[0]).join(', ') : 'No data';
            let invKeys = (datasets.inventory && datasets.inventory.length > 0) ? Object.keys(datasets.inventory[0]).join(', ') : 'No data';
            let schedKeys = (datasets.schedule && datasets.schedule.length > 0) ? Object.keys(datasets.schedule[0]).join(', ') : 'No data';

            tbody.innerHTML = `<tr><td colspan="14" style="text-align:left; color: var(--danger); padding: 2rem;">
            <strong>Không tìm thấy bất kỳ dữ liệu hợp lệ nào. (Tồn kho, hàng nhập và lịch giao không khớp ngàm dữ liệu, hoặc tất cả đều bằng 0).</strong><br/><br/>
            <div style="font-family: monospace; font-size:12px; color: var(--text-muted);">
                <strong>--- TRÌNH KIỂM TRA LỖI NỘI BỘ ---</strong><br/>
                - Schedule Headers: ${schedKeys}<br/>
                - Monthly Headers: ${monthlyKeys}<br/>
                - Inventory Headers: ${invKeys}<br/>
                - Lịch Giao Hàng quét được: ${validSAPs.size} mã hợp lệ<br/>
                - Mapping File quét được: ${mappingMap ? mappingMap.size : 0} cặp quy đổi.<br/>
                - Master List đăng ký được: ${allItems.size} mã sản phẩm.
            </div>
            <p>Vui lòng chụp màn hình đoạn mã màu xám này và gửi lại để kỹ sư hoàn tất căn chỉnh file.</p>
        </td></tr>`;
        }

        resultsSection.style.display = 'block';
        if (finalResults.length > 0) {
            btnExport.style.display = 'inline-block';
            
             // --- LƯU LỊCH SỬ TÍNH TOÁN NGAY LẬP TỨC ĐỂ XEM LẠI Ở TAB "LỊCH SỬ TẢI LÊN" (EXPIRES QUA ĐÊM) ---
             saveToDB('soq_latest_filename', scheduleFileName);
             saveToDB('soq_latest_html', tbody.innerHTML);
             saveToDB('soq_latest_array', finalResults);

             // --- LƯU LÊN FIREBASE (CLOUD STORAGE) ---
             if (typeof firebase !== 'undefined') {
                 let userName = inputUserName ? inputUserName.value.trim() : "Hệ thống";
                 if (!userName) userName = "Ẩn danh";

                 const now = new Date();
                 const dateStr = now.toISOString().split('T')[0]; // YYYY-MM-DD

                 const payload = {
                     results: finalResults,
                     filename: scheduleFileName,
                     timestamp: now.getTime(),
                     dateStr: dateStr,
                     userName: userName
                 };

                 firebase.database().ref('latest_soq').set(payload)
                     .then(() => console.log("Đã cập nhật SOQ mới nhất lên Cloud."))
                     .catch(err => console.error("Lỗi lưu Cloud:", err));
             }
        }
    } catch (err) {
        console.error("Lỗi tính toán SOQ:", err);
        alert("Lỗi tính toán: " + err.message + "\n\nBạn hãy kiểm tra xem các file đã được tải lên đầy đủ chưa nhé!");
        btnCalculate.disabled = false;
        btnCalculate.textContent = "Tiến hành tính SOQ";
    }
});

// Export to Excel (Bypass Security Block for local file:///)
btnExport.addEventListener('click', () => {
    // Map lại keys sang Tiếng Việt chuyên nghiệp để xuất Excel
    const exportData = finalResults.map(item => ({
        'Mã SAP (Store)': item.sap,
        'Tên Cửa Hàng': item.store,
        'Tên Sản Phẩm': item.product,
        'Trung Bình Bán/Ngày': item.ads,
        'Xu Hướng Bán (%)': item.trend,
        'ADS T2-T6': item.ads_weekday,
        'ADS T7-CN': item.ads_weekend,
        'Tăng trưởng theo leadtime (%)': item.growth,
        'Leadtime': item.leadtime,
        'Total Demand': item.demand,
        'Tồn (Inv)': item.inventory,
        'Nhập (Input)': item.input,
        'Giảm trừ (Penalty)': item.penalty,
        'SOQ': item.soq
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "SOQ_Results");

    // Khử dấu tiếng Việt và ký tự lạ để tránh Browser chặn tải
    let safeName = String(scheduleFileName).normalize('NFD').replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9_\-]/g, "_");
    let exportName = `SOQ_Data_${safeName}.xlsx`;

    // Custom File downloader to bypass 'Cần có quyền tải xuống' warning
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
        alert("Lỗi tải file: Trình duyệt của bạn khóa quyền tải cục bộ. Hãy mở trang này bằng Chrome nhé!");
    }
});

// --- BỘ LỌC TÌM KIẾM ---
const searchStoreInput = document.getElementById('search-store');
const searchProductInput = document.getElementById('search-product');

function filterTable() {
    if (!searchStoreInput || !searchProductInput) return;
    const storeQuery = searchStoreInput.value.toLowerCase();
    const productQuery = searchProductInput.value.toLowerCase();
    const rows = document.querySelectorAll('#soq-tbody tr');

    rows.forEach(row => {
        if (row.cells.length < 3) return; // Skip special rows like empty data messages
        const sap = row.cells[0].textContent.toLowerCase();
        const storeName = row.cells[1].textContent.toLowerCase();
        const productName = row.cells[2].textContent.toLowerCase();

        const matchStore = sap.includes(storeQuery) || storeName.includes(storeQuery);
        const matchProduct = productName.includes(productQuery);

        if (matchStore && matchProduct) {
            row.style.display = '';
        } else {
            row.style.display = 'none';
        }
    });
}

if (searchStoreInput && searchProductInput) {
    searchStoreInput.addEventListener('input', filterTable);
    searchProductInput.addEventListener('input', filterTable);
}



// --- COLLAPSE SIDEBAR ---
const btnToggleSidebar = document.getElementById('btn-toggle-sidebar');
const sidebar = document.querySelector('.sidebar');

if (btnToggleSidebar && sidebar) {
    btnToggleSidebar.addEventListener('click', () => {
        if (sidebar.style.display === 'none') {
            sidebar.style.display = 'flex';
            btnToggleSidebar.innerHTML = '<span>◄</span> Ẩn Menu trái';
        } else {
            sidebar.style.display = 'none';
            btnToggleSidebar.innerHTML = '<span>►</span> Hiện Menu trái';
        }
    });
}
// --- ĐIỀU CHUYỂN MENU TAB LỊCH SỬ VÀ BẢNG TÍNH ---
const navDashboard = document.getElementById('nav-dashboard');
const navHistory = document.getElementById('nav-history');

if (navHistory && navDashboard) {
    navHistory.addEventListener('click', async (e) => {
        e.preventDefault();
        document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
        navHistory.classList.add('active');
        
        // Ẩn khu vực tải file
        document.querySelector('.upload-section').style.display = 'none';
        
        let tbody = document.getElementById('soq-tbody');
        let titleSpan = document.querySelector('.results-section h2');
        let btnExport = document.getElementById('btn-export');

        // Hiện section kết quả trước để người dùng thấy đang load
        document.getElementById('results-section').style.display = 'block';
        tbody.innerHTML = `<tr><td colspan="14" style="text-align:center; padding: 2rem;">🔄 Đang tải lịch sử từ Cloud...</td></tr>`;

        // 1. Kiểm tra Firebase trước (Shared History)
        if (typeof firebase !== 'undefined') {
            firebase.database().ref('latest_soq').once('value').then(async (snapshot) => {
                const data = snapshot.val();
                const todayStr = new Date().toISOString().split('T')[0];

                if (data && data.dateStr === todayStr) {
                    // Dữ liệu hợp lệ (trong ngày)
                    finalResults = data.results;
                    scheduleFileName = data.filename;

                    // Render bảng từ Array
                    renderTableFromArray(data.results);
                    btnExport.style.display = 'inline-block';

                    let timeStr = new Date(data.timestamp).toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' });
                    titleSpan.innerHTML = `Kết Quả Dự Báo <span style="font-size: 0.6em; background: rgba(76, 175, 80, 0.2); color: #4caf50; border: 1px solid #4caf50; padding: 4px 8px; border-radius: 4px; margin-left: 10px; vertical-align: middle;">Shared: ${data.userName} (${timeStr})</span>`;
                } else {
                    // Không có dữ liệu Cloud hôm nay -> Fallback về Local Cache của chính mình
                    loadLocalHistoryFallback(tbody, titleSpan, btnExport);
                }
            }).catch(err => {
                console.error("Lỗi tải Cloud:", err);
                loadLocalHistoryFallback(tbody, titleSpan, btnExport);
            });
        } else {
            loadLocalHistoryFallback(tbody, titleSpan, btnExport);
        }
    });

    // Hàm bổ trợ Render bảng từ mảng dữ liệu
    function renderTableFromArray(arr) {
        let tbody = document.getElementById('soq-tbody');
        tbody.innerHTML = '';
        arr.forEach(item => {
            let tr = document.createElement('tr');

            // --- Bổ sung xử lý màu sắc và biểu tượng như bản SOQ gốc ---
            
            // 1. Phân tích Xu hướng (Trend)
            let trendVal = String(item.trend || '-').trim();
            // Xử lý cả trường hợp số dạng "4.2%" hoặc "+4.2%" hoặc "▲ 4.2%"
            let trendNum = parseFloat(trendVal.replace(/[▲▼+%\s]/g, ''));
            let trendHtml = `<span>${trendVal}</span>`;
            
            if (trendVal.toLowerCase().includes('new') || trendVal.toLowerCase().includes('mới')) {
                trendHtml = `<span style="color: var(--success)">▲ Mới bán</span>`;
            } else if (!isNaN(trendNum)) {
                if (Math.abs(trendNum) < 1e-6) {
                    trendHtml = `<span>0.0%</span>`;
                } else if (trendNum > 0 || trendVal.includes('+') || trendVal.includes('▲')) {
                    trendHtml = `<span style="color: var(--success)">▲ ${Math.abs(trendNum).toFixed(1)}%</span>`;
                } else if (trendNum < 0 || trendVal.includes('-') || trendVal.includes('▼')) {
                    trendHtml = `<span style="color: var(--danger)">▼ ${Math.abs(trendNum).toFixed(1)}%</span>`;
                }
            }

            // 2. Phân tích Tăng trưởng (Growth)
            let growthVal = String(item.growth || '-').trim();
            let growthNum = parseFloat(growthVal.replace(/[▲▼+%\s]/g, ''));
            let growthHtml = `<span>${growthVal}</span>`;
            
            if (growthVal.toLowerCase().includes('new') || growthVal.toLowerCase().includes('mới')) {
                growthHtml = `<span style="color: var(--success)">${growthVal}</span>`;
            } else if (!isNaN(growthNum)) {
                if (growthNum > 1e-6) { // Positive
                    growthHtml = `<span style="color: var(--success)">+${growthNum.toFixed(1)}%</span>`;
                } else if (growthNum < -1e-6) { // Negative
                    growthHtml = `<span style="color: var(--danger)">-${Math.abs(growthNum).toFixed(1)}%</span>`;
                } else {
                    growthHtml = `<span>0.0%</span>`;
                }
            }

            // Mapping lại các cột từ safe-keys (Firebase) sang giao diện
            tr.innerHTML = `
                <td>${item.sap || ''}</td>
                <td>${item.store || ''}</td>
                <td>${item.product || ''}</td>
                <td>${item.ads || '0.00'}</td>
                <td><b>${trendHtml}</b></td>
                <td title="${item.tip_weekday || ''}">${item.ads_weekday || '0.00'}</td>
                <td title="${item.tip_weekend || ''}">${item.ads_weekend || '0.00'}</td>
                <td><b>${growthHtml}</b></td>
                <td><b><span title="${item.tip_leadtime || ''}">${item.leadtime || ''}</span></b></td>
                <td title="${item.tip_demand || ''}">${item.demand || '0.00'}</td>
                <td class="warning" title="${item.tip_inventory || ''}">${item.inventory || 0}</td>
                <td class="highlight" title="${item.tip_input || ''}">${item.input || 0}</td>
                <td style="color:${parseFloat(item.penalty) < 0 ? 'var(--danger)' : ''}" title="${item.tip_penalty || ''}">${item.penalty || '0'}</td>
                <td class="highlight">${item.soq || 0}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    // Hàm bổ trợ Load Local
    async function loadLocalHistoryFallback(tbody, titleSpan, btnExport) {
        let histArr = await loadFromDB('soq_latest_array'); // Không dùng histHtml từ Cache vì có thể bị stale style
        let histName = await loadFromDB('soq_latest_filename');

        if (histArr && !histArr.invalidated) {
            renderTableFromArray(histArr); // Render lại từ mảng để áp dụng Style mới nhất
            finalResults = histArr;
            if (histName && !histName.invalidated) scheduleFileName = histName;
            btnExport.style.display = 'inline-block';
            titleSpan.innerHTML = `Kết Quả Dự Báo <span style="font-size: 0.6em; background: rgba(255,152,0,0.2); color: #ff9800; border: 1px solid #ff9800; padding: 4px 8px; border-radius: 4px; margin-left: 10px; vertical-align: middle;">Local: Bản lưu máy bạn</span>`;
        } else {
            btnExport.style.display = 'none';
            tbody.innerHTML = `<tr><td colspan="14" style="text-align:center; padding: 2.5rem; color: #ff9800; font-size: 1.1em;"><i class="fas fa-history" style="font-size: 2em; display: block; margin-bottom: 10px; opacity: 0.5;"></i>Không có lịch sử chia sẻ hoặc lịch sử máy bạn đã hết hạn trong ngày hôm nay.</td></tr>`;
            titleSpan.innerHTML = `Kết Quả Dự Báo`;
        }
    }

    navDashboard.addEventListener('click', (e) => {
        e.preventDefault();
        document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
        navDashboard.classList.add('active');
        
        // Hiện lại khu vực Tải file
        document.querySelector('.upload-section').style.display = 'block';
        
        let titleSpan = document.querySelector('.results-section h2');
        if (titleSpan && titleSpan.querySelector('span')) { 
            // Dọn dẹp View Lịch sử (Ép người dùng bấn Tính SOQ lại để tải lại Live Data an toàn)
            titleSpan.innerHTML = `Kết Quả Dự Báo`;
            document.getElementById('soq-tbody').innerHTML = ''; 
            document.getElementById('results-section').style.display = 'none';
            finalResults = [];
        }
    });
}
