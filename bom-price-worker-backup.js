66448ba1a7c7
Content-Disposition: form-data; name="index.js"

var __defProp = Object.defineProperty;
var __name = (target, value) => __defProp(target, "name", { value, configurable: true });

// src/index.ts
var HTML_PAGE = `<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>آپدیت قیمت BOM</title>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        .result-table tr:hover td { background-color: rgba(34,211,238,0.08) !important; }
        .result-table tr:nth-child(even) td { background-color: rgba(31,41,55,0.5); }
        .result-table th { cursor: pointer; user-select: none; }
        .result-table th:hover { background-color: rgba(34,211,238,0.15); }
        .toast { position: fixed; top: 20px; left: 50%; transform: translateX(-50%); background: #059669; color: white; padding: 12px 24px; border-radius: 8px; font-weight: bold; z-index: 9999; opacity: 0; transition: opacity 0.3s; }
        .toast.show { opacity: 1; }
        .search-box { background: #1f2937; border: 1px solid #4b5563; border-radius: 8px; padding: 8px 12px; color: white; width: 100%; }
        .search-box:focus { outline: none; border-color: #22d3ee; }
    </style>
</head>
<body class="bg-gray-900 text-white min-h-screen p-8">
    <div class="max-w-6xl mx-auto bg-gray-800 p-6 rounded-xl shadow-lg">
        <h1 class="text-2xl font-bold mb-6 text-center text-cyan-400">سیستم هوشمند بروزرسانی قیمت دستگاه (BOM)</h1>
        
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6 bg-gray-700 p-4 rounded-lg">
            <div>
                <label class="block mb-2 font-semibold text-cyan-300">1. فایل لیست خرید (بازار):</label>
                <input type="file" id="kharidFile" accept=".xlsx, .xls, .csv" onchange="loadColumns('kharid')" class="w-full bg-gray-800 p-2 rounded border border-gray-600 text-sm">
            </div>
            <div>
                <label class="block mb-2 font-semibold text-cyan-300">2. فایل لیست قطعات دستگاه (BOM):</label>
                <input type="file" id="abzarFile" accept=".xlsx, .xls, .csv" onchange="loadColumns('abzar')" class="w-full bg-gray-800 p-2 rounded border border-gray-600 text-sm">
            </div>
        </div>

        <div class="mb-6 bg-gray-700 p-4 rounded-lg border border-cyan-500">
            <h3 class="font-bold text-yellow-300 mb-3">تنظیمات محاسباتی و فیلتر زمان:</h3>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                    <label class="block text-xs mb-1">محاسبه قیمت بر اساس:</label>
                    <select id="calcMethod" class="w-full bg-gray-800 p-2 rounded border border-gray-600 text-sm">
                        <option value="last">آخرین قیمت خرید</option>
                        <option value="avg">میانگین قیمت خریدها</option>
                    </select>
                </div>
                <div>
                    <label class="block text-xs mb-1">محاسبه قیمت‌ها تا تاریخ (اختیاری):</label>
                    <input type="text" id="targetDateFilter" placeholder="مثلاً 1404/05/30 یا خالی برای همه" class="w-full bg-gray-800 p-2 rounded border border-gray-600 text-sm text-right">
                </div>
            </div>
        </div>

        <div id="mappingSection" class="hidden grid grid-cols-1 md:grid-cols-2 gap-6 mb-6 bg-gray-700 p-4 rounded-lg border border-gray-600">
            <div>
                <h3 class="font-bold text-yellow-300 mb-3">تنظیم ستون‌های فایل خرید:</h3>
                <div class="mb-3"><label class="block text-xs mb-1">ستون کد کالا:</label><select id="mapKharidCode" class="w-full bg-gray-800 p-2 rounded border border-gray-600 text-sm"></select></div>
                <div class="mb-3"><label class="block text-xs mb-1">ستون تاریخ خرید:</label><select id="mapKharidDate" class="w-full bg-gray-800 p-2 rounded border border-gray-600 text-sm"></select></div>
                <div><label class="block text-xs mb-1">ستون قیمت خرید:</label><select id="mapKharidPrice" class="w-full bg-gray-800 p-2 rounded border border-gray-600 text-sm"></select></div>
            </div>
            <div>
                <h3 class="font-bold text-yellow-300 mb-3">تنظیم ستون‌های فایل قطعات:</h3>
                <div class="mb-3"><label class="block text-xs mb-1">ستون کد کالا:</label><select id="mapAbzarCode" class="w-full bg-gray-800 p-2 rounded border border-gray-600 text-sm"></select></div>
                <div class="mb-3"><label class="block text-xs mb-1">ستون نام کالا (اختیاری):</label><select id="mapAbzarName" class="w-full bg-gray-800 p-2 rounded border border-gray-600 text-sm"></select></div>
                <div><label class="block text-xs mb-1">ستون تعداد / ضریب مصرف:</label><select id="mapAbzarQty" class="w-full bg-gray-800 p-2 rounded border border-gray-600 text-sm"></select></div>
            </div>
        </div>

        <button onclick="processFiles()" class="w-full bg-cyan-600 hover:bg-cyan-500 font-bold py-3 rounded transition shadow-lg">پردازش و تطبیق نهایی اطلاعات</button>
        <div id="filterSection" class="hidden mt-4 bg-gray-700 p-4 rounded-lg">
            <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                    <label class="block text-xs mb-1 text-gray-400">جستجو:</label>
                    <input type="text" id="searchInput" placeholder="نام یا کد کالا..." oninput="applyFilters()" class="search-box">
                </div>
                <div>
                    <label class="block text-xs mb-1 text-gray-400">فیلتر وضعیت:</label>
                    <select id="statusFilter" onchange="applyFilters()" class="w-full bg-gray-800 p-2 rounded border border-gray-600 text-sm">
                        <option value="all">همه</option>
                        <option value="found">فقط یافت شده</option>
                        <option value="notfound">فقط یافت نشده</option>
                    </select>
                </div>
                <div class="flex items-end">
                    <button onclick="clearFilters()" class="bg-gray-600 hover:bg-gray-500 text-white font-bold py-2 px-4 rounded transition">پاک کردن فیلترها</button>
                </div>
            </div>
            <div id="partsCounter" class="mt-3 text-sm text-gray-400"></div>
        </div>
        <div id="result" class="mt-8 overflow-x-auto"></div>
    </div>

    <script>
        let kharidRawData = [];
        document.addEventListener('DOMContentLoaded', loadSettings);
        let abzarRawData = [];

        async function loadColumns(type) {
            const fileInput = document.getElementById(type + 'File').files[0];
            if (!fileInput) return;
            const data = await readExcel(fileInput);
            if (type === 'kharid') kharidRawData = data;
            else abzarRawData = data;

            if (data.length > 0) {
                const keys = Object.keys(data[0]);
                let options = '<option value="">-- انتخاب کنید --</option>';
                keys.forEach(k => { options += '<option value="' + k + '">' + k + '</option>'; });

                if (type === 'kharid') {
                    document.getElementById('mapKharidCode').innerHTML = options;
                    document.getElementById('mapKharidDate').innerHTML = options;
                    document.getElementById('mapKharidPrice').innerHTML = options;
                    setBestMatch('mapKharidCode', keys, ['کد', 'code', 'شناسه']);
                    setBestMatch('mapKharidDate', keys, ['تاریخ', 'date', 'زمان']);
                    setBestMatch('mapKharidPrice', keys, ['قیمت', 'مبلغ', 'فی', 'price']);
                } else {
                    document.getElementById('mapAbzarCode').innerHTML = options;
                    document.getElementById('mapAbzarName').innerHTML = options;
                    document.getElementById('mapAbzarQty').innerHTML = options;
                    setBestMatch('mapAbzarCode', keys, ['کد', 'code', 'شناسه']);
                    setBestMatch('mapAbzarName', keys, ['نام', 'شرح', 'عنوان', 'name']);
                    setBestMatch('mapAbzarQty', keys, ['تعداد', 'ضریب', 'مقدار', 'qty']);
                }
                document.getElementById('mappingSection').classList.remove('hidden');
            }
        }

        function setBestMatch(elementId, keys, keywords) {
            const found = keys.find(k => keywords.some(kw => k.toLowerCase().includes(kw)));
            if (found) document.getElementById(elementId).value = found;
        }

        function normalizeDate(dStr) {
            if (!dStr) return '';
            return String(dStr).trim().replace(/[\\/\\-\\.]/g, '');
        }

        async function processFiles() {
            if (kharidRawData.length === 0 || abzarRawData.length === 0) {
                alert('لطفاً ابتدا هر دو فایل را انتخاب کنید.');
                return;
            }

            const rawTargetDate = document.getElementById('targetDateFilter').value.trim();
            const isFilterActive = rawTargetDate && !rawTargetDate.includes('مثلاً') && rawTargetDate !== '1404/05/30';
            const normalizedTargetDate = isFilterActive ? normalizeDate(rawTargetDate) : '';

            const kCodeKey = document.getElementById('mapKharidCode').value;
            const kDateKey = document.getElementById('mapKharidDate').value;
            const kPriceKey = document.getElementById('mapKharidPrice').value;
            const aCodeKey = document.getElementById('mapAbzarCode').value;
            const aNameKey = document.getElementById('mapAbzarName').value;
            const aQtyKey = document.getElementById('mapAbzarQty').value;
            const calcMethod = document.getElementById('calcMethod').value;

            if (!kCodeKey || !kPriceKey || !aCodeKey) {
                alert('لطفاً ستون‌های ضروری (کد و قیمت) را مشخص کنید.');
                return;
            }

            const priceMapObj = {};
            kharidRawData.forEach(row => {
                const code = String(row[kCodeKey] || '').trim().split('.')[0];
                const rawDate = kDateKey ? String(row[kDateKey] || '').trim() : '';
                const normDate = normalizeDate(rawDate);
                const price = parseFloat(String(row[kPriceKey]).replace(/,/g, '')) || 0;
                
                if (!code) return;
                
                if (isFilterActive && normalizedTargetDate && normDate) {
                    if (normDate > normalizedTargetDate) return;
                }

                if (!priceMapObj[code]) {
                    priceMapObj[code] = { prices: [], lastDate: rawDate, normLastDate: normDate, lastPrice: price };
                }
                priceMapObj[code].prices.push(price);
                
                if (!priceMapObj[code].normLastDate || (normDate && normDate >= priceMapObj[code].normLastDate)) {
                    priceMapObj[code].normLastDate = normDate;
                    priceMapObj[code].lastDate = rawDate;
                    priceMapObj[code].lastPrice = price;
                }
            });

            const finalPriceMap = {};
            Object.keys(priceMapObj).forEach(code => {
                const item = priceMapObj[code];
                let resolvedPrice = calcMethod === 'avg' ? Math.round(item.prices.reduce((a, b) => a + b, 0) / item.prices.length) : item.lastPrice;
                finalPriceMap[code] = { date: item.lastDate || 'موجود', price: resolvedPrice, count: item.prices.length };
            });

            let grandTotal = 0;
            const processed = abzarRawData.map((row, idx) => {
                const code = String(row[aCodeKey] || '').trim().split('.')[0];
                const name = aNameKey ? String(row[aNameKey] || '') : '';
                const qty = aQtyKey ? (parseFloat(String(row[aQtyKey]).replace(/,/g, '')) || 1) : 1;
                const match = finalPriceMap[code];
                const unitPrice = match ? match.price : 0;
                const latestDate = match ? match.date : '';
                const purchaseCount = match ? match.count : 0;
                const totalPrice = match ? unitPrice * qty : 0;
                if (match) grandTotal += totalPrice;

                return {
                    'ردیف': idx + 1, 'کد کالا': code, 'نام کالا': name, 'تعداد': qty,
                    'تاریخ مرجع': latestDate, 'تعداد دفعات خرید': purchaseCount || 'یافت نشد',
                    'قیمت واحد': unitPrice,
                    'قیمت کل': totalPrice
                };
            });

            window._processedData = processed;
            window._grandTotal = grandTotal;
            saveSettings();
            renderTable(processed, grandTotal);
        }

        function renderTable(data, grandTotal) {
            var html = '<h2 class="text-xl font-bold mb-4 text-green-400">نتیجه پردازش و محاسبات قطعات:</h2>';
            html += '<div id="partsCounterTop" class="mb-3 text-sm text-gray-400"></div>';
            html += '<table class="w-full text-right border-collapse bg-gray-700 rounded overflow-hidden text-sm result-table" id="resultTable">';
            html += '<tr class="bg-gray-600"><th class="p-2" onclick="sortTable(0)">ردیف</th><th class="p-2" onclick="sortTable(1)">کد کالا</th><th class="p-2" onclick="sortTable(2)">نام کالا</th><th class="p-2" onclick="sortTable(3)">ضریب مصرف</th><th class="p-2" onclick="sortTable(4)">تاریخ مرجع</th><th class="p-2" onclick="sortTable(5)">تعداد دفعات خرید</th><th class="p-2" onclick="sortTable(6)">قیمت واحد</th><th class="p-2" onclick="sortTable(7)">قیمت کل</th></tr>';
            data.forEach(function(r, i) {
                var notFound = (r['قیمت واحد'] === 0);
                var qtyVal = r['تعداد'];
                var dateVal = r['تاریخ مرجع'] || '';
                var priceVal = notFound ? '' : r['قیمت واحد'];
                var bgClass = notFound ? 'bg-red-900/30' : '';
                html += '<tr class="border-t border-gray-600 ' + bgClass + '">';
                html += '<td class="p-2">' + r['ردیف'] + '</td>';
                html += '<td class="p-2">' + r['کد کالا'] + '</td>';
                html += '<td class="p-2">' + r['نام کالا'] + '</td>';
                html += '<td class="p-2"><input type="number" class="w-16 bg-gray-800 border border-gray-500 rounded px-1 py-0.5 text-center" id="qty_' + i + '" value="' + qtyVal + '" oninput="autoUpdate(' + i + ')"></td>';
                html += '<td class="p-2"><input type="text" class="w-24 bg-gray-800 border border-gray-500 rounded px-1 py-0.5 text-center" id="date_' + i + '" value="' + dateVal + '"></td>';
                html += '<td class="p-2 text-center text-yellow-300 font-semibold">' + r['تعداد دفعات خرید'] + '</td>';
                html += '<td class="p-2"><input type="number" class="w-24 bg-gray-800 border border-gray-500 rounded px-1 py-0.5 text-center" id="price_' + i + '" value="' + priceVal + '" placeholder="' + (notFound ? 'قیمت را وارد کنید' : '') + '" oninput="autoUpdate(' + i + ')"></td>';
                html += '<td class="p-2 font-bold" id="total_' + i + '">' + (r['قیمت کل'] > 0 ? r['قیمت کل'].toLocaleString() : 'خرید یافت نشد') + '</td>';
                html += '</tr>';
            });
            html += '</table>';
            html += '<div class="mt-4 flex items-center gap-4">';
            html += '<button onclick="recalculate()" class="bg-yellow-500 hover:bg-yellow-400 text-gray-900 font-bold py-2 px-6 rounded transition shadow-lg">محاسبه مجدد</button>';
            html += '<button onclick="exportToExcel()" class="bg-green-600 hover:bg-green-500 text-white font-bold py-2 px-6 rounded transition shadow-lg">خروجی اکسل</button>';
            html += '<div class="text-xl font-bold text-cyan-300">جمع کل نهایی دستگاه: <span id="grandTotal">' + grandTotal.toLocaleString() + '</span> ریال</div>';
            html += '</div>';
            document.getElementById('result').innerHTML = html;
            document.getElementById('filterSection').classList.remove('hidden');
            updatePartsCounter(data);
            window._sortState = { col: -1, asc: true };
        }

        function recalculate() {
            var data = window._processedData;
            if (!data) return;
            var grandTotal = 0;
            data.forEach(function(r, i) {
                var qtyEl = document.getElementById('qty_' + i);
                var priceEl = document.getElementById('price_' + i);
                var dateEl = document.getElementById('date_' + i);
                var totalEl = document.getElementById('total_' + i);
                var qty = parseFloat(qtyEl.value) || 0;
                var price = parseFloat(priceEl.value) || 0;
                var date = dateEl.value || r['تاریخ مرجع'];
                r['تعداد'] = qty;
                r['قیمت واحد'] = price;
                r['تاریخ مرجع'] = date || 'دستی';
                var total = qty * price;
                r['قیمت کل'] = total;
                if (price > 0) grandTotal += total;
                totalEl.textContent = price > 0 ? total.toLocaleString() : 'خرید یافت نشد';
            });
            window._grandTotal = grandTotal;
            document.getElementById('grandTotal').textContent = grandTotal.toLocaleString();
        }



        function autoUpdate(idx) {
            var data = window._processedData;
            if (!data) return;
            var r = data[idx];
            var qty = parseFloat(document.getElementById('qty_' + idx).value) || 0;
            var price = parseFloat(document.getElementById('price_' + idx).value) || 0;
            var total = qty * price;
            r['تعداد'] = qty;
            r['قیمت واحد'] = price;
            r['قیمت کل'] = total;
            var totalEl = document.getElementById('total_' + idx);
            totalEl.textContent = price > 0 ? total.toLocaleString() : 'خرید یافت نشد';
            totalEl.className = 'p-2 font-bold ' + (price > 0 ? 'text-green-300' : 'text-red-400');
            recalculateGrandTotal();
        }

        function recalculateGrandTotal() {
            var data = window._processedData;
            if (!data) return;
            var grandTotal = 0;
            var foundTotal = 0;
            var notFoundCount = 0;
            var foundCount = 0;
            data.forEach(function(r) {
                var price = r['قیمت واحد'] || r['قیمت واحد'] || 0;
                var qty = r['تعداد'] || 0;
                var total = price * qty;
                r['قیمت کل'] = total;
                if (price > 0) {
                    grandTotal += total;
                    foundTotal += total;
                    foundCount++;
                } else {
                    notFoundCount++;
                }
            });
            window._grandTotal = grandTotal;
            document.getElementById('grandTotal').textContent = grandTotal.toLocaleString();
            updatePartsCounter(data);
        }

        function updatePartsCounter(data) {
            var found = 0, notFound = 0;
            data.forEach(function(r) {
                var price = r['قیمت واحد'] || r['قیمت واحد'] || 0;
                if (price > 0) found++; else notFound++;
            });
            var html = 'کل قطعات: ' + data.length + ' | ';
            html += '<span class="text-green-400">یافت شده: ' + found + '</span> | ';
            html += '<span class="text-red-400">یافت نشده: ' + notFound + '</span>';
            var el1 = document.getElementById('partsCounterTop');
            var el2 = document.getElementById('partsCounter');
            if (el1) el1.innerHTML = html;
            if (el2) el2.innerHTML = html;
        }

        function applyFilters() {
            var search = document.getElementById('searchInput').value.toLowerCase();
            var status = document.getElementById('statusFilter').value;
            var rows = document.querySelectorAll('#resultTable tr');
            rows.forEach(function(row, idx) {
                if (idx === 0) return; // skip header
                var cells = row.querySelectorAll('td');
                if (cells.length < 3) return;
                var code = (cells[1].textContent || '').toLowerCase();
                var name = (cells[2].textContent || '').toLowerCase();
                var priceCell = cells[6];
                var priceInput = priceCell.querySelector('input');
                var priceVal = priceInput ? parseFloat(priceInput.value) || 0 : 0;
                var matchSearch = !search || code.includes(search) || name.includes(search);
                var matchStatus = status === 'all' || (status === 'found' && priceVal > 0) || (status === 'notfound' && priceVal === 0);
                row.style.display = (matchSearch && matchStatus) ? '' : 'none';
            });
        }

        function clearFilters() {
            document.getElementById('searchInput').value = '';
            document.getElementById('statusFilter').value = 'all';
            applyFilters();
        }

        function showToast(msg) {
            var t = document.createElement('div');
            t.className = 'toast';
            t.textContent = msg;
            document.body.appendChild(t);
            setTimeout(function() { t.classList.add('show'); }, 10);
            setTimeout(function() { t.classList.remove('show'); setTimeout(function() { t.remove(); }, 300); }, 2000);
        }

        function saveSettings() {
            var settings = {
                calcMethod: document.getElementById('calcMethod').value,
                targetDate: document.getElementById('targetDateFilter').value
            };
            try { localStorage.setItem('bom_settings', JSON.stringify(settings)); } catch(e) {}
        }

        function loadSettings() {
            try {
                var s = JSON.parse(localStorage.getItem('bom_settings'));
                if (s) {
                    if (s.calcMethod) document.getElementById('calcMethod').value = s.calcMethod;
                    if (s.targetDate) document.getElementById('targetDateFilter').value = s.targetDate;
                }
            } catch(e) {}
        }

        var _sortState = { col: -1, asc: true };
        function sortTable(colIdx) {
            var data = window._processedData;
            if (!data) return;
            if (_sortState.col === colIdx) { _sortState.asc = !_sortState.asc; } else { _sortState.col = colIdx; _sortState.asc = true; }
            var keys = ['ردیف', 'کد کالا', 'نام کالا', 'تعداد', 'تاریخ مرجع', 'تعداد دفعات خرید', 'قیمت واحد', 'قیمت کل'];
            var key = keys[colIdx];
            data.sort(function(a, b) {
                var va = a[key] || 0, vb = b[key] || 0;
                if (typeof va === 'string') return _sortState.asc ? va.localeCompare(vb) : vb.localeCompare(va);
                return _sortState.asc ? va - vb : vb - va;
            });
            var gt = window._grandTotal || 0;
            renderTable(data, gt);
            showToast('مرتب شد!');
        }

        function exportToExcel() {
            var data = window._processedData;
            if (!data || data.length === 0) {
                alert('داده‌ای برای خروجی وجود ندارد.');
                return;
            }

            var wsData = [];
            var headers = ['ردیف', 'کد کالا', 'نام کالا', 'تعداد', 'تاریخ مرجع', 'تعداد دفعات خرید', 'قیمت واحد', 'قیمت کل'];
            wsData.push(headers);

            data.forEach(function(r) {
                wsData.push([
                    r['ردیف'],
                    r['کد کالا'],
                    r['نام کالا'],
                    r['تعداد'],
                    r['تاریخ مرجع'] || '',
                    r['تعداد دفعات خرید'],
                    r['قیمت واحد'] || 0,
                    r['قیمت کل'] || 0
                ]);
            });

            // Add grand total row
            wsData.push([]);
            wsData.push(['', '', '', '', '', '', 'جمع کل نهایی:', window._grandTotal || 0]);

            var ws = XLSX.utils.aoa_to_sheet(wsData);

            // Set RTL
            ws['!cols'] = [
                { wch: 8 },   // ردیف
                { wch: 12 },  // کد کالا
                { wch: 30 },  // نام کالا
                { wch: 10 },  // تعداد
                { wch: 15 },  // تاریخ مرجع
                { wch: 18 },  // تعداد دفعات خرید
                { wch: 15 },  // قیمت واحد
                { wch: 18 }   // قیمت کل
            ];

            var wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'نتیجه BOM');

            // Generate file name with date
            var dateStr = new Date().toISOString().slice(0, 10);
            XLSX.writeFile(wb, 'BOM_Result_' + dateStr + '.xlsx');
        }
        function readExcel(file) {
            return new Promise((resolve) => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    resolve(XLSX.utils.sheet_to_json(firstSheet));
                };
                reader.readAsArrayBuffer(file);
            });
        }
    </script>
</body>
</html>`;

var userSessions = /* @__PURE__ */ new Map();

var index_default = {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (request.method === "POST" && url.pathname === "/webhook") {
      try {
        const update = await request.json();
        if (update.callback_query) {
          const chatId = update.callback_query.message.chat.id;
          const data = update.callback_query.data;
          let session = userSessions.get(chatId) || {};
          if (data === "method_last" || data === "method_avg") {
            session.calcMethod = data === "method_last" ? "last" : "avg";
            session.step = "WAITING_DATE";
            userSessions.set(chatId, session);
            await answerCallbackQuery(env.TELEGRAM_BOT_TOKEN, update.callback_query.id, "روش محاسبه ثبت شد.");
            await sendMessage(env.TELEGRAM_BOT_TOKEN, chatId, "📅 لطفاً **تاریخ خرید** موردنظر را وارد کنید تا محاسبات تا آن تاریخ انجام شود.\n\n*(مثلاً تاریخ را به شکل `1404/05/30` وارد کنید، یا اگر محدودیتی ندارید کلمه **همه** را بفرستید)*");
          }
          return new Response("OK");
        }
        if (update.message) {
          const chatId = update.message.chat.id;
          const text = update.message.text || "";
          let session = userSessions.get(chatId) || {};
          if (text === "/start") {
            userSessions.set(chatId, { step: "WAITING_METHOD" });
            await sendMethodSelectionMenu(env.TELEGRAM_BOT_TOKEN, chatId);
          } else if (session.step === "WAITING_DATE") {
            session.targetDate = text.trim() === "همه" || text.trim() === "all" ? "" : text.trim();
            session.step = "WAITING_FILES";
            userSessions.set(chatId, session);
            await sendMessage(env.TELEGRAM_BOT_TOKEN, chatId, "✅ تاریخ ثبت شد.\n\nاکنون لطفاً **دو فایل اکسل** خود را به ترتیب ارسال کنید:\n1. ابتدا فایل لیست خرید (بازار)\n2. سپس فایل لیست قطعات (BOM)");
          } else if (update.message.document) {
            const doc = update.message.document;
            const fileName = doc.file_name || "";
            const fileId = doc.file_id;
            const fileRes = await fetch(`https://api.telegram.org/bot${env.TELEGRAM_BOT_TOKEN}/getFile?file_id=${fileId}`);
            const fileJson = await fileRes.json();
            if (fileJson.ok) {
              const filePath = fileJson.result.file_path;
              const downloadRes = await fetch(`https://api.telegram.org/file/bot${env.TELEGRAM_BOT_TOKEN}/${filePath}`);
              const fileBuffer = await downloadRes.arrayBuffer();
              if (!session.kharidBuffer) {
                session.kharidBuffer = fileBuffer;
                userSessions.set(chatId, session);
                await sendMessage(env.TELEGRAM_BOT_TOKEN, chatId, `✅ فایل خرید (${fileName}) دریافت شد.\nاکنون لطفاً فایل لیست قطعات (BOM) را ارسال کنید.`);
              } else {
                session.abzarBuffer = fileBuffer;
                userSessions.set(chatId, session);
                await sendMessage(env.TELEGRAM_BOT_TOKEN, chatId, `✅ فایل قطعات (${fileName}) دریافت شد.\n⏳ در حال پردازش و محاسبه نهایی...`);
                await processAndSendResult(env.TELEGRAM_BOT_TOKEN, chatId, session);
              }
            }
          }
        }
        return new Response("OK");
      } catch (err) {
        return new Response("Error", { status: 500 });
      }
    }
    return new Response(HTML_PAGE, {
      headers: { "Content-Type": "text/html;charset=UTF-8" }
    });
  }
};

async function sendMethodSelectionMenu(token, chatId) {
  await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      chat_id: chatId,
      text: "سلام! خوش آمدید.\n\nلطفاً روش محاسبه قیمت‌ها را انتخاب کنید:",
      reply_markup: {
        inline_keyboard: [
          [
            { text: "📌 آخرین قیمت خرید", callback_data: "method_last" },
            { text: "📊 میانگین قیمت خرید", callback_data: "method_avg" }
          ]
        ]
      }
    })
  });
}
__name(sendMethodSelectionMenu, "sendMethodSelectionMenu");

function normalizeDate(dStr) {
  if (!dStr) return '';
  return String(dStr).trim().replace(/[\\/\\-\\.]/g, '');
}
__name(normalizeDate, "normalizeDate");

async function processAndSendResult(token, chatId, session) {
  try {
    const kharidWorkbook = XLSX.read(new Uint8Array(session.kharidBuffer), { type: "array" });
    const kharidData = XLSX.utils.sheet_to_json(kharidWorkbook.Sheets[kharidWorkbook.SheetNames[0]]);
    const abzarWorkbook = XLSX.read(new Uint8Array(session.abzarBuffer), { type: "array" });
    const abzarData = XLSX.utils.sheet_to_json(abzarWorkbook.Sheets[abzarWorkbook.SheetNames[0]]);
    
    const findKey = /* @__PURE__ */ __name((keys, keywords) => keys.find((k) => keywords.some((kw) => k.toLowerCase().includes(kw))) || keys[0], "findKey");
    const kKeys = kharidData.length > 0 ? Object.keys(kharidData[0]) : [];
    const aKeys = abzarData.length > 0 ? Object.keys(abzarData[0]) : [];
    
    const kCodeKey = findKey(kKeys, ["کد", "code", "شناسه"]);
    const kDateKey = findKey(kKeys, ["تاریخ", "date", "زمان"]);
    const kPriceKey = findKey(kKeys, ["قیمت", "مبلغ", "فی", "price"]);
    const aCodeKey = findKey(aKeys, ["کد", "code", "شناسه"]);
    const aQtyKey = findKey(aKeys, ["تعداد", "ضریب", "مقدار", "qty"]);
    
    const targetDateFilter = session.targetDate || "";
    const isFilterActive = targetDateFilter && targetDateFilter !== "همه";
    const normalizedTargetDate = isFilterActive ? normalizeDate(targetDateFilter) : "";
    
    const priceMapObj = {};
    kharidData.forEach((row) => {
      const code = String(row[kCodeKey] || "").trim().split('.')[0];
      const rawDate = kDateKey ? String(row[kDateKey] || "").trim() : "";
      const normDate = normalizeDate(rawDate);
      const price = parseFloat(String(row[kPriceKey] || "").replace(/,/g, "")) || 0;
      
      if (!code) return;
      if (isFilterActive && normalizedTargetDate && normDate && normDate > normalizedTargetDate) return;
      
      if (!priceMapObj[code]) {
        priceMapObj[code] = { prices: [], lastDate: rawDate, normLastDate: normDate, lastPrice: price };
      }
      priceMapObj[code].prices.push(price);
      
      if (!priceMapObj[code].normLastDate || (normDate && normDate >= priceMapObj[code].normLastDate)) {
        priceMapObj[code].normLastDate = normDate;
        priceMapObj[code].lastDate = rawDate;
        priceMapObj[code].lastPrice = price;
      }
    });

    let grandTotal = 0;
    let reportLines = [
      `📊 *نتیجه پردازش نهایی BOM*`,
      `🔹 روش: ${session.calcMethod === "avg" ? "میانگین قیمت‌ها" : "آخرین قیمت"}`,
      `🔹 تا تاریخ: ${targetDateFilter || "بدون محدودیت"}`,
      ``
    ];
    
    abzarData.forEach((row, idx) => {
      const code = String(row[aCodeKey] || "").trim().split('.')[0];
      const qty = aQtyKey ? parseFloat(String(row[aQtyKey] || "").replace(/,/g, "")) || 1 : 1;
      const match = priceMapObj[code];
      let unitPrice = 0;
      if (match) {
        if (session.calcMethod === "avg") {
          unitPrice = Math.round(match.prices.reduce((a, b) => a + b, 0) / match.prices.length);
        } else {
          unitPrice = match.lastPrice;
        }
      }
      const totalPrice = unitPrice * qty;
      grandTotal += totalPrice;
      reportLines.push(`${idx + 1}. کد: ${code} | تعداد: ${qty} | فی: ${unitPrice.toLocaleString()} | کل: ${totalPrice.toLocaleString()}`);
    });
    
    reportLines.push(`\n💰 *جمع کل نهایی دستگاه:* ${grandTotal.toLocaleString()} تومان`);
    await sendMessage(token, chatId, reportLines.join("\n"));
    userSessions.delete(chatId);
  } catch (calcErr) {
    await sendMessage(token, chatId, "❌ خطا در پردازش فایل‌های اکسل. لطفاً از درست بودن ساختار ستون‌ها اطمینان حاصل کنید.");
  }
}
__name(processAndSendResult, "processAndSendResult");

async function sendMessage(token, chatId, text) {
  await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ chat_id: chatId, text, parse_mode: "Markdown" })
  });
}
__name(sendMessage, "sendMessage");

async function answerCallbackQuery(token, callbackQueryId, text) {
  await fetch(`https://api.telegram.org/bot${token}/answerCallbackQuery`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ callback_query_id: callbackQueryId, text })
  });
}
__name(answerCallbackQuery, "answerCallbackQuery");

export {
  index_default as default
};
--f9efb10de8da2973c71d97f04c279c2f25ef99a4c0f97e6166448ba1a7c7--