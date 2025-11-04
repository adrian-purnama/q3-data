// Dashboard Configuration
const CONFIG = {
    accentColor: '#2563eb',
    gridColor: '#e5e5e5',
    textColor: '#666666',
    primaryTextColor: '#1a1a1a'
};

// Global data storage
let dashboardData = {
    headers: [],
    rows: [],
    parsed: []
};

// Chart instances
let charts = {};

// Column mapping (keeping original Indonesian names)
const COLUMN_MAP = {
    rfqNumber: 'NO. PENAWARAN',
    date: 'DATE',
    sales: 'MARKERTING',
    customer: 'CUSTOMER NAME',
    hargaNew: 'HARGA (NEW)',
    status: 'KETERANGAN',
    total: 'TOTAL'
};

// ============================================
// CSV Loading & Parsing
// ============================================

async function loadCSVFile(filePath) {
    try {
        const response = await fetch(filePath);
        const text = await response.text();
        return parseCSV(text);
    } catch (error) {
        console.error('Error loading CSV:', error);
        return null;
    }
}

function parseCSV(text) {
    const lines = text.split('\n').filter(line => line.trim());
    if (lines.length === 0) return { headers: [], rows: [] };
    
    // Parse header - handle quoted fields
    const headers = parseCSVLine(lines[0]).map(h => h.trim());
    const rows = [];
    
    for (let i = 1; i < lines.length; i++) {
        const row = parseCSVLine(lines[i]);
        if (row.length > 0 && row.some(cell => cell.trim())) {
            rows.push(row);
        }
    }
    
    return { headers, rows };
}

function parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        
        if (char === '"') {
            inQuotes = !inQuotes;
        } else if (char === ',' && !inQuotes) {
            result.push(current.trim());
            current = '';
        } else {
            current += char;
        }
    }
    result.push(current.trim());
    
    return result;
}

// ============================================
// Excel Loading & Parsing
// ============================================

function loadExcelFile(file) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Get first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON (header: 1 means first row is header)
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            
            if (jsonData.length === 0) {
                showError('File Excel kosong atau tidak valid.');
                return;
            }
            
            // First row is headers
            const headers = jsonData[0].map(h => String(h).trim());
            const rows = jsonData.slice(1).filter(row => row.some(cell => cell && String(cell).trim()));
            
            processData(headers, rows);
            
        } catch (error) {
            console.error('Error parsing Excel:', error);
            showError('Error membaca file Excel: ' + error.message);
        }
    };
    
    reader.onerror = function() {
        showError('Error membaca file.');
    };
    
    reader.readAsArrayBuffer(file);
}

function processData(headers, rows) {
    dashboardData.headers = headers;
    dashboardData.rows = rows;
    dashboardData.parsed = parseData(headers, rows);
    
    console.log('Data loaded:', dashboardData.parsed.length, 'records');
    console.log('Headers:', headers);
    
    // Show preview
    showPreview(headers, rows.slice(0, 10));
    
    // Validate aggregators (console.log)
    validateAllAggregators();
    
    // Update data status
    const dataStatus = document.getElementById('dataStatus');
    if (dataStatus) {
        dataStatus.textContent = `Data berhasil dimuat: ${dashboardData.parsed.length} records`;
        dataStatus.style.color = '#10b981';
    }
    
    // Initialize dashboard (with delay to ensure DOM is ready)
    setTimeout(() => {
        initializeDashboard();
    }, 100);
}

function parseData(headers, rows) {
    // Find column indices
    const findColumn = (name) => {
        return headers.findIndex(h => h === name || h.trim() === name.trim());
    };
    
    const columnMap = {
        rfqNumber: findColumn(COLUMN_MAP.rfqNumber),
        date: findColumn(COLUMN_MAP.date),
        sales: findColumn(COLUMN_MAP.sales),
        customer: findColumn(COLUMN_MAP.customer),
        hargaNew: findColumn(COLUMN_MAP.hargaNew),
        status: findColumn(COLUMN_MAP.status),
        total: findColumn(COLUMN_MAP.total)
    };
    
    return rows.map(row => {
        const parseNumber = (val) => {
            if (!val || val === '') return 0;
            const str = String(val).trim();
            if (!str || str === '') return 0;
            // Remove non-numeric chars except dots and commas
            const cleaned = str.replace(/[^\d,\.]/g, '').replace(/,/g, '');
            const num = parseFloat(cleaned);
            return isNaN(num) ? 0 : num;
        };
        
        const parseDate = (val) => {
            if (!val) return null;
            const str = String(val).trim();
            if (!str) return null;
            
            // Try to parse as Excel date number
            if (typeof val === 'number') {
                // Excel dates are days since 1900-01-01
                const excelEpoch = new Date(1900, 0, 1);
                excelEpoch.setDate(excelEpoch.getDate() + val - 2); // -2 because Excel's epoch bug
                return excelEpoch;
            }
            
            // Try parsing date string formats
            const dateStr = str;
            // Format: DD-MMM-YY or DD/MM/YYYY or YYYY-MM-DD
            let parsedDate = new Date(dateStr);
            if (!isNaN(parsedDate.getTime())) {
                return parsedDate;
            }
            
            // Try DD-MMM-YY format
            const parts = dateStr.split('-');
            if (parts.length === 3) {
                const day = parseInt(parts[0]);
                const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
                const month = monthNames.indexOf(parts[1]);
                const year = 2000 + parseInt(parts[2]);
                if (month >= 0 && !isNaN(day) && !isNaN(year)) {
                    return new Date(year, month, day);
                }
            }
            
            return null;
        };
        
        const status = String(row[columnMap.status] || '').trim().toUpperCase();
        // Konversi = KETERANGAN != "TIDAK JADI OC"
        const isConverted = status !== 'TIDAK JADI OC' && status !== '';
        
        const hargaNew = parseNumber(row[columnMap.hargaNew]);
        const total = parseNumber(row[columnMap.total]);
        // Use TOTAL if available, otherwise HARGA (NEW)
        const amount = total > 0 ? total : hargaNew;
        
        return {
            date: parseDate(row[columnMap.date]),
            rfqNumber: String(row[columnMap.rfqNumber] || '').trim(),
            customer: String(row[columnMap.customer] || '').trim(),
            sales: String(row[columnMap.sales] || '').trim(),
            hargaNew: hargaNew,
            total: total,
            amount: amount,
            status: status,
            isConverted: isConverted,
            rawRow: row // Keep original row data
        };
    }).filter(item => item.customer || item.sales); // Filter out completely empty rows
}

// ============================================
// Preview Section
// ============================================

function showPreview(headers, previewRows) {
    const previewSection = document.getElementById('previewSection');
    const columnInfo = document.getElementById('columnInfo');
    const previewTableHead = document.getElementById('previewTableHead');
    const previewTableBody = document.getElementById('previewTableBody');
    
    if (!previewSection) return;
    
    previewSection.style.display = 'block';
    
    // Show column names
    columnInfo.innerHTML = `
        <h4>Nama Kolom (${headers.length} kolom):</h4>
        <ul>
            ${headers.map(h => `<li>${h || '(kosong)'}</li>`).join('')}
        </ul>
    `;
    
    // Show preview table
    previewTableHead.innerHTML = `
        <tr>
            ${headers.map(h => `<th>${h || ''}</th>`).join('')}
        </tr>
    `;
    
    previewTableBody.innerHTML = previewRows.map(row => {
        return `
            <tr>
                ${headers.map((_, idx) => `<td>${row[idx] || ''}</td>`).join('')}
            </tr>
        `;
    }).join('');
}

// ============================================
// Data Filtering
// ============================================

function filterData(filters = {}) {
    let filtered = [...dashboardData.parsed];
    
    if (filters.customer && filters.customer.trim() !== '') {
        filtered = filtered.filter(d => 
            d.customer && d.customer.toLowerCase().includes(filters.customer.toLowerCase())
        );
    }
    
    if (filters.sales && filters.sales.trim() !== '') {
        const salesFilterValue = filters.sales.trim();
        filtered = filtered.filter(d => {
            if (!d.sales) return false;
            // Exact match (case-insensitive, trimmed)
            return d.sales.trim().toLowerCase() === salesFilterValue.toLowerCase();
        });
    }
    
    if (filters.status) {
        filtered = filtered.filter(d => 
            d.status.toLowerCase().includes(filters.status.toLowerCase())
        );
    }
    
    if (filters.dateFrom) {
        const fromDate = new Date(filters.dateFrom);
        filtered = filtered.filter(d => {
            if (!d.date) return false;
            return d.date >= fromDate;
        });
    }
    
    if (filters.dateTo) {
        const toDate = new Date(filters.dateTo);
        toDate.setHours(23, 59, 59, 999); // End of day
        filtered = filtered.filter(d => {
            if (!d.date) return false;
            return d.date <= toDate;
        });
    }
    
    return filtered;
}

function getFilterValues(widgetId) {
    const filters = {};
    
    const customerFilter = document.getElementById(`filter${widgetId}-customer`);
    const salesFilter = document.getElementById(`filter${widgetId}-sales`);
    const dateFromFilter = document.getElementById(`filter${widgetId}-date-from`);
    const dateToFilter = document.getElementById(`filter${widgetId}-date-to`);
    
    if (customerFilter && customerFilter.value) filters.customer = customerFilter.value;
    if (salesFilter && salesFilter.value) filters.sales = salesFilter.value;
    if (dateFromFilter && dateFromFilter.value) filters.dateFrom = dateFromFilter.value;
    if (dateToFilter && dateToFilter.value) filters.dateTo = dateToFilter.value;
    
    return filters;
}

// ============================================
// Populate Filter Options
// ============================================

function populateFilterOptions() {
    if (!dashboardData.parsed || dashboardData.parsed.length === 0) return;
    
    const dateFromInput = document.getElementById('rfq-customer-sales-date-from');
    const dateToInput = document.getElementById('rfq-customer-sales-date-to');
    
    const dates = dashboardData.parsed
        .map(d => d.date)
        .filter(d => d instanceof Date && !isNaN(d.getTime()));
    
    const formatDateForInput = (date) => {
        if (!date) return '';
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    };
    
    if (dates.length > 0) {
        const minDate = new Date(Math.min(...dates.map(d => d.getTime())));
        const maxDate = new Date(Math.max(...dates.map(d => d.getTime())));
        const minDateStr = formatDateForInput(minDate);
        const maxDateStr = formatDateForInput(maxDate);
    
        if (dateFromInput) {
            if (!dateFromInput.value) dateFromInput.value = minDateStr;
            dateFromInput.setAttribute('min', minDateStr);
            dateFromInput.setAttribute('max', maxDateStr);
        }
        
        if (dateToInput) {
            if (!dateToInput.value) dateToInput.value = maxDateStr;
            dateToInput.setAttribute('min', minDateStr);
            dateToInput.setAttribute('max', maxDateStr);
        }
    }
    
    // Rebuild dropdown options so the widget immediately shows correlated counts
    try {
        refreshRFQLinkedDropdowns({ preserveSelections: false });
    } catch (error) {
        console.error('Failed to refresh RFQ dropdowns:', error);
    }
    
    const salesSelect = document.getElementById('rfq-customer-sales-sales');
    const customerSelect = document.getElementById('rfq-customer-sales-customer');
    if (salesSelect) rfqFilterState.sales = (salesSelect.value || '').trim();
    if (customerSelect) rfqFilterState.customer = (customerSelect.value || '').trim();
}

// ============================================
// KPI Calculations
// ============================================

function calculateKPIs(data) {
    const totalRFQ = data.length;
    const converted = data.filter(d => d.isConverted).length;
    const conversionRate = totalRFQ > 0 ? (converted / totalRFQ) * 100 : 0;
    
    return {
        totalRFQ,
        converted,
        conversionRate
    };
}

function formatCurrency(value) {
    if (value >= 1000000000) {
        return (value / 1000000000).toFixed(2) + 'B';
    } else if (value >= 1000000) {
        return (value / 1000000).toFixed(1) + 'M';
    } else if (value >= 1000) {
        return (value / 1000).toFixed(1) + 'K';
    }
    return value.toLocaleString('id-ID');
}

function updateKPICards() {
    const kpis = calculateKPIs(dashboardData.parsed);
    
    // Summary KPIs
    document.getElementById('kpiTotalRFQ').textContent = kpis.totalRFQ.toLocaleString('id-ID');
    document.getElementById('kpiConversion').textContent = kpis.conversionRate.toFixed(2) + '%';
    document.getElementById('kpiConverted').textContent = kpis.converted.toLocaleString('id-ID');
    
    // Update KPI Card 1: RFQ per Customer per Sales
    const kpi1Data = aggregateRFQPerCustomerSales(dashboardData.parsed);
    const kpi1Total = kpi1Data.length;
    const kpi1Top = kpi1Data[0];
    
    const kpi1TotalEl = document.getElementById('kpi1TotalCombinations');
    const kpi1TopCombinationEl = document.getElementById('kpi1TopCombination');
    const kpi1TopCountEl = document.getElementById('kpi1TopCount');
    
    if (kpi1TotalEl) kpi1TotalEl.textContent = kpi1Total.toLocaleString('id-ID');
    if (kpi1TopCombinationEl && kpi1Top) {
        const label = `${kpi1Top.customer} - ${kpi1Top.sales}`;
        kpi1TopCombinationEl.textContent = label.length > 30 ? label.substring(0, 27) + '...' : label;
    }
    if (kpi1TopCountEl && kpi1Top) kpi1TopCountEl.textContent = kpi1Top.count.toLocaleString('id-ID');
    
    // Update KPI Card 2: Top Customers
    const kpi2Data = aggregateTopCustomers(dashboardData.parsed);
    const kpi2Total = kpi2Data.length;
    const kpi2Top = kpi2Data[0];
    
    const kpi2TotalEl = document.getElementById('kpi2TotalCustomers');
    const kpi2TopCustomerEl = document.getElementById('kpi2TopCustomer');
    const kpi2TopCountEl = document.getElementById('kpi2TopCount');
    
    if (kpi2TotalEl) kpi2TotalEl.textContent = kpi2Total.toLocaleString('id-ID');
    if (kpi2TopCustomerEl && kpi2Top) {
        kpi2TopCustomerEl.textContent = kpi2Top.customer.length > 25 ? kpi2Top.customer.substring(0, 22) + '...' : kpi2Top.customer;
    }
    if (kpi2TopCountEl && kpi2Top) kpi2TopCountEl.textContent = kpi2Top.count.toLocaleString('id-ID');
    
    // Update KPI Card 3: Conversion per Customer
    const kpi3Data = aggregateConversionPerCustomer(dashboardData.parsed);
    const kpi3Avg = kpi3Data.length > 0 
        ? kpi3Data.reduce((sum, r) => sum + r.conversionRate, 0) / kpi3Data.length 
        : 0;
    const kpi3Best = kpi3Data[0];
    
    const kpi3AvgEl = document.getElementById('kpi3AvgConversion');
    const kpi3BestCustomerEl = document.getElementById('kpi3BestCustomer');
    const kpi3BestRateEl = document.getElementById('kpi3BestRate');
    
    if (kpi3AvgEl) kpi3AvgEl.textContent = kpi3Avg.toFixed(2) + '%';
    if (kpi3BestCustomerEl && kpi3Best) {
        kpi3BestCustomerEl.textContent = kpi3Best.customer.length > 25 ? kpi3Best.customer.substring(0, 22) + '...' : kpi3Best.customer;
    }
    if (kpi3BestRateEl && kpi3Best) kpi3BestRateEl.textContent = kpi3Best.conversionRate.toFixed(2) + '%';
    
    // Update KPI Card 4: Amount per Customer per Sales
    const kpi4Data = aggregateAmountPerCustomerSales(dashboardData.parsed);
    const kpi4Total = kpi4Data.reduce((sum, r) => sum + r.amount, 0);
    const kpi4Top = kpi4Data[0];
    
    const kpi4TotalEl = document.getElementById('kpi4TotalAmount');
    const kpi4TopCombinationEl = document.getElementById('kpi4TopCombination');
    const kpi4TopAmountEl = document.getElementById('kpi4TopAmount');
    
    if (kpi4TotalEl) kpi4TotalEl.textContent = 'Rp ' + formatCurrency(kpi4Total);
    if (kpi4TopCombinationEl && kpi4Top) {
        const label = `${kpi4Top.customer} - ${kpi4Top.sales}`;
        kpi4TopCombinationEl.textContent = label.length > 30 ? label.substring(0, 27) + '...' : label;
    }
    if (kpi4TopAmountEl && kpi4Top) kpi4TopAmountEl.textContent = 'Rp ' + formatCurrency(kpi4Top.amount);
}

// ============================================
// Aggregator Functions for KPI Validation
// ============================================

/**
 * Aggregator 1: Jumlah RFQ per Customer per Sales
 * Menghitung total jumlah RFQ yang dikelompokkan per customer dan sales
 */
function aggregateRFQPerCustomerSales(data) {
    const grouped = {};
    
    data.forEach(item => {
        if (!item.customer || !item.sales) return;
        const key = `${item.customer}|${item.sales}`;
        if (!grouped[key]) {
            grouped[key] = {
                customer: item.customer,
                sales: item.sales,
                count: 0
            };
        }
        grouped[key].count++;
    });
    
    const result = Object.values(grouped)
        .sort((a, b) => b.count - a.count);
    
    console.log('=== AGGREGATOR 1: Jumlah RFQ per Customer per Sales ===');
    console.log('Total kombinasi customer-sales:', result.length);
    console.log('Top 15:', result.slice(0, 15));
    console.log('Detail lengkap:', result);
    console.log('');
    
    return result;
}

/**
 * Aggregator 2: Customer dengan RFQ Terbanyak
 * Menghitung total RFQ per customer (tanpa memisahkan per sales)
 */
function aggregateTopCustomers(data) {
    const customerData = {};
    
    data.forEach(item => {
        if (!item.customer) return;
        if (!customerData[item.customer]) {
            customerData[item.customer] = 0;
        }
        customerData[item.customer]++;
    });
    
    const result = Object.entries(customerData)
        .map(([customer, count]) => ({
            customer,
            count
        }))
        .sort((a, b) => b.count - a.count);
    
    console.log('=== AGGREGATOR 2: Customer dengan RFQ Terbanyak ===');
    console.log('Total customer:', result.length);
    console.log('Top 10:', result.slice(0, 10));
    console.log('Detail lengkap:', result);
    console.log('');
    
    return result;
}

/**
 * Aggregator 3: % Konversi RFQ per Customer
 * Menghitung persentase konversi (JADI OC / Total) per customer
 */
function aggregateConversionPerCustomer(data) {
    const customerData = {};
    
    data.forEach(item => {
        if (!item.customer) return;
        if (!customerData[item.customer]) {
            customerData[item.customer] = {
                customer: item.customer,
                total: 0,
                converted: 0
            };
        }
        customerData[item.customer].total++;
        if (item.isConverted) {
            customerData[item.customer].converted++;
        }
    });
    
    const result = Object.values(customerData)
        .map(item => ({
            customer: item.customer,
            total: item.total,
            converted: item.converted,
            notConverted: item.total - item.converted,
            conversionRate: item.total > 0 ? (item.converted / item.total) * 100 : 0
        }))
        .sort((a, b) => b.conversionRate - a.conversionRate);
    
    console.log('=== AGGREGATOR 3: % Konversi RFQ per Customer ===');
    console.log('Total customer:', result.length);
    console.log('Top 10 (by conversion rate):', result.slice(0, 10));
    console.log('Summary:', {
        totalCustomers: result.length,
        customersWithConversion: result.filter(r => r.converted > 0).length,
        customersWithoutConversion: result.filter(r => r.converted === 0).length,
        avgConversionRate: result.reduce((sum, r) => sum + r.conversionRate, 0) / result.length
    });
    console.log('Detail lengkap:', result);
    console.log('');
    
    return result;
}

/**
 * Aggregator 4: Total Amount RFQ per Customer per Sales
 * Menghitung total amount (TOTAL atau HARGA (NEW)) per customer dan sales
 */
function aggregateAmountPerCustomerSales(data) {
    const grouped = {};
    
    data.forEach(item => {
        if (!item.customer || !item.sales) return;
        const key = `${item.customer}|${item.sales}`;
        if (!grouped[key]) {
            grouped[key] = {
                customer: item.customer,
                sales: item.sales,
                amount: 0,
                count: 0
            };
        }
        grouped[key].amount += item.amount;
        grouped[key].count++;
    });
    
    const result = Object.values(grouped)
        .sort((a, b) => b.amount - a.amount);
    
    console.log('=== AGGREGATOR 4: Total Amount RFQ per Customer per Sales ===');
    console.log('Total kombinasi customer-sales:', result.length);
    console.log('Top 15 (by amount):', result.slice(0, 15).map(r => ({
        customer: r.customer,
        sales: r.sales,
        amount: r.amount,
        count: r.count,
        avgAmount: r.amount / r.count
    })));
    console.log('Summary:', {
        totalCombinations: result.length,
        totalAmount: result.reduce((sum, r) => sum + r.amount, 0),
        avgAmount: result.reduce((sum, r) => sum + r.amount, 0) / result.length,
        maxAmount: Math.max(...result.map(r => r.amount)),
        minAmount: Math.min(...result.map(r => r.amount))
    });
    console.log('Detail lengkap:', result);
    console.log('');
    
    return result;
}

/**
 * Function untuk memvalidasi semua aggregator sekaligus
 * Dipanggil setelah data di-load
 */
function validateAllAggregators() {
    if (!dashboardData.parsed || dashboardData.parsed.length === 0) {
        console.warn('Tidak ada data untuk divalidasi');
        return;
    }
    
    console.log('========================================');
    console.log('VALIDASI AGGREGATOR - SEMUA DATA');
    console.log('Total records:', dashboardData.parsed.length);
    console.log('========================================');
    console.log('');
    
    // Validasi tanpa filter
    aggregateRFQPerCustomerSales(dashboardData.parsed);
    aggregateTopCustomers(dashboardData.parsed);
    aggregateConversionPerCustomer(dashboardData.parsed);
    aggregateAmountPerCustomerSales(dashboardData.parsed);
    
    console.log('========================================');
    console.log('VALIDASI SELESAI');
    console.log('========================================');
}

// ============================================
// Widget: RFQ per Customer per Sales (Stacked Bar Chart)
// ============================================

// Store selected customer and sales for click interaction
let selectedCustomerSales = {
    customer: null,
    sales: null
};

// Store axis orientation: 'customer-sales' (X=Customer, stacked by Sales) or 'sales-customer' (X=Sales, stacked by Customer)
let chartAxisOrientation = 'customer-sales';

// Store simple mode state: when true, only highlight the biggest value in each stack
let chartSimpleMode = false;

// Store show all state: when true, show all customers/sales; when false, show top N
let chartShowAll = false;
// Store custom show count for RFQ Customer Sales widget
let rfqCustomerSalesShowCount = 15;
// Store custom show count for Top Customer RFQ Volume widget
let topCustomerRFQVolumeShowCount = 20;

function createRFQCustomerSalesChart() {
    const ctx = document.getElementById('chartRFQCustomerSales');
    if (!ctx || !dashboardData.parsed || dashboardData.parsed.length === 0) return;
    
    // Get filters
    const dateFromInput = document.getElementById('rfq-customer-sales-date-from');
    const dateToInput = document.getElementById('rfq-customer-sales-date-to');
    const customerFilter = document.getElementById('rfq-customer-sales-customer');
    const salesFilter = document.getElementById('rfq-customer-sales-sales');
    const statusJadiOC = document.getElementById('rfq-customer-sales-status-jadi-oc');
    const statusTidakJadiOC = document.getElementById('rfq-customer-sales-status-tidak-jadi-oc');
    
    const filters = {
        dateFrom: dateFromInput?.value || '',
        dateTo: dateToInput?.value || '',
        customer: customerFilter?.value?.trim() || '',
        sales: salesFilter?.value?.trim() || ''
    };
    
    // Debug: Log filter values
    const debugBothChecked = statusJadiOC?.checked && statusTidakJadiOC?.checked;
    if (debugBothChecked) {
        console.log('=== FILTER VALUES ===');
        console.log('Sales filter:', filters.sales || '(all)');
        console.log('Customer filter:', filters.customer || '(all)');
        console.log('Date from:', filters.dateFrom || '(all)');
        console.log('Date to:', filters.dateTo || '(all)');
    }
    
    // Filter data
    let filteredData = filterData(filters);
    
    // Apply status filter (JADI OC / TIDAK JADI OC)
    const includeJadiOC = statusJadiOC?.checked || false;
    const includeTidakJadiOC = statusTidakJadiOC?.checked || false;
    
    // Store original count before status filtering
    const countBeforeStatusFilter = filteredData.length;
    
    if (includeJadiOC && includeTidakJadiOC) {
        // Both selected - show ALL data (no status filter applied)
        // filteredData already contains all data after other filters
        // Don't apply any status filter - show everything
    } else if (includeJadiOC && !includeTidakJadiOC) {
        // Only JADI OC - show only converted
        filteredData = filteredData.filter(d => d.isConverted === true);
    } else if (!includeJadiOC && includeTidakJadiOC) {
        // Only TIDAK JADI OC - show only not converted
        filteredData = filteredData.filter(d => d.isConverted === false);
    } else {
        // Neither selected - should not happen (we prevent this), but show nothing
        filteredData = [];
    }
    
    // Debug: Log counts - verify no filtering when both checked
    if (includeJadiOC && includeTidakJadiOC) {
        console.log('=== BOTH STATUSES CHECKED ===');
        console.log('Count before status filter:', countBeforeStatusFilter);
        console.log('Count after status filter (should be same):', filteredData.length);
        console.log('Match:', filteredData.length === countBeforeStatusFilter ? '✓ YES' : '✗ NO - DIFFERENCE: ' + Math.abs(filteredData.length - countBeforeStatusFilter));
        
        // Check specific sales if filtered
        if (filters.sales) {
            const salesFilterValue = filters.sales.trim().toLowerCase();
            const salesExactCount = filteredData.filter(d => {
                const sales = d.sales ? d.sales.trim().toLowerCase() : '';
                return sales === salesFilterValue;
            }).length;
            const salesPartialCount = filteredData.filter(d => {
                const sales = d.sales ? d.sales.toLowerCase() : '';
                return sales.includes(salesFilterValue);
            }).length;
            
            console.log(`Filtered for sales "${filters.sales}":`);
            console.log(`  - Exact match (normalized): ${salesExactCount} rows`);
            console.log(`  - Partial match: ${salesPartialCount} rows`);
            
            // Show sample of sales values
            const uniqueSales = [...new Set(filteredData.map(d => d.sales ? d.sales.trim() : '').filter(s => s))];
            console.log(`  - Unique sales in filtered data (first 10):`, uniqueSales.slice(0, 10));
            
            // Verify all FAJAR rows if that's what we're filtering
            if (salesFilterValue === 'fajar') {
                const allFajar = dashboardData.parsed.filter(d => {
                    const sales = d.sales ? d.sales.trim().toLowerCase() : '';
                    return sales === 'fajar';
                });
                console.log(`  - Total FAJAR in all data: ${allFajar.length}`);
                console.log(`  - FAJAR with customer: ${allFajar.filter(d => d.customer && d.customer.trim()).length}`);
                console.log(`  - FAJAR isConverted=true: ${allFajar.filter(d => d.isConverted).length}`);
                console.log(`  - FAJAR isConverted=false: ${allFajar.filter(d => !d.isConverted).length}`);
            }
        }
    }
    
    // Group by customer, then by sales
    const customerSalesMap = {};
    const allSales = new Set();
    let totalCountedInChart = 0;
    let rowsSkipped = 0;
    
    filteredData.forEach(item => {
        // Normalize customer and sales (trim whitespace)
        const customer = item.customer ? item.customer.trim() : '';
        const sales = item.sales ? item.sales.trim() : '';
        
        // Skip rows without customer or sales (they can't be grouped in the chart)
        if (!customer || !sales) {
            rowsSkipped++;
            return;
        }
        
        if (!customerSalesMap[customer]) {
            customerSalesMap[customer] = {};
        }
        
        if (!customerSalesMap[customer][sales]) {
            customerSalesMap[customer][sales] = 0;
        }
        
        customerSalesMap[customer][sales]++;
        allSales.add(sales);
        totalCountedInChart++;
    });
    
    // Convert to array and sort customers by total RFQ
    const allCustomers = Object.keys(customerSalesMap).sort((a, b) => {
        const totalA = Object.values(customerSalesMap[a]).reduce((sum, val) => sum + val, 0);
        const totalB = Object.values(customerSalesMap[b]).reduce((sum, val) => sum + val, 0);
        return totalB - totalA;
    });
    
    // Check if there's no data after filtering
    if (allCustomers.length === 0 || filteredData.length === 0) {
        // Destroy existing chart
        if (charts.rfqCustomerSales) {
            charts.rfqCustomerSales.destroy();
        }
        
        // Create empty chart with message
        charts.rfqCustomerSales = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: [],
                datasets: []
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: { enabled: false }
                },
                scales: {
                    x: { display: false },
                    y: { display: false }
                }
            }
        });
        
        // Update summary to show empty message
        const summaryEl = document.getElementById('rfq-customer-sales-summary');
        if (summaryEl) {
            let emptyMessage = 'No data available';
            if (filters.sales && !includeTidakJadiOC && includeJadiOC) {
                emptyMessage = `No JADI OC records found for sales: ${filters.sales}`;
            } else if (filters.sales && !includeJadiOC && includeTidakJadiOC) {
                emptyMessage = `No TIDAK JADI OC records found for sales: ${filters.sales}`;
            } else if (filters.sales) {
                emptyMessage = `No records found for sales: ${filters.sales}`;
            } else if (!includeTidakJadiOC && includeJadiOC) {
                emptyMessage = 'No JADI OC records found';
            } else if (!includeJadiOC && includeTidakJadiOC) {
                emptyMessage = 'No TIDAK JADI OC records found';
            }
            summaryEl.querySelector('.summary-text').textContent = emptyMessage;
        }
        
        // Hide table if shown
        const tableContainer = document.getElementById('rfq-customer-sales-table-container');
        if (tableContainer) tableContainer.style.display = 'none';
        
        return;
    }
    
    // Calculate total sum across all customers (not just top 15)
    const totalSumAllCustomers = allCustomers.reduce((sum, customer) => {
        return sum + Object.values(customerSalesMap[customer]).reduce((customerSum, val) => customerSum + val, 0);
    }, 0);
    
    // Debug: Verify counts match when both statuses are checked
    if (includeJadiOC && includeTidakJadiOC) {
        console.log('=== DEBUG: Both Statuses Checked ===');
        console.log('Total filtered data rows:', filteredData.length);
        console.log('Total counted in chart (with customer/sales):', totalCountedInChart);
        console.log('Rows skipped (no customer/sales):', rowsSkipped);
        console.log('Total sum across all customers:', totalSumAllCustomers);
        console.log('Expected sum (filteredData.length - rowsSkipped):', filteredData.length - rowsSkipped);
        console.log('Match:', totalSumAllCustomers === (filteredData.length - rowsSkipped) ? 'YES ✓' : 'NO ✗ - DIFF: ' + Math.abs(totalSumAllCustomers - (filteredData.length - rowsSkipped)));
        
        // If sales filter is applied, verify count for that sales
        if (filters.sales) {
            const salesFilterValue = filters.sales.trim().toLowerCase();
            const salesDataCount = filteredData.filter(d => {
                const sales = d.sales ? d.sales.trim().toLowerCase() : '';
                const customer = d.customer ? d.customer.trim() : '';
                return sales === salesFilterValue && customer;
            }).length;
            
            // Calculate chart count - sum across all customers for this sales
            let salesChartCount = 0;
            Object.keys(customerSalesMap).forEach(customer => {
                Object.keys(customerSalesMap[customer]).forEach(salesKey => {
                    if (salesKey.trim().toLowerCase() === salesFilterValue) {
                        salesChartCount += customerSalesMap[customer][salesKey];
                    }
                });
            });
            
            console.log(`=== Sales "${filters.sales}" Verification ===`);
            console.log(`Filtered data rows with customer/sales: ${salesDataCount}`);
            console.log(`Total in chart (sum across ALL customers): ${salesChartCount}`);
            console.log(`Expected: ${salesDataCount}`);
            console.log(`Match: ${salesChartCount === salesDataCount ? '✓ YES' : '✗ NO - DIFF: ' + Math.abs(salesChartCount - salesDataCount)}`);
            
            // Breakdown by status
            const jadiOC = filteredData.filter(d => {
                const sales = d.sales ? d.sales.trim().toLowerCase() : '';
                return sales === salesFilterValue && d.isConverted;
            }).length;
            const tidakJadiOC = filteredData.filter(d => {
                const sales = d.sales ? d.sales.trim().toLowerCase() : '';
                return sales === salesFilterValue && !d.isConverted;
            }).length;
            console.log(`  JADI OC: ${jadiOC}, TIDAK JADI OC: ${tidakJadiOC}, Total: ${jadiOC + tidakJadiOC}`);
            console.log(`  (Should equal ${salesDataCount})`);
        }
        console.log('=====================================');
    }
    
    // Determine chart orientation
    const isCustomerSales = chartAxisOrientation === 'customer-sales';
    
    let labels, datasets, xAxisLabel, stackedByLabel;
    
    if (isCustomerSales) {
        // X-axis = Customer, stacked by Sales
        // Show top N or all customers based on toggle and custom count
        const showCountInput = document.getElementById('rfq-customer-sales-show-count');
        const showCount = showCountInput ? parseInt(showCountInput.value) || 15 : 15;
        rfqCustomerSalesShowCount = showCount;
        const customers = chartShowAll ? allCustomers : allCustomers.slice(0, showCount);
        const salesArray = Array.from(allSales).sort();
        
        // Generate colors for sales
        const colorConfigs = generateSalesColors(salesArray.length);
        window.salesColorConfigs = colorConfigs;
        
        const salesToShow = filters.sales ? [filters.sales.trim()] : salesArray;
        
        labels = customers.map(c => c.length > 30 ? c.substring(0, 27) + '...' : c);
        
        // Calculate data first
        const dataArrays = salesToShow.map((sales, index) => {
            const colorIndex = filters.sales ? salesArray.indexOf(sales) : index;
            const colorConfig = colorIndex >= 0 ? colorConfigs[colorIndex % colorConfigs.length] : colorConfigs[0];
            
            return {
                sales: sales,
                data: customers.map(customer => {
                    const matchingSalesKey = Object.keys(customerSalesMap[customer] || {}).find(s => {
                        return s.trim().toLowerCase() === sales.trim().toLowerCase();
                    });
                    return matchingSalesKey ? customerSalesMap[customer][matchingSalesKey] : 0;
                }),
                colorConfig: colorConfig
            };
        });
        
        // Apply simple mode highlighting if enabled
        datasets = dataArrays.map((item, index) => {
            let backgroundColor, borderColor, borderWidth;
            
            if (chartSimpleMode) {
                // Find the biggest value for each customer across all sales
                backgroundColor = item.data.map((value, customerIndex) => {
                    const allValuesForCustomer = dataArrays.map(d => d.data[customerIndex]);
                    const maxValue = Math.max(...allValuesForCustomer);
                    const isMax = value === maxValue && value > 0;
                    
                    if (isMax) {
                        // Highlight the biggest value
                        return item.colorConfig.highlight || item.colorConfig.base.replace('0.85', '1');
                    } else {
                        // Dim other values
                        return item.colorConfig.base.replace('0.85', '0.2');
                    }
                });
                
                borderColor = item.data.map((value, customerIndex) => {
                    const allValuesForCustomer = dataArrays.map(d => d.data[customerIndex]);
                    const maxValue = Math.max(...allValuesForCustomer);
                    const isMax = value === maxValue && value > 0;
                    
                    return isMax ? (item.colorConfig.highlight || item.colorConfig.base.replace('0.8', '1')) : 'transparent';
                });
                
                borderWidth = item.data.map((value, customerIndex) => {
                    const allValuesForCustomer = dataArrays.map(d => d.data[customerIndex]);
                    const maxValue = Math.max(...allValuesForCustomer);
                    const isMax = value === maxValue && value > 0;
                    
                    return isMax ? 2 : 0;
                });
            } else {
                // Normal mode: full colors
                backgroundColor = item.colorConfig.base;
                borderColor = item.colorConfig.base.replace('0.8', '1');
                borderWidth = 1;
            }
            
            return {
                label: item.sales,
                data: item.data,
                backgroundColor: backgroundColor,
                borderColor: borderColor,
                borderWidth: borderWidth,
                originalColor: item.colorConfig.base,
                highlightColor: item.colorConfig.highlight
            };
        });
        xAxisLabel = 'Customer';
        stackedByLabel = 'Sales';
        
        // Store for click interaction
        window.currentChartData = { xAxisValues: customers, stackedValues: salesToShow, orientation: 'customer-sales' };
    } else {
        // X-axis = Sales, stacked by Customer
        // Create salesCustomerMap: { sales: { customer: count } }
        const salesCustomerMap = {};
        filteredData.forEach(item => {
            const customer = item.customer ? item.customer.trim() : '';
            const sales = item.sales ? item.sales.trim() : '';
            if (!customer || !sales) return;
            
            if (!salesCustomerMap[sales]) {
                salesCustomerMap[sales] = {};
            }
            if (!salesCustomerMap[sales][customer]) {
                salesCustomerMap[sales][customer] = 0;
            }
            salesCustomerMap[sales][customer]++;
        });
        
        // Get all sales and sort
        const allSalesSorted = Object.keys(salesCustomerMap).sort((a, b) => {
            const totalA = Object.values(salesCustomerMap[a] || {}).reduce((sum, val) => sum + val, 0);
            const totalB = Object.values(salesCustomerMap[b] || {}).reduce((sum, val) => sum + val, 0);
            return totalB - totalA;
        });
        
        // Show top N or all sales based on toggle and custom count
        const showCountInput = document.getElementById('rfq-customer-sales-show-count');
        const showCount = showCountInput ? parseInt(showCountInput.value) || 15 : 15;
        rfqCustomerSalesShowCount = showCount;
        const sales = chartShowAll ? allSalesSorted : allSalesSorted.slice(0, showCount);
        
        // Store for summary calculation
        window.allSalesSorted = allSalesSorted;
        window.currentSalesCount = sales.length;
        
        // Get all customers for this sales set
        const allCustomersForSales = new Set();
        sales.forEach(s => {
            Object.keys(salesCustomerMap[s] || {}).forEach(c => allCustomersForSales.add(c));
        });
        const customersArray = Array.from(allCustomersForSales).sort();
        
        // Generate colors for customers
        const colorConfigs = generateSalesColors(customersArray.length);
        window.customerColorConfigs = colorConfigs;
        
        const salesToShow = filters.sales ? [filters.sales.trim()] : sales;
        
        labels = salesToShow.map(s => s.length > 30 ? s.substring(0, 27) + '...' : s);
        
        // Calculate data first
        const dataArrays = customersArray.map((customer, index) => {
            const colorConfig = colorConfigs[index % colorConfigs.length];
            
            return {
                customer: customer,
                data: salesToShow.map(sales => {
                    return salesCustomerMap[sales]?.[customer] || 0;
                }),
                colorConfig: colorConfig
            };
        });
        
        // Apply simple mode highlighting if enabled
        datasets = dataArrays.map((item, index) => {
            let backgroundColor, borderColor, borderWidth;
            
            if (chartSimpleMode) {
                // Find the biggest value for each sales across all customers
                backgroundColor = item.data.map((value, salesIndex) => {
                    const allValuesForSales = dataArrays.map(d => d.data[salesIndex]);
                    const maxValue = Math.max(...allValuesForSales);
                    const isMax = value === maxValue && value > 0;
                    
                    if (isMax) {
                        // Highlight the biggest value
                        return item.colorConfig.highlight || item.colorConfig.base.replace('0.85', '1');
                    } else {
                        // Dim other values
                        return item.colorConfig.base.replace('0.85', '0.2');
                    }
                });
                
                borderColor = item.data.map((value, salesIndex) => {
                    const allValuesForSales = dataArrays.map(d => d.data[salesIndex]);
                    const maxValue = Math.max(...allValuesForSales);
                    const isMax = value === maxValue && value > 0;
                    
                    return isMax ? (item.colorConfig.highlight || item.colorConfig.base.replace('0.8', '1')) : 'transparent';
                });
                
                borderWidth = item.data.map((value, salesIndex) => {
                    const allValuesForSales = dataArrays.map(d => d.data[salesIndex]);
                    const maxValue = Math.max(...allValuesForSales);
                    const isMax = value === maxValue && value > 0;
                    
                    return isMax ? 2 : 0;
                });
            } else {
                // Normal mode: full colors
                backgroundColor = item.colorConfig.base;
                borderColor = item.colorConfig.base.replace('0.8', '1');
                borderWidth = 1;
            }
            
            return {
                label: item.customer.length > 30 ? item.customer.substring(0, 27) + '...' : item.customer,
                data: item.data,
                backgroundColor: backgroundColor,
                borderColor: borderColor,
                borderWidth: borderWidth,
                originalColor: item.colorConfig.base,
                highlightColor: item.colorConfig.highlight
            };
        });
        xAxisLabel = 'Sales';
        stackedByLabel = 'Customer';
        
        // Store for click interaction
        window.currentChartData = { xAxisValues: salesToShow, stackedValues: customersArray, orientation: 'sales-customer' };
    }
    
    // Destroy existing chart
    if (charts.rfqCustomerSales) {
        charts.rfqCustomerSales.destroy();
    }
    
    // Create stacked bar chart
    charts.rfqCustomerSales = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            resizeDelay: 100,
            animation: false,
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: {
                        usePointStyle: true,
                        padding: 12,
                        font: { size: 11 },
                        color: CONFIG.textColor
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.85)',
                    padding: 10,
                    titleFont: { size: 12 },
                    bodyFont: { size: 11 },
                    callbacks: {
                        label: function(context) {
                            const label = context.dataset.label || '';
                            const value = context.parsed.y || 0;
                            return `${label}: ${value} RFQ`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    stacked: true,
                    grid: {
                        display: false
                    },
                    ticks: {
                        font: { size: 11 },
                        color: CONFIG.textColor,
                        maxRotation: 45,
                        minRotation: 45,
                        autoSkip: true,
                        maxTicksLimit: 20
                    },
                    categoryPercentage: 0.85,
                    barPercentage: 0.95,
                    maxBarThickness: 50
                },
                y: {
                    stacked: true,
                    grid: {
                        display: true,
                        color: CONFIG.gridColor,
                        lineWidth: 0.5
                    },
                    ticks: {
                        font: { size: 11 },
                        color: CONFIG.textColor,
                        beginAtZero: true,
                        stepSize: 1
                    }
                }
            },
            layout: {
                padding: {
                    left: 20,
                    right: 20,
                    top: 15,
                    bottom: 15
                }
            },
            onClick: (event, activeElements) => {
                if (activeElements.length > 0) {
                    const element = activeElements[0];
                    const datasetIndex = element.datasetIndex;
                    const index = element.index;
                    
                    // Store customers and salesArray for use in highlight function
                    window.currentChartData = { customers, salesArray: salesToShow };
                    
                    const clickedCustomer = customers[index];
                    const clickedSales = salesToShow[datasetIndex];
                    
                    selectedCustomerSales.customer = clickedCustomer;
                    selectedCustomerSales.sales = clickedSales;
                    
                    // Show filtered table
                    showFilteredTable(clickedCustomer, clickedSales);
                    
                    // Highlight the clicked stack
                    setTimeout(() => {
                        highlightStackBar(element);
                    }, 10);
                }
            }
        }
    });
    
    // Update summary - pass all customers, not just the displayed ones
    updateRFQCustomerSalesSummary(allCustomers, customerSalesMap, filteredData.length);
    
    // Update title based on orientation
    const titleEl = document.getElementById('rfq-customer-sales-title');
    if (titleEl) {
        if (isCustomerSales) {
            titleEl.textContent = 'Jumlah RFQ per Customer per Sales';
        } else {
            titleEl.textContent = 'Jumlah RFQ per Sales per Customer';
        }
    }
    
    // Update toggle all button text based on current orientation
    const toggleAllBtn = document.getElementById('rfq-customer-sales-toggle-all');
    if (toggleAllBtn) {
        const showCountInput = document.getElementById('rfq-customer-sales-show-count');
        const showCount = showCountInput ? parseInt(showCountInput.value) || 15 : 15;
        if (chartShowAll) {
            toggleAllBtn.textContent = `📊 Top ${showCount}`;
            toggleAllBtn.title = `Show Top ${showCount} Only`;
        } else {
            toggleAllBtn.textContent = '📈 Show All';
            toggleAllBtn.title = 'Show All ' + (isCustomerSales ? 'Customers' : 'Sales');
        }
    }
}

function generateSalesColors(count) {
    const colors = [];
    
    // Distinct color palette with different hues for each sales
    const colorPalette = [
        { hue: 217, saturation: 75, lightness: 55 }, // Blue
        { hue: 158, saturation: 70, lightness: 50 }, // Teal/Green
        { hue: 340, saturation: 70, lightness: 55 }, // Pink/Red
        { hue: 45, saturation: 85, lightness: 55 },  // Yellow/Orange
        { hue: 270, saturation: 65, lightness: 55 }, // Purple
        { hue: 195, saturation: 75, lightness: 55 }, // Cyan
        { hue: 25, saturation: 80, lightness: 55 },  // Orange
        { hue: 145, saturation: 70, lightness: 50 }, // Green
        { hue: 300, saturation: 70, lightness: 55 }, // Magenta
        { hue: 210, saturation: 75, lightness: 60 }, // Light Blue
        { hue: 10, saturation: 75, lightness: 55 },  // Red-Orange
        { hue: 180, saturation: 70, lightness: 50 }, // Turquoise
        { hue: 290, saturation: 65, lightness: 55 }, // Violet
        { hue: 50, saturation: 85, lightness: 60 },  // Light Yellow
        { hue: 230, saturation: 70, lightness: 55 }, // Indigo
    ];
    
    for (let i = 0; i < count; i++) {
        const colorIndex = i % colorPalette.length;
        const color = colorPalette[colorIndex];
        
        // Create base color and lighter shade for highlight
        colors.push({
            base: `hsla(${color.hue}, ${color.saturation}%, ${color.lightness}%, 0.85)`,
            highlight: `hsla(${color.hue}, ${color.saturation}%, ${Math.min(color.lightness + 20, 85)}%, 1)`
        });
    }
    
    return colors;
}

function updateRFQCustomerSalesSummary(customers, customerSalesMap, totalFilteredRows = null) {
    const summaryEl = document.getElementById('rfq-customer-sales-summary');
    if (!summaryEl || customers.length === 0) return;
    
    const topCustomer = customers[0];
    // Calculate top customer total - sum across all sales for this customer
    const topCustomerTotal = Object.values(customerSalesMap[topCustomer] || {}).reduce((sum, val) => sum + val, 0);
    
    // Calculate total across ALL customers in the map (not just displayed ones)
    const allCustomersInMap = Object.keys(customerSalesMap);
    const totalAllCustomers = allCustomersInMap.reduce((sum, customer) => {
        return sum + Object.values(customerSalesMap[customer] || {}).reduce((customerSum, val) => customerSum + val, 0);
    }, 0);
    
    // Get total data points from all parsed data (for percentage calculation)
    const totalDataPoints = dashboardData.parsed ? dashboardData.parsed.length : 0;
    
    // Check if sales filter is applied
    const salesFilter = document.getElementById('rfq-customer-sales-sales');
    const selectedSales = salesFilter?.value?.trim() || '';
    
    // If sales filter is applied, calculate total for that sales only
    // IMPORTANT: Sum across ALL customers in customerSalesMap, not just displayed top 15
    let salesTotal = 0;
    if (selectedSales) {
        const salesFilterValue = selectedSales.toLowerCase();
        // Iterate through ALL customers in the map, not just the displayed ones
        allCustomersInMap.forEach(customer => {
            const customerSales = customerSalesMap[customer] || {};
            Object.keys(customerSales).forEach(salesKey => {
                if (salesKey.trim().toLowerCase() === salesFilterValue) {
                    salesTotal += customerSales[salesKey];
                }
            });
        });
        
        // Debug: Log the calculation
        console.log(`=== Summary Calculation for "${selectedSales}" ===`);
        console.log(`Total customers in map: ${allCustomersInMap.length}`);
        console.log(`Displayed customers: ${customers.length}`);
        console.log(`Sales total (all customers): ${salesTotal}`);
        console.log(`Total all customers (all sales): ${totalAllCustomers}`);
    }
    
    // Calculate percentages
    const topCustomerPercentage = totalAllCustomers > 0 ? ((topCustomerTotal / totalAllCustomers) * 100).toFixed(1) : '0.0';
    const totalPercentage = totalDataPoints > 0 ? ((totalAllCustomers / totalDataPoints) * 100).toFixed(1) : '0.0';
    const salesTotalPercentage = totalDataPoints > 0 && selectedSales ? ((salesTotal / totalDataPoints) * 100).toFixed(1) : null;
    
    // If totalFilteredRows is provided, use that for "from" calculation instead of totalDataPoints
    const displayTotal = totalFilteredRows !== null ? totalFilteredRows : totalDataPoints;
    
    let summaryText = `Top Customer: ${topCustomer} (${topCustomerTotal} RFQ, ${topCustomerPercentage}%)`;
    
    // If total filtered rows is provided and both statuses are checked, show verification
    const statusJadiOC = document.getElementById('rfq-customer-sales-status-jadi-oc');
    const statusTidakJadiOC = document.getElementById('rfq-customer-sales-status-tidak-jadi-oc');
    const bothChecked = statusJadiOC?.checked && statusTidakJadiOC?.checked;
    
    if (selectedSales) {
        // Show sales-specific total (sum across ALL customers, not just displayed)
        summaryText += ` | Sales "${selectedSales}": ${salesTotal} RFQ`;
        if (salesTotalPercentage !== null && displayTotal > 0) {
            const salesFromTotal = (salesTotal / displayTotal * 100).toFixed(1);
            summaryText += ` (${salesTotal} from ${displayTotal}, ${salesFromTotal}%)`;
        }
        if (allCustomersInMap.length > customers.length) {
            summaryText += ` | Across ${allCustomersInMap.length} customers (showing ${customers.length} out of ${allCustomersInMap.length} in chart)`;
        }
    } else if (bothChecked && totalFilteredRows !== null) {
        summaryText += ` | Total: ${totalAllCustomers} RFQ (${totalAllCustomers} from ${displayTotal}, ${totalPercentage}%)`;
        if (allCustomersInMap.length > customers.length) {
            summaryText += ` | Showing ${customers.length} out of ${allCustomersInMap.length} customers`;
        }
    } else {
        // Determine what's being displayed (customers or sales) based on orientation
        const isCustomerSales = chartAxisOrientation === 'customer-sales';
        const showCountInput = document.getElementById('rfq-customer-sales-show-count');
        const showCount = showCountInput ? parseInt(showCountInput.value) || 15 : 15;
        const displayedCount = isCustomerSales ? customers.length : (window.currentSalesCount || (window.allSalesSorted ? (chartShowAll ? window.allSalesSorted.length : Math.min(window.allSalesSorted.length, showCount)) : 0));
        const totalCount = isCustomerSales ? allCustomersInMap.length : (window.allSalesSorted ? window.allSalesSorted.length : 0);
        
        if (totalCount > displayedCount || chartShowAll === false) {
            const entityName = isCustomerSales ? 'customers' : 'sales';
            summaryText += ` | Total: ${totalAllCustomers} RFQ (${totalAllCustomers} from ${displayTotal}, ${totalPercentage}%) | Showing ${displayedCount} out of ${totalCount} ${entityName}`;
        } else {
            summaryText += ` | Total: ${totalAllCustomers} RFQ (${totalAllCustomers} from ${displayTotal}, ${totalPercentage}%)`;
        }
    }
    
    summaryEl.querySelector('.summary-text').textContent = summaryText;
}

function showFilteredTable(customer, sales) {
    const tableContainer = document.getElementById('rfq-customer-sales-table-container');
    const tableBody = document.getElementById('rfq-customer-sales-table-body');
    
    if (!tableContainer || !tableBody) return;
    
    // Filter data by selected customer and sales
    let filteredData = dashboardData.parsed.filter(item => 
        item.customer === customer && item.sales === sales
    );
    
    // Get additional filters
    const dateFromInput = document.getElementById('rfq-customer-sales-date-from');
    const dateToInput = document.getElementById('rfq-customer-sales-date-to');
    
    if (dateFromInput?.value) {
        const fromDate = new Date(dateFromInput.value);
        filteredData = filteredData.filter(d => {
            if (!d.date) return false;
            return d.date >= fromDate;
        });
    }
    
    if (dateToInput?.value) {
        const toDate = new Date(dateToInput.value);
        toDate.setHours(23, 59, 59, 999);
        filteredData = filteredData.filter(d => {
            if (!d.date) return false;
            return d.date <= toDate;
        });
    }
    
    // Format date for display
    const formatDate = (date) => {
        if (!date) return '';
        return new Date(date).toLocaleDateString('id-ID', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit'
        });
    };
    
    // Populate table
    tableBody.innerHTML = filteredData.map(item => `
        <tr>
            <td>${item.rfqNumber || ''}</td>
            <td>${formatDate(item.date)}</td>
            <td>${item.sales || ''}</td>
            <td>${item.customer || ''}</td>
        </tr>
    `).join('');
    
    if (filteredData.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="4" style="text-align: center; color: var(--text-secondary);">Tidak ada data</td></tr>';
    }
    
    // Show table container
    tableContainer.style.display = 'block';
    
    // Store filtered data for "View All Rows" button
    window.currentFilteredData = filteredData;
}

function highlightStackBar(element) {
    const chart = charts.rfqCustomerSales;
    if (!chart || !window.currentChartData) return;
    
    const { customers } = window.currentChartData;
    
    // Reset all datasets to original colors first
    chart.data.datasets.forEach((dataset) => {
        if (Array.isArray(dataset.backgroundColor)) {
            dataset.backgroundColor = dataset.backgroundColor.map(() => dataset.originalColor);
        } else {
            // Convert to array
            dataset.backgroundColor = customers.map(() => dataset.originalColor);
        }
    });
    
    // Highlight the clicked stack
    const clickedDataset = chart.data.datasets[element.datasetIndex];
    if (clickedDataset) {
        if (!Array.isArray(clickedDataset.backgroundColor)) {
            clickedDataset.backgroundColor = customers.map(() => clickedDataset.originalColor);
        }
        clickedDataset.backgroundColor[element.index] = clickedDataset.highlightColor;
    }
    
    chart.update('none'); // Update without animation
}

function showAllFilteredRows() {
    if (!window.currentFilteredData) return;
    
    showFilteredTable(selectedCustomerSales.customer, selectedCustomerSales.sales);
}

// ============================================
// Chart: Top Customers
// ============================================

function createTopCustomersChart() {
    const ctx = document.getElementById('chartTopCustomers');
    if (!ctx) return;
    
    const filters = getFilterValues('2');
    const data = filterData(filters);
    
    // Aggregate tanpa console.log (untuk chart rendering)
    const customerData = {};
    data.forEach(item => {
        if (!item.customer) return;
        if (!customerData[item.customer]) {
            customerData[item.customer] = 0;
        }
        customerData[item.customer]++;
    });
    
    const aggregated = Object.entries(customerData)
        .map(([customer, count]) => ({ customer, count }))
        .sort((a, b) => b.count - a.count);
    const sorted = aggregated.slice(0, 10);
    
    const labels = sorted.map(item => item.customer.length > 25 ? item.customer.substring(0, 22) + '...' : item.customer);
    const counts = sorted.map(item => item.count);
    
    if (charts.topCustomers) {
        charts.topCustomers.destroy();
    }
    
    charts.topCustomers = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Jumlah RFQ',
                data: counts,
                backgroundColor: CONFIG.accentColor + '60',
                borderColor: CONFIG.accentColor,
                borderWidth: 1.5
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            indexAxis: 'y',
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    padding: 8,
                    titleFont: { size: 12 },
                    bodyFont: { size: 11 }
                }
            },
            scales: {
                x: {
                    grid: { display: true, color: CONFIG.gridColor, lineWidth: 0.5 },
                    ticks: { font: { size: 10 }, color: CONFIG.textColor, beginAtZero: true }
                },
                y: {
                    grid: { display: false },
                    ticks: { font: { size: 10 }, color: CONFIG.textColor }
                }
            }
        }
    });
}

// ============================================
// Chart: Top Customer by RFQ Volume
// ============================================

function createTopCustomerRFQVolumeChart() {
    const ctx = document.getElementById('chartTopCustomerRFQVolume');
    if (!ctx || !dashboardData.parsed || dashboardData.parsed.length === 0) return;
    
    // Get status checkboxes
    const statusJadiOC = document.getElementById('top-customer-status-jadi-oc');
    const statusTidakJadiOC = document.getElementById('top-customer-status-tidak-jadi-oc');
    const includeJadiOC = statusJadiOC?.checked || false;
    const includeTidakJadiOC = statusTidakJadiOC?.checked || false;
    
    // Ensure at least one is checked
    if (!includeJadiOC && !includeTidakJadiOC) {
        if (statusJadiOC) statusJadiOC.checked = true;
        return createTopCustomerRFQVolumeChart(); // Retry with default
    }
    
    // Count RFQ per customer (count all, but we'll filter what to display)
    const customerDetails = {};
    
    dashboardData.parsed.forEach(item => {
        if (!item.customer || item.customer.trim() === '') return;
        
        const customer = item.customer.trim();
        
        if (!customerDetails[customer]) {
            customerDetails[customer] = {
                jadiOC: 0,
                tidakJadiOC: 0,
                items: []
            };
        }
        
        // Track conversion breakdown
        if (item.isConverted) {
            customerDetails[customer].jadiOC++;
        } else {
            customerDetails[customer].tidakJadiOC++;
        }
        
        customerDetails[customer].items.push(item);
    });
    
    // Convert to array and calculate counts
    let sortedCustomers = Object.entries(customerDetails)
        .map(([customer, details]) => ({
            customer,
            count: details.jadiOC + details.tidakJadiOC,
            jadiOC: details.jadiOC,
            tidakJadiOC: details.tidakJadiOC,
            items: details.items
        }));
    
    // Sort based on which checkboxes are selected
    if (includeJadiOC && includeTidakJadiOC) {
        // Both checked: Sort by total count (JADI OC + TIDAK JADI OC)
        sortedCustomers = sortedCustomers
            .filter(c => c.count > 0) // Only show customers with at least one RFQ
            .sort((a, b) => {
                // Primary sort: total count (descending)
                if (b.count !== a.count) return b.count - a.count;
                // Secondary sort: customer name (ascending) for consistency
                return a.customer.localeCompare(b.customer);
            });
    } else if (includeJadiOC && !includeTidakJadiOC) {
        // Only JADI OC checked: Sort by JADI OC count
        sortedCustomers = sortedCustomers
            .filter(c => c.jadiOC > 0) // Only show customers with JADI OC RFQ
            .sort((a, b) => {
                // Primary sort: JADI OC count (descending)
                if (b.jadiOC !== a.jadiOC) return b.jadiOC - a.jadiOC;
                // Secondary sort: customer name (ascending) for consistency
                return a.customer.localeCompare(b.customer);
            });
    } else if (!includeJadiOC && includeTidakJadiOC) {
        // Only TIDAK JADI OC checked: Sort by TIDAK JADI OC count
        sortedCustomers = sortedCustomers
            .filter(c => c.tidakJadiOC > 0) // Only show customers with TIDAK JADI OC RFQ
            .sort((a, b) => {
                // Primary sort: TIDAK JADI OC count (descending)
                if (b.tidakJadiOC !== a.tidakJadiOC) return b.tidakJadiOC - a.tidakJadiOC;
                // Secondary sort: customer name (ascending) for consistency
                return a.customer.localeCompare(b.customer);
            });
    } else {
        // Neither checked (shouldn't happen, but fallback to total count)
        sortedCustomers = sortedCustomers
            .filter(c => c.count > 0)
            .sort((a, b) => {
                if (b.count !== a.count) return b.count - a.count;
                return a.customer.localeCompare(b.customer);
            });
    }
    
    // Get custom show count
    const showCountInput = document.getElementById('top-customer-show-count');
    const showCount = showCountInput ? parseInt(showCountInput.value) || 20 : 20;
    topCustomerRFQVolumeShowCount = showCount;
    
    // Show top N customers
    const topCustomers = sortedCustomers.slice(0, showCount);
    
    // Prepare chart data - grouped bar chart
    const labels = topCustomers.map(item => {
        // Truncate long customer names
        const name = item.customer;
        return name.length > 40 ? name.substring(0, 37) + '...' : name;
    });
    
    // Prepare datasets based on checkboxes
    const datasets = [];
    
    if (includeJadiOC) {
        datasets.push({
            label: 'JADI OC',
            data: topCustomers.map(item => item.jadiOC),
            backgroundColor: '#10b981', // Green for converted
            borderColor: '#059669',
            borderWidth: 1
        });
    }
    
    if (includeTidakJadiOC) {
        datasets.push({
            label: 'TIDAK JADI OC',
            data: topCustomers.map(item => item.tidakJadiOC),
            backgroundColor: '#ef4444', // Red for not converted
            borderColor: '#dc2626',
            borderWidth: 1
        });
    }
    
    // Destroy existing chart if it exists
    if (charts.topCustomerRFQVolume) {
        charts.topCustomerRFQVolume.destroy();
    }
    
    // Create grouped horizontal bar chart
    charts.topCustomerRFQVolume = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: datasets
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: datasets.length > 1,
                    position: 'top',
                    labels: {
                        usePointStyle: true,
                        padding: 12,
                        font: { size: 11 },
                        color: CONFIG.textColor
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.85)',
                    padding: 10,
                    titleFont: { size: 13, weight: 'bold' },
                    bodyFont: { size: 12 },
                    callbacks: {
                        label: function(context) {
                            const index = context.dataIndex;
                            const customer = topCustomers[index];
                            const label = context.dataset.label;
                            const value = context.parsed.x;
                            return `${label}: ${value} RFQ`;
                        },
                        afterBody: function(context) {
                            const index = context[0].dataIndex;
                            const customer = topCustomers[index];
                            return [
                                `Total: ${customer.count} RFQ`,
                                `(${customer.jadiOC} JADI OC, ${customer.tidakJadiOC} TIDAK JADI OC)`
                            ];
                        }
                    }
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    stacked: false, // Grouped bars (side by side)
                    grid: {
                        display: true,
                        color: CONFIG.gridColor,
                        lineWidth: 0.5
                    },
                    ticks: {
                        font: { size: 11 },
                        color: CONFIG.textColor,
                        stepSize: 1
                    },
                    title: {
                        display: true,
                        text: 'Number of RFQ',
                        font: { size: 12, weight: 'bold' },
                        color: CONFIG.primaryTextColor
                    }
                },
                y: {
                    stacked: false, // Grouped bars (side by side)
                    grid: {
                        display: false
                    },
                    ticks: {
                        font: { size: 10 },
                        color: CONFIG.textColor
                    }
                }
            },
            onClick: (event, elements) => {
                if (elements.length > 0) {
                    const element = elements[0];
                    const index = element.index;
                    const customer = topCustomers[index];
                    showTopCustomerRFQTable(customer);
                }
            },
            onHover: (event, elements) => {
                event.native.target.style.cursor = elements.length > 0 ? 'pointer' : 'default';
            }
        }
    });
    
    // Update summary
    // Update summary with the FULL sorted list (before slicing to top N)
    // This ensures summary shows totals for all customers matching the filter, not just the displayed ones
    updateTopCustomerRFQSummary(sortedCustomers, includeJadiOC, includeTidakJadiOC);
    
    // Store data for click interaction
    window.topCustomerRFQData = topCustomers;
}

function showTopCustomerRFQTable(customerData) {
    const tableContainer = document.getElementById('top-customer-rfq-table-container');
    const tableBody = document.getElementById('top-customer-rfq-table-body');
    const tableTitle = document.getElementById('top-customer-rfq-table-title');
    const breakdownBody = document.getElementById('top-customer-rfq-breakdown-body');
    const breakdownThead = document.getElementById('top-customer-rfq-breakdown-thead');
    const breakdownTitle = document.getElementById('top-customer-rfq-breakdown-title');
    const breakdownTypeSelect = document.getElementById('top-customer-breakdown-type');
    
    if (!tableContainer || !tableBody || !tableTitle || !breakdownBody || !breakdownThead || !breakdownTitle) return;
    
    // Create summary title
    const summary = `${customerData.customer} – ${customerData.count} RFQ (${customerData.jadiOC} JADI OC, ${customerData.tidakJadiOC} TIDAK JADI OC)`;
    tableTitle.textContent = summary;
    
    // Store customerData for breakdown updates
    window.currentCustomerData = customerData;
    
    // Initial breakdown by selected type
    updateTopCustomerBreakdown(customerData, breakdownTypeSelect ? breakdownTypeSelect.value : 'marketing');
    
    // Setup breakdown type selector change handler
    if (breakdownTypeSelect && !breakdownTypeSelect.hasAttribute('data-listener-attached')) {
        breakdownTypeSelect.setAttribute('data-listener-attached', 'true');
        breakdownTypeSelect.addEventListener('change', () => {
            updateTopCustomerBreakdown(window.currentCustomerData, breakdownTypeSelect.value);
        });
    }
    
    // Get headers
    const headers = dashboardData.headers;
    const remarkIndex = headers.findIndex(h => h.trim() === 'REMARK (KAROSERI)');
    
    // Format date for display
    const formatDate = (date) => {
        if (!date) return '';
        return new Date(date).toLocaleDateString('id-ID', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit'
        });
    };
    
    // Format currency
    const formatCurrency = (value) => {
        if (!value || value === 0) return '';
        return new Intl.NumberFormat('id-ID', {
            style: 'currency',
            currency: 'IDR',
            minimumFractionDigits: 0,
            maximumFractionDigits: 0
        }).format(value);
    };
    
    // Sort items by date (newest first), then by sales
    const sortedItems = [...customerData.items].sort((a, b) => {
        // First sort by sales
        const salesA = (a.sales || '').trim();
        const salesB = (b.sales || '').trim();
        if (salesA !== salesB) {
            return salesA.localeCompare(salesB);
        }
        // Then by date (newest first)
        if (!a.date && !b.date) return 0;
        if (!a.date) return 1;
        if (!b.date) return -1;
        return new Date(b.date) - new Date(a.date);
    });
    
    // Populate detailed table
    tableBody.innerHTML = sortedItems.map(item => {
        // Get REMARK from rawRow if available
        let remark = '';
        if (item.rawRow && remarkIndex >= 0 && remarkIndex < item.rawRow.length) {
            remark = String(item.rawRow[remarkIndex] || '').trim();
        }
        
        return `
            <tr>
                <td class="px-4 py-3 border-b border-gray-200 text-gray-900">${item.rfqNumber || ''}</td>
                <td class="px-4 py-3 border-b border-gray-200 text-gray-700">${formatDate(item.date)}</td>
                <td class="px-4 py-3 border-b border-gray-200 text-gray-700">${item.sales || ''}</td>
                <td class="px-4 py-3 border-b border-gray-200 text-gray-700">${item.customer || ''}</td>
                <td class="px-4 py-3 border-b border-gray-200 text-gray-700">${remark}</td>
                <td class="px-4 py-3 border-b border-gray-200 text-gray-700 text-right">${formatCurrency(item.hargaNew)}</td>
                <td class="px-4 py-3 border-b border-gray-200 text-gray-700">${item.status || ''}</td>
            </tr>
        `;
    }).join('');
    
    if (sortedItems.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="7" class="px-4 py-3 text-center text-gray-500">Tidak ada data</td></tr>';
    }
    
    // Show table container
    tableContainer.style.display = 'block';
    
    // Scroll to table
    tableContainer.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// ============================================
// Chart: Conversion Rate per Customer
// ============================================

function createConversionRateCustomerChart() {
    const ctx = document.getElementById('chartConversionRateCustomer');
    if (!ctx || !dashboardData.parsed || dashboardData.parsed.length === 0) return;
    
    // Get filter values
    const statusFilter = document.getElementById('conversion-status-filter');
    const dateFromInput = document.getElementById('conversion-date-from');
    const dateToInput = document.getElementById('conversion-date-to');
    const salesFilter = document.getElementById('conversion-sales-filter');
    const customerFilter = document.getElementById('conversion-customer-filter');
    const showCountInput = document.getElementById('conversion-show-count');
    
    const statusValue = statusFilter?.value || 'all';
    const dateFrom = dateFromInput?.value || '';
    const dateTo = dateToInput?.value || '';
    const selectedSales = salesFilter?.value?.trim() || '';
    const selectedCustomer = customerFilter?.value?.trim() || '';
    const showCount = showCountInput ? parseInt(showCountInput.value) || 20 : 20;
    
    // Filter data based on filters (EXCEPT status filter - status is only for display, not calculation)
    // Conversion rate must be calculated from ALL RFQs for each customer (within date/sales/customer filters)
    let filteredData = dashboardData.parsed.filter(item => {
        // Date filter
        if (dateFrom || dateTo) {
            if (!item.date) return false;
            const itemDate = new Date(item.date);
            if (dateFrom && itemDate < new Date(dateFrom)) return false;
            if (dateTo) {
                const toDate = new Date(dateTo);
                toDate.setHours(23, 59, 59, 999); // Include entire end date
                if (itemDate > toDate) return false;
            }
        }
        
        // Sales filter
        if (selectedSales) {
            const itemSales = item.sales ? item.sales.trim().toLowerCase() : '';
            if (itemSales !== selectedSales.toLowerCase()) return false;
        }
        
        // Customer filter
        if (selectedCustomer) {
            const itemCustomer = item.customer ? item.customer.trim() : '';
            if (itemCustomer !== selectedCustomer) return false;
        }
        
        // NOTE: Status filter is NOT applied here - conversion rate must use ALL RFQs
        return true;
    });
    
    // Aggregate conversion rate per customer
    // Count ALL RFQs for each customer (both JADI OC and TIDAK JADI OC) to calculate accurate conversion rate
    const customerDetails = {};
    
    filteredData.forEach(item => {
        if (!item.customer || item.customer.trim() === '') return;
        
        const customer = item.customer.trim();
        
        if (!customerDetails[customer]) {
            customerDetails[customer] = {
                totalRFQ: 0,
                convertedRFQ: 0,
                tidakJadiOCRfq: 0,
                items: []
            };
        }
        
        // Count ALL RFQs regardless of status
        customerDetails[customer].totalRFQ++;
        
        // Track conversion separately
        if (item.isConverted) {
            customerDetails[customer].convertedRFQ++;
        } else {
            customerDetails[customer].tidakJadiOCRfq++;
        }
        
        customerDetails[customer].items.push(item);
    });
    
    // Calculate conversion rate and convert to array
    const conversionData = Object.entries(customerDetails)
        .map(([customer, details]) => {
            const conversionRate = details.totalRFQ > 0 
                ? (details.convertedRFQ / details.totalRFQ) * 100 
                : 0;
            
            return {
                customer,
                totalRFQ: details.totalRFQ,
                convertedRFQ: details.convertedRFQ,
                tidakJadiOCRfq: details.tidakJadiOCRfq,
                conversionRate: conversionRate,
                items: details.items
            };
        })
        // Sort by conversion rate (highest to lowest), then by total RFQ (descending)
        .sort((a, b) => {
            if (Math.abs(b.conversionRate - a.conversionRate) > 0.01) {
                return b.conversionRate - a.conversionRate;
            }
            return b.totalRFQ - a.totalRFQ;
        });
    
    // Filter customers with at least 1 RFQ (required for meaningful conversion rate)
    const validCustomers = conversionData.filter(c => c.totalRFQ > 0);
    
    // Apply status filter to displayed customers if needed
    // IMPORTANT: Status filter only affects which customers are SHOWN, not the conversion rate calculation
    // Conversion rate is ALWAYS calculated as: convertedRFQ / totalRFQ (where totalRFQ includes both JADI OC and TIDAK JADI OC)
    let displayCustomers = validCustomers;
    if (statusValue === 'jadi-oc') {
        // Show only customers that have at least one JADI OC (but conversion rate still uses all their RFQs)
        displayCustomers = validCustomers.filter(c => c.convertedRFQ > 0);
    } else if (statusValue === 'tidak-jadi-oc') {
        // Show only customers that have at least one TIDAK JADI OC (but conversion rate still uses all their RFQs)
        displayCustomers = validCustomers.filter(c => c.tidakJadiOCRfq > 0);
    }
    // If statusValue === 'all', show all customers
    
    // Show top N customers
    const topCustomers = displayCustomers.slice(0, showCount);
    
    if (topCustomers.length === 0) {
        // Destroy existing chart if it exists
        if (charts.conversionRateCustomer) {
            charts.conversionRateCustomer.destroy();
        }
        
        // Show empty message
        const summaryEl = document.getElementById('conversion-rate-summary');
        if (summaryEl) {
            summaryEl.querySelector('.summary-text').textContent = 'No data available with current filters.';
        }
        return;
    }
    
    // Prepare chart data - horizontal bar chart
    const labels = topCustomers.map(item => {
        const name = item.customer;
        return name.length > 40 ? name.substring(0, 37) + '...' : name;
    });
    
    const conversionRates = topCustomers.map(item => item.conversionRate);
    
    // Color based on conversion rate
    const backgroundColor = conversionRates.map(rate => {
        if (rate >= 50) return '#10b981'; // Green for high conversion
        if (rate >= 25) return '#f59e0b'; // Orange for medium conversion
        return '#ef4444'; // Red for low conversion
    });
    
    // Destroy existing chart if it exists
    if (charts.conversionRateCustomer) {
        charts.conversionRateCustomer.destroy();
    }
    
    // Chart configuration
    const CONFIG = {
        primaryColor: '#3b82f6',
        textColor: '#6b7280',
        primaryTextColor: '#111827',
        gridColor: 'rgba(0, 0, 0, 0.05)'
    };
    
    // Create horizontal bar chart
    charts.conversionRateCustomer = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Conversion Rate (%)',
                data: conversionRates,
                backgroundColor: backgroundColor,
                borderColor: backgroundColor.map(c => c === '#10b981' ? '#059669' : c === '#f59e0b' ? '#d97706' : '#dc2626'),
                borderWidth: 1
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            animation: {
                duration: 200
            },
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.85)',
                    padding: 10,
                    titleFont: { size: 13, weight: 'bold' },
                    bodyFont: { size: 12 },
                    callbacks: {
                        label: function(context) {
                            const index = context.dataIndex;
                            const customer = topCustomers[index];
                            const tidakJadiOC = customer.tidakJadiOCRfq !== undefined 
                                ? customer.tidakJadiOCRfq 
                                : (customer.totalRFQ - customer.convertedRFQ);
                            return [
                                `Conversion Rate: ${customer.conversionRate.toFixed(2)}%`,
                                `Converted: ${customer.convertedRFQ} JADI OC`,
                                `Total: ${customer.totalRFQ} RFQ (${customer.convertedRFQ} JADI OC + ${tidakJadiOC} TIDAK JADI OC)`
                            ];
                        }
                    }
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    max: 100,
                    grid: {
                        display: true,
                        color: CONFIG.gridColor,
                        lineWidth: 0.5
                    },
                    ticks: {
                        font: { size: 11 },
                        color: CONFIG.textColor,
                        callback: function(value) {
                            return value + '%';
                        }
                    },
                    title: {
                        display: true,
                        text: 'Conversion Rate (%)',
                        font: { size: 12, weight: 'bold' },
                        color: CONFIG.primaryTextColor
                    }
                },
                y: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        font: { size: 10 },
                        color: CONFIG.textColor
                    }
                }
            },
            onClick: (event, elements) => {
                if (elements.length > 0) {
                    const element = elements[0];
                    const index = element.index;
                    const customer = topCustomers[index];
                    showConversionRateTable(customer);
                }
            },
            onHover: (event, elements) => {
                event.native.target.style.cursor = elements.length > 0 ? 'pointer' : 'default';
            }
        },
        plugins: [{
            id: 'datalabels',
            afterDatasetsDraw: (chart) => {
                const ctx = chart.ctx;
                ctx.save();
                
                chart.data.datasets.forEach((dataset, i) => {
                    const meta = chart.getDatasetMeta(i);
                    meta.data.forEach((bar, index) => {
                        const value = dataset.data[index];
                        const xPos = bar.x;
                        const yPos = bar.y;
                        
                        // Draw percentage label on the bar
                        ctx.fillStyle = '#111827';
                        ctx.font = 'bold 11px Inter, system-ui, sans-serif';
                        ctx.textAlign = 'left';
                        ctx.textBaseline = 'middle';
                        
                        // Only show label if bar is wide enough
                        if (xPos > 10) {
                            ctx.fillText(`${value.toFixed(1)}%`, xPos + 5, yPos);
                        }
                    });
                });
                
                ctx.restore();
            }
        }]
    });
    
    // Update summary
    updateConversionRateSummary(topCustomers, displayCustomers.length);
    
    // Store data for click interaction
    window.conversionRateCustomerData = topCustomers;
}

function showConversionRateTable(customerData) {
    const tableContainer = document.getElementById('conversion-rate-table-container');
    const tableBody = document.getElementById('conversion-rate-table-body');
    const tableTitle = document.getElementById('conversion-rate-table-title');
    
    if (!tableContainer || !tableBody || !tableTitle) return;
    
    // Create summary title
    const summary = `${customerData.customer} – Conversion Rate: ${customerData.conversionRate.toFixed(2)}% (${customerData.convertedRFQ} converted from ${customerData.totalRFQ} RFQ)`;
    tableTitle.textContent = summary;
    
    // Get headers
    const headers = dashboardData.headers;
    const remarkIndex = headers.findIndex(h => h.trim() === 'REMARK (KAROSERI)');
    
    // Format date for display
    const formatDate = (date) => {
        if (!date) return '';
        return new Date(date).toLocaleDateString('id-ID', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit'
        });
    };
    
    // Format currency
    const formatCurrency = (value) => {
        if (!value || value === 0) return '';
        return new Intl.NumberFormat('id-ID', {
            style: 'currency',
            currency: 'IDR',
            minimumFractionDigits: 0,
            maximumFractionDigits: 0
        }).format(value);
    };
    
    // Sort items by date (newest first)
    const sortedItems = [...customerData.items].sort((a, b) => {
        if (!a.date && !b.date) return 0;
        if (!a.date) return 1;
        if (!b.date) return -1;
        return new Date(b.date) - new Date(a.date);
    });
    
    // Populate detailed table
    tableBody.innerHTML = sortedItems.map(item => {
        // Get REMARK from rawRow if available
        let remark = '';
        if (item.rawRow && remarkIndex >= 0 && remarkIndex < item.rawRow.length) {
            remark = String(item.rawRow[remarkIndex] || '').trim();
        }
        
        return `
            <tr>
                <td class="px-4 py-3 border-b border-gray-200 text-gray-900">${item.rfqNumber || ''}</td>
                <td class="px-4 py-3 border-b border-gray-200 text-gray-700">${formatDate(item.date)}</td>
                <td class="px-4 py-3 border-b border-gray-200 text-gray-700">${item.sales || ''}</td>
                <td class="px-4 py-3 border-b border-gray-200 text-gray-700">${item.customer || ''}</td>
                <td class="px-4 py-3 border-b border-gray-200 text-gray-700">${remark}</td>
                <td class="px-4 py-3 border-b border-gray-200 text-gray-700 text-right">${formatCurrency(item.hargaNew)}</td>
                <td class="px-4 py-3 border-b border-gray-200 text-gray-700">${item.status || ''}</td>
            </tr>
        `;
    }).join('');
    
    if (sortedItems.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="7" class="px-4 py-3 text-center text-gray-500">Tidak ada data</td></tr>';
    }
    
    // Show table container
    tableContainer.style.display = 'block';
    
    // Scroll to table
    tableContainer.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function updateConversionRateSummary(displayedCustomers, totalCustomers) {
    const summaryEl = document.getElementById('conversion-rate-summary');
    if (!summaryEl) return;
    
    const displayedCount = displayedCustomers.length;
    const avgConversionRate = displayedCustomers.length > 0
        ? displayedCustomers.reduce((sum, c) => sum + c.conversionRate, 0) / displayedCustomers.length
        : 0;
    
    let summaryText = `Showing ${displayedCount} out of ${totalCustomers} customers`;
    summaryText += ` | Average Conversion Rate: ${avgConversionRate.toFixed(2)}%`;
    
    summaryEl.querySelector('.summary-text').textContent = summaryText;
}

function updateTopCustomerBreakdown(customerData, breakdownType) {
    const breakdownBody = document.getElementById('top-customer-rfq-breakdown-body');
    const breakdownThead = document.getElementById('top-customer-rfq-breakdown-thead');
    const breakdownTitle = document.getElementById('top-customer-rfq-breakdown-title');
    
    if (!breakdownBody || !breakdownThead || !breakdownTitle) return;
    
    const headers = dashboardData.headers;
    const remarkIndex = headers.findIndex(h => h.trim() === 'REMARK (KAROSERI)');
    
    // Format currency helper
    const formatCurrency = (value) => {
        if (!value || value === 0) return '';
        return new Intl.NumberFormat('id-ID', {
            style: 'currency',
            currency: 'IDR',
            minimumFractionDigits: 0,
            maximumFractionDigits: 0
        }).format(value);
    };
    
    // Format date helper
    const formatDateMonth = (date) => {
        if (!date) return '';
        return new Date(date).toLocaleDateString('id-ID', {
            year: 'numeric',
            month: 'short'
        });
    };
    
    let breakdownMap = {};
    let title = '';
    let theadHTML = '';
    let bodyHTML = '';
    
    switch (breakdownType) {
        case 'marketing':
            title = 'Breakdown by MARKERTING:';
            theadHTML = `
                <tr>
                    <th class="px-4 py-2 text-left font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">CUSTOMER NAME</th>
                    <th class="px-4 py-2 text-left font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">MARKERTING</th>
                    <th class="px-4 py-2 text-center font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">TOTAL RFQ</th>
                    <th class="px-4 py-2 text-center font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">JADI OC</th>
                    <th class="px-4 py-2 text-center font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">TIDAK JADI OC</th>
                </tr>
            `;
            
            customerData.items.forEach(item => {
                const sales = item.sales ? item.sales.trim() : '(No Sales)';
                if (!breakdownMap[sales]) {
                    breakdownMap[sales] = { value: sales, total: 0, jadiOC: 0, tidakJadiOC: 0 };
                }
                breakdownMap[sales].total++;
                if (item.isConverted) breakdownMap[sales].jadiOC++;
                else breakdownMap[sales].tidakJadiOC++;
            });
            
            bodyHTML = Object.values(breakdownMap)
                .sort((a, b) => b.total - a.total)
                .map(data => `
                    <tr>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-900">${customerData.customer}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-700">${data.value}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-900 text-center font-semibold">${data.total}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-700 text-center">${data.jadiOC}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-700 text-center">${data.tidakJadiOC}</td>
                    </tr>
                `).join('');
            break;
            
        case 'harga':
            title = 'Breakdown by HARGA (NEW):';
            theadHTML = `
                <tr>
                    <th class="px-4 py-2 text-left font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">CUSTOMER NAME</th>
                    <th class="px-4 py-2 text-left font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">HARGA RANGE</th>
                    <th class="px-4 py-2 text-center font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">TOTAL RFQ</th>
                    <th class="px-4 py-2 text-center font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">JADI OC</th>
                    <th class="px-4 py-2 text-center font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">TIDAK JADI OC</th>
                    <th class="px-4 py-2 text-right font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">TOTAL AMOUNT</th>
                </tr>
            `;
            
            customerData.items.forEach(item => {
                const harga = item.hargaNew || 0;
                let range = '';
                if (harga === 0) {
                    range = 'No Price';
                } else if (harga < 100000000) {
                    range = '< 100M';
                } else if (harga < 300000000) {
                    range = '100M - 300M';
                } else if (harga < 500000000) {
                    range = '300M - 500M';
                } else if (harga < 1000000000) {
                    range = '500M - 1B';
                } else {
                    range = '> 1B';
                }
                
                if (!breakdownMap[range]) {
                    breakdownMap[range] = { value: range, total: 0, jadiOC: 0, tidakJadiOC: 0, totalAmount: 0 };
                }
                breakdownMap[range].total++;
                breakdownMap[range].totalAmount += harga;
                if (item.isConverted) breakdownMap[range].jadiOC++;
                else breakdownMap[range].tidakJadiOC++;
            });
            
            bodyHTML = Object.values(breakdownMap)
                .sort((a, b) => {
                    const order = { 'No Price': 0, '< 100M': 1, '100M - 300M': 2, '300M - 500M': 3, '500M - 1B': 4, '> 1B': 5 };
                    return (order[a.value] || 999) - (order[b.value] || 999);
                })
                .map(data => `
                    <tr>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-900">${customerData.customer}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-700">${data.value}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-900 text-center font-semibold">${data.total}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-700 text-center">${data.jadiOC}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-700 text-center">${data.tidakJadiOC}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-700 text-right">${formatCurrency(data.totalAmount)}</td>
                    </tr>
                `).join('');
            break;
            
        case 'remark':
            title = 'Breakdown by REMARK (KAROSERI):';
            theadHTML = `
                <tr>
                    <th class="px-4 py-2 text-left font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">CUSTOMER NAME</th>
                    <th class="px-4 py-2 text-left font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">REMARK (KAROSERI)</th>
                    <th class="px-4 py-2 text-center font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">TOTAL RFQ</th>
                    <th class="px-4 py-2 text-center font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">JADI OC</th>
                    <th class="px-4 py-2 text-center font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">TIDAK JADI OC</th>
                </tr>
            `;
            
            customerData.items.forEach(item => {
                let remark = '';
                if (item.rawRow && remarkIndex >= 0 && remarkIndex < item.rawRow.length) {
                    remark = String(item.rawRow[remarkIndex] || '').trim();
                }
                if (!remark) remark = '(No Remark)';
                const remarkDisplay = remark.length > 60 ? remark.substring(0, 57) + '...' : remark;
                
                if (!breakdownMap[remark]) {
                    breakdownMap[remark] = { value: remark, display: remarkDisplay, total: 0, jadiOC: 0, tidakJadiOC: 0 };
                }
                breakdownMap[remark].total++;
                if (item.isConverted) breakdownMap[remark].jadiOC++;
                else breakdownMap[remark].tidakJadiOC++;
            });
            
            bodyHTML = Object.values(breakdownMap)
                .sort((a, b) => b.total - a.total)
                .map(data => `
                    <tr>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-900">${customerData.customer}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-700" title="${data.value}">${data.display}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-900 text-center font-semibold">${data.total}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-700 text-center">${data.jadiOC}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-700 text-center">${data.tidakJadiOC}</td>
                    </tr>
                `).join('');
            break;
            
        case 'date':
            title = 'Breakdown by DATE:';
            theadHTML = `
                <tr>
                    <th class="px-4 py-2 text-left font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">CUSTOMER NAME</th>
                    <th class="px-4 py-2 text-left font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">MONTH/YEAR</th>
                    <th class="px-4 py-2 text-center font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">TOTAL RFQ</th>
                    <th class="px-4 py-2 text-center font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">JADI OC</th>
                    <th class="px-4 py-2 text-center font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">TIDAK JADI OC</th>
                </tr>
            `;
            
            customerData.items.forEach(item => {
                const dateStr = item.date ? formatDateMonth(item.date) : '(No Date)';
                
                if (!breakdownMap[dateStr]) {
                    breakdownMap[dateStr] = { value: dateStr, total: 0, jadiOC: 0, tidakJadiOC: 0 };
                }
                breakdownMap[dateStr].total++;
                if (item.isConverted) breakdownMap[dateStr].jadiOC++;
                else breakdownMap[dateStr].tidakJadiOC++;
            });
            
            bodyHTML = Object.values(breakdownMap)
                .sort((a, b) => {
                    if (a.value === '(No Date)') return 1;
                    if (b.value === '(No Date)') return -1;
                    return b.value.localeCompare(a.value);
                })
                .map(data => `
                    <tr>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-900">${customerData.customer}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-700">${data.value}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-900 text-center font-semibold">${data.total}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-700 text-center">${data.jadiOC}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-700 text-center">${data.tidakJadiOC}</td>
                    </tr>
                `).join('');
            break;
            
        case 'status':
            title = 'Breakdown by KETERANGAN:';
            theadHTML = `
                <tr>
                    <th class="px-4 py-2 text-left font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">CUSTOMER NAME</th>
                    <th class="px-4 py-2 text-left font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">KETERANGAN</th>
                    <th class="px-4 py-2 text-center font-semibold text-gray-600 uppercase text-xs tracking-wider border-b-2 border-gray-200 bg-gray-50">TOTAL RFQ</th>
                </tr>
            `;
            
            customerData.items.forEach(item => {
                const status = item.status || '(No Status)';
                
                if (!breakdownMap[status]) {
                    breakdownMap[status] = { value: status, total: 0 };
                }
                breakdownMap[status].total++;
            });
            
            bodyHTML = Object.values(breakdownMap)
                .sort((a, b) => b.total - a.total)
                .map(data => `
                    <tr>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-900">${customerData.customer}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-700">${data.value}</td>
                        <td class="px-4 py-2 border-b border-gray-200 text-gray-900 text-center font-semibold">${data.total}</td>
                    </tr>
                `).join('');
            break;
    }
    
    // Update UI
    breakdownTitle.textContent = title;
    breakdownThead.innerHTML = theadHTML;
    breakdownBody.innerHTML = bodyHTML || '<tr><td colspan="5" class="px-4 py-2 text-center text-gray-500">Tidak ada data</td></tr>';
}

function updateTopCustomerRFQSummary(sortedCustomers, includeJadiOC, includeTidakJadiOC) {
    const summaryEl = document.getElementById('top-customer-rfq-summary');
    if (!summaryEl) return;
    
    const totalCustomers = sortedCustomers.length;
    // Get the actual displayed count from the input or use default
    const showCountInput = document.getElementById('top-customer-show-count');
    const showCount = showCountInput ? parseInt(showCountInput.value) || 20 : 20;
    const displayedCount = Math.min(totalCustomers, showCount);
    
    // Calculate totals based on which checkboxes are selected
    const totalJadiOC = sortedCustomers.reduce((sum, c) => sum + c.jadiOC, 0);
    const totalTidakJadiOC = sortedCustomers.reduce((sum, c) => sum + c.tidakJadiOC, 0);
    const totalRFQ = sortedCustomers.reduce((sum, c) => sum + c.count, 0);
    
    // Determine the relevant total based on selection
    let relevantTotal = 0;
    if (includeJadiOC && includeTidakJadiOC) {
        relevantTotal = totalRFQ; // Show total when both are selected
    } else if (includeJadiOC && !includeTidakJadiOC) {
        relevantTotal = totalJadiOC; // Show only JADI OC total
    } else if (!includeJadiOC && includeTidakJadiOC) {
        relevantTotal = totalTidakJadiOC; // Show only TIDAK JADI OC total
    } else {
        relevantTotal = totalRFQ; // Fallback
    }
    
    let summaryText = `Showing ${displayedCount} out of ${totalCustomers} customers`;
    summaryText += ` | Total: ${relevantTotal} RFQ`;
    
    if (includeJadiOC && includeTidakJadiOC) {
        summaryText += ` (${totalJadiOC} JADI OC, ${totalTidakJadiOC} TIDAK JADI OC)`;
    } else if (includeJadiOC) {
        summaryText += ` (${totalJadiOC} JADI OC)`;
    } else if (includeTidakJadiOC) {
        summaryText += ` (${totalTidakJadiOC} TIDAK JADI OC)`;
    }
    
    summaryEl.querySelector('.summary-text').textContent = summaryText;
}

// ============================================
// Chart: Conversion % per Customer
// ============================================

function createConversionPerCustomerChart() {
    const ctx = document.getElementById('chartConversionPerCustomer');
    if (!ctx) return;
    
    const filters = getFilterValues('3');
    const data = filterData(filters);
    
    // Aggregate tanpa console.log (untuk chart rendering)
    const customerData = {};
    data.forEach(item => {
        if (!item.customer) return;
        if (!customerData[item.customer]) {
            customerData[item.customer] = { total: 0, converted: 0 };
        }
        customerData[item.customer].total++;
        if (item.isConverted) {
            customerData[item.customer].converted++;
        }
    });
    
    const aggregated = Object.values(customerData)
        .map(item => ({
            customer: item.customer,
            total: item.total,
            converted: item.converted,
            conversionRate: item.total > 0 ? (item.converted / item.total) * 100 : 0
        }))
        .sort((a, b) => b.conversionRate - a.conversionRate);
    const sorted = aggregated.slice(0, 10);
    
    const labels = sorted.map(item => item.customer.length > 25 ? item.customer.substring(0, 22) + '...' : item.customer);
    const rates = sorted.map(item => item.conversionRate);
    
    if (charts.conversionPerCustomer) {
        charts.conversionPerCustomer.destroy();
    }
    
    charts.conversionPerCustomer = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: '% Konversi',
                data: rates,
                backgroundColor: '#10b98180',
                borderColor: '#10b981',
                borderWidth: 1.5
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            indexAxis: 'y',
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    padding: 8,
                    titleFont: { size: 12 },
                    bodyFont: { size: 11 },
                    callbacks: {
                        label: function(context) {
                            const idx = context.dataIndex;
                            const item = sorted[idx];
                            return `${item.rate.toFixed(2)}% (${item.converted}/${item.total})`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: { display: true, color: CONFIG.gridColor, lineWidth: 0.5 },
                    ticks: { 
                        font: { size: 10 }, 
                        color: CONFIG.textColor, 
                        beginAtZero: true,
                        callback: function(value) { return value.toFixed(0) + '%'; }
                    }
                },
                y: {
                    grid: { display: false },
                    ticks: { font: { size: 10 }, color: CONFIG.textColor }
                }
            }
        }
    });
}

// ============================================
// Chart: Total Amount per Customer per Sales
// ============================================

function createAmountPerCustomerSalesChart() {
    const ctx = document.getElementById('chartAmountPerCustomerSales');
    if (!ctx) return;
    
    const filters = getFilterValues('4');
    const data = filterData(filters);
    
    // Aggregate tanpa console.log (untuk chart rendering)
    const grouped = {};
    data.forEach(item => {
        if (!item.customer || !item.sales) return;
        const key = `${item.customer}|${item.sales}`;
        if (!grouped[key]) {
            grouped[key] = { customer: item.customer, sales: item.sales, amount: 0, count: 0 };
        }
        grouped[key].amount += item.amount;
        grouped[key].count++;
    });
    
    const aggregated = Object.values(grouped).sort((a, b) => b.amount - a.amount);
    const sorted = aggregated.slice(0, 15);
    
    const labels = sorted.map(item => `${item.customer} - ${item.sales}`);
    const amounts = sorted.map(item => item.amount);
    
    if (charts.amountPerCustomerSales) {
        charts.amountPerCustomerSales.destroy();
    }
    
    charts.amountPerCustomerSales = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels.map(l => l.length > 40 ? l.substring(0, 37) + '...' : l),
            datasets: [{
                label: 'Total Amount (Rp)',
                data: amounts,
                backgroundColor: CONFIG.accentColor + '80',
                borderColor: CONFIG.accentColor,
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            indexAxis: 'y',
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    padding: 8,
                    titleFont: { size: 12 },
                    bodyFont: { size: 11 },
                    callbacks: {
                        label: function(context) {
                            return 'Rp ' + formatCurrency(context.parsed.x);
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: { display: true, color: CONFIG.gridColor, lineWidth: 0.5 },
                    ticks: { 
                        font: { size: 10 }, 
                        color: CONFIG.textColor, 
                        beginAtZero: true,
                        callback: function(value) { return 'Rp ' + formatCurrency(value); }
                    }
                },
                y: {
                    grid: { display: false },
                    ticks: { font: { size: 9 }, color: CONFIG.textColor }
                }
            }
        }
    });
}

// ============================================
// Original Data Table
// ============================================

function renderOriginalTable() {
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');
    
    if (!tableHead || !tableBody) return;
    
    // Show original headers
    tableHead.innerHTML = `
        <tr>
            ${dashboardData.headers.map(h => `<th>${h || ''}</th>`).join('')}
        </tr>
    `;
    
    // Render all rows (will be filtered by table filters)
    updateTableRows();
}

function updateTableRows() {
    const tableBody = document.getElementById('tableBody');
    if (!tableBody) return;
    
    // Get table filters
    const customerFilter = document.getElementById('table-filter-customer')?.value || '';
    const salesFilter = document.getElementById('table-filter-sales')?.value || '';
    const statusFilter = document.getElementById('table-filter-status')?.value || '';
    const dateFromFilter = document.getElementById('table-filter-date-from')?.value || '';
    const dateToFilter = document.getElementById('table-filter-date-to')?.value || '';
    
    const filters = {
        customer: customerFilter,
        sales: salesFilter,
        status: statusFilter,
        dateFrom: dateFromFilter,
        dateTo: dateToFilter
    };
    
    const filteredData = filterData(filters);
    
    tableBody.innerHTML = filteredData.map(item => {
        return `
            <tr>
                ${item.rawRow.map(cell => `<td>${cell || ''}</td>`).join('')}
            </tr>
        `;
    }).join('');
    
    if (filteredData.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="100%" class="loading">Tidak ada data yang sesuai filter</td></tr>';
    }
}

// ============================================
// Drag and Drop
// ============================================

function initDragAndDrop() {
    const grid = document.getElementById('dashboardGrid');
    if (!grid) return;
    
    // Remove existing listeners
    const existingWidgets = document.querySelectorAll('.widget');
    existingWidgets.forEach(widget => {
        widget.draggable = true;
    });
    
    // Re-initialize drag handlers
    grid.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        
        const dragging = document.querySelector('.dragging');
        if (!dragging) return;
        
        const afterElement = getDragAfterElement(grid, e.clientY);
        if (afterElement == null) {
            grid.appendChild(dragging);
        } else {
            grid.insertBefore(dragging, afterElement);
        }
    });
    
    // Add drag handlers to all widgets (including dynamically added ones)
    const observer = new MutationObserver(() => {
        document.querySelectorAll('.widget').forEach(widget => {
            if (!widget.hasAttribute('data-drag-initialized')) {
                widget.draggable = true;
                widget.setAttribute('data-drag-initialized', 'true');
                
                widget.addEventListener('dragstart', (e) => {
                    widget.classList.add('dragging');
                    e.dataTransfer.effectAllowed = 'move';
                    e.stopPropagation();
                });
                
                widget.addEventListener('dragend', () => {
                    widget.classList.remove('dragging');
                });
            }
        });
    });
    
    observer.observe(grid, { childList: true, subtree: true });
    
    // Initialize existing widgets
    document.querySelectorAll('.widget').forEach(widget => {
        widget.draggable = true;
        widget.setAttribute('data-drag-initialized', 'true');
        
        widget.addEventListener('dragstart', (e) => {
            widget.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
            e.stopPropagation();
        });
        
        widget.addEventListener('dragend', () => {
            widget.classList.remove('dragging');
        });
    });
}

function getDragAfterElement(container, y) {
    const draggableElements = [...container.querySelectorAll('.widget:not(.dragging)')];
    
    return draggableElements.reduce((closest, child) => {
        const box = child.getBoundingClientRect();
        const offset = y - box.top - box.height / 2;
        
        if (offset < 0 && offset > closest.offset) {
            return { offset: offset, element: child };
        } else {
            return closest;
        }
    }, { offset: Number.NEGATIVE_INFINITY }).element;
}

// ============================================
// Widget Toggle
// ============================================

function initWidgetToggles() {
    // Remove existing listeners to avoid duplicates
    document.querySelectorAll('.widget-toggle').forEach(button => {
        const newButton = button.cloneNode(true);
        button.parentNode.replaceChild(newButton, button);
    });
    
    document.querySelectorAll('.widget-toggle').forEach(button => {
        button.addEventListener('click', (e) => {
            e.stopPropagation();
            e.preventDefault();
            const widget = button.closest('.widget');
            widget.classList.toggle('collapsed');
            button.textContent = widget.classList.contains('collapsed') ? '+' : '−';
            
            // Recreate chart when expanded
            if (!widget.classList.contains('collapsed')) {
                const widgetId = widget.dataset.widgetId;
                if (widgetId === 'widget-rfq-customer-sales' && dashboardData.parsed.length > 0) {
                    createRFQCustomerSalesChart();
                } else if (widgetId === 'chart-rfq-per-customer-sales' && dashboardData.parsed.length > 0) {
                    createRFQPerCustomerSalesChart();
                } else if (widgetId === 'chart-top-customers' && dashboardData.parsed.length > 0) {
                    createTopCustomersChart();
                } else if (widgetId === 'chart-conversion-per-customer' && dashboardData.parsed.length > 0) {
                    createConversionPerCustomerChart();
                } else if (widgetId === 'chart-amount-per-customer-sales' && dashboardData.parsed.length > 0) {
                    createAmountPerCustomerSalesChart();
                } else if (widgetId === 'widget-top-customer-rfq-volume' && dashboardData.parsed.length > 0) {
                    createTopCustomerRFQVolumeChart();
                }
            }
        });
    });
}

// ============================================
// Filter Event Listeners
// ============================================

function initFilterListeners() {
    // Chart filter listeners
    ['1', '2', '3', '4'].forEach(id => {
        ['customer', 'sales', 'date-from', 'date-to'].forEach(filterType => {
            const filterId = `filter${id}-${filterType}`;
            const filterEl = document.getElementById(filterId);
            if (filterEl) {
                filterEl.addEventListener('change', () => {
                    if (id === '1') createRFQPerCustomerSalesChart();
                    else if (id === '2') createTopCustomersChart();
                    else if (id === '3') createConversionPerCustomerChart();
                    else if (id === '4') createAmountPerCustomerSalesChart();
                });
            }
        });
    });
    
    // Table filter listeners
    ['customer', 'sales', 'status', 'date-from', 'date-to'].forEach(filterType => {
        const filterEl = document.getElementById(`table-filter-${filterType}`);
        if (filterEl) {
            filterEl.addEventListener('input', updateTableRows);
        }
    });
    
    // Reset button
    const resetBtn = document.getElementById('table-filter-reset');
    if (resetBtn) {
        resetBtn.addEventListener('click', () => {
            document.getElementById('table-filter-customer').value = '';
            document.getElementById('table-filter-sales').value = '';
            document.getElementById('table-filter-status').value = '';
            document.getElementById('table-filter-date-from').value = '';
            document.getElementById('table-filter-date-to').value = '';
            updateTableRows();
        });
    }
}

// ============================================
// Error Handling
// ============================================

function showError(message) {
    const grid = document.getElementById('dashboardGrid');
    const errorDiv = document.createElement('div');
    errorDiv.className = 'widget';
    errorDiv.style.gridColumn = '1 / -1';
    errorDiv.style.borderColor = '#ef4444';
    errorDiv.innerHTML = `
        <div class="widget-content" style="text-align: center; padding: 32px; color: #ef4444;">
            <p style="font-weight: 500; margin-bottom: 8px;">Error</p>
            <p style="font-size: 14px; color: #666;">${message}</p>
        </div>
    `;
    grid.insertBefore(errorDiv, grid.firstChild);
}

// ============================================
// Login & Authentication
// ============================================

// Hard-coded credentials (front-end only, not secure)
const LOGIN_CREDENTIALS = {
    username: 'stm',
    password: 'cigoongthebest123'
};

function checkLoginStatus() {
    const isLoggedIn = localStorage.getItem('loggedIn') === 'true';
    const loginScreen = document.getElementById('loginScreen');
    const dashboardContainer = document.getElementById('dashboardContainer');
    
    if (isLoggedIn) {
        // Show dashboard
        if (loginScreen) loginScreen.style.display = 'none';
        if (dashboardContainer) dashboardContainer.style.display = 'block';
    } else {
        // Show login screen
        if (loginScreen) loginScreen.style.display = 'flex';
        if (dashboardContainer) dashboardContainer.style.display = 'none';
    }
    
    return isLoggedIn;
}

function handleLogin(event) {
    event.preventDefault();
    
    const username = document.getElementById('username').value.trim();
    const password = document.getElementById('password').value;
    const errorDiv = document.getElementById('loginError');
    
    // Validate credentials
    if (username === LOGIN_CREDENTIALS.username && password === LOGIN_CREDENTIALS.password) {
        // Save login status
        localStorage.setItem('loggedIn', 'true');
        
        // Hide error if shown
        if (errorDiv) errorDiv.style.display = 'none';
        
        // Show dashboard
        checkLoginStatus();
        
        // Clear form
        document.getElementById('loginForm').reset();
    } else {
        // Show error
        if (errorDiv) {
            errorDiv.style.display = 'block';
            errorDiv.textContent = 'Username atau password salah';
        }
        
        // Shake animation for error
        const loginCard = document.querySelector('.login-card');
        if (loginCard) {
            loginCard.style.animation = 'shake 0.3s';
            setTimeout(() => {
                loginCard.style.animation = '';
            }, 300);
        }
    }
}

function handleLogout() {
    // Clear login status
    localStorage.removeItem('loggedIn');
    
    // Show login screen
    checkLoginStatus();
}

function initLogin() {
    // Check login status on page load
    checkLoginStatus();
    
    // Setup login form handler
    const loginForm = document.getElementById('loginForm');
    if (loginForm) {
        loginForm.addEventListener('submit', handleLogin);
    }
    
    // Setup logout button
    const logoutButton = document.getElementById('logoutButton');
    if (logoutButton) {
        logoutButton.addEventListener('click', handleLogout);
    }
    
    // Allow Enter key to submit login form
    document.getElementById('username')?.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            document.getElementById('password')?.focus();
        }
    });
    
    document.getElementById('password')?.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            loginForm?.dispatchEvent(new Event('submit'));
        }
    });
}

// ============================================
// Insights Summary
// ============================================

function showInsightsSummary() {
    const insightsContainer = document.getElementById('insightsSummary');
    if (!insightsContainer || !dashboardData.parsed || dashboardData.parsed.length === 0) {
        return;
    }
    
    // 1. RFQ per Customer per Sales
    const kpi1 = aggregateRFQPerCustomerSales(dashboardData.parsed);
    const topRFQCustomerSales = kpi1.slice(0, 10);
    
    // 2. Customer dengan RFQ terbanyak
    const kpi2 = aggregateTopCustomers(dashboardData.parsed);
    const topCustomers = kpi2.slice(0, 10);
    
    // 3. Prosentase konversi per Customer
    const kpi3 = aggregateConversionPerCustomer(dashboardData.parsed);
    const topConversionCustomers = kpi3
        .filter(item => item.total > 0)
        .slice(0, 10);
    
    // 4. Total Amount RFQ per Customer per Sales
    const kpi4 = aggregateAmountPerCustomerSales(dashboardData.parsed);
    const topAmountCustomerSales = kpi4.slice(0, 10);
    
    let html = `
        <div class="insight-section">
            <h4>1. RFQ per Customer per Sales</h4>
            <p style="margin-bottom: 12px; color: var(--text-secondary); font-size: var(--font-size-sm);">
                Total kombinasi: <strong>${kpi1.length.toLocaleString('id-ID')}</strong>
            </p>
            <table class="insight-table">
                <thead>
                    <tr>
                        <th style="width: 40%;">Customer</th>
                        <th style="width: 25%;">Sales</th>
                        <th style="width: 35%; text-align: right;">Jumlah RFQ</th>
                    </tr>
                </thead>
                <tbody>
                    ${topRFQCustomerSales.map(item => `
                        <tr>
                            <td>${item.customer.length > 30 ? item.customer.substring(0, 27) + '...' : item.customer}</td>
                            <td>${item.sales.length > 20 ? item.sales.substring(0, 17) + '...' : item.sales}</td>
                            <td style="text-align: right; font-weight: var(--font-weight-semibold);">${item.count.toLocaleString('id-ID')}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
        
        <div class="insight-section">
            <h4>2. Customer dengan RFQ Terbanyak</h4>
            <p style="margin-bottom: 12px; color: var(--text-secondary); font-size: var(--font-size-sm);">
                Total customer: <strong>${kpi2.length.toLocaleString('id-ID')}</strong>
            </p>
            <table class="insight-table">
                <thead>
                    <tr>
                        <th style="width: 70%;">Customer</th>
                        <th style="width: 30%; text-align: right;">Jumlah RFQ</th>
                    </tr>
                </thead>
                <tbody>
                    ${topCustomers.map(item => `
                        <tr>
                            <td>${item.customer.length > 50 ? item.customer.substring(0, 47) + '...' : item.customer}</td>
                            <td style="text-align: right; font-weight: var(--font-weight-semibold);">${item.count.toLocaleString('id-ID')}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
        
        <div class="insight-section">
            <h4>3. Prosentase Konversi RFQ menjadi Order per Customer</h4>
            <p style="margin-bottom: 12px; color: var(--text-secondary); font-size: var(--font-size-sm);">
                Customer dengan konversi tertinggi (minimal 1 RFQ)
            </p>
            <table class="insight-table">
                <thead>
                    <tr>
                        <th style="width: 50%;">Customer</th>
                        <th style="width: 15%; text-align: center;">Total RFQ</th>
                        <th style="width: 15%; text-align: center;">Converted</th>
                        <th style="width: 20%; text-align: right;">% Konversi</th>
                    </tr>
                </thead>
                <tbody>
                    ${topConversionCustomers.map(item => `
                        <tr>
                            <td>${item.customer.length > 40 ? item.customer.substring(0, 37) + '...' : item.customer}</td>
                            <td style="text-align: center;">${item.total.toLocaleString('id-ID')}</td>
                            <td style="text-align: center; color: ${item.converted > 0 ? '#10b981' : 'var(--text-secondary)'};">
                                <strong>${item.converted.toLocaleString('id-ID')}</strong>
                            </td>
                            <td style="text-align: right; font-weight: var(--font-weight-semibold); color: ${item.conversionRate >= 50 ? '#10b981' : item.conversionRate >= 25 ? '#f59e0b' : 'var(--text-primary)'};">
                                ${item.conversionRate.toFixed(2)}%
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
        
        <div class="insight-section">
            <h4>4. Total Amount RFQ per Customer per Sales</h4>
            <p style="margin-bottom: 12px; color: var(--text-secondary); font-size: var(--font-size-sm);">
                Total kombinasi: <strong>${kpi4.length.toLocaleString('id-ID')}</strong> | 
                Total amount: <strong>Rp ${formatCurrency(kpi4.reduce((sum, r) => sum + r.amount, 0))}</strong>
            </p>
            <table class="insight-table">
                <thead>
                    <tr>
                        <th style="width: 40%;">Customer</th>
                        <th style="width: 25%;">Sales</th>
                        <th style="width: 35%; text-align: right;">Total Amount</th>
                    </tr>
                </thead>
                <tbody>
                    ${topAmountCustomerSales.map(item => `
                        <tr>
                            <td>${item.customer.length > 30 ? item.customer.substring(0, 27) + '...' : item.customer}</td>
                            <td>${item.sales.length > 20 ? item.sales.substring(0, 17) + '...' : item.sales}</td>
                            <td style="text-align: right; font-weight: var(--font-weight-semibold);">Rp ${formatCurrency(item.amount)}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;
    
    insightsContainer.innerHTML = html;
}

// ============================================
// Initialization
// ============================================

function initializeDashboard() {
    // Only initialize if logged in
    if (!checkLoginStatus()) {
        return;
    }
    
    // Populate filter options
    populateFilterOptions();
    
    // Initialize widget interactions
    initDragAndDrop();
    initWidgetToggles();
    
    // Create charts
    if (dashboardData.parsed.length > 0) {
        createRFQCustomerSalesChart();
        createTopCustomerRFQVolumeChart();
        createConversionRateCustomerChart();
    }
    
    // Populate filter dropdowns for Conversion Rate widget
    populateConversionRateFilters();
    
    // Setup filter listeners for Conversion Rate widget
    const conversionStatusFilter = document.getElementById('conversion-status-filter');
    const conversionDateFrom = document.getElementById('conversion-date-from');
    const conversionDateTo = document.getElementById('conversion-date-to');
    const conversionSalesFilter = document.getElementById('conversion-sales-filter');
    const conversionCustomerFilter = document.getElementById('conversion-customer-filter');
    const conversionShowCount = document.getElementById('conversion-show-count');
    
    if (conversionStatusFilter) {
        conversionStatusFilter.addEventListener('change', () => {
            createConversionRateCustomerChart();
            const tableContainer = document.getElementById('conversion-rate-table-container');
            if (tableContainer) tableContainer.style.display = 'none';
        });
    }
    
    if (conversionDateFrom) {
        conversionDateFrom.addEventListener('change', () => {
            populateConversionRateFilters({ preserveSalesSelection: true });
            createConversionRateCustomerChart();
            const tableContainer = document.getElementById('conversion-rate-table-container');
            if (tableContainer) tableContainer.style.display = 'none';
        });
    }

    if (conversionDateTo) {
        conversionDateTo.addEventListener('change', () => {
            populateConversionRateFilters({ preserveSalesSelection: true });
            createConversionRateCustomerChart();
            const tableContainer = document.getElementById('conversion-rate-table-container');
            if (tableContainer) tableContainer.style.display = 'none';
        });
    }

    if (conversionSalesFilter) {
        conversionSalesFilter.addEventListener('change', () => {
            populateConversionRateCustomerDropdown({ preserveSelection: false });
            createConversionRateCustomerChart();
            const tableContainer = document.getElementById('conversion-rate-table-container');
            if (tableContainer) tableContainer.style.display = 'none';
        });
    }
    
    if (conversionCustomerFilter) {
        conversionCustomerFilter.addEventListener('change', () => {
            createConversionRateCustomerChart();
            const tableContainer = document.getElementById('conversion-rate-table-container');
            if (tableContainer) tableContainer.style.display = 'none';
        });
    }
    
    if (conversionShowCount) {
        conversionShowCount.addEventListener('change', () => {
            const value = parseInt(conversionShowCount.value);
            if (value && value > 0 && value <= 1000) {
                createConversionRateCustomerChart();
                const tableContainer = document.getElementById('conversion-rate-table-container');
                if (tableContainer) tableContainer.style.display = 'none';
            } else {
                conversionShowCount.value = 20;
            }
        });
    }
    
    // Setup filter listeners for RFQ Customer Sales widget
    const dateFromInput = document.getElementById('rfq-customer-sales-date-from');
    const dateToInput = document.getElementById('rfq-customer-sales-date-to');
    const customerFilter = document.getElementById('rfq-customer-sales-customer');
    const salesFilter = document.getElementById('rfq-customer-sales-sales');
    const viewAllBtn = document.getElementById('rfq-customer-sales-view-all');
    
    // Setup chart update listeners for all filters
    [dateFromInput, dateToInput, customerFilter].forEach(el => {
        if (el) {
            el.addEventListener('change', () => {
                createRFQCustomerSalesChart();
                // Hide table when filter changes
                const tableContainer = document.getElementById('rfq-customer-sales-table-container');
                if (tableContainer) tableContainer.style.display = 'none';
            });
        }
    });
    
    // Status filter checkboxes
    const statusJadiOC = document.getElementById('rfq-customer-sales-status-jadi-oc');
    const statusTidakJadiOC = document.getElementById('rfq-customer-sales-status-tidak-jadi-oc');
    
    [statusJadiOC, statusTidakJadiOC].forEach(checkbox => {
        if (checkbox) {
            checkbox.addEventListener('change', () => {
                // Ensure at least one is checked
                if (!statusJadiOC.checked && !statusTidakJadiOC.checked) {
                    // If both unchecked, re-check the default (JADI OC)
                    statusJadiOC.checked = true;
                }
                
                // Preserve current filter values before updating
                const salesFilter = document.getElementById('rfq-customer-sales-sales');
                const customerFilter = document.getElementById('rfq-customer-sales-customer');
                const savedSalesValue = salesFilter ? salesFilter.value : '';
                const savedCustomerValue = customerFilter ? customerFilter.value : '';
                
                // Update customer dropdown counts based on status filter (only if sales is selected)
                // This updates the counts in customer dropdown without resetting selections
                if (salesFilter && savedSalesValue && customerFilter) {
                    // Get current status checkbox states
                    const statusJadiOC = document.getElementById('rfq-customer-sales-status-jadi-oc');
                    const statusTidakJadiOC = document.getElementById('rfq-customer-sales-status-tidak-jadi-oc');
                    const includeJadiOC = statusJadiOC?.checked || false;
                    const includeTidakJadiOC = statusTidakJadiOC?.checked || false;
                    
                    // Get all unique customers
                    const allCustomers = [...new Set(dashboardData.parsed.map(d => d.customer ? d.customer.trim() : '').filter(c => c))].sort();
                    
                    // Count RFQ per customer for selected sales (with status filter)
                    const customerCounts = {};
                    const selectedSalesNormalized = savedSalesValue.trim().toLowerCase();
                    
                    dashboardData.parsed.forEach(item => {
                        const itemSales = item.sales ? item.sales.trim() : '';
                        const itemCustomer = item.customer ? item.customer.trim() : '';
                        
                        if (itemSales.toLowerCase() === selectedSalesNormalized && itemCustomer) {
                            let include = false;
                            if (includeJadiOC && includeTidakJadiOC) {
                                include = true;
                            } else if (includeJadiOC && !includeTidakJadiOC) {
                                include = item.isConverted;
                            } else if (!includeJadiOC && includeTidakJadiOC) {
                                include = !item.isConverted;
                            }
                            
                            if (include) {
                                customerCounts[itemCustomer] = (customerCounts[itemCustomer] || 0) + 1;
                            }
                        }
                    });
                    
                    // Filter customers that have count > 0, then sort by count (descending), then alphabetically
                    const sortedCustomers = allCustomers.filter(c => customerCounts[c] > 0).sort((a, b) => {
                        const countA = customerCounts[a] || 0;
                        const countB = customerCounts[b] || 0;
                        if (countB !== countA) return countB - countA;
                        return a.localeCompare(b);
                    });
                    
                    // Build new options - always start with "Semua Customer"
                    let customerOptions = '<option value="">Semua Customer</option>';
                    
                    // Only add customer options if there are customers with count > 0
                    if (sortedCustomers.length > 0) {
                        customerOptions += sortedCustomers.map(c => {
                            const count = customerCounts[c] || 0;
                            const escapedCustomer = c.replace(/"/g, '&quot;');
                            return `<option value="${escapedCustomer}">${c} (${count})</option>`;
                        }).join('');
                    }
                    // If sortedCustomers is empty, dropdown will only show "Semua Customer"
                    
                    // Update dropdown and restore selection
                    customerFilter.innerHTML = customerOptions;
                    // Always reset to "Semua Customer" when dropdown is updated based on status filter
                    // This ensures the dropdown shows the correct state (empty if no matches)
                    customerFilter.value = '';
                    
                    // Only try to restore if the saved customer still exists in the filtered list
                    if (savedCustomerValue && sortedCustomers.length > 0) {
                        const option = Array.from(customerFilter.options).find(opt => opt.value === savedCustomerValue);
                        if (option) {
                            customerFilter.value = savedCustomerValue;
                        }
                    }
                } else if (customerFilter && !savedSalesValue) {
                    // If no sales is selected but customer dropdown exists, make sure it shows all customers
                    // This shouldn't happen often, but just in case
                    const allCustomers = [...new Set(dashboardData.parsed.map(d => d.customer ? d.customer.trim() : '').filter(c => c))].sort();
                    let customerOptions = '<option value="">Semua Customer</option>';
                    customerOptions += allCustomers.map(c => {
                        const escapedCustomer = c.replace(/"/g, '&quot;');
                        return `<option value="${escapedCustomer}">${c}</option>`;
                    }).join('');
                    customerFilter.innerHTML = customerOptions;
                    // Try to restore saved customer value
                    if (savedCustomerValue) {
                        const option = Array.from(customerFilter.options).find(opt => opt.value === savedCustomerValue);
                        if (option) {
                            customerFilter.value = savedCustomerValue;
                        }
                    }
                }
                
                createRFQCustomerSalesChart();
                // Hide table when filter changes
                const tableContainer = document.getElementById('rfq-customer-sales-table-container');
                if (tableContainer) tableContainer.style.display = 'none';
            });
        }
    });
    
    // Sales filter chart update - get fresh reference after populateFilterOptions
    const salesFilterAfterInit = document.getElementById('rfq-customer-sales-sales');
    if (salesFilterAfterInit) {
        salesFilterAfterInit.addEventListener('change', () => {
            createRFQCustomerSalesChart();
            // Hide table when filter changes
            const tableContainer = document.getElementById('rfq-customer-sales-table-container');
            if (tableContainer) tableContainer.style.display = 'none';
        });
    }
    
    if (viewAllBtn) {
        viewAllBtn.addEventListener('click', showAllFilteredRows);
    }
    
    // Toggle axis button
    const toggleAxisBtn = document.getElementById('rfq-customer-sales-toggle-axis');
    if (toggleAxisBtn) {
        toggleAxisBtn.addEventListener('click', () => {
            // Toggle orientation
            chartAxisOrientation = chartAxisOrientation === 'customer-sales' ? 'sales-customer' : 'customer-sales';
            // Recreate chart with new orientation
            createRFQCustomerSalesChart();
            // Hide table when toggled
            const tableContainer = document.getElementById('rfq-customer-sales-table-container');
            if (tableContainer) tableContainer.style.display = 'none';
        });
    }
    
    // Toggle simple mode button
    const toggleSimpleBtn = document.getElementById('rfq-customer-sales-toggle-simple');
    if (toggleSimpleBtn) {
        toggleSimpleBtn.addEventListener('click', () => {
            // Toggle simple mode
            chartSimpleMode = !chartSimpleMode;
            // Update button text
            toggleSimpleBtn.textContent = chartSimpleMode ? '📊 Normal Mode' : '✨ Simple Mode';
            toggleSimpleBtn.title = chartSimpleMode ? 'Switch to Normal Mode' : 'Toggle Simple Mode (Highlight Biggest Value)';
            // Recreate chart with new mode
            createRFQCustomerSalesChart();
            // Hide table when toggled
            const tableContainer = document.getElementById('rfq-customer-sales-table-container');
            if (tableContainer) tableContainer.style.display = 'none';
        });
    }
    
    // Toggle show all button
    const toggleAllBtn = document.getElementById('rfq-customer-sales-toggle-all');
    if (toggleAllBtn) {
        toggleAllBtn.addEventListener('click', () => {
            // Toggle show all
            chartShowAll = !chartShowAll;
            // Update button text
            const isCustomerSales = chartAxisOrientation === 'customer-sales';
            const showCountInput = document.getElementById('rfq-customer-sales-show-count');
            const showCount = showCountInput ? parseInt(showCountInput.value) || 15 : 15;
            if (chartShowAll) {
                toggleAllBtn.textContent = `📊 Top ${showCount}`;
                toggleAllBtn.title = `Show Top ${showCount} Only`;
            } else {
                toggleAllBtn.textContent = '📈 Show All';
                toggleAllBtn.title = 'Show All ' + (isCustomerSales ? 'Customers' : 'Sales');
            }
            // Recreate chart with new mode
            createRFQCustomerSalesChart();
            // Hide table when toggled
            const tableContainer = document.getElementById('rfq-customer-sales-table-container');
            if (tableContainer) tableContainer.style.display = 'none';
        });
    }
    
    // Setup show count input listener for RFQ Customer Sales widget
    const rfqCustomerSalesShowCountInput = document.getElementById('rfq-customer-sales-show-count');
    if (rfqCustomerSalesShowCountInput) {
        rfqCustomerSalesShowCountInput.addEventListener('change', () => {
            const value = parseInt(rfqCustomerSalesShowCountInput.value);
            if (value && value > 0 && value <= 1000) {
                rfqCustomerSalesShowCount = value;
                createRFQCustomerSalesChart();
                // Hide table when count changes
                const tableContainer = document.getElementById('rfq-customer-sales-table-container');
                if (tableContainer) tableContainer.style.display = 'none';
            } else {
                // Reset to default if invalid
                rfqCustomerSalesShowCountInput.value = rfqCustomerSalesShowCount;
            }
        });
    }
    
    // Setup filter listeners for Top Customer RFQ Volume widget
    const topCustomerStatusJadiOC = document.getElementById('top-customer-status-jadi-oc');
    const topCustomerStatusTidakJadiOC = document.getElementById('top-customer-status-tidak-jadi-oc');
    const topCustomerShowCountInput = document.getElementById('top-customer-show-count');
    
    [topCustomerStatusJadiOC, topCustomerStatusTidakJadiOC].forEach(checkbox => {
        if (checkbox) {
            checkbox.addEventListener('change', () => {
                // Ensure at least one is checked
                const jadiOC = topCustomerStatusJadiOC?.checked || false;
                const tidakJadiOC = topCustomerStatusTidakJadiOC?.checked || false;
                
                if (!jadiOC && !tidakJadiOC) {
                    // If both unchecked, re-check the default (JADI OC)
                    if (topCustomerStatusJadiOC) topCustomerStatusJadiOC.checked = true;
                }
                
                createTopCustomerRFQVolumeChart();
                // Hide table when filter changes
                const tableContainer = document.getElementById('top-customer-rfq-table-container');
                if (tableContainer) tableContainer.style.display = 'none';
            });
        }
    });
    
    // Setup show count input listener for Top Customer RFQ Volume widget
    if (topCustomerShowCountInput) {
        topCustomerShowCountInput.addEventListener('change', () => {
            const value = parseInt(topCustomerShowCountInput.value);
            if (value && value > 0 && value <= 1000) {
                topCustomerRFQVolumeShowCount = value;
                createTopCustomerRFQVolumeChart();
                // Hide table when count changes
                const tableContainer = document.getElementById('top-customer-rfq-table-container');
                if (tableContainer) tableContainer.style.display = 'none';
            } else {
                // Reset to default if invalid
                topCustomerShowCountInput.value = topCustomerRFQVolumeShowCount;
            }
        });
    }
    
    console.log('Dashboard ready. Data available:', dashboardData.parsed.length, 'records');
}

function getConversionFilterBaseData() {
    if (!dashboardData.parsed || dashboardData.parsed.length === 0) return [];

    const dateFromInput = document.getElementById('conversion-date-from');
    const dateToInput = document.getElementById('conversion-date-to');
    const parsedDateFrom = dateFromInput?.value ? new Date(dateFromInput.value) : null;
    const parsedDateTo = dateToInput?.value ? new Date(dateToInput.value) : null;

    if (parsedDateTo) {
        parsedDateTo.setHours(23, 59, 59, 999);
    }

    return dashboardData.parsed.filter((item) => {
        if (!parsedDateFrom && !parsedDateTo) return true;
        if (!item.date) return false;
        const itemDate = item.date instanceof Date ? item.date : new Date(item.date);
        if (parsedDateFrom && itemDate < parsedDateFrom) return false;
        if (parsedDateTo && itemDate > parsedDateTo) return false;
        return true;
    });
}

function populateConversionRateFilters({ preserveSalesSelection = true } = {}) {
    if (!dashboardData.parsed || dashboardData.parsed.length === 0) return;

    const salesFilter = document.getElementById('conversion-sales-filter');
    if (salesFilter) {
        const baseData = getConversionFilterBaseData();
        const counts = {};

        baseData.forEach((item) => {
            const salesName = (item.sales || '').trim();
            if (!salesName) return;
            counts[salesName] = (counts[salesName] || 0) + 1;
        });

        const entries = Object.entries(counts)
            .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));

        const previousValue = preserveSalesSelection ? salesFilter.value : '';
        let salesOptions = '<option value="">Semua Sales</option>';
        entries.forEach(([name, count]) => {
            const escapedSales = name.replace(/"/g, '&quot;');
            salesOptions += `<option value="${escapedSales}">${formatOptionLabel(name, count)}</option>`;
        });

        salesFilter.innerHTML = salesOptions;
        if (previousValue && counts[previousValue]) {
            salesFilter.value = previousValue;
        } else if (!preserveSalesSelection) {
            salesFilter.value = '';
        }
    }

    populateConversionRateCustomerDropdown({ preserveSelection: true });
}

function populateConversionRateCustomerDropdown({ preserveSelection = true } = {}) {
    const customerFilter = document.getElementById('conversion-customer-filter');
    const salesFilter = document.getElementById('conversion-sales-filter');

    if (!customerFilter) return;

    const baseData = getConversionFilterBaseData();
    const selectedSales = salesFilter?.value?.trim().toLowerCase() || '';
    const counts = {};

    baseData.forEach((item) => {
        const customerName = (item.customer || '').trim();
        if (!customerName) return;
        if (selectedSales) {
            const itemSales = (item.sales || '').trim().toLowerCase();
            if (itemSales !== selectedSales) return;
        }
        counts[customerName] = (counts[customerName] || 0) + 1;
    });

    const entries = Object.entries(counts)
        .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));

    const previousValue = preserveSelection ? customerFilter.value : '';
    let customerOptions = '<option value="">Semua Customer</option>';
    entries.forEach(([name, count]) => {
        const escapedCustomer = name.replace(/"/g, '&quot;');
        customerOptions += `<option value="${escapedCustomer}">${formatOptionLabel(name, count)}</option>`;
    });

    customerFilter.innerHTML = customerOptions;

    if (previousValue && counts[previousValue]) {
        customerFilter.value = previousValue;
    } else if (!preserveSelection) {
        customerFilter.value = '';
    }
}

// ============================================
// File Input Handler
// ============================================

document.addEventListener('DOMContentLoaded', async function() {
    // Initialize login system first
    initLogin();
    
    // Try to auto-load CSV file if exists
    if (checkLoginStatus()) {
        const csvData = await loadCSVFile('./RECAP PENAWARAN 2025.csv');
        if (csvData && csvData.headers.length > 0) {
            processData(csvData.headers, csvData.rows);
            const fileName = document.getElementById('fileName');
            if (fileName) fileName.textContent = 'RECAP PENAWARAN 2025.csv (loaded)';
        }
    }
    
    // Setup file input handler
    const fileInput = document.getElementById('excelFile');
    const fileName = document.getElementById('fileName');
    
    if (fileInput) {
        fileInput.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (file) {
                fileName.textContent = file.name;
                
                // Check file type
                if (file.name.endsWith('.csv')) {
                    const reader = new FileReader();
                    reader.onload = function(e) {
                        const csvData = parseCSV(e.target.result);
                        if (csvData) {
                            processData(csvData.headers, csvData.rows);
                        }
                    };
                    reader.readAsText(file);
                } else {
                    // Excel file
                    loadExcelFile(file);
                }
            }
        });
    }
});
const rfqFilterState = {
    sales: '',
    customer: ''
};

function getRFQBaseFilteredData(overrides = {}) {
    const dateFromInput = document.getElementById('rfq-customer-sales-date-from');
    const dateToInput = document.getElementById('rfq-customer-sales-date-to');
    const statusConverted = document.getElementById('rfq-customer-sales-status-jadi-oc');
    const statusNotConverted = document.getElementById('rfq-customer-sales-status-tidak-jadi-oc');

    const dateFrom = overrides.dateFrom || dateFromInput?.value || '';
    const dateTo = overrides.dateTo || dateToInput?.value || '';
    const includeConverted = overrides.includeConverted !== undefined ? overrides.includeConverted : (statusConverted?.checked ?? true);
    const includeNotConverted = overrides.includeNotConverted !== undefined ? overrides.includeNotConverted : (statusNotConverted?.checked ?? true);

    const parsedDateFrom = dateFrom ? new Date(dateFrom) : null;
    const parsedDateTo = dateTo ? new Date(dateTo) : null;
    if (parsedDateTo) parsedDateTo.setHours(23, 59, 59, 999);

    return dashboardData.parsed.filter((item) => {
        if (!includeConverted && item.isConverted) return false;
        if (!includeNotConverted && !item.isConverted) return false;

        if (parsedDateFrom && (!item.date || item.date < parsedDateFrom)) return false;
        if (parsedDateTo && (!item.date || item.date > parsedDateTo)) return false;

        return true;
    });
}

function formatOptionLabel(name, count) {
    if (!name) return '';
    return `${name} (${count} RFQ)`;
}

function updateRFQSalesOptions({ preserveSelection = true } = {}) {
    const salesSelect = document.getElementById('rfq-customer-sales-sales');
    const customerValue = rfqFilterState.customer?.trim().toLowerCase();
    if (!salesSelect) return;

    const baseData = getRFQBaseFilteredData();
    const counts = {};

    baseData.forEach((item) => {
        const salesName = (item.sales || '').trim();
        if (!salesName) return;
        if (customerValue && (!item.customer || item.customer.trim().toLowerCase() !== customerValue)) return;
        counts[salesName] = (counts[salesName] || 0) + 1;
    });

    const entries = Object.entries(counts)
        .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));

    const previousValue = preserveSelection ? salesSelect.value : '';
    let optionsHtml = '<option value="">Semua Sales</option>';
    entries.forEach(([name, count]) => {
        const escaped = name.replace(/"/g, '&quot;');
        optionsHtml += `<option value="${escaped}">${formatOptionLabel(name, count)}</option>`;
    });

    salesSelect.innerHTML = optionsHtml;
    if (previousValue && counts[previousValue]) {
        salesSelect.value = previousValue;
    } else if (!preserveSelection) {
        salesSelect.value = '';
    }
}

function updateRFQCustomerOptions({ preserveSelection = true } = {}) {
    const customerSelect = document.getElementById('rfq-customer-sales-customer');
    const salesValue = rfqFilterState.sales?.trim().toLowerCase();
    if (!customerSelect) return;

    const baseData = getRFQBaseFilteredData();
    const counts = {};

    baseData.forEach((item) => {
        const customerName = (item.customer || '').trim();
        if (!customerName) return;
        if (salesValue && (!item.sales || item.sales.trim().toLowerCase() !== salesValue)) return;
        counts[customerName] = (counts[customerName] || 0) + 1;
    });

    const entries = Object.entries(counts)
        .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));

    const previousValue = preserveSelection ? customerSelect.value : '';
    let optionsHtml = '<option value="">Semua Customer</option>';
    entries.forEach(([name, count]) => {
        const escaped = name.replace(/"/g, '&quot;');
        optionsHtml += `<option value="${escaped}">${formatOptionLabel(name, count)}</option>`;
    });

    customerSelect.innerHTML = optionsHtml;
    if (previousValue && counts[previousValue]) {
        customerSelect.value = previousValue;
    } else if (!preserveSelection) {
        customerSelect.value = '';
    }
}

function refreshRFQLinkedDropdowns({ preserveSelections = true } = {}) {
    updateRFQSalesOptions({ preserveSelection: preserveSelections });
    updateRFQCustomerOptions({ preserveSelection: preserveSelections });
}

// ============================================
// Smart RFQ Dropdowns (bidirectional, with counts)
// ============================================
(function setupSmartRFQDropdowns(){
    const salesSelect = document.getElementById('rfq-customer-sales-sales');
    const customerSelect = document.getElementById('rfq-customer-sales-customer');
    const dateFrom = document.getElementById('rfq-customer-sales-date-from');
    const dateTo = document.getElementById('rfq-customer-sales-date-to');
    const statusYes = document.getElementById('rfq-customer-sales-status-jadi-oc');
    const statusNo = document.getElementById('rfq-customer-sales-status-tidak-jadi-oc');

    if (!salesSelect || !customerSelect) return;

    const hideDetailTable = () => {
        const tableContainer = document.getElementById('rfq-customer-sales-table-container');
        if (tableContainer) tableContainer.style.display = 'none';
    };

    // Initialize state from current selects
    rfqFilterState.sales = (salesSelect.value || '').trim();
    rfqFilterState.customer = (customerSelect.value || '').trim();

    // Initial population with current context
    try { refreshRFQLinkedDropdowns({ preserveSelections: true }); } catch(e) {}

    // Attach listeners
    salesSelect.addEventListener('change', () => {
        rfqFilterState.sales = (salesSelect.value || '').trim();
        // Update the linked customer list
        try { updateRFQCustomerOptions({ preserveSelection: true }); } catch(e) {}
        // Re-render the chart and hide detail table
        try { createRFQCustomerSalesChart(); } catch(e) {}
        hideDetailTable();
    });

    customerSelect.addEventListener('change', () => {
        rfqFilterState.customer = (customerSelect.value || '').trim();
        // Update the linked sales list
        try { updateRFQSalesOptions({ preserveSelection: true }); } catch(e) {}
        // Re-render the chart and hide detail table
        try { createRFQCustomerSalesChart(); } catch(e) {}
        hideDetailTable();
    });

    const refreshOnContext = () => {
        try { refreshRFQLinkedDropdowns({ preserveSelections: true }); } catch(e) {}
    };

    if (dateFrom) dateFrom.addEventListener('change', refreshOnContext);
    if (dateTo) dateTo.addEventListener('change', refreshOnContext);
    if (statusYes) statusYes.addEventListener('change', refreshOnContext);
    if (statusNo) statusNo.addEventListener('change', refreshOnContext);
})();
