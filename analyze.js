// CSV Analysis Script
// Load and analyze RECAP PENAWARAN 2025.csv

async function loadCSV() {
    const response = await fetch('./RECAP PENAWARAN 2025.csv');
    const text = await response.text();
    return text;
}

function parseCSV(text) {
    const lines = text.split('\n').filter(line => line.trim());
    if (lines.length === 0) return { headers: [], rows: [] };
    
    // Parse header - handle quoted fields
    const headers = parseCSVLine(lines[0]);
    
    // Parse data rows
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

function detectFieldTypes(headers, rows) {
    const types = {};
    const sampleSize = Math.min(100, rows.length);
    
    headers.forEach((header, idx) => {
        const sample = rows.slice(0, sampleSize).map(row => row[idx] || '');
        const nonEmpty = sample.filter(v => v && v.trim());
        
        if (nonEmpty.length === 0) {
            types[header] = 'empty';
            return;
        }
        
        // Check for date patterns
        if (nonEmpty.some(v => /^\d{2}-\w{3}-\d{2}/.test(v.trim()))) {
            types[header] = 'date';
            return;
        }
        
        // Check for numeric (price values with commas/spaces)
        const numericPattern = /^[\d\s,\.\(\)]+$/;
        const pricePattern = /^[\d\s,\.]+$/;
        if (nonEmpty.some(v => {
            const cleaned = v.replace(/[^\d,\.]/g, '');
            return cleaned && !isNaN(parseFloat(cleaned.replace(/,/g, '')));
        })) {
            types[header] = 'numeric';
            return;
        }
        
        // Default to string
        types[header] = 'string';
    });
    
    return types;
}

function identifyColumns(headers) {
    const mapping = {
        rfqValue: null,
        customerName: null,
        salesName: null,
        rfqStatus: null,
        orderConversion: null
    };
    
    headers.forEach(header => {
        const h = header.toUpperCase().trim();
        
        // RFQ Value - look for HARGA (NEW) first, then HARGA, PRICE, VALUE
        if (h.includes('HARGA') && h.includes('NEW')) {
            mapping.rfqValue = header;
        } else if (!mapping.rfqValue && (h.includes('HARGA') || h.includes('PRICE') || h.includes('VALUE'))) {
            mapping.rfqValue = header;
        }
        
        // Customer Name
        if (h.includes('CUSTOMER') && h.includes('NAME')) {
            mapping.customerName = header;
        } else if (!mapping.customerName && (h.includes('CUSTOMER') || h.includes('CLIENT'))) {
            mapping.customerName = header;
        }
        
        // Sales Name - MARKERTING (typo in data), MARKETING, SALES
        if (h.includes('MARKERTING') || h.includes('MARKETING')) {
            mapping.salesName = header;
        } else if (!mapping.salesName && (h.includes('SALES') || h.includes('PERSON'))) {
            mapping.salesName = header;
        }
        
        // RFQ Status - KETERANGAN is the main status column
        if (h.includes('KETERANGAN')) {
            mapping.rfqStatus = header;
        } else if (!mapping.rfqStatus && (h.includes('STATUS') || (h.includes('PROGRESS') && !h.includes('QUANTITY')))) {
            mapping.rfqStatus = header;
        }
        
        // Order Conversion - TOTAL (final order value)
        if (h.includes('TOTAL') && !h.includes('QUANTITY')) {
            mapping.orderConversion = header;
        } else if (!mapping.orderConversion && h.includes('QUANTITY') && !h.includes('SATUAN')) {
            mapping.orderConversion = header;
        }
    });
    
    return mapping;
}

function analyzeData(headers, rows, columnMapping) {
    const analysis = {
        totalRows: rows.length,
        preview: rows.slice(0, 10),
        fieldNames: headers,
        fieldTypes: detectFieldTypes(headers, rows),
        uniqueCounts: {},
        distributions: {}
    };
    
    // Unique counts for customer and sales
    if (columnMapping.customerName) {
        const custIdx = headers.indexOf(columnMapping.customerName);
        const customers = new Set(rows.map(row => (row[custIdx] || '').trim()).filter(v => v));
        analysis.uniqueCounts.customers = customers.size;
        analysis.distributions.rfqPerCustomer = {};
        rows.forEach(row => {
            const cust = (row[custIdx] || '').trim();
            if (cust) {
                analysis.distributions.rfqPerCustomer[cust] = 
                    (analysis.distributions.rfqPerCustomer[cust] || 0) + 1;
            }
        });
    }
    
    if (columnMapping.salesName) {
        const salesIdx = headers.indexOf(columnMapping.salesName);
        const sales = new Set(rows.map(row => (row[salesIdx] || '').trim()).filter(v => v));
        analysis.uniqueCounts.sales = sales.size;
        analysis.distributions.rfqPerSales = {};
        rows.forEach(row => {
            const salesPerson = (row[salesIdx] || '').trim();
            if (salesPerson) {
                analysis.distributions.rfqPerSales[salesPerson] = 
                    (analysis.distributions.rfqPerSales[salesPerson] || 0) + 1;
            }
        });
    }
    
    // RFQ Conversion Rate
    if (columnMapping.rfqStatus) {
        const statusIdx = headers.indexOf(columnMapping.rfqStatus);
        let converted = 0;
        let notConverted = 0;
        let total = 0;
        const statusCounts = {};
        
        rows.forEach(row => {
            const status = (row[statusIdx] || '').trim().toUpperCase();
            if (status) {
                total++;
                statusCounts[status] = (statusCounts[status] || 0) + 1;
                
                // Check for converted status first (JADI OC)
                if (status.includes('JADI OC') && !status.includes('TIDAK')) {
                    converted++;
                } 
                // Check for not converted (TIDAK JADI OC)
                else if (status.includes('TIDAK JADI') || status.includes('NOT CONVERTED')) {
                    notConverted++;
                }
                // If it's just "OC" or "CONVERTED" without "TIDAK", count as converted
                else if ((status === 'OC' || status.includes('CONVERTED')) && !status.includes('TIDAK')) {
                    converted++;
                }
            }
        });
        
        analysis.distributions.conversion = {
            total: total,
            converted: converted,
            notConverted: notConverted,
            conversionRate: total > 0 ? ((converted / total) * 100).toFixed(2) + '%' : '0%',
            statusBreakdown: statusCounts
        };
        
        // Also check by looking at TOTAL column for non-zero values
        if (columnMapping.orderConversion) {
            const totalIdx = headers.indexOf(columnMapping.orderConversion);
            let ordersWithValue = 0;
            rows.forEach(row => {
                const totalVal = (row[totalIdx] || '').trim();
                const numVal = parseFloat(totalVal.replace(/[^\d,\.]/g, '').replace(/,/g, ''));
                if (numVal && numVal > 0) {
                    ordersWithValue++;
                }
            });
            analysis.distributions.ordersWithValue = ordersWithValue;
        }
    }
    
    return analysis;
}

function formatOutput(analysis, columnMapping) {
    console.log('='.repeat(80));
    console.log('CSV DATASET ANALYSIS SUMMARY');
    console.log('='.repeat(80));
    console.log();
    
    console.log('BASIC INFO:');
    console.log(`  Total Rows (excluding header): ${analysis.totalRows}`);
    console.log(`  Total Fields: ${analysis.fieldNames.length}`);
    console.log();
    
    console.log('FIELD NAMES:');
    analysis.fieldNames.forEach((name, idx) => {
        console.log(`  ${(idx + 1).toString().padStart(2)}. ${name.padEnd(25)} (${analysis.fieldTypes[name]})`);
    });
    console.log();
    
    console.log('COLUMN MAPPING IDENTIFICATION:');
    console.log(`  RFQ Value Column:        ${columnMapping.rfqValue || 'NOT FOUND'}`);
    console.log(`  Customer Name Column:    ${columnMapping.customerName || 'NOT FOUND'}`);
    console.log(`  Sales Name Column:       ${columnMapping.salesName || 'NOT FOUND'}`);
    console.log(`  RFQ Status Column:       ${columnMapping.rfqStatus || 'NOT FOUND'}`);
    console.log(`  Order Conversion Column: ${columnMapping.orderConversion || 'NOT FOUND'}`);
    console.log();
    
    console.log('UNIQUE COUNTS:');
    console.log(`  Unique Customers: ${analysis.uniqueCounts.customers || 'N/A'}`);
    console.log(`  Unique Sales:     ${analysis.uniqueCounts.sales || 'N/A'}`);
    console.log();
    
    console.log('PREVIEW (First 10 Rows):');
    console.log('-'.repeat(80));
    analysis.preview.forEach((row, idx) => {
        console.log(`Row ${idx + 1}:`);
        analysis.fieldNames.forEach((header, hIdx) => {
            const value = (row[hIdx] || '').trim();
            const displayValue = value.length > 50 ? value.substring(0, 47) + '...' : value;
            console.log(`  ${header.padEnd(20)}: ${displayValue}`);
        });
        console.log();
    });
    
    console.log('DISTRIBUTION INFO:');
    console.log('-'.repeat(80));
    
    if (analysis.distributions.rfqPerCustomer) {
        console.log('\nRFQ Count per Customer (Top 10):');
        const customerEntries = Object.entries(analysis.distributions.rfqPerCustomer)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 10);
        customerEntries.forEach(([customer, count]) => {
            console.log(`  ${customer.padEnd(40)}: ${count} RFQs`);
        });
    }
    
    if (analysis.distributions.rfqPerSales) {
        console.log('\nRFQ Count per Salesperson:');
        const salesEntries = Object.entries(analysis.distributions.rfqPerSales)
            .sort((a, b) => b[1] - a[1]);
        salesEntries.forEach(([sales, count]) => {
            console.log(`  ${sales.padEnd(40)}: ${count} RFQs`);
        });
    }
    
    if (analysis.distributions.conversion) {
        const conv = analysis.distributions.conversion;
        console.log('\nRFQ Conversion Rate:');
        console.log(`  Total RFQs:        ${conv.total}`);
        console.log(`  Converted:         ${conv.converted}`);
        console.log(`  Not Converted:     ${conv.notConverted}`);
        console.log(`  Conversion Rate:   ${conv.conversionRate}`);
        if (analysis.distributions.ordersWithValue) {
            console.log(`  Orders with Value: ${analysis.distributions.ordersWithValue}`);
        }
        if (conv.statusBreakdown) {
            console.log('\n  Status Breakdown:');
            Object.entries(conv.statusBreakdown)
                .sort((a, b) => b[1] - a[1])
                .slice(0, 5)
                .forEach(([status, count]) => {
                    console.log(`    ${status.padEnd(25)}: ${count}`);
                });
        }
    }
    
    console.log();
    console.log('='.repeat(80));
}

// Main execution
async function main() {
    try {
        console.log('Loading CSV file...');
        const csvText = await loadCSV();
        
        console.log('Parsing CSV...');
        const { headers, rows } = parseCSV(csvText);
        
        console.log('Identifying columns...');
        const columnMapping = identifyColumns(headers);
        
        console.log('Analyzing data...');
        const analysis = analyzeData(headers, rows, columnMapping);
        
        formatOutput(analysis, columnMapping);
        
        // Also return for potential use
        return { headers, rows, columnMapping, analysis };
        
    } catch (error) {
        console.error('Error:', error);
    }
}

// Run if in Node.js environment
if (typeof window === 'undefined') {
    const fs = require('fs');
    // Override fetch for Node.js
    global.fetch = async (url) => {
        const content = fs.readFileSync(url.replace('./', ''), 'utf-8');
        return {
            text: async () => content
        };
    };
    main();
} else {
    // Browser environment - expose main function
    window.analyzeCSV = main;
}

