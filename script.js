// ================================
// STATE MANAGEMENT
// ================================

let uploadedData = null;
let currentFile = null;

// ================================
// DOM ELEMENTS
// ================================

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const removeFileBtn = document.getElementById('removeFile');
const previewSection = document.getElementById('previewSection');
const previewTable = document.getElementById('previewTable');
const tableHead = document.getElementById('tableHead');
const tableBody = document.getElementById('tableBody');
const runAnalysisBtn = document.getElementById('runAnalysisBtn');
const clearAllBtn = document.getElementById('clearAllBtn');
const resultsSection = document.getElementById('resultsSection');
const statusMessage = document.getElementById('statusMessage');
const exportBtn = document.getElementById('exportBtn');

// ================================
// EVENT LISTENERS
// ================================

// File input change
fileInput.addEventListener('change', handleFileSelect);

// Drag and drop
uploadArea.addEventListener('dragover', handleDragOver);
uploadArea.addEventListener('dragleave', handleDragLeave);
uploadArea.addEventListener('drop', handleDrop);
uploadArea.addEventListener('click', () => fileInput.click());

// Remove file
removeFileBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    clearFile();
});

// Run analysis
runAnalysisBtn.addEventListener('click', runAnalysis);

// Clear all
clearAllBtn.addEventListener('click', clearAllData);

// Export report
exportBtn.addEventListener('click', exportReport);

// ================================
// FILE HANDLING
// ================================

function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadArea.classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadArea.classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadArea.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function processFile(file) {
    // Validate file
    const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        'application/vnd.ms-excel',
                        'text/csv'];
    
    if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls|csv)$/)) {
        showStatus('Please upload a valid Excel or CSV file', 'error');
        return;
    }

    if (file.size > 10 * 1024 * 1024) { // 10MB limit
        showStatus('File size must be less than 10MB', 'error');
        return;
    }

    currentFile = file;
    displayFileInfo(file);
    readFile(file);
}

function displayFileInfo(file) {
    fileName.textContent = file.name;
    fileSize.textContent = `Size: ${(file.size / 1024).toFixed(2)} KB`;
    fileInfo.style.display = 'block';
    runAnalysisBtn.disabled = false;
    showStatus(`File "${file.name}" uploaded successfully`, 'success');
}

function clearFile() {
    currentFile = null;
    uploadedData = null;
    fileInput.value = '';
    fileInfo.style.display = 'none';
    previewSection.style.display = 'none';
    runAnalysisBtn.disabled = true;
    showStatus('File removed', 'success');
}

function readFile(file) {
    const reader = new FileReader();
    
    reader.onload = (e) => {
        try {
            let data = [];
            
            if (file.name.endsWith('.csv')) {
                data = parseCSV(e.target.result);
            } else {
                const workbook = XLSX.read(e.target.result, { type: 'binary' });
                const sheetName = workbook.SheetNames[0];
                data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
            }
            
            if (data.length === 0) {
                showStatus('File is empty or could not be parsed', 'error');
                clearFile();
                return;
            }
            
            uploadedData = data;
            displayPreview(data);
            showStatus(`Successfully loaded ${data.length} rows from file`, 'success');
            
        } catch (error) {
            console.error('Error reading file:', error);
            showStatus('Error reading file. Please ensure it is a valid Excel or CSV file', 'error');
            clearFile();
        }
    };
    
    reader.onerror = () => {
        showStatus('Error reading file', 'error');
        clearFile();
    };
    
    if (file.name.endsWith('.csv')) {
        reader.readAsText(file);
    } else {
        reader.readAsBinaryString(file);
    }
}

function parseCSV(text) {
    const lines = text.trim().split('\n');
    if (lines.length === 0) return [];
    
    const headers = lines[0].split(',').map(h => h.trim());
    const data = [];
    
    for (let i = 1; i < Math.min(lines.length, 1000); i++) {
        const obj = {};
        const values = lines[i].split(',');
        headers.forEach((header, index) => {
            obj[header] = values[index] ? values[index].trim() : '';
        });
        data.push(obj);
    }
    
    return data;
}

function displayPreview(data) {
    if (!data || data.length === 0) return;
    
    // Get headers
    const headers = Object.keys(data[0]);
    
    // Clear existing
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';
    
    // Add headers
    const headerRow = document.createElement('tr');
    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    tableHead.appendChild(headerRow);
    
    // Add first 10 rows
    data.slice(0, 10).forEach(row => {
        const tr = document.createElement('tr');
        headers.forEach(header => {
            const td = document.createElement('td');
            td.textContent = row[header] || '';
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });
    
    previewSection.style.display = 'block';
}

// ================================
// ANALYSIS ENGINE
// ================================

function runAnalysis() {
    if (!uploadedData || uploadedData.length === 0) {
        showStatus('Please upload a file first', 'error');
        return;
    }
    
    // Get configuration
    const analysisType = document.getElementById('analysisType').value;
    const timeframe = document.getElementById('timeframe').value;
    const threshold = parseInt(document.getElementById('threshold').value);
    const recommendations = document.getElementById('recommendations').value === 'yes';
    
    // Show loading
    showStatus('Analyzing your data... This may take a moment', 'loading');
    runAnalysisBtn.disabled = true;
    
    // Simulate processing time
    setTimeout(() => {
        try {
            const results = performAnalysis(uploadedData, analysisType, threshold);
            displayResults(results, recommendations, threshold);
            showStatus('Analysis completed successfully!', 'success');
        } catch (error) {
            console.error('Analysis error:', error);
            showStatus('Error during analysis. Please check your data format', 'error');
        } finally {
            runAnalysisBtn.disabled = false;
        }
    }, 1500);
}

function performAnalysis(data, analysisType, threshold) {
    const results = {
        totalSkus: data.length,
        highPerformers: 0,
        midPerformers: 0,
        lowPerformers: 0,
        recommendations: []
    };
    
    // Analyze data
    data.forEach((sku, index) => {
        let score = calculateSKUScore(sku, analysisType);
        
        if (score >= 70) {
            results.highPerformers++;
        } else if (score >= threshold) {
            results.midPerformers++;
        } else {
            results.lowPerformers++;
        }
    });
    
    // Generate recommendations
    if (analysisType === 'profitability' || analysisType === 'combined') {
        results.recommendations.push(
            'Focus marketing efforts on high-performing SKUs to maximize profit margins',
            'Review pricing strategy for mid-performing SKUs',
            'Consider discontinuing or repositioning low-performing SKUs'
        );
    }
    
    if (analysisType === 'demand' || analysisType === 'combined') {
        results.recommendations.push(
            'Increase inventory levels for high-demand SKUs',
            'Adjust demand forecasting parameters for volatile SKUs',
            'Implement bundling strategies for complementary products'
        );
    }
    
    if (analysisType === 'inventory' || analysisType === 'combined') {
        results.recommendations.push(
            'Optimize reorder points based on turnover rates',
            'Implement just-in-time inventory for slow-moving SKUs',
            'Consolidate warehouse space usage'
        );
    }
    
    if (analysisType === 'abc' || analysisType === 'combined') {
        results.recommendations.push(
            'Apply stricter inventory controls to A-class SKUs',
            'Implement periodic reviews for B-class SKUs',
            'Simplify management of C-class SKUs'
        );
    }
    
    // Add insights based on data
    if (results.lowPerformers > results.highPerformers * 2) {
        results.recommendations.push(
            'High ratio of low performers detected - consider product portfolio review'
        );
    }
    
    // Add general recommendations
    results.recommendations.push(
        'Implement real-time monitoring for inventory levels',
        'Schedule quarterly reviews of SKU performance metrics',
        'Establish KPIs for each SKU category'
    );
    
    return results;
}

function calculateSKUScore(sku, analysisType) {
    let score = Math.random() * 100; // Base score
    
    // Try to extract numeric values from the row
    const values = Object.values(sku)
        .map(v => parseFloat(v))
        .filter(v => !isNaN(v) && v > 0);
    
    if (values.length > 0) {
        const avgValue = values.reduce((a, b) => a + b, 0) / values.length;
        score = Math.min(100, (avgValue / 100) * 100);
    }
    
    // Apply analysis type adjustments
    switch(analysisType) {
        case 'profitability':
            score += Math.random() * 20 - 10;
            break;
        case 'demand':
            score += Math.random() * 15 - 7.5;
            break;
        case 'inventory':
            score += Math.random() * 25 - 12.5;
            break;
        case 'abc':
            score += Math.random() * 20 - 10;
            break;
        default:
            score += Math.random() * 20 - 10;
    }
    
    return Math.min(100, Math.max(0, score));
}

function displayResults(results, showRecommendations, threshold) {
    // Update metrics
    document.getElementById('totalSkus').textContent = results.totalSkus;
    document.getElementById('highPerformers').textContent = results.highPerformers;
    document.getElementById('midPerformers').textContent = results.midPerformers;
    document.getElementById('lowPerformers').textContent = results.lowPerformers;
    
    // Display recommendations
    const recommendationsList = document.getElementById('recommendationsList');
    recommendationsList.innerHTML = '';
    
    if (showRecommendations && results.recommendations.length > 0) {
        results.recommendations.slice(0, 8).forEach(rec => {
            const li = document.createElement('li');
            li.textContent = rec;
            recommendationsList.appendChild(li);
        });
    }
    
    resultsSection.style.display = 'block';
    
    // Scroll to results
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ================================
// EXPORT FUNCTIONALITY
// ================================

function exportReport() {
    if (!uploadedData) {
        showStatus('No analysis results to export', 'error');
        return;
    }
    
    const analysisType = document.getElementById('analysisType').value;
    const results = {
        timestamp: new Date().toLocaleString(),
        analysisType: analysisType,
        totalSkus: document.getElementById('totalSkus').textContent,
        highPerformers: document.getElementById('highPerformers').textContent,
        midPerformers: document.getElementById('midPerformers').textContent,
        lowPerformers: document.getElementById('lowPerformers').textContent
    };
    
    // Create CSV content
    let csvContent = 'SKU Optimization Analysis Report\n';
    csvContent += `Generated: ${results.timestamp}\n`;
    csvContent += `Analysis Type: ${results.analysisType}\n\n`;
    csvContent += 'Summary Metrics\n';
    csvContent += `Total SKUs Analyzed,${results.totalSkus}\n`;
    csvContent += `High Performers,${results.highPerformers}\n`;
    csvContent += `Mid Performers,${results.midPerformers}\n`;
    csvContent += `Low Performers,${results.lowPerformers}\n\n`;
    
    // Add recommendations
    const recList = document.getElementById('recommendationsList');
    if (recList.children.length > 0) {
        csvContent += 'Key Recommendations\n';
        Array.from(recList.children).forEach(li => {
            csvContent += `"${li.textContent}"\n`;
        });
    }
    
    // Download
    downloadFile(csvContent, 'sku-optimization-report.csv', 'text/csv');
    showStatus('Report exported successfully', 'success');
}

function downloadFile(content, filename, type) {
    const element = document.createElement('a');
    element.setAttribute('href', `data:${type};charset=utf-8,${encodeURIComponent(content)}`);
    element.setAttribute('download', filename);
    element.style.display = 'none';
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
}

// ================================
// UI UTILITIES
// ================================

function showStatus(message, type = 'info') {
    statusMessage.textContent = message;
    statusMessage.className = `status-message ${type}`;
    statusMessage.style.display = 'flex';
    
    if (type !== 'loading') {
        setTimeout(() => {
            statusMessage.style.display = 'none';
        }, 4000);
    }
}

function clearAllData() {
    if (confirm('Are you sure you want to clear all data and start over?')) {
        clearFile();
        uploadedData = null;
        previewSection.style.display = 'none';
        resultsSection.style.display = 'none';
        document.getElementById('analysisType').value = 'profitability';
        document.getElementById('timeframe').value = 'monthly';
        document.getElementById('threshold').value = '20';
        document.getElementById('recommendations').value = 'yes';
        showStatus('All data cleared', 'success');
    }
}

// ================================
// INITIALIZATION
// ================================

document.addEventListener('DOMContentLoaded', () => {
    console.log('SKU Optimization application loaded successfully');
    showStatus('Ready to analyze! Upload an Excel or CSV file to get started', 'success');
});
