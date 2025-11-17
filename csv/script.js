// Get DOM elements
const excelInput = document.getElementById('excel-input');
const processBtn = document.getElementById('process-btn');
const clearBtn = document.getElementById('clear-btn');
const previewSection = document.getElementById('preview-section');
const previewTable = document.getElementById('preview-table');
const downloadBtn = document.getElementById('download-btn');
const copyCsvBtn = document.getElementById('copy-csv-btn');
const rowCount = document.getElementById('row-count');
const colCount = document.getElementById('col-count');
const pasteStatus = document.getElementById('paste-status');
const componentInput = document.getElementById('component-input');
const themeToggle = document.getElementById('theme-toggle');
const themeIcon = themeToggle ? themeToggle.querySelector('.theme-icon') : null;

let csvData = [];
let csvFiles = []; // Array to store multiple files
let componentValue = '';

// Maximum rows per file (Jira limitation)
const MAX_ROWS_PER_FILE = 248;

// Theme management
function initTheme() {
    const savedTheme = localStorage.getItem('theme') || 'light';
    document.documentElement.setAttribute('data-theme', savedTheme);
    updateThemeIcon(savedTheme);
}

function toggleTheme() {
    const currentTheme = document.documentElement.getAttribute('data-theme');
    const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', newTheme);
    localStorage.setItem('theme', newTheme);
    updateThemeIcon(newTheme);
}

function updateThemeIcon(theme) {
    if (themeIcon) {
        themeIcon.textContent = theme === 'dark' ? '☀️' : '🌙';
    }
}

// Initialize theme on load
if (themeToggle) {
    initTheme();
    themeToggle.addEventListener('click', toggleTheme);
}

// Sub-task names (as in context.csv)
const SUB_TASK_NAMES = [
    'Introduction',
    'Learn the Fundamentals of the AWS Cloud',
    'Serverless Fundamentals',
    'IAM Fundamentals',
    'Serverless Design',
    'Advanced Serverless Concepts',
    'Final steps'
];

const TASK_SUMMARY = 'Deep Dive into Serverless';

// Function to show status
function showStatus(message, type = 'success') {
    pasteStatus.textContent = message;
    pasteStatus.className = `paste-status ${type}`;
    pasteStatus.style.display = 'block';
    
    if (type === 'success') {
        setTimeout(() => {
            pasteStatus.style.display = 'none';
        }, 3000);
    }
}

// Function to process Excel data
function processExcelData(input) {
    const lines = input.trim().split('\n');
    const data = [];
    
    for (let line of lines) {
        // Split row into cells
        // Excel usually uses tab to separate cells
        const cells = line.split('\t');
        
        // If no tab, try splitting by multiple spaces
        if (cells.length === 1) {
            const spaceSplit = line.split(/\s{2,}/);
            if (spaceSplit.length > 1) {
                data.push(spaceSplit.map(cell => cell.trim()));
            } else {
                // If single column, just add the whole row
                data.push([line.trim()]);
            }
        } else {
            data.push(cells.map(cell => cell.trim()));
        }
    }
    
    return data;
}

// Function to extract unique emails
function extractUniqueEmails(data) {
    if (data.length === 0) return [];
    
    const emails = new Set();
    
    // Find email column index
    const headers = data[0] ? data[0].map(h => String(h).toLowerCase().trim()) : [];
    let emailIndex = -1;
    
    // Search for email column
    headers.forEach((header, index) => {
        if (header.includes('assignee') || header.includes('email')) {
            emailIndex = index;
        }
    });
    
    // If not found by name, search by content (@)
    if (emailIndex === -1) {
        for (let i = 0; i < (data[0]?.length || 0); i++) {
            if (data.length > 1 && data[1] && data[1][i] && String(data[1][i]).includes('@')) {
                emailIndex = i;
                break;
            }
        }
    }
    
    // If still not found, use first column
    if (emailIndex === -1) emailIndex = 0;
    
    // Extract unique emails
    const startRow = headers.length > 0 && headers.some(h => h.includes('assignee') || h.includes('email')) ? 1 : 0;
    
    for (let i = startRow; i < data.length; i++) {
        const email = data[i] && data[i][emailIndex] ? String(data[i][emailIndex]).trim() : '';
        if (email && email.includes('@')) {
            emails.add(email);
        }
    }
    
    return Array.from(emails).sort();
}

// Function to generate Task + Sub-tasks structure
function generateTaskStructure(emails, component) {
    if (!emails || emails.length === 0 || !component) return [];
    
    const result = [];
    let issueId = 1;
    
    // Add headers
    result.push(['Assignee', 'Issue Type', 'Issue ID', 'Parent ID', 'Summary', 'Component']);
    
    // Generate structure for each email
    emails.forEach(email => {
        const taskId = issueId++;
        
        // Add Task
        result.push([
            email,
            'Task',
            taskId.toString(),
            '', // Parent ID empty for Task
            TASK_SUMMARY,
            component
        ]);
        
        // Add Sub-tasks
        SUB_TASK_NAMES.forEach(subTaskName => {
            result.push([
                email,
                'Sub-task',
                issueId.toString(),
                taskId.toString(), // Parent ID = Task ID
                subTaskName,
                component
            ]);
            issueId++;
        });
    });
    
    return result;
}

// Function to convert data to CSV format
function convertToCSV(data) {
    if (data.length === 0) return '';
    
    // Find maximum number of columns
    const maxCols = Math.max(...data.map(row => row.length));
    
    // Normalize data (add empty cells if needed)
    const normalizedData = data.map(row => {
        const normalizedRow = [...row];
        while (normalizedRow.length < maxCols) {
            normalizedRow.push('');
        }
        return normalizedRow;
    });
    
    // Convert to CSV format
    return normalizedData.map(row => {
        return row.map(cell => {
            // Escape cells if they contain commas, quotes or line breaks
            const cellStr = String(cell || '');
            if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
                return '"' + cellStr.replace(/"/g, '""') + '"';
            }
            return cellStr;
        }).join(',');
    }).join('\n');
}

// Function to split data into parts (248 rows each)
function splitDataIntoFiles(data) {
    if (data.length === 0) return [];
    
    const files = [];
    const header = data[0]; // Header
    const rows = data.slice(1); // Data without header
    
    // If rows less than or equal to MAX_ROWS_PER_FILE, return one file
    if (rows.length <= MAX_ROWS_PER_FILE) {
        return [data];
    }
    
    // Split into parts
    for (let i = 0; i < rows.length; i += MAX_ROWS_PER_FILE) {
        const chunk = rows.slice(i, i + MAX_ROWS_PER_FILE);
        files.push([header, ...chunk]);
    }
    
    return files;
}

// Function to display preview
function displayPreview(data) {
    if (data.length === 0) {
        previewSection.style.display = 'none';
        showStatus('No data found', 'error');
        return;
    }
    
    previewSection.style.display = 'block';
    
    // Find maximum number of columns
    const maxCols = Math.max(...data.map(row => row.length));
    
    // Normalize data
    const normalizedData = data.map(row => {
        const normalizedRow = [...row];
        while (normalizedRow.length < maxCols) {
            normalizedRow.push('');
        }
        return normalizedRow;
    });
    
    // Clear table
    previewTable.innerHTML = '';
    
    // Create headers (if data exists)
    let useFirstRowAsHeaders = false;
    let displayData = [...normalizedData];
    
    if (normalizedData.length > 0) {
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        
        // Use first row as headers if it looks like headers
        const firstRow = normalizedData[0];
        useFirstRowAsHeaders = firstRow.some(cell => cell && cell.toString().trim().length > 0) && normalizedData.length > 1;
        
        if (useFirstRowAsHeaders) {
            // Use first row as headers
            firstRow.forEach(cell => {
                const th = document.createElement('th');
                th.textContent = cell || '';
                headerRow.appendChild(th);
            });
            
            // Remove first row from display data
            displayData = normalizedData.slice(1);
        } else {
            // Create standard headers
            for (let i = 0; i < maxCols; i++) {
                const th = document.createElement('th');
                th.textContent = `Column ${i + 1}`;
                headerRow.appendChild(th);
            }
        }
        
        thead.appendChild(headerRow);
        previewTable.appendChild(thead);
        
        // Create table body
        const tbody = document.createElement('tbody');
        
        displayData.forEach((row, rowIndex) => {
            const tr = document.createElement('tr');
            
            row.forEach((cell, cellIndex) => {
                const td = document.createElement('td');
                td.textContent = cell || '';
                // Add alternating color for better readability
                if (rowIndex % 2 === 0) {
                    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
                    tr.style.backgroundColor = isDark ? '#2a2a3e' : '#fafafa';
                }
                tr.appendChild(td);
            });
            
            tbody.appendChild(tr);
        });
        
        previewTable.appendChild(tbody);
    }
    
    // Update information
    const totalRows = normalizedData.length;
    const uniqueStudents = new Set();
    let taskCount = 0;
    let subtaskCount = 0;
    
    // Count unique students and tasks (skip header)
    for (let i = 1; i < normalizedData.length; i++) {
        if (normalizedData[i][0]) {
            uniqueStudents.add(normalizedData[i][0]);
        }
        if (normalizedData[i][1] === 'Task') {
            taskCount++;
        } else if (normalizedData[i][1] === 'Sub-task') {
            subtaskCount++;
        }
    }
    
    // Split into files if needed
    csvFiles = splitDataIntoFiles(normalizedData);
    const fileCount = csvFiles.length;
    
    rowCount.textContent = `📊 Total rows: ${totalRows - 1}`;
    if (fileCount > 1) {
        colCount.textContent = `📁 Files: ${fileCount} | 👥 Students: ${uniqueStudents.size} | Tasks: ${taskCount} | Sub-tasks: ${subtaskCount}`;
    } else {
        colCount.textContent = `👥 Students: ${uniqueStudents.size} | Tasks: ${taskCount} | Sub-tasks: ${subtaskCount}`;
    }
    
    // Save data
    csvData = normalizedData;
    
    // Show success status
    if (fileCount > 1) {
        showStatus(`✅ Generated ${taskCount} Tasks and ${subtaskCount} Sub-tasks for ${uniqueStudents.size} students! Data split into ${fileCount} files (max ${MAX_ROWS_PER_FILE} rows per file)`, 'success');
    } else {
        showStatus(`✅ Generated ${taskCount} Tasks and ${subtaskCount} Sub-tasks for ${uniqueStudents.size} students!`, 'success');
    }
}

// Process button click handler
processBtn.addEventListener('click', () => {
    // Save field values
    componentValue = componentInput.value.trim();
    
    const input = excelInput.value.trim();
    
    if (!input) {
        showStatus('Please paste data from Excel', 'error');
        return;
    }
    
    if (!componentValue) {
        showStatus('Please enter Component', 'error');
        componentInput.focus();
        return;
    }
    
    const rawData = processExcelData(input);
    // Extract unique emails
    const uniqueEmails = extractUniqueEmails(rawData);
    
    if (uniqueEmails.length === 0) {
        showStatus('No emails found in pasted data', 'error');
        return;
    }
    
    // Generate Task + Sub-tasks structure
    const generatedData = generateTaskStructure(uniqueEmails, componentValue);
    displayPreview(generatedData);
});

// Clear button click handler
clearBtn.addEventListener('click', () => {
    excelInput.value = '';
    previewSection.style.display = 'none';
    csvData = [];
    csvFiles = [];
    componentInput.value = '';
    componentValue = '';
    pasteStatus.style.display = 'none';
    pasteStatus.textContent = '';
    pasteStatus.className = 'paste-status';
});

// Download CSV button click handler
downloadBtn.addEventListener('click', () => {
    if (!csvFiles || csvFiles.length === 0) {
        showStatus('No data to download. Please process data first.', 'error');
        return;
    }
    
    const timestamp = new Date().getTime();
    
    // Download all files
    csvFiles.forEach((fileData, index) => {
        const csvContent = convertToCSV(fileData);
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        
        // Form file name
        let fileName = `data_${timestamp}`;
        if (csvFiles.length > 1) {
            fileName += `_part${index + 1}`;
        }
        fileName += '.csv';
        
        link.setAttribute('href', url);
        link.setAttribute('download', fileName);
        link.style.visibility = 'hidden';
        
        document.body.appendChild(link);
        
        // Delay for sequential download
        setTimeout(() => {
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
        }, index * 200);
    });
    
    if (csvFiles.length > 1) {
        showStatus(`✅ Downloaded ${csvFiles.length} CSV files!`, 'success');
    } else {
        showStatus('✅ CSV file successfully downloaded!', 'success');
    }
});

// Copy CSV button click handler
copyCsvBtn.addEventListener('click', async () => {
    if (!csvFiles || csvFiles.length === 0) {
        showStatus('No data to copy. Please process data first.', 'error');
        return;
    }
    
    // If multiple files, copy only the first one
    const csvContent = convertToCSV(csvFiles[0]);
    
    try {
        await navigator.clipboard.writeText(csvContent);
        copyCsvBtn.textContent = '✅ Copied!';
        if (csvFiles.length > 1) {
            showStatus(`✅ Copied first file (of ${csvFiles.length}). Use download for other files.`, 'success');
        } else {
            showStatus('✅ CSV data copied to clipboard!', 'success');
        }
        setTimeout(() => {
            copyCsvBtn.textContent = '📋 Copy CSV';
        }, 2000);
    } catch (err) {
        // Fallback for older browsers
        const textArea = document.createElement('textarea');
        textArea.value = csvContent;
        textArea.style.position = 'fixed';
        textArea.style.left = '-999999px';
        document.body.appendChild(textArea);
        textArea.select();
        try {
            document.execCommand('copy');
            copyCsvBtn.textContent = '✅ Copied!';
            if (csvFiles.length > 1) {
                showStatus(`✅ Copied first file (of ${csvFiles.length}). Use download for other files.`, 'success');
            } else {
                showStatus('✅ CSV data copied to clipboard!', 'success');
            }
            setTimeout(() => {
                copyCsvBtn.textContent = '📋 Copy CSV';
            }, 2000);
        } catch (err) {
            showStatus('Failed to copy. Try downloading the file.', 'error');
        }
        document.body.removeChild(textArea);
    }
});

// Auto-process on paste (Ctrl+V / Cmd+V) - only show email count
excelInput.addEventListener('paste', (e) => {
    setTimeout(() => {
        const input = excelInput.value.trim();
        if (input) {
            const rawData = processExcelData(input);
            const uniqueEmails = extractUniqueEmails(rawData);
            if (uniqueEmails.length > 0) {
                showStatus(`✅ Found ${uniqueEmails.length} unique emails. Enter Component and click "Process Data"`, 'success');
            } else {
                showStatus('No emails found in pasted data', 'error');
            }
        }
    }, 10);
});

// Text input change handler
excelInput.addEventListener('input', (e) => {
    const input = excelInput.value.trim();
    if (input.length === 0) {
        previewSection.style.display = 'none';
        pasteStatus.style.display = 'none';
    }
});

// Enter key support for processing
excelInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        processBtn.click();
    }
});
