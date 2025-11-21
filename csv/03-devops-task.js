// Get DOM elements - wait for DOM to load
let pasteArea, processBtn, clearBtn, preview, downloadBtn, copyCsvBtn, dataPreview, alerts;
let emailPasteArea, emailPreview;
let componentInput, themeToggle, themeIcon;
let rowCount, colCount;

function initDOM() {
    pasteArea = document.getElementById('pasteArea');
    emailPasteArea = document.getElementById('emailPasteArea');
    processBtn = document.getElementById('process-btn');
    clearBtn = document.getElementById('clear-btn');
    preview = document.getElementById('preview');
    downloadBtn = document.getElementById('download-btn');
    copyCsvBtn = document.getElementById('copy-csv-btn');
    dataPreview = document.getElementById('dataPreview');
    emailPreview = document.getElementById('emailPreview');
    alerts = document.getElementById('alerts');
    componentInput = document.getElementById('component-input');
    themeToggle = document.getElementById('theme-toggle');
    themeIcon = themeToggle ? themeToggle.querySelector('.theme-icon') : null;
    rowCount = document.getElementById('row-count');
    colCount = document.getElementById('col-count');
    
    if (!pasteArea || !processBtn || !clearBtn || !preview || !alerts) {
        console.error('Required DOM elements not found');
        return false;
    }
    return true;
}

let csvData = [];
let csvFiles = [];
let parsedTemplateData = [];
let parsedEmails = [];
let isParsing = false;
let isParsingEmails = false;

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

// Initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        if (initDOM()) {
            setupEventListeners();
        }
    });
} else {
    if (initDOM()) {
        setupEventListeners();
    }
}

function setupEventListeners() {
    if (themeToggle) {
        initTheme();
        themeToggle.addEventListener('click', toggleTheme);
    }
    
    // Prevent form submission (we handle it manually via Process button)
    const devopsForm = document.getElementById('devopsForm');
    if (devopsForm) {
        devopsForm.addEventListener('submit', (e) => {
            e.preventDefault();
            // Trigger Process button click instead
            if (processBtn) {
                processBtn.click();
            }
        });
    }
    
    // Paste area handling
    if (pasteArea) {
        pasteArea.addEventListener('click', () => {
            pasteArea.focus();
        });

        pasteArea.addEventListener('paste', (e) => {
            e.preventDefault();
            const pastedText = (e.clipboardData || window.clipboardData).getData('text');
            
            console.log('Paste event - pasted text length:', pastedText.length);
            console.log('Paste event - first 200 chars:', pastedText.substring(0, 200));
            
            // Check if pasted text is placeholder or status message
            const placeholderTexts = ['Data pasted', 'processing', 'Click here', 'Expected columns', 'Successfully loaded'];
            const isPlaceholder = placeholderTexts.some(placeholder => pastedText.toLowerCase().includes(placeholder.toLowerCase()));
            
            if (isPlaceholder || pastedText.trim().length < 50) {
                console.log('Paste event: Ignoring placeholder or too short text');
                if (alerts) {
                    showAlert('⚠️ Please paste actual Excel data, not status messages', 'error');
                }
                pasteArea.innerHTML = `
                    <p><strong>Click here and paste the table</strong></p>
                    <p style="font-size: 12px; color: var(--text-muted);">Expected columns: Epic, Task, Description</p>
                `;
                return;
            }
            
            // Prevent double parsing
            if (isParsing) {
                console.log('Paste event: Already parsing, skipping');
                return;
            }
            
            // Set flag and process
            isParsing = true;
            pasteArea.innerHTML = '<p style="color: #10b981;"><strong>✅ Data pasted, processing...</strong></p>';
            
            // Call parseExcelData
            parseExcelData(pastedText);
        });

        pasteArea.addEventListener('keydown', (e) => {
            if ((e.ctrlKey || e.metaKey) && e.key === 'v') {
                // Paste event will handle it, so we don't need to do anything here
                // This is just a fallback in case paste event doesn't fire
                if (!isParsing) {
                    setTimeout(() => {
                        // Only parse if paste event didn't handle it and we're not already parsing
                        if (!isParsing) {
                            const text = pasteArea.textContent || pasteArea.innerText;
                            // Filter out placeholder text and success messages
                            const placeholderTexts = ['Click here and paste the table', 'Expected columns:', 'Click here and paste', 'Data pasted', 'Successfully loaded', 'processing'];
                            const isPlaceholder = placeholderTexts.some(placeholder => text.toLowerCase().includes(placeholder.toLowerCase()));
                            if (text && text.trim().length > 50 && !isPlaceholder) {
                                console.log('keydown handler: Processing paste via keydown fallback');
                                isParsing = true;
                                pasteArea.innerHTML = '<p style="color: #10b981;"><strong>✅ Data pasted, processing...</strong></p>';
                                parseExcelData(text);
                            }
                        }
                    }, 100);
                }
            }
        });
    }
    
    // Email paste area handling
    if (emailPasteArea) {
        emailPasteArea.addEventListener('click', () => {
            emailPasteArea.focus();
        });

        emailPasteArea.addEventListener('paste', (e) => {
            e.preventDefault();
            const pastedText = (e.clipboardData || window.clipboardData).getData('text');
            
            console.log('Email paste event - pasted text length:', pastedText.length);
            
            // Check if pasted text is placeholder or status message
            const placeholderTexts = ['Data pasted', 'processing', 'Click here', 'One email per row'];
            const isPlaceholder = placeholderTexts.some(placeholder => pastedText.toLowerCase().includes(placeholder.toLowerCase()));
            
            if (isPlaceholder || pastedText.trim().length === 0) {
                console.log('Email paste event: Ignoring placeholder or empty text');
                if (alerts) {
                    showAlert('⚠️ Please paste actual email data', 'error');
                }
                emailPasteArea.innerHTML = `
                    <p><strong>Click here and paste emails</strong></p>
                    <p style="font-size: 12px; color: var(--text-muted);">One email per row (or column)</p>
                `;
                return;
            }
            
            // Prevent double parsing
            if (isParsingEmails) {
                console.log('Email paste event: Already parsing, skipping');
                return;
            }
            
            // Set flag and process
            isParsingEmails = true;
            emailPasteArea.innerHTML = '<p style="color: #10b981;"><strong>✅ Emails pasted, processing...</strong></p>';
            
            // Call parseEmails
            parseEmails(pastedText);
        });

        emailPasteArea.addEventListener('keydown', (e) => {
            if ((e.ctrlKey || e.metaKey) && e.key === 'v') {
                if (!isParsingEmails) {
                    setTimeout(() => {
                        if (!isParsingEmails) {
                            const text = emailPasteArea.textContent || emailPasteArea.innerText;
                            const placeholderTexts = ['Click here and paste emails', 'One email per row', 'Data pasted', 'Successfully loaded', 'processing'];
                            const isPlaceholder = placeholderTexts.some(placeholder => text.toLowerCase().includes(placeholder.toLowerCase()));
                            if (text && text.trim().length > 0 && !isPlaceholder) {
                                console.log('Email keydown handler: Processing paste via keydown fallback');
                                isParsingEmails = true;
                                emailPasteArea.innerHTML = '<p style="color: #10b981;"><strong>✅ Emails pasted, processing...</strong></p>';
                                parseEmails(text);
                            }
                        }
                    }, 100);
                }
            }
        });
    }
    
    // Process button
    if (processBtn) {
        processBtn.addEventListener('click', () => {
            console.log('Process button clicked');
            
            // Validate inputs
            const component = componentInput ? componentInput.value.trim() : '';
            
            if (!component) {
                showAlert('⚠️ Please enter Component', 'error');
                return;
            }
            
            if (!parsedTemplateData || parsedTemplateData.length === 0) {
                showAlert('⚠️ Please paste tasks data from Excel first', 'error');
                return;
            }
            
            if (!parsedEmails || parsedEmails.length === 0) {
                showAlert('⚠️ Please paste emails from Excel', 'error');
                return;
            }
            
            // Generate CSV structure
            const generatedData = generateDevOpsStructure(parsedTemplateData, parsedEmails, component);
            
            if (!generatedData || generatedData.length === 0) {
                showAlert('⚠️ Failed to generate CSV data', 'error');
                return;
            }
            
            csvData = generatedData;
            csvFiles = splitDataIntoFiles(csvData);
            
            // Display preview
            displayPreview(csvData);
            
            // Show download buttons
            if (downloadBtn) downloadBtn.style.display = 'inline-block';
            if (copyCsvBtn) copyCsvBtn.style.display = 'inline-block';
            
            showAlert('✅ CSV generated successfully!', 'success');
        });
    }
    
    // Clear button
    if (clearBtn) {
        clearBtn.addEventListener('click', () => {
            console.log('Clear button clicked');
            
            // Clear inputs
            if (componentInput) componentInput.value = '';
            
            // Clear paste area for tasks
            if (pasteArea) {
                pasteArea.innerHTML = `
                    <p><strong>Click here and paste the table</strong></p>
                    <p style="font-size: 12px; color: var(--text-muted);">Expected columns: Epic, Task, Description</p>
                `;
                pasteArea.classList.remove('has-data');
            }
            
            // Clear paste area for emails
            if (emailPasteArea) {
                emailPasteArea.innerHTML = `
                    <p><strong>Click here and paste emails</strong></p>
                    <p style="font-size: 12px; color: var(--text-muted);">One email per row (or column)</p>
                `;
                emailPasteArea.classList.remove('has-data');
            }
            
            // Clear previews
            if (dataPreview) {
                dataPreview.innerHTML = '';
                dataPreview.classList.add('hidden');
            }
            
            if (emailPreview) {
                emailPreview.innerHTML = '';
                emailPreview.classList.add('hidden');
            }
            
            if (preview) {
                preview.innerHTML = '<div class="loading">Paste data from Excel to generate CSV</div>';
            }
            
            // Hide download buttons
            if (downloadBtn) downloadBtn.style.display = 'none';
            if (copyCsvBtn) copyCsvBtn.style.display = 'none';
            
            // Clear data
            csvData = [];
            csvFiles = [];
            parsedTemplateData = [];
            parsedEmails = [];
            
            // Clear alerts
            if (alerts) alerts.innerHTML = '';
            
            // Clear statistics
            if (rowCount) rowCount.textContent = '';
            if (colCount) colCount.textContent = '';
        });
    }
    
    // Download button
    if (downloadBtn) {
        downloadBtn.addEventListener('click', () => {
            if (csvFiles.length === 0) {
                showAlert('⚠️ No data to download', 'error');
                return;
            }
            
            csvFiles.forEach((fileData, index) => {
                const csv = convertToCSV(fileData);
                const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
                const link = document.createElement('a');
                const url = URL.createObjectURL(blob);
                link.setAttribute('href', url);
                link.setAttribute('download', `devops-tasks${csvFiles.length > 1 ? `-part${index + 1}` : ''}.csv`);
                link.style.visibility = 'hidden';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            });
            
            showAlert(`✅ Downloaded ${csvFiles.length} file(s)`, 'success');
        });
    }
    
    // Copy CSV button
    if (copyCsvBtn) {
        copyCsvBtn.addEventListener('click', () => {
            if (csvFiles.length === 0) {
                showAlert('⚠️ No data to copy', 'error');
                return;
            }
            
            const csv = convertToCSV(csvFiles[0]);
            navigator.clipboard.writeText(csv).then(() => {
                showAlert('✅ CSV copied to clipboard!', 'success');
            }).catch(err => {
                console.error('Failed to copy:', err);
                showAlert('⚠️ Failed to copy to clipboard', 'error');
            });
        });
    }
}

// Function to show alert messages
function showAlert(message, type = 'info') {
    if (!alerts) return;
    
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type}`;
    alertDiv.textContent = message;
    
    alerts.innerHTML = '';
    alerts.appendChild(alertDiv);
    
    // Auto-remove after 5 seconds
    setTimeout(() => {
        if (alertDiv.parentNode) {
            alertDiv.parentNode.removeChild(alertDiv);
        }
    }, 5000);
}

// Function to find column index by name (case-insensitive, flexible)
function findColumnIndex(headers, possibleNames) {
    for (let i = 0; i < headers.length; i++) {
        const header = headers[i].trim().toLowerCase();
        for (const name of possibleNames) {
            if (header === name.toLowerCase() || header.includes(name.toLowerCase())) {
                return i;
            }
        }
    }
    return -1;
}

// Function to parse Excel data
function parseExcelData(text) {
    console.log('parseExcelData called with text length:', text.length);
    
    if (!text || text.trim().length === 0) {
        console.log('parseExcelData: Empty text');
        if (alerts) {
            showAlert('⚠️ No data to parse', 'error');
        }
        isParsing = false;
        return;
    }
    
    try {
        // Split into lines
        const lines = text.split(/\r?\n/).filter(line => line.trim().length > 0);
        console.log('parseExcelData: Total lines after split:', lines.length);
        
        if (lines.length < 2) {
            console.log('parseExcelData: Not enough lines (need at least header + 1 data row)');
            if (alerts) {
                showAlert('⚠️ Not enough data. Please paste at least a header row and one data row.', 'error');
            }
            isParsing = false;
            if (pasteArea) {
                pasteArea.innerHTML = `
                    <p><strong>Click here and paste the table</strong></p>
                    <p style="font-size: 12px; color: var(--text-muted);">Expected columns: Epic, Task, Description</p>
                `;
            }
            return;
        }
        
        // Parse header row - try different separators
        let headers = [];
        const firstLine = lines[0];
        let separator = '\t';
        let separatorRegex = /\t/;
        
        // Try tab separator first (most common from Excel)
        if (firstLine.includes('\t')) {
            headers = firstLine.split('\t').map(h => h.trim());
            separator = '\t';
            separatorRegex = /\t/;
        } else if (firstLine.includes(',')) {
            headers = firstLine.split(',').map(h => h.trim().replace(/^"|"$/g, ''));
            separator = ',';
            separatorRegex = /,/;
        } else if (firstLine.includes(';')) {
            headers = firstLine.split(';').map(h => h.trim().replace(/^"|"$/g, ''));
            separator = ';';
            separatorRegex = /;/;
        } else {
            headers = [firstLine.trim()];
        }
        
        const expectedColCount = headers.length;
        console.log('parseExcelData: Headers found:', headers);
        console.log('parseExcelData: Expected column count:', expectedColCount);
        
        // Find column indices - ожидаем колонки: Epic, Task, Description
        const epicIdx = findColumnIndex(headers, ['epic', 'epic name', 'epic name']);
        const taskIdx = findColumnIndex(headers, ['task', 'summary', 'title']);
        const descriptionIdx = findColumnIndex(headers, ['description', 'desc', 'details']);
        
        console.log('parseExcelData: Column indices - epic:', epicIdx, 'task:', taskIdx, 'description:', descriptionIdx);
        
        if (epicIdx === -1 || taskIdx === -1 || descriptionIdx === -1) {
            const missingCols = [];
            if (epicIdx === -1) missingCols.push('Epic');
            if (taskIdx === -1) missingCols.push('Task');
            if (descriptionIdx === -1) missingCols.push('Description');
            
            console.log('parseExcelData: Missing columns:', missingCols);
            if (alerts) {
                showAlert(`⚠️ Missing required columns: ${missingCols.join(', ')}. Expected columns: Epic, Task, Description`, 'error');
            }
            isParsing = false;
            if (pasteArea) {
                pasteArea.innerHTML = `
                    <p><strong>Click here and paste the table</strong></p>
                    <p style="font-size: 12px; color: var(--text-muted);">Expected columns: Epic, Task, Description</p>
                `;
            }
            return;
        }
        
        // Parse data rows - handle multiline cells properly
        const result = [];
        console.log('Starting to parse rows. Total lines:', lines.length);
        
        // Find the actual number of separators in first data row (not from header)
        let actualSeparatorsInDataRow = -1;
        for (let i = 1; i < lines.length; i++) {
            const sepCount = (lines[i].match(separatorRegex) || []).length;
            if (sepCount > 0) {
                actualSeparatorsInDataRow = sepCount;
                console.log(`First data row (line ${i}) has ${sepCount} separators`);
                break;
            }
        }
        
        // If we couldn't find any row with separators, fall back to expected count
        if (actualSeparatorsInDataRow === -1) {
            actualSeparatorsInDataRow = expectedColCount - 1;
        }
        
        console.log('Using separator count for new row detection:', actualSeparatorsInDataRow);
        
        let accumulatedLine = '';
        
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i];
            const separatorCount = (line.match(separatorRegex) || []).length;
            
            // If this is the first line or accumulated line is empty, start accumulating
            if (accumulatedLine === '') {
                accumulatedLine = line;
            } else {
                // Check if current line looks like start of a new row
                // A new row should have the same number of separators as the first data row
                if (separatorCount >= actualSeparatorsInDataRow) {
                    // This is a new row, so process the accumulated line first
                    let cells = accumulatedLine.split(separator);
                    if (separator !== '\t' && cells.length === 1) {
                        cells = accumulatedLine.split(separatorRegex);
                    }
                    
                    // Trim cells but preserve content for Description column
                    cells = cells.map((c, idx) => {
                        // For Description column, only trim leading/trailing spaces on each line, not the whole content
                        if (idx === descriptionIdx) {
                            // Preserve line breaks - just remove quotes
                            return c.replace(/^"|"$/g, '');
                        }
                        // For other columns, trim normally
                        return c.trim().replace(/^"|"$/g, '');
                    });
                    
                    if (cells.length > Math.max(epicIdx, taskIdx, descriptionIdx)) {
                        const rowData = {
                            epic: cells[epicIdx] || '',
                            task: cells[taskIdx] || '',
                            description: cells[descriptionIdx] || ''
                        };
                        result.push(rowData);
                        console.log(`Row ${result.length} parsed: epic="${rowData.epic.substring(0,20)}", task="${rowData.task.substring(0,20)}", desc length=${rowData.description.length}, linebreaks=${(rowData.description.match(/\n/g) || []).length}`);
                    }
                    
                    // Start new row
                    accumulatedLine = line;
                } else {
                    // This line doesn't have enough separators, it's a continuation (multiline cell)
                    // Append to accumulated line with newline
                    accumulatedLine += '\n' + line;
                }
            }
        }
        
        // Don't forget to process the last accumulated row
        if (accumulatedLine !== '') {
            let cells = accumulatedLine.split(separator);
            if (separator !== '\t' && cells.length === 1) {
                cells = accumulatedLine.split(separatorRegex);
            }
            
            // Trim cells but preserve content for Description column
            cells = cells.map((c, idx) => {
                // For Description column, only trim leading/trailing spaces on each line, not the whole content
                if (idx === descriptionIdx) {
                    // Preserve line breaks - just remove quotes
                    return c.replace(/^"|"$/g, '');
                }
                // For other columns, trim normally
                return c.trim().replace(/^"|"$/g, '');
            });
            
            if (cells.length > Math.max(epicIdx, taskIdx, descriptionIdx)) {
                const rowData = {
                    epic: cells[epicIdx] || '',
                    task: cells[taskIdx] || '',
                    description: cells[descriptionIdx] || ''
                };
                result.push(rowData);
                console.log(`Row ${result.length} parsed (last): epic="${rowData.epic.substring(0,20)}", task="${rowData.task.substring(0,20)}", desc length=${rowData.description.length}, linebreaks=${(rowData.description.match(/\n/g) || []).length}`);
            }
        }
        
        console.log('parseExcelData: Parsed', result.length, 'rows');
        
        if (result.length === 0) {
            console.log('parseExcelData: No valid rows parsed');
            if (alerts) {
                showAlert('⚠️ No valid data rows found', 'error');
            }
            isParsing = false;
            if (pasteArea) {
                pasteArea.innerHTML = `
                    <p><strong>Click here and paste the table</strong></p>
                    <p style="font-size: 12px; color: var(--text-muted);">Expected columns: Epic, Task, Description</p>
                `;
            }
            return;
        }
        
        // Store parsed data
        parsedTemplateData = result;
        
        // Show success message
        if (alerts) {
            showAlert(`✅ Successfully loaded ${result.length} row(s)`, 'success');
        }
        
        // Update paste area
        if (pasteArea) {
            pasteArea.classList.add('has-data');
            pasteArea.innerHTML = `<p style="color: #10b981;"><strong>✅ ${result.length} row(s) loaded</strong></p>`;
        }
        
        // Show data preview
        if (dataPreview) {
            let html = '<h3>📋 Parsed Data Preview:</h3>';
            html += '<div class="data-preview"><table>';
            html += '<thead><tr><th>Epic</th><th>Task</th><th>Description</th></tr></thead>';
            html += '<tbody>';
            result.slice(0, 10).forEach(row => {
                html += `
                    <tr>
                        <td>${escapeHtml(row.epic || '')}</td>
                        <td>${escapeHtml(row.task || '')}</td>
                        <td>${escapeHtml(row.description || '')}</td>
                    </tr>
                `;
            });
            if (result.length > 10) {
                html += `<tr><td colspan="2" style="text-align: center; color: var(--text-muted);">... and ${result.length - 10} more rows</td></tr>`;
            }
            html += '</tbody></table></div>';
            dataPreview.innerHTML = html;
            dataPreview.classList.remove('hidden');
        }
        
        isParsing = false;
        
    } catch (error) {
        console.error('parseExcelData error:', error);
        if (alerts) {
            showAlert('⚠️ Error parsing data: ' + error.message, 'error');
        }
        isParsing = false;
        if (pasteArea) {
            pasteArea.innerHTML = `
                <p><strong>Click here and paste the table</strong></p>
                <p style="font-size: 12px; color: var(--text-muted);">Expected columns: Task, Description</p>
            `;
        }
    }
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Function to convert email to Name Surname format
// serhii_serebrennykov@epam.com -> Serhii Serebrennykov
// serhii_serebrennykov1@epam.com -> Serhii Serebrennykov1
function emailToNameSurname(email) {
    if (!email || !email.includes('@')) {
        return email; // Return as is if not a valid email
    }
    
    // Extract part before @
    const localPart = email.split('@')[0];
    
    // Replace underscores with spaces
    const withSpaces = localPart.replace(/_/g, ' ');
    
    // Convert to Title Case (first letter of each word uppercase)
    const nameSurname = withSpaces.split(' ').map(word => {
        if (word.length === 0) return word;
        return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    }).join(' ');
    
    return nameSurname;
}

// Function to parse emails from Excel
function parseEmails(text) {
    console.log('parseEmails called with text length:', text.length);
    
    if (!text || text.trim().length === 0) {
        console.log('parseEmails: Empty text');
        if (alerts) {
            showAlert('⚠️ No email data to parse', 'error');
        }
        isParsingEmails = false;
        return;
    }
    
    try {
        // Split into lines
        const lines = text.split(/\r?\n/).filter(line => line.trim().length > 0);
        console.log('parseEmails: Total lines after split:', lines.length);
        
        if (lines.length === 0) {
            console.log('parseEmails: No lines found');
            if (alerts) {
                showAlert('⚠️ No email data found', 'error');
            }
            isParsingEmails = false;
            if (emailPasteArea) {
                emailPasteArea.innerHTML = `
                    <p><strong>Click here and paste emails</strong></p>
                    <p style="font-size: 12px; color: var(--text-muted);">One email per row (or column)</p>
                `;
            }
            return;
        }
        
        // Extract emails - try different separators
        const emails = [];
        
        for (const line of lines) {
            let cells = [];
            
            // Try different separators
            if (line.includes('\t')) {
                cells = line.split('\t').map(c => c.trim());
            } else if (line.includes(',')) {
                cells = line.split(',').map(c => c.trim().replace(/^"|"$/g, ''));
            } else if (line.includes(';')) {
                cells = line.split(';').map(c => c.trim().replace(/^"|"$/g, ''));
            } else {
                cells = [line.trim()];
            }
            
            // Extract emails from cells (simple email validation)
            cells.forEach(cell => {
                const trimmedCell = cell.trim();
                // Simple email validation (contains @ and has at least one dot after @)
                if (trimmedCell && trimmedCell.includes('@') && trimmedCell.includes('.')) {
                    emails.push(trimmedCell);
                }
            });
        }
        
        console.log('parseEmails: Parsed', emails.length, 'emails');
        
        if (emails.length === 0) {
            console.log('parseEmails: No valid emails found');
            if (alerts) {
                showAlert('⚠️ No valid emails found. Please paste emails (e.g., user@example.com)', 'error');
            }
            isParsingEmails = false;
            if (emailPasteArea) {
                emailPasteArea.innerHTML = `
                    <p><strong>Click here and paste emails</strong></p>
                    <p style="font-size: 12px; color: var(--text-muted);">One email per row (or column)</p>
                `;
            }
            return;
        }
        
        // Store parsed emails
        parsedEmails = emails;
        
        // Show success message
        if (alerts) {
            showAlert(`✅ Successfully loaded ${emails.length} email(s)`, 'success');
        }
        
        // Update paste area
        if (emailPasteArea) {
            emailPasteArea.classList.add('has-data');
            emailPasteArea.innerHTML = `<p style="color: #10b981;"><strong>✅ ${emails.length} email(s) loaded</strong></p>`;
        }
        
        // Show email preview
        if (emailPreview) {
            let html = '<h3>📧 Parsed Emails Preview:</h3>';
            html += '<div class="data-preview"><table>';
            html += '<thead><tr><th>Email</th></tr></thead>';
            html += '<tbody>';
            emails.slice(0, 20).forEach(email => {
                html += `<tr><td>${escapeHtml(email)}</td></tr>`;
            });
            if (emails.length > 20) {
                html += `<tr><td style="text-align: center; color: var(--text-muted);">... and ${emails.length - 20} more emails</td></tr>`;
            }
            html += '</tbody></table></div>';
            emailPreview.innerHTML = html;
            emailPreview.classList.remove('hidden');
        }
        
        isParsingEmails = false;
        
    } catch (error) {
        console.error('parseEmails error:', error);
        if (alerts) {
            showAlert('⚠️ Error parsing emails: ' + error.message, 'error');
        }
        isParsingEmails = false;
        if (emailPasteArea) {
            emailPasteArea.innerHTML = `
                <p><strong>Click here and paste emails</strong></p>
                <p style="font-size: 12px; color: var(--text-muted);">One email per row (or column)</p>
            `;
        }
    }
}

// Function to generate DevOps CSV structure
function generateDevOpsStructure(templateData, emails, component) {
    console.log('generateDevOpsStructure called with:', {
        templateDataLength: templateData ? templateData.length : 0,
        emailsLength: emails ? emails.length : 0,
        component: component
    });
    
    if (!templateData || templateData.length === 0 || !emails || emails.length === 0 || !component) {
        console.log('generateDevOpsStructure: Early return - missing required parameters');
        return [];
    }
    
    const result = [];
    
    // Add headers - adjust based on your Jira CSV import requirements
    result.push(['Assignee', 'Issue Type', 'Epic Name', 'Epic Link', 'Summary', 'Description', 'Component']);
    
    // Generate tasks for each email
    emails.forEach(email => {
        // Convert email to Name Surname format for Epic Name and Epic Link
        const nameSurname = emailToNameSurname(email);
        
        // First, add Epic (one per student) - placed before all tasks
        // Epic Name is generated from email (Name Surname)
        result.push([
            email,       // Assignee (email)
            'Epic',      // Issue Type = Epic
            nameSurname, // Epic Name (Name Surname generated from email - one per student)
            '',          // Epic Link empty for Epic
            nameSurname, // Summary = Epic Name (same as Epic Name)
            '',          // Description empty for Epic
            component    // Component
        ]);
        
        // Then, add all Tasks for this email
        templateData.forEach((row, index) => {
            const task = row.task ? row.task.trim() : '';
            // Don't trim description to preserve line breaks and formatting
            const description = row.description || '';
            
            // Debug: log first task to check description
            if (index === 0) {
                console.log('First task description check:', {
                    description: description,
                    descriptionLength: description ? description.length : 0,
                    hasLineBreaks: description.includes('\n'),
                    lineBreakCount: (description.match(/\n/g) || []).length
                });
            }
            
            // Skip empty tasks
            if (!task) return;
            
            // Create Task with Epic Name empty and Epic Link = Name Surname
            result.push([
                email,       // Assignee (keep email)
                'Task',      // Issue Type
                '',          // Epic Name empty for Tasks
                nameSurname, // Epic Link (Name Surname generated from email)
                task,        // Summary
                description, // Description (with line breaks preserved)
                component    // Component
            ]);
        });
    });
    
    console.log('generateDevOpsStructure: Generated', result.length, 'rows');
    return result;
}

// Function to convert data to CSV format
function convertToCSV(data) {
    if (data.length === 0) return '';
    
    const maxCols = Math.max(...data.map(row => row.length));
    const normalizedData = data.map(row => {
        const normalizedRow = [...row];
        while (normalizedRow.length < maxCols) {
            normalizedRow.push('');
        }
        return normalizedRow;
    });
    
    return normalizedData.map(row => {
        return row.map(cell => {
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
    const header = data[0];
    const rows = data.slice(1);
    
    if (rows.length <= MAX_ROWS_PER_FILE) {
        return [data];
    }
    
    for (let i = 0; i < rows.length; i += MAX_ROWS_PER_FILE) {
        const chunk = rows.slice(i, i + MAX_ROWS_PER_FILE);
        files.push([header, ...chunk]);
    }
    
    return files;
}

// Function to display preview
function displayPreview(data) {
    if (!data || data.length === 0) {
        if (preview) {
            preview.innerHTML = '<div class="loading">No data to display</div>';
        }
        return;
    }
    
    // Calculate statistics
    const totalRows = data.length - 1; // Exclude header
    const totalFiles = csvFiles.length;
    const totalTasks = totalRows;
    
    // Update statistics badges
    if (rowCount) {
        rowCount.textContent = `📊 ${totalRows} rows`;
    }
    if (colCount) {
        colCount.textContent = `📁 ${totalFiles} file${totalFiles > 1 ? 's' : ''}`;
    }
    
    // Generate table HTML
    let html = '<table>';
    
    // Header row
    html += '<thead><tr>';
    if (data[0]) {
        data[0].forEach(header => {
            html += `<th>${escapeHtml(String(header || ''))}</th>`;
        });
    }
    html += '</tr></thead>';
    
    // Data rows (limit to first 100 for performance)
    html += '<tbody>';
    const displayRows = data.slice(1, 101);
    displayRows.forEach(row => {
        html += '<tr>';
        row.forEach(cell => {
            html += `<td>${escapeHtml(String(cell || ''))}</td>`;
        });
        html += '</tr>';
    });
    
    if (data.length > 101) {
        html += `<tr><td colspan="${data[0] ? data[0].length : 1}" style="text-align: center; color: var(--text-muted); padding: 20px;">... and ${data.length - 101} more rows (will be included in CSV)</td></tr>`;
    }
    
    html += '</tbody></table>';
    
    if (preview) {
        preview.innerHTML = html;
    }
}

