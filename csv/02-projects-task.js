// Get DOM elements - wait for DOM to load
let pasteArea, processBtn, clearBtn, preview, downloadBtn, copyCsvBtn, dataPreview, alerts;
let componentInput, labelInput, teamsInput, themeToggle, themeIcon;

function initDOM() {
    pasteArea = document.getElementById('pasteArea');
    processBtn = document.getElementById('process-btn');
    clearBtn = document.getElementById('clear-btn');
    preview = document.getElementById('preview');
    downloadBtn = document.getElementById('download-btn');
    copyCsvBtn = document.getElementById('copy-csv-btn');
    dataPreview = document.getElementById('dataPreview');
    alerts = document.getElementById('alerts');
    componentInput = document.getElementById('component-input');
    labelInput = document.getElementById('label-input');
    teamsInput = document.getElementById('teams-input');
    themeToggle = document.getElementById('theme-toggle');
    themeIcon = themeToggle ? themeToggle.querySelector('.theme-icon') : null;
    
    if (!pasteArea || !processBtn || !clearBtn || !preview || !alerts) {
        console.error('Required DOM elements not found');
        return false;
    }
    return true;
}

let csvData = [];
let csvFiles = [];
let parsedTemplateData = [];
let isParsing = false;

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
                    <p style="font-size: 12px; color: var(--text-muted);">Expected columns: Epic, Story, Sprint, Description</p>
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
                    }, 200);
                }
            }
        });
    }
    
    // Process button click handler
    if (processBtn) {
        processBtn.addEventListener('click', () => {
            const component = componentInput.value.trim();
            const label = labelInput.value.trim();
            const numTeams = parseInt(teamsInput.value);
            
            if (!component) {
                showAlert('Please enter Component', 'error');
                componentInput.focus();
                return;
            }
            
            if (!label) {
                showAlert('Please enter Label', 'error');
                labelInput.focus();
                return;
            }
            
            if (!numTeams || numTeams < 1) {
                showAlert('Please enter valid Number of Teams', 'error');
                teamsInput.focus();
                return;
            }
            
            if (!parsedTemplateData || parsedTemplateData.length === 0) {
                showAlert('Please paste Excel data first', 'error');
                return;
            }
            
            const generatedData = generateProjectsStructure(parsedTemplateData, component, label, numTeams);
            displayPreview(generatedData);
        });
    }
    
    // Clear button click handler
    if (clearBtn) {
        clearBtn.addEventListener('click', () => {
            pasteArea.innerHTML = `
                <p><strong>Click here and paste the table</strong></p>
                <p style="font-size: 12px; color: var(--text-muted);">Expected columns: Epic, Story, Sprint, Description</p>
            `;
            pasteArea.classList.remove('has-data');
            preview.innerHTML = '<div class="loading">Paste data from Excel to generate CSV</div>';
            csvData = [];
            csvFiles = [];
            parsedTemplateData = [];
            componentInput.value = '';
            labelInput.value = '';
            teamsInput.value = '';
            if (dataPreview) dataPreview.classList.add('hidden');
            if (downloadBtn) downloadBtn.style.display = 'none';
            if (copyCsvBtn) copyCsvBtn.style.display = 'none';
            if (alerts) alerts.innerHTML = '';
        });
    }
    
    // Download CSV button click handler
    if (downloadBtn) {
        downloadBtn.addEventListener('click', () => {
            if (!csvFiles || csvFiles.length === 0) {
                showAlert('No data to download. Please process data first.', 'error');
                return;
            }
            
            const timestamp = new Date().getTime();
            
            csvFiles.forEach((fileData, index) => {
                const csvContent = convertToCSV(fileData);
                const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                const link = document.createElement('a');
                const url = URL.createObjectURL(blob);
                
                let fileName = `projects_task_${timestamp}`;
                if (csvFiles.length > 1) {
                    fileName += `_part${index + 1}`;
                }
                fileName += '.csv';
                
                link.setAttribute('href', url);
                link.setAttribute('download', fileName);
                link.style.visibility = 'hidden';
                
                document.body.appendChild(link);
                
                setTimeout(() => {
                    link.click();
                    document.body.removeChild(link);
                    URL.revokeObjectURL(url);
                }, index * 200);
            });
            
            if (csvFiles.length > 1) {
                showAlert(`✅ Downloaded ${csvFiles.length} CSV files!`, 'success');
            } else {
                showAlert('✅ CSV file successfully downloaded!', 'success');
            }
        });
    }
    
    // Copy CSV button click handler
    if (copyCsvBtn) {
        copyCsvBtn.addEventListener('click', async () => {
            if (!csvFiles || csvFiles.length === 0) {
                showAlert('No data to copy. Please process data first.', 'error');
                return;
            }
            
            const csvContent = convertToCSV(csvFiles[0]);
            
            try {
                await navigator.clipboard.writeText(csvContent);
                copyCsvBtn.textContent = '✅ Copied!';
                if (csvFiles.length > 1) {
                    showAlert(`✅ First file (of ${csvFiles.length}) copied. Use download for other files.`, 'success');
                } else {
                    showAlert('✅ CSV data copied to clipboard!', 'success');
                }
                setTimeout(() => {
                    copyCsvBtn.textContent = '📋 Copy CSV';
                }, 2000);
            } catch (err) {
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
                        showAlert(`✅ First file (of ${csvFiles.length}) copied. Use download for other files.`, 'success');
                    } else {
                        showAlert('✅ CSV data copied to clipboard!', 'success');
                    }
                    setTimeout(() => {
                        copyCsvBtn.textContent = '📋 Copy CSV';
                    }, 2000);
                } catch (err) {
                    showAlert('Failed to copy. Try downloading the file.', 'error');
                }
                document.body.removeChild(textArea);
            }
        });
    }
}

// Function to show alert
function showAlert(message, type = 'info') {
    if (!alerts) {
        console.error('Alerts element not found');
        return;
    }
    const alert = document.createElement('div');
    alert.className = `alert alert-${type}`;
    alert.textContent = message;
    alerts.innerHTML = '';
    alerts.appendChild(alert);
    
    setTimeout(() => {
        if (alert.parentNode) {
            alert.remove();
        }
    }, 5000);
}


// Function to parse Excel data
function parseExcelData(text) {
    console.log('parseExcelData: Starting parse, text length:', text.length);
    console.log('parseExcelData: isParsing flag:', isParsing);
    
    // Check if text is placeholder or status message
    const placeholderTexts = ['Data pasted', 'processing', 'Click here', 'Expected columns', 'Successfully loaded'];
    const isPlaceholder = placeholderTexts.some(placeholder => text.toLowerCase().includes(placeholder.toLowerCase()));
    
    if (isPlaceholder || !text || text.trim().length < 50) {
        console.log('parseExcelData: Text is placeholder or too short, ignoring');
        isParsing = false;
        if (alerts) {
            showAlert('⚠️ Please paste actual Excel data. The text appears to be a status message, not data.', 'error');
        }
        if (pasteArea) {
            pasteArea.innerHTML = `
                <p><strong>Click here and paste the table</strong></p>
                <p style="font-size: 12px; color: var(--text-muted);">Expected columns: Epic, Story, Sprint, Description</p>
            `;
        }
        return;
    }
    
    try {
        const lines = text.split('\n').filter(line => line.trim());
        console.log('parseExcelData: Total lines after split:', lines.length);
        
        if (lines.length === 0) {
            showAlert('Error: No data to process', 'error');
            isParsing = false;
            if (pasteArea) {
                pasteArea.innerHTML = `
                    <p><strong>Click here and paste the table</strong></p>
                    <p style="font-size: 12px; color: var(--text-muted);">Expected columns: Epic, Story, Sprint, Description</p>
                `;
            }
            return;
        }
        
        // Parse header
        const headerLine = lines[0];
        console.log('parseExcelData: Header line:', headerLine);
        
        let headers = headerLine.split('\t').map(h => h.trim());
        if (headers.length === 1) {
            headers = headerLine.split(/[;,\t]/).map(h => h.trim());
        }
        console.log('parseExcelData: Headers found:', headers);
        
        // Find column indices - ожидаем колонки: Epic, Story, Sprint, Description
        const epicIdx = findColumnIndex(headers, ['epic']);
        const storyIdx = findColumnIndex(headers, ['story']);
        const sprintIdx = findColumnIndex(headers, ['sprint']);
        const descriptionIdx = findColumnIndex(headers, ['description', 'desc']);
        
        console.log('parseExcelData: Column indices - epic:', epicIdx, 'story:', storyIdx, 'sprint:', sprintIdx, 'description:', descriptionIdx);
        
        if (epicIdx === -1 || storyIdx === -1 || sprintIdx === -1 || descriptionIdx === -1) {
            const missingCols = [];
            if (epicIdx === -1) missingCols.push('Epic');
            if (storyIdx === -1) missingCols.push('Story');
            if (sprintIdx === -1) missingCols.push('Sprint');
            if (descriptionIdx === -1) missingCols.push('Description');
            
            showAlert(`Error: Required columns not found: ${missingCols.join(', ')}. Found headers: ${headers.join(', ')}`, 'error');
            isParsing = false;
            if (pasteArea) {
                pasteArea.innerHTML = `
                    <p><strong>Click here and paste the table</strong></p>
                    <p style="font-size: 12px; color: var(--text-muted);">Expected columns: Epic, Story, Sprint, Description</p>
                `;
            }
            return;
        }
        
        // Parse data rows
        const result = [];
        console.log('Starting to parse rows. Total lines:', lines.length);
        
        for (let i = 1; i < lines.length; i++) {
            let cells = lines[i].split('\t').map(c => c.trim());
            if (cells.length === 1) {
                cells = lines[i].split(/[;,\t]/).map(c => c.trim());
            }
            
            if (cells.length > Math.max(epicIdx, storyIdx, sprintIdx, descriptionIdx)) {
                const rowData = {
                    epic: cells[epicIdx] || '',
                    story: cells[storyIdx] || '',
                    sprint: cells[sprintIdx] || '',
                    description: cells[descriptionIdx] || ''
                };
                result.push(rowData);
            }
        }
        
        console.log('Total lines:', lines.length);
        console.log('Parsed data count:', result.length);
        console.log('First 5 rows:', result.slice(0, 5));
        
        // Store in global variable
        parsedTemplateData = result;
        
        // Update UI
        if (pasteArea) {
            pasteArea.innerHTML = `<p style="color: #10b981;"><strong>✅ Successfully loaded ${result.length} records</strong></p>`;
            pasteArea.classList.add('has-data');
        }
        
        // Show data preview
        updateDataPreview(result);
        
        showAlert(`Successfully loaded ${result.length} records`, 'success');
        
        isParsing = false;
        console.log('parseExcelData: Parsing completed successfully');
        
    } catch (error) {
        showAlert('Data processing error: ' + error.message, 'error');
        console.error(error);
        isParsing = false;
    }
}

function findColumnIndex(headers, possibleNames) {
    for (let i = 0; i < headers.length; i++) {
        const header = headers[i].toLowerCase();
        for (const name of possibleNames) {
            if (header.includes(name.toLowerCase())) {
                return i;
            }
        }
    }
    return -1;
}

function updateDataPreview(data) {
    if (!dataPreview) return;
    if (!data || data.length === 0) {
        dataPreview.classList.add('hidden');
        return;
    }
    
    let html = '<div class="data-preview">';
    html += `<h3>📊 Data Preview (${data.length} records):</h3>`;
    html += '<table>';
    html += '<thead><tr><th>Epic</th><th>Story</th><th>Sprint</th><th>Description</th></tr></thead>';
    html += '<tbody>';
    
    const rowsToShow = data.slice(0, 10);
    rowsToShow.forEach(row => {
        html += `<tr>
            <td>${escapeHtml(row.epic || '')}</td>
            <td>${escapeHtml(row.story || '')}</td>
            <td>${escapeHtml(row.sprint || '')}</td>
            <td>${escapeHtml(row.description || '')}</td>
        </tr>`;
    });
    
    if (data.length > 10) {
        html += `<tr><td colspan="4" style="text-align: center; color: var(--text-muted);">... and ${data.length - 10} more records</td></tr>`;
    }
    
    html += '</tbody></table></div>';
    
    dataPreview.innerHTML = html;
    dataPreview.classList.remove('hidden');
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Function to generate CSV structure for multiple teams
function generateProjectsStructure(templateData, component, label, numTeams) {
    if (!templateData || templateData.length === 0 || !component || !label || !numTeams) return [];
    
    const result = [];
    
    // Add headers
    result.push(['Epic Name', 'Epic Link', 'Label', 'Issue Type', 'Summary', 'Description', 'Component']);
    
    // Group template data by Epic
    const epicGroups = {};
    templateData.forEach(row => {
        const epic = row.epic.trim();
        if (!epic) return;
        
        if (!epicGroups[epic]) {
            epicGroups[epic] = [];
        }
        
        epicGroups[epic].push(row);
    });
    
    // Generate for each team
    for (let teamNum = 1; teamNum <= numTeams; teamNum++) {
        const teamComponent = `Team-${teamNum}_${component}`;
        
        // Process each epic group
        Object.keys(epicGroups).forEach(epicName => {
            const stories = epicGroups[epicName];
            
            // Generate Epic Name with team number
            let baseEpicName = epicName.trim();
            baseEpicName = baseEpicName.replace(/\s+T\d+_[A-Z0-9]+$/gi, '').trim();
            const epicNameWithTeam = `${baseEpicName} T${teamNum}_${component}`;
            
            // Get description from first story (if available)
            const firstStory = stories[0];
            const epicDescription = firstStory?.description || '';
            
            // Add Epic
            result.push([
                epicNameWithTeam,
                '', // Epic Link empty for Epic
                label, // Label only for Epics
                'Epic',
                epicNameWithTeam,
                epicDescription,
                teamComponent
            ]);
            
            // Add Stories for this Epic
            stories.forEach(story => {
                result.push([
                    '', // Epic Name empty for Stories
                    epicNameWithTeam, // Epic Link points to Epic Name
                    '', // Label empty for Stories
                    'Story',
                    story.story || '',
                    story.description || '',
                    teamComponent
                ]);
            });
        });
    }
    
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
    if (!preview) return;
    if (data.length === 0) {
        preview.innerHTML = '<div class="loading">No data to display</div>';
        return;
    }
    
    const maxCols = Math.max(...data.map(row => row.length));
    const normalizedData = data.map(row => {
        const normalizedRow = [...row];
        while (normalizedRow.length < maxCols) {
            normalizedRow.push('');
        }
        return normalizedRow;
    });
    
    let html = '<table>';
    
    // Headers
    if (normalizedData.length > 0) {
        html += '<thead><tr>';
        normalizedData[0].forEach(cell => {
            html += `<th>${escapeHtml(cell || '')}</th>`;
        });
        html += '</tr></thead>';
    }
    
    // Body
    html += '<tbody>';
    const displayData = normalizedData.slice(1);
    displayData.forEach((row, rowIndex) => {
        html += '<tr>';
        row.forEach(cell => {
            html += `<td>${escapeHtml(cell || '')}</td>`;
        });
        html += '</tr>';
    });
    html += '</tbody></table>';
    
    preview.innerHTML = html;
    
    // Update information
    const totalRows = normalizedData.length;
    let epicCount = 0;
    let storyCount = 0;
    let taskCount = 0;
    
    for (let i = 1; i < normalizedData.length; i++) {
        const issueType = normalizedData[i][3] || '';
        if (issueType.toLowerCase() === 'epic') {
            epicCount++;
        } else if (issueType.toLowerCase() === 'story') {
            storyCount++;
        } else if (issueType.toLowerCase() === 'task') {
            taskCount++;
        }
    }
    
    csvFiles = splitDataIntoFiles(normalizedData);
    const fileCount = csvFiles.length;
    
    // Show download buttons
    if (downloadBtn) downloadBtn.style.display = 'inline-block';
    if (copyCsvBtn) copyCsvBtn.style.display = 'inline-block';
    
    csvData = normalizedData;
    
    if (fileCount > 1) {
        showAlert(`✅ Generated ${epicCount} Epics, ${storyCount} Stories, and ${taskCount} Tasks! Data split into ${fileCount} files (max ${MAX_ROWS_PER_FILE} rows per file)`, 'success');
    } else {
        showAlert(`✅ Generated ${epicCount} Epics, ${storyCount} Stories, and ${taskCount} Tasks!`, 'success');
    }
}

