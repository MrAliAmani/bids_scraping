const socket = io('/', {
    transports: ['websocket'],
    upgrade: false
});

let scripts = [];
let logModal;

document.addEventListener('DOMContentLoaded', () => {
    // Setup socket error handling
    socket.on('connect_error', (error) => {
        console.error('Socket connection error:', error);
    });
    
    socket.on('connect', () => {
        console.log('Socket connected successfully');
    });

    logModal = new bootstrap.Modal(document.getElementById('logModal'));
    
    // Initialize event listeners
    document.getElementById('startAllBtn').addEventListener('click', startAllScripts);
    document.getElementById('stopAllBtn').addEventListener('click', stopAllScripts);
    document.getElementById('startAppBtn').addEventListener('click', startApp);
    document.getElementById('stopAppBtn').addEventListener('click', stopApp);

    // Set up Socket.IO listeners
    socket.on('script_output', handleScriptOutput);
    socket.on('progress_update', handleProgressUpdate);
    socket.on('main_log', handleMainLog);
    socket.on('script_update', (data) => {
        console.log('[DEBUG] Received script_update:', data);
        const scriptIndex = scripts.findIndex(s => s.name === data.script);
        if (scriptIndex >= 0) {
            console.log('[DEBUG] Before update - Script state:', {
                name: scripts[scriptIndex].name,
                status: scripts[scriptIndex].status,
                excel_status: scripts[scriptIndex].excel_status,
                excel_progress: scripts[scriptIndex].excel_progress
            });
            
            scripts[scriptIndex].status = data.status;
            scripts[scriptIndex].excel_status = data.excel_status;
            scripts[scriptIndex].excel_progress = data.excel_progress;
            
            console.log('[DEBUG] After update - Script state:', {
                name: scripts[scriptIndex].name,
                status: scripts[scriptIndex].status,
                excel_status: scripts[scriptIndex].excel_status,
                excel_progress: scripts[scriptIndex].excel_progress
            });
            
            updateScriptDisplay();
        } else {
            console.log('[DEBUG] Script not found:', data.script);
        }
    });

    // Start polling for updates
    fetchInitialData();
    setInterval(fetchUpdates, 2000);
    setInterval(checkAppStatus, 5000);
});

function fetchInitialData() {
    fetch('/api/scripts')
        .then(response => response.json())
        .then(data => {
            console.log('[DEBUG] Fetched initial data:', data);
            scripts = data.map(script => ({
                ...script,
                excel_status: script.excel_status || 'Pending',
                excel_progress: script.excel_progress || 0
            }));
            updateScriptDisplay();
            updateStatistics();
        })
        .catch(error => console.error('Error fetching initial data:', error));
}

function fetchUpdates() {
    Promise.all([
        fetch('/api/scripts').then(r => r.json()),
        fetch('/api/master/status').then(r => r.json())
    ]).then(([scriptsData, statusData]) => {
        console.log('[DEBUG] Fetched updates:', scriptsData);
        scripts = scriptsData.map(script => ({
            ...script,
            excel_status: script.excel_status || 'Pending',
            excel_progress: script.excel_progress || 0
        }));
        updateScriptDisplay();
        updateStatistics(statusData);
    }).catch(error => console.error('Error fetching updates:', error));
}

function updateScriptDisplay() {
    const container = document.getElementById('scriptContainer');
    container.innerHTML = scripts.map(script => createScriptCard(script)).join('');
    
    // Add event listeners to new elements
    scripts.forEach(script => {
        const card = document.getElementById(`script-${script.name}`);
        if (card) {
            card.querySelector('.start-btn')?.addEventListener('click', () => startScript(script.name));
            card.querySelector('.stop-btn')?.addEventListener('click', () => stopScript(script.name));
            card.querySelector('.logs-btn')?.addEventListener('click', () => showLogs(script));
        }
    });
}

function formatRuntime(runtime) {
    if (!runtime) return '00:00:00';
    // If runtime contains milliseconds or microseconds, remove them
    if (runtime.includes('.')) {
        runtime = runtime.split('.')[0];
    }
    return runtime;
}

function createScriptCard(script) {
    const statusClass = getStatusClass(script.status);
    const excelStatusClass = getExcelStatusClass(script.excel_status || 'Pending');
    const progressWidth = script.progress || 0;
    const excelProgress = script.excel_progress || 0;
    const formattedRuntime = formatRuntime(script.runtime);
    
    // Extract script name without path and extension
    const scriptBaseName = script.name.split('/').pop().replace('.py', '');
    // Truncate name if longer than 25 characters
    const truncatedName = scriptBaseName.length > 25 ? scriptBaseName.substring(0, 22) + '...' : scriptBaseName;
    
    return `
        <div class="col-md-4 mb-4" id="script-${script.name}">
            <div class="script-card ${script.status === 'Running' || script.excel_status === 'Running' ? 'processing' : ''}">
                <div class="d-flex justify-content-between align-items-center mb-2">
                    <div class="text-truncate" style="max-width: 70%;">
                        <h5 class="mb-0" title="${scriptBaseName}">${truncatedName}</h5>
                    </div>
                    <div class="d-flex flex-column align-items-end">
                        <span class="status-badge ${statusClass}">${script.status}</span>
                    </div>
                </div>
                
                <div class="progress mb-2">
                    <div class="progress-bar" role="progressbar" 
                         style="width: ${progressWidth}%" 
                         aria-valuenow="${progressWidth}" 
                         aria-valuemin="0" 
                         aria-valuemax="100">
                        ${progressWidth}%
                    </div>
                </div>

                <div class="excel-status">
                    <div class="d-flex justify-content-between align-items-center">
                        <span class="status-badge ${excelStatusClass} small">Excel: ${script.excel_status || 'Pending'}</span>
                        <span class="small text-muted">${excelProgress}%</span>
                    </div>
                    <div class="progress" style="margin-top: 5px;">
                        <div class="progress-bar ${excelStatusClass}" role="progressbar" 
                             style="width: ${excelProgress}%" 
                             aria-valuenow="${excelProgress}" 
                             aria-valuemin="0" 
                             aria-valuemax="100">
                        </div>
                    </div>
                </div>

                <div class="d-flex justify-content-between align-items-center mb-3">
                    <span class="text-muted">Runtime: ${formattedRuntime}</span>
                </div>

                <div class="control-buttons d-flex gap-2">
                    <button class="btn btn-primary flex-grow-1 start-btn" ${script.status === 'Running' || script.excel_status === 'Running' ? 'disabled' : ''}>
                        <i class="fas fa-play me-2"></i>Start
                    </button>
                    <button class="btn btn-danger flex-grow-1 stop-btn" ${script.status !== 'Running' ? 'disabled' : ''}>
                        <i class="fas fa-stop me-2"></i>Stop
                    </button>
                    <button class="btn btn-secondary flex-grow-1 logs-btn">
                        <i class="fas fa-list me-2"></i>Logs
                    </button>
                </div>
            </div>
        </div>
    `;
}

function getStatusClass(status) {
    switch (status.toLowerCase()) {
        case 'running': return 'status-running';
        case 'success': return 'status-success';
        case 'error': return 'status-error';
        default: return 'status-pending';
    }
}

function getExcelStatusClass(status) {
    switch (status.toLowerCase()) {
        case 'running': return 'status-running';
        case 'done': return 'status-success';
        default: return 'status-pending';
    }
}

function updateStatistics(data = {}) {
    document.getElementById('runningCount').textContent = data.running || 0;
    document.getElementById('completedCount').textContent = data.completed || 0;
    document.getElementById('failedCount').textContent = data.failed || 0;
    document.getElementById('totalCount').textContent = data.total || scripts.length;
}

function startScript(scriptName) {
    fetch('/api/start', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ script: scriptName })
    }).then(() => fetchUpdates());
}

function stopScript(scriptName) {
    fetch('/api/stop', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ script: scriptName })
    })
    .then(response => response.json())
    .then(data => {
        if (data.status === 'success') {
            appendToMainLog(`Script ${scriptName} stopped successfully`);
            // Immediately fetch updates to show new script starting
            fetchUpdates();
        } else {
            appendToMainLog(`Error stopping script ${scriptName}: ${data.message}`);
        }
    })
    .catch(error => {
        console.error('Error stopping script:', error);
        appendToMainLog(`Error stopping script ${scriptName}`);
    });
}

function startAllScripts() {
    fetch('/api/start', { method: 'POST' })
        .then(() => fetchUpdates());
}

function stopAllScripts() {
    // Get all running scripts
    const runningScripts = scripts.filter(script => script.status === 'Running');
    
    if (runningScripts.length === 0) {
        appendToMainLog('No running scripts to stop');
        return;
    }

    appendToMainLog(`Stopping ${runningScripts.length} running scripts...`);
    
    // Create an array of promises for each stop request
    const stopPromises = runningScripts.map(script => {
        return fetch('/api/stop', { 
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ script: script.name })
        })
        .then(response => response.json())
        .then(data => {
            if (data.status === 'success') {
                appendToMainLog(`Script ${script.name} stopped successfully`);
            } else {
                appendToMainLog(`Error stopping script ${script.name}: ${data.message}`);
            }
        })
        .catch(error => {
            console.error(`Error stopping script ${script.name}:`, error);
            appendToMainLog(`Error stopping script ${script.name}`);
        });
    });

    // Wait for all stop requests to complete
    Promise.all(stopPromises)
        .then(() => {
            appendToMainLog('All stop requests completed');
            fetchUpdates();  // Update the UI
        })
        .catch(error => {
            console.error('Error in stop all operation:', error);
            appendToMainLog('Error stopping all scripts');
        });
}

function showLogs(script) {
    const logViewer = document.getElementById('logViewer');
    const scriptPath = script.name.replace(/\\/g, '/');  // Normalize path separators
    
    // Show loading message with spinner
    logViewer.innerHTML = `
        <div class="text-center p-4">
            <div class="spinner-border text-primary mb-2" role="status">
                <span class="visually-hidden">Loading...</span>
            </div>
            <p>Loading logs for ${scriptPath}...</p>
        </div>
    `;
    
    // Get yesterday's date for the log file path
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const dateStr = yesterday.toISOString().split('T')[0];
    
    // Get script base name without extension
    const scriptBaseName = scriptPath.split('/').pop().replace('.py', '');
    
    // Construct log file path
    const logPath = `${dateStr}/${scriptBaseName}.log`;
    
    // Fetch logs from the server
    fetch(`/api/logs/${encodeURIComponent(scriptPath)}`)
        .then(response => response.json())
        .then(data => {
            if (data.status === 'success' && data.content) {
                // Display the logs exactly as they appear in PowerShell
                logViewer.innerHTML = `
                    <div class="log-content">
                        <div class="d-flex justify-content-between align-items-center mb-2">
                            <h6 class="mb-0">Log file: ${logPath}</h6>
                            <button onclick="copyLogs()" class="btn btn-sm btn-outline-secondary">
                                <i class="fas fa-copy me-1"></i>Copy
                            </button>
                        </div>
                        <pre class="m-0 p-3 bg-dark text-light powershell-log" style="white-space: pre-wrap; font-family: 'Consolas', monospace; font-size: 0.9rem; max-height: 500px; overflow-y: auto;">${data.content}</pre>
                    </div>
                `;

                // Scroll to the bottom of the log viewer
                const preElement = logViewer.querySelector('pre');
                preElement.scrollTop = preElement.scrollHeight;
            } else {
                // Show appropriate message based on status
                const icon = data.status === 'pending' ? 'clock' : 'info-circle';
                logViewer.innerHTML = `
                    <div class="alert alert-info m-3">
                        <i class="fas fa-${icon} me-2"></i>
                        ${data.message || 'No logs available yet. Logs will appear here once the script starts running.'}
                    </div>
                `;
            }
        })
        .catch(error => {
            console.error('Error fetching logs:', error);
            logViewer.innerHTML = `
                <div class="alert alert-danger m-3">
                    <i class="fas fa-exclamation-circle me-2"></i>
                    Error loading logs. Please try again.
                </div>
            `;
        });
    
    logModal.show();
}

// Add copy functionality
function copyLogs() {
    const preElement = document.querySelector('.log-content pre');
    if (preElement) {
        const text = preElement.textContent;
        navigator.clipboard.writeText(text).then(() => {
            // Show a temporary success message
            const button = document.querySelector('.log-content button');
            const originalText = button.innerHTML;
            button.innerHTML = '<i class="fas fa-check me-1"></i>Copied!';
            button.classList.add('btn-success');
            button.classList.remove('btn-outline-secondary');
            setTimeout(() => {
                button.innerHTML = originalText;
                button.classList.remove('btn-success');
                button.classList.add('btn-outline-secondary');
            }, 2000);
        });
    }
}

function handleScriptOutput(data) {
    const scriptIndex = scripts.findIndex(s => s.name === data.script);
    if (scriptIndex >= 0) {
        if (!scripts[scriptIndex].logs) scripts[scriptIndex].logs = [];
        scripts[scriptIndex].logs.push(data.output);
        updateScriptDisplay();
    }
    appendToMainLog(data.output);
}

function handleProgressUpdate(data) {
    const scriptIndex = scripts.findIndex(s => s.name === data.script);
    if (scriptIndex >= 0) {
        if (data.type === 'excel_processing') {
            scripts[scriptIndex].excel_status = data.status;
            scripts[scriptIndex].excel_progress = data.progress;
        } else {
            scripts[scriptIndex].progress = data.progress;
        }
        if (!scripts[scriptIndex].logs) scripts[scriptIndex].logs = [];
        scripts[scriptIndex].logs.push(data.message);
        updateScriptDisplay();
    }
    appendToMainLog(data.message);
}

function handleMainLog(data) {
    appendToMainLog(data.message);
}

function appendToMainLog(message) {
    const mainLog = document.getElementById('mainLog');
    const logEntry = document.createElement('div');
    logEntry.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;
    mainLog.appendChild(logEntry);
    mainLog.scrollTop = mainLog.scrollHeight;
}

function checkAppStatus() {
    fetch('/api/app/status')
        .then(response => response.json())
        .then(data => {
            const btn = document.getElementById('appStatusBtn');
            if (data.status === 'running') {
                btn.classList.remove('btn-danger');
                btn.classList.add('btn-success');
                btn.innerHTML = '<i class="fas fa-circle me-2"></i>App Running';
            } else {
                btn.classList.remove('btn-success');
                btn.classList.add('btn-danger');
                btn.innerHTML = '<i class="fas fa-circle me-2"></i>App Stopped';
            }
        })
        .catch(error => console.error('Error checking app status:', error));
}

function startApp() {
    fetch('/api/app/start', { method: 'POST' })
        .then(response => response.json())
        .then(data => {
            if (data.status === 'started') {
                checkAppStatus();
                appendToMainLog('Application started successfully');
            }
        })
        .catch(error => {
            console.error('Error starting app:', error);
            appendToMainLog('Error starting application');
        });
}

function stopApp() {
    fetch('/api/app/stop', { method: 'POST' })
        .then(response => response.json())
        .then(data => {
            if (data.status === 'stopped') {
                checkAppStatus();
                appendToMainLog('Application stopped successfully');
                // Update UI to reflect stopped state
                const btn = document.getElementById('appStatusBtn');
                btn.classList.remove('btn-success');
                btn.classList.add('btn-danger');
                btn.innerHTML = '<i class="fas fa-circle me-2"></i>App Stopped';
            }
        })
        .catch(error => {
            console.error('Error stopping app:', error);
            appendToMainLog('Error stopping application');
        });
}

// Add PowerShell log styles to the existing styles
const style = document.createElement('style');
style.textContent = `
    .powershell-log {
        background-color: #012456 !important;
        color: #EEEDF0 !important;
        padding: 10px !important;
        font-family: 'Consolas', monospace !important;
        line-height: 1.2 !important;
    }
    .powershell-log::-webkit-scrollbar {
        width: 12px;
    }
    .powershell-log::-webkit-scrollbar-track {
        background: #012456;
    }
    .powershell-log::-webkit-scrollbar-thumb {
        background-color: #EEEDF0;
        border: 3px solid #012456;
        border-radius: 6px;
    }
`;
document.head.appendChild(style); 