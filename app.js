let file1Obj = null;
let file2Obj = null;

let resultAvailable = [];
let resultMissing = [];
let activeTab = 'avail';

if (window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

// ─── Theme Toggling ──────────────────────────────────────────────────────────
function toggleTheme() {
    const isLight = document.body.classList.toggle('light-mode');
    localStorage.setItem('theme', isLight ? 'light' : 'dark');
    
    const icon = document.getElementById('theme-icon');
    
    icon.style.transform = 'rotate(360deg)';
    setTimeout(() => {
        icon.style.transition = 'none';
        icon.style.transform = 'rotate(0deg)';
        setTimeout(() => icon.style.transition = '', 50);
    }, 700);
    
    if (isLight) {
        icon.classList.remove('fa-moon');
        icon.classList.add('fa-sun');
    } else {
        icon.classList.remove('fa-sun');
        icon.classList.add('fa-moon');
    }
}

// ─── UI State ────────────────────────────────────────────────────────────────
function toggleUI() {
    const mode = document.querySelector('input[name="mode"]:checked').value;

    const f1Card        = document.getElementById('file1-card');
    const col1Container = document.getElementById('col1-container');
    const f1Title       = document.getElementById('file1-title');
    const col1Title     = document.getElementById('col1-title');

    const f2Card        = document.getElementById('file2-card');
    const col2Container = document.getElementById('col2-container');
    const f2Title       = document.getElementById('file2-title');
    const col2Title     = document.getElementById('col2-title');

    const metric1Title  = document.getElementById('metric1-title');
    const metric2Title  = document.getElementById('metric2-title');
    const metric1Card   = document.getElementById('metric1-card');
    const metricsGrid   = document.getElementById('metrics-grid');

    // Reset textual content
    f1Title.innerHTML    = '<i class="fa-regular fa-folder-open"></i> Old File';
    f2Title.innerHTML    = '<i class="fa-regular fa-folder-open"></i> New File';
    col1Title.innerText  = 'Old File Column';
    col2Title.innerText  = 'New File Column';
    metric1Title.innerText = 'File 1 Count';
    if(metric2Title) metric2Title.innerText = 'File 2 Count';
    
    // Reset Visibility smoothly
    f1Card.classList.remove('hidden');
    if(metric1Card) metric1Card.classList.remove('hidden');
    if(metricsGrid) {
        metricsGrid.classList.remove('grid-cols-2');
        metricsGrid.classList.add('grid-cols-3');
    }
    
    // Hide column dropdowns by default (only show after file is uploaded)
    col1Container.classList.remove('max-h-0', 'opacity-0'); // reset
    col2Container.classList.remove('max-h-0', 'opacity-0'); // reset
    col1Container.style.maxHeight = '0px';
    col1Container.style.opacity = '0';
    col2Container.style.maxHeight = '0px';
    col2Container.style.opacity = '0';
    
    const f2Input = document.getElementById('file2-input');
    f2Input.removeAttribute('multiple');

        if (mode === 'Extract numbers from the file') {
        f1Card.classList.add('hidden');
        if(metric1Card) metric1Card.classList.add('hidden');
        if(metricsGrid) {
            metricsGrid.classList.remove('grid-cols-3');
            metricsGrid.classList.add('grid-cols-2');
        }
        document.getElementById('metric1-val').innerText = '-';
        if(metric2Title) metric2Title.innerText = 'File Count';
        
        f2Input.setAttribute('multiple', 'true');
        f2Title.innerHTML = '<i class="fa-solid fa-layer-group"></i> Select Files (Multi-Select)';
    } else if (mode === 'Check Availability of Numbers') {
        f1Title.innerHTML      = '<i class="fa-solid fa-check-double"></i> Sorted numbers';
        f2Title.innerHTML      = '<i class="fa-solid fa-database"></i> All numbers';
        col1Title.innerText    = 'Sorted Column';
        col2Title.innerText    = 'All Numbers Column';
        metric1Title.innerText = 'Sorted List Count';
    }
}

// ─── Extraction Engine ───────────────────────────────────────────────────────
async function extractListFromFile(file, colId) {
    if (!file) return [];
    const filename = file.name.toLowerCase();
    let numbers = [];

    if (filename.endsWith('.txt')) {
        const text = await file.text();
        const lines = text.split('\n');
        for (let line of lines) {
            line = line.trim().replace(/-/g, "");
            if (line) {
                const matches = line.match(/\b\d{8,}\b/g);
                if (matches) numbers.push(...matches);
                // Fallback: if no regex match but it's a number, just add it
                else if (/^\d+$/.test(line)) numbers.push(line);
            }
        }
    } else if (filename.endsWith('.pdf')) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            const pageText = textContent.items.map(item => item.str).join(' ');
            const cleanText = pageText.replace(/-/g, "");
            const matches = cleanText.match(/\b\d{8,}\b/g);
            if (matches) numbers.push(...matches);
        }
    } else if (filename.endsWith('.xlsx') || filename.endsWith('.xls')) {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        let colIdx = parseInt(colId);
        let readAll = isNaN(colIdx);

        for (let row of jsonData) {
            if (readAll) {
                for (let c = 0; c < row.length; c++) {
                    let val = String(row[c] || "").trim();
                    if (val && !['nan', '', 'number', 'none'].includes(val.toLowerCase())) {
                        val = val.replace(/-/g, "");
                        const matches = val.match(/\d{8,}/g);
                        if (matches) numbers.push(...matches);
                    }
                }
            } else {
                if (row.length > colIdx) {
                    let val = String(row[colIdx] || "").trim();
                    if (val && !['nan', '', 'number', 'none'].includes(val.toLowerCase())) {
                        val = val.replace(/-/g, "");
                        const matches = val.match(/\d{8,}/g);
                        if (matches) numbers.push(...matches);
                    }
                }
            }
        }
    }
    return numbers;
}

// ─── File Upload Handlers ────────────────────────────────────────────────────
async function handleFileUpload(event, type) {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    const isOld  = type === 'old';
    const mode = document.querySelector('input[name="mode"]:checked').value;
    
    let file;
    if (isOld) {
        file1Obj = files[0];
        file = file1Obj;
        document.getElementById('file1-name').innerText = file.name;
    } else {
        if (mode === 'Extract numbers from the file') {
            file2Obj = Array.from(files);
            file = files[0]; // Use first file to detect columns
            document.getElementById('file2-name').innerText = files.length > 1 ? `${files.length} files selected` : file.name;
        } else {
            file2Obj = files[0];
            file = file2Obj;
            document.getElementById('file2-name').innerText = file.name;
        }
    }

    const selectEl = document.getElementById(`col${isOld ? 1 : 2}-select`);
    const filename = file.name.toLowerCase();

    selectEl.innerHTML = '';

    if (filename.endsWith('.txt')) {
        selectEl.innerHTML = '<option value="TXT">Text File (All Lines)</option>';
        selectEl.disabled = false;
        const container = document.getElementById(`col${isOld ? 1 : 2}-container`);
        container.classList.remove('max-h-0', 'opacity-0');
        container.style.maxHeight = '300px';
        container.style.opacity = '1';
    } else if (filename.endsWith('.pdf')) {
        selectEl.innerHTML = '<option value="PDF_AUTO">PDF Auto-Extract</option>';
        selectEl.disabled = false;
        const container = document.getElementById(`col${isOld ? 1 : 2}-container`);
        container.classList.remove('max-h-0', 'opacity-0');
        container.style.maxHeight = '300px';
        container.style.opacity = '1';
    } else if (filename.endsWith('.xlsx') || filename.endsWith('.xls')) {
        showLoading(true, "Reading Excel Columns...");
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            
            if (jsonData.length > 0) {
                const range = XLSX.utils.decode_range(sheet['!ref']);
                const numCols = range.e.c + 1;
                
                // Smart Header Detection: Find the row in the first 10 rows with the most filled columns
                let bestRow = [];
                let maxFilled = -1;
                for (let r = 0; r < Math.min(10, jsonData.length); r++) {
                    const row = jsonData[r] || [];
                    let filledCount = 0;
                    for (let c = 0; c < numCols; c++) {
                        if (row[c] !== undefined && String(row[c]).trim() !== "") {
                            filledCount++;
                        }
                    }
                    if (filledCount > maxFilled) {
                        maxFilled = filledCount;
                        bestRow = row;
                    }
                }
                const headerRow = bestRow;
                
                for (let i = 0; i < numCols; i++) {
                    const opt = document.createElement('option');
                    opt.value = i;
                    opt.textContent = headerRow[i] ? String(headerRow[i]) : `Column ${i+1}`;
                    selectEl.appendChild(opt);
                }
                selectEl.disabled = false;
                
                if (numCols > 2) {
                    selectEl.value = 2; // Default 3rd col
                } else if (numCols > 0) {
                    selectEl.value = 0;
                }
            } else {
                selectEl.disabled = true;
            }
            
            // Animate contextual dropdown open using inline styles for bulletproof CDN support
            const container = document.getElementById(`col${isOld ? 1 : 2}-container`);
            container.classList.remove('max-h-0', 'opacity-0');
            container.style.maxHeight = '300px';
            container.style.opacity = '1';
            
        } catch (err) {
            showAlert('Error', 'Failed to read Excel file: ' + err.message);
        }
        showLoading(false);
    }
}

// ─── Main Processing ─────────────────────────────────────────────────────────
async function processFiles() {
    const mode = document.querySelector('input[name="mode"]:checked').value;

    if (mode !== 'Extract numbers from the file') {
        if (!file1Obj || !file2Obj) {
            showAlert('Warning', 'Please load both files first.');
            return;
        }
    } else {
        if (!file2Obj) {
            showAlert('Warning', 'Please load the file first.');
            return;
        }
    }

    const btn = document.getElementById('process-btn');
    btn.disabled = true;
    showLoading(true, "Extracting numbers instantly...");

    try {
        const col2Id = document.getElementById('col2-select').value;
        let newCleaned = [];
        
        if (mode === 'Extract numbers from the file' && Array.isArray(file2Obj)) {
            for (let f of file2Obj) {
                const arr = await extractListFromFile(f, col2Id);
                newCleaned.push(...arr);
            }
        } else {
            const f2 = Array.isArray(file2Obj) ? file2Obj[0] : file2Obj;
            newCleaned = await extractListFromFile(f2, col2Id);
        }
        
        let oldCleaned = [];
        if (mode !== 'Extract numbers from the file') {
            const col1Id = document.getElementById('col1-select').value;
            oldCleaned = await extractListFromFile(file1Obj, col1Id);
        }

        const newSet = new Set(newCleaned);
        const oldSet = new Set(oldCleaned);

        resultAvailable = [];
        resultMissing = [];

        if (mode === "Newly added numbers") {
            // Numbers in new that are NOT in old
            let diff = [];
            for (let num of newCleaned) {
                if (!oldSet.has(num)) diff.push(num);
            }
            // Unique and sorted
            resultAvailable = [...new Set(diff)].sort();
        } else if (mode === "Extract numbers from the file") {
            // Unique from new
            resultAvailable = [...new Set(newCleaned)];
        } else if (mode === "Check Availability of Numbers") {
            // Which from OLD are in NEW?
            let avail = [];
            let missing = [];
            for (let num of oldCleaned) {
                if (newSet.has(num)) avail.push(num);
                else missing.push(num);
            }
            resultAvailable = [...new Set(avail)];
            resultMissing = [...new Set(missing)];
        }
        
        updateUI(oldCleaned.length, newCleaned.length);

    } catch (err) {
        showAlert('Error', err.message);
    }

    btn.disabled = false;
    showLoading(false);
    
    // Show results panel gracefully on mobile
    const resultsPanel = document.getElementById('results-panel');
    if (resultsPanel) {
        resultsPanel.classList.remove('hidden');
        // trigger reflow for smooth animation
        void resultsPanel.offsetWidth;
        resultsPanel.classList.remove('opacity-0', 'translate-y-10');
        resultsPanel.classList.add('opacity-100', 'translate-y-0');
        // smooth scroll down to results on small screens
        if (window.innerWidth < 768) {
            setTimeout(() => resultsPanel.scrollIntoView({ behavior: 'smooth', block: 'start' }), 300);
        }
    }
}

// ─── Update Results UI ────────────────────────────────────────────────────────
function updateUI(oldLen, newLen) {
    const mode = document.querySelector('input[name="mode"]:checked').value;

    if (mode !== 'Extract numbers from the file') {
        document.getElementById('metric1-val').innerText = oldLen;
    }
    document.getElementById('metric2-val').innerText  = newLen;
    document.getElementById('metric3-val').innerText  = resultAvailable.length;

    document.getElementById('txt-avail').value   = resultAvailable.join('\n');
    document.getElementById('txt-missing').value = resultMissing.join('\n');

    const tabMissing = document.getElementById('tab-missing');
    if (resultMissing.length > 0) {
        tabMissing.classList.remove('hidden');
        tabMissing.innerText = `Missing Numbers (${resultMissing.length})`;
    } else {
        tabMissing.classList.add('hidden');
    }

    switchTab('avail');
    document.getElementById('verify-btn').disabled = (resultAvailable.length === 0);
    document.getElementById('split-btn').disabled = (resultAvailable.length === 0);
}

// ─── Tab Switching ────────────────────────────────────────────────────────────
function switchTab(tab) {
    activeTab = tab;
    const tabAvail   = document.getElementById('tab-avail');
    const tabMissing = document.getElementById('tab-missing');
    const txtAvail   = document.getElementById('txt-avail');
    const txtMissing = document.getElementById('txt-missing');

    const accentClass  = ['border-accent', 'text-textMain'];
    const inactClass   = ['border-transparent', 'text-textMuted'];

    if (tab === 'avail') {
        accentClass.forEach(c => tabAvail.classList.add(c));
        inactClass.forEach(c  => tabAvail.classList.remove(c));
        inactClass.forEach(c  => tabMissing.classList.add(c));
        accentClass.forEach(c => tabMissing.classList.remove(c));
        txtAvail.classList.remove('hidden');
        txtMissing.classList.add('hidden');
    } else {
        accentClass.forEach(c => tabMissing.classList.add(c));
        inactClass.forEach(c  => tabMissing.classList.remove(c));
        inactClass.forEach(c  => tabAvail.classList.add(c));
        accentClass.forEach(c => tabAvail.classList.remove(c));
        txtMissing.classList.remove('hidden');
        txtAvail.classList.add('hidden');
    }
}

function getActiveList() {
    return activeTab === 'avail' ? resultAvailable : resultMissing;
}

// ─── Master List Verification ─────────────────────────────────────────────────
function verifyAgainstMaster() {
    if (resultAvailable.length === 0) {
        showAlert('Warning', 'No results to verify yet. Process files first.');
        return;
    }
    document.getElementById('master-input').click();
}

async function handleMasterUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    showLoading(true, "Verifying against Master List...");
    const vBtn = document.getElementById('verify-btn');
    vBtn.disabled = true;

    try {
        const masterCleaned = await extractListFromFile(file, "all"); // Read ALL columns for verify
        const masterSet = new Set(masterCleaned);

        let avail = [];
        let missing = [];

        for (let num of resultAvailable) {
            if (masterSet.has(num)) avail.push(num);
            else missing.push(num);
        }

        resultAvailable = avail;
        resultMissing = missing;

        updateUI(document.getElementById('metric1-val').innerText, masterCleaned.length);
        showAlert('Verification Complete',
            `✅ Available in master: ${resultAvailable.length}\n❌ Missing from master: ${resultMissing.length}`);

    } catch (err) {
        showAlert('Error', 'Verification failed:\n' + err.message);
    }

    showLoading(false);
    vBtn.disabled = (resultAvailable.length === 0);
    event.target.value = '';
}

// ─── Export & Verify ─────────────────────────────────────────────────────────
async function copyToClipboard() {
    const list = getActiveList();
    if (!list.length) { showAlert('Info', 'No results to copy.'); return; }
    try {
        await navigator.clipboard.writeText(list.join('\n'));
        showAlert('Copied!', `${list.length} numbers copied to clipboard.`);
    } catch {
        showAlert('Error', 'Clipboard access denied. Select all text in the box and copy manually.');
    }
}

function saveTXT() {
    const list = getActiveList();
    if (!list.length) { showAlert('Info', 'No results to save.'); return; }
    downloadBlob(new Blob([list.join('\n')], { type: 'text/plain' }), 'extracted_numbers.txt');
}

function saveExcel() {
    const list = getActiveList();
    if (!list.length) { showAlert('Info', 'No results to save.'); return; }
    // No # counting as requested
    const ws = XLSX.utils.json_to_sheet(list.map((n) => ({ Numbers: n })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Results');
    XLSX.writeFile(wb, 'extracted_numbers.xlsx');
}

function savePDF() {
    const list = getActiveList();
    if (!list.length) { showAlert('Info', 'No results to save.'); return; }

    const { jsPDF } = window.jspdf;
    const doc   = new jsPDF();
    const label = activeTab === 'avail' ? 'Available Numbers' : 'Missing Numbers';

    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text(label, 105, 14, null, null, 'center');
    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    doc.text(`Total: ${list.length}`, 105, 20, null, null, 'center');
    doc.setFontSize(11);

    let y = 30;
    for (let i = 0; i < list.length; i++) {
        if (y > 280) { doc.addPage(); y = 20; }
        doc.text(`${i + 1}. ${list[i]}`, 10, y);
        y += 7;
    }
    doc.save('extracted_numbers.pdf');
}

function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a   = Object.assign(document.createElement('a'), { href: url, download: filename });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// ─── Alert / Loading UI ───────────────────────────────────────────────────────
function showAlert(title, msg) {
    document.getElementById('alert-title').innerText = title;
    document.getElementById('alert-msg').innerText   = msg;
    const modal = document.getElementById('alert-modal');
    modal.classList.remove('hidden');
    requestAnimationFrame(() => {
        modal.classList.remove('opacity-0');
        modal.firstElementChild.classList.remove('scale-95');
    });
}

function closeAlert() {
    const modal = document.getElementById('alert-modal');
    modal.classList.add('opacity-0');
    modal.firstElementChild.classList.add('scale-95');
    setTimeout(() => modal.classList.add('hidden'), 300);
}

function showLoading(show, text="Processing...") {
    const ov = document.getElementById('loading-overlay');
    if (show) {
        document.getElementById('loading-text').innerText = text;
        ov.classList.remove('hidden');
    } else {
        ov.classList.add('hidden');
    }
}

// ─── Split Feature ────────────────────────────────────────────────────────────
function openSplitModal() {
    const modal = document.getElementById('split-modal');
    modal.classList.remove('hidden');
    requestAnimationFrame(() => {
        modal.classList.remove('opacity-0');
        modal.firstElementChild.classList.remove('scale-95');
    });
}

function closeSplitModal() {
    const modal = document.getElementById('split-modal');
    modal.classList.remove('opacity-100');
    modal.classList.add('opacity-0');
    setTimeout(() => modal.classList.add('hidden'), 300);
}

// ─── Help Modal ──────────────────────────────────────────────────────────────
function openHelpModal() {
    const modal = document.getElementById('help-modal');
    modal.classList.remove('hidden');
    // trigger reflow
    void modal.offsetWidth;
    modal.classList.remove('opacity-0');
    modal.classList.add('opacity-100');
}

function closeHelpModal() {
    const modal = document.getElementById('help-modal');
    modal.classList.remove('opacity-100');
    modal.classList.add('opacity-0');
    setTimeout(() => modal.classList.add('hidden'), 300);
}

async function executeSplit() {
    const parts = parseInt(document.getElementById('split-parts').value);
    if (isNaN(parts) || parts < 2) {
        showAlert('Error', 'Please enter a valid number of parts (minimum 2).');
        return;
    }

    const list = getActiveList();
    if (list.length === 0) {
        showAlert('Error', 'No data to split.');
        return;
    }

    const chunkSize = Math.floor(list.length / parts);
    const remainder = list.length % parts;

    closeSplitModal();
    showLoading(true, "Generating split files...");

    let idx = 0;
    for (let i = 0; i < parts; i++) {
        const size = chunkSize + (i < remainder ? 1 : 0);
        const chunk = list.slice(idx, idx + size);
        idx += size;
        
        const blob = new Blob([chunk.join('\n')], { type: 'text/plain' });
        downloadBlob(blob, `split_part_${i + 1}.txt`);
        
        await new Promise(r => setTimeout(r, 200));
    }

    showLoading(false);
    showAlert('Success', `Successfully downloaded ${parts} files!`);
}

// ─── Init ─────────────────────────────────────────────────────────────────────
toggleUI();

window.addEventListener('load', () => {
    // Theme persistence check
    const savedTheme = localStorage.getItem('theme');
    if (savedTheme === 'light') {
        document.body.classList.add('light-mode');
        const icon = document.getElementById('theme-icon');
        if(icon) {
            icon.classList.remove('fa-moon');
            icon.classList.add('fa-sun');
        }
    }

    // Elegant splash screen dismissal (shortened per request)
    setTimeout(() => {
        const loader = document.getElementById('initial-loader');
        if (loader) {
            loader.style.opacity = '0';
            setTimeout(() => loader.remove(), 700);
        }
    }, 400); // 0.4s initial splash display
});
