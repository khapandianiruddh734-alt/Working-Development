/* --- GLOBAL STATE & CONFIG --- */
let currentTool = null;
let globalFileList = []; 
let rawDupData = [];
let activeDupFileName = "";

const forcedCorrections = {
    "utapam": "Uttapam", "uttapam": "Uttapam", "muglai": "Mughlai", "mughlai": "Mughlai",
    "schez": "Schezwan", "schz": "Schezwan", "schzwn": "Schezwan", "rot": "Roti",
    "nan": "Naan", "nun": "Naan", "naa": "Naan", "naan": "Naan", "tika": "Tikka",
    "tikka": "Tikka", "seek": "Seekh", "seekh": "Seekh", "agli": "Aglio", "aglio": "Aglio",
    "piza": "Pizza", "pizaa": "Pizza", "pulaw": "Pulao", "pulao": "Pulao", "plao": "Pulao",
    "zeera": "Jeera", "jeera": "Jeera", "kabab": "Kebab", "kebab": "Kebab",
    "chkn": "Chicken", "chk": "Chicken", "pnr": "Paneer", "paner": "Paneer", "panner": "Paneer",
    "btr": "Butter", "buttr": "Butter", "veg": "Veg", "vge": "Veg", "frd": "Fried",
    "brgr": "Burger", "sndwch": "Sandwich", "sambar": "Sambhar", "sambhar": "Sambhar",
    "kadai": "Kadhai", "kadhai": "Kadhai"
};

const dictionary = [
    "Chicken", "Paneer", "Rice", "Dal", "Roti", "Naan", "Kulcha", "Paratha",
    "Butter", "Cheese", "Cream", "Curd", "Yogurt", "Milk", "Mayo",
    "Tikka", "Masala", "Makhani", "Kadhai", "Handi", "Tawa", "Bhuna", "Lababdar",
    "Biryani", "Pulao", "Jeera", "Fried", "Schezwan", "Hakka", "Manchurian",
    "Noodles", "Soup", "Salad", "Raita", "Papad", "Pickle", "Chutney",
    "Pizza", "Pasta", "Burger", "Sandwich", "Fries", "Momos", "Roll", "Wrap",
    "Aglio", "Olio", "Arrabiata", "Alfredo", "Pesto", "Lasagna", "Risotto",
    "Uttapam", "Mughlai", "Sambhar", "Idli", "Dosa", "Vada", "Pav", "Bhaji",
    "Coke", "Pepsi", "Soda", "Water", "Juice", "Shake", "Lassi", "Mojito",
    "Aloo", "Gobi", "Matar", "Palak", "Methi", "Bhindi", "Corn", "Mushroom",
    "Onion", "Tomato", "Garlic", "Ginger", "Chilli", "Capsicum",
    "Sauteed", "Grilled", "Roasted", "Steamed", "Boiled", "Baked", "Tandoori",
    "Seekh", "Kebab", "Reshmi", "Malai", "Afghani", "Achari", "Amritsari",
    "Regular", "Medium", "Large", "Small", "Full", "Half", "Quarter", "Pcs"
];

const tools = {
    'jpg-to-pdf': { id: 'jpg-to-pdf', title: "JPG to PDF", desc: "Add multiple images to combine into a single PDF.", accept: "image/jpeg, image/png, image/jpg", handler: convertJpgToPdf, multiple: true },
    'word-to-pdf': { id: 'word-to-pdf', title: "Word to PDF", desc: "Convert DOCX to PDF.", accept: ".docx", handler: convertWordToPdf },
    'excel-to-pdf': { id: 'excel-to-pdf', title: "Excel to PDF", desc: "Convert Sheets to PDF.", accept: ".xlsx, .xls", handler: convertExcelToPdf },
    'pdf-to-jpg': { id: 'pdf-to-jpg', title: "PDF to JPG", desc: "Extract pages as images.", accept: ".pdf", handler: convertPdfToJpg },
    'compress-pdf': { id: 'compress-pdf', title: "Compress PDF", desc: "Reduce file size while preserving text quality.", accept: ".pdf", handler: compressPDF },
    'pdf-image-to-excel': { id: 'pdf-image-to-excel', title: "PDF/Image to Excel", desc: "Extract data to Excel.", accept: ".pdf, image/*", handler: convertPdfImageToExcel },
    'clean-excel': { id: 'clean-excel', title: "Clean Excel Sheet", desc: "Remove accents and special symbols.", accept: ".xlsx, .csv", handler: cleanExcelFile },
    'duplicate-remover': { id: 'duplicate-remover', title: "Safe Duplicate Remover", desc: "Highlight or Remove duplicate rows safely.", accept: ".xlsx, .xls, .csv", handler: null },
    'menu-fixer': { id: 'menu-fixer', title: "Menu Fixer (Beta)", desc: "Auto-correct restaurant menu spellings (e.g. Utapam -> Uttapam).", accept: ".xlsx, .xls", handler: processMenuFixer }
};

/* --- CORE UI FUNCTIONS --- */
function loadTool(id) {
    console.log('Loading tool:', id);
    currentTool = tools[id];
    if(!currentTool) {
        console.error('Tool not found:', id);
        return;
    }
    
    globalFileList = [];
    document.getElementById('menu-view').classList.add('hidden');
    document.getElementById('workspace-view').classList.remove('hidden');
    document.getElementById('tool-title').innerText = currentTool.title;
    document.getElementById('tool-desc').innerText = currentTool.desc;
    document.getElementById('file-formats').innerText = "Supports: " + currentTool.accept;
    document.getElementById('success-details').innerText = "";
    
    const fileInput = document.getElementById('file-input');
    fileInput.accept = currentTool.accept;

    if (currentTool.multiple) {
        fileInput.setAttribute('multiple', '');
        document.querySelector('#drop-area span.text-lg').innerText = "Click to Add File(s)";
    } else {
        fileInput.removeAttribute('multiple');
        document.querySelector('#drop-area span.text-lg').innerText = "Click to Choose File";
    }

    if (currentTool.id === 'compress-pdf') {
        document.getElementById('compress-options-area').classList.remove('hidden');
    }

    resetUI();
}

function resetUI() {
    document.getElementById('status-area').classList.add('hidden');
    document.getElementById('loading-spinner').classList.add('hidden');
    document.getElementById('success-msg').classList.add('hidden');
    document.getElementById('error-msg').classList.add('hidden');
    document.getElementById('drop-area').classList.remove('hidden');
    document.getElementById('preview-area').classList.add('hidden');
    document.getElementById('dup-options-area').classList.add('hidden');
    document.getElementById('compress-options-area').classList.add('hidden');
    document.getElementById('preview-grid').innerHTML = '';
    document.getElementById('file-input').value = '';
}

/* --- FILE HANDLING --- */
async function handleFile(files) {
    if (files.length === 0) return;

    if (currentTool.id === 'jpg-to-pdf') {
        globalFileList = [...globalFileList, ...Array.from(files)];
        updatePreviewUI();
        return;
    }

    if (currentTool.id === 'duplicate-remover') {
        handleDuplicateFile(files[0]);
        return;
    }

    if (currentTool.id === 'compress-pdf') {
        if (files[0].size > 50 * 1024 * 1024) {
            alert("File too large. Please use a PDF under 50MB.");
            return;
        }
        globalFileList = Array.from(files);
        // Options are shown, wait for user to start
        return;
    }

    startProcessingImplementation(files);
}

function updatePreviewUI() {
    const previewArea = document.getElementById('preview-area');
    const previewGrid = document.getElementById('preview-grid');
    if (globalFileList.length > 0) {
        previewArea.classList.remove('hidden');
        previewGrid.innerHTML = '';
        globalFileList.forEach(file => {
            if (file.type.startsWith('image/')) {
                const img = document.createElement('img');
                img.src = URL.createObjectURL(file);
                img.className = 'preview-thumb shadow-sm';
                previewGrid.appendChild(img);
            }
        });
    }
}

async function startCompress() {
    if (globalFileList.length === 0) {
        alert("Please upload a PDF first.");
        return;
    }
    document.getElementById('compress-options-area').classList.add('hidden');
    startProcessingImplementation(globalFileList);
}

async function startProcessingImplementation(filesToProcess) {
    document.getElementById('drop-area').classList.add('hidden');
    document.getElementById('status-area').classList.remove('hidden');
    document.getElementById('loading-spinner').classList.remove('hidden');

    try {
        await currentTool.handler(filesToProcess);
        document.getElementById('loading-spinner').classList.add('hidden');
        document.getElementById('success-msg').classList.remove('hidden');
    } catch (err) {
        console.error(err);
        document.getElementById('loading-spinner').classList.add('hidden');
        document.getElementById('error-msg').innerText = "Error: " + err.message;
        document.getElementById('error-msg').classList.remove('hidden');
        document.getElementById('drop-area').classList.remove('hidden');
    }
}

/* --- MENU FIXER LOGIC --- */
function normalizeMenuString(str) {
    return str
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "") 
        .replace(/—/g, "-").replace(/–/g, "-") 
        .replace(/["]/g, '').replace(/[']/g, '');
}

function getSimilarity(s1, s2) {
    let longer = s1, shorter = s2;
    if (s1.length < s2.length) { longer = s2; shorter = s1; }
    let longerLength = longer.length;
    if (longerLength === 0) return 1.0;
    return (longerLength - editDistance(longer, shorter)) / parseFloat(longerLength);
}

function editDistance(s1, s2) {
    s1 = s1.toLowerCase(); s2 = s2.toLowerCase();
    let costs = new Array();
    for (let i = 0; i <= s1.length; i++) {
        let lastValue = i;
        for (let j = 0; j <= s2.length; j++) {
            if (i == 0) costs[j] = j;
            else {
                if (j > 0) {
                    let newValue = costs[j - 1];
                    if (s1.charAt(i - 1) != s2.charAt(j - 1))
                        newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
                    costs[j - 1] = lastValue;
                    lastValue = newValue;
                }
            }
        }
        if (i > 0) costs[s2.length] = lastValue;
    }
    return costs[s2.length];
}

function solveWord(rawWord) {
    let clean = normalizeMenuString(rawWord);
    let lower = clean.toLowerCase();
    if (forcedCorrections[lower]) return forcedCorrections[lower];
    if (!isNaN(clean) || clean.length < 3) return clean;
    for (let word of dictionary) {
        if (word.toLowerCase() === lower) return word;
    }
    let bestMatch = clean;
    let highestScore = 0;
    for (let correct of dictionary) {
        let score = getSimilarity(clean, correct);
        if (score > highestScore) {
            highestScore = score;
            bestMatch = correct;
        }
    }
    if (highestScore > 0.80) return bestMatch;
    return clean;
}

async function processMenuFixer(files) {
    const file = files[0];
    const buffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const worksheet = workbook.worksheets[0];
    let totalChanges = 0;

    worksheet.eachRow((row) => {
        row.eachCell((cell) => {
            if (cell.type === ExcelJS.ValueType.String || typeof cell.value === 'string') {
                let originalStr = cell.value.toString();
                let spacingFixed = originalStr.replace(/\s\s+/g, ' ').trim();
                let words = spacingFixed.split(' ');
                let richText = [];
                let hasChange = false;

                words.forEach((word, index) => {
                    let fixed = solveWord(word);
                    let space = (index < words.length - 1) ? ' ' : '';
                    if (fixed !== word) {
                        richText.push({ text: fixed + space, font: { bold: true, color: { argb: 'FFFF0000' }, name: 'Calibri' } });
                        hasChange = true;
                        totalChanges++;
                    } else {
                        richText.push({ text: fixed + space, font: { color: { argb: 'FF000000' }, name: 'Calibri' } });
                    }
                });

                if (hasChange) {
                    cell.value = { richText: richText };
                } else if (originalStr !== spacingFixed) {
                    cell.value = spacingFixed;
                }
            }
        });
    });

    const outBuffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([outBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    document.getElementById('success-details').innerText = `Fixed ${totalChanges} typos!`;
    setupDownload(blob, "Menu_Fixed.xlsx");
}

/* --- DUPLICATE REMOVER LOGIC --- */
function handleDuplicateFile(file) {
    activeDupFileName = file.name.split('.')[0];
    const reader = new FileReader();
    document.getElementById('drop-area').classList.add('hidden');
    document.getElementById('status-area').classList.remove('hidden');
    document.getElementById('loading-spinner').classList.remove('hidden');

    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheet];
        rawDupData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        document.getElementById('loading-spinner').classList.add('hidden');
        
        if(rawDupData.length > 0) {
            document.getElementById('dup-options-area').classList.remove('hidden');
        } else {
            document.getElementById('error-msg').innerText = "File is empty.";
            document.getElementById('error-msg').classList.remove('hidden');
        }
    };
    reader.readAsArrayBuffer(file);
}

function processDuplicateData(mode) {
    if(!rawDupData || rawDupData.length === 0) return;
    const criteria = document.querySelector('input[name="dup-criteria"]:checked').value;
    const seen = new Set();
    const duplicatesMap = new Map();
    rawDupData.forEach((row, index) => {
        let fingerprint = "";
        if (criteria === 'col1') fingerprint = row[0] ? String(row[0]).trim().toLowerCase() : ""; 
        else fingerprint = JSON.stringify(row);
        if (fingerprint === "") return; 
        if (seen.has(fingerprint)) duplicatesMap.set(index, true);
        else seen.add(fingerprint);
    });
    const finalRows = [];
    rawDupData.forEach((row, index) => {
        const isDuplicate = duplicatesMap.has(index);
        if (mode === 'remove') {
            if (!isDuplicate) finalRows.push(row);
        } else {
            const styledRow = row.map(cellValue => {
                let cellObj = { v: cellValue, t: 's' };
                if(typeof cellValue === 'number') cellObj.t = 'n';
                if (isDuplicate) {
                    cellObj.s = { fill: { fgColor: { rgb: "FFFF00" } }, font: { color: { rgb: "DC2626" }, bold: true } };
                }
                return cellObj;
            });
            finalRows.push(styledRow);
        }
    });
    const ws = XLSX.utils.aoa_to_sheet(finalRows);
    const colWidths = finalRows[0] ? finalRows[0].map(() => ({ wch: 25 })) : [];
    ws['!cols'] = colWidths;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Processed_Data");
    const suffix = mode === 'remove' ? "_Cleaned" : "_Highlighted";
    XLSX.writeFile(wb, `${activeDupFileName}${suffix}.xlsx`);
    document.getElementById('dup-options-area').classList.add('hidden');
    document.getElementById('success-msg').classList.remove('hidden');
    document.getElementById('download-btn').style.display = 'none';
}

/* --- CONVERSION TOOLS --- */
async function convertJpgToPdf(files) {
    const doc = new window.jspdf.jsPDF();
    let processedCount = 0;
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        if (!file.type.startsWith('image/')) continue;
        const imageData = await readImageDimensions(file);
        const pdfWidth = doc.internal.pageSize.getWidth();
        const pdfHeight = doc.internal.pageSize.getHeight();
        const ratio = Math.min(pdfWidth / imageData.width, pdfHeight / imageData.height);
        const imgWidth = imageData.width * ratio;
        const imgHeight = imageData.height * ratio;
        const x = (pdfWidth - imgWidth) / 2;
        const y = (pdfHeight - imgHeight) / 2;
        if (processedCount > 0) doc.addPage();
        doc.addImage(imageData.img, 'JPEG', x, y, imgWidth, imgHeight);
        processedCount++;
    }
    if (processedCount === 0) throw new Error("No valid images selected.");
    setupDownload(doc.output('blob'), "combined-images.pdf");
}

function readImageDimensions(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const img = new Image();
            img.src = e.target.result;
            img.onload = () => resolve({ img, width: img.width, height: img.height });
            img.onerror = reject;
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

async function compressPDF(files) {
    const file = files[0];
    const originalSize = file.size;
    const quality = document.getElementById('quality-slider').value / 100;
    const data = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument(data).promise;
    const doc = new window.jspdf.jsPDF();
    
    document.getElementById('progress-bar').classList.remove('hidden');
    document.getElementById('total-pages').innerText = pdf.numPages;
    
    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const viewport = page.getViewport({ scale: 1.5 });  // Higher scale for better text quality
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        await page.render({ canvasContext: context, viewport: viewport }).promise;
        const imgData = canvas.toDataURL('image/jpeg', quality);
        if (i > 1) doc.addPage([viewport.width, viewport.height]);
        else { doc.internal.pageSize.width = viewport.width; doc.internal.pageSize.height = viewport.height; }
        doc.addImage(imgData, 'JPEG', 0, 0, viewport.width, viewport.height);
        
        document.getElementById('current-page').innerText = i;
        document.getElementById('progress-fill').style.width = (i / pdf.numPages * 100) + '%';
    }
    
    document.getElementById('progress-bar').classList.add('hidden');
    const compressedBlob = doc.output('blob');
    const compressedSize = compressedBlob.size;
    const savings = ((originalSize - compressedSize) / originalSize * 100).toFixed(1);
    document.getElementById('success-details').innerText = `Compressed from ${(originalSize / 1024 / 1024).toFixed(2)} MB to ${(compressedSize / 1024 / 1024).toFixed(2)} MB (${savings}% reduction)!`;
    setupDownload(compressedBlob, "compressed.pdf");
}

async function cleanExcelFile(files) {
    const file = files[0];
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        if(!sheet['!ref']) return;
        const range = XLSX.utils.decode_range(sheet['!ref']);
        for(let R = range.s.r; R <= range.e.r; ++R) {
            for(let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({r:R, c:C});
                const cell = sheet[cellAddress];
                if(cell && cell.t === 's') { cell.v = cleanString(cell.v); }
            }
        }
    });
    XLSX.writeFile(workbook, "Cleaned_" + file.name);
}

function cleanString(str) {
    if (!str) return str;
    let cleaned = str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    cleaned = cleaned.replace(/–/g, "-").replace(/—/g, "-").replace(/[^\x00-\x7F]/g, "");
    return cleaned;
}

async function convertWordToPdf(files) {
    const file = files[0];
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
    const pdfBlob = await html2pdf().from(result.value).output('blob');
    setupDownload(pdfBlob, "converted.pdf");
}

async function convertExcelToPdf(files) {
    const file = files[0];
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const htmlTable = XLSX.utils.sheet_to_html(firstSheet);
    const pdfBlob = await html2pdf().from(htmlTable).output('blob');
    setupDownload(pdfBlob, "converted.pdf");
}

async function convertPdfToJpg(files) {
    const file = files[0];
    const data = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument(data).promise;
    const zip = new JSZip();
    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const viewport = page.getViewport({ scale: 1.5 });
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.height = viewport.height;
        canvas.width = viewport.width;
        await page.render({ canvasContext: context, viewport: viewport }).promise;
        zip.file(`page-${i}.jpg`, canvas.toDataURL('image/jpeg').split(',')[1], {base64: true});
    }
    setupDownload(await zip.generateAsync({type:"blob"}), "pages.zip");
}

async function convertPdfImageToExcel(files) {
    const file = files[0];
    let extractedText = "";
    if (file.type === "application/pdf") {
        const data = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument(data).promise;
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            extractedText += textContent.items.map(i => i.str).join(' ') + "\n";
        }
    } else if (file.type.startsWith("image/")) {
        const worker = Tesseract.createWorker();
        await worker.load(); await worker.loadLanguage('eng'); await worker.initialize('eng');
        const { data: { text } } = await worker.recognize(file);
        extractedText = text;
        await worker.terminate();
    }
    const rows = extractedText.split('\n').map(line => [line]);
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    XLSX.writeFile(workbook, "extracted_data.xlsx");
}

/* --- DOWNLOAD HELPER --- */
function setupDownload(blob, filename) {
    const btn = document.getElementById('download-btn');
    document.getElementById('download-btn').style.display = 'inline-block';
    btn.onclick = () => saveAs(blob, filename);
}

// Initialize slider listener
document.addEventListener('DOMContentLoaded', () => {
    const slider = document.getElementById('quality-slider');
    if (slider) {
        slider.addEventListener('input', (e) => {
            document.getElementById('quality-value').innerText = e.target.value + '%';
        });
    }
});