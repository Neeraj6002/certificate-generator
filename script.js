// Core application state
let templateImg = new Image();
let excelData = [];
let positions = {};
let textStyles = {};
let selectedFields = new Set();
let hoverField = null;

// DOM elements
const canvas = document.getElementById('preview');
const ctx = canvas.getContext('2d');

// UI state
let isDragging = false;
let selectedField = null;
let offsetX, offsetY;
let MOVE_STEP = 5;

// Google Fonts list (exclude system fonts). These will be loaded on demand.
let googleFonts = [
    'Inter', 'Roboto', 'Open Sans', 'Lato', 'Montserrat', 'Merriweather',
    'Oswald', 'Raleway', 'Source Sans Pro', 'Playfair Display', 'Ubuntu',
    'Nunito', 'Pacifico', 'Lobster', 'Bree Serif', 'Poppins', 'Quicksand',
    'Dancing Script', 'Crimson Text', 'Fira Sans', 'Work Sans', 'PT Sans',
    'Libre Baskerville', 'Cabin', 'Dosis'
];

// Cache of loaded fonts
const loadedFonts = new Set();
// Font loading preview animation state
let fontPreviewLoading = false;
let _fontPreviewRAF = null;
let _fontPreviewAngle = 0;
// Persisted data and removed fonts
let persistedState = {};
let removedFonts = new Set();

// Debounced save
let _saveTimeout = null;
function scheduleSave(delay = 1000) {
    if (_saveTimeout) clearTimeout(_saveTimeout);
    _saveTimeout = setTimeout(() => {
        saveState();
    }, delay);
}

function saveState() {
    try {
        const state = {
            textStyles,
            positions
        };
        localStorage.setItem('certgen_state', JSON.stringify(state));
        localStorage.setItem('certgen_removed_fonts', JSON.stringify(Array.from(removedFonts)));
    } catch (e) {
        console.warn('Failed to save state', e);
    }
}

function loadState() {
    try {
        const s = localStorage.getItem('certgen_state');
        if (s) persistedState = JSON.parse(s) || {};
        const rf = localStorage.getItem('certgen_removed_fonts');
        if (rf) {
            const arr = JSON.parse(rf) || [];
            removedFonts = new Set(arr);
            // remove from googleFonts list
            googleFonts = googleFonts.filter(f => !removedFonts.has(f));
        }
    } catch (e) {
        console.warn('Failed to load persisted state', e);
    }
}

// Helper: timeout wrapper for promises
function promiseTimeout(promise, ms) {
    let t = null;
    const timeout = new Promise((_, reject) => {
        t = setTimeout(() => reject(new Error('Timeout')), ms);
    });
    return Promise.race([promise, timeout]).then(res => { clearTimeout(t); return res; });
}

// Small helper to dynamically load a script and return a promise that resolves when loaded
function loadScript(src, attrs = {}) {
    return new Promise((resolve, reject) => {
        try {
            const s = document.createElement('script');
            s.src = src;
            s.async = true;
            Object.keys(attrs).forEach(k => s.setAttribute(k, attrs[k]));
            s.onload = () => resolve();
            s.onerror = (e) => reject(new Error('Failed to load ' + src));
            document.head.appendChild(s);
        } catch (err) {
            reject(err);
        }
    });
}

// Ensure JSZip is available. If not, dynamically load a known CDN copy.
async function ensureJSZip() {
    if (typeof window.JSZip !== 'undefined') return;
    // Try a reliable CDN (jsDelivr). No SRI to avoid blocking if hashes mismatch.
    const cdn = 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js';
    try {
        await loadScript(cdn);
        // wait a tick for global to be available
        await new Promise(r => setTimeout(r, 50));
        if (typeof window.JSZip === 'undefined') throw new Error('JSZip not available after dynamic load');
    } catch (e) {
        console.error('Failed to load JSZip dynamically:', e);
        throw e;
    }
}

// Verify fonts at startup and remove ones that fail to load
async function verifyAndTrimFonts() {
    const keep = [];
    for (const f of googleFonts) {
        try {
            await promiseTimeout(loadFontFamily(f), 4000);
            keep.push(f);
        } catch (e) {
            removedFonts.add(f);
        }
    }
    googleFonts = keep;
    saveState();
}

// Eagerly load all fonts on startup, but with controlled concurrency
async function loadAllFontsOnStartup() {
    const fontSelect = document.getElementById('fontSelect');
    const fontSearch = document.getElementById('fontSearch');
    if (fontSelect) {
        fontSelect.disabled = true;
        fontSelect.innerHTML = '';
        const opt = document.createElement('option');
        opt.textContent = 'Loading fonts...';
        fontSelect.appendChild(opt);
    }

    const concurrency = 4;
    const loaded = [];

    // Process fonts in batches for proper concurrency control
    const fontsToLoad = [...googleFonts];

    while (fontsToLoad.length > 0) {
        // Take a batch of fonts to load concurrently
        const batch = fontsToLoad.splice(0, concurrency);

        // Create promises for the batch
        const batchPromises = batch.map(async (f) => {
            try {
                await promiseTimeout(loadFontFamily(f), 6000);
                return { font: f, success: true };
            } catch (e) {
                return { font: f, success: false };
            }
        });

        // Wait for entire batch to complete before starting next
        const results = await Promise.allSettled(batchPromises);

        // Process results
        results.forEach(result => {
            if (result.status === 'fulfilled' && result.value.success) {
                loaded.push(result.value.font);
            } else if (result.status === 'fulfilled' && !result.value.success) {
                removedFonts.add(result.value.font);
            }
        });
    }

    // Update googleFonts to only those loaded
    googleFonts = loaded;
    saveState();

    if (fontSelect) {
        loadGoogleFonts();
        fontSelect.disabled = false;
        if (fontSearch) fontSearch.disabled = false;
    }
}

// Initialize application when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
    setupEventListeners();
    loadState();
    // Eagerly load all fonts from the googleFonts list, then populate the picker
    loadAllFontsOnStartup().then(() => {
        loadGoogleFonts();
    }).catch(() => {
        loadGoogleFonts();
    });
    checkFirstVisit();
    initializeTheme();
});

function initializeApp() {
    canvas.style.display = 'none';
    document.getElementById('canvasPlaceholder').style.display = 'flex';
    updateMoveStep();
}

function setupEventListeners() {
    // File uploads
    document.getElementById('excelInput').addEventListener('change', handleExcelUpload);
    document.getElementById('templateInput').addEventListener('change', handleTemplateUpload);
    
    // Style controls
    document.getElementById('fieldSelect').addEventListener('change', handleFieldSelect);
    document.getElementById('fontSearch').addEventListener('input', handleFontSearch);
    document.getElementById('fontSelect').addEventListener('change', handleFontChange);
    document.getElementById('fontSize').addEventListener('change', handleFontSizeChange);
    document.getElementById('fontSizeRange').addEventListener('input', handleFontSizeRangeChange);
    document.getElementById('textColor').addEventListener('change', handleColorChange);
    
    // Alignment buttons
    document.querySelectorAll('.align-btn').forEach(btn => {
        btn.addEventListener('click', handleAlignmentChange);
    });
    
    // Canvas interactions
    canvas.addEventListener('mousedown', handleMouseDown);
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
    canvas.addEventListener('mousemove', handleCanvasHover);
    canvas.addEventListener('mouseleave', () => {
        canvas.style.cursor = 'default';
        hoverField = null;
    });
    
    // Button events
    document.getElementById('generateBtn').addEventListener('click', generateCertificates);
    document.getElementById('resetAllBtn').addEventListener('click', resetAllPositions);
    
    // Theme controls
    document.getElementById('themeToggle').addEventListener('click', toggleTheme);
    document.querySelectorAll('.theme-option').forEach(option => {
        option.addEventListener('click', () => {
            setTheme(option.dataset.theme);
        });
    });
    
  /*   // Patch notes
    document.getElementById('patchNotesBtn').addEventListener('click', showPatchNotes);
    document.getElementById('closePatchNotes').addEventListener('click', closePatchNotes);
    document.getElementById('gotItBtn').addEventListener('click', closePatchNotes);
    document.getElementById('patchNotesOverlay').addEventListener('click', function(e) {
        if (e.target === this) closePatchNotes();
    });
    
    // Keyboard shortcuts
    document.addEventListener('keydown', function(e) {
        // Escape key - close modals and dropdowns
        if (e.key === 'Escape') {
            closePatchNotes();
            document.getElementById('themeDropdown').classList.remove('show');
        }

        // Only handle shortcuts when not in an input field
        const isInputFocused = ['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement.tagName);
        if (isInputFocused) return;

        // Ctrl/Cmd + G - Generate certificates
        if ((e.ctrlKey || e.metaKey) && e.key === 'g') {
            e.preventDefault();
            generateCertificates();
        }

        // Ctrl/Cmd + R - Reset all positions (when not refreshing page)
        if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === 'R') {
            e.preventDefault();
            resetAllPositions();
        }

        // Arrow keys - move selected field
        if (selectedField && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
            e.preventDefault();
            const direction = e.key.replace('Arrow', '').toLowerCase();
            moveText(direction);
        }
    }); */

    // Prevent accidental page close when data is loaded
    window.addEventListener('beforeunload', function(e) {
        if (excelData.length > 0 || templateImg.src) {
            e.preventDefault();
            e.returnValue = '';
            return '';
        }
    });
    
    // Close theme dropdown when clicking outside
    document.addEventListener('click', function(e) {
        const themeDropdown = document.getElementById('themeDropdown');
        const themeBtn = document.getElementById('themeToggle');
        
        if (!themeBtn.contains(e.target) && !themeDropdown.contains(e.target)) {
            themeDropdown.classList.remove('show');
        }
    });
    
    // Movement buttons removed (position controls no longer present)
}

// File handling
function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const label = e.target.previousElementSibling;
    label.innerHTML = '<i class="fas fa-spinner fa-spin"></i><span>Loading...</span>';

    // Validate file type
    const validExtensions = ['.xlsx', '.xls'];
    const fileName = file.name.toLowerCase();
    const hasValidExtension = validExtensions.some(ext => fileName.endsWith(ext));

    if (!hasValidExtension) {
        label.innerHTML = '<i class="fas fa-exclamation-triangle"></i><span>Invalid file type</span><small>Please use .xlsx or .xls</small>';
        showNotification('Please upload a valid Excel file (.xlsx or .xls)', 'error');
        return;
    }

    // Validate file size (max 10MB)
    const maxSize = 10 * 1024 * 1024; // 10MB
    if (file.size > maxSize) {
        label.innerHTML = '<i class="fas fa-exclamation-triangle"></i><span>File too large</span><small>Max 10MB allowed</small>';
        showNotification('File is too large. Please use a file smaller than 10MB.', 'error');
        return;
    }

    // Guard: ensure SheetJS (XLSX) library is available
    if (typeof XLSX === 'undefined') {
        console.error('XLSX library is not loaded. Make sure the SheetJS script is included before script.js');
        label.innerHTML = '<i class="fas fa-exclamation-triangle"></i><span>Library missing</span><small>Please refresh the page</small>';
        showNotification('Excel library failed to load. Please refresh the page.', 'error');
        return;
    }

    const reader = new FileReader();

    reader.onerror = function() {
        console.error('FileReader error:', reader.error);
        label.innerHTML = '<i class="fas fa-exclamation-triangle"></i><span>Read error</span><small>Please try again</small>';
        showNotification('Failed to read the file. Please try again.', 'error');
    };

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Validate workbook has sheets
            if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                throw new Error('No sheets found in the Excel file');
            }

            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            excelData = XLSX.utils.sheet_to_json(sheet);

            if (excelData.length === 0) {
                throw new Error('Excel file is empty or has no data rows');
            }

            const headings = Object.keys(excelData[0]);

            if (headings.length === 0) {
                throw new Error('No columns found in the Excel file');
            }

            // Warn if very large dataset
            if (excelData.length > 1000) {
                showNotification(`Loading ${excelData.length} records. Generation may take a while.`, 'warning');
            }

            positions = {};
            textStyles = {};
            selectedFields.clear();

            headings.forEach((field, index) => {
                positions[field] = { x: 100, y: 100 + (index * 60) };
                textStyles[field] = {
                    font: 'Inter',
                    size: 24,
                    align: 'left',
                    color: '#000000'
                };
                selectedFields.add(field);
            });

            // persist initial layout
            scheduleSave(500);

            populateFieldOptions(headings);
            updatePreview();

            label.classList.add('success');
            label.innerHTML = '<i class="fas fa-check"></i><span>Excel Loaded</span><small>(' + excelData.length + ' records, ' + headings.length + ' fields)</small>';
            showNotification(`Loaded ${excelData.length} records with ${headings.length} fields`, 'success');

        } catch (error) {
            console.error('Error reading Excel file:', error);
            label.innerHTML = '<i class="fas fa-exclamation-triangle"></i><span>Error loading file</span><small>' + (error.message || 'Please try again') + '</small>';
            showNotification(error.message || 'Failed to parse Excel file. Please check the file format.', 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

function handleTemplateUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const label = e.target.previousElementSibling;
    label.innerHTML = '<i class="fas fa-spinner fa-spin"></i><span>Loading...</span>';

    // Validate file is an image
    if (!file.type.startsWith('image/')) {
        label.innerHTML = '<i class="fas fa-exclamation-triangle"></i><span>Invalid file type</span><small>Please use an image file</small>';
        showNotification('Please upload a valid image file (PNG, JPG, etc.)', 'error');
        return;
    }

    // Validate file size (max 20MB for images)
    const maxSize = 20 * 1024 * 1024; // 20MB
    if (file.size > maxSize) {
        label.innerHTML = '<i class="fas fa-exclamation-triangle"></i><span>File too large</span><small>Max 20MB allowed</small>';
        showNotification('Image is too large. Please use a file smaller than 20MB.', 'error');
        return;
    }

    const reader = new FileReader();

    reader.onerror = function() {
        console.error('FileReader error:', reader.error);
        label.innerHTML = '<i class="fas fa-exclamation-triangle"></i><span>Read error</span><small>Please try again</small>';
        showNotification('Failed to read the image file. Please try again.', 'error');
    };

    reader.onload = function(e) {
        templateImg.src = e.target.result;

        templateImg.onerror = function() {
            label.innerHTML = '<i class="fas fa-exclamation-triangle"></i><span>Invalid image</span><small>Could not load image</small>';
            showNotification('Failed to load the image. Please try a different file.', 'error');
        };

        templateImg.onload = function() {
            // Validate image dimensions
            if (templateImg.width < 100 || templateImg.height < 100) {
                label.innerHTML = '<i class="fas fa-exclamation-triangle"></i><span>Image too small</span><small>Min 100x100 required</small>';
                showNotification('Image is too small. Please use an image at least 100x100 pixels.', 'error');
                return;
            }

            const maxWidth = 800;
            const maxHeight = 600;
            let { width, height } = templateImg;

            if (width > maxWidth || height > maxHeight) {
                const ratio = Math.min(maxWidth / width, maxHeight / height);
                width *= ratio;
                height *= ratio;
            }

            canvas.width = width;
            canvas.height = height;

            canvas.style.display = 'block';
            document.getElementById('canvasPlaceholder').style.display = 'none';

            updatePreview();

            label.classList.add('success');
            label.innerHTML = '<i class="fas fa-check"></i><span>Template Loaded</span><small>(' + templateImg.width + 'x' + templateImg.height + ')</small>';
            showNotification('Template loaded successfully', 'success');
        };
    };
    reader.readAsDataURL(file);
}

// Field management
function populateFieldOptions(headings) {
    const fieldSelect = document.getElementById('fieldSelect');
    const fieldCheckboxes = document.getElementById('fieldCheckboxes');
    
    fieldSelect.innerHTML = '<option value="">Select a field to edit</option>';
    fieldCheckboxes.innerHTML = '';

    headings.forEach(field => {
        const option = document.createElement('option');
        option.value = field;
        option.textContent = field;
        fieldSelect.appendChild(option);

        const label = document.createElement('label');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = true;
        checkbox.value = field;
        checkbox.addEventListener('change', function() {
            if (this.checked) {
                selectedFields.add(field);
            } else {
                selectedFields.delete(field);
            }
            updatePreview();
        });
        
        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(field));
        fieldCheckboxes.appendChild(label);
    });

    if (headings.length > 0) {
        selectedField = headings[0];
        fieldSelect.value = selectedField;
        updateStyleControls();
    }
}

// Style controls
function handleFieldSelect() {
    selectedField = document.getElementById('fieldSelect').value;
    updateStyleControls();
    updatePreview();
    scheduleSave();
}

function handleFontSearch(e) {
    loadGoogleFonts(e.target.value);
}

function handleFontChange() {
    if (!selectedField) return;

    const selectedFont = document.getElementById('fontSelect').value;
    // Show loading UI
    showFontLoading(selectedFont);
    // Attempt to load the font from Google Fonts (or cached)
    loadFontFamily(selectedFont).then(() => {
        textStyles[selectedField].font = selectedFont;
        // Force a preview update after the font is available so canvas renders immediately
        updatePreview();
        // A short timeout to ensure the browser's font rendering has caught up
        setTimeout(updatePreview, 50);
        scheduleSave();
    }).catch(err => {
        console.warn(err);
        // Remove failing font from picker so users won't repeatedly choose it
        googleFonts = googleFonts.filter(f => f !== selectedFont);
        removedFonts.add(selectedFont);
        scheduleSave();
        // Fallback to Inter
        textStyles[selectedField].font = 'Inter';
        updatePreview();
    }).finally(() => {
        hideFontLoading();
    });
}

function handleFontSizeChange() {
    if (!selectedField) return;

    let size = parseInt(document.getElementById('fontSize').value);
    // Clamp value to valid range
    size = Math.max(8, Math.min(200, size || 16));
    textStyles[selectedField].size = size;
    document.getElementById('fontSizeRange').value = Math.min(size, 72); // Range max is 72
    document.getElementById('fontSize').value = size;
    updatePreview();
    scheduleSave();
}

function handleFontSizeRangeChange() {
    if (!selectedField) return;

    const size = parseInt(document.getElementById('fontSizeRange').value);
    textStyles[selectedField].size = size;
    document.getElementById('fontSize').value = size;
    updatePreview();
    scheduleSave();
}

function handleColorChange() {
    if (!selectedField) return;

    const color = document.getElementById('textColor').value;
    textStyles[selectedField].color = color;
    document.getElementById('colorValue').textContent = color;
    updatePreview();
    scheduleSave();
}

function handleAlignmentChange(e) {
    if (!selectedField) return;

    document.querySelectorAll('.align-btn').forEach(btn => btn.classList.remove('active'));
    e.target.closest('.align-btn').classList.add('active');

    const alignment = e.target.closest('.align-btn').dataset.align;
    textStyles[selectedField].align = alignment;
    document.getElementById('textAlign').value = alignment;
    updatePreview();
    scheduleSave();
}

function updateMoveStep() {
    const el = document.getElementById('moveStep');
    if (el) {
        const v = parseInt(el.value);
        if (!isNaN(v)) MOVE_STEP = v;
    }
}

// Canvas interactions
function handleMouseDown(e) {
    const rect = canvas.getBoundingClientRect();
    const x = (e.clientX - rect.left) * (canvas.width / rect.width);
    const y = (e.clientY - rect.top) * (canvas.height / rect.height);
    
    for (let field of selectedFields) {
        const pos = positions[field];
        const styles = textStyles[field];
        const sampleData = excelData.length > 0 ? excelData[0] : {};
        const text = sampleData[field] || field;
        
        ctx.font = `${styles.size}px "${styles.font}"`;
        const metrics = ctx.measureText(text);
        const textWidth = metrics.width;
        const textHeight = styles.size;
        
        let textX = pos.x;
        if (styles.align === 'center') textX -= textWidth / 2;
        else if (styles.align === 'right') textX -= textWidth;
        
        const padding = 10;
        if (x >= textX - padding && 
            x <= textX + textWidth + padding && 
            y >= pos.y - textHeight - padding && 
            y <= pos.y + padding) {
            
            isDragging = true;
            selectedField = field;
            offsetX = x - pos.x;
            offsetY = y - pos.y;
            
            document.getElementById('fieldSelect').value = field;
            updateStyleControls();
            canvas.style.cursor = 'grabbing';
            break;
        }
    }
}

function handleMouseMove(e) {
    if (!isDragging || !selectedField) return;
    
    const rect = canvas.getBoundingClientRect();
    const x = (e.clientX - rect.left) * (canvas.width / rect.width);
    const y = (e.clientY - rect.top) * (canvas.height / rect.height);

    positions[selectedField] = {
        x: Math.max(0, Math.min(x - offsetX, canvas.width)),
        y: Math.max(20, Math.min(y - offsetY, canvas.height - 20))
    };
    
    updatePreview();
}

function handleMouseUp() {
    if (isDragging) {
        isDragging = false;
        canvas.style.cursor = 'grab';
        // Save position after drag operation completes
        scheduleSave();
    }
}

function handleCanvasHover(e) {
    if (isDragging) return;
    
    const rect = canvas.getBoundingClientRect();
    const x = (e.clientX - rect.left) * (canvas.width / rect.width);
    const y = (e.clientY - rect.top) * (canvas.height / rect.height);
    
    let foundField = null;
    
    for (let field of selectedFields) {
        const pos = positions[field];
        const styles = textStyles[field];
        const sampleData = excelData.length > 0 ? excelData[0] : {};
        const text = sampleData[field] || field;
        
        ctx.font = `${styles.size}px "${styles.font}"`;
        const metrics = ctx.measureText(text);
        const textWidth = metrics.width;
        const textHeight = styles.size;
        
        let textX = pos.x;
        if (styles.align === 'center') textX -= textWidth / 2;
        else if (styles.align === 'right') textX -= textWidth;
        
        const padding = 10;
        if (x >= textX - padding && 
            x <= textX + textWidth + padding && 
            y >= pos.y - textHeight - padding && 
            y <= pos.y + padding) {
            foundField = field;
            break;
        }
    }
    
    canvas.style.cursor = foundField ? 'grab' : 'default';
    hoverField = foundField;
}

// Position controls
function moveText(direction) {
    if (!selectedField) {
        showNotification('Please select a field first!', 'warning');
        return;
    }

    const pos = positions[selectedField];
    switch (direction) {
        case 'up':
            pos.y = Math.max(20, pos.y - MOVE_STEP);
            break;
        case 'down':
            pos.y = Math.min(canvas.height - 20, pos.y + MOVE_STEP);
            break;
        case 'left':
            pos.x = Math.max(0, pos.x - MOVE_STEP);
            break;
        case 'right':
            pos.x = Math.min(canvas.width, pos.x + MOVE_STEP);
            break;
    }
    
    updatePreview();
}

function resetPosition() {
    const headings = Object.keys(positions);
    const idx = headings.indexOf(selectedField);
    positions[selectedField] = { x: 100, y: 100 + (idx * 60) };
    updatePreview();
    scheduleSave();
}

function resetAllPositions() {
    const headings = Object.keys(positions);
    if (headings.length === 0) {
        showNotification('No fields to reset. Please upload an Excel file first.', 'warning');
        return;
    }

    // Show confirmation dialog
    if (!confirm('Are you sure you want to reset all field positions to default?')) {
        return;
    }

    headings.forEach((field, index) => {
        positions[field] = { x: 100, y: 100 + (index * 60) };
    });
    updatePreview();
    scheduleSave();
    showNotification('All positions have been reset to default', 'info');
}

// Font management
function loadGoogleFonts(searchTerm = '') {
    const fontSelect = document.getElementById('fontSelect');
    fontSelect.innerHTML = '';

    const filteredFonts = googleFonts.filter(font =>
        font.toLowerCase().includes(searchTerm.toLowerCase())
    );

    filteredFonts.forEach(font => {
        const option = document.createElement('option');
        option.value = font;
        option.textContent = font;
        fontSelect.appendChild(option);
    });

    if (selectedField && textStyles[selectedField]) {
        fontSelect.value = textStyles[selectedField].font || 'Inter';
    }
}

// Lazy-load a Google font using WebFont loader (or fallback to a link tag)
function loadFontFamily(fontFamily) {
    return new Promise((resolve, reject) => {
        // If already loaded, ensure it's available to the Font Loading API
        if (loadedFonts.has(fontFamily)) {
            if (window.document && document.fonts && document.fonts.load) {
                document.fonts.load(`1em "${fontFamily}"`).then(() => resolve()).catch(() => resolve());
            } else {
                return resolve();
            }
            return;
        }

        const finishResolve = () => {
            // Try to ensure the face is usable by the canvas by waiting on document.fonts if available
            if (window.document && document.fonts && document.fonts.load) {
                // request a lightweight load and then resolve regardless of success after a short timeout
                const p = document.fonts.load(`1em "${fontFamily}"`);
                const to = setTimeout(() => {
                    loadedFonts.add(fontFamily);
                    resolve();
                }, 2000);
                p.then(() => { clearTimeout(to); loadedFonts.add(fontFamily); resolve(); }).catch(() => { clearTimeout(to); loadedFonts.add(fontFamily); resolve(); });
            } else {
                loadedFonts.add(fontFamily);
                resolve();
            }
        };

        if (typeof WebFont !== 'undefined') {
            WebFont.load({
                google: { families: [fontFamily] },
                active: function() {
                    finishResolve();
                },
                inactive: function() {
                    reject(new Error('WebFont failed to load ' + fontFamily));
                }
            });
            return;
        }

        // Fallback: inject a link tag for the font
        const link = document.createElement('link');
        link.rel = 'stylesheet';
        link.href = `https://fonts.googleapis.com/css2?family=${fontFamily.replace(/ /g, '+')}:wght@400;700&display=swap`;
        link.onload = function() { finishResolve(); };
        link.onerror = function() { reject(new Error('Link failed to load ' + fontFamily)); };
        document.head.appendChild(link);
    });
}

// UI helpers: show a tiny loading state in the font select while a font is loading
function showFontLoading(fontFamily) {
    const fontSelect = document.getElementById('fontSelect');
    const fontSearch = document.getElementById('fontSearch');
    if (!fontSelect) return;
    fontSelect.disabled = true;
    if (fontSearch) fontSearch.disabled = true;
    // store current selection in data attr
    fontSelect.dataset.prev = fontSelect.value || '';

    // create or show spinner element next to select
    let spinner = document.getElementById('fontLoadingSpinner');
    if (!spinner) {
        spinner = document.createElement('span');
        spinner.id = 'fontLoadingSpinner';
        spinner.className = 'font-spinner';
        spinner.setAttribute('aria-hidden', 'true');
        fontSelect.insertAdjacentElement('afterend', spinner);
    }
    spinner.style.display = 'inline-block';
    // Start canvas preview animation
    startFontPreviewLoadingAnimation();
}

function hideFontLoading() {
    const fontSelect = document.getElementById('fontSelect');
    const fontSearch = document.getElementById('fontSearch');
    if (!fontSelect) return;
    fontSelect.disabled = false;
    if (fontSearch) fontSearch.disabled = false;
    // restore selection if possible
    const prev = fontSelect.dataset.prev;
    if (prev) fontSelect.value = prev;
    delete fontSelect.dataset.prev;
    const spinner = document.getElementById('fontLoadingSpinner');
    if (spinner) spinner.style.display = 'none';
    // Stop canvas preview animation
    stopFontPreviewLoadingAnimation();
}

function startFontPreviewLoadingAnimation() {
    // Always cancel any existing animation frame first to prevent leaks
    if (_fontPreviewRAF) {
        cancelAnimationFrame(_fontPreviewRAF);
        _fontPreviewRAF = null;
    }

    if (fontPreviewLoading) return;
    fontPreviewLoading = true;
    _fontPreviewAngle = 0;

    function frame() {
        // Safety check - stop if loading was cancelled
        if (!fontPreviewLoading) {
            _fontPreviewRAF = null;
            return;
        }
        _fontPreviewAngle += 0.12;
        updatePreview();
        _fontPreviewRAF = requestAnimationFrame(frame);
    }
    _fontPreviewRAF = requestAnimationFrame(frame);
}

function stopFontPreviewLoadingAnimation() {
    fontPreviewLoading = false;
    if (_fontPreviewRAF) {
        cancelAnimationFrame(_fontPreviewRAF);
        _fontPreviewRAF = null;
    }
    _fontPreviewAngle = 0;
    updatePreview();
}

// Preload a small subset of trending fonts in the background for faster UX
function preloadCommonFonts() {
    const preloadFonts = ['Inter', 'Poppins', 'Roboto', 'Montserrat'];
    // Load them in sequence to avoid hammering network
    (async () => {
        for (const f of preloadFonts) {
            try {
                await loadFontFamily(f);
                // console.info('Preloaded font', f);
            } catch (e) {
                // ignore preload failures
            }
        }
    })();
}

// UI updates
function updateStyleControls() {
    if (!selectedField || !textStyles[selectedField]) return;
    
    const styles = textStyles[selectedField];
    
    document.getElementById('fontSelect').value = styles.font || 'Inter';
    document.getElementById('fontSize').value = styles.size;
    document.getElementById('fontSizeRange').value = styles.size;
    document.getElementById('textAlign').value = styles.align;
    document.getElementById('textColor').value = styles.color;
    document.getElementById('colorValue').textContent = styles.color;
    
    document.querySelectorAll('.align-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.align === styles.align);
    });
}

function updatePreview() {
    if (!templateImg.src || !canvas.width || !canvas.height) return;
    
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(templateImg, 0, 0, canvas.width, canvas.height);
    
    const sampleData = excelData.length > 0 ? excelData[0] : {};
    
    for (let field of selectedFields) {
        const styles = textStyles[field];
        const pos = positions[field];
        const text = sampleData[field] || field;
        
        ctx.font = `${styles.size}px "${styles.font}"`;
        ctx.textAlign = styles.align;
        ctx.fillStyle = styles.color;

        // If this is the selected field, possibly show loading placeholder
        const isSelected = field === selectedField;

        if (isSelected && fontPreviewLoading) {
            // Animated placeholder text while font loads
            const dots = Math.floor((_fontPreviewAngle * 5) % 4);
            const placeholder = 'Loading' + '.'.repeat(dots);
            // Use a reliable fallback font for the placeholder (Inter)
            ctx.font = `${styles.size}px "Inter"`;
            ctx.fillStyle = 'rgba(71,85,105,0.9)'; // neutral/secondary color
            ctx.textAlign = styles.align;
            ctx.fillText(placeholder, pos.x, pos.y);
        } else {
            ctx.fillText(text, pos.x, pos.y);
        }

        if (isSelected) {
            let metrics = ctx.measureText(text);
            let textWidth = metrics.width;
            if (isSelected && fontPreviewLoading) {
                const dots = Math.floor((_fontPreviewAngle * 5) % 4);
                const placeholder = 'Loading' + '.'.repeat(dots);
                ctx.font = `${styles.size}px "Inter"`;
                metrics = ctx.measureText(placeholder);
                textWidth = metrics.width;
            }
            const textHeight = styles.size;

            let textX = pos.x;
            if (styles.align === 'center') textX -= textWidth / 2;
            else if (styles.align === 'right') textX -= textWidth;

            ctx.strokeStyle = '#6366f1';
            ctx.lineWidth = 2;
            ctx.setLineDash([5, 5]);
            ctx.strokeRect(textX - 5, pos.y - textHeight - 5, textWidth + 10, textHeight + 10);
            ctx.setLineDash([]);

            ctx.fillStyle = '#6366f1';
            ctx.beginPath();
            ctx.arc(pos.x, pos.y, 4, 0, 2 * Math.PI);
            ctx.fill();

            // If loading, draw a small rotating spinner to the right of the text
            if (fontPreviewLoading) {
                const spinnerX = textX + textWidth + 16;
                const spinnerY = pos.y - textHeight / 2;
                const r = Math.max(6, Math.min(12, Math.round(textHeight / 3)));
                const primary = getComputedStyle(document.documentElement).getPropertyValue('--primary-color') || '#06b6d4';

                ctx.save();
                ctx.translate(spinnerX, spinnerY);
                ctx.rotate(_fontPreviewAngle);
                ctx.lineWidth = 2;
                ctx.strokeStyle = primary.trim();
                ctx.beginPath();
                ctx.arc(0, 0, r, 0, Math.PI * 1.25);
                ctx.stroke();
                ctx.restore();
            }
        }
    }
}

// Helper function to yield to the event loop for UI updates
function yieldToMain() {
    return new Promise(resolve => setTimeout(resolve, 0));
}

// Certificate generation
async function generateCertificates() {
    if (!excelData.length) {
        showNotification('Please upload an Excel file first!', 'warning');
        return;
    }
    if (!templateImg.src) {
        showNotification('Please upload a template image first!', 'warning');
        return;
    }
    if (selectedFields.size === 0) {
        showNotification('Please select at least one field to display!', 'warning');
        return;
    }

    const btn = document.getElementById('generateBtn');
    const originalContent = btn.innerHTML;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Generating...';
    btn.disabled = true;

 try {
    await ensureJSZip();
    // Verify JSZip is actually available
    if (typeof JSZip === 'undefined') {
        throw new Error('JSZip library not available');
    }
} catch (e) {
    console.error('JSZip load error:', e);
    showNotification('Failed to load compression library. Please refresh the page and try again.', 'error');
    btn.innerHTML = originalContent;
    btn.disabled = false;
    return;
}

    const zip = new JSZip();
    let completedCertificates = 0;
    let failedCertificates = 0;
    const errors = [];
    const totalCertificates = excelData.length;

    // Use for...of with async/await for proper UI updates
    for (let index = 0; index < excelData.length; index++) {
        const data = excelData[index];

        try {
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            ctx.drawImage(templateImg, 0, 0, canvas.width, canvas.height);

            for (let field of selectedFields) {
                const styles = textStyles[field];
                const pos = positions[field];
                // Handle null/undefined values gracefully
                const rawValue = data[field];
                const text = rawValue !== null && rawValue !== undefined ? String(rawValue) : 'N/A';

                ctx.font = `${styles.size}px "${styles.font}"`;
                ctx.textAlign = styles.align;
                ctx.fillStyle = styles.color;
                ctx.fillText(text, pos.x, pos.y);
            }

            // Generate unique filename with sanitization
            const fileNameParts = Array.from(selectedFields)
                .slice(0, 2)
                .map(f => {
                    const val = data[f];
                    const strVal = val !== null && val !== undefined ? String(val) : 'NA';
                    // Sanitize: remove special chars, limit length
                    return strVal.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_').substring(0, 50);
                });

            // Add index to prevent duplicate filenames
            const fileName = `${fileNameParts.join('_')}_${index + 1}_certificate.png`;

            const dataURL = canvas.toDataURL('image/png');
            const base64Data = dataURL.replace(/^data:image\/png;base64,/, '');
            zip.file(fileName, base64Data, { base64: true });

            completedCertificates++;
        } catch (error) {
            console.error(`Error generating certificate for row ${index + 1}:`, error);
            failedCertificates++;
            errors.push(`Row ${index + 1}: ${error.message}`);
        }

        // Update progress and yield to event loop every 5 certificates
        if (index % 5 === 0 || index === excelData.length - 1) {
            btn.innerHTML = `<i class="fas fa-spinner fa-spin"></i> Generating... ${index + 1}/${totalCertificates}`;
            await yieldToMain();
        }
    }

    // Check if we have any certificates to zip
    if (completedCertificates === 0) {
        showNotification('Failed to generate any certificates. Please check your data.', 'error');
        btn.innerHTML = originalContent;
        btn.disabled = false;
        updatePreview();
        return;
    }

    // Generate the ZIP file
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Creating ZIP...';
    await yieldToMain();

    try {
        const blob = await zip.generateAsync({
            type: 'blob',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        }, (metadata) => {
            // Progress callback for ZIP generation
            const percent = Math.round(metadata.percent);
            btn.innerHTML = `<i class="fas fa-spinner fa-spin"></i> Compressing... ${percent}%`;
        });

        saveAs(blob, 'certificates.zip');

        btn.innerHTML = originalContent;
        btn.disabled = false;

        if (failedCertificates > 0) {
            showNotification(
                `Generated ${completedCertificates} certificates. ${failedCertificates} failed.`,
                'warning'
            );
            console.warn('Failed certificates:', errors);
        } else {
            showNotification(`Successfully generated ${completedCertificates} certificates!`, 'success');
        }

        updatePreview();
    } catch (error) {
        console.error('Error generating ZIP:', error);
        showNotification('An error occurred while creating the ZIP file. Please try again.', 'error');
        btn.innerHTML = originalContent;
        btn.disabled = false;
    }
}

// Notification system to replace alert()
function showNotification(message, type = 'info') {
    // Remove existing notification if any
    const existingNotification = document.querySelector('.notification-toast');
    if (existingNotification) {
        existingNotification.remove();
    }

    const notification = document.createElement('div');
    notification.className = `notification-toast notification-${type}`;

    const icons = {
        success: 'fa-check-circle',
        error: 'fa-exclamation-circle',
        warning: 'fa-exclamation-triangle',
        info: 'fa-info-circle'
    };

    notification.innerHTML = `
        <i class="fas ${icons[type] || icons.info}"></i>
        <span>${message}</span>
        <button class="notification-close" aria-label="Close notification">
            <i class="fas fa-times"></i>
        </button>
    `;

    document.body.appendChild(notification);

    // Add close button handler
    notification.querySelector('.notification-close').addEventListener('click', () => {
        notification.classList.add('notification-hiding');
        setTimeout(() => notification.remove(), 300);
    });

    // Trigger animation
    requestAnimationFrame(() => {
        notification.classList.add('notification-show');
    });

    // Auto-remove after delay (longer for errors)
    const duration = type === 'error' ? 6000 : type === 'warning' ? 5000 : 4000;
    setTimeout(() => {
        if (notification.parentNode) {
            notification.classList.add('notification-hiding');
            setTimeout(() => notification.remove(), 300);
        }
    }, duration);
}

// Patch notes
/* function checkFirstVisit() {
    const hasSeenPatchNotes = localStorage.getItem('certgen_patch_notes_v2');
    const patchNotesBtn = document.getElementById('patchNotesBtn');
    
    if (!hasSeenPatchNotes) {
        setTimeout(() => {
            showPatchNotes();
        }, 1000);
    } else {
        patchNotesBtn.classList.add('seen');
    }
} */

function showPatchNotes() {
    const overlay = document.getElementById('patchNotesOverlay');
    overlay.classList.add('show');
    overlay.removeAttribute('hidden');
    document.body.style.overflow = 'hidden';
    
    document.getElementById('patchNotesBtn').classList.add('seen');
}

/* function closePatchNotes() {
    const overlay = document.getElementById('patchNotesOverlay');
    const dontShowAgain = document.getElementById('dontShowAgain');
    
    overlay.classList.remove('show');
    document.body.style.overflow = '';
    
    if (dontShowAgain.checked) {
        localStorage.setItem('certgen_patch_notes_v2', 'seen');
    }
    
    localStorage.setItem('certgen_patch_notes_v2_session', 'seen');
    
    setTimeout(() => {
        overlay.setAttribute('hidden', 'true');
    }, 300);
} */

// Theme management
function initializeTheme() {
    const savedTheme = localStorage.getItem('certgen_theme') || 'system';
    applyTheme(savedTheme);
    updateThemeIcon(savedTheme);
    
    if (window.matchMedia) {
        window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => {
            const currentTheme = localStorage.getItem('certgen_theme') || 'system';
            if (currentTheme === 'system') {
                applyTheme('system');
            }
        });
    }
}

function toggleTheme() {
    const themeDropdown = document.getElementById('themeDropdown');
    themeDropdown.classList.toggle('show');
}

function setTheme(theme) {
    localStorage.setItem('certgen_theme', theme);
    applyTheme(theme);
    updateThemeIcon(theme);
    
    document.getElementById('themeDropdown').classList.remove('show');
    
    document.querySelectorAll('.theme-option').forEach(option => {
        option.classList.remove('active');
    });
    
    document.querySelector(`[data-theme="${theme}"]`).classList.add('active');
}

function applyTheme(theme) {
    const html = document.documentElement;
    
    if (theme === 'dark') {
        html.setAttribute('data-theme', 'dark');
    } else if (theme === 'light') {
        html.removeAttribute('data-theme');
    } else if (theme === 'system') {
        if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
            html.setAttribute('data-theme', 'dark');
        } else {
            html.removeAttribute('data-theme');
        }
    }
}

function updateThemeIcon(theme) {
    const themeIcon = document.getElementById('themeIcon');
    const isDark = theme === 'dark' || (theme === 'system' && window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches);
    
    if (theme === 'system') {
        themeIcon.className = 'fas fa-desktop';
    } else if (isDark) {
        themeIcon.className = 'fas fa-moon';
    } else {
        themeIcon.className = 'fas fa-sun';
    }
    
    document.querySelectorAll('.theme-option').forEach(option => {
        option.classList.remove('active');
    });
    
    document.querySelector(`[data-theme="${theme}"]`)?.classList.add('active');
}