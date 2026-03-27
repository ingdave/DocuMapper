// Variables globales
let excelData = null;
let excelHeaders = [];
let wordPlaceholders = [];
let templateId = null;
let mappings = {};
let formatOptions = {};

const API_BASE = '/api';

// Función para formatear números en el cliente (preview)
function formatNumberClient(value, formatType) {
    const num = typeof value === 'number' ? value : parseFloat(value);
    if (isNaN(num)) return String(value || '');

    const hasDecimals = num % 1 !== 0;
    const decimals = hasDecimals ? 2 : 0;

    function fmtThousands(n) {
        const parts = Math.abs(n).toFixed(decimals).split('.');
        parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
        const formatted = parts.length > 1 ? parts[0] + ',' + parts[1] : parts[0];
        return n < 0 ? '-' + formatted : formatted;
    }

    switch (formatType) {
        case 'currency':
            return '$ ' + fmtThousands(num);
        case 'thousands':
            return fmtThousands(num);
        default:
            return String(value);
    }
}

// ===== UTILIDADES =====
function showToast(message, type = 'info') {
    const toastElement = document.getElementById('liveToast');
    const toastTitle = document.getElementById('toastTitle');
    const toastMessage = document.getElementById('toastMessage');
    const toastIcon = document.getElementById('toastIcon');

    // restablecer clases
    toastTitle.parentElement.classList.remove('bg-success', 'bg-danger', 'bg-info', 'bg-warning', 'text-white');
    
    if (type === 'success') {
        toastTitle.parentElement.classList.add('bg-success', 'text-white');
        toastTitle.textContent = 'Éxito';
        toastIcon.className = 'bi bi-check-circle-fill me-2';
    } else if (type === 'error') {
        toastTitle.parentElement.classList.add('bg-danger', 'text-white');
        toastTitle.textContent = 'Error';
        toastIcon.className = 'bi bi-exclamation-triangle-fill me-2';
    } else if (type === 'warning') {
        toastTitle.parentElement.classList.add('bg-warning', 'text-dark');
        toastTitle.textContent = 'Advertencia';
        toastIcon.className = 'bi bi-exclamation-circle-fill me-2';
    } else {
        toastTitle.parentElement.classList.add('bg-info', 'text-white');
        toastTitle.textContent = 'Información';
        toastIcon.className = 'bi bi-info-circle-fill me-2';
    }

    toastMessage.textContent = message;
    
    const toast = new bootstrap.Toast(toastElement);
    toast.show();
}

// Interceptar consola para mostrar en UI
const originalLog = console.log;
const originalWarn = console.warn;
const originalError = console.error;

console.log = (...args) => {
    originalLog(...args);
    showToast(args.join(' '), 'info');
};

console.warn = (...args) => {
    originalWarn(...args);
    showToast(args.join(' '), 'warning');
};

console.error = (...args) => {
    originalError(...args);
    showToast(args.join(' '), 'error');
};

function showStep(stepNumber) {
    for (let i = 1; i <= 5; i++) {
        const stepElement = document.getElementById(`step${i}`);
        if (stepElement) {
            stepElement.style.display = i === stepNumber ? 'block' : 'none';
        }
    }
    updateNavigationUI(stepNumber);
}

function updateNavigationUI(currentStep) {
    const navItems = document.querySelectorAll('.step-nav-item');
    navItems.forEach(item => {
        const step = parseInt(item.getAttribute('data-step'));
        item.classList.remove('active', 'completed');
        
        if (step === currentStep) {
            item.classList.add('active');
        } else if (step < currentStep) {
            item.classList.add('completed');
        }
    });
}

function navigateToStep(stepNumber) {
    // Validar si el usuario puede ir a ese paso
    if (stepNumber === 1) {
        showStep(1);
    } else if (stepNumber === 2 && excelData) {
        showStep(2);
    } else if (stepNumber === 3 && wordPlaceholders.length > 0) {
        showStep(3);
    } else if (stepNumber === 4 && Object.keys(mappings).length === wordPlaceholders.length && wordPlaceholders.length > 0) {
        showStep(4);
    } else if (stepNumber === 5 && generatedFilenames.length > 0) {
        showStep(5);
    } else {
        const messages = {
            2: 'Primero debes cargar un archivo Excel.',
            3: 'Primero debes cargar una plantilla Word.',
            4: 'Primero debes completar el mapeo de todos los campos.',
            5: 'Primero debes generar los documentos.'
        };
        showToast(messages[stepNumber] || 'No puedes saltar a este paso todavía.', 'warning');
    }
}

function showStatus(elementId, message, isSuccess, isError = false) {
    const element = document.getElementById(elementId);
    element.textContent = message;
    element.classList.remove('success', 'error');
    if (isSuccess) {
        element.classList.add('success');
    } else if (isError) {
        element.classList.add('error');
    }
}

function showError(message) {
    showToast(message, 'error');
}

// ===== EXCEL UPLOAD =====
const excelDropArea = document.getElementById('excelDropArea');
const excelInput = document.getElementById('excelInput');

excelDropArea.addEventListener('click', () => excelInput.click());

excelInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        uploadExcel(e.target.files[0]);
    }
});

excelDropArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    excelDropArea.classList.add('dragging');
});

excelDropArea.addEventListener('dragleave', () => {
    excelDropArea.classList.remove('dragging');
});

excelDropArea.addEventListener('drop', (e) => {
    e.preventDefault();
    excelDropArea.classList.remove('dragging');
    if (e.dataTransfer.files.length > 0) {
        uploadExcel(e.dataTransfer.files[0]);
    }
});

async function uploadExcel(file) {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        showError('Por favor carga un archivo Excel (.xlsx o .xls)');
        return;
    }

    showStatus('excelStatus', 'Cargando archivo...', false);

    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch(`${API_BASE}/excel/parse`, {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (result.success) {
            excelData = result.data;
            excelHeaders = result.headers;

            // Validar que haya datos
            if (!excelHeaders || excelHeaders.length === 0) {
                showError('El archivo Excel no tiene encabezados. Asegúrate de que la primera fila tenga los nombres de las columnas.');
                return;
            }

            if (!excelData || excelData.length === 0) {
                showError('El archivo Excel no tiene datos. Asegúrate de agregar filas con datos debajo de los encabezados.');
                return;
            }

            showStatus('excelStatus', `✓ ${result.totalRows} filas cargadas correctamente (${excelHeaders.length} columnas)`, true);

            // Avanzar a paso 2
            setTimeout(() => {
                showStep(2);
            }, 500);
        } else {
            showError(result.error || 'Error al procesar Excel');
        }
    } catch (error) {
        showError(error.message);
    }
}

// ===== WORD UPLOAD =====
const wordDropArea = document.getElementById('wordDropArea');
const wordInput = document.getElementById('wordInput');

wordDropArea.addEventListener('click', () => wordInput.click());

wordInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        uploadWord(e.target.files[0]);
    }
});

wordDropArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    wordDropArea.classList.add('dragging');
});

wordDropArea.addEventListener('dragleave', () => {
    wordDropArea.classList.remove('dragging');
});

wordDropArea.addEventListener('drop', (e) => {
    e.preventDefault();
    wordDropArea.classList.remove('dragging');
    if (e.dataTransfer.files.length > 0) {
        uploadWord(e.dataTransfer.files[0]);
    }
});

async function uploadWord(file) {
    if (!file.name.match(/\.docx$/i)) {
        showError('Por favor carga un archivo Word (.docx)');
        return;
    }

    showStatus('wordStatus', 'Analizando plantilla...', false);

    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch(`${API_BASE}/word/extract-variables`, {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (result.success) {
            wordPlaceholders = result.placeholders;
            templateId = result.templateId;

            showStatus('wordStatus', `✓ ${result.totalPlaceholders} placeholders encontrados`, true);

            // Avanzar a paso 3
            setTimeout(() => {
                createMappingUI();
                showStep(3);
            }, 500);
        } else {
            // Mostrar mensaje de error detallado
            let errorMsg = result.error || 'Error al procesar Word';
            if (result.details) {
                errorMsg += `\n\nDetalles: ${result.details}`;
            }
            showError(errorMsg);
        }
    } catch (error) {
        showError('Error de conexión: ' + error.message);
    }
}

// ===== MAPPING UI =====
function createMappingUI() {
    const container = document.getElementById('mappingContainer');
    container.innerHTML = '';

    // Validar que exista Excel
    if (!excelHeaders || excelHeaders.length === 0) {
        container.innerHTML = '<div class="alert alert-danger"><strong>Error:</strong> No se encontraron encabezados en el Excel.<br><br>Asegúrate de que:<br>✓ El Excel esté cargado (Paso 1)<br>✓ La primeira fila tenga los nombres de las columnas<br><br>Vuelve al Paso 1 e intenta cargar el Excel nuevamente.</div>';
        return;
    }

    // Validar que exista Word
    if (!wordPlaceholders || wordPlaceholders.length === 0) {
        container.innerHTML = '<div class="alert alert-danger"><strong>Error:</strong> No se encontraron placeholders en la plantilla Word.<br><br>Asegúrate de que:<br>✓ La plantilla Word esté cargada (Paso 2)<br>✓ Contenga placeholders en formato {PLACEHOLDER}<br><br>Vuelve al Paso 2 e intenta cargar la plantilla nuevamente.</div>';
        return;
    }

    wordPlaceholders.forEach(placeholder => {
        const item = document.createElement('div');
        item.className = 'mapping-item p-1 mb-1';

        const row = document.createElement('div');
        row.className = 'row g-1 align-items-end';

        const col1 = document.createElement('div');
        col1.className = 'col-12 col-sm-4';

        const labelLeft = document.createElement('label');
        labelLeft.className = 'form-label fw-bold';
        labelLeft.innerHTML = `Placeholder: <strong class="text-primary">{${placeholder}}</strong>`;
        col1.appendChild(labelLeft);

        const col2 = document.createElement('div');
        col2.className = 'col-12 col-sm-4';

        const labelRight = document.createElement('label');
        labelRight.className = 'form-label fw-bold';
        labelRight.textContent = 'Columna Excel:';
        col2.appendChild(labelRight);

        const selectRight = document.createElement('select');
        selectRight.className = 'form-select';
        selectRight.innerHTML = '<option value="">-- Selecciona columna Excel --</option>';

        excelHeaders.forEach(header => {
            const option = document.createElement('option');
            option.value = header;
            option.textContent = header;
            selectRight.appendChild(option);
        });

        // Col3: Selector de formato (solo visible para columnas numéricas)
        const col3 = document.createElement('div');
        col3.className = 'col-12 col-sm-4';
        col3.style.display = 'none';

        const labelFormat = document.createElement('label');
        labelFormat.className = 'form-label fw-bold';
        labelFormat.textContent = 'Formato:';
        col3.appendChild(labelFormat);

        const selectFormat = document.createElement('select');
        selectFormat.className = 'form-select';
        selectFormat.innerHTML = `
            <option value="raw">Sin formato (tal cual)</option>
            <option value="currency">Moneda ($ X.XXX)</option>
            <option value="thousands">Con separador (X.XXX)</option>
        `;

        selectFormat.addEventListener('change', (e) => {
            formatOptions[placeholder] = e.target.value;
        });

        col3.appendChild(selectFormat);

        selectRight.addEventListener('change', (e) => {
            mappings[placeholder] = e.target.value;

            // Mostrar/ocultar selector de formato según si la columna es numérica
            if (e.target.value && excelData && excelData.length > 0) {
                const isNumeric = excelData.slice(0, 5).some(r => typeof r[e.target.value] === 'number');
                if (isNumeric) {
                    col3.style.display = '';
                } else {
                    col3.style.display = 'none';
                    formatOptions[placeholder] = 'raw';
                    selectFormat.value = 'raw';
                }
            } else {
                col3.style.display = 'none';
                formatOptions[placeholder] = 'raw';
                selectFormat.value = 'raw';
            }

            updateMappingStatus();
        });

        col2.appendChild(selectRight);

        row.appendChild(col1);
        row.appendChild(col2);
        row.appendChild(col3);
        item.appendChild(row);
        container.appendChild(item);
    });

    updateMappingStatus();
}

function updateMappingStatus() {
    const mapped = Object.keys(mappings).filter(k => mappings[k]).length;
    document.getElementById('placeholdersMapped').textContent = mapped;
    document.getElementById('placeholdersTotal').textContent = wordPlaceholders.length;

    // Habilitar botón siguiente si todos están mapeados
    const allMapped = mapped === wordPlaceholders.length;

    // Si todos están mapeados, mostrar opción de continuar
    if (allMapped) {
        // Mostrar paso 4
        document.getElementById('filesToGenerate').textContent = excelData.length;

        // Poblar selector de columna para nombre de archivo
        const filenameSelect = document.getElementById('filenameColumnSelect');
        filenameSelect.innerHTML = '<option value="">-- Selecciona valor del prefijo --</option>';
        excelHeaders.forEach(header => {
            const option = document.createElement('option');
            option.value = header;
            option.textContent = header;
            filenameSelect.appendChild(option);
        });

        // Auto-seleccionar mes actual
        const meses = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC'];
        document.getElementById('monthSelect').value = meses[new Date().getMonth()];

        showStep(4);
    }
}

// ===== PREVIEW =====
async function showPreview() {
    const modal = document.getElementById('previewModal');
    const previewData = document.getElementById('previewData');
    const previewLoading = document.getElementById('previewLoading');

    previewLoading.style.display = 'block';
    previewData.innerHTML = '';
    
    // Show Bootstrap Modal
    const bsModal = new bootstrap.Modal(modal);
    bsModal.show();

    try {
        // Mostrar datos de muestra del primer registro
        const sampleData = excelData[0];
        const mappedData = {};

        for (const [placeholder, columnName] of Object.entries(mappings)) {
            let value = sampleData[columnName];
            const format = formatOptions[placeholder] || 'raw';
            if (format !== 'raw' && value !== undefined && value !== null && value !== '') {
                value = formatNumberClient(value, format);
            }
            mappedData[placeholder] = (value !== undefined && value !== null && value !== '') ? value : '(vacío)';
        }

        // Mostrar en la UI
        previewData.innerHTML = '';
        for (const [placeholder, value] of Object.entries(mappedData)) {
            const item = document.createElement('div');
            item.className = 'preview-item';
            item.innerHTML = `
                <div class="preview-label">{${placeholder}}</div>
                <div class="preview-value">${value}</div>
            `;
            previewData.appendChild(item);
        }

        previewLoading.style.display = 'none';
    } catch (error) {
        showError(error.message);
        const bsModal = bootstrap.Modal.getInstance(modal);
        if (bsModal) bsModal.hide();
    }
}

function closePreview() {
    const modal = document.getElementById('previewModal');
    const bsModal = bootstrap.Modal.getInstance(modal);
    if (bsModal) bsModal.hide();
}

window.onclick = function (event) {
    const modal = document.getElementById('previewModal');
    if (event.target === modal) {
        const bsModal = bootstrap.Modal.getInstance(modal);
        if (bsModal) bsModal.hide();
    }
};

// ===== GENERACIÓN DE DOCUMENTOS =====
async function generateAllDocuments() {
    const button = event.target;
    const progressDiv = document.getElementById('generationProgress');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');

    button.disabled = true;
    progressDiv.style.display = 'block';

    try {
        const totalFiles = excelData.length;

        // Validar mapeos
        const allMapped = Object.keys(mappings).length === wordPlaceholders.length &&
            Object.values(mappings).every(v => v);

        if (!allMapped) {
            showError('Debes mapear todos los placeholders antes de generar');
            button.disabled = false;
            progressDiv.style.display = 'none';
            return;
        }

        const initialName = document.getElementById('initialNameInput').value.trim() || 'Informe_supervision';

        // Enviar petición al servidor para generar todos los documentos
        const generationResponse = await fetch(`${API_BASE}/mapping/generate`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                templateId,
                mappings,
                excelData,
                filenameColumn: document.getElementById('filenameColumnSelect').value,
                month: document.getElementById('monthSelect').value,
                initialName: initialName,
                formatOptions: formatOptions
            })
        });

        const result = await generationResponse.json();

        if (result.success) {
            progressText.textContent = `✓ ${result.totalGenerated} documentos generados exitosamente`;

            // Guardar filenames para descargas en lote
            generatedFilenames = result.allFilenames || [];

            setTimeout(() => {
                showStep(5);
                document.getElementById('successCount').textContent = result.totalGenerated;

                // Crear lista de descargas
                const downloadList = document.getElementById('downloadList');
                downloadList.innerHTML = '';

                result.files.forEach((file, index) => {
                    if (file.filename) {
                        const col = document.createElement('div');
                        col.className = 'col-4';

                        const item = document.createElement('div');
                        item.className = 'download-item h-100';
                        item.innerHTML = `
                            <div class="d-flex flex-column gap-1 w-100 text-center">
                                <small class="text-muted text-truncate w-100 d-block mb-1">${file.filename}</small>
                                <a href="${file.downloadUrl}" download class="btn btn-sm btn-success text-white w-100 mt-2">
                                    <i class="bi bi-download"></i> Descargar
                                </a>
                            </div>
                        `;
                        col.appendChild(item);
                        downloadList.appendChild(col);
                    }
                });
            }, 500);
        } else {
            showError(result.error || 'Error al generar documentos');
        }

    } catch (error) {
        showError(error.message);
    } finally {
        button.disabled = false;
    }
}

// ===== RESET =====
function resetForm() {
    excelData = null;
    excelHeaders = [];
    wordPlaceholders = [];
    templateId = null;
    mappings = {};
    formatOptions = {};

    document.getElementById('excelInput').value = '';
    document.getElementById('wordInput').value = '';
    document.getElementById('excelStatus').classList.remove('success', 'error');
    document.getElementById('wordStatus').classList.remove('success', 'error');
    document.getElementById('mappingContainer').innerHTML = '';

    showStep(1);
}

// ===== DESCARGAS EN LOTE =====
let generatedFilenames = [];

async function downloadAllDocx() {
    if (generatedFilenames.length === 0) {
        showError('No hay documentos generados para descargar');
        return;
    }

    const button = event.target;
    const originalText = button.textContent;
    button.disabled = true;
    button.textContent = 'Preparando ZIP...';

    try {
        const response = await fetch(`${API_BASE}/mapping/download-all-docx`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filenames: generatedFilenames })
        });

        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || 'Error creating ZIP');
        }

        const result = await response.json();

        if (result.downloadUrl) {
            const link = document.createElement('a');
            link.href = result.downloadUrl;
            link.click();
            showStatus('downloadList', `✓ ZIP de ${generatedFilenames.length} documentos DOCX descargado`, true);
        }
    } catch (error) {
        showError(error.message);
    } finally {
        button.disabled = false;
        button.textContent = originalText;
    }
}

async function downloadAllPdf() {
    if (generatedFilenames.length === 0) {
        showError('No hay documentos generados para descargar');
        return;
    }

    const button = event.target;
    const originalText = button.innerHTML;
    button.disabled = true;
    button.innerHTML = '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span> Convirtiendo a PDF...';

    try {
        const response = await fetch(`${API_BASE}/mapping/download-all-pdf`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filenames: generatedFilenames })
        });

        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || 'Error creando PDF');
        }

        const result = await response.json();

        if (result.downloadUrl) {
            const link = document.createElement('a');
            link.href = result.downloadUrl;
            link.click();
            showStatus('downloadList', `✓ ZIP de ${generatedFilenames.length} documentos PDF descargado`, true);
        }
    } catch (error) {
        showError(error.message);
    } finally {
        button.disabled = false;
        button.innerHTML = originalText;
    }
}
// Inicializar
showStep(1);
