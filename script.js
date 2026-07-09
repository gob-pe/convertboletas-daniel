let globalWorkbook = null;
let processedDataBlocks = [];
let visibleSheetsList = [];
let excludedSheets = new Set();


// Cabeceras estrictas del Excel de salida
const masterHeaders = [
    "AÑO", "DNI", "NOMBRES", "APELLIDOS", "TIPO CESANTÍA", "COD_MOD", "CARGO", 
    "TIPO_PENSIONISTA", "TIPO_PENSION", "NIVEL_MAG", "GRUPO_OCUP", "HORAS", 
    "TIEMPO_SERVICIO", "TIPO_SEGURO", "FECHA_REGIS", "CUENTA BANCARIA", "ESPECIAL"
];

const fileInput = document.getElementById('fileInput');
const btnProcess = document.getElementById('btnProcess');
const btnDownload = document.getElementById('btnDownload');
const alertBox = document.getElementById('alertBox');

function showStep(step) {
    [1, 2, 3].forEach(i => document.getElementById('step' + i).classList.add('hidden'));
    document.getElementById('step' + step).classList.remove('hidden');
}

function showError(msg) {
    alertBox.textContent = msg;
    alertBox.classList.remove('hidden');
    if (fileInput) fileInput.style.display = '';
}

// Función auxiliar para buscar el siguiente valor no vacío en una fila
function findNextValue(row, startIndex) {
    for (let k = startIndex + 1; k < row.length; k++) {
        let val = String(row[k] || '').trim();
        if (val !== '') return val;
    }
    return '';
}

// Función auxiliar para normalizar y limpiar texto (remover acentos, mayúsculas, múltiples espacios y caracteres raros)
function cleanText(str) {
    if (str === undefined || str === null) return '';
    return String(str)
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "") // Remueve acentos
        .replace(/[\s\u00a0\u200b]+/g, " ") // Convierte cualquier espacio (normal, no-breaking, zero-width) en un espacio regular único
        .toUpperCase()
        .trim();
}

// Función auxiliar para dividir un nombre completo en nombres y apellidos si vienen combinados
function splitFullName(fullName) {
    if (!fullName) return { nombres: "", apellidos: "" };
    let parts = fullName.trim().split(/\s+/);
    if (parts.length >= 3) {
        // En Perú/Latinoamérica, usualmente son 2 apellidos seguidos por nombres
        let apellidos = parts.slice(0, 2).join(" ");
        let nombres = parts.slice(2).join(" ");
        return { nombres, apellidos };
    } else if (parts.length === 2) {
        return { nombres: parts[1], apellidos: parts[0] };
    }
    return { nombres: fullName, apellidos: "" };
}

// Función auxiliar para actualizar la barra de estado de los años detectados en las hojas
function updateYearStatusBar() {
    const container = document.getElementById('yearStatusContainer');
    if (!container) return;
    container.innerHTML = '';
    
    visibleSheetsList.forEach(sheetName => {
        const ws = globalWorkbook.Sheets[sheetName];
        // Convertimos a JSON sin encabezados para buscar rápidamente conceptos y totales
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        
        let hasConcepts = false;
        let hasTotals = false;
        
        for (let i = 0; i < rows.length; i++) {
            let row = rows[i];
            if (!row || row.length === 0) continue;
            let firstVal = String(row[0] || '').trim();
            if (firstVal.startsWith('+') || firstVal.startsWith('-')) {
                hasConcepts = true;
            }
            if (['T-RENUM', 'T-DSCTO', 'T-LIQUI'].includes(firstVal)) {
                hasTotals = true;
            }
        }
        
        // Crear píldora/cápsula minimalista de estado
        const pill = document.createElement('span');
        pill.className = 'inline-flex items-center gap-1.5 text-xs font-bold font-mono px-2.5 py-1 rounded-full border transition-all duration-300 cursor-pointer select-none';
        
        // Si la hoja está excluida por el usuario
        if (excludedSheets.has(sheetName)) {
            pill.className += ' bg-slate-100 text-slate-400 border-slate-300 border-dashed line-through opacity-60';
            pill.innerHTML = `${sheetName} ✗`;
            pill.title = 'Excluido del proceso. Clic para volver a incluir.';
        } else {
            // Hojas activas
            if (hasConcepts && hasTotals) {
                pill.className += ' bg-emerald-50 text-emerald-700 border-emerald-200 hover:bg-emerald-100';
                pill.innerHTML = `${sheetName} <svg class="w-3 h-3 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="3.5" d="M5 13l4 4L19 7"></path></svg>`;
                pill.title = 'Activo. Estructura correcta. Clic para excluir del proceso.';
            } else if (!hasConcepts) {
                pill.className += ' bg-rose-50 text-rose-700 border-rose-200 hover:bg-rose-100';
                pill.innerHTML = `${sheetName} ⚠`;
                pill.title = `Activo. Hoja vacía o sin conceptos válidos (+/-). Clic para excluir del proceso.`;
            } else {
                pill.className += ' bg-amber-50 text-amber-700 border-amber-200 hover:bg-amber-100';
                pill.innerHTML = `${sheetName} ⚠`;
                pill.title = `Activo. Faltan totales de control (T-RENUM/T-DSCTO/T-LIQUI). Clic para excluir del proceso.`;
            }
        }
        
        // Evento click interactivo para alternar inclusión/exclusión
        pill.addEventListener('click', () => {
            if (excludedSheets.has(sheetName)) {
                excludedSheets.delete(sheetName);
            } else {
                excludedSheets.add(sheetName);
            }
            updateYearStatusBar();
        });
        
        container.appendChild(pill);
    });
}

// 1. CARGA DE ARCHIVO Y EXTRACCIÓN DINÁMICA MEJORADA
fileInput.addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    alertBox.classList.add('hidden');
    excludedSheets.clear();
    
    const reader = new FileReader();
    reader.onload = function(evt) {
        try {
            const data = evt.target.result;
            globalWorkbook = XLSX.read(data, { type: 'binary' });
            
            visibleSheetsList = globalWorkbook.SheetNames.filter(name => {
                const sheetInfo = globalWorkbook.Workbook?.Sheets?.find(s => s.name === name);
                return sheetInfo ? sheetInfo.Hidden !== 1 : true;
            });

            if(visibleSheetsList.length === 0) throw new Error("El archivo no contiene hojas visibles.");

            const firstWs = globalWorkbook.Sheets[visibleSheetsList[0]];
            const rows = XLSX.utils.sheet_to_json(firstWs, { header: 1, defval: "" });
            
            let ext = { nombres: "", apellidos: "", codmod: "", horas: "", cargo: "", cesantia: "", dni: "" };

            // Escaneo de las primeras 15 filas
            for(let i = 0; i < Math.min(15, rows.length); i++) {
                let row = rows[i];
                if(!row) continue;
                
                for(let j = 0; j < row.length; j++) {
                    // Limpiamos la celda: sin acentos, sin espacios extra y sin dos puntos ":"
                    let cell = cleanText(row[j]).replace(/:/g, "").trim();
                    if(!cell) continue;
                    
                    // Comparación exacta limpia para evitar falsos positivos
                    if (cell === "NOMBRES") {
                        ext.nombres = findNextValue(row, j);
                    } else if (cell === "APELLIDOS") {
                        ext.apellidos = findNextValue(row, j);
                    } else if (cell === "NOMBRES Y APELLIDOS" || cell === "APELLIDOS Y NOMBRES" || cell === "NOMBRES Y APELLIDOS DEL TRABAJADOR") {
                        let fullName = findNextValue(row, j);
                        let split = splitFullName(fullName);
                        ext.nombres = split.nombres;
                        ext.apellidos = split.apellidos;
                    } else if (cell === "CODIGO MODULAR" || cell === "COD MODULAR") {
                        ext.codmod = findNextValue(row, j);
                    } else if (cell === "HORAS") {
                        ext.horas = findNextValue(row, j);
                    } else if (cell === "CARGO") {
                        ext.cargo = findNextValue(row, j);
                    } else if (cell === "TIPO CESANTIA" || cell === "CESANTIA") {
                        ext.cesantia = findNextValue(row, j);
                    } else if (cell === "DNI") {
                        ext.dni = findNextValue(row, j);
                    }
                }
            }

            document.getElementById('f_nombres').value = ext.nombres;
            document.getElementById('f_apellidos').value = ext.apellidos;
            document.getElementById('f_codmod').value = ext.codmod;
            document.getElementById('f_horas').value = ext.horas;
            document.getElementById('f_cargo').value = ext.cargo;
            document.getElementById('f_cesantia').value = ext.cesantia;
            document.getElementById('f_dni').value = ext.dni; 

            // Secuencia de animación de carga premium
            const uploadView = document.getElementById('uploadView');
            const loadingView = document.getElementById('loadingView');
            const slidingDoc = document.getElementById('slidingDoc');
            const loadingSpinner = document.getElementById('loadingSpinner');
            const importedBadge = document.getElementById('importedBadge');
            const loadingStatusText = document.getElementById('loadingStatusText');
            const loadingSubText = document.getElementById('loadingSubText');

            // Deshabilitar clics/arrastres en el input mientras carga y ocultarlo para prevenir tooltips nativos congelados
            fileInput.classList.add('pointer-events-none');
            fileInput.style.display = 'none';

            // 1. Mostrar estado "Procesando" y activar animación de documento deslizante
            uploadView.classList.add('hidden');
            loadingView.classList.remove('hidden');
            loadingView.classList.add('flex');
            slidingDoc.classList.add('animate-document-slide');

            setTimeout(() => {
                // 2. Cambiar a estado "Éxito" (Tucked inside folder & badge pop)
                slidingDoc.classList.remove('animate-document-slide');
                slidingDoc.classList.add('success-doc-position');
                loadingSpinner.classList.add('hidden');
                importedBadge.classList.remove('scale-0');
                importedBadge.classList.add('scale-100');
                
                loadingStatusText.textContent = "¡Carga Completa!";
                loadingSubText.textContent = "Se importaron los datos maestros con éxito.";

                setTimeout(() => {
                    // 3. Restaurar UI de carga a su estado original (para futuras subidas) y pasar al Paso 2
                    uploadView.classList.remove('hidden');
                    loadingView.classList.add('hidden');
                    loadingView.classList.remove('flex');
                    slidingDoc.classList.remove('success-doc-position');
                    loadingSpinner.classList.remove('hidden');
                    importedBadge.classList.remove('scale-100');
                    importedBadge.classList.add('scale-0');
                    loadingStatusText.textContent = "Procesando planilla...";
                    loadingSubText.textContent = "Normalizando columnas y mapeando meses";
                    fileInput.classList.remove('pointer-events-none');
                    fileInput.style.display = '';

                    // Actualizar barra de estado de años detectados
                    updateYearStatusBar();
                    showStep(2);
                }, 1200); // 1.2 segundos mostrando éxito
            }, 1800); // 1.8 segundos de carga simulada
        } catch (err) {
            showError('Error al procesar el archivo. Asegúrese de que es un archivo Excel válido.');
            console.error(err);
        }
    };
    reader.readAsBinaryString(file);
});

// 2. PROCESAMIENTO CON CONTROL DE DECIMALES
btnProcess.addEventListener('click', () => {
    processedDataBlocks = [];
    alertBox.classList.add('hidden');
    
    const fNombres = document.getElementById('f_nombres');
    const fApellidos = document.getElementById('f_apellidos');
    const fDni = document.getElementById('f_dni');
    const fCesantia = document.getElementById('f_cesantia');

    let nombresVal = fNombres.value.trim();
    let apellidosVal = fApellidos.value.trim();
    let dniVal = fDni.value.trim();
    let cesantiaVal = fCesantia.value.trim();

    // Resetear clases invalid
    const allInputs = document.querySelectorAll('.input-field');
    allInputs.forEach(input => input.classList.remove('invalid'));

    let errors = [];

    // Validar nombres y apellidos
    if (!nombresVal) {
        fNombres.classList.add('invalid');
        errors.push("NOMBRES");
    }
    if (!apellidosVal) {
        fApellidos.classList.add('invalid');
        errors.push("APELLIDOS");
    }

    // Validar DNI
    if (dniVal) {
        // Auto-completado inteligente: si tiene 7 dígitos numéricos, rellenar con un cero a la izquierda
        if (dniVal.length === 7 && /^\d+$/.test(dniVal)) {
            dniVal = '0' + dniVal;
            fDni.value = dniVal;
        }
        // Validar longitud (8 para DNI, 9 para CE)
        if (dniVal.length !== 8 && dniVal.length !== 9) {
            fDni.classList.add('invalid');
            errors.push("DNI (debe tener 8 o 9 dígitos)");
        }
    }

    // Validar Tipo Cesantía (es obligatorio)
    if (!cesantiaVal) {
        fCesantia.classList.add('invalid');
        errors.push("TIPO CESANTÍA");
    }

    if (errors.length > 0) {
        showError("Los siguientes campos son obligatorios o tienen formato incorrecto: " + errors.join(", ") + ".");
        window.scrollTo({ top: 0, behavior: 'smooth' });
        return;
    }

    const md = {
        nombres: nombresVal,
        apellidos: apellidosVal,
        dni: dniVal,
        codmod: document.getElementById('f_codmod').value.trim(),
        cargo: document.getElementById('f_cargo').value.trim(),
        cesantia: cesantiaVal,
        horas: document.getElementById('f_horas').value.trim(),
        tipopen1: document.getElementById('f_tipopen_1').value.trim(),
        tipopen2: document.getElementById('f_tipopen_2').value.trim(),
        nivel: document.getElementById('f_nivel').value.trim(),
        grupo: document.getElementById('f_grupo').value.trim(),
        tiempo: document.getElementById('f_tiempo').value.trim(),
        seguro: document.getElementById('f_seguro').value.trim(),
        fecha: document.getElementById('f_fecha').value.trim(),
        cuenta: document.getElementById('f_cuenta').value.trim(),
        especial: document.getElementById('f_especial').value.trim()
    };

    const blocksByYear = {};

    // Filtrar únicamente los años/hojas activos elegidos por el usuario
    const activeSheets = visibleSheetsList.filter(name => !excludedSheets.has(name));
    
    if (activeSheets.length === 0) {
        showError("Debe incluir al menos un año/hoja en la barra de estado para poder ejecutar la consolidación.");
        return;
    }

    activeSheets.forEach((sheetName) => {
        const ws = globalWorkbook.Sheets[sheetName];
        
        let anioActual = sheetName.trim();
        let matchSheetName = anioActual.match(/\b(19\d{2}|20\d{2})\b/);
        if (matchSheetName) {
            anioActual = matchSheetName[1];
        } else {
            // Si el nombre de la hoja no contiene un año, buscamos en celdas probables de cabecera
            let foundYear = null;
            const possibleCells = ['B7', 'C7', 'D7', 'B8', 'C8'];
            for (let cellRef of possibleCells) {
                if (ws[cellRef] && ws[cellRef].v !== undefined && ws[cellRef].v !== null) {
                    let cellText = String(ws[cellRef].v);
                    let match = cellText.match(/\b(19\d{2}|20\d{2})\b/);
                    if (match) {
                        foundYear = match[1];
                        break;
                    }
                }
            }
            if (foundYear) {
                anioActual = foundYear;
            }
        }

        let rowMeta = [
            anioActual, md.dni, md.nombres, md.apellidos, md.cesantia, md.codmod, md.cargo, 
            md.tipopen1, md.tipopen2, md.nivel, md.grupo, md.horas, 
            md.tiempo, md.seguro, md.fecha, md.cuenta, md.especial
        ];

        // Inicializamos el bloque temporal con arrays vacíos de 12 meses para los totales
        let tempBlock = { 
            meta: rowMeta, 
            totals: { "T-RENUM": Array(12).fill(""), "T-DSCTO": Array(12).fill(""), "T-LIQUI": Array(12).fill("") }, 
            concepts: [] 
        };

        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        
        for (let i = 0; i < rows.length; i++) {
            let row = rows[i];
            if (!row || row.length === 0) continue;

            let cIdx = -1;
            let conceptName = "";
            let isTotal = false;

            for (let j = 0; j <= 5; j++) {
                let val = String(row[j] || '').trim();
                let upper = val.toUpperCase();
                if (val.startsWith('+') || val.startsWith('-') || ['T-RENUM', 'T-DSCTO', 'T-LIQUI'].includes(upper)) {
                    cIdx = j; conceptName = upper;
                    isTotal = ['T-RENUM', 'T-DSCTO', 'T-LIQUI'].includes(upper);
                    break;
                }
            }

            if (cIdx !== -1) {
                let months = [];
                for (let m = cIdx + 1; m <= cIdx + 12; m++) {
                    let val = row[m];
                    if (val === undefined || val === null || String(val).trim() === '.-' || String(val).trim() === '') {
                        months.push("");
                    } else {
                        let numVal = Number(val);
                        if (Number.isNaN(numVal)) {
                            months.push(val);
                        } else {
                            months.push(Number(numVal.toFixed(2)));
                        }
                    }
                }

                if (isTotal) { 
                    tempBlock.totals[conceptName] = months; 
                } else { 
                    tempBlock.concepts.push({ name: row[cIdx], months: months }); 
                }
            }
        }

        // Si la hoja contiene al menos un total o concepto procesable, la unificamos por año
        const hasData = tempBlock.concepts.length > 0 || tempBlock.totals["T-RENUM"].some(v => v !== "");
        if (!hasData) return;

        // Fusión o inicialización por año (anioActual)
        if (!blocksByYear[anioActual]) {
            blocksByYear[anioActual] = tempBlock;
        } else {
            // 1. Fusionar totales mes por mes
            ['T-RENUM', 'T-DSCTO', 'T-LIQUI'].forEach(t => {
                for (let m = 0; m < 12; m++) {
                    let valA = blocksByYear[anioActual].totals[t][m];
                    let valB = tempBlock.totals[t][m];
                    
                    if (valA !== "" && valB !== "") {
                        blocksByYear[anioActual].totals[t][m] = Number((Number(valA) + Number(valB)).toFixed(2));
                    } else if (valB !== "") {
                        blocksByYear[anioActual].totals[t][m] = valB;
                    }
                }
            });

            // 2. Fusionar conceptos individuales
            tempBlock.concepts.forEach(newConcept => {
                let existingConcept = blocksByYear[anioActual].concepts.find(
                    c => cleanText(c.name) === cleanText(newConcept.name)
                );
                
                if (existingConcept) {
                    for (let m = 0; m < 12; m++) {
                        let valA = existingConcept.months[m];
                        let valB = newConcept.months[m];
                        
                        if (valA !== "" && valB !== "") {
                            existingConcept.months[m] = Number((Number(valA) + Number(valB)).toFixed(2));
                        } else if (valB !== "") {
                            existingConcept.months[m] = valB;
                        }
                    }
                } else {
                    blocksByYear[anioActual].concepts.push(newConcept);
                }
            });
        }
    });

    // Convertir el objeto indexado por años de vuelta a un array ordenado cronológicamente
    processedDataBlocks = Object.values(blocksByYear).sort((a, b) => {
        let yearA = Number(a.meta[0]) || 0;
        let yearB = Number(b.meta[0]) || 0;
        return yearA - yearB;
    });

    // Ordenar los conceptos dentro de cada bloque para que todos los '+' vayan antes que los '-'
    processedDataBlocks.forEach(block => {
        block.concepts.sort((a, b) => {
            const aIsPlus = a.name.trim().startsWith('+');
            const bIsPlus = b.name.trim().startsWith('+');
            if (aIsPlus && !bIsPlus) return -1;
            if (!aIsPlus && bIsPlus) return 1;
            return 0; // Mantiene el orden relativo original para el mismo signo
        });
    });

    if (processedDataBlocks.length === 0) {
        showError("No se pudieron detectar conceptos válidos (+/-) en el archivo procesado.");
        return;
    }
    
    renderPreview();
    showStep(3);
});

// 3. RENDERIZADO VISUAL DEL RESULTADO
function renderPreview() {
    const tbody = document.getElementById('previewTableBody');
    tbody.innerHTML = '';

    processedDataBlocks.forEach(b => {
        let headRow = `<tr class="bg-slate-200 text-slate-700 font-bold">`;
        masterHeaders.forEach(h => headRow += `<td class="px-3 py-2 border-r border-slate-300">${h}</td>`);
        tbody.innerHTML += headRow + `</tr>`;

        let valRow = `<tr class="bg-white">`;
        b.meta.forEach(v => valRow += `<td class="px-3 py-2 border-r border-slate-200">${v}</td>`);
        tbody.innerHTML += valRow + `</tr><tr><td colspan="17" class="h-4 bg-slate-50"></td></tr>`;
        
        tbody.innerHTML += `<tr class="font-bold text-teal-800 bg-teal-50"><td class="px-3 py-2 border-r border-teal-200">MImponible</td><td colspan="16"></td></tr>`;
        ['T-RENUM', 'T-DSCTO', 'T-LIQUI'].forEach(t => {
            let rowHtml = `<td class="px-3 py-2 border-r border-slate-200 font-bold">${t}</td>`;
            (b.totals[t] || Array(12).fill("")).forEach(val => {
                rowHtml += `<td class="px-3 py-2 border-r border-slate-200 text-right">${val === '' ? '-' : val}</td>`;
            });
            tbody.innerHTML += `<tr class="bg-white">${rowHtml}<td colspan="4"></td></tr>`;
        });
        
        tbody.innerHTML += `<tr class="bg-slate-800 text-white font-bold"><td class="px-3 py-2">CONCEPTOS</td><td class="px-3 py-2 text-center border-l border-slate-600">ENE</td><td class="px-3 py-2 text-center border-l border-slate-600">FEB</td><td class="px-3 py-2 text-center border-l border-slate-600">MAR</td><td class="px-3 py-2 text-center border-l border-slate-600">ABR</td><td class="px-3 py-2 text-center border-l border-slate-600">MAY</td><td class="px-3 py-2 text-center border-l border-slate-600">JUN</td><td class="px-3 py-2 text-center border-l border-slate-600">JUL</td><td class="px-3 py-2 text-center border-l border-slate-600">AGO</td><td class="px-3 py-2 text-center border-l border-slate-600">SEP</td><td class="px-3 py-2 text-center border-l border-slate-600">OCT</td><td class="px-3 py-2 text-center border-l border-slate-600">NOV</td><td class="px-3 py-2 text-center border-l border-slate-600">DIC</td><td colspan="4"></td></tr>`;
        
        b.concepts.forEach(c => {
            let rowHtml = `<td class="px-3 py-2 border-r border-slate-200">${c.name}</td>`;
            c.months.forEach(val => {
                rowHtml += `<td class="px-3 py-2 border-r border-slate-200 text-right ${val === '' ? 'text-slate-400' : ''}">${val === '' ? '-' : val}</td>`;
            });
            tbody.innerHTML += `<tr class="bg-white hover:bg-slate-50">${rowHtml}<td colspan="4"></td></tr>`;
        });
        
        tbody.innerHTML += `<tr><td colspan="17" class="h-8 bg-slate-50"></td></tr>`;
    });
}

// Helper para generar celdas con estilo para el Excel descargado (usando xlsx-js-style)
function makeCell(val, isHeader = false) {
    if (val === undefined || val === null) val = "";
    
    const isNum = typeof val === 'number';
    const cellObj = {
        v: val,
        t: isNum ? 'n' : 's'
    };
    
    if (isHeader) {
        cellObj.s = {
            fill: { fgColor: { rgb: "1B365D" } }, // Azul marino idéntico al solicitado
            font: { color: { rgb: "FFFFFF" }, bold: true, name: "Calibri", sz: 11 },
            alignment: { horizontal: "center", vertical: "center", wrapText: true },
            border: {
                top: { style: "thin", color: { rgb: "cbd5e1" } },
                bottom: { style: "thin", color: { rgb: "cbd5e1" } },
                left: { style: "thin", color: { rgb: "cbd5e1" } },
                right: { style: "thin", color: { rgb: "cbd5e1" } }
            }
        };
    } else {
        cellObj.s = {
            font: { name: "Calibri", sz: 11 },
            alignment: { horizontal: isNum ? "right" : "left", vertical: "center" },
            border: {
                top: { style: "thin", color: { rgb: "e2e8f0" } },
                bottom: { style: "thin", color: { rgb: "e2e8f0" } },
                left: { style: "thin", color: { rgb: "e2e8f0" } },
                right: { style: "thin", color: { rgb: "e2e8f0" } }
            }
        };
        // Forzar formato numérico de dos decimales en Excel
        if (isNum) {
            cellObj.s.numFmt = "0.00";
        }
    }
    return cellObj;
}

// 4. EXPORTACIÓN FINAL A EXCEL CON ESTILOS
btnDownload.addEventListener('click', () => {
    const wb = XLSX.utils.book_new();
    const wsData = [];
    const headersMeses = ["CONCEPTOS", "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

    processedDataBlocks.forEach(b => {
        // Cabeceras maestras principales (Fila 1)
        wsData.push(masterHeaders.map(h => makeCell(h, true)));
        
        // Datos maestros (Fila 2)
        wsData.push(b.meta.map(v => makeCell(v, false)));
        
        // Fila vacía separadora
        wsData.push([]);
        
        // Fila de cabecera MImponible (con primera celda azul marino)
        const mimponRow = Array(17).fill("").map(() => makeCell("", false));
        mimponRow[0] = makeCell("MImponible", true);
        wsData.push(mimponRow);
        
        // Totales (T-RENUM, T-DSCTO, T-LIQUI)
        ['T-RENUM', 'T-DSCTO', 'T-LIQUI'].forEach(t => {
            const rowVals = [t, ...(b.totals[t] || Array(12).fill(""))];
            const rowCells = rowVals.map((val, idx) => {
                if (idx === 0) return makeCell(val, true); // Etiqueta azul marino
                return makeCell(val === '' ? '' : Number(val), false);
            });
            while(rowCells.length < 17) rowCells.push(makeCell("", false));
            wsData.push(rowCells);
        });
        
        // Fila de cabecera de meses
        const monthHeaderRow = headersMeses.map(m => makeCell(m, true));
        while(monthHeaderRow.length < 17) monthHeaderRow.push(makeCell("", true));
        wsData.push(monthHeaderRow);
        
        // Conceptos individuales
        b.concepts.forEach(c => {
            const rowVals = [c.name, ...c.months];
            const rowCells = rowVals.map((val, idx) => {
                if (idx === 0) return makeCell(val, false); // Nombre del concepto
                return makeCell(val === '' ? '' : Number(val), false);
            });
            while(rowCells.length < 17) rowCells.push(makeCell("", false));
            wsData.push(rowCells);
        });
        
        // Filas vacías entre bloques de años
        wsData.push([]); wsData.push([]);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    
    // Configurar anchos de columna óptimos
    ws['!cols'] = [
        { wch: 8 },   // AÑO
        { wch: 12 },  // DNI
        { wch: 25 },  // NOMBRES
        { wch: 25 },  // APELLIDOS
        { wch: 15 },  // TIPO CESANTÍA
        { wch: 15 },  // COD_MOD
        { wch: 20 },  // CARGO
        { wch: 18 },  // TIPO_PENSIONISTA
        { wch: 15 },  // TIPO_PENSION
        { wch: 12 },  // NIVEL_MAG
        { wch: 12 },  // GRUPO_OCUP
        { wch: 10 },  // HORAS
        { wch: 15 },  // TIEMPO_SERVICIO
        { wch: 15 },  // TIPO_SEGURO
        { wch: 15 },  // FECHA_REGIS
        { wch: 20 },  // CUENTA BANCARIA
        { wch: 30 }   // ESPECIAL
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Consolidado");

    // Construir un nombre de archivo dinámico y seguro a partir de Nombres y Apellidos
    const nombresInput = document.getElementById('f_nombres').value.trim();
    const apellidosInput = document.getElementById('f_apellidos').value.trim();
    let primerNombre = nombresInput.split(/\s+/)[0] || "";
    let primerApellido = apellidosInput.split(/\s+/)[0] || "";
    
    let safeFileName = "Planilla_Consolidada";
    if (primerApellido || primerNombre) {
        let namePart = cleanText(`${primerApellido}_${primerNombre}`).replace(/[^A-Z0-9_]/g, "");
        if (namePart) {
            safeFileName += `_${namePart}`;
        }
    }
    
    XLSX.writeFile(wb, `${safeFileName}.xlsx`);
});

// 5. NAVEGACIÓN DE REGRESO A EDICIÓN
document.getElementById('btnEditForm').addEventListener('click', () => {
    showStep(2);
});



