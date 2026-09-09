// Configurar worker de pdf.js
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// ==== ESTADO GLOBAL ====
let archivosProcesados = [];
let datosExcel = null;
let pdfListo = false;

// ==== DRAG & DROP IMÁGENES/PDF ====
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');

['dragenter', 'dragover'].forEach(ev => {
    dropZone.addEventListener(ev, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('dragover');
    });
});

['dragleave', 'drop'].forEach(ev => {
    dropZone.addEventListener(ev, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('dragover');
    });
});

dropZone.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        const dt = new DataTransfer();
        for (let i = 0; i < files.length; i++) {
            dt.items.add(files[i]);
        }
        fileInput.files = dt.files;
        procesarArchivos();
    }
});

dropZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', procesarArchivos);

// ==== DRAG & DROP EXCEL ====
const dropZoneExcel = document.getElementById('dropZoneExcel');
const fileExcel = document.getElementById('fileExcel');

['dragenter', 'dragover'].forEach(ev => {
    dropZoneExcel.addEventListener(ev, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZoneExcel.classList.add('dragover');
    });
});

['dragleave', 'drop'].forEach(ev => {
    dropZoneExcel.addEventListener(ev, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZoneExcel.classList.remove('dragover');
    });
});

dropZoneExcel.addEventListener('drop', (e) => {
    const file = e.dataTransfer.files[0];
    if (file) {
        fileExcel.files = e.dataTransfer.files;
        leerExcel(file);
    }
});

dropZoneExcel.addEventListener('click', () => fileExcel.click());
fileExcel.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        leerExcel(e.target.files[0]);
    }
});

// ==== MOSTRAR/OCULTAR OVERLAY DE CARGA ====
function mostrarCarga() {
    document.getElementById('loadingOverlay').classList.remove('hidden');
}

function ocultarCarga() {
    document.getElementById('loadingOverlay').classList.add('hidden');
}

// ==== LEER EXCEL (SOLO LEE Y GUARDA, SIN VALIDAR) ====
function leerExcel(file) {
    mostrarCarga();
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            // Convertir a matriz de filas (array de arrays)
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

            // Buscar fila de encabezados (DOCUMENTO y OBSERVACIÓN)
            let headerRow = -1;
            let colDoc = -1;
            let colObs = -1;

            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                for (let j = 0; j < row.length; j++) {
                    const cell = String(row[j]).toUpperCase();
                    if (cell.includes('DOCUMENTO')) colDoc = j;
                    if (cell.includes('OBSERVACIÓN') || cell.includes('OBSERVACION')) colObs = j;
                }
                if (colDoc !== -1 && colObs !== -1) {
                    headerRow = i;
                    break;
                }
            }

            if (headerRow === -1) {
                throw new Error('No se encontraron las columnas "DOCUMENTO" y "OBSERVACIÓN".');
            }

            // Leer datos
            datosExcel = [];
            for (let i = headerRow + 1; i < rows.length; i++) {
                const row = rows[i];
                const doc = String(row[colDoc] ?? '').trim();
                const obs = String(row[colObs] ?? '').trim().toUpperCase();
                if (doc) {
                    datosExcel.push({ documento: doc, observacion: obs, fila: i + 1 });
                }
            }

            // Actualizar UI del drop zone
            dropZoneExcel.classList.add('file-loaded');
            dropZoneExcel.querySelector('.drop-title').innerHTML = `<i class="fas fa-check-circle"></i> ${file.name} cargado correctamente`;
            setTimeout(() => dropZoneExcel.classList.remove('file-loaded'), 1500);

        } catch (error) {
            console.error(error);
            mostrarError('Error al leer el Excel: ' + error.message);
        } finally {
            ocultarCarga();
        }
    };
    reader.onerror = function() {
        ocultarCarga();
        mostrarError('No se pudo leer el archivo.');
    };
    reader.readAsArrayBuffer(file);
}

// ==== PROCESAR ARCHIVOS (IMÁGENES/PDF) - SOLO PROCESA, NO PREVIEW ====
async function procesarArchivos() {
    const files = fileInput.files;
    if (files.length === 0) return;

    dropZone.classList.add('file-loaded');
    dropZone.querySelector('.drop-title').innerHTML = `<i class="fas fa-spinner fa-spin"></i> Procesando ${files.length} archivos...`;

    archivosProcesados = [];

    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        await procesarArchivo(file);
    }

    archivosProcesados.sort((a, b) => a.numero - b.numero);

    dropZone.classList.remove('file-loaded');
    dropZone.querySelector('.drop-title').innerHTML = `<i class="fas fa-check-circle"></i> ${files.length} archivo(s) cargado(s)`;
    setTimeout(() => dropZone.classList.remove('file-loaded'), 1500);

    // Asegurarse de ocultar preview y mensajes (aún no validamos)
    document.getElementById('secPreview').classList.add('hidden');
    document.getElementById('mensajeValidacion').classList.add('hidden');
    pdfListo = false;
    document.getElementById('btnDescargar').disabled = true;
}

// ==== PROCESAR UN ARCHIVO INDIVIDUAL ====
async function procesarArchivo(file) {
    const nombreCompleto = file.name;
    const match = nombreCompleto.match(/^(\d+)\.\s*(\d+)/);
    if (!match) {
        alert(`El archivo "${nombreCompleto}" no sigue el formato "1. 10299841.jpg". Será omitido.`);
        return;
    }

    const numero = parseInt(match[1], 10);
    const documento = match[2];

    let imgDataUrl = '';

    if (file.type === 'application/pdf' || nombreCompleto.toLowerCase().endsWith('.pdf')) {
        imgDataUrl = await convertirPDFaImagen(file);
    } else {
        imgDataUrl = await leerArchivoComoImagen(file);
    }

    if (imgDataUrl) {
        archivosProcesados.push({
            numero,
            documento,
            imgDataUrl,
            fileName: nombreCompleto
        });
    }
}

// ==== RECORTE AUTOMÁTICO DE BORDES ====
function recortarBordesAutomatico(canvas) {
    const ctx = canvas.getContext('2d');
    const width = canvas.width;
    const height = canvas.height;
    const imageData = ctx.getImageData(0, 0, width, height);
    const data = imageData.data;

    let minX = width, minY = height, maxX = 0, maxY = 0;

    for (let y = 0; y < height; y++) {
        for (let x = 0; x < width; x++) {
            const i = (y * width + x) * 4;
            const r = data[i];
            const g = data[i + 1];
            const b = data[i + 2];

            const esBlanco = r > 240 && g > 240 && b > 240;
            const esNegro = r < 15 && g < 15 && b < 15;

            if (!esBlanco && !esNegro) {
                if (x < minX) minX = x;
                if (x > maxX) maxX = x;
                if (y < minY) minY = y;
                if (y > maxY) maxY = y;
            }
        }
    }

    if (maxX <= minX || maxY <= minY) return canvas;

    const margen = 10;
    minX = Math.max(0, minX - margen);
    minY = Math.max(0, minY - margen);
    maxX = Math.min(width, maxX + margen);
    maxY = Math.min(height, maxY + margen);

    const recorteWidth = maxX - minX;
    const recorteHeight = maxY - minY;
    const nuevoCanvas = document.createElement('canvas');
    nuevoCanvas.width = recorteWidth;
    nuevoCanvas.height = recorteHeight;
    const nuevoCtx = nuevoCanvas.getContext('2d');
    nuevoCtx.drawImage(canvas, minX, minY, recorteWidth, recorteHeight, 0, 0, recorteWidth, recorteHeight);

    return nuevoCanvas;
}

// ==== CONVERTIR PDF A IMAGEN ====
async function convertirPDFaImagen(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = async function(e) {
            try {
                const pdfData = new Uint8Array(e.target.result);
                const pdf = await pdfjsLib.getDocument({ data: pdfData }).promise;
                const page = await pdf.getPage(1);
                const viewport = page.getViewport({ scale: 2 });
                const canvas = document.createElement('canvas');
                canvas.width = viewport.width;
                canvas.height = viewport.height;
                const ctx = canvas.getContext('2d');
                await page.render({ canvasContext: ctx, viewport }).promise;

                const canvasRecortado = recortarBordesAutomatico(canvas);
                resolve(canvasRecortado.toDataURL('image/jpeg', 0.85));
            } catch (error) {
                console.error('Error al convertir PDF a imagen:', error);
                resolve('');
            }
        };
        reader.readAsArrayBuffer(file);
    });
}

// ==== LEER IMAGEN COMO DATAURL ====
function leerArchivoComoImagen(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            const img = new Image();
            img.onload = function() {
                const canvas = document.createElement('canvas');
                canvas.width = img.width;
                canvas.height = img.height;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0);

                const canvasRecortado = recortarBordesAutomatico(canvas);
                resolve(canvasRecortado.toDataURL('image/jpeg', 0.85));
            };
            img.src = e.target.result;
        };
        reader.readAsDataURL(file);
    });
}

// ==== VALIDAR Y PROCESAR (HACER TODO AQUÍ) ====
function validarYProcesar() {
    const mensajeDiv = document.getElementById('mensajeValidacion');
    mensajeDiv.classList.add('hidden');

    // Verificar que haya archivos cargados
    if (archivosProcesados.length === 0) {
        mostrarError('No hay imágenes/PDF cargados.');
        return;
    }

    // Verificar que haya Excel cargado
    if (!datosExcel || datosExcel.length === 0) {
        mostrarError('No se ha cargado un archivo Excel válido.');
        return;
    }

    // 1. Detectar documentos duplicados en el Excel
    const mapDocRows = {};
    const mensajesDuplicados = [];
    datosExcel.forEach(item => {
        if (!mapDocRows[item.documento]) {
            mapDocRows[item.documento] = [];
        }
        mapDocRows[item.documento].push(item.fila);
    });

    let hayDuplicados = false;
    for (const doc in mapDocRows) {
        if (mapDocRows[doc].length > 1) {
            hayDuplicados = true;
            mensajesDuplicados.push(`Documento ${doc} repetido en filas: ${mapDocRows[doc].join(', ')}`);
        }
    }

    // 2. Construir mapa de Excel: documento -> observacion
    const mapaExcel = {};
    datosExcel.forEach(item => {
        mapaExcel[item.documento] = item.observacion;
    });

    // 3. Validar cada imagen
    const errores = [];
    archivosProcesados.forEach(item => {
        const doc = item.documento;
        const obs = mapaExcel[doc];

        // ¿Existe el documento en el Excel?
        if (obs === undefined) {
            errores.push(`Documento ${doc}: No existe en el Excel.`);
            return;
        }

        // ¿La observación es válida?
        const obsNormalizada = obs.toUpperCase().replace(/\s+/g, ' ').trim();
        const validas = ['COMFACAUCA EN LINEA', 'CORR.BCRIO'];
        if (!validas.includes(obsNormalizada)) {
            errores.push(`Documento ${doc}: Observación "${obs}" no permitida.`);
        }
    });

    // 4. Mostrar errores o preparar descarga
    if (errores.length > 0) {
        mensajeDiv.innerHTML = '<strong>ERRORES DE VALIDACIÓN:</strong><br>' + errores.join('<br>');
        mensajeDiv.classList.remove('hidden');
        mensajeDiv.classList.add('error');
        pdfListo = false;
        document.getElementById('btnDescargar').disabled = true;
        return;
    } else {
        // Si hay duplicados, mostrar advertencia (pero permitir continuar)
        let mensajeFinal = '';
        if (hayDuplicados) {
            mensajeFinal = '<strong>⚠️ Advertencia:</strong> Existen documentos duplicados en el Excel:<br>' + mensajesDuplicados.join('<br>') + '<br><br>La validación de imágenes es correcta. Puedes descargar el PDF.';
            mensajeDiv.classList.add('warning');
        } else {
            mensajeFinal = '<strong>✅ Validación correcta.</strong> Todos los documentos coinciden. Ahora puedes descargar el PDF.';
            mensajeDiv.classList.add('success');
        }
        mensajeDiv.innerHTML = mensajeFinal;
        mensajeDiv.classList.remove('hidden');

        // Mostrar preview y habilitar descarga
        mostrarPrevisualizacion();
        pdfListo = true;
        document.getElementById('btnDescargar').disabled = false;
    }
}

// ==== DESCARGAR PDF (SOLO SI YA SE VALIDÓ) ====
function descargarPDF() {
    if (!pdfListo) {
        alert('Primero debes validar los documentos (haz clic en "Validar y Procesar").');
        return;
    }
    organizarDocumentos();
}

// ==== MOSTRAR ERROR ====
function mostrarError(mensaje) {
    const mensajeDiv = document.getElementById('mensajeValidacion');
    mensajeDiv.innerHTML = '<strong>ERROR:</strong> ' + mensaje;
    mensajeDiv.classList.remove('hidden');
    mensajeDiv.classList.add('error');
    pdfListo = false;
    document.getElementById('btnDescargar').disabled = true;
}

// ==== MOSTRAR PREVISUALIZACIÓN ====
function mostrarPrevisualizacion() {
    const secPreview = document.getElementById('secPreview');
    const gridPreview = document.getElementById('gridPreview');
    gridPreview.innerHTML = '';

    for (let i = 0; i < archivosProcesados.length; i += 4) {
        const grupo = archivosProcesados.slice(i, i + 4);
        const hojaDiv = document.createElement('div');
        hojaDiv.className = 'hoja-preview';

        grupo.forEach(item => {
            const itemDiv = document.createElement('div');
            itemDiv.className = 'item-preview';
            itemDiv.innerHTML = `
                <div class="item-header">${item.numero}. ${item.documento}</div>
                <img src="${item.imgDataUrl}" alt="${item.fileName}">
            `;
            hojaDiv.appendChild(itemDiv);
        });

        for (let j = grupo.length; j < 4; j++) {
            const emptyDiv = document.createElement('div');
            emptyDiv.className = 'item-preview empty';
            hojaDiv.appendChild(emptyDiv);
        }

        gridPreview.appendChild(hojaDiv);
    }

    secPreview.classList.remove('hidden');
}

// ==== GENERAR PDF FINAL ====
async function organizarDocumentos() {
    const boton = document.getElementById('btnDescargar');
    const textoOriginal = boton.innerHTML;
    boton.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Generando...';
    boton.disabled = true;

    try {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'letter' });

        const pageWidth = 216;
        const pageHeight = 279;
        const margin = 10;
        const gap = 2;
        const headerHeight = 8;

        const cellWidth = (pageWidth - 2 * margin - gap) / 2;
        const cellHeight = (pageHeight - 2 * margin - gap - headerHeight) / 2;

        // Pre-cargar imágenes
        const imagenes = [];
        for (let i = 0; i < archivosProcesados.length; i++) {
            const item = archivosProcesados[i];
            const img = new Image();
            await new Promise((resolve, reject) => {
                img.onload = resolve;
                img.onerror = reject;
                img.src = item.imgDataUrl;
            });
            imagenes.push(img);
        }

        // Generar páginas
        for (let i = 0; i < archivosProcesados.length; i += 4) {
            const grupo = archivosProcesados.slice(i, i + 4);

            if (i > 0) doc.addPage();

            for (let j = 0; j < grupo.length; j++) {
                const item = grupo[j];
                const img = imagenes[i + j];

                const col = j % 2;
                const fila = Math.floor(j / 2);

                const x = margin + col * (cellWidth + gap);
                const y = margin + fila * (cellHeight + gap + headerHeight);

                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.text(`${item.numero}. ${item.documento}`, x + 2, y + 3);

                const imgWidth = img.width;
                const imgHeight = img.height;
                const ratio = Math.min((cellWidth - 4) / imgWidth, (cellHeight - headerHeight - 2) / imgHeight);
                const newWidth = imgWidth * ratio;
                const newHeight = imgHeight * ratio;
                const imgX = x + (cellWidth - newWidth) / 2;
                const imgY = y + headerHeight + (cellHeight - headerHeight - newHeight) / 2;

                doc.addImage(item.imgDataUrl, 'JPEG', imgX, imgY, newWidth, newHeight);
            }
        }

        doc.save('documentos_validados.pdf');
    } catch (error) {
        console.error(error);
        alert('Error al generar PDF. Revisa la consola.');
    } finally {
        boton.innerHTML = textoOriginal;
        boton.disabled = false;
    }
}

// ==== LIMPIAR TODO ====
function limpiarTodo() {
    archivosProcesados = [];
    datosExcel = null;
    pdfListo = false;

    document.getElementById('fileInput').value = '';
    document.getElementById('fileExcel').value = '';
    document.getElementById('btnDescargar').disabled = true;

    const dropZone = document.getElementById('dropZone');
    dropZone.classList.remove('file-loaded', 'dragover');
    dropZone.querySelector('.drop-title').innerHTML = 'Arrastra imágenes/PDF aquí o haz clic';

    const dropZoneExcel = document.getElementById('dropZoneExcel');
    dropZoneExcel.classList.remove('file-loaded', 'dragover');
    dropZoneExcel.querySelector('.drop-title').innerHTML = 'Arrastra el Excel (.xlsx) aquí o haz clic';

    document.getElementById('secPreview').classList.add('hidden');
    document.getElementById('gridPreview').innerHTML = '';
    document.getElementById('mensajeValidacion').classList.add('hidden');
    document.getElementById('mensajeValidacion').classList.remove('warning', 'error', 'success');
}

// ==== SCROLL ====
function scrollToBottom() {
    window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' });
}

function scrollToTop() {
    window.scrollTo({ top: 0, behavior: 'smooth' });
}
