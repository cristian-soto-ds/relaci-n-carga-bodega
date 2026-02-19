/**
 * Magaya HTML to Excel Converter - PRO Visual
 * Lógica de conversión y generación de Excel con estilos.
 * Codificación: ISO-8859-1 para compatibilidad con archivos Magaya.
 */

(function () {
    'use strict';

    // --- Referencias del DOM ---
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const resultArea = document.getElementById('result-area');
    const fileNameDisplay = document.getElementById('file-name-display');
    const outputNamePreview = document.getElementById('output-name-preview');
    const downloadBtn = document.getElementById('download-btn');
    const resetBtn = document.getElementById('reset-btn');
    const errorArea = document.getElementById('error-area');
    const errorMsg = document.getElementById('error-msg');
    const quoteElement = document.getElementById('daily-quote');

    // --- Estado de la aplicación ---
    let parsedData = [];
    let originalFileName = "Reporte";

    // CONFIGURACIÓN: Columnas a eliminar automáticamente
    const COLUMNS_TO_REMOVE = ["Piezas en Almacén", "Volumen en Almacén (m³)", "Peso en Almacén (kg)"];

    // --- Frases inspiradoras ---
    const quotes = [
        "\"La excelencia no es un acto, sino un hábito.\" – Aristóteles",
        "\"La calidad significa hacer lo correcto cuando nadie está mirando.\" – Henry Ford",
        "\"El éxito es la suma de pequeños esfuerzos repetidos día tras día.\" – Robert Collier",
        "\"La logística es el puente entre la idea y la realidad.\"",
        "\"El liderazgo es la capacidad de transformar la visión en realidad.\" – Warren Bennis",
        "\"No encuentres la falta, encuentra el remedio.\" – Henry Ford",
        "\"La eficiencia es hacer mejor lo que ya se está haciendo.\" – Peter Drucker",
        "\"El único lugar donde el éxito viene antes que el trabajo es en el diccionario.\" – Vidal Sassoon"
    ];

    // --- Utilidades ---
    function getCleanText(doc, selector) {
        const el = doc.querySelector(selector);
        return el ? el.textContent.trim() : null;
    }

    function showError(msg) {
        errorMsg.textContent = msg;
        errorArea.classList.remove('hidden');
    }

    function setRandomQuote() {
        if (quoteElement) {
            const randomQuote = quotes[Math.floor(Math.random() * quotes.length)];
            quoteElement.innerText = randomQuote;
        }
    }

    // --- Lógica de extracción de datos HTML ---
    function parseMagayaHTML(htmlString) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlString, 'text/html');
        parsedData = [];

        // 1. Extraer Encabezados (Filas 1-4)
        const companyName = getCleanText(doc, '.docname') || "Reporte";
        const titles = Array.from(doc.querySelectorAll('.title')).map(el => el.textContent.trim());
        const reportDate = getCleanText(doc, 'td[align="right"]');

        parsedData.push([companyName]);
        titles.forEach(t => parsedData.push([t]));
        if (reportDate) parsedData.push([reportDate]);

        // 2. Localizar la Tabla de Datos Correcta
        const tables = Array.from(doc.querySelectorAll('table'));
        let dataTable = null;
        let maxRows = 0;

        for (let table of tables) {
            const headerCells = table.querySelectorAll('td.sec.bar');
            if (headerCells.length > 5) {
                const rowCount = table.querySelectorAll('tr').length;
                const hasNestedTables = table.querySelector('table') !== null;
                if (!hasNestedTables || rowCount > maxRows) {
                    maxRows = rowCount;
                    dataTable = table;
                }
            }
        }
        if (!dataTable) {
            for (let table of tables) {
                if (table.querySelectorAll('td.sec.bar').length > 5) {
                    dataTable = table;
                    break;
                }
            }
        }
        if (!dataTable) throw new Error("No se encontró tabla de datos.");

        // 3. Filtrar Columnas y Guardar Índices Activos
        let finalHeaders = [];
        let activeIndices = [];

        const headerCells = dataTable.querySelectorAll('td.sec.bar');
        headerCells.forEach((cell, index) => {
            const text = cell.textContent.trim();
            if (!COLUMNS_TO_REMOVE.includes(text)) {
                finalHeaders.push(text);
                activeIndices.push(index);
            }
        });
        parsedData.push(finalHeaders);

        // 4. Extraer Filas de Datos
        const rows = dataTable.querySelectorAll('tr');
        let processingData = false;

        rows.forEach(row => {
            if (row.querySelector('.docname') || row.querySelector('.title') || row.querySelector('table') || row.innerText.trim() === '') return;

            if (row.querySelector('.sec.bar')) {
                processingData = true;
                return;
            }
            if (!processingData) return;

            const cells = row.querySelectorAll('td');
            if (cells.length === 0) return;

            let rowData = [];
            let isGroupRow = false;

            const firstCellText = cells[0]?.textContent.trim();
            if (firstCellText === companyName || titles.includes(firstCellText) || firstCellText === reportDate) return;

            // FILTRO DE FILAS DUPLICADAS (Subtotales redundantes)
            const hasUnderline = row.querySelector('u') !== null;
            if (hasUnderline && firstCellText !== "Total") {
                return;
            }

            const colspan = parseInt(cells[0].getAttribute('colspan') || "1");
            if (colspan > 5) isGroupRow = true;

            if (isGroupRow) {
                rowData = [firstCellText];
                for (let i = 1; i < finalHeaders.length; i++) rowData.push("");
            } else {
                cells.forEach((cell, index) => {
                    if (activeIndices.includes(index)) {
                        let text = cell.textContent.trim().replace(/\u00A0/g, " ");
                        const isDate = /\d{1,2}\/\d{1,2}\/\d{2,4}/.test(text);

                        if (!isDate && /^-?[\d,]+(\.\d+)?$/.test(text) && text !== "") {
                            const cleanNum = text.replace(/,/g, '');
                            !isNaN(parseFloat(cleanNum)) ? rowData.push(parseFloat(cleanNum)) : rowData.push(text);
                        } else {
                            rowData.push(text);
                        }
                    }
                });
            }

            if (rowData.length > 0) {
                if (firstCellText === "Total") {
                    parsedData.push([]);
                }
                parsedData.push(rowData);
            }
        });
    }

    // --- Generación del Excel (Estilos y Formato) ---
    function generateExcelPro() {
        if (parsedData.length === 0) return;

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(parsedData);

        const borderStyle = {
            top: { style: "thin", color: { rgb: "BDBDBD" } },
            bottom: { style: "thin", color: { rgb: "BDBDBD" } },
            left: { style: "thin", color: { rgb: "BDBDBD" } },
            right: { style: "thin", color: { rgb: "BDBDBD" } }
        };

        const styles = {
            mainHeader: {
                font: { bold: true, sz: 14, color: { rgb: "1E3A8A" }, name: "Calibri" },
                alignment: { horizontal: "center", vertical: "center" }
            },
            subHeader: {
                font: { bold: true, sz: 12, color: { rgb: "374151" }, name: "Calibri" },
                alignment: { horizontal: "center", vertical: "center" }
            },
            tableHeader: {
                fill: { fgColor: { rgb: "1E3A8A" } },
                font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11, name: "Calibri" },
                alignment: { horizontal: "center", vertical: "center", wrapText: true },
                border: borderStyle
            },
            groupRow: {
                fill: { fgColor: { rgb: "F3F4F6" } },
                font: { bold: true, color: { rgb: "1D4ED8" }, sz: 11, name: "Calibri" },
                border: borderStyle,
                alignment: { horizontal: "left", vertical: "center" }
            },
            totalRow: {
                fill: { fgColor: { rgb: "FEF9C3" } },
                font: { bold: true, sz: 11, name: "Calibri" },
                border: { top: { style: "double" }, bottom: { style: "medium" } },
                alignment: { vertical: "center" }
            },
            normalCell: {
                font: { name: "Calibri", sz: 11 },
                border: borderStyle,
                alignment: { vertical: "center" }
            }
        };

        const range = XLSX.utils.decode_range(ws['!ref']);
        const headerRowIndex = 4;

        const merges = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 9 } },
            { s: { r: 1, c: 0 }, e: { r: 1, c: 9 } },
            { s: { r: 2, c: 0 }, e: { r: 2, c: 9 } },
            { s: { r: 3, c: 0 }, e: { r: 3, c: 9 } }
        ];

        for (let R = range.s.r; R <= range.e.r; ++R) {
            const firstCellVal = ws[XLSX.utils.encode_cell({ r: R, c: 0 })] ? ws[XLSX.utils.encode_cell({ r: R, c: 0 })].v : "";
            const isTotalRow = String(firstCellVal).trim() === "Total";
            const secondCellVal = ws[XLSX.utils.encode_cell({ r: R, c: 1 })] ? ws[XLSX.utils.encode_cell({ r: R, c: 1 })].v : "";
            const isGroupRow = R > headerRowIndex && firstCellVal && (!secondCellVal || secondCellVal === "") && !isTotalRow;

            if (isGroupRow) {
                merges.push({ s: { r: R, c: 0 }, e: { r: R, c: 9 } });
            }

            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                if (!ws[cellRef]) continue;
                let cell = ws[cellRef];

                if (R === 0) cell.s = styles.mainHeader;
                else if (R <= 3) cell.s = styles.subHeader;
                else if (R === headerRowIndex) {
                    cell.s = styles.tableHeader;
                } else {
                    if (isTotalRow) {
                        cell.s = styles.totalRow;
                        if (typeof cell.v === 'number') cell.s.alignment = { horizontal: "right" };
                    } else if (isGroupRow) {
                        cell.s = styles.groupRow;
                    } else {
                        cell.s = JSON.parse(JSON.stringify(styles.normalCell));
                        if (typeof cell.v === 'number') {
                            cell.s.alignment = { horizontal: "right", vertical: "center" };
                            cell.z = "#,##0.00";
                        } else {
                            if (String(cell.v).includes("/")) cell.s.alignment = { horizontal: "center", vertical: "center" };
                        }

                        if (C === 1) {
                            cell.s.alignment = { horizontal: "center", vertical: "center" };
                            delete cell.z;
                        }
                        if (C === 5) {
                            cell.s.alignment = { horizontal: "left", vertical: "center" };
                        }
                        if (C === 6) {
                            cell.s.alignment = { horizontal: "center", vertical: "center" };
                            cell.z = "0";
                        }
                        if (C === 7 || C === 8) {
                            cell.s.alignment = { horizontal: "center", vertical: "center" };
                            if (typeof cell.v === 'number') {
                                if (cell.v % 1 === 0) {
                                    cell.z = "0";
                                } else {
                                    cell.z = "0.00";
                                }
                            }
                        }
                    }
                }
            }
        }

        ws['!merges'] = merges;

        const specificWidths = {
            0: 10.57,
            3: 27.57,
            5: 17.57
        };

        const colWidths = [];
        parsedData.forEach(row => {
            row.forEach((cell, i) => {
                if (row === parsedData[0] || row === parsedData[1]) return;
                if (specificWidths[i] === undefined) {
                    const len = (cell ? cell.toString().length : 0);
                    colWidths[i] = Math.max(colWidths[i] || 0, len + 2);
                }
            });
        });

        ws['!cols'] = [];
        const maxCol = Math.max(parsedData[4].length, 6);

        for (let i = 0; i < maxCol; i++) {
            if (specificWidths[i] !== undefined) {
                ws['!cols'][i] = { wch: specificWidths[i] };
            } else {
                let w = colWidths[i] || 10;
                ws['!cols'][i] = { wch: Math.min(Math.max(w, 10), 60) };
            }
        }

        XLSX.utils.book_append_sheet(wb, ws, "CARGA EN BODEGA");
        XLSX.writeFile(wb, `${originalFileName}.xlsx`);
    }

    // --- Procesamiento de archivos ---
    function processFile(file) {
        originalFileName = file.name.replace(/\.[^/.]+$/, "").toUpperCase();

        const reader = new FileReader();
        reader.onload = function (e) {
            try {
                parseMagayaHTML(e.target.result);
                dropZone.classList.add('hidden');
                resultArea.classList.remove('hidden');
                fileNameDisplay.textContent = file.name;
                outputNamePreview.textContent = `${originalFileName}.xlsx`;
                errorArea.classList.add('hidden');
            } catch (err) {
                console.error(err);
                showError("Error al leer el archivo. Verifica que sea un HTML válido de Magaya.");
            }
        };
        reader.readAsText(file, "ISO-8859-1");
    }

    // --- Reinicio de la vista ---
    function resetView() {
        resultArea.classList.add('hidden');
        dropZone.classList.remove('hidden');
        fileInput.value = '';
        parsedData = [];
        errorArea.classList.add('hidden');
        setRandomQuote();
    }

    // --- Inicialización: Event Listeners ---
    function initEventListeners() {
        dropZone.addEventListener('click', () => fileInput.click());
        dropZone.addEventListener('dragover', function (e) {
            e.preventDefault();
            dropZone.classList.add('drag-active');
        });
        dropZone.addEventListener('dragleave', function () {
            dropZone.classList.remove('drag-active');
        });
        dropZone.addEventListener('drop', function (e) {
            e.preventDefault();
            dropZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) processFile(e.dataTransfer.files[0]);
        });
        fileInput.addEventListener('change', function (e) {
            if (e.target.files.length) processFile(e.target.files[0]);
        });
        resetBtn.addEventListener('click', resetView);
        downloadBtn.addEventListener('click', generateExcelPro);
    }

    // --- Inicio ---
    setRandomQuote();
    initEventListeners();
})();
