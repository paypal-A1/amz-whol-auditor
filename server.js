const express = require('express');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const { GoogleGenAI } = require('@google/genai');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

const upload = multer({ storage: multer.memoryStorage() });

// --------------------------------------------------------------
// 1. FUNCIÓN PARA BUSCAR VALOR POR NOMBRE (multi-idioma)
// --------------------------------------------------------------
function getColumnValue(row, posiblesNombres) {
    if (!row) return null;
    for (const nombre of posiblesNombres) {
        if (row[nombre] !== undefined && row[nombre] !== null && row[nombre] !== '') {
            return row[nombre];
        }
    }
    return null;
}

// --------------------------------------------------------------
// 2. FUNCIÓN PARA LLAMAR A GEMINI CON REINTENTOS INTELIGENTES
// --------------------------------------------------------------
async function callGeminiWithRetry(prompt, maxRetries = 4) {
    let lastError;
    const backoffDelays = [5000, 10000, 20000, 40000];

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            const response = await ai.models.generateContent({
                model: 'gemini-3.5-flash-lite',
                contents: prompt,
                config: {
                    responseMimeType: 'application/json'
                }
            });
            return response;
        } catch (error) {
            lastError = error;
            
            if (error.status === 503) {
                const waitTime = backoffDelays[attempt - 1] || 5000;
                console.log(`⏳ Gemini saturado (503), esperando ${waitTime/1000}s antes de reintentar (${attempt}/${maxRetries})`);
                await new Promise(resolve => setTimeout(resolve, waitTime));
                continue;
            }
            
            if (error.status === 429) {
                const mensaje = error.message || '';
                if (mensaje.includes('per day') || mensaje.includes('Daily')) {
                    console.error(`❌ LÍMITE DIARIO ALCANZADO: ${error.message}`);
                    const dailyLimitError = new Error('DAILY_LIMIT_REACHED');
                    dailyLimitError.isDailyLimit = true;
                    dailyLimitError.details = error.message;
                    throw dailyLimitError;
                }
                
                const waitTime = backoffDelays[attempt - 1] || 5000;
                console.log(`⏳ Cuota por minuto excedida (429), esperando ${waitTime/1000}s antes de reintentar (${attempt}/${maxRetries})`);
                await new Promise(resolve => setTimeout(resolve, waitTime));
                continue;
            }
            
            throw error;
        }
    }
    throw lastError;
}

// --------------------------------------------------------------
// 3. FUNCIÓN PARA EVALUAR VIABILIDAD
// --------------------------------------------------------------
function evaluarViabilidad(texto) {
    if (!texto) return 'neutral';
    const inicio = texto.trim().substring(0, 20);
    if (inicio.includes('✅') || inicio.includes('Apto') || inicio.includes('OK') || inicio.includes('Positivo')) {
        return 'positivo';
    }
    if (inicio.includes('❌') || inicio.includes('No apto') || inicio.includes('KO') || inicio.includes('Negativo') || inicio.includes('Inviable')) {
        return 'negativo';
    }
    return 'neutral';
}

function getColorStatus(fila) {
    const resKeepa = fila['Resumen Keepa'] || '';
    const resIA = fila['Resumen IA'] || '';
    const statusKeepa = evaluarViabilidad(String(resKeepa));
    const statusIA = evaluarViabilidad(String(resIA));
    if (statusKeepa === 'positivo' && statusIA === 'positivo') return 'verde';
    if (statusKeepa === 'negativo' || statusIA === 'negativo') return 'rojo';
    return 'amarillo';
}

// --------------------------------------------------------------
// 4. FUNCIÓN PARA CREAR HIPERVÍNCULO
// --------------------------------------------------------------
function createHyperlinkFromText(text) {
    if (!text || text === '') return { text: text || '', hyperlink: null };
    const str = String(text).trim();
    const emailMatch = str.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    if (emailMatch) {
        return { text: emailMatch[1], hyperlink: 'mailto:' + emailMatch[1] };
    }
    const urlMatch = str.match(/(https?:\/\/[^\s,;]+)/);
    if (urlMatch) {
        return { text: urlMatch[1], hyperlink: urlMatch[1] };
    }
    return { text: str, hyperlink: null };
}

// --------------------------------------------------------------
// 5. GENERAR DESCRIPCIÓN DE COLUMNA
// --------------------------------------------------------------
function getColumnDescription(colName, config) {
    const { roiAlto, roiMedio, roiBajo } = config;
    const descripciones = {
        'Título': 'Nombre completo del producto en Amazon',
        'ASIN': 'Amazon Standard Identification Number (clic para abrir en Amazon)',
        'Break-Even ($)': 'Punto de equilibrio (0% ROI). Fórmula: Precio Buy Box - FBA - Comisión - Envío - Prep',
        'Compra Máx (30%) ($)': `Precio máximo para ${roiAlto}% de ROI. Fórmula: Break-Even / (1 + ${roiAlto}/100)`,
        '% Desc. Req (30%)': `Descuento necesario para ${roiAlto}% de ROI`,
        'Compra Máx (20%) ($)': `Precio máximo para ${roiMedio}% de ROI`,
        '% Desc. Req (20%)': `Descuento necesario para ${roiMedio}% de ROI`,
        'Compra Máx (15%) ($)': `Precio máximo para ${roiBajo}% de ROI`,
        '% Desc. Req (15%)': `Descuento necesario para ${roiBajo}% de ROI`,
        'Est. # Ventas Mensual': 'Unidades estimadas mensuales',
        'Est. $ Ventas Mensual': 'Ingresos mensuales estimados',
        'Resumen Keepa': 'Resumen basado en datos Keepa y cálculos. Comienza con ✅ ⚠️ ❌',
        'Resumen IA': 'Resumen basado en investigación de IA. Comienza con ✅ ⚠️ ❌',
        'Admite Wholesale': 'Indica si la marca tiene programa mayorista en EE.UU.',
        'Tipo de Proveedor': 'Clasificación: Marca Directa, Distribuidor Autorizado, Mayorista Nacional',
        'Teléfono de Contacto': 'Teléfono de ventas/wholesale en EE.UU.',
        'Correo / Formulario': 'Email o enlace al formulario de apertura de cuenta',
        'Links Proveedores Potenciales': 'Enlaces a proveedores, distribuidores o formularios B2B',
        'Requisitos de Apertura': 'Requisitos para abrir cuenta mayorista (Tax ID, MOQ, etc.)',
        'Fabricante/Matriz': 'Fabricante real o corporación matriz',
        'Rutas de Distribución': 'Lista detallada de distribuidores autorizados',
        'Riesgo IP / Claims': 'Análisis de riesgo de Propiedad Intelectual',
        'Estrategia de Margen': 'Análisis de márgenes estimados y viabilidad financiera',
        'Conclusión General': 'Análisis integral combinando Keepa, cálculos e investigación de IA'
    };
    const compraMaxMatch = colName.match(/^Compra Máx \((\d+)%\) \(\$\)$/);
    if (compraMaxMatch) {
        const roi = compraMaxMatch[1];
        return `Precio máximo para ${roi}% de ROI. Fórmula: Break-Even / (1 + ${roi}/100)`;
    }
    const descReqMatch = colName.match(/^% Desc\. Req \((\d+)%\)$/);
    if (descReqMatch) {
        const roi = descReqMatch[1];
        return `Descuento necesario para ${roi}% de ROI`;
    }
    return descripciones[colName] || 'Columna generada por el sistema';
}

// --------------------------------------------------------------
// 6. FUNCIÓN PARA CREAR EL EXCEL CON EXCELJS
// --------------------------------------------------------------
async function createExcelWithStyles(filasProcesadas, config) {
    // --- Crear mapa ASIN -> URL ---
    const asinToUrl = {};
    filasProcesadas.forEach(row => {
        if (row['ASIN'] && row['URL: Amazon']) {
            asinToUrl[row['ASIN']] = row['URL: Amazon'];
        }
    });

    // --- Reordenar filas ---
    const grupos = { verde: [], amarillo: [], rojo: [] };
    filasProcesadas.forEach(row => {
        const color = getColorStatus(row);
        grupos[color].push(row);
    });

    // Ordenar dentro de cada grupo: por marca, luego por Est. # Ventas Mensual (desc), luego Est. $ Ventas Mensual (desc)
    const ordenarGrupo = (grupo) => {
        return grupo.sort((a, b) => {
            const marcaA = a['Marca'] || '';
            const marcaB = b['Marca'] || '';
            if (marcaA !== marcaB) return marcaA.localeCompare(marcaB);
            const ventasA = parseFloat(a['Est. # Ventas Mensual']) || 0;
            const ventasB = parseFloat(b['Est. # Ventas Mensual']) || 0;
            if (ventasA !== ventasB) return ventasB - ventasA;
            const dineroA = parseFloat(a['Est. $ Ventas Mensual']) || 0;
            const dineroB = parseFloat(b['Est. $ Ventas Mensual']) || 0;
            return dineroB - dineroA;
        });
    };

    const filasOrdenadas = [
        ...ordenarGrupo(grupos.verde),
        ...ordenarGrupo(grupos.amarillo),
        ...ordenarGrupo(grupos.rojo)
    ];

    // --- Definir orden de columnas (sin URL: Amazon) ---
    const todasLasColumnas = Object.keys(filasOrdenadas[0] || {});
    const bloque1 = ['Título', 'ASIN', 'Break-Even ($)', 'Compra Máx (30%) ($)', '% Desc. Req (30%)', 'Compra Máx (20%) ($)', '% Desc. Req (20%)', 'Compra Máx (15%) ($)', '% Desc. Req (15%)', 'Est. # Ventas Mensual', 'Est. $ Ventas Mensual'];
    const bloque2 = ['Resumen Keepa', 'Resumen IA'];
    const bloque3 = ['Admite Wholesale', 'Tipo de Proveedor', 'Teléfono de Contacto', 'Correo / Formulario', 'Links Proveedores Potenciales', 'Requisitos de Apertura', 'Fabricante/Matriz', 'Rutas de Distribución', 'Riesgo IP / Claims', 'Estrategia de Margen', 'Conclusión General'];
    const bloquesSet = new Set([...bloque1, ...bloque2, ...bloque3]);
    const bloque4 = todasLasColumnas.filter(col => !bloquesSet.has(col) && !col.includes('--- SEPARADOR ---') && col !== 'Viabilidad' && col !== 'URL: Amazon');

    const ordenFinal = [...bloque1, ...bloque2, ...bloque3, ...bloque4];
    const headers = ordenFinal.filter(col => todasLasColumnas.includes(col));

    // --- Crear workbook ---
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'AMZ Wholesale Auditor Pro';
    workbook.created = new Date();
    const worksheet = workbook.addWorksheet('Resultados Wholesale', {
        properties: { tabColor: { argb: 'FFD700' } }
    });

    // Configurar columnas
    worksheet.columns = headers.map(col => ({
        header: col,
        key: col,
        width: (bloque2.includes(col) || bloque3.includes(col)) ? 50 : 13
    }));
    // Título mantiene ancho automático
    const titleIndex = headers.indexOf('Título');
    if (titleIndex !== -1) {
        worksheet.getColumn(titleIndex + 1).width = 30;
    }

    // Agregar datos
    filasOrdenadas.forEach(row => {
        const rowData = {};
        headers.forEach(col => {
            const value = row[col] !== undefined && row[col] !== null ? row[col] : '';
            rowData[col] = value;
        });
        worksheet.addRow(rowData);
    });

    // Alto de fila (45)
    for (let rowNum = 1; rowNum <= worksheet.rowCount; rowNum++) {
        worksheet.getRow(rowNum).height = 45;
    }

    // Estilos encabezado
    const headerRow = worksheet.getRow(1);
    headerRow.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    const coloresBloques = [
        { inicio: 0, fin: bloque1.length - 1, color: 'FF5D6D7E' },
        { inicio: bloque1.length, fin: bloque1.length + bloque2.length - 1, color: 'FF1A237E' },
        { inicio: bloque1.length + bloque2.length, fin: bloque1.length + bloque2.length + bloque3.length - 1, color: 'FF283593' },
        { inicio: bloque1.length + bloque2.length + bloque3.length, fin: headers.length - 1, color: 'FF424242' }
    ];

    for (let i = 0; i < headers.length; i++) {
        const cell = headerRow.getCell(i + 1);
        let color = 'FF424242';
        for (const bloque of coloresBloques) {
            if (i >= bloque.inicio && i <= bloque.fin) {
                color = bloque.color;
                break;
            }
        }
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    }

    // Congelar paneles
    worksheet.views = [{ state: 'frozen', ySplit: 1, xSplit: 2 }];

    // ---- Formatos, hipervínculos y colores ----
    const colIndexMap = {};
    headers.forEach((col, idx) => colIndexMap[col] = idx + 1);

    for (let rowNum = 2; rowNum <= worksheet.rowCount; rowNum++) {
        const row = worksheet.getRow(rowNum);
        row.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

        // Color de fila según viabilidad (usando la clasificación ya hecha)
        const rowData = filasOrdenadas[rowNum - 2];
        const colorStatus = getColorStatus(rowData);
        let bgColor = null;
        if (colorStatus === 'verde') bgColor = 'FFC6EFCE';
        else if (colorStatus === 'rojo') bgColor = 'FFFFC7CE';
        else bgColor = 'FFFFEB9C';

        for (let colIdx = 0; colIdx < headers.length; colIdx++) {
            const colName = headers[colIdx];
            const cell = row.getCell(colIdx + 1);
            const value = cell.value;

            if (bgColor) {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
            }

            // Formato moneda/porcentaje
            let format = null;
            if (colName.includes('($)') || colName.includes('Break-Even') || colName.includes('Compra Máx') || colName.includes('Est. $') ||
                colName === 'Caja de Compra: Actual' || colName === 'Caja de Compra: Promedio de 30 días' ||
                colName === 'Caja de Compra: Promedio de 90 días' || colName === 'Caja de Compra: Promedio de 180 días' ||
                colName === 'Amazon: Promedio de 30 días' || colName === 'Amazon: Promedio de 90 días' ||
                colName === 'Tarifa FBA Pick&Pack') {
                format = '$#,##0.00';
            } else if (colName.includes('%') || colName.includes('Desc. Req')) {
                format = '0.00%';
            }
            if (format && typeof value === 'number') {
                cell.numFmt = format;
            }

            // Hipervínculo en ASIN
            if (colName === 'ASIN' && value) {
                const asin = String(value).trim();
                const url = asinToUrl[asin];
                if (url && url.startsWith('http')) {
                    cell.value = { text: asin, hyperlink: url };
                    cell.font = { color: { argb: 'FF0000FF' }, underline: true };
                }
            }

            // Hipervínculos en otras columnas
            if (colName === 'Correo / Formulario' || colName === 'Links Proveedores Potenciales') {
                if (value && typeof value === 'string') {
                    const { text, hyperlink } = createHyperlinkFromText(value);
                    if (hyperlink) {
                        cell.value = { text: text, hyperlink: hyperlink };
                        cell.font = { color: { argb: 'FF0000FF' }, underline: true };
                    } else {
                        cell.value = text;
                    }
                }
            }
        }
    }

    // Ancho de columnas
    worksheet.columns.forEach((col, idx) => {
        const header = col.header;
        if (header === 'Título') {
            col.width = 30;
        } else if (bloque2.includes(header) || bloque3.includes(header)) {
            let maxLen = header.length;
            col.eachCell({ includeEmpty: true }, (cell) => {
                const val = cell.value;
                if (val) {
                    const str = typeof val === 'object' ? (val.text || '') : String(val);
                    if (str.length > maxLen) maxLen = str.length;
                }
            });
            col.width = Math.min(Math.max(maxLen + 2, 20), 60);
        } else {
            col.width = 13; // Bloques 1 y 4
        }
    });

    // ---- Hoja de significado ----
    const meaningSheet = workbook.addWorksheet('📘 Significado de Columnas', {
        properties: { tabColor: { argb: 'FF2196F3' } }
    });
    meaningSheet.columns = [
        { header: 'Nombre de Columna', key: 'columna', width: 35 },
        { header: 'Descripción', key: 'descripcion', width: 70 }
    ];
    headers.forEach(col => {
        const desc = getColumnDescription(col, config);
        meaningSheet.addRow({ columna: col, descripcion: desc });
    });
    const meaningHeaderRow = meaningSheet.getRow(1);
    meaningHeaderRow.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    meaningHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1976D2' } };
    meaningHeaderRow.alignment = { horizontal: 'center', vertical: 'middle' };
    meaningHeaderRow.height = 22;

    const buffer = await workbook.xlsx.writeBuffer();
    return buffer;
}

// --------------------------------------------------------------
// 7. MOTOR PRINCIPAL DE PROCESAMIENTO
// --------------------------------------------------------------
async function procesarInventarioWholesale(fileBuffer, config) {
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
    if (rows.length === 0) {
        throw new Error('El archivo Excel no contiene datos.');
    }

    const encabezadosOriginales = Object.keys(rows[0]);

    const {
        prepFee,
        inboundShippingPound,
        supplierShippingUnit,
        roiAlto,
        roiMedio,
        roiBajo,
        priceBasis,
        minSalesMonthly
    } = config;

    const filasProcesadas = [];
    const productosPorMarca = {};

    console.log(`📊 Procesando ${rows.length} filas del Excel...`);

    for (const row of rows) {
        const titulo = getColumnValue(row, ['Title', 'Título']) || 'Sin Título';
        const asin = getColumnValue(row, ['ASIN']) || 'Desconocido';
        const marca = getColumnValue(row, ['Brand', 'Marca']) || 'Genérico';

        const ventasMensuales = parseFloat(
            getColumnValue(row, [
                'Tendencias de ventas mensuales: Comprados el mes pasado',
                'Bought past month',
                'Monthly Sales',
                'Sales Drops (30 days)'
            ]) || 0
        );

        if (ventasMensuales < minSalesMonthly) {
            console.log(`⏭️ Descartando producto con ventas ${ventasMensuales} < ${minSalesMonthly}: ${asin}`);
            continue;
        }

        let precioBuyBox = 0;
        if (priceBasis === '90day') {
            precioBuyBox = parseFloat(
                getColumnValue(row, [
                    'Caja de Compra: Promedio de 90 días',
                    'Buy Box: 90 days avg',
                    'Amazon 90 days avg'
                ]) || 0
            );
        } else {
            precioBuyBox = parseFloat(
                getColumnValue(row, [
                    'Caja de Compra: Actual',
                    'Buy Box: Current',
                    'Precio Actual'
                ]) || 0
            );
        }
        if (!precioBuyBox || precioBuyBox === 0) {
            console.log(`⏭️ Descartando producto sin precio: ${asin}`);
            continue;
        }

        const pesoGramos = parseFloat(
            getColumnValue(row, [
                'Paquete: Peso (g)',
                'Weight (g)'
            ]) || 0
        );
        const pesoLibras = pesoGramos * 0.00220462;
        const costoEnvioAmazon = pesoLibras * inboundShippingPound;

        const referralFee = precioBuyBox * 0.15;
        const fbaFee = parseFloat(
            getColumnValue(row, [
                'Tarifa FBA Pick&Pack',
                'FBA Pick & Pack Fee'
            ]) || 0
        );

        const breakEven = precioBuyBox - fbaFee - referralFee - costoEnvioAmazon - prepFee - supplierShippingUnit;

        const calcularCompraMax = (roi) => breakEven / (1 + (roi / 100));
        const calcularDescuentoFraccion = (precioMax) => (precioBuyBox - precioMax) / precioBuyBox;

        const maxAlto = calcularCompraMax(roiAlto);
        const maxMedio = calcularCompraMax(roiMedio);
        const maxBajo = calcularCompraMax(roiBajo);
        const descAlto = calcularDescuentoFraccion(maxAlto);
        const descMedio = calcularDescuentoFraccion(maxMedio);
        const descBajo = calcularDescuentoFraccion(maxBajo);

        const fbaElegibles = parseInt(
            getColumnValue(row, ['Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBA']) || 0
        );
        const fbmElegibles = parseInt(
            getColumnValue(row, ['Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBM']) || 0
        );
        
        const pctMejorVendedor30d = parseFloat(
            getColumnValue(row, ['Caja de Compra: % Mejor vendedor 30 días'])
        );
        
        let estVentasUnidades = 0;
        if (pctMejorVendedor30d && pctMejorVendedor30d > 0 && (fbaElegibles + fbmElegibles) > 0) {
            const pct = pctMejorVendedor30d / 100;
            const ventasRestantes = ventasMensuales * (1 - pct);
            const competidoresRestantes = fbaElegibles + fbmElegibles;
            estVentasUnidades = ventasRestantes / competidoresRestantes;
        }
        const estVentasDolares = estVentasUnidades * precioBuyBox;

        const filaConMetricas = {};
        for (const key of encabezadosOriginales) {
            filaConMetricas[key] = row[key];
        }
        
        filaConMetricas['Break-Even ($)'] = breakEven;
        filaConMetricas[`Compra Máx (${roiAlto}%) ($)`] = maxAlto;
        filaConMetricas[`% Desc. Req (${roiAlto}%)`] = descAlto;
        filaConMetricas[`Compra Máx (${roiMedio}%) ($)`] = maxMedio;
        filaConMetricas[`% Desc. Req (${roiMedio}%)`] = descMedio;
        filaConMetricas[`Compra Máx (${roiBajo}%) ($)`] = maxBajo;
        filaConMetricas[`% Desc. Req (${roiBajo}%)`] = descBajo;
        filaConMetricas['Est. # Ventas Mensual'] = Math.round(estVentasUnidades);
        filaConMetricas['Est. $ Ventas Mensual'] = estVentasDolares;
        
        filaConMetricas['Resumen Keepa'] = '';
        filaConMetricas['Resumen IA'] = '';
        filaConMetricas['Admite Wholesale'] = '';
        filaConMetricas['Tipo de Proveedor'] = '';
        filaConMetricas['Teléfono de Contacto'] = '';
        filaConMetricas['Correo / Formulario'] = '';
        filaConMetricas['Links Proveedores Potenciales'] = '';
        filaConMetricas['Requisitos de Apertura'] = '';
        filaConMetricas['Fabricante/Matriz'] = '';
        filaConMetricas['Rutas de Distribución'] = '';
        filaConMetricas['Riesgo IP / Claims'] = '';
        filaConMetricas['Estrategia de Margen'] = '';
        filaConMetricas['Conclusión General'] = '';

        filasProcesadas.push(filaConMetricas);

        if (!productosPorMarca[marca]) productosPorMarca[marca] = [];
        productosPorMarca[marca].push({ asin, title: titulo, rowRef: filaConMetricas });
    }

    console.log(`✅ ${filasProcesadas.length} productos aprobados para análisis.`);
    console.log(`📦 Marcas identificadas: ${Object.keys(productosPorMarca).length}`);

    // ---- AUDITORÍA CON IA ----
    const marcas = Object.keys(productosPorMarca);
    marcas.sort((a, b) => productosPorMarca[b].length - productosPorMarca[a].length);

    let solicitudesRealizadas = 0;
    const startTime = Date.now();
    const LIMITE_DIARIO = 1500;
    let limiteAlcanzado = false;

    for (let i = 0; i < marcas.length; i++) {
        const nombreMarca = marcas[i];
        const productos = productosPorMarca[nombreMarca];

        if (limiteAlcanzado || solicitudesRealizadas >= LIMITE_DIARIO) {
            console.log(`⛔ Límite diario de ${LIMITE_DIARIO} solicitudes alcanzado.`);
            break;
        }

        console.log(`\n🔍 Procesando marca ${i+1}/${marcas.length}: "${nombreMarca}" (${productos.length} productos)`);
        console.log(`📊 Solicitudes realizadas: ${solicitudesRealizadas}/${LIMITE_DIARIO}`);

        if (i > 0) {
            console.log(`⏳ Esperando 5s antes de siguiente marca...`);
            await new Promise(resolve => setTimeout(resolve, 5000));
        }

        try {
            const prompt = `
                Actúa como un detective de proveedores para Amazon Wholesale. Analiza en profundidad la marca "${nombreMarca}".
                
                Productos asociados: ${JSON.stringify(productos.map(p => ({ asin: p.asin, title: p.title })))}
                
                Investiga y proporciona un análisis detallado con la siguiente estructura EXACTA (formato JSON).
                SI NO ENCUENTRAS INFORMACIÓN, USA null o "No encontrado".
                NO INVENTES DATOS. Los enlaces deben ser reales y verificados.
                IMPORTANTE: Para "resumenKeepa" y "resumenIA", comienza el texto con ✅ si es positivo, ⚠️ si es neutro, o ❌ si es negativo.
                
                {
                    "resumenKeepa": "Resumen corto (máx 1 línea) basado en los datos de Keepa y cálculos financieros. Evalúa demanda, competencia y márgenes. Comienza con ✅, ⚠️ o ❌.",
                    "resumenIA": "Resumen corto (máx 1 línea) basado en la investigación de la IA sobre la marca. Evalúa wholesale, contactos y riesgos. Comienza con ✅, ⚠️ o ❌.",
                    "admiteWholesale": "Sí" o "No" o "No encontrado",
                    "tipoProveedor": "Marca Directa" o "Distribuidor Autorizado" o "Mayorista Nacional" o "No encontrado",
                    "telefono": "Número de teléfono de ventas/wholesale en EE.UU. o null",
                    "contacto": "Email de wholesale o enlace al formulario de apertura de cuenta o null",
                    "links": "Enlaces directos a páginas de proveedores o formularios B2B (separados por comas). Solo enlaces reales.",
                    "requisitos": "Requisitos de apertura de cuenta (Tax ID, Resale Certificate, MOQ, etc.) o null",
                    "fabricante": "Nombre del fabricante real o corporación matriz. Si es marca propia, indica 'Marca propia'.",
                    "rutas_distribucion": "Lista detallada de distribuidores autorizados en EE.UU. Incluye: 1) Nombre, 2) Tipo, 3) Enlace web (solo reales), 4) Notas sobre requisitos.",
                    "riesgo_ip": "Análisis del riesgo de Propiedad Intelectual: 1) Protección de marca en Amazon, 2) Número de vendedores FBA, 3) Recomendación.",
                    "estrategia_margen": "Análisis de márgenes: 1) Estimación de precio de compra, 2) Margen bruto estimado tras FBA, 3) Recomendación de viabilidad.",
                    "conclusion": "Análisis INTEGRAL Y DETALLADO (mínimo 200 palabras) combinando los datos de Keepa, los cálculos financieros y la investigación de IA. Debe incluir: análisis de demanda y competencia, viabilidad de márgenes, quién está detrás de la marca, rutas de distribución, riesgo de IP, y una recomendación final clara: CONTACTAR, EVITAR o INVESTIGAR MÁS."
                }
            `;

            const response = await callGeminiWithRetry(prompt);
            solicitudesRealizadas++;

            let textoLimpio = response.text;
            textoLimpio = textoLimpio.replace(/```json/gi, '');
            textoLimpio = textoLimpio.replace(/```/g, '');
            textoLimpio = textoLimpio.trim();

            const datosIA = JSON.parse(textoLimpio);

            for (const prod of productos) {
                const info = datosIA[prod.asin] || datosIA;
                if (info) {
                    prod.rowRef['Resumen Keepa'] = info.resumenKeepa || '';
                    prod.rowRef['Resumen IA'] = info.resumenIA || '';
                    prod.rowRef['Admite Wholesale'] = info.admiteWholesale || '';
                    prod.rowRef['Tipo de Proveedor'] = info.tipoProveedor || '';
                    prod.rowRef['Teléfono de Contacto'] = info.telefono || '';
                    prod.rowRef['Correo / Formulario'] = info.contacto || '';
                    prod.rowRef['Links Proveedores Potenciales'] = info.links || '';
                    prod.rowRef['Requisitos de Apertura'] = info.requisitos || '';
                    prod.rowRef['Fabricante/Matriz'] = info.fabricante || '';
                    prod.rowRef['Rutas de Distribución'] = info.rutas_distribucion || '';
                    prod.rowRef['Riesgo IP / Claims'] = info.riesgo_ip || '';
                    prod.rowRef['Estrategia de Margen'] = info.estrategia_margen || '';
                    prod.rowRef['Conclusión General'] = info.conclusion || '';
                }
            }
            
            const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
            console.log(`✅ Marca "${nombreMarca}" procesada. Solicitudes: ${solicitudesRealizadas}/${LIMITE_DIARIO} en ${elapsed}s`);

        } catch (error) {
            if (error.isDailyLimit) {
                console.log(`⛔ Límite diario de solicitudes alcanzado. Deteniendo procesamiento.`);
                limiteAlcanzado = true;
                break;
            }
            console.error(`❌ Error procesando marca "${nombreMarca}":`, error.message);
        }
    }

    const totalTime = ((Date.now() - startTime) / 1000).toFixed(1);
    console.log(`\n📊 Resumen final:`);
    console.log(`   - Solicitudes realizadas: ${solicitudesRealizadas}/${LIMITE_DIARIO}`);
    console.log(`   - Tiempo total: ${totalTime}s`);
    console.log(`   - Marcas procesadas: ${solicitudesRealizadas}`);
    console.log(`   - Marcas pendientes: ${marcas.length - solicitudesRealizadas}`);

    // ---- GENERAR EXCEL ----
    const buffer = await createExcelWithStyles(filasProcesadas, config);
    
    return {
        buffer: buffer,
        solicitudesRealizadas,
        marcasProcesadas: solicitudesRealizadas,
        marcasPendientes: marcas.length - solicitudesRealizadas,
        limiteAlcanzado
    };
}

// --------------------------------------------------------------
// 8. ENDPOINT
// --------------------------------------------------------------
app.post('/api/audit-excel', upload.single('excelFile'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No se ha cargado ningún archivo Excel.' });
        }

        console.log(`📁 Archivo recibido: ${req.file.originalname} (${req.file.size} bytes)`);

        const config = {
            prepFee: parseFloat(req.body.prepFee || 1.50),
            inboundShippingPound: parseFloat(req.body.inboundShippingPound || 1.00),
            supplierShippingUnit: parseFloat(req.body.supplierShippingUnit || 0.00),
            roiAlto: parseFloat(req.body.roiAlto || 30),
            roiMedio: parseFloat(req.body.roiMedio || 20),
            roiBajo: parseFloat(req.body.roiBajo || 15),
            priceBasis: req.body.priceBasis || '90day',
            minSalesMonthly: parseFloat(req.body.minSalesMonthly || 100)
        };

        console.log('⚙️ Configuración:', config);

        const resultado = await procesarInventarioWholesale(req.file.buffer, config);

        const priceLabel = config.priceBasis === '90day' ? '90day' : 'actual';
        const nombreOriginal = req.file.originalname || 'keepa_export.xlsx';
        const nombreArchivo = `analisis_wholesale_${priceLabel}_${nombreOriginal}`;
        console.log(`📤 Enviando archivo: ${nombreArchivo}`);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(nombreArchivo)}`);
        res.send(resultado.buffer);

        console.log('✅ Proceso completado exitosamente.');

    } catch (error) {
        console.error("❌ Error crítico procesando Excel:", error);
        res.status(500).json({ error: 'Ocurrió un error interno al procesar el archivo Excel: ' + error.message });
    }
});

// --------------------------------------------------------------
// 9. INICIAR SERVIDOR
// --------------------------------------------------------------
app.listen(PORT, () => {
    console.log(`🚀 Servidor corriendo en el puerto ${PORT}`);
    console.log(`📊 Modelo: gemini-3.5-flash-lite`);
    console.log(`📅 Límite diario: 1,500 solicitudes/día`);
    console.log(`📄 Nombre archivo: Incluye criterio (90day/actual)`);
    console.log(`📊 Estimación ventas: Fija con % Mejor vendedor 30 días`);
    console.log(`🎨 Colores por bloque: Gris, Azul, Azul claro, Gris`);
    console.log(`🎨 Viabilidad por fila: Verde/amarillo/rojo según resúmenes`);
    console.log(`📘 Hoja de significados: Incluida`);
    console.log(`❄️ Paneles congelados: Fila 1 y columnas A-B`);
    console.log(`🔗 ASIN clickeable: Sí (ocultando URL: Amazon)`);
    console.log(`📏 Ancho columnas: 13 (desde ASIN hasta Est. $)`);
    console.log(`📊 Orden filas: Verde → Amarillo → Rojo, por Marca y ventas`);
});
