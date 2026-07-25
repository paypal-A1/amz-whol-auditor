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
// 2. FUNCIÓN PARA LLAMAR A GEMINI CON REINTENTOS
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

// --------------------------------------------------------------
// 4. FUNCIÓN PARA DETERMINAR COLOR DE FILA
// --------------------------------------------------------------
function getColorStatus(fila) {
    if (fila['Restriction Code'] === 'NOT_ELIGIBLE') {
        return 'rojo_oscuro';
    }

    const resKeepa = fila['Resumen Keepa'] || '';
    const resIA = fila['Resumen IA'] || '';
    const statusKeepa = evaluarViabilidad(String(resKeepa));
    const statusIA = evaluarViabilidad(String(resIA));
    if (statusKeepa === 'positivo' && statusIA === 'positivo') return 'verde';
    if (statusKeepa === 'negativo' || statusIA === 'negativo') return 'rojo';
    return 'amarillo';
}

// --------------------------------------------------------------
// 5. FUNCIÓN PARA CREAR HIPERVÍNCULO
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
// 6. GENERAR DESCRIPCIÓN DE COLUMNA
// --------------------------------------------------------------
function getColumnDescription(colName, config) {
    const { roiAlto, roiMedio, roiBajo } = config;
    const descripciones = {
        'Título': 'Nombre completo del producto en Amazon',
        'ASIN': 'Amazon Standard Identification Number (clic para abrir en Amazon)',
        'Marca': 'Marca del producto (agrupador principal en el orden de filas)',
        'Restriction Code': 'Código de restricción de Amazon (ALLOWED, APPROVAL_REQUIRED, NOT_ELIGIBLE)',
        'Restriction Message': 'Mensaje detallado de la restricción proporcionado por Amazon',
        'Units Req.': 'Número de unidades que Amazon exige para aprobar la solicitud de autorización (si aplica)',
        'Compra Máx (ROI_FBM%) ($) FBM': `Precio máximo para ${roiBajo}% de ROI en logística FBM. Fórmula: Ingreso Neto FBM / (1 + ${roiBajo}/100)`,
        '% Desc. Req (ROI_FBM%) FBM': `Descuento necesario para ${roiBajo}% de ROI en logística FBM`,
        'Compra Máx (30%) ($) FBA': `Precio máximo para ${roiAlto}% de ROI en logística FBA. Fórmula: Ingreso Neto FBA / (1 + ${roiAlto}/100)`,
        '% Desc. Req (30%) FBA': `Descuento necesario para ${roiAlto}% de ROI en logística FBA`,
        'Compra Máx (20%) ($) FBA': `Precio máximo para ${roiMedio}% de ROI en logística FBA. Fórmula: Ingreso Neto FBA / (1 + ${roiMedio}/100)`,
        '% Desc. Req (20%) FBA': `Descuento necesario para ${roiMedio}% de ROI en logística FBA`,
        'Est. # Ventas Mensual': 'Unidades estimadas mensuales (orden descendente dentro de cada marca)',
        'Est. $ Ventas Mensual': 'Ingresos mensuales estimados (orden descendente dentro de cada marca, como desempate)',
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
    const compraMaxMatch = colName.match(/^Compra Máx \((\d+)%\) \(\$\) FBA$/);
    if (compraMaxMatch) {
        const roi = compraMaxMatch[1];
        return `Precio máximo para ${roi}% de ROI en logística FBA. Fórmula: Ingreso Neto FBA / (1 + ${roi}/100)`;
    }
    const descReqMatch = colName.match(/^% Desc\. Req \((\d+)%\) FBA$/);
    if (descReqMatch) {
        const roi = descReqMatch[1];
        return `Descuento necesario para ${roi}% de ROI en logística FBA`;
    }
    const compraMaxFBMMatch = colName.match(/^Compra Máx \((\d+)%\) \(\$\) FBM$/);
    if (compraMaxFBMMatch) {
        const roi = compraMaxFBMMatch[1];
        return `Precio máximo para ${roi}% de ROI en logística FBM. Fórmula: Ingreso Neto FBM / (1 + ${roi}/100)`;
    }
    const descReqFBMMatch = colName.match(/^% Desc\. Req \((\d+)%\) FBM$/);
    if (descReqFBMMatch) {
        const roi = descReqFBMMatch[1];
        return `Descuento necesario para ${roi}% de ROI en logística FBM`;
    }
    return descripciones[colName] || 'Columna generada por el sistema';
}

// --------------------------------------------------------------
// 7. FUNCIÓN PARA CREAR EL EXCEL CON EXCELJS
// --------------------------------------------------------------
async function createExcelWithStyles(filasProcesadas, config) {
    // --- Crear mapa ASIN -> URL ---
    const asinToUrl = {};
    filasProcesadas.forEach(row => {
        if (row['ASIN'] && row['URL: Amazon']) {
            asinToUrl[row['ASIN']] = row['URL: Amazon'];
        }
    });

    // --- Reordenar filas: POR MARCA PRIMERO, luego ventas ---
    const grupos = { verde: [], amarillo: [], rojo: [], rojo_oscuro: [] };
    filasProcesadas.forEach(row => {
        const color = getColorStatus(row);
        grupos[color].push(row);
    });
    
    function calcularMaxVentasPorMarca(grupo) {
        const mapa = {};
        grupo.forEach(row => {
            const marca = row['Marca'] || '';
            const ventas = parseFloat(row['Est. # Ventas Mensual']) || 0;
            if (!mapa[marca] || ventas > mapa[marca]) {
                mapa[marca] = ventas;
            }
        });
        return mapa;
    }
    
    const ordenarGrupo = (grupo) => {
        const maxVentasPorMarca = calcularMaxVentasPorMarca(grupo);
        return grupo.sort((a, b) => {
            const marcaA = a['Marca'] || '';
            const marcaB = b['Marca'] || '';
            const maxA = maxVentasPorMarca[marcaA] || 0;
            const maxB = maxVentasPorMarca[marcaB] || 0;
            if (maxA !== maxB) return maxB - maxA;
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
        ...ordenarGrupo(grupos.rojo),
        ...ordenarGrupo(grupos.rojo_oscuro)
    ];

    // --- Definir orden de columnas (NUEVO) ---
    const todasLasColumnas = Object.keys(filasOrdenadas[0] || {});
    const bloque1 = [
        'Título', 'ASIN', 'Marca', 'Restriction Code', 'Restriction Message', 'Units Req.',
        'Compra Máx (ROI_FBM%) ($) FBM', '% Desc. Req (ROI_FBM%) FBM',
        'Compra Máx (30%) ($) FBA', '% Desc. Req (30%) FBA',
        'Compra Máx (20%) ($) FBA', '% Desc. Req (20%) FBA',
        'Est. # Ventas Mensual', 'Est. $ Ventas Mensual'
    ];
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

    worksheet.columns = headers.map(col => ({
        header: col,
        key: col,
        width: (bloque2.includes(col) || bloque3.includes(col)) ? 50 :
               (col === 'Título') ? 30 :
               (col === 'Units Req.') ? 6 :
               (col.includes('FBM') || col.includes('FBA')) ? 15 : 13
    }));

    // Agregar datos
    filasOrdenadas.forEach((row) => {
        const rowData = {};
        headers.forEach(col => {
            const value = row[col] !== undefined && row[col] !== null ? row[col] : '';
            rowData[col] = value;
        });
        worksheet.addRow(rowData);
    });

    // Alto de fila
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

    // Congelar paneles: 3 columnas (Título, ASIN, Marca) y 1 fila
    worksheet.views = [{ state: 'frozen', ySplit: 1, xSplit: 3 }];

    // ---- Formatos, hipervínculos y colores ----
    for (let rowNum = 2; rowNum <= worksheet.rowCount; rowNum++) {
        const row = worksheet.getRow(rowNum);
        row.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

        const rowData = filasOrdenadas[rowNum - 2];
        const colorStatus = getColorStatus(rowData);
        let bgColor = null;
        if (colorStatus === 'verde') bgColor = 'FFC6EFCE';
        else if (colorStatus === 'rojo') bgColor = 'FFFFC7CE';
        else if (colorStatus === 'rojo_oscuro') bgColor = 'FFFF0000';
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

            // Hipervínculo a Keepa en Título
            if (colName === 'Título' && value) {
                const asin = rowData['ASIN'];
                if (asin) {
                    const keepaUrl = `https://keepa.com/#!product/1-${asin}`;
                    cell.value = { text: value, hyperlink: keepaUrl };
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
        } else if (header === 'Units Req.') {
            col.width = 6;
        } else if (header.includes('FBM') || header.includes('FBA')) {
            col.width = 15;
        } else {
            col.width = 13;
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
// FUNCIÓN PARA CONSULTAR RESTRICCIÓN INDIVIDUAL (usada por /api/restriccion)
// --------------------------------------------------------------
async function consultarRestriccionAmazon(asin) {
    try {
        const clientId = process.env.AMZ_CLIENT_ID;
        const clientSecret = process.env.AMZ_CLIENT_SECRET;
        const refreshToken = process.env.AMZ_REFRESH_TOKEN;
        const sellerId = process.env.AMZ_SELLER_ID;

        if (!clientId || !clientSecret || !refreshToken || !sellerId) {
            throw new Error('Faltan variables de entorno de Amazon');
        }

        const tokenResponse = await fetch('https://api.amazon.com/auth/o2/token', {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
                grant_type: 'refresh_token',
                refresh_token: refreshToken,
                client_id: clientId,
                client_secret: clientSecret,
            })
        });

        if (!tokenResponse.ok) {
            const errorText = await tokenResponse.text();
            throw new Error(`Error al obtener token: ${tokenResponse.status} - ${errorText}`);
        }

        const tokenData = await tokenResponse.json();
        const accessToken = tokenData.access_token;

        const marketplaceId = 'ATVPDKIKX0DER';
        const conditionType = 'new_new';
        const url = `https://sellingpartnerapi-na.amazon.com/listings/2021-08-01/restrictions?sellerId=${sellerId}&asin=${asin}&marketplaceIds=${marketplaceId}&conditionType=${conditionType}`;

        const restrictionsResponse = await fetch(url, {
            headers: {
                'x-amz-access-token': accessToken,
                'Content-Type': 'application/json'
            }
        });

        if (!restrictionsResponse.ok) {
            const errorText = await restrictionsResponse.text();
            throw new Error(`Error al consultar restricciones: ${restrictionsResponse.status} - ${errorText}`);
        }

        const restrictionsData = await restrictionsResponse.json();

        let restrictionCode = 'ALLOWED';
        let restrictionMessage = '';
        if (restrictionsData.restrictions && restrictionsData.restrictions.length > 0) {
            const reasons = restrictionsData.restrictions[0].reasons || [];
            if (reasons.length > 0) {
                restrictionCode = reasons[0].reasonCode || 'ALLOWED';
                restrictionMessage = reasons[0].message || '';
            }
        }

        return { restrictionCode, restrictionMessage };

    } catch (error) {
        console.error(`❌ Error consultando restricción para ${asin}:`, error.message);
        return { restrictionCode: 'ERROR', restrictionMessage: error.message };
    }
}

// --------------------------------------------------------------
// 8. ENDPOINT /api/restriccion (para que el frontend consulte restricciones)
// --------------------------------------------------------------
app.get('/api/restriccion', async (req, res) => {
    const asin = req.query.asin;
    if (!asin) {
        return res.status(400).json({ error: 'Falta el parámetro asin' });
    }

    try {
        const resultado = await consultarRestriccionAmazon(asin);
        res.json({
            asin: asin,
            restriction_code: resultado.restrictionCode,
            restriction_message: resultado.restrictionMessage
        });
    } catch (error) {
        console.error(`❌ Error en /api/restriccion para ${asin}:`, error.message);
        res.status(500).json({ error: error.message });
    }
});

// --------------------------------------------------------------
// 9. MOTOR PRINCIPAL DE PROCESAMIENTO (MODIFICADO: FBA + FBM)
// --------------------------------------------------------------
async function procesarInventarioWholesale(fileBuffer, config, restriccionesMap, unidadesMap) {
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

        // --- Obtener restricción desde el mapa recibido (frontend) ---
        const restData = restriccionesMap[asin] || { restriction_code: 'NO_CONSULTADO', restriction_message: '' };
        const restrictionCode = restData.restriction_code;
        const restrictionMessage = restData.restriction_message;

        // --- Obtener unidades desde el mapa recibido (frontend) ---
        const unidades = unidadesMap[asin] || '';

        // --- Procesar ventas y cálculos ---
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

        // --- Peso para envíos ---
        const pesoGramos = parseFloat(
            getColumnValue(row, [
                'Paquete: Peso (g)',
                'Weight (g)'
            ]) || 0
        );
        const pesoLibras = pesoGramos * 0.00220462;
        const costoEnvioAmazon = pesoLibras * inboundShippingPound;

        // --- Comisión de referencia (leída del Excel) ---
        let comisionReferencia = parseFloat(
            getColumnValue(row, ['% de comisión de referencia']) || 15
        );
        if (comisionReferencia > 1) {
            comisionReferencia = comisionReferencia / 100;
        }
        const referralFee = precioBuyBox * comisionReferencia;

        // --- Tarifa FBA ---
        const fbaFee = parseFloat(
            getColumnValue(row, [
                'Tarifa FBA Pick&Pack',
                'FBA Pick & Pack Fee'
            ]) || 0
        );

        // --- COSTO DE ENVÍO FBM (según tabla) ---
        let costoEnvioCliente = 0;
        if (pesoLibras > 0) {
            if (pesoLibras <= 0.5) {
                costoEnvioCliente = 4.80;
            } else if (pesoLibras <= 1.0) {
                costoEnvioCliente = 5.70;
            } else if (pesoLibras <= 2.0) {
                costoEnvioCliente = 8.50;
            } else if (pesoLibras <= 3.0) {
                costoEnvioCliente = 10.20;
            } else if (pesoLibras <= 5.0) {
                costoEnvioCliente = 13.50;
            } else if (pesoLibras <= 10.0) {
                costoEnvioCliente = 19.50;
            } else if (pesoLibras <= 15.0) {
                costoEnvioCliente = 25.00;
            } else {
                costoEnvioCliente = 25.00 + (pesoLibras - 15) * 1.20;
            }
        }

        // --- INGRESO NETO FBA (sin ROI) ---
        const ingresoNetoFBA = precioBuyBox - fbaFee - referralFee - costoEnvioAmazon - prepFee - supplierShippingUnit;

        // --- INGRESO NETO FBM (sin ROI) ---
        const ingresoNetoFBM = precioBuyBox - referralFee - costoEnvioCliente - prepFee - supplierShippingUnit;

        // --- Cálculo de compra máxima y descuento para FBA ---
        const compraMaxFBA1 = ingresoNetoFBA / (1 + roiAlto / 100);
        const descReqFBA1 = (precioBuyBox - compraMaxFBA1) / precioBuyBox;

        const compraMaxFBA2 = ingresoNetoFBA / (1 + roiMedio / 100);
        const descReqFBA2 = (precioBuyBox - compraMaxFBA2) / precioBuyBox;

        // --- Cálculo de compra máxima y descuento para FBM ---
        const compraMaxFBM = ingresoNetoFBM / (1 + roiBajo / 100);
        const descReqFBM = (precioBuyBox - compraMaxFBM) / precioBuyBox;

        // --- Est. Ventas (FBA + FBM elegibles) ---
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
            const pct = pctMejorVendedor30d;
            const ventasRestantes = ventasMensuales * (1 - pct);
            const competidoresRestantes = fbaElegibles + fbmElegibles;
            estVentasUnidades = ventasRestantes / competidoresRestantes;
        }
        const estVentasDolares = estVentasUnidades * precioBuyBox;

        // --- Construir fila ---
        const filaConMetricas = {};
        for (const key of encabezadosOriginales) {
            filaConMetricas[key] = row[key];
        }
        
        // ---- Asignar las columnas ----
        filaConMetricas['Restriction Code'] = restrictionCode;
        filaConMetricas['Restriction Message'] = restrictionMessage;
        filaConMetricas['Units Req.'] = unidades;

        // FBM
        filaConMetricas['Compra Máx (ROI_FBM%) ($) FBM'] = compraMaxFBM;
        filaConMetricas['% Desc. Req (ROI_FBM%) FBM'] = descReqFBM;

        // FBA
        filaConMetricas['Compra Máx (30%) ($) FBA'] = compraMaxFBA1;
        filaConMetricas['% Desc. Req (30%) FBA'] = descReqFBA1;
        filaConMetricas['Compra Máx (20%) ($) FBA'] = compraMaxFBA2;
        filaConMetricas['% Desc. Req (20%) FBA'] = descReqFBA2;

        filaConMetricas['Est. # Ventas Mensual'] = Math.round(estVentasUnidades);
        filaConMetricas['Est. $ Ventas Mensual'] = estVentasDolares;
        
        // Columnas de IA se llenarán después
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

    // ---- AUDITORÍA CON IA (con salto para marcas NOT_ELIGIBLE) ----
    const marcas = Object.keys(productosPorMarca);
    marcas.sort((a, b) => productosPorMarca[b].length - productosPorMarca[a].length);

    let solicitudesRealizadas = 0;
    const startTime = Date.now();
    const LIMITE_DIARIO = 1500;
    let limiteAlcanzado = false;

    for (let i = 0; i < marcas.length; i++) {
        const nombreMarca = marcas[i];
        const productos = productosPorMarca[nombreMarca];

        const todosNoElegibles = productos.every(p => {
            const restCode = p.rowRef['Restriction Code'] || '';
            return restCode === 'NOT_ELIGIBLE';
        });

        if (todosNoElegibles) {
            console.log(`⏭️ Saltando IA para marca "${nombreMarca}" (todos NOT_ELIGIBLE)`);
            continue;
        }

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

    return {
        filasProcesadas,
        config,
        solicitudesRealizadas,
        marcasProcesadas: solicitudesRealizadas,
        marcasPendientes: marcas.length - solicitudesRealizadas,
        limiteAlcanzado
    };
}

// --------------------------------------------------------------
// 10. ENDPOINT PRINCIPAL /api/audit-excel
// --------------------------------------------------------------
app.post('/api/audit-excel', upload.single('excelFile'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No se ha cargado ningún archivo Excel.' });
        }

        console.log(`📁 Archivo recibido: ${req.file.originalname} (${req.file.size} bytes)`);

        // --- Recibir mapas desde el frontend ---
        let restriccionesMap = {};
        let unidadesMap = {};

        if (req.body.restricciones) {
            try {
                restriccionesMap = JSON.parse(req.body.restricciones);
                console.log(`📦 Restricciones recibidas: ${Object.keys(restriccionesMap).length} ASINs`);
            } catch (e) {
                console.warn('⚠️ No se pudo parsear el mapa de restricciones:', e.message);
            }
        }

        if (req.body.unidades) {
            try {
                unidadesMap = JSON.parse(req.body.unidades);
                console.log(`📦 Unidades recibidas: ${Object.keys(unidadesMap).length} ASINs`);
            } catch (e) {
                console.warn('⚠️ No se pudo parsear el mapa de unidades:', e.message);
            }
        }

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

        // --- Procesar Excel ---
        const resultado = await procesarInventarioWholesale(
            req.file.buffer,
            config,
            restriccionesMap,
            unidadesMap
        );
        const { filasProcesadas } = resultado;

        // --- Generar Excel final ---
        const buffer = await createExcelWithStyles(filasProcesadas, config);

        const priceLabel = config.priceBasis === '90day' ? '90day' : 'actual';
        const nombreOriginal = req.file.originalname || 'keepa_export.xlsx';
        const nombreArchivo = `analisis_wholesale_${priceLabel}_${nombreOriginal}`;
        console.log(`📤 Enviando archivo con unidades: ${nombreArchivo}`);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(nombreArchivo)}`);
        res.send(buffer);

        console.log('✅ Proceso completado exitosamente.');

    } catch (error) {
        console.error("❌ Error crítico procesando Excel:", error);
        res.status(500).json({ error: 'Ocurrió un error interno al procesar el archivo Excel: ' + error.message });
    }
});

// --------------------------------------------------------------
// 11. INICIAR SERVIDOR
// --------------------------------------------------------------
app.listen(PORT, () => {
    console.log(`🚀 Servidor corriendo en el puerto ${PORT}`);
    console.log(`📊 Modelo: gemini-3.5-flash-lite`);
    console.log(`📅 Límite diario: 1,500 solicitudes/día`);
    console.log(`📄 Nombre archivo: Incluye criterio (90day/actual)`);
    console.log(`📊 Estimación ventas: Fija con % Mejor vendedor 30 días`);
    console.log(`🎨 Colores por bloque: Gris, Azul, Azul claro, Gris`);
    console.log(`🎨 Viabilidad por fila: Verde/amarillo/rojo/rojo oscuro según restricción y resúmenes`);
    console.log(`📘 Hoja de significados: Incluida`);
    console.log(`❄️ Paneles congelados: Fila 1 y columnas A-B-C (Título, ASIN, Marca)`);
    console.log(`🔗 ASIN clickeable: Sí (ocultando URL: Amazon)`);
    console.log(`🔗 Título clickeable a Keepa: Sí`);
    console.log(`📏 Ancho columnas: 13 (estándar), 15 para FBA/FBM, 6 para Units Req.`);
    console.log(`📊 Orden filas: Verde → Amarillo → Rojo → Rojo oscuro, por Marca → Ventas (desc) → Dinero (desc)`);
    console.log(`📦 Columnas FBA y FBM integradas.`);
});
