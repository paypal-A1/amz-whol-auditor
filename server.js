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


// Permitir CORS para desarrollo local
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Content-Type');
    if (req.method === 'OPTIONS') {
        return res.sendStatus(200);
    }
    next();
});
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
// 3. FUNCIÓN PARA EVALUAR VIABILIDAD (MODIFICADA)
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
// 4. FUNCIÓN PARA DETERMINAR COLOR DE FILA (MODIFICADA)
// --------------------------------------------------------------
function getColorStatus(fila) {
    // 1. NOT_ELIGIBLE → rojo oscuro (prioridad máxima)
    if (fila['Restriction Code'] === 'NOT_ELIGIBLE') {
        return 'rojo_oscuro';
    }

    // 2. Si % Desc. Req (FBA 30%) > 70% → rojo (inviable por alto descuento)
    const descReq30 = parseFloat(fila['% Desc. Req (30%) FBA']) || 0;
    if (descReq30 > 0.70) {
        return 'rojo';
    }

    // 3. Si no, evaluar los resúmenes de IA (como antes)
    const resCuantitativo = fila['Resumen IA Cuantitativo'] || '';
    const resCualitativo = fila['Resumen IA Cualitativo'] || '';
    const statusCuantitativo = evaluarViabilidad(String(resCuantitativo));
    const statusCualitativo = evaluarViabilidad(String(resCualitativo));
    if (statusCuantitativo === 'positivo' && statusCualitativo === 'positivo') return 'verde';
    if (statusCuantitativo === 'negativo' || statusCualitativo === 'negativo') return 'rojo';
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
// 6. GENERAR DESCRIPCIÓN DE COLUMNA (MODIFICADA)
// --------------------------------------------------------------
function getColumnDescription(colName, config) {
    const { roiAlto, roiMedio, roiBajo } = config;
    const descripciones = {
        // ---- Columnas principales ----
        'Título': 'Nombre completo del producto en Amazon',
        'ASIN': 'Amazon Standard Identification Number (clic para abrir en Amazon)',
        'Marca': 'Marca del producto (agrupador principal en el orden de filas)',
        'Restriction Code': 'Código de restricción de Amazon (ALLOWED, APPROVAL_REQUIRED, NOT_ELIGIBLE)',
        'Restriction Message': 'Mensaje detallado de la restricción proporcionado por Amazon',
        'Units Req.': 'Número de unidades que Amazon exige para aprobar la solicitud de autorización (si aplica)',
        'Break-Even ($)': 'Punto de equilibrio (0% ROI). Fórmula: Precio Buy Box - FBA - Comisión - Envío - Prep',
        'Compra Máx (ROI_FBM%) ($) FBM': `Precio máximo para ${roiBajo}% de ROI en logística FBM. Fórmula: Ingreso Neto FBM / (1 + ${roiBajo}/100)`,
        '% Desc. Req (ROI_FBM%) FBM': `Descuento necesario para ${roiBajo}% de ROI en logística FBM`,
        'Compra Máx (30%) ($) FBA': `Precio máximo para ${roiAlto}% de ROI en logística FBA. Fórmula: Ingreso Neto FBA / (1 + ${roiAlto}/100)`,
        '% Desc. Req (30%) FBA': `Descuento necesario para ${roiAlto}% de ROI en logística FBA`,
        'Compra Máx (20%) ($) FBA': `Precio máximo para ${roiMedio}% de ROI en logística FBA. Fórmula: Ingreso Neto FBA / (1 + ${roiMedio}/100)`,
        '% Desc. Req (20%) FBA': `Descuento necesario para ${roiMedio}% de ROI en logística FBA`,
        'Est. # Ventas Mensual': 'Unidades estimadas mensuales (orden descendente dentro de cada marca)',
        'Est. $ Ventas Mensual': 'Ingresos mensuales estimados (orden descendente dentro de cada marca, como desempate)',
        'Resumen IA Cuantitativo': 'Análisis basado en datos numéricos de Keepa y cálculos financieros. Evalúa demanda, competencia, márgenes y viabilidad logística. Comienza con ✅ ⚠️ ❌',
        'Resumen IA Cualitativo': 'Análisis basado en investigación web de la marca. Evalúa programa wholesale, contactos, requisitos, riesgo IP y estrategia. Comienza con ✅ ⚠️ ❌',
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
        'Conclusión General': 'Análisis integral combinando Keepa, cálculos e investigación de IA',

        // ---- Columnas de Keepa (mejoradas) ----
        'Clasificación de Ventas: Actual': 'Posición actual del producto en el ranking de ventas de su categoría (Best Sellers Rank). Cuanto más bajo sea el número, mejores ventas.',
        'Clasificación de Ventas: Promedio de 90 días': 'Promedio del BSR en los últimos 90 días. Indica la estabilidad de la demanda; si es similar al actual, la demanda es constante.',
        'Clasificación de Ventas: Descensos en los últimos 30 días': 'Número de veces que el producto ha empeorado su BSR en los últimos 30 días. Un valor alto puede indicar caída de demanda.',
        'Tendencias de ventas mensuales: Comprados el mes pasado': 'Unidades totales vendidas en el último mes según datos de Keepa (aproximación basada en BSR histórico).',
        'Opiniones: Cantidad de valoraciones': 'Número total de reseñas (reviews) del producto. Una cantidad alta suele indicar producto consolidado y confiable.',
        'Caja de Compra: Actual': 'Precio actual del vendedor que tiene la Buy Box. Es el precio de referencia para tus cálculos de rentabilidad.',
        'Caja de Compra: Promedio de 30 días': 'Precio promedio de la Buy Box en los últimos 30 días. Ayuda a ver la tendencia de precio a corto plazo.',
        'Caja de Compra: Promedio de 90 días': 'Precio promedio de la Buy Box en los últimos 90 días. Muestra la estabilidad del precio a mediano plazo.',
        'Caja de Compra: Promedio de 180 días': 'Precio promedio de la Buy Box en los últimos 180 días. Útil para evaluar si el precio ha subido o bajado significativamente.',
        'Caja de Compra: Vendedor Caja de Compra': 'Nombre o Seller ID del vendedor que actualmente gana la Buy Box. Si es Amazon, indica alta competencia directa.',
        'Caja de Compra: % Amazon 30 días': 'Porcentaje del tiempo (en los últimos 30 días) que Amazon ha tenido la Buy Box. Si es alto (>50%), Amazon compite directamente contigo.',
        'Caja de Compra: % Amazon 90 días': 'Porcentaje del tiempo (en los últimos 90 días) que Amazon ha tenido la Buy Box.',
        'Caja de Compra: % Mejor vendedor 30 días': 'Porcentaje de la Buy Box controlado por el vendedor principal en los últimos 30 días. Se usa para calcular el reparto de ventas estimadas.',
        'Caja de Compra: % Mejor vendedor 90 días': 'Porcentaje de la Buy Box controlado por el vendedor principal en los últimos 90 días.',
        'Caja de Compra: Es FBA': 'Indica si el vendedor que gana la Buy Box utiliza logística FBA (Sí/No). Clave para saber si tu competencia directa es FBA o FBM.',
        'Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBA': 'Número de vendedores con ofertas en estado "Nuevo" que usan FBA y son elegibles para competir por la Buy Box. Se usa para repartir ventas estimadas.',
        'Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBM': 'Número de vendedores con ofertas en estado "Nuevo" que usan FBM (envío por vendedor) y son elegibles para la Buy Box.',
        'Amazon: Actual': 'Precio actual de la oferta de Amazon (si Amazon vende el producto). Si aparece, es tu competidor directo.',
        'Amazon: Promedio de 30 días': 'Precio promedio de Amazon en los últimos 30 días.',
        'Amazon: Promedio de 90 días': 'Precio promedio de Amazon en los últimos 90 días.',
        'Tarifa FBA Pick&Pack': 'Tarifa que cobra Amazon por preparar y enviar el producto (Pick & Pack Fee). Se usa para calcular el costo FBA.',
        '% de comisión de referencia': 'Comisión que Amazon cobra sobre el precio de venta (referral fee). Específica de la categoría del producto (ej. 15% en juguetes, 12% en electrónica).',
        'Recuento total de Ofertas': 'Número total de vendedores que ofrecen el producto (suma de FBA + FBM + otros).',
        'Recuento ofertas nuevas FBA: Actual': 'Cantidad actual de vendedores FBA con ofertas en estado "Nuevo".',
        'Recuento ofertas nuevas FBM: Actual': 'Cantidad actual de vendedores FBM con ofertas en estado "Nuevo".',
        'Códigos de producto: UPC': 'Código Universal de Producto (UPC/EAN). Identificador único del producto.',
        'Paquete: Dimensión (cm³)': 'Volumen del paquete en centímetros cúbicos (largo × ancho × alto). Se usa para determinar si el producto es oversize (afecta tarifas FBA).',
        'Paquete: Peso (g)': 'Peso del paquete en gramos. Se convierte a libras para calcular el costo de envío a Amazon (inbound shipping) y el costo de envío FBM.',
        'Es HazMat': 'Indica si el producto está clasificado como material peligroso (Hazmat). Si es "Sí", tiene restricciones adicionales de almacenamiento y envío.',
        'Es sensible al calor': 'Indica si el producto es sensible al calor (meltable). Si es "Sí", Amazon restringe su almacenamiento en ciertas épocas (afecta disponibilidad FBA).',
        'Producto para adultos': 'Indica si el producto está clasificado como para adultos (Adult). Puede limitar la visibilidad y requerir aprobación especial.'
    };

    // ---- Manejo de columnas dinámicas (Compra Máx y % Desc. Req) ----
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
// 7. FUNCIÓN PARA CREAR EL EXCEL CON EXCELJS (MODIFICADA)
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

    // --- Definir orden de columnas (MODIFICADO) ---
    const todasLasColumnas = Object.keys(filasOrdenadas[0] || {});
    const bloque1 = [
        'Título', 'ASIN', 'Marca', 'Restriction Code', 'Restriction Message', 'Units Req.',
        `Compra Máx (${config.roiBajo}%) ($) FBM`,
        `% Desc. Req (${config.roiBajo}%) FBM`,
        `Compra Máx (${config.roiAlto}%) ($) FBA`,
        `% Desc. Req (${config.roiAlto}%) FBA`,
        `Compra Máx (${config.roiMedio}%) ($) FBA`,
        `% Desc. Req (${config.roiMedio}%) FBA`,
        'Est. # Ventas Mensual', 'Est. $ Ventas Mensual'
    ];
    const bloque2 = ['Resumen IA Cuantitativo', 'Resumen IA Cualitativo'];
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

    // ---- HOJA DE PRODUCTOS NOT_ELIGIBLE ----
    const notEligibleBrands = filasProcesadas
        .filter(row => row['Restriction Code'] === 'NOT_ELIGIBLE')
        .map(row => row['Marca'])
        .filter(marca => marca && marca !== '')
        .map(marca => marca.trim());
    
    const uniqueBrands = [...new Set(notEligibleBrands)];
    const notEligibleText = uniqueBrands.map(m => `-${m}`).join(' ');
    
    const notEligibleSheet = workbook.addWorksheet('productos NOT_ELIGIBLE', {
        properties: { tabColor: { argb: 'FFFF0000' } }
    });
    notEligibleSheet.getCell('A1').value = notEligibleText;
    notEligibleSheet.getCell('A1').alignment = { wrapText: true, vertical: 'middle' };
    notEligibleSheet.getColumn('A').width = 80;
    
    worksheet.columns = headers.map(col => ({
        header: col,
        key: col,
        width: (bloque2.includes(col) || bloque3.includes(col)) ? 50 :
               (col === 'Título') ? 30 :
               (col === 'Units Req.') ? 6 :
               (col.includes('Compra Máx') || col.includes('% Desc. Req') || col.includes('Est. # Ventas Mensual') || col.includes('Est. $ Ventas Mensual')) ? 10 : 12
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

            // Hipervínculo a Seller Central en Restriction Code
            if (colName === 'Restriction Code' && value) {
                const asin = rowData['ASIN'];
                if (asin) {
                    const sellerUrl = `https://sellercentral.amazon.com/hz/approvalrequest/restrictions/approve?asin=${asin}`;
                    cell.value = { text: value, hyperlink: sellerUrl };
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
        } else if (header.includes('Compra Máx') || header.includes('% Desc. Req') || header.includes('Est. # Ventas Mensual') || header.includes('Est. $ Ventas Mensual')) {
            col.width = 11;
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



// ============================================================
// NUEVO ENDPOINT UNIFICADO: /api/product-details-batch
// Devuelve: restriction, hazmat, size_tier, amazon_in_buybox,
// fba_count, fbm_count, y métricas financieras (compra/dto)
// ============================================================
app.get('/api/product-details-batch', async (req, res) => {
    const asinsParam = req.query.asins;
    if (!asinsParam) {
        return res.status(400).json({ error: 'Falta el parámetro asins' });
    }

    const asins = asinsParam.split(',').map(a => a.trim().toUpperCase()).filter(Boolean);
    if (asins.length === 0) {
        return res.status(400).json({ error: 'No hay ASINs válidos' });
    }
    if (asins.length > 30) {
        return res.status(400).json({ error: 'Máximo 30 ASINs por llamada' });
    }

    console.log(`📦 Consultando detalles para ${asins.length} ASINs`);

    // Configuración de ROIs (igual que en tu frontend)
    const ROI_FBM = 20;   // roiBajo
    const ROI_FBA_ALTO = 30; // roiAlto
    const ROI_FBA_MEDIO = 20; // roiMedio

    try {
        const resultados = {};

        // Procesar en lotes de 5 para no saturar la API de Amazon
        const batchSize = 10;
        for (let i = 0; i < asins.length; i += batchSize) {
            const batch = asins.slice(i, i + batchSize);
            const promesas = batch.map(async (asin) => {
                try {
                    const resultado = await consultarDetallesProducto(asin, {
                        roiFBM: ROI_FBM,
                        roiFBAAlto: ROI_FBA_ALTO,
                        roiFBAMedio: ROI_FBA_MEDIO
                    });
                    return { asin, resultado };
                } catch (error) {
                    console.error(`❌ Error en ${asin}:`, error.message);
                    return { asin, resultado: { error: error.message } };
                }
            });

            const respuestas = await Promise.all(promesas);
            respuestas.forEach(({ asin, resultado }) => {
                resultados[asin] = resultado;
            });

            // Pausa entre lotes para respetar rate limits
            if (i + batchSize < asins.length) {
                await new Promise(resolve => setTimeout(resolve, 500));
            }
        }

        res.json({
            success: true,
            total: asins.length,
            resultados
        });

    } catch (error) {
        console.error('❌ Error en /api/product-details-batch:', error);
        res.status(500).json({ error: 'Error interno' });
    }
});

// ============================================================
// FUNCIÓN PARA CONSULTAR TODOS LOS DETALLES DE UN PRODUCTO
// ============================================================
// ============================================================
// FUNCIÓN PARA CONSULTAR TODOS LOS DETALLES DE UN PRODUCTO
// ============================================================

async function consultarDetallesProducto(asin, rois) {
    const { roiFBM, roiFBAAlto, roiFBAMedio } = rois;
    const resultado = {};

    // ---- CONFIGURACIÓN (cambia esto según necesites) ----
    const SALTAR_NOT_ELIGIBLE = true;    // true = no consultar
    const SALTAR_APPROVAL_REQUIRED = true; // true = no consultar (cambia a false para consultar)
    const SALTAR_ALLOWED = false;         // false = consultar siempre

    // ---- 1. RESTRICCIÓN ----
    let restrictionCode = '';
    let restrictionMessage = '';
    try {
        const restriccion = await consultarRestriccionAmazon(asin);
        restrictionCode = restriccion.restrictionCode;
        restrictionMessage = restriccion.restrictionMessage;
        resultado.restriction_code = restrictionCode;
        resultado.restriction_message = restrictionMessage;
    } catch (e) {
        resultado.restriction_code = 'ERROR';
        resultado.restriction_message = e.message;
        restrictionCode = 'ERROR';
    }

    // ---- 2. DECIDIR SI SALTAMOS CONSULTAS SEGÚN LA CONFIGURACIÓN ----
    const saltarConsultas = 
        (restrictionCode === 'NOT_ELIGIBLE' && SALTAR_NOT_ELIGIBLE) ||
        (restrictionCode === 'APPROVAL_REQUIRED' && SALTAR_APPROVAL_REQUIRED) ||
        (restrictionCode === 'ALLOWED' && SALTAR_ALLOWED);

    if (saltarConsultas) {
        // Asignar valores por defecto (vacíos o cero)
        resultado.hazmat = false;
        resultado.size_tier = 'STANDARD';
        resultado.peso_libras = 0;
        resultado.precio_lista = 0;
        resultado.bsr = null;
        resultado.brand = '';
        resultado.buy_box_price = 0;
        resultado.amazon_in_buybox = false;
        resultado.fba_count = 0;
        resultado.fbm_count = 0;
        resultado.fba_eligible_count = 0;
        resultado.fbm_eligible_count = 0;
        resultado.compra_max_fbm = 0;
        resultado.desc_req_fbm = 0;
        resultado.compra_max_fba_alto = 0;
        resultado.desc_req_fba_alto = 0;
        resultado.compra_max_fba_medio = 0;
        resultado.desc_req_fba_medio = 0;
        return resultado;
    }

    // ---- 3. HAZMAT y DIMENSIONES (desde Catalog API) ----
    try {
        const catalogData = await consultarCatalogoAmazon(asin);
        resultado.hazmat = catalogData.hazmat || false;
        resultado.size_tier = catalogData.sizeTier || 'STANDARD';
        resultado.peso_libras = catalogData.pesoLibras || 0;
        resultado.precio_lista = catalogData.listPrice || 0;
        resultado.bsr = catalogData.bsr || null;
        resultado.brand = catalogData.brand || '';
    } catch (e) {
        resultado.hazmat = false;
        resultado.size_tier = 'UNKNOWN';
        resultado.peso_libras = 0;
        resultado.precio_lista = 0;
        resultado.bsr = null;
        resultado.brand = '';
    }

    // ---- 4. COMPETENCIA (precios, ofertas, Buy Box) ----
    try {
        const competencia = await consultarCompetenciaAmazon(asin);
        resultado.buy_box_price = competencia.buyBoxPrice || 0;
        resultado.amazon_in_buybox = competencia.amazonInBuybox || false;
        resultado.fba_count = competencia.fbaCount || 0;
        resultado.fbm_count = competencia.fbmCount || 0;
        resultado.fba_eligible_count = competencia.fbaEligibleCount || 0;
        resultado.fbm_eligible_count = competencia.fbmEligibleCount || 0;
    } catch (e) {
        console.warn(`⚠️ Error en consultarCompetenciaAmazon para ${asin}:`, e.message);
        if (e.message.includes('429')) {
            console.warn(`❌ Error en consultarCompetenciaAmazon para ${asin}: HTTP 429`);
        }
        resultado.buy_box_price = 0;
        resultado.amazon_in_buybox = false;
        resultado.fba_count = 0;
        resultado.fbm_count = 0;
        resultado.fba_eligible_count = 0;
        resultado.fbm_eligible_count = 0;
    }

    //keepa
        // ---- CONFIGURACIÓN DE SALTOS (ajústala según necesites) ----
    //const SALTAR_NOT_ELIGIBLE = true;       // true = no consultar Keepa para NOT_ELIGIBLE
    //const SALTAR_APPROVAL_REQUIRED = true;  // true = no consultar Keepa para APPROVAL_REQUIRED
    //const SALTAR_ALLOWED = false;           // false = consultar siempre para ALLOWED

    //let restrictionCode = resultado.restriction_code || '';

    // Decidir si consultar Keepa según la restricción
    let debeConsultarKeepa = false;   // <--- CAMBIA ESTE NOMBRE


    if (restrictionCode === 'ALLOWED' && !SALTAR_ALLOWED) {
        debeConsultarKeepa = true;
    } else if (restrictionCode === 'APPROVAL_REQUIRED' && !SALTAR_APPROVAL_REQUIRED) {
        debeConsultarKeepa = true;
    } else if (restrictionCode === 'NOT_ELIGIBLE' && !SALTAR_NOT_ELIGIBLE) {
        debeConsultarKeepa = true;
    } else if (restrictionCode === 'ERROR' || restrictionCode === '') {
        // Si la restricción falló, puedes decidir si consultar o no
        debeConsultarKeepa = true; // Cambia a false si prefieres no gastar tokens
    }

    
    // ---- 5. KEEPA (Buy Box percentages) ----
    if (debeConsultarKeepa) {    // <--- USA EL NUEVO NOMBRE AQUÍ
        try {
            console.log(`🔍 Llamando a Keepa para ${asin}...`);
            const keepaData = await consultarKeepa(asin);
            if (keepaData) {
                resultado.amazon_percentage = keepaData.amazon_percentage;  // en lugar de amazon_buybox_percentage
                //resultado.amazon_buybox_percentage = keepaData.amazon_percentage;
                resultado.best_seller_percentage = keepaData.best_seller_percentage;
                resultado.keepa_tokens_left = keepaData.tokens_left;
            } else {
                resultado.amazon_percentage = null;
                resultado.best_seller_percentage = null;
            }
        } catch (e) {
            console.warn(`⚠️ Error consultando Keepa para ${asin}:`, e.message);
            resultado.amazon_percentage = null;
            resultado.best_seller_percentage = null;
        }
    } else {
        resultado.amazon_percentage = null;
        resultado.best_seller_percentage = null;
    }




    
    // ---- 5. CÁLCULOS FINANCIEROS (solo si hay precio) ----
    const precioBuyBox = resultado.buy_box_price;
    const pesoLibras = resultado.peso_libras || 0;
    const referralFee = precioBuyBox * 0.15;
    const prepFee = 1.50;
    const inboundShippingPound = 1.00;
    const supplierShippingUnit = 0.00;

    let costoEnvioCliente = 0;
    if (pesoLibras > 0) {
        if (pesoLibras <= 0.5) costoEnvioCliente = 4.80;
        else if (pesoLibras <= 1.0) costoEnvioCliente = 5.70;
        else if (pesoLibras <= 2.0) costoEnvioCliente = 8.50;
        else if (pesoLibras <= 3.0) costoEnvioCliente = 10.20;
        else if (pesoLibras <= 5.0) costoEnvioCliente = 13.50;
        else if (pesoLibras <= 10.0) costoEnvioCliente = 19.50;
        else if (pesoLibras <= 15.0) costoEnvioCliente = 25.00;
        else costoEnvioCliente = 25.00 + (pesoLibras - 15) * 1.20;
    }

    const fbaFee = pesoLibras <= 1 ? 3.50 : 5.50;
    const costoEnvioAmazon = pesoLibras * inboundShippingPound;
    const ingresoNetoFBA = precioBuyBox - fbaFee - referralFee - costoEnvioAmazon - prepFee - supplierShippingUnit;
    const ingresoNetoFBM = precioBuyBox - referralFee - costoEnvioCliente - prepFee - supplierShippingUnit;

    resultado.compra_max_fbm = Math.max(0, ingresoNetoFBM / (1 + roiFBM / 100));
    resultado.desc_req_fbm = precioBuyBox > 0 ? (precioBuyBox - resultado.compra_max_fbm) / precioBuyBox : 0;

    resultado.compra_max_fba_alto = Math.max(0, ingresoNetoFBA / (1 + roiFBAAlto / 100));
    resultado.desc_req_fba_alto = precioBuyBox > 0 ? (precioBuyBox - resultado.compra_max_fba_alto) / precioBuyBox : 0;

    resultado.compra_max_fba_medio = Math.max(0, ingresoNetoFBA / (1 + roiFBAMedio / 100));
    resultado.desc_req_fba_medio = precioBuyBox > 0 ? (precioBuyBox - resultado.compra_max_fba_medio) / precioBuyBox : 0;

    return resultado;
}

// ============================================================
// FUNCIÓN REAL PARA CONSULTAR CATÁLOGO (Hazmat, dimensiones, BSR, BRAND)
// ============================================================
async function consultarCatalogoAmazon(asin) {
    try {
        const clientId = process.env.AMZ_CLIENT_ID;
        const clientSecret = process.env.AMZ_CLIENT_SECRET;
        const refreshToken = process.env.AMZ_REFRESH_TOKEN;

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
            throw new Error('Error al obtener token de acceso');
        }

        const tokenData = await tokenResponse.json();
        const accessToken = tokenData.access_token;

        const marketplaceId = 'ATVPDKIKX0DER';
        const url = `https://sellingpartnerapi-na.amazon.com/catalog/2022-04-01/items/${asin}?marketplaceIds=${marketplaceId}&includedData=summaries,dimensions`;

        const response = await fetch(url, {
            headers: {
                'x-amz-access-token': accessToken,
                'Content-Type': 'application/json'
            }
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();
        //console.log(`📦 CATALOG RAW para ${asin}:`, JSON.stringify(data, null, 2)); // <-- AQUÍ


        // const item = data.items?.[0] || {};  // <-- ELIMINA ESTA LÍNEA
        const attributes = data.attributes || {};
        const dimensions = data.dimensions?.[0] || {};
        const summaries = data.summaries?.[0] || {};

        

        // ✅ BRAND (extraída correctamente desde summaries o attributes)
        //const brand = summaries.brand || attributes.brand?.[0]?.value || '';
        //const brand = attributes.brand?.[0]?.value || '';
        
        // Extraer marca
        const brand = summaries.brand || summaries.manufacturer || '';

        // Hazmat
        let hazmat = false;
        if (attributes.supplier_declared_dg_hz_regulation) {
            hazmat = attributes.supplier_declared_dg_hz_regulation[0]?.value !== 'not_applicable';
        }

        // Peso en libras
        const pesoKg = dimensions.package?.weight?.value || 0;
        const pesoLibras = pesoKg * 2.20462;

        // Dimensiones
        const alto = dimensions.package?.height?.value || 0;
        const ancho = dimensions.package?.width?.value || 0;
        const largo = dimensions.package?.length?.value || 0;
        const mayorDimension = Math.max(alto, ancho, largo);

        // Size Tier
        let sizeTier = 'STANDARD';
        if (pesoLibras > 20 || mayorDimension > 18) {
            sizeTier = 'OVERSIZE';
        } else if (pesoLibras > 10 || mayorDimension > 15) {
            sizeTier = 'OVERSIZE';
        }

        const bsr = summaries.salesRank?.[0]?.rank || 0;
        const listPrice = summaries.listPrice?.Amount || 0;

        console.log(`🔍 Catalog data for ${asin}: brand = "${brand}"`);

        return {
            hazmat,
            sizeTier,
            pesoLibras,
            listPrice,
            bsr,
            brand  // ✅ AHORA SÍ DEVUELVE LA MARCA
        };

    } catch (error) {
        console.error(`❌ Error en consultarCatalogoAmazon para ${asin}:`, error.message);
        console.log(`🔍 Catalog data for ${asin}: brand = "${brand}"`);
        return {
            hazmat: false,
            sizeTier: 'UNKNOWN',
            pesoLibras: 0,
            listPrice: 0,
            bsr: 0,
            brand: ''  // ✅ También incluir brand vacío en el error
        };
    }
}

// ============================================================
// FUNCIÓN REAL PARA CONSULTAR COMPETENCIA (precios, ofertas, Buy Box)
// ============================================================
async function consultarCompetenciaAmazon(asin) {
    try {
        const clientId = process.env.AMZ_CLIENT_ID;
        const clientSecret = process.env.AMZ_CLIENT_SECRET;
        const refreshToken = process.env.AMZ_REFRESH_TOKEN;

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
            throw new Error('Error al obtener token de acceso');
        }

        const tokenData = await tokenResponse.json();
        const accessToken = tokenData.access_token;

        const marketplaceId = 'ATVPDKIKX0DER';
        const url = `https://sellingpartnerapi-na.amazon.com/products/pricing/v0/items/${asin}/offers?ItemCondition=New&MarketplaceId=${marketplaceId}`;

        const response = await fetch(url, {
            headers: {
                'x-amz-access-token': accessToken,
                'Content-Type': 'application/json'
            }
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();
        const payload = data.payload || {};

        // Precio de Buy Box
        let buyBoxPrice = 0;
        const buyBoxPrices = payload.Summary?.BuyBoxPrices || [];
        if (buyBoxPrices.length > 0) {
            buyBoxPrice = buyBoxPrices[0]?.LandedPrice?.Amount || 0;
        }

        // ¿Amazon tiene la Buy Box?
        let amazonInBuybox = false;
        const offers = payload.Offers || [];
        for (const offer of offers) {
            if (offer.IsBuyBoxWinner && offer.IsFulfilledByAmazon) {
                amazonInBuybox = true;
                break;
            }
        }

        // Conteo de ofertas FBA y FBM
        let fbaCount = 0;
        let fbmCount = 0;
        const numberOfOffers = payload.Summary?.NumberOfOffers || [];
        for (const offer of numberOfOffers) {
            if (offer.condition === 'new') {
                if (offer.fulfillmentChannel === 'Amazon') {
                    fbaCount = offer.OfferCount || 0;
                } else if (offer.fulfillmentChannel === 'Merchant') {
                    fbmCount = offer.OfferCount || 0;
                }
            }
        }

        // Ofertas elegibles para Buy Box
        let fbaEligibleCount = 0;
        let fbmEligibleCount = 0;
        const eligibleOffers = payload.Summary?.BuyBoxEligibleOffers || [];
        for (const offer of eligibleOffers) {
            if (offer.condition === 'new') {
                if (offer.fulfillmentChannel === 'Amazon') {
                    fbaEligibleCount = offer.OfferCount || 0;
                } else if (offer.fulfillmentChannel === 'Merchant') {
                    fbmEligibleCount = offer.OfferCount || 0;
                }
            }
        }

        return {
            buyBoxPrice,
            amazonInBuybox,
            fbaCount,
            fbmCount,
            fbaEligibleCount,
            fbmEligibleCount
        };

    } catch (error) {
        console.error(`❌ Error en consultarCompetenciaAmazon para ${asin}:`, error.message);
        return {
            buyBoxPrice: 0,
            amazonInBuybox: false,
            fbaCount: 0,
            fbmCount: 0,
            fbaEligibleCount: 0,
            fbmEligibleCount: 0
        };
    }
}








//keepa

// ============================================================
// FUNCIÓN PARA CONSULTAR KEEPA (Buy Box %)
// ============================================================
async function consultarKeepa(asin) {
    const apiKey = process.env.KEEPA_API_KEY;
    //console.log(`🔑 KEEPA_API_KEY existe? ${!!process.env.KEEPA_API_KEY}`);
    if (!apiKey) {
        console.warn(`⚠️ KEEPA_API_KEY no configurada para ${asin}`);
        return null;
    }

    try {
        const url = `https://api.keepa.com/product?key=${apiKey}&domain=1&asin=${asin}&stats=90&buybox=1`;
        
        const response = await fetch(url, {
            method: 'GET',
            headers: { 'Content-Type': 'application/json' },
            timeout: 30000
        });

        if (!response.ok) {
            const errorText = await response.text();
            console.warn(`⚠️ Keepa error para ${asin}: ${response.status} - ${errorText}`);
            return null;
        }

        
        const data = await response.json();
        console.log(`📡 Keepa response for ${asin}:`, JSON.stringify(data).substring(0, 200));

        if (data.tokensLeft !== undefined && data.tokensLeft < 3) {
            console.warn(`⚠️ Keepa tokens bajos (${data.tokensLeft}) para ${asin}`);
            return null;
        }

        const product = data.products?.[0];
        console.log(`📦 Producto Keepa para ${asin}:`, JSON.stringify(product).substring(0, 500));
        if (!product || !product.buyBox || !Array.isArray(product.buyBox)) {
            console.log(`⚠️ Keepa sin BuyBox para ${asin}`);
            return null;
        }

        const AMAZON_SELLER_ID_US = 'ATVPDKIKX0DER';

        // Obtener el objeto buyBoxStats (si existe)
        const buyBoxStats = product.buyBoxStats || {};
        let amazonPercentage = 0;
        let bestSellerPercentage = 0;
        let bestSellerId = null;
        
        // Recorrer todos los vendedores en buyBoxStats
        for (const [sellerId, stats] of Object.entries(buyBoxStats)) {
            const percentage = parseFloat(stats.percentageWon) || 0;
            // Actualizar el mejor vendedor
            if (percentage > bestSellerPercentage) {
                bestSellerPercentage = percentage;
                bestSellerId = sellerId;
            }
            // Detectar si es Amazon
            if (sellerId === 'ATVPDKIKX0DER') {
                amazonPercentage = percentage;
            }
        }
        
        // Si Amazon no aparece en buyBoxStats pero buyBoxIsAmazon es true, entonces Amazon tiene el 100%
        if (amazonPercentage === 0 && product.buyBoxIsAmazon === true) {
            amazonPercentage = 100;
            // Si además no había mejor vendedor, asignar Amazon como mejor
            if (bestSellerPercentage === 0) {
                bestSellerPercentage = 100;
                bestSellerId = 'ATVPDKIKX0DER';
            }
        }
        
        return {
            amazon_percentage: amazonPercentage,
            best_seller_percentage: bestSellerPercentage,
            best_seller_id: bestSellerId,
            tokens_left: data.tokensLeft || 0
        };

        return {
            amazon_percentage: amazonPercentage,
            best_seller_percentage: bestSellerPercentage,
            best_seller_id: bestSellerId,
            tokens_left: data.tokensLeft || 0
        };

    } catch (error) {
        console.error(`❌ Error en consultarKeepa para ${asin}:`, error.message);
        return null;
    }
}








// --------------------------------------------------------------
// 9. MOTOR PRINCIPAL DE PROCESAMIENTO (MODIFICADO: NUEVOS RESUMENES IA + saltarIA)
// --------------------------------------------------------------
async function procesarInventarioWholesale(fileBuffer, config, restriccionesMap, unidadesMap, saltarIA = false) {
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

        // --- Precio promedio 90 días para volatilidad ---
        const precioPromedio90 = parseFloat(
            getColumnValue(row, [
                'Caja de Compra: Promedio de 90 días',
                'Buy Box: 90 days avg',
                'Amazon 90 days avg'
            ]) || 0
        );

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
        
        // FBM (usando config.roiBajo)
        filaConMetricas[`Compra Máx (${config.roiBajo}%) ($) FBM`] = compraMaxFBM;
        filaConMetricas[`% Desc. Req (${config.roiBajo}%) FBM`] = descReqFBM;
        
        // FBA (usando config.roiAlto y config.roiMedio)
        filaConMetricas[`Compra Máx (${config.roiAlto}%) ($) FBA`] = compraMaxFBA1;
        filaConMetricas[`% Desc. Req (${config.roiAlto}%) FBA`] = descReqFBA1;
        filaConMetricas[`Compra Máx (${config.roiMedio}%) ($) FBA`] = compraMaxFBA2;
        filaConMetricas[`% Desc. Req (${config.roiMedio}%) FBA`] = descReqFBA2;
        
        filaConMetricas['Est. # Ventas Mensual'] = Math.round(estVentasUnidades);
        filaConMetricas['Est. $ Ventas Mensual'] = estVentasDolares;
        
        // Columnas de IA se llenarán después (si no se salta la IA)
        filaConMetricas['Resumen IA Cuantitativo'] = '';
        filaConMetricas['Resumen IA Cualitativo'] = '';
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

    // ---- AUDITORÍA CON IA (SALTAR SI saltarIA es true) ----
    if (!saltarIA) {
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
                // --- DATOS CUANTITATIVOS (por producto) ---
                const productosInfo = productos.map(p => {
                    const row = p.rowRef;
                    return {
                        asin: p.asin,
                        title: p.title,
                        ventasMensuales: row['Tendencias de ventas mensuales: Comprados el mes pasado'] || 0,
                        bsr: row['Clasificación de Ventas: Actual'] || 'N/A',
                        precioActual: row['Caja de Compra: Actual'] || 0,
                        precioPromedio90: row['Caja de Compra: Promedio de 90 días'] || 0,
                        fbaElegibles: row['Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBA'] || 0,
                        fbmElegibles: row['Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBM'] || 0,
                        estVentasUnidades: row['Est. # Ventas Mensual'] || 0,
                        estVentasDolares: row['Est. $ Ventas Mensual'] || 0,
                        compraMaxFBM: row[`Compra Máx (${config.roiBajo}%) ($) FBM`] || 0,
                        descReqFBM: row[`% Desc. Req (${config.roiBajo}%) FBM`] || 0,
                        compraMaxFBA1: row[`Compra Máx (${config.roiAlto}%) ($) FBA`] || 0,
                        descReqFBA1: row[`% Desc. Req (${config.roiAlto}%) FBA`] || 0,
                        compraMaxFBA2: row[`Compra Máx (${config.roiMedio}%) ($) FBA`] || 0,
                        descReqFBA2: row[`% Desc. Req (${config.roiMedio}%) FBA`] || 0,
                        pesoLibras: (row['Paquete: Peso (g)'] || 0) * 0.00220462,
                        comisionReferencia: (parseFloat(row['% de comisión de referencia']) || 15) / 100
                    };
                });
            
                // --- PROMPT CUANTITATIVO (con análisis de %DescReq) ---
                const promptCuantitativo = `
                    Eres un analista financiero experto en Amazon.
                    Analiza los siguientes datos de Keepa y cálculos para la marca "${nombreMarca}".
                    
                    Para CADA producto, evalúa su viabilidad basándote en TODOS los factores disponibles.
                    No te centres solo en el % de descuento requerido; intégralo con el resto de indicadores.
                    
                    Considera:
                    - **Demanda**: ventas mensuales reales, BSR (cuanto más bajo, mejor).
                    - **Competencia**: número de vendedores FBA y FBM elegibles. Si hay muchos, el margen se comprime.
                    - **Márgenes**: compara el % de descuento requerido para FBA y para FBM. Si FBM requiere mucho menos descuento, sugiere FBM.
                    - **Logística**: si el peso supera 15 lb, el producto es oversize y FBA es caro; recomienda FBM.
                    - **Volatilidad**: si el precio actual difiere >10% del promedio 90 días, alerta de posible pico temporal.
                    - **Ventas estimadas**: si <10 unidades/mes o <$500/mes, es baja rotación.
                    
                    Clasifica el producto en una de estas zonas, pero no uses el % de descuento como único criterio; combínalo con los demás:
                    
                    - **Zona Esmeralda**: todos los indicadores son positivos (alta demanda, poca competencia, buen margen, peso ligero, precio estable). → ✅ "Producto excelente. Priorizar búsqueda de proveedores."
                    - **Zona Verde**: mayoría de indicadores positivos, pero algún factor neutral (ej. competencia moderada). → ✅ "Producto viable. Buscar proveedores estándar."
                    - **Zona Amarilla**: varios indicadores en zona media (ej. descuento aceptable pero peso alto o competencia intensa). → ⚠️ "Producto marginal. Negociar volumen o buscar promociones, o considerar FBM."
                    - **Zona Roja**: varios indicadores negativos (descuento >60%, peso >15 lb, competencia alta, baja demanda). → ❌ "Producto inviable. Descartar automáticamente."
                    
                    Si el producto cae en zona esmeralda o verde, indica si el precio actual es estable (no >10% por encima del promedio 90 días).
                    
                    Devuelve un objeto JSON con un campo "resumenes" que sea un ARRAY de strings,
                    donde cada string es el resumen de UN producto en el MISMO orden en que aparecen listados.
                    Cada resumen debe comenzar con ✅, ⚠️ o ❌, seguido de una justificación breve que mencione los factores clave.
                    
                    Productos:
                    ${productosInfo.map((p, idx) => `
                    ${idx+1}. ASIN: ${p.asin} | Título: ${p.title} | Ventas: ${p.ventasMensuales} | BSR: ${p.bsr}
                    | Precio actual: $${p.precioActual.toFixed(2)} | Prom 90d: $${p.precioPromedio90.toFixed(2)}
                    | Vendedores FBA: ${p.fbaElegibles} | FBM: ${p.fbmElegibles}
                    | Est. ventas: ${p.estVentasUnidades} uds ($${p.estVentasDolares.toFixed(2)})
                    | Compra FBM (${config.roiBajo}%): $${p.compraMaxFBM.toFixed(2)} (desc: ${(p.descReqFBM*100).toFixed(1)}%)
                    | Compra FBA (${config.roiAlto}%): $${p.compraMaxFBA1.toFixed(2)} (desc: ${(p.descReqFBA1*100).toFixed(1)}%)
                    | Compra FBA (${config.roiMedio}%): $${p.compraMaxFBA2.toFixed(2)} (desc: ${(p.descReqFBA2*100).toFixed(1)}%)
                    | Peso: ${p.pesoLibras.toFixed(2)} lb | Comisión: ${p.comisionReferencia*100}%
                    | **% Desc. Req (30%) FBA**: ${(p.descReqFBA1*100).toFixed(1)}%
                    `).join('\n')}
                    
                    Responde SOLO con el objeto JSON, sin texto adicional.
                    Ejemplo: {"resumenes": ["✅ Producto excelente. Buena demanda (500 uds/mes), baja competencia (2 FBA), peso ligero (1 lb) y descuento viable (35%). Priorizar proveedores.", "⚠️ Producto marginal. Competencia alta (12 FBA) y peso oversize (22 lb), aunque el descuento es aceptable (52%). Considerar FBM o negociar volumen.", ...]}
                `;
                
                // --- PROMPT CUALITATIVO (investigación web) ---
                const promptCualitativo = `
                    Eres un detective de proveedores para Amazon Wholesale.
                    Investiga en profundidad la marca "${nombreMarca}" (NO uses datos de Keepa).
                    Busca información sobre: programa wholesale en EE.UU., contactos de ventas mayoristas, 
                    requisitos de apertura de cuenta (Tax ID, Resale Certificate, MOQ, etc.), 
                    fabricante/matriz, distribuidores autorizados, riesgo de IP (marcas, vendedores no autorizados), 
                    y estrategia de margen estimada para revendedores.
                    
                    Devuelve un objeto JSON con EXACTAMENTE esta estructura:
                    {
                        "resumenCualitativo": "resumen de UNA LÍNEA que comience con ✅ (buena), ⚠️ (media) o ❌ (mala/muy restrictiva), seguido de conclusión",
                        "admiteWholesale": "Sí" o "No" o "No encontrado",
                        "tipoProveedor": "Marca Directa" o "Distribuidor Autorizado" o "Mayorista Nacional" o "No encontrado",
                        "telefono": "número o null",
                        "contacto": "email o enlace o null",
                        "links": "enlace o null",
                        "requisitos": "requisitos o null",
                        "fabricante": "nombre o null",
                        "rutas_distribucion": "lista resumida o null",
                        "riesgo_ip": "resumen o null",
                        "estrategia_margen": "resumen o null"
                    }
                    Si no encuentras información, usa null.
                    Responde SOLO con el objeto JSON.
                `;
                
                // --- LLAMADAS A GEMINI EN PARALELO ---
                const [respCuant, respCual] = await Promise.all([
                    callGeminiWithRetry(promptCuantitativo),
                    callGeminiWithRetry(promptCualitativo)
                ]);

                solicitudesRealizadas += 2; // dos llamadas por marca
            
                // Procesar respuesta cuantitativa
                let datosCuantitativos = { resumenes: [] };
                try {
                    let textCuant = respCuant.text;
                    textCuant = textCuant.replace(/```json/gi, '').replace(/```/g, '').trim();
                    const parsed = JSON.parse(textCuant);
                    if (Array.isArray(parsed.resumenes) && parsed.resumenes.length === productos.length) {
                        datosCuantitativos = parsed;
                    } else if (Array.isArray(parsed)) {
                        datosCuantitativos = { resumenes: parsed };
                    } else {
                        const values = Object.values(parsed).filter(v => typeof v === 'string' && v.length > 0);
                        if (values.length === productos.length) {
                            datosCuantitativos = { resumenes: values };
                        } else {
                            throw new Error('No se pudo extraer un array de resúmenes');
                        }
                    }
                } catch (e) {
                    console.error('❌ Error parseando respuesta cuantitativa:', e.message);
                    console.log('📄 Respuesta cruda:', respCuant.text.substring(0, 500));
                    datosCuantitativos = { resumenes: productos.map(() => '⚠️ Error en análisis cuantitativo') };
                }
            
                // Procesar respuesta cualitativa
                let datosCualitativos = {
                    resumenCualitativo: '⚠️ Error en análisis cualitativo',
                    admiteWholesale: '',
                    tipoProveedor: '',
                    telefono: null,
                    contacto: null,
                    links: null,
                    requisitos: null,
                    fabricante: null,
                    rutas_distribucion: null,
                    riesgo_ip: null,
                    estrategia_margen: null
                };
                try {
                    let textCual = respCual.text;
                    textCual = textCual.replace(/```json/gi, '').replace(/```/g, '').trim();
                    const parsed = JSON.parse(textCual);
                    if (parsed.resumenCualitativo) datosCualitativos.resumenCualitativo = parsed.resumenCualitativo;
                    if (parsed.admiteWholesale) datosCualitativos.admiteWholesale = parsed.admiteWholesale;
                    if (parsed.tipoProveedor) datosCualitativos.tipoProveedor = parsed.tipoProveedor;
                    if (parsed.telefono) datosCualitativos.telefono = parsed.telefono;
                    if (parsed.contacto) datosCualitativos.contacto = parsed.contacto;
                    if (parsed.links) datosCualitativos.links = parsed.links;
                    if (parsed.requisitos) datosCualitativos.requisitos = parsed.requisitos;
                    if (parsed.fabricante) datosCualitativos.fabricante = parsed.fabricante;
                    if (parsed.rutas_distribucion) datosCualitativos.rutas_distribucion = parsed.rutas_distribucion;
                    if (parsed.riesgo_ip) datosCualitativos.riesgo_ip = parsed.riesgo_ip;
                    if (parsed.estrategia_margen) datosCualitativos.estrategia_margen = parsed.estrategia_margen;
                } catch (e) {
                    console.error('❌ Error parseando respuesta cualitativa:', e.message);
                    console.log('📄 Respuesta cruda:', respCual.text.substring(0, 500));
                }
            
                // Asegurar que tenemos suficientes resúmenes cuantitativos
                while (datosCuantitativos.resumenes.length < productos.length) {
                    datosCuantitativos.resumenes.push('⚠️ Sin análisis cuantitativo');
                }
            
                // ---- ASIGNAR RESÚMENES A CADA PRODUCTO (SOLO SI NO ES NOT_ELIGIBLE) ----
                productos.forEach((prod, idx) => {
                    const isNotEligible = prod.rowRef['Restriction Code'] === 'NOT_ELIGIBLE';
                    const descReq30 = parseFloat(prod.rowRef['% Desc. Req (30%) FBA']) || 0;
                    const descartePorDescuento = descReq30 > 0.70;
                
                    if (isNotEligible || descartePorDescuento) {
                        prod.rowRef['Resumen IA Cuantitativo'] = descartePorDescuento ? '❌ Descartado por alto % de descuento' : '';
                        prod.rowRef['Resumen IA Cualitativo'] = '';
                    } else {
                        const cuantText = datosCuantitativos.resumenes[idx] || '⚠️ Sin análisis cuantitativo';
                        prod.rowRef['Resumen IA Cuantitativo'] = cuantText;
                        prod.rowRef['Resumen IA Cualitativo'] = datosCualitativos.resumenCualitativo || '⚠️ Sin análisis cualitativo';
                        prod.rowRef['Admite Wholesale'] = datosCualitativos.admiteWholesale || '';
                        prod.rowRef['Tipo de Proveedor'] = datosCualitativos.tipoProveedor || '';
                        prod.rowRef['Teléfono de Contacto'] = datosCualitativos.telefono || '';
                        prod.rowRef['Correo / Formulario'] = datosCualitativos.contacto || '';
                        prod.rowRef['Links Proveedores Potenciales'] = datosCualitativos.links || '';
                        prod.rowRef['Requisitos de Apertura'] = datosCualitativos.requisitos || '';
                        prod.rowRef['Fabricante/Matriz'] = datosCualitativos.fabricante || '';
                        prod.rowRef['Rutas de Distribución'] = datosCualitativos.rutas_distribucion || '';
                        prod.rowRef['Riesgo IP / Claims'] = datosCualitativos.riesgo_ip || '';
                        prod.rowRef['Estrategia de Margen'] = datosCualitativos.estrategia_margen || '';
                    }
                });
            
                const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
                console.log(`✅ Marca "${nombreMarca}" procesada. Solicitudes: ${solicitudesRealizadas}/${LIMITE_DIARIO} en ${elapsed}s`);
            
            } catch (error) {
                if (error.isDailyLimit) {
                    console.log(`⛔ Límite diario de solicitudes alcanzado. Deteniendo procesamiento.`);
                    limiteAlcanzado = true;
                    break;
                }
                console.error(`❌ Error procesando marca "${nombreMarca}":`, error.message);
                productos.forEach(prod => {
                    if (prod.rowRef['Restriction Code'] !== 'NOT_ELIGIBLE') {
                        prod.rowRef['Resumen IA Cuantitativo'] = '⚠️ Error en análisis cuantitativo';
                        prod.rowRef['Resumen IA Cualitativo'] = '⚠️ Error en análisis cualitativo';
                    }
                });
            }
        }

        const totalTime = ((Date.now() - startTime) / 1000).toFixed(1);
        console.log(`\n📊 Resumen final (IA):`);
        console.log(`   - Solicitudes realizadas: ${solicitudesRealizadas}/${LIMITE_DIARIO}`);
        console.log(`   - Tiempo total: ${totalTime}s`);
        console.log(`   - Marcas procesadas: ${solicitudesRealizadas}`);
        console.log(`   - Marcas pendientes: ${marcas.length - solicitudesRealizadas}`);
    } else {
        console.log(`⏭️ IA saltada por solicitud del usuario (checkbox).`);
    }

    return {
        filasProcesadas,
        config,
        solicitudesRealizadas: 0, // si se saltó, 0
        marcasProcesadas: 0,
        marcasPendientes: 0,
        limiteAlcanzado: false
    };
}

// --------------------------------------------------------------
// 10. ENDPOINT PRINCIPAL /api/audit-excel (MODIFICADO)
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
        let saltarIA = false;

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

        if (req.body.saltarIA === 'true') {
            saltarIA = true;
            console.log(`⏭️ Se saltará el análisis de IA por solicitud del usuario.`);
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
            unidadesMap,
            saltarIA
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





// ============================================================
// ENDPOINT TEMPORAL: /api/catalog-raw
// Devuelve la respuesta CRUDA de la API de Catalog para un ASIN
// ============================================================
app.get('/api/catalog-raw', async (req, res) => {
    const asin = req.query.asin;
    if (!asin) {
        return res.status(400).json({ error: 'Falta el parámetro asin' });
    }

    try {
        const clientId = process.env.AMZ_CLIENT_ID;
        const clientSecret = process.env.AMZ_CLIENT_SECRET;
        const refreshToken = process.env.AMZ_REFRESH_TOKEN;

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
            throw new Error('Error al obtener token de acceso');
        }

        const tokenData = await tokenResponse.json();
        const accessToken = tokenData.access_token;

        const marketplaceId = 'ATVPDKIKX0DER';
        const url = `https://sellingpartnerapi-na.amazon.com/catalog/2022-04-01/items/${asin}?marketplaceIds=${marketplaceId}&includedData=summaries,dimensions`;

        const response = await fetch(url, {
            headers: {
                'x-amz-access-token': accessToken,
                'Content-Type': 'application/json'
            }
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();
        res.json(data);

    } catch (error) {
        console.error(`❌ Error en /api/catalog-raw para ${asin}:`, error.message);
        res.status(500).json({ error: error.message });
    }
});




// ============================================================
// ENDPOINT TEMPORAL: /api/keepa-raw
// Devuelve la respuesta CRUDA de Keepa para un ASIN
// ============================================================
app.get('/api/keepa-raw', async (req, res) => {
    const asin = req.query.asin;
    if (!asin) {
        return res.status(400).json({ error: 'Falta el parámetro asin' });
    }

    const apiKey = process.env.KEEPA_API_KEY;
    if (!apiKey) {
        return res.status(500).json({ error: 'KEEPA_API_KEY no configurada' });
    }

    try {
        const url = `https://api.keepa.com/product?key=${apiKey}&domain=1&asin=${asin}&stats=90&buybox=1`;
        const response = await fetch(url);
        const data = await response.json();
        res.json(data);
    } catch (error) {
        console.error(`❌ Error en /api/keepa-raw para ${asin}:`, error.message);
        res.status(500).json({ error: error.message });
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
    console.log(`🤖 Resúmenes IA: Cuantitativo y Cualitativo (saltable vía checkbox)`);
});
