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
// 3. FUNCIÓN PARA CREAR HIPERVÍNCULO (extrae primera URL/email)
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
// 4. FUNCIÓN PARA GENERAR DESCRIPCIÓN DE COLUMNA (con detección dinámica)
// --------------------------------------------------------------
function getColumnDescription(colName, config) {
    const { roiAlto, roiMedio, roiBajo } = config;
    
    // Descripciones para columnas originales (genéricas)
    const descripcionesGenericas = {
        'Título': 'Nombre completo del producto en Amazon',
        'Clasificación de Ventas: Actual': 'Posición actual en el ranking de ventas de la categoría (Best Sellers Rank)',
        'Clasificación de Ventas: Promedio de 90 días': 'Promedio del ranking de ventas de los últimos 90 días',
        'Clasificación de Ventas: Descensos en los últimos 30 días': 'Número de caídas en el ranking (drops) en los últimos 30 días (proxy de velocidad de ventas)',
        'Tendencias de ventas mensuales: Comprados el mes pasado': 'Volumen de ventas estimado en el último mes (datos de Amazon)',
        'Opiniones: Cantidad de valoraciones': 'Número total de reseñas del producto',
        'Caja de Compra: Actual': 'Precio actual de la Buy Box',
        'Caja de Compra: Promedio de 30 días': 'Precio promedio de la Buy Box en los últimos 30 días',
        'Caja de Compra: Promedio de 90 días': 'Precio promedio de la Buy Box en los últimos 90 días',
        'Caja de Compra: Promedio de 180 días': 'Precio promedio de la Buy Box en los últimos 180 días',
        'Caja de Compra: Vendedor Caja de Compra': 'Vendedor que tiene actualmente la Buy Box (nombre y ID)',
        'Caja de Compra: % Amazon 30 días': 'Porcentaje de tiempo que Amazon tuvo la Buy Box en los últimos 30 días',
        'Caja de Compra: % Amazon 90 días': 'Porcentaje de tiempo que Amazon tuvo la Buy Box en los últimos 90 días',
        'Caja de Compra: % Mejor vendedor 30 días': 'Porcentaje de tiempo que el mejor vendedor tuvo la Buy Box en los últimos 30 días',
        'Caja de Compra: % Mejor vendedor 90 días': 'Porcentaje de tiempo que el mejor vendedor tuvo la Buy Box en los últimos 90 días',
        'Caja de Compra: Es FBA': 'Indica si el vendedor actual usa FBA (yes/no)',
        'Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBA': 'Número de vendedores FBA elegibles para la Buy Box',
        'Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBM': 'Número de vendedores FBM elegibles para la Buy Box',
        'Amazon: Actual': 'Precio actual del producto vendido por Amazon (si aplica)',
        'Amazon: Promedio de 30 días': 'Precio promedio de Amazon en los últimos 30 días',
        'Amazon: Promedio de 90 días': 'Precio promedio de Amazon en los últimos 90 días',
        'Tarifa FBA Pick&Pack': 'Tarifa que cobra Amazon por almacenar, empacar y enviar el producto (FBA fee)',
        '% de comisión de referencia': 'Comisión que cobra Amazon por la venta (Referral Fee)',
        'Recuento total de Ofertas': 'Número total de ofertas disponibles para el producto',
        'Recuento ofertas nuevas FBA: Actual': 'Número de vendedores FBA actualmente activos',
        'Recuento ofertas nuevas FBM: Actual': 'Número de vendedores FBM actualmente activos',
        'URL: Amazon': 'Enlace directo al producto en Amazon',
        'ASIN': 'Amazon Standard Identification Number (identificador único del producto)',
        'Códigos de producto: UPC': 'Universal Product Code (código de barras)',
        'Marca': 'Marca del producto',
        'Paquete: Dimensión (cm³)': 'Volumen del paquete en centímetros cúbicos',
        'Paquete: Peso (g)': 'Peso del paquete en gramos',
        'Es HazMat': 'Indica si el producto es considerado material peligroso (yes/no)',
        'Es sensible al calor': 'Indica si el producto es sensible a altas temperaturas (yes/no)',
        'Producto para adultos': 'Indica si el producto está clasificado para adultos (yes/no)',
        '--- SEPARADOR ---': 'Separador visual entre columnas originales y columnas generadas por el sistema'
    };
    
    // Descripciones base para columnas nuevas
    const descripcionesNuevasBase = {
        'Break-Even ($)': 'Punto de equilibrio (0% ROI). Precio máximo de compra para no perder dinero. Fórmula: Precio Buy Box - FBA Fee - Referral Fee - Envío Interno - Prep Fee - Envío Proveedor',
        'Est. # Ventas Mensual': 'Estimación de unidades que podrías vender al mes considerando la competencia real. Fórmula: Ventas Totales * (1 - %MejorVendedor30d) / (FBA Elegibles + FBM Elegibles)',
        'Est. $ Ventas Mensual': 'Estimación de ingresos brutos mensuales (Est. # Ventas Mensual * Precio Buy Box)',
        'Viabilidad': 'Evaluación general de viabilidad del producto según la IA (Alta, Media, Baja, Inviable)',
        'Admite Wholesale': 'Indica si la marca tiene programa de distribuidores mayoristas en EE.UU.',
        'Tipo de Proveedor': 'Clasificación del tipo de proveedor: Marca Directa, Distribuidor Autorizado, Mayorista Nacional',
        'Teléfono de Contacto': 'Número de teléfono del departamento de ventas/wholesale de la marca',
        'Correo / Formulario': 'Email de contacto o enlace al formulario de apertura de cuenta',
        'Links Proveedores Potenciales': 'Enlaces a páginas de proveedores, distribuidores o formularios B2B',
        'Requisitos de Apertura': 'Requisitos necesarios para abrir cuenta mayorista (Tax ID, Resale Certificate, MOQ, etc.)',
        'Fabricante/Matriz': 'Fabricante real o corporación matriz detrás de la marca',
        'Rutas de Distribución': 'Lista detallada de distribuidores autorizados con enlaces y requisitos',
        'Riesgo IP / Claims': 'Análisis del riesgo de Propiedad Intelectual y reclamos de marca',
        'Estrategia de Margen': 'Análisis de márgenes estimados y recomendación de viabilidad financiera',
        'Conclusión General': 'Resumen ejecutivo detallado con recomendación final de acción'
    };
    
    // Verificar si es una columna original
    if (descripcionesGenericas[colName]) {
        return descripcionesGenericas[colName];
    }
    
    // Verificar si es una columna nueva exacta
    if (descripcionesNuevasBase[colName]) {
        return descripcionesNuevasBase[colName];
    }
    
    // Detectar patrones dinámicos: Compra Máx (X%) ($)
    const compraMaxMatch = colName.match(/^Compra Máx \((\d+)%\) \(\$\)$/);
    if (compraMaxMatch) {
        const roi = compraMaxMatch[1];
        return `Precio máximo de compra para lograr ${roi}% de ROI. Fórmula: Break-Even / (1 + (${roi} / 100))`;
    }
    
    // Detectar patrones dinámicos: % Desc. Req (X%)
    const descReqMatch = colName.match(/^% Desc\. Req \((\d+)%\)$/);
    if (descReqMatch) {
        const roi = descReqMatch[1];
        return `Porcentaje de descuento que debes obtener del proveedor para lograr ${roi}% de ROI. Fórmula: ((Precio Buy Box - Compra Máx) / Precio Buy Box) * 100 (mostrado como fracción en la celda)`;
    }
    
    // Si no coincide con nada, devolver genérico
    return 'Columna generada por el sistema';
}

// --------------------------------------------------------------
// 5. FUNCIÓN PARA CREAR EL EXCEL CON EXCELJS (con colores y formatos)
// --------------------------------------------------------------
async function createExcelWithStyles(filasProcesadas, config, nombreOriginal) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'AMZ Wholesale Auditor Pro';
    workbook.created = new Date();
    
    // ----- HOJA 1: RESULTADOS -----
    const worksheet = workbook.addWorksheet('Resultados Wholesale', {
        properties: { tabColor: { argb: 'FFD700' } }
    });
    
    const columnas = Object.keys(filasProcesadas[0] || {});
    
    // Definir columnas
    worksheet.columns = columnas.map(col => ({
        header: col,
        key: col,
        width: Math.min(Math.max(col.length + 5, 15), 40)
    }));
    
    // Agregar datos
    filasProcesadas.forEach(row => {
        const rowData = {};
        columnas.forEach(col => {
            let value = row[col];
            // Si es undefined o null, dejamos vacío
            if (value === undefined || value === null) {
                rowData[col] = '';
                return;
            }
            rowData[col] = value;
        });
        worksheet.addRow(rowData);
    });
    
    // ----- APLICAR ESTILOS A LA FILA DE ENCABEZADO (fila 1) -----
    const headerRow = worksheet.getRow(1);
    headerRow.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    headerRow.height = 25;
    
    // ----- FORMATO DE CELDAS (moneda y porcentaje) Y HIPERVÍNCULOS -----
    const colIndexMap = {};
    columnas.forEach((col, idx) => {
        colIndexMap[col] = idx + 1; // ExcelJS usa índice 1-based
    });
    
    // Aplicar formatos y procesar hipervínculos
    for (let rowNum = 2; rowNum <= worksheet.rowCount; rowNum++) {
        const row = worksheet.getRow(rowNum);
        
        // Determinar color de fondo según viabilidad
        let bgColor = null;
        const viabilidadCol = columnas.indexOf('Viabilidad');
        if (viabilidadCol !== -1) {
            const cell = row.getCell(viabilidadCol + 1);
            const viabilidad = cell.value ? String(cell.value).trim() : '';
            if (viabilidad === 'Alta') bgColor = 'FFC6EFCE'; // Verde claro
            else if (viabilidad === 'Media') bgColor = 'FFFFEB9C'; // Amarillo claro
            else if (viabilidad === 'Baja' || viabilidad === 'Inviable') bgColor = 'FFFFC7CE'; // Rojo claro
        }
        
        // Recorrer celdas de la fila
        for (let colIdx = 0; colIdx < columnas.length; colIdx++) {
            const colName = columnas[colIdx];
            const cell = row.getCell(colIdx + 1);
            const value = cell.value;
            
            // Aplicar color de fondo si se determinó
            if (bgColor) {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
            }
            
            // Aplicar formato de moneda o porcentaje
            let format = null;
            if (colName.includes('$') || colName.includes('Precio') || colName.includes('Buy Box') || 
                colName.includes('Break-Even') || colName.includes('Compra Máx') || colName.includes('Est. $') ||
                colName === 'Caja de Compra: Actual' || colName === 'Caja de Compra: Promedio de 30 días' ||
                colName === 'Caja de Compra: Promedio de 90 días' || colName === 'Caja de Compra: Promedio de 180 días' ||
                colName === 'Amazon: Promedio de 30 días' || colName === 'Amazon: Promedio de 90 días' ||
                colName === 'Tarifa FBA Pick&Pack') {
                format = '$#,##0.00';
            } else if (colName.includes('%') || colName.includes('ROI') || colName.includes('Desc. Req')) {
                format = '0.00%';
            }
            
            if (format && typeof value === 'number') {
                cell.numFmt = format;
            }
            
            // Procesar hipervínculos (solo en columnas específicas)
            if (colName === 'URL: Amazon' || colName === 'Correo / Formulario' || colName === 'Links Proveedores Potenciales') {
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
    
    // ----- AJUSTAR ANCHO DE COLUMNAS -----
    worksheet.columns.forEach(col => {
        let maxLength = col.header.length;
        col.eachCell({ includeEmpty: true }, (cell) => {
            const cellValue = cell.value;
            if (cellValue) {
                const str = String(cellValue);
                if (str.length > maxLength) maxLength = str.length;
            }
        });
        col.width = Math.min(Math.max(maxLength + 2, 12), 50);
    });
    
    // ----- HOJA 2: SIGNIFICADO DE COLUMNAS -----
    const meaningSheet = workbook.addWorksheet('📘 Significado de Columnas', {
        properties: { tabColor: { argb: 'FF2196F3' } }
    });
    
    // Encabezados de la hoja de significado
    meaningSheet.columns = [
        { header: 'Nombre de Columna', key: 'columna', width: 35 },
        { header: 'Descripción', key: 'descripcion', width: 70 }
    ];
    
    // Agregar filas con significado
    columnas.forEach(col => {
        const desc = getColumnDescription(col, config);
        meaningSheet.addRow({ columna: col, descripcion: desc });
    });
    
    // Estilo para encabezado de la hoja de significado
    const meaningHeaderRow = meaningSheet.getRow(1);
    meaningHeaderRow.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    meaningHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1976D2' } };
    meaningHeaderRow.alignment = { horizontal: 'center', vertical: 'middle' };
    meaningHeaderRow.height = 22;
    
    // Generar buffer
    const buffer = await workbook.xlsx.writeBuffer();
    return buffer;
}

// --------------------------------------------------------------
// 6. MOTOR PRINCIPAL DE PROCESAMIENTO
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

    // --------------------------------------------------------------
    // 7. FILTRADO Y CÁLCULOS MATEMÁTICOS
    // --------------------------------------------------------------
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
        
        // ---- CÁLCULO DE ESTIMACIÓN DE VENTAS (SOLO CON 30 DÍAS, SIN FALLBACK) ----
        const pctMejorVendedor30d = parseFloat(
            getColumnValue(row, ['Caja de Compra: % Mejor vendedor 30 días'])
        );
        
        let estVentasUnidades = 0;
        if (pctMejorVendedor30d && pctMejorVendedor30d > 0 && (fbaElegibles + fbmElegibles) > 0) {
            const pct = pctMejorVendedor30d / 100;
            const ventasRestantes = ventasMensuales * (1 - pct);
            const competidoresRestantes = fbaElegibles + fbmElegibles;
            estVentasUnidades = ventasRestantes / competidoresRestantes;
        } else {
            estVentasUnidades = 0;
        }
        
        const estVentasDolares = estVentasUnidades * precioBuyBox;

        // ---- CONSTRUIR FILA CON SEPARADOR ----
        const filaConMetricas = {};
        
        // 1. Columnas originales
        for (const key of encabezadosOriginales) {
            filaConMetricas[key] = row[key];
        }
        
        // 2. Separador visual
        filaConMetricas['--- SEPARADOR ---'] = '';
        
        // 3. Columnas matemáticas
        filaConMetricas['Break-Even ($)'] = breakEven;
        filaConMetricas[`Compra Máx (${roiAlto}%) ($)`] = maxAlto;
        filaConMetricas[`% Desc. Req (${roiAlto}%)`] = descAlto;
        filaConMetricas[`Compra Máx (${roiMedio}%) ($)`] = maxMedio;
        filaConMetricas[`% Desc. Req (${roiMedio}%)`] = descMedio;
        filaConMetricas[`Compra Máx (${roiBajo}%) ($)`] = maxBajo;
        filaConMetricas[`% Desc. Req (${roiBajo}%)`] = descBajo;
        filaConMetricas['Est. # Ventas Mensual'] = Math.round(estVentasUnidades);
        filaConMetricas['Est. $ Ventas Mensual'] = estVentasDolares;
        
        // 4. Columnas de IA (incluyendo Viabilidad)
        filaConMetricas['Viabilidad'] = '';
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

    // --------------------------------------------------------------
    // 8. AUDITORÍA CON IA (PROMPT MEJORADO CON VIABILIDAD)
    // --------------------------------------------------------------
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
            console.log(`⛔ Límite diario de ${LIMITE_DIARIO} solicitudes alcanzado. No se procesarán más marcas.`);
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
                NO INVENTES DATOS. Los enlaces deben ser reales y verificados. Si no estás segura de un enlace, omítelo.
                
                {
                    "viabilidad": "Alta" o "Media" o "Baja" o "Inviable" (evalúa si el producto es viable para vender en Amazon basado en competencia, márgenes estimados y riesgo de IP),
                    "admiteWholesale": "Sí" o "No" o "No encontrado",
                    "tipoProveedor": "Marca Directa" o "Distribuidor Autorizado" o "Mayorista Nacional" o "No encontrado",
                    "telefono": "Número de teléfono de ventas/wholesale en EE.UU. o null si no se encuentra",
                    "contacto": "Email de wholesale o enlace al formulario de apertura de cuenta o null",
                    "links": "Enlaces directos a páginas de proveedores, distribuidores o formularios B2B (separados por comas) o null. Solo incluye enlaces reales y verificados.",
                    "requisitos": "Requisitos de apertura de cuenta (Tax ID, Resale Certificate, MOQ, etc.) o null",
                    "fabricante": "Nombre del fabricante real o corporación matriz. Si es una marca propia, indica 'Marca propia'.",
                    "rutas_distribucion": "Lista detallada de distribuidores autorizados en EE.UU. Incluye: 1) Nombre, 2) Tipo, 3) Enlace web (solo reales), 4) Notas sobre requisitos. Formato: texto extenso con viñetas.",
                    "riesgo_ip": "Análisis del riesgo de Propiedad Intelectual: 1) Si la marca es conocida por proteger activamente sus listados en Amazon (IP Claims), 2) Si el listado tiene pocos vendedores FBA, 3) Recomendación.",
                    "estrategia_margen": "Análisis de márgenes: 1) Estimación del precio de compra al distribuidor, 2) Margen bruto estimado después de tarifas FBA, 3) Recomendación sobre viabilidad.",
                    "conclusion": "Resumen ejecutivo DETALLADO (mínimo 200 palabras) que combine toda la información anterior. Debe incluir: quién está detrás de la marca, las mejores rutas de distribución, si vale la pena contactarlos, el riesgo de IP, y una recomendación final de acción (Contactar, Evitar, o Investigar más). Usa párrafos y listas con viñetas."
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
                    prod.rowRef['Viabilidad'] = info.viabilidad || '';
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

    // --------------------------------------------------------------
    // 9. GENERAR EXCEL CON EXCELJS (colores, formatos, hipervínculos)
    // --------------------------------------------------------------
    const buffer = await createExcelWithStyles(filasProcesadas, config, '');
    
    return {
        buffer: buffer,
        solicitudesRealizadas,
        marcasProcesadas: solicitudesRealizadas,
        marcasPendientes: marcas.length - solicitudesRealizadas,
        limiteAlcanzado
    };
}

// --------------------------------------------------------------
// 10. ENDPOINT /api/audit-excel
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

        // ---- NOMBRE DE ARCHIVO CON CRITERIO USADO ----
        const priceLabel = config.priceBasis === '90day' ? '90day' : 'actual';
        const nombreArchivo = `analisis_wholesale_${priceLabel}_${req.file.originalname}`;
        console.log(`📤 Enviando archivo: ${nombreArchivo}`);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=${nombreArchivo}`);
        res.send(resultado.buffer);

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
    console.log(`⏱️  Delay entre marcas: 5 segundos (para 15 RPM)`);
    console.log(`🔗 Links clickeables: Sí (primera URL/email)`);
    console.log(`📊 Formato de moneda y porcentaje: Sí (ExcelJS)`);
    console.log(`📝 Análisis extendido: Sí (12 campos incluyendo Viabilidad)`);
    console.log(`🎨 Colores: Encabezado negro/blanco, filas según viabilidad (ExcelJS)`);
    console.log(`📘 Hoja de significados: Incluida con descripciones dinámicas`);
    console.log(`📄 Nombre archivo: Incluye criterio (90day/actual)`);
    console.log(`📊 Estimación ventas: Fija con % Mejor vendedor 30 días`);
});
