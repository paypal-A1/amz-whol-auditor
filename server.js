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
    
    const descripcionesGenericas = {
        'Título': 'Nombre completo del producto en Amazon',
        'URL: Amazon': 'Enlace directo al producto en Amazon',
        'ASIN': 'Amazon Standard Identification Number (identificador único del producto)',
        'Marca': 'Marca del producto',
        'Clasificación de Ventas: Actual': 'Posición actual en el ranking de ventas de la categoría (Best Sellers Rank)',
        'Clasificación de Ventas: Promedio de 90 días': 'Promedio del ranking de ventas de los últimos 90 días',
        'Clasificación de Ventas: Descensos en los últimos 30 días': 'Número de caídas en el ranking (drops) en los últimos 30 días',
        'Tendencias de ventas mensuales: Comprados el mes pasado': 'Volumen de ventas estimado en el último mes (datos de Amazon)',
        'Opiniones: Cantidad de valoraciones': 'Número total de reseñas del producto',
        'Caja de Compra: Actual': 'Precio actual de la Buy Box',
        'Caja de Compra: Promedio de 30 días': 'Precio promedio de la Buy Box en los últimos 30 días',
        'Caja de Compra: Promedio de 90 días': 'Precio promedio de la Buy Box en los últimos 90 días',
        'Caja de Compra: Promedio de 180 días': 'Precio promedio de la Buy Box en los últimos 180 días',
        'Caja de Compra: Vendedor Caja de Compra': 'Vendedor que tiene actualmente la Buy Box',
        'Caja de Compra: % Amazon 30 días': 'Porcentaje de tiempo que Amazon tuvo la Buy Box (30 días)',
        'Caja de Compra: % Amazon 90 días': 'Porcentaje de tiempo que Amazon tuvo la Buy Box (90 días)',
        'Caja de Compra: % Mejor vendedor 30 días': 'Porcentaje de tiempo que el mejor vendedor tuvo la Buy Box (30 días)',
        'Caja de Compra: % Mejor vendedor 90 días': 'Porcentaje de tiempo que el mejor vendedor tuvo la Buy Box (90 días)',
        'Caja de Compra: Es FBA': 'Indica si el vendedor actual usa FBA',
        'Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBA': 'Número de vendedores FBA elegibles para la Buy Box',
        'Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBM': 'Número de vendedores FBM elegibles para la Buy Box',
        'Amazon: Actual': 'Precio actual del producto vendido por Amazon',
        'Amazon: Promedio de 30 días': 'Precio promedio de Amazon (30 días)',
        'Amazon: Promedio de 90 días': 'Precio promedio de Amazon (90 días)',
        'Tarifa FBA Pick&Pack': 'Tarifa que cobra Amazon por almacenar, empacar y enviar el producto',
        '% de comisión de referencia': 'Comisión que cobra Amazon por la venta (Referral Fee)',
        'Recuento total de Ofertas': 'Número total de ofertas disponibles para el producto',
        'Recuento ofertas nuevas FBA: Actual': 'Número de vendedores FBA actualmente activos',
        'Recuento ofertas nuevas FBM: Actual': 'Número de vendedores FBM actualmente activos',
        'Códigos de producto: UPC': 'Universal Product Code (código de barras)',
        'Paquete: Dimensión (cm³)': 'Volumen del paquete en centímetros cúbicos',
        'Paquete: Peso (g)': 'Peso del paquete en gramos',
        'Es HazMat': 'Indica si el producto es considerado material peligroso',
        'Es sensible al calor': 'Indica si el producto es sensible a altas temperaturas',
        'Producto para adultos': 'Indica si el producto está clasificado para adultos'
    };
    
    const descripcionesNuevas = {
        'Break-Even ($)': 'Punto de equilibrio (0% ROI). Precio máximo de compra para no perder dinero. Fórmula: Precio Buy Box - FBA Fee - Referral Fee - Envío Interno - Prep Fee - Envío Proveedor',
        'Est. # Ventas Mensual': 'Estimación de unidades mensuales. Fórmula: Ventas Totales × (1 - %MejorVendedor30d) / (FBA Elegibles + FBM Elegibles)',
        'Est. $ Ventas Mensual': 'Estimación de ingresos brutos mensuales (Est. # Ventas Mensual × Precio Buy Box)',
        'Resumen Keepa': 'Resumen concluyente basado en datos de Keepa y cálculos financieros. Evalúa demanda, competencia y márgenes. Comienza con ✅ (positivo), ⚠️ (neutro) o ❌ (negativo)',
        'Resumen IA': 'Resumen concluyente basado en la investigación de la IA sobre la marca. Evalúa wholesale, contactos, riesgos. Comienza con ✅ (positivo), ⚠️ (neutro) o ❌ (negativo)',
        'Admite Wholesale': 'Indica si la marca tiene programa de distribuidores mayoristas en EE.UU.',
        'Tipo de Proveedor': 'Clasificación: Marca Directa, Distribuidor Autorizado, Mayorista Nacional',
        'Teléfono de Contacto': 'Número de teléfono del departamento de ventas/wholesale',
        'Correo / Formulario': 'Email de contacto o enlace al formulario de apertura de cuenta',
        'Links Proveedores Potenciales': 'Enlaces a páginas de proveedores, distribuidores o formularios B2B',
        'Requisitos de Apertura': 'Requisitos necesarios para abrir cuenta mayorista (Tax ID, Resale Certificate, MOQ, etc.)',
        'Fabricante/Matriz': 'Fabricante real o corporación matriz detrás de la marca',
        'Rutas de Distribución': 'Lista detallada de distribuidores autorizados con enlaces y requisitos',
        'Riesgo IP / Claims': 'Análisis del riesgo de Propiedad Intelectual y reclamos de marca',
        'Estrategia de Margen': 'Análisis de márgenes estimados y recomendación de viabilidad financiera',
        'Conclusión General': 'Análisis integral combinando datos de Keepa, cálculos financieros e investigación de IA. Recomendación final de acción.'
    };
    
    if (descripcionesGenericas[colName]) return descripcionesGenericas[colName];
    if (descripcionesNuevas[colName]) return descripcionesNuevas[colName];
    
    const compraMaxMatch = colName.match(/^Compra Máx \((\d+)%\) \(\$\)$/);
    if (compraMaxMatch) {
        const roi = compraMaxMatch[1];
        return `Precio máximo de compra para lograr ${roi}% de ROI. Fórmula: Break-Even / (1 + (${roi} / 100))`;
    }
    
    const descReqMatch = colName.match(/^% Desc\. Req \((\d+)%\)$/);
    if (descReqMatch) {
        const roi = descReqMatch[1];
        return `Descuento necesario para lograr ${roi}% de ROI. Fórmula: ((Precio Buy Box - Compra Máx) / Precio Buy Box) (mostrado como porcentaje)`;
    }
    
    return 'Columna generada por el sistema';
}

// --------------------------------------------------------------
// 5. FUNCIÓN PARA EVALUAR VIABILIDAD (basada en resúmenes)
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
// 6. FUNCIÓN PARA CREAR EL EXCEL CON EXCELJS (con colores y formatos)
// --------------------------------------------------------------
async function createExcelWithStyles(filasProcesadas, config, nombreOriginal) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'AMZ Wholesale Auditor Pro';
    workbook.created = new Date();
    
    const worksheet = workbook.addWorksheet('Resultados Wholesale', {
        properties: { tabColor: { argb: 'FFD700' } }
    });
    
    const columnasOriginales = Object.keys(filasProcesadas[0] || {});
    
    // Definir el orden de columnas según bloques
    const bloque1 = ['Título', 'URL: Amazon', 'ASIN', 'Break-Even ($)', 'Compra Máx (30%) ($)', '% Desc. Req (30%)', 'Compra Máx (20%) ($)', '% Desc. Req (20%)', 'Compra Máx (15%) ($)', '% Desc. Req (15%)', 'Est. # Ventas Mensual', 'Est. $ Ventas Mensual'];
    const bloque2 = ['Resumen Keepa', 'Resumen IA'];
    const bloque3 = ['Admite Wholesale', 'Tipo de Proveedor', 'Teléfono de Contacto', 'Correo / Formulario', 'Links Proveedores Potenciales', 'Requisitos de Apertura', 'Fabricante/Matriz', 'Rutas de Distribución', 'Riesgo IP / Claims', 'Estrategia de Margen', 'Conclusión General'];
    // Bloque 4: resto de columnas originales (las que no están en bloque1)
    const bloque1Set = new Set(bloque1);
    const bloque4 = columnasOriginales.filter(col => !bloque1Set.has(col) && col !== '--- SEPARADOR ---' && col !== 'Viabilidad');
    
    const ordenColumnas = [...bloque1, ...bloque2, ...bloque3, ...bloque4];
    const headers = ordenColumnas.filter(col => columnasOriginales.includes(col));
    
    // Configurar columnas
    worksheet.columns = headers.map(col => ({
        header: col,
        key: col,
        width: (bloque2.includes(col) || bloque3.includes(col)) ? 60 : 18
    }));
    
    // Agregar datos en el orden definido
    filasProcesadas.forEach(row => {
        const rowData = {};
        headers.forEach(col => {
            let value = row[col];
            if (value === undefined || value === null) value = '';
            rowData[col] = value;
        });
        worksheet.addRow(rowData);
    });
    
    // ---- ALTURA DE FILA (45) ----
    for (let rowNum = 1; rowNum <= worksheet.rowCount; rowNum++) {
        worksheet.getRow(rowNum).height = 45;
    }
    
    // ---- ESTILOS ENCABEZADO (fila 1) POR BLOQUES ----
    const headerRow = worksheet.getRow(1);
    headerRow.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    
    // Asignar colores por bloque
    const coloresBloques = [
        { inicio: 0, fin: bloque1.length - 1, color: 'FF1565C0' },  // Azul oscuro
        { inicio: bloque1.length, fin: bloque1.length + bloque2.length - 1, color: 'FF2E7D32' }, // Verde oscuro
        { inicio: bloque1.length + bloque2.length, fin: bloque1.length + bloque2.length + bloque3.length - 1, color: 'FFE65100' }, // Naranja
        { inicio: bloque1.length + bloque2.length + bloque3.length, fin: headers.length - 1, color: 'FF424242' } // Gris oscuro
    ];
    
    for (let i = 0; i < headers.length; i++) {
        const cell = headerRow.getCell(i + 1);
        let color = 'FF424242'; // fallback
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
    
    // ---- FORMATOS Y HIPERVÍNCULOS ----
    const colIndexMap = {};
    headers.forEach((col, idx) => colIndexMap[col] = idx + 1);
    
    for (let rowNum = 2; rowNum <= worksheet.rowCount; rowNum++) {
        const row = worksheet.getRow(rowNum);
        row.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        
        // Evaluar viabilidad para color de fila
        const resKeepa = row.getCell(colIndexMap['Resumen Keepa'] || 1).value || '';
        const resIA = row.getCell(colIndexMap['Resumen IA'] || 2).value || '';
        const statusKeepa = evaluarViabilidad(String(resKeepa));
        const statusIA = evaluarViabilidad(String(resIA));
        
        let bgColor = null;
        if (statusKeepa === 'positivo' && statusIA === 'positivo') {
            bgColor = 'FFC6EFCE'; // Verde claro
        } else if (statusKeepa === 'negativo' || statusIA === 'negativo') {
            bgColor = 'FFFFC7CE'; // Rojo claro
        } else if (statusKeepa === 'neutral' && statusIA === 'neutral') {
            bgColor = 'FFFFEB9C'; // Amarillo claro
        } else {
            bgColor = 'FFFFEB9C'; // Amarillo claro por defecto
        }
        
        for (let colIdx = 0; colIdx < headers.length; colIdx++) {
            const colName = headers[colIdx];
            const cell = row.getCell(colIdx + 1);
            const value = cell.value;
            
            if (bgColor) {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
            }
            
            // Formato moneda o porcentaje
            let format = null;
            if (colName.includes('$') || colName.includes('Break-Even') || colName.includes('Compra Máx') || colName.includes('Est. $') ||
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
            
            // Hipervínculos
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
    
    // ---- ANCHO DE COLUMNAS (ajuste fino) ----
    worksheet.columns.forEach((col, idx) => {
        const header = col.header;
        if (bloque2.includes(header) || bloque3.includes(header)) {
            // Ancho automático para bloques de IA (máximo 60)
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
            // Bloques 1 y 4: ancho fijo de 18
            col.width = 18;
        }
    });
    
    // ---- HOJA DE SIGNIFICADO ----
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

    // --------------------------------------------------------------
    // 8. FILTRADO Y CÁLCULOS MATEMÁTICOS
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
        
        // Columnas matemáticas (ya no incluimos separador ni viabilidad)
        filaConMetricas['Break-Even ($)'] = breakEven;
        filaConMetricas[`Compra Máx (${roiAlto}%) ($)`] = maxAlto;
        filaConMetricas[`% Desc. Req (${roiAlto}%)`] = descAlto;
        filaConMetricas[`Compra Máx (${roiMedio}%) ($)`] = maxMedio;
        filaConMetricas[`% Desc. Req (${roiMedio}%)`] = descMedio;
        filaConMetricas[`Compra Máx (${roiBajo}%) ($)`] = maxBajo;
        filaConMetricas[`% Desc. Req (${roiBajo}%)`] = descBajo;
        filaConMetricas['Est. # Ventas Mensual'] = Math.round(estVentasUnidades);
        filaConMetricas['Est. $ Ventas Mensual'] = estVentasDolares;
        
        // Nuevas columnas de resumen IA
        filaConMetricas['Resumen Keepa'] = '';
        filaConMetricas['Resumen IA'] = '';
        
        // Columnas IA detalladas
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
    // 9. AUDITORÍA CON IA (PROMPT MEJORADO CON NUEVOS CAMPOS)
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

    // --------------------------------------------------------------
    // 10. GENERAR EXCEL CON EXCELJS
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
// 11. ENDPOINT /api/audit-excel
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
// 12. INICIAR SERVIDOR
// --------------------------------------------------------------
app.listen(PORT, () => {
    console.log(`🚀 Servidor corriendo en el puerto ${PORT}`);
    console.log(`📊 Modelo: gemini-3.5-flash-lite`);
    console.log(`📅 Límite diario: 1,500 solicitudes/día`);
    console.log(`⏱️  Delay entre marcas: 5 segundos (para 15 RPM)`);
    console.log(`🔗 Links clickeables: Sí`);
    console.log(`📊 Formatos: Sí`);
    console.log(`📝 Análisis extendido: Sí (con resúmenes y conclusión integral)`);
    console.log(`🎨 Colores por bloque: Sí (azul, verde, naranja, gris)`);
    console.log(`🎨 Viabilidad por fila: Sí (verde/amarillo/rojo según resúmenes)`);
    console.log(`📘 Hoja de significados: Incluida`);
    console.log(`📄 Nombre archivo: Incluye criterio (90day/actual)`);
});
