const express = require('express');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');
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
// 3. FUNCIÓN PARA APLICAR FORMATO A LAS CELDAS DEL EXCEL
// --------------------------------------------------------------
function applyCellFormat(worksheet, colName, format) {
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: XLSX.utils.decode_col(colName) });
        if (worksheet[cellAddress]) {
            worksheet[cellAddress].z = format;
        }
    }
}

// --------------------------------------------------------------
// 4. FUNCIÓN PARA CREAR HIPERVÍNCULO EN EXCEL
// --------------------------------------------------------------
function createHyperlink(text, url) {
    return {
        v: text || url,
        l: { Target: url },
        t: 's'
    };
}

// --------------------------------------------------------------
// 5. MOTOR PRINCIPAL DE PROCESAMIENTO
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
    // 6. FILTRADO Y CÁLCULOS MATEMÁTICOS (con nombres dinámicos)
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
        const calcularDescuento = (precioMax) => ((precioBuyBox - precioMax) / precioBuyBox) * 100;

        const maxAlto = calcularCompraMax(roiAlto);
        const maxMedio = calcularCompraMax(roiMedio);
        const maxBajo = calcularCompraMax(roiBajo);

        const descAlto = calcularDescuento(maxAlto);
        const descMedio = calcularDescuento(maxMedio);
        const descBajo = calcularDescuento(maxBajo);

        const fbaElegibles = parseInt(
            getColumnValue(row, ['Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBA']) || 0
        );
        const fbmElegibles = parseInt(
            getColumnValue(row, ['Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBM']) || 0
        );
        const competidoresTotales = fbaElegibles + fbmElegibles + 1;
        const estVentasUnidades = ventasMensuales / competidoresTotales;
        const estVentasDolares = estVentasUnidades * precioBuyBox;

        const filaConMetricas = {};
        
        // Columnas originales
        for (const key of encabezadosOriginales) {
            filaConMetricas[key] = row[key];
        }
        
        // Columnas matemáticas con nombres dinámicos
        filaConMetricas['Break-Even ($)'] = breakEven;
        filaConMetricas[`Compra Máx (${roiAlto}%) ($)`] = maxAlto;
        filaConMetricas[`% Desc. Req (${roiAlto}%)`] = descAlto;
        filaConMetricas[`Compra Máx (${roiMedio}%) ($)`] = maxMedio;
        filaConMetricas[`% Desc. Req (${roiMedio}%)`] = descMedio;
        filaConMetricas[`Compra Máx (${roiBajo}%) ($)`] = maxBajo;
        filaConMetricas[`% Desc. Req (${roiBajo}%)`] = descBajo;
        filaConMetricas['Est. # Ventas Mensual'] = Math.round(estVentasUnidades);
        filaConMetricas['Est. $ Ventas Mensual'] = estVentasDolares;

        // Columnas de IA (inicializadas vacías)
        filaConMetricas['Admite Wholesale'] = '';
        filaConMetricas['Tipo de Proveedor'] = '';
        filaConMetricas['Teléfono de Contacto'] = '';
        filaConMetricas['Correo / Formulario'] = '';
        filaConMetricas['Links Proveedores Potenciales'] = '';
        filaConMetricas['Requisitos de Apertura'] = '';
        filaConMetricas['Dictamen de Salud'] = '';
        filaConMetricas['Riesgo de IP / Alerta'] = '';
        filaConMetricas['Conclusión General'] = '';

        filasProcesadas.push(filaConMetricas);

        if (!productosPorMarca[marca]) productosPorMarca[marca] = [];
        productosPorMarca[marca].push({ asin, title: titulo, rowRef: filaConMetricas });
    }

    console.log(`✅ ${filasProcesadas.length} productos aprobados para análisis.`);
    console.log(`📦 Marcas identificadas: ${Object.keys(productosPorMarca).length}`);

    // --------------------------------------------------------------
    // 7. AUDITORÍA CON IA (PROMPT MEJORADO Y DETALLADO)
    // --------------------------------------------------------------
    const marcas = Object.keys(productosPorMarca);
    marcas.sort((a, b) => productosPorMarca[b].length - productosPorMarca[a].length);

    let solicitudesRealizadas = 0;
    const startTime = Date.now();
    const LIMITE_DIARIO = 1500;
    let limiteAlcanzado = false;
    let motivoDetencion = '';

    for (let i = 0; i < marcas.length; i++) {
        const nombreMarca = marcas[i];
        const productos = productosPorMarca[nombreMarca];

        if (limiteAlcanzado || solicitudesRealizadas >= LIMITE_DIARIO) {
            console.log(`⛔ Límite diario de ${LIMITE_DIARIO} solicitudes alcanzado.`);
            motivoDetencion = `Límite diario de ${LIMITE_DIARIO} solicitudes alcanzado (procesadas ${solicitudesRealizadas}).`;
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
                Actúa como un detective de proveedores y experto en Amazon Wholesale. Investiga a fondo la marca "${nombreMarca}" y sus productos asociados.

                Productos a analizar: ${JSON.stringify(productos.map(p => ({ asin: p.asin, title: p.title })))}

                **INVESTIGA Y RESPONDE ESTRICTAMENTE EN FORMATO JSON. NO INVENTES DATOS. Si no encuentras información, usa null o "No encontrado".**

                Debes devolver un objeto donde cada ASIN sea una clave, y el valor sea otro objeto con estos campos:

                {
                    "admiteWholesale": "Sí" o "No" o "No encontrado" (basado en si la marca tiene programa de distribuidores en EE.UU.),
                    "tipoProveedor": "Marca Directa" o "Distribuidor Autorizado" o "Mayorista Nacional" o "No encontrado",
                    "telefono": "Número de teléfono de ventas/wholesale en EE.UU. o null si no se encuentra",
                    "contacto": "Email de wholesale o enlace al formulario de apertura de cuenta o null",
                    "links": "Enlaces directos a páginas de proveedores, distribuidores o formularios B2B (máximo 3 enlaces separados por comas) o null",
                    "requisitos": "Requisitos de apertura de cuenta (Tax ID, Resale Certificate, MOQ, etc.) o null",
                    "dictamenSalud": "SALUDABLE" o "MONOPOLIO" o "AMAZON" o "PRICE TANKING" o "No evaluado",
                    "riesgoIP": "Ninguno" o "Alerta de reclamos conocidos" o "No encontrado",
                    "conclusion": "ANÁLISIS EXTENSO Y DETALLADO en texto libre (mínimo 150 palabras) que incluya: quien es el fabricante real, rutas de distribución autorizadas, análisis de viabilidad de márgenes, estrategia de compra recomendada, y cualquier otra observación relevante sobre la marca. Debe ser una investigación profunda como si fueras un detective de proveedores."
                }

                **Instrucciones adicionales:**
                - Busca el fabricante real (si es una marca de una corporación más grande, como P&G, Nestlé, etc.)
                - Identifica distribuidores autorizados (McLane, Core-Mark, KeHE, etc.)
                - Evalúa el riesgo de IP Claims (propiedad intelectual)
                - Analiza si el margen es viable considerando los costos típicos de distribución
                - **SI NO ENCUENTRAS INFORMACIÓN, NO INVENTES. USA null o "No encontrado".**
            `;

            const response = await callGeminiWithRetry(prompt);
            solicitudesRealizadas++;

            let textoLimpio = response.text;
            textoLimpio = textoLimpio.replace(/```json/gi, '');
            textoLimpio = textoLimpio.replace(/```/g, '');
            textoLimpio = textoLimpio.trim();

            const datosIA = JSON.parse(textoLimpio);

            for (const prod of productos) {
                const dataAsin = datosIA[prod.asin];
                if (dataAsin) {
                    prod.rowRef['Admite Wholesale'] = dataAsin.admiteWholesale || '';
                    prod.rowRef['Tipo de Proveedor'] = dataAsin.tipoProveedor || '';
                    prod.rowRef['Teléfono de Contacto'] = dataAsin.telefono || '';
                    prod.rowRef['Correo / Formulario'] = dataAsin.contacto || '';
                    prod.rowRef['Links Proveedores Potenciales'] = dataAsin.links || '';
                    prod.rowRef['Requisitos de Apertura'] = dataAsin.requisitos || '';
                    prod.rowRef['Dictamen de Salud'] = dataAsin.dictamenSalud || '';
                    prod.rowRef['Riesgo de IP / Alerta'] = dataAsin.riesgoIP || '';
                    prod.rowRef['Conclusión General'] = dataAsin.conclusion || '';
                }
            }
            
            const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
            console.log(`✅ Marca "${nombreMarca}" procesada. Solicitudes: ${solicitudesRealizadas}/${LIMITE_DIARIO} en ${elapsed}s`);

        } catch (error) {
            if (error.isDailyLimit) {
                console.log(`⛔ Límite diario de solicitudes alcanzado. Deteniendo procesamiento.`);
                limiteAlcanzado = true;
                motivoDetencion = `Límite diario excedido: ${error.details || 'No se pueden hacer más solicitudes hoy.'}`;
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
    if (limiteAlcanzado) {
        console.log(`   - ⚠️ Proceso detenido por límite de cuota.`);
    }

    // --------------------------------------------------------------
    // 8. GENERAR EXCEL CON FORMATOS Y HIPERVÍNCULOS
    // --------------------------------------------------------------
    // Convertir a hoja de cálculo
    const nuevaHoja = XLSX.utils.json_to_sheet(filasProcesadas);
    const nuevoLibro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(nuevoLibro, nuevaHoja, 'Resultados Wholesale');

    // Obtener todas las columnas
    const columnas = Object.keys(filasProcesadas[0] || {});

    // Aplicar formatos de celda
    columnas.forEach(col => {
        // Formato de moneda para columnas que contienen precios
        if (col.includes('($)') || col.includes('Break-Even') || col.includes('Buy Box') || col.includes('Est. $')) {
            applyCellFormat(nuevaHoja, col, '$#,##0.00');
        }
        // Formato de porcentaje para columnas que contienen %
        if (col.includes('%') || col.includes('Desc. Req')) {
            applyCellFormat(nuevaHoja, col, '0.00%');
        }
    });

    // Convertir hipervínculos en columnas específicas
    const columnasConLinks = ['URL: Amazon', 'Correo / Formulario', 'Links Proveedores Potenciales'];
    const range = XLSX.utils.decode_range(nuevaHoja['!ref']);
    
    columnasConLinks.forEach(colName => {
        const colIndex = columnas.indexOf(colName);
        if (colIndex === -1) return;
        
        for (let row = range.s.r + 1; row <= range.e.r; row++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: colIndex });
            const cell = nuevaHoja[cellAddress];
            if (cell && cell.v && typeof cell.v === 'string') {
                // Si es una URL o contiene http, crear hipervínculo
                const valor = cell.v.toString().trim();
                if (valor.startsWith('http://') || valor.startsWith('https://') || valor.includes('@')) {
                    // Si es email, crear mailto
                    let url = valor;
                    if (valor.includes('@') && !valor.startsWith('mailto:')) {
                        url = `mailto:${valor}`;
                    }
                    nuevaHoja[cellAddress] = createHyperlink(valor, url);
                }
            }
        }
    });

    // Ajustar ancho de columnas (opcional)
    const colWidths = [];
    columnas.forEach((col, idx) => {
        let maxLen = col.length;
        for (let row = range.s.r + 1; row <= Math.min(range.e.r, range.s.r + 50); row++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: idx });
            const cell = nuevaHoja[cellAddress];
            if (cell && cell.v) {
                const val = cell.v.toString();
                if (val.length > maxLen) maxLen = val.length;
            }
        }
        colWidths.push({ wch: Math.min(Math.max(maxLen + 2, 12), 50) });
    });
    nuevaHoja['!cols'] = colWidths;

    return {
        buffer: XLSX.write(nuevoLibro, { type: 'buffer', bookType: 'xlsx' }),
        solicitudesRealizadas,
        marcasProcesadas: solicitudesRealizadas,
        marcasPendientes: marcas.length - solicitudesRealizadas,
        limiteAlcanzado,
        motivoDetencion
    };
}

// --------------------------------------------------------------
// 9. ENDPOINT /api/audit-excel
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

        console.log(`📤 Enviando archivo procesado al cliente...`);
        console.log(`   - Marcas procesadas: ${resultado.marcasProcesadas}`);
        console.log(`   - Marcas pendientes: ${resultado.marcasPendientes}`);
        if (resultado.limiteAlcanzado) {
            console.log(`   - ⚠️ ${resultado.motivoDetencion}`);
        }
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=analisis_wholesale_${req.file.originalname}`);
        
        res.setHeader('X-Processing-Info', JSON.stringify({
            marcasProcesadas: resultado.marcasProcesadas,
            marcasPendientes: resultado.marcasPendientes,
            limiteAlcanzado: resultado.limiteAlcanzado,
            motivo: resultado.motivoDetencion || 'Completado'
        }));
        
        res.send(resultado.buffer);

        console.log('✅ Proceso completado exitosamente.');

    } catch (error) {
        console.error("❌ Error crítico procesando Excel:", error);
        res.status(500).json({ error: 'Ocurrió un error interno al procesar el archivo Excel: ' + error.message });
    }
});

// --------------------------------------------------------------
// 10. INICIAR SERVIDOR
// --------------------------------------------------------------
app.listen(PORT, () => {
    console.log(`🚀 Servidor corriendo en el puerto ${PORT}`);
    console.log(`📊 Modelo: gemini-3.5-flash-lite`);
    console.log(`📅 Límite diario: 1,500 solicitudes/día`);
    console.log(`⏱️  Delay entre marcas: 5 segundos (para 15 RPM)`);
    console.log(`🔄 Reintentos: 4 intentos para errores 503 y 429 (no diarios)`);
});
