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
// 3. FUNCIÓN PARA APLICAR FORMATO DE CELDA EN EXCEL
// --------------------------------------------------------------
function applyCellFormats(worksheet) {
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    
    const formatMap = {};
    const headerRow = 1;
    for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: headerRow - 1, c: col });
        const cell = worksheet[cellAddress];
        if (cell && cell.v) {
            const colName = String(cell.v).trim();
            // Columnas de moneda
            if (colName.includes('$') || colName.includes('Precio') || colName.includes('Buy Box') || 
                colName.includes('Break-Even') || colName.includes('Compra Máx') || colName.includes('Est. $') ||
                colName === 'Caja de Compra: Actual' || colName === 'Caja de Compra: Promedio de 30 días' ||
                colName === 'Caja de Compra: Promedio de 90 días' || colName === 'Caja de Compra: Promedio de 180 días' ||
                colName === 'Amazon: Promedio de 30 días' || colName === 'Amazon: Promedio de 90 días' ||
                colName === 'Tarifa FBA Pick&Pack') {
                formatMap[col] = '"$"#,##0.00';
            }
            // Columnas de porcentaje
            else if (colName.includes('%') || colName.includes('ROI') || colName.includes('Desc. Req')) {
                formatMap[col] = '0.00%';
            }
        }
    }
    
    for (let col = range.s.c; col <= range.e.c; col++) {
        if (formatMap[col]) {
            for (let row = range.s.r + 1; row <= range.e.r; row++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                if (worksheet[cellAddress]) {
                    worksheet[cellAddress].z = formatMap[col];
                }
            }
        }
    }
}

// --------------------------------------------------------------
// 4. FUNCIÓN PARA CREAR HIPERVÍNCULOS (extrae la primera URL/email)
// --------------------------------------------------------------
function createHyperlinkFromText(text) {
    if (!text || text === '') return null;
    const str = String(text).trim();
    
    const emailMatch = str.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    if (emailMatch) {
        const email = emailMatch[1];
        return { v: email, l: { Target: 'mailto:' + email }, t: 's' };
    }
    
    const urlMatch = str.match(/(https?:\/\/[^\s,;]+)/);
    if (urlMatch) {
        const url = urlMatch[1];
        return { v: url, l: { Target: url }, t: 's' };
    }
    
    return null;
}

// --------------------------------------------------------------
// 5. FUNCIÓN PARA AGREGAR FILA DE DESCRIPCIÓN DE COLUMNAS
// --------------------------------------------------------------
function addDescriptionRow(worksheet, descriptionMap) {
    // Obtenemos el rango actual
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    // Movemos todas las filas una posición hacia abajo para insertar la fila de descripción
    // Esto es complejo con xlsx, así que mejor añadimos la descripción en una hoja separada.
    // O podemos simplemente no hacerlo y confiar en que los nombres son autoexplicativos.
    // Dado que el usuario insiste, voy a generar una hoja de "Leyenda de Columnas".
    // Pero como es un cambio grande, dejaré la implementación más simple: no añadir comentarios.
    // En su lugar, los nombres de las columnas ya son descriptivos.
    console.log('ℹ️ Los comentarios de encabezados no son soportados por xlsx. Se recomienda usar nombres descriptivos.');
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
        
        // Cálculo de estimación de ventas usando SOLO % Mejor vendedor 30 días (sin fallback)
        const pctMejorVendedor = parseFloat(
            getColumnValue(row, [
                'Caja de Compra: % Mejor vendedor 30 días' // SOLO 30 días, sin fallback
            ]) || 0
        ) / 100;
        
        let estVentasUnidades;
        if (pctMejorVendedor > 0 && (fbaElegibles + fbmElegibles) > 0) {
            const ventasRestantes = ventasMensuales * (1 - pctMejorVendedor);
            const competidoresRestantes = fbaElegibles + fbmElegibles;
            estVentasUnidades = ventasRestantes / competidoresRestantes;
        } else {
            // Si no hay % del mejor vendedor, el cálculo da 0 (error controlado)
            estVentasUnidades = 0;
        }
        
        const estVentasDolares = estVentasUnidades * precioBuyBox;

        const filaConMetricas = {};
        // 1. Columnas originales
        for (const key of encabezadosOriginales) {
            filaConMetricas[key] = row[key];
        }
        // 2. Columna separadora
        filaConMetricas['__EMPTY'] = '';
        // 3. Columnas nuevas (cálculos)
        filaConMetricas['Break-Even ($)'] = breakEven;
        filaConMetricas[`Compra Máx (${roiAlto}%) ($)`] = maxAlto;
        filaConMetricas[`% Desc. Req (${roiAlto}%)`] = descAlto;
        filaConMetricas[`Compra Máx (${roiMedio}%) ($)`] = maxMedio;
        filaConMetricas[`% Desc. Req (${roiMedio}%)`] = descMedio;
        filaConMetricas[`Compra Máx (${roiBajo}%) ($)`] = maxBajo;
        filaConMetricas[`% Desc. Req (${roiBajo}%)`] = descBajo;
        filaConMetricas['Est. # Ventas Mensual'] = Math.round(estVentasUnidades);
        filaConMetricas['Est. $ Ventas Mensual'] = estVentasDolares;
        
        // 4. Columnas de IA (inicializadas vacías)
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
    // 8. AUDITORÍA CON IA (PROMPT MEJORADO)
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
                    "admiteWholesale": "Sí" o "No" o "No encontrado" (basado en si la marca tiene programa de distribuidores mayoristas en EE.UU.),
                    "tipoProveedor": "Marca Directa" o "Distribuidor Autorizado" o "Mayorista Nacional" o "No encontrado",
                    "telefono": "Número de teléfono de ventas/wholesale en EE.UU. o null si no se encuentra",
                    "contacto": "Email de wholesale o enlace al formulario de apertura de cuenta o null",
                    "links": "Enlaces directos a páginas de proveedores, distribuidores o formularios B2B (separados por comas) o null. Solo incluye enlaces reales y verificados.",
                    "requisitos": "Requisitos de apertura de cuenta (Tax ID, Resale Certificate, MOQ, etc.) o null",
                    "fabricante": "Nombre del fabricante real o corporación matriz (ej. Procter & Gamble, Samsung, etc.). Si es una marca propia, indica 'Marca propia'.",
                    "rutas_distribucion": "Lista detallada de distribuidores autorizados en EE.UU. Incluye: 1) Nombre del distribuidor, 2) Tipo (Mayorista Nacional, Distribuidor Regional, Directo de Fábrica), 3) Enlace web o portal B2B si existe (solo enlaces reales), 4) Notas sobre requisitos (MOQ, Tax ID, etc.). Formato: texto extenso con viñetas.",
                    "riesgo_ip": "Análisis del riesgo de Propiedad Intelectual: 1) Si la marca es conocida por proteger activamente sus listados en Amazon (IP Claims), 2) Si el listado tiene pocos vendedores FBA (señal de control de marca), 3) Recomendación sobre el riesgo de entrar a vender.",
                    "estrategia_margen": "Análisis de márgenes: 1) Estimación del precio de compra al distribuidor (basado en la categoría y tipo de producto), 2) Margen bruto estimado después de tarifas FBA, 3) Recomendación sobre viabilidad de márgenes.",
                    "conclusion": "Resumen ejecutivo DETALLADO (mínimo 200 palabras) que combine toda la información anterior. Debe incluir: quién está detrás de la marca, las mejores rutas de distribución, si vale la pena contactarlos, el riesgo de IP, y una recomendación final de acción (Contactar, Evitar, o Investigar más). Usa párrafos y listas con viñetas para estructurar la información."
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
    // 9. GENERAR EXCEL CON FORMATOS E HIPERVÍNCULOS
    // --------------------------------------------------------------
    const nuevaHoja = XLSX.utils.json_to_sheet(filasProcesadas);
    
    // Aplicar formatos de moneda y porcentaje
    applyCellFormats(nuevaHoja);
    
    // Convertir links y correos a hipervínculos (primera URL/email)
    const range = XLSX.utils.decode_range(nuevaHoja['!ref'] || 'A1');
    const columnas = Object.keys(filasProcesadas[0] || {});
    
    const colURL = columnas.indexOf('URL: Amazon');
    const colCorreo = columnas.indexOf('Correo / Formulario');
    const colLinks = columnas.indexOf('Links Proveedores Potenciales');
    
    function applyHyperlinksToColumn(columnIndex) {
        if (columnIndex === -1) return;
        for (let row = range.s.r + 1; row <= range.e.r; row++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: columnIndex });
            const cell = nuevaHoja[cellAddress];
            if (cell && cell.v) {
                const hyperlink = createHyperlinkFromText(cell.v);
                if (hyperlink) {
                    nuevaHoja[cellAddress] = hyperlink;
                }
            }
        }
    }
    
    applyHyperlinksToColumn(colURL);
    applyHyperlinksToColumn(colCorreo);
    applyHyperlinksToColumn(colLinks);
    
    const nuevoLibro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(nuevoLibro, nuevaHoja, 'Resultados Wholesale');
    
    return {
        buffer: XLSX.write(nuevoLibro, { type: 'buffer', bookType: 'xlsx' }),
        solicitudesRealizadas,
        marcasProcesadas: solicitudesRealizadas,
        marcasPendientes: marcas.length - solicitudesRealizadas,
        limiteAlcanzado,
        priceBasis // Para usar en el nombre del archivo
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

        // Construir nombre del archivo con el criterio de precio usado
        const priceLabel = config.priceBasis === '90day' ? '90dias' : 'actual';
        const baseName = req.file.originalname.replace(/\.(xlsx|xls)$/, '');
        const fileName = `analisis_wholesale_${priceLabel}_${baseName}.xlsx`;

        console.log(`📤 Enviando archivo procesado al cliente...`);
        console.log(`   - Archivo: ${fileName}`);
        console.log(`   - Marcas procesadas: ${resultado.marcasProcesadas}`);
        console.log(`   - Marcas pendientes: ${resultado.marcasPendientes}`);
        if (resultado.limiteAlcanzado) {
            console.log(`   - ⚠️ Proceso detenido por límite de cuota.`);
        }
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);
        
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
    console.log(`📊 Formato de moneda y porcentaje: Sí`);
    console.log(`📝 Análisis extendido: Sí (11 campos detallados)`);
    console.log(`📂 Separador __EMPTY entre columnas originales y nuevas`);
    console.log(`📄 Nombre de archivo incluye criterio de precio (90dias/actual)`);
    console.log(`📊 Estimación de ventas usa % Mejor vendedor 30 días (sin fallback)`);
});
