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
// 2. FUNCIÓN PARA REINTENTAR LLAMADAS A LA API DE GEMINI
//    CON BACKOFF ESCALONADO (maneja 429 y 503)
// --------------------------------------------------------------
async function callGeminiWithRetry(prompt, maxRetries = 4) {
    let lastError;
    // Backoff: 6s, 12s, 24s, 48s (para respetar 10 RPM)
    const backoffDelays = [6000, 12000, 24000, 48000];

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            console.log(`📡 Llamando a Gemini (intento ${attempt}/${maxRetries})...`);
            const response = await ai.models.generateContent({
                model: 'gemini-3.6-flash', // <-- Modelo elegido por calidad
                contents: prompt,
                config: {
                    responseMimeType: 'application/json'
                }
            });
            console.log(`✅ Respuesta recibida (intento ${attempt})`);
            return response;
        } catch (error) {
            lastError = error;
            
            // Si es error de cuota (429) o de disponibilidad (503), esperamos y reintentamos
            if (error.status === 429 || error.status === 503) {
                const waitTime = backoffDelays[attempt - 1] || 60000;
                const errorType = error.status === 429 ? 'cuota excedida' : 'saturado';
                console.log(`⏳ Gemini ${errorType} (${error.status}), esperando ${waitTime/1000}s antes de reintentar (${attempt}/${maxRetries})`);
                await new Promise(resolve => setTimeout(resolve, waitTime));
                continue;
            }
            
            // Si es error 404 (modelo no encontrado) o similar, no reintentamos
            console.error(`❌ Error no recuperable (${error.status}):`, error.message);
            throw error;
        }
    }
    console.error(`❌ Todos los ${maxRetries} reintentos fallaron. Último error:`, lastError);
    throw lastError;
}

// --------------------------------------------------------------
// 3. MOTOR PRINCIPAL DE PROCESAMIENTO
// --------------------------------------------------------------
async function procesarInventarioWholesale(fileBuffer, config) {
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Leer datos preservando todas las columnas
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
    if (rows.length === 0) {
        throw new Error('El archivo Excel no contiene datos.');
    }

    // Obtener encabezados en el orden exacto
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
    // 4. PROCESAR CADA FILA (cálculos y agrupación)
    // --------------------------------------------------------------
    for (const row of rows) {
        // -- Lectura dinámica de campos --
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
        const maxAlto = calcularCompraMax(roiAlto);
        const maxMedio = calcularCompraMax(roiMedio);
        const maxBajo = calcularCompraMax(roiBajo);

        const fbaElegibles = parseInt(
            getColumnValue(row, ['Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBA']) || 0
        );
        const fbmElegibles = parseInt(
            getColumnValue(row, ['Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBM']) || 0
        );
        const competidoresTotales = fbaElegibles + fbmElegibles + 1;
        const estVentasUnidades = ventasMensuales / competidoresTotales;
        const estVentasDolares = estVentasUnidades * precioBuyBox;

        // --- Construir objeto con orden original + nuevas columnas al final ---
        const filaConMetricas = {};
        // 1. Columnas originales en el orden exacto
        for (const key of encabezadosOriginales) {
            filaConMetricas[key] = row[key];
        }
        // 2. Nuevas columnas al final
        filaConMetricas['Break-Even ($)'] = breakEven.toFixed(2);
        filaConMetricas['Compra Máx (ROI Alto) ($)'] = maxAlto.toFixed(2);
        filaConMetricas['ROI Alto Alcanzado (%)'] = `${roiAlto}%`;
        filaConMetricas['Compra Máx (ROI Medio) ($)'] = maxMedio.toFixed(2);
        filaConMetricas['ROI Medio Alcanzado (%)'] = `${roiMedio}%`;
        filaConMetricas['Compra Máx (ROI Bajo) ($)'] = maxBajo.toFixed(2);
        filaConMetricas['ROI Bajo Alcanzado (%)'] = `${roiBajo}%`;
        filaConMetricas['Est. # Ventas Mensual'] = Math.round(estVentasUnidades);
        filaConMetricas['Est. $ Ventas Mensual'] = estVentasDolares.toFixed(2);

        // Espacios para IA (se rellenan después)
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

        // Agrupar por marca para la IA
        if (!productosPorMarca[marca]) productosPorMarca[marca] = [];
        productosPorMarca[marca].push({ asin, title: titulo, rowRef: filaConMetricas });
    }

    console.log(`✅ ${filasProcesadas.length} productos aprobados para análisis.`);
    console.log(`📦 Marcas identificadas: ${Object.keys(productosPorMarca).length}`);

    // --------------------------------------------------------------
    // 5. AUDITORÍA CON IA (AGRUPADA POR MARCA CON DELAY Y BACKOFF)
    // --------------------------------------------------------------
    const marcas = Object.keys(productosPorMarca);
    // Ordenar marcas por cantidad de productos (las que tienen más productos primero)
    marcas.sort((a, b) => productosPorMarca[b].length - productosPorMarca[a].length);

    let solicitudesRealizadas = 0;
    const startTime = Date.now();
    const LIMITE_DIARIO = 1500; // Límite diario de gemini-3.6-flash (por seguridad)

    // Delay de 6 segundos entre marcas para respetar 10 RPM (gemini-3.6-flash)
    const DELAY_ENTRE_MARCAS = 6000;

    for (let i = 0; i < marcas.length; i++) {
        const nombreMarca = marcas[i];
        const productos = productosPorMarca[nombreMarca];

        // Si alcanzamos el límite diario, detener el procesamiento
        if (solicitudesRealizadas >= LIMITE_DIARIO) {
            console.log(`⛔ Límite diario de ${LIMITE_DIARIO} solicitudes alcanzado. No se procesarán más marcas.`);
            break;
        }

        console.log(`\n🔍 Procesando marca ${i+1}/${marcas.length}: "${nombreMarca}" (${productos.length} productos)`);
        console.log(`📊 Solicitudes realizadas: ${solicitudesRealizadas}/${LIMITE_DIARIO}`);

        // Delay de 6 segundos entre marcas (excepto la primera)
        if (i > 0) {
            console.log(`⏳ Esperando ${DELAY_ENTRE_MARCAS/1000}s antes de siguiente marca...`);
            await new Promise(resolve => setTimeout(resolve, DELAY_ENTRE_MARCAS));
        }

        try {
            const prompt = `
                Analiza la marca comercial de Amazon: "${nombreMarca}".
                Para los siguientes productos asociados: ${JSON.stringify(productos.map(p => ({ asin: p.asin, title: p.title })))}
                
                Entrega una respuesta estrictamente en formato JSON plano (un objeto cuyas llaves sean los ASINs proporcionados). El objeto por cada ASIN debe contener obligatoriamente estos campos:
                {
                    "admiteWholesale": "Sí" o "No" o "Desconocido",
                    "tipoProveedor": "Marca Directa" o "Distribuidor Autorizado" o "Mayorista Nacional",
                    "telefono": "Número de teléfono de ventas en EE.UU.",
                    "contacto": "Email o enlace al formulario de apertura",
                    "links": "Enlaces web de proveedores",
                    "requisitos": "Notas de apertura (Tax ID, MOQ, etc.)",
                    "dictamenSalud": "SALUDABLE" o "MONOPOLIO" o "AMAZON" o "PRICE TANKING",
                    "riesgoIP": "Ninguno" o "Alerta de reclamos conocidos",
                    "conclusion": "Resumen ejecutivo combinando viabilidad comercial"
                }
            `;

            // Llamada a Gemini con reintentos automáticos
            const response = await callGeminiWithRetry(prompt);
            solicitudesRealizadas++;

            // --- LIMPIAR JSON (evita SyntaxError por markdown) ---
            let textoLimpio = response.text;
            textoLimpio = textoLimpio.replace(/```json/gi, '');
            textoLimpio = textoLimpio.replace(/```/g, '');
            textoLimpio = textoLimpio.trim();

            const datosIA = JSON.parse(textoLimpio);

            // Inyectar datos en cada fila
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
                } else {
                    console.log(`⚠️ No se encontraron datos para ASIN ${prod.asin} en la respuesta de IA`);
                }
            }
            
            const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
            console.log(`✅ Marca "${nombreMarca}" procesada. Solicitudes: ${solicitudesRealizadas}/${LIMITE_DIARIO} en ${elapsed}s`);

        } catch (error) {
            // Si es error de límite diario, detener todo
            if (error.message && error.message.includes('DAILY_LIMIT')) {
                console.log(`⛔ Límite diario alcanzado. Deteniendo procesamiento.`);
                break;
            }
            console.error(`❌ Error procesando marca "${nombreMarca}":`, error.message);
            // Continuar con la siguiente marca
        }
    }

    const totalTime = ((Date.now() - startTime) / 1000).toFixed(1);
    console.log(`\n📊 Resumen final:`);
    console.log(`   - Solicitudes realizadas: ${solicitudesRealizadas}/${LIMITE_DIARIO}`);
    console.log(`   - Tiempo total: ${totalTime}s`);
    console.log(`   - Marcas procesadas: ${solicitudesRealizadas}`);
    console.log(`   - Marcas pendientes: ${marcas.length - solicitudesRealizadas}`);

    // --------------------------------------------------------------
    // 6. GENERAR EXCEL (SIEMPRE, incluso con error)
    // --------------------------------------------------------------
    const nuevaHoja = XLSX.utils.json_to_sheet(filasProcesadas);
    const nuevoLibro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(nuevoLibro, nuevaHoja, 'Resultados Wholesale');
    
    return XLSX.write(nuevoLibro, { type: 'buffer', bookType: 'xlsx' });
}

// --------------------------------------------------------------
// 7. ENDPOINT /api/audit-excel
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

        const outputBuffer = await procesarInventarioWholesale(req.file.buffer, config);

        console.log(`📤 Enviando archivo procesado al cliente...`);
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=analisis_wholesale_${req.file.originalname}`);
        res.send(outputBuffer);

        console.log('✅ Proceso completado exitosamente.');

    } catch (error) {
        console.error("❌ Error crítico procesando Excel:", error);
        // Aunque haya error, intentamos devolver un mensaje claro
        res.status(500).json({ error: 'Ocurrió un error interno al procesar el archivo Excel: ' + error.message });
    }
});

// --------------------------------------------------------------
// 8. INICIAR SERVIDOR
// --------------------------------------------------------------
app.listen(PORT, () => {
    console.log(`🚀 Servidor corriendo en el puerto ${PORT}`);
    console.log(`📊 Modelo: gemini-3.6-flash (calidad prioritaria)`);
    console.log(`📅 Límite diario: 1500 solicitudes/día`);
    console.log(`⏱️  Delay entre marcas: 6 segundos (para 10 RPM)`);
});
