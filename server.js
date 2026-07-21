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
// --------------------------------------------------------------
async function callGeminiWithRetry(prompt, maxRetries = 3) {
    let lastError;
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            const response = await ai.models.generateContent({
                model: 'gemini-3.5-flash',  // <-- CAMBIADO A 3.5 FLASH
                contents: prompt,
                config: { 
                    responseMimeType: 'application/json',
                    // No usamos temperature, top_p, top_k (ya no recomendados)
                    thinkingLevel: 'medium'   // 'low' para más velocidad, 'high' para más profundidad
                }
            });
            return response;
        } catch (error) {
            lastError = error;
            if (error.status === 429) {
                const waitTime = Math.pow(2, attempt) * 1000 + Math.random() * 1000;
                console.log(`⏳ Cuota excedida, reintentando en ${waitTime/1000}s (intento ${attempt}/${maxRetries})`);
                await new Promise(resolve => setTimeout(resolve, waitTime));
                continue;
            }
            throw error;
        }
    }
    throw lastError;
}

// --------------------------------------------------------------
// 3. MOTOR PRINCIPAL DE PROCESAMIENTO
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

    // --------------------------------------------------------------
    // 4. PROCESAR CADA FILA (cálculos y agrupación)
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

        if (ventasMensuales < minSalesMonthly) continue;

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
        if (!precioBuyBox || precioBuyBox === 0) continue;

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

        const roiAltoStr = `${roiAlto}%`;
        const roiMedioStr = `${roiMedio}%`;
        const roiBajoStr = `${roiBajo}%`;

        const fbaElegibles = parseInt(
            getColumnValue(row, ['Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBA']) || 0
        );
        const fbmElegibles = parseInt(
            getColumnValue(row, ['Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBM']) || 0
        );
        const competidoresTotales = fbaElegibles + fbmElegibles + 1;
        const estVentasUnidades = ventasMensuales / competidoresTotales;
        const estVentasDolares = estVentasUnidades * precioBuyBox;

        // Construir objeto con el orden original + nuevas columnas al final
        const filaConMetricas = {};
        for (const key of encabezadosOriginales) {
            filaConMetricas[key] = row[key];
        }
        filaConMetricas['Break-Even ($)'] = breakEven.toFixed(2);
        filaConMetricas['Compra Máx (ROI Alto) ($)'] = maxAlto.toFixed(2);
        filaConMetricas['ROI Alto Alcanzado (%)'] = roiAltoStr;
        filaConMetricas['Compra Máx (ROI Medio) ($)'] = maxMedio.toFixed(2);
        filaConMetricas['ROI Medio Alcanzado (%)'] = roiMedioStr;
        filaConMetricas['Compra Máx (ROI Bajo) ($)'] = maxBajo.toFixed(2);
        filaConMetricas['ROI Bajo Alcanzado (%)'] = roiBajoStr;
        filaConMetricas['Est. # Ventas Mensual'] = Math.round(estVentasUnidades);
        filaConMetricas['Est. $ Ventas Mensual'] = estVentasDolares.toFixed(2);

        // Espacios para IA
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

    // --------------------------------------------------------------
    // 5. AUDITORÍA CON IA (con delay de 4 segundos entre marcas)
    // --------------------------------------------------------------
    const marcas = Object.keys(productosPorMarca);
    for (let i = 0; i < marcas.length; i++) {
        const nombreMarca = marcas[i];
        const productos = productosPorMarca[nombreMarca];

        // Delay de 4 segundos entre marcas (para respetar 15 RPM)
        if (i > 0) {
            await new Promise(resolve => setTimeout(resolve, 4000));
        }

        try {
            // Prompt mejorado para Gemini 3.5 Flash
            const prompt = `
                Eres un experto en análisis de marcas para Amazon Wholesale.
                Analiza la marca comercial de Amazon: "${nombreMarca}".
                Los siguientes productos (ASINs) pertenecen a esta marca:
                ${JSON.stringify(productos.map(p => ({ asin: p.asin, title: p.title })), null, 2)}

                Para cada ASIN, investiga y proporciona la siguiente información en un objeto JSON. 
                La respuesta debe ser un objeto cuyas claves sean los ASINs. Cada valor debe tener esta estructura:
                {
                    "admiteWholesale": "Sí" o "No" o "Desconocido",
                    "tipoProveedor": "Marca Directa" o "Distribuidor Autorizado" o "Mayorista Nacional" o "Desconocido",
                    "telefono": "Número de teléfono de ventas en EE.UU. (si se encuentra, si no, 'No encontrado')",
                    "contacto": "Email de contacto o enlace al formulario de apertura de cuenta mayorista",
                    "links": "Enlaces web relevantes (sitio oficial, página de mayoristas, etc.)",
                    "requisitos": "Requisitos de apertura de cuenta (por ejemplo, Tax ID, MOQ, etc.)",
                    "dictamenSalud": "SALUDABLE" o "MONOPOLIO" o "AMAZON" o "PRICE TANKING",
                    "riesgoIP": "Ninguno" o "Alerta de reclamos conocidos",
                    "conclusion": "Resumen ejecutivo que combine viabilidad comercial, estabilidad del listado y oportunidades de wholesale."
                }
                Si no encuentras información para algún campo, usa "No encontrado" o "Desconocido".
                Responde ÚNICAMENTE con el JSON, sin texto adicional.
            `;

            const response = await callGeminiWithRetry(prompt);
            const datosIA = JSON.parse(response.text);

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
        } catch (error) {
            console.error(`Error de IA en marca ${nombreMarca}:`, error);
            // Si falla, dejamos los campos vacíos (ya están como '')
        }
    }

    // --------------------------------------------------------------
    // 6. GENERAR NUEVO EXCEL
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

        const outputBuffer = await procesarInventarioWholesale(req.file.buffer, config);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=analisis_wholesale_completo.xlsx');
        res.send(outputBuffer);

    } catch (error) {
        console.error("Error crítico procesando Excel:", error);
        res.status(500).json({ error: 'Ocurrió un error interno al procesar e indexar el archivo Excel.' });
    }
});

// --------------------------------------------------------------
// 8. INICIAR SERVIDOR
// --------------------------------------------------------------
app.listen(PORT, () => {
    console.log(`Servidor corriendo en el puerto ${PORT}`);
});
