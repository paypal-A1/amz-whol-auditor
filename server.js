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
// 1. FUNCIÓN CLAVE: Busca el valor de una columna por su nombre
//    sin importar la posición, probando múltiples sinónimos.
// --------------------------------------------------------------
function getColumnValue(row, posiblesNombres) {
    if (!row) return null;
    for (const nombre of posiblesNombres) {
        // Si la fila tiene esa clave y no está vacía
        if (row[nombre] !== undefined && row[nombre] !== null && row[nombre] !== '') {
            return row[nombre];
        }
    }
    return null; // No se encontró ningún nombre
}

// --------------------------------------------------------------
// 2. MOTOR PRINCIPAL DE PROCESAMIENTO
// --------------------------------------------------------------
async function procesarInventarioWholesale(fileBuffer, config) {
    // Leer el Excel subido
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet);

    // Extraer configuración del panel web
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
    // 3. RECORRER CADA FILA DEL EXCEL (Lectura dinámica de columnas)
    // --------------------------------------------------------------
    for (const row of rows) {
        // --- A. DATOS BÁSICOS (Título, ASIN, Marca) ---
        const titulo = getColumnValue(row, ['Title', 'Título']) || 'Sin Título';
        const asin = getColumnValue(row, ['ASIN']) || 'Desconocido';
        const marca = getColumnValue(row, ['Brand', 'Marca']) || 'Genérico';

        // --- B. VENTAS MENSUALES (La columna correcta en español) ---
        const ventasMensuales = parseFloat(
            getColumnValue(row, [
                'Tendencias de ventas mensuales: Comprados el mes pasado', // Español (tu Excel)
                'Bought past month', // Inglés
                'Monthly Sales', // Alternativo
                'Sales Drops (30 days)' // Fallback (aunque no es el volumen exacto)
            ]) || 0
        );

        // Filtro de descarte por volumen mínimo
        if (ventasMensuales < minSalesMonthly) continue;

        // --- C. PRECIO DE LA BUY BOX (Según selector: Actual o Promedio 90 días) ---
        let precioBuyBox = 0;
        if (priceBasis === '90day') {
            precioBuyBox = parseFloat(
                getColumnValue(row, [
                    'Caja de Compra: Promedio de 90 días', // Español
                    'Buy Box: 90 days avg', // Inglés
                    'Amazon 90 days avg'
                ]) || 0
            );
        } else {
            precioBuyBox = parseFloat(
                getColumnValue(row, [
                    'Caja de Compra: Actual', // Español
                    'Buy Box: Current', // Inglés
                    'Precio Actual'
                ]) || 0
            );
        }

        // Si no hay precio, saltamos esta fila (no se puede calcular)
        if (!precioBuyBox || precioBuyBox === 0) continue;

        // --- D. PESO (para calcular el flete a Amazon) ---
        const pesoGramos = parseFloat(
            getColumnValue(row, [
                'Paquete: Peso (g)', // Español
                'Weight (g)' // Inglés
            ]) || 0
        );
        const pesoLibras = pesoGramos * 0.00220462;
        const costoEnvioAmazon = pesoLibras * inboundShippingPound;

        // --- E. COMISIONES Y TARIFAS FIJAS DE AMAZON ---
        const referralFee = precioBuyBox * 0.15; // 15% estándar
        const fbaFee = parseFloat(
            getColumnValue(row, [
                'Tarifa FBA Pick&Pack', // Español (tu Excel)
                'FBA Pick & Pack Fee' // Inglés
            ]) || 0
        );

        // --- F. CÁLCULO DEL PUNTO DE EQUILIBRIO (Break-Even / 0% ROI) ---
        const breakEven = precioBuyBox 
                        - fbaFee 
                        - referralFee 
                        - costoEnvioAmazon 
                        - prepFee 
                        - supplierShippingUnit;

        // --- G. CÁLCULO DE PRECIOS MÁXIMOS Y DESCUENTOS PARA CADA ROI ---
        const calcularCompraMax = (roi) => breakEven / (1 + (roi / 100));
        const calcularDescuento = (precioMax) => ((precioBuyBox - precioMax) / precioBuyBox) * 100;

        const maxAlto = calcularCompraMax(roiAlto);
        const maxMedio = calcularCompraMax(roiMedio);
        const maxBajo = calcularCompraMax(roiBajo);

        // --- H. COMPETENCIA REAL (FBA + FBM Elegibles para la Buy Box) ---
        const fbaElegibles = parseInt(
            getColumnValue(row, [
                'Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBA'
            ]) || 0
        );
        const fbmElegibles = parseInt(
            getColumnValue(row, [
                'Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBM'
            ]) || 0
        );

        // Sumamos todos los competidores + 1 (nosotros)
        const competidoresTotales = fbaElegibles + fbmElegibles + 1;
        const estVentasUnidades = ventasMensuales / competidoresTotales;
        const estVentasDolares = estVentasUnidades * precioBuyBox;

        // --------------------------------------------------------------
        // 4. CONSTRUIR LA FILA FINAL (Mantiene el orden original + añade al final)
        // --------------------------------------------------------------
        // IMPORTANTE: Al usar ...row (spread) primero, TODAS las columnas
        // originales de Keepa se mantienen. Luego añadimos las nuestras.
        // Esto asegura que nuestras columnas SIEMPRE queden al final,
        // sin importar cuántas columnas nuevas tenga Keepa en el futuro.
        const filaConMetricas = {
            ...row, // <-- Aquí se copian todas las columnas originales (en su orden)

            // NUEVAS COLUMNAS (se añaden al final automáticamente)
            'Break-Even ($)': breakEven.toFixed(2),
            'Compra Máx (ROI Alto) ($)': maxAlto.toFixed(2),
            '% Desc. Req (ROI Alto)': `${calcularDescuento(maxAlto).toFixed(1)}%`,
            'Compra Máx (ROI Medio) ($)': maxMedio.toFixed(2),
            '% Desc. Req (ROI Medio)': `${calcularDescuento(maxMedio).toFixed(1)}%`,
            'Compra Máx (ROI Bajo) ($)': maxBajo.toFixed(2),
            '% Desc. Req (ROI Bajo)': `${calcularDescuento(maxBajo).toFixed(1)}%`,
            'Est. # Ventas Mensual': Math.round(estVentasUnidades),
            'Est. $ Ventas Mensual': estVentasDolares.toFixed(2),

            // Espacios para la IA (se rellenarán después)
            'Admite Wholesale': '',
            'Tipo de Proveedor': '',
            'Teléfono de Contacto': '',
            'Correo / Formulario': '',
            'Links Proveedores Potenciales': '',
            'Requisitos de Apertura': '',
            'Dictamen de Salud': '',
            'Riesgo de IP / Alerta': '',
            'Conclusión General': ''
        };

        filasProcesadas.push(filaConMetricas);

        // Agrupar por marca para la auditoría de IA
        if (!productosPorMarca[marca]) productosPorMarca[marca] = [];
        productosPorMarca[marca].push({ 
            asin, 
            title: titulo, 
            rowRef: filaConMetricas 
        });
    }

    // --------------------------------------------------------------
    // 5. AUDITORÍA CON IA (AGRUPADA POR MARCA)
    // --------------------------------------------------------------
    for (const [nombreMarca, productos] of Object.entries(productosPorMarca)) {
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

            const response = await ai.models.generateContent({
                model: 'gemini-2.5-flash',
                contents: prompt,
                config: { responseMimeType: 'application/json' }
            });

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
        }
    }

    // --------------------------------------------------------------
    // 6. GENERAR EL NUEVO EXCEL
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
