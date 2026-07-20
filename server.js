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

async function procesarInventarioWholesale(fileBuffer, config) {
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet);

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

    for (const row of rows) {
        const ventasMensuales = parseFloat(row['Sales Drops (30 days)'] || row['Ventas Mensuales Estimadas'] || 0);
        if (ventasMensuales < minSalesMonthly) continue;

        const precioBuyBox = priceBasis === '90day'
            ? parseFloat(row['Amazon 90 days avg'] || row['Buy Box 90 days avg'] || 0)
            : parseFloat(row['Buy Box: Current'] || row['Precio Actual'] || 0);

        if (!precioBuyBox || precioBuyBox === 0) continue;

        const pesoGramos = parseFloat(row['Weight (g)'] || 0);
        const pesoLibras = pesoGramos * 0.00220462;
        const costoEnvioAmazon = pesoLibras * inboundShippingPound;

        const referralFee = precioBuyBox * 0.15;
        const fbaFee = parseFloat(row['FBA Pick & Pack Fee'] || 0);

        const breakEven = precioBuyBox - fbaFee - referralFee - costoEnvioAmazon - prepFee - supplierShippingUnit;

        const calcularCompraMax = (roiObjetivo) => breakEven / (1 + (roiObjetivo / 100));
        const calcularDescuento = (precioMaximo) => ((precioBuyBox - precioMaximo) / precioBuyBox) * 100;

        const maxAlto = calcularCompraMax(roiAlto);
        const maxMedio = calcularCompraMax(roiMedio);
        const maxBajo = calcularCompraMax(roiBajo);

        const fbaElegibles = parseInt(row['Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBA'] || 0);
        const fbmElegibles = parseInt(row['Recuento de ofertas elegibles para la Caja de Compra: Nuevo FBM'] || 0);
        
        const competidoresTotales = fbaElegibles + fbmElegibles + 1; 
        const estVentasUnidades = ventasMensuales / competidoresTotales;
        const estVentasDolares = estVentasUnidades * precioBuyBox;

        const marca = row['Brand'] || row['Marca'] || 'Genérico';
        const asin = row['ASIN'] || 'Desconocido';

        const filaConMetricas = {
            ...row,
            'Break-Even ($)': breakEven.toFixed(2),
            'Compra Máx (ROI Alto) ($)': maxAlto.toFixed(2),
            '% Desc. Req (ROI Alto)': `${calcularDescuento(maxAlto).toFixed(1)}%`,
            'Compra Máx (ROI Medio) ($)': maxMedio.toFixed(2),
            '% Desc. Req (ROI Medio)': `${calcularDescuento(maxMedio).toFixed(1)}%`,
            'Compra Máx (ROI Bajo) ($)': maxBajo.toFixed(2),
            '% Desc. Req (ROI Bajo)': `${calcularDescuento(maxBajo).toFixed(1)}%`,
            'Est. # Ventas Mensual': Math.round(estVentasUnidades),
            'Est. $ Ventas Mensual': estVentasDolares.toFixed(2),
            'Admite Wholesale': '', 'Tipo de Proveedor': '', 'Teléfono de Contacto': '',
            'Correo / Formulario': '', 'Links Proveedores Potenciales': '',
            'Requisitos de Apertura': '', 'Dictamen de Salud': '', 'Riesgo de IP / Alerta': '',
            'Conclusión General': ''
        };

        filasProcesadas.push(filaConMetricas);

        if (!productosPorMarca[marca]) productosPorMarca[marca] = [];
        productosPorMarca[marca].push({ asin, title: row['Title'] || '', rowRef: filaConMetricas });
    }

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
                    prod.rowRef['Admite Wholesale'] = dataAsin.admiteWholesale;
                    prod.rowRef['Tipo de Proveedor'] = dataAsin.tipoProveedor;
                    prod.rowRef['Teléfono de Contacto'] = dataAsin.telefono;
                    prod.rowRef['Correo / Formulario'] = dataAsin.contacto;
                    prod.rowRef['Links Proveedores Potenciales'] = dataAsin.links;
                    prod.rowRef['Requisitos de Apertura'] = dataAsin.requisitos;
                    prod.rowRef['Dictamen de Salud'] = dataAsin.dictamenSalud;
                    prod.rowRef['Riesgo de IP / Alerta'] = dataAsin.riesgoIP;
                    prod.rowRef['Conclusión General'] = dataAsin.conclusion;
                }
            }
        } catch (error) {
            console.error(`Error de IA en marca ${nombreMarca}:`, error);
        }
    }

    const nuevaHoja = XLSX.utils.json_to_sheet(filasProcesadas);
    const nuevoLibro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(nuevoLibro, nuevaHoja, 'Resultados Wholesale');
    
    return XLSX.write(nuevoLibro, { type: 'buffer', bookType: 'xlsx' });
}

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

app.listen(PORT, () => {
    console.log(`Servidor corriendo en el puerto ${PORT}`);
});
