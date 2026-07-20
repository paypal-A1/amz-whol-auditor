const express = require('express');
const dotenv = require('dotenv');
const { GoogleGenAI } = require('@google/genai');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Configuración de almacenamiento en memoria para procesar el Excel temporalmente
const upload = multer({ storage: multer.memoryStorage() });

app.post('/api/audit-excel', upload.single('excelFile'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No se ha cargado ningún archivo Excel.' });
        }

        // Recibir variables comerciales desde el panel de configuración
        const { minRoi, priceDropTolerance, customRules } = req.body;

        // Leer el archivo desde el buffer de memoria
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convertir la hoja a formato JSON JSON (objeto por fila)
        const rows = XLSX.utils.sheet_to_json(worksheet);

        if (rows.length === 0) {
            return res.status(400).json({ error: 'El archivo Excel no contiene filas de datos.' });
        }

        // Procesar las filas evaluando los datos de Keepa y agregando los campos finales
        for (let row of rows) {
            try {
                const interaction = await ai.interactions.create({
                    model: "gemini-3.5-flash",
                    systemInstruction: 
                        `Eres un auditor experto en Amazon Wholesale. Tu tarea es analizar las columnas de la fila suministrada.\n` +
                        `Criterios de evaluación configurados:\n` +
                        `- ROI Mínimo Requerido: ${minRoi}%\n` +
                        `- Tolerancia a Caídas de Precio (Price Tanking): ${priceDropTolerance}%\n` +
                        `- Reglas Adicionales: ${customRules || 'Ninguna'}.\n\n` +
                        `Analiza las métricas de Keepa presentes en la fila. Debes retornar estrictamente un objeto JSON válido con tres campos cortos:\n` +
                        `{\n` +
                        `  "estabilidad": "Estable / Inestable / Alerta Tanking",\n` +
                        `  "competencia": "Baja / Moderada / Saturada (Vendedores FBA)",\n` +
                        `  "dictamen": "Aprobado / Rechazado / Revisar Manualmente"\n` +
                        `}`,
                    input: `Métricas de la fila actual: ${JSON.stringify(row)}`,
                    generationConfig: {
                        thinkingLevel: "medium",
                        responseMimeType: "application/json"
                    }
                });

                const result = JSON.parse(interaction.output_text);

                // Agregar dinámicamente nuevas columnas al final de la fila identificada
                row['AUDIT: Estabilidad Precio'] = result.estabilidad || 'N/A';
                row['AUDIT: Competencia FBA'] = result.competencia || 'N/A';
                row['AUDIT: Dictamen Final'] = result.dictamen || 'N/A';

            } catch (err) {
                // En caso de error en una fila específica, no romper el flujo del archivo completo
                row['AUDIT: Estabilidad Precio'] = 'Error de Análisis';
                row['AUDIT: Competencia FBA'] = 'Error de Análisis';
                row['AUDIT: Dictamen Final'] = 'Omitido';
            }
        }

        // Re-generar la hoja de cálculo con las nuevas columnas posicionadas al final
        const updatedWorksheet = XLSX.utils.json_to_sheet(rows);
        workbook.Sheets[sheetName] = updatedWorksheet;

        // Escribir el nuevo archivo de Excel en un buffer de salida
        const outputBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

        // Enviar cabeceras de descarga de archivos binarios al navegador
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=auditoria_wholesale_completada.xlsx');
        res.send(outputBuffer);

    } catch (error) {
        console.error("Error procesando Excel:", error);
        res.status(500).json({ error: 'Ocurrió un error interno al procesar e indexar el archivo Excel.' });
    }
});

app.listen(PORT, () => {
    console.log(`Servidor corriendo en el puerto ${PORT}`);
});
