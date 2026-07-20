const express = require('express');
const dotenv = require('dotenv');
const { GoogleGenAI } = require('@google/genai');
const path = require('path');

// Cargar variables de entorno (como la API Key de Gemini)
dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

// Inicializar el cliente oficial moderno de Gemini
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Middleware para entender JSON y servir la carpeta pública
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Ruta principal para procesar los datos de Amazon Wholesale y Keepa
app.post('/api/audit', async (req, res) => {
    try {
        const { keepaData } = req.body;

        if (!keepaData) {
            return res.status(400).json({ error: 'Faltan los datos de Keepa para el análisis.' });
        }

        // Llamada usando la arquitectura de interacciones optimizada para agentes
        const interaction = await ai.interactions.create({
            model: "gemini-3.5-flash",
            systemInstruction: 
                "Eres un auditor experto en Amazon Wholesale. Analiza los datos de Keepa provistos. " +
                "Evalúa: 1) Estabilidad del precio (si hay Price Tanking). 2) Comportamiento de la Buy Box. " +
                "3) Nivel de competencia de vendedores FBA. Genera un diagnóstico preciso y directo.",
            input: `Analiza este bloque de datos del producto: ${JSON.stringify(keepaData)}`,
            generationConfig: {
                thinkingLevel: "medium", // Nivel balanceado ideal para evaluar lógica comercial
                responseMimeType: "application/json" // Asegura recibir un JSON limpio
            }
        });

        // Retornar la respuesta al frontend
        const result = JSON.parse(interaction.output_text);
        res.json(result);

    } catch (error) {
        console.error("Error en el servidor:", error);
        res.status(500).json({ error: 'Hubo un error al procesar el análisis con Gemini.' });
    }
});

// Iniciar el servidor
app.listen(PORT, () => {
    console.log(`Servidor corriendo en el puerto ${PORT}`);
});
