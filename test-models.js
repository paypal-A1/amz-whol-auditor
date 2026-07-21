const { GoogleGenAI } = require('@google/genai');
require('dotenv').config();

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Lista de modelos a probar
const modelos = [
    'gemini-1.5-flash',
    'gemini-2.5-flash',
    'gemini-2.5-flash-lite',
    'gemini-2.5-pro',
    'gemini-3.5-flash',
    'gemini-3.5-flash-lite',
    'gemini-3.6-flash',
    'gemini-2.0-flash',
    'gemini-2.0-flash-lite'
];

// Función para probar disponibilidad de un modelo
async function testDisponibilidad(modelo) {
    try {
        const response = await ai.models.generateContent({
            model: modelo,
            contents: 'Di solo la palabra "hola" en una línea.',
            config: { responseMimeType: 'text/plain' }
        });
        return { disponible: true, error: null };
    } catch (error) {
        return { disponible: false, error: error.message, status: error.status };
    }
}

// Función para medir RPM (Requests Per Minute)
async function testRPM(modelo, maxRequests = 25) {
    let requests = 0;
    let limiteRPM = null;
    let errores = [];

    console.log(`\n📡 Probando RPM para ${modelo}...`);

    for (let i = 1; i <= maxRequests; i++) {
        try {
            await ai.models.generateContent({
                model: modelo,
                contents: 'Responde solo con la palabra "ok" en una línea.',
                config: { responseMimeType: 'text/plain' }
            });
            requests++;
            process.stdout.write(`.`); // Indicador visual
        } catch (error) {
            if (error.status === 429) {
                // Extraer el límite del mensaje de error
                const match = error.message.match(/limit:\s*(\d+)/i);
                if (match) {
                    limiteRPM = parseInt(match[1]);
                }
                errores.push({ intento: i, error: error.message });
                console.log(`\n⛔ Error 429 en intento ${i}. Límite detectado: ${limiteRPM || 'desconocido'} RPM`);
                break;
            } else {
                errores.push({ intento: i, error: error.message });
                console.log(`\n❌ Error ${error.status} en intento ${i}: ${error.message}`);
                break;
            }
        }
        // Pequeño delay para no saturar demasiado rápido (0.5s)
        await new Promise(resolve => setTimeout(resolve, 500));
    }

    return { requests, limiteRPM, errores };
}

// Función principal
async function main() {
    console.log('🚀 Iniciando prueba de modelos Gemini...\n');
    console.log(`📋 Modelos a probar: ${modelos.length}\n`);

    const resultados = [];

    for (const modelo of modelos) {
        console.log(`\n🔍 Probando modelo: ${modelo}`);
        
        // Paso 1: Disponibilidad
        const { disponible, error, status } = await testDisponibilidad(modelo);
        
        if (!disponible) {
            console.log(`❌ No disponible: ${error}`);
            resultados.push({
                modelo,
                disponible: false,
                status,
                rpm: null,
                error
            });
            continue;
        }

        console.log(`✅ Disponible`);

        // Paso 2: Medir RPM
        const { requests, limiteRPM, errores } = await testRPM(modelo);
        
        console.log(`📊 Solicitudes exitosas: ${requests}`);
        console.log(`📊 Límite RPM detectado: ${limiteRPM || 'No detectado (posiblemente sin límite o límite alto)'}`);

        resultados.push({
            modelo,
            disponible: true,
            status: 200,
            rpm: limiteRPM,
            requestsExitosas: requests,
            errores: errores.length > 0 ? errores : null
        });
    }

    // Mostrar resumen final
    console.log('\n\n📊 ===== RESUMEN FINAL =====');
    console.table(resultados.map(r => ({
        Modelo: r.modelo,
        Disponible: r.disponible ? '✅' : '❌',
        'Status': r.status || 'N/A',
        'RPM (real)': r.rpm || 'N/A',
        'Solicitudes OK': r.requestsExitosas || 0
    })));

    console.log('\n✅ Prueba completada.');
}

main().catch(console.error);
