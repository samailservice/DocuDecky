import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Funzione di supporto per gestire i tentativi multipli in caso di sovraccarico (503)
async function generateWithRetry(fileBytes: Buffer, mimeType: string, prompt: string, maxRetries = 3) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = await ai.models.generateContent({
        model: 'gemini-3.8-flash',
        contents: [
          {
            inlineData: {
              mimeType: mimeType,
              data: fileBytes.toString('base64'),
            },
          },
          prompt,
        ],
      });
      return response.text || '';
    } catch (error: any) {
      console.warn(`⚠️ Tentativo \({attempt}/\){maxRetries} fallito (Status: ${error?.status || '503/Altro'}):`, error.message);
      
      // Se è l'ultimo tentativo, rilancia l'errore
      if (attempt === maxRetries) {
        throw error;
      }
      
      // Attesa esponenziale prima del prossimo tentativo (es. 1.5s, 3s)
      await new Promise((resolve) => setTimeout(resolve, attempt * 1500));
    }
  }
  throw new Error('Superato il limite massimo di tentativi con Gemini.');
}

export async function POST(req: Request) {
  try {
    const formData = await req.formData();
    const file = formData.get('file') as File | null;
    const sector = (formData.get('sector') as string) || 'Business / Aziendale';
    const objective = (formData.get('objective') as string) || 'Presentazione PPTX';
    const userAnswers = (formData.get('userAnswers') as string) || '';

    if (!file) {
      return NextResponse.json({ message: 'Nessun file caricato' }, { status: 400 });
    }

    const fileBuffer = await file.arrayBuffer();
    const fileBytes = Buffer.from(fileBuffer);

    const prompt = `
Agisci come un esperto analista aziendale e content strategist per presentazioni professionali.
Analizza il documento allegato e genera una sintesi approfondita strutturata per la creazione di slide, integrando rigorosamente questi dati forniti dall'utente:

- **Obiettivo specifico**: ${objective}
- **Settore di riferimento**: ${sector}
- **Ulteriori chiarimenti / Risposte**: ${userAnswers || 'Nessuna'}

Istruzioni operative:
1. Estrai i punti chiave, i dati di dettaglio, le strategie e le metriche finanziarie o operative direttamente dal documento allegato.
2. Adatta e organizza i contenuti in sezioni e punti elenco professionali coerenti con il settore (\({sector}) e l'obiettivo (\){objective}).
3. **Importante**: Non restituire messaggi di benvenuto, preamboli o richieste generiche. Produci direttamente la struttura dei contenuti delle slide in modo dettagliato.
4. Se e solo se mancano informazioni assolutamente cruciali nel documento per procedere, restituisci ESCLUSIVAMENTE un oggetto JSON nel formato:
{
  "needsInput": true,
  "questions": ["Domanda 1", "Domanda 2"]
}
    `;

    // Esegue la chiamata sfruttando il meccanismo di retry automatico (fino a 3 tentativi)
    const textResponse = await generateWithRetry(fileBytes, file.type || 'application/pdf', prompt);

    try {
      const jsonMatch = textResponse.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        const parsed = JSON.parse(jsonMatch[0]);
        if (parsed.needsInput && Array.isArray(parsed.questions)) {
          return NextResponse.json(parsed);
        }
      }
    } catch (e) {
      // Ignora se non è un JSON valido e procedi con il testo normale
    }

    return new NextResponse(textResponse, {
      headers: {
        'Content-Type': 'text/plain; charset=utf-8',
        'Content-Disposition': 'attachment; filename="presentazione-strutturata.txt"',
      },
    });

  } catch (error: any) {
    console.error('Errore API Convert Definitivo:', error);
    return NextResponse.json(
      { message: 'I server di Gemini sono momentaneamente sovraccarichi (503). Riprova tra qualche istante.' },
      { status: 503 }
    );
  }
}
