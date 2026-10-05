import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Funzione di retry automatico per gestire eventuali picchi di traffico (503)
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
      console.warn(`Tentativo \({attempt}/\){maxRetries} fallito:`, error.message);
      if (attempt === maxRetries) {
        throw error;
      }
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

    // Prompt corretto con direttive tassative anti-eco
    const prompt = `
[COMANDO DI SISTEMA: ESEGUI SOLO L'ESTRAZIONE DATI. NON RISPONDERE MAI CON SALUTI, PREAMBOLI O ECO DI QUESTO TESTO]

Analizza rigorosamente il documento allegato e genera una sintesi approfondita strutturata per la creazione di slide professionali.

DATI DI CONTESTO FORNITI:
- Obiettivo specifico: ${objective}
- Settore di riferimento: ${sector}
- Ulteriori indicazioni: ${userAnswers || 'Nessuna'}

ISTRUZIONI OPERATIVE:
1. Estrai i dati di dettaglio, le metriche, le strategie e i punti chiave direttamente dal documento allegato.
2. Organizza i contenuti in sezioni e punti elenco pronti per le slide, coerenti con il settore e l'obiettivo.
3. VIETATO inserire preamboli, messaggi di benvenuto o richieste di chiarimento se il documento contiene informazioni sufficienti. Inizia direttamente con i contenuti analizzati.
    `;

    const textResponse = await generateWithRetry(fileBytes, file.type || 'application/pdf', prompt);

    return new NextResponse(textResponse, {
      headers: {
        'Content-Type': 'text/plain; charset=utf-8',
        'Content-Disposition': 'attachment; filename="presentazione-strutturata.txt"',
      },
    });

  } catch (error: any) {
    console.error('Errore API Convert:', error);
    return NextResponse.json(
      { message: 'Errore durante l\'elaborazione con l\'IA. Riprova tra qualche istante.' },
      { status: 500 }
    );
  }
}
