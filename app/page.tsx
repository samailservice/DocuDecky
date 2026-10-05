import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

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

    // Prompt strutturato secondo le direttive del workflow (Punto 1)
    const prompt = `
Sei l'assistente IA di DocuDecky. Il tuo compito è analizzare il documento caricato ed estrarre i dati di dettaglio e la sintesi chiave combinandoli con i seguenti dati di contesto:

1. **Obiettivo specifico**: ${objective}
2. **Settore di riferimento**: ${sector}
3. **Ulteriori chiarimenti / Risposte dell'utente**: ${userAnswers || 'Nessuno finora'}

Istruzioni per l'analisi (Punto 1 del workflow):
- Leggi attentamente il documento allegato.
- Estrai i dati di dettaglio, le strategie chiave, i risultati, le metriche e i dati finanziari rilevanti.
- Collega le informazioni estratte all'obiettivo e al settore specificati.

Se ritieni che manchino informazioni fondamentali per procedere con una presentazione di alta qualità, restituisci ESCLUSIVAMENTE un oggetto JSON in questo formato esatto (senza altri testi attorno):
{
  "needsInput": true,
  "questions": [
    "Domanda specifica 1 per chiarire un punto mancante",
    "Domanda specifica 2..."
  ]
}

Se invece hai tutte le informazioni necessarie, genera una sintesi approfondita e strutturata dei contenuti che andranno a comporre le slide.
    `;

    // Chiamata all'API Gemini utilizzando il modello supportato
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [
        {
          inlineData: {
            mimeType: file.type || 'application/pdf',
            data: fileBytes.toString('base64'),
          },
        },
        prompt,
      ],
    });

    const textResponse = response.text || '';

    // Verifica se l'IA richiede ulteriori input dall'utente
    try {
      const jsonMatch = textResponse.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        const parsed = JSON.parse(jsonMatch[0]);
        if (parsed.needsInput && Array.isArray(parsed.questions)) {
          return NextResponse.json(parsed);
        }
      }
    } catch (e) {
      // Se non è JSON valido, procediamo con il flusso standard
    }

    // Risposta di successo con i dati elaborati
    return new NextResponse(textResponse, {
      headers: {
        'Content-Type': 'text/plain; charset=utf-8',
        'Content-Disposition': 'attachment; filename="presentazione-strutturata.txt"',
      },
    });

  } catch (error: any) {
    console.error('Errore API Convert:', error);
    return NextResponse.json(
      { message: error.message || 'Errore interno durante l\'elaborazione con l\'IA' },
      { status: 500 }
    );
  }
}
