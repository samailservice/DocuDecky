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

    // Prompt corretto e rigoroso per evitare che l'IA ripeta le istruzioni
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

    // Verifica se l'IA richiede input aggiuntivi tramite JSON
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

    // Restituisce la sintesi strutturata per le slide
    return new NextResponse(textResponse, {
      headers: {
        'Content-Type': 'text/plain; charset=utf-8',
        'Content-Disposition': 'attachment; filename="presentazione-strutturata.txt"',
      },
    });

  } catch (error: any) {
    console.error('Errore API Convert:', error);
    return NextResponse.json(
      { message: error.message || 'Errore interno durante l\'elaborazione' },
      { status: 500 }
    );
  }
}
