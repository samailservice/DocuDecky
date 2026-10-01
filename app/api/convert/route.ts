import { NextRequest, NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';

const apiKey = process.env.GEMINI_API_KEY || '';
const ai = new GoogleGenAI({ apiKey });

// Modello principale gemini-3.8-flash seguito dai modelli di riserva validi
const CANDIDATE_MODELS = [
  'gemini-3.8-flash',
  'gemini-2.0-flash',
  'gemini-1.5-flash',
];

const SYSTEM_PROMPTS: { [key: string]: string } = {
  'Business / Aziendale': 'Sei l’analista finanziario e business DocuDecky. Estrai KPI, metriche di bilancio, punti di forza e sintesi esecutiva.',
  'Educativo / Accademico': 'Sei il docente ed esperto accademico DocuDecky. Estrai concetti chiave, formule, definizioni e schemi di studio.',
  'Tecnologico / Startup': 'Sei il Solution Architect e Startup Mentor DocuDecky. Estrai architetture, specifiche tecniche e roadmap.',
  'Creativo / Marketing': 'Sei il Marketing Director DocuDecky. Estrai strategie di posizionamento, target, valori chiave e punti salienti.',
  university: 'Sei il docente ed esperto accademico DocuDecky. Estrai concetti chiave, formule, definizioni e schemi di studio.',
  legal: 'Sei l’esperto legale DocuDecky. Estrai clausole di rischio, obblighi, scadenze e sintesi normative in punti chiari.',
  economy: 'Sei l’analista finanziario DocuDecky. Estrai KPI, numeri chiave, metriche di bilancio e driver di crescita.',
  it: 'Sei il Solution Architect DocuDecky. Estrai architetture, specifiche tecniche, requisiti software e roadmap.',
  hr: 'Sei l’HR Director DocuDecky. Estrai competenze, punti salienti dei verbali, action item e policy aziendali.',
  medical: 'Sei il consulente medico-scientifico DocuDecky. Estrai evidenze scientifiche, sintomi, diagnosi e linee guida.',
  realestate: 'Sei l’esperto immobiliare DocuDecky. Estrai dati catastali, dettagli dell’immobile, perizie e condizioni contrattuali.'
};

async function generateContentWithFallback(contents: any[]) {
  let lastError: any = null;

  for (const model of CANDIDATE_MODELS) {
    try {
      const response = await ai.models.generateContent({
        model: model,
        contents: contents,
      });

      if (response && response.text) {
        return response.text;
      }
    } catch (err: any) {
      console.warn(`Modello ${model} non disponibile o in errore. Prova modello successivo:`, err?.message || err);
      lastError = err;
    }
  }

  throw lastError || new Error('Tutti i modelli AI configurati hanno restituito errore.');
}

export async function POST(req: NextRequest) {
  try {
    if (!process.env.GEMINI_API_KEY) {
      return NextResponse.json(
        { message: 'Chiave GEMINI_API_KEY mancante su Vercel. Aggiungila nelle Environment Variables.' },
        { status: 500 }
      );
    }

    const formData = await req.formData();
    const file = formData.get('file') as File;
    const sector = (formData.get('sector') as string) || 'Business / Aziendale';
    const objective = (formData.get('objective') as string) || (formData.get('goal') as string) || 'Presentazione PPTX';

    if (!file) {
      return NextResponse.json({ message: 'Nessun file caricato.' }, { status: 400 });
    }

    const agentPrompt = SYSTEM_PROMPTS[sector] || SYSTEM_PROMPTS['Business / Aziendale'];

    const promptText = `
${agentPrompt}

Analizza il seguente documento e genera una risposta ESCLUSIVAMENTE in formato JSON valido con questa struttura:
{
  "title": "Titolo principale",
  "summary": "Sintesi esecutiva in 3 punti",
  "slides": [
    {
      "slideTitle": "Titolo Diapositiva",
      "bulletPoints": ["Punto 1", "Punto 2", "Punto 3"]
    }
  ],
  "promptHelpers": [
    "Prompt pronto 1 per approfondire su ChatGPT",
    "Prompt pronto 2 per espandere il testo",
    "Prompt pronto 3 per simulare domande"
  ]
}

Obiettivo richiesto: ${objective}.
    `;

    let contents: any[];

    if (file.type.includes('pdf') || file.name.endsWith('.pdf')) {
      const arrayBuffer = await file.arrayBuffer();
      const base64Data = Buffer.from(arrayBuffer).toString('base64');

      contents = [
        {
          inlineData: {
            mimeType: 'application/pdf',
            data: base64Data,
          },
        },
        promptText,
      ];
    } else {
      const textContent = await file.text();
      contents = [
        `\({promptText}\n\nTesto del Documento:\n\){textContent.substring(0, 30000)}`,
      ];
    }

    const responseText = await generateContentWithFallback(contents);
    const cleanJson = responseText.replace(/```json|```/g, '').trim();

    let parsedData;
    try {
      parsedData = JSON.parse(cleanJson);
    } catch {
      parsedData = {
        title: file.name,
        summary: responseText,
        slides: [],
        promptHelpers: [],
      };
    }

    return NextResponse.json({ success: true, data: parsedData });
  } catch (error: any) {
    console.error('Errore durante la conversione:', error);

    let cleanErrorMessage = 'Errore nell’elaborazione con l’IA. Riprova tra qualche istante.';
    if (typeof error?.message === 'string') {
      try {
        const parsedErr = JSON.parse(error.message);
        cleanErrorMessage = parsedErr?.error?.message || error.message;
      } catch {
        cleanErrorMessage = error.message;
      }
    }

    return NextResponse.json(
      { message: cleanErrorMessage },
      { status: 500 }
    );
  }
}
