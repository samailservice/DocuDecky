import { NextRequest, NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';

// Inizializzazione del client Gemini
const apiKey = process.env.GEMINI_API_KEY || '';
const ai = new GoogleGenAI({ apiKey });

const SYSTEM_PROMPTS: Record = {
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

    // Gestione nativa dei file PDF tramite Base64
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
      // Per file di testo (.txt, .md, ecc.)
      const textContent = await file.text();
      contents = [
        `\({promptText}\n\nTesto del Documento:\n\){textContent.substring(0, 30000)}`,
      ];
    }

    // Utilizzo del modello ufficiale valido gemini-2.5-flash
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: contents,
    });

    const responseText = response.text || '';
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
    return NextResponse.json(
      { message: error?.message || 'Errore interno durante l\'elaborazione.' },
      { status: 500 }
    );
  }
}
