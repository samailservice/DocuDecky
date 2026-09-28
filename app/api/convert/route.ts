import { NextRequest, NextResponse } from 'next/server';
import { ai } from '../../../lib/gemini';

const SYSTEM_PROMPTS = {
  university: 'Sei il docente ed esperto accademico DocuDecky. Estrai concetti chiave, formule, definizioni e schemi di studio.',
  legal: 'Sei l’esperto legale DocuDecky. Estrai clausole di rischio, obblighi, scadenze e sintesi normative in punti chiari.',
  economy: 'Sei l’analista finanziario DocuDecky. Estrai KPI, numeri chiave, metriche di bilancio e driver di crescita.',
  it: 'Sei il Solution Architect DocuDecky. Estrai architetture, specifiche tecniche, requisiti software e roadmap.',
  hr: 'Sei l’HR Director DocuDecky. Estrai competenze, punti salienti del verbali, action item e policy aziendali.',
  medical: 'Sei il consulente medico-scientifico DocuDecky. Estrai evidenze scientifiche, sintomi, diagnosi e linee guida.',
  realestate: 'Sei l’esperto immobiliare DocuDecky. Estrai dati catastali, dettagli dell’immobile, perizie e condizioni contrattuali.'
};

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get('file') as File;
    const sector = (formData.get('sector') as string) || 'university';
    const goal = (formData.get('goal') as string) || 'presentation';

    if (!file) {
      return NextResponse.json({ error: 'Nessun file caricato.' }, { status: 400 });
    }

    const textContent = await file.text();
    const agentPrompt = SYSTEM_PROMPTS[sector] || SYSTEM_PROMPTS.university;

    const prompt = `
    ${agentPrompt}
    
    Analizza il seguente testo estratto dal documento e genera una risposta in formato JSON con la seguente struttura:
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

    Obiettivo richiesto dall'utente: ${goal}.
    Testo del Documento:
    ${textContent.substring(0, 15000)}
    `;

    const response = await ai.models.generateContent({
      model: 'gemini-1.5-flash',
      contents: prompt,
    });

    const responseText = response.text || '';
    const cleanJson = responseText.replace(/```json|```/g, '').trim();
    const parsedData = JSON.parse(cleanJson);

    return NextResponse.json({ success: true, data: parsedData });
  } catch (error: any) {
    return NextResponse.json({ error: error.message || 'Errore nella conversione' }, { status: 500 });
  }
}
