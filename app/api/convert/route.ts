import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import pptxgen from 'pptxgenjs';

// Inizializza il client Google GenAI
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

/**
 * Funzione unificata con tentativi automatici su Gemini e fallback su Groq per qualsiasi prompt testuale.
 */
async function callAIWithFallback(promptText: string) {
  const timeoutMs = 90000; // 90 secondi per evitare i limiti di Vercel
  const startTime = Date.now();
  let attempt = 0;

  // 1. Tentativo con Gemini
  while (true) {
    const elapsed = Date.now() - startTime;
    if (elapsed >= timeoutMs) {
      break; // Passa al fallback se scade il tempo
    }

    try {
      attempt++;
      const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: promptText,
      });
      if (response.text) return response.text;
    } catch (error: any) {
      const currentElapsed = Date.now() - startTime;
      if (currentElapsed >= timeoutMs) break;

      const delay = Math.min(1000 * Math.pow(2, attempt), 15000);
      console.warn(`[Gemini Tentativo \({attempt}] Server occupato o errore. Nuovo tentativo tra\){delay / 1000}s...`);
      await new Promise((resolve) => setTimeout(resolve, delay));
    }
  }

  // 2. Fallback su Groq (se Gemini fallisce o va in timeout)
  console.warn('Gemini non disponibile o timeout scaduto. Attivazione fallback su Groq...');
  const groqApiKey = process.env.GROQ_API_KEY;
  if (!groqApiKey) {
    throw new Error('Timeout Gemini raggiunto e chiave API Groq (GROQ_API_KEY) non configurata per il fallback.');
  }

  const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${groqApiKey}`,
    },
    body: JSON.stringify({
      model: 'llama-3.1-8b-instant',
      messages: [
        { role: 'system', content: 'Sei un analista aziendale esperto nella sintesi di documenti e nella creazione di presentazioni professionali.' },
        { role: 'user', content: promptText }
      ],
      temperature: 0.7,
    }),
  });

  if (!res.ok) {
    const errData = await res.json().catch(() => ({}));
    throw new Error(`Errore API Groq: ${errData.error?.message || res.statusText}`);
  }

  const data = await res.json();
  return data.choices?.[0]?.message?.content || 'Generazione completata tramite Groq';
}

export async function POST(req: Request) {
  try {
    const formData = await req.formData();
    const file = formData.get('file') as File;
    const sector = (formData.get('sector') as string) || 'Business / Aziendale';
    const objective = (formData.get('objective') as string) || 'Presentazione PPTX';

    if (!file) {
      return NextResponse.json({ message: 'Nessun file caricato' }, { status: 400 });
    }

    // Estrazione del testo dal file caricato
    const fileBuffer = await file.arrayBuffer();
    const rawText = Buffer.from(fileBuffer).toString('utf-8');
    const cleanText = rawText.replace(/[^\x20-\x7E\sÀ-ÿ]/g, ' ').substring(0, 15000);

    // ==========================================
    // STEP 1: Creazione della sintesi temporanea del documento
    // ==========================================
    console.log('Step 1: Generazione sintesi del documento in corso...');
    const step1Prompt = `
Analizza attentamente il seguente documento estratto da "${file.name}" ed estrai una sintesi dettagliata e strutturata dei punti chiave, dei dati economici/principali e delle sezioni rilevanti. 
Questa sintesi servirà come base per creare una presentazione professionale.

Documento:
${cleanText}
    `.trim();

    const documentSynthesis = await callAIWithFallback(step1Prompt);

    // ==========================================
    // STEP 2: Generazione della presentazione strutturata basata sulla sintesi
    // ==========================================
    console.log('Step 2: Generazione della struttura delle slide basata sulla sintesi...');
    const step2Prompt = `
Basandoti esclusivamente sulla seguente sintesi del documento, crea la struttura testuale completa per una presentazione professionale (\({objective}) mirata al settore "\){sector}".

L'output deve essere suddiviso in modo chiaro in slide consecutive (es. Slide 1, Slide 2, ecc.), fornendo per ciascuna un Titolo e dei Punti Elenco descrittivi e professionali.

Sintesi del documento:
${documentSynthesis}
    `.trim();

    const finalPresentationText = await callAIWithFallback(step2Prompt);

    // ==========================================
    // STEP 3: Generazione del file PowerPoint (.pptx)
    // ==========================================
    const pptx = new pptxgen();
    
    // Slide di copertina
    const coverSlide = pptx.addSlide();
    coverSlide.addText(`Presentazione: ${file.name}`, { x: 1, y: 1.5, fontSize: 24, bold: true, color: '363636' });
    coverSlide.addText(`Settore: \({sector} | Obiettivo:\){objective}`, { x: 1, y: 2.5, fontSize: 14, color: '666666' });

    // Suddivisione del testo generato per popolare le slide successive
    const contentSlide = pptx.addSlide();
    contentSlide.addText("Contenuto Chiave & Sintesi", { x: 1, y: 0.8, fontSize: 20, bold: true, color: '363636' });
    contentSlide.addText(finalPresentationText.substring(0, 1000), { x: 1, y: 1.5, fontSize: 12, color: '444444', w: '85%', h: '70%' });

    const pptxBuffer = await pptx.write({ outputType: 'arraybuffer' });

    return new NextResponse(pptxBuffer as unknown as BodyInit, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `attachment; filename="${file.name.split('.')[0] || 'documento'}-presentazione.pptx"`,
      },
    });

  } catch (error: any) {
    console.error('Errore durante la conversione procedurale:', error);
    return NextResponse.json(
      { message: error.message || 'Errore interno del server durante la generazione.' },
      { status: 500 }
    );
  }
}
