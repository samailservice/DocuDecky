import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import pptxgen from 'pptxgenjs';

// Inizializza il client Google GenAI
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

/**
 * Step 1 & 2: Chiamata a Gemini basata su testo puro (con tentativi automatici e timeout a 90s)
 */
async function callGeminiTextWithRetry(promptText: string, timeoutMs = 90000) {
  const startTime = Date.now();
  let attempt = 0;

  while (true) {
    const elapsed = Date.now() - startTime;
    if (elapsed >= timeoutMs) {
      throw new Error('TIMEOUT_EXCEEDED');
    }

    try {
      attempt++;
      const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: promptText,
      });
      return response.text || '';
    } catch (error: any) {
      const currentElapsed = Date.now() - startTime;
      if (currentElapsed >= timeoutMs) {
        throw new Error('TIMEOUT_EXCEEDED');
      }

      const delay = Math.min(1000 * Math.pow(2, attempt), 15000);
      console.warn(`[Gemini Tentativo \({attempt}] Server occupato o errore. Nuovo tentativo tra\){delay / 1000}s...`);
      await new Promise((resolve) => setTimeout(resolve, delay));
    }
  }
}

/**
 * Fallback su Groq basato su testo puro (usa il modello universale llama3-8b-8192)
 */
async function callGroqTextFallback(promptText: string) {
  const apiKey = process.env.GROQ_API_KEY;
  if (!apiKey) {
    throw new Error('Timeout Gemini raggiunto e chiave API Groq (GROQ_API_KEY) non configurata per il fallback.');
  }

  console.log('Attivazione fallback testuale su Groq in corso...');
  const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: 'llama3-8b-8192',
      messages: [
        { role: 'system', content: 'Sei un assistente esperto nella strutturazione di presentazioni professionali e analisi di bilancio.' },
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

    // Estrazione del testo dal file caricato (gestione sicura del buffer)
    const fileBuffer = await file.arrayBuffer();
    const rawText = Buffer.from(fileBuffer).toString('utf-8');
    
    // Pulizia di base per rimuovere caratteri binari superflui se il file è un PDF
    const cleanText = rawText.replace(/[^\x20-\x7E\sÀ-ÿ]/g, ' ').substring(0, 15000);

    // Creazione del prompt strutturato per l'IA
    const prompt = `
Sei un analista aziendale ed esperto di comunicazione. 
Analizza il seguente documento estratto da "\({file.name}" e crea la struttura testuale per una presentazione professionale (\){objective}) mirata al settore "${sector}".

Suddividi chiaramente l'output in slide (es. Slide 1: Titolo, Punti chiave, ecc.).

Contenuto del documento:
${cleanText}
    `.trim();

    let resultText = '';

    try {
      // 1. Tenta la generazione testuale con Gemini
      resultText = await callGeminiTextWithRetry(prompt, 90000);
    } catch (geminiError: any) {
      // 2. Se Gemini fallisce o va in timeout, passa a Groq con lo stesso identico prompt testuale
      console.warn('Gemini non disponibile o timeout scaduto. Reindirizzamento a Groq...');
      resultText = await callGroqTextFallback(prompt);
    }

    // Generazione del file PowerPoint (.pptx) con pptxgenjs
    const pptx = new pptxgen();
    
    const slide = pptx.addSlide();
    slide.addText(`Presentazione: ${file.name}`, { x: 1, y: 1, fontSize: 22, bold: true, color: '363636' });
    slide.addText(resultText.substring(0, 1000), { x: 1, y: 2, fontSize: 13, color: '555555', w: '80%' });

    const pptxBuffer = await pptx.write({ outputType: 'arraybuffer' });

    return new NextResponse(pptxBuffer as unknown as BodyInit, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `attachment; filename="${file.name.split('.')[0] || 'documento'}-presentazione.pptx"`,
      },
    });

  } catch (error: any) {
    console.error('Errore durante la conversione:', error);
    return NextResponse.json(
      { message: error.message || 'Errore interno del server durante la generazione.' },
      { status: 500 }
    );
  }
}
