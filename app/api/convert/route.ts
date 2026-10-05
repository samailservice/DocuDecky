import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import pptxgen from 'pptxgenjs';

// Inizializza il client Google GenAI
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

/**
 * Esegue la chiamata a Gemini con tentativi automatici in caso di server occupato
 * e un timeout massimo di 5 minuti (300.000 ms).
 */
async function callGeminiWithRetry(promptText: string, timeoutMs = 300000) {
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

      // Backoff esponenziale con un'attesa massima di 30 secondi tra i tentativi
      const delay = Math.min(1000 * Math.pow(2, attempt), 30000);
      console.warn(`[Gemini Tentativo \({attempt}] Server occupato o errore. Nuovo tentativo tra\){delay / 1000}s...`);
      await new Promise((resolve) => setTimeout(resolve, delay));
    }
  }
}

/**
 * Funzione di fallback che interroga ChatGPT (OpenAI API) al termine del timeout di Gemini.
 */
async function callChatGPTFallback(promptText: string) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    throw new Error('Timeout Gemini raggiunto e chiave API OpenAI (OPENAI_API_KEY) non configurata per il fallback.');
  }

  console.log('Attivazione fallback su ChatGPT in corso...');
  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: 'gpt-4o-mini',
      messages: [
        { role: 'system', content: 'Sei un assistente esperto nella strutturazione di presentazioni professionali.' },
        { role: 'user', content: promptText }
      ],
      temperature: 0.7,
    }),
  });

  if (!res.ok) {
    const errData = await res.json().catch(() => ({}));
    throw new Error(`Errore API ChatGPT: ${errData.error?.message || res.statusText}`);
  }

  const data = await res.json();
  return data.choices?.[0]?.message?.content || 'Generazione completata tramite ChatGPT';
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

    const fileBuffer = await file.arrayBuffer();
    const textContent = Buffer.from(fileBuffer).toString('utf-8');

    const prompt = `
      Analizza il seguente documento e strutturalo per una presentazione professionale (\({objective}) mirata al settore "\){sector}".
      Estrai i punti chiave suddividendoli in slide con titoli e punti elenco chiari.
      Documento:
      ${textContent.substring(0, 15000)}
    `;

    let resultText = '';

    try {
      // 1. Tenta la generazione con Gemini (con retry e timeout di 5 minuti)
      resultText = await callGeminiWithRetry(prompt, 300000);
    } catch (geminiError: any) {
      // 2. Se scade il timeout, passa a ChatGPT
      console.warn('Gemini non disponibile o timeout scaduto. Reindirizzamento a ChatGPT...');
      resultText = await callChatGPTFallback(prompt);
    }

    // Generazione del file PowerPoint (.pptx)
    const pptx = new pptxgen();
    
    const slide = pptx.addSlide();
    slide.addText(`Presentazione: ${file.name}`, { x: 1, y: 1, fontSize: 22, bold: true, color: '363636' });
    slide.addText(resultText.substring(0, 1000), { x: 1, y: 2, fontSize: 13, color: '555555', w: '80%' });

    // Modificato in 'arraybuffer' per soddisfare i requisiti TypeScript di Next.js (BodyInit)
    const pptxBuffer = await pptx.write({ outputType: 'arraybuffer' });

    return new NextResponse(pptxBuffer, {
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
