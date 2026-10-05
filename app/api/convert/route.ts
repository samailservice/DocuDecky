import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import pptxgen from 'pptxgenjs';

// Inizializza il client Google GenAI
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

/**
 * Esegue la chiamata a Gemini con tentativi automatici in caso di server occupato
 * e un timeout massimo di 5 minuti (300.000 ms), gestendo direttamente file nativi (PDF, DOCX, ecc.).
 */
async function callGeminiWithRetry(base64Data: string, mimeType: string, promptText: string, timeoutMs = 300000) {
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
        contents: [
          {
            inlineData: {
              data: base64Data,
              mimeType: mimeType,
            },
          },
          {
            text: promptText,
          },
        ],
      });
      return response.text || '';
    } catch (error: any) {
      const currentElapsed = Date.now() - startTime;
      if (currentElapsed >= timeoutMs) {
        throw new Error('TIMEOUT_EXCEEDED');
      }

      // Backoff esponenziale con un'attesa massima di 30 secondi tra i tentativi
      const delay = Math.min(1000 * Math.pow(2, attempt), 30000);
      console.warn(`[Gemini Tentativo \({attempt}] Server occupato o errore:\){error.message || error}. Nuovo tentativo tra ${delay / 1000}s...`);
      await new Promise((resolve) => setTimeout(resolve, delay));
    }
  }
}

/**
 * Funzione di fallback che interroga ChatGPT (OpenAI API) al termine del timeout di Gemini.
 */
async function callChatGPTFallback(fileName: string, promptText: string) {
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
        { role: 'user', content: `Il documento si chiama "\({fileName}".\){promptText}` }
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

    // Conversione del file in Base64 per consentire a Gemini di leggere nativamente PDF e documenti
    const fileBuffer = await file.arrayBuffer();
    const base64Data = Buffer.from(fileBuffer).toString('base64');
    const mimeType = file.type || 'application/pdf';

    const prompt = `Analizza questo documento (\({file.name}) e strutturalo per una presentazione professionale (\){objective}) mirata al settore "${sector}". 
Estrai i punti chiave suddividendoli in slide chiare, con titoli e punti elenco strutturati.`;

    let resultText = '';

    try {
      // 1. Tenta la generazione con Gemini (con retry, timeout di 5 minuti e supporto nativo file)
      resultText = await callGeminiWithRetry(base64Data, mimeType, prompt, 300000);
    } catch (geminiError: any) {
      // 2. Se scade il timeout o fallisce permanentemente, passa a ChatGPT
      console.warn('Gemini non disponibile o timeout scaduto. Reindirizzamento a ChatGPT...');
      resultText = await callChatGPTFallback(file.name, prompt);
    }

    // Generazione del file PowerPoint (.pptx)
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
