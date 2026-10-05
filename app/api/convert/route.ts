import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import pptxgen from 'pptxgenjs';

// Inizializza il client Google GenAI
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

async function callGeminiWithRetry(base64Data: string, mimeType: string, promptText: string, timeoutMs = 90000) {
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

      const delay = Math.min(1000 * Math.pow(2, attempt), 15000);
      console.warn(`[Gemini Tentativo \({attempt}] Server occupato o errore. Nuovo tentativo tra\){delay / 1000}s...`);
      await new Promise((resolve) => setTimeout(resolve, delay));
    }
  }
}

/**
 * Funzione di fallback che interroga Groq API con il modello classico e universale llama3-8b-8192.
 */
async function callGroqFallback(fileName: string, promptText: string) {
  const apiKey = process.env.GROQ_API_KEY;
  if (!apiKey) {
    throw new Error('Timeout Gemini raggiunto e chiave API Groq (GROQ_API_KEY) non configurata per il fallback.');
  }

  console.log('Attivazione fallback gratuito su Groq in corso...');
  const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: 'llama3-8b-8192',
      messages: [
        { role: 'system', content: 'Sei un assistente esperto nella strutturazione di presentazioni professionali.' },
        { role: 'user', content: `Il documento si chiama "\({fileName}".\){promptText}` }
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

    const fileBuffer = await file.arrayBuffer();
    const base64Data = Buffer.from(fileBuffer).toString('base64');
    const mimeType = file.type || 'application/pdf';

    const prompt = `Analizza questo documento (\({file.name}) e strutturalo per una presentazione professionale (\){objective}) mirata al settore "${sector}". 
Estrai i punti chiave suddividendoli in slide chiare, con titoli e punti elenco strutturati.`;

    let resultText = '';

    try {
      resultText = await callGeminiWithRetry(base64Data, mimeType, prompt, 90000);
    } catch (geminiError: any) {
      console.warn('Gemini non disponibile o timeout scaduto. Reindirizzamento a Groq...');
      resultText = await callGroqFallback(file.name, prompt);
    }

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
