import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import pptxgen from 'pptxgenjs';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

/**
 * STEP 2: Sintesi Agente AI con logica esatta del workflow
 * - Loop su Gemini (Fallimento < 3 volte)
 * - Switch su Groq (Fallimento >= 3 volte)
 */
async function getAISynthesisWithWorkflow(promptText: string): Promise<{ provider: string; text: string }> {
  let attempt = 0;
  const maxGeminiAttempts = 3;

  // Loop su Gemini (Tentativi < 3)
  while (attempt < maxGeminiAttempts) {
    try {
      attempt++;
      console.log(`[Step 2] Gemini - Tentativo ${attempt}/3 in corso...`);
      const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: promptText,
      });
      if (response.text) {
        return { provider: 'Gemini', text: response.text };
      }
    } catch (error: any) {
      console.warn(`[Step 2] Gemini - Tentativo ${attempt} fallito:`, error.message);
      if (attempt >= maxGeminiAttempts) {
        console.warn('[Step 2] Fallimento >= 3 volte su Gemini. Switch automatico su Groq...');
        break;
      }
      // Attesa breve prima del prossimo tentativo nel loop
      await new Promise((resolve) => setTimeout(resolve, 2000));
    }
  }

  // Switch su Groq (Fallback dopo 3 tentativi falliti su Gemini)
  const groqApiKey = process.env.GROQ_API_KEY;
  if (!groqApiKey) {
    throw new Error('Gemini non disponibile dopo 3 tentativi e chiave API Groq (GROQ_API_KEY) non configurata.');
  }

  console.log('[Step 2] Switch su Groq (llama-3.1-8b-instant) in corso...');
  const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${groqApiKey}`,
    },
    body: JSON.stringify({
      model: 'llama-3.1-8b-instant',
      messages: [
        { role: 'system', content: 'Sei un analista aziendale esperto nella sintesi di documenti.' },
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
  const groqText = data.choices?.[0]?.message?.content;
  if (!groqText) throw new Error('Risposta vuota ricevuta da Groq');

  return { provider: 'Groq', text: groqText };
}

export async function POST(req: Request) {
  try {
    const formData = await req.formData();
    const file = formData.get('file') as File;
    const sector = (formData.get('sector') as string) || 'Business / Aziendale';
    const objective = (formData.get('objective') as string) || 'Presentazione PPTX';
    const outputFormat = (formData.get('format') as string) || 'pptx'; // 'pptx' o 'docx'

    if (!file) {
      return NextResponse.json({ message: 'Nessun file caricato' }, { status: 400 });
    }

    // ====================================================
    // STEP 1: UPLOAD & ANALISI DOCUMENTO (Parsing)
    // ====================================================
    const fileBuffer = await file.arrayBuffer();
    const rawText = Buffer.from(fileBuffer).toString('utf-8');
    const cleanText = rawText.replace(/[^\x20-\x7E\sÀ-ÿ]/g, ' ').substring(0, 15000);

    const analysisPrompt = `Analizza attentamente il documento estratto da "\({file.name}" ed estrai una sintesi dettagliata dei punti chiave, dei dati finanziari e delle sezioni rilevanti:\n\n\){cleanText}`;

    // ====================================================
    // STEP 2 & 3: SINTESI AGENTE AI & PROMPT INTERMEDIO
    // ====================================================
    const { provider, text: synthesisText } = await getAISynthesisWithWorkflow(analysisPrompt);
    
    // Creazione del Prompt Intermedio basato sulla sintesi pulita
    const intermediatePrompt = `
Basandoti sulla sintesi chiave del documento, genera la struttura finale completa per l'obiettivo "\({objective}" nel settore "\){sector}". 
Fornisci i contenuti suddivisi in sezioni e slide dettagliate con titoli e punti elenco professionali.

Sintesi di riferimento:
${synthesisText}
    `.trim();

    // Elaborazione finale sfruttando lo stesso motore resiliente del workflow
    const { text: finalContent } = await getAISynthesisWithWorkflow(intermediatePrompt);

    // ====================================================
    // STEP 4 & 5: SCELTA OUTPUT & DOWNLOAD AUTOMATICO
    // ====================================================
    if (outputFormat === 'docx' || objective.toLowerCase().includes('docx') || objective.toLowerCase().includes('pitch')) {
      // Diramazione: Generazione Documento / Pitch (.docx)
      const docBuffer = Buffer.from(`REPORT DI SINTESI / PITCH\nMotore IA: \({provider}\nSettore:\){sector}\n\n${finalContent}`, 'utf-8');
      return new NextResponse(docBuffer, {
        headers: {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          'Content-Disposition': `attachment; filename="${file.name.split('.')[0] || 'documento'}-sintesi.docx"`,
        },
      });
    } else {
      // Diramazione: Generazione Presentazione Grafica (.pptx)
      const pptx = new pptxgen();
      
      const coverSlide = pptx.addSlide();
      coverSlide.addText(`Presentazione: ${file.name}`, { x: 1, y: 1.5, fontSize: 24, bold: true, color: '363636' });
      coverSlide.addText(`Settore: \({sector} | Elaborato con:\){provider}`, { x: 1, y: 2.5, fontSize: 13, color: '666666' });

      const contentSlide = pptx.addSlide();
      contentSlide.addText("Contenuti e Punti Chiave", { x: 1, y: 0.8, fontSize: 20, bold: true, color: '363636' });
      contentSlide.addText(finalContent.substring(0, 1500), { x: 1, y: 1.5, fontSize: 12, color: '444444', w: '85%', h: '70%' });

      const pptxBuffer = await pptx.write({ outputType: 'arraybuffer' });

      return new NextResponse(pptxBuffer as unknown as BodyInit, {
        headers: {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          'Content-Disposition': `attachment; filename="${file.name.split('.')[0] || 'documento'}-presentazione.pptx"`,
        },
      });
    }

  } catch (error: any) {
    console.error('Errore nel workflow di automazione:', error);
    return NextResponse.json(
      { message: error.message || 'Errore interno del server durante l\'elaborazione.' },
      { status: 500 }
    );
  }
}
