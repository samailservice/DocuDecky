import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import pptxgen from 'pptxgenjs';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

async function callAIWithWorkflow(promptText: string): Promise<{ provider: string; text: string }> {
  let attempt = 0;
  const maxGeminiAttempts = 3;

  while (attempt < maxGeminiAttempts) {
    try {
      attempt++;
      console.log(`[AI Workflow] Gemini - Tentativo ${attempt}/3 in corso...`);
      const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: promptText,
      });
      if (response.text) {
        return { provider: 'Gemini', text: response.text };
      }
    } catch (error: any) {
      console.warn(`[AI Workflow] Gemini - Tentativo ${attempt} fallito:`, error.message);
      if (attempt >= maxGeminiAttempts) {
        console.warn('[AI Workflow] Fallimento >= 3 volte su Gemini. Switch automatico su Groq...');
        break;
      }
      await new Promise((resolve) => setTimeout(resolve, 2000));
    }
  }

  const groqApiKey = process.env.GROQ_API_KEY;
  if (!groqApiKey) {
    throw new Error('Gemini non disponibile dopo 3 tentativi e chiave API Groq (GROQ_API_KEY) non configurata.');
  }

  console.log('[AI Workflow] Switch su Groq (openai/gpt-oss-20b) in corso...');
  const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${groqApiKey}`,
    },
    body: JSON.stringify({
      model: 'openai/gpt-oss-20b',
      messages: [
        { role: 'system', content: 'Sei un analista aziendale esperto nella creazione di presentazioni e report strategici.' },
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
    const outputFormat = (formData.get('format') as string) || 'pptx';
    const userAnswers = (formData.get('userAnswers') as string) || '';

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
    const { provider, text: synthesisText } = await callAIWithWorkflow(analysisPrompt);

    // ====================================================
    // STEP 4: SCELTA OUTPUT & GENERAZIONE CON I 3 PILASTRI
    // ====================================================
    // Pulizia delle stringhe condizionali per evitare errori di parsing nel template literal
    const userNotesSection = userAnswers ? `\nInformazioni integrative fornite dall'utente: ${userAnswers}` : '';

    const step4Prompt = `
Agisci come un consulente strategico senior. Devi valutare se le informazioni attuali sono sufficienti o se mancano dettagli critici.
I 3 pilastri fondamentali da considerare sono:
1. OBIETTIVO SPECIFICO: "${objective}"
2. SETTORE DI RIFERIMENTO: "${sector}"
3. SINTESI CHIAVE DEL DOCUMENTO:
${synthesisText}
${userNotesSection}

REGOLA CRITICA PER LA DIALOGO UTENTE:
Se ritieni che manchino indicazioni fondamentali o dati strategici per personalizzare al meglio il lavoro, NON inventarli ma restituisci un oggetto JSON con questo formato esatto:
{
  "needsInput": true,
  "questions": [
    "Prima domanda specifica da fare all'utente?",
    "Seconda domanda specifica da fare all'utente?"
  ]
}

Se invece hai tutte le informazioni necessarie per procedere, genera la struttura completa dell'output suddivisa in sezioni/slide professionali con Titoli e Punti Elenco.
    `.trim();

    const { text: aiResponse } = await callAIWithWorkflow(step4Prompt);

    if (aiResponse.trim().startsWith('{') && aiResponse.includes('needsInput')) {
      try {
        const jsonMatch = aiResponse.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          const parsed = JSON.parse(jsonMatch[0]);
          if (parsed.needsInput) {
            return NextResponse.json({
              needsInput: true,
              questions: parsed.questions || ['Potresti fornire maggiori dettagli sul target di riferimento?'],
            }, { status: 200 });
          }
        }
      } catch (e) {
        // Ignora e prosegue
      }
    }

    // ====================================================
    // STEP 5: DOWNLOAD AUTOMATICO (PPTX o DOCX)
    // ====================================================
    if (outputFormat === 'docx' || objective.toLowerCase().includes('docx') || objective.toLowerCase().includes('pitch')) {
      const docBuffer = Buffer.from(`REPORT DI SINTESI / PITCH\nMotore IA: \({provider}\nSettore:\){sector}\nObiettivo: \({objective}\n\n\){aiResponse}`, 'utf-8');
      return new NextResponse(docBuffer, {
        headers: {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          'Content-Disposition': `attachment; filename="${file.name.split('.')[0] || 'documento'}-sintesi.docx"`,
        },
      });
    } else {
      const pptx = new pptxgen();
      
      const coverSlide = pptx.addSlide();
      coverSlide.addText(`Presentazione: ${file.name}`, { x: 1, y: 1.5, fontSize: 24, bold: true, color: '363636' });
      coverSlide.addText(`Settore: \({sector} | Obiettivo:\){objective}\nElaborato con: ${provider}`, { x: 1, y: 2.5, fontSize: 13, color: '666666' });

      const contentSlide = pptx.addSlide();
      contentSlide.addText("Contenuti e Punti Chiave", { x: 1, y: 0.8, fontSize: 20, bold: true, color: '363636' });
      contentSlide.addText(aiResponse.substring(0, 1500), { x: 1, y: 1.5, fontSize: 12, color: '444444', w: '85%', h: '70%' });

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
