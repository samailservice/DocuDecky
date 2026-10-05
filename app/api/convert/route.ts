import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import pptxgen from 'pptxgenjs';
import fs from 'fs/promises';
import path from 'path';
import os from 'os';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Funzione di retry con gestione dedicata dell'errore di quota (429)
async function generateWithRetry(contents: any, maxRetries = 3) {
  const model = 'gemini-3.8-flash';

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = await ai.models.generateContent({
        model: model,
        contents: contents,
      });
      return response.text || '';
    } catch (error: any) {
      console.warn(`Modello \({model} - Tentativo\){attempt}/${maxRetries} fallito:`, error.message);
      
      // Se la quota giornaliera gratuita è esaurita (429), interrompi subito i retry e segnalalo
      if (error?.status === 429 || error?.message?.includes('RESOURCE_EXHAUSTED') || error?.message?.includes('quota')) {
        throw new Error('QUOTA_EXHAUSTED');
      }

      if (attempt === maxRetries) throw error;
      await new Promise((resolve) => setTimeout(resolve, attempt * 2000));
    }
  }
  throw new Error('Superato il limite massimo di tentativi con Gemini.');
}

export async function POST(req: Request) {
  let tempFilePath: string | null = null;

  try {
    const formData = await req.formData();
    const file = formData.get('file') as File | null;
    const sector = (formData.get('sector') as string) || 'Business / Aziendale';
    const objective = (formData.get('objective') as string) || 'Presentazione PPTX';
    const userAnswers = (formData.get('userAnswers') as string) || '';

    if (!file) {
      return NextResponse.json({ message: 'Nessun file caricato' }, { status: 400 });
    }

    const fileBuffer = await file.arrayBuffer();
    const fileBytes = Buffer.from(fileBuffer);

    // =========================================================================
    // STEP 1: Prima query - Sottopone il file e genera il prompt di sintesi
    // =========================================================================
    const prompt1 = [
      {
        inlineData: {
          mimeType: file.type || 'application/pdf',
          data: fileBytes.toString('base64'),
        },
      },
      `Sei un analista di sistemi e prompt engineer esperto. Analizza il documento allegato e scrivi un prompt funzionale, dettagliato e strutturato (inclusi dati finanziari, metriche e punti chiave) ottimizzato per guidare la generazione di una presentazione di sintesi professionale.
      
      Contesto utente:
      - Obiettivo: ${objective}
      - Settore: ${sector}
      - Note: ${userAnswers || 'Nessuna'}

      Restituisci esclusivamente il testo del prompt e dei dati strutturati, senza preamboli o saluti.`
    ];

    const generatedPromptText = await generateWithRetry(prompt1);

    // =========================================================================
    // STEP 2: Salvataggio dell'output in file temporaneo
    // =========================================================================
    const tempDir = os.tmpdir();
    tempFilePath = path.join(tempDir, `docudecky-prompt-${Date.now()}.txt`);
    await fs.writeFile(tempFilePath, generatedPromptText, 'utf-8');

    const fileContentForQuery2 = await fs.readFile(tempFilePath, 'utf-8');

    // =========================================================================
    // STEP 3: Seconda query - Crea la presentazione basata sul file temporaneo
    // =========================================================================
    const prompt2 = `
    Agisci come un esperto di corporate storytelling. Utilizzando rigorosamente il testo e le istruzioni presenti nel file di riferimento qui sotto, genera la struttura finale della presentazione suddivisa in esattamente 6-7 sezioni/slide.
    
    Usa tassativamente questo formato esatto per ogni slide:
    === SLIDE N: [Titolo della Slide] ===
    - [Punto elenco 1 con dettagli e metriche]
    - [Punto elenco 2]
    - [Punto elenco 3]

    Testo di riferimento (dal file temporaneo):
    ${fileContentForQuery2}
    `;

    const finalSlidesText = await generateWithRetry(prompt2);

    // =========================================================================
    // STEP 4: Generazione file PowerPoint (.pptx)
    // =========================================================================
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';

    const slideChunks = finalSlidesText.split('=== SLIDE').filter(Boolean);

    if (slideChunks.length === 0) {
      const slide = pptx.addSlide();
      slide.addText(finalSlidesText, { x: 0.8, y: 0.8, w: '85%', h: '80%', fontSize: 14, color: '333333' });
    } else {
      for (const chunk of slideChunks) {
        const lines = chunk.split('\n').map((l) => l.trim()).filter(Boolean);
        const headerLine = lines[0].replace(/^[:=]+/, '').trim();
        const bulletPoints = lines.slice(1).map((l) => l.replace(/^[-*]\s*/, ''));

        const slide = pptx.addSlide();

        slide.addText(headerLine, {
          x: 0.8,
          y: 0.6,
          w: '85%',
          h: 0.8,
          fontSize: 22,
          bold: true,
          color: '4A154B',
        });

        if (bulletPoints.length > 0) {
          const formattedBullets = bulletPoints.map((bp) => ({
            text: bp,
            options: { bullet: true, fontSize: 14, color: '333333', spaceAfter: 10 },
          }));

          slide.addText(formattedBullets, {
            x: 0.8,
            y: 1.6,
            w: '85%',
            h: '70%',
          });
        }
      }
    }

    const pptxBuffer = await pptx.write({ outputType: 'nodebuffer' });

    // =========================================================================
    // STEP 5: Cancellazione file temporaneo
    // =========================================================================
    if (tempFilePath) {
      try {
        await fs.unlink(tempFilePath);
        tempFilePath = null;
      } catch (e) {}
    }

    return new NextResponse(pptxBuffer as unknown as BodyInit, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': 'attachment; filename="presentazione-docudecky.pptx"',
      },
    });

  } catch (error: any) {
    if (tempFilePath) {
      try {
        await fs.unlink(tempFilePath);
      } catch (e) {}
    }

    console.error('Errore API Workflow Due Step:', error);

    // Gestione specifica per quota esaurita
    if (error.message === 'QUOTA_EXHAUSTED' || error?.status === 429) {
      return NextResponse.json(
        { message: 'Hai raggiunto il limite massimo giornaliero di richieste gratuite della chiave API Gemini (20 richieste/giorno). Riprova domani o inserisci una chiave con un piano a pagamento.' },
        { status: 429 }
      );
    }

    return NextResponse.json(
      { message: error.message || 'Errore interno durante il flusso di elaborazione a due step.' },
      { status: 500 }
    );
  }
}
