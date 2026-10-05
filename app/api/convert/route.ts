import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import pptxgen from 'pptxgenjs';
import fs from 'fs/promises';
import path from 'path';
import os from 'os';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Funzione di retry avanzata con fallback di modello in caso di errore 503 (High Demand)
async function generateWithRetry(contents: any, maxRetries = 4) {
  const models = ['gemini-3.8-flash', 'gemini-1.5-flash'];

  for (const model of models) {
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        const response = await ai.models.generateContent({
          model: model,
          contents: contents,
        });
        return response.text || '';
      } catch (error: any) {
        console.warn(`Modello \({model} - Tentativo\){attempt}/${maxRetries} fallito:`, error.message);
        if (attempt === maxRetries && model === models[models.length - 1]) {
          throw error;
        }
        await new Promise((resolve) => setTimeout(resolve, attempt * 2500));
      }
    }
  }
  throw new Error('Superato il limite massimo di tentativi con i modelli Gemini.');
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
    // STEP 1: Prima query - Sottopone il file e chiede di creare un prompt funzionale
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
    // STEP 2: Salvataggio dell'output della prima query in un file temporaneo
    // =========================================================================
    const tempDir = os.tmpdir();
    tempFilePath = path.join(tempDir, `docudecky-prompt-${Date.now()}.txt`);
    await fs.writeFile(tempFilePath, generatedPromptText, 'utf-8');

    // Lettura del testo dal file temporaneo appena creato
    const fileContentForQuery2 = await fs.readFile(tempFilePath, 'utf-8');

    // =========================================================================
    // STEP 3: Seconda query - Usa il testo del file temporaneo per creare la presentazione
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
    // STEP 4: Generazione del file PowerPoint (.pptx)
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
    // STEP 5: Cancellazione del file temporaneo (avvenuta generazione con successo)
    // =========================================================================
    if (tempFilePath) {
      try {
        await fs.unlink(tempFilePath);
        tempFilePath = null;
      } catch (e) {
        console.warn('Impossibile rimuovere il file temporaneo:', e);
      }
    }

    return new NextResponse(pptxBuffer as unknown as BodyInit, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': 'attachment; filename="presentazione-docudecky.pptx"',
      },
    });

  } catch (error: any) {
    // Pulizia di sicurezza del file temporaneo in caso di errore
    if (tempFilePath) {
      try {
        await fs.unlink(tempFilePath);
      } catch (e) {}
    }

    console.error('Errore API Workflow Due Step:', error);
    return NextResponse.json(
      { message: 'I server di Google sono temporaneamente sovraccarichi (503). Riprova tra qualche istante.' },
      { status: 503 }
    );
  }
}
