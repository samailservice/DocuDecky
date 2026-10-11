import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import pptxgen from 'pptxgenjs';
import fs from 'fs/promises';
import path from 'path';
import os from 'os';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

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
      console.warn(`Modello ${model} - Tentativo ${attempt}/${maxRetries} fallito:`, error.message);
      
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
  let tempFile1: string | null = null;
  let tempFile2: string | null = null;
  let tempFile3: string | null = null;

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
    const tempDir = os.tmpdir();

    // =========================================================================
    // STEP 1: Estrazione dati e sintesi dal documento
    // =========================================================================
    const prompt1 = [
      {
        inlineData: {
          mimeType: file.type || 'application/pdf',
          data: fileBytes.toString('base64'),
        },
      },
      `Sei un analista aziendale esperto. Analizza il documento allegato estraendo tutti i dati chiave, metriche, bilanci e strategie.
      Contesto:
      - Obiettivo: ${objective}
      - Settore: ${sector}
      - Note: ${userAnswers || 'Nessuna'}
      Restituisci esclusivamente la sintesi analitica dettagliata, senza preamboli.`
    ];

    const rawAnalysisText = await generateWithRetry(prompt1);

    tempFile1 = path.join(tempDir, `docudecky-step1-${Date.now()}.txt`);
    await fs.writeFile(tempFile1, rawAnalysisText, 'utf-8');
    const contentStep1 = await fs.readFile(tempFile1, 'utf-8');

    // =========================================================================
    // STEP 2: Creazione della struttura iniziale delle slide
    // =========================================================================
    const prompt2 = `
    Basandoti esclusivamente sui dati di sintesi contenuti nel testo sottostante, suddividi i contenuti in esattamente 6-7 sezioni/slide.
    Usa rigorosamente questo formato per ogni slide:
    === SLIDE N: [Titolo della Slide] ===
    - [Punto elenco 1 con dettagli]
    - [Punto elenco 2]

    Dati di sintesi:
    ${contentStep1}
    `;

    const rawStructureText = await generateWithRetry(prompt2);

    tempFile2 = path.join(tempDir, `docudecky-step2-${Date.now()}.txt`);
    await fs.writeFile(tempFile2, rawStructureText, 'utf-8');
    const contentStep2 = await fs.readFile(tempFile2, 'utf-8');

    // =========================================================================
    // STEP 3: Raffinamento con vincolo tassativo per layout, grafici e commenti
    // =========================================================================
    const prompt3 = `
    Sei un esperto di data visualization e corporate storytelling. 
    Modifica la presentazione applicando un layout comune e professionale a tutte le slide. 
    È TASSATIVO E OBBLIGATORIO inserire i tag per i grafici [CHART: ...] in almeno 2 o 3 slide che trattano dati finanziari, ricavi, costi o ripartizioni percentuali, usando dati numerici reali presenti nel testo.

    Per ogni slide, segui rigorosamente questa struttura esatta:
    === SLIDE N: [Titolo della Slide] ===
    - [Punto elenco 1]
    - [Punto elenco 2]
    [CHART: bar | Titolo del Grafico | Categoria1,Categoria2,Categoria3 | 100,200,300] (Usa 'bar' o 'pie' come tipo. Includi questa riga se la slide contiene dati numerici idonei).
    [COMMENTO: Nota dettagliata o script per il relatore].

    Struttura di partenza da ottimizzare:
    ${contentStep2}
    `;

    const finalOptimizedText = await generateWithRetry(prompt3);

    tempFile3 = path.join(tempDir, `docudecky-step3-${Date.now()}.txt`);
    await fs.writeFile(tempFile3, finalOptimizedText, 'utf-8');
    const finalContent = await fs.readFile(tempFile3, 'utf-8');

    // =========================================================================
    // STEP 4: Generazione del file PowerPoint (.pptx) con parsing dei grafici
    // =========================================================================
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';

    const slideChunks = finalContent.split('=== SLIDE').filter(Boolean);

    if (slideChunks.length === 0) {
      const slide = pptx.addSlide();
      slide.addText(finalContent, { x: 0.8, y: 0.8, w: '85%', h: '80%', fontSize: 14, color: '333333' });
    } else {
      for (const chunk of slideChunks) {
        const lines = chunk.split('\n').map((l) => l.trim()).filter(Boolean);
        const headerLine = lines[0].replace(/^[:=]+/, '').trim();
        
        const bulletPoints: string[] = [];
        let chartConfig: any = null;
        let slideComment = '';

        for (const line of lines.slice(1)) {
          if (line.startsWith('[CHART:')) {
            const match = line.match(/\[CHART:\s*([^\vert{}]+)\Vert{}\s*([^\vert{}]+)\Vert{}\s*([^\vert{}]+)\Vert{}\s*([^\]]+)\]/i);
            if (match) {
              const chartTypeInput = match[1].trim().toLowerCase();
              const chartType = chartTypeInput.includes('pie') ? pptx.ChartType.pie : pptx.ChartType.bar;
              const chartTitle = match[2].trim();
              const categories = match[3].split(',').map((c) => c.trim());
              const values = match[4].split(',').map((v) => parseFloat(v.trim().replace('€', '').replace('M', '')) || 0);

              chartConfig = {
                type: chartType,
                data: [
                  {
                    name: chartTitle,
                    labels: categories,
                    values: values,
                  },
                ],
                options: {
                  title: chartTitle,
                  showLegend: true,
                  legendPos: 'b',
                },
              };
            }
          } else if (line.startsWith('[COMMENTO:')) {
            slideComment = line.replace('[COMMENTO:', '').replace(']', '').trim();
          } else if (line.startsWith('-') || line.startsWith('*')) {
            bulletPoints.push(line.replace(/^[-*]\s*/, ''));
          }
        }

        const slide = pptx.addSlide();

        // Layout comune del titolo
        slide.addText(headerLine, {
          x: 0.8,
          y: 0.6,
          w: '85%',
          h: 0.8,
          fontSize: 22,
          bold: true,
          color: '4A154B',
        });

        // Se l'IA ha generato il grafico, dividiamo la slide: testo a sinistra, grafico a destra
        if (chartConfig && bulletPoints.length > 0) {
          const formattedBullets = bulletPoints.map((bp) => ({
            text: bp,
            options: { bullet: true, fontSize: 13, color: '333333', spaceAfter: 8 },
          }));

          slide.addText(formattedBullets, {
            x: 0.8,
            y: 1.6,
            w: 5.0,
            h: 4.5,
          });

          slide.addChart(chartConfig.type, chartConfig.data, {
            x: 6.2,
            y: 1.6,
            w: 6.0,
            h: 4.5,
            ...chartConfig.options,
          });
        } else {
          // Altrimenti visualizziamo i punti elenco a tutta larghezza
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

        // Inserimento note/commenti del relatore
        if (slideComment) {
          slide.addNotes(slideComment);
        }
      }
    }

    const pptxBuffer = await pptx.write({ outputType: 'nodebuffer' });

    // =========================================================================
    // STEP 5: Pulizia file temporanei
    // =========================================================================
    for (const filePath of [tempFile1, tempFile2, tempFile3]) {
      if (filePath) {
        try { await fs.unlink(filePath); } catch (e) {}
      }
    }

    return new NextResponse(pptxBuffer as unknown as BodyInit, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': 'attachment; filename="presentazione-docudecky.pptx"',
      },
    });

  } catch (error: any) {
    for (const filePath of [tempFile1, tempFile2, tempFile3]) {
      if (filePath) {
        try { await fs.unlink(filePath); } catch (e) {}
      }
    }

    console.error('Errore API Workflow Tre Step:', error);

    if (error.message === 'QUOTA_EXHAUSTED' || error?.status === 429) {
      return NextResponse.json(
        { message: 'Hai raggiunto il limite massimo giornaliero di richieste gratuite della chiave API Gemini. Riprova domani o usa una chiave a pagamento.' },
        { status: 429 }
      );
    }

    return NextResponse.json(
      { message: error.message || 'Errore interno durante il flusso a tre step.' },
      { status: 500 }
    );
  }
}
