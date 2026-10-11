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
    // STEP 1: Estrazione dati e sintesi avanzata
    // =========================================================================
    const prompt1 = [
      {
        inlineData: {
          mimeType: file.type || 'application/pdf',
          data: fileBytes.toString('base64'),
        },
      },
      `Sei un analista aziendale esperto. Analizza il documento allegato estraendo tutti i dati chiave, metriche, bilanci e strategie (es. ricavi, costi operativi, stato patrimoniale, KPI).
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
    // STEP 2: Generazione JSON Strutturato per Layout Visivo e Grafici
    // =========================================================================
    const promptJson = `
    Sei un esperto di corporate storytelling e data visualization. 
    Analizza i dati di sintesi forniti e genera una presentazione professionale compatibile con Google Slide composta da esattamente 6 o 7 slide.
    
    Devi restituire **ESCLUSIVAMENTE** un oggetto JSON valido (senza blocchi markdown tipo \`\`\`json) strutturato esattamente così:

    {
      "slides": [
        {
          "title": "Titolo della Slide",
          "subtitle": "Breve sottotitolo descrittivo",
          "cards": [
            {
              "title": "Titolo Blocco / Metrica",
              "content": "Testo esplicativo o dettaglio numerico"
            }
          ],
          "chart": {
            "type": "bar",
            "title": "Titolo del Grafico",
            "categories": ["Cat 1", "Cat 2", "Cat 3"],
            "values": [10, 20, 30]
          },
          "comment": "Nota dettagliata per lo speaker."
        }
      ]
    }

    Regole fondamentali:
    1. Invece di semplici elenchi, organizza i contenuti in 2 o 3 "cards" (oggetti con title e content) per dare un impatto visivo a blocchi.
    2. Includi l'oggetto "chart" (con tipo "bar" o "pie", categorie e valori numerici reali estratti dal testo) nelle slide che presentano dati finanziari, ripartizioni di costi o ricavi.
    3. Restituisci SOLO il JSON puro.

    Dati di sintesi di partenza:
    ${contentStep1}
    `;

    const jsonResponseText = await generateWithRetry(promptJson);

    tempFile2 = path.join(tempDir, `docudecky-step2-${Date.now()}.txt`);
    await fs.writeFile(tempFile2, jsonResponseText, 'utf-8');

    let cleanJsonStr = jsonResponseText.trim();
    if (cleanJsonStr.startsWith('```json')) {
      cleanJsonStr = cleanJsonStr.replace(/^```json/, '').replace(/```$/, '').trim();
    } else if (cleanJsonStr.startsWith('```')) {
      cleanJsonStr = cleanJsonStr.replace(/^```/, '').replace(/```$/, '').trim();
    }

    let presentationData: any;
    try {
      presentationData = JSON.parse(cleanJsonStr);
    } catch (e) {
      console.error('Errore parsing JSON dall\'IA', e);
      throw new Error('L\'intelligenza artificiale non ha restituito un formato JSON valido. Riprova.');
    }

    // =========================================================================
    // STEP 3: Generazione PowerPoint con Card Grafiche e Box Dinamici
    // =========================================================================
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';

    const slidesArray = presentationData.slides || [];

    if (slidesArray.length === 0) {
      const slide = pptx.addSlide();
      slide.addText("Nessun contenuto generato.", { x: 0.8, y: 0.8, w: '85%', h: '80%', fontSize: 14, color: '333333' });
    } else {
      for (const slideData of slidesArray) {
        const slide = pptx.addSlide();

        // 1. Header Unificato di Classe Enterprise
        slide.addText(slideData.title || 'Slide Professionale', {
          x: 0.8,
          y: 0.5,
          w: '85%',
          h: 0.6,
          fontSize: 22,
          bold: true,
          color: '4A154B', // Branding color
        });

        if (slideData.subtitle) {
          slide.addText(slideData.subtitle, {
            x: 0.8,
            y: 1.1,
            w: '85%',
            h: 0.4,
            fontSize: 13,
            color: '666666',
            italic: true,
          });
        }

        const cards: any[] = slideData.cards || [];
        const hasChart = slideData.chart && slideData.chart.categories && slideData.chart.values;

        // 2. Layout Visivo a Card / Riquadri Colorati + Grafico opzionale
        if (hasChart) {
          // Se c'è un grafico: mettiamo le card a sinistra e il grafico a destra
          let cardY = 1.6;
          for (const card of cards.slice(0, 3)) {
            // Disegna il riquadro di sfondo della card
            slide.addShape(pptx.ShapeType.roundRect, {
              x: 0.8,
              y: cardY,
              w: 5.2,
              h: 1.1,
              fill: { color: 'F4F1F5' },
              line: { color: 'E2D9E4', width: 1 },
            });
            // Testo interno alla card
            slide.addText([
              { text: (card.title || '') + '\n', options: { bold: true, fontSize: 12, color: '4A154B' } },
              { text: card.content || '', options: { fontSize: 10, color: '333333' } }
            ], {
              x: 1.0,
              y: cardY + 0.1,
              w: 4.8,
              h: 0.9,
            });
            cardY += 1.25;
          }

          // Inserimento grafico nativo a destra
          const chartType = String(slideData.chart.type).toLowerCase() === 'pie' ? pptx.ChartType.pie : pptx.ChartType.bar;
          const chartValues = slideData.chart.values.map((v: any) => parseFloat(v) || 0);

          slide.addChart(
            chartType,
            [
              {
                name: slideData.chart.title || 'Dati',
                labels: slideData.chart.categories,
                values: chartValues,
              },
            ],
            {
              x: 6.4,
              y: 1.6,
              w: 5.8,
              h: 4.8,
              title: slideData.chart.title || '',
              showLegend: true,
              legendPos: 'b',
            }
          );

        } else {
          // Se non c'è grafico, distribuiamo le card in una griglia visiva pulita
          let startX = 0.8;
          let startY = 1.6;
          
          cards.forEach((card, index) => {
            const col = index % 2;
            const row = Math.floor(index / 2);
            
            const posX = startX + col * 5.8;
            const posY = startY + row * 1.8;

            // Box Card di sfondo
            slide.addShape(pptx.ShapeType.roundRect, {
              x: posX,
              y: posY,
              w: 5.5,
              h: 1.6,
              fill: { color: 'F9F8FA' },
              line: { color: 'E2D9E4', width: 1 },
            });

            slide.addText([
              { text: (card.title || '') + '\n', options: { bold: true, fontSize: 13, color: '4A154B' } },
              { text: card.content || '', options: { fontSize: 11, color: '333333' } }
            ], {
              x: posX + 0.2,
              y: posY + 0.15,
              w: 5.1,
              h: 1.3,
            });
          });
        }

        // 3. Aggiunta Note per il Relatore
        if (slideData.comment) {
          slide.addNotes(slideData.comment);
        }
      }
    }

    const pptxBuffer = await pptx.write({ outputType: 'nodebuffer' });

    for (const filePath of [tempFile1, tempFile2]) {
      if (filePath) {
        try { await fs.unlink(filePath); } catch (e) {}
      }
    }

    return new NextResponse(pptxBuffer as unknown as BodyInit, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': 'attachment; filename="presentazione-docudecky-visiva.pptx"',
      },
    });

  } catch (error: any) {
    for (const filePath of [tempFile1, tempFile2]) {
      if (filePath) {
        try { await fs.unlink(filePath); } catch (e) {}
      }
    }

    console.error('Errore API Workflow Visivo:', error);

    if (error.message === 'QUOTA_EXHAUSTED' || error?.status === 429) {
      return NextResponse.json(
        { message: 'Hai raggiunto il limite massimo giornaliero di richieste gratuite della chiave API Gemini. Riprova domani o usa una chiave a pagamento.' },
        { status: 429 }
      );
    }

    return NextResponse.json(
      { message: error.message || 'Errore interno durante la generazione della presentazione.' },
      { status: 500 }
    );
  }
}
