import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import pptxgen from 'pptxgenjs';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

async function generateWithRetry(contents: any, maxRetries = 3) {
  const modelsToTry = ['gemini-2.5-flash', 'gemini-1.5-flash', 'gemini-1.5-pro'];

  for (const model of modelsToTry) {
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

        if (attempt === maxRetries) {
          break;
        }

        await new Promise((resolve) => setTimeout(resolve, attempt * 2500));
      }
    }
  }

  throw new Error('I server di Google sono temporaneamente sovraccarichi su tutti i modelli disponibili. Riprova tra qualche istante.');
}

export async function POST(req: Request) {
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

    const prompt = [
      {
        inlineData: {
          mimeType: file.type || 'application/pdf',
          data: fileBytes.toString('base64'),
        },
      },
      `Sei un analista aziendale e un esperto di corporate storytelling. Analizza il documento allegato estraendo i dati chiave e genera una presentazione professionale compatibile con Google Slide composta da esattamente 6 o 7 slide.
      
      Contesto utente:
      - Obiettivo: ${objective}
      - Settore: ${sector}
      - Note: ${userAnswers || 'Nessuna'}

      Devi restituire **ESCLUSIVAMENTE** un oggetto JSON valido (senza blocchi markdown tipo \`\`\`json) strutturato esattamente così:

      {
        "slides": [
          {
            "title": "Titolo della Slide",
            "subtitle": "Sottotitolo chiaro ed esplicativo",
            "cards": [
              {
                "title": "Titolo del Blocco / Metrica",
                "content": "Testo descrittivo dettagliato o valore numerico"
              }
            ],
            "chart": {
              "type": "bar",
              "title": "Titolo del Grafico",
              "categories": ["Cat 1", "Cat 2", "Cat 3"],
              "values": [10, 20, 30]
            },
            "comment": "Nota dettagliata per il relatore."
          }
        ]
      }

      Regole fondamentali:
      1. Struttura i contenuti di ogni slide in 2 o 3 "cards" (oggetti con title e content).
      2. Includi l'oggetto "chart" (con tipo "bar" o "pie", categorie e valori numerici reali) nelle slide che presentano dati finanziari, costi o ricavi. Nelle altre slide ometti completamente la chiave "chart".
      3. Restituisci SOLO il JSON puro.`
    ];

    const jsonResponseText = await generateWithRetry(prompt);

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

    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';

    const slidesArray = presentationData.slides || [];

    if (slidesArray.length === 0) {
      const slide = pptx.addSlide();
      slide.addText("Nessun contenuto generato.", { x: 0.8, y: 0.8, w: '85%', h: '80%', fontSize: 14, color: '333333' });
    } else {
      for (const slideData of slidesArray) {
        const slide = pptx.addSlide();

        slide.background = { color: 'F8FAFC' };

        slide.addText(slideData.title || 'Slide Professionale', {
          x: 0.8,
          y: 0.5,
          w: '85%',
          h: 0.5,
          fontSize: 22,
          bold: true,
          color: '0F172A',
        });

        if (slideData.subtitle) {
          slide.addText(slideData.subtitle, {
            x: 0.8,
            y: 1.0,
            w: '85%',
            h: 0.4,
            fontSize: 13,
            color: '475569',
            italic: true,
          });
        }

        const cards: any[] = slideData.cards || [];
        const hasChart = slideData.chart && slideData.chart.categories && slideData.chart.values;

        if (hasChart) {
          let cardY = 1.5;
          for (const card of cards.slice(0, 3)) {
            slide.addShape(pptx.ShapeType.roundRect, {
              x: 0.8,
              y: cardY,
              w: 5.2,
              h: 1.2,
              fill: { color: 'FFFFFF' },
              line: { color: 'E2E8F0', width: 1 },
            });

            slide.addShape(pptx.ShapeType.rect, {
              x: 0.8,
              y: cardY,
              w: 0.12,
              h: 1.2,
              fill: { color: '4F46E5' },
            });

            slide.addText([
              { text: (card.title || '') + '\n', options: { bold: true, fontSize: 12, color: '1E293B' } },
              { text: card.content || '', options: { fontSize: 10, color: '475569', lineSpacing: 12 } }
            ], {
              x: 1.1,
              y: cardY + 0.1,
              w: 4.7,
              h: 1.0,
            });

            cardY += 1.35;
          }

          const chartType = String(slideData.chart.type).toLowerCase() === 'pie' ? pptx.ChartType.pie : pptx.ChartType.bar;
          const chartValues = slideData.chart.values.map((v: any) => parseFloat(v) || 0);

          slide.addChart(
            chartType,
            [
              {
                name: slideData.chart.title || 'Valori',
                labels: slideData.chart.categories,
                values: chartValues,
              },
            ],
            {
              x: 6.4,
              y: 1.5,
              w: 5.8,
              h: 4.8,
              title: slideData.chart.title || '',
              showLegend: true,
              legendPos: 'b',
              chartColors: ['4F46E5', '06B6D4', '10B981', 'F59E0B'],
            }
          );

        } else {
          let startX = 0.8;
          let startY = 1.5;
          
          cards.forEach((card, index) => {
            const col = index % 2;
            const row = Math.floor(index / 2);
            
            const posX = startX + col * 5.8;
            const posY = startY + row * 1.85;

            slide.addShape(pptx.ShapeType.roundRect, {
              x: posX,
              y: posY,
              w: 5.5,
              h: 1.65,
              fill: { color: 'FFFFFF' },
              line: { color: 'E2E8F0', width: 1 },
            });

            slide.addShape(pptx.ShapeType.rect, {
              x: posX,
              y: posY,
              w: 0.12,
              h: 1.65,
              fill: { color: '4F46E5' },
            });

            slide.addText([
              { text: (card.title || '') + '\n', options: { bold: true, fontSize: 13, color: '1E293B' } },
              { text: card.content || '', options: { fontSize: 11, color: '475569', lineSpacing: 14 } }
            ], {
              x: posX + 0.3,
              y: posY + 0.15,
              w: 5.0,
              h: 1.35,
            });
          });
        }

        if (slideData.comment) {
          slide.addNotes(slideData.comment);
        }
      }
    }

    const pptxBuffer = await pptx.write({ outputType: 'nodebuffer' });

    return new NextResponse(pptxBuffer as unknown as BodyInit, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': 'attachment; filename="presentazione-docudecky-pro.pptx"',
      },
    });

  } catch (error: any) {
    console.error('Errore API Workflow Ottimizzato:', error);

    if (error.message === 'QUOTA_EXHAUSTED' || error?.status === 429) {
      return NextResponse.json(
        { message: 'Hai raggiunto il limite massimo giornaliero di richieste gratuite della chiave API Gemini (20/giorno). Riprova domani o usa una chiave a pagamento.' },
        { status: 429 }
      );
    }

    return NextResponse.json(
      { message: error.message || 'Errore interno durante la generazione della presentazione.' },
      { status: 500 }
    );
  }
}
