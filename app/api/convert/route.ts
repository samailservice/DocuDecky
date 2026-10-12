import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import pptxgen from 'pptxgenjs';

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
      console.warn(`Tentativo ${attempt}/${maxRetries} fallito:`, error.message);
      
      if (error?.status === 429 || error?.message?.includes('RESOURCE_EXHAUSTED') || error?.message?.includes('quota')) {
        throw new Error('QUOTA_EXHAUSTED');
      }

      if (attempt === maxRetries) throw error;
      await new Promise((resolve) => setTimeout(resolve, 1500));
    }
  }
  throw new Error('Superato il limite massimo di tentativi.');
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
      `Sei un analista aziendale senior e un esperto di corporate storytelling. Analizza il documento allegato estraendo i dati chiave e genera una presentazione professionale in stile consulenza strategica composta da esattamente 6 o 7 slide.
      
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
    const totalSlides = slidesArray.length;

    if (totalSlides === 0) {
      const slide = pptx.addSlide();
      slide.addText("Nessun contenuto generato.", { x: 0.8, y: 0.8, w: '85%', h: '80%', fontSize: 14, color: '333333' });
    } else {
      slidesArray.forEach((slideData: any, idx: number) => {
        const slide = pptx.addSlide();

        // 1. Sfondo professionale ardesia chiarissimo
        slide.background = { color: 'F8FAFC' };

        // 2. Barra superiore decorativa corporate (Brand Accent)
        slide.addShape(pptx.ShapeType.rect, {
          x: 0,
          y: 0,
          w: '100%',
          h: 0.1,
          fill: { color: '2563EB' } // Blu Corporate
        });

        // 3. Titolo principale della slide
        slide.addText(slideData.title || 'Slide Professionale', {
          x: 0.8,
          y: 0.45,
          w: '85%',
          h: 0.5,
          fontSize: 22,
          bold: true,
          color: '0F172A',
        });

        // 4. Sottotitolo descrittivo
        if (slideData.subtitle) {
          slide.addText(slideData.subtitle, {
            x: 0.8,
            y: 0.95,
            w: '85%',
            h: 0.4,
            fontSize: 13,
            color: '64748B',
            italic: true,
          });
        }

        // 5. Piè di pagina elegante con numerazione slide
        slide.addText(`DocuDecky Pro  |  Slide ${idx + 1} di ${totalSlides}`, {
          x: 0.8,
          y: 6.9,
          w: '85%',
          h: 0.3,
          fontSize: 9,
          color: '94A3B8',
        });

        const cards: any[] = slideData.cards || [];
        const hasChart = slideData.chart && slideData.chart.categories && slideData.chart.values;

        // 6. Layout con Grafico Affiancato (Split Screen)
        if (hasChart) {
          let cardY = 1.5;
          const maxCardsToShow = Math.min(cards.length, 3);
          
          for (let i = 0; i < maxCardsToShow; i++) {
            const card = cards[i];

            // Card di sfondo bianca con bordo sottile
            slide.addShape(pptx.ShapeType.roundRect, {
              x: 0.8,
              y: cardY,
              w: 5.2,
              h: 1.55,
              fill: { color: 'FFFFFF' },
              line: { color: 'CBD5E1', width: 1 },
              rectRadius: 0.08
            });

            // Barra d'accento laterale della card
            slide.addShape(pptx.ShapeType.rect, {
              x: 0.8,
              y: cardY,
              w: 0.12,
              h: 1.55,
              fill: { color: i % 2 === 0 ? '2563EB' : '4F46E5' },
            });

            // Testo della card
            slide.addText([
              { text: (card.title || '') + '\n', options: { bold: true, fontSize: 12, color: '0F172A' } },
              { text: card.content || '', options: { fontSize: 10.5, color: '334155', lineSpacing: 14 } }
            ], {
              x: 1.15,
              y: cardY + 0.15,
              w: 4.7,
              h: 1.25,
            });

            cardY += 1.75;
          }

          // Inserimento del grafico a destra
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
              x: 6.35,
              y: 1.5,
              w: 6.35,
              h: 4.9,
              title: slideData.chart.title || '',
              titleColor: '0F172A',
              titleFontSize: 13,
              showLegend: true,
              legendPos: 'b',
              chartColors: ['2563EB', '06B6D4', '10B981', 'F59E0B'],
              valAxisLabelColor: '64748B',
              catAxisLabelColor: '64748B',
            }
          );

        } else {
          // 7. Layout a Griglia per le Card (2x2 o 1x3)
          const startX = 0.8;
          const startY = 1.5;
          
          cards.forEach((card, index) => {
            const col = index % 2;
            const row = Math.floor(index / 2);
            
            const posX = startX + col * 5.9;
            const posY = startY + row * 2.55;

            // Sfondo card
            slide.addShape(pptx.ShapeType.roundRect, {
              x: posX,
              y: posY,
              w: 5.6,
              h: 2.35,
              fill: { color: 'FFFFFF' },
              line: { color: 'CBD5E1', width: 1 },
              rectRadius: 0.08
            });

            // Barra verticale laterale
            slide.addShape(pptx.ShapeType.rect, {
              x: posX,
              y: posY,
              w: 0.12,
              h: 2.35,
              fill: { color: index === 0 ? '2563EB' : index === 1 ? '4F46E5' : '059669' },
            });

            // Contenuto testuale formattato
            slide.addText([
              { text: (card.title || '') + '\n', options: { bold: true, fontSize: 13, color: '0F172A' } },
              { text: card.content || '', options: { fontSize: 11, color: '334155', lineSpacing: 16 } }
            ], {
              x: posX + 0.35,
              y: posY + 0.2,
              w: 5.0,
              h: 1.95,
            });
          });
        }

        // Note per il relatore
        if (slideData.comment) {
          slide.addNotes(slideData.comment);
        }
      });
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
        { message: 'Hai raggiunto il limite massimo giornaliero di richieste della chiave API Gemini. Verifica il tuo piano a consumo.' },
        { status: 429 }
      );
    }

    return NextResponse.json(
      { message: error.message || 'Errore interno durante la generazione della presentazione.' },
      { status: 500 }
    );
  }
}
