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
      `Sei un analista finanziario senior e un esperto di corporate storytelling. Analizza il documento allegato e genera una presentazione di esattamente 7 slide strutturata rigorosamente secondo il modello standard di bilancio e relazione agli stakeholder (stile McKinsey/BCG):
      
      Contesto utente:
      - Obiettivo: ${objective}
      - Settore: ${sector}
      - Note: ${userAnswers || 'Nessuna'}

      Devi restituire **ESCLUSIVAMENTE** un oggetto JSON valido (senza blocchi markdown tipo \`\`\`json) strutturato con esattamente 7 slide in questo ordine preciso:
      - Slide 1: Copertina istituzionale (Titolo, Sottotitolo, Informazioni Societarie)[cite: 4].
      - Slide 2: Executive Summary / Highlights con 4 metriche chiave (Fatturato, EBIT, Utile, Cassa) e commento della governance[cite: 4].
      - Slide 3: Analisi Commerciale & Mix di Prodotto (con oggetto "chart" a torta o barre sui segmenti di ricavo)[cite: 4].
      - Slide 4: Conto Economico & Struttura Costi Operativi (spaccato materie prime, personale, servizi)[cite: 4].
      - Slide 5: Assetto Patrimoniale (Confronto Attivo vs Passivo e Patrimonio Netto)[cite: 4].
      - Slide 6: Indicatori Gestionali e KPI (ROE, ROI, ROS, Debt/Equity)[cite: 4].
      - Slide 7: Prospettive Future, Destinazione Utile e Piano Capex 2026[cite: 4].

      Schema JSON richiesto:
      {
        "slides": [
          {
            "slideNumber": 1,
            "category": "ASSEMBLEA DEI SOCI & STAKEHOLDER",
            "title": "Bilancio d'Esercizio 2025",
            "subtitle": "Risultati Economico-Finanziari e Strategia di Sviluppo",
            "cards": [
              { "title": "INFORMAZIONI SOCIETARIE", "content": "Sede Legale: ... \nP.IVA: ... \nCapitale Sociale: ..." }
            ]
          },
          {
            "slideNumber": 2,
            "category": "SINTESI DIREZIONALE",
            "title": "Executive Summary & Highlights 2025",
            "subtitle": "Crescita a doppia cifra e solidità patrimoniale",
            "cards": [
              { "title": "FATTURATO TOTALE", "content": "€ 2,2M (+12% vs 2024)" },
              { "title": "EBIT OPERATIVO", "content": "€ 350K (15,9% ricavi)" },
              { "title": "UTILE NETTO", "content": "€ 180K (8,2% margine)" },
              { "title": "CASSA & LIQUIDITÀ", "content": "€ 400K disponibili" }
            ],
            "comment": "Commento della Governance: eccellente espansione con redditività in crescita."
          },
          {
            "slideNumber": 3,
            "category": "MERCATO & VENDITE",
            "title": "Analisi Commerciale & Mix di Prodotto",
            "subtitle": "Espansione geografica e posizionamento di nicchia",
            "cards": [
              { "title": "Guanti Tecnici Industriali (60%)", "content": "€ 1.320.000 - Automotive e Siderurgico." },
              { "title": "Guanti in Pelle Lusso (40%)", "content": "€ 880.000 - Manifattura artigianale toscana." }
            ],
            "chart": {
              "type": "pie",
              "title": "Ripartizione Ricavi per Segmento",
              "categories": ["Tecnici Industriali", "Pelle Lusso"],
              "values": [60, 40]
            }
          },
          {
            "slideNumber": 4,
            "category": "PERFORMANCE FINANZIARIA",
            "title": "Conto Economico & Struttura Costi Operativi",
            "subtitle": "Controllo rigoroso dei costi di produzione e marginalità",
            "cards": [
              { "title": "Materie Prime", "content": "€ 900.000 (Pelli, Kevlar, Accessori)" },
              { "title": "Costo Personale", "content": "€ 600.000 (24 FTE operativi)" },
              { "title": "Servizi e Ammortamenti", "content": "€ 350.000 (Utenze e Impianti)" }
            ]
          },
          {
            "slideNumber": 5,
            "category": "ASSETTO PATRIMONIALE",
            "title": "Stato Patrimoniale: Impieghi vs Fonti",
            "subtitle": "Equilibrio finanziario e copertura delle immobilizzazioni",
            "cards": [
              { "title": "ATTIVO (Impieghi)", "content": "Totale € 1.900.000 (Fisso € 650K, Circolante € 1,25M)" },
              { "title": "PASSIVO & NETTO (Fonti)", "content": "Patrimonio Netto € 800K, Debiti € 950K, TFR/Fondi € 150K" }
            ]
          },
          {
            "slideNumber": 6,
            "category": "INDICATORI GESTIONALI",
            "title": "Analisi dei KPI e Solidità Finanziaria",
            "subtitle": "Metriche di redditività e indebitamento controllato",
            "cards": [
              { "title": "ROE: 22,5%", "content": "Eccellente redditività del capitale proprio." },
              { "title": "ROI: 18,4%", "content": "Elevata efficienza operativa caratteristica." },
              { "title": "ROS: 15,9%", "content": "EBIT margin di assoluto rilievo." },
              { "title": "DEBT / EQUITY: 1,18", "content": "Leva finanziaria equilibrata e sostenibile." }
            ]
          },
          {
            "slideNumber": 7,
            "category": "PROSPETTIVE FUTURE",
            "title": "Destinazione Utile & Piano Capex 2026",
            "subtitle": "Autofinanziamento per la robotizzazione e remunerazione soci",
            "cards": [
              { "title": "Riserva Legale (5%)", "content": "€ 9.000 (Accantonamento obbligatorio)" },
              { "title": "Riserva Straordinaria / Capex", "content": "€ 121.000 (Taglio laser robotizzato 2026)" },
              { "title": "Dividendi ai Soci", "content": "€ 50.000 (Payout 27,8%)" }
            ]
          }
        ]
      }

      Regole fondamentali:
      1. Restituisci esattamente 7 slide con la struttura indicata[cite: 4].
      2. Includi categorie descrittive in MAIUSCOLETTO (es. "MERCATO & VENDITE").
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
      slide.addText("Nessun contenuto generato.", { x: 0.8, y: 0.8, w: '85%', h: '80%', fontSize: 14, color: '333333', fontFace: 'Arial' });
    } else {
      slidesArray.forEach((slideData: any, idx: number) => {
        const slide = pptx.addSlide();

        // Sfondo pulito
        slide.background = { color: 'F8FAFC' };

        // Header superiore corporate con categoria e numerazione slide
        slide.addShape(pptx.ShapeType.rect, {
          x: 0,
          y: 0,
          w: '100%',
          h: 0.08,
          fill: { color: '2563EB' }
        });

        slide.addText(`${slideData.category || 'RELAZIONE AZIENDALE'}`, {
          x: 0.8,
          y: 0.35,
          w: '80%',
          h: 0.3,
          fontSize: 10,
          bold: true,
          color: '2563EB',
          fontFace: 'Arial',
        });

        slide.addText(`Slide ${idx + 1} di ${totalSlides}`, {
          x: 10.5,
          y: 0.35,
          w: 2.0,
          h: 0.3,
          fontSize: 10,
          bold: true,
          color: '64748B',
          align: 'right',
          fontFace: 'Arial',
        });

        // Titolo e sottotitolo slide
        slide.addText(slideData.title || 'Titolo Slide', {
          x: 0.8,
          y: 0.65,
          w: '85%',
          h: 0.5,
          fontSize: 20,
          bold: true,
          color: '0F172A',
          fontFace: 'Arial',
        });

        if (slideData.subtitle) {
          slide.addText(slideData.subtitle, {
            x: 0.8,
            y: 1.15,
            w: '85%',
            h: 0.35,
            fontSize: 12,
            color: '64748B',
            italic: true,
            fontFace: 'Arial',
          });
        }

        // Footer di pagina
        slide.addText(`GloveCraft S.r.l.  |  Documento Ufficiale per Assemblea dei Soci`, {
          x: 0.8,
          y: 6.9,
          w: '85%',
          h: 0.3,
          fontSize: 8.5,
          color: '94A3B8',
          fontFace: 'Arial',
        });

        const cards: any[] = slideData.cards || [];
        const hasChart = slideData.chart && slideData.chart.categories && slideData.chart.values;

        if (hasChart) {
          // Layout con grafico a destra e card a sinistra
          let cardY = 1.6;
          cards.slice(0, 3).forEach((card, i) => {
            slide.addShape(pptx.ShapeType.roundRect, {
              x: 0.8,
              y: cardY,
              w: 5.2,
              h: 1.5,
              fill: { color: 'FFFFFF' },
              line: { color: 'CBD5E1', width: 1 },
              rectRadius: 0.06
            });

            slide.addShape(pptx.ShapeType.rect, {
              x: 0.8,
              y: cardY,
              w: 0.1,
              h: 1.5,
              fill: { color: i % 2 === 0 ? '2563EB' : '4F46E5' },
            });

            slide.addText([
              { text: (card.title || '') + '\n', options: { bold: true, fontSize: 11, color: '0F172A', fontFace: 'Arial' } },
              { text: card.content || '', options: { fontSize: 10, color: '334155', lineSpacing: 13, fontFace: 'Arial' } }
            ], {
              x: 1.1,
              y: cardY + 0.12,
              w: 4.7,
              h: 1.25,
            });

            cardY += 1.68;
          });

          const chartType = String(slideData.chart.type).toLowerCase() === 'pie' ? pptx.ChartType.pie : pptx.ChartType.bar;
          const chartValues = slideData.chart.values.map((v: any) => parseFloat(v) || 0);
          const chartCategories = slideData.chart.categories.map((c: any) => String(c));

          slide.addChart(
            chartType,
            [
              {
                name: slideData.chart.title || 'Valori',
                labels: chartCategories,
                values: chartValues,
              },
            ],
            {
              x: 6.3,
              y: 1.6,
              w: 6.3,
              h: 4.8,
              title: slideData.chart.title || '',
              titleColor: '0F172A',
              titleFontSize: 12,
              titleFontFace: 'Arial',
              showLegend: true,
              legendPos: 'b',
              chartColors: ['2563EB', '06B6D4', '10B981', 'F59E0B'],
              valAxisLabelColor: '64748B',
              catAxisLabelColor: '64748B',
            }
          );

        } else {
          // Layout a griglia per le card informative
          const startX = 0.8;
          const startY = 1.6;
          const numCards = cards.length;
          
          cards.forEach((card, index) => {
            let col = index % 2;
            let row = Math.floor(index / 2);
            let wCard = 5.6;
            let hCard = numCards === 3 && index === 2 ? 2.3 : 2.2;

            if (numCards === 4) {
              hCard = 2.4;
            }

            const posX = startX + col * 5.9;
            const posY = startY + row * 2.55;

            slide.addShape(pptx.ShapeType.roundRect, {
              x: posX,
              y: posY,
              w: wCard,
              h: hCard,
              fill: { color: 'FFFFFF' },
              line: { color: 'CBD5E1', width: 1 },
              rectRadius: 0.06
            });

            slide.addShape(pptx.ShapeType.rect, {
              x: posX,
              y: posY,
              w: 0.1,
              h: hCard,
              fill: { color: index === 0 ? '2563EB' : index === 1 ? '4F46E5' : index === 2 ? '059669' : 'D97706' },
            });

            slide.addText([
              { text: (card.title || '') + '\n', options: { bold: true, fontSize: 12, color: '0F172A', fontFace: 'Arial' } },
              { text: card.content || '', options: { fontSize: 10.5, color: '334155', lineSpacing: 14, fontFace: 'Arial' } }
            ], {
              x: posX + 0.3,
              y: posY + 0.18,
              w: 5.1,
              h: hCard - 0.3,
            });
          });
        }

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
        { message: 'Hai raggiunto il limite massimo giornaliero di richieste. Verifica il tuo piano a consumo.' },
        { status: 429 }
      );
    }

    return NextResponse.json(
      { message: error.message || 'Errore interno durante la generazione della presentazione.' },
      { status: 500 }
    );
  }
}
