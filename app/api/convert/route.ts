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
    const objective = (formData.get('objective') as string) || 'Presentazione Bilancio';
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
      `Sei un analista finanziario senior e un esperto di corporate storytelling. Analizza il bilancio allegato e genera una presentazione aziendale strutturata rigorosamente in esattamente 7 slide in stile consulenza strategica (stile McKinsey/BCG).
      
      Contesto utente:
      - Obiettivo: ${objective}
      - Settore: ${sector}
      - Note: ${userAnswers || 'Nessuna'}

      Restituisci **ESCLUSIVAMENTE** un oggetto JSON valido (senza blocchi markdown tipo \`\`\`json) strutturato con esattamente 7 slide in questo ordine:
      - Slide 1: Copertina istituzionale (Titolo, Sottotitolo, Descrizione, Azienda, Info Societarie).
      - Slide 2: Sintesi e Metriche Chiave (Narrativa e 4 metriche con valori e variazioni).
      - Slide 3: Pilastri Strategici / Mercato (Due elementi chiave contrapposti e nota di evidenza).
      - Slide 4: Conto Economico & Tabella Strutturata (Spaccato voci di conto economico e highlight efficienza).
      - Slide 5: Confronto Bilanciato (Impieghi vs Fonti, Totali e commento di sintesi).
      - Slide 6: Dashboard Indicatori e KPI (ROE, ROI, ROS, Debt/Equity con relative descrizioni).
      - Slide 7: Allocazione Risorse e Piano Strategico (Slogan, piano d'azione capex e ripartizione percentuali).

      Schema JSON richiesto:
      {
        "slides": [
          {
            "slideNumber": 1,
            "category": "ASSEMBLEA DEI SOCI",
            "title": "Bilancio d'Esercizio",
            "subtitle": "Risultati Economico-Finanziari e Strategia",
            "description": "Descrizione sintetica degli obiettivi e dei risultati dell'esercizio.",
            "company": "Nome Azienda S.r.l.",
            "infoItems": ["Sede Legale: ...", "P.IVA: ...", "Capitale Sociale: ..."]
          },
          {
            "slideNumber": 2,
            "section": "SEZIONE 01",
            "title": "Sintesi e Metriche Chiave",
            "narrative": "Commento principale della direzione sui risultati d'esercizio.",
            "metrics": [
              { "title": "FATTURATO", "value": "€ 2,2M", "note": "+12% vs anno prec." },
              { "title": "EBIT", "value": "€ 350K", "note": "15,9% ricavi" },
              { "title": "UTILE NETTO", "value": "€ 180K", "note": "8,2% margine" },
              { "title": "CASSA", "value": "€ 400K", "note": "Liquidità immediata" }
            ]
          },
          {
            "slideNumber": 3,
            "section": "SEZIONE 02",
            "title": "Pilastri Strategici & Mercato",
            "elementA": { "title": "Segmento Core", "metric": "60% (€ 1.32M)", "points": ["Punto chiave 1", "Punto chiave 2"] },
            "elementB": { "title": "Segmento Secondario", "metric": "40% (€ 880K)", "points": ["Punto chiave 1", "Punto chiave 2"] },
            "note": "Nota di evidenza trasversale o punto di forza."
          },
          {
            "slideNumber": 4,
            "section": "SEZIONE 03",
            "title": "Conto Economico & Tabella Strutturata",
            "highlight": "Riflessione sintetica sulla gestione dei costi e dei margini operativi.",
            "tableRows": [
              { "label": "Valore della Produzione", "value": "€ 2.200.000", "isTotal": true },
              { "label": "Materie Prime", "value": "- € 900.000", "isTotal": false },
              { "label": "Costo del Personale", "value": "- € 600.000", "isTotal": false },
              { "label": "Servizi ed Energia", "value": "- € 250.000", "isTotal": false },
              { "label": "Ammortamenti", "value": "- € 100.000", "isTotal": false },
              { "label": "Risultato Netto d'Esercizio", "value": "€ 180.000", "isTotal": true }
            ]
          },
          {
            "slideNumber": 5,
            "section": "SEZIONE 04",
            "title": "Confronto Bilanciato: Impieghi vs Fonti",
            "comment": "Commento riassuntivo sull'equilibrio patrimoniale e finanziario.",
            "sideA": { "title": "ATTIVO (Impieghi)", "total": "€ 1.900.000", "items": ["Attivo Fisso: € 650K", "Attivo Circolante: € 1,25M"] },
            "sideB": { "title": "PASSIVO & NETTO (Fonti)", "total": "€ 1.900.000", "items": ["Patrimonio Netto: € 800K", "Debiti Totali: € 950K", "TFR/Fondi: € 150K"] }
          },
          {
            "slideNumber": 6,
            "section": "SEZIONE 05",
            "title": "Dashboard Indicatori e KPI",
            "comment": "Valutazione complessiva degli indici prestazionali aziendali.",
            "kpis": [
              { "title": "ROE", "value": "22,5%", "desc": "Redditività capitale proprio." },
              { "title": "ROI", "value": "18,4%", "desc": "Efficienza operativa caratteristica." },
              { "title": "ROS", "value": "15,9%", "desc": "EBIT margin di settore." },
              { "title": "DEBT / EQUITY", "value": "1,18", "desc": "Leva finanziaria equilibrata." }
            ]
          },
          {
            "slideNumber": 7,
            "section": "SEZIONE 06",
            "title": "Allocazione Risorse e Piano Strategico",
            "slogan": "Visione di lungo termine e consolidamento tecnologico",
            "initiative": { "title": "Piano Investimenti 2026", "desc": "Investimento in automazione e innovazione di processo." },
            "allocations": [
              { "label": "Riserva Legale", "quota": "€ 9K (5%)" },
              { "label": "Riserva Straordinaria", "quota": "€ 121K (67,2%)" },
              { "label": "Dividendi ai Soci", "quota": "€ 50K (27,8%)" }
            ]
          }
        ]
      }

      Restituisci SOLO il JSON puro.`
    ];

    const jsonResponseText = await generateWithRetry(prompt);
    let cleanJsonStr = jsonResponseText.trim();
    if (cleanJsonStr.startsWith('```json')) {
      cleanJsonStr = cleanJsonStr.replace(/^```json/, '').replace(/```$/, '').trim();
    } else if (cleanJsonStr.startsWith('```')) {
      cleanJsonStr = cleanJsonStr.replace(/^```/, '').replace(/```$/, '').trim();
    }

    const presentationData = JSON.parse(cleanJsonStr);
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';

    const slidesArray = presentationData.slides || [];
    const totalSlides = slidesArray.length;

    if (slidesArray.length === 0) {
      const slide = pptx.addSlide();
      slide.addText("Nessun contenuto generato.", { x: 0.8, y: 0.8, w: '85%', h: '80%', fontSize: 14, color: '333333', fontFace: 'Arial' });
    } else {
      slidesArray.forEach((slideData: any, idx: number) => {
        const slide = pptx.addSlide();
        slide.background = { color: 'F8FAFC' };

        // Header superiore corporate
        slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: '2563EB' } });

        if (slideData.section) {
          slide.addText(slideData.section, { x: 0.8, y: 0.35, w: 4.0, h: 0.3, fontSize: 10, bold: true, color: '2563EB', fontFace: 'Arial' });
        }

        slide.addText(`Slide ${idx + 1} di ${totalSlides}`, { x: 10.5, y: 0.35, w: 2.0, h: 0.3, fontSize: 10, bold: true, color: '64748B', align: 'right', fontFace: 'Arial' });
        slide.addText(slideData.title || 'Titolo Slide', { x: 0.8, y: 0.65, w: '85%', h: 0.5, fontSize: 20, bold: true, color: '0F172A', fontFace: 'Arial' });

        if (idx === 0) {
          // SLIDE 1: Copertina
          if (slideData.category) {
            slide.addText(slideData.category, { x: 0.8, y: 1.3, w: '85%', h: 0.3, fontSize: 11, bold: true, color: '2563EB', fontFace: 'Arial' });
          }
          if (slideData.subtitle) {
            slide.addText(slideData.subtitle, { x: 0.8, y: 1.7, w: '85%', h: 0.4, fontSize: 14, bold: true, color: '334155', fontFace: 'Arial' });
          }
          if (slideData.description) {
            slide.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 2.3, w: 8.0, h: 1.5, fill: { color: 'FFFFFF' }, line: { color: 'CBD5E1', width: 1 }, rectRadius: 0.06 });
            slide.addText(slideData.description, { x: 1.1, y: 2.5, w: 7.4, h: 1.1, fontSize: 12, color: '475569', lineSpacing: 16, fontFace: 'Arial' });
          }
          slide.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 4.2, w: 8.0, h: 1.8, fill: { color: 'FFFFFF' }, line: { color: 'CBD5E1', width: 1 }, rectRadius: 0.06 });
          const infoText = (slideData.infoItems || []).join('\n');
          slide.addText([
            { text: "INFORMAZIONI SOCIETARIE\n", options: { bold: true, fontSize: 11, color: '0F172A' } },
            { text: infoText, options: { fontSize: 10.5, color: '334155', lineSpacing: 14 } }
          ], { x: 1.1, y: 4.4, w: 7.4, h: 1.4, fontFace: 'Arial' });

          if (slideData.company) {
            slide.addText(slideData.company, { x: 0.8, y: 6.8, w: '85%', h: 0.3, fontSize: 10, bold: true, color: '64748B', fontFace: 'Arial' });
          }

        } else if (idx === 1) {
          // SLIDE 2: Metriche Chiave
          if (slideData.narrative) {
            slide.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 1.3, w: 11.7, h: 1.0, fill: { color: 'FFFFFF' }, line: { color: 'CBD5E1', width: 1 }, rectRadius: 0.06 });
            slide.addText([
              { text: "NARRATIVA / COMMENTO PRINCIPALE\n", options: { bold: true, fontSize: 10, color: '2563EB' } },
              { text: slideData.narrative, options: { fontSize: 11, color: '334155' } }
            ], { x: 1.1, y: 1.45, w: 11.1, h: 0.7, fontFace: 'Arial' });
          }

          const metrics = slideData.metrics || [];
          metrics.slice(0, 4).forEach((m: any, mIdx: number) => {
            const posX = 0.8 + mIdx * 3.0;
            slide.addShape(pptx.ShapeType.roundRect, { x: posX, y: 2.6, w: 2.7, h: 3.8, fill: { color: 'FFFFFF' }, line: { color: 'CBD5E1', width: 1 }, rectRadius: 0.06 });
            slide.addShape(pptx.ShapeType.rect, { x: posX, y: 2.6, w: 2.7, h: 0.45, fill: { color: '2563EB' } });
            slide.addText(m.title || '', { x: posX + 0.1, y: 2.68, w: 2.5, h: 0.3, fontSize: 10, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Arial' });
            slide.addText(m.value || '', { x: posX + 0.1, y: 3.5, w: 2.5, h: 0.8, fontSize: 18, bold: true, color: '0F172A', align: 'center', fontFace: 'Arial' });
            slide.addText(m.note || '', { x: posX + 0.1, y: 4.8, w: 2.5, h: 1.2, fontSize: 10.5, color: '64748B', align: 'center', lineSpacing: 14, fontFace: 'Arial' });
          });

        } else if (idx === 2) {
          // SLIDE 3: Pilastri Strategici
          const elA = slideData.elementA || {};
          const elB = slideData.elementB || {};

          slide.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 1.3, w: 5.7, h: 4.2, fill: { color: 'FFFFFF' }, line: { color: 'CBD5E1', width: 1 }, rectRadius: 0.06 });
          slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.3, w: 0.12, h: 4.2, fill: { color: '2563EB' } });
          slide.addText([
            { text: (elA.title || 'Elemento A') + '\n', options: { bold: true, fontSize: 13, color: '0F172A' } },
            { text: (elA.metric || '') + '\n\n', options: { bold: true, fontSize: 16, color: '2563EB' } },
            { text: (elA.points || []).map((p: string) => `• ${p}`).join('\n'), options: { fontSize: 11, color: '334155', lineSpacing: 16 } }
          ], { x: 1.1, y: 1.5, w: 5.2, h: 3.8, fontFace: 'Arial' });

          slide.addShape(pptx.ShapeType.roundRect, { x: 6.8, y: 1.3, w: 5.7, h: 4.2, fill: { color: 'FFFFFF' }, line: { color: 'CBD5E1', width: 1 }, rectRadius: 0.06 });
          slide.addShape(pptx.ShapeType.rect, { x: 6.8, y: 1.3, w: 0.12, h: 4.2, fill: { color: '4F46E5' } });
          slide.addText([
            { text: (elB.title || 'Elemento B') + '\n', options: { bold: true, fontSize: 13, color: '0F172A' } },
            { text: (elB.metric || '') + '\n\n', options: { bold: true, fontSize: 16, color: '4F46E5' } },
            { text: (elB.points || []).map((p: string) => `• ${p}`).join('\n'), options: { fontSize: 11, color: '334155', lineSpacing: 16 } }
          ], { x: 7.1, y: 1.5, w: 5.2, h: 3.8, fontFace: 'Arial' });

          if (slideData.note) {
            slide.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 5.7, w: 11.7, h: 0.9, fill: { color: 'EFF6FF' }, line: { color: 'BFDBFE', width: 1 }, rectRadius: 0.06 });
            slide.addText(`Nota di Evidenza: ${slideData.note}`, { x: 1.1, y: 5.9, w: 11.1, h: 0.5, fontSize: 11, bold: true, color: '1E40AF', fontFace: 'Arial' });
          }

        } else if (idx === 3) {
          // SLIDE 4: Tabella Conto Economico
          const rows = slideData.tableRows || [];
          let tableY = 1.3;

          slide.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: tableY, w: 7.5, h: 5.2, fill: { color: 'FFFFFF' }, line: { color: 'CBD5E1', width: 1 }, rectRadius: 0.06 });
          slide.addText("VOCE DI CONTO ECONOMICO", { x: 1.1, y: tableY + 0.2, w: 4.5, h: 0.3, fontSize: 10, bold: true, color: '64748B', fontFace: 'Arial' });
          slide.addText("VALORE (€)", { x: 5.8, y: tableY + 0.2, w: 2.2, h: 0.3, fontSize: 10, bold: true, color: '64748B', align: 'right', fontFace: 'Arial' });

          let rowY = tableY + 0.6;
          rows.forEach((r: any) => {
            slide.addText(r.label || '', { x: 1.1, y: rowY, w: 4.5, h: 0.35, fontSize: r.isTotal ? 11 : 10, bold: r.isTotal, color: r.isTotal ? '0F172A' : '334155', fontFace: 'Arial' });
            slide.addText(r.value || '', { x: 5.8, y: rowY, w: 2.2, h: 0.35, fontSize: r.isTotal ? 11 : 10, bold: r.isTotal, color: r.isTotal ? '2563EB' : '334155', align: 'right', fontFace: 'Arial' });
            rowY += 0.5;
          });

          slide.addShape(pptx.ShapeType.roundRect, { x: 8.6, y: tableY, w: 3.9, h: 5.2, fill: { color: 'FFFFFF' }, line: { color: 'CBD5E1', width: 1 }, rectRadius: 0.06 });
          slide.addShape(pptx.ShapeType.rect, { x: 8.6, y: tableY, w: 3.9, h: 0.5, fill: { color: '2563EB' } });
          slide.addText("HIGHLIGHT EFFICIENZA", { x: 8.8, y: tableY + 0.12, w: 3.5, h: 0.3, fontSize: 10, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Arial' });
          slide.addText(slideData.highlight || '', { x: 8.9, y: tableY + 0.8, w: 3.3, h: 4.0, fontSize: 11, color: '334155', lineSpacing: 16, fontFace: 'Arial' });

        } else if (idx === 4) {
          // SLIDE 5: Stato Patrimoniale
          const sideA = slideData.sideA || {};
          const sideB = slideData.sideB || {};

          slide.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 1.3, w: 5.7, h: 4.2, fill: { color: 'FFFFFF' }, line: { color: 'CBD5E1', width: 1 }, rectRadius: 0.06 });
          slide.addText(sideA.title || '', { x: 1.1, y: 1.5, w: 5.1, h: 0.4, fontSize: 12, bold: true, color: '2563EB', fontFace: 'Arial' });
          slide.addText(`Totale: ${sideA.total || ''}`, { x: 1.1, y: 1.9, w: 5.1, h: 0.4, fontSize: 14, bold: true, color: '0F172A', fontFace: 'Arial' });
          slide.addText((sideA.items || []).map((it: string) => `• ${it}`).join('\n'), { x: 1.1, y: 2.5, w: 5.1, h: 2.5, fontSize: 11, color: '334155', lineSpacing: 16, fontFace: 'Arial' });

          slide.addShape(pptx.ShapeType.roundRect, { x: 6.8, y: 1.3, w: 5.7, h: 4.2, fill: { color: 'FFFFFF' }, line: { color: 'CBD5E1', width: 1 }, rectRadius: 0.06 });
          slide.addText(sideB.title || '', { x: 7.1, y: 1.5, w: 5.1, h: 0.4, fontSize: 12, bold: true, color: '4F46E5', fontFace: 'Arial' });
          slide.addText(`Totale: ${sideB.total || ''}`, { x: 7.1, y: 1.9, w: 5.1, h: 0.4, fontSize: 14, bold: true, color: '0F172A', fontFace: 'Arial' });
          slide.addText((sideB.items || []).map((it: string) => `• ${it}`).join('\n'), { x: 7.1, y: 2.5, w: 5.1, h: 2.5, fontSize: 11, color: '334155', lineSpacing: 16, fontFace: 'Arial' });

          if (slideData.comment) {
            slide.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 5.7, w: 11.7, h: 0.9, fill: { color: 'EFF6FF' }, line: { color: 'BFDBFE', width: 1 }, rectRadius: 0.06 });
            slide.addText(`Sintesi di Confronto: ${slideData.comment}`, { x: 1.1, y: 5.9, w: 11.1, h: 0.5, fontSize: 11, bold: true, color: '1E40AF', fontFace: 'Arial' });
          }

        } else if (idx === 5) {
          // SLIDE 6: Dashboard KPI
          const kpis = slideData.kpis || [];
          kpis.slice(0, 4).forEach((k: any, kIdx: number) => {
            const col = kIdx % 2;
            const row = Math.floor(kIdx / 2);
            const posX = 0.8 + col * 5.9;
            const posY = 1.3 + row * 2.4;

            slide.addShape(pptx.ShapeType.roundRect, { x: posX, y: posY, w: 5.6, h: 2.2, fill: { color: 'FFFFFF' }, line: { color: 'CBD5E1', width: 1 }, rectRadius: 0.06 });
            slide.addShape(pptx.ShapeType.rect, { x: posX, y: posY, w: 0.1, h: 2.2, fill: { color: kIdx === 0 ? '2563EB' : kIdx === 1 ? '4F46E5' : '059669' } });
            
            slide.addText(k.title || '', { x: posX + 0.3, y: posY + 0.2, w: 4.0, h: 0.3, fontSize: 11, bold: true, color: '64748B', fontFace: 'Arial' });
            slide.addText(k.value || '', { x: posX + 0.3, y: posY + 0.5, w: 5.0, h: 0.6, fontSize: 20, bold: true, color: '2563EB', fontFace: 'Arial' });
            slide.addText(k.desc || '', { x: posX + 0.3, y: posY + 1.2, w: 5.0, h: 0.8, fontSize: 11, color: '334155', lineSpacing: 14, fontFace: 'Arial' });
          });

          if (slideData.comment) {
            slide.addText(slideData.comment, { x: 0.8, y: 6.2, w: 11.7, h: 0.4, fontSize: 11, italic: true, color: '64748B', fontFace: 'Arial' });
          }

        } else if (idx === 6) {
          // SLIDE 7: Allocazione Risorse
          if (slideData.slogan) {
            slide.addText(slideData.slogan, { x: 0.8, y: 1.3, w: 11.7, h: 0.4, fontSize: 13, bold: true, color: '2563EB', fontFace: 'Arial' });
          }

          const init = slideData.initiative || {};
          slide.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 1.8, w: 5.7, h: 4.5, fill: { color: 'FFFFFF' }, line: { color: 'CBD5E1', width: 1 }, rectRadius: 0.06 });
          slide.addText([
            { text: "PIANO D'AZIONE STRATEGICO\n", options: { bold: true, fontSize: 11, color: '64748B' } },
            { text: (init.title || '') + '\n\n', options: { bold: true, fontSize: 14, color: '0F172A' } },
            { text: init.desc || '', options: { fontSize: 11, color: '334155', lineSpacing: 16 } }
          ], { x: 1.1, y: 2.1, w: 5.1, h: 3.9, fontFace: 'Arial' });

          const allocs = slideData.allocations || [];
          let allocY = 1.8;
          allocs.forEach((a: any) => {
            slide.addShape(pptx.ShapeType.roundRect, { x: 6.8, y: allocY, w: 5.7, h: 1.35, fill: { color: 'FFFFFF' }, line: { color: 'CBD5E1', width: 1 }, rectRadius: 0.06 });
            slide.addText([
              { text: (a.label || '') + '\n', options: { bold: true, fontSize: 11, color: '0F172A' } },
              { text: a.quota || '', options: { fontSize: 14, bold: true, color: '2563EB' } }
            ], { x: 7.1, y: allocY + 0.2, w: 5.1, h: 0.95, fontFace: 'Arial' });
            allocY += 1.55;
          });
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
    console.error('Errore API Locale:', error);

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
