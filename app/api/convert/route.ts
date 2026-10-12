import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import pptxgen from 'pptxgenjs';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

async function generateWithRetry(contents: any, maxRetries = 3) {
  const model = 'gemini-2.5-flash';

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = await ai.models.generateContent({
        model: model,
        contents: contents,
      });
      return response.text || '';
    } catch (error: any) {
      console.warn(`Tentativo ${attempt}/${maxRetries} fallito:`, error.message);
      if (error?.status === 429 || error?.message?.includes('RESOURCE_EXHAUSTED')) {
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
    const objective = (formData.get('objective') as string) || 'presentazione_pptx';
    const userAnswers = (formData.get('userAnswers') as string) || '';

    if (!file) {
      return NextResponse.json({ message: 'Nessun file caricato' }, { status: 400 });
    }

    const fileBuffer = await file.arrayBuffer();
    const fileBytes = Buffer.from(fileBuffer);

    let objectiveInstructions = '';
    switch (objective) {
      case 'documento_docx':
        objectiveInstructions = 'Genera un documento testuale approfondito e strutturato in sezioni (Executive Summary, Analisi Dettagliata, Raccomandazioni e Conclusioni).';
        break;
      case 'pitch_presentazione':
        objectiveInstructions = 'Genera la struttura e i testi per un Pitch Deck ad alto impatto (Problema, Soluzione, Mercato, Business Model, Traction, Team).';
        break;
      case 'google_slides':
        objectiveInstructions = 'Genera la scaletta di slide ottimizzata per Google Slides con titoli chiari e punti sintetici.';
        break;
      case 'presentazione_pptx':
      default:
        objectiveInstructions = 'Genera una presentazione professionale di slide per PowerPoint (PPTX) con titoli, punti chiave e dati salienti.';
        break;
    }

    const promptParts = [
      {
        inlineData: {
          mimeType: file.type || 'application/pdf',
          data: fileBytes.toString('base64'),
        },
      },
      `Sei un analista strategico ed esperto di dominio nel settore "${sector}".
      
Analizza accuratamente il documento allegato per identificarne l'argomento centrale, i numeri chiave e il contesto operativo.

Obiettivo richiesto dall'utente: "${objective}" (${objectiveInstructions}).
Note dell'utente: "${userAnswers || 'Nessuna'}".

ISTRUZIONI DI RISPOSTA:
1. Se nel documento o nel contesto mancano dati fondamentali per completare un'analisi accurata, restituisci ESCLUSIVAMENTE un JSON valido in questo formato:
{"needsInput": true, "questions": ["Domanda 1...", "Domanda 2..."]}

2. Se hai tutte le informazioni necessarie, genera il contenuto strutturato completo per l'obiettivo scelto.

Restituisci SOLO il testo o il JSON richiesto.`
    ];

    const responseText = await generateWithRetry(promptParts);
    let cleanText = responseText.trim();

    // Pulizia sicura da blocchi markdown senza interruzioni di riga nelle regex
    if (cleanText.includes('```json')) {
      cleanText = cleanText.replace(/```json/g, '').replace(/```/g, '').trim();
    } else if (cleanText.includes('```')) {
      cleanText = cleanText.replace(/```/g, '').trim();
    }

    // Controllo se l'IA richiede input aggiuntivi
    try {
      const parsed = JSON.parse(cleanText);
      if (parsed && parsed.needsInput) {
        return NextResponse.json(parsed);
      }
    } catch (e) {
      // Procedi con la generazione dell'output finale se non è JSON
    }

    // OUTPUT 1: Report Testuale / Documento (.docx / .txt)
    if (objective === 'documento_docx') {
      const docBuffer = Buffer.from(responseText, 'utf-8');
      return new NextResponse(docBuffer as unknown as BodyInit, {
        headers: {
          'Content-Type': 'text/plain; charset=utf-8',
          'Content-Disposition': 'attachment; filename="report-strategico.txt"',
        },
      });
    }

    // OUTPUT 2: Presentazioni PPTX / Google Slides / Pitch Deck
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';

    // Slide 1: Copertina
    const slide1 = pptx.addSlide();
    slide1.background = { color: '0F172A' };
    slide1.addText(sector.toUpperCase(), { x: 0.8, y: 1.0, w: '85%', h: 0.4, fontSize: 12, bold: true, color: '38BDF8', fontFace: 'Arial' });
    slide1.addText(
      objective === 'pitch_presentazione' ? 'Pitch Deck Strategico' : 'Presentazione Esecutiva',
      { x: 0.8, y: 1.6, w: '85%', h: 0.8, fontSize: 30, bold: true, color: 'FFFFFF', fontFace: 'Arial' }
    );
    slide1.addText(`Elaborazione basata su: ${file.name}`, { x: 0.8, y: 5.5, w: '85%', h: 0.4, fontSize: 11, color: '94A3B8', fontFace: 'Arial' });

    // Slide 2: Contenuto Principale
    const slide2 = pptx.addSlide();
    slide2.background = { color: 'F8FAFC' };
    slide2.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: '2563EB' } });
    slide2.addText('SINTESI STRATEGICA & ANALISI', { x: 0.8, y: 0.5, w: '85%', h: 0.4, fontSize: 18, bold: true, color: '0F172A', fontFace: 'Arial' });

    const paragraphs = responseText.split('\n').filter((p) => p.trim().length > 0);
    const bodyContent = paragraphs.slice(0, 12).join('\n');

    slide2.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 1.2, w: 11.7, h: 5.2, fill: { color: 'FFFFFF' }, line: { color: 'CBD5E1', width: 1 }, rectRadius: 0.06 });
    slide2.addText(bodyContent, { x: 1.1, y: 1.4, w: 11.1, h: 4.8, fontSize: 11, color: '334155', lineSpacing: 16, fontFace: 'Arial' });

    const pptxBuffer = await pptx.write({ outputType: 'nodebuffer' });

    const filename = objective === 'pitch_presentazione' ? 'pitch-deck.pptx' : 'presentazione-docudecky.pptx';

    return new NextResponse(pptxBuffer as unknown as BodyInit, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `attachment; filename="${filename}"`,
      },
    });

  } catch (error: any) {
    console.error('Errore API Convert:', error);
    return NextResponse.json(
      { message: error.message || 'Errore interno durante la generazione.' },
      { status: 500 }
    );
  }
}
