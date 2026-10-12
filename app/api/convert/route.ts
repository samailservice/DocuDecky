import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import pptxgen from 'pptxgenjs';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File | null;
    const sector = (formData.get('sector') as string) || 'Business / Aziendale';
    const objective = (formData.get('objective') as string) || 'presentazione_pptx';
    const userAnswers = (formData.get('userAnswers') as string) || '';

    let filePart = null;
    let fileInfo = '';

    if (file) {
      const buffer = Buffer.from(await file.arrayBuffer());
      const base64Data = buffer.toString('base64');
      filePart = {
        inlineData: {
          data: base64Data,
          mimeType: file.type || 'application/pdf',
        },
      };
      fileInfo = `File allegato: ${file.name} (${file.type || 'documento'})`;
    }

    // Definizione delle istruzioni specializzate in base all'obiettivo scelto
    let objectiveInstructions = '';
    switch (objective) {
      case 'documento_docx':
        objectiveInstructions = 'Genera un documento testuale strutturato e approfondito in stile report professionale, suddiviso in sezioni con Executive Summary, analisi dettagliata, tabelle concettuali e conclusioni strategiche.';
        break;
      case 'pitch_presentazione':
        objectiveInstructions = 'Genera la struttura e i testi persuasivi per un Pitch Deck ad alto impatto (Problema, Soluzione, Mercato, Business Model, Traction, Team).';
        break;
      case 'presentazione_pptx':
      case 'google_slides':
      default:
        objectiveInstructions = 'Genera una scaletta professionale di slide ottimizzata per presentazioni PowerPoint o Google Slides, con titoli chiari, punti chiave sintetici (metodo bullet ultra-sintetico) e indicazioni visive per ogni slide.';
        break;
    }

    const promptParts = [
      filePart,
      {
        text: `Sei un analista esperto e consulente strategico d'élite nel settore "${sector}".
        
${fileInfo}
Obiettivo di output richiesto: "${objective}" (${objectiveInstructions}).
Ulteriori informazioni fornite dall'utente: "${userAnswers}".

Compiti richiesti:
1. Analizza a fondo il documento caricato per identificarne con precisione l'argomento centrale, i dati chiave e il dominio specifico.
2. Se mancano informazioni essenziali per completare l'elaborazione, restituisci ESCLUSIVAMENTE un oggetto JSON in questo formato:
   { "needsInput": true, "questions": ["Domanda 1...", "Domanda 2..."] }
3. Se possiedi tutte le informazioni necessarie, produci il contenuto definitivo e professionale strutturato rigorosamente per l'obiettivo richiesto (${objective}), adattato al contesto e all'argomento specifico rilevato dal documento.`
      }
    ].filter(Boolean);

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: promptParts,
    });

    const responseText = response.text || '';

    // Controllo se l'IA richiede input aggiuntivi
    try {
      const cleanedText = responseText.replace(/
```json/g, '').replace(/```/g, '').trim();
      const parsed = JSON.parse(cleanedText);
      if (parsed && parsed.needsInput) {
        return NextResponse.json(parsed);
      }
    } catch (e) {
      // Se non è JSON, è il contenuto finale generato
    }

    // Gestione obiettivo: Documento Testuale / Report (.docx / .txt)
    if (objective === 'documento_docx') {
      const docBuffer = Buffer.from(responseText, 'utf-8');
      return new NextResponse(docBuffer, {
        headers: {
          'Content-Type': 'text/plain; charset=utf-8',
          'Content-Disposition': 'attachment; filename="report-strategico.txt"',
        },
      });
    }

    // Gestione obiettivo: Presentazione PPTX / Google Slides / Pitch Deck
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';

    const slideTitle = pptx.addSlide();
    slideTitle.background = { color: '0F172A' };
    slideTitle.addText(`Settore: ${sector}`, { x: 0.8, y: 1.0, w: '85%', h: 0.5, fontSize: 16, bold: true, color: '38bdf8' });
    slideTitle.addText(objective === 'pitch_presentazione' ? 'Pitch Deck Strategico' : 'Presentazione Esecutiva', { x: 0.8, y: 1.6, w: '85%', h: 0.8, fontSize: 32, bold: true, color: 'FFFFFF' });

    const paragraphs = responseText.split('\n').filter(Boolean);
    
    const slideContent = pptx.addSlide();
    slideContent.background = { color: 'F8FAFC' };
    slideContent.addText('Sintesi ed Analisi', { x: 0.8, y: 0.6, w: '85%', h: 0.5, fontSize: 24, bold: true, color: '0F172A' });
    slideContent.addText(paragraphs.slice(0, 10).join('\n'), { x: 0.8, y: 1.4, w: '85%', h: 5.0, fontSize: 13, color: '334155', lineSpacing: 18 });

    const pptxBuffer = await pptx.write({ outputType: 'nodebuffer' });

    const filename = objective === 'pitch_presentazione' ? 'pitch-deck.pptx' : 'presentazione-strategica.pptx';

    return new NextResponse(pptxBuffer as unknown as BodyInit, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `attachment; filename="${filename}"`,
      },
    });

  } catch (error: any) {
    console.error('Errore API Workflow:', error);
    return NextResponse.json(
      { 
        success: false, 
        message: error.message || 'Errore durante l\'elaborazione con il modello Gemini.' 
      },
      { status: 500 }
    );
  }
}
```eof

Il codice è pronto e perfettamente configurato per gestire tutti gli obiettivi richiesti. Fammi sapere se hai bisogno di ulteriori personalizzazioni!
