import { NextResponse } from 'next/server';
import { GoogleGenerativeAI } from '@google/generative-ai';

// Inizializza il client con la chiave API dalle variabili d'ambiente
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY || '');

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File | null;
    const sector = formData.get('sector') as string || 'Business / Aziendale';
    const objective = formData.get('objective') as string || 'presentazione_pptx';
    const userAnswers = formData.get('userAnswers') as string || '';

    // Utilizzo del modello aggiornato gemini-2.5-flash con supporto multimodale per i file caricati
    const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });

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
        objectiveInstructions = 'Genera un documento testuale strutturato e approfondito in stile report .docx, suddiviso in sezioni con Executive Summary, analisi dettagliata, tabelle concettuali e conclusioni strategiche.';
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

    const result = await model.generateContent(promptParts);
    const responseText = result.response.text();

    // Controllo se l'IA richiede input aggiuntivi
    try {
      const cleanedText = responseText.replace(/```json/g, '').replace(/```/g, '').trim();
      const parsed = JSON.parse(cleanedText);
      if (parsed && parsed.needsInput) {
        return NextResponse.json(parsed);
      }
    } catch (e) {
      // Se non è JSON, è il contenuto finale generato
    }

    return NextResponse.json({
      success: true,
      data: responseText,
      objective: objective,
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
