import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import { google } from 'googleapis';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Configurazione autenticazione Google API (Drive & Slides)
const auth = new google.auth.GoogleAuth({
  credentials: {
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n'),
  },
  scopes: [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/presentations',
  ],
});

const drive = google.drive({ version: 'v3', auth });
const slides = google.slides({ version: 'v1', auth });

async function generateWithRetry(contents: any, maxRetries = 3) {
  const model = 'gemini-3.8-flash';
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = await ai.models.generateContent({ model, contents });
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
    const objective = (formData.get('objective') as string) || 'Presentazione Bilancio';
    const userAnswers = (formData.get('userAnswers') as string) || '';

    if (!file) {
      return NextResponse.json({ message: 'Nessun file caricato' }, { status: 400 });
    }

    const rawTemplateEnv = process.env.GOOGLE_DRIVE_TEMPLATE_ID;
    if (!rawTemplateEnv) {
      return NextResponse.json({ message: 'Manca la variabile d\'ambiente GOOGLE_DRIVE_TEMPLATE_ID' }, { status: 500 });
    }

    // Estrazione automatica e sicura del riferimento al template anche se viene inserito l'intero URL
    let TEMPLATE_ID = rawTemplateEnv.trim();
    const urlMatch = TEMPLATE_ID.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (urlMatch && urlMatch[1]) {
      TEMPLATE_ID = urlMatch[1];
    }

    const fileBuffer = await file.arrayBuffer();
    const fileBytes = Buffer.from(fileBuffer);

    // 1. Chiamata a Gemini per estrarre i dati mappati sui tag del template
    const prompt = [
      {
        inlineData: {
          mimeType: file.type || 'application/pdf',
          data: fileBytes.toString('base64'),
        },
      },
      `Sei un analista finanziario senior. Analizza il bilancio allegato e mappa i dati estratti per popolare un template di presentazione aziendale a 7 slide.
      
      Contesto utente:
      - Obiettivo: ${objective}
      - Settore: ${sector}
      - Note: ${userAnswers || 'Nessuna'}

      Restituisci **ESCLUSIVAMENTE** un oggetto JSON valido (senza markdown) strutturato con i tag esatti per la sostituzione:
      {
        "title_1": "Titolo della Copertina",
        "subtitle_1": "Sottotitolo istituzionale",
        "company_1": "Nome Azienda S.r.l.",
        "metric_val_1": "€ 2,2M",
        "metric_note_1": "+12% vs 2024",
        "metric_val_2": "€ 350K",
        "metric_note_2": "EBIT Operativo",
        "metric_val_3": "€ 180K",
        "metric_note_3": "Utile Netto",
        "metric_val_4": "€ 400K",
        "metric_note_4": "Liquidità Cassa",
        "narrative_1": "Commento direzionale sui risultati d'esercizio...",
        "segment_a_title": "Guanti Tecnici Industriali",
        "segment_a_val": "60% (€ 1.32M)",
        "segment_b_title": "Guanti in Pelle Lusso",
        "segment_b_val": "40% (€ 880K)",
        "table_row_1": "Valore della Produzione: € 2.200.000",
        "table_row_2": "Materie Prime: - € 900.000",
        "table_row_3": "Costo del Personale: - € 600.000",
        "table_row_4": "Servizi e Ammortamenti: - € 350.000",
        "table_row_5": "Risultato Netto: € 180.000",
        "kpi_roe": "22,5%",
        "kpi_roi": "18,4%",
        "kpi_ros": "15,9%",
        "kpi_de": "1,18",
        "capex_title": "Taglio Laser Robotizzato 2026",
        "capex_desc": "Investimento in automazione industriale finanziato con riserva straordinaria."
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

    const data = JSON.parse(cleanJsonStr);

    // 2. Copia il file template da Google Drive in un file temporaneo di lavoro
    const copyResponse = await drive.files.copy({
      fileId: TEMPLATE_ID,
      requestBody: {
        name: `DocuDecky_Presentazione_${Date.now()}`,
      },
    });

    const newFileId = copyResponse.data.id;
    if (!newFileId) {
      throw new Error('Impossibile duplicare il template Google Drive.');
    }

    // 3. Esegue la sostituzione automatica dei tag (es. {{title_1}}) con i dati elaborati
    const requests = Object.keys(data).map((key) => ({
      replaceAllText: {
        containsText: {
          text: `{{${key}}}`,
          matchCase: true,
        },
        replaceText: String(data[key] || ''),
      },
    }));

    await slides.presentations.batchUpdate({
      presentationId: newFileId,
      requestBody: { requests },
    });

    // 4. Esporta la presentazione risultante direttamente in formato .pptx
    const exportResponse = await drive.files.export(
      {
        fileId: newFileId,
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      },
      { responseType: 'stream' }
    );

    const chunks: any[] = [];
    for await (const chunk of exportResponse.data) {
      chunks.push(chunk);
    }
    const pptxBuffer = Buffer.concat(chunks);

    // 5. Elimina il file temporaneo creato su Google Drive per pulizia
    await drive.files.delete({ fileId: newFileId }).catch(() => {});

    return new NextResponse(pptxBuffer as unknown as BodyInit, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': 'attachment; filename="presentazione-docudecky-pro.pptx"',
      },
    });

  } catch (error: any) {
    console.error('Errore API Google Drive / Slides:', error);
    return NextResponse.json(
      { message: error.message || 'Errore interno durante l\'elaborazione del template.' },
      { status: 500 }
    );
  }
}
